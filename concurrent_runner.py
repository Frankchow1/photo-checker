"""并发执行引擎：线程池 + 重试 + 自适应降速 + 失败补偿。

设计目标（按 5~10 并发考虑）：
1. 并发数可在界面设置，默认 6，允许 1~16；
2. 单条失败不影响整体，可重试（指数退避 + 随机抖动，避免雪崩）；
3. 遇到限流/服务端错误自动降并发，稳定后再慢慢恢复；
4. 一轮跑完后对“仍失败”的条目做一次低并发补偿重跑；
5. 任何时刻可停止，已完成结果不丢。

与 ai_client 的缓存配合：重试/补偿/断点续跑都不会重复计费，
因为成功结果按“模型+prompt+遮蔽后图片哈希”落盘。
"""

import random
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional

VERSION = "1.0.0"

DEFAULT_WORKERS = 6
HARD_MIN_WORKERS = 1
HARD_MAX_WORKERS = 16

# 可重试：限流、服务端错误、网络异常、返回不合规
RETRYABLE_HINTS = (
    "http_429",
    "http_500",
    "http_502",
    "http_503",
    "http_504",
    "request_error",
    "timeout",
    "parse_error",
    "schema_invalid",
)
# 不可重试：重试多少次结果都一样，重试只是浪费时间和钱
FATAL_HINTS = (
    "image_error",
    "photo_missing",
    "http_400",
    "http_401",
    "http_403",
    "http_404",
)
# 触发降并发的信号
THROTTLE_HINTS = ("http_429", "http_503", "http_504", "timeout")


@dataclass
class RunnerConfig:
    workers: int = DEFAULT_WORKERS
    max_retries: int = 3               # 首次之外的重试次数
    base_backoff: float = 1.5          # 退避基数（秒）
    max_backoff: float = 30.0
    jitter: float = 0.4                # 随机抖动比例，防止所有线程同时重试
    adaptive: bool = True              # 遇限流自动降并发
    min_workers: int = 2               # 降并发的下限
    recover_after_success: int = 25    # 连续成功多少条后尝试恢复 1 个并发
    cooldown_on_rate_limit: float = 3.0  # 被限流后全局冷却
    enable_compensation: bool = True   # 一轮跑完后对失败项补偿重跑
    compensation_workers: int = 2      # 补偿轮的低并发
    compensation_delay: float = 5.0    # 补偿轮开始前的等待

    def normalized(self) -> "RunnerConfig":
        workers = max(HARD_MIN_WORKERS, min(HARD_MAX_WORKERS, int(self.workers or DEFAULT_WORKERS)))
        min_workers = max(HARD_MIN_WORKERS, min(workers, int(self.min_workers or 1)))
        return RunnerConfig(
            workers=workers,
            max_retries=max(0, min(8, int(self.max_retries))),
            base_backoff=max(0.2, float(self.base_backoff)),
            max_backoff=max(1.0, float(self.max_backoff)),
            jitter=max(0.0, min(1.0, float(self.jitter))),
            adaptive=bool(self.adaptive),
            min_workers=min_workers,
            recover_after_success=max(5, int(self.recover_after_success)),
            cooldown_on_rate_limit=max(0.0, float(self.cooldown_on_rate_limit)),
            enable_compensation=bool(self.enable_compensation),
            compensation_workers=max(1, min(workers, int(self.compensation_workers))),
            compensation_delay=max(0.0, float(self.compensation_delay)),
        )


def classify_result(result: Any) -> str:
    """把一次调用结果分为 ok / retryable / fatal。"""
    if not isinstance(result, dict):
        return "ok"
    if not result.get("error"):
        return "ok"

    status = str(result.get("call_status") or "").lower()
    reason = str(result.get("reason") or "").lower()
    blob = status + " " + reason

    for hint in FATAL_HINTS:
        if hint in blob:
            return "fatal"
    for hint in RETRYABLE_HINTS:
        if hint in blob:
            return "retryable"
    # 其他不明异常默认当可重试（宁可多试一次，也不要静默丢掉）
    return "retryable"


def should_throttle(result: Any) -> bool:
    if not isinstance(result, dict):
        return False
    blob = (str(result.get("call_status") or "") + " " + str(result.get("reason") or "")).lower()
    return any(h in blob for h in THROTTLE_HINTS)


class AdaptiveLimiter:
    """可伸缩的并发闸门：被限流就收窄，稳定一段时间后再放宽。"""

    def __init__(self, limit: int, min_limit: int, max_limit: int, recover_after: int = 25):
        self._cv = threading.Condition()
        self.limit = max(1, limit)
        self.min_limit = max(1, min_limit)
        self.max_limit = max(self.limit, max_limit)
        self.recover_after = recover_after
        self.active = 0
        self._success_streak = 0
        self.shrink_events = 0
        self.grow_events = 0

    def acquire(self, should_stop: Optional[Callable[[], bool]] = None) -> bool:
        with self._cv:
            while self.active >= self.limit:
                if should_stop and should_stop():
                    return False
                self._cv.wait(timeout=0.3)
            self.active += 1
            return True

    def release(self):
        with self._cv:
            self.active = max(0, self.active - 1)
            self._cv.notify()

    def penalize(self) -> int:
        with self._cv:
            self._success_streak = 0
            if self.limit > self.min_limit:
                self.limit -= 1
                self.shrink_events += 1
            return self.limit

    def reward(self) -> int:
        with self._cv:
            self._success_streak += 1
            if self._success_streak >= self.recover_after and self.limit < self.max_limit:
                self.limit += 1
                self._success_streak = 0
                self.grow_events += 1
                self._cv.notify()
            return self.limit


@dataclass
class RunSummary:
    total: int = 0
    succeeded: int = 0
    failed: int = 0
    retried_items: int = 0
    total_retries: int = 0
    compensated: int = 0
    compensated_ok: int = 0
    stopped: bool = False
    elapsed_sec: float = 0.0
    start_workers: int = 0
    final_workers: int = 0
    throttle_events: int = 0
    failures: List[Dict[str, Any]] = field(default_factory=list)

    def as_dict(self) -> Dict[str, Any]:
        d = dict(self.__dict__)
        d["failures"] = list(self.failures)
        return d


class ConcurrentRunner:
    """并发跑任务，带重试与补偿。

    task_fn(item) -> dict；约定：失败时 dict 里 error=True、call_status 说明原因。
    on_result(done, total, item, result, meta)：每完成一条回调一次（已加锁，串行）。
    on_event(event_dict)：重试/降速/补偿等运行事件。
    should_stop()：返回 True 则尽快停。
    """

    def __init__(
        self,
        config: Optional[RunnerConfig] = None,
        on_result: Optional[Callable[..., None]] = None,
        on_event: Optional[Callable[[Dict[str, Any]], None]] = None,
        should_stop: Optional[Callable[[], bool]] = None,
    ):
        self.cfg = (config or RunnerConfig()).normalized()
        self.on_result = on_result
        self.on_event = on_event
        self.should_stop = should_stop or (lambda: False)

        self.limiter = AdaptiveLimiter(
            limit=self.cfg.workers,
            min_limit=self.cfg.min_workers if self.cfg.adaptive else self.cfg.workers,
            max_limit=self.cfg.workers,
            recover_after=self.cfg.recover_after_success,
        )
        self._lock = threading.Lock()
        self._done = 0
        self._cooldown_until = 0.0
        self.summary = RunSummary(start_workers=self.cfg.workers, final_workers=self.cfg.workers)

    # ---------- 内部工具 ----------

    def _emit(self, event: Dict[str, Any]):
        if self.on_event:
            try:
                self.on_event(event)
            except Exception:
                pass

    def _sleep_backoff(self, attempt: int):
        delay = min(self.cfg.max_backoff, self.cfg.base_backoff * (2 ** (attempt - 1)))
        delay = delay * (1.0 + random.uniform(-self.cfg.jitter, self.cfg.jitter))
        delay = max(0.1, delay)
        end = time.time() + delay
        while time.time() < end:
            if self.should_stop():
                return
            time.sleep(min(0.25, end - time.time()))

    def _wait_cooldown(self):
        while True:
            with self._lock:
                remain = self._cooldown_until - time.time()
            if remain <= 0 or self.should_stop():
                return
            time.sleep(min(0.25, remain))

    def _trigger_cooldown(self):
        if self.cfg.cooldown_on_rate_limit <= 0:
            return
        with self._lock:
            self._cooldown_until = max(self._cooldown_until, time.time() + self.cfg.cooldown_on_rate_limit)

    # ---------- 单条执行 ----------

    def _run_one(self, index: int, item: Any, task_fn: Callable[[Any], Dict[str, Any]], total: int, phase: str):
        if self.should_stop():
            return

        attempts = 0
        result: Dict[str, Any] = {}
        max_attempts = 1 + self.cfg.max_retries

        while attempts < max_attempts:
            if self.should_stop():
                return
            self._wait_cooldown()

            if not self.limiter.acquire(self.should_stop):
                return
            attempts += 1
            try:
                result = task_fn(item)
            except Exception as e:
                result = {
                    "error": True,
                    "call_status": "request_error",
                    "reason": "任务执行异常: %s" % str(e)[:300],
                    "confidence": 0.0,
                }
            finally:
                self.limiter.release()

            kind = classify_result(result)

            if kind == "ok":
                if self.cfg.adaptive:
                    self.limiter.reward()
                break

            if should_throttle(result):
                with self._lock:
                    self.summary.throttle_events += 1
                if self.cfg.adaptive:
                    new_limit = self.limiter.penalize()
                    self._emit({
                        "type": "throttled",
                        "workers": new_limit,
                        "detail": str(result.get("call_status", "")),
                    })
                self._trigger_cooldown()

            if kind == "fatal" or attempts >= max_attempts:
                break

            with self._lock:
                self.summary.total_retries += 1
            self._emit({
                "type": "retry",
                "index": index,
                "attempt": attempts,
                "max_attempts": max_attempts,
                "status": str(result.get("call_status", "")),
                "phase": phase,
            })
            self._sleep_backoff(attempts)

        ok = not (isinstance(result, dict) and result.get("error"))
        meta = {
            "attempts": attempts,
            "retries": max(0, attempts - 1),
            "phase": phase,
            "ok": ok,
            "workers_now": self.limiter.limit,
            "status": (result or {}).get("call_status", ""),
        }

        with self._lock:
            self._done += 1
            done = self._done
            if attempts > 1:
                self.summary.retried_items += 1
            if ok:
                self.summary.succeeded += 1
                if phase == "compensation":
                    self.summary.compensated_ok += 1
            else:
                self.summary.failed += 1
                self.summary.failures.append({
                    "index": index,
                    "item": item,
                    "status": meta["status"],
                    "reason": (result or {}).get("reason", ""),
                    "attempts": attempts,
                })
            if self.on_result:
                try:
                    self.on_result(done, total, item, result, meta)
                except Exception:
                    pass

    # ---------- 对外入口 ----------

    def run(self, items: List[Any], task_fn: Callable[[Any], Dict[str, Any]]) -> Dict[str, Any]:
        items = list(items or [])
        total = len(items)
        self.summary = RunSummary(
            total=total,
            start_workers=self.cfg.workers,
            final_workers=self.cfg.workers,
        )
        self._done = 0

        if total == 0:
            return self.summary.as_dict()

        started = time.time()
        self._emit({"type": "start", "total": total, "workers": self.cfg.workers})

        with ThreadPoolExecutor(max_workers=self.cfg.workers) as pool:
            futures = [
                pool.submit(self._run_one, i, item, task_fn, total, "main")
                for i, item in enumerate(items)
            ]
            for f in futures:
                try:
                    f.result()
                except Exception:
                    pass

        # ---- 补偿轮：对仍失败且可重试的条目，用低并发再跑一次 ----
        if (
            self.cfg.enable_compensation
            and not self.should_stop()
            and self.summary.failures
        ):
            retryable = [
                f for f in self.summary.failures
                if classify_result({"error": True, "call_status": f.get("status", ""), "reason": f.get("reason", "")}) == "retryable"
            ]
            if retryable:
                self._emit({
                    "type": "compensation_start",
                    "count": len(retryable),
                    "workers": self.cfg.compensation_workers,
                    "delay": self.cfg.compensation_delay,
                })
                end = time.time() + self.cfg.compensation_delay
                while time.time() < end and not self.should_stop():
                    time.sleep(0.2)

                # 补偿轮刷新计数：把这批从 failed 里摸掉，重新记录
                failed_indices = {f["index"] for f in retryable}
                with self._lock:
                    self.summary.failures = [f for f in self.summary.failures if f["index"] not in failed_indices]
                    self.summary.failed -= len(retryable)
                    self._done -= len(retryable)
                    self.summary.compensated = len(retryable)

                # 补偿轮回到低并发，避免再次撞限流
                self.limiter.limit = self.cfg.compensation_workers
                with ThreadPoolExecutor(max_workers=self.cfg.compensation_workers) as pool:
                    futures = [
                        pool.submit(self._run_one, f["index"], f["item"], task_fn, total, "compensation")
                        for f in retryable
                    ]
                    for fu in futures:
                        try:
                            fu.result()
                        except Exception:
                            pass

        self.summary.elapsed_sec = round(time.time() - started, 2)
        self.summary.final_workers = self.limiter.limit
        self.summary.stopped = bool(self.should_stop())
        self._emit({"type": "done", **{k: v for k, v in self.summary.as_dict().items() if k != "failures"}})
        return self.summary.as_dict()


def suggest_workers(total_items: int) -> int:
    """给个保守建议值：小批量没必要开大并发。"""
    if total_items <= 10:
        return 2
    if total_items <= 50:
        return 4
    if total_items <= 500:
        return 6
    return 8
