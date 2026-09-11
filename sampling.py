import math
import random
from typing import List, Dict, Any, Tuple

# AQL (GB/T 2828.1 / ISO 2859-1) 正常检验一次抽样方案（一般检验水平 II）
AQL_TABLE = [
    (8, 8),
    (15, 3),
    (25, 5),
    (50, 8),
    (90, 13),
    (150, 20),
    (280, 32),
    (500, 50),
    (1200, 80),
    (3200, 125),
    (10000, 200),
    (35000, 315),
    (150000, 500),
    (500000, 800),
]

def calculate_aql_sample_size(total_pairs: int) -> int:
    """根据总数计算 AQL 国际/国家标准抽检数量"""
    if total_pairs <= 0:
        return 0
    for max_n, sample_n in AQL_TABLE:
        if total_pairs <= max_n:
            return min(sample_n, total_pairs)
    return min(1250, total_pairs)

def calculate_confidence_sample_size(total_pairs: int, confidence: float = 0.95, margin_error: float = 0.05, p: float = 0.95) -> int:
    """基于有限总体 95% 置信度与指定误差限计算科学样本量"""
    if total_pairs <= 0:
        return 0
    # Z-score for 95% is ~1.96
    z = 1.96 if confidence >= 0.95 else 1.645
    # Cochran's formula for finite population
    numerator = total_pairs * (z ** 2) * p * (1 - p)
    denominator = ((total_pairs - 1) * (margin_error ** 2)) + ((z ** 2) * p * (1 - p))
    sample_size = math.ceil(numerator / denominator)
    return max(1, min(sample_size, total_pairs))

def stratified_sample(items: List[Any], sample_size: int) -> List[Any]:
    """
    分层随机抽样 (Stratified Random Sampling)
    将整个列表划分为等宽层，在各层内部随机抽样，确保覆盖全生命周期/全批次任务，避免集中在某一头部或尾部。
    """
    total = len(items)
    if sample_size >= total:
        return list(items)
    if sample_size <= 0:
        return []

    # 将 items 分为 sample_size 个区间，每个区间抽取 1 个元素
    step = total / sample_size
    sampled = []
    for i in range(sample_size):
        start_idx = int(i * step)
        end_idx = int((i + 1) * step)
        end_idx = max(start_idx + 1, min(end_idx, total))
        chosen_idx = random.randint(start_idx, end_idx - 1)
        sampled.append(items[chosen_idx])

    return sampled
