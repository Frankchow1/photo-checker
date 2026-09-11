#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
门头照复核包生成与多规则筛选 GUI 工具 (Photo Review Builder GUI)
===================================================================
功能特点：
1. 文件与目录配置：支持 Excel (.xlsx, .xls) 与 CSV/TSV 文件，自动读取 Sheet 与表头。
2. 原始数据列依据映射：可直观指定商户号、商户名、地址、提交时间、作业员、任务号、文件名等原始数据列，支持智能自动推荐。
3. 相当商户（同店）判断规则：
   - 支持根据商户号、商户名、地址任意组合判断是否为相同商户；
   - 支持“过滤排除相同商户（过滤同店，重点审查跨商户造假）”、“仅保留相同商户”、“全部保留”等业务模式；
   - 支持作业员关系过滤（同一作业员 / 跨作业员 / 不限）。
4. 判断的时间等规则：
   - 支持时间差天数计算（自动计算两张照片的拍摄/提交时间绝对差值）；
   - 支持最大时间差天数（如 <= 7天窗口）、最小时间差天数筛选；
   - 支持任务提交时间起始/截止日期区间限制。
5. 优先级与算法指标阈值：
   - 复核优先级高/中/低/未验证多选；
   - 支持 pHash 距离、ORB 内点数、ORB 比例等指标阈值过滤。
6. 实时统计与打包生成：
   - 支持毫秒级统计筛选匹配数（同商户数、跨商户数、预计包数）；
   - 多线程并发分包打包，生成包含精美离线审核系统的 HTML、图片压缩包及汇总 Excel。
"""

from __future__ import annotations

import os
import sys
import math
import json
import time
import shutil
import zipfile
import threading
import subprocess
from pathlib import Path
from typing import List, Tuple, Dict, Optional, Any, Set

import pandas as pd
from PIL import Image, ImageOps

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ==============================================================================
# 内嵌离线 HTML 模板 (与 build_review_packages_v3 保持完全一致并兼容)
# ==============================================================================

HTML_TEMPLATE = r'''<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>门头照复核</title>
<style>
:root{--bg:#f6f6f4;--surface:#fff;--soft:#f1f0ee;--border:#e4e3e0;--text:#2c2c2b;--muted:#77736e;--blue:#2783de;--blue-soft:#e5f2fc;--red:#d9554b;--red-soft:#fce9e7;--orange:#c66a25;--orange-soft:#fbebde;--green:#2f8f62;--green-soft:#e8f1ec;--shadow:0 1px 2px rgba(0,0,0,.04),0 6px 20px rgba(0,0,0,.04)}
*{box-sizing:border-box}html,body{height:100%;margin:0}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",Arial,sans-serif;background:var(--bg);color:var(--text);font-size:14px;overflow:hidden}button,input,select,textarea{font:inherit;color:inherit}button{cursor:pointer}.app{height:100%;display:grid;grid-template-rows:60px 1fr}.topbar{display:flex;align-items:center;gap:18px;background:var(--surface);border-bottom:1px solid var(--border);padding:8px 14px}.brand{display:flex;align-items:center;gap:10px;min-width:230px}.brand-mark{width:34px;height:34px;border-radius:9px;background:var(--blue-soft);color:var(--blue);display:grid;place-items:center;font-size:18px}.brand h1{font-size:17px;margin:0}.brand small{display:block;color:var(--muted);margin-top:2px}.progress{display:flex;align-items:center;gap:10px;flex:1}.track{height:7px;background:var(--soft);border-radius:99px;overflow:hidden;width:min(320px,45vw)}.bar{height:100%;background:var(--blue);width:0}.progress span{color:var(--muted);font-variant-numeric:tabular-nums}.top-actions{display:flex;gap:7px}.btn{min-height:38px;border:1px solid var(--border);border-radius:8px;background:var(--surface);padding:7px 11px}.btn:hover{background:var(--soft)}.btn.primary{background:var(--blue);border-color:var(--blue);color:#fff}.btn:focus-visible,.queue-item:focus-visible,.decision:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible{outline:3px solid rgba(39,131,222,.22);outline-offset:1px}.layout{min-height:0;display:grid;grid-template-columns:300px minmax(590px,1fr) 330px;gap:10px;padding:10px}.panel{min-height:0;background:var(--surface);border:1px solid var(--border);border-radius:11px;box-shadow:var(--shadow);overflow:hidden}.queue{display:grid;grid-template-rows:auto 1fr}.queue-head{padding:12px;border-bottom:1px solid var(--border)}.queue-head h2{font-size:14px;margin:0 0 9px}.queue-controls{display:grid;gap:7px}.date-row{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:6px}.date-row input{min-width:0;padding:0 4px;font-size:12px}.filter-label{font-size:12px;color:var(--muted);margin-bottom:-3px}.queue-controls input,.queue-controls select,.reviewer input{width:100%;height:38px;border:1px solid var(--border);border-radius:8px;background:var(--surface);padding:0 9px}.reset-filter{width:100%;min-height:34px}.task-list{overflow:auto;padding:6px}.queue-item{width:100%;text-align:left;border:1px solid transparent;background:transparent;border-radius:8px;padding:9px;margin:2px 0}.queue-item:hover{background:var(--soft)}.queue-item.active{background:var(--blue-soft);border-color:#abd2f2}.queue-item.done .seq::after{content:' ✓';color:var(--green)}.queue-top{display:flex;align-items:center;justify-content:space-between;gap:6px}.seq{font-weight:700;font-size:13px}.tag{font-size:12px;border-radius:5px;padding:2px 5px}.tag.high{color:var(--red);background:var(--red-soft)}.tag.medium{color:var(--orange);background:var(--orange-soft)}.tag.low{color:var(--muted);background:var(--soft)}.task-pair{font-size:12px;color:var(--muted);line-height:1.45;margin-top:5px}.task-pair div{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.center{display:grid;grid-template-rows:auto 1fr auto}.center-head{display:flex;align-items:center;justify-content:space-between;padding:10px 12px;border-bottom:1px solid var(--border)}.current-title{display:flex;align-items:center;gap:8px}.current-title strong{font-size:15px}.position{color:var(--muted)}.compare{min-height:0;display:grid;grid-template-columns:1fr 1fr;gap:1px;background:var(--border)}.photo-card{min-width:0;min-height:0;background:var(--surface);display:grid;grid-template-rows:auto minmax(260px,1fr) 126px}.photo-head{padding:8px 10px;border-bottom:1px solid var(--border);min-width:0}.photo-head strong{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.photo-head span{display:block;color:var(--muted);font-size:12px;margin-top:2px}.image-frame{min-height:0;background:#171717;display:flex;align-items:center;justify-content:center;overflow:hidden;padding:6px}.image-frame img{display:block;max-width:100%;max-height:100%;width:auto;height:auto;object-fit:contain;cursor:zoom-in}.missing{color:#ddd;text-align:center}.watermark-box{height:126px;background:#111;border-top:1px solid var(--border);position:relative;overflow:hidden}.watermark-box::before{content:'底部水印放大';position:absolute;left:7px;top:6px;z-index:2;background:rgba(0,0,0,.68);color:white;border-radius:5px;padding:3px 6px;font-size:11px}.watermark-box img{width:100%;height:100%;object-fit:cover;object-position:center bottom;cursor:zoom-in}.watermark-box.hidden{display:none}.metrics{display:flex;gap:16px;align-items:center;padding:9px 12px;min-height:42px;border-top:1px solid var(--border);overflow:auto;white-space:nowrap}.metric small{color:var(--muted);margin-right:4px}.metric b{font-variant-numeric:tabular-nums}.review{display:grid;grid-template-rows:auto 1fr auto}.review-head{padding:11px 12px;border-bottom:1px solid var(--border)}.reviewer label{display:block;color:var(--muted);font-size:12px;margin-bottom:5px}.details{overflow:auto;padding:10px 12px}.detail-card{background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:9px;margin-bottom:8px}.detail-card h3{font-size:13px;margin:0 0 7px}.kv{margin:6px 0}.kv small{display:block;color:var(--muted);font-size:12px;margin-bottom:1px}.kv div{line-height:1.4;word-break:break-word}.reason{background:var(--blue-soft);color:#205f94;padding:9px;border-radius:8px;line-height:1.45}.review-actions{padding:10px 12px;border-top:1px solid var(--border)}.decisions{display:grid;grid-template-columns:1fr 1fr;gap:6px}.decision{min-height:42px;border:1px solid var(--border);border-radius:8px;background:var(--surface);text-align:left;padding:7px 8px}.decision:hover{background:var(--soft)}.decision.selected{background:var(--blue-soft);border-color:var(--blue);box-shadow:inset 0 0 0 1px var(--blue)}.decision b{display:inline-block;width:18px;color:var(--muted)}.decision.full{grid-column:1/-1}.notes{width:100%;height:50px;resize:none;border:1px solid var(--border);border-radius:8px;padding:7px 9px;margin-top:7px}.nav{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:7px}.modal{position:fixed;inset:0;background:rgba(0,0,0,.9);z-index:20;display:none;align-items:center;justify-content:center}.modal.open{display:flex}.modal img{max-width:95vw;max-height:91vh;object-fit:contain;transform:scale(1)}.modal-tools{position:absolute;right:16px;top:14px;display:flex;gap:7px}.modal-tools .btn{background:rgba(255,255,255,.96);color:#222}.toast{position:fixed;left:50%;bottom:22px;transform:translateX(-50%);background:#222;color:#fff;border-radius:7px;padding:8px 13px;opacity:0;pointer-events:none;z-index:30}.toast.show{opacity:1}.empty{padding:24px;text-align:center;color:var(--muted)}
@media(max-width:1100px){body{overflow:auto}.app{height:auto;min-height:100%}.layout{grid-template-columns:220px minmax(540px,1fr)}.review{grid-column:1/-1;min-height:620px}.photo-card{grid-template-rows:auto 400px 120px}}
@media(max-width:760px){.topbar{height:auto;flex-wrap:wrap}.app{grid-template-rows:auto 1fr}.progress{order:3;width:100%}.layout{display:block}.queue{height:300px;margin-bottom:10px}.center{height:auto;min-height:0;grid-template-rows:auto auto auto;margin-bottom:10px}.compare{grid-template-columns:1fr}.photo-card{grid-template-rows:auto 300px 110px}.review{min-height:760px}.top-actions{width:100%;overflow:auto}.track{width:100%}}
@media(prefers-color-scheme:dark){:root{--bg:#191919;--surface:#202020;--soft:#383836;--border:rgba(255,255,255,.18);--text:#fff;--muted:rgba(255,255,255,.65);--blue:#5e9fe8;--blue-soft:rgba(94,159,232,.14);--red:#e97366;--red-soft:rgba(233,115,102,.14);--orange:#de9255;--orange-soft:rgba(222,146,85,.14);--green:#72bc8f;--green-soft:rgba(114,188,143,.14);--shadow:none}.reason{color:#a8d2f5}}
</style>
</head>
<body>
<div class="app">
<header class="topbar">
  <div class="brand"><div class="brand-mark">◎</div><div><h1>门头照复核</h1><small id="batchLabel"></small></div></div>
  <div class="progress"><div class="track"><div class="bar" id="progressBar"></div></div><span id="progressText"></span></div>
  <div class="top-actions"><button class="btn" id="watermarkBtn">隐藏水印放大</button><button class="btn" id="importBtn">导入进度</button><button class="btn" id="backupBtn">备份 JSON</button><button class="btn primary" id="exportBtn">导出 CSV</button><input type="file" id="importFile" accept="application/json" hidden></div>
</header>
<div class="layout">
  <aside class="queue panel"><div class="queue-head"><h2>任务号列表</h2><div class="queue-controls"><input id="taskSearch" placeholder="搜索任务号/组名/文件名"><select id="operatorFilter"><option value="">全部作业员</option></select><input id="addressSearch" placeholder="筛选地址关键词"><div class="filter-label">提交时间（任一照片命中）</div><div class="date-row"><input type="date" id="startDate" aria-label="开始日期"><input type="date" id="endDate" aria-label="结束日期"></div><select id="filter"><option>全部</option><option>未复核</option><option>高优先级</option><option>中优先级</option><option>已确认复用</option><option>同店正常</option></select><button class="btn reset-filter" id="resetFilters">重置筛选</button></div></div><div class="task-list" id="taskList"></div></aside>
  <main class="center panel"><div class="center-head"><div class="current-title"><strong id="groupId"></strong><span class="tag" id="priority"></span><span class="position" id="positionText"></span></div><span class="position">点击图片可放大</span></div>
    <section class="compare">
      <article class="photo-card"><div class="photo-head"><strong id="file1"></strong><span id="taskHead1"></span></div><div class="image-frame"><img id="img1" alt="候选照片1"><div class="missing" id="missing1" hidden>照片缺失</div></div><div class="watermark-box" id="watermarkBox1"><img id="watermark1" alt="照片1底部水印"></div></article>
      <article class="photo-card"><div class="photo-head"><strong id="file2"></strong><span id="taskHead2"></span></div><div class="image-frame"><img id="img2" alt="候选照片2"><div class="missing" id="missing2" hidden>照片缺失</div></div><div class="watermark-box" id="watermarkBox2"><img id="watermark2" alt="照片2底部水印"></div></article>
    </section><div class="metrics" id="metrics"></div>
  </main>
  <aside class="review panel"><div class="review-head"><div class="reviewer"><label for="reviewer">复核人</label><input id="reviewer" placeholder="请输入姓名"></div></div><div class="details"><div id="detail1" class="detail-card"></div><div id="detail2" class="detail-card"></div><div class="reason" id="reason"></div></div><div class="review-actions"><div class="decisions"><button class="decision" data-value="确认复用"><b>1</b>确认复用</button><button class="decision" data-value="正常不同照片"><b>2</b>正常不同</button><button class="decision" data-value="同店正常拍摄"><b>3</b>同店正常</button><button class="decision" data-value="系统图"><b>4</b>系统图</button><button class="decision full" data-value="无法判断"><b>5</b>无法判断</button></div><textarea id="notes" class="notes" placeholder="可选备注"></textarea><div class="nav"><button class="btn" id="prevBtn">← 上一组</button><button class="btn primary" id="nextBtn">下一组 →</button></div></div></aside>
</div>
</div>
<div class="modal" id="modal"><div class="modal-tools"><button class="btn" id="zoomOut">−</button><button class="btn" id="zoomReset">100%</button><button class="btn" id="zoomIn">＋</button><button class="btn" id="closeModal">关闭</button></div><img id="modalImage" alt="大图预览"></div><div class="toast" id="toast"></div>
<script>
const BATCH_ID=__BATCH_ID_JSON__;const ITEMS=__CANDIDATES_JSON__;const VERDICTS=['确认复用','正常不同照片','同店正常拍摄','系统图','无法判断'];const storageKey='photo-review-v2:'+BATCH_ID;
let state={reviewer:'',reviews:{},currentId:ITEMS[0]?.id||'',filter:'全部',taskSearch:'',operator:'',addressSearch:'',startDate:'',endDate:'',showWatermark:true};try{const saved=JSON.parse(localStorage.getItem(storageKey)||'null');if(saved)state={...state,...saved};if(state.search&&!state.taskSearch)state.taskSearch=state.search}catch(e){}
let visible=[];let zoom=1;const $=id=>document.getElementById(id);function safe(v){return v===null||v===undefined||v===''?'—':String(v)}function esc(v){return safe(v).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}function reviewFor(item){return state.reviews[item.id]||{verdict:'',notes:'',reviewer:'',reviewTime:''}}function save(){localStorage.setItem(storageKey,JSON.stringify(state))}
function dateKey(v){const m=String(v||'').match(/(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);return m?`${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`:''}function renderOperatorOptions(){const values=[...new Set(ITEMS.flatMap(i=>[i.auditor1,i.auditor2]).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'zh-CN'));$('operatorFilter').innerHTML='<option value="">全部作业员</option>'+values.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');$('operatorFilter').value=state.operator}function rebuildVisible(){const tq=state.taskSearch.trim().toLowerCase();const aq=state.addressSearch.trim().toLowerCase();visible=ITEMS.filter(item=>{const r=reviewFor(item);let ok=true;if(state.filter==='未复核')ok=!r.verdict;else if(state.filter==='高优先级')ok=item.priority==='高';else if(state.filter==='中优先级')ok=item.priority==='中';else if(state.filter==='已确认复用')ok=r.verdict==='确认复用';else if(state.filter==='同店正常')ok=r.verdict==='同店正常拍摄';if(!ok)return false;if(state.operator&&![item.auditor1,item.auditor2].includes(state.operator))return false;if(tq&&![item.task1,item.task2,item.file1,item.file2,item.group].join(' ').toLowerCase().includes(tq))return false;if(aq&&![item.address1,item.address2].join(' ').toLowerCase().includes(aq))return false;if(state.startDate||state.endDate){const dates=[dateKey(item.time1),dateKey(item.time2)].filter(Boolean);const dateOk=dates.some(d=>(!state.startDate||d>=state.startDate)&&(!state.endDate||d<=state.endDate));if(!dateOk)return false}return true});if(!visible.some(x=>x.id===state.currentId))state.currentId=visible[0]?.id||''}
function current(){return ITEMS.find(x=>x.id===state.currentId)||null}function priorityClass(v){return v==='高'?'high':v==='中'?'medium':'low'}
function renderList(){const html=visible.map((item,index)=>{const r=reviewFor(item);return`<button class="queue-item ${item.id===state.currentId?'active':''} ${r.verdict?'done':''}" data-id="${esc(item.id)}"><div class="queue-top"><span class="seq">${index+1}. ${esc(item.group)}</span><span class="tag ${priorityClass(item.priority)}">${esc(item.priority)}</span></div><div class="task-pair"><div>${esc(item.task1)}</div><div>↔ ${esc(item.task2)}</div></div></button>`}).join('');$('taskList').innerHTML=html||'<div class="empty">没有符合条件的任务</div>';document.querySelectorAll('.queue-item').forEach(btn=>btn.onclick=()=>{state.currentId=btn.dataset.id;render()});const active=document.querySelector('.queue-item.active');if(active)active.scrollIntoView({block:'nearest'})}
function detailHtml(title,item,side){return`<h3>${title}</h3><div class="kv"><small>任务号</small><div>${esc(item['task'+side])}</div></div><div class="kv"><small>商户号</small><div>${esc(item['merchantNo'+side])}</div></div><div class="kv"><small>商户名</small><div>${esc(item['merchantName'+side])}</div></div><div class="kv"><small>导入地址</small><div>${esc(item['address'+side])}</div></div><div class="kv"><small>采集项</small><div>${esc(item['item'+side])}</div></div><div class="kv"><small>提交时间</small><div>${esc(item['time'+side])}</div></div><div class="kv"><small>作业员</small><div>${esc(item['auditor'+side])}</div></div>`}function metric(label,value){return`<div class="metric"><small>${label}</small><b>${esc(value)}</b></div>`}
function setImage(img,missing,src){img.hidden=false;missing.hidden=true;img.onerror=()=>{img.hidden=true;missing.hidden=false};img.src=src}
function render(){rebuildVisible();renderList();renderOperatorOptions();$('batchLabel').textContent=BATCH_ID;$('reviewer').value=state.reviewer;$('filter').value=state.filter;$('taskSearch').value=state.taskSearch;$('addressSearch').value=state.addressSearch;$('startDate').value=state.startDate;$('endDate').value=state.endDate;const reviewed=ITEMS.filter(x=>reviewFor(x).verdict).length;$('progressText').textContent=`已复核 ${reviewed} / ${ITEMS.length}`;$('progressBar').style.width=(ITEMS.length?reviewed/ITEMS.length*100:0)+'%';const item=current();if(!item){$('groupId').textContent='暂无任务';return}const r=reviewFor(item);const pos=visible.findIndex(x=>x.id===item.id);$('groupId').textContent=item.group;$('priority').textContent=item.priority+'优先级';$('priority').className='tag '+priorityClass(item.priority);$('positionText').textContent=`${pos+1} / ${visible.length}`;$('file1').textContent=item.file1;$('file2').textContent=item.file2;$('taskHead1').textContent=item.task1;$('taskHead2').textContent=item.task2;setImage($('img1'),$('missing1'),item.image1);setImage($('img2'),$('missing2'),item.image2);$('watermark1').src=item.image1;$('watermark2').src=item.image2;$('watermarkBox1').classList.toggle('hidden',!state.showWatermark);$('watermarkBox2').classList.toggle('hidden',!state.showWatermark);$('watermarkBtn').textContent=state.showWatermark?'隐藏水印放大':'显示水印放大';$('detail1').innerHTML=detailHtml('照片 1',item,'1');$('detail2').innerHTML=detailHtml('照片 2',item,'2');$('reason').innerHTML=esc(item.reason)+(item.diffDays?` <span style="font-size:12px;color:var(--muted);">(间隔 ${item.diffDays} 天)</span>`:'');$('metrics').innerHTML=metric('pHash',item.phash)+metric('最佳区域',item.phashArea)+metric('dHash',item.dhash)+metric('ORB内点',item.orbInliers)+metric('内点比例',item.orbRatio);$('notes').value=r.notes||'';document.querySelectorAll('.decision').forEach(b=>b.classList.toggle('selected',b.dataset.value===r.verdict));save();preloadNext()}
function move(delta){if(!visible.length)return;let i=visible.findIndex(x=>x.id===state.currentId);i=Math.min(Math.max(i+delta,0),visible.length-1);state.currentId=visible[i].id;render()}function setVerdict(value){const item=current();if(!item)return;state.reviews[item.id]={...reviewFor(item),verdict:value,notes:$('notes').value,reviewer:state.reviewer,reviewTime:new Date().toISOString()};save();toast('已记录：'+value);move(1)}function preloadNext(){const i=visible.findIndex(x=>x.id===state.currentId);const n=visible[i+1];if(n){new Image().src=n.image1;new Image().src=n.image2}}function toast(t){$('toast').textContent=t;$('toast').classList.add('show');setTimeout(()=>$('toast').classList.remove('show'),1200)}
function csvCell(v){const s=v==null?'':String(v);return'"'+s.replaceAll('"','""')+'"'}function download(name,text,type){const blob=new Blob([text],{type});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),800)}function exportCsv(){const h=['batch_id','candidate_id','group','priority','file1','file2','task1','task2','merchant_no_1','merchant_no_2','merchant_1','merchant_2','address1','address2','diff_days','verdict','notes','reviewer','review_time','phash','orb_inliers','orb_ratio'];const rows=[h.join(',')];ITEMS.forEach(item=>{const r=reviewFor(item);rows.push([BATCH_ID,item.id,item.group,item.priority,item.file1,item.file2,item.task1,item.task2,item.merchantNo1,item.merchantNo2,item.merchantName1,item.merchantName2,item.address1,item.address2,item.diffDays,r.verdict,r.notes,r.reviewer,r.reviewTime,item.phash,item.orbInliers,item.orbRatio].map(csvCell).join(','))});download(`${BATCH_ID}_复核结果.csv`,'\ufeff'+rows.join('\r\n'),'text/csv;charset=utf-8')}function backup(){download(`${BATCH_ID}_进度备份.json`,JSON.stringify({batchId:BATCH_ID,state},null,2),'application/json')}function importBackup(file){const reader=new FileReader();reader.onload=()=>{try{const data=JSON.parse(reader.result);if(data.batchId!==BATCH_ID)throw new Error('批次不一致');state={...state,...data.state};save();render();toast('进度已恢复')}catch(e){alert('导入失败：'+e.message)}};reader.readAsText(file)}
function openModal(src){zoom=1;$('modalImage').src=src;$('modalImage').style.transform='scale(1)';$('zoomReset').textContent='100%';$('modal').classList.add('open')}function changeZoom(d){zoom=Math.min(Math.max(zoom+d,.5),4);$('modalImage').style.transform=`scale(${zoom})`;$('zoomReset').textContent=Math.round(zoom*100)+'%'}
$('reviewer').onchange=e=>{state.reviewer=e.target.value.trim();save()};$('taskSearch').oninput=e=>{state.taskSearch=e.target.value;render()};$('operatorFilter').onchange=e=>{state.operator=e.target.value;render()};$('addressSearch').oninput=e=>{state.addressSearch=e.target.value;render()};$('startDate').onchange=e=>{state.startDate=e.target.value;render()};$('endDate').onchange=e=>{state.endDate=e.target.value;render()};$('filter').onchange=e=>{state.filter=e.target.value;render()};$('resetFilters').onclick=()=>{state.filter='全部';state.taskSearch='';state.operator='';state.addressSearch='';state.startDate='';state.endDate='';render()};$('watermarkBtn').onclick=()=>{state.showWatermark=!state.showWatermark;render()};$('notes').onchange=()=>{const item=current();if(item){state.reviews[item.id]={...reviewFor(item),notes:$('notes').value,reviewer:state.reviewer};save()}};document.querySelectorAll('.decision').forEach(b=>b.onclick=()=>setVerdict(b.dataset.value));$('prevBtn').onclick=()=>move(-1);$('nextBtn').onclick=()=>move(1);$('exportBtn').onclick=exportCsv;$('backupBtn').onclick=backup;$('importBtn').onclick=()=>$('importFile').click();$('importFile').onchange=e=>e.target.files[0]&&importBackup(e.target.files[0]);['img1','img2','watermark1','watermark2'].forEach(id=>$(id).onclick=()=>openModal($(id).src));$('closeModal').onclick=()=>$('modal').classList.remove('open');$('modal').onclick=e=>{if(e.target===$('modal'))$('modal').classList.remove('open')};$('zoomIn').onclick=()=>changeZoom(.25);$('zoomOut').onclick=()=>changeZoom(-.25);$('zoomReset').onclick=()=>{zoom=1;$('modalImage').style.transform='scale(1)';$('zoomReset').textContent='100%'};document.addEventListener('keydown',e=>{if(['INPUT','TEXTAREA','SELECT'].includes(document.activeElement.tagName))return;if(e.key>='1'&&e.key<='5'){setVerdict(VERDICTS[Number(e.key)-1]);e.preventDefault()}else if(e.key==='ArrowLeft'){move(-1);e.preventDefault()}else if(e.key==='ArrowRight'){move(1);e.preventDefault()}else if(e.key==='Escape')$('modal').classList.remove('open')});render();
</script>
</body>
</html>'''

README_TEMPLATE = '''门头照复核包使用说明（任务列表版）

1. 先完整解压 ZIP，不能在压缩包内直接打开。
2. 推荐使用最新版 Chrome、Edge 或 Safari 浏览器。
3. 双击“开始复核.html”。本工具完全离线运行，不会上传任何照片到外部网络。
4. 首次打开请输入复核人姓名。
5. 快捷键支持：
   - 数字键 1：确认复用
   - 数字键 2：正常不同照片
   - 数字键 3：同店正常拍摄
   - 数字键 4：系统图
   - 数字键 5：无法判断
   - 方向键 ←/→：切换上一组 / 下一组
6. 浏览器会本地暂存复核进度，但结束工作前务必点击“导出 CSV”，并建议点击“备份 JSON”。
7. 将导出的 CSV 发回复核任务负责人。
'''


# ==============================================================================
# 核心数据引擎 (Engine)
# ==============================================================================

def clean_val(value: Any) -> str:
    if pd.isna(value) or value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def format_num(value: Any, digits: Optional[int] = None) -> Any:
    if pd.isna(value) or value == "" or value is None:
        return ""
    try:
        res = float(value)
        if digits is not None:
            return round(res, digits)
        return int(res) if res.is_integer() else res
    except Exception:
        return clean_val(value)


class ReviewPackageEngine:
    """负责数据加载、规则过滤、分包切分及图片缩放打包的核心引擎"""

    @staticmethod
    def detect_text_params(file_path: str) -> Tuple[str, str]:
        encodings = ["utf-8-sig", "utf-8", "gb18030", "gbk"]
        best_enc = "utf-8"
        sample_bytes = b""
        with open(file_path, "rb") as f:
            sample_bytes = f.read(16384)
        for enc in encodings:
            try:
                sample_bytes.decode(enc)
                best_enc = enc
                break
            except UnicodeDecodeError:
                continue
        text_sample = sample_bytes.decode(best_enc, errors="ignore")
        first_line = text_sample.splitlines()[0] if text_sample.splitlines() else ""
        if first_line.count("\t") > first_line.count(","):
            delimiter = "\t"
        elif first_line.count(";") > first_line.count(","):
            delimiter = ";"
        else:
            delimiter = ","
        return best_enc, delimiter

    @classmethod
    def inspect_file(cls, file_path: str) -> Dict[str, Any]:
        """快速提取表格的 Sheet 列表和首行表头"""
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")

        ext = os.path.splitext(file_path)[1].lower()
        is_excel = ext in [".xlsx", ".xls", ".xlsm"]
        sheets: List[str] = []
        columns: List[str] = []

        if is_excel:
            xl = pd.ExcelFile(file_path)
            sheets = xl.sheet_names
            first_sheet = sheets[0] if sheets else 0
            df_preview = pd.read_excel(xl, sheet_name=first_sheet, nrows=2)
            columns = [str(c) for c in df_preview.columns]
        else:
            enc, sep = cls.detect_text_params(file_path)
            df_preview = pd.read_csv(file_path, sep=sep, encoding=enc, nrows=2)
            columns = [str(c) for c in df_preview.columns]

        return {
            "is_excel": is_excel,
            "sheets": sheets,
            "columns": columns,
        }

    @classmethod
    def read_dataset(cls, file_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """加载表格完整数据"""
        ext = os.path.splitext(file_path)[1].lower()
        if ext in [".xlsx", ".xls", ".xlsm"]:
            sheet = sheet_name if sheet_name else 0
            return pd.read_excel(file_path, sheet_name=sheet)
        else:
            enc, sep = cls.detect_text_params(file_path)
            return pd.read_csv(file_path, sep=sep, encoding=enc, low_memory=False)

    @classmethod
    def get_columns(cls, file_path: str, sheet_name: Optional[str] = None) -> List[str]:
        """快速提取指定 sheet 的列名（仅读取表头）"""
        ext = os.path.splitext(file_path)[1].lower()
        if ext in [".xlsx", ".xls", ".xlsm"]:
            sheet = sheet_name if sheet_name and sheet_name != "(默认)" else 0
            df_preview = pd.read_excel(file_path, sheet_name=sheet, nrows=1)
            return [str(c) for c in df_preview.columns]
        else:
            enc, sep = cls.detect_text_params(file_path)
            df_preview = pd.read_csv(file_path, sep=sep, encoding=enc, nrows=1)
            return [str(c) for c in df_preview.columns]

    @classmethod
    def guess_column_mapping(cls, columns: List[str]) -> Dict[str, str]:
        """智能猜测各个逻辑字段对应的实际列名，支持两两比对候选表与原始单表"""
        col_map = {str(c).strip().lower(): str(c) for c in columns}

        def find_match(*patterns: str) -> str:
            # 1. 优先完全精准匹配
            for pat in patterns:
                p_lower = pat.lower()
                for k, orig in col_map.items():
                    if k == p_lower:
                        return orig
            # 2. 前缀/后缀或包含
            for pat in patterns:
                p_lower = pat.lower()
                for k, orig in col_map.items():
                    if p_lower in k:
                        return orig
            return ""

        has_pair = any(str(c).endswith("_1") or str(c).endswith("1") for c in columns) and any(str(c).endswith("_2") or str(c).endswith("2") for c in columns)

        mapping = {
            "is_pair_table": "1" if has_pair else "0",
            "file1": find_match("文件名1", "完整路径1", "file1", "image1", "doc_url_1", "doc_url", "item_name", "文件名"),
            "file2": find_match("文件名2", "完整路径2", "file2", "image2", "doc_url_2", "doc_url", "item_name", "文件名"),
            "path1": find_match("完整路径1", "本地路径1", "path1", "doc_url_1", "doc_url"),
            "path2": find_match("完整路径2", "本地路径2", "path2", "doc_url_2", "doc_url"),
            "task1": find_match("task_no_1", "task1", "任务号1", "任务编号1", "task_no", "task_id"),
            "task2": find_match("task_no_2", "task2", "任务号2", "任务编号2", "task_no", "task_id"),
            "merchant_no_1": find_match("merchant_no_1", "商户号1", "商户编号1", "merchant_no", "merchants_no"),
            "merchant_no_2": find_match("merchant_no_2", "商户号2", "商户编号2", "merchant_no", "merchants_no"),
            "merchant_name_1": find_match("merchant_1", "merchants_name_1", "商户名1", "商户名称1", "merchant", "merchants_name"),
            "merchant_name_2": find_match("merchant_2", "merchants_name_2", "商户名2", "商户名称2", "merchant", "merchants_name"),
            "address1": find_match("import_address_1", "address1", "地址1", "导入地址1", "import_address", "live_address"),
            "address2": find_match("import_address_2", "address2", "地址2", "导入地址2", "import_address", "live_address"),
            "time1": find_match("submit_time_1", "upload_time_1", "提交时间1", "时间1", "submit_time", "upload_time"),
            "time2": find_match("submit_time_2", "upload_time_2", "提交时间2", "时间2", "submit_time", "upload_time"),
            "diff_days": find_match("时间差天", "time_diff_days", "时间差"),
            "auditor1": find_match("auditor_name_1", "auditor1", "作业员1", "审核人1", "auditor_name", "auditor"),
            "auditor2": find_match("auditor_name_2", "auditor2", "作业员2", "审核人2", "auditor_name", "auditor"),
            "item1": find_match("item_name_1", "item1", "采集项1", "item_name"),
            "item2": find_match("item_name_2", "item2", "采集项2", "item_name"),
            "priority": find_match("复核优先级", "priority", "优先级"),
            "reason": find_match("判定依据", "reason", "依据"),
            "biz_type": find_match("候选业务类型", "biz_type", "业务类型"),
            "group": find_match("候选组", "group", "组名"),
            "seq": find_match("复核序号", "seq", "序号"),
            "phash": find_match("pHash最小距离", "phash_min", "phash"),
            "phash_area": find_match("最佳pHash区域", "phash_area"),
            "dhash": find_match("dHash距离", "dhash"),
            "orb_inliers": find_match("ORB内点数", "orb_inliers"),
            "orb_ratio": find_match("ORB内点比例", "orb_ratio"),
        }
        return mapping

    @classmethod
    def apply_filter_rules(
        cls,
        df: pd.DataFrame,
        mapping: Dict[str, str],
        rules: Dict[str, Any],
    ) -> Tuple[pd.DataFrame, Dict[str, int]]:
        """
        核心规则计算与筛选：
        1. 相当商户判断（商户号/商户名/地址一致性）；
        2. 商户筛选策略（排除同商户/仅同商户/全部）；
        3. 作业员关系筛选；
        4. 时间差及提交时间范围筛选；
        5. 优先级及算法阈值筛选；
        6. 排序与 limit 上限。
        """
        data = df.copy()
        stats = {
            "total_input": len(data),
            "same_merchant_count": 0,
            "cross_merchant_count": 0,
            "time_filtered_count": 0,
            "priority_filtered_count": 0,
            "final_count": 0,
        }
        if data.empty:
            return data, stats

        # -------------------------------------------------------------
        # 1. 相当商户（同店）判断
        # -------------------------------------------------------------
        # 判断依据：商户号、商户名、地址
        mno1 = mapping.get("merchant_no_1", "")
        mno2 = mapping.get("merchant_no_2", "")
        mname1 = mapping.get("merchant_name_1", "")
        mname2 = mapping.get("merchant_name_2", "")
        addr1 = mapping.get("address1", "")
        addr2 = mapping.get("address2", "")

        check_by_no = rules.get("same_merchant_by_no", True) and mno1 in data.columns and mno2 in data.columns
        check_by_name = rules.get("same_merchant_by_name", True) and mname1 in data.columns and mname2 in data.columns
        check_by_addr = rules.get("same_merchant_by_addr", False) and addr1 in data.columns and addr2 in data.columns

        same_merchant_series = pd.Series(False, index=data.index)

        if check_by_no:
            s1 = data[mno1].astype(str).str.strip()
            s2 = data[mno2].astype(str).str.strip()
            valid_no = (s1 != "") & (s1 != "nan") & (s1 != "None")
            same_merchant_series |= (valid_no & (s1 == s2))

        if check_by_name:
            s1 = data[mname1].astype(str).str.strip()
            s2 = data[mname2].astype(str).str.strip()
            valid_name = (s1 != "") & (s1 != "nan") & (s1 != "None")
            same_merchant_series |= (valid_name & (s1 == s2))

        if check_by_addr:
            s1 = data[addr1].astype(str).str.strip()
            s2 = data[addr2].astype(str).str.strip()
            valid_addr = (s1 != "") & (s1 != "nan") & (s1 != "None")
            same_merchant_series |= (valid_addr & (s1 == s2))

        # 动态打标到 DataFrame 列中，方便后续审查与导出
        data["_is_same_merchant"] = same_merchant_series
        stats["same_merchant_count"] = int(same_merchant_series.sum())
        stats["cross_merchant_count"] = int(len(data) - stats["same_merchant_count"])

        # 根据用户选择的商户过滤模式执行过滤
        merchant_filter_mode = rules.get("merchant_filter_mode", "exclude_same")
        if merchant_filter_mode == "exclude_same":
            # 排除相同商户（重点审查跨商户造假）
            data = data[~data["_is_same_merchant"]].copy()
        elif merchant_filter_mode == "only_same":
            # 仅保留相同商户（同店跨期核查）
            data = data[data["_is_same_merchant"]].copy()
        # else "all": 保留全部

        # -------------------------------------------------------------
        # 2. 作业员关系筛选
        # -------------------------------------------------------------
        aud1 = mapping.get("auditor1", "")
        aud2 = mapping.get("auditor2", "")
        operator_mode = rules.get("operator_mode", "all")
        if aud1 in data.columns and aud2 in data.columns and operator_mode != "all":
            a1 = data[aud1].astype(str).str.strip()
            a2 = data[aud2].astype(str).str.strip()
            valid_aud = (a1 != "") & (a1 != "nan") & (a2 != "") & (a2 != "nan")
            if operator_mode == "only_same_operator":
                data = data[valid_aud & (a1 == a2)].copy()
            elif operator_mode == "only_diff_operator":
                data = data[valid_aud & (a1 != a2)].copy()

        # -------------------------------------------------------------
        # 3. 判断的时间等规则（时间差计算 & 绝对区间）
        # -------------------------------------------------------------
        diff_col = mapping.get("diff_days", "")
        t1_col = mapping.get("time1", "")
        t2_col = mapping.get("time2", "")

        # 确保有 _diff_days 数值列
        if diff_col in data.columns:
            data["_diff_days"] = pd.to_numeric(data[diff_col], errors="coerce")
        elif t1_col in data.columns and t2_col in data.columns:
            t1 = pd.to_datetime(data[t1_col], errors="coerce")
            t2 = pd.to_datetime(data[t2_col], errors="coerce")
            data["_diff_days"] = (t1 - t2).abs().dt.total_seconds() / 86400.0
        else:
            data["_diff_days"] = float("nan")

        # 最大时间差过滤 (例如 <= 7天)
        max_diff = rules.get("max_diff_days", None)
        if max_diff is not None and not data["_diff_days"].isna().all():
            data = data[data["_diff_days"].isna() | (data["_diff_days"] <= max_diff)].copy()

        # 最小时间差过滤 (例如 >= 0.01天，排除同秒连拍)
        min_diff = rules.get("min_diff_days", None)
        if min_diff is not None and not data["_diff_days"].isna().all():
            data = data[data["_diff_days"].isna() | (data["_diff_days"] >= min_diff)].copy()

        # 绝对日期区间 (start_date ~ end_date)
        start_date = rules.get("start_date", "")
        end_date = rules.get("end_date", "")
        if (start_date or end_date) and (t1_col in data.columns or t2_col in data.columns):
            cond = pd.Series(True, index=data.index)
            if t1_col in data.columns:
                t1_dates = pd.to_datetime(data[t1_col], errors="coerce").dt.strftime("%Y-%m-%d")
            else:
                t1_dates = pd.Series("", index=data.index)
            if t2_col in data.columns:
                t2_dates = pd.to_datetime(data[t2_col], errors="coerce").dt.strftime("%Y-%m-%d")
            else:
                t2_dates = pd.Series("", index=data.index)

            if start_date:
                cond &= ((t1_dates >= start_date) | (t2_dates >= start_date))
            if end_date:
                cond &= ((t1_dates <= end_date) | (t2_dates <= end_date))
            data = data[cond].copy()

        stats["time_filtered_count"] = len(data)

        # -------------------------------------------------------------
        # 4. 优先级与算法指标阈值
        # -------------------------------------------------------------
        priorities = rules.get("priorities", ["高", "中"])
        p_col = mapping.get("priority", "")
        if p_col in data.columns and priorities:
            data = data[data[p_col].astype(str).str.strip().isin(priorities)].copy()

        # pHash 最小距离阈值
        phash_max = rules.get("phash_max", None)
        phash_col = mapping.get("phash", "")
        if phash_max is not None and phash_col in data.columns:
            phash_vals = pd.to_numeric(data[phash_col], errors="coerce")
            data = data[phash_vals.isna() | (phash_vals <= phash_max)].copy()

        # ORB 内点数阈值
        orb_inliers_min = rules.get("orb_inliers_min", None)
        inlier_col = mapping.get("orb_inliers", "")
        if orb_inliers_min is not None and inlier_col in data.columns:
            inliers = pd.to_numeric(data[inlier_col], errors="coerce")
            data = data[inliers.isna() | (inliers >= orb_inliers_min)].copy()

        # ORB 内点比例阈值
        orb_ratio_min = rules.get("orb_ratio_min", None)
        ratio_col = mapping.get("orb_ratio", "")
        if orb_ratio_min is not None and ratio_col in data.columns:
            ratios = pd.to_numeric(data[ratio_col], errors="coerce")
            data = data[ratios.isna() | (ratios >= orb_ratio_min)].copy()

        # -------------------------------------------------------------
        # 5. 排序与限制
        # -------------------------------------------------------------
        order_map = {"高": 0, "中": 1, "低": 2, "未验证": 3}
        if p_col in data.columns:
            data["_order"] = data[p_col].astype(str).str.strip().map(order_map).fillna(9)
        else:
            data["_order"] = 9

        sort_cols = ["_order"]
        ascending_flags = [True]

        if phash_col in data.columns:
            data["_phash_num"] = pd.to_numeric(data[phash_col], errors="coerce")
            sort_cols.append("_phash_num")
            ascending_flags.append(True)

        if inlier_col in data.columns:
            data["_orb_inliers_num"] = pd.to_numeric(data[inlier_col], errors="coerce")
            sort_cols.append("_orb_inliers_num")
            ascending_flags.append(False)

        if ratio_col in data.columns:
            data["_orb_ratio_num"] = pd.to_numeric(data[ratio_col], errors="coerce")
            sort_cols.append("_orb_ratio_num")
            ascending_flags.append(False)

        data = data.sort_values(sort_cols, ascending=ascending_flags, na_position="last")
        data = data.drop(columns=[c for c in ["_order", "_phash_num", "_orb_inliers_num", "_orb_ratio_num"] if c in data.columns])

        limit = rules.get("limit", 0)
        if limit > 0:
            data = data.head(limit)

        data = data.reset_index(drop=True)
        stats["final_count"] = len(data)
        return data, stats

    @classmethod
    def resolve_photo_path(cls, photo_root: Path, row: pd.Series, mapping: Dict[str, str], side: int) -> Optional[Path]:
        """寻找真实图片文件的绝对路径"""
        f_col = mapping.get(f"file{side}", "")
        p_col = mapping.get(f"path{side}", "")

        # 1. 优先尝试本地照片根目录 + 文件名
        filename = clean_val(row.get(f_col)) if f_col else ""
        if filename:
            # 去除可能的 url 前缀或路径
            fname = Path(filename).name
            direct = photo_root / fname
            if direct.is_file():
                return direct

        # 2. 尝试从绝对路径字段读取
        path_str = clean_val(row.get(p_col)) if p_col else ""
        if path_str:
            abs_p = Path(path_str)
            if abs_p.is_file():
                return abs_p

        # 3. 尝试直接以 filename 构造成绝对路径
        if filename:
            maybe_abs = Path(filename)
            if maybe_abs.is_file():
                return maybe_abs

        return None

    @classmethod
    def prepare_photo(cls, source: Path, destination: Path, max_px: int, quality: int):
        """转码与等比缩放照片到目标打包目录"""
        if destination.exists() and destination.stat().st_size > 0:
            return
        destination.parent.mkdir(parents=True, exist_ok=True)
        with Image.open(source) as image:
            image = ImageOps.exif_transpose(image).convert("RGB")
            image.thumbnail((max_px, max_px), Image.Resampling.LANCZOS)
            image.save(destination, "JPEG", quality=quality, optimize=True, progressive=True)

    @classmethod
    def build_record(
        cls,
        row: pd.Series,
        mapping: Dict[str, str],
        image1_rel: str,
        image2_rel: str,
        index: int,
    ) -> Dict[str, Any]:
        """将一行数据转换为离线 HTML 复核系统所需的标准化 JSON 数据"""
        seq_col = mapping.get("seq", "")
        group_col = mapping.get("group", "")
        seq = clean_val(row.get(seq_col)) if seq_col else str(index + 1)
        group = clean_val(row.get(group_col)) if group_col else f"PAIR-{seq}"

        # 判定依据
        reason_col = mapping.get("reason", "")
        biz_col = mapping.get("biz_type", "")
        reason = clean_val(row.get(reason_col)) if reason_col else ""
        biz_type = clean_val(row.get(biz_col)) if biz_col else ""
        if biz_type and biz_type not in reason:
            reason = f"[{biz_type}] {reason}" if reason else biz_type

        diff_days_val = row.get("_diff_days")
        diff_days_str = f"{diff_days_val:.2f}" if pd.notna(diff_days_val) else ""

        return {
            "id": f"{group}-{seq}",
            "group": group,
            "priority": clean_val(row.get(mapping.get("priority", ""))) or "高",
            "reason": reason,
            "diffDays": diff_days_str,
            "file1": Path(clean_val(row.get(mapping.get("file1", "")))).name,
            "file2": Path(clean_val(row.get(mapping.get("file2", "")))).name,
            "image1": image1_rel,
            "image2": image2_rel,
            "task1": clean_val(row.get(mapping.get("task1", ""))),
            "task2": clean_val(row.get(mapping.get("task2", ""))),
            "merchantNo1": clean_val(row.get(mapping.get("merchant_no_1", ""))),
            "merchantNo2": clean_val(row.get(mapping.get("merchant_no_2", ""))),
            "merchantName1": clean_val(row.get(mapping.get("merchant_name_1", ""))),
            "merchantName2": clean_val(row.get(mapping.get("merchant_name_2", ""))),
            "address1": clean_val(row.get(mapping.get("address1", ""))),
            "address2": clean_val(row.get(mapping.get("address2", ""))),
            "item1": clean_val(row.get(mapping.get("item1", ""))),
            "item2": clean_val(row.get(mapping.get("item2", ""))),
            "auditor1": clean_val(row.get(mapping.get("auditor1", ""))),
            "auditor2": clean_val(row.get(mapping.get("auditor2", ""))),
            "time1": clean_val(row.get(mapping.get("time1", ""))),
            "time2": clean_val(row.get(mapping.get("time2", ""))),
            "phash": format_num(row.get(mapping.get("phash", ""))),
            "phashArea": clean_val(row.get(mapping.get("phash_area", ""))),
            "dhash": format_num(row.get(mapping.get("dhash", ""))),
            "orbInliers": format_num(row.get(mapping.get("orb_inliers", ""))),
            "orbRatio": format_num(row.get(mapping.get("orb_ratio", "")), 4),
        }

    @classmethod
    def make_package(
        cls,
        batch_data: pd.DataFrame,
        package_dir: Path,
        photo_root: Path,
        batch_id: str,
        mapping: Dict[str, str],
        max_px: int,
        quality: int,
        on_photo_progress: Optional[Any] = None,
    ) -> Tuple[int, int]:
        """打包单个复核包文件夹"""
        if package_dir.exists():
            shutil.rmtree(package_dir)
        photos_dir = package_dir / "photos"
        photos_dir.mkdir(parents=True, exist_ok=True)

        records = []
        missing = []

        total_rows = len(batch_data)
        for idx, (_, row) in enumerate(batch_data.iterrows()):
            src1 = cls.resolve_photo_path(photo_root, row, mapping, 1)
            src2 = cls.resolve_photo_path(photo_root, row, mapping, 2)
            name1 = Path(clean_val(row.get(mapping.get("file1", "")))).name or f"row_{idx}_1.jpg"
            name2 = Path(clean_val(row.get(mapping.get("file2", "")))).name or f"row_{idx}_2.jpg"

            complete = True
            rel_paths = []
            for side, (src, name) in enumerate([(src1, name1), (src2, name2)], start=1):
                if src is None or not src.is_file():
                    missing.append({
                        "批次": batch_id,
                        "序号": idx + 1,
                        "照片侧": side,
                        "文件名": name,
                        "任务号": clean_val(row.get(mapping.get(f"task{side}", ""))),
                    })
                    complete = False
                    rel_paths.append("")
                else:
                    dest = photos_dir / name
                    cls.prepare_photo(src, dest, max_px, quality)
                    rel_paths.append("photos/" + name)

            if complete:
                records.append(cls.build_record(row, mapping, rel_paths[0], rel_paths[1], idx))

            if on_photo_progress:
                on_photo_progress(idx + 1, total_rows)

        # 写入离线 HTML 与配置
        data_json = json.dumps(records, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
        batch_json = json.dumps(batch_id, ensure_ascii=False)
        html_content = HTML_TEMPLATE.replace("__BATCH_ID_JSON__", batch_json).replace("__CANDIDATES_JSON__", data_json)
        (package_dir / "开始复核.html").write_text(html_content, encoding="utf-8")
        (package_dir / "使用说明.txt").write_text(README_TEMPLATE, encoding="utf-8")

        if missing:
            pd.DataFrame(missing).to_csv(package_dir / "缺失照片.csv", index=False, encoding="utf-8-sig")

        return len(records), len(missing)

    @classmethod
    def zip_directory(cls, dir_path: Path) -> Path:
        """把指定目录压缩为同名 .zip 文件"""
        zip_path = dir_path.with_suffix(".zip")
        if zip_path.exists():
            zip_path.unlink()
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
            for p in dir_path.rglob("*"):
                if p.is_file():
                    archive.write(p, p.relative_to(dir_path.parent))
        return zip_path


# ==============================================================================
# GUI 现代化应用程序 (ReviewBuilderApp)
# ==============================================================================

class ReviewBuilderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("门头照复核包生成与规则筛选工作台 v3.0")
        self.root.geometry("1080x880")
        self.root.minsize(980, 720)

        # 数据状态
        self.inspect_data: Optional[Dict[str, Any]] = None
        self.raw_df: Optional[pd.DataFrame] = None
        self.filtered_df: Optional[pd.DataFrame] = None
        self.mapping_vars: Dict[str, tk.StringVar] = {}
        self.column_options: List[str] = ["(未选择)"]
        self.is_running = False
        self.cancel_requested = False

        self._setup_style()
        self._build_ui()
        self._init_defaults()

    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # 设定字体与颜色
        base_font = ("-apple-system", 12)
        bold_font = ("-apple-system", 12, "bold")
        head_font = ("-apple-system", 14, "bold")

        style.configure(".", font=base_font)
        style.configure("TLabel", font=base_font)
        style.configure("TButton", font=base_font, padding=4)
        style.configure("Header.TLabel", font=head_font, foreground="#1d1d1f")
        style.configure("Section.TLabelframe", padding=10)
        style.configure("Section.TLabelframe.Label", font=bold_font, foreground="#005fb8")
        style.configure("Action.TButton", font=bold_font, foreground="#ffffff", background="#0071e3")

    def _build_ui(self):
        # 整体采用带垂直滚动的 Canvas 容器，以适应各种屏幕高度
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient=tk.VERTICAL, command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas, padding=12)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas_window = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        def _on_canvas_configure(e):
            canvas.itemconfig(canvas_window, width=e.width)

        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.configure(yscrollcommand=scrollbar.set)

        # 鼠标滚轮绑定 (Mac & Windows)
        def _on_mousewheel(event):
            if sys.platform == "darwin":
                canvas.yview_scroll(int(-1 * event.delta), "units")
            else:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 顶部说明栏
        top_frame = ttk.Frame(self.scrollable_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(top_frame, text="门头照复核包生成与规则筛选工作台", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            top_frame,
            text="将相似候选数据按「相当商户判定」「时间窗口」「优先级」等规则筛选，并一键拆分打包为离线复核系统",
            foreground="#666666"
        ).pack(anchor="w", pady=(2, 0))

        # 1. 文件与目录配置卡片
        self._build_io_card()

        # 2. 原始数据列映射设置卡片
        self._build_mapping_card()

        # 3. 相当商户（同店）判断规则卡片
        self._build_merchant_rules_card()

        # 4. 时间与优先级过滤卡片
        self._build_time_priority_card()

        # 5. 分包与打包输出设置卡片
        self._build_package_settings_card()

        # 6. 控制按钮、进度条与日志区
        self._build_action_and_log_card()

    # --------------------------------------------------------------------------
    # 卡片 1: 文件与目录配置
    # --------------------------------------------------------------------------
    def _build_io_card(self):
        frame = ttk.LabelFrame(self.scrollable_frame, text="1. 文件与目录配置", style="Section.TLabelframe")
        frame.pack(fill=tk.X, pady=6)

        grid = ttk.Frame(frame)
        grid.pack(fill=tk.X, expand=True)
        grid.columnconfigure(1, weight=1)

        # 输入表格文件
        ttk.Label(grid, text="候选表格文件:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.var_input_file = tk.StringVar()
        entry_input = ttk.Entry(grid, textvariable=self.var_input_file)
        entry_input.grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(grid, text="浏览表格...", command=self._browse_input_file).grid(row=0, column=2, padx=4, pady=4)

        # Sheet 选择（针对 Excel）
        ttk.Label(grid, text="工作表 (Sheet):").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.var_sheet = tk.StringVar()
        self.combo_sheet = ttk.Combobox(grid, textvariable=self.var_sheet, state="readonly")
        self.combo_sheet.grid(row=1, column=1, sticky="w", padx=4, pady=4)
        self.combo_sheet.bind("<<ComboboxSelected>>", lambda e: self._on_sheet_changed())

        # 照片根目录
        ttk.Label(grid, text="照片存放目录:").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        self.var_photo_root = tk.StringVar()
        ttk.Entry(grid, textvariable=self.var_photo_root).grid(row=2, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(grid, text="选择目录...", command=self._browse_photo_root).grid(row=2, column=2, padx=4, pady=4)

        # 输出保存目录
        ttk.Label(grid, text="复核包输出目录:").grid(row=3, column=0, sticky="w", padx=4, pady=4)
        self.var_output_root = tk.StringVar()
        ttk.Entry(grid, textvariable=self.var_output_root).grid(row=3, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(grid, text="选择目录...", command=self._browse_output_root).grid(row=3, column=2, padx=4, pady=4)

    # --------------------------------------------------------------------------
    # 卡片 2: 原始数据列依据映射 (全自动适配 + 高级折叠微调)
    # --------------------------------------------------------------------------
    def _build_mapping_card(self):
        self.frame_mapping = ttk.LabelFrame(self.scrollable_frame, text="2. 原始数据列依据 (全自动识别)", style="Section.TLabelframe")
        self.frame_mapping.pack(fill=tk.X, pady=6)

        # 状态概览栏
        self.status_box = ttk.Frame(self.frame_mapping)
        self.status_box.pack(fill=tk.X, padx=4, pady=(2, 4))

        self.lbl_mapping_status = ttk.Label(
            self.status_box,
            text="✅ 原始数据列已全自动识别绑定",
            font=("-apple-system", 12, "bold"),
            foreground="#2f8f62"
        )
        self.lbl_mapping_status.pack(anchor="w")

        self.lbl_mapping_summary = ttk.Label(
            self.status_box,
            text="正在解析表格字段...",
            foreground="#555555"
        )
        self.lbl_mapping_summary.pack(anchor="w", pady=(2, 4))

        # 折叠控制按钮
        ctrl_row = ttk.Frame(self.frame_mapping)
        ctrl_row.pack(fill=tk.X, padx=4, pady=(2, 4))

        self.mapping_expanded = False
        self.btn_toggle_mapping = ttk.Button(
            ctrl_row,
            text="⚙ 展开高级字段映射微调 (通常无需修改)",
            command=self._toggle_mapping_view
        )
        self.btn_toggle_mapping.pack(side=tk.LEFT)

        ttk.Button(ctrl_row, text="↺ 重新推测映射", command=self._re_guess_mapping).pack(side=tk.RIGHT)

        # 高级映射网格（默认折叠隐藏，点击展开）
        self.grid_mapping_container = ttk.Frame(self.frame_mapping)

        mapping_fields = [
            ("商户号 1 (merchant_no_1):", "merchant_no_1", "商户号 2 (merchant_no_2):", "merchant_no_2"),
            ("商户名 1 (merchant_1):", "merchant_name_1", "商户名 2 (merchant_2):", "merchant_name_2"),
            ("照片名 1 (file1):", "file1", "照片名 2 (file2):", "file2"),
            ("完整路径 1 (完整路径1):", "path1", "完整路径 2 (完整路径2):", "path2"),
            ("提交时间 1 (submit_time_1):", "time1", "提交时间 2 (submit_time_2):", "time2"),
            ("任务号 1 (task_no_1):", "task1", "任务号 2 (task_no_2):", "task2"),
            ("地址 1 (import_address_1):", "address1", "地址 2 (import_address_2):", "address2"),
            ("作业员 1 (auditor_name_1):", "auditor1", "作业员 2 (auditor_name_2):", "auditor2"),
            ("时间差天数列 (时间差天):", "diff_days", "复核优先级列 (复核优先级):", "priority"),
            ("判定依据列 (判定依据):", "reason", "候选业务类型列 (候选业务类型):", "biz_type"),
            ("pHash距离 (pHash最小距离):", "phash", "ORB内点数 (ORB内点数):", "orb_inliers"),
            ("ORB比例 (ORB内点比例):", "orb_ratio", "候选组列 (候选组):", "group"),
        ]

        self.mapping_combos: Dict[str, ttk.Combobox] = {}

        for row_idx, (l1, key1, l2, key2) in enumerate(mapping_fields):
            # 左列
            ttk.Label(self.grid_mapping_container, text=l1).grid(row=row_idx, column=0, sticky="w", padx=(4, 2), pady=3)
            v1 = tk.StringVar(value="(未选择)")
            self.mapping_vars[key1] = v1
            cb1 = ttk.Combobox(self.grid_mapping_container, textvariable=v1, values=self.column_options, state="readonly", width=22)
            cb1.grid(row=row_idx, column=1, sticky="ew", padx=(2, 16), pady=3)
            self.mapping_combos[key1] = cb1

            # 右列
            ttk.Label(self.grid_mapping_container, text=l2).grid(row=row_idx, column=2, sticky="w", padx=(4, 2), pady=3)
            v2 = tk.StringVar(value="(未选择)")
            self.mapping_vars[key2] = v2
            cb2 = ttk.Combobox(self.grid_mapping_container, textvariable=v2, values=self.column_options, state="readonly", width=22)
            cb2.grid(row=row_idx, column=3, sticky="ew", padx=(2, 4), pady=3)
            self.mapping_combos[key2] = cb2

        self.grid_mapping_container.columnconfigure(1, weight=1)
        self.grid_mapping_container.columnconfigure(3, weight=1)

    def _toggle_mapping_view(self):
        if self.mapping_expanded:
            self.grid_mapping_container.pack_forget()
            self.btn_toggle_mapping.config(text="⚙ 展开高级字段映射微调 (通常无需修改)")
            self.mapping_expanded = False
        else:
            self.grid_mapping_container.pack(fill=tk.X, expand=True, padx=4, pady=(6, 4))
            self.btn_toggle_mapping.config(text="▲ 收起高级字段映射")
            self.mapping_expanded = True

    # --------------------------------------------------------------------------
    # 卡片 3: 相当商户（同店）判断规则
    # --------------------------------------------------------------------------
    def _build_merchant_rules_card(self):
        frame = ttk.LabelFrame(self.scrollable_frame, text="3. 「相当商户 (同店)」判断与过滤规则", style="Section.TLabelframe")
        frame.pack(fill=tk.X, pady=6)

        # 3.1 相当商户判定依据
        cond_box = ttk.LabelFrame(frame, text="相当商户判定依据 (满足任一项即视为相同门店)")
        cond_box.pack(fill=tk.X, padx=4, pady=4)

        cond_row = ttk.Frame(cond_box)
        cond_row.pack(fill=tk.X, padx=6, pady=4)

        self.var_match_no = tk.BooleanVar(value=True)
        self.var_match_name = tk.BooleanVar(value=True)
        self.var_match_addr = tk.BooleanVar(value=False)

        ttk.Checkbutton(cond_row, text="商户号一致 (去除空白后一致)", variable=self.var_match_no).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(cond_row, text="商户名一致 (去除空白后一致)", variable=self.var_match_name).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(cond_row, text="导入地址完全一致", variable=self.var_match_addr).pack(side=tk.LEFT, padx=10)

        # 3.2 过滤模式选择
        mode_box = ttk.LabelFrame(frame, text="相当商户处理策略")
        mode_box.pack(fill=tk.X, padx=4, pady=4)

        self.var_merchant_filter_mode = tk.StringVar(value="exclude_same")
        m_row = ttk.Frame(mode_box)
        m_row.pack(fill=tk.X, padx=6, pady=4)

        ttk.Radiobutton(
            m_row,
            text="排除相同商户（过滤同店，重点核查跨商户复用/门头造假）★推荐",
            variable=self.var_merchant_filter_mode,
            value="exclude_same"
        ).pack(anchor="w", pady=2)

        ttk.Radiobutton(
            m_row,
            text="仅保留相同商户（审查同商户跨期拍摄是否合规/偷懒）",
            variable=self.var_merchant_filter_mode,
            value="only_same"
        ).pack(anchor="w", pady=2)

        ttk.Radiobutton(
            m_row,
            text="全部保留（不做商户过滤，保留所有候选对）",
            variable=self.var_merchant_filter_mode,
            value="all"
        ).pack(anchor="w", pady=2)

        # 3.3 作业员关系过滤
        op_box = ttk.LabelFrame(frame, text="作业员维度筛选")
        op_box.pack(fill=tk.X, padx=4, pady=4)

        self.var_operator_mode = tk.StringVar(value="all")
        op_row = ttk.Frame(op_box)
        op_row.pack(fill=tk.X, padx=6, pady=4)

        ttk.Radiobutton(op_row, text="不限作业员关系", variable=self.var_operator_mode, value="all").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(op_row, text="仅限同一作业员（查个人多单复用）", variable=self.var_operator_mode, value="only_same_operator").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(op_row, text="仅限跨作业员（查不同人互借图库）", variable=self.var_operator_mode, value="only_diff_operator").pack(side=tk.LEFT, padx=10)

    # --------------------------------------------------------------------------
    # 卡片 4: 时间与优先级过滤
    # --------------------------------------------------------------------------
    def _build_time_priority_card(self):
        frame = ttk.LabelFrame(self.scrollable_frame, text="4. 「判断的时间」与优先级/算法指标规则", style="Section.TLabelframe")
        frame.pack(fill=tk.X, pady=6)

        # 时间规则子框
        time_box = ttk.LabelFrame(frame, text="时间窗口与时间差过滤 (若表中无时间差列，将自动根据两列提交时间计算)")
        time_box.pack(fill=tk.X, padx=4, pady=4)

        t_grid = ttk.Frame(time_box)
        t_grid.pack(fill=tk.X, padx=6, pady=4)

        ttk.Label(t_grid, text="最大时间差 (天):").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.var_max_diff_days = tk.StringVar(value="7")
        ttk.Entry(t_grid, textvariable=self.var_max_diff_days, width=8).grid(row=0, column=1, sticky="w", padx=4, pady=4)
        ttk.Label(t_grid, text="天内 (例如填 7 即 7 天窗口；留空表示不限)").grid(row=0, column=2, sticky="w", padx=4, pady=4)

        ttk.Label(t_grid, text="最小时间差 (天):").grid(row=0, column=3, sticky="w", padx=(16, 4), pady=4)
        self.var_min_diff_days = tk.StringVar(value="")
        ttk.Entry(t_grid, textvariable=self.var_min_diff_days, width=8).grid(row=0, column=4, sticky="w", padx=4, pady=4)
        ttk.Label(t_grid, text="天 (用于排除几秒内连拍；留空不限)").grid(row=0, column=5, sticky="w", padx=4, pady=4)

        # 绝对日期筛选
        ttk.Label(t_grid, text="任务提交起始日期:").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.var_start_date = tk.StringVar(value="")
        ttk.Entry(t_grid, textvariable=self.var_start_date, width=12).grid(row=1, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(t_grid, text="截止日期:").grid(row=1, column=2, sticky="e", padx=4, pady=4)
        self.var_end_date = tk.StringVar(value="")
        ttk.Entry(t_grid, textvariable=self.var_end_date, width=12).grid(row=1, column=3, sticky="w", padx=4, pady=4)
        ttk.Label(t_grid, text="(格式: YYYY-MM-DD，留空不限)").grid(row=1, column=4, columnspan=2, sticky="w", padx=4, pady=4)

        # 优先级与特征指标子框
        algo_box = ttk.LabelFrame(frame, text="复核优先级与算法特征指标过滤")
        algo_box.pack(fill=tk.X, padx=4, pady=4)

        a_row = ttk.Frame(algo_box)
        a_row.pack(fill=tk.X, padx=6, pady=4)

        ttk.Label(a_row, text="纳入复核优先级:").pack(side=tk.LEFT, padx=(0, 8))
        self.var_prio_high = tk.BooleanVar(value=True)
        self.var_prio_med = tk.BooleanVar(value=True)
        self.var_prio_low = tk.BooleanVar(value=False)
        self.var_prio_unverified = tk.BooleanVar(value=False)

        ttk.Checkbutton(a_row, text="高", variable=self.var_prio_high).pack(side=tk.LEFT, padx=6)
        ttk.Checkbutton(a_row, text="中", variable=self.var_prio_med).pack(side=tk.LEFT, padx=6)
        ttk.Checkbutton(a_row, text="低", variable=self.var_prio_low).pack(side=tk.LEFT, padx=6)
        ttk.Checkbutton(a_row, text="未验证", variable=self.var_prio_unverified).pack(side=tk.LEFT, padx=6)

        # 阈值
        th_row = ttk.Frame(algo_box)
        th_row.pack(fill=tk.X, padx=6, pady=(4, 6))

        self.var_phash_enabled = tk.BooleanVar(value=False)
        self.var_phash_val = tk.StringVar(value="10")
        ttk.Checkbutton(th_row, text="pHash 最小距离 <=", variable=self.var_phash_enabled).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(th_row, textvariable=self.var_phash_val, width=6).pack(side=tk.LEFT, padx=(0, 16))

        self.var_orb_inliers_enabled = tk.BooleanVar(value=False)
        self.var_orb_inliers_val = tk.StringVar(value="80")
        ttk.Checkbutton(th_row, text="ORB 内点数 >=", variable=self.var_orb_inliers_enabled).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(th_row, textvariable=self.var_orb_inliers_val, width=6).pack(side=tk.LEFT, padx=(0, 16))

        self.var_orb_ratio_enabled = tk.BooleanVar(value=False)
        self.var_orb_ratio_val = tk.StringVar(value="0.50")
        ttk.Checkbutton(th_row, text="ORB 内点比例 >=", variable=self.var_orb_ratio_enabled).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(th_row, textvariable=self.var_orb_ratio_val, width=6).pack(side=tk.LEFT, padx=(0, 4))

    # --------------------------------------------------------------------------
    # 卡片 5: 分包与打包输出设置
    # --------------------------------------------------------------------------
    def _build_package_settings_card(self):
        frame = ttk.LabelFrame(self.scrollable_frame, text="5. 分包与打包输出设置", style="Section.TLabelframe")
        frame.pack(fill=tk.X, pady=6)

        grid = ttk.Frame(frame)
        grid.pack(fill=tk.X, padx=6, pady=4)

        ttk.Label(grid, text="每个复核包候选数 (Batch Size):").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.var_batch_size = tk.StringVar(value="100")
        ttk.Entry(grid, textvariable=self.var_batch_size, width=8).grid(row=0, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(grid, text="候选总数上限限制 (Limit, 0不限):").grid(row=0, column=2, sticky="w", padx=(16, 4), pady=4)
        self.var_limit = tk.StringVar(value="0")
        ttk.Entry(grid, textvariable=self.var_limit, width=8).grid(row=0, column=3, sticky="w", padx=4, pady=4)

        ttk.Label(grid, text="打包图片最长边 (px):").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.var_max_px = tk.StringVar(value="1600")
        ttk.Entry(grid, textvariable=self.var_max_px, width=8).grid(row=1, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(grid, text="JPEG 压缩质量 (1-100):").grid(row=1, column=2, sticky="w", padx=(16, 4), pady=4)
        self.var_quality = tk.StringVar(value="82")
        ttk.Entry(grid, textvariable=self.var_quality, width=8).grid(row=1, column=3, sticky="w", padx=4, pady=4)

        opt_row = ttk.Frame(frame)
        opt_row.pack(fill=tk.X, padx=6, pady=(4, 6))

        self.var_zip_archive = tk.BooleanVar(value=True)
        self.var_export_filtered_excel = tk.BooleanVar(value=True)

        ttk.Checkbutton(opt_row, text="自动为每个复核包压缩为 .zip 文件", variable=self.var_zip_archive).pack(side=tk.LEFT, padx=(4, 16))
        ttk.Checkbutton(opt_row, text="同时导出《过滤后候选明细.xlsx》（含相当商户与时间差判定结果）", variable=self.var_export_filtered_excel).pack(side=tk.LEFT, padx=4)

    # --------------------------------------------------------------------------
    # 卡片 6: 操作、统计、进度与日志
    # --------------------------------------------------------------------------
    def _build_action_and_log_card(self):
        frame = ttk.LabelFrame(self.scrollable_frame, text="6. 规则计算、统计预览与一键执行", style="Section.TLabelframe")
        frame.pack(fill=tk.X, pady=6)

        # 统计面板
        self.stat_bar = ttk.Label(
            frame,
            text="[统计状态]: 请先加载表格文件，点击「统计匹配数量」可秒级预览规则筛选结果",
            foreground="#005fb8",
            font=("-apple-system", 12, "bold")
        )
        self.stat_bar.pack(anchor="w", padx=6, pady=4)

        # 按钮行
        btn_bar = ttk.Frame(frame)
        btn_bar.pack(fill=tk.X, padx=4, pady=6)

        self.btn_calc_stats = ttk.Button(btn_bar, text="📊 统计匹配数量 (毫秒级预估)", command=self._on_calc_stats_clicked)
        self.btn_calc_stats.pack(side=tk.LEFT, padx=4)

        self.btn_start = ttk.Button(btn_bar, text="🚀 开始生成复核包", style="Action.TButton", command=self._on_start_clicked)
        self.btn_start.pack(side=tk.LEFT, padx=8)

        self.btn_cancel = ttk.Button(btn_bar, text="⏹ 取消生成", state=tk.DISABLED, command=self._on_cancel_clicked)
        self.btn_cancel.pack(side=tk.LEFT, padx=4)

        ttk.Button(btn_bar, text="📂 打开输出目录", command=self._open_output_dir).pack(side=tk.RIGHT, padx=4)

        # 进度条
        p_frame = ttk.Frame(frame)
        p_frame.pack(fill=tk.X, padx=6, pady=4)
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(p_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.lbl_progress = ttk.Label(p_frame, text="0 / 0", width=16)
        self.lbl_progress.pack(side=tk.RIGHT)

        # 日志文本框
        log_box = ttk.Frame(frame)
        log_box.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        self.log_text = tk.Text(log_box, height=8, wrap="word", bg="#1e1e1e", fg="#e0e0e0", insertbackground="#fff")
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll = ttk.Scrollbar(log_box, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=log_scroll.set)

    def log(self, message: str):
        now = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{now}] {message}\n")
        self.log_text.see(tk.END)

    def _init_defaults(self):
        """初始化预设路径（优先寻找两两相似候选表，其次寻找合并表）"""
        try:
            base_dir = Path(__file__).resolve().parent

            # 1. 优先寻找桌面或当前目录中的两两相似候选大表
            desktop_candidates = [
                Path("/Users/jun/Desktop/照片相似候选_100k_7天窗口.xlsx"),
                Path("/Users/jun/Desktop/照片相似候选_100k.xlsx"),
            ]
            chosen_f = None
            for cand_p in desktop_candidates:
                if cand_p.is_file():
                    chosen_f = cand_p
                    break

            # 2. 其次在当前工作目录查找包含“候选”或“相似”的表
            if chosen_f is None:
                local_xlsx = [f for f in base_dir.glob("*.xlsx") if not f.name.startswith("~$") and not f.name.startswith(".~")]
                for f in local_xlsx:
                    if "候选" in f.name or "相似" in f.name:
                        chosen_f = f
                        break
                if chosen_f is None and local_xlsx:
                    chosen_f = local_xlsx[0]

            if chosen_f and chosen_f.is_file():
                self.var_input_file.set(str(chosen_f))
                self._load_file_metadata(str(chosen_f), silent=True)

            # 照片目录
            local_photos = base_dir / "pos-jianhang2_已合并_photos"
            if local_photos.is_dir():
                self.var_photo_root.set(str(local_photos))
            else:
                self.var_photo_root.set(str(base_dir))

            # 输出目录
            default_out = base_dir / "review_packages"
            self.var_output_root.set(str(default_out))

            self.log("程序初始化完成，已自动就绪。")
        except Exception as e:
            self.log(f"预加载默认配置时提示: {e}")

    # --------------------------------------------------------------------------
    # 文件浏览与元数据解析
    # --------------------------------------------------------------------------
    def _browse_input_file(self):
        f = filedialog.askopenfilename(
            title="选择候选数据表格",
            filetypes=[("表格文件", "*.xlsx *.xls *.csv *.tsv *.txt"), ("所有文件", "*.*")]
        )
        if f:
            self.var_input_file.set(f)
            self._load_file_metadata(f, silent=False)

    def _browse_photo_root(self):
        d = filedialog.askdirectory(title="选择照片根目录")
        if d:
            self.var_photo_root.set(d)

    def _browse_output_root(self):
        d = filedialog.askdirectory(title="选择复核包输出目录")
        if d:
            self.var_output_root.set(d)

    def _load_file_metadata(self, file_path: str, silent: bool = False):
        if not file_path or not os.path.isfile(file_path):
            return

        self.lbl_mapping_status.config(text="⏳ 正在后台解析表格列结构...", foreground="#005fb8")
        self.lbl_mapping_summary.config(text=f"正在分析文件: {Path(file_path).name} ...")
        self.raw_df = None

        def _worker():
            try:
                inspect_data = ReviewPackageEngine.inspect_file(file_path)
                sheets = inspect_data.get("sheets", [])
                best_sheet = "(默认)"
                if sheets:
                    best_sheet = sheets[0]
                    for s in sheets:
                        if "相似" in s or "候选" in s:
                            best_sheet = s
                            break

                cols = inspect_data.get("columns", [])
                if best_sheet and sheets and best_sheet != sheets[0]:
                    cols = ReviewPackageEngine.get_columns(file_path, best_sheet)

                def _ui_callback():
                    self.inspect_data = inspect_data
                    if sheets:
                        self.combo_sheet["values"] = sheets
                        self.var_sheet.set(best_sheet)
                        self.combo_sheet.config(state="readonly")
                    else:
                        self.combo_sheet["values"] = ["(默认)"]
                        self.var_sheet.set("(默认)")
                        self.combo_sheet.config(state="disabled")

                    self.column_options = ["(未选择)"] + cols
                    for cb in self.mapping_combos.values():
                        cb["values"] = self.column_options

                    self._re_guess_mapping()
                    self.log(f"成功读取表格「{Path(file_path).name}」列信息: 共 {len(cols)} 列。")

                self.root.after(0, _ui_callback)
            except Exception as e:
                def _err_callback():
                    self.lbl_mapping_status.config(text="⚠️ 表格解析异常", foreground="#d9554b")
                    self.lbl_mapping_summary.config(text=str(e))
                    if not silent:
                        messagebox.showerror("文件读取失败", f"无法解析所选文件元数据：\n{e}")
                    else:
                        self.log(f"预解析表格元数据跳过: {e}")
                self.root.after(0, _err_callback)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_sheet_changed(self):
        file_path = self.var_input_file.get()
        sheet = self.var_sheet.get()
        if not file_path or not os.path.isfile(file_path):
            return

        self.lbl_mapping_status.config(text="⏳ 正在切换并解析工作表...", foreground="#005fb8")

        def _worker():
            try:
                cols = ReviewPackageEngine.get_columns(file_path, sheet)
                def _ui():
                    self.column_options = ["(未选择)"] + cols
                    for cb in self.mapping_combos.values():
                        cb["values"] = self.column_options
                    self._re_guess_mapping()
                    self.raw_df = None
                    self.log(f"已切换工作表至「{sheet}」，共 {len(cols)} 列。")
                self.root.after(0, _ui)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Sheet读取失败", str(e)))

        threading.Thread(target=_worker, daemon=True).start()

    def _re_guess_mapping(self):
        cols = [c for c in self.column_options if c != "(未选择)"]
        if not cols:
            self.lbl_mapping_status.config(text="⚠️ 暂未获取到列头信息", foreground="#c66a25")
            self.lbl_mapping_summary.config(text="请选择有效表格文件")
            return

        guessed = ReviewPackageEngine.guess_column_mapping(cols)
        for key, var in self.mapping_vars.items():
            best = guessed.get(key, "")
            var.set(best if best in cols else "(未选择)")

        is_pair = guessed.get("is_pair_table") == "1"
        mno = guessed.get("merchant_no_1") or guessed.get("merchant_name_1") or "—"
        time_c = guessed.get("time1") or guessed.get("diff_days") or "—"
        task_c = guessed.get("task1") or "—"
        file_c = guessed.get("file1") or "—"

        if is_pair:
            self.lbl_mapping_status.config(
                text="✅ 自动识别为【两两相似候选表】（字段已 100% 全自动绑定就绪，无需手动操作）",
                foreground="#2f8f62"
            )
            self.lbl_mapping_summary.config(
                text=f"已自动绑定：商户 [{mno}] ｜ 时间 [{time_c}] ｜ 任务 [{task_c}] ｜ 照片 [{file_c}] ｜ 算法特征指标全部就绪"
            )
        else:
            self.lbl_mapping_status.config(
                text="ℹ️ 自动识别为【原始任务明细单表】（单条记录，已自动抓取基准列）",
                foreground="#005fb8"
            )
            self.lbl_mapping_summary.config(
                text=f"已自动抓取基准列：商户号 [{mno}] ｜ 提交时间 [{time_c}] ｜ 任务号 [{task_c}] ｜ 照片链接 [{file_c}]"
            )

    def _gather_mapping(self) -> Dict[str, str]:
        mapping = {}
        for k, v in self.mapping_vars.items():
            val = v.get().strip()
            mapping[k] = "" if val == "(未选择)" else val
        return mapping

    def _gather_rules(self) -> Dict[str, Any]:
        priorities = []
        if self.var_prio_high.get():
            priorities.append("高")
        if self.var_prio_med.get():
            priorities.append("中")
        if self.var_prio_low.get():
            priorities.append("低")
        if self.var_prio_unverified.get():
            priorities.append("未验证")

        def _float_or_none(s: str) -> Optional[float]:
            s = s.strip()
            try:
                return float(s) if s else None
            except Exception:
                return None

        def _int_or_zero(s: str) -> int:
            try:
                return int(s.strip())
            except Exception:
                return 0

        rules = {
            "same_merchant_by_no": self.var_match_no.get(),
            "same_merchant_by_name": self.var_match_name.get(),
            "same_merchant_by_addr": self.var_match_addr.get(),
            "merchant_filter_mode": self.var_merchant_filter_mode.get(),
            "operator_mode": self.var_operator_mode.get(),
            "max_diff_days": _float_or_none(self.var_max_diff_days.get()),
            "min_diff_days": _float_or_none(self.var_min_diff_days.get()),
            "start_date": self.var_start_date.get().strip(),
            "end_date": self.var_end_date.get().strip(),
            "priorities": priorities,
            "phash_max": _float_or_none(self.var_phash_val.get()) if self.var_phash_enabled.get() else None,
            "orb_inliers_min": _float_or_none(self.var_orb_inliers_val.get()) if self.var_orb_inliers_enabled.get() else None,
            "orb_ratio_min": _float_or_none(self.var_orb_ratio_val.get()) if self.var_orb_ratio_enabled.get() else None,
            "limit": _int_or_zero(self.var_limit.get()),
        }
        return rules

    # --------------------------------------------------------------------------
    # 统计匹配数量
    # --------------------------------------------------------------------------
    def _on_calc_stats_clicked(self):
        file_path = self.var_input_file.get().strip()
        if not file_path or not os.path.isfile(file_path):
            messagebox.showwarning("提示", "请先选择有效的输入表格文件！")
            return

        self.btn_calc_stats.config(state=tk.DISABLED)
        self.stat_bar.config(text="正在读取并计算统计，请稍候...")
        self.log("开始进行规则匹配统计计算...")

        def _worker():
            try:
                if self.raw_df is None:
                    sheet = self.var_sheet.get()
                    if sheet == "(默认)":
                        sheet = None
                    self.raw_df = ReviewPackageEngine.read_dataset(file_path, sheet)

                mapping = self._gather_mapping()
                rules = self._gather_rules()

                if mapping.get("is_pair_table") != "1" and (mapping.get("file1") == mapping.get("file2") or not mapping.get("file2")):
                    def _warn_single_stat():
                        self.stat_bar.config(text="⚠️ 提示：当前选择的是【单任务原始表】，需使用【两两相似候选表】")
                        self.log("提示：当前表格为【单任务原始表】（单条记录），无法直接进行两两复核打包。")
                        self.btn_calc_stats.config(state=tk.NORMAL)
                        messagebox.showinfo(
                            "需要选择「相似候选表」",
                            "您当前选择的是【原始任务明细单表】（每行仅有一条任务的照片）。\n\n"
                            "门头照复核系统需要输入【两两相似候选表】（包含 照片1 ↔ 照片2 对比数据的表格，例如桌面的《照片相似候选_100k_7天窗口.xlsx》）。\n\n"
                            "请在上方点击「浏览表格...」选择对应的相似候选表。"
                        )
                    self.root.after(0, _warn_single_stat)
                    return

                filtered, stats = ReviewPackageEngine.apply_filter_rules(self.raw_df, mapping, rules)
                self.filtered_df = filtered

                batch_size = max(1, int(self.var_batch_size.get().strip() or 100))
                expected_batches = math.ceil(len(filtered) / batch_size) if len(filtered) > 0 else 0

                msg = (
                    f"【统计结果】输入总数: {stats['total_input']:,} 条 | "
                    f"相当商户: {stats['same_merchant_count']:,} 条 | 跨商户: {stats['cross_merchant_count']:,} 条 | "
                    f"最终符合规则将打包: {stats['final_count']:,} 条 | "
                    f"预计生成: {expected_batches} 个复核包"
                )

                def _update_ui():
                    self.stat_bar.config(text=msg)
                    self.log(msg)
                    self.btn_calc_stats.config(state=tk.NORMAL)

                self.root.after(0, _update_ui)
            except Exception as e:
                def _err():
                    self.stat_bar.config(text="统计计算失败，请检查列映射与参数。")
                    self.log(f"统计计算异常: {e}")
                    self.btn_calc_stats.config(state=tk.NORMAL)
                    messagebox.showerror("统计失败", str(e))
                self.root.after(0, _err)

        threading.Thread(target=_worker, daemon=True).start()

    # --------------------------------------------------------------------------
    # 开始生成复核包
    # --------------------------------------------------------------------------
    def _on_start_clicked(self):
        file_path = self.var_input_file.get().strip()
        photo_root = Path(self.var_photo_root.get().strip())
        output_root = Path(self.var_output_root.get().strip())

        if not file_path or not os.path.isfile(file_path):
            messagebox.showwarning("提示", "请选择有效的候选表格文件！")
            return
        if not photo_root.is_dir():
            messagebox.showwarning("提示", f"照片目录不存在：{photo_root}\n请先选择或创建照片存放目录！")
            return

        try:
            batch_size = max(1, int(self.var_batch_size.get().strip() or 100))
            max_px = max(200, int(self.var_max_px.get().strip() or 1600))
            quality = min(100, max(10, int(self.var_quality.get().strip() or 82)))
        except ValueError:
            messagebox.showwarning("参数错误", "批次大小、图片尺寸或质量请输入有效正整数！")
            return

        self.is_running = True
        self.cancel_requested = False
        self.btn_start.config(state=tk.DISABLED)
        self.btn_calc_stats.config(state=tk.DISABLED)
        self.btn_cancel.config(state=tk.NORMAL)
        self.progress_var.set(0)
        self.lbl_progress.config(text="正在加载与过滤...")

        def _worker():
            try:
                # 1. 准备数据
                if self.raw_df is None:
                    self.log(f"正在读取表格文件: {file_path} ...")
                    sheet = self.var_sheet.get()
                    if sheet == "(默认)":
                        sheet = None
                    self.raw_df = ReviewPackageEngine.read_dataset(file_path, sheet)

                mapping = self._gather_mapping()
                rules = self._gather_rules()

                if mapping.get("is_pair_table") != "1" and (mapping.get("file1") == mapping.get("file2") or not mapping.get("file2")):
                    def _warn_single_start():
                        self.log("提示：当前表格为【单任务原始表】，无法直接进行两两复核打包。")
                        self._reset_buttons()
                        messagebox.showinfo(
                            "需要选择「相似候选表」",
                            "您当前选择的是【原始任务明细单表】。\n\n"
                            "门头照复核系统需要输入【两两相似候选表】（包含 照片1 ↔ 照片2 对比数据的表格，例如桌面的《照片相似候选_100k_7天窗口.xlsx》）。\n\n"
                            "请在上方点击「浏览表格...」选择对应的相似候选表。"
                        )
                    self.root.after(0, _warn_single_start)
                    return

                self.log("正在应用规则筛选数据...")
                data, stats = ReviewPackageEngine.apply_filter_rules(self.raw_df, mapping, rules)
                self.filtered_df = data

                total_items = len(data)
                if total_items == 0:
                    raise RuntimeError("筛选后没有任何符合条件的候选数据！请调整判定规则或优先级后重试。")

                output_root.mkdir(parents=True, exist_ok=True)

                # 导出过滤后候选明细 Excel
                if self.var_export_filtered_excel.get():
                    filtered_xlsx = output_root / "过滤后候选明细.xlsx"
                    self.log(f"正在保存过滤后的候选明细至: {filtered_xlsx.name} ...")
                    data.to_excel(filtered_xlsx, index=False)

                batches = math.ceil(total_items / batch_size)
                self.log(f"筛选完成！符合条件总数: {total_items} 条，将拆分为 {batches} 个复核包。")

                summary_list = []
                total_processed = 0

                for b_idx in range(batches):
                    if self.cancel_requested:
                        self.log("用户取消了打包操作。")
                        break

                    start_idx = b_idx * batch_size
                    end_idx = min(start_idx + batch_size, total_items)
                    batch_data = data.iloc[start_idx:end_idx]

                    batch_id = f"门头照复核_{b_idx + 1:03d}"
                    pkg_dir = output_root / batch_id

                    def _on_photo(cur, tot):
                        nonlocal total_processed
                        total_processed = start_idx + cur
                        pct = (total_processed / total_items) * 100
                        self.root.after(0, lambda: self._update_progress(pct, f"{total_processed} / {total_items}"))

                    inc, mis = ReviewPackageEngine.make_package(
                        batch_data,
                        pkg_dir,
                        photo_root,
                        batch_id,
                        mapping,
                        max_px,
                        quality,
                        on_photo_progress=_on_photo,
                    )

                    zip_path_str = ""
                    if self.var_zip_archive.get():
                        z_path = ReviewPackageEngine.zip_directory(pkg_dir)
                        zip_path_str = str(z_path)

                    summary_list.append({
                        "批次": batch_id,
                        "候选输入": len(batch_data),
                        "成功打包": inc,
                        "缺失照片": mis,
                        "目录": str(pkg_dir),
                        "ZIP压缩包": zip_path_str,
                    })

                    self.log(f"[{batch_id}] 生成完毕: 成功 {inc} 对, 缺失 {mis} 张图片" + (f", 已压缩为 ZIP" if zip_path_str else ""))

                # 生成复核包生成汇总.xlsx
                if summary_list:
                    summary_df = pd.DataFrame(summary_list)
                    summary_path = output_root / "复核包生成汇总.xlsx"
                    summary_df.to_excel(summary_path, index=False)
                    self.log(f"已生成打包总清单: {summary_path}")

                first_html = output_root / "门头照复核_001" / "开始复核.html"

                def _finish_ui():
                    self.progress_var.set(100)
                    self.lbl_progress.config(text=f"{total_processed} / {total_items}")
                    self.log("🎉 所有复核包生成完毕！")
                    self._on_worker_finished(output_root, first_html if first_html.exists() else None)

                self.root.after(0, _finish_ui)

            except Exception as e:
                def _err_ui():
                    self.log(f"❌ 生成过程中出错: {e}")
                    messagebox.showerror("生成失败", f"处理时发生错误：\n{e}")
                    self._reset_buttons()

                self.root.after(0, _err_ui)

        threading.Thread(target=_worker, daemon=True).start()

    def _update_progress(self, percent: float, text: str):
        self.progress_var.set(percent)
        self.lbl_progress.config(text=text)

    def _on_cancel_clicked(self):
        if self.is_running:
            self.cancel_requested = True
            self.btn_cancel.config(state=tk.DISABLED)
            self.log("已请求中止，将在完成当前任务后安全退出...")

    def _reset_buttons(self):
        self.is_running = False
        self.btn_start.config(state=tk.NORMAL)
        self.btn_calc_stats.config(state=tk.NORMAL)
        self.btn_cancel.config(state=tk.DISABLED)

    def _on_worker_finished(self, output_root: Path, first_html: Optional[Path]):
        self._reset_buttons()
        resp = messagebox.askyesno(
            "复核包生成成功！",
            f"所有复核包已成功生成并保存至：\n{output_root}\n\n是否立即在浏览器中打开首个复核页面测试？"
        )
        if resp and first_html and first_html.is_file():
            if sys.platform == "darwin":
                subprocess.Popen(["open", str(first_html)])
            elif sys.platform.startswith("win"):
                os.startfile(str(first_html))
            else:
                subprocess.Popen(["xdg-open", str(first_html)])

    def _open_output_dir(self):
        out = self.var_output_root.get().strip()
        if not out or not os.path.isdir(out):
            messagebox.showinfo("提示", "输出目录尚未生成或不存在。")
            return
        if sys.platform == "darwin":
            subprocess.Popen(["open", out])
        elif sys.platform.startswith("win"):
            os.startfile(out)
        else:
            subprocess.Popen(["xdg-open", out])


# ==============================================================================
# 程序入口
# ==============================================================================

def main():
    root = tk.Tk()
    app = ReviewBuilderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
