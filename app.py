# -*- coding: utf-8 -*-
"""
SVOD TIZIMI v1.3
Mahalliy (local) CSV -> nazorat -> Excel svod tizimi.
Manba fayllar o'zgartirilmaydi va o'chirilmaydi.
"""
from __future__ import annotations

import csv
import hashlib
import io
import json
import logging
import math
import os
import re
import socket
import sys
import tempfile
import threading
import urllib.request
import webbrowser
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from functools import lru_cache
from pathlib import Path
from typing import Dict, List, Tuple
from xml.etree import ElementTree as ET

from flask import Flask, jsonify, redirect, render_template_string, request, send_file
from werkzeug.exceptions import HTTPException
import xlsxwriter
from reference import BUCKETS, DISPLAY_BUCKETS, parse_reference_xlsx

APP_NAME = "SVOD TIZIMI"
APP_VERSION = "1.3"
PORT = 8765
MAX_UPLOAD_MB = 350
HOSTED_UPLOAD_MB = 4
HOSTED = bool(os.environ.get("VERCEL") or os.environ.get("VERCEL_ENV"))

BASE_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
LOG_DIR = (
    Path(os.environ.get("LOCALAPPDATA", BASE_DIR)) / "SVOD_TIZIMI" / "logs"
    if getattr(sys, "frozen", False)
    else BASE_DIR / "logs"
)
log_handler = logging.StreamHandler(sys.stderr)
if not HOSTED:
    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        log_handler = logging.FileHandler(LOG_DIR / "app.log", encoding="utf-8")
    except OSError:
        # Serverless deployments cannot write beside app.py. Keep startup working.
        pass
logging.basicConfig(
    handlers=[log_handler],
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

REQUIRED_HEADERS = [
    "Вилоят",
    "Райгаз",
    "Маҳалла",
    "Абонент код",
    "Абонент",
    "Еҳтиёж",
    "Сўнги реализация",
    "Бугунги реализация",
]

INTERVALS = ["0-30 (НАЗОРАТ)", "31-35", "36-40", "41-45", "46-50",
             "51-55", "56-60", "61-65", "66-70", "71+"]

DATE_FORMATS = ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y", "%Y.%m.%d")
XLSX_NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

# Excel varag'idagi maksimal qatorlar soni (sarlavha bilan).
EXCEL_MAX_ROWS = 1_048_576

# Sarlavhalarni solishtirishda imlo farqlarini tekislash:
# "Эҳтиёж" / "Еҳтиёж", "Маҳалла" / "Махалла", "Сўнги" / "Сунги" va h.k.
HEADER_FOLD = str.maketrans({"э": "е", "ё": "е", "ҳ": "х", "қ": "к", "ғ": "г", "ў": "у"})

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = (HOSTED_UPLOAD_MB if HOSTED else MAX_UPLOAD_MB) * 1024 * 1024


@app.after_request
def protect_response(response):
    """Uploaded operational data must not be cached, sniffed, or framed."""
    response.headers["Cache-Control"] = "no-store, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; base-uri 'none'; form-action 'self'; frame-ancestors 'none'; "
        "object-src 'none'; img-src 'self' data:; style-src 'self' 'unsafe-inline'; "
        "script-src 'self' 'unsafe-inline'"
    )
    return response


def hosted_request() -> bool:
    return HOSTED or request.host.split(":", 1)[0] not in ("127.0.0.1", "localhost")


HTML = r"""
<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#102f45">
<title>SVOD TIZIMI</title>
<style>
:root {
  --navy:#102f45;--navy-2:#174d66;--blue:#137a8b;--green:#15805d;--green-dark:#0d6549;
  --mint:#e9f8f2;--light:#f3f7f8;--surface:#ffffff;--ink:#142b36;--muted:#637b86;
  --border:#d8e4e7;--border-strong:#b8ced3;--danger:#b42318;--warn:#9a5b09;
  --shadow:0 24px 70px rgba(16,47,69,.12);--shadow-soft:0 10px 30px rgba(16,47,69,.08);
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;min-height:100vh;background:
  radial-gradient(circle at 8% 0%,rgba(19,122,139,.12),transparent 30rem),
  radial-gradient(circle at 92% 18%,rgba(21,128,93,.10),transparent 26rem),var(--light);
  font-family:Inter,"Segoe UI",Arial,sans-serif;color:var(--ink);line-height:1.5}
button,input{font:inherit}
.top{position:relative;overflow:hidden;background:linear-gradient(125deg,#0c293d 0%,var(--navy-2) 58%,#126e70 100%);color:white;padding:31px 18px 33px;border-bottom:1px solid rgba(255,255,255,.12)}
.top::after{content:"";position:absolute;width:360px;height:360px;border:70px solid rgba(255,255,255,.045);border-radius:50%;right:-100px;top:-210px;pointer-events:none}
.top-inner{position:relative;z-index:1;max-width:1120px;margin:auto;display:flex;align-items:center;justify-content:space-between;gap:28px}
.brand{min-width:0;display:flex;align-items:center;gap:15px}
.brand-mark{width:48px;height:48px;display:grid;place-items:center;flex:0 0 auto;border-radius:15px;background:rgba(255,255,255,.13);border:1px solid rgba(255,255,255,.22);box-shadow:inset 0 1px 0 rgba(255,255,255,.18)}
.brand-mark svg{width:25px;height:25px}.brand-copy{min-width:0}
.eyebrow{font-size:11px;font-weight:800;letter-spacing:.16em;text-transform:uppercase;color:#84e0c4;margin-bottom:3px}
h1{margin:0;font-size:clamp(27px,4vw,34px);line-height:1.1;letter-spacing:-.025em}.sub{opacity:.83;font-size:14px;margin-top:6px}
.nav{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end}
.nav a{color:rgba(255,255,255,.88);text-decoration:none;font-weight:750;font-size:13px;letter-spacing:.025em;padding:10px 14px;border:1px solid rgba(255,255,255,.22);border-radius:11px;white-space:nowrap;transition:transform .2s,background .2s,border-color .2s,color .2s}
.nav a:hover{transform:translateY(-2px);border-color:rgba(255,255,255,.48)}.nav a.active{background:white;color:var(--navy);border-color:white;box-shadow:0 7px 20px rgba(0,0,0,.14)}
.wrap{max-width:1120px;margin:30px auto;padding:0 18px}
.card{background:rgba(255,255,255,.94);border:1px solid rgba(184,206,211,.8);border-radius:24px;padding:clamp(18px,3vw,30px);box-shadow:var(--shadow);backdrop-filter:blur(12px);animation:liftIn .55s cubic-bezier(.2,.8,.2,1) both}
.workspace-head{display:flex;align-items:flex-start;justify-content:space-between;gap:24px;margin-bottom:20px}
.workspace-head h2{font-size:clamp(21px,3vw,27px);letter-spacing:-.025em;margin:0 0 5px}.workspace-head p{margin:0;max-width:680px}
.trust-chip{display:flex;align-items:center;gap:8px;flex:0 1 auto;max-width:430px;background:var(--mint);color:var(--green-dark);border:1px solid #bde6d8;border-radius:999px;padding:8px 12px;font-size:12px;font-weight:800;line-height:1.35}
.trust-dot{width:8px;height:8px;border-radius:50%;background:var(--green);box-shadow:0 0 0 4px rgba(21,128,93,.12)}
.drop{position:relative;overflow:hidden;border:2px dashed #8eb4bb;border-radius:18px;padding:40px 22px;text-align:center;cursor:pointer;background:linear-gradient(180deg,#fbfefe,#f5fafb);transition:transform .22s,border-color .22s,background .22s,box-shadow .22s;outline:none}
.drop::before{content:"";position:absolute;inset:0;background:linear-gradient(120deg,transparent 30%,rgba(255,255,255,.7),transparent 70%);transform:translateX(-110%);transition:transform .7s;pointer-events:none}
.drop:hover{transform:translateY(-2px);border-color:var(--blue);box-shadow:var(--shadow-soft)}.drop:hover::before{transform:translateX(110%)}
.drop.drag{transform:scale(1.008);border-color:var(--green);background:var(--mint);box-shadow:0 0 0 5px rgba(21,128,93,.09)}
.upload-icon{width:58px;height:58px;margin:0 auto 14px;display:grid;place-items:center;border-radius:18px;background:linear-gradient(145deg,#dff4f1,#edf8fb);color:var(--green);box-shadow:inset 0 0 0 1px rgba(21,128,93,.12)}
.upload-icon svg{width:29px;height:29px}.drop strong{font-size:20px;display:block;margin-bottom:7px;color:var(--navy)}
.drop-actions{position:relative;display:flex;justify-content:center;gap:10px;margin-top:15px;flex-wrap:wrap}
.muted{color:var(--muted);font-size:14px}
#fileInput{display:none}
.reference{margin-top:18px;padding:17px 18px;border:1px solid var(--border);border-radius:15px;background:#f8fbfc;transition:border-color .2s,background .2s,transform .2s}
.reference.drag{border-color:var(--green);background:var(--mint);transform:translateY(-1px)}
.reference-title{display:flex;align-items:center;gap:10px;margin-bottom:5px}.reference-badge{display:inline-grid;place-items:center;width:26px;height:26px;border-radius:8px;background:#e5f1f4;color:var(--blue);font-size:14px;font-weight:900}
.reference label{font-weight:800;display:block;color:var(--navy)}
.reference input{max-width:100%;color:var(--muted)}
.reference-row{display:flex;align-items:center;gap:10px;margin-top:12px;flex-wrap:wrap}
.selected-reference{font-weight:750;color:var(--green-dark);overflow-wrap:anywhere;margin-top:9px}
.file-summary{display:none;align-items:center;justify-content:space-between;gap:15px;margin-top:19px;padding:0 3px 9px;border-bottom:1px solid var(--border);font-size:13px;color:var(--muted)}.file-summary.show{display:flex}.file-summary b{color:var(--navy)}
.filelist{max-height:250px;overflow:auto}
.file{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:11px 4px;border-bottom:1px solid #eaf0f2;font-size:14px;animation:rowIn .25s ease both}
.file-name{min-width:0;overflow-wrap:anywhere}.file-tools{display:flex;align-items:center;gap:10px;white-space:nowrap}
.actions{display:flex;gap:11px;align-items:center;margin-top:21px;flex-wrap:wrap}
button{border:0;border-radius:11px;padding:12px 18px;font-size:15px;font-weight:800;cursor:pointer;transition:transform .18s,box-shadow .18s,background .18s;color .18s}
button:hover:not(:disabled){transform:translateY(-2px)}button:active:not(:disabled){transform:translateY(0)}
.primary{background:linear-gradient(135deg,var(--green),#1a9470);color:white;box-shadow:0 8px 18px rgba(21,128,93,.22)}.primary:hover:not(:disabled){box-shadow:0 12px 24px rgba(21,128,93,.28)}.primary:disabled{opacity:.46;cursor:not-allowed;box-shadow:none}
.secondary{background:#e8f0f2;color:#244d5b}.secondary:hover:not(:disabled){background:#dce9ec}
.small{padding:7px 11px;font-size:13px}.remove{background:#fff0ee;color:var(--danger)}.remove:hover:not(:disabled){background:#ffe3df}
.status{position:relative;margin-top:18px;padding:14px 16px 14px 44px;border-radius:12px;display:none;white-space:pre-wrap;animation:statusIn .25s ease both}
.status::before{position:absolute;left:16px;top:14px;font-weight:900}.status.ok{display:block;background:#eaf8f1;color:#126342;border:1px solid #b9e2d0}.status.ok::before{content:"✓"}
.status.err{display:block;background:#fff1ef;color:var(--danger);border:1px solid #f0c9c4}.status.err::before{content:"!"}
.status.work{display:block;background:#fff8e9;color:var(--warn);border:1px solid #efd89f}.status.work::before{content:"";width:15px;height:15px;border:2px solid currentColor;border-right-color:transparent;border-radius:50%;top:16px;animation:spin .75s linear infinite}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:18px}
.kpi{background:linear-gradient(145deg,#f9fcfc,#f1f7f8);border:1px solid var(--border);border-radius:14px;padding:15px;animation:rowIn .35s ease both}.kpi:nth-child(2){animation-delay:.05s}.kpi:nth-child(3){animation-delay:.1s}.kpi:nth-child(4){animation-delay:.15s}
.kpi span{font-size:12px;color:var(--muted);font-weight:700}.kpi b{font-size:25px;line-height:1.15;display:block;color:var(--navy);margin-top:4px;font-variant-numeric:tabular-nums}
.notice{margin-top:20px;padding:15px 17px;background:#edf6f8;border:1px solid #d2e7eb;border-left:4px solid var(--blue);border-radius:11px;font-size:13px;line-height:1.55;color:#365965}
.footer{color:#718892;font-size:12px;text-align:center;margin:19px}.footer strong{color:#46636e}
/* Minimal, compact workspace */
.top{padding:19px 18px 20px;background:linear-gradient(115deg,#102f45,#15556a 62%,#146f70)}
.top::after{display:none}.top-inner{max-width:1180px;gap:20px}.brand{gap:0}.brand-mark{display:none}
.eyebrow{font-size:10px;letter-spacing:.2em;margin-bottom:2px}h1{font-size:29px}.sub{font-size:13px;margin-top:4px}
.nav{gap:12px;padding:0;background:transparent;border:0;border-radius:0}
.nav a{color:#fff;border:1px solid rgba(255,255,255,.34);border-radius:10px;padding:10px 14px;font-size:12px;box-shadow:0 5px 14px rgba(4,28,40,.2)}
.nav a:nth-child(1){background:#2563eb}.nav a:nth-child(2){background:#d97706}.nav a:nth-child(3){background:#15803d}
.nav a:hover{transform:translateY(-1px);filter:brightness(1.1)}.nav a.active{color:#fff;border-color:#fff;box-shadow:0 0 0 2px rgba(255,255,255,.25),0 6px 16px rgba(4,28,40,.28)}
.wrap{max-width:1180px;margin:22px auto}.card{padding:24px;border-radius:20px;box-shadow:0 16px 45px rgba(16,47,69,.1)}
.workspace-head{align-items:center;margin-bottom:17px}.workspace-head h2{font-size:23px}.workspace-head p{font-size:13px}
.drop{padding:20px 22px;border-radius:15px;border-color:#74acd2;background:linear-gradient(135deg,#eef7ff,#f8fcff);display:grid;grid-template-columns:48px minmax(0,1fr) auto;align-items:center;gap:18px;text-align:left;box-shadow:inset 0 0 0 1px rgba(37,99,235,.04)}.drop:hover{border-color:#2563eb;background:linear-gradient(135deg,#e8f3ff,#f5faff)}.upload-icon{display:none}.drop strong{font-size:18px;margin:0;color:var(--navy)}
.drop-symbol{width:48px;height:48px;display:grid;place-items:center;border-radius:14px;background:linear-gradient(145deg,#dceff4,#e8f7f1);color:var(--blue);box-shadow:inset 0 0 0 1px rgba(19,122,139,.15)}.drop-symbol svg{width:25px;height:25px}.drop-copy{min-width:0}
.format-row{display:flex;justify-content:flex-start;gap:6px;margin:8px 0 0;flex-wrap:wrap}
.format-badge{padding:4px 8px;border:1px solid #c9dce0;border-radius:7px;background:#fff;color:#41616c;font-size:11px;font-weight:800;letter-spacing:.05em}
.drop-actions{justify-content:flex-end;margin:0}.file-pick{background:var(--navy);color:#fff;box-shadow:0 6px 14px rgba(16,47,69,.16)}
.file-pick:hover:not(:disabled){background:#19455c;box-shadow:0 8px 18px rgba(16,47,69,.2)}
.reference{padding:14px 16px;border-color:#e7c66b;border-left:4px solid #d99a12;border-radius:13px;background:linear-gradient(135deg,#fff8df,#fffdf3);display:flex;align-items:center;justify-content:space-between;gap:18px;flex-wrap:wrap}.reference-badge{display:none}.reference-title{margin:0;min-width:0}.reference input[type=file]{display:none}
.reference-row{margin:0 0 0 auto}.selected-reference{flex-basis:100%;margin-top:0}.custom-file{background:#fff;color:var(--navy);border:1px solid #bcd0d5;box-shadow:0 2px 7px rgba(16,47,69,.05)}
.custom-file:hover:not(:disabled){background:#f2f8f9;border-color:#91b3ba}
.actions{padding-top:2px}.actions .primary{min-width:245px}.actions .secondary{border:1px solid #d7e3e6}
.toast-region{position:fixed;z-index:1000;top:16px;left:50%;width:min(440px,calc(100vw - 28px));transform:translateX(-50%);display:grid;gap:9px;pointer-events:none}
.toast{display:flex;align-items:flex-start;gap:11px;padding:12px 14px;border:1px solid #d6e3e5;border-radius:12px;background:rgba(255,255,255,.97);color:var(--ink);box-shadow:0 16px 42px rgba(16,47,69,.18);opacity:0;transform:translateY(-16px) scale(.98);transition:opacity .22s,transform .22s;backdrop-filter:blur(12px)}
.toast.show{opacity:1;transform:none}.toast-mark{display:grid;place-items:center;width:23px;height:23px;flex:0 0 auto;border-radius:50%;font-size:12px;font-weight:900}.toast-copy{min-width:0}.toast-title{font-size:13px;font-weight:850}.toast-message{font-size:12px;color:var(--muted);margin-top:1px;overflow-wrap:anywhere}
.toast.success{border-color:#b8e0d1}.toast.success .toast-mark{background:#dcf5ea;color:var(--green-dark)}.toast.error{border-color:#efc7c2}.toast.error .toast-mark{background:#ffebe8;color:var(--danger)}.toast.info .toast-mark{background:#e7f2f4;color:var(--blue)}
:focus-visible{outline:3px solid rgba(19,122,139,.34);outline-offset:3px}
@keyframes liftIn{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:none}}
@keyframes rowIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
@keyframes statusIn{from{opacity:0;transform:translateY(-4px)}to{opacity:1;transform:none}}
@keyframes spin{to{transform:rotate(360deg)}}
@media(max-width:760px){.top{padding-top:18px}.top-inner{align-items:flex-start;flex-direction:column}.nav{justify-content:flex-start;width:100%}.nav a{flex:1;text-align:center}.workspace-head{flex-direction:column;align-items:flex-start}.trust-chip{align-self:flex-start}.grid{grid-template-columns:repeat(2,1fr)}.drop{grid-template-columns:1fr;padding:20px 15px;text-align:center;gap:12px}.drop-symbol{margin:auto}.format-row,.drop-actions{justify-content:center}.reference{align-items:flex-start}.reference-row{width:100%;margin-left:0}.card{border-radius:17px}.actions .muted{flex-basis:100%}}
@media(max-width:460px){.brand-mark{width:42px;height:42px}.sub{font-size:13px}.nav{gap:7px}.nav a{padding:9px 8px;font-size:11px}.grid{grid-template-columns:1fr}.file{align-items:flex-start;flex-direction:column}.file-tools{width:100%;justify-content:space-between}.actions button{width:100%}.reference-row input{width:100%}}
@media(prefers-reduced-motion:reduce){*,*::before,*::after{animation-duration:.001ms!important;animation-iteration-count:1!important;scroll-behavior:auto!important;transition-duration:.001ms!important}}
</style>
</head>
<body>
<div id="toastRegion" class="toast-region" role="status" aria-live="polite" aria-atomic="true"></div>
<div class="top"><div class="top-inner">
  <div class="brand">
    <div class="brand-copy"><div class="eyebrow">Nazorat • Tahlil • Hisobot</div><h1>SVOD TIZIMI</h1>
    <div class="sub">CSV, ZIP yoki papkani tekshirish va tayyor Excel svod yaratish</div></div>
  </div>
  <nav class="nav" aria-label="Asosiy bo‘limlar">
    <a class="active" aria-current="page" href="/">SVOD HISOBOTI</a>
    <a href="/sotuvlar">SOTILGAN GAZLAR</a>
    <a href="/egaz">E-GAZ HISOBOTI</a>
    <a href="/gnp-taqqoslash">GNP + MFY SVOD</a>
  </nav>
</div></div>
<main class="wrap">
  <div class="card">
    <div class="workspace-head">
      <div><h2>Yangi hisobot tayyorlash</h2></div>
    </div>
    <div id="drop" class="drop" role="button" tabindex="0" aria-label="CSV, XLSX, ZIP fayllari yoki papkani tanlash">
      <span class="drop-symbol" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4"/><path d="m7 9 5-5 5 5"/><path d="M5 20h14"/></svg></span>
      <div class="drop-copy"><strong>Fayl yoki papkani shu yerga tashlang</strong>
        <div class="format-row" aria-label="Qabul qilinadigan formatlar"><span class="format-badge">CSV</span><span class="format-badge">XLSX</span><span class="format-badge">ZIP</span><span class="format-badge">PAPKA</span></div>
      </div>
      <input id="fileInput" type="file" accept=".csv,.xlsx,.zip,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/zip" multiple>
      <input id="folderInput" type="file" webkitdirectory directory multiple hidden>
      <div class="drop-actions"><button id="filePick" class="file-pick small" type="button">Fayllarni tanlash</button><button id="folderPick" class="secondary small" type="button">Papkani tanlash</button></div>
    </div>

    <div id="referenceDrop" class="reference">
      <div class="reference-title"><span class="reference-badge" aria-hidden="true">↔</span><label for="referenceInput">Tayyor davomat jadvali bilan solishtirish (ixtiyoriy)</label></div>
      <div class="reference-row">
        <input id="referenceInput" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden>
        <button id="referencePick" class="custom-file small" type="button">Etalon XLSX tanlash</button>
        <button id="referenceClear" class="remove small" type="button" hidden>Etalonni olib tashlash</button>
      </div>
      <div id="referenceSelected" class="selected-reference" hidden></div>
    </div>

    <div id="fileSummary" class="file-summary"><span><b id="fileCount">0 ta fayl</b> tanlandi</span><span id="fileSize">0 KB</span></div>
    <div id="filelist" class="filelist"></div>

    <div class="actions">
      <button id="go" class="primary" disabled>Tekshirish va Excel yaratish</button>
      <button id="clear" class="secondary">Tozalash</button>
    </div>

    <div id="status" class="status" role="status" aria-live="polite"></div>
    <div id="kpis" class="grid" style="display:none"></div>

  </div>
  <div class="footer"><strong>SVOD TIZIMI v1.3</strong> — {{ privacy_notice }}</div>
</main>

<script>
let files = [];
let referenceFile = null;
const hostedMode = {{ hosted | tojson }};
const maxUploadBytes = {{ upload_limit_bytes }};
const uploadLimitMb = {{ upload_limit_mb }};
const largeFileHint = hostedMode ? 'Katta fayllarni lokal dasturda ishlating.' : 'Fayllarni kichraytiring.';
const drop = document.getElementById('drop');
const inp = document.getElementById('fileInput');
const folderInput = document.getElementById('folderInput');
const folderPick = document.getElementById('folderPick');
const referenceDrop = document.getElementById('referenceDrop');
const referenceInput = document.getElementById('referenceInput');
const referenceClear = document.getElementById('referenceClear');
const referenceSelected = document.getElementById('referenceSelected');
const list = document.getElementById('filelist');
const go = document.getElementById('go');
const clearBtn = document.getElementById('clear');
const status = document.getElementById('status');
const kpis = document.getElementById('kpis');
const fileSummary = document.getElementById('fileSummary');
const fileCount = document.getElementById('fileCount');
const fileSize = document.getElementById('fileSize');
const filePick = document.getElementById('filePick');
const referencePick = document.getElementById('referencePick');
const toastRegion = document.getElementById('toastRegion');
let audioContext = null;

function playTone(kind='add'){
  try{
    const AudioEngine=window.AudioContext||window.webkitAudioContext;
    if(!AudioEngine)return;
    if(!audioContext)audioContext=new AudioEngine();
    if(audioContext.state==='suspended')audioContext.resume();
    const sequences={add:[[620,0,.10]],done:[[520,0,.11],[760,.12,.16]],error:[[220,0,.16]]};
    const now=audioContext.currentTime;
    (sequences[kind]||sequences.add).forEach(([frequency,delay,duration])=>{
      const oscillator=audioContext.createOscillator();
      const gain=audioContext.createGain();
      const start=now+delay;
      oscillator.type='sine';oscillator.frequency.setValueAtTime(frequency,start);
      gain.gain.setValueAtTime(.0001,start);
      gain.gain.exponentialRampToValueAtTime(.075,start+.012);
      gain.gain.exponentialRampToValueAtTime(.0001,start+duration);
      oscillator.connect(gain);gain.connect(audioContext.destination);
      oscillator.start(start);oscillator.stop(start+duration+.02);
    });
  }catch(_){}
}
function showToast(message,type='info',sound=''){
  const toast=document.createElement('div');toast.className='toast '+type;
  const mark=document.createElement('span');mark.className='toast-mark';mark.textContent=type==='success'?'✓':type==='error'?'!':'i';
  const copy=document.createElement('div');copy.className='toast-copy';
  const title=document.createElement('div');title.className='toast-title';title.textContent=type==='success'?'Muvaffaqiyatli':type==='error'?'Xatolik':'Ma’lumot';
  const body=document.createElement('div');body.className='toast-message';body.textContent=message;
  copy.append(title,body);toast.append(mark,copy);toastRegion.appendChild(toast);
  requestAnimationFrame(()=>toast.classList.add('show'));
  setTimeout(()=>{toast.classList.remove('show');setTimeout(()=>toast.remove(),230);},type==='error'?5200:3400);
  if(sound)playTone(sound);
}

// Fayl maydondan tashqariga tashlansa, brauzer uni ochib sahifani yo'qotmasin.
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('drop', e => e.preventDefault());

drop.onclick = () => inp.click();
filePick.onclick = e => {e.stopPropagation();inp.click();};
drop.onkeydown = e => {
  if(e.key==='Enter'||e.key===' '){e.preventDefault();inp.click();}
};
drop.ondragover = e => { e.preventDefault(); drop.classList.add('drag'); };
drop.ondragleave = () => drop.classList.remove('drag');
drop.ondrop = async e => {
  e.preventDefault(); drop.classList.remove('drag');
  try{
    const droppedFiles=await collectDroppedFiles(e.dataTransfer);
    const accepted=droppedFiles.filter(f=>/\.(csv|xlsx|zip)$/i.test(f.name));
    if(!accepted.length){
      status.className='status err';
      status.textContent='Tashlangan joyda CSV, XLSX yoki ZIP fayl topilmadi.';
      showToast(status.textContent,'error','error');
      return;
    }
    addFiles(accepted);
  }catch(e){
    status.className='status err';
    status.textContent='Papkani o‘qib bo‘lmadi. “Papkani tanlash” tugmasidan foydalaning.';
    showToast(status.textContent,'error','error');
  }
};
inp.onchange = () => { addFiles([...inp.files]); inp.value = ''; };
folderPick.onclick=e=>{e.stopPropagation();folderInput.click();};
folderInput.onchange=()=>{addFiles([...folderInput.files]);folderInput.value='';};

function entryFile(entry){
  return new Promise((resolve,reject)=>entry.file(resolve,reject));
}
function directoryBatch(reader){
  return new Promise((resolve,reject)=>reader.readEntries(resolve,reject));
}
async function collectEntryFiles(entry,result){
  if(entry.isFile){
    const file=await entryFile(entry);
    file._relativePath=entry.fullPath||file.name;
    result.push(file);
    return;
  }
  if(!entry.isDirectory)return;
  const reader=entry.createReader();
  while(true){
    const entries=await directoryBatch(reader);
    if(!entries.length)break;
    for(const child of entries)await collectEntryFiles(child,result);
  }
}
async function collectDroppedFiles(dataTransfer){
  const entries=[...dataTransfer.items]
    .map(item=>item.webkitGetAsEntry?item.webkitGetAsEntry():null).filter(Boolean);
  if(!entries.length)return [...dataTransfer.files];
  const result=[];
  for(const entry of entries)await collectEntryFiles(entry,result);
  return result;
}

function addFiles(newFiles){
  let added=0;
  for(const f of newFiles){
    const lower=f.name.toLowerCase();
    if(!lower.endsWith('.csv') && !lower.endsWith('.xlsx') && !lower.endsWith('.zip')) continue;
    const displayName=f.webkitRelativePath||f._relativePath||f.name;
    const key = displayName + ':' + f.size + ':' + f.lastModified;
    f._displayName=displayName;
    if(!files.some(x => x._key === key)){ f._key = key; files.push(f); added++; }
  }
  render();
  if(added)showToast(added+' ta fayl qo‘shildi. Tekshirishga tayyor.','success','add');
  return added;
}
function render(){
  list.innerHTML = '';
  files.forEach((f,i)=>{
    const d=document.createElement('div'); d.className='file';
    d.innerHTML=`<span class="file-name">${i+1}. ${escapeHtml(f._displayName||f.name)}</span>`+
      `<span class="file-tools"><span>${(f.size/1024).toFixed(1)} KB</span>`+
      `<button class="remove small" type="button" data-remove="${i}">Olib tashlash</button></span>`;
    list.appendChild(d);
  });
  list.querySelectorAll('[data-remove]').forEach(button=>{
    button.onclick=()=>{files.splice(Number(button.dataset.remove),1);render();};
  });
  go.disabled = files.length===0;
  const totalBytes=files.reduce((sum,file)=>sum+file.size,0);
  fileSummary.classList.toggle('show',files.length>0);
  fileCount.textContent=files.length+' ta fayl';
  fileSize.textContent=formatBytes(totalBytes);
}
function escapeHtml(s){return s.replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]));}
function formatBytes(bytes){
  if(bytes<1024)return bytes+' B';
  if(bytes<1024*1024)return (bytes/1024).toFixed(1)+' KB';
  return (bytes/(1024*1024)).toFixed(2)+' MB';
}
function setReference(file){
  if(file && (!file.name.toLowerCase().endsWith('.xlsx') || file.size>10*1024*1024)){
    status.className='status err';
    status.textContent='Solishtirish uchun 10 MB gacha .xlsx fayl tanlang yoki tashlang.';
    showToast(status.textContent,'error','error');
    return false;
  }
  referenceFile=file||null;
  referenceClear.hidden=!referenceFile;
  referenceSelected.hidden=!referenceFile;
  referenceSelected.textContent=referenceFile?'Tanlandi: '+referenceFile.name:'';
  if(referenceFile)showToast('Etalon Excel tanlandi: '+referenceFile.name,'success','add');
  return true;
}
referencePick.onclick=()=>referenceInput.click();
referenceInput.onchange=()=>{setReference(referenceInput.files[0]||null);};
referenceDrop.ondragover=e=>{e.preventDefault();e.stopPropagation();referenceDrop.classList.add('drag');};
referenceDrop.ondragleave=e=>{if(!referenceDrop.contains(e.relatedTarget))referenceDrop.classList.remove('drag');};
referenceDrop.ondrop=e=>{
  e.preventDefault();e.stopPropagation();referenceDrop.classList.remove('drag');
  const file=[...e.dataTransfer.files].find(f=>f.name.toLowerCase().endsWith('.xlsx'));
  if(!file){
    status.className='status err';
    status.textContent='Bu maydonga solishtirish uchun .xlsx fayl tashlang.';
    showToast(status.textContent,'error','error');
    return;
  }
  referenceInput.value='';
  setReference(file);
};
referenceClear.onclick=()=>{referenceInput.value='';setReference(null);};
clearBtn.onclick=()=>{files=[];inp.value='';folderInput.value='';referenceInput.value='';setReference(null);render();status.className='status';status.textContent='';kpis.style.display='none';};
go.onclick=async()=>{
  if(!files.length)return;
  if(referenceFile && (!referenceFile.name.toLowerCase().endsWith('.xlsx') || referenceFile.size>10*1024*1024)){
    status.className='status err';
    status.textContent='Solishtirish uchun 10 MB gacha .xlsx fayl tanlang.';
    showToast(status.textContent,'error','error');
    return;
  }
  if(files.reduce((sum,f)=>sum+f.size+1024,referenceFile?referenceFile.size+1024:0)>maxUploadBytes){
    status.className='status err';
    status.textContent='Fayllar jami '+uploadLimitMb+' MB limitdan oshdi. '+largeFileHint;
    showToast(status.textContent,'error','error');
    return;
  }
  go.disabled=true;
  go.setAttribute('aria-busy','true');
  go.textContent='Tayyorlanmoqda…';
  status.className='status work';
  const t0=Date.now();
  const tick=()=>{status.textContent='Tekshirilmoqda va Excel tayyorlanmoqda... '
    +Math.round((Date.now()-t0)/1000)+' s\n(katta fayllarda 1-2 daqiqa davom etishi mumkin)';};
  tick();
  const timer=setInterval(tick,1000);
  kpis.style.display='none';
  const fd=new FormData();
  files.forEach(f=>fd.append('files',f,f.name));
  if(referenceFile)fd.append('reference',referenceFile,referenceFile.name);
  try{
    let res;
    try{
      res=await fetch('/generate',{method:'POST',body:fd});
    }catch(e){
      throw new Error("Server bilan aloqa yo'q. Sahifani yangilang va qayta urinib ko'ring.");
    }
    if(!res.ok){
      if(res.status===413) throw new Error('Fayl yoki tayyor Excel server limitidan oshdi. '+largeFileHint);
      const txt=await res.text();
      let msg=txt||('Xatolik: HTTP '+res.status);
      try{const j=JSON.parse(txt);msg=j.error||msg;}catch(e){}
      throw new Error(msg);
    }
    const blob=await res.blob();
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);
    a.download=res.headers.get('X-Output-Name')||'SVOD.xlsx';
    document.body.appendChild(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(a.href),1000);

    const reportType=res.headers.get('X-Report-Type')||'duration';
    const total=res.headers.get('X-Total-Rows')||'-';
    const ray=res.headers.get('X-Raygaz-Count')||'-';
    const mah=res.headers.get('X-Mahalla-Count')||'-';
    const dup=res.headers.get('X-Duplicate-Count')||'-';
    const mismatch=res.headers.get('X-Mismatch-Count');
    status.className='status ok';
    if(reportType==='gas-sales'){
      const accepted=res.headers.get('X-Accepted-Total')||'-';
      const sold=res.headers.get('X-Sold-Total')||'-';
      const returned=res.headers.get('X-Returned-Total')||'-';
      const percent=res.headers.get('X-Sales-Percent')||'-';
      status.textContent='Tayyor. Qabul va sotuv bo‘yicha Excel svod yaratildi va yuklandi.';
      kpis.innerHTML=`
        <div class="kpi"><span>Qabul qilindi</span><b>${accepted}</b></div>
        <div class="kpi"><span>Sotildi</span><b>${sold}</b></div>
        <div class="kpi"><span>Sotilmagan gaz</span><b>${returned}</b></div>
        <div class="kpi"><span>Sotuv foizi</span><b>${percent}%</b></div>`;
    }else{
      status.textContent='Tayyor. Excel fayl yaratildi va yuklandi.'+
        (mismatch!==null?' Solishtirish varag‘ida '+mismatch+' ta farq aniqlandi.':'');
      kpis.innerHTML=`
        <div class="kpi"><span>Jami yozuv</span><b>${total}</b></div>
        <div class="kpi"><span>Raygaz</span><b>${ray}</b></div>
        <div class="kpi"><span>Mahalla</span><b>${mah}</b></div>
        <div class="kpi"><span>Dublikat</span><b>${dup}</b></div>`;
    }
    kpis.style.display='grid';
    showToast(reportType==='gas-sales'?'Excel hisobot tayyorlandi va yuklandi.':'SVOD Excel tayyorlandi va yuklandi.','success','done');
  }catch(e){
    status.className='status err';
    status.textContent=e.message;
    showToast(e.message||'Hisobot yaratishda xatolik yuz berdi.','error','error');
  }finally{
    clearInterval(timer);
    go.removeAttribute('aria-busy');
    go.textContent='Tekshirish va Excel yaratish';
    go.disabled=files.length===0;
  }
};
</script>
</body>
</html>
"""


SALES_HTML = r"""
<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#102f45">
<title>Sotilgan gazlar — SVOD TIZIMI</title>
<style>
:root{
  --navy:#102f45;--navy-2:#174d66;--blue:#137a8b;--green:#15805d;--green-dark:#0d6549;
  --mint:#e9f8f2;--light:#f3f7f8;--surface:#fff;--border:#d8e4e7;--danger:#b42318;--warn:#9a5b09;--ink:#142b36;--muted:#637b86;
  --shadow:0 24px 70px rgba(16,47,69,.12);--shadow-soft:0 10px 30px rgba(16,47,69,.08);
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}body{margin:0;min-height:100vh;background:radial-gradient(circle at 7% 0%,rgba(19,122,139,.12),transparent 30rem),radial-gradient(circle at 93% 20%,rgba(21,128,93,.1),transparent 28rem),var(--light);font-family:Inter,"Segoe UI",Arial,sans-serif;color:var(--ink);line-height:1.5}
button,input{font:inherit}.top{position:relative;overflow:hidden;background:linear-gradient(125deg,#0c293d 0%,var(--navy-2) 58%,#126e70 100%);color:white;padding:31px 18px 33px;border-bottom:1px solid rgba(255,255,255,.12)}
.top::after{content:"";position:absolute;width:360px;height:360px;border:70px solid rgba(255,255,255,.045);border-radius:50%;right:-100px;top:-210px;pointer-events:none}
.top-inner{position:relative;z-index:1;max-width:1180px;margin:auto;display:flex;align-items:center;justify-content:space-between;gap:28px}
.brand{min-width:0;display:flex;align-items:center;gap:15px}.brand-mark{width:48px;height:48px;display:grid;place-items:center;flex:0 0 auto;border-radius:15px;background:rgba(255,255,255,.13);border:1px solid rgba(255,255,255,.22)}.brand-mark svg{width:25px;height:25px}.brand-copy{min-width:0}.eyebrow{font-size:11px;font-weight:800;letter-spacing:.16em;text-transform:uppercase;color:#84e0c4;margin-bottom:3px}h1{margin:0;font-size:clamp(27px,4vw,34px);line-height:1.1;letter-spacing:-.025em}.sub{opacity:.83;font-size:14px;margin-top:6px}
.nav{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end}
.nav a{color:rgba(255,255,255,.88);text-decoration:none;font-weight:750;font-size:13px;letter-spacing:.025em;padding:10px 14px;border:1px solid rgba(255,255,255,.22);border-radius:11px;white-space:nowrap;transition:transform .2s,background .2s,border-color .2s,color .2s}.nav a:hover{transform:translateY(-2px);border-color:rgba(255,255,255,.48)}.nav a.active{background:white;color:var(--navy);border-color:white;box-shadow:0 7px 20px rgba(0,0,0,.14)}
.wrap{max-width:1500px;margin:30px auto;padding:0 18px}.card{background:rgba(255,255,255,.95);border:1px solid rgba(184,206,211,.8);border-radius:24px;padding:clamp(18px,3vw,30px);box-shadow:var(--shadow);backdrop-filter:blur(12px);animation:liftIn .55s cubic-bezier(.2,.8,.2,1) both}
.intro{display:flex;align-items:flex-start;justify-content:space-between;gap:24px;margin-bottom:21px}.intro h2{margin:0 0 6px;font-size:clamp(21px,3vw,27px);letter-spacing:-.025em}.intro p{margin:0;max-width:920px}.mode-chip{display:flex;align-items:center;gap:8px;flex:0 1 auto;max-width:430px;background:var(--mint);color:var(--green-dark);border:1px solid #bde6d8;border-radius:999px;padding:8px 12px;font-size:12px;font-weight:800;line-height:1.35}.mode-chip span{width:8px;height:8px;flex:0 0 auto;border-radius:50%;background:var(--green);box-shadow:0 0 0 4px rgba(21,128,93,.12)}.muted{color:var(--muted);font-size:14px;line-height:1.5}
.snapshots{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.snapshot{border:1px solid var(--border);border-radius:18px;padding:18px;background:linear-gradient(160deg,#fcfefe,#f7fafb);transition:box-shadow .2s,transform .2s}.snapshot:focus-within{box-shadow:var(--shadow-soft)}
.snapshot h2{display:flex;align-items:center;font-size:19px;margin:0 0 5px;color:var(--navy)}.step{display:inline-grid;place-items:center;width:30px;height:30px;border-radius:10px;background:linear-gradient(145deg,var(--blue),#1498a5);color:#fff;margin-right:9px;font-size:14px;box-shadow:0 5px 13px rgba(19,122,139,.22)}
.drop{position:relative;overflow:hidden;margin-top:14px;border:2px dashed #8eb4bb;border-radius:15px;padding:29px 16px;text-align:center;cursor:pointer;background:linear-gradient(180deg,#fbfefe,#f4fafb);transition:transform .22s,border-color .22s,background .22s,box-shadow .22s;outline:none}
.drop-icon{width:44px;height:44px;display:grid;place-items:center;margin:0 auto 10px;border-radius:14px;color:var(--green);background:#e4f5ef}.drop-icon svg{width:23px;height:23px}.drop:hover{transform:translateY(-2px);border-color:var(--blue);box-shadow:var(--shadow-soft)}.drop.drag{transform:scale(1.008);border-color:var(--green);background:var(--mint);box-shadow:0 0 0 5px rgba(21,128,93,.09)}.drop strong{display:block;margin-bottom:5px;color:var(--navy)}
.drop input{display:none}.drop-actions{display:flex;justify-content:center;gap:8px;margin-top:12px;flex-wrap:wrap}
button{border:0;border-radius:11px;padding:12px 18px;font-size:15px;font-weight:800;cursor:pointer;transition:transform .18s,box-shadow .18s,background .18s}button:hover:not(:disabled){transform:translateY(-2px)}button:active:not(:disabled){transform:none}.primary{background:linear-gradient(135deg,var(--green),#1a9470);color:white;box-shadow:0 8px 18px rgba(21,128,93,.22)}.primary:hover:not(:disabled){box-shadow:0 12px 24px rgba(21,128,93,.28)}.primary:disabled{opacity:.45;cursor:not-allowed;box-shadow:none}.secondary{background:#e8f0f2;color:#244d5b}.secondary:hover:not(:disabled){background:#dce9ec}.small{padding:7px 11px;font-size:13px}.remove{background:#fff0ee;color:var(--danger)}
.filelist{margin-top:12px;max-height:180px;overflow:auto}
.file{display:flex;justify-content:space-between;align-items:center;gap:10px;padding:10px 3px;border-bottom:1px solid #eaf0f2;font-size:13px;animation:rowIn .25s ease both}
.file-name{min-width:0;overflow-wrap:anywhere}.file-tools{white-space:nowrap;display:flex;align-items:center;gap:8px}
.actions{display:flex;gap:12px;align-items:center;margin-top:20px;flex-wrap:wrap}
.status{position:relative;margin-top:18px;padding:14px 16px 14px 44px;border-radius:12px;display:none;white-space:pre-wrap;animation:statusIn .25s ease both}.status::before{position:absolute;left:16px;top:14px;font-weight:900}.status.ok{display:block;background:#eaf8f1;color:#126342;border:1px solid #b9e2d0}.status.ok::before{content:"✓"}.status.err{display:block;background:#fff1ef;color:var(--danger);border:1px solid #f0c9c4}.status.err::before{content:"!"}.status.work{display:block;background:#fff8e9;color:var(--warn);border:1px solid #efd89f}.status.work::before{content:"";width:15px;height:15px;border:2px solid currentColor;border-right-color:transparent;border-radius:50%;top:16px;animation:spin .75s linear infinite}
.results{display:none;margin-top:24px}.results.show{display:block;animation:liftIn .45s cubic-bezier(.2,.8,.2,1) both}.dates{display:inline-flex;align-items:center;font-weight:800;color:var(--blue);margin-bottom:14px;padding:8px 12px;background:#eaf4f6;border-radius:10px;font-size:13px}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.kpi{background:linear-gradient(145deg,#f9fcfc,#f1f7f8);border:1px solid var(--border);border-radius:14px;padding:15px;animation:rowIn .35s ease both}.kpi:nth-child(2){animation-delay:.05s}.kpi:nth-child(3){animation-delay:.1s}.kpi:nth-child(4){animation-delay:.15s}.kpi span{font-size:12px;color:var(--muted);font-weight:700}.kpi b{font-size:25px;display:block;color:var(--navy);margin-top:3px;font-variant-numeric:tabular-nums}
.section-head{display:flex;justify-content:space-between;align-items:end;gap:12px;margin:25px 0 10px}.section-head h2{font-size:19px;margin:0}.tools{display:flex;gap:9px;align-items:center;flex-wrap:wrap}
.search{border:1px solid #b7cbd0;border-radius:10px;padding:10px 12px;min-width:280px;color:inherit;background:#fbfdfd;transition:border-color .2s,box-shadow .2s}.search:focus{border-color:var(--blue);box-shadow:0 0 0 4px rgba(19,122,139,.1);outline:none}
.table-wrap{overflow:auto;border:1px solid var(--border);border-radius:13px;max-height:520px;box-shadow:0 6px 22px rgba(16,47,69,.05)}table{border-collapse:separate;border-spacing:0;width:100%;font-size:14px;background:white}th,td{padding:11px 12px;text-align:left;border-bottom:1px solid #e8eef0;white-space:nowrap}th{position:sticky;top:0;z-index:1;background:var(--navy);color:white;font-size:12px;letter-spacing:.015em}tbody tr{transition:background .15s}tbody tr:hover{background:#f1f8f8}td.num,th.num{text-align:right;font-variant-numeric:tabular-nums}
.gas-combined-wrap{overflow:visible;max-height:none}
.gas-combined-table{table-layout:fixed}
.gas-combined-table th,.gas-combined-table td{padding:10px 8px;white-space:normal;overflow-wrap:anywhere;line-height:1.3}
.gas-combined-table th.num,.gas-combined-table td.num{white-space:nowrap;overflow-wrap:normal}
.gas-combined-table th:nth-child(1),.gas-combined-table td:nth-child(1){width:18%}
.gas-combined-table th:nth-child(2),.gas-combined-table td:nth-child(2){width:30%}
.gas-combined-table th:nth-child(3),.gas-combined-table td:nth-child(3),
.gas-combined-table th:nth-child(4),.gas-combined-table td:nth-child(4),
.gas-combined-table th:nth-child(5),.gas-combined-table td:nth-child(5){width:11%}
.gas-combined-table th:nth-child(6),.gas-combined-table td:nth-child(6){width:6%}
.gas-combined-table th:nth-child(7),.gas-combined-table td:nth-child(7){width:13%}
.gas-total td{background:#fff5b8;font-weight:800}.gas-grand-total td{background:#f7d0b6;font-weight:800}.empty{text-align:center;color:var(--muted);padding:25px}.control-note{margin-top:11px;padding:13px 15px;border:1px solid #f0dfb8;border-left:4px solid var(--warn);background:#fff9ec;border-radius:9px;font-size:13px;line-height:1.5}.footer{color:#718892;font-size:12px;text-align:center;margin:19px}.footer strong{color:#46636e}
/* Minimal, compact workspace */
.top{padding:19px 18px 20px;background:linear-gradient(115deg,#102f45,#15556a 62%,#146f70)}.top::after{display:none}
.top-inner{gap:20px}.brand{gap:0}.brand-mark{display:none}.eyebrow{font-size:10px;letter-spacing:.2em;margin-bottom:2px}h1{font-size:29px}.sub{font-size:13px;margin-top:4px}
.nav{gap:12px;padding:0;background:transparent;border:0;border-radius:0}.nav a{color:#fff;border:1px solid rgba(255,255,255,.34);border-radius:10px;padding:10px 14px;font-size:12px;box-shadow:0 5px 14px rgba(4,28,40,.2)}.nav a:nth-child(1){background:#2563eb}.nav a:nth-child(2){background:#d97706}.nav a:nth-child(3){background:#15803d}.nav a:hover{transform:translateY(-1px);filter:brightness(1.1)}.nav a.active{color:#fff;border-color:#fff;box-shadow:0 0 0 2px rgba(255,255,255,.25),0 6px 16px rgba(4,28,40,.28)}
.wrap{margin:22px auto}.card{padding:24px;border-radius:20px;box-shadow:0 16px 45px rgba(16,47,69,.1)}.intro{align-items:center;margin-bottom:17px}.intro h2{font-size:23px}
.snapshot{padding:15px;border-radius:15px}.snapshot:first-child{border-color:#b8d8ee;border-top:4px solid #2f80c8;background:linear-gradient(145deg,#f2f9ff,#fbfdff)}.snapshot:last-child{border-color:#b7dfcc;border-top:4px solid #23936a;background:linear-gradient(145deg,#f0faf5,#fbfefc)}.snapshot:first-child .step{background:linear-gradient(145deg,#267bc0,#3d9bd5)}.snapshot:last-child .step{background:linear-gradient(145deg,#15805d,#2aa67a)}.drop{padding:18px;display:grid;grid-template-columns:44px minmax(0,1fr) auto;align-items:center;gap:14px;text-align:left}.drop-icon{display:none}#oldDrop{border-color:#79add1;background:linear-gradient(135deg,#eaf5fd,#f8fcff)}#newDrop{border-color:#75b997;background:linear-gradient(135deg,#e9f8f1,#f8fdfb)}#oldDrop:hover,#oldDrop.drag{border-color:#2f80c8;background:#e4f2fc;box-shadow:0 0 0 5px rgba(47,128,200,.1)}#newDrop:hover,#newDrop.drag{border-color:#15805d;background:#e3f6ed;box-shadow:0 0 0 5px rgba(21,128,93,.1)}.drop-symbol{width:44px;height:44px;display:grid;place-items:center;border-radius:13px;background:#dceef9;color:#2f80c8;box-shadow:inset 0 0 0 1px rgba(47,128,200,.16)}#newDrop .drop-symbol{background:#dff3e9;color:var(--green);box-shadow:inset 0 0 0 1px rgba(21,128,93,.14)}.drop-symbol svg{width:23px;height:23px}.drop-copy{min-width:0}.drop strong{margin:0}.format-row{display:flex;justify-content:flex-start;gap:6px;margin:8px 0 0;flex-wrap:wrap}.format-badge{padding:4px 8px;border:1px solid #c9dce0;border-radius:7px;background:#fff;color:#41616c;font-size:11px;font-weight:800;letter-spacing:.05em}.drop-actions{justify-content:flex-end;margin:0}.file-pick{background:var(--navy);color:#fff;box-shadow:0 6px 14px rgba(16,47,69,.16)}.file-pick:hover:not(:disabled){background:#19455c}
.actions .primary{min-width:210px}.actions .secondary{border:1px solid #d7e3e6}
.toast-region{position:fixed;z-index:1000;top:16px;left:50%;width:min(440px,calc(100vw - 28px));transform:translateX(-50%);display:grid;gap:9px;pointer-events:none}.toast{display:flex;align-items:flex-start;gap:11px;padding:12px 14px;border:1px solid #d6e3e5;border-radius:12px;background:rgba(255,255,255,.97);color:var(--ink);box-shadow:0 16px 42px rgba(16,47,69,.18);opacity:0;transform:translateY(-16px) scale(.98);transition:opacity .22s,transform .22s;backdrop-filter:blur(12px)}.toast.show{opacity:1;transform:none}.toast-mark{display:grid;place-items:center;width:23px;height:23px;flex:0 0 auto;border-radius:50%;font-size:12px;font-weight:900}.toast-copy{min-width:0}.toast-title{font-size:13px;font-weight:850}.toast-message{font-size:12px;color:var(--muted);margin-top:1px;overflow-wrap:anywhere}.toast.success{border-color:#b8e0d1}.toast.success .toast-mark{background:#dcf5ea;color:var(--green-dark)}.toast.error{border-color:#efc7c2}.toast.error .toast-mark{background:#ffebe8;color:var(--danger)}.toast.info .toast-mark{background:#e7f2f4;color:var(--blue)}
:focus-visible{outline:3px solid rgba(19,122,139,.34);outline-offset:3px}
@keyframes liftIn{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:none}}@keyframes rowIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}@keyframes statusIn{from{opacity:0;transform:translateY(-4px)}to{opacity:1;transform:none}}@keyframes spin{to{transform:rotate(360deg)}}
@media(max-width:800px){.top{padding-top:18px}.top-inner{align-items:flex-start;flex-direction:column}.nav{justify-content:flex-start;width:100%}.nav a{flex:1;text-align:center}.intro{flex-direction:column;align-items:flex-start}.mode-chip{align-self:flex-start}.snapshots{grid-template-columns:1fr}.grid{grid-template-columns:repeat(2,1fr)}.section-head{align-items:stretch;flex-direction:column}.search{min-width:0;width:100%}.card{border-radius:17px}}
@media(max-width:600px){.drop{grid-template-columns:1fr;padding:18px 14px;text-align:center;gap:11px}.drop-symbol{margin:auto}.format-row,.drop-actions{justify-content:center}}
@media(max-width:480px){.brand-mark{width:42px;height:42px}.nav{gap:7px}.nav a{padding:9px 8px;font-size:11px}.grid{grid-template-columns:1fr}.file{align-items:flex-start;flex-direction:column}.file-tools{width:100%;justify-content:space-between}.actions button,.tools button{width:100%}.gas-combined-table{min-width:780px}}
@media(prefers-reduced-motion:reduce){*,*::before,*::after{animation-duration:.001ms!important;animation-iteration-count:1!important;scroll-behavior:auto!important;transition-duration:.001ms!important}}
</style>
</head>
<body>
<div id="toastRegion" class="toast-region" role="status" aria-live="polite" aria-atomic="true"></div>
<div class="top"><div class="top-inner">
  <div class="brand">
    <div class="brand-copy"><div class="eyebrow">Nazorat • Tahlil • Hisobot</div><h1>SVOD TIZIMI</h1><div class="sub">Ikki kun holatidan sotilgan gazlarni aniqlash</div></div>
  </div>
  <nav class="nav" aria-label="Asosiy bo‘limlar">
    <a href="/">SVOD HISOBOTI</a>
    <a class="active" aria-current="page" href="/sotuvlar">SOTILGAN GAZLAR</a>
    <a href="/egaz">E-GAZ HISOBOTI</a>
    <a href="/gnp-taqqoslash">GNP + MFY SVOD</a>
  </nav>
</div></div>

<main class="wrap">
  <div class="card">
    <div class="intro">
      <div><h2>Kechagi va yangi holatni solishtirish</h2></div>
    </div>

    <div class="snapshots">
      <section class="snapshot">
        <h2><span class="step">1</span>Kechagi holat</h2>
        <div id="oldDrop" class="drop" tabindex="0" role="button" aria-label="Kechagi holat fayllarini tanlash">
          <span class="drop-symbol" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4"/><path d="m7 9 5-5 5 5"/><path d="M5 20h14"/></svg></span>
          <div class="drop-copy"><strong>Kechagi fayllarni shu yerga tashlang</strong>
            <div class="format-row"><span class="format-badge">CSV</span><span class="format-badge">XLSX</span><span class="format-badge">ZIP</span><span class="format-badge">PAPKA</span></div>
          </div>
          <input id="oldInput" type="file" accept=".csv,.xlsx,.zip,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/zip" multiple>
          <input id="oldFolder" type="file" webkitdirectory directory multiple>
          <div class="drop-actions"><button class="file-pick small" type="button">Fayllarni tanlash</button><button class="secondary small folder-pick" type="button">Papkani tanlash</button></div>
        </div>
        <div id="oldList" class="filelist"></div>
      </section>

      <section class="snapshot">
        <h2><span class="step">2</span>Yangi holat</h2>
        <div id="newDrop" class="drop" tabindex="0" role="button" aria-label="Yangi holat fayllarini tanlash">
          <span class="drop-symbol" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4"/><path d="m7 9 5-5 5 5"/><path d="M5 20h14"/></svg></span>
          <div class="drop-copy"><strong>Yangi fayllarni shu yerga tashlang</strong>
            <div class="format-row"><span class="format-badge">CSV</span><span class="format-badge">XLSX</span><span class="format-badge">ZIP</span><span class="format-badge">PAPKA</span></div>
          </div>
          <input id="newInput" type="file" accept=".csv,.xlsx,.zip,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/zip" multiple>
          <input id="newFolder" type="file" webkitdirectory directory multiple>
          <div class="drop-actions"><button class="file-pick small" type="button">Fayllarni tanlash</button><button class="secondary small folder-pick" type="button">Papkani tanlash</button></div>
        </div>
        <div id="newList" class="filelist"></div>
      </section>
    </div>

    <div class="actions">
      <button id="compare" class="primary" disabled>Sotuvlarni hisoblash</button>
      <button id="clear" class="secondary" type="button">Tozalash</button>
    </div>
    <div id="status" class="status" role="status" aria-live="polite"></div>

    <section id="results" class="results">
      <div id="dates" class="dates"></div>
      <div id="kpis" class="grid"></div>

      <div id="summarySection">
        <div class="section-head"><h2 id="summaryTitle">Raygazlar bo‘yicha sotuv</h2></div>
        <div class="table-wrap"><table>
          <thead><tr id="summaryHead"><th>Raygaz</th><th class="num">Sotilgan abonent</th><th class="num">Jami ehtiyoj</th></tr></thead>
          <tbody id="summaryBody"></tbody>
        </table></div>
      </div>

      <div id="detailsSection">
        <div class="section-head">
          <h2 id="detailsTitle">Sotilgan abonentlar</h2>
          <div class="tools"><input id="search" class="search" type="search" aria-label="Natijalar ichidan qidirish" placeholder="Kod, abonent, Raygaz yoki mahalla..."><button id="downloadExcel" class="primary" type="button" hidden>Rasmdagidek Excel yuklash</button><button id="download" class="secondary" type="button">CSV yuklab olish</button></div>
        </div>
        <div id="detailsTableWrap" class="table-wrap"><table id="detailsTable">
          <thead><tr id="detailsHead"><th>№</th><th>Raygaz</th><th>Mahalla</th><th>Abonent kod</th><th>Abonent</th><th>Ehtiyoj</th><th>Eski realizatsiya</th><th>Yangi realizatsiya</th></tr></thead>
          <tbody id="detailsBody"></tbody>
        </table></div>
      </div>

      <div id="controlSection">
        <div class="section-head"><h2 id="controlTitle">Nazorat</h2></div>
        <div id="controlNote" class="control-note">Yangi holatda topilmagan kod sotuv deb taxmin qilinmaydi. U va kechagi holatda bo‘lmagan yangi kodlar quyida alohida ko‘rsatiladi.</div>
        <div class="table-wrap" style="margin-top:10px;max-height:300px"><table>
          <thead><tr id="controlHead"><th>Holat</th><th>Abonent kod</th><th>Raygaz</th><th>Mahalla</th><th>Abonent</th></tr></thead>
          <tbody id="controlBody"></tbody>
        </table></div>
      </div>
    </section>
  </div>
  <div class="footer"><strong>SVOD TIZIMI v1.3</strong> — {{ privacy_notice }}</div>
</main>

<script>
const hostedMode={{ hosted | tojson }};
const maxUploadBytes={{ upload_limit_bytes }};
const uploadLimitMb={{ upload_limit_mb }};
const states={old:[],new:[]};
let lastResult=null;
const statusEl=document.getElementById('status');
const compareBtn=document.getElementById('compare');
const resultsEl=document.getElementById('results');
const toastRegion=document.getElementById('toastRegion');
let audioContext=null;

function playTone(kind='add'){
  try{
    const AudioEngine=window.AudioContext||window.webkitAudioContext;
    if(!AudioEngine)return;
    if(!audioContext)audioContext=new AudioEngine();
    if(audioContext.state==='suspended')audioContext.resume();
    const sequences={add:[[620,0,.10]],done:[[520,0,.11],[760,.12,.16]],error:[[220,0,.16]]};
    const now=audioContext.currentTime;
    (sequences[kind]||sequences.add).forEach(([frequency,delay,duration])=>{
      const oscillator=audioContext.createOscillator();const gain=audioContext.createGain();const start=now+delay;
      oscillator.type='sine';oscillator.frequency.setValueAtTime(frequency,start);gain.gain.setValueAtTime(.0001,start);
      gain.gain.exponentialRampToValueAtTime(.075,start+.012);gain.gain.exponentialRampToValueAtTime(.0001,start+duration);
      oscillator.connect(gain);gain.connect(audioContext.destination);oscillator.start(start);oscillator.stop(start+duration+.02);
    });
  }catch(_){}
}
function showToast(message,type='info',sound=''){
  const toast=document.createElement('div');toast.className='toast '+type;
  const mark=document.createElement('span');mark.className='toast-mark';mark.textContent=type==='success'?'✓':type==='error'?'!':'i';
  const copy=document.createElement('div');copy.className='toast-copy';const title=document.createElement('div');title.className='toast-title';title.textContent=type==='success'?'Muvaffaqiyatli':type==='error'?'Xatolik':'Ma’lumot';
  const body=document.createElement('div');body.className='toast-message';body.textContent=message;copy.append(title,body);toast.append(mark,copy);toastRegion.appendChild(toast);
  requestAnimationFrame(()=>toast.classList.add('show'));setTimeout(()=>{toast.classList.remove('show');setTimeout(()=>toast.remove(),230);},type==='error'?5200:3400);
  if(sound)playTone(sound);
}

window.addEventListener('dragover',e=>e.preventDefault());
window.addEventListener('drop',e=>e.preventDefault());

function entryFile(entry){return new Promise((resolve,reject)=>entry.file(resolve,reject));}
function directoryBatch(reader){return new Promise((resolve,reject)=>reader.readEntries(resolve,reject));}
async function collectEntryFiles(entry,result){
  if(entry.isFile){const file=await entryFile(entry);file._relativePath=entry.fullPath||file.name;result.push(file);return;}
  if(!entry.isDirectory)return;
  const reader=entry.createReader();
  while(true){const entries=await directoryBatch(reader);if(!entries.length)break;for(const child of entries)await collectEntryFiles(child,result);}
}
async function collectDroppedFiles(dataTransfer){
  const entries=[...dataTransfer.items].map(item=>item.webkitGetAsEntry?item.webkitGetAsEntry():null).filter(Boolean);
  if(!entries.length)return [...dataTransfer.files];
  const result=[];for(const entry of entries)await collectEntryFiles(entry,result);return result;
}
function addFiles(kind,newFiles){
  let added=0;
  for(const f of newFiles){
    if(!/\.(csv|xlsx|zip)$/i.test(f.name))continue;
    const displayName=f.webkitRelativePath||f._relativePath||f.name;
    const key=displayName+':'+f.size+':'+f.lastModified;
    f._displayName=displayName;f._key=key;
    if(!states[kind].some(x=>x._key===key)){states[kind].push(f);added++;}
  }
  renderFiles(kind);updateCompareButton();
  if(added)showToast(added+' ta '+(kind==='old'?'kechagi':'yangi')+' fayl qo‘shildi.','success','add');
}
function renderFiles(kind){
  const list=document.getElementById(kind+'List');list.textContent='';
  states[kind].forEach((f,index)=>{
    const row=document.createElement('div');row.className='file';
    const name=document.createElement('span');name.className='file-name';name.textContent=(index+1)+'. '+(f._displayName||f.name);
    const tools=document.createElement('span');tools.className='file-tools';
    const size=document.createElement('span');size.textContent=(f.size/1024).toFixed(1)+' KB';
    const remove=document.createElement('button');remove.type='button';remove.className='remove small';remove.textContent='Olib tashlash';
    remove.onclick=()=>{states[kind].splice(index,1);renderFiles(kind);updateCompareButton();};
    tools.append(size,remove);row.append(name,tools);list.appendChild(row);
  });
}
function updateCompareButton(){compareBtn.disabled=!states.old.length||!states.new.length;}
function setupPicker(kind){
  const drop=document.getElementById(kind+'Drop');
  const input=document.getElementById(kind+'Input');
  const folder=document.getElementById(kind+'Folder');
  const filePick=drop.querySelector('.file-pick');
  const folderPick=drop.querySelector('.folder-pick');
  drop.onclick=e=>{if(!e.target.closest('button'))input.click();};
  drop.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();input.click();}};
  input.onchange=()=>{addFiles(kind,[...input.files]);input.value='';};
  filePick.onclick=e=>{e.stopPropagation();input.click();};
  folderPick.onclick=e=>{e.stopPropagation();folder.click();};
  folder.onchange=()=>{addFiles(kind,[...folder.files]);folder.value='';};
  drop.ondragover=e=>{e.preventDefault();drop.classList.add('drag');};
  drop.ondragleave=()=>drop.classList.remove('drag');
  drop.ondrop=async e=>{
    e.preventDefault();e.stopPropagation();drop.classList.remove('drag');
    try{
      const found=(await collectDroppedFiles(e.dataTransfer)).filter(f=>/\.(csv|xlsx|zip)$/i.test(f.name));
      if(!found.length)throw new Error('CSV, XLSX yoki ZIP fayl topilmadi. Eski .xls fayl bo‘lsa, uni .xlsx qilib saqlang.');
      addFiles(kind,found);
    }catch(err){const message=err.message||'Papkani o‘qib bo‘lmadi.';showStatus(message,'err');showToast(message,'error','error');}
  };
}
setupPicker('old');setupPicker('new');

function showStatus(message,type){statusEl.className='status '+type;statusEl.textContent=message;}
function appendCells(row,values,numberColumns=[]){
  values.forEach((value,index)=>{const cell=document.createElement('td');if(numberColumns.includes(index))cell.className='num';cell.textContent=value;row.appendChild(cell);});
}
function emptyRow(body,colspan,message){const row=document.createElement('tr');const cell=document.createElement('td');cell.colSpan=colspan;cell.className='empty';cell.textContent=message;row.appendChild(cell);body.appendChild(row);}
function isGasSummary(){return lastResult&&lastResult.report_type==='gas-sales-summary';}
function deltaText(value){const number=Number(value)||0;return number>0?'+'+number:String(number);}
function setTableHead(id,labels,numberColumns=[]){
  const row=document.getElementById(id);row.textContent='';
  labels.forEach((label,index)=>{const cell=document.createElement('th');if(numberColumns.includes(index))cell.className='num';cell.textContent=label;row.appendChild(cell);});
}
function renderSummary(data){
  const body=document.getElementById('summaryBody');body.textContent='';
  if(isGasSummary()){
    if(!data.length){emptyRow(body,4,'Raygaz ma’lumoti topilmadi.');return;}
    data.forEach(item=>{const row=document.createElement('tr');appendCells(row,[item.raygaz,item.old_returned,item.new_returned,deltaText(item.returned_delta)],[1,2,3]);body.appendChild(row);});
    return;
  }
  if(!data.length){emptyRow(body,3,'Tasdiqlangan sotuv topilmadi.');return;}
  data.forEach(item=>{const row=document.createElement('tr');appendCells(row,[item.raygaz,item.sales_count,item.need_total],[1,2]);body.appendChild(row);});
}
function filteredDetails(){
  if(!lastResult)return[];
  const q=document.getElementById('search').value.trim().toLocaleLowerCase('uz');
  const source=isGasSummary()?lastResult.comparison_rows:lastResult.details;
  if(!q)return source;
  if(isGasSummary())return source.filter(item=>[item.organization,item.inspector,item.accepted,item.sold,item.returned,item.returned_delta].some(v=>String(v).toLocaleLowerCase('uz').includes(q)));
  return lastResult.details.filter(item=>[item.raygaz,item.mahalla,item.code,item.subscriber,item.need,item.old_last_sale,item.new_last_sale].some(v=>String(v).toLocaleLowerCase('uz').includes(q)));
}
function renderDetails(){
  const body=document.getElementById('detailsBody');body.textContent='';const data=filteredDetails();
  if(isGasSummary()){
    if(!data.length){emptyRow(body,7,'Qidiruv bo‘yicha ma’lumot topilmadi.');return;}
    data.forEach(item=>{const row=document.createElement('tr');if(item.is_grand_total)row.className='gas-grand-total';else if(item.is_total)row.className='gas-total';appendCells(row,[item.organization,item.inspector,item.accepted,item.sold,item.returned,item.percent,item.is_total?deltaText(item.returned_delta):''],[2,3,4,5,6]);body.appendChild(row);});
    return;
  }
  if(!data.length){emptyRow(body,8,lastResult&&lastResult.details.length?'Qidiruv bo‘yicha ma’lumot topilmadi.':'Tasdiqlangan sotuv topilmadi.');return;}
  data.forEach((item,index)=>{const row=document.createElement('tr');appendCells(row,[index+1,item.raygaz,item.mahalla,item.code,item.subscriber,item.need,item.old_last_sale,item.new_last_sale]);body.appendChild(row);});
}
function renderControls(data){
  const body=document.getElementById('controlBody');body.textContent='';
  if(isGasSummary()){
    if(!data.length){emptyRow(body,3,'Yo‘qolgan, qo‘shilgan yoki kamaygan chilangar qatori yo‘q.');return;}
    data.forEach(item=>{const row=document.createElement('tr');appendCells(row,[item.status,item.organization,item.inspector]);body.appendChild(row);});
    return;
  }
  if(!data.length){emptyRow(body,5,'Nazorat farqi yo‘q.');return;}
  data.forEach(item=>{const row=document.createElement('tr');appendCells(row,[item.status,item.code,item.raygaz,item.mahalla,item.subscriber]);body.appendChild(row);});
}
function renderResult(data){
  lastResult=data;
  if(isGasSummary()){
    document.getElementById('detailsTableWrap').classList.add('gas-combined-wrap');
    document.getElementById('detailsTable').classList.add('gas-combined-table');
    document.getElementById('summarySection').hidden=true;
    document.getElementById('detailsSection').hidden=false;
    document.getElementById('controlSection').hidden=true;
    const newPeriod=/^0+$/.test(data.new_period.trim())?'Yangi holat':data.new_period;
    document.getElementById('dates').textContent='Kechagi hisobot: '+data.old_period+'  →  Yangi hisobot: '+newPeriod;
    document.getElementById('kpis').innerHTML=
      '<div class="kpi"><span>Sotilmagan gaz farqi</span><b>'+deltaText(data.returned_delta)+'</b></div>'+
      '<div class="kpi"><span>Qabul farqi</span><b>'+deltaText(data.accepted_delta)+'</b></div>'+
      '<div class="kpi"><span>Sotuv farqi</span><b>'+deltaText(data.sold_delta)+'</b></div>'+
      '<div class="kpi"><span>O‘zgargan qator</span><b>'+data.changed_count+'</b></div>';
    document.getElementById('detailsTitle').textContent='Kechagi holat va yangi holat bilan solishtirish';
    document.getElementById('search').placeholder='Raygaz yoki chilangarni qidiring...';
    document.getElementById('downloadExcel').hidden=false;
    setTableHead('detailsHead',['Raygaz','Chilangar','Kechagi qabul','Kechagi sotuv','Kechagi sotilmagan','%','Farq (yangi − kechagi)'],[2,3,4,5,6]);
  }else{
    document.getElementById('detailsTableWrap').classList.remove('gas-combined-wrap');
    document.getElementById('detailsTable').classList.remove('gas-combined-table');
    document.getElementById('summarySection').hidden=false;
    document.getElementById('detailsSection').hidden=false;
    document.getElementById('controlSection').hidden=false;
    document.getElementById('dates').textContent='Kechagi holat: '+data.old_date+'  →  Yangi holat: '+data.new_date;
    document.getElementById('kpis').innerHTML=
      '<div class="kpi"><span>Sotilgan abonent</span><b>'+data.total_sales+'</b></div>'+
      '<div class="kpi"><span>Sotuv qilgan Raygaz</span><b>'+data.raygaz_count+'</b></div>'+
      '<div class="kpi"><span>Yangi holatda yo‘q</span><b>'+data.missing_count+'</b></div>'+
      '<div class="kpi"><span>Yangi qo‘shilgan kod</span><b>'+data.added_count+'</b></div>';
    document.getElementById('summaryTitle').textContent='Raygazlar bo‘yicha sotuv';
    document.getElementById('detailsTitle').textContent='Sotilgan abonentlar';
    document.getElementById('controlTitle').textContent='Nazorat';
    document.getElementById('controlNote').textContent='Yangi holatda topilmagan kod sotuv deb taxmin qilinmaydi. U va kechagi holatda bo‘lmagan yangi kodlar quyida alohida ko‘rsatiladi.';
    document.getElementById('search').placeholder='Kod, abonent, Raygaz yoki mahalla...';
    document.getElementById('downloadExcel').hidden=true;
    setTableHead('summaryHead',['Raygaz','Sotilgan abonent','Jami ehtiyoj'],[1,2]);
    setTableHead('detailsHead',['№','Raygaz','Mahalla','Abonent kod','Abonent','Ehtiyoj','Eski realizatsiya','Yangi realizatsiya']);
    setTableHead('controlHead',['Holat','Abonent kod','Raygaz','Mahalla','Abonent']);
  }
  if(!isGasSummary())renderSummary(data.summary);
  renderDetails();
  if(!isGasSummary())renderControls(data.controls);
  resultsEl.classList.add('show');
}
document.getElementById('search').oninput=renderDetails;

compareBtn.onclick=async()=>{
  if(!states.old.length||!states.new.length)return;
  const totalBytes=[...states.old,...states.new].reduce((sum,f)=>sum+f.size+1024,0);
  if(totalBytes>maxUploadBytes){const message='Fayllar jami '+uploadLimitMb+' MB limitdan oshdi.';showStatus(message,'err');showToast(message,'error','error');return;}
  const fd=new FormData();states.old.forEach(f=>fd.append('old_files',f,f.name));states.new.forEach(f=>fd.append('new_files',f,f.name));
  compareBtn.disabled=true;compareBtn.setAttribute('aria-busy','true');compareBtn.textContent='Hisoblanmoqda…';resultsEl.classList.remove('show');showStatus('Ikki holat tekshirilmoqda va sotuvlar hisoblanmoqda...','work');
  try{
    const response=await fetch('/sotuvlar/compare',{method:'POST',body:fd});
    const text=await response.text();let data;try{data=JSON.parse(text);}catch(_){throw new Error(text||('HTTP '+response.status));}
    if(!response.ok)throw new Error(data.error||('HTTP '+response.status));
    renderResult(data);
    if(data.report_type==='gas-sales-summary'){
      showStatus('Tayyor. Qabul, sotuv va sotilmagan gaz miqdorlari solishtirildi.','ok');
    }else{
      showStatus('Tayyor. '+data.total_sales+' ta tasdiqlangan sotuv aniqlandi.','ok');
    }
    showToast('Solishtirish yakunlandi. Natijalar tayyor.','success','done');
  }catch(err){const message=err.message||'Hisoblashda xatolik yuz berdi.';showStatus(message,'err');showToast(message,'error','error');}
  finally{compareBtn.removeAttribute('aria-busy');compareBtn.textContent='Sotuvlarni hisoblash';updateCompareButton();}
};

document.getElementById('clear').onclick=()=>{
  states.old=[];states.new=[];lastResult=null;renderFiles('old');renderFiles('new');updateCompareButton();
  statusEl.className='status';statusEl.textContent='';resultsEl.classList.remove('show');document.getElementById('search').value='';document.getElementById('downloadExcel').hidden=true;
};

function csvCell(value){const s=String(value??'');return /[";,\n\r]/.test(s)?'"'+s.replaceAll('"','""')+'"':s;}
document.getElementById('downloadExcel').onclick=async()=>{
  if(!isGasSummary()||!states.old.length||!states.new.length)return;
  const button=document.getElementById('downloadExcel');
  const fd=new FormData();states.old.forEach(f=>fd.append('old_files',f,f.name));states.new.forEach(f=>fd.append('new_files',f,f.name));
  button.disabled=true;button.setAttribute('aria-busy','true');button.textContent='Excel tayyorlanmoqda…';showStatus('Rasmdagidek solishtirish Exceli tayyorlanmoqda...','work');
  try{
    const response=await fetch('/sotuvlar/compare-xlsx',{method:'POST',body:fd});
    if(!response.ok){const text=await response.text();let message=text||('HTTP '+response.status);try{message=JSON.parse(text).error||message;}catch(_){}throw new Error(message);}
    const blob=await response.blob();const link=document.createElement('a');link.href=URL.createObjectURL(blob);link.download=response.headers.get('X-Output-Name')||'QABUL_SOTUV_SOLISHTIRISH.xlsx';document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(link.href),1000);
    showStatus('Tayyor. G ustunida har bir ЖАМИ qatorining sotilmagan gaz farqi yozilgan Excel yuklandi.','ok');
    showToast('Solishtirish Exceli tayyorlandi va yuklandi.','success','done');
  }catch(err){const message=err.message||'Excel tayyorlashda xatolik yuz berdi.';showStatus(message,'err');showToast(message,'error','error');}
  finally{button.disabled=false;button.removeAttribute('aria-busy');button.textContent='Rasmdagidek Excel yuklash';}
};
document.getElementById('download').onclick=()=>{
  if(!lastResult)return;
  let headers,lines,filename;
  if(isGasSummary()){
    headers=['Райгаз','Чилангар','Кечаги қабул','Кечаги сотув','Кечаги сотилмаган','%','Фарқ (янги - кечаги)'];
    lines=[headers.map(csvCell).join(';')];
    lastResult.comparison_rows.forEach(item=>lines.push([item.organization,item.inspector,item.accepted,item.sold,item.returned,item.percent,item.is_total?item.returned_delta:''].map(csvCell).join(';')));
    filename='QABUL_SOTUV_SOLISHTIRISH.csv';
  }else{
    headers=['№','Райгаз','Маҳалла','Абонент код','Абонент','Еҳтиёж','Кечаги сўнгги реализация','Янги сўнгги реализация'];
    lines=[headers.map(csvCell).join(';')];
    lastResult.details.forEach((item,index)=>lines.push([index+1,item.raygaz,item.mahalla,item.code,item.subscriber,item.need,item.old_last_sale,item.new_last_sale].map(csvCell).join(';')));
    filename='SOTUVLAR_'+lastResult.old_date+'_'+lastResult.new_date+'.csv';
  }
  const blob=new Blob(['\ufeff'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'});
  const link=document.createElement('a');link.href=URL.createObjectURL(blob);link.download=filename;document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(link.href),1000);
  showToast('CSV natija yuklandi.','success','done');
};
</script>
</body>
</html>
"""


EGAZ_HTML = r"""
<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#102f45">
<title>E-GAZ hisoboti — SVOD TIZIMI</title>
<style>
:root{--navy:#102f45;--navy-2:#174d66;--blue:#137a8b;--green:#15805d;--green-dark:#0d6549;--light:#f3f7f8;--surface:#fff;--border:#d8e4e7;--danger:#b42318;--warn:#9a5b09;--ink:#142b36;--muted:#637b86;--shadow:0 16px 45px rgba(16,47,69,.1)}
*{box-sizing:border-box}body{margin:0;min-height:100vh;background:radial-gradient(circle at 8% 0%,rgba(19,122,139,.12),transparent 30rem),radial-gradient(circle at 92% 18%,rgba(21,128,93,.1),transparent 26rem),var(--light);font-family:Inter,"Segoe UI",Arial,sans-serif;color:var(--ink);line-height:1.5}button,input{font:inherit}
.top{background:linear-gradient(115deg,#102f45,#15556a 62%,#146f70);color:#fff;padding:19px 18px 20px}.top-inner{max-width:1180px;margin:auto;display:flex;align-items:center;justify-content:space-between;gap:20px}.eyebrow{font-size:10px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:#84e0c4;margin-bottom:2px}h1{margin:0;font-size:29px;line-height:1.1;letter-spacing:-.025em}.sub{opacity:.83;font-size:13px;margin-top:4px}
.nav{display:flex;gap:12px;flex-wrap:wrap;justify-content:flex-end;padding:0;background:transparent;border:0;border-radius:0}.nav a{color:#fff;text-decoration:none;font-weight:750;font-size:12px;letter-spacing:.025em;padding:10px 14px;border:1px solid rgba(255,255,255,.34);border-radius:10px;white-space:nowrap;box-shadow:0 5px 14px rgba(4,28,40,.2);transition:transform .18s,filter .18s}.nav a:nth-child(1){background:#2563eb}.nav a:nth-child(2){background:#d97706}.nav a:nth-child(3){background:#15803d}.nav a:hover{transform:translateY(-1px);filter:brightness(1.1)}.nav a.active{color:#fff;border-color:#fff;box-shadow:0 0 0 2px rgba(255,255,255,.25),0 6px 16px rgba(4,28,40,.28)}
.wrap{max-width:1050px;margin:22px auto;padding:0 18px}.card{background:rgba(255,255,255,.96);border:1px solid rgba(184,206,211,.8);border-radius:20px;padding:clamp(20px,4vw,34px);box-shadow:var(--shadow)}.intro{display:flex;justify-content:space-between;align-items:flex-start;gap:22px;margin-bottom:20px}.intro h2{margin:0 0 6px;font-size:25px;letter-spacing:-.025em}.muted{color:var(--muted);font-size:14px}.mode-chip{display:flex;align-items:center;gap:8px;max-width:390px;background:#e9f8f2;color:var(--green-dark);border:1px solid #bde6d8;border-radius:999px;padding:8px 12px;font-size:12px;font-weight:800}.mode-chip span{width:8px;height:8px;flex:0 0 auto;border-radius:50%;background:var(--green)}
.drop{position:relative;border:2px dashed #69b690;border-radius:17px;padding:28px 24px;display:flex;align-items:center;justify-content:center;gap:18px;text-align:left;cursor:pointer;background:linear-gradient(135deg,#e9f8f1,#f8fdfb);box-shadow:inset 0 0 0 1px rgba(21,128,93,.04);transition:.2s;outline:none}.drop:hover,.drop.drag{border-color:var(--green);background:#e2f6ec;box-shadow:0 0 0 5px rgba(21,128,93,.1);transform:translateY(-2px)}.drop input{display:none}.drop-symbol{width:48px;height:48px;display:grid;place-items:center;flex:0 0 auto;border-radius:14px;background:#d9f1e5;color:var(--green);box-shadow:inset 0 0 0 1px rgba(21,128,93,.16)}.drop-symbol svg{width:25px;height:25px}.drop-copy{min-width:0}.drop strong{display:block;color:var(--navy);font-size:19px;margin:0 0 6px}.drop-meta{display:flex;align-items:center;gap:10px;flex-wrap:wrap}.badge{display:inline-block;margin:0;padding:5px 9px;border:1px solid #9fcfb9;border-radius:7px;background:#fff;color:#28654d;font-size:11px;font-weight:800;letter-spacing:.06em}
.file{display:none;align-items:center;justify-content:space-between;gap:12px;margin-top:14px;padding:12px 14px;border:1px solid var(--border);border-radius:12px;background:#f9fcfc}.file.show{display:flex}.file-name{min-width:0;overflow-wrap:anywhere;font-size:13px;font-weight:700}.file-size{white-space:nowrap;color:var(--muted);font-size:12px}.actions{display:flex;align-items:center;gap:11px;flex-wrap:wrap;margin-top:18px}button{border:0;border-radius:11px;padding:12px 18px;font-weight:800;cursor:pointer;transition:.18s}.primary{background:linear-gradient(135deg,var(--green),#1a9470);color:#fff;box-shadow:0 8px 18px rgba(21,128,93,.22)}.primary:disabled{opacity:.45;cursor:not-allowed;box-shadow:none}.secondary{background:#e8f0f2;color:#244d5b}.status{position:relative;display:none;margin-top:18px;padding:14px 16px 14px 44px;border-radius:12px;white-space:pre-wrap}.status::before{position:absolute;left:16px;top:14px;font-weight:900}.status.ok{display:block;background:#eaf8f1;color:#126342;border:1px solid #b9e2d0}.status.ok::before{content:"✓"}.status.err{display:block;background:#fff1ef;color:var(--danger);border:1px solid #f0c9c4}.status.err::before{content:"!"}.status.work{display:block;background:#fff8e9;color:var(--warn);border:1px solid #efd89f}.status.work::before{content:"";width:15px;height:15px;border:2px solid currentColor;border-right-color:transparent;border-radius:50%;top:16px;animation:spin .75s linear infinite}
.results{display:none;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:20px}.results.show{display:grid}.kpi{background:linear-gradient(145deg,#f9fcfc,#f1f7f8);border:1px solid var(--border);border-radius:14px;padding:15px}.kpi span{font-size:12px;color:var(--muted);font-weight:700}.kpi b{display:block;color:var(--navy);font-size:24px;margin-top:3px;font-variant-numeric:tabular-nums}.note{margin-top:18px;padding:13px 15px;border-left:4px solid var(--blue);border-radius:9px;background:#edf7f8;color:#365b66;font-size:13px}.footer{text-align:center;color:#718892;font-size:12px;margin:19px}.footer strong{color:#46636e}@keyframes spin{to{transform:rotate(360deg)}}
@media(max-width:800px){.top-inner,.intro{align-items:flex-start;flex-direction:column}.nav{justify-content:flex-start;width:100%}.nav a{flex:1;text-align:center}.results{grid-template-columns:repeat(2,1fr)}}@media(max-width:500px){.nav a{font-size:10px;padding:9px 6px}.drop{flex-direction:column;padding:22px 16px;text-align:center}.drop-meta{justify-content:center}.results{grid-template-columns:1fr}.file{align-items:flex-start;flex-direction:column}.actions button{width:100%}}
</style>
</head>
<body>
<header class="top"><div class="top-inner">
  <div><div class="eyebrow">Nazorat • Tahlil • Hisobot</div><h1>SVOD TIZIMI</h1><div class="sub">E-GAZ eksportidan tayyor shakldagi Excel yaratish</div></div>
  <nav class="nav" aria-label="Asosiy bo‘limlar">
    <a href="/">SVOD HISOBOTI</a>
    <a href="/sotuvlar">SOTILGAN GAZLAR</a>
    <a class="active" aria-current="page" href="/egaz">E-GAZ HISOBOTI</a>
    <a href="/gnp-taqqoslash">GNP + MFY SVOD</a>
  </nav>
</div></header>
<main class="wrap">
  <section class="card">
    <div class="intro">
      <div><h2>E-GAZ hisobotini tayyorlash</h2></div>
    </div>
    <div id="drop" class="drop" tabindex="0" role="button" aria-label="E-GAZ XLSX faylini tanlash">
      <span class="drop-symbol" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4"/><path d="m7 9 5-5 5 5"/><path d="M5 20h14"/></svg></span>
      <div class="drop-copy"><strong>E-GAZ Billing Excel faylini shu yerga tashlang</strong>
        <div class="drop-meta"><span class="muted">yoki bosib kompyuterdan tanlang</span><span class="badge">XLSX</span></div>
      </div>
      <input id="input" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">
    </div>
    <div id="file" class="file"><div id="fileName" class="file-name"></div><div id="fileSize" class="file-size"></div></div>
    <div class="actions">
      <button id="generate" class="primary" disabled>Hisobotni yaratish va yuklash</button>
      <button id="clear" class="secondary" type="button">Tozalash</button>
    </div>
    <div id="status" class="status" role="status" aria-live="polite"></div>
    <div id="results" class="results">
      <div class="kpi"><span>Chilangar qatori</span><b id="rows">0</b></div>
      <div class="kpi"><span>Qabul qilindi</span><b id="accepted">0</b></div>
      <div class="kpi"><span>Sotildi</span><b id="sold">0</b></div>
      <div class="kpi"><span>Sotilmagan gaz</span><b id="unsold">0</b></div>
    </div>
    <div class="note">Davr fayl nomidan olinadi: <b>YYYY-MM-DD</b> yoki <b>1-31 Avgust</b>. Masalan, <b>1-31 Avgust.xlsx</b> → <b>1-31 Avgust holatiga</b>.</div>
  </section>
  <div class="footer"><strong>SVOD TIZIMI v1.3</strong> — {{ privacy_notice }}</div>
</main>
<script>
const maxUploadBytes={{ upload_limit_bytes }};let selected=null;
const input=document.getElementById('input'),drop=document.getElementById('drop'),button=document.getElementById('generate'),statusEl=document.getElementById('status'),fileEl=document.getElementById('file'),results=document.getElementById('results');
function showStatus(message,type){statusEl.className='status '+type;statusEl.textContent=message}
function selectFile(file){if(!file||!/\.xlsx$/i.test(file.name)){showStatus('Faqat .xlsx formatidagi E-GAZ Billing faylini tanlang.','err');return}selected=file;document.getElementById('fileName').textContent=file.name;document.getElementById('fileSize').textContent=(file.size/1024).toFixed(1)+' KB';fileEl.classList.add('show');button.disabled=false;statusEl.className='status';results.classList.remove('show')}
drop.onclick=()=>input.click();drop.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();input.click()}};input.onchange=()=>{selectFile(input.files[0]);input.value=''};
drop.ondragover=e=>{e.preventDefault();drop.classList.add('drag')};drop.ondragleave=()=>drop.classList.remove('drag');drop.ondrop=e=>{e.preventDefault();drop.classList.remove('drag');selectFile(e.dataTransfer.files[0])};window.addEventListener('dragover',e=>e.preventDefault());window.addEventListener('drop',e=>e.preventDefault());
document.getElementById('clear').onclick=()=>{selected=null;button.disabled=true;fileEl.classList.remove('show');statusEl.className='status';results.classList.remove('show')};
button.onclick=async()=>{if(!selected)return;if(selected.size>maxUploadBytes){showStatus('Fayl ruxsat etilgan hajmdan katta.','err');return}const fd=new FormData();fd.append('file',selected,selected.name);button.disabled=true;button.textContent='Hisobot tayyorlanmoqda…';showStatus('E-GAZ ma’lumotlari tekshirilib, Excel yaratilmoqda...','work');
try{const response=await fetch('/egaz/generate',{method:'POST',body:fd});if(!response.ok){const text=await response.text();let message=text||('HTTP '+response.status);try{message=JSON.parse(text).error||message}catch(_){}throw new Error(message)}const blob=await response.blob();const link=document.createElement('a');link.href=URL.createObjectURL(blob);link.download=response.headers.get('X-Output-Name')||'E-GAZ_HISOBOTI.xlsx';document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(link.href),1000);document.getElementById('rows').textContent=response.headers.get('X-Egaz-Rows')||'0';document.getElementById('accepted').textContent=response.headers.get('X-Egaz-Accepted')||'0';document.getElementById('sold').textContent=response.headers.get('X-Egaz-Sold')||'0';document.getElementById('unsold').textContent=response.headers.get('X-Egaz-Unsold')||'0';results.classList.add('show');showStatus('Tayyor. E-GAZ hisoboti yaratildi va yuklandi.','ok')}catch(error){showStatus(error.message||'Hisobot yaratishda xatolik yuz berdi.','err')}finally{button.disabled=!selected;button.textContent='Hisobotni yaratish va yuklash'}};
</script>
</body>
</html>
"""


ANALYSIS_PAGE_HTML = r"""
<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#102f45">
<title>{{ page_title }} — SVOD TIZIMI</title>
<style>
:root{--navy:#102f45;--navy2:#174d66;--blue:#137a8b;--green:#15805d;--light:#f3f7f8;--ink:#142b36;--muted:#637b86;--border:#d8e4e7;--danger:#b42318}
*{box-sizing:border-box}body{margin:0;background:var(--light);font:15px/1.5 "Segoe UI",Arial,sans-serif;color:var(--ink)}
.top{background:linear-gradient(125deg,#0c293d,var(--navy2) 58%,#126e70);color:#fff;padding:22px 18px}
.top-inner{max-width:1120px;margin:auto;display:flex;align-items:center;justify-content:space-between;gap:24px}.brand h1{margin:0;font-size:28px}.sub{font-size:13px;opacity:.85}
.nav{display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end}.nav a{color:#fff;text-decoration:none;font-size:12px;font-weight:800;padding:9px 11px;border:1px solid #ffffff44;border-radius:10px;white-space:nowrap}.nav a.active{background:#fff;color:var(--navy)}
.wrap{max-width:1040px;margin:28px auto;padding:0 16px}.card{background:#fff;border:1px solid var(--border);border-radius:20px;padding:clamp(18px,3vw,30px);box-shadow:0 14px 40px #102f4512}.intro{margin:0 0 22px;color:var(--muted);font-size:14px}
.files{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}.filebox{display:block;border:1px dashed #8eb4bb;border-radius:14px;padding:18px;background:#f9fcfc}.filebox strong{display:block;margin-bottom:7px}.filebox input{display:block;width:100%;margin-top:10px}.hint{font-size:12px;color:var(--muted);margin-top:7px}
.files.gnp-upload .filebox{position:relative;display:flex;flex-direction:column;min-width:0;min-height:224px;padding:22px;border:1px solid #cbdfe5;border-radius:18px;overflow:hidden;box-shadow:0 8px 22px rgba(16,47,69,.06);transition:transform .2s,box-shadow .2s,border-color .2s}
.files.gnp-upload .filebox::before{content:"";position:absolute;left:0;right:0;top:0;height:5px;background:var(--upload-accent)}
.files.gnp-upload .filebox:hover{transform:translateY(-2px);box-shadow:0 14px 28px rgba(16,47,69,.1)}
.files.gnp-upload .filebox.dragover{transform:translateY(-2px);border-color:var(--upload-accent);border-style:solid;box-shadow:0 0 0 4px rgba(19,122,139,.13),0 14px 28px rgba(16,47,69,.1)}
.files.gnp-upload .filebox.selected{border:2px solid var(--upload-accent);box-shadow:0 0 0 4px rgba(19,122,139,.09),0 12px 26px rgba(16,47,69,.09)}
.filebox-egaz.selected{background:linear-gradient(145deg,#eaf3ff,#f8fbff)}
.filebox-gnp.selected{background:linear-gradient(145deg,#e8f7ef,#f7fcf9)}
.filebox-egaz{--upload-accent:#2877c7;--upload-soft:#edf5ff;background:linear-gradient(145deg,#f4f9ff,#fff)}
.filebox-gnp{--upload-accent:#16805c;--upload-soft:#eaf8f1;background:linear-gradient(145deg,#f1fbf5,#fff)}
.upload-kicker{display:flex;align-items:center;gap:9px;margin-bottom:12px;color:var(--upload-accent);font-size:11px;font-weight:850;letter-spacing:.13em;text-transform:uppercase}
.upload-step{display:grid;place-items:center;width:29px;height:29px;border-radius:9px;background:var(--upload-soft);font-size:12px;letter-spacing:0}
.filebox.selected .upload-step{background:var(--upload-accent);color:#fff}
.files.gnp-upload .filebox strong{margin:0 0 5px;color:var(--navy);font-size:19px;line-height:1.3}
.filebox-desc{max-width:52ch;color:#526d78;font-size:13px;line-height:1.55}
.upload-drop-hint{margin-top:7px;color:#738a93;font-size:11px}
.file-picker{display:flex;align-items:center;gap:12px;min-width:0;margin-top:auto;padding-top:20px}
.file-select{flex:0 0 auto;background:var(--upload-accent);border-radius:9px;padding:9px 14px;font-size:13px;box-shadow:0 5px 12px rgba(16,47,69,.12)}
.file-select:hover{filter:brightness(1.06)}
.file-name{min-width:0;overflow:hidden;color:#657e88;font-size:12px;text-overflow:ellipsis;white-space:nowrap}
.filebox.selected .file-name{color:var(--upload-accent);font-weight:800}
.files.gnp-upload .filebox input[type=file]{display:none}
.selection-toast{position:fixed;z-index:1000;top:14px;left:50%;display:flex;align-items:center;gap:11px;width:max-content;max-width:calc(100vw - 28px);padding:12px 17px;border:1px solid #b8e0d1;border-radius:14px;background:rgba(245,253,249,.98);color:#105d40;box-shadow:0 14px 38px rgba(16,47,69,.2);font-size:14px;font-weight:800;opacity:0;transform:translate(-50%,-150%);transition:opacity .22s,transform .22s;pointer-events:none;backdrop-filter:blur(10px)}
.selection-toast.show{opacity:1;transform:translate(-50%,0)}
.selection-toast-icon{display:grid;place-items:center;width:24px;height:24px;flex:0 0 auto;border-radius:50%;background:#d9f3e7;color:#0d6549;font-size:14px}
.actions{display:flex;gap:10px;align-items:center;margin-top:20px;flex-wrap:wrap}button{border:0;border-radius:11px;padding:12px 18px;font:inherit;font-weight:800;background:var(--green);color:#fff;cursor:pointer}button:disabled{opacity:.55;cursor:wait}.status{margin:18px 0 0;padding:12px 14px;border-radius:10px;background:#f3f7f8;color:var(--muted);min-height:44px}.status.err{background:#fff1f0;color:var(--danger)}.status.ok{background:#e9f8f2;color:#0d6549}.footer{max-width:1040px;margin:0 auto 20px;padding:0 16px;text-align:center;font-size:12px;color:var(--muted)}
@media(max-width:760px){.top-inner{align-items:flex-start;flex-direction:column}.nav{justify-content:flex-start}.files{grid-template-columns:1fr}.files.gnp-upload .filebox{min-height:205px;padding:20px}.nav a{font-size:11px;padding:8px 9px}}
@media(max-width:480px){.selection-toast{top:10px;padding:10px 13px;font-size:12px}}
</style>
</head>
<body>
<header class="top"><div class="top-inner"><div class="brand"><h1>SVOD TIZIMI</h1><div class="sub">Nazorat · tahlil · hisobot</div></div>
<nav class="nav" aria-label="Asosiy bo‘limlar">
  <a class="{{ 'active' if active == 'main' else '' }}" href="/">SVOD HISOBOTI</a>
  <a class="{{ 'active' if active == 'sales' else '' }}" href="/sotuvlar">SOTILGAN GAZLAR</a>
  <a class="{{ 'active' if active == 'egaz' else '' }}" href="/egaz">E-GAZ HISOBOTI</a>
  <a class="{{ 'active' if active == 'gnp' else '' }}" href="/gnp-taqqoslash">GNP + MFY SVOD</a>
</nav></div></header>
<div id="selectionToast" class="selection-toast" role="status" aria-live="polite" aria-atomic="true"><span class="selection-toast-icon" aria-hidden="true">✓</span><span id="selectionToastText"></span></div>
<main class="wrap"><section class="card">
  <p class="intro">{{ page_description }}</p>
  {% if mode == 'mfy' %}
  <div class="files"><label class="filebox"><strong>E-GAZ CSV fayli</strong><span>Sana, tuman/Raygaz, ariza raqami, MFY va summa ustunlari bo‘lgan .csv</span><input id="csvFile" type="file" accept=".csv,text/csv"><span class="hint" id="csvName">Fayl tanlanmagan</span></label></div>
  {% else %}
  <div class="files gnp-upload">
    <section class="filebox filebox-egaz" aria-labelledby="csvHeading">
      <div class="upload-kicker"><span class="upload-step">01</span><span>E-GAZ CSV / ZIP</span></div>
      <strong id="csvHeading">E-GAZ sotuvlari</strong>
      <span class="filebox-desc">Bitta yoki bir nechta CSV faylni, yoki ularni jamlagan ZIP faylni tanlang. Sana, tuman, ariza raqami va MFY ustunlari kerak. Summa shart emas.</span>
      <span class="upload-drop-hint">Fayllarni shu kartaga sudrab tashlashingiz ham mumkin.</span>
      <input id="csvFile" type="file" accept=".csv,.zip,text/csv,application/zip" multiple>
      <div class="file-picker"><button id="csvPick" class="file-select" type="button" aria-controls="csvFile">E-GAZ faylini tanlash</button><span class="file-name" id="csvName" aria-live="polite">Hali fayl tanlanmagan</span></div>
    </section>
    <section class="filebox filebox-gnp" aria-labelledby="gnpHeading">
      <div class="upload-kicker"><span class="upload-step">02</span><span>GNP XLSX</span></div>
      <strong id="gnpHeading">GNP buyurtmalari</strong>
      <span class="filebox-desc">Ariza raqami, buyurtma ballonlari va o‘tkazilgan ballonlar ustunlari bo‘lgan Excel faylni tanlang.</span>
      <span class="upload-drop-hint">Bitta XLSX faylni shu kartaga sudrab tashlang.</span>
      <input id="gnpFile" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">
      <div class="file-picker"><button id="gnpPick" class="file-select" type="button" aria-controls="gnpFile">GNP Excel faylini tanlash</button><span class="file-name" id="gnpName" aria-live="polite">Hali fayl tanlanmagan</span></div>
    </section>
  </div>
  {% endif %}
  <div class="actions"><button id="generate" type="button" disabled>Excel hisobotini yuklab olish</button></div>
  <div id="status" class="status" role="status" aria-live="polite">{{ upload_hint }}</div>
</section></main>
<footer class="footer">Mahalliy rejimda fayllar kompyuteringizda qayta ishlanadi. {{ privacy_notice }}</footer>
<script>
const mode={{ mode|tojson }},limit={{ upload_limit_bytes }},csvInput=document.getElementById('csvFile'),gnpInput=document.getElementById('gnpFile'),button=document.getElementById('generate'),statusEl=document.getElementById('status');
let toastTimer=null,confirmationAudioContext=null;
function prepareConfirmationAudio(){const AudioContextClass=window.AudioContext||window.webkitAudioContext;if(!AudioContextClass)return;try{if(!confirmationAudioContext)confirmationAudioContext=new AudioContextClass();if(confirmationAudioContext.state==='suspended')confirmationAudioContext.resume().catch(()=>{})}catch(error){}}
async function playGentleConfirmationSound(){const AudioContextClass=window.AudioContext||window.webkitAudioContext;if(!AudioContextClass)return;try{if(!confirmationAudioContext)confirmationAudioContext=new AudioContextClass();const context=confirmationAudioContext;if(context.state==='suspended')await context.resume();const start=context.currentTime+.01,master=context.createGain();master.gain.setValueAtTime(.0001,start);master.gain.exponentialRampToValueAtTime(.065,start+.025);master.gain.exponentialRampToValueAtTime(.0001,start+.58);master.connect(context.destination);[[784,0],[988,.12]].forEach(([frequency,offset])=>{const tone=context.createOscillator(),volume=context.createGain();tone.type='sine';tone.frequency.setValueAtTime(frequency,start+offset);volume.gain.setValueAtTime(.0001,start+offset);volume.gain.exponentialRampToValueAtTime(.35,start+offset+.018);volume.gain.exponentialRampToValueAtTime(.0001,start+offset+.28);tone.connect(volume);volume.connect(master);tone.start(start+offset);tone.stop(start+offset+.3)});window.setTimeout(()=>master.disconnect(),700)}catch(error){}}
function showSelectionToast(message){const toast=document.getElementById('selectionToast');document.getElementById('selectionToastText').textContent=message;toast.classList.add('show');window.clearTimeout(toastTimer);toastTimer=window.setTimeout(()=>toast.classList.remove('show'),3200);playGentleConfirmationSound()}
function update(changedInput=null){const csvs=Array.from(csvInput.files),hasGnp=!!gnpInput?.files.length,csvBox=document.querySelector('.filebox-egaz'),gnpBox=document.querySelector('.filebox-gnp');document.getElementById('csvName').textContent=csvs.length?csvs.map(file=>file.name).join(', '):'E-GAZ fayli tanlanmagan';if(gnpInput)document.getElementById('gnpName').textContent=gnpInput.files[0]?.name||'GNP XLSX tanlanmagan';csvBox?.classList.toggle('selected',csvs.length>0);gnpBox?.classList.toggle('selected',hasGnp);button.disabled=!csvs.length||(mode==='gnp'&&!hasGnp);const missing=[];if(!csvs.length)missing.push('E-GAZ CSV/ZIP');if(mode==='gnp'&&!hasGnp)missing.push('GNP XLSX');if(missing.length)setStatus('Tanlanmagan fayl: '+missing.join(' va ')+'.','');else setStatus('Ikkala fayl tanlandi. Hisobot yaratishga tayyor.','ok');if(changedInput){let message=changedInput===csvInput?(csvs.length===1?'E-GAZ fayli biriktirildi.':'E-GAZ fayllari biriktirildi ('+csvs.length+' ta).'):'GNP buyurtma fayli biriktirildi.';if(csvs.length&&hasGnp)message='Ikkala fayl biriktirildi. Hisobot tayyor.';showSelectionToast(message)}}
function setStatus(text,type=''){statusEl.className='status '+type;statusEl.textContent=text}
function assignDroppedFiles(input,incoming){const dropped=Array.from(incoming||[]);if(!dropped.length)return;const isCsv=input===csvInput,allowed=isCsv?/\.(csv|zip)$/i:/\.xlsx$/i;if(dropped.some(file=>!allowed.test(file.name))){setStatus(isCsv?'E-GAZ uchun .csv yoki .zip fayl tashlang.':'GNP uchun .xlsx fayl tashlang.','err');return}if(!isCsv&&dropped.length!==1){setStatus('GNP uchun bir vaqtda faqat bitta .xlsx fayl tashlang.','err');return}const existing=isCsv&&input.multiple?Array.from(input.files):[];const seen=new Set();const allFiles=[...existing,...dropped].filter(file=>{const key=[file.name,file.size,file.lastModified].join('|');if(seen.has(key))return false;seen.add(key);return true});try{const transfer=new DataTransfer();allFiles.forEach(file=>transfer.items.add(file));input.files=transfer.files;input.dispatchEvent(new Event('change',{bubbles:true}))}catch(error){setStatus('Faylni biriktirib bo‘lmadi. “Fayl tanlash” tugmasidan foydalaning.','err')}}
csvInput.addEventListener('change',event=>update(event.currentTarget));if(gnpInput)gnpInput.addEventListener('change',event=>update(event.currentTarget));document.getElementById('csvPick')?.addEventListener('click',()=>{prepareConfirmationAudio();csvInput.click()});document.getElementById('gnpPick')?.addEventListener('click',()=>{prepareConfirmationAudio();gnpInput.click()});
document.querySelectorAll('.files.gnp-upload .filebox').forEach(box=>{const input=box.querySelector('input[type="file"]');box.addEventListener('dragenter',event=>{event.preventDefault();box.classList.add('dragover')});box.addEventListener('dragover',event=>{event.preventDefault();if(event.dataTransfer)event.dataTransfer.dropEffect='copy';box.classList.add('dragover')});box.addEventListener('dragleave',event=>{if(!event.relatedTarget||!box.contains(event.relatedTarget))box.classList.remove('dragover')});box.addEventListener('drop',event=>{event.preventDefault();box.classList.remove('dragover');assignDroppedFiles(input,event.dataTransfer?.files)})});
button.addEventListener('click',async()=>{const csvs=Array.from(csvInput.files),gnp=gnpInput?.files[0],csvSize=csvs.reduce((total,file)=>total+file.size,0);if(!csvs.length||csvSize>limit||(gnp&&gnp.size+csvSize>limit)){setStatus('Fayllar hajmi ruxsat etilgan limitdan oshdi.','err');return}const fd=new FormData();for(const csv of csvs)fd.append('csv_file',csv,csv.name);if(gnp)fd.append('gnp_file',gnp,gnp.name);button.disabled=true;setStatus('Fayllar tekshirilib, Excel tayyorlanmoqda...');try{const url=mode==='mfy'?'/mfy-svod/generate':'/gnp-taqqoslash/generate';const res=await fetch(url,{method:'POST',body:fd});if(!res.ok){const t=await res.text();let message=t;try{message=JSON.parse(t).error||t}catch(_){}throw new Error(message||('HTTP '+res.status))}const blob=await res.blob();const link=document.createElement('a');link.href=URL.createObjectURL(blob);link.download=res.headers.get('X-Output-Name')||'SVOD_HISOBOTI.xlsx';document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(link.href),1000);const rows=res.headers.get('X-Source-Rows');const amount=res.headers.get('X-Source-Amount');const matched=res.headers.get('X-Matched');const mismatches=res.headers.get('X-Mismatches');let msg='Excel tayyor va yuklandi.';if(rows)msg+=' Sotuv satrlari: '+rows+'.';if(amount&&mode==='mfy')msg+=' Jami summa: '+amount+' so‘m.';if(matched!==null)msg+=' MOS: '+matched+'.';if(mismatches!==null)msg+=' Tafovutli arizalar: '+mismatches+'.';setStatus(msg,'ok')}catch(e){setStatus(e.message||'Hisobot yaratishda xatolik yuz berdi.','err')}finally{button.disabled=false}});
</script></body></html>
"""


def clean_header(s: str) -> str:
    return (s or "").replace("\ufeff", "").strip()


def header_key(s: str) -> str:
    s = clean_header(s).strip('"').lower().translate(HEADER_FOLD)
    return re.sub(r"\s+", " ", s).strip()


REQUIRED_BY_KEY = {header_key(h): h for h in REQUIRED_HEADERS}


def decode_csv(data: bytes) -> str:
    # UTF-8 first; CP1251 fallback for older exports.
    for enc in ("utf-8-sig", "cp1251"):
        try:
            text = data.decode(enc)
            if "\x00" in text:
                raise ValueError("CSV faylida NUL belgisi bor. Fayl matnli CSV sifatida qayta saqlansin.")
            return text
        except UnicodeDecodeError:
            pass
    raise ValueError("CSV kodirovkasi o'qilmadi. UTF-8 yoki Windows-1251 CSV kerak.")


def detect_delimiter(text: str) -> str:
    # Sarlavha qatorida ustun nomlari ichida ajratuvchi uchramaydi,
    # shuning uchun eng ko'p uchragan belgi ajratuvchi hisoblanadi.
    header_line = next((ln for ln in text.splitlines() if ln.strip()), "")
    counts = {d: header_line.count(d) for d in (";", ",", "\t")}
    best = max(counts, key=counts.get)
    # Original files use semicolon.
    return best if counts[best] > 0 else ";"


@lru_cache(maxsize=65536)
def _parse_date_text(v: str) -> datetime | None:
    # Excel qayta saqlaganda vaqt qo'shiladi: "17.08.2026 0:00", "2026-08-17T00:00:00".
    date_part = re.split(r"[ T]", v, maxsplit=1)[0]
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(date_part, fmt)
        except ValueError:
            continue
    return None


def parse_date(value: str, file_name: str, row_no: int, column: str) -> datetime:
    v = (value or "").strip()
    # Sanalar ko'p takrorlanadi — kesh katta fayllarda tahlilni bir necha barobar tezlashtiradi.
    parsed = _parse_date_text(v)
    if parsed is not None:
        return parsed
    raise ValueError(
        f"{file_name}: {row_no}-qator, '{column}' sanasi noto'g'ri: {v!r}. "
        "Qabul qilinadi: YYYY-MM-DD, DD.MM.YYYY, DD/MM/YYYY yoki DD-MM-YYYY."
    )


def interval_for(days: int) -> str:
    if days < 0:
        raise ValueError("Manfiy kun farqi aniqlandi (bugungi sana so'nggi realizatsiyadan oldin).")
    if days <= 30:
        return "0-30 (НАЗОРАТ)"
    if days <= 35:
        return "31-35"
    if days <= 40:
        return "36-40"
    if days <= 45:
        return "41-45"
    if days <= 50:
        return "46-50"
    if days <= 55:
        return "51-55"
    if days <= 60:
        return "56-60"
    if days <= 65:
        return "61-65"
    if days <= 70:
        return "66-70"
    return "71+"


def sha256_hex(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


class MemoryUpload:
    """FileStorage-compatible wrapper for one supported member read from a ZIP."""

    def __init__(self, filename: str, data: bytes):
        self.filename = filename
        self._data = data

    def read(self) -> bytes:
        return self._data


def _xlsx_column_number(address: str) -> int:
    match = re.match(r"[A-Z]+", address or "")
    if not match:
        return 0
    number = 0
    for char in match.group():
        number = number * 26 + ord(char) - ord("A") + 1
    return number


def _xlsx_cell_text(cell: ET.Element, strings: List[str]) -> str:
    value = cell.find("x:v", XLSX_NS)
    cell_type = cell.get("t")
    if cell_type == "s" and value is not None:
        try:
            return strings[int(value.text or "0")]
        except (ValueError, IndexError) as exc:
            raise ValueError("Excel matnlar jadvali buzilgan.") from exc
    if cell_type == "inlineStr":
        inline = cell.find("x:is", XLSX_NS)
        return "".join(inline.itertext()) if inline is not None else ""
    return value.text or "" if value is not None else ""


def _xlsx_date_value(value: str, date_1904: bool) -> str:
    text = str(value or "").strip()
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        return text
    serial = float(text)
    if not 1 <= serial <= 100000:
        return text
    base = datetime(1904, 1, 1) if date_1904 else datetime(1899, 12, 30)
    return (base + timedelta(days=serial)).strftime("%Y-%m-%d")


def xlsx_to_csv_text(data: bytes, filename: str) -> str:
    """Read the first worksheet containing all source columns and return canonical CSV text."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            names = archive.namelist()
            sheet_names = sorted(
                (name for name in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)),
                key=lambda name: int(re.search(r"sheet(\d+)\.xml", name).group(1)),
            )
            if not sheet_names:
                raise ValueError(f"{filename}: Excel faylida varaq topilmadi.")

            strings = []
            if "xl/sharedStrings.xml" in names:
                info = archive.getinfo("xl/sharedStrings.xml")
                if info.file_size > 100 * 1024 * 1024:
                    raise ValueError(f"{filename}: Excel matnlar qismi juda katta.")
                shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                strings = ["".join(item.itertext()) for item in shared.findall("x:si", XLSX_NS)]

            date_1904 = False
            if "xl/workbook.xml" in names:
                workbook = ET.fromstring(archive.read("xl/workbook.xml"))
                properties = workbook.find("x:workbookPr", XLSX_NS)
                date_1904 = properties is not None and properties.get("date1904") in ("1", "true")

            for sheet_name in sheet_names:
                info = archive.getinfo(sheet_name)
                if info.file_size > 400 * 1024 * 1024:
                    raise ValueError(f"{filename}: Excel varag‘i juda katta.")
                root = ET.fromstring(archive.read(sheet_name))
                rows = []
                for row_node in root.findall("x:sheetData/x:row", XLSX_NS):
                    values = {}
                    for cell in row_node.findall("x:c", XLSX_NS):
                        column = _xlsx_column_number(cell.get("r", ""))
                        if column:
                            values[column] = _xlsx_cell_text(cell, strings)
                    rows.append((int(row_node.get("r", len(rows) + 1)), values))

                header_index = None
                column_by_header = {}
                for index, (_, values) in enumerate(rows):
                    found = defaultdict(list)
                    for column, value in values.items():
                        canonical = REQUIRED_BY_KEY.get(header_key(value))
                        if canonical:
                            found[canonical].append(column)
                    if all(header in found for header in REQUIRED_HEADERS):
                        repeated = [header for header, columns in found.items() if len(columns) > 1]
                        if repeated:
                            raise ValueError(
                                f"{filename}: Excel sarlavhasida ustun takrorlangan: {', '.join(repeated)}"
                            )
                        header_index = index
                        column_by_header = {header: found[header][0] for header in REQUIRED_HEADERS}
                        break

                if header_index is None:
                    continue

                output = io.StringIO()
                writer = csv.writer(output, delimiter=";", lineterminator="\n")
                writer.writerow(REQUIRED_HEADERS)
                date_headers = {"Сўнги реализация", "Бугунги реализация"}
                for _, values in rows[header_index + 1:]:
                    record = []
                    for header in REQUIRED_HEADERS:
                        value = values.get(column_by_header[header], "")
                        if header in date_headers:
                            value = _xlsx_date_value(value, date_1904)
                        record.append(value)
                    if any(str(value).strip() for value in record):
                        writer.writerow(record)
                return output.getvalue()
    except ValueError:
        raise
    except (zipfile.BadZipFile, KeyError, ET.ParseError, OSError) as exc:
        raise ValueError(f"{filename}: haqiqiy va buzilmagan .xlsx fayl kerak.") from exc

    raise ValueError(
        f"{filename}: Excel varaqlarida kerakli ustunlar topilmadi: {', '.join(REQUIRED_HEADERS)}"
    )


GAS_SALES_HEADERS = ["Учреждение", "Инспектор", "Принял", "Реализовал", "Вернул"]
GAS_SALES_HEADER_ALIASES = {
    header_key("Учреждение"): "Учреждение",
    header_key("Ташкилот"): "Учреждение",
    header_key("Райгаз"): "Учреждение",
    header_key("Инспектор"): "Инспектор",
    header_key("Чилангар"): "Инспектор",
    header_key("Ходим"): "Инспектор",
    header_key("Принял"): "Принял",
    header_key("Қабул қилинди"): "Принял",
    header_key("Кабул килинди"): "Принял",
    header_key("Реализовал"): "Реализовал",
    header_key("Сотилди"): "Реализовал",
    header_key("Вернул"): "Вернул",
    header_key("Қайтарилди"): "Вернул",
    header_key("Сотилмаган газ"): "Вернул",
}


def _gas_sales_integer(value: str, filename: str, row_no: int, column: str) -> int:
    text = str(value or "").strip().replace("\u00a0", "").replace(" ", "").replace(",", ".")
    # E-GAZ Excel eksportlarida miqdor katagida qo'shtirnoq belgisi ba'zan
    # bo'sh qiymat o'rnida keladi. Uni nol deb olamiz; parser bu holatni
    # hisobotdagi ogohlantirishlarda ham ko'rsatadi.
    if text in {'"', "“", "”"}:
        return 0
    try:
        number = float(text)
    except ValueError as exc:
        raise ValueError(
            f"{filename}: {row_no}-qator, '{column}' qiymati son emas: {value!r}."
        ) from exc
    if not math.isfinite(number):
        raise ValueError(
            f"{filename}: {row_no}-qator, '{column}' chekli son bo‘lishi kerak: {value!r}."
        )
    rounded = round(number)
    if number < 0 or abs(number - rounded) > 1e-8:
        raise ValueError(
            f"{filename}: {row_no}-qator, '{column}' musbat butun son bo‘lishi kerak: {value!r}."
        )
    return int(rounded)


def parse_gas_sales_xlsx(data: bytes, filename: str) -> dict | None:
    """Parse inspector-level accepted/sold/returned workbooks; return None for other XLSX layouts."""
    recognized_layout = False
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            names = archive.namelist()
            sheet_names = sorted(
                (name for name in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)),
                key=lambda name: int(re.search(r"sheet(\d+)\.xml", name).group(1)),
            )
            if not sheet_names:
                raise ValueError(f"{filename}: Excel faylida varaq topilmadi.")

            strings = []
            if "xl/sharedStrings.xml" in names:
                info = archive.getinfo("xl/sharedStrings.xml")
                if info.file_size > 100 * 1024 * 1024:
                    raise ValueError(f"{filename}: Excel matnlar qismi juda katta.")
                shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                strings = ["".join(item.itertext()) for item in shared.findall("x:si", XLSX_NS)]

            for sheet_name in sheet_names:
                info = archive.getinfo(sheet_name)
                if info.file_size > 400 * 1024 * 1024:
                    raise ValueError(f"{filename}: Excel varag‘i juda katta.")
                root = ET.fromstring(archive.read(sheet_name))
                rows = []
                for row_node in root.findall("x:sheetData/x:row", XLSX_NS):
                    values = {}
                    for cell in row_node.findall("x:c", XLSX_NS):
                        column = _xlsx_column_number(cell.get("r", ""))
                        if column:
                            values[column] = _xlsx_cell_text(cell, strings)
                    rows.append((int(row_node.get("r", len(rows) + 1)), values))

                header_index = None
                columns = {}
                for index, (_, values) in enumerate(rows):
                    found = defaultdict(list)
                    for column, value in values.items():
                        canonical = GAS_SALES_HEADER_ALIASES.get(header_key(value))
                        if canonical:
                            found[canonical].append(column)
                    if all(header in found for header in GAS_SALES_HEADERS):
                        repeated = [header for header in GAS_SALES_HEADERS if len(found[header]) > 1]
                        if repeated:
                            raise ValueError(
                                f"{filename}: Excel sarlavhasida ustun takrorlangan: {', '.join(repeated)}"
                            )
                        header_index = index
                        columns = {header: found[header][0] for header in GAS_SALES_HEADERS}
                        break
                if header_index is None:
                    continue
                recognized_layout = True

                period_candidates = []
                for _, values in rows[:header_index]:
                    for value in values.values():
                        text = clean_header(str(value))
                        if text and re.search(r"(?:ҳ|х)олатига|holatiga", text, re.IGNORECASE):
                            period_candidates.append(text)
                period = period_candidates[-1] if period_candidates else Path(filename).stem

                details = []
                district_total_entries = defaultdict(list)
                district_totals = {}
                stated_total = None
                warnings = []
                for row_no, values in rows[header_index + 1:]:
                    organization = clean_header(values.get(columns["Учреждение"], ""))
                    inspector = clean_header(values.get(columns["Инспектор"], ""))
                    number_values = [values.get(columns[name], "") for name in GAS_SALES_HEADERS[2:]]
                    organization_key = header_key(organization).strip("-— ")
                    inspector_key = header_key(inspector).strip("-— ")
                    if not organization and not inspector and not any(str(value).strip() for value in number_values):
                        continue
                    if not any(str(value).strip() for value in number_values):
                        continue

                    quoted_columns = [
                        column_name
                        for column_name, value in zip(GAS_SALES_HEADERS[2:], number_values)
                        if str(value or "").strip() in {'"', "“", "”"}
                    ]
                    if quoted_columns:
                        warnings.append(
                            f"{filename}: {row_no}-qator miqdor ustunidagi qo'shtirnoq belgisi "
                            f"0 deb olindi ({len(quoted_columns)} ta katak)."
                        )

                    # Some Uzbek exports end with an overall "Umumiy jami" row
                    # that only fills one total column (for example, returned
                    # gas) and leaves the other quantity cells blank.
                    if not inspector and organization_key in {
                        "жами", "умумий жами", "итого", "всего"
                    }:
                        total = {}
                        for field, value, column_name in zip(
                            ("accepted", "sold", "returned"), number_values,
                            GAS_SALES_HEADERS[2:],
                        ):
                            if str(value or "").strip():
                                total[field] = _gas_sales_integer(
                                    value, filename, row_no, column_name
                                )
                        if "accepted" in total and "sold" in total:
                            if total["sold"] > total["accepted"]:
                                raise ValueError(
                                    f"{filename}: {row_no}-qator balans mos emas: "
                                    f"Реализовал ({total['sold']}) Принял "
                                    f"({total['accepted']}) dan katta."
                                )
                            # Preserve the established rule that returned gas is
                            # checked as accepted minus sold when both exist.
                            total["returned"] = total["accepted"] - total["sold"]
                        if total:
                            stated_total = total
                        continue

                    accepted = _gas_sales_integer(number_values[0], filename, row_no, "Принял")
                    sold = _gas_sales_integer(number_values[1], filename, row_no, "Реализовал")
                    if sold > accepted:
                        raise ValueError(
                            f"{filename}: {row_no}-qator balans mos emas: "
                            f"Реализовал ({sold}) Принял ({accepted}) dan katta."
                        )

                    # E-GAZ eksportidagi “Вернул” ustuni sotilmagan gaz qoldig‘ini
                    # har doim bermaydi (masalan 3063 − 2374 o‘rniga 0 kelishi mumkin).
                    # Shablondagi sariq “sotilmagan gaz” esa qat’iy ravishda qabul
                    # qilingan miqdordan sotilgan miqdorni ayirish orqali hisoblanadi.
                    returned = accepted - sold
                    record_total = {"accepted": accepted, "sold": sold, "returned": returned}

                    if re.search(r"(?:^|\s)(жами|итого|всего)(?:\s|$)", inspector_key):
                        district_total_entries[header_key(organization)].append({
                            "name": organization,
                            "source_row": row_no,
                            **record_total,
                        })
                        continue
                    if not organization or not inspector:
                        warnings.append(f"{filename}: {row_no}-qator tashkilot yoki inspektorsiz qoldirildi.")
                        continue
                    details.append({
                        "organization": organization,
                        "inspector": inspector,
                        "accepted": accepted,
                        "sold": sold,
                        "returned": returned,
                        "source_file": filename,
                        "source_row": row_no,
                        "is_unassigned": False,
                    })

                # Ba’zi hisobotlarda bir raygaz uchun oraliq JAMI va yakuniy JAMI
                # qatorlari birga keladi. Oraliq qatordagi miqdor yo‘qolmasligi uchun
                # uni inspektorga biriktirilmagan alohida yozuv sifatida saqlaymiz.
                for district_key, entries in district_total_entries.items():
                    district_totals[district_key] = entries[-1]
                    for fragment in entries[:-1]:
                        details.append({
                            "organization": fragment["name"],
                            "inspector": "Тақсимланмаган (манбада оралиқ ЖАМИ)",
                            "accepted": fragment["accepted"],
                            "sold": fragment["sold"],
                            "returned": fragment["returned"],
                            "source_file": filename,
                            "source_row": fragment["source_row"],
                            "is_unassigned": True,
                        })

                if not details:
                    continue
                return {
                    "filename": filename,
                    "period": period,
                    "rows": details,
                    "district_totals": district_totals,
                    "stated_total": stated_total,
                    "warnings": warnings,
                    "sha256": sha256_hex(data),
                }
    except ValueError:
        raise
    except (zipfile.BadZipFile, KeyError, ET.ParseError, OSError) as exc:
        raise ValueError(f"{filename}: haqiqiy va buzilmagan .xlsx fayl kerak.") from exc
    if recognized_layout:
        raise ValueError(
            f"{filename}: qabul/sotuv sarlavhalari topildi, lekin inspektor ma’lumotlari yo‘q."
        )
    return None


def _analysis_order_number(value: str) -> str:
    text = clean_header(str(value or "")).strip("'\" ").replace("\u00a0", "")
    text = re.sub(r"\s+", "", text)
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    if text.isdigit():
        return str(int(text))
    return text


def _analysis_group_key(value: str) -> str:
    return re.sub(r"\s+", " ", clean_header(value)).strip().casefold()


def _analysis_decimal(value: str, filename: str, row_no: int, label: str) -> Decimal:
    text = str(value or "").strip().replace("\u00a0", "").replace(" ", "")
    if not text:
        raise ValueError(f"{filename}: {row_no}-qator, '{label}' bo'sh.")
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        tail = text.rsplit(",", 1)[-1]
        text = text.replace(",", ".") if len(tail) in (1, 2) else text.replace(",", "")
    try:
        number = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"{filename}: {row_no}-qator, '{label}' son emas: {value!r}.") from exc
    if not number.is_finite() or number < 0:
        raise ValueError(f"{filename}: {row_no}-qator, '{label}' musbat chekli qiymat bo'lishi kerak.")
    return number


def _analysis_date(value: str, date_1904: bool = False) -> str:
    text = clean_header(str(value or ""))
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        text = _xlsx_date_value(text, date_1904)
    for date_format in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y", "%Y.%m.%d"):
        try:
            return datetime.strptime(text[:10], date_format).date().isoformat()
        except ValueError:
            continue
    raise ValueError(f"Sana formati o'qilmadi: {value!r}.")


def parse_egaz_sales_csv(data: bytes, filename: str, require_amount: bool = True) -> list:
    """Read E-GAZ sale rows used by the MFY summary and GNP reconciliation pages."""
    text = decode_csv(data)
    portal_message = re.sub(r"\s+", " ", text.strip())
    portal_message = portal_message.translate(str.maketrans({"’": "'", "‘": "'", "ʼ": "'", "ʻ": "'"})).casefold()
    if re.fullmatch(r"ma'lumot(?:lar)? topilmadi[.!]?", portal_message):
        raise ValueError(
            f"{filename}: faylda CSV sotuvlari yo‘q — E-GAZ portali «Ma'lumot topilmadi» xabarini bergan. "
            "E-GAZda sana/tuman filtrlarini tekshirib, natijali hisobotni qayta eksport qiling."
        )
    delimiter = detect_delimiter(text)
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)
    aliases = {
        header_key("Дата"): "date",
        header_key("Sana"): "date",
        header_key("Учреждение"): "district",
        header_key("Raygaz"): "district",
        header_key("Заявка №"): "order_no",
        header_key("Ariza №"): "order_no",
        header_key("Махалля"): "mahalla",
        header_key("Mahalla"): "mahalla",
        header_key("Сумма"): "amount",
        header_key("Summa"): "amount",
    }
    field_by_name = {}
    for name in reader.fieldnames or []:
        canonical = aliases.get(header_key(name))
        if canonical:
            if canonical in field_by_name:
                raise ValueError(f"{filename}: '{name}' sarlavhasi takrorlangan.")
            field_by_name[canonical] = name
    required = {"date", "district", "order_no", "mahalla"}
    if require_amount:
        required.add("amount")
    missing = sorted(required - field_by_name.keys())
    if missing and not require_amount:
        descriptions = {
            "date": "Sana", "district": "Tuman/Raygaz", "order_no": "Ariza raqami", "mahalla": "MFY",
        }
        missing_labels = ", ".join(descriptions[name] for name in missing)
        raise ValueError(
            f"{filename}: CSV sarlavhasida kerakli ustun topilmadi: {missing_labels}. "
            "Ustun nomlari faylning birinchi qatorida bo'lishi kerak."
        )
    if missing:
        raise ValueError(
            f"{filename}: E-GAZ CSV sarlavhalari topilmadi yoki noto‘g‘ri. Kerakli ustunlar: "
            "Sana (Дата), tuman/Raygaz (Учреждение), ariza raqami (Заявка №), "
            "MFY (Махалля) va summa (Сумма)."
        )

    records = []
    for row_no, source in enumerate(reader, start=2):
        if not any(str(value or "").strip() for value in source.values()):
            continue
        order_no = _analysis_order_number(source.get(field_by_name["order_no"], ""))
        district = clean_header(source.get(field_by_name["district"], ""))
        mahalla = clean_header(source.get(field_by_name["mahalla"], ""))
        if not order_no or not district or not mahalla:
            raise ValueError(f"{filename}: {row_no}-qator, ariza/tuman/MFY qiymati bo'sh.")
        try:
            date = _analysis_date(source.get(field_by_name["date"], ""))
        except ValueError as exc:
            raise ValueError(f"{filename}: {row_no}-qator, {exc}") from exc
        amount = (
            _analysis_decimal(source.get(field_by_name["amount"], ""), filename, row_no, "Сумма")
            if require_amount else Decimal(0)
        )
        records.append({
            "date": date,
            "district": district,
            "mahalla": mahalla,
            "order_no": order_no,
            "amount": amount,
            "source_row": row_no,
        })
    if not records:
        raise ValueError(f"{filename}: CSV ichida sotuv yozuvlari topilmadi.")
    return records


def parse_gnp_orders_xlsx(data: bytes, filename: str) -> list:
    """Read GNP order rows and preserve both ordered and transferred cylinder counts."""
    header_aliases = {}
    for alias in ("Ariza №", "Заявка №", "Заявка N"):
        header_aliases[header_key(alias)] = "order_no"
    for alias in ("Qabul qiluvchi (RAYGAZ)", "Qabul qiluvchi Raygaz", "Учреждение"):
        header_aliases[header_key(alias)] = "district"
    for alias in ("Summa", "Сумма"):
        header_aliases[header_key(alias)] = "amount"
    for alias in ("Ballonlar soni", "Баллонлар сони"):
        header_aliases[header_key(alias)] = "requested"
    for alias in ("O'tkazilgan ballonlar", "O‘tkazilgan ballonlar", "Oʻtkazilgan ballonlar"):
        header_aliases[header_key(alias)] = "transferred"
    for alias in ("Yaratilgan", "Создано"):
        header_aliases[header_key(alias)] = "created_date"
    for alias in ("Qabul sanasi", "Дата приема"):
        header_aliases[header_key(alias)] = "accepted_date"
    for alias in ("MFY", "Mahalla", "Mahalla nomi", "Махалля"):
        header_aliases[header_key(alias)] = "mahalla"

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            names = archive.namelist()
            sheet_names = sorted(
                (name for name in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)),
                key=lambda name: int(re.search(r"sheet(\d+)\.xml", name).group(1)),
            )
            if not sheet_names:
                raise ValueError(f"{filename}: Excel varag'i topilmadi.")
            strings = []
            if "xl/sharedStrings.xml" in names:
                strings = [
                    "".join(item.itertext())
                    for item in ET.fromstring(archive.read("xl/sharedStrings.xml")).findall("x:si", XLSX_NS)
                ]
            date_1904 = False
            if "xl/workbook.xml" in names:
                book = ET.fromstring(archive.read("xl/workbook.xml"))
                properties = book.find("x:workbookPr", XLSX_NS)
                date_1904 = properties is not None and properties.get("date1904") in ("1", "true")

            for sheet_name in sheet_names:
                root = ET.fromstring(archive.read(sheet_name))
                rows = []
                for row_node in root.findall("x:sheetData/x:row", XLSX_NS):
                    values = {}
                    for cell in row_node.findall("x:c", XLSX_NS):
                        column = _xlsx_column_number(cell.get("r", ""))
                        if column:
                            values[column] = _xlsx_cell_text(cell, strings)
                    rows.append((int(row_node.get("r", len(rows) + 1)), values))

                header_index = None
                columns = {}
                for index, (_, values) in enumerate(rows):
                    found = defaultdict(list)
                    for column, value in values.items():
                        canonical = header_aliases.get(header_key(value))
                        if canonical:
                            found[canonical].append(column)
                    required = {"order_no", "district", "amount", "requested", "transferred"}
                    if required.issubset(found) and ("created_date" in found or "accepted_date" in found):
                        repeated = [name for name in required if len(found[name]) > 1]
                        if len(found.get("mahalla", [])) > 1:
                            repeated.append("mahalla")
                        if repeated:
                            raise ValueError(f"{filename}: sarlavhada ustunlar takrorlangan: {', '.join(repeated)}")
                        header_index = index
                        columns = {name: found[name][0] for name in required}
                        columns["date"] = found.get("created_date", found.get("accepted_date"))[0]
                        if found.get("mahalla"):
                            columns["mahalla"] = found["mahalla"][0]
                        break
                if header_index is None:
                    continue

                orders = []
                seen = set()
                for row_no, values in rows[header_index + 1:]:
                    order_no = _analysis_order_number(values.get(columns["order_no"], ""))
                    if not order_no:
                        if any(str(value).strip() for value in values.values()):
                            continue
                        continue
                    district = clean_header(values.get(columns["district"], ""))
                    if not district:
                        raise ValueError(f"{filename}: {row_no}-qator, Qabul qiluvchi (RAYGAZ) bo'sh.")
                    try:
                        date = _analysis_date(values.get(columns["date"], ""), date_1904)
                    except ValueError as exc:
                        raise ValueError(f"{filename}: {row_no}-qator, {exc}") from exc
                    amount = _analysis_decimal(values.get(columns["amount"], ""), filename, row_no, "Summa")
                    requested = _gas_sales_integer(values.get(columns["requested"], ""), filename, row_no, "Ballonlar soni")
                    transferred = _gas_sales_integer(values.get(columns["transferred"], ""), filename, row_no, "O'tkazilgan ballonlar")
                    if transferred > requested:
                        raise ValueError(f"{filename}: {row_no}-qator, o'tkazilgan ballonlar buyurtmadagi sondan ko'p.")
                    duplicate_key = (order_no, date)
                    if duplicate_key in seen:
                        raise ValueError(f"{filename}: {row_no}-qator, {order_no} arizasi shu sanada takrorlangan.")
                    seen.add(duplicate_key)
                    orders.append({
                        "order_no": order_no,
                        "district": district,
                        "mahalla": clean_header(values.get(columns.get("mahalla", 0), "")),
                        "date": date,
                        "amount": amount,
                        "requested": requested,
                        "transferred": transferred,
                        "unit_price": amount / requested if requested else Decimal(0),
                        "source_row": row_no,
                    })
                if orders:
                    return orders
    except ValueError:
        raise
    except (zipfile.BadZipFile, KeyError, ET.ParseError, OSError) as exc:
        raise ValueError(f"{filename}: buzilmagan .xlsx fayl kerak.") from exc
    raise ValueError(
        f"{filename}: GNP buyurtma ustunlari topilmadi: Ariza №, Qabul qiluvchi, Summa, "
        "Ballonlar soni, O'tkazilgan ballonlar va Yaratilgan/Qabul sanasi."
    )


def aggregate_egaz_sales(records: list) -> dict:
    districts = {}
    mfy_groups = {}
    order_days = {}
    order_numbers = {}
    dates = set()
    total_amount = Decimal(0)
    for record in records:
        district_key = _analysis_group_key(record["district"])
        mahalla_key = _analysis_group_key(record["mahalla"])
        mfy_key = (district_key, mahalla_key)
        if mfy_key not in mfy_groups:
            mfy_groups[mfy_key] = {
                "district": record["district"], "mahalla": record["mahalla"],
                "sales": 0, "amount": Decimal(0), "applications": set(),
            }
        group = mfy_groups[mfy_key]
        group["sales"] += 1
        group["amount"] += record["amount"]
        group["applications"].add(record["order_no"])
        if district_key not in districts:
            districts[district_key] = {
                "district": record["district"], "sales": 0, "amount": Decimal(0),
                "applications": set(), "mahalla_keys": set(),
            }
        district = districts[district_key]
        district["sales"] += 1
        district["amount"] += record["amount"]
        district["applications"].add(record["order_no"])
        district["mahalla_keys"].add(mahalla_key)
        day_key = (record["order_no"], record["date"])
        if day_key not in order_days:
            order_days[day_key] = {
                "order_no": record["order_no"], "date": record["date"],
                "sales": 0, "amount": Decimal(0), "districts": set(), "mahallas": set(), "files": set(),
            }
        order_days[day_key]["sales"] += 1
        order_days[day_key]["amount"] += record["amount"]
        order_days[day_key]["districts"].add(record["district"])
        order_days[day_key]["mahallas"].add(record["mahalla"])
        if record.get("source_file"):
            order_days[day_key]["files"].add(record["source_file"])
        if record["order_no"] not in order_numbers:
            order_numbers[record["order_no"]] = {
                "sales": 0, "amount": Decimal(0), "dates": set(), "districts": set(),
                "mahallas": set(), "files": set(),
            }
        order_numbers[record["order_no"]]["sales"] += 1
        order_numbers[record["order_no"]]["amount"] += record["amount"]
        order_numbers[record["order_no"]]["dates"].add(record["date"])
        order_numbers[record["order_no"]]["districts"].add(record["district"])
        order_numbers[record["order_no"]]["mahallas"].add(record["mahalla"])
        if record.get("source_file"):
            order_numbers[record["order_no"]]["files"].add(record["source_file"])
        dates.add(record["date"])
        total_amount += record["amount"]

    mfy_rows = []
    for group in mfy_groups.values():
        mfy_rows.append({**group, "applications": len(group["applications"])})
    district_rows = []
    for district in districts.values():
        district_rows.append({
            "district": district["district"], "sales": district["sales"],
            "amount": district["amount"], "applications": len(district["applications"]),
            "mahalla_count": len(district["mahalla_keys"]),
        })
    return {
        "records": records, "dates": sorted(dates), "sales": len(records),
        "amount": total_amount, "district_rows": district_rows, "mfy_rows": mfy_rows,
        "order_days": order_days, "order_numbers": order_numbers,
    }


def compare_egaz_with_gnp(sales: dict, orders: list) -> dict:
    comparisons = []
    gnp_numbers = set()
    for order in orders:
        number = order["order_no"]
        gnp_numbers.add(number)
        actual = sales["order_days"].get((number, order["date"]), {
            "sales": 0, "amount": Decimal(0), "districts": set(), "mahallas": set(), "files": set(),
        })
        other_days = [
            item for (sale_no, sale_date), item in sales["order_days"].items()
            if sale_no == number and sale_date != order["date"]
        ]
        count_difference = actual["sales"] - order["transferred"]
        district_names = sorted(actual["districts"])
        mahalla_names = sorted(actual["mahallas"])
        expected_mahalla = clean_header(order.get("mahalla", ""))
        district_mismatch = bool(district_names) and any(
            _analysis_group_key(name) != _analysis_group_key(order["district"])
            for name in district_names
        )
        mahalla_mismatch = bool(expected_mahalla and mahalla_names) and any(
            _analysis_group_key(name) != _analysis_group_key(expected_mahalla)
            for name in mahalla_names
        )
        other_day_sales = sum(item["sales"] for item in other_days)
        if not actual["sales"] and other_days:
            status = "SANA MOS EMAS"
        elif not actual["sales"] and order["transferred"]:
            status = "E-GAZDA TOPILMADI"
        elif not actual["sales"]:
            status = "0 SOTUV - E-GAZDA YO'Q"
        elif district_mismatch:
            status = "TUMAN TAFOVUTI"
        elif mahalla_mismatch:
            status = "MFY TAFOVUTI"
        elif count_difference != 0:
            status = "SONI TAFOVUTI"
        elif other_day_sales:
            status = "BOSHQA KUNDA HAM BOR"
        else:
            status = "MOS"
        comparisons.append({
            **order,
            "csv_sales": actual["sales"],
            "count_difference": count_difference,
            "csv_districts": district_names,
            "csv_mahallas": mahalla_names,
            "csv_files": sorted(actual["files"]),
            "district_mismatch": district_mismatch,
            "district_match": "MOS EMAS" if district_mismatch else "MOS" if district_names else "CSVda sotuv yo'q",
            "mahalla_mismatch": mahalla_mismatch,
            "mahalla_match": (
                "MOS EMAS" if mahalla_mismatch else "MOS" if expected_mahalla and mahalla_names
                else "GNPda MFY yo'q - E-GAZ MFYlari ko'rsatildi" if not expected_mahalla
                else "CSVda sotuv yo'q"
            ),
            "other_day_sales": other_day_sales,
            "other_day_files": sorted({name for item in other_days for name in item.get("files", set())}),
            "status": status,
        })
    csv_only = [
        {"order_no": number, **summary}
        for number, summary in sales["order_numbers"].items()
        if number not in gnp_numbers
    ]
    return {
        "orders": comparisons,
        "csv_only": csv_only,
        "matched": sum(row["status"] == "MOS" for row in comparisons),
        "mismatches": sum(row["status"] not in ("MOS", "0 SOTUV - E-GAZDA YO'Q") for row in comparisons),
        "zero_transfer_no_csv": sum(row["status"] == "0 SOTUV - E-GAZDA YO'Q" for row in comparisons),
        "same_day_found": sum(row["csv_sales"] > 0 for row in comparisons),
        "date_count": len(sales["dates"]),
    }


def ensure_unique_reports(reports: List[dict], label: str) -> None:
    """Block accidental double counting of byte-identical source workbooks."""
    seen = {}
    for report in reports:
        digest = report["sha256"]
        if digest in seen:
            raise ValueError(
                f"{label}: '{report['filename']}' fayli '{seen[digest]}' bilan aynan bir xil. "
                "Bir xil hisobotni ikki marta yuklash bloklandi."
            )
        seen[digest] = report["filename"]


def source_uploads(file_items, expanded_limit: int):
    """Yield direct CSV/XLSX files and supported ZIP members without disk extraction."""
    expanded_bytes = 0
    for item in file_items:
        upload_name = Path(item.filename or "noma'lum").name
        lower_name = upload_name.lower()
        if lower_name.endswith((".csv", ".xlsx")):
            yield item
            continue
        if not lower_name.endswith(".zip"):
            raise ValueError(f"{upload_name}: faqat .csv, .xlsx yoki .zip fayl qabul qilinadi.")

        try:
            item.stream.seek(0)
            with zipfile.ZipFile(item.stream) as archive:
                members = [info for info in archive.infolist()
                           if not info.is_dir()
                           and info.filename.lower().endswith((".csv", ".xlsx"))]
                if not members:
                    raise ValueError(f"{upload_name}: ZIP ichida CSV yoki XLSX fayl topilmadi.")
                if len(members) > 5000:
                    raise ValueError(f"{upload_name}: ZIP ichida 5000 tadan ko'p fayl bor.")
                for info in members:
                    if info.flag_bits & 0x1:
                        raise ValueError(f"{upload_name}: parolli ZIP qabul qilinmaydi.")
                    expanded_bytes += info.file_size
                    if expanded_bytes > expanded_limit:
                        limit_mb = expanded_limit // (1024 * 1024)
                        raise ValueError(
                            f"ZIP ichidagi CSV fayllar ochilganda {limit_mb} MB limitdan oshdi."
                        )
                    member_name = re.split(r"[\\/]", info.filename)[-1]
                    safe_name = f"{Path(upload_name).stem}__{member_name}"
                    yield MemoryUpload(safe_name, archive.read(info))
        except (zipfile.BadZipFile, zipfile.LargeZipFile, OSError) as exc:
            raise ValueError(f"{upload_name}: haqiqiy va buzilmagan ZIP fayl kerak.") from exc


def parse_files(file_items) -> Tuple[List[dict], List[dict], List[str], List[dict]]:
    all_rows: List[dict] = []
    file_stats: List[dict] = []
    warnings: List[str] = []
    source_manifest: List[dict] = []
    today_values = set()
    content_hashes = {}

    source_seen = False
    for item in file_items:
        source_seen = True
        filename = Path(item.filename or "noma'lum.csv").name
        is_xlsx = filename.lower().endswith(".xlsx")
        if not filename.lower().endswith((".csv", ".xlsx")):
            raise ValueError(f"{filename}: faqat .csv yoki .xlsx fayl qabul qilinadi.")

        data = item.read()
        if not data:
            raise ValueError(f"{filename}: fayl bo'sh.")
        digest = sha256_hex(data)
        if digest in content_hashes:
            raise ValueError(
                f"{filename}: fayl mazmuni '{content_hashes[digest]}' bilan aynan bir xil. "
                "Bir xil faylni ikki marta yuklash bloklandi."
            )
        content_hashes[digest] = filename
        if is_xlsx:
            text = xlsx_to_csv_text(data, filename)
            delim = ";"
        else:
            text = decode_csv(data)
            delim = detect_delimiter(text)

        reader = csv.DictReader(io.StringIO(text), delimiter=delim)
        if not reader.fieldnames:
            raise ValueError(f"{filename}: sarlavha qatori topilmadi.")

        original_fieldnames = reader.fieldnames
        normalized = [REQUIRED_BY_KEY.get(header_key(x), clean_header(x)) for x in original_fieldnames]
        rename = dict(zip(original_fieldnames, normalized))
        missing = [h for h in REQUIRED_HEADERS if h not in normalized]
        if missing:
            raise ValueError(
                f"{filename}: kerakli ustun(lar) yo'q: {', '.join(missing)}. "
                f"Faylda topilgan ustunlar: {', '.join(clean_header(x) for x in original_fieldnames)}"
            )
        repeated = [h for h in REQUIRED_HEADERS if normalized.count(h) > 1]
        if repeated:
            raise ValueError(
                f"{filename}: ustun(lar) bir necha marta uchraydi: {', '.join(repeated)}"
            )

        local_codes = Counter()
        local_diffs = []
        local_blank = 0
        local_count = 0
        local_under31 = 0

        for raw in reader:
            # line_num aniq qator raqamini beradi (bo'sh qatorlar va ko'p qatorli qiymatlar hisobga olinadi).
            excel_row_no = reader.line_num
            overflow = raw.get(None)
            if overflow and any(str(value).strip() for value in overflow):
                raise ValueError(
                    f"{filename}: {excel_row_no}-qator ustunlari sarlavha soniga mos emas. "
                    "Ajratuvchi belgilarni va qo‘shtirnoqlarni tekshiring."
                )
            row = {rename.get(k, clean_header(k or "")): (v if v is not None else "") for k, v in raw.items()}
            # Completely empty lines are ignored, not data deletion.
            if all(str(v).strip() == "" for v in row.values()):
                continue

            local_count += 1
            blank_required = [h for h in REQUIRED_HEADERS if str(row.get(h, "")).strip() == ""]
            if blank_required:
                local_blank += len(blank_required)
                raise ValueError(
                    f"{filename}: {excel_row_no}-qator: bo'sh majburiy maydon: "
                    + ", ".join(blank_required)
                )

            d1 = parse_date(row["Сўнги реализация"], filename, excel_row_no, "Сўнги реализация")
            d2 = parse_date(row["Бугунги реализация"], filename, excel_row_no, "Бугунги реализация")
            days = (d2.date() - d1.date()).days
            if days < 0:
                raise ValueError(
                    f"{filename}: {excel_row_no}-qator: kun farqi manfiy ({days})."
                )
            interval = interval_for(days)
            if days <= 30:
                local_under31 += 1

            code = str(row["Абонент код"]).strip()
            local_codes[code] += 1
            local_diffs.append(days)
            today_values.add(d2.date())

            # Preserve the 8 official source columns exactly as text/date values,
            # then append only derived technical columns.
            all_rows.append({
                "Вилоят": str(row["Вилоят"]).strip(),
                "Райгаз": str(row["Райгаз"]).strip(),
                "Маҳалла": str(row["Маҳалла"]).strip(),
                "Абонент код": code,
                "Абонент": str(row["Абонент"]).strip(),
                "Еҳтиёж": str(row["Еҳтиёж"]).strip(),
                "Сўнги реализация": d1,
                "Бугунги реализация": d2,
                "Кун фарқи": days,
                "Оралиқ": interval,
                "Манба файл": filename,
                "_source_row": excel_row_no,
            })

        if local_count == 0:
            raise ValueError(f"{filename}: ma'lumot qatori topilmadi.")

        local_dups = sum(v - 1 for v in local_codes.values() if v > 1)
        if local_under31:
            warnings.append(
                f"{filename}: {local_under31} ta yozuv 0-30 kun oralig'ida. "
                "Ular 'НАЗОРАТ' guruhi sifatida saqlandi."
            )

        file_stats.append({
            "Файл": filename,
            "Қатор сони": local_count,
            "Амалий min": min(local_diffs),
            "Амалий max": max(local_diffs),
            "Бўш майдон": local_blank,
            "Файл ичи дубликат": local_dups,
            "SHA-256": digest,
            "Ҳолат": "OK",
        })
        source_manifest.append({"file": filename, "sha256": digest, "rows": local_count})
        logging.info("SOURCE | %s | rows=%s | sha256=%s", filename, local_count, digest)

    if not source_seen:
        raise ValueError("Kamida bitta CSV yoki CSV fayllari bor ZIP tanlang.")

    # A single reporting/snapshot date is required to prevent mixing two days.
    if len(today_values) != 1:
        vals = ", ".join(sorted(d.isoformat() for d in today_values))
        raise ValueError(
            "“Бугунги реализация” sanasi barcha fayllarda bir xil emas. "
            f"Aniqlangan sanalar: {vals}. Turli kun ma'lumotlarini bitta svodga aralashtirish bloklandi."
        )

    return all_rows, file_stats, warnings, source_manifest


def _unique_rows_by_code(rows: List[dict], snapshot_label: str) -> Dict[str, dict]:
    """Build a reliable comparison index and reject ambiguous duplicate codes."""
    grouped = defaultdict(list)
    for row in rows:
        grouped[row["Абонент код"]].append(row)
    duplicate_codes = sorted(code for code, values in grouped.items() if len(values) > 1)
    if duplicate_codes:
        example = ", ".join(duplicate_codes[:10])
        more = f" (yana {len(duplicate_codes) - 10} ta)" if len(duplicate_codes) > 10 else ""
        raise ValueError(
            f"{snapshot_label}: abonent kodi takrorlangan, aniq solishtirib bo‘lmaydi: "
            f"{example}{more}. Dublikatlarni manba faylda tekshiring."
        )
    return {code: values[0] for code, values in grouped.items()}


def _numeric_need(value: str) -> Decimal | None:
    text = str(value or "").strip().replace(" ", "").replace(",", ".")
    if not re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text):
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def _display_number(value: Decimal) -> str:
    if value == value.to_integral_value():
        return str(int(value))
    return format(value.normalize(), "f")


def compare_sales_rows(old_rows: List[dict], new_rows: List[dict]) -> dict:
    """Compare two complete daily snapshots using subscriber code as the stable key."""
    if not old_rows or not new_rows:
        raise ValueError("Kechagi va yangi holatda kamida bittadan yozuv bo‘lishi kerak.")

    old_date = old_rows[0]["Бугунги реализация"].date()
    new_date = new_rows[0]["Бугунги реализация"].date()
    if new_date <= old_date:
        raise ValueError(
            f"Yangi holat sanasi ({new_date.isoformat()}) kechagi holat sanasidan "
            f"({old_date.isoformat()}) keyin bo‘lishi kerak. Fayllar joyini tekshiring."
        )

    old_by_code = _unique_rows_by_code(old_rows, "Kechagi holat")
    new_by_code = _unique_rows_by_code(new_rows, "Yangi holat")
    shared_codes = old_by_code.keys() & new_by_code.keys()

    details = []
    controls = []
    summary_counts = Counter()
    summary_needs = defaultdict(lambda: Decimal("0"))
    summary_need_complete = defaultdict(lambda: True)

    for code in shared_codes:
        old_row = old_by_code[code]
        new_row = new_by_code[code]
        old_last = old_row["Сўнги реализация"].date()
        new_last = new_row["Сўнги реализация"].date()

        changed_fields = []
        for field, label in (
                ("Райгаз", "Raygaz"), ("Маҳалла", "mahalla"),
                ("Абонент", "abonent"), ("Еҳтиёж", "ehtiyoj")):
            if str(old_row[field]).strip() != str(new_row[field]).strip():
                changed_fields.append(label)
        if changed_fields:
            controls.append({
                "status": "Abonent ma’lumoti o‘zgargan: " + ", ".join(changed_fields),
                "code": code,
                "raygaz": new_row["Райгаз"],
                "mahalla": new_row["Маҳалла"],
                "subscriber": new_row["Абонент"],
            })

        # A later realization date inside the two-snapshot window is direct evidence of a sale.
        if new_last > old_last and new_last >= old_date:
            detail = {
                "raygaz": new_row["Райгаз"],
                "mahalla": new_row["Маҳалла"],
                "code": code,
                "subscriber": new_row["Абонент"],
                "need": new_row["Еҳтиёж"],
                "old_last_sale": old_last.isoformat(),
                "new_last_sale": new_last.isoformat(),
            }
            details.append(detail)
            raygaz = new_row["Райгаз"]
            summary_counts[raygaz] += 1
            need_number = _numeric_need(new_row["Еҳтиёж"])
            if need_number is None:
                summary_need_complete[raygaz] = False
            else:
                summary_needs[raygaz] += need_number
        elif new_last != old_last:
            status = ("Realizatsiya sanasi orqaga o‘zgargan"
                      if new_last < old_last else "Oldingi realizatsiya sanasi tuzatilgan")
            controls.append({
                "status": status,
                "code": code,
                "raygaz": new_row["Райгаз"],
                "mahalla": new_row["Маҳалла"],
                "subscriber": new_row["Абонент"],
            })

    missing_codes = old_by_code.keys() - new_by_code.keys()
    for code in missing_codes:
        row = old_by_code[code]
        controls.append({
            "status": "Yangi holatda topilmadi",
            "code": code,
            "raygaz": row["Райгаз"],
            "mahalla": row["Маҳалла"],
            "subscriber": row["Абонент"],
        })

    added_codes = new_by_code.keys() - old_by_code.keys()
    for code in added_codes:
        row = new_by_code[code]
        controls.append({
            "status": "Yangi qo‘shilgan kod",
            "code": code,
            "raygaz": row["Райгаз"],
            "mahalla": row["Маҳалла"],
            "subscriber": row["Абонент"],
        })

    sort_key = lambda item: (header_key(item["raygaz"]), header_key(item["mahalla"]), item["code"])
    details.sort(key=sort_key)
    controls.sort(key=lambda item: (item["status"], *sort_key(item)))
    summary = []
    for raygaz in sorted(summary_counts, key=lambda name: (-summary_counts[name], header_key(name))):
        need_total = (_display_number(summary_needs[raygaz])
                      if summary_need_complete[raygaz] else "Aralash qiymat")
        summary.append({
            "raygaz": raygaz,
            "sales_count": summary_counts[raygaz],
            "need_total": need_total,
        })

    logging.info(
        "SALES | old=%s | new=%s | sales=%s | missing=%s | added=%s",
        old_date, new_date, len(details), len(missing_codes), len(added_codes),
    )
    return {
        "old_date": old_date.isoformat(),
        "new_date": new_date.isoformat(),
        "total_sales": len(details),
        "raygaz_count": len(summary_counts),
        "missing_count": len(missing_codes),
        "added_count": len(added_codes),
        "summary": summary,
        "details": details,
        "controls": controls,
    }


def compare_gas_sales_reports(old_reports: List[dict], new_reports: List[dict]) -> dict:
    """Compare cumulative accepted/sold/returned summaries by district and inspector."""
    def combine(reports):
        combined = {}
        for report in reports:
            for row in report["rows"]:
                key = (
                    header_key(row["organization"]),
                    header_key(row["inspector"]),
                    bool(row.get("is_unassigned")),
                )
                item = combined.setdefault(key, {
                    "organization": re.sub(r"\s+", " ", row["organization"]).strip(),
                    "inspector": row["inspector"],
                    "accepted": 0,
                    "sold": 0,
                    "returned": 0,
                })
                for field in ("accepted", "sold", "returned"):
                    item[field] += row[field]
        return combined

    old_map = combine(old_reports)
    new_map = combine(new_reports)
    all_keys = sorted(
        set(old_map) | set(new_map),
        key=lambda key: (
            key[0],
            (new_map.get(key) or old_map.get(key))["organization"],
            key[1],
        ),
    )

    details = []
    controls = []
    for key in all_keys:
        old = old_map.get(key)
        new = new_map.get(key)
        shown = new or old
        values = {}
        for field in ("accepted", "sold", "returned"):
            old_value = old[field] if old else 0
            new_value = new[field] if new else 0
            values[f"old_{field}"] = old_value
            values[f"new_{field}"] = new_value
            values[f"{field}_delta"] = new_value - old_value

        if old is None:
            status = "Yangi holatda qo‘shilgan"
        elif new is None:
            status = "Yangi holatda yo‘q"
        elif any(values[f"{field}_delta"] < 0 for field in ("accepted", "sold", "returned")):
            status = "Miqdor kamaygan — tekshiring"
        else:
            status = "O‘zgargan"

        if old is None or new is None or status.startswith("Miqdor kamaygan"):
            controls.append({
                "status": status,
                "organization": shown["organization"],
                "inspector": shown["inspector"],
            })
        if old is None or new is None or any(values[f"{field}_delta"] for field in ("accepted", "sold", "returned")):
            details.append({
                "organization": shown["organization"],
                "inspector": shown["inspector"],
                **values,
            })

    district_names = {}
    old_districts = defaultdict(Counter)
    new_districts = defaultdict(Counter)
    for mapping, target in ((old_map, old_districts), (new_map, new_districts)):
        for key, row in mapping.items():
            district_key = key[0]
            district_names.setdefault(district_key, row["organization"])
            for field in ("accepted", "sold", "returned"):
                target[district_key][field] += row[field]

    district_keys = sorted(district_names, key=lambda key: district_names[key])
    summary = []
    for key in district_keys:
        old_values = old_districts[key]
        new_values = new_districts[key]
        summary.append({
            "raygaz": district_names[key],
            "old_sold": old_values["sold"],
            "new_sold": new_values["sold"],
            "sold_delta": new_values["sold"] - old_values["sold"],
            "old_returned": old_values["returned"],
            "new_returned": new_values["returned"],
            "returned_delta": new_values["returned"] - old_values["returned"],
            "old_accepted": old_values["accepted"],
            "new_accepted": new_values["accepted"],
            "accepted_delta": new_values["accepted"] - old_values["accepted"],
        })
    summary.sort(key=lambda item: (-abs(item["returned_delta"]), item["raygaz"]))
    details.sort(key=lambda item: (-abs(item["returned_delta"]), item["organization"], item["inspector"]))

    # Vebdagi bitta jadval va yuklanadigan Excel kechagi hisobot shaklini saqlaydi:
    # avval kechagi chilangarlar, keyin o‘sha raygazning ЖАМИ qatori va unda
    # yangi sotilmagan gaz − kechagi sotilmagan gaz farqi.
    old_source_order = {report["filename"]: index for index, report in enumerate(old_reports)}
    ordered_old_rows = sorted(
        (row for report in old_reports for row in report["rows"]),
        key=lambda row: (old_source_order.get(row["source_file"], 0), row["source_row"]),
    )
    old_district_order = []
    old_details_by_district = defaultdict(list)
    for row in ordered_old_rows:
        district_key = header_key(row["organization"])
        if district_key not in old_details_by_district:
            old_district_order.append(district_key)
        old_details_by_district[district_key].append(row)

    comparison_rows = []
    for key in old_district_order:
        district_name = district_names[key]
        for row in old_details_by_district[key]:
            comparison_rows.append({
                "organization": district_name,
                "inspector": row["inspector"],
                "accepted": row["accepted"],
                "sold": row["sold"],
                "returned": row["returned"],
                "percent": round(row["sold"] * 100 / row["accepted"]) if row["accepted"] else 0,
                "returned_delta": None,
                "is_total": False,
                "is_grand_total": False,
            })
        old_values = old_districts[key]
        new_values = new_districts[key]
        comparison_rows.append({
            "organization": district_name,
            "inspector": f"--- {district_name} ЖАМИ ---",
            "accepted": old_values["accepted"],
            "sold": old_values["sold"],
            "returned": old_values["returned"],
            "percent": round(old_values["sold"] * 100 / old_values["accepted"])
            if old_values["accepted"] else 0,
            "returned_delta": new_values["returned"] - old_values["returned"],
            "is_total": True,
            "is_grand_total": False,
        })

    for key in district_keys:
        if key in old_details_by_district:
            continue
        new_values = new_districts[key]
        district_name = district_names[key]
        comparison_rows.append({
            "organization": district_name,
            "inspector": "Кечаги ҳисоботда йўқ",
            "accepted": 0,
            "sold": 0,
            "returned": 0,
            "percent": 0,
            "returned_delta": new_values["returned"],
            "is_total": True,
            "is_grand_total": False,
        })

    def total(mapping, field):
        return sum(row[field] for row in mapping.values())

    old_accepted = total(old_map, "accepted")
    old_sold = total(old_map, "sold")
    old_returned = total(old_map, "returned")
    new_returned = total(new_map, "returned")
    comparison_rows.append({
        "organization": "Жами",
        "inspector": "",
        "accepted": old_accepted,
        "sold": old_sold,
        "returned": old_returned,
        "percent": round(old_sold * 100 / old_accepted) if old_accepted else 0,
        "returned_delta": new_returned - old_returned,
        "is_total": True,
        "is_grand_total": True,
    })

    old_period = "; ".join(dict.fromkeys(report["period"] for report in old_reports))
    new_period = "; ".join(dict.fromkeys(report["period"] for report in new_reports))
    return {
        "report_type": "gas-sales-summary",
        "old_period": old_period,
        "new_period": new_period,
        "accepted_delta": total(new_map, "accepted") - total(old_map, "accepted"),
        "sold_delta": total(new_map, "sold") - total(old_map, "sold"),
        "returned_delta": total(new_map, "returned") - total(old_map, "returned"),
        "changed_count": len(details),
        "raygaz_count": len(summary),
        "summary": summary,
        "comparison_rows": comparison_rows,
        "details": details,
        "controls": controls,
    }


def build_gas_sales_comparison_xlsx(old_reports: List[dict], new_reports: List[dict]) -> bytes:
    """Create the image-style comparison: district unsold-gas delta in column G on total rows."""
    old_rows = [row for report in old_reports for row in report["rows"]]
    new_rows = [row for report in new_reports for row in report["rows"]]
    if not old_rows or not new_rows:
        raise ValueError("Solishtirish uchun ikkala hisobotda ham ma’lumot bo‘lishi kerak.")

    def district_totals(rows):
        totals = defaultdict(Counter)
        names = {}
        for row in rows:
            key = header_key(row["organization"])
            names.setdefault(key, re.sub(r"\s+", " ", row["organization"]).strip())
            for field in ("accepted", "sold", "returned"):
                totals[key][field] += row[field]
        return totals, names

    old_totals, old_names = district_totals(old_rows)
    new_totals, new_names = district_totals(new_rows)
    source_order = {report["filename"]: index for index, report in enumerate(old_reports)}
    ordered_rows = sorted(old_rows, key=lambda row: (
        source_order.get(row["source_file"], 0), row["source_row"]
    ))
    district_order = []
    district_details = defaultdict(list)
    for row in ordered_rows:
        key = header_key(row["organization"])
        if key not in district_details:
            district_order.append(key)
        district_details[key].append(row)

    period = "; ".join(dict.fromkeys(report["period"] for report in new_reports))
    old_period = "; ".join(dict.fromkeys(report["period"] for report in old_reports))
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {
        "in_memory": True,
        "strings_to_formulas": False,
        "strings_to_urls": False,
    })
    wb.set_properties({
        "title": f"{old_period} va {period} qabul-sotuv solishtirishi",
        "subject": "Raygaz JAMI qatorlarida Вернул (sotilmagan gaz) farqi",
        "author": APP_NAME,
    })

    fmt_title = wb.add_format({"bold": True, "font_size": 16, "align": "center",
                               "valign": "vcenter", "text_wrap": True})
    fmt_info = wb.add_format({"bold": True, "font_size": 14, "align": "center",
                              "valign": "vcenter"})
    fmt_period = wb.add_format({"bold": True, "bg_color": "#19B5E6", "border": 1,
                                "align": "center", "valign": "vcenter"})
    fmt_group = wb.add_format({"bold": True, "border": 1, "align": "center",
                               "valign": "vcenter", "text_wrap": True})
    fmt_group_yellow = wb.add_format({"bold": True, "bg_color": "#FFF200", "border": 1,
                                      "align": "center", "valign": "vcenter", "text_wrap": True})
    fmt_compare_header = wb.add_format({"bold": True, "bg_color": "#F4B183", "border": 1,
                                        "align": "center", "valign": "vcenter", "text_wrap": True})
    fmt_header = wb.add_format({"bold": True, "border": 1, "align": "center",
                                "valign": "vcenter", "text_wrap": True})
    fmt_header_yellow = wb.add_format({"bold": True, "bg_color": "#FFF200", "border": 1,
                                       "align": "center", "valign": "vcenter"})
    fmt_cell = wb.add_format({"border": 1, "valign": "vcenter"})
    fmt_int = wb.add_format({"border": 1, "num_format": "#,##0", "align": "center"})
    fmt_pct = wb.add_format({"border": 1, "num_format": "0", "align": "center"})
    fmt_total_text = wb.add_format({"bold": True, "bg_color": "#FFF200", "border": 1,
                                    "align": "center", "valign": "vcenter"})
    fmt_total_int = wb.add_format({"bold": True, "bg_color": "#FFF200", "border": 1,
                                   "num_format": "#,##0", "align": "center"})
    fmt_delta_zero = wb.add_format({"bold": True, "bg_color": "#FFD966", "border": 1,
                                    "num_format": "+#,##0;-#,##0;0", "align": "center"})
    fmt_delta = wb.add_format({"bold": True, "bg_color": "#FFD966", "font_color": "#C00000",
                               "border": 1, "num_format": "+#,##0;-#,##0;0", "align": "center"})
    fmt_grand_text = wb.add_format({"bold": True, "bg_color": "#F4B183", "border": 1,
                                    "align": "center", "valign": "vcenter"})
    fmt_grand_int = wb.add_format({"bold": True, "bg_color": "#F4B183", "border": 1,
                                   "num_format": "#,##0", "align": "center"})
    fmt_grand_delta = wb.add_format({"bold": True, "bg_color": "#F4B183", "border": 1,
                                     "num_format": "+#,##0;-#,##0;0", "align": "center"})
    fmt_note = wb.add_format({"italic": True, "font_color": "#666666", "text_wrap": True})

    ws = wb.add_worksheet("Солиштириш")
    ws.merge_range(
        "A1:F2",
        '"Худудгаз Андижон" ГТФ тасарруфидаги шаҳар-тумангаз таъминоти '
        "суюлтирилган газ бўлими чилангар-таъминотчилари томонидан суюлтирилган "
        "газни E-газ дастурида қабул қилиш ва сотиш тўғрисида",
        fmt_title,
    )
    ws.merge_range("A3:F3", "М А Ъ Л У М О Т", fmt_info)
    # Shablon sarlavhasida fayl nomi emas, kechagi/eski hisobot davri turadi.
    # Masalan, yangi eksport "0000000.xlsx" deb nomlangan bo‘lsa ham ko‘k qatorda
    # "1-31 Август холатига" kabi eski hisobotdagi davr saqlanadi.
    ws.merge_range("C4:F4", old_period, fmt_period)
    ws.merge_range("A5:B5", "Шаҳар-туман номи ва Чилангарларни Ф.И.Ш", fmt_group)
    ws.write("C5", "E-газ дастурида қабул қилинган газ", fmt_group)
    ws.write("D5", "E-газ дастурида сотилган газ", fmt_group)
    ws.write("E5", "E-газ дастурида сотилмаган газ", fmt_group_yellow)
    ws.write("F5", "%", fmt_group)
    ws.write("G5", "Солиштириш\n(янги − кечаги сотилмаган газ)", fmt_compare_header)
    ws.write_row("A6", ["Учреждение", "Инспектор", "Принял", "Реализовал", "Вернул"], fmt_header)
    ws.write("F6", "%", fmt_header)
    ws.write("G6", "Фарқ", fmt_compare_header)
    ws.set_row(0, 33)
    ws.set_row(1, 33)
    ws.set_row(4, 82)
    ws.set_row(5, 24)

    row_index = 6
    for key in district_order:
        district_name = old_names[key]
        for row in district_details[key]:
            accepted = row["accepted"]
            sold = row["sold"]
            returned = row["returned"]
            ws.write(row_index, 0, district_name, fmt_cell)
            ws.write(row_index, 1, row["inspector"], fmt_cell)
            ws.write_number(row_index, 2, accepted, fmt_int)
            ws.write_number(row_index, 3, sold, fmt_int)
            ws.write_number(row_index, 4, returned, fmt_int)
            ws.write_number(row_index, 5, round(sold * 100 / accepted) if accepted else 0, fmt_pct)
            ws.write_blank(row_index, 6, None, fmt_cell)
            row_index += 1

        previous = old_totals[key]
        current = new_totals[key]
        delta = current["returned"] - previous["returned"]
        ws.write(row_index, 0, district_name, fmt_total_text)
        ws.write(row_index, 1, f"--- {district_name} ЖАМИ ---", fmt_total_text)
        ws.write_number(row_index, 2, previous["accepted"], fmt_total_int)
        ws.write_number(row_index, 3, previous["sold"], fmt_total_int)
        ws.write_number(row_index, 4, previous["returned"], fmt_total_int)
        ws.write_number(row_index, 5,
                        round(previous["sold"] * 100 / previous["accepted"])
                        if previous["accepted"] else 0,
                        fmt_total_int)
        ws.write_number(row_index, 6, delta, fmt_delta_zero if delta == 0 else fmt_delta)
        row_index += 1

    new_only_keys = [key for key in new_totals if key not in old_totals]
    if new_only_keys:
        row_index += 1
        ws.merge_range(row_index, 0, row_index, 6, "КЕЧАГИ ҲОЛАТДА ЙЎҚ РАЙГАЗЛАР", fmt_compare_header)
        row_index += 1
        for key in sorted(new_only_keys, key=lambda item: new_names[item]):
            delta = new_totals[key]["returned"]
            ws.write(row_index, 0, new_names[key], fmt_total_text)
            ws.write(row_index, 1, "Кечаги ҳисоботда йўқ", fmt_total_text)
            ws.write_blank(row_index, 2, None, fmt_total_int)
            ws.write_blank(row_index, 3, None, fmt_total_int)
            ws.write_blank(row_index, 4, None, fmt_total_int)
            ws.write_blank(row_index, 5, None, fmt_total_int)
            ws.write_number(row_index, 6, delta, fmt_delta)
            row_index += 1

    old_accepted = sum(item["accepted"] for item in old_totals.values())
    old_sold = sum(item["sold"] for item in old_totals.values())
    old_returned = sum(item["returned"] for item in old_totals.values())
    new_returned = sum(item["returned"] for item in new_totals.values())
    ws.merge_range(row_index, 0, row_index, 1, "Жами", fmt_grand_text)
    ws.write_number(row_index, 2, old_accepted, fmt_grand_int)
    ws.write_number(row_index, 3, old_sold, fmt_grand_int)
    ws.write_number(row_index, 4, old_returned, fmt_grand_int)
    ws.write_number(row_index, 5,
                    round(old_sold * 100 / old_accepted) if old_accepted else 0,
                    fmt_grand_int)
    ws.write_number(row_index, 6, new_returned - old_returned, fmt_grand_delta)
    row_index += 1

    ws.merge_range(row_index + 1, 0, row_index + 1, 6,
                   f"Изоҳ: G устун = янги “Вернул” − кечаги “Вернул” (сотилмаган газ). Кечаги давр: {old_period}.",
                   fmt_note)

    ws.set_column("A:A", 32)
    ws.set_column("B:B", 49)
    ws.set_column("C:E", 14)
    ws.set_column("F:F", 8)
    ws.set_column("G:G", 18)
    ws.freeze_panes(6, 2)
    ws.autofilter(5, 0, max(6, row_index - 1), 6)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(4, 5)
    wb.close()
    return output.getvalue()


def _analysis_sheet_name(value: str, used: set) -> str:
    name = re.sub(r"[\\/*?:\[\]]", " ", clean_header(value)).strip() or "Tuman"
    name = name[:31]
    candidate = name
    suffix = 2
    while candidate.casefold() in used:
        marker = f" ({suffix})"
        candidate = name[:31 - len(marker)] + marker
        suffix += 1
    used.add(candidate.casefold())
    return candidate


def _analysis_formats(workbook):
    return {
        "title": workbook.add_format({"bold": True, "font_size": 16, "font_color": "#FFFFFF", "bg_color": "#102F45", "valign": "vcenter"}),
        "header": workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#174D66", "border": 1, "text_wrap": True, "valign": "vcenter"}),
        "text": workbook.add_format({"border": 1, "valign": "top"}),
        "note": workbook.add_format({"italic": True, "font_color": "#365B66", "bg_color": "#EDF7F8", "border": 1, "text_wrap": True, "valign": "vcenter"}),
        "integer": workbook.add_format({"border": 1, "num_format": "#,##0"}),
        "money": workbook.add_format({"border": 1, "num_format": "#,##0.00"}),
        "total_label": workbook.add_format({"bold": True, "bg_color": "#E9F8F2", "border": 1}),
        "total_integer": workbook.add_format({"bold": True, "bg_color": "#E9F8F2", "border": 1, "num_format": "#,##0"}),
        "total_money": workbook.add_format({"bold": True, "bg_color": "#E9F8F2", "border": 1, "num_format": "#,##0.00"}),
        "ok": workbook.add_format({"bold": True, "font_color": "#0D6549", "bg_color": "#E9F8F2", "border": 1}),
        "bad": workbook.add_format({"bold": True, "font_color": "#B42318", "bg_color": "#FFF1F0", "border": 1}),
    }


def _analysis_write_table_header(sheet, headers, formats, row=2):
    sheet.write_row(row, 0, headers, formats["header"])
    sheet.set_row(row, 32)
    sheet.freeze_panes(row + 1, 0)


def build_mfy_svod_xlsx(sales: dict, source_filename: str) -> Tuple[bytes, dict]:
    """Create an all-district MFY workbook plus a separate worksheet per district."""
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    formats = _analysis_formats(workbook)
    used_names = {"umumiy", "tumanlar", "mfylar"}

    sheet = workbook.add_worksheet("Umumiy")
    sheet.merge_range("A1:D1", "E-GAZ SOTUVLARI — UMUMIY SVOD", formats["title"])
    sheet.set_row(0, 30)
    metrics = [
        ("Manba fayl", source_filename),
        ("Hisobot sanasi", ", ".join(sales["dates"])),
        ("Sotuv satrlari / ballon", sales["sales"]),
        ("Jami summa, so'm", float(sales["amount"])),
        ("Tuman va shaharlar", len(sales["district_rows"])),
        ("Tuman–MFY juftliklari", len(sales["mfy_rows"])),
        ("Ariza raqamlari", len(sales["order_numbers"])),
    ]
    for index, (label, value) in enumerate(metrics, start=3):
        sheet.write(index, 0, label, formats["total_label"])
        fmt = formats["money"] if label == "Jami summa, so'm" else formats["integer"] if isinstance(value, int) else formats["text"]
        sheet.write(index, 1, value, fmt)
    sheet.set_column("A:A", 30)
    sheet.set_column("B:B", 54)
    sheet.set_column("C:D", 18)

    district_headers = ["Tuman / shahar", "Sotuv soni", "Arizalar", "MFY soni", "Jami summa, so'm"]
    district_rows = sorted(sales["district_rows"], key=lambda row: _analysis_group_key(row["district"]))
    sheet = workbook.add_worksheet("Tumanlar")
    sheet.merge_range(0, 0, 0, len(district_headers) - 1, "TUMAN VA SHAHARLAR BO'YICHA", formats["title"])
    sheet.set_row(0, 30)
    _analysis_write_table_header(sheet, district_headers, formats)
    for row_index, item in enumerate(district_rows, start=3):
        sheet.write(row_index, 0, item["district"], formats["text"])
        sheet.write_number(row_index, 1, item["sales"], formats["integer"])
        sheet.write_number(row_index, 2, item["applications"], formats["integer"])
        sheet.write_number(row_index, 3, item["mahalla_count"], formats["integer"])
        sheet.write_number(row_index, 4, float(item["amount"]), formats["money"])
    total_row = len(district_rows) + 3
    sheet.write(total_row, 0, "JAMI", formats["total_label"])
    sheet.write_number(total_row, 1, sales["sales"], formats["total_integer"])
    sheet.write_number(total_row, 2, len(sales["order_numbers"]), formats["total_integer"])
    sheet.write_number(total_row, 3, len(sales["mfy_rows"]), formats["total_integer"])
    sheet.write_number(total_row, 4, float(sales["amount"]), formats["total_money"])
    sheet.autofilter(2, 0, max(2, total_row - 1), len(district_headers) - 1)
    sheet.set_column("A:A", 32)
    sheet.set_column("B:D", 16)
    sheet.set_column("E:E", 22)

    mfy_headers = ["Tuman / shahar", "MFY", "Sotuv soni", "Arizalar", "Jami summa, so'm"]
    mfy_rows = sorted(sales["mfy_rows"], key=lambda row: (
        _analysis_group_key(row["district"]), _analysis_group_key(row["mahalla"])
    ))
    sheet = workbook.add_worksheet("MFYlar")
    sheet.merge_range(0, 0, 0, len(mfy_headers) - 1, "TUMAN VA MFY BO'YICHA SOTUVLAR", formats["title"])
    sheet.set_row(0, 30)
    _analysis_write_table_header(sheet, mfy_headers, formats)
    for row_index, item in enumerate(mfy_rows, start=3):
        sheet.write(row_index, 0, item["district"], formats["text"])
        sheet.write(row_index, 1, item["mahalla"], formats["text"])
        sheet.write_number(row_index, 2, item["sales"], formats["integer"])
        sheet.write_number(row_index, 3, item["applications"], formats["integer"])
        sheet.write_number(row_index, 4, float(item["amount"]), formats["money"])
    sheet.autofilter(2, 0, len(mfy_rows) + 2, len(mfy_headers) - 1)
    sheet.set_column("A:B", 30)
    sheet.set_column("C:D", 15)
    sheet.set_column("E:E", 22)

    mfy_by_district = defaultdict(list)
    for item in mfy_rows:
        mfy_by_district[_analysis_group_key(item["district"])].append(item)
    for district in district_rows:
        district_items = mfy_by_district[_analysis_group_key(district["district"])]
        sheet_name = _analysis_sheet_name(district["district"], used_names)
        sheet = workbook.add_worksheet(sheet_name)
        headers = ["MFY", "Sotuv soni", "Arizalar", "Jami summa, so'm"]
        sheet.merge_range(0, 0, 0, 3, f"{district['district']} — MFY SVOD", formats["title"])
        sheet.set_row(0, 30)
        _analysis_write_table_header(sheet, headers, formats)
        for row_index, item in enumerate(sorted(district_items, key=lambda row: -row["amount"]), start=3):
            sheet.write(row_index, 0, item["mahalla"], formats["text"])
            sheet.write_number(row_index, 1, item["sales"], formats["integer"])
            sheet.write_number(row_index, 2, item["applications"], formats["integer"])
            sheet.write_number(row_index, 3, float(item["amount"]), formats["money"])
        subtotal = len(district_items) + 3
        sheet.write(subtotal, 0, "TUMAN JAMI", formats["total_label"])
        sheet.write_number(subtotal, 1, district["sales"], formats["total_integer"])
        sheet.write_number(subtotal, 2, district["applications"], formats["total_integer"])
        sheet.write_number(subtotal, 3, float(district["amount"]), formats["total_money"])
        sheet.autofilter(2, 0, max(2, subtotal - 1), len(headers) - 1)
        sheet.set_column("A:A", 32)
        sheet.set_column("B:C", 16)
        sheet.set_column("D:D", 22)

    workbook.close()
    return output.getvalue(), {
        "rows": sales["sales"], "amount": sales["amount"],
        "districts": len(sales["district_rows"]), "mfy": len(sales["mfy_rows"]),
        "applications": len(sales["order_numbers"]), "dates": sales["dates"],
    }


def build_gnp_comparison_xlsx(sales: dict, comparison: dict, csv_filename: str, gnp_filename: str) -> Tuple[bytes, dict]:
    """Build a count-only, date-aware GNP/E-GAZ comparison workbook."""
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    formats = _analysis_formats(workbook)
    all_orders = comparison["orders"]
    no_sale_status = "0 SOTUV - E-GAZDA YO'Q"
    mismatches = [row for row in all_orders if row["status"] not in ("MOS", no_sale_status)]
    zero_transfer = [row for row in all_orders if row["status"] == no_sale_status]
    gnp_numbers = {row["order_no"] for row in all_orders}
    same_day_keys = {(row["order_no"], row["date"]) for row in all_orders if row["csv_sales"] > 0}
    gnp_dates = sorted({row["date"] for row in all_orders})
    csv_dates = sorted(sales["dates"])
    common_dates = sorted(set(gnp_dates).intersection(csv_dates))
    if not common_dates:
        date_status = "YO'Q - GNP sanalari E-GAZ CSV sanalarida topilmadi"
    elif set(gnp_dates).issubset(csv_dates):
        date_status = "HA - barcha GNP sanalari E-GAZ CSVda bor"
    else:
        date_status = "QISMAN - ayrim GNP sanalari E-GAZ CSVda yo'q"

    gnp_transferred = sum(row["transferred"] for row in all_orders)
    egaz_for_gnp = sum(row["csv_sales"] for row in all_orders)
    count_difference = egaz_for_gnp - gnp_transferred
    gnp_has_mahalla = any(clean_header(row.get("mahalla", "")) for row in all_orders)

    sheet = workbook.add_worksheet("Umumiy")
    sheet.merge_range("A1:B1", "E-GAZ VA GNP - SOTUV SONI TAQQOSLASH", formats["title"])
    sheet.set_row(0, 30)
    summary_values = [
        ("E-GAZ CSV manba fayllari", csv_filename),
        ("GNP XLSX manba fayli", gnp_filename),
        ("E-GAZ hisobot sanalari", ", ".join(csv_dates)),
        ("GNP buyurtma sanalari", ", ".join(gnp_dates)),
        ("GNP sanalari E-GAZ CSVda bormi", date_status),
        ("E-GAZ jami sotuv satrlari", sales["sales"]),
        ("E-GAZ tuman / shaharlar soni", len(sales["district_rows"])),
        ("E-GAZ tuman-MFY juftliklari soni", len(sales["mfy_rows"])),
        ("GNP faylida MFY ustuni", "Bor" if gnp_has_mahalla else "Yo'q - E-GAZ MFYlari ko'rsatildi"),
        ("GNP arizalari", len(all_orders)),
        ("Sana bir xil bo'lib, E-GAZda topilgan arizalar", comparison["same_day_found"]),
        ("MOS - to'liq tekshiruvdan o'tgan arizalar", comparison["matched"]),
        ("Tafovutli / tekshirish kerak arizalar", len(mismatches)),
        ("0 o'tkazma va E-GAZda sotuv satri yo'q", len(zero_transfer)),
        ("GNP o'tkazgan ballonlar", gnp_transferred),
        ("E-GAZ sotuvlari - GNP arizalari, ayni sana", egaz_for_gnp),
        ("Soni farqi (E-GAZ - GNP)", count_difference),
        ("GNPda bo'lmagan E-GAZ arizalari", len(comparison["csv_only"])),
    ]
    for row_index, (label, value) in enumerate(summary_values, start=3):
        sheet.write(row_index, 0, label, formats["total_label"])
        if isinstance(value, int):
            sheet.write_number(row_index, 1, value, formats["integer"])
        else:
            fmt = formats["bad"] if label == "GNP sanalari E-GAZ CSVda bormi" and value.startswith("YO'Q") else formats["text"]
            sheet.write(row_index, 1, value, fmt)
    note_row = len(summary_values) + 4
    note = "MOS faqat ariza raqami va sana topilib, tuman hamda o'tkazilgan/sotilgan son teng bo'lganda beriladi."
    if not gnp_has_mahalla:
        note += " GNPda MFY yo'q: MFYlar E-GAZ CSVdan ko'rsatildi, ular GNP bilan solishtirilmagan."
    sheet.merge_range(note_row, 0, note_row, 1, note, formats["bad"] if not common_dates else formats["text"])
    sheet.set_row(note_row, 38)
    sheet.set_column("A:A", 48)
    sheet.set_column("B:B", 82)

    headers = [
        "Ariza №", "GNP sana", "Tuman / shahar", "GNP MFY", "Buyurtma ballonlari",
        "O'tkazilgan ballonlar", "E-GAZ sotuvlari (shu sana)", "Soni farqi",
        "E-GAZ tuman(lar)i", "E-GAZ CSV MFY(lar)i", "E-GAZ manba fayli(lar)i",
        "Tuman tekshiruvi", "MFY tekshiruvi", "Boshqa sanadagi E-GAZ sotuvlari",
        "Boshqa sana manba fayllari", "Holat",
    ]

    def write_comparison_sheet(sheet_name, rows, title):
        ws = workbook.add_worksheet(sheet_name)
        ws.merge_range(0, 0, 0, len(headers) - 1, title, formats["title"])
        ws.set_row(0, 30)
        _analysis_write_table_header(ws, headers, formats)
        integer_columns = {4, 5, 6, 7, 13}
        for row_index, item in enumerate(rows, start=3):
            values = [
                item["order_no"], item["date"], item["district"], item.get("mahalla", ""),
                item["requested"], item["transferred"], item["csv_sales"], item["count_difference"],
                ", ".join(item["csv_districts"]), ", ".join(item["csv_mahallas"]),
                ", ".join(item["csv_files"]), item["district_match"], item["mahalla_match"],
                item["other_day_sales"], ", ".join(item["other_day_files"]), item["status"],
            ]
            for column, value in enumerate(values):
                if column in integer_columns:
                    ws.write_number(row_index, column, value, formats["integer"])
                else:
                    fmt = formats["ok"] if column == 15 and item["status"] == "MOS" else formats["bad"] if column == 15 and item["status"] != no_sale_status else formats["text"]
                    ws.write(row_index, column, value, fmt)
        ws.autofilter(2, 0, max(2, len(rows) + 2), len(headers) - 1)
        for col, width in enumerate([14, 13, 28, 24, 18, 19, 23, 15, 34, 32, 38, 22, 39, 26, 38, 29]):
            ws.set_column(col, col, width)

    write_comparison_sheet("Arizalar", all_orders, "GNP ARIZALARI - E-GAZ BILAN SOLISHTIRISH")
    write_comparison_sheet("Tafovutlar", mismatches, "TAFOVUTLI ARIZALAR - TEKSHIRISH KERAK")
    write_comparison_sheet("0 sotuv", zero_transfer, "0 O'TKAZMA - E-GAZDA SOTUV SATRI YO'Q")

    sheet = workbook.add_worksheet("CSV arizalari")
    csv_headers = [
        "Ariza №", "CSV sanalari", "Tuman / shahar(lar)i", "E-GAZ MFY(lar)i",
        "Manba fayli(lar)i", "Sotuv soni", "GNPda ariza bor", "GNP sana bilan topildi",
    ]
    sheet.merge_range(0, 0, 0, len(csv_headers) - 1, "E-GAZ CSV ARIZALARI - SONI BO'YICHA", formats["title"])
    sheet.set_row(0, 30)
    _analysis_write_table_header(sheet, csv_headers, formats)
    for row_index, (number, item) in enumerate(sorted(sales["order_numbers"].items()), start=3):
        has_same_day_gnp = any((number, date) in same_day_keys for date in item["dates"])
        values = [
            number, ", ".join(sorted(item["dates"])), ", ".join(sorted(item["districts"])),
            ", ".join(sorted(item["mahallas"])), ", ".join(sorted(item["files"])), item["sales"],
            "Ha" if number in gnp_numbers else "Yo'q", "Ha" if has_same_day_gnp else "Yo'q",
        ]
        for column, value in enumerate(values):
            if column == 5:
                sheet.write_number(row_index, column, value, formats["integer"])
            else:
                sheet.write(row_index, column, value, formats["text"])
    sheet.autofilter(2, 0, max(2, len(sales["order_numbers"]) + 2), len(csv_headers) - 1)
    for col, width in enumerate([15, 22, 36, 32, 38, 15, 18, 24]):
        sheet.set_column(col, col, width)

    district_groups = {}
    for row in all_orders:
        key = (row["date"], _analysis_group_key(row["district"]))
        item = district_groups.setdefault(key, {
            "date": row["date"], "name": row["district"], "orders": 0,
            "mismatches": 0, "zero": 0, "transferred": 0, "csv_sales": 0,
        })
        item["orders"] += 1
        item["mismatches"] += int(row["status"] not in ("MOS", no_sale_status))
        item["zero"] += int(row["status"] == no_sale_status)
        item["transferred"] += row["transferred"]
        item["csv_sales"] += row["csv_sales"]
    district_headers = [
        "GNP sana", "GNP tuman / shahar", "GNP arizalari", "Tafovutli arizalar",
        "0 o'tkazma / sotuv yo'q", "GNP o'tkazgan ballon", "E-GAZ sotgan ballon (shu sana)",
        "Soni farqi (E-GAZ - GNP)",
    ]
    sheet = workbook.add_worksheet("Tumanlar")
    sheet.merge_range(0, 0, 0, len(district_headers) - 1, "GNP BUYURTMALARI - TUMAN / SHAHAR KESIMI", formats["title"])
    sheet.set_row(0, 30)
    _analysis_write_table_header(sheet, district_headers, formats)
    sorted_groups = sorted(district_groups.values(), key=lambda row: (row["date"], _analysis_group_key(row["name"])))
    for row_index, item in enumerate(sorted_groups, start=3):
        values = [item["date"], item["name"], item["orders"], item["mismatches"], item["zero"], item["transferred"], item["csv_sales"], item["csv_sales"] - item["transferred"]]
        for column, value in enumerate(values):
            if column < 2:
                sheet.write(row_index, column, value, formats["text"])
            else:
                sheet.write_number(row_index, column, value, formats["integer"])
    total_row = len(sorted_groups) + 3
    sheet.write(total_row, 0, "JAMI", formats["total_label"])
    totals = [len(all_orders), len(mismatches), len(zero_transfer), gnp_transferred, egaz_for_gnp, count_difference]
    for column, value in enumerate(totals, start=2):
        sheet.write_number(total_row, column, value, formats["total_integer"])
    sheet.autofilter(2, 0, max(2, total_row - 1), len(district_headers) - 1)
    for col, width in enumerate([13, 32, 16, 20, 23, 22, 26, 24]):
        sheet.set_column(col, col, width)

    comparison_by_day = {(row["order_no"], row["date"]): row for row in all_orders}
    mfy_audit = defaultdict(lambda: {"gnp": set(), "matched": set(), "mismatched": set()})
    egaz_district_days = {}
    egaz_mfy_days = {}
    district_applications = defaultdict(set)
    for record in sales["records"]:
        date = record["date"]
        district_key = _analysis_group_key(record["district"])
        mahalla_key = _analysis_group_key(record["mahalla"])
        district_day_key = (date, district_key)
        district_day = egaz_district_days.setdefault(district_day_key, {
            "date": date, "district": record["district"], "sales": 0,
            "applications": set(), "mahallas": set(),
        })
        district_day["sales"] += 1
        district_day["applications"].add(record["order_no"])
        district_day["mahallas"].add(mahalla_key)
        district_applications[district_key].add(record["order_no"])
        mfy_day_key = (date, district_key, mahalla_key)
        mfy_day = egaz_mfy_days.setdefault(mfy_day_key, {
            "date": date, "district": record["district"], "mahalla": record["mahalla"],
            "sales": 0, "applications": set(),
        })
        mfy_day["sales"] += 1
        mfy_day["applications"].add(record["order_no"])

        order_key = (record["order_no"], record["date"])
        comparison_row = comparison_by_day.get(order_key)
        if comparison_row is None:
            continue
        outcome = "matched" if comparison_row["status"] == "MOS" else "mismatched"
        mfy_key = (record["date"], district_key, mahalla_key)
        mfy_audit[mfy_key]["gnp"].add(order_key)
        mfy_audit[mfy_key][outcome].add(order_key)

    used_names = {"umumiy", "arizalar", "tafovutlar", "0 sotuv", "csv arizalari", "tumanlar", "egaz tumanlar", "mfylar"}
    sales_districts = sorted(egaz_district_days.values(), key=lambda row: (row["date"], _analysis_group_key(row["district"])))
    sheet = workbook.add_worksheet("E-GAZ tumanlar")
    district_headers = ["E-GAZ sana", "E-GAZ tuman / shahar", "Sotuv soni", "Takrorlanmas arizalar", "MFY soni"]
    sheet.merge_range(0, 0, 0, len(district_headers) - 1, "E-GAZ SOTUVLARI - SANA VA TUMAN / SHAHAR KESIMI", formats["title"])
    sheet.set_row(0, 30)
    _analysis_write_table_header(sheet, district_headers, formats)
    for row_index, item in enumerate(sales_districts, start=3):
        values = [item["date"], item["district"], item["sales"], len(item["applications"]), len(item["mahallas"])]
        for column, value in enumerate(values):
            if column < 2:
                sheet.write(row_index, column, value, formats["text"])
            else:
                sheet.write_number(row_index, column, value, formats["integer"])
    total_row = len(sales_districts) + 3
    sheet.write(total_row, 0, "JAMI", formats["total_label"])
    for column, value in enumerate([sales["sales"], len(sales["order_numbers"]), len(egaz_mfy_days)], start=2):
        sheet.write_number(total_row, column, value, formats["total_integer"])
    sheet.autofilter(2, 0, max(2, total_row - 1), len(district_headers) - 1)
    sheet.set_column("A:A", 14)
    sheet.set_column("B:B", 36)
    sheet.set_column("C:E", 20)

    mfy_rows = sorted(egaz_mfy_days.values(), key=lambda row: (
        row["date"], _analysis_group_key(row["district"]), _analysis_group_key(row["mahalla"])
    ))
    mfy_headers = [
        "Sana", "Tuman / shahar", "E-GAZ MFY", "E-GAZ sotuv satrlari",
        "E-GAZdagi noyob arizalar", "GNPda topildi (ariza+sana)",
        "GNPda topilmadi (ariza+sana)", "MOS arizalar", "Tafovutli arizalar",
    ]
    sheet = workbook.add_worksheet("MFYlar")
    sheet.merge_range(0, 0, 0, len(mfy_headers) - 1, "E-GAZ SOTUVLARI - SANA VA MFY KESIMI", formats["title"])
    sheet.set_row(0, 30)
    mfy_note = (
        "Sotuv satri — E-GAZ qatori; noyob ariza — takrorlanmagan ariza raqami. Bir raqam bir nechta MFYda bo'lsa, "
        "MFY satrlarida takror ko'rinadi. GNPda topildi/topilmadi va MOS/tafovut ariza+sana juftligi bo'yicha; "
        "JAMI juftlikni bir marta sanaydi. MOS GNP buyurtmasining tuman va sotuv soni tekshiruvi. "
        "GNP faylida MFY ustuni bo'lmasa, MFY mosligi tekshirilmaydi. Sanalar farq qilsa, juftlik topilmaydi."
    )
    sheet.merge_range(1, 0, 1, len(mfy_headers) - 1, mfy_note, formats["note"])
    sheet.set_row(1, 54)
    _analysis_write_table_header(sheet, mfy_headers, formats)
    all_egaz_keys = set()
    all_gnp_keys = set()
    for row_index, item in enumerate(mfy_rows, start=3):
        audit = mfy_audit[(item["date"], _analysis_group_key(item["district"]), _analysis_group_key(item["mahalla"]))]
        egaz_keys = {(number, item["date"]) for number in item["applications"]}
        not_found_keys = egaz_keys - audit["gnp"]
        all_egaz_keys.update(egaz_keys)
        all_gnp_keys.update(audit["gnp"])
        values = [
            item["date"], item["district"], item["mahalla"], item["sales"],
            len(item["applications"]), len(audit["gnp"]), len(not_found_keys),
            len(audit["matched"]), len(audit["mismatched"]),
        ]
        for column, value in enumerate(values):
            if column < 3:
                sheet.write(row_index, column, value, formats["text"])
            else:
                sheet.write_number(row_index, column, value, formats["integer"])
    total_row = len(mfy_rows) + 3
    all_keys = set().union(*(audit["gnp"] for audit in mfy_audit.values())) if mfy_audit else set()
    matched_keys = set().union(*(audit["matched"] for audit in mfy_audit.values())) if mfy_audit else set()
    mismatched_keys = set().union(*(audit["mismatched"] for audit in mfy_audit.values())) if mfy_audit else set()
    sheet.write(total_row, 0, "JAMI", formats["total_label"])
    for column, value in enumerate([
        sales["sales"], len(sales["order_numbers"]), len(all_keys),
        len(all_egaz_keys - all_gnp_keys), len(matched_keys), len(mismatched_keys),
    ], start=3):
        sheet.write_number(total_row, column, value, formats["total_integer"])
    sheet.autofilter(2, 0, max(2, total_row - 1), len(mfy_headers) - 1)
    sheet.set_column("A:A", 14)
    sheet.set_column("B:C", 34)
    sheet.set_column("D:D", 20)
    sheet.set_column("E:E", 24)
    sheet.set_column("F:G", 29)
    sheet.set_column("H:I", 15)

    mfy_by_district = defaultdict(list)
    for item in mfy_rows:
        mfy_by_district[_analysis_group_key(item["district"])].append(item)
    district_labels = {}
    for district in sales_districts:
        district_labels.setdefault(_analysis_group_key(district["district"]), district["district"])
    for district_key, district_name in sorted(district_labels.items(), key=lambda item: item[0]):
        district_mfys = sorted(mfy_by_district[district_key], key=lambda row: (row["date"], -row["sales"], _analysis_group_key(row["mahalla"])))
        sheet_name = _analysis_sheet_name(district_name, used_names)
        sheet = workbook.add_worksheet(sheet_name)
        headers = [
            "Sana", "E-GAZ MFY", "Sotuv satrlari", "E-GAZdagi noyob arizalar",
            "GNPda topildi (ariza+sana)", "GNPda topilmadi (ariza+sana)", "MOS arizalar", "Tafovutli arizalar",
        ]
        sheet.merge_range(0, 0, 0, len(headers) - 1, f"{district_name} - MFY SOTUVLARI VA TEKSHIRUVI", formats["title"])
        sheet.set_row(0, 30)
        sheet.merge_range(1, 0, 1, len(headers) - 1, mfy_note, formats["note"])
        sheet.set_row(1, 54)
        _analysis_write_table_header(sheet, headers, formats)
        district_egaz_keys = set()
        district_gnp_keys = set()
        for row_index, item in enumerate(district_mfys, start=3):
            audit = mfy_audit[(item["date"], district_key, _analysis_group_key(item["mahalla"]))]
            egaz_keys = {(number, item["date"]) for number in item["applications"]}
            district_egaz_keys.update(egaz_keys)
            district_gnp_keys.update(audit["gnp"])
            values = [
                item["date"], item["mahalla"], item["sales"], len(item["applications"]),
                len(audit["gnp"]), len(egaz_keys - audit["gnp"]),
                len(audit["matched"]), len(audit["mismatched"]),
            ]
            for column, value in enumerate(values):
                if column < 2:
                    sheet.write(row_index, column, value, formats["text"])
                else:
                    sheet.write_number(row_index, column, value, formats["integer"])
        subtotal = len(district_mfys) + 3
        district_keys = set().union(*(mfy_audit[(item["date"], district_key, _analysis_group_key(item["mahalla"]))]["gnp"] for item in district_mfys)) if district_mfys else set()
        district_matched_keys = set().union(*(mfy_audit[(item["date"], district_key, _analysis_group_key(item["mahalla"]))]["matched"] for item in district_mfys)) if district_mfys else set()
        district_mismatch_keys = set().union(*(mfy_audit[(item["date"], district_key, _analysis_group_key(item["mahalla"]))]["mismatched"] for item in district_mfys)) if district_mfys else set()
        district_sales = sum(item["sales"] for item in district_mfys)
        district_apps = district_applications[district_key]
        sheet.write(subtotal, 1, "TUMAN JAMI", formats["total_label"])
        subtotal_values = [
            district_sales, len(district_apps), len(district_keys),
            len(district_egaz_keys - district_gnp_keys), len(district_matched_keys), len(district_mismatch_keys),
        ]
        for column, value in enumerate(subtotal_values, start=2):
            sheet.write_number(subtotal, column, value, formats["total_integer"])
        sheet.autofilter(2, 0, max(2, subtotal - 1), len(headers) - 1)
        sheet.set_column("A:A", 14)
        sheet.set_column("B:B", 34)
        sheet.set_column("C:C", 20)
        sheet.set_column("D:D", 24)
        sheet.set_column("E:F", 29)
        sheet.set_column("G:H", 15)

    workbook.close()
    return output.getvalue(), {
        "orders": len(all_orders), "matched": comparison["matched"],
        "mismatches": len(mismatches), "zero_transfer_no_csv": len(zero_transfer),
        "same_day_found": comparison["same_day_found"], "csv_only": len(comparison["csv_only"]),
        "gnp_transferred": gnp_transferred, "egaz_for_gnp": egaz_for_gnp,
        "count_difference": count_difference, "date_status": date_status,
    }


def build_gas_sales_xlsx(reports: List[dict]) -> Tuple[bytes, dict]:
    """Build a dedicated summary for inspector-level accepted/sold/returned reports."""
    rows = [row for report in reports for row in report["rows"]]
    if not rows:
        raise ValueError("Qabul va sotuv hisobotida ma’lumot yo‘q.")

    districts = {}
    for row in rows:
        key = header_key(row["organization"])
        item = districts.setdefault(key, {
            "name": re.sub(r"\s+", " ", row["organization"]).strip(),
            "accepted": 0,
            "sold": 0,
            "returned": 0,
            "inspectors": set(),
        })
        item["accepted"] += row["accepted"]
        item["sold"] += row["sold"]
        item["returned"] += row["returned"]
        if not row.get("is_unassigned"):
            item["inspectors"].add(header_key(row["inspector"]))

    district_rows = sorted(districts.values(), key=lambda item: (-item["accepted"], item["name"]))
    total_accepted = sum(row["accepted"] for row in rows)
    total_sold = sum(row["sold"] for row in rows)
    total_returned = sum(row["returned"] for row in rows)
    inspector_count = len({
        (header_key(row["organization"]), header_key(row["inspector"]))
        for row in rows if not row.get("is_unassigned")
    })
    periods = list(dict.fromkeys(report["period"] for report in reports))
    period_label = "; ".join(periods)

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {
        "constant_memory": True,
        "tmpdir": tempfile.gettempdir(),
        "strings_to_formulas": False,
        "strings_to_urls": False,
    })
    wb.set_properties({
        "title": f"{period_label} qabul va sotuv svodi",
        "subject": "Suyultirilgan gaz qabul, sotuv va qoldiq hisoboti",
        "author": APP_NAME,
        "comments": "Manba hisobotidagi inspektor qatorlari saqlangan va raygaz kesimida jamlangan.",
    })
    wb.set_calc_mode("auto")

    fmt_title = wb.add_format({"bold": True, "font_size": 16, "font_color": "white",
                               "bg_color": "#17365D", "align": "center", "valign": "vcenter"})
    fmt_header = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78",
                                "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    fmt_cell = wb.add_format({"border": 1, "valign": "top"})
    fmt_center = wb.add_format({"border": 1, "align": "center", "valign": "vcenter"})
    fmt_int = wb.add_format({"border": 1, "num_format": "#,##0", "align": "center"})
    fmt_pct = wb.add_format({"border": 1, "num_format": "0.00%", "align": "center"})
    fmt_total = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1,
                               "num_format": "#,##0", "align": "center"})
    fmt_total_label = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    fmt_total_pct = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1,
                                   "num_format": "0.00%", "align": "center"})
    fmt_ok = wb.add_format({"border": 1, "bg_color": "#E2F0D9", "font_color": "#375623",
                            "align": "center"})
    fmt_warn = wb.add_format({"border": 1, "bg_color": "#FFF2CC", "font_color": "#7F6000",
                              "text_wrap": True, "valign": "top"})
    fmt_hash = wb.add_format({"border": 1, "font_name": "Consolas", "font_size": 8})

    summary_headers = ["Райгаз", "Қабул қилинди", "Сотилди", "Қайтарилди", "Сотув %"]

    ws = wb.add_worksheet("Свод")
    ws.set_tab_color("#17365D")
    ws.merge_range("A1:E2", f"{period_label} — ҚАБУЛ ВА СОТУВ СВОДИ", fmt_title)
    ws.write_row("A4", ["Кўрсаткич", "Қиймат"], fmt_header)
    metrics = [
        ("Райгазлар сони", len(district_rows)),
        ("Чилангарлар сони", inspector_count),
        ("Қабул қилинди", total_accepted),
        ("Сотилди", total_sold),
        ("Қайтарилди", total_returned),
    ]
    for row_index, (label, value) in enumerate(metrics, start=4):
        ws.write(row_index, 0, label, fmt_cell)
        ws.write_number(row_index, 1, value, fmt_int)
    ws.write(9, 0, "Сотув фоизи", fmt_cell)
    ws.write_formula(9, 1, "=IFERROR(B8/B7,0)", fmt_pct,
                     total_sold / total_accepted if total_accepted else 0)

    table_start = 12
    ws.write_row(table_start, 0, summary_headers, fmt_header)
    for offset, item in enumerate(district_rows, start=1):
        row_index = table_start + offset
        excel_row = row_index + 1
        ws.write(row_index, 0, item["name"], fmt_cell)
        ws.write_number(row_index, 1, item["accepted"], fmt_int)
        ws.write_number(row_index, 2, item["sold"], fmt_int)
        ws.write_number(row_index, 3, item["returned"], fmt_int)
        ws.write_formula(row_index, 4, f"=IFERROR(C{excel_row}/B{excel_row},0)", fmt_pct,
                         item["sold"] / item["accepted"] if item["accepted"] else 0)
    total_row = table_start + len(district_rows) + 1
    ws.write(total_row, 0, "ЖАМИ", fmt_total_label)
    ws.write_number(total_row, 1, total_accepted, fmt_total)
    ws.write_number(total_row, 2, total_sold, fmt_total)
    ws.write_number(total_row, 3, total_returned, fmt_total)
    ws.write_formula(total_row, 4, f"=IFERROR(C{total_row + 1}/B{total_row + 1},0)", fmt_total_pct,
                     total_sold / total_accepted if total_accepted else 0)
    ws.set_column("A:A", 34)
    ws.set_column("B:E", 17)
    ws.freeze_panes(table_start + 1, 1)
    ws.autofilter(table_start, 0, max(table_start + 1, total_row - 1), 4)

    chart = wb.add_chart({"type": "column", "subtype": "stacked"})
    chart.add_series({
        "name": "Сотилди",
        "categories": ["Свод", table_start + 1, 0, total_row - 1, 0],
        "values": ["Свод", table_start + 1, 2, total_row - 1, 2],
        "fill": {"color": "#70AD47"},
        "border": {"color": "#548235"},
    })
    chart.add_series({
        "name": "Қайтарилди",
        "categories": ["Свод", table_start + 1, 0, total_row - 1, 0],
        "values": ["Свод", table_start + 1, 3, total_row - 1, 3],
        "fill": {"color": "#FFC000"},
        "border": {"color": "#BF9000"},
    })
    chart.set_title({"name": "Райгазлар бўйича сотув ва қайтарилган газ"})
    chart.set_y_axis({"major_gridlines": {"visible": False}, "num_format": "#,##0"})
    chart.set_legend({"position": "bottom"})
    chart.set_size({"width": 720, "height": 360})
    ws.insert_chart("G4", chart)

    ws = wb.add_worksheet("Райгаз")
    ws.set_tab_color("#5B9BD5")
    ws.write_row(0, 0, [*summary_headers[:-1], "Чилангар сони", summary_headers[-1]], fmt_header)
    for row_index, item in enumerate(district_rows, start=1):
        excel_row = row_index + 1
        ws.write(row_index, 0, item["name"], fmt_cell)
        ws.write_number(row_index, 1, item["accepted"], fmt_int)
        ws.write_number(row_index, 2, item["sold"], fmt_int)
        ws.write_number(row_index, 3, item["returned"], fmt_int)
        ws.write_number(row_index, 4, len(item["inspectors"]), fmt_int)
        ws.write_formula(row_index, 5, f"=IFERROR(C{excel_row}/B{excel_row},0)", fmt_pct,
                         item["sold"] / item["accepted"] if item["accepted"] else 0)
    ray_total_row = len(district_rows) + 1
    ws.write(ray_total_row, 0, "ЖАМИ", fmt_total_label)
    for column, value in enumerate((total_accepted, total_sold, total_returned, inspector_count), start=1):
        ws.write_number(ray_total_row, column, value, fmt_total)
    ws.write_formula(ray_total_row, 5, f"=IFERROR(C{ray_total_row + 1}/B{ray_total_row + 1},0)",
                     fmt_total_pct, total_sold / total_accepted if total_accepted else 0)
    ws.set_column("A:A", 34)
    ws.set_column("B:F", 17)
    ws.freeze_panes(1, 1)
    ws.autofilter(0, 0, max(1, ray_total_row - 1), 5)

    ws = wb.add_worksheet("Чилангарлар")
    ws.set_tab_color("#70AD47")
    detail_headers = ["Давр", "Райгаз", "Инспектор", "Қабул қилинди", "Сотилди",
                      "Қайтарилди", "Сотув %", "Манба файл", "Манба қатор"]
    ws.write_row(0, 0, detail_headers, fmt_header)
    report_period = {report["filename"]: report["period"] for report in reports}
    for row_index, row in enumerate(sorted(rows, key=lambda item: (
            header_key(item["organization"]), -item["accepted"], header_key(item["inspector"]))), start=1):
        excel_row = row_index + 1
        ws.write(row_index, 0, report_period[row["source_file"]], fmt_center)
        ws.write(row_index, 1, re.sub(r"\s+", " ", row["organization"]).strip(), fmt_cell)
        ws.write(row_index, 2, row["inspector"], fmt_cell)
        ws.write_number(row_index, 3, row["accepted"], fmt_int)
        ws.write_number(row_index, 4, row["sold"], fmt_int)
        ws.write_number(row_index, 5, row["returned"], fmt_int)
        ws.write_formula(row_index, 6, f"=IFERROR(E{excel_row}/D{excel_row},0)", fmt_pct,
                         row["sold"] / row["accepted"] if row["accepted"] else 0)
        ws.write(row_index, 7, row["source_file"], fmt_cell)
        ws.write_number(row_index, 8, row["source_row"], fmt_int)
    ws.autofilter(0, 0, len(rows), len(detail_headers) - 1)
    ws.freeze_panes(1, 3)
    ws.set_column("A:A", 23)
    ws.set_column("B:B", 34)
    ws.set_column("C:C", 43)
    ws.set_column("D:G", 16)
    ws.set_column("H:H", 30)
    ws.set_column("I:I", 13)

    ws = wb.add_worksheet("Назорат")
    ws.set_tab_color("#FFC000")
    ws.merge_range("A1:N2", "МАНБА ФАЙЛЛАР ВА БАЛАНС НАЗОРАТИ", fmt_title)
    control_headers = ["Файл", "Давр", "Чилангар қаторлари", "Ҳисоб: қабул", "Файл: қабул",
                       "Фарқ", "Ҳисоб: сотув", "Файл: сотув", "Фарқ", "Ҳисоб: қайтарилди",
                       "Файл: қайтарилди", "Фарқ", "Ҳолат", "SHA-256"]
    ws.write_row(3, 0, control_headers, fmt_header)
    all_warnings = []
    for row_index, report in enumerate(reports, start=4):
        calculated = {
            "accepted": sum(row["accepted"] for row in report["rows"]),
            "sold": sum(row["sold"] for row in report["rows"]),
            "returned": sum(row["returned"] for row in report["rows"]),
        }
        if report["stated_total"] is None:
            stated = calculated
            differences = {key: 0 for key in calculated}
        else:
            stated = {key: report["stated_total"].get(key) for key in calculated}
            differences = {
                key: None if stated[key] is None else stated[key] - calculated[key]
                for key in calculated
            }
        report_warnings = list(report["warnings"])
        for key, district_total in report["district_totals"].items():
            district_calculated = {
                field: sum(row[field] for row in report["rows"]
                           if header_key(row["organization"]) == key)
                for field in ("accepted", "sold", "returned")
            }
            if any(district_total[field] != district_calculated[field] for field in district_calculated):
                report_warnings.append(
                    f"{report['filename']}: {district_total['name']} jami qatori inspektorlar yig‘indisiga mos emas."
                )
        status = "MOS" if not report_warnings and not any(
            value not in (None, 0) for value in differences.values()
        ) else "TEKSHIRING"
        values = [
            report["filename"], report["period"], len(report["rows"]),
            calculated["accepted"], stated["accepted"], differences["accepted"],
            calculated["sold"], stated["sold"], differences["sold"],
            calculated["returned"], stated["returned"], differences["returned"],
        ]
        for column, value in enumerate(values):
            if value is None:
                ws.write_blank(row_index, column, None, fmt_int)
            elif isinstance(value, int):
                ws.write_number(row_index, column, value, fmt_int)
            else:
                ws.write(row_index, column, value, fmt_cell)
        ws.write(row_index, 12, status, fmt_ok if status == "MOS" else fmt_warn)
        ws.write(row_index, 13, report["sha256"], fmt_hash)
        all_warnings.extend(report_warnings)
    if all_warnings:
        warning_row = 6 + len(reports)
        ws.write(warning_row, 0, "Огоҳлантиришлар", fmt_total_label)
        ws.merge_range(warning_row, 1, warning_row + max(1, len(all_warnings)), 13,
                       "\n".join("• " + warning for warning in all_warnings), fmt_warn)
    ws.set_column("A:B", 29)
    ws.set_column("C:M", 15)
    ws.set_column("N:N", 68)
    ws.freeze_panes(4, 0)

    wb.close()
    data = output.getvalue()
    summary = {
        "report_type": "gas-sales",
        "total": len(rows),
        "raygaz": len(district_rows),
        "inspectors": inspector_count,
        "accepted": total_accepted,
        "sold": total_sold,
        "returned": total_returned,
        "sales_percent": round(total_sold * 100 / total_accepted, 2) if total_accepted else 0,
        "period": period_label,
        "files": len(reports),
    }
    logging.info("GAS SALES OUTPUT | rows=%s | raygaz=%s | accepted=%s | sold=%s | returned=%s",
                 len(rows), len(district_rows), total_accepted, total_sold, total_returned)
    return data, summary


EGAZ_MONTHS_CYRILLIC = [
    "", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]
EGAZ_MONTHS_LATIN = [
    "", "Yanvar", "Fevral", "Mart", "Aprel", "May", "Iyun",
    "Iyul", "Avgust", "Sentyabr", "Oktyabr", "Noyabr", "Dekabr",
]
EGAZ_MONTH_ALIASES = {
    "yanvar": 1, "janvar": 1, "январ": 1, "январь": 1,
    "fevral": 2, "феврал": 2, "февраль": 2,
    "mart": 3, "март": 3,
    "aprel": 4, "апрел": 4, "апрель": 4,
    "may": 5, "mai": 5, "май": 5,
    "iyun": 6, "iun": 6, "июн": 6, "июнь": 6,
    "iyul": 7, "iul": 7, "июл": 7, "июль": 7,
    "avgust": 8, "avgyst": 8, "august": 8, "август": 8,
    "sentyabr": 9, "sentabr": 9, "сентябр": 9, "сентябрь": 9,
    "oktyabr": 10, "oktabr": 10, "октябр": 10, "октябрь": 10,
    "noyabr": 11, "ноябр": 11, "ноябрь": 11,
    "dekabr": 12, "декабр": 12, "декабрь": 12,
}
EGAZ_MONTH_MAX_DAYS = [0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]


def egaz_period_from_filename(filename: str) -> dict:
    """Return the reporting period encoded in an E-GAZ export filename."""
    source_name = Path(filename).name
    iso_match = re.search(
        r"(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})", source_name
    )
    if iso_match:
        try:
            date_value = datetime(
                int(iso_match.group("year")),
                int(iso_match.group("month")),
                int(iso_match.group("day")),
            )
        except ValueError as exc:
            raise ValueError(f"{filename}: fayl nomidagi sana noto‘g‘ri.") from exc
        start_day = 1
        day = date_value.day
        month = date_value.month
        title_month = EGAZ_MONTHS_CYRILLIC[month]
        title_suffix = "ҳолатига"
        year = date_value.year
    else:
        range_match = re.search(
            r"(?<!\d)(?P<start>\d{1,2})\s*[-–—_]\s*(?P<day>\d{1,2})"
            r"\s*[-_. ]*\s*(?P<month>[^\W\d_]+)",
            Path(filename).stem,
            re.IGNORECASE,
        )
        month = (
            EGAZ_MONTH_ALIASES.get(range_match.group("month").casefold())
            if range_match else None
        )
        if not range_match or month is None:
            raise ValueError(
                f"{filename}: fayl nomidan hisobot davri topilmadi. "
                "Nomda 2026-09-23 yoki 1-31 Avgust ko‘rinishini yozing."
            )
        start_day = int(range_match.group("start"))
        day = int(range_match.group("day"))
        if not 1 <= start_day <= day <= EGAZ_MONTH_MAX_DAYS[month]:
            raise ValueError(f"{filename}: fayl nomidagi kun oralig‘i noto‘g‘ri.")
        title_month = EGAZ_MONTHS_CYRILLIC[month]
        title_suffix = "\u04b3\u043e\u043b\u0430\u0442\u0438\u0433\u0430"
        year = None

    latin_month = EGAZ_MONTHS_LATIN[month]
    return {
        "start_day": start_day,
        "day": day,
        "month": month,
        "year": year,
        "title": f"{start_day}-{day} {title_month} {title_suffix}",
        "sheet": f"{start_day}-{day} {title_month}",
        "filename": f"E-GAZ_{start_day}-{day}_{latin_month}_holatiga.xlsx",
    }


def build_egaz_xlsx(report: dict, source_filename: str) -> Tuple[bytes, dict]:
    """Build the compact E-GAZ workbook used by the separate E-GAZ page."""
    rows = report.get("rows") or []
    if not rows:
        raise ValueError("E-GAZ hisobotida chilangar ma’lumotlari topilmadi.")
    period = egaz_period_from_filename(source_filename)

    grouped = {}
    district_order = []
    for row in rows:
        key = header_key(row["organization"])
        if key not in grouped:
            grouped[key] = {"name": row["organization"], "rows": []}
            district_order.append(key)
        grouped[key]["rows"].append(row)

    district_totals = report.get("district_totals") or {}
    table_rows = []
    for key in district_order:
        group = grouped[key]
        table_rows.extend(("detail", item) for item in group["rows"])
        calculated = {
            "accepted": sum(item["accepted"] for item in group["rows"]),
            "sold": sum(item["sold"] for item in group["rows"]),
        }
        total = district_totals.get(key) or {
            "name": group["name"],
            "accepted": calculated["accepted"],
            "sold": calculated["sold"],
        }
        total = dict(total)
        total["returned"] = total["accepted"] - total["sold"]
        table_rows.append(("subtotal", total))

    total_accepted = sum(item["accepted"] for kind, item in table_rows if kind == "subtotal")
    total_sold = sum(item["sold"] for kind, item in table_rows if kind == "subtotal")
    total_unsold = total_accepted - total_sold
    inspector_rows = sum(1 for kind, _ in table_rows if kind == "detail")

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {
        "strings_to_formulas": False,
        "strings_to_urls": False,
        "tmpdir": tempfile.gettempdir(),
    })
    wb.set_properties({
        "title": period["title"],
        "subject": "E-GAZ qabul, sotuv va sotilmagan gaz hisoboti",
        "author": APP_NAME,
    })
    wb.set_calc_mode("auto")
    ws = wb.add_worksheet(period["sheet"][:31])

    border = {"border": 1, "border_color": "#000000"}
    fmt_title = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 15,
                               "align": "center", "valign": "vcenter"})
    fmt_title_long = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 12,
                                    "align": "center", "valign": "vcenter"})
    fmt_title_small = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 13,
                                     "align": "center", "valign": "vcenter"})
    fmt_period = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 12,
                                "bg_color": "#00B0F0", "align": "center", "valign": "vcenter", **border})
    fmt_header_text = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 11,
                                     "align": "center", "valign": "vcenter", "text_wrap": True, **border})
    fmt_header_unsold = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 11,
                                       "bg_color": "#FFFF00", "align": "center", "valign": "vcenter",
                                       "text_wrap": True, **border})
    fmt_header_ru = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 10,
                                   "align": "center", "valign": "vcenter", **border})
    fmt_text = wb.add_format({"font_name": "Arial", "font_size": 10, "align": "center",
                              "valign": "vcenter", **border})
    fmt_number = wb.add_format({"font_name": "Arial", "font_size": 10, "align": "center",
                                "valign": "vcenter", "num_format": "0", **border})
    fmt_subtotal_text = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 10,
                                       "bg_color": "#FFFF00", "align": "center", "valign": "vcenter", **border})
    fmt_subtotal_number = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 10,
                                         "bg_color": "#FFFF00", "align": "center", "valign": "vcenter",
                                         "num_format": "0", **border})
    fmt_grand_text = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 11,
                                    "bg_color": "#F4B183", "align": "center", "valign": "vcenter", **border})
    fmt_grand_number = wb.add_format({"bold": True, "font_name": "Arial", "font_size": 11,
                                      "bg_color": "#F4B183", "align": "center", "valign": "vcenter",
                                      "num_format": "0", **border})

    ws.merge_range("A1:F1", '"Худудгаз Андижон" ГТБ тасарруфидаги шаҳар-тумангаз таъминоти суюлтирилган газ бўлими чилангар-', fmt_title_long)
    ws.merge_range("A2:F2", "таъминотчилари томонидан суюлтирилган газни E-gaz дастурида қабул қилиш ва сотиш тўғрисида", fmt_title_small)
    ws.merge_range("A3:F3", "М А Ъ Л У М О Т", fmt_title)
    ws.merge_range("C5:F5", period["title"], fmt_period)
    ws.merge_range("A6:B6", "Шаҳар-туман номи ва Чилангарларни Ф.И.Ш", fmt_header_text)
    ws.write("C6", "E-gaz дастурида қабул қилинган газ", fmt_header_text)
    ws.write("D6", "E-gaz дастурида сотилган газ", fmt_header_text)
    ws.write("E6", "E-gaz дастурида сотилмаган газ", fmt_header_unsold)
    ws.write("F6", "%", fmt_header_text)
    ws.write_row("A7", ["Учреждение", "Инспектор", "Принял", "Реализовал", "Вернул", "%"], fmt_header_ru)

    row_index = 7
    for kind, item in table_rows:
        excel_row = row_index + 1
        if kind == "detail":
            ws.write(row_index, 0, item["organization"], fmt_text)
            ws.write(row_index, 1, item["inspector"], fmt_text)
            number_format = fmt_number
        else:
            organization = item.get("name") or grouped.get(header_key(item.get("name", "")), {}).get("name", "")
            ws.write(row_index, 0, organization, fmt_subtotal_text)
            ws.write(row_index, 1, f"--- {organization} ЖАМИ ---", fmt_subtotal_text)
            number_format = fmt_subtotal_number
        ws.write_number(row_index, 2, item["accepted"], number_format)
        ws.write_number(row_index, 3, item["sold"], number_format)
        ws.write_formula(row_index, 4, f"=C{excel_row}-D{excel_row}", number_format,
                         item["accepted"] - item["sold"])
        percent = round(item["sold"] * 100 / item["accepted"]) if item["accepted"] else 0
        ws.write_formula(row_index, 5, f"=IFERROR(D{excel_row}/C{excel_row}*100,0)", number_format, percent)
        row_index += 1

    grand_excel_row = row_index + 1
    ws.merge_range(row_index, 0, row_index, 1, "Жами", fmt_grand_text)
    ws.write_number(row_index, 2, total_accepted, fmt_grand_number)
    ws.write_number(row_index, 3, total_sold, fmt_grand_number)
    ws.write_formula(row_index, 4, f"=C{grand_excel_row}-D{grand_excel_row}", fmt_grand_number, total_unsold)
    grand_percent = round(total_sold * 100 / total_accepted) if total_accepted else 0
    ws.write_formula(row_index, 5, f"=IFERROR(D{grand_excel_row}/C{grand_excel_row}*100,0)",
                     fmt_grand_number, grand_percent)

    ws.set_column("A:A", 30)
    ws.set_column("B:B", 55)
    ws.set_column("C:C", 10.5)
    ws.set_column("D:D", 15)
    ws.set_column("E:E", 10.5)
    ws.set_column("F:F", 9)
    ws.set_row(0, 23)
    ws.set_row(1, 22)
    ws.set_row(2, 22)
    ws.set_row(4, 22)
    ws.set_row(5, 78)
    ws.set_row(6, 21)
    ws.freeze_panes(7, 2)
    ws.autofilter(6, 0, max(7, row_index - 1), 5)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(0, 6)
    ws.set_margins(left=0.2, right=0.2, top=0.35, bottom=0.35)
    wb.close()

    data = output.getvalue()
    summary = {
        "rows": inspector_rows,
        "districts": len(district_order),
        "accepted": total_accepted,
        "sold": total_sold,
        "unsold": total_unsold,
        "percent": grand_percent,
        "period": period["title"],
        "output_name": period["filename"],
    }
    logging.info("EGAZ OUTPUT | rows=%s | districts=%s | accepted=%s | sold=%s | unsold=%s",
                 inspector_rows, len(district_order), total_accepted, total_sold, total_unsold)
    return data, summary


def build_xlsx(all_rows: List[dict], file_stats: List[dict], warnings: List[str],
               source_manifest: List[dict], reference: dict | None = None) -> Tuple[bytes, dict]:
    total = len(all_rows)
    if total == 0:
        raise ValueError("Ma'lumot yo'q.")
    if total >= EXCEL_MAX_ROWS:
        raise ValueError(
            f"Jami {total:,} ta yozuv — Excel varag'i sig'imidan ({EXCEL_MAX_ROWS - 1:,} qator) ko'p. "
            "Fayllarni bir necha qismga bo'lib svod qiling."
        )

    snapshot_date = all_rows[0]["Бугунги реализация"].date()
    raygaz_names = sorted({r["Райгаз"] for r in all_rows})
    mahalla_keys = sorted({(r["Райгаз"], r["Маҳалла"]) for r in all_rows})
    code_counter = Counter(r["Абонент код"] for r in all_rows)
    duplicate_codes = {k: v for k, v in code_counter.items() if v > 1}
    duplicate_extra = sum(v - 1 for v in duplicate_codes.values())

    interval_counts = Counter(r["Оралиқ"] for r in all_rows)
    raygaz_counts = defaultdict(Counter)
    mahalla_counts = defaultdict(Counter)
    for r in all_rows:
        raygaz_counts[r["Райгаз"]][r["Оралиқ"]] += 1
        mahalla_counts[(r["Райгаз"], r["Маҳалла"])][r["Оралиқ"]] += 1

    # New summaries start at day 31. Old references may include 1-30;
    # day 0 stays in the control group and is never counted as 1-30.
    detailed_counts = defaultdict(Counter)
    detailed_names = {}
    zero_day_count = 0
    for r in all_rows:
        key = header_key(r["Райгаз"])
        detailed_names.setdefault(key, r["Райгаз"])
        days = r["Кун фарқи"]
        if days == 0:
            zero_day_count += 1
            continue
        bucket = "1-30" if days <= 30 else (r["Оралиқ"] if days <= 70 else "71+")
        detailed_counts[key][bucket] += 1
    detailed_keys = sorted(detailed_names, key=lambda key: (-sum(detailed_counts[key].values()), detailed_names[key]))
    if reference:
        detailed_keys = ([key for key in reference["districts"] if key in detailed_names]
                         + [key for key in detailed_keys if key not in reference["districts"]])

    raygaz_sorted = sorted(raygaz_names, key=lambda x: (-sum(raygaz_counts[x].values()), x))
    mahalla_sorted = sorted(mahalla_keys, key=lambda x: (x[0], -sum(mahalla_counts[x].values()), x[1]))

    output = io.BytesIO()
    # constant_memory: katta bazada qatorlar vaqtinchalik faylga yoziladi (barcha varaqlar
    # qatorma-qator tartibda to'ldiriladi). strings_to_*: "=..." yoki "http..." bilan
    # boshlangan manba matnlari formula/havolaga aylanib buzilmasin.
    wb = xlsxwriter.Workbook(output, {
        "constant_memory": True,
        "tmpdir": tempfile.gettempdir(),
        "strings_to_formulas": False,
        "strings_to_urls": False,
    })
    wb.set_properties({
        "title": f"{snapshot_date.isoformat()} holati bo'yicha svod",
        "subject": "Avtomatik CSV nazorat va svod",
        "author": "SVOD TIZIMI",
        "comments": "Manba yozuvlar o'chirilmagan. Hosila ustunlar sanalardan hisoblangan.",
    })
    wb.set_calc_mode("auto")

    # Formats
    fmt_title = wb.add_format({"bold": True, "font_size": 16, "font_color": "white",
                               "bg_color": "#17365D", "align": "center", "valign": "vcenter"})
    fmt_header = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78",
                                "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    fmt_cell = wb.add_format({"border": 1, "valign": "top"})
    fmt_center = wb.add_format({"border": 1, "align": "center", "valign": "vcenter"})
    fmt_int = wb.add_format({"border": 1, "num_format": "#,##0", "align": "center"})
    fmt_pct = wb.add_format({"border": 1, "num_format": "0.00%", "align": "center"})
    fmt_date = wb.add_format({"border": 1, "num_format": "yyyy-mm-dd", "align": "center"})
    fmt_total = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "num_format": "#,##0"})
    fmt_total_pct = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "num_format": "0.00%"})
    fmt_label = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
    fmt_ok = wb.add_format({"border": 1, "bg_color": "#E2F0D9", "font_color": "#375623", "align": "center"})
    fmt_warn = wb.add_format({"border": 1, "bg_color": "#FFF2CC", "font_color": "#7F6000",
                              "align": "left", "valign": "top", "text_wrap": True})
    fmt_hash = wb.add_format({"border": 1, "font_name": "Consolas", "font_size": 8})
    fmt_note = wb.add_format({"italic": True, "font_color": "#666666", "text_wrap": True})
    fmt_dup = wb.add_format({"border": 1, "bg_color": "#FCE8E6"})
    fmt_unknown = wb.add_format({"border": 1, "bg_color": "#F2F4F7", "font_color": "#667085", "align": "center"})
    fmt_diff = wb.add_format({"border": 1, "bg_color": "#FCE8E6", "font_color": "#B42318", "num_format": "+#,##0;-#,##0;0"})

    # Свод
    ws = wb.add_worksheet("Свод")
    ws.set_tab_color("#17365D")
    ws.merge_range("A1:F2", f"{snapshot_date.strftime('%d.%m.%Y')} ҲОЛАТИ БЎЙИЧА СВОД", fmt_title)
    ws.write_row("A4", ["Кўрсаткич", "Қиймат"], fmt_header)
    metrics = [
        ("Жами абонент", total),
        ("Райгазлар сони", len(raygaz_names)),
        ("Маҳаллалар сони", len(mahalla_keys)),
        ("Манба файллар сони", len(file_stats)),
        ("Дубликат қўшимча ёзув", duplicate_extra),
    ]
    for idx, (name, val) in enumerate(metrics, start=4):
        ws.write(idx, 0, name, fmt_cell)
        ws.write_number(idx, 1, val, fmt_int)

    ws.write_row("A11", ["Оралиқ (кун)", "Абонент сони", "Улуши"], fmt_header)
    svod_labels = [x for x in INTERVALS if interval_counts[x] > 0 or x != "0-30 (НАЗОРАТ)"]
    start = 11
    for r_idx, lab in enumerate(svod_labels, start=start):
        cnt = interval_counts[lab]
        ws.write(r_idx, 0, lab, fmt_center)
        ws.write_number(r_idx, 1, cnt, fmt_int)
        ws.write_number(r_idx, 2, cnt / total if total else 0, fmt_pct)
    total_row = start + len(svod_labels)
    ws.write(total_row, 0, "ЖАМИ", fmt_total)
    ws.write_number(total_row, 1, total, fmt_total)
    ws.write_number(total_row, 2, 1, fmt_total_pct)
    note_row = total_row + 3
    ws.write(note_row, 0, "Изоҳ", fmt_label)
    note_text = (
        "Оралиқлар файл номидан эмас, “Бугунги реализация” ва “Сўнги реализация” "
        "саналарининг ҳақиқий кун фарқидан ҳисобланди. Манба қаторлар ўчирилмади."
    )
    ws.merge_range(note_row, 1, note_row + 1, 5, note_text, fmt_note)
    if warnings:
        wr = note_row + 3
        ws.write(wr, 0, "Огоҳлантириш", fmt_label)
        ws.merge_range(wr, 1, wr + max(1, len(warnings)), 5, "\n".join("• " + x for x in warnings), fmt_warn)

    ws.set_column("A:A", 26)
    ws.set_column("B:C", 18)
    ws.set_column("D:F", 15)

    # Chart
    chart = wb.add_chart({"type": "column"})
    last = start + len(svod_labels) - 1
    chart.add_series({
        "name": "Абонент сони",
        "categories": ["Свод", start, 0, last, 0],
        "values": ["Свод", start, 1, last, 1],
        "fill": {"color": "#5B9BD5"},
        "border": {"color": "#2F5597"},
        "data_labels": {"value": True},
    })
    chart.set_title({"name": "Оралиқлар бўйича абонентлар"})
    chart.set_legend({"none": True})
    chart.set_y_axis({"major_gridlines": {"visible": False}, "num_format": "#,##0"})
    chart.set_size({"width": 720, "height": 360})
    ws.insert_chart("E4", chart)

    # One row per district, in the same 31-35 ... 71+ layout as the supplied screenshot.
    ws = wb.add_worksheet("Давомилик")
    ws.set_tab_color("#00A6A6")
    detailed_headers = ["Райгаз", *DISPLAY_BUCKETS, "Жами"]
    ws.write_row(0, 0, detailed_headers, fmt_header)
    ws.set_row(0, 26)
    for row_index, key in enumerate(detailed_keys, start=1):
        name = reference["names"].get(key, detailed_names.get(key, key)) if reference else detailed_names[key]
        ws.write(row_index, 0, name, fmt_cell)
        for column, bucket in enumerate(DISPLAY_BUCKETS, start=1):
            ws.write_number(row_index, column, detailed_counts[key][bucket], fmt_int)
        ws.write_formula(row_index, 10, f"=SUM(B{row_index + 1}:J{row_index + 1})",
                         fmt_int, sum(detailed_counts[key][bucket] for bucket in DISPLAY_BUCKETS))
    total_row = len(detailed_keys) + 1
    ws.write(total_row, 0, "ЖАМИ", fmt_total)
    for column, bucket in enumerate(DISPLAY_BUCKETS, start=1):
        excel_col = xlsxwriter.utility.xl_col_to_name(column)
        cached = sum(detailed_counts[key][bucket] for key in detailed_keys)
        ws.write_formula(total_row, column, f"=SUM({excel_col}2:{excel_col}{total_row})", fmt_total, cached)
    ws.write_formula(total_row, 10, f"=SUM(K2:K{total_row})", fmt_total,
                     sum(detailed_counts[key][bucket] for key in detailed_keys for bucket in DISPLAY_BUCKETS))
    ws.merge_range(total_row + 2, 0, total_row + 2, 10,
                   f"{snapshot_date.strftime('%d.%m.%Y')} ҳолати. 0–30 кунлик ёзувлар бу жадвалга киритилмади; "
                   "улар 'Свод' ва 'База'да сақланган.", fmt_note)
    ws.set_row(total_row + 2, 28)
    ws.set_column("A:A", 31)
    ws.set_column("B:K", 13)
    ws.freeze_panes(1, 1)
    ws.autofilter(0, 0, max(1, total_row - 1), 10)

    comparison_mismatches = None
    if reference:
        compare_buckets = reference["buckets"]
        compare_keys = list(reference["districts"]) + [key for key in detailed_keys if key not in reference["districts"]]
        comparison_mismatches = 0
        for key in compare_keys:
            cell_differences = sum(
                detailed_counts[key][bucket]
                != (reference["districts"][key][bucket] if key in reference["districts"] else 0)
                for bucket in compare_buckets
            )
            # Even an all-zero row is structurally different when it exists on only one side.
            if ((key not in reference["districts"] or key not in detailed_names)
                    and cell_differences == 0):
                cell_differences = 1
            comparison_mismatches += cell_differences

        # Keep the source matrix intact. The red suffix is the current amount
        # being subtracted, not the already-calculated remainder.
        ws = wb.add_worksheet("Солиштириш")
        ws.set_tab_color("#E69138")
        shown_headers = ([reference["first_header"]]
                         + [reference["bucket_headers"][bucket] for bucket in compare_buckets]
                         + reference["special_headers"]
                         + (["Жами"] if reference["has_total_column"] else []))
        last_shown_col = len(shown_headers) - 1
        title_start = max(1, last_shown_col - 3)
        fmt_matrix_title = wb.add_format({"bold": True, "bg_color": "#A7DDE8", "border": 1,
                                          "align": "center", "valign": "vcenter"})
        fmt_matrix_red = wb.add_format({"font_color": "#E60000"})
        matrix_formats = {}
        for color, bg in (("plain", "#FFFFFF"), ("gray", "#A6A6A6"), ("yellow", "#FFFF00")):
            matrix_formats[color] = {
                "header": wb.add_format({"bold": True, "bg_color": bg, "border": 1,
                                          "align": "center", "valign": "vcenter", "text_wrap": True}),
                "number": wb.add_format({"bg_color": bg, "border": 1, "align": "center",
                                          "valign": "vcenter", "num_format": "0"}),
                "total": wb.add_format({"bold": True, "bg_color": bg, "border": 1,
                                         "align": "center", "valign": "vcenter", "num_format": "0"}),
            }

        def matrix_color(bucket):
            if bucket in ("31-35", "36-40"):
                return "gray"
            if bucket in ("66-70", "71+"):
                return "yellow"
            return "plain"

        def write_matrix_count(row, col, baseline, subtracted, cell_format):
            if subtracted:
                ws.write_rich_string(row, col, str(baseline), fmt_matrix_red,
                                     f" -{subtracted}", cell_format)
            else:
                ws.write_number(row, col, baseline, cell_format)

        ws.merge_range(0, title_start, 0, last_shown_col,
                       reference["title"] or f"{snapshot_date.strftime('%d.%m.%Y')} Давомилик",
                       fmt_matrix_title)
        ws.set_row(0, 22)
        for col, label in enumerate(shown_headers):
            color = matrix_color(compare_buckets[col - 1]) if 1 <= col <= len(compare_buckets) else "plain"
            ws.write(1, col, label, matrix_formats[color]["header"])
        ws.set_row(1, 96)

        for offset, key in enumerate(compare_keys):
            row_index = offset + 2
            name = reference["names"].get(key, detailed_names.get(key, key))
            ws.write_string(row_index, 0, name, matrix_formats["plain"]["number"])
            for col, bucket in enumerate(compare_buckets, start=1):
                expected = reference["districts"][key][bucket] if key in reference["districts"] else 0
                actual = detailed_counts[key][bucket] if key in detailed_names else 0
                write_matrix_count(row_index, col, expected, actual,
                                   matrix_formats[matrix_color(bucket)]["number"])
            next_col = len(compare_buckets) + 1
            for extra in reference["special_values"].get(key, [None] * len(reference["special_headers"])):
                if extra is None:
                    ws.write_blank(row_index, next_col, None, matrix_formats["plain"]["number"])
                else:
                    ws.write_number(row_index, next_col, extra, matrix_formats["plain"]["number"])
                next_col += 1
            if reference["has_total_column"]:
                expected = sum(reference["districts"][key][bucket] for bucket in compare_buckets) if key in reference["districts"] else 0
                actual = sum(detailed_counts[key][bucket] for bucket in compare_buckets) if key in detailed_names else 0
                write_matrix_count(row_index, next_col, expected, actual,
                                   matrix_formats["plain"]["number"])

        matrix_total_row = len(compare_keys) + 2
        ws.write_string(matrix_total_row, 0, "Жами", matrix_formats["plain"]["total"])
        for col, bucket in enumerate(compare_buckets, start=1):
            expected = reference["totals"][bucket]
            actual = sum(detailed_counts[key][bucket] for key in detailed_keys)
            write_matrix_count(matrix_total_row, col, expected, actual,
                               matrix_formats[matrix_color(bucket)]["total"])
        next_col = len(compare_buckets) + 1
        for index in range(len(reference["special_headers"])):
            expected = sum((values[index] or 0) for values in reference["special_values"].values())
            ws.write_number(matrix_total_row, next_col, expected, matrix_formats["plain"]["total"])
            next_col += 1
        if reference["has_total_column"]:
            expected = sum(reference["totals"].values())
            actual = sum(detailed_counts[key][bucket] for key in detailed_keys for bucket in compare_buckets)
            write_matrix_count(matrix_total_row, next_col, expected, actual,
                               matrix_formats["plain"]["total"])
        ws.merge_range(matrix_total_row + 2, 0, matrix_total_row + 2, last_shown_col,
                       "Қизил сон — янги жадвалдаги айириладиган миқдор: эталон − янги миқдор. "
                       "Қизил ёзув бўлмаса, янги миқдор 0. "
                       "CSVда белгиланмаган махсус тоифалар эталондаги қиймати билан кўрсатилди, солиштирилмади.",
                       fmt_note)
        missing_reference = [reference["names"].get(key, detailed_names.get(key, key))
                             for key in compare_keys if key not in reference["districts"]]
        missing_actual = [reference["names"].get(key, detailed_names.get(key, key))
                          for key in compare_keys if key not in detailed_names]
        if missing_reference or missing_actual:
            ws.merge_range(matrix_total_row + 3, 0, matrix_total_row + 4, last_shown_col,
                           "Эталонда йўқ: " + (", ".join(missing_reference) or "—") + ". "
                           "CSVда йўқ: " + (", ".join(missing_actual) or "—") + ".", fmt_warn)
        ws.set_column(0, 0, 31)
        ws.set_column(1, len(compare_buckets), 11)
        if last_shown_col > len(compare_buckets):
            ws.set_column(len(compare_buckets) + 1, last_shown_col, 13)
        ws.freeze_panes(2, 1)
        ws.set_landscape()
        ws.fit_to_pages(1, 1)
        ws.print_area(0, 0, matrix_total_row, last_shown_col)

        ws = wb.add_worksheet("Фарқлар")
        ws.set_tab_color("#9C6ADE")
        last_bucket_col = len(compare_buckets)
        total_col = last_bucket_col + 1
        status_col = total_col + 1
        last_bucket_letter = xlsxwriter.utility.xl_col_to_name(last_bucket_col)
        ws.merge_range(0, 0, 1, status_col, "ЭТАЛОН ВА ДАСТУР НАТИЖАСИНИ СОЛИШТИРИШ", fmt_title)
        ws.write(2, 0, "Эталон", fmt_label)
        ws.merge_range(2, 1, 2, status_col, reference["title"] or "Юкланган Excel", fmt_cell)
        ws.write(3, 0, "Дастур санаси", fmt_label)
        ws.write(3, 1, snapshot_date.strftime("%d.%m.%Y"), fmt_cell)
        ws.write(4, 0, "Фарқлар сони", fmt_label)
        ws.write_number(4, 1, comparison_mismatches, fmt_diff if comparison_mismatches else fmt_int)
        ws.write(4, 2, "Эталон жами", fmt_label)
        ws.write_number(4, 3, sum(reference["totals"].values()), fmt_int)
        ws.write(4, 4, "Дастур жами", fmt_label)
        ws.write_number(4, 5, sum(detailed_counts[key][bucket]
                                  for key in detailed_keys for bucket in compare_buckets), fmt_int)

        def section_value(section, key, bucket):
            expected = reference["districts"][key][bucket] if key in reference["districts"] else 0
            actual = detailed_counts[key][bucket] if key in detailed_names else 0
            if section == "reference":
                return expected
            if section == "actual":
                return actual
            return expected - actual

        section_rows = {}
        section_start = 6
        for section, title in (("reference", "ЭТАЛОН"), ("actual", "ДАСТУР"), ("difference", "ҚОЛДИҚ (ЭТАЛОН − ЯНГИ НАТИЖА)")):
            ws.merge_range(section_start, 0, section_start, status_col, title, fmt_label)
            header_row = section_start + 1
            data_start = section_start + 2
            ws.write_row(header_row, 0, ["Райгаз", *compare_buckets, "Жами", "Ҳолат"], fmt_header)
            ws.set_row(header_row, 26)
            section_rows[section] = data_start
            for offset, key in enumerate(compare_keys):
                row_index = data_start + offset
                name = reference["names"].get(key, detailed_names.get(key, key))
                has_reference = key in reference["districts"]
                has_actual = key in detailed_names
                ws.write(row_index, 0, name, fmt_cell)
                for col, bucket in enumerate(compare_buckets, start=1):
                    expected = reference["districts"][key][bucket] if has_reference else 0
                    actual = detailed_counts[key][bucket] if has_actual else 0
                    if section == "reference":
                        ws.write_number(row_index, col, expected, fmt_int if has_reference else fmt_unknown)
                    elif section == "actual":
                        ws.write_number(row_index, col, actual, fmt_int if has_actual else fmt_unknown)
                    else:
                        ref_row = section_rows["reference"] + offset + 1
                        actual_row = section_rows["actual"] + offset + 1
                        excel_col = xlsxwriter.utility.xl_col_to_name(col)
                        delta = expected - actual
                        ws.write_formula(row_index, col,
                                         f"={excel_col}{ref_row}-{excel_col}{actual_row}",
                                         fmt_diff if delta else fmt_int, delta)
                row_number = row_index + 1
                cached_total = sum(section_value(section, key, bucket) for bucket in compare_buckets)
                ws.write_formula(row_index, total_col, f"=SUM(B{row_number}:{last_bucket_letter}{row_number})",
                                 fmt_diff if section == "difference" and cached_total else fmt_int, cached_total)
                if section == "difference":
                    if not has_reference:
                        status = "Эталонда йўқ"
                    elif not has_actual:
                        status = "CSVда йўқ"
                    else:
                        status = "Мос" if all(detailed_counts[key][bucket] == reference["districts"][key][bucket]
                                               for bucket in compare_buckets) else "Фарқ бор"
                    ws.write(row_index, status_col, status, fmt_ok if status == "Мос" else fmt_warn)
            summary_row = data_start + len(compare_keys)
            ws.write(summary_row, 0, "ЖАМИ", fmt_total)
            for col in range(1, total_col + 1):
                excel_col = xlsxwriter.utility.xl_col_to_name(col)
                buckets_for_col = compare_buckets if col == total_col else (compare_buckets[col - 1],)
                cached = sum(section_value(section, key, bucket)
                             for key in compare_keys for bucket in buckets_for_col)
                ws.write_formula(summary_row, col,
                                 f"=SUM({excel_col}{data_start + 1}:{excel_col}{summary_row})", fmt_total,
                                 cached)
            section_start = summary_row + 3
        ws.set_column("A:A", 32)
        ws.set_column(1, total_col, 13)
        ws.set_column(status_col, status_col, 19)
        ws.freeze_panes(8, 1)
        ws.merge_range(section_start, 0, section_start, status_col,
                       "Қолдиқ = эталон − янги натижа. Эталонда ёки CSVда йўқ район 0 билан кўрсатилган ва Ҳолатда белгиланган. "
                       f"0 кунлик {zero_day_count} та ёзув 1-30 кунга қўшилмади.", fmt_note)
        ws.set_row(section_start, 34)

    # Райгаз
    ws = wb.add_worksheet("Райгаз")
    ws.set_tab_color("#5B9BD5")
    ray_labels = [x for x in INTERVALS if interval_counts[x] > 0 or x != "0-30 (НАЗОРАТ)"]
    ray_headers = ["Райгаз"] + ray_labels + ["Жами"]
    ws.write_row(0, 0, ray_headers, fmt_header)
    for r_idx, rg in enumerate(raygaz_sorted, start=1):
        ws.write(r_idx, 0, rg, fmt_cell)
        for c_idx, lab in enumerate(ray_labels, start=1):
            ws.write_number(r_idx, c_idx, raygaz_counts[rg][lab], fmt_int)
        ws.write_number(r_idx, len(ray_labels) + 1, sum(raygaz_counts[rg].values()), fmt_int)
    total_r = len(raygaz_sorted) + 1
    ws.write(total_r, 0, "ЖАМИ", fmt_total)
    for c_idx, lab in enumerate(ray_labels, start=1):
        ws.write_number(total_r, c_idx, interval_counts[lab], fmt_total)
    ws.write_number(total_r, len(ray_labels) + 1, total, fmt_total)
    ws.autofilter(0, 0, max(1, total_r - 1), len(ray_headers) - 1)
    ws.freeze_panes(1, 1)
    ws.set_column(0, 0, 30)
    ws.set_column(1, len(ray_headers)-1, 11)

    # Маҳалла
    ws = wb.add_worksheet("Маҳалла")
    ws.set_tab_color("#70AD47")
    mah_headers = ["Райгаз", "Маҳалла"] + ray_labels + ["Жами"]
    ws.write_row(0, 0, mah_headers, fmt_header)
    for r_idx, key in enumerate(mahalla_sorted, start=1):
        rg, mah = key
        ws.write(r_idx, 0, rg, fmt_cell)
        ws.write(r_idx, 1, mah, fmt_cell)
        for c_idx, lab in enumerate(ray_labels, start=2):
            ws.write_number(r_idx, c_idx, mahalla_counts[key][lab], fmt_int)
        ws.write_number(r_idx, len(ray_labels) + 2, sum(mahalla_counts[key].values()), fmt_int)
    total_m = len(mahalla_sorted) + 1
    ws.write(total_m, 0, "ЖАМИ", fmt_total)
    ws.write(total_m, 1, "", fmt_total)
    for c_idx, lab in enumerate(ray_labels, start=2):
        ws.write_number(total_m, c_idx, interval_counts[lab], fmt_total)
    ws.write_number(total_m, len(ray_labels) + 2, total, fmt_total)
    ws.autofilter(0, 0, max(1, total_m - 1), len(mah_headers)-1)
    ws.freeze_panes(1, 2)
    ws.set_column(0, 1, 30)
    ws.set_column(2, len(mah_headers)-1, 10)

    # Назорат
    ws = wb.add_worksheet("Назорат")
    ws.set_tab_color("#FFC000")
    ws.merge_range("A1:H2", "ФАЙЛЛАР БЎЙИЧА НАЗОРАТ ВА ТЕКШИРУВ", fmt_title)
    control_headers = ["Файл", "Қатор сони", "Амалий min", "Амалий max",
                       "Бўш майдон", "Файл ичи дубликат", "Ҳолат", "SHA-256"]
    ws.write_row(3, 0, control_headers, fmt_header)
    for r_idx, st in enumerate(file_stats, start=4):
        vals = [st["Файл"], st["Қатор сони"], st["Амалий min"], st["Амалий max"],
                st["Бўш майдон"], st["Файл ичи дубликат"], st["Ҳолат"], st["SHA-256"]]
        for c_idx, val in enumerate(vals):
            if c_idx in (1,2,3,4,5):
                ws.write_number(r_idx, c_idx, int(val), fmt_int)
            elif c_idx == 6:
                ws.write(r_idx, c_idx, val, fmt_ok)
            elif c_idx == 7:
                ws.write(r_idx, c_idx, val, fmt_hash)
            else:
                ws.write(r_idx, c_idx, val, fmt_cell)
    rr = 5 + len(file_stats)
    control_summary = [
        ("Жами ёзув", total),
        ("Жами қўшимча дубликат", duplicate_extra),
        ("Сана хатолари", 0),
        ("Snapshot сана", snapshot_date.isoformat()),
        ("Текширув хулосаси", "Тузилмавий хато аниқланмади"),
    ]
    for i, (k, v) in enumerate(control_summary):
        ws.write(rr+i, 0, k, fmt_label)
        if isinstance(v, int):
            ws.write_number(rr+i, 1, v, fmt_int)
        else:
            ws.write(rr+i, 1, v, fmt_ok if i == len(control_summary)-1 else fmt_cell)
    if warnings:
        wr = rr + len(control_summary) + 1
        ws.write(wr, 0, "Огоҳлантиришлар", fmt_label)
        ws.merge_range(wr, 1, wr + max(1, len(warnings)), 7, "\n".join("• "+x for x in warnings), fmt_warn)
    ws.set_column("A:A", 28)
    ws.set_column("B:G", 18)
    ws.set_column("H:H", 68)
    ws.freeze_panes(3, 0)

    # Дубликатлар — preserve, don't delete
    ws = wb.add_worksheet("Дубликатлар")
    ws.set_tab_color("#C00000")
    dup_headers = ["Абонент код", "Нечта", "Райгаз", "Маҳалла", "Абонент", "Манба файл", "Манба қатор"]
    ws.write_row(0, 0, dup_headers, fmt_header)
    dr = 1
    if duplicate_codes:
        for r in all_rows:
            if r["Абонент код"] in duplicate_codes:
                ws.write(dr, 0, r["Абонент код"], fmt_dup)
                ws.write_number(dr, 1, duplicate_codes[r["Абонент код"]], fmt_dup)
                ws.write(dr, 2, r["Райгаз"], fmt_dup)
                ws.write(dr, 3, r["Маҳалла"], fmt_dup)
                ws.write(dr, 4, r["Абонент"], fmt_dup)
                ws.write(dr, 5, r["Манба файл"], fmt_dup)
                ws.write_number(dr, 6, r["_source_row"], fmt_dup)
                dr += 1
    else:
        ws.write(1, 0, "Дубликат аниқланмади", fmt_ok)
    ws.set_column("A:B", 18)
    ws.set_column("C:F", 30)
    ws.set_column("G:G", 14)
    ws.freeze_panes(1, 0)

    # База: 8 official + 3 derived technical columns.
    ws = wb.add_worksheet("База")
    ws.set_tab_color("#A5A5A5")
    base_headers = REQUIRED_HEADERS + ["Кун фарқи", "Оралиқ", "Манба файл"]
    ws.write_row(0, 0, base_headers, fmt_header)
    for r_idx, r in enumerate(all_rows, start=1):
        ws.write(r_idx, 0, r["Вилоят"], fmt_cell)
        ws.write(r_idx, 1, r["Райгаз"], fmt_cell)
        ws.write(r_idx, 2, r["Маҳалла"], fmt_cell)
        ws.write_string(r_idx, 3, r["Абонент код"], fmt_center)
        ws.write(r_idx, 4, r["Абонент"], fmt_cell)
        ws.write(r_idx, 5, r["Еҳтиёж"], fmt_center)
        ws.write_datetime(r_idx, 6, r["Сўнги реализация"], fmt_date)
        ws.write_datetime(r_idx, 7, r["Бугунги реализация"], fmt_date)
        ws.write_number(r_idx, 8, r["Кун фарқи"], fmt_int)
        ws.write(r_idx, 9, r["Оралиқ"], fmt_center)
        ws.write(r_idx, 10, r["Манба файл"], fmt_cell)
    ws.autofilter(0, 0, total, 10)
    ws.freeze_panes(1, 0)
    ws.set_column("A:A", 20)
    ws.set_column("B:C", 28)
    ws.set_column("D:D", 16)
    ws.set_column("E:E", 36)
    ws.set_column("F:F", 12)
    ws.set_column("G:H", 16)
    ws.set_column("I:J", 14)
    ws.set_column("K:K", 26)

    wb.close()
    data = output.getvalue()

    summary = {
        "total": total,
        "raygaz": len(raygaz_names),
        "mahalla": len(mahalla_keys),
        "duplicates": duplicate_extra,
        "snapshot_date": snapshot_date.isoformat(),
        "files": len(file_stats),
        "comparison_mismatches": comparison_mismatches,
    }
    logging.info("OUTPUT | total=%s | raygaz=%s | mahalla=%s | duplicates=%s",
                 total, len(raygaz_names), len(mahalla_keys), duplicate_extra)
    return data, summary


@app.get("/")
def index():
    hosted = hosted_request()
    limit_mb = HOSTED_UPLOAD_MB if hosted else MAX_UPLOAD_MB
    return render_template_string(
        HTML,
        hosted=hosted,
        upload_limit_mb=limit_mb,
        upload_limit_bytes=limit_mb * 1024 * 1024,
        upload_hint=(f"Vercel: jami fayl hajmi {limit_mb} MB gacha. Katta fayllar uchun lokal dasturni ishlating."
                     if hosted else f"Lokal rejim: jami fayl hajmi {limit_mb} MB gacha."),
        privacy_notice=("Fayllar Vercel serverida qayta ishlanadi; doimiy saqlanmaydi."
                        if hosted else "lokal ishlaydi, fayllar tashqi serverga yuborilmaydi."),
    )


@app.get("/egaz")
def egaz_page():
    hosted = hosted_request()
    limit_mb = HOSTED_UPLOAD_MB if hosted else MAX_UPLOAD_MB
    return render_template_string(
        EGAZ_HTML,
        hosted=hosted,
        upload_limit_mb=limit_mb,
        upload_limit_bytes=limit_mb * 1024 * 1024,
        upload_hint=(f"Vercel: XLSX fayl {limit_mb} MB gacha."
                     if hosted else f"Lokal rejim: XLSX fayl {limit_mb} MB gacha."),
        privacy_notice=("Fayl Vercel serverida qayta ishlanadi; doimiy saqlanmaydi."
                        if hosted else "lokal ishlaydi, fayl tashqi serverga yuborilmaydi."),
    )


@app.post("/egaz/generate")
def generate_egaz():
    try:
        limit_mb = HOSTED_UPLOAD_MB if hosted_request() else MAX_UPLOAD_MB
        limit_bytes = limit_mb * 1024 * 1024
        if request.content_length and request.content_length > limit_bytes:
            return jsonify({"error": f"E-GAZ fayli {limit_mb} MB dan oshmasin."}), 413

        item = request.files.get("file")
        if item is None or not item.filename:
            raise ValueError("E-GAZ Billing XLSX faylini tanlang.")
        filename = Path(item.filename).name
        if not filename.lower().endswith(".xlsx"):
            raise ValueError("E-GAZ hisoboti uchun .xlsx fayl kerak.")
        data = item.read()
        if not data:
            raise ValueError(f"{filename}: fayl bo‘sh.")
        if len(data) > limit_bytes:
            return jsonify({"error": f"E-GAZ fayli {limit_mb} MB dan oshmasin."}), 413

        report = parse_gas_sales_xlsx(data, filename)
        if report is None:
            raise ValueError(
                f"{filename}: E-GAZ ustunlari topilmadi. "
                "Учреждение / Инспектор / Принял / Реализовал / Вернул ustunlari kerak."
            )
        xlsx_bytes, summary = build_egaz_xlsx(report, filename)
        if hosted_request() and len(xlsx_bytes) > HOSTED_UPLOAD_MB * 1024 * 1024:
            return jsonify({"error": "Tayyor Excel server yuklab olish limitidan oshdi."}), 413
        response = send_file(
            io.BytesIO(xlsx_bytes),
            as_attachment=True,
            download_name=summary["output_name"],
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response.headers["X-Output-Name"] = summary["output_name"]
        response.headers["X-Egaz-Rows"] = str(summary["rows"])
        response.headers["X-Egaz-Districts"] = str(summary["districts"])
        response.headers["X-Egaz-Accepted"] = str(summary["accepted"])
        response.headers["X-Egaz-Sold"] = str(summary["sold"])
        response.headers["X-Egaz-Unsold"] = str(summary["unsold"])
        return response
    except HTTPException:
        raise
    except ValueError as exc:
        logging.warning("E-GAZ request rejected: %s", exc)
        return jsonify({"error": str(exc)}), 400
    except Exception:
        logging.exception("Unexpected E-GAZ generate error")
        return jsonify({"error": "Kutilmagan server xatosi yuz berdi. Log faylini tekshiring."}), 500


@app.get("/mfy-svod")
def mfy_svod_page():
    return redirect("/gnp-taqqoslash", code=302)


@app.get("/gnp-taqqoslash")
def gnp_comparison_page():
    hosted = hosted_request()
    limit_mb = HOSTED_UPLOAD_MB if hosted else MAX_UPLOAD_MB
    return render_template_string(
        ANALYSIS_PAGE_HTML,
        active="gnp",
        mode="gnp",
        page_title="GNP + E-GAZ + MFY SVOD",
        page_description="Arizalar sana bo'yicha solishtiriladi, natija tuman va E-GAZ MFY kesimida ajratiladi. CSVda summa ustuni shart emas.",
        upload_limit_bytes=limit_mb * 1024 * 1024,
        upload_hint=f"Ikkala faylni tanlang. Umumiy hajmi {limit_mb} MB gacha.",
        privacy_notice=("Fayllar Vercel serverida qayta ishlanadi; doimiy saqlanmaydi."
                        if hosted else "Fayllar tashqi serverga yuborilmaydi."),
    )


@app.post("/mfy-svod/generate")
def generate_mfy_svod():
    try:
        limit_mb = HOSTED_UPLOAD_MB if hosted_request() else MAX_UPLOAD_MB
        limit_bytes = limit_mb * 1024 * 1024
        if request.content_length and request.content_length > limit_bytes:
            return jsonify({"error": f"CSV fayli {limit_mb} MB dan oshmasin."}), 413
        item = request.files.get("csv_file")
        if item is None or not item.filename:
            raise ValueError("E-GAZ CSV faylini tanlang.")
        filename = Path(item.filename).name
        if not filename.lower().endswith(".csv"):
            raise ValueError("MFY svodi uchun .csv fayl kerak.")
        data = item.read()
        if not data:
            raise ValueError(f"{filename}: fayl bo'sh.")
        if len(data) > limit_bytes:
            return jsonify({"error": f"CSV fayli {limit_mb} MB dan oshmasin."}), 413

        records = parse_egaz_sales_csv(data, filename)
        sales = aggregate_egaz_sales(records)
        xlsx_bytes, summary = build_mfy_svod_xlsx(sales, filename)
        if hosted_request() and len(xlsx_bytes) > HOSTED_UPLOAD_MB * 1024 * 1024:
            return jsonify({"error": "Tayyor Excel server yuklab olish limitidan oshdi."}), 413
        output_name = f"MFY_SVOD_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        response = send_file(
            io.BytesIO(xlsx_bytes), as_attachment=True, download_name=output_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response.headers["X-Output-Name"] = output_name
        response.headers["X-Source-Rows"] = str(summary["rows"])
        response.headers["X-Source-Amount"] = f"{summary['amount']:,.2f}"
        return response
    except HTTPException:
        raise
    except ValueError as exc:
        logging.warning("MFY summary request rejected: %s", exc)
        return jsonify({"error": str(exc)}), 400
    except Exception:
        logging.exception("Unexpected MFY summary error")
        return jsonify({"error": "Kutilmagan server xatosi yuz berdi. Log faylini tekshiring."}), 500


@app.post("/gnp-taqqoslash/generate")
def generate_gnp_comparison():
    try:
        limit_mb = HOSTED_UPLOAD_MB if hosted_request() else MAX_UPLOAD_MB
        limit_bytes = limit_mb * 1024 * 1024
        if request.content_length and request.content_length > limit_bytes:
            return jsonify({"error": f"Ikkala fayl hajmi jami {limit_mb} MB dan oshmasin."}), 413
        csv_items = [item for item in request.files.getlist("csv_file") if item and item.filename]
        gnp_item = request.files.get("gnp_file")
        if not csv_items:
            raise ValueError("Bitta yoki bir nechta E-GAZ CSV faylini tanlang.")
        if len(csv_items) > 500:
            raise ValueError("Bir urinishda 500 tadan ko'p E-GAZ faylini yuklamang.")
        if gnp_item is None or not gnp_item.filename:
            raise ValueError("GNP buyurtmalari XLSX faylini tanlang.")
        gnp_filename = Path(gnp_item.filename).name
        for item in csv_items:
            csv_upload_name = Path(item.filename).name
            if not csv_upload_name.lower().endswith((".csv", ".zip")):
                raise ValueError(f"{csv_upload_name}: E-GAZ sotuvlari uchun .csv yoki CSV fayllari joylangan .zip kerak.")
        if not gnp_filename.lower().endswith(".xlsx"):
            raise ValueError("GNP buyurtmalari uchun .xlsx fayl kerak.")
        gnp_data = gnp_item.read()
        if not gnp_data:
            raise ValueError("GNP XLSX fayli bo'sh.")
        records = []
        csv_filenames = []
        csv_digests = {}
        csv_bytes = 0
        for csv_item in source_uploads(csv_items, limit_bytes):
            csv_filename = Path(csv_item.filename).name
            if not csv_filename.lower().endswith(".csv"):
                raise ValueError(f"{csv_filename}: ZIP ichida faqat E-GAZ .csv fayllari bo'lishi kerak.")
            csv_data = csv_item.read()
            if not csv_data:
                raise ValueError(f"{csv_filename}: CSV fayli bo'sh.")
            digest = hashlib.sha256(csv_data).hexdigest()
            if digest in csv_digests:
                raise ValueError(
                    f"{csv_filename}: mazmuni '{csv_digests[digest]}' fayli bilan bir xil; "
                    "dublikat E-GAZ fayli ikki marta hisoblanmaydi."
                )
            csv_digests[digest] = csv_filename
            csv_bytes += len(csv_data)
            if csv_bytes + len(gnp_data) > limit_bytes:
                return jsonify({"error": f"CSV va XLSX fayllari ochilganda jami {limit_mb} MB dan oshmasin."}), 413
            file_records = parse_egaz_sales_csv(csv_data, csv_filename, require_amount=False)
            records.extend({**record, "source_file": csv_filename} for record in file_records)
            csv_filenames.append(csv_filename)
        if not records:
            raise ValueError("E-GAZ CSV fayllarida sotuv yozuvlari topilmadi.")
        csv_filename = ", ".join(csv_filenames)
        sales = aggregate_egaz_sales(records)
        orders = parse_gnp_orders_xlsx(gnp_data, gnp_filename)
        comparison = compare_egaz_with_gnp(sales, orders)
        xlsx_bytes, summary = build_gnp_comparison_xlsx(sales, comparison, csv_filename, gnp_filename)
        if hosted_request() and len(xlsx_bytes) > HOSTED_UPLOAD_MB * 1024 * 1024:
            return jsonify({"error": "Tayyor Excel server yuklab olish limitidan oshdi."}), 413
        output_name = f"GNP_EGAZ_MFY_SVOD_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        response = send_file(
            io.BytesIO(xlsx_bytes), as_attachment=True, download_name=output_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response.headers["X-Output-Name"] = output_name
        response.headers["X-Source-Rows"] = str(sales["sales"])
        response.headers["X-Mismatches"] = str(summary["mismatches"])
        response.headers["X-Matched"] = str(summary["matched"])
        response.headers["X-Same-Day"] = str(summary["same_day_found"])
        response.headers["X-Zero-Transfer"] = str(summary["zero_transfer_no_csv"])
        return response
    except HTTPException:
        raise
    except ValueError as exc:
        logging.warning("GNP comparison request rejected: %s", exc)
        return jsonify({"error": str(exc)}), 400
    except Exception:
        logging.exception("Unexpected GNP comparison error")
        return jsonify({"error": "Kutilmagan server xatosi yuz berdi. Log faylini tekshiring."}), 500


@app.get("/sotuvlar")
def sales_page():
    hosted = hosted_request()
    limit_mb = HOSTED_UPLOAD_MB if hosted else MAX_UPLOAD_MB
    return render_template_string(
        SALES_HTML,
        hosted=hosted,
        upload_limit_mb=limit_mb,
        upload_limit_bytes=limit_mb * 1024 * 1024,
        upload_hint=(f"Vercel: ikkala holat fayllari jami {limit_mb} MB gacha."
                     if hosted else f"Lokal rejim: ikkala holat fayllari jami {limit_mb} MB gacha."),
        privacy_notice=("Fayllar Vercel serverida qayta ishlanadi; doimiy saqlanmaydi."
                        if hosted else "lokal ishlaydi, fayllar tashqi serverga yuborilmaydi."),
    )


@app.post("/sotuvlar/compare")
def compare_sales():
    try:
        limit_mb = HOSTED_UPLOAD_MB if hosted_request() else MAX_UPLOAD_MB
        limit_bytes = limit_mb * 1024 * 1024
        if request.content_length and request.content_length > limit_bytes:
            return jsonify({"error": f"Ikkala holat fayllari jami {limit_mb} MB dan oshmasin."}), 413

        def buffer_uploads(file_items):
            buffered = []
            for item in source_uploads(file_items, limit_bytes):
                name = Path(item.filename or "noma'lum").name
                buffered.append(MemoryUpload(name, item.read()))
            return buffered

        def classify(buffered):
            gas_reports = []
            standard = []
            for item in buffered:
                if item.filename.lower().endswith(".xlsx"):
                    report = parse_gas_sales_xlsx(item.read(), item.filename)
                    if report is not None:
                        gas_reports.append(report)
                        continue
                standard.append(item)
            return gas_reports, standard

        old_gas, old_standard = classify(buffer_uploads(request.files.getlist("old_files")))
        new_gas, new_standard = classify(buffer_uploads(request.files.getlist("new_files")))
        if old_gas or new_gas:
            if not old_gas or not new_gas or old_standard or new_standard:
                raise ValueError(
                    "Solishtirishning ikkala tomoniga ham bir xil turdagi fayl yuklang: "
                    "faqat “Принял / Реализовал / Вернул” XLSX yoki faqat abonent holati fayllari."
                )
            ensure_unique_reports(old_gas, "Kechagi holat")
            ensure_unique_reports(new_gas, "Yangi holat")
            return jsonify(compare_gas_sales_reports(old_gas, new_gas))

        old_rows, _, _, _ = parse_files(old_standard)
        new_rows, _, _, _ = parse_files(new_standard)
        return jsonify(compare_sales_rows(old_rows, new_rows))
    except HTTPException:
        raise
    except ValueError as exc:
        logging.warning("Sales comparison request rejected: %s", exc)
        return jsonify({"error": str(exc)}), 400
    except Exception:
        logging.exception("Unexpected sales comparison error")
        return jsonify({"error": "Kutilmagan server xatosi yuz berdi. Log faylini tekshiring."}), 500


@app.post("/sotuvlar/compare-xlsx")
def compare_sales_xlsx():
    try:
        limit_mb = HOSTED_UPLOAD_MB if hosted_request() else MAX_UPLOAD_MB
        limit_bytes = limit_mb * 1024 * 1024
        if request.content_length and request.content_length > limit_bytes:
            return jsonify({"error": f"Ikkala holat fayllari jami {limit_mb} MB dan oshmasin."}), 413

        def load_reports(field_name):
            reports = []
            for item in source_uploads(request.files.getlist(field_name), limit_bytes):
                filename = Path(item.filename or "noma'lum").name
                if not filename.lower().endswith(".xlsx"):
                    raise ValueError(
                        "Rasmdagidek solishtirish Exceli uchun ikkala tomonga ham "
                        "“Принял / Реализовал / Вернул” formatidagi XLSX fayl yuklang."
                    )
                report = parse_gas_sales_xlsx(item.read(), filename)
                if report is None:
                    raise ValueError(
                        f"{filename}: “Учреждение / Инспектор / Принял / Реализовал / Вернул” "
                        "ustunlari topilmadi."
                    )
                reports.append(report)
            if not reports:
                raise ValueError("Solishtirish uchun ikkala tomonga ham XLSX fayl yuklang.")
            return reports

        old_reports = load_reports("old_files")
        new_reports = load_reports("new_files")
        ensure_unique_reports(old_reports, "Kechagi holat")
        ensure_unique_reports(new_reports, "Yangi holat")
        xlsx_bytes = build_gas_sales_comparison_xlsx(old_reports, new_reports)
        if hosted_request() and len(xlsx_bytes) > HOSTED_UPLOAD_MB * 1024 * 1024:
            return jsonify({"error": "Excel fayl server yuklab olish limitidan oshdi."}), 413
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"QABUL_SOTUV_SOLISHTIRISH_{stamp}.xlsx"
        resp = send_file(
            io.BytesIO(xlsx_bytes),
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        resp.headers["X-Output-Name"] = out_name
        return resp
    except HTTPException:
        raise
    except ValueError as exc:
        logging.warning("Gas sales Excel comparison request rejected: %s", exc)
        return jsonify({"error": str(exc)}), 400
    except Exception:
        logging.exception("Unexpected gas sales Excel comparison error")
        return jsonify({"error": "Kutilmagan server xatosi yuz berdi. Log faylini tekshiring."}), 500


@app.post("/generate")
def generate():
    try:
        if hosted_request() and request.content_length and request.content_length > HOSTED_UPLOAD_MB * 1024 * 1024:
            return jsonify({"error": f"Vercel limiti: jami fayl hajmi {HOSTED_UPLOAD_MB} MB dan oshmasin."}), 413
        files = request.files.getlist("files")
        expanded_limit = (HOSTED_UPLOAD_MB if hosted_request() else MAX_UPLOAD_MB) * 1024 * 1024
        buffered_uploads = []
        for item in source_uploads(files, expanded_limit):
            name = Path(item.filename or "noma'lum").name
            buffered_uploads.append(MemoryUpload(name, item.read()))

        gas_sales_reports = []
        standard_uploads = []
        for item in buffered_uploads:
            if item.filename.lower().endswith(".xlsx"):
                report = parse_gas_sales_xlsx(item.read(), item.filename)
                if report is not None:
                    gas_sales_reports.append(report)
                    continue
            standard_uploads.append(item)

        if gas_sales_reports and standard_uploads:
            raise ValueError(
                "Ikki xil hisobot turi birga tanlangan. “Принял / Реализовал / Вернул” XLSX "
                "fayllarini abonent holati CSV/XLSX fayllaridan alohida yuklang."
            )

        reference_file = request.files.get("reference")
        if gas_sales_reports:
            ensure_unique_reports(gas_sales_reports, "Hisobotlar")
            if reference_file and reference_file.filename:
                raise ValueError(
                    "Qabul va sotuv hisobotiga davomat etaloni qo‘shilmaydi. Etalonni olib tashlang."
                )
            xlsx_bytes, summary = build_gas_sales_xlsx(gas_sales_reports)
        else:
            all_rows, file_stats, warnings, manifest = parse_files(standard_uploads)
            reference = None
            if reference_file and reference_file.filename:
                reference_name = Path(reference_file.filename).name
                if not reference_name.lower().endswith(".xlsx"):
                    raise ValueError("Solishtirish uchun .xlsx fayl kerak.")
                reference = parse_reference_xlsx(reference_file.read(), header_key)
            xlsx_bytes, summary = build_xlsx(all_rows, file_stats, warnings, manifest, reference)

        if hosted_request() and len(xlsx_bytes) > HOSTED_UPLOAD_MB * 1024 * 1024:
            return jsonify({"error": "Excel fayl Vercel yuklab olish limitidan oshdi. Lokal dasturni ishlating."}), 413
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if summary.get("report_type") == "gas-sales":
            out_name = f"QABUL_SOTUV_{stamp}.xlsx"
        else:
            out_name = f"SVOD_{summary['snapshot_date']}_{stamp}.xlsx"
        resp = send_file(
            io.BytesIO(xlsx_bytes),
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        resp.headers["X-Output-Name"] = out_name
        resp.headers["X-Total-Rows"] = str(summary["total"])
        resp.headers["X-Raygaz-Count"] = str(summary["raygaz"])
        resp.headers["X-Report-Type"] = summary.get("report_type", "duration")
        if summary.get("report_type") == "gas-sales":
            resp.headers["X-Inspector-Count"] = str(summary["inspectors"])
            resp.headers["X-Accepted-Total"] = str(summary["accepted"])
            resp.headers["X-Sold-Total"] = str(summary["sold"])
            resp.headers["X-Returned-Total"] = str(summary["returned"])
            resp.headers["X-Sales-Percent"] = f"{summary['sales_percent']:.2f}"
        else:
            resp.headers["X-Mahalla-Count"] = str(summary["mahalla"])
            resp.headers["X-Duplicate-Count"] = str(summary["duplicates"])
            if summary["comparison_mismatches"] is not None:
                resp.headers["X-Mismatch-Count"] = str(summary["comparison_mismatches"])
        return resp
    except HTTPException:
        raise
    except ValueError as exc:
        logging.warning("Generate request rejected: %s", exc)
        return jsonify({"error": str(exc)}), 400
    except Exception:
        logging.exception("Unexpected generate error")
        return jsonify({"error": "Kutilmagan server xatosi yuz berdi. Log faylini tekshiring."}), 500


@app.errorhandler(413)
def too_large(_):
    limit_mb = HOSTED_UPLOAD_MB if hosted_request() else MAX_UPLOAD_MB
    return jsonify({"error": f"Yuklanayotgan fayllar juda katta. Limit: {limit_mb} MB."}), 413


def open_browser():
    webbrowser.open(f"http://127.0.0.1:{PORT}")


def port_is_busy(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        return s.connect_ex(("127.0.0.1", port)) == 0


def already_running(port: int) -> bool:
    """Portda aynan SVOD TIZIMI ishlayotganini tekshiradi."""
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/", timeout=3) as r:
            return APP_NAME.encode() in r.read(4096)
    except Exception:
        return False


if __name__ == "__main__":
    print(f"{APP_NAME} v{APP_VERSION}")
    if port_is_busy(PORT):
        if already_running(PORT):
            print("Dastur allaqachon ishlab turibdi. Brauzer ochilmoqda...")
            open_browser()
            sys.exit(0)
        print(f"XATO: {PORT}-port boshqa dastur tomonidan band. "
              "O'sha dasturni yoping yoki kompyuterni qayta ishga tushiring.")
        sys.exit(1)
    print(f"Brauzer: http://127.0.0.1:{PORT}")
    print("Dastur ishlayotgan paytda ushbu oynani yopmang.")
    logging.info("START | %s v%s | port=%s", APP_NAME, APP_VERSION, PORT)
    if os.environ.get("SVOD_NO_BROWSER") != "1":
        threading.Timer(1.2, open_browser).start()
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False)
