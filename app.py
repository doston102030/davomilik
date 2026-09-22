# -*- coding: utf-8 -*-
"""
SVOD TIZIMI v1.0
Mahalliy (local) CSV -> nazorat -> Excel svod tizimi.
Manba fayllar o'zgartirilmaydi va o'chirilmaydi.
"""
from __future__ import annotations

import csv
import hashlib
import io
import json
import logging
import os
import re
import socket
import sys
import tempfile
import threading
import urllib.request
import webbrowser
from collections import Counter, defaultdict
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import Dict, List, Tuple

from flask import Flask, jsonify, render_template_string, request, send_file
import xlsxwriter

APP_NAME = "SVOD TIZIMI"
APP_VERSION = "1.0"
PORT = 8765
MAX_UPLOAD_MB = 350

BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
logging.basicConfig(
    filename=LOG_DIR / "app.log",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8",
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

# Excel varag'idagi maksimal qatorlar soni (sarlavha bilan).
EXCEL_MAX_ROWS = 1_048_576

# Sarlavhalarni solishtirishda imlo farqlarini tekislash:
# "Эҳтиёж" / "Еҳтиёж", "Маҳалла" / "Махалла", "Сўнги" / "Сунги" va h.k.
HEADER_FOLD = str.maketrans({"э": "е", "ё": "е", "ҳ": "х", "қ": "к", "ғ": "г", "ў": "у"})

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


HTML = r"""
<!doctype html>
<html lang="uz">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SVOD TIZIMI</title>
<style>
:root {
  --navy:#17365D; --blue:#1F4E78; --green:#2E7D32; --light:#F5F8FC;
  --border:#D8E1EC; --danger:#B42318; --warn:#B54708;
}
*{box-sizing:border-box}
body{margin:0;background:var(--light);font-family:Segoe UI,Arial,sans-serif;color:#172B4D}
.top{background:linear-gradient(135deg,var(--navy),var(--blue));color:white;padding:24px 18px}
.top-inner{max-width:980px;margin:auto}
h1{margin:0 0 6px;font-size:30px}.sub{opacity:.9}
.wrap{max-width:980px;margin:26px auto;padding:0 18px}
.card{background:white;border:1px solid var(--border);border-radius:16px;padding:24px;box-shadow:0 8px 28px rgba(20,47,82,.08)}
.drop{border:2px dashed #87A6C8;border-radius:14px;padding:42px 20px;text-align:center;cursor:pointer;background:#FAFCFF;transition:.15s}
.drop.drag{border-color:var(--green);background:#F0FAF1}
.drop strong{font-size:20px;display:block;margin-bottom:8px}
.muted{color:#60758A;font-size:14px}
#fileInput{display:none}
.filelist{margin-top:18px;max-height:250px;overflow:auto;border-top:1px solid var(--border)}
.file{display:flex;justify-content:space-between;gap:12px;padding:10px 4px;border-bottom:1px solid #EDF1F6;font-size:14px}
.actions{display:flex;gap:12px;align-items:center;margin-top:20px;flex-wrap:wrap}
button{border:0;border-radius:10px;padding:13px 20px;font-size:16px;font-weight:700;cursor:pointer}
.primary{background:var(--green);color:white}.primary:disabled{opacity:.45;cursor:not-allowed}
.secondary{background:#E9EFF7;color:#24496F}
.status{margin-top:18px;padding:14px 16px;border-radius:10px;display:none;white-space:pre-wrap}
.status.ok{display:block;background:#EAF6EC;color:#1B5E20;border:1px solid #B9DDBE}
.status.err{display:block;background:#FFF0EE;color:var(--danger);border:1px solid #F4C7C3}
.status.work{display:block;background:#FFF7E8;color:var(--warn);border:1px solid #F2D59D}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:18px}
.kpi{background:#F7FAFD;border:1px solid var(--border);border-radius:12px;padding:14px}
.kpi b{font-size:23px;display:block;color:var(--navy)}
.notice{margin-top:18px;padding:14px;background:#F2F7FF;border-left:4px solid var(--blue);border-radius:8px;font-size:14px;line-height:1.5}
.footer{color:#73879C;font-size:12px;text-align:center;margin:18px}
@media(max-width:720px){.grid{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<div class="top"><div class="top-inner">
  <h1>SVOD TIZIMI</h1>
  <div class="sub">CSV fayllarni tekshirish va tayyor Excel svod yaratish</div>
</div></div>
<div class="wrap">
  <div class="card">
    <div id="drop" class="drop">
      <strong>CSV fayllarni shu yerga tashlang</strong>
      <div class="muted">yoki fayllarni tanlash uchun shu maydonni bosing</div>
      <input id="fileInput" type="file" accept=".csv,text/csv" multiple>
    </div>

    <div id="filelist" class="filelist"></div>

    <div class="actions">
      <button id="go" class="primary" disabled>Tekshirish va Excel yaratish</button>
      <button id="clear" class="secondary">Tozalash</button>
      <span class="muted">Manba CSV fayllar o'zgartirilmaydi.</span>
    </div>

    <div id="status" class="status"></div>
    <div id="kpis" class="grid" style="display:none"></div>

    <div class="notice">
      <b>Qat'iy nazorat:</b> kerakli ustun yo'q bo'lsa, sana noto'g'ri bo'lsa yoki
      “Бугунги реализация” sanalari bir xil bo'lmasa tizim xatoni ko'rsatadi va
      noto'g'ri svod chiqarmaydi. Dublikatlar o'chirilmaydi — alohida varaqda ko'rsatiladi.
      Barcha hisoblar sanalarning haqiqiy farqidan olinadi, fayl nomiga ishonilmaydi.
    </div>
  </div>
  <div class="footer">SVOD TIZIMI v1.0 — lokal ishlaydi, fayllar tashqi serverga yuborilmaydi.</div>
</div>

<script>
let files = [];
const drop = document.getElementById('drop');
const inp = document.getElementById('fileInput');
const list = document.getElementById('filelist');
const go = document.getElementById('go');
const clearBtn = document.getElementById('clear');
const status = document.getElementById('status');
const kpis = document.getElementById('kpis');

// Fayl maydondan tashqariga tashlansa, brauzer uni ochib sahifani yo'qotmasin.
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('drop', e => e.preventDefault());

drop.onclick = () => inp.click();
drop.ondragover = e => { e.preventDefault(); drop.classList.add('drag'); };
drop.ondragleave = () => drop.classList.remove('drag');
drop.ondrop = e => {
  e.preventDefault(); drop.classList.remove('drag');
  addFiles([...e.dataTransfer.files]);
};
inp.onchange = () => { addFiles([...inp.files]); inp.value = ''; };

function addFiles(newFiles){
  for(const f of newFiles){
    if(!f.name.toLowerCase().endsWith('.csv')) continue;
    const key = f.name + ':' + f.size + ':' + f.lastModified;
    if(!files.some(x => x._key === key)){ f._key = key; files.push(f); }
  }
  render();
}
function render(){
  list.innerHTML = '';
  files.forEach((f,i)=>{
    const d=document.createElement('div'); d.className='file';
    d.innerHTML=`<span>${i+1}. ${escapeHtml(f.name)}</span><span>${(f.size/1024).toFixed(1)} KB</span>`;
    list.appendChild(d);
  });
  go.disabled = files.length===0;
}
function escapeHtml(s){return s.replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]));}
clearBtn.onclick=()=>{files=[];inp.value='';render();status.className='status';status.textContent='';kpis.style.display='none';};
go.onclick=async()=>{
  if(!files.length)return;
  go.disabled=true;
  status.className='status work';
  const t0=Date.now();
  const tick=()=>{status.textContent='Tekshirilmoqda va Excel tayyorlanmoqda... '
    +Math.round((Date.now()-t0)/1000)+' s\n(katta fayllarda 1-2 daqiqa davom etishi mumkin)';};
  tick();
  const timer=setInterval(tick,1000);
  kpis.style.display='none';
  const fd=new FormData();
  files.forEach(f=>fd.append('files',f,f.name));
  try{
    let res;
    try{
      res=await fetch('/generate',{method:'POST',body:fd});
    }catch(e){
      throw new Error("Server bilan aloqa yo'q. ISHGA_TUSHIRISH.bat oynasi ochiq ekanini tekshiring va sahifani yangilang.");
    }
    if(!res.ok){
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

    const total=res.headers.get('X-Total-Rows')||'-';
    const ray=res.headers.get('X-Raygaz-Count')||'-';
    const mah=res.headers.get('X-Mahalla-Count')||'-';
    const dup=res.headers.get('X-Duplicate-Count')||'-';
    status.className='status ok';
    status.textContent='Tayyor. Excel fayl yaratildi va yuklandi.';
    kpis.innerHTML=`
      <div class="kpi"><span>Jami yozuv</span><b>${total}</b></div>
      <div class="kpi"><span>Raygaz</span><b>${ray}</b></div>
      <div class="kpi"><span>Mahalla</span><b>${mah}</b></div>
      <div class="kpi"><span>Dublikat</span><b>${dup}</b></div>`;
    kpis.style.display='grid';
  }catch(e){
    status.className='status err';
    status.textContent=e.message;
  }finally{
    clearInterval(timer);
    go.disabled=files.length===0;
  }
};
</script>
</body>
</html>
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
            return data.decode(enc)
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


def parse_files(file_items) -> Tuple[List[dict], List[dict], List[str], List[dict]]:
    all_rows: List[dict] = []
    file_stats: List[dict] = []
    warnings: List[str] = []
    source_manifest: List[dict] = []
    today_values = set()

    if not file_items:
        raise ValueError("Kamida bitta CSV fayl tanlang.")

    for item in file_items:
        filename = Path(item.filename or "noma'lum.csv").name
        if not filename.lower().endswith(".csv"):
            raise ValueError(f"{filename}: faqat .csv fayl qabul qilinadi.")

        data = item.read()
        if not data:
            raise ValueError(f"{filename}: fayl bo'sh.")
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

        digest = sha256_hex(data)
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

    # A single reporting/snapshot date is required to prevent mixing two days.
    if len(today_values) != 1:
        vals = ", ".join(sorted(d.isoformat() for d in today_values))
        raise ValueError(
            "“Бугунги реализация” sanasi barcha fayllarda bir xil emas. "
            f"Aniqlangan sanalar: {vals}. Turli kun ma'lumotlarini bitta svodga aralashtirish bloklandi."
        )

    return all_rows, file_stats, warnings, source_manifest


def build_xlsx(all_rows: List[dict], file_stats: List[dict], warnings: List[str],
               source_manifest: List[dict]) -> Tuple[bytes, dict]:
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
    }
    logging.info("OUTPUT | total=%s | raygaz=%s | mahalla=%s | duplicates=%s",
                 total, len(raygaz_names), len(mahalla_keys), duplicate_extra)
    return data, summary


@app.get("/")
def index():
    return render_template_string(HTML)


@app.post("/generate")
def generate():
    try:
        files = request.files.getlist("files")
        all_rows, file_stats, warnings, manifest = parse_files(files)
        xlsx_bytes, summary = build_xlsx(all_rows, file_stats, warnings, manifest)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
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
        resp.headers["X-Mahalla-Count"] = str(summary["mahalla"])
        resp.headers["X-Duplicate-Count"] = str(summary["duplicates"])
        return resp
    except Exception as exc:
        logging.exception("Generate error")
        return jsonify({"error": str(exc)}), 400


@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": f"Yuklanayotgan fayllar juda katta. Limit: {MAX_UPLOAD_MB} MB."}), 413


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
    threading.Timer(1.2, open_browser).start()
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False)
