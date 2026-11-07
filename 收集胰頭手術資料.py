# -*- coding: utf-8 -*-
"""
胰臟頭部手術（Whipple / PD）術前資料收集 —— 交班單專案的分支工具

連線用同資料夾的 eip_session.py（從交班單專案抽出來的 EipSession 與
open_patient）：登入 EIP 一次，開 EMR 病歷首頁（rkeprd 開不起來會換 web9），
之後每位病人依序抓：

  1. findOpn 手術記錄清單（範例一）
       找最近一次胰臟頭部手術（Whipple / PD / PPPD / RPD …，見 PANC_HEAD_NAME_RE）
  2. findOpn&dt=0YYYMMDD&tm=N 手術記錄內容（範例二）
       術前診斷、手術名稱、手術開始時間
  3. findResd&resdtmonth=月數&resdtype=報告類型 累積報告
       手術開始前最新一筆的各項檢驗。要查哪些項目可以自選（候選清單見
       lab_items.py），預設 ALB / CREA / BILIT：
         DCHEM 生化（範例三，31 項）　DCBC 血液 CBC/DC 與凝血（範例五，21 項）
       resdtmonth = 手術至今的月數 ＋ 往前多查的月數（EXTRA_MONTHS，預設 3），
       一份報告一頁就涵蓋全部；比這範圍更早的檢驗不採用。
       同一份報告的項目在同一頁，多勾幾項不會增加查詢次數；
       兩份報告都有勾，每位就多開一頁。
  4. findVts&caseno=all&pbvtype=HWS 生命象徵（範例四，可關掉）
       手術開始前最新一筆「身高、體重都有值」的量測，用它算 BMI
  5. EMR 首頁表頭的生日 -> 手術當天的年齡

最後輸出成 Excel：
  ‧「術前資料」工作表：病歷號、術前診斷、手術日期、手術名稱、手術當下年齡、
     術前身高、術前體重、術前BMI、各項術前檢驗、備註
  ‧「資料時間（核對用）」工作表：每個數值是哪一天量的，方便抽查

用法（整個資料夾放在一起即可，不需要交班單程式）：
    雙擊「執行.bat」 -> 開啟操作介面（收集介面.py），勾選檢驗項目後執行
命令列：
    py -3 收集胰頭手術資料.py                      讀同資料夾的 病歷號清單.txt
    py -3 收集胰頭手術資料.py 我的清單.txt
    py -3 收集胰頭手術資料.py --histno 1234567 --histno 12345678
    py -3 收集胰頭手術資料.py --save-html          另存每個網頁，解析不對時給我看
    py -3 收集胰頭手術資料.py --visible            開有畫面的 Chrome
    py -3 收集胰頭手術資料.py --delay 6            放慢：每開一頁至少隔 6 秒（預設 3）
    py -3 收集胰頭手術資料.py --restart            不續跑，全部重查
    py -3 收集胰頭手術資料.py --labs ALB,CREA,BILIT,HGB,PLT   自選檢驗項目
    py -3 收集胰頭手術資料.py --no-vitals          不查身高體重

速度與續跑：每開一頁至少間隔 eip_session.REQUEST_INTERVAL 秒，每位病人之間
再停 PATIENT_DELAY 秒。連續 MAX_CONSECUTIVE_FAIL 位打不開 EMR（或伺服器回
403/429）就當作被鎖、自動停止。每查完一位就存進 收集進度.json，中斷後再
執行會跳過已查完的病人；全部查完會自動刪掉這個檔。

病歷號清單.txt：一行一個病歷號（逗號、空白分隔也可以），# 開頭是註解。

第一次使用：
    py -3 install_requirement.py          安裝 selenium、requests、openpyxl
    EIP 帳密填在 eip帳號.txt（沒填的話第一次執行會詢問並自動存檔）
輸出的 Excel 與 --save-html 存下來的網頁都含病人資料，用完請妥善保管或刪除。
"""

from __future__ import annotations

import argparse
import html as _html
import os
import re
import sys
import time
import traceback
from datetime import date, datetime, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

DEFAULT_LIST = os.path.join(BASE_DIR, "病歷號清單.txt")
from eip_session import EMR_PATH  # noqa: E402

import lab_items  # noqa: E402

LAB_ITEMS = lab_items.DEFAULT_LABS      # 沒指定時查這幾項

# ==================== 胰臟頭部手術的判定 ====================
# 手術名稱（不分大小寫）含下列字樣就算。要多收別的術式，加進這裡即可。
PANC_HEAD_NAME_RE = re.compile(
    r"whipple"
    r"|pancreat(?:ic)?o?duodenectom"      # pancreaticoduodenectomy / pancreatoduodenectomy
    r"|胰十二指腸",
    re.I)
# 縮寫要大寫、且前後不接英文字母（避免誤中 PDS、PDA、Updated 之類）
PANC_HEAD_ABBR_RE = re.compile(
    r"(?<![A-Za-z])(?:PD|PPPD|SSPPD|PRPD|RPD|LPD|RAPD|LAPD)(?![A-Za-z])")
# 只用縮寫命中時，名稱裡若有這些字就不算（例：PD catheter 是腹膜透析導管）
PANC_HEAD_ABBR_EXCLUDE_RE = re.compile(
    r"catheter|dialysis|tenckhoff|peritoneal|透析", re.I)


def is_panc_head(name: str) -> bool:
    n = name or ""
    if PANC_HEAD_NAME_RE.search(n):
        return True
    return bool(PANC_HEAD_ABBR_RE.search(n)) and not PANC_HEAD_ABBR_EXCLUDE_RE.search(n)


# ============================ 小工具 ============================

_TAG_RE = re.compile(r"<[^>]+>")


def clean(s) -> str:
    """去標籤、解 HTML 實體、&nbsp; 換空白、壓縮空白。"""
    t = _TAG_RE.sub(" ", str(s or ""))
    t = _html.unescape(t).replace("\xa0", " ")
    return re.sub(r"\s+", " ", t).strip()


def roc_dt(d: date) -> str:
    """2026-10-20 -> '01151020'（findOpn 的 dt 參數：0 + 民國年3碼 + 月日）"""
    return f"{d.year - 1911:04d}{d.month:02d}{d.day:02d}"


def to_number(v):
    """'0.42' -> 0.42；'<5' 之類保留原字串。"""
    s = str(v or "").strip()
    try:
        return float(s)
    except ValueError:
        return s


def age_at(birth: date | None, d: date | None):
    if not (birth and d):
        return None
    return d.year - birth.year - ((d.month, d.day) < (birth.month, birth.day))


def read_text_any(path: str) -> str:
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-16", "cp950", "big5hkscs"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", errors="replace")


def read_histno_list(path: str):
    return parse_histno_text(read_text_any(path))


def parse_histno_text(text: str):
    """一行一個病歷號（逗號、空白、頓號分隔也可以），# 開頭是註解，去重複。"""
    out = []
    for line in (text or "").splitlines():
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        for tok in re.split(r"[\s,，、;]+", s):
            if tok and re.fullmatch(r"\w+", tok) and tok not in out:
                out.append(tok)
    return out


# ==================== 解析：EMR 首頁表頭 -> 生日 ====================

def parse_birthday(page: str):
    """表頭：王小明 12345678 67歲(19590101) 男 A000- 00(住院中) -> date(1959,1,1)"""
    m = re.search(r"歲\s*\(\s*(\d{8})\s*\)", clean(page))
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y%m%d").date()
    except ValueError:
        return None


# ==================== 解析：手術記錄清單（範例一） ====================

_OPN_LINK_RE = re.compile(
    r"action=findOpn&(?:amp;)?histno=\w+&(?:amp;)?dt=(\d+)&(?:amp;)?tm=(\d+)[^>]*>\s*(\d{8})\s*<",
    re.I)
_TD_RE = re.compile(r"<td[^>]*>(.*?)</td>", re.S | re.I)


def _op_names(cell: str):
    """'1.SURDGS  02123&nbsp;&nbsp;Robot-assisted Whipple's operation<br>2.  &nbsp;&nbsp;'
    -> ["Robot-assisted Whipple's operation"]"""
    names = []
    for seg in re.split(r"<br\s*/?>", cell, flags=re.I):
        parts = re.split(r"(?:&nbsp;|&#160;|\xa0){2,}", seg, maxsplit=1)
        name = clean(parts[1]) if len(parts) > 1 else re.sub(r"^\d+\.\s*", "", clean(seg))
        if name:
            names.append(name)
    return names


def parse_op_list(page: str):
    """回傳 [{date, dt, tm, names, surgeon}]，順序照網頁。"""
    ops = []
    for chunk in re.split(r"<tr[^>]*>", page or "", flags=re.I):
        link = _OPN_LINK_RE.search(chunk)
        if not link:
            continue
        cells = _TD_RE.findall(chunk)
        try:
            d = datetime.strptime(link.group(3), "%Y%m%d").date()
        except ValueError:
            continue
        ops.append({
            "date": d,
            "dt": link.group(1),
            "tm": link.group(2),
            "names": _op_names(cells[1]) if len(cells) > 1 else [],
            "surgeon": clean(cells[2]) if len(cells) > 2 else "",
        })
    return ops


def pick_panc_head(ops):
    """最近一次胰臟頭部手術；沒有回 None。"""
    hits = [op for op in ops if any(is_panc_head(n) for n in op["names"])]
    if not hits:
        return None
    return max(hits, key=lambda op: (op["date"], int(op["tm"] or 0)))


# ==================== 解析：手術記錄內容（範例二） ====================

_FIELD_RE = re.compile(r"<font[^>]*>\s*([^<>：:]{1,40})[：:]\s*</font>(.*?)</td>", re.S | re.I)
_OP_NAME_KEYS = ("手術名稱一", "手術名稱二", "手術名稱三", "手術名稱四", "手術名稱五")


def parse_op_detail(page: str) -> dict:
    """回傳 {'術前診斷':..., '手術名稱': [...], 'start': datetime|None, 'fields': {...}}"""
    fields = {}
    for label, value in _FIELD_RE.findall(page or ""):
        label = clean(label)
        if label not in fields:
            v = clean(value)
            fields[label] = "" if re.fullmatch(r"-+", v) else v

    start = None
    m = re.search(r"(\d{8})\s*(\d{4})?", fields.get("手術開始日期", ""))
    if m:
        try:
            start = datetime.strptime(m.group(1) + (m.group(2) or "0000"), "%Y%m%d%H%M")
        except ValueError:
            start = None
    return {
        "術前診斷": fields.get("術前診斷", ""),
        "手術名稱": [fields[k] for k in _OP_NAME_KEYS if fields.get(k)],
        "start": start,
        "has_start_time": bool(m and m.group(2)),
        "fields": fields,
    }


# ==================== 解析：生化累積報告（範例三） ====================

def _parse_lab_stamp(s: str):
    s = s.strip()
    for fmt in ("%y-%m-%d %H:%M", "%y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S",
                "%y-%m-%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _visible_lab_table(page: str):
    """備援：從看得到的表格讀 {項目: [(時間, 值)]}。"""
    head = re.search(r"<thead[^>]*>(.*?)</thead>", page or "", re.S | re.I)
    if not head:
        return {}
    cols = []
    for c in re.findall(r"<th[^>]*>(.*?)</th>", head.group(1), re.S | re.I):
        code = re.search(r"getImg\('([A-Za-z0-9_]+)'\)", c)      # 表頭顯示 eGFR(M)，代碼是 EGFR
        cols.append(code.group(1) if code else clean(c))
    out = {}
    for row in re.findall(r"<tr[^>]*class=\"?resdtr\"?[^>]*>(.*?)</tr>", page, re.S | re.I):
        cells = [clean(c) for c in _TD_RE.findall(row)]
        if not cells:
            continue
        t = _parse_lab_stamp(cells[0])
        if not t:
            continue
        for name, val in zip(cols[1:], cells[1:]):
            out.setdefault(name, []).append((t, val))
    return out


def parse_lab_series(page: str, item: str):
    """[(時間, 原始字串)]，含空值。優先讀隱藏列 <td id="ALB">26-10-16 08:02/2.8|...</td>。"""
    m = re.search(r'<td[^>]*\bid="%s"[^>]*>(.*?)</td>' % re.escape(item), page or "", re.S | re.I)
    if m:
        series = []
        for ent in _html.unescape(_TAG_RE.sub("", m.group(1))).split("|"):
            if "/" not in ent:
                continue
            stamp, val = ent.split("/", 1)
            t = _parse_lab_stamp(stamp)
            if t:
                series.append((t, val.strip()))
        return series
    return _visible_lab_table(page).get(item, [])


def latest_lab_before(series, cutoff: datetime, since: datetime | None = None):
    """cutoff（含）以前最新一筆有值的檢驗，回傳 (時間, 值) 或 (None, None)。

    since：比這更早的不算（依月數查的報告用，跟網頁的查詢範圍一致）。
    """
    ok = [(t, v) for t, v in series
          if t <= cutoff and (since is None or t >= since) and v and v.strip() not in ("-", "")]
    if not ok:
        return None, None
    t, v = max(ok, key=lambda x: x[0])
    return t, to_number(v)


# ==================== 解析：身高體重（範例四） ====================

def _vital_hidden(page: str, name: str):
    tag = re.search(r'<input[^>]*\bname="%s"[^>]*>' % name, page or "", re.I)
    if not tag:
        return None
    val = re.search(r'\bvalue="([^"]*)"', tag.group(0))
    if not val:
        return None
    out = []
    for ent in _html.unescape(val.group(1)).split("|"):
        if "/" not in ent:
            continue
        stamp, v = ent.split("/", 1)
        try:
            t = datetime.strptime(stamp.strip(), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
        out.append((t, v.strip()))
    return out


def _vital_table(page: str):
    """備援：從看得到的表格讀 [(時間, 身高, 體重)]。"""
    rows = []
    for row in re.findall(r"<tr[^>]*>(.*?)</tr>", page or "", re.S | re.I):
        cells = [clean(c) for c in _TD_RE.findall(row)]
        if len(cells) < 3:
            continue
        try:
            t = datetime.strptime(re.sub(r"\s+", " ", cells[0]), "%Y-%m-%d %H:%M")
        except ValueError:
            continue
        rows.append((t, re.sub(r"[^\d.]", "", cells[1]), re.sub(r"[^\d.]", "", cells[2])))
    return rows


def parse_vitals(page: str):
    """[(時間, 身高cm, 體重kg)]，讀不到的值是 None。"""
    hs, ws = _vital_hidden(page, "Height"), _vital_hidden(page, "Weight")
    if hs is not None and ws is not None and len(hs) == len(ws):
        raw = [(th, h, w) for (th, h), (_, w) in zip(hs, ws)]
    else:
        raw = _vital_table(page)

    def num(x):
        try:
            v = float(x)
        except (TypeError, ValueError):
            return None
        return v if v > 0 else None

    return [(t, num(h), num(w)) for t, h, w in raw]


def latest_vitals_before(rows, cutoff: datetime):
    """cutoff（含）以前最新一筆身高、體重都有值的量測 -> (時間, 身高, 體重)。"""
    ok = [r for r in rows if r[0] <= cutoff and r[1] and r[2]]
    if not ok:
        return None, None, None
    return max(ok, key=lambda r: r[0])


def bmi(height_cm, weight_kg):
    if not (height_cm and weight_kg):
        return None
    return round(weight_kg / (height_cm / 100) ** 2, 2)


# ============================ 連線 ============================

class BlockedError(RuntimeError):
    """伺服器明確拒絕（HTTP 403 / 429），很可能是被鎖了，整批要停下來。"""


# driver.get() 本身會等網頁載完，這裡只是保險多等一下內容出現。
# 等不到通常是這位病人沒有這類資料（例如那年沒驗生化），不必等太久。
MARKER_WAIT = 5.0

# （依年份查的報告才用）手術年份那頁查不到的檢驗，只有手術在年初這幾個月內才去翻前一年
# （術前抽血多半在幾週內；這樣不常驗的項目不會讓每位都多查一頁）。
PREV_YEAR_MONTHS = 3

# 依「月數」查的報告（目前生化、血液都是）：resdtmonth 填「手術日到今天經過幾個月」，
# 再往前多查這麼多個月，才涵蓋得到手術前幾週抽的血。超過這個範圍的舊資料
# 不會被抓到（等於術前檢驗只採手術前這幾個月內的）。介面上可以改。
EXTRA_MONTHS = 3


def months_since(d: date, today: date | None = None) -> int:
    """d 到今天至少要往回幾個月才涵蓋得到（不足一個月算一個月）。

    今天 11/07、手術 10/20 -> 1（11/07 往回 1 個月是 10/07，已涵蓋 10/20）
    今天 11/25、手術 10/20 -> 2（往回 1 個月是 10/25，還不到 10/20）
    """
    today = today or date.today()
    n = (today.year - d.year) * 12 + (today.month - d.month)
    if today.day > d.day:
        n += 1
    return max(n, 0)


def months_before(today: date, n: int) -> datetime:
    """today 往回 n 個月的那天 00:00（月底會自動對齊，例 3/31 往回 1 個月 -> 2/28）。"""
    import calendar
    y, m = divmod(today.year * 12 + today.month - 1 - n, 12)
    d = min(today.day, calendar.monthrange(y, m + 1)[1])
    return datetime(y, m + 1, d)


def report_periods(rtype, cutoff: datetime, extra_months=None, today=None):
    """這份報告要開哪幾頁：回傳 resdtmonth 參數值的 list。

    依月數的報告（目前全部）：[手術至今的月數 + 往前多查的月數]，一頁就涵蓋全部
    依年份的報告（保留的舊做法）：[手術年份]，手術在年初再加 [前一年]
    """
    if lab_items.PERIOD.get(rtype) == "months":
        extra = EXTRA_MONTHS if extra_months is None else max(0, int(extra_months))
        return [max(1, months_since(cutoff.date(), today) + extra)]
    if cutoff.month <= PREV_YEAR_MONTHS:
        return [cutoff.year, cutoff.year - 1]
    return [cutoff.year]


def get_page(eip, url, marker, log=print, wait=None) -> str:
    """瀏覽器整頁導覽，等 marker（regex）出現。每次開網頁前都會先 throttle()。

    網頁有開起來但沒看到 marker，通常代表這位病人沒有這類資料（例如當年
    沒驗生化），直接回傳內容，不再多打一次——避免無謂增加請求次數。
    只有瀏覽器本身開啟失敗時，才改用 requests 帶 AJAX 標頭再試一次。
    """
    import eip_session

    pat = re.compile(marker, re.I)
    wait = MARKER_WAIT if wait is None else wait
    html_text = ""
    try:
        eip_session.throttle()
        eip.driver.get(url)
        deadline = time.time() + wait
        while True:
            html_text = eip.driver.page_source or ""
            if pat.search(html_text):
                return html_text
            if time.time() > deadline:
                return html_text
            time.sleep(0.4)
    except Exception as e:
        log(f"      [瀏覽器] 開啟失敗（{e}），改用 requests")

    try:
        eip.sync_cookies()
        eip_session.throttle()
        r = eip.sess.get(url, timeout=20, headers={
            "X-Requested-With": "XMLHttpRequest",
            "Accept": "text/html, */*; q=0.01",
        })
        if r.status_code in (403, 429):
            raise BlockedError(f"伺服器回 HTTP {r.status_code}")
        r.raise_for_status()
        return r.text
    except BlockedError:
        raise
    except Exception as e:
        log(f"      [requests] 失敗（{e}）")
    return html_text


def _save(save_dir, histno, tag, text):
    if not (save_dir and text):
        return
    os.makedirs(save_dir, exist_ok=True)
    with open(os.path.join(save_dir, f"{histno}_{tag}.html"), "w", encoding="utf-8") as f:
        f.write(text)


# ============================ 主流程：一位病人 ============================

def collect_one(histno, eip, log=print, save_dir=None, labs=LAB_ITEMS, vitals=True,
                extra_months=None) -> dict:
    """查一位病人，回傳一列結果。任何一步失敗都只把原因寫進「備註」，不中斷。

    labs：要查的檢驗（'DCHEM:ALB' 或只寫 'ALB'，見 lab_items.py）；
    vitals=False 就不查身高體重。
    """
    import eip_session

    labs = lab_items.normalize(labs)
    # _labs / _vitals 記下這一列查過哪些東西，續跑時用來判斷要不要重查
    row = {"病歷號": histno, "備註": [], "_labs": list(labs), "_vitals": bool(vitals)}
    note = row["備註"].append

    # --- EMR 首頁：生日、決定用哪個網域 ---
    emr_html, host = eip_session.open_patient(histno, eip, log=log)
    _save(save_dir, histno, "0_findEmr", emr_html)
    if not host:
        note("EMR 病歷首頁打不開")
        return row
    birth = parse_birthday(emr_html)
    row["生日"] = birth
    if not birth:
        note("表頭讀不到生日")

    base = f"{host}{EMR_PATH}"

    # --- 1. 手術記錄清單 ---
    page = get_page(eip, f"{base}?action=findOpn&histno={histno}",
                    r"手術記錄清單|opnlist|action=findOpn&(?:amp;)?histno", log=log)
    _save(save_dir, histno, "1_findOpn清單", page)
    ops = parse_op_list(page)
    op = pick_panc_head(ops)
    if not op:
        note(f"手術記錄清單（{len(ops)} 筆）裡沒有胰臟頭部手術")
        return row
    row["手術日期"] = op["date"]
    row["清單手術名稱"] = "; ".join(op["names"])
    log(f"      胰頭手術 {op['date']:%Y-%m-%d}：{row['清單手術名稱']}")
    if sum(1 for o in ops if any(is_panc_head(n) for n in o["names"])) > 1:
        note("有多次胰頭手術，取最近一次")

    # --- 2. 手術記錄內容 ---
    dt = op["dt"] or roc_dt(op["date"])
    page = get_page(eip, f"{base}?action=findOpn&histno={histno}&dt={dt}&tm={op['tm'] or 1}",
                    r"術前診斷", log=log)
    _save(save_dir, histno, "2_findOpn內容", page)
    det = parse_op_detail(page)
    row["術前診斷"] = det["術前診斷"]
    row["手術名稱"] = "; ".join(det["手術名稱"]) or row["清單手術名稱"]
    if not det["術前診斷"]:
        note("手術記錄讀不到術前診斷")

    # 「術前」的切點：手術開始時間；讀不到時間就只收手術日前一天以前的資料
    if det["start"] and det["has_start_time"] and det["start"].date() == op["date"]:
        cutoff = det["start"]
    else:
        cutoff = datetime.combine(op["date"], datetime.min.time()) - timedelta(seconds=1)
        note("讀不到手術開始時間，只採手術前一天（含）以前的檢驗與量測")
    row["手術開始時間"] = det["start"]
    row["手術當下年齡"] = age_at(birth, op["date"])

    # --- 3. 檢驗：每份有勾的報告各開一頁（同一份報告的項目都在同一頁）。
    #     resdtmonth 依手術至今的月數（見 report_periods） ---
    for rtype, codes in lab_items.by_report(labs).items():
        url_type = lab_items.URL_TYPE.get(rtype, rtype)
        for period in report_periods(rtype, cutoff, extra_months):
            missing = [c for c in codes if row.get(f"術前{rtype}:{c}") is None]
            if not missing:
                break
            page = get_page(eip, f"{base}?action=findResd&histno={histno}"
                                 f"&resdtmonth={period}&resdtype={url_type}",
                            r'resdtable|累積報告', log=log)
            _save(save_dir, histno, f"3_findResd_{url_type}_{period}", page)
            since = (months_before(date.today(), period)
                     if lab_items.PERIOD.get(rtype) == "months" else None)
            for c in missing:
                t, v = latest_lab_before(parse_lab_series(page, c), cutoff, since)
                if t:
                    row[f"術前{rtype}:{c}"], row[f"{rtype}:{c}時間"] = v, t
    lost = [lab_items.header(k) for k in labs if row.get(f"術前{k}") is None]
    if lost:
        note(f"術前查無：{'、'.join(lost)}")

    # --- 4. 身高體重 ---
    if vitals:
        page = get_page(eip, f"{base}?action=findVts&histno={histno}&caseno=all&pbvtype=HWS",
                        r'name="Weight"|pbvlisttable', log=log)
        _save(save_dir, histno, "4_findVts", page)
        t, h, w = latest_vitals_before(parse_vitals(page), cutoff)
        if t:
            row["術前身高"], row["術前體重"], row["身高體重時間"] = h, w, t
            row["術前BMI"] = bmi(h, w)
        else:
            note("術前身高體重查無資料")

    parts = [f"年齡 {row.get('手術當下年齡')}"]
    if vitals:
        parts.append(f"BMI {row.get('術前BMI')}")
    parts += [f"{lab_items.header(k)} {row.get(f'術前{k}', '-')}" for k in labs]
    log("      " + "　".join(parts))
    return row


# ============================ Excel ============================

def excel_columns(labs=LAB_ITEMS, vitals=True):
    """依勾選的項目組出兩張工作表的欄位 [(key, 欄寬, 數字格式, 欄名)]。"""
    labs = lab_items.normalize(labs)
    dt = "yyyy-mm-dd hh:mm"
    main = [("病歷號", 12, "@", "病歷號"),
            ("術前診斷", 40, None, "術前診斷"),
            ("手術日期", 12, "yyyy-mm-dd", "手術日期"),
            ("手術名稱", 40, None, "手術名稱"),
            ("手術當下年齡", 8, "0", "手術當下年齡")]
    src = [("病歷號", 12, "@", "病歷號"),
           ("生日", 12, "yyyy-mm-dd", "生日"),
           ("手術開始時間", 17, dt, "手術開始時間"),
           ("清單手術名稱", 40, None, "清單手術名稱")]
    if vitals:
        main += [("術前身高", 9, "0.0", "術前身高(cm)"),
                 ("術前體重", 9, "0.0", "術前體重(kg)"),
                 ("術前BMI", 9, "0.00", "術前BMI")]
        src.append(("身高體重時間", 17, dt, "身高體重時間"))
    for k in labs:
        h = lab_items.header(k)
        main.append((f"術前{k}", 9, None, f"術前{h}"))
        src.append((f"{k}時間", 17, dt, f"{h}時間"))
    main.append(("備註", 50, None, "備註"))
    return main, src


def write_excel(rows, out_path, labs=LAB_ITEMS, vitals=True):
    """寫 Excel。目標檔正被 Excel 開著（存不進去）就改存成另一個檔名，回傳實際路徑。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    main_cols, src_cols = excel_columns(labs, vitals)
    wb = Workbook()
    sheets = ((wb.active, "術前資料", main_cols), (wb.create_sheet(), "資料時間（核對用）", src_cols))
    for ws, title, cols in sheets:
        ws.title = title
        for c, (key, width, _, label) in enumerate(cols, 1):
            cell = ws.cell(row=1, column=c, value=label)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(c)].width = width
        for r, row in enumerate(rows, 2):
            for c, (key, _, fmt, _label) in enumerate(cols, 1):
                v = row.get(key)
                if key == "備註":
                    v = "；".join(v or [])
                cell = ws.cell(row=r, column=c, value=v)
                if fmt:
                    cell.number_format = fmt
                cell.alignment = Alignment(vertical="top", wrap_text=key in ("術前診斷", "手術名稱", "備註", "清單手術名稱"))
        ws.freeze_panes = "B2"
    base, ext = os.path.splitext(out_path)
    for n in range(0, 20):
        path = out_path if n == 0 else f"{base}_{n}{ext}"
        try:
            wb.save(path)
            return path
        except PermissionError:
            continue
    raise PermissionError(f"存不進 {out_path}（檔案是不是正被 Excel 開著？）")


# ============================ 執行 ============================

# ---- 速度與斷點續跑 ----
PATIENT_DELAY = 5.0          # 每位病人之間額外多停幾秒（網頁之間的間隔見 eip_session.REQUEST_INTERVAL）
MAX_CONSECUTIVE_FAIL = 3     # 連續這麼多位病人連 EMR 首頁都打不開，就當作被鎖、停止整批
PROGRESS_FILE = os.path.join(BASE_DIR, "收集進度.json")

# 備註裡出現這些字代表這位沒查完，續跑時要重查
_RETRY_NOTES = ("查詢失敗", "EMR 病歷首頁打不開", "疑似被鎖", "尚未查詢")


def needs_retry(row) -> bool:
    return any(k in n for n in row.get("備註", []) for k in _RETRY_NOTES)


def covers(row, labs, vitals) -> bool:
    """上次查好的這一列，有沒有涵蓋這次勾選的項目（沒有就要重查這位）。"""
    if not row.get("手術日期"):
        return True                      # 沒有胰頭手術，檢驗項目無關
    have = set(lab_items.normalize(row.get("_labs", LAB_ITEMS)))
    return set(lab_items.normalize(labs)) <= have and (not vitals or row.get("_vitals", True))


def _migrate(row) -> dict:
    """舊版進度檔：項目只存代碼（'術前ALB'），換成新版的 '術前DCHEM:ALB'。"""
    old = row.get("_labs", ["ALB", "CREA", "BILIT"])
    if all(":" in str(x) for x in old) and "_labs" in row:
        return row
    for code in old:
        code = str(code)
        if ":" in code:
            continue
        lid = lab_items.to_id(code)
        if not lid:
            continue
        for a, b in ((f"術前{code}", f"術前{lid}"), (f"{code}時間", f"{lid}時間")):
            if a in row:
                row[b] = row.pop(a)
    row["_labs"] = lab_items.normalize(old)
    return row


def _to_json(v):
    if isinstance(v, datetime):
        return {"__dt__": v.isoformat()}
    if isinstance(v, date):
        return {"__d__": v.isoformat()}
    return v


def _from_json(v):
    if isinstance(v, dict) and "__dt__" in v:
        return datetime.fromisoformat(v["__dt__"])
    if isinstance(v, dict) and "__d__" in v:
        return date.fromisoformat(v["__d__"])
    return v


def load_progress(path=PROGRESS_FILE) -> dict:
    """{病歷號: 一列結果}，只留已經查完的。"""
    import json
    if not os.path.exists(path):
        return {}
    try:
        data = json.load(open(path, encoding="utf-8"))
    except Exception:
        return {}
    rows = {h: _migrate({k: _from_json(v) for k, v in r.items()}) for h, r in data.items()}
    return {h: r for h, r in rows.items() if not needs_retry(r)}


def save_progress(done: dict, path=PROGRESS_FILE):
    import json
    data = {h: {k: _to_json(v) for k, v in r.items()} for h, r in done.items()}
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def estimate_minutes(n, labs=LAB_ITEMS, vitals=True, interval=None, patient_delay=None):
    """粗估 n 位要跑幾分鐘（每位約開 4～6 頁，加上每頁載入時間）。"""
    import eip_session
    interval = eip_session.REQUEST_INTERVAL if interval is None else interval
    patient_delay = PATIENT_DELAY if patient_delay is None else patient_delay
    pages = 3 + len(lab_items.by_report(labs)) + (1 if vitals else 0)
    per = pages * (interval + 1.5) + patient_delay
    return n * per / 60


def run(histnos, out_path, log=print, headless=True, save_dir=None, restart=False,
        labs=LAB_ITEMS, vitals=True, stop_event=None, progress=None,
        acc=None, pwd=None, request_interval=None, patient_delay=None, extra_months=None):
    """整批查詢並輸出 Excel，回傳 (rows, 實際存檔路徑)。

    extra_months：檢驗報告除了手術至今的月數，再往前多查幾個月（預設 EXTRA_MONTHS）。

    stop_event：threading.Event，介面按「停止」時 set()，查完手上這位就停。
    progress(i, n, histno)：每開始查一位叫一次，介面用來更新進度條。
    """
    import eip_session

    labs = lab_items.normalize(labs)
    if request_interval is not None:
        eip_session.REQUEST_INTERVAL = max(0.0, float(request_interval))
    pdelay = PATIENT_DELAY if patient_delay is None else max(0.0, float(patient_delay))

    done = {} if restart else load_progress()
    done = {h: r for h, r in done.items() if covers(r, labs, vitals)}
    skip = [h for h in histnos if h in done]
    todo = [h for h in histnos if h not in done]
    if skip:
        log(f"■ 續跑：{len(skip)} 位上次已查完，這次跳過；剩 {len(todo)} 位要查")
    log(f"■ 檢驗項目：{'、'.join(lab_items.header(k) for k in labs) or '（不查）'}"
        f"{'　＋身高體重' if vitals else ''}")
    if any(lab_items.PERIOD.get(r) == "months" for r in lab_items.by_report(labs)):
        em = EXTRA_MONTHS if extra_months is None else extra_months
        log(f"■ 檢驗查詢範圍：手術至今的月數 ＋ 往前多查 {em} 個月")
    log(f"■ 速度：每開一頁至少間隔 {eip_session.REQUEST_INTERVAL:g} 秒，"
        f"每位病人之間再多停 {pdelay:g} 秒"
        f"（{len(todo)} 位約 {estimate_minutes(len(todo), labs, vitals, patient_delay=pdelay):.0f} 分鐘）")

    stopped, by_user = "", False
    saved_path = out_path
    try:
        if todo:
            with eip_session.EipSession(log=log, headless=headless, acc=acc, pwd=pwd) as eip:
                log(f"[2/4] 開始查詢，共 {len(todo)} 位")
                fails = 0
                for i, histno in enumerate(todo, 1):
                    if i > 1:
                        if stop_event is not None:
                            stop_event.wait(pdelay)
                        else:
                            time.sleep(pdelay)
                    if stop_event is not None and stop_event.is_set():
                        stopped, by_user = "已依使用者要求停止", True
                        break
                    if progress:
                        progress(i - 1, len(todo), histno)
                    log(f"    ({i}/{len(todo)}) {histno}")
                    try:
                        row = collect_one(histno, eip, log=log, save_dir=save_dir,
                                          labs=labs, vitals=vitals, extra_months=extra_months)
                    except BlockedError as e:
                        stopped = f"{e}，疑似被鎖"
                        break
                    except Exception as e:
                        log(f"      [!] 查詢失敗：{e}")
                        row = {"病歷號": histno, "備註": [f"查詢失敗：{e}"]}
                    done[histno] = row
                    save_progress(done)

                    fails = fails + 1 if "EMR 病歷首頁打不開" in row.get("備註", []) else 0
                    if fails >= MAX_CONSECUTIVE_FAIL:
                        stopped = f"連續 {fails} 位病人都打不開 EMR，疑似被鎖"
                        break
                if progress and not stopped:
                    progress(len(todo), len(todo), "")
    finally:
        # 中途出錯、被鎖、按停止或 Ctrl+C 中斷，已查到的也照樣存檔
        if by_user:
            log(f"\n[!] {stopped}。已查到的資料都保存了，下次執行會接著查。")
        elif stopped:
            log(f"\n[!] {stopped}，已停止。已查到的資料都保存了。")
            log("    建議等一段時間（或請資訊室解鎖）後，把每頁間隔調大"
                "（命令列用 --delay）再執行一次，會自動從沒查完的病人接著查。")
        rows = []
        for h in histnos:
            r = done.get(h) or {"病歷號": h, "備註": ["尚未查詢（程式中途停止）"]}
            if stopped and not by_user and "EMR 病歷首頁打不開" in r.get("備註", []):
                r["備註"].append("疑似被鎖")
            rows.append(r)
        log("[3/4] 寫入 Excel")
        saved_path = write_excel(rows, out_path, labs=labs, vitals=vitals)
        if saved_path != out_path:
            log(f"    {os.path.basename(out_path)} 正被開著，改存成 {os.path.basename(saved_path)}")
        ok = sum(1 for r in rows if r.get("手術日期"))
        pending = sum(1 for r in rows if needs_retry(r))
        log(f"[4/4] {ok} / {len(rows)} 位找到胰頭手術 -> {saved_path}")
        if pending:
            log(f"    還有 {pending} 位沒查完，再執行一次會只查這些。")
        elif os.path.exists(PROGRESS_FILE):
            os.remove(PROGRESS_FILE)      # 全部查完，進度檔用不到了
    return rows, saved_path


def main():
    global PATIENT_DELAY
    ap = argparse.ArgumentParser(description="收集胰臟頭部手術（Whipple/PD）術前資料，輸出 Excel")
    ap.add_argument("list", nargs="?", default=None, help="病歷號清單 txt（預設：同資料夾的 病歷號清單.txt）")
    ap.add_argument("--histno", action="append", default=[], help="直接指定病歷號，可重複")
    ap.add_argument("--out", default=None, help="輸出 Excel 路徑")
    ap.add_argument("--visible", action="store_true", help="開有畫面的 Chrome")
    ap.add_argument("--save-html", action="store_true",
                    help="把每個網頁存進「除錯HTML」資料夾（含病人資料，用完請刪）")
    ap.add_argument("--delay", type=float, default=None,
                    help="每開一頁至少間隔幾秒（預設 3）。被鎖過就調大，例如 --delay 6")
    ap.add_argument("--patient-delay", type=float, default=None,
                    help=f"每位病人之間再多停幾秒（預設 {PATIENT_DELAY:g}）")
    ap.add_argument("--restart", action="store_true",
                    help="不理會上次的進度，全部重查")
    ap.add_argument("--labs", default=",".join(LAB_ITEMS),
                    help="要查的檢驗代碼，逗號分隔（預設 ALB,CREA,BILIT）。只寫代碼會自動對到"
                         "所屬報告，也可以寫 DCBC:HGB。候選："
                         + "；".join(f"{lab_items.REPORT_TITLE[r]} {','.join(lab_items.CODES[r])}"
                                    for r in lab_items.REPORT_TYPES))
    ap.add_argument("--no-vitals", action="store_true", help="不查身高體重")
    ap.add_argument("--extra-months", type=int, default=None,
                    help=f"檢驗報告除了手術至今的月數，再往前多查幾個月（預設 {EXTRA_MONTHS}）")
    a = ap.parse_args()

    import eip_session
    if a.delay is not None:
        eip_session.REQUEST_INTERVAL = max(0.0, a.delay)
    if a.patient_delay is not None:
        PATIENT_DELAY = max(0.0, a.patient_delay)

    histnos = list(a.histno)
    list_path = a.list or (DEFAULT_LIST if not histnos else None)
    if list_path:
        if not os.path.exists(list_path):
            print(f"找不到病歷號清單：{list_path}")
            print("請建立這個檔案，一行一個病歷號，或改用 --histno 指定。")
            return 1
        for h in read_histno_list(list_path):
            if h not in histnos:
                histnos.append(h)
    if not histnos:
        print("清單裡沒有病歷號。")
        return 1

    try:
        import openpyxl  # noqa: F401
    except ImportError:
        print("缺少 openpyxl，請先執行：py -3 -m pip install openpyxl")
        return 1

    out = a.out or os.path.join(BASE_DIR, f"胰頭手術術前資料_{datetime.now():%Y%m%d_%H%M}.xlsx")
    save_dir = os.path.join(BASE_DIR, "除錯HTML") if a.save_html else None
    run(histnos, out, headless=not a.visible, save_dir=save_dir, restart=a.restart,
        labs=lab_items.normalize(a.labs), vitals=not a.no_vitals, extra_months=a.extra_months)
    if save_dir:
        print(f"網頁已存在 {save_dir}（含病人資料，用完請刪）")
    return 0


if __name__ == "__main__":
    code = 1
    try:
        code = main()
    except KeyboardInterrupt:
        print("\n已中斷。已查到的資料都存了，再執行一次會接著查。")
    except Exception:
        traceback.print_exc()
    if len(sys.argv) == 1:           # 直接雙擊執行時，讓視窗停住看結果
        input("\n按 Enter 結束...")
    sys.exit(code)
