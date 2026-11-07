# -*- coding: utf-8 -*-
"""
術前檢驗的候選項目。

EMR 的累積報告網址是
    qemr.cfm?action=findResd&histno=病歷號&resdtmonth=年份&resdtype=報告類型
不同的 resdtype 是不同的報告頁，各自有一組欄位：
    DCHEM  生化（CHEM 累積報告，範例三）—— 31 項
    DCBC   血液 CBC/DC 與凝血（範例五）—— 21 項
兩份都用相對月份查：resdtmonth 填「手術至今經過幾個月，再往前多查幾個月」。
候選項目就是兩頁表頭的全部欄位。

每個項目用「報告類型:代碼」當識別（例 DCHEM:ALB、DCBC:HGB）。代碼就是網頁上
getImg('ALB') 與隱藏列 <td id="ALB"> 用的字；只寫代碼（ALB、HGB）也可以，
會自動對到它所屬的報告。

院內報告日後若多了新欄位，介面上各分頁的「其他代碼」直接打代碼即可；
要讓它出現在勾選清單裡，再加進下面對應的分類。新的報告類型（別的 resdtype）
加進 REPORTS 即可，介面會自動多一個分頁。
"""

from __future__ import annotations

import re

# (報告類型 resdtype, 介面上的名稱, [(分類, [(代碼, 說明), ...]), ...])
REPORTS = [
    ("DCHEM", "生化 CHEM", [
        ("營養／蛋白", [
            ("ALB", "白蛋白 Albumin"),
            ("TP", "總蛋白 Total protein"),
        ]),
        ("肝膽", [
            ("BILIT", "總膽紅素 T-bil"),
            ("DBILI", "直接膽紅素 D-bil"),
            ("AST", "AST (GOT)"),
            ("ALT", "ALT (GPT)"),
            ("ALKP", "鹼性磷酸酶 ALP"),
            ("GGT", "γ-GT"),
            ("LDH", "乳酸脫氫酶 LDH"),
        ]),
        ("腎功能", [
            ("CREA", "肌酸酐 Creatinine"),
            ("BUN", "尿素氮 BUN"),
            ("EGFR", "eGFR(M)"),
            ("UA", "尿酸 Uric acid"),
        ]),
        ("電解質", [
            ("NA", "鈉 Na"),
            ("K", "鉀 K"),
            ("CL", "氯 Cl"),
            ("CA", "鈣 Ca"),
            ("IP", "無機磷 P"),
            ("CO2", "CO2（HCO3）"),
        ]),
        ("血糖／血脂", [
            ("GLU", "血糖 Glucose"),
            ("CHOL", "總膽固醇"),
            ("TG", "三酸甘油酯 TG"),
            ("HDLC", "HDL-C"),
            ("LDLC", "LDL-C"),
            ("RISKF", "膽固醇/HDL 比值"),
        ]),
        ("心肌／發炎", [
            ("CRP", "C 反應蛋白 CRP"),
            ("CK", "肌酸激酶 CK"),
            ("CKMB", "CK-MB"),
            ("TROP", "Troponin"),
        ]),
        ("鐵", [
            ("IRON", "血清鐵 Fe"),
            ("TIBC", "總鐵結合能力 TIBC"),
        ]),
    ]),
    ("DCBC", "血液 CBC/DC／凝血", [
        ("血球計數", [
            ("HGB", "血紅素 Hb"),
            ("HCT", "血比容 Hct"),
            ("PLT", "血小板 Platelet"),
            ("WBC", "白血球 WBC"),
            ("RBC", "紅血球 RBC"),
        ]),
        ("紅血球指數", [
            ("MCV", "MCV"),
            ("MCH", "MCH"),
            ("MCHC", "MCHC"),
            ("RDW", "RDW"),
        ]),
        ("白血球分類 DC（%）", [
            ("SEG", "嗜中性球 Segment"),
            ("BAND", "帶狀球 Band"),
            ("LYM", "淋巴球 Lymphocyte"),
            ("MONO", "單核球 Monocyte"),
            ("EOS", "嗜酸性球 Eosinophil"),
            ("BASO", "嗜鹼性球 Basophil"),
        ]),
        ("凝血", [
            ("INRPT", "INR"),
            ("PT", "PT（秒）"),
            ("APTT", "aPTT（秒）"),
            ("FIBRINOGEN", "Fibrinogen"),
            ("DDIMER", "D-dimer"),
            ("FDP", "FDP"),
        ]),
    ]),
]

# resdtmonth 參數怎麼填："months" = 往回查幾個月（目前兩份報告都用這個）；
# "year" = 西元年份（舊做法，保留給日後若有報告只能用年份查）
PERIOD = {
    "DCHEM": "months",
    "DCBC": "months",
}

# 網址裡實際用的 resdtype。用月數查時，網頁標題會變成「CBCN…」、頁尾寫
# DCBCN，那是伺服器自己標的，網址照樣用 DCBC（已實測可以）。
# 左邊是勾選設定與進度檔用的識別，別動。
URL_TYPE = {
    "DCHEM": "DCHEM",
    "DCBC": "DCBC",
}

# 各報告「預設」按鈕勾的項目（介面第一次打開也勾這些）
DEFAULTS = {
    "DCHEM": ("ALB", "CREA", "BILIT"),
    "DCBC": (),
}
# 各報告的「常用」按鈕（沒有就不顯示）
COMMON = {
    "DCBC": ("HGB", "PLT", "WBC", "INRPT"),
}

REPORT_TYPES = tuple(r for r, _, _ in REPORTS)
REPORT_TITLE = {r: t for r, t, _ in REPORTS}
GROUPS = {r: g for r, _, g in REPORTS}
CODES = {r: tuple(c for _, items in g for c, _ in items) for r, _, g in REPORTS}
NAMES = {f"{r}:{c}": n for r, _, g in REPORTS for _, items in g for c, n in items}
ALL_IDS = tuple(f"{r}:{c}" for r in REPORT_TYPES for c in CODES[r])

# 只寫代碼時對到哪份報告（同一個代碼若出現在兩份報告，取前面那份）
_CODE_OWNER = {}
for _r in REPORT_TYPES:
    for _c in CODES[_r]:
        _CODE_OWNER.setdefault(_c, _r)

DEFAULT_LABS = tuple(f"{r}:{c}" for r in REPORT_TYPES for c in DEFAULTS.get(r, ()))

# Excel 欄名：代碼跟網頁表頭顯示的不一樣的，在這裡對應
HEADER = {"DCHEM:EGFR": "eGFR(M)"}

HEADER["DCBC:INRPT"] = "INR"

_CODE_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,15}$")
_SPLIT_RE = re.compile(r"[\s,，、;；/]+")


def split_id(lab_id: str):
    """'DCBC:HGB' -> ('DCBC', 'HGB')"""
    r, _, c = lab_id.partition(":")
    return r, c


def header(lab_id: str) -> str:
    """Excel 欄名與訊息用的顯示名。代碼重複出現在兩份報告時加上報告類型。"""
    if lab_id in HEADER:
        return HEADER[lab_id]
    r, c = split_id(lab_id)
    if _CODE_OWNER.get(c, r) != r:
        return f"{c}({r})"
    return c


def to_id(token: str, default_report: str | None = None) -> str | None:
    """把一個字整理成 '報告:代碼'。看不懂回 None。

    'alb' -> 'DCHEM:ALB'   'hgb' -> 'DCBC:HGB'   'dcbc:xyz' -> 'DCBC:XYZ'
    清單外的代碼：有給 default_report 就歸到那份報告，否則歸到 DCHEM。
    """
    t = str(token or "").strip().upper()
    if ":" in t:
        r, c = split_id(t)
        if r and _CODE_RE.match(c):
            return f"{r}:{c}"
        return None
    if not _CODE_RE.match(t):
        return None
    if default_report:
        return f"{default_report}:{t}"
    return f"{_CODE_OWNER.get(t, REPORT_TYPES[0])}:{t}"


def normalize(items, default_report: str | None = None) -> list[str]:
    """整理使用者選的項目：轉成 '報告:代碼'、去重複、照候選清單的順序排，
    清單外的接在後面。

    items 可以是 list，也可以是 'ALB, crea  HGB' 這種字串。看不懂的字會被丟掉。
    舊版只存代碼（'ALB'）的設定與進度檔，也會被正確轉換。
    """
    if isinstance(items, str):
        items = _SPLIT_RE.split(items)
    picked = []
    for tok in items or ():
        lid = to_id(tok, default_report)
        if lid and lid not in picked:
            picked.append(lid)
    known = [i for i in ALL_IDS if i in picked]
    return known + [i for i in picked if i not in ALL_IDS]


def invalid_tokens(text: str) -> list[str]:
    """「其他代碼」欄裡不像代碼的字，介面用來提醒使用者。"""
    return [t for t in _SPLIT_RE.split(text or "") if t and to_id(t, REPORT_TYPES[0]) is None]


def by_report(lab_ids) -> dict[str, list[str]]:
    """{報告類型: [代碼, ...]}，照 REPORT_TYPES 的順序，沒選的報告不列。"""
    out = {}
    for lid in normalize(lab_ids):
        r, c = split_id(lid)
        out.setdefault(r, []).append(c)
    order = {r: i for i, r in enumerate(REPORT_TYPES)}
    return dict(sorted(out.items(), key=lambda kv: order.get(kv[0], len(order))))


def page_columns(page: str) -> list[str]:
    """從實際的累積報告網頁讀出有哪些欄位代碼（getImg('XXX')）。"""
    return list(dict.fromkeys(re.findall(r"getImg\('([A-Za-z0-9_]+)'\)", page or "")))
