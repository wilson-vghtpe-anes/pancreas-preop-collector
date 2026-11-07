# -*- coding: utf-8 -*-
"""
EIP 登入與 EMR 病歷首頁 —— 從交班單專案抽出來的最小連線模組。

來源：
  EipSession     <- handover_core.EipSession（改成從 eip帳號.txt 讀帳密）
  open_patient() <- anes_record.open_patient()

不需要 PyMuPDF，只需要 selenium 與 requests。

帳號密碼放在同資料夾的 eip帳號.txt：
    帳號=XXXXXXXX
    密碼=XXXXXXXX
檔案不存在或沒填時，第一次執行會在視窗裡詢問，並幫你存進這個檔。
"""

from __future__ import annotations

import getpass
import os
import re
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIAL_FILE = os.path.join(BASE_DIR, "eip帳號.txt")

LOGIN_PAGE = "https://eip.vghtpe.gov.tw/login.php"

# EMR 病歷首頁。兩個網域都有同一支 qemr.cfm，rkeprd 開不起來就換 web9
EMR_HOSTS = ("https://rkeprd.vghtpe.gov.tw", "https://web9.vghtpe.gov.tw")
EMR_PATH = "/emr/qemr/qemr.cfm"
PAGE_WAIT = 12.0          # 等網頁內容出現的秒數上限

# 每兩次開網頁之間至少間隔幾秒。院內系統對短時間大量查詢會鎖 IP，
# 被鎖過就調大一點；收集胰頭手術資料.py 的 --delay 會改這個值。
REQUEST_INTERVAL = 3.0

_last_request = 0.0


def throttle():
    """距離上一次開網頁不到 REQUEST_INTERVAL 秒就先等，所有網路請求前都要叫。"""
    global _last_request
    wait = _last_request + REQUEST_INTERVAL - time.monotonic()
    if wait > 0:
        time.sleep(wait)
    _last_request = time.monotonic()


# ============================ 帳號密碼 ============================

def _read_text_any(path: str) -> str:
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-16", "cp950", "big5hkscs"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", errors="replace")


def read_credentials(path: str = CREDENTIAL_FILE):
    """讀 eip帳號.txt，回傳 (帳號, 密碼)；沒有檔案或沒填就是空字串，不會詢問。

    檔案格式寬鬆：「帳號=…」「密碼=…」（冒號也可以），或者只寫兩行，
    第一行帳號、第二行密碼。# 開頭是註解。
    """
    acc = pwd = ""
    if os.path.exists(path):
        plain = []
        for line in _read_text_any(path).splitlines():
            s = line.strip()
            if not s or s.startswith("#"):
                continue
            m = re.match(r"(帳號|密碼|account|password)\s*[=:：]\s*(.*)$", s, re.I)
            if m:
                key, val = m.group(1).lower(), m.group(2).strip()
                if key in ("帳號", "account"):
                    acc = val
                else:
                    pwd = val
            else:
                plain.append(s)
        if not (acc or pwd) and len(plain) >= 2:
            acc, pwd = plain[0], plain[1]
    return acc, pwd


def save_credentials(acc: str, pwd: str, path: str = CREDENTIAL_FILE):
    with open(path, "w", encoding="utf-8") as f:
        f.write("# EIP 登入帳號密碼（收集胰頭手術資料使用），請勿外流\n")
        f.write(f"帳號={acc}\n密碼={pwd}\n")


def load_credentials(path: str = CREDENTIAL_FILE):
    """命令列用：讀 eip帳號.txt，沒有或沒填就在主控台詢問，並存起來。"""
    acc, pwd = read_credentials(path)
    if acc and pwd:
        return acc, pwd

    print(f"找不到 EIP 帳號密碼（{os.path.basename(path)}），請輸入：")
    acc = acc or input("  EIP 帳號：").strip()
    pwd = pwd or getpass.getpass("  EIP 密碼（輸入時不會顯示）：").strip()
    if not (acc and pwd):
        raise RuntimeError("沒有輸入 EIP 帳號或密碼")
    save_credentials(acc, pwd, path)
    print(f"  已存進 {os.path.basename(path)}，下次就不用再輸入。")
    return acc, pwd


# ============================ 登入 ============================

class EipSession:
    """登入 EIP 一次，之後 EMR 各頁都共用同一個瀏覽器／連線。

    以 with 使用，離開時自動關掉 Chrome：

        with EipSession(log=print) as eip:
            html, host = open_patient(histno, eip)

    .driver  已登入的 headless Chrome
    .sess    複製了 cookie 的 requests.Session
    """

    def __init__(self, log=print, headless=True, acc=None, pwd=None):
        self.log = log
        self.headless = headless
        self.acc, self.pwd = acc, pwd
        self.driver = None
        self.sess = None

    def __enter__(self):
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service as ChromeService
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait
        import requests

        if not (self.acc and self.pwd):
            self.acc, self.pwd = load_credentials()

        opts = Options()
        if self.headless:
            opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--log-level=3")
        opts.add_argument("--silent")
        opts.add_experimental_option("excludeSwitches",
                                     ["enable-logging", "enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)

        self.log("[1/4] 開啟 Chrome 並登入 EIP")
        self.driver = webdriver.Chrome(options=opts,
                                       service=ChromeService(log_path=os.devnull))
        try:
            self.driver.get(LOGIN_PAGE)
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, "login_name")))
            self.driver.find_element(By.ID, "login_name").send_keys(self.acc)
            self.driver.find_element(By.ID, "password").send_keys(self.pwd)
            self.driver.find_element(By.ID, "loginBtn").click()
            WebDriverWait(self.driver, 30).until(
                lambda d: "/login" not in d.current_url)
            self.log(f"    登入成功：{self.driver.current_url}")

            self.sess = requests.Session()
            self.sess.headers.update({"User-Agent": "Mozilla/5.0"})
            self.sess.trust_env = False       # 內網 IP 不走系統 proxy
            self.sess.proxies = {}
            self.sync_cookies()
            return self
        except Exception:
            self.close()
            raise

    def sync_cookies(self):
        """把瀏覽器目前網域的 cookie 搬進 requests.Session。

        driver.get_cookies() 只給得到「當下這個網域」的 cookie，所以換站台
        （EIP -> EMR）之後要再叫一次，requests 才帶得動登入狀態。
        """
        if not (self.driver and self.sess):
            return
        for c in self.driver.get_cookies():
            try:
                self.sess.cookies.set(name=c.get("name"), value=c.get("value"),
                                      domain=c.get("domain", ""),
                                      path=c.get("path", "/"))
            except Exception:
                pass

    def close(self):
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    def __exit__(self, *exc):
        self.close()
        return False


# ============================ EMR 病歷首頁 ============================

# 表頭那一列長這樣：王小明 12345678 67歲(19590101) 男 A000- 00(住院中) JR
_HEAD_RE = re.compile(r'<ul[^>]*class="[^"]*jd_menu[^"]*"[^>]*>\s*(<li.*?</li>)', re.S | re.I)


def open_patient(histno, eip, log=print):
    """導覽到 EMR 病歷首頁（action=findEmr），回傳 (html, host)。

    表頭那一列出現就算成功。兩個網域都試過還是不行就回 ("", "")。
    成功時已經同步過 cookie，後面可以直接用 eip.sess 打 requests。
    """
    for host in EMR_HOSTS:
        url = f"{host}{EMR_PATH}?action=findEmr&histno={histno}"
        try:
            throttle()
            eip.driver.get(url)
        except Exception as e:
            log(f"      findEmr@{host.split('//')[-1]} 開不起來（{e}）")
            continue
        deadline = time.time() + PAGE_WAIT
        while True:
            html = eip.driver.page_source or ""
            if _HEAD_RE.search(html):
                eip.sync_cookies()
                return html, host
            if time.time() > deadline:
                break
            time.sleep(0.4)
        log(f"      findEmr@{host.split('//')[-1]} 沒看到病歷表頭")
    return "", ""
