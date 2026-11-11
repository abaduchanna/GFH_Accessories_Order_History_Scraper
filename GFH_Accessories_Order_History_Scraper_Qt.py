#!/usr/bin/env python3
"""
GFH Telecom — Accessories Order History Scraper (PySide6/Qt build)
=====================================================================
Same tool as the Tkinter build - logs into the CPWH Wireless portal,
extracts the accessories order history, and saves it to Excel. Only the
UI toolkit differs; see GFH_Accessories_Order_History_Scraper.py's
module docstring for the full feature description.
"""

import subprocess
import sys


def _pkg_version(dist):
    try:
        import importlib.metadata as _md
        return _md.version(dist)
    except Exception:
        return None


def _ensure(pip_name, imp_name):
    if getattr(sys, "frozen", False):
        return
    if _pkg_version(pip_name) is not None:
        return
    try:
        __import__(imp_name.split(".")[0])
    except ImportError:
        try:
            print(f"Installing {pip_name}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", pip_name, "-q"])
        except Exception as e:
            print(f"  [WARN] could not install {pip_name}: {e}")


for _p, _i in [("selenium", "selenium"), ("pandas", "pandas"), ("openpyxl", "openpyxl"),
               ("pdfplumber", "pdfplumber"), ("requests", "requests"), ("Pillow", "PIL")]:
    _ensure(_p, _i)

import os
import io
import re
import time
import queue
import shutil
import threading
import traceback
import zipfile
import urllib.request
from datetime import datetime, date, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── LAZY heavy imports (opening-hang fix) ─────────────────────────────────
pd = None
webdriver = None
By = None
WebDriverWait = None
EC = None
Service = None
PatternFill = None
Font = None
Alignment = None
Border = None
Side = None
get_column_letter = None


def _ensure_heavy_imports():
    global pd, webdriver, By, WebDriverWait, EC, Service
    global PatternFill, Font, Alignment, Border, Side, get_column_letter
    if pd is None:
        import pandas as _pd_mod
        pd = _pd_mod
    if webdriver is None:
        from selenium import webdriver as _w
        from selenium.webdriver.common.by import By as _By
        from selenium.webdriver.support.ui import WebDriverWait as _W
        from selenium.webdriver.support import expected_conditions as _EC
        from selenium.webdriver.edge.service import Service as _S
        webdriver, By, WebDriverWait, EC, Service = _w, _By, _W, _EC, _S
    if get_column_letter is None:
        from openpyxl.styles import (PatternFill as _PF, Font as _F,
                                      Alignment as _A, Border as _B, Side as _Sd)
        from openpyxl.utils import get_column_letter as _GCL
        PatternFill, Font, Alignment, Border, Side = _PF, _F, _A, _B, _Sd
        get_column_letter = _GCL

from theme_manager_qt import ThemeManagerQt

# ============================================================================
# BRAND / CONFIGURATION
# ============================================================================
APP_TITLE = "GFH Telecom — Accessories Order History Scraper"

ICON_ICO_NAME = "gfh_icon.ico"
WORDMARK_PNG_NAME = "GFH_Telecom_Logo.png"
EMBEDDED_ICON_B64 = (
    open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "embedded_icon_b64.txt"), "r").read().strip()
    if not getattr(sys, "frozen", False)
    else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "embedded_icon_b64.txt"), "r").read().strip()
)

# ── Edge automation profile + port ───────────────────────────────────────────
AUTOMATION_PROFILE_DIR = r"C:\GFH_Edge_Automation_Profile_Scraper"
REMOTE_DEBUGGING_PORT = 9226
ATTACH_TO_OPEN_EDGE = True

LOGIN_URL = "https://www.cpwhwireless.com/login"
USERNAME = "nofalgodil@gfhtelecom.com"
PASSWORD = "Verizon123"

downloads_folder = r"C:\Users\AbadUmairChanna\Downloads"
zip_path = os.path.join(downloads_folder, "edgedriver_win64.zip")
driver_path = os.path.join(downloads_folder, "msedgedriver.exe")
download_dir = downloads_folder


def get_script_dir():
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass and meipass.strip():
            return meipass
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _resource_path(name):
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, name)


# ============================================================================
# EDGE DRIVER — version-aware auto-management
# ============================================================================
def _get_installed_edge_version():
    try:
        edge_paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]
        for p in edge_paths:
            if os.path.exists(p):
                result = subprocess.run(
                    ["powershell", "-NoProfile", "-Command",
                     f"(Get-Item '{p}').VersionInfo.ProductVersion"],
                    capture_output=True, text=True, timeout=15
                )
                if result.returncode == 0 and result.stdout.strip():
                    return result.stdout.strip()
        try:
            import winreg
            reg_paths = [
                r"Software\Microsoft\Edge\BLBeacon",
                r"Software\WOW6432Node\Microsoft\Edge\BLBeacon",
            ]
            for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                for reg_path in reg_paths:
                    try:
                        key = winreg.OpenKey(root, reg_path)
                        version, _ = winreg.QueryValueEx(key, "version")
                        if version:
                            return version
                    except Exception:
                        pass
        except Exception:
            pass
    except Exception as e:
        print(f"  [WARN] Could not detect Edge version: {e}")
    return None


def _get_cached_driver_version():
    if not os.path.exists(driver_path):
        return None
    try:
        result = subprocess.run([driver_path, "--version"], capture_output=True, text=True, timeout=15)
        if result.returncode == 0:
            for tok in result.stdout.split():
                if tok and tok[0].isdigit() and "." in tok:
                    return tok
    except Exception:
        pass
    return None


def _download_matching_driver(major_version):
    if major_version:
        index_url = f"https://msedgedriver.azureedge.net/LATEST_RELEASE_{major_version}_WINDOWS"
        print(f"  Fetching latest driver build for Edge {major_version}.x ...")
    else:
        index_url = "https://msedgedriver.azureedge.net/LATEST_RELEASE"
        print("  Fetching latest stable Edge driver build ...")
    try:
        req = urllib.request.Request(index_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as r:
            full_version = r.read().decode("utf-8").strip()
    except Exception as e:
        print(f"  [ERROR] Could not fetch driver version index: {e}")
        return False

    download_url = f"https://msedgedriver.azureedge.net/{full_version}/edgedriver_win64.zip"
    print(f"  Downloading Edge WebDriver {full_version} ...")
    try:
        for stale in (zip_path, driver_path):
            if os.path.exists(stale):
                try:
                    os.remove(stale)
                except Exception:
                    pass
        urllib.request.urlretrieve(download_url, zip_path)
        print("  Extraction ...")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(downloads_folder)
        print("  Driver downloaded and extracted.")
        return True
    except Exception as e:
        print(f"  [ERROR] Download/extract failed: {e}")
        return False


def get_edge_exe_path():
    possible_paths = [
        shutil.which("msedge"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for path in possible_paths:
        if path and os.path.exists(path):
            return path
    return None


def is_port_open(host="127.0.0.1", port=REMOTE_DEBUGGING_PORT, timeout=1):
    import socket
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except Exception:
        return False


def open_vpn_setup_browser(url="about:blank", log=print):
    edge_path = get_edge_exe_path()
    if not edge_path:
        log("Microsoft Edge executable not found.")
        return False

    os.makedirs(AUTOMATION_PROFILE_DIR, exist_ok=True)

    args = [
        edge_path,
        f"--remote-debugging-port={REMOTE_DEBUGGING_PORT}",
        f"--user-data-dir={AUTOMATION_PROFILE_DIR}",
        "--profile-directory=Default",
        "--no-first-run",
        "--no-default-browser-check",
        url,
    ]

    try:
        subprocess.Popen(args)
        log("Opened dedicated Edge automation browser.")

        for _ in range(20):
            if is_port_open():
                log("Automation Edge remote connection is ready.")
                return True
            time.sleep(0.5)

        log("Edge opened, but remote debugging port is not ready yet.")
        return False
    except Exception as e:
        log(f"Failed to open Edge: {e}")
        return False


def create_edge_driver(log=print):
    if ATTACH_TO_OPEN_EDGE:
        if not is_port_open():
            log("Automation Edge is not open.")
            log("Opening VPN Browser Setup now.")
            open_vpn_setup_browser(log=log)

        if not is_port_open():
            raise RuntimeError(
                "Automation Edge is not available on remote debugging port. "
                "Click Open VPN Browser Setup first and keep that Edge window open."
            )

        options = webdriver.EdgeOptions()
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{REMOTE_DEBUGGING_PORT}")

        driver = webdriver.Edge(options=options)
        driver.set_page_load_timeout(90)
        return driver

    options = _build_edge_options()
    options.add_argument(f"--user-data-dir={AUTOMATION_PROFILE_DIR}")
    options.add_argument("--profile-directory=Default")
    driver = webdriver.Edge(options=options)
    driver.set_page_load_timeout(90)
    return driver


def _build_edge_options():
    edge_options = webdriver.EdgeOptions()
    edge_options.add_argument("--disable-gpu")
    edge_options.add_argument("--no-sandbox")
    edge_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    edge_options.add_experimental_option("prefs", prefs)
    return edge_options


# ============================================================================
# SCRAPER LOGIC (prints redirected to GUI log)
# ============================================================================
def setup_driver():
    print("Setting up Microsoft Edge driver...")
    os.makedirs(download_dir, exist_ok=True)

    try:
        driver = create_edge_driver(log=print)
        try:
            driver.maximize_window()
        except Exception:
            pass
        print("  Edge driver initialized via attach-to-open pattern.")
        return driver
    except Exception as e:
        print(f"  Attach-to-open Edge failed: {e}")
        print("  Falling back to standalone driver...")

    edge_version = _get_installed_edge_version()
    if edge_version:
        print(f"  Installed Edge version: {edge_version}")

    try:
        print("  Trying Selenium Manager (auto-download)...")
        driver = webdriver.Edge(options=_build_edge_options())
        driver.maximize_window()
        print("  Edge driver initialized via Selenium Manager.")
        return driver
    except Exception as e:
        print(f"  Selenium Manager failed: {e}")
        return None


def login(driver, manual_login_confirm):
    print(f"\nNavigating to {LOGIN_URL}...")
    driver.get(LOGIN_URL)
    time.sleep(3)

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        username_field = driver.find_element(
            By.CSS_SELECTOR,
            "input[type='email'], input[name='username'], input[id='username'], input[name='email'], input[id='email']"
        )
        username_field.clear()
        username_field.send_keys(USERNAME)
        print("Username entered")

        password_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        password_field.clear()
        password_field.send_keys(PASSWORD)
        print("Password entered")

        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btnLogin"))
        )
        login_button.click()
        print("Login button clicked")

        time.sleep(5)
        print(f"Current URL after login: {driver.current_url}")
        return True

    except Exception as e:
        print(f"Login failed: {e}")
        print("\nAttempting manual login...")
        print("Please log in manually in the browser window.")
        manual_login_confirm()
        return True


def navigate_to_my_account(driver):
    print("\nNavigating to My Account...")
    try:
        my_account_selectors = [
            "a[href='dashboard']",
            "a[href*='dashboard']",
            "//a[contains(@href, 'dashboard')]",
            "//a[contains(text(), 'My Account')]"
        ]
        for selector in my_account_selectors:
            try:
                if selector.startswith("//"):
                    my_account = driver.find_element(By.XPATH, selector)
                else:
                    my_account = driver.find_element(By.CSS_SELECTOR, selector)
                my_account.click()
                print("Clicked My Account link")
                time.sleep(3)
                return True
            except Exception:
                continue
        print("Could not find My Account link automatically")
        return False
    except Exception as e:
        print(f"Error navigating to My Account: {e}")
        return False


# ----------------------------------------------------------------------------
# INLINE TRACKING EXTRACTION
# ----------------------------------------------------------------------------
TRACKING_PATTERNS = [
    re.compile(r"\b(1Z\s*[0-9A-Z]{3}\s*[0-9A-Z]{3}\s*[0-9A-Z]{2}\s*[0-9A-Z]{6}\s*[0-9A-Z]{2})\b"),
    re.compile(r"\b(1Z[0-9A-Z]{16})\b"),
    re.compile(r"\b([A-Z]{2}\d{9}[A-Z]{2})\b"),
    re.compile(r"\b(\d{20,22})\b"),
    re.compile(r"\b(\d{15,16})\b"),
    re.compile(r"\b([A-Z]{3}\d{6,})\b"),
]

CARRIER_KEYWORDS = ("fedex", "ups.com", "usps", "dhl", "ontrac", "lasership")


def _extract_tracking_from_row(row):
    try:
        for a in row.find_elements(By.TAG_NAME, "a"):
            href = (a.get_attribute("href") or "").strip()
            if not href:
                continue
            href_lower = href.lower()
            if any(k in href_lower for k in CARRIER_KEYWORDS):
                for pat in TRACKING_PATTERNS:
                    m = pat.search(href)
                    if m:
                        return m.group(1).strip()
                txt = (a.text or "").strip()
                if txt and re.match(r"^[A-Z0-9][A-Z0-9\s-]{7,}$", txt, flags=re.IGNORECASE):
                    return re.sub(r"\s+", "", txt)
    except Exception:
        pass

    try:
        labeled = row.find_elements(
            By.XPATH,
            ".//*[contains(translate(@id, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tracking') "
            "or contains(translate(@class, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tracking')]"
        )
        for el in labeled:
            txt = (el.text or "").strip()
            if not txt:
                continue
            cleaned = re.sub(r"^.*?tracking\s*(number)?\s*[:]?\s*", "", txt, flags=re.IGNORECASE).strip()
            if not cleaned or len(cleaned) < 6:
                continue
            for pat in TRACKING_PATTERNS:
                m = pat.search(cleaned)
                if m:
                    return m.group(1).strip()
            if 8 <= len(cleaned) <= 35 and re.match(r"^[A-Z0-9][A-Z0-9\- ]+$", cleaned, flags=re.IGNORECASE):
                return re.sub(r"\s+", "", cleaned)
    except Exception:
        pass

    try:
        order_id_text = ""
        try:
            id_cells = row.find_elements(By.TAG_NAME, "td")
            if len(id_cells) >= 2:
                order_id_text = (id_cells[1].text or "").strip()
        except Exception:
            pass

        for cell in row.find_elements(By.TAG_NAME, "td"):
            txt = (cell.text or "").strip()
            if not txt:
                continue
            if txt.startswith("$") or re.match(r"^\d{1,4}[-/]\d{1,2}[-/]\d{2,4}", txt):
                continue
            if order_id_text and txt == order_id_text:
                continue
            for pat in TRACKING_PATTERNS:
                m = pat.search(txt)
                if m:
                    candidate = m.group(1).strip()
                    if candidate and candidate != order_id_text:
                        return candidate
    except Exception:
        pass

    return ""


# ----------------------------------------------------------------------------
# DATE RANGE FILTERING
# ----------------------------------------------------------------------------
def parse_order_date(date_str):
    if not date_str:
        return None
    s = date_str.strip()
    for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def get_date_range(mode, from_str, to_str):
    today = datetime.now().replace(hour=23, minute=59, second=59, microsecond=0)
    if mode == "all":
        return None, None
    if mode == "7d":
        return today - timedelta(days=7), today
    if mode == "30d":
        return today - timedelta(days=30), today
    if mode == "90d":
        return today - timedelta(days=90), today
    if mode == "ytd":
        return datetime(today.year, 1, 1, 0, 0, 0), today
    if mode == "custom":
        start = None
        end = today
        if from_str:
            for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
                try:
                    start = datetime.strptime(from_str.strip(), fmt).replace(hour=0, minute=0, second=0)
                    break
                except ValueError:
                    pass
        if to_str:
            for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
                try:
                    end = datetime.strptime(to_str.strip(), fmt).replace(hour=23, minute=59, second=59)
                    break
                except ValueError:
                    pass
        return start, end
    return None, None


def filter_orders_by_date(orders, start_date, end_date):
    if start_date is None and end_date is None:
        return orders
    out = []
    for o in orders:
        d = parse_order_date(o.get("Order Date", ""))
        if d is None:
            out.append(o)
            continue
        if start_date and d < start_date:
            continue
        if end_date and d > end_date:
            continue
        out.append(o)
    return out


# ----------------------------------------------------------------------------
# PDF-BASED TRACKING EXTRACTION
# ----------------------------------------------------------------------------
PDF_TRACKING_PATTERN = re.compile(r"\b(1Z[0-9A-Z]{6,})\b")
PDF_TRACKING_PATTERN_GENERIC = re.compile(
    r"\b(1Z[0-9A-Z]{6,}"
    r"|[A-Z]{2}\d{9}[A-Z]{2}"
    r"|\d{20,22}"
    r"|\d{15,16})\b"
)


def _extract_tracking_from_pdf(pdf_source):
    try:
        import pdfplumber
    except ImportError:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pdfplumber", "-q"])
            import pdfplumber
        except Exception as e:
            print(f"    [WARN] pdfplumber not available: {e}")
            return ""

    try:
        if isinstance(pdf_source, (str, bytes)) and not isinstance(pdf_source, bytes):
            opener = lambda: pdfplumber.open(pdf_source)
        else:
            opener = lambda: pdfplumber.open(io.BytesIO(pdf_source.read() if hasattr(pdf_source, "read") else pdf_source))

        all_text = ""
        with opener() as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                all_text += page_text + "\n"
    except Exception as e:
        print(f"    [WARN] PDF parse failed: {e}")
        return ""

    if not all_text.strip():
        return ""

    ups_matches = PDF_TRACKING_PATTERN.findall(all_text)
    if ups_matches:
        return ups_matches[0].strip()

    generic_matches = PDF_TRACKING_PATTERN_GENERIC.findall(all_text)
    if generic_matches:
        for m in generic_matches:
            if len(m) >= 12:
                return m.strip()

    return ""


def _fetch_one_pdf_tracking(order, session):
    url = order["_invoice_link"]
    try:
        resp = session.get(url, timeout=30)
        if resp.status_code != 200:
            return order, "", f"HTTP {resp.status_code}"
        tracking = _extract_tracking_from_pdf(io.BytesIO(resp.content))
        return order, tracking, ""
    except Exception as e:
        return order, "", str(e)


def fetch_tracking_from_invoices(orders_data, ask_yes_no, should_stop, max_workers=8):
    candidates = [
        o for o in orders_data
        if not o.get("Tracking Number")
        and o.get("_invoice_link")
        and o["_invoice_link"].startswith("http")
    ]
    if not candidates:
        already_have = sum(1 for o in orders_data if o.get("Tracking Number"))
        print(f"All {already_have}/{len(orders_data)} orders already have tracking numbers — no PDF fetch needed.")
        return

    est_seconds = max(1, (len(candidates) * 2) // max(1, max_workers))
    do_fetch = ask_yes_no(
        "Fetch Tracking Numbers from Invoices",
        f"{len(candidates)} order(s) need tracking numbers fetched from invoice PDFs.\n\n"
        f"Fetch them in parallel ({max_workers} workers, ~{est_seconds}s estimated)?\n"
        f"No PDF files will be saved to disk."
    )
    if not do_fetch:
        print("Skipped PDF-based tracking fetch.")
        return

    try:
        import requests
    except ImportError:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "requests", "-q"])
            import requests
        except Exception as e:
            print(f"[ERROR] requests library not available: {e}")
            print("        Cannot fetch tracking numbers from PDFs without it.")
            return

    print(f"\nFetching tracking from {len(candidates)} invoice PDFs in parallel ({max_workers} workers)...")
    fetched = 0
    t0 = time.time()

    session = requests.Session()
    adapter = requests.adapters.HTTPAdapter(pool_connections=max_workers, pool_maxsize=max_workers * 2)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({"User-Agent": "Mozilla/5.0"})

    total = len(candidates)
    completed = 0
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(_fetch_one_pdf_tracking, o, session): o for o in candidates}
        for fut in as_completed(futures):
            if should_stop():
                for f in futures:
                    f.cancel()
                print("  Stop requested — aborting parallel PDF fetch.")
                break
            try:
                order, tracking, err = fut.result()
            except Exception as e:
                order = futures[fut]
                tracking, err = "", str(e)
            completed += 1
            order_id = order.get("Order ID", "?")
            if tracking:
                order["Tracking Number"] = tracking
                fetched += 1
                print(f"  [{completed}/{total}] {order_id}: tracking = {tracking}")
            elif err:
                print(f"  [{completed}/{total}] {order_id}: no tracking ({err})")
            else:
                print(f"  [{completed}/{total}] {order_id}: no tracking in PDF")

    elapsed = time.time() - t0
    total_with_tracking = sum(1 for o in orders_data if o.get("Tracking Number"))
    print(f"\nFetched {fetched}/{len(candidates)} tracking numbers from PDFs in {elapsed:.1f}s")
    print(f"Total orders with tracking: {total_with_tracking}/{len(orders_data)}")


def extract_order_history(driver):
    print("\nExtracting order history...")
    orders_data = []
    try:
        if "dashboard" not in driver.current_url.lower():
            navigate_to_my_account(driver)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.table-striped, table.table"))
        )
        table = driver.find_element(By.CSS_SELECTOR, "table.table-striped, table.table")
        rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
        print(f"Found {len(rows)} order rows")

        for idx, row in enumerate(rows, 1):
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 6:
                order_id = cells[1].text.strip()
                order_date = cells[2].text.strip()
                amount = cells[3].text.strip()
                order_status = cells[5].text.strip()

                invoice_link = ""
                try:
                    invoice_elem = cells[4].find_element(By.TAG_NAME, "a")
                    invoice_link = invoice_elem.get_attribute("href")
                except Exception:
                    pass

                view_link = ""
                try:
                    view_elem = cells[6].find_element(By.XPATH, ".//a[contains(text(), 'View Order')]")
                    view_link = view_elem.get_attribute("href")
                except Exception:
                    pass

                tracking_number = _extract_tracking_from_row(row)

                order = {
                    "#": idx,
                    "Order ID": order_id,
                    "Order Date": order_date,
                    "Amount": amount,
                    "Order Status": order_status,
                    "Tracking Number": tracking_number,
                    "_invoice_link": invoice_link,
                    "_view_order_link": view_link,
                }
                orders_data.append(order)
                if tracking_number:
                    print(f"  Extracted order #{idx}: {order_id}  |  {order_date}  |  {amount}  |  {order_status}  |  tracking={tracking_number}")
                else:
                    print(f"  Extracted order #{idx}: {order_id}  |  {order_date}  |  {amount}  |  {order_status}  |  tracking=(none)")

        print(f"\nSuccessfully extracted {len(orders_data)} orders")
        return orders_data

    except Exception as e:
        print(f"Error extracting orders: {e}")
        print(f"Current URL: {driver.current_url}")
        return []


# ============================================================================
# EXCEL EXPORT (professional styling)
# ============================================================================
def save_to_excel(orders_data):
    _ensure_heavy_imports()
    if not orders_data:
        print("\nNo order data to save")
        return None

    export_rows = []
    for o in orders_data:
        export_rows.append({k: v for k, v in o.items() if not k.startswith("_")})
    df_export = pd.DataFrame(export_rows)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"cpwh_order_history_{timestamp}.xlsx"
    filepath = os.path.join(download_dir, filename)

    NAVY_FILL = PatternFill(start_color="FF161632", end_color="FF161632", fill_type="solid")
    ROW_ALT_FILL = PatternFill(start_color="FFF2F4F7", end_color="FFF2F4F7", fill_type="solid")
    ROW_WHITE_FILL = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type="solid")
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
    BODY_FONT = Font(name="Calibri", size=10, color="FF1C1C1C")
    ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    THIN_BORDER = Border(
        left=Side(style="thin", color="FFD0D5DD"),
        right=Side(style="thin", color="FFD0D5DD"),
        top=Side(style="thin", color="FFD0D5DD"),
        bottom=Side(style="thin", color="FFD0D5DD"),
    )

    with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
        df_export.to_excel(writer, sheet_name='Order History', index=False)
        worksheet = writer.sheets['Order History']

        max_row = worksheet.max_row
        max_col = worksheet.max_column

        for col_idx in range(1, max_col + 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill = NAVY_FILL
            cell.font = HEADER_FONT
            cell.alignment = ALIGN_CENTER
            cell.border = THIN_BORDER
        worksheet.row_dimensions[1].height = 26

        center_cols = {"#"}
        for r in range(2, max_row + 1):
            row_fill = ROW_ALT_FILL if (r % 2 == 0) else ROW_WHITE_FILL
            for c in range(1, max_col + 1):
                cell = worksheet.cell(row=r, column=c)
                cell.fill = row_fill
                cell.font = BODY_FONT
                cell.border = THIN_BORDER
                header_val = worksheet.cell(row=1, column=c).value
                if header_val in center_cols:
                    cell.alignment = ALIGN_CENTER
                else:
                    cell.alignment = ALIGN_LEFT
            worksheet.row_dimensions[r].height = 18

        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            max_length = 0
            for r in range(1, max_row + 1):
                val = worksheet.cell(row=r, column=col_idx).value
                if val is None:
                    continue
                length = sum(2 if ord(ch) > 127 else 1 for ch in str(val))
                if length > max_length:
                    max_length = length
            adjusted_width = min(max(max_length + 3, 10), 45)
            worksheet.column_dimensions[col_letter].width = adjusted_width

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
        worksheet.sheet_properties.tabColor = "161632"

    print(f"\n{'='*60}")
    print("DATA SAVED SUCCESSFULLY!")
    print(f"{'='*60}")
    print(f"File location: {filepath}")
    print(f"Total orders:  {len(orders_data)}")
    print(f"Columns:       {', '.join(df_export.columns)}")
    print(f"{'='*60}")

    print("\nORDER SUMMARY:")
    summary_cols = [c for c in ["Order ID", "Order Date", "Amount", "Order Status", "Tracking Number"] if c in df_export.columns]
    print(df_export[summary_cols].to_string(index=False))

    return filepath


# ============================================================================
# MAIN SCRAPE ROUTINE
# ============================================================================
def run_scrape(log, ask_yes_no, ask_ok, should_stop, date_range=("all", "", ""), max_workers=8):
    _ensure_heavy_imports()
    driver = None
    try:
        driver = setup_driver()
        if not driver:
            log("Driver setup failed — see log above.")
            return

        def manual_login_confirm():
            ask_ok("Manual Login Required",
                   "Please complete login manually in the browser window,\nthen click OK to continue.")

        if login(driver, manual_login_confirm):
            print("\nLogin successful!")
            orders = extract_order_history(driver)

            if orders:
                mode, from_str, to_str = date_range
                start_date, end_date = get_date_range(mode, from_str, to_str)
                if start_date is not None or end_date is not None:
                    before = len(orders)
                    orders = filter_orders_by_date(orders, start_date, end_date)
                    print(f"\nDate filter ({mode}): {before} -> {len(orders)} orders in range")
                    if not orders:
                        print("No orders fall in the selected date range. Nothing to export.")
                        return
                    for idx, o in enumerate(orders, 1):
                        o["#"] = idx

                fetch_tracking_from_invoices(orders, ask_yes_no, should_stop, max_workers=max_workers)
                save_to_excel(orders)
            else:
                print("\nNo orders found in the order history.")
                print("The browser will stay open for you to check manually.")
        else:
            print("\nLogin had issues. Browser will stay open.")

        print("\n" + "="*60)
        print("Browser will remain open for 15 seconds")
        print("="*60)
        for i in range(15, 0, -1):
            if should_stop():
                break
            time.sleep(1)

    except Exception as e:
        print(f"\nError: {e}")
        print(traceback.format_exc())
    finally:
        if driver:
            try:
                driver.quit()
                print("\nBrowser closed.")
            except Exception:
                pass


# ==========================================================
# QT UI LAYER
# ==========================================================
from PySide6.QtCore import Qt, QObject, Signal, QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QRadioButton, QButtonGroup, QLineEdit, QComboBox,
    QPlainTextEdit, QMessageBox, QGroupBox
)

from header_manager_qt import FixedHeaderManagerQt

ICON_ICO_B64 = (
    open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "icon_ico_b64.txt"), "r").read().strip()
    if not getattr(sys, "frozen", False)
    else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "icon_ico_b64.txt"), "r").read().strip()
)


def _extract_embedded_icon(b64, filename):
    try:
        if not b64:
            return None
        import base64 as _b64, tempfile
        target = os.path.join(tempfile.gettempdir(), filename)
        with open(target, "wb") as fh:
            fh.write(_b64.b64decode(b64))
        return target if os.path.isfile(target) else None
    except Exception:
        return None


def _set_window_icon_qt(window):
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        ico_path = os.path.join(meipass, ICON_ICO_NAME)
        if os.path.exists(ico_path):
            window.setWindowIcon(QIcon(ico_path))
            return
    base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else get_script_dir()
    ico_path = os.path.join(base_dir, ICON_ICO_NAME)
    if os.path.exists(ico_path):
        window.setWindowIcon(QIcon(ico_path))
        return
    try:
        ico_path = _extract_embedded_icon(EMBEDDED_ICON_B64, "gfh_app_icon.ico")
        if ico_path:
            window.setWindowIcon(QIcon(ico_path))
            return
    except Exception:
        pass
    try:
        ico_path = _extract_embedded_icon(ICON_ICO_B64, "gfh_app_icon2.ico")
        if ico_path:
            window.setWindowIcon(QIcon(ico_path))
    except Exception:
        pass


class UIBridge(QObject):
    log = Signal(str)
    done = Signal()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        _set_window_icon_qt(self)
        self._stop_flag = False
        self._busy = False

        self.theme_manager = ThemeManagerQt(default="dark", app_name="GFH-Accessories-Order-History-Scraper")

        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.header_mgr = FixedHeaderManagerQt(outer, title="GFH Accessories Order History Scraper")
        logo_path = _resource_path(WORDMARK_PNG_NAME)
        if os.path.exists(logo_path):
            self.header_mgr.set_logo(logo_path=logo_path, text="GFH")
        self.header_mgr.add_theme_toggle(self.theme_manager, callback=self._apply_theme)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(20, 12, 20, 6)
        outer.addWidget(body)

        info_row = QHBoxLayout()
        info_row.addWidget(QLabel(f"Login user: {USERNAME}"))
        info_row.addWidget(QLabel(f"    Excel output: {download_dir}"))
        info_row.addStretch(1)
        body_layout.addLayout(info_row)

        date_box = QGroupBox("Date Range Filter")
        date_layout = QVBoxLayout(date_box)
        row1 = QHBoxLayout()
        self.date_group = QButtonGroup(self)
        presets = [
            ("All Dates", "all"), ("7 Days", "7d"), ("30 Days", "30d"),
            ("90 Days", "90d"), ("This Year", "ytd"), ("Custom", "custom"),
        ]
        self._date_radios = {}
        for label, value in presets:
            rb = QRadioButton(label)
            if value == "all":
                rb.setChecked(True)
            rb.toggled.connect(self._on_date_range_change)
            row1.addWidget(rb)
            self.date_group.addButton(rb)
            self._date_radios[value] = rb
        row1.addStretch(1)
        date_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("From:"))
        self.from_entry = QLineEdit()
        self.from_entry.setFixedWidth(100)
        self.from_entry.setEnabled(False)
        row2.addWidget(self.from_entry)
        row2.addWidget(QLabel("To:"))
        self.to_entry = QLineEdit()
        self.to_entry.setFixedWidth(100)
        self.to_entry.setEnabled(False)
        row2.addWidget(self.to_entry)
        hint = QLabel("(MM/DD/YYYY)")
        hint.setStyleSheet("color: #888;")
        row2.addWidget(hint)
        row2.addStretch(1)
        row2.addWidget(QLabel("Workers:"))
        self.workers_combo = QComboBox()
        self.workers_combo.addItems([str(n) for n in (2, 4, 6, 8, 10, 12, 16)])
        self.workers_combo.setCurrentText("8")
        row2.addWidget(self.workers_combo)
        date_layout.addLayout(row2)
        body_layout.addWidget(date_box)

        btn_row = QHBoxLayout()
        self.start_btn = QPushButton("Start Scrape")
        self.start_btn.clicked.connect(self._start)
        btn_row.addWidget(self.start_btn)
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.clicked.connect(self._stop)
        self.stop_btn.setEnabled(False)
        btn_row.addWidget(self.stop_btn)
        self.vpn_btn = QPushButton("Open Edge Browser")
        self.vpn_btn.clicked.connect(self._open_edge_browser)
        btn_row.addWidget(self.vpn_btn)
        btn_row.addStretch(1)
        body_layout.addLayout(btn_row)

        log_wrap = QWidget()
        log_layout = QVBoxLayout(log_wrap)
        log_layout.setContentsMargins(20, 4, 20, 6)
        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("background: #0d0d1f; color: #d7e3f0; font-family: Consolas;")
        log_layout.addWidget(self.log_box)
        outer.addWidget(log_wrap, 1)

        self.header_mgr.add_copyright(self.theme_manager, outer)

        self._bridge = UIBridge()
        self._bridge.log.connect(self._append_log)
        self._bridge.done.connect(self._on_done)

        self._apply_theme()
        self.resize(1200, 820)
        QTimer.singleShot(0, self.showMaximized)

    def _apply_theme(self):
        self.theme_manager.apply_theme_to_app(QApplication.instance())
        self.header_mgr.update_for_theme()

    def _on_date_range_change(self):
        custom = self._date_radios["custom"].isChecked()
        self.from_entry.setEnabled(custom)
        self.to_entry.setEnabled(custom)

    def _log(self, msg):
        self._bridge.log.emit(msg)

    def _append_log(self, msg):
        if msg and msg.strip("\r\n") != "":
            self.log_box.appendPlainText(f"[{datetime.now():%H:%M:%S}] {msg.rstrip(chr(10))}")

    def _ask_yes_no(self, title, message):
        result_holder = {}
        event = threading.Event()

        def _show():
            result_holder["v"] = (QMessageBox.question(self, title, message) == QMessageBox.Yes)
            event.set()

        QTimer.singleShot(0, _show)
        event.wait()
        return result_holder.get("v", False)

    def _ask_ok(self, title, message):
        event = threading.Event()

        def _show():
            QMessageBox.information(self, title, message)
            event.set()

        QTimer.singleShot(0, _show)
        event.wait()

    def _should_stop(self):
        return self._stop_flag

    def _open_edge_browser(self):
        if is_port_open():
            QMessageBox.information(self, "Edge Already Open",
                                     "The automation Edge browser is already running.\n"
                                     "You can click Start Scrape now.")
            return
        self._log("Opening dedicated Edge automation browser...")
        if open_vpn_setup_browser(log=self._log):
            self._log("Edge is ready. Click Start Scrape to begin.")
        else:
            QMessageBox.warning(self, "Edge Failed",
                                 "Could not open the automation Edge browser.\n"
                                 "Make sure Microsoft Edge is installed.")

    def _start(self):
        if self._busy:
            return
        self._busy = True
        self._stop_flag = False
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.log_box.setPlainText("")

        mode = next(v for v, rb in self._date_radios.items() if rb.isChecked())
        if mode == "custom":
            from_str = self.from_entry.text().strip()
            to_str = self.to_entry.text().strip()
        else:
            from_str, to_str = "", ""
        try:
            workers = int(self.workers_combo.currentText())
        except Exception:
            workers = 8
        workers = max(1, min(workers, 32))

        class _PrintRedirector:
            def __init__(self, emit):
                self._emit = emit

            def write(self, text):
                if text and text.strip("\r\n") != "":
                    self._emit(text.rstrip("\n"))

            def flush(self):
                pass

        sys.stdout = _PrintRedirector(self._log)

        def _worker():
            try:
                run_scrape(
                    self._log, self._ask_yes_no, self._ask_ok, self._should_stop,
                    date_range=(mode, from_str, to_str),
                    max_workers=workers,
                )
            finally:
                sys.stdout = sys.__stdout__
                self._bridge.done.emit()

        threading.Thread(target=_worker, daemon=True).start()

    def _stop(self):
        self._stop_flag = True
        self._log("Stop requested — will halt after the current step.")

    def _on_done(self):
        self._busy = False
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    def closeEvent(self, event):
        self._stop_flag = True
        event.accept()


def _enable_dpi_awareness() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _close_pyi_splash():
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def main():
    _enable_dpi_awareness()
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("GFHTelecom.AccessoriesOrderHistoryScraper")
    except Exception:
        pass
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    _close_pyi_splash()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
