#uberone.py
"""
UBER EATS — EXTENDED (Phase 2): signed-in Uber One availability check.

STANDALONE. It does NOT import or modify main.py / scraper.py / excel_formatter.py.
It only reuses two PURE helpers (read-only) from the existing project:
    timezone_utils.extract_store_id, determine_timezone_and_state
so the approved Phase-1 code is left completely untouched.

WHAT IT DOES, per store row in input.xlsx:
    1. open the store on a Chrome profile that is ALREADY logged in to Uber Eats
    2. check for the Uber One icon  -> column  U1 Icon   ("U1icon_present" / "")
    3. set delivery location = address_2mile, reload, look for an estimated
       delivery-time line -> column  2mile   ("Available" / "Not Available")
    4. set delivery location = address_5mile, same -> column  5mile
    5. next store

WHY A DEDICATED, LOGGED-IN PROFILE
    Uber One status and the delivery-time line only render when signed in. Phase 1
    used a throwaway profile (mkdtemp) that is never logged in - useless here. This
    script points Chrome at a PERSISTENT profile folder (CHROME_PROFILE_DIR) that
    you log into ONCE by hand; every later run reuses those cookies.

FIRST-TIME SETUP (do this once)
    python uberone.py --login
    -> Chrome opens on ubereats.com in the dedicated profile. Sign in manually
       (solve any OTP/captcha), confirm you can see your account, then close it.
    From then on:  python uberone.py   (uses the saved session)

ACCOUNT RISK
    This drives a logged-in account and changes its delivery address many times per
    run. Uber may flag/suspend such an account. Run it on a dedicated account the
    client accepts that risk for. Pacing lowers the risk; it does not remove it.

SELECTORS
    Uber's location dialog uses rotating hashed CSS classes, so every selector
    below is text-anchored / structural and there are SEVERAL fallbacks per step.
    The exact live flow is confirmed on the first test run - each step logs what it
    saw to uberone.log, which is the place to look when a step misses.
"""

import os
import re
import sys
import json
import time
import random
import logging
import argparse
import tempfile
from datetime import datetime

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# read-only reuse of pure helpers from the approved code (no edits to those files)
from timezone_utils import extract_store_id, determine_timezone_and_state

# =============================================================================
# CONFIG
# =============================================================================
# -----------------------------------------------------------------------------
# BROWSER SELECTION - resolved at RUNTIME, never baked into the .exe.
#
# This ships as a frozen .exe, so the source cannot be edited after packaging.
# The browser is therefore decided each run, in this order (first that applies):
#   1.  --browser chrome|edge          (command-line flag)
#   2.  UE_BROWSER environment variable
#   3.  settings.ini placed NEXT TO the .exe:   [browser]\n name = chrome
#   4.  DEFAULT_BROWSER below, and if that browser is not installed the launcher
#       automatically falls back to the other one.
#
# Default is Chrome: it is the portable choice for a later move to cloud / Mac
# (Edge headless on Linux is a pain), while a Windows PC without Chrome silently
# falls back to Edge, which is always present. So the same .exe "just works" on a
# Mac, a cloud box, or any of the 20+ Windows machines with no per-machine edit.
# -----------------------------------------------------------------------------
DEFAULT_BROWSER = "chrome"        # "chrome" | "edge"

PROFILE_ROOT = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
                else os.path.dirname(os.path.abspath(__file__)))


def _app_dir():
    return PROFILE_ROOT


def resolve_browser(cli_value=None):
    """Decide the PREFERRED browser at runtime (see order above)."""
    if cli_value and cli_value.strip().lower() in ("chrome", "edge"):
        return cli_value.strip().lower()
    env = os.environ.get("UE_BROWSER", "").strip().lower()
    if env in ("chrome", "edge"):
        return env
    try:
        import configparser
        ini = os.path.join(_app_dir(), "settings.ini")
        if os.path.exists(ini):
            cp = configparser.ConfigParser()
            cp.read(ini)
            v = cp.get("browser", "name", fallback="").strip().lower()
            if v in ("chrome", "edge"):
                return v
    except Exception:
        pass
    return DEFAULT_BROWSER


def profile_dir_for(browser):
    """One persistent, logged-in profile folder per browser, next to the program
    (so switching browsers never mixes a Chrome session into Edge)."""
    return os.path.join(_app_dir(), f"ue_profile_{browser}")


# =============================================================================
# REAL CHROME PROFILE (loaded IN PLACE, selected by the sheet's e-mail)
# =============================================================================
# The tool opens the team member's ACTUAL Chrome profile - the one already signed
# in to Uber Eats, with their extensions etc. - by pointing Chrome straight at the
# real "User Data" folder and the matching profile sub-folder. Nothing is copied
# or moved, so Chrome does NOT show "verify it's you". The only requirement is that
# all Chrome windows are CLOSED first (Chrome locks a profile that is open).
#
# The e-mail -> folder map comes from "User Data\Local State" (profile.info_cache),
# where each profile records the Google account (user_name) it is signed in to.
# =============================================================================

def find_chrome_user_data_dir():
    """Locate Chrome's real 'User Data' folder for the current OS, or None."""
    home = os.path.expanduser("~")
    plat = sys.platform
    if plat.startswith("win"):
        local = os.environ.get("LOCALAPPDATA", os.path.join(home, "AppData", "Local"))
        candidates = [os.path.join(local, "Google", "Chrome", "User Data")]
    elif plat == "darwin":
        candidates = [os.path.join(home, "Library", "Application Support", "Google", "Chrome")]
    else:
        candidates = [os.path.join(home, ".config", "google-chrome"),
                      os.path.join(home, ".config", "chromium")]
    for c in candidates:
        if os.path.isdir(c):
            return c
    return None


def list_profile_emails(user_data_dir):
    """Return {email_lower: profile_folder_name} from Local State's info_cache."""
    out = {}
    try:
        with open(os.path.join(user_data_dir, "Local State"), "r",
                  encoding="utf-8", errors="ignore") as f:
            ls = json.load(f)
        cache = ls.get("profile", {}).get("info_cache", {}) or {}
        for folder, info in cache.items():
            email = (info.get("user_name") or "").strip().lower()
            if email:
                out[email] = folder
    except Exception as e:
        log.warning(f"could not read Chrome Local State: {e}")
    return out


def find_profile_for_email(user_data_dir, email):
    """Return (profile_folder_or_None, {email: folder} mapping)."""
    mapping = list_profile_emails(user_data_dir)
    return mapping.get((email or "").strip().lower()), mapping


def resolve_real_profile(email):
    """Return (real_user_data_dir, profile_folder) for the Chrome profile matching
    `email`, to be opened IN PLACE. Blank e-mail -> the 'Default' profile.
    Raises RuntimeError with actionable guidance when it cannot be resolved.
    """
    user_data = find_chrome_user_data_dir()
    if not user_data:
        raise RuntimeError("Could not find Chrome's 'User Data' folder on this PC. "
                           "Is Google Chrome installed for this Windows user?")
    email = (email or "").strip()
    if not email:
        return user_data, "Default"
    folder, mapping = find_profile_for_email(user_data, email)
    if not folder:
        known = ", ".join(sorted(mapping.keys())) if mapping else \
            "(none - no Chrome profile is signed in with a Google account)"
        raise RuntimeError(
            f"No Chrome profile on this PC is signed in with '{email}'.\n"
            f"       Chrome profiles found: {known}\n"
            f"       Fix: open Chrome, pick the profile logged in to Uber Eats, make sure\n"
            f"       it is signed in to Chrome with '{email}' (Settings - You and Google),\n"
            f"       put a matching e-mail in input.xlsx, or leave it blank to use Default.")
    return user_data, folder

INPUT_FILE   = "input.xlsx"
STORE_URL_COL = "website_url"
ADDR_2MI_COL  = "address_2mile"
ADDR_5MI_COL  = "address_5mile"

PAGE_LOAD_TIMEOUT = 60
SETTLE_SECS       = 5        # after a navigation / reload
BETWEEN_STORES    = (6, 11)  # random polite gap (seconds) between stores
BETWEEN_ADDR      = (3, 6)   # gap between the 2-mile and 5-mile checks

# "Available" is defined by the client as: the reloaded page shows an ESTIMATED
# DELIVERY TIME. So the signal is strictly a minutes-based ETA - nothing vaguer.
# "Schedule delivery for later" is deliberately NOT counted: it is the option
# shown when order-now delivery is not offered, i.e. the opposite of available.
_DELIVERY_ESTIMATE_PATTERNS = [
    re.compile(r'\b\d{1,3}\s*[-–]\s*\d{1,3}\s*min\b', re.I),          # "25 - 40 min"
    re.compile(r'\bearliest arrival[^\n]{0,20}\d{1,3}\s*min\b', re.I),# "Earliest arrival 35 min"
    re.compile(r'\b\d{1,3}\s*min\b', re.I),                           # "35 min"
]
# Text that means the address is out of range / not served.
_NOT_AVAILABLE_MARKERS = (
    "delivery unavailable", "unavailable for delivery", "outside of delivery range",
    "outside the delivery area", "out of delivery range",
    "not available", "no longer available", "out of range",
    "too far", "doesn't deliver here", "does not deliver",
    "not delivering to", "no stores", "nothing available",
)

# ---- logging -----------------------------------------------------------------
_log_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
    else os.path.dirname(os.path.abspath(__file__))

# Module logger only. We do NOT configure the root logger at import time - that
# would hijack logging when main.py imports this module (duplicate console lines
# + a stray uberone.log). When run standalone, _setup_standalone_logging() wires
# the handlers; when imported, log records propagate to whatever main.py set up
# (so Phase-2 messages land in scraper.log alongside everything else).
log = logging.getLogger("ue.uberone")


def _setup_standalone_logging():
    """Attach file+console handlers - only when this module is run on its own."""
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    fh = logging.FileHandler(os.path.join(_log_dir, "uberone.log"))
    fh.setFormatter(fmt)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    root.addHandler(fh)
    root.addHandler(ch)


# =============================================================================
# DRIVER  (dedicated, persistent, logged-in profile)
# =============================================================================
def _launch(browser, headless):
    """Launch ONE browser. Raises if it (or its driver) can't start."""
    is_edge = browser == "edge"
    profile = profile_dir_for(browser)
    os.makedirs(profile, exist_ok=True)

    opts = EdgeOptions() if is_edge else ChromeOptions()
    opts.add_argument(f"--user-data-dir={profile}")     # <- login persistence
    opts.add_argument("--profile-directory=Default")
    opts.add_argument("--window-size=1440,1000")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    if headless:
        opts.add_argument("--headless=new")

    # A driver bundled next to the program wins (offline / locked-down PCs);
    # otherwise Selenium Manager fetches the matching driver automatically.
    driver_name = "msedgedriver.exe" if is_edge else "chromedriver.exe"
    local_driver = os.path.join(_app_dir(), driver_name)
    Launcher = webdriver.Edge if is_edge else webdriver.Chrome
    ServiceCls = EdgeService if is_edge else ChromeService
    if os.path.exists(local_driver):
        return Launcher(service=ServiceCls(local_driver), options=opts), profile
    return Launcher(options=opts), profile


def build_driver(preferred=None, headless=False):
    """Start the PREFERRED browser; if it isn't installed, fall back to the other.

    Returns (driver, browser_used). This is what makes one frozen .exe work
    unchanged on a Mac, a cloud box, or a Windows PC that has only Edge - it
    detects what's actually present instead of failing.

    The fallback only triggers when the preferred browser cannot start at all
    (not installed). A browser that starts but simply isn't logged in is NOT a
    fallback case - ensure_login() handles that and asks for a one-time --login.
    """
    preferred = (preferred or resolve_browser()).strip().lower()
    order = [preferred, "edge" if preferred == "chrome" else "chrome"]

    last_err = None
    for browser in order:
        try:
            driver, profile = _launch(browser, headless)
            if browser != preferred:
                log.warning(f"'{preferred}' unavailable; fell back to '{browser}'.")
            log.info(f"driver started: {browser} | profile={profile}")
            driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
            try:
                driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                    "source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"
                })
            except Exception:
                pass
            return driver, browser
        except Exception as e:
            last_err = e
            log.error(f"could not start {browser}: {e}")
    raise RuntimeError(f"No usable browser (tried {order}). Last error: {last_err}")


# The "Log in" / "Sign up" controls Uber Eats renders ONLY when signed out.
# Their presence on a loaded page is the authoritative 'logged out' signal - that
# is exactly what separates the two header states: signed-out shows these buttons;
# signed-in shows a real delivery address + cart instead. Matched as actual header
# controls (exact button/link text), NOT loose body text, so the words "log in"
# sitting in a footer or marketing block can never cause a false 'logged out'.
_LOGIN_CTA_XPATHS = (
    "//a[normalize-space(.)='Log in']",
    "//button[normalize-space(.)='Log in']",
    "//a[normalize-space(.)='Sign up']",
    "//button[normalize-space(.)='Sign up']",
    "//a[normalize-space(.)='Sign in']",
    "//button[normalize-space(.)='Sign in']",
    "//*[@data-testid='login-button']",
    "//*[@data-testid='signup-button']",
    "//a[contains(@href,'/login')]",
)


def _has_login_cta(driver):
    """True if a VISIBLE 'Log in' / 'Sign up' control is on the page (= signed OUT).

    Visibility is required so a hidden /login link inside a collapsed menu (which
    can exist even while signed in) does not count.
    """
    for xp in _LOGIN_CTA_XPATHS:
        try:
            for el in driver.find_elements(By.XPATH, xp):
                try:
                    if el.is_displayed():
                        return True
                except Exception:
                    return True   # found the control; if visibility is unreadable, treat as present
        except Exception:
            continue
    return False


# URL fragments that mean the browser is sitting on Uber's sign-in / auth flow.
# This is GROUND TRUTH: when a logged-out profile opens a login-gated page, Uber
# redirects here (observed: https://auth.uber.com/v2/?next_url=...login-redirect).
_AUTH_URL_MARKERS = ("auth.uber.com", "accounts.uber.com", "login-redirect",
                     "/login", "signin")


def is_on_login_flow(driver):
    """True while the browser is on Uber's sign-in / auth pages."""
    try:
        cur = (driver.current_url or "").lower()
    except Exception:
        return False
    return any(m in cur for m in _AUTH_URL_MARKERS)


def looks_logged_out(driver):
    """Cheap mid-run check on the page already open (no navigation).

    Only POSITIVE signs of being signed out count (auth URL or a visible
    'Log in' / 'Sign up' control), so a page that simply lacks an account icon
    never stops the run by mistake.
    """
    return is_on_login_flow(driver) or _has_login_cta(driver)


def _has_account_marker(driver):
    """POSITIVE proof of a signed-in session (not merely the absence of a CTA).

    Strongest signal: a 'Sign out' / 'Log out' control exists only when signed in.
    Backed up by an account/profile/cart element in the header.
    """
    try:
        body = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
    except Exception:
        body = ""
    if "sign out" in body or "log out" in body:
        return True
    # NOTE: "Enter delivery address" alone is NOT proof of being signed out - a
    # brand-new account that never saved an address shows it too. So the visible
    # account / cart controls are checked FIRST; a signed-out page has neither.
    try:
        for el in driver.find_elements(
                By.CSS_SELECTOR,
                "[data-testid*='account' i], [aria-label*='account' i], "
                "[aria-label*='profile' i], [data-testid*='cart' i], [aria-label*='cart' i]"):
            try:
                if el.is_displayed():          # hidden menu items don't count
                    return True
            except Exception:
                continue
    except Exception:
        pass
    # no visible account/cart control: signed out (header shows only
    # "Enter delivery address", as seen on the live site)
    return False


def is_logged_in(driver):
    """True ONLY with POSITIVE proof of a real signed-in session.

    Order of checks (strict - a false 'logged in' is what broke the run):
      1. URL is on the auth/sign-in flow  -> NOT logged in (ground truth)
      2. page is blank / failed           -> NOT logged in
      3. visible 'Log in' / 'Sign up' CTA -> NOT logged in
      4. an account marker is present     -> logged in
      5. none of the above (ambiguous)    -> NOT logged in  (so the tool prompts
                                             for a manual sign-in instead of
                                             starting and failing on auth)
    """
    if is_on_login_flow(driver):
        log.info("login check: on auth.uber.com sign-in flow -> NOT logged in")
        return False
    try:
        body = (driver.find_element(By.TAG_NAME, "body").text or "").strip()
    except Exception:
        return False
    if not body:
        return False                      # blank / failed page: cannot claim a session
    if _has_login_cta(driver):
        log.info("login check: 'Log in'/'Sign up' present -> NOT logged in")
        return False
    if _has_account_marker(driver):
        log.info("login check: account/sign-out marker present -> logged in")
        return True
    log.info("login check: no positive session proof -> treating as NOT logged in")
    return False


def ensure_login(driver, url=None):
    """Confirm a REAL signed-in session. Returns bool.

    Checks on `url` when given (the first store URL from the sheet, so the browser
    starts straight on the real page); otherwise on the normal Uber Eats homepage.
    Never navigates to /account or the accounts/sign-in page. is_logged_in() treats
    an auth.uber.com URL (a logged-out store page redirects there) and a visible
    'Log in'/'Sign up' as NOT logged in, and requires a positive account marker to
    confirm a session.
    """
    target = url if (url and str(url).strip().lower() not in ("", "nan")) \
        else "https://www.ubereats.com/"
    try:
        driver.get(target)
        time.sleep(SETTLE_SECS)
    except Exception:
        pass
    if is_logged_in(driver):
        log.info("Session is logged in.")
        return True
    log.warning("Session is NOT logged in - sign in by hand "
                "(the tool will wait), or run `python uberone.py --login` first.")
    return False


# =============================================================================
# UBER ONE ICON
# =============================================================================
# Wording that belongs to Uber One ADVERTISING ("Join Uber One", "Try Uber One
# free") rather than the store's own badge. A signed-in page can show these
# anywhere, so a match whose text/label contains one of them is ignored.
_U1_PROMO_WORDS = ("join", "try ", "free trial", "sign up", "subscribe",
                   "get uber one", "save with", "become a member", "start your")

_U1_BADGE_CSS = ("[data-testid*='uber-one' i], [data-testid*='uberone' i], "
                 "[aria-label*='uber one' i], img[alt*='Uber One' i], "
                 "svg[aria-label*='Uber One' i], [title*='Uber One' i]")
_U1_TEXT_XPATH = ("//*[contains(translate(normalize-space(.),'UBERONE','uberone'),'uber one')]"
                  "[string-length(normalize-space(.))<40]")


def _u1_label(el):
    parts = []
    for attr in ("aria-label", "alt", "title", "data-testid"):
        try:
            v = el.get_attribute(attr)
            if v:
                parts.append(v)
        except Exception:
            pass
    try:
        parts.append(el.text or "")
    except Exception:
        pass
    return " ".join(parts).strip()


def _is_promo(label):
    low = (label or "").lower()
    return any(w in low for w in _U1_PROMO_WORDS)


# Where the Uber One logo sits on a store page: in the first line under the
# store name - "$5.99 Delivery Fee  •  (U1 logo)  •  Info".
# 1st entry = the exact path taken from the live page; the rest are the same
# spot without the brittle positions, in case Uber adds/removes a span.
_U1_ICON_XPATHS = (
    "/html/body/div[1]/div[3]/div[1]/div[2]/main/div/div[1]/section/div[2]/div[1]/p[1]/span[8]/*[name()='svg']",
    "//main//section/div[2]/div[1]/p[1]/span/*[name()='svg']",
    "//main//section//p[contains(., 'Delivery Fee') or contains(., 'delivery fee')]//span/*[name()='svg']",
)


# The Uber One logo's drawing, as seen on the live page (circle with the "1").
# The SAME spot can hold a different icon (seen on a store whose address was not
# set: a path starting "M19.315 2.93"), so only this shape counts as Uber One.
# If Uber redraws the logo, add the new path start here (it is printed in
# scraper.log as "uber-one: OTHER icon ...").
_U1_PATH_STARTS = ("M5.889 21.146",)


def _is_u1_svg(el):
    try:
        html = el.get_attribute("outerHTML") or ""
    except Exception:
        return False
    return any(sig in html for sig in _U1_PATH_STARTS)


def _svg_fingerprint(el):
    """Short description of an <svg> for the log (its first path's shape)."""
    try:
        return (el.get_attribute("outerHTML") or "")[:160].replace("\n", " ")
    except Exception:
        return "?"


def detect_uber_one(driver, wait_secs=2):
    """True if the Uber One logo is shown in the store's fee line.

    Checks the exact spot given by the client first, then the same spot found
    more loosely. Waits a few seconds for it to render. Whatever is found there
    is written to the log so it can be checked against the screenshot.
    """
    deadline = time.time() + wait_secs
    while True:
        for i, xp in enumerate(_U1_ICON_XPATHS):
            try:
                for el in driver.find_elements(By.XPATH, xp):
                    try:
                        if not el.is_displayed():
                            continue
                    except Exception:
                        continue
                    label = _u1_label(el)
                    if _is_promo(label):
                        continue
                    if _is_u1_svg(el):
                        log.info(f"uber-one: logo found (rule {i + 1})")
                        return True
                    log.info(f"uber-one: OTHER icon in that spot (not Uber One): {_svg_fingerprint(el)}")
            except Exception:
                continue
        if time.time() >= deadline:
            break
        time.sleep(0.5)
    log.info("uber-one: no logo in the fee line")
    return False

def _find_clickable(driver, selectors):
    """First visible + enabled element among `selectors` (in priority order), or None."""
    for how, sel in selectors:
        try:
            for el in driver.find_elements(how, sel):
                if el.is_displayed() and el.is_enabled():
                    return el, sel
        except Exception:
            continue
    return None, None


def _click_first(driver, selectors, wait=6, stop_if=None):
    """Click the first matching control. ALL selectors are checked together every
    0.25 s until `wait` seconds have passed in TOTAL (the old code waited `wait`
    seconds per selector, so a step that is simply absent cost up to 18 s).
    stop_if: optional callable(driver) -> True to give up early (step not needed)."""
    deadline = time.time() + wait
    while True:
        el, sel = _find_clickable(driver, selectors)
        if el is not None:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                driver.execute_script("arguments[0].click();", el)
                log.info(f"clicked via {sel[:60]}")
                return True
            except Exception:
                pass
        if stop_if is not None:
            try:
                if stop_if(driver):
                    return False
            except Exception:
                pass
        if time.time() >= deadline:
            return False
        time.sleep(0.25)


def _wait_for(driver, selectors, wait):
    """Wait up to `wait` s for any selector to show a visible element; return it or None."""
    deadline = time.time() + wait
    while True:
        el = _first_visible(driver, selectors)
        if el is not None or time.time() >= deadline:
            return el
        time.sleep(0.25)


def _dialog_open(driver):
    try:
        return any(e.is_displayed() for e in driver.find_elements(By.CSS_SELECTOR, "div[role='dialog']"))
    except Exception:
        return False


def _wait_page_ready(driver, max_wait=SETTLE_SECS):
    """Return as soon as the store page is drawn (title present), max `max_wait` s."""
    deadline = time.time() + max_wait
    while time.time() < deadline:
        try:
            if driver.execute_script("return document.readyState") == "complete" and (
                    driver.find_elements(By.TAG_NAME, "h1")
                    or driver.find_elements(By.CSS_SELECTOR, "a[href*='/store/']")):
                time.sleep(0.5)          # small buffer for the header/fee line to hydrate
                return True
        except Exception:
            pass
        time.sleep(0.25)
    return False


def _first_visible(driver, selectors):
    for how, sel in selectors:
        for e in driver.find_elements(how, sel):
            try:
                if e.is_displayed():
                    return e
            except Exception:
                continue
    return None


# =============================================================================
# SET DELIVERY LOCATION  (exact flow from the client's screenshots)
#
#   1. click the delivery ADDRESS BAR in the header        (screenshot 2)
#   2. the "Addresses" dialog opens; type into "Search for
#      an address"                                          (screenshot 3)
#   3. pick the FIRST suggestion                            (screenshot 4)
#   4. "Choose your building" dialog -> click "Other"       (screenshot 5)
#   5. "Address info" dialog -> fill the required
#      "Additional address information" -> click Save       (screenshot 6)
#   6. the page reloads at the new location
# Each step has several fallbacks; anything Uber may skip (e.g. the building
# dialog for an already-saved address) is optional and simply passed over.
# =============================================================================
_ADDR_BAR_SELECTORS = [
    (By.CSS_SELECTOR, "[data-testid='delivery-address-button']"),
    (By.CSS_SELECTOR, "[data-testid*='address' i]"),
    # header button showing the current address + "Now" and the "·" separator
    (By.XPATH, "//header//button[contains(.,'·') and contains(.,'Now')]"),
    (By.XPATH, "//button[contains(.,'·') and (contains(.,'Now') or contains(.,'Ave') or contains(.,'St'))]"),
    # no address set yet: the header button reads "Enter delivery address"
    (By.XPATH, "//header//button[contains(translate(.,'DELIVERYADS','deliveryads'),'delivery address')]"),
    (By.XPATH, "//*[self::button or self::a or @role='button'][contains(translate(.,'ENTRDLIVYADS','entrdlivyads'),'enter delivery address')]"),
    (By.XPATH, "//header//*[normalize-space(text())='Enter delivery address']"),
]


def _read_header_address(driver):
    """The address currently shown in the header address button ('' if unreadable)."""
    for how, sel in _ADDR_BAR_SELECTORS[:3]:
        try:
            for el in driver.find_elements(how, sel):
                txt = (el.text or "").strip()
                if txt and el.is_displayed():
                    return txt.split("\n")[0].split("·")[0].strip()
        except Exception:
            continue
    return ""
_ADDR_SEARCH_INPUT = [
    (By.CSS_SELECTOR, "input[placeholder*='Search for an address' i]"),
    (By.CSS_SELECTOR, "input[placeholder*='address' i]"),
    (By.XPATH, "//div[@role='dialog']//input[@type='text']"),
    (By.CSS_SELECTOR, "input[type='text']"),
]
_SUGGESTION_SELECTORS = [
    (By.CSS_SELECTOR, "div[role='dialog'] [role='option']"),
    (By.CSS_SELECTOR, "[role='option']"),
    (By.XPATH, "//div[@role='dialog']//ul//li"),
    (By.XPATH, "//div[@role='listbox']//*[@role='option']"),
]
# ---- choosing the RIGHT suggestion --------------------------------------
# Typing into Uber's box first shows the account's SAVED addresses, then the
# real search results replace them a moment later. Taking "the first option"
# too early picked an old saved address (seen in the log: wanted '3540 Grand
# Ave', got '650 NW 27th Ave'). So we wait until an option actually matches
# the address we typed (same house number + street name) and click THAT one.
_STREET_SKIP = {"n", "s", "e", "w", "ne", "nw", "se", "sw", "north", "south", "east",
                "west", "st", "street", "ave", "avenue", "rd", "road", "blvd", "dr",
                "drive", "ln", "lane", "ct", "hwy", "pkwy", "way", "pl", "suite", "ste"}


def clean_address(addr):
    """Make a sheet address typeable for Uber's search box.

      'P.S. 019, 98-02 Roosevelt Ave, Corona (~1.8 mi)'  ->  '98-02 Roosevelt Ave, Corona'
      '153-39 Hillside Ave., Jamaica, NY 11432'           ->  '153-39 Hillside Ave, Jamaica, NY 11432'
      '2340 Fox Hills Dr, Los Angeles, CA90064'          ->  '2340 Fox Hills Dr, Los Angeles, CA 90064'
    (school / building names in front, '(~x mi)' notes and stray dots confuse
    Uber's search: it answered "can't find this address" for them)."""
    a = str(addr or "").strip()
    a = re.sub(r"\([^)]*\)", " ", a)                       # '(~1.8 mi)'
    a = re.sub(r"\b([A-Za-z]{2,})\.(?=\s|,|$)", r"\1", a)  # 'Ave.' -> 'Ave'
    a = re.sub(r"\b([A-Z]{2})(\d{5})\b", r"\1 \2", a)      # 'CA90064' -> 'CA 90064'
    parts = [p.strip() for p in a.split(",") if p.strip()]
    # drop a leading name ('P.S. 019', 'Main Mall') when a later part starts
    # with a house number
    if parts and not re.match(r"\d", parts[0]):
        for i, p in enumerate(parts[1:], 1):
            if re.match(r"\d[\d-]*\s+[A-Za-z]", p):     # '98-02 Roosevelt Ave' (not a bare zip)
                parts = parts[i:]
                break
    return re.sub(r"\s+", " ", ", ".join(parts)).strip()


def _addr_key(text):
    """('3540', 'grand') from '3540 Grand Ave Suite A, Gurnee, IL' - (None, word) if no number."""
    low = re.sub(r"[\u2019'`]", "", (text or "").lower())   # Luther's == Luthers
    m = re.match(r"\s*(\d{1,6}(?:-\d{1,6})?)\b", low)   # house number = FIRST token only
    house = m.group(1) if m else None
    tail = low[m.end():] if m else low
    street = next((w for w in re.findall(r"[a-z0-9]+", tail)
                   if w not in _STREET_SKIP and not w.isdigit()), None)
    return house, street


def _matches_address(option_text, address):
    """True when an option / header text is the same street address as `address`."""
    want_h, want_s = _addr_key(address)
    got_h, got_s = _addr_key(option_text)
    if want_h:                                   # we typed a house number -> must match
        if got_h != want_h:
            return False
        low = (option_text or "").lower()
        if not (want_s and got_s) or want_s == got_s or want_s in low:
            return True
        # same house number, street written differently ('Highway 292' vs 'SC-292'):
        # accept when they share the route number
        want_nums = set(re.findall(r"\d+", (address or "").lower().split(",")[0])) - {want_h}
        return bool(want_nums & set(re.findall(r"\d+", low)))
    # no house number (e.g. 'Queens-Midtown Tunnel ...'): first words must match
    opt = re.sub(r"[\u2019'`]", "", (option_text or "").lower())
    if not (want_s and want_s in opt):
        return False
    first = (address or "").lower().split(",")[0]
    types = ("avenue", "ave", "street", "st", "road", "rd", "boulevard", "blvd", "drive", "dr",
             "lane", "ln", "highway", "hwy", "parkway", "pkwy", "place", "pl", "court", "ct", "way")
    asked = [t for t in types if re.search(rf"\b{t}\b", first)]
    return not asked or any(re.search(rf"\b{t}\b", opt) for t in types)


def _visible_options(driver):
    out = []
    for how, sel in _SUGGESTION_SELECTORS:
        try:
            for e in driver.find_elements(how, sel):
                if e.is_displayed() and (e.text or "").strip():
                    out.append(e)
        except Exception:
            continue
        if out:
            break
    return out


_LAST_PICK = {"matched": False}


def _pick_matching_option(driver, address, wait=6):
    """Wait for a suggestion that matches `address` and return it (or the first
    option after `wait` s, or None when no list shows at all)."""
    deadline = time.time() + wait
    last = []
    while True:
        last = _visible_options(driver)
        for e in last:
            try:
                if _matches_address(e.text, address):
                    return e, True
            except Exception:
                continue
        if time.time() >= deadline:
            return (last[0], False) if last else (None, False)
        time.sleep(0.3)


_OTHER_SELECTORS = [
    (By.XPATH, "//div[@role='dialog']//button[normalize-space()='Other']"),
    (By.XPATH, "//button[.//*[normalize-space()='Other']]"),
    (By.XPATH, "//*[normalize-space(text())='Other']/ancestor-or-self::button[1]"),
    (By.XPATH, "//*[normalize-space(text())='Other']"),
]
_ADDR_INFO_REQUIRED = [
    (By.CSS_SELECTOR, "input[placeholder*='Street address' i]"),
    (By.CSS_SELECTOR, "textarea[placeholder*='Street address' i]"),
    (By.CSS_SELECTOR, "input[placeholder*='building name' i]"),
    (By.XPATH, "//*[contains(text(),'Additional address')]/following::input[1]"),
    (By.XPATH, "//*[contains(text(),'Additional address')]/following::textarea[1]"),
]
_SAVE_SELECTORS = [
    (By.XPATH, "//div[@role='dialog']//button[normalize-space()='Save']"),
    (By.XPATH, "//button[normalize-space()='Save']"),
    (By.XPATH, "//button[contains(.,'Save')]"),
]


def set_delivery_location(driver, address):
    """Run the 6-step location flow for `address`. Returns True if it completed.

    Whether the store actually DELIVERS there is decided afterwards by
    check_availability(); this function only sets the location.
    """
    try:
        # 1. open the address bar
        if not _click_first(driver, _ADDR_BAR_SELECTORS):
            log.warning("could not open the address bar")
            return False

        # 2. type the address into the Addresses search box (as soon as it shows)
        inp = _wait_for(driver, _ADDR_SEARCH_INPUT, 5)
        if inp is None:
            log.warning("address search box not found")
            return False
        try:
            inp.clear()
        except Exception:
            pass
        try:
            inp.click()
        except Exception:
            # something (a tooltip / overlay) covers the box - focus it directly
            driver.execute_script("arguments[0].focus();", inp)
        address = clean_address(address)
        inp.send_keys(address)
        time.sleep(1.0)  # let autocomplete start, then take the list as soon as it shows

        # 3. pick the suggestion that MATCHES what we typed (not just the first one)
        opt, matched = _pick_matching_option(driver, address)
        if opt is not None:
            label = (opt.text or "").split("\n")[0][:60]
            driver.execute_script("arguments[0].click();", opt)
            _LAST_PICK["matched"] = matched
            if matched:
                log.info(f"picked matching suggestion: {label}")
            else:
                log.warning(f"no suggestion matched {address!r} - took the first one: {label}")
        else:
            inp.send_keys(Keys.ENTER)
            log.info("no suggestion list; pressed ENTER")
        # A SAVED address is applied straight away - the dialog just closes.
        # Then there is no 'Other' / 'Save' step to wait for. (Checked over 2 s so
        # the short gap before the 'Choose your building' dialog isn't mistaken
        # for "closed".)
        closed_checks = 0
        for _ in range(8):
            time.sleep(0.25)
            closed_checks = closed_checks + 1 if not _dialog_open(driver) else 0
        if closed_checks >= 6:
            log.info("address applied directly (saved address)")
            _wait_page_ready(driver)
            return True

        # 4. "Choose your building" -> Other   (optional; skipped for saved addr)
        if _click_first(driver, _OTHER_SELECTORS, wait=3,
                        stop_if=lambda d: not _dialog_open(d)
                        or _first_visible(d, _ADDR_INFO_REQUIRED) is not None):
            log.info("selected building type 'Other'")
            time.sleep(0.6)

        # 5. "Address info": the additional-info field is marked required, so fill
        #    it (reusing the address) before Save, or Save may stay disabled.
        req = _wait_for(driver, _ADDR_INFO_REQUIRED, 1.5)
        if req is not None:
            try:
                req.clear()
            except Exception:
                pass
            req.send_keys(address)
            log.info("filled required address-info field")
            time.sleep(0.5)

        if not _click_first(driver, _SAVE_SELECTORS, wait=5,
                            stop_if=lambda d: not _dialog_open(d)):
            if _dialog_open(driver):
                log.warning("Save button not found/clicked")

        # page reloads at the new location: wait for the dialog to go, then the page
        deadline = time.time() + 6
        while _dialog_open(driver) and time.time() < deadline:
            time.sleep(0.25)
        _wait_page_ready(driver)
        return True
    except Exception as e:
        log.warning(f"set_delivery_location failed for {address!r}: {e}")
        return False


# Only the TOP of the store page is read for availability (header, delivery
# toggle, fee/ETA row). The menu below can contain item text like "15 min" or
# "Not available" (sold-out items) that would otherwise be misread.
AVAILABILITY_TEXT_WINDOW = 4000


def check_availability(driver):
    """After a location is set, decide Available / Not Available.

    Available  = an estimated delivery-time line is present ("25 - 40 min", etc.).
    Not Available = an explicit out-of-range / not-served message, OR no delivery
                    estimate in the top part of the page.
    """
    try:
        body = driver.find_element(By.TAG_NAME, "body").text
    except Exception:
        body = ""
    body = body[:AVAILABILITY_TEXT_WINDOW]
    low = body.lower()

    for marker in _NOT_AVAILABLE_MARKERS:
        if marker in low:
            log.info(f"availability: Not Available (marker '{marker}')")
            _LAST_AVAIL_REASON["why"] = "marker"
            return "Not Available"

    for rx in _DELIVERY_ESTIMATE_PATTERNS:
        if rx.search(body):
            log.info(f"availability: Available (matched /{rx.pattern}/)")
            _LAST_AVAIL_REASON["why"] = "eta"
            return "Available"

    log.info("availability: Not Available (no delivery estimate found)")
    _LAST_AVAIL_REASON["why"] = "no_eta"
    return "Not Available"


# why the last check_availability() verdict was reached: marker | eta | no_eta
_LAST_AVAIL_REASON = {"why": ""}


# =============================================================================
# ONE RADIUS CHECK (used by main.py for the 2-mile and the 5-mile address)
#
#   1. make sure the STORE page is on screen (re-open it if not)
#   2. set the delivery address (address bar -> search -> first suggestion ->
#      Other -> address info -> Save)
#   3. confirm the header now shows THAT address - if Uber did not take it,
#      report "Address not set" instead of a misleading Available/Not Available
#   4. Save can drop us on the home feed: if so, re-open the store page
#   5. read availability on the store page
# =============================================================================
def _on_store(driver, url):
    try:
        return extract_store_id(driver.current_url) == extract_store_id(url)
    except Exception:
        return False


def open_store(driver, url):
    """Open (or re-open) the store page in the current window and let it settle."""
    driver.get(url)
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "h1")))
    except Exception:
        pass
    _wait_page_ready(driver)


_open_store = open_store   # old name, kept for compatibility


def close_overlays(driver):
    """Close an open dialog (e.g. the store-info panel Phase 1 leaves open)
    WITHOUT reloading the page. Returns True when nothing is left covering it."""
    from selenium.webdriver.common.action_chains import ActionChains
    for _ in range(2):
        try:
            if not any(e.is_displayed() for e in driver.find_elements(By.CSS_SELECTOR, "div[role='dialog']")):
                return True
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(1)
        except Exception:
            break
    try:
        return not any(e.is_displayed() for e in driver.find_elements(By.CSS_SELECTOR, "div[role='dialog']"))
    except Exception:
        return False


def apply_address(driver, url, address):
    """Set `address` as the delivery address while on the store page, confirm the
    header shows it, and make sure we end up back on the store page.
    Returns True / False. Used for the STORE address (before scraping) and for
    the 2-mile / 5-mile addresses."""
    try:
        from scraper import _read_pin_address, _pin_matches_store   # read-only helpers
    except Exception:
        _read_pin_address = _pin_matches_store = None

    address = clean_address(address)
    if not _on_store(driver, url):
        open_store(driver, url)
    if not close_overlays(driver):
        open_store(driver, url)

    ok = set_delivery_location(driver, address)
    if not ok:
        # usually the page was not fully drawn yet - reload the store once and retry
        log.info(f"address flow did not complete for {address!r} - reloading store and retrying once")
        open_store(driver, url)
        close_overlays(driver)
        ok = set_delivery_location(driver, address)
    if not ok:
        log.warning(f"address flow did not complete for {address!r}")
        return False

    def _accepted(pin_text):
        if not pin_text or _matches_address(pin_text, address):
            return True
        # We clicked the suggestion that matched our address, but Uber shows a
        # neighbouring building on the SAME street ('23-71 19th St' -> '24-23 19th St').
        # That is Uber snapping the pin, not a wrong address - accept it.
        if _LAST_PICK.get("matched"):
            w, g = _addr_key(address)[1], _addr_key(pin_text)[1]
            if w and g and w == g:
                log.info(f"Uber placed the pin at {pin_text!r} on the same street - accepted")
                return True
        return False

    pin = _read_header_address(driver) or (_read_pin_address(driver) if _read_pin_address else "")
    if not _accepted(pin):
        # wrong address went in (e.g. an old saved one) - try once more from a fresh page
        log.info(f"header shows {pin!r}, wanted {address!r} - retrying once")
        open_store(driver, url)
        close_overlays(driver)
        if set_delivery_location(driver, address):
            pin = _read_header_address(driver) or pin
    if not _accepted(pin):
        log.warning(f"address not applied: header shows {pin!r}, wanted {address!r}")
        return False
    log.info(f"address applied (header: {pin or 'unreadable'})")

    if not _on_store(driver, url):
        log.info("left the store page after Save - re-opening it")
        open_store(driver, url)
    return True


# =============================================================================
# 2-MILE / 5-MILE  =  "is the store VISIBLE in the Flowers section at that address"
#
#   1. click "Uber Eats" in the header            -> home page
#   2. set the 2-mile (or 5-mile) address in the header address bar, Save
#   3. click "Flowers" in the left menu
#   4. our store's card (its link) is on that page  -> Available
#      not there                                     -> Not Available
#
# One Flowers page shows EVERY store visible at that address, so each address
# is checked ONCE and the answer is used for all our stores that need it.
# =============================================================================
UBER_HOME = "https://www.ubereats.com/feed?diningMode=DELIVERY"

_HOME_LOGO = [
    (By.CSS_SELECTOR, "header a[href='/']"),
    (By.CSS_SELECTOR, "header a[aria-label*='Uber Eats' i]"),
    (By.CSS_SELECTOR, "a[data-testid*='logo' i]"),
    (By.XPATH, "//header//a[.//*[name()='svg'] and (@href='/' or contains(@href,'/feed'))]"),
]
_FLOWERS_LINK = [
    (By.XPATH, "//nav//a[normalize-space(.)='Flowers']"),
    (By.XPATH, "//a[normalize-space(.)='Flowers']"),
    (By.XPATH, "//*[normalize-space(text())='Flowers']/ancestor::a[1]"),
    (By.XPATH, "//*[normalize-space(text())='Flowers']/ancestor::*[@role='button' or @role='link' or self::button][1]"),
    # a link that GOES to the flowers section, whatever its text/layout
    (By.XPATH, "//a[contains(translate(@href,'FLOWERS','flowers'),'flower') and not(contains(@href,'/store/'))]"),
    # last resort: the word itself (the click bubbles up to its link)
    (By.XPATH, "//*[normalize-space(text())='Flowers']"),
]
_FLOWERS_URL = {"url": ""}      # remembered after the first click (used if the menu is hidden)


def _on_home(driver):
    try:
        from urllib.parse import urlparse
        path = urlparse(driver.current_url or "").path.rstrip("/")
        return path in ("", "/feed") or path.endswith("/feed")
    except Exception:
        return False


def go_home(driver):
    """Step 1: open the Uber Eats home page.

    Opened directly by its address - exactly where clicking "Uber Eats" in the
    header leads - because a click can land on the wrong header link. If the
    home page still isn't showing, try clicking the header logo as well."""
    driver.get(UBER_HOME)
    _wait_for_feed(driver)
    if not _on_home(driver):
        _click_first(driver, _HOME_LOGO, wait=3)
        _wait_for_feed(driver)
    if not _on_home(driver):
        log.warning(f"home page did not open - browser is on {driver.current_url[:90]}")
        _snapshot(driver, "home_not_open")
        return False
    return True


def _snapshot(driver, name):
    """Save a screenshot next to scraper.log so a problem can be seen afterwards."""
    try:
        path = os.path.join(os.getcwd(), f"{name}_{time.strftime('%H%M%S')}.png")
        driver.save_screenshot(path)
        log.info(f"screenshot saved: {path}")
    except Exception:
        pass


def _wait_for_feed(driver, max_wait=12):
    deadline = time.time() + max_wait
    while time.time() < deadline:
        try:
            if driver.execute_script("return document.readyState") == "complete" and \
                    driver.find_elements(By.CSS_SELECTOR, "a[href*='/store/'], main"):
                time.sleep(0.6)
                return True
        except Exception:
            pass
        time.sleep(0.3)
    return False


def set_address_on_home(driver, address):
    """Step 2: put `address` in the header address bar (same flow as before) and
    confirm the header shows it. Returns True / False."""
    address = clean_address(address)
    for attempt in (1, 2):
        if attempt == 2:
            log.info(f"retrying the address {address!r} from a fresh home page")
            driver.get(UBER_HOME)
            _wait_for_feed(driver)
        close_overlays(driver)
        if not set_delivery_location(driver, address):
            continue
        pin = _read_header_address(driver)
        if not pin or _matches_address(pin, address):
            log.info(f"home address set (header: {pin or 'unreadable'})")
            return True
        if _LAST_PICK.get("matched"):
            w, g = _addr_key(address)[1], _addr_key(pin)[1]
            if w and g and w == g:
                log.info(f"Uber placed the pin at {pin!r} on the same street - accepted")
                return True
        log.info(f"header shows {pin!r}, wanted {address!r}")
    log.warning(f"address not set on home page: {address!r}")
    return False


def open_flowers(driver):
    """Step 3: click 'Flowers' in the left menu of the home page."""
    el, _ = _find_clickable(driver, _FLOWERS_LINK)
    if el is None:
        # the left menu can be collapsed on a narrow window - open it
        _click_first(driver, [(By.CSS_SELECTOR, "button[aria-label*='menu' i]")], wait=2)
        el = _wait_for(driver, _FLOWERS_LINK, 3)
    if el is not None:
        try:
            href = el.get_attribute("href")
            if href:
                _FLOWERS_URL["url"] = href
            driver.execute_script("arguments[0].click();", el)
        except Exception:
            el = None
    if el is None:
        if not _FLOWERS_URL["url"]:
            log.warning("could not find 'Flowers' in the left menu")
            _snapshot(driver, "flowers_not_found")
            return False
        driver.get(_FLOWERS_URL["url"])
    _wait_for_feed(driver)
    return True


_COLLECT_STORE_IDS_JS = """
    const ids = new Set();
    for (const a of document.querySelectorAll("a[href*='/store/']")) {
        const m = a.getAttribute('href').match(/\\/store\\/[^\\/]+\\/([^\\/?#]+)/);
        if (m) ids.add(m[1].toLowerCase());
    }
    return Array.from(ids);
"""


def visible_store_ids(driver, wanted=None, max_scrolls=25):
    """Step 4: every store id linked on the page. Scrolls down so lazily loaded
    rows appear; stops early once all `wanted` ids are found or nothing new loads."""
    wanted = {w.lower() for w in (wanted or ())}
    seen, stale = set(), 0
    for _ in range(max_scrolls):
        try:
            now = set(driver.execute_script(_COLLECT_STORE_IDS_JS) or [])
        except Exception:
            now = set()
        new = now - seen
        seen |= now
        if wanted and wanted <= seen:
            break
        stale = stale + 1 if not new else 0
        if stale >= 2:
            break
        try:
            at_bottom = driver.execute_script(
                "window.scrollBy(0, window.innerHeight * 0.9);"
                "return (window.innerHeight + window.scrollY) >= document.body.scrollHeight - 5;")
        except Exception:
            at_bottom = True
        time.sleep(0.7)
        if at_bottom and not new:
            break
    return seen


_CARD_TEXTS_JS = """
    const out = [];
    for (const a of document.querySelectorAll("a[href*='/store/']")) {
        const r = a.getBoundingClientRect();
        if (!r.width || !r.height) continue;
        const t = (a.innerText || a.getAttribute('aria-label') || '').replace(/\\s+/g, ' ').trim();
        if (t) out.push(t.toLowerCase());
    }
    return out;
"""


def _flowers_page_names(driver, brand, max_scrolls=4):
    """Names of the stores shown on the Flowers screen. Looks at the screen,
    scrolls a little further only while our brand isn't seen yet."""
    names = []
    for _ in range(max_scrolls):
        try:
            names = list(dict.fromkeys(names + (driver.execute_script(_CARD_TEXTS_JS) or [])))
        except Exception:
            pass
        if any(brand in n for n in names):
            break
        try:
            driver.execute_script("window.scrollBy(0, window.innerHeight * 0.9);")
        except Exception:
            break
        time.sleep(0.6)
    return names


def check_address_visibility(driver, address, wanted_ids, brand="bomi"):
    """2-mile / 5-mile for one address.

      first address : home page -> set address -> Flowers
      next addresses: already on the Flowers screen -> just change the address
                      there (the screen reloads for the new address) - no extra
                      home page / Flowers clicks, which saves ~20 s per address.
    Then: is the store NAME (e.g. 'Bomi Flowers...') on the Flowers screen?

    Returns (ok, visible_ids, visible_names). ok: True = checked,
    False = address not accepted, None = Flowers screen could not be opened.
    """
    on_flowers = bool(_FLOWERS_URL["url"]) and \
        (driver.current_url or "").split("?")[0] == _FLOWERS_URL["url"].split("?")[0]
    if not on_flowers and not go_home(driver):
        return None, set(), []
    if not set_address_on_home(driver, address):
        return False, set(), []
    if _FLOWERS_URL["url"]:
        if (driver.current_url or "").split("?")[0] != _FLOWERS_URL["url"].split("?")[0]:
            driver.get(_FLOWERS_URL["url"])
            _wait_for_feed(driver)
    else:
        if not _on_home(driver):
            go_home(driver)
        if not open_flowers(driver):
            return None, set(), []
        _FLOWERS_URL["url"] = _FLOWERS_URL["url"] or (driver.current_url or "")
    try:
        driver.execute_script("window.scrollTo(0, 0);")
    except Exception:
        pass
    names = _flowers_page_names(driver, brand)
    try:
        ids = set(driver.execute_script(_COLLECT_STORE_IDS_JS) or [])
    except Exception:
        ids = set()
    ours = [n for n in names if brand in n]
    log.info(f"Flowers screen for {address!r}: {len(names)} store(s) shown; "
             f"{brand!r} stores shown: {[n[:50] for n in ours] or 'none'}")
    return True, ids, names


def check_radius(driver, url, address, store_status=""):
    """Returns 'Available' | 'Not Available' | 'Not Available (store closed)' |
    'Address not set' | 'No address' | 'Error'.

    store_status is the Phase 1 verdict. A closed store shows no delivery time
    at ANY address, so 'no ETA' there is reported as '(store closed)' instead of
    being mistaken for 'out of range'. An explicit out-of-range message still
    counts as plain 'Not Available'."""
    if not address or str(address).strip().lower() in ("", "nan"):
        return "No address"
    try:
        if not apply_address(driver, url, address):
            return "Address not set"

        verdict = check_availability(driver)
        closed = str(store_status or "").strip().lower() in ("closed", "force closed")
        if verdict == "Not Available" and _LAST_AVAIL_REASON["why"] == "no_eta" and closed:
            return "Not Available (store closed)"
        return verdict
    except Exception as e:
        log.warning(f"radius check failed for {address!r}: {e}")
        return "Error"


# =============================================================================
# PER-STORE
# =============================================================================
def process_store(driver, row):
    url = str(row.get(STORE_URL_COL, "")).strip()
    a2  = str(row.get(ADDR_2MI_COL, "")).strip()
    a5  = str(row.get(ADDR_5MI_COL, "")).strip()
    zipc = ""
    for k in row.keys():
        if str(k).strip().lower() in ("zipcode", "zip", "zip code", "postal code"):
            zipc = str(row.get(k, "")).strip()
            break
    state, tz = determine_timezone_and_state(zipc) if zipc else ("Unknown State", "Unknown Timezone")

    result = {
        "Store Name": (url.split("/store/")[1].split("/")[0].replace("-", " ").title()
                       if "/store/" in url else "Unknown Store"),
        "State": state, "Timezone": tz, "Website_link": url,
        "U1 Icon": "", "2mile": "", "5mile": "",
    }
    if not url or url.lower() == "nan":
        return result

    log.info(f"--- store: {result['Store Name']} ---")
    driver.get("about:blank")
    driver.get(url)
    time.sleep(SETTLE_SECS)

    # Uber One badge decides both the U1 Icon column AND whether the 2mi/5mi
    # checks run: per the client's flow, the address-bar / availability check is
    # only meaningful when the store shows Uber One. No badge -> the store isn't
    # an Uber One store, so the radius columns are marked accordingly, no clicks.
    has_u1 = detect_uber_one(driver)
    result["U1 Icon"] = "U1icon_present" if has_u1 else ""

    if not has_u1:
        result["2mile"] = result["5mile"] = "No Uber One"
        log.info("no Uber One badge -> skipping 2mi/5mi checks")
        return result

    if a2 and a2.lower() != "nan":
        log.info(f"2-mile check: {a2}")
        set_delivery_location(driver, a2)
        result["2mile"] = check_availability(driver)
    else:
        result["2mile"] = "No address"

    time.sleep(random.uniform(*BETWEEN_ADDR))

    # reload the store fresh before the 5-mile change so we start clean
    driver.get(url)
    time.sleep(SETTLE_SECS)
    if a5 and a5.lower() != "nan":
        log.info(f"5-mile check: {a5}")
        set_delivery_location(driver, a5)
        result["5mile"] = check_availability(driver)
    else:
        result["5mile"] = "No address"

    log.info(f"result: U1={result['U1 Icon'] or '-'} | 2mi={result['2mile']} | 5mi={result['5mile']}")
    return result


# =============================================================================
# OUTPUT
# =============================================================================
def write_output(records):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    cols = ["Store Name", "State", "Timezone", "U1 Icon", "2mile", "5mile", "Website_link"]
    df = pd.DataFrame(records, columns=cols)

    folder = f"uberone {datetime.now():%d-%m-%y}"
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"uberone_{datetime.now():%d-%m-%Y_%I-%M-%S_%p}.xlsx")
    df.to_excel(path, index=False)

    wb = load_workbook(path); ws = wb.active
    thin = Border(*[Side(style="thin", color="D9D9D9")]*4)
    head = PatternFill("solid", fgColor="1F4E79")
    ok   = PatternFill("solid", fgColor="C6EFCE")   # green  = Available
    no   = PatternFill("solid", fgColor="F7B6B6")   # red    = Not Available
    widths = {"Store Name":30,"State":14,"Timezone":10,"U1 Icon":16,"2mile":14,"5mile":14,"Website_link":16}
    for i, name in enumerate(cols, 1):
        c = ws.cell(row=1, column=i); c.fill = head
        c.font = Font(name="Segoe UI", bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = widths.get(name, 14)
    ws.row_dimensions[1].height = 24
    ci_2, ci_5 = cols.index("2mile")+1, cols.index("5mile")+1
    for r in range(2, ws.max_row+1):
        for i in range(1, len(cols)+1):
            cell = ws.cell(row=r, column=i)
            cell.border = thin
            cell.alignment = Alignment(horizontal="center" if i in (ci_2,ci_5,4) else "left",
                                       vertical="center", wrap_text=True)
        for ci in (ci_2, ci_5):
            v = str(ws.cell(row=r, column=ci).value or "")
            if v == "Available":     ws.cell(row=r, column=ci).fill = ok
            elif v == "Not Available": ws.cell(row=r, column=ci).fill = no
    ws.freeze_panes = "A2"
    wb.save(path)
    return path


# =============================================================================
# MAIN
# =============================================================================
def main():
    _setup_standalone_logging()   # standalone run -> its own file+console handlers
    ap = argparse.ArgumentParser()
    ap.add_argument("--login", action="store_true",
                    help="open the browser to sign in to Uber Eats once, then exit")
    ap.add_argument("--browser", choices=["chrome", "edge"], default=None,
                    help="force the browser for this run (overrides settings.ini / env)")
    args = ap.parse_args()

    if getattr(sys, "frozen", False):
        os.chdir(os.path.dirname(sys.executable))
    else:
        os.chdir(os.path.dirname(os.path.abspath(__file__)))

    preferred = resolve_browser(args.browser)
    log.info(f"preferred browser: {preferred}")

    if args.login:
        d, used = build_driver(preferred=preferred, headless=False)
        d.get("https://www.ubereats.com/")
        print(f"\n[{used}] Sign in to Uber Eats in the window that opened, then press ENTER here.")
        try:
            input()
        except EOFError:
            time.sleep(120)
        print("logged in:", is_logged_in(d))
        d.quit()
        return

    if not os.path.exists(INPUT_FILE):
        log.error(f"{INPUT_FILE} not found"); sys.exit(1)
    df = pd.read_excel(INPUT_FILE)
    if STORE_URL_COL not in df.columns:
        log.error(f"input needs a '{STORE_URL_COL}' column"); sys.exit(1)

    driver, _used = build_driver(preferred=preferred, headless=False)
    try:
        if not ensure_login(driver):
            sys.exit(2)
        records = []
        for idx, row in df.iterrows():
            if not str(row.get(STORE_URL_COL, "")).strip():
                continue
            try:
                records.append(process_store(driver, row))
            except Exception as e:
                log.error(f"row {idx} failed: {e}")
            time.sleep(random.uniform(*BETWEEN_STORES))
    finally:
        try: driver.quit()
        except Exception: pass

    if records:
        out = write_output(records)
        log.info(f"saved: {out}")
    log.info("done.")


if __name__ == "__main__":
    main()