#main.py  -  EXTENDED VERSION (runs inside the logged-in Chrome)
"""
Uber Eats scraper - EXTENDED (signed-in) version.

Per store, in ONE window of the logged-in Chrome profile:
    1. Phase 1 scrape (status, tags, hours, phone, rating ...)  - unchanged logic
    2. Uber One badge?   no  -> 2mile / 5mile = 'No Uber One'
                         yes -> set the 2-mile address, check delivery
                                set the 5-mile address, check delivery
    3. next store row, same window

The browser is never hidden. A robot check pauses the run until a person
solves it in that window. Other Chrome windows/tabs are never touched, and at
the end only the scraper's own window is closed - Chrome and the login stay.

Settings (port, profile folder, tab/window, data folder ...) live in
settings.ini or UE_* environment variables - see settings.py. The same code
runs on a Windows PC and on a cloud VM/container.
"""
import logging
import time
import os
import html
import sys
import multiprocessing
import random
import pandas as pd
from urllib.parse import unquote

# Local modules
import settings
from browser_session import ScraperSession, body_text_lower
from scraper import scrape_ubereats_store_dynamically, get_failure_reason
from excel_formatter import format_and_export_excel
from excel_formatter import deliver_report_via_email
from timezone_utils import determine_timezone_and_state, extract_store_id
import scraper as _scraper_mod

# Phase 2 (Uber One) checks live in uberone.py. The import is guarded so any
# problem loading it can never stop the Phase 1 scrape.
ENABLE_UBER_ONE = True
try:
    import uberone as _uberone
except Exception:
    _uberone = None

# Extended version = always the signed-in Chrome:
#   * store-cookie reset uses the strict location-only allow-list (login untouched)
#   * the address-pin check is off (signed in, the pin is the account's address)
_scraper_mod.LOGGED_IN_MODE = True
_scraper_mod.REQUIRE_PIN_MATCH = False
# Before every new store URL: delete ONLY the previous store's delivery-location
# cookies (uev2.loc / uev2.diningMode) + location storage keys, so the next
# store's address goes in clean. Login cookies are never touched (strict list
# in scraper.clear_location_state, LOGGED_IN_MODE).
_scraper_mod.PRESERVE_SESSION = False

# Logging Setup  (log file lives in the data folder - next to the .exe by default,
# or on the mounted volume when hosted on cloud)
DATA_DIR = settings.data_dir()
logger = logging.getLogger()
logger.setLevel(logging.INFO)
file_handler = logging.FileHandler(os.path.join(DATA_DIR, "scraper.log"), encoding="utf-8")
file_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(file_handler)
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(logging.Formatter("%(message)s"))
logger.addHandler(console_handler)
# Keep the granular load/verify trail (logger name "scraper.loadwatch") in the
# log FILE only - the console output stays exactly as it was.
# CONSOLE = short and client-friendly: only this file's own messages (one line
# per store + summary) and the "please solve the robot check" prompts.
# Every technical step (clicks, addresses, cookies, logo checks) still goes to
# scraper.log in full.
console_handler.addFilter(
    lambda r: r.name == "root" or (r.name == "ue.session" and r.levelno >= logging.WARNING))
_detail = logging.getLogger("ue.detail")      # file-only detail lines


# ---------------------------------------------------------------------------
# Robot / captcha check -> pause for a human (the browser is always visible)
# ---------------------------------------------------------------------------
_BOT_WALL_MARKERS = getattr(_scraper_mod, "_BOT_WALL_MARKERS", ())


def _bot_wall_showing(driver):
    low = body_text_lower(driver)
    return any(m in low for m in _BOT_WALL_MARKERS)


# ---------------------------------------------------------------------------
# Login gate (before scraping) + watchdog (during the run)
# ---------------------------------------------------------------------------
_u1_login = {"checked": False, "ok": False}


def _login_instructions():
    acc = os.environ.get("UE_ACCOUNT_NO", "").strip()
    if acc:      # started by run_all.py for one account
        exe = (f"{os.path.basename(sys.executable)} login {acc}" if getattr(sys, "frozen", False)
               else f"python run_all.py login {acc}")
    else:
        exe = (f"{os.path.basename(sys.executable)} login" if getattr(sys, "frozen", False)
               else "python main.py login")
    logging.warning("=" * 70)
    logging.warning(f"{'Account ' + acc if acc else 'The scraper'} is NOT signed in to Uber Eats.")
    logging.warning("One-time fix:")
    logging.warning("  1. Close the scraper's Chrome completely.")
    logging.warning(f"  2. Run:  {exe}")
    logging.warning("  3. Sign in to Uber Eats in the Chrome that opens (finish any OTP),")
    logging.warning("     check your address + search + cart show at the top, then close that Chrome.")
    logging.warning("  4. Run the scraper again.")
    logging.warning("=" * 70)


def wait_for_login(session, first_url=None):
    """Check (once, on the first store) that the scraper's Chrome is signed in.

    Signing in is NOT done in the scraper's window: Uber's sign-in page does not
    load properly in a browser that is being controlled. It is done once with
    `main.py login` (plain Chrome, same profile). Returns True / False.
    """
    if _uberone is None:
        return True
    ok = False
    # Two tries: a slow first page load can look "signed out" for a moment.
    for attempt in (1, 2):
        try:
            ok = _uberone.ensure_login(session.driver, first_url)
        except Exception as e:
            logging.warning(f"login check failed: {e}")
        if ok:
            break
        if attempt == 1:
            _detail.info("Login not confirmed on the first try - checking once more")
            time.sleep(5)
    _u1_login.update(checked=True, ok=ok)
    if ok:
        # run_all.py uses this marker to know the account is ready (never removed
        # automatically, so a one-off slow page can't force a re-login)
        try:
            with open(os.path.join(settings.profile_dir(), "signed_in.txt"), "w") as fh:
                fh.write(time.strftime("signed in (checked %Y-%m-%d %H:%M)"))
        except Exception:
            pass
    if not ok:
        session.bring_to_front()
        _login_instructions()
    return ok


def login_watchdog(session, current_url):
    """Every few stores: if Uber signed the account out mid-run, stop the Uber One
    checks (they need a login) but let the normal scrape carry on."""
    if _uberone is None or not _u1_login["ok"]:
        return
    try:
        if not _uberone.looks_logged_out(session.driver):
            return
    except Exception:
        return
    logging.error("The Uber Eats account was SIGNED OUT mid-run - Uber One columns for the "
                  "remaining stores will read 'Login required'. The normal scrape continues.")
    _u1_login.update(checked=True, ok=False)


# ---------------------------------------------------------------------------
# Uber One (same window, right after the Phase 1 scrape of the store)
# ---------------------------------------------------------------------------
def enrich_with_uber_one(session, url, row, a2_col, a5_col, store_status=""):
    """U1 Icon / 2mile / 5mile for one store, in the scraper window.

      1. close the info panel Phase 1 left open (no reload)
      2. Uber One badge?   no  -> 'No Uber One' for both radius columns
      3. yes -> set the 2-mile address, check delivery availability
      4.      -> set the 5-mile address, check delivery availability
    Fully failure-isolated: a problem only marks these three columns.
    """
    out = {'U1 Icon': '', '2mile': '', '5mile': ''}
    if not ENABLE_UBER_ONE or _uberone is None:
        return out
    if _u1_login["checked"] and not _u1_login["ok"]:
        out['2mile'] = out['5mile'] = 'Login required'
        return out

    driver = session.driver
    a2 = str(row.get(a2_col, '')).strip() if a2_col else ''
    a5 = str(row.get(a5_col, '')).strip() if a5_col else ''
    try:
        if not _uberone.close_overlays(driver):
            _uberone.open_store(driver, url)        # panel would not close - reload once
        if _bot_wall_showing(driver):
            if session.wait_for_human_check(_bot_wall_showing):
                _uberone.open_store(driver, url)
            else:
                out['2mile'] = out['5mile'] = 'Robot check'
                return out

        has_u1 = _uberone.detect_uber_one(driver)
        out['U1 Icon'] = 'U1icon_present' if has_u1 else ''
        if not has_u1:
            out['2mile'] = out['5mile'] = 'No Uber One'
            return out
        if not CHECK_RADIUS["on"]:
            # 2-mile / 5-mile switched off in input.xlsx (quick daily run)
            out['2mile'] = out['5mile'] = 'Not checked'
            return out
        # 2-mile / 5-mile are checked LATER on the home page -> Flowers, once per
        # address for all stores together (see run_radius_phase).
        sid = extract_store_id(url) or url
        out['2mile'] = 'No address' if not a2 or a2.lower() == 'nan' else 'Pending'
        out['5mile'] = 'No address' if not a5 or a5.lower() == 'nan' else 'Pending'
        _RADIUS_TODO[sid] = (a2, a5)
        _detail.info(f"U1 [{str(url)[-40:]}]: icon present - 2/5-mile queued")
    except Exception as e:
        logging.error(f"Uber One check failed for {url}: {e}")
        if not out['2mile']: out['2mile'] = 'Error'
        if not out['5mile']: out['5mile'] = 'Error'
    return out


# 2-mile / 5-mile on or off for this run - set from input.xlsx (column
# 'check_2_5_mile': Yes / No). Off = the quick daily run.
CHECK_RADIUS = {"on": True}

_RADIUS_COL_NAMES = {'check_2_5_mile', 'check 2/5 mile', 'check_2mile_5mile', '2_5_mile',
                     'check 2_5 mile', 'run_2_5_mile', 'check_radius', 'two_five_mile'}

# The name looked for on the Flowers screen (store cards read "Bomi Flowers, Plants & Gifts ...")
BRAND = (settings.get("store_brand") or "bomi").strip().lower()

# store_id -> (address_2mile, address_5mile) for Uber One stores of this run
_RADIUS_TODO = {}


def run_radius_phase(session, results_by_key):
    """2-mile / 5-mile visibility for every queued Uber One store.

    Each DIFFERENT address is checked once (home page -> set address -> Flowers),
    and the list of stores visible there answers every store that uses that
    address - so shared addresses cost nothing extra.
    """
    if not _RADIUS_TODO:
        return
    # address (as typed, normalised) -> list of (store_id, column)
    jobs = {}
    for sid, (a2, a5) in _RADIUS_TODO.items():
        for addr, col in ((a2, '2mile'), (a5, '5mile')):
            if addr and str(addr).lower() != 'nan':
                key = _uberone.clean_address(addr).lower()
                jobs.setdefault(key, {"addr": addr, "targets": []})["targets"].append((sid, col))

    total = len(jobs)
    logging.info(f"Checking 2-mile / 5-mile: {total} address(es) for {len(_RADIUS_TODO)} Uber One store(s)...")
    driver = session.driver
    for n, job in enumerate(jobs.values(), 1):
        session.ensure_tab()
        driver = session.driver
        wanted = {sid for sid, _ in job["targets"]}
        ok, ids, names = False, set(), []
        try:
            ok, ids, names = _uberone.check_address_visibility(driver, job["addr"], wanted, BRAND)
            if _bot_wall_showing(driver) and session.wait_for_human_check(_bot_wall_showing):
                ok, ids, names = _uberone.check_address_visibility(driver, job["addr"], wanted, BRAND)
        except Exception as e:
            _detail.warning(f"2/5-mile check failed for {job['addr']!r}: {e}")
            ok = None
        seen = 0
        for sid, col in job["targets"]:
            rec = results_by_key.get(sid)
            if rec is None:
                continue
            if ok is None:
                rec[col] = 'Error'
            elif not ok:
                rec[col] = 'Address not set'
            elif sid.lower() in ids or any(BRAND in n for n in names):
                # the store's name is shown on the Flowers screen at this address
                rec[col] = 'Available'
                seen += 1
            else:
                closed = str(rec.get('Store_Status', '')).strip().lower() in ('closed', 'force closed')
                rec[col] = 'Not Available (store closed)' if closed else 'Not Available'
        short = str(job["addr"])[:45]
        state = ("address not set" if ok is False else "error" if ok is None
                 else ("store visible" if seen else "store NOT visible"))
        logging.info(f"[2/5-mile {n}/{total}] {short}  ->  {state}")
        time.sleep(random.uniform(1, 2))
    _RADIUS_TODO.clear()


def _store_name(url):
    if '/store/' not in url:
        return 'Unknown Store'
    raw = url.split('/store/')[1].split('/')[0]
    return html.unescape(unquote(raw).replace('-', ' ')).title()


def _find_col(cols, names):
    for c in cols:
        if str(c).strip().lower() in names:
            return c
    return None


if __name__ == "__main__":
    # CRITICAL for a PyInstaller build: without freeze_support(), a frozen
    # Windows .exe can relaunch itself in a loop. Must be the first line here.
    multiprocessing.freeze_support()

    # One-time sign-in:  main.py login   (or UberEatsScraper.exe login)
    if len(sys.argv) > 1 and sys.argv[1].strip().lower() == "login":
        from browser_session import open_login_browser
        if not open_login_browser():
            logging.error("The scraper's Chrome is still open - close it completely, then run login again.")
            time.sleep(10)
            sys.exit(1)
        logging.info(f"Chrome opened with the scraper profile ({settings.profile_dir()}).")
        logging.info("Sign in to Uber Eats there, check address + search + cart show at the top,")
        logging.info("then CLOSE that Chrome. After that, run the scraper normally.")
        time.sleep(5)
        sys.exit(0)

    # Work from the data folder (program folder by default, a mounted volume on
    # cloud) so input.xlsx / output folders are always found.
    os.chdir(DATA_DIR)

    input_file = "input.xlsx"
    if not os.path.exists(input_file):
        logging.error(f"input.xlsx not found in {DATA_DIR}")
        time.sleep(10)          # let the person read it before the console closes
        sys.exit(1)

    input_df = pd.read_excel(input_file)

    # Zip column detection (tolerant header spelling)
    zip_column = _find_col(input_df.columns,
                           {'zipcode', 'zip code', 'zip_code', 'zip', 'postal code', 'postalcode', 'postal_code'})
    if zip_column is None:
        logging.warning(
            f"No zip/postal code column found (columns present: {list(input_df.columns)}). "
            "Every store will show as 'Unknown Timezone'.")

    # --- Duplicate-row protection (pre-scrape), by stable Uber Eats store ID ---
    if 'website_url' in input_df.columns:
        input_df['_dedup_store_id'] = input_df['website_url'].apply(extract_store_id)
        has_id = input_df['_dedup_store_id'].ne('')
        id_rows = input_df[has_id]
        blank_rows = input_df[~has_id]
        dupe_mask = id_rows.duplicated(subset='_dedup_store_id', keep=False)
        if dupe_mask.any():
            dupe_urls = sorted(id_rows.loc[dupe_mask, 'website_url'].astype(str).unique())
            logging.warning(
                f"Found {int(dupe_mask.sum())} duplicate row(s) in input.xlsx by store ID - "
                f"keeping only the first occurrence of each. Duplicated URLs: {dupe_urls}")
        before_rows = len(input_df)
        id_rows = id_rows.drop_duplicates(subset='_dedup_store_id', keep='first')
        input_df = pd.concat([id_rows, blank_rows]).sort_index().drop(columns='_dedup_store_id')
        logging.info(f"Loaded {len(input_df)} row(s) from input.xlsx "
                     f"({before_rows - len(input_df)} duplicate store row(s) removed)")

    # Store's own address - typed into Uber's address bar before each store is
    # read, so the page loads with full details.
    store_addr_column = _find_col(input_df.columns, {'store_address', 'store address', 'store_addr',
                                                     'address', 'store _address'})
    if store_addr_column is None:
        # The current sheet keeps the store's address in 'Store _Name'. Use it
        # when its values really are addresses (house number + comma).
        import re as _re
        cand = _find_col(input_df.columns, {'store _name', 'store_name', 'store name'})
        if cand is not None:
            vals = [str(v) for v in input_df[cand].dropna().astype(str) if str(v).strip()]
            if vals and all(_re.match(r'\s*\d+\s+\S', v) and ',' in v for v in vals):
                store_addr_column = cand
                logging.info(f"Using '{cand}' as the store address column.")
    if store_addr_column is None:
        logging.warning("input.xlsx has no 'store_address' column - stores will be read "
                        "without setting their address first.")

    # Phase 2 (Uber One) address columns
    # Yes / No switch for the slow 2-mile / 5-mile part (first filled cell wins)
    _rc = _find_col(input_df.columns, _RADIUS_COL_NAMES)
    if _rc is not None:
        _vals = [str(v).strip().lower() for v in input_df[_rc].dropna() if str(v).strip()]
        if _vals:
            CHECK_RADIUS["on"] = _vals[0] in ("yes", "y", "true", "1", "on")
    logging.info("2-mile / 5-mile check: " + ("ON (slower run)" if CHECK_RADIUS["on"]
                                               else "OFF (quick run - columns show 'Not checked')"))

    a2_column = _find_col(input_df.columns, {'address_2mile', '2mile', 'address 2mile', 'addr_2mile'})
    a5_column = _find_col(input_df.columns, {'address_5mile', '5mile', 'address 5mile', 'addr_5mile'})
    if ENABLE_UBER_ONE and _uberone is not None and not a2_column and not a5_column:
        logging.warning("Uber One is ON but input.xlsx has no address_2mile / address_5mile "
                        "column - 2mile/5mile will read 'No address'.")

    # WORKER MODE (started by run_all.py, one per Uber account): scrape this
    # account's share once, write the results to a JSON file, and let run_all.py
    # build the ONE merged report + email. No Excel / email / repeat runs here.
    RESULTS_JSON = os.environ.get("UE_RESULTS_JSON", "").strip()
    WORKER_MODE = bool(RESULTS_JSON)

    def _first(col, default):
        try:
            return input_df[col].dropna().iloc[0] if col in input_df.columns else default
        except Exception:
            return default

    # Extract settings safely
    dynamic_recipient = str(_first('email', "recipient@gmail.com"))
    total_expected_runs = 1 if WORKER_MODE else int(_first('runs', 1))
    run_interval = int(_first('interval', 0))
    run_unit = str(_first('unit', "minutes")).lower()
    sleep_seconds = run_interval * (3600 if "hour" in run_unit else (1 if "second" in run_unit else 60))

    # --- ROBUSTNESS TUNING ---
    BASE_DELAY_SECS       = 2      # polite gap between stores on a clean pass
    DELAY_JITTER_SECS     = 2      # random extra gap so traffic isn't perfectly robotic
    # In the signed-in Chrome a "fresh session" is just a fresh WINDOW - the
    # browser, its cookies and the login are never thrown away.
    MAX_CONSECUTIVE_FAILS = 6
    SESSION_COOLDOWN_SECS = 45
    RECOVERY_ROUNDS       = 2
    RECOVERY_WAIT_SECS    = 15
    UNRESOLVED_STATUS     = "Temporarily Unavailable"
    LOGIN_CHECK_EVERY     = max(1, settings.get_int("login_check_every"))

    def build_row_payload(row, status_override=None):
        """Full export record for a store that could not be scraped."""
        url = str(row.get('website_url', '')).strip()
        state, tz = determine_timezone_and_state(row.get(zip_column, '') if zip_column else '')
        return {
            'Store_Status': status_override or "Closed",
            'Status_Note': "",
            'Rating': "", 'Review Count': "",
            'Tag 1': "", 'Tag 2': "", 'Tag 3': "", 'Tag 4': "", 'Tag 5': "", 'Tag 6': "", 'Tag 7': "",
            'All Tags': "", 'Address': "", 'Mobile': "", 'Timings': "",
            'U1 Icon': "", '2mile': "", '5mile': "",
            'Timezone': tz, 'State': state, 'Store Name': _store_name(url),
            'Website_link': url, 'Zipcode': str(row.get(zip_column, '')) if zip_column else ''
        }

    def scrape_one_row(session, row, nav_attempts=3, pin_attempts=3, load_settle=5):
        """Scrape a single row in the scraper window. Returns (payload_or_None, resolved)."""
        url = str(row.get('website_url', '')).strip()
        if not url or url.lower() == 'nan':
            return None, True
        session.ensure_tab()
        state, tz = determine_timezone_and_state(row.get(zip_column, '') if zip_column else '')

        # Step 1 for every store: put the STORE'S OWN address, then read the page.
        store_addr = str(row.get(store_addr_column, '')).strip() if store_addr_column else ''
        addr_result = {"set": None}

        # Speed: if this address was typed in an earlier run, put Uber's saved
        # location in place BEFORE the store page opens - it then loads already
        # at the right address and nothing needs typing.
        def pre_nav(driver, _url):
            if store_addr and store_addr.lower() != 'nan' and _uberone is not None:
                _uberone.put_saved_location(driver, store_addr)
        _scraper_mod.PRE_NAV_HOOK = pre_nav

        def put_store_address(driver, _url):
            if not store_addr or store_addr.lower() == 'nan' or _uberone is None:
                return
            ok = _uberone.apply_address(driver, _url, store_addr)
            if not ok:
                _detail.warning(f"Store address could not be set for {_store_name(_url)}: {store_addr}")
                # The page (and the Uber One logo) only shows properly WITH an
                # address - use the 2-mile address instead so the store still
                # reads correctly.
                a2 = str(row.get(a2_column, '')).strip() if a2_column else ''
                if a2 and a2.lower() != 'nan':
                    ok = _uberone.apply_address(driver, _url, a2)
                    if ok:
                        _detail.info(f"Used the 2-mile address for {_store_name(_url)} instead")
            addr_result["set"] = ok

        payload = scrape_ubereats_store_dynamically(
            session.driver, url, tz, nav_attempts=nav_attempts,
            pin_attempts=pin_attempts, load_settle=load_settle,
            before_scrape=put_store_address)
        if addr_result["set"] is False:
            note = payload.get('Status_Note', '')
            payload['Status_Note'] = (note + " | " if note else "") + "store address not set"
        payload.update({
            'Timezone': tz, 'State': state, 'Store Name': _store_name(url),
            'Website_link': url,
            'Zipcode': str(row.get(zip_column, '')) if zip_column else ''
        })
        resolved = payload.get('Store_Status') != "Load Failed"

        # Uber One: same window, right after the Phase 1 scrape of this store.
        if resolved:
            try:
                payload.update(enrich_with_uber_one(
                    session, url, row, a2_column, a5_column,
                    store_status=payload.get('Store_Status', '')))
            except Exception as e:
                logging.error(f"Uber One enrich error for {url}: {e}")
        return payload, resolved

    # ---- Attach to the logged-in Chrome and open the scraper's own window ----
    try:
        session = ScraperSession()
    except Exception as e:
        logging.error("=" * 70)
        logging.error("Could not start / attach to the scraper's Chrome:")
        for line in str(e).splitlines():
            logging.error("  " + line)
        logging.error("=" * 70)
        time.sleep(10)
        sys.exit(2)

    # Worker started by run_all.py: if the main program is closed (cmd window
    # closed, Stop button), this account stops too and closes its Chrome -
    # nothing keeps running in the background.
    _parent = os.environ.get("UE_PARENT_PID", "").strip()
    if WORKER_MODE and _parent.isdigit():
        import threading

        def _parent_alive(pid):
            if sys.platform.startswith("win"):
                import ctypes
                k32 = ctypes.windll.kernel32
                h = k32.OpenProcess(0x00100000, False, pid)        # SYNCHRONIZE
                if not h:
                    return False
                still = k32.WaitForSingleObject(h, 0) == 0x102     # WAIT_TIMEOUT = running
                k32.CloseHandle(h)
                return still
            try:
                os.kill(pid, 0)
                return True
            except OSError:
                return False

        def _watch_parent(pid=int(_parent)):
            while True:
                time.sleep(3)
                if not _parent_alive(pid):
                    try:
                        session.release()
                    except Exception:
                        pass
                    os._exit(3)

        threading.Thread(target=_watch_parent, daemon=True).start()

    # A robot check during navigation pauses the run for a person (never hidden).
    _scraper_mod.HUMAN_CHECK_HOOK = lambda d: session.wait_for_human_check(_bot_wall_showing)

    try:
        # ---- Verify Uber Eats login BEFORE scraping anything ----
        _first_url = next(
            (str(r.get('website_url', '')).strip() for _, r in input_df.iterrows()
             if str(r.get('website_url', '')).strip()
             and str(r.get('website_url', '')).strip().lower() != 'nan'), None)
        if ENABLE_UBER_ONE and _uberone is not None:
            logging.info("Checking Uber Eats login before starting...")
            if not wait_for_login(session, _first_url):
                logging.error("Not signed in to Uber Eats. Nothing was scraped.")
                time.sleep(15)          # let the person read the steps above
                sys.exit(2)
            logging.info("Uber Eats: signed in. Starting the run.")
            # one-time diagnostic: cookie NAMES (never values) -> scraper.log,
            # so the location-cookie allow-list can be checked against reality
            _scraper_mod.log_cookie_names(session.driver)

        # --- THE MAIN LOOP ---
        for current_run in range(1, total_expected_runs + 1):
            logging.info(f"--- STARTING RUN [{current_run}/{total_expected_runs}] ---")
            results_by_key = {}   # store_id -> payload (ONE record per store)
            _RADIUS_TODO.clear()
            pending = []          # rows that still need a clean scrape
            consecutive_fails = 0
            done_count = 0

            # ---- PASS 1: every store, same window ----
            for idx, row in input_df.iterrows():
                url = str(row.get('website_url', '')).strip()
                if not url or url.lower() == 'nan':
                    continue
                _detail.info(f"Processing Row [{idx+1}]: {url}")

                try:
                    payload, resolved = scrape_one_row(session, row)
                except Exception as e:
                    _detail.error(f"Failed to scrape {url}: {e}")
                    payload, resolved = None, False

                # ONE line per store on screen
                n_total = int(input_df['website_url'].astype(str).str.strip()
                              .replace('nan', '').astype(bool).sum())
                if payload is not None and resolved:
                    u1 = "yes" if payload.get('U1 Icon') else "no"
                    line = f"{payload.get('Store_Status', '')} | Uber One: {u1}"
                    if payload.get('U1 Icon') and payload.get('2mile') == 'Pending':
                        line += " | 2/5 mile: checked at the end"
                else:
                    line = "could not load - will retry at the end"
                logging.info(f"[{done_count + 1}/{n_total}] {_store_name(url)}  ->  {line}")

                if payload is not None and resolved:
                    results_by_key[extract_store_id(url) or url] = payload
                    consecutive_fails = 0
                else:
                    pending.append(row)
                    consecutive_fails += 1

                done_count += 1
                if done_count % LOGIN_CHECK_EVERY == 0:
                    login_watchdog(session, url)

                # Self-heal: repeated misses -> cool down and swap in a brand-new
                # window. Chrome, its cookies and the login are kept.
                if consecutive_fails >= MAX_CONSECUTIVE_FAILS:
                    logging.warning("Repeated load failures - pausing, then opening a fresh scraper window...")
                    time.sleep(SESSION_COOLDOWN_SECS)
                    session.fresh_tab()
                    consecutive_fails = 0

                time.sleep(BASE_DELAY_SECS + random.uniform(0, DELAY_JITTER_SECS))

            # ---- RECOVERY PASSES: only the stragglers, fresh window, more patience ----
            for rnd in range(1, RECOVERY_ROUNDS + 1):
                if not pending:
                    break
                logging.info(f"--- RECOVERY PASS {rnd}/{RECOVERY_ROUNDS}: retrying {len(pending)} store(s) ---")
                time.sleep(RECOVERY_WAIT_SECS)
                session.fresh_tab()
                still_pending = []
                for row in pending:
                    url = str(row.get('website_url', '')).strip()
                    try:
                        payload, resolved = scrape_one_row(session, row, nav_attempts=4,
                                                           pin_attempts=4, load_settle=8)
                    except Exception as e:
                        logging.error(f"Recovery scrape error for {url}: {e}")
                        payload, resolved = None, False
                    if payload is not None and resolved:
                        results_by_key[extract_store_id(url) or url] = payload
                        logging.info(f"Recovered on retry: {url}")
                    else:
                        still_pending.append(row)
                    time.sleep(BASE_DELAY_SECS + random.uniform(1, DELAY_JITTER_SECS + 1))
                pending = still_pending

            # ---- 2-MILE / 5-MILE: home page -> address -> Flowers (once per address) ----
            if ENABLE_UBER_ONE and _uberone is not None and _u1_login["ok"]:
                try:
                    run_radius_phase(session, results_by_key)
                except Exception as e:
                    logging.error(f"2/5-mile phase failed: {e}")
            for rec in results_by_key.values():          # anything left unchecked
                for col in ('2mile', '5mile'):
                    if rec.get(col) == 'Pending':
                        rec[col] = 'Error'

            # ---- FINAL SAFETY NET: never emit "Load Failed" ----
            failed_stores = []
            for row in pending:
                url = str(row.get('website_url', '')).strip()
                if not url or url.lower() == 'nan':
                    continue
                reason = get_failure_reason(url)
                logging.warning(f"Unresolved after all retries -> '{UNRESOLVED_STATUS}': {url}")
                results_by_key[extract_store_id(url) or url] = build_row_payload(row, status_override=UNRESOLVED_STATUS)
                failed_stores.append({
                    "name": _store_name(url),
                    "zip": str(row.get(zip_column, '')) if zip_column else '',
                    "url": url,
                    "reason": reason,
                })

            master_records = list(results_by_key.values())

            # ---- RUN SUMMARY + FAILED-STORE LIST (log file) ----
            total_rows = sum(
                1 for _, r in input_df.iterrows()
                if str(r.get('website_url', '')).strip()
                and str(r.get('website_url', '')).strip().lower() != 'nan')
            n_failed = len(failed_stores)
            n_ok = len(master_records) - n_failed

            logging.info("=" * 70)
            logging.info(f"RUN SUMMARY [{current_run}/{total_expected_runs}]")
            logging.info(f"  Stores in input file : {total_rows}")
            logging.info(f"  Scraped successfully : {n_ok}")
            logging.info(f"  Not scraped          : {n_failed}")
            logging.info(f"  Rows in report       : {len(master_records)}")

            forced = [r for r in master_records
                      if str(r.get('Store_Status', '')).strip().lower() == 'force closed']
            logging.info(f"  Force closed         : {len(forced)}")
            if forced:
                logging.warning("-" * 70)
                logging.warning(f"FORCE-CLOSED STORES ({len(forced)}) - page says not accepting orders:")
                for i, fr in enumerate(forced, 1):
                    note = str(fr.get('Status_Note', '') or 'no status line read')
                    logging.warning(f"  {i:>3}. {fr.get('Store Name', '?')}  ->  {note}")
                    logging.warning(f"       url : {fr.get('Website_link', '')}")

            u1 = [r for r in master_records if r.get('U1 Icon')]
            logging.info(f"  Uber One stores      : {len(u1)}")
            if total_rows and len(master_records) != total_rows:
                logging.error(f"  ROW COUNT MISMATCH - expected {total_rows}, got {len(master_records)}")

            if failed_stores:
                logging.warning("-" * 70)
                logging.warning(f"STORES NOT SCRAPED ({n_failed}) - listed with reason:")
                for i, fs in enumerate(failed_stores, 1):
                    zip_part = f" | zip {fs['zip']}" if fs['zip'] else ""
                    logging.warning(f"  {i:>3}. {fs['name']}{zip_part}")
                    logging.warning(f"       url    : {fs['url']}")
                    logging.warning(f"       reason : {fs['reason']}")
                logging.warning("-" * 70)
            else:
                logging.info("  Not-scraped list     : none - every store scraped cleanly")
            logging.info("=" * 70)

            if WORKER_MODE:
                import json as _json
                with open(RESULTS_JSON, "w", encoding="utf-8") as fh:
                    _json.dump({"records": master_records, "failed": failed_stores},
                               fh, default=str, ensure_ascii=False)
                logging.info("Results handed to the main program.")
            elif master_records:
                saved_file = format_and_export_excel(master_records)
                logging.info(f"Report saved: {os.path.abspath(saved_file)}")
                deliver_report_via_email(saved_file, dynamic_recipient,
                                         store_count=len(master_records))

            if current_run < total_expected_runs:
                time.sleep(sleep_seconds)
    finally:
        # Close ONLY the scraper's window; Chrome, other windows and the login stay.
        session.release()

    logging.info("All sessions complete.")
    # Finished: close the console / terminal. Chrome itself stays as it is.
    logging.shutdown()
    os._exit(0)