#scraper.py
import re, json, time, random, logging
from datetime import datetime, timedelta, timezone
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from timezone_utils import (calculate_status, extract_store_id, _covered_days,
                            classify_force_closed, decide_store_status)
from config import TZ_OFFSETS
from zoneinfo import ZoneInfo

# File-only diagnostic logger. Records sent here are written to scraper.log
# (with timestamps) but are kept OFF the console by a filter installed in
# main.py - so the console output stays exactly as it was, while the log file
# gains the minute-by-minute load/verify trail.
_flog = logging.getLogger("scraper.loadwatch")

# Why the most recent load attempt for a given URL failed. Written by
# _load_store_page, read by main.py so the end-of-run failed-store list can
# tell the client EXACTLY why each store could not be scraped.
_LAST_FAILURE_REASON = {}


def get_failure_reason(url, default="Unknown - see scraper.log"):
    """Reason the last load attempt for `url` failed."""
    return _LAST_FAILURE_REASON.get(url, default)

# Map the coarse tz LABELS (PT/MST/CT/EST) to real IANA zones so the store-local
# time is computed WITH daylight saving. The old fixed TZ_OFFSETS ignored DST and
# skewed the open/closed boundary by an hour in summer.
TZ_LABEL_TO_IANA = {
    "PT": "America/Los_Angeles", "PST": "America/Los_Angeles", "PDT": "America/Los_Angeles",
    "MST": "America/Denver", "MDT": "America/Denver",
    "CT": "America/Chicago", "CST": "America/Chicago", "CDT": "America/Chicago",
    "EST": "America/New_York", "EDT": "America/New_York",
}


# Tag scraper
def scrape_tags(driver):
    """
    Scrapes tags from the Uber Eats store info dialog.
    Returns a list of clean tag strings.
    """
    raw_tags_list = []
    try:
        tag_element = driver.find_element(
            By.XPATH,
            "//*[@id='wrapper']/div[4]/div/div/div[2]/div[3]/div[2]/div[2]/div[2]"
        )
        full_tags_string = tag_element.text.strip()

        if full_tags_string:
            normalized_string = full_tags_string.replace("·", "•").replace("|", "•")
            raw_tags_list = [tag.strip() for tag in normalized_string.split("•") if tag.strip()]
            raw_tags_list = [re.sub(r'^[^A-Za-z0-9]+|[^A-Za-z0-9]+$', '', tag) for tag in raw_tags_list]
            raw_tags_list = [tag for tag in raw_tags_list if tag]

    except Exception as e:
        _flog.warning(f"Tag scraping failed: {e}")

    return raw_tags_list


# ==========================================
# Navigation guard (anti stale-duplicate)
# ==========================================
def _backoff(attempt, base=3.0, cap=20.0):
    """Escalating wait + jitter between retries, so we don't hammer Uber Eats
    (which triggers throttling) and so retries spread out enough for a
    rate-limit window to pass."""
    return min(cap, base * attempt) + random.uniform(0, 2)


# ---------------------------------------------------------------------------
# LOCATION STATE CONTROL  (the real fix)
#
# Uber Eats keeps the delivery-location pin in COOKIES / localStorage, NOT in
# the URL. Navigating to a new store URL therefore renders the correct store but
# keeps the PREVIOUS store's delivery address in the header pin - and store
# status is evaluated against that pin. That is the stale-address bug.
#
# Restarting the browser fixes it, but also destroys the bot-clearance cookies,
# which is what triggers the captcha. Instead we delete ONLY the location-ish
# cookies and storage keys, leaving the session (and its clearance) intact.
# ---------------------------------------------------------------------------

# Uber Eats location cookies are prefixed 'uev2.' (e.g. uev2.loc, uev2.diningMode).
# Matched by hint rather than a hard-coded list, so a rename by Uber does not
# silently break the reset. Session/auth cookies do not contain these hints.
_LOCATION_COOKIE_HINTS = ("loc", "diningmode", "address", "geo", "latlng", "place")

# Never delete these even if a hint matches - they carry the session and the
# anti-bot clearance whose loss causes captchas.
_PROTECTED_COOKIE_NAMES = ("sid", "csid", "jwt-session", "_userUuid")

# When True, clear_location_state() does NOTHING. The caller (main.py) sets this
# for LOGGED-IN runs. Two reasons it must be off when signed in:
#   1. Wiping location cookies / localStorage repeatedly on an authenticated
#      session can destabilise or sign out the account (the "cookies get dumped"
#      symptom) - the session must be left completely untouched.
#   2. The stale-pin reset it performs is no longer needed: "Delivery
#      unavailable" / fulfillment phrases are already IGNORED for status, status
#      and hours are address-independent, and the Uber One 2-mile / 5-mile checks
#      REQUIRE the address we set to stay in place.
# For a Phase-1-only (logged-out) run this stays False and behaviour is unchanged.
PRESERVE_SESSION = False

# ---------------------------------------------------------------------------
# EXTENDED (logged-in) MODE - set by main.py when running in the signed-in Chrome.
#
# LOGGED_IN_MODE
#   The store-cookie reset switches to a STRICT allow-list: only Uber Eats'
#   own delivery-location cookies (uev2.loc*, uev2.diningMode*) and location
#   keys in web storage are removed. Every other cookie - the sign-in cookies
#   (sid, csid, jwt-session ...), the bot-clearance cookies, and everything on
#   auth.uber.com - is left alone, so this step cannot log the account out.
#   It deletes cookies one by one by NAME; it never clears the browser.
#
# REQUIRE_PIN_MATCH
#   Phase B of _load_store_page insists the header address pin equals the store's
#   own address. When signed in, the pin shows the ACCOUNT's saved address, never
#   the store's, so that check would reject every store. Logged-in mode turns it
#   off; the store-id check (Phase A) still guarantees the right store loaded, and
#   status/hours are address-independent (fulfillment lines are already ignored).
#
# HUMAN_CHECK_HOOK
#   callable(driver) -> bool. When a robot/captcha wall shows up, the run pauses
#   and lets a person solve it in the visible Chrome instead of giving up.
# ---------------------------------------------------------------------------
LOGGED_IN_MODE = False
REQUIRE_PIN_MATCH = True
HUMAN_CHECK_HOOK = None

_LOGGED_IN_COOKIE_RE = re.compile(r'^uev2\.(loc|diningmode)', re.IGNORECASE)
_LOGGED_IN_PROTECTED = set(_PROTECTED_COOKIE_NAMES) | {
    "jwt-session", "sid", "csid", "_userUuid", "uev2.id.session", "uev2.id.xp",
    "uev2.ts.session", "udi-id", "udi-fingerprint",
}


def clear_location_state(driver):
    """Delete ONLY the delivery-location cookies + storage keys.

    Must run while the browser is on an ubereats.com page (cookies are
    domain-scoped). Session/bot-clearance cookies are preserved, so this does
    NOT provoke a new captcha the way a browser restart does.

    No-op when PRESERVE_SESSION is set (settings: clear_location_between_stores
    = false). In LOGGED_IN_MODE only the strict allow-list above is removed.

    Returns the list of cookie names removed (recorded in the log file).
    """
    if PRESERVE_SESSION:
        _flog.info("LOCATION-RESET skipped (session-preserving mode)")
        return []
    removed = []
    try:
        for cookie in driver.get_cookies():
            name = cookie.get("name", "")
            if name in _PROTECTED_COOKIE_NAMES:
                continue
            if LOGGED_IN_MODE:
                if name in _LOGGED_IN_PROTECTED or not _LOGGED_IN_COOKIE_RE.match(name):
                    continue
            elif not any(hint in name.lower() for hint in _LOCATION_COOKIE_HINTS):
                continue
            try:
                driver.delete_cookie(name)
                removed.append(name)
            except Exception:
                pass
    except Exception as e:
        _flog.warning(f"clear_location_state: cookie sweep failed: {e}")

    # Location is often mirrored into web storage; clear those keys too.
    # (Only keys whose NAME looks like location - never a blanket clear().)
    try:
        driver.execute_script("""
            const rx = /loc|address|geo|latlng|place|dining/i;
            for (const store of [window.localStorage, window.sessionStorage]) {
                try {
                    for (const k of Object.keys(store)) {
                        if (rx.test(k) && !/auth|token|session|user/i.test(k)) store.removeItem(k);
                    }
                } catch (e) {}
            }
        """)
    except Exception:
        pass

    _flog.info(f"LOCATION-RESET removed {len(removed)} cookie(s): {removed or 'none'}")
    return removed


def log_cookie_names(driver):
    """Write the NAMES (never values) of the cookies this page can see to the log,
    so the allow-list above can be checked against what Uber actually sets."""
    try:
        names = sorted({c.get("name", "") for c in driver.get_cookies()})
        _flog.info(f"COOKIE NAMES visible on {driver.current_url[:60]}: {names}")
        return names
    except Exception as e:
        _flog.warning(f"could not list cookie names: {e}")
        return []


# The header pin button, across markup variants Uber Eats has used.
# Ordered most-specific first; the last two are structural fallbacks that find
# the address button in the header even when data-testid/aria-label are absent.
_PIN_SELECTORS = (
    "[data-testid='delivery-address-button']",
    "[data-testid='address-button']",
    "header button[aria-label*='address' i]",
    "header button[aria-label*='location' i]",
    "nav button[aria-label*='address' i]",
    "header button:has(svg) + button",
)


def _read_pin_address(driver):
    """Best-effort read of the delivery-location pin text in the header.

    Returns "" when the pin cannot be located. An unreadable pin is treated as
    'unknown', never as 'stale' - a selector change must not wrongly reject 200
    stores.
    """
    for sel in _PIN_SELECTORS:
        try:
            txt = (driver.find_element(By.CSS_SELECTOR, sel).text or "").strip()
            if txt:
                return txt.split("\n")[0].strip()
        except Exception:
            continue
    return ""


_HOUSE_NUM_RE = re.compile(r"\b(\d{1,6})\b")
# Words skipped when picking the "first real street word", so that abbreviated
# and spelled-out forms of the SAME address produce the same signature:
#   '421 E Speer Blvd'  and  '421 East Speer Boulevard'  -> ('421', 'speer')
# Directionals must appear in BOTH forms, otherwise 'E' -> 'speer' but
# 'East' -> 'east', and a valid store would be falsely rejected.
_STREET_STOPWORDS = {
    "st", "street", "ave", "avenue", "rd", "road", "blvd", "boulevard",
    "dr", "drive", "ln", "lane", "ct", "court", "hwy", "highway", "pkwy",
    "parkway", "way", "pl", "place", "ste", "suite", "unit", "apt", "fl", "floor",
    # directionals - abbreviated AND spelled out
    "n", "s", "e", "w", "ne", "nw", "se", "sw",
    "north", "south", "east", "west",
    "northeast", "northwest", "southeast", "southwest",
}


def _addr_signature(text):
    """Reduce an address to (house_number, first_street_word).

    '51 Jackson St'                      -> ('51', 'jackson')
    '51 Jackson Street, South River, NJ' -> ('51', 'jackson')
    """
    if not text:
        return (None, None)
    low = text.lower()
    m = _HOUSE_NUM_RE.search(low)
    house = m.group(1) if m else None
    tail = low[m.end():] if m else low
    for word in re.findall(r"[a-z]+", tail):
        if word not in _STREET_STOPWORDS:
            return (house, word)
    return (house, None)


def _pin_matches_store(pin_text, store_address):
    """True when the header pin refers to the same street address as the store.

    Returns True when either side is unknown, so a parsing gap never causes a
    false rejection.
    """
    if not pin_text or not store_address:
        return True
    p_house, p_street = _addr_signature(pin_text)
    s_house, s_street = _addr_signature(store_address)
    if not p_house or not s_house:
        return True
    if p_house != s_house:
        return False
    if p_street and s_street and p_street != s_street:
        return False
    return True


def _read_store_address(driver):
    """Read the store's own address as shown on the store page.

    Primary source is the live DOM's copy-address button, which carries the full
    address in its aria-label, e.g.
        aria-label="Store address: 421 East Speer Boulevard, Denver, CO 80203 Copy"
    That attribute is far more stable than any class-based selector.
    """
    try:
        el = driver.find_element(By.CSS_SELECTOR, "button[aria-label^='Store address']")
        label = (el.get_attribute("aria-label") or "").strip()
        m = re.search(r"Store address:\s*(.+?)\s*(?:Copy)?$", label, re.IGNORECASE)
        if m and m.group(1).strip():
            return m.group(1).strip()
    except Exception:
        pass

    for sel in ("[data-testid='store-address']", "h1 ~ div"):
        try:
            txt = (driver.find_element(By.CSS_SELECTOR, sel).text or "").strip()
            if txt and _HOUSE_NUM_RE.search(txt):
                return txt.split("\n")[0].strip()
        except Exception:
            continue
    try:
        body = driver.find_element(By.TAG_NAME, "body").text
        m = re.search(r"^\s*(\d{1,6}\s+[A-Za-z].{3,60},\s*[A-Z]{2}\s*\d{5})",
                      body, re.MULTILINE)
        if m:
            return m.group(1).strip()
    except Exception:
        pass
    return ""


# Text that means Uber Eats served a bot wall / captcha instead of the store.
_BOT_WALL_MARKERS = (
    "access to this page has been denied",
    "verify you are a human",
    "before you continue",
    "unusual traffic",
    "are you a robot",
)

# The toast in the screenshot: it fires on the PREVIOUS (stale) store page when
# a navigation didn't actually switch stores. Its presence next to a store-id
# mismatch is the tell-tale sign of the "went back to the previous website" bug.
_LOCATION_UPDATE_MARKERS = (
    "the store location has been updated",
    "store location has been updated",
    "location has been updated",
)


def _settle_until_ready(driver, max_wait=5.0):
    """Wait only as long as the store content actually needs, instead of a flat
    time.sleep(5) on every store.

    A fixed 5s sleep costs 200 stores x 5s = ~17 minutes of pure waiting per run.
    Most pages are interactive in well under a second. This polls for a real
    store element and returns the moment it appears, falling back to `max_wait`
    only when the page is genuinely slow. Same safety, a fraction of the time.
    """
    deadline = time.time() + max_wait
    while time.time() < deadline:
        try:
            ready = (driver.find_elements(By.CSS_SELECTOR, "button[aria-label^='Store address']")
                     or driver.find_elements(By.TAG_NAME, "h1"))
            if ready:
                time.sleep(0.4)          # small buffer for hydration
                return True
        except Exception:
            pass
        time.sleep(0.25)
    return False


def _load_store_page(driver, url, nav_attempts=3, pin_attempts=3,
                     wait_secs=15, settle_secs=5):
    """
    Load `url` and confirm BOTH invariants before any scraping happens.

    The two failure modes are independent, so they get INDEPENDENT budgets.
    Previously they shared one counter, which meant three stale-pin retries
    silently ate the page-load budget and a slow page had no attempts left.

      PHASE A - "the page won't open" (nav_attempts)
          about:blank -> get(url) -> confirm the store-id in current_url.
          Retried on timeout, redirect, or a stale/previous page.
          A bot wall aborts Phase A immediately: it cannot clear inside this
          session, so burning the remaining attempts is pure waste.

      PHASE B - "the location pin is stale" (pin_attempts)
          Right store, wrong delivery address. Clear the location cookies and
          driver.refresh() - a refresh is cheaper than a fresh navigation, and
          the session (and its bot clearance) is never touched.

    Returns True only when the intended store is confirmed loaded at its own
    address. On failure, the reason is recorded for the run's failed-store list.
    """
    target_id = extract_store_id(url)

    # ---------------- PHASE A: open the right store page ----------------
    on_store = False
    for attempt in range(1, nav_attempts + 1):
        # Kill the previous store's delivery pin while we are still on the
        # ubereats domain from the last store (cookies are domain-scoped).
        try:
            if "ubereats.com" in (driver.current_url or ""):
                clear_location_state(driver)
        except Exception:
            pass

        try:
            # about:blank only empties the TAB so the previous store can't be
            # mistaken for this one. It does NOT clear cookies or log anyone out.
            driver.get("about:blank")          # a failed nav leaves an EMPTY page
            driver.get(url)
            WebDriverWait(driver, wait_secs).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            _settle_until_ready(driver, settle_secs)
        except Exception:
            msg = f"[nav attempt {attempt}/{nav_attempts}] page load timed out: {url}"
            _flog.debug(msg)
            _flog.warning(msg)
            _LAST_FAILURE_REASON[url] = "Page did not open (load timed out)"
            time.sleep(_backoff(attempt))
            continue

        landed_id = extract_store_id(driver.current_url)
        if target_id and landed_id == target_id:
            on_store = True
            break

        # Wrong store: stale page, home redirect, or a bot wall.
        body_low = ""
        try:
            body_low = driver.find_element(By.TAG_NAME, "body").text.lower()
        except Exception:
            pass

        if any(marker in body_low for marker in _LOCATION_UPDATE_MARKERS):
            _flog.warning(
                f"LOCATION-UPDATE notification present on STALE page for {url} "
                f"(store did not switch; landed_id='{landed_id}', expected='{target_id}', "
                f"nav attempt {attempt}/{nav_attempts})"
            )

        if any(marker in body_low for marker in _BOT_WALL_MARKERS):
            bw = f"[nav attempt {attempt}/{nav_attempts}] BOT WALL detected on {url}"
            _flog.debug(bw)
            _flog.warning(bw)
            # Extended mode: the browser is visible - let a person solve it, then
            # retry this store instead of abandoning it.
            if HUMAN_CHECK_HOOK is not None:
                try:
                    if HUMAN_CHECK_HOOK(driver):
                        _flog.info(f"BOT WALL cleared by a person - retrying {url}")
                        continue
                except Exception as e:
                    _flog.warning(f"human-check hook failed: {e}")
            # A bot wall cannot clear inside the SAME session - the clearance is
            # exactly what we lack. Abort now; the caller escalates to a new
            # session. Retrying here would waste the remaining attempts.
            _flog.warning(f"GIVE-UP (early, bot wall) - escalating to a new session: {url}")
            _LAST_FAILURE_REASON[url] = "Bot wall / captcha - needs a fresh session"
            return False

        miss = f"[nav attempt {attempt}/{nav_attempts}] landed on '{landed_id}', expected '{target_id}'"
        _flog.debug(miss)
        _flog.warning(miss)
        _LAST_FAILURE_REASON[url] = (
            f"Landed on '{landed_id or 'unknown'}' instead of '{target_id}' (stale page / redirect)"
        )
        time.sleep(_backoff(attempt))

    if not on_store:
        _flog.warning(f"GIVE-UP after {nav_attempts} nav attempt(s) - page never opened: {url}")
        _LAST_FAILURE_REASON.setdefault(
            url, f"Page never opened after {nav_attempts} attempt(s)")
        return False

    # ---------------- PHASE B: make the delivery pin match ----------------
    if not REQUIRE_PIN_MATCH:
        # Signed in: the pin is the account's address, by design. The store-id
        # check above already proved the right store is on screen.
        _flog.info(f"VERIFIED store '{target_id}' (nav ok; pin check off in logged-in mode)")
        _LAST_FAILURE_REASON.pop(url, None)
        return True

    pin = store_addr = ""
    for attempt in range(1, pin_attempts + 1):
        pin = _read_pin_address(driver)
        store_addr = _read_store_address(driver)

        if _pin_matches_store(pin, store_addr):
            _flog.info(
                f"VERIFIED store '{target_id}' (nav ok, pin attempt {attempt}/{pin_attempts}) "
                f"| pin='{pin or 'unknown'}' | store='{store_addr or 'unknown'}'"
            )
            _LAST_FAILURE_REASON.pop(url, None)
            return True

        _flog.warning(
            f"STALE-PIN on correct store '{target_id}': pin='{pin}' != store='{store_addr}' "
            f"(pin attempt {attempt}/{pin_attempts}) - clearing location cookies and refreshing"
        )
        _LAST_FAILURE_REASON[url] = (
            f"Delivery pin stuck on '{pin}' but store is at '{store_addr}' "
            f"(location cookie would not reset)"
        )

        if attempt == pin_attempts:
            break                       # budget spent; don't pay for one more reload

        try:
            clear_location_state(driver)
            driver.refresh()            # cheaper than about:blank + full navigation
            WebDriverWait(driver, wait_secs).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            _settle_until_ready(driver, settle_secs)
        except Exception:
            _flog.warning(f"pin-reset refresh failed on attempt {attempt}: {url}")
        time.sleep(_backoff(attempt, base=1.5, cap=6.0))

    _flog.warning(
        f"GIVE-UP after {pin_attempts} pin attempt(s) - delivery pin never matched: {url}"
    )
    return False


# ---------------------------------------------------------------------------
# STORE STATUS LINE READER  (forced-closure detection)
#
# Uber Eats prints ONE status line directly under the store title, inside the
# header <section>. The SAME node holds three very different meanings:
#
#   " Open until 10:00 PM"              -> open right now
#   " Closed - Available at 11:00 AM"   -> SCHEDULED closed (outside posted hours)
#   " Delivery unavailable"             -> FORCED closed (owner paused / store
#                                          turned off the tablet / temporarily
#                                          not accepting orders)
#
# A forced closure is NOT a schedule problem: the posted hours still say the
# store should be serving, so it must be reported separately from a store that
# is simply shut for the night. That is the whole point of this reader.
#
# WHY NOT THE CAPTURED XPATHS
#   /html/body/div[1]/div[3]/div[1]/div[2]/main/div/div/section/div[2]/div[1]/p[2]/span[2]
#   #main-content > div > div > section > div._ak._ci._cg._fx... > p:nth-child(4) > span:nth-child(2)
# Both are single-store snapshots: every div[n] / nth-child(n) index shifts as
# soon as the page has one badge more or less (note the two captures already
# disagree - p[2] vs p:nth-child(4), and p:nth-child(3) vs p:nth-child(4) between
# the two states), and ._ak._ci._cg._fx are hashed build classes that change on
# every Uber Eats redeploy.
#
# So the selectors below keep ONLY the parts that are actually stable -
# the #main-content / <main> landmark, the header <section>, and the
# data-testid="rich-text" hook - and drop every positional index and hashed
# class. When even data-testid disappears, the last strategy works the other way
# round: it finds the NODE THAT CONTAINS THE STATUS WORDS and walks UP through
# its ancestor generations (ancestor::p[1] -> ancestor::section[1]) to read the
# whole status line. Text-anchored + ancestor-generated, so it needs no path at
# all.
# ---------------------------------------------------------------------------

# Words that only ever appear in a status line, used as the text anchor for the
# ancestor-walk strategy.
_STATUS_ANCHOR_WORDS = (
    "delivery unavailable", "unavailable", "not accepting", "closed", "open",
    "available at", "opens at", "temporarily",
)

# lower-cased translate() helper for case-insensitive XPath 1.0 matching
_XP_LOWER = ("translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', "
             "'abcdefghijklmnopqrstuvwxyz')")


# TIER 1 - the store status line by landmark + data-testid. Index-free and
# class-free, so it holds across redeploys and across stores.
_STATUS_XPATHS_PRIMARY = (
    # generalised from the captured paths: landmark + header section + p + hook.
    "//*[@id='main-content']//section//p//span[@data-testid='rich-text']",
    "//main//section//p//span[@data-testid='rich-text']",
    # same hook anywhere in the store header.
    "//*[@id='main-content']//span[@data-testid='rich-text']",
    "//main//span[@data-testid='rich-text']",
    # the store subtitle test id, if Uber Eats brings it back.
    "//*[contains(@data-testid, 'store-subtitle')]",
    # the whole status paragraph: the <p> living in the same section as the <h1>.
    "//main//h1/ancestor::section[1]//p",
)


def _status_xpaths_ancestor_generated():
    """TIER 2 - pure ancestor generation, used only when TIER 1 reads nothing.

    No path knowledge at all: anchor on a node whose own text carries a status
    word, then climb one ancestor generation at a time (p -> section) to the
    element that holds the whole status line.
    """
    xp = [f"//main//*[contains({_XP_LOWER}, '{w}')]/ancestor::p[1]"
          for w in _STATUS_ANCHOR_WORDS]
    xp += [f"//main//*[contains({_XP_LOWER}, '{w}')]/ancestor::section[1]"
           for w in ("delivery unavailable", "not accepting")]
    return tuple(xp)


_STATUS_XPATHS_FALLBACK = _status_xpaths_ancestor_generated()

# ---------------------------------------------------------------------------
# WHY "Delivery unavailable" IS NOT A FORCE-CLOSED SIGNAL
#
# It describes the DELIVERY ADDRESS, not the store.
#
# _load_store_page() calls clear_location_state() before every navigation (the
# stale-pin fix), so most stores load with NO delivery address set at all. Uber
# Eats cannot quote delivery to an address that does not exist, so it prints
# " Delivery unavailable" - on a perfectly healthy, open store. And when a pin
# DOES survive the sweep it is the previous store's address, frequently in
# another state and far outside this store's delivery radius, which produces the
# very same line.
#
# Result: the message appears on a large share of normal stores and says nothing
# about whether the store is accepting orders. It is recorded as context in
# Status_Note, and it NEVER decides the status.
# ---------------------------------------------------------------------------
_FULFILLMENT_PHRASES = (
    "delivery unavailable",
    "pickup unavailable",
    "delivery and pickup unavailable",
    "unavailable for delivery",
    "unavailable for pickup",
    "out of range",
    "too far",
    "does not deliver",
    "doesn't deliver",
    "currently unavailable",      # ambiguous in practice - treated as fulfillment
    "temporarily unavailable",    # ambiguous in practice - treated as fulfillment
)

# FORCED closed: the store ITSELF says it is switched off. These lines are about
# the store's own order-taking state and do not change with the delivery pin, so
# they are safe to act on. Deliberately short - a phrase only belongs here once
# it is confirmed to be address-independent.
_STORE_OFF_PHRASES = (
    "temporarily closed",
    "not accepting orders",
    "not currently accepting orders",
    "no longer accepting orders",
    "not taking orders",
    "orders paused",
    "closed for now",
    "too busy",
)

# SCHEDULED closed: "Closed - Available at 11:00 AM" / "Closed until 9:00 AM".
# Uber Eats is telling us WHEN it reopens, which is exactly what a normal
# out-of-hours closure looks like.
_SCHEDULED_CLOSED_RE = re.compile(
    r"closed\b[^a-z]{0,6}(?:.{0,20}?)\b(?:available|opens?|reopens?|until)\b",
    re.IGNORECASE,
)
# NOTE: "opens in ..." is deliberately NOT here. "Opens in 51 min" means the
# store is CLOSED and reopens shortly - the opposite of open. It used to sit in
# this list and wrongly marked such stores Open.
_OPEN_PHRASES = ("open until", "open now", "open 24 hours", "open ·", "open -")


def _classify_status_text(text):
    """Map one status line to a signal.

    Returns one of:
        'Force Closed'  the store says it is not taking orders (address-independent)
        'Open'          the store says it is serving now
        'Closed'        the store says closed, with or without a reopen time
        'Fulfillment'   a delivery/pickup-availability line -> says NOTHING about
                        the store; recorded as context, never used as a status
        'Unknown'       not a status line

    Order matters. 'Fulfillment' is tested BEFORE the closed/open words because
    "Delivery unavailable" must not fall through and be read as a closure.
    """
    low = " ".join((text or "").lower().split())
    if not low:
        return "Unknown"
    for phrase in _STORE_OFF_PHRASES:
        if phrase in low:
            return "Force Closed"
    for phrase in _FULFILLMENT_PHRASES:
        if phrase in low:
            return "Fulfillment"
    if _SCHEDULED_CLOSED_RE.search(low):
        return "Closed"
    if any(p in low for p in _OPEN_PHRASES):
        return "Open"
    # Bare "Closed" with no reopen time given: still a closure, just unexplained.
    if re.search(r"\bclosed\b", low):
        return "Closed"
    if re.search(r"\bopen\b", low):
        return "Open"
    return "Unknown"


def read_surface_status(driver):
    """Read the store's own status line and classify it.

    Returns (signal, raw_text):
        signal   'Open' | 'Force Closed' | 'Closed' | 'Unknown'
        raw_text the exact line read from the page (kept for the log)

    Only SHORT, visible nodes are considered. A status line is a handful of
    words; anything longer is a container that swallowed half the page and would
    match 'closed' from an unrelated menu item.
    """
    fulfillment_text = [""]      # context only; never decides the status

    def scan(xpaths):
        found_signal, found_text = "Unknown", ""
        for xpath in xpaths:
            try:
                elements = driver.find_elements(By.XPATH, xpath)
            except Exception:
                continue
            # The header can carry a dozen rich-text spans (price, fee, rating,
            # tags) before the status line, so the cap has to clear all of them.
            for el in elements[:15]:
                try:
                    if not el.is_displayed():
                        continue
                    txt = " ".join((el.text or "").split())
                except Exception:
                    continue
                if not txt or len(txt) > 120:
                    continue          # too long to be the status line
                signal = _classify_status_text(txt)
                if signal == "Fulfillment":
                    # Remember it for the note, then keep looking for a line that
                    # actually describes the STORE.
                    if not fulfillment_text[0]:
                        fulfillment_text[0] = txt
                    continue
                if signal == "Force Closed":
                    _flog.info(f"SURFACE-STATUS 'Force Closed' <- '{txt}' via {xpath[:70]}")
                    return "Force Closed", txt   # forced beats everything: stop now
                if signal != "Unknown" and found_signal == "Unknown":
                    found_signal, found_text = signal, txt
            # A tier that produced a readable line has already seen the node the
            # forced text would live in (it is the SAME span), so once this tier
            # has an answer there is nothing left to gain from later selectors in
            # it - but keep going within the tier until the tier is exhausted, so
            # a forced line in a sibling node is still caught.
        return found_signal, found_text

    signal, text = scan(_STATUS_XPATHS_PRIMARY)
    if signal == "Unknown":
        # Nothing readable via the data-testid hook -> markup changed. Fall back to
        # finding the status TEXT and generating the path from its ancestors.
        signal, text = scan(_STATUS_XPATHS_FALLBACK)
        if signal != "Unknown":
            _flog.info(f"SURFACE-STATUS via ancestor-generated fallback: '{signal}' <- '{text}'")

    if signal != "Unknown":
        _flog.info(f"SURFACE-STATUS '{signal}' <- '{text}'"
                   + (f" | fulfillment line: '{fulfillment_text[0]}'" if fulfillment_text[0] else ""))
    else:
        _flog.warning(
            "SURFACE-STATUS: no store-state line matched any strategy"
            + (f" (only a fulfillment line was present: '{fulfillment_text[0]}')"
               if fulfillment_text[0] else "")
        )
    return signal, text, fulfillment_text[0]


# ---------------------------------------------------------------------------
# STORE INFO "NEXT OPEN" LINE READER
#
# Inside the store-info panel, ABOVE the weekly-hours dropdown, Uber prints one
# line that is its own live verdict on the store:
#
#     "Open until 10:00 PM"     "Opens Monday 10:30 AM"     "Opens Sunday 8:00 AM"
#
# This is the signal the client identified. It is NOT the delivery line, so it
# does not depend on the pin - it is present whether or not an address is set.
# Read it after the info panel is open, before expanding the hours dropdown.
# ---------------------------------------------------------------------------
# Clock fragment: "10:30 AM", "10 AM", "10AM" - minutes optional.
_CLK = r'\d{1,2}(?::\d{2})?\s*[AP]M'
_NEXT_OPEN_RE = re.compile(
    r'(open\s+until\b[^\n]{0,30}'
    r'|open\s+now\b|open\s+24\s*(?:/\s*7|hours?)\b'
    r'|opens?\s+in\s+\d+\s*(?:min|hour|hr)[a-z]*'                    # relative: "Opens in 51 min"
    r'|(?:opens?|available)\s+(?:today\s+|tomorrow\s+)?at\s+' + _CLK  # "Opens at 10 AM" / "Available at 11:00 AM"
    + r'|opens?\s+(?:today|tomorrow|mon|tue|wed|thu|fri|sat|sun)[a-z]*\s+(?:at\s+)?' + _CLK  # "Opens Monday 10:30 AM"
    + r'|opens?\s+' + _CLK                                           # "Opens 2:00 PM"
    + r'|available\s+at\s+' + _CLK + r')',                           # "Available at 11 AM"
    re.IGNORECASE)


def read_next_open_line(driver):
    """Return the store-info 'Open until ...' / 'Opens <Day> <Time>' line, or ''.

    Text-anchored: it scans short visible strings in the store header / info
    dialog for the one that matches the next-open pattern, rather than depending
    on a positional path. The dialog markup uses the same rotating hashed classes
    as the rest of the page, so a class-based selector would not survive a
    redeploy - the pattern match does.
    """
    # SPEED: the short visible texts are collected INSIDE the browser in one call.
    # The old loop asked Chrome about every element of the whole menu one by one
    # (thousands of round-trips, ~60 s per store when the line was absent).
    # Same scopes, same order, same "short text only" rule.
    scopes = (
        "div[role='dialog'] *",
        "#main-content section *",
        "main *",
    )
    js = """
        const out = [];
        for (const el of document.querySelectorAll(arguments[0])) {
            if (!el.getClientRects().length) continue;
            const cs = getComputedStyle(el);
            if (cs.visibility === 'hidden' || cs.display === 'none') continue;
            const t = (el.innerText || '').replace(/\\s+/g, ' ').trim();
            if (t && t.length <= 60) out.push(t);
        }
        return out;"""
    for scope in scopes:
        try:
            texts = driver.execute_script(js, scope) or []
        except Exception:
            continue
        for txt in texts:
            m = _NEXT_OPEN_RE.search(txt)
            if m:
                found = m.group(1).strip()
                _flog.info(f"NEXT-OPEN line read: '{found}'")
                return found
    _flog.info("NEXT-OPEN line: none found")
    return ""


# ==========================================
# Unified Store Scraper (Status merged)
# ==========================================
def _first_clickable_now(driver, selectors, total_wait, priority_n=0, priority_wait=0.0):
    """SPEED: check all selectors together (in the same priority order) every
    0.25 s, for `total_wait` seconds IN TOTAL. Previously each selector waited
    3 s on its own, so a selector list where only the 10th matches cost ~30 s.
    priority_n / priority_wait: for the first `priority_wait` seconds accept only
    the first `priority_n` (most specific) selectors, so a generic fallback can
    never win just because the specific button rendered a moment later."""
    start = time.time()
    deadline = start + total_wait
    while True:
        active = (selectors[:priority_n]
                  if priority_n and time.time() - start < priority_wait else selectors)
        for how, sel in active:
            try:
                for el in driver.find_elements(how, sel):
                    if el.is_displayed() and el.is_enabled():
                        return el, how, sel
            except Exception:
                continue
        if time.time() >= deadline:
            return None, None, None
        time.sleep(0.25)


def _wait_visible(driver, by, sel, max_wait):
    deadline = time.time() + max_wait
    while time.time() < deadline:
        try:
            if any(e.is_displayed() for e in driver.find_elements(by, sel)):
                return True
        except Exception:
            pass
        time.sleep(0.25)
    return False


def scrape_ubereats_store_dynamically(driver, url, target_tz,
                                     nav_attempts=3, pin_attempts=3, load_settle=5,
                                     before_scrape=None):
    """before_scrape: optional callable(driver, url), run once the store page is
    open and BEFORE anything is read - the extended version uses it to set the
    store's own address so the page loads fully."""
    extracted_fields = {
        'Store_Status': "Closed", 'Status_Note': "", 'Rating': "", 'Review Count': "",
        'Tag 1': "", 'Tag 2': "", 'Tag 3': "", 'Tag 4': "", 'Tag 5': "", 'Tag 6': "", 'Tag 7': "",
        'All Tags': "", 'Address': "", 'Mobile': "", 'Timings': ""
    }

    # Navigate WITH verification. If we cannot confirm we landed on the intended
    # store, return immediately with a distinct status instead of scraping
    # whatever stale page is on screen - this is what used to create duplicates.
    # nav_attempts / pin_attempts are INDEPENDENT budgets; recovery passes raise them.
    if not _load_store_page(driver, url, nav_attempts=nav_attempts,
                            pin_attempts=pin_attempts, settle_secs=load_settle):
        extracted_fields['Store_Status'] = "Load Failed"
        return extracted_fields

    if before_scrape is not None:
        try:
            before_scrape(driver, url)
        except Exception as e:
            _flog.warning(f"before_scrape step failed for {url}: {e}")

    # Dynamic 404/Unavailable Error page validation
    # NOTE: kept the indicators specific. A bare "error" substring was removed because
    # the word can legitimately appear on a healthy page (cookie banners, JS strings)
    # and would wrongly flag a live store as Deactivated.
    try:
        page_text_lower = driver.find_element(By.TAG_NAME, "body").text.lower()
        if any(indicator in page_text_lower for indicator in [
            "can't find this page", "page not found", "404 error",
            "this page could not be found", "something went wrong"
        ]):
            extracted_fields['Store_Status'] = "Deactivated"
            return extracted_fields
    except Exception:
        pass

    # State 2 detection: page LOADS but has NO menu items.
    # Uber Eats serves a generic placeholder ("Nothing to eat here...") for stores
    # that exist but have no published menu. Without this branch, such a store would
    # fall through and get mislabeled as "Closed", indistinguishable from a truly
    # closed store. We give it a distinct status so it can be acted on separately.
    try:
        empty_page_text = driver.find_element(By.TAG_NAME, "body").text.lower()
        empty_menu_markers = [
            "nothing to eat here",
            "let's discover something delicious",
            "let\u2019s discover something delicious",  # curly apostrophe variant
        ]
        if any(marker in empty_page_text for marker in empty_menu_markers):
            extracted_fields['Store_Status'] = "Menu Empty"
            return extracted_fields
    except Exception:
        pass

    # Detect Historical/Permanent Closure Banners
    # e.g. "Closed on Uber Eats as of Apr 6, 2026." This means the store has been
    # removed from the platform. Per client spec this is treated as "Deactivated"
    # (same bucket as a 404 page), since both mean the store is gone from Uber Eats.
    # We keep the banner text in Timings so the closure date is still visible.
    banner_text_found = ""
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
        match = re.search(r'(Closed on Uber Eats as of [A-Za-z]+\s+\d{1,2},\s+\d{4})', body_text, re.IGNORECASE)
        if match:
            banner_text_found = match.group(1)
    except Exception:
        pass

    if banner_text_found:
        extracted_fields['Store_Status'] = "Deactivated"
        extracted_fields['Timings'] = banner_text_found
        return extracted_fields

    # ----------------------------------------------------------------------
    # Live store surface status capture
    #
    # live_surface now has FOUR values, not three:
    #   'Open'          page says it is serving now
    #   'Closed'        page says closed AND tells us when it reopens (schedule)
    #   'Force Closed'  page says "Delivery unavailable" / "not accepting orders"
    #                   -> switched off by the store, schedule is irrelevant
    #   'Unknown'       no readable status line
    # ----------------------------------------------------------------------
    live_surface, surface_text, fulfillment_text = "Unknown", "", ""
    try:
        live_surface, surface_text, fulfillment_text = read_surface_status(driver)
    except Exception as e:
        _flog.warning(f"surface status read failed for {url}: {e}")

    # Read the next-open line NOW, from the store header, BEFORE the info panel
    # is opened. The countdown / "Open until ..." subtitle ("Opens in 51 min")
    # lives on the surface, and once the info modal opens it is hidden behind it -
    # so a read taken only after the click misses exactly this case. Whatever we
    # capture here is preferred; the post-click read is a fallback for the
    # info-panel-only "Opens <Day> <Time>" variant.
    surface_next_open = ""
    try:
        surface_next_open = read_next_open_line(driver)
    except Exception as e:
        _flog.warning(f"surface next-open read failed for {url}: {e}")

    # Fallback: the precise reader found nothing (markup changed beyond every
    # strategy). Fall back to a coarse body-text sniff so a status is still
    # produced rather than a blank. Store-off phrases only - fulfillment wording
    # is excluded here too, for exactly the same reason.
    if live_surface == "Unknown":
        try:
            body_txt = driver.find_element(By.TAG_NAME, "body").text
            head_txt = body_txt.split("More info")[0] if "More info" in body_txt else body_txt[:2000]
            low = head_txt.lower()
            if any(p in low for p in _STORE_OFF_PHRASES):
                live_surface, surface_text = "Force Closed", "(body-text fallback)"
            elif any(k in low for k in ("open until", "open now", "open 24 hours")):
                live_surface, surface_text = "Open", "(body-text fallback)"
            elif any(k in low for k in ("closed until", "closed now", "closed •", "available at")):
                live_surface, surface_text = "Closed", "(body-text fallback)"
        except Exception:
            pass

    # The fulfillment line is kept as CONTEXT so a human can see it in the sheet,
    # clearly marked as an address artefact rather than a store verdict.
    extracted_fields['Status_Note'] = surface_text
    if fulfillment_text and not surface_text:
        extracted_fields['Status_Note'] = f"[address-dependent] {fulfillment_text}"

    # Merge into Store_Status (schedule engine below can still refine it)
    if live_surface == "Open":
        extracted_fields['Store_Status'] = "Active"
    elif live_surface == "Force Closed":
        extracted_fields['Store_Status'] = "Force Closed"
    elif live_surface == "Closed":
        extracted_fields['Store_Status'] = "Closed"
    else:
        extracted_fields['Store_Status'] = "Unknown"

    # --- SCHEMA METADATA SCANNER ---
    try:
        script_elements = driver.find_elements(By.XPATH, "//script[@type='application/ld+json']")
        for element in script_elements:
            try:
                content = element.get_attribute("textContent")
                if not content or "aggregateRating" not in content:
                    continue
                data = json.loads(content.strip())
                if isinstance(data, list): data = data[0]
                if isinstance(data, dict):
                    rating_data = data.get('aggregateRating', {})
                    if isinstance(rating_data, dict):
                        extracted_fields['Rating'] = str(rating_data.get('ratingValue', '')).strip()
                        extracted_fields['Review Count'] = str(rating_data.get('ratingCount', '')).strip()

                    addr = data.get('address', {})
                    if isinstance(addr, dict):
                        extracted_fields['Address'] = f"{addr.get('streetAddress', '')}, {addr.get('addressLocality', '')}, {addr.get('addressRegion', '')} {addr.get('postalCode', '')}".strip(', ')
                    extracted_fields['Mobile'] = data.get('telephone', '')
                    break
            except Exception:
                continue
    except Exception:
        pass

    # --- Surface Rating & Review Splitter ---
    if not extracted_fields['Rating'] or not extracted_fields['Review Count']:
        try:
            rating_selectors = [
                "//div[contains(@data-testid, 'store-subtitle')]",
                "//main//button[contains(., 'ratings') or contains(., 'Rating') or contains(., 'reviews')]",
                "//main//div[contains(., 'ratings') or contains(., 'Rating') or contains(., 'reviews')]",
                "//main//*[contains(., '(') and (contains(., 'ratings') or contains(., 'reviews'))]"
            ]

            raw_rating_text = ""
            for rxpath in rating_selectors:
                elements = driver.find_elements(By.XPATH, rxpath)
                for el in elements:
                    if el.is_displayed():
                        txt = el.text.strip()
                        if txt and any(c.isdigit() for c in txt):
                            raw_rating_text = txt
                            break
                if raw_rating_text: break

            if not raw_rating_text:
                elements = driver.find_elements(By.XPATH, "//main//*[contains(., '(') and contains(., ')')]")
                for el in elements:
                    if el.is_displayed():
                        txt = el.text.strip()
                        if txt and len(txt) < 200 and any(char.isdigit() for char in txt):
                            raw_rating_text = txt
                            break

            if raw_rating_text:
                rating_match = re.search(r'([0-5]\.\d|[0-5])', raw_rating_text)
                if rating_match:
                    extracted_fields['Rating'] = rating_match.group(1)

                count_match = re.search(r'(\d[\d,;\+]*)\+?\s*(?:rating|review)', raw_rating_text, re.IGNORECASE)
                if count_match:
                    found_val = count_match.group(1).strip()
                    if len(found_val) < 10:
                        extracted_fields['Review Count'] = found_val
        except Exception as e:
            _flog.warning(f"Surface rating parsing warning: {e}")

    # --- Info link click ---
    info_clicked = False
    info_link_selectors = [
        "//button[contains(., 'More info') or contains(., 'Store info')]",
        "//main//a[contains(text(), 'info') or contains(text(), 'Info')]",
        "//*[contains(text(), 'More info') or contains(text(), 'Store info')]"
    ]
    info_el, _, _ = _first_clickable_now(
        driver, [(By.XPATH, x) for x in info_link_selectors], total_wait=4)
    if info_el is not None:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", info_el)
            driver.execute_script("arguments[0].click();", info_el)
            # wait for the info panel itself instead of a flat 3 s
            _wait_visible(driver, By.CSS_SELECTOR, "div[role='dialog']", 3)
            time.sleep(0.7)
            info_clicked = True
        except Exception:
            pass

    # --- Next-open line (client's force-closed signal) ---
    # Read it now: the info panel is open but the hours dropdown is NOT yet
    # expanded, which is exactly where "Opens <Day> <Time>" / "Open until ..."
    # lives. Kept for the status decision further down.
    # Prefer the surface read (taken before the modal covered the header). Only
    # fall back to a fresh in-modal read when the surface gave nothing - that is
    # the "Opens <Day> <Time>" row that lives inside the info panel.
    next_open_line = surface_next_open
    if not next_open_line and info_clicked:
        try:
            next_open_line = read_next_open_line(driver)
        except Exception as e:
            _flog.warning(f"next-open read failed: {e}")
    if next_open_line:
        _flog.info(f"NEXT-OPEN used: '{next_open_line}' "
                   f"(surface='{surface_next_open}')")

    # --- Tag extraction ---
    raw_tags_list = []
    if info_clicked:
        raw_tags_list = scrape_tags(driver)

    extracted_fields['All Tags'] = ", ".join(raw_tags_list)
    for i, tag in enumerate(raw_tags_list):
        extracted_fields[f'Tag {i+1}'] = tag

    # --- STEP 2: DROPDOWN TOGGLE ENGINE & TIMINGS EXTRACTION ---
    scraped_hours_raw = ""
    if info_clicked:
        # Toggle buttons, ordered MOST-STABLE first.
        #
        # The hours toggle is the <button aria-expanded="false"> in the store-info
        # dialog whose label reads like "Available Thursday 9:00 AM" / "Open until
        # ...". Matching on that TEXT + aria-expanded survives Uber Eats' rotating
        # obfuscated class names (._d3 ._cd) and shifting child indexes.
        #
        # The last two entries are the exact CSS/XPath paths captured from a live
        # page. They are kept as a literal fallback, but they are brittle by nature
        # (nth-child + hashed classes + absolute path) and WILL break on redeploy -
        # which is why they run last, after the semantic selectors.
        dropdown_selectors = [
            # -- semantic / resilient --
            (By.XPATH, "//button[@aria-expanded]["
                       ".//*[contains(text(), 'Available') or contains(text(), 'Open until') "
                       "or contains(text(), 'Open 24 Hours') or contains(text(), 'Closed until') "
                       "or contains(text(), 'Closed')]]"),
            (By.XPATH, "//div[@role='dialog']//button[@aria-expanded='false']"),
            (By.XPATH, "//*[contains(text(), 'Open until') or contains(text(), 'Open 24 Hours') "
                       "or contains(text(), 'Closed until')]/ancestor::button"),
            (By.XPATH, "//svg[contains(@aria-label, 'Select operating hours')]/ancestor::button"),
            (By.XPATH, "//div[@role='dialog']//button[descendant::svg]"),
            (By.XPATH, "//button[descendant::svg[@data-testid='plus-icon']]"),
            (By.XPATH, "//button[contains(@aria-expanded, 'false') and descendant::div]"),
            (By.XPATH, "//div[@role='dialog']//button"),

            # -- DYNAMIC variants of the live paths below --
            # Same UI as the captured paths, but with the two unstable parts removed
            # so they match EVERY store using this layout, not just one:
            #   * child indexes dropped  (the captured paths disagree with each other:
            #     CSS says div:nth-child(8), XPath says div[4] - the index drifts)
            #   * hashed build classes (._db._dc._cd ...) matched by prefix via
            #     [class*=...] instead of an exact, rotating class list
            # The captured paths target the inner <div>; 'ancestor::button[1]' walks
            # back up to the real clickable <button>.
            (By.XPATH, "//*[@id='wrapper']//button[@aria-expanded]/div/div[2]/ancestor::button[1]"),
            (By.XPATH, "//*[@id='wrapper']//button[.//div[contains(text(), 'Available') "
                       "or contains(text(), 'Open') or contains(text(), 'Closed')]]"),
            (By.XPATH, "//*[@id='wrapper']//div[contains(text(), 'Available') and contains(text(), ':')]"
                       "/ancestor::button[1]"),
            (By.CSS_SELECTOR, "#wrapper button[aria-expanded] > div > div:nth-of-type(2)"),
            (By.CSS_SELECTOR, "#wrapper div[class*='_ps'] > button > div > div[class*='_cd']"),
            (By.CSS_SELECTOR, "#wrapper div[class*='_ps'] > button[aria-expanded]"),

            # -- exact paths captured from the live page (brittle fallback) --
            (By.CSS_SELECTOR,
             "#wrapper > div:nth-child(8) > div > div > div._cg._ak._ce._ep._eo._f8._f9._pe "
             "> div._ae > div._ps > button:nth-child(6) > div "
             "> div._db._dc._dd._de._al._am._cn._bb._pz._q0._q1._cd"),
            (By.XPATH,
             "//*[@id=\"wrapper\"]/div[4]/div/div/div[2]/div[3]/div[2]/button[2]/div/div[2]"),
        ]

        toggle_btn, how, sel = _first_clickable_now(driver, dropdown_selectors, total_wait=4,
                                                    priority_n=1, priority_wait=1.5)
        if toggle_btn is not None:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", toggle_btn)
                time.sleep(0.3)
                driver.execute_script("arguments[0].click();", toggle_btn)
                # wait for the hours table to appear instead of a flat 2.5 s
                _wait_visible(driver, By.TAG_NAME, "table", 2.5)
                time.sleep(0.4)
                _flog.info(f"HOURS-TOGGLE clicked via {how}: {sel[:60]}")
            except Exception:
                _flog.warning("HOURS-TOGGLE: click failed")
        else:
            _flog.warning("HOURS-TOGGLE: no selector matched - hours table may not expand")

        # Stronger, more resilient selectors for hours table
        hours_container_selectors = [
            "//h2[contains(text(), 'Hours')]/following-sibling::div//table",
            "//div[contains(@class, 'hours')]/table",
            "//section[contains(., 'Hours')]//table",
            "//*[@id='wrapper']/div[4]/div/div/div[2]/div[3]/div[2]/table",
            "//table[@class='_of _i3 _og']",
        ]
        for container_xpath in hours_container_selectors:
            try:
                elements = driver.find_elements(By.XPATH, container_xpath)
                if elements and elements[0].is_displayed():
                    rows = elements[0].find_elements(By.TAG_NAME, "tr")
                    parsed_rows, dropped_rows = [], []
                    inside_day_header = False   # last kept row was a day with no times
                    for row in rows:
                        # BOTH th and td. The old code read only <td>, so a table
                        # that puts the weekday in a <th> lost the day entirely -
                        # and hours with no day attached evaluate to Closed for
                        # every store.
                        cells = [c.text.strip() for c in
                                 row.find_elements(By.XPATH, "./th|./td") if c.text.strip()]
                        if not cells:
                            continue
                        line = " ".join(cells)

                        # Keep weekday-bearing rows. Section headers ("Delivery
                        # hours", "Pickup hours", "Special hours") are dropped and
                        # logged, so it is visible exactly what was removed.
                        #
                        # A times-only row ("11:00 AM - 10:00 PM") is kept ONLY
                        # while it directly follows a day-only header row - that is
                        # the two-row layout where the day sits in its own row and
                        # the row below carries its hours. Outside that position a
                        # bare times row has no day to attach to and is noise.
                        has_day = bool(_covered_days(line))
                        has_time = bool(re.search(
                            r'\d{1,2}:\d{2}\s?(?:AM|PM)|\b24\s*(?:/\s*7|hours?)\b',
                            line, re.IGNORECASE))
                        is_closed_word = 'closed' in line.lower()

                        if has_day:
                            parsed_rows.append(line)
                            inside_day_header = not has_time and not is_closed_word
                        elif inside_day_header and (has_time or is_closed_word):
                            parsed_rows.append(line)
                            inside_day_header = False
                        else:
                            dropped_rows.append(line)

                    if dropped_rows:
                        _flog.info(f"HOURS non-weekday row(s) dropped: {dropped_rows}")
                    scraped_hours_raw = " ; ".join(parsed_rows)
                    if scraped_hours_raw:
                        _flog.info(f"HOURS parsed {len(parsed_rows)} weekday row(s): "
                                   f"{scraped_hours_raw[:200]}")
                        break
            except Exception:
                continue

    # Timings column
    extracted_fields['Timings'] = scraped_hours_raw if scraped_hours_raw else "N/A"

    # ======================================================================
    # UNIFIED STATUS DECISION
    #
    # Every status string read off the page - the header subtitle (surface_text),
    # Uber's next-open line (next_open_line), the raw fulfillment line - is handed
    # to ONE function that applies the full priority table (store-off > open >
    # opens-in-X > opens-<day> > closed > schedule). This replaces the old chain
    # of scattered elif branches, so every phrasing is handled in one tested
    # place. It runs whether or not the hours table loaded: with no schedule the
    # store-off / open / closed wordings still resolve; only the open-vs-closed
    # cross-check needs the hours.
    # ======================================================================
    try:
        iana_tz = TZ_LABEL_TO_IANA.get(target_tz, "America/New_York")
        store_local_time = datetime.now(timezone.utc).astimezone(ZoneInfo(iana_tz))

        status_lines = [next_open_line, surface_text, fulfillment_text]
        final_status, final_note = decide_store_status(
            status_lines, scraped_hours_raw, store_local_time)

        extracted_fields['Store_Status'] = final_status
        if final_note:
            extracted_fields['Status_Note'] = final_note
        _flog.info(f"STATUS DECISION -> {final_status} | lines={[l for l in status_lines if l]} "
                   f"| note='{final_note}'")
    except Exception as e:
        _flog.warning(f"status decision failed for {url}: {e}")
        _flog.warning(f"Status calculations warning for {url}: {e}")

    return extracted_fields









