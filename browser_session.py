#browser_session.py
"""
EXTENDED VERSION - run inside an ALREADY-RUNNING, LOGGED-IN Chrome.

How it works
------------
* The scraper's Chrome uses its own dedicated profile folder (settings
  `profile_dir`), signed in to Uber Eats once by hand and remembered after that.
  Chrome 136+ blocks automation of the everyday default profile, so a dedicated
  folder is required - it behaves like any normal Chrome profile otherwise.
* That Chrome is started with a DevTools port (settings `debug_port`). If it is
  not running yet, this module starts it - VISIBLE, never headless, so a person
  can solve a robot check - and leaves it running after the scrape.
* The scraper ATTACHES to that Chrome and opens ONE new tab (or window, setting
  `open_in`). Every store is loaded in that same tab. Other tabs are never
  touched, so people can keep working in the same Chrome.
* At the end only the scraper's own tab is closed. Chrome itself, the login and
  every other tab stay as they were. We never call driver.quit() here, because
  that is the call that can take the whole browser down with it.

Cloud: the same code runs on a VM / container. Chrome runs on a virtual display
(Xvfb) viewed through noVNC, so robot checks are still solved by a human. Only
settings.ini / UE_* environment variables change.
"""
import os
import sys
import time
import json
import shutil
import logging
import subprocess
import urllib.request

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

import settings

log = logging.getLogger("ue.session")


# ---------------------------------------------------------------------------
# Chrome process
# ---------------------------------------------------------------------------
# Keep the scraper tab working at full speed when it is NOT the visible tab
# (someone is using another tab/window). Without these Chrome throttles timers
# in background tabs and pages load half-way.
_ANTI_THROTTLE_FLAGS = (
    "--disable-background-timer-throttling",
    "--disable-backgrounding-occluded-windows",
    "--disable-renderer-backgrounding",
)


def _debug_url(path="/json/version"):
    return f"http://{settings.get('debug_host')}:{settings.get_int('debug_port')}{path}"


def devtools_alive(timeout=2.0):
    """True when a Chrome is listening on the DevTools port."""
    try:
        # Direct connection - never send a localhost call through a proxy.
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with opener.open(_debug_url(), timeout=timeout) as r:
            return r.status == 200 and "Browser" in json.loads(r.read().decode("utf-8", "ignore"))
    except Exception:
        return False


def find_chrome():
    """Path to the Chrome executable (settings `chrome_path` wins)."""
    configured = settings.get("chrome_path")
    if configured and os.path.exists(configured):
        return configured
    candidates = []
    if sys.platform.startswith("win"):
        for base in (os.environ.get("PROGRAMFILES", r"C:\Program Files"),
                     os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
                     os.environ.get("LOCALAPPDATA", "")):
            if base:
                candidates.append(os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"))
    elif sys.platform == "darwin":
        candidates.append("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
    else:
        for name in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser"):
            p = shutil.which(name)
            if p:
                candidates.append(p)
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return None


def chrome_launch_args():
    """Command-line flags for the scraper's Chrome."""
    args = [
        f"--remote-debugging-port={settings.get_int('debug_port')}",
        f"--user-data-dir={settings.profile_dir()}",
        "--no-first-run",
        "--no-default-browser-check",
        "--window-size=" + (settings.get("window_size") or "1920,1000"),
        *_ANTI_THROTTLE_FLAGS,
    ]
    if settings.get("window_position"):              # side-by-side windows (run_all.py)
        args.append("--window-position=" + settings.get("window_position"))
    if not sys.platform.startswith("win") and hasattr(os, "geteuid") and os.geteuid() == 0:
        # containers often run as root, where Chrome's sandbox cannot start
        args += ["--no-sandbox", "--disable-dev-shm-usage"]
    return args


def ensure_chrome_running(wait_secs=25):
    """Start the scraper's Chrome if nothing is listening on the DevTools port.

    Chrome is started DETACHED: it keeps running after the scraper finishes, so
    the login and the user's own tabs survive.

    Returns True when THIS call started Chrome (its first window is then used by
    the scraper - no extra window), False when Chrome was already running.
    """
    if devtools_alive():
        log.info(f"Attaching to the Chrome already running on port {settings.get('debug_port')}.")
        return False

    exe = find_chrome()
    if not exe:
        raise RuntimeError("Google Chrome was not found. Install it, or set chrome_path in settings.ini.")

    os.makedirs(settings.profile_dir(), exist_ok=True)
    cmd = [exe, *chrome_launch_args(), "about:blank"]
    log.info(f"Starting the scraper's Chrome (profile: {settings.profile_dir()})")
    kwargs = dict(stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
    if sys.platform.startswith("win"):
        kwargs["creationflags"] = 0x00000008 | 0x00000200   # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    else:
        kwargs["start_new_session"] = True
    subprocess.Popen(cmd, **kwargs)

    deadline = time.time() + wait_secs
    while time.time() < deadline:
        if devtools_alive():
            return True
        time.sleep(0.5)

    # Most common cause: this same profile is already open in a Chrome that was
    # started WITHOUT the port (e.g. from a normal shortcut). Chrome then hands
    # the request to that window and the port never opens.
    raise RuntimeError(
        "Chrome started but its control port did not open.\n"
        "       The scraper profile is probably already open in a Chrome window started\n"
        "       from a normal shortcut, or by 'main.py login'. Close THAT Chrome window\n"
        "       and run again - the scraper opens it itself.")


def open_login_browser():
    """First-time sign-in: open the scraper profile in PLAIN Chrome - no control
    port, nothing attached - so Uber's sign-in page sees an ordinary browser.
    The person signs in, then closes that Chrome. Returns False if the scraper's
    Chrome is still running (it must be closed first)."""
    if devtools_alive():
        return False
    exe = find_chrome()
    if not exe:
        raise RuntimeError("Google Chrome was not found. Install it, or set chrome_path in settings.ini.")
    os.makedirs(settings.profile_dir(), exist_ok=True)
    cmd = [exe, f"--user-data-dir={settings.profile_dir()}", "--no-first-run",
           "--no-default-browser-check", "--start-maximized", "https://www.ubereats.com/login"]
    kwargs = dict(stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
    if sys.platform.startswith("win"):
        kwargs["creationflags"] = 0x00000008 | 0x00000200
    else:
        kwargs["start_new_session"] = True
    subprocess.Popen(cmd, **kwargs)
    return True


# ---------------------------------------------------------------------------
# Attach + own tab
# ---------------------------------------------------------------------------
def _attach_driver():
    opts = Options()
    opts.debugger_address = f"{settings.get('debug_host')}:{settings.get_int('debug_port')}"
    # SPEED: driver.get() returns once the page's HTML is ready instead of after
    # every image/tracker has loaded (Uber pages took ~30 s to "finish"). The
    # scraper already waits for the store content itself before reading.
    opts.page_load_strategy = "eager"
    base = settings.app_base_dir()
    for name in ("chromedriver.exe", "chromedriver"):
        local = os.path.join(base, name)
        if os.path.exists(local):
            log.info(f"Using bundled chromedriver: {local}")
            return webdriver.Chrome(service=Service(local), options=opts)
    return webdriver.Chrome(options=opts)          # Selenium Manager matches the installed Chrome


class ScraperSession:
    """One attached Chrome + the scraper's own tab inside it."""

    def __init__(self, page_load_timeout=60):
        started_here = ensure_chrome_running()
        self.started_here = started_here        # True -> the scraper owns this Chrome
        self.driver = _attach_driver()
        self.page_load_timeout = page_load_timeout
        self.open_in = "window" if settings.get("open_in").lower() == "window" else "tab"
        self.handle = None
        if started_here:
            self.use_first_window()     # Chrome was just opened -> its window IS ours
        else:
            self.open_tab()             # Chrome already in use -> exactly one new window

    def use_first_window(self):
        """Work in the window Chrome opened with (about:blank) - no second window."""
        d = self.driver
        handles = d.window_handles
        mine = handles[0]
        for h in handles:                       # prefer the blank start page if
            try:                                # Chrome also restored old tabs
                d.switch_to.window(h)
                if d.current_url in ("about:blank", "chrome://newtab/"):
                    mine = h
                    break
            except Exception:
                continue
        d.switch_to.window(mine)
        self.handle = mine
        d.set_page_load_timeout(self.page_load_timeout)
        self.keep_active()
        log.info("Scraper is using the window of the Chrome it just started.")

    # -- tab lifecycle -------------------------------------------------------
    def open_tab(self):
        """Open the scraper's own tab (or window) and work only in it."""
        d = self.driver
        try:
            d.switch_to.new_window(self.open_in)
        except Exception:
            # the window the driver was pointing at is gone -> anchor on any open one
            handles = d.window_handles
            if not handles:
                raise RuntimeError("The scraper's Chrome has no open windows - was it closed?")
            d.switch_to.window(handles[-1])
            d.switch_to.new_window(self.open_in)
        self.handle = d.current_window_handle
        d.set_page_load_timeout(self.page_load_timeout)
        self.keep_active()
        log.info(f"Scraper {self.open_in} opened in the logged-in Chrome.")

    def keep_active(self):
        """Ask Chrome to treat our tab as active/focused even when it is in the
        background, so pages keep rendering while people use other tabs."""
        for cmd, params in (("Page.setWebLifecycleState", {"state": "active"}),
                            ("Emulation.setFocusEmulationEnabled", {"enabled": True})):
            try:
                self.driver.execute_cdp_cmd(cmd, params)
            except Exception:
                pass

    def ensure_tab(self):
        """Called before every store. Re-opens the scraper tab if someone closed
        it, and re-attaches if the whole Chrome was closed / crashed."""
        try:
            handles = self.driver.window_handles
        except Exception:
            # chromedriver lost Chrome entirely (closed, crashed, PC slept)
            log.warning("Lost the connection to Chrome - starting/attaching again.")
            self.reconnect()
            return
        if self.handle in handles:
            try:
                self.driver.switch_to.window(self.handle)
            except Exception:
                pass
            return
        log.warning(f"The scraper {self.open_in} was closed - opening a new one and carrying on.")
        self.open_tab()

    def reconnect(self):
        try:
            self.driver.service.stop()
        except Exception:
            pass
        started_here = ensure_chrome_running()
        self.started_here = self.started_here or started_here
        self.driver = _attach_driver()
        self.use_first_window() if started_here else self.open_tab()

    def fresh_tab(self):
        """Self-heal: replace the scraper tab with a brand-new one (the old
        'restart the browser' step - without logging out or touching other tabs)."""
        old = self.handle
        try:
            self.driver.switch_to.new_window(self.open_in)
            new = self.driver.current_window_handle
            if old in self.driver.window_handles:
                self.driver.switch_to.window(old)
                self.driver.close()
            self.driver.switch_to.window(new)
            self.handle = new
            self.driver.set_page_load_timeout(self.page_load_timeout)
            self.keep_active()
        except Exception:
            self.open_tab()

    def bring_to_front(self):
        try:
            self.driver.switch_to.window(self.handle)
            self.driver.execute_cdp_cmd("Page.bringToFront", {})
        except Exception:
            pass

    def release(self):
        """End of run.
        * Scraper started Chrome itself -> close that Chrome completely (all its
          windows). The login is saved in the profile folder, so it is kept.
        * Chrome was already open (someone is using it) -> close ONLY the
          scraper's window; their windows stay.
        """
        d = self.driver
        if self.started_here:
            try:
                d.execute_cdp_cmd("Browser.close", {})
            except Exception:
                pass
            try:
                d.service.stop()
            except Exception:
                pass
            for _ in range(20):                 # wait until Chrome is really gone
                if not devtools_alive(timeout=0.5):
                    break
                time.sleep(0.5)
            log.info("Chrome closed (all windows). The Uber login stays saved in the profile.")
            return
        try:
            handles = d.window_handles
            if self.handle in handles:
                d.switch_to.window(self.handle)
                if len(handles) > 1:
                    d.close()
                else:
                    d.get("about:blank")     # last tab: closing it would close Chrome
        except Exception:
            pass
        try:
            d.service.stop()                  # stop chromedriver only; never d.quit()
        except Exception:
            pass
        log.info(f"Scraper {self.open_in} closed. Chrome and the login are left running.")

    # -- human checks --------------------------------------------------------
    def wait_for_human_check(self, still_blocked, timeout=None, poll=5):
        """Robot/captcha check on screen: bring the tab forward and wait for a
        person to solve it. `still_blocked(driver)` -> True while the check shows.
        Returns True once cleared, False on timeout."""
        timeout = timeout or settings.get_int("human_check_wait_secs")
        self.bring_to_front()
        log.warning("=" * 70)
        log.warning("Uber is showing a ROBOT / HUMAN CHECK in the scraper's Chrome.")
        span = f"{timeout // 60} min" if timeout >= 60 else f"{timeout} s"
        log.warning(f"Please solve it in that window. Waiting up to {span}...")
        log.warning("=" * 70)
        waited = 0
        while waited < timeout:
            time.sleep(poll)
            waited += poll
            try:
                if not still_blocked(self.driver):
                    log.info("Human check cleared - continuing.")
                    self.keep_active()
                    return True
            except Exception:
                pass
            if waited % 60 == 0:
                log.info(f"...still waiting for the human check ({waited // 60} min)")
        log.error("Human check was not solved in time.")
        return False


def body_text_lower(driver):
    try:
        return (driver.find_element(By.TAG_NAME, "body").text or "").lower()
    except Exception:
        return ""