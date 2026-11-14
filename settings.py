#settings.py
"""
Runtime settings for the EXTENDED (logged-in) Uber Eats scraper.

Nothing here is baked into the .exe / container image. Every value is resolved
at RUNTIME, first match wins:

    1. environment variable  UE_<NAME>        (cloud / Docker / scheduler)
    2. settings.ini next to the program, section [scraper]
    3. the DEFAULTS below

So the SAME code runs on a team member's Windows PC and on a cloud VM - only
settings.ini (or env vars) differ. No source edits, no rebuild.
"""
import os
import sys
import configparser


def app_base_dir():
    """Folder of the .exe (frozen) or of this source file."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


DEFAULTS = {
    # Folder that holds input.xlsx and receives the output folders + scraper.log.
    # Blank = the program folder (same as before). On cloud, point it at a mounted volume.
    "data_dir": "",

    # Chrome DevTools port the scraper attaches to. The scraper's Chrome is
    # started with this port open (the container does it, or
    # the scraper does it itself if nothing is listening yet).
    "debug_host": "127.0.0.1",
    "debug_port": "9222",

    # DEDICATED Chrome profile folder that stays logged in to Uber Eats.
    # Chrome 136+ refuses remote control of the everyday default profile, so this
    # must be its own folder. Blank = "ue_chrome_profile" next to the program.
    "profile_dir": "",

    # Full path to chrome.exe / google-chrome. Blank = auto-detect.
    "chrome_path": "",

    # Where the scraper works inside that Chrome:
    #   window -> a new window of the SAME logged-in profile (default). People
    #             can keep working in their own windows; Chrome does not slow
    #             down a separate window the way it slows a background tab.
    #   tab    -> a new tab in the existing window instead.
    "open_in": "window",

    # Between stores: delete ONLY the Uber Eats delivery-location cookies and
    # location storage keys. Login cookies are never touched. Set to false if
    # it ever causes trouble - the run still works without it.
    "clear_location_between_stores": "true",

    # When Uber shows a robot / captcha check, the run PAUSES and waits this many
    # seconds for a person to solve it in the visible Chrome window.
    "human_check_wait_secs": "600",

    # How long to wait for the one-time manual sign-in.
    "login_wait_secs": "600",

    # Re-check that the account is still signed in every N stores.
    "login_check_every": "15",

    # How many Uber accounts run side by side (run_all.py). Each account has its
    # own Chrome profile folder: profiles/account1, profiles/account2, ...
    "accounts": "3",

    # Window size / position - set per account by run_all.py so the windows
    # don't sit on top of each other. Normally left blank.
    "window_size": "",
    "window_position": "",
}


def _ini_values():
    path = os.path.join(app_base_dir(), "settings.ini")
    if not os.path.exists(path):
        return {}
    cp = configparser.ConfigParser()
    try:
        cp.read(path, encoding="utf-8")
        if cp.has_section("scraper"):
            return {k.lower(): v for k, v in cp.items("scraper")}
    except Exception:
        pass
    return {}


_INI = _ini_values()


def get(name):
    env = os.environ.get(f"UE_{name.upper()}")
    if env is not None and env.strip() != "":
        return env.strip()
    v = _INI.get(name.lower())
    if v is not None and v.strip() != "":
        return v.strip()
    return DEFAULTS.get(name, "")


def get_int(name):
    try:
        return int(float(get(name)))
    except Exception:
        return int(DEFAULTS[name])


def get_bool(name):
    return str(get(name)).strip().lower() in ("1", "true", "yes", "on")


def data_dir():
    d = get("data_dir") or app_base_dir()
    os.makedirs(d, exist_ok=True)
    return d


def profile_dir():
    d = get("profile_dir") or os.path.join(app_base_dir(), "ue_chrome_profile")
    return os.path.abspath(d)