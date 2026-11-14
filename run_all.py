#run_all.py
"""
UberEats Scraper - MAIN PROGRAM (this is what UberEatsScraper.exe runs).

Uses several Uber accounts AT THE SAME TIME so 200 stores finish in ~1.5 h:

    input.xlsx (all stores)
        -> split evenly between the accounts (account 1, 2, 3 ...)
        -> one Chrome window per account, all working side by side
        -> results merged into ONE Excel report + ONE email

How to use (nothing else to learn):
    UberEatsScraper.exe            run the scraper (also what a schedule starts)
    UberEatsScraper.exe login      one-time: sign in each account, one by one
    UberEatsScraper.exe login 2    sign in only account 2 again

Number of accounts: settings.ini -> accounts = 3

Folder layout next to the .exe (created automatically):
    profiles/account1 ...   the signed-in Chrome profile of each account (keep!)
    work/account1 ...       each account's share of stores + its own scraper.log
    ubereats_extracted <date>/   the merged reports (same as before)
    run.log                 one short log of every run
"""
import os
import sys
import json
import time
import shutil
import logging
import threading
import subprocess
import multiprocessing

import pandas as pd

import settings

BASE = settings.app_base_dir()
PROFILES = os.path.join(BASE, "profiles")
WORK = os.path.join(BASE, "work")
BASE_PORT = 9222
FINISH_TIMEOUT_HRS = 6          # a run that takes longer than this is stopped

log = logging.getLogger("runall")


def setup_logging():
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    fh = logging.FileHandler(os.path.join(BASE, "run.log"), encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(logging.Formatter("%(message)s"))
    root.addHandler(fh)
    root.addHandler(ch)


def n_accounts():
    return max(1, settings.get_int("accounts"))


def account_env(k, n):
    """Settings for account k (1-based): own profile, own port, own work folder,
    own spot on the screen."""
    work = os.path.join(WORK, f"account{k}")
    os.makedirs(work, exist_ok=True)
    # Windows are cascaded (each 1280 px wide, shifted a little), NOT squeezed side
    # by side: below ~1000 px Uber switches to its narrow/mobile layout, which
    # would change the page the scraper reads. Overlapping is fine - Chrome is
    # told not to slow down covered windows.
    width = 1920          # wide enough that Uber shows its left menu (with "Flowers")
    env = dict(os.environ)
    env.update({
        "UE_PROFILE_DIR": os.path.join(PROFILES, f"account{k}"),
        "UE_DEBUG_PORT": str(BASE_PORT + k),
        "UE_DATA_DIR": work,
        "UE_RESULTS_JSON": os.path.join(work, "results.json"),
        "UE_WINDOW_SIZE": f"{width},900",
        "UE_WINDOW_POSITION": f"{(k - 1) * 120},{(k - 1) * 60}",
        "UE_OPEN_IN": "window",
        "UE_ACCOUNT_NO": str(k),
        "UE_PARENT_PID": str(os.getpid()),       # account stops when this program stops
        "PYTHONIOENCODING": "utf-8",
    })
    return env


SIGNED_IN_MARK = "signed_in.txt"


def account_ready(k):
    """True when account k has been set up before: either the sign-in step was
    done, or its Chrome profile already exists from earlier runs. A profile that
    exists is ALWAYS used - an existing login is never thrown away or asked for
    again. (If it really is signed out, that account's run says so and its
    stores are given to the other accounts.)"""
    prof = os.path.join(PROFILES, f"account{k}")
    return (os.path.exists(os.path.join(prof, SIGNED_IN_MARK))
            or os.path.isdir(os.path.join(prof, "Default")))


def login_hint(k):
    base = "UberEatsScraper.exe login" if getattr(sys, "frozen", False) else "python run_all.py login"
    return f"{base} {k}" if k else base


def worker_command():
    """How to start one account's scraper: the same .exe (or main.py)."""
    if getattr(sys, "frozen", False):
        return [sys.executable, "--worker"]
    return [sys.executable, os.path.join(BASE, "run_all.py"), "--worker"]


# ---------------------------------------------------------------------------
# One-time sign-in
# ---------------------------------------------------------------------------
def sign_in(only=None):
    from browser_session import find_chrome
    exe = settings.get("chrome_path") or find_chrome()
    if not exe:
        log.error("Google Chrome was not found on this PC.")
        return
    n = n_accounts()
    for k in ([only] if only else range(1, n + 1)):
        prof = os.path.join(PROFILES, f"account{k}")
        os.makedirs(prof, exist_ok=True)
        log.info("=" * 60)
        log.info(f"ACCOUNT {k} of {n}")
        log.info(f"  A Chrome window opens on the Uber Eats sign-in page.")
        log.info(f"  Sign in with the Uber account you use for ACCOUNT {k} (finish any OTP).")
        log.info(f"  Check your address + search + cart show at the top, then CLOSE that Chrome.")
        log.info("=" * 60)
        p = subprocess.Popen([exe, f"--user-data-dir={prof}", "--no-first-run",
                              "--no-default-browser-check", "https://www.ubereats.com/login"])
        # wait until that Chrome is closed (process ended and its profile lock gone)
        locks = [os.path.join(prof, "lockfile"), os.path.join(prof, "SingletonLock")]
        time.sleep(5)
        while p.poll() is None or any(os.path.lexists(x) for x in locks):
            time.sleep(2)
        with open(os.path.join(prof, SIGNED_IN_MARK), "w") as fh:
            fh.write(time.strftime("signed in %Y-%m-%d %H:%M"))
        log.info(f"Account {k}: done.")
    log.info("All accounts signed in. You can now run the scraper.")


# ---------------------------------------------------------------------------
# Split -> run in parallel -> merge
# ---------------------------------------------------------------------------
SHARED_COLS = ("email", "runs", "interval", "unit", "check_2_5_mile")


def split_input(df, n):
    """Remove duplicate stores, then deal rows out like cards (1,2,3,1,2,3...)
    so every account gets a similar mix of stores."""
    from timezone_utils import extract_store_id
    df = df[df["website_url"].astype(str).str.strip().str.lower().ne("nan")
            & df["website_url"].astype(str).str.strip().ne("")].copy()
    df["_id"] = df["website_url"].apply(extract_store_id)
    before = len(df)
    df = df.drop_duplicates(subset="_id", keep="first").drop(columns="_id")
    if before != len(df):
        log.info(f"{before - len(df)} duplicate store row(s) skipped.")
    # the report settings sit on the first row only - copy them to every row
    for c in SHARED_COLS:
        if c in df.columns:
            first = df[c].dropna()
            if len(first):
                df[c] = first.iloc[0]
    df = df.reset_index(drop=True)
    n = min(n, max(1, len(df)))
    return [df.iloc[i::n].reset_index(drop=True) for i in range(n)], len(df)


# Lines from an account that the client does not need to see on screen (they
# are still in that account's work/accountN/scraper.log).
_QUIET = ("Loaded ", "Using '", "Checking Uber Eats login", "Uber Eats: signed in",
          "--- STARTING RUN", "=====", "RUN SUMMARY", "  Stores in input", "  Scraped successfully",
          "  Not scraped", "  Rows in report", "  Force closed", "  Uber One stores",
          "  Not-scraped list", "Results handed", "All sessions complete", "2-mile / 5-mile check:")


def _pump(proc, k):
    """Show each account's one-line-per-store progress, tagged with the account."""
    for raw in iter(proc.stdout.readline, b""):
        line = raw.decode("utf-8", "replace").rstrip()
        if line and not line.startswith(_QUIET):
            print(f"[Account {k}] {line}", flush=True)


def run_parts(parts, accounts):
    """Run part i on account accounts[i], all at once. Returns {k: result or None}."""
    procs = {}
    _RUNNING.clear()
    for part, k in zip(parts, accounts):
        env = account_env(k, n_accounts())
        res = env["UE_RESULTS_JSON"]
        if os.path.exists(res):
            os.remove(res)
        part.to_excel(os.path.join(env["UE_DATA_DIR"], "input.xlsx"), index=False)
        kwargs = dict(env=env, cwd=env["UE_DATA_DIR"], stdout=subprocess.PIPE,
                      stderr=subprocess.STDOUT)
        if sys.platform.startswith("win"):
            kwargs["creationflags"] = 0x08000000        # CREATE_NO_WINDOW: no extra consoles
        p = subprocess.Popen(worker_command(), **kwargs)
        threading.Thread(target=_pump, args=(p, k), daemon=True).start()
        procs[k] = (p, res, len(part))
        _RUNNING[k] = p
        log.info(f"Account {k}: started with {len(part)} store(s).")
        time.sleep(8)          # stagger the starts so the Chromes don't open at the same second

    deadline = time.time() + FINISH_TIMEOUT_HRS * 3600
    results = {}
    for k, (p, res, _) in procs.items():
        # poll instead of p.wait(): on Windows a blocking wait ignores Ctrl+C
        while p.poll() is None:
            if time.time() > deadline:
                log.error(f"Account {k}: took too long - stopped.")
                p.kill()
                break
            time.sleep(1)
        data = None
        if os.path.exists(res):
            try:
                with open(res, encoding="utf-8") as fh:
                    data = json.load(fh)
            except Exception as e:
                log.error(f"Account {k}: could not read its results ({e}).")
        if data is None:
            log.error(f"Account {k}: could not run (exit code {p.returncode}) - the reason is shown "
                      f"above and in work/account{k}/scraper.log. If it says 'NOT signed in', run:  "
                      f"{login_hint(k)}")
        results[k] = data
    return results


_RUNNING = {}      # account number -> its running process (for Ctrl+C)


def stop_and_explain(msg_lines):
    """Show a problem and KEEP the window open so a person can read it
    (double-clicked .exe windows otherwise vanish instantly)."""
    log.error("=" * 60)
    for line in msg_lines:
        log.error(line)
    log.error("=" * 60)
    try:
        if sys.stdin and sys.stdin.isatty():
            input("Press ENTER to close this window...")
        else:
            time.sleep(30)
    except Exception:
        time.sleep(30)


def _close_account_chrome(k):
    """Close account k's Chrome by closing all its windows through its control port."""
    import urllib.request
    port = BASE_PORT + k
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    try:
        with opener.open(f"http://127.0.0.1:{port}/json/list", timeout=2) as r:
            targets = json.loads(r.read().decode("utf-8", "ignore"))
        for t in targets:
            if t.get("type") == "page":
                try:
                    opener.open(f"http://127.0.0.1:{port}/json/close/{t['id']}", timeout=2).close()
                except Exception:
                    pass
    except Exception:
        pass


def stop_everything():
    """Ctrl+C: stop every account and close the scraper's Chrome windows."""
    log.warning("Ctrl+C pressed - stopping all accounts and closing their Chrome windows...")
    for k, p in list(_RUNNING.items()):
        try:
            if p.poll() is None:
                p.kill()
        except Exception:
            pass
    for k in range(1, n_accounts() + 1):
        _close_account_chrome(k)
    log.warning("Stopped. No report was made for this run.")


def one_run(run_no, total_runs):
    from excel_formatter import format_and_export_excel, deliver_report_via_email

    input_path = os.path.join(BASE, "input.xlsx")
    if not os.path.exists(input_path):
        stop_and_explain([f"input.xlsx was not found in {BASE}",
                          "Put input.xlsx in the same folder as UberEatsScraper.exe and start again."])
        return
    df = pd.read_excel(input_path)

    # Only use accounts that have been signed in at least once.
    ready = [k for k in range(1, n_accounts() + 1) if account_ready(k)]
    for k in range(1, n_accounts() + 1):
        if k not in ready:
            log.warning(f"Account {k} is not set up yet - skipped. To add it, run:  "
                        f"{login_hint(k)}")
    if not ready:
        # First time on this PC: offer to do the sign-in right here, so a
        # double-click is all a person needs.
        interactive = False
        try:
            interactive = bool(sys.stdin and sys.stdin.isatty())
        except Exception:
            pass
        if interactive:
            log.warning("=" * 60)
            log.warning("No Uber account is signed in on this PC yet (one-time setup).")
            log.warning("Press ENTER to sign in now - a Chrome window opens for each account.")
            log.warning("=" * 60)
            try:
                input()
            except Exception:
                return
            sign_in()
            ready = [k for k in range(1, n_accounts() + 1) if account_ready(k)]
        if not ready:
            stop_and_explain([
                "No Uber account is signed in on this PC yet.",
                "One-time setup: double-click '1 - Sign in Uber accounts.bat'",
                f"(or run:  {login_hint(None)} )",
                "Sign in to each account in the Chrome window that opens, close it,",
                "then start the scraper again."])
            return
    parts, total = split_input(df, len(ready))
    accounts = ready[:len(parts)]
    recipient = str(df["email"].dropna().iloc[0]) if "email" in df.columns and df["email"].notna().any() \
        else "recipient@gmail.com"

    started = time.time()
    log.info("=" * 60)
    log.info(f"RUN {run_no}/{total_runs}: {total} store(s) on {len(parts)} account(s)")
    rc = "check_2_5_mile"
    on = True
    if rc in df.columns and df[rc].notna().any():
        on = str(df[rc].dropna().iloc[0]).strip().lower() in ("yes", "y", "true", "1", "on")
    log.info("2-mile / 5-mile check: " + ("ON (slower run)" if on else "OFF (quick run)"))
    log.info("Press Ctrl+C at any time to stop everything.")
    log.info("=" * 60)

    results = run_parts(parts, accounts)

    # An account that failed (e.g. signed out): give its stores to a working account.
    ok_accounts = [k for k in accounts if results.get(k) is not None]
    failed = [k for k in accounts if results.get(k) is None]
    if failed and ok_accounts:
        redo = pd.concat([parts[k - 1] for k in failed], ignore_index=True)
        helper = ok_accounts[0]
        log.warning(f"Re-running {len(redo)} store(s) of account(s) {failed} on account {helper}.")
        extra = run_parts([redo], [helper]).get(helper)
        if extra:
            results[helper]["records"] += extra.get("records", [])
            results[helper]["failed"] += extra.get("failed", [])

    records, not_scraped = [], []
    for k in ok_accounts:
        records += results[k].get("records", [])
        not_scraped += results[k].get("failed", [])

    mins = (time.time() - started) / 60
    log.info("=" * 60)
    log.info(f"RUN {run_no} SUMMARY  ({mins:.0f} min)")
    log.info(f"  Stores in input file : {total}")
    log.info(f"  In the report        : {len(records)}")
    log.info(f"  Not scraped          : {len(not_scraped)}")
    log.info(f"  Uber One stores      : {sum(1 for r in records if r.get('U1 Icon'))}")
    for fs in not_scraped:
        log.info(f"    - {fs.get('name')} : {fs.get('reason')}")
    if len(records) < total:
        log.error(f"  {total - len(records)} store(s) MISSING (an account could not run) - see above.")
    log.info("=" * 60)

    if records:
        os.chdir(BASE)
        saved = format_and_export_excel(records)
        log.info(f"Report saved: {os.path.abspath(saved)}")
        deliver_report_via_email(saved, recipient, store_count=len(records))


def main():
    multiprocessing.freeze_support()

    # Worker: one account's share (started by this same program).
    if "--worker" in sys.argv:
        sys.argv = [sys.argv[0]]
        import runpy
        runpy.run_module("main", run_name="__main__")
        return

    os.chdir(BASE)
    setup_logging()

    if len(sys.argv) > 1 and sys.argv[1].lower() == "login":
        only = int(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[2].isdigit() else None
        sign_in(only)
        time.sleep(3)
        return

    df = pd.read_excel(os.path.join(BASE, "input.xlsx")) if os.path.exists(os.path.join(BASE, "input.xlsx")) else None
    total_runs, wait_secs = 1, 0
    if df is not None:
        def first(c, d):
            return df[c].dropna().iloc[0] if c in df.columns and df[c].notna().any() else d
        total_runs = int(first("runs", 1))
        unit = str(first("unit", "minutes")).lower()
        wait_secs = int(first("interval", 0)) * (3600 if "hour" in unit else (1 if "second" in unit else 60))

    try:
        for r in range(1, total_runs + 1):
            one_run(r, total_runs)
            if r < total_runs:
                log.info(f"Next run in {wait_secs // 60} min...  (Ctrl+C to stop)")
                time.sleep(wait_secs)
    except KeyboardInterrupt:
        stop_everything()
        logging.shutdown()
        os._exit(1)

    log.info("All runs complete. This window closes in 15 seconds.")
    logging.shutdown()
    time.sleep(15)
    os._exit(0)


if __name__ == "__main__":
    main()