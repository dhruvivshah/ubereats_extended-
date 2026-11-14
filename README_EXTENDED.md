# Uber Eats scraper – Extended (logged-in) version

## What changed from the standard version

| | Standard | Extended |
|---|---|---|
| Chrome | Throwaway logged-out profile, closed at the end | The logged-in scraper profile. The scraper opens **its own window** in it and closes only that window at the end |
| Your other Chrome work | n/a (current build force-closed Chrome) | Untouched. Keep working in your own windows and tabs |
| Between stores | `about:blank` + location-cookie sweep | `about:blank` (only empties the window) + deletes **only** the `uev2.loc*` / `uev2.diningMode*` cookies and location storage keys. Login cookies are never touched |
| Robot check | Store skipped, browser restarted | Run **pauses**, brings the window forward, waits for a person, then retries the same store |
| Repeated failures | Browser restarted (lost cookies, new captcha) | Fresh **window** in the same Chrome. Login and cookies kept |
| Uber One | — | `U1 Icon`, `2mile`, `5mile` columns (same output file) |
| Settings | Hard-coded | `settings.ini` or `UE_*` environment variables, so nothing is rebuilt for cloud |

## Files

- `main.py` – run flow (Phase 1 scrape, then Uber One, per store, in one window)
- `browser_session.py` – **new.** Starts or attaches to the scraper's Chrome, owns the scraper window, pauses for robot checks, reconnects if Chrome was closed
- `settings.py` + `settings.ini` – **new.** All settings that differ between a PC and the cloud
- `scraper.py` – Phase 1 logic unchanged. Adds logged-in mode (strict cookie allow-list, the pin-match check turned off, a robot-check hook)
- `uberone.py` – badge detection (store area only, Uber One ads ignored), address flow, and the 2-mile/5-mile check
- `start_chrome.bat` – opens the scraper's Chrome (Windows)
- `cloud/` – Dockerfile + entrypoint for hosting
- `config.py`, `excel_formatter.py`, `timezone_utils.py` – unchanged

## One-time setup on a PC

1. Put the program folder somewhere fixed. `start_chrome.bat`, `settings.ini`, `input.xlsx` go next to the .exe.
2. Double-click `start_chrome.bat`. A Chrome window opens with a **new, separate profile** (the `ue_chrome_profile` folder).
3. Sign in to Uber Eats in that window. Optionally sign in to Google too, so bookmarks and passwords sync in. This is done **once**; the folder remembers the login.
4. From now on, open this Chrome with `start_chrome.bat`, not the normal Chrome icon. You can use it for normal work.

Why not the everyday Chrome profile: from Chrome 136, Google blocks automation of the default profile, and cookies can't be copied into another profile (they're encrypted per profile). A dedicated profile, signed in once, is the only reliable way.

## Each run

Run the .exe or `python main.py`.
- If the scraper's Chrome is already open, the scraper attaches to it and opens its own window.
- If it isn't open, the scraper starts it (visible) and leaves it running afterwards.
- If the account is signed out, the run waits up to 10 minutes for you to sign in in the scraper window.
- On a robot check, the scraper window comes to the front and the run waits for you.
- Please don't click inside the scraper's window while it is working.

## Output values (Uber One columns)

| Column | Values |
|---|---|
| U1 Icon | `U1icon_present` or blank |
| 2mile / 5mile | `Available` (green), `Not Available` (red), `Not Available (store closed)` (store shut at scrape time, so no delivery time is shown at any address), `No Uber One`, `No address`, `Address not set` (Uber didn't accept the address), `Login required`, `Robot check`, `Error` |

## Cloud

See `cloud/README_CLOUD.md`.
