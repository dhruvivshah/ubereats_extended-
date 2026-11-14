#timezone_utils.py
# ==========================================
# 2. TIMEZONE LOOKUP
# ==========================================
import logging
import re
import pandas as pd
from datetime import datetime, timedelta
from config import STATE_MAP   # fixed import
import zipcodes


# ==========================================
# 0. DUPLICATE-DETECTION HELPERS
# ==========================================
_STORE_ID_RE = re.compile(r'/store/[^/]+/([^/?]+)')


def extract_store_id(url):
    """
    Uber Eats URLs look like:
        https://www.ubereats.com/store/<slug>/<store-id>?diningMode=DELIVERY
    The segment AFTER the slug is the stable per-listing ID - two URLs with
    different slugs (e.g. a store renamed/relisted on Uber Eats) but the same
    trailing ID are the same listing. Falls back to the slug alone, then the
    full normalized URL, if the URL doesn't match the expected shape.
    """
    if not url:
        return ""
    url = str(url).strip()
    m = _STORE_ID_RE.search(url)
    if m:
        return m.group(1).split('?')[0].lower()
    if '/store/' in url:
        return url.split('/store/')[1].split('/')[0].split('?')[0].lower()
    return url.rstrip('/').lower()


def normalize_text(s):
    """Lowercase, strip to alphanumerics only - for tolerant text comparisons
       (e.g. matching store names/addresses despite punctuation or spacing
       differences)."""
    return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


def determine_timezone_and_state(zip_val):
    if pd.isna(zip_val) or str(zip_val).strip() == "" or str(zip_val).lower() == 'nan':
        return "Unknown State", "Unknown Timezone"
        
    zip_str = str(zip_val).split('.')[0].strip().split('-')[0].strip()
    clean_zip = re.sub(r'[^0-9]', '', zip_str)[:5].zfill(5)
    
    if len(clean_zip) != 5:
        return "Unknown State", "Unknown Timezone"
        
    try:
        
        data = zipcodes.matching(clean_zip)
        if data:
            z_info = data[0]
            abbr = z_info.get('state', 'Unknown').upper()
            state_full = STATE_MAP.get(abbr, abbr)
            raw_tz = z_info.get('timezone', '')
            tz_str = str(raw_tz).lower()

            if 'pacific' in tz_str or raw_tz in [-8, '-8', -8.0]: 
                return state_full, "PT"
            elif 'mountain' in tz_str or raw_tz in [-7, '-7', -7.0]: 
                return state_full, "MST"
            elif 'central' in tz_str or raw_tz in [-6, '-6', -6.0]: 
                return state_full, "CT"
            elif 'eastern' in tz_str or raw_tz in [-5, '-5', -5.0]: 
                return state_full, "EST"

            if abbr in ['CA', 'OR', 'WA', 'NV']:
                return state_full, "PT"
            elif abbr in ['AZ', 'CO', 'UT', 'NM', 'WY', 'MT', 'ID']:
                return state_full, "MST"
            elif abbr in ['TX', 'OK', 'KS', 'NE', 'SD', 'ND', 'MN', 'IA', 'MO', 'AR', 'LA', 'MS', 'AL', 'TN', 'KY', 'IL', 'WI']:
                return state_full, "CT"
            elif abbr in ['FL', 'GA', 'SC', 'NC', 'VA', 'WV', 'MD', 'DE', 'PA', 'NJ', 'NY', 'CT', 'RI', 'MA', 'VT', 'NH', 'ME', 'OH', 'MI', 'IN']:
                return state_full, "EST"
    except Exception:
        pass
    return "Unknown State", "Unknown Timezone"

# ==========================================================
# 1. HELPER ENGINE FOR TIMING RESOLUTION
# ==========================================================

def expand_days(day_part_lower, days_order):
    """Expand ranges like 'Monday - Friday' or 'Every Day' into explicit day list."""
    if "every day" in day_part_lower or "daily" in day_part_lower:
        return days_order
    if "-" in day_part_lower:
        try:
            start_d, end_d = [d.strip() for d in day_part_lower.split('-')]
            if start_d in days_order and end_d in days_order:
                idx_s = days_order.index(start_d)
                idx_e = days_order.index(end_d)
                if idx_s <= idx_e:
                    return days_order[idx_s:idx_e+1]
                else:  # wrap around (e.g. Friday - Monday)
                    return days_order[idx_s:] + days_order[:idx_e+1]
        except Exception:
            return []
    elif day_part_lower in days_order:
        return [day_part_lower]
    return []

def time_in_range(current_time, start_time, end_time):
    """Check if current_time falls between start_time and end_time, handling overnight ranges."""
    if start_time <= end_time:
        return start_time <= current_time <= end_time
    else:  # Overnight rollover case (e.g., 11:00 PM to 2:00 AM)
        return current_time >= start_time or current_time <= end_time


DAYS_ORDER = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

# A full "10:00 AM - 10:00 PM" pair, and the position of the first clock time in
# a string (used to split "Saturday 11:00 AM - 10:00 PM" into day part + times).
_TIME_RANGE_RE = re.compile(
    r'(\d{1,2}:\d{2}\s?(?:AM|PM))\s*-\s*(\d{1,2}:\d{2}\s?(?:AM|PM))', re.IGNORECASE)
_FIRST_TIME_RE = re.compile(r'\d{1,2}:\d{2}\s?(?:AM|PM)', re.IGNORECASE)


def _covered_days(day_part):
    """Expand a day phrase into explicit day names.

    'monday'          -> ['monday']
    'monday - friday' -> ['monday'..'friday']
    'friday - monday' -> wraps around the week
    'every day'/'daily' -> all seven
    """
    low = " ".join((day_part or "").lower().split())
    if not low:
        return []
    if "every day" in low or "everyday" in low or "daily" in low:
        return list(DAYS_ORDER)

    # Pick out day names in the order they appear, so 'Mon - Fri', 'Monday to
    # Friday' and 'Monday-Friday' all resolve the same way.
    found = []
    for m in re.finditer(r'[a-z]+', low):
        w = m.group(0)
        for d in DAYS_ORDER:
            if d == w or (len(w) >= 3 and d.startswith(w)):
                if d not in found:
                    found.append(d)
                break

    if not found:
        return []
    is_range = ("-" in low or " to " in low or "–" in low or "—" in low)
    if is_range and len(found) >= 2:
        i_s, i_e = DAYS_ORDER.index(found[0]), DAYS_ORDER.index(found[-1])
        if i_s <= i_e:
            return DAYS_ORDER[i_s:i_e + 1]
        return DAYS_ORDER[i_s:] + DAYS_ORDER[:i_e + 1]   # wraps (Fri - Mon)
    return found


def calculate_status(timing_string, day_name, store_local_time):
    """Is the store inside its posted hours right now?

    Handles BOTH layouts the hours table can produce, which is the fix here:

      "Saturday ; 11:00 AM - 10:00 PM"            day and times in separate tokens
      "Saturday 11:00 AM - 10:00 PM"              day and times in the SAME token

    The second form used to return "Closed" for a store that was open: a token
    containing a day name was treated purely as a header (set current day, then
    `continue`), so the times sitting in that same token were never read. Since
    the scraper flattens each table row with " ".join(cells), that is exactly the
    shape it produces - so every store built this way read as Closed.

    Also handles "Monday Closed" without letting the following token's times leak
    onto Monday, and multiple ranges per day (split shifts).
    """
    if not timing_string or not timing_string.strip():
        return "Closed"

    target_day = (day_name or "").lower().strip()
    current_time_only = (store_local_time.time()
                         if isinstance(store_local_time, datetime) else store_local_time)

    def ranges_cover_now(text):
        """True if any time range in `text` contains the current time."""
        if re.search(r'\b24\s*(?:/\s*7|hours?)\b', text, re.IGNORECASE):
            return True
        for start_str, end_str in _TIME_RANGE_RE.findall(text):
            try:
                start_t = datetime.strptime(start_str.upper().replace(" ", ""), "%I:%M%p").time()
                end_t = datetime.strptime(end_str.upper().replace(" ", ""), "%I:%M%p").time()
            except Exception as e:
                logging.warning(f"Failed to parse time range '{start_str}-{end_str}': {e}")
                continue
            if time_in_range(current_time_only, start_t, end_t):
                return True
        return False

    tokens = [t.strip() for t in re.split(r'[;\n]', timing_string) if t.strip()]
    pending_days = None          # day phrase awaiting its times in a later token

    for token in tokens:
        low = token.lower()
        has_time = bool(_FIRST_TIME_RE.search(token))
        day_part = token[:_FIRST_TIME_RE.search(token).start()] if has_time else token
        days_here = _covered_days(day_part)

        if days_here and has_time:
            # Day AND times in one token - evaluate immediately. This is the case
            # that used to be silently skipped.
            if target_day in days_here and ranges_cover_now(token):
                return "Open"
            pending_days = None
            continue

        if days_here and not has_time:
            # "Saturday 24 Hours" - a complete statement with no clock times in it.
            if re.search(r'\b24\s*(?:/\s*7|hours?)\b', low):
                if target_day in days_here:
                    return "Open"
                pending_days = None
                continue
            if "closed" in low:
                # "Monday Closed" - a complete statement. Do NOT leave Monday
                # pending, or the next token's times would be applied to it.
                pending_days = None
            else:
                pending_days = days_here
            continue

        if has_time and pending_days:
            # Times belonging to the day phrase seen in an earlier token.
            if target_day in pending_days and "closed" not in low and ranges_cover_now(token):
                return "Open"
            continue

        # A bare "24 hours"/"Open 24 Hours" row under a pending day.
        if pending_days and target_day in pending_days and re.search(
                r'\b24\s*(?:/\s*7|hours?)\b', low):
            return "Open"

    return "Closed"

# ==========================================================
# FORCE-CLOSED DETECTION via Uber's own "next open" line
# ==========================================================
# Uber Eats prints one authoritative status line in the store info panel BEFORE
# the weekly-hours dropdown is expanded:
#
#     "Open until 10:00 PM"        -> serving now
#     "Opens Monday 10:30 AM"      -> not serving; next opens then
#     "Opens Sunday 8:00 AM"       -> not serving; next opens then
#     "Closed"                     -> not serving; no time given
#
# The "Opens <Day> <Time>" value is Uber's own computed NEXT-OPEN, and - unlike
# "Delivery unavailable" - it does NOT depend on the delivery pin, so it is safe
# to act on. The trick is to compare it against the store's POSTED schedule:
#
#   * schedule says the store should be OPEN right now, yet Uber says it next
#     opens at some future time   -> the store has been PAUSED  -> Force Closed
#   * Uber's next-open SKIPS a scheduled opening (opens Monday when the store has
#     Sunday hours)               -> paused through that slot   -> Force Closed
#   * Uber's next-open MATCHES the schedule's next opening       -> normal Closed
#
# A normal overnight closure has Uber and the schedule agreeing to the minute
# (Uber derives its line from the same hours), so any real gap is the tell.

_DAY_ABBR = {
    "mon": "monday", "tue": "tuesday", "wed": "wednesday", "thu": "thursday",
    "fri": "friday", "sat": "saturday", "sun": "sunday",
}


def _parse_clock(text):
    """'10:30 AM' / '8:00AM' / '10 AM' / '10AM' -> datetime.time, or None.

    Minutes are optional: Uber writes both '10:00 AM' and a bare '10 AM'.
    """
    m = re.search(r'(\d{1,2})(?::(\d{2}))?\s*([AP]M)', text, re.IGNORECASE)
    if not m:
        return None
    minute = m.group(2) or "00"
    try:
        return datetime.strptime(f"{m.group(1)}:{minute} {m.group(3).upper()}",
                                 "%I:%M %p").time()
    except Exception:
        return None


def parse_schedule_slots(timing_string):
    """Posted hours -> {day_name: [(open_time, close_time), ...]}.

    Reuses the same day/time token walk as calculate_status, so it accepts every
    layout that produces (day+time in one token, day then time in the next,
    ranges, Every Day, etc.).
    """
    slots = {d: [] for d in DAYS_ORDER}
    if not timing_string or not str(timing_string).strip():
        return slots

    tokens = [t.strip() for t in re.split(r'[;\n]', str(timing_string)) if t.strip()]
    pending_days = None

    for token in tokens:
        low = token.lower()
        tmatch = _FIRST_TIME_RE.search(token)
        has_time = bool(tmatch)
        day_part = token[:tmatch.start()] if has_time else token
        days_here = _covered_days(day_part)

        if days_here and not has_time:
            pending_days = None if "closed" in low else days_here
            continue

        ranges = _TIME_RANGE_RE.findall(token) if has_time else []
        parsed = []
        for s, e in ranges:
            st, en = _parse_clock(s), _parse_clock(e)
            if st and en:
                parsed.append((st, en))

        apply_to = days_here if days_here else (pending_days or [])
        if has_time and apply_to and "closed" not in low:
            for d in apply_to:
                slots[d].extend(parsed)
        # a day+time token closes its own pending; a lone time token consumes it
        pending_days = None
    return slots


def next_scheduled_open(timing_string, store_local_dt, horizon_days=7):
    """Earliest opening strictly after `store_local_dt`, or None.

    Walks today..+7 days over the posted slots and returns the first opening
    datetime in the future. An opening earlier today that has already passed is
    ignored (that slot is over or currently running - either way it is not the
    NEXT open).
    """
    slots = parse_schedule_slots(timing_string)
    if not any(slots.values()):
        return None
    best = None
    for offset in range(horizon_days + 1):
        day_dt = store_local_dt + timedelta(days=offset)
        day_name = day_dt.strftime("%A").lower()
        for open_t, _close_t in slots.get(day_name, []):
            cand = day_dt.replace(hour=open_t.hour, minute=open_t.minute,
                                  second=0, microsecond=0)
            if cand > store_local_dt and (best is None or cand < best):
                best = cand
    return best


def parse_uber_next_open(text, store_local_dt):
    """Uber's info line -> ('open', None) | ('opens', datetime) | ('closed', None) | (None, None).

    'Open until ...' / 'Open now' / 'Open 24 Hours'  -> ('open', None)
    'Opens Monday 10:30 AM' / 'Opens 2:00 PM' / 'Opens today 2 PM'
                                                     -> ('opens', datetime)
    'Closed' with no time                            -> ('closed', None)
    """
    if not text:
        return (None, None)
    low = " ".join(str(text).lower().split())

    if re.search(r'\bopen\s+until\b|\bopen\s+now\b|\bopen\s+24\b|\bopen\s*·', low):
        return ("open", None)

    # Relative countdown: "Opens in 51 min" / "Opens in 2 hours". Uber is saying
    # the store is CLOSED now and reopens after that gap. Return it as its own
    # kind, with the absolute reopen time, so the caller can treat a short
    # near-term reopen (a likely brief pause) differently from a definite
    # day/time reopen.
    m_rel = re.search(r'opens?\s+in\s+(\d+)\s*(min|hour|hr)', low)
    if m_rel:
        qty = int(m_rel.group(1))
        mins = qty if m_rel.group(2).startswith("min") else qty * 60
        return ("opens_soon", store_local_dt + timedelta(minutes=mins))

    # Absolute reopen. Uber phrases this several ways, all handled here:
    #   "Opens Monday 10:30 AM"   "Opens 2:00 PM"   "Opens today 2 PM"
    #   "Opens at 10 AM"          "Available at 11:00 AM"
    #   "Closed • Available at 11:00 AM"  (schedule reopen shown next to Closed)
    if ("opens" in low or "opening" in low or "available at" in low
            or re.search(r'\bcloses?\b.*\bopens?\b', low)):
        clock = _parse_clock(low)
        # which weekday?
        target_day = None
        if "today" in low:
            target_day = store_local_dt.strftime("%A").lower()
        elif "tomorrow" in low:
            target_day = (store_local_dt + timedelta(days=1)).strftime("%A").lower()
        else:
            for full in DAYS_ORDER:
                if full in low:
                    target_day = full
                    break
            if not target_day:
                for ab, full in _DAY_ABBR.items():
                    if re.search(rf'\b{ab}\b', low):
                        target_day = full
                        break

        if target_day and clock:
            # next date on that weekday, at/after now
            today_idx = store_local_dt.weekday()
            tgt_idx = DAYS_ORDER.index(target_day)
            delta = (tgt_idx - today_idx) % 7
            cand = (store_local_dt + timedelta(days=delta)).replace(
                hour=clock.hour, minute=clock.minute, second=0, microsecond=0)
            if cand <= store_local_dt:
                cand += timedelta(days=7)   # that weekday already passed today
            return ("opens", cand)

        if clock and not target_day:
            # time only -> assume today, roll to tomorrow if already past
            cand = store_local_dt.replace(hour=clock.hour, minute=clock.minute,
                                          second=0, microsecond=0)
            if cand <= store_local_dt:
                cand += timedelta(days=1)
            return ("opens", cand)
        return ("opens", None)   # said "opens" but unparseable time

    if re.search(r'\bclosed\b', low):
        return ("closed", None)
    return (None, None)


# Minimum gap between Uber's next-open and the schedule's next-open before we
# call it a pause. A genuine overnight closure has them agreeing to the minute,
# so 45 min comfortably clears parsing/rounding noise while still catching a
# same-day pause ("opens 2 PM" vs scheduled 10 AM) and any day-skip.
_PAUSE_GAP = timedelta(minutes=45)


def classify_force_closed(uber_line, timing_string, store_local_dt):
    """Decide Force Closed vs Closed vs Open from Uber's next-open line + schedule.

    Returns (status, note) where status is 'Open' | 'Force Closed' | 'Closed' |
    None (no verdict - caller keeps its existing logic).
    """
    kind, uber_open_dt = parse_uber_next_open(uber_line, store_local_dt)
    if kind == "open":
        return ("Open", "")
    if kind not in ("opens", "closed", "opens_soon"):
        return (None, "")

    day_name = store_local_dt.strftime("%A")
    open_now = calculate_status(timing_string, day_name, store_local_dt) == "Open"

    # RELATIVE COUNTDOWN ("Opens in 51 min"): a short, near-term reopen.
    #   * schedule says OPEN now  -> the two disagree, but a sub-hour countdown is
    #     ambiguous (a brief pause, a lunch gap the posted hours don't show, or
    #     the store about to resume) - so FLAG it for review rather than declaring
    #     a definite forced closure.
    #   * schedule says CLOSED now -> closed and reopening soon is consistent, an
    #     ordinary out-of-hours closure.
    if kind == "opens_soon":
        if open_now:
            when = uber_open_dt.strftime("%I:%M %p").lstrip("0") if uber_open_dt else "soon"
            return ("Flagged",
                    f"posted hours say OPEN now, but Uber reopens {when} (opening soon)")
        return ("Closed", "")

    # RULE 1 (high confidence): the posted hours say OPEN right now, but Uber's
    # own line says the store next opens at a specific later day/time. Paused.
    if open_now:
        when = (uber_open_dt.strftime("%a %I:%M %p").lstrip("0")
                if uber_open_dt else "later")
        return ("Force Closed", f"paused - posted hours say OPEN now; Uber reopens {when}")

    # RULE 2: closed now, and Uber's next-open is meaningfully LATER than the
    # schedule's next-open -> Uber has skipped a scheduled opening -> paused.
    if kind == "opens" and uber_open_dt is not None:
        sched_open_dt = next_scheduled_open(timing_string, store_local_dt)
        if sched_open_dt is not None and uber_open_dt - sched_open_dt >= _PAUSE_GAP:
            return ("Force Closed",
                    f"paused - schedule reopens {sched_open_dt.strftime('%a %I:%M %p').lstrip('0')}, "
                    f"Uber reopens {uber_open_dt.strftime('%a %I:%M %p').lstrip('0')}")

    # Otherwise Uber and the schedule agree -> ordinary out-of-hours closure.
    return ("Closed", "")


# ==========================================================================
# UNIFIED STORE-STATUS DECISION  (single source of truth)
# ==========================================================================
# One function that takes every status string read off the page plus the posted
# schedule, and returns the final verdict. All the scattered rules live here so
# every phrasing is handled in ONE place and can be exhaustively tested without a
# browser. Priority order (first match wins):
#
#   1. STORE OFF   "not accepting orders", "temporarily closed", "paused", ...
#                  -> Force Closed. The store's own "switched off" statement.
#                  Wins over everything, including a schedule "Open until" line
#                  the paused store still displays.
#   2. OPEN NOW    "Open until 10 PM", "Open now", "Open 24 hours", bare "Open"
#                  -> Open.
#   3. OPENS IN X  relative countdown "Opens in 10 min", "Opens in 2 hours"
#                  -> schedule says open now  -> Flagged (opening soon, review)
#                  -> schedule says closed    -> Closed  (consistent)
#   4. OPENS <WHEN> absolute "Opens Monday 8 AM", "Opens 12:30 PM"
#                  -> schedule says open now  -> Force Closed (paused)
#                  -> reopen skips a scheduled slot -> Force Closed
#                  -> otherwise               -> Closed
#   5. CLOSED      "Closed until 9 AM", bare "Closed"
#                  -> schedule says open now  -> Flagged (disagreement)
#                  -> otherwise               -> Closed
#   6. nothing readable -> fall back to the schedule's own verdict.
#
# FULFILLMENT lines ("delivery unavailable", "out of range") are IGNORED - they
# describe the delivery address, not the store, and were the original source of
# false Force-Closed rows.

STORE_OFF_PHRASES = (
    "not accepting orders", "not currently accepting orders",
    "no longer accepting orders", "not taking orders", "isn't accepting orders",
    "orders paused", "ordering paused", "temporarily closed", "closed for now",
    "too busy", "paused", "on a break", "taking a break",
)

OPEN_PHRASES = ("open until", "open now", "open 24 hours", "open 24/7",
                "open ·", "open -", "open •")

FULFILLMENT_PHRASES = (
    "delivery unavailable", "pickup unavailable", "delivery and pickup unavailable",
    "unavailable for delivery", "unavailable for pickup", "out of range",
    "too far", "does not deliver", "doesn't deliver", "currently unavailable",
    "temporarily unavailable",
)

_REL_OPEN_RE = re.compile(r'opens?\s+in\s+\d+\s*(?:min|hour|hr)', re.IGNORECASE)


def _norm(s):
    return " ".join(str(s or "").split())


def decide_store_status(status_lines, timing_string, store_local_dt):
    """Final store status from all page status strings + the posted schedule.

    status_lines : list of short status strings read from the page (header
                   subtitle, info-panel line, next-open line) in any order;
                   '' / None entries are ignored.
    Returns (status, note) with status in
            {'Open','Closed','Force Closed','Flagged'}.
    """
    lines = [_norm(s) for s in status_lines if s and _norm(s)]
    # de-dup, keep order
    seen, uniq = set(), []
    for s in lines:
        k = s.lower()
        if k not in seen:
            seen.add(k); uniq.append(s)
    pairs = [(s, s.lower()) for s in uniq]

    open_now = calculate_status(
        timing_string, store_local_dt.strftime("%A"), store_local_dt) == "Open"

    def flagnote(base):
        return f"{base} (posted hours say OPEN now)" if open_now else base

    # 1. STORE OFF - highest priority, address-independent.
    for s, low in pairs:
        # a fulfillment line can contain "unavailable" too; store-off phrases are
        # checked on their own explicit wording, and fulfillment is excluded next.
        if any(p in low for p in STORE_OFF_PHRASES):
            return ("Force Closed", flagnote(s))

    # 2. OPEN NOW.
    for s, low in pairs:
        if any(p in low for p in OPEN_PHRASES):
            return ("Open", s)
        # bare "Open" with no closed/opens wording
        if re.search(r'\bopen\b', low) and not re.search(r'\bopens?\b\s+(in|at|on|mon|tue|wed|thu|fri|sat|sun|today|tomorrow|\d)', low) and "closed" not in low:
            return ("Open", s)

    # 3. RELATIVE COUNTDOWN "Opens in X".
    for s, low in pairs:
        if _REL_OPEN_RE.search(low):
            _, dt = parse_uber_next_open(s, store_local_dt)
            if open_now:
                when = dt.strftime("%I:%M %p").lstrip("0") if dt else "soon"
                return ("Flagged", flagnote(f"{s} (Uber reopens {when})"))
            return ("Closed", s)

    # 4. ABSOLUTE "Opens <Day>/<Time>".
    for s, low in pairs:
        kind, dt = parse_uber_next_open(s, store_local_dt)
        if kind == "opens" and dt is not None:
            if open_now:
                when = dt.strftime("%a %I:%M %p").lstrip("0")
                return ("Force Closed", f"{s} - paused; posted hours say OPEN now (Uber reopens {when})")
            sched_dt = next_scheduled_open(timing_string, store_local_dt)
            if sched_dt is not None and dt - sched_dt >= _PAUSE_GAP:
                return ("Force Closed",
                        f"{s} - paused; schedule reopens {sched_dt.strftime('%a %I:%M %p').lstrip('0')}, "
                        f"Uber reopens {dt.strftime('%a %I:%M %p').lstrip('0')}")
            return ("Closed", s)
        if kind == "opens" and dt is None:
            # said "opens" but time unparseable -> treat like a closure
            return ("Flagged" if open_now else "Closed", flagnote(s))

    # 5. CLOSED wording.
    for s, low in pairs:
        if re.search(r'\bclosed\b', low):
            if open_now:
                return ("Flagged", flagnote(s))
            return ("Closed", s)

    # 6. Nothing recognized -> schedule verdict.
    return ("Open" if open_now else "Closed", "")






















# #timezone_utils.py all good just updating the file to remove the excess junk 
# # ==========================================
# # 2. TIMEZONE LOOKUP
# # ==========================================
# import logging
# import re
# import pandas as pd
# from datetime import datetime
# from config import STATE_MAP   # fixed import
# import zipcodes


# # ==========================================
# # 0. DUPLICATE-DETECTION HELPERS
# # ==========================================
# _STORE_ID_RE = re.compile(r'/store/[^/]+/([^/?]+)')


# def extract_store_id(url):
#     """
#     Uber Eats URLs look like:
#         https://www.ubereats.com/store/<slug>/<store-id>?diningMode=DELIVERY
#     The segment AFTER the slug is the stable per-listing ID - two URLs with
#     different slugs (e.g. a store renamed/relisted on Uber Eats) but the same
#     trailing ID are the same listing. Falls back to the slug alone, then the
#     full normalized URL, if the URL doesn't match the expected shape.
#     """
#     if not url:
#         return ""
#     url = str(url).strip()
#     m = _STORE_ID_RE.search(url)
#     if m:
#         return m.group(1).split('?')[0].lower()
#     if '/store/' in url:
#         return url.split('/store/')[1].split('/')[0].split('?')[0].lower()
#     return url.rstrip('/').lower()


# def normalize_text(s):
#     """Lowercase, strip to alphanumerics only - for tolerant text comparisons
#        (e.g. matching store names/addresses despite punctuation or spacing
#        differences)."""
#     return re.sub(r'[^a-z0-9]', '', str(s or '').lower())


# def determine_timezone_and_state(zip_val):
#     if pd.isna(zip_val) or str(zip_val).strip() == "" or str(zip_val).lower() == 'nan':
#         return "Unknown State", "Unknown Timezone"
        
#     zip_str = str(zip_val).split('.')[0].strip().split('-')[0].strip()
#     clean_zip = re.sub(r'[^0-9]', '', zip_str)[:5].zfill(5)
    
#     if len(clean_zip) != 5:
#         return "Unknown State", "Unknown Timezone"
        
#     try:
        
#         data = zipcodes.matching(clean_zip)
#         if data:
#             z_info = data[0]
#             abbr = z_info.get('state', 'Unknown').upper()
#             state_full = STATE_MAP.get(abbr, abbr)
#             raw_tz = z_info.get('timezone', '')
#             tz_str = str(raw_tz).lower()

#             if 'pacific' in tz_str or raw_tz in [-8, '-8', -8.0]: 
#                 return state_full, "PT"
#             elif 'mountain' in tz_str or raw_tz in [-7, '-7', -7.0]: 
#                 return state_full, "MST"
#             elif 'central' in tz_str or raw_tz in [-6, '-6', -6.0]: 
#                 return state_full, "CT"
#             elif 'eastern' in tz_str or raw_tz in [-5, '-5', -5.0]: 
#                 return state_full, "EST"

#             if abbr in ['CA', 'OR', 'WA', 'NV']:
#                 return state_full, "PT"
#             elif abbr in ['AZ', 'CO', 'UT', 'NM', 'WY', 'MT', 'ID']:
#                 return state_full, "MST"
#             elif abbr in ['TX', 'OK', 'KS', 'NE', 'SD', 'ND', 'MN', 'IA', 'MO', 'AR', 'LA', 'MS', 'AL', 'TN', 'KY', 'IL', 'WI']:
#                 return state_full, "CT"
#             elif abbr in ['FL', 'GA', 'SC', 'NC', 'VA', 'WV', 'MD', 'DE', 'PA', 'NJ', 'NY', 'CT', 'RI', 'MA', 'VT', 'NH', 'ME', 'OH', 'MI', 'IN']:
#                 return state_full, "EST"
#     except Exception:
#         pass
#     return "Unknown State", "Unknown Timezone"

# # ==========================================================
# # 1. HELPER ENGINE FOR TIMING RESOLUTION
# # ==========================================================

# def expand_days(day_part_lower, days_order):
#     """Expand ranges like 'Monday - Friday' or 'Every Day' into explicit day list."""
#     if "every day" in day_part_lower or "daily" in day_part_lower:
#         return days_order
#     if "-" in day_part_lower:
#         try:
#             start_d, end_d = [d.strip() for d in day_part_lower.split('-')]
#             if start_d in days_order and end_d in days_order:
#                 idx_s = days_order.index(start_d)
#                 idx_e = days_order.index(end_d)
#                 if idx_s <= idx_e:
#                     return days_order[idx_s:idx_e+1]
#                 else:  # wrap around (e.g. Friday - Monday)
#                     return days_order[idx_s:] + days_order[:idx_e+1]
#         except Exception:
#             return []
#     elif day_part_lower in days_order:
#         return [day_part_lower]
#     return []

# def time_in_range(current_time, start_time, end_time):
#     """Check if current_time falls between start_time and end_time, handling overnight ranges."""
#     if start_time <= end_time:
#         return start_time <= current_time <= end_time
#     else:  # Overnight rollover case (e.g., 11:00 PM to 2:00 AM)
#         return current_time >= start_time or current_time <= end_time


# def calculate_status(timing_string, day_name, store_local_time):
#     """
#     Groups Day headers with all their associated Time blocks correctly,
#     handles multiple ranges per day, and expands 'Everyday/Daily'.
#     """
#     if not timing_string or not timing_string.strip():
#         return "Closed"
        
#     days_order = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
#     target_day = day_name.lower().strip()
#     current_time_only = store_local_time.time() if isinstance(store_local_time, datetime) else store_local_time

#     tokens = [t.strip() for t in re.split(r'[;\n]', timing_string) if t.strip()]
#     current_day_range = None

#     for token in tokens:
#         is_day = (
#             any(d in token.lower() for d in days_order)
#             or "every day" in token.lower()
#             or "everyday" in token.lower()
#             or "daily" in token.lower()
#         )
        
#         if is_day:
#             current_day_range = token
#             continue

#         if current_day_range:
#             day_part_lower = current_day_range.lower().strip()
#             covered_days = []

#             if "every day" in day_part_lower or "everyday" in day_part_lower or "daily" in day_part_lower:
#                 covered_days = days_order
#             elif "-" in day_part_lower:
#                 try:
#                     start_d, end_d = [d.strip() for d in day_part_lower.split('-')]
#                     if start_d in days_order and end_d in days_order:
#                         idx_s = days_order.index(start_d)
#                         idx_e = days_order.index(end_d)
#                         if idx_s <= idx_e:
#                             covered_days = days_order[idx_s:idx_e+1]
#                         else:
#                             covered_days = days_order[idx_s:] + days_order[:idx_e+1]
#                 except Exception as e:
#                     logging.warning(f"Failed to expand day range '{day_part_lower}': {e}")
#                     continue
#             elif day_part_lower in days_order:
#                 covered_days = [day_part_lower]

#             if target_day in covered_days:
#                 if "closed" in token.lower():
#                     continue
#                 if "24" in token.lower() or "24/7" in token.lower():
#                     return "Open"

#                 ranges = re.findall(
#                     r'(\d{1,2}:\d{2}\s?(?:AM|PM))\s*-\s*(\d{1,2}:\d{2}\s?(?:AM|PM))',
#                     token,
#                     re.IGNORECASE
#                 )
#                 for start_str, end_str in ranges:
#                     try:
#                         start_t = datetime.strptime(start_str.upper(), "%I:%M %p").time()
#                         end_t = datetime.strptime(end_str.upper(), "%I:%M %p").time()
#                         if time_in_range(current_time_only, start_t, end_t):
#                             return "Open"
#                     except Exception as e:
#                         logging.warning(f"Failed to parse time range in token '{token}': {e}")
#                         continue

#     return "Closed"