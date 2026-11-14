#excel_formatter.py
import os
import pandas as pd
import logging
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from config import SMTP_SERVER,SMTP_PORT,SENDER_EMAIL ,SENDER_PASSWORD
from timezone_utils import extract_store_id
import re
from math import ceil


# ==========================================================================
# PRESENTATION RULES - one place that decides how every column is displayed
# ==========================================================================
# Horizontal alignment per column. Short codes and numbers centre; anything
# read as a sentence goes left. Applied to EVERY cell in the used range, empty
# cells included, so a column never looks ragged where a value happens to be
# blank.
COL_ALIGN = {
    'Sr. No': 'center', 'Timezone': 'center', 'State': 'left', 'Store Name': 'left',
    'Store_Status': 'center', 'Status_Note': 'left',
    'Rating': 'center', 'Review Count': 'center',
    'Tag 1': 'center', 'Tag 2': 'center', 'Tag 3': 'center', 'Tag 4': 'center',
    'Tag 5': 'center', 'Tag 6': 'center', 'Tag 7': 'center',
    'All Tags': 'left', 'Address': 'left', 'Mobile': 'center',
    'Timings': 'left', 'Website_link': 'center',
    # Phase 2 (Uber One) columns
    'U1 Icon': 'center', '2mile': 'center', '5mile': 'center',
}

# Columns that must NEVER wrap. A 599-character store URL wrapped into a 22-wide
# column becomes 32 lines and a 500-point row - which is exactly what made the
# sheet unreadable. The link is shown as one line and carries the real URL as a
# clickable hyperlink instead.
NO_WRAP_COLS = {'Sr. No', 'Website_link', 'Mobile', 'Rating', 'Review Count'}

# Hard ceiling on how tall a data row may grow, in wrapped lines. Without a cap,
# one long field dictates the height of the whole row.
MAX_ROW_LINES = 8
LINE_HEIGHT_PTS = 15
ROW_PADDING_PTS = 8

_TIME_RE = re.compile(r'\d{1,2}:\d{2}\s?(?:AM|PM)', re.IGNORECASE)
_DAY_RE = re.compile(
    r'\b(mon|tue|wed|thu|fri|sat|sun)[a-z]*\b|every\s*day|daily', re.IGNORECASE)


def format_timings_for_display(raw):
    """Turn the scraped hours string into one readable line per day.

    In:   'Sunday ; Closed ; Monday - Saturday ; 10:00 AM - 4:00 PM'
    Out:  'Sunday: Closed
           Monday - Saturday: 10:00 AM - 4:00 PM'

    The semicolon-run is technically complete but reads as a wall of text, and
    wraps at arbitrary points mid-time. One day per line wraps predictably and
    can be scanned down the column.
    """
    if not raw or not str(raw).strip():
        return raw
    text = str(raw)
    # Keep any banner text (e.g. "Closed on Uber Eats as of ...") untouched.
    if not _DAY_RE.search(text):
        return text

    tokens = [t.strip(" \t•") for t in re.split(r'[;\n]', text) if t.strip(" \t•")]
    lines, pending_day = [], None

    for tok in tokens:
        has_time = bool(_TIME_RE.search(tok))
        has_day = bool(_DAY_RE.search(tok))
        is_closed = 'closed' in tok.lower()

        if has_day and has_time:
            m = _TIME_RE.search(tok)
            day_part = tok[:m.start()].strip(" -:–")
            time_part = tok[m.start():].strip()
            lines.append(f"{day_part}: {time_part}" if day_part else time_part)
            pending_day = None
        elif has_day and is_closed:
            lines.append(f"{re.sub(r'(?i)closed', '', tok).strip(' -:')}: Closed")
            pending_day = None
        elif has_day:
            pending_day = tok
        elif pending_day and (has_time or is_closed):
            lines.append(f"{pending_day}: {tok}")
            pending_day = None
        else:
            lines.append(tok)

    if pending_day:
        lines.append(pending_day)
    return "\n".join(lines)


def _wrapped_line_count(value, width, wraps):
    """How many display lines a cell needs at this column width."""
    if value is None or str(value) == "":
        return 1
    if not wraps:
        return 1
    usable = max(4, int(width) - 2)
    total = 0
    for para in str(value).split('\n'):
        total += 1 if not para else ceil(len(para) / usable)
    return max(1, total)



# ==========================================
# 4. EXCEL FORMATTING & WORKBOOK STYLING ENGINE
# ==========================================
def format_and_export_excel(records):
    df = pd.DataFrame(records)
    
    # Count before dropping
    before_count = len(df)

    # Pass 0: same store ID, pulled from the URL - catches exact repeat
    # scrapes of the identical Uber Eats listing regardless of how the
    # scraped Store Name/Address text happened to format on each pass.
    # Belt-and-suspenders alongside main.py's own pre-scrape input.xlsx dedup.
    if 'Website_link' in df.columns:
        df['_store_id'] = df['Website_link'].apply(extract_store_id)
        has_store_id = df['_store_id'].ne('')
        removed0 = df[has_store_id & df.duplicated(subset='_store_id', keep='first')]
        if not removed0.empty:
            logging.warning(
                f"Pass 0 (same store ID): dropping {len(removed0)} duplicate row(s): "
                f"{sorted(removed0['Store Name'].astype(str).unique())}"
            )
        df = pd.concat([
            df[~has_store_id],
            df[has_store_id].drop_duplicates(subset='_store_id', keep='first')
        ], ignore_index=True).drop(columns='_store_id')

    # Pass 1: exact repeat scrapes of the same store (identical name + address).
    removed1 = df[df.duplicated(subset=['Store Name', 'Address'], keep='first')]
    if not removed1.empty:
        logging.warning(
            f"Pass 1 (same name + address): dropping {len(removed1)} duplicate row(s): "
            f"{sorted(removed1['Store Name'].astype(str).unique())}"
        )
    df = df.drop_duplicates(subset=['Store Name', 'Address'], keep='first')

    # Pass 2: same physical store listed under a different name/URL in the
    # input sheet -- Store Name differs but Address + Mobile are identical,
    # so Pass 1 alone misses these. Only applied to rows that actually HAVE
    # an address/phone (Deactivated / Menu Empty / etc. rows are blank on
    # both, and blank == blank would wrongly collapse all of those together).
    if 'Address' in df.columns and 'Mobile' in df.columns:
        has_contact_info = (
            df['Address'].astype(str).str.strip().ne('') &
            df['Mobile'].astype(str).str.strip().ne('')
        )
        df_contact = df[has_contact_info]
        removed2 = df_contact[df_contact.duplicated(subset=['Address', 'Mobile'], keep='first')]
        if not removed2.empty:
            logging.warning(
                f"Pass 2 (same address + phone): dropping {len(removed2)} duplicate row(s): "
                f"{sorted(removed2['Store Name'].astype(str).unique())}"
            )
        df_contact = df_contact.drop_duplicates(subset=['Address', 'Mobile'], keep='first')
        df_no_contact = df[~has_contact_info]
        df = pd.concat([df_contact, df_no_contact], ignore_index=True)

    # Pass 3 (SOFT, log-only - does NOT remove rows): same address on file,
    # but Pass 2 couldn't confirm it with a matching phone number (blank
    # Mobile, or the numbers differ). This closes the blind spot where a
    # store with no listed phone slides through Pass 2 untouched - but it's
    # only logged, never auto-merged, because a shared street address (food
    # court, strip mall, multi-unit plaza) can legitimately hold several
    # distinct restaurants. Auto-removing here risks deleting a real store.
    # Review these manually and re-run with a corrected input.xlsx if any
    # turn out to be genuine duplicates.
    if 'Address' in df.columns:
        addr_nonblank = df[df['Address'].astype(str).str.strip().ne('')]
        soft_dupes = addr_nonblank[addr_nonblank.duplicated(subset='Address', keep=False)]
        if not soft_dupes.empty:
            pairs = soft_dupes[['Store Name', 'Address']].drop_duplicates().to_dict('records')
            logging.warning(
                "POSSIBLE duplicates (same address, unconfirmed by phone - review "
                f"manually, NOT auto-removed): {pairs}"
            )

    # Count after dropping
    after_count = len(df)
    removed_count = before_count - after_count

    if removed_count > 0:
        logging.warning(f"Removed {removed_count} duplicate row(s) total before export.")

    # Convert back to list of dicts for your existing Excel writing logic
    master_records = df.to_dict(orient='records')
    
    def get_status_sort_val(status_val):
        s = str(status_val).strip().lower()
        if s == 'flagged': return -1     # mismatches float to the very top
        if '404' in s or s in ('deactivated', 'off-platform'): return 0
        # Force Closed = store switched itself off ("Delivery unavailable" /
        # "not accepting orders"). Actionable, so it sits ABOVE the ordinary
        # out-of-hours closures.
        if s == 'force closed': return 1
        if s == 'closed': return 2       # Closed stores near the top (below deactivated)
        if s in ('active', 'open'): return 3  # Open / Active stores below closed
        if s == 'menu empty': return 4   # Menu empty at the bottom
        return 5  # Unknown/other statuses at the very bottom
        
    df['status_sort'] = df['Store_Status'].apply(get_status_sort_val)
    
    # 'Status_Note' carries the store's OWN status line, exactly as the page
    # printed it ("Delivery unavailable", "Closed - Available at 11:00 AM"). It is
    # what makes a Force Closed row auditable: the client can see the store's own
    # words next to the verdict instead of taking the label on trust.
    # 'Sr. No' numbers the stores CONTINUOUSLY down the whole sheet, across zone
    # blocks - not restarting per zone - so "store 87" identifies one row in the
    # file with no further qualification. It is worth having precisely because
    # Excel's own row numbers cannot do that job here: the banner and separator
    # rows consume row numbers, so the 87th store is never on row 87.
    # Phase 2 (Uber One) adds three columns at the end:
    #   'U1 Icon' - 'U1icon_present' when the store shows the Uber One badge
    #   '2mile' / '5mile' - 'Available' / 'Not Available' delivery at those
    #   addresses (green / red). Blank for every store when the extended run is
    #   off, so the report is identical to before whenever Phase 2 isn't used.
    original_columns = [
    'Sr. No',
    'Timezone', 'State', 'Store Name', 'Store_Status', 'Status_Note', 'Rating', 'Review Count',
    'Tag 1', 'Tag 2', 'Tag 3', 'Tag 4', 'Tag 5', 'Tag 6', 'Tag 7',
    'All Tags', 'Address', 'Mobile', 'Timings', 'Website_link',
    'U1 Icon', '2mile', '5mile'
]

    
    # Zone blocks: active-store zones surface first (so you don't scroll past a
    # wall of closed stores to find ones needing attention), all-closed zones
    # after. Within each bucket, order follows the fixed geographic sequence
    # EST -> CT -> MST -> PT -> Unknown.
    unique_tzs = list(df['Timezone'].unique()) if 'Timezone' in df.columns else ['EST']
    tz_order = ['EST', 'CT', 'MST', 'PT', 'Unknown Timezone']
    unique_tzs.sort(key=lambda x: tz_order.index(x) if x in tz_order else 99)

    active_blocks = []
    all_closed_blocks = []
    for tz in unique_tzs:
        df_tz = df[df['Timezone'] == tz]
        if all(str(status).strip().lower() in ['closed', 'force closed', 'deactivated', 'off-platform', 'menu empty'] for status in df_tz['Store_Status']):
            all_closed_blocks.append(tz)
        else:
            active_blocks.append(tz)

    final_block_order = active_blocks + all_closed_blocks
    
    header_rows_idx = []
    blank_rows_idx = []
    store_rows_metadata = [] 
    final_rows = []
    current_row_counter = 1
    store_serial = 0            # Sr. No - counts STORES only, ignoring banners

    IANA_TZ_MAP = {
        'EST': 'America/New_York', 'EDT': 'America/New_York',
        'CST': 'America/Chicago', 'CDT': 'America/Chicago',
        'MST': 'America/Denver', 'MDT': 'America/Denver',
        'PST': 'America/Los_Angeles', 'PDT': 'America/Los_Angeles',
        'CT': 'America/Chicago', 'PT': 'America/Los_Angeles'
    }

    for tz in final_block_order:
        # Deterministic order: status bucket first, then store name as a tiebreak.
        #
        # Without the tiebreak, rows inside one status bucket kept whatever order
        # the scrape happened to produce - and that order changes between runs,
        # because stores that failed on pass 1 are re-scraped in the recovery
        # passes and land at the END of the results. Two runs over identical data
        # could therefore emit different row orders, which is what makes the sheet
        # feel like it "fluctuates". mergesort is used because pandas' default
        # quicksort is not stable and would reshuffle equal keys on its own.
        df_tz = df[df['Timezone'] == tz].copy()
        df_tz['_name_key'] = df_tz['Store Name'].astype(str).str.strip().str.lower()
        df_tz = (df_tz.sort_values(by=['status_sort', '_name_key'], kind='mergesort')
                      .drop(columns='_name_key'))
        
        current_row_counter += 1
        header_rows_idx.append((current_row_counter, tz))
        
        if tz != "Unknown Timezone":
            iana_tz_name = IANA_TZ_MAP.get(tz, 'America/New_York')
            tz_local_time = datetime.now(timezone.utc).astimezone(ZoneInfo(iana_tz_name))
            formatted_tz_time = tz_local_time.strftime("%A, %d-%m-%Y %I:%M %p")
            header_text = f"--- {tz} TIME ZONE STORES --- (Scraped at: {formatted_tz_time} {tz})"
        else:
            header_text = "--- UNKNOWN TIMEZONE / NO ZIPCODE PROVIDED STORES ---"
        
        header_row = {col: None for col in original_columns}
        # Banner text goes in the FIRST column, whatever that column happens to
        # be. The merge anchors at column 1, and only the anchor cell's value is
        # rendered - writing the text into 'Timezone' silently loses it the
        # moment another column is inserted ahead of it.
        header_row[original_columns[0]] = header_text
        final_rows.append(header_row)

        for _, row in df_tz.iterrows():
            current_row_counter += 1
            status_val = str(row.get('Store_Status', 'Closed')).strip()
            store_rows_metadata.append((current_row_counter, status_val))
            store_row = {col: row.get(col, None) for col in original_columns}
            # Hours become one line per day here, before the value is written, so
            # the wrapping and the row height are both predictable.
            store_row['Timings'] = format_timings_for_display(store_row.get('Timings'))
            store_serial += 1
            store_row['Sr. No'] = store_serial
            final_rows.append(store_row)

        if tz != final_block_order[-1]:
            current_row_counter += 1
            blank_rows_idx.append(current_row_counter)
            sep_row = {col: None for col in original_columns}
            sep_row[original_columns[0]] = "--------------- "
            final_rows.append(sep_row)
            
    date_folder_suffix = datetime.now().strftime("%d-%m-%y")
    target_directory = f"ubereats_extracted {date_folder_suffix}"
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)
        
    file_timestamp = datetime.now().strftime("%d-%m-%Y_%I-%M-%S_%p")
    output_filename = f"ubereats_scraper_{file_timestamp}.xlsx"
    full_export_path = os.path.join(target_directory, output_filename)
    
    df_output = pd.DataFrame(final_rows, columns=original_columns)
    df_output.to_excel(full_export_path, index=False)
    
    wb = load_workbook(full_export_path)
    ws = wb.active
    
    # Styling definitions
    classic_gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    dark_header_fill = PatternFill(start_color="ECECEC", end_color="ECECEC", fill_type="solid")  # Grubhub header gray
    alert_red_fill = PatternFill(start_color="F7B6B6", end_color="F7B6B6", fill_type="solid")  # Grubhub red
    decommissioned_fill = PatternFill(start_color="D5DBDC", end_color="D5DBDC", fill_type="solid") 
    error_404_fill = PatternFill(start_color="D5DBDC", end_color="D5DBDC", fill_type="solid")  # slate = dead 
    rating_orange_fill = PatternFill(start_color="FBD3A6", end_color="FBD3A6", fill_type="solid")  # Grubhub orange 
    menu_empty_fill = PatternFill(start_color="BBD5EC", end_color="BBD5EC", fill_type="solid")  # palette blue
    flagged_fill = PatternFill(start_color="FBE7A3", end_color="FBE7A3", fill_type="solid")  # Grubhub gold
    # Force Closed gets its OWN hue (violet), not a second shade of red. Two
    # near-identical reds would defeat the purpose: the client has to be able to
    # tell "shut for the night" from "switched off" at a glance.
    force_closed_fill = PatternFill(start_color="D5BDE8", end_color="D5BDE8", fill_type="solid")
    # Phase 2 availability: a single cell (not the whole row) goes green/red so the
    # 2mile / 5mile result pops regardless of the row's own status colour.
    avail_green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    not_avail_fill   = PatternFill(start_color="F7B6B6", end_color="F7B6B6", fill_type="solid")

    font_white_bold = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    font_red_dark = Font(name="Segoe UI", size=11, bold=True, color="9C0006")
    font_decomm_bold = Font(name="Segoe UI", size=11, bold=True, color="7F0000", italic=True)
    font_error_bold = Font(name="Segoe UI", size=11, bold=True, color="9C6500", italic=True)
    font_normal = Font(name="Segoe UI", size=11, bold=False, color="000000")
    font_zone_header = Font(name="Segoe UI", size=12, bold=True, color="000000")
    font_orange_dark = Font(name="Segoe UI", size=11, bold=True, color="7F6000")
    font_menu_empty = Font(name="Segoe UI", size=11, bold=True, color="1F4E78", italic=True)
    font_flagged = Font(name="Segoe UI", size=11, bold=True, color="000000")
    font_dark_bold = Font(name="Segoe UI", size=11, bold=True, color="1A202C")
    font_force_closed = Font(name="Segoe UI", size=11, bold=True, color="4A148C")

    thin_border = Border(left=Side(style='thin', color='D9D9D9'),
                         right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'),
                         bottom=Side(style='thin', color='D9D9D9'))
    
    # Header row styling
    ws.row_dimensions[1].height = 26
    for col_idx in range(1, len(original_columns) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = dark_header_fill
        cell.font = Font(name="Segoe UI", size=11, bold=True, color="1A202C")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Zone headers: one consistent style across all zones. The point of a
    # header row is to read clearly as "this is a section divider" -- that
    # comes from it looking different from the data rows below it, not from
    # each zone having its own color.
    zone_header_fill = PatternFill(start_color="AEC5DE", end_color="AEC5DE", fill_type="solid")  # Grubhub navy
    zone_header_font = Font(name="Segoe UI", size=12, bold=True, color="1A202C")
    for h_row, h_tz in header_rows_idx:
        ws.row_dimensions[h_row].height = 24
        ws.merge_cells(start_row=h_row, start_column=1, end_row=h_row, end_column=len(original_columns))
        target_cell = ws.cell(row=h_row, column=1)
        target_cell.font = zone_header_font
        target_cell.fill = zone_header_fill
        # LEFT, not centred. The banner is merged across all 19 columns, so
        # centring puts its text in the middle of that span - far off the right
        # of the screen at normal scroll, and off the page when printed. Left
        # with a small indent keeps it where the eye already is.
        target_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col_idx in range(1, len(original_columns) + 1):
            ws.cell(row=h_row, column=col_idx).border = thin_border

    # ----------------------------------------------------------------------
    # Column widths: keep them TIGHT and CAPPED so the full header row fits on
    # screen with minimal horizontal scrolling. Long values (tags, address,
    # timings, links) WRAP inside the cell instead of stretching the column.
    # Merged banner / separator / legend rows live in column A but span every
    # column, so they must not be allowed to widen the Timezone column.
    # ----------------------------------------------------------------------
    width_caps = {
        # Each cap is at least as wide as its own header text, so no header wraps
        # mid-word ("Timezon / e", "Store_Statu / s").
        'Sr. No': 7, 'Timezone': 12, 'State': 14, 'Store Name': 30, 'Store_Status': 16,
        'Status_Note': 28,
        'Rating': 8, 'Review Count': 10,
        'Tag 1': 13, 'Tag 2': 13, 'Tag 3': 13, 'Tag 4': 13,
        'Tag 5': 13, 'Tag 6': 13, 'Tag 7': 13,
        'All Tags': 24, 'Address': 26, 'Mobile': 18,
        'Timings': 30, 'Website_link': 16,
        'U1 Icon': 16, '2mile': 14, '5mile': 14,
    }
    merged_row_set = {h_row for h_row, _ in header_rows_idx} | set(blank_rows_idx)

    for col_idx, col_name in enumerate(original_columns, start=1):
        cap = width_caps.get(col_name, 12)

        # Timezone: fixed to fit the word on one line. Its long merged banner text
        # and the wrappable "Unknown Timezone" value are intentionally ignored.
        if col_name == "Timezone":
            ws.column_dimensions[get_column_letter(col_idx)].width = cap
            continue

        # Measure real content length, skipping merged banner/separator/legend rows.
        # Long wrapping values are measured by their longest LINE, not their total
        # length - a 4-line address should not demand a 60-wide column.
        max_length = 0
        for r in range(1, ws.max_row + 1):
            if r in merged_row_set:
                continue
            val = ws.cell(row=r, column=col_idx).value
            if val:
                longest_line = max(len(part) for part in str(val).split('\n'))
                max_length = max(max_length, longest_line)

        # The HEADER sets the floor. Measuring content alone is what made
        # "Review Count" 10 wide and "Store_Status" 14 - both then wrapped their
        # own header mid-word ("Review / Count", "Store_Statu / s"), which is what
        # made the top of the sheet look broken. Header text is bold, so it needs
        # roughly 15% more room than its character count suggests.
        header_floor = ceil(len(str(col_name)) * 1.15) + 2

        # A non-wrapping column must fit its content outright - it has no second
        # line to fall back on, so anything too narrow silently clips (a 12-digit
        # phone number losing its first digits). Still bounded by the cap:
        # Website_link holds a 599-character URL at this point (it is shortened to
        # a hyperlink later), and without the cap it demanded a 601-wide column.
        if col_name in NO_WRAP_COLS:
            floor = max(header_floor, min(cap, max_length + 2))
        else:
            floor = header_floor

        # Content gets the same bold allowance as the header: data rows for
        # Closed / Flagged / Force Closed / Deactivated stores are all bold, and
        # measuring them at regular width is what made short tag values wrap
        # mid-word ("Flower / s", "Grocer / y").
        content_need = ceil(max_length * 1.12) + 2
        width = max(floor, min(cap, max(6, content_need)))
        ws.column_dimensions[get_column_letter(col_idx)].width = width



    center_columns = ['Sr. No', 'Timezone', 'State', 'Store_Status', 'Rating', 'Review Count', 'Mobile',
                'Tag 1', 'Tag 2', 'Tag 3', 'Tag 4', 'Tag 5', 'Tag 6', 'Tag 7',
                'U1 Icon', '2mile', '5mile']


    center_columns_idx = [original_columns.index(c) + 1 for c in center_columns]

    
    for r_num, store_status in store_rows_metadata:
        notice = str(store_status).strip()

        if notice == "Flagged":
            for col_idx in range(1, len(original_columns) + 1):
                cell = ws.cell(row=r_num, column=col_idx)
                cell.border = thin_border
                cell.fill = flagged_fill
                cell.font = font_flagged
                cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
                                        vertical="center", wrap_text=True)
            ws.row_dimensions[r_num].height = 28

        elif notice == "Deactivated":
            for col_idx in range(1, len(original_columns) + 1):
                cell = ws.cell(row=r_num, column=col_idx)
                cell.border = thin_border
                cell.fill = error_404_fill
                cell.font = font_dark_bold
                cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
                                        vertical="center", wrap_text=True)
            ws.row_dimensions[r_num].height = 28

        elif notice == "Off-Platform":
            # Column numbers derived from the header list, not hardcoded - adding a
            # column (Status_Note) must not silently bold the wrong cells.
            _emph_cols = [original_columns.index(c) + 1 for c in ('Store_Status', 'Mobile')
                          if c in original_columns]
            for col_idx in range(1, len(original_columns) + 1):
                cell = ws.cell(row=r_num, column=col_idx)
                cell.border = thin_border
                cell.fill = decommissioned_fill
                cell.font = font_decomm_bold if col_idx in _emph_cols else font_normal
                cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
                                        vertical="center", wrap_text=True)
            ws.row_dimensions[r_num].height = 30

        elif notice == "Force Closed":
            for col_idx in range(1, len(original_columns) + 1):
                cell = ws.cell(row=r_num, column=col_idx)
                cell.border = thin_border
                cell.fill = force_closed_fill
                cell.font = font_force_closed
                cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
                                        vertical="center", wrap_text=True)
            ws.row_dimensions[r_num].height = 28

        elif notice == "Closed":
            for col_idx in range(1, len(original_columns) + 1):
                cell = ws.cell(row=r_num, column=col_idx)
                cell.border = thin_border
                cell.fill = alert_red_fill
                cell.font = font_dark_bold
                cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
                                        vertical="center", wrap_text=True)

        elif notice == "Menu Empty":
            for col_idx in range(1, len(original_columns) + 1):
                cell = ws.cell(row=r_num, column=col_idx)
                cell.border = thin_border
                cell.fill = menu_empty_fill
                cell.font = font_dark_bold
                cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
                                        vertical="center", wrap_text=True)
            ws.row_dimensions[r_num].height = 28

        elif notice in ("Active", "Open"):
            for col_idx in range(1, len(original_columns) + 1):
                cell = ws.cell(row=r_num, column=col_idx)
                cell.border = thin_border
                cell.font = font_normal
                cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
                                        vertical="center", wrap_text=True)

        else:  # Unknown
            for col_idx in range(1, len(original_columns) + 1):
                cell = ws.cell(row=r_num, column=col_idx)
                cell.border = thin_border
                cell.font = font_normal
                cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
                                        vertical="center", wrap_text=True)

        # --- Rating flagging (< 3.5) ---
        rating_col_idx = original_columns.index("Rating") + 1
        rating_val = ws.cell(row=r_num, column=rating_col_idx).value
        try:
            if rating_val and float(rating_val) < 3.5:
                cell = ws.cell(row=r_num, column=rating_col_idx)
                cell.fill = rating_orange_fill
                cell.font = font_dark_bold
        except ValueError:
            pass

        # Row heights are set by the uniform presentation pass below, which knows
        # which columns wrap and caps the result. The old estimator counted the
        # full 599-character store URL against a 22-wide column and produced
        # 500-point rows.

    # ======================================================================
    # UNIFORM PRESENTATION PASS
    #
    # Runs LAST over the whole data range, so alignment is decided in exactly one
    # place instead of being re-specified inside each status branch. Every cell
    # in every data row gets the same treatment - including empty cells, which is
    # what stops a column looking ragged wherever a value happens to be blank.
    #
    # Fills and fonts set by the status branches above are left alone; only
    # alignment, wrapping, borders and row height are normalised here.
    # ======================================================================
    merged_rows = {h_row for h_row, _ in header_rows_idx} | set(blank_rows_idx)
    last_data_row = max((r for r, _ in store_rows_metadata), default=1)

    link_font = Font(name="Segoe UI", size=11, color="1F4E79", underline="single")

    for r_num in range(2, last_data_row + 1):
        if r_num in merged_rows:
            continue

        max_lines = 1
        for col_idx, col_name in enumerate(original_columns, start=1):
            cell = ws.cell(row=r_num, column=col_idx)
            wraps = col_name not in NO_WRAP_COLS
            width = ws.column_dimensions[get_column_letter(col_idx)].width or 12

            # --- Website_link: one line, always -------------------------------
            # The visible text drops the query string (which is what made these
            # 599 characters long); the full original URL stays reachable as a
            # hyperlink, so nothing is lost.
            if col_name == 'Website_link' and cell.value:
                full_url = str(cell.value)
                if full_url.lower().startswith("http"):
                    cell.value = "Open store"
                    cell.hyperlink = full_url
                    cell.font = link_font

            # --- numbers stored as numbers, so they sort and filter properly ---
            # --- phone stays TEXT ---------------------------------------------
            # "+19147719453" in a General cell is read as a formula by Excel and
            # LibreOffice, which turns a phone number into a negative integer.
            # An explicit text format pins it.
            elif col_name == 'Mobile' and cell.value not in (None, ""):
                cell.number_format = '@'

            elif col_name in ('Rating', 'Review Count') and cell.value not in (None, ""):
                raw = str(cell.value).replace(",", "").replace("+", "").strip()
                try:
                    cell.value = float(raw) if col_name == 'Rating' else int(float(raw))
                    cell.number_format = '0.0' if col_name == 'Rating' else '#,##0'
                except ValueError:
                    pass

            cell.alignment = Alignment(
                horizontal=COL_ALIGN.get(col_name, 'left'),
                vertical='center',
                wrap_text=wraps,
                indent=1 if COL_ALIGN.get(col_name) == 'left' else 0,
            )
            cell.border = thin_border

            # --- Phase 2 availability cell colour (green / red) ---------------
            if col_name in ('2mile', '5mile'):
                v = str(cell.value or '').strip().lower()
                if v == 'available':
                    cell.fill = avail_green_fill
                    cell.font = font_dark_bold
                elif v == 'not available':
                    cell.fill = not_avail_fill
                    cell.font = font_dark_bold

            lines = _wrapped_line_count(cell.value, width, wraps)
            if lines > max_lines:
                max_lines = lines

        max_lines = min(max_lines, MAX_ROW_LINES)
        ws.row_dimensions[r_num].height = max_lines * LINE_HEIGHT_PTS + ROW_PADDING_PTS

    # Blank row styling
    for b_row in blank_rows_idx:
        ws.row_dimensions[b_row].height = 15
        ws.merge_cells(start_row=b_row, start_column=1, end_row=b_row, end_column=len(original_columns))
        b_cell = ws.cell(row=b_row, column=1)
        b_cell.font = Font(name="Segoe UI", size=9, italic=True, color="7F7F7F")
        b_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws.freeze_panes = "A2"

    # ------------------------------------------------------------------
    # Print / PDF setup. Without this, 19 columns split across several pages in
    # an order nobody can follow, and only page 1 carries the header row.
    # ------------------------------------------------------------------
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0           # as many pages tall as needed
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"             # header repeats on every page
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4


    # ----------------------------------------------------------------------
    # Make the header visible without horizontal scrolling on open:
    #  - open at the top-left cell (A1 selected),
    #  - set a zoom level that fits the full row of columns on a standard screen.
    # The zoom is derived from the total column width: if everything already fits
    # in a typical laptop-width view (~180 units at 100%), keep 100%; otherwise
    # scale down just enough (never below 70% so text stays readable).
    # ----------------------------------------------------------------------
    total_width = sum(
        (ws.column_dimensions[get_column_letter(c)].width or 8)
        for c in range(1, len(original_columns) + 1)
    )
    TARGET_VISIBLE_UNITS = 210  # approx. one standard laptop screen at 100% zoom
    if total_width > TARGET_VISIBLE_UNITS:
        # Floor of 75%: below that the text stops being comfortably readable, and
        # a little horizontal scrolling beats squinting.
        fit_zoom = max(75, int(TARGET_VISIBLE_UNITS / total_width * 100))
    else:
        fit_zoom = 100
    ws.sheet_view.zoomScale = fit_zoom
    ws.sheet_view.topLeftCell = "A1"
    if ws.sheet_view.selection:
        ws.sheet_view.selection[0].activeCell = "A1"
        ws.sheet_view.selection[0].sqref = "A1"

    # Legend block at bottom
    legend_start = ws.max_row + 2
    ws.merge_cells(start_row=legend_start, start_column=1, end_row=legend_start, end_column=len(original_columns))
    legend_cell = ws.cell(row=legend_start, column=1)
    legend_cell.value = "LEGEND / COLOR KEY"
    legend_cell.font = Font(name="Segoe UI Semibold", size=12, bold=True, color="1A202C")
    legend_cell.alignment = Alignment(horizontal="left", vertical="center")

    legend_entries = [
        ("AEC5DE", "Time Zone Section Header"),
        (None,     "Open / Active store  (plain white row)"),
        ("FBE7A3", "FLAGGED  —  page shows Closed but posted hours say it should be OPEN now"),
        ("D5BDE8", "FORCE CLOSED  —  store switched itself off (\"Delivery unavailable\" / not accepting orders). See Status_Note"),
        ("F7B6B6", "Closed  —  store is outside its operating hours right now"),
        ("D5DBDC", "Deactivated  —  404 page or removed from Uber Eats"),
        ("BBD5EC", "Menu Empty  —  page loads but no items are listed"),
        ("FBD3A6", "Low Rating cell  —  store rating is below 3.5"),
        ("C6EFCE", "Available  —  (Uber One) store delivers to the 2-mile / 5-mile address"),
        ("F7B6B6", "Not Available  —  (Uber One) store does not deliver to that address"),
    ]

    for offset, (swatch_hex, meaning) in enumerate(legend_entries, start=1):
        r = legend_start + offset
        swatch = ws.cell(row=r, column=1)
        swatch.value = ""
        swatch.fill = PatternFill(start_color=(swatch_hex or "FFFFFF"),
                                  end_color=(swatch_hex or "FFFFFF"), fill_type="solid")
        swatch.border = thin_border
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=len(original_columns))
        desc = ws.cell(row=r, column=2)
        desc.value = meaning
        desc.font = Font(name="Segoe UI", size=11, bold=False, color="1A202C")
        desc.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r].height = 20

    wb.save(full_export_path)
    return full_export_path

def deliver_report_via_email(attachment_path, target_recipient, store_count=None):
    if not target_recipient or pd.isna(target_recipient) or str(target_recipient).strip() == "":
        logging.error("No valid recipient email provided.")
        return

    recipient_clean = str(target_recipient).strip()

    # The scraper sends FROM the developer's own SMTP account (SENDER_EMAIL in
    # config), so a copy of this email lands in that account's Sent folder on
    # every run - a tamper-proof usage log the client never sees and cannot
    # delete. The subject below is written so that Sent folder is scannable at a
    # glance: WHO the report went to, from WHICH machine, HOW MANY stores, and
    # WHEN - no need to open the mail or the attachment to know what a run did.
    import socket
    from datetime import datetime as _dt
    try:
        machine = socket.gethostname()
    except Exception:
        machine = "unknown-pc"
    stamp = _dt.now().strftime("%d-%m-%Y %H:%M")
    count_part = f"{store_count} stores | " if store_count is not None else ""

    recipient_clean = str(target_recipient).strip()

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = recipient_clean
    filename_raw = os.path.basename(attachment_path)
    # e.g.  [Scraper] JK-RECEPTION-PC | 212 stores | to client@theirs.com | 12-09-2026 19:25
    msg['Subject'] = (f"[Scraper] {machine} | {count_part}"
                      f"to {recipient_clean} | {stamp}")
    body = (
        "Automated Uber Eats scraper run.\n\n"
        f"Machine   : {machine}\n"
        f"Recipient : {recipient_clean}\n"
        + (f"Stores    : {store_count}\n" if store_count is not None else "")
        + f"Time      : {stamp}\n"
        f"File      : {filename_raw}\n"
    )
    msg.attach(MIMEText(body, 'plain'))

    try:
        if os.path.exists(attachment_path):
            with open(attachment_path, "rb") as file_obj:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file_obj.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{filename_raw}"')
                msg.attach(part)

        # Use config port instead of hardcoding
        smtp_session = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        smtp_session.set_debuglevel(0)
        smtp_session.login(SENDER_EMAIL, SENDER_PASSWORD)
        smtp_session.sendmail(SENDER_EMAIL, recipient_clean, msg.as_string())
        smtp_session.quit()

        logging.info(f"Report delivered to {recipient_clean}")
    except Exception as error:
        logging.error(f"Email failure: {error}")






































# #excel_formatter.py #working till store paused
# import os
# import pandas as pd
# import logging
# from datetime import datetime, timezone
# from zoneinfo import ZoneInfo
# from openpyxl import load_workbook
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter
# import smtplib
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# from email.mime.base import MIMEBase
# from email import encoders
# from email.header import Header
# from config import SMTP_SERVER,SMTP_PORT,SENDER_EMAIL ,SENDER_PASSWORD
# from timezone_utils import extract_store_id
# import re
# from math import ceil


# # ==========================================================================
# # PRESENTATION RULES - one place that decides how every column is displayed
# # ==========================================================================
# # Horizontal alignment per column. Short codes and numbers centre; anything
# # read as a sentence goes left. Applied to EVERY cell in the used range, empty
# # cells included, so a column never looks ragged where a value happens to be
# # blank.
# COL_ALIGN = {
#     'Sr. No': 'center', 'Timezone': 'center', 'State': 'left', 'Store Name': 'left',
#     'Store_Status': 'center', 'Status_Note': 'left',
#     'Rating': 'center', 'Review Count': 'center',
#     'Tag 1': 'center', 'Tag 2': 'center', 'Tag 3': 'center', 'Tag 4': 'center',
#     'Tag 5': 'center', 'Tag 6': 'center', 'Tag 7': 'center',
#     'All Tags': 'left', 'Address': 'left', 'Mobile': 'center',
#     'Timings': 'left', 'Website_link': 'center',
# }

# # Columns that must NEVER wrap. A 599-character store URL wrapped into a 22-wide
# # column becomes 32 lines and a 500-point row - which is exactly what made the
# # sheet unreadable. The link is shown as one line and carries the real URL as a
# # clickable hyperlink instead.
# NO_WRAP_COLS = {'Sr. No', 'Website_link', 'Mobile', 'Rating', 'Review Count'}

# # Hard ceiling on how tall a data row may grow, in wrapped lines. Without a cap,
# # one long field dictates the height of the whole row.
# MAX_ROW_LINES = 8
# LINE_HEIGHT_PTS = 15
# ROW_PADDING_PTS = 8

# _TIME_RE = re.compile(r'\d{1,2}:\d{2}\s?(?:AM|PM)', re.IGNORECASE)
# _DAY_RE = re.compile(
#     r'\b(mon|tue|wed|thu|fri|sat|sun)[a-z]*\b|every\s*day|daily', re.IGNORECASE)


# def format_timings_for_display(raw):
#     """Turn the scraped hours string into one readable line per day.

#     In:   'Sunday ; Closed ; Monday - Saturday ; 10:00 AM - 4:00 PM'
#     Out:  'Sunday: Closed
#            Monday - Saturday: 10:00 AM - 4:00 PM'

#     The semicolon-run is technically complete but reads as a wall of text, and
#     wraps at arbitrary points mid-time. One day per line wraps predictably and
#     can be scanned down the column.
#     """
#     if not raw or not str(raw).strip():
#         return raw
#     text = str(raw)
#     # Keep any banner text (e.g. "Closed on Uber Eats as of ...") untouched.
#     if not _DAY_RE.search(text):
#         return text

#     tokens = [t.strip(" \t•") for t in re.split(r'[;\n]', text) if t.strip(" \t•")]
#     lines, pending_day = [], None

#     for tok in tokens:
#         has_time = bool(_TIME_RE.search(tok))
#         has_day = bool(_DAY_RE.search(tok))
#         is_closed = 'closed' in tok.lower()

#         if has_day and has_time:
#             m = _TIME_RE.search(tok)
#             day_part = tok[:m.start()].strip(" -:–")
#             time_part = tok[m.start():].strip()
#             lines.append(f"{day_part}: {time_part}" if day_part else time_part)
#             pending_day = None
#         elif has_day and is_closed:
#             lines.append(f"{re.sub(r'(?i)closed', '', tok).strip(' -:')}: Closed")
#             pending_day = None
#         elif has_day:
#             pending_day = tok
#         elif pending_day and (has_time or is_closed):
#             lines.append(f"{pending_day}: {tok}")
#             pending_day = None
#         else:
#             lines.append(tok)

#     if pending_day:
#         lines.append(pending_day)
#     return "\n".join(lines)


# def _wrapped_line_count(value, width, wraps):
#     """How many display lines a cell needs at this column width."""
#     if value is None or str(value) == "":
#         return 1
#     if not wraps:
#         return 1
#     usable = max(4, int(width) - 2)
#     total = 0
#     for para in str(value).split('\n'):
#         total += 1 if not para else ceil(len(para) / usable)
#     return max(1, total)



# # ==========================================
# # 4. EXCEL FORMATTING & WORKBOOK STYLING ENGINE
# # ==========================================
# def format_and_export_excel(records):
#     df = pd.DataFrame(records)
    
#     # Count before dropping
#     before_count = len(df)

#     # Pass 0: same store ID, pulled from the URL - catches exact repeat
#     # scrapes of the identical Uber Eats listing regardless of how the
#     # scraped Store Name/Address text happened to format on each pass.
#     # Belt-and-suspenders alongside main.py's own pre-scrape input.xlsx dedup.
#     if 'Website_link' in df.columns:
#         df['_store_id'] = df['Website_link'].apply(extract_store_id)
#         has_store_id = df['_store_id'].ne('')
#         removed0 = df[has_store_id & df.duplicated(subset='_store_id', keep='first')]
#         if not removed0.empty:
#             logging.warning(
#                 f"Pass 0 (same store ID): dropping {len(removed0)} duplicate row(s): "
#                 f"{sorted(removed0['Store Name'].astype(str).unique())}"
#             )
#         df = pd.concat([
#             df[~has_store_id],
#             df[has_store_id].drop_duplicates(subset='_store_id', keep='first')
#         ], ignore_index=True).drop(columns='_store_id')

#     # Pass 1: exact repeat scrapes of the same store (identical name + address).
#     removed1 = df[df.duplicated(subset=['Store Name', 'Address'], keep='first')]
#     if not removed1.empty:
#         logging.warning(
#             f"Pass 1 (same name + address): dropping {len(removed1)} duplicate row(s): "
#             f"{sorted(removed1['Store Name'].astype(str).unique())}"
#         )
#     df = df.drop_duplicates(subset=['Store Name', 'Address'], keep='first')

#     # Pass 2: same physical store listed under a different name/URL in the
#     # input sheet -- Store Name differs but Address + Mobile are identical,
#     # so Pass 1 alone misses these. Only applied to rows that actually HAVE
#     # an address/phone (Deactivated / Menu Empty / etc. rows are blank on
#     # both, and blank == blank would wrongly collapse all of those together).
#     if 'Address' in df.columns and 'Mobile' in df.columns:
#         has_contact_info = (
#             df['Address'].astype(str).str.strip().ne('') &
#             df['Mobile'].astype(str).str.strip().ne('')
#         )
#         df_contact = df[has_contact_info]
#         removed2 = df_contact[df_contact.duplicated(subset=['Address', 'Mobile'], keep='first')]
#         if not removed2.empty:
#             logging.warning(
#                 f"Pass 2 (same address + phone): dropping {len(removed2)} duplicate row(s): "
#                 f"{sorted(removed2['Store Name'].astype(str).unique())}"
#             )
#         df_contact = df_contact.drop_duplicates(subset=['Address', 'Mobile'], keep='first')
#         df_no_contact = df[~has_contact_info]
#         df = pd.concat([df_contact, df_no_contact], ignore_index=True)

#     # Pass 3 (SOFT, log-only - does NOT remove rows): same address on file,
#     # but Pass 2 couldn't confirm it with a matching phone number (blank
#     # Mobile, or the numbers differ). This closes the blind spot where a
#     # store with no listed phone slides through Pass 2 untouched - but it's
#     # only logged, never auto-merged, because a shared street address (food
#     # court, strip mall, multi-unit plaza) can legitimately hold several
#     # distinct restaurants. Auto-removing here risks deleting a real store.
#     # Review these manually and re-run with a corrected input.xlsx if any
#     # turn out to be genuine duplicates.
#     if 'Address' in df.columns:
#         addr_nonblank = df[df['Address'].astype(str).str.strip().ne('')]
#         soft_dupes = addr_nonblank[addr_nonblank.duplicated(subset='Address', keep=False)]
#         if not soft_dupes.empty:
#             pairs = soft_dupes[['Store Name', 'Address']].drop_duplicates().to_dict('records')
#             logging.warning(
#                 "POSSIBLE duplicates (same address, unconfirmed by phone - review "
#                 f"manually, NOT auto-removed): {pairs}"
#             )

#     # Count after dropping
#     after_count = len(df)
#     removed_count = before_count - after_count

#     if removed_count > 0:
#         logging.warning(f"Removed {removed_count} duplicate row(s) total before export.")

#     # Convert back to list of dicts for your existing Excel writing logic
#     master_records = df.to_dict(orient='records')
    
#     def get_status_sort_val(status_val):
#         s = str(status_val).strip().lower()
#         if s == 'flagged': return -1     # mismatches float to the very top
#         if '404' in s or s in ('deactivated', 'off-platform'): return 0
#         # Force Closed = store switched itself off ("Delivery unavailable" /
#         # "not accepting orders"). Actionable, so it sits ABOVE the ordinary
#         # out-of-hours closures.
#         if s == 'force closed': return 1
#         if s == 'closed': return 2       # Closed stores near the top (below deactivated)
#         if s in ('active', 'open'): return 3  # Open / Active stores below closed
#         if s == 'menu empty': return 4   # Menu empty at the bottom
#         return 5  # Unknown/other statuses at the very bottom
        
#     df['status_sort'] = df['Store_Status'].apply(get_status_sort_val)
    
#     # 'Status_Note' carries the store's OWN status line, exactly as the page
#     # printed it ("Delivery unavailable", "Closed - Available at 11:00 AM"). It is
#     # what makes a Force Closed row auditable: the client can see the store's own
#     # words next to the verdict instead of taking the label on trust.
#     # 'Sr. No' numbers the stores CONTINUOUSLY down the whole sheet, across zone
#     # blocks - not restarting per zone - so "store 87" identifies one row in the
#     # file with no further qualification. It is worth having precisely because
#     # Excel's own row numbers cannot do that job here: the banner and separator
#     # rows consume row numbers, so the 87th store is never on row 87.
#     original_columns = [
#     'Sr. No',
#     'Timezone', 'State', 'Store Name', 'Store_Status', 'Status_Note', 'Rating', 'Review Count',
#     'Tag 1', 'Tag 2', 'Tag 3', 'Tag 4', 'Tag 5', 'Tag 6', 'Tag 7',
#     'All Tags', 'Address', 'Mobile', 'Timings', 'Website_link'
# ]

    
#     # Zone blocks: active-store zones surface first (so you don't scroll past a
#     # wall of closed stores to find ones needing attention), all-closed zones
#     # after. Within each bucket, order follows the fixed geographic sequence
#     # EST -> CT -> MST -> PT -> Unknown.
#     unique_tzs = list(df['Timezone'].unique()) if 'Timezone' in df.columns else ['EST']
#     tz_order = ['EST', 'CT', 'MST', 'PT', 'Unknown Timezone']
#     unique_tzs.sort(key=lambda x: tz_order.index(x) if x in tz_order else 99)

#     active_blocks = []
#     all_closed_blocks = []
#     for tz in unique_tzs:
#         df_tz = df[df['Timezone'] == tz]
#         if all(str(status).strip().lower() in ['closed', 'force closed', 'deactivated', 'off-platform', 'menu empty'] for status in df_tz['Store_Status']):
#             all_closed_blocks.append(tz)
#         else:
#             active_blocks.append(tz)

#     final_block_order = active_blocks + all_closed_blocks
    
#     header_rows_idx = []
#     blank_rows_idx = []
#     store_rows_metadata = [] 
#     final_rows = []
#     current_row_counter = 1
#     store_serial = 0            # Sr. No - counts STORES only, ignoring banners

#     IANA_TZ_MAP = {
#         'EST': 'America/New_York', 'EDT': 'America/New_York',
#         'CST': 'America/Chicago', 'CDT': 'America/Chicago',
#         'MST': 'America/Denver', 'MDT': 'America/Denver',
#         'PST': 'America/Los_Angeles', 'PDT': 'America/Los_Angeles',
#         'CT': 'America/Chicago', 'PT': 'America/Los_Angeles'
#     }

#     for tz in final_block_order:
#         # Deterministic order: status bucket first, then store name as a tiebreak.
#         #
#         # Without the tiebreak, rows inside one status bucket kept whatever order
#         # the scrape happened to produce - and that order changes between runs,
#         # because stores that failed on pass 1 are re-scraped in the recovery
#         # passes and land at the END of the results. Two runs over identical data
#         # could therefore emit different row orders, which is what makes the sheet
#         # feel like it "fluctuates". mergesort is used because pandas' default
#         # quicksort is not stable and would reshuffle equal keys on its own.
#         df_tz = df[df['Timezone'] == tz].copy()
#         df_tz['_name_key'] = df_tz['Store Name'].astype(str).str.strip().str.lower()
#         df_tz = (df_tz.sort_values(by=['status_sort', '_name_key'], kind='mergesort')
#                       .drop(columns='_name_key'))
        
#         current_row_counter += 1
#         header_rows_idx.append((current_row_counter, tz))
        
#         if tz != "Unknown Timezone":
#             iana_tz_name = IANA_TZ_MAP.get(tz, 'America/New_York')
#             tz_local_time = datetime.now(timezone.utc).astimezone(ZoneInfo(iana_tz_name))
#             formatted_tz_time = tz_local_time.strftime("%A, %d-%m-%Y %I:%M %p")
#             header_text = f"--- {tz} TIME ZONE STORES --- (Scraped at: {formatted_tz_time} {tz})"
#         else:
#             header_text = "--- UNKNOWN TIMEZONE / NO ZIPCODE PROVIDED STORES ---"
        
#         header_row = {col: None for col in original_columns}
#         # Banner text goes in the FIRST column, whatever that column happens to
#         # be. The merge anchors at column 1, and only the anchor cell's value is
#         # rendered - writing the text into 'Timezone' silently loses it the
#         # moment another column is inserted ahead of it.
#         header_row[original_columns[0]] = header_text
#         final_rows.append(header_row)

#         for _, row in df_tz.iterrows():
#             current_row_counter += 1
#             status_val = str(row.get('Store_Status', 'Closed')).strip()
#             store_rows_metadata.append((current_row_counter, status_val))
#             store_row = {col: row.get(col, None) for col in original_columns}
#             # Hours become one line per day here, before the value is written, so
#             # the wrapping and the row height are both predictable.
#             store_row['Timings'] = format_timings_for_display(store_row.get('Timings'))
#             store_serial += 1
#             store_row['Sr. No'] = store_serial
#             final_rows.append(store_row)

#         if tz != final_block_order[-1]:
#             current_row_counter += 1
#             blank_rows_idx.append(current_row_counter)
#             sep_row = {col: None for col in original_columns}
#             sep_row[original_columns[0]] = "--------------- "
#             final_rows.append(sep_row)
            
#     date_folder_suffix = datetime.now().strftime("%d-%m-%y")
#     target_directory = f"ubereats_extracted {date_folder_suffix}"
#     if not os.path.exists(target_directory):
#         os.makedirs(target_directory)
        
#     file_timestamp = datetime.now().strftime("%d-%m-%Y_%I-%M-%S_%p")
#     output_filename = f"ubereats_scraper_{file_timestamp}.xlsx"
#     full_export_path = os.path.join(target_directory, output_filename)
    
#     df_output = pd.DataFrame(final_rows, columns=original_columns)
#     df_output.to_excel(full_export_path, index=False)
    
#     wb = load_workbook(full_export_path)
#     ws = wb.active
    
#     # Styling definitions
#     classic_gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
#     dark_header_fill = PatternFill(start_color="ECECEC", end_color="ECECEC", fill_type="solid")  # Grubhub header gray
#     alert_red_fill = PatternFill(start_color="F7B6B6", end_color="F7B6B6", fill_type="solid")  # Grubhub red
#     decommissioned_fill = PatternFill(start_color="D5DBDC", end_color="D5DBDC", fill_type="solid") 
#     error_404_fill = PatternFill(start_color="D5DBDC", end_color="D5DBDC", fill_type="solid")  # slate = dead 
#     rating_orange_fill = PatternFill(start_color="FBD3A6", end_color="FBD3A6", fill_type="solid")  # Grubhub orange 
#     menu_empty_fill = PatternFill(start_color="BBD5EC", end_color="BBD5EC", fill_type="solid")  # palette blue
#     flagged_fill = PatternFill(start_color="FBE7A3", end_color="FBE7A3", fill_type="solid")  # Grubhub gold
#     # Force Closed gets its OWN hue (violet), not a second shade of red. Two
#     # near-identical reds would defeat the purpose: the client has to be able to
#     # tell "shut for the night" from "switched off" at a glance.
#     force_closed_fill = PatternFill(start_color="D5BDE8", end_color="D5BDE8", fill_type="solid")

#     font_white_bold = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
#     font_red_dark = Font(name="Segoe UI", size=11, bold=True, color="9C0006")
#     font_decomm_bold = Font(name="Segoe UI", size=11, bold=True, color="7F0000", italic=True)
#     font_error_bold = Font(name="Segoe UI", size=11, bold=True, color="9C6500", italic=True)
#     font_normal = Font(name="Segoe UI", size=11, bold=False, color="000000")
#     font_zone_header = Font(name="Segoe UI", size=12, bold=True, color="000000")
#     font_orange_dark = Font(name="Segoe UI", size=11, bold=True, color="7F6000")
#     font_menu_empty = Font(name="Segoe UI", size=11, bold=True, color="1F4E78", italic=True)
#     font_flagged = Font(name="Segoe UI", size=11, bold=True, color="000000")
#     font_dark_bold = Font(name="Segoe UI", size=11, bold=True, color="1A202C")
#     font_force_closed = Font(name="Segoe UI", size=11, bold=True, color="4A148C")

#     thin_border = Border(left=Side(style='thin', color='D9D9D9'),
#                          right=Side(style='thin', color='D9D9D9'),
#                          top=Side(style='thin', color='D9D9D9'),
#                          bottom=Side(style='thin', color='D9D9D9'))
    
#     # Header row styling
#     ws.row_dimensions[1].height = 26
#     for col_idx in range(1, len(original_columns) + 1):
#         cell = ws.cell(row=1, column=col_idx)
#         cell.fill = dark_header_fill
#         cell.font = Font(name="Segoe UI", size=11, bold=True, color="1A202C")
#         cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#         cell.border = thin_border

#     # Zone headers: one consistent style across all zones. The point of a
#     # header row is to read clearly as "this is a section divider" -- that
#     # comes from it looking different from the data rows below it, not from
#     # each zone having its own color.
#     zone_header_fill = PatternFill(start_color="AEC5DE", end_color="AEC5DE", fill_type="solid")  # Grubhub navy
#     zone_header_font = Font(name="Segoe UI", size=12, bold=True, color="1A202C")
#     for h_row, h_tz in header_rows_idx:
#         ws.row_dimensions[h_row].height = 24
#         ws.merge_cells(start_row=h_row, start_column=1, end_row=h_row, end_column=len(original_columns))
#         target_cell = ws.cell(row=h_row, column=1)
#         target_cell.font = zone_header_font
#         target_cell.fill = zone_header_fill
#         # LEFT, not centred. The banner is merged across all 19 columns, so
#         # centring puts its text in the middle of that span - far off the right
#         # of the screen at normal scroll, and off the page when printed. Left
#         # with a small indent keeps it where the eye already is.
#         target_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
#         for col_idx in range(1, len(original_columns) + 1):
#             ws.cell(row=h_row, column=col_idx).border = thin_border

#     # ----------------------------------------------------------------------
#     # Column widths: keep them TIGHT and CAPPED so the full header row fits on
#     # screen with minimal horizontal scrolling. Long values (tags, address,
#     # timings, links) WRAP inside the cell instead of stretching the column.
#     # Merged banner / separator / legend rows live in column A but span every
#     # column, so they must not be allowed to widen the Timezone column.
#     # ----------------------------------------------------------------------
#     width_caps = {
#         # Each cap is at least as wide as its own header text, so no header wraps
#         # mid-word ("Timezon / e", "Store_Statu / s").
#         'Sr. No': 7, 'Timezone': 12, 'State': 14, 'Store Name': 30, 'Store_Status': 16,
#         'Status_Note': 28,
#         'Rating': 8, 'Review Count': 10,
#         'Tag 1': 13, 'Tag 2': 13, 'Tag 3': 13, 'Tag 4': 13,
#         'Tag 5': 13, 'Tag 6': 13, 'Tag 7': 13,
#         'All Tags': 24, 'Address': 26, 'Mobile': 18,
#         'Timings': 30, 'Website_link': 16,
#     }
#     merged_row_set = {h_row for h_row, _ in header_rows_idx} | set(blank_rows_idx)

#     for col_idx, col_name in enumerate(original_columns, start=1):
#         cap = width_caps.get(col_name, 12)

#         # Timezone: fixed to fit the word on one line. Its long merged banner text
#         # and the wrappable "Unknown Timezone" value are intentionally ignored.
#         if col_name == "Timezone":
#             ws.column_dimensions[get_column_letter(col_idx)].width = cap
#             continue

#         # Measure real content length, skipping merged banner/separator/legend rows.
#         # Long wrapping values are measured by their longest LINE, not their total
#         # length - a 4-line address should not demand a 60-wide column.
#         max_length = 0
#         for r in range(1, ws.max_row + 1):
#             if r in merged_row_set:
#                 continue
#             val = ws.cell(row=r, column=col_idx).value
#             if val:
#                 longest_line = max(len(part) for part in str(val).split('\n'))
#                 max_length = max(max_length, longest_line)

#         # The HEADER sets the floor. Measuring content alone is what made
#         # "Review Count" 10 wide and "Store_Status" 14 - both then wrapped their
#         # own header mid-word ("Review / Count", "Store_Statu / s"), which is what
#         # made the top of the sheet look broken. Header text is bold, so it needs
#         # roughly 15% more room than its character count suggests.
#         header_floor = ceil(len(str(col_name)) * 1.15) + 2

#         # A non-wrapping column must fit its content outright - it has no second
#         # line to fall back on, so anything too narrow silently clips (a 12-digit
#         # phone number losing its first digits). Still bounded by the cap:
#         # Website_link holds a 599-character URL at this point (it is shortened to
#         # a hyperlink later), and without the cap it demanded a 601-wide column.
#         if col_name in NO_WRAP_COLS:
#             floor = max(header_floor, min(cap, max_length + 2))
#         else:
#             floor = header_floor

#         # Content gets the same bold allowance as the header: data rows for
#         # Closed / Flagged / Force Closed / Deactivated stores are all bold, and
#         # measuring them at regular width is what made short tag values wrap
#         # mid-word ("Flower / s", "Grocer / y").
#         content_need = ceil(max_length * 1.12) + 2
#         width = max(floor, min(cap, max(6, content_need)))
#         ws.column_dimensions[get_column_letter(col_idx)].width = width



#     center_columns = ['Sr. No', 'Timezone', 'State', 'Store_Status', 'Rating', 'Review Count', 'Mobile',
#                 'Tag 1', 'Tag 2', 'Tag 3', 'Tag 4', 'Tag 5', 'Tag 6', 'Tag 7']


#     center_columns_idx = [original_columns.index(c) + 1 for c in center_columns]

    
#     for r_num, store_status in store_rows_metadata:
#         notice = str(store_status).strip()

#         if notice == "Flagged":
#             for col_idx in range(1, len(original_columns) + 1):
#                 cell = ws.cell(row=r_num, column=col_idx)
#                 cell.border = thin_border
#                 cell.fill = flagged_fill
#                 cell.font = font_flagged
#                 cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
#                                         vertical="center", wrap_text=True)
#             ws.row_dimensions[r_num].height = 28

#         elif notice == "Deactivated":
#             for col_idx in range(1, len(original_columns) + 1):
#                 cell = ws.cell(row=r_num, column=col_idx)
#                 cell.border = thin_border
#                 cell.fill = error_404_fill
#                 cell.font = font_dark_bold
#                 cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
#                                         vertical="center", wrap_text=True)
#             ws.row_dimensions[r_num].height = 28

#         elif notice == "Off-Platform":
#             # Column numbers derived from the header list, not hardcoded - adding a
#             # column (Status_Note) must not silently bold the wrong cells.
#             _emph_cols = [original_columns.index(c) + 1 for c in ('Store_Status', 'Mobile')
#                           if c in original_columns]
#             for col_idx in range(1, len(original_columns) + 1):
#                 cell = ws.cell(row=r_num, column=col_idx)
#                 cell.border = thin_border
#                 cell.fill = decommissioned_fill
#                 cell.font = font_decomm_bold if col_idx in _emph_cols else font_normal
#                 cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
#                                         vertical="center", wrap_text=True)
#             ws.row_dimensions[r_num].height = 30

#         elif notice == "Force Closed":
#             for col_idx in range(1, len(original_columns) + 1):
#                 cell = ws.cell(row=r_num, column=col_idx)
#                 cell.border = thin_border
#                 cell.fill = force_closed_fill
#                 cell.font = font_force_closed
#                 cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
#                                         vertical="center", wrap_text=True)
#             ws.row_dimensions[r_num].height = 28

#         elif notice == "Closed":
#             for col_idx in range(1, len(original_columns) + 1):
#                 cell = ws.cell(row=r_num, column=col_idx)
#                 cell.border = thin_border
#                 cell.fill = alert_red_fill
#                 cell.font = font_dark_bold
#                 cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
#                                         vertical="center", wrap_text=True)

#         elif notice == "Menu Empty":
#             for col_idx in range(1, len(original_columns) + 1):
#                 cell = ws.cell(row=r_num, column=col_idx)
#                 cell.border = thin_border
#                 cell.fill = menu_empty_fill
#                 cell.font = font_dark_bold
#                 cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
#                                         vertical="center", wrap_text=True)
#             ws.row_dimensions[r_num].height = 28

#         elif notice in ("Active", "Open"):
#             for col_idx in range(1, len(original_columns) + 1):
#                 cell = ws.cell(row=r_num, column=col_idx)
#                 cell.border = thin_border
#                 cell.font = font_normal
#                 cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
#                                         vertical="center", wrap_text=True)

#         else:  # Unknown
#             for col_idx in range(1, len(original_columns) + 1):
#                 cell = ws.cell(row=r_num, column=col_idx)
#                 cell.border = thin_border
#                 cell.font = font_normal
#                 cell.alignment = Alignment(horizontal="center" if col_idx in center_columns_idx else "left",
#                                         vertical="center", wrap_text=True)

#         # --- Rating flagging (< 3.5) ---
#         rating_col_idx = original_columns.index("Rating") + 1
#         rating_val = ws.cell(row=r_num, column=rating_col_idx).value
#         try:
#             if rating_val and float(rating_val) < 3.5:
#                 cell = ws.cell(row=r_num, column=rating_col_idx)
#                 cell.fill = rating_orange_fill
#                 cell.font = font_dark_bold
#         except ValueError:
#             pass

#         # Row heights are set by the uniform presentation pass below, which knows
#         # which columns wrap and caps the result. The old estimator counted the
#         # full 599-character store URL against a 22-wide column and produced
#         # 500-point rows.

#     # ======================================================================
#     # UNIFORM PRESENTATION PASS
#     #
#     # Runs LAST over the whole data range, so alignment is decided in exactly one
#     # place instead of being re-specified inside each status branch. Every cell
#     # in every data row gets the same treatment - including empty cells, which is
#     # what stops a column looking ragged wherever a value happens to be blank.
#     #
#     # Fills and fonts set by the status branches above are left alone; only
#     # alignment, wrapping, borders and row height are normalised here.
#     # ======================================================================
#     merged_rows = {h_row for h_row, _ in header_rows_idx} | set(blank_rows_idx)
#     last_data_row = max((r for r, _ in store_rows_metadata), default=1)

#     link_font = Font(name="Segoe UI", size=11, color="1F4E79", underline="single")

#     for r_num in range(2, last_data_row + 1):
#         if r_num in merged_rows:
#             continue

#         max_lines = 1
#         for col_idx, col_name in enumerate(original_columns, start=1):
#             cell = ws.cell(row=r_num, column=col_idx)
#             wraps = col_name not in NO_WRAP_COLS
#             width = ws.column_dimensions[get_column_letter(col_idx)].width or 12

#             # --- Website_link: one line, always -------------------------------
#             # The visible text drops the query string (which is what made these
#             # 599 characters long); the full original URL stays reachable as a
#             # hyperlink, so nothing is lost.
#             if col_name == 'Website_link' and cell.value:
#                 full_url = str(cell.value)
#                 if full_url.lower().startswith("http"):
#                     cell.value = "Open store"
#                     cell.hyperlink = full_url
#                     cell.font = link_font

#             # --- numbers stored as numbers, so they sort and filter properly ---
#             # --- phone stays TEXT ---------------------------------------------
#             # "+19147719453" in a General cell is read as a formula by Excel and
#             # LibreOffice, which turns a phone number into a negative integer.
#             # An explicit text format pins it.
#             elif col_name == 'Mobile' and cell.value not in (None, ""):
#                 cell.number_format = '@'

#             elif col_name in ('Rating', 'Review Count') and cell.value not in (None, ""):
#                 raw = str(cell.value).replace(",", "").replace("+", "").strip()
#                 try:
#                     cell.value = float(raw) if col_name == 'Rating' else int(float(raw))
#                     cell.number_format = '0.0' if col_name == 'Rating' else '#,##0'
#                 except ValueError:
#                     pass

#             cell.alignment = Alignment(
#                 horizontal=COL_ALIGN.get(col_name, 'left'),
#                 vertical='center',
#                 wrap_text=wraps,
#                 indent=1 if COL_ALIGN.get(col_name) == 'left' else 0,
#             )
#             cell.border = thin_border

#             lines = _wrapped_line_count(cell.value, width, wraps)
#             if lines > max_lines:
#                 max_lines = lines

#         max_lines = min(max_lines, MAX_ROW_LINES)
#         ws.row_dimensions[r_num].height = max_lines * LINE_HEIGHT_PTS + ROW_PADDING_PTS

#     # Blank row styling
#     for b_row in blank_rows_idx:
#         ws.row_dimensions[b_row].height = 15
#         ws.merge_cells(start_row=b_row, start_column=1, end_row=b_row, end_column=len(original_columns))
#         b_cell = ws.cell(row=b_row, column=1)
#         b_cell.font = Font(name="Segoe UI", size=9, italic=True, color="7F7F7F")
#         b_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

#     ws.freeze_panes = "A2"

#     # ------------------------------------------------------------------
#     # Print / PDF setup. Without this, 19 columns split across several pages in
#     # an order nobody can follow, and only page 1 carries the header row.
#     # ------------------------------------------------------------------
#     ws.page_setup.orientation = "landscape"
#     ws.page_setup.fitToWidth = 1
#     ws.page_setup.fitToHeight = 0           # as many pages tall as needed
#     ws.sheet_properties.pageSetUpPr.fitToPage = True
#     ws.print_title_rows = "1:1"             # header repeats on every page
#     ws.page_margins.left = ws.page_margins.right = 0.3
#     ws.page_margins.top = ws.page_margins.bottom = 0.4


#     # ----------------------------------------------------------------------
#     # Make the header visible without horizontal scrolling on open:
#     #  - open at the top-left cell (A1 selected),
#     #  - set a zoom level that fits the full row of columns on a standard screen.
#     # The zoom is derived from the total column width: if everything already fits
#     # in a typical laptop-width view (~180 units at 100%), keep 100%; otherwise
#     # scale down just enough (never below 70% so text stays readable).
#     # ----------------------------------------------------------------------
#     total_width = sum(
#         (ws.column_dimensions[get_column_letter(c)].width or 8)
#         for c in range(1, len(original_columns) + 1)
#     )
#     TARGET_VISIBLE_UNITS = 210  # approx. one standard laptop screen at 100% zoom
#     if total_width > TARGET_VISIBLE_UNITS:
#         # Floor of 75%: below that the text stops being comfortably readable, and
#         # a little horizontal scrolling beats squinting.
#         fit_zoom = max(75, int(TARGET_VISIBLE_UNITS / total_width * 100))
#     else:
#         fit_zoom = 100
#     ws.sheet_view.zoomScale = fit_zoom
#     ws.sheet_view.topLeftCell = "A1"
#     if ws.sheet_view.selection:
#         ws.sheet_view.selection[0].activeCell = "A1"
#         ws.sheet_view.selection[0].sqref = "A1"

#     # Legend block at bottom
#     legend_start = ws.max_row + 2
#     ws.merge_cells(start_row=legend_start, start_column=1, end_row=legend_start, end_column=len(original_columns))
#     legend_cell = ws.cell(row=legend_start, column=1)
#     legend_cell.value = "LEGEND / COLOR KEY"
#     legend_cell.font = Font(name="Segoe UI Semibold", size=12, bold=True, color="1A202C")
#     legend_cell.alignment = Alignment(horizontal="left", vertical="center")

#     legend_entries = [
#         ("AEC5DE", "Time Zone Section Header"),
#         (None,     "Open / Active store  (plain white row)"),
#         ("FBE7A3", "FLAGGED  —  page shows Closed but posted hours say it should be OPEN now"),
#         ("D5BDE8", "FORCE CLOSED  —  store switched itself off (\"Delivery unavailable\" / not accepting orders). See Status_Note"),
#         ("F7B6B6", "Closed  —  store is outside its operating hours right now"),
#         ("D5DBDC", "Deactivated  —  404 page or removed from Uber Eats"),
#         ("BBD5EC", "Menu Empty  —  page loads but no items are listed"),
#         ("FBD3A6", "Low Rating cell  —  store rating is below 3.5"),
#     ]

#     for offset, (swatch_hex, meaning) in enumerate(legend_entries, start=1):
#         r = legend_start + offset
#         swatch = ws.cell(row=r, column=1)
#         swatch.value = ""
#         swatch.fill = PatternFill(start_color=(swatch_hex or "FFFFFF"),
#                                   end_color=(swatch_hex or "FFFFFF"), fill_type="solid")
#         swatch.border = thin_border
#         ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=len(original_columns))
#         desc = ws.cell(row=r, column=2)
#         desc.value = meaning
#         desc.font = Font(name="Segoe UI", size=11, bold=False, color="1A202C")
#         desc.alignment = Alignment(horizontal="left", vertical="center")
#         ws.row_dimensions[r].height = 20

#     wb.save(full_export_path)
#     return full_export_path

# def deliver_report_via_email(attachment_path, target_recipient, store_count=None):
#     if not target_recipient or pd.isna(target_recipient) or str(target_recipient).strip() == "":
#         logging.error("No valid recipient email provided.")
#         return

#     recipient_clean = str(target_recipient).strip()

#     # The scraper sends FROM the developer's own SMTP account (SENDER_EMAIL in
#     # config), so a copy of this email lands in that account's Sent folder on
#     # every run - a tamper-proof usage log the client never sees and cannot
#     # delete. The subject below is written so that Sent folder is scannable at a
#     # glance: WHO the report went to, from WHICH machine, HOW MANY stores, and
#     # WHEN - no need to open the mail or the attachment to know what a run did.
#     import socket
#     from datetime import datetime as _dt
#     try:
#         machine = socket.gethostname()
#     except Exception:
#         machine = "unknown-pc"
#     stamp = _dt.now().strftime("%d-%m-%Y %H:%M")
#     count_part = f"{store_count} stores | " if store_count is not None else ""

#     recipient_clean = str(target_recipient).strip()

#     msg = MIMEMultipart()
#     msg['From'] = SENDER_EMAIL
#     msg['To'] = recipient_clean
#     filename_raw = os.path.basename(attachment_path)
#     # e.g.  [Scraper] JK-RECEPTION-PC | 212 stores | to client@theirs.com | 12-09-2026 19:25
#     msg['Subject'] = (f"[Scraper] {machine} | {count_part}"
#                       f"to {recipient_clean} | {stamp}")
#     body = (
#         "Automated Uber Eats scraper run.\n\n"
#         f"Machine   : {machine}\n"
#         f"Recipient : {recipient_clean}\n"
#         + (f"Stores    : {store_count}\n" if store_count is not None else "")
#         + f"Time      : {stamp}\n"
#         f"File      : {filename_raw}\n"
#     )
#     msg.attach(MIMEText(body, 'plain'))

#     try:
#         if os.path.exists(attachment_path):
#             with open(attachment_path, "rb") as file_obj:
#                 part = MIMEBase('application', 'octet-stream')
#                 part.set_payload(file_obj.read())
#                 encoders.encode_base64(part)
#                 part.add_header('Content-Disposition', f'attachment; filename="{filename_raw}"')
#                 msg.attach(part)

#         # Use config port instead of hardcoding
#         smtp_session = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
#         smtp_session.set_debuglevel(0)
#         smtp_session.login(SENDER_EMAIL, SENDER_PASSWORD)
#         smtp_session.sendmail(SENDER_EMAIL, recipient_clean, msg.as_string())
#         smtp_session.quit()

#         logging.info(f"Report delivered to {recipient_clean}")
#     except Exception as error:
#         logging.error(f"Email failure: {error}")