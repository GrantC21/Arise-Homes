"""
PO Viewer auto-renamer for Arise Homes.

Watches the Windows Downloads folder. When a PDF whose filename contains
"PO Viewer" appears, it reads fixed areas of page 1, looks up the vendor /
subdivision / PO type in po_rename_config.txt, and renames the file in place
to:

    [VendorShort]_[SubdivisionAbbrev][Lot]_[Address]_[PO Type]_[DD Mon YYYY].pdf

Several purchase order templates are read by the same code. Rather than
keying off template-specific wording, the jobsite is located by finding the
unit marker (Lot / Homesite / BLDG) and reading the block's shape around
it, and the column positions are measured per document. That covers the
ERP export, the Excel purchase order printed to PDF, and multi-family POs,
at whatever scale or margin they happen to use. A PDF with no jobsite
marker at all is flagged rather than guessed at.

If anything can't be confidently determined (unexpected layout, a value not
yet in the config table, or a filename collision), the file is left alone
content-wise and instead renamed to "ERROR - PO Viewer.pdf" (or
"ERROR - PO Viewer (1).pdf", "(2)", ... if that name is already taken) so
it's obviously flagged and nothing gets overwritten or silently mis-named.

The lookup tables are re-read whenever po_rename_config.txt changes, so
adding a new vendor / subdivision / PO type takes effect immediately -
there's no need to restart the watcher after editing it.

Run this with pythonw.exe (no console window) via a Task Scheduler
"At log on" trigger. See SETUP.md for step-by-step instructions.
"""

import argparse
import logging
import os
import queue
import re
import sys
import threading
import time
import traceback
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path

import pdfplumber
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

# ----------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "po_rename_config.txt"
DOWNLOADS_DIR = Path.home() / "Downloads"

# Days of log history to keep. Each day gets its own file; anything older
# than this is deleted automatically.
LOG_RETENTION_DAYS = 7


# Logs live in a "logs" subfolder of the tool folder: still everything in
# one place, but the daily rotations don't clutter the folder you actually
# open to edit the config. If that folder is synced (OneDrive etc.) the log
# is re-uploaded whenever it changes, which is a fine trade at this volume.
LOG_DIR = SCRIPT_DIR / "logs"
try:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
except OSError:
    LOG_DIR = SCRIPT_DIR
LOG_PATH = LOG_DIR / "po_rename_log.txt"


def _tidy_old_logs():
    """
    Moves logs left in the tool folder by an earlier version into the logs
    subfolder, so upgrading doesn't leave the old ones lying around.

    Only touches this tool's own log files, never overwrites anything
    already in the subfolder, and gives up quietly - a file the running
    watcher still holds open is simply left where it is.
    """
    if LOG_DIR == SCRIPT_DIR:
        return
    for stray in SCRIPT_DIR.glob("po_rename_log.txt*"):
        if not stray.is_file():
            continue
        destination = LOG_DIR / stray.name
        if destination.exists():
            continue
        try:
            stray.rename(destination)
        except OSError:
            pass


_tidy_old_logs()

TRIGGER_TEXT = "po viewer"          # filename must contain this (case-insensitive)
ERROR_STEM = "ERROR - PO Viewer"    # base name used when a file can't be processed

# Fallback x-position (points from the left edge) separating the Vendor
# block from the jobsite / summary-table block, used only when the gutter
# between the two columns can't be measured on the page itself. Templates
# differ in scale and margins, so the measured value is preferred.
RIGHT_COLUMN_X_MIN = 280


class SafeTimedRotatingFileHandler(TimedRotatingFileHandler):
    """
    Daily-rotating log handler that won't take the watcher down if the log
    file is momentarily locked.

    On Windows the log can be held open by something else at the moment we
    try to roll it over - a `--file` retry run from Command Prompt while the
    background watcher is up, or a sync client (OneDrive) mid-upload.
    Rather than raising, keep appending to the current file and retry the
    rollover a little later.
    """

    def doRollover(self):
        try:
            super().doRollover()
        except OSError:
            if self.stream is None:
                try:
                    self.stream = self._open()
                except OSError:
                    pass
            self.rolloverAt = int(time.time()) + 3600


# A logging failure must never crash the renamer.
logging.raiseExceptions = False

_handler = SafeTimedRotatingFileHandler(
    str(LOG_PATH),
    when="midnight",
    backupCount=LOG_RETENTION_DAYS,
    encoding="utf-8",
    delay=True,          # don't create/lock the file until something is logged
)
_handler.setFormatter(logging.Formatter("%(asctime)s  %(levelname)s  %(message)s"))

_logger = logging.getLogger("po_rename")
_logger.setLevel(logging.INFO)
_logger.addHandler(_handler)
_logger.propagate = False


def log(msg, level=logging.INFO):
    _logger.log(level, msg)
    print(msg)   # no-op under pythonw.exe, useful when run from a console


# ----------------------------------------------------------------------
# Config table loading
# ----------------------------------------------------------------------

def normalize_ws(s):
    return re.sub(r"\s+", " ", (s or "")).strip()


def load_config(path):
    """
    Parses po_rename_config.txt. Sections are marked with a line starting
    with '#' (e.g. "#Vendors"). Within a section, each line is
    "Full Name<delim>Short/Abbrev Value" where <delim> is "|" or "_"
    (whichever is present - "|" takes priority if both appear).
    Returns dict with keys 'vendors', 'subdivisions', 'po_types'.
    """
    section_map = {
        "vendors": "vendors",
        "subdivisions": "subdivisions",
        "po types": "po_types",
        "potypes": "po_types",
    }
    vendors = {}
    subdivisions = []  # list of (full_name_lower, abbrev) - order preserved
    po_types = {}

    current = None
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("#"):
            key = line.lstrip("#").strip().lower()
            current = section_map.get(key)
            continue
        if current is None:
            continue

        if "|" in line:
            full, _, short = line.partition("|")
        elif "_" in line:
            full, _, short = line.partition("_")
        else:
            log(f"Config: skipping unparseable line: {raw_line!r}", logging.WARNING)
            continue

        full = normalize_ws(full)
        short = normalize_ws(short)
        if not full or not short:
            log(f"Config: skipping incomplete line: {raw_line!r}", logging.WARNING)
            continue

        if current == "vendors":
            vendors[full.lower()] = short
        elif current == "subdivisions":
            subdivisions.append((full.lower(), short))
        elif current == "po_types":
            po_types[full.lower()] = short

    # Longest subdivision name first, so "Stoneridge South" is checked
    # before any shorter/overlapping name.
    subdivisions.sort(key=lambda t: -len(t[0]))

    return {"vendors": vendors, "subdivisions": subdivisions, "po_types": po_types}


def describe_config(config):
    return (
        f"{len(config['vendors'])} vendors, "
        f"{len(config['subdivisions'])} subdivisions, "
        f"{len(config['po_types'])} PO types"
    )


class ConfigLoader:
    """
    Serves the lookup tables, re-reading po_rename_config.txt whenever it
    changes on disk. This means adding a new vendor / subdivision / PO type
    takes effect on the very next PDF - no need to restart the watcher.
    """

    def __init__(self, path):
        self.path = path
        self._mtime = None
        self._config = None

    def get(self):
        try:
            mtime = self.path.stat().st_mtime
        except OSError as exc:
            if self._config is not None:
                log(f"Config unreadable ({exc}); using the last good copy.", logging.WARNING)
                return self._config
            raise

        if self._config is not None and mtime == self._mtime:
            return self._config

        new_config = load_config(self.path)

        # A config with nothing in it usually means we caught the file
        # mid-save (Notepad truncates then rewrites). Keep the previous
        # tables rather than failing every PDF until the next edit.
        if self._config is not None and not any(new_config.values()):
            log("Config read back empty - keeping the previous tables.", logging.WARNING)
            return self._config

        self._config = new_config
        self._mtime = mtime
        log(f"Loaded config: {describe_config(new_config)}.")
        return self._config


# ----------------------------------------------------------------------
# PDF field extraction
# ----------------------------------------------------------------------

def get_visual_lines(page, x_min=None, x_max=None, tol=2.5):
    """
    Reconstructs the page's visual reading order: groups words into rows by
    vertical position (not the PDF's internal text-object order, which can
    scatter labels and values apart), then sorts each row left-to-right.
    Returns a list of (top, text) tuples, top-to-bottom.
    """
    words = page.extract_words()
    if x_min is not None:
        words = [w for w in words if w["x0"] >= x_min]
    if x_max is not None:
        words = [w for w in words if w["x0"] <= x_max]
    words.sort(key=lambda w: (w["top"], w["x0"]))

    rows = []
    current_row, current_top = [], None
    for w in words:
        if current_top is None or abs(w["top"] - current_top) <= tol:
            current_row.append(w)
            current_top = w["top"] if current_top is None else current_top
        else:
            rows.append(current_row)
            current_row, current_top = [w], w["top"]
    if current_row:
        rows.append(current_row)

    lines = []
    for row in rows:
        row_sorted = sorted(row, key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in row_sorted)
        top = min(w["top"] for w in row)
        lines.append((top, text))
    return lines


def _find_vendor(left_lines):
    """First line under the "Vendor:" label in the left column."""
    vendor_label_top = None
    for top, text in left_lines:
        if text.strip().rstrip(":").lower() == "vendor":
            vendor_label_top = top
            break
    if vendor_label_top is None:
        return None
    below = sorted(
        (t for t in left_lines if t[0] > vendor_label_top + 1),
        key=lambda t: t[0],
    )
    if below and (below[0][0] - vendor_label_top) < 40:
        return below[0][1].strip()
    return None


# A row of the summary table starts with a short label and a colon
# ("PO Type:", "Floor Plan:", "Region:"). A line without one, sitting just
# below, is the tail of a value that wrapped.
SUMMARY_LABEL_RE = re.compile(r"^[^:]{1,25}:")

# How far below a value a wrapped remainder can sit, and how many such
# lines to accept. Rows in the summary table are ~15pt apart, so this
# reaches the next line but not the block further down the page.
WRAP_MAX_GAP = 20.0
WRAP_MAX_LINES = 3


def _find_summary_value(right_lines, label_pattern):
    """
    Value of one row of the right-hand summary table.

    A long value wraps onto following lines within its cell - e.g.
    "*Trim Material - Interior Trim and" / "Doors" - so any unlabelled line
    immediately beneath is joined back on. Reading only the first line would
    mean keeping a lookup entry that stops wherever the text happened to
    wrap, which breaks as soon as the wording or column width changes.
    """
    ordered = sorted(right_lines, key=lambda t: t[0])
    for i, (top, text) in enumerate(ordered):
        m = re.search(label_pattern, text.strip(), re.IGNORECASE)
        if not m:
            continue
        parts = [m.group(1).strip()]
        previous_top = top
        for next_top, next_text in ordered[i + 1:][:WRAP_MAX_LINES]:
            candidate = next_text.strip()
            if (next_top - previous_top) > WRAP_MAX_GAP:
                break
            if not candidate or SUMMARY_LABEL_RE.match(candidate):
                break
            parts.append(candidate)
            previous_top = next_top
        return normalize_ws(" ".join(parts))
    return None


def _find_po_type(right_lines):
    """The ERP labels this row "PO Type:" and the Excel sheet just "Type:",
    so match the part they share."""
    return _find_summary_value(right_lines, r"\bType:\s*(.+)$")


def _find_region(right_lines):
    """
    The "Region:" row, e.g. "JOHNSON COUNTY SF" or "JOHNSON COUNTY MF".

    This is how a multi-family purchase order is told from a single-family
    one, which decides whether a plain lot number is really a lot or a
    building-and-side code. See resolve_subdivision_and_lot.
    """
    return _find_summary_value(right_lines, r"\bRegion:\s*(.+)$")


def find_column_split(page, default=RIGHT_COLUMN_X_MIN):
    """
    Measures the gutter between the Vendor column and the jobsite column.

    Templates place these blocks at different scales and margins, so a fixed
    cutoff eventually lands inside one of the columns. Instead, look at the
    band of text starting at the "Vendor:" label and take the middle of the
    widest horizontal gap in it - that's the gutter, wherever it happens to
    fall on a given template.
    """
    words = page.extract_words()
    vendor_top = None
    for w in words:
        if w["text"].strip().rstrip(":").lower() == "vendor":
            vendor_top = w["top"]
            break
    if vendor_top is None:
        return default

    band = [w for w in words if vendor_top - 5 <= w["top"] <= vendor_top + 80]
    if not band:
        return default

    spans = sorted((w["x0"], w["x1"]) for w in band)
    best_gap, best_mid = 0.0, default
    cursor = spans[0][1]
    for x0, x1 in spans[1:]:
        if x0 - cursor > best_gap:
            best_gap, best_mid = x0 - cursor, (cursor + x0) / 2
        cursor = max(cursor, x1)

    # Too narrow to be a column gutter - don't trust it.
    if best_gap < 30:
        return default
    return best_mid


def _find_label_top(lines, label):
    """Vertical position of a standalone label line such as "Vendor:"."""
    for top, text in lines:
        if text.strip().rstrip(":").lower() == label:
            return top
    return None


def _jobsite_block(right_lines, vendor_top):
    """
    The right-hand jobsite lines, top to bottom.

    Starts at the "Ship To:" heading when there is one, otherwise at the
    Vendor label's height - the jobsite always sits alongside the vendor
    block. Either way the summary table above (PO number, Region, Date) is
    excluded, so a PO number like "PO-C4-WALLS-BLDG10-1" can't be mistaken
    for a building number.
    """
    ordered = sorted(right_lines, key=lambda t: t[0])
    anchor = _find_label_top(ordered, "ship to")
    if anchor is None:
        for top, text in ordered:
            if "ship to" in text.strip().lower():
                anchor = top
                break
    if anchor is None:
        anchor = vendor_top
    if anchor is None:
        return ordered
    return [t for t in ordered if -5 <= (t[0] - anchor) < 110 and t[0] > anchor]


def _extract_jobsite(right_lines, vendor_top):
    """
    Reads the subdivision, unit number and street address off the jobsite.

    The templates arrange these three around the unit marker in every
    combination seen so far - the plat sometimes leads the marker line and
    sometimes follows it, and the address sometimes shares that line and
    sometimes sits on the line above:

        18460 W 195th Ter                   18505 W 195th Ter
        GARRETT RANCH THIRD PLAT, Lot 31    Lot 39, GARRETT RANCH THIRD PLAT

        Garrett Ranch                       Stoneridge North MF
        Homesite 31, 18460 W 195th Ter      Building 5, 26055-26057 W. 82nd Ter

    Position is therefore no guide. Each candidate is judged on what it
    looks like instead: the piece shaped like a street address is the
    address, and the remaining non-address piece is the subdivision.

    Returns (plot_raw, address_raw), either of which may be None.
    """
    block = _jobsite_block(right_lines, vendor_top)
    for i, (_, raw) in enumerate(block):
        text = raw.strip()
        match = UNIT_RE.search(text)
        if not match:
            continue

        marker = text[match.start():match.end()].strip()
        before = text[:match.start()].strip(" ,")
        after = text[match.end():].strip(" ,")
        above = block[i - 1][1].strip() if i > 0 else ""

        address = next(
            (part for part in (after, before, above)
             if part and looks_like_street_address(part)),
            None,
        )
        subdivision = next(
            (part for part in (before, after, above)
             if part and part != address and not looks_like_street_address(part)),
            "",
        )
        plot = f"{subdivision}, {marker}".strip(" ,")
        return (plot or None), address
    return None, None


def extract_raw_fields(pdf_path):
    """
    Returns a dict with vendor_raw, po_type_raw, address_raw, plot_raw.
    Any field that can't be located is None.
    """
    empty = {"vendor_raw": None, "po_type_raw": None, "address_raw": None,
             "plot_raw": None, "region_raw": None}

    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            return empty
        page = pdf.pages[0]
        split = find_column_split(page)
        left_lines = get_visual_lines(page, x_max=split)
        right_lines = get_visual_lines(page, x_min=split)

        vendor_top = _find_label_top(left_lines, "vendor")
        plot_raw, address_raw = _extract_jobsite(right_lines, vendor_top)
        if plot_raw is None:
            # No jobsite marker anywhere - not a purchase order we know how
            # to read. Leave every field unset so it gets flagged.
            return empty
        return {"vendor_raw": _find_vendor(left_lines),
                "po_type_raw": _find_po_type(right_lines),
                "address_raw": address_raw,
                "plot_raw": plot_raw,
                "region_raw": _find_region(right_lines)}


# ----------------------------------------------------------------------
# Lookup + filename construction
# ----------------------------------------------------------------------

def resolve_vendor(vendor_raw, vendor_table):
    if not vendor_raw:
        return None
    # "Century Building Solutions - Spring Hill" -> "Century Building Solutions"
    truncated = normalize_ws(vendor_raw.split(" - ")[0])
    return vendor_table.get(truncated.lower())


# The marker naming the specific home on the jobsite line. Single-family
# POs write "Lot 31" (ERP) or "Homesite 31" (Excel); multi-family and villa
# POs write "BLDG 10". Anchored to a space or line start so it can't match
# inside a PO number such as "PO-C4-WALLS-BLDG10-1".
UNIT_RE = re.compile(
    r"(?:^|\s)(?:Lot|Homesite|Bldg\.?|Building|Unit)\s*#?\s*(\d+[A-Za-z]?)\b",
    re.IGNORECASE,
)


# Markers that name a building rather than a single-family lot. Building
# numbers are joined to the subdivision with a dash and lot numbers run
# straight on, so a villa reads "159-10" where a lot reads "GR31".
BUILDING_MARKERS = ("bldg", "building", "unit")

# The Region row marks a purchase order as multi-family.
MULTIFAMILY_RE = re.compile(r"\bMF\b", re.IGNORECASE)

# On a multi-family PO the ERP can write the jobsite as a plain lot number
# that actually encodes the building and which half of the duplex it is -
# "Lot 32" means building 3, side 2. The side is not wanted in the filename.
DUPLEX_SIDES = ("1", "2")


def is_multifamily(region_raw):
    return bool(region_raw and MULTIFAMILY_RE.search(region_raw))


def resolve_subdivision_and_lot(plot_raw, subdivision_table, region_raw=None):
    """
    Returns (subdivision abbreviation, unit suffix) - e.g. ("GR", "31") for
    a lot, or ("159", "-10") for a building. The two concatenate to form the
    second field of the filename.

    A marker that names a building already gives the number wanted. A plain
    lot number normally does too - except on a multi-family PO, where it is
    the building followed by the duplex side: "Lot 32" in an MF region is
    building 3, side 2, and becomes "-3".

    That rewrite is deliberately confined to MF regions. Applied to a
    single-family PO it would turn lot 31 into building 3, so anything that
    doesn't fit the pattern - a single digit, or a last digit that isn't a
    side - is reported as unresolved rather than guessed at.
    """
    if not plot_raw:
        return None, None
    text_lower = plot_raw.lower()
    matches = [(full, abbr) for full, abbr in subdivision_table if full in text_lower]
    if len(matches) != 1:
        return None, None
    _, abbr = matches[0]
    unit_match = UNIT_RE.search(plot_raw)
    if not unit_match:
        return None, None
    marker = unit_match.group(0).strip().split()[0].rstrip(".").lower()
    number = unit_match.group(1)

    if marker.startswith(BUILDING_MARKERS):
        return abbr, f"-{number}"

    if is_multifamily(region_raw):
        if not number.isdigit() or len(number) < 2 or number[-1] not in DUPLEX_SIDES:
            return None, None
        return abbr, f"-{number[:-1]}"

    return abbr, number


def resolve_po_type(po_type_raw, po_type_table):
    if not po_type_raw:
        return None
    return po_type_table.get(normalize_ws(po_type_raw).lower())


# A street address starts with a house number (optionally suffixed, e.g.
# "123A") followed by the street name. Duplexes are billed against both
# halves at once and write the pair as a range - "26055-26057 W. 82nd Ter" -
# so a second number after a dash is allowed too.
#
# The address is picked positionally, so this guards against grabbing a
# neighbouring line (e.g. "Arise Homes LLC" or a subdivision name) when a
# jobsite block's shape differs from the usual template.
STREET_ADDRESS_RE = re.compile(r"^\d+[A-Za-z]?(?:\s*-\s*\d+[A-Za-z]?)?\s+\S")


def looks_like_street_address(s):
    return bool(STREET_ADDRESS_RE.match(s or ""))


def sanitize_part(s):
    s = re.sub(r'[\\/:*?"<>|]', "-", s)
    return s.strip(" .")


def build_target_filename(vendor_short, subdivision_abbr, unit_suffix, address, po_type_value, date_str):
    parts = [
        vendor_short,
        f"{subdivision_abbr}{unit_suffix}",
        address,
        po_type_value,
        date_str,
    ]
    return "_".join(sanitize_part(p) for p in parts) + ".pdf"


def unique_error_name(folder):
    candidate = f"{ERROR_STEM}.pdf"
    if not (folder / candidate).exists():
        return candidate
    i = 1
    while True:
        candidate = f"{ERROR_STEM} ({i}).pdf"
        if not (folder / candidate).exists():
            return candidate
        i += 1


# ----------------------------------------------------------------------
# Processing
# ----------------------------------------------------------------------

def looks_like_complete_pdf(path):
    """
    True if the file ends with the PDF end-of-file marker, i.e. the download
    has written the whole document. Lets a finished PDF be picked up
    immediately instead of sitting through the size-stability wait.
    """
    try:
        with path.open("rb") as fh:
            fh.seek(0, os.SEEK_END)
            tail_size = min(1024, fh.tell())
            fh.seek(-tail_size, os.SEEK_END)
            return b"%%EOF" in fh.read(tail_size)
    except OSError:
        return False


# How long to keep waiting for a download to finish before giving up on it.
DOWNLOAD_TIMEOUT = 60

# How long a file lacking the %%EOF marker must hold the same size before
# it counts as finished. A download can pause mid-transfer, so this has to
# be comfortably longer than a normal network stall - checking twice in
# quick succession would call a paused download "complete" and try to read
# a half-written PDF.
STABLE_SECONDS = 3.0

# How long to keep waiting for a file that can't even be opened for
# reading before giving up on that check and letting the content and size
# checks decide.
LOCK_GRACE_SECONDS = 5.0

# path -> (last observed size, when it first reached that size), and
# path -> when the file was first looked at. Only touched by the worker.
_pending_sizes = {}
_first_seen = {}


def forget_pending(path):
    """Drops the bookkeeping for a file we're done tracking."""
    _pending_sizes.pop(path, None)
    _first_seen.pop(path, None)
    _finish_warned.discard(path)


WINDOWS = os.name == "nt"

# How long to wait for the internet-source stamp before renaming anyway.
#
# Kept short on purpose. Only files that actually came from the internet
# are ever stamped, so a PDF saved or printed locally - an Excel purchase
# order, say - will never get one and pays this wait in full. When a stamp
# is coming it lands about a second after the bytes do, so waiting longer
# buys very little and slows the local case down for nothing.
BROWSER_FINISH_GRACE = 3.0

_finish_warned = set()


def is_file_readable(path):
    """
    True when the file can be opened for reading.

    Deliberately read-only. An earlier version asked for write access to
    detect whether the browser still held the file, but that probe runs
    every time readiness is checked - repeatedly, while the download is
    being finalised - and asking for write access to a file Windows is busy
    scanning can make the scan itself fail.
    """
    try:
        with path.open("rb"):
            return True
    except OSError:
        return False


def browser_has_finished(path):
    """
    True when a download's source has finished with the file, not merely
    finished writing its bytes.

    On Windows a file that came from the internet is stamped with a
    Zone.Identifier stream. The call that writes it is the same one that
    runs the antivirus scan, so its presence means the scan is done and the
    file is ours to rename. Renaming before then makes that call fail and
    the browser reports the whole download as failed - "Couldn't download -
    Virus scan failed".

    A file that did not come from the internet is never stamped, so this
    stays False for something saved or printed locally out of Excel. That
    is what BROWSER_FINISH_GRACE bounds: such a file waits it out once and
    is then renamed normally.

    Off Windows there is no such marker, so the size and content checks
    decide on their own.
    """
    if not WINDOWS:
        return True
    try:
        return os.path.exists(f"{path}:Zone.Identifier")
    except OSError:
        return True


def is_download_complete(path):
    """
    Single, non-blocking readiness check.

    A finished PDF ends with the %%EOF marker, so the usual case is decided
    instantly and correctly. A file without that marker is only accepted
    once its size has held steady for STABLE_SECONDS.
    """
    try:
        size = path.stat().st_size
    except OSError:
        return False
    if size == 0:
        return False

    now = time.time()
    first = _first_seen.setdefault(path, now)
    waited = now - first

    if not is_file_readable(path) and waited < LOCK_GRACE_SECONDS:
        return False

    # Bytes on disk is not the same as the browser being done: its last step
    # is the Windows call that scans the file and stamps it. Renaming inside
    # that window makes the download fail outright, so wait for the stamp -
    # but not forever, in case a setup never produces one.
    if not browser_has_finished(path):
        if waited < BROWSER_FINISH_GRACE:
            return False
        if path not in _finish_warned:
            _finish_warned.add(path)
            log(f"  '{path.name}' has no internet-source mark (saved locally "
                f"rather than downloaded?); renaming after "
                f"{BROWSER_FINISH_GRACE:.0f}s.")

    if looks_like_complete_pdf(path):
        forget_pending(path)
        return True

    previous = _pending_sizes.get(path)
    if previous is None or previous[0] != size:
        _pending_sizes[path] = (size, now)
        return False
    return (now - previous[1]) >= STABLE_SECONDS


def process_file(path, config):
    folder = path.parent

    if not path.exists():
        # A duplicate filesystem event for a file another event already
        # renamed/handled - nothing to do.
        return

    log(f"Detected: {path.name}")

    try:
        fields = extract_raw_fields(path)
    except Exception as exc:  # corrupt/unreadable PDF, etc.
        log(f"  Failed to read PDF: {exc}", logging.ERROR)
        _mark_error(path, folder)
        return

    vendor_short = resolve_vendor(fields["vendor_raw"], config["vendors"])
    subdivision_abbr, unit_suffix = resolve_subdivision_and_lot(
        fields["plot_raw"], config["subdivisions"], fields.get("region_raw"))
    po_type_value = resolve_po_type(fields["po_type_raw"], config["po_types"])
    address = normalize_ws(fields["address_raw"]) if fields["address_raw"] else None

    missing = []
    if not vendor_short:
        missing.append(f"vendor (read: {fields['vendor_raw']!r})")
    if not subdivision_abbr or not unit_suffix:
        missing.append(
            f"subdivision/lot (read: {fields['plot_raw']!r}, "
            f"region: {fields.get('region_raw')!r})")
    if not po_type_value:
        missing.append(f"PO type (read: {fields['po_type_raw']!r})")
    if not address:
        missing.append("address")
    elif not looks_like_street_address(address):
        missing.append(f"address doesn't look like a street address (read: {address!r})")

    if missing:
        log(f"  Could not resolve: {'; '.join(missing)}", logging.WARNING)
        _mark_error(path, folder)
        return

    date_str = datetime.now().strftime("%d %b %Y")
    target_name = build_target_filename(vendor_short, subdivision_abbr, unit_suffix, address, po_type_value, date_str)
    target_path = folder / target_name

    if target_path.exists():
        log(f"  Target filename already exists (duplicate PO?): {target_name}", logging.WARNING)
        _mark_error(path, folder)
        return

    path.rename(target_path)
    log(f"  Renamed to: {target_name}")


def _mark_error(path, folder):
    if path.stem.lower().startswith(ERROR_STEM.lower()):
        # Already flagged (e.g. a --file retry that still can't resolve).
        # Renaming again would just shuffle it between "ERROR - PO Viewer.pdf"
        # and "ERROR - PO Viewer (1).pdf" and lose track of which file you
        # were retrying, so leave the name alone.
        log("  Still unresolved - leaving the existing ERROR filename as-is.", logging.WARNING)
        return
    try:
        error_name = unique_error_name(folder)
        path.rename(folder / error_name)
        log(f"  Renamed to: {error_name}", logging.WARNING)
    except OSError as exc:
        log(f"  Could not rename to error name: {exc}", logging.ERROR)


# ----------------------------------------------------------------------
# Watcher
# ----------------------------------------------------------------------

# The folder is re-scanned to catch POs that arrived without a usable
# filesystem event: Windows delivers directory notifications through a
# fixed-size buffer, and a burst of downloads (plus the renames this tool
# makes in the same folder) can overflow it and drop notifications
# silently.
#
# Rather than polling on a timer, a scan is scheduled for this long after
# the folder last changed, and each new change pushes it back - so a run of
# downloads produces one scan once they settle. A dropped notification is
# recoverable because the events that did survive still arm the scan.
#
# There is deliberately no periodic fallback: with nothing happening in the
# folder, nothing runs at all.
SCAN_DEBOUNCE = 5.0

# Set whenever the watched folder changes, and on shutdown so the waiting
# scan thread can exit.
_activity_event = threading.Event()


def note_activity():
    """Records that the watched folder changed, arming the debounced scan."""
    _activity_event.set()


_queue_lock = threading.Lock()
_in_flight = set()          # queued or being worked on right now
_abandoned = {}             # path -> (size, mtime) we already gave up on


def is_candidate(path):
    """Whether a file is a PO this tool should try to rename."""
    if path.suffix.lower() != ".pdf":
        return False
    stem = path.stem.lower()
    # Never re-trigger on our own error output - renaming a file inside the
    # watched folder fires an event, and an ERROR file's name still contains
    # "po viewer", which would otherwise make it reprocess itself in a loop.
    if stem.startswith(ERROR_STEM.lower()):
        return False
    return TRIGGER_TEXT in stem


def enqueue_file(work_queue, path):
    """Queues a PO unless it's already waiting, so the scan can't pile up
    duplicates of a file the events already reported."""
    with _queue_lock:
        if path in _in_flight:
            return False
        _in_flight.add(path)
    work_queue.put((path, time.time() + DOWNLOAD_TIMEOUT))
    return True


def release_file(path, abandoned_stat=None):
    """Marks a PO as no longer in flight. When given a stat, remembers that
    this exact file was given up on, so the scan doesn't retry it forever -
    but a later download reusing the name has a different size or timestamp
    and gets picked up normally."""
    with _queue_lock:
        _in_flight.discard(path)
        if abandoned_stat is None:
            _abandoned.pop(path, None)
        else:
            _abandoned[path] = abandoned_stat


def scan_folder(work_queue, folder):
    """Queues any PO sitting in the folder that isn't already in hand."""
    try:
        entries = list(folder.iterdir())
    except OSError as exc:
        log(f"Could not scan {folder}: {exc}", logging.WARNING)
        return
    for path in sorted(entries):
        if not path.is_file() or not is_candidate(path):
            continue
        try:
            stat = path.stat()
        except OSError:
            continue
        with _queue_lock:
            if _abandoned.get(path) == (stat.st_size, stat.st_mtime):
                continue
        enqueue_file(work_queue, path)


def scan_loop(work_queue, folder, stop_event):
    """
    Scans once the folder has gone quiet for SCAN_DEBOUNCE, once per burst.

    Blocks until something actually changes rather than waking on a timer,
    so an idle machine does no work at all - not even a tick.
    """
    while not stop_event.is_set():
        try:
            _activity_event.wait()          # nothing to do until the folder changes
            if stop_event.is_set():
                return
            _activity_event.clear()

            # Hold off while activity continues; each change restarts the wait.
            while _activity_event.wait(SCAN_DEBOUNCE):
                if stop_event.is_set():
                    return
                _activity_event.clear()

            scan_folder(work_queue, folder)
        except Exception as exc:
            log(f"Unexpected error scanning {folder}: {exc}", logging.ERROR)


class PoViewerHandler(FileSystemEventHandler):
    """
    Filters filesystem events and hands matching PDFs to a worker thread.

    The actual processing deliberately does NOT happen here: watchdog
    dispatches events on a single thread, so waiting on a slow or stalled
    download inline would hold up every PO queued behind it.

    Events are the fast path, not the only path - see SCAN_DEBOUNCE.
    """

    def __init__(self, work_queue):
        self.work_queue = work_queue

    def on_any_event(self, event):
        # Arm the debounced scan on ANY change in the folder, not just the
        # ones that look like a PO. A browser writing its temp file or
        # creating a placeholder is evidence that downloads are happening,
        # so if one PO's notification was dropped, a neighbouring event
        # still schedules the scan that finds it.
        note_activity()

    def on_created(self, event):
        self._maybe_handle(event.src_path, event.is_directory)

    def on_moved(self, event):
        # Chrome (and some other browsers) download to a temp name like
        # "PO Viewer.crdownload" and rename it to the final name on
        # completion - that shows up as a "moved" event, not "created".
        self._maybe_handle(event.dest_path, event.is_directory)

    def _maybe_handle(self, raw_path, is_directory):
        if is_directory:
            return
        path = Path(raw_path)
        if is_candidate(path):
            enqueue_file(self.work_queue, path)


def worker_loop(work_queue, config_loader, stop_event):
    """
    Processes queued PDFs off the watchdog dispatch thread.

    Renames run one at a time, so two POs can never race for the same target
    filename. A download that isn't finished yet is put back on the queue
    rather than waited on, so a slow or abandoned file can't hold up the POs
    behind it.
    """
    while not stop_event.is_set():
        try:
            path, deadline = work_queue.get(timeout=0.5)
        except queue.Empty:
            continue
        try:
            if not path.exists():
                forget_pending(path)
                release_file(path)
                continue
            if is_download_complete(path):
                forget_pending(path)
                release_file(path)
                process_file(path, config_loader.get())
            elif time.time() < deadline:
                time.sleep(0.2)
                work_queue.put((path, deadline))   # stays in flight
            else:
                forget_pending(path)
                try:
                    stat = path.stat()
                    release_file(path, (stat.st_size, stat.st_mtime))
                except OSError:
                    release_file(path)
                log(f"Gave up waiting for '{path.name}' to finish downloading.",
                    logging.WARNING)
        except Exception as exc:
            release_file(path)
            log(f"Unexpected error processing {path.name}: {exc}", logging.ERROR)
        finally:
            work_queue.task_done()


# How long to keep waiting at startup for a folder that isn't there yet.
#
# Launched at log on, this can start before the rest of the desktop is
# ready. The tool folder in particular often lives in a synced location
# (OneDrive) that hasn't been mounted yet, and giving up at that moment is
# why a start would occasionally do nothing at all.
STARTUP_WAIT = 180.0

# How long to wait before trying the folder watch again when restarting it
# failed, so a folder that is temporarily gone doesn't produce a error a
# second until it comes back.
WATCH_RETRY_BACKOFF = 30.0


def wait_for_path(path, description, timeout=STARTUP_WAIT):
    """Waits for a path to appear. True if it did, False if it never came."""
    deadline = time.time() + timeout
    announced = False
    while not path.exists():
        if time.time() >= deadline:
            return False
        if not announced:
            # Logging may itself be unavailable right now for the same
            # reason - the handler retries on the next record, so this line
            # appears as soon as the folder is back.
            log(f"Waiting for {description} to become available: {path}")
            announced = True
        time.sleep(2.0)
    return True


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--file",
        help="Reprocess a single existing PDF (e.g. an ERROR file after you've "
             "fixed the config table) instead of starting the folder watcher.",
    )
    parser.add_argument(
        "--log",
        nargs="?",
        type=int,
        const=40,
        metavar="LINES",
        help="Show where the log lives and print its last LINES lines "
             "(default 40), then exit.",
    )
    args = parser.parse_args()

    if args.log is not None:
        print(f"Log file: {LOG_PATH}")
        print(f"Keeping {LOG_RETENTION_DAYS} days of history in: {LOG_DIR}")
        if not LOG_PATH.exists():
            print("(nothing logged yet)")
            return
        lines = LOG_PATH.read_text(encoding="utf-8", errors="replace").splitlines()
        print(f"--- last {min(args.log, len(lines))} of {len(lines)} lines ---")
        for line in lines[-args.log:]:
            print(line)
        return

    if not wait_for_path(CONFIG_PATH, "the config file"):
        log(f"Config file never appeared: {CONFIG_PATH}", logging.ERROR)
        sys.exit(1)

    config_loader = ConfigLoader(CONFIG_PATH)
    config_loader.get()  # load once up front so problems surface at startup

    if args.file:
        # Manual retry mode: process one file and exit. Doesn't require the
        # filename to contain "PO Viewer" since you're pointing at it directly.
        target = Path(args.file)
        if not target.exists():
            log(f"File not found: {target}", logging.ERROR)
            sys.exit(1)
        process_file(target, config_loader.get())
        return

    if not wait_for_path(DOWNLOADS_DIR, "the Downloads folder"):
        log(f"Downloads folder never appeared: {DOWNLOADS_DIR}", logging.ERROR)
        sys.exit(1)

    log(f"Watching {DOWNLOADS_DIR} for files containing '{TRIGGER_TEXT}' "
        f"(re-scanning {SCAN_DEBOUNCE:.0f}s after downloads settle) ...")
    log(f"Logging to {LOG_PATH} (keeping {LOG_RETENTION_DAYS} days).")

    work_queue = queue.Queue()
    stop_event = threading.Event()
    worker = threading.Thread(
        target=worker_loop,
        args=(work_queue, config_loader, stop_event),
        daemon=True,
    )
    worker.start()

    observer = Observer()
    observer.schedule(PoViewerHandler(work_queue), str(DOWNLOADS_DIR), recursive=False)
    observer.start()

    # Catch anything already sitting there - a PO downloaded while the
    # watcher was stopped, or one whose event went missing on a previous run.
    scan_folder(work_queue, DOWNLOADS_DIR)
    scanner = threading.Thread(
        target=scan_loop,
        args=(work_queue, DOWNLOADS_DIR, stop_event),
        daemon=True,
    )
    scanner.start()

    retry_watch_at = 0.0
    try:
        while True:
            time.sleep(1)

            # Everything above runs on background threads, and a thread that
            # dies takes its job with it silently - the process stays up,
            # looking healthy, doing nothing. Notice that and put it back.
            if not observer.is_alive() and time.time() >= retry_watch_at:
                log("Folder watch stopped; restarting it.", logging.WARNING)
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass
                try:
                    observer = Observer()
                    observer.schedule(PoViewerHandler(work_queue),
                                      str(DOWNLOADS_DIR), recursive=False)
                    observer.start()
                    # Pick up anything that landed while it was down.
                    scan_folder(work_queue, DOWNLOADS_DIR)
                    retry_watch_at = 0.0
                    log("Folder watch restarted.")
                except Exception as exc:
                    # Can't restart yet - the folder may be briefly gone.
                    # Back off rather than retrying, and logging, every second.
                    retry_watch_at = time.time() + WATCH_RETRY_BACKOFF
                    log(f"Could not restart the folder watch ({exc}); "
                        f"trying again in {WATCH_RETRY_BACKOFF:.0f}s.",
                        logging.ERROR)

            if not worker.is_alive():
                log("Rename worker stopped; restarting it.", logging.WARNING)
                worker = threading.Thread(
                    target=worker_loop,
                    args=(work_queue, config_loader, stop_event),
                    daemon=True,
                )
                worker.start()

            if not scanner.is_alive():
                log("Folder scan stopped; restarting it.", logging.WARNING)
                scanner = threading.Thread(
                    target=scan_loop,
                    args=(work_queue, DOWNLOADS_DIR, stop_event),
                    daemon=True,
                )
                scanner.start()
    except KeyboardInterrupt:
        log("Stopping (interrupted).")
    finally:
        observer.stop()
        observer.join()
        stop_event.set()
        note_activity()          # wake the scan thread so it can exit
        worker.join(timeout=5)
        scanner.join(timeout=5)
        log("Watcher stopped.")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except BaseException as exc:
        # Under pythonw.exe there is no console, so without this an
        # unexpected exit leaves nothing at all to explain why the watcher
        # disappeared. Record it before going.
        log(f"Stopped unexpectedly: {exc!r}\n{traceback.format_exc()}",
            logging.ERROR)
        raise
