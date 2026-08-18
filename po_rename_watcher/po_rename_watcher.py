"""
PO Viewer auto-renamer for Arise Homes.

Watches the Windows Downloads folder. When a PDF whose filename contains
"PO Viewer" appears, it reads fixed areas of page 1, looks up the vendor /
subdivision / PO type in po_rename_config.txt, and renames the file in place
to:

    [VendorShort]_[SubdivisionAbbrev][Lot]_[Address]_[PO Type]_[DD Mon YYYY].pdf

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


def _default_log_dir():
    """
    Somewhere local to write the log.

    Deliberately NOT next to the script: this tool typically lives in a
    OneDrive-synced folder, and rewriting the log on every rename would make
    OneDrive re-upload it constantly. %LOCALAPPDATA% is machine-local, so
    the churn stays off the sync engine.
    """
    base = os.environ.get("LOCALAPPDATA")
    if base:
        candidate = Path(base) / "PoRenameWatcher"
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            return candidate
        except OSError:
            pass
    return SCRIPT_DIR


LOG_DIR = _default_log_dir()
LOG_PATH = LOG_DIR / "po_rename_log.txt"

TRIGGER_TEXT = "po viewer"          # filename must contain this (case-insensitive)
ERROR_STEM = "ERROR - PO Viewer"    # base name used when a file can't be processed

# Right-column x-position cutoff (points from left edge of the page) used to
# separate the Vendor block (left column) from the Ship To / summary-table
# block (right column). This matches the ERP's fixed PO Viewer template.
RIGHT_COLUMN_X_MIN = 280


class SafeTimedRotatingFileHandler(TimedRotatingFileHandler):
    """
    Daily-rotating log handler that won't take the watcher down if the log
    file is momentarily locked.

    On Windows a second process - e.g. a `--file` retry run from Command
    Prompt while the background watcher is up - can hold the log open and
    make the rollover rename fail. Rather than raising, keep appending to
    the current file and retry the rollover a little later.
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


def extract_raw_fields(pdf_path):
    """
    Returns a dict with vendor_raw, po_type_raw, address_raw, plot_raw.
    Any field that can't be located is None.
    """
    result = {"vendor_raw": None, "po_type_raw": None, "address_raw": None, "plot_raw": None}

    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            return result
        page = pdf.pages[0]
        left_lines = get_visual_lines(page, x_max=RIGHT_COLUMN_X_MIN)
        right_lines = get_visual_lines(page, x_min=RIGHT_COLUMN_X_MIN)

        # --- PO Type: right-hand summary table ---
        # Restricted to the right column (and matched anywhere in the line
        # rather than anchored at its start), so left-column content sitting
        # at the same height can't push the label out of position.
        for _, text in right_lines:
            m = re.search(r"PO Type:\s*(.+)$", text.strip(), re.IGNORECASE)
            if m:
                result["po_type_raw"] = m.group(1).strip()
                break

        # --- Vendor: left column, first line under the "Vendor:" label ---
        vendor_label_top = None
        for top, text in left_lines:
            if text.strip().rstrip(":").lower() == "vendor":
                vendor_label_top = top
                break
        if vendor_label_top is not None:
            below = sorted(
                (t for t in left_lines if t[0] > vendor_label_top + 1),
                key=lambda t: t[0],
            )
            if below and (below[0][0] - vendor_label_top) < 40:
                result["vendor_raw"] = below[0][1].strip()

        # --- Ship To block: right column, address + plot/lot lines ---
        shipto_label_top = None
        for top, text in right_lines:
            if "ship to" in text.strip().lower():
                shipto_label_top = top
                break
        if shipto_label_top is not None:
            block = sorted(
                (t for t in right_lines if 0 < (t[0] - shipto_label_top) < 100),
                key=lambda t: t[0],
            )
            for i, (top, text) in enumerate(block):
                if re.search(r"\bLot\s*\d+", text, re.IGNORECASE):
                    result["plot_raw"] = text.strip()
                    if i > 0:
                        result["address_raw"] = block[i - 1][1].strip()
                    break

    return result


# ----------------------------------------------------------------------
# Lookup + filename construction
# ----------------------------------------------------------------------

def resolve_vendor(vendor_raw, vendor_table):
    if not vendor_raw:
        return None
    # "Century Building Solutions - Spring Hill" -> "Century Building Solutions"
    truncated = normalize_ws(vendor_raw.split(" - ")[0])
    return vendor_table.get(truncated.lower())


def resolve_subdivision_and_lot(plot_raw, subdivision_table):
    if not plot_raw:
        return None, None
    text_lower = plot_raw.lower()
    matches = [(full, abbr) for full, abbr in subdivision_table if full in text_lower]
    if len(matches) != 1:
        return None, None
    _, abbr = matches[0]
    lot_match = re.search(r"\bLot\s*(\d+[A-Za-z]?)", plot_raw, re.IGNORECASE)
    if not lot_match:
        return None, None
    return abbr, lot_match.group(1)


def resolve_po_type(po_type_raw, po_type_table):
    if not po_type_raw:
        return None
    return po_type_table.get(normalize_ws(po_type_raw).lower())


# A street address starts with a house number (optionally suffixed, e.g. "123A")
# followed by the street name. The address is picked positionally - the line
# directly above the Plat/Lot line - so this guards against grabbing a
# neighbouring line (e.g. "Arise Homes LLC") when the Ship To block's shape
# differs from the usual template.
STREET_ADDRESS_RE = re.compile(r"^\d+[A-Za-z]?\s+\S")


def looks_like_street_address(s):
    return bool(STREET_ADDRESS_RE.match(s or ""))


def sanitize_part(s):
    s = re.sub(r'[\\/:*?"<>|]', "-", s)
    return s.strip(" .")


def build_target_filename(vendor_short, subdivision_abbr, lot, address, po_type_value, date_str):
    parts = [
        vendor_short,
        f"{subdivision_abbr}{lot}",
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

# path -> (last observed size, when it first reached that size).
# Only touched by the worker thread.
_pending_sizes = {}


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
    if looks_like_complete_pdf(path):
        _pending_sizes.pop(path, None)
        return True

    now = time.time()
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
    subdivision_abbr, lot = resolve_subdivision_and_lot(fields["plot_raw"], config["subdivisions"])
    po_type_value = resolve_po_type(fields["po_type_raw"], config["po_types"])
    address = normalize_ws(fields["address_raw"]) if fields["address_raw"] else None

    missing = []
    if not vendor_short:
        missing.append(f"vendor (read: {fields['vendor_raw']!r})")
    if not subdivision_abbr or not lot:
        missing.append(f"subdivision/lot (read: {fields['plot_raw']!r})")
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
    target_name = build_target_filename(vendor_short, subdivision_abbr, lot, address, po_type_value, date_str)
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

class PoViewerHandler(FileSystemEventHandler):
    """
    Filters filesystem events and hands matching PDFs to a worker thread.

    The actual processing deliberately does NOT happen here: watchdog
    dispatches events on a single thread, so waiting on a slow or stalled
    download inline would hold up every PO queued behind it.
    """

    def __init__(self, work_queue):
        self.work_queue = work_queue

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
        if path.suffix.lower() != ".pdf":
            return
        stem_lower = path.stem.lower()
        if stem_lower.startswith(ERROR_STEM.lower()):
            # Never re-trigger on our own error output - renaming a file
            # inside the watched folder fires a "moved" event, and an
            # ERROR file's name still contains "po viewer", which would
            # otherwise cause it to reprocess itself in a loop.
            return
        if TRIGGER_TEXT not in stem_lower:
            return
        self.work_queue.put((path, time.time() + DOWNLOAD_TIMEOUT))


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
                _pending_sizes.pop(path, None)
                continue
            if is_download_complete(path):
                _pending_sizes.pop(path, None)
                process_file(path, config_loader.get())
            elif time.time() < deadline:
                time.sleep(0.2)
                work_queue.put((path, deadline))
            else:
                _pending_sizes.pop(path, None)
                log(f"Gave up waiting for '{path.name}' to finish downloading.",
                    logging.WARNING)
        except Exception as exc:
            log(f"Unexpected error processing {path.name}: {exc}", logging.ERROR)
        finally:
            work_queue.task_done()


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

    if not CONFIG_PATH.exists():
        log(f"Config file not found: {CONFIG_PATH}", logging.ERROR)
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

    if not DOWNLOADS_DIR.exists():
        log(f"Downloads folder not found: {DOWNLOADS_DIR}", logging.ERROR)
        sys.exit(1)

    log(f"Watching {DOWNLOADS_DIR} for files containing '{TRIGGER_TEXT}' ...")
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
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    finally:
        observer.stop()
        observer.join()
        stop_event.set()
        worker.join(timeout=5)


if __name__ == "__main__":
    main()
