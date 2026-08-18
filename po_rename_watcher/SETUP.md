# PO Viewer Auto-Renamer — Windows Setup

This folder contains everything you need. It watches your Downloads folder,
and whenever a PDF containing "PO Viewer" in its name appears, it reads the
Vendor / Plot & Lot / Address / PO Type off page 1 and renames the file to:

```
[VendorShort]_[SubdivisionAbbrev][Lot]_[Address]_[PO Type]_[DD Mon YYYY].pdf
```

e.g. `Century_GR31_18460 W 195th Ter_Attic Ladder_17 Aug 2026.pdf`

If it can't confidently read a field, or a value isn't in your lookup table
yet, or two POs would produce the same filename, it renames the file to
`ERROR - PO Viewer.pdf` (or `ERROR - PO Viewer (1).pdf`, etc.) instead of
guessing — nothing is ever overwritten or silently mis-named.

### Two PO layouts are supported

Both are recognised automatically — you don't pick or configure anything,
just make sure the file has "PO Viewer" somewhere in its name:

| | ERP "PO Viewer" export | Excel PO printed to PDF |
|---|---|---|
| PO type row | `PO Type:` | `Type:` |
| Jobsite block | under a `Ship To:` heading | no heading |
| Subdivision | `GARRETT RANCH THIRD PLAT, Lot 31` | `Garrett Ranch` on its own line |
| Lot number | `Lot 31` | `Homesite 31` |

A PDF that matches neither layout is flagged as an ERROR rather than
guessed at.

Note that a vendor can be spelled differently between the two (the ERP says
`McCray Lumber`, the Excel sheet says `McCray Lumber Co`). Add a `#VENDORS`
row for each spelling you run into — both can point at the same short name.

## 1. Copy this folder to your PC

Copy the whole `po_rename_watcher` folder anywhere you like — a Documents
folder, a synced OneDrive folder, wherever you keep your scripts. Just pick a
permanent home, because the scheduled task will point at it.

It should contain:
- `po_rename_watcher.py`
- `po_rename_config.txt`
- `requirements.txt`

Throughout the rest of these instructions, **"the tool folder"** means
wherever you put it. A quick way to get its exact path later: open the folder
in File Explorer, click once in the address bar, and copy the text.

## 2. Install Python (one-time)

If you don't already have Python:
1. Go to https://www.python.org/downloads/
2. Download and run the latest Python 3 installer.
3. **Important:** on the first install screen, check the box **"Add python.exe to PATH"** before clicking Install.

## 3. Install the required libraries (one-time)

Open **Command Prompt** (search "cmd" in the Start menu) and run the
following, substituting your own tool folder path. Keep the quotes — they
matter if the path contains spaces:

```
cd /d "C:\Path\To\po_rename_watcher"
pip install -r requirements.txt
```

## 4. Test it manually first

With Command Prompt still open in that folder, run:

```
python po_rename_watcher.py
```

You should see something like:

```
Loaded config: 32 vendors, 9 subdivisions, 46 PO types.
Watching C:\Users\<you>\Downloads for files containing 'po viewer' ...
Logging to ...\po_rename_log.txt (keeping 7 days).
```

Leave it running, then download a real PO Viewer PDF from the ERP. Watch it
get picked up and renamed right there in the console. Press `Ctrl+C` to stop
it once you're satisfied it works.

## 5. Make it start automatically and run in the background

We'll use Task Scheduler so it starts automatically when you log in and runs
silently (no window) from then on.

1. Search **"Task Scheduler"** in the Start menu and open it.
2. Click **Create Task...** (right panel) — not "Create Basic Task."
3. **General tab:**
   - Name: `PO Viewer Auto-Renamer`
   - Check **"Run only when user is logged on"**
4. **Triggers tab:**
   - Click **New...**
   - Begin the task: **At log on**
   - Specific user: your account (should be pre-filled)
   - Click OK
5. **Actions tab:**
   - Click **New...**
   - Action: **Start a program**
   - Program/script: the full path to `pythonw.exe` — find it by running
     `where pythonw` in Command Prompt (typically something like
     `C:\Users\<you>\AppData\Local\Programs\Python\Python312\pythonw.exe`)
   - Add arguments: `po_rename_watcher.py`
   - Start in: your tool folder path
   - Click OK

   Paste these as plain text with **no quotes around them** — Task Scheduler
   treats each box as a literal value, not a command line, so quotes make it
   look for a file that doesn't exist. Paths containing spaces are fine
   unquoted here.
6. **Conditions tab:** uncheck "Start the task only if the computer is on AC
   power" if this is a laptop, otherwise leave defaults.
7. Click **OK** to save. It'll ask for your Windows password — enter it.

`pythonw.exe` (instead of `python.exe`) runs with no visible console window,
so it just quietly runs in the background from now on.

**To test it starts correctly:** right-click the task in Task Scheduler and
choose **Run**, then check the log (see below) — it should show
"Watching ... " within a couple seconds.

## 6. Day-to-day use

Nothing to do — download PO Viewer PDFs from the ERP as usual. They'll be
renamed in place in Downloads within a few seconds.

### Checking what happened

Every file it sees — and what it renamed it to, or why it couldn't — is
logged with a timestamp. The quickest way to read it, from Command Prompt in
the tool folder:

```
python po_rename_watcher.py --log
```

That prints the log's location and its last 40 lines. Add a number for more
(`--log 200`).

The log file is `po_rename_log.txt`, in the tool folder alongside the script
and config.

A new log file starts each day and **7 days of history is kept** — older days
are deleted automatically, so it never grows without bound. Previous days are
kept beside it as `po_rename_log.txt.2026-08-17` and so on.

### Fixing an ERROR file

1. Run `python po_rename_watcher.py --log` and find why it failed — usually a vendor,
   subdivision, or PO type that isn't in `po_rename_config.txt` yet, printed
   next to "Could not resolve."
2. Open `po_rename_config.txt` in Notepad, add the missing row under the
   right section (`#Vendors`, `#Subdivisions`, or `#PO Types`), save. The
   running watcher picks this up automatically — no restart.
3. Retry that one file without redownloading it — open Command Prompt in
   the tool folder and run:
   ```
   python po_rename_watcher.py --file "%USERPROFILE%\Downloads\ERROR - PO Viewer.pdf"
   ```
   (Adjust the filename if it's `ERROR - PO Viewer (1).pdf`, etc.)

   If it still can't be resolved, the file keeps its current ERROR name and
   the log says why — it won't get renumbered on each retry.

### Editing the lookup table

`po_rename_config.txt` is a plain text file — open it in Notepad any time.
Each line is `Full Name | Short Value` (or `Full Name_Short Value` — either
works). Add new vendors, subdivisions, or PO types as you run into them.

Just save the file — **no restart needed**. The watcher notices the change
and picks up your edit on the very next PDF it processes.

### A note on the "Vendor" match rule

The tool automatically ignores everything after " - " in the vendor name on
the PDF (e.g. "Century Building Solutions **- Spring Hill**" is matched as
just "Century Building Solutions") — so multiple branches/cities of the same
vendor all match one row in your table. Make sure your `#Vendors` entries
use the name *without* that suffix, exactly as you already have them.
