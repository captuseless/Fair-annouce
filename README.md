# Fair Announcement Scheduler

Automatically plays audio files at scheduled times. Mutes all other system audio while an announcement plays, then restores it when finished.

> **Windows only** — audio control relies on Windows Core Audio (pycaw).

---

## Files

| File | Purpose |
|---|---|
| `xlsxBuilder.py` | GUI to build the Excel schedule file |
| `fairAnnouce2.py` | Runs the schedule and plays audio |
| `listBuilder.py` | Legacy GUI that builds CSV schedule files |
| `fairAnnouce.py` | Legacy scheduler that reads CSV files |

---

## Setup

**1. Install Python 3.12+** from [python.org](https://python.org)

**2. Create and activate a virtual environment**

```bat
python -m venv venv
venv\Scripts\activate
```

**3. Install dependencies**

```bat
pip install -r requirements.txt
```

---

## Building a Schedule (xlsxBuilder.py)

Run the GUI to create or edit an Excel schedule file:

```bat
python xlsxBuilder.py
```

**Steps:**

1. Click **Open / New XLSX** — either open an existing `.xlsx` file or choose a location to create a new one.
2. Click **Pick Date & Time** — select a date from the calendar and set the hour, minute, and AM/PM.
3. Click **Browse** — select the audio file to play (mp3, wav, ogg, flac, m4a).
4. Click **Add to Schedule** — the entry appears in the table.
5. Repeat steps 2–4 for each announcement.
6. Click **Save Excel File** — writes all entries to disk.

To remove an entry, click its row in the table then click **Delete Selected**, and save again.

The Excel file uses two columns:

| datetime | file_path |
|---|---|
| `MM/DD/YYYY HH:MM` | Full path to audio file |

---

## Running the Scheduler (fairAnnouce2.py)

**Option A — pass the path on the command line (overrides the default)**

```bat
python fairAnnouce2.py "C:\Schedules\fair2026.xlsx"
```

**Option B — hardcode a default path**

Open `fairAnnouce2.py` and update the `DEFAULT_SCHEDULE` constant near the top:

```python
DEFAULT_SCHEDULE = 'C:\\Schedules\\fair2026.xlsx'
```

Then run without arguments and it will use that path automatically:

```bat
python fairAnnouce2.py
```

**Get help / see usage**

```bat
python fairAnnouce2.py --help
```

The script will:
- Wait until each scheduled time arrives
- Mute all audio except the media player playing the announcement
- Unmute everything once the audio finishes
- Move on to the next entry

Entries whose scheduled time has already passed are skipped automatically.

> Only one instance of `fairAnnouce2.py` can run at a time. A `script.lock` file prevents duplicates and is removed automatically when the script exits.

---

## Supported Audio Formats

mp3, wav, ogg, flac, m4a — any format supported by your system's default media player.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` with the venv active |
| `openpyxl` not found | Run `pip install openpyxl` |
| Audio not muting | Run the script as Administrator |
| `script.lock` exists on startup | A previous run did not exit cleanly — delete `script.lock` manually |
| Entry skipped immediately | The scheduled time is in the past — update the date/time in xlsxBuilder |
