# Law Study Bot

Automatically watches your Google Drive lecture folders, processes new PPTX/PDF files through Claude, and pushes structured study notes into your Notion workspace.

---

## Features

- Watches 4 subject-specific Google Drive folders for new PPTX and PDF files
- Extracts slide text + speaker notes (PPTX) or page text (PDF)
- Sends content to Claude (Opus 4.6 with adaptive thinking) and receives structured JSON notes
- Creates pages in subject-specific Notion databases with Summary, Key Concepts, Legislation, Exam Notes, and Tutorial Prep
- Populates a Case Bank database with AGLC4-cited case entries
- Creates Revision pages for quick exam prep
- Tracks processed files locally (`processed_files.json`) so nothing is double-processed
- Two modes: `--watch` (poll every 5 min) or `--run-once` (great for cron)

---

## Prerequisites

- Python 3.11+
- An [Anthropic API key](https://console.anthropic.com/)
- A Google Cloud service account with Drive access
- A Notion integration token and shared databases

---

## Installation

```bash
git clone <this-repo>
cd Study-Ai-Bot
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

---

## Configuration

Create a `.env` file in the project root:

```env
# Anthropic
ANTHROPIC_API_KEY=sk-ant-...

# Google Drive — service account JSON file path
GOOGLE_SERVICE_ACCOUNT_JSON=path/to/service_account.json

# Google Drive folder IDs (from the URL: drive.google.com/drive/folders/<ID>)
DRIVE_FOLDER_PROPERTY=
DRIVE_FOLDER_CORPS=
DRIVE_FOLDER_EES=
DRIVE_FOLDER_IBO=

# Notion integration token
NOTION_API_KEY=secret_...

# Notion database IDs (32-char hex from the database URL)
NOTION_DB_CASE_BANK=
NOTION_DB_PROPERTY=
NOTION_DB_CORPS=
NOTION_DB_EES=
NOTION_DB_IBO=
NOTION_DB_REVISION=
```

---

## Getting a Google Service Account JSON

1. Go to [Google Cloud Console](https://console.cloud.google.com/) and create or select a project.
2. Enable the **Google Drive API**: APIs & Services → Enable APIs → search "Google Drive API" → Enable.
3. Go to **IAM & Admin → Service Accounts** → Create Service Account.
   - Give it a name (e.g. `law-study-bot`).
   - Skip optional role/user steps for now.
4. Click the new service account → **Keys** tab → **Add Key → Create new key → JSON**.
   - A `.json` file downloads — this is your `GOOGLE_SERVICE_ACCOUNT_JSON`.
5. **Share each Drive folder** with the service account's email address (found in the JSON as `client_email`) with **Viewer** access.
   - Right-click the folder in Google Drive → Share → paste the email.

---

## Setting Up the Notion Integration

1. Go to [https://www.notion.so/my-integrations](https://www.notion.so/my-integrations) → **New integration**.
   - Name: `Law Study Bot`
   - Capabilities: **Read content**, **Update content**, **Insert content**
   - Copy the **Internal Integration Token** → set as `NOTION_API_KEY`.
2. For **each** Notion database (Case Bank, Property, Corps, EES, IBO, Revision):
   - Open the database page in Notion.
   - Click `···` (top-right) → **Connections** → search your integration → **Connect**.
   - Copy the database ID from the URL:
     `https://www.notion.so/<workspace>/<DATABASE_ID>?v=...`
     (32 hex characters before the `?`)

### Required Database Properties

The bot writes to these properties — add them to each database if they don't exist:

**Subject databases** (Property, Corps, EES, IBO):

| Property | Type |
|---|---|
| Name | Title |
| Subject | Text |
| Topic | Text |
| Date Processed | Text |

**Case Bank**:

| Property | Type |
|---|---|
| Name | Title |
| Citation | Text |
| Subject | Text |
| Topic | Text |

**Revision**:

| Property | Type |
|---|---|
| Name | Title |

---

## Usage

### Watch mode (runs continuously, polls every 5 minutes)

```bash
python law_study_bot.py --watch
```

### Run-once mode (process new files then exit — good for cron)

```bash
python law_study_bot.py --run-once
```

### Cron example (run every 10 minutes)

```cron
*/10 * * * * /path/to/.venv/bin/python /path/to/law_study_bot.py --run-once >> /path/to/bot.log 2>&1
```

---

## How It Works

```
Google Drive folder
      │
      │  (new .pptx / .pdf detected)
      ▼
Text extraction
  • PPTX → slide text + speaker notes  (python-pptx)
  • PDF  → page text                   (PyMuPDF)
      │
      ▼
Claude API  (claude-opus-4-6, adaptive thinking)
  System: "You are a law school study assistant…"
  Returns: JSON with subject, topic, summary, key_concepts,
           cases (AGLC4), legislation, exam_notes, tutorial_prep
      │
      ├─→ Notion subject DB  (lecture notes page)
      ├─→ Notion Case Bank   (one page per case)
      └─→ Notion Revision DB (exam notes page)
      │
      ▼
processed_files.json  (prevents re-processing)
```

---

## Error Handling

- **Claude API failure**: file is skipped and NOT marked as processed — it will be retried on the next run.
- **Notion write failure**: logged as an error but does not crash the bot.
- **Missing env vars**: the affected folder/database is skipped with a warning; other folders continue normally.

---

## File Structure

```
Study-Ai-Bot/
├── law_study_bot.py       # Main script
├── requirements.txt       # Python dependencies
├── README.md              # This file
├── .env                   # Your credentials (never commit this)
└── processed_files.json   # Auto-created; tracks processed Drive file IDs
```

---

## Dependencies

| Package | Purpose |
|---|---|
| `anthropic` | Claude API client |
| `google-api-python-client` | Google Drive API |
| `google-auth` | Service account authentication |
| `python-pptx` | PPTX text extraction |
| `PyMuPDF` | PDF text extraction |
| `notion-client` | Notion API client |
| `python-dotenv` | Load `.env` credentials |
