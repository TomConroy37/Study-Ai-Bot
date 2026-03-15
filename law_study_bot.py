#!/usr/bin/env python3
"""
Law Study Bot — watches Google Drive folders for lecture files, processes them
through the Claude API, and writes structured notes into Notion databases.
"""

import argparse
import io
import json
import logging
import os
import sys
import time
from datetime import date, datetime
from pathlib import Path

import anthropic
import fitz  # PyMuPDF
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from notion_client import Client as NotionClient
from docx import Document
from pptx import Presentation

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

load_dotenv()
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("law_study_bot")

PROCESSED_FILES_PATH = Path("processed_files.json")
POLL_INTERVAL_SECONDS = 300  # 5 minutes

# Map env-var names → human subject labels → Notion DB env-var names
SUBJECT_CONFIG = {
    "DRIVE_FOLDER_PROPERTY": {
        "label": "Property Law",
        "notion_db_env": "NOTION_DB_PROPERTY",
    },
    "DRIVE_FOLDER_CORPS": {
        "label": "Business Organisations",
        "notion_db_env": "NOTION_DB_CORPS",
    },
    "DRIVE_FOLDER_EES": {
        "label": "Enhancing Employability Skills",
        "notion_db_env": "NOTION_DB_EES",
    },
    "DRIVE_FOLDER_IBO": {
        "label": "International Business Operations",
        "notion_db_env": "NOTION_DB_IBO",
    },
}

CLAUDE_MODEL = "claude-opus-4-6"  # Most capable model; user can override via env

SYSTEM_PROMPT = (
    "You are a law school study assistant. Extract structured notes strictly from the lecture content provided. "
    "STRICT ACCURACY RULES — these are non-negotiable:\n"
    "1. Only include cases that are explicitly named in the lecture text. Do not add, infer, or recall cases from your training data.\n"
    "2. Only include legislation that is explicitly referenced in the lecture text. Do not add related or commonly associated statutes.\n"
    "3. Case citations must be copied exactly as they appear in the lecture. If a citation is not given in the lecture, omit the citation field rather than guessing.\n"
    "4. If a field (e.g. cases, legislation) has no content in the lecture, return an empty array — never fabricate entries to fill it.\n"
    "5. Summaries, key concepts, exam notes, and tutorial prep must be grounded in the lecture content only — do not add external legal commentary.\n"
    "Write tutorial_prep in plain conversational English — short spoken phrases, not formal legal prose. "
    "Format any citations that ARE present in the lecture in AGLC4 format."
)

NOTE_SCHEMA = """{
  "subject": "Property Law",
  "topic": "Native Title",
  "summary": "3-4 sentence overview of the lecture",
  "key_concepts": ["concept 1", "concept 2"],
  "cases": [
    {
      "name": "Mabo v Queensland (No 2)",
      "citation": "[1992] HCA 23",
      "principle": "one sentence legal principle",
      "relevance": "why it matters to this topic"
    }
  ],
  "legislation": ["s 223 Native Title Act 1993 (Cth)"],
  "exam_notes": "key points likely to appear in exams",
  "tutorial_prep": "conversational dot points for tutorial discussion — plain English, not formal legal prose"
}"""

# ---------------------------------------------------------------------------
# Persistence helpers
# ---------------------------------------------------------------------------


def load_processed() -> set[str]:
    """Return the set of Drive file IDs already processed."""
    if PROCESSED_FILES_PATH.exists():
        return set(json.loads(PROCESSED_FILES_PATH.read_text()))
    return set()


def save_processed(processed: set[str]) -> None:
    PROCESSED_FILES_PATH.write_text(json.dumps(sorted(processed), indent=2))


# ---------------------------------------------------------------------------
# Google Drive helpers
# ---------------------------------------------------------------------------


def build_drive_service():
    sa_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_path:
        log.error("GOOGLE_SERVICE_ACCOUNT_JSON not set in .env")
        sys.exit(1)
    creds = service_account.Credentials.from_service_account_file(
        sa_path, scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def list_new_files(drive_service, folder_id: str, processed: set[str]) -> list[dict]:
    """List PPTX/PDF files in *folder_id* that haven't been processed yet."""
    query = (
        f"'{folder_id}' in parents "
        "and (mimeType='application/vnd.openxmlformats-officedocument.presentationml.presentation' "
        "or mimeType='application/pdf' "
        "or mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document') "
        "and trashed=false"
    )
    results = (
        drive_service.files()
        .list(q=query, fields="files(id, name, mimeType)")
        .execute()
    )
    files = results.get("files", [])
    return [f for f in files if f["id"] not in processed]


def download_file(drive_service, file_id: str) -> bytes:
    """Download a Drive file and return its raw bytes."""
    request = drive_service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Text extraction
# ---------------------------------------------------------------------------


def extract_text_pptx(data: bytes) -> str:
    """Extract slide text + speaker notes from a PPTX file."""
    prs = Presentation(io.BytesIO(data))
    parts: list[str] = []
    for slide_num, slide in enumerate(prs.slides, start=1):
        slide_texts: list[str] = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if text:
                        slide_texts.append(text)
        notes = ""
        if slide.has_notes_slide:
            notes_tf = slide.notes_slide.notes_text_frame
            notes = notes_tf.text.strip() if notes_tf else ""
        parts.append(f"--- Slide {slide_num} ---")
        parts.extend(slide_texts)
        if notes:
            parts.append(f"[Speaker notes: {notes}]")
    return "\n".join(parts)


def extract_text_docx(data: bytes) -> str:
    """Extract text from a Word document, preserving heading/paragraph structure."""
    doc = Document(io.BytesIO(data))
    lines: list[str] = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            lines.append(text)
    return "\n".join(lines)


def extract_text_pdf(data: bytes) -> str:
    """Extract text from a PDF using PyMuPDF."""
    doc = fitz.open(stream=data, filetype="pdf")
    pages: list[str] = []
    for page_num, page in enumerate(doc, start=1):
        text = page.get_text().strip()
        if text:
            pages.append(f"--- Page {page_num} ---\n{text}")
    doc.close()
    return "\n\n".join(pages)


def extract_text(filename: str, data: bytes) -> str:
    ext = Path(filename).suffix.lower()
    if ext == ".pptx":
        return extract_text_pptx(data)
    elif ext == ".pdf":
        return extract_text_pdf(data)
    elif ext == ".docx":
        return extract_text_docx(data)
    raise ValueError(f"Unsupported file type: {ext}")


# ---------------------------------------------------------------------------
# Claude processing
# ---------------------------------------------------------------------------


def process_with_claude(
    claude_client: anthropic.Anthropic, subject_label: str, filename: str, text: str
) -> dict:
    """Send lecture text to Claude and return the parsed JSON notes."""
    user_message = (
        f"Subject: {subject_label}\n"
        f"File: {filename}\n\n"
        f"Lecture content (your ONLY source — do not use knowledge outside this text):\n{text}\n\n"
        f"Return ONLY a valid JSON object matching this schema (no markdown fences).\n"
        f"Include only cases and legislation that appear explicitly in the lecture content above:\n{NOTE_SCHEMA}"
    )

    log.info("  Calling Claude API…")
    with claude_client.messages.stream(
        model=CLAUDE_MODEL,
        max_tokens=4096,
        thinking={"type": "adaptive"},
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_message}],
    ) as stream:
        response = stream.get_final_message()

    # Extract the text block (thinking blocks come first; skip them)
    raw_text = ""
    for block in response.content:
        if block.type == "text":
            raw_text = block.text.strip()
            break

    # Strip accidental markdown fences
    if raw_text.startswith("```"):
        raw_text = raw_text.split("```")[1]
        if raw_text.startswith("json"):
            raw_text = raw_text[4:]

    return json.loads(raw_text)


# ---------------------------------------------------------------------------
# Notion helpers
# ---------------------------------------------------------------------------


def rich_text(content: str) -> list[dict]:
    """Wrap a plain string in a Notion rich_text array (respecting 2000-char limit)."""
    chunks: list[dict] = []
    for i in range(0, len(content), 2000):
        chunks.append({"type": "text", "text": {"content": content[i : i + 2000]}})
    return chunks


def bullet_items(items: list[str]) -> list[dict]:
    """Build a list of Notion bulleted_list_item blocks."""
    return [
        {
            "object": "block",
            "type": "bulleted_list_item",
            "bulleted_list_item": {
                "rich_text": rich_text(item),
            },
        }
        for item in items
    ]


def heading2(text: str) -> dict:
    return {
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": rich_text(text)},
    }


def paragraph(text: str) -> dict:
    return {
        "object": "block",
        "type": "paragraph",
        "paragraph": {"rich_text": rich_text(text)},
    }


def append_blocks_chunked(
    notion: NotionClient, page_id: str, blocks: list[dict]
) -> None:
    """Notion limits appending to 100 blocks per call; chunk as needed."""
    for i in range(0, len(blocks), 100):
        notion.blocks.children.append(block_id=page_id, children=blocks[i : i + 100])


def write_subject_page(
    notion: NotionClient, db_id: str, notes: dict, processed_date: str
) -> None:
    """Create a lecture notes page in the subject-specific Notion database."""
    title = f"{notes['topic']} — {processed_date}"
    log.info(f"  Creating subject page: «{title}»")

    page = notion.pages.create(
        parent={"database_id": db_id},
        properties={
            "Name": {"title": rich_text(title)},
            "Subject": {
                "rich_text": rich_text(notes.get("subject", ""))
            },
            "Topic": {
                "rich_text": rich_text(notes.get("topic", ""))
            },
            "Date Processed": {
                "rich_text": rich_text(processed_date)
            },
        },
    )
    page_id = page["id"]

    blocks: list[dict] = [
        heading2("Summary"),
        paragraph(notes.get("summary", "")),
        heading2("Key Concepts"),
    ]
    blocks += bullet_items(notes.get("key_concepts", []))
    blocks.append(heading2("Legislation"))
    blocks += bullet_items(notes.get("legislation", []))
    blocks.append(heading2("Exam Notes"))
    blocks.append(paragraph(notes.get("exam_notes", "")))
    blocks.append(heading2("Tutorial Prep"))
    blocks.append(paragraph(notes.get("tutorial_prep", "")))

    append_blocks_chunked(notion, page_id, blocks)


def write_case_bank_entries(
    notion: NotionClient, case_bank_db_id: str, notes: dict
) -> None:
    """Create or update one page per case in the Case Bank database."""
    for case in notes.get("cases", []):
        case_name = case.get("name", "Unknown case")
        log.info(f"  Writing case: {case_name}")

        # Check if case already exists (simple search by title prefix)
        search = notion.databases.query(
            database_id=case_bank_db_id,
            filter={
                "property": "Name",
                "title": {"equals": case_name},
            },
        )
        existing_pages = search.get("results", [])

        properties = {
            "Name": {"title": rich_text(case_name)},
            "Citation": {"rich_text": rich_text(case.get("citation", ""))},
            "Subject": {"rich_text": rich_text(notes.get("subject", ""))},
            "Topic": {"rich_text": rich_text(notes.get("topic", ""))},
        }
        body_blocks: list[dict] = [
            heading2("Legal Principle"),
            paragraph(case.get("principle", "")),
            heading2("Relevance"),
            paragraph(case.get("relevance", "")),
        ]

        if existing_pages:
            page_id = existing_pages[0]["id"]
            notion.pages.update(page_id=page_id, properties=properties)
            # Clear old body and rewrite
            existing_blocks = notion.blocks.children.list(block_id=page_id).get(
                "results", []
            )
            for block in existing_blocks:
                try:
                    notion.blocks.delete(block_id=block["id"])
                except Exception:
                    pass
            append_blocks_chunked(notion, page_id, body_blocks)
        else:
            page = notion.pages.create(
                parent={"database_id": case_bank_db_id},
                properties=properties,
            )
            append_blocks_chunked(notion, page["id"], body_blocks)


def write_revision_page(
    notion: NotionClient, revision_db_id: str, notes: dict, processed_date: str
) -> None:
    """Create a revision page in NOTION_DB_REVISION."""
    title = f"{notes['topic']} — Revision"
    log.info(f"  Creating revision page: «{title}»")

    page = notion.pages.create(
        parent={"database_id": revision_db_id},
        properties={
            "Name": {"title": rich_text(title)},
        },
    )
    append_blocks_chunked(
        notion,
        page["id"],
        [
            heading2("Exam Notes"),
            paragraph(notes.get("exam_notes", "")),
        ],
    )


# ---------------------------------------------------------------------------
# Core processing loop
# ---------------------------------------------------------------------------


def process_folder(
    drive_service,
    claude_client: anthropic.Anthropic,
    notion: NotionClient,
    folder_env_key: str,
    config: dict,
    processed: set[str],
) -> set[str]:
    """Process all new files in one Drive folder. Returns updated processed set."""
    folder_id = os.environ.get(folder_env_key, "").strip()
    if not folder_id:
        log.warning(f"{folder_env_key} not set — skipping.")
        return processed

    subject_label = config["label"]
    notion_db_id = os.environ.get(config["notion_db_env"], "").strip()
    if not notion_db_id:
        log.warning(f"{config['notion_db_env']} not set — skipping {subject_label}.")
        return processed

    case_bank_db_id = os.environ.get("NOTION_DB_CASE_BANK", "").strip()
    revision_db_id = os.environ.get("NOTION_DB_REVISION", "").strip()

    new_files = list_new_files(drive_service, folder_id, processed)
    if not new_files:
        log.info(f"[{subject_label}] No new files.")
        return processed

    for file_info in new_files:
        filename = file_info["name"]
        file_id = file_info["id"]
        log.info(f"Found new file: {filename}")

        # --- Extract text ---
        try:
            log.info("  Downloading…")
            data = download_file(drive_service, file_id)
            log.info("  Extracting text…")
            text = extract_text(filename, data)
        except Exception as exc:
            log.error(f"  Extraction failed for {filename}: {exc}")
            continue  # Don't mark as processed; will retry next run

        # --- Claude ---
        try:
            log.info("  Processing with Claude…")
            notes = process_with_claude(claude_client, subject_label, filename, text)
        except Exception as exc:
            log.error(f"  Claude API failed for {filename}: {exc}")
            continue  # Don't mark as processed; will retry next run

        topic = notes.get("topic", "Unknown Topic")
        processed_date = date.today().isoformat()

        # --- Notion writes ---
        try:
            write_subject_page(notion, notion_db_id, notes, processed_date)
        except Exception as exc:
            log.error(f"  Notion subject page write failed: {exc}")

        if case_bank_db_id:
            try:
                write_case_bank_entries(notion, case_bank_db_id, notes)
            except Exception as exc:
                log.error(f"  Case bank write failed: {exc}")

        if revision_db_id:
            try:
                write_revision_page(notion, revision_db_id, notes, processed_date)
            except Exception as exc:
                log.error(f"  Revision page write failed: {exc}")

        log.info(f"Written to Notion: {topic}")

        # Mark as processed only after all writes attempted
        processed.add(file_id)
        save_processed(processed)

    return processed


def run_once() -> None:
    """Process any new files across all watched folders, then exit."""
    drive_service = build_drive_service()
    claude_client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    notion = NotionClient(auth=os.environ.get("NOTION_API_KEY"))
    processed = load_processed()

    for folder_env_key, config in SUBJECT_CONFIG.items():
        processed = process_folder(
            drive_service, claude_client, notion, folder_env_key, config, processed
        )

    log.info("Run complete.")


def watch_loop() -> None:
    """Poll all Drive folders every POLL_INTERVAL_SECONDS until interrupted."""
    drive_service = build_drive_service()
    claude_client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    notion = NotionClient(auth=os.environ.get("NOTION_API_KEY"))

    log.info(f"Watching folders — polling every {POLL_INTERVAL_SECONDS // 60} minutes.")
    while True:
        processed = load_processed()
        for folder_env_key, config in SUBJECT_CONFIG.items():
            processed = process_folder(
                drive_service, claude_client, notion, folder_env_key, config, processed
            )
        log.info(f"Sleeping {POLL_INTERVAL_SECONDS // 60} min…")
        time.sleep(POLL_INTERVAL_SECONDS)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Law Study Bot — auto-process lecture files from Google Drive into Notion."
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--watch",
        action="store_true",
        help="Poll Drive folders every 5 minutes in a loop.",
    )
    group.add_argument(
        "--run-once",
        action="store_true",
        help="Process new files once and exit (good for cron).",
    )
    args = parser.parse_args()

    if args.watch:
        watch_loop()
    else:
        run_once()


if __name__ == "__main__":
    main()
