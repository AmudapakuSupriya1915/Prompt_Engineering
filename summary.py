"""Generate zero-shot and source-cited Groq summaries from Project ORION documents."""

import os
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from dotenv import load_dotenv
from groq import Groq


ROOT = Path(__file__).resolve().parent
load_dotenv(ROOT / ".env")
SOURCE_FILES = (
    "01_Messy_Project_Notes.docx",
    "02_Email_Thread_And_Chat.docx",
    "03_Vendor_And_Budget_Notes.docx",
    "04_Service_Desk_Incident_Notes.docx",
)
OUTPUT_FILE = ROOT / "summary_output.md"
MODEL_NAME = os.getenv("GROQ_MODEL", "openai/gpt-oss-20b")
WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def get_text(element):
    """Return visible text from a Word paragraph or table cell."""
    return "".join(node.text or "" for node in element.findall(".//w:t", WORD_NS)).strip()


def extract_docx(path):
    """Extract paragraphs and tables while retaining Word heading labels."""
    with zipfile.ZipFile(path) as document:
        root = ET.fromstring(document.read("word/document.xml"))

    body = root.find("w:body", WORD_NS)
    if body is None:
        raise ValueError(f"No document body found in {path.name}")

    sections = []
    current_heading = "Document start"

    for element in list(body):
        element_type = element.tag.rsplit("}", 1)[-1]
        if element_type == "p":
            text = get_text(element)
            if not text:
                continue

            style = element.find("w:pPr/w:pStyle", WORD_NS)
            style_name = (
                style.get(f"{{{WORD_NS['w']}}}val", "") if style is not None else ""
            )
            if style_name.lower().startswith(("heading", "title", "subtitle")):
                current_heading = text
                sections.append(f"\n### {text}")
            else:
                sections.append(f"[{current_heading}] {text}")
        elif element_type == "tbl":
            rows = []
            for row in element.findall("./w:tr", WORD_NS):
                cells = [get_text(cell) for cell in row.findall("./w:tc", WORD_NS)]
                if any(cells):
                    rows.append(" | ".join(cells))
            if rows:
                sections.append(f"\n[Table under {current_heading}]\n" + "\n".join(rows))

    return "\n".join(sections)


def load_source_documents():
    """Read the four source reports and label each with its filename."""
    documents = []
    for filename in SOURCE_FILES:
        path = ROOT / filename
        if not path.is_file():
            raise FileNotFoundError(f"Required source document not found: {filename}")
        documents.append(f"## Source document: {filename}\n{extract_docx(path)}")
    return "\n\n".join(documents)


def build_zero_shot_prompt(source_text):
    """Build the unstructured baseline prompt for the four reports."""
    return (
        "Summarize these documents.\n\n"
        "The following text comes from four project source documents.\n\n"
        f"{source_text}"
    )


def build_structured_prompt(zero_shot_summary, source_text):
    """Build a constrained project-brief prompt grounded in the source text."""
    return f"""You are an ORION delivery analyst. Use only the four source documents below. Do not invent facts. If dates or claims conflict, show the conflict and mark it unconfirmed. Cite each factual item with the exact source filename and passage heading shown in the text.

Create a concise, structured project brief with these sections:
1. Executive summary.
2. Phase 1 scope and Phase 2 exclusions.
3. Demo date/time and production launch status.
4. Milestones, owners, and deadlines.
5. Incidents, current status, impact, and workarounds.
6. QA results and data-quality counts, with dates for changing counts.
7. Vendor quotes, tax, approval status, and budget uncertainties.
8. Risks and unresolved decisions.
9. At least five omissions or errors in the zero-shot summary below, each supported by a source citation. If there are fewer than five errors, list source-backed omissions instead of inventing errors.
10. Three specific improvements to the prompt or output, based on the comparison.

Zero-shot summary to review:
{zero_shot_summary}

Source documents:
{source_text}
"""


def call_model(client, prompt):
    """Send one prompt to Groq and return non-empty response text."""
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
        max_completion_tokens=8192,
    )
    choice = response.choices[0]
    content = (choice.message.content or "").strip()
    if not content:
        raise RuntimeError(
            f"The model returned no text (finish reason: {choice.finish_reason})."
        )
    return content


def format_report(zero_shot_summary, structured_result):
    """Format both model responses as a Markdown report."""
    return (
        "# ORION Document Summary\n\n"
        "## Zero-Shot Summary\n\n"
        f"{zero_shot_summary}\n\n"
        "## Source-Cited Brief, Review, and Improvements\n\n"
        f"{structured_result}\n"
    )


def main():
    """Extract the source reports, call Groq twice, and save the Markdown report."""
    source_text = load_source_documents()
    api_key = os.getenv("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError("GROQ_API_KEY is missing. Add it to the project .env file.")

    client = Groq(api_key=api_key)
    print(f"Sending the four source reports to Groq model {MODEL_NAME}.")

    zero_shot_summary = call_model(client, build_zero_shot_prompt(source_text))
    structured_result = call_model(
        client, build_structured_prompt(zero_shot_summary, source_text)
    )
    OUTPUT_FILE.write_text(
        format_report(zero_shot_summary, structured_result), encoding="utf-8"
    )
    print(f"Saved both model outputs to {OUTPUT_FILE.name}")


if __name__ == "__main__":
    main()


if __name__ == "__main__":
    main()