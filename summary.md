# ORION Report Generator

`summary.py` reads the four Project ORION Word source documents, sends two prompts to the Groq chat-completions API, and saves both results to `summary_output.md`. The first call creates the requested zero-shot baseline. The second asks for a source-cited project brief, a critique of the baseline, and three improvements. The learner worksheet is deliberately excluded from model input.

## Setup and Run

Install the required packages in the active Python environment:

```bash
python -m pip install groq python-dotenv
```

Set `GROQ_API_KEY` in this folder's `.env` file. Optionally set `GROQ_MODEL`; the default is `openai/gpt-oss-20b`.

```bash
python summary.py
```

The script expects these files next to `summary.py`:

- `01_Messy_Project_Notes.docx`
- `02_Email_Thread_And_Chat.docx`
- `03_Vendor_And_Budget_Notes.docx`
- `04_Service_Desk_Incident_Notes.docx`

It writes or replaces `summary_output.md` in the same folder. Each run sends two requests to Groq and may incur API usage charges. The script does not read `05_Learner_Worksheet.docx`.

## Functions

- **`get_text(element)`** extracts visible text from a Word XML paragraph or table cell.
- **`extract_docx(path)`** opens a `.docx` ZIP archive, parses its Word XML, and returns paragraphs and tables with document headings preserved as citation labels.
- **`load_source_documents()`** verifies each required input file exists, extracts it, and prepends its filename to the source text.
- **`build_zero_shot_prompt(source_text)`** prepares the baseline prompt, “Summarize these documents,” followed by the four extracted reports.
- **`build_structured_prompt(zero_shot_summary, source_text)`** specifies the analyst role, source restrictions, required sections, uncertainty handling, exact citations, critique criteria, and improvement suggestions.
- **`call_model(client, prompt)`** makes one deterministic Groq chat-completions request, enforces an output-token limit, and raises an error if the response is empty.
- **`format_report(zero_shot_summary, structured_result)`** combines both model responses into the final Markdown document.
- **`main()`** loads the source text and API key, runs the two model calls in sequence, and writes the report.

## Report Contents

The zero-shot section records the baseline output. The structured section requests an executive summary; scope; demo and launch status; milestones and owners; incident and QA status; data counts; vendor and budget details; risks and unresolved decisions; source-backed omissions or errors; and three prompt/output improvements. Claims in the structured result should cite the source filename and passage heading. Review generated content against the original files before using it as an official project record.