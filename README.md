# Document Translation Pipeline

## Overview
This project extracts content from PowerPoint decks, translates only the human‑readable text while preserving every formatting detail, and reassembles a fully formatted PPTX in the target language. It consists of three main scripts:
- `extractor.py` – walks a source `.pptx`, exporting slide masters, layouts, shapes, tables, charts, SmartArt, speaker notes, and links into a rich JSON structure.
- `translator.py` – feeds the JSON to OpenAI, translating text elements but keeping all metadata intact.
- `reassembler.py` – loads the translated JSON and writes the translated text back into a copy of the original PPTX template.

## Prerequisites
- Python 3.8+
- PowerPoint template to translate (`.pptx`)
- OpenAI API key with access to the `gpt-4o-mini` model (default)
- Recommended Python packages (install with pip):
  ```
  python3 -m pip install python-pptx lxml python-dotenv openai
  ```

## Setup
1. **Get the source**
   - Download or copy the project folder onto your machine.
   - Open a terminal in the project directory.
2. **Create a virtual environment (optional but recommended)**
   ```
  python3 -m venv venv
  source venv/bin/activate
   ```
3. **Configure environment variables**
   - Copy `.env.example` to `.env` if available, otherwise create `.env`.
   - Add your OpenAI key:
     ```
     OPENAI_API_KEY=sk-...
     ```
   - The scripts load `.env` automatically via `python-dotenv`.

## Usage
### 1. Extract source content
Run `extractor.py` to convert a PPTX into JSON:
```
python3 extractor.py
```
Key behaviors:
- Reads the PPTX path from the `pptx_file` variable in `extractor.py` (default: `BI SAM_Negotiations.pptx`).
- Writes a comprehensive JSON file (`extracted_content_with_layouts.json`) that includes slide masters, slide backgrounds, placeholder geometry, element metadata, and text runs.
- Update the file paths in the `__main__` block if you need different source/target names.

### 2. Translate JSON content
Use `translator.py` to translate the extracted JSON while keeping metadata untouched:
```
python3 translator.py extracted_content_with_layouts.json \
  -o extracted_content_with_layouts_translated_spanish.json \
  -l Spanish
```
Important details:
- Loads `OPENAI_API_KEY` from `.env` unless `--api-key` is provided.
- Translates in batches via `gpt-4o-mini`, validating that response JSON matches the input structure.
- Preserves slide masters, backgrounds, SmartArt structures, chart/table defaults, and all formatting details.
- Tracks basic statistics (API calls, tokens, texts translated) and prints them on completion.

### 3. Reassemble the translated deck
Use `reassembler.py` to merge translations back into the original PPTX template:
```
python3 reassembler.py BI SAM_Negotiations.pptx \
  extracted_content_with_layouts_translated_spanish.json \
  Output.pptx
```
Notes:
- All arguments must be on one line: `python3 reassembler.py <original.pptx> <translated.json> [output.pptx]`.
- Matches shapes by `shape_id` and replaces only text content, keeping formatting, fills, shadows, and layout assignments intact.
- Updates tables, charts, SmartArt, and speaker notes when translations are present.

## End-to-End Workflow
1. Place the original deck (e.g., `BI SAM_Negotiations.pptx`) in the project folder.
2. `python3 extractor.py` → produces `extracted_content_with_layouts.json`.
3. `python3 translator.py extracted_content_with_layouts.json -l Spanish` → outputs `extracted_content_with_layouts_translated_spanish.json`.
4. `python3 reassembler.py BI SAM_Negotiations.pptx extracted_content_with_layouts_translated_spanish.json Output.pptx` → yields the localized PowerPoint.

## Troubleshooting
- **“OPENAI_API_KEY not found”** – ensure `.env` exists and contains a valid key, or pass `--api-key` explicitly.
- **“unrecognized arguments” when running `reassembler.py`** – provide all arguments on a single command line; there should be no newline before the output path.
- **Large decks / rate limits** – the translator throttles slightly between slides; adjust `time.sleep` in `translate_slide` if you hit API limits.
- **Regenerating outputs** – `.gitignore` excludes generated JSON and PPTX files; re-run the extractor/translator as needed.

## Additional Notes
- All scripts include comments describing their logic; consult the source files for deeper customization.
- If you add new metadata fields to the extractor JSON, ensure the translator preserves them (via deep copies) and extend the reassembler if those fields need to drive PowerPoint updates.
- Keep version control history clean by tracking only source code, configuration, and the original template; regenerate intermediate artefacts when needed.

