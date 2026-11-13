# DeepL Word Document Translator - Documentation

## Overview

The project translates Microsoft Word (`.docx`) documents with the DeepL API while keeping their formatting. All front ends share one engine:

| File | Role |
|---|---|
| `docx_translate.py` | Shared translation engine (no UI) |
| `dl_translator.py` | Desktop GUI (tkinter) |
| `app.py` + `templates/index.html` | Web front end (Flask) |
| `op2.py` | Command-line tool (`tl_fast.py` and `tldl.py` are kept as aliases) |

## Dependencies

```
deepl          DeepL API client
lxml           .docx XML processing
python-dotenv  Loads DEEPL_API_KEY from .env
flask          Web front end only
tkinter        GUI / file dialog (standard library)
```

Microsoft Word and pywin32 are **not** required.

```bash
pip install -r requirements.txt
```

## Configuration

The API key is read from the `DEEPL_API_KEY` environment variable or from the `.env` file **next to the scripts** (not the current working directory), so the tools work from any folder. Never put the key in source code.

Optional settings for the web app:

| Variable | Purpose |
|---|---|
| `FLASK_SECRET_KEY` | Session signing key. A random one is generated at start-up if unset. |
| `FLASK_DEBUG=1` | Enables Flask debug mode. Off by default; never enable it on a shared network. |

## How the engine works (`docx_translate.py`)

1. **Open the package.** A `.docx` is a ZIP. The engine reads these parts:
   `word/document.xml`, every `word/headerN.xml` and `word/footerN.xml`, `word/footnotes.xml` and `word/endnotes.xml`.
2. **Collect paragraphs.** Every `<w:p>` in those parts is a candidate, including paragraphs inside tables and text boxes (`<w:txbxContent>`), and the legacy VML copy of each text box. Text belonging to a paragraph nested inside another (a text box anchored in a paragraph) is kept separate.
3. **Filter.** Paragraphs with no letters at all (numbers, dates, symbols) are skipped.
4. **Encode formatting.** Each text run becomes an `<s>` tag, e.g. `The <s>quick</s> brown fox`. Whitespace at the edges of runs is kept outside the tags so spaces between differently formatted words survive.
5. **De-duplicate and batch.** Identical paragraphs are sent once. Texts go to DeepL in batches of up to 50 per request with `tag_handling="xml"`, so DeepL translates around the tags. No artificial delays are needed; the DeepL client retries on rate-limit responses.
6. **Write back.** The translated pieces are put back into the original runs in reading order, so bold/italic/fonts/colours/hyperlinks stay attached. If DeepL reorders words across tags, the text order is always correct and only the formatting may shift to a neighbouring word.
7. **Save.** The ZIP is rewritten entry by entry, keeping the original order and every untouched part byte for byte. Input and output can be file paths or in-memory file objects.

Public API:

```python
import docx_translate as dt
client = dt.make_client(dt.load_api_key())
stats = dt.translate_docx("in.docx", "out.docx", "FR", client,
                          progress=lambda done, total: ..., log=print)
# stats == {"paragraphs": <unique paragraphs sent>, "characters": <characters sent>}
```

Other helpers: `fetch_target_languages(client)`, `normalize_lang(code)` (maps bare `EN`/`PT`/`ZH` to `EN-GB`/`PT-PT`/`ZH-HANS`), `output_path_for(path, lang)`, and the fallback list `TARGET_LANGUAGES`.

## Desktop GUI (`dl_translator.py`)

- Loads the key from `.env`; **Save to .env** writes it next to the scripts.
- Fetches the current target-language list from DeepL when a key is available and falls back to the built-in list otherwise. Languages are shown as `French (FR)`.
- Translation runs on a worker thread. The worker never touches widgets: it posts log, progress and completion messages to a queue that the Tk main loop drains every 100 ms, because Tk is not thread-safe.
- On failure, any partially written output file is deleted.

Output name: `{original}_{lang}_{unix_timestamp}.docx`, e.g. `report_fr_1729705432.docx`, saved next to the source file.

## Web front end (`app.py`)

- Uploads are processed entirely in memory. Nothing is written to disk, so simultaneous users can't overwrite each other's files.
- The browser generates a job id per submission and polls `/progress/<job_id>`, so each user sees only their own progress.
- `/quota` shows the account's character usage.
- Offers every DeepL target language (loaded at start-up; restart to pick up new ones), with a type-to-filter search box. French is preselected (`DEFAULT_LANGUAGE`).
- Uploads are limited to 50 MB, and only languages from that list are accepted.

Run with `python app.py`, then open http://127.0.0.1:5000.

## Command line (`op2.py`)

```bash
python op2.py                         # file dialog + language prompt
python op2.py report.docx --lang de   # non-interactive
```

## Supported languages

The GUI and CLI use DeepL's live list. The fallback list is:
AR, BG, CS, DA, DE, EL, EN-GB, EN-US, ES, ET, FI, FR, HE, HU, ID, IT, JA, KO, LT, LV, NB, NL, PL, PT-BR, PT-PT, RO, RU, SK, SL, SV, TH, TR, UK, VI, ZH-HANS, ZH-HANT.

Bare `EN`, `PT` and `ZH` are mapped to `EN-GB`, `PT-PT` and `ZH-HANS`.

## Limitations

- Only `.docx` (not `.doc`, `.pdf`).
- Comments, charts and SmartArt text are not translated.
- When DeepL reorders words across formatting boundaries, formatting can land on a neighbouring word (text is always in the right order).
- One document per run.
