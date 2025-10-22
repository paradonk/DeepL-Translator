# DeepL Word Document Translator - Documentation

## Overview

`dl_translator.py` is a GUI-based application that translates Microsoft Word documents using the DeepL API. It provides a user-friendly interface for translating documents while preserving formatting and handling various content types including paragraphs, tables, and text boxes.

## Features

- **Graphical User Interface** - Easy-to-use tkinter-based GUI
- **Comprehensive Translation** - Translates paragraphs, table cells, and text boxes
- **API Key Management** - Secure storage in .env file with show/hide toggle
- **27 Supported Languages** - Wide range of target languages
- **Progress Tracking** - Real-time progress bar and detailed status log
- **Multi-threaded** - Non-blocking UI during translation
- **Automatic File Naming** - Output files include language code and timestamp

## Dependencies

```python
deepl              # DeepL API client
python-docx        # Word document manipulation
pywin32            # Windows COM automation for text boxes
python-dotenv      # Environment variable management
tkinter            # GUI framework (standard library)
```

## Installation

```bash
pip install deepl python-docx pywin32 python-dotenv
```

## Usage

### Running the Application

```bash
python dl_translator.py
```

### Workflow

1. **Enter API Key** - Input your DeepL API key (can be saved to .env)
2. **Select Document** - Browse and select a .docx file
3. **Choose Target Language** - Select from 27 available languages
4. **Start Translation** - Click "Start Translation" button
5. **Monitor Progress** - Watch progress bar and status log
6. **Access Output** - Translated file saved with format: `{original}_{lang}_{timestamp}.docx`

## Application Structure

### Class: `TranslatorGUI`

Main application class managing the GUI and translation logic.

#### Initialization (`__init__`, lines 14-31)

```python
def __init__(self, root):
    # Initialize window
    # Load API key from .env
    # Setup UI components
```

**Variables:**
- `file_path` - Selected Word document path
- `api_key` - DeepL API key
- `target_lang` - Target language code (default: "FR")
- `is_translating` - Translation state flag

#### UI Components (`setup_ui`, lines 33-101)

1. **API Key Section** (lines 38-49)
   - Entry field with password masking
   - Show/Hide toggle button
   - Save to .env button

2. **File Selection** (lines 51-60)
   - Read-only path display
   - Browse button for file dialog

3. **Language Selection** (lines 62-78)
   - Dropdown combo box
   - 27 language options

4. **Progress Tracking** (lines 80-88)
   - Progress bar (determinate mode)
   - Counter label (X / Y format)

5. **Status Log** (lines 90-96)
   - Scrollable text area
   - Read-only display
   - Auto-scroll to bottom

6. **Translate Button** (lines 98-101)
   - Start/Stop translation
   - State management

### Key Methods

#### `toggle_api_key()` (lines 103-107)
Toggles API key visibility between masked and plain text.

#### `save_api_key()` (lines 109-121)
Saves API key to .env file with validation.

#### `select_file()` (lines 123-130)
Opens file dialog for .docx selection and updates UI.

#### `log(message)` (lines 132-137)
Appends messages to status log with auto-scroll.

#### `start_translation()` (lines 139-162)
- Validates inputs (API key, file path)
- Initializes UI for translation state
- Spawns translation thread

#### `translate_document()` (lines 164-280)
Main translation logic executed in separate thread.

**Translation Process:**

1. **Initialization** (lines 165-173)
   - Load file path and target language
   - Initialize DeepL translator
   - Define translation filter

2. **Document Loading** (lines 179-196)
   - Load .docx using python-docx
   - Collect paragraphs to translate
   - Collect table cell paragraphs
   - Calculate total items
   - Update progress display

3. **Paragraph Translation** (lines 198-210)
   - Iterate through paragraphs
   - Translate text via DeepL API
   - Update paragraph in-place
   - Log progress
   - 0.4 second delay between calls

4. **Table Cell Translation** (lines 212-224)
   - Iterate through table paragraphs
   - Translate text via DeepL API
   - Update cell in-place
   - Log progress
   - 0.4 second delay between calls

5. **Save Temporary File** (lines 226-229)
   - Save document as temp file
   - Prepare for COM automation

6. **Text Box Translation** (lines 231-252)
   - Open document via Word COM
   - Find shapes with text
   - Translate text box content
   - Update text boxes
   - 0.4 second delay between calls

7. **Finalization** (lines 254-270)
   - Save final document with timestamp
   - Close Word COM objects
   - Delete temporary file
   - Display success message

8. **Error Handling** (lines 272-279)
   - Catch and log exceptions
   - Reset UI state
   - Show error dialog

### Helper Function

#### `should_translate(text)` (lines 175-177)
Determines if text should be translated.

**Logic:**
- Returns `False` for empty/whitespace text
- Returns `False` for numeric-only content (e.g., "123.45")
- Returns `True` otherwise

## Supported Languages

| Language    | Code | Language   | Code |
|-------------|------|------------|------|
| Bulgarian   | BG   | Italian    | IT   |
| Czech       | CS   | Japanese   | JA   |
| Danish      | DA   | Lithuanian | LT   |
| German      | DE   | Latvian    | LV   |
| Greek       | EL   | Dutch      | NL   |
| English     | EN   | Polish     | PL   |
| Spanish     | ES   | Portuguese | PT   |
| Estonian    | ET   | Romanian   | RO   |
| Finnish     | FI   | Russian    | RU   |
| French      | FR   | Slovak     | SK   |
| Hungarian   | HU   | Slovenian  | SL   |
| Indonesian  | ID   | Swedish    | SV   |
|             |      | Turkish    | TR   |
|             |      | Ukrainian  | UK   |
|             |      | Chinese    | ZH   |

## Output File Naming

Format: `{original_name}_{language_code}_{unix_timestamp}.docx`

**Example:**
- Input: `report.docx`
- Output: `report_fr_1729705432.docx`

## Rate Limiting

The application includes a **0.4 second delay** between translation requests to avoid API rate limiting (lines 208, 222, 250).

## Error Handling

- **Validation Errors** - Missing API key or file path
- **Translation Errors** - Individual item failures logged but don't stop process
- **Critical Errors** - Caught and displayed in error dialog
- **Resource Cleanup** - Ensures Word COM objects are properly closed

## Threading Model

- **Main Thread** - GUI event loop
- **Worker Thread** - Translation process (daemon thread)
- **Thread Safety** - UI updates via `update_idletasks()`

## Technical Notes

### Why Word COM Automation?

The `python-docx` library cannot access text boxes in Word documents. The application uses Windows COM automation (`win32com.client`) to:
1. Open the document in Word
2. Iterate through Shapes collection
3. Translate text box content
4. Save the modified document

### Temporary File Workflow

1. Translate paragraphs and tables with `python-docx`
2. Save as `{name}_temp.docx`
3. Open temp file with Word COM
4. Translate text boxes
5. Save as final file
6. Delete temp file

## Limitations

- **Windows Only** - Requires Microsoft Word installed (COM automation)
- **DOCX Format** - Only supports .docx files (not .doc)
- **Sequential Processing** - Translates items one at a time
- **No Batch Mode** - Single document per run

## Security Considerations

- API key masked by default with show/hide toggle
- API key stored in .env file (should be in .gitignore)
- No validation of API key format or permissions

## Future Improvements

- Batch document processing
- Parallel translation for better performance
- Support for additional file formats
- Translation memory/cache
- Source language detection
- Formality settings for supported languages
- Glossary support

## License

Not specified in the file.

## Author

Not specified in the file.
