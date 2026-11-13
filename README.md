# DeepL Word Document Translator

A Python application for translating Microsoft Word documents with the DeepL API, available as a desktop GUI, a web app and a command-line tool. It translates paragraphs, tables, text boxes, headers, footers and footnotes while keeping the original formatting.

![Python](https://img.shields.io/badge/python-3.9+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## Features

- **Full-Featured GUI** - Easy-to-use tkinter interface (plus a Flask web app and a CLI)
- **Comprehensive Translation**:
  - Document paragraphs and table cells
  - Text boxes
  - All headers, footers, footnotes and endnotes
- **Formatting Preservation** - Bold, italic, fonts, colours and links stay on the right words
- **30+ Languages** - Live list from DeepL, including EN-GB/EN-US and PT-PT/PT-BR
- **Fast** - Paragraphs are sent in batches and repeated text is translated once
- **Cross-platform** - No Microsoft Word needed
- **Real-Time Progress Tracking** - Progress bar and detailed status log
- **API Key Management** - Stored in `.env`, with show/hide toggle
- **Smart Filtering** - Skips empty text and text with no letters (numbers, dates)

## Screenshots

### Main Interface
- API key entry with show/hide toggle
- File browser for document selection
- Language dropdown selector
- Real-time progress bar
- Detailed status log

## Requirements

- Python 3.9+
- DeepL API key ([Get one here](https://www.deepl.com/pro-api))

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/deepl-word-translator.git
cd deepl-word-translator
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Set up your DeepL API key:
   - Copy `.env.example` to `.env`
   - Add your DeepL API key:
   ```
   DEEPL_API_KEY=your_api_key_here
   ```

## Usage

### Starting the Application

```bash
python dl_translator.py
```

### Translation Workflow

1. **Enter API Key**
   - Paste your DeepL API key in the text field
   - Click "Save to .env" to store it permanently
   - Use "Show" button to toggle visibility

2. **Select Document**
   - Click "Browse..." button
   - Choose a `.docx` file
   - File path will display in the interface

3. **Choose Language**
   - Select target language from dropdown
   - The list is loaded from DeepL when your key is set

4. **Start Translation**
   - Click "Start Translation" button
   - Monitor progress bar and status log
   - Wait for completion message

5. **Access Output**
   - Find translated file in same directory
   - Format: `{original}_{lang}_{timestamp}.docx`
   - Example: `report_fr_1729705432.docx`

## Other Ways to Run

```bash
python app.py                          # web app at http://127.0.0.1:5000
python op2.py report.docx --lang de    # command line
```

## Supported Languages

The GUI and CLI show DeepL's current list. Built-in fallback:

| Code | Language | Code | Language | Code | Language |
|------|----------|------|----------|------|----------|
| AR | Arabic | HU | Hungarian | PT-BR | Portuguese (Brazilian) |
| BG | Bulgarian | ID | Indonesian | PT-PT | Portuguese (European) |
| CS | Czech | IT | Italian | RO | Romanian |
| DA | Danish | JA | Japanese | RU | Russian |
| DE | German | KO | Korean | SK | Slovak |
| EL | Greek | LT | Lithuanian | SL | Slovenian |
| EN-GB | English (British) | LV | Latvian | SV | Swedish |
| EN-US | English (American) | NB | Norwegian (Bokmål) | TH | Thai |
| ES | Spanish | NL | Dutch | TR | Turkish |
| ET | Estonian | PL | Polish | UK | Ukrainian |
| FI | Finnish | HE | Hebrew | VI | Vietnamese |
| FR | French | ZH-HANS | Chinese (Simplified) | ZH-HANT | Chinese (Traditional) |

## How It Works

1. **Open** - A `.docx` is a ZIP of XML files; the body, headers, footers, footnotes and endnotes are read with `lxml`
2. **Collect** - Every paragraph is gathered, including those in tables and text boxes
3. **Tag** - Each formatted run is wrapped in a tag so DeepL translates around the formatting
4. **Translate** - Unique paragraphs are sent in batches of up to 50 with DeepL's XML tag handling
5. **Write back** - Translations are put back into the original runs, keeping their formatting
6. **Save** - The package is rewritten with everything else untouched

## Technical Details

### Dependencies

- **deepl** - Official DeepL API client
- **lxml** - Word document XML processing
- **python-dotenv** - Environment variable management
- **flask** - Web app
- **tkinter** - GUI framework (included in Python)

### Architecture

- **Engine** - `docx_translate.py`, shared by all front ends
- **GUI** - `dl_translator.py`; translation runs on a worker thread and UI updates go through a queue
- **Web** - `app.py`; uploads are processed in memory with per-user progress
- **CLI** - `op2.py`

Works the same on Windows, macOS and Linux.

## File Output Format

Translated files use the naming pattern:

```
{original_filename}_{language_code}_{unix_timestamp}.docx
```

**Examples:**
- `annual_report.docx` → `annual_report_de_1729705432.docx`
- `meeting_notes.docx` → `meeting_notes_ja_1729705433.docx`

## Troubleshooting

### "Please enter your DeepL API key!"
- Ensure you've entered a valid API key
- Check that the key isn't expired
- Verify you have available quota

### "Translation failed" errors
- Check your internet connection
- Verify API key has sufficient quota
- Ensure document isn't corrupted or password-protected

### "Could not load language list"
- The key is missing or invalid, or you are offline; the built-in list is used instead

### Import errors on startup
```bash
# Install all requirements
pip install -r requirements.txt
```

## Security Best Practices

- ✅ Never commit `.env` file to Git
- ✅ Keep API keys private and secure
- ✅ Use the password masking in the GUI
- ✅ Regularly rotate API keys
- ❌ Don't share screenshots with visible API keys

## Getting a DeepL API Key

1. Visit [DeepL API Portal](https://www.deepl.com/pro-api)
2. Sign up for free or paid account
3. Navigate to account settings
4. Copy your authentication key
5. Add to `.env` file or paste in GUI

**Free Tier:** 500,000 characters/month

## Documentation

For detailed technical documentation, see [dl_translator_documentation.md](dl_translator_documentation.md)

## Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## Roadmap

- [ ] Batch document processing
- [x] Cross-platform text box support
- [ ] Translation memory/cache
- [ ] Source language auto-detection
- [ ] Formality level settings
- [ ] Custom glossary support
- [ ] PDF export option

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [DeepL](https://www.deepl.com/) for their excellent translation API
- [lxml](https://lxml.de/) library maintainers

## Support

- **Issues**: [GitHub Issues](https://github.com/yourusername/deepl-word-translator/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/deepl-word-translator/discussions)

---

**Note**: This tool is not affiliated with or endorsed by DeepL SE.
