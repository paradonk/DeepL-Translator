# DeepL Word Document Translator

A Python GUI application for translating Microsoft Word documents using the DeepL API. Features a user-friendly interface with progress tracking, comprehensive content translation (paragraphs, tables, and text boxes), and automatic formatting preservation.

![Python](https://img.shields.io/badge/python-3.7+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## Features

- **Full-Featured GUI** - Easy-to-use tkinter interface
- **Comprehensive Translation**:
  - Document paragraphs
  - Table cells
  - Text boxes (via Word COM automation)
- **27 Supported Languages** - Wide range of target languages
- **Real-Time Progress Tracking** - Progress bar and detailed status log
- **API Key Management** - Secure storage with show/hide toggle
- **Formatting Preservation** - Maintains original document structure
- **Smart Filtering** - Skips numbers and empty content
- **Multi-threaded** - Non-blocking UI during translation

## Screenshots

### Main Interface
- API key entry with show/hide toggle
- File browser for document selection
- Language dropdown selector
- Real-time progress bar
- Detailed status log

## Requirements

- Python 3.7+
- Microsoft Word (for text box translation - Windows only)
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
   - 27 languages available

4. **Start Translation**
   - Click "Start Translation" button
   - Monitor progress bar and status log
   - Wait for completion message

5. **Access Output**
   - Find translated file in same directory
   - Format: `{original}_{lang}_{timestamp}.docx`
   - Example: `report_fr_1729705432.docx`

## Supported Languages

| Code | Language   | Code | Language   | Code | Language    |
|------|------------|------|------------|------|-------------|
| BG   | Bulgarian  | IT   | Italian    | RO   | Romanian    |
| CS   | Czech      | JA   | Japanese   | RU   | Russian     |
| DA   | Danish     | LT   | Lithuanian | SK   | Slovak      |
| DE   | German     | LV   | Latvian    | SL   | Slovenian   |
| EL   | Greek      | NL   | Dutch      | SV   | Swedish     |
| EN   | English    | PL   | Polish     | TR   | Turkish     |
| ES   | Spanish    | PT   | Portuguese | UK   | Ukrainian   |
| ET   | Estonian   | FR   | French     | ZH   | Chinese     |
| FI   | Finnish    | HU   | Hungarian  | ID   | Indonesian  |

## How It Works

### Translation Process

1. **Document Loading** - Opens .docx using `python-docx`
2. **Content Collection** - Identifies paragraphs and table cells
3. **Batch 1: Paragraphs** - Translates all paragraphs via DeepL API
4. **Batch 2: Tables** - Translates all table cell content
5. **Save Temporary** - Saves intermediate file
6. **Text Boxes** - Uses Word COM to access and translate text boxes
7. **Final Save** - Outputs completed file with timestamp

### Smart Content Filtering

The application intelligently skips:
- Empty or whitespace-only text
- Pure numbers (e.g., "123.45")
- Content that shouldn't be translated

### Rate Limiting

Includes a **0.4-second delay** between API calls to respect DeepL rate limits and prevent quota exhaustion.

## Technical Details

### Dependencies

- **deepl** - Official DeepL API client
- **python-docx** - Word document manipulation
- **pywin32** - Windows COM automation for text boxes
- **python-dotenv** - Environment variable management
- **tkinter** - GUI framework (included in Python)

### Architecture

- **GUI Layer** - tkinter-based interface
- **Translation Logic** - DeepL API integration
- **Document Processing** - python-docx + Word COM
- **Threading** - Async translation with UI updates

### Platform Support

| Platform | Paragraph/Table Translation | Text Box Translation |
|----------|----------------------------|----------------------|
| Windows  | ✅ Full Support            | ✅ Full Support      |
| macOS    | ✅ Full Support            | ❌ Not Supported     |
| Linux    | ✅ Full Support            | ❌ Not Supported     |

*Text box translation requires Microsoft Word (COM automation)*

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

### Text boxes not translating
- **Windows**: Ensure Microsoft Word is installed
- **macOS/Linux**: Text box translation not supported (COM limitation)
- Close any open Word instances before translating

### Import errors on startup
```bash
# Install all requirements
pip install -r requirements.txt

# Windows-specific: Install pywin32
pip install pywin32
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
- [ ] Cross-platform text box support
- [ ] Translation memory/cache
- [ ] Source language auto-detection
- [ ] Formality level settings
- [ ] Custom glossary support
- [ ] PDF export option

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [DeepL](https://www.deepl.com/) for their excellent translation API
- [python-docx](https://python-docx.readthedocs.io/) library maintainers
- [pywin32](https://github.com/mhammond/pywin32) project

## Support

- **Issues**: [GitHub Issues](https://github.com/yourusername/deepl-word-translator/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/deepl-word-translator/discussions)

---

**Note**: This tool is not affiliated with or endorsed by DeepL SE.
