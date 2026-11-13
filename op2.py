"""Command-line DeepL .docx translator.

Usage:
    python op2.py                          # pick a file in a dialog, type the language
    python op2.py report.docx --lang fr    # no prompts
"""
import argparse
import sys
from pathlib import Path

import docx_translate as dt


def pick_file() -> str:
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(title="Select Word File", filetypes=[("Word Documents", "*.docx")])
    root.destroy()
    return path


def print_progress(done: int, total: int, width: int = 30) -> None:
    filled = int(width * done / total) if total else width
    percent = int(100 * done / total) if total else 100
    print(f"\r[{'#' * filled}{'-' * (width - filled)}] {percent}% ({done}/{total})",
          end="\n" if done >= total else "", flush=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Translate a Word document with DeepL.")
    parser.add_argument("file", nargs="?", help="path to a .docx file (opens a dialog if omitted)")
    parser.add_argument("--lang", help="target language code, e.g. fr, de, en-us, pt-br, zh-hans")
    args = parser.parse_args()

    api_key = dt.load_api_key()
    if not api_key:
        sys.exit("DeepL API key not found. Set DEEPL_API_KEY in .env.")

    file_path = args.file or pick_file()
    if not file_path:
        sys.exit("No file selected.")
    input_path = Path(file_path)
    if not input_path.is_file() or input_path.suffix.lower() != ".docx":
        sys.exit("File not found or not a .docx file.")

    client = dt.make_client(api_key)
    try:
        languages = dict(dt.fetch_target_languages(client))
    except Exception:
        languages = dict(dt.TARGET_LANGUAGES)

    lang = args.lang or input("Enter target language code (e.g. fr, de, en-us, pt-br): ")
    lang = dt.normalize_lang(lang)
    if lang not in languages:
        sys.exit(f"Unsupported language code: {lang}. Choose from: {', '.join(sorted(languages))}")

    output_path = dt.output_path_for(input_path, lang)
    print(f"Translating {input_path.name} to {languages[lang]} ({lang})...")
    try:
        stats = dt.translate_docx(input_path, output_path, lang, client, progress=print_progress, log=print)
    except Exception as e:
        output_path.unlink(missing_ok=True)
        sys.exit(f"Translation failed: {e}")

    print(f"\n✅ Translation complete ({stats['characters']:,} characters). Saved to: {output_path}")


if __name__ == "__main__":
    main()
