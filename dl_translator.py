import os
import sys
import time
import deepl
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from dotenv import load_dotenv, set_key
from docx import Document
import win32com.client as win32
import threading

class TranslatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DeepL Word Document Translator")
        self.root.geometry("700x600")
        self.root.resizable(True, True)

        # Variables
        self.file_path = tk.StringVar()
        self.api_key = tk.StringVar()
        self.target_lang = tk.StringVar(value="FR")
        self.is_translating = False

        # Load existing API key
        load_dotenv()
        existing_key = os.getenv("DEEPL_API_KEY", "")
        self.api_key.set(existing_key)

        self.setup_ui()

    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # === API Key Section ===
        ttk.Label(main_frame, text="DeepL API Key:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5))

        api_frame = ttk.Frame(main_frame)
        api_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        self.api_entry = ttk.Entry(api_frame, textvariable=self.api_key, width=50, show="*")
        self.api_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(api_frame, text="Show", command=self.toggle_api_key, width=8).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(api_frame, text="Save to .env", command=self.save_api_key, width=12).pack(side=tk.LEFT)

        # === File Selection ===
        ttk.Label(main_frame, text="Word Document:", font=('Arial', 10, 'bold')).grid(
            row=2, column=0, sticky=tk.W, pady=(10, 5))

        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Entry(file_frame, textvariable=self.file_path, width=50, state='readonly').pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(file_frame, text="Browse...", command=self.select_file, width=12).pack(side=tk.LEFT)

        # === Language Selection ===
        ttk.Label(main_frame, text="Target Language:", font=('Arial', 10, 'bold')).grid(
            row=4, column=0, sticky=tk.W, pady=(10, 5))

        lang_map = {
            'Bulgarian': 'BG', 'Czech': 'CS', 'Danish': 'DA', 'German': 'DE',
            'Greek': 'EL', 'English': 'EN', 'Spanish': 'ES', 'Estonian': 'ET',
            'Finnish': 'FI', 'French': 'FR', 'Hungarian': 'HU', 'Indonesian': 'ID',
            'Italian': 'IT', 'Japanese': 'JA', 'Lithuanian': 'LT', 'Latvian': 'LV',
            'Dutch': 'NL', 'Polish': 'PL', 'Portuguese': 'PT', 'Romanian': 'RO',
            'Russian': 'RU', 'Slovak': 'SK', 'Slovenian': 'SL', 'Swedish': 'SV',
            'Turkish': 'TR', 'Ukrainian': 'UK', 'Chinese': 'ZH'
        }

        self.lang_combo = ttk.Combobox(main_frame, textvariable=self.target_lang,
                                       values=list(lang_map.values()), width=15, state='readonly')
        self.lang_combo.grid(row=5, column=0, sticky=tk.W, pady=(0, 10))

        # === Progress Bar ===
        ttk.Label(main_frame, text="Progress:", font=('Arial', 10, 'bold')).grid(
            row=6, column=0, sticky=tk.W, pady=(10, 5))

        self.progress = ttk.Progressbar(main_frame, mode='determinate', length=660)
        self.progress.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 5))

        self.progress_label = ttk.Label(main_frame, text="0 / 0", foreground="gray")
        self.progress_label.grid(row=8, column=0, sticky=tk.W, pady=(0, 10))

        # === Status Log ===
        ttk.Label(main_frame, text="Status Log:", font=('Arial', 10, 'bold')).grid(
            row=9, column=0, sticky=tk.W, pady=(10, 5))

        self.log_text = scrolledtext.ScrolledText(main_frame, width=80, height=12,
                                                   state='disabled', wrap=tk.WORD)
        self.log_text.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        # === Translate Button ===
        self.translate_btn = ttk.Button(main_frame, text="Start Translation",
                                       command=self.start_translation, width=20)
        self.translate_btn.grid(row=11, column=0, columnspan=3, pady=(10, 0))

    def toggle_api_key(self):
        if self.api_entry.cget('show') == '*':
            self.api_entry.config(show='')
        else:
            self.api_entry.config(show='*')

    def save_api_key(self):
        key = self.api_key.get().strip()
        if not key:
            messagebox.showwarning("Warning", "API key is empty!")
            return

        env_path = Path(".env")
        if not env_path.exists():
            env_path.write_text("")

        set_key(str(env_path), "DEEPL_API_KEY", key)
        messagebox.showinfo("Success", "API key saved to .env file!")
        self.log("API key saved to .env file")

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx")]
        )
        if file_path:
            self.file_path.set(file_path)
            self.log(f"Selected file: {Path(file_path).name}")

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update_idletasks()

    def start_translation(self):
        # Validation
        if not self.api_key.get().strip():
            messagebox.showerror("Error", "Please enter your DeepL API key!")
            return

        if not self.file_path.get():
            messagebox.showerror("Error", "Please select a Word document!")
            return

        if self.is_translating:
            messagebox.showwarning("Warning", "Translation already in progress!")
            return

        # Start translation in a separate thread
        self.is_translating = True
        self.translate_btn.config(state='disabled', text="Translating...")
        self.progress['value'] = 0
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

        thread = threading.Thread(target=self.translate_document, daemon=True)
        thread.start()

    def translate_document(self):
        try:
            input_path = Path(self.file_path.get())
            target_lang = self.target_lang.get()

            self.log(f"Starting translation to {target_lang}...")
            self.log("=" * 60)

            # Initialize translator
            translator = deepl.Translator(self.api_key.get().strip())

            def should_translate(text):
                text = text.strip()
                return text and not text.replace('.', '', 1).isdigit()

            # Load document
            self.log("Loading document...")
            doc = Document(str(input_path))

            # Collect all translatable items
            paragraphs = [p for p in doc.paragraphs if should_translate(p.text)]
            tables_texts = [
                para for table in doc.tables
                for row in table.rows
                for cell in row.cells
                for para in cell.paragraphs
                if should_translate(para.text)
            ]
            total = len(paragraphs) + len(tables_texts)
            count = 0

            self.log(f"Found {len(paragraphs)} paragraphs and {len(tables_texts)} table cells to translate")
            self.progress_label.config(text=f"{count} / {total}")

            # Translate paragraphs
            for para in paragraphs:
                original = para.text.strip()
                try:
                    result = translator.translate_text(original, target_lang=target_lang)
                    para.text = result.text
                    count += 1
                    self.log(f"[{count}/{total}] Para: {original[:50]}...")
                    self.progress['value'] = (count / total) * 100
                    self.progress_label.config(text=f"{count} / {total}")
                    time.sleep(0.4)
                except Exception as e:
                    self.log(f"ERROR in para: {str(e)}")

            # Translate table cells
            for para in tables_texts:
                original = para.text.strip()
                try:
                    result = translator.translate_text(original, target_lang=target_lang)
                    para.text = result.text
                    count += 1
                    self.log(f"[{count}/{total}] Table: {original[:50]}...")
                    self.progress['value'] = (count / total) * 100
                    self.progress_label.config(text=f"{count} / {total}")
                    time.sleep(0.4)
                except Exception as e:
                    self.log(f"ERROR in table: {str(e)}")

            # Save temp file
            temp_path = input_path.with_name(f"{input_path.stem}_temp.docx")
            doc.save(str(temp_path))
            self.log("Saved temporary file, processing text boxes...")

            # Translate text boxes via Word COM
            word = win32.gencache.EnsureDispatch("Word.Application")
            word.Visible = False
            doc_com = word.Documents.Open(str(temp_path.absolute()))

            shape_texts = [s for s in doc_com.Shapes if s.TextFrame.HasText > 0 and should_translate(s.TextFrame.TextRange.Text)]
            total += len(shape_texts)
            self.log(f"Found {len(shape_texts)} text boxes to translate")
            self.progress_label.config(text=f"{count} / {total}")

            for shape in shape_texts:
                original = shape.TextFrame.TextRange.Text.strip()
                try:
                    result = translator.translate_text(original, target_lang=target_lang)
                    shape.TextFrame.TextRange.Text = result.text
                    count += 1
                    self.log(f"[{count}/{total}] Text box: {original[:50]}...")
                    self.progress['value'] = (count / total) * 100
                    self.progress_label.config(text=f"{count} / {total}")
                    time.sleep(0.4)
                except Exception as e:
                    self.log(f"ERROR in text box: {str(e)}")

            # Save final output
            timestamp = int(time.time())
            short_code = target_lang.lower()
            final_path = input_path.with_name(f"{input_path.stem}_{short_code}_{timestamp}.docx")
            doc_com.SaveAs(str(final_path.absolute()))
            doc_com.Close(False)
            word.Quit()

            # Clean up temp file
            if temp_path.exists():
                temp_path.unlink()

            self.log("=" * 60)
            self.log(f"✅ Translation complete!")
            self.log(f"Saved to: {final_path}")

            messagebox.showinfo("Success", f"Translation complete!\n\nSaved to:\n{final_path.name}")

        except Exception as e:
            self.log(f"CRITICAL ERROR: {str(e)}")
            messagebox.showerror("Error", f"Translation failed:\n{str(e)}")

        finally:
            self.is_translating = False
            self.translate_btn.config(state='normal', text="Start Translation")
            self.progress['value'] = 100

def main():
    root = tk.Tk()
    app = TranslatorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
