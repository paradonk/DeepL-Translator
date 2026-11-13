"""DeepL Word Document Translator - desktop GUI."""
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from dotenv import set_key

import docx_translate as dt


class TranslatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("DeepL Word Document Translator")
        self.root.geometry("700x600")
        self.root.resizable(True, True)

        self.file_path = tk.StringVar()
        self.api_key = tk.StringVar(value=dt.load_api_key())
        self.target_lang = tk.StringVar()
        self.is_translating = False
        self.languages = list(dt.TARGET_LANGUAGES)

        # The worker thread never touches widgets; it posts messages here and
        # the Tk main loop applies them (Tk is not thread-safe).
        self.events = queue.Queue()

        self.setup_ui()
        self.set_languages(self.languages, default="FR")
        self.root.after(100, self.process_events)
        if self.api_key.get():
            self.refresh_languages()

    # ------------------------------------------------------------------ UI

    def setup_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(10, weight=1)

        bold = ('Arial', 10, 'bold')

        ttk.Label(main_frame, text="DeepL API Key:", font=bold).grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        api_frame = ttk.Frame(main_frame)
        api_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.api_entry = ttk.Entry(api_frame, textvariable=self.api_key, width=50, show="*")
        self.api_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.show_btn = ttk.Button(api_frame, text="Show", command=self.toggle_api_key, width=8)
        self.show_btn.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(api_frame, text="Save to .env", command=self.save_api_key, width=12).pack(side=tk.LEFT)

        ttk.Label(main_frame, text="Word Document:", font=bold).grid(row=2, column=0, sticky=tk.W, pady=(10, 5))
        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        ttk.Entry(file_frame, textvariable=self.file_path, width=50, state='readonly').pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(file_frame, text="Browse...", command=self.select_file, width=12).pack(side=tk.LEFT)

        ttk.Label(main_frame, text="Target Language:", font=bold).grid(row=4, column=0, sticky=tk.W, pady=(10, 5))
        self.lang_combo = ttk.Combobox(main_frame, textvariable=self.target_lang, width=35, state='readonly')
        self.lang_combo.grid(row=5, column=0, sticky=tk.W, pady=(0, 10))

        ttk.Label(main_frame, text="Progress:", font=bold).grid(row=6, column=0, sticky=tk.W, pady=(10, 5))
        self.progress = ttk.Progressbar(main_frame, mode='determinate', maximum=100)
        self.progress.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        self.progress_label = ttk.Label(main_frame, text="0 / 0", foreground="gray")
        self.progress_label.grid(row=8, column=0, sticky=tk.W, pady=(0, 10))

        ttk.Label(main_frame, text="Status Log:", font=bold).grid(row=9, column=0, sticky=tk.W, pady=(10, 5))
        self.log_text = scrolledtext.ScrolledText(main_frame, width=80, height=12, state='disabled', wrap=tk.WORD)
        self.log_text.grid(row=10, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        self.translate_btn = ttk.Button(main_frame, text="Start Translation", command=self.start_translation, width=20)
        self.translate_btn.grid(row=11, column=0, pady=(10, 0))

    def set_languages(self, languages, default=None):
        current = self.selected_code() or default
        self.languages = languages
        self.lang_combo['values'] = [f"{name} ({code})" for code, name in languages]
        codes = [code for code, _ in languages]
        if current in codes:
            self.lang_combo.current(codes.index(current))
        elif codes:
            self.lang_combo.current(0)

    def selected_code(self):
        index = self.lang_combo.current()
        return self.languages[index][0] if index >= 0 else None

    def toggle_api_key(self):
        hidden = self.api_entry.cget('show') == '*'
        self.api_entry.config(show='' if hidden else '*')
        self.show_btn.config(text="Hide" if hidden else "Show")

    def save_api_key(self):
        key = self.api_key.get().strip()
        if not key:
            messagebox.showwarning("Warning", "API key is empty!")
            return
        dt.ENV_PATH.touch(exist_ok=True)
        set_key(str(dt.ENV_PATH), "DEEPL_API_KEY", key)
        messagebox.showinfo("Success", f"API key saved to {dt.ENV_PATH}")
        self.log(f"API key saved to {dt.ENV_PATH}")
        self.refresh_languages()

    def select_file(self):
        file_path = filedialog.askopenfilename(title="Select Word Document",
                                               filetypes=[("Word Documents", "*.docx")])
        if file_path:
            self.file_path.set(file_path)
            self.log(f"Selected file: {Path(file_path).name}")

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    # ------------------------------------------------------- worker bridge

    def post(self, kind, *args):
        """Called from worker threads only."""
        self.events.put((kind, args))

    def process_events(self):
        try:
            while True:
                kind, args = self.events.get_nowait()
                if kind == "log":
                    self.log(args[0])
                elif kind == "progress":
                    done, total = args
                    self.progress['value'] = (done / total * 100) if total else 100
                    self.progress_label.config(text=f"{done} / {total}")
                elif kind == "languages":
                    self.set_languages(args[0])
                elif kind == "done":
                    self.finish_translation(*args)
        except queue.Empty:
            pass
        self.root.after(100, self.process_events)

    def refresh_languages(self):
        key = self.api_key.get().strip()

        def work():
            try:
                self.post("languages", dt.fetch_target_languages(dt.make_client(key)))
            except Exception as e:
                self.post("log", f"Could not load language list from DeepL ({e}); using built-in list.")

        threading.Thread(target=work, daemon=True).start()

    # ---------------------------------------------------------- translation

    def start_translation(self):
        if self.is_translating:
            messagebox.showwarning("Warning", "Translation already in progress!")
            return
        if not self.api_key.get().strip():
            messagebox.showerror("Error", "Please enter your DeepL API key!")
            return
        if not self.file_path.get():
            messagebox.showerror("Error", "Please select a Word document!")
            return
        if not self.selected_code():
            messagebox.showerror("Error", "Please select a target language!")
            return

        self.is_translating = True
        self.translate_btn.config(state='disabled', text="Translating...")
        self.progress['value'] = 0
        self.progress_label.config(text="0 / 0")
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

        args = (Path(self.file_path.get()), self.selected_code(), self.api_key.get().strip())
        threading.Thread(target=self.translate_document, args=args, daemon=True).start()

    def translate_document(self, input_path, target_lang, api_key):
        """Runs in a worker thread: no widget access here, only self.post()."""
        output_path = dt.output_path_for(input_path, target_lang)
        try:
            self.post("log", f"Translating {input_path.name} to {target_lang}...")
            self.post("log", "=" * 60)
            client = dt.make_client(api_key)
            stats = dt.translate_docx(
                input_path, output_path, target_lang, client,
                progress=lambda done, total: self.post("progress", done, total),
                log=lambda msg: self.post("log", msg),
            )
            self.post("log", "=" * 60)
            self.post("log", f"Translation complete ({stats['characters']:,} characters sent).")
            self.post("log", f"Saved to: {output_path}")
            self.post("done", output_path, None)
        except Exception as e:
            if output_path.exists():
                try:
                    output_path.unlink()  # don't leave a half-written file behind
                except OSError:
                    pass
            self.post("log", f"ERROR: {e}")
            self.post("done", None, str(e))

    def finish_translation(self, output_path, error):
        self.is_translating = False
        self.translate_btn.config(state='normal', text="Start Translation")
        if error:
            messagebox.showerror("Error", f"Translation failed:\n{error}")
        else:
            self.progress['value'] = 100
            messagebox.showinfo("Success", f"Translation complete!\n\nSaved to:\n{output_path.name}")


def main():
    root = tk.Tk()
    TranslatorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
