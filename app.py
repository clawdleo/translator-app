"""
Document Translator - Standalone Desktop App
Translates PPTX and DOCX files using DeepL API
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

# Add bundled modules path for PyInstaller
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))

from translator import Translator
from pptx_processor import PPTXProcessor
from docx_processor import DOCXProcessor

# DeepL API key (free tier)
DEEPL_API_KEY = 'e87352a7-9518-4019-bb38-73f09eb2581b:fx'

LANGUAGES = {
    'Slovenian': 'sl',
    'Croatian': 'hr',
    'Serbian': 'sr',
    'German': 'de',
    'French': 'fr',
    'Spanish': 'es',
    'Italian': 'it',
    'English': 'en',
}


class TranslatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Translator")
        self.root.geometry("500x400")
        self.root.resizable(False, False)
        
        # Configure style
        self.style = ttk.Style()
        self.style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'))
        self.style.configure('Status.TLabel', font=('Segoe UI', 10))
        
        self.selected_file = None
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title = ttk.Label(main_frame, text="📄 Document Translator", style='Title.TLabel')
        title.pack(pady=(0, 20))
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="Select File", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.file_label = ttk.Label(file_frame, text="No file selected", wraplength=400)
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        browse_btn = ttk.Button(file_frame, text="Browse...", command=self.browse_file)
        browse_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Language selection
        lang_frame = ttk.LabelFrame(main_frame, text="Translate to", padding="10")
        lang_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.lang_var = tk.StringVar(value="Slovenian")
        lang_combo = ttk.Combobox(lang_frame, textvariable=self.lang_var, 
                                   values=list(LANGUAGES.keys()), state="readonly", width=30)
        lang_combo.pack(fill=tk.X)
        
        # Translate button
        self.translate_btn = ttk.Button(main_frame, text="Translate Document", 
                                         command=self.start_translation)
        self.translate_btn.pack(pady=(0, 15))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready", style='Status.TLabel')
        self.status_label.pack()
        
        # Supported formats info
        info_label = ttk.Label(main_frame, text="Supports: .pptx (PowerPoint) and .docx (Word)",
                               foreground="gray")
        info_label.pack(side=tk.BOTTOM, pady=(20, 0))
        
    def browse_file(self):
        filetypes = [
            ("Supported files", "*.pptx *.docx"),
            ("PowerPoint", "*.pptx"),
            ("Word", "*.docx"),
            ("All files", "*.*")
        ]
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            self.selected_file = filename
            # Show just filename, not full path
            display_name = Path(filename).name
            size_mb = os.path.getsize(filename) / 1024 / 1024
            self.file_label.config(text=f"{display_name} ({size_mb:.1f} MB)")
            
    def start_translation(self):
        if not self.selected_file:
            messagebox.showwarning("No file", "Please select a file first.")
            return
            
        if not os.path.exists(self.selected_file):
            messagebox.showerror("Error", "Selected file no longer exists.")
            return
            
        # Disable button and start progress
        self.translate_btn.config(state='disabled')
        self.progress.start(10)
        self.status_label.config(text="Translating... This may take a few minutes for large files.")
        
        # Run translation in background thread
        thread = threading.Thread(target=self.do_translation)
        thread.daemon = True
        thread.start()
        
    def do_translation(self):
        try:
            input_path = Path(self.selected_file)
            lang_code = LANGUAGES[self.lang_var.get()]
            
            # Generate output filename
            output_path = input_path.parent / f"{input_path.stem}_{self.lang_var.get()}{input_path.suffix}"
            
            # Initialize translator
            self.update_status("Initializing translator...")
            translator = Translator(target_lang=lang_code, deepl_api_key=DEEPL_API_KEY)
            
            # Process based on file type
            ext = input_path.suffix.lower()
            
            if ext == '.pptx':
                self.update_status("Processing PowerPoint...")
                processor = PPTXProcessor(translator, status_callback=self.update_status)
                processor.process_file(str(input_path), str(output_path))
            elif ext == '.docx':
                self.update_status("Processing Word document...")
                processor = DOCXProcessor(translator, status_callback=self.update_status)
                processor.process_file(str(input_path), str(output_path))
            else:
                raise ValueError(f"Unsupported file type: {ext}")
            
            # Success
            self.translation_complete(True, str(output_path))
            
        except Exception as e:
            self.translation_complete(False, str(e))
            
    def update_status(self, message):
        self.root.after(0, lambda: self.status_label.config(text=message))
        
    def translation_complete(self, success, result):
        def update_ui():
            self.progress.stop()
            self.translate_btn.config(state='normal')
            
            if success:
                self.status_label.config(text="✓ Translation complete!")
                messagebox.showinfo("Success", f"Translated file saved:\n\n{result}")
            else:
                self.status_label.config(text="✗ Translation failed")
                messagebox.showerror("Error", f"Translation failed:\n\n{result}")
                
        self.root.after(0, update_ui)


def main():
    root = tk.Tk()
    
    # Set icon if available
    try:
        if sys.platform == 'win32':
            root.iconbitmap(default='')
    except:
        pass
    
    app = TranslatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
