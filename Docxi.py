import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import re
import os
import win32com.client

class DocxiApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Script Automation Tool")
        self.root.geometry("1200x900")
        self.root.configure(bg="#ffffff") 

        self.file_path = ""

        # Toolbar part
        self.toolbar = tk.Frame(root, bg="#007bff")
        self.toolbar.pack(side=tk.TOP, fill=tk.X)

        self.select_file_button = tk.Button(self.toolbar, text="📁 Select File", command=self.select_file, bg="#007bff", fg="white", font=("Arial", 14), relief="flat")
        self.select_file_button.pack(side=tk.LEFT, padx=5, pady=5)
        self.select_file_button.bind("<Enter>", lambda e: self.select_file_button.config(bg="#0056b3"))
        self.select_file_button.bind("<Leave>", lambda e: self.select_file_button.config(bg="#007bff"))

        self.save_button = tk.Button(self.toolbar, text="💾 Save As", command=self.save_file, bg="#28a745", fg="white", font=("Arial", 14), relief="flat")
        self.save_button.pack(side=tk.LEFT, padx=5, pady=5)
        self.save_button.bind("<Enter>", lambda e: self.save_button.config(bg="#218838"))
        self.save_button.bind("<Leave>", lambda e: self.save_button.config(bg="#28a745"))

        # File Selection Frame
        self.file_frame = tk.Frame(root, bg="#f0f0f0")
        self.file_frame.pack(pady=10, padx=10, fill=tk.X)

        self.file_label = tk.Label(self.file_frame, text="No file selected", bg="#f0f0f0", font=("Arial", 14))
        self.file_label.pack(side=tk.LEFT, padx=5)

        # Processing Options Frame
        self.options_frame = tk.LabelFrame(root, text="Processing Options", bg="#f0f0f0", font=("Arial", 14))
        self.options_frame.pack(pady=10, padx=10, fill=tk.X)

        self.remove_chinese_var = tk.BooleanVar()
        self.remove_empty_var = tk.BooleanVar()

        self.remove_chinese_var.trace_add("write", lambda *args: self.process_file())
        self.remove_empty_var.trace_add("write", lambda *args: self.process_file())

        self.remove_chinese_checkbox = tk.Checkbutton(self.options_frame, text="Remove Lines with Chinese Characters", variable=self.remove_chinese_var, bg="#f0f0f0", font=("Arial", 12))
        self.remove_chinese_checkbox.pack(pady=5)

        self.remove_empty_checkbox = tk.Checkbutton(self.options_frame, text="Remove Empty Lines", variable=self.remove_empty_var, bg="#f0f0f0", font=("Arial", 12))
        self.remove_empty_checkbox.pack(pady=5)

        # Realtime preview area with scrolling thingie
        self.preview_area = scrolledtext.ScrolledText(root, width=90, height=20, font=("Arial", 12), wrap=tk.WORD, bd=2, relief="groove")
        self.preview_area.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # Status Bar
        self.status_label = tk.Label(root, text="", bg="#f0f0f0", font=("Arial", 12))
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if self.file_path:
            self.file_label.config(text=self.file_path)
            self.process_file()

    def process_file(self):
        with open(self.file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()

        processed_lines = []
        for line in lines:
            if self.remove_chinese_var.get() and re.search(r'[\u4e00-\u9fff]', line):
                continue
            if self.remove_empty_var.get() and not line.strip():
                continue
            processed_lines.append(line)

        self.preview_area.delete(1.0, tk.END)
        self.preview_area.insert(tk.END, ''.join(processed_lines))

    def save_file(self):
        # prefill the save dialog with the text file name for qol
        default_name = os.path.splitext(os.path.basename(self.file_path))[0] + ".docx"
        output_path = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=default_name, filetypes=[("Word Documents", "*.docx")])
        if output_path:
            self.create_docx(output_path)

    def create_docx(self, output_path):
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False 
            doc = word.Documents.Add()

            content = self.preview_area.get(1.0, tk.END).strip()
            if content:
                doc.Content.Text = content

            # Will set font to Arial
            for paragraph in doc.Content.Paragraphs:
                paragraph.Range.Font.Name = "Arial"

            # Save the document
            doc.SaveAs(output_path)
            doc.Close()
            word.Quit()

            self.status_label.config(text="File saved successfully!")
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocxiApp(root)
    root.mainloop()
