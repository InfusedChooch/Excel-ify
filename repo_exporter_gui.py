#!/usr/bin/env python3
# repo_exporter_gui.py

import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from repo_exporter import build_excel


class RepoExporterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Repo Exporter to Excel")
        self.root.geometry("500x200")
        self.root.resizable(False, False)

        self.source_folder = tk.StringVar()
        self.output_folder = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        pad = {"padx": 10, "pady": 5}

        tk.Label(self.root, text="📂 Source Folder").grid(row=0, column=0, sticky="w", **pad)
        tk.Entry(self.root, textvariable=self.source_folder, width=50).grid(row=1, column=0, columnspan=2, **pad)
        tk.Button(self.root, text="Browse", command=self.select_source).grid(row=1, column=2, **pad)

        tk.Label(self.root, text="📁 Destination Folder").grid(row=2, column=0, sticky="w", **pad)
        tk.Entry(self.root, textvariable=self.output_folder, width=50).grid(row=3, column=0, columnspan=2, **pad)
        tk.Button(self.root, text="Browse", command=self.select_destination).grid(row=3, column=2, **pad)

        tk.Button(self.root, text="Run Export", command=self.run_export,
                  bg="green", fg="white", height=2).grid(row=4, column=0, columnspan=3, **pad)

    def select_source(self):
        folder = filedialog.askdirectory(title="Choose Project Folder")
        if folder:
            self.source_folder.set(folder)

    def select_destination(self):
        folder = filedialog.askdirectory(title="Choose Output Folder")
        if folder:
            self.output_folder.set(folder)

    def run_export(self):
        src = self.source_folder.get()
        out = self.output_folder.get()

        if not src or not out:
            messagebox.showerror("Missing Information", "Please select both source and destination folders.")
            return

        try:
            src_path = Path(src)
            dest_file = Path(out) / "project_source_export.xlsx"
            build_excel(src_path, dest_file)
            messagebox.showinfo("Success", f"Export complete:\n{dest_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = RepoExporterGUI(root)
    root.mainloop()
