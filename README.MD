# 📦 Repo Exporter

`repo_exporter.py` is a drop-in utility that scans any codebase and exports all human-readable source files into a single, well-organized Excel workbook (`.xlsx`). This is ideal for sharing, auditing, documentation, or code review purposes without needing direct access to the codebase.

---

## 🔧 Features

* Recursively finds all readable code/text files in a project
* Skips non-relevant folders (`venv`, `.git`, `__pycache__`, etc.)
* Outputs:

  * 📄 **Summary** tab with file paths, sizes, timestamps
  * 🢾 **README** tab with usage and dependency info
  * 📁 One Excel tab per file, wrapped in code fences
* Built-in support for common formats: `.py`, `.md`, `.json`, `.yaml`, `.csv`, etc.
* Filters out binary files like images or compiled assets
* Supports inclusion of `requirements.txt`, `pyproject.toml`, or `environment.yml`

---

## 🚀 Usage

```bash
python repo_exporter.py [project_root] [-o output.xlsx]
```

### Examples:

Export the current directory:

```bash
python repo_exporter.py
```

Export a specific project folder:

```bash
python repo_exporter.py ../my-project -o exported_code.xlsx
```

---

## 📁 Output Structure

The resulting Excel file includes:

| Sheet Name  | Contents                                              |
| ----------- | ----------------------------------------------------- |
| `Summary`   | A table with file paths, sizes, timestamps, sheet IDs |
| `README`    | Script usage notes and any detected dependency files  |
| `<file>.py` | One sheet per file, with content inside code fences   |

---

## 📌 Notes

* Excel limits sheet names to 31 characters; long paths are truncated safely.
* File content is rendered in monospaced font with gray code-style background.
* The tool ignores binary files and non-source folders by default.
* Intended for local-only use; does not transmit or upload data.

---

## 🧪 Requirements

Install the required libraries with:

```bash
pip install pandas openpyxl
```

---

## 📆 Supported File Types

By default, the script includes these file extensions:

```
.py, .md, .txt, .json, .yaml, .yml, .toml,
.ini, .cfg, .csv, .tsv, .js, .ts, .html, .css,
.sh, .bat, .ps1, .sql
```

You can easily expand this list by editing the `INCLUDE_EXTS` set in the script.

---

## 🔐 License

MIT License — Use freely and modify for your own workflow needs.

---

## 🧠 Author

Crafted by J, based on a concept of fast, portable, and readable repo exports.
