<div align="center">

# 🧹 Dynamic Data Cleaning Automation Tool using AI with Python

**An Interactive Desktop GUI Application for Automated Excel Preprocessing using Python, Tkinter & Pandas**

[![Python](https://img.shields.io/badge/Python-3.13%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-Data_Manipulation-150458?style=for-the-badge&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![Tkinter](https://img.shields.io/badge/Tkinter-Desktop_GUI-FF6F00?style=for-the-badge&logo=python&logoColor=white)](https://docs.python.org/3/library/tkinter.html)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel_Engine-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](https://openpyxl.readthedocs.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-0d9488?style=for-the-badge)](https://opensource.org/licenses/MIT)

[📖 Project Overview](#-project-overview) •
[✨ Key Features](#-key-features) •
[🔄 Workflow](#-how-it-works-workflow) •
[📊 Case Study](#-before--after-data-transformation) •
[🚀 Getting Started](#-getting-started) •
[📬 Contact](#-author--contact)

</div>

---

## 📖 Project Overview

In real-world data analytics workflows, raw Excel spreadsheets extracted from ERP systems, CRM platforms, and manual entry forms frequently suffer from structural inconsistencies—such as duplicate entries, completely blank rows, trailing/leading whitespaces, and erratic text capitalization.

The **Dynamic Data Cleaning Automation Tool** is a lightweight, zero-code desktop utility built with **Python**, **Tkinter**, and **Pandas**. It empowers both data analysts and non-technical stakeholders to browse local Excel files (`.xlsx` / `.xls`), toggle customizable preprocessing rules via an intuitive interface, and export clean, analysis-ready datasets in a single click.

### 🎯 Core Objectives
- **Eliminate Repetitive Effort:** Automate routine spreadsheet cleaning without writing ad-hoc scripts or Excel formulas for every new file.
- **Ensure Non-Destructive Processing:** Preserve original raw data by automatically exporting processed files with a `Cleaned_` prefix.
- **Type-Safe String Normalization:** Dynamically detect text (`object`) columns so numeric quantities, IDs, and dates remain untouched.

---

## ✨ Key Features

- 📂 **Interactive File & Directory Browser:** Native OS file dialogs (`filedialog.askopenfilename` & `askdirectory`) to select `.xlsx`/`.xls` files and destination folders effortlessly.
- 🧹 **Modular One-Click Cleaning Options:**
  - **Remove Blank Rows:** Purges records where all column values are empty (`dropna(how='all')`) while safely retaining partial records.
  - **Remove Duplicate Rows:** Eliminates redundant transaction entries across the dataset (`drop_duplicates()`).
  - **Trim Leading & Trailing Spaces:** Strips hidden whitespaces exclusively from text columns (`.str.strip()`).
  - **Title Case Standardization:** Normalizes inconsistent casing (e.g., `"laptop"`, `"LAPTOP"` $\rightarrow$ `"Laptop"`) across text fields (`.str.title()`).
- ⚡ **Quick Reset Controls:** Dedicated buttons to clear file paths or uncheck all cleaning options for rapid multi-file batch runs.
- 🛡️ **Defensive Error Handling:** Built-in validation alerts (`messagebox.showerror`) for missing file paths, permission locks, or corrupted workbooks.

---

## 🛠️ Tech Stack

| Technology / Library | Role in Application |
| :--- | :--- |
| **Python 3.x** | Core programming language managing application logic and state binding |
| **Tkinter** | Desktop Graphical User Interface (`650x480` responsive layout frames, checkboxes, dialogs) |
| **Pandas** | High-performance DataFrame ingestion, `object` dtype filtering, and vectorized cleaning |
| **Openpyxl** | Backend I/O engine for reading and writing Microsoft Excel workbooks |
| **OS Module** | Cross-platform file path resolution (`os.path.basename`, `os.path.join`) |

---

## 🔄 How It Works (Workflow)

```text
┌──────────────────────┐     ┌──────────────────────┐     ┌──────────────────────┐
│  1. Select Excel File│ ──► │ 2. Choose Output Dir │ ──► │ 3. Toggle Checkboxes │
│   (.xlsx / .xls)     │     │   (Target Folder)    │     │ (Cleaning Rules 1-4) │
└──────────────────────┘     └──────────────────────┘     └──────────┬───────────┘
                                                                     │
┌──────────────────────┐     ┌──────────────────────┐                ▼
│  6. Auto-Save Output │ ◄── │ 5. Pandas Processing │ ◄── ┌──────────────────────┐
│ Cleaned_<filename>   │     │ dropna/dedup/str ops │     │ 4. Click 'Clean Data'│
└──────────────────────┘     └──────────────────────┘     └──────────────────────┘
```

---

## 📊 Before & After Data Transformation

Below is an empirical verification using the included sample dataset (`raw_sales_data.xlsx` $\rightarrow$ `Cleaned_raw_sales_data.xlsx`):

### ❌ Raw Input (`raw_sales_data.xlsx`)
*Contains leading/trailing spaces, mixed casing, a completely blank row, and duplicate records.*

| Order_ID | Customer_Name | Product | Region | Price |
| :---: | :---: | :---: | :---: | :---: |
| `101` | `" Alice "` | `"laptop"` | `"north"` | `1200` |
| `102` | `"Bob"` | `" Mouse "` | `"SOUTH"` | `25` |
| `103` | `"charlie"` | `"KEYBOARD"` | `" East"` | `45` |
| `101` | `" Alice "` | `"laptop"` | `"NORTH"` | `1200` |
| `NaN` | `NaN` | `NaN` | `NaN` | `NaN` |
| `104` | `"David "` | `"monitor"` | `"west "` | `300` |
| `105` | `"Eva"` | `"Headphones"` | `NaN` | `80` |
| `106` | `" frank"` | `"webcam"` | `"North "` | `NaN` |
| `102` | `"Bob"` | `" Mouse "` | `"SOUTH"` | `25` |

### ✅ Cleaned Output (`Cleaned_raw_sales_data.xlsx`)
*Blank row purged, duplicate Order 102 removed, whitespace stripped, and text standardized to Title Case.*

| Order_ID | Customer_Name | Product | Region | Price |
| :---: | :---: | :---: | :---: | :---: |
| `101` | `"Alice"` | `"Laptop"` | `"North"` | `1200` |
| `102` | `"Bob"` | `"Mouse"` | `"South"` | `25` |
| `103` | `"Charlie"` | `"Keyboard"` | `"East"` | `45` |
| `101` | `"Alice"` | `"Laptop"` | `"North"` | `1200` |
| `104` | `"David"` | `"Monitor"` | `"West"` | `300` |
| `105` | `"Eva"` | `"Headphones"` | `NaN` | `80` |
| `106` | `"Frank"` | `"Webcam"` | `"North"` | `NaN` |

---

## 📂 Repository Structure

```bash
📦 Dynamic-Data-Cleaning-Automation-Tool
 ┣ 📜 GUI-Based Automation Using AI with Python.ipynb  # Main Jupyter Notebook source code
 ┣ 📜 app.py                                           # Standalone Python script (optional extraction)
 ┣ 📊 raw_sales_data.xlsx                              # Sample uncleaned input dataset
 ┣ 📊 Cleaned_raw_sales_data.xlsx                      # Sample processed output dataset
 ┣ 📄 Dynamic Data Cleaning Automation Tool.docx       # Project documentation
 ┗ 📜 README.md                                        # Repository documentation
```

---

## 🚀 Getting Started

### 1. Prerequisites
Ensure you have **Python 3.8+** installed on your system. `tkinter` and `os` come pre-installed with standard Python distributions.

### 2. Clone the Repository
```bash
git clone https://github.com/your-username/dynamic-data-cleaning-automation.git
cd dynamic-data-cleaning-automation
```

### 3. Install Dependencies
Install `pandas` and `openpyxl` via `pip`:
```bash
pip install pandas openpyxl
```

### 4. Run the Application
You can launch the GUI directly from **Jupyter Notebook** or as a Python script:

**Option A — Via Jupyter Notebook:**
```bash
jupyter notebook "GUI-Based Automation Using AI with Python.ipynb"
```

**Option B — Via Terminal (if saved as `app.py`):**
```bash
python app.py
```

---

## 🔮 Future Enhancements Roadmap

- [ ] **Interactive Data Preview:** Embed a `ttk.Treeview` table inside the GUI to preview top rows before and after cleaning.
- [ ] **Pipeline Order Toggle:** Execute whitespace trimming and Title Case normalization prior to deduplication so case-variant duplicates (e.g., `"north"` vs `"NORTH"`) merge automatically.
- [ ] **Visual Summary Dashboard:** Generate quick post-cleaning metrics (rows removed, nulls remaining, column data types).
- [ ] **Streamlit Web App & Standalone `.exe`:** Deploy a browser version using Streamlit and package a zero-dependency Windows executable via `PyInstaller`.

---

## 📬 Author & Contact

<div align="left">

**Anurag Kumar Poddar**  
*Data Analyst*  
📞 **Phone:** +91 9748748405  

</div>

---
<div align="center">
  <i>If you found this project useful, please consider giving the repository a ⭐!</i>
</div>
