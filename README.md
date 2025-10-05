# 📘 Report Templating Automation

Automated Word report generation from structured Excel data and organized image folders.  
This project uses **Python**, **docxtpl**, and **python-docx** to fill a Word template (`ReportTemplate.docx`) with dynamic content, including variables, tables, and embedded pictures — producing a professional `CompletedReport.docx` in one command.

---

## 🚀 Features

- 🔁 Auto-fills a Word template with values from `ReportData.xlsx`
- 🖼️ Inserts multiple images from categorized folders (e.g. `entrance`, `sewer_layout`)
- 📊 Imports Excel sheet data as real Word tables (e.g. *Domestic Water*, *Storm Water*)
- 🧹 Cleans up intermediate working files automatically
- 📂 Easily extendable for new data or templates

---

## 📁 Project Structure

```
report_templating/
│
├── create_report.py              # Main Python script to generate the report
├── run_me_once.ps1               # Optional PowerShell helper to set up environment
├── ReportTemplate.docx           # Word template containing placeholders/macros
├── CompletedReport.docx          # Generated output file (auto-created)
├── .gitignore
│
├── data_source/
│   ├── ReportData.xlsx           # Input data file with variables & tables
│   └── pictures/
│       ├── approved_plan/
│       ├── domestic_water_layout/
│       ├── entrance/
│       ├── sewer_layout/
│       └── storm_water_layout/
│           ├── plan001.png
│           ├── layout.png
│           ├── entrance001.png
│           ├── sewer_layout.png
│           └── storm_water001.png
```

---

## ⚙️ Setup Instructions

### 1. 📦 Install Requirements

Make sure you have **Python 3.9+** installed.  
Then install dependencies using pip:

```bash
pip install -r requirements.txt
```

If you don’t have a `requirements.txt`, create one with:

```bash
pip install docxtpl python-docx pandas openpyxl
pip freeze > requirements.txt
```

---

### 2. 🧰 Folder Setup

Your project must contain:

- `ReportTemplate.docx` — the Word template with placeholders like  
  `{{date_of_report}}`, `{{render_images(domestic_water_layout_pictures)}}`,  
  and table insertion markers like `<<DOMESTIC_WATER_TABLE>>`.
- `data_source/ReportData.xlsx` — Excel file with:
  - **General** sheet for variable-value pairs  
  - **DomesticWater** and **StormWater** sheets for table data
- `data_source/pictures/` — folders of categorized images (png/jpg/jpeg).

---

### 3. 🧠 Template Usage

Your Word template uses Jinja-style placeholders from **docxtpl**, e.g.:

```jinja2
{{ date_of_report }}
{{ report_author }}
{{ location_of_report }}
```

For image collections, use:

```jinja2
{{ render_images(domestic_water_layout_pictures) }}
```

For dynamic tables (handled programmatically):

```text
<<DOMESTIC_WATER_TABLE>>
<<STORM_WATER_TABLE>>
```

---

### 4. ▶️ Running the Report

From the project root:

```bash
python create_report.py
```

This will:

1. Read data from `data_source/ReportData.xlsx`
2. Collect and insert all categorized images
3. Fill all placeholders in `ReportTemplate.docx`
4. Build real Word tables for each Excel sheet
5. Save the completed report as `CompletedReport.docx`
6. Automatically delete the intermediate working file

---

### 5. 🧹 Optional PowerShell Helper

If you’re on Windows, you can initialize the project and run the script with:

```powershell
.
un_me_once.ps1
```

This script can be extended to:
- Create your Python virtual environment  
- Install dependencies  
- Launch the report creation process

---

## 🧩 Example Output

The final document (`CompletedReport.docx`) will include:

- Auto-filled header sections (e.g. project name, author, date)
- Embedded and captioned images grouped by category
- Fully formatted tables rendered from your Excel data
- Clean professional layout ready for distribution

---

## 🧱 Customization

### ➕ Adding New Picture Categories

To add another section with images:
1. Create a new folder under `data_source/pictures/` (e.g. `fire_suppression`).
2. Add images (`.png`, `.jpg`, `.jpeg`) inside that folder.
3. In your template, insert a new block:

   ```jinja2
   {{ render_images(fire_suppression_pictures) }}
   ```

4. The Python script automatically detects and populates it.

---

### ➕ Adding New Data Tables

To include a new Excel sheet as a table:
1. Add a new sheet in `ReportData.xlsx` (e.g. `FireWater`).
2. Add your header row and data rows.
3. Insert a marker in Word:  
   ```
   <<FIREWATER_TABLE>>
   ```
4. Update the script to call:
   ```python
   insert_table_at_marker(doc_final, marker="<<FIREWATER_TABLE>>", df=excel["FireWater"])
   ```

---

## 🧠 Tech Overview

| Component | Purpose |
|------------|----------|
| **docxtpl** | Template rendering and variable substitution in Word |
| **python-docx** | Inserting tables programmatically |
| **pandas** | Reading Excel data and converting it to Word-friendly structures |
| **PowerShell** | Optional automation wrapper for Windows users |

---

## 🧾 License

This project is open-source under the **MIT License**.  
You’re free to modify, reuse, and distribute with attribution.

---

## 👨‍💻 Author

**Marius Vorster**  
📍 South Africa  
💼 [GitHub Profile](https://github.com/mariusvrstr)

---

> ✨ *“Automate your documentation — spend your time thinking, not formatting.”*
