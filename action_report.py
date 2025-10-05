import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from pathlib import Path

# === Base Configuration ===
script_dir = Path(__file__).resolve().parent
base_dir = script_dir / "data_source"
pictures_base_dir = base_dir / "pictures"
template_path = script_dir / "ReportTemplate.docx"
output_path = script_dir / "CompletedReport.docx"
excel_file = base_dir / "ReportData.xlsx"

# === Load Excel key-value pairs ===
df = pd.read_excel(excel_file)
context = {row["Variable Name"]: row["Variable Value"] for _, row in df.iterrows()}

# === Load Word template ===
doc = DocxTemplate(template_path)

# === Picture folders to process ===
picture_folders = [
    "approved_plan",
    "domestic_water_layout",
    "entrance",
    "sewer_layout",
    "storm_water_layout"
]

# === Supported extensions ===
image_extensions = (".png", ".jpg", ".jpeg")

# === Iterate through each folder ===
for folder_name in picture_folders:
    folder_path = pictures_base_dir / folder_name
    var_name = f"{folder_name}_pictures"

    # Check folder existence
    if not folder_path.exists():
        print(f"⚠️ Folder not found: {folder_path}")
        context[var_name] = None
        continue

    # Gather all valid image files (sorted alphabetically)
    image_files = sorted(
        [f for f in folder_path.glob("*") if f.suffix.lower() in image_extensions],
        key=lambda x: x.name.lower()
    )

    if not image_files:
        print(f"⚠️ No images found in: {folder_name}")
        context[var_name] = None
        continue

    # Convert to InlineImage objects
    image_list = []
    for img_file in image_files:
        try:
            image_list.append(InlineImage(doc, str(img_file), width=Mm(80)))
        except Exception as e:
            print(f"⚠️ Skipped image {img_file}: {e}")

    context[var_name] = image_list if image_list else None

# === Render and save ===
doc.render(context)
doc.save(output_path)

print(f"✅ Report generated successfully: {output_path}")
