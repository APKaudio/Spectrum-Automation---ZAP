import pdfplumber
import csv
import re
import os
import webbrowser

input_pdf = "SB PDF.pdf"
output_csv = "SB PDF.csv"

entries = []
last_known_group = None

with pdfplumber.open(input_pdf) as pdf:
    for page in pdf.pages:
        # Read and clean the page text for group names
        lines = page.extract_text().splitlines()
        lines = [line.strip() for line in lines if line.strip()]

        # Find group headers like "Simple Plan IEM's (8 frequencies)"
        group_headers = [(i, line) for i, line in enumerate(lines)
                         if re.match(r".+\(\d+ frequencies\)", line)]

        # Extract tables visually
        tables = page.extract_tables()

        group_index = 0
        for table in tables:
            # Use next group header if available, otherwise carry forward
            if group_index < len(group_headers):
                last_known_group = group_headers[group_index][1]
                group_index += 1

            group_name = last_known_group

            for row in table:
                # Skip empty or all-null rows
                if not row or all(cell is None or cell.strip() == "" for cell in row):
                    continue

                # Skip headers like "Model Band Name Preset Spacing Frequency"
                if "Model" in row[0] and "Frequency" in row[-1]:
                    continue

                # Clean and flatten the row
                clean_row = [cell.replace("\n", " ").strip() if cell else "" for cell in row]
                while len(clean_row) < 6:
                    clean_row.append("")

                model, band, name, preset, spacing, frequency = clean_row

                # Skip rows that mistakenly repeat the group name
                if model.strip() == group_name.strip():
                    continue

                entries.append({
                    "Group": group_name,
                    "Model": model,
                    "Band": band,
                    "Name": name,
                    "Preset": preset,
                    "Spacing": spacing,
                    "Frequency": frequency
                })

# Optional: Deduplicate rows
seen = set()
unique_entries = []
for e in entries:
    row_key = tuple(e.values())
    if row_key not in seen:
        seen.add(row_key)
        unique_entries.append(e)

# Write to CSV
with open(output_csv, "w", newline="", encoding="utf-8") as csvfile:
    fieldnames = ["Group", "Model", "Band", "Name", "Preset", "Spacing", "Frequency"]
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(unique_entries)

print(f"✅ Done. Wrote {len(unique_entries)} clean rows to '{output_csv}'")

# Open the CSV in the default viewer
csv_path = os.path.abspath(output_csv)
webbrowser.open(f'file://{csv_path}')
