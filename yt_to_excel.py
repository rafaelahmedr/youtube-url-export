import json
import subprocess
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Channel URL
channel_url = "https://www.youtube.com/@CotesAS"

# Run yt-dlp and capture JSON lines
result = subprocess.run(
    ["python", "-m", "yt_dlp", "-j", "--flat-playlist", channel_url],
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True,
    encoding="utf-8",
)

# Parse JSON lines
videos = []
for line in result.stdout.splitlines():
    try:
        v = json.loads(line)
        title = v.get("title", "")
        video_id = v.get("id", "")
        url = f"https://www.youtube.com/watch?v={video_id}"
        videos.append((title, url))
    except:
        continue

# Create Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "YouTube Videos"

# Headers
ws.append(["Title", "URL"])
for cell in ws[1]:
    cell.font = Font(bold=True)

# Add rows
for title, url in videos:
    ws.append([title, url])

# Make URLs clickable
for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
    cell = row[0]
    cell.hyperlink = cell.value
    cell.style = "Hyperlink"

# Adjust column widths
for col in ws.columns:
    max_length = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    ws.column_dimensions[col_letter].width = min(max_length + 2, 80)

# Save Excel
output_file = r"C:\Users\RAH\Downloads\videos.xlsx"
wb.save(output_file)

print(f"Done! Saved as {output_file}")
