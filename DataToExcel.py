from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side
import requests
from AccountData import get_cash_info
# Create a new workbook and select the active worksheet
wb = Workbook()
ws = wb.active

# Get data
cash_info = get_cash_info()

# Styles
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
title_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # light green
label_fill = PatternFill(start_color="EAF4EA", end_color="EAF4EA", fill_type="solid")  # more opaque green tint

# Merge title across 3 cells
ws.merge_cells('A1:C1')
cell = ws['A1']
cell.value = "Cash Info"
cell.font = Font(bold=True)
cell.fill = title_fill
cell.border = thin_border

# Custom label mapping
label_map = {
    "total": "Total Account Value",
    "free": "Cash",
    "invested": "Currently Invested",
    "blocked": "Blocked Amount",
    "pieCash": "Cash in Pies",
    "result": "Realised Return",
    "ppl": "Open Position P&L"
}

# Write data
row = 2
for key, label in label_map.items():
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    label_cell = ws.cell(row=row, column=1, value=label)
    value_cell = ws.cell(row=row, column=3, value=cash_info.get(key, "N/A"))

    # Apply fills and borders
    label_cell.fill = label_fill
    value_cell.fill = label_fill
    for col in range(1, 4):
        ws.cell(row=row, column=col).border = thin_border

    row += 1

# Optional: Adjust column widths
ws.column_dimensions['A'].width = 10
ws.column_dimensions['B'].width = 10
ws.column_dimensions['C'].width = 10

# Save the workbook
wb.save("AccountAnalysis.xlsx")