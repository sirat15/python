import xlsxwriter

title_list = ["Sr No.", "ID", "Description", "Steps", "Expected Results", "Actual Results"]
row = 6  # Starting row
col = 0  # Starting column

workbook = xlsxwriter.Workbook("creating_workbook_excel.xlsx")
worksheet_data = workbook.add_worksheet("data")
worksheet_analysis = workbook.add_worksheet("analysis")

title_list_format = workbook.add_format(
    {
        "bg_color": "#76ff7b",
        "bold" : True,
        "align" : "center",
        "border" : 1,
        "border_color" : "#000000"

    }
)

# Loop through title_list and write each title in a new column
for index, item in enumerate(title_list):
    worksheet_data.write(row, col + index, item, title_list_format)

worksheet_data.set_column(col, index, 20)
print(worksheet_data)
workbook.close()