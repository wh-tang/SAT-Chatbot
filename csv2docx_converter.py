import docx
import csv

ori_path = "/Users/Wan Hee/Documents/Academic/2021-2022/Individual Project/Roy/empatheticPersonasZH.csv"
new_path = '/Users/Wan Hee/Documents/Academic/2021-2022/Individual Project/SAT-Chatbot/translation.docx'

doc = docx.Document()

with open(ori_path, newline='', encoding='utf8') as f:
    csv_reader = csv.reader(f) 

    csv_headers = next(csv_reader)
    csv_cols = len(csv_headers)

    table = doc.add_table(rows=2, cols=csv_cols)
    hdr_cells = table.rows[0].cells

    for i in range(csv_cols):
        hdr_cells[i].text = csv_headers[i]

    for row in csv_reader:
        row_cells = table.add_row().cells
        for i in range(csv_cols):
            row_cells[i].text = row[i]

doc.add_page_break()
doc.save(new_path)