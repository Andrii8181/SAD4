from docx import Document
from datetime import datetime

def export_results(data, results):
    doc = Document()
    doc.add_heading("Перевірка на нормальність (Шапіро–Вілка)", 0)

    doc.add_paragraph(f"Дата аналізу: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    doc.add_paragraph("Показник: (введіть назву)")
    doc.add_paragraph("Одиниці виміру: (введіть одиниці)")

    doc.add_heading("Сирі дані", level=1)
    table = doc.add_table(rows=len(data[0])+1, cols=len(data))
    hdr_cells = table.rows[0].cells
    for i in range(len(data)):
        hdr_cells[i].text = f"Фактор {i+1}"
    for row in range(len(data[0])):
        cells = table.add_row().cells
        for col in range(len(data)):
            try:
                cells[col].text = str(data[col][row])
            except IndexError:
                cells[col].text = ""

    doc.add_heading("Результати перевірки", level=1)
    for res in results:
        doc.add_paragraph(
            f"{res['column']}: W={res['W']}, p={res['p']} → {res['result']}"
        )

    doc.save("Результати_аналізу.docx")
