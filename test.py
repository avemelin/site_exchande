import os
import fitz
import re
import openpyxl
from flask import Flask, render_template, request, send_file

app = Flask(__name__)


@app.route("/", methods=["GET", "POST"])
def upload_and_process():
    if request.method == "POST":
        uploaded_file = request.files["file"]
        if uploaded_file.filename != "":
            # Сохраняем загруженный файл
            pdf_path = os.path.join("uploads", uploaded_file.filename)
            uploaded_file.save(pdf_path)
            pdf_document = fitz.open(pdf_path)
            num_pages = len(pdf_document)
            num_rows = 0

            first_page = pdf_document[0]

            # Извлекаем текст с первой страницы
            text = first_page.get_text()

            # Разбиваем текст на строки
            lines = text.split('\n')

            # Получаем первую строку (если она существует)
            lin = []
            for i in range(29, len(lines) - 1):
                lin.append(lines[i])

            print(lin)
            patternnn = r'\d{6}'
            pattern = r'\d{6}'
            pattern_1 = r'Перевод'
            pattern_2 = r'Прочие операции'
            pattern_3 = r'Неизвестная'
            pattern_4 = r'Прочие расходы'

            # Создаем новую рабочую книгу
            workbook = openpyxl.Workbook()

            # Выбираем активный лист (по умолчанию это лист с именем "Sheet")
            sheet = workbook.active

            # записывается в столбец А дата
            date = []
            for i in range(2, len(lin) - 1):
                if re.match(patternnn, lin[i]):
                    date.append(lin[i - 1])
            print(date)
            for g in range(len(date)):
                sheet[f'A{g + 1}'] = date[g]

            # записывает в В1-10 текст
            text = []
            for i in range(len(lin) - 1):
                if re.match(patternnn, lin[i]):
                    text.append(lin[i + 1])

            for g in range(len(text)):
                sheet[f'B{g + 1}'] = text[g]

            # записывается в столбец С сумма операции
            sum = []
            sum_end = []
            for i in range(len(lin) - 1):
                if re.match(pattern_1, lin[i]) or re.match(pattern_2, lin[i]) or re.match(pattern_3,
                                                                                              lin[i]) or re.match(
                            pattern_4, lin[i]):
                    sum.append(lin[i + 1])
            for k in sum:
                k1 = k.replace('\xa0', ' ')
                sum_end.append(k1)

            for g in range(len(sum_end)):
                sheet[f'C{g + 1}'] = sum_end[g]
            # Циклом проходимся по всем страницам документа, начиная со второй
            for page in range(1, num_pages):
                first_page = pdf_document[page]

                # Извлекаем текст с первой страницы
                text = first_page.get_text()

                # Разбиваем текст на строки
                lines = text.split('\n')

                # Получаем первую строку (если она существует)
                lin = []
                for i in range(7, len(lines) - 1):
                    lin.append(lines[i])




                # записывается в столбец А дата
                date = []
                for i in range(2, len(lin) - 1):
                    if re.match(pattern, lin[i]):
                        date.append(lin[i - 1])

                for g in range(len(date)):
                    sheet[f'A{num_rows + g + 1}'] = date[g]

                # записывает в В1-10 текст
                text = []
                for i in range(len(lin) - 1):
                    if re.match(pattern, lin[i]):
                        text.append(lin[i + 1])
                print(text)
                for g in range(len(text)):
                    sheet[f'B{num_rows + g + 1}'] = text[g]

                # записывается в столбец С сумма операции
                sum = []
                sum_end = []
                for i in range(len(lin) - 1):
                    if re.match(pattern_1, lin[i]) or re.match(pattern_2, lin[i]) or re.match(pattern_3,
                                                                                              lin[i]) or re.match(
                            pattern_4, lin[i]):
                        sum.append(lin[i + 1])
                for k in sum:
                    k1 = k.replace('\xa0', ' ')
                    sum_end.append(k1)

                for g in range(len(sum_end)):
                    sheet[f'C{num_rows + g + 1}'] = sum_end[g]
                sheet.title = f'Лист {page}'
                num_rows = sheet.max_row

                workbook.save('example_2.xlsx')

            result_path = os.path.join("uploads", "result.xlsx")
            workbook.save(result_path)
            return send_file(result_path, as_attachment=True)

    return render_template("upload.html")


if __name__ == "__main__":
    app.run()
