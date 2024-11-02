from flask import Flask, render_template, request, redirect, url_for
import csv
import io

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/result', methods=['POST'])
def result():
    if request.method == 'POST':
        file1 = request.files['file1']
        file2 = request.files['file2']

        # Перевірка першого обраного файлу та встановлення file1_type
        file1_name = file1.filename
        file1_startswith_serviceOrder = file1_name.startswith("serviceOrder")
        file1_startswith_repair_data = file1_name.startswith("repair_data")

        if file1_startswith_serviceOrder:
            file1_type = "ІТ4"
        elif file1_startswith_repair_data:
            file1_type = "GSX"
        else:
            return "Перший обраний файл повинен починатися на 'serviceOrder' або 'repair_data'."

        # Перевірка другого обраного файлу
        file2_name = file2.filename
        file2_startswith_orderreturn = file2_name.startswith("orderreturn")

        if not file2_startswith_orderreturn:
            return "Другий обраний файл повинен починатися на 'orderreturn'."

        # Зчитуємо дані з обох файлів
        data1 = []
        data2 = []

        with io.TextIOWrapper(file1, encoding='utf-8', newline='') as text_file:
            csv_reader = csv.reader(text_file)
            for row in csv_reader:
                data1.append(row)
        print(len(data1))
        with io.TextIOWrapper(file2, encoding='utf-8', newline='') as text_file:
            csv_reader = csv.reader(text_file)
            for row in csv_reader:
                data2.append(row)

        # Отримуємо заголовки та дані
        headers1 = data1[0]
        headers2 = data2[0]

        # Створюємо таблицю з новими заголовками
        new_headers = [
            "Spare Part Number",
            "Spare Part description",
            "Asbis invoice",
            "Date of Spare Part received",
            "Cost of Spare part",
            "Repair Number it4profit",
            "Serial Number",
            "ASBIS RMA",
            "GSX Repair number",
            "GSX Dispatch ID",
            "Return order (GSX)",
            "Coverage status",
            "Labor Cost"
        ]

        # Ініціалізуємо таблицю як пустий список
        new_data = []

        # Перебираємо рядки першого файлу та заповнюємо нову таблицю
        # Перебираємо рядки першого файлу та заповнюємо нову таблицю
        for row in data1[1:]:
            if file1_type == "ІТ4":
                # Додавання Serial Number до нової таблиці
                new_row = [row[headers1.index("Serial number of claimed product")]]
            else:
                # Додавання порожнього значення, оскільки Serial Number не визначений для GSX
                new_row = [""]

            # Додавання інших значень з файлу 1 до нового рядку
            new_row.extend(row)

            # Додавання порожніх значень для інших стовпців, які потрібно заповнити
            while len(new_row) < len(new_headers):
                new_row.append("")

            # Додавання рядка до нової таблиці
            new_data.append(new_row)

        # Отримуємо фінальну таблицю
        final_headers = new_headers
        final_data = new_data

        # Відображення результатів
        return render_template('index.html', headers=final_headers, data=final_data)


if __name__ == '__main__':
    app.run(debug=True)
