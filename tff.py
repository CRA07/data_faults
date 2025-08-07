from flask import Flask, request
import openpyxl
import os
from datetime import datetime


app = Flask(__name__)
EXCEL_FILE = "brak_report.xlsx"

# Ключевые слова для поиска
PRODUCTS = ["масло монарды", "тоник жасмина", "крем ромашки"]
DEFECTS = ["не тот дозатор", "протек", "не хватает", "сломался", "без крышки"]

# Создаём Excel-файл, если его ещё нет
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Брак"
        sheet.append(["Время", "Автор", "Продукт", "Брак", "Сообщение"])
        wb.save(EXCEL_FILE)

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json

    if not data or "message" not in data:
        return "Invalid", 400

    msg = data["message"]
    text = msg.get("content", "").lower()
    author = msg.get("author", {}).get("name", "Неизвестно")
    timestamp = msg.get("created_at", datetime.now().isoformat())
    time_str = datetime.fromisoformat(timestamp).strftime("%Y-%m-%d %H:%M:%S")

    # Ищем ключевые слова в сообщении
    found_product = next((p for p in PRODUCTS if p in text), "")
    found_defect = next((d for d in DEFECTS if d in text), "")

    # Если найден и продукт, и дефект — записываем в Excel
    if found_product and found_defect:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        sheet = wb.active
        sheet.append([time_str, author, found_product, found_defect, text])
        wb.save(EXCEL_FILE)
        return "Записано", 200

    return "Нет нужной информации", 200

if __name__ == "__main__":
    init_excel()

    app.run(port=2283)
