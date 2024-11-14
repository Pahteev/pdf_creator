from reportlab.lib.pagesizes import mm
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import mm
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF

import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog


def generate_pdf(product_name, article_number, maker, logo_path="", file_name='output.pdf'):
    # Размер страницы 60x60 мм
    page_width, page_height = 58 * mm, 60 * mm

    # Создаем документ с нужными размерами
    doc = SimpleDocTemplate(file_name, pagesize=(page_width, page_height),
                            rightMargin=3 * mm, leftMargin=3 * mm, topMargin=3 * mm, bottomMargin=3 * mm)

    # Подключаем шрифт Roboto
    pdfmetrics.registerFont(TTFont('Roboto', 'Roboto-Regular.ttf'))  # Указываем путь к файлу шрифта

    # Получаем стили для текста и устанавливаем шрифт Roboto
    styles = getSampleStyleSheet()
    styleN = styles['Normal']
    styleN.fontName = 'Roboto'  # Используем подключенный шрифт
    styleN.fontSize = 10  # Уменьшаем размер шрифта

    # Создаем объекты Paragraph для каждого текста (с автоматическим переносом)
    elements = []
    elements.append(Paragraph(f"Производитель: {maker}", styleN))
    elements.append(Paragraph(f"Артикул: {article_number}", styleN))
    elements.append(Paragraph(f"Наименование: {product_name}", styleN))
    # elements.append(Paragraph(f"Для: autodoc.ru", styleN))

    # Добавляем отступ сверху перед логотипом
    elements.append(Spacer(1, 10 * mm))  # Отступ 10 мм сверху

    # Чтение и масштабирование SVG логотипа
    drawing = svg2rlg(logo_path)

    # Масштабируем логотип, чтобы его ширина была максимум 40 мм
    scale_factor = (40 * mm) / drawing.width  # Рассчитываем масштабирование по ширине
    drawing.width = 40 * mm
    drawing.height = drawing.height * scale_factor  # Пропорционально изменяем высоту
    drawing.scale(scale_factor, scale_factor)  # Применяем масштабирование
    elements.append(drawing)  # Добавляем логотип

    # Генерируем PDF с элементами
    doc.build(elements)


def get_data_from_catalog():
    df = pd.read_excel("files/Price_autodoc.xlsx")
    data_info = {}
    for i in range(len(df)):
        new_key = df['Артикул'][i]
        data_info[new_key] = {"maker": df['Производитель'][i],
                              "name": df['Наименование'][i]}
    with open("data/data_info.json", "w", encoding="utf-8") as f:
        json.dump(data_info, f, indent=2, ensure_ascii=False)


def read_json(file_name: str, file_path: str = ""):
    with open(f"{file_path}{file_name}.json", "r", encoding="utf-8") as r:
        return json.load(r)


def select_files():
    # Инициализируем главное окно tkinter (оно не будет показано)
    root = tk.Tk()
    root.withdraw()  # Скрываем основное окно

    # Открываем диалоговое окно для выбора файлов (можно выбрать несколько файлов)
    file_paths = filedialog.askopenfilenames(title="Выберите файлы с заказами",
                                             filetypes=[("Excel Files", "*.xlsx *.xls")])

    return file_paths


def read_orders_autodoc(file_path):
    """Открываем заказы автодок"""
    return pd.read_excel(file_path, header=4)


def main():
    data_file = read_json(file_name="data_info", file_path="data/")
    # Вызываем функцию для выбора файлов
    selected_files = select_files()
    df = pd.DataFrame()
    # Выводим выбранные пути файлов
    for file_path in selected_files:
        df_autodoc = read_orders_autodoc(file_path)
        df = pd.concat([df_autodoc, df], ignore_index=True)
    for i in range(len(df)):
        name = data_file[df['Артикул'][i]]['name']
        maker = data_file[df['Артикул'][i]]['maker']
        generate_pdf(name, df['Артикул'][i], maker, logo_path="static/img/logo_autoto.svg", file_name=f"output_img/{df['Артикул'][i]}.pdf")


if __name__ == "__main__":
    main()
