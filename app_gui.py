import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import mm
from svglib.svglib import svg2rlg

import json


def generate_pdf(product_name, article_number, maker, logo_path="", file_name='output.pdf'):
    """Формирование pdf наклеек"""

    # Размер страницы 60x60 мм
    page_width, page_height = 58 * mm, 60 * mm

    # Создаем документ с нужными размерами
    doc = SimpleDocTemplate(file_name, pagesize=(page_width, page_height),
                            rightMargin=3 * mm, leftMargin=3 * mm, topMargin=2 * mm, bottomMargin=1 * mm)

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

    # Добавляем отступ сверху перед логотипом
    elements.append(Spacer(1, 5 * mm))
    # Добавляем логотип (PNG)
    if logo_path:
        if '.png' in logo_path:
            logo_width = 20 * mm  # Максимальная ширина логотипа
            logo = Image(logo_path)  # Высота вычисляется автоматически
            logo.hAlign = 'CENTER'  # Выравнивание по центру
            elements.append(logo)
            # Генерируем PDF с элементами
            doc.build(elements)
        elif '.svg' in logo_path:
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
        else:
            print(f"Ошибка в наименовании логотипа {logo_path}")


def read_json(file_name: str, file_path: str = ""):
    with open(f"{file_path}{file_name}.json", "r", encoding="utf-8") as r:
        return json.load(r)


# Основная функция
def process_marketplace(marketplace, logo_path, file_path):
    data_file = read_json(file_name="info_codes", file_path="data/")
    df = pd.read_excel(file_path)
    print(df)
    result = df.groupby(["DETAIL_NUM", "Марка"], as_index=False).agg({
        "ORDER_Q_TY": "sum",
    })
    print(result)
    for i in range(len(result)):
        result.loc[i, 'DETAIL_NUM'] = result['DETAIL_NUM'][i].replace("'", "")
        if result['DETAIL_NUM'][i] in data_file[result['Марка'][i]]:
            name = data_file[result['Марка'][i]][result['DETAIL_NUM'][i]]
            maker = result['Марка'][i]
            for n in range(result['ORDER_Q_TY'][i]):
                generate_pdf(name, result['DETAIL_NUM'][i], maker, logo_path=logo_path,
                             file_name=f"output_img/{maker}_{result['DETAIL_NUM'][i]}_{n}.pdf")
        else:
            messagebox.showwarning("Ошибка", f"Ключа {result['DETAIL_NUM'][i]} нет в словаре марки {result['Марка'][i]}")
    # Итоговый месседжбокс в случае успешного завершения работы программы
    messagebox.showinfo("Информация", f"Обработка для {marketplace}\nЗакончена успешно!!!")


# Функция обработки после выбора
def start_process():
    # Получаем выбранный маркетплейс
    marketplace = marketplace_var.get()
    if not marketplace:
        messagebox.showwarning("Предупреждение", "Выберите маркетплейс!")
        return

    # Задаём логотип в зависимости от маркетплейса
    logo_paths = {
        "Автодок": "static/img/logo_autodoc.svg",
        "Автостелс": "static/img/logo_autostels.png",
        "АвтоТО": "static/img/logo_autoto.svg"
    }
    logo_path = logo_paths.get(marketplace)

    # Открываем окно выбора файла
    file_path = filedialog.askopenfilename(title=f"Выберите файл для {marketplace}")
    if not file_path:
        messagebox.showwarning("Предупреждение", "Файл не выбран!")
        return

    # Вызываем основную функцию
    process_marketplace(marketplace, logo_path, file_path)


# Создаём главное окно
root = tk.Tk()
root.title("Выбор маркетплейса")

# Переменная для хранения выбора
marketplace_var = tk.StringVar()

# Метка
label = tk.Label(root, text="Выберите маркетплейс:", font=("Arial", 14))
label.pack(pady=10)

# Выпадающий список
marketplace_dropdown = ttk.Combobox(
    root, textvariable=marketplace_var, values=["Автодок", "Автостелс", "АвтоТО"], state="readonly", font=("Arial", 12)
)
marketplace_dropdown.pack(pady=5)

# Кнопка для запуска обработки
process_button = tk.Button(root, text="Выбрать и продолжить", command=start_process, font=("Arial", 12))
process_button.pack(pady=20)

# Запуск цикла Tkinter
root.mainloop()
