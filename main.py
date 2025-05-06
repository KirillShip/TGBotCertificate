import os
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, CallbackContext
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
import random
import textwrap
import zipfile

# Регистрируем профессиональные шрифты
pdfmetrics.registerFont(TTFont('Geometria', 'Geometria-Medium.ttf'))
pdfmetrics.registerFont(TTFont('Geometria-Bold', 'Geometria-Bold.ttf'))

TOKEN = "7543427762:AAEI6rB_iOpdfl8W6SRb7579ux9PkGSw7Nc"

# Глобальная переменная для хранения типа сертификата
CERT_TYPE = None


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    keyboard = [["Сертификат выпускника", "Сертификат о прохождении модулей"]]
    reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True)
    await update.message.reply_text("Выберите тип сертификата:",
                                    reply_markup=reply_markup)


async def choose_cert_type(update: Update,
                           context: ContextTypes.DEFAULT_TYPE) -> None:
    global CERT_TYPE
    cert_type = update.message.text
    if cert_type == "Сертификат выпускника":
        CERT_TYPE = "graduate"
        await update.message.reply_text(
            "Выбран сертификат выпускника. Теперь отправьте Excel-файл с данными.\n"
            "Обязательные столбцы:\n"
            "- ФИО\n"
            "- Список модулей (через запятую)\n"
            "- Дата выпуска (дд.мм.гггг)")
    elif cert_type == "Сертификат о прохождении модулей":
        CERT_TYPE = "module"
        await update.message.reply_text(
            "Выбран сертификат о прохождении модулей. Теперь отправьте Excel-файл с данными.\n"
            "Обязательные столбцы:\n"
            "- ФИО\n"
            "- Список модулей (через запятую)\n"
            "- Дата выпуска (дд.мм.гггг)")
    else:
        await update.message.reply_text(
            "Пожалуйста, выберите тип сертификата из предложенных вариантов.")


async def handle_excel(update: Update,
                       context: ContextTypes.DEFAULT_TYPE) -> None:
    global CERT_TYPE

    if CERT_TYPE is None:
        await update.message.reply_text(
            "Сначала выберите тип сертификата командой /start")
        return

    if not update.message.document or not update.message.document.file_name.endswith(
            '.xlsx'):
        await update.message.reply_text(
            "Пожалуйста, отправьте файл в формате .xlsx")
        return

    try:
        # Скачиваем файл
        excel_file = await update.message.document.get_file()
        excel_path = "temp_data.xlsx"
        await excel_file.download_to_drive(excel_path)

        # Читаем данные
        df = pd.read_excel(excel_path)
        required_columns = ['ФИО', 'Список модулей', 'Дата выпуска']
        if not all(col in df.columns for col in required_columns):
            raise ValueError(
                f"Файл должен содержать столбцы: {', '.join(required_columns)}"
            )

        # Создаем архив
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w',
                             zipfile.ZIP_DEFLATED) as zip_file:
            for _, row in df.iterrows():
                pdf_data = generate_certificate(name=row['ФИО'],
                                                modules=row['Список модулей'],
                                                date=row['Дата выпуска'],
                                                cert_type=CERT_TYPE)
                zip_file.writestr(
                    f"Сертификат_{row['ФИО'].replace(' ', '_')}.pdf",
                    pdf_data.getvalue())

        # Отправляем архив
        zip_buffer.seek(0)
        await update.message.reply_document(
            document=zip_buffer,
            filename="Сертификаты.zip",
            caption=f"Готово! Создано {len(df)} сертификатов.")

    except Exception as e:
        await update.message.reply_text(f"Ошибка: {str(e)}")
    finally:
        if os.path.exists(excel_path):
            os.remove(excel_path)


def generate_certificate(name: str, modules: str, date: str, cert_type: str):
    # Выбираем шаблон в зависимости от типа сертификата
    template_file = "Сертификат выпускника.pdf" if cert_type == "graduate" else "Сертификат модулей.pdf"
    template = PdfReader(template_file)
    template_page = template.pages[0]

    # Создаем новый слой для изменений
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=[1080, 1200])

    # Заливка старых данных (разные цвета для разных типов)
    if cert_type == "graduate":
        # Цвета для сертификата выпускника
        can.setFillColorRGB(1, 1, 1)
        can.rect(55, 140, 510, 90, stroke=0, fill=1)  # МОДУЛИ
        can.rect(55, 280, 500, 35, stroke=0, fill=1)  # ФИО
        can.rect(670, 157, 80, 20, stroke=0, fill=1)  # ДАТА
        can.setFillColorRGB(0, 0.72, 0.72)  # Бирюзовый
        can.rect(700, 505, 100, 25, stroke=0, fill=1)  # НОМЕР
    else:
        # Цвета для сертификата о прохождении модулей
        can.setFillColorRGB(1, 1, 1)
        can.rect(55, 140, 510, 90, stroke=0, fill=1)  # МОДУЛИ
        can.rect(55, 280, 500, 35, stroke=0, fill=1)  # ФИО
        can.rect(670, 157, 80, 20, stroke=0, fill=1)  # ДАТА
        can.setFillColorRGB(0, 0.3, 0.4)
        can.rect(700, 505, 100, 25, stroke=0, fill=1)  # НОМЕР

    can.setFillColorRGB(0, 0, 0)  # Черный текст

    # ФИО выпускника
    can.setFont("Geometria-Bold", 25)
    can.drawString(55, 287, name)

    # Список модулей с правильным переносом по словам
    can.setFont("Geometria", 18)
    text = can.beginText(55, 220)
    text.setLeading(18)

    modules_text = str(modules)
    if isinstance(modules, str):
        modules_list = [m.strip() + ',' for m in modules.split(',')]
        modules_list[-1] = modules_list[-1].replace(',', '')
    else:
        modules_list = modules

    current_line = ""
    for module in modules_list:
        if len(current_line) + len(module) > 55:
            text.textLine(current_line.strip())
            current_line = module + " "
        else:
            current_line += module + " "

    if current_line:
        text.textLine(current_line.strip())

    can.drawText(text)

    # Дата выдачи
    can.setFont("Geometria-Bold", 12)
    can.drawString(673, 159, date)

    # Номер сертификата
    cert_number = f"N°{random.randint(100000, 999999)}"
    can.setFont("Geometria-Bold", 15)
    if cert_type == "graduate":
        can.setFillColorRGB(1, 1, 1)  # Белый текст для бирюзовой плашки
    else:
        can.setFillColorRGB(1, 1, 1)  # Белый текст для красной плашки
    can.drawString(713, 517, cert_number)

    can.save()

    # Наложение изменений
    overlay = PdfReader(packet)
    template_page.merge_page(overlay.pages[0])

    # Сохраняем результат
    output = PdfWriter()
    output.add_page(template_page)

    result_buffer = BytesIO()
    output.write(result_buffer)
    result_buffer.seek(0)

    return result_buffer


def main():
    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(
        MessageHandler(
            filters.Text(
                ["Сертификат выпускника", "Сертификат о прохождении модулей"]),
            choose_cert_type))
    app.add_handler(
        MessageHandler(
            filters.Document.FileExtension("xlsx") & ~filters.COMMAND,
            handle_excel))

    app.run_polling()


if __name__ == "__main__":
    from telegram import ReplyKeyboardMarkup
    main()
