import zipfile
import rarfile
import py7zr
import pytesseract
import fitz  # PyMuPDF
import pandas as pd
from docx import Document
from openpyxl import load_workbook
from PIL import Image, ImageDraw
import os
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from concurrent.futures import ThreadPoolExecutor
output_directory_text_images = ""
output_directory_numbering = ""
# Настройка логирования
logging.basicConfig(filename="process.log", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
def extract_data_from_txt(txt_path):
    """
    Извлекает текст из текстового файла (.txt).
    :param txt_path: Путь к текстовому файлу
    :return: Извлечённый текст
    """
    try:
        with open(txt_path, "r", encoding="utf-8") as file:
            data = file.read()
        logging.info(f"Текст успешно извлечён из {txt_path}")
        return data
    except Exception as e:
        logging.error(f"Ошибка при извлечении текста из {txt_path}: {str(e)}")
        return ""
def select_output_directory_for_text_images():
    global output_directory_text_images
    output_directory_text_images = filedialog.askdirectory(title="Выберите директорию для вывода текста и изображений")
    if output_directory_text_images:
        logging.info(f"Директория для текста и изображений установлена: {output_directory_text_images}")

def select_output_directory_for_numbering():
    global output_directory_numbering
    output_directory_numbering = filedialog.askdirectory(title="Выберите директорию для вывода с нумерацией")
    if output_directory_numbering:
        logging.info(f"Директория для нумерации установлена: {output_directory_numbering}")
def extract_data_from_docx(docx_path):
    """
    Извлекает текст из документа Word (.docx).
    :param docx_path: Путь к документу Word
    :return: Извлечённый текст
    """
    try:
        doc = Document(docx_path)
        data = "\n".join([paragraph.text for paragraph in doc.paragraphs])
        logging.info(f"Текст успешно извлечён из {docx_path}")
        return data
    except Exception as e:
        logging.error(f"Ошибка при извлечении текста из {docx_path}: {str(e)}")
        return ""


def extract_data_from_pdf(pdf_path):
    """
    Извлекает текст из PDF.
    :param pdf_path: Путь к PDF файлу
    :return: Извлечённый текст
    """
    try:
        data = ""
        with fitz.open(pdf_path) as pdf_file:
            for page in pdf_file:
                data += page.get_text()
        logging.info(f"Текст успешно извлечён из {pdf_path}")
        return data
    except Exception as e:
        logging.error(f"Ошибка при извлечении текста из {pdf_path}: {str(e)}")
        return ""


def extract_data_from_xlsx(xlsx_path):
    """
    Извлекает текст из Excel (.xlsx).
    :param xlsx_path: Путь к Excel файлу
    :return: Извлечённый текст
    """
    try:
        workbook = load_workbook(xlsx_path, data_only=True)
        data = ""
        for sheet in workbook.sheetnames:
            sheet_data = "\n".join(
                ["\t".join([str(cell.value) if cell.value else "" for cell in row]) for row in workbook[sheet].rows]
            )
            data += f"\nЛист {sheet}:\n{sheet_data}\n"
        logging.info(f"Текст успешно извлечён из {xlsx_path}")
        return data
    except Exception as e:
        logging.error(f"Ошибка при извлечении текста из {xlsx_path}: {str(e)}")
        return ""

def extract_archive(file_path, extract_to):
    if not file_path or not extract_to:
        messagebox.showerror("Ошибка", "Необходимо указать расположение архива и место для разархивации.")
        return False

    success = False
    try:
        if file_path.endswith('.zip'):
            with zipfile.ZipFile(file_path, 'r') as archive:
                archive.extractall(extract_to)
        elif file_path.endswith('.rar'):
            with rarfile.RarFile(file_path, 'r') as archive:
                archive.extractall(extract_to)
        elif file_path.endswith('.7z'):
            with py7zr.SevenZipFile(file_path, mode='r') as archive:
                archive.extractall(extract_to)
        success = True
        logging.info(f"Файл {file_path} успешно извлечен в {extract_to}")
        messagebox.showinfo("Успех", f"Архив успешно извлечен в {extract_to}")
    except Exception as e:
        logging.error(f"Ошибка извлечения файла {file_path}: {str(e)}")
        messagebox.showerror("Ошибка", "Ошибка при извлечении архива.")
    return success

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with fitz.open(pdf_path) as pdf:
            for page_num in range(len(pdf)):
                page = pdf[page_num]
                text += page.get_text("text")
                if not text.strip():
                    pix = page.get_pixmap()
                    image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    text += pytesseract.image_to_string(image)
    except Exception as e:
        logging.error(f"Ошибка извлечения текста из PDF {pdf_path}: {str(e)}")
    return text

def extract_text_from_docx(docx_path):
    try:
        doc = Document(docx_path)
        return "\n".join([para.text for para in doc.paragraphs])
    except Exception as e:
        logging.error(f"Ошибка извлечения текста из DOCX {docx_path}: {str(e)}")
    return ""

def extract_data_from_excel(xlsx_path):
    try:
        workbook = load_workbook(xlsx_path, data_only=True)
        sheet = workbook.active
        data = [[cell.value for cell in row] for row in sheet.iter_rows()]
        return data
    except Exception as e:
        logging.error(f"Ошибка извлечения данных из Excel {xlsx_path}: {str(e)}")
    return []

def compare_with_reference(data, reference_path):
    try:
        reference = pd.read_excel(reference_path)
        matched_data = [item for item in data if item in reference.values]
        return matched_data
    except Exception as e:
        logging.error(f"Ошибка при загрузке справочника {reference_path}: {str(e)}")
    return []


def create_inventory(matched_data, output_path):
    """
    Создает опись документов в формате .docx и сохраняет её.
    :param matched_data: Данные для внесения в опись
    :param output_path: Путь для сохранения описи
    """
    try:
        doc = Document()
        doc.add_heading('Опись документов', 0)

        for item in matched_data:
            doc.add_paragraph(
                f"Наименование: {item[0]}\nОбозначение: {item[1]}\nКоличество листов: {item[2]}\nФормат: {item[3]}"
            )

        doc.save(output_path)
        logging.info(f"Опись успешно сохранена в {output_path}")
        messagebox.showinfo("Успех", "Опись успешно создана.")
    except Exception as e:
        logging.error(f"Ошибка при создании описи: {str(e)}")
        messagebox.showerror("Ошибка", "Ошибка при создании описи.")

def apply_number_to_file(image_path, number, output_path):
    try:
        with Image.open(image_path) as img:
            draw = ImageDraw.Draw(img)
            draw.text((10, 10), str(number), fill="black")
            img.save(output_path)
            logging.info(f"Файл {image_path} сохранен с номером {number} в {output_path}")
            messagebox.showinfo("Успех", "Файл успешно обновлен с номером.")
    except Exception as e:
        logging.error(f"Ошибка нанесения номера на файл: {str(e)}")
        messagebox.showerror("Ошибка", "Ошибка при нанесении номера.")

def rename_file_with_dialog():
    try:
        current_file_path = filedialog.askopenfilename(title="Выберите файл для переименования")
        if not current_file_path:
            return

        new_name = simpledialog.askstring("Новое имя", "Введите новое имя для файла:")
        if not new_name:
            return

        save_directory = filedialog.askdirectory(title="Выберите директорию для сохранения")
        if not save_directory:
            return

        new_file_path = os.path.join(save_directory, new_name + os.path.splitext(current_file_path)[1])
        os.rename(current_file_path, new_file_path)
        logging.info(f"Файл {current_file_path} переименован и сохранен как {new_file_path}")
        messagebox.showinfo("Успех", f"Файл успешно переименован и перемещен в {new_file_path}")

    except Exception as e:
        logging.error(f"Ошибка при переименовании файла: {str(e)}")
        messagebox.showerror("Ошибка", "Ошибка при переименовании файла.")


class DocumentProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Processor")

        # Поля для путей
        self.archive_path = tk.StringVar()
        self.reference_path = tk.StringVar()
        self.output_directory = tk.StringVar()
        self.numbers_path = tk.StringVar()
        self.files_directory = tk.StringVar()  # Новое поле

        # Элементы интерфейса
        tk.Label(root, text="Путь к архиву:").grid(row=0, column=0, sticky="w")
        tk.Entry(root, textvariable=self.archive_path, width=50).grid(row=0, column=1)
        tk.Button(root, text="Обзор", command=self.select_archive).grid(row=0, column=2)

        tk.Label(root, text="Путь к справочнику:").grid(row=1, column=0, sticky="w")
        tk.Entry(root, textvariable=self.reference_path, width=50).grid(row=1, column=1)
        tk.Button(root, text="Обзор", command=self.select_reference).grid(row=1, column=2)

        tk.Label(root, text="Директория для результатов:").grid(row=2, column=0, sticky="w")
        tk.Entry(root, textvariable=self.output_directory, width=50).grid(row=2, column=1)
        tk.Button(root, text="Обзор", command=self.select_output_directory).grid(row=2, column=2)

        tk.Label(root, text="Путь к файлу с номерами:").grid(row=3, column=0, sticky="w")
        tk.Entry(root, textvariable=self.numbers_path, width=50).grid(row=3, column=1)
        tk.Button(root, text="Обзор", command=self.select_numbers).grid(row=3, column=2)

        tk.Label(root, text="Директория с файлами:").grid(row=4, column=0, sticky="w")  # Новая строка
        tk.Entry(root, textvariable=self.files_directory, width=50).grid(row=4, column=1)
        tk.Button(root, text="Обзор", command=self.select_files_directory).grid(row=4, column=2)

        # Кнопки для функций
        tk.Button(root, text="Извлечь архив", command=self.run_extraction).grid(row=5, column=0, pady=10)
        tk.Button(root, text="Извлечь текст и изображения", command=self.run_extract_text_and_images).grid(row=5,
                                                                                                           column=1,
                                                                                                           pady=10)
        tk.Button(root, text="Сформировать опись", command=self.run_inventory).grid(row=6, column=0, pady=10)
        tk.Button(root, text="Нанести номера", command=self.run_apply_numbers).grid(row=6, column=1, pady=10)
        tk.Button(root, text="Переименовать файлы", command=self.run_rename_files).grid(row=6, column=2, pady=10)

    def select_archive(self):
        file_path = filedialog.askopenfilename(filetypes=[("Archive files", "*.zip *.rar *.7z")])
        if file_path:
            self.archive_path.set(file_path)
            logging.info(f"Архив выбран: {file_path}")

    def select_files_directory(self):
        """Метод для выбора директории с файлами для извлечения текста и изображений."""
        directory = filedialog.askdirectory()
        if directory:
            self.files_directory.set(directory)
            logging.info(f"Директория с файлами выбрана: {directory}")
    def select_reference(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.reference_path.set(file_path)
            logging.info(f"Справочник выбран: {file_path}")

    def select_output_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_directory.set(directory)
            logging.info(f"Директория для результатов выбрана: {directory}")

    def select_numbers(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.numbers_path.set(file_path)
            logging.info(f"Файл с номерами выбран: {file_path}")

    def run_extraction(self):
        extract_archive(self.archive_path.get(), self.output_directory.get())

    def run_extract_text_and_images(self):
        """Метод для извлечения текста и изображений из файлов различных форматов."""
        directory = self.files_directory.get()  # Использование новой директории
        if not directory:
            logging.error("Директория для извлечения не указана.")
            return

        extracted_data = ""
        for file_name in os.listdir(directory):
            file_path = os.path.join(directory, file_name)
            _, ext = os.path.splitext(file_name)

            if ext.lower() == ".xlsx":
                extracted_data += f"\n--- Data from {file_name} ---\n"
                extracted_data += extract_data_from_excel(file_path)

            elif ext.lower() == ".docx":
                extracted_data += f"\n--- Data from {file_name} ---\n"
                extracted_data += extract_data_from_docx(file_path)

            elif ext.lower() == ".txt":
                extracted_data += f"\n--- Data from {file_name} ---\n"
                extracted_data += extract_data_from_txt(file_path)

            elif ext.lower() == ".pdf":
                extracted_data += f"\n--- Data from {file_name} ---\n"
                extracted_data += extract_data_from_pdf(file_path)

        output_path = os.path.join(directory, "extracted_data.txt")
        with open(output_path, "w", encoding="utf-8") as output_file:
            output_file.write(extracted_data)

        logging.info("Извлечение текста и изображений завершено.")
        messagebox.showinfo("Успех", "Извлечение завершено.")

    def run_inventory(self):
        """Метод для создания описи документов."""
        reference_path = self.reference_path.get()
        output_directory = self.output_directory.get()
        extracted_data = extract_data_from_excel(reference_path)
        matched_data = compare_with_reference(extracted_data, reference_path)
        output_path = os.path.join(output_directory, "inventory.docx")
        create_inventory(matched_data, output_path)

    def run_apply_numbers(self):
        """Метод для нанесения номеров на файлы."""
        numbers_path = self.numbers_path.get()
        directory = self.files_directory.get()
        numbers_data = extract_data_from_excel(numbers_path)

        for number_info in numbers_data:
            file_name = number_info[0]  # Assuming file name is in the first column
            number = number_info[1]  # Assuming number is in the second column
            file_path = os.path.join(directory, file_name)
            output_path = os.path.join(directory, f"numbered_{file_name}")
            apply_number_to_file(file_path, number, output_path)

    def run_rename_files(self):
        rename_file_with_dialog()

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentProcessorApp(root)
    root.mainloop()
