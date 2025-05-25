# pdf.py
import sys
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QLabel, QProgressBar, QMessageBox, QCheckBox, QFileDialog
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
import qdarkstyle
import re
import pytesseract
from PIL import Image, ImageEnhance
import logging

logging.basicConfig(filename='ocr_log.txt', level=logging.DEBUG, format='%(asctime)s - %(message)s')

class PDFWordSearch(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Word Search")
        self.setGeometry(100, 100, 800, 600)
        self.words = []
        self.found_words = []
        self.pdf_path = None
        self.page_texts = []
        self.init_ui()
        self.apply_dark_theme()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.select_pdf_button = QPushButton("انتخاب فایل PDF")
        self.select_pdf_button.setFont(self.get_font("Vazir", 12))
        self.select_pdf_button.clicked.connect(self.select_pdf_file)
        layout.addWidget(self.select_pdf_button)

        self.pdf_path_label = QLabel("هیچ فایلی انتخاب نشده است")
        self.pdf_path_label.setFont(self.get_font("Vazir", 12))
        layout.addWidget(self.pdf_path_label)

        self.label_entered = QLabel("تعداد کلمات/اعداد وارد شده: 0")
        self.label_entered.setFont(self.get_font("Vazir", 12))
        layout.addWidget(self.label_entered)

        self.label_found = QLabel("تعداد کلمات/اعداد پیدا شده: 0")
        self.label_found.setFont(self.get_font("Vazir", 12))
        layout.addWidget(self.label_found)

        self.table = QTableWidget()
        self.table.setRowCount(0)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["کلمه/عدد", "وضعیت"])
        self.table.setFont(self.get_font("Vazir", 10))
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(1, 50)
        self.table.cellChanged.connect(self.update_entered_count)
        layout.addWidget(self.table)

        self.search_button = QPushButton("جستجو")
        self.search_button.setFont(self.get_font("Vazir", 12))
        self.search_button.clicked.connect(self.search_words)
        layout.addWidget(self.search_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.add_row()

    def get_font(self, font_name, size):
        font = QFont(font_name, size)
        if not font.exactMatch():
            font = QFont("Arial", size)
        return font

    def apply_dark_theme(self):
        self.setStyleSheet(qdarkstyle.load_stylesheet(qt_api='pyqt6'))
        self.setStyleSheet(self.styleSheet() + """
            QTableWidget::item { color: white; }
            QLabel { color: white; }
            QPushButton { background-color: #555; color: white; }
        """)

    def select_pdf_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل PDF", "", "PDF Files (*.pdf)")
        if file_path:
            self.pdf_path = file_path
            self.pdf_path_label.setText(f"فایل انتخاب شده: {file_path}")

    def add_row(self):
        if self.table.rowCount() >= 100:
            return
        row_count = self.table.rowCount()
        self.table.insertRow(row_count)
        word_item = QTableWidgetItem()
        word_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.table.setItem(row_count, 0, word_item)
        checkbox = QCheckBox()
        checkbox.setEnabled(False)
        checkbox_widget = QWidget()
        checkbox_layout = QVBoxLayout(checkbox_widget)
        checkbox_layout.addWidget(checkbox)
        checkbox_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        checkbox_layout.setContentsMargins(0, 0, 0, 0)
        self.table.setCellWidget(row_count, 1, checkbox_widget)

    def update_entered_count(self):
        self.words = []
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item and item.text().strip():
                self.words.append(item.text().strip())
                if row == self.table.rowCount() - 1 and self.table.rowCount() < 100:
                    self.add_row()
        self.label_entered.setText(f"تعداد کلمات/اعداد وارد شده: {len(self.words)}")
        logging.debug(f"Entered words/numbers: {self.words}")

    def normalize_numbers(self, text):
        persian_to_latin = str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')
        return text.translate(persian_to_latin)

    def normalize_text(self, text):
        # حذف کاراکترهای نامرئی و نویز
        text = text.replace('\u200c', ' ').replace('\u200b', '').replace('\ufeff', '')
        # فقط حروف فارسی، اعداد، اسلش، خط تیره و فاصله نگه داشته شوند
        text = re.sub(r'[^\u0600-\u06FF0-9۰-۹/\-\s]', ' ', text)
        text = re.sub(r'\s+', ' ', text.strip())
        return text

    def preprocess_image(self, img):
        # بهبود کیفیت تصویر برای OCR
        img = img.convert('L')  # تبدیل به خاکستری
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.5)  # افزایش کنتراست
        enhancer = ImageEnhance.Sharpness(img)
        img = enhancer.enhance(2.0)  # افزایش وضوح
        return img

    def extract_text_from_page(self, page):
        text = page.get_text("text")
        if not text.strip():
            try:
                pix = page.get_pixmap(dpi=300)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.rgb)
                img = self.preprocess_image(img)
                text = pytesseract.image_to_string(img, lang='fas+eng', config='--psm 6 --oem 1')
                logging.debug(f"OCR text page {page.number + 1}: {text[:500]}...")
            except Exception as e:
                logging.error(f"OCR error on page {page.number + 1}: {str(e)}")
                text = ""
        normalized_text = self.normalize_text(text)
        logging.debug(f"Normalized text page {page.number + 1}: {normalized_text[:500]}...")
        return text, normalized_text

    def is_date_format(self, word):
        """بررسی اینکه آیا ورودی فرمت تاریخ (مانند ۱۴۰۳/۰۷/۰۹) دارد"""
        date_pattern = r'^[\u06F0-\u06F90-9]{4}[\/\-][\u06F0-\u06F90-9]{1,2}[\/\-][\u06F0-\u06F90-9]{1,2}$'
        return bool(re.match(date_pattern, word))

    def search_date_in_text(self, word, raw_text, normalized_text, result, index):
        """جستجوی تاریخ با فرمت‌های مختلف"""
        raw_word = word.strip()
        normalized_word = self.normalize_numbers(raw_word)

        date_patterns = [
            re.escape(raw_word),  # دقیقاً همان ورودی
            re.escape(normalized_word),  # با اعداد لاتین
            raw_word.replace('/', '-'),  # اسلش به خط تیره
            normalized_word.replace('/', '-'),
            raw_word.replace('/', ' '),  # اسلش به فاصله
            normalized_word.replace('/', ' '),
            # الگوی عمومی برای تاریخ
            r'\b[\u06F0-\u06F90-9]{4}[\s\/\-][\u06F0-\u06F90-9]{1,2}[\s\/\-][\u06F0-\u06F90-9]{1,2}\b',
        ]

        for pattern in date_patterns:
            if (re.search(pattern, raw_text, re.UNICODE) or
                re.search(pattern, normalized_text, re.UNICODE)):
                result[index] = True
                logging.debug(f"Found date pattern '{pattern}' for word '{word}'")
                return

    def search_word_in_text(self, word, raw_text, normalized_text, result, index):
        """جستجوی کلمات عمومی یا تاریخ"""
        if self.is_date_format(word):
            self.search_date_in_text(word, raw_text, normalized_text, result, index)
            return

        # جستجوی کلمات عمومی با تطبیق دقیق
        raw_word = word.strip()
        normalized_word = self.normalize_numbers(raw_word)

        # فقط کلمه دقیق یا با تغییرات کوچک
        patterns = [
            re.escape(raw_word),  # کلمه دقیق
            re.escape(normalized_word),  # با اعداد لاتین
        ]

        for pattern in patterns:
            # استفاده از \b برای اطمینان از تطبیق کلمه کامل
            if (re.search(r'\b' + pattern + r'\b', raw_text, re.UNICODE) or
                re.search(r'\b' + pattern + r'\b', normalized_text, re.UNICODE)):
                result[index] = True
                logging.debug(f"Found word pattern '{pattern}' for word '{word}'")
                return

    def search_words(self):
        if not self.words:
            QMessageBox.warning(self, "خطا", "لطفاً حداقل یک کلمه یا عدد وارد کنید!")
            return

        if not self.pdf_path:
            QMessageBox.warning(self, "خطا", "لطفاً یک فایل PDF انتخاب کنید!")
            return

        try:
            doc = fitz.open(self.pdf_path)
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در باز کردن PDF: {str(e)}")
            return

        self.page_texts = []
        with open("debug_text.txt", "w", encoding="utf-8") as f:
            for i, page in enumerate(doc):
                raw_text, normalized_text = self.extract_text_from_page(page)
                self.page_texts.append((raw_text, normalized_text))
                f.write(f"\n--- Page {i+1} ---\n{raw_text}\n--- Normalized ---\n{normalized_text}\n")

        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(self.words))
        self.found_words = []
        results = [False] * len(self.words)

        for idx, word in enumerate(self.words):
            for raw_text, normalized_text in self.page_texts:
                self.search_word_in_text(word, raw_text, normalized_text, results, idx)
                if results[idx]:  # اگر کلمه پیدا شد، نیازی به ادامه جستجو نیست
                    break

        for idx, word in enumerate(self.words):
            if results[idx]:
                self.found_words.append(word)
            checkbox = self.table.cellWidget(idx, 1).layout().itemAt(0).widget()
            checkbox.setChecked(results[idx])
            checkbox.setStyleSheet(f"QCheckBox::indicator {{ background-color: {'green' if results[idx] else 'red'}; }}")
            self.progress_bar.setValue(idx + 1)
            QApplication.processEvents()

        doc.close()
        self.progress_bar.setVisible(False)
        self.label_found.setText(f"تعداد کلمات/اعداد پیدا شده: {len(self.found_words)}")
        QMessageBox.information(self, "اتمام", "جستجو با موفقیت به پایان رسید!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFWordSearch()
    window.show()
    sys.exit(app.exec())
