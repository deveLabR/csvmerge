import os
import sys
import glob
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QHBoxLayout, QVBoxLayout, QFileDialog, QMessageBox,
    QSpacerItem, QSizePolicy, QMenuBar, QMenu, QTextEdit
)
from PySide6.QtGui import QAction, QIcon


# https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class CsvMerger(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("CSV Merger 1.0")
        self.setGeometry(100, 100, 500, 400)
        self.setWindowIcon(QIcon(resource_path("merge_icon.ico")))

        # Main Layout
        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(10, 10, 10, 10)

        # Menu Bar
        self.menu_bar = QMenuBar(self)
        about_menu = QMenu("Довідка", self)
        about_action = QAction("Про програму", self)
        about_action.triggered.connect(self.show_about_dialog)
        about_menu.addAction(about_action)
        self.menu_bar.addMenu(about_menu)
        self.main_layout.setMenuBar(self.menu_bar)

        # Set fixed width for labels
        label_width = 100

        # Input Folder Widgets
        self.input_folder_label = QLabel("Вхідна папка:")
        self.input_folder_label.setFixedWidth(label_width)
        self.input_folder_entry = QLineEdit()
        self.input_button = QPushButton("Обрати")
        self.input_button.clicked.connect(self.select_input_folder)

        # Input Folder Layout
        self.input_layout = QHBoxLayout()
        self.input_layout.setSpacing(5)
        self.input_layout.addWidget(self.input_folder_label)
        self.input_layout.addWidget(self.input_folder_entry)
        self.input_layout.addWidget(self.input_button)

        # Output Folder Widgets
        self.output_folder_label = QLabel("Вихідна папка:")
        self.output_folder_label.setFixedWidth(label_width)
        self.output_folder_entry = QLineEdit()
        self.output_button = QPushButton("Обрати")
        self.output_button.clicked.connect(self.select_output_folder)

        # Output Folder Layout
        self.output_layout = QHBoxLayout()
        self.output_layout.setSpacing(5)
        self.output_layout.addWidget(self.output_folder_label)
        self.output_layout.addWidget(self.output_folder_entry)
        self.output_layout.addWidget(self.output_button)

        # Status Area
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setMinimumHeight(150)
        self.status_text.setPlaceholderText("Тут відображатиметься статус операцій...")

        # Merge Button
        self.merge_button = QPushButton("Об'єднати та зберегти")
        self.merge_button.clicked.connect(self.merge_and_save)

        # Center the button
        self.button_layout = QHBoxLayout()
        self.button_layout.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.button_layout.addWidget(self.merge_button)
        self.button_layout.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

        # Add widgets to main layout
        self.main_layout.addLayout(self.input_layout)
        self.main_layout.addLayout(self.output_layout)
        self.main_layout.addWidget(self.status_text)
        self.main_layout.addLayout(self.button_layout)

        # Set main layout
        self.setLayout(self.main_layout)

    def log_status(self, message):
        """Add message to Status Area"""
        self.status_text.append(message)
        self.status_text.verticalScrollBar().setValue(
            self.status_text.verticalScrollBar().maximum()
        )
        # Update interface
        QApplication.processEvents()

    @staticmethod
    def check_folder_empty(folder_path):
        """Перевіряє чи папка існує і чи містить CSV файли"""
        if not os.path.exists(folder_path):
            return False, "Папка не існує"

        if not os.path.isdir(folder_path):
            return False, "Вказаний шлях не є папкою"

        csv_files = glob.glob(os.path.join(folder_path, '*.csv'))
        if not csv_files:
            return False, "У папці відсутні CSV файли"

        return True, f"Знайдено {len(csv_files)} CSV файлів"

    def select_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Виберіть папку з CSV файлами")
        if folder:
            is_valid, message = self.check_folder_empty(folder)
            if is_valid:
                self.input_folder_entry.setText(folder)
                self.log_status(f"Обрано вхідну папку: {folder}")
                self.log_status(message)
            else:
                self.log_status(f"❌ Помилка вибору вхідної папки: {message}")
                QMessageBox.warning(self, "Попередження",
                                    f"Обрана папка не підходить: {message}\nБудь ласка, оберіть іншу папку.")
                self.input_folder_entry.clear()

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Виберіть папку для збереження результату")
        if folder:
            if os.path.exists(folder) and os.path.isdir(folder):
                self.output_folder_entry.setText(folder)
                self.log_status(f"Обрано вихідну папку: {folder}")

                # Перевіряємо наявність прав на запис
                test_file_path = os.path.join(folder, 'test_write_permission.tmp')
                try:
                    with open(test_file_path, 'w') as f:
                        f.write('')
                    os.remove(test_file_path)
                except Exception:
                    self.log_status("❌ Попередження: відсутні права на запис у вихідну папку")
                    QMessageBox.warning(self, "Попередження",
                                        "Відсутні права на запис у вибрану папку. Будь ласка, оберіть іншу папку.")
                    self.output_folder_entry.clear()
            else:
                self.log_status("❌ Помилка: обраний шлях не є папкою")
                QMessageBox.warning(self, "Помилка",
                                    "Обраний шлях не є папкою. Будь ласка, оберіть коректну папку.")
                self.output_folder_entry.clear()

    def merge_and_save(self):
        input_folder = self.input_folder_entry.text()
        output_folder = self.output_folder_entry.text()

        if not input_folder or not output_folder:
            self.log_status("❌ Помилка: Будь ласка, виберіть обидві папки.")
            QMessageBox.warning(self, "Помилка", "Будь ласка, виберіть обидві папки.")
            return

        # Повторна перевірка вхідної папки перед об'єднанням
        is_valid, message = self.check_folder_empty(input_folder)
        if not is_valid:
            self.log_status(f"❌ Помилка: {message}")
            QMessageBox.critical(self, "Помилка", message)
            return

        try:
            self.log_status("🔄 Починаємо об'єднання файлів...")
            self.merge_csv(input_folder, output_folder)
            self.log_status("✅ Файли успішно об'єднані та збережені!")
            QMessageBox.information(self, "Успіх", "Файли успішно об'єднані та збережені!")
        except Exception as e:
            error_message = f"❌ Помилка: {str(e)}"
            self.log_status(error_message)
            QMessageBox.critical(self, "Помилка", error_message)

    def merge_csv(self, input_folder, output_folder):
        csv_files = glob.glob(os.path.join(input_folder, '*.csv'))

        if not csv_files:
            raise FileNotFoundError("У вхідній папці немає CSV файлів.")

        self.log_status(f"Знайдено {len(csv_files)} CSV файлів")

        dfs = []
        total_rows = 0
        for i, csv_file in enumerate(csv_files, 1):
            try:
                self.log_status(f"Обробка файлу {i}/{len(csv_files)}: {os.path.basename(csv_file)}")
                df = pd.read_csv(csv_file, encoding='cp1251', sep=';', quotechar='"', engine='python')
                total_rows += len(df)
                self.log_status(f"Прочитано {len(df)} рядків з файлу {os.path.basename(csv_file)}")
                dfs.append(df)
            except Exception as e:
                self.log_status(f"❌ Помилка при читанні файлу {os.path.basename(csv_file)}: {str(e)}")
                raise

        self.log_status("Об'єднання файлів...")
        combined_df = pd.concat(dfs, ignore_index=True)

        output_file_path = os.path.join(output_folder, 'merged_output.csv')
        combined_df.to_csv(output_file_path, index=False, encoding='utf-8-sig', sep=';', quotechar='"')
        self.log_status(f"Збережено результат у файл: {output_file_path}")
        self.log_status(f"Загальна кількість об'єднаних рядків: {total_rows}")

    def show_about_dialog(self):
        about_text = (
            "<b>CSV Merger 1.0</b><br><br>"
            "Автор: deveLabR<br>"
            "Ліцензія: MIT<br>"
            "Репозиторій: <a href='https://github.com/deveLabR/csvmerge'>GitHub</a><br>"
        )
        QMessageBox.about(self, "Про програму", about_text)


if __name__ == "__main__":
    app = QApplication([])
    app.setWindowIcon(QIcon(resource_path("merge_icon.ico")))
    window = CsvMerger()
    window.show()
    app.exec()
