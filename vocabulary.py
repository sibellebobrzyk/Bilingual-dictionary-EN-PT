import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QListWidget, QMessageBox, QFormLayout, QTextEdit
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

excel_file = "dictionary.xlsx"

def load_spreadsheet():
    try:
        return pd.read_excel(excel_file)
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Portuguese", "English", "Meaning"])
        df.to_excel(excel_file, index=False)
        return df

class DictionaryApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Bilingual Dictionary")
        self.setGeometry(100, 100, 800, 600)
        self.setStyleSheet("""
            QWidget {
                background-color: #F5E6F7;
                border-radius: 15px;
            }
            QLabel {
                color: #4A235A;
            }
            QLineEdit, QTextEdit {
                border: 1px solid #DCC6E0;
                border-radius: 10px;
                padding: 8px;
                background-color: #FFFFFF;
            }
            QPushButton {
                background-color: #9B59B6;
                color: white;
                border-radius: 10px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #76448A;
            }
            QListWidget {
                border: 1px solid #DCC6E0;
                border-radius: 10px;
                background-color: #FFFFFF;
            }
        """)

        title_font = QFont("Arial", 16, QFont.Bold)
        normal_font = QFont("Arial", 11)

        main_layout = QVBoxLayout()

        title = QLabel("Vocabulary EN-PT")
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)

        form_layout = QFormLayout()

        self.input_portuguese = QLineEdit()
        self.input_portuguese.setFont(normal_font)
        self.input_portuguese.setPlaceholderText("Enter the word in Portuguese")
        form_layout.addRow(QLabel("Portuguese:"), self.input_portuguese)

        self.input_english = QLineEdit()
        self.input_english.setFont(normal_font)
        self.input_english.setPlaceholderText("Enter the word in English")
        form_layout.addRow(QLabel("English:"), self.input_english)

        self.input_meaning = QTextEdit()
        self.input_meaning.setFont(normal_font)
        self.input_meaning.setPlaceholderText("Enter the meaning of the word")
        form_layout.addRow(QLabel("Meaning:"), self.input_meaning)

        main_layout.addLayout(form_layout)

        self.save_button = QPushButton("Save")
        self.save_button.setFont(normal_font)
        self.save_button.clicked.connect(self.save_word)
        main_layout.addWidget(self.save_button)

        search_layout = QHBoxLayout()

        self.search_input = QLineEdit()
        self.search_input.setFont(normal_font)
        self.search_input.setPlaceholderText("Search for a word or expression")
        search_layout.addWidget(self.search_input)

        self.search_button = QPushButton("Search")
        self.search_button.setFont(normal_font)
        self.search_button.clicked.connect(self.search_word)
        search_layout.addWidget(self.search_button)

        main_layout.addLayout(search_layout)

        self.results_list = QListWidget()
        self.results_list.setFont(normal_font)
        main_layout.addWidget(self.results_list)

        self.setLayout(main_layout)
        self.update_list()

    def save_word(self):
        portuguese = self.input_portuguese.text().strip()
        english = self.input_english.text().strip()
        meaning = self.input_meaning.toPlainText().strip()

        if portuguese and english and meaning:
            new_entry = pd.DataFrame({"Portuguese": [portuguese], "English": [english], "Meaning": [meaning]})
            df = load_spreadsheet()
            df = pd.concat([df, new_entry], ignore_index=True)
            df.to_excel(excel_file, index=False)
            self.input_portuguese.clear()
            self.input_english.clear()
            self.input_meaning.clear()
            self.update_list()
            QMessageBox.information(self, "Success", "Word saved successfully!")
        else:
            QMessageBox.warning(self, "Warning", "Please fill in all fields!")

    def search_word(self):
        term = self.search_input.text().strip().lower()
        if term:
            df = load_spreadsheet()
            results = df[
                df["Portuguese"].str.lower().str.contains(term) |
                df["English"].str.lower().str.contains(term) |
                df["Meaning"].str.lower().str.contains(term)
            ]
            self.results_list.clear()
            for _, row in results.iterrows():
                self.results_list.addItem(
                    f"{row['Portuguese']} - {row['English']} ({row['Meaning']})"
                )
        else:
            QMessageBox.warning(self, "Warning", "Please enter a term to search!")

    def update_list(self):
        df = load_spreadsheet()
        self.results_list.clear()
        for _, row in df.iterrows():
            self.results_list.addItem(
                f"{row['Portuguese']} - {row['English']} ({row['Meaning']})"
            )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DictionaryApp()
    window.show()
    sys.exit(app.exec_())
