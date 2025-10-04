# ppt_view.py
from PySide6.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QPushButton, QTextEdit, QLabel, QComboBox

class PPTView(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PPT ローカル翻訳")
        self.setGeometry(200, 200, 1000, 600)

        self.slide_select = QComboBox()
        self.input_text = QTextEdit()
        self.output_text = QTextEdit()
        self.open_btn = QPushButton("📂 PPTを開く")
        self.translate_btn = QPushButton("🚀 翻訳")
        self.replace_btn = QPushButton("🔁 置換")
        self.save_btn = QPushButton("💾 保存")

        layout = QVBoxLayout()
        layout.addWidget(self.open_btn)
        layout.addWidget(QLabel("スライド選択"))
        layout.addWidget(self.slide_select)
        layout.addWidget(QLabel("元スライドテキスト"))
        layout.addWidget(self.input_text)
        layout.addWidget(self.translate_btn)
        layout.addWidget(QLabel("翻訳プレビュー"))
        layout.addWidget(self.output_text)
        layout.addWidget(self.replace_btn)
        layout.addWidget(self.save_btn)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
