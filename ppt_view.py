# ppt_view.py
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QPushButton,
    QTextEdit, QLabel, QComboBox, QListWidget, QListWidgetItem, QCheckBox
)
import torch
import openvino.runtime as ov

class PPTView(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PPT ローカル翻訳")
        self.setGeometry(200, 200, 1000, 600)

        # ------------------------
        # 🎛 ウィジェット定義
        # ------------------------
        self.device_label = QLabel("🔍 デバイス情報を取得中…")
        self.device_label.setStyleSheet("color: #00aa88; font-weight: bold;")

        self.slide_select = QComboBox()
        self.input_text = QTextEdit()
        self.output_text = QTextEdit()
        self.open_btn = QPushButton("📂 PPTを開く")
        self.translate_btn = QPushButton("🚀 翻訳")
        self.replace_btn = QPushButton("🔁 部分置換")
        self.save_btn = QPushButton("💾 保存")

        self.text_list_widget = QListWidget()

        # ------------------------
        # 🧩 レイアウト構成
        # ------------------------
        layout = QVBoxLayout()
        layout.addWidget(self.device_label)
        layout.addWidget(self.open_btn)
        layout.addWidget(QLabel("スライド選択"))
        layout.addWidget(self.slide_select)
        layout.addWidget(QLabel("元スライドテキスト"))
        layout.addWidget(self.input_text)
       
        layout.addWidget(self.translate_btn)
        layout.addWidget(QLabel("翻訳プレビュー"))
        layout.addWidget(self.output_text)
        layout.addWidget(QLabel("置換テキスト選択"))
        layout.addWidget(self.text_list_widget)
        layout.addWidget(self.replace_btn)
        layout.addWidget(self.save_btn)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # ------------------------
        # 🚀 初期化処理
        # ------------------------
        self.update_device_info()

    # ----------------------------
    # 💡 デバイス情報を検出してUIに表示
    # ----------------------------
    def update_device_info(self):
        info_text = self.detect_devices()
        self.device_label.setText(info_text)

    def detect_devices(self):
        try:
            if torch.cuda.is_available():
                return f"💻 使用デバイス: GPU ({torch.cuda.get_device_name(0)})"
            elif torch.backends.mps.is_available():
                return "💻 使用デバイス: GPU (Apple MPS)"

            core = ov.Core()
            available = [d.upper() for d in core.available_devices]
            if any("NPU" in d for d in available):
                return "💻 使用デバイス: NPU (OpenVINO)"
            elif any("GPU" in d for d in available):
                return "💻 使用デバイス: GPU (OpenVINO)"
            else:
                return "💻 使用デバイス: CPU"

        except Exception:
            return "💻 使用デバイス: CPU"

    # ----------------------------
    # 💡 テキストリスト更新（チェックボックス付き）
    # ----------------------------
    def update_text_list(self, texts):
        self.text_list_widget.clear()
        for t in texts:
            item = QListWidgetItem()
            checkbox = QCheckBox(t)
            self.text_list_widget.addItem(item)
            self.text_list_widget.setItemWidget(item, checkbox)
