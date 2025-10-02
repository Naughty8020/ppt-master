import sys
import os
import subprocess
import time

from PySide6.QtWidgets import QFileDialog
from model import PPTModel, TranslatorModel
from view import PPTView

# Windowsのみ
if sys.platform == "win32":
    import win32com.client


# ---------------------------
# OS別 PowerPoint操作関数
# ---------------------------
def open_ppt(file_path: str):
    """PowerPointファイルをOSに応じて開く"""
    file_path = os.path.abspath(file_path)
    if sys.platform == "darwin":  # macOS
        subprocess.run(["open", file_path])
    elif sys.platform == "win32":  # Windows
        os.startfile(file_path)
    else:  # Linuxなど
        subprocess.run(["xdg-open", file_path])


def close_ppt():
    """開いているPowerPointをOSに応じて閉じる"""
    if sys.platform == "win32":
        try:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            ppt_app.Quit()
            time.sleep(1)
            print("PowerPointを閉じました（Windows）。")
        except Exception as e:
            print(f"PowerPointの終了に失敗しました: {e}")
    elif sys.platform == "darwin":
        subprocess.run([
            "osascript", "-e",
            'tell application "Microsoft PowerPoint" to quit'
        ])
        time.sleep(1)
        print("PowerPointを閉じました（macOS）。")
    else:
        print("PowerPointの閉じる操作はこのOSでは未対応です。")


# ---------------------------
# PPTコントローラー
# ---------------------------
class PPTController:
    def __init__(self, view: PPTView):
        self.view = view
        self.model = None
        self.edited_ppt_path = None
        self.translator = TranslatorModel()

        # イベント接続
        view.open_btn.clicked.connect(self.load_ppt)
        view.slide_select.currentIndexChanged.connect(self.display_slide_text)
        view.translate_btn.clicked.connect(self.translate_slide)
        view.save_btn.clicked.connect(self.save_ppt)
        view.open_in_app_btn.clicked.connect(self.open_in_app)

    # ---------------------------
    # PPT読み込み
    # ---------------------------
    def load_ppt(self):
        path, _ = QFileDialog.getOpenFileName(
            self.view, "PPTを選択", "", "PowerPoint Files (*.pptx)"
        )
        if not path:
            return

        self.model = PPTModel(path)
        slides_text = self.model.extract_slides_text()

        self.view.slide_select.clear()
        self.view.slide_select.addItems([f"Slide {i+1}" for i in range(len(slides_text))])
        self.view.model_path_label.setText(f"PPTパス: {path}")
        self.view.input_text.setText(slides_text[0])
        self.view.output_text.clear()

    # ---------------------------
    # スライド表示
    # ---------------------------
    def display_slide_text(self):
        if not self.model:
            return
        idx = self.view.slide_select.currentIndex()
        slides_text = self.model.extract_slides_text()
        if 0 <= idx < len(slides_text):
            self.view.input_text.setText(slides_text[idx])
            self.view.output_text.clear()

    # ---------------------------
    # 翻訳
    # ---------------------------
    def translate_slide(self):
        if not self.model:
            return
        idx = self.view.slide_select.currentIndex()
        text = self.view.input_text.toPlainText()
        translated = self.translator.translate_text(text)
        self.view.output_text.setText(translated)
        self.model.update_slide_text(idx, translated)
        self.edited_ppt_path = self.model.save()  # 翻訳後PPT保存

    # ---------------------------
    # 保存
    # ---------------------------
    def save_ppt(self):
        if not self.model:
            return
        close_ppt()  # OS判定済み
        path = self.model.save()
        self.edited_ppt_path = path
        self.view.output_text.append(f"PPT保存完了: {path}")

    # ---------------------------
    # PPTアプリで開く
    # ---------------------------
    def open_in_app(self):
        """元 or 翻訳後PPTを開く"""
        if not self.model:
            return

        if self.view.view_box.currentText() == "元のPPT":
            file_to_open = self.model.ppt_path
        elif self.view.view_box.currentText() == "翻訳後PPT":
            if not self.edited_ppt_path:
                self.view.output_text.append("まず翻訳・保存してください")
                return
            file_to_open = self.edited_ppt_path
        else:
            return

        open_ppt(file_to_open)
