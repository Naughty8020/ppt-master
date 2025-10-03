import sys, os, subprocess, time
from PySide6.QtWidgets import QFileDialog
from model import PPTModel, TranslatorModel
from view import PPTView

if sys.platform == "win32":
    import win32com.client

# ---------------------------
# OS別 PPT操作
# ---------------------------
def open_ppt(file_path: str):
    file_path = os.path.abspath(file_path)
    if sys.platform == "darwin":
        subprocess.run(["open", file_path])
    elif sys.platform == "win32":
        os.startfile(file_path)
    else:
        subprocess.run(["xdg-open", file_path])

def close_ppt():
    if sys.platform == "win32":
        try:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            ppt_app.Quit()
            time.sleep(1)
        except Exception as e:
            print(f"PowerPoint終了失敗: {e}")
    elif sys.platform == "darwin":
        subprocess.run([
            "osascript", "-e",
            'tell application "Microsoft PowerPoint" to quit'
        ])
        time.sleep(1)

# ---------------------------
# PPTコントローラー
# ---------------------------
class PPTController:
    def __init__(self, view: PPTView):
        self.view = view
        self.model = None
        self.edited_ppt_path = None
        self.translator = TranslatorModel(
            src_lang="ja_XX",  # 入力：日本語
            tgt_lang="en_XX"   # 出力：英語
        )

        # イベント接続
        view.open_btn.clicked.connect(self.load_ppt)
        view.slide_select.currentIndexChanged.connect(self.display_slide_text)
        view.translate_btn.clicked.connect(self.translate_slide)
        view.toneup_btn.clicked.connect(self.toneup_slide)
        view.save_btn.clicked.connect(self.save_ppt)
        view.open_in_app_btn.clicked.connect(self.open_in_app)

    def load_ppt(self):
        path, _ = QFileDialog.getOpenFileName(self.view, "PPTを選択", "", "PowerPoint Files (*.pptx)")
        if not path:
            return
        self.model = PPTModel(path)
        slides_text = self.model.extract_slides_text()
        self.view.slide_select.clear()
        self.view.slide_select.addItems([f"Slide {i+1}" for i in range(len(slides_text))])
        self.view.model_path_label.setText(f"PPTパス: {path}")
        self.view.input_text.setText("\n".join(slides_text[0]))
        self.view.output_text.clear()

    def display_slide_text(self):
        if not self.model: return
        idx = self.view.slide_select.currentIndex()
        slides_text = self.model.extract_slides_text()
        if 0 <= idx < len(slides_text):
            self.view.input_text.setText("\n".join(slides_text[idx]))
            self.view.output_text.clear()

    def translate_slide(self):
        if not self.model: return
        idx = self.view.slide_select.currentIndex()
        slides_text = self.model.extract_slides_text()
        original_texts = slides_text[idx]

        translated_texts = []
        for text in original_texts:
            # 改行を保持したまま1行ごとに翻訳
            lines = text.split("\n")
            translated_lines = [self.translator.translate_text(line) if line.strip() else "" for line in lines]
            translated_texts.append("\n".join(translated_lines))

        self.view.output_text.setText("\n".join(translated_texts))
        self.model.update_slide_text(idx, translated_texts)
        self.edited_ppt_path = self.model.save()

    def toneup_slide(self):
        if not self.model: return
        idx = self.view.slide_select.currentIndex()
        slides_text = self.model.extract_slides_text()
        original_texts = slides_text[idx]
        toned_texts = [self.translator.tone_up(t) for t in original_texts]
        self.view.output_text.setText("\n".join(toned_texts))
        self.model.update_slide_text(idx, toned_texts)
        self.edited_ppt_path = self.model.save()

    def save_ppt(self):
        if not self.model:
            return
        close_ppt()
        idx = self.view.slide_select.currentIndex()
        if 0 <= idx < len(self.model.extract_slides_text()):
            edited_texts = self.view.output_text.toPlainText().split("\n")
            self.model.update_slide_text(idx, edited_texts)
            path = self.model.save()
            self.edited_ppt_path = path
            self.view.output_text.append(f"PPT保存完了: {path}")

    def open_in_app(self):
        if not self.model: return
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
