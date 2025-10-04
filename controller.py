# controller.py
from ppt_com_model import PowerPointCOM
from translator_model import TranslatorModel
from ppt_view import PPTView
from PySide6.QtWidgets import QFileDialog

class PPTController:
    def __init__(self, view: PPTView):
        self.view = view
        self.ppt = None
        self.translator = TranslatorModel()
        self.slides_text = []
        self.preview_translations = []

        view.open_btn.clicked.connect(self.load_ppt)
        view.translate_btn.clicked.connect(self.translate_slide)
        view.replace_btn.clicked.connect(self.replace_slide)
        view.save_btn.clicked.connect(self.save_ppt)

    def load_ppt(self):
        path, _ = QFileDialog.getOpenFileName(self.view, "PPTを選択", "", "PowerPoint (*.pptx)")
        if not path:
            return
        self.ppt = PowerPointCOM(path)
        self.slides_text = self.ppt.extract_texts()
        self.view.slide_select.clear()
        self.view.slide_select.addItems([f"Slide {i+1}" for i in range(len(self.slides_text))])
        self.view.input_text.setText("\n".join(self.slides_text[0]))
        self.preview_translations = []

    def translate_slide(self):
        if not self.ppt:
            return
        idx = self.view.slide_select.currentIndex()
        originals = self.slides_text[idx]
        self.preview_translations = [self.translator.translate_text(t) for t in originals]
        self.view.output_text.setText("\n".join(self.preview_translations))

    def replace_slide(self):
        if not self.ppt or not self.preview_translations:
            return
        idx = self.view.slide_select.currentIndex()
        self.ppt.replace_text_preserve_format(idx, self.slides_text[idx], self.preview_translations)
        self.view.input_text.setText("\n".join(self.preview_translations))

    def save_ppt(self):
        if not self.ppt:
            return
        path = self.ppt.save_as()
        self.view.output_text.append(f"\n✅ 保存完了: {path}")
        self.ppt.close()
