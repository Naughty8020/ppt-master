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
        self.slides_text = []  # 元テキスト保持
        self.preview_translations = []  # 翻訳文保持

        # ボタン・スライド選択連携
        view.open_btn.clicked.connect(self.load_ppt)
        view.translate_btn.clicked.connect(self.translate_slide)
        view.replace_btn.clicked.connect(self.replace_slide_partial)
        view.save_btn.clicked.connect(self.save_ppt)
        view.slide_select.currentIndexChanged.connect(self.on_slide_change)

    def load_ppt(self):
        path, _ = QFileDialog.getOpenFileName(self.view, "PPTを選択", "", "PowerPoint (*.pptx)")
        if not path:
            return
        self.ppt = PowerPointCOM(path)
        self.slides_text = self.ppt.extract_texts()
        self.view.slide_select.clear()
        self.view.slide_select.addItems([f"Slide {i+1}" for i in range(len(self.slides_text))])
        self.on_slide_change(0)
        self.preview_translations = []

    def on_slide_change(self, idx):
        """スライド切替時は元テキスト表示、チェックリストは空に"""
        if not self.ppt:
            return
        texts = self.slides_text[idx]
        self.view.input_text.setText("\n".join(texts))
        self.view.text_list_widget.clear()  # 翻訳前はチェックリスト表示しない

    def translate_slide(self):
        """スライド内テキストを翻訳してチェックリストに反映"""
        if not self.ppt:
            return
        idx = self.view.slide_select.currentIndex()
        originals = self.slides_text[idx]

        # 翻訳
        self.preview_translations = [self.translator.translate_text(t) for t in originals]
        self.view.output_text.setText("\n".join(self.preview_translations))

        # チェックボックスリストを翻訳文で更新
        self.view.update_text_list(self.preview_translations)

    def replace_slide_partial(self):
        """チェックされた箇所だけ置換"""
        if not self.ppt or not self.preview_translations:
            return

        idx = self.view.slide_select.currentIndex()
        originals = self.slides_text[idx]
        translations = self.preview_translations.copy()

        # チェックされているインデックスだけ置換
        selected_indices = []
        for i in range(self.view.text_list_widget.count()):
            item = self.view.text_list_widget.item(i)
            checkbox = self.view.text_list_widget.itemWidget(item)
            if checkbox.isChecked():
                selected_indices.append(i)

        # チェックされていない部分は元テキストに戻す
        for i in range(len(translations)):
            if i not in selected_indices:
                translations[i] = originals[i]

        # スライド置換
        self.ppt.replace_text_preserve_format(idx, originals, translations)
        # 入力欄更新
        self.view.input_text.setText("\n".join(translations))

    def save_ppt(self):
        if not self.ppt:
            return
        path = self.ppt.save_as()
        self.view.output_text.append(f"\n✅ 保存完了: {path}")
        self.ppt.close()
