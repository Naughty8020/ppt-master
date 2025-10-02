from pptx import Presentation
from transformers import AutoModelForSeq2SeqLM, AutoTokenizer
import torch

# ---------------------------
# PPTモデル
# ---------------------------
class PPTModel:
    def __init__(self, ppt_path):
        self.ppt_path = ppt_path
        self.presentation = Presentation(ppt_path)
        self.edited_ppt_path = None

    def extract_slides_text(self):
        slides_text = []
        for slide in self.presentation.slides:
            texts = [
                shape.text
                for shape in slide.shapes
                if hasattr(shape, "text") and shape.has_text_frame
            ]
            slides_text.append("\n".join(texts))
        return slides_text

    def update_slide_text(self, slide_idx, new_text):
        if slide_idx < 0 or slide_idx >= len(self.presentation.slides):
            return
        slide = self.presentation.slides[slide_idx]
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.has_text_frame:
                shape.text = new_text
                break

    def save(self, suffix="_translated"):
        self.edited_ppt_path = self.ppt_path.replace(".pptx", f"{suffix}.pptx")
        self.presentation.save(self.edited_ppt_path)
        return self.edited_ppt_path


# ---------------------------
# 翻訳・トーンアップモデル
# ---------------------------
class TranslatorModel:
    def __init__(self):
        model_name = "Helsinki-NLP/opus-mt-ja-en"  # 翻訳用
        print("Loading model…")
        self.tokenizer = AutoTokenizer.from_pretrained(model_name)
        self.model = AutoModelForSeq2SeqLM.from_pretrained(model_name)
        # 動的量子化で軽量化
        self.model = torch.quantization.quantize_dynamic(
            self.model, {torch.nn.Linear}, dtype=torch.qint8
        )
        print("Model loaded!")

    def translate_text(self, text: str) -> str:
        if not text.strip():
            return ""
        inputs = self.tokenizer(text, return_tensors="pt", truncation=True)
        with torch.no_grad():
            outputs = self.model.generate(**inputs, max_length=512)
        return self.tokenizer.decode(outputs[0], skip_special_tokens=True)

    def tone_up(self, text: str, style_prompt: str = "文体を丁寧に、明るく、トーンアップしてください") -> str:
        if not text.strip():
            return ""
        input_text = f"{style_prompt}: {text}"
        inputs = self.tokenizer(input_text, return_tensors="pt", truncation=True)
        with torch.no_grad():
            outputs = self.model.generate(**inputs, max_length=512)
        return self.tokenizer.decode(outputs[0], skip_special_tokens=True)
