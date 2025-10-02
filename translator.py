from transformers import AutoModelForSeq2SeqLM, AutoTokenizer
import torch

class TranslatorModel:
    def __init__(self, src_lang="ja", tgt_lang="en"):
        self.src_lang = src_lang
        self.tgt_lang = tgt_lang
        model_name = "staka/fugumt-en-ja"  # 日本語→英語

        # トークナイザーとモデルをロード
        self.tokenizer = AutoTokenizer.from_pretrained(model_name)
        self.model = AutoModelForSeq2SeqLM.from_pretrained(model_name)

        # 動的量子化をモデルに適用（推論を軽量化）
        self.model = torch.quantization.quantize_dynamic(
            self.model,              # 量子化するモデル
            {torch.nn.Linear},       # 量子化対象のレイヤー（線形層）
            dtype=torch.qint8         # 量子化後のデータ型（8ビット整数）
        )

    def translate_text(self, text: str) -> str:
        # トークン化
        inputs = self.tokenizer(text, return_tensors="pt", truncation=True)
        with torch.no_grad():
            # 翻訳生成
            outputs = self.model.generate(**inputs, max_length=512)
        # デコードして文字列で返す
        return self.tokenizer.decode(outputs[0], skip_special_tokens=True)


# # 使用例
# if __name__ == "__main__":
#     translator = TranslatorModel()
#     translated_text = translator.translate_text("こんにちは、元気ですか？")
#     print(translated_text)
