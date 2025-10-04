import win32com.client
import os
import pythoncom
import traceback

def _safe_call(fn, *args, **kwargs):
    try:
        return fn(*args, **kwargs)
    except Exception:
        return None

class PowerPointCOM:
    def __init__(self, path: str):
        pythoncom.CoInitialize()
        self.path = os.path.abspath(path)
        self.app = win32com.client.Dispatch("PowerPoint.Application")
        self.app.Visible = True
        try:
            self.presentation = self.app.Presentations.Open(self.path, WithWindow=True)
        except Exception as e:
            print(f"❌ PowerPointファイルを開けませんでした: {e}")
            self.app.Quit()
            raise

    # ----------------------------------------
    # 再帰的にすべてのテキストを持つ shape を yield
    # ----------------------------------------
    def _iterate_text_shapes(self, shape_or_slide):
        """スライドまたはシェイプから全テキストを再帰取得"""
        if hasattr(shape_or_slide, "Shapes"):
            for s in shape_or_slide.Shapes:
                yield from self._iterate_text_shapes(s)
        elif getattr(shape_or_slide, "Type", None) == 6:  # group
            for s in shape_or_slide.GroupItems:
                yield from self._iterate_text_shapes(s)
        elif hasattr(shape_or_slide, "HasTextFrame") and shape_or_slide.HasTextFrame and shape_or_slide.TextFrame.HasText:
            yield shape_or_slide

    # ----------------------------------------
    # runs コレクションを Python リスト化
    # ----------------------------------------
    def _get_runs_list(self, text_range):
        if text_range is None:
            return None
        runs = []
        try:
            runs_col = text_range.Runs()
            count = int(runs_col.Count)
            for i in range(1, count + 1):
                runs.append(runs_col.Item(i))
        except Exception:
            try:
                for r in text_range.Runs:
                    runs.append(r)
            except Exception:
                return None
        return runs

    # ----------------------------------------
    # スライド全体の段落テキスト抽出（空白保持）
    # ----------------------------------------
    def extract_texts(self):
        slides_text = []
        for slide in self.presentation.Slides:
            slide_paras = []
            for shape in self._iterate_text_shapes(slide):
                try:
                    full = shape.TextFrame.TextRange.Text
                    paras = [p for p in full.split('\r') if p]
                    slide_paras.extend(paras)
                except Exception:
                    continue
            slides_text.append(slide_paras)
        return slides_text

    # ----------------------------------------
    # 翻訳文を runs に分配
    # ----------------------------------------
    def _distribute_translation(self, orig_run_texts, translated):
        n = len(orig_run_texts)
        if n == 0:
            return []
        if not translated:
            return [''] * n

        source_has_space = any(' ' in s for s in orig_run_texts)
        if source_has_space:
            trans_words = translated.split()
            total_words = max(1, sum(len(s.split()) if s.strip() else 1 for s in orig_run_texts))
            assigned = []
            idx = 0
            total = len(trans_words)
            for i, s in enumerate(orig_run_texts):
                if i == n - 1:
                    part = ' '.join(trans_words[idx:]) if idx < total else ''
                else:
                    want = max(1, round((len(s.split()) / total_words) * total))
                    part = ' '.join(trans_words[idx:idx + want])
                    idx += want
                assigned.append(part)
            if idx < total:
                assigned[-1] = (assigned[-1] + ' ' + ' '.join(trans_words[idx:])).strip()
            return assigned
        else:
            total_chars = sum(len(s) for s in orig_run_texts)
            total_chars = total_chars if total_chars > 0 else 1
            L = len(translated)
            assigned = []
            pos = 0
            for i, s in enumerate(orig_run_texts):
                if i == n - 1:
                    part = translated[pos:]
                else:
                    take = max(1, round((len(s) / total_chars) * L))
                    part = translated[pos:pos + take]
                    pos += take
                assigned.append(part)
            return assigned

    # ----------------------------------------
    # テキスト置換本体
    # ----------------------------------------
    def replace_text_preserve_format(self, slide_idx, originals, translations, log_misses=False):
        misses = []
        replaced_count = 0
        slide = self.presentation.Slides(slide_idx + 1)
        idx = 0
        total_units = len(originals)

        for shape in self._iterate_text_shapes(slide):
            try:
                if idx >= total_units:
                    break
                full_text = shape.TextFrame.TextRange.Text
                shape_paras = [p for p in full_text.split('\r') if p]

                runs = self._get_runs_list(shape.TextFrame.TextRange) or []
                combined_runs = ''.join([r.Text for r in runs]) if runs else None

                for p in shape_paras:
                    if idx >= total_units:
                        break
                    orig = originals[idx]
                    trans = translations[idx]
                    success = False

                    # normalize for comparison
                    def norm(s):
                        return s.replace('\u3000', ' ').strip()

                    if combined_runs and orig and (norm(orig) in norm(combined_runs)):
                        start = norm(combined_runs).find(norm(orig))
                        end = start + len(norm(orig))
                        pos = 0
                        run_positions = []
                        for r in runs:
                            rtext = r.Text
                            run_positions.append((pos, pos + len(rtext)))
                            pos += len(rtext)
                        run_indices = [i for i, (spos, epos) in enumerate(run_positions)
                                       if not (epos <= start or spos >= end)]
                        if run_indices:
                            orig_run_texts = [runs[i].Text for i in run_indices]
                            pieces = self._distribute_translation(orig_run_texts, trans)
                            for ri, piece in zip(run_indices, pieces):
                                try:
                                    runs[ri].Text = piece
                                except Exception:
                                    pass
                            success = True

                    if not success:
                        try:
                            if norm(orig) in norm(full_text):
                                new_full = full_text.replace(orig, trans, 1)
                                shape.TextFrame.TextRange.Text = new_full
                                # runs 再取得（PowerPointがrun構造を再生成するため）
                                runs = self._get_runs_list(shape.TextFrame.TextRange) or []
                                combined_runs = ''.join([r.Text for r in runs]) if runs else None
                                full_text = new_full
                                success = True
                        except Exception:
                            success = False

                    if not success:
                        misses.append((orig, "not found in shape"))
                    else:
                        replaced_count += 1
                    idx += 1

            except Exception as e:
                traceback.print_exc()
                continue

        if log_misses and misses:
            print("⚠️ 置換されなかった箇所一覧:")
            for m in misses:
                print(" -", m[0], ":", m[1])

        return replaced_count, misses

    # ----------------------------------------
    # 保存と終了
    # ----------------------------------------
    def save_as(self, suffix="_edited"):
        new_path = self.path.replace(".pptx", f"{suffix}.pptx")
        try:
            self.presentation.SaveAs(new_path)
            print(f"✅ 保存成功: {new_path}")
            return new_path
        except Exception as e:
            print(f"❌ 保存失敗: {e}")
            return None

    def close(self):
        try:
            if self.presentation:
                self.presentation.Close()
            if self.app:
                self.app.Quit()
            pythoncom.CoUninitialize()
        except Exception as e:
            print(f"⚠️ 終了時エラー: {e}")
        finally:
            os.system("taskkill /IM POWERPNT.EXE /F >nul 2>&1")
