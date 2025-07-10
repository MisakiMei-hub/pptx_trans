from pptx import Presentation
from deep_translator import GoogleTranslator

prs = Presentation("AI.pptx")
translator = GoogleTranslator(source='zh-TW', target='zh-CN')

for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    original_text = run.text
                    if isinstance(original_text, str) and original_text.strip():
                        try:
                            translated = translator.translate(original_text)
                            run.text = str(translated)
                        except Exception as e:
                            print(f"翻译出错：{original_text} -> {e}")


prs.save("translated_output.pptx")
print("✅ 翻译完成，已保存为 translated_output.pptx")

