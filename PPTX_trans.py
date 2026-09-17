import os
import threading
from tkinter import Tk, filedialog, messagebox, StringVar
from tkinter import ttk, Label, Button
from pptx import Presentation
from deep_translator import GoogleTranslator


def translate_ppt(input_file, target_lang, progress_var, status_var):
    try:
        base_name, ext = os.path.splitext(input_file)
        output_file = base_name + f"_trans_{target_lang}.pptx"

        prs = Presentation(input_file)
        translator = GoogleTranslator(source="zh-TW", target=target_lang)

        total_shapes = sum(len(slide.shapes) for slide in prs.slides)
        processed_shapes = 0

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
                processed_shapes += 1
                progress = int((processed_shapes / total_shapes) * 100)
                progress_var.set(progress)
                status_var.set(f"正在翻译... {progress}%")
                root.update_idletasks()

        prs.save(output_file)
        status_var.set("✅ 翻译完成！")
        messagebox.showinfo("完成", f"翻译完成，已保存为：\n{output_file}")

    except Exception as e:
        messagebox.showerror("错误", f"翻译过程中出错：{e}")
        status_var.set("❌ 出错")


def start_translation():
    file_path = filedialog.askopenfilename(filetypes=[("PowerPoint 文件", "*.pptx")])
    if not file_path:
        return

    target_lang = lang_var.get()
    status_var.set(f"开始翻译为 {lang_display[target_lang]} ...")
    progress_var.set(0)

    threading.Thread(
        target=translate_ppt,
        args=(file_path, target_lang, progress_var, status_var),
        daemon=True,
    ).start()


# ----------------- GUI -----------------
root = Tk()
root.title("PPT 翻译器 (繁体 → 任意语言)")
root.geometry("450x250")
root.resizable(False, False)

# 微软风格
style = ttk.Style(root)
style.theme_use("clam")
style.configure("TButton", font=("Segoe UI", 11), padding=6, relief="flat", background="#0078D7", foreground="white")
style.map("TButton",
          background=[("active", "#005A9E")],
          foreground=[("active", "white")])
style.configure("TLabel", font=("Segoe UI", 11))
style.configure("TProgressbar", thickness=18)

# 状态显示
status_var = StringVar(value="请选择一个 PPTX 文件")
progress_var = StringVar(value="0")

Label(root, text="PPT 翻译工具", font=("Segoe UI", 14, "bold")).pack(pady=10)
Label(root, textvariable=status_var).pack(pady=5)

# 翻译语言选择
lang_display = {
    "zh-CN": "简体中文",
    "en": "英语",
    "ja": "日语",
    "ko": "韩语",
    "fr": "法语",
    "de": "德语",
}
lang_var = StringVar(value="zh-CN")
lang_menu = ttk.Combobox(root, textvariable=lang_var, values=list(lang_display.keys()), state="readonly", font=("Segoe UI", 11))
lang_menu.pack(pady=5)
lang_menu.set("zh-CN")

# 进度条
progress_bar = ttk.Progressbar(root, length=350, mode="determinate", maximum=100)
progress_bar.pack(pady=10)
progress_var.trace_add("write", lambda *args: progress_bar.config(value=int(progress_var.get())))

# 按钮
Button(root, text="选择 PPT 文件并翻译", command=start_translation).pack(pady=10)

root.mainloop()
