import os
import time

from opencv_button_click import process_buttons_in_sequence

desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
A_name = desktop_path + "\A组"
B_name = desktop_path + "\B组"

for root, _, files in os.walk(A_name):
    for file in files:
        # 检查文件名是否以点开头（通常表示隐藏文件）且是exe文件
        if file.lower().endswith('.exe') and not file.startswith('.'):
            exe_path = os.path.join(root, "._cache_" + file)
            print(f"找到并打开exe文件: {exe_path}")
            try:
                # 打开找到的exe文件
                os.startfile(exe_path)
                time.sleep(15)
                from FK import set_window_topmost

                set_window_topmost("必读公告", topmost=True)
                BUTTON_SEQUENCE = [
                    # 点击程序确定按钮
                    ("img/openexe/yes.PNG", 3),
                    ("img/openexe/login.PNG", 3)
                ]
                MATCH_THRESHOLD = 0.85  # 匹配阈值
                TIMEOUT = 180  # 每个按钮的超时时间(秒)
                # 执行按钮序列
                process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)
            except Exception as e:
                print(f"打开文件时出错: {e}")