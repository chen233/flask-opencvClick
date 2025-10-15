from datetime import datetime
import pyautogui
import time
import pyperclip
import platform  # 用于适配不同系统的粘贴快捷键


def auto_input_filename(filename):
    try:
        # 粘贴文件名到输入框
        pyperclip.copy(filename)  # 把文件名复制到剪贴板
        time.sleep(0.3)  # 确保剪贴板同步完成

        # 适配系统粘贴快捷键（Windows/Linux用Ctrl+V，Mac用Command+V）
        if platform.system() == "Darwin":  # Mac系统
            pyautogui.hotkey('command', 'v')
        else:  # Windows/Linux系统
            pyautogui.hotkey('ctrl', 'v')

        print(f"已成功输入文件名：{filename}")
        return True

    except Exception as e:
        print(f"输入文件名失败：{str(e)}")
        return False


auto_input_filename(datetime.now().strftime('%Y-%m-%d_%H-%M-%S'))
