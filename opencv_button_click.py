import cv2
import numpy as np
import mss
import os
import pyautogui
import time
import pyperclip
import platform  # 用于适配不同系统的粘贴快捷键


def auto_input_filename(filename):
    try:
        # 粘贴文件名到输入框
        pyperclip.copy(filename)  # 把文件名复制到剪贴板
        time.sleep(1)  # 确保剪贴板同步完成

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

def find_button(template_path, threshold=0.8, timeout=10):
    """查找单个按钮并返回中心坐标"""
    if not os.path.exists(template_path):
        print(f"错误：模板图片 '{template_path}' 不存在！")
        return None

    template = cv2.imread(template_path, 0)
    w, h = template.shape[::-1]
    start_time = time.time()

    with mss.mss() as sct:
        monitor = sct.monitors[1]

        while time.time() - start_time < timeout:
            sct_img = sct.grab(monitor)
            screen = np.array(sct_img)
            screen_gray = cv2.cvtColor(screen, cv2.COLOR_BGR2GRAY)

            result = cv2.matchTemplate(screen_gray, template, cv2.TM_CCOEFF_NORMED)
            locations = np.where(result >= threshold)

            if len(locations[0]) > 0:
                top_left = (locations[1][0], locations[0][0])
                center_x = top_left[0] + w // 2
                center_y = top_left[1] + h // 2
                return (center_x, center_y)

            time.sleep(0.3)

    print(f"超时：未找到 {template_path}")
    return None


def click_button(position, delay_after=1):
    """点击按钮并在之后等待指定时间"""
    if position:
        pyautogui.moveTo(position[0], position[1], duration=0.2)
        pyautogui.click()
        print(f"已点击位置：({position[0]}, {position[1]})")
        time.sleep(delay_after)  # 点击后等待，给界面反应时间
        return True
    return False


def process_buttons_in_sequence(button_sequence, threshold=0.85, timeout=15):
    """按顺序处理按钮列表中的每个按钮"""
    for i, (template_path, delay_after) in enumerate(button_sequence, 1):
        button_pos = find_button(template_path, threshold, timeout)

        # 如果找到按钮并点击成功，则继续下一个
        if click_button(button_pos, delay_after):
            pass

    return True

def open_vm():
    # 需要先打开AB组
    BUTTON_SEQUENCE = [
        # 点击虚拟机管理按钮
        ("img/VMguanli/VMguanli1.PNG", 1),
        # 点击刷新虚拟机
        ("img/VMguanli/VMF5.PNG", 1),
        # 点击全选
        ("img/VMguanli/VMallIn.PNG", 1),
        # 虚拟机开机
        ("img/VMguanli/VMopen.PNG", 600),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 1800  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)
def close_vm():
    # 需要先打开AB组
    BUTTON_SEQUENCE = [
        # 点击虚拟机管理按钮
        ("img/VMguanli/VMguanli1.PNG", 1),
        # 点击刷新虚拟机
        ("img/VMguanli/VMF5.PNG", 1),
        # 点击全选
        ("img/VMguanli/VMallIn.PNG", 1),
        # 虚拟机guan机
        ("img/VMguanli/VMclose.PNG", 60),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 180  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)


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


def outExcel(timeNow):
    BUTTON_SEQUENCE = [
        ("img/kongZhiTai/kongZhiTai1.PNG", 3),
        ("img/kongZhiTai/outTxt.PNG", 3),
        ("img/kongZhiTai/outTxt1.PNG", 3),
        ("img/kongZhiTai/outTxt2.png", 5),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 10  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)
    auto_input_filename(timeNow)
    BUTTON_SEQUENCE1 = [
        ("img/kongZhiTai/save.PNG", 3),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 10  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE1, MATCH_THRESHOLD, TIMEOUT)
    pyautogui.press('esc')
    BUTTON_SEQUENCE2 = [
        ("img/kongZhiTai/cheak.PNG", 3),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 10  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE2, MATCH_THRESHOLD, TIMEOUT)

def click_AB():
    BUTTON_SEQUENCE = [
        # 点击账号管理
        ("img/UserManagement/UserMan1.PNG", 1),
        ("img/UserManagement/Userallin.PNG", 1),
        ("img/UserManagement/UserF5.PNG", 1),
        ("img/UserManagement/Loginjiankong.PNG", 1),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 180  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)

def close_AB():
    BUTTON_SEQUENCE = [
        # 点击账号管理
        ("img/UserManagement/UserMan1.PNG", 1),
        ("img/UserManagement/Userallin.PNG", 1),
        ("img/UserManagement/UserF5.PNG", 1),
        ("img/UserManagement/StopLogin.PNG", 1),

    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 10  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)


def open_exe(exe_files):
    # 执行AB组任务,exe_name是A组或B组
    for root, _, files in os.walk(exe_files):
        for file in files:
            # 检查文件名是否以点开头（通常表示隐藏文件）且是exe文件
            if file.lower().endswith('.exe') and not file.startswith('.'):
                exe_path = os.path.join(root, "._cache_"+file)
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

def close_exe():
    BUTTON_SEQUENCE = [
        ("img/kongZhiTai/kongZhiTai1.PNG", 1),
        ("img/kongZhiTai/cheak1.PNG", 1),
        ("img/kongZhiTai/cheak.PNG", 1),
        ("img/kongZhiTai/allIn.PNG", 1),
        ("img/kongZhiTai/F5.PNG", 5),
        ("img/kongZhiTai/closeGame.PNG", 5),
        ("img/kongZhiTai/closeGame1.PNG", 5),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 10  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)
