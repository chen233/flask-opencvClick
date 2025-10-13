import cv2
import numpy as np
import pyautogui
import time
import mss
import os



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
        # 点击程序确定按钮
        ("button1_template.png", 3),
        # 点击虚拟机管理按钮
        ("button2_template.png", 3),
        # 点击刷新虚拟机
        ("button2_template.png", 3),
        # 点击全选
        ("button2_template.png", 3),
        # 虚拟机开机
        ("button2_template.png", 3),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 1800  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)


def open_exe(exe_files):
    # 执行AB组任务,exe_name是A组或B组
    for root, _, files in os.walk(exe_files):
        for file in files:
            if file.lower().endswith('.exe'):
                exe_path = os.path.join(root, file)
                print(f"找到并打开exe文件: {exe_path}")
                try:
                    # 打开找到的exe文件
                    os.startfile(exe_path)
                except Exception as e:
                    print(f"打开文件时出错: {e}")

def close_exe(exe_files):
    BUTTON_SEQUENCE = [
        # 点击程序确定按钮
        ("button1_template.png", 3),
        # 点击虚拟机管理按钮
        ("button2_template.png", 3),
        # 点击刷新虚拟机
        ("button2_template.png", 3),
        # 点击全选
        ("button2_template.png", 3),
        # 虚拟机开机
        ("button2_template.png", 3),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 1800  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)

