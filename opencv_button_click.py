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
        # 点击虚拟机管理按钮
        ("img/虚拟机管理/虚拟机管理1.PNG", 1),
        # 点击刷新虚拟机
        ("img/虚拟机管理/虚拟机刷新.PNG", 1),
        # 点击全选
        ("img/虚拟机管理/虚拟机全选.PNG", 1),
        # 虚拟机开机
        ("img/虚拟机管理/虚拟机开机.PNG", 1),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 1800  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)
def close_vm():
    # 需要先打开AB组
    BUTTON_SEQUENCE = [
        # 点击虚拟机管理按钮
        ("img/虚拟机管理/虚拟机管理1.PNG", 1),
        # 点击刷新虚拟机
        ("img/虚拟机管理/虚拟机刷新.PNG", 1),
        # 点击全选
        ("img/虚拟机管理/虚拟机全选.PNG", 1),
        # 虚拟机开机
        ("img/虚拟机管理/虚拟机关机.PNG", 1),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 1800  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)


def open_exe(exe_files):
    # 执行AB组任务,exe_name是A组或B组
    for root, _, files in os.walk(exe_files):
        for file in files:
            # 检查文件名是否以点开头（通常表示隐藏文件）且是exe文件
            if file.lower().endswith('.exe') and not file.startswith('.'):
                exe_path = os.path.join(root, file)
                print(f"找到并打开exe文件: {exe_path}")
                BUTTON_SEQUENCE = [
                    # 点击程序确定按钮
                    ("img/开启程序/确认.PNG", 1),
                    ("img/开启程序/登录.PNG", 1)
                ]
                MATCH_THRESHOLD = 0.85  # 匹配阈值
                TIMEOUT = 1800  # 每个按钮的超时时间(秒)
                # 执行按钮序列
                process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)
                try:
                    # 打开找到的exe文件
                    os.startfile(exe_path)
                except Exception as e:
                    print(f"打开文件时出错: {e}")

def close_exe():
    BUTTON_SEQUENCE = [
        ("img/控制台/控制台1.PNG", 1),
        ("img/控制台/全选.PNG", 1),
        ("img/控制台/刷新.PNG", 1),
        ("img/控制台/关闭游戏.PNG", 1),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 1800  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)

def close_program(window_title):
    """
    关闭指定标题的程序
    参数:
        window_title: 窗口标题（支持模糊匹配）
    """
    hwnd = None

    def callback(wnd, param):
        nonlocal hwnd
        if window_title in win32gui.GetWindowText(wnd) and win32gui.IsWindowVisible(wnd):
            hwnd = wnd
        return True

    win32gui.EnumWindows(callback, None)

    if not hwnd:
        print(f"未找到标题包含 '{window_title}' 的窗口")
        return False

    # 获取窗口对应的进程ID
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    # 根据进程ID打开进程
    process_handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
    if process_handle:
        # 终止进程
        win32api.TerminateProcess(process_handle, 0)
        win32api.CloseHandle(process_handle)
        print(f"程序（窗口标题含'{window_title}'）已关闭")
        return True
    else:
        print(f"无法打开进程以关闭程序（窗口标题含'{window_title}'）")
        return False
def click_AB():
    BUTTON_SEQUENCE = [
        # 点击账号管理
        ("img/账号管理/账号管理1.PNG", 1),
        ("img/账号管理/账号管理全选.PNG", 1),
        ("img/账号管理/账号管理刷新.PNG", 1),
        ("img/账号管理/监控登录.PNG", 1),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 1800  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)

def close_AB():
    BUTTON_SEQUENCE = [
        # 点击账号管理
        ("img/账号管理/账号管理1.PNG", 1),
        ("img/账号管理/账号管理全选.PNG", 1),
        ("img/账号管理/账号管理刷新.PNG", 1),
        ("img/账号管理/停止登录.PNG", 1),
    ]
    MATCH_THRESHOLD = 0.85  # 匹配阈值
    TIMEOUT = 1800  # 每个按钮的超时时间(秒)
    # 执行按钮序列
    process_buttons_in_sequence(BUTTON_SEQUENCE, MATCH_THRESHOLD, TIMEOUT)

