import hashlib
import time
from flask import Flask, render_template, request, jsonify
from apscheduler.schedulers.background import BackgroundScheduler
import datetime
import json
import os
from functools import wraps
import psutil
import excelChange
import opencv_button_click
import platform
import win32gui
import win32con

app = Flask(__name__)
scheduler = BackgroundScheduler()
scheduler.start()

# 配置文件路径
LOG_FILE = "log.txt"
STATUS_FILE = "status.json"
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
A_name = desktop_path + "\A组"
B_name = desktop_path + "\B组"
outTxtName = desktop_path + "\导出信息\\"

# 授权文件（隐藏文件+复杂名称，降低被发现的概率）
LICENSE_FILE = ".app_license_v3.dat"
# 加密盐值（可自定义，增加破解难度）
SECRET_SALT = "chen233"


def encrypt(text):
    """简单加密：盐值+MD5哈希（不可逆，适合存储时间戳）"""
    return hashlib.md5(f"{text}_{SECRET_SALT}".encode()).hexdigest()


def get_file_create_time(file_path):
    """获取文件创建时间（作为删除后重置的备用校验）"""
    if os.path.exists(file_path):
        # Windows系统用st_ctime（创建时间），Linux用st_birthtime
        try:
            return os.stat(file_path).st_birthtime
        except AttributeError:
            return os.stat(file_path).st_ctime
    return None


def check_license():
    # 首次运行：创建授权文件（记录真实时间戳，而非依赖文件系统时间）
    if not os.path.exists(LICENSE_FILE):
        now = datetime.datetime.now()
        timestamp = now.timestamp()  # 真实时间戳（未加密）
        encrypted_time = encrypt(str(timestamp))  # 加密存储，防止篡改
        check_code = encrypt(f"{timestamp}_check")  # 基于真实时间生成校验码

        with open(LICENSE_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "encrypted_time": encrypted_time,  # 加密后的首次运行时间
                "raw_timestamp": timestamp,  # 新增：存储原始时间戳（用于计算）
                "check_code": check_code,
                "registered": False
            }, f)
        return True

    # 已存在授权文件：校验合法性
    try:
        with open(LICENSE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        if data["registered"]:
            return True

        # 校验文件是否被篡改（基于原始时间戳）
        expected_check = encrypt(f"{data['raw_timestamp']}_check")
        if data["check_code"] != expected_check:
            return False

        # 计算试用期（基于文件内存储的原始时间戳，而非文件系统时间）
        first_run = datetime.datetime.fromtimestamp(data["raw_timestamp"])
        days_used = (datetime.datetime.now() - first_run).days
        if days_used > 30:
            return False

        return True

    except (json.JSONDecodeError, KeyError):
        return False

def activate():
    if os.path.exists(LICENSE_FILE):
        try:
            with open(LICENSE_FILE, "r+", encoding="utf-8") as f:
                data = json.load(f)
                data["registered"] = True
                # 基于原始时间戳生成激活校验码
                data["check_code"] = encrypt("activated_" + str(data["raw_timestamp"]))
                f.seek(0)
                f.truncate()
                json.dump(data, f)
            return True
        except:
            return False
    return False

# 初始化文件（首次运行时创建）
def init_files():
    if not os.path.exists(LOG_FILE):
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            f.write("=== 日志记录开始 ===\n")

    if not os.path.exists(STATUS_FILE):
        default_status = {
            "workflow_running": False,
            "config": {
                "time1": "00:00",
                "time2": "00:00",
                "time3": "00:00",
                "time4": "00:00",
                "time5": "00:00",
                "time6": "00:00",
                "time7": "00:00",
                "time8": "00:00",
                "time9": "00:00",
                "time10": "00:00",
                "time11": "00:00",
                "time12": "00:00",
                "shutdown_time": "23:59"  # 添加关机时间配置

            },
            "shutdown_scheduled": False  # 记录是否已设置定时关机
        }
        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            json.dump(default_status, f, ensure_ascii=False, indent=2)
    else:
        with open(STATUS_FILE, 'r', encoding='utf-8') as file:
            try:
                # 读取并解析文件
                data = json.load(file)
            except json.JSONDecodeError as e:
                raise ValueError(f"状态文件格式错误（JSON解析失败）：{str(e)}")


init_files()
def set_window_topmost(window_title, topmost=True):
    # 查找窗口句柄
    hwnd = None
    def callback(wnd, param):
        nonlocal hwnd
        if window_title in win32gui.GetWindowText(wnd) and win32gui.IsWindowVisible(wnd):
            nonlocal hwnd
            hwnd = wnd
        return True

    win32gui.EnumWindows(callback, None)

    if not hwnd:
        print(f"未找到标题包含 '{window_title}' 的窗口")
        return False

    # 设置窗口置顶属性
    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    if topmost:
        # 添加置顶属性
        if not (ex_style & win32con.WS_EX_TOPMOST):
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style | win32con.WS_EX_TOPMOST)
            print(f"窗口 '{window_title}' 已设置为置顶")
    else:
        # 移除置顶属性
        if ex_style & win32con.WS_EX_TOPMOST:
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style & ~win32con.WS_EX_TOPMOST)
            print(f"窗口 '{window_title}' 已取消置顶")

    # 刷新窗口
    win32gui.SetWindowPos(
        hwnd,
        win32con.HWND_TOPMOST if topmost else win32con.HWND_NOTOPMOST,
        0, 0, 0, 0,
        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
    )
    return True

def close_window(window_title):
    """关闭指定标题的窗口（发送正常关闭消息，而非强制终止）"""
    # 复用窗口查找逻辑（和置顶函数保持一致）
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

    # 发送关闭消息（模拟用户点击窗口的关闭按钮）
    # 先发送WM_CLOSE消息（正常关闭）
    win32gui.SendMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    print(f"已向窗口 '{window_title}' 发送关闭消息")

    # 等待窗口关闭（最多等待3秒）
    import time
    for _ in range(30):  # 30*0.1=3秒
        if not win32gui.IsWindow(hwnd):  # 检查窗口是否已关闭
            print(f"窗口 '{window_title}' 已成功关闭")
            return True
        time.sleep(0.1)

    # 如果正常关闭失败，尝试强制关闭
    print(f"窗口 '{window_title}' 未响应关闭消息，尝试强制关闭...")
    win32gui.PostMessage(hwnd, win32con.WM_QUIT, 0, 0)
    return True


# 关机执行函数
def shutdown_machine():
    """执行关机操作，根据操作系统类型适配"""
    # 将标题包含"yoo"的窗口置顶
    set_window_topmost("yoo", topmost=True)
    opencv_button_click.close_AB()
    set_window_topmost("yoo", topmost=False)
    opencv_button_click.close_exe()
    try:
        sys_name = platform.system()
        opencv_button_click.close_vm()
        print('已关闭虚拟机，开始执行关机任务')
        if sys_name == "Windows":
            os.system("shutdown /s /t 60")  # Windows系统，60秒后关机
        elif sys_name == "Linux" or sys_name == "Darwin":  # Linux或macOS
            os.system("shutdown -h +1")  # 1分钟后关机
        write_log("定时关机指令已执行")
    except Exception as e:
        write_log(f"执行关机时出错: {str(e)}")


# 设置定时关机API
@app.route('/api/set_shutdown', methods=['POST'])
def set_shutdown():
    try:
        shutdown_time = request.json.get('time')
        if not shutdown_time:
            return jsonify({"success": False, "message": "请指定关机时间"})

        # 清除现有关机任务
        for job in scheduler.get_jobs():
            if job.name == "shutdown_task":
                scheduler.remove_job(job.id)

        # 添加新的关机任务
        hour, minute = map(int, shutdown_time.split(':'))
        scheduler.add_job(
            shutdown_machine,
            'cron',
            hour=hour,
            minute=minute,
            name="shutdown_task"  # 给任务命名，方便后续管理
        )

        # 更新状态
        status = read_status()
        status["config"]["shutdown_time"] = shutdown_time
        status["shutdown_scheduled"] = True
        write_status(status)
        write_log(f"已设置定时关机: {shutdown_time}")

        return jsonify({"success": True, "message": f"定时关机已设置为 {shutdown_time}"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


# 取消定时关机API
@app.route('/api/cancel_shutdown', methods=['POST'])
def cancel_shutdown():
    try:
        # 清除关机任务
        for job in scheduler.get_jobs():
            if job.name == "shutdown_task":
                scheduler.remove_job(job.id)

        # 更新状态
        status = read_status()
        status["shutdown_scheduled"] = False
        write_status(status)
        write_log("定时关机已取消")

        # 取消系统中已设置的关机计划
        sys_name = platform.system()
        if sys_name == "Windows":
            os.system("shutdown /a")  # 取消Windows关机计划
        elif sys_name == "Linux" or sys_name == "Darwin":
            os.system("shutdown -c")  # 取消Linux/macOS关机计划

        return jsonify({"success": True, "message": "定时关机已取消"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


# 装饰器：确保文件操作的安全性（避免并发写入冲突）
# 修复文件锁装饰器（正确获取文件路径参数）
def file_lock(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        # 从参数中获取文件路径（支持位置参数和关键字参数）
        file_path = None
        # 检查位置参数（优先取第一个字符串参数作为文件路径）
        for arg in args:
            if isinstance(arg, str) and (arg.endswith('.txt') or arg.endswith('.json')):
                file_path = arg
                break
        # 检查关键字参数
        if not file_path:
            for key in ['log_file', 'status_file']:
                if key in kwargs and isinstance(kwargs[key], str):
                    file_path = kwargs[key]
                    break
        # 默认锁文件路径
        file_path = file_path or "temp.lock"

        lock_file = f"{file_path}.lock"
        # 等待锁释放（最多等待5秒，避免死锁）
        wait_time = 0
        while os.path.exists(lock_file) and wait_time < 5:
            time.sleep(0.1)
            wait_time += 0.1
        if os.path.exists(lock_file):
            raise Exception("文件锁定超时，无法获取操作权限")

        try:
            open(lock_file, "w").close()  # 创建锁
            return func(*args, **kwargs)
        finally:
            if os.path.exists(lock_file):
                os.remove(lock_file)  # 释放锁

    return wrapper


# 日志操作：追加日志到TXT文件
@file_lock
def write_log(message, log_file=LOG_FILE):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}\n"
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(log_entry)
    return log_entry


# 状态操作：读取JSON状态文件
@file_lock
def read_status(status_file=STATUS_FILE):
    with open(status_file, "r", encoding="utf-8") as f:
        return json.load(f)


# 状态操作：写入JSON状态文件
@file_lock
def write_status(data, status_file=STATUS_FILE):
    with open(status_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# 首页：渲染前端页面
@app.route('/')
def index():
    return render_template('index.html')


def check_vm_running(vm_process_names):
    """
    检查VM程序是否运行
    :param vm_process_names: VM程序的进程名列表（如["VirtualBoxVM", "vmware-vmx"]）
    :return: 运行中的进程名列表
    """
    running = []
    for proc in psutil.process_iter(['name']):
        try:
            proc_name = proc.info['name']
            # 检查进程名是否包含在目标列表中（不区分大小写）
            for vm_name in vm_process_names:
                if vm_name.lower() in proc_name.lower():
                    running.append(proc_name)
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return list(set(running))  # 去重


# 工作流执行函数（实际业务逻辑）
def workflow_task(desc):
    global task_result
    vm_programs = [
        "VirtualBoxVM",  # VirtualBox虚拟机进程
        "vmware-vmx",  # VMware虚拟机进程
        "qemu-system",  # QEMU/KVM进程
        "hyper-v"  # Hyper-V相关进程
    ]

    running_vms = check_vm_running(vm_programs)
    if running_vms:
        print(f"检测到运行中的VM程序：{', '.join(running_vms)}")
    else:
        print("未检测到运行中的VM程序")

    # ===== 这里添加实际业务逻辑 =====
    # 例如：调用接口、执行脚本、处理数据等
    if desc == "A组开始":
        opencv_button_click.open_exe(A_name)
        set_window_topmost("yoo", topmost=True)
        set_window_topmost("yoo", topmost=False)
        if running_vms:
            print(f"检测到运行中的VM程序：{', '.join(running_vms)}")
        else:
            print("未检测到运行中的VM程序")
            opencv_button_click.open_vm()
        set_window_topmost("yoo", topmost=True)
        set_window_topmost("yoo", topmost=False)
        opencv_button_click.click_AB()
        task_result = f"[{desc}] 业务逻辑执行完成"  # 模拟业务结果
    elif desc == "A组停止":
        set_window_topmost("yoo", topmost=True)
        set_window_topmost("yoo", topmost=False)
        timeNow = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        opencv_button_click.outExcel(outTxtName+timeNow)
        excelChange.txt_to_excel_with_history(outTxtName+timeNow+".txt", "角色数据_历史追踪.xlsx", group="A组")
        opencv_button_click.close_AB()

        opencv_button_click.close_exe()
        close_window("yoo")
        task_result = f"[{desc}] else业务逻辑执行完成"  # 模拟业务结果
    elif desc == "B组开始":
        opencv_button_click.open_exe(B_name)
        set_window_topmost("yoo", topmost=True)
        set_window_topmost("yoo", topmost=False)
        if running_vms:
            print(f"检测到运行中的VM程序：{', '.join(running_vms)}")
        else:
            print("未检测到运行中的VM程序")
            opencv_button_click.open_vm()
        set_window_topmost("yoo", topmost=True)
        set_window_topmost("yoo", topmost=False)
        opencv_button_click.click_AB()
        task_result = f"[{desc}] 业务逻辑执行完成"  # 模拟业务结果
    elif desc == "B组停止":
        set_window_topmost("yoo", topmost=True)
        set_window_topmost("yoo", topmost=False)
        timeNow = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        opencv_button_click.outExcel(outTxtName+timeNow)
        excelChange.txt_to_excel_with_history(outTxtName+timeNow+".txt", "角色数据_历史追踪.xlsx", group="B组")
        opencv_button_click.close_AB()
        opencv_button_click.close_exe()
        close_window("yoo")
        task_result = f"[{desc}] else业务逻辑执行完成"  # 模拟业务结果

    # ==============================

    # 记录详细日志（包含执行结果）
    log = write_log(f"执行任务: {desc} -> {task_result}")
    print(log.strip())  # 控制台打印，方便调试


# 启动工作流
@app.route('/api/start_workflow', methods=['POST'])
def start_workflow():
    try:
        status = read_status()
        if status["workflow_running"]:
            return jsonify({"success": False, "message": "工作流已在运行中"})

        # 清除现有任务
        for job in scheduler.get_jobs():
            scheduler.remove_job(job.id)

        # 添加新任务（从状态文件读取配置）
        config = status["config"]
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time1"].split(':')[0],
            minute=config["time1"].split(':')[1],
            args=["A组开始"]
        )
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time2"].split(':')[0],
            minute=config["time2"].split(':')[1],
            args=["A组停止"]
        )
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time3"].split(':')[0],
            minute=config["time3"].split(':')[1],
            args=["B组开始"]
        )
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time4"].split(':')[0],
            minute=config["time4"].split(':')[1],
            args=["B组停止"]
        )
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time5"].split(':')[0],
            minute=config["time5"].split(':')[1],
            args=["A组开始"]
        )
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time6"].split(':')[0],
            minute=config["time6"].split(':')[1],
            args=["A组停止"]
        )
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time7"].split(':')[0],
            minute=config["time7"].split(':')[1],
            args=["B组开始"]
        )
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time8"].split(':')[0],
            minute=config["time8"].split(':')[1],
            args=["B组停止"]
        )
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time9"].split(':')[0],
            minute=config["time9"].split(':')[1],
            args=["A组停止"]
        )
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time10"].split(':')[0],
            minute=config["time10"].split(':')[1],
            args=["B组开始"]
        )
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time11"].split(':')[0],
            minute=config["time11"].split(':')[1],
            args=["B组停止"]
        )
        scheduler.add_job(
            workflow_task, 'cron',
            hour=config["time12"].split(':')[0],
            minute=config["time12"].split(':')[1],
            args=["B组停止"]
        )
        # 更新状态为运行中
        status["workflow_running"] = True
        write_status(status)
        write_log("工作流已启动")
        return jsonify({"success": True, "message": "工作流启动成功"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


# 停止工作流
@app.route('/api/stop_workflow', methods=['POST'])
def stop_workflow():
    try:
        status = read_status()
        if not status["workflow_running"]:
            return jsonify({"success": False, "message": "工作流未在运行中"})

        # 清除所有任务
        for job in scheduler.get_jobs():
            scheduler.remove_job(job.id)

        # 更新状态为已停止
        status["workflow_running"] = False
        write_status(status)
        write_log("工作流已停止")
        return jsonify({"success": True, "message": "工作流已停止"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


# 保存配置
@app.route('/api/save_config', methods=['POST'])
def save_config():
    try:
        new_config = request.json
        status = read_status()
        status["config"].update(new_config)  # 只更新传入的配置项
        write_status(status)
        write_log(f"配置已更新: {new_config}")
        return jsonify({"success": True, "message": "配置保存成功"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

# 获取最新日志（支持分页，避免一次性加载过大文件）
@app.route('/api/get_log', methods=['GET'])
def get_log():
    try:
        # 读取最后100行日志（可根据需求调整）
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            lines = f.readlines()[-100:]  # 取最后100行
        return jsonify({
            "success": True,
            "logs": [line.strip() for line in lines if line.strip()]
        })
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


# 获取当前状态
# 确保后端接口正确实现
@app.route('/api/get_status', methods=['GET'])
def get_status():
    try:
        status = read_status()  # 读取status.json
        return jsonify({
            "success": True,
            "running": status.get("workflow_running", False),
            "config": status.get("config", {
                "time1": "14:00",
                "time2": "19:00",
                "time3": "21:00",
                "time4": "23:59",
                "time5": "14:00",
                "time6": "19:00",
                "time7": "21:00",
                "time8": "23:59",
                "time9": "19:00",
                "time10": "21:00",
                "time11": "23:59",
                "time12": "23:59",
                "shutdown_time": "23:59"  # 添加关机时间
            }),
            "shutdown_scheduled": status.get("shutdown_scheduled", False)  # 添加关机状态
        })
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


if __name__ == '__main__':
    if not check_license():
        print("提示：试用已到期，请联系作者付款激活")
        print("提示：试用已到期，请联系作者付款激活")
        print("提示：试用已到期，请联系作者付款激活")
        print("提示：试用已到期，请联系作者付款激活")
        time.sleep(200000)
        exit(1)  # 直接退出程序

    # 下面是你的主程序逻辑
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    A_name = desktop_path + "\A组"
    B_name = desktop_path + "\B组"
    app.run(debug=True)
