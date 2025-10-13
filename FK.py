from flask import Flask, render_template, request, jsonify
from apscheduler.schedulers.background import BackgroundScheduler
import datetime
import json
import os
from functools import wraps
import time
import psutil
import opencv_button_click

app = Flask(__name__)
scheduler = BackgroundScheduler()
scheduler.start()

# 配置文件路径
LOG_FILE = "log.txt"
STATUS_FILE = "status.json"
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
A_name = desktop_path + "\A组"
B_name = desktop_path + "\B组"


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
                "time4": "00:00"
            }
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
        if running_vms:
            print(f"检测到运行中的VM程序：{', '.join(running_vms)}")
        else:
            print("未检测到运行中的VM程序")
            opencv_button_click.open_vm()
        task_result = f"[{desc}] 业务逻辑执行完成"  # 模拟业务结果
    elif desc == "A组停止":
        opencv_button_click.close_exe(A_name)
        task_result = f"[{desc}] else业务逻辑执行完成"  # 模拟业务结果
    elif desc == "B组开始":
        opencv_button_click.open_exe(B_name)
        if running_vms:
            print(f"检测到运行中的VM程序：{', '.join(running_vms)}")
        else:
            print("未检测到运行中的VM程序")
            opencv_button_click.open_vm()
        task_result = f"[{desc}] 业务逻辑执行完成"  # 模拟业务结果
    elif desc == "B组结束":
        opencv_button_click.close_exe(B_name)
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
                "time4": "23:59"
            })
        })
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})


if __name__ == '__main__':
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    A_name = desktop_path + "\A组"
    B_name = desktop_path + "\B组"
    print(A_name)
    app.run(debug=False)
