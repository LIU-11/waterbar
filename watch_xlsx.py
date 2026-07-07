#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
watch_xlsx.py - 桌面Excel文件实时监控 + 每日定时兜底
1. 检测到 7月杯数达成看板-墨柠.xlsx 被修改后，自动调用 auto_refresh.py 刷新数据
2. 每天 9:20 自动执行一次兜底刷新（防止监控遗漏）

使用方法:
  python watch_xlsx.py

建议设置为开机自启动（见部署手册第四章）。
"""
import os, sys, time, subprocess
from pathlib import Path
from datetime import datetime, date

# ==================== 配置 ====================
DESKTOP_DIR = os.path.join(os.path.expanduser("~"), "Desktop")
XLSX_FILENAME = "7月杯数达成看板-墨柠.xlsx"
REPO_DIR = os.path.dirname(os.path.abspath(__file__))
AUTO_REFRESH = os.path.join(REPO_DIR, "auto_refresh.py")
PYTHON_EXE = sys.executable
LOG_FILE = os.path.join(REPO_DIR, "auto_refresh.log")

# 文件修改后等待时间（秒），确保Excel完全写入
COOLDOWN = 5
# 防抖：同一文件最短刷新间隔（秒）
MIN_INTERVAL = 60
# 每日兜底刷新时间（24小时制）
DAILY_REFRESH_HOUR = 9
DAILY_REFRESH_MINUTE = 20


def log(msg, level="INFO"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except:
        pass


def should_run_daily_refresh(last_daily_date):
    """检查是否到了每日兜底刷新时间"""
    now = datetime.now()
    today = date.today()
    # 如果今天还没执行过，且当前时间 >= 9:20
    if last_daily_date != today:
        if now.hour > DAILY_REFRESH_HOUR or (now.hour == DAILY_REFRESH_HOUR and now.minute >= DAILY_REFRESH_MINUTE):
            return True
    return False


def run_refresh():
    """执行一次刷新"""
    try:
        subprocess.run([PYTHON_EXE, AUTO_REFRESH], encoding="utf-8", timeout=120)
    except subprocess.TimeoutExpired:
        log("刷新超时（120秒），跳过", "WARN")
    except Exception as e:
        log(f"刷新失败: {e}", "ERROR")


def main():
    log("=== 桌面Excel监控启动 ===")
    log(f"监控目录: {DESKTOP_DIR}")
    log(f"监控文件: {XLSX_FILENAME}")
    log(f"每日兜底刷新时间: {DAILY_REFRESH_HOUR:02d}:{DAILY_REFRESH_MINUTE:02d}")

    target_path = os.path.join(DESKTOP_DIR, XLSX_FILENAME)
    last_mtime = 0
    last_refresh = 0
    last_daily_date = None  # 记录上次每日刷新的日期

    # 启动时检查：如果今天还没刷新过且已过9:20，立即执行一次
    if should_run_daily_refresh(last_daily_date):
        log("启动时检测到今日尚未刷新，立即执行")
        run_refresh()
        last_daily_date = date.today()
        last_refresh = time.time()

    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
    except ImportError:
        log("watchdog未安装，使用轮询模式", "WARN")
        # 轮询模式
        while True:
            try:
                # 检查每日兜底
                if should_run_daily_refresh(last_daily_date):
                    log("执行每日兜底刷新")
                    run_refresh()
                    last_daily_date = date.today()
                    last_refresh = time.time()

                # 检查文件变化
                if os.path.exists(target_path):
                    mtime = os.path.getmtime(target_path)
                    now = time.time()
                    if mtime > last_mtime and (now - last_refresh) > MIN_INTERVAL:
                        last_mtime = mtime
                        last_refresh = now
                        log(f"检测到文件变化: {target_path}")
                        time.sleep(COOLDOWN)
                        run_refresh()
                time.sleep(3)
            except KeyboardInterrupt:
                log("监控已停止")
                break
            except Exception as e:
                log(f"监控异常: {e}", "ERROR")
                time.sleep(10)
        return

    # watchdog 模式
    class XlsxHandler(FileSystemEventHandler):
        def on_modified(self, event):
            nonlocal last_mtime, last_refresh
            if event.is_directory:
                return
            if os.path.basename(event.src_path) != XLSX_FILENAME:
                return
            if event.src_path.startswith("~$"):
                return
            try:
                mtime = os.path.getmtime(event.src_path)
            except:
                return
            now = time.time()
            if mtime > last_mtime and (now - last_refresh) > MIN_INTERVAL:
                last_mtime = mtime
                last_refresh = now
                log(f"检测到文件变化: {event.src_path}")
                time.sleep(COOLDOWN)
                run_refresh()

        on_created = on_modified

    observer = Observer()
    handler = XlsxHandler()
    observer.schedule(handler, DESKTOP_DIR, recursive=False)
    observer.start()
    log("watchdog监控已启动，等待文件变化...")

    try:
        while True:
            # 每秒检查每日兜底
            if should_run_daily_refresh(last_daily_date):
                log("执行每日兜底刷新")
                run_refresh()
                last_daily_date = date.today()
                last_refresh = time.time()
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        log("监控已停止")

    observer.join()


if __name__ == "__main__":
    main()
