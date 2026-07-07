#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
watch_xlsx.py - 桌面Excel文件实时监控
检测到 7月杯数达成看板-墨柠.xlsx 被修改后，自动调用 auto_refresh.py 刷新数据

使用方法:
  python watch_xlsx.py

建议设置为开机自启动（见部署手册第四章）。
"""
import os, sys, time, subprocess
from pathlib import Path

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


def log(msg, level="INFO"):
    from datetime import datetime
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except:
        pass


def main():
    log("=== 桌面Excel监控启动 ===")
    log(f"监控目录: {DESKTOP_DIR}")
    log(f"监控文件: {XLSX_FILENAME}")

    target_path = os.path.join(DESKTOP_DIR, XLSX_FILENAME)
    last_mtime = 0
    last_refresh = 0

    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
    except ImportError:
        log("watchdog未安装，使用轮询模式", "WARN")
        log("安装watchdog可获得更好的性能: pip install watchdog")
        # 轮询模式
        while True:
            try:
                if os.path.exists(target_path):
                    mtime = os.path.getmtime(target_path)
                    now = time.time()
                    if mtime > last_mtime and (now - last_refresh) > MIN_INTERVAL:
                        last_mtime = mtime
                        last_refresh = now
                        log(f"检测到文件变化: {target_path}")
                        time.sleep(COOLDOWN)
                        subprocess.run([PYTHON_EXE, AUTO_REFRESH], encoding="utf-8")
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
                subprocess.run([PYTHON_EXE, AUTO_REFRESH], encoding="utf-8")

        on_created = on_modified

    observer = Observer()
    handler = XlsxHandler()
    observer.schedule(handler, DESKTOP_DIR, recursive=False)
    observer.start()
    log("watchdog监控已启动，等待文件变化...")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        log("监控已停止")

    observer.join()


if __name__ == "__main__":
    main()
