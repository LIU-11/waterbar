#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
auto_refresh.py - 看板数据自动刷新脚本
从Excel提取数据 -> 生成内嵌HTML -> 复制到deploy_dist/

使用方法:
  1. 自动模式: python auto_refresh.py
     (自动查找桌面上最新的 7月杯数达成看板-墨柠.xlsx)
  2. 指定文件: python auto_refresh.py "C:/path/to/file.xlsx"
"""
import os, sys, json, glob, time, shutil, importlib.util
from datetime import datetime

# ==================== 配置 ====================
DESKTOP_DIR = os.path.join(os.path.expanduser("~"), "Desktop")
REPO_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_JSON = os.path.join(REPO_DIR, "dashboard_data.json")
INDEX_HTML = os.path.join(REPO_DIR, "index.html")
DEPLOY_DIR = os.path.join(REPO_DIR, "..", "deploy_dist")
DEPLOY_HTML = os.path.join(DEPLOY_DIR, "index.html")
UPDATE_SCRIPT = os.path.join(REPO_DIR, "update_dashboard.py")
GEN_SCRIPT = os.path.join(REPO_DIR, "gen_dashboard.py")
XLSX_FILENAME = "7月杯数达成看板-墨柠.xlsx"

# CloudStudio 部署链接
CLOUDSTUDIO_URL = "https://48172fff7d584b119950fc79bd435fb6.app.codebuddy.work"

# 日志文件
LOG_FILE = os.path.join(REPO_DIR, "auto_refresh.log")


def log(msg, level="INFO"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except:
        pass


def find_xlsx():
    """在桌面找到 7月杯数达成看板-墨柠.xlsx"""
    path = os.path.join(DESKTOP_DIR, XLSX_FILENAME)
    if os.path.exists(path) and not path.startswith("~$"):
        return path
    # 备选：找桌面最新的包含"杯数"的xlsx（排除临时文件）
    pattern = os.path.join(DESKTOP_DIR, "*杯数*.xlsx")
    files = [f for f in glob.glob(pattern) if not os.path.basename(f).startswith("~$")]
    if files:
        files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
        return files[0]
    return None


def wait_file_ready(filepath, timeout=30):
    """等待文件写入完成（Excel保存时可能有短暂锁定）"""
    start = time.time()
    while time.time() - start < timeout:
        try:
            with open(filepath, "a"):
                pass
            return True
        except (PermissionError, OSError):
            time.sleep(1)
    return False


def update_data(xlsx_path):
    """从Excel提取数据，生成内嵌HTML"""
    log(f"正在处理: {xlsx_path}")

    if not os.path.exists(UPDATE_SCRIPT):
        log(f"update_dashboard.py 不存在: {UPDATE_SCRIPT}", "ERROR")
        return False
    if not os.path.exists(GEN_SCRIPT):
        log(f"gen_dashboard.py 不存在: {GEN_SCRIPT}", "ERROR")
        return False

    # 等待文件就绪
    if not wait_file_ready(xlsx_path):
        log("文件被锁定，请关闭Excel后重试", "ERROR")
        return False

    # 1. 加载 update_dashboard 模块，提取数据
    try:
        spec = importlib.util.spec_from_file_location("update_dashboard", UPDATE_SCRIPT)
        update_mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(update_mod)
        data = update_mod.extract_data(xlsx_path)
        log(f"数据提取完成: {len(data.get('store_info', {}))} 门店, {len(data.get('available_dates_current', []))} 天")
    except Exception as e:
        log(f"数据提取失败: {e}", "ERROR")
        return False

    # 2. 加载 gen_dashboard 模块，生成内嵌数据模式的 HTML（无fetch）
    try:
        spec2 = importlib.util.spec_from_file_location("gen_dashboard", GEN_SCRIPT)
        gen_mod = importlib.util.module_from_spec(spec2)
        spec2.loader.exec_module(gen_mod)
        gen_mod.generate_html(data, INDEX_HTML, external_data=False)
        log(f"内嵌HTML生成完成: {INDEX_HTML}")
    except Exception as e:
        log(f"HTML生成失败: {e}", "ERROR")
        return False

    # 3. 复制到 deploy_dist/ 供 CloudStudio 部署
    try:
        os.makedirs(DEPLOY_DIR, exist_ok=True)
        shutil.copy2(INDEX_HTML, DEPLOY_HTML)
        log(f"已复制到部署目录: {DEPLOY_HTML}")
    except Exception as e:
        log(f"复制到部署目录失败: {e}", "ERROR")
        return False

    return True


def main():
    log("=== 看板数据自动刷新 ===")

    # 1. 找到xlsx文件
    if len(sys.argv) > 1:
        xlsx_path = sys.argv[1]
    else:
        xlsx_path = find_xlsx()

    if not xlsx_path or not os.path.exists(xlsx_path):
        log(f"未找到xlsx文件，请在桌面放置 {XLSX_FILENAME}", "ERROR")
        return 1

    # 2. 更新数据 + 生成内嵌HTML + 复制到deploy_dist
    if update_data(xlsx_path):
        log(f"=== 刷新完成！CloudStudio部署由WorkBuddy自动化处理 ===")
        log(f"访问: {CLOUDSTUDIO_URL}")
    else:
        log("刷新失败", "ERROR")

    return 0


if __name__ == "__main__":
    sys.exit(main())
