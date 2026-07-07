#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
auto_refresh.py - 看板数据自动刷新脚本
从Excel提取数据 -> 生成JSON+HTML -> 推送GitHub

使用方法:
  1. 自动模式: python auto_refresh.py
     (自动查找桌面上最新的 7月杯数达成看板-墨柠.xlsx)
  2. 指定文件: python auto_refresh.py "C:/path/to/file.xlsx"
"""
import os, sys, json, glob, time, importlib.util
from datetime import datetime

# ==================== 配置 ====================
DESKTOP_DIR = os.path.join(os.path.expanduser("~"), "Desktop")
REPO_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_JSON = os.path.join(REPO_DIR, "dashboard_data.json")
INDEX_HTML = os.path.join(REPO_DIR, "index.html")
UPDATE_SCRIPT = os.path.join(REPO_DIR, "update_dashboard.py")
GEN_SCRIPT = os.path.join(REPO_DIR, "gen_dashboard.py")
XLSX_FILENAME = "7月杯数达成看板-墨柠.xlsx"

# GitHub 配置
GITHUB_USERNAME = "LIU-11"
REPO_NAME = "waterbar"
GITHUB_PAGES_URL = f"https://{GITHUB_USERNAME}.github.io/{REPO_NAME}/"

# Token 从 github_token.txt 读取（该文件已 gitignore，不会推送到 GitHub）
_TOKEN_FILE = os.path.join(REPO_DIR, "github_token.txt")
GITHUB_TOKEN = ""
if os.path.exists(_TOKEN_FILE):
    with open(_TOKEN_FILE, "r", encoding="utf-8") as _f:
        GITHUB_TOKEN = _f.read().strip()

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
            # 尝试以追加模式打开（如果文件被锁定会失败）
            with open(filepath, "a"):
                pass
            return True
        except (PermissionError, OSError):
            time.sleep(1)
    return False


def update_data(xlsx_path):
    """从Excel提取数据，生成JSON和HTML"""
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

    # 2. 加载 gen_dashboard 模块，生成外部数据模式的 HTML + JSON
    try:
        spec2 = importlib.util.spec_from_file_location("gen_dashboard", GEN_SCRIPT)
        gen_mod = importlib.util.module_from_spec(spec2)
        spec2.loader.exec_module(gen_mod)
        gen_mod.generate_html(data, INDEX_HTML, external_data=True)
        log(f"HTML+JSON 生成完成: {INDEX_HTML}")
    except Exception as e:
        log(f"HTML生成失败: {e}", "ERROR")
        return False

    return True


def push_to_github():
    """通过GitHub API推送 index.html 和 dashboard_data.json"""
    if not GITHUB_TOKEN:
        log("未配置GITHUB_TOKEN，跳过推送（数据已更新到本地）", "WARN")
        log(f"请手动查看: {INDEX_HTML}")
        return False

    import base64, urllib.request

    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
    }

    files_to_push = [
        (INDEX_HTML, "index.html"),
        (DATA_JSON, "dashboard_data.json"),
    ]

    for local_path, repo_path in files_to_push:
        if not os.path.exists(local_path):
            log(f"文件不存在: {local_path}", "ERROR")
            continue

        # 读取文件内容并编码
        with open(local_path, "rb") as f:
            content = base64.b64encode(f.read()).decode("ascii")

        # 获取当前文件SHA（已存在的文件需要SHA才能更新）
        url = f"https://api.github.com/repos/{GITHUB_USERNAME}/{REPO_NAME}/contents/{repo_path}"
        sha = None
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req) as resp:
                sha = json.loads(resp.read()).get("sha")
        except Exception:
            pass  # 文件不存在，是新建

        # 上传文件
        payload = json.dumps({
            "message": f"自动刷新 {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "content": content,
            "sha": sha,
            "branch": "main"
        }).encode("utf-8")

        req2 = urllib.request.Request(
            url, data=payload,
            headers={**headers, "Content-Type": "application/json"},
            method="PUT"
        )
        try:
            with urllib.request.urlopen(req2) as resp:
                result = json.loads(resp.read())
                log(f"推送成功: {repo_path} ({len(content)} bytes)")
        except Exception as e:
            log(f"推送失败: {repo_path} - {e}", "ERROR")
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

    # 2. 更新数据
    if not update_data(xlsx_path):
        return 1

    # 3. 推送到GitHub
    if push_to_github():
        log(f"=== 刷新完成！访问: {GITHUB_PAGES_URL} ===")
    else:
        log("数据已更新到本地，但推送GitHub失败", "WARN")

    return 0


if __name__ == "__main__":
    sys.exit(main())
