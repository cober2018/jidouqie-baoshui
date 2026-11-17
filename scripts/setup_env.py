"""准备 skill 用的 Python 环境，成功时只在标准输出打印一行：环境里 python 的路径。
用法（Mac / Linux）: python3 setup_env.py
用法（Windows）    : py -3 setup_env.py   或   python setup_env.py

只在第一次运行时安装，以后秒过。失败时退出码非 0，标准错误里第一行是【错误类型】，
后面是给 AI 助手看的处理办法，详见 references/环境安装.md。

环境放在哪里（按顺序尝试，第一个能写的就用）：
  1. 环境变量 JIDOUQIE_BAOSHUI_VENV 指定的目录
  2. 用户缓存目录：Mac/Linux ~/.cache/jidouqie-baoshui/venv，Windows %LOCALAPPDATA%\\jidouqie-baoshui\\venv
  3. 当前文件夹下的 .jidouqie-baoshui-env（Codex 等沙箱只允许写当前文件夹时用这个）
"""
import os, sys, subprocess

PKGS = ["xlrd", "xlwt", "xlutils", "pdfplumber", "openpyxl"]
IMPORTS = "import xlrd, xlwt, xlutils, pdfplumber, openpyxl"
MIRRORS = [None, "https://pypi.tuna.tsinghua.edu.cn/simple", "https://mirrors.aliyun.com/pypi/simple"]
NET_WORDS = ("Network is unreachable", "Failed to establish", "Temporary failure in name resolution",
             "Name or service not known", "nodename nor servname", "timed out", "ConnectTimeout",
             "ReadTimeout", "No matching distribution", "Could not fetch URL", "ProxyError", "SSLError",
             "Connection refused", "Max retries exceeded", "getaddrinfo failed")


def fail(kind, *lines):
    print(f"【{kind}】", file=sys.stderr)
    for ln in lines: print(ln, file=sys.stderr)
    print("（处理办法见 skill 的 references/环境安装.md）", file=sys.stderr)
    sys.exit(2)


def py_in(venv):
    return os.path.join(venv, "Scripts", "python.exe") if os.name == "nt" else os.path.join(venv, "bin", "python")


def ready(py):
    return os.path.exists(py) and subprocess.run([py, "-c", IMPORTS], capture_output=True).returncode == 0


def candidates():
    if os.environ.get("JIDOUQIE_BAOSHUI_VENV"):
        yield os.environ["JIDOUQIE_BAOSHUI_VENV"]; return   # 指定了就只用这一个
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~\\AppData\\Local")
        yield os.path.join(base, "jidouqie-baoshui", "venv")
    else:
        yield os.path.join(os.environ.get("XDG_CACHE_HOME") or os.path.expanduser("~/.cache"), "jidouqie-baoshui", "venv")
    yield os.path.join(os.getcwd(), ".jidouqie-baoshui-env")


def writable(path):
    d = path
    while not os.path.exists(d):
        parent = os.path.dirname(d)
        if parent == d: return False
        d = parent
    try:
        os.makedirs(path, exist_ok=True)
        probe = os.path.join(path, ".write_test")
        open(probe, "w").close(); os.remove(probe)
        return True
    except OSError:
        return False


def main():
    if sys.version_info < (3, 8):
        fail("Python版本太旧", f"当前是 Python {sys.version.split()[0]}，需要 3.8 或更新的版本。")

    # 已经装好的直接用
    for v in candidates():
        if ready(py_in(v)):
            print(py_in(v)); return

    venv = next((v for v in candidates() if writable(v)), None)
    if not venv:
        fail("没有可写的目录", "缓存目录和当前文件夹都写不进去。",
             "请在有写权限的文件夹里运行，或设置环境变量 JIDOUQIE_BAOSHUI_VENV 指向一个可写目录。")

    py = py_in(venv)
    if not os.path.exists(py):
        r = subprocess.run([sys.executable, "-m", "venv", venv], capture_output=True, text=True)
        if r.returncode != 0:
            fail("缺少venv组件", f"创建虚拟环境失败：{(r.stderr or r.stdout).strip()[-300:]}",
                 "Ubuntu/Debian 需要先装 python3-venv（sudo apt install python3-venv）；",
                 "Mac/Windows 建议用 python.org 官网安装包重新装 Python。")

    last = ""
    for mirror in MIRRORS:
        cmd = [py, "-m", "pip", "install", "-q", "--disable-pip-version-check", "--timeout", "30"] + PKGS
        if mirror:
            cmd += ["-i", mirror]
        print(f"正在安装组件{'（换用镜像 ' + mirror + '）' if mirror else ''}……", file=sys.stderr)
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode == 0 and ready(py):
            print(py); return
        last = (r.stderr or r.stdout).strip()
    if any(w in last for w in NET_WORDS):
        fail("无法联网下载", "下载 xlrd、pdfplumber 等组件失败，官方源和国内镜像都连不上。",
             "如果是在 Codex 等有沙箱的客户端里：请申请联网权限后重新运行这条命令。",
             "否则请检查网络或代理设置。", "最后的报错：" + last[-300:])
    fail("组件安装失败", "最后的报错：" + last[-500:])


if __name__ == "__main__":
    main()
