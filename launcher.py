# launcher.py
import os
import sys
import threading
import time
import webbrowser
import socket
import urllib.request
import tempfile
import textwrap
import subprocess
import traceback

# === 关键：显式导入，让 PyInstaller 知道要把 weekdata_app 打包进去 ===
try:
    import weekdata_app.app_main  # noqa: F401
except Exception:
    # 开发阶段 / 某些环境里可能没装成包，这里不影响运行
    pass


def resource_path(rel_path: str):
    """
    兼容 PyInstaller / 源码：
    - 冻结态：从 _MEIPASS 目录开始拼接
    - 源码态：从当前文件所在目录开始拼接
    """
    if hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, rel_path)


def find_free_port(preferred=8501, max_tries=20):
    for p in range(preferred, preferred + max_tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", p))
                return p
            except OSError:
                continue
    # 兜底：让系统随机分配一个
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def wait_and_open_browser(port: int):
    url = f"http://localhost:{port}"
    # 轮询等待 Streamlit 服务起来
    for _ in range(30):  # 最多 30 秒
        try:
            with urllib.request.urlopen(url, timeout=1):
                webbrowser.open_new_tab(url)
                return
        except Exception:
            time.sleep(1)
    # 实在没起来也打开一次，让用户看到错误页
    webbrowser.open_new_tab(url)


def _get_streamlit_entry_script() -> str:
    """
    返回给 `streamlit run` 用的入口脚本路径：

    优先：
        1. weekdata_app/app_main.py 真实存在 → 直接用它（源码态）
    否则：
        2. 在临时目录生成一个很小的 wrapper.py：
           里面只是从 weekdata_app.app_main 导入 main 并执行。

    这样打包态就不需要依赖物理 app_main.py 文件了。
    """
    # 方案 1：开发环境（源码阶段）
    candidate = resource_path("weekdata_app/app_main.py")
    if os.path.exists(candidate):
        return candidate

    # 方案 2：打包后兜底，用一个临时脚本包装 main()
    wrapper_code = textwrap.dedent("""
        from weekdata_app.app_main import main

        if __name__ == "__main__":
            main()
    """).lstrip()

    tmp = tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8", delete=False, suffix=".py"
    )
    tmp.write(wrapper_code)
    tmp.close()
    return tmp.name


def _get_log_file() -> str:
    """
    出问题时把 traceback 写到一个日志文件里，方便你排查。
    - 源码态：和 launcher.py 同目录
    - 打包态：和 .app 里的可执行文件同目录（Contents/MacOS）
    """
    if getattr(sys, "frozen", False):
        base = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "app_error.log")


def run_streamlit_in_process(port: int):
    """
    打包态：在当前进程内启动 streamlit（PyInstaller 推荐方式）
    """
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENTMODE"] = "false"
    from streamlit.web import cli as stcli

    app_path = _get_streamlit_entry_script()
    sys.argv = [
        "streamlit",
        "run",
        app_path,
        f"--server.port={port}",
        "--server.headless=true",
        "--global.developmentMode=false",
    ]
    stcli.main()


def run_dev_subprocess(port: int):
    """
    源码态：用子进程跑 `python -m streamlit run ...`，方便在命令行调试
    """
    app_path = _get_streamlit_entry_script()
    subprocess.run(
        [
            sys.executable,
            "-m",
            "streamlit",
            "run",
            app_path,
            f"--server.port={port}",
            "--server.headless=true",
        ],
        check=False,
    )


def main():
    port = find_free_port(8501)

    # 先起浏览器等待线程（它会轮询，不会立刻打开失败页）
    threading.Thread(target=wait_and_open_browser, args=(port,), daemon=True).start()

    try:
        # 打包后走 in-process，源码走子进程
        if getattr(sys, "frozen", False):
            run_streamlit_in_process(port)
        else:
            run_dev_subprocess(port)
    except Exception:
        # 避免“黑盒秒退”：把错误写入日志
        log_file = _get_log_file()
        try:
            with open(log_file, "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
        except Exception:
            pass
        # 再抛出，让进程退出（否则可能卡死）
        raise


if __name__ == "__main__":
    main()
