"""
Markdown -> Word (.docx) 转换工具
==================================

依赖安装(在终端执行):

    pip install customtkinter pypandoc

功能特性:
  - 使用 Pandoc 将 Markdown 转换为 Word (.docx)
  - 完整保留 LaTeX 数学公式(转为 Word 原生 OMML 公式)
  - 现代风格 GUI 界面(customtkinter)
  - 启动时自动检测 Pandoc, 未安装时可一键自动安装(winget / GitHub MSI)
  - 多线程转换, 界面不卡顿
"""

import os
import sys
import tempfile
import threading
import subprocess
import urllib.request
import json

import customtkinter as ctk
from tkinter import filedialog, messagebox


# ============================================================
# Pandoc 检测
# ============================================================
def _check_pandoc():
    """检测本地是否安装了 Pandoc. 返回 (已安装: bool, 路径或信息: str)"""
    try:
        import pypandoc
        path = pypandoc.get_pandoc_path()
        if path:
            return True, path
        return False, "Pandoc 未找到"
    except ImportError:
        return False, "pypandoc 未安装"
    except Exception as e:
        return False, str(e)


def _try_install_pandoc_winget():
    """尝试通过 winget 安装 Pandoc (Windows 10/11 自带). 返回是否成功."""
    try:
        result = subprocess.run(
            ["winget", "install", "--id", "Pandoc.Pandoc",
             "--accept-source-agreements", "--accept-package-agreements", "--silent"],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode == 0:
            _refresh_pandoc_path()
            ok, _ = _check_pandoc()
            return ok
        return False
    except Exception:
        return False


def _try_install_pandoc_msi(status_callback=None):
    """从 GitHub 下载最新 Pandoc MSI 并静默安装. 返回是否成功."""
    try:
        if status_callback:
            status_callback("正在查询最新 Pandoc 版本...")

        api_url = "https://api.github.com/repos/jgm/pandoc/releases/latest"
        req = urllib.request.Request(api_url, headers={"User-Agent": "md2word"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            release = json.loads(resp.read().decode())

        msi_url = None
        for asset in release.get("assets", []):
            name = asset["browser_download_url"]
            if name.endswith("-windows-x86_64.msi"):
                msi_url = name
                break

        if not msi_url:
            if status_callback:
                status_callback("未找到 Pandoc MSI 下载链接")
            return False

        if status_callback:
            status_callback("正在下载 Pandoc 安装包...")

        tmpdir = tempfile.gettempdir()
        msi_path = os.path.join(tmpdir, "pandoc-installer.msi")
        urllib.request.urlretrieve(msi_url, msi_path)

        if status_callback:
            status_callback("正在安装 Pandoc (可能需要管理员权限)...")

        result = subprocess.run(
            ["msiexec", "/i", msi_path, "/quiet", "/norestart"],
            capture_output=True, text=True, timeout=180,
        )

        try:
            os.remove(msi_path)
        except OSError:
            pass

        if result.returncode == 0:
            _refresh_pandoc_path()
            ok, _ = _check_pandoc()
            return ok
        return False
    except Exception:
        return False


def _refresh_pandoc_path():
    """让 pypandoc 重新扫描 Pandoc 路径."""
    import pypandoc
    pypandoc.pandoc_path = None
    if hasattr(pypandoc.get_pandoc_path, "cache_clear"):
        pypandoc.get_pandoc_path.cache_clear()


# ============================================================
# 转换引擎
# ============================================================
def convert_md_to_docx(input_path: str, output_path: str):
    """使用 pypandoc 将 .md 转换为 .docx. 返回 (成功: bool, 消息: str)"""
    try:
        import pypandoc
        pypandoc.convert_file(
            input_path,
            "docx",
            outputfile=output_path,
            extra_args=[
                "--from=markdown+tex_math_dollars+tex_math_single_backslash",
                "--to=docx",
            ],
        )
        return True, f"转换成功!\n文件已保存至: {output_path}"
    except ImportError:
        return False, "缺少 pypandoc 模块, 请执行: pip install pypandoc"
    except RuntimeError as e:
        msg = str(e)
        if "Pandoc" in msg or "pandoc" in msg:
            return False, "未检测到 Pandoc 引擎. 请重启程序后自动安装."
        return False, f"转换出错:\n{msg}"
    except Exception as e:
        return False, f"转换出错:\n{str(e)}"


# ============================================================
# GUI 应用
# ============================================================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Markdown -> Word 转换器")
        self.geometry("640x460")
        self.minsize(520, 380)

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.input_path = ctk.StringVar()
        self.output_path = ctk.StringVar()
        self.status_text = ctk.StringVar(value="就绪 - 请选择输入文件和保存位置")
        self._pandoc_ready = False

        self._build_ui()
        self.after(500, self._startup_check)

    # ========== UI 布局 ==========
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        # 标题
        title = ctk.CTkLabel(
            self,
            text="Markdown -> Word 转换器",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        title.grid(row=0, column=0, pady=(24, 20), padx=24, sticky="w")

        # 输入文件行
        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=1, column=0, padx=24, pady=(0, 12), sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(input_frame, text="输入文件", width=80).grid(
            row=0, column=0, padx=(12, 8), pady=12, sticky="w"
        )
        ctk.CTkEntry(
            input_frame, textvariable=self.input_path, state="readonly"
        ).grid(row=0, column=1, padx=(0, 8), pady=12, sticky="ew")
        ctk.CTkButton(
            input_frame, text="选择...", width=80, command=self._choose_input
        ).grid(row=0, column=2, padx=(0, 12), pady=12)

        # 输出文件行
        output_frame = ctk.CTkFrame(self)
        output_frame.grid(row=2, column=0, padx=24, pady=(0, 12), sticky="ew")
        output_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(output_frame, text="保存位置", width=80).grid(
            row=0, column=0, padx=(12, 8), pady=12, sticky="w"
        )
        ctk.CTkEntry(
            output_frame, textvariable=self.output_path, state="readonly"
        ).grid(row=0, column=1, padx=(0, 8), pady=12, sticky="ew")
        ctk.CTkButton(
            output_frame, text="选择...", width=80, command=self._choose_output
        ).grid(row=0, column=2, padx=(0, 12), pady=12)

        # 转换按钮
        self.convert_btn = ctk.CTkButton(
            self,
            text="开始转换",
            height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._start_conversion,
        )
        self.convert_btn.grid(row=3, column=0, padx=24, pady=(8, 12), sticky="ew")

        # 进度条
        self.progress = ctk.CTkProgressBar(self, mode="indeterminate")
        self.progress.grid(row=4, column=0, padx=24, pady=(0, 8), sticky="ew")

        # 状态栏
        self.status_label = ctk.CTkLabel(
            self,
            textvariable=self.status_text,
            font=ctk.CTkFont(size=12),
            text_color="gray",
            wraplength=580,
            justify="left",
        )
        self.status_label.grid(row=5, column=0, padx=24, pady=(4, 20), sticky="w")

    # ========== 回调 ==========
    def _choose_input(self):
        path = filedialog.askopenfilename(
            title="选择 Markdown 文件",
            filetypes=[("Markdown 文件", "*.md"), ("所有文件", "*.*")],
        )
        if path:
            self.input_path.set(path)
            if not self.output_path.get():
                base = os.path.splitext(path)[0]
                self.output_path.set(base + ".docx")

    def _choose_output(self):
        path = filedialog.asksaveasfilename(
            title="保存 Word 文件",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def _startup_check(self):
        """启动时检测 Pandoc. 未安装则弹窗引导自动安装."""
        ok, info = _check_pandoc()
        if ok:
            self._pandoc_ready = True
            self.status_text.set(f"Pandoc 已就绪 - {info}")
            return

        self.status_text.set("未检测到 Pandoc 引擎")
        choice = messagebox.askyesno(
            "未检测到 Pandoc",
            "本工具依赖 Pandoc 引擎进行文件转换.\n\n"
            "是否自动下载并安装 Pandoc?\n\n"
            "(安装过程可能需要管理员权限,\n"
            "请在弹出的 UAC 窗口中点击 [是])",
        )

        if choice:
            self._auto_install_pandoc()
        else:
            self.status_text.set(
                "未安装 Pandoc - 可重启程序后自动安装, "
                "或手动前往 https://pandoc.org/installing.html 下载"
            )

    def _auto_install_pandoc(self):
        """后台线程执行 Pandoc 自动安装."""
        self._set_running(True)

        def _install_worker():
            # 优先 winget (Windows 10/11 自带, 快速安静)
            self.after(0, lambda: self.status_text.set("正在通过 winget 安装 Pandoc..."))
            if _try_install_pandoc_winget():
                self.after(0, self._on_install_done)
                return

            # 回退到 MSI 下载安装
            self.after(0, lambda: self.status_text.set("winget 不可用, 改从 GitHub 下载安装..."))
            _try_install_pandoc_msi(
                status_callback=lambda msg: self.after(0, self.status_text.set, msg)
            )
            self.after(0, self._on_install_done)

        threading.Thread(target=_install_worker, daemon=True).start()

    def _on_install_done(self):
        self._set_running(False)
        ok, info = _check_pandoc()
        if ok:
            self._pandoc_ready = True
            self.status_text.set(f"Pandoc 安装成功! - {info}")
            messagebox.showinfo("安装完成", "Pandoc 已安装成功, 现在可以开始转换文件了.")
        else:
            self.status_text.set(
                "Pandoc 自动安装失败 - 请手动前往 "
                "https://pandoc.org/installing.html 下载安装"
            )
            messagebox.showerror(
                "安装失败",
                "Pandoc 自动安装失败.\n\n"
                "请手动前往 https://pandoc.org/installing.html\n"
                "下载并安装 Pandoc 后, 重新启动本程序.",
            )

    def _start_conversion(self):
        src = self.input_path.get().strip()
        dst = self.output_path.get().strip()

        if not src:
            messagebox.showwarning("提示", "请先选择输入 Markdown 文件.")
            return
        if not os.path.isfile(src):
            messagebox.showerror("错误", f"输入文件不存在:\n{src}")
            return
        if not dst:
            messagebox.showwarning("提示", "请先选择保存位置.")
            return
        if not self._pandoc_ready:
            ok, _ = _check_pandoc()
            if not ok:
                messagebox.showwarning(
                    "未检测到 Pandoc",
                    "Pandoc 引擎未安装, 无法执行转换.\n\n"
                    "请重启程序, 启动时会自动提示安装.",
                )
                return
            self._pandoc_ready = True

        self._set_running(True)
        self.status_text.set("转换中, 请稍候...")

        def _run():
            success, msg = convert_md_to_docx(src, dst)
            self.after(0, lambda: self._on_done(success, msg))

        threading.Thread(target=_run, daemon=True).start()

    def _on_done(self, success: bool, message: str):
        self._set_running(False)
        if success:
            self.status_text.set(message)
            messagebox.showinfo("完成", message)
        else:
            self.status_text.set(f"转换失败 - {message}")
            messagebox.showerror("转换失败", message)

    def _set_running(self, running: bool):
        if running:
            self.convert_btn.configure(state="disabled", text="工作中...")
            self.progress.start()
        else:
            self.convert_btn.configure(state="normal", text="开始转换")
            self.progress.stop()


# ============================================================
# 入口
# ============================================================
if __name__ == "__main__":
    app = App()
    app.mainloop()
