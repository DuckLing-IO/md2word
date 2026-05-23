# MD2Word — Markdown to Word Converter / Markdown 转 Word 转换器

[English](#english) | [中文](#chinese)

---

<a name="english"></a>
## English

A Windows desktop tool that converts Markdown (`.md`) files to Word (`.docx`) documents with full LaTeX math formula preservation. Built with Pandoc + customtkinter.

### Features

- **One-click conversion** — select a `.md` file, choose where to save, click convert
- **LaTeX math preservation** — formulas are converted to Word native OMML format, not images
- **Auto-install Pandoc** — if Pandoc is missing, the app offers to install it automatically via winget or GitHub
- **Modern GUI** — clean interface with progress bar and real-time status feedback
- **Portable EXE** — single 30MB executable, no Python or dependencies required on the target machine

### Quick Start

#### Option 1: Run the EXE (recommended for end users)

Download `MD2Word.exe` from the [Releases](https://github.com/DuckLing-IO/md2word/releases) page, double-click to run.

#### Option 2: Run from source

```bash
pip install customtkinter pypandoc
python md2docx.py
```

Pandoc will be auto-installed on first launch if not detected.

### Build EXE from source

```bash
pip install customtkinter pypandoc pyinstaller
pyinstaller --onefile --windowed --name "MD2Word" --collect-data customtkinter --clean md2docx.py
```

The output will be at `dist/MD2Word.exe`.

### Requirements

| Component | Required for EXE | Required for source |
|-----------|:---:|:---:|
| Windows 10/11 | Yes | Yes |
| Pandoc | Auto-installed | Auto-installed |
| Python 3.9+ | No (bundled) | Yes |
| customtkinter | No (bundled) | Yes |
| pypandoc | No (bundled) | Yes |

### How It Works

```
Markdown (.md) + LaTeX formulas
       |
       v
   Pandoc engine
  (pypandoc bridge)
       |
       v
Word (.docx) with native OMML formulas
```

The app uses `pandoc` with the `tex_math_dollars` and `tex_math_single_backslash` extensions to properly parse LaTeX math delimiters (`$...$`, `$$...$$`, `\(...\)`, `\[...\]`) and convert them to Word's native Office Math Markup Language (OMML).

### Pandoc Auto-Install Flow

```
Startup
  |
  +--[Pandoc found]--> Ready
  |
  +--[Pandoc missing]--> Dialog: "Install now?"
        |
        +--[Yes]--> Try winget install
        |             |
        |             +--[OK]--> Ready
        |             |
        |             +--[Fail]--> Download MSI from GitHub
        |                            |
        |                            +--[OK]--> Ready
        |                            |
        |                            +--[Fail]--> Show manual download link
        |
        +--[No]--> Show manual download link
```

---

<a name="chinese"></a>
## 中文

一款 Windows 桌面工具，用于将 Markdown (`.md`) 文件转换为 Word (`.docx`) 文档，并完整保留 LaTeX 数学公式。基于 Pandoc + customtkinter 构建。

### 功能特性

- **一键转换** — 选择 `.md` 文件，指定保存位置，点击转换
- **LaTeX 公式保留** — 公式转为 Word 原生 OMML 格式，非图片
- **自动安装 Pandoc** — 如未检测到 Pandoc，程序可自动通过 winget 或 GitHub 下载安装
- **现代风格界面** — 简洁界面，带进度条和实时状态反馈
- **便携 EXE** — 单个 30MB 可执行文件，目标电脑无需安装 Python 或任何依赖

### 快速开始

#### 方式一: 直接运行 EXE (推荐)

从 [Releases](https://github.com/DuckLing-IO/md2word/releases) 页面下载 `MD2Word.exe`，双击运行。

#### 方式二: 从源码运行

```bash
pip install customtkinter pypandoc
python md2docx.py
```

首次启动时如未检测到 Pandoc 将自动提示安装。

### 从源码构建 EXE

```bash
pip install customtkinter pypandoc pyinstaller
pyinstaller --onefile --windowed --name "MD2Word" --collect-data customtkinter --clean md2docx.py
```

构建产物在 `dist/MD2Word.exe`。

### 环境要求

| 组件 | EXE 需要 | 源码需要 |
|-----------|:---:|:---:|
| Windows 10/11 | 是 | 是 |
| Pandoc | 自动安装 | 自动安装 |
| Python 3.9+ | 否 (已内嵌) | 是 |
| customtkinter | 否 (已内嵌) | 是 |
| pypandoc | 否 (已内嵌) | 是 |

### 工作原理

```
Markdown (.md) + LaTeX 公式
       |
       v
    Pandoc 引擎
   (pypandoc 桥接)
       |
       v
Word (.docx) 带有原生 OMML 公式
```

程序使用 `pandoc` 配合 `tex_math_dollars` 和 `tex_math_single_backslash` 扩展，正确解析 LaTeX 数学定界符 (`$...$`、`$$...$$`、`\(...\)`、`\[...\]`)，并将其转换为 Word 原生的 Office Math Markup Language (OMML) 格式。

### Pandoc 自动安装流程

```
启动程序
  |
  +--[已找到 Pandoc]--> 就绪
  |
  +--[未找到 Pandoc]--> 弹窗: "是否自动安装?"
        |
        +--[是]--> 尝试 winget 安装
        |            |
        |            +--[成功]--> 就绪
        |            |
        |            +--[失败]--> 从 GitHub 下载 MSI
        |                           |
        |                           +--[成功]--> 就绪
        |                           |
        |                           +--[失败]--> 显示手动下载链接
        |
        +--[否]--> 显示手动下载链接
```

### 项目结构

```
md2word/
  md2docx.py       — 完整源代码 (单文件)
  dist/
    MD2Word.exe    — 打包好的可执行文件
  build/           — PyInstaller 构建中间产物
  MD2Word.spec     — PyInstaller 规格文件
  CLAUDE.md        — Claude Code 项目文档
  README.md        — 本文件
```

### License

MIT
