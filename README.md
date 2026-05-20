# 元数据清除工具

> 一键清除 Office / WPS / PDF 文件中隐藏的个人和公司信息

![平台](https://img.shields.io/badge/平台-Windows%20%7C%20macOS%20%7C%20Linux-brightgreen)
![Python](https://img.shields.io/badge/Python-3.8+-blue)
![许可](https://img.shields.io/badge/许可-MIT-green)

---

## 功能

**清除的元数据：**

| 清除内容 | 示例 |
|----------|------|
| 作者 / 修改者 | 张三、李四 |
| 公司名 / 管理者 | XX有限公司、王经理 |
| 创建时间 / 修改时间 | 2024-01-15 14:30 |
| 版本号 / 修订号 | 3.1、42 |
| 标题 / 主题 / 描述 | 机密报告、财务数据 |
| 自定义属性（custom.xml） | 第三方插件写入的路径、标识等（直接丢弃） |

---

## 批注和修订记录

工具**能自动检测**文档中的批注和修订记录，但**无法自动清除**（嵌在正文 XML 中，结构复杂）。

检测到批注/修订时，会弹出提示，指导你去 Office 或 WPS 中手动清理：

1. **审阅 → 接受所有修订**（清除修订记录中的作者名）
2. **审阅 → 删除文档中的所有批注**
3. 保存后重新拖入工具处理

---

## 快速开始

### Windows

```
双击 setup.bat  →  安装依赖  →  双击 metadata_cleaner.py
```

如果未安装 Python，脚本会弹出对话框引导你从清华镜像下载（国内速度快）。安装 Python 时**务必勾选**「Add Python to PATH」。

### macOS

```
终端运行 bash setup.sh  →  安装依赖  →  双击 metadata_cleaner.py
```

---

## 使用方法

### 图形界面（单文件 / 批量均可）

双击 `metadata_cleaner.py`：

1. 点击 **「添加文件」**，选择要处理的文件（按住 `Ctrl` 或 `Cmd` 可多选，**支持批量**）
2. 点击 **「清除元数据」**
3. 如有批注/修订会先弹出警告
4. 确认后自动处理，进度条显示实时状态，完成弹窗提示

### 拖拽批量处理

把**一个或多个**文件框选后直接拖到 `metadata_cleaner.py` 图标上，无需打开界面，自动处理全部并显示结果。

### 命令行批量处理

```bash
# 处理当前目录下所有 Word 文档
python3 metadata_cleaner.py *.docx

# 处理所有支持的 Office 和 PDF 文件
python3 metadata_cleaner.py *.docx *.xlsx *.pptx *.wps *.pdf
```

---

## 支持的文件格式

| 类型 | 扩展名 |
|------|--------|
| Word 文档 | `.docx` |
| Excel 表格 | `.xlsx` |
| PowerPoint 演示 | `.pptx` |
| WPS 文字 | `.wps` |
| WPS 表格 | `.et` |
| WPS 演示 | `.dps` |
| PDF 文件 | `.pdf` |

> 旧版二进制格式（`.doc` / `.xls` / `.ppt` / 旧 `.wps`）不支持。请先"另存为"新格式。

---

## 安全设计

| 机制 | 说明 |
|------|------|
| **原子替换** | 先写临时文件，成功后再替换；中途断电原文件无损 |
| **流式处理** | 逐条读写 ZIP，500MB+ 文件不会 OOM |
| **defusedxml** | 安全 XML 解析库，防御 XXE / Billion Laughs 攻击 |
| **分类报错** | 权限不足、格式不支持等均有中文提示 |

---

## 常见问题

<details>
<summary><b>处理后文件会损坏吗？</b></summary>
不会。原子替换机制保证即使写入中途断电，原文件也完好无损。
</details>

<details>
<summary><b>能恢复清除的元数据吗？</b></summary>
不能，清除是永久性的。建议处理敏感文件前先备份。
</details>

<details>
<summary><b>提示"文件正被其他程序占用"？</b></summary>
先关闭 Office / WPS 中打开的该文件，再重试。
</details>

<details>
<summary><b>提示"不是有效的 Office/WPS 文档"？</b></summary>
文件是旧版二进制格式，本工具不支持。请在 Office 中"另存为" .docx 格式。
</details>

<details>
<summary><b>旧格式 .doc / .xls / .ppt 能处理吗？</b></summary>
先用 Office 打开，「文件 → 另存为」 .docx / .xlsx / .pptx，再处理。
</details>

<details>
<summary><b>pip 安装太慢？</b></summary>
setup 脚本已配置清华镜像。如仍慢，手动切换：
<pre>pip config set global.index-url https://mirrors.aliyun.com/pypi/simple/</pre>
</details>
