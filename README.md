# PDF 工具箱 — 使用手册

基于 **PyMuPDF 开源版**（`pip install pymupdf`）。每个功能独立文件夹，进入对应目录双击 **`run.bat`**（或 `运行.bat`）即可。

---

## 第一步：安装（必做）

在**本目录**双击 **`install.bat`**，自动安装全部 Python 依赖。  
（不会安装 Python 本体；需本机已有 Python 3.8+）

---

## 功能目录

| 文件夹 | 功能 | 用法 |
|--------|------|------|
| `pdf2word/` | PDF → Word | 拖 `.pdf` 到 `run.bat` |
| `pdf2excel/` | PDF → Excel | 同上 |
| `pdf2ppt/` | PDF → PPT | 同上 |
| `pdf2image/` | PDF → PNG 图片 | 同上，输出到子文件夹 |
| `pdf2md/` | PDF → Markdown | 同上 |
| `pdf2json/` | PDF → JSON | 同上 |
| `pdf2txt/` | PDF → 纯文本 | 同上 |
| `word2pdf/` | Word → PDF | 拖 `.docx` 到 `run.bat` |
| `excel2pdf/` | Excel → PDF | 拖 `.xlsx` |
| `ppt2pdf/` | PPT → PDF | 拖 `.pptx` |
| `image2pdf/` | 图片 → PDF | **推荐拖 `拖入图片.vbs`** 或根目录 **`拖拽-图片转PDF.vbs`**；也可双击 **`双击运行.bat`** |
| `encrypt/` | PDF 加密 | 拖 PDF，输入密码 |
| `decrypt/` | PDF 解密 | 拖加密 PDF，输入密码 |

各文件夹内有 **README.md** 说明细节。

---

## 目录结构

```
├── install.bat          ← 首次必运行
├── requirements.txt
├── README.md            ← 本手册
├── core/                程序代码（勿删）
│   ├── pdf_tool.py
│   └── converters.py
├── common/              公共脚本
│   ├── run.bat          各功能共用启动器
│   ├── install_deps.py
│   └── find_python.bat
├── pdf2word/
│   ├── run.bat          ← 双击或拖文件到这里（运行.bat 相同）
│   └── README.md
├── pdf2excel/
│   …（共 13 个功能文件夹）
└── image2pdf/
```

---

## 说明

- 转换过程会显示**进度条**（按页/按张）。
- 窗口**不会闪退**：结束会暂停，可查看成功或报错。
- 启动前会**自动检查** pymupdf、openpyxl、tqdm 等依赖，缺了会尝试安装。
- **PDF→Word**：每页 PDF 对应 Word 一段内容；有文字则只提取文字，扫描/纯图页才插入整页图片，避免页数膨胀。
- **PDF→Markdown**：除文字外，每页会导出预览图到 `{文件名}_images/` 并在 md 中引用。
- **Word / Excel / PPT → PDF**：均可**不装 Office**（自动**简易模式**）。转换时会显示阶段提示；简易模式下按段/表/页显示进度条。仅保留文字/表格数据/幻灯片文字与部分图片，**版式、图表、动画不保证**。要**高还原**请装 [LibreOffice](https://www.libreoffice.org/)（免费）或 Microsoft Office。
- 旧版 `.doc` / `.xls` / `.ppt` 简易模式不支持，请先另存为 `.docx` / `.xlsx` / `.pptx`。
- **图片 → PDF** 仅需 PyMuPDF，支持多图合并。

---

## 命令行（可选）

```bat
cd 本目录
python core\pdf_tool.py pdf2txt "文件.pdf"
python core\pdf_tool.py image2pdf 图1.png 图2.png
python core\pdf_tool.py -h
```

---

## 安装 LibreOffice（免费，推荐）

用于 **Word/Excel/PPT 高还原转 PDF**（比简易模式更接近原件）。

1. 打开官网：https://www.libreoffice.org/download/download-libreoffice/
2. 选择 **Windows**，下载安装包（约 300MB+）
3. 双击安装，默认选项即可，安装路径一般为：  
   `C:\Program Files\LibreOffice\`
4. 安装完成后**无需登录**，本工具会自动查找 `soffice.exe`（含注册表、非 C 盘安装如 `D:\LibreOffice`）

**若仍提示未检测到：**

1. 运行 **`common\check_libreoffice.bat`** 查看是否识别到路径  
2. 编辑 **`config\libreoffice.txt`**，写入一行安装目录，例如：  
   ```
   D:\LibreOffice
   ```  
3. 或设置系统环境变量 **`LIBREOFFICE_HOME`** = `D:\LibreOffice`（安装目录，非 program 子文件夹）

5. 重新用 `word2pdf\run.bat` 转换

## 安装 Microsoft Office（可选，付费）

若已购买 **Microsoft 365 / Office**：

1. 从 https://www.microsoft.com/microsoft-365 安装 Word、Excel、PowerPoint
2. 在本工具目录打开命令行，执行：  
   `pip install comtypes`
3. 未装 LibreOffice 时，Windows 上会优先尝试用 Office COM 导出 PDF

**不必同时装两套**；有 LibreOffice 即可满足大多数高还原需求。

## 图片多选拖放（image2pdf）

| 方式 | 说明 |
|------|------|
| **`拖入图片.vbs`**（image2pdf 文件夹内） | **最推荐**，多图拖放最稳定 |
| **`拖拽-图片转PDF.vbs`**（根目录） | 同上 |
| **`双击运行.bat`** | 双击多选；或拖到 bat 图标上（部分电脑对 bat 拖放不稳定） |

**.vbs 用法**：把多张图片拖到 **`.vbs` 文件图标** 上，不要拖到文件夹空白处。

## 为什么不能“直接拖到文件夹”？

这是 **Windows 的机制**，不是本工具限制：

| 拖到哪里 | 会发生什么 |
|----------|------------|
| 文件夹**空白处** | 只会复制/移动文件，**不会**自动转换 |
| **`run.bat` 文件图标上** | 才会运行转换（可多图） |
| 根目录 **`拖拽-图片转PDF.vbs`** | **推荐**，多图拖放 |

请把图片拖到 **`.vbs` 或 `.bat` 文件图标** 上，不要只拖进文件夹里。

**PDF 转 Word 失败？**  
→ 含图片 PDF 已增强；请先 `install.bat` 更新依赖后重试 `pdf2word\run.bat`。

---

## 常见问题

**install.bat 报 ot / --version / 乱码？**  
→ 已修复：批处理改为纯英文，请重新下载/覆盖后只运行新版 **install.bat**。

**功能都用不了？**  
→ 先运行根目录 **install.bat**，再进对应文件夹用 **run.bat**。

**提示缺少 openpyxl / docx？**  
→ 重新运行 `install.bat`。

**报错 `unrecognized arguments: xxx.docx`？**  
→ 缺少子命令。请用 **`word2pdf\run.bat`** 拖入 docx，或：  
  `python core\pdf_tool.py word2pdf "F:\Desktop\11.docx"`  
→ 现已支持只传文件路径时**自动识别**（`.docx` → `word2pdf`）。

**为什么 Word 转 PDF 以前提示要装 Office？**  
→ `.docx` 是复杂排版文件，PyMuPDF **不能**像 Word 一样排版，高还原导出只能靠 Word/LibreOffice。  
→ 现已支持**不装 Office 的简易模式**（仅正文文字进 PDF）。  

**Excel/PPT 也要高还原？**  
→ 安装 [LibreOffice](https://www.libreoffice.org/)（免费）或 Office；不装则用简易模式（效果一般）。

**加密/解密？**  
→ 进入 `encrypt` 或 `decrypt` 文件夹，拖 PDF 到 `run.bat`，按提示输入密码。
