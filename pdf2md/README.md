# PDF 转 Markdown

## 使用方法
1. 首次：运行上级目录 **install.bat**
2. 将 PDF **拖到本文件夹「运行.bat」**
3. 或双击「运行.bat」按提示操作

## 输出
- 默认生成 `.md`（与源文件同目录）
- 每页会导出预览图到 **`{文件名}_images/`** 文件夹，并在 Markdown 中用 `![...](相对路径)` 引用

## 命令行
```bat
cd ..
python core\pdf_tool.py pdf2md "文件路径"
```
