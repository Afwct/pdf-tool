# PPT 转 PDF

## 使用方法
1. 首次：运行上级目录 **install.bat**
2. 将 PPT **拖到本文件夹「运行.bat」**
3. 或双击「运行.bat」按提示操作（image2pdf 可多选图）

## 输出
- 默认生成 .pdf（与源文件同目录）

## 进度
- 命令行会显示转换阶段（LibreOffice / Office / 简易模式）
- 简易模式下按幻灯片页显示进度条

## 命令行
```bat
cd ..
python core\pdf_tool.py ppt2pdf "文件路径"
```
