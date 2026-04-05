# Enron 邮件处理工具

一个基于 Tkinter 的 GUI 工具，用于解析 Enron 邮件数据集，提取邮件内容并导出为 Word 文档。

## 功能
- 扫描 Enron 数据集目录，筛选有效邮件文件
- 提取发件人、收件人、主题及纯文本内容（修复编码问题）
- 导出为 Word 文档（.docx）
- 实时显示处理进度和日志

## 依赖
- Python 3.x
- python-docx（自动安装）

## 使用方法
1. 运行 `fix_data.py`
2. 选择 Enron 数据集根目录
3. 选择输出 Word 文档路径
4. 点击“开始处理”
