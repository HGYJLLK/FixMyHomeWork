# FixMyHomeWork

## 项目简介

- 你有没有遇到过你的同学交给你的作业文件名格式混乱？
- 这是一个基于Flask的Web应用程序，用于批量重命名Word文档。
- 该工具可以从Excel表格中读取学生姓名和学号信息，自动识别Word文件名中的姓名，并将文件重命名为你设置的格式。

## 安装要求

### Python环境
- Python 3.7+

### 依赖包
```bash
pip install flask flask-cors pandas openpyxl pathlib
```

## 使用方法

### 1. 启动程序

```bash
python app.py
```
在浏览器中访问：`http://localhost:5002`

**⚠️ 免责声明**：使用本工具前请务必备份原始文件。开发者不对因使用本工具导致的任何数据损失承担责任。
