# Docx-LaTeX 转换工具

一个简洁的 Web 界面，支持 Word 与 LaTeX 文档的双向转换。

## 功能

- **Word → LaTeX**: 将 .docx 文件转换为 .tex 代码
- **LaTeX → Word**: 将 .tex 文件转换为 .docx 文档

## 快速开始

### 1. 安装依赖

```bash
cd docx-tools
pip install -r requirements.txt
```

### 2. 启动服务

```bash
python app.py
```

### 3. 打开浏览器

访问 http://127.0.0.1:5000

## 使用方法

1. 选择对应的转换类型卡片
2. 拖拽或点击上传文件
3. 点击"转换"按钮
4. 下载转换后的文件

## 系统要求

- Python 3.8+
- Flask
- python-docx
- lxml

## 项目结构

```
docx-tools/
├── app.py              # Flask Web 应用
├── requirements.txt   # Python 依赖
└── README.md          # 本文件
```

## 技术说明

- 前端: 纯 HTML/CSS/JavaScript，无需额外依赖
- 后端: Flask 微框架
- 转换引擎: 基于 word2latex 和 latex2word 模块

## 截图预览

界面包含两个卡片：
- **Word → LaTeX**: 蓝色渐变主题，支持数学公式、表格等转换
- **LaTeX → Word**: 青色渐变主题，支持多种 LaTeX 元素

## License

MIT
