# DocFlow智能填表系统 

一个基于LLM的智能文档理解与自动填表系统，支持从多个数据源文档中提取信息，自动填写 Excel 或 Word 模板。

## 功能特性

- 📂 多格式数据源支持：.docx、.xlsx、.md、.txt
- 📋 模板自动填表：支持 .xlsx 和 .docx 格式
- 🤖 AI 智能提取：基于 DeepSeek API 的自然语言理解
- 🚀 Web 界面：简洁易用的前端交互
- ⚡ 高效处理：支持大文件分块处理和并发提取

## 系统要求

- Python 3.8+
- pip

## 快速开始

### 1. 安装依赖

//cd 项目存放路径
cd project
//安装依赖
pip install -r requirements.txt


### 2. 启动服务
cd backend
python main.py

服务将在 `http://localhost:8000` 启动

### 3. 打开 Web 界面

在浏览器中访问：`http://localhost:8000`

## 使用流程

### 步骤 1：上传数据源文档
- 点击或拖拽上传包含数据的文档（支持多个）
- 系统自动提取文本内容
- 显示已上传文件数量和字符数
⚡记得点击“上传并提取文本”，再次上传时先点击“清空数据源”，然后再上传

### 步骤 2：上传模板并填表
- 上传需要填写的模板文件（.xlsx 或 .docx）
- 可选：输入用户要求（如日期范围、特定条件等）
- 点击"开始填表"按钮
- 系统自动分析数据源并填写模板
- 下载填写完成的文件

## 项目结构一览

```
project/
├── backend/
│   ├── main.py              # FastAPI 主入口
│   ├── agent.py             # AI 数据提取逻辑
│   ├── extractor.py         # 文档解析模块
│   ├── filler.py            # 表格填写模块
│   └── uploads/             # 上传文件存储
├── frontend/
│   ├── index.html           # Web 界面
│   └── app.js               # 前端交互逻辑
├── outputs/                 # 输出文件存储
├── requirements.txt         # Python 依赖
└── README.md               # 本文件
```

## 核心模块说明

### agent.py
- `_call()`: 调用 DeepSeek API
- `_parse_json()`: 解析 API 返回的 JSON
- `_analyze_keywords()`: 分析用户要求和模板结构
- `_filter_text()`: 按关键词过滤文本
- `_extract()`: 从文本块中提取结构化数据
- `extract_and_fill()`: 主流程函数

### extractor.py
- `extract_file()`: 通用文件提取
- `extract_text_from_docx()`: Word 文档提取
- `extract_text_from_xlsx()`: Excel 文件提取
- `get_xlsx_structure()`: 获取 Excel 模板结构
- `get_docx_structure()`: 获取 Word 模板结构

### filler.py
- `fill_xlsx()`: 填写 Excel 模板
- `fill_docx()`: 填写 Word 模板

## API 接口

### POST /api/upload-sources
上传数据源文档

**请求：** multipart/form-data，files 字段包含多个文件

**响应：**
```json
{
  "uploaded": 2,
  "total_sources": 2,
  "details": [
    {
      "filename": "data.xlsx",
      "status": "成功",
      "chars": 5000
    }
  ]
}
```

### POST /api/fill-template
上传模板并自动填表

**请求：** multipart/form-data
- `template`: 模板文件
- `requirement`: 用户要求（可选）

**响应：**
```json
{
  "status": "成功",
  "output_file": "filled_abc123_template.xlsx",
  "elapsed_seconds": 12.5,
  "download_url": "/api/download/filled_abc123_template.xlsx"
}
```

### GET /api/download/{filename}
下载填写完成的文件

### GET /api/sources
查看已上传的数据源列表

### DELETE /api/sources
清空所有数据源

## 可能遇到的问题与解决办法

### Q: API Key 在哪里获取？
A: 访问 [DeepSeek 官网](https://platform.deepseek.com) 注册账户并获取 API Key；当前项目中硬编码了本队的API KEY，无需再配置

### Q: 支持哪些文件格式？
A: 
- 数据源：.docx、.xlsx、.xls、.md、.txt
- 模板：.xlsx、.xls、.docx

### Q: 如何处理大文件？
A: 系统自动对大于 1MB 的 Excel 文件进行按需过滤，避免内存溢出

### Q: 填表失败怎么办？
A: 
1. 检查 API Key 是否正确设置
2. 检查网络连接
3. 查看后端日志获取详细错误信息
4. 确保模板文件格式正确

## 技术栈

- 后端：FastAPI + Python 3.8+
- 前端：HTML5 + Vanilla JavaScript
- 文档处理：python-docx、openpyxl
- AI 服务：DeepSeek API
- 并发处理：ThreadPoolExecutor

## 性能优化

- 大文件自动分块处理（2000 字符/块）
- 并发提取（最多 6 个线程）
- 智能关键词过滤，减少 API 调用
- 结果去重，避免重复数据

## 注意事项

1. API 配额：DeepSeek API 有调用限制，请合理使用
2. 文件大小：建议单个文件不超过 50MB
3. 隐私保护：上传的文件会临时存储在服务器，处理完成后自动删除
4. 网络要求：需要能访问 DeepSeek API 服务

