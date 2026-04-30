# 发票查验识别系统

基于 Flask + 阿里云百炼视觉大模型的发票自动化处理工具，支持发票真伪查验与批量识别整理。

## 功能

- **发票查验**：自动填写国家税务局查验平台，识别验证码，返回查验结果
- **发票识别整理**：支持 PDF / OFD / XML / 图片等格式，提取发票关键字段，导出 Excel
- **多页 PDF 支持**：每页识别为一张独立发票，批量处理
- **二维码优先解析**：扫描发票二维码直接提取字段，无需调用 VLM，节省 token

## 技术栈

| 层 | 技术 |
|---|---|
| 后端 | Python / Flask |
| 前端 | 单页 HTML（无框架） |
| 浏览器自动化 | Playwright |
| 视觉模型 | 阿里云百炼（Qwen-VL 系列） |
| PDF 解析 | PyMuPDF |
| OFD 解析 | easyofd |

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
playwright install chromium
```

可选（安装后自动启用二维码扫描，减少 VLM 调用）：

```bash
pip install pyzbar opencv-python-headless
```

### 2. 配置环境变量

复制模板并填入你的 API Key：

```bash
cp .env.example .env
```

编辑 `.env`：

```
ALIBABA_API_KEY=your_api_key_here
VLM_MODEL=qwen-vl-max
RECOGNIZE_VLM_MODEL=qwen3-vl-flash
```

API Key 获取地址：https://bailian.console.aliyun.com/

### 3. 启动服务

```bash
python invoice_app.py
```

默认监听 `http://localhost:5000`，浏览器打开即可使用。

## 模型说明

| 环境变量 | 用途 | 推荐值 |
|---|---|---|
| `VLM_MODEL` | 发票查验（验证码识别 + 兜底） | `qwen-vl-max` |
| `RECOGNIZE_VLM_MODEL` | 发票识别整理（字段提取） | `qwen3-vl-flash` |

修改后需重启服务生效。模型名称须为阿里云百炼平台支持的视觉语言模型。

## 注意事项

- `.env` 文件含 API Key，**请勿提交到版本控制**
- 发票查验依赖国家税务局查验平台，查验频率请遵守平台限制
- 本项目仅用于学习和合规的发票管理用途
