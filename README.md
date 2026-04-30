# InvoiceChecker 发票查验识别系统

基于 Flask + 阿里云百炼视觉大模型的发票自动化处理工具，支持发票真伪查验与批量识别整理，导出 Excel。

## 功能概览

| 功能 | 说明 |
|---|---|
| 发票查验 | 自动提取四要素，调用国家税务局平台核验真伪，识别验证码，流式返回结果 |
| 发票识别整理 | 批量提取发票完整字段（买卖方、金额、税号、商品明细等），导出 Excel |
| 多格式支持 | PDF / OFD / XML / JPG / PNG / WEBP 等 |
| 多页 PDF | 每页识别为一张独立发票 |
| 二维码优先 | 扫描发票防伪二维码直接提取关键字段，无需调用 VLM，节省 token |

## 实现原理

### 发票查验

```
上传文件
  │
  ├─ ① 提取四要素（发票代码、号码、开票日期、金额）
  │     优先级：二维码扫描 → PDF 文字层正则 → XML 正则 → VLM 视觉识别
  │
  ├─ ② Playwright 打开税务局查验平台，自动填写四要素
  │
  ├─ ③ 验证码识别
  │     优先：ddddocr 本地模型（无网络消耗）
  │     兜底：调用 VLM（qwen-vl-max）识别图片验证码
  │     失败自动重试，最多 10 次
  │
  └─ ④ 抓取查验结果，通过 SSE 流式推送到前端
```

### 发票识别整理

```
上传文件（支持批量）
  │
  ├─ PDF
  │   ├─ ① 二维码扫描（pyzbar，优先级最高）
  │   ├─ ② PDF 文字层正则提取（零 token）
  │   ├─ ③ VLM 视觉识别补全（字段不完整时调用）
  │   └─ ④ 二维码数据覆盖关键字段（号码/金额/日期以二维码为准）
  │
  ├─ OFD → easyofd 转图片 → VLM 识别
  │
  ├─ XML → 正则解析（支持多种电子发票 XML 方言）
  │
  └─ 图片（JPG/PNG 等）
      ├─ ① 二维码扫描
      └─ ② VLM 视觉识别 → 二维码数据覆盖

最终：regex + VLM 融合（号码/金额以 regex 为准，买卖方/商品明细等文字字段由 VLM 补全）
      → 汇总所有发票 → 导出 Excel
```

### 验证码识别策略

优先使用 **ddddocr**（本地离线模型，速度快、无 API 消耗），识别失败时自动降级为 **VLM**（`qwen-vl-max`）。整个查验过程最多重试 10 次，并通过 SSE 实时向前端推送进度日志。

## 技术栈

| 层 | 技术 |
|---|---|
| 后端 | Python / Flask |
| 前端 | 单页 HTML（无框架） |
| 浏览器自动化 | Playwright（Chromium） |
| 视觉模型 | 阿里云百炼 Qwen-VL 系列 |
| PDF 解析 | PyMuPDF（fitz） |
| OFD 解析 | easyofd |
| 验证码识别 | ddddocr + VLM 兜底 |
| 二维码扫描 | pyzbar（可选） |
| 数据导出 | pandas + openpyxl |

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

```env
ALIBABA_API_KEY=your_api_key_here
VLM_MODEL=qwen-vl-max
RECOGNIZE_VLM_MODEL=qwen3-vl-flash
```

> API Key 获取：https://bailian.console.aliyun.com/

### 3. 启动服务

```bash
python invoice_app.py
```

默认监听 `http://localhost:5000`，浏览器打开即可使用。

## 模型配置

| 环境变量 | 用途 | 推荐值 |
|---|---|---|
| `VLM_MODEL` | 查验用（验证码识别 + 提取四要素兜底） | `qwen-vl-max` |
| `RECOGNIZE_VLM_MODEL` | 识别整理用（完整字段提取） | `qwen3-vl-flash` |

可选模型：`qwen-vl-max` / `qwen-vl-plus` / `qwen2-vl-72b-instruct` / `qwen3-vl-flash` 等（须为阿里云百炼平台支持的视觉语言模型）。

**修改后需重启服务生效。**

## 注意事项

- `.env` 文件含 API Key，**请勿提交到版本控制**
- 发票查验依赖国家税务局查验平台，请遵守平台使用频率限制
- 本项目仅用于学习研究和合规的发票管理用途
