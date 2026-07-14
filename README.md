# 送货单 → Excel 自动识别（delivery-ocr）

上传送货单图片，调用视觉大模型自动识别，生成**固定表头** Excel 下载。纯前端静态站点（GitHub Pages），无服务器、无运行成本，全部在浏览器本地完成。

## 功能
- 多图上传（3–4 张，自动压缩至 ≤2000px / JPEG q85，10MB→约 1–2MB）
- 调用任意 **OpenAI 兼容** 视觉大模型（预设来自 portfolio 仓的 3 个模型）
- 严格提示词约束输出固定 JSON 字段
- 一键生成 Excel：台头公司名称 + 固定表头 + 数据 + 金额合计
- 固定表头：`单据编号 | 日期 | 客户/收货单位 | 名称及规格 | 规格型号/颜色 | 数量 | 单位 | 单价 | 金额 | 备注`

## 模型预设
页面下拉直接切换（移植自 `portfolio/scripts/call_llm.py`）：

| 预设 | Key 环境变量 | 端点 | 视觉? |
|---|---|---|---|
| **NVIDIA GLM-5.2**（默认，推荐） | `NVIDIA_API_KEY` | `integrate.api.nvidia.com/v1` | ✅ |
| 商汤 DeepSeek-V4-Flash（兜底） | `SENSENOVA_API_KEY` | `token.sensenova.cn/v1` | ❌ 文本 |
| NVIDIA Nemotron-3-Ultra-550B（兜底） | `NVIDIA_API_KEY` | `integrate.api.nvidia.com/v1` | ❌ 文本 |

> 只有 GLM-5.2 是多模态视觉模型，能识别送货单图片。其余两个为文本模型，仅作为备选保留。如需其他模型，选「自定义」手动填写。

## 使用
1. GitHub Pages 打开站点
2. 填写：**API Base URL / API Key / 模型名 / 台头公司名称**
3. 上传送货单图片（可多选）
4. 点「识别并生成 Excel」→ 预览表格 → 点「下载 Excel」
5. API Key 仅存浏览器 sessionStorage（关页即清），不上传任何服务器

## 部署（GitHub Pages）
- 仓库根目录即站点根；Settings → Pages → Source 选 `main` 分支 `/ (root)`
- 推送后自动生效，无需构建

## 文件结构
```
delivery-ocr/
├── index.html          # 页面 UI
├── css/style.css
├── js/app.js           # 主流程：压缩→调API→解析→生成
├── js/prompt.js        # 严格提示词（固定表头字段）
├── js/excel.js         # SheetJS 封装（台头+表头+数据+合计）
├── vendor/xlsx.full.min.js  # 内置 SheetJS（不走 CDN）
└── README.md
```

## 安全说明
- API Key 在前端代码中可见，**个人自用、勿公开分享页面链接**
- 识别过程为浏览器直连你填写的模型接口，图片不上传本站服务器

## 自定义
- 改表头/提示词：编辑 `js/prompt.js` 与 `js/excel.js` 的 `EXCEL_HEADERS`
- 换模型接口：站点 UI 填对应 Base URL / Key / Model
