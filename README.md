# 送货单 → Excel 自动识别（delivery-ocr）

上传送货单图片，调用视觉大模型自动识别，生成**固定表头** Excel 下载。纯前端静态站点（GitHub Pages），无服务器、无运行成本，全部在浏览器本地完成。

## 功能
- 多图上传（3–4 张，自动压缩至 ≤2000px / JPEG q85，10MB→约 1–2MB）
- 调用任意 **OpenAI 兼容** 视觉大模型（预设来自 portfolio 仓的 3 个模型）
- 严格提示词约束输出固定 JSON 字段
- 一键生成 Excel：固定表头（含台头公司名称）+ 数据 + 金额合计（台头公司即送货方，由 LLM 从图片识别、写在备注右边）
- 固定表头：`单据编号 | 日期 | 客户/收货单位 | 名称及规格 | 规格型号/颜色 | 数量 | 单位 | 单价 | 金额 | 备注 | 台头公司名称`

## 模型预设
页面下拉直接切换（移植自 `portfolio/scripts/call_llm.py`）：

| 预设 | Key 框 | 端点（经 Worker 代理） | 视觉? |
|---|---|---|---|
| **NVIDIA GLM-5.2**（默认，推荐） | NVIDIA Key | `integrate.api.nvidia.com/v1` | ✅ |
| **NVIDIA Qwen 3.5-397B VLM** | NVIDIA Key（同 Key） | `integrate.api.nvidia.com/v1` | ✅ |
| **硅基流动 Qwen2.5-VL-72B** | 硅基流动 Key | `api.siliconflow.cn/v1` | ✅ |
| **硅基流动 DeepSeek-OCR** | 硅基流动 Key（同 Key） | `api.siliconflow.cn/v1` | ✅（OCR 专用·免费） |

> 前两个共用同一个 NVIDIA Key，免费。硅基流动需在其平台注册获取 Key。
> **所有请求均经自建 Cloudflare Worker（`cors-proxy.homjanon.workers.dev`）转发**——浏览器无法直连 NVIDIA/SiliconFlow（CORS 限制），Worker 负责透传 POST 请求体与 Authorization 头。Worker 源码见 `homjanon.github.io/proxy-worker.js`，需自行部署到 Cloudflare。

## 使用
1. GitHub Pages 打开站点
2. 选择模型预设 → 在下方对应 Key 框里填入 API Key（只需填 1 次，永久记住）
3. 上传送货单图片（可多选）
4. 点「识别并生成 Excel」→ 预览表格 → 点「下载 Excel」
5. API Key 存浏览器 localStorage（持久化，关页不丢、不清），仅你本机可见，不上传任何服务器

## 部署（GitHub Pages）
- 仓库根目录即站点根；Settings → Pages → Source 选 `main` 分支 `/ (root)`
- 推送后自动生效，无需构建

## CORS 代理（必读）
静态站点部署在 `*.github.io`，浏览器的同源策略**禁止**直接 `fetch` `integrate.api.nvidia.com` / `api.siliconflow.cn`（会报 `Failed to fetch` / CORS 错误）。解决方案是经一个自建的 Cloudflare Worker 转发：

1. 源码：`homjanon.github.io/proxy-worker.js`（已升级为支持 POST + 透传 Authorization 头）
2. 部署：Cloudflare 控制台 → Workers → 粘贴代码 → Deploy，得到 `https://cors-proxy.homjanon.workers.dev`
   （或 `wrangler deploy`，仓库根放 `wrangler.toml` 指向该文件）
3. 站点 `js/app.js` 的 `MODEL_PRESETS` 已把三个模型端点写成经该 Worker 转发的完整 URL，**无需再手动配置**
4. 若你换了自己的 Worker 域名，只改 `app.js` 顶部 `WORKER` 常量一处即可

> 免责：个人自用 Worker，请求经它中转（含你的 API Key），请确保 Worker 不被他人滥用。

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
- API Key 由你手动填入、仅存本机 localStorage，**不写进代码、不上传服务器**；页面仅供个人自用，勿公开分享链接
- 识别过程为浏览器 → 你的 Worker → 模型接口，图片不经过本站 GitHub 服务器

## 自定义
- 改表头/提示词：编辑 `js/prompt.js` 与 `js/excel.js` 的 `EXCEL_HEADERS`
- 换模型接口：站点 UI 填对应 Base URL / Key / Model
