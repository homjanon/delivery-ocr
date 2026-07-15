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

| 预设 | Key 框 | 端点 | 视觉? | 国内可用 |
|---|---|---|---|---|
| **硅基流动 Qwen2.5-VL-72B** | 硅基流动 Key | `api.siliconflow.cn/v1`（直连✅） | ✅ | ✅ 直连 |
| **硅基流动 DeepSeek-OCR** | 硅基流动 Key（同 Key） | `api.siliconflow.cn/v1`（直连✅） | ✅（OCR 专用·免费） | ✅ 直连 |
| **NVIDIA GLM-5.2** | NVIDIA Key | `integrate.api.nvidia.com/v1`（需代理） | ✅ | ⚠️ 需国内可达代理 |
| **NVIDIA Qwen 3.5-397B VLM** | NVIDIA Key（同 Key） | `integrate.api.nvidia.com/v1`（需代理） | ✅ | ⚠️ 需国内可达代理 |

> 硅基流动接口原生支持 CORS（预检返回 `Access-Control-Allow-Origin: *`），浏览器可**直连**，无需任何代理，国内立即可用（默认预设已设为硅基流动）。
> NVIDIA 接口 OPTIONS 预检不带 CORS 头，浏览器无法直连，必须走代理。默认走 `cors-proxy.homjanon.workers.dev`，但 **Cloudflare 在大陆常被墙**；如需在国内用 NVIDIA，把 `app.js` 顶部 `WORKER` 常量改为你自建的国内可达代理地址（在阿里云函数计算 / 腾讯云函数部署 `proxy-worker.js`）。

## 使用
1. GitHub Pages 打开站点
2. 选择模型预设 → 在下方对应 Key 框里填入 API Key（只需填 1 次，永久记住）
3. 上传送货单图片（可多选）
4. 点「识别并生成 Excel」→ 预览表格 → 点「下载 Excel」
5. API Key 存浏览器 localStorage（持久化，关页不丢、不清），仅你本机可见，不上传任何服务器

## 部署（GitHub Pages）
- 仓库根目录即站点根；Settings → Pages → Source 选 `main` 分支 `/ (root)`
- 推送后自动生效，无需构建

## 关于 CORS 与代理（国内用户必读）
- **硅基流动 `api.siliconflow.cn`**：预检返回 `Access-Control-Allow-Origin: *`，浏览器可**直连**，无需代理。
- **NVIDIA `integrate.api.nvidia.com`**：OPTIONS 预检不带 CORS 头，浏览器禁止直连，必须走代理转发 POST 体与 Authorization。
- **代理选型**：`proxy-worker.js` 是标准 JS（fetch + CORS 透传），可部署到任何支持 Web 标准的运行时：
  1. Cloudflare Workers（已部署 `cors-proxy.homjanon.workers.dev`）——但 **Cloudflare 在大陆常被墙**，从国内访问会 `Failed to fetch`；
  2. 国内可达替代：阿里云函数计算 FC / 腾讯云云函数 SCF / 自有国内服务器。部署后把 `app.js` 的 `WORKER` 常量改成你的地址（`?url=` 内放完整端点，需 encodeURIComponent）。

> `proxy-worker.js` 源码见 `homjanon.github.io/proxy-worker.js`。个人自用代理，请求经它中转（含你的 API Key），请确保不被他人滥用。

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
- 识别过程：硅基流动模型为浏览器**直连**；NVIDIA 模型经你的代理转发；图片不经过本站 GitHub 服务器

## 自定义
- 改表头/提示词：编辑 `js/prompt.js` 与 `js/excel.js` 的 `EXCEL_HEADERS`
- 换模型接口：站点 UI 填对应 Base URL / Key / Model
