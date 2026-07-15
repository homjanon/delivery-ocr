# 送货单 → Excel 自动识别（delivery-ocr）

上传送货单图片，调用视觉大模型自动识别，生成**固定表头** Excel 下载。纯前端静态站点（GitHub Pages），无服务器、无运行成本，全部在浏览器本地完成。

## 功能
- 多图上传（3–4 张，自动压缩至 ≤2000px / JPEG q85，10MB→约 1–2MB）
- 调用任意 **OpenAI 兼容** 视觉大模型（预设来自 portfolio 仓的 3 个模型）
- 严格提示词约束输出固定 JSON 字段
- 一键生成 Excel：固定表头（含台头公司名称）+ 数据 + 金额合计（台头公司即送货方，由 LLM 从图片识别、写在备注右边）
- 固定表头：`单据编号 | 日期 | 客户/收货单位 | 名称及规格 | 规格型号/颜色 | 数量 | 单位 | 单价 | 金额 | 备注 | 台头公司名称`

## 模型预设
页面下拉直接切换，全部为**国内可达厂商、浏览器直连（原生 CORS，无需代理）**：

| 预设 | Key 框 | 端点 | 视觉? | 费用 |
|---|---|---|---|---|
| **智谱 GLM-4.6V-Flash** | 智谱 Key | `open.bigmodel.cn/api/paas/v4/chat/completions` | ✅ | 免费（有时限流） |
| **硅基流动 Qwen3.5-397B-A17B** | 硅基流动 Key | `api.siliconflow.cn/v1/chat/completions` | ✅ | 付费·适中 |
| **硅基流动 Qwen3.5-35B-A3B** | 硅基流动 Key（同框） | `api.siliconflow.cn/v1/chat/completions` | ✅ | 付费·便宜 |

> 智谱 / 硅基流动均经实测：OPTIONS 预检返回 `Access-Control-Allow-Origin`，浏览器可**直连**，国内无需任何代理。
> 默认预设为**智谱 GLM-4.6V-Flash**（免费）。若繁忙，可切换至硅基流动付费模型（CORS 直连、稳定可靠）。

## 使用
1. GitHub Pages 打开站点
2. 选择模型预设 → 在下方对应 Key 框里填入 API Key（只需填 1 次，永久记住）
3. 上传送货单图片（可多选）
4. 点「识别并生成 Excel」→ 预览表格 → 点「下载 Excel」
5. API Key 存浏览器 localStorage（持久化，关页不丢、不清），仅你本机可见，不上传任何服务器

## 部署（GitHub Pages）
- 仓库根目录即站点根；Settings → Pages → Source 选 `main` 分支 `/ (root)`
- 推送后自动生效，无需构建

## 关于 CORS
所有模型厂商（智谱 / 商汤 / 硅基流动）均已实测：OPTIONS 预检返回 `Access-Control-Allow-Origin` 及 `Allow-Headers`（含 `Authorization`、`Content-Type`）。浏览器可**直接 fetch**，**无需任何代理**。这是从 Cloudflare Workers（国内被墙）切换为国内厂商直连的根本原因，完整流程更简洁、零基础设施成本。

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
- 识别过程：全部为浏览器**直连**模型 API，不经过任何代理或第三方服务器

## 自定义
- 改表头/提示词：编辑 `js/prompt.js` 与 `js/excel.js` 的 `EXCEL_HEADERS`
- 换模型接口：站点 UI 填对应 Base URL / Key / Model
