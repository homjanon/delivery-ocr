# 送货单 → Excel 自动识别（delivery-ocr）

上传送货单图片，调用视觉大模型自动识别，生成**固定表头** Excel。提供两种形态：

- **网页版**（本目录）：纯前端静态站点（GitHub Pages），无服务器、无运行成本，全部在浏览器本地完成。
- **本地版**（`local/` 目录）：纯本地 Python 脚本，零网页、零 CORS 限制，兼容 Windows 7 / 10，可批量识别整个文件夹。详见 [`local/打包说明.md`](local/打包说明.md)。

---

## 一、网页版（GitHub Pages）

### 功能
- 多图上传（3–4 张，自动压缩至 ≤2000px / JPEG q85，10MB→约 1–2MB）
- 调用 **OpenAI 兼容** 视觉大模型（预设见下表）
- 严格提示词约束输出固定 JSON 字段
- 一键生成 Excel：固定表头（含台头公司名称）+ 数据 + 金额合计
- **日期列固化格式（2026-08-16）**：日期转为真实日期值并锁定单元格格式 `yyyy-mm-dd`（如 `2026-08-05`），微信 / QQ / Excel / WPS / 手机表格打开显示一致，不再出现 QQ 打开日期变样的问题
- **执行总耗时实时计时**：「③ 识别结果预览」标题右侧徽章，从点击识别开始实时跳秒，完成/失败定格

固定表头：

```
单据编号 | 日期 | 台头公司名称 | 名称及规格 | 规格型号/颜色 | 数量 | 单位 | 单价 | 金额 | 备注 | 客户/收货单位
```

### 模型预设（网页版）
页面下拉直接切换；**硅基流动、智谱浏览器直连（原生 CORS）；Agnes 默认经自有 Cloudflare 代理 `proxy.hellohopo.dpdns.org` 加速**：

| 预设 | Key 框 | 端点 | 视觉? | 费用 |
|---|---|---|---|---|
| **Agnes 2.0-Flash（默认）** | Agnes Key | `apihub.agnes-ai.com/v1/chat/completions` | ✅ | 免费（无限期·高峰排队） |
| **硅基流动 Qwen3.5-35B-A3B** | 硅基流动 Key | `api.siliconflow.cn/v1/chat/completions` | ✅ | 付费·便宜 |
| **智谱 GLM-4.6v** | 智谱 Key | `open.bigmodel.cn/api/paas/v4/chat/completions` | ✅ | 付费旗舰（需余额） |

> 默认预设为 **Agnes 2.0-Flash**（免费无限期，高峰偶有排队）；**硅基流动 Qwen3.5-35B-A3B**、**智谱 GLM-4.6v**（付费旗舰）为备用。三个模型均关闭思考模式（`enable_thinking:false`）并设 `max_tokens:8192`，避免 JSON 被截断。
> ⚡ **SSE 流式 + 超时降级（2026-08-08）**：识别改为流式输出（内容边生成边在「③ 识别结果预览」上方滚动显示，体感提速）；**空闲超时 40s**（**指"连续 40s 零数据"而非总时长上限**——一旦收到任何输出就重置计时，慢但持续吐字的批次不会被误切；只有 40s 内一点数据都没有——含"有响应头但无正文"的情况——才中断并自动切换下一个有 Key 的模型）；**最多 3 轮循环切换**（每个有 Key 的模型按顺序反复尝试，任意一轮成功即停，可点「停止」按钮随时中止）；模型偶发输出 JSON 对象而非数组时会自动归一化为数组。
> 📊 **实测参考（2026-08-08）**：Agnes 2.0-Flash 最快（秒级出字）；智谱 GLM-4.6v 可用；硅基流动 Qwen3.5-35B-A3B 可用但偏慢（4 张图一批约 **2 分钟**，属正常，耐心等流式滚动即可）。
> 💡 **报错定位**：每个通道失败都会打印「✗ 通道名 失败：原因」；**401 = Key 无效/已注销**（需到平台重新生成并填入），**403 = 无权限/账户余额不足**（Qwen3.5-35B-A3B、GLM-4.6v 为付费模型，需账户有余额）。
> 历史上曾内置「硅基流动 Qwen3.5-397B-A17B」「商汤 SenseNova」预设，已移除。

### 使用
1. GitHub Pages 打开站点
2. 选择模型预设 → 在对应 Key 框填入 API Key（只需填 1 次，永久记住）
3. 上传送货单图片（可多选）
4. 点「识别并生成 Excel」→ 右侧徽章实时显示耗时 → 预览表格 → 点「下载 Excel」
5. API Key 存浏览器 localStorage（仅本机可见，不上传任何服务器）

### 部署（GitHub Pages）
- 仓库根目录即站点根；Settings → Pages → Source 选 `main` 分支 `/ (root)`
- 推送后自动生效，无需构建

### 关于 CORS（重要）
- **硅基流动** 实测支持浏览器直连：OPTIONS 预检返回 `Access-Control-Allow-Origin` 及 `Allow-Headers`（含 `Authorization`、`Content-Type`），浏览器可直接 `fetch`，**无需代理**。
- **Agnes 默认走自有 Cloudflare 代理**（`https://proxy.hellohopo.dpdns.org/?url=<目标>`）加速；其原生端点 `apihub.agnes-ai.com` 实测也带 CORS 头（`Access-Control-Allow-Origin: *`），可作直连兜底。
- **Agnes 端点域名注意**：必须用 `apihub.agnes-ai.com`（实测 OPTIONS 预检 `Access-Control-Allow-Origin: *`）。其另一域名 `api.agnes-ai.com` 预检返回 404、**无 CORS 头**，浏览器直连会失败，切勿使用。Agnes 2.0-Flash 实测**支持 base64 `data:` 内联图片**（与本项目发送方式一致），且免费无限期开放。
- **商汤 SenseNova 不支持浏览器直连**：其 Token 端点 OPTIONS 预检返回 404，浏览器报 `Failed to fetch`，因此**网页版已移除商汤预设**。若需使用商汤，请改用 [`local/` 本地版](local/打包说明.md)（后端调用无 CORS 限制，商汤可接）。
- 这是从 Cloudflare Workers（国内被墙）切换为国内厂商直连的根本原因：流程更简洁、零基础设施成本。

---

## 二、本地版（local/，可选）

适合不想开网页、要批量处理、或需使用商汤等模型的场景。纯本地 Python，逻辑与网页版一致（同提示词 / 表头 / 合计 / 解析容错）。

| 项目 | 说明 |
|------|------|
| 形态 | 命令行 / Tkinter 桌面 GUI，输出 `送货单_日期.xlsx` |
| 系统 | Windows 7（需 Python 3.9）/ Windows 10（Python 3.9 或 3.11）/ **Android（Termux，直接跑源码）** |
| CORS | 无限制，商汤亦可接 |
| 打包 | PyInstaller 打成 exe 双击运行（见 `local/打包说明.md`） |

> **打包一次、两套通用**：用 **Python 3.9** 打的 exe（可在 Win10 上装 3.9 打包），**Win7 与 Win10 都能直接双击用**，Win7 端无需安装 Python。切勿用 Python 3.11 打包（Win7 跑不了）。
> **安卓**：用 Termux 终端直接跑 `ocr_delivery.py`（无需打包），商汤等任意模型可接；步骤见 `local/打包说明.md` 第五节。

入口与完整用法见 **[`local/打包说明.md`](local/打包说明.md)**。

---

## 文件结构
```
delivery-ocr/
├── index.html                 # 页面 UI（含执行总耗时计时徽章）
├── css/style.css
├── js/app.js                  # 主流程：压缩→调API→解析→生成→计时
├── js/prompt.js               # 严格提示词（固定表头字段）
├── js/excel.js                # SheetJS 封装（台头+表头+数据+合计）
├── vendor/xlsx.full.min.js    # 内置 SheetJS（不走 CDN）
├── local/                     # 本地 Python 版（零网页、零CORS，兼容 Win7）
│   ├── ocr_delivery.py        #   命令行主程序（单文件）
│   ├── gui_delivery.py        #   桌面 GUI 窗口版（Tkinter，零额外依赖）
│   ├── requirements.txt       #   依赖（锁 Win7/Py3.9 兼容版本）
│   ├── .env.example           #   密钥模板
│   ├── .gitignore
│   └── 打包说明.md            #   使用 + exe 打包指南
└── README.md
```

## 安全说明
- API Key 由你手动填入：网页版仅存本机 localStorage；本地版仅存同目录 `.env`，**均不写进代码、不上传服务器**。
- 网页版识别过程全部为浏览器**直连**模型 API，不经过任何代理或第三方服务器。
- 页面 / 工具仅供个人自用，勿公开分享链接。

## 自定义
- 改表头/提示词：编辑 `js/prompt.js` 与 `js/excel.js` 的 `EXCEL_HEADERS`（本地版在 `local/ocr_delivery.py` 内同步定义）
- 换模型接口：网页版 UI 填对应 Base URL / Key / Model；本地版改 `local/ocr_delivery.py` 的 `MODEL_PRESETS`
