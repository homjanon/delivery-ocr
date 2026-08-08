// 主流程：上传 → 压缩图片 → 调视觉大模型 → 解析 JSON → 预览 → 生成 Excel
// 所有模型均浏览器直连（各厂商原生支持 CORS，已实测预检返回 Access-Control-Allow-Origin），无需任何代理。
(function () {
  const $ = id => document.getElementById(id);
  const status = $("status");
  const logEl = $("log");
  let lastWorkbook = null;

  // 模型预设（baseUrl 已含完整 /chat/completions 端点，调用时直接 fetch(baseUrl)）
  const MODEL_PRESETS = {
    qwen35b: {
      name: "硅基流动 Qwen3.5-35B-A3B（视觉·便宜·直连✅）",
      baseUrl: "https://api.siliconflow.cn/v1/chat/completions",
      model: "Qwen/Qwen3.5-35B-A3B", key: "siliconflow"
    },

    agnes: {
      name: "Agnes 2.0-Flash（视觉·免费·代理加速✅）",
      baseUrl: "https://apihub.agnes-ai.com/v1/chat/completions",
      model: "agnes-2.0-flash", key: "agnes",
      proxy: "https://proxy.hellohopo.dpdns.org/"
    },
  };

  function log(msg) { logEl.textContent += msg + "\n"; logEl.scrollTop = logEl.scrollHeight; }
  function setStatus(msg) { status.textContent = msg; }

  // 预设下拉变更即保存
  $("preset").addEventListener("change", () => localStorage.setItem("do_preset", $("preset").value));

  // 恢复预设（默认 Agnes 2.0-Flash）
  const savedPreset = localStorage.getItem("do_preset");
  $("preset").value = (savedPreset && MODEL_PRESETS[savedPreset]) ? savedPreset : "agnes";

  // 恢复 API Key（持久化）：siliconflow / agnes
  ["siliconflow", "agnes"].forEach(k => {
    const el = $(k + "ApiKey");
    el.value = localStorage.getItem("do_" + k + "Key") || "";
    el.addEventListener("input", () => localStorage.setItem("do_" + k + "Key", el.value));
  });

  // ——— 图片压缩 ———
  function compressImage(file, maxDim = 2000, quality = 0.85) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const img = new Image();
        img.onload = () => {
          let { width, height } = img;
          const scale = Math.min(1, maxDim / Math.max(width, height));
          width = Math.round(width * scale); height = Math.round(height * scale);
          const canvas = document.createElement("canvas");
          canvas.width = width; canvas.height = height;
          canvas.getContext("2d").drawImage(img, 0, 0, width, height);
          resolve(canvas.toDataURL("image/jpeg", quality));
        };
        img.onerror = reject;
        img.src = reader.result;
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  // 预览缩略图
  $("fileInput").addEventListener("change", async (e) => {
    const files = [...e.target.files];
    const prev = $("preview"); prev.innerHTML = "";
    for (const f of files) {
      const url = URL.createObjectURL(f);
      const im = document.createElement("img"); im.src = url; prev.appendChild(im);
    }
  });

  // 从模型回复中提取 JSON 数组
  function extractJSON(text) {
    let t = (text || "").trim();
    t = t.replace(/<think>[\s\S]*?<\/think>/gi, "");
    // 剥离可能不闭合的代码围栏（模型被截断时常见）
    t = t.replace(/^```(?:json)?\s*/i, "").replace(/```\s*$/i, "");
    const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fence) t = fence[1].trim();
    const s = t.indexOf("["), e = t.lastIndexOf("]");
    if (s !== -1 && e !== -1 && e > s) t = t.slice(s, e + 1);
    if (!t) throw new Error("模型未输出任何内容");
    // 截断容错：以 [ 开头但缺结尾 ] 时，尝试补 ] 修复
    if (t.startsWith("[") && !t.endsWith("]")) {
      t = t.replace(/,\s*$/, "").replace(/\s+$/, "") + "]";
    }
    try { return JSON.parse(t); }
    catch (err) {
      throw new Error("JSON 不完整/被截断（可能触及 max_tokens）：" + err.message);
    }
  }

  // 容错归一化：模型偶发输出 JSON 对象而非数组时，自动提取数组字段 / 包成单元素数组
  function normalizeRows(v) {
    if (Array.isArray(v)) return v;
    if (v && typeof v === "object") {
      for (const k of Object.keys(v)) {
        if (Array.isArray(v[k])) return v[k];
      }
      return [v];
    }
    return null;
  }

  // ——— CORS 代理（仅配置了 proxy 的预设使用）：代理?url=<目标> ———
  function proxiedUrl(base, proxy) {
    const p = (proxy || "").trim();
    if (!p) return base;
    const enc = encodeURIComponent(base);
    if (p.indexOf("url=") >= 0) return p + enc;
    return p + (p.indexOf("?") >= 0 ? "&" : "?") + "url=" + enc;
  }

  // ——— 单次调用（SSE 流式）：onToken(t, full) 增量回调 → 流式；否则非流式。90s（流式 120s）超时 ———
  async function callLLMOnce(cfg, messages, opts) {
    const stream = !!opts.onToken;
    const body = { model: cfg.model, messages, temperature: 0 };
    body.max_tokens = 8192;
    if (cfg.thinking) body.thinking = { type: cfg.thinking };
    else body.chat_template_kwargs = { enable_thinking: false };
    body.stream = stream;
    // 空闲超时：流式每收到一块数据即重置计时，只有连续 IDLE ms 零数据（真卡死）才中断
    const ctl = new AbortController();
    const IDLE = stream ? 120000 : 90000;
    let to = null;
    const arm = () => { if (to) clearTimeout(to); to = setTimeout(() => ctl.abort(), IDLE); };
    arm();
    const url = (cfg.proxy && !opts.direct) ? proxiedUrl(cfg.baseUrl, cfg.proxy) : cfg.baseUrl;
    let resp;
    try {
      resp = await fetch(url, { method: "POST", headers: { "Content-Type": "application/json", "Authorization": "Bearer " + cfg.apiKey }, body: JSON.stringify(body), signal: ctl.signal });
    } catch (e) { if (to) clearTimeout(to); return { error: "net", msg: String(e) }; }
    if (to) clearTimeout(to);
    if (!resp.ok) { let t = ""; try { t = await resp.text(); } catch (_) {} return { error: "http", status: resp.status, msg: (t || "").slice(0, 300) }; }
    if (stream) {
      try {
        if (!resp.body) return { error: "parse", msg: "no stream body" };
        const reader = resp.body.getReader();
        const dec = new TextDecoder();
        let buf = "", full = "";
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          arm();
          buf += dec.decode(value, { stream: true });
          let idx;
          while ((idx = buf.indexOf("\n")) >= 0) {
            const line = buf.slice(0, idx).trim(); buf = buf.slice(idx + 1);
            if (line.indexOf("data:") !== 0) continue;
            const data = line.slice(5).trim();
            if (!data || data === "[DONE]") continue;
            try {
              const j = JSON.parse(data);
              const delta = (j.choices && j.choices[0] && j.choices[0].delta && j.choices[0].delta.content) || "";
              if (delta) { full += delta; opts.onToken(delta, full); }
            } catch (e) {}
          }
        }
        if (!full) return { error: "empty" };
        if (opts.json) {
          try { return { content: extractJSON(full) }; }
          catch (e) { return { error: "json", content: full, msg: e.message }; }
        }
        return { content: full };
      } catch (e) { return { error: "net", msg: String(e) }; }
    }
    let j; try { j = await resp.json(); } catch (e) { return { error: "parse", msg: String(e) }; }
    const c = (j && j.choices && j.choices[0] && j.choices[0].message) ? (j.choices[0].message.content || j.choices[0].message.reasoning || "") : "";
    if (c === "") return { error: "empty", raw: j };
    if (opts.json) {
      try { return { content: extractJSON(c) }; }
      catch (e) { return { error: "json", content: c, msg: e.message }; }
    }
    return { content: c };
  }

  // ——— 带自动降级的调用：主通道失败 → 按预设顺序切备用（无 Key 跳过）；末位带代理的通道再补一次直连兜底 ———
  async function callLLM(messages, opts) {
    opts = opts || {};
    const cur = $("preset").value;
    const order = [cur].concat(Object.keys(MODEL_PRESETS).filter(p => p !== cur));
    let lastErr = null;
    for (let i = 0; i < order.length; i++) {
      const p = order[i];
      const p0 = MODEL_PRESETS[p];
      const apiKey = $(p0.key + "ApiKey").value.trim();
      if (!apiKey) continue;
      const cfg = Object.assign({}, p0, { apiKey });
      if (i > 0) {
        if (opts.onReset) opts.onReset();
        log("⚠️ 主通道不可用，自动切换 " + cfg.name);
      }
      const attempts = ((i === order.length - 1) && !!cfg.proxy) ? [false, true] : [false];
      for (const direct of attempts) {
        const r = await callLLMOnce(cfg, messages, Object.assign({}, opts, { direct }));
        if (!r.error) return r;
        lastErr = r;
        log("✗ " + cfg.name + (direct ? "（直连兜底）" : "") + " 失败：" + errMsg(r));
      }
    }
    return lastErr || { error: "nokey" };
  }

  function errMsg(e) {
    if (e.status === 401) return "API Key 无效或已失效（401）：请检查本页对应 Key 框是否填入完整正确的 Key（前后无空格、格式完整）" + (e.msg ? "｜" + e.msg.slice(0, 120) : "");
    if (e.status === 403) return "无权限或账户余额不足（403）：glm-4.6v / Qwen3.5-35B-A3B 为付费模型，需账户有余额并已开通" + (e.msg ? "｜" + e.msg.slice(0, 120) : "");
    const hints = {
      nokey: "未配置可用的 API Key",
      net: "网络/代理失败（" + (e.msg || "Key 无效、代理不可达或浏览器拦截") + "）",
      http: "接口返回 " + (e.status || "") + "：" + (e.msg || ""),
      json: "模型未按约定输出 JSON（" + (e.msg || "") + "）",
      empty: "模型返回为空（可能服务繁忙/被限流）",
      timeout: "请求超时，已自动尝试备用通道"
    };
    return (hints[e.error] || e.msg || e.error || "未知错误") + (e.content ? "\n（原文前 300 字：\n" + e.content.slice(0, 300) + "）" : "");
  }

  // ——— 主流程：识别（SSE 流式 + 自动降级） ———
  $("runBtn").addEventListener("click", async () => {
    const presetKey = $("preset").value;
    const cfg = MODEL_PRESETS[presetKey];
    const apiKey = $(cfg.key + "ApiKey").value.trim();
    const files = [...$("fileInput").files];
    const streamBox = $("streamBox");

    if (!apiKey) { setStatus("请填写对应模型的 API Key"); return; }
    if (!files.length) { setStatus("请先上传送货单图片"); return; }

    logEl.textContent = "";
    lastWorkbook = null;
    $("downloadBtn").disabled = true;
    $("tableWrap").innerHTML = "";
    if (streamBox) { streamBox.style.display = "block"; streamBox.textContent = "等待模型输出…"; }
    $("runBtn").disabled = true;
    setStatus("压缩图片中…");

    // 执行总耗时计时（点击开始 → 完成停表）
    const elapsedEl = $("elapsed");
    const t0 = Date.now();
    const tick = () => { elapsedEl.textContent = "⏱ " + ((Date.now() - t0) / 1000).toFixed(1) + "s"; };
    elapsedEl.textContent = "⏱ 计时中…";
    const timer = setInterval(tick, 200);

    let images;
    try {
      // 多图时自动降低压缩质量与分辨率，避免免费层限流
      const maxDim = files.length > 2 ? 1200 : 2000;
      const quality = files.length > 2 ? 0.7 : 0.85;
      images = await Promise.all(files.map(f => compressImage(f, maxDim, quality)));
      log(`已压缩 ${images.length} 张图片（maxDim=${maxDim} quality=${quality}）`);
    } catch (err) {
      setStatus("图片处理失败：" + err.message); $("runBtn").disabled = false;
      clearInterval(timer); tick();
      return;
    }

    // 图像部分：统一 {url: ...} 格式送图（base64 内联，Agnes 实测兼容）
    const imgParts = images.map(b64 => ({ type: "image_url", image_url: { url: b64 } }));

    // 消息构造：标准 system（提示词）+ user（指令 + 图片）
    const messages = [
      { role: "system", content: SYSTEM_PROMPT },
      { role: "user", content: [{ type: "text", text: "请识别以下送货单图片，严格按系统提示词输出 JSON 数组。" }, ...imgParts] }
    ];

    setStatus("调用模型中（SSE 流式）…");
    log("channel=" + cfg.name + "  model=" + cfg.model + (cfg.proxy ? "（经 Cloudflare 代理）" : "（直连）"));
    try {
      const r = await callLLM(messages, {
        onToken: (t, full) => {
          if (streamBox) { streamBox.textContent = full; streamBox.scrollTop = streamBox.scrollHeight; }
        },
        onReset: () => { if (streamBox) streamBox.innerHTML = ""; }
      });
      if (r.error) throw new Error(errMsg(r));
      const rows = normalizeRows(r.content);
      if (!rows) throw new Error("返回不是数组（模型输出格式异常，请重试或换通道）");

      // 规整字段
      const norm = rows.map(r2 => {
        const o = {};
        EXCEL_HEADERS.forEach(h => { o[h] = (r2[h] === undefined ? "" : r2[h]); });
        return o;
      });
      log(`识别成功，共 ${norm.length} 行`);
      if (streamBox) streamBox.style.display = "none";
      renderTable(norm);
      lastWorkbook = buildWorkbook(norm);
      $("downloadBtn").disabled = false;
      setStatus(`完成：识别 ${norm.length} 行，可下载`);
    } catch (err) {
      log("错误：" + err.message);
      setStatus("失败：" + err.message);
    } finally {
      clearInterval(timer);
      tick();
      $("runBtn").disabled = false;
    }
  });

  function renderTable(rows) {
    let html = "<table><thead><tr>" +
      EXCEL_HEADERS.map(h => `<th>${h}</th>`).join("") +
      "</tr></thead><tbody>";
    rows.forEach(r => {
      html += "<tr>" + EXCEL_HEADERS.map(h => `<td>${r[h]}</td>`).join("") + "</tr>";
    });
    html += "</tbody></table>";
    $("tableWrap").innerHTML = html;
  }

  $("downloadBtn").addEventListener("click", () => {
    if (lastWorkbook) {
      const name = "送货单_" + new Date().toISOString().slice(0, 10) + ".xlsx";
      downloadWorkbook(lastWorkbook, name);
    }
  });
})();
