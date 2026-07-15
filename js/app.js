// 主流程：上传 → 压缩图片 → 调视觉大模型 → 解析 JSON → 预览 → 生成 Excel
(function () {
  const $ = id => document.getElementById(id);
  const status = $("status");
  const logEl = $("log");
  let lastWorkbook = null;
  let _baseUrl = "", _model = "";

  // 端点说明：
  //  · 硅基流动 api.siliconflow.cn 原生支持 CORS（预检返回 Access-Control-Allow-Origin: *），
  //    浏览器可直连，无需任何代理 → SF_DIRECT。
  //  · NVIDIA integrate.api.nvidia.com 的预检不带 ACAO，浏览器无法直连，必须走代理。
  //    默认走 Cloudflare Worker，但 Cloudflare 在国内常被墙；如需国内可用，把 WORKER 改成
  //    你自建的国内可达代理（阿里云函数计算 / 腾讯云函数等部署 proxy-worker.js 的地址）。
  const WORKER = "https://cors-proxy.homjanon.workers.dev/?url=";
  const nvidiaEp = encodeURIComponent("https://integrate.api.nvidia.com/v1/chat/completions");
  const SF_DIRECT = "https://api.siliconflow.cn/v1/chat/completions";

  // 模型预设（baseUrl 已含完整 /chat/completions 端点，调用时直接 fetch(baseUrl)）
  const MODEL_PRESETS = {
    sfds: { name: "硅基流动 Qwen2.5-VL-72B（视觉·直连✅）",
      baseUrl: SF_DIRECT, model: "Qwen/Qwen2.5-VL-72B-Instruct" },
    sfocr: { name: "硅基流动 DeepSeek-OCR（OCR专用·免费·直连✅）",
      baseUrl: SF_DIRECT, model: "deepseek-ai/DeepSeek-OCR", grounding: true },
    glm52: { name: "NVIDIA GLM-5.2（需国内可达代理）",
      baseUrl: WORKER + nvidiaEp, model: "z-ai/glm-5.2" },
    qw397: { name: "NVIDIA Qwen 3.5-397B VLM（需国内可达代理）",
      baseUrl: WORKER + nvidiaEp, model: "qwen/qwen3.5-397b-a17b" },
  };

  function log(msg) { logEl.textContent += msg + "\n"; logEl.scrollTop = logEl.scrollHeight; }
  function setStatus(msg) { status.textContent = msg; }

  // 应用预设：记录 baseUrl / model
  function applyPreset(presetKey, save = true) {
    const preset = MODEL_PRESETS[presetKey];
    if (preset) {
      _baseUrl = preset.baseUrl;
      _model = preset.model;
    }
    if (save) {
      localStorage.setItem("do_preset", presetKey);
      localStorage.setItem("do_baseUrl", _baseUrl);
      localStorage.setItem("do_model", _model);
    }
  }

  // 预设下拉变更
  $("preset").addEventListener("change", () => applyPreset($("preset").value));

  // 恢复设置
  const savedPreset = localStorage.getItem("do_preset");
  const savedBaseUrl = localStorage.getItem("do_baseUrl");
  const savedModel = localStorage.getItem("do_model");
  if (savedPreset && MODEL_PRESETS[savedPreset]) {
    $("preset").value = savedPreset;
    applyPreset(savedPreset, false);
  } else {
    $("preset").value = "sfds";
    applyPreset("sfds", false);
  }
  if (savedBaseUrl && savedPreset === "custom") { _baseUrl = savedBaseUrl; }
  if (savedModel && savedPreset === "custom") { _model = savedModel; }

  // 恢复 API Key（持久化）
  $("nvidiaApiKey").value = localStorage.getItem("do_nvidiaKey") || "";
  $("siliconflowApiKey").value = localStorage.getItem("do_siliconflowKey") || "";
  // 输入即保存
  $("nvidiaApiKey").addEventListener("input", () =>
    localStorage.setItem("do_nvidiaKey", $("nvidiaApiKey").value));
  $("siliconflowApiKey").addEventListener("input", () =>
    localStorage.setItem("do_siliconflowKey", $("siliconflowApiKey").value));

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
    let t = text.trim();
    const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fence) t = fence[1].trim();
    const s = t.indexOf("["), e = t.lastIndexOf("]");
    if (s !== -1 && e !== -1 && e > s) t = t.slice(s, e + 1);
    return JSON.parse(t);
  }

  // ——— 主流程：识别 ———
  $("runBtn").addEventListener("click", async () => {
    const baseUrl = _baseUrl.replace(/\/$/, "");
    // 按预设自动选取对应的 API Key 框
    const preset = $("preset").value;
    const apiKey = (preset === "sfds" || preset === "sfocr") ? $("siliconflowApiKey").value.trim() : $("nvidiaApiKey").value.trim();
    const model = _model.trim();
    const files = [...$("fileInput").files];

    if (!apiKey) { setStatus("请填写 API Key"); return; }
    if (!files.length) { setStatus("请先上传送货单图片"); return; }

    // 保存设置
    localStorage.setItem("do_baseUrl", baseUrl);
    localStorage.setItem("do_model", model);
    logEl.textContent = "";
    lastWorkbook = null;
    $("downloadBtn").disabled = true;
    $("tableWrap").innerHTML = "";
    $("runBtn").disabled = true;
    setStatus("压缩图片中…");

    let images;
    try {
      images = await Promise.all(files.map(f => compressImage(f)));
      log(`已压缩 ${images.length} 张图片`);
    } catch (err) {
      setStatus("图片处理失败：" + err.message); $("runBtn").disabled = false; return;
    }

    // 构造消息：DeepSeek-OCR 等 OCR 专用模型指令需放在 user 消息（不单独用 system 角色）；
    // 注意：不放 <|grounding|> 前缀，避免返回 <|ref|> 标签污染 JSON。其余模型用 system + user 标准结构。
    const imgParts = images.map(b64 => ({ type: "image_url", image_url: { url: b64 } }));
    let messages;
    if (MODEL_PRESETS[preset] && MODEL_PRESETS[preset].grounding) {
      messages = [{
        role: "user",
        content: [
          { type: "text", text: SYSTEM_PROMPT + "\n请按上述提示词识别以下送货单图片，严格输出 JSON 数组。" },
          ...imgParts
        ]
      }];
    } else {
      messages = [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: [{ type: "text", text: "请识别以下送货单图片，严格按系统提示词输出 JSON 数组。" }, ...imgParts] }
      ];
    }

    setStatus("调用模型中…");
    log("POST " + baseUrl + "  model=" + model);
    try {
      const resp = await fetch(baseUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json", "Authorization": "Bearer " + apiKey },
        body: JSON.stringify({
          model,
          messages,
          temperature: 0
        })
      });
      if (!resp.ok) {
        const errText = await resp.text();
        throw new Error("HTTP " + resp.status + " " + errText.slice(0, 300));
      }
      const data = await resp.json();
      const raw = data.choices?.[0]?.message?.content || "";
      log("模型返回（前 200 字）：\n" + raw.slice(0, 200));

      let rows;
      try { rows = extractJSON(raw); }
      catch (e) { throw new Error("JSON 解析失败，模型未按约定输出：" + e.message); }
      if (!Array.isArray(rows)) throw new Error("返回不是数组");

      // 规整字段
      rows = rows.map(r => {
        const o = {};
        EXCEL_HEADERS.forEach(h => { o[h] = (r[h] === undefined ? "" : r[h]); });
        return o;
      });
      log(`识别成功，共 ${rows.length} 行`);
      renderTable(rows);
      lastWorkbook = buildWorkbook(rows);
      $("downloadBtn").disabled = false;
      setStatus(`完成：识别 ${rows.length} 行，可下载`);
    } catch (err) {
      log("错误：" + err.message);
      setStatus("失败：" + err.message);
    } finally {
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
