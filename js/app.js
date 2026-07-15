// 主流程：上传 → 压缩图片 → 调视觉大模型 → 解析 JSON → 预览 → 生成 Excel
// 所有模型均浏览器直连（各厂商原生支持 CORS，已实测预检返回 Access-Control-Allow-Origin），无需任何代理。
(function () {
  const $ = id => document.getElementById(id);
  const status = $("status");
  const logEl = $("log");
  let lastWorkbook = null;

  // 模型预设（baseUrl 已含完整 /chat/completions 端点，调用时直接 fetch(baseUrl)）
  const MODEL_PRESETS = {
    zhipu: {
      name: "智谱 GLM-4.6V-Flash（视觉·直连✅）",
      baseUrl: "https://open.bigmodel.cn/api/paas/v4/chat/completions",
      model: "glm-4.6v-flash", key: "zhipu"
    },
    qwen35: {
      name: "硅基流动 Qwen3.5-397B-A17B（视觉·直连✅）",
      baseUrl: "https://api.siliconflow.cn/v1/chat/completions",
      model: "Qwen/Qwen3.5-397B-A17B", key: "siliconflow"
    },
  };

  function log(msg) { logEl.textContent += msg + "\n"; logEl.scrollTop = logEl.scrollHeight; }
  function setStatus(msg) { status.textContent = msg; }

  // 预设下拉变更即保存
  $("preset").addEventListener("change", () => localStorage.setItem("do_preset", $("preset").value));

  // 恢复预设（默认智谱 GLM-4.6V-Flash）
  const savedPreset = localStorage.getItem("do_preset");
  $("preset").value = (savedPreset && MODEL_PRESETS[savedPreset]) ? savedPreset : "zhipu";

  // 恢复 API Key（持久化）：两个框 zhipu / siliconflow
  ["zhipu", "siliconflow"].forEach(k => {
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
    let t = text.trim();
    const fence = t.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fence) t = fence[1].trim();
    const s = t.indexOf("["), e = t.lastIndexOf("]");
    if (s !== -1 && e !== -1 && e > s) t = t.slice(s, e + 1);
    return JSON.parse(t);
  }

  // ——— 主流程：识别 ———
  $("runBtn").addEventListener("click", async () => {
    const presetKey = $("preset").value;
    const cfg = MODEL_PRESETS[presetKey];
    const baseUrl = cfg.baseUrl.replace(/\/$/, "");
    const model = cfg.model;
    const apiKey = $(cfg.key + "ApiKey").value.trim();
    const files = [...$("fileInput").files];

    if (!apiKey) { setStatus("请填写对应模型的 API Key"); return; }
    if (!files.length) { setStatus("请先上传送货单图片"); return; }

    logEl.textContent = "";
    lastWorkbook = null;
    $("downloadBtn").disabled = true;
    $("tableWrap").innerHTML = "";
    $("runBtn").disabled = true;
    setStatus("压缩图片中…");

    let images;
    try {
      // 多图时自动降低压缩质量与分辨率，避免免费层限流
      const maxDim = files.length > 2 ? 1200 : 2000;
      const quality = files.length > 2 ? 0.7 : 0.85;
      images = await Promise.all(files.map(f => compressImage(f, maxDim, quality)));
      log(`已压缩 ${images.length} 张图片（maxDim=${maxDim} quality=${quality}）`);
    } catch (err) {
      setStatus("图片处理失败：" + err.message); $("runBtn").disabled = false; return;
    }

    // 图像部分：所有模型用标准 {url: ...} 格式
    const imgParts = images.map(b64 => ({ type: "image_url", image_url: { url: b64 } }));

    // 消息构造：标准 system（提示词）+ user（指令 + 图片）
    const messages = [
      { role: "system", content: SYSTEM_PROMPT },
      { role: "user", content: [{ type: "text", text: "请识别以下送货单图片，严格按系统提示词输出 JSON 数组。" }, ...imgParts] }
    ];

    setStatus("调用模型中…");
    log("POST " + baseUrl + "  model=" + model);
    try {
      const body = { model, messages, temperature: 0, max_tokens: 2048 };
      const resp = await fetch(baseUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json", "Authorization": "Bearer " + apiKey },
        body: JSON.stringify(body)
      });
      if (!resp.ok) {
        const errText = await resp.text();
        throw new Error("HTTP " + resp.status + " " + errText.slice(0, 300));
      }
      const data = await resp.json();

      // 标准 OpenAI 格式解析
      const raw = data?.choices?.[0]?.message?.content || "";
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
