// 主流程：上传 → 压缩图片 → 调视觉大模型 → 解析 JSON → 预览 → 生成 Excel
// 所有模型均浏览器直连（各厂商原生支持 CORS，已实测预检返回 Access-Control-Allow-Origin），无需任何代理。
(function () {
  const $ = id => document.getElementById(id);
  const status = $("status");
  const logEl = $("log");
  let lastWorkbook = null;

  // 模型预设（baseUrl 已含完整 /chat/completions 端点，调用时直接 fetch(baseUrl)）
  //   key   : 对应 API Key 输入框 id 前缀（zhipu → #zhipuApiKey）
  //   vendor: 'sensenova' 时走商汤专属格式（image_url 为字符串；回复在 data.choices[0].message）
  const MODEL_PRESETS = {
    zhipu: {
      name: "智谱 GLM-4.6V-Flash（视觉·直连✅）",
      baseUrl: "https://open.bigmodel.cn/api/paas/v4/chat/completions",
      model: "glm-4.6v-flash", key: "zhipu"
    },
    sensenova: {
      name: "商汤 SenseNova U1（视觉·直连✅）",
      baseUrl: "https://api.sensenova.cn/v1/llm/chat-completions",
      model: "SenseNova-U1", key: "sensenova", vendor: "sensenova"
    },
    qwen35: {
      name: "硅基流动 Qwen2.5-VL-72B（视觉·推荐·直连✅）",
      baseUrl: "https://api.siliconflow.cn/v1/chat/completions",
      model: "Qwen/Qwen2.5-VL-72B", key: "siliconflow"
    },
    sfocr: {
      name: "硅基流动 DeepSeek-OCR（OCR专用·免费·直连✅）",
      baseUrl: "https://api.siliconflow.cn/v1/chat/completions",
      model: "deepseek-ai/DeepSeek-OCR", key: "siliconflow"
    },
  };

  function log(msg) { logEl.textContent += msg + "\n"; logEl.scrollTop = logEl.scrollHeight; }
  function setStatus(msg) { status.textContent = msg; }

  // 预设下拉变更即保存
  $("preset").addEventListener("change", () => localStorage.setItem("do_preset", $("preset").value));

  // 恢复预设（默认智谱 GLM-4.6V-Flash）
  const savedPreset = localStorage.getItem("do_preset");
  $("preset").value = (savedPreset && MODEL_PRESETS[savedPreset]) ? savedPreset : "zhipu";

  // 恢复 API Key（持久化）：三个框 zhipu / sensenova / siliconflow
  ["zhipu", "sensenova", "siliconflow"].forEach(k => {
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
      images = await Promise.all(files.map(f => compressImage(f)));
      log(`已压缩 ${images.length} 张图片`);
    } catch (err) {
      setStatus("图片处理失败：" + err.message); $("runBtn").disabled = false; return;
    }

    // 图像部分：商汤 SenseNova 的 image_url 是「字符串」；其余为 { url: ... }
    const imgParts = images.map(b64 => cfg.vendor === "sensenova"
      ? { type: "image_url", image_url: b64 }
      : { type: "image_url", image_url: { url: b64 } });

    // 消息构造：标准 system（提示词）+ user（指令 + 图片）
    const messages = [
      { role: "system", content: SYSTEM_PROMPT },
      { role: "user", content: [{ type: "text", text: "请识别以下送货单图片，严格按系统提示词输出 JSON 数组。" }, ...imgParts] }
    ];

    setStatus("调用模型中…");
    log("POST " + baseUrl + "  model=" + model);
    try {
      const body = { model, messages, temperature: 0 };
      if (cfg.vendor === "sensenova") body.max_new_tokens = 2048; // 商汤用此参数控制长度
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

      // 解析回复文本：商汤在 data.choices[0].message（字符串）；其余在 choices[0].message.content
      let raw;
      if (cfg.vendor === "sensenova") {
        if (data.error) throw new Error("商汤错误：" + (data.error.message || JSON.stringify(data.error)));
        if (data.status && data.status.code !== 0) throw new Error("商汤状态：" + (data.status.message || data.status.code));
        raw = data?.data?.choices?.[0]?.message || "";
      } else {
        raw = data?.choices?.[0]?.message?.content || "";
      }
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
