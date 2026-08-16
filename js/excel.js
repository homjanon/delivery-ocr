// Excel 生成：固定表头（含台头公司名称）+ 数据 + 合计
const EXCEL_HEADERS = [
  "单据编号", "日期", "台头公司名称", "名称及规格", "规格型号/颜色",
  "数量", "单位", "单价", "金额", "备注", "客户/收货单位"
];

function buildWorkbook(rows) {
  const XLSX = window.XLSX;
  const aoa = [];
  aoa.push(EXCEL_HEADERS.slice());  // 第 1 行：表头

  // 2026-08-05（兼容 2026/8/5）→ Excel 日期序列号；非法/非标准返回 null
  function dateToSerial(s) {
    const m = String(s).trim().match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
    if (!m) return null;
    const y = +m[1], mo = +m[2], d = +m[3];
    const dt = new Date(Date.UTC(y, mo - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return null;
    return Math.floor((dt.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
  }

  rows.forEach(r => {
    aoa.push(EXCEL_HEADERS.map(h => {
      let v = r[h];
      if (h === "数量" || h === "单价" || h === "金额") {
        v = (v === undefined || v === null || v === "") ? 0 : Number(v);
        if (!isFinite(v)) v = 0;
      } else if (h === "日期" && typeof v === "string" && v.trim()) {
        const ser = dateToSerial(v);
        if (ser !== null) v = ser;   // 转为真实日期序列号，下方统一设 yyyy-mm-dd 格式
      }
      return v === undefined ? "" : v;
    }));
  });

  // 合计行（金额求和）
  const amtIdx = EXCEL_HEADERS.indexOf("金额");
  const total = rows.reduce((s, r) => s + (Number(r["金额"]) || 0), 0);
  const totalRow = EXCEL_HEADERS.map((h, i) =>
    i === 0 ? "合计" : (i === amtIdx ? Math.round(total * 100) / 100 : ""));
  aoa.push(totalRow);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  // 日期列固化为 yyyy-mm-dd：真实日期值，微信 / QQ / Excel / WPS 打开显示一致
  const dateIdx = EXCEL_HEADERS.indexOf("日期");
  for (let r = 1; r < aoa.length; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: dateIdx })];
    if (cell && cell.t === "n") cell.z = "yyyy-mm-dd";
  }
  ws["!cols"] = EXCEL_HEADERS.map(() => ({ wch: 16 }));
  // 表头加粗（第 1 行，索引 0）
  for (let c = 0; c < EXCEL_HEADERS.length; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
    if (cell) cell.s = { font: { bold: true }, fill: { fgColor: "FFEAF1F5" }, alignment: { horizontal: "center" } };
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "送货单");
  return wb;
}

function downloadWorkbook(wb, filename) {
  window.XLSX.writeFile(wb, filename);
}
