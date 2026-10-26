/* Phần dùng chung: chuyển thẻ, đọc file, nhận diện dòng tiêu đề, khớp tên cột, vẽ bảng, tải file */
(function () {
  "use strict";
  const $ = s => document.querySelector(s);
  const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c]));
  const boDau = s => window.LamSach.boDau(s);
  const rong = v => v === null || v === undefined || String(v).trim() === "";

  // ---------- chuyển thẻ (nhớ thẻ đang mở qua #) ----------
  function moThe(k) {
    document.querySelectorAll("#the button").forEach(b => b.classList.toggle("on", b.dataset.the === k));
    document.querySelectorAll(".tab").forEach(t => t.hidden = t.dataset.tab !== k);
    try { history.replaceState(null, "", "#" + k); } catch (e) {}
  }
  document.querySelectorAll("#the button").forEach(b => b.addEventListener("click", () => moThe(b.dataset.the)));
  const dau = location.hash.slice(1);
  if (document.querySelector(`.tab[data-tab="${dau}"]`)) moThe(dau);

  // ---------- đọc file ----------
  const docBuf = f => new Promise((ok, loi) => { const r = new FileReader(); r.onload = e => ok(new Uint8Array(e.target.result)); r.onerror = loi; r.readAsArrayBuffer(f); });
  const docWb = buf => XLSX.read(buf, {type: "array", cellDates: true});
  // tìm dòng tiêu đề trong 10 dòng đầu (bỏ qua dòng tiêu đề báo cáo kiểu "BÁO CÁO DOANH SỐ...")
  function bangSheet(ws) {
    const all = XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: "", blankrows: true});
    while (all.length && all[all.length - 1].every(rong)) all.pop();
    if (!all.length) return {hd: [], rows: [], dongTieuDe: 0};
    let best = 0, diem = -1;
    for (let i = 0; i < Math.min(10, all.length); i++) {
      const r = all[i], chu = r.filter(v => typeof v === "string" && v.trim() && isNaN(+v)).length, day = r.filter(v => !rong(v)).length;
      const d = chu * 2 + day - (day === 1 ? 5 : 0);
      if (d > diem) { diem = d; best = i; }
    }
    const n = Math.max(...all.map(r => r.length));
    const hd = Array.from({length: n}, (_, j) => String(all[best][j] ?? "").trim() || `Cột ${j + 1}`);
    const rows = all.slice(best + 1).map(r => Array.from({length: n}, (_, j) => r[j] ?? ""));
    return {hd, rows, dongTieuDe: best + 1};
  }
  // khoá so khớp tên cột: bỏ dấu, bỏ phần trong ngoặc (đơn vị), bỏ ký tự lạ, gom từ đồng nghĩa
  const DONG_NGHIA = [
    ["sdt", ["sdt", "sodienthoai", "dienthoai", "sodt", "dt", "phone", "mobile", "didong", "sdtlienhe"]],
    ["hoten", ["hoten", "hovaten", "fullname", "tennhanvien", "tennv"]],
    ["email", ["email", "mail", "thudientu"]],
    ["manv", ["manv", "manhanvien", "msnv", "masonhanvien"]],
    ["makh", ["makh", "makhachhang"]],
    ["sohd", ["sohd", "sohoadon", "mahd", "mahoadon", "sohopdong"]],
    ["ngaysinh", ["ngaysinh", "ns", "dob"]],
    ["stt", ["stt", "tt", "sothutu"]],
  ];
  function khoaCot(h) {
    const k = boDau(String(h || "")).replace(/\(.*?\)/g, "").replace(/[^a-z0-9]/g, "");
    for (const [g, ds] of DONG_NGHIA) if (ds.includes(k)) return g;
    return k;
  }

  // ---------- kéo thả ----------
  function ganDrop(drop, input, cb) {
    input.addEventListener("change", e => { if (e.target.files.length) cb([...e.target.files]); e.target.value = ""; });
    ["dragenter", "dragover"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add("over"); }));
    ["dragleave", "drop"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove("over"); }));
    drop.addEventListener("drop", e => e.dataTransfer.files.length && cb([...e.dataTransfer.files]));
  }

  // ---------- hiển thị giá trị ----------
  function hien(v) {
    if (v instanceof Date && !isNaN(v)) return `${String(v.getDate()).padStart(2, "0")}/${String(v.getMonth() + 1).padStart(2, "0")}/${v.getFullYear()}`;
    if (typeof v === "number") return (Math.round(v * 100) / 100).toLocaleString("vi-VN", {maximumFractionDigits: 2});
    return String(v ?? "");
  }
  // số từ ô: số thật, hoặc chữ kiểu "1.250.000", "15tr"
  function so(v) {
    if (typeof v === "number") return v;
    if (rong(v)) return null;
    const r = window.LamSach.tien(v);
    return typeof r.v === "number" && !r.loi ? r.v : null;
  }
  // bảng HTML: cols = [{ten}], rows = [{o:[...], cls:[...], trCls, rn}]
  function veBang(el, cols, rows, max = 500) {
    if (!rows.length) { el.innerHTML = `<p class="note" style="padding:14px">Không có dòng nào ở mục này.</p>`; return; }
    let h = `<table><thead><tr><th class="rn">#</th>${cols.map(c => `<th>${esc(c)}</th>`).join("")}</tr></thead><tbody>`;
    rows.slice(0, max).forEach((r, i) => {
      h += `<tr class="${r.trCls || ""}"><td class="rn">${r.rn ?? i + 1}</td>${r.o.map((v, j) => {
        const c = (r.cls && r.cls[j]) || (typeof v === "number" ? "n" : "");
        return `<td class="${c}"${r.tip && r.tip[j] ? ` title="${esc(r.tip[j])}"` : ""}>${r.html && r.html[j] ? r.html[j] : esc(hien(v))}</td>`; }).join("")}</tr>`;
    });
    h += "</tbody></table>";
    if (rows.length > max) h += `<p class="note" style="padding:8px">Đang hiện ${max}/${rows.length} dòng. File tải về có đủ tất cả.</p>`;
    el.innerHTML = h;
  }
  // sheet từ mảng 2 chiều, giữ kiểu: số có dấu phân cách, ngày dd/mm/yyyy
  function sheetTu(aoa, rong) {
    const ws = XLSX.utils.aoa_to_sheet(aoa, {cellDates: true, dateNF: "dd/mm/yyyy"});
    const rg = XLSX.utils.decode_range(ws["!ref"] || "A1");
    for (let r = 1; r <= rg.e.r; r++) for (let c = 0; c <= rg.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({r, c})];
      if (cell && cell.t === "n" && Math.abs(cell.v) >= 1000 && !cell.z) cell.z = Number.isInteger(cell.v) ? "#,##0" : "#,##0.00";
      if (cell && cell.t === "d") cell.z = "dd/mm/yyyy";
    }
    ws["!cols"] = (aoa[0] || []).map((h, j) => ({wch: rong ? rong[j] : Math.min(40, Math.max(8, String(h).length + 2, ...aoa.slice(1, 200).map(r => hien(r[j]).length + 2)))}));
    return ws;
  }
  function taiWb(sheets, ten) {
    const wb = XLSX.utils.book_new();
    sheets.forEach(([n, ws]) => XLSX.utils.book_append_sheet(wb, ws, n.slice(0, 31)));
    XLSX.writeFile(wb, ten);
  }
  const fileMau = (sheets) => { const wb = XLSX.utils.book_new(); sheets.forEach(([n, aoa]) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), n)); return XLSX.write(wb, {bookType: "xlsx", type: "array"}); };

  window.TL = {$, esc, boDau, rong, moThe, docBuf, docWb, bangSheet, khoaCot, ganDrop, hien, so, veBang, sheetTu, taiWb, fileMau};
})();
