/* Dùng chung cho các trang công cụ kế toán: hiện file mẫu gốc (chưa xử lý) để người dùng so với kết quả, kèm nút tải về */
(function () {
  "use strict";
  const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c]));
  const hien = o => {
    if (!o) return "";
    if (o.t === "n") return o.z && /d|m|y/i.test(o.z) && !/#/.test(o.z) ? (o.w || String(o.v)) : (Math.round(o.v * 100) / 100).toLocaleString("vi-VN");
    if (o.t === "d" || o.v instanceof Date) { const d = o.v; return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`; }
    return String(o.v ?? "");
  };
  function bang(ws, max) {
    if (!ws || !ws["!ref"]) return `<p class="note">Sheet trống.</p>`;
    const rg = XLSX.utils.decode_range(ws["!ref"]), het = Math.min(rg.e.r, rg.s.r + max - 1);
    let h = `<table><thead><tr><th></th>`;
    for (let c = rg.s.c; c <= rg.e.c; c++) h += `<th style="text-align:center">${XLSX.utils.encode_col(c)}</th>`;
    h += "</tr></thead><tbody>";
    for (let r = rg.s.r; r <= het; r++) {
      h += `<tr><td class="muted" style="text-align:right">${r + 1}</td>`;
      for (let c = rg.s.c; c <= rg.e.c; c++) {
        const o = ws[XLSX.utils.encode_cell({r, c})];
        h += `<td class="${o && o.t === "n" ? "n" : ""}">${esc(hien(o))}</td>`;
      }
      h += "</tr>";
    }
    return h + `</tbody></table>${het < rg.e.r ? `<p class="note" style="padding:8px">Đang hiện ${max}/${rg.e.r - rg.s.r + 1} dòng. Tải file mẫu về để xem đủ.</p>` : ""}`;
  }
  async function xemMau(el, files, ghiChu) {
    const ds = await Promise.all(files.map(async f => ({f, buf: new Uint8Array(await f.arrayBuffer())})));
    let chon = 0;
    const ve = () => {
      const {f, buf} = ds[chon], wb = XLSX.read(buf, {type: "array", cellDates: true});
      el.innerHTML = `<details open><summary><b>📄 File mẫu gốc, chưa xử lý</b> <span class="muted">· ${esc(ghiChu || "so với kết quả bên dưới để thấy công cụ đã làm gì")}</span></summary>
        <div style="display:grid;gap:10px;margin-top:10px">
          <div class="row sb"><div class="loc">${ds.map((x, i) => `<button type="button" data-g="${i}" class="${i === chon ? "" : "ghost"}">${esc(x.f.name)}</button>`).join("")}</div>
            <button type="button" class="ghost sm" data-tai>⬇ Tải ${ds.length > 1 ? "file đang xem" : "file mẫu"} về máy</button></div>
          ${wb.SheetNames.map(n => `<div class="grid" style="max-height:360px">${bang(wb.Sheets[n], 60)}</div>`).join("")}
        </div></details>`;
      el.querySelectorAll("[data-g]").forEach(b => b.onclick = () => { chon = +b.dataset.g; ve(); });
      el.querySelector("[data-tai]").onclick = () => {
        const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([buf])); a.download = f.name; a.click();
        setTimeout(() => URL.revokeObjectURL(a.href), 2000);
      };
    };
    el.hidden = false; ve();
  }
  window.KT = {xemMau};
})();
