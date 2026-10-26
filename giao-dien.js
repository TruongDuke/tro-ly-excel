/* Giao diện: đọc file, cấu hình cột, xem trước, tải file đã làm sạch */
(function () {
  "use strict";
  const $ = s => document.querySelector(s);
  const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}[c]));
  const L = window.LamSach;
  let wb = null, tenFile = "", rows = null, loai = [], kq = null, loc = "tat";

  // ---------- file mẫu lộn xộn ----------
  function fileMau() {
    const nfd = s => s.normalize("NFD");
    const d = [
      ["STT", "Họ và tên", "SĐT", "Ngày sinh", "Lương", "Email", "Phòng ban"],
      [1, "  nguyễn  văn   AN ", "912345678", "15/3/1995", "12.500.000", " An.Nguyen@Gmail.COM ", "Kinh doanh"],
      [2, "TRẦN THỊ thu hà", "0987 654 321", "1996-07-02", "15tr", "thuha@congty.vn", "Kế toán"],
      [3, nfd("Lê Hoàng Phúc"), "+84 903.111.222", "02.09.94", "9,8tr", "phuc.le@congty.vn", "Kinh doanh"],
      [4, "phạm   minh châu.", "84976543210", "12-25-1993", "11 triệu", "chau@congty", "Nhân sự"],
      [5, "Võ Thanh Tùng", "0912 345 67", "31/02/1990", "13.000.000 đ", "tung.vo@congty.vn", "Kỹ thuật"],
      [6, "ĐẶNG THỊ MAI", 935222333, 34567, "8500k", "mai.dang@congty.vn", "Kế toán"],
      [7, "bùi quốc  bảo", "0868-777-888", "ngày 5 tháng 6 năm 1998", "10,000,000", "bao.bui@congty.vn", "Kỹ thuật"],
      [8, "  Nguyễn Văn An", "0912345678", "15/03/1995", "12500000", "an.nguyen@gmail.com", "Kinh doanh"],
      [],
      [9, "hoàng thị LAN anh", "(024) 3826 1234", "20/11/1997", "1tr2", "lananh@congty.vn", "Hành chính"],
      [10, "Trương Văn Đức", "0399 888 777", "1999/1/9", "14,5tr", "duc.truong@congty.vn", "Kỹ thuật"],
      [11, "lý  thị  kim   ngân", "0777 123 456", "7/8/2000", "khoảng 9tr", "ngan.ly@congty.vn", "Kinh doanh"],
      [12, nfd("Đỗ Quang Huy"), "0345678901", "3/4/91", "16.200.000", "HUY.DO@CONGTY.VN", "Ban giám đốc"],
      [13, "TRẦN THỊ THU HÀ", "0987654321", "02/07/1996", "15.000.000", "thuha@congty.vn", "Kế toán"],
    ];
    const ws = XLSX.utils.aoa_to_sheet(d); ws["!cols"] = [{wch: 5}, {wch: 24}, {wch: 18}, {wch: 22}, {wch: 14}, {wch: 26}, {wch: 14}];
    const b = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(b, ws, "Danh sách nhân viên");
    return XLSX.write(b, {bookType: "xlsx", type: "array"});
  }

  // ---------- đọc file ----------
  function moFile(buf, ten) {
    try { wb = XLSX.read(buf, {type: "array", cellDates: true, raw: false}); }
    catch (e) { alert("Không đọc được file này: " + e.message); return; }
    tenFile = ten;
    $("#sheet").innerHTML = wb.SheetNames.map((n, i) => `<option value="${i}">${esc(n)}</option>`).join("");
    $("#sheetRow").hidden = wb.SheetNames.length < 2 && false;
    napSheet(0);
  }
  function napSheet(i) {
    const ws = wb.Sheets[wb.SheetNames[i]];
    rows = XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: "", blankrows: true});
    while (rows.length && rows[rows.length - 1].every(v => v === "")) rows.pop();
    if (rows.length < 2) { alert("Sheet này chưa có dữ liệu (cần dòng tiêu đề và ít nhất 1 dòng)."); return; }
    const n = Math.max(...rows.map(r => r.length)); rows = rows.map(r => Array.from({length: n}, (_, j) => r[j] ?? ""));
    loai = rows[0].map((h, j) => {
      const t = L.boDau(String(h));
      if (/^(stt|tt|so thu tu|#)$/.test(t.trim())) return "giu";
      return L.doanLoai(String(h), rows.slice(1).map(r => r[j]));
    });
    $("#fileInfo").textContent = `${tenFile} · ${rows.length - 1} dòng · ${n} cột`;
    $("#sheetRow").hidden = false; $("#cauHinh").hidden = false; $("#ketQua").hidden = false;
    veCot(); chay();
  }

  // ---------- cấu hình cột ----------
  function veCot() {
    const opts = Object.entries(L.LOAI).map(([k, o]) => [k, o.ten]);
    $("#cols").innerHTML = rows[0].map((h, j) => `<div class="col"><b title="${esc(h)}">${esc(h || "(cột " + (j + 1) + ")")}</b>
      <select data-j="${j}">${opts.map(([k, t]) => `<option value="${k}"${k === loai[j] ? " selected" : ""}>${t}</option>`).join("")}</select>
      <span class="muted">Ví dụ: ${esc(String(rows.slice(1).map(r => r[j]).find(v => v !== "") ?? "").slice(0, 28))}</span></div>`).join("");
    $("#cols").querySelectorAll("select").forEach(s => s.addEventListener("change", () => { loai[+s.dataset.j] = s.value; chay(); }));
    const ht = loai.indexOf("hoten"), sd = loai.indexOf("sdt"), em = loai.indexOf("email");
    const goiY = sd >= 0 ? sd : em >= 0 ? em : ht;
    $("#khoaTrung").innerHTML = `<option value="-1">Toàn bộ dòng giống hệt</option>` + rows[0].map((h, j) => `<option value="${j}"${j === goiY ? " selected" : ""}>${esc(h || "(cột " + (j + 1) + ")")}</option>`).join("");
  }

  // ---------- chạy & hiển thị ----------
  function chay() {
    kq = L.lamSach(rows, loai, {nfc: $("#oNFC").checked, xoaTrong: $("#oTrong").checked, xoaTrung: $("#oTrung").checked, khoa: +$("#khoaTrung").value});
    const conLai = kq.ds.filter(d => !d.xoa).length;
    $("#tong").innerHTML = `<div class="o"><b>${conLai}</b>dòng giữ lại / ${kq.tong}</div><div class="s"><b>${kq.soSua}</b>ô đã sửa</div>
      <div class="l"><b>${kq.soLoi}</b>ô cần kiểm tra</div><div class="x"><b>${kq.soTrung + kq.soTrong}</b>dòng xoá (${kq.soTrung} trùng, ${kq.soTrong} trống)</div>`;
    veBang();
  }
  function cotXuat() {   // danh sách cột khi xuất: chèn Họ / Tên đệm / Tên sau cột họ tên
    const out = [];
    kq.hd.forEach((h, j) => { out.push({j, ten: h}); if (loai[j] === "hoten" && $("#oTach").checked) ["Họ", "Tên đệm", "Tên"].forEach((t, k) => out.push({j, tach: k, ten: t})); });
    return out;
  }
  function veBang() {
    const cs = cotXuat();
    let ds = kq.ds;
    if (loc === "doi") ds = ds.filter(d => d.xoa || d.o.some(o => o.doi || o.loi));
    if (loc === "loi") ds = ds.filter(d => d.o.some(o => o.loi));
    const MAX = 500;
    let h = `<table><thead><tr><th class="rn">Dòng</th>${cs.map(c => `<th>${esc(c.ten)}</th>`).join("")}<th>Ghi chú</th></tr></thead><tbody>`;
    ds.slice(0, MAX).forEach(d => {
      h += `<tr class="${d.xoa ? "xoa" : ""}"><td class="rn">${d.goc}</td>`;
      cs.forEach(c => {
        const o = d.o[c.j];
        if (c.tach !== undefined) { h += `<td>${esc(L.tachTen(o.hien)[c.tach])}</td>`; return; }
        const cls = o.loi ? "loi" : o.doi ? "sua" : "";
        const tip = o.loi || o.ghi || "";
        h += `<td class="${cls}"${tip ? ` title="${esc(tip)}"` : ""}>${esc(o.loi ? o.cu : o.hien)}${o.doi && !o.loi && o.cu ? `<span class="cu">${esc(o.cu)}</span>` : ""}</td>`;
      });
      const ghi = d.xoa || d.o.map(o => o.loi || o.ghi).filter(Boolean).join("; ");
      h += `<td>${esc(ghi)}</td></tr>`;
    });
    h += `</tbody></table>`;
    if (ds.length > MAX) h += `<p class="note" style="padding:8px">Đang hiện ${MAX}/${ds.length} dòng. File tải về có đủ tất cả.</p>`;
    if (!ds.length) h = `<p class="note" style="padding:14px">Không có dòng nào ở mục này.</p>`;
    $("#grid").innerHTML = h;
  }

  // ---------- xuất file ----------
  const serial = o => (Date.UTC(o.y, o.m - 1, o.d) - Date.UTC(1899, 11, 30)) / 864e5;
  function taiFile() {
    const cs = cotXuat(), ws = {};
    const put = (r, c, cell) => { ws[XLSX.utils.encode_cell({r, c})] = cell; };
    cs.forEach((c, k) => put(0, k, {t: "s", v: c.ten}));
    let r = 1;
    kq.ds.filter(d => !d.xoa).forEach(d => {
      cs.forEach((c, k) => {
        const o = d.o[c.j], lo = loai[c.j];
        if (c.tach !== undefined) { const t = L.tachTen(o.hien)[c.tach]; if (t) put(r, k, {t: "s", v: t}); return; }
        if (o.v === "" || o.v === null || o.v === undefined) return;
        if (!o.loi && lo === "ngay" && typeof o.v === "object") put(r, k, {t: "n", v: serial(o.v), z: "dd/mm/yyyy"});
        else if (!o.loi && lo === "tien" && typeof o.v === "number") put(r, k, {t: "n", v: o.v, z: Number.isInteger(o.v) ? "#,##0" : "#,##0.00"});
        else if (typeof o.v === "number" && lo === "giu") put(r, k, {t: "n", v: o.v});
        else put(r, k, {t: "s", v: String(o.loi ? o.cu : o.hien)});
      });
      r++;
    });
    ws["!ref"] = XLSX.utils.encode_range({s: {r: 0, c: 0}, e: {r: Math.max(0, r - 1), c: cs.length - 1}});
    ws["!cols"] = cs.map((c, k) => ({wch: Math.min(40, Math.max(8, c.ten.length + 2, ...kq.ds.slice(0, 200).map(d => String(c.tach !== undefined ? L.tachTen(d.o[c.j].hien)[c.tach] : d.o[c.j].hien).length + 2)))}));
    const nk = XLSX.utils.aoa_to_sheet([["Dòng gốc", "Cột", "Giá trị cũ", "Giá trị mới", "Ghi chú"], ...kq.log]);
    nk["!cols"] = [{wch: 9}, {wch: 18}, {wch: 28}, {wch: 28}, {wch: 50}];
    const b = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(b, ws, "Đã làm sạch");
    XLSX.utils.book_append_sheet(b, nk, "Nhật ký thay đổi");
    XLSX.writeFile(b, tenFile.replace(/\.(xlsx|xls|csv)$/i, "") + "_da-lam-sach.xlsx");
  }

  // ---------- sự kiện ----------
  const docFile = f => { const rd = new FileReader(); rd.onload = e => moFile(new Uint8Array(e.target.result), f.name); rd.readAsArrayBuffer(f); };
  $("#file").addEventListener("change", e => e.target.files[0] && docFile(e.target.files[0]));
  const dz = $("#drop");
  ["dragenter", "dragover"].forEach(t => dz.addEventListener(t, e => { e.preventDefault(); dz.classList.add("over"); }));
  ["dragleave", "drop"].forEach(t => dz.addEventListener(t, e => { e.preventDefault(); dz.classList.remove("over"); }));
  dz.addEventListener("drop", e => e.dataTransfer.files[0] && docFile(e.dataTransfer.files[0]));
  $("#mau").addEventListener("click", () => moFile(fileMau(), "danh-sach-mau.xlsx"));
  $("#sheet").addEventListener("change", e => napSheet(+e.target.value));
  ["#oNFC", "#oTach", "#oTrong", "#oTrung", "#khoaTrung"].forEach(id => $(id).addEventListener("change", () => rows && chay()));
  document.querySelectorAll("[data-loc]").forEach(b => b.addEventListener("click", () => {
    loc = b.dataset.loc; document.querySelectorAll("[data-loc]").forEach(x => x.classList.toggle("ghost", x !== b)); veBang(); }));
  document.querySelector('[data-loc="tat"]').classList.remove("ghost");
  $("#tai").addEventListener("click", taiFile);
  window._tl = {moFile, fileMau, get kq() { return kq; }, taiFile};   // để kiểm thử
})();
