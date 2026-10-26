/* Đối chiếu 2 danh sách theo cột mã: khớp, lệch, chỉ có ở A, chỉ có ở B */
(function () {
  "use strict";
  const {$, esc, boDau, rong, docBuf, docWb, bangSheet, khoaCot, ganDrop, hien, so, veBang, sheetTu, taiWb, fileMau} = window.TL;
  const S = {A: null, B: null};   // {ten, wb, sheet, hd, rows}
  let kq = null, loc = "tat";

  function mau() {
    const A = [["SỔ CHI TIẾT CÔNG NỢ PHẢI THU - THÁNG 10/2026"], [],
      ["Số HĐ", "Khách hàng", "Ngày HĐ", "Số tiền phải thu"],
      ["HD0101", "Công ty TNHH Hoa Sen", "02/10/2026", 48500000], ["HD0102", "Công ty CP Minh Quang", "03/10/2026", 125000000],
      ["HD0103", "Cửa hàng Bình An", "05/10/2026", 7800000], ["HD0104", "Công ty TNHH Đông Á", "08/10/2026", 62000000],
      ["HD0105", "Nhà sách Phương Nam", "10/10/2026", 15400000], ["HD0106", "Công ty CP An Phát", "12/10/2026", 230000000],
      ["HD0107", "Siêu thị Thành Đô", "15/10/2026", 33600000], ["HD0108", "Công ty TNHH Tân Phú", "18/10/2026", 9900000],
      ["HD0109", "Cửa hàng Kim Long", "20/10/2026", 4250000], ["HD0110", "Công ty CP Hưng Thịnh", "25/10/2026", 87300000]];
    const B = [["Ngày GD", "Nội dung chuyển khoản", "Số hoá đơn", "Số tiền (VNĐ)"],
      ["05/10/2026", "HOA SEN TT HD0101", "HD0101", "48.500.000"], ["07/10/2026", "MINH QUANG TT HD 0102 DOT 1", "HD0102", "60.000.000"],
      ["21/10/2026", "MINH QUANG TT HD0102 DOT 2", "HD0102", "65.000.000"], ["09/10/2026", "BINH AN CK", "hd0103", "7.800.000"],
      ["14/10/2026", "DONG A THANH TOAN", "HD0104", "60.000.000"], ["13/10/2026", "PHUONG NAM", "HD 0105", "15.400.000"],
      ["16/10/2026", "AN PHAT TT HD0106", "HD0106", "230.000.000"], ["22/10/2026", "TAN PHU CK", "HD0108", "9.900.000"],
      ["24/10/2026", "CK KHONG RO NOI DUNG", "HD0199", "12.000.000"]];
    return {A: {ten: "So-cong-no-T10.xlsx", buf: fileMau([["Công nợ", A]])}, B: {ten: "Sao-ke-ngan-hang-T10.xlsx", buf: fileMau([["Sao kê", B]])}};
  }

  async function napFile(ben, f) {
    try { S[ben] = {ten: f.name || f.ten, wb: docWb(f.buf || await docBuf(f))}; }
    catch (e) { alert(`Không đọc được file: ${e.message}`); return; }
    const sel = $("#dSheet" + ben);
    sel.innerHTML = S[ben].wb.SheetNames.map((n, i) => `<option value="${i}">Sheet: ${esc(n)}</option>`).join("");
    sel.hidden = S[ben].wb.SheetNames.length < 2;
    napSheet(ben, 0);
  }
  function napSheet(ben, i) {
    const b = bangSheet(S[ben].wb.Sheets[S[ben].wb.SheetNames[i]]);
    Object.assign(S[ben], {sheet: S[ben].wb.SheetNames[i], hd: b.hd, rows: b.rows.filter(r => !r.every(rong))});
    $("#dTen" + ben).textContent = `${S[ben].ten} · ${S[ben].rows.length} dòng`;
    if (S.A && S.B) cauHinh();
  }

  // ---------- cấu hình: cột mã và các cặp cột so sánh ----------
  const laMa = h => /^(sohd|manv|makh|sdt|ma|so|id|code)/.test(khoaCot(h)) || /ma|so hd|so hoa don|so phieu|so ct|chung tu/.test(boDau(h));
  function cauHinh() {
    const ka = S.A.hd.map(khoaCot), kb = S.B.hd.map(khoaCot);
    let ia = ka.findIndex((k, i) => kb.includes(k) && laMa(S.A.hd[i])), ib;
    if (ia < 0) ia = ka.findIndex(k => kb.includes(k));
    if (ia < 0) { ia = Math.max(0, S.A.hd.findIndex(laMa)); ib = Math.max(0, S.B.hd.findIndex(laMa)); } else ib = kb.indexOf(ka[ia]);
    $("#dKhoaA").innerHTML = S.A.hd.map((h, i) => `<option value="${i}"${i === ia ? " selected" : ""}>${esc(h)}</option>`).join("");
    $("#dKhoaB").innerHTML = S.B.hd.map((h, i) => `<option value="${i}"${i === ib ? " selected" : ""}>${esc(h)}</option>`).join("");
    veCot();
    $("#dCauHinh").hidden = false;
    chay();
  }
  function ghepCot(j) {   // cột B tương ứng với cột A thứ j: trùng tên, hoặc tên này chứa tên kia (Số tiền / Số tiền phải thu)
    const a = khoaCot(S.A.hd[j]), kb = S.B.hd.map(khoaCot);
    let i = kb.indexOf(a);
    if (i < 0 && a.length >= 4) i = kb.findIndex(b => b.length >= 4 && (a.includes(b) || b.includes(a)));
    return i;
  }
  function veCot() {
    const ka = +$("#dKhoaA").value, kb = +$("#dKhoaB").value;
    $("#dCot").innerHTML = S.A.hd.map((h, j) => j === ka ? "" : `<div class="col"><b title="${esc(h)}"><span class="tag a">A</span> ${esc(h)}</b>
      <select data-cj="${j}"><option value="-1">Không so sánh</option>${S.B.hd.map((hb, i) => i === kb ? "" : `<option value="${i}"${i === ghepCot(j) ? " selected" : ""}>B: ${esc(hb)}</option>`).join("")}</select></div>`).join("");
    $("#dCot").querySelectorAll("select").forEach(s => s.addEventListener("change", chay));
  }

  // ---------- đối chiếu ----------
  const chuanKhoa = v => {
    if (typeof v === "number") return String(v);
    const s = String(v ?? "").trim();
    const sd = window.LamSach.sdt(s);
    if (!sd.loi && sd.v && /\d{9}/.test(s.replace(/\D/g, ""))) return sd.v;
    return boDau(s).replace(/[\s.\-_/]/g, "").toUpperCase();
  };
  const chuanChu = v => boDau(hien(v)).replace(/\s+/g, " ").trim();
  function gom(ben, k, cap) {   // gộp dòng trùng mã: cộng dồn cột số, giữ giá trị đầu cho cột chữ
    const m = new Map();
    S[ben].rows.forEach((r, i) => {
      const key = chuanKhoa(r[k]); if (!key) return;
      if (!m.has(key)) m.set(key, {goc: r[k], rows: [r], dong: [i + 1], v: cap.map(c => r[c])});
      else { const g = m.get(key); g.rows.push(r); g.dong.push(i + 1);
        cap.forEach((c, t) => { const a = so(g.v[t]), b = so(r[c]); if (a !== null && b !== null) g.v[t] = a + b; }); }
    });
    return m;
  }
  function chay() {
    const ka = +$("#dKhoaA").value, kb = +$("#dKhoaB").value;
    const cap = [...$("#dCot").querySelectorAll("select")].map(s => [+s.dataset.cj, +s.value]).filter(p => p[1] >= 0);
    const A = gom("A", ka, cap.map(p => p[0])), B = gom("B", kb, cap.map(p => p[1]));
    const ds = [];
    A.forEach((ga, key) => {
      const gb = B.get(key);
      if (!gb) { ds.push({loai: "chiA", key, ma: ga.goc, a: ga, b: null}); return; }
      const lech = cap.map((p, t) => {
        const x = ga.v[t], y = gb.v[t], nx = so(x), ny = so(y);
        if (nx !== null && ny !== null) return Math.abs(nx - ny) > 0.5 ? {so: true, d: ny - nx} : null;
        return chuanChu(x) !== chuanChu(y) ? {so: false} : null;
      });
      ds.push({loai: lech.some(Boolean) ? "lech" : "khop", key, ma: ga.goc, a: ga, b: gb, lech});
    });
    B.forEach((gb, key) => { if (!A.has(key)) ds.push({loai: "chiB", key, ma: gb.goc, a: null, b: gb}); });
    const thuTu = {lech: 0, chiA: 1, chiB: 2, khop: 3};
    ds.sort((x, y) => thuTu[x.loai] - thuTu[y.loai]);
    const dem = l => ds.filter(d => d.loai === l).length;
    const trungA = [...A.values()].filter(g => g.rows.length > 1).length, trungB = [...B.values()].filter(g => g.rows.length > 1).length;
    kq = {ds, cap, ka, kb, A, B, trungA, trungB};
    $("#dKq").hidden = false;
    const tongLech = cap.map((p, t) => ds.filter(d => d.loai === "lech" && d.lech[t] && d.lech[t].so).reduce((s, d) => s + d.lech[t].d, 0));
    $("#dTong").innerHTML = `<div class="o"><b>${dem("khop")}</b>khớp</div><div class="l"><b>${dem("lech")}</b>lệch giá trị</div>
      <div class="a"><b>${dem("chiA")}</b>chỉ có ở A</div><div class="bb"><b>${dem("chiB")}</b>chỉ có ở B</div>
      ${cap.map((p, t) => tongLech[t] ? `<div class="s"><b>${hien(tongLech[t])}</b>tổng chênh lệch "${esc(S.A.hd[p[0]])}" (B trừ A)</div>` : "").join("")}`;
    $("#dLoc").innerHTML = [["tat", "Tất cả"], ["lech", "Lệch"], ["chiA", "Chỉ ở A"], ["chiB", "Chỉ ở B"], ["khop", "Khớp"]].map(([k, t]) => `<button class="${k === loc ? "" : "ghost"}" data-dl="${k}" type="button">${t}</button>`).join("");
    $("#dLoc").querySelectorAll("[data-dl]").forEach(b => b.addEventListener("click", () => { loc = b.dataset.dl; chay(); }));
    if (trungA || trungB) $("#dTong").innerHTML += `<div class="x"><b>${trungA + trungB}</b>mã xuất hiện nhiều lần (đã cộng dồn số tiền)</div>`;
    ve();
  }
  const TEN = {khop: "✓ Khớp", lech: "≠ Lệch", chiA: "Chỉ có ở A", chiB: "Chỉ có ở B"};
  function ve() {
    const {ds, cap} = kq;
    const cols = ["Kết quả", "Mã"].concat(...cap.map(p => [`A: ${S.A.hd[p[0]]}`, `B: ${S.B.hd[p[1]]}`]), cap.some(p => true) ? ["Ghi chú"] : []);
    const rows = ds.filter(d => loc === "tat" || d.loai === loc).map(d => {
      const o = [TEN[d.loai], d.ma], cls = [d.loai === "lech" ? "loi" : d.loai === "chiA" ? "ca" : d.loai === "chiB" ? "cb" : "", ""];
      cap.forEach((p, t) => {
        const sx = d.a ? d.a.v[t] : "", sy = d.b ? d.b.v[t] : "";
        o.push(so(sx) ?? sx); o.push(so(sy) ?? sy);   // số viết dạng chữ ("60.000.000") hiện như số
        const l = d.lech && d.lech[t];
        const c0 = d.loai === "chiA" ? "ca" : d.loai === "chiB" ? "cb" : l ? "loi" : "";
        cls.push(c0, c0);
      });
      const ghi = [];
      if (d.lech) d.lech.forEach((l, t) => { if (l && l.so) ghi.push(`${S.A.hd[cap[t][0]]}: B ${l.d > 0 ? "nhiều" : "ít"} hơn A ${hien(Math.abs(l.d))}`); else if (l) ghi.push(`${S.A.hd[cap[t][0]]} khác nhau`); });
      if (d.a && d.a.rows.length > 1) ghi.push(`A có ${d.a.rows.length} dòng mã này (dòng ${d.a.dong.join(", ")})`);
      if (d.b && d.b.rows.length > 1) ghi.push(`B có ${d.b.rows.length} dòng mã này (đã cộng dồn)`);
      o.push(ghi.join("; ")); cls.push("");
      return {o, cls};
    });
    veBang($("#dGrid"), cols, rows);
  }

  function tai() {
    if (!kq) return;
    const {ds, cap} = kq;
    const cols = ["Kết quả", "Mã"].concat(...cap.map(p => [`A: ${S.A.hd[p[0]]}`, `B: ${S.B.hd[p[1]]}`, "Chênh lệch (B - A)"]));
    const dong = l => ds.filter(d => !l || d.loai === l).map(d => [TEN[d.loai], hien(d.ma)].concat(...cap.map((p, t) => {
      const x = d.a ? d.a.v[t] : "", y = d.b ? d.b.v[t] : "", nx = so(x), ny = so(y);
      return [nx ?? x, ny ?? y, nx !== null && ny !== null ? ny - nx : ""]; })));
    const goc = (ben, l) => [S[ben].hd, ...ds.filter(d => d.loai === l).flatMap(d => d[ben === "A" ? "a" : "b"].rows)];
    taiWb([["Tất cả", sheetTu([cols, ...dong()])], ["Lệch", sheetTu([cols, ...dong("lech")])],
      ["Chỉ có ở A", sheetTu(goc("A", "chiA"))], ["Chỉ có ở B", sheetTu(goc("B", "chiB"))], ["Khớp", sheetTu([cols, ...dong("khop")])]],
      "doi-chieu_" + S.A.ten.replace(/\.\w+$/, "") + "_vs_" + S.B.ten.replace(/\.\w+$/, "") + ".xlsx");
  }

  ganDrop($("#dDropA"), $("#dFileA"), fs => napFile("A", fs[0]));
  ganDrop($("#dDropB"), $("#dFileB"), fs => napFile("B", fs[0]));
  $("#dSheetA").addEventListener("change", e => napSheet("A", +e.target.value));
  $("#dSheetB").addEventListener("change", e => napSheet("B", +e.target.value));
  $("#dKhoaA").addEventListener("change", () => { veCot(); chay(); });
  $("#dKhoaB").addEventListener("change", () => { veCot(); chay(); });
  $("#dMau").addEventListener("click", async () => { const m = mau(); S.A = S.B = null; await napFile("A", m.A); await napFile("B", m.B); });
  $("#dTai").addEventListener("click", tai);
  window._dc = {get kq() { return kq; }, tai};
})();
