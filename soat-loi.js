/* ================= SOÁT LỖI FILE EXCEL (chạy ngay trong trình duyệt, không cần Claude) ================= */
const SOAT = (() => {
  const esc = window.TL.esc;
  const colNum = s => s.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const colName = n => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const REF = /(?<![A-Za-z_\d.!'])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![\d(A-Za-z_])/g;
  const fmt = n => typeof n === "number" ? n.toLocaleString("vi-VN", {maximumFractionDigits: 2}) : n;
  function pattern(f, r, c) {           // công thức dạng tương đối (R1C1) để so giữa các dòng
    return f.replace(/"[^"]*"/g, '""').replace(REF, (m, ac, col, ar, row) =>
      (ar ? "R" + row : "R[" + (+row - r) + "]") + (ac ? "C" + colNum(col) : "C[" + (colNum(col) - c) + "]"));
  }
  const looksNumText = v => typeof v === "string" && /\d/.test(v) &&
    (/^\s*-?\d{1,3}([.,]\d{3})+\s*$/.test(v) || /^\s*-?\d+([.,]\d+)?\s*$/.test(v));
  const isNum = cell => cell && cell.t === "n" && !cell.f;
  const parseNumText = v => { const t = String(v).trim();
    return /^-?\d{1,3}([.,]\d{3})+$/.test(t) ? +t.replace(/[.,]/g, "") : +t.replace(",", "."); };
  const so = x => (Math.round(x * 100) / 100).toLocaleString("vi-VN", {maximumFractionDigits: 2});
  const ERR = /^#(REF!|DIV\/0!|VALUE!|N\/A|NAME\?|NUM!|NULL!)/;

  function audit(wb) {
    const found = [], stats = {sheets: 0, cells: 0, formulas: 0};
    const meta = (wb.Workbook && wb.Workbook.Sheets) || [];
    const refsTo = {};                   // "Sheet!A1" -> [ô có công thức tham chiếu tới]
    wb.SheetNames.forEach((name, si) => {
      const ws = wb.Sheets[name]; if (!ws || !ws["!ref"]) return;
      stats.sheets++;
      const rg = XLSX.utils.decode_range(ws["!ref"]);
      const get = (r, c) => ws[XLSX.utils.encode_cell({r: r - 1, c: c - 1})];
      const valOf = (r, c, asText) => { const x = get(r, c); if (!x) return 0; if (x.t === "n") return x.v;
        return asText && x.t === "s" && looksNumText(x.v) ? parseNumText(x.v) : 0; };
      const sumR = (r1, c1, r2, c2, asText) => { let t = 0; for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) t += valOf(r, c, asText); return t; };
      let filled = 0;
      for (const k in ws) if (k[0] !== "!") {
        filled++; const cell = ws[k];
        if (cell.f) { stats.formulas++;
          const dc = XLSX.utils.decode_cell(k);
          const re = /(?:(?:'([^']+)'|([A-Za-z0-9_À-ỹ]+))!)?(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?/g; let m;
          while ((m = re.exec(cell.f))) {
            const sh = m[1] || m[2] || name;
            const a = XLSX.utils.decode_cell(m[3].replace(/\$/g, "")), b = m[4] ? XLSX.utils.decode_cell(m[4].replace(/\$/g, "")) : a;
            if ((b.r - a.r + 1) * (b.c - a.c + 1) > 5000) continue;
            for (let r = a.r; r <= b.r; r++) for (let c = a.c; c <= b.c; c++)
              (refsTo[sh + "!" + XLSX.utils.encode_cell({r, c})] ||= []).push(name + "!" + XLSX.utils.encode_cell(dc));
          }
        }
      }
      stats.cells += filled;
      const hidden = meta[si] && meta[si].Hidden;
      if (hidden && filled >= 2) found.push({sheet: name, addr: "Cả sheet", lvl: "info", key: "an",
        ttl: "Sheet đang bị ẩn nhưng có " + filled + " ô dữ liệu",
        why: "Dễ có số liệu cũ hoặc bản nháp bị công thức dùng nhầm. Kiểm tra sheet này có còn cần không."});
      (ws["!cols"] || []).forEach((cl, ci) => { if (cl && cl.hidden) {
        let n = 0; for (let r = rg.s.r + 1; r <= rg.e.r + 1; r++) if (get(r, ci + 1)) n++;
        if (n) found.push({sheet: name, addr: "Cột " + colName(ci + 1), lvl: "info", key: "an",
          ttl: "Cột bị ẩn có " + n + " ô dữ liệu", why: "Người đọc báo cáo không thấy cột này. Kiểm tra cột ẩn có ảnh hưởng tới tổng không."}); } });

      for (let r = rg.s.r + 1; r <= rg.e.r + 1; r++) for (let c = rg.s.c + 1; c <= rg.e.c + 1; c++) {
        const cell = get(r, c); if (!cell) continue;
        const addr = colName(c) + r;
        const add = o => found.push(Object.assign({sheet: name, addr}, o));
        // 1. lỗi hàm
        if (cell.t === "e" || (typeof cell.v === "string" && ERR.test(cell.v)) || (cell.f && cell.f.includes("#REF!"))) {
          add({lvl: "bad", key: "loi", ttl: "Ô đang báo lỗi " + (cell.w || cell.v || "#REF!"),
            why: "Các ô cộng dồn từ ô này cũng sẽ lỗi hoặc sai.", f: cell.f});
          continue;
        }
        // 2. số lưu dạng chữ
        if (cell.t === "s" && looksNumText(cell.v)) {
          let numericNear = 0;
          for (const dr of [-2, -1, 1, 2]) if (isNum(get(r + dr, c)) || (get(r + dr, c) || {}).f) numericNear++;
          for (const dc of [-2, -1, 1, 2]) if (isNum(get(r, c + dc)) || (get(r, c + dc) || {}).f) numericNear++;
          if (numericNear >= 2) {
            const users = refsTo[name + "!" + addr] || [];
            const tac = [];
            for (const u of users) { const [ush, ua] = u.split("!"); if (ush !== name) continue;
              const uc = ws[ua]; if (uc && typeof uc.v === "number" && /^SUM\(/i.test(uc.f || ""))
                tac.push(ua + " đang ra " + so(uc.v) + ", đúng phải là " + so(uc.v + parseNumText(cell.v))); }
            add({lvl: users.length ? "bad" : "warn", key: "chu", tac: tac.join("; "),
              ttl: "Số \"" + cell.v.trim() + "\" đang lưu dạng chữ",
              why: users.length ? "Excel không cộng ô này, nên " + users.slice(0, 3).join(", ") + (users.length > 3 ? "…" : "") +
                " đang ra kết quả sai. Gõ lại thành số." : "Excel không tính ô này vào phép cộng. Gõ lại thành số.",
              v: cell.v});
          }
          continue;
        }
        // 3. SUM bỏ sót
        const sm = cell.f && cell.f.match(/^SUM\(\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\)$/i);
        if (sm) {
          const c1 = colNum(sm[1].toUpperCase()), r1 = +sm[2], c2 = colNum(sm[3].toUpperCase()), r2 = +sm[4];
          const miss = [];
          if (c1 === c2 && c1 === c && r2 < r) {             // cộng dọc
            for (let rr = r2 + 1; rr < r; rr++) { const x = get(rr, c); if (x && (x.t === "n" || looksNumText(x.v))) miss.push(colName(c) + rr); }
            for (let rr = r1 - 1; rr > 0; rr--) { const x = get(rr, c); if (!x || x.t !== "n" || x.f) break; miss.unshift(colName(c) + rr); }
          } else if (r1 === r2 && r1 === r && c2 < c) {      // cộng ngang
            for (let cc = c2 + 1; cc < c; cc++) { const x = get(r, cc); if (x && x.t === "n") miss.push(colName(cc) + r); }
          }
          if (miss.length) {
            const them = miss.reduce((t, a) => { const d = XLSX.utils.decode_cell(a); return t + valOf(d.r + 1, d.c + 1, true); }, 0);
            add({lvl: "bad", key: "sum", ttl: "Hàm SUM bỏ sót " + miss.length + " ô có số",
            tac: typeof cell.v === "number" ? "Đang ra " + so(cell.v) + ", đúng phải là " + so(cell.v + them) + " (thiếu " + so(them) + ")" : "",
            why: "Phạm vi cộng chưa gồm " + miss.join(", ") + ". Tổng đang thiếu.", f: cell.f}); continue; }
        }
        // 4. ô gõ đè / công thức lệch so với hàng xóm cùng cột hoặc cùng dòng
        const pat = (rr, cc) => { const x = get(rr, cc); return x && x.f ? pattern(x.f, rr, cc) : null; };
        const pairs = [[pat(r - 1, c), pat(r + 1, c), "cột"], [pat(r - 2, c), pat(r - 1, c), "cột"],
          [pat(r + 1, c), pat(r + 2, c), "cột"], [pat(r, c - 1), pat(r, c + 1), "dòng"]];
        const ref = pairs.find(p => p[0] && p[0] === p[1]);
        if (ref) {
          if (isNum(cell)) {
            const mm = ref[0].match(/^SUM\(R\[(-?\d+)\]C\[(-?\d+)\]:R\[(-?\d+)\]C\[(-?\d+)\]\)$/);
            let tac = "";
            if (mm) { const dung = sumR(r + +mm[1], c + +mm[2], r + +mm[3], c + +mm[4], true);
              if (Math.abs(dung - cell.v) > 1e-9) tac = "Cộng lại các ô đúng phải là " + so(dung) + ", ô đang ghi " + so(cell.v) + " (lệch " + (cell.v > dung ? "+" : "") + so(cell.v - dung) + ")"; }
            add({lvl: "bad", key: "de", tac, ttl: "Ô bị gõ đè số " + fmt(cell.v) + " thay cho công thức",
              why: "Các ô xung quanh cùng " + ref[2] + " đều là công thức. Khi số liệu thay đổi, ô này không tự cập nhật.",
              v: cell.v}); continue;
          }
          const tongCot = cell.f && new RegExp("^(SUM|SUBTOTAL\\(\\d+,)\\(?\\$?" + colName(c) + "\\$?\\d+:\\$?" + colName(c) + "\\$?\\d+\\)$", "i").test(cell.f);
          const kep = (pat(r - 1, c) && pat(r - 1, c) === pat(r + 1, c)) || (pat(r, c - 1) && pat(r, c - 1) === pat(r, c + 1));
          if (cell.f && !tongCot && kep && pattern(cell.f, r, c) !== ref[0]) {
            const nb = ref[2] === "cột" ? (pat(r - 1, c) === ref[0] ? colName(c) + (r - 1) : colName(c) + (r + 1))
                                         : colName(c - 1) + r;
            const nbF = (get(XLSX.utils.decode_cell(nb).r + 1, XLSX.utils.decode_cell(nb).c + 1) || {}).f;
            add({lvl: "warn", key: "lech", ttl: "Công thức khác với các ô cùng " + ref[2],
              why: "Ô " + nb + " là =" + nbF + ", còn ô này là =" + cell.f + ". Kiểm tra có bị sót hoặc lệch ô không.",
              f: cell.f}); continue;
          }
        }
        // 5. số gõ cứng trong công thức
        if (cell.f) {
          const lits = cell.f.replace(/"[^"]*"/g, "").replace(REF, "").match(/(?<![A-Za-z_\d.$])\d+(?:\.\d+)?(?![\d(A-Za-z_])/g) || [];
          const hard = lits.filter(x => x.includes(".") ? true : +x >= 100 && ![100, 1000, 365].includes(+x));
          if (hard.length) add({lvl: "info", key: "cung", ttl: "Công thức có số gõ cứng (ví dụ " + hard[0].replace(".", ",") + ")",
            why: "Nếu tỷ lệ hoặc mức này thay đổi, phải sửa từng ô. Nên đặt số này vào một ô riêng rồi tham chiếu tới.",
            f: cell.f});
        }
      }
    });
    // gộp "số gõ cứng" lặp lại trong cùng cột
    const out = [], seen = {};
    for (const f of found) {
      if (f.key === "cung") { const k = f.sheet + "|" + f.addr.replace(/\d+/g, "");
        if (seen[k]) { seen[k].n++; seen[k].last = f.addr; continue; } seen[k] = f; f.n = 1; }
      out.push(f);
    }
    out.forEach(f => { if (f.n > 1) f.addr = f.addr + ":" + f.last; });
    const order = {bad: 0, warn: 1, info: 2};
    out.sort((a, b) => order[a.lvl] - order[b.lvl]);
    return {found: out, stats};
  }

  function showCell(cell) {
    if (!cell) return {t: "", cls: ""};
    if (cell.t === "e") return {t: cell.w || "#LỖI", cls: "n"};
    if (cell.t === "n") { const pct = cell.z && String(cell.z).includes("%");
      return {t: pct ? (cell.v * 100).toLocaleString("vi-VN", {maximumFractionDigits: 1}) + "%" : so(cell.v), cls: "n"}; }
    if (cell.t === "s" && looksNumText(cell.v)) return {t: cell.v, cls: "txtnum"};
    return {t: cell.v == null ? "" : String(cell.v), cls: ""};
  }
  function gridHTML(wb, sheet, si, marks, hidden) {
    const ws = wb.Sheets[sheet]; if (!ws || !ws["!ref"]) return "";
    const rg = XLSX.utils.decode_range(ws["!ref"]);
    const r2 = Math.min(rg.e.r, rg.s.r + 39), c2 = Math.min(rg.e.c, rg.s.c + 11);
    let h = `<tr><th class="rh"></th>`;
    for (let c = rg.s.c; c <= c2; c++) h += `<th>${colName(c + 1)}</th>`;
    h += "</tr>";
    for (let r = rg.s.r; r <= r2; r++) {
      h += `<tr><th class="rh">${r + 1}</th>`;
      for (let c = rg.s.c; c <= c2; c++) {
        const a = XLSX.utils.encode_cell({r, c}), cell = ws[a], d = showCell(cell), m = marks[a];
        const tip = cell && cell.f ? ` title="=${esc(cell.f)}"` : cell && cell.t === "s" && looksNumText(cell.v) ? ` title="Số đang lưu dạng chữ"` : "";
        h += `<td id="sl-${si}-${a}" class="${d.cls}${m ? " m-" + m : ""}"${tip}>${esc(d.t)}</td>`;
      }
      h += "</tr>";
    }
    const more = rg.e.r > r2 || rg.e.c > c2 ? " · chỉ hiện 40 dòng, 12 cột đầu" : "";
    return `<div class="sheetview"><div class="cap"><b>${esc(sheet)}${hidden ? " (sheet đang ẩn)" : ""}</b><span>Rê chuột vào ô để xem công thức${more}</span></div>
      <div class="gridwrap"><table class="xl">${h}</table></div></div>`;
  }

  function render(root, name, res, isSample, wb) {
    const {found, stats} = res;
    const n = l => found.filter(f => f.lvl === l).length;
    const groups = {};
    found.forEach(f => (groups[f.sheet] ||= []).push(f));
    let html = `<div class="status"><h3>${esc(name)}</h3>${isSample ? '<span class="stag">File mẫu</span>' : ""}
      <span class="meta">${stats.sheets} sheet · ${stats.cells.toLocaleString("vi-VN")} ô có dữ liệu · ${stats.formulas.toLocaleString("vi-VN")} công thức</span></div>
      <div class="tally"><div class="t-bad"><b>${n("bad")}</b><span>Lỗi làm sai số</span></div>
      <div class="t-warn"><b>${n("warn")}</b><span>Cần xem lại</span></div>
      <div class="t-info"><b>${n("info")}</b><span>Lưu ý</span></div></div>`;
    if (!found.length) html += `<div class="okbox">Không thấy lỗi nào trong các loại lỗi trang này soát. Vẫn nên xem lại các con số quan trọng bằng mắt.</div>`;
    const meta = (wb.Workbook && wb.Workbook.Sheets) || [];
    const lvlRank = {bad: 3, warn: 2, info: 1};
    if (found.length) html += `<div class="legend"><span><i style="background:var(--bad)"></i>Ô làm sai số</span><span><i style="background:var(--warn)"></i>Cần xem lại</span><span><i style="background:var(--info)"></i>Lưu ý</span><span><i style="background:var(--grid)"></i>Tam giác xanh: số lưu dạng chữ</span></div>`;
    wb.SheetNames.forEach((sh, si) => {
      if (!groups[sh]) return;
      const marks = {};
      groups[sh].forEach(f => { const a = /^[A-Z]+\d+/.exec(f.addr); if (!a) return;
        const cells = f.addr.includes(":") ? (() => { const r = XLSX.utils.decode_range(f.addr), o = [];
          for (let rr = r.s.r; rr <= r.e.r; rr++) for (let cc = r.s.c; cc <= r.e.c; cc++) o.push(XLSX.utils.encode_cell({r: rr, c: cc})); return o; })() : [a[0]];
        cells.forEach(c => { if (!marks[c] || lvlRank[f.lvl] > lvlRank[marks[c]]) marks[c] = f.lvl; }); });
      html += gridHTML(wb, sh, si, marks, meta[si] && meta[si].Hidden);
    });
    for (const sh in groups) {
      const si = wb.SheetNames.indexOf(sh);
      html += `<div class="sheet"><h4>${esc(sh)}</h4>` + groups[sh].map(f => `
        <div class="item ${f.lvl}" data-cell="sl-${si}-${esc((/^[A-Z]+\d+/.exec(f.addr) || [""])[0])}" tabindex="0" role="button" aria-label="Xem ô ${esc(f.addr)} trong bảng">
          <span class="addr">${esc(f.addr)}</span>
          <div class="ttl">${esc(f.ttl)}</div>${f.tac ? `<div class="tac">${esc(f.tac)}</div>` : ""}<div class="why">${esc(f.why)}</div>
          ${(f.f || f.v !== undefined) ? `<div class="ev">${f.f ? `<span>Công thức <code>=${esc(f.f)}</code></span>` : ""}${f.v !== undefined && !f.f ? `<span>Giá trị <code>${esc(f.v)}</code></span>` : ""}</div>` : ""}
        </div>`).join("") + `</div>`;
    }
    if (found.length) html += `<div class="sec"><div class="row"><h3 style="margin:0;font-size:15px">Nhờ Claude chỉ cách sửa</h3><button class="sm" id="sl-ask" type="button">Mở Claude với danh sách lỗi →</button></div>
      <p class="note">Bấm để mở Claude (cần tài khoản Claude, gói miễn phí cũng được) với lệnh đã soạn sẵn, bạn chỉ cần bấm Gửi. Claude chỉ nhận danh sách lỗi (địa chỉ ô, giá trị, công thức), không nhận cả file.</p></div>`;
    root.innerHTML = html;
    root.querySelectorAll(".item[data-cell]").forEach(it => {
      const go = () => { const td = document.getElementById(it.dataset.cell); if (!td) return;
        td.scrollIntoView({block: "center", inline: "center", behavior: "smooth"});
        td.classList.remove("flash"); void td.offsetWidth; td.classList.add("flash"); };
      it.addEventListener("click", go); it.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go(); } });
    });
    const ask = root.querySelector("#sl-ask");
    if (ask) ask.addEventListener("click", () => {
      const list = found.map(f => `- [${f.lvl === "bad" ? "làm sai số" : f.lvl === "warn" ? "cần xem" : "lưu ý"}] ${f.sheet}!${f.addr}: ${f.ttl}. ${f.tac ? " Ảnh hưởng: " + f.tac + "." : ""} ${f.why}${f.f ? " Công thức: =" + f.f : ""}${f.v !== undefined && !f.f ? " Giá trị: " + f.v : ""}`).join("\n");
      const p = `Bạn là kế toán trưởng giàu kinh nghiệm, đang giúp một nhân viên văn phòng Việt Nam không rành Excel.
Một công cụ đã soát file "${name}" và tìm ra các lỗi sau (chỉ dựa vào danh sách này, không bịa thêm):
${list}

Viết bằng tiếng Việt, xưng "mình", gọi "bạn":
1. Một câu kết luận: file có gửi được chưa, lỗi nào ảnh hưởng con số báo cáo nhiều nhất.
2. Với từng lỗi "làm sai số": giải thích ngắn vì sao sai, và các bước bấm cụ thể trong Excel để sửa.
3. Gộp ngắn các mục "cần xem" và "lưu ý" trong 2-3 câu.`;
      const url = "https://claude.ai/new?q=" + encodeURIComponent(p);
      if (url.length < 12000) window.open(url, "_blank", "noopener");
      else navigator.clipboard.writeText(p).then(() => alert("Danh sách lỗi dài, đã chép lệnh. Mở claude.ai rồi dán vào khung chat."));
    });
  }

  function load(root, buf, name, isSample) {
    try {
      const wb = XLSX.read(buf, {type: "array", cellFormula: true, cellStyles: true, cellNF: true});
      render(root, name, audit(wb), isSample, wb);
    } catch (e) {
      root.innerHTML = `<div class="item bad"><span class="addr">!</span><div class="ttl">Không đọc được file này</div>
        <div class="why">Kiểm tra file có phải Excel (.xlsx, .xls) và không bị đặt mật khẩu. Chi tiết: ${esc(e.message)}</div></div>`;
    }
  }

  // file mẫu: báo cáo doanh thu quý có cài sẵn lỗi
  function sampleBook() {
    const n = (v) => ({t: "n", v, z: "#,##0"}), fx = (f, v, z) => (typeof v === "string" ? {t: "e", f, v: 7, w: v} : {t: "n", f, v, z: z || "#,##0"});
    const pc = (f, v) => fx(f, v, "0.0%");
    const data = [
      ["DOANH THU THEO CHI NHÁNH - QUÝ III/2026 (triệu đồng)"],
      ["Chi nhánh", "Tháng 7", "Tháng 8", "Tháng 9", "Cộng quý", "% kế hoạch", "Thưởng (1,5%)"],
      ["Hà Nội", n(5620), n(6310), n(6950), fx("SUM(B3:D3)", 18880), pc("E3/18500", 18880 / 18500), fx("E3*0.015", 283.2, "#,##0.0")],
      ["TP. Hồ Chí Minh", n(5480), n(6280), n(7120), n(19500), pc("E4/19000", 19500 / 19000), fx("E4*0.015", 292.5, "#,##0.0")],
      ["Đà Nẵng", n(2150), {t: "s", v: "2.530"}, n(2960), fx("SUM(B5:D5)", 5110), pc("E5/6500", 5110 / 6500), fx("E5*0.015", 76.65, "#,##0.0")],
      ["Cần Thơ", n(1800), n(1950), n(2210), fx("SUM(B6:C6)", 3750), pc("E6/6000", 3750 / 6000), fx("E6*0.015", 56.25, "#,##0.0")],
      ["Hải Phòng", n(1420), n(1610), n(1880), fx("SUM(B7:D7)", 4910), fx("E7/0", "#DIV/0!"), fx("E7*0.015", 73.65, "#,##0.0")],
      ["Nha Trang", n(950), n(1120), n(1305), fx("SUM(B8:D8)", 3375), pc("E8/3300", 3375 / 3300), fx("E8*0.015", 50.625, "#,##0.0")],
      ["Tổng cộng", fx("SUM(B3:B6)", 15050), fx("SUM(C3:C8)", 17270), fx("SUM(D3:D8)", 22425), fx("SUM(E3:E8)", 55525), "", fx("SUM(G3:G8)", 832.875, "#,##0.0")],
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [{wch: 18}, {wch: 9}, {wch: 9}, {wch: 9}, {wch: 10}, {wch: 11}, {wch: 13}];
    const draft = XLSX.utils.aoa_to_sheet([["Số liệu tạm tháng 9 - bản cũ"], ["Hà Nội", 7210], ["TP. Hồ Chí Minh", 7350], ["Đà Nẵng", 3100], ["Cần Thơ", 2400]]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Doanh thu Q3");
    XLSX.utils.book_append_sheet(wb, draft, "nhap T9");
    wb.Workbook = {Sheets: [{Hidden: 0}, {Hidden: 1}]};
    return XLSX.write(wb, {bookType: "xlsx", type: "array"});
  }
  return {load, sampleBook};
})();


(function () {
  const {$, ganDrop, docBuf} = window.TL;
  const root = $("#sKq");
  const mo = (buf, ten, mau) => { root.hidden = false; SOAT.load(root, buf, ten, mau); };
  ganDrop($("#sDrop"), $("#sFile"), async fs => mo(await docBuf(fs[0]), fs[0].name, false));
  $("#sMau").addEventListener("click", () => mo(SOAT.sampleBook(), "Bao-cao-doanh-thu-Q3-mau.xlsx", true));
})();
