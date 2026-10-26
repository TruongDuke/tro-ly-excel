/* Gộp nhiều file / nhiều sheet thành một bảng, tự khớp cột theo tên */
(function () {
  "use strict";
  const {$, esc, docBuf, docWb, bangSheet, khoaCot, ganDrop, veBang, sheetTu, taiWb, fileMau, rong} = window.TL;
  let files = [];   // [{ten, wb}]
  let kq = null;

  function mau() {
    return [
      {ten: "Doanh-so-Ha-Noi.xlsx", buf: fileMau([["T10", [["Mã NV", "Họ và tên", "SĐT", "Doanh số"], ["HN01", "Nguyễn Văn An", "0912345678", 125000000], ["HN02", "Trần Thị Hà", "0987654321", 98000000], ["HN03", "Lê Hoàng Phúc", "0903111222", 143500000]]]])},
      {ten: "Doanh-so-HCM.xlsx", buf: fileMau([["Tháng 10", [["Họ tên", "Mã nhân viên", "Số điện thoại", "Doanh số (VNĐ)"], ["Phạm Minh Châu", "SG01", "0976543210", 156000000], ["Võ Thanh Tùng", "SG02", "0935222333", 87500000]]]])},
      {ten: "Doanh-so-Da-Nang.xlsx", buf: fileMau([["DN", [["BÁO CÁO DOANH SỐ CHI NHÁNH ĐÀ NẴNG - THÁNG 10/2026"], [], ["STT", "Mã NV", "Họ và tên", "Điện thoại", "Doanh số", "Ghi chú"], [1, "DN01", "Bùi Quốc Bảo", "0868777888", 76000000, "Nhân viên mới"], [2, "DN02", "Hoàng Thị Lan Anh", "0399888777", 91200000, ""]]]])},
    ];
  }

  async function them(fs) {
    for (const f of fs) {
      try { files.push({ten: f.name || f.ten, wb: docWb(f.buf || await docBuf(f))}); }
      catch (e) { alert(`Không đọc được ${f.name}: ${e.message}`); }
    }
    gop();
  }

  function gop() {
    $("#gDs").innerHTML = files.map((f, i) => `<div><span>📄 ${esc(f.ten)} <span class="muted">· ${f.wb.SheetNames.length} sheet</span></span><button class="ghost sm" data-xoa="${i}" type="button">Bỏ</button></div>`).join("");
    $("#gDs").querySelectorAll("[data-xoa]").forEach(b => b.addEventListener("click", () => { files.splice(+b.dataset.xoa, 1); gop(); }));
    if (!files.length) { $("#gKq").hidden = true; return; }
    const moiSheet = $("#gMoiSheet").checked, boTrong = $("#gBoTrong").checked, nguon = $("#gNguon").checked;
    const cot = [], viTri = new Map(), khop = [];   // cot: tên cột gộp; viTri: khoá -> chỉ số
    const dong = [];
    files.forEach(f => {
      (moiSheet ? f.wb.SheetNames : f.wb.SheetNames.slice(0, 1)).forEach(sn => {
        const b = bangSheet(f.wb.Sheets[sn]);
        if (!b.hd.length) return;
        const map = b.hd.map(h => {
          const k = khoaCot(h);
          if (!viTri.has(k)) { viTri.set(k, cot.length); cot.push(h); }
          return viTri.get(k);
        });
        khop.push({ten: f.ten, sheet: sn, hd: b.hd, map, tieuDe: b.dongTieuDe});
        b.rows.forEach(r => {
          if (boTrong && r.every(rong)) return;
          const o = []; r.forEach((v, j) => { o[map[j]] = v; });
          dong.push({nguon: `${f.ten} › ${sn}`, o});
        });
      });
    });
    // cột chỉ có ở một số file
    const coMat = cot.map((_, j) => khop.filter(k => k.map.includes(j)).length);
    kq = {cot, dong, khop, coMat, nguon};
    $("#gKq").hidden = false;
    $("#gTong").innerHTML = `<div class="o"><b>${dong.length}</b>dòng sau khi gộp</div><div><b>${files.length}</b>file · ${khop.length} sheet</div>
      <div><b>${cot.length}</b>cột chung</div><div class="${coMat.some(n => n < khop.length) ? "s" : "o"}"><b>${coMat.filter(n => n < khop.length).length}</b>cột không có ở mọi file</div>`;
    $("#gKhop").innerHTML = `<details${khop.length <= 4 ? " open" : ""}><summary class="muted" style="cursor:pointer">Xem cách máy khớp cột của từng file</summary><div class="grid" style="margin-top:8px;max-height:300px"><table><thead><tr><th>File › sheet</th>${cot.map((c, j) => `<th${coMat[j] < khop.length ? ' style="color:var(--sua-ink)"' : ""}>${esc(c)}</th>`).join("")}</tr></thead><tbody>
      ${khop.map(k => `<tr><td>${esc(k.ten)} › ${esc(k.sheet)}${k.tieuDe > 1 ? ` <span class="muted">(tiêu đề ở dòng ${k.tieuDe})</span>` : ""}</td>${cot.map((_, j) => { const i = k.map.indexOf(j); return `<td class="${i < 0 ? "loi" : ""}">${i < 0 ? "không có" : esc(k.hd[i])}</td>`; }).join("")}</tr>`).join("")}
      </tbody></table></div></details>`;
    veBang($("#gGrid"), (nguon ? ["Nguồn"] : []).concat(cot), dong.map(d => ({o: (nguon ? [d.nguon] : []).concat(cot.map((_, j) => d.o[j] ?? ""))})));
  }

  function tai() {
    if (!kq) return;
    const hd = (kq.nguon ? ["Nguồn"] : []).concat(kq.cot);
    const aoa = [hd, ...kq.dong.map(d => (kq.nguon ? [d.nguon] : []).concat(kq.cot.map((_, j) => d.o[j] ?? "")))];
    const kh = [["File", "Sheet", "Dòng tiêu đề", ...kq.cot], ...kq.khop.map(k => [k.ten, k.sheet, k.tieuDe, ...kq.cot.map((_, j) => { const i = k.map.indexOf(j); return i < 0 ? "(không có)" : k.hd[i]; })])];
    taiWb([["Đã gộp", sheetTu(aoa)], ["Cách khớp cột", sheetTu(kh)]], "da-gop_" + kq.dong.length + "-dong.xlsx");
  }

  ganDrop($("#gDrop"), $("#gFile"), them);
  $("#gMau").addEventListener("click", () => { files = []; them(mau()); });
  ["#gMoiSheet", "#gBoTrong", "#gNguon"].forEach(id => $(id).addEventListener("change", gop));
  $("#gTai").addEventListener("click", tai);
  window._gop = {get kq() { return kq; }, tai};
})();
