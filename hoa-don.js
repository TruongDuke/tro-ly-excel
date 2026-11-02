/* Hoá đơn điện tử XML (chuẩn TT78, QĐ 1450/QĐ-TCT) → bảng kê Excel, kèm soát lỗi từng hoá đơn */
(function () {
  "use strict";
  const {$, esc, boDau, docBuf, ganDrop, hien, veBang, sheetTu, taiWb} = window.TL;
  let hds = [], hong = [], kq = null, loc = "tat";

  // ---------- đọc XML theo tên thẻ, bỏ qua namespace ----------
  const con = (el, ten) => el && [...el.children].find(c => c.localName === ten);
  const di = (el, duong) => duong.split("/").reduce((e, t) => con(e, t), el);
  const chu = (el, duong) => (di(el, duong)?.textContent || "").trim();
  const so = (el, duong) => { const s = chu(el, duong).replace(/,/g, ""); return s === "" || isNaN(+s) ? null : +s; };
  const tatCa = (el, ten) => el ? [...el.getElementsByTagNameNS("*", ten)] : [];

  function docHoaDon(xml, tenFile) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("file XML bị hỏng");
    const goc = tatCa(doc, "HDon")[0];
    const dl = goc && con(goc, "DLHDon");
    if (!dl) throw new Error("không phải hoá đơn điện tử theo chuẩn TT78 (không có thẻ HDon/DLHDon)");
    const tc = con(dl, "TTChung"), nd = con(dl, "NDHDon"), tt = con(nd, "TToan");
    const ngay = chu(tc, "NLap").slice(0, 10).split("-");
    const tg = so(tc, "TGia") || 1, dvt = chu(tc, "DVTTe") || "VND";
    const lq = con(tc, "TTHDLQuan");
    const dong = tatCa(con(nd, "DSHHDVu"), "HHDVu").map(h => ({
      tchat: chu(h, "TChat"), stt: chu(h, "STT"), ma: chu(h, "MHHDVu"), ten: chu(h, "THHDVu"), dvt: chu(h, "DVTinh"),
      sl: so(h, "SLuong"), gia: so(h, "DGia"), ck: so(h, "STCKhau"), tien: so(h, "ThTien"), ts: chu(h, "TSuat"),
    }));
    const suat = tatCa(con(tt, "THTTLTSuat"), "LTSuat").map(l => ({ts: chu(l, "TSuat"), tien: so(l, "ThTien"), thue: so(l, "TThue")}));
    return {
      file: tenFile, mau: chu(tc, "KHMSHDon"), kh: chu(tc, "KHHDon"), so: chu(tc, "SHDon"), loaiHd: chu(tc, "THDon"),
      ngay: ngay.length === 3 ? new Date(+ngay[0], +ngay[1] - 1, +ngay[2]) : null, dvt, tg, httt: chu(tc, "HTTToan"),
      lienQuan: lq ? {tc: chu(lq, "TCHDon"), mau: chu(lq, "KHMSHDCLQuan"), kh: chu(lq, "KHHDCLQuan"), so: chu(lq, "SHDCLQuan")} : null,
      ban: {ten: chu(nd, "NBan/Ten"), mst: chu(nd, "NBan/MST"), dc: chu(nd, "NBan/DChi")},
      mua: {ten: chu(nd, "NMua/Ten") || chu(nd, "NMua/HVTNMHang"), mst: chu(nd, "NMua/MST"), dc: chu(nd, "NMua/DChi")},
      dong, suat,
      chuaThue: so(tt, "TgTCThue"), thue: so(tt, "TgTThue"), ck: so(tt, "TTCKTMai"), phi: so(tt, "TgTPhi"),
      tong: so(tt, "TgTTTBSo"), bangChu: chu(tt, "TgTTTBChu"),
      mcqt: chu(goc, "MCCQT"),
    };
  }

  // ---------- đọc số tiền bằng chữ → số, để so với số tiền bằng số ----------
  const SO = {khong: 0, mot: 1, mot_: 1, hai: 2, ba: 3, bon: 4, tu: 4, nam: 5, lam: 5, nham: 5, sau: 6, bay: 7, tam: 8, chin: 9};
  const BO_QUA = new Set(["le", "linh", "dong", "chan", "vnd", "vietnam", "viet", "va", "am", "tron"]);
  function docChu(s) {
    const tu = boDau(s).replace(/[^a-z ]/g, " ").split(/\s+/).filter(Boolean);
    if (!tu.length) return null;
    let kq = 0, nhom = 0, cho = null;
    for (const t of tu) {
      if (t in SO) { if (cho !== null) return null; cho = SO[t]; }
      else if (t === "muoi") { nhom += cho === null ? 10 : cho * 10; cho = null; }
      else if (t === "tram") { if (cho === null) return null; nhom += cho * 100; cho = null; }
      else if (t === "nghin" || t === "ngan" || t === "trieu") { kq += ((nhom + (cho || 0)) || 1) * (t === "trieu" ? 1e6 : 1e3); nhom = 0; cho = null; }
      else if (t === "ty" || t === "ti") { kq = (kq + nhom + (cho || 0) || 1) * 1e9; nhom = 0; cho = null; }
      else if (BO_QUA.has(t)) continue;
      else return null;   // gặp chữ lạ (xu, phẩy, đô la...) thì không kết luận
    }
    return kq + nhom + (cho || 0);
  }

  // ---------- soát từng hoá đơn ----------
  const mstHopLe = m => /^\d{10}(-\d{3})?$/.test(m) || /^\d{12}$/.test(m);
  const tienVnd = (h, v) => v === null ? null : Math.round(v * (h.dvt === "VND" ? 1 : h.tg));
  function soat() {
    // công ty mình: MST xuất hiện ở cùng một phía trên phần lớn hoá đơn
    const dem = (phia) => { const m = new Map(); hds.forEach(h => h[phia].mst && m.set(h[phia].mst, (m.get(h[phia].mst) || 0) + 1)); return [...m].sort((a, b) => b[1] - a[1])[0] || ["", 0]; };
    const [mstMua, nMua] = dem("mua"), [mstBan, nBan] = dem("ban");
    let minh = null;
    if (hds.length >= 3) {
      if (nMua >= nBan && nMua / hds.length >= .6) minh = {phia: "mua", mst: mstMua, chieu: "Đầu vào (mua vào)"};
      else if (nBan / hds.length >= .6) minh = {phia: "ban", mst: mstBan, chieu: "Đầu ra (bán ra)"};
    }
    const khoa = h => [h.ban.mst, h.mau, h.kh, String(+h.so || h.so)].join("|");
    const daGap = new Map(), biThay = new Map();
    hds.forEach(h => { if (h.lienQuan && h.lienQuan.so) biThay.set([h.ban.mst, h.lienQuan.mau, h.lienQuan.kh, String(+h.lienQuan.so || h.lienQuan.so)].join("|"), h); });

    hds.forEach(h => {
      const loi = [], luuY = [], k = khoa(h);
      h.trung = daGap.has(k);
      if (h.trung) loi.push(`Trùng với file ${daGap.get(k).file}, không cộng vào tổng`); else daGap.set(k, h);
      if (h.tong !== null && h.chuaThue !== null) {
        const thue = h.thue || 0, phi = h.phi || 0, ck = h.ck || 0;
        if (Math.abs(h.chuaThue + thue + phi - h.tong) > 1 && Math.abs(h.chuaThue + thue + phi - ck - h.tong) > 1)
          loi.push(`Tổng thanh toán ${hien(h.tong)} khác chưa thuế + thuế = ${hien(h.chuaThue + thue + phi)}`);
      }
      if (h.bangChu && h.tong !== null && h.dvt === "VND") {
        const v = docChu(h.bangChu);
        if (v !== null && v !== Math.round(Math.abs(h.tong))) loi.push(`Số tiền bằng chữ đọc ra ${hien(v)}, khác bằng số ${hien(h.tong)}`);
      }
      if (minh && h[minh.phia].mst !== minh.mst)
        loi.push(`MST ${minh.phia === "mua" ? "người mua" : "người bán"} là ${h[minh.phia].mst || "(trống)"}, khác MST công ty ${minh.mst}`);
      // cộng dòng hàng: hàng hoá, khuyến mại cộng; chiết khấu trừ; dòng ghi chú bỏ qua
      const dh = h.dong.filter(d => d.tien !== null && d.tchat !== "4");
      if (dh.length && h.chuaThue !== null) {
        const cong = dh.reduce((s, d) => s + (d.tchat === "3" ? -Math.abs(d.tien) : d.tien), 0);
        if (Math.abs(cong - h.chuaThue) > 2) luuY.push(`Cộng các dòng hàng được ${hien(cong)}, khác tiền chưa thuế ${hien(h.chuaThue)}`);
      }
      h.suat.forEach(s => {
        const r = /^(?:KHAC:)?(\d+(?:[.,]\d+)?)%$/.exec(s.ts.replace(/\s/g, ""));
        if (r && s.tien !== null && s.thue !== null) {
          const dung = Math.round(s.tien * parseFloat(r[1].replace(",", ".")) / 100);
          if (Math.abs(dung - s.thue) > Math.max(10, h.dong.length)) luuY.push(`Thuế suất ${s.ts}: ${hien(s.tien)} × ${s.ts} = ${hien(dung)}, hoá đơn ghi ${hien(s.thue)}`);
        }
      });
      if (/^C/i.test(h.kh) && !h.mcqt) luuY.push("Hoá đơn có mã nhưng file không có mã của cơ quan thuế (tải lại từ cổng thuế)");
      const nam = /^[CK](\d\d)/i.exec(h.kh);
      if (nam && h.ngay && +nam[1] !== h.ngay.getFullYear() % 100) luuY.push(`Ký hiệu ${h.kh} là năm 20${nam[1]} nhưng ngày lập năm ${h.ngay.getFullYear()}`);
      ["ban", "mua"].forEach(p => { if (h[p].mst && !mstHopLe(h[p].mst)) luuY.push(`MST ${p === "ban" ? "người bán" : "người mua"} "${h[p].mst}" sai định dạng`); });
      if (h.lienQuan) luuY.push(`${h.lienQuan.tc === "1" ? "Thay thế" : h.lienQuan.tc === "2" ? "Điều chỉnh" : "Liên quan"} hoá đơn ${h.lienQuan.kh} số ${h.lienQuan.so}`);
      const t = biThay.get(k);
      if (t && t !== h) loi.push(`Đã bị ${t.lienQuan.tc === "2" ? "điều chỉnh" : "thay thế"} bởi hoá đơn số ${t.so}, xem lại trước khi kê khai`);
      if (h.dvt !== "VND") luuY.push(`Tiền ${h.dvt}, đã quy đổi theo tỷ giá ${hien(h.tg)}`);
      h.loi = loi; h.luuY = luuY;
    });
    return minh;
  }

  // ---------- nạp file: .xml hoặc .zip tải từ cổng thuế ----------
  const giai = new TextDecoder("utf-8");
  async function them(fs) {
    for (const f of fs) {
      const ten = f.name || f.ten;
      try {
        const buf = f.buf || await docBuf(f);
        if (/\.zip$/i.test(ten)) {
          const z = XLSX.CFB.read(buf, {type: "array"});
          const xs = z.FileIndex.map((e, i) => [z.FullPaths[i], e]).filter(([p, e]) => /\.xml$/i.test(p) && e.content && e.content.length);
          if (!xs.length) hong.push({file: ten, ly: "file nén không có file XML"});
          xs.forEach(([p, e]) => napXml(giai.decode(e.content), `${ten} › ${p.split("/").pop()}`));
        } else napXml(giai.decode(buf), ten);
      } catch (e) { hong.push({file: ten, ly: e.message}); }
    }
    chay();
  }
  function napXml(xml, ten) {
    try { hds.push(docHoaDon(xml, ten)); } catch (e) { hong.push({file: ten, ly: e.message}); }
  }

  function chay() {
    $("#hDs").innerHTML = hds.length || hong.length ? `<p class="note">Đã nạp ${hds.length} hoá đơn${hong.length ? ` · ${hong.length} file không đọc được` : ""}. <button class="ghost sm" id="hXoa" type="button">Xoá hết</button></p>` : "";
    const x = $("#hXoa"); if (x) x.addEventListener("click", () => { hds = []; hong = []; chay(); });
    if (!hds.length && !hong.length) { $("#hKq").hidden = true; return; }
    hds.sort((a, b) => (a.ngay || 0) - (b.ngay || 0) || a.kh.localeCompare(b.kh) || (+a.so - +b.so));
    const minh = soat();
    const tinh = hds.filter(h => !h.trung);
    const cong = f => tinh.reduce((s, h) => s + (tienVnd(h, f(h)) || 0), 0);
    const theoSuat = new Map();
    tinh.forEach(h => (h.suat.length ? h.suat : [{ts: h.mau === "2" ? "Không thuế (HĐ bán hàng)" : "Không ghi", tien: h.chuaThue ?? h.tong, thue: 0}]).forEach(s => {
      const o = theoSuat.get(s.ts) || {n: 0, tien: 0, thue: 0};
      o.n++; o.tien += tienVnd(h, s.tien) || 0; o.thue += tienVnd(h, s.thue) || 0; theoSuat.set(s.ts, o);
    }));
    const canXem = hds.filter(h => h.loi.length || h.luuY.length);
    kq = {minh, tinh, theoSuat, canXem, tong: {chua: cong(h => h.chuaThue), thue: cong(h => h.thue), tt: cong(h => h.tong)}};

    $("#hKq").hidden = false;
    $("#hTong").innerHTML = `<div class="o"><b>${tinh.length}</b>hoá đơn${hds.length > tinh.length ? ` (bỏ ${hds.length - tinh.length} trùng)` : ""}</div>
      <div><b>${hien(kq.tong.chua)}</b>tiền chưa thuế</div><div><b>${hien(kq.tong.thue)}</b>tiền thuế GTGT</div><div><b>${hien(kq.tong.tt)}</b>tổng thanh toán</div>
      <div class="${hds.some(h => h.loi.length) ? "l" : canXem.length ? "s" : "o"}"><b>${canXem.length}</b>hoá đơn cần xem lại</div>
      ${hong.length ? `<div class="x"><b>${hong.length}</b>file không đọc được</div>` : ""}`;
    $("#hMinh").innerHTML = minh ? `Máy nhận ra đây là hoá đơn <b>${minh.chieu.toLowerCase()}</b> của MST <span class="pill">${esc(minh.mst)}</span>. Hoá đơn ghi MST khác sẽ bị đánh dấu.` : "";
    veBang($("#hSuat"), ["Thuế suất", "Số hoá đơn", "Tiền chưa thuế", "Tiền thuế"],
      [...kq.theoSuat].map(([ts, o]) => ({o: [ts, o.n, o.tien, o.thue]})).concat([{o: ["Cộng", tinh.length, [...kq.theoSuat.values()].reduce((s, o) => s + o.tien, 0), kq.tong.thue], trCls: "cong"}]));
    $("#hLoc").innerHTML = [["tat", `Tất cả (${hds.length})`], ["xem", `Cần xem lại (${canXem.length})`]].concat(hong.length ? [["hong", `Không đọc được (${hong.length})`]] : [])
      .map(([k, t]) => `<button type="button" data-loc="${k}" class="${loc === k ? "" : "ghost"}">${t}</button>`).join("");
    $("#hLoc").querySelectorAll("button").forEach(b => b.addEventListener("click", () => { loc = b.dataset.loc; chay(); }));
    ve();
  }

  const ghiChu = h => h.loi.concat(h.luuY).join(" · ");
  function ve() {
    if (loc === "hong") { veBang($("#hGrid"), ["File", "Lý do"], hong.map(x => ({o: [x.file, x.ly]}))); return; }
    const ds = loc === "xem" ? kq.canXem : hds;
    veBang($("#hGrid"), ["Ngày lập", "Ký hiệu", "Số HĐ", "Kiểm tra", "Người bán", "MST bán", "Người mua", "MST mua", "Chưa thuế", "Thuế suất", "Tiền thuế", "Tổng thanh toán"],
      ds.map(h => ({
        o: [h.ngay, h.kh, h.so, ghiChu(h) || "Ổn", h.ban.ten, h.ban.mst, h.mua.ten, h.mua.mst, tienVnd(h, h.chuaThue), h.suat.map(s => s.ts).join(", "), tienVnd(h, h.thue), tienVnd(h, h.tong)],
        cls: Object.assign([], {3: h.loi.length ? "loi" : h.luuY.length ? "sua" : ""}),
        tip: Object.assign([], {3: ghiChu(h), 4: h.ban.dc, 6: h.mua.dc}),
        trCls: h.trung ? "xoa" : "",
      })));
  }

  function tai() {
    if (!kq) return;
    const bk = [["STT", "Mẫu số", "Ký hiệu", "Số hoá đơn", "Ngày lập", "Tên người bán", "MST người bán", "Tên người mua", "MST người mua", "Tiền chưa thuế", "Thuế suất", "Tiền thuế GTGT", "Tổng thanh toán", "Loại tiền", "Tỷ giá", "Mã CQT", "Hình thức TT", "Kiểm tra", "File"]];
    hds.forEach((h, i) => bk.push([i + 1, h.mau, h.kh, h.so, h.ngay, h.ban.ten, h.ban.mst, h.mua.ten, h.mua.mst, tienVnd(h, h.chuaThue), h.suat.map(s => s.ts).join(", "), tienVnd(h, h.thue), tienVnd(h, h.tong), h.dvt, h.tg, h.mcqt, h.httt, (h.trung ? "[TRÙNG] " : "") + (ghiChu(h) || "Ổn"), h.file]));
    bk.push(["", "", "", "", "", "Cộng (không tính hoá đơn trùng)", "", "", "", kq.tong.chua, "", kq.tong.thue, kq.tong.tt]);
    const ts = [["Ngày lập", "Ký hiệu", "Số hoá đơn", "Người bán", "MST người bán", "Người mua", "MST người mua", "Thuế suất", "Tiền chưa thuế", "Tiền thuế"]];
    kq.tinh.forEach(h => h.suat.forEach(s => ts.push([h.ngay, h.kh, h.so, h.ban.ten, h.ban.mst, h.mua.ten, h.mua.mst, s.ts, tienVnd(h, s.tien), tienVnd(h, s.thue)])));
    ts.push([], ["Tổng hợp theo thuế suất", "", "", "", "", "", "", "Thuế suất", "Tiền chưa thuế", "Tiền thuế"]);
    kq.theoSuat.forEach((o, k) => ts.push(["", "", "", "", "", "", "", k, o.tien, o.thue]));
    const ct = [["Ký hiệu", "Số hoá đơn", "Ngày lập", "Người bán", "STT", "Mã hàng", "Tên hàng hoá, dịch vụ", "ĐVT", "Số lượng", "Đơn giá", "Chiết khấu", "Thành tiền", "Thuế suất", "Tính chất"]];
    const TCHAT = {1: "Hàng hoá, dịch vụ", 2: "Khuyến mại", 3: "Chiết khấu thương mại", 4: "Ghi chú"};
    kq.tinh.forEach(h => h.dong.forEach(d => ct.push([h.kh, h.so, h.ngay, h.ban.ten, d.stt, d.ma, d.ten, d.dvt, d.sl, d.gia, d.ck, d.tien, d.ts, TCHAT[d.tchat] || d.tchat])));
    const cx = [["Ký hiệu", "Số hoá đơn", "Ngày lập", "Người bán", "Mức", "Nội dung"]];
    kq.canXem.forEach(h => { h.loi.forEach(l => cx.push([h.kh, h.so, h.ngay, h.ban.ten, "Lỗi", l])); h.luuY.forEach(l => cx.push([h.kh, h.so, h.ngay, h.ban.ten, "Lưu ý", l])); });
    const sheets = [["Bảng kê hoá đơn", sheetTu(bk)], ["Theo thuế suất", sheetTu(ts)], ["Chi tiết hàng hoá", sheetTu(ct)], ["Cần xem lại", sheetTu(cx)]];
    if (hong.length) sheets.push(["Không đọc được", sheetTu([["File", "Lý do"], ...hong.map(x => [x.file, x.ly])])]);
    const ngay = kq.tinh.map(h => h.ngay).filter(Boolean), thang = ngay.length ? `${String(ngay[0].getMonth() + 1).padStart(2, "0")}-${ngay[0].getFullYear()}` : "";
    taiWb(sheets, `bang-ke-hoa-don${thang ? "_" + thang : ""}_${kq.tinh.length}-hd.xlsx`);
  }

  // ---------- hoá đơn mẫu: đầu vào của một công ty, cài sẵn vài lỗi hay gặp ----------
  function xmlMau(o) {
    const dong = o.dong.map((d, i) => `<HHDVu><TChat>${d[5] || 1}</TChat><STT>${i + 1}</STT><MHHDVu>${d[0]}</MHHDVu><THHDVu>${d[1]}</THHDVu><DVTinh>${d[2]}</DVTinh><SLuong>${d[3]}</SLuong><DGia>${d[4]}</DGia><ThTien>${d[3] * d[4]}</ThTien><TSuat>${o.ts}</TSuat></HHDVu>`).join("");
    const chua = o.chua ?? o.dong.reduce((s, d) => s + d[3] * d[4], 0), thue = o.thue ?? Math.round(chua * parseFloat(o.ts) / 100), tong = o.tong ?? chua + thue;
    return `<?xml version="1.0" encoding="UTF-8"?><HDon><DLHDon Id="data"><TTChung><PBan>2.0.0</PBan><THDon>Hóa đơn giá trị gia tăng</THDon><KHMSHDon>1</KHMSHDon><KHHDon>${o.kh}</KHHDon><SHDon>${o.so}</SHDon><NLap>${o.ngay}</NLap><DVTTe>VND</DVTTe><TGia>1</TGia><HTTToan>TM/CK</HTTToan>${o.lq ? `<TTHDLQuan><TCHDon>1</TCHDon><LHDCLQuan>1</LHDCLQuan><KHMSHDCLQuan>1</KHMSHDCLQuan><KHHDCLQuan>${o.lq[0]}</KHHDCLQuan><SHDCLQuan>${o.lq[1]}</SHDCLQuan></TTHDLQuan>` : ""}</TTChung>
<NDHDon><NBan><Ten>${o.ban[0]}</Ten><MST>${o.ban[1]}</MST><DChi>${o.ban[2]}</DChi></NBan><NMua><Ten>${o.mua?.[0] || "Công ty TNHH Thương mại Minh Phát"}</Ten><MST>${o.mua?.[1] || "0109876543"}</MST><DChi>Số 12 Trần Duy Hưng, Cầu Giấy, Hà Nội</DChi></NMua>
<DSHHDVu>${dong}</DSHHDVu><TToan><THTTLTSuat><LTSuat><TSuat>${o.ts}</TSuat><ThTien>${chua}</ThTien><TThue>${thue}</TThue></LTSuat></THTTLTSuat><TgTCThue>${chua}</TgTCThue><TgTThue>${thue}</TgTThue><TgTTTBSo>${tong}</TgTTTBSo><TgTTTBChu>${o.chu}</TgTTTBChu></TToan></NDHDon></DLHDon>
${o.khongMa ? "" : "<MCCQT>00" + o.so.padStart(6, "0") + "A1B2C3D4E5F6A7B8C9D0E1F2</MCCQT>"}</HDon>`;
  }
  function mau() {
    const vpp = ["Công ty CP Văn phòng phẩm Hồng Hà", "0100100223", "25 Lý Thường Kiệt, Hoàn Kiếm, Hà Nội"];
    const vt = ["Tổng Công ty Viễn thông Viettel", "0100109106-002", "Lô D26 Khu đô thị mới Cầu Giấy, Hà Nội"];
    const xd = ["Công ty TNHH Xăng dầu Petrolimex Hà Nội", "0100107370", "1 Khâm Thiên, Đống Đa, Hà Nội"];
    const md = ["Công ty CP Điện máy Thành Đạt", "0105556789", "88 Láng Hạ, Đống Đa, Hà Nội"];
    const ds = [
      ["HD_C26THH_118.xml", {kh: "C26THH", so: "118", ngay: "2026-10-03", ban: vpp, ts: "8%", dong: [["GA4", "Giấy A4 Double A 70gsm", "Ram", 20, 72000], ["BB01", "Bút bi Thiên Long TL-027", "Hộp", 5, 65000]], chu: "Một triệu chín trăm linh sáu nghìn hai trăm đồng"}],
      ["HD_C26TVT_2045.xml", {kh: "C26TVT", so: "2045", ngay: "2026-10-05", ban: vt, ts: "10%", dong: [["CUOC", "Cước Internet cáp quang tháng 10/2026", "Tháng", 1, 1200000]], chu: "Một triệu ba trăm hai mươi nghìn đồng chẵn"}],
      ["HD_C26TXD_7781.xml", {kh: "C26TXD", so: "7781", ngay: "2026-10-09", ban: xd, ts: "10%", dong: [["RON95", "Xăng RON 95-V", "Lít", 120, 23500]], chu: "Ba triệu một trăm hai nghìn đồng"}],
      ["HD_C26TXD_7781 (1).xml", {kh: "C26TXD", so: "7781", ngay: "2026-10-09", ban: xd, ts: "10%", dong: [["RON95", "Xăng RON 95-V", "Lít", 120, 23500]], chu: "Ba triệu một trăm hai nghìn đồng"}],
      ["HD_C26TMD_356.xml", {kh: "C26TMD", so: "356", ngay: "2026-10-14", ban: md, ts: "10%", dong: [["DH12", "Điều hoà Daikin 12000BTU FTKB35", "Bộ", 2, 11500000], ["LD", "Công lắp đặt", "Bộ", 2, 500000]], tong: 26200000, chu: "Hai mươi sáu triệu hai trăm nghìn đồng"}],
      ["HD_C26THH_131.xml", {kh: "C26THH", so: "131", ngay: "2026-10-20", ban: vpp, ts: "8%", dong: [["MUC", "Mực in HP 107A", "Hộp", 3, 450000]], chu: "Một triệu bốn trăm năm mươi tám nghìn đồng", khongMa: true}],
      ["HD_C26TMD_360.xml", {kh: "C26TMD", so: "360", ngay: "2026-10-22", ban: md, ts: "10%", mua: ["Hộ kinh doanh Nguyễn Văn Hùng", "8012345678"], dong: [["TL", "Tủ lạnh Samsung RT35", "Chiếc", 1, 8900000]], chu: "Chín triệu bảy trăm chín mươi nghìn đồng"}],
      ["HD_C26TVT_2190.xml", {kh: "C26TVT", so: "2190", ngay: "2026-10-28", ban: vt, ts: "10%", dong: [["SIM", "Cước di động trả sau tháng 10/2026", "Tháng", 1, 850000]], chu: "Chín trăm ba mươi nghìn năm trăm đồng"}],
    ];
    const enc = new TextEncoder();
    return ds.map(([ten, o]) => ({ten, buf: enc.encode(xmlMau(o))}));
  }

  ganDrop($("#hDrop"), $("#hFile"), them);
  $("#hMau").addEventListener("click", () => { hds = []; hong = []; loc = "tat"; them(mau()); });
  $("#hTai").addEventListener("click", tai);
  window._hoaDon = {get kq() { return kq; }, get hds() { return hds; }, docChu, tai};
})();
