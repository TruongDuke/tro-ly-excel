/* Lõi làm sạch dữ liệu tiếng Việt: hàm thuần, chạy được cả trong trình duyệt lẫn Node (để kiểm thử). */
(function (goc) {
  "use strict";
  const boDau = s => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D").toLowerCase();
  const rong = v => v === null || v === undefined || String(v).trim() === "";

  // ---------- họ tên ----------
  function hoTen(v) {
    if (rong(v)) return {v: ""};
    let s = String(v).normalize("NFC").replace(/[ \s]+/g, " ").trim();
    s = s.replace(/[.,;:_]+$/g, "").trim();
    if (/\d/.test(s)) return {v: s, loi: "Họ tên có chữ số"};
    s = s.toLocaleLowerCase("vi").split(" ").map(w => w.split("-").map(p => p ? p[0].toLocaleUpperCase("vi") + p.slice(1) : p).join("-")).join(" ");
    return {v: s};
  }
  function tachTen(s) {
    const p = String(s || "").trim().split(/\s+/).filter(Boolean);
    if (p.length === 0) return ["", "", ""];
    if (p.length === 1) return ["", "", p[0]];
    return [p[0], p.slice(1, -1).join(" "), p[p.length - 1]];
  }

  // ---------- số điện thoại ----------
  function sdt(v) {
    if (rong(v)) return {v: ""};
    const goc = String(v).trim();
    let d = goc.replace(/[^\d+]/g, "");
    if (d.startsWith("+84")) d = "0" + d.slice(3);
    else if (d.startsWith("0084")) d = "0" + d.slice(4);
    else if (/^84\d{9}$/.test(d)) d = "0" + d.slice(2);
    d = d.replace(/\+/g, "");
    if (/^[35789]\d{8}$/.test(d)) d = "0" + d;          // mất số 0 đầu khi Excel coi là số
    if (/^0[35789]\d{8}$/.test(d)) return {v: d};        // di động 10 số
    if (/^02\d{9}$/.test(d)) return {v: d};              // cố định 11 số
    return {v: goc, loi: "Số điện thoại không đúng 10 số"};
  }

  // ---------- ngày tháng (người Việt viết ngày trước tháng) ----------
  const hopLe = (y, m, d) => { if (m < 1 || m > 12 || d < 1 || y < 1900 || y > 2100) return false;
    const t = new Date(Date.UTC(y, m - 1, d)); return t.getUTCMonth() === m - 1 && t.getUTCDate() === d; };
  function ngay(v) {
    if (rong(v)) return {v: ""};
    if (v instanceof Date && !isNaN(v)) return {v: {y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate()}};
    if (typeof v === "number" && v > 20000 && v < 80000) {          // số ngày kiểu Excel
      const t = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5);
      return {v: {y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate()}};
    }
    let s = boDau(String(v)).trim().replace(/^(ngay|thu \w+,?)\s*/, "");
    s = s.replace(/\s*thang\s*/g, "/").replace(/\s*nam\s*/g, "/").replace(/\s+/g, "/");
    const p = s.split(/[\/\-.]+/).filter(Boolean).map(x => x.replace(/\D/g, ""));
    if (p.length !== 3 || p.some(x => !x)) return {v: String(v), loi: "Không đọc được ngày"};
    let [a, b, c] = p.map(Number), y, m, d, ghi;
    if (p[0].length === 4) { y = a; m = b; d = c; }                 // 2024-03-15
    else { d = a; m = b; y = c; if (p[2].length <= 2) y = y < 50 ? 2000 + y : 1900 + y; }
    if (!hopLe(y, m, d) && hopLe(y, d, m)) { [d, m] = [m, d]; ghi = "Đã đảo ngày/tháng (có thể nhập theo kiểu Mỹ)"; }
    if (!hopLe(y, m, d)) return {v: String(v), loi: "Ngày không hợp lệ"};
    return {v: {y, m, d}, loi: ghi};
  }
  const ngayChu = o => o && typeof o === "object" ? `${String(o.d).padStart(2, "0")}/${String(o.m).padStart(2, "0")}/${o.y}` : String(o ?? "");

  // ---------- số tiền ----------
  function tien(v) {
    if (rong(v)) return {v: ""};
    if (typeof v === "number") return {v: v};
    const goc = String(v);
    let s = boDau(goc).replace(/\s+/g, "").replace(/(vnd|dong|d|₫)$/,"").replace(/₫/g, "");
    let m;
    if ((m = s.match(/^(-?\d+(?:[.,]\d+)?)(ty)(\d{0,3})$/))) return {v: Math.round(parseFloat(m[1].replace(",", ".")) * 1e9 + (m[3] ? +m[3].padEnd(3, "0") * 1e6 : 0))};
    if ((m = s.match(/^(-?\d+(?:[.,]\d+)?)(tr|trieu|m)(\d{0,3})$/))) return {v: Math.round(parseFloat(m[1].replace(",", ".")) * 1e6 + (m[3] ? +m[3].padEnd(3, "0") * 1e3 : 0))};
    if ((m = s.match(/^(-?\d+(?:[.,]\d+)?)(k|nghin|ngan)$/))) return {v: Math.round(parseFloat(m[1].replace(",", ".")) * 1e3)};
    if (/^-?\d{1,3}([.,]\d{3})+$/.test(s)) return {v: +s.replace(/[.,]/g, "")};
    if (/^-?\d{1,3}(\.\d{3})+,\d+$/.test(s)) return {v: +s.replace(/\./g, "").replace(",", ".")};
    if (/^-?\d+([.,]\d+)?$/.test(s)) return {v: +s.replace(",", ".")};
    return {v: goc, loi: "Không đọc được số tiền"};
  }

  // ---------- email ----------
  function email(v) {
    if (rong(v)) return {v: ""};
    const s = String(v).trim().toLowerCase().replace(/\s+/g, "");
    return /^[^@\s]+@[^@\s]+\.[a-z]{2,}$/.test(s) ? {v: s} : {v: s, loi: "Email sai định dạng"};
  }
  function chu(v) { if (rong(v)) return {v: ""}; return {v: String(v).normalize("NFC").replace(/[ \s]+/g, " ").trim()}; }

  const LOAI = {
    hoten: {ten: "Họ tên", lam: hoTen}, sdt: {ten: "Số điện thoại", lam: sdt}, ngay: {ten: "Ngày tháng", lam: ngay},
    tien: {ten: "Số tiền / số", lam: tien}, email: {ten: "Email", lam: email}, chu: {ten: "Chữ thường (chỉ bỏ khoảng trắng thừa)", lam: chu}, giu: {ten: "Giữ nguyên", lam: v => ({v})},
  };

  // đoán loại cột theo tiêu đề rồi theo nội dung
  function doanLoai(tieuDe, giaTri) {
    const t = boDau(tieuDe || "");
    if (/(ho ?ten|ho va ten|ten khach|ten nhan vien|full ?name|^ten$|name)/.test(t)) return "hoten";
    if (/(sdt|so dt|dien thoai|phone|mobile|di dong|zalo)/.test(t)) return "sdt";
    if (/(email|e-mail|mail)/.test(t)) return "email";
    if (/(ngay|date|sinh|dob)/.test(t)) return "ngay";
    if (/(tien|luong|doanh|gia|thu nhap|so du|cong no|amount|price|tong)/.test(t)) return "tien";
    const mau = giaTri.filter(v => !rong(v)).slice(0, 40);
    if (!mau.length) return "chu";
    const ti = f => mau.filter(f).length / mau.length;
    if (ti(v => !sdt(v).loi) > .6) return "sdt";
    if (ti(v => !email(v).loi) > .6) return "email";
    if (ti(v => typeof v === "string" && /[\/\-.]/.test(v) && !ngay(v).loi && ngay(v).v) > .6) return "ngay";
    if (ti(v => typeof v !== "number" && !tien(v).loi) > .6) return "tien";
    return "chu";
  }

  // chạy làm sạch: rows = mảng mảng (dòng 0 là tiêu đề); loai = mảng mã loại từng cột
  function lamSach(rows, loai, opt) {
    const hd = rows[0].map(x => String(x ?? "").trim() || "(không tên)");
    const out = [], ds = [], log = [];
    let soSua = 0, soLoi = 0, soTrong = 0, soTrung = 0;
    const daThay = new Map();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const dong = {goc: i + 1, o: [], xoa: null};
      if (r.every(rong)) { if (opt.xoaTrong) { dong.xoa = "Dòng trống"; soTrong++; } }
      hd.forEach((_, j) => {
        let v0 = r[j];
        if (opt.nfc && typeof v0 === "string") v0 = v0.normalize("NFC");
        const kq = LOAI[loai[j]].lam(v0 === undefined ? "" : v0);
        const truoc = rong(r[j]) ? "" : (r[j] instanceof Date ? ngayChu({y: r[j].getFullYear(), m: r[j].getMonth() + 1, d: r[j].getDate()}) : String(r[j]));
        const sau = loai[j] === "ngay" ? ngayChu(kq.v) : (loai[j] === "tien" && typeof kq.v === "number" ? kq.v.toLocaleString("vi-VN") : String(kq.v ?? ""));
        const doi = loai[j] !== "giu" && truoc !== sau && !(loai[j] === "tien" && String(r[j]) === String(kq.v));
        const o = {v: kq.v, hien: sau, cu: truoc, doi, loi: kq.loi && !(kq.loi.startsWith("Đã đảo") ) ? kq.loi : null, ghi: kq.loi && kq.loi.startsWith("Đã đảo") ? kq.loi : null};
        if (o.ghi) o.doi = true;
        dong.o.push(o);
      });
      if (!dong.xoa && opt.xoaTrung) {
        const k = opt.khoa === -1 ? dong.o.map(o => boDau(o.hien)).join("|") : boDau(dong.o[opt.khoa]?.hien || "");
        if (k.replace(/\|/g, "")) { if (daThay.has(k)) { dong.xoa = `Trùng với dòng ${daThay.get(k)}`; soTrung++; } else daThay.set(k, dong.goc); }
      }
      if (!dong.xoa) dong.o.forEach((o, j) => { if (o.doi) { soSua++; log.push([dong.goc, hd[j], o.cu, o.hien, o.ghi || ""]); } if (o.loi) { soLoi++; log.push([dong.goc, hd[j], o.cu, o.cu, "CẦN KIỂM TRA: " + o.loi]); } });
      else log.push([dong.goc, "(cả dòng)", "", "", "Đã xoá: " + dong.xoa]);
      ds.push(dong);
    }
    return {hd, ds, log, soSua, soLoi, soTrong, soTrung, tong: rows.length - 1};
  }

  goc.LamSach = {hoTen, tachTen, sdt, ngay, ngayChu, tien, email, doanLoai, lamSach, LOAI, boDau};
})(typeof window !== "undefined" ? window : globalThis);
