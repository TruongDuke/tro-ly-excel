/* Dán chữ từ ChatGPT/Claude -> file Word chuẩn trình bày Việt Nam (Times New Roman, lề NĐ30, đề mục đánh số, mục lục) */
(function () {
  "use strict";
  const {$, esc} = window.TL;

  // ---------- 1. đọc nguồn thành danh sách khối ----------
  // khối: {t:"h",lv,runs} {t:"p",runs} {t:"li",so,lv,runs} {t:"bang",rows:[[runs]],dauBang} {t:"code",text} {t:"quote",runs} {t:"hr"}
  // run: {text,b,i,code}
  function tuMarkdown(md) {
    const html = window.marked.parse(md, {gfm: true, breaks: false});
    return tuHtml(html);
  }
  function tuHtml(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const ds = [];
    duyet(doc.body, ds, 0);
    return ds;
  }
  const KHOI = /^(P|H[1-6]|UL|OL|LI|TABLE|PRE|BLOCKQUOTE|HR|DIV|SECTION|ARTICLE|MAIN|HEADER|FOOTER)$/;
  function runsCua(node, kieu = {}, out = []) {
    node.childNodes.forEach(n => {
      if (n.nodeType === 3) { if (n.nodeValue) out.push({text: n.nodeValue, ...kieu}); return; }
      if (n.nodeType !== 1) return;
      const tg = n.tagName;
      if (tg === "BR") { out.push({text: "\n", ...kieu}); return; }
      if (tg === "UL" || tg === "OL" || tg === "TABLE" || tg === "PRE") return;   // xử lý riêng như khối con
      const k = {...kieu};
      if (tg === "STRONG" || tg === "B") k.b = true;
      if (tg === "EM" || tg === "I") k.i = true;
      if (tg === "CODE") k.code = true;
      const fw = n.style && n.style.fontWeight; if (fw === "bold" || +fw >= 600) k.b = true;
      runsCua(n, k, out);
    });
    return out;
  }
  function duyet(el, ds, lv) {
    let tho = [];   // chữ rời nằm thẳng trong div
    const xa = () => { if (tho.some(r => r.text.trim())) ds.push({t: "p", runs: tho}); tho = []; };
    el.childNodes.forEach(n => {
      if (n.nodeType === 3) { if (n.nodeValue.trim()) tho.push({text: n.nodeValue}); return; }
      if (n.nodeType !== 1) return;
      const tg = n.tagName;
      if (!KHOI.test(tg)) { tho.push(...runsCua({childNodes: [n]})); return; }
      xa();
      if (/^H[1-6]$/.test(tg)) ds.push({t: "h", lv: +tg[1], runs: runsCua(n)});
      else if (tg === "P") ds.push({t: "p", runs: runsCua(n)});
      else if (tg === "UL" || tg === "OL") danhSach(n, ds, lv, tg === "OL");
      else if (tg === "TABLE") {
        const rows = [...n.querySelectorAll("tr")].map(tr => [...tr.children].map(td => runsCua(td)));
        const dauBang = !!n.querySelector("thead") || !!n.querySelector("tr > th");
        if (rows.length) ds.push({t: "bang", rows, dauBang});
      }
      else if (tg === "PRE") ds.push({t: "code", text: n.textContent.replace(/\n$/, "")});
      else if (tg === "BLOCKQUOTE") { const con = []; duyet(n, con, lv); con.forEach(c => { if (c.t === "p") c.t = "quote"; ds.push(c); }); }
      else if (tg === "HR") ds.push({t: "hr"});
      else duyet(n, ds, lv);
    });
    xa();
  }
  function danhSach(ul, ds, lv, so) {
    [...ul.children].filter(li => li.tagName === "LI").forEach(li => {
      // đoạn <p> bên trong li (danh sách "rời") gộp thành một dòng
      const runs = runsCua(li).filter((r, i, a) => !(i === a.length - 1 && !r.text.trim()));
      ds.push({t: "li", so, lv, runs: catKhoangTrang(runs)});
      li.querySelectorAll(":scope > ul, :scope > ol, :scope > p > ul, :scope > p > ol").forEach(c => danhSach(c, ds, lv + 1, c.tagName === "OL"));
      li.querySelectorAll(":scope > table, :scope > pre").forEach(c => duyet({childNodes: [c]}, ds, lv));
    });
  }
  function catKhoangTrang(runs) {
    runs = runs.map(r => ({...r, text: r.text.replace(/\s*\n\s*/g, r.text.includes("\n\n") ? "\n" : " ")}));
    if (runs.length) { runs[0].text = runs[0].text.replace(/^\s+/, ""); runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, ""); }
    return runs.filter(r => r.text);
  }

  // ---------- 2. dọn và chuẩn hoá ----------
  const EMOJI = /(?:\p{Extended_Pictographic}|\p{Regional_Indicator})(?:️|‍(?:\p{Extended_Pictographic}))*️?|[☀-➿]️?/gu;
  const MO_AI = /^(chắc chắn|tất nhiên|dĩ nhiên|tuyệt vời|rất hay|câu hỏi hay|được rồi|ok[,! ]|vâng|dạ[,! ]|sure|certainly|of course|great question|here (is|are)|dưới đây là|sau đây là|mình (sẽ )?(gửi|giúp|soạn|viết)|tôi (sẽ )?(gửi|giúp|soạn|viết)|em (xin )?(gửi|soạn))/i;
  const KET_AI = /(bạn có muốn|anh\/chị có muốn|bạn cần (mình|tôi)|nếu (bạn|anh\/chị) (cần|muốn)|hy vọng|hi vọng|hãy cho (mình|tôi|em) biết|mình có thể (giúp|làm|soạn)|tôi có thể (giúp|làm|soạn)|chúc (bạn|anh|chị)|would you like|let me know|i hope|feel free)/i;
  const chu = runs => runs.map(r => r.text).join("");
  function suaChu(s, dem) {
    const t0 = s;
    s = s.normalize("NFC")
      .replace(/\*\*|__(?=\S)|(?<=\S)__/g, "")                       // dấu ** sót lại
      .replace(/^#{1,6}\s+/gm, "")
      .replace(/[ \t ]{2,}/g, " ")
      .replace(/\s+([,.;:!?%)\]])/g, "$1")                         // cách trước dấu câu
      .replace(/([(\[])\s+/g, "$1")
      .replace(/([,;:!?])(?=[A-Za-zÀ-ỹ])/g, "$1 ")                  // thiếu cách sau dấu câu
      .replace(/\.(?=[A-ZÀ-Ỹ][a-zà-ỹ])/g, ". ");
    if (s !== t0) dem.dauCau++;
    return s;
  }
  function suaRuns(runs, o, dem) {
    return runs.map(r => {
      let t = r.text;
      if (o.boEmoji) { const n = (t.match(EMOJI) || []).length; if (n) { dem.emoji += n; t = t.replace(EMOJI, "").replace(/^\s+/, m => m.includes("\n") ? m : ""); } }
      if (!r.code) t = suaChu(t, dem);
      return {...r, text: t};
    }).filter(r => r.text !== "");
  }
  function catDau(runs) {   // bỏ khoảng trắng đầu/cuối cả đoạn
    runs = runs.slice();
    while (runs.length && !runs[0].text.trim()) runs.shift();
    while (runs.length && !runs[runs.length - 1].text.trim()) runs.pop();
    if (runs.length) { runs[0] = {...runs[0], text: runs[0].text.replace(/^\s+/, "")}; const k = runs.length - 1; runs[k] = {...runs[k], text: runs[k].text.replace(/\s+$/, "")}; }
    return runs;
  }
  const LA_MA = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX"];
  const SO_SAN = /^\s*((phần|chương|mục|bước|phase|part)\s+[\dIVXLC]+\s*[:.\-–]?|[IVXLC]+[.)]|\d+(\.\d+)*[.)]?|[a-zđ][.)])\s+/i;

  function chuanHoa(khoi, o) {
    const dem = {sao: 0, thang: 0, emoji: 0, dauCau: 0, cauAi: 0, deMuc: 0, bang: 0, ds: 0};
    let ds = khoi.map(k => {
      const c = {...k};
      if (c.runs) c.runs = catDau(suaRuns(c.runs, o, dem));
      if (c.rows) c.rows = c.rows.map(r => r.map(cell => catDau(suaRuns(cell, o, dem))));
      return c;
    }).filter(k => k.t === "hr" || k.t === "bang" || k.t === "code" || chu(k.runs).trim());

    // câu mở đầu / kết thúc kiểu chatbot
    if (o.boCauAi) {
      ds.forEach((k, i) => {
        if (k.t !== "p") return;
        const s = chu(k.runs).trim();
        if (i <= 1 && MO_AI.test(s) && s.length < 260) { k.bo = "Câu mở đầu của AI"; dem.cauAi++; }
        else if (i >= ds.length - 3 && KET_AI.test(s) && s.length < 400) { k.bo = "Câu hỏi lại của AI"; dem.cauAi++; }
      });
      // đường kẻ ngang ngay trước câu kết bị bỏ
      ds.forEach((k, i) => { if (k.t === "hr" && (i === ds.length - 1 || ds.slice(i + 1).every(x => x.bo) || ds.slice(0, i).every(x => x.bo))) k.bo = "Đường kẻ thừa"; });
    }
    const con = ds.filter(k => !k.bo && k.t !== "hr");

    // tiêu đề văn bản: H1 duy nhất ở đầu, hoặc dòng đầu toàn chữ đậm ngắn
    const h = con.filter(k => k.t === "h");
    let tieuDe = null;
    if (con[0] && con[0].t === "h" && h.filter(k => k.lv === con[0].lv).length === 1) tieuDe = con[0];
    else if (con[0] && con[0].t === "p" && con[0].runs.every(r => r.b || !r.text.trim()) && chu(con[0].runs).length < 140) tieuDe = con[0];
    if (tieuDe) tieuDe.tieuDe = true;

    // dồn cấp đề mục về 1, 2, 3...
    const cap = [...new Set(h.filter(k => !k.tieuDe).map(k => k.lv))].sort((a, b) => a - b);
    ds.forEach(k => { if (k.t === "h" && !k.tieuDe) k.cap = Math.min(cap.indexOf(k.lv) + 1, 4); });

    // đánh số đề mục kiểu Việt: I. / 1. / 1.1. / a)
    const deMuc = ds.filter(k => k.t === "h" && !k.tieuDe && !k.bo);
    const daCoSo = deMuc.filter(k => SO_SAN.test(chu(k.runs))).length;
    const coSo = deMuc.length && daCoSo >= deMuc.length / 2;
    if (o.danhSo) {
      const dang = [0, 0, 0, 0, 0];
      deMuc.forEach(k => {
        if (coSo) k.runs = boSoCu(k.runs);
        dang[k.cap]++; for (let j = k.cap + 1; j < dang.length; j++) dang[j] = 0;
        const so = k.cap === 1 ? LA_MA[dang[1]] + ". " : k.cap === 2 ? dang[2] + ". " : k.cap === 3 ? `${dang[2]}.${dang[3]}. ` : String.fromCharCode(96 + dang[4]) + ") ";
        k.so = so; dem.deMuc++;
      });
    }
    ds.forEach(k => { if (k.t === "bang" && !k.bo) dem.bang++; if (k.t === "li" && !k.bo) dem.ds++; });
    return {ds, tieuDe, dem};
  }
  function boSoCu(runs) {
    runs = runs.map(r => ({...r}));
    const r0 = runs.find(r => r.text.trim()); if (!r0) return runs;
    r0.text = r0.text.replace(SO_SAN, "");
    return runs;
  }
  // đếm dấu markdown thừa trên chữ gốc (để cho người dùng thấy đã dọn bao nhiêu)
  function demGoc(md) {
    return {
      sao: (md.match(/\*\*/g) || []).length,
      thang: (md.match(/^#{1,6}\s/gm) || []).length,
      gach: (md.match(/^\s*\|?\s*:?-{3,}/gm) || []).length,
    };
  }

  // ---------- 3. xem trước như trang Word ----------
  function runsHtml(runs) {
    return runs.map(r => {
      let t = esc(r.text).replace(/\n/g, "<br>");
      if (r.code) t = `<code>${t}</code>`;
      if (r.i) t = `<i>${t}</i>`;
      if (r.b) t = `<b>${t}</b>`;
      return t;
    }).join("");
  }
  function xemTruoc(kq, o) {
    let h = "", ol = {}, ml = "";
    if (o.mucLuc && kq.ds.some(k => k.t === "h" && !k.tieuDe && !k.bo)) {
      const muc = kq.ds.filter(k => k.t === "h" && !k.tieuDe && !k.bo && k.cap <= 3);
      ml = `<div class="w-ml"><div class="w-mlt">MỤC LỤC</div>${muc.map(k => `<div class="w-mli c${k.cap}"><span>${esc(k.so || "")}${esc(chu(k.runs))}</span><i></i><span class="w-trg">…</span></div>`).join("")}</div>`;
      if (!kq.tieuDe) h += ml;
    }
    kq.ds.forEach((k, i) => {
      if (k.t !== "li") ol = {};
      if (k.bo) { h += `<div class="w-bo" title="${esc(k.bo)} – đã bỏ"><span>${esc(k.bo)}</span>${k.t === "hr" ? "<hr>" : runsHtml(k.runs || [])}</div>`; return; }
      if (k.tieuDe) { h += `<p class="w-td">${runsHtml(k.runs)}</p>` + ml; return; }
      if (k.t === "h") h += `<p class="w-h c${k.cap}">${k.so ? `<span class="w-so">${esc(k.so)}</span>` : ""}${runsHtml(k.runs)}</p>`;
      else if (k.t === "p") h += `<p class="w-p">${runsHtml(k.runs)}</p>`;
      else if (k.t === "quote") h += `<p class="w-p w-q">${runsHtml(k.runs)}</p>`;
      else if (k.t === "li") {
        const key = k.lv; ol[key] = k.so ? (ol[key] || 0) + 1 : 0; Object.keys(ol).forEach(x => { if (+x > key) delete ol[x]; });
        const dau = k.so ? (k.lv ? String.fromCharCode(96 + ol[key]) + ")" : ol[key] + ".") : (k.lv ? "+" : "-");
        h += `<p class="w-li" style="--lv:${k.lv}"><span class="w-dau">${dau}</span>${runsHtml(k.runs)}</p>`;
      }
      else if (k.t === "bang") h += `<div class="w-bangwrap"><table class="w-bang">${k.rows.map((r, ri) => `<tr>${r.map(c => ri === 0 && k.dauBang ? `<th>${runsHtml(c)}</th>` : `<td${laSo(chu(c)) ? ' class="n"' : ""}>${runsHtml(c)}</td>`).join("")}</tr>`).join("")}</table></div>`;
      else if (k.t === "code") h += `<pre class="w-code">${esc(k.text)}</pre>`;
      else if (k.t === "hr") h += `<hr class="w-hr">`;
    });
    return h;
  }
  const laSo = s => /^[-+]?[\d.,\s]+(%|đ|₫|vnđ|triệu|tỷ)?$/i.test(s.trim()) && /\d/.test(s);

  // ---------- 4. xuất .docx ----------
  const CM = 567;   // 1 cm = 567 twip
  function taoDocx(kq, o) {
    const D = window.docx;
    const co = o.coChu * 2;   // half-point
    const TNR = "Times New Roman";
    const tr = (r, them = {}) => new D.TextRun({text: r.text.replace(/\n/g, " "), bold: r.b || them.bold, italics: r.i || them.italics, font: r.code ? "Courier New" : TNR, size: r.code ? co - 4 : them.size, allCaps: them.allCaps,
      ...(r.text.includes("\n") ? {} : {})});
    // giữ xuống dòng mềm bên trong đoạn
    const runsTo = (runs, them) => {
      const out = [];
      runs.forEach(r => r.text.split("\n").forEach((part, j) => { if (j) out.push(new D.TextRun({break: 1})); if (part) out.push(tr({...r, text: part}, them)); }));
      return out;
    };
    const vien = {style: D.BorderStyle.SINGLE, size: 4, color: "000000"};
    const children = [];
    if (kq.tieuDe) children.push(new D.Paragraph({alignment: D.AlignmentType.CENTER, indent: {firstLine: 0}, spacing: {after: 240}, keepNext: true, children: runsTo(kq.tieuDe.runs, {bold: true, allCaps: true, size: co + 2})}));
    if (o.mucLuc && kq.ds.some(k => k.t === "h" && !k.tieuDe && !k.bo)) {
      children.push(new D.Paragraph({alignment: D.AlignmentType.CENTER, indent: {firstLine: 0}, spacing: {before: 120, after: 120}, children: [new D.TextRun({text: "MỤC LỤC", bold: true, font: TNR, size: co})]}));
      children.push(new D.TableOfContents("Mục lục", {hyperlink: true, headingStyleRange: "1-3"}));
      children.push(new D.Paragraph({children: [new D.PageBreak()]}));
    }
    let olId = 0, dangOl = false;
    kq.ds.forEach(k => {
      if (k.bo || k.tieuDe) return;
      if (k.t !== "li") dangOl = false;
      if (k.t === "h") {
        const lv = [null, D.HeadingLevel.HEADING_1, D.HeadingLevel.HEADING_2, D.HeadingLevel.HEADING_3, D.HeadingLevel.HEADING_4][k.cap];
        const ch = [];
        if (k.so) ch.push(new D.TextRun({text: k.so, font: TNR}));
        ch.push(...runsTo(k.runs));
        children.push(new D.Paragraph({heading: lv, children: ch}));
      } else if (k.t === "p" || k.t === "quote") {
        children.push(new D.Paragraph({style: k.t === "quote" ? "TrichDan" : undefined, children: runsTo(k.runs)}));
      } else if (k.t === "li") {
        if (k.so && !dangOl && k.lv === 0) { olId++; dangOl = true; }
        children.push(new D.Paragraph(k.so ? {numbering: {reference: "so", level: Math.min(k.lv, 2), instance: olId}, style: "DongDS", children: runsTo(k.runs)}
          : {numbering: {reference: "gach", level: Math.min(k.lv, 2)}, style: "DongDS", children: runsTo(k.runs)}));
      } else if (k.t === "bang") {
        const n = Math.max(...k.rows.map(r => r.length));
        // cột rộng theo độ dài chữ (giới hạn để cột ngắn không quá hẹp)
        const dai = Array.from({length: n}, (_, j) => Math.min(40, Math.max(6, ...k.rows.map(r => chu(r[j] || []).length))));
        const tong = dai.reduce((a, b) => a + b, 0), RONG = 9355;
        const cw = dai.map(d => Math.round(RONG * d / tong));
        children.push(new D.Table({
          width: {size: RONG, type: D.WidthType.DXA}, columnWidths: cw, layout: D.TableLayoutType.FIXED,
          borders: {top: vien, bottom: vien, left: vien, right: vien, insideHorizontal: vien, insideVertical: vien},
          rows: k.rows.map((r, ri) => new D.TableRow({tableHeader: ri === 0 && k.dauBang, cantSplit: true,
            children: Array.from({length: n}, (_, j) => {
              const c = r[j] || [], dau = ri === 0 && k.dauBang;
              return new D.TableCell({width: {size: cw[j], type: D.WidthType.DXA},
                shading: dau ? {type: D.ShadingType.CLEAR, fill: "E7E6E6", color: "auto"} : undefined,
                margins: {top: 60, bottom: 60, left: 100, right: 100}, verticalAlign: D.VerticalAlign.CENTER,
                children: [new D.Paragraph({style: "TrongBang", alignment: dau ? D.AlignmentType.CENTER : laSo(chu(c)) ? D.AlignmentType.RIGHT : D.AlignmentType.LEFT,
                  children: runsTo(c, {bold: dau || undefined, size: co - 2})})],
              });
            })}))
        }));
        children.push(new D.Paragraph({style: "TrongBang", children: []}));
      } else if (k.t === "code") {
        k.text.split("\n").forEach(line => children.push(new D.Paragraph({style: "Ma", children: [new D.TextRun({text: line || " ", font: "Courier New", size: co - 4})]})));
      }
    });

    const thuong = {font: TNR, size: co, color: "000000"};
    const hd = (id, name, run, para) => ({id, name, basedOn: "Normal", next: "Normal", quickFormat: true, run: {...thuong, ...run}, paragraph: {keepNext: true, keepLines: true, ...para}});
    const doc = new D.Document({
      creator: "Trợ lý văn phòng", title: kq.tieuDe ? chu(kq.tieuDe.runs) : "Văn bản",
      features: {updateFields: o.mucLuc},
      styles: {
        default: {document: {run: thuong, paragraph: {spacing: {line: Math.round(240 * o.gianDong), before: 120, after: 0}}}},
        paragraphStyles: [
          {id: "Normal", name: "Normal", run: thuong, paragraph: {alignment: D.AlignmentType.JUSTIFIED, indent: {firstLine: Math.round(1 * CM)}, spacing: {line: Math.round(240 * o.gianDong), before: 120, after: 0}}},
          hd("Heading1", "Heading 1", {bold: true, allCaps: true}, {spacing: {before: 240, after: 120}, indent: {firstLine: 0}, alignment: D.AlignmentType.LEFT}),
          hd("Heading2", "Heading 2", {bold: true}, {spacing: {before: 180, after: 60}, indent: {firstLine: 0}, alignment: D.AlignmentType.LEFT}),
          hd("Heading3", "Heading 3", {bold: true, italics: true}, {spacing: {before: 120, after: 60}, indent: {firstLine: 0}, alignment: D.AlignmentType.LEFT}),
          hd("Heading4", "Heading 4", {italics: true}, {spacing: {before: 120}, indent: {firstLine: 0}}),
          {id: "DongDS", name: "Dong danh sach", basedOn: "Normal", run: thuong, paragraph: {alignment: D.AlignmentType.JUSTIFIED, spacing: {before: 60}}},
          {id: "TrongBang", name: "Trong bang", basedOn: "Normal", run: {...thuong, size: co - 2}, paragraph: {alignment: D.AlignmentType.LEFT, indent: {firstLine: 0}, spacing: {line: 240, before: 0, after: 0}}},
          {id: "TrichDan", name: "Trich dan", basedOn: "Normal", run: {...thuong, italics: true}, paragraph: {indent: {left: Math.round(1 * CM), firstLine: 0}}},
          {id: "Ma", name: "Ma lenh", basedOn: "Normal", run: {font: "Courier New", size: co - 4}, paragraph: {alignment: D.AlignmentType.LEFT, indent: {left: Math.round(0.5 * CM), firstLine: 0}, spacing: {line: 240, before: 0}, shading: {type: D.ShadingType.CLEAR, fill: "F2F2F2", color: "auto"}}},
        ],
      },
      numbering: {config: [
        {reference: "gach", levels: ["-", "+", "•"].map((t, i) => ({level: i, format: D.LevelFormat.BULLET, text: t, alignment: D.AlignmentType.LEFT,
          style: {run: {font: TNR}, paragraph: {indent: {left: Math.round((1 + i * 0.6) * CM), hanging: Math.round(0.5 * CM)}}}}))},
        {reference: "so", levels: [[D.LevelFormat.DECIMAL, "%1."], [D.LevelFormat.LOWER_LETTER, "%2)"], [D.LevelFormat.LOWER_ROMAN, "%3."]].map(([f, t], i) => ({level: i, format: f, text: t, alignment: D.AlignmentType.LEFT,
          style: {run: {font: TNR}, paragraph: {indent: {left: Math.round((1.2 + i * 0.6) * CM), hanging: Math.round(0.6 * CM)}}}}))},
      ]},
      sections: [{
        properties: {
          page: {size: {width: 11906, height: 16838}, margin: {top: 2 * CM, bottom: 2 * CM, left: 3 * CM, right: Math.round(1.5 * CM), header: Math.round(1 * CM), footer: Math.round(1 * CM)}},
          titlePage: true,
        },
        headers: {
          default: new D.Header({children: [new D.Paragraph({alignment: D.AlignmentType.CENTER, indent: {firstLine: 0}, children: [new D.TextRun({children: [D.PageNumber.CURRENT], font: TNR, size: co - 2})]})]}),
          first: new D.Header({children: [new D.Paragraph({children: []})]}),
        },
        children,
      }],
    });
    return D.Packer.toBlob(doc);
  }
  const tenFile = kq => {
    const s = kq.tieuDe ? window.TL.boDau(chu(kq.tieuDe.runs)).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 60) : "";
    return (s || "van-ban") + ".docx";
  };

  // ---------- 5. mẫu: câu trả lời điển hình của ChatGPT ----------
  const MAU = `Chắc chắn rồi! 😊 Dưới đây là **bản báo cáo** hoàn chỉnh mà bạn có thể dùng ngay:

---

# 📊 BÁO CÁO KẾT QUẢ TRIỂN KHAI CHUYỂN ĐỔI SỐ QUÝ III/2026

## 🎯 1. Mục tiêu đặt ra
Trong quý III/2026 , phòng Hành chính - Nhân sự đặt ra **3 mục tiêu chính**:
- Số hoá **100%** hồ sơ nhân sự đang lưu bằng giấy
- Áp dụng chữ ký số cho văn bản nội bộ ,giảm thời gian trình ký
- Đào tạo kỹ năng sử dụng AI cho **toàn bộ** nhân viên khối văn phòng

## ✅ 2. Kết quả đạt được

### 2.1. Số hoá hồ sơ
Đã số hoá **1.248/1.300** hồ sơ (đạt *96%*).Phần còn lại là hồ sơ nhân viên đã nghỉ việc trước năm 2015 , cần xin ý kiến lưu trữ.

### 2.2. Bảng tổng hợp chỉ tiêu

| Chỉ tiêu | Kế hoạch | Thực hiện | Tỷ lệ |
|---|---:|---:|---:|
| Hồ sơ số hoá | 1.300 | 1.248 | 96% |
| Văn bản ký số | 450 | 512 | 114% |
| Nhân viên được đào tạo | 85 | 79 | 93% |
| Chi phí (triệu đồng) | 120 | 104,5 | 87% |

### 2.3. Đánh giá chung
> Thời gian trình ký trung bình giảm từ **3 ngày** xuống còn **4 giờ**.

## ⚠️ 3. Khó khăn, vướng mắc
1. Một số nhân viên lớn tuổi còn ngại sử dụng phần mềm mới
2. Đường truyền tại chi nhánh Đà Nẵng chưa ổn định
   - Thường xuyên mất kết nối vào giờ cao điểm
   - Ảnh hưởng tới việc ký số từ xa
3. Kinh phí mua chữ ký số cho cấp phó phòng chưa được duyệt

## 🚀 4. Kế hoạch quý IV/2026
- **Hoàn thành** số hoá 52 hồ sơ còn lại trước 30/11/2026
- Tổ chức thêm 2 buổi đào tạo cho 6 nhân viên chưa tham gia
- Đề xuất nâng cấp đường truyền chi nhánh Đà Nẵng

---

Bạn có muốn mình chuyển báo cáo này sang dạng **slide trình chiếu** hoặc viết thêm phần **kiến nghị** không? 😊`;

  // ---------- 6. giao diện ----------
  let nguon = {kieu: "md", md: ""}, kq = null;
  const o = () => ({boEmoji: $("#wEmoji").checked, boCauAi: $("#wCauAi").checked, danhSo: $("#wSo").checked, mucLuc: $("#wMucLuc").checked, coChu: +$("#wCo").value, gianDong: +$("#wGian").value});
  function chay() {
    const txt = $("#wNhap").value;
    if (!txt.trim()) { $("#wKq").hidden = true; kq = null; return; }
    const khoi = nguon.kieu === "html" && nguon.text === txt ? tuHtml(nguon.html) : tuMarkdown(txt);
    const op = o();
    kq = chuanHoa(khoi, op);
    const g = nguon.kieu === "html" && nguon.text === txt ? {sao: 0, thang: 0, gach: 0} : demGoc(txt), d = kq.dem;
    const the = [
      [g.sao / 2 + g.thang, "dấu ** và # đã dọn"],
      [d.emoji, "biểu tượng cảm xúc đã bỏ"],
      [d.cauAi, "câu thừa của AI đã bỏ"],
      [d.deMuc, "đề mục đã đánh số"],
      [d.dauCau, "chỗ sửa dấu cách, dấu câu"],
      [d.bang, "bảng chuyển thành bảng Word"],
    ];
    $("#wTong").innerHTML = the.map(([n, l]) => `<div class="${n ? "o" : "x"}"><b>${Math.round(n)}</b>${l}</div>`).join("");
    $("#wTrang").innerHTML = xemTruoc(kq, op);
    $("#wTrang").style.setProperty("--co", op.coChu + "pt");
    $("#wTrang").style.setProperty("--gd", op.gianDong);
    $("#wKq").hidden = false;
  }
  async function tai() {
    if (!kq) return;
    const b = $("#wTai"); b.disabled = true; const t = b.textContent; b.textContent = "Đang tạo file…";
    try {
      const blob = await taoDocx(kq, o());
      const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = tenFile(kq);
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    } catch (e) { alert("Không tạo được file: " + e.message); }
    b.disabled = false; b.textContent = t;
  }
  $("#wNhap").addEventListener("paste", e => {
    const html = e.clipboardData && e.clipboardData.getData("text/html"), txt = e.clipboardData && e.clipboardData.getData("text/plain");
    // dán từ cửa sổ chat (bôi đen rồi Ctrl+C) thì có định dạng HTML, giữ lại đề mục, bảng, chữ đậm
    if (html && /<(h[1-6]|table|ul|ol|strong|b)\b/i.test(html) && !/^\s*(#|\*\*|\|)/m.test(txt || "")) {
      e.preventDefault();
      const ta = $("#wNhap"); ta.value = txt; nguon = {kieu: "html", html, text: txt}; chay();
    }
  });
  let hen; $("#wNhap").addEventListener("input", () => { if (nguon.kieu === "html" && nguon.text !== $("#wNhap").value) nguon = {kieu: "md"}; clearTimeout(hen); hen = setTimeout(chay, 250); });
  ["#wEmoji", "#wCauAi", "#wSo", "#wMucLuc", "#wCo", "#wGian"].forEach(id => $(id).addEventListener("change", chay));
  $("#wMau").addEventListener("click", () => { nguon = {kieu: "md"}; $("#wNhap").value = MAU; chay(); });
  $("#wXoa").addEventListener("click", () => { nguon = {kieu: "md"}; $("#wNhap").value = ""; chay(); $("#wNhap").focus(); });
  $("#wTai").addEventListener("click", tai);
  window._w = {get kq() { return kq; }, tuMarkdown, tuHtml, chuanHoa, taoDocx, MAU, o};
})();
