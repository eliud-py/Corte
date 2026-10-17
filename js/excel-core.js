
const ColorTransfer = (() => {
  const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const XMLH = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const parse = t => new DOMParser().parseFromString(t, "application/xml");
  const ser = n => new XMLSerializer().serializeToString(n);
  const toXml = doc => { const x = ser(doc); return /^\s*<\?xml/.test(x) ? x : XMLH + x; };

  const colNum = r => { let n = 0; for (const ch of r.replace(/\d/g, "")) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  const colName = n => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  // Texto comparable: sin tildes, sin mayúsculas, espacios simples
  const clean = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ").trim().toLowerCase();
  // Código: texto sin espacios; "123.0" == "123"; conserva ceros a la izquierda
  const normId = v => { let s = String(v ?? "").trim().replace(/^'/, ""); if (/^-?\d+\.0+$/.test(s)) s = s.split(".")[0]; return s.toLowerCase(); };

  const ID_HEADERS = ["itemid", "id", "codigo", "codigoproducto", "codigodeproducto", "sku", "referencia", "code", "itemcode"];
  const NAME_RE = /(nombre|name|descrip|producto|product|articulo)/;
  const NOT_NAME_RE = /(\bid\b|codigo|code|sku|ref)/;

  /* ---------- Lectura del libro ---------- */
  async function loadBook(file) {
    let zip;
    try { zip = await JSZip.loadAsync(await file.arrayBuffer()); }
    catch (e) { throw new Error('"' + file.name + '" no es un Excel válido o está dañado.'); }
    const rd = async p => {
      const f = zip.file(p); if (!f) return null;
      const d = parse(await f.async("string"));
      if (d.getElementsByTagName("parsererror").length) throw new Error('No se pudo leer "' + p + '" dentro de "' + file.name + '".');
      return d;
    };
    const wb = await rd("xl/workbook.xml"), rels = await rd("xl/_rels/workbook.xml.rels"), styles = await rd("xl/styles.xml");
    if (!wb || !rels || !styles) throw new Error('"' + file.name + '" no tiene estructura de libro de Excel.');
    const relMap = {};
    for (const r of rels.getElementsByTagName("Relationship")) relMap[r.getAttribute("Id")] = r.getAttribute("Target");
    const sheets = [...wb.getElementsByTagName("sheet")].map(s => {
      const t = relMap[s.getAttribute("r:id")] || "";
      return { name: s.getAttribute("name"), path: t.startsWith("/") ? t.slice(1) : "xl/" + t };
    }).filter(s => zip.file(s.path));
    if (!sheets.length) throw new Error('"' + file.name + '" no tiene hojas utilizables.');
    const ss = await rd("xl/sharedStrings.xml");
    const strs = ss ? [...ss.getElementsByTagName("si")].map(si =>
      [...si.getElementsByTagName("t")].filter(t => t.parentNode.nodeName !== "rPh").map(t => t.textContent).join("")) : [];
    return { zip, sheets, strs, styles, rd, name: file.name };
  }

  const cellVal = (c, strs) => {
    const t = c.getAttribute("t"), v = c.getElementsByTagName("v")[0];
    if (t === "inlineStr") return [...c.getElementsByTagName("t")].map(x => x.textContent).join("");
    if (!v) return "";
    return t === "s" ? (strs[+v.textContent] ?? "") : v.textContent;
  };

  // Fila de encabezados: la que contenga una columna de código (en las primeras 40 filas)
  function findHeader(doc, strs) {
    let n = 0;
    for (const row of doc.getElementsByTagName("row")) {
      if (++n > 40) break;
      const headers = {}; let idCol = 0, lastCol = 0;
      for (const c of row.children) {
        const cn = colNum(c.getAttribute("r")), h = clean(cellVal(c, strs));
        headers[cn] = h; if (h) lastCol = Math.max(lastCol, cn);
        if (!idCol && ID_HEADERS.includes(h.replace(/\s/g, ""))) idCol = cn;
      }
      if (!idCol) continue;
      const nameCol = +Object.keys(headers).find(k => +k !== idCol && NAME_RE.test(headers[k]) && !NOT_NAME_RE.test(headers[k])) || 0;
      return { rowNum: +row.getAttribute("r"), idCol, nameCol, headers, lastCol };
    }
    return null;
  }

  /* Detección de colores */
  function isMark(fill) {
    const p = fill.getElementsByTagName("patternFill")[0];
    if (!p) return false;
    const t = p.getAttribute("patternType");
    if (t === "none" || t === "gray125") return false;
    const c = p.getElementsByTagName("fgColor")[0];
    if (!c) return false;
    const rgb = (c.getAttribute("rgb") || "").toUpperCase();
    if (rgb && rgb.slice(-6) === "FFFFFF") return false;
    if (c.getAttribute("theme") === "0" && !+c.getAttribute("tint")) return false;
    if (["64", "9"].includes(c.getAttribute("indexed"))) return false;
    return true;
  }
  function fillOf(styles, s) {
    const xf = styles.getElementsByTagName("cellXfs")[0].children[s];
    if (!xf) return null;
    const f = styles.getElementsByTagName("fills")[0].children[+xf.getAttribute("fillId") || 0];
    return f && isMark(f) ? f : null;
  }
  function hexOf(fill) {
    const c = fill.getElementsByTagName("fgColor")[0], rgb = c && c.getAttribute("rgb");
    return rgb ? "#" + rgb.slice(-6) : null;
  }

  // Lee TODAS las filas de datos: código, nombre, celdas coloreadas y color de fila completa
  async function collectRows(book, sheet) {
    const doc = await book.rd(sheet.path), h = findHeader(doc, book.strs);
    if (!h) return null;
    const rows = [];
    for (const row of doc.getElementsByTagName("row")) {
      if (+row.getAttribute("r") <= h.rowNum) continue;
      let id = "", name = "";
      const marks = [];
      for (const c of row.children) {
        const cn = colNum(c.getAttribute("r"));
        if (cn === h.idCol) id = normId(cellVal(c, book.strs));
        if (cn === h.nameCol) name = clean(cellVal(c, book.strs));
        const f = fillOf(book.styles, +c.getAttribute("s") || 0);
        if (f) marks.push({ head: h.headers[cn] || "", cn, fill: f });
      }
      const rowFill = row.getAttribute("customFormat") === "1" ? fillOf(book.styles, +row.getAttribute("s") || 0) : null;
      if (id) rows.push({ id, name, marks, rowFill });
    }
    return { h, rows };
  }

  /* Estilos: agrega rellenos sin tocar los existentes  */
  function styleCtx(styles) {
    const fills = styles.getElementsByTagName("fills")[0], xfs = styles.getElementsByTagName("cellXfs")[0];
    const fc = new Map(), xc = new Map();
    return {
      xf(s, fill) {
        const fk = ser(fill), k = s + "|" + fk;
        if (xc.has(k)) return xc.get(k);
        let fid = fc.get(fk);
        if (fid === undefined) {
          fid = fills.children.length;
          fills.appendChild(styles.importNode(fill, true));
          fills.setAttribute("count", fid + 1);
          fc.set(fk, fid);
        }
        const x = (xfs.children[s] || xfs.children[0]).cloneNode(true);
        x.setAttribute("fillId", fid); x.setAttribute("applyFill", "1");
        xfs.appendChild(x);
        const id = xfs.children.length - 1;
        xfs.setAttribute("count", id + 1);
        xc.set(k, id);
        return id;
      }
    };
  }
  function setFill(doc, row, col, fill, ctx) {
    let cell = null, before = null;
    for (const c of row.children) {
      const n = colNum(c.getAttribute("r"));
      if (n === col) { cell = c; break; }
      if (n > col) { before = c; break; }
    }
    if (!cell) {
      cell = doc.createElementNS(NS, "c");
      cell.setAttribute("r", colName(col) + row.getAttribute("r"));
      row.insertBefore(cell, before);
    }
    cell.setAttribute("s", ctx.xf(+cell.getAttribute("s") || 0, fill));
  }
  const count = (doc, tag) => doc.getElementsByTagName(tag).length;

  /* ---------- Una hoja: viejo -> nuevo ---------- */
  async function processPair(o, n, so, sn, ctx, res) {
    const ro = await collectRows(o, so);
    if (!ro) { res.warnings.push('Hoja "' + so.name + '": no se encontró columna de código (Item Id) en el Excel anterior.'); return; }
    const doc = await n.rd(sn.path), hn = findHeader(doc, n.strs);
    if (!hn) { res.warnings.push('Hoja "' + sn.name + '": no se encontró columna de código (Item Id) en el Excel nuevo.'); return; }

    const wantName = res.opts.useName, useName = wantName && !!(ro.h.nameCol && hn.nameCol);
    if (wantName && !useName) res.warnings.push('Hoja "' + sn.name + '": no se halló columna de nombre en ambos archivos; se comparó solo por código.');
    const alnum = t => t.replace(/[^a-z0-9]/g, "");
    // Nombres iguales (ignorando símbolos/espacios) o uno contenido en el otro
    const same = (x, y) => {
      if (!useName) return true;
      x = alnum(x); y = alnum(y);
      return x === y || (x.length >= 4 && y.length >= 4 && (x.includes(y) || y.includes(x)));
    };

    const byId = new Map(), markedNames = new Map();
    let markedRows = 0;
    for (const r of ro.rows) {
      if (!byId.has(r.id)) byId.set(r.id, []);
      byId.get(r.id).push(r);
      if (r.marks.length || r.rowFill) { markedRows++; if (!markedNames.has(r.id)) markedNames.set(r.id, r.name); }
    }
    const dupKeys = [...byId.values()].filter(l => l.length > 1 && l.some(r => r.marks.length || r.rowFill)).length;
    if (dupKeys) res.warnings.push('Hoja "' + sn.name + '": ' + dupKeys + " producto(s) aparecen repetidos con color; se emparejaron por orden de aparición.");
    const common = new Set(), nameDiff = new Set();

    const headToCol = {};
    for (const [cn, t] of Object.entries(hn.headers)) if (t && !(t in headToCol)) headToCol[t] = +cn;
    const rowsBefore = count(doc, "row");

    const seen = new Map();
    let matched = 0, cells = 0, newCount = 0;
    for (const row of doc.getElementsByTagName("row")) {
      if (+row.getAttribute("r") <= hn.rowNum) continue;
      let id = "", name = "";
      for (const c of row.children) {
        const cn = colNum(c.getAttribute("r"));
        if (cn === hn.idCol) id = normId(cellVal(c, n.strs));
        if (cn === hn.nameCol) name = clean(cellVal(c, n.strs));
      }
      if (!id) continue;
      newCount++;
      if (markedNames.has(id)) common.add(id);
      const list = (byId.get(id) || []).filter(r => same(r.name, name));
      if (!list.length) {
        if (useName && markedNames.has(id)) {
          nameDiff.add(id);
          res.mismatch.push("Código " + id + ': antes "' + markedNames.get(id) + '" → ahora "' + name + '" (hoja ' + sn.name + ")");
        }
        continue;
      }
      const k = id + "|" + alnum(name), i = seen.get(k) || 0; seen.set(k, i + 1);
      const e = list.length === 1 ? list[0] : list[i];
      if (!e || !(e.marks.length || e.rowFill)) continue;

      matched++;
      const done = new Set(), rowKeys = new Map();
      for (const mk of e.marks) {
        const col = (mk.head && headToCol[mk.head]) || (mk.cn <= hn.lastCol ? mk.cn : 0);
        if (!col) continue;   // columna sin equivalente en el nuevo
        setFill(doc, row, col, mk.fill, ctx); done.add(col); cells++;
        rowKeys.set(ser(mk.fill), hexOf(mk.fill));
      }
      if (e.rowFill) {   // fila completa pintada en el anterior
        for (let col = 1; col <= hn.lastCol; col++) if (!done.has(col)) { setFill(doc, row, col, e.rowFill, ctx); cells++; }
        rowKeys.set(ser(e.rowFill), hexOf(e.rowFill));
      }
      for (const [key, hex] of rowKeys) { const cur = res.colors.get(key) || { hex, n: 0 }; cur.n++; res.colors.set(key, cur); }
    }

    res.diag.push('Hoja "' + sn.name + '": con color en el anterior ' + markedNames.size + " código(s); aparecen en el nuevo " + common.size +
      "; con nombre distinto " + nameDiff.size + ". Columnas usadas → anterior: código \"" + ro.h.headers[ro.h.idCol] + "\"" +
      (ro.h.nameCol ? ", nombre \"" + ro.h.headers[ro.h.nameCol] + "\"" : ", sin nombre") + "; nuevo: código \"" + hn.headers[hn.idCol] + "\"" +
      (hn.nameCol ? ", nombre \"" + hn.headers[hn.nameCol] + "\"" : ", sin nombre") + ".");
    // Hoja sin cambios: se deja EXACTAMENTE como venía
    if (!cells) { res.rowsKept += rowsBefore; res.stats.sheets++; res.sheets.push([so.name, ro.rows.length, newCount, markedRows, 0]); return; }

    // Solo se agregaron estilos a celdas: no se toca ningún dato. Único freno: que el XML quede inválido.
    const out = toXml(doc);
    if (parse(out).getElementsByTagName("parsererror").length) {
      res.warnings.push('Hoja "' + sn.name + '": no se pudo guardar con colores y se dejó sin cambios.');
      res.rowsKept += rowsBefore; res.sheets.push([so.name, ro.rows.length, newCount, markedRows, 0]);
      return;
    }
    n.zip.file(sn.path, out);
    res.rowsKept += rowsBefore;
    res.stats.matches += matched; res.stats.colored += matched; res.stats.cells += cells; res.stats.sheets++;
    res.sheets.push([so.name, ro.rows.length, newCount, markedRows, matched]);
  }

  /*Proceso completo */
  async function run(oldFile, newFile, progress = () => {}, opts = { useName: true }) {
    const ext = (newFile.name.match(/\.(xlsx|xlsm)$/i) || [, "xlsx"])[1].toLowerCase();
    progress(10, "Leyendo archivo anterior…");
    const o = await loadBook(oldFile);
    progress(25, "Leyendo archivo nuevo…");
    const n = await loadBook(newFile);

    const res = { opts, diag: [], warnings: [], mismatch: [], sheets: [], colors: new Map(), rowsKept: 0, stats: { sheets: 0, matches: 0, colored: 0, cells: 0 } };
    const byName = new Map(n.sheets.map(s => [s.name.trim().toLowerCase(), s]));
    const used = new Set(), pairs = [], lostOld = [];
    for (const s of o.sheets) {
      const m = byName.get(s.name.trim().toLowerCase());
      if (m) { pairs.push([s, m]); used.add(m); } else lostOld.push(s);
    }
    const lostNew = n.sheets.filter(s => !used.has(s));
    if (lostOld.length === 1 && lostNew.length === 1) {
      res.warnings.push('La hoja "' + lostOld[0].name + '" del anterior se comparó con "' + lostNew[0].name + '" del nuevo (nombres distintos).');
      pairs.push([lostOld[0], lostNew[0]]); lostOld.length = lostNew.length = 0;
    }
    lostOld.forEach(s => res.warnings.push('La hoja "' + s.name + '" del anterior no existe en el nuevo: se omitió.'));
    lostNew.forEach(s => res.warnings.push('La hoja "' + s.name + '" del nuevo no existe en el anterior: queda sin cambios.'));
    if (!pairs.length) throw new Error("Ninguna hoja del Excel anterior coincide con las del nuevo.");

    const ctx = styleCtx(n.styles);
    let i = 0;
    for (const [so, sn] of pairs) {
      progress(35 + 45 * (i++ / pairs.length), 'Comparando hoja "' + sn.name + '"…');
      await processPair(o, n, so, sn, ctx, res);
    }
    if (!res.sheets.length) throw new Error("No se encontró la columna de código (Item Id) en ninguna hoja comparable.");
    if (!res.sheets.some(s => s[3] > 0)) throw new Error("No se encontró ningún producto con color en el Excel anterior.");
    if (!res.stats.cells) throw new Error("Ningún producto con color del anterior coincide con el nuevo" + (opts.useName ? " (código y nombre)" : " (código)") + ". No se generó archivo. DETALLE: " + res.diag.concat(res.warnings).join(" · "));

    // Cantidad de artículos (filas con código) en cada archivo
    const cnt = async (b, sh) => { const r = await collectRows(b, sh); return r ? r.rows.length : 0; };
    let totOld = 0, totNew = 0;
    for (const sh of o.sheets) totOld += await cnt(o, sh);
    for (const sh of n.sheets) totNew += await cnt(n, sh);
    for (const sh of lostOld) res.sheets.push([sh.name, await cnt(o, sh), "—", "—", "—"]);
    for (const sh of lostNew) res.sheets.push([sh.name, "—", await cnt(n, sh), "—", "—"]);

    progress(85, "Generando archivo de salida…");
    n.zip.file("xl/styles.xml", toXml(n.styles));
    const mime = ext === "xlsm" ? "application/vnd.ms-excel.sheet.macroEnabled.12" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    const blob = await n.zip.generateAsync({ type: "blob", compression: "DEFLATE", mimeType: mime });
    return {
      blob, name: newFile.name.replace(/\.(xlsx|xlsm)$/i, "") + "_colores_recuperados." + ext,
      stats: res.stats, totals: { old: totOld, new: totNew }, sheets: res.sheets, warnings: res.warnings, mismatch: res.mismatch,
      colors: [...res.colors.values()], rowsKept: res.rowsKept
    };
  }

  return { run };
})();
if (typeof module !== "undefined") module.exports = ColorTransfer;
