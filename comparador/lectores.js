// Lectores de listas de precios. Los usa la página (en el navegador) y procesar.js (en Node),
// así las dos leen cada archivo exactamente igual.
//
// Un producto leído es un arreglo:
//   [proveedor, codigo, descripcion, detalle, categoria, presentacion, compra, precioLista, unidades, precioAnterior]
// precioLista es el de la lista, sin descuento ni ajuste de IVA; unidades es cuántas trae la presentación (o null).
(function (raiz) {
  const UMBRAL_RARO = 0.25;
  const MESES = { ENERO: 1, FEBRERO: 2, MARZO: 3, ABRIL: 4, MAYO: 5, JUNIO: 6, JULIO: 7, AGOSTO: 8,
    SEPTIEMBRE: 9, SETIEMBRE: 9, OCTUBRE: 10, NOVIEMBRE: 11, DICIEMBRE: 12 };

  const norm = s => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/\s+/g, " ").trim();
  const limpio = s => String(s ?? "").replace(/\s+/g, " ").trim();
  const capital = s => (s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : "");
  const fmtCant = n => String(Math.round(n * 100) / 100);

  // "$ 6,530.00", "6.530,00", "27511" -> número
  function numero(v) {
    if (typeof v === "number") return Number.isFinite(v) ? v : null;
    if (v == null) return null;
    let s = String(v).replace(/[$\s]/g, "");
    if (!/^-?[\d.,]+$/.test(s) || !/\d/.test(s)) return null;
    const c = s.lastIndexOf(","), p = s.lastIndexOf(".");
    if (c >= 0 && p >= 0) s = c > p ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
    else if (c >= 0) s = /,\d{1,2}$/.test(s) ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
    else if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, "");
    const n = parseFloat(s);
    return Number.isFinite(n) ? n : null;
  }

  function iso(y, m, d) {
    y = +y; m = +m; d = +d;
    if (y < 2000 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) return null;
    return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  // Busca una fecha en los primeros textos del archivo ("24-10-2026", "1/10/2026", "26 DE OCTUBRE 2026").
  function buscarFecha(textos) {
    for (const t of textos) {
      if (t instanceof Date && !isNaN(t)) return iso(t.getFullYear(), t.getMonth() + 1, t.getDate());
      const s = norm(t);
      let m = s.match(/\b(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})\b/);
      if (m && iso(m[3], m[2], m[1])) return iso(m[3], m[2], m[1]);
      m = s.match(/\b(\d{1,2}) DE ([A-Z]+),? (?:DE |DEL )?(\d{4})\b/);
      if (m && MESES[m[2]]) return iso(m[3], MESES[m[2]], m[1]);
    }
    return null;
  }

  // Cantidad por presentación escrita en la descripción: "Cx5000", "CJx60", "PQx1000", "BTxx400", " x100u.".
  // No toma medidas ("34x18x20"), kilos ("x1KG") ni paquetes de paquetes ("x100PAQ").
  const PACK = /(?<!\d)x+\s*(\d{1,3}(?:\.\d{3})+|\d+)(?![\d.,]|\s*(?:KG|K\b|GR|PAQ|MT|CM|MM|CC|LT))/gi;
  const PACK_COMPRA = /x\s*(\d+)\s*u|\(\s*(\d+)\s*u/i;
  const POR_UNIDAD = /(\bu\.|\bu$|\bunidad|\bunidades)/i;
  const POR_KILO = /x\s*KG/i;
  function cantidadEnTexto(desc, compra) {
    if (POR_KILO.test(desc)) return null;
    const enCompra = PACK_COMPRA.exec(compra);
    if (enCompra) return +(enCompra[1] || enCompra[2]);
    if (POR_UNIDAD.test(compra)) return 1;
    const enDesc = [...desc.matchAll(PACK)];
    return enDesc.length ? +enDesc[enDesc.length - 1][1].replace(/\./g, "") : null;
  }

  // ---------- Planillas (Excel / CSV / Google Sheets descargado) ----------

  // Hojas visibles con filas, como arreglos de celdas. XLSX es la librería SheetJS.
  function hojasVisibles(wb, XLSX) {
    return wb.SheetNames.map((nombre, i) => ({
      nombre,
      oculta: ((wb.Workbook && wb.Workbook.Sheets && wb.Workbook.Sheets[i] && wb.Workbook.Sheets[i].Hidden) || 0) !== 0,
      filas: XLSX.utils.sheet_to_json(wb.Sheets[nombre], { header: 1, raw: true, defval: null, blankrows: false }),
    })).filter(h => !h.oculta && h.filas.length);
  }

  // Ubica las columnas configuradas en una fila de títulos. null si no es la fila de títulos.
  function indiceColumnas(fila, columnas) {
    const celdas = fila.map(norm);
    const idx = {};
    for (const campo of ["codigo", "desc", "precio", "cant", "um", "detalle", "compra"]) {
      const titulo = columnas[campo];
      idx[campo] = -1;
      if (!titulo) continue;
      const t = norm(titulo);
      let i = celdas.indexOf(t);
      if (i < 0) i = celdas.findIndex(c => c && c.startsWith(t));
      idx[campo] = i;
    }
    return idx.desc >= 0 && idx.precio >= 0 ? idx : null;
  }

  // Fila que parece de títulos: la que más títulos conocidos tiene (descripción, precio, código…);
  // si ninguna, la de más textos que tenga números debajo.
  const TITULOS = /DESCRIP|PRODUCTO|ARTICULO|PRECIO|IMPORTE|COD|CANT|U\/M|C\/U|UNIDAD/;
  function adivinarTitulos(filas) {
    let mejor = 0, puntaje = 0;
    for (let i = 0; i < Math.min(filas.length, 30); i++) {
      const textos = filas[i].filter(c => typeof c === "string" && c.trim() && numero(c) == null);
      const siguiente = filas.slice(i + 1, i + 4).some(f => f.some(c => numero(c) != null && numero(c) > 0));
      if (textos.length < 2 || !siguiente) continue;
      const p = textos.filter(c => TITULOS.test(norm(c))).length * 100 + textos.length;
      if (p > puntaje) { mejor = i; puntaje = p; }
    }
    return mejor;
  }

  function leerPlanilla(hojas, cfg) {
    const items = [];
    const multi = hojas.length > 1;
    let hojasConTitulos = 0;
    for (const h of hojas) {
      let idx = null;
      let cat = multi ? limpio(h.nombre) : "";
      for (const fila of h.filas) {
        if (!idx) { idx = indiceColumnas(fila, cfg.columnas); if (idx) hojasConTitulos++; continue; }
        const precio = numero(fila[idx.precio]);
        const desc = limpio(fila[idx.desc]).replace(/[*\-\s]+$/, "");
        if (precio == null || precio <= 0 || !desc) {
          const hasta = Math.max(...Object.values(idx)) + 1;
          const llenas = fila.slice(0, hasta).filter(c => c != null && String(c).trim());
          if (precio == null && llenas.length === 1 && typeof llenas[0] === "string") cat = limpio(llenas[0]);
          continue;
        }
        const um = idx.um >= 0 ? limpio(fila[idx.um]) : "";
        const cant = idx.cant >= 0 ? numero(fila[idx.cant]) : null;
        const compraTxt = idx.compra >= 0 ? limpio(fila[idx.compra]) : "";
        let unidades = null;
        if (/^kg$/i.test(um)) unidades = null;
        else if (cant) unidades = cant;
        else if (cfg.unidadesFijas) unidades = cfg.unidadesFijas;
        else if (cfg.cantEnDesc) unidades = cantidadEnTexto(desc, compraTxt);
        let pres = "";
        if (um && cant) pres = `${capital(um)} x ${fmtCant(cant)} u.`;
        else if (um) pres = capital(um);
        else if (cfg.unidadesFijas) pres = `Precio x ${cfg.unidadesFijas} u.`;
        items.push([
          cfg.nombre,
          idx.codigo >= 0 ? limpio(fila[idx.codigo]) : "",
          desc,
          idx.detalle >= 0 ? limpio(fila[idx.detalle]) : "",
          cat ? capital(cat).replace(/\b\w/g, l => l.toUpperCase()) : "",
          pres,
          compraTxt.toLowerCase().replace(/^por\s+/, "comprando "),
          Math.round(precio * 100) / 100,
          unidades,
          null,
        ]);
      }
    }
    const primeras = hojas.slice(0, 1).flatMap(h => h.filas.slice(0, 6).flat()).filter(c => c != null);
    return { items, fecha: buscarFecha(primeras), titulosEncontrados: hojasConTitulos > 0 };
  }

  // ---------- PDF de Vassoi ----------

  const FILA_VASSOI = /^(.+?)\s+(RECT\.|CIRC\.|OVAL\.|BOB\.)\s+([\d\s]+?)\s+\$\s*([\d.]+,\d+)$/;
  function leerVassoi(lineas, cfg) {
    const items = [];
    let categoria = "";
    lineas.forEach((linea, i) => {
      linea = limpio(linea);
      const m = linea.match(FILA_VASSOI);
      if (!m) {
        if ((lineas[i + 1] || "").trim().startsWith("MODELO")) categoria = linea;
        return;
      }
      const nums = m[3].split(" ");
      const forma = m[2].replace(".", "").toLowerCase();
      const capac = nums.length === 6 || nums.length === 8 ? `${nums[nums.length - 3]} cc` : "";
      let medidas = "";
      if (nums.length === 8) medidas = [nums[1], nums[3]].join(" x ");
      else if (nums.length === 6) medidas = nums[1];
      const cat = categoria.replace(/\s*\(.*\)/, "").trim();
      const modelo = m[1].replace(/\*/g, "").trim();
      items.push([
        cfg.nombre, modelo, `${cat} ${modelo} ${capac}`.trim(),
        [forma, medidas ? `boca ${medidas} mm` : "", `caja x ${nums[nums.length - 2]} u.`].filter(Boolean).join(" · "),
        capital(cat).replace(/\b\w/g, l => l.toUpperCase()),
        "Precio x 100 u.", "", numero(m[4]), 100, null,
      ]);
    });
    return { items, fecha: buscarFecha(lineas.slice(0, 5)), titulosEncontrados: items.length > 0 };
  }

  // Texto de un PDF (pdf.js) en líneas: agrupa los trozos por altura y los ordena de izquierda a derecha.
  async function lineasPDF(doc) {
    const lineas = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const pagina = await doc.getPage(p);
      const { items } = await pagina.getTextContent();
      const filas = [];
      for (const it of items) {
        if (!it.str || !it.str.trim()) continue;
        const y = it.transform[5], x = it.transform[4];
        let f = filas.find(f => Math.abs(f.y - y) < 3);
        if (!f) filas.push(f = { y, partes: [] });
        f.partes.push({ x, s: it.str });
      }
      filas.sort((a, b) => b.y - a.y);
      for (const f of filas) lineas.push(limpio(f.partes.sort((a, b) => a.x - b.x).map(t => t.s).join(" ")));
    }
    return lineas;
  }

  // ---------- Comparar con la lista anterior ----------

  const clave = it => `${it[0]}|${it[1] || it[2]}`;

  // Anota en cada producto nuevo el precio de lista anterior y resume los cambios.
  function compararConAnterior(nuevos, anteriores) {
    const previo = new Map(anteriores.map(it => [clave(it), it[7]]));
    const vistos = new Set();
    const raros = [];
    const variaciones = [];
    let nuevosProductos = 0;
    for (const it of nuevos) {
      const k = clave(it);
      vistos.add(k);
      if (!previo.has(k)) { nuevosProductos++; continue; }
      const antes = previo.get(k);
      it[9] = antes;
      if (!antes) continue;
      const v = it[7] / antes - 1;
      variaciones.push(v);
      if (Math.abs(v) >= UMBRAL_RARO) raros.push(it);
    }
    variaciones.sort((a, b) => a - b);
    return {
      nuevos: nuevosProductos,
      faltan: anteriores.filter(it => !vistos.has(clave(it))).length,
      raros,
      mediana: variaciones.length ? variaciones[Math.floor(variaciones.length / 2)] : null,
      comparados: variaciones.length,
    };
  }

  const L = { UMBRAL_RARO, numero, buscarFecha, cantidadEnTexto, hojasVisibles, adivinarTitulos, indiceColumnas,
    leerPlanilla, leerVassoi, lineasPDF, compararConAnterior, clave, norm };
  if (typeof module !== "undefined" && module.exports) module.exports = L;
  else raiz.Lectores = L;
})(typeof window !== "undefined" ? window : globalThis);
