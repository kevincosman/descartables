// Lee las listas de cada proveedor y arma comparador.html con los datos adentro.
// Uso (desde esta carpeta):  npm install   (una sola vez)   y después   node procesar.js
// Las listas se buscan en la carpeta de arriba (descartables/); Bolsatodo se lee de su Google Sheets.
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const pdfjs = require("pdfjs-dist/legacy/build/pdf.js");
const L = require("./lectores.js");

const AQUI = __dirname;
const LISTAS = path.join(AQUI, "..");
const IVA = 0.21;

// Reglas de cada proveedor. descuento: fracción que se resta al precio de lista.
// columnas: título de cada columna en la planilla (se busca en todas las hojas visibles).
const PROVEEDORES = [
  { nombre: "Darnel", archivo: "LISTA DARNEL PILARICA OCTUBRE 26 A18.xlsx", con_iva: false, descuento: 0.48,
    formato: "planilla",
    columnas: { codigo: "CODIGO PILARICA", um: "U/M", cant: "C/U", desc: "DESCRIPCION", detalle: "COMENT", precio: "PRECIO" } },
  { nombre: "Vassoi", archivo: "LISTA VASSOI 01-10-2026.pdf", con_iva: false, descuento: 0, formato: "vassoi" },
  { nombre: "Bolsatodo", sheet: "13NKMIVL9uDetW99qLbsnoWxORnemxOX0EZ2In5R78o8", con_iva: false, descuento: 0,
    formato: "planilla", cantEnDesc: true,
    columnas: { codigo: "COD. ARTICULO", desc: "DESCRIPCION", compra: "DESC. ADICIONAL", precio: "PRECIO" },
    ayuda: "En Google Sheets: Archivo → Descargar → Microsoft Excel (.xlsx), y subí ese archivo." },
];

async function leer(cfg) {
  if (cfg.formato === "vassoi") {
    const datos = new Uint8Array(fs.readFileSync(path.join(LISTAS, cfg.archivo)));
    const doc = await pdfjs.getDocument({ data: datos, verbosity: 0 }).promise;
    return L.leerVassoi(await L.lineasPDF(doc), cfg);
  }
  let wb;
  if (cfg.sheet) {
    const r = await fetch(`https://docs.google.com/spreadsheets/d/${cfg.sheet}/export?format=xlsx`);
    if (!r.ok) throw new Error(`No pude bajar el Google Sheets de ${cfg.nombre} (${r.status})`);
    wb = XLSX.read(new Uint8Array(await r.arrayBuffer()), { type: "array", cellDates: true });
  } else {
    wb = XLSX.read(fs.readFileSync(path.join(LISTAS, cfg.archivo)), { cellDates: true });
  }
  return L.leerPlanilla(L.hojasVisibles(wb, XLSX), cfg);
}

async function main() {
  const previo = fs.existsSync(path.join(AQUI, "datos.json"))
    ? JSON.parse(fs.readFileSync(path.join(AQUI, "datos.json"), "utf8")) : null;
  const salida = { generado: new Date().toISOString().slice(0, 10), iva: IVA, proveedores: [], items: [] };
  for (const cfg of PROVEEDORES) {
    const { items, fecha } = await leer(cfg);
    const fechaLista = fecha || cfg.fecha || salida.generado;
    const antes = previo ? previo.proveedores.find(p => p.nombre === cfg.nombre) : null;
    const anteriores = previo ? previo.items.filter(it => it[0] === cfg.nombre) : [];
    if (antes && antes.fecha === fechaLista) {
      // Misma lista que la vez pasada: se conserva la comparación con la anterior.
      const guardado = new Map(anteriores.map(it => [L.clave(it), it[9]]));
      for (const it of items) it[9] = guardado.get(L.clave(it)) ?? null;
    } else {
      L.compararConAnterior(items, anteriores);
    }
    const { archivo, sheet, ...config } = cfg;
    salida.proveedores.push({ ...config, fecha: fechaLista });
    salida.items.push(...items);
    const conUnidad = items.filter(it => it[8]).length;
    console.error(`${cfg.nombre}: ${items.length} productos, ${conUnidad} con precio por unidad (lista del ${fechaLista})`);
  }
  const plantilla = fs.readFileSync(path.join(AQUI, "plantilla.html"), "utf8");
  const lectores = fs.readFileSync(path.join(AQUI, "lectores.js"), "utf8");
  const datos = JSON.stringify(salida).replace(/<\//g, "<\\/");
  const pagina = plantilla.replace("/*__LECTORES__*/", () => lectores).replace("/*__DATOS__*/null", () => datos);
  fs.writeFileSync(path.join(AQUI, "comparador.html"), pagina);
  // Copia para subir a Netlify (arrastrando la carpeta "publicar"): ahí no hay carga de listas ni correcciones.
  fs.mkdirSync(path.join(AQUI, "publicar"), { recursive: true });
  fs.writeFileSync(path.join(AQUI, "publicar", "index.html"), `<!doctype html><html lang="es"><head><meta charset="utf-8">`
    + `<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"></head><body>${pagina}</body></html>`);
  fs.writeFileSync(path.join(AQUI, "datos.json"), JSON.stringify(salida));
}

main().catch(e => { console.error(e.message); process.exit(1); });
