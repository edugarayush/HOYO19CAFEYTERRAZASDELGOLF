/* Hoyo 19 · carta web
   Fuente de datos, en este orden:
   1) La Google Sheet, a través del Apps Script (SHEET_URL). Se edita la planilla y la web cambia sola.
   2) Si SHEET_URL está vacío o Google no responde: carta-hoyo-19.xlsx de esta carpeta (respaldo). */

/* ===== Configuración ===== */
const SHEET_URL = "https://script.google.com/macros/s/AKfycbw9KR3wU1r5C8d4jv6m2Q0LcUIS7WtAqefPTh4NbDkxXqbvhoxlwP8Kw8KQfcIWpZbm/exec";   // URL del Apps Script (termina en /exec). Vacío = usa solo el Excel de respaldo
const EXCEL = "carta-hoyo-19.xlsx";
const ESPERA_SHEET_MS = 8000; // si la planilla tarda más que esto, se usa el respaldo
// Nombre de cada página de la carta (columna "Página" del Excel). Una página nueva sin nombre acá sale como "Carta".
const PAGINAS = { 1: "Café & Delicias", 2: "Cocina", 3: "Bebidas" };

/* ===== Utilidades ===== */
const $ = s => document.querySelector(s);
const norm = s => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().trim();
const txt = v => (v === null || v === undefined) ? "" : String(v).trim();
const esc = s => s.replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const precioNum = v => {
  if (typeof v === "number") return v > 0 ? v : null;
  const n = Number(String(v ?? "").replace(/[^\d,]/g, "").replace(",", "."));
  return n > 0 ? n : null;
};
const fmt = n => "$ " + Math.round(n).toLocaleString("es-AR");
const slug = s => norm(s).replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");

/* ===== Íconos de categoría (como en el PDF; las que no tienen ícono allá, acá tampoco) ===== */
const ICONOS = [
  ["cafeteria", '<path d="M4 9h12v4a5 5 0 0 1-5 5H9a5 5 0 0 1-5-5z"/><path d="M16 10h1.5a2.5 2.5 0 0 1 0 5H16"/><path d="M3 21h16"/><path d="M8 3c-1 1 1 2 0 3M12 3c-1 1 1 2 0 3"/>'],
  ["desayun", '<path d="M3 15c2-6 6-9 9-9s7 3 9 9c-3-1-5-1-6 1-1-3-2-4-3-4s-2 1-3 4c-1-2-3-2-6-1z"/><path d="M9 9.5l1 3M15 9.5l-1 3"/>'],
  ["jugo", '<path d="M7 8h10l-1.2 13H8.2z"/><path d="M12 8l2.5-5H18"/><path d="M7.4 12h9.2"/>'],
  ["sandwich", '<path d="M3 10l9-5 9 5-9 4z"/><path d="M3 14l9 4 9-4"/><path d="M3 10v4M21 10v4"/>'],
  ["empanada", '<path d="M3 16a9 8 0 0 1 18 0z"/><path d="M4.5 16l1-1.6 1 1.6 1-1.6 1 1.6 1-1.6 1 1.6 1-1.6 1 1.6 1-1.6 1 1.6 1-1.6 1 1.6 1-1.6 1 1.6"/>'],
  ["picada", '<path d="M3 16l12-8 6 4v4z"/><path d="M3 16v3h18v-3"/><circle cx="14" cy="13.5" r="1"/><circle cx="10" cy="14.5" r=".6"/>'],
  ["papa", '<path d="M6 10h12l-2 11H8z"/><path d="M8 10V4M11 10V3M14 10V4M16.5 10V5"/>'],
  ["ensalada", '<path d="M5 19C5 10 10 5 20 4c-1 10-6 15-15 15z"/><path d="M5 19l9-9"/>'],
  ["plato", '<path d="M3 18h18"/><path d="M5 18a7 7 0 0 1 14 0"/><path d="M12 11V9"/><circle cx="12" cy="7.6" r="1.2"/>'],
  ["postre", '<path d="M4 20h16v-7H4z"/><path d="M4 16c2 1 4 1 6 0s4-1 6 0 3 1 4 0"/><path d="M12 13V9.5"/><path d="M12 7.5c.9 0 1.2-1.6 0-3-1.2 1.4-.9 3 0 3z"/>'],
  ["aperitivo", '<path d="M4 4h16l-8 9z"/><path d="M12 13v7M8 20h8"/>'],
  ["gin", '<path d="M5 4h14l-7 8z"/><path d="M12 12v8M8.5 20h7"/><path d="M15 2l-2 5"/>'],
  ["trago", '<path d="M6 5h12l-1.5 15h-9z"/><path d="M6.6 11h10.8"/><path d="M14.5 2l-1.2 7"/>'],
  ["whisky", '<path d="M5 7h14l-1.5 13h-11z"/><path d="M8 12h4v4H8zM12.5 10.5h3.5V14h-3.5z"/>'],
  ["espumante", '<path d="M10 3h4l-.5 9a1.5 1.5 0 0 1-3 0z"/><path d="M12 13.5V20M9 20h6"/>'],
  ["vino", '<path d="M7 3h10c0 6-2 9-5 9s-5-3-5-9z"/><path d="M12 12v8M8 20h8"/><path d="M7.4 6.5h9.2"/>'],
  ["cerveza", '<path d="M5 8h10v12H5z"/><path d="M15 10.5h2a2 2 0 0 1 2 2v3a2 2 0 0 1-2 2h-2"/><path d="M5 8a2.5 2.5 0 0 1 5-1.2A2.5 2.5 0 0 1 15 8"/><path d="M8.5 11v6M11.5 11v6"/>'],
  ["sin alcohol", '<path d="M10 2h4v4l2 3v12H8V9l2-3z"/><path d="M8 13h8"/>'],
];
const icono = cat => {
  const n = norm(cat);
  const hit = ICONOS.find(([k]) => n.includes(k));
  return hit ? `<svg viewBox="0 0 24 24" aria-hidden="true">${hit[1]}</svg>` : "";
};

/* ===== Lectura del Excel =====
   Busca la fila de títulos (la que tiene "Producto" y "Precio"), así las filas de título de arriba no molestan.
   Columnas reconocidas (sin importar mayúsculas ni tildes): Página, Categoría, Producto,
   Descripción / detalle, Precio ($), Observaciones. Opcional: Activo (NO = no se muestra). */
const COLS = {
  pagina: ["pagina", "seccion", "carta"],
  categoria: ["categoria"],
  producto: ["producto", "plato", "nombre"],
  descripcion: ["descripcion", "detalle"],
  precio: ["precio"],
  obs: ["observacion", "nota"],
  activo: ["activo", "disponible"],
};
/* Cada fuente devuelve una lista de hojas, y cada hoja es una lista de filas (como en la planilla). */
async function hojasDesdeSheet() {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ESPERA_SHEET_MS);
  try {
    const res = await fetch(SHEET_URL, { cache: "no-store", signal: ctrl.signal });
    if (!res.ok) throw new Error("Apps Script respondió " + res.status);
    const data = await res.json();
    if (!data.ok || !Array.isArray(data.filas)) throw new Error("Respuesta inesperada del Apps Script");
    return [data.filas];
  } finally { clearTimeout(t); }
}
async function hojasDesdeExcel() {
  const res = await fetch(EXCEL + "?v=" + Date.now(), { cache: "no-store" });
  if (!res.ok) throw new Error("No se encontró " + EXCEL + " (" + res.status + ")");
  const wb = XLSX.read(await res.arrayBuffer(), { type: "array" });
  return wb.SheetNames.map(n => XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null, blankrows: false }));
}

async function leerCarta() {
  if (SHEET_URL) {
    try {
      const p = productosDe(await hojasDesdeSheet());
      if (p.length) return { productos: p, fuente: "Google Sheet" };
      console.warn("La planilla no trajo productos; uso el Excel de respaldo");
    } catch (err) { console.warn("No se pudo leer la planilla; uso el Excel de respaldo.", err); }
  }
  return { productos: productosDe(await hojasDesdeExcel()), fuente: EXCEL };
}

function productosDe(hojas) {
  const productos = [];
  for (const filas of hojas) {
    const iTit = filas.findIndex(f => {
      const n = f.map(norm);
      return n.some(c => c.startsWith("producto")) && n.some(c => c.startsWith("precio"));
    });
    if (iTit < 0) continue;
    const tit = filas[iTit].map(norm);
    const idx = {};
    for (const [k, alias] of Object.entries(COLS)) idx[k] = tit.findIndex(c => alias.some(a => c.startsWith(a)));
    const val = (f, k) => idx[k] >= 0 ? f[idx[k]] : null;
    let pagAnt = null, catAnt = "";
    for (const f of filas.slice(iTit + 1)) {
      const producto = txt(val(f, "producto"));
      if (!producto) continue;
      if (norm(val(f, "activo")) === "no") continue;
      const pag = txt(val(f, "pagina")) || pagAnt || "";
      const cat = txt(val(f, "categoria")) || catAnt || "Otros";
      pagAnt = pag; catAnt = cat;
      productos.push({
        pagina: pag, categoria: cat, producto,
        descripcion: txt(val(f, "descripcion")),
        precio: precioNum(val(f, "precio")),
        obs: txt(val(f, "obs")),
      });
    }
  }
  return productos;
}

/* Agrupa en páginas → categorías, respetando el orden del Excel */
function agrupar(productos) {
  const paginas = [];
  for (const p of productos) {
    const nombre = PAGINAS[p.pagina] || (isNaN(p.pagina) && p.pagina) || "Carta";
    let pag = paginas.find(x => x.nombre === nombre);
    if (!pag) paginas.push(pag = { nombre, id: slug(nombre), categorias: [] });
    let cat = pag.categorias.find(c => c.nombre === p.categoria);
    if (!cat) pag.categorias.push(cat = { nombre: p.categoria, id: pag.id + "--" + slug(p.categoria), items: [] });
    cat.items.push(p);
  }
  return paginas;
}

/* ===== Render ===== */
let PAGS = [], tabActiva = "todo";

function resaltar(texto, q) {
  const t = esc(texto);
  if (!q) return t;
  // resalta ignorando tildes: busca la posición en la versión normalizada
  const n = norm(texto), i = n.indexOf(q);
  if (i < 0) return t;
  return esc(texto.slice(0, i)) + "<mark>" + esc(texto.slice(i, i + q.length)) + "</mark>" + esc(texto.slice(i + q.length));
}

function tituloCategoria(nombre) {
  // "Empanadas (c/u)" → EMPANADAS con "(c/u)" en chico, como en el PDF
  const m = nombre.match(/^(.*?)\s*(\(.*\))$/);
  const base = m ? m[1] : nombre, extra = m ? ` <small>${esc(m[2])}</small>` : "";
  const ic = icono(nombre);
  return `<h3 class="cat-titulo${ic ? "" : " sin-icono"}">${ic}<span>${esc(base)}${extra}</span></h3>`;
}

function renderItem(p, q) {
  if (p.precio === null) { // sin precio → línea de nota (guarniciones, salsas…)
    return `<li class="nota"><b>${resaltar(p.producto, q)}:</b> ${resaltar(p.descripcion || p.obs, q)}</li>`;
  }
  return `<li class="item">
    <div class="fila"><span class="nombre">${resaltar(p.producto, q)}</span><span class="guia"></span><span class="precio">${fmt(p.precio)}</span></div>
    ${p.descripcion ? `<p class="desc">${resaltar(p.descripcion, q)}</p>` : ""}
    ${p.obs ? `<p class="obs">${esc(p.obs)}</p>` : ""}
  </li>`;
}

function render() {
  const q = norm($("#buscar").value);
  const coincide = p => !q || [p.producto, p.descripcion, p.categoria, p.obs].some(s => norm(s).includes(q));
  let html = "", total = 0;
  for (const pag of PAGS) {
    if (!q && tabActiva !== "todo" && tabActiva !== pag.id) continue;
    let cuerpo = "";
    for (const cat of pag.categorias) {
      const enCat = norm(cat.nombre).includes(q);
      const items = cat.items.filter(p => enCat || coincide(p));
      if (!items.length) continue;
      total += items.length;
      cuerpo += `<section class="categoria" id="${cat.id}">${tituloCategoria(cat.nombre)}<ul class="items">${items.map(p => renderItem(p, q)).join("")}</ul></section>`;
    }
    if (cuerpo) html += `<article class="hoja" id="${pag.id}"><h2 class="hoja-titulo">${esc(pag.nombre)}</h2>${cuerpo}</article>`;
  }
  $("#carta").innerHTML = html || `<p class="estado">No encontramos “${esc($("#buscar").value)}” en la carta.</p>`;
  $("#limpiar").hidden = !q;
  renderChips(q);
}

function renderTabs() {
  const tabs = [{ id: "todo", nombre: "Todo" }, ...PAGS];
  $("#tabs").innerHTML = tabs.map(t =>
    `<button class="tab" role="tab" data-tab="${t.id}" aria-selected="${t.id === tabActiva}">${esc(t.nombre)}</button>`).join("");
}

function renderChips(q) {
  const cats = q ? [] : PAGS.filter(p => tabActiva === "todo" || p.id === tabActiva).flatMap(p => p.categorias);
  $("#chips").innerHTML = cats.map(c => `<button class="chip" data-ir="${c.id}">${esc(c.nombre)}</button>`).join("");
  $("#chips").hidden = !cats.length;
}

/* ===== Eventos ===== */
$("#tabs").addEventListener("click", e => {
  const b = e.target.closest("[data-tab]"); if (!b) return;
  tabActiva = b.dataset.tab;
  $("#buscar").value = "";
  renderTabs(); render();
  window.scrollTo({ top: $(".barra").offsetTop, behavior: "smooth" });
});
$("#chips").addEventListener("click", e => {
  const b = e.target.closest("[data-ir]"); if (!b) return;
  document.getElementById(b.dataset.ir)?.scrollIntoView({ behavior: "smooth", block: "start" });
  b.scrollIntoView({ behavior: "smooth", inline: "center", block: "nearest" });
});
$("#buscar").addEventListener("input", render);
$("#limpiar").addEventListener("click", () => { $("#buscar").value = ""; render(); $("#buscar").focus(); });

/* ===== Inicio ===== */
(async () => {
  try {
    const { productos, fuente } = await leerCarta();
    if (!productos.length) throw new Error("La carta no tiene productos (¿cambió la fila de títulos?)");
    PAGS = agrupar(productos);
    renderTabs(); render();
    console.info(`Carta Hoyo 19: ${productos.length} filas leídas de ${fuente}`);
  } catch (err) {
    console.error(err);
    const local = location.protocol === "file:";
    $("#carta").innerHTML = `<p class="estado">No pudimos cargar la carta.${local
      ? "<br>Abrí la carpeta con un servidor local (ver README), no con doble clic."
      : "<br>Probá recargar la página en unos segundos."}</p>`;
  }
})();
