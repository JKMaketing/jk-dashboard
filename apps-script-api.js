/**
 * JK Marketing — Dashboard API
 * Google Apps Script Web App
 *
 * Deploy: Extensiones → Apps Script → Implementar → Nueva implementación
 *         Tipo: App web | Acceso: Cualquier usuario
 *
 * Reemplaza DEPLOY_URL en sales-dashboard.html con la URL generada.
 */

// ── IDs de las hojas ───────────────────────────────────────────────────
const SHEETS = {
  elrey: {
    id: '1puftbarNBpZJu6ELGeZWutnnU_W71b0-yuPQW5q14mk',
    tabJunio: 'Junio',
    tabMarketing: 'Marketing',
    tabInforme: 'Informe_Mes',
  },
  liliKpis: {
    id: '1DQO1nI8ucA8eiEGPKtxCk0PfTBJC3PMWWC8Zu5AWqqk',
  },
  jamundi: {
    id: '1HRqhDS_63PWVutrRCHbEVEFtk0-n_r86ErlBMpAc0KE',
  },
};

// ── Punto de entrada GET ───────────────────────────────────────────────
function doGet(e) {
  try {
    const result = {
      ok: true,
      timestamp: new Date().toISOString(),
      elrey: getElReyData(),
      lili_jun26: getLiliKpisData(),
      jamundi: getJamundiData(),
    };
    return respond(result);
  } catch (err) {
    return respond({ ok: false, error: err.toString(), stack: err.stack });
  }
}

// ── Helpers ────────────────────────────────────────────────────────────
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheetValues(id, tabName) {
  const ss = SpreadsheetApp.openById(id);
  const sh = ss.getSheetByName(tabName);
  if (!sh) return [];
  return sh.getDataRange().getValues();
}

function cleanNum(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  return parseFloat(String(v).replace(/[^0-9.\-]/g, '')) || 0;
}

function cleanStr(v) {
  return String(v || '').trim();
}

// ── El Rey General ─────────────────────────────────────────────────────
function getElReyData() {
  const rows = sheetValues(SHEETS.elrey.id, SHEETS.elrey.tabJunio);
  const informe = sheetValues(SHEETS.elrey.id, SHEETS.elrey.tabInforme);

  const SEDES_KNOWN    = ['CERRITOS', 'CARTAGO', 'JAMUNDÍ'];
  const PRODUCTS_KNOWN = ['GRAN DUQUE', 'EMPERADOR', 'BASE CAMA', 'ESPALDAR',
                          'SALA BOSTON', 'SALA VENECIA', 'COMEDOR FLOR MORADO',
                          'SALA ROMA', 'IMPERIO'];
  const PAGOS_KNOWN    = ['EFECTIVO', 'TRANSFERENCIA', 'DAVIVIENDA', 'BANCOLOMBIA',
                          'BOLD', 'ADDI', 'SISTECREDITO', 'MIXTO', 'CONTRA ENTREGA'];
  const ASESORES_KNOWN = ['EDUARDO', 'NICOL', 'JUAN GABRIEL', 'SARA GIRALDO', 'RONAL'];

  const out = {
    proyeccion: {},
    sedes_actual: [],
    asesores_actual: [],
    productos: [],
    pagos: [],
    diario: [],
    metas_asesores: [],
    metas_sedes: [],
  };

  // ── Scan pestaña Junio ─────────────────────────────────────────────
  let sedesBlock = false;
  let asesoresActualBlock = false;
  let productosBlock = false;
  let pagosBlock = false;
  let diarioBlock = false;
  let proyeccionBlock = false;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const col0 = cleanStr(r[0]).toUpperCase();
    const col1 = cleanStr(r[1]).toUpperCase();

    // Proyección (Fecha corte row)
    if (col0 === 'FECHA CORTE' || (col1 && col1.includes('SOLO CAMBIAR'))) {
      proyeccionBlock = true;
    }
    if (proyeccionBlock) {
      if (col0 === 'DÍAS LABORADOS' || col0 === 'DIAS LABORADOS') {
        out.proyeccion.diasLaborados = cleanNum(r[1]);
      }
      if (col0 === 'TOTAL VENDIDO A LA FECHA') {
        out.proyeccion.totalVentas = cleanNum(r[1]);
      }
      if (col0 === 'TOTAL RECAUDADO A LA FECHA') {
        out.proyeccion.totalRecaudo = cleanNum(r[1]);
      }
      if (col0 === 'PROMEDIO DIARIO VENDIDO') {
        out.proyeccion.promedioDiario = cleanNum(r[1]);
      }
      if (col0 === 'PROYECCIÓN CIERRE VENTAS' || col0 === 'PROYECCION CIERRE VENTAS') {
        out.proyeccion.proyVentas = cleanNum(r[1]);
      }
      if (col0 === 'PROYECCIÓN CIERRE RECAUDO' || col0 === 'PROYECCION CIERRE RECAUDO') {
        out.proyeccion.proyRecaudo = cleanNum(r[1]);
      }
    }

    // Sedes actuales (SEDE | CANTIDAD | PRECIO DE VENTA header)
    if (col0 === 'SEDE' && cleanStr(r[1]).toUpperCase() === 'CANTIDAD') {
      sedesBlock = true; continue;
    }
    if (sedesBlock && SEDES_KNOWN.includes(col0)) {
      out.sedes_actual.push({ name: r[0], cantidad: cleanNum(r[1]), valor: cleanNum(r[2]) });
    }
    if (sedesBlock && !SEDES_KNOWN.includes(col0) && col0 !== '' && col0 !== 'SEDE') {
      sedesBlock = false;
    }

    // Asesores actuales (ASESOR | CANTIDAD | TOTAL header — primera ocurrencia)
    if (col0 === 'ASESOR' && cleanStr(r[1]).toUpperCase() === 'CANTIDAD' && out.asesores_actual.length === 0) {
      asesoresActualBlock = true; continue;
    }
    if (asesoresActualBlock && ASESORES_KNOWN.includes(col0)) {
      out.asesores_actual.push({ name: r[0], cantidad: cleanNum(r[1]), valor: cleanNum(r[2]) });
    }
    if (asesoresActualBlock && !ASESORES_KNOWN.includes(col0) && col0 !== '' && col0 !== 'ASESOR') {
      asesoresActualBlock = false;
    }

    // Metas asesores (ASESOR | TOTAL — segunda ocurrencia: solo 2 columnas)
    if (col0 === 'ASESOR' && cleanStr(r[1]).toUpperCase() === 'TOTAL' && out.metas_asesores.length === 0 && !r[2]) {
      // siguiente bloque son metas
      let j = i + 1;
      while (j < rows.length && ASESORES_KNOWN.includes(cleanStr(rows[j][0]).toUpperCase())) {
        out.metas_asesores.push({ name: rows[j][0], meta: cleanNum(rows[j][1]) });
        j++;
      }
    }

    // Productos (PRODUCTO | CANTIDAD | TOTAL VENDIDO header)
    if (col0 === 'PRODUCTO' && cleanStr(r[1]).toUpperCase() === 'CANTIDAD') {
      productosBlock = true; continue;
    }
    if (productosBlock && PRODUCTS_KNOWN.includes(col0)) {
      out.productos.push({ name: r[0], cantidad: cleanNum(r[1]), valor: cleanNum(r[2]) });
    }
    if (productosBlock && !PRODUCTS_KNOWN.includes(col0) && col0 !== '' && col0 !== 'PRODUCTO') {
      productosBlock = false;
    }

    // Formas de pago (EFECTIVO | valor)
    if (PAGOS_KNOWN.includes(col0)) {
      out.pagos.push({ name: r[0], valor: cleanNum(r[1]) });
    }

    // Ventas diarias (FECHA | TOTAL VENTAS | TOTAL PRECIO VENTA | RECAUDO | PENDIENTE)
    if (col0 === 'FECHA' && cleanStr(r[1]).toUpperCase().includes('TOTAL')) {
      diarioBlock = true; continue;
    }
    if (diarioBlock && r[0] instanceof Date) {
      const fecha = Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'dd/MM');
      const ventas = cleanNum(r[1]);
      if (ventas > 0) {
        out.diario.push({
          fecha,
          ventas,
          precio: cleanNum(r[2]),
          recaudo: cleanNum(r[3]),
          pendiente: cleanNum(r[4]),
        });
      }
    }
  }

  // ── Scan Informe_Mes para metas por sede ───────────────────────────
  for (let i = 0; i < informe.length; i++) {
    const r = informe[i];
    const col0 = cleanStr(r[0]).toUpperCase();
    if (SEDES_KNOWN.includes(col0)) {
      // Solo primera ocurrencia (sede con cantidad + valor = metas)
      if (!out.metas_sedes.find(s => s.name === r[0])) {
        out.metas_sedes.push({ name: r[0], meta_u: cleanNum(r[1]), meta_v: cleanNum(r[2]) });
      }
    }
  }

  // Si metas_sedes vacío, intentar desde el tab Junio (últimas filas)
  if (out.metas_sedes.length === 0) {
    // Las metas de sede aparecen en el bloque final de Junio
    // donde CANTIDAD es grande (>50) y no corresponde a actual
    for (let i = rows.length - 1; i >= 0; i--) {
      const r = rows[i];
      const col0 = cleanStr(r[0]).toUpperCase();
      if (SEDES_KNOWN.includes(col0) && cleanNum(r[1]) > 30) {
        if (!out.metas_sedes.find(s => cleanStr(s.name).toUpperCase() === col0)) {
          out.metas_sedes.unshift({ name: r[0], meta_u: cleanNum(r[1]), meta_v: cleanNum(r[2]) });
        }
      }
    }
  }

  return out;
}

// ── Lili Valencia KPIs Online ──────────────────────────────────────────
function getLiliKpisData() {
  const MES_MAP = {
    1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
    7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'
  };

  const ss = SpreadsheetApp.openById(SHEETS.liliKpis.id);

  // Read MARKETING_DIARIO tab directly — more reliable than formula-driven DASHBOARD
  let sh = null;
  ss.getSheets().forEach(s => {
    const n = s.getName().toUpperCase().replace(/[\s_]/g, '');
    if (n === 'MARKETINGDIARIO' || n === 'MARKETING_DIARIO') sh = s;
  });
  if (!sh) return { daily: [], monthly: {} };

  const rows = sh.getDataRange().getValues();

  // Header: Fecha | Contactos | Cotizaciones | Ventas registradas | Ventas WhatsApp | Ventas Instagram | Valor total de venta | Anotaciones | ...
  let headerRow = -1;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const c0 = cleanStr(r[0]).toUpperCase();
    const c1 = cleanStr(r[1]).toUpperCase();
    if (c0 === 'FECHA' && c1 === 'CONTACTOS') { headerRow = i; break; }
  }
  if (headerRow < 0) return { daily: [], monthly: {} };

  const hdr = rows[headerRow].map(v => cleanStr(v).toUpperCase());
  const iC   = hdr.indexOf('CONTACTOS');
  const iQ   = hdr.indexOf('COTIZACIONES');
  // "Ventas registradas" = total ventas (col 3); also try plain "VENTAS"
  let iV = hdr.findIndex((h,i) => i >= 3 && (h === 'VENTAS' || h.startsWith('VENTAS R')));
  if (iV < 0) iV = 3;
  // "Valor total de venta" (col 6); also try plain "VALOR"
  let iVal = hdr.findIndex((h,i) => i >= 4 && (h === 'VALOR' || h.startsWith('VALOR T') || h.startsWith('VALOR TOTAL')));
  if (iVal < 0) iVal = 6;
  const iWA  = hdr.findIndex(h => h.includes('WHATSAPP'));
  const iIG  = hdr.findIndex(h => h.includes('INSTAGRAM'));

  function parseFechaLili(raw) {
    if (raw instanceof Date) return raw;
    const s = String(raw).trim();
    // dd/mm/yyyy
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return new Date(parseInt(m[3]), parseInt(m[2])-1, parseInt(m[1]));
    return new Date(s);
  }

  const daily = [];
  let totC = 0, totQ = 0, totV = 0, totVal = 0;
  let mesLabel = '';

  for (let i = headerRow + 1; i < rows.length; i++) {
    const r = rows[i];
    const rawFecha = r[0];
    if (!rawFecha) continue;

    const dateObj = parseFechaLili(rawFecha);
    if (isNaN(dateObj)) continue;

    const c   = cleanNum(r[iC]);
    const q   = cleanNum(r[iQ]);
    const v   = cleanNum(r[iV]);
    const val = cleanNum(r[iVal]);

    // Skip rows with no data at all
    if (c === 0 && v === 0 && val === 0) continue;

    const fechaStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM');
    if (!mesLabel) {
      const mo = dateObj.getMonth() + 1;
      const y  = String(dateObj.getFullYear()).slice(2);
      mesLabel = MES_MAP[mo] + ' ' + y;
    }

    daily.push({
      fecha:       fechaStr,
      contactos:   c,
      cotizaciones:q,
      ventas:      v,
      valor:       val,
      whatsapp:    iWA >= 0 ? cleanNum(r[iWA]) : null,
      instagram:   iIG >= 0 ? cleanNum(r[iIG]) : null,
    });

    totC   += c;
    totQ   += q;
    totV   += v;
    totVal += val;
  }

  return {
    daily,
    monthly: {
      mes:          mesLabel,
      contactos:    totC,
      cotizaciones: totQ,
      ventas:       totV,
      valor:        totVal,
    }
  };
}

// ── Jamundí Imperial ───────────────────────────────────────────────────
function getJamundiData() {
  const MES_MAP = {
    1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
    7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'
  };
  const ASESORES_KNOWN = ['EDUARDO','NICOL','JUAN GABRIEL','SARA GIRALDO','RONAL'];
  const SEDES_KNOWN    = ['CERRITOS','CARTAGO','JAMUNDÍ'];
  const PRODUCTS_KNOWN = ['GRAN DUQUE','EMPERADOR','BASE CAMA','ESPALDAR','SALA BOSTON',
                          'SALA VENECIA','COMEDOR FLOR MORADO','SALA ROMA','IMPERIO'];
  const PAGOS_KNOWN    = ['EFECTIVO','TRANSFERENCIA','DAVIVIENDA','BANCOLOMBIA',
                          'BOLD','ADDI','SISTECREDITO','MIXTO','CONTRA ENTREGA'];

  const ss = SpreadsheetApp.openById(SHEETS.jamundi.id);

  // Buscar pestañas MARKETING e INFORME_MES
  let mktSh = null, informeSh = null;
  ss.getSheets().forEach(s => {
    const n = s.getName().toUpperCase();
    if (n === 'MARKETING') mktSh = s;
    if (n.includes('INFORME')) informeSh = s;
  });

  let mesLabel = '';
  const asesoresMap = {};

  function parseFecha(raw) {
    if (raw instanceof Date) return raw;
    const s = cleanStr(raw);
    const p = s.split('/');
    if (p.length === 3) return new Date(p[2], parseInt(p[1])-1, parseInt(p[0]));
    return new Date(s);
  }

  // ── Lee pestaña MARKETING: bloques por asesor ───────────────────────
  if (mktSh) {
    const rows = mktSh.getDataRange().getValues();
    for (let i = 0; i < rows.length; i++) {
      const r  = rows[i];
      const c0 = cleanStr(r[0]).toUpperCase();
      const c1 = cleanStr(r[1]).toUpperCase();
      const c2 = cleanStr(r[2]).toUpperCase();
      const c3 = cleanStr(r[3]).toUpperCase();
      if (c0 === 'FECHA' && c1.startsWith('CONTACTOS') && c2.startsWith('COTIZACIONES') && c3.startsWith('VENTAS')) {
        const asesor = c1.replace('CONTACTOS','').trim() || 'UNKNOWN';
        if (!asesoresMap[asesor]) asesoresMap[asesor] = [];
        for (let j = i + 1; j < rows.length; j++) {
          const dr = rows[j];
          if (!dr[0]) continue;
          if (cleanStr(dr[0]).toUpperCase() === 'FECHA') break;
          const dateObj = parseFecha(dr[0]);
          if (isNaN(dateObj)) break;
          const c   = cleanNum(dr[1]);
          const q   = cleanNum(dr[2]);
          const v   = cleanNum(dr[3]);
          const val = cleanNum(dr[4]);
          if (c === 0 && v === 0 && val === 0) continue;
          const key = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM');
          if (!mesLabel) mesLabel = MES_MAP[dateObj.getMonth()+1] + ' ' + String(dateObj.getFullYear()).slice(2);
          asesoresMap[asesor].push({fecha:key, contactos:c, cotizaciones:q, ventas:v, valor:val});
        }
      }
    }
  }

  // Agrega totales diarios de todos los asesores
  const dailyTotals = {};
  Object.values(asesoresMap).forEach(rows => {
    rows.forEach(r => {
      if (!dailyTotals[r.fecha]) dailyTotals[r.fecha] = {contactos:0,cotizaciones:0,ventas:0,valor:0};
      dailyTotals[r.fecha].contactos    += r.contactos;
      dailyTotals[r.fecha].cotizaciones += r.cotizaciones;
      dailyTotals[r.fecha].ventas       += r.ventas;
      dailyTotals[r.fecha].valor        += r.valor;
    });
  });

  const sortDates = obj => Object.keys(obj).sort((a,b) => {
    const [da,ma] = a.split('/').map(Number);
    const [db,mb] = b.split('/').map(Number);
    return (ma*100+da)-(mb*100+db);
  });

  const daily = sortDates(dailyTotals).map(f => ({fecha:f, ...dailyTotals[f]}));

  const asesores_mkt = ASESORES_KNOWN
    .filter(name => asesoresMap[name] && asesoresMap[name].length > 0)
    .map(name => ({
      name,
      contactos:    asesoresMap[name].reduce((s,r)=>s+r.contactos,0),
      cotizaciones: asesoresMap[name].reduce((s,r)=>s+r.cotizaciones,0),
      ventas:       asesoresMap[name].reduce((s,r)=>s+r.ventas,0),
      valor:        asesoresMap[name].reduce((s,r)=>s+r.valor,0),
      daily:        asesoresMap[name],
    }));

  // ── Lee pestaña INFORME_MES ──────────────────────────────────────────
  const out = {
    daily, asesores_mkt,
    monthly: {mes: mesLabel || 'Jun 26'},
    asesores_actual: [], metas_asesores: [],
    proyeccion: {}, diario_ventas: [],
    productos: [], pagos: [],
    sedes_actual: [], metas_sedes: [],
  };

  if (informeSh) {
    const rows      = informeSh.getDataRange().getValues();
    let asesActDone = false, metasAsDone   = false;
    let sedesActDone= false, metasSedDone  = false;
    let diarioBlock = false, prodBlock     = false;

    for (let i = 0; i < rows.length; i++) {
      const r  = rows[i];
      const c0 = cleanStr(r[0]).toUpperCase();
      const c1 = cleanStr(r[1]).toUpperCase();
      const c2 = cleanStr(r[2]).toUpperCase();

      // Tabla diaria
      if (c0 === 'FECHA' && c1.includes('TOTAL') && !diarioBlock) { diarioBlock = true; continue; }
      if (diarioBlock) {
        if (c0 === 'FECHA' || c0 === 'FECHA CORTE') { diarioBlock = false; }
        else if (r[0]) {
          const dateObj = r[0] instanceof Date ? r[0] : parseFecha(r[0]);
          if (!isNaN(dateObj)) {
            const v = cleanNum(r[1]);
            if (v > 0) out.diario_ventas.push({
              fecha: Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM'),
              ventas: v, precio: cleanNum(r[2]), recaudo: cleanNum(r[3]), pendiente: cleanNum(r[4]),
            });
          } else diarioBlock = false;
        }
        continue;
      }

      // Proyección
      if (c0 === 'DÍAS LABORADOS' || c0 === 'DIAS LABORADOS') out.proyeccion.diasLaborados  = cleanNum(r[1]);
      if (c0 === 'TOTAL VENDIDO A LA FECHA')                   out.proyeccion.totalVentas    = cleanNum(r[1]);
      if (c0 === 'TOTAL RECAUDADO A LA FECHA')                 out.proyeccion.totalRecaudo   = cleanNum(r[1]);
      if (c0 === 'PROMEDIO DIARIO VENDIDO')                    out.proyeccion.promedioDiario = cleanNum(r[1]);
      if (c0.includes('PROYECCI') && c0.includes('VENTAS'))   out.proyeccion.proyVentas     = cleanNum(r[1]);
      if (c0.includes('PROYECCI') && c0.includes('RECAUDO'))  out.proyeccion.proyRecaudo    = cleanNum(r[1]);

      // Productos
      if (c0 === 'PRODUCTO' && c1 === 'CANTIDAD') { prodBlock = true; continue; }
      if (prodBlock && PRODUCTS_KNOWN.includes(c0)) out.productos.push({name:r[0], cantidad:cleanNum(r[1]), valor:cleanNum(r[2])});
      if (prodBlock && !PRODUCTS_KNOWN.includes(c0) && c0 !== '') prodBlock = false;

      // Pagos
      if (PAGOS_KNOWN.includes(c0)) out.pagos.push({name:r[0], valor:cleanNum(r[1])});

      // Asesores actual: ASESOR | CANTIDAD | TOTAL (3 cols)
      if (c0 === 'ASESOR' && c1 === 'CANTIDAD' && c2 === 'TOTAL' && !asesActDone) {
        let j = i + 1;
        while (j < rows.length && ASESORES_KNOWN.includes(cleanStr(rows[j][0]).toUpperCase())) {
          out.asesores_actual.push({name:rows[j][0], cantidad:cleanNum(rows[j][1]), valor:cleanNum(rows[j][2])});
          j++;
        }
        asesActDone = true;
      }

      // Metas asesores: ASESOR | TOTAL (2 cols, ≥ 1M)
      if (c0 === 'ASESOR' && c1 === 'TOTAL' && !cleanStr(r[2]) && !metasAsDone) {
        const temp = [];
        let j = i + 1;
        while (j < rows.length && ASESORES_KNOWN.includes(cleanStr(rows[j][0]).toUpperCase())) {
          const val = cleanNum(rows[j][1]);
          if (val >= 1000000) temp.push({name:rows[j][0], meta:val});
          j++;
        }
        if (temp.length) { out.metas_asesores = temp; metasAsDone = true; }
      }

      // Sedes actual (primera vez)
      if (c0 === 'SEDE' && c1 === 'CANTIDAD' && !sedesActDone) {
        let j = i + 1;
        while (j < rows.length && SEDES_KNOWN.includes(cleanStr(rows[j][0]).toUpperCase())) {
          out.sedes_actual.push({name:rows[j][0], cantidad:cleanNum(rows[j][1]), valor:cleanNum(rows[j][2])});
          j++;
        }
        sedesActDone = true;
      }

      // Metas sedes (segunda aparición con valores grandes)
      if (c0 === 'SEDE' && c1 === 'CANTIDAD' && sedesActDone && !metasSedDone) {
        const temp = [];
        let j = i + 1;
        while (j < rows.length && SEDES_KNOWN.includes(cleanStr(rows[j][0]).toUpperCase())) {
          const mu = cleanNum(rows[j][1]), mv = cleanNum(rows[j][2]);
          if (mu > 10 || mv > 10000000) temp.push({name:rows[j][0], meta_u:mu, meta_v:mv});
          j++;
        }
        if (temp.length) { out.metas_sedes = temp; metasSedDone = true; }
      }
    }
  }

  const totC   = daily.reduce((s,r)=>s+r.contactos,0);
  const totQ   = daily.reduce((s,r)=>s+r.cotizaciones,0);
  const totV   = out.diario_ventas.reduce((s,r)=>s+r.ventas,0) || daily.reduce((s,r)=>s+r.ventas,0);
  const totVal = out.diario_ventas.reduce((s,r)=>s+r.precio,0) || daily.reduce((s,r)=>s+r.valor,0);
  out.monthly  = {mes: mesLabel || 'Jun 26', contactos: totC, cotizaciones: totQ, ventas: totV, valor: totVal};

  return out;
}
