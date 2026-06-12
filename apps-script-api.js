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
};

// ── Punto de entrada GET ───────────────────────────────────────────────
function doGet(e) {
  try {
    const result = {
      ok: true,
      timestamp: new Date().toISOString(),
      elrey: getElReyData(),
      lili_jun26: getLiliKpisData(),
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
  const ss = SpreadsheetApp.openById(SHEETS.liliKpis.id);
  const sheets = ss.getSheets();
  const result = [];

  const MES_MAP = {
    1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
    7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'
  };

  sheets.forEach(sh => {
    const name = sh.getName();
    const rows = sh.getDataRange().getValues();
    // Look for a TOTALES or summary row
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (cleanStr(r[0]).toUpperCase() === 'TOTALES' || cleanStr(r[0]).toUpperCase() === 'TOTAL') {
        // Try to figure out month from sheet name or first date in sheet
        const contactos = cleanNum(r[1]);
        const cotizaciones = cleanNum(r[2]);
        const ventas = cleanNum(r[3]);
        const valor = cleanNum(r[5]) || cleanNum(r[6]);
        if (ventas > 0 || contactos > 0) {
          result.push({ tab: name, contactos, cotizaciones, ventas, valor });
        }
        break;
      }
    }
  });

  return result;
}
