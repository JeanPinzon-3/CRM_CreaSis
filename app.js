// =======================================================================
// Panel de Operaciones — lógica de la aplicación
// =======================================================================

let registros = [];
let nextNum = 1;
const el = id => document.getElementById(id);

function genId(){ return 'REG-' + String(nextNum++).padStart(5,'0'); }

function normalizeHeader(h){
  return String(h).toLowerCase().replace(/[\n\r]+/g,' ').replace(/\s+/g,' ').trim();
}

function excelSerialToDate(n){
  // Convierte un número serial de Excel (celda de fecha con formato no
  // estándar que SheetJS no reconoció como Date) a un objeto Date.
  const utcDays = Math.floor(n - 25569);
  const utcValue = utcDays * 86400;
  return new Date(utcValue * 1000);
}

function toDateString(v){
  if(v === undefined || v === null || v === '') return '';
  if(v instanceof Date && !isNaN(v)) return v.toISOString().slice(0,10);
  // Número serial de Excel (rango razonable: años ~1950-2100) que no fue
  // convertido a Date automáticamente por algún formato de celda atípico.
  if(typeof v === 'number' && v > 18000 && v < 73050){
    const d = excelSerialToDate(v);
    if(!isNaN(d)) return d.toISOString().slice(0,10);
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m) return m[0];
  const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(m2) return `${m2[3]}-${String(m2[1]).padStart(2,'0')}-${String(m2[2]).padStart(2,'0')}`;
  return s.slice(0,10);
}

function parseNumberLoose(v){
  // Acepta "12", "12.5", "12,5" (coma decimal) y "1.234,56" / "1,234.56"
  // (separador de miles) sin confundirlos.
  if(typeof v === 'number') return v;
  let s = String(v).trim();
  if(s === '') return NaN;
  const hasComma = s.includes(','), hasDot = s.includes('.');
  if(hasComma && hasDot){
    // El último separador que aparece es el decimal; el otro es de miles.
    if(s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g,'').replace(',', '.');
    else s = s.replace(/,/g,'');
  } else if(hasComma){
    s = s.replace(',', '.');
  }
  return parseFloat(s);
}

function inferTipo(sheetName, fileName){
  const s = (sheetName + ' ' + fileName).toLowerCase();
  if(s.includes('falla')) return 'Falla de Job';
  if(s.includes('control_pasos') || s.includes('paso')) return 'Despliegue';
  if(s.includes('cierre')) return 'Cierre Mensual';
  if(s.includes('graffana') || s.includes('grafana')) return 'Incidente Monitoreo';
  if(s.includes('abs')) return 'Incidente ABS';
  if(s.includes('version')) return 'Versionamiento';
  return 'Registro';
}

const LOG_KEYWORDS = ['fecha','servidor','instancia','job','proceso','responsable','operador',
  'estado','resultado','diagnostico','escalamiento','cliente','incidente','reportado',
  'asunto','tiempo','solicitud','observaciones','aplicaci'];

function countKeywordMatches(headerTexts){
  const headerSet = new Set(headerTexts.map(normalizeHeader).filter(Boolean));
  let matches = 0;
  for(const kw of LOG_KEYWORDS){
    for(const h of headerSet){ if(h.includes(kw)){ matches++; break; } }
  }
  return matches;
}

// Convierte una hoja de SheetJS en una lista de objetos fila->valor,
// resolviendo dos problemas comunes de Excel que hacían que se perdiera
// información al importar:
//
// 1) CELDAS COMBINADAS: SheetJS solo pone el valor en la celda superior-
//    izquierda del rango combinado; el resto queda vacío. Aquí se "rellena"
//    el valor combinado en todas las celdas del rango antes de armar las
//    filas, para que no se pierdan fechas/servidores que abarcan varias
//    filas.
// 2) ENCABEZADO NO ESTÁ EN LA FILA 1: si hay filas de título/logo antes de
//    los encabezados reales, se busca entre las primeras 15 filas cuál es
//    la que más coincide con palabras clave de bitácora, en vez de asumir
//    siempre la fila 1.
// 3) ENCABEZADOS DUPLICADOS: si dos columnas tienen el mismo nombre, se
//    distinguen agregando la letra de columna de Excel, para que ninguna
//    se sobrescriba.
function sheetToRows(sheet){
  const aoa = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''});
  if(!aoa.length) return { rows:[], headerRow:-1 };

  // Rellenar celdas combinadas con el valor de la celda superior-izquierda.
  const merges = sheet['!merges'] || [];
  merges.forEach(m=>{
    if(!aoa[m.s.r]) return;
    const val = aoa[m.s.r][m.s.c];
    if(val === undefined || val === '') return;
    for(let r=m.s.r; r<=m.e.r; r++){
      if(!aoa[r]) aoa[r] = [];
      for(let c=m.s.c; c<=m.e.c; c++){
        if(aoa[r][c] === undefined || aoa[r][c] === '') aoa[r][c] = val;
      }
    }
  });

  // Buscar la fila de encabezado real entre las primeras 15 filas.
  // Se exige que existan al menos 2 coincidencias de palabras clave Y que
  // una de ellas sea "fecha": esto evita que hojas de referencia/listas
  // desplegables (que reutilizan nombres de columna como RESULTADO,
  // RESPONSABLE, AMBIENTE, pero no tienen fecha) se confundan con una
  // bitácora real.
  const searchLimit = Math.min(aoa.length, 15);
  let headerRow = 0, bestScore = -1;
  for(let i=0; i<searchLimit; i++){
    const texts = (aoa[i]||[]).map(v=>String(v||''));
    const hasFecha = texts.some(t=>normalizeHeader(t).includes('fecha'));
    const score = hasFecha ? countKeywordMatches(texts) : -1;
    if(score > bestScore){ bestScore = score; headerRow = i; }
  }
  if(bestScore < 2) return { rows:[], headerRow:-1 }; // no parece bitácora

  const headerCells = aoa[headerRow] || [];
  const seen = new Map();
  const headers = headerCells.map((h,c)=>{
    let name = String(h||'').trim();
    if(!name) name = 'Columna ' + XLSX.utils.encode_col(c);
    if(seen.has(name)){
      const n = seen.get(name) + 1;
      seen.set(name, n);
      name = name + ' (col ' + XLSX.utils.encode_col(c) + ')';
    } else {
      seen.set(name, 0);
    }
    return name;
  });

  const CORE_KEYWORDS = ['fecha','nombre aplicaci','nombre job','nombre reporte','asunto'];
  const rows = [];
  for(let r=headerRow+1; r<aoa.length; r++){
    const line = aoa[r];
    if(!line) continue;
    const obj = {};
    let hasCoreContent = false;
    headers.forEach((h,c)=>{
      const v = line[c] !== undefined ? line[c] : '';
      obj[h] = v;
      if(String(v).trim() !== ''){
        const nh = normalizeHeader(h);
        if(CORE_KEYWORDS.some(kw=>nh.includes(kw))) hasCoreContent = true;
      }
    });
    // Filas de plantilla sin usar (comunes en formatos de Excel prellenados
    // con valores por defecto como "0" en columnas de conteo) se descartan:
    // solo se importan filas que sí tengan fecha o un proceso/job real.
    if(hasCoreContent) rows.push(obj);
  }
  return { rows, headerRow };
}

// Convierte una fila cruda del Excel (objeto columna->valor tal como la
// entrega SheetJS) en un registro del panel. Cualquier columna que no
// encaje en los campos conocidos (incluyendo listas desplegables de
// selección múltiple con nombres personalizados) se conserva íntegra en
// `extra`, y la fila completa sin tocar se guarda en `raw` para poder
// inspeccionarla con el botón "Ver original".
function rowToRegistro(row, sheetName, fileName){
  const usedKeys = new Set();

  function pick(keywords){
    const keys = Object.keys(row);
    for(const kw of keywords){
      for(const k of keys){
        if(usedKeys.has(k)) continue;
        if(normalizeHeader(k).includes(kw)){
          const v = row[k];
          if(v !== undefined && v !== null && String(v).trim() !== ''){
            usedKeys.add(k);
            return v;
          }
        }
      }
    }
    return '';
  }

  const fecha = toDateString(pick(['fecha imp','fecha sol','fecha']));
  const proceso = String(pick(['nombre aplicaci','nombre job','nombre reporte','asunto','proceso','job']) || 'Sin nombre');
  const servidor = String(pick(['servidor','ambiente','instancia']));
  const responsable = String(pick(['responsable','operador','resp.','reportado por']));

  const estadoRaw = String(pick(['estado final','estado','resultado']));
  const estadoLower = estadoRaw.toLowerCase();
  let estado = 'Otro';
  if(estadoLower.includes('exito')) estado = 'Exitoso';
  else if(estadoLower.includes('fall')) estado = 'Fallido';
  else if(estadoLower.includes('pendient')) estado = 'Pendiente';

  // Escalamiento: puede haber varias columnas "ESCALAMIENTO..."; se marcan
  // todas como usadas y basta que UNA tenga SI para marcar el registro
  // como escalado.
  let escalamiento = 'No';
  Object.keys(row).forEach(k=>{
    if(usedKeys.has(k)) return;
    if(normalizeHeader(k).includes('escalamiento')){
      usedKeys.add(k);
      const v = String(row[k]).trim().toLowerCase();
      if(v === 'si' || v === 'sí') escalamiento = 'Sí';
    }
  });

  // Tiempo: se prioriza calcular la duración real a partir de columnas de
  // "hora inicio" / "hora final" (ej. H. INICIO / H. FINAL), en vez de
  // confiar en una columna de fórmula tipo "T.DESPLIEGUE": en archivos
  // reales esa columna suele venir desactualizada o sin fórmula en algunas
  // filas, mientras que inicio/final casi siempre son datos capturados a
  // mano y por lo tanto más confiables.
  let tiempo = '';
  {
    const keys = Object.keys(row);
    let inicioKey = null, finalKey = null;
    for(const k of keys){
      const h = normalizeHeader(k);
      if(inicioKey===null && (h.includes('hora inicio') || h.includes('h. inicio') || h.includes('h.inicio') || /(^|\s)inicio(\s|$)/.test(h))) inicioKey = k;
      if(finalKey===null && (h.includes('hora final') || h.includes('h. final') || h.includes('h.final') || /(^|\s)final(\s|$)/.test(h) || /(^|\s)termino(\s|$)/.test(h))) finalKey = k;
    }
    // Se marcan como usadas siempre (aunque el cálculo falle) para que no
    // terminen colándose como "actividad" numérica más adelante.
    if(inicioKey) usedKeys.add(inicioKey);
    if(finalKey) usedKeys.add(finalKey);
    if(inicioKey && finalKey){
      const vIni = parseNumberLoose(row[inicioKey]);
      const vFin = parseNumberLoose(row[finalKey]);
      if(!isNaN(vIni) && !isNaN(vFin) && vFin >= vIni){
        tiempo = vFin - vIni;
      }
    }
  }
  // Cualquier columna que hable de tiempo/duración (ej. "T.DESPLIEGUE") se
  // excluye de "actividades" siempre, se haya usado o no para calcular el
  // tiempo final, para que no aparezca como si fuera un conteo de tareas.
  Object.keys(row).forEach(k=>{
    if(usedKeys.has(k)) return;
    const h = normalizeHeader(k);
    if(h.includes('tiempo') || h.includes('despliegue') || h.includes('duracion') || h.includes('demora') || h.includes('duración')){
      if(tiempo === ''){
        const v = parseNumberLoose(row[k]);
        if(!isNaN(v)) tiempo = v;
      }
      usedKeys.add(k);
    }
  });

  const tipoDirect = pick(['tipo solicitud','tipo']);
  const tipo = tipoDirect ? String(tipoDirect) : inferTipo(sheetName, fileName);
  const programado = String(pick(['programado']));

  const diagnostico = String(pick(['diagnostico','observaciones','accion a tomar']));
  const accion = String(pick(['respuesta escalamiento','accion','resultado']));

  // Actividades: columnas de conteo tipo checkbox (ETL, SCRIPT, Modificacion
  // Job, Creacion Job, Restauracion BD...) que traen un número (0,1,2...)
  // indicando cuántas veces se hizo esa actividad en el registro. Se
  // guardan aparte (no como texto libre) para poder sumarlas y filtrarlas
  // en la pestaña "Actividades por tipo". La clave se normaliza aquí UNA
  // sola vez con cleanLabel (en vez de en cada render): si dos columnas
  // del Excel limpian al mismo nombre (ej. con distintos saltos de línea),
  // sus valores se suman directamente.
  const actividades = {};
  Object.keys(row).forEach(k=>{
    if(usedKeys.has(k)) return;
    const v = row[k];
    if(v === undefined || v === null) return;
    const s = String(v).trim();
    if(s === '') return;
    if(/^-?\d+([.,]\d+)?$/.test(s)){
      const num = parseNumberLoose(s);
      if(!isNaN(num)){
        const label = cleanLabel(k);
        actividades[label] = (actividades[label] || 0) + num;
        usedKeys.add(k);
      }
    }
  });

  // Cualquier columna que no se haya usado (texto libre, listas de
  // selección múltiple, campos personalizados, etc.) se conserva aquí
  // para no perder información.
  const extraParts = [];
  Object.keys(row).forEach(k=>{
    if(usedKeys.has(k)) return;
    const v = row[k];
    if(v !== undefined && v !== null && String(v).trim() !== ''){
      extraParts.push(String(k).trim() + ': ' + String(v).trim());
    }
  });

  return {
    id: genId(), fecha, proceso, servidor, responsable, estado, escalamiento, tiempo, tipo, programado,
    diagnostico, accion, extra: extraParts.join(' | '), actividades,
    origen: fileName + ' · ' + sheetName,
    raw: row, // fila 100% original, sin procesar, tal como la entregó SheetJS
  };
}

function estadoClass(estado){
  const map = {'Exitoso':'b-exitoso','Fallido':'b-fallido','Pendiente':'b-pendiente'};
  return map[estado] || 'b-otro';
}

const _escapeScratch = document.createElement('div');
function escapeHtml(str){
  _escapeScratch.textContent = str==null ? '' : String(str);
  return _escapeScratch.innerHTML;
}

function uniqueValues(field){
  return [...new Set(registros.map(r=>r[field]).filter(Boolean))].sort();
}

function refreshFilterOptions(){
  const tipoSel = el('filterTipo'), origenSel = el('filterOrigen');
  const curTipo = tipoSel.value, curOrigen = origenSel.value;
  tipoSel.innerHTML = '<option value="">Todos los tipos</option>' +
    uniqueValues('tipo').map(t=>`<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`).join('');
  origenSel.innerHTML = '<option value="">Todos los archivos</option>' +
    uniqueValues('origen').map(o=>`<option value="${escapeHtml(o)}">${escapeHtml(o)}</option>`).join('');
  tipoSel.value = curTipo; origenSel.value = curOrigen;
}

function getFilteredRegistros(){
  const search = el('search').value.trim().toLowerCase();
  const fEstado = el('filterEstado').value;
  const fTipo = el('filterTipo').value;
  const fOrigen = el('filterOrigen').value;
  return registros.filter(r=>{
    const matchSearch = !search || [r.proceso,r.servidor,r.responsable].join(' ').toLowerCase().includes(search);
    const matchEstado = !fEstado || r.estado===fEstado;
    const matchTipo = !fTipo || r.tipo===fTipo;
    const matchOrigen = !fOrigen || r.origen===fOrigen;
    return matchSearch && matchEstado && matchTipo && matchOrigen;
  });
}

function activeFilterLabels(){
  const labels = [];
  if(el('search').value.trim()) labels.push('Búsqueda: "' + el('search').value.trim() + '"');
  if(el('filterEstado').value) labels.push('Estado: ' + el('filterEstado').value);
  if(el('filterTipo').value) labels.push('Tipo: ' + el('filterTipo').value);
  if(el('filterOrigen').value) labels.push('Archivo: ' + el('filterOrigen').value);
  return labels;
}

function renderStats(list){
  const total = list.length;
  const totalGeneral = registros.length;
  const fallidos = list.filter(r=>r.estado==='Fallido').length;
  const exitosos = list.filter(r=>r.estado==='Exitoso').length;
  const escalados = list.filter(r=>r.escalamiento==='Sí').length;
  const tiempos = list.map(r=>r.tiempo).filter(t=>typeof t === 'number');
  const promedio = tiempos.length ? Math.round(tiempos.reduce((a,b)=>a+b,0)/tiempos.length) : 0;
  const filtrado = total !== totalGeneral;

  el('stats').innerHTML = `
    <div class="stat total"><div class="n">${total}</div><div class="l">${filtrado ? 'Registros (filtrado)' : 'Total registros'}</div></div>
    <div class="stat fallido"><div class="n">${fallidos}</div><div class="l">Fallidos</div></div>
    <div class="stat exitoso"><div class="n">${exitosos}</div><div class="l">Exitosos</div></div>
    <div class="stat escalado"><div class="n">${escalados}</div><div class="l">Con escalamiento</div></div>
    <div class="stat tiempo"><div class="n">${promedio}</div><div class="l">Min. promedio</div></div>
  `;
}

function renderTable(list){
  if(list.length===0){
    el('tableWrap').innerHTML = `<div class="empty"><b>No hay registros para mostrar</b>Importa uno o varios Excel, crea un registro nuevo, o ajusta los filtros.</div>`;
    return;
  }

  const rows = list.map(r=>`
    <tr>
      <td class="mono-cell" data-label="ID">${r.id}</td>
      <td data-label="Fecha">${r.fecha||'—'}</td>
      <td data-label="Proceso/Job"><strong>${escapeHtml(r.proceso)}</strong></td>
      <td data-label="Servidor">${escapeHtml(r.servidor||'—')}</td>
      <td data-label="Responsable">${escapeHtml(r.responsable||'—')}</td>
      <td data-label="Tipo">${escapeHtml(r.tipo)}</td>
      <td data-label="Estado"><span class="badge ${estadoClass(r.estado)}">${r.estado}</span></td>
      <td data-label="Escalamiento"><span class="${r.escalamiento==='Sí'?'esc-si':'esc-no'}">${r.escalamiento}</span></td>
      <td data-label="Tiempo (min)">${r.tiempo!==''?r.tiempo:'—'}</td>
      <td data-label="Info. adicional" class="cell-desc" title="${escapeHtml(r.extra||'')}">${r.extra ? escapeHtml(r.extra.slice(0,60)) + (r.extra.length>60?'…':'') : '—'}</td>
      <td data-label="Origen" class="cell-origen">${escapeHtml(r.origen)}</td>
      <td data-label="Acciones">
        <div class="row-actions">
          <button onclick="viewRaw('${r.id}')">Ver original</button>
          <button onclick="editRegistro('${r.id}')">Editar</button>
          <button onclick="deleteRegistro('${r.id}')">Eliminar</button>
        </div>
      </td>
    </tr>
  `).join('');

  el('tableWrap').innerHTML = `
    <table>
      <thead>
        <tr>
          <th>ID</th><th>Fecha</th><th>Proceso/Job</th><th>Servidor</th><th>Responsable</th>
          <th>Tipo</th><th>Estado</th><th>Escalamiento</th><th>Tiempo</th><th>Info. adicional</th><th>Origen</th><th></th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

// --- Gráficos con Chart.js: cada panel guarda su tipo elegido (barras
// horizontales/verticales o torta) y los últimos datos, para poder
// redibujar al instante cuando el usuario cambia el tipo sin recalcular. ---
if(typeof ChartDataLabels !== 'undefined'){ Chart.register(ChartDataLabels); }

const chartState = {
  // Estos dos arrancan en barras verticales para que se parezcan a los
  // gráficos de referencia (conteo por ambiente / por tipo de solicitud).
  chartAmbienteConteo: { type:'vbar' },
  chartTipoSolicitud: { type:'vbar' },
  // Estos arrancan en horizontales porque los nombres de servicio son
  // largos y se leen mejor en filas que en columnas angostas.
  chartServicioActividad: { type:'hbar' },
  chartServicioProgramado: { type:'hbar' },
};
const PIE_PALETTE = ['#4FD1C5','#5B8DEF','#F5A524','#E5484D','#3DD68C','#9F7AEA','#F56565','#38B2AC','#ED8936','#667EEA','#48BB78','#ECC94B','#FC8181','#4299E1','#B794F4'];
const COLOR_MAP = { 'c-danger':'#E5484D', 'c-info':'#5B8DEF', 'c-warn':'#F5A524', '':'#4FD1C5' };
const SERIES_COLORS = ['#5B8DEF', '#1F3A93', '#3DD68C', '#F5A524'];

function cleanLabel(s){
  return String(s).replace(/[\r\n]+/g,' ').replace(/\s+/g,' ').trim();
}

function renderChartData(id, items, opts={}){
  chartState[id] = chartState[id] || { type:'hbar' };
  chartState[id].grouped = false;
  chartState[id].items = items;
  chartState[id].opts = opts;
  drawChart(id);
}

// Gráfico de barras agrupadas (varias series por categoría), ej.
// "Solicitudes por tipo" separado por Programado / No Programado.
function renderGroupedChart(id, categories, series, opts={}){
  chartState[id] = chartState[id] || { type:'vbar' };
  chartState[id].grouped = true;
  chartState[id].categories = categories;
  chartState[id].series = series;
  chartState[id].opts = opts;
  drawChart(id);
}

function datalabelsFor(suffix, color){
  return {
    color: color || '#E8ECF1',
    font: { size:10, weight:'600' },
    formatter: v => (v===0 ? '' : v + (suffix||'')),
  };
}

// Llena la tabla lateral de "Etiqueta | Cantidad | %" que acompaña a cada
// gráfico, para poder ver el desglose exacto (y el porcentaje sobre el
// total) sin depender de pasar el mouse sobre las barras.
function renderSideTable(id, items, opts={}){
  const table = document.getElementById(id + '-table');
  if(!table) return;
  if(!items || !items.length){ table.innerHTML = ''; return; }
  const suffix = opts.suffix || '';
  const total = items.reduce((a,b)=>a + (b.value||0), 0);
  const rows = items.map(i=>{
    const pct = total ? (i.value/total*100) : 0;
    return `<tr><td title="${escapeHtml(i.label)}">${escapeHtml(i.label)}</td><td class="num">${i.value}${suffix}</td><td class="num">${pct.toFixed(1)}%</td></tr>`;
  }).join('');
  table.innerHTML =
    '<thead><tr><th>' + (opts.labelHeader || 'Etiqueta') + '</th><th class="num">' + (opts.valueHeader || 'Cant.') + '</th><th class="num">%</th></tr></thead>' +
    '<tbody>' + rows + '<tr class="total-row"><td>Total</td><td class="num">' + total + suffix + '</td><td class="num">100%</td></tr></tbody>';
}

function drawChart(id){
  const state = chartState[id];
  if(!state) return;
  const canvas = document.getElementById(id);
  const emptyEl = document.getElementById(id + '-empty');
  if(!canvas) return;
  if(emptyEl && emptyEl.dataset.defaultText === undefined){ emptyEl.dataset.defaultText = emptyEl.textContent; }
  if(state.instance){ try{ state.instance.destroy(); }catch(e){} state.instance = null; }

  try{
    const gridColor = '#2A3340', tickColor = '#8894A3', textColor = '#E8ECF1';
    const opts = state.opts || {};
    const suffix = opts.suffix || '';

    if(state.grouped){
      const categories = state.categories || [];
      const series = state.series || [];
      const hasData = categories.length && series.some(s=>(s.data||[]).some(v=>v>0));
      if(!hasData){
        canvas.style.display = 'none';
        if(emptyEl){ emptyEl.style.display = 'block'; emptyEl.textContent = emptyEl.dataset.defaultText; }
        renderSideTable(id, []);
        return;
      }
      canvas.style.display = 'block';
      if(emptyEl) emptyEl.style.display = 'none';
      renderSideTable(
        id,
        categories.map((c,i)=>({ label:c, value: series.reduce((acc,s)=>acc+(s.data[i]||0),0) })),
        { labelHeader: opts.labelHeader || 'Tipo', valueHeader: 'Total' }
      );

      let cfg;
      if(state.type === 'pie'){
        const totals = categories.map((c,i)=> series.reduce((acc,s)=>acc+(s.data[i]||0),0));
        cfg = {
          type:'pie',
          data:{ labels: categories, datasets:[{ data: totals, backgroundColor: categories.map((_,i)=>PIE_PALETTE[i % PIE_PALETTE.length]), borderColor:'#171D25', borderWidth:2 }] },
          options:{
            animation:false, maintainAspectRatio:false,
            plugins:{
              legend:{ position:'right', labels:{ color:textColor, boxWidth:11, font:{size:10.5}, padding:8 } },
              tooltip:{ callbacks:{ label: ctx => ctx.label + ': ' + ctx.parsed + suffix } },
              datalabels: datalabelsFor(suffix, '#0F1318'),
            }
          }
        };
      } else {
        const horizontal = state.type === 'hbar';
        const stacked = !!opts.stacked;
        const seriesColors = SERIES_COLORS.concat(PIE_PALETTE);
        cfg = {
          type:'bar',
          data:{
            labels: categories,
            datasets: series.map((s,i)=>({
              label: s.name, data: s.data,
              backgroundColor: seriesColors[i % seriesColors.length],
              borderRadius: stacked ? 2 : 3, maxBarThickness: stacked ? 40 : 30,
            }))
          },
          options:{
            animation:false,
            indexAxis: horizontal ? 'y' : 'x',
            maintainAspectRatio:false,
            plugins:{
              legend:{ display:true, position:'top', align:'end', labels:{ color:textColor, boxWidth:11, font:{size:10.5} } },
              tooltip:{ callbacks:{ label: ctx => ctx.dataset.label + ': ' + (horizontal ? ctx.parsed.x : ctx.parsed.y) + suffix } },
              datalabels: datalabelsFor(suffix),
            },
            scales:{
              x:{ stacked, ticks:{ color: tickColor, font:{size:10} }, grid:{ color: horizontal ? gridColor : 'transparent' } },
              y:{ stacked, ticks:{ color: tickColor, font:{size:10} }, grid:{ color: horizontal ? 'transparent' : gridColor } }
            }
          }
        };
      }
      state.instance = new Chart(canvas.getContext('2d'), cfg);
      return;
    }

    const items = state.items || [];
    if(!items.length){
      canvas.style.display = 'none';
      if(emptyEl){ emptyEl.style.display = 'block'; emptyEl.textContent = emptyEl.dataset.defaultText; }
      renderSideTable(id, []);
      return;
    }
    canvas.style.display = 'block';
    if(emptyEl) emptyEl.style.display = 'none';
    renderSideTable(id, items, { labelHeader: opts.labelHeader || 'Etiqueta', valueHeader: opts.valueHeader || 'Cant.', suffix: opts.suffix });

    const labels = items.map(i=>i.label);
    const data = items.map(i=>i.value);
    const baseColor = COLOR_MAP[opts.colorClass] || COLOR_MAP[''];

    let cfg;
    if(state.type === 'pie'){
      cfg = {
        type:'pie',
        data:{ labels, datasets:[{ data, backgroundColor: labels.map((_,i)=>PIE_PALETTE[i % PIE_PALETTE.length]), borderColor:'#171D25', borderWidth:2 }] },
        options:{
          animation:false,
          maintainAspectRatio:false,
          plugins:{
            legend:{ position:'right', labels:{ color:textColor, boxWidth:11, font:{size:10.5}, padding:8 } },
            tooltip:{ callbacks:{ label: ctx => ctx.label + ': ' + ctx.parsed + suffix } },
            datalabels: datalabelsFor(suffix, '#0F1318'),
          }
        }
      };
    } else {
      const horizontal = state.type === 'hbar';
      cfg = {
        type:'bar',
        data:{ labels, datasets:[{ data, backgroundColor: baseColor, borderRadius:4, maxBarThickness:34 }] },
        options:{
          animation:false,
          indexAxis: horizontal ? 'y' : 'x',
          maintainAspectRatio:false,
          plugins:{
            legend:{ display:false },
            tooltip:{ callbacks:{ label: ctx => (horizontal ? ctx.parsed.x : ctx.parsed.y) + suffix } },
            datalabels: datalabelsFor(suffix),
          },
          scales:{
            x:{ ticks:{ color: tickColor, font:{size:10.5} }, grid:{ color: horizontal ? gridColor : 'transparent' } },
            y:{ ticks:{ color: tickColor, font:{size:10.5} }, grid:{ color: horizontal ? 'transparent' : gridColor } }
          }
        }
      };
    }
    state.instance = new Chart(canvas.getContext('2d'), cfg);
  } catch(err){
    // Un gráfico que falla no debe dejar en blanco ni tumbar a los demás:
    // se muestra el mensaje de "sin datos" y el detalle queda en consola
    // para poder diagnosticarlo (F12 → Consola).
    console.error('No se pudo dibujar el gráfico "' + id + '":', err);
    canvas.style.display = 'none';
    if(emptyEl){
      emptyEl.style.display = 'block';
      emptyEl.textContent = 'No se pudo dibujar este gráfico (ver consola del navegador para más detalle).';
    }
    renderSideTable(id, []);
  }
}

function setChartType(id, type){
  chartState[id] = chartState[id] || { type:'hbar' };
  chartState[id].type = type;
  drawChart(id);
}

document.querySelectorAll('.chart-type-toggle').forEach(group=>{
  const target = group.dataset.target;
  group.querySelectorAll('.ctbtn').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      group.querySelectorAll('.ctbtn').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active');
      setChartType(target, btn.dataset.type);
    });
  });
});

function renderAnalysis(list){
  // Pasos por servicio / aplicación — cuántos pasos solicitó cada
  // servicio (SSAS, COMRIC, SISPOS...), el gráfico "estrella" del resumen.
  const porServicio = {};
  list.filter(r=>r.proceso).forEach(r=>{
    porServicio[r.proceso] = (porServicio[r.proceso]||0) + 1;
  });
  const itemsServicio = Object.entries(porServicio)
    .map(([label,value])=>({label,value}))
    .sort((a,b)=>b.value-a.value)
    .slice(0,15);
  renderChartData('chartServicio', itemsServicio, {colorClass:'', labelHeader:'Servicio', valueHeader:'Pasos'});

  // Fallas por servidor
  const porServidor = {};
  list.filter(r=>r.estado==='Fallido' && r.servidor).forEach(r=>{
    porServidor[r.servidor] = (porServidor[r.servidor]||0) + 1;
  });
  const itemsServidor = Object.entries(porServidor)
    .map(([label,value])=>({label,value}))
    .sort((a,b)=>b.value-a.value)
    .slice(0,8);
  renderChartData('chartServidor', itemsServidor, {colorClass:'c-danger', labelHeader:'Servidor', valueHeader:'Fallas'});

  // Cantidad de pasos (registros) por ambiente / servidor — conteo total,
  // no solo fallas, similar al gráfico "Recuento por AMBIENTE" de Power BI.
  const conteoPorAmbiente = {};
  list.filter(r=>r.servidor).forEach(r=>{
    conteoPorAmbiente[r.servidor] = (conteoPorAmbiente[r.servidor]||0) + 1;
  });
  const itemsAmbienteConteo = Object.entries(conteoPorAmbiente)
    .map(([label,value])=>({label,value}))
    .sort((a,b)=>b.value-a.value);
  renderChartData('chartAmbienteConteo', itemsAmbienteConteo, {colorClass:'', labelHeader:'Ambiente', valueHeader:'Pasos'});

  // Tiempo promedio por responsable
  const tiemposPorResp = {};
  list.filter(r=>r.responsable && typeof r.tiempo === 'number').forEach(r=>{
    if(!tiemposPorResp[r.responsable]) tiemposPorResp[r.responsable] = [];
    tiemposPorResp[r.responsable].push(r.tiempo);
  });
  const itemsResp = Object.entries(tiemposPorResp)
    .map(([label,vals])=>({label, value:Math.round(vals.reduce((a,b)=>a+b,0)/vals.length)}))
    .sort((a,b)=>b.value-a.value)
    .slice(0,8);
  renderChartData('chartResponsable', itemsResp, {colorClass:'c-info', suffix:' min', labelHeader:'Responsable', valueHeader:'Prom.'});

  // Tiempo promedio por ambiente / servidor
  const tiemposPorAmbiente = {};
  list.filter(r=>r.servidor && typeof r.tiempo === 'number').forEach(r=>{
    if(!tiemposPorAmbiente[r.servidor]) tiemposPorAmbiente[r.servidor] = [];
    tiemposPorAmbiente[r.servidor].push(r.tiempo);
  });
  const itemsAmbiente = Object.entries(tiemposPorAmbiente)
    .map(([label,vals])=>({label, value:Math.round(vals.reduce((a,b)=>a+b,0)/vals.length)}))
    .sort((a,b)=>b.value-a.value)
    .slice(0,8);
  renderChartData('chartAmbiente', itemsAmbiente, {colorClass:'c-info', suffix:' min', labelHeader:'Ambiente', valueHeader:'Prom.'});

  renderTipoSolicitud(list);
}

// Solicitudes por tipo (ej. "Paso a pruebas", "Paso Producción"...),
// separadas por la columna PROGRAMADO / No Programado del Excel — igual
// al gráfico agrupado de referencia.
function renderTipoSolicitud(list){
  const conData = list.filter(r=>r.tipo);
  const porTipo = {};
  const progValues = new Set();
  conData.forEach(r=>{
    const prog = r.programado || 'No especificado';
    progValues.add(prog);
    if(!porTipo[r.tipo]) porTipo[r.tipo] = {};
    porTipo[r.tipo][prog] = (porTipo[r.tipo][prog]||0) + 1;
  });
  const categories = Object.keys(porTipo).sort((a,b)=>{
    const totalA = Object.values(porTipo[a]).reduce((x,y)=>x+y,0);
    const totalB = Object.values(porTipo[b]).reduce((x,y)=>x+y,0);
    return totalB - totalA;
  });
  const preferredOrder = ['No Programado','Programado'];
  const progList = Array.from(progValues).sort((a,b)=>{
    const ia = preferredOrder.indexOf(a), ib = preferredOrder.indexOf(b);
    if(ia===-1 && ib===-1) return a.localeCompare(b,'es');
    if(ia===-1) return 1;
    if(ib===-1) return -1;
    return ia-ib;
  });
  const series = progList.map(p=>({ name:p, data: categories.map(c=> (porTipo[c][p]||0)) }));
  renderGroupedChart('chartTipoSolicitud', categories, series, { labelHeader:'Tipo solicitud' });
}

// --- Selección local de actividades (dentro de la pestaña "Actividades
// por tipo"): checkboxes para ver una, varias o todas las actividades
// (Script, ETL, MSI...). Por defecto empiezan todas marcadas; al importar
// más archivos, cualquier actividad nueva que aparezca se marca también
// por defecto, sin tocar lo que el usuario ya haya desmarcado.
let activitySelection = new Set();
let knownActivityLabels = new Set();

function syncActivitySelection(labels){
  labels.forEach(l=>{
    if(!knownActivityLabels.has(l)){ activitySelection.add(l); knownActivityLabels.add(l); }
  });
  Array.from(knownActivityLabels).forEach(l=>{
    if(!labels.includes(l)){ knownActivityLabels.delete(l); activitySelection.delete(l); }
  });
}

function renderActivityFilterList(){
  const labels = getActivityLabels();
  syncActivitySelection(labels);
  const wrap = el('act_filter_list');
  if(!labels.length){
    wrap.innerHTML = '<div class="chart-empty">No hay columnas de actividad (conteos tipo ETL, Job, Script...) detectadas todavía.</div>';
    return;
  }
  wrap.innerHTML = labels.map(l=>`
    <label><input type="checkbox" class="act-filter-check" value="${escapeHtml(l)}" ${activitySelection.has(l)?'checked':''}> ${escapeHtml(l)}</label>
  `).join('');
}

el('act_filter_list').addEventListener('change', (e)=>{
  if(!e.target.classList.contains('act-filter-check')) return;
  if(e.target.checked) activitySelection.add(e.target.value);
  else activitySelection.delete(e.target.value);
  renderActividades(getFilteredRegistros());
});

el('btnActSelectAll').onclick = () => {
  getActivityLabels().forEach(l=>activitySelection.add(l));
  renderActivityFilterList();
  renderActividades(getFilteredRegistros());
};
el('btnActSelectNone').onclick = () => {
  activitySelection.clear();
  renderActivityFilterList();
  renderActividades(getFilteredRegistros());
};

// Suma, sobre los registros filtrados, cada columna de conteo (ETL,
// SCRIPT, Modificacion Job, Creacion Job, Restauracion BD...) que esté
// marcada en el selector de la pestaña. Así se responde "¿cuántos ETL,
// cuántos Job...?" viendo una actividad a la vez, varias, o todas.
function renderActividades(list){
  syncActivitySelection(getActivityLabels());
  const totales = {};
  list.forEach(r=>{
    const act = r.actividades || {};
    Object.entries(act).forEach(([label,v])=>{
      if(!activitySelection.has(label)) return;
      totales[label] = (totales[label]||0) + (typeof v === 'number' ? v : 0);
    });
  });
  const items = Object.entries(totales)
    .map(([label,value])=>({label,value}))
    .filter(i=>i.value > 0)
    .sort((a,b)=>b.value-a.value);
  renderChartData('chartActividades', items, {colorClass:'c-warn', labelHeader:'Actividad', valueHeader:'Total'});
  renderServiciosPorActividad(list);
}

// Qué servicio/aplicación pidió cada tipo de actividad marcada en el
// selector de arriba (Script, ETL, MSI...). Barras apiladas por servicio
// (top 10), un color por actividad — si solo hay una actividad marcada,
// queda como una barra simple por servicio.
function renderServiciosPorActividad(list){
  const selected = Array.from(activitySelection);
  if(!selected.length){
    renderGroupedChart('chartServicioActividad', [], [], {});
    return;
  }
  const porServicio = {}; // servicio -> { actividad: total }
  list.forEach(r=>{
    if(!r.proceso) return;
    const act = r.actividades || {};
    Object.entries(act).forEach(([label,v])=>{
      if(!activitySelection.has(label)) return;
      const val = typeof v === 'number' ? v : 0;
      if(val <= 0) return;
      if(!porServicio[r.proceso]) porServicio[r.proceso] = {};
      porServicio[r.proceso][label] = (porServicio[r.proceso][label]||0) + val;
    });
  });
  const categories = Object.keys(porServicio)
    .map(servicio=>({ servicio, total: Object.values(porServicio[servicio]).reduce((a,b)=>a+b,0) }))
    .sort((a,b)=>b.total-a.total)
    .slice(0,10)
    .map(x=>x.servicio);
  const series = selected
    .filter(label=>categories.some(c=>(porServicio[c]||{})[label] > 0))
    .map(label=>({ name:label, data: categories.map(c=>(porServicio[c] && porServicio[c][label]) || 0) }));
  renderGroupedChart('chartServicioActividad', categories, series, { labelHeader:'Servicio', stacked:true });
}

// Qué servicio/aplicación pidió pasos Programados vs No Programados —
// top 10 servicios, barras agrupadas (igual estilo que "Solicitudes por
// tipo" en el Resumen).
function renderServiciosPorProgramado(list){
  const porServicio = {}; // servicio -> { Programado: n, 'No Programado': n }
  const progValues = new Set();
  list.filter(r=>r.proceso).forEach(r=>{
    const prog = r.programado || 'No especificado';
    progValues.add(prog);
    if(!porServicio[r.proceso]) porServicio[r.proceso] = {};
    porServicio[r.proceso][prog] = (porServicio[r.proceso][prog]||0) + 1;
  });
  const categories = Object.keys(porServicio)
    .map(servicio=>({ servicio, total: Object.values(porServicio[servicio]).reduce((a,b)=>a+b,0) }))
    .sort((a,b)=>b.total-a.total)
    .slice(0,10)
    .map(x=>x.servicio);
  const preferredOrder = ['No Programado','Programado'];
  const progList = Array.from(progValues).sort((a,b)=>{
    const ia = preferredOrder.indexOf(a), ib = preferredOrder.indexOf(b);
    if(ia===-1 && ib===-1) return a.localeCompare(b,'es');
    if(ia===-1) return 1;
    if(ib===-1) return -1;
    return ia-ib;
  });
  const series = progList.map(p=>({ name:p, data: categories.map(c=>(porServicio[c] && porServicio[c][p]) || 0) }));
  renderGroupedChart('chartServicioProgramado', categories, series, { labelHeader:'Servicio' });
}

function updateViews(){
  const list = getFilteredRegistros();
  renderStats(list);
  renderAnalysis(list);
  renderActivityFilterList();
  renderActividades(list);
  renderServiciosPorProgramado(list);
  renderTable(list);
}

function renderAll(){ refreshFilterOptions(); updateViews(); }

// --- Pestañas: Resumen / Actividades por tipo / Registros ---
document.querySelectorAll('.tab-btn').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
    document.querySelectorAll('.tabpanel').forEach(p=>p.classList.remove('active'));
    btn.classList.add('active');
    el('tab-' + btn.dataset.tab).classList.add('active');
  });
});

// --- Modal de edición / creación ---
function openModal(reg){
  el('modalTitle').textContent = reg ? 'Editar registro' : 'Nuevo registro';
  el('regId').value = reg ? reg.id : '';
  el('f_proceso').value = reg ? reg.proceso : '';
  el('f_servidor').value = reg ? reg.servidor : '';
  el('f_responsable').value = reg ? reg.responsable : '';
  el('f_estado').value = reg ? reg.estado : 'Exitoso';
  el('f_escalamiento').value = reg ? reg.escalamiento : 'No';
  el('f_fecha').value = reg ? reg.fecha : '';
  el('f_tiempo').value = reg ? reg.tiempo : '';
  el('f_tipo').value = reg ? reg.tipo : '';
  el('f_diagnostico').value = reg ? reg.diagnostico : '';
  el('f_accion').value = reg ? reg.accion : '';
  el('f_extra').value = reg ? (reg.extra || '') : '';
  el('overlay').classList.add('open');
}
function closeModal(){ el('overlay').classList.remove('open'); }

el('btnNew').onclick = () => openModal(null);
el('btnCancel').onclick = closeModal;
el('overlay').addEventListener('click', e => { if(e.target.id==='overlay') closeModal(); });

el('regForm').addEventListener('submit', e => {
  e.preventDefault();
  const id = el('regId').value;
  const data = {
    proceso: el('f_proceso').value.trim(),
    servidor: el('f_servidor').value.trim(),
    responsable: el('f_responsable').value.trim(),
    estado: el('f_estado').value,
    escalamiento: el('f_escalamiento').value,
    fecha: el('f_fecha').value,
    tiempo: el('f_tiempo').value ? parseFloat(el('f_tiempo').value) : '',
    tipo: el('f_tipo').value.trim() || 'Registro',
    diagnostico: el('f_diagnostico').value.trim(),
    accion: el('f_accion').value.trim(),
    extra: el('f_extra').value.trim(),
  };
  if(id){
    const r = registros.find(x=>x.id===id);
    Object.assign(r, data);
  } else {
    registros.push({ id: genId(), origen: 'Manual', ...data });
  }
  closeModal();
  renderAll();
});

function editRegistro(id){ const r = registros.find(x=>x.id===id); if(r) openModal(r); }
function deleteRegistro(id){
  if(!confirm('¿Eliminar este registro?')) return;
  registros = registros.filter(x=>x.id!==id);
  renderAll();
}

// --- Modal "Ver original": muestra la fila 100% cruda tal como la leyó SheetJS ---
function viewRaw(id){
  const r = registros.find(x=>x.id===id);
  if(!r) return;
  const raw = r.raw;
  if(!raw){
    el('rawContent').innerHTML = '<div class="raw-empty">Este registro se creó manualmente en el panel, no proviene de un archivo Excel.</div>';
  } else {
    const entries = Object.entries(raw).filter(([k,v]) => v !== undefined && v !== null && String(v).trim() !== '');
    if(entries.length === 0){
      el('rawContent').innerHTML = '<div class="raw-empty">La fila no tenía ninguna celda con contenido.</div>';
    } else {
      el('rawContent').innerHTML = `
        <table class="raw-table">
          ${entries.map(([k,v]) => `
            <tr>
              <td class="raw-key">${escapeHtml(k)}</td>
              <td class="raw-val">${escapeHtml(v instanceof Date ? v.toISOString() : v)}</td>
            </tr>
          `).join('')}
        </table>
      `;
    }
  }
  el('rawOverlay').classList.add('open');
}
el('btnCloseRaw').onclick = () => el('rawOverlay').classList.remove('open');
el('rawOverlay').addEventListener('click', e => { if(e.target.id==='rawOverlay') el('rawOverlay').classList.remove('open'); });

// La búsqueda de texto libre se "debounce" (espera a que la persona pare de
// escribir) porque cada actualización redibuja 9 gráficos de Chart.js y la
// tabla completa; sin esto, cada tecla presionada dispararía ese trabajo.
// Los selects sí actualizan al instante: no se "escribe" en ellos tecla a
// tecla, así que no hay nada que debounce-ar.
function debounce(fn, wait){
  let timer;
  return function(...args){
    clearTimeout(timer);
    timer = setTimeout(()=>fn.apply(this, args), wait);
  };
}
el('search').addEventListener('input', debounce(updateViews, 250));
['filterEstado','filterTipo','filterOrigen'].forEach(id=>{
  el(id).addEventListener('change', updateViews);
});

el('btnClear').onclick = () => {
  if(registros.length===0) return;
  if(!confirm('Esto borrará todos los registros cargados. ¿Continuar?')) return;
  registros = [];
  renderAll();
};

// --- Import ---
el('btnImport').onclick = () => el('fileInput').click();

function readFileAsync(file){
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = evt => resolve(evt.target.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

el('fileInput').addEventListener('change', async function(e){
  const files = Array.from(e.target.files);
  if(!files.length) return;
  const progress = el('progress');
  let done = 0;
  let importedCount = 0;
  const importedSheets = [];   // { archivo, hoja, filas }
  const skippedSheets = [];    // { archivo, hoja, motivo }

  for(const file of files){
    try{
      const bin = await readFileAsync(file);
      const wb = XLSX.read(bin, {type:'array', cellDates:true});
      for(const sheetName of wb.SheetNames){
        const sheet = wb.Sheets[sheetName];
        const { rows } = sheetToRows(sheet);
        if(!rows.length){
          skippedSheets.push({archivo:file.name, hoja:sheetName, motivo:'no parece una bitácora de registros (no se encontró fila de encabezado reconocible)'});
          continue;
        }
        let countThisSheet = 0;
        rows.forEach(row => {
          registros.push(rowToRegistro(row, sheetName, file.name));
          importedCount++;
          countThisSheet++;
        });
        importedSheets.push({archivo:file.name, hoja:sheetName, filas:countThisSheet});
      }
    }catch(err){
      console.error('Error leyendo ' + file.name, err);
      skippedSheets.push({archivo:file.name, hoja:'(todo el archivo)', motivo:'no se pudo leer el archivo — ¿es un Excel válido?'});
    }
    done++;
    progress.style.width = (done/files.length*100) + '%';
  }

  setTimeout(()=>{ progress.style.width = '0%'; }, 500);
  e.target.value = '';
  renderAll();

  // Resumen de la importación: siempre se informa qué se cargó y qué se
  // omitió, para que sea fácil detectar si algo no entró como se esperaba.
  console.log('Hojas importadas:', importedSheets);
  console.log('Hojas omitidas:', skippedSheets);
  if(importedCount===0){
    alert('No se encontraron hojas con formato de bitácora reconocible en el/los archivo(s) seleccionado(s).\n\n' +
      skippedSheets.map(s=>`• ${s.archivo} — "${s.hoja}": ${s.motivo}`).join('\n'));
  } else if(skippedSheets.length > 0){
    alert(`Se importaron ${importedCount} registros de ${importedSheets.length} hoja(s).\n\n` +
      `Se omitieron ${skippedSheets.length} hoja(s) por no parecer bitácoras de registros:\n` +
      skippedSheets.map(s=>`• ${s.archivo} — "${s.hoja}"`).join('\n') +
      `\n\nSi alguna de estas SÍ debería haberse importado, revísala: puede que sus encabezados usen nombres muy distintos a los esperados (fecha, servidor, responsable, estado...).`);
  }
});

// --- Export Excel (respeta filtros activos) ---
el('btnExport').onclick = () => {
  const list = getFilteredRegistros();
  if(list.length===0){ alert('No hay registros para exportar con los filtros actuales.'); return; }
  const data = list.map(r=>({
    ID:r.id, Fecha:r.fecha, 'Proceso/Job':r.proceso, Servidor:r.servidor, Responsable:r.responsable,
    Tipo:r.tipo, Estado:r.estado, Escalamiento:r.escalamiento, 'Tiempo (min)':r.tiempo,
    Diagnostico:r.diagnostico, Accion:r.accion, 'Informacion adicional':r.extra||'', Origen:r.origen
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Registros');
  XLSX.writeFile(wb, 'panel_operaciones_' + new Date().toISOString().slice(0,10) + '.xlsx');
};

// --- Exportar HTML interactivo (archivo único, autónomo) ---
// Genera un .html independiente con los registros filtrados embebidos, con
// sus propios filtros, gráficos y tabla — para compartir con personal que
// no tiene acceso a este panel. Necesita internet la primera vez que se
// abre (carga Chart.js desde el mismo CDN que usa este panel), pero no
// depende de este panel ni de Excel para nada más.
el('btnExportHtml').onclick = () => {
  const list = getFilteredRegistros();
  if(list.length===0){ alert('No hay registros para exportar con los filtros actuales.'); return; }
  const html = buildInteractiveExportHtml(list);
  const blob = new Blob([html], {type:'text/html;charset=utf-8'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'panel_operaciones_interactivo_' + new Date().toISOString().slice(0,10) + '.html';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 2000);
};

function buildInteractiveExportHtml(list){
  // Se deja afuera el campo "raw" (fila 100% original) para que el archivo
  // no pese de más; todo lo demás que ya se procesó sí viaja, incluidas
  // las actividades (para los gráficos) y la info adicional.
  const dataForExport = list.map(r=>({
    fecha: r.fecha, proceso: r.proceso, servidor: r.servidor, responsable: r.responsable,
    estado: r.estado, escalamiento: r.escalamiento, tiempo: r.tiempo, tipo: r.tipo,
    programado: r.programado, extra: r.extra, actividades: r.actividades, origen: r.origen,
  }));
  // Escapar "<" evita que un dato con "</script>" (por ejemplo escrito a
  // mano en "Info. adicional") rompa el archivo generado al abrirlo.
  const dataJson = JSON.stringify(dataForExport).replace(/</g, '\\u003c');
  const generatedAt = new Date().toLocaleString('es-CO', {dateStyle:'long', timeStyle:'short'});
  const filtros = activeFilterLabels();
  const filtrosTxt = filtros.length ? filtros.join(' · ') : 'Ninguno (todos los registros)';

  const parts = [];
  parts.push('<!DOCTYPE html>');
  parts.push('<html lang="es">');
  parts.push('<head>');
  parts.push('<meta charset="UTF-8">');
  parts.push('<meta name="viewport" content="width=device-width, initial-scale=1.0">');
  parts.push('<title>Panel de Operaciones — Vista interactiva</title>');
  parts.push('<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"><' + '/script>');
  parts.push('<script src="https://cdnjs.cloudflare.com/ajax/libs/chartjs-plugin-datalabels/2.2.0/chartjs-plugin-datalabels.min.js"><' + '/script>');
  parts.push('<style>' + EXPORT_CSS + '</style>');
  parts.push('</head>');
  parts.push('<body>');
  parts.push('<div class="topbar">');
  parts.push('  <div class="brand"><div class="dot"></div><div>');
  parts.push('    <h1>Panel de <span>Operaciones</span> <small class="tag">vista interactiva</small></h1>');
  parts.push('    <small>Generado el ' + escapeHtml(generatedAt) + ' · Filtros aplicados al exportar: ' + escapeHtml(filtrosTxt) + '</small>');
  parts.push('  </div></div>');
  parts.push('</div>');
  parts.push('<main>');
  parts.push('  <div class="filterbar"><div class="filters">');
  parts.push('    <input type="text" id="search" placeholder="Buscar proceso, servidor, responsable...">');
  parts.push('    <select id="filterEstado"><option value="">Todos los estados</option><option>Exitoso</option><option>Fallido</option><option>Pendiente</option><option>Otro</option></select>');
  parts.push('    <select id="filterTipo"><option value="">Todos los tipos</option></select>');
  parts.push('    <select id="filterOrigen"><option value="">Todos los archivos</option></select>');
  parts.push('    <span class="count-info" id="countInfo"></span>');
  parts.push('  </div></div>');
  parts.push('  <div class="stats" id="stats"></div>');

  parts.push('  <h2 class="section-title">Resumen</h2>');
  parts.push(chartPanelHtml('chartServicio', 'Pasos por servicio / aplicación', 'Top 15 servicios/aplicaciones con más pasos solicitados', {wide:true, tall:true}));
  parts.push(chartPanelHtml('chartTipoSolicitud', 'Solicitudes por tipo', 'Separado por Programado / No Programado', {wide:true, tall:true}));
  parts.push('  <div class="analysis">');
  parts.push(chartPanelHtml('chartAmbienteConteo', 'Cantidad de pasos por ambiente', 'Total de registros por servidor / ambiente'));
  parts.push(chartPanelHtml('chartAmbiente', 'Tiempo promedio por ambiente', 'Minutos promedio de atención, por servidor / ambiente'));
  parts.push(chartPanelHtml('chartServidor', 'Fallas por servidor / instancia', 'Top servidores con más registros en estado "Fallido"'));
  parts.push(chartPanelHtml('chartResponsable', 'Tiempo promedio por responsable', 'Minutos promedio de atención, por responsable'));
  parts.push('  </div>');

  parts.push('  <h2 class="section-title">Actividades por tipo</h2>');
  parts.push('  <div class="chart-panel chart-panel-wide">');
  parts.push('    <h2>Actividades por tipo</h2><div class="sub">Suma de cada columna de conteo (ETL, Job, Script, Reporte...)</div>');
  parts.push('    <div class="activity-filter-bar">');
  parts.push('      <div class="activity-filter-head"><label>Mostrar actividades</label><div class="activity-filter-actions">');
  parts.push('        <button type="button" class="ghost tiny" id="actSelectAll">Todas</button>');
  parts.push('        <button type="button" class="ghost tiny" id="actSelectNone">Ninguna</button>');
  parts.push('      </div></div>');
  parts.push('      <div class="checkbox-grid" id="actFilterList"></div>');
  parts.push('    </div>');
  parts.push('    <div class="chart-body">');
  parts.push('      <div class="chart-canvas-wrap chart-canvas-wrap-tall"><canvas id="chartActividades"></canvas><div class="chart-empty" id="chartActividades-empty" style="display:none">Sin datos.</div></div>');
  parts.push('      <div class="chart-side-table-wrap tall"><table class="side-table" id="chartActividades-table"></table></div>');
  parts.push('    </div>');
  parts.push('  </div>');
  parts.push(chartPanelHtml('chartServicioActividad', 'Servicios por actividad', 'Top 10 servicios, según las actividades marcadas arriba', {wide:true, tall:true}));
  parts.push(chartPanelHtml('chartServicioProgramado', 'Servicios por tipo de paso', 'Top 10 servicios, separados por Programado / No Programado', {wide:true, tall:true}));

  parts.push('  <h2 class="section-title">Registros</h2>');
  parts.push('  <div class="panel"><div class="panel-head"><h2>Registros</h2></div><div id="tableWrap"></div></div>');
  parts.push('  <p class="hint">Este archivo es una copia independiente e interactiva de los datos exportados — no está conectado al panel original, así que los cambios que hagas en uno no afectan al otro. Los filtros de arriba y la selección de actividades funcionan localmente sobre los ' + dataForExport.length + ' registro(s) que trae este archivo. Necesita conexión a internet para cargar las librerías de gráficos.</p>');
  parts.push('</main>');
  parts.push('<script>');
  parts.push('const DATA = ' + dataJson + ';');
  parts.push(EXPORT_JS);
  parts.push('<' + '/script>');
  parts.push('</body>');
  parts.push('</html>');
  return parts.join('\n');
}

function chartPanelHtml(id, title, sub, opts){
  opts = opts || {};
  const cls = 'chart-panel' + (opts.wide ? ' chart-panel-wide' : '');
  const canvasCls = 'chart-canvas-wrap' + (opts.tall ? ' chart-canvas-wrap-tall' : '');
  const tableCls = 'chart-side-table-wrap' + (opts.tall ? ' tall' : '');
  return '  <div class="' + cls + '"><h2>' + title + '</h2><div class="sub">' + sub + '</div>' +
    '<div class="chart-body">' +
      '<div class="' + canvasCls + '"><canvas id="' + id + '"></canvas>' +
      '<div class="chart-empty" id="' + id + '-empty" style="display:none">Sin datos para los filtros actuales.</div></div>' +
      '<div class="' + tableCls + '"><table class="side-table" id="' + id + '-table"></table></div>' +
    '</div></div>';
}

const EXPORT_CSS = [
':root{--bg:#0F1318;--panel:#171D25;--panel-2:#1E2630;--border:#2A3340;--text:#E8ECF1;--muted:#8894A3;',
'--accent:#4FD1C5;--accent-dim:#2C6E68;--warn:#F5A524;--danger:#E5484D;--ok:#3DD68C;--info:#5B8DEF;',
'--mono:"JetBrains Mono","Courier New",monospace;--sans:"Inter",-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;}',
'*{box-sizing:border-box;}',
'body{margin:0;background:radial-gradient(circle at 20% 0%, rgba(79,209,197,.06), transparent 40%),var(--bg);color:var(--text);font-family:var(--sans);min-height:100vh;}',
'.topbar{padding:20px 28px;border-bottom:1px solid var(--border);}',
'.brand{display:flex;align-items:center;gap:12px;}',
'.dot{width:10px;height:10px;border-radius:50%;background:var(--accent);flex-shrink:0;}',
'.brand h1{font-size:15px;margin:0;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);font-weight:600;}',
'.brand h1 span{color:var(--text);}',
'.brand h1 .tag{font-size:10px;color:var(--accent);border:1px solid var(--accent-dim);border-radius:20px;padding:2px 8px;text-transform:none;letter-spacing:0;margin-left:8px;vertical-align:middle;}',
'.brand small{display:block;color:var(--muted);font-size:11px;margin-top:2px;letter-spacing:0;text-transform:none;}',
'main{padding:24px 28px 60px;max-width:1360px;margin:0 auto;}',
'.filterbar{background:var(--panel);border:1px solid var(--border);border-radius:14px;padding:14px 16px;margin-bottom:16px;}',
'.filters{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}',
'.filters input,.filters select{background:var(--panel-2);border:1px solid var(--border);color:var(--text);border-radius:8px;padding:8px 10px;font-size:13px;font-family:var(--sans);}',
'.count-info{margin-left:auto;font-size:12px;color:var(--muted);}',
'button{font-family:var(--sans);font-size:13px;font-weight:600;padding:9px 14px;border-radius:8px;border:1px solid var(--border);background:var(--panel-2);color:var(--text);cursor:pointer;}',
'button:hover{border-color:var(--accent);color:var(--accent);} button.ghost{background:transparent;}',
'button.tiny{padding:5px 10px;font-size:11.5px;}',
'.section-title{font-size:13px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;font-weight:700;margin:28px 0 12px;padding-top:14px;border-top:1px solid var(--border);}',
'.section-title:first-of-type{margin-top:0;padding-top:0;border-top:none;}',
'.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px;margin-bottom:22px;}',
'.stat{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:16px 18px;}',
'.stat .n{font-family:var(--mono);font-size:26px;font-weight:700;}',
'.stat .l{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-top:4px;}',
'.stat.total .n{color:var(--accent);} .stat.fallido .n{color:var(--danger);} .stat.exitoso .n{color:var(--ok);}',
'.stat.escalado .n{color:var(--warn);} .stat.tiempo .n{color:var(--info);}',
'.analysis{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:22px;}',
'.chart-panel{background:var(--panel);border:1px solid var(--border);border-radius:14px;padding:16px 18px;margin-bottom:16px;}',
'.chart-panel h2{font-size:13px;margin:0 0 4px;color:var(--muted);font-weight:600;letter-spacing:.03em;}',
'.chart-panel .sub{font-size:11.5px;color:var(--muted);margin-bottom:14px;}',
'.chart-panel-wide{max-width:none;}',
'.chart-body{display:flex;gap:16px;align-items:stretch;}',
'.chart-body .chart-canvas-wrap{flex:1 1 60%;min-width:0;}',
'.chart-canvas-wrap{position:relative;height:250px;}',
'.chart-canvas-wrap-tall{height:340px;}',
'.chart-empty{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;color:var(--muted);font-size:12.5px;text-align:center;padding:0 20px;}',
'.chart-side-table-wrap{flex:0 0 210px;max-height:250px;overflow-y:auto;border:1px solid var(--border);border-radius:10px;background:var(--panel-2);}',
'.chart-side-table-wrap.tall{max-height:340px;}',
'table.side-table{width:100%;border-collapse:collapse;font-size:12px;}',
'table.side-table th{position:sticky;top:0;background:var(--panel-2);text-align:left;padding:8px 10px;font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;border-bottom:1px solid var(--border);}',
'table.side-table th.num{text-align:right;}',
'table.side-table td{padding:7px 10px;border-bottom:1px solid var(--border);color:var(--text);}',
'table.side-table td.num{text-align:right;font-family:var(--mono);color:var(--muted);white-space:nowrap;}',
'table.side-table tr.total-row td{font-weight:700;color:var(--text);border-top:1px solid var(--border);border-bottom:none;}',
'table.side-table tr:last-child td{border-bottom:none;}',
'.activity-filter-bar{margin-bottom:16px;}',
'.activity-filter-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;}',
'.activity-filter-head label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;}',
'.activity-filter-actions{display:flex;gap:6px;}',
'.checkbox-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px 14px;max-height:180px;overflow-y:auto;padding:10px;background:var(--panel-2);border:1px solid var(--border);border-radius:8px;}',
'.checkbox-grid label{display:flex;align-items:center;gap:7px;font-size:12.5px;color:var(--text);cursor:pointer;}',
'.checkbox-grid input{accent-color:var(--accent);width:14px;height:14px;flex-shrink:0;}',
'.panel{background:var(--panel);border:1px solid var(--border);border-radius:14px;overflow:hidden;}',
'.panel-head{padding:16px 18px;border-bottom:1px solid var(--border);}',
'.panel-head h2{font-size:14px;margin:0;color:var(--muted);font-weight:600;letter-spacing:.03em;}',
'table{width:100%;border-collapse:collapse;font-size:13px;}',
'#tableWrap{overflow-x:auto;}',
'thead th{text-align:left;padding:10px 12px;font-size:10.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid var(--border);background:var(--panel-2);white-space:nowrap;}',
'tbody td{padding:11px 12px;border-bottom:1px solid var(--border);vertical-align:top;}',
'tbody tr:hover{background:rgba(79,209,197,.04);} tbody tr:last-child td{border-bottom:none;}',
'.cell-desc{max-width:260px;color:var(--muted);font-size:12.5px;}',
'.badge{display:inline-flex;align-items:center;gap:5px;padding:3px 9px;border-radius:20px;font-size:11.5px;font-weight:600;white-space:nowrap;}',
'.badge::before{content:"";width:6px;height:6px;border-radius:50%;background:currentColor;}',
'.b-exitoso{background:rgba(61,214,140,.12);color:var(--ok);} .b-fallido{background:rgba(229,72,77,.12);color:var(--danger);}',
'.b-pendiente{background:rgba(245,165,36,.12);color:var(--warn);} .b-otro{background:rgba(136,148,163,.12);color:var(--muted);}',
'.esc-si{color:var(--danger);font-weight:700;} .esc-no{color:var(--muted);}',
'.empty{padding:60px 20px;text-align:center;color:var(--muted);}',
'.empty b{color:var(--text);display:block;margin-bottom:6px;font-size:15px;}',
'.hint{font-size:12px;color:var(--muted);margin-top:16px;line-height:1.6;border-top:1px dashed var(--border);padding-top:14px;}',
'@media (max-width:900px){ .chart-body{flex-direction:column;} .chart-side-table-wrap{flex:1 1 auto;max-height:180px;} }',
'@media (max-width:760px){ .analysis{grid-template-columns:1fr;} .checkbox-grid{grid-template-columns:1fr;} thead{display:none;} tbody tr{display:block;border-bottom:6px solid var(--bg);padding:10px 0;}',
'tbody td{display:flex;justify-content:space-between;gap:10px;border:none;padding:6px 14px;text-align:right;}',
'tbody td::before{content:attr(data-label);color:var(--muted);font-size:11px;text-transform:uppercase;text-align:left;} }',
].join('\n');

const EXPORT_JS = [
'(function(){',
'  var el = function(id){ return document.getElementById(id); };',
'  function escapeHtml(str){ var d=document.createElement("div"); d.textContent = str==null?"":String(str); return d.innerHTML; }',
'  function estadoClass(estado){ if(estado==="Exitoso") return "b-exitoso"; if(estado==="Fallido") return "b-fallido"; if(estado==="Pendiente") return "b-pendiente"; return "b-otro"; }',
'  function debounce(fn, wait){ var t; return function(){ var args=arguments, ctx=this; clearTimeout(t); t=setTimeout(function(){ fn.apply(ctx,args); }, wait); }; }',
'  if(typeof ChartDataLabels !== "undefined"){ Chart.register(ChartDataLabels); }',
'',
'  var PIE_PALETTE = ["#4FD1C5","#5B8DEF","#F5A524","#E5484D","#3DD68C","#9F7AEA","#F56565","#38B2AC","#ED8936","#667EEA","#48BB78","#ECC94B","#FC8181","#4299E1","#B794F4"];',
'  var SERIES_COLORS = ["#5B8DEF","#1F3A93","#3DD68C","#F5A524"].concat(PIE_PALETTE);',
'',
'  function uniqueValues(field){',
'    var set = {}; var out = [];',
'    DATA.forEach(function(r){ var v=r[field]; if(v && !set[v]){ set[v]=true; out.push(v); } });',
'    return out.sort(function(a,b){ return String(a).localeCompare(String(b),"es"); });',
'  }',
'',
'  function refreshFilterOptions(){',
'    var tipoSel = el("filterTipo"), origenSel = el("filterOrigen");',
'    var curTipo = tipoSel.value, curOrigen = origenSel.value;',
'    tipoSel.innerHTML = "<option value=\\"\\">Todos los tipos</option>" + uniqueValues("tipo").map(function(t){ return "<option value=\\""+escapeHtml(t)+"\\">"+escapeHtml(t)+"</option>"; }).join("");',
'    origenSel.innerHTML = "<option value=\\"\\">Todos los archivos</option>" + uniqueValues("origen").map(function(o){ return "<option value=\\""+escapeHtml(o)+"\\">"+escapeHtml(o)+"</option>"; }).join("");',
'    tipoSel.value = curTipo; origenSel.value = curOrigen;',
'  }',
'',
'  function getFiltered(){',
'    var search = el("search").value.trim().toLowerCase();',
'    var fEstado = el("filterEstado").value, fTipo = el("filterTipo").value, fOrigen = el("filterOrigen").value;',
'    return DATA.filter(function(r){',
'      var matchSearch = !search || [r.proceso,r.servidor,r.responsable].join(" ").toLowerCase().indexOf(search) !== -1;',
'      var matchEstado = !fEstado || r.estado===fEstado;',
'      var matchTipo = !fTipo || r.tipo===fTipo;',
'      var matchOrigen = !fOrigen || r.origen===fOrigen;',
'      return matchSearch && matchEstado && matchTipo && matchOrigen;',
'    });',
'  }',
'',
'  function statCard(cls,label,value){ return "<div class=\\"stat "+cls+"\\"><div class=\\"n\\">"+value+"</div><div class=\\"l\\">"+label+"</div></div>"; }',
'  function renderStats(list){',
'    var total = list.length;',
'    var fallidos = list.filter(function(r){return r.estado==="Fallido";}).length;',
'    var exitosos = list.filter(function(r){return r.estado==="Exitoso";}).length;',
'    var escalados = list.filter(function(r){return r.escalamiento==="Sí";}).length;',
'    var tiempos = list.map(function(r){return r.tiempo;}).filter(function(t){return typeof t==="number";});',
'    var promedio = tiempos.length ? Math.round(tiempos.reduce(function(a,b){return a+b;},0)/tiempos.length) : 0;',
'    el("stats").innerHTML = statCard("total","Total registros",total) + statCard("fallido","Fallidos",fallidos) + statCard("exitoso","Exitosos",exitosos) + statCard("escalado","Con escalamiento",escalados) + statCard("tiempo","Tiempo promedio (min)",promedio);',
'  }',
'',
'  function getActivityLabels(){',
'    var set = {}; var out = [];',
'    DATA.forEach(function(r){ Object.keys(r.actividades||{}).forEach(function(k){ if(!set[k]){ set[k]=true; out.push(k); } }); });',
'    return out.sort(function(a,b){ return a.localeCompare(b,"es"); });',
'  }',
'',
'  var activitySelection = {};',
'  var knownActivityLabels = {};',
'  function syncActivitySelection(labels){',
'    labels.forEach(function(l){ if(!knownActivityLabels[l]){ activitySelection[l]=true; knownActivityLabels[l]=true; } });',
'  }',
'  function selectedActivityLabels(){',
'    return Object.keys(activitySelection).filter(function(l){ return activitySelection[l]; });',
'  }',
'',
'  function renderActivityFilterList(){',
'    var labels = getActivityLabels();',
'    syncActivitySelection(labels);',
'    var wrap = el("actFilterList");',
'    if(!labels.length){ wrap.innerHTML = "<div class=\\"chart-empty\\" style=\\"position:static\\">No hay columnas de actividad detectadas.</div>"; return; }',
'    wrap.innerHTML = labels.map(function(l){',
'      return "<label><input type=\\"checkbox\\" class=\\"act-check\\" value=\\""+escapeHtml(l)+"\\" "+(activitySelection[l]?"checked":"")+"> "+escapeHtml(l)+"</label>";',
'    }).join("");',
'  }',
'',
'  function renderSideTable(id, items, opts){',
'    opts = opts || {};',
'    var table = el(id + "-table");',
'    if(!table) return;',
'    if(!items || !items.length){ table.innerHTML = ""; return; }',
'    var suffix = opts.suffix || "";',
'    var total = items.reduce(function(a,b){ return a + (b.value||0); }, 0);',
'    var rows = items.map(function(i){',
'      var pct = total ? (i.value/total*100) : 0;',
'      return "<tr><td title=\\""+escapeHtml(i.label)+"\\">"+escapeHtml(i.label)+"</td><td class=\\"num\\">"+i.value+suffix+"</td><td class=\\"num\\">"+pct.toFixed(1)+"%</td></tr>";',
'    }).join("");',
'    table.innerHTML = "<thead><tr><th>"+(opts.labelHeader||"Etiqueta")+"</th><th class=\\"num\\">"+(opts.valueHeader||"Cant.")+"</th><th class=\\"num\\">%</th></tr></thead>" +',
'      "<tbody>"+rows+"<tr class=\\"total-row\\"><td>Total</td><td class=\\"num\\">"+total+suffix+"</td><td class=\\"num\\">100%</td></tr></tbody>";',
'  }',
'',
'  var chartInstances = {};',
'  function destroyChart(id){ if(chartInstances[id]){ chartInstances[id].destroy(); chartInstances[id]=null; } }',
'  function setEmpty(id, isEmpty){',
'    var canvas = el(id), empty = el(id+"-empty");',
'    if(canvas) canvas.style.display = isEmpty ? "none" : "block";',
'    if(empty) empty.style.display = isEmpty ? "flex" : "none";',
'  }',
'',
'  function drawBar(id, items, opts){',
'    opts = opts || {};',
'    var canvas = el(id);',
'    destroyChart(id);',
'    if(!canvas) return;',
'    if(!items.length){ setEmpty(id, true); renderSideTable(id, []); return; }',
'    setEmpty(id, false);',
'    renderSideTable(id, items, opts);',
'    var horizontal = opts.horizontal !== false;',
'    var labels = items.map(function(i){return i.label;}), data = items.map(function(i){return i.value;});',
'    var suffix = opts.suffix || "";',
'    chartInstances[id] = new Chart(canvas.getContext("2d"), { type:"bar",',
'      data:{ labels:labels, datasets:[{ data:data, backgroundColor:opts.color||"#4FD1C5", borderRadius:4, maxBarThickness:30 }] },',
'      options:{ animation:false, indexAxis: horizontal?"y":"x", maintainAspectRatio:false,',
'        plugins:{ legend:{display:false}, datalabels:{ color:"#E8ECF1", font:{size:10,weight:"600"}, formatter:function(v){ return v===0?"":v+suffix; } } },',
'        scales:{ x:{ ticks:{color:"#8894A3",font:{size:10}}, grid:{color: horizontal?"#2A3340":"transparent"} }, y:{ ticks:{color:"#8894A3",font:{size:10}}, grid:{color: horizontal?"transparent":"#2A3340"} } }',
'      }',
'    });',
'  }',
'',
'  function drawGrouped(id, categories, series, opts){',
'    opts = opts || {};',
'    var canvas = el(id);',
'    destroyChart(id);',
'    if(!canvas) return;',
'    var hasData = categories.length && series.some(function(s){ return s.data.some(function(v){ return v>0; }); });',
'    if(!hasData){ setEmpty(id, true); renderSideTable(id, []); return; }',
'    setEmpty(id, false);',
'    renderSideTable(id, categories.map(function(c,i){',
'      return { label:c, value: series.reduce(function(acc,s){ return acc+(s.data[i]||0); },0) };',
'    }), { labelHeader: opts.labelHeader||"Categoría", valueHeader:"Total" });',
'    var stacked = !!opts.stacked;',
'    var horizontal = !!opts.horizontal;',
'    chartInstances[id] = new Chart(canvas.getContext("2d"), { type:"bar",',
'      data:{ labels:categories, datasets: series.map(function(s,i){ return { label:s.name, data:s.data, backgroundColor: SERIES_COLORS[i%SERIES_COLORS.length], borderRadius:3, maxBarThickness: stacked?40:28 }; }) },',
'      options:{ animation:false, indexAxis: horizontal?"y":"x", maintainAspectRatio:false,',
'        plugins:{ legend:{display:true,position:"top",align:"end",labels:{color:"#E8ECF1",boxWidth:11,font:{size:10.5}}}, datalabels:{ color:"#E8ECF1", font:{size:10,weight:"600"}, formatter:function(v){ return v===0?"":v; } } },',
'        scales:{ x:{ stacked:stacked, ticks:{color:"#8894A3",font:{size:10}}, grid:{color: horizontal?"#2A3340":"transparent"} }, y:{ stacked:stacked, ticks:{color:"#8894A3",font:{size:10}}, grid:{color: horizontal?"transparent":"#2A3340"} } }',
'      }',
'    });',
'  }',
'',
'  function topEntries(map, n){',
'    return Object.keys(map).map(function(k){ return {label:k, value:map[k]}; }).sort(function(a,b){ return b.value-a.value; }).slice(0, n||9999);',
'  }',
'  function avgByGroup(list, groupField, valueField, n){',
'    var groups = {};',
'    list.filter(function(r){ return r[groupField] && typeof r[valueField]==="number"; }).forEach(function(r){',
'      if(!groups[r[groupField]]) groups[r[groupField]] = [];',
'      groups[r[groupField]].push(r[valueField]);',
'    });',
'    return Object.keys(groups).map(function(k){',
'      var vals = groups[k];',
'      return { label:k, value: Math.round(vals.reduce(function(a,b){return a+b;},0)/vals.length) };',
'    }).sort(function(a,b){ return b.value-a.value; }).slice(0, n||9999);',
'  }',
'',
'  function renderAnalysis(list){',
'    var porServicio = {};',
'    list.forEach(function(r){ if(r.proceso) porServicio[r.proceso]=(porServicio[r.proceso]||0)+1; });',
'    drawBar("chartServicio", topEntries(porServicio,15), { color:"#4FD1C5", horizontal:true, labelHeader:"Servicio", valueHeader:"Pasos" });',
'',
'    var porServidorFallas = {};',
'    list.filter(function(r){return r.estado==="Fallido"&&r.servidor;}).forEach(function(r){ porServidorFallas[r.servidor]=(porServidorFallas[r.servidor]||0)+1; });',
'    drawBar("chartServidor", topEntries(porServidorFallas,8), { color:"#E5484D", horizontal:true, labelHeader:"Servidor", valueHeader:"Fallas" });',
'',
'    var porAmbienteConteo = {};',
'    list.forEach(function(r){ if(r.servidor) porAmbienteConteo[r.servidor]=(porAmbienteConteo[r.servidor]||0)+1; });',
'    drawBar("chartAmbienteConteo", topEntries(porAmbienteConteo), { color:"#4FD1C5", horizontal:false, labelHeader:"Ambiente", valueHeader:"Pasos" });',
'',
'    drawBar("chartResponsable", avgByGroup(list,"responsable","tiempo",8), { color:"#5B8DEF", horizontal:true, suffix:" min", labelHeader:"Responsable", valueHeader:"Prom." });',
'    drawBar("chartAmbiente", avgByGroup(list,"servidor","tiempo",8), { color:"#5B8DEF", horizontal:true, suffix:" min", labelHeader:"Ambiente", valueHeader:"Prom." });',
'',
'    renderTipoSolicitud(list);',
'  }',
'',
'  function programadoSeries(list, groupField, topN){',
'    var porGrupo = {}, progSet = {}, progList0 = [];',
'    list.filter(function(r){ return r[groupField]; }).forEach(function(r){',
'      var prog = r.programado || "No especificado";',
'      if(!progSet[prog]){ progSet[prog]=true; progList0.push(prog); }',
'      if(!porGrupo[r[groupField]]) porGrupo[r[groupField]] = {};',
'      porGrupo[r[groupField]][prog] = (porGrupo[r[groupField]][prog]||0) + 1;',
'    });',
'    var categories = Object.keys(porGrupo).map(function(g){',
'      var total=0; Object.keys(porGrupo[g]).forEach(function(k){ total+=porGrupo[g][k]; });',
'      return {g:g, total:total};',
'    }).sort(function(a,b){ return b.total-a.total; });',
'    if(topN) categories = categories.slice(0, topN);',
'    categories = categories.map(function(x){ return x.g; });',
'    var order = ["No Programado","Programado"];',
'    progList0.sort(function(a,b){',
'      var ia=order.indexOf(a), ib=order.indexOf(b);',
'      if(ia===-1 && ib===-1) return a.localeCompare(b,"es");',
'      if(ia===-1) return 1; if(ib===-1) return -1;',
'      return ia-ib;',
'    });',
'    var series = progList0.map(function(p){ return { name:p, data: categories.map(function(c){ return (porGrupo[c][p]||0); }) }; });',
'    return { categories: categories, series: series };',
'  }',
'',
'  function renderTipoSolicitud(list){',
'    var r = programadoSeries(list, "tipo");',
'    drawGrouped("chartTipoSolicitud", r.categories, r.series, { labelHeader:"Tipo solicitud", horizontal:false, stacked:false });',
'  }',
'',
'  function renderServiciosPorProgramado(list){',
'    var r = programadoSeries(list, "proceso", 10);',
'    drawGrouped("chartServicioProgramado", r.categories, r.series, { labelHeader:"Servicio", horizontal:true, stacked:false });',
'  }',
'',
'  function renderActividades(list){',
'    syncActivitySelection(getActivityLabels());',
'    var totales = {};',
'    list.forEach(function(r){',
'      var act = r.actividades || {};',
'      Object.keys(act).forEach(function(label){',
'        if(!activitySelection[label]) return;',
'        var v = act[label];',
'        totales[label] = (totales[label]||0) + (typeof v==="number"?v:0);',
'      });',
'    });',
'    var items = topEntries(totales).filter(function(i){ return i.value>0; });',
'    drawBar("chartActividades", items, { color:"#F5A524", horizontal:true, labelHeader:"Actividad", valueHeader:"Total" });',
'    renderServiciosPorActividad(list);',
'  }',
'',
'  function renderServiciosPorActividad(list){',
'    var selected = selectedActivityLabels();',
'    if(!selected.length){ drawGrouped("chartServicioActividad", [], [], {}); return; }',
'    var porServicio = {};',
'    list.forEach(function(r){',
'      if(!r.proceso) return;',
'      var act = r.actividades || {};',
'      Object.keys(act).forEach(function(label){',
'        if(!activitySelection[label]) return;',
'        var val = typeof act[label]==="number" ? act[label] : 0;',
'        if(val<=0) return;',
'        if(!porServicio[r.proceso]) porServicio[r.proceso]={};',
'        porServicio[r.proceso][label] = (porServicio[r.proceso][label]||0)+val;',
'      });',
'    });',
'    var categories = Object.keys(porServicio).map(function(s){',
'      var total=0; Object.keys(porServicio[s]).forEach(function(k){total+=porServicio[s][k];});',
'      return {servicio:s,total:total};',
'    }).sort(function(a,b){return b.total-a.total;}).slice(0,10).map(function(x){return x.servicio;});',
'    var series = selected.filter(function(label){',
'      return categories.some(function(c){ return (porServicio[c]||{})[label] > 0; });',
'    }).map(function(label){',
'      return { name:label, data: categories.map(function(c){ return (porServicio[c] && porServicio[c][label]) || 0; }) };',
'    });',
'    drawGrouped("chartServicioActividad", categories, series, { labelHeader:"Servicio", stacked:true, horizontal:true });',
'  }',
'',
'  function renderTable(list){',
'    if(!list.length){ el("tableWrap").innerHTML = "<div class=\\"empty\\"><b>Sin resultados</b>Ajusta los filtros de arriba.</div>"; return; }',
'    var rows = list.map(function(r){',
'      return "<tr>" +',
'        "<td data-label=\\"Fecha\\">"+escapeHtml(r.fecha||"—")+"</td>" +',
'        "<td data-label=\\"Proceso\\">"+escapeHtml(r.proceso)+"</td>" +',
'        "<td data-label=\\"Servidor\\">"+escapeHtml(r.servidor||"—")+"</td>" +',
'        "<td data-label=\\"Responsable\\">"+escapeHtml(r.responsable||"—")+"</td>" +',
'        "<td data-label=\\"Tipo\\">"+escapeHtml(r.tipo||"—")+"</td>" +',
'        "<td data-label=\\"Estado\\"><span class=\\"badge "+estadoClass(r.estado)+"\\">"+escapeHtml(r.estado)+"</span></td>" +',
'        "<td data-label=\\"Escal.\\" class=\\""+(r.escalamiento==="Sí"?"esc-si":"esc-no")+"\\">"+escapeHtml(r.escalamiento||"No")+"</td>" +',
'        "<td data-label=\\"Tiempo\\">"+(r.tiempo!==""&&r.tiempo!=null ? r.tiempo+" min" : "—")+"</td>" +',
'        "<td data-label=\\"Info. adicional\\" class=\\"cell-desc\\">"+escapeHtml(r.extra||"—")+"</td>" +',
'      "</tr>";',
'    }).join("");',
'    el("tableWrap").innerHTML = "<table><thead><tr><th>Fecha</th><th>Proceso/Job</th><th>Servidor</th><th>Responsable</th><th>Tipo</th><th>Estado</th><th>Escal.</th><th>Tiempo</th><th>Info. adicional</th></tr></thead><tbody>"+rows+"</tbody></table>";',
'  }',
'',
'  function update(){',
'    var list = getFiltered();',
'    renderStats(list);',
'    renderAnalysis(list);',
'    renderActivityFilterList();',
'    renderActividades(list);',
'    renderServiciosPorProgramado(list);',
'    renderTable(list);',
'    el("countInfo").textContent = list.length + " de " + DATA.length + " registros";',
'  }',
'',
'  el("actFilterList").addEventListener("change", function(e){',
'    if(!e.target.classList.contains("act-check")) return;',
'    activitySelection[e.target.value] = e.target.checked;',
'    renderActividades(getFiltered());',
'  });',
'  el("actSelectAll").onclick = function(){',
'    getActivityLabels().forEach(function(l){ activitySelection[l]=true; });',
'    renderActivityFilterList();',
'    renderActividades(getFiltered());',
'  };',
'  el("actSelectNone").onclick = function(){',
'    Object.keys(activitySelection).forEach(function(l){ activitySelection[l]=false; });',
'    renderActivityFilterList();',
'    renderActividades(getFiltered());',
'  };',
'',
'  refreshFilterOptions();',
'  el("search").addEventListener("input", debounce(update, 200));',
'  ["filterEstado","filterTipo","filterOrigen"].forEach(function(id){ el(id).addEventListener("change", update); });',
'  update();',
'})();',
].join('\n');

// --- Helpers de actividades para el informe (agrupan por etiqueta limpia,
// igual que el gráfico, para que "SCRIPT" de distintos meses/hojas se
// trate como una sola actividad aunque el encabezado original varíe un
// poco en espacios/saltos de línea) ---
const MESES_ES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

function getActivityLabels(){
  const set = new Set();
  registros.forEach(r=>{
    Object.keys(r.actividades||{}).forEach(k=> set.add(k));
  });
  return Array.from(set).sort((a,b)=>a.localeCompare(b,'es'));
}

function activityValueForLabel(r, label){
  const v = (r.actividades || {})[label];
  return typeof v === 'number' ? v : 0;
}

function periodoLabelFromValue(p){
  const [y,m] = p.split('-');
  return (MESES_ES[parseInt(m,10)-1] || m) + ' ' + y;
}

function slugify(s){
  return String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[^a-z0-9]+/g,'_').replace(/^_+|_+$/g,'');
}

// --- Modal: configurar informe ---
function openReportModal(){
  const periodSel = el('rep_periodo');
  const periods = new Set();
  registros.forEach(r=>{ if(r.fecha && /^\d{4}-\d{2}/.test(r.fecha)) periods.add(r.fecha.slice(0,7)); });
  const sortedPeriods = Array.from(periods).sort();
  periodSel.innerHTML = '<option value="">Todo el periodo (según filtros actuales)</option>' +
    sortedPeriods.map(p=>`<option value="${p}">${escapeHtml(periodoLabelFromValue(p))}</option>`).join('');

  const labels = getActivityLabels();
  const wrap = el('rep_actividades_list');
  if(!labels.length){
    wrap.innerHTML = '<div class="chart-empty">No se detectaron columnas de actividad (ETL, Job, Script...) en los registros importados.</div>';
  } else {
    wrap.innerHTML = labels.map(l=>`
      <label><input type="checkbox" class="rep-act-check" value="${escapeHtml(l)}"> ${escapeHtml(l)}</label>
    `).join('');
  }

  document.querySelector('input[name="rep_tipo"][value="general"]').checked = true;
  el('rep_actividades_wrap').style.display = 'none';
  el('reportOverlay').classList.add('open');
}

el('btnReport').onclick = openReportModal;
el('btnCancelReport').onclick = () => el('reportOverlay').classList.remove('open');

document.querySelectorAll('input[name="rep_tipo"]').forEach(radio=>{
  radio.addEventListener('change', ()=>{
    const val = document.querySelector('input[name="rep_tipo"]:checked').value;
    el('rep_actividades_wrap').style.display = (val==='actividad') ? 'block' : 'none';
  });
});

el('btnGenerateReport').onclick = () => {
  const periodo = el('rep_periodo').value;
  const tipo = document.querySelector('input[name="rep_tipo"]:checked').value;
  const selectedActs = Array.from(document.querySelectorAll('.rep-act-check:checked')).map(c=>c.value);

  if(tipo==='actividad' && selectedActs.length===0){
    alert('Selecciona al menos una actividad, o cambia a "Resumen general".');
    return;
  }

  let list = getFilteredRegistros();
  if(periodo) list = list.filter(r=> (r.fecha||'').slice(0,7) === periodo);
  if(tipo==='actividad'){
    list = list.filter(r=> selectedActs.some(label => activityValueForLabel(r,label) > 0));
  }

  if(list.length===0){
    alert('No hay registros que cumplan los criterios seleccionados (periodo / actividad) para generar el informe.');
    return;
  }

  el('reportOverlay').classList.remove('open');
  generarInformePDF(list, {
    periodo,
    periodoLabel: periodo ? periodoLabelFromValue(periodo) : 'Todo el periodo',
    tipo,
    selectedActs,
  });
};

// --- Captura de gráficos como imagen para el PDF ---
// Los canvas que están en una pestaña oculta (display:none) miden 0x0, así
// que Chart.js no puede dibujarlos bien. Por eso, mientras se capturan las
// imágenes se muestran las 3 pestañas a la vez (de forma síncrona, sin
// pintar en pantalla) y luego se restaura la vista normal.
function withAllTabsVisible(fn){
  const panels = document.querySelectorAll('.tabpanel');
  const prevDisplay = Array.from(panels).map(p=>p.style.display);
  panels.forEach(p=>{ p.style.display = 'block'; });
  let result;
  try{
    result = fn();
  } finally {
    panels.forEach((p,i)=>{ p.style.display = prevDisplay[i]; });
  }
  return result;
}

function captureChartImage(id){
  const canvas = document.getElementById(id);
  if(!canvas || canvas.style.display === 'none') return null;
  try{
    return canvas.toDataURL('image/png', 1.0);
  }catch(e){
    console.error('No se pudo capturar el gráfico ' + id, e);
    return null;
  }
}

// Renderiza los 4 gráficos con los datos exactos del informe (mismo
// periodo / actividad elegidos, no necesariamente lo que está visible en
// pantalla ahora mismo), captura cada uno como imagen, y restaura la vista
// normal del panel (según los filtros activos) al terminar.
function captureReportCharts(list){
  const currentList = getFilteredRegistros();
  let images = {};
  try{
    withAllTabsVisible(()=>{
      renderAnalysis(list);
      renderActividades(list);
      renderServiciosPorProgramado(list);
      images = {
        servicio: captureChartImage('chartServicio'),
        servidor: captureChartImage('chartServidor'),
        ambienteConteo: captureChartImage('chartAmbienteConteo'),
        responsable: captureChartImage('chartResponsable'),
        ambiente: captureChartImage('chartAmbiente'),
        tipoSolicitud: captureChartImage('chartTipoSolicitud'),
        actividades: captureChartImage('chartActividades'),
        servicioActividad: captureChartImage('chartServicioActividad'),
        servicioProgramado: captureChartImage('chartServicioProgramado'),
      };
    });
  } catch(err){
    console.error('No se pudieron capturar los gráficos para el PDF:', err);
  } finally {
    // Pase lo que pase arriba, la vista normal del panel (según los
    // filtros activos) siempre se restaura al terminar.
    renderAnalysis(currentList);
    renderActividades(currentList);
    renderServiciosPorProgramado(currentList);
  }
  return images;
}

// Ubica una imagen dentro de una caja máxima (maxW x maxH), conservando su
// proporción original, y devuelve el tamaño final usado.
function addImageFit(doc, dataUrl, x, y, maxW, maxH){
  const props = doc.getImageProperties(dataUrl);
  const ratio = props.width / props.height;
  let w = maxW, h = maxW / ratio;
  if(h > maxH){ h = maxH; w = maxH * ratio; }
  doc.addImage(dataUrl, 'PNG', x, y, w, h);
  return { w, h };
}


function generarInformePDF(list, meta){
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation:'landscape', unit:'mm', format:'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const fechaGen = new Date().toLocaleDateString('es-CO', { year:'numeric', month:'long', day:'numeric' });
  const esActividad = meta.tipo === 'actividad';

  doc.setFont('helvetica','bold'); doc.setFontSize(16);
  doc.text('Informe de Operaciones', 14, 16);
  doc.setFont('helvetica','normal'); doc.setFontSize(10); doc.setTextColor(100);
  doc.text('Generado el ' + fechaGen + '   ·   Periodo: ' + meta.periodoLabel, 14, 22);

  const tipoLabel = esActividad ? ('Por actividad específica: ' + meta.selectedActs.join(', ')) : 'Resumen general (todos los procedimientos)';
  doc.text('Tipo de informe: ' + tipoLabel, 14, 27);

  const filtros = activeFilterLabels();
  doc.text(filtros.length ? 'Filtros aplicados: ' + filtros.join(' · ') : 'Filtros aplicados: ninguno', 14, 32);

  // Resumen
  const total = list.length;
  const fallidos = list.filter(r=>r.estado==='Fallido').length;
  const exitosos = list.filter(r=>r.estado==='Exitoso').length;
  const escalados = list.filter(r=>r.escalamiento==='Sí').length;
  const tiempos = list.map(r=>r.tiempo).filter(t=>typeof t === 'number');
  const promedio = tiempos.length ? Math.round(tiempos.reduce((a,b)=>a+b,0)/tiempos.length) : 0;

  doc.setTextColor(20); doc.setFont('helvetica','bold'); doc.setFontSize(11);
  doc.text('Resumen', 14, 41);
  doc.setFont('helvetica','normal'); doc.setFontSize(10);
  doc.text(`Total: ${total}      Fallidos: ${fallidos}      Exitosos: ${exitosos}      Con escalamiento: ${escalados}      Tiempo promedio: ${promedio} min`, 14, 47);

  let y = 47;

  // Si el informe es "por actividad", mostrar el total de cada actividad
  // seleccionada dentro del periodo/filtros elegidos.
  if(esActividad){
    const totalesAct = meta.selectedActs.map(label=>{
      let sum = 0, registrosConEsa = 0;
      list.forEach(r=>{
        const v = activityValueForLabel(r, label);
        if(v > 0){ sum += v; registrosConEsa++; }
      });
      return { label, sum, registrosConEsa };
    });
    y += 9;
    doc.setFont('helvetica','bold'); doc.setFontSize(11);
    doc.text('Totales de actividad en el periodo', 14, y);
    doc.setFont('helvetica','normal'); doc.setFontSize(10);
    totalesAct.forEach(t=>{
      y += 6;
      doc.text(`${t.label}: ${t.sum} en total, sobre ${t.registrosConEsa} registro(s)`, 14, y);
    });
  }

  // Top servidores / responsables / ambiente (tiempo)
  const porServidor = {};
  list.filter(r=>r.estado==='Fallido' && r.servidor).forEach(r=>{ porServidor[r.servidor] = (porServidor[r.servidor]||0)+1; });
  const topServidores = Object.entries(porServidor).sort((a,b)=>b[1]-a[1]).slice(0,5);

  const tiemposPorResp = {};
  list.filter(r=>r.responsable && typeof r.tiempo === 'number').forEach(r=>{
    if(!tiemposPorResp[r.responsable]) tiemposPorResp[r.responsable] = [];
    tiemposPorResp[r.responsable].push(r.tiempo);
  });
  const topResponsables = Object.entries(tiemposPorResp)
    .map(([k,vals])=>[k, Math.round(vals.reduce((a,b)=>a+b,0)/vals.length)])
    .sort((a,b)=>b[1]-a[1]).slice(0,5);

  const tiemposPorAmbiente = {};
  list.filter(r=>r.servidor && typeof r.tiempo === 'number').forEach(r=>{
    if(!tiemposPorAmbiente[r.servidor]) tiemposPorAmbiente[r.servidor] = [];
    tiemposPorAmbiente[r.servidor].push(r.tiempo);
  });
  const topAmbientes = Object.entries(tiemposPorAmbiente)
    .map(([k,vals])=>[k, Math.round(vals.reduce((a,b)=>a+b,0)/vals.length)])
    .sort((a,b)=>b[1]-a[1]).slice(0,5);

  const col1 = 14, col2 = pageWidth/3 + 4, col3 = (pageWidth/3)*2 - 6;
  y += 10;
  doc.setFont('helvetica','bold');
  doc.text('Top servidores con más fallas', col1, y);
  doc.text('Tiempo promedio por responsable', col2, y);
  doc.text('Tiempo promedio por ambiente', col3, y);
  doc.setFont('helvetica','normal');
  const maxRows = Math.max(topServidores.length, topResponsables.length, topAmbientes.length, 1);
  for(let i=0;i<maxRows;i++){
    y += 6;
    if(topServidores[i]) doc.text(`${topServidores[i][0]} — ${topServidores[i][1]} fallas`, col1, y);
    if(topResponsables[i]) doc.text(`${topResponsables[i][0]} — ${topResponsables[i][1]} min`, col2, y);
    if(topAmbientes[i]) doc.text(`${topAmbientes[i][0]} — ${topAmbientes[i][1]} min`, col3, y);
  }
  if(!topServidores.length && !topResponsables.length && !topAmbientes.length){ y += 6; doc.text('Sin datos suficientes.', col1, y); }

  // Páginas de gráficos: se capturan los 7 gráficos del panel (pasos por
  // servicio, fallas por servidor, cantidad de pasos por ambiente, tiempo
  // por responsable, tiempo por ambiente, solicitudes por tipo,
  // actividades por tipo), calculados exactamente con los datos de este
  // informe (periodo / actividad elegidos), no lo que se ve en pantalla.
  const chartImages = captureReportCharts(list);
  const gMargin = 14;

  // Página 1: el gráfico "estrella" (pasos por servicio/aplicación), a
  // todo el ancho, con más alto porque puede traer hasta 15 barras.
  doc.addPage();
  doc.setTextColor(20); doc.setFont('helvetica','bold'); doc.setFontSize(13);
  doc.text('Gráficos', 14, 16);
  doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(100);
  doc.text('Calculados sobre los ' + total + ' registro(s) de este informe.', 14, 21);
  doc.setTextColor(20); doc.setFont('helvetica','bold'); doc.setFontSize(11);
  doc.text('Pasos por servicio / aplicación', gMargin, 30);
  if(chartImages.servicio){
    addImageFit(doc, chartImages.servicio, gMargin, 34, pageWidth - gMargin*2, 165);
  } else {
    doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(140);
    doc.text('Sin datos suficientes para este gráfico.', gMargin, 45);
  }

  // Página 2: los 4 gráficos de barra simple, en cuadrícula 2x2.
  doc.addPage();
  const gTitles = ['Fallas por servidor / instancia', 'Cantidad de pasos por ambiente', 'Tiempo promedio por responsable', 'Tiempo promedio por ambiente'];
  const gImages = [chartImages.servidor, chartImages.ambienteConteo, chartImages.responsable, chartImages.ambiente];
  const gGap = 8;
  const cellW = (pageWidth - gMargin*2 - gGap) / 2;
  const cellH = 78;
  const rowsY = [16, 16 + cellH + 14];
  for(let i=0;i<4;i++){
    const col = i % 2, row = Math.floor(i/2);
    const x = gMargin + col*(cellW+gGap);
    const yTop = rowsY[row];
    doc.setTextColor(20); doc.setFont('helvetica','bold'); doc.setFontSize(10);
    doc.text(gTitles[i], x, yTop);
    if(gImages[i]){
      addImageFit(doc, gImages[i], x, yTop + 3, cellW, cellH);
    } else {
      doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(140);
      doc.text('Sin datos suficientes para este gráfico.', x, yTop + 12);
    }
  }

  // Páginas siguientes: los gráficos anchos (solicitudes por tipo,
  // actividades por tipo, servicios por actividad, servicios por tipo de
  // paso), apilados de a 2 por página a todo el ancho.
  const wideTitles = ['Solicitudes por tipo', 'Actividades por tipo', 'Servicios por actividad', 'Servicios por tipo de paso'];
  const wideImages = [chartImages.tipoSolicitud, chartImages.actividades, chartImages.servicioActividad, chartImages.servicioProgramado];
  const wideW = pageWidth - gMargin*2, wideH = 80;
  for(let i=0;i<wideTitles.length;i+=2){
    doc.addPage();
    doc.setTextColor(20); doc.setFont('helvetica','bold'); doc.setFontSize(13);
    doc.text('Gráficos (continuación)', 14, 16);
    const wideRowsY = [26, 26 + wideH + 14];
    for(let j=0;j<2 && i+j<wideTitles.length;j++){
      const idx = i+j;
      const x = gMargin, yTop = wideRowsY[j];
      doc.setTextColor(20); doc.setFont('helvetica','bold'); doc.setFontSize(10);
      doc.text(wideTitles[idx], x, yTop);
      if(wideImages[idx]){
        addImageFit(doc, wideImages[idx], x, yTop + 3, wideW, wideH);
      } else {
        doc.setFont('helvetica','normal'); doc.setFontSize(9); doc.setTextColor(140);
        doc.text('Sin datos suficientes para este gráfico.', x, yTop + 12);
      }
    }
  }
  doc.addPage();

  // Tabla de registros: en modo "actividad" la última columna muestra solo
  // las actividades seleccionadas por fila; en modo general muestra la
  // info adicional completa, como antes.
  doc.setTextColor(20); doc.setFont('helvetica','bold'); doc.setFontSize(13);
  doc.text('Detalle de registros', 14, 16);
  const lastColHeader = esActividad ? 'Actividades seleccionadas' : 'Info. adicional';
  const tableRows = list.map(r=>{
    const lastCol = esActividad
      ? (meta.selectedActs.map(label=>{
          const v = activityValueForLabel(r, label);
          return v>0 ? (label + ': ' + v) : null;
        }).filter(Boolean).join(' · ') || '—')
      : (r.extra || '—');
    return [r.id, r.fecha||'—', r.proceso, r.servidor||'—', r.responsable||'—', r.tipo, r.estado, r.escalamiento, r.tiempo!==''?r.tiempo:'—', lastCol];
  });
  doc.autoTable({
    startY: 22,
    head: [['ID','Fecha','Proceso/Job','Servidor','Responsable','Tipo','Estado','Escal.','Tiempo', lastColHeader]],
    body: tableRows,
    styles: { fontSize:7, cellPadding:2, overflow:'linebreak' },
    headStyles: { fillColor:[23,29,37], textColor:255 },
    alternateRowStyles: { fillColor:[245,247,249] },
    columnStyles: { 9: { cellWidth: 55 } },
    margin: { left:14, right:14 },
  });

  const filenameParts = ['informe_operaciones'];
  if(meta.periodo) filenameParts.push(meta.periodo);
  if(esActividad) filenameParts.push(meta.selectedActs.slice(0,2).map(slugify).join('_'));
  filenameParts.push(new Date().toISOString().slice(0,10));
  doc.save(filenameParts.join('_') + '.pdf');
}

renderAll();
