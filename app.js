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

// Detecta si una hoja del Excel "parece" una bitácora de registros
// (fecha + servidor/job/responsable/estado...) en vez de un calendario,
// resumen dinámico u hoja de referencia.
function isLogSheet(rows){
  if(!rows.length) return false;
  return countKeywordMatches(rows.slice(0,5).flatMap(r=>Object.keys(r))) >= 2;
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

  const diagnostico = String(pick(['diagnostico','observaciones','accion a tomar']));
  const accion = String(pick(['respuesta escalamiento','accion','resultado']));

  // Actividades: columnas de conteo tipo checkbox (ETL, SCRIPT, Modificacion
  // Job, Creacion Job, Restauracion BD...) que traen un número (0,1,2...)
  // indicando cuántas veces se hizo esa actividad en el registro. Se
  // guardan aparte (no como texto libre) para poder sumarlas y filtrarlas
  // en la pestaña "Actividades por tipo".
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
        actividades[String(k).trim()] = num;
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
    id: genId(), fecha, proceso, servidor, responsable, estado, escalamiento, tiempo, tipo,
    diagnostico, accion, extra: extraParts.join(' | '), actividades,
    origen: fileName + ' · ' + sheetName,
    raw: row, // fila 100% original, sin procesar, tal como la entregó SheetJS
  };
}

function estadoClass(estado){
  const map = {'Exitoso':'b-exitoso','Fallido':'b-fallido','Pendiente':'b-pendiente'};
  return map[estado] || 'b-otro';
}

function escapeHtml(str){
  const d = document.createElement('div');
  d.textContent = str==null ? '' : String(str);
  return d.innerHTML;
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
const chartState = {};   // id -> { type, items, opts, instance }
const PIE_PALETTE = ['#4FD1C5','#5B8DEF','#F5A524','#E5484D','#3DD68C','#9F7AEA','#F56565','#38B2AC','#ED8936','#667EEA','#48BB78','#ECC94B','#FC8181','#4299E1','#B794F4'];
const COLOR_MAP = { 'c-danger':'#E5484D', 'c-info':'#5B8DEF', 'c-warn':'#F5A524', '':'#4FD1C5' };

function cleanLabel(s){
  return String(s).replace(/[\r\n]+/g,' ').replace(/\s+/g,' ').trim();
}

function renderChartData(id, items, opts={}){
  chartState[id] = chartState[id] || { type:'hbar' };
  chartState[id].items = items;
  chartState[id].opts = opts;
  drawChart(id);
}

function drawChart(id){
  const state = chartState[id];
  if(!state) return;
  const canvas = document.getElementById(id);
  const emptyEl = document.getElementById(id + '-empty');
  if(!canvas) return;
  if(state.instance){ state.instance.destroy(); state.instance = null; }

  const items = state.items || [];
  if(!items.length){
    canvas.style.display = 'none';
    if(emptyEl) emptyEl.style.display = 'block';
    return;
  }
  canvas.style.display = 'block';
  if(emptyEl) emptyEl.style.display = 'none';

  const opts = state.opts || {};
  const labels = items.map(i=>i.label);
  const data = items.map(i=>i.value);
  const baseColor = COLOR_MAP[opts.colorClass] || COLOR_MAP[''];
  const suffix = opts.suffix || '';
  const gridColor = '#2A3340', tickColor = '#8894A3', textColor = '#E8ECF1';

  let cfg;
  if(state.type === 'pie'){
    cfg = {
      type:'pie',
      data:{ labels, datasets:[{ data, backgroundColor: labels.map((_,i)=>PIE_PALETTE[i % PIE_PALETTE.length]), borderColor:'#171D25', borderWidth:2 }] },
      options:{
        maintainAspectRatio:false,
        plugins:{
          legend:{ position:'right', labels:{ color:textColor, boxWidth:11, font:{size:10.5}, padding:8 } },
          tooltip:{ callbacks:{ label: ctx => ctx.label + ': ' + ctx.parsed + suffix } }
        }
      }
    };
  } else {
    const horizontal = state.type === 'hbar';
    cfg = {
      type:'bar',
      data:{ labels, datasets:[{ data, backgroundColor: baseColor, borderRadius:4, maxBarThickness:34 }] },
      options:{
        indexAxis: horizontal ? 'y' : 'x',
        maintainAspectRatio:false,
        plugins:{
          legend:{ display:false },
          tooltip:{ callbacks:{ label: ctx => (horizontal ? ctx.parsed.x : ctx.parsed.y) + suffix } }
        },
        scales:{
          x:{ ticks:{ color: tickColor, font:{size:10.5} }, grid:{ color: horizontal ? gridColor : 'transparent' } },
          y:{ ticks:{ color: tickColor, font:{size:10.5} }, grid:{ color: horizontal ? 'transparent' : gridColor } }
        }
      }
    };
  }
  state.instance = new Chart(canvas.getContext('2d'), cfg);
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
  // Fallas por servidor
  const porServidor = {};
  list.filter(r=>r.estado==='Fallido' && r.servidor).forEach(r=>{
    porServidor[r.servidor] = (porServidor[r.servidor]||0) + 1;
  });
  const itemsServidor = Object.entries(porServidor)
    .map(([label,value])=>({label,value}))
    .sort((a,b)=>b.value-a.value)
    .slice(0,8);
  renderChartData('chartServidor', itemsServidor, {colorClass:'c-danger'});

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
  renderChartData('chartResponsable', itemsResp, {colorClass:'c-info', suffix:' min'});

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
  renderChartData('chartAmbiente', itemsAmbiente, {colorClass:'c-info', suffix:' min'});
}

// Suma, sobre los registros filtrados, cada columna de conteo (ETL,
// SCRIPT, Modificacion Job, Creacion Job, Restauracion BD...) detectada
// al importar. Así se responde "¿cuántos ETL, cuántos Job...?" de forma
// directa y ya filtrada por búsqueda/estado/tipo/archivo.
function renderActividades(list){
  const totales = {};
  list.forEach(r=>{
    const act = r.actividades || {};
    Object.entries(act).forEach(([k,v])=>{
      const label = cleanLabel(k);
      totales[label] = (totales[label]||0) + (typeof v === 'number' ? v : 0);
    });
  });
  const items = Object.entries(totales)
    .map(([label,value])=>({label,value}))
    .filter(i=>i.value > 0)
    .sort((a,b)=>b.value-a.value);
  renderChartData('chartActividades', items, {colorClass:'c-warn'});
}

function updateViews(){
  const list = getFilteredRegistros();
  renderStats(list);
  renderAnalysis(list);
  renderActividades(list);
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

['search','filterEstado','filterTipo','filterOrigen'].forEach(id=>{
  el(id).addEventListener('input', updateViews);
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
    reader.readAsBinaryString(file);
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
      const wb = XLSX.read(bin, {type:'binary', cellDates:true});
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

// --- Helpers de actividades para el informe (agrupan por etiqueta limpia,
// igual que el gráfico, para que "SCRIPT" de distintos meses/hojas se
// trate como una sola actividad aunque el encabezado original varíe un
// poco en espacios/saltos de línea) ---
const MESES_ES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

function getActivityLabels(){
  const set = new Set();
  registros.forEach(r=>{
    Object.keys(r.actividades||{}).forEach(k=> set.add(cleanLabel(k)));
  });
  return Array.from(set).sort((a,b)=>a.localeCompare(b,'es'));
}

function activityValueForLabel(r, label){
  let sum = 0;
  Object.entries(r.actividades||{}).forEach(([k,v])=>{
    if(cleanLabel(k)===label) sum += (typeof v==='number' ? v : 0);
  });
  return sum;
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

// --- Generación del PDF (parametrizada por periodo / tipo de informe) ---
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
      const sum = list.reduce((acc,r)=>acc + activityValueForLabel(r,label), 0);
      const registrosConEsa = list.filter(r=>activityValueForLabel(r,label) > 0).length;
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

  // Tabla de registros: en modo "actividad" la última columna muestra solo
  // las actividades seleccionadas por fila; en modo general muestra la
  // info adicional completa, como antes.
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
    startY: y + 10,
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
