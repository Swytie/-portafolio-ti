const PM_SHEETS = ['Alexia C','Alexis L','Zaira F','Cynthia M','Alejandra R'];
const PM_COLORS = ['#B07CC6','#2563EB','#1A9E6A','#C8A200','#D94040'];
const STATUS_COLORS = {
  'Ejecución':'#2563EB','Activo':'#1A9E6A','Planeación':'#C8A200',
  'Cierre':'#1A9E6A','En proceso de cierre':'#C8A200','Cerrado':'#1A9E6A',
  'Cancelado':'#D94040','En aprobación':'#803F85',
  'Inicio':'#A889AB','N/A':'#A889AB','-':'#A889AB','':'#A889AB'
};

// Orden oficial de fases y estados (ver guía de catálogo del portafolio)
const FASE_ORDER = ['Inicio','Planeación','Ejecución','Cierre'];
const ESTADO_ORDER = ['Parametrización','En aprobación','Iniciativa (RPP)','Inicio','Planeación','Ejecución','Cierre','En proceso de cierre','Cerrado','Cancelado','Activo'];
const sortByOrder = (arr, orderList) => {
  const norm = s => String(s||'').replace(/^(\d+|N\/A)\s*/i,'').trim();
  return [...arr].sort((a,b) => {
    const ia = orderList.indexOf(norm(a));
    const ib = orderList.indexOf(norm(b));
    const ra = ia === -1 ? 999 : ia;
    const rb = ib === -1 ? 999 : ib;
    if (ra !== rb) return ra - rb;
    return String(a).localeCompare(String(b));
  });
};

const SHAREPOINT_URL = 'https://portafolio-proxy.jorgeguerrerp90.workers.dev/';
const AUTO_REFRESH_MS = 60 * 1000;

let ALL = [], FILTERED = [], CHARTS = {};
let APP_INITIALIZED = false, CURRENT_TAB = 'general';
let LAST_VIEW_HASH = null;
const hashData = arr => JSON.stringify(arr);
const DEFAULT_SORT_COL = '_idx', DEFAULT_SORT_DIR = 1;
let SORT_COL = DEFAULT_SORT_COL, SORT_DIR = DEFAULT_SORT_DIR, SEARCH = '';
let TBL_FILTERS = {pm:'', tipo:'', fase:'', unidad:'', estado:''};
let refreshTimer = null, countdownTimer = null, nextRefreshAt = null;
let MACTI_PCT = null;

const $ = id => document.getElementById(id);
const pmColor = pm => PM_COLORS[PM_SHEETS.indexOf(pm)] || '#803F85';

const fmt = v => {
  if (!v || isNaN(v)) return '—';
  if (v >= 1e6) return '$' + (v/1e6).toFixed(1) + 'M';
  if (v >= 1000) return '$' + (v/1000).toFixed(0) + 'K';
  return '$' + Math.round(v).toLocaleString('es-MX');
};
const pct = v => Math.round((v||0)*100) + '%';
const toNum = v => {
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(String(v).replace(/[,$\s%]/g,''));
  return isNaN(n) ? 0 : n;
};
const normalizeTipoProy = v => {
  const s = String(v||'').trim();
  if (/^ryr$/i.test(s)) return 'RyR';
  if (/^estrat[ée]gico$/i.test(s)) return 'Estratégico';
  return s;
};
const badgeClass = e => {
  const s = (e||'').toLowerCase();
  if (s.includes('ejec')) return 'b-ejec';
  if (s.includes('plan')) return 'b-plan';
  if (s.includes('cierr')||s.includes('cerr')) return 'b-cierre';
  if (s.includes('cancel')) return 'b-cancel';
  if (s.includes('apro')||s.includes('inicio')) return 'b-apro';
  return 'b-other';
};
const desvClass = v => v > 0 ? 'desv-ok' : v < -0.1 ? 'desv-neg' : v < 0 ? 'desv-pos' : 'desv-ok';
const desvStr = v => {
  if (!v && v !== 0) return '—';
  const p = Math.round(v * 100);
  return (p > 0 ? '+' : '') + p + '%';
};
const excelDateToJS = v => new Date(Math.round((Number(v) - 25569) * 86400 * 1000) + new Date().getTimezoneOffset() * 60 * 1000);
const isGoLiveOverdue = d => {
  const v = d._fecha_golive;
  if (!v || v === '' || v === 'nan' || v === '-') return false;
  const s = String(v);
  if (s.toUpperCase().includes('INICIATIVA') || s.toUpperCase().includes('TBD')) return false;
  let date;
  if (!isNaN(v) && Number(v) > 1000) date = excelDateToJS(v);
  else { date = new Date(v); if (isNaN(date.getTime())) return false; }
  const today = new Date(); today.setHours(0,0,0,0);
  if (date >= today) return false;
  const fase = (d._fase || '').toLowerCase();
  const estado = (d._estado || '').toLowerCase();
  if (estado.includes('cancel')) return false;
  return !fase.includes('cierr') && !fase.includes('cerr');
};

const App = {

  splashConnect() {
    const splash = document.getElementById('splash');
    splash.classList.add('hide');
    setTimeout(() => { splash.style.display = 'none'; App.connectSharePoint(); }, 500);
  },

  splashManual() {
    document.getElementById('splash-file').click();
  },

  splashFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
      App.parseExcel(e.target.result);
      $('last-update').textContent = 'Cargado: ' + new Date().toLocaleTimeString('es-MX');
      const splash = document.getElementById('splash');
      splash.classList.add('hide');
      setTimeout(() => { splash.style.display = 'none'; }, 500);
    };
    reader.readAsArrayBuffer(file);
  },

  async connectSharePoint() {
    const btn = document.getElementById('btn-auto');
    btn.innerHTML = '<i class="ti ti-loader"></i> Conectando...';
    btn.disabled = true;
    const ok = await App.fetchSharePoint();
    if (ok) {
      btn.innerHTML = '<i class="ti ti-refresh"></i> Actualizar';
      btn.disabled = false;
      btn.onclick = () => { App.fetchSharePoint(); App.scheduleRefresh(); };
      App.scheduleRefresh();
    } else {
      btn.innerHTML = '<i class="ti ti-upload"></i> Cargar manual';
      btn.disabled = false;
      btn.onclick = () => document.getElementById('file-input').click();
    }
  },

  scheduleRefresh() {
    if (refreshTimer) clearInterval(refreshTimer);
    nextRefreshAt = Date.now() + AUTO_REFRESH_MS;
    refreshTimer = setInterval(() => {
      nextRefreshAt = Date.now() + AUTO_REFRESH_MS;
      App.fetchSharePoint();
    }, AUTO_REFRESH_MS);
    App.startCountdown();
  },

  startCountdown() {
    if (countdownTimer) clearInterval(countdownTimer);
    const tick = () => {
      const el = $('next-update');
      if (!el || !nextRefreshAt) return;
      const diff = Math.max(0, nextRefreshAt - Date.now());
      const m = Math.floor(diff / 60000);
      const s = Math.floor((diff % 60000) / 1000);
      el.textContent = `Próxima sync en ${m}:${String(s).padStart(2,'0')}`;
    };
    tick();
    countdownTimer = setInterval(tick, 1000);
  },

  async fetchSharePoint() {
    try {
      const res = await fetch(`${SHAREPOINT_URL}?t=${Date.now()}`, {cache:'no-store'});
      if (!res.ok) throw new Error(res.status);
      const buffer = await res.arrayBuffer();
      App.parseExcel(buffer);
      $('last-update').textContent = 'Actualizado: ' + new Date().toLocaleTimeString('es-MX');
      const splash = document.getElementById('splash');
      if (splash) { splash.classList.add('hide'); setTimeout(()=>{ splash.style.display='none'; },500); }
      return true;
    } catch(e) {
      console.error('SharePoint error:', e);
      alert('No se pudo conectar a SharePoint/OneDrive. Verifica que el vínculo siga siendo válido y que el archivo esté compartido como "Cualquier persona con el vínculo".');
      return false;
    }
  },

  loadFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
      App.parseExcel(e.target.result);
      $('last-update').textContent = 'Cargado: ' + new Date().toLocaleTimeString('es-MX');
    };
    reader.readAsArrayBuffer(file);
  },

  parseMACTI(wb) {
  if (!wb.SheetNames.includes('MACTI Checklist')) return;
  const raw = XLSX.utils.sheet_to_json(wb.Sheets['MACTI Checklist'], {header:1, defval:''});
  let totalReq = 0, totalSub = 0;
  window._mactiMap = {};
  const clean = v => String(v||'').replace(/[\u202a\u202b\u202c\u202d\u202e\u200b\u200c\u200d\ufeff]/g,'').trim();
  for (let i = 3; i < raw.length; i++) {
    const row = raw[i];
    const id     = clean(row[2]);
    const nombre = clean(row[3]).toLowerCase();
    const req    = typeof row[10] === 'number' ? row[10] : 0;
    const sub    = typeof row[11] === 'number' ? row[11] : 0;
    if (req > 0) { totalReq += req; totalSub += sub; }
    if (req > 0) {
      const val = Math.round((sub / req) * 10000) / 100;
      if (id && id !== '' && id !== 'None' && id !== '-') window._mactiMap[id] = val;
      if (nombre) window._mactiMap[nombre] = val;
    }
  }
  MACTI_PCT = totalReq > 0 ? Math.round((totalSub / totalReq) * 10000) / 100 : null;
},


  parseExcel(buffer) {
  const wb = XLSX.read(buffer, {type:'array'});
  ALL = [];
  App.parseMACTI(wb);
  PM_SHEETS.forEach(sheet => {
    if (!wb.SheetNames.includes(sheet)) return;
    const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheet], {header:1, defval:''});
    let hdr = -1;
    for (let i = 0; i < raw.length; i++) {
      if (raw[i].some(c => String(c).includes('Tipo de') || String(c).includes('Nombre del') || String(c).includes('Esfuerzo'))) { hdr = i; break; }
    }
    if (hdr < 0) return;
    const hdrs = raw[hdr];
    const SKIP = ['Gerencia TI','Gerencia Expansión TI','Información general','Tipo de Esfuerzo','Tipo de esfuerzo'];
    const isAlejandra = sheet === 'Alejandra R';
    const tipoCol = isAlejandra ? 1 : 0;
    const cleanStr = v => String(v||'').replace(/[\u202a\u202b\u202c\u202d\u202e\u200b\u200c\u200d\ufeff]/g,'').trim();
    for (let i = hdr+1; i < raw.length; i++) {
      const row = raw[i];
      const tipo = String(row[tipoCol]||'').trim();
      if (!tipo || tipo === 'nan' || SKIP.includes(tipo)) continue;
      const obj = {pm: sheet, tipo};
      hdrs.forEach((h,idx) => { obj[String(h).trim()] = row[idx] ?? ''; });
      const find = (...keys) => {
        for (const k of keys) {
          const h = hdrs.find(h => String(h).toLowerCase().includes(k.toLowerCase()));
          if (h !== undefined) return obj[String(h).trim()];
        }
        return '';
      };
      obj._nombre       = cleanStr(find('Nombre del proyecto','Nombre de proyecto'));
      obj._unidad       = cleanStr(find('Unidades a desplegar','Unidad'));
      obj._fase         = cleanStr(find('Fase actual','Fase'));
      obj._estado       = cleanStr(find('Estado','Estatus'));
      obj._tipo_proy    = normalizeTipoProy(cleanStr(find('Tipo de Proyecto','Tipo de proyecto')));
      obj._pct          = toNum(find('% real','% Real','%Real'));
      obj._pct_plan     = toNum(find('% planeado','% Planeado'));
      obj._desviacion   = obj._pct - obj._pct_plan;
      obj._monto        = toNum(find('Monto Aprobado TI','Monto TI','Monto aprobado TI'));
      obj._erogado      = toNum(find('Presupuesto erogado'));
      obj._comprometido = toNum(find('Presupuesto Comprometido','Presupuesto comprometido'));
      obj._disponible   = toNum(find('Presupuesto Disponible','Presupuesto disponible'));
      obj._rrpp         = cleanStr(find('RRPP'));
      obj._id           = cleanStr(find('# Proyecto en Oracle','# Proyecto','# proyecto en Oracle'));
      obj._direccion    = cleanStr(find('Dirección solicitante','Dirección'));
      obj._lider        = cleanStr(find('Lider TI','Líder TI'));
      obj._fecha_ini    = find('Fecha de inicio');
      obj._fecha_cie    = find('Fecha de cierre');
      obj._fecha_golive = find('Fecha Go Live');
      obj._comentarios  = cleanStr(find('Descripción','Comentarios'));
      obj._sponsor      = cleanStr(find('Sponsor'));
      obj._pm_gestor    = cleanStr(find('PM / Gestor','PM/Gestor'));
      obj._ruta_critica = cleanStr(find('Ruta Crítica','RUTA CRÍTICA'));
      obj._tipo_servicio= cleanStr(find('Tipo de Servicio'));
      obj._es_iniciativa = String(obj._fecha_golive||'').toUpperCase().includes('INICIATIVA');
      obj._tiene_rrpp    = obj._rrpp && obj._rrpp !== '-' && obj._rrpp !== '' && obj._rrpp !== 'nan';
      const cleanId     = obj._id && obj._id !== '-' && obj._id !== '' && obj._id !== 'nan' ? obj._id : null;
      const cleanNombre = obj._nombre.toLowerCase();
      const mactiKey    = cleanId || cleanNombre;
      obj._macti = (window._mactiMap && mactiKey && window._mactiMap[mactiKey] !== undefined)
        ? window._mactiMap[mactiKey]
        : null;
      if (obj._nombre && obj._nombre !== 'nan') { obj._idx = ALL.length; ALL.push(obj); }
    }
  });
  App.buildTabs();
  if (!APP_INITIALIZED) {
    APP_INITIALIZED = true;
    App.showTab('general');
  } else {
    App.refreshCurrentView();
  }
},

  refreshCurrentView() {
    const relevant = CURRENT_TAB === 'general' ? ALL : ALL.filter(d=>d.pm===CURRENT_TAB);
    const hash = hashData(relevant);
    if (hash === LAST_VIEW_HASH) return;
    LAST_VIEW_HASH = hash;
    if (CURRENT_TAB === 'general') {
      if (!$('f-pm')) { App.showTab('general'); return; }
      const pmSel = $('f-pm').value, unidadSel = $('f-unidad').value, tipoProySel = $('f-tipo-proy').value;
      const pms = [...new Set(ALL.map(d=>d.pm))];
      const unidades = [...new Set(ALL.map(d=>d._unidad).filter(Boolean))].sort();
      const tipoProy = [...new Set(ALL.map(d=>d._tipo_proy).filter(Boolean))].sort();
      $('f-pm').innerHTML = '<option value="">Todos los PMs</option>' + pms.map(p=>`<option ${p===pmSel?'selected':''}>${p}</option>`).join('');
      $('f-unidad').innerHTML = '<option value="">Todas las unidades</option>' + unidades.map(u=>`<option ${u===unidadSel?'selected':''}>${u}</option>`).join('');
      $('f-tipo-proy').innerHTML = '<option value="">Todas las categorías</option>' + tipoProy.map(t=>`<option ${t===tipoProySel?'selected':''}>${t}</option>`).join('');
      App.filterGeneral();
    } else if (ALL.some(d=>d.pm===CURRENT_TAB)) {
      App.destroyCharts();
      App.renderPM(CURRENT_TAB);
    } else {
      App.showTab('general');
    }
  },

  buildTabs() {
    const pms = [...new Set(ALL.map(d => d.pm))];
    $('tabs').innerHTML = `
      <div class="tab active" onclick="App.showTab('general')" id="tab-general">
        <i class="ti ti-layout-dashboard"></i> Vista General
      </div>
      ${pms.map(pm => {
        const color = pmColor(pm);
        return `
          <div class="tab" onclick="App.showTab('${pm}')" id="tab-${pm.replace(/ /g,'_')}">
            <span style="width:22px;height:22px;border-radius:50%;background:${color};color:#fff;display:inline-flex;align-items:center;justify-content:center;flex-shrink:0">
              <i class="ti ti-user" style="font-size:12px"></i>
            </span>
            ${pm}
          </div>`;
      }).join('')}
    `;
  },

  showTab(tab) {
    CURRENT_TAB = tab;
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    const tid = tab === 'general' ? 'tab-general' : 'tab-'+tab.replace(/ /g,'_');
    if ($(tid)) $(tid).classList.add('active');
    SEARCH = ''; SORT_COL = DEFAULT_SORT_COL; SORT_DIR = DEFAULT_SORT_DIR;
    TBL_FILTERS = {pm:'', tipo:'', fase:'', unidad:'', estado:''};
    App.destroyCharts();
    tab === 'general' ? App.renderGeneral() : App.renderPM(tab);
    LAST_VIEW_HASH = hashData(tab === 'general' ? ALL : ALL.filter(d=>d.pm===tab));
  },

  renderGeneral() {
    $('main').innerHTML = `
      <div class="toolbar">
        <span class="filter-label"><i class="ti ti-filter"></i> Filtrar:</span>
        <select class="fsel" id="f-pm" onchange="App.filterGeneral()"><option value="">Todos los PMs</option></select>
        <select class="fsel" id="f-tipo" onchange="App.filterGeneral()">
          <option value="">Todos los tipos</option><option>Proyecto</option><option>Backlog</option>
        </select>
        <select class="fsel" id="f-tipo-proy" onchange="App.filterGeneral()"><option value="">Todas las categorías</option></select>
        <select class="fsel" id="f-fase" onchange="App.filterGeneral()">
          <option value="">Todas las fases</option>
          <option>01 Inicio</option><option>02 Planeación</option><option>03 Ejecución</option><option>05 Cierre</option>
        </select>
        <select class="fsel" id="f-unidad" onchange="App.filterGeneral()"><option value="">Todas las unidades</option></select>
        <select class="fsel" id="f-estado" onchange="App.filterGeneral()">
          <option value="">Todos los estados</option>
          <option>Activo</option><option>Ejecución</option><option>Planeación</option>
          <option>Cierre</option><option>Cancelado</option><option>En aprobación</option>
        </select>
        <button class="btn-primary" onclick="App.exportMSR()" style="margin-left:auto">
          <i class="ti ti-file-download"></i> Exportar MSR_R
        </button>
      </div>
      <div class="kpi-grid" id="kpis"></div>
      <div class="status-grid" id="status-grid"></div>
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-bar"></i><span class="chart-title">Monto Aprobado TI por PM</span></div>
          <div class="chart-wrap"><canvas id="ch-pm"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-donut"></i><span class="chart-title">Por Estado</span></div>
          <div class="chart-wrap"><canvas id="ch-estado"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-pie"></i><span class="chart-title">RyR vs Estratégico</span></div>
          <div class="chart-wrap"><canvas id="ch-tipo"></canvas></div>
        </div>
      </div>
      <div class="charts-row-2">
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-bar"></i><span class="chart-title">Fase de los Proyectos</span></div>
          <div class="chart-wrap-lg"><canvas id="ch-fase"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-trending-up"></i><span class="chart-title">Avance Real vs Planeado por PM</span></div>
          <div class="chart-wrap-lg"><canvas id="ch-avance"></canvas></div>
        </div>
      </div>
      <div class="table-card">
        <div class="table-header">
          <div class="table-header-left"><i class="ti ti-table"></i><span class="table-title">Detalle de Proyectos</span></div>
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
            <div class="search-box">
              <i class="ti ti-search"></i>
              <input type="text" id="search-input" placeholder="Buscar proyecto..." value="${SEARCH}" oninput="App.onSearch(this.value)">
            </div>
            <span class="table-count" id="tbl-count"></span>
          </div>
        </div>
        <div class="tbl-filter-bar" id="tbl-filter-bar"></div>
        <div class="table-scroll-top" id="tbl-scroll-top"><div class="table-scroll-top-inner" id="tbl-scroll-top-inner"></div></div>
        <div class="table-scroll" id="tbl-scroll"><div id="tbl"></div></div>
      </div>`;
    const pms = [...new Set(ALL.map(d => d.pm))];
    const unidades = [...new Set(ALL.map(d => d._unidad).filter(Boolean))].sort();
    const tipoProy = [...new Set(ALL.map(d => d._tipo_proy).filter(Boolean))].sort();
    $('f-pm').innerHTML = '<option value="">Todos los PMs</option>' + pms.map(p => `<option>${p}</option>`).join('');
    $('f-unidad').innerHTML = '<option value="">Todas las unidades</option>' + unidades.map(u => `<option>${u}</option>`).join('');
    $('f-tipo-proy').innerHTML = '<option value="">Todas las categorías</option>' + tipoProy.map(t => `<option>${t}</option>`).join('');
    App.filterGeneral();
  },

  filterGeneral() {
    const pm       = $('f-pm').value;
    const tipo     = $('f-tipo').value;
    const tipoProy = $('f-tipo-proy').value;
    const fase     = $('f-fase').value;
    const unidad   = $('f-unidad').value;
    const estado = $('f-estado').value;
    FILTERED = ALL.filter(d => {
      if (pm       && d.pm !== pm) return false;
      if (tipo     && !d.tipo.toLowerCase().startsWith(tipo.toLowerCase())) return false;
      if (tipoProy && d._tipo_proy !== tipoProy) return false;
      if (fase     && d._fase !== fase) return false;
      if (unidad   && d._unidad !== unidad) return false;
      if (estado   && d._estado !== estado) return false;
      return true;
    });
    App.renderKPIs(FILTERED);
    App.renderStatusGrid(FILTERED);
    App.renderCharts(FILTERED);
    App.renderTblFilters(FILTERED);
    App.renderTable(App.applyTblFilters(FILTERED));
  },

  onSearch(val) {
    SEARCH = val.toLowerCase();
    App.renderTable(App.applyTblFilters(FILTERED));
  },

  renderTblFilters(data) {
  const unidades = [...new Set(data.map(d=>d._unidad).filter(Boolean))].sort();
  const estados  = sortByOrder([...new Set(data.map(d=>d._estado).filter(Boolean))], ESTADO_ORDER);
  const fases    = sortByOrder([...new Set(data.map(d=>d._fase).filter(Boolean))], FASE_ORDER);
  const tipos    = [...new Set(data.map(d=>d._tipo_proy).filter(Boolean))].sort();
  const pms      = [...new Set(data.map(d=>d.pm).filter(Boolean))];
  const sel = (id, opts, label) => `
    <select class="tbl-fsel" id="tf-${id}" onchange="App.onTblFilter('${id}',this.value)">
      <option value="">${label}</option>
      ${opts.map(o=>`<option value="${o}" ${TBL_FILTERS[id]===o?'selected':''}>${o}</option>`).join('')}
    </select>`;
  const pmSel = pms.length > 1 ? sel('pm', pms, 'PM') : '';
  const sortOptions = [
    {v:'_idx|1',     label:'Orden original (Excel)'},
    {v:'_nombre|1',  label:'Proyecto (A-Z)'},
    {v:'_nombre|-1', label:'Proyecto (Z-A)'},
    {v:'_estado|1',  label:'Estado (orden oficial)'},
    {v:'_fase|1',    label:'Fase (orden oficial)'},
    {v:'_fecha_golive|1', label:'Go Live (próxima primero)'},
    {v:'_monto|-1',  label:'Monto (mayor a menor)'},
    {v:'_pct|-1',    label:'% Avance (mayor a menor)'},
  ];
  const sortVal = `${SORT_COL}|${SORT_DIR}`;
  const sortSel = `
    <span class="tbl-fsel-label">Ordenar por:</span>
    <select class="tbl-fsel" id="tf-sort" onchange="App.onSortSelect(this.value)">
      ${sortOptions.map(o=>`<option value="${o.v}" ${sortVal===o.v?'selected':''}>${o.label}</option>`).join('')}
    </select>`;
  $('tbl-filter-bar').innerHTML = `
    <div class="tbl-filter-inner">
      <i class="ti ti-adjustments-horizontal" style="font-size:13px;color:var(--t3)"></i>
      ${pmSel}
      ${sel('tipo', tipos, 'Tipo')}
      ${sel('fase', fases, 'Fase')}
      ${sel('unidad', unidades, 'Unidad')}
      ${sel('estado', estados, 'Estado')}
      ${sortSel}
      <button class="tbl-clear-btn" onclick="App.clearTblFilters()"><i class="ti ti-x"></i> Limpiar</button>
    </div>`;
},

  onSortSelect(val) {
    if (!val) return;
    const [col, dir] = val.split('|');
    SORT_COL = col;
    SORT_DIR = Number(dir);
    App.renderTable(App.applyTblFilters(FILTERED));
  },

  onTblFilter(key, val) {
    TBL_FILTERS[key] = val;
    App.renderTable(App.applyTblFilters(FILTERED));
  },

  clearTblFilters() {
    TBL_FILTERS = {pm:'', tipo:'', fase:'', unidad:'', estado:''};
    SORT_COL = DEFAULT_SORT_COL; SORT_DIR = DEFAULT_SORT_DIR;
    App.renderTblFilters(FILTERED);
    App.renderTable(App.applyTblFilters(FILTERED));
  },

  applyTblFilters(data) {
    return data.filter(d => {
      if (TBL_FILTERS.pm     && d.pm !== TBL_FILTERS.pm) return false;
      if (TBL_FILTERS.tipo   && d._tipo_proy !== TBL_FILTERS.tipo) return false;
      if (TBL_FILTERS.fase   && d._fase !== TBL_FILTERS.fase) return false;
      if (TBL_FILTERS.unidad && d._unidad !== TBL_FILTERS.unidad) return false;
      if (TBL_FILTERS.estado && d._estado !== TBL_FILTERS.estado) return false;
      return true;
    });
  },

  renderKPIs(data) {
    const monto   = data.reduce((s,d)=>s+d._monto,0);
    const erogado = data.reduce((s,d)=>s+d._erogado,0);
    const disp    = data.reduce((s,d)=>s+d._disponible,0);
    const comp    = data.reduce((s,d)=>s+d._comprometido,0);
    const activos = data.filter(d=>d.tipo.toLowerCase().startsWith('proyecto'));
    const avg     = activos.length ? activos.reduce((s,d)=>s+d._pct,0)/activos.length : 0;
    const pctE    = monto > 0 ? Math.round(erogado/monto*100) : 0;
    const mactiHtml = MACTI_PCT !== null
      ? `<div class="kpi"><i class="ti ti-clipboard-check kpi-icon"></i><div class="kpi-label">MACTI</div><div class="kpi-val" style="color:${MACTI_PCT>=80?'#1A9E6A':MACTI_PCT>=50?'#C8A200':'#D94040'}">${MACTI_PCT}%</div><div class="kpi-sub">Cumplimiento</div></div>`
      : '';
    $('kpis').innerHTML = `
      <div class="kpi"><i class="ti ti-briefcase kpi-icon"></i><div class="kpi-label">Total Proyectos</div><div class="kpi-val">${data.length}</div><div class="kpi-sub">${activos.length} activos</div></div>
      <div class="kpi"><i class="ti ti-coin kpi-icon"></i><div class="kpi-label">Monto Aprobado TI</div><div class="kpi-val">${fmt(monto)}</div><div class="kpi-sub">MXN</div></div>
      <div class="kpi"><i class="ti ti-receipt kpi-icon"></i><div class="kpi-label">Erogado</div><div class="kpi-val">${fmt(erogado)}</div><div class="kpi-sub">${pctE}% del monto</div><div class="kpi-bar"><div class="kpi-bar-fill" style="width:${Math.min(pctE,100)}%"></div></div></div>
      <div class="kpi"><i class="ti ti-credit-card kpi-icon"></i><div class="kpi-label">Comprometido</div><div class="kpi-val">${fmt(comp)}</div><div class="kpi-sub">MXN</div></div>
      <div class="kpi"><i class="ti ti-trending-up kpi-icon"></i><div class="kpi-label">Avance Promedio</div><div class="kpi-val">${pct(avg)}</div><div class="kpi-sub">proyectos activos</div><div class="kpi-bar"><div class="kpi-bar-fill" style="width:${Math.round(avg*100)}%"></div></div></div>
      ${mactiHtml}`;
  },

  renderStatusGrid(data) {
    const activos    = data.filter(d=>['Activo','Ejecución'].some(s=>d._estado.includes(s))).length;
    const cerrados   = data.filter(d=>['Cierre','Cancelado'].some(s=>d._estado.includes(s))).length;
    const detenidos  = data.filter(d=>d._estado.includes('En aprobación')).length;
    const sinInic    = data.filter(d=>d._fase.includes('01') || d._estado.includes('Inicio')).length;
    const conRRPP    = data.filter(d=>d._tiene_rrpp).length;
    const iniciativas= data.filter(d=>d._es_iniciativa).length;
    $('status-grid').innerHTML = `
      <div class="status-card"><div class="status-dot" style="background:#1A9E6A"></div><div class="status-info"><div class="status-label">Proyectos Activos</div><div class="status-val">${activos}</div></div></div>
      <div class="status-card"><div class="status-dot" style="background:#803F85"></div><div class="status-info"><div class="status-label">Cerrados / Cancelados</div><div class="status-val">${cerrados}</div></div></div>
      <div class="status-card"><div class="status-dot" style="background:#C8A200"></div><div class="status-info"><div class="status-label">En Aprobación</div><div class="status-val">${detenidos}</div></div></div>
      <div class="status-card"><div class="status-dot" style="background:#A889AB"></div><div class="status-info"><div class="status-label">Por Iniciar</div><div class="status-val">${sinInic}</div></div></div>
      <div class="status-card"><div class="status-dot" style="background:#2563EB"></div><div class="status-info"><div class="status-label">Con RRPP</div><div class="status-val">${conRRPP}</div></div></div>
      <div class="status-card"><div class="status-dot" style="background:#9E5EA3"></div><div class="status-info"><div class="status-label">Iniciativas</div><div class="status-val">${iniciativas}</div></div></div>`;
  },

  destroyCharts() {
    ['pm','estado','tipo','fase','avance'].forEach(k => { if (CHARTS[k]) { CHARTS[k].destroy(); delete CHARTS[k]; } });
  },

  upsertChart(key, canvasId, config) {
    const existing = CHARTS[key];
    if (existing) {
      existing.data.labels = config.data.labels;
      config.data.datasets.forEach((ds, i) => {
        if (existing.data.datasets[i]) Object.assign(existing.data.datasets[i], ds);
        else existing.data.datasets[i] = ds;
      });
      existing.data.datasets.length = config.data.datasets.length;
      existing.update();
    } else {
      CHARTS[key] = new Chart($(canvasId), config);
    }
  },

  renderCharts(data) {
    const pms = [...new Set(ALL.map(d=>d.pm))];
    const montos = pms.map(p=>data.filter(d=>d.pm===p).reduce((s,d)=>s+d._monto,0));
    App.upsertChart('pm','ch-pm', {
      type:'bar',
      data:{labels:pms, datasets:[{data:montos, backgroundColor:pms.map(p=>pmColor(p)), borderRadius:5}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}},
        scales:{x:{ticks:{font:{size:10}}}, y:{ticks:{callback:v=>fmt(v),font:{size:9}}}}}
    });
    const est = {};
    data.forEach(d=>{const e=d._estado||'Sin estado'; est[e]=(est[e]||0)+1;});
    const eK = Object.keys(est).filter(k=>est[k]>0);
    App.upsertChart('estado','ch-estado', {
      type:'doughnut',
      data:{labels:eK, datasets:[{data:eK.map(k=>est[k]), backgroundColor:eK.map(k=>STATUS_COLORS[k]||'#A889AB'), borderWidth:0}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom',labels:{font:{size:9},boxWidth:8}}}}
    });
    const tip = {};
    data.forEach(d=>{const t=(d._tipo_proy||'').toLowerCase().includes('estrateg')?'Estratégico':'RyR'; tip[t]=(tip[t]||0)+1;});
    const tK = Object.keys(tip);
    App.upsertChart('tipo','ch-tipo', {
      type:'doughnut',
      data:{labels:tK, datasets:[{data:tK.map(k=>tip[k]), backgroundColor:['#803F85','#1A9E6A'], borderWidth:0}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom',labels:{font:{size:9},boxWidth:8}}}}
    });
    const fases = {};
    data.forEach(d=>{const f=d._fase||'Sin fase'; fases[f]=(fases[f]||0)+1;});
    const fK = Object.keys(fases).sort();
    App.upsertChart('fase','ch-fase', {
      type:'bar',
      data:{labels:fK, datasets:[{data:fK.map(k=>fases[k]), backgroundColor:'#803F85', borderRadius:4}]},
      options:{indexAxis:'y', responsive:true, maintainAspectRatio:false,
        plugins:{legend:{display:false}},
        scales:{x:{ticks:{font:{size:9}}}, y:{ticks:{font:{size:10}}}}}
    });
    const pmsAvance = [...new Set(ALL.map(d=>d.pm))];
    const realAvg = pmsAvance.map(p=>{
      const rows = data.filter(d=>d.pm===p && d.tipo.toLowerCase().startsWith('proyecto'));
      return rows.length ? Math.round(rows.reduce((s,d)=>s+d._pct,0)/rows.length*100) : 0;
    });
    const planAvg = pmsAvance.map(p=>{
      const rows = data.filter(d=>d.pm===p && d.tipo.toLowerCase().startsWith('proyecto'));
      return rows.length ? Math.round(rows.reduce((s,d)=>s+d._pct_plan,0)/rows.length*100) : 0;
    });
    App.upsertChart('avance','ch-avance', {
      type:'bar',
      data:{labels:pmsAvance, datasets:[
        {label:'Real', data:realAvg, backgroundColor:pmsAvance.map(p=>pmColor(p)), borderRadius:4},
        {label:'Planeado', data:planAvg, backgroundColor:'#E8DCE9', borderRadius:4},
      ]},
      options:{responsive:true, maintainAspectRatio:false,
        plugins:{legend:{position:'bottom',labels:{font:{size:9},boxWidth:8}}},
        scales:{x:{ticks:{font:{size:10}}}, y:{ticks:{callback:v=>v+'%',font:{size:9}}, max:100}}}
    });
  },

  renderPM(pm) {
    const data = ALL.filter(d => d.pm === pm);
    const ini = pm.split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase();
    const color = pmColor(pm);
    const totalMonto = data.reduce((s,d)=>s+d._monto,0);
    const totalErog  = data.reduce((s,d)=>s+d._erogado,0);
    const totalDisp  = data.reduce((s,d)=>s+d._disponible,0);
    const activos    = data.filter(d=>d.tipo.toLowerCase().startsWith('proyecto'));
    const avg        = activos.length ? activos.reduce((s,d)=>s+d._pct,0)/activos.length : 0;
    const avgPlan    = activos.length ? activos.reduce((s,d)=>s+d._pct_plan,0)/activos.length : 0;
    const desv       = avg - avgPlan;
    const unidades   = [...new Set(data.map(d=>d._unidad).filter(Boolean))];
    const conRRPP    = data.filter(d=>d._tiene_rrpp).length;
    const iniciativas= data.filter(d=>d._es_iniciativa).length;
    $('main').innerHTML = `
      <div class="pm-header" style="border-left:4px solid ${color}">
        <div class="pm-avatar" style="background:${color}">${ini}</div>
        <div>
          <div class="pm-name">${pm}</div>
          <div class="pm-meta">
            <span><i class="ti ti-briefcase"></i> ${data.length} proyectos</span>
            <span><i class="ti ti-building-hospital"></i> ${unidades.length} unidades</span>
            <span><i class="ti ti-trending-up"></i> ${pct(avg)} avance</span>
            ${conRRPP ? `<span style="color:${color}"><i class="ti ti-file-invoice"></i> ${conRRPP} con RRPP</span>` : ''}
            ${iniciativas ? `<span style="color:var(--t2)"><i class="ti ti-bulb"></i> ${iniciativas} iniciativas</span>` : ''}
            <span class="${desvClass(desv)}"><i class="ti ti-arrows-diff"></i> Desviación: ${desvStr(desv)}</span>
          </div>
        </div>
        <div class="pm-header-kpis">
          <div class="pm-mini-kpi"><span class="pm-mini-label"><i class="ti ti-coin"></i> Monto Aprobado TI</span><span class="pm-mini-val" style="color:${color}">${fmt(totalMonto)}</span></div>
          <div class="pm-mini-kpi"><span class="pm-mini-label"><i class="ti ti-receipt"></i> Erogado</span><span class="pm-mini-val" style="color:${color}">${fmt(totalErog)}</span></div>
          <div class="pm-mini-kpi"><span class="pm-mini-label"><i class="ti ti-wallet"></i> Disponible</span><span class="pm-mini-val" style="color:${color}">${fmt(totalDisp)}</span></div>
        </div>
        <button class="btn-primary" onclick="App.exportMSR()" style="margin-left:auto;background:${color}">
          <i class="ti ti-file-download"></i> Exportar MSR_R
        </button>
      </div>
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-bar"></i><span class="chart-title">Monto por Unidad</span></div>
          <div class="chart-wrap"><canvas id="ch-pm"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-donut"></i><span class="chart-title">Por Estado</span></div>
          <div class="chart-wrap"><canvas id="ch-estado"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-pie"></i><span class="chart-title">Por Fase</span></div>
          <div class="chart-wrap"><canvas id="ch-tipo"></canvas></div>
        </div>
      </div>
      <div class="table-card">
        <div class="table-header">
          <div class="table-header-left"><i class="ti ti-table"></i><span class="table-title">Proyectos de ${pm}</span></div>
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
            <div class="search-box">
              <i class="ti ti-search"></i>
              <input type="text" id="search-input" placeholder="Buscar proyecto..." value="${SEARCH}" oninput="App.onSearch(this.value)">
            </div>
            <span class="table-count" id="tbl-count"></span>
          </div>
        </div>
        <div class="tbl-filter-bar" id="tbl-filter-bar"></div>
        <div class="table-scroll-top" id="tbl-scroll-top"><div class="table-scroll-top-inner" id="tbl-scroll-top-inner"></div></div>
        <div class="table-scroll" id="tbl-scroll"><div id="tbl"></div></div>
      </div>`;
    FILTERED = data;
    App.renderChartsPM(data, color);
    App.renderTblFilters(data);
    App.renderTable(App.applyTblFilters(data));
  },

  renderChartsPM(data, color) {
    const uds = [...new Set(data.map(d=>d._unidad).filter(Boolean))];
    const montos = uds.map(u=>data.filter(d=>d._unidad===u).reduce((s,d)=>s+d._monto,0));
    App.upsertChart('pm','ch-pm', {
      type:'bar',
      data:{labels:uds, datasets:[{data:montos, backgroundColor:color||'#803F85', borderRadius:5}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}},
        scales:{x:{ticks:{font:{size:9},maxRotation:30}}, y:{ticks:{callback:v=>fmt(v),font:{size:9}}}}}
    });
    const est = {};
    data.forEach(d=>{const e=d._estado||'Sin estado'; est[e]=(est[e]||0)+1;});
    const eK = Object.keys(est).filter(k=>est[k]>0);
    App.upsertChart('estado','ch-estado', {
      type:'doughnut',
      data:{labels:eK, datasets:[{data:eK.map(k=>est[k]), backgroundColor:eK.map(k=>STATUS_COLORS[k]||'#A889AB'), borderWidth:0}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom',labels:{font:{size:9},boxWidth:8}}}}
    });
    const fases = {};
    data.forEach(d=>{const f=d._fase||'Sin fase'; fases[f]=(fases[f]||0)+1;});
    const fK = Object.keys(fases).filter(k=>fases[k]>0);
    App.upsertChart('tipo','ch-tipo', {
      type:'doughnut',
      data:{labels:fK, datasets:[{data:fK.map(k=>fases[k]), backgroundColor:PM_COLORS, borderWidth:0}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom',labels:{font:{size:9},boxWidth:8}}}}
    });
  },

  sortTable(col) {
    if (SORT_COL === col) SORT_DIR *= -1;
    else { SORT_COL = col; SORT_DIR = col === '_nombre' ? 1 : -1; }
    App.renderTable(App.applyTblFilters(FILTERED));
  },

  renderTable(data) {
  const search = SEARCH.toLowerCase();
  const visible = search
    ? data.filter(d => d._nombre.toLowerCase().includes(search) || (d._unidad||'').toLowerCase().includes(search) || d.pm.toLowerCase().includes(search))
    : data;
  $('tbl-count').textContent = visible.length + ' registros';
  if (!visible.length) {
    $('tbl').innerHTML = '<div class="empty-state"><i class="ti ti-search empty-icon"></i><div class="empty-title">Sin resultados</div></div>';
    return;
  }
  const ORDER_MAP = {_estado: ESTADO_ORDER, _fase: FASE_ORDER};
  const normOrder = s => String(s||'').replace(/^(\d+|N\/A)\s*/i,'').trim();
  const sorted = [...visible].sort((a,b) => {
    const orderList = ORDER_MAP[SORT_COL];
    if (orderList) {
      const ia = orderList.indexOf(normOrder(a[SORT_COL]));
      const ib = orderList.indexOf(normOrder(b[SORT_COL]));
      const ra = ia === -1 ? 999 : ia;
      const rb = ib === -1 ? 999 : ib;
      if (ra !== rb) return (ra - rb) * SORT_DIR;
      return String(a[SORT_COL]||'').localeCompare(String(b[SORT_COL]||'')) * SORT_DIR;
    }
    const va = a[SORT_COL] || '';
    const vb = b[SORT_COL] || '';
    if (typeof va === 'number') return (va - vb) * SORT_DIR;
    return String(va).localeCompare(String(vb)) * SORT_DIR;
  });
  const cols = [
    {label:'', key:'_alert', noSort:true},
    {label:'# Proyecto', key:'_id'},
    {label:'Nombre del Proyecto', key:'_nombre'},
    {label:'PM', key:'pm'},
    {label:'Unidad', key:'_unidad'},
    {label:'Fase', key:'_fase'},
    {label:'Estado', key:'_estado'},
    {label:'Tipo', key:'_tipo_proy'},
    {label:'RRPP', key:'_rrpp'},
    {label:'Fecha Inicio', key:'_fecha_ini'},
    {label:'Fecha Cierre', key:'_fecha_cie'},
    {label:'Fecha Go Live', key:'_fecha_golive'},
    {label:'Monto Aprobado TI', key:'_monto'},
    {label:'Erogado', key:'_erogado'},
    {label:'Disponible', key:'_disponible'},
    {label:'% Real', key:'_pct'},
    {label:'% Plan', key:'_pct_plan'},
    {label:'Desviación', key:'_desviacion'},
    {label:'% MACTI', key:'_macti'},
  ];
  const arrow = key => {
    if (SORT_COL !== key) return '<i class="ti ti-arrows-sort" style="opacity:.3;font-size:11px"></i>';
    return SORT_DIR === 1
      ? '<i class="ti ti-sort-ascending" style="font-size:11px;color:var(--p)"></i>'
      : '<i class="ti ti-sort-descending" style="font-size:11px;color:var(--p)"></i>';
  };
  const fmtFecha = v => {
    if (!v || v === '' || v === 'nan' || v === '-') return '—';
    const s = String(v);
    if (s.toUpperCase().includes('INICIATIVA') || s.toUpperCase().includes('TBD')) {
      return `<span style="font-size:10px;background:#F3EAF4;color:#803F85;padding:2px 6px;border-radius:4px">Iniciativa</span>`;
    }
    if (!isNaN(v) && Number(v) > 1000) {
      const d = new Date(Math.round((Number(v) - 25569) * 86400 * 1000) + new Date().getTimezoneOffset() * 60 * 1000);
      return d.toLocaleDateString('es-MX', {day:'2-digit', month:'short', year:'numeric'});
    }
    try {
      const d = new Date(v);
      if (!isNaN(d.getTime()) && d.getFullYear() > 1970) {
        return d.toLocaleDateString('es-MX', {day:'2-digit', month:'short', year:'numeric'});
      }
    } catch(e) {}
    return s;
  };
  $('tbl').innerHTML = `
    <table>
      <thead><tr>
        ${cols.map(c=>c.noSort
          ? `<th style="width:28px"></th>`
          : `<th onclick="App.sortTable('${c.key}')" style="cursor:pointer;user-select:none">${c.label} ${arrow(c.key)}</th>`
        ).join('')}
      </tr></thead>
      <tbody>
        ${sorted.map(d=>{
          const overdue = isGoLiveOverdue(d);
          const overdueTitle = 'Fecha Go Live pasada y la fase aún no es Cierre';
          return `
          <tr ${overdue ? `style="background:#FDECEC" title="${overdueTitle}"` : ''}>
            <td style="width:28px;text-align:center">${overdue ? `<i class="ti ti-alert-triangle-filled" style="color:#D94040;font-size:14px" title="Fecha Go Live pasada"></i>` : ''}</td>
            <td><span style="font-family:'DM Mono',monospace;font-size:11px">${d._id||'—'}</span></td>
            <td title="${d._nombre}">${d._nombre||'—'}</td>
            <td><span class="chip" style="background:${pmColor(d.pm)}18;color:${pmColor(d.pm)};border:1px solid ${pmColor(d.pm)}33">${d.pm}</span></td>
            <td title="${d._unidad}">${d._unidad||'—'}</td>
            <td>${d._fase||'—'}</td>
            <td><span class="badge ${badgeClass(d._estado)}">${d._estado||'—'}</span></td>
            <td>${d._tipo_proy||'—'}</td>
            <td><span style="font-family:'DM Mono',monospace;font-size:11px">${d._rrpp||'—'}</span></td>
            <td style="font-size:11px">${fmtFecha(d._fecha_ini)}</td>
            <td style="font-size:11px">${fmtFecha(d._fecha_cie)}</td>
            <td style="font-size:11px">${fmtFecha(d._fecha_golive)}</td>
            <td>${fmt(d._monto)}</td>
            <td>${fmt(d._erogado)}</td>
            <td>${fmt(d._disponible)}</td>
            <td><div class="prog"><div class="prog-bar"><div class="prog-fill" style="width:${Math.round((d._pct||0)*100)}%;background:${pmColor(d.pm)}"></div></div><span class="prog-val">${pct(d._pct)}</span></div></td>
            <td><span class="prog-val">${pct(d._pct_plan)}</span></td>
            <td><span class="${desvClass(d._desviacion)}">${desvStr(d._desviacion)}</span></td>
            <td>${d._macti !== null && d._macti !== undefined
              ? `<span style="font-size:11px;font-weight:600;color:${d._macti>=80?'#1A9E6A':d._macti>=50?'#C8A200':'#D94040'}">${d._macti.toFixed(2)}%</span>`
              : '<span style="color:var(--t3);font-size:11px">—</span>'
            }</td>
          </tr>`;
        }).join('')}
      </tbody>
    </table>`;

  const scrollBottom = $('tbl-scroll');
  const scrollTop = $('tbl-scroll-top');
  const scrollTopInner = $('tbl-scroll-top-inner');
  if (scrollBottom && scrollTop && scrollTopInner) {
    const table = scrollBottom.querySelector('table');
    scrollTopInner.style.width = table.scrollWidth + 'px';
    scrollTop.onscroll = () => { scrollBottom.scrollLeft = scrollTop.scrollLeft; };
    scrollBottom.onscroll = () => { scrollTop.scrollLeft = scrollBottom.scrollLeft; };
  }
},

  async exportMSR() {
    if (!FILTERED.length) return;
    if (typeof ExcelJS === 'undefined') {
      alert('La librería ExcelJS no cargó. Verifica que el dashboard esté abierto con conexión a internet.');
      return;
    }
    try {

    const P1='FF5C2B60', P2='FF803F85', DB1='FF1F3864', DB2='FF2E4E8A';

    const fmtDate = v => {
      if (!v || v==='' || v==='nan' || v==='-') return '';
      const s = String(v);
      if (s.toUpperCase().includes('INICIATIVA')||s.toUpperCase().includes('TBD')) return s;
      if (!isNaN(v) && Number(v)>1000) {
        const d = new Date(Math.round((Number(v)-25569)*86400*1000)+new Date().getTimezoneOffset()*60*1000);
        return d.toLocaleDateString('es-MX',{day:'2-digit',month:'short',year:'numeric'});
      }
      try { const d=new Date(v); if(!isNaN(d.getTime())&&d.getFullYear()>1970) return d.toLocaleDateString('es-MX',{day:'2-digit',month:'short',year:'numeric'}); } catch(e){}
      return s;
    };

    const SEC = [
      {label:'Información del Proyecto', bg:P1,  sub:P2,  start:0,  end:10},
      {label:'Presupuesto',              bg:DB1, sub:DB2, start:11, end:17},
    ];
    const COLS = [
      {label:'# Proyecto',          sec:0, v:d=>d._id||'',        w:14},
      {label:'Nombre del Proyecto', sec:0, v:d=>d._nombre||'',    w:42},
      {label:'PM',                  sec:0, v:d=>d.pm||'',         w:14},
      {label:'Unidad',              sec:0, v:d=>d._unidad||'',    w:22},
      {label:'Tipo',                sec:0, v:d=>d._tipo_proy||'', w:16},
      {label:'RRPP',                sec:0, v:d=>d._rrpp||'',      w:10},
      {label:'Fase',                sec:0, v:d=>d._fase||'',      w:18},
      {label:'Estado',              sec:0, v:d=>d._estado||'',    w:20},
      {label:'Fecha Inicio',        sec:0, v:d=>fmtDate(d._fecha_ini),     w:16},
      {label:'Fecha Cierre',        sec:0, v:d=>fmtDate(d._fecha_cie),     w:16},
      {label:'Fecha Go Live',       sec:0, v:d=>fmtDate(d._fecha_golive),  w:16},
      {label:'Monto Aprobado TI',   sec:1, v:d=>d._monto||0,        num:true, w:20},
      {label:'Erogado',             sec:1, v:d=>d._erogado||0,      num:true, w:16},
      {label:'Disponible',          sec:1, v:d=>d._disponible||0,   num:true, w:16},
      {label:'% Real',              sec:1, v:d=>d._pct!=null?Math.round(d._pct*100)+'%':'',                                                        center:true, w:10},
      {label:'% Plan',              sec:1, v:d=>d._pct_plan!=null?Math.round(d._pct_plan*100)+'%':'',                                              center:true, w:10},
      {label:'Desviación',          sec:1, v:d=>d._desviacion!=null?(Math.round(d._desviacion*100)>0?'+':'')+Math.round(d._desviacion*100)+'%':'', center:true, w:12},
      {label:'% MACTI',             sec:1, v:d=>d._macti!=null?d._macti.toFixed(2)+'%':'',                                                         center:true, w:10},
    ];

    const thinBorder = {top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}};

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Portafolio');
    ws.columns = COLS.map(col => ({width: col.w || 16}));

    // Fila 1 — encabezados de sección
    const r1 = ws.addRow(COLS.map((col, i) => SEC.find(s=>s.start===i)?.label || ''));
    r1.height = 30;
    r1.eachCell({includeEmpty:true}, (cell, cn) => {
      const sec = SEC.find(s => (cn-1) >= s.start && (cn-1) <= s.end) || SEC[0];
      cell.fill      = {type:'pattern', pattern:'solid', fgColor:{argb:sec.bg}};
      cell.font      = {name:'Calibri', bold:true, size:12, color:{argb:'FFFFFFFF'}};
      cell.alignment = {horizontal:'center', vertical:'middle', wrapText:true};
      cell.border    = thinBorder;
    });
    SEC.forEach(sec => { ws.mergeCells(1, sec.start+1, 1, sec.end+1); });

    // Fila 2 — nombres de columna
    const r2 = ws.addRow(COLS.map(col => col.label));
    r2.height = 26;
    r2.eachCell({includeEmpty:true}, (cell, cn) => {
      const sec = SEC.find(s => (cn-1) >= s.start && (cn-1) <= s.end) || SEC[0];
      cell.fill      = {type:'pattern', pattern:'solid', fgColor:{argb:sec.sub}};
      cell.font      = {name:'Calibri', bold:true, size:10, color:{argb:'FFFFFFFF'}};
      cell.alignment = {horizontal:'center', vertical:'middle', wrapText:true};
      cell.border    = thinBorder;
    });

    // Filas de datos
    FILTERED.forEach(d => {
      const dataRow = ws.addRow(COLS.map(col => col.v(d)));
      dataRow.height = 18;
      dataRow.eachCell({includeEmpty:true}, (cell, cn) => {
        const col = COLS[cn-1];
        cell.font   = {name:'Calibri', size:10};
        cell.border = thinBorder;
        if (col.num) {
          cell.numFmt    = '"$"#,##0.00';
          cell.alignment = {horizontal:'right', vertical:'middle'};
        } else if (col.center) {
          cell.alignment = {horizontal:'center', vertical:'middle'};
        } else {
          cell.alignment = {horizontal:'left', vertical:'middle'};
        }
      });
    });

    ws.views = [{state:'frozen', ySplit:2}];

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Portafolio_TI_${new Date().toISOString().slice(0,10)}.xlsx`;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);

    } catch(err) {
      alert('Error al exportar: ' + err.message);
      console.error(err);
    }
  },
};
