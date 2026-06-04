const PM_SHEETS = ['Alexia C','Alexis L','Zaira F','Cynthia M','Alejandra R'];
const PM_COLORS = ['#803F85','#9E5EA3','#1A9E6A','#C8A200','#D94040'];
const STATUS_COLORS = {
  'Ejecución':'#2563EB','Activo':'#1A9E6A','Planeación':'#C8A200',
  'Cierre':'#1A9E6A','Cancelado':'#D94040','En aprobación':'#803F85',
  'Inicio':'#A889AB','N/A':'#A889AB','-':'#A889AB','':'#A889AB'
};

const SHAREPOINT_URL = 'https://christusmx-my.sharepoint.com/:x:/g/personal/jorgez_guerrero_christus_mx/IQB2sHTOZELzSZnL9YDxAk-LAXYSs26VmOC0IKQxCU1lIAI?e=8i20Ff&download=1';
const AUTO_REFRESH_MS = 5 * 60 * 1000;

let ALL = [], FILTERED = [], CHARTS = {};
let SORT_COL = '_monto', SORT_DIR = -1, SEARCH = '';
let refreshTimer = null;

const $ = id => document.getElementById(id);

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
const badgeClass = e => {
  const s = (e||'').toLowerCase();
  if (s.includes('ejec')) return 'b-ejec';
  if (s.includes('plan')) return 'b-plan';
  if (s.includes('cierr')||s.includes('cerr')) return 'b-cierre';
  if (s.includes('cancel')) return 'b-cancel';
  if (s.includes('apro')||s.includes('inicio')) return 'b-apro';
  return 'b-other';
};
const desvClass = v => v > 0 ? 'desv-pos' : v < -0.05 ? 'desv-neg' : 'desv-ok';
const desvStr = v => {
  if (!v && v !== 0) return '—';
  const p = Math.round(v * 100);
  return (p > 0 ? '+' : '') + p + '%';
};

const App = {


    async connectSharePoint() {
  const btn = document.getElementById('btn-auto');
  btn.innerHTML = '<i class="ti ti-loader" aria-hidden="true"></i> Conectando...';
  btn.disabled = true;
  const ok = await App.fetchSharePoint();
  if (ok) {
    btn.innerHTML = '<i class="ti ti-refresh" aria-hidden="true"></i> Actualizar';
    btn.disabled = false;
    btn.onclick = () => App.fetchSharePoint();
    if (refreshTimer) clearInterval(refreshTimer);
    refreshTimer = setInterval(App.fetchSharePoint, AUTO_REFRESH_MS);
  } else {
    btn.innerHTML = '<i class="ti ti-upload" aria-hidden="true"></i> Cargar manual';
    btn.disabled = false;
    btn.onclick = () => document.getElementById('file-input').click();
  }
},

async fetchSharePoint() {
  try {
    const res = await fetch(SHAREPOINT_URL, {
      mode: 'no-cors',
      credentials: 'include',
      cache: 'no-store'
    });
    const buffer = await res.arrayBuffer();
    App.parseExcel(buffer);
    $('last-update').textContent = 'Actualizado: ' + new Date().toLocaleTimeString('es-MX');
    return true;
  } catch(e) {
    console.error('Error:', e);
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

  parseExcel(buffer) {
    const wb = XLSX.read(buffer, {type:'array'});
    ALL = [];
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
        obj._nombre       = String(find('Nombre del proyecto','Nombre de proyecto')||'').trim();
        obj._unidad       = String(find('Unidades a desplegar','Unidad')||'').trim();
        obj._fase         = String(find('Fase actual','Fase')||'').trim();
        obj._estado       = String(find('Estado','Estatus')||'').trim();
        obj._tipo_proy    = String(find('Tipo de Proyecto','Tipo de proyecto')||'').trim();
        obj._pct          = toNum(find('% real','% Real','%Real'));
        obj._pct_plan     = toNum(find('% planeado','% Planeado'));
        obj._desviacion   = obj._pct - obj._pct_plan;
        obj._monto        = toNum(find('Monto Aprobado TI','Monto TI','Monto aprobado TI'));
        obj._erogado      = toNum(find('Presupuesto erogado'));
        obj._comprometido = toNum(find('Presupuesto Comprometido','Presupuesto comprometido'));
        obj._disponible   = toNum(find('Presupuesto Disponible','Presupuesto disponible'));
        obj._rrpp         = String(find('RRPP')||'').trim();
        obj._id           = String(find('# Proyecto en Oracle','# Proyecto','# proyecto en Oracle')||'').trim();
        obj._direccion    = String(find('Dirección solicitante','Dirección')||'').trim();
        obj._lider        = String(find('Lider TI','Líder TI')||'').trim();
        obj._fecha_ini    = find('Fecha de inicio');
        obj._fecha_cie    = find('Fecha de cierre');
        obj._fecha_golive = find('Fecha Go Live');
        obj._comentarios  = String(find('Descripción','Comentarios')||'').trim();
        obj._sponsor      = String(find('Sponsor')||'').trim();
        obj._pm_gestor    = String(find('PM / Gestor','PM/Gestor')||'').trim();
        obj._ruta_critica = String(find('Ruta Crítica','RUTA CRÍTICA')||'').trim();
        obj._tipo_servicio= String(find('Tipo de Servicio')||'').trim();
        if (obj._nombre && obj._nombre !== 'nan') ALL.push(obj);
      }
    });
    App.buildTabs();
    App.showTab('general');
  },

  buildTabs() {
    const pms = [...new Set(ALL.map(d => d.pm))];
    $('tabs').innerHTML = `
      <div class="tab active" onclick="App.showTab('general')" id="tab-general">
        <i class="ti ti-layout-dashboard" aria-hidden="true"></i> Vista General
      </div>
      ${pms.map(pm => `
        <div class="tab" onclick="App.showTab('${pm}')" id="tab-${pm.replace(/ /g,'_')}">
          <i class="ti ti-user-circle" aria-hidden="true"></i> ${pm}
        </div>`).join('')}
    `;
  },

  showTab(tab) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    const tid = tab === 'general' ? 'tab-general' : 'tab-'+tab.replace(/ /g,'_');
    if ($(tid)) $(tid).classList.add('active');
    SEARCH = ''; SORT_COL = '_monto'; SORT_DIR = -1;
    tab === 'general' ? App.renderGeneral() : App.renderPM(tab);
  },

  renderGeneral() {
    $('main').innerHTML = `
      <div class="toolbar">
        <span class="filter-label"><i class="ti ti-filter" aria-hidden="true"></i> Filtrar:</span>
        <select class="fsel" id="f-pm" onchange="App.filterGeneral()"><option value="">Todos los PMs</option></select>
        <select class="fsel" id="f-tipo" onchange="App.filterGeneral()">
          <option value="">Todos los tipos</option><option>Proyecto</option><option>Backlog</option>
        </select>
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
          <i class="ti ti-file-download" aria-hidden="true"></i> Exportar MSR_R
        </button>
      </div>
      <div class="kpi-grid" id="kpis"></div>
      <div class="status-grid" id="status-grid"></div>
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-bar" aria-hidden="true"></i><span class="chart-title">Monto Aprobado TI por PM</span></div>
          <div class="chart-wrap"><canvas id="ch-pm" role="img" aria-label="Monto aprobado TI por PM"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-donut" aria-hidden="true"></i><span class="chart-title">Por Estado</span></div>
          <div class="chart-wrap"><canvas id="ch-estado" role="img" aria-label="Proyectos por estado"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-pie" aria-hidden="true"></i><span class="chart-title">RyR vs Estratégico</span></div>
          <div class="chart-wrap"><canvas id="ch-tipo" role="img" aria-label="Proyectos por tipo"></canvas></div>
        </div>
      </div>
      <div class="charts-row-2">
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-bar" aria-hidden="true"></i><span class="chart-title">Fase de los Proyectos</span></div>
          <div class="chart-wrap-lg"><canvas id="ch-fase" role="img" aria-label="Proyectos por fase"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-trending-up" aria-hidden="true"></i><span class="chart-title">Avance Real vs Planeado por PM</span></div>
          <div class="chart-wrap-lg"><canvas id="ch-avance" role="img" aria-label="Avance real vs planeado"></canvas></div>
        </div>
      </div>
      <div class="table-card">
        <div class="table-header">
          <div class="table-header-left"><i class="ti ti-table" aria-hidden="true"></i><span class="table-title">Detalle de Proyectos</span></div>
          <div style="display:flex;align-items:center;gap:8px">
            <div class="search-box">
              <i class="ti ti-search" aria-hidden="true"></i>
              <input type="text" id="search-input" placeholder="Buscar proyecto..." oninput="App.onSearch(this.value)">
            </div>
            <span class="table-count" id="tbl-count"></span>
          </div>
        </div>
        <div class="table-scroll"><div id="tbl"></div></div>
      </div>`;
    const pms = [...new Set(ALL.map(d => d.pm))];
    const unidades = [...new Set(ALL.map(d => d._unidad).filter(Boolean))].sort();
    $('f-pm').innerHTML = '<option value="">Todos los PMs</option>' + pms.map(p => `<option>${p}</option>`).join('');
    $('f-unidad').innerHTML = '<option value="">Todas las unidades</option>' + unidades.map(u => `<option>${u}</option>`).join('');
    App.filterGeneral();
  },

  filterGeneral() {
    const pm     = $('f-pm').value;
    const tipo   = $('f-tipo').value;
    const fase   = $('f-fase').value;
    const unidad = $('f-unidad').value;
    const estado = $('f-estado').value;
    FILTERED = ALL.filter(d => {
      if (pm     && d.pm !== pm) return false;
      if (tipo   && !d.tipo.toLowerCase().startsWith(tipo.toLowerCase())) return false;
      if (fase   && d._fase !== fase) return false;
      if (unidad && d._unidad !== unidad) return false;
      if (estado && d._estado !== estado) return false;
      return true;
    });
    App.renderKPIs(FILTERED);
    App.renderStatusGrid(FILTERED);
    App.renderCharts(FILTERED);
    App.renderTable(FILTERED);
  },

  onSearch(val) {
    SEARCH = val.toLowerCase();
    App.renderTable(FILTERED);
  },

  renderKPIs(data) {
    const monto   = data.reduce((s,d)=>s+d._monto,0);
    const erogado = data.reduce((s,d)=>s+d._erogado,0);
    const disp    = data.reduce((s,d)=>s+d._disponible,0);
    const comp    = data.reduce((s,d)=>s+d._comprometido,0);
    const activos = data.filter(d=>d.tipo.toLowerCase().startsWith('proyecto'));
    const avg     = activos.length ? activos.reduce((s,d)=>s+d._pct,0)/activos.length : 0;
    const pctE    = monto > 0 ? Math.round(erogado/monto*100) : 0;
    $('kpis').innerHTML = `
      <div class="kpi"><i class="ti ti-briefcase kpi-icon" aria-hidden="true"></i><div class="kpi-label">Total Proyectos</div><div class="kpi-val">${data.length}</div><div class="kpi-sub">${activos.length} activos</div></div>
      <div class="kpi"><i class="ti ti-coin kpi-icon" aria-hidden="true"></i><div class="kpi-label">Monto Aprobado TI</div><div class="kpi-val">${fmt(monto)}</div><div class="kpi-sub">MXN</div></div>
      <div class="kpi"><i class="ti ti-receipt kpi-icon" aria-hidden="true"></i><div class="kpi-label">Erogado</div><div class="kpi-val">${fmt(erogado)}</div><div class="kpi-sub">${pctE}% del monto</div><div class="kpi-bar"><div class="kpi-bar-fill" style="width:${Math.min(pctE,100)}%"></div></div></div>
      <div class="kpi"><i class="ti ti-credit-card kpi-icon" aria-hidden="true"></i><div class="kpi-label">Comprometido</div><div class="kpi-val">${fmt(comp)}</div><div class="kpi-sub">MXN</div></div>
      <div class="kpi"><i class="ti ti-trending-up kpi-icon" aria-hidden="true"></i><div class="kpi-label">Avance Promedio</div><div class="kpi-val">${pct(avg)}</div><div class="kpi-sub">proyectos activos</div><div class="kpi-bar"><div class="kpi-bar-fill" style="width:${Math.round(avg*100)}%"></div></div></div>`;
  },

  renderStatusGrid(data) {
    const activos   = data.filter(d=>['Activo','Ejecución'].some(s=>d._estado.includes(s))).length;
    const cerrados  = data.filter(d=>['Cierre','Cancelado'].some(s=>d._estado.includes(s))).length;
    const detenidos = data.filter(d=>d._estado.includes('En aprobación')).length;
    const sinInic   = data.filter(d=>d._fase.includes('01') || d._estado.includes('Inicio')).length;
    $('status-grid').innerHTML = `
      <div class="status-card"><div class="status-dot" style="background:#1A9E6A"></div><div class="status-info"><div class="status-label">Proyectos Activos</div><div class="status-val">${activos}</div></div></div>
      <div class="status-card"><div class="status-dot" style="background:#803F85"></div><div class="status-info"><div class="status-label">Cerrados / Cancelados</div><div class="status-val">${cerrados}</div></div></div>
      <div class="status-card"><div class="status-dot" style="background:#C8A200"></div><div class="status-info"><div class="status-label">En Aprobación</div><div class="status-val">${detenidos}</div></div></div>
      <div class="status-card"><div class="status-dot" style="background:#A889AB"></div><div class="status-info"><div class="status-label">Por Iniciar</div><div class="status-val">${sinInic}</div></div></div>`;
  },

  destroyCharts() {
    ['pm','estado','tipo','fase','avance'].forEach(k => { if (CHARTS[k]) { CHARTS[k].destroy(); delete CHARTS[k]; } });
  },

  renderCharts(data) {
    App.destroyCharts();
    const pms = [...new Set(ALL.map(d=>d.pm))];
    const montos = pms.map(p=>data.filter(d=>d.pm===p).reduce((s,d)=>s+d._monto,0));
    CHARTS.pm = new Chart($('ch-pm'), {
      type:'bar',
      data:{labels:pms, datasets:[{data:montos, backgroundColor:PM_COLORS, borderRadius:5}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}},
        scales:{x:{ticks:{font:{size:10}}}, y:{ticks:{callback:v=>fmt(v),font:{size:9}}}}}
    });
    const est = {};
    data.forEach(d=>{const e=d._estado||'Sin estado'; est[e]=(est[e]||0)+1;});
    const eK = Object.keys(est).filter(k=>est[k]>0);
    CHARTS.estado = new Chart($('ch-estado'), {
      type:'doughnut',
      data:{labels:eK, datasets:[{data:eK.map(k=>est[k]), backgroundColor:eK.map(k=>STATUS_COLORS[k]||'#A889AB'), borderWidth:0}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom',labels:{font:{size:9},boxWidth:8}}}}
    });
    const tip = {};
    data.forEach(d=>{const t=(d._tipo_proy||'').toLowerCase().includes('estrateg')?'Estratégico':'RyR'; tip[t]=(tip[t]||0)+1;});
    const tK = Object.keys(tip);
    CHARTS.tipo = new Chart($('ch-tipo'), {
      type:'doughnut',
      data:{labels:tK, datasets:[{data:tK.map(k=>tip[k]), backgroundColor:['#803F85','#1A9E6A'], borderWidth:0}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom',labels:{font:{size:9},boxWidth:8}}}}
    });
    const fases = {};
    data.forEach(d=>{const f=d._fase||'Sin fase'; fases[f]=(fases[f]||0)+1;});
    const fK = Object.keys(fases).sort();
    CHARTS.fase = new Chart($('ch-fase'), {
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
    CHARTS.avance = new Chart($('ch-avance'), {
      type:'bar',
      data:{labels:pmsAvance, datasets:[
        {label:'Real', data:realAvg, backgroundColor:'#803F85', borderRadius:4},
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
    const totalMonto = data.reduce((s,d)=>s+d._monto,0);
    const totalErog  = data.reduce((s,d)=>s+d._erogado,0);
    const totalDisp  = data.reduce((s,d)=>s+d._disponible,0);
    const activos    = data.filter(d=>d.tipo.toLowerCase().startsWith('proyecto'));
    const avg        = activos.length ? activos.reduce((s,d)=>s+d._pct,0)/activos.length : 0;
    const avgPlan    = activos.length ? activos.reduce((s,d)=>s+d._pct_plan,0)/activos.length : 0;
    const desv       = avg - avgPlan;
    const unidades   = [...new Set(data.map(d=>d._unidad).filter(Boolean))];
    $('main').innerHTML = `
      <div class="pm-header">
        <div class="pm-avatar">${ini}</div>
        <div>
          <div class="pm-name">${pm}</div>
          <div class="pm-meta">
            <span><i class="ti ti-briefcase" aria-hidden="true"></i> ${data.length} proyectos</span>
            <span><i class="ti ti-building-hospital" aria-hidden="true"></i> ${unidades.length} unidades</span>
            <span><i class="ti ti-trending-up" aria-hidden="true"></i> ${pct(avg)} avance</span>
            <span class="${desvClass(desv)}"><i class="ti ti-arrows-diff" aria-hidden="true"></i> Desviación: ${desvStr(desv)}</span>
          </div>
        </div>
        <div class="pm-header-kpis">
          <div class="pm-mini-kpi"><span class="pm-mini-label"><i class="ti ti-coin" aria-hidden="true"></i> Monto Aprobado TI</span><span class="pm-mini-val">${fmt(totalMonto)}</span></div>
          <div class="pm-mini-kpi"><span class="pm-mini-label"><i class="ti ti-receipt" aria-hidden="true"></i> Erogado</span><span class="pm-mini-val">${fmt(totalErog)}</span></div>
          <div class="pm-mini-kpi"><span class="pm-mini-label"><i class="ti ti-wallet" aria-hidden="true"></i> Disponible</span><span class="pm-mini-val">${fmt(totalDisp)}</span></div>
        </div>
        <button class="btn-primary" onclick="App.exportMSR()" style="margin-left:auto">
          <i class="ti ti-file-download" aria-hidden="true"></i> Exportar MSR_R
        </button>
      </div>
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-bar" aria-hidden="true"></i><span class="chart-title">Monto por Unidad</span></div>
          <div class="chart-wrap"><canvas id="ch-pm" role="img" aria-label="Monto por unidad"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-donut" aria-hidden="true"></i><span class="chart-title">Por Estado</span></div>
          <div class="chart-wrap"><canvas id="ch-estado" role="img" aria-label="Proyectos por estado"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-header"><i class="ti ti-chart-pie" aria-hidden="true"></i><span class="chart-title">Por Fase</span></div>
          <div class="chart-wrap"><canvas id="ch-tipo" role="img" aria-label="Proyectos por fase"></canvas></div>
        </div>
      </div>
      <div class="table-card">
        <div class="table-header">
          <div class="table-header-left"><i class="ti ti-table" aria-hidden="true"></i><span class="table-title">Proyectos de ${pm}</span></div>
          <div style="display:flex;align-items:center;gap:8px">
            <div class="search-box">
              <i class="ti ti-search" aria-hidden="true"></i>
              <input type="text" id="search-input" placeholder="Buscar proyecto..." oninput="App.onSearch(this.value)">
            </div>
            <span class="table-count" id="tbl-count"></span>
          </div>
        </div>
        <div class="table-scroll"><div id="tbl"></div></div>
      </div>`;
    FILTERED = data;
    App.renderChartsPM(data);
    App.renderTable(data);
  },

  renderChartsPM(data) {
    App.destroyCharts();
    const uds = [...new Set(data.map(d=>d._unidad).filter(Boolean))];
    const montos = uds.map(u=>data.filter(d=>d._unidad===u).reduce((s,d)=>s+d._monto,0));
    CHARTS.pm = new Chart($('ch-pm'), {
      type:'bar',
      data:{labels:uds, datasets:[{data:montos, backgroundColor:'#803F85', borderRadius:5}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}},
        scales:{x:{ticks:{font:{size:9},maxRotation:30}}, y:{ticks:{callback:v=>fmt(v),font:{size:9}}}}}
    });
    const est = {};
    data.forEach(d=>{const e=d._estado||'Sin estado'; est[e]=(est[e]||0)+1;});
    const eK = Object.keys(est).filter(k=>est[k]>0);
    CHARTS.estado = new Chart($('ch-estado'), {
      type:'doughnut',
      data:{labels:eK, datasets:[{data:eK.map(k=>est[k]), backgroundColor:eK.map(k=>STATUS_COLORS[k]||'#A889AB'), borderWidth:0}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom',labels:{font:{size:9},boxWidth:8}}}}
    });
    const fases = {};
    data.forEach(d=>{const f=d._fase||'Sin fase'; fases[f]=(fases[f]||0)+1;});
    const fK = Object.keys(fases).filter(k=>fases[k]>0);
    CHARTS.tipo = new Chart($('ch-tipo'), {
      type:'doughnut',
      data:{labels:fK, datasets:[{data:fK.map(k=>fases[k]), backgroundColor:PM_COLORS, borderWidth:0}]},
      options:{responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom',labels:{font:{size:9},boxWidth:8}}}}
    });
  },

  sortTable(col) {
    if (SORT_COL === col) SORT_DIR *= -1;
    else { SORT_COL = col; SORT_DIR = col === '_nombre' ? 1 : -1; }
    App.renderTable(FILTERED);
  },

  renderTable(data) {
    const search = SEARCH.toLowerCase();
    const visible = search
      ? data.filter(d => d._nombre.toLowerCase().includes(search) || d._unidad.toLowerCase().includes(search) || d.pm.toLowerCase().includes(search))
      : data;
    $('tbl-count').textContent = visible.length + ' registros';
    if (!visible.length) {
      $('tbl').innerHTML = '<div class="empty-state"><i class="ti ti-search empty-icon"></i><div class="empty-title">Sin resultados</div></div>';
      return;
    }
    const sorted = [...visible].sort((a,b) => {
      const va = a[SORT_COL] || '';
      const vb = b[SORT_COL] || '';
      if (typeof va === 'number') return (va - vb) * SORT_DIR;
      return String(va).localeCompare(String(vb)) * SORT_DIR;
    });
    const cols = [
      {label:'Nombre del Proyecto', key:'_nombre'},
      {label:'PM', key:'pm'},
      {label:'Unidad', key:'_unidad'},
      {label:'Fase', key:'_fase'},
      {label:'Estado', key:'_estado'},
      {label:'Tipo', key:'_tipo_proy'},
      {label:'Monto Aprobado TI', key:'_monto'},
      {label:'Erogado', key:'_erogado'},
      {label:'Disponible', key:'_disponible'},
      {label:'% Real', key:'_pct'},
      {label:'% Plan', key:'_pct_plan'},
      {label:'Desviación', key:'_desviacion'},
    ];
    const arrow = key => {
      if (SORT_COL !== key) return '<i class="ti ti-arrows-sort" style="opacity:.3;font-size:11px"></i>';
      return SORT_DIR === 1
        ? '<i class="ti ti-sort-ascending" style="font-size:11px;color:var(--p)"></i>'
        : '<i class="ti ti-sort-descending" style="font-size:11px;color:var(--p)"></i>';
    };
    $('tbl').innerHTML = `
      <table>
        <thead><tr>
          ${cols.map(c=>`<th onclick="App.sortTable('${c.key}')" style="cursor:pointer;user-select:none">${c.label} ${arrow(c.key)}</th>`).join('')}
        </tr></thead>
        <tbody>${sorted.map(d=>`
          <tr>
            <td title="${d._nombre}">${d._nombre||'—'}</td>
            <td><span class="chip">${d.pm}</span></td>
            <td title="${d._unidad}">${d._unidad||'—'}</td>
            <td>${d._fase||'—'}</td>
            <td><span class="badge ${badgeClass(d._estado)}">${d._estado||'—'}</span></td>
            <td>${d._tipo_proy||'—'}</td>
            <td>${fmt(d._monto)}</td>
            <td>${fmt(d._erogado)}</td>
            <td>${fmt(d._disponible)}</td>
            <td><div class="prog"><div class="prog-bar"><div class="prog-fill" style="width:${Math.round((d._pct||0)*100)}%"></div></div><span class="prog-val">${pct(d._pct)}</span></div></td>
            <td><span class="prog-val">${pct(d._pct_plan)}</span></td>
            <td><span class="${desvClass(d._desviacion)}">${desvStr(d._desviacion)}</span></td>
          </tr>`).join('')}
        </tbody>
      </table>`;
  },

  exportMSR() {
    if (!FILTERED.length) return;
    const wb = XLSX.utils.book_new();
    const PURPLE='FF7030A0', BLUE='FF2E75B6', ORANGE='FFE97132', GREEN='FF375623', DBLUE='FF1F3864';
    const h1 = bg => ({font:{name:'Aptos Display',bold:true,sz:11,color:{rgb:'FFFFFFFF'}},fill:{fgColor:{rgb:bg},patternType:'solid'},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}}});
    const h2 = bg => ({font:{name:'Aptos Display',bold:true,sz:11,color:{rgb:'FFFFFFFF'}},fill:{fgColor:{rgb:bg},patternType:'solid'},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{top:{style:'thin'},bottom:{style:'thin'},left:{style:'thin'},right:{style:'thin'}}});
    const ds = {font:{name:'Aptos Display',sz:10},alignment:{horizontal:'left',vertical:'center'},border:{bottom:{style:'thin',color:{rgb:'FFD9D9D9'}},right:{style:'thin',color:{rgb:'FFD9D9D9'}}}};
    const row1 = Array(48).fill({v:'',s:{}});
    row1[0]={v:'Información General | Nivel Proyecto',s:h1(PURPLE)};
    row1[10]={v:'Información General | Nivel Tarea',s:h1(PURPLE)};
    row1[20]={v:'Información de Compra',s:h1(BLUE)};
    row1[29]={v:'Planeación & Seguimiento',s:h1(ORANGE)};
    row1[36]={v:'Plan de Erogaciones',s:h1(GREEN)};
    row1[43]={v:'Presupuesto',s:h1(DBLUE)};
    const cols=[['Gerencia TI',PURPLE],['Año de gestión',PURPLE],['ID Proyecto',PURPLE],['Nombre de proyecto',PURPLE],['Fecha Inicio',PURPLE],['Fecha Cierre',PURPLE],['Área Solicitante',PURPLE],['Unidad Médica',PURPLE],['Fase del proyecto',PURPLE],['% Avance Real',PURPLE],['RPP',PURPLE],['Estatus',PURPLE],['Responsable TI',PURPLE],['ID Tarea',PURPLE],['Nombre de tarea',PURPLE],['Categoría',PURPLE],['Concepto',PURPLE],['Descripción',PURPLE],['Departamento/Área',PURPLE],['Tipo',PURPLE],['Proveedor',BLUE],['#Solicitud/#Requisición',BLUE],['Fecha de solicitud',BLUE],['Estatus de solicitud',BLUE],['OC Asociada',BLUE],['Monto OC',BLUE],['Fecha generación de OC',BLUE],['Estatus de OC',BLUE],['Factura',BLUE],['Fecha KickOff',ORANGE],['Preconfiguracion de Equipos',ORANGE],['Entrega de equipos en Unidad',ORANGE],['Instalación Física de Equipos',ORANGE],['Configuración de Equipos',ORANGE],['Pruebas',ORANGE],['Comentario de Estatus',ORANGE],['Fecha Go Live',ORANGE],['Mes 1',GREEN],['Mes 2',GREEN],['Mes 3',GREEN],['Mes 4',GREEN],['Mes 5',GREEN],['Mes 6',GREEN],['Mes "N"',GREEN],['Monto Aprobado',DBLUE],['Total Erogado',DBLUE],['Total Comprometido',DBLUE],['Remanente',DBLUE]];
    const row2 = cols.map(([v,bg])=>({v,s:h2(bg)}));
    const rows = [row1,row2];
    FILTERED.forEach(d=>{
      const c = v=>({v:v||'',s:ds});
      const r = Array(48).fill(c(''));
      r[0]=c(d.pm); r[1]=c(2026); r[2]=c(d._id); r[3]=c(d._nombre);
      r[4]=c(d._fecha_ini); r[5]=c(d._fecha_cie); r[6]=c(d._direccion); r[7]=c(d._unidad);
      r[8]=c(d._fase); r[9]=c(d._pct?Math.round(d._pct*100)+'%':'');
      r[10]=c(d._rrpp); r[11]=c(d._estado); r[12]=c(d._lider); r[19]=c(d._tipo_proy);
      r[29]=c(d._fecha_golive); r[35]=c(d._comentarios);
      r[43]=c(d._monto||0); r[44]=c(d._erogado||0); r[45]=c(d._comprometido||0); r[46]=c(d._disponible||0);
      rows.push(r);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:9}},{s:{r:0,c:10},e:{r:0,c:19}},{s:{r:0,c:20},e:{r:0,c:28}},{s:{r:0,c:29},e:{r:0,c:35}},{s:{r:0,c:36},e:{r:0,c:42}},{s:{r:0,c:43},e:{r:0,c:46}}];
    ws['!cols']=[{wch:20},{wch:14},{wch:12},{wch:41},{wch:16},{wch:15},{wch:12},{wch:21},{wch:24},{wch:14},{wch:13},{wch:20},{wch:23},{wch:14},{wch:25},{wch:19},{wch:19},{wch:21},{wch:25},{wch:20},{wch:17},{wch:25},{wch:27},{wch:26},{wch:19},{wch:17},{wch:19},{wch:26},{wch:19},{wch:26},{wch:34},{wch:35},{wch:45},{wch:40},{wch:23},{wch:29},{wch:27},{wch:13},{wch:13},{wch:13},{wch:13},{wch:13},{wch:13},{wch:17},{wch:26},{wch:21},{wch:27},{wch:19}];
    ws['!rows']=[{hpt:33.6},{hpt:32.1}];
    XLSX.utils.book_append_sheet(wb,ws,'Tracker');
    XLSX.writeFile(wb,`MSR_R_Portafolio_${new Date().toISOString().slice(0,10)}.xlsx`,{cellStyles:true});
  },
};
