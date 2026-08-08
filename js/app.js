/* =====================================================================
   Mapa Zelar — aplicacao principal
   Carga da planilha (GitHub ao vivo ou snapshot embutido), filtros,
   mapa, tabela, estatisticas, tema claro/escuro.
   ===================================================================== */
(function () {
  'use strict';

  /* ---------- estado global ---------- */
  var ALL = [];            /* unidades enriquecidas */
  var FILTERED = [];
  var state = {
    q: '', cliente: {}, unidade: {}, segmento: {}, estado: {}, consultor: {},
    status: {}, viSemData: true, viMais: '', viMenos: '', enMenos: '', enDe: '', enAte: '',
    atMin: '', atMax: '', atExact: false, atExactVal: '',
    kmMin: '', kmMax: ''
  };
  var order = { k: 'cl', dir: 1 };
  var map = null, markers = null, tileLight = null, tileDark = null;
  var FMT = new Intl.NumberFormat('pt-BR');

  /* ---------- datas ---------- */
  var MESES = { JANEIRO: 0, FEVEREIRO: 1, MARCO: 2, MARÇO: 2, ABRIL: 3, MAIO: 4, JUNHO: 5, JULHO: 6, AGOSTO: 7, SETEMBRO: 8, OUTUBRO: 9, NOVEMBRO: 10, DEZEMBRO: 11 };
  var MESES_ABR = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

  function parseMes(v) {
    var m = String(v || '').toUpperCase().match(/^([A-ZÇ]+)\s*\/\s*(\d{2})$/);
    if (!m || MESES[m[1]] === undefined) return null;
    return new Date(2000 + (+m[2]), MESES[m[1]], 1);
  }
  function fmtMes(d) { return d ? MESES_ABR[d.getMonth()] + '/' + String(d.getFullYear()).slice(2) : ''; }
  function monthDiff(a, b) { return (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth()); }

  /* ---------- status ---------- */
  var STATUS_META = {
    novo:    { label: 'Ativos novos',       color: '#7c3aed', cls: 'b-novo' },
    nunca:   { label: 'Nunca visitado',     color: '#ef4444', cls: 'b-nunca' },
    velho:   { label: '12+ meses sem visita', color: '#f97316', cls: 'b-velho' },
    medio:   { label: '3–12 meses',         color: '#d97706', cls: 'b-medio' },
    recente: { label: 'Menos de 3 meses',   color: '#16a34a', cls: 'b-recente' },
    semdata: { label: 'Sem data',           color: '#94a3b8', cls: 'b-semdata' }
  };
  var STATUS_KEYS = ['novo', 'nunca', 'velho', 'medio', 'recente', 'semdata'];

  function statusOf(u, now) {
    var v = String(u.vi).toUpperCase();
    if (v === 'ATIVOS NOVOS') return 'novo';
    if (v === 'NUNCA VISITADO') return 'nunca';
    var d = parseMes(u.vi);
    if (!d) return 'semdata';
    var m = monthDiff(d, now);
    return m > 12 ? 'velho' : (m >= 3 ? 'medio' : 'recente');
  }

  /* ---------- enriquecimento ---------- */
  function enrich(u) {
    var now = new Date();
    u._viDate = parseMes(u.vi);
    u._enDate = parseMes(u.en);
    u._status = statusOf(u, now);
    u._at = parseInt(u.at, 10) || 0;
    u._km = parseFloat(u.km) || 0;
    return u;
  }

  /* ---------- carga de dados ---------- */
  function fetchXlsx(url, ms) {
    return new Promise(function (resolve, reject) {
      var ctl = new AbortController();
      var t = setTimeout(function () { ctl.abort(); }, ms || 20000);
      fetch(url, { signal: ctl.signal })
        .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.arrayBuffer(); })
        .then(function (buf) { clearTimeout(t); resolve(buf); })
    });
  }

  function parseXlsx(buf) {
    var wb = XLSX.read(buf, { type: 'array' });
    var sheet = wb.Sheets['Mapa'] || wb.Sheets[wb.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    var out = [];
    rows.forEach(function (r) {
      var geo = GEO.geocodeRow(r['CIDADE'], r['ESTADO'], r['LOCALIZAÇÃO']);
      if (!geo) return;
      out.push(enrich({
        cl: String(r['CLIENTE'] || ''), un: String(r['UNIDADE'] || ''),
        en: String(r['ENTREGA'] || ''), vi: String(r['ULTIMA VISITA'] || ''),
        sg: String(r['SEGMENTO'] || ''), cn: String(r['CNPJ'] || ''),
        ci: String(r['CIDADE'] || ''), es: String(r['ESTADO'] || ''),
        at: r['ATIVOS'] || 0, km: r['KM'] || '', tm: String(r['TEMPO'] || ''),
        f: String(r['TELEFONE'] || ''), n: String(r['NOME'] || ''),
        e: String(r['E-MAIL'] || ''), ca: String(r['CARGO'] || ''),
        co: String(r['CONSULTOR'] || ''), lo: String(r['LOCALIZAÇÃO'] || ''),
        lat: geo.lat, lng: geo.lng, pr: geo.prec, ap: geo.aprox
      }));
    });
    return out;
  }

  function loadData() {
    var setSource = function (mode, extra) {
      var el = document.getElementById('dataSource');
      el.classList.remove('live', 'snap');
      el.classList.add(mode);
      document.getElementById('dataSourceText').textContent = extra;
      document.getElementById('footSource').textContent = extra;
    };
    var isDefault = CONFIG.GITHUB_USER.indexOf('SEU-') === 0;
    var attempts = isDefault ? [] : CONFIG_URLS;
    if (!attempts.length) {
      ALL = SNAPSHOT.map(enrich);
      setSource('snap', 'Dados embutidos (planilha não carregada)');
      showBanner('Site rodando com dados embutidos. Para carregar a planilha do GitHub ao vivo, configure seu usuário e repositório em <b>js/config.js</b>.');
      buildUI();
      return;
    }
    var chain = Promise.reject();
    var done = false;
    attempts.forEach(function (url) {
      chain = chain.catch(function () {
        return fetchXlsx(url).then(function (buf) {
          if (done) return;
          done = true;
          ALL = parseXlsx(buf);
          if (!ALL.length) throw new Error('planilha vazia');
          setSource('live', 'XLSX ao vivo · ' + new Date().toLocaleString('pt-BR'));
        });
      });
    });
    chain.catch(function () {
      ALL = SNAPSHOT.map(enrich);
      setSource('snap', 'Dados embutidos (planilha não carregada)');
      showBanner('Não foi possível baixar a planilha do GitHub. Exibindo dados embutidos.');
    }).then(function () {
      buildUI();
    });
  }

  function showBanner(html) {
    var b = document.getElementById('banner');
    b.innerHTML = html;
    b.classList.remove('hidden');
  }

  /* ---------- filtros ---------- */
  function clean(s) { return (s || '').toLowerCase().trim(); }

  function applyFilters() {
    var now = new Date();
    var q = clean(state.q);
    var st = state;
    FILTERED = ALL.filter(function (u) {
      if (q) {
        var hay = clean([u.cl, u.un, u.ci, u.es, u.sg, u.co].join(' '));
        if (hay.indexOf(q) === -1) return false;
      }
      if (Object.keys(st.cliente).length && !st.cliente[u.cl]) return false;
      if (Object.keys(st.unidade).length && !st.unidade[u.un]) return false;
      if (Object.keys(st.segmento).length && !st.segmento[u.sg]) return false;
      if (Object.keys(st.estado).length && !st.estado[u.es]) return false;
      if (Object.keys(st.consultor).length && !st.consultor[u.co]) return false;
      if (Object.keys(st.status).length && !st.status[u._status]) return false;

      if (st.atExact && st.atExactVal !== '' && u._at !== +st.atExactVal) return false;
      if (!st.atExact) {
        if (st.atMin !== '' && u._at < +st.atMin) return false;
        if (st.atMax !== '' && u._at > +st.atMax) return false;
      }
      if (st.kmMin !== '' && u._km < +st.kmMin) return false;
      if (st.kmMax !== '' && u._km > +st.kmMax) return false;

      if (u._status === 'semdata' && !st.viSemData) return false;

      if (st.viMais !== '' && (u._status === 'nunca' || !u._viDate)) return false;
      if (st.viMais !== '' && u._viDate) {
        var m = monthDiff(u._viDate, now);
        if (!(m > +st.viMais)) return false;
      }
      if (st.viMenos !== '' && u._viDate) {
        if (monthDiff(u._viDate, now) > +st.viMenos) return false;
      }
      if (st.enMenos !== '' && u._enDate) {
        if (monthDiff(u._enDate, now) > +st.enMenos) return false;
      }
      if (st.enDe && u._enDate && u._enDate.getFullYear() < +st.enDe) return false;
      if (st.enAte && u._enDate && u._enDate.getFullYear() > +st.enAte) return false;
      return true;
    });
    render();
  }

  function clearFilters() {
    state = {
      q: '', cliente: {}, unidade: {}, segmento: {}, estado: {}, consultor: {},
      status: {}, viSemData: true, viMais: '', viMenos: '', enMenos: '', enDe: '', enAte: '',
      atMin: '', atMax: '', atExact: false, atExactVal: '',
      kmMin: '', kmMax: ''
    };
    document.getElementById('q').value = '';
    document.getElementById('atMin').value = '';
    document.getElementById('atMax').value = '';
    document.getElementById('atExact').checked = false;
    document.getElementById('atExactVal').value = '';
    document.getElementById('kmMin').value = '';
    document.getElementById('kmMax').value = '';
    document.getElementById('viSemData').checked = true;
    document.getElementById('viMais').value = '';
    document.getElementById('viMenos').value = '';
    document.getElementById('enMenos').value = '';
    document.getElementById('enDe').value = '';
    document.getElementById('enAte').value = '';
    document.getElementById('exactRow').classList.add('hide');
    document.querySelectorAll('.multi').forEach(function (m) { m.classList.remove('open'); m.querySelector('.multi-pop').classList.add('hidden'); });
    renderMultiAll();
    renderChipsAll();
    applyFilters();
  }

  /* ---------- componentes: multi-select ---------- */
  function buildCounts(field) {
    var c = {};
    ALL.forEach(function (u) { var v = u[field] || '—'; c[v] = (c[v] || 0) + 1; });
    return c;
  }
  function sortedKeys(counts) {
    return Object.keys(counts).sort(function (a, b) { return counts[b] - counts[a] || a.localeCompare(b, 'pt'); });
  }

  function renderMultiAll() {
    document.querySelectorAll('.multi').forEach(function (el) {
      var name = el.getAttribute('data-multi');
      var sel = Object.keys(state[name]);
      var label = el.querySelector('.multi-label');
      var count = el.querySelector('.multi-count');
      var total = ALL.length;
      if (!sel.length) { label.textContent = name === 'cliente' ? 'Todos' : name === 'unidade' ? 'Todas' : 'Todos'; count.textContent = total; }
      else { label.textContent = sel.length + ' selecionado' + (sel.length > 1 ? 's' : ''); count.textContent = ''; }
    });
  }

  function initMulti() {
    var mapCfg = { cliente: { f: 'cl', label: 'Cliente' }, unidade: { f: 'un', label: 'Unidade' }, consultor: { f: 'co', label: 'Consultor' } };
    document.querySelectorAll('.multi').forEach(function (el) {
      var name = el.getAttribute('data-multi');
      var cfg = mapCfg[name];
      var btn = el.querySelector('.multi-btn');
      var pop = el.querySelector('.multi-pop');
      var search = el.querySelector('.multi-search');
      var opts = el.querySelector('.multi-options');
      btn.addEventListener('click', function (e) {
        e.stopPropagation();
        var open = el.classList.toggle('open');
        pop.classList.toggle('hidden', !open);
        if (open) { renderOpts(''); search.focus(); }
      });
      search.addEventListener('input', function () { renderOpts(clean(search.value)); });
      el.querySelector('[data-multi-all]').addEventListener('click', function () {
        state[name] = {}; buildCounts(cfg.f) && sortedKeys(buildCounts(cfg.f)).forEach(function (k) { state[name][k] = 1; });
        renderOpts(clean(search.value)); renderMultiAll(); applyFilters();
      });
      el.querySelector('[data-multi-none]').addEventListener('click', function () {
        state[name] = {}; renderOpts(clean(search.value)); renderMultiAll(); applyFilters();
      });
      function renderOpts(q) {
        var counts = buildCounts(cfg.f);
        var keys = sortedKeys(counts).filter(function (k) { return !q || clean(k).indexOf(q) > -1; });
        opts.innerHTML = '';
        keys.forEach(function (k) {
          var row = document.createElement('label');
          row.className = 'opt';
          var cb = document.createElement('input');
          cb.type = 'checkbox';
          cb.checked = !!state[name][k];
          cb.addEventListener('change', function () {
            if (cb.checked) state[name][k] = 1; else delete state[name][k];
            renderMultiAll(); applyFilters();
          });
          var span = document.createElement('span');
          span.textContent = k;
          if (state[name][k]) span.className = 'on';
          var n = document.createElement('span');
          n.className = 'n';
          n.textContent = FMT.format(counts[k]);
          row.appendChild(cb); row.appendChild(span); row.appendChild(n);
          opts.appendChild(row);
        });
        if (!keys.length) { opts.innerHTML = '<div class="opt" style="color:var(--muted)">Sem resultados</div>'; }
      }
    });
    document.addEventListener('click', function (e) {
      document.querySelectorAll('.multi').forEach(function (m) {
        if (!m.contains(e.target)) { m.classList.remove('open'); m.querySelector('.multi-pop').classList.add('hidden'); }
      });
    });
  }

  /* ---------- componentes: chips ---------- */
  function renderChipsAll() {
    renderChips('segmento', function (u) { return u.sg; });
    renderChips('estado', function (u) { return u.es; });
    renderChips('status', function (u) { return u._status; }, true);
  }
  function renderChips(name, fn, meta) {
    var el = document.querySelector('[data-chips="' + name + '"]');
    var counts = {};
    ALL.forEach(function (u) { var v = fn(u); counts[v] = (counts[v] || 0) + 1; });
    var keys = meta ? STATUS_KEYS.filter(function (k) { return counts[k]; }) : sortedKeys(counts);
    el.innerHTML = '';
    keys.forEach(function (k) {
      var chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'chip' + (state[name][k] ? ' on' : '');
      if (meta) { var sw = document.createElement('span'); sw.className = 'swatch'; sw.style.background = STATUS_META[k].color; chip.appendChild(sw); }
      var label = document.createElement('span');
      label.textContent = meta ? STATUS_META[k].label : k;
      var n = document.createElement('span');
      n.className = 'n';
      n.textContent = FMT.format(counts[k]);
      chip.appendChild(label); chip.appendChild(n);
      chip.addEventListener('click', function () {
        if (state[name][k]) delete state[name][k]; else state[name][k] = 1;
        renderChipsAll();
        applyFilters();
      });
      el.appendChild(chip);
    });
  }

  /* ---------- inputs ---------- */
  function bindInputs() {
    var q = document.getElementById('q');
    var t = null;
    q.addEventListener('input', function () { clearTimeout(t); t = setTimeout(function () { state.q = q.value; applyFilters(); }, 220); });

    var num = [
      ['atMin', function (v) { state.atMin = v; }],
      ['atMax', function (v) { state.atMax = v; }],
      ['atExactVal', function (v) { state.atExactVal = v; }],
      ['kmMin', function (v) { state.kmMin = v; }],
      ['kmMax', function (v) { state.kmMax = v; }]
    ];
    num.forEach(function (p) {
      document.getElementById(p[0]).addEventListener('input', function (e) { p[1](e.target.value); applyFilters(); });
    });
    document.getElementById('atExact').addEventListener('change', function (e) {
      state.atExact = e.target.checked;
      document.getElementById('exactRow').classList.toggle('hide', !e.target.checked);
      applyFilters();
    });
    document.getElementById('viSemData').addEventListener('change', function (e) { state.viSemData = e.target.checked; applyFilters(); });
    var sel = [
      ['viMais', function (v) { state.viMais = v; }],
      ['viMenos', function (v) { state.viMenos = v; }],
      ['enMenos', function (v) { state.enMenos = v; }],
      ['enDe', function (v) { state.enDe = v; }],
      ['enAte', function (v) { state.enAte = v; }]
    ];
    sel.forEach(function (p) {
      document.getElementById(p[0]).addEventListener('change', function (e) { p[1](e.target.value); applyFilters(); });
    });
    ['clearFilters', 'clearFilters2'].forEach(function (id) {
      document.getElementById(id).addEventListener('click', clearFilters);
    });
    document.querySelectorAll('.preset').forEach(function (b) {
      b.addEventListener('click', function () {
        clearFilters();
        var p = b.getAttribute('data-preset');
        var now = new Date();
        if (p === 'nunca') state.status = { nunca: 1 };
        else if (p === 'velho') state.status = { velho: 1, nunca: 1 };
        else if (p === 'novo') state.status = { novo: 1 };
        else if (p === 'recente') state.status = { recente: 1 };
        if (p === 'recente' || p === 'velho') {
          state.viSemData = true;
          document.getElementById('viSemData').checked = true;
        }
        void now;
        renderChipsAll();
        applyFilters();
      });
    });
  }

  function fillEntregaSelects() {
    var years = {};
    ALL.forEach(function (u) { if (u._enDate) years[u._enDate.getFullYear()] = 1; });
    var opts = Object.keys(years).sort();
    opts.unshift('');
    ['enDe', 'enAte'].forEach(function (id) {
      var el = document.getElementById(id);
      el.innerHTML = opts.map(function (y) { return '<option value="' + y + '">' + (y ? y : '—') + '</option>'; }).join('');
    });
  }

  /* ---------- stats ---------- */
  function renderStats() {
    var el = document.getElementById('stats');
    var total = FILTERED.length;
    var ativos = FILTERED.reduce(function (s, u) { return s + u._at; }, 0);
    var c = {};
    FILTERED.forEach(function (u) { c[u._status] = (c[u._status] || 0) + 1; });
    var cards = [
      { cls: 'blue', v: FMT.format(total), l: 'Unidades no filtro' },
      { cls: 'blue', v: FMT.format(ativos), l: 'Ativos instalados' },
      { cls: 'red', v: FMT.format(c.nunca || 0), l: 'Nunca visitados' },
      { cls: 'orange', v: FMT.format((c.velho || 0) + (c.nunca || 0)), l: '12+ meses sem visita' },
      { cls: 'violet', v: FMT.format(c.novo || 0), l: 'Ativos novos' },
      { cls: 'green', v: FMT.format(c.recente || 0), l: 'Visitados < 3 meses' }
    ];
    el.innerHTML = cards.map(function (c) {
      return '<div class="stat ' + c.cls + '"><div class="v">' + c.v + '</div><div class="l">' + c.l + '</div></div>';
    }).join('');
  }

  /* ---------- mapa ---------- */
  function initMap() {
    var theme = document.documentElement.getAttribute('data-theme');
    var light = theme !== 'dark';
    tileLight = L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', {
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> &copy; <a href="https://carto.com/">CARTO</a>',
      subdomains: 'abcd', maxZoom: 19
    });
    tileDark = L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png', {
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> &copy; <a href="https://carto.com/">CARTO</a>',
      subdomains: 'abcd', maxZoom: 19
    });
    map = L.map('map', {
      center: [-19.47, -47.71], zoom: 5,
      layers: [light ? tileLight : tileDark],
      zoomControl: true
    });
    markers = L.markerClusterGroup({
      maxClusterRadius: 46,
      iconCreateFunction: function (cluster) {
        var n = cluster.getChildCount();
        var s = n > 200 ? 54 : n > 80 ? 48 : 42;
        return L.divIcon({ html: '<div class="cluster" style="width:' + s + 'px;height:' + s + 'px;font-size:' + (n > 99 ? 12 : 14) + 'px">' + n + '</div>', className: '', iconSize: [s, s] });
      }
    });
    map.addLayer(markers);
    var legend = document.getElementById('legend');
    legend.innerHTML = '<div class="title">Situação da visita</div>' + STATUS_KEYS.map(function (k) {
      return '<div><i style="background:' + STATUS_META[k].color + '"></i>' + STATUS_META[k].label + '</div>';
    }).join('');
    window.addEventListener('resize', function () { map.invalidateSize(); });
  }

  function renderMarkers() {
    markers.clearLayers();
    var icons = {};
    FILTERED.forEach(function (u) {
      if (!u.lat || !u.lng) return;
      var st = u._status;
      if (!icons[st]) {
        icons[st] = L.divIcon({
          html: '<div class="pin" style="--c:' + STATUS_META[st].color + '"></div>',
          className: '', iconSize: [18, 18], iconAnchor: [9, 9], popupAnchor: [0, -10]
        });
      }
      var mk = L.marker([u.lat, u.lng], { icon: icons[st], title: u.un });
      mk.bindPopup(popupHtml(u));
      markers.addLayer(mk);
    });
    if (!FILTERED.length) return;
    if (FILTERED.length === ALL.length) { map.setView([-19.47, -47.71], 5); }
    else { map.fitBounds(markers.getBounds(), { padding: [26, 26], maxZoom: 10 }); }
  }

  function popupHtml(u) {
    var st = STATUS_META[u._status];
    return '<div class="pop">' +
      '<h3>' + esc(u.un || u.cl) + '</h3>' +
      '<div class="sub">' + esc(u.cl) + '</div>' +
      '<div class="badges"><span class="badge ' + st.cls + '">' + st.label + '</span>' +
      (u.sg ? '<span class="badge" style="background:var(--brand-soft);color:var(--brand2)">' + esc(u.sg) + '</span>' : '') +
      (u.ap ? '<span class="badge b-semdata">posição aproximada</span>' : '') + '</div>' +
      prow('Cidade', esc(u.ci) + (u.es ? ' - ' + esc(u.es) : '')) +
      prow('Ativos', '<b>' + FMT.format(u._at) + '</b>') +
      prow('Última visita', u.vi ? esc(u.vi) : '—') +
      prow('Entrega', u.en ? esc(u.en) : '—') +
      (u.co && u.co !== '-' ? prow('Consultor', esc(u.co)) : '') +
      (u.f && u.f !== '-' ? prow('Telefone', esc(u.f)) : '') +
      (u.e ? prow('E-mail', esc(u.e)) : '') +
      (u.km ? prow('KM', FMT.format(u._km) + (u.tm ? ' · ' + esc(u.tm) : '')) : '') +
      (u.lo ? prow('Local', esc(u.lo)) : '') +
      '<a class="gmaps" target="_blank" rel="noopener" href="https://www.google.com/maps/search/?api=1&query=' + u.lat + ',' + u.lng + '">Abrir no Google Maps ↗</a>' +
      '</div>';
  }
  function prow(k, v) { return '<div class="prow"><span style="color:var(--muted)">' + k + '</span><span style="text-align:right">' + v + '</span></div>'; }
  function esc(s) { return String(s || '').replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

  /* ---------- tabela ---------- */
  function renderTable() {
    var body = document.getElementById('gridBody');
    document.getElementById('countBadge').textContent = FMT.format(FILTERED.length);
    document.getElementById('tableInfo').textContent = FILTERED.length + ' unidade' + (FILTERED.length === 1 ? '' : 's');
    var rows = FILTERED.slice().sort(function (a, b) {
      var va = a[order.k], vb = b[order.k];
      if (order.k === 'at') { va = a._at; vb = b._at; }
      else if (order.k === 'km') { va = a._km; vb = b._km; }
      else if (order.k === 'vi') { va = a._viDate ? a._viDate.getTime() : (a.vi === 'NUNCA VISITADO' ? -1 : -2); vb = b._viDate ? b._viDate.getTime() : (b.vi === 'NUNCA VISITADO' ? -1 : -2); }
      if (va < vb) return -1 * order.dir;
      if (va > vb) return 1 * order.dir;
      return 0;
    });
    body.innerHTML = rows.map(function (u) {
      var st = STATUS_META[u._status];
      return '<tr>' +
        '<td>' + esc(u.cl) + '</td>' +
        '<td><b>' + esc(u.un) + '</b></td>' +
        '<td>' + esc(u.ci) + (u.es ? '/' + esc(u.es) : '') + '</td>' +
        '<td>' + esc(u.sg || '—') + '</td>' +
        '<td class="num"><b>' + FMT.format(u._at) + '</b></td>' +
        '<td><span class="st" style="background:' + st.color + '1f;color:' + st.color + '">' + st.label + '</span> ' + esc(u.vi) + '</td>' +
        '<td>' + esc(u.en || '—') + '</td>' +
        '<td>' + esc(u.co && u.co !== '-' ? u.co : '—') + '</td>' +
        '<td class="num">' + (u._km ? FMT.format(u._km) : '—') + '</td>' +
        '</tr>';
    }).join('');
  }

  function exportCsv() {
    var cols = ['Cliente', 'Unidade', 'Entrega', 'Ultima visita', 'Situacao', 'Segmento', 'CNPJ', 'Cidade', 'Estado', 'Ativos', 'KM', 'Tempo', 'Telefone', 'Nome', 'E-mail', 'Cargo', 'Consultor', 'Localizacao', 'Latitude', 'Longitude'];
    var lines = [cols.join(';')];
    FILTERED.forEach(function (u) {
      lines.push([u.cl, u.un, u.en, u.vi, STATUS_META[u._status].label, u.sg, u.cn, u.ci, u.es, u._at, u.km, u.tm, u.f, u.n, u.e, u.ca, u.co, u.lo, u.lat, u.lng].join(';'));
    });
    var blob = new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'mapa-zelar-filtrado.csv';
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function initTable() {
    document.querySelectorAll('.grid th').forEach(function (th) {
      th.addEventListener('click', function () {
        var k = th.getAttribute('data-k');
        if (order.k === k) order.dir *= -1; else { order.k = k; order.dir = 1; }
        renderTable();
      });
    });
    document.getElementById('exportCsv').addEventListener('click', exportCsv);
  }

  /* ---------- tabs ---------- */
  function initTabs() {
    document.querySelectorAll('.tab').forEach(function (t) {
      t.addEventListener('click', function () {
        document.querySelectorAll('.tab').forEach(function (x) { x.classList.remove('active'); });
        document.querySelectorAll('.view').forEach(function (v) { v.classList.remove('active'); });
        t.classList.add('active');
        document.getElementById('view-' + t.getAttribute('data-view')).classList.add('active');
        if (t.getAttribute('data-view') === 'mapa') setTimeout(function () { map.invalidateSize(); }, 60);
      });
    });
  }

  /* ---------- tema ---------- */
  function applyTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    try { localStorage.setItem('zelar-theme', theme); } catch (e) {}
    if (map) {
      if (theme === 'dark') { if (map.hasLayer(tileLight)) map.removeLayer(tileLight); if (!map.hasLayer(tileDark)) map.addLayer(tileDark); }
      else { if (map.hasLayer(tileDark)) map.removeLayer(tileDark); if (!map.hasLayer(tileLight)) map.addLayer(tileLight); }
    }
  }
  function initTheme() {
    var saved = null; try { saved = localStorage.getItem('zelar-theme'); } catch (e) {}
    var pref = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    applyTheme(saved || pref);
    document.getElementById('themeToggle').addEventListener('click', function () {
      applyTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
    });
  }

  /* ---------- render total ---------- */
  function render() {
    renderStats();
    renderMarkers();
    renderTable();
  }

  function buildUI() {
    initTheme();
    initMap();
    initMulti();
    renderChipsAll();
    bindInputs();
    fillEntregaSelects();
    initTable();
    initTabs();
    applyFilters();
    renderMultiAll();
  }

  loadData();
})();
