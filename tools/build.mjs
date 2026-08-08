/* =====================================================================
   Regenera js/data.js (copia embutida dos dados) a partir da planilha.
   Uso:  node tools/build.mjs
   Requer: npm install --no-save xlsx open-location-code
   Mantem a MESMA logica de geocodificacao do js/geo.js (fallback para
   o centroide do municipio e do estado).
   ===================================================================== */
import fs from 'node:fs';
import path from 'node:path';
import OpenLocationCode from 'open-location-code';
import XLSX from 'xlsx';

const OLC = new (OpenLocationCode.OpenLocationCode || OpenLocationCode.default || OpenLocationCode)();

const root = path.resolve(import.meta.dirname, '..');
const xlsxPath = path.join(root, 'Mapa - Zelar.xlsx');
const outPath = path.join(root, 'js', 'data.js');

const wb = XLSX.readFile(xlsxPath);
const sheet = wb.Sheets['Mapa'] || wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
console.log(`Planilha: ${rows.length} linhas`);

const SEP = '+';
const SEP_POS = 8;
const UFS = 'RO|AC|AM|RR|PA|AP|TO|MA|PI|CE|RN|PB|PE|AL|SE|BA|MG|ES|RJ|SP|PR|SC|RS|MS|MT|GO|DF';
const ESTADOS = {
  RO: [-10.83, -63.34], AC: [-8.77, -70.55], AM: [-3.47, -65.1], RR: [1.99, -61.33], PA: [-4.47, -53.9],
  AP: [1.42, -51.77], TO: [-9.46, -48.2], MA: [-5.45, -45.24], PI: [-7.72, -42.72], CE: [-5.35, -39.66],
  RN: [-5.81, -36.57], PB: [-7.24, -36.76], PE: [-8.41, -37.84], AL: [-9.62, -36.77], SE: [-10.57, -37.46],
  BA: [-12.94, -41.61], MG: [-18.51, -44.56], ES: [-19.18, -40.33], RJ: [-22.25, -42.68], SP: [-22.14, -48.79],
  PR: [-24.89, -51.55], SC: [-27.24, -50.22], RS: [-30.56, -53.51], MS: [-20.44, -54.65], MT: [-12.68, -56.09],
  GO: [-15.83, -49.25], DF: [-15.78, -47.93]
};

const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const MUNI = (() => {
  try {
    const m = JSON.parse(fs.readFileSync(path.join(root, 'muni.json'), 'utf8'));
    const out = {};
    m.forEach(r => { out[norm(r.nome) + '|' + String(r.uf).toUpperCase()] = [r.lat, r.lng]; });
    return out;
  } catch (e) { return null; }
})();
if (!MUNI) console.log('Aviso: muni.json ausente — sem fallback de municipio.');

function resolveMuni(city, uf) {
  if (!city) return null;
  const key = norm(city) + '|' + String(uf || '').toUpperCase();
  if (MUNI && MUNI[key]) return MUNI[key];
  return ESTADOS[String(uf || '').toUpperCase()] || null;
}

function extractCode(loc) {
  const s = String(loc || '');
  let m = s.match(/([2-9CFGHJMPQRVWX]{4,8})\s*\+([2-9CFGHJMPQRVWX]{2,3})/);
  if (m) return m[1] + m[2];
  m = s.match(/([2-9CFGHJMPQRVWX]{8})\b/);
  return m ? m[1] : null;
}

function textCityRef(loc, uf) {
  const s = String(loc || '');
  const re = new RegExp('([^,;\\-]{3,60}?)\\s*[-,]\\s*(' + UFS + ')\\b', 'g');
  const ms = s.match(re);
  if (!ms || !ms.length) return null;
  const mm = ms[ms.length - 1].match(/([^,;-]{2,60}?)\s*[-,]\s*(' + UFS + ')\b/);
  if (!mm) return null;
  const city = mm[1].trim().replace(/^[\s-]+/, '');
  if (city.length < 3 || /^(jardim|distrito|centro|vila|fazenda|parque|rodovia|avenida|rua|br|go|mt|sp|mg|rj|pr|sc|rs|ba|ma|pa)/i.test(city)) return null;
  return resolveMuni(city, mm[2]);
}

function normalizeDigits(d) {
  if (d.length === 10) return d.slice(0, 8) + SEP + d.slice(8);
  if (d.length === 8) return d.slice(0, 6) + SEP + d.slice(6);
  if (d.length === 6) return d.slice(0, 4) + SEP + d.slice(4);
  return null;
}

function haversineKm(a, b) {
  const R = 6371, dLat = (b[0] - a[0]) * Math.PI / 180, dLng = (b[1] - a[1]) * Math.PI / 180;
  const s = Math.sin(dLat / 2) ** 2 + Math.cos(a[0] * Math.PI / 180) * Math.cos(b[0] * Math.PI / 180) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(s));
}

function decodeCode(digits, ref) {
  const cands = [];
  const push = (d, desc) => {
    try {
      const c = normalizeDigits(d);
      if (!c) return;
      const full = OLC.recoverNearest(c, ref[0], ref[1]);
      cands.push([OLC.decode(full), d, desc]);
    } catch (e) {}
  };
  if (digits.length === 10) push(digits, 'full10');
  else if (digits.length === 8) push(digits, 'short8');
  else if (digits.length === 7) { push(digits.slice(0, 6), 'short6'); push(digits.slice(0, 7), 'short8'); }
  else if (digits.length === 6) push(digits, 'short6');
  else if (digits.length === 5) push(digits.slice(0, 4), 'short4');
  if (!cands.length) return null;
  cands.sort((a, b) => haversineKm([a[0].latitudeCenter, a[0].longitudeCenter], ref) - haversineKm([b[0].latitudeCenter, b[0].longitudeCenter], ref));
  return cands[0];
}

function geocodeRow(cidade, uf, localizacao) {
  const digits = extractCode(localizacao);
  const sheetRef = resolveMuni(cidade, uf);
  const ref = textCityRef(localizacao, uf) || sheetRef;
  if (digits && ref) {
    const res = decodeCode(digits, ref);
    if (res) {
      return {
        lat: +res[0].latitudeCenter.toFixed(6),
        lng: +res[0].longitudeCenter.toFixed(6),
        prec: res[2],
        code: res[1],
        aprox: false
      };
    }
  }
  if (sheetRef) {
    return { lat: +sheetRef[0].toFixed(6), lng: +sheetRef[1].toFixed(6), prec: 'cidade', code: null, aprox: true };
  }
  return null;
}

const out = [];
const semGeo = [];
rows.forEach((r, i) => {
  const geo = geocodeRow(r['CIDADE'], r['ESTADO'], r['LOCALIZAÇÃO']);
  if (!geo) { semGeo.push(`L${i + 2}: ${r['CLIENTE']} ${r['CIDADE']}/${r['ESTADO']}`); return; }
  out.push({
    cl: String(r['CLIENTE'] || ''), un: String(r['UNIDADE'] || ''),
    en: String(r['ENTREGA'] || ''), vi: String(r['ULTIMA VISITA'] || ''),
    sg: String(r['SEGMENTO'] || ''), cn: String(r['CNPJ'] || ''),
    ci: String(r['CIDADE'] || ''), es: String(r['ESTADO'] || ''),
    at: r['ATIVOS'] || 0, km: r['KM'] || '', tm: String(r['TEMPO'] || ''),
    f: String(r['TELEFONE'] || ''), n: String(r['NOME'] || ''),
    e: String(r['E-MAIL'] || ''), ca: String(r['CARGO'] || ''),
    co: String(r['CONSULTOR'] || ''), lo: String(r['LOCALIZAÇÃO'] || ''),
    lat: geo.lat, lng: geo.lng, pr: geo.prec, ap: geo.aprox
  });
});

if (semGeo.length) console.log(`Sem coordenadas (${semGeo.length}):\n  ` + semGeo.join('\n  '));

const muniKeys = Object.keys(MUNI || {});
if (!muniKeys.length) console.log('Aviso: sem municipios em muni.json — js/geo.js nao tera fallback.');
const js = '/* Snapshot gerado por tools/build.mjs em ' + new Date().toISOString().slice(0, 10) +
  ' — ' + out.length + ' unidades e ' + muniKeys.length + ' municipios. Nao edite manualmente. */\n' +
  'var MUNI = ' + JSON.stringify(MUNI || {}) + ';\n' +
  'var ESTADOS = ' + JSON.stringify(ESTADOS) + ';\n' +
  'var SNAPSHOT = ' + JSON.stringify(out) + ';\n';
fs.writeFileSync(outPath, js);
console.log(`Gerado js/data.js com ${out.length} unidades, ${muniKeys.length} municipios e ${Object.keys(ESTADOS).length} estados.`);
