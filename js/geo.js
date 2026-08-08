/* =====================================================================
   geo.js — decodificacao de Plus Codes (Open Location Code) + geocodigo
   de municipios brasileiros (centroides IBGE embutidos em data.js).
   Implementacao do algoritmo OLC baseada na referencia oficial do
   Google (open-location-code, Apache 2.0).
   ===================================================================== */
var GEO = (function () {
  'use strict';

  var SEP = '+';
  var SEP_POS = 8;
  var ALPHABET = '23456789CFGHJMPQRVWX';
  var BASE = ALPHABET.length;
  var LAT_MAX = 90;
  var LNG_MAX = 180;
  var PAIR_RES = [20.0, 1.0, 0.05, 0.0025, 0.000125];
  var UFS = 'RO|AC|AM|RR|PA|AP|TO|MA|PI|CE|RN|PB|PE|AL|SE|BA|MG|ES|RJ|SP|PR|SC|RS|MS|MT|GO|DF';

  function clipLat(lat) { return Math.max(-90, Math.min(90, lat)); }
  function normLng(lng) { return (((lng + 180) % 360) + 360) % 360 - 180; }

  function decodePairs(code, offset) {
    var i = 0, value = 0;
    while (i * 2 + offset < code.length) {
      value += ALPHABET.indexOf(code.charAt(i * 2 + offset)) * PAIR_RES[i];
      i += 1;
    }
    return [value, value + PAIR_RES[i - 1]];
  }

  /* Decodifica um codigo completo (8+2) */
  function decodeFull(code) {
    code = String(code).replace(SEP, '');
    var lat = decodePairs(code, 0);
    var lng = decodePairs(code, 1);
    return {
      latLo: lat[0] - LAT_MAX, latHi: lat[1] - LAT_MAX,
      lngLo: lng[0] - LNG_MAX, lngHi: lng[1] - LNG_MAX,
      latCenter: (lat[0] - LAT_MAX) + (lat[1] - lat[0]) / 2,
      lngCenter: (lng[0] - LNG_MAX) + (lng[1] - lng[0]) / 2,
      codeLength: code.length
    };
  }

  function encodePairs(lat, lng, codeLength) {
    var digits = '', i = 0;
    while (i < codeLength / 2) {
      digits += ALPHABET.charAt(Math.floor(lat / PAIR_RES[i]));
      digits += ALPHABET.charAt(Math.floor(lng / PAIR_RES[i]));
      lat %= PAIR_RES[i];
      lng %= PAIR_RES[i];
      i += 1;
    }
    return digits;
  }

  function encode(lat, lng, codeLength) {
    codeLength = codeLength || 10;
    lat = clipLat(lat) + 90;
    lng = normLng(lng) + 180;
    var code = encodePairs(lat, lng, codeLength);
    return code.substr(0, SEP_POS) + SEP + code.substr(SEP_POS);
  }

  function isShort(code) {
    if (!code) return false;
    var p = code.indexOf(SEP);
    return p > -1 && p < SEP_POS && code.length > p + 1;
  }

  /* Recupera um codigo curto usando um ponto de referencia proximo.
     Retorna o objeto decodificado (mesmo comportamento do lib oficial). */
  function recoverNearest(shortCode, refLat, refLng) {
    refLat = clipLat(refLat);
    refLng = normLng(refLng);
    shortCode = shortCode.toUpperCase();
    var paddingLength = SEP_POS - shortCode.indexOf(SEP);
    var resolution = Math.pow(BASE, 2 - paddingLength / 2);
    var areaToEdge = resolution / 2;
    var roundedLat = Math.floor(refLat / resolution) * resolution;
    var roundedLng = Math.floor(refLng / resolution) * resolution;
    var ca = decodeFull(encode(roundedLat, roundedLng).substr(0, paddingLength) + shortCode);
    var diff = ca.latCenter - refLat;
    if (diff > areaToEdge) ca.latCenter -= resolution;
    else if (diff < -areaToEdge) ca.latCenter += resolution;
    diff = ca.lngCenter - refLng;
    if (diff > areaToEdge) ca.lngCenter -= resolution;
    else if (diff < -areaToEdge) ca.lngCenter += resolution;
    return decodeFull(encode(ca.latCenter, ca.lngCenter, ca.codeLength));
  }

  /* --- utilitarios de normalizacao --- */
  function norm(s) {
    return String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toUpperCase().replace(/[^A-Z0-9]/g, '');
  }

  function resolveMuni(city, uf) {
    if (!city) return null;
    var key = norm(city) + '|' + String(uf || '').toUpperCase();
    if (MUNI[key]) return MUNI[key];
    var est = String(uf || '').toUpperCase();
    return ESTADOS[est] || null;
  }

  /* Extrai o plus code do campo "LOCALIZACAO" (texto livre) */
  function extractCode(loc) {
    var s = String(loc || '');
    var m = s.match(/([2-9CFGHJMPQRVWX]{4,8})\s*\+([2-9CFGHJMPQRVWX]{2,3})/);
    if (m) return m[1] + m[2];
    m = s.match(/([2-9CFGHJMPQRVWX]{8})\b/);
    return m ? m[1] : null;
  }

  /* Tenta achar "Cidade - UF" no proprio texto do localizacao
     (mais confiavel que a coluna CIDADE, que tem erros de digitacao) */
  function textCityRef(loc, uf) {
    var s = String(loc || '');
    var re = new RegExp('([^,;\\-]{3,60}?)\\s*[-,]\\s*(' + UFS + ')\\b', 'g');
    var ms = s.match(re);
    if (!ms || !ms.length) return null;
    var mm = ms[ms.length - 1].match(/([^,;-]{2,60}?)\s*[-,]\s*(' + UFS + ')\b/);
    if (!mm) return null;
    var city = mm[1].trim().replace(/^[\s-]+/, '');
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
    var R = 6371, dLat = (b[0] - a[0]) * Math.PI / 180, dLng = (b[1] - a[1]) * Math.PI / 180;
    var s = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
      Math.cos(a[0] * Math.PI / 180) * Math.cos(b[0] * Math.PI / 180) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
    return 2 * R * Math.asin(Math.sqrt(s));
  }

  /* Decodifica com varias interpretacoes e escolhe a mais proxima do ref */
  function decodeCode(digits, ref) {
    var cands = [];
    function push(d, desc) {
      try {
        var code = normalizeDigits(d);
        if (!code) return;
        var ca = recoverNearest(code, ref[0], ref[1]);
        cands.push([ca, d, desc]);
      } catch (e) {}
    }
    if (digits.length === 10) push(digits, 'full10');
    else if (digits.length === 8) push(digits, 'short8');
    else if (digits.length === 7) { push(digits.slice(0, 6), 'short6'); push(digits.slice(0, 7), 'short8'); }
    else if (digits.length === 6) push(digits, 'short6');
    else if (digits.length === 5) push(digits.slice(0, 4), 'short4');
    if (!cands.length) return null;
    cands.sort(function (a, b) {
      return haversineKm([a[0].latCenter, a[0].lngCenter], ref) - haversineKm([b[0].latCenter, b[0].lngCenter], ref);
    });
    return cands[0];
  }

  /* Resolve a localizacao completa de uma linha da planilha */
  function geocodeRow(cidade, uf, localizacao) {
    var digits = extractCode(localizacao);
    var sheetRef = resolveMuni(cidade, uf);
    var ref = textCityRef(localizacao, uf) || sheetRef;
    if (digits && ref) {
      var res = decodeCode(digits, ref);
      if (res) {
        return {
          lat: +res[0].latCenter.toFixed(6),
          lng: +res[0].lngCenter.toFixed(6),
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

  return { geocodeRow: geocodeRow, norm: norm };
})();
