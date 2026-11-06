/* Leitweg-ID pruefen: checks format and check digits in the browser. Nothing is sent anywhere.
   Rules: Leitweg-ID Format-Spezifikation 2.0.2 (KoSIT, 28.07.2021), check digits ISO/IEC 7064 MOD 97-10. */
(function (root) {
  'use strict';
  var LAENDER = {
    '01': 'Schleswig-Holstein', '02': 'Hamburg', '03': 'Niedersachsen', '04': 'Bremen', '05': 'Nordrhein-Westfalen',
    '06': 'Hessen', '07': 'Rheinland-Pfalz', '08': 'Baden-Württemberg', '09': 'Bayern', '10': 'Saarland', '11': 'Berlin',
    '12': 'Brandenburg', '13': 'Mecklenburg-Vorpommern', '14': 'Sachsen', '15': 'Sachsen-Anhalt', '16': 'Thüringen', '99': 'Bund'
  };
  // Grobadressierung: Land (2) + Regierungsbezirk (1) + Landkreis (2) + Gemeinde (3, 4 or 7 digits).
  var GROB_LENGTHS = [2, 3, 5, 8, 9, 12];

  function mod97(digits) {
    var r = 0;
    for (var i = 0; i < digits.length; i++) { r = (r * 10 + (digits.charCodeAt(i) - 48)) % 97; }
    return r;
  }
  function toDigits(s) {
    var out = '';
    for (var i = 0; i < s.length; i++) {
      var c = s.charAt(i);
      out += (c >= 'A' && c <= 'Z') ? String(c.charCodeAt(0) - 55) : c;
    }
    return out;
  }
  function expectedCheck(grob, fein) {
    var p = 98 - mod97(toDigits(grob + fein) + '00');
    return (p < 10 ? '0' : '') + p;
  }

  // Returns { ok, normalized, parts, land, checks: [{ ok, text }], expected }
  function check(input) {
    var raw = String(input == null ? '' : input);
    var notes = [];
    var s = raw.replace(/[\u2010\u2011\u2012\u2013\u2014\u2212]/g, function () { notes.push('dash'); return '-'; });
    s = s.replace(/\s+/g, function (m) { if (m.length) { notes.push('space'); } return ''; }).toUpperCase();
    var res = { ok: false, normalized: s, parts: null, land: null, checks: [], expected: null };
    function add(ok, text) { res.checks.push({ ok: ok, text: text }); return ok; }

    if (s === '') { add(false, 'Bitte eine Leitweg-ID eingeben.'); return res; }
    if (notes.indexOf('space') >= 0) { add(true, 'Leerzeichen entfernt. In der Rechnung steht die Leitweg-ID ohne Leerzeichen.'); }
    if (notes.indexOf('dash') >= 0) { add(true, 'Langer Strich durch den normalen Bindestrich ersetzt. Erlaubt ist nur der Bindestrich.'); }
    if (!add(/^[0-9A-Z-]+$/.test(s), /^[0-9A-Z-]+$/.test(s) ? 'Nur erlaubte Zeichen: Ziffern, Buchstaben A bis Z und Bindestriche.'
      : 'Unerlaubte Zeichen. Erlaubt sind Ziffern, Buchstaben A bis Z (ohne Umlaute) und Bindestriche.')) { return res; }
    if (!add(s.length >= 5 && s.length <= 46, 'Länge ' + s.length + ' Zeichen (erlaubt sind 5 bis 46).')) { return res; }

    var parts = s.split('-');
    if (parts.length < 2 || parts.length > 3) {
      add(false, 'Eine Leitweg-ID hat zwei oder drei Teile, getrennt durch Bindestriche: Grobadressierung, optional Feinadressierung, Prüfziffer.');
      return res;
    }
    var grob = parts[0];
    var fein = parts.length === 3 ? parts[1] : '';
    var pruef = parts[parts.length - 1];
    res.parts = { grob: grob, fein: fein, pruef: pruef };

    var grobOk = /^\d+$/.test(grob) && GROB_LENGTHS.indexOf(grob.length) >= 0;
    add(grobOk, grobOk ? 'Grobadressierung ' + grob + ': ' + grob.length + ' Ziffern, gültige Länge.'
      : 'Grobadressierung "' + grob + '": muss aus 2, 3, 5, 8, 9 oder 12 Ziffern bestehen.');
    var land = /^\d{2}/.test(grob) ? LAENDER[grob.slice(0, 2)] : null;
    res.land = land || null;
    add(!!land, land ? 'Die ersten beiden Ziffern ' + grob.slice(0, 2) + ' stehen für: ' + land + '.'
      : 'Die ersten beiden Ziffern müssen 01 bis 16 (ein Bundesland) oder 99 (Bund) sein.');
    if (parts.length === 3) {
      var feinOk = /^[0-9A-Z]{1,30}$/.test(fein);
      add(feinOk, feinOk ? 'Feinadressierung ' + fein + ': ' + fein.length + ' Zeichen, gültig.'
        : 'Feinadressierung "' + fein + '": 1 bis 30 Zeichen, nur Ziffern und Buchstaben.');
    } else {
      add(true, 'Keine Feinadressierung. Das ist erlaubt.');
    }
    var pruefOk = /^\d{2}$/.test(pruef);
    add(pruefOk, pruefOk ? 'Prüfziffer ' + pruef + ': zwei Ziffern.' : 'Die Prüfziffer am Ende muss aus genau zwei Ziffern bestehen.');

    if (/^\d+$/.test(grob) && /^[0-9A-Z]*$/.test(fein)) {
      res.expected = expectedCheck(grob, fein);
    }
    if (pruefOk && res.expected !== null) {
      var sumOk = mod97(toDigits(grob + fein + pruef)) === 1;
      add(sumOk, sumOk ? 'Die Prüfziffer stimmt (Modulo 97-10).'
        : 'Die Prüfziffer stimmt nicht. Zu den Zeichen davor passt die Prüfziffer ' + res.expected + '. Bitte die Leitweg-ID beim Auftraggeber nachfragen, nicht selbst ändern.');
    }
    res.ok = res.checks.every(function (c) { return c.ok; });
    return res;
  }

  var api = { check: check, expectedCheck: expectedCheck, LAENDER: LAENDER };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; }
  root.LeitwegPruefen = api;

  if (typeof document === 'undefined') { return; }
  function render() {
    var input = document.getElementById('lw-input');
    var out = document.getElementById('lw-result');
    if (!input || !out) { return; }
    var r = check(input.value);
    out.textContent = '';
    var head = document.createElement('p');
    head.className = r.ok ? 'lw-ok' : 'lw-bad';
    head.textContent = r.ok ? 'Gültig: ' + r.normalized + (r.land ? ' (' + r.land + ')' : '')
      : (r.normalized ? 'Nicht gültig: ' + r.normalized : 'Noch keine Eingabe.');
    out.appendChild(head);
    var ul = document.createElement('ul');
    r.checks.forEach(function (c) {
      var li = document.createElement('li');
      li.textContent = (c.ok ? 'OK: ' : 'Fehler: ') + c.text;
      ul.appendChild(li);
    });
    out.appendChild(ul);
    if (r.ok) {
      var hint = document.createElement('p');
      hint.textContent = 'Eine gültige Prüfziffer zeigt nur, dass die Leitweg-ID richtig aufgebaut ist. Ob sie vergeben ist und zu Ihrem Auftrag gehört, weiß nur Ihr Auftraggeber.';
      out.appendChild(hint);
    }
  }
  document.addEventListener('DOMContentLoaded', function () {
    var form = document.getElementById('lw-form');
    if (!form) { return; }
    form.addEventListener('submit', function (e) { e.preventDefault(); render(); });
    var ex = document.getElementById('lw-example');
    if (ex) {
      ex.addEventListener('click', function (e) {
        e.preventDefault();
        document.getElementById('lw-input').value = '04011000-1234512345-06';
        render();
      });
    }
  });
})(typeof window !== 'undefined' ? window : this);
