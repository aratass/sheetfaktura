/* SheetFaktura für Excel 1.0.1. Open source components: assets/lib/LICENSES.txt. */
window.SF_WEB = {"version":"1.0.1","libs":[{"global":"eInvoiceEU","src":"assets/lib/e-invoice-eu-core-3.3.3.min.js","integrity":"sha384-B0PLO8jTGFjm0irRadzSsbk5l0NTFYl0C3Q2dsVtTJFM94gAxoJD8Vggo+87olAc"},{"global":"PDFLib","src":"assets/lib/cantoo-pdf-lib-2.11.1.min.js","integrity":"sha384-BcRbGROz3cx4LSKcOw2L5YxAxiJPWIEmXRMrYbXu87aSiU5uMQ4qB9v8FEjrBhhq"},{"global":"fontkit","src":"assets/lib/pdf-lib-fontkit-1.1.1.umd.min.js","integrity":"sha384-2p6U+1mmqF10USehFeRiyG2ESG9FwIqN+jxULn5w9jjQIihSn9Pt13dVCn/Hawjn"},{"global":"EINVOICE_FONTS","src":"assets/excel-fonts.79cfc8eb6a.js","integrity":"sha384-+EfcfO8243Fyknsji0OpYQzquqQK3nwuN9MvjMjX7xcGV519o5bIjT4+8ZllcE6V"}]};
window.SF_PAGE_EN = {"navStart":"Home","navGuide":"Guide","navHelp":"Help","navPrivacy":"Privacy","navTerms":"Terms","navImprint":"Legal notice","h1":"Create e-invoices from Excel","lead":"You write your invoices in Excel? With the free SheetFaktura template you turn your Excel invoice into an XRechnung (XML) and a ZUGFeRD PDF. The file is read only in your browser and is not uploaded. You do not need a Google account.","s1t":"1. Download the template","dlEmpty":"Empty template","dlExample":"Example invoice","s1":"Excel files (.xlsx) of 14 and 15 KB.","s2t":"2. Fill it in Excel","s2":"Fill in the yellow fields and save the file as .xlsx. LibreOffice and Numbers work too.","s3t":"3. Pick the file here","s3":"SheetFaktura checks the data, names every problem with its cell and creates the XRechnung and the ZUGFeRD PDF.","toolTitle":"Turn your Excel invoice into XRechnung and ZUGFeRD","privacyNote":"Your file is read only in your browser. The XRechnung and the PDF are made there too; nothing is uploaded. The check covers the format and is not tax advice.","tipsTitle":"Good to know","tip1":"Excel first opens files from the internet in Protected View. Click <i>Enable Editing</i> there to fill in the template.","tip2":"For each new invoice, copy the filled sheet (right-click the sheet name, <i>Move or Copy</i>, <i>Create a copy</i>) and change the customer, invoice number, date and line items. If the file has several invoice sheets, pick the right one here.","tip3":"For more line items insert rows above <i>Summe Positionen / Items net total</i> and copy the formula of the <i>Net</i> column from the row above.","tip4":"The buyer reference is required: the Leitweg-ID for public bodies, for companies for example their order number. Without a Leitweg-ID the invoice needs the email address at which your customer receives e-invoices.","tip5":"Agree with your customer which file they get: the XRechnung is the pure data format, the ZUGFeRD PDF can also be read as a document. Both contain the same invoice.","faqTitle":"Questions","sheetsTitle":"Prefer Google Sheets?","sheetsText":"SheetFaktura is also available as an <a href=\"index.html#english\">add-on for Google Sheets</a> with the same template and the same checks. Both ways are explained in the guide <a href=\"xrechnung-aus-excel.html\">XRechnung aus Excel erstellen</a> (in German). When e-invoices are required: <a href=\"e-rechnung-pflicht.html\">E-Rechnungspflicht 2025 bis 2028</a> (in German).","disclaimer":"SheetFaktura checks the format of your invoice; it does not replace tax advice.","faqUploadQ":"Is my data uploaded?","faqUploadA":"No. Your Excel file is read only in your browser, and the XRechnung and the PDF are made there too. Nothing is sent to us. Only when you enter a licence key does your browser check the key with our payment provider Polar. More in the <a href=\"privacy.html\">privacy policy</a>.","faqVersionsQ":"Which Excel versions work?","faqVersionsA":"Any version that saves files in .xlsx format: Excel 2007 and later, Microsoft 365, Excel for Mac and Excel on the web. Save an older .xls file with <i>Save As</i> as <i>Excel Workbook (.xlsx)</i>. The file must not be protected with a password.","faqLibreQ":"Does it work with LibreOffice or Numbers?","faqLibreA":"Yes, if you save the file as .xlsx. In LibreOffice Calc: <i>File</i>, <i>Save As</i>, file type <i>Excel 2007-365 (.xlsx)</i>. In Numbers: <i>File</i>, <i>Export To</i>, <i>Excel</i>.","faqPdfQ":"Can I convert an existing PDF invoice?","faqPdfA":"No. A PDF cannot be turned into an e-invoice directly, because it lacks the structured invoice data. You enter the details once in the Excel template and get both files from it: the XRechnung and the ZUGFeRD PDF.","faqLayoutQ":"Can I use my existing Excel invoice?","faqLayoutA":"SheetFaktura reads only its own template, not any layout. Copy your details into the yellow fields once. For the next invoices you copy the sheet and change only the customer, number, date and line items.","faqSmallQ":"Do I need e-invoices as a small business (Kleinunternehmer)?","faqSmallA":"Small businesses under section 19 of the German VAT Act (UStG) do not have to issue e-invoices, but they must be able to receive them. If a customer still wants an XRechnung, choose the VAT type <i>E Steuerfrei / exempt</i> in the template and enter a reason for the exemption, for example „Kein Ausweis der Umsatzsteuer, da Kleinunternehmer gemäß § 19 UStG“. More (in German): <a href=\"e-rechnung-kleinunternehmer.html\">E-Rechnung für Kleinunternehmer</a>.","faqPriceQ":"What does it cost?","faqPriceA":"5 e-invoices are free, counted in your browser. After that 7.90 EUR a month, cancel monthly, or 69 EUR a year, plus VAT. One licence key works here and in the Google Sheets add-on. XML and PDF at no extra cost, no limit on the number of invoices.","faqErrorsQ":"What do I do when a problem is shown?","faqErrorsA":"Each message names the cell, for example <i>Cell C6</i>. Correct the entry in Excel, save the file and pick it here again."};
/*
 * E-Rechnung core. Pure functions shared by the sidebar (browser), the Apps Script
 * server and the Node tests: template layout, sheet parsing, checks with German and
 * English messages, EN 16931 totals, mapping to the @e-invoice-eu/core input format,
 * and the licence decisions.
 *
 * Syntax stays at ES2017 (no optional chaining, no BigInt literals) so the same file
 * also loads on the Apps Script V8 runtime. Money math uses BigInt through BigInt().
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) {
    module.exports = api;
  } else {
    root.EInvoiceCore = api;
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var VERSION = '1.0.1';

  // ---------------------------------------------------------------------------
  // Configuration. Values in CAPITALS are placeholders filled in at deployment.
  // ---------------------------------------------------------------------------
  var CONFIG = {
    appName: 'SheetFaktura',
    freeInvoices: 5,
    priceMonthly: '7,90 EUR',
    priceYearly: '69 EUR',
    licenceRecheckHours: 24,
    licenceRetryMinutes: 2,
    licenceGraceDays: 7,
    polarTimeoutMs: 12000, // a licence check that hangs ends as a network error; the grace period decides (LIC-3)
    libTimeoutMs: 30000, // per library script from jsDelivr
    polar: {
      apiBase: 'https://api.polar.sh',
      organizationId: 'POLAR_ORGANIZATION_ID',
      benefitId: '',
      checkoutMonthly: 'POLAR_CHECKOUT_LINK_MONTHLY',
      checkoutYearly: 'POLAR_CHECKOUT_LINK_YEARLY',
      customerPortal: 'POLAR_CUSTOMER_PORTAL_LINK'
    },
    links: {
      help: 'https://sheetfaktura.de/hilfe.html',
      privacy: 'https://sheetfaktura.de/datenschutz.html',
      terms: 'https://sheetfaktura.de/agb.html',
      // English pages for the English interface (UI-8).
      helpEn: 'https://sheetfaktura.de/help.html',
      privacyEn: 'https://sheetfaktura.de/privacy.html',
      termsEn: 'https://sheetfaktura.de/terms.html'
    }
  };

  // Deep-merges deployment values (deploy.config.json) into CONFIG.
  function configure(overrides) {
    (function merge(target, src) {
      Object.keys(src || {}).forEach(function (k) {
        if (src[k] && typeof src[k] === 'object' && !Array.isArray(src[k]) && target[k] && typeof target[k] === 'object') {
          merge(target[k], src[k]);
        } else {
          target[k] = src[k];
        }
      });
    }(CONFIG, overrides));
    return CONFIG;
  }
  // Lists config values that still hold a placeholder (capital letters with an underscore, or DOMAIN).
  function configPlaceholders() {
    var found = [];
    (function walk(o, p) {
      Object.keys(o).forEach(function (k) {
        var v = o[k];
        if (v && typeof v === 'object') { walk(v, p + k + '.'); }
        else if (typeof v === 'string' && (/^[A-Z][A-Z0-9_]*_[A-Z0-9_]+$/.test(v) || v.indexOf('://DOMAIN/') >= 0)) { found.push(p + k); }
      });
    }(CONFIG, ''));
    return found;
  }

  // ---------------------------------------------------------------------------
  // Decimal arithmetic on strings, backed by BigInt. A decimal is {i: BigInt, s: scale}.
  // ---------------------------------------------------------------------------
  function big(n) { return BigInt(n); }
  var POW10 = [];
  function pow10(n) {
    if (!POW10[n]) {
      var r = big(1);
      for (var k = 0; k < n; k++) { r = r * big(10); }
      POW10[n] = r;
    }
    return POW10[n];
  }
  function dParse(str) {
    var m = /^([+-])?(\d*)(?:\.(\d*))?$/.exec(String(str));
    if (!m || ((m[2] || '') === '' && (m[3] || '') === '')) { throw new Error('Not a decimal: ' + str); }
    var frac = m[3] || '';
    var v = big((m[2] || '0') + frac);
    if (m[1] === '-') { v = -v; }
    return { i: v, s: frac.length };
  }
  function dAlign(a, b) {
    if (a.s === b.s) { return [a, b]; }
    if (a.s < b.s) { return [{ i: a.i * pow10(b.s - a.s), s: b.s }, b]; }
    return [a, { i: b.i * pow10(a.s - b.s), s: a.s }];
  }
  function dAdd(a, b) { var x = dAlign(a, b); return { i: x[0].i + x[1].i, s: x[0].s }; }
  function dSub(a, b) { var x = dAlign(a, b); return { i: x[0].i - x[1].i, s: x[0].s }; }
  function dMul(a, b) { return { i: a.i * b.i, s: a.s + b.s }; }
  function dCmp(a, b) { var x = dAlign(a, b); return x[0].i < x[1].i ? -1 : (x[0].i > x[1].i ? 1 : 0); }
  function dRound(a, places) {
    if (a.s <= places) { return { i: a.i * pow10(places - a.s), s: places }; }
    var f = pow10(a.s - places);
    var q = a.i / f;
    var r = a.i % f;
    if (r < big(0)) { r = -r; }
    if (r * big(2) >= f) { q = q + (a.i < big(0) ? big(-1) : big(1)); }
    return { i: q, s: places };
  }
  function dToString(a) {
    var neg = a.i < big(0);
    var digits = (neg ? -a.i : a.i).toString();
    if (a.s > 0) {
      while (digits.length <= a.s) { digits = '0' + digits; }
      digits = digits.slice(0, digits.length - a.s) + '.' + digits.slice(digits.length - a.s);
    }
    return (neg && /[1-9]/.test(digits) ? '-' : '') + digits;
  }
  // Public helpers working on decimal strings.
  function add(a, b) { return dToString(dAdd(dParse(a), dParse(b))); }
  function sub(a, b) { return dToString(dSub(dParse(a), dParse(b))); }
  function mul(a, b) { return dToString(dMul(dParse(a), dParse(b))); }
  function percentOf(base, percent) { // base * percent / 100, unrounded
    var p = dMul(dParse(base), dParse(percent));
    return dToString({ i: p.i, s: p.s + 2 });
  }
  function round2(a) { return dToString(dRound(dParse(a), 2)); }
  // a / b rounded half away from zero to the given places.
  function divRound(a, b, places) {
    var x = dParse(a);
    var y = dParse(b);
    var num = x.i * pow10(y.s) * pow10(places);
    var den = y.i * pow10(x.s);
    if (den < big(0)) { num = -num; den = -den; }
    var q = num / den;
    var r = num % den;
    if (r < big(0)) { r = -r; }
    if (r * big(2) >= den) { q = q + (num < big(0) ? big(-1) : big(1)); }
    return dToString({ i: q, s: places });
  }
  function cmp(a, b) { return dCmp(dParse(a), dParse(b)); }
  function isZero(a) { return dParse(a).i === big(0); }
  function normalizeDecimal(a) { // strip trailing zeros: "1.500" -> "1.5", "2.00" -> "2"
    var s = dToString(dParse(a));
    if (s.indexOf('.') >= 0) { s = s.replace(/0+$/, '').replace(/\.$/, ''); }
    return s === '-0' ? '0' : s;
  }
  function sum(list) {
    var acc = { i: big(0), s: 0 };
    for (var k = 0; k < list.length; k++) { acc = dAdd(acc, dParse(list[k])); }
    return dToString(acc);
  }

  // A JavaScript number from the sheet to an exact decimal string (15 significant digits).
  function numberToDecimalString(n) {
    if (typeof n !== 'number' || !isFinite(n)) { return null; }
    var x = Number(n.toPrecision(15));
    var s = String(x);
    if (/e/i.test(s)) {
      s = x.toFixed(20).replace(/0+$/, '').replace(/\.$/, '');
    }
    return s;
  }

  // ---------------------------------------------------------------------------
  // Parsing cell values.
  // ---------------------------------------------------------------------------
  function isBlank(v) {
    return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  }
  function text(v) {
    if (v === null || v === undefined) { return ''; }
    if (typeof v === 'number') { return numberToDecimalString(v) || ''; }
    // Control characters (vertical tab, form feed and others from pasted text) are not allowed in XML (INV-7).
    return String(v).replace(/\u00a0/g, ' ').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ' ').trim();
  }
  // Accepts numbers and German or English formatted strings. Returns a decimal string or null.
  function parseNumber(v) {
    if (typeof v === 'number') { return numberToDecimalString(v); }
    if (typeof v === 'boolean' || isBlank(v)) { return null; }
    var s = String(v).replace(/[\s  ']/g, '').replace(/(EUR|€|%)$/i, '').replace(/^(EUR|€)/i, '');
    if (s === '') { return null; }
    var lastDot = s.lastIndexOf('.');
    var lastComma = s.lastIndexOf(',');
    if (lastDot >= 0 && lastComma >= 0) {
      if (lastComma > lastDot) { s = s.replace(/\./g, '').replace(',', '.'); } else { s = s.replace(/,/g, ''); }
    } else if (lastComma >= 0) {
      s = (s.split(',').length > 2) ? s.replace(/,/g, '') : s.replace(',', '.');
    } else if (lastDot >= 0 && s.split('.').length > 2) {
      s = s.replace(/\./g, '');
    }
    if (!/^[+-]?(\d+(\.\d*)?|\.\d+)$/.test(s)) { return null; }
    if (s.charAt(0) === '+') { s = s.slice(1); }
    if (s.charAt(0) === '.' || s.slice(0, 2) === '-.') { s = s.replace('.', '0.'); }
    return normalizeDecimal(s);
  }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  // Years before 2000 are typos for an invoice tool: "011026" typed as a number is the serial of 9 March 1930 (SRV-11).
  function validYmd(y, m, d) {
    if (y < 2000 || y > 2200 || m < 1 || m > 12 || d < 1) { return false; }
    var dim = [31, (y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0)) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    return d <= dim[m - 1];
  }
  // Accepts "YYYY-MM-DD", "DD.MM.YYYY", "D.M.YY" and spreadsheet serial numbers. Returns "YYYY-MM-DD" or null.
  function parseDate(v) {
    if (isBlank(v)) { return null; }
    var y, m, d, mm;
    if (typeof v === 'number') {
      if (v < 36526 || v > 200000) { return null; } // 36526 is 1 January 2000
      var ms = Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000;
      var dt = new Date(ms);
      return dt.getUTCFullYear() + '-' + pad2(dt.getUTCMonth() + 1) + '-' + pad2(dt.getUTCDate());
    }
    if (v instanceof Date) {
      if (isNaN(v.getTime()) || v.getFullYear() < 2000) { return null; }
      return v.getFullYear() + '-' + pad2(v.getMonth() + 1) + '-' + pad2(v.getDate());
    }
    var s = String(v).trim();
    if ((mm = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ].*)?$/.exec(s))) {
      y = +mm[1]; m = +mm[2]; d = +mm[3];
    } else if ((mm = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/.exec(s))) {
      d = +mm[1]; m = +mm[2]; y = +mm[3];
      if (mm[3].length === 2) { y += 2000; }
    } else {
      return null;
    }
    if (!validYmd(y, m, d)) { return null; }
    return y + '-' + pad2(m) + '-' + pad2(d);
  }

  // ---------------------------------------------------------------------------
  // Code lists.
  // ---------------------------------------------------------------------------
  var DOC_TYPES = [
    { code: '380', de: 'Rechnung', en: 'Invoice' },
    { code: '381', de: 'Rechnungskorrektur (Gutschrift)', en: 'Credit note' },
    { code: '326', de: 'Teilrechnung', en: 'Partial invoice' }
  ];
  var VAT_CATEGORIES = {
    S: { de: 'Regelsatz', en: 'Standard rate', pick: 'Regelsatz / standard', zeroRate: false },
    Z: { de: 'Nullsatz', en: 'Zero rate', pick: 'Nullsatz / zero', zeroRate: true },
    E: { de: 'Steuerfrei', en: 'Exempt', pick: 'Steuerfrei / exempt', zeroRate: true, reasonRequired: true },
    AE: { de: 'Reverse Charge', en: 'Reverse Charge', pick: 'Reverse Charge', zeroRate: true, code: 'VATEX-EU-AE',
      textDe: 'Steuerschuldnerschaft des Leistungsempfängers', textEn: 'Reverse charge: the recipient is liable for VAT' },
    K: { de: 'Innergemeinschaftliche Lieferung', en: 'Intra-EU supply', pick: 'EU-Lieferung / intra-EU', zeroRate: true, code: 'VATEX-EU-IC',
      textDe: 'Steuerfreie innergemeinschaftliche Lieferung', textEn: 'VAT exempt intra-community supply' },
    G: { de: 'Ausfuhr', en: 'Export', pick: 'Ausfuhr / export', zeroRate: true, code: 'VATEX-EU-G',
      textDe: 'Steuerfreie Ausfuhrlieferung', textEn: 'VAT exempt export outside the EU' },
    O: { de: 'Nicht steuerbar', en: 'Not subject to VAT', pick: 'Nicht steuerbar / out of scope', zeroRate: true, noRate: true, code: 'VATEX-EU-O',
      textDe: 'Nicht im Inland steuerbare Leistung', textEn: 'Not subject to VAT' }
  };
  var VAT_CATEGORY_ORDER = ['S', 'Z', 'E', 'AE', 'K', 'G', 'O'];
  var UNITS = [
    { code: 'H87', de: 'Stück', en: 'piece', sde: 'Stk.', sen: 'pc', aliases: ['STK', 'STÜCK', 'STUECK', 'ST', 'PCS', 'PIECE', 'C62'] },
    { code: 'HUR', de: 'Stunde', en: 'hour', sde: 'Std.', sen: 'h', aliases: ['STD', 'STUNDE', 'STUNDEN', 'H', 'HOUR', 'HOURS'] },
    { code: 'DAY', de: 'Tag', en: 'day', sde: 'Tag', sen: 'day', aliases: ['TAG', 'TAGE', 'D', 'DAYS'] },
    { code: 'WEE', de: 'Woche', en: 'week', sde: 'Wo.', sen: 'wk', aliases: ['WOCHE', 'WOCHEN', 'WEEK'] },
    { code: 'MON', de: 'Monat', en: 'month', sde: 'Mon.', sen: 'mo', aliases: ['MONAT', 'MONATE', 'MONTH'] },
    { code: 'ANN', de: 'Jahr', en: 'year', sde: 'Jahr', sen: 'yr', aliases: ['JAHR', 'JAHRE', 'YEAR'] },
    { code: 'MIN', de: 'Minute', en: 'minute', sde: 'Min.', sen: 'min', aliases: ['MINUTE', 'MINUTEN'] },
    { code: 'LS', de: 'Pauschale', en: 'lump sum', sde: 'pausch.', sen: 'lump sum', aliases: ['PAUSCHAL', 'PAUSCHALE', 'PSCH', 'PA'] },
    { code: 'KGM', de: 'Kilogramm', en: 'kilogram', sde: 'kg', sen: 'kg', aliases: ['KG'] },
    { code: 'TNE', de: 'Tonne', en: 'tonne', sde: 't', sen: 't', aliases: ['T', 'TONNE'] },
    { code: 'MTR', de: 'Meter', en: 'metre', sde: 'm', sen: 'm', aliases: ['M', 'METER'] },
    { code: 'KMT', de: 'Kilometer', en: 'kilometre', sde: 'km', sen: 'km', aliases: ['KM'] },
    { code: 'MTK', de: 'Quadratmeter', en: 'square metre', sde: 'm²', sen: 'm²', aliases: ['M2', 'QM', 'M²'] },
    { code: 'MTQ', de: 'Kubikmeter', en: 'cubic metre', sde: 'm³', sen: 'm³', aliases: ['M3', 'CBM', 'M³'] },
    { code: 'LTR', de: 'Liter', en: 'litre', sde: 'l', sen: 'l', aliases: ['L', 'LITER'] },
    // XPK (package, UN/ECE Rec 21). 1.0.0 offered 'Packung (PK)', which is no valid unit code; sheets made then still
    // hold that text in their dropdown list, so PK stays an alias.
    { code: 'XPK', de: 'Packung', en: 'pack', sde: 'Pck.', sen: 'pk', aliases: ['PACKUNG', 'PAKET', 'PCK', 'PK'] },
    { code: 'SET', de: 'Satz', en: 'set', sde: 'Satz', sen: 'set', aliases: ['SATZ'] },
    { code: 'KWH', de: 'Kilowattstunde', en: 'kilowatt hour', sde: 'kWh', sen: 'kWh', aliases: ['KWH'] }
  ];
  // UN/ECE Recommendation 20 and 21 unit codes as listed by @e-invoice-eu/core 3.3.3 (invoiceSchema
  // $defs.codeLists.UNECERec20), the list the library checks when it creates the files. Bundled so that Check gives
  // the same answer before the libraries have loaded; test/unit/fixes.test.js keeps it equal to the library's list.
  var UNIT_CODES = (
    '10 11 13 14 15 20 21 22 23 24 25 27 28 33 34 35 37 38 40 41 56 57 58 59 60 61 74 77 80 81 85 87 89 91 1I 2A 2B ' +
    '2C 2G 2H 2I 2J 2K 2L 2M 2N 2P 2Q 2R 2U 2X 2Y 2Z 3B 3C 4C 4G 4H 4K 4L 4M 4N 4O 4P 4Q 4R 4T 4U 4W 4X 5A 5B 5E 5J ' +
    'A10 A11 A12 A13 A14 A15 A16 A17 A18 A19 A2 A20 A21 A22 A23 A24 A26 A27 A28 A29 A3 A30 A31 A32 A33 A34 A35 A36 ' +
    'A37 A38 A39 A4 A40 A41 A42 A43 A44 A45 A47 A48 A49 A5 A53 A54 A55 A56 A59 A6 A68 A69 A7 A70 A71 A73 A74 A75 A76 ' +
    'A8 A84 A85 A86 A87 A88 A89 A9 A90 A91 A93 A94 A95 A96 A97 A98 A99 AA AB ACR ACT AD AE AH AI AK AL AMH AMP ANN ' +
    'APZ AQ AS ASM ASU ATM AWG AY AZ B1 B10 B11 B12 B13 B14 B15 B16 B17 B18 B19 B20 B21 B22 B23 B24 B25 B26 B27 B28 ' +
    'B29 B3 B30 B31 B32 B33 B34 B35 B4 B41 B42 B43 B44 B45 B46 B47 B48 B49 B50 B52 B53 B54 B55 B56 B57 B58 B59 B60 ' +
    'B61 B62 B63 B64 B66 B67 B68 B69 B7 B70 B71 B72 B73 B74 B75 B76 B77 B78 B79 B8 B80 B81 B82 B83 B84 B85 B86 B87 ' +
    'B88 B89 B90 B91 B92 B93 B94 B95 B96 B97 B98 B99 BAR BB BFT BHP BIL BLD BLL BP BPM BQL BTU BUA BUI C0 C10 C11 C12 ' +
    'C13 C14 C15 C16 C17 C18 C19 C20 C21 C22 C23 C24 C25 C26 C27 C28 C29 C3 C30 C31 C32 C33 C34 C35 C36 C37 C38 C39 ' +
    'C40 C41 C42 C43 C44 C45 C46 C47 C48 C49 C50 C51 C52 C53 C54 C55 C56 C57 C58 C59 C60 C61 C62 C63 C64 C65 C66 C67 ' +
    'C68 C69 C7 C70 C71 C72 C73 C74 C75 C76 C78 C79 C8 C80 C81 C82 C83 C84 C85 C86 C87 C88 C89 C9 C90 C91 C92 C93 C94 ' +
    'C95 C96 C97 C99 CCT CDL CEL CEN CG CGM CKG CLF CLT CMK CMQ CMT CNP CNT COU CTG CTM CTN CUR CWA CWI D03 D04 D1 ' +
    'D10 D11 D12 D13 D15 D16 D17 D18 D19 D2 D20 D21 D22 D23 D24 D25 D26 D27 D29 D30 D31 D32 D33 D34 D36 D41 D42 D43 ' +
    'D44 D45 D46 D47 D48 D49 D5 D50 D51 D52 D53 D54 D55 D56 D57 D58 D59 D6 D60 D61 D62 D63 D65 D68 D69 D73 D74 D77 ' +
    'D78 D80 D81 D82 D83 D85 D86 D87 D88 D89 D91 D93 D94 D95 DAA DAD DAY DB DBM DBW DD DEC DG DJ DLT DMA DMK DMO DMQ ' +
    'DMT DN DPC DPR DPT DRA DRI DRL DT DTN DWT DZN DZP E01 E07 E08 E09 E10 E12 E14 E15 E16 E17 E18 E19 E20 E21 E22 ' +
    'E23 E25 E27 E28 E30 E31 E32 E33 E34 E35 E36 E37 E38 E39 E4 E40 E41 E42 E43 E44 E45 E46 E47 E48 E49 E50 E51 E52 ' +
    'E53 E54 E55 E56 E57 E58 E59 E60 E61 E62 E63 E64 E65 E66 E67 E68 E69 E70 E71 E72 E73 E74 E75 E76 E77 E78 E79 E80 ' +
    'E81 E82 E83 E84 E85 E86 E87 E88 E89 E90 E91 E92 E93 E94 E95 E96 E97 E98 E99 EA EB EQ F01 F02 F03 F04 F05 F06 F07 ' +
    'F08 F10 F11 F12 F13 F14 F15 F16 F17 F18 F19 F20 F21 F22 F23 F24 F25 F26 F27 F28 F29 F30 F31 F32 F33 F34 F35 F36 ' +
    'F37 F38 F39 F40 F41 F42 F43 F44 F45 F46 F47 F48 F49 F50 F51 F52 F53 F54 F55 F56 F57 F58 F59 F60 F61 F62 F63 F64 ' +
    'F65 F66 F67 F68 F69 F70 F71 F72 F73 F74 F75 F76 F77 F78 F79 F80 F81 F82 F83 F84 F85 F86 F87 F88 F89 F90 F91 F92 ' +
    'F93 F94 F95 F96 F97 F98 F99 FAH FAR FBM FC FF FH FIT FL FNU FOT FP FR FS FTK FTQ G01 G04 G05 G06 G08 G09 G10 G11 ' +
    'G12 G13 G14 G15 G16 G17 G18 G19 G2 G20 G21 G23 G24 G25 G26 G27 G28 G29 G3 G30 G31 G32 G33 G34 G35 G36 G37 G38 ' +
    'G39 G40 G41 G42 G43 G44 G45 G46 G47 G48 G49 G50 G51 G52 G53 G54 G55 G56 G57 G58 G59 G60 G61 G62 G63 G64 G65 G66 ' +
    'G67 G68 G69 G70 G71 G72 G73 G74 G75 G76 G77 G78 G79 G80 G81 G82 G83 G84 G85 G86 G87 G88 G89 G90 G91 G92 G93 G94 ' +
    'G95 G96 G97 G98 G99 GB GBQ GDW GE GF GFI GGR GIA GIC GII GIP GJ GL GLD GLI GLL GM GO GP GQ GRM GRN GRO GV GWH ' +
    'H03 H04 H05 H06 H07 H08 H09 H10 H11 H12 H13 H14 H15 H16 H18 H19 H20 H21 H22 H23 H24 H25 H26 H27 H28 H29 H30 H31 ' +
    'H32 H33 H34 H35 H36 H37 H38 H39 H40 H41 H42 H43 H44 H45 H46 H47 H48 H49 H50 H51 H52 H53 H54 H55 H56 H57 H58 H59 ' +
    'H60 H61 H62 H63 H64 H65 H66 H67 H68 H69 H70 H71 H72 H73 H74 H75 H76 H77 H79 H80 H81 H82 H83 H84 H85 H87 H88 H89 ' +
    'H90 H91 H92 H93 H94 H95 H96 H98 H99 HA HAD HBA HBX HC HDW HEA HGM HH HIU HKM HLT HM HMO HMQ HMT HPA HTZ HUR HWE ' +
    'IA IE INH INK INQ ISD IU IUG IV J10 J12 J13 J14 J15 J16 J17 J18 J19 J2 J20 J21 J22 J23 J24 J25 J26 J27 J28 J29 ' +
    'J30 J31 J32 J33 J34 J35 J36 J38 J39 J40 J41 J42 J43 J44 J45 J46 J47 J48 J49 J50 J51 J52 J53 J54 J55 J56 J57 J58 ' +
    'J59 J60 J61 J62 J63 J64 J65 J66 J67 J68 J69 J70 J71 J72 J73 J74 J75 J76 J78 J79 J81 J82 J83 J84 J85 J87 J90 J91 ' +
    'J92 J93 J95 J96 J97 J98 J99 JE JK JM JNT JOU JPS JWL K1 K10 K11 K12 K13 K14 K15 K16 K17 K18 K19 K2 K20 K21 K22 ' +
    'K23 K26 K27 K28 K3 K30 K31 K32 K33 K34 K35 K36 K37 K38 K39 K40 K41 K42 K43 K45 K46 K47 K48 K49 K50 K51 K52 K53 ' +
    'K54 K55 K58 K59 K6 K60 K61 K62 K63 K64 K65 K66 K67 K68 K69 K70 K71 K73 K74 K75 K76 K77 K78 K79 K80 K81 K82 K83 ' +
    'K84 K85 K86 K87 K88 K89 K90 K91 K92 K93 K94 K95 K96 K97 K98 K99 KA KAT KB KBA KCC KDW KEL KGM KGS KHY KHZ KI KIC ' +
    'KIP KJ KJO KL KLK KLX KMA KMH KMK KMQ KMT KNI KNM KNS KNT KO KPA KPH KPO KPP KR KSD KSH KT KTN KUR KVA KVR KVT ' +
    'KW KWH KWN KWO KWS KWT KWY KX L10 L11 L12 L13 L14 L15 L16 L17 L18 L19 L2 L20 L21 L23 L24 L25 L26 L27 L28 L29 L30 ' +
    'L31 L32 L33 L34 L35 L36 L37 L38 L39 L40 L41 L42 L43 L44 L45 L46 L47 L48 L49 L50 L51 L52 L53 L54 L55 L56 L57 L58 ' +
    'L59 L60 L63 L64 L65 L66 L67 L68 L69 L70 L71 L72 L73 L74 L75 L76 L77 L78 L79 L80 L81 L82 L83 L84 L85 L86 L87 L88 ' +
    'L89 L90 L91 L92 L93 L94 L95 L96 L98 L99 LA LAC LBR LBT LD LEF LF LH LK LM LN LO LP LPA LR LS LTN LTR LUB LUM LUX ' +
    'LY M1 M10 M11 M12 M13 M14 M15 M16 M17 M18 M19 M20 M21 M22 M23 M24 M25 M26 M27 M29 M30 M31 M32 M33 M34 M35 M36 ' +
    'M37 M38 M39 M4 M40 M41 M42 M43 M44 M45 M46 M47 M48 M49 M5 M50 M51 M52 M53 M55 M56 M57 M58 M59 M60 M61 M62 M63 ' +
    'M64 M65 M66 M67 M68 M69 M7 M70 M71 M72 M73 M74 M75 M76 M77 M78 M79 M80 M81 M82 M83 M84 M85 M86 M87 M88 M89 M9 ' +
    'M90 M91 M92 M93 M94 M95 M96 M97 M98 M99 MAH MAL MAM MAR MAW MBE MBF MBR MC MCU MD MGM MHZ MIK MIL MIN MIO MIU ' +
    'MKD MKM MKW MLD MLT MMK MMQ MMT MND MNJ MON MPA MQD MQH MQM MQS MQW MRD MRM MRW MSK MTK MTQ MTR MTS MTZ MVA MWH ' +
    'N1 N10 N11 N12 N13 N14 N15 N16 N17 N18 N19 N20 N21 N22 N23 N24 N25 N26 N27 N28 N29 N3 N30 N31 N32 N33 N34 N35 ' +
    'N36 N37 N38 N39 N40 N41 N42 N43 N44 N45 N46 N47 N48 N49 N50 N51 N52 N53 N54 N55 N56 N57 N58 N59 N60 N61 N62 N63 ' +
    'N64 N65 N66 N67 N68 N69 N70 N71 N72 N73 N74 N75 N76 N77 N78 N79 N80 N81 N82 N83 N84 N85 N86 N87 N88 N89 N90 N91 ' +
    'N92 N93 N94 N95 N96 N97 N98 N99 NA NAR NCL NEW NF NIL NIU NL NM3 NMI NMP NPT NT NTU NU NX OA ODE ODG ODK ODM OHM ' +
    'ON ONZ OPM OT OZA OZI P1 P10 P11 P12 P13 P14 P15 P16 P17 P18 P19 P2 P20 P21 P22 P23 P24 P25 P26 P27 P28 P29 P30 ' +
    'P31 P32 P33 P34 P35 P36 P37 P38 P39 P40 P41 P42 P43 P44 P45 P46 P47 P48 P49 P5 P50 P51 P52 P53 P54 P55 P56 P57 ' +
    'P58 P59 P60 P61 P62 P63 P64 P65 P66 P67 P68 P69 P70 P71 P72 P73 P74 P75 P76 P77 P78 P79 P80 P81 P82 P83 P84 P85 ' +
    'P86 P87 P88 P89 P90 P91 P92 P93 P94 P95 P96 P97 P98 P99 PAL PD PFL PGL PI PLA PO PQ PR PS PTD PTI PTL PTN Q10 ' +
    'Q11 Q12 Q13 Q14 Q15 Q16 Q17 Q18 Q19 Q20 Q21 Q22 Q23 Q24 Q25 Q26 Q27 Q28 Q29 Q30 Q31 Q32 Q33 Q34 Q35 Q36 Q37 Q38 ' +
    'Q39 Q40 Q41 Q42 Q3 QA QAN QB QR QTD QTI QTL QTR R1 R9 RH RM ROM RP RPM RPS RT S3 S4 SAN SCO SCR SEC SET SG SIE ' +
    'SM3 SMI SQ SQR SR STC STI STK STL STN STW SW SX SYR T0 T3 TAH TAN TI TIC TIP TKM TMS TNE TP TPI TPR TQD TRL TST ' +
    'TTS U1 U2 UB UC VA VLT VP W2 WA WB WCD WE WEB WEE WG WHR WM WSD WTT X1 YDK YDQ YRD Z11 Z9 ZP ZZ X1A X1B X1D X1F ' +
    'X1G X1W X2C X3A X3H X43 X44 X4A X4B X4C X4D X4F X4G X4H X5H X5L X5M X6H X6P X7A X7B X8A X8B X8C XAA XAB XAC XAD ' +
    'XAE XAF XAG XAH XAI XAJ XAL XAM XAP XAT XAV XB4 XBA XBB XBC XBD XBE XBF XBG XBH XBI XBJ XBK XBL XBM XBN XBO XBP ' +
    'XBQ XBR XBS XBT XBU XBV XBW XBX XBY XBZ XCA XCB XCC XCD XCE XCF XCG XCH XCI XCJ XCK XCL XCM XCN XCO XCP XCQ XCR ' +
    'XCS XCT XCU XCV XCW XCX XCY XCZ XDA XDB XDC XDG XDH XDI XDJ XDK XDL XDM XDN XDP XDR XDS XDT XDU XDV XDW XDX XDY ' +
    'XEC XED XEE XEF XEG XEH XEI XEN XFB XFC XFD XFE XFI XFL XFO XFP XFR XFT XFW XFX XGB XGI XGL XGR XGU XGY XGZ XHA ' +
    'XHB XHC XHG XHN XHR XIA XIB XIC XID XIE XIF XIG XIH XIK XIL XIN XIZ XJB XJC XJG XJR XJT XJY XKG XKI XLE XLG XLT ' +
    'XLU XLV XLZ XMA XMB XMC XME XMR XMS XMT XMW XMX XNA XNE XNF XNG XNS XNT XNU XNV XO1 XO2 XO3 XO4 XO5 XO6 XO7 XO8 ' +
    'XO9 XOA XOB XOC XOD XOE XOF XOG XOH XOI XOK XOJ XOL XOM XON XOP XOQ XOR XOS XOV XOW XOT XOU XOX XOY XOZ XP1 XP2 ' +
    'XP3 XP4 XPA XPB XPC XPD XPE XPF XPG XPH XPI XPJ XPK XPL XPN XPO XPP XPR XPT XPU XPV XPX XPY XPZ XQA XQB XQC XQD ' +
    'XQF XQG XQH XQJ XQK XQL XQM XQN XQP XQQ XQR XQS XRD XRG XRJ XRK XRL XRO XRT XRZ XSA XSB XSC XSD XSE XSH XSI XSK ' +
    'XSL XSM XSO XSP XSS XST XSU XSV XSW XSX XSY XSZ XT1 XTB XTC XTD XTE XTG XTI XTK XTL XTN XTO XTR XTS XTT XTU XTV ' +
    'XTW XTY XTZ XUC XUN XVA XVG XVI XVK XVL XVO XVP XVQ XVN XVR XVS XVY XWA XWB XWC XWD XWF XWG XWH XWJ XWK XWL XWM ' +
    'XWN XWP XWQ XWR XWS XWT XWU XWV XWW XWX XWY XWZ XXA XXB XXC XXD XXF XXG XXH XXJ XXK XYA XYB XYC XYD XYF XYG XYH ' +
    'XYJ XYK XYL XYM XYN XYP XYQ XYR XYS XYT XYV XYW XYX XYY XYZ XZA XZB XZC XZD XZF XZG XZH XZJ XZK XZL XZM XZN XZP ' +
    'XZQ XZR XZS XZT XZU XZV XZW XZX XZY XZZ').split(' ');
  var UNIT_CODE_SET = {};
  UNIT_CODES.forEach(function (c) { UNIT_CODE_SET[c] = true; });
  var COUNTRY_CODES = ('AD AE AF AG AI AL AM AO AQ AR AS AT AU AW AX AZ BA BB BD BE BF BG BH BI BJ BL BM BN BO BQ BR BS BT ' +
    'BV BW BY BZ CA CC CD CF CG CH CI CK CL CM CN CO CR CU CV CW CX CY CZ DE DJ DK DM DO DZ EC EE EG EH ER ES ET FI FJ FK ' +
    'FM FO FR GA GB GD GE GF GG GH GI GL GM GN GP GQ GR GS GT GU GW GY HK HM HN HR HT HU ID IE IL IM IN IO IQ IR IS IT JE ' +
    'JM JO JP KE KG KH KI KM KN KP KR KW KY KZ LA LB LC LI LK LR LS LT LU LV LY MA MC MD ME MF MG MH MK ML MM MN MO MP MQ ' +
    'MR MS MT MU MV MW MX MY MZ NA NC NE NF NG NI NL NO NP NR NU NZ OM PA PE PF PG PH PK PL PM PN PR PS PT PW PY QA RE RO ' +
    'RS RU RW SA SB SC SD SE SG SH SI SJ SK SL SM SN SO SR SS ST SV SX SY SZ TC TD TF TG TH TJ TK TL TM TN TO TR TT TV TW ' +
    'TZ UA UG UM US UY UZ VA VC VE VG VI VN VU WF WS YE YT ZA ZM ZW XI 1A').split(' ');
  var COUNTRY_ALIASES = {
    'DEUTSCHLAND': 'DE', 'GERMANY': 'DE', 'ÖSTERREICH': 'AT', 'OESTERREICH': 'AT', 'AUSTRIA': 'AT',
    'SCHWEIZ': 'CH', 'SWITZERLAND': 'CH', 'FRANKREICH': 'FR', 'FRANCE': 'FR', 'NIEDERLANDE': 'NL',
    'NETHERLANDS': 'NL', 'BELGIEN': 'BE', 'BELGIUM': 'BE', 'LUXEMBURG': 'LU', 'LUXEMBOURG': 'LU',
    'ITALIEN': 'IT', 'ITALY': 'IT', 'SPANIEN': 'ES', 'SPAIN': 'ES', 'POLEN': 'PL', 'POLAND': 'PL',
    'TSCHECHIEN': 'CZ', 'CZECHIA': 'CZ', 'DÄNEMARK': 'DK', 'DENMARK': 'DK', 'LITAUEN': 'LT', 'LITHUANIA': 'LT',
    'GROSSBRITANNIEN': 'GB', 'GROßBRITANNIEN': 'GB', 'VEREINIGTES KÖNIGREICH': 'GB', 'UNITED KINGDOM': 'GB',
    'USA': 'US', 'VEREINIGTE STAATEN': 'US', 'UNITED STATES': 'US', 'SCHWEDEN': 'SE', 'SWEDEN': 'SE'
  };
  var EU_VAT_PREFIXES = ['AT', 'BE', 'BG', 'CY', 'CZ', 'DE', 'DK', 'EE', 'EL', 'ES', 'FI', 'FR', 'HR', 'HU', 'IE', 'IT',
    'LT', 'LU', 'LV', 'MT', 'NL', 'PL', 'PT', 'RO', 'SE', 'SI', 'SK', 'XI'];
  var IBAN_LENGTHS = { AD: 24, AE: 23, AL: 28, AT: 20, AZ: 28, BA: 20, BE: 16, BG: 22, BH: 22, BR: 29, BY: 28, CH: 21,
    CR: 22, CY: 28, CZ: 24, DE: 22, DK: 18, DO: 28, EE: 20, EG: 29, ES: 24, FI: 18, FO: 18, FR: 27, GB: 22, GE: 22,
    GI: 23, GL: 18, GR: 27, GT: 28, HR: 21, HU: 28, IE: 22, IL: 23, IQ: 23, IS: 26, IT: 27, JO: 30, KW: 30, KZ: 20,
    LB: 28, LC: 32, LI: 21, LT: 20, LU: 20, LV: 21, MC: 27, MD: 24, ME: 22, MK: 19, MR: 27, MT: 31, MU: 30, NL: 18,
    NO: 15, PK: 24, PL: 28, PS: 29, PT: 25, QA: 29, RO: 24, RS: 22, SA: 24, SC: 31, SE: 24, SI: 19, SK: 24, SM: 27,
    ST: 25, SV: 28, TL: 23, TN: 24, TR: 26, UA: 29, VA: 22, VG: 24, XK: 20 };

  function mod97(numericString) {
    var r = 0;
    for (var k = 0; k < numericString.length; k++) {
      r = (r * 10 + (numericString.charCodeAt(k) - 48)) % 97;
    }
    return r;
  }
  function lettersToDigits(s) {
    var out = '';
    for (var k = 0; k < s.length; k++) {
      var c = s.charAt(k);
      out += (c >= 'A' && c <= 'Z') ? String(c.charCodeAt(0) - 55) : c;
    }
    return out;
  }
  function normalizeIban(v) { return text(v).replace(/[\s-]/g, '').toUpperCase(); }
  function isValidIban(v) {
    var s = normalizeIban(v);
    if (!/^[A-Z]{2}\d{2}[A-Z0-9]{8,30}$/.test(s)) { return false; }
    var len = IBAN_LENGTHS[s.slice(0, 2)];
    if (len && s.length !== len) { return false; }
    return mod97(lettersToDigits(s.slice(4) + s.slice(0, 4))) === 1;
  }
  function normalizeVatId(v) {
    var s = text(v).replace(/[\s.-]/g, '').toUpperCase();
    // The Swiss UID as officially written, CHE-123.456.789 MWST (or TVA, IVA), becomes CHE123456789 (INV-10).
    var ch = /^(CHE\d{9})(MWST|TVA|IVA)$/.exec(s);
    return ch ? ch[1] : s;
  }
  function isPlausibleVatId(v) {
    var s = normalizeVatId(v);
    if (!/^[A-Z]{2}[A-Z0-9+*]{2,13}$/.test(s)) { return false; }
    var p = s.slice(0, 2);
    if (p === 'DE') { return /^DE\d{9}$/.test(s); }
    if (p === 'AT') { return /^ATU\d{8}$/.test(s); }
    return p === 'EL' || COUNTRY_CODES.indexOf(p) >= 0;
  }
  // Leitweg-ID (German public sector routing id): coarse-fine-check digits, ISO 7064 MOD 97-10.
  function leitwegShape(v) {
    return /^\d{2,12}(-[A-Za-z0-9]{1,30})?-\d{2}$/.test(text(v));
  }
  function isValidLeitwegId(v) {
    var s = text(v).toUpperCase();
    if (!leitwegShape(s)) { return false; }
    var parts = s.split('-');
    var check = parts.pop();
    return mod97(lettersToDigits(parts.join('') + check)) === 1;
  }
  function normalizeCountry(v) {
    var s = text(v).toUpperCase();
    if (s === '') { return null; }
    var m = /\(([A-Z0-9]{2})\)$/.exec(s);
    if (m) { s = m[1]; }
    if (COUNTRY_ALIASES[s]) { return COUNTRY_ALIASES[s]; }
    return COUNTRY_CODES.indexOf(s) >= 0 ? s : null;
  }
  function normalizeUnit(v, extraCodes) {
    var s = text(v);
    if (s === '') { return 'H87'; }
    var m = /\(([A-Za-z0-9]{2,3})\)\s*$/.exec(s);
    // A trailing dot is dropped and the short labels the PDF prints are accepted: "Stk.", "Std." (INV-10).
    var up = (m ? m[1] : s).toUpperCase().replace(/\.$/, '');
    for (var k = 0; k < UNITS.length; k++) {
      var u = UNITS[k];
      if (u.code === up || u.aliases.indexOf(up) >= 0 || u.de.toUpperCase() === up || u.en.toUpperCase() === up ||
          u.sde.toUpperCase().replace(/\.$/, '') === up || u.sen.toUpperCase() === up) {
        return u.code;
      }
    }
    // Any other code only when it is an official one: a guessed code like (KAR) passed the check in 1.0.0 and then
    // made the library refuse the invoice with "validation failed".
    if (UNIT_CODE_SET[up] || (extraCodes && extraCodes.indexOf(up) >= 0)) { return up; }
    return null;
  }
  function unitLabel(code, lang, short) {
    for (var k = 0; k < UNITS.length; k++) {
      if (UNITS[k].code === code) { return UNITS[k][(short ? 's' : '') + (lang === 'en' ? 'en' : 'de')]; }
    }
    return code;
  }
  function normalizeCategory(v) {
    var s = text(v).toUpperCase();
    if (s === '') { return 'S'; }
    var first = s.split(/[\s:(]/)[0];
    return VAT_CATEGORIES[first] ? first : null;
  }
  function normalizeDocType(v) {
    var s = text(v);
    if (s === '') { return '380'; }
    var m = /^(\d{3})/.exec(s);
    if (!m) { return null; }
    for (var k = 0; k < DOC_TYPES.length; k++) { if (DOC_TYPES[k].code === m[1]) { return m[1]; } }
    return null;
  }
  // The XRechnung pattern (BR-DE-28), plus no dot at the start or right before the @.
  function isValidEmail(v) {
    var s = text(v);
    return /^[^@\s]+@([^@.\s]+\.)+[^@.\s]+$/.test(s) && !/^\./.test(s) && !/\.@/.test(s);
  }
  function countDigits(v) { var m = text(v).match(/\d/g); return m ? m.length : 0; }

  // ---------------------------------------------------------------------------
  // Template definition. Column A holds hidden keys, B labels, C values.
  // ---------------------------------------------------------------------------
  var TEMPLATE_MARKER = 'einvoice-template:1';
  var TEMPLATE_SHEET_NAME = 'E-Rechnung';
  var LINE_COLUMNS = [
    { key: 'description', col: 2, de: 'Beschreibung', en: 'Description', required: true, width: 260 },
    { key: 'quantity', col: 3, de: 'Menge', en: 'Quantity', required: true, width: 150, format: '#,##0.###' },
    { key: 'unit', col: 4, de: 'Einheit', en: 'Unit', width: 100 },
    { key: 'price', col: 5, de: 'Einzelpreis netto', en: 'Unit price net', required: true, width: 100, format: '#,##0.00##' },
    { key: 'discount', col: 6, de: 'Rabatt %', en: 'Discount %', width: 70, format: '0.##' },
    { key: 'category', col: 7, de: 'Steuerart', en: 'VAT type', width: 170 },
    { key: 'rate', col: 8, de: 'USt %', en: 'VAT %', width: 55, format: '0.##' },
    { key: 'net', col: 9, de: 'Netto', en: 'Net', width: 95, format: '#,##0.00' },
    { key: 'itemNo', col: 10, de: 'Art.-Nr.', en: 'Item no.', width: 90 }
  ];
  var LINE_COL = {};
  LINE_COLUMNS.forEach(function (c) { LINE_COL[c.key] = c.col; });
  // Line columns formatted as plain text in the template (description, unit, VAT type, item no.).
  var LINE_TEXT_COLUMNS = LINE_COLUMNS.filter(function (c) { return !c.format && c.key !== 'net'; }).map(function (c) { return c.col; });

  function docTypeChoices() { return DOC_TYPES.map(function (t) { return t.code + ' ' + t.de + ' / ' + t.en; }); }
  function categoryChoices() {
    return VAT_CATEGORY_ORDER.map(function (k) { return k + ' ' + VAT_CATEGORIES[k].pick; });
  }
  function unitChoices() { return UNITS.map(function (u) { return u.de + ' (' + u.code + ')'; }); }

  var SECTIONS = [
    { key: 'doc', de: 'Rechnung', en: 'Invoice', fields: [
      { key: 'doc.type', de: 'Dokumentart', en: 'Document type', type: 'choice', choices: 'docTypes',
        def: '380 Rechnung / Invoice', required: true },
      { key: 'doc.number', de: 'Rechnungsnummer', en: 'Invoice number', required: true },
      { key: 'doc.date', de: 'Rechnungsdatum', en: 'Invoice date', type: 'date', required: true },
      { key: 'doc.deliveryDate', de: 'Leistungsdatum', en: 'Delivery date', type: 'date',
        hintDe: 'Tag der Lieferung oder Leistung. Alternativ den Leistungszeitraum ausfüllen.',
        hintEn: 'Date of delivery or service. Alternatively fill in the service period.' },
      { key: 'doc.periodStart', de: 'Leistungszeitraum von', en: 'Service period from', type: 'date' },
      { key: 'doc.periodEnd', de: 'Leistungszeitraum bis', en: 'Service period to', type: 'date' },
      { key: 'doc.dueDate', de: 'Fällig am', en: 'Due date', type: 'date' },
      { key: 'doc.buyerReference', de: 'Käuferreferenz oder Leitweg-ID', en: 'Buyer reference or Leitweg-ID', required: true,
        hintDe: 'Bei Behörden die Leitweg-ID. Bei Firmen eine Referenz des Kunden, z. B. seine Bestellnummer oder Kostenstelle.',
        hintEn: 'For public bodies the Leitweg-ID. For companies a reference of the customer, e.g. order number or cost centre.' },
      { key: 'doc.orderReference', de: 'Bestellnummer des Kunden', en: 'Customer order number' },
      { key: 'doc.precedingInvoice', de: 'Bezug auf Rechnung Nr.', en: 'Corrects invoice no.',
        hintDe: 'Nur bei Rechnungskorrektur: Nummer der ursprünglichen Rechnung.',
        hintEn: 'Credit notes only: number of the original invoice.' },
      { key: 'doc.precedingInvoiceDate', de: 'Datum dieser Rechnung', en: 'Date of that invoice', type: 'date' },
      { key: 'doc.note', de: 'Bemerkung', en: 'Note' }
    ] },
    { key: 'seller', de: 'Ihre Firma (Verkäufer)', en: 'Your company (seller)', fields: [
      { key: 'seller.name', de: 'Firmenname', en: 'Company name', required: true },
      { key: 'seller.street', de: 'Straße und Hausnummer', en: 'Street and number', required: true },
      { key: 'seller.zip', de: 'PLZ', en: 'Postcode', required: true },
      { key: 'seller.city', de: 'Ort', en: 'City', required: true },
      { key: 'seller.country', de: 'Land (z. B. DE)', en: 'Country code (e.g. DE)', def: 'DE', required: true },
      { key: 'seller.vatId', de: 'USt-IdNr.', en: 'VAT ID',
        hintDe: 'USt-IdNr. oder Steuernummer angeben.', hintEn: 'Enter the VAT ID or the tax number.' },
      { key: 'seller.taxNumber', de: 'Steuernummer', en: 'Tax number' },
      { key: 'seller.registration', de: 'Handelsregisternummer', en: 'Company registration no.' },
      { key: 'seller.contactName', de: 'Ansprechpartner', en: 'Contact person', required: true },
      { key: 'seller.phone', de: 'Telefon', en: 'Phone', required: true },
      { key: 'seller.email', de: 'E-Mail', en: 'Email', required: true },
      { key: 'seller.iban', de: 'IBAN', en: 'IBAN', required: true },
      { key: 'seller.bic', de: 'BIC', en: 'BIC' },
      { key: 'seller.accountHolder', de: 'Kontoinhaber', en: 'Account holder' }
    ] },
    { key: 'buyer', de: 'Kunde (Käufer)', en: 'Customer (buyer)', fields: [
      { key: 'buyer.name', de: 'Firmenname', en: 'Company name', required: true },
      { key: 'buyer.contactName', de: 'Ansprechpartner', en: 'Contact person' },
      { key: 'buyer.street', de: 'Straße und Hausnummer', en: 'Street and number', required: true },
      { key: 'buyer.zip', de: 'PLZ', en: 'Postcode', required: true },
      { key: 'buyer.city', de: 'Ort', en: 'City', required: true },
      { key: 'buyer.country', de: 'Land (z. B. DE)', en: 'Country code (e.g. DE)', def: 'DE', required: true },
      { key: 'buyer.vatId', de: 'USt-IdNr. des Kunden', en: 'Customer VAT ID',
        hintDe: 'Pflicht bei Reverse Charge und innergemeinschaftlicher Lieferung.',
        hintEn: 'Required for reverse charge and intra-EU supplies.' },
      { key: 'buyer.email', de: 'E-Mail für E-Rechnungen', en: 'Email for e-invoices',
        hintDe: 'Elektronische Adresse des Kunden. Nicht nötig, wenn eine Leitweg-ID angegeben ist.',
        hintEn: 'Electronic address of the customer. Not needed when a Leitweg-ID is given.' },
      { key: 'buyer.customerNumber', de: 'Kundennummer', en: 'Customer number' }
    ] },
    { key: 'lines', de: 'Positionen', en: 'Line items', lines: true },
    { key: 'totals', de: 'Abschluss', en: 'Totals', fields: [
      { key: 'totals.discountPercent', de: 'Rabatt auf alle Positionen %', en: 'Discount on all items %', type: 'number' },
      { key: 'totals.discountReason', de: 'Grund des Rabatts', en: 'Discount reason', def: 'Rabatt' },
      { key: 'totals.prepaid', de: 'Bereits bezahlt (EUR)', en: 'Already paid (EUR)', type: 'number',
        hintDe: 'Bei einer Schlussrechnung die Anzahlungen brutto. Bei einem Steuersatz wird die darin enthaltene USt automatisch ausgewiesen.',
        hintEn: 'For a final invoice the advance payments including VAT. With one VAT rate the VAT contained is stated automatically.' },
      { key: 'tax.exemptionReason', de: 'Grund der Steuerbefreiung', en: 'Reason for VAT exemption',
        hintDe: 'Pflicht bei Steuerart E (steuerfrei). Wird auch bei Steuerart O verwendet, falls angegeben.',
        hintEn: 'Required for VAT type E (exempt). Also used for VAT type O when given.' }
    ] },
    { key: 'payment', de: 'Zahlung', en: 'Payment', fields: [
      { key: 'payment.terms', de: 'Zahlungsbedingungen', en: 'Payment terms', def: 'Zahlbar innerhalb von 14 Tagen ohne Abzug.' },
      { key: 'payment.skontoPercent', de: 'Skonto %', en: 'Cash discount %', type: 'number' },
      { key: 'payment.skontoDays', de: 'Skonto Tage', en: 'Cash discount days', type: 'number' },
      { key: 'payment.reference', de: 'Verwendungszweck', en: 'Payment reference',
        hintDe: 'Leer lassen für die Rechnungsnummer.', hintEn: 'Leave empty to use the invoice number.' }
    ] }
  ];
  var FIELD_BY_KEY = {};
  SECTIONS.forEach(function (s) { (s.fields || []).forEach(function (f) { FIELD_BY_KEY[f.key] = f; }); });

  // Fields formatted as plain text in the template (everything except dates and numbers).
  function isTextField(key) {
    var f = FIELD_BY_KEY[key];
    return !!f && f.type !== 'date' && f.type !== 'number';
  }
  function fieldLabel(key, lang) {
    var f = FIELD_BY_KEY[key];
    if (!f) {
      for (var k = 0; k < LINE_COLUMNS.length; k++) {
        if ('line.' + LINE_COLUMNS[k].key === key) { return LINE_COLUMNS[k][lang === 'en' ? 'en' : 'de']; }
      }
      return key;
    }
    return f[lang === 'en' ? 'en' : 'de'];
  }
  function colLetter(col) { // 1-based
    var s = '';
    while (col > 0) { var m = (col - 1) % 26; s = String.fromCharCode(65 + m) + s; col = Math.floor((col - 1) / 26); }
    return s;
  }
  function a1(row, col) { return colLetter(col) + row; }

  // Builds the template as plain data. The Apps Script server applies it to a sheet.
  // opts: {lineRows: number of empty line rows, example: object from exampleInvoice() or null}
  function buildTemplatePlan(opts) {
    opts = opts || {};
    var lineRows = opts.lineRows || 10;
    var width = 10;
    var values = [];
    var rowTypes = [];
    var notes = [];
    var validations = [];
    var numberFormats = [];
    var inputRanges = [];
    var fieldRows = {};
    function blankRow() { var r = []; for (var k = 0; k < width; k++) { r.push(''); } return r; }
    function push(row, type) { values.push(row); rowTypes.push(type); return values.length; }

    var r0 = blankRow(); r0[0] = TEMPLATE_MARKER; r0[1] = 'E-Rechnung (XRechnung / ZUGFeRD)';
    push(r0, 'title');
    var r1 = blankRow();
    r1[1] = 'Gelbe Felder ausfüllen, Pflichtfelder haben einen Stern (*). Danach in der Seitenleiste auf "Prüfen" klicken. / ' +
      'Fill in the yellow fields; required fields have a star (*). Then click "Check" in the sidebar.';
    push(r1, 'subtitle');
    push(blankRow(), 'blank');

    var linesFirst = null;
    var linesLast = null;
    SECTIONS.forEach(function (section) {
      var h = blankRow();
      h[0] = '#section:' + section.key;
      h[1] = section.de + ' / ' + section.en;
      if (section.lines) {
        h[0] = '#lines';
        LINE_COLUMNS.forEach(function (c) {
          h[c.col - 1] = c.de + ' / ' + c.en + (c.required ? ' *' : '');
        });
        var hr = push(h, 'linesHeader');
        notes.push({ a1: a1(hr, 2), text: 'Weitere Positionen: Zeilen oberhalb von "Summe Positionen" einfügen und die Formel der Spalte Netto ' +
          'aus der Zeile darüber kopieren. / More items: insert rows above "Items net total" and copy the formula of the Net column ' +
          'from the row above.' });
        linesFirst = hr + 1;
        for (var n = 0; n < lineRows; n++) {
          var lr = blankRow();
          lr[LINE_COL.unit - 1] = unitChoices()[0];
          lr[LINE_COL.category - 1] = categoryChoices()[0];
          lr[LINE_COL.rate - 1] = 19;
          var rowNo = values.length + 1;
          lr[LINE_COL.net - 1] = lineNetFormula(rowNo);
          push(lr, 'line');
        }
        linesLast = values.length;
        var end = blankRow();
        end[0] = '#lines-end';
        end[1] = 'Summe Positionen netto / Items net total';
        end[LINE_COL.net - 1] = linesTotalFormula(hr);
        push(end, 'linesEnd');
        push(blankRow(), 'blank');
        return;
      }
      push(h, 'section');
      section.fields.forEach(function (f) {
        var row = blankRow();
        row[0] = f.key;
        row[1] = f.de + ' / ' + f.en + (f.required ? ' *' : '');
        row[2] = f.def !== undefined ? f.def : '';
        var rn = push(row, 'field');
        fieldRows[f.key] = rn;
        inputRanges.push(a1(rn, 3));
        if (f.hintDe) { notes.push({ a1: a1(rn, 3), text: f.hintDe + ' / ' + f.hintEn }); }
        if (f.type === 'date') {
          validations.push({ a1: a1(rn, 3), type: 'date' });
          numberFormats.push({ a1: a1(rn, 3), format: 'dd.mm.yyyy' });
        }
        if (f.type === 'choice' && f.choices === 'docTypes') {
          validations.push({ a1: a1(rn, 3), type: 'list', values: docTypeChoices(), strict: true });
        }
        if (f.type === 'number') { numberFormats.push({ a1: a1(rn, 3), format: '#,##0.00' }); }
        // Plain text keeps leading zeros (postcode 04109, phone 030...) and stops date guessing.
        if (f.type !== 'date' && f.type !== 'number') { numberFormats.push({ a1: a1(rn, 3), format: '@' }); }
      });
      push(blankRow(), 'blank');
    });

    var linesRange = function (col) { return a1(linesFirst, col) + ':' + a1(linesLast, col); };
    inputRanges.push(a1(linesFirst, 2) + ':' + a1(linesLast, LINE_COL.rate));
    inputRanges.push(linesRange(LINE_COL.itemNo));
    validations.push({ a1: linesRange(LINE_COL.unit), type: 'list', values: unitChoices(), strict: false });
    validations.push({ a1: linesRange(LINE_COL.category), type: 'list', values: categoryChoices(), strict: true });
    LINE_COLUMNS.forEach(function (c) {
      var fmt = c.format || (c.key === 'net' ? null : '@');
      if (fmt) { numberFormats.push({ a1: linesRange(c.col), format: fmt }); }
    });
    numberFormats.push({ a1: a1(linesLast + 1, LINE_COL.net), format: '#,##0.00' });

    var plan = {
      sheetName: TEMPLATE_SHEET_NAME,
      width: width,
      values: values,
      rowTypes: rowTypes,
      notes: notes,
      validations: validations,
      numberFormats: numberFormats,
      inputRanges: inputRanges,
      fieldRows: fieldRows,
      linesFirstRow: linesFirst,
      linesLastRow: linesLast,
      columnWidths: [40].concat(LINE_COLUMNS.map(function (c) { return c.width; })),
      hiddenColumns: [1]
    };
    if (opts.example) { fillTemplate(plan, opts.example); }
    return plan;
  }
  // Sum from the header row to the row above the total, so rows inserted anywhere in between count (SRV-10).
  function linesTotalFormula(headerRow) {
    var c = colLetter(LINE_COL.net);
    return '=SUM(' + c + headerRow + ':INDEX(' + c + ':' + c + ',ROW()-1))';
  }
  function lineNetFormula(rowNo) {
    var q = a1(rowNo, LINE_COL.quantity);
    var p = a1(rowNo, LINE_COL.price);
    var d = a1(rowNo, LINE_COL.discount);
    var base = 'ROUND(' + q + '*' + p + ',2)';
    return '=IF(OR(' + q + '="",' + p + '=""),"",' + base + '-ROUND(' + base + '*N(' + d + ')/100,2))';
  }
  // Writes data {fields: {key: value}, lines: [{description, quantity, unit, price, discount, category, rate, itemNo}]}
  // into a plan, inserting line rows when needed. Mutates and returns the plan.
  function fillTemplate(plan, data) {
    var fields = data.fields || {};
    Object.keys(fields).forEach(function (key) {
      var r = plan.fieldRows[key];
      if (!r) { throw new Error('Unknown template field: ' + key); }
      plan.values[r - 1][2] = fields[key];
    });
    var lines = data.lines || [];
    var capacity = plan.linesLastRow - plan.linesFirstRow + 1;
    if (lines.length > capacity) {
      var extra = lines.length - capacity;
      for (var e = 0; e < extra; e++) {
        var copy = plan.values[plan.linesLastRow - 1].slice();
        plan.values.splice(plan.linesLastRow, 0, copy);
        plan.rowTypes.splice(plan.linesLastRow, 0, 'line');
        plan.linesLastRow += 1;
      }
      // Re-number formulas and field rows after the insert.
      for (var r = plan.linesFirstRow; r <= plan.linesLastRow; r++) {
        plan.values[r - 1][LINE_COL.net - 1] = lineNetFormula(r);
      }
      plan.values[plan.linesLastRow][LINE_COL.net - 1] = linesTotalFormula(plan.linesFirstRow - 1);
      Object.keys(plan.fieldRows).forEach(function (k) {
        if (plan.fieldRows[k] > plan.linesFirstRow) { plan.fieldRows[k] += extra; }
      });
    }
    lines.forEach(function (line, idx) {
      var row = plan.values[plan.linesFirstRow - 1 + idx];
      LINE_COLUMNS.forEach(function (c) {
        if (c.key === 'net') { return; }
        if (line[c.key] !== undefined && line[c.key] !== null) { row[c.col - 1] = line[c.key]; }
      });
    });
    return plan;
  }

  // ---------------------------------------------------------------------------
  // Messages.
  // ---------------------------------------------------------------------------
  var MESSAGES = {
    NO_TEMPLATE: { de: 'Kein Rechnungsblatt gefunden. Bitte zuerst "Leere Vorlage anlegen" oder "Beispiel anlegen" ausführen.',
      en: 'No invoice sheet found. Please use "Create empty template" or "Create example" first.' },
    CHOOSE_SHEET: { de: 'Diese Tabelle hat mehrere Rechnungsblätter. Bitte das Blatt öffnen, für das die E-Rechnung erstellt werden soll, und erneut prüfen.',
      en: 'This spreadsheet has several invoice sheets. Please open the sheet you want to create the e-invoice for and check again.' },
    REQUIRED: { de: '{label} fehlt.', en: '{label} is missing.' },
    BAD_DATE: { de: '{label}: kein gültiges Datum. Bitte als TT.MM.JJJJ eingeben.', en: '{label}: not a valid date. Please use DD.MM.YYYY.' },
    BAD_NUMBER: { de: '{label}: keine gültige Zahl.', en: '{label}: not a valid number.' },
    BAD_DOC_TYPE: { de: 'Dokumentart unbekannt. Bitte aus der Liste wählen.', en: 'Unknown document type. Please pick one from the list.' },
    BAD_COUNTRY: { de: '{label}: bitte den zweistelligen Ländercode angeben, z. B. DE.', en: '{label}: please enter the two letter country code, e.g. DE.' },
    BAD_EMAIL: { de: '{label}: keine gültige E-Mail-Adresse.', en: '{label}: not a valid email address.' },
    BAD_PHONE: { de: '{label}: mindestens 3 Ziffern nötig.', en: '{label}: at least 3 digits needed.' },
    BAD_IBAN: { de: 'IBAN ist ungültig (Prüfsumme oder Länge stimmt nicht).', en: 'IBAN is not valid (checksum or length wrong).' },
    BAD_VAT_ID: { de: '{label}: Format ungültig. Erwartet z. B. DE123456789.', en: '{label}: invalid format. Expected e.g. DE123456789.' },
    SELLER_TAX_ID: { de: 'USt-IdNr. oder Steuernummer des Verkäufers angeben.', en: 'Enter the seller VAT ID or tax number.' },
    SELLER_VAT_ID_NEEDED: { de: 'Bei Steuerart {cat} ist die USt-IdNr. des Verkäufers Pflicht.', en: 'VAT type {cat} requires the seller VAT ID.' },
    BUYER_VAT_ID_NEEDED: { de: 'Bei Steuerart {cat} ist die USt-IdNr. des Kunden Pflicht.', en: 'VAT type {cat} requires the customer VAT ID.' },
    BUYER_EMAIL_NEEDED: { de: 'E-Mail für E-Rechnungen des Kunden fehlt (nötig, wenn keine Leitweg-ID angegeben ist).',
      en: 'Customer email for e-invoices is missing (needed when no Leitweg-ID is given).' },
    NO_DELIVERY_DATE: { de: 'Leistungsdatum oder Leistungszeitraum fehlt.', en: 'Delivery date or service period is missing.' },
    PERIOD_INCOMPLETE: { de: 'Leistungszeitraum: bitte Beginn und Ende angeben.', en: 'Service period: please enter start and end.' },
    PERIOD_ORDER: { de: 'Leistungszeitraum: Ende liegt vor dem Beginn.', en: 'Service period: end is before start.' },
    DUE_BEFORE_ISSUE: { de: 'Fälligkeitsdatum liegt vor dem Rechnungsdatum.', en: 'Due date is before the invoice date.' },
    NO_LINES: { de: 'Keine Positionen. Bitte mindestens eine Position ausfüllen.', en: 'No line items. Please fill in at least one.' },
    LINE_REQUIRED: { de: 'Position {n}: {label} fehlt.', en: 'Item {n}: {label} is missing.' },
    LINE_BAD_NUMBER: { de: 'Position {n}: {label} ist keine gültige Zahl.', en: 'Item {n}: {label} is not a valid number.' },
    LINE_ZERO_QTY: { de: 'Position {n}: Menge darf nicht 0 sein.', en: 'Item {n}: quantity must not be 0.' },
    LINE_NEGATIVE_PRICE: { de: 'Position {n}: Einzelpreis darf nicht negativ sein. Für Abzüge eine negative Menge verwenden.',
      en: 'Item {n}: unit price must not be negative. Use a negative quantity for deductions.' },
    LINE_BAD_DISCOUNT: { de: 'Position {n}: Rabatt muss zwischen 0 und 100 liegen.', en: 'Item {n}: discount must be between 0 and 100.' },
    LINE_BAD_UNIT: { de: 'Position {n}: Einheit unbekannt. Bitte aus der Liste wählen oder den Code in Klammern angeben, z. B. (H87).',
      en: 'Item {n}: unknown unit. Pick one from the list or give the code in brackets, e.g. (H87).' },
    LINE_BAD_CATEGORY: { de: 'Position {n}: Steuerart unbekannt. Bitte aus der Liste wählen.', en: 'Item {n}: unknown VAT type. Please pick one from the list.' },
    LINE_RATE_NEEDED: { de: 'Position {n}: USt-Satz muss bei Regelsatz größer als 0 sein.', en: 'Item {n}: VAT rate must be above 0 for the standard rate.' },
    LINE_RATE_ZERO: { de: 'Position {n}: bei Steuerart {cat} muss der USt-Satz 0 sein.', en: 'Item {n}: VAT type {cat} needs a VAT rate of 0.' },
    LINE_NET_DIFFERS: { de: 'Position {n}: Netto im Blatt ({s}) passt nicht zu Menge mal Einzelpreis ({c}). Die E-Rechnung verwendet {c}. Bitte die Formel in der Spalte Netto wiederherstellen.',
      en: 'Item {n}: the net in the sheet ({s}) does not match quantity times unit price ({c}). The e-invoice uses {c}. Please restore the formula in the Net column.' },
    DATE_FAR: { de: 'Rechnungsdatum {d} liegt mehr als ein Jahr vom heutigen Datum entfernt. Bitte prüfen.',
      en: 'Invoice date {d} is more than a year away from today. Please check.' },
    LINE_RATE_FRACTION: { de: 'Position {n}: USt-Satz {r} % ist zu klein. Bitte den Prozentsatz als Zahl eingeben, z. B. 7 statt 0,07.',
      en: 'Item {n}: VAT rate {r} % is too small. Please enter the percentage as a number, e.g. 7 instead of 0.07.' },
    LINE_DISCOUNT_FRACTION: { de: 'Position {n}: Rabatt {d} % ist sehr klein. Falls {x} % gemeint sind, bitte {x} eingeben.',
      en: 'Item {n}: discount {d} % is very small. If you mean {x} %, please enter {x}.' },
    DISCOUNT_FRACTION: { de: 'Rabatt auf alle Positionen: {d} % ist sehr klein. Falls {x} % gemeint sind, bitte {x} eingeben.',
      en: 'Discount on all items: {d} % is very small. If you mean {x} %, please enter {x}.' },
    SKONTO_FRACTION: { de: 'Skonto: {d} % ist sehr klein. Falls {x} % gemeint sind, bitte {x} eingeben.',
      en: 'Cash discount: {d} % is very small. If you mean {x} %, please enter {x}.' },
    O_MIXED: { de: 'Steuerart O (nicht steuerbar) darf nicht mit anderen Steuerarten in einer Rechnung stehen.',
      en: 'VAT type O (not subject to VAT) cannot be combined with other VAT types in one invoice.' },
    O_VAT_IDS: { de: 'Bei Steuerart O werden die USt-IdNrn. nicht in die E-Rechnung übernommen.',
      en: 'With VAT type O the VAT IDs are left out of the e-invoice.' },
    O_SELLER_ID: { de: 'Bei Steuerart O (nicht steuerbar) wird die USt-IdNr. nicht übernommen. Bitte die Steuernummer oder die Handelsregisternummer angeben.',
      en: 'With VAT type O (not subject to VAT) the VAT ID is left out. Please enter the tax number or the company registration number.' },
    EXEMPTION_REASON_NEEDED: { de: 'Grund der Steuerbefreiung fehlt (Pflicht bei Steuerart E).', en: 'Reason for VAT exemption is missing (required for VAT type E).' },
    BAD_DISCOUNT: { de: 'Rabatt auf alle Positionen muss zwischen 0 und 100 liegen.', en: 'Discount on all items must be between 0 and 100.' },
    SKONTO_INCOMPLETE: { de: 'Skonto: bitte Prozent und Tage angeben.', en: 'Cash discount: please enter percent and days.' },
    SKONTO_BAD: { de: 'Skonto: Prozent muss zwischen 0 und 100 liegen, Tage eine ganze Zahl über 0.',
      en: 'Cash discount: percent must be between 0 and 100, days a whole number above 0.' },
    PAYMENT_TERMS_HASH: { de: 'Zahlungsbedingungen: Zeilen, die mit # beginnen, sind für Skonto reserviert und müssen genau so aussehen: #SKONTO#TAGE=14#PROZENT=2.00#. Einfacher sind die Felder "Skonto %" und "Skonto Tage".',
      en: 'Payment terms: lines starting with # are reserved for cash discounts and must look exactly like #SKONTO#TAGE=14#PROZENT=2.00#. The fields "Cash discount %" and "Cash discount days" are easier.' },
    FORMULA_ERROR: { de: '{label}: Die Zelle zeigt einen Formelfehler ({v}). Bitte die Formel korrigieren oder den Wert direkt eingeben.',
      en: '{label}: the cell shows a formula error ({v}). Please fix the formula or type the value.' },
    LINE_FORMULA_ERROR: { de: 'Position {n}: {label} zeigt einen Formelfehler ({v}). Bitte die Formel korrigieren oder den Wert direkt eingeben.',
      en: 'Item {n}: {label} shows a formula error ({v}). Please fix the formula or type the value.' },
    CREDIT_NOTE_REFERENCE: { de: 'Rechnungskorrektur: Nummer der ursprünglichen Rechnung fehlt ("Bezug auf Rechnung Nr.").',
      en: 'Credit note: number of the original invoice is missing ("Corrects invoice no.").' },
    CREDIT_NOTE_TERMS: { de: 'Rechnungskorrektur: Die Zahlungsbedingungen fordern zur Zahlung auf. Bitte durch einen Hinweis zur Erstattung ersetzen, zum Beispiel "Der Betrag wird Ihnen erstattet."',
      en: 'Credit note: the payment terms ask for payment. Please replace them with a refund note, for example "The amount will be refunded to you."' },
    LEITWEG_CHECKSUM: { de: 'Die Käuferreferenz sieht wie eine Leitweg-ID aus, aber die Prüfziffer stimmt nicht. Sie wird als normale Referenz übernommen.',
      en: 'The buyer reference looks like a Leitweg-ID but its check digits are wrong. It is used as a plain reference.' },
    PREPAID_VAT_UNKNOWN: { de: 'Bereits bezahlt: Bei mehreren Steuersätzen kann die darin enthaltene Umsatzsteuer nicht berechnet werden. Bei einer Schlussrechnung nach Anzahlungsrechnungen bitte die Umsatzsteuer der Anzahlungen je Steuersatz in der Bemerkung angeben.',
      en: 'Already paid: with several VAT rates the VAT contained cannot be calculated. In a final invoice after advance invoices, please state the VAT of the advance payments per rate in the note.' },
    TOTAL_ZERO: { de: 'Der Gesamtbetrag ist 0,00 EUR. Bitte Mengen und Preise prüfen.', en: 'The total is 0.00 EUR. Please check quantities and prices.' },
    CREDIT_NOTE_NOT_POSITIVE: { de: 'Der Gesamtbetrag ist 0 oder negativ. Bei einer Rechnungskorrektur (381) die Mengen und Preise positiv eingeben, die Dokumentart kehrt das Vorzeichen um.',
      en: 'The total is zero or negative. For a credit note (381) enter positive quantities and prices, the document type reverses the sign.' },
    PREPAID_RANGE: { de: 'Bereits bezahlt muss zwischen 0 und dem Gesamtbetrag ({g} EUR) liegen.', en: 'Already paid must be between 0 and the total ({g} EUR).' },
    TOTAL_NOT_POSITIVE: { de: 'Der Gesamtbetrag ist 0 oder negativ. Für Gutschriften die Dokumentart "381 Rechnungskorrektur" wählen.',
      en: 'The total is zero or negative. For credit notes choose document type "381 Credit note".' },
    CURRENCY_EUR: { de: 'Alle Beträge werden in EUR ausgegeben.', en: 'All amounts are issued in EUR.' },
    ROW_NOT_USED: { de: 'Zeile {row} wird nicht in die E-Rechnung übernommen. Weitere Positionen bitte oberhalb von "Summe Positionen" einfügen, andere Angaben in ein vorhandenes Feld schreiben oder die Zeile leeren.',
      en: 'Row {row} is not taken into the e-invoice. Insert more items above "Items net total", put other details into an existing field, or clear the row.' },
    DUPLICATE_FIELD: { de: '{label} steht mehrfach im Blatt (Zeilen {rows}). Bitte nur eine dieser Zeilen behalten.',
      en: '{label} appears more than once in the sheet (rows {rows}). Please keep only one of these rows.' },
    DUPLICATE_TEMPLATE: { de: 'Das Blatt enthält eine zweite Rechnungsvorlage (ab Zeile {row}). Bitte für jede Rechnung ein eigenes Blatt verwenden.',
      en: 'The sheet contains a second invoice template (from row {row}). Please use a separate sheet for each invoice.' },
    REQUIRED_ROW_MISSING: { de: 'Die Zeile "{label}" fehlt im Blatt. Bitte die Löschung rückgängig machen oder ein neues Blatt mit "Leere Vorlage anlegen" erstellen.',
      en: 'The row "{label}" is missing from the sheet. Please undo the deletion or create a new sheet with "Create empty template".' }
  };
  function formatMessage(issue, lang) {
    var l = lang === 'en' ? 'en' : 'de';
    var tpl = (MESSAGES[issue.code] || { de: issue.code, en: issue.code })[l];
    var params = issue.params || {};
    return tpl.replace(/\{(\w+)\}/g, function (_, k) {
      if (k === 'label' && issue.field) { return fieldLabel(issue.field, l); }
      var v = params[k];
      if (v && typeof v === 'object') { v = v[l]; } // numbers written per language, see localNumber()
      return v !== undefined && v !== null ? String(v) : '';
    });
  }

  // ---------------------------------------------------------------------------
  // Sheet -> model -> issues.
  // ---------------------------------------------------------------------------
  // values: 2D array as read from the sheet (dates already "YYYY-MM-DD" strings or Date objects).
  // Besides fields and lines it reports what the template does not read, so nothing typed there is lost silently:
  // ignoredRows (rows with values in C to J that are neither a field nor a line, SRV-3), fieldRows (every row of a
  // field key, to find copied rows, SRV-8) and secondTemplateRow (a second template block in the same sheet).
  var IGNORED_ROW_COLUMNS = [3, 5, 6, 9, 10];
  function parseSheet(values) {
    var raw = {};
    var cells = {};
    var lines = [];
    var fieldRows = {};
    var ignoredRows = [];
    var foundMarker = false;
    var markers = 0;
    var lineBlocks = 0;
    var secondTemplateRow = null;
    var state = 'fields';
    for (var r = 0; r < values.length; r++) {
      var row = values[r] || [];
      var key = text(row[0]);
      if (key === TEMPLATE_MARKER || key === '#lines') {
        if (key === TEMPLATE_MARKER) { foundMarker = true; markers += 1; } else { lineBlocks += 1; state = 'lines'; }
        if ((markers > 1 || lineBlocks > 1) && !secondTemplateRow) { secondTemplateRow = r + 1; }
        continue;
      }
      if (key === '#lines-end') { state = 'fields'; continue; }
      var isField = key !== '' && key.charAt(0) !== '#' && !!FIELD_BY_KEY[key];
      // In the line block every row is a line row, also when something was typed into hidden column A (SRV-9).
      if (state === 'lines' && !isField && key.charAt(0) !== '#') {
        var line = { row: r + 1, cells: {} };
        LINE_COLUMNS.forEach(function (c) {
          line[c.key] = row[c.col - 1];
          line.cells[c.key] = a1(r + 1, c.col);
        });
        if (!isBlank(line.description) || !isBlank(line.quantity) || !isBlank(line.price)) { lines.push(line); }
        continue;
      }
      if (isField) {
        state = 'fields';
        (fieldRows[key] = fieldRows[key] || []).push(r + 1);
        if (fieldRows[key].length === 1) {
          raw[key] = row[2];
          cells[key] = a1(r + 1, 3);
        }
        continue;
      }
      // Columns C, E, F, I and J. D, G and H are left out: empty line rows carry the dropdown defaults there.
      for (var k = 0; k < IGNORED_ROW_COLUMNS.length; k++) {
        if (!isBlank(row[IGNORED_ROW_COLUMNS[k] - 1])) {
          ignoredRows.push({ row: r + 1, cell: a1(r + 1, IGNORED_ROW_COLUMNS[k]), value: row[IGNORED_ROW_COLUMNS[k] - 1] });
          break;
        }
      }
    }
    return { foundMarker: foundMarker, raw: raw, cells: cells, lines: lines, fieldRows: fieldRows, ignoredRows: ignoredRows,
      secondTemplateRow: secondTemplateRow };
  }

  // A decimal for a message, written the German and the English way.
  function localNumber(dec) { return { de: formatDecimal(dec, 'de', 0, 4), en: formatDecimal(dec, 'en', 0, 4) }; }
  function localMoney(dec) { return { de: formatMoney(dec, 'de'), en: formatMoney(dec, 'en') }; }
  // Percent values between 0 and 1 are almost always "7%" typed into a number cell, which Sheets stores as 0.07.
  function looksLikeFraction(dec) { return dec !== null && cmp(dec, '0') > 0 && cmp(dec, '1') < 0; }
  function fractionParams(dec, extra) {
    var p = { d: localNumber(dec), x: localNumber(normalizeDecimal(mul(dec, '100'))) };
    Object.keys(extra || {}).forEach(function (k) { p[k] = extra[k]; });
    return p;
  }

  function issue(level, code, field, cell, params) {
    var o = { level: level, code: code };
    if (field) { o.field = field; }
    if (cell) { o.cell = cell; }
    if (params) { o.params = params; }
    return o;
  }

  var FORMULA_ERROR_VALUE = /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|ERROR!)$/;

  // Turns parsed sheet data into a normalized model and a list of issues.
  // opts.unitCodes: optional list of extra valid UN/ECE Rec 20 codes (from the library schema).
  function buildModel(parsed, opts) {
    opts = opts || {};
    var issues = [];
    var raw = parsed.raw;
    var cells = parsed.cells;
    function t(key) { return text(raw[key]); }
    // A required value is missing; when its whole row is gone, say so (SRV-9).
    function missing(key) {
      issues.push(cells[key] ? issue('error', 'REQUIRED', key, cells[key]) : issue('error', 'REQUIRED_ROW_MISSING', key, null));
    }
    function req(key) {
      if (t(key) === '') { missing(key); }
      return t(key);
    }
    function date(key, required) {
      var v = raw[key];
      if (isBlank(v)) {
        if (required) { missing(key); }
        return null;
      }
      var d = parseDate(v);
      if (!d) { issues.push(issue('error', 'BAD_DATE', key, cells[key])); }
      return d;
    }
    function num(key) {
      var v = raw[key];
      if (isBlank(v)) { return null; }
      var n = parseNumber(v);
      if (n === null) { issues.push(issue('error', 'BAD_NUMBER', key, cells[key])); }
      return n;
    }
    function country(key, required) {
      var v = raw[key];
      if (isBlank(v)) {
        if (required) { missing(key); }
        return null;
      }
      var c = normalizeCountry(v);
      if (!c) { issues.push(issue('error', 'BAD_COUNTRY', key, cells[key])); }
      return c;
    }

    // What the template does not read (SRV-3, SRV-8).
    if (parsed.secondTemplateRow) {
      issues.push(issue('error', 'DUPLICATE_TEMPLATE', null, a1(parsed.secondTemplateRow, 2), { row: parsed.secondTemplateRow }));
    } else {
      Object.keys(parsed.fieldRows || {}).forEach(function (key) {
        var rowsOfKey = parsed.fieldRows[key];
        if (rowsOfKey.length > 1) {
          issues.push(issue('error', 'DUPLICATE_FIELD', key, a1(rowsOfKey[1], 3), { rows: rowsOfKey.join(', ') }));
        }
      });
    }
    // A number there is most likely an item or an amount that would be lost: refuse. Text (a note, an extra address
    // line, a heading) only gets a warning, so customised sheets keep working.
    (parsed.ignoredRows || []).forEach(function (x) {
      var isNum = typeof x.value === 'number' || (typeof x.value === 'string' && parseNumber(x.value) !== null);
      issues.push(issue(isNum ? 'error' : 'warning', 'ROW_NOT_USED', null, x.cell, { row: x.row }));
    });

    var model = { version: VERSION, currency: 'EUR' };
    var docType = normalizeDocType(raw['doc.type']);
    if (!docType) { issues.push(issue('error', 'BAD_DOC_TYPE', 'doc.type', cells['doc.type'])); docType = '380'; }
    model.type = docType;
    model.number = req('doc.number');
    model.issueDate = date('doc.date', true);
    model.deliveryDate = date('doc.deliveryDate', false);
    model.periodStart = date('doc.periodStart', false);
    model.periodEnd = date('doc.periodEnd', false);
    model.dueDate = date('doc.dueDate', false);
    model.buyerReference = req('doc.buyerReference');
    model.orderReference = t('doc.orderReference');
    model.precedingInvoice = t('doc.precedingInvoice');
    model.precedingInvoiceDate = date('doc.precedingInvoiceDate', false);
    model.note = t('doc.note');

    var seller = {
      name: req('seller.name'), street: req('seller.street'), zip: req('seller.zip'), city: req('seller.city'),
      country: country('seller.country', true), vatId: normalizeVatId(raw['seller.vatId']),
      taxNumber: t('seller.taxNumber'), registration: t('seller.registration'),
      contactName: req('seller.contactName'), phone: req('seller.phone'), email: req('seller.email'),
      iban: normalizeIban(raw['seller.iban']), bic: t('seller.bic').replace(/\s/g, '').toUpperCase(),
      accountHolder: t('seller.accountHolder')
    };
    if (seller.iban === '') { missing('seller.iban'); }
    else if (!isValidIban(seller.iban)) { issues.push(issue('error', 'BAD_IBAN', 'seller.iban', cells['seller.iban'])); }
    if (seller.phone && countDigits(seller.phone) < 3) { issues.push(issue('error', 'BAD_PHONE', 'seller.phone', cells['seller.phone'])); }
    if (seller.email && !isValidEmail(seller.email)) { issues.push(issue('error', 'BAD_EMAIL', 'seller.email', cells['seller.email'])); }
    if (seller.vatId && !isPlausibleVatId(seller.vatId)) { issues.push(issue('error', 'BAD_VAT_ID', 'seller.vatId', cells['seller.vatId'])); }
    model.seller = seller;

    var buyer = {
      name: req('buyer.name'), contactName: t('buyer.contactName'), street: req('buyer.street'),
      zip: req('buyer.zip'), city: req('buyer.city'), country: country('buyer.country', true),
      vatId: normalizeVatId(raw['buyer.vatId']), email: t('buyer.email'), customerNumber: t('buyer.customerNumber')
    };
    if (buyer.vatId && !isPlausibleVatId(buyer.vatId)) { issues.push(issue('error', 'BAD_VAT_ID', 'buyer.vatId', cells['buyer.vatId'])); }
    if (buyer.email && !isValidEmail(buyer.email)) { issues.push(issue('error', 'BAD_EMAIL', 'buyer.email', cells['buyer.email'])); }
    model.buyer = buyer;

    // Buyer reference: Leitweg-ID detection.
    model.leitwegId = false;
    // A reference with valid Leitweg check digits is a Leitweg-ID, with or without a buyer e-mail (public bodies often
    // give both). Wrong check digits are only flagged without a buyer e-mail: with one it is most likely a B2B order
    // number of the same shape (INV-12).
    if (model.buyerReference && leitwegShape(model.buyerReference)) {
      if (isValidLeitwegId(model.buyerReference)) { model.leitwegId = true; }
      else if (!buyer.email) { issues.push(issue('warning', 'LEITWEG_CHECKSUM', 'doc.buyerReference', cells['doc.buyerReference'])); }
    }
    if (!model.leitwegId && !buyer.email) {
      issues.push(issue('error', 'BUYER_EMAIL_NEEDED', 'buyer.email', cells['buyer.email']));
    }

    // Dates.
    if (!model.deliveryDate && !model.periodStart && !model.periodEnd) {
      issues.push(issue('error', 'NO_DELIVERY_DATE', 'doc.deliveryDate', cells['doc.deliveryDate']));
    }
    if ((model.periodStart && !model.periodEnd) || (!model.periodStart && model.periodEnd)) {
      var missing = model.periodStart ? 'doc.periodEnd' : 'doc.periodStart';
      issues.push(issue('error', 'PERIOD_INCOMPLETE', missing, cells[missing]));
    }
    if (model.periodStart && model.periodEnd && model.periodEnd < model.periodStart) {
      issues.push(issue('error', 'PERIOD_ORDER', 'doc.periodEnd', cells['doc.periodEnd']));
    }
    // opts.today (YYYY-MM-DD, from the sidebar): an invoice date more than a year away is most likely a typo.
    if (opts.today && model.issueDate &&
        Math.abs(Date.parse(model.issueDate + 'T00:00:00Z') - Date.parse(opts.today + 'T00:00:00Z')) > 365 * 86400000) {
      issues.push(issue('warning', 'DATE_FAR', 'doc.date', cells['doc.date'],
        { d: { de: formatDate(model.issueDate, 'de'), en: formatDate(model.issueDate, 'en') } }));
    }
    if (model.dueDate && model.issueDate && model.dueDate < model.issueDate) {
      issues.push(issue('error', 'DUE_BEFORE_ISSUE', 'doc.dueDate', cells['doc.dueDate']));
    }
    if (model.type === '381' && !model.precedingInvoice) {
      issues.push(issue('error', 'CREDIT_NOTE_REFERENCE', 'doc.precedingInvoice', cells['doc.precedingInvoice']));
    }

    // Lines.
    model.lines = [];
    var categories = {};
    parsed.lines.forEach(function (pl, idx) {
      var n = idx + 1;
      var line = { n: n, row: pl.row, cells: pl.cells };
      line.description = text(pl.description);
      if (!line.description) { issues.push(issue('error', 'LINE_REQUIRED', 'line.description', pl.cells.description, { n: n })); }
      line.quantity = isBlank(pl.quantity) ? null : parseNumber(pl.quantity);
      if (isBlank(pl.quantity)) { issues.push(issue('error', 'LINE_REQUIRED', 'line.quantity', pl.cells.quantity, { n: n })); }
      else if (line.quantity === null) { issues.push(issue('error', 'LINE_BAD_NUMBER', 'line.quantity', pl.cells.quantity, { n: n })); }
      else if (isZero(line.quantity)) { issues.push(issue('error', 'LINE_ZERO_QTY', 'line.quantity', pl.cells.quantity, { n: n })); }
      line.price = isBlank(pl.price) ? null : parseNumber(pl.price);
      if (isBlank(pl.price)) { issues.push(issue('error', 'LINE_REQUIRED', 'line.price', pl.cells.price, { n: n })); }
      else if (line.price === null) { issues.push(issue('error', 'LINE_BAD_NUMBER', 'line.price', pl.cells.price, { n: n })); }
      else if (cmp(line.price, '0') < 0) { issues.push(issue('error', 'LINE_NEGATIVE_PRICE', 'line.price', pl.cells.price, { n: n })); }
      line.discount = null;
      if (!isBlank(pl.discount)) {
        line.discount = parseNumber(pl.discount);
        if (line.discount === null) { issues.push(issue('error', 'LINE_BAD_NUMBER', 'line.discount', pl.cells.discount, { n: n })); }
        else if (cmp(line.discount, '0') < 0 || cmp(line.discount, '100') > 0) {
          issues.push(issue('error', 'LINE_BAD_DISCOUNT', 'line.discount', pl.cells.discount, { n: n }));
          line.discount = null;
        } else if (isZero(line.discount)) { line.discount = null; }
        else if (looksLikeFraction(line.discount)) {
          issues.push(issue('warning', 'LINE_DISCOUNT_FRACTION', 'line.discount', pl.cells.discount, fractionParams(line.discount, { n: n })));
        }
      }
      // The Netto column holds a formula; a number typed over it shows in the sheet but is not used (INV-13).
      var sheetNet = isBlank(pl.net) ? null : parseNumber(pl.net);
      var sheetDiscount = isBlank(pl.discount) ? '0' : parseNumber(pl.discount);
      if (sheetNet !== null && sheetDiscount !== null && line.quantity !== null && line.price !== null) {
        var lineGross = round2(mul(line.quantity, line.price));
        var lineCalc = sub(lineGross, round2(percentOf(lineGross, sheetDiscount)));
        var netDiff = sub(sheetNet, lineCalc);
        if (cmp(netDiff, '0.01') > 0 || cmp(netDiff, '-0.01') < 0) {
          issues.push(issue('warning', 'LINE_NET_DIFFERS', 'line.net', pl.cells.net, { n: n, s: localMoney(sheetNet), c: localMoney(lineCalc) }));
        }
      }
      line.unit = normalizeUnit(pl.unit, opts.unitCodes);
      if (!line.unit) { issues.push(issue('error', 'LINE_BAD_UNIT', 'line.unit', pl.cells.unit, { n: n })); line.unit = 'H87'; }
      line.category = normalizeCategory(pl.category);
      if (!line.category) { issues.push(issue('error', 'LINE_BAD_CATEGORY', 'line.category', pl.cells.category, { n: n })); line.category = 'S'; }
      var cat = VAT_CATEGORIES[line.category];
      var rate = isBlank(pl.rate) ? '0' : parseNumber(pl.rate);
      if (rate === null) { issues.push(issue('error', 'LINE_BAD_NUMBER', 'line.rate', pl.cells.rate, { n: n })); rate = '0'; }
      if (line.category === 'S' && cmp(rate, '0') <= 0) {
        issues.push(issue('error', 'LINE_RATE_NEEDED', 'line.rate', pl.cells.rate, { n: n }));
      }
      // No VAT rate in the EU is below 1 %: 0.07 is "7%" typed into the cell (SRV-1).
      if (line.category === 'S' && looksLikeFraction(rate)) {
        issues.push(issue('error', 'LINE_RATE_FRACTION', 'line.rate', pl.cells.rate, { n: n, r: localNumber(rate) }));
      }
      if (cat.zeroRate && !isZero(rate)) {
        issues.push(issue('error', 'LINE_RATE_ZERO', 'line.rate', pl.cells.rate, { n: n, cat: line.category }));
        rate = '0';
      }
      line.rate = cat.noRate ? null : rate;
      line.itemNo = text(pl.itemNo);
      categories[line.category] = true;
      model.lines.push(line);
    });
    if (model.lines.length === 0) { issues.push(issue('error', 'NO_LINES', null, null)); }
    model.categories = VAT_CATEGORY_ORDER.filter(function (k) { return categories[k]; });

    // Document level values.
    model.discountPercent = num('totals.discountPercent');
    if (model.discountPercent !== null && (cmp(model.discountPercent, '0') < 0 || cmp(model.discountPercent, '100') > 0)) {
      issues.push(issue('error', 'BAD_DISCOUNT', 'totals.discountPercent', cells['totals.discountPercent']));
      model.discountPercent = null;
    }
    if (model.discountPercent !== null && isZero(model.discountPercent)) { model.discountPercent = null; }
    if (looksLikeFraction(model.discountPercent)) {
      issues.push(issue('warning', 'DISCOUNT_FRACTION', 'totals.discountPercent', cells['totals.discountPercent'], fractionParams(model.discountPercent)));
    }
    model.discountReason = t('totals.discountReason') || 'Rabatt';
    model.prepaid = num('totals.prepaid');
    if (model.prepaid !== null && isZero(model.prepaid)) { model.prepaid = null; }
    model.exemptionReason = t('tax.exemptionReason');
    model.payment = {
      terms: t('payment.terms'),
      skontoPercent: num('payment.skontoPercent'),
      skontoDays: num('payment.skontoDays'),
      reference: t('payment.reference') || model.number
    };
    // Lines starting with # are XRechnung Skonto lines (BR-DE-18); anything else there makes the file invalid (INV-3).
    if (splitTerms(model.payment.terms).skonto.some(function (l) { return !SKONTO_LINE.test(l); })) {
      issues.push(issue('error', 'PAYMENT_TERMS_HASH', 'payment.terms', cells['payment.terms']));
    }
    // A credit note asks for no payment; the template's "Zahlbar innerhalb von 14 Tagen" reads like a payment demand (PDF-7).
    if (model.type === '381' && /^\s*(zahlbar|payable)\b/i.test(model.payment.terms || '')) {
      issues.push(issue('warning', 'CREDIT_NOTE_TERMS', 'payment.terms', cells['payment.terms']));
    }
    var sp = model.payment.skontoPercent;
    var sd = model.payment.skontoDays;
    if ((sp !== null) !== (sd !== null)) {
      var mf = sp === null ? 'payment.skontoPercent' : 'payment.skontoDays';
      issues.push(issue('error', 'SKONTO_INCOMPLETE', mf, cells[mf]));
    } else if (sp !== null) {
      if (cmp(sp, '0') <= 0 || cmp(sp, '100') >= 0 || !/^\d+$/.test(sd) || cmp(sd, '0') <= 0) {
        issues.push(issue('error', 'SKONTO_BAD', 'payment.skontoPercent', cells['payment.skontoPercent']));
      } else if (looksLikeFraction(sp)) {
        issues.push(issue('warning', 'SKONTO_FRACTION', 'payment.skontoPercent', cells['payment.skontoPercent'], fractionParams(sp)));
      }
    }

    // VAT rules that depend on the categories in use.
    var cats = model.categories;
    var has = function (c) { return cats.indexOf(c) >= 0; };
    if (has('O') && cats.length > 1) { issues.push(issue('error', 'O_MIXED', null, null)); }
    if (has('O') && (seller.vatId || buyer.vatId)) { issues.push(issue('warning', 'O_VAT_IDS', null, null)); }
    var needsSellerTaxId = cats.some(function (c) { return c !== 'O'; });
    if (needsSellerTaxId && !seller.vatId && !seller.taxNumber) {
      issues.push(issue('error', 'SELLER_TAX_ID', 'seller.vatId', cells['seller.vatId']));
    }
    // With O the VAT ID must stay out (BR-O-02), but the seller still needs an identifier (BR-CO-26): the tax number
    // (sent as seller identifier BT-29) or the register number (BT-30).
    if (has('O') && !seller.taxNumber && !seller.registration) {
      issues.push(issue('error', 'O_SELLER_ID', 'seller.taxNumber', cells['seller.taxNumber']));
    }
    ['K', 'G'].forEach(function (c) {
      if (has(c) && !seller.vatId) {
        issues.push(issue('error', 'SELLER_VAT_ID_NEEDED', 'seller.vatId', cells['seller.vatId'], { cat: c }));
      }
    });
    ['AE', 'K'].forEach(function (c) {
      if (has(c) && !buyer.vatId) {
        issues.push(issue('error', 'BUYER_VAT_ID_NEEDED', 'buyer.vatId', cells['buyer.vatId'], { cat: c }));
      }
    });
    if (has('E') && !model.exemptionReason) {
      issues.push(issue('error', 'EXEMPTION_REASON_NEEDED', 'tax.exemptionReason', cells['tax.exemptionReason']));
    }
    // Intra-EU supplies need a deliver-to address (BR-IC-12; XRechnung BR-DE-10/11 want city and postcode).
    // Version 1 delivers to the buyer's address.
    model.deliverTo = has('K') ? { name: buyer.name, street: buyer.street, city: buyer.city, zip: buyer.zip, country: buyer.country } : null;

    // Formula errors (#N/A from a failed lookup and the like) in cells no other check has flagged (SRV-5).
    var flagged = {};
    issues.forEach(function (i) { if (i.cell) { flagged[i.cell] = true; } });
    Object.keys(raw).forEach(function (key) {
      if (cells[key] && !flagged[cells[key]] && FORMULA_ERROR_VALUE.test(text(raw[key]))) {
        issues.push(issue('error', 'FORMULA_ERROR', key, cells[key], { v: text(raw[key]) }));
      }
    });
    parsed.lines.forEach(function (pl, idx) {
      LINE_COLUMNS.forEach(function (c) {
        if (c.key !== 'net' && !flagged[pl.cells[c.key]] && FORMULA_ERROR_VALUE.test(text(pl[c.key]))) {
          issues.push(issue('error', 'LINE_FORMULA_ERROR', 'line.' + c.key, pl.cells[c.key], { n: idx + 1, v: text(pl[c.key]) }));
        }
      });
    });

    var hasErrors = issues.some(function (i) { return i.level === 'error'; });
    if (!hasErrors) {
      model.totals = computeTotals(model);
      var sign = cmp(model.totals.taxInclusive, '0');
      if (sign <= 0) {
        issues.push(issue('error', model.type === '381' ? 'CREDIT_NOTE_NOT_POSITIVE' : (sign === 0 ? 'TOTAL_ZERO' : 'TOTAL_NOT_POSITIVE'), null, null));
      }
      // A typo in "Bereits bezahlt" would give a negative or a raised amount due (INV-8).
      var pre = model.totals.prepaid;
      if (pre && (cmp(pre, '0') < 0 || cmp(pre, model.totals.taxInclusive) > 0)) {
        issues.push(issue('error', 'PREPAID_RANGE', 'totals.prepaid', cells['totals.prepaid'],
          { g: { de: formatMoney(model.totals.taxInclusive, 'de'), en: formatMoney(model.totals.taxInclusive, 'en') } }));
      }
      // A final invoice must state the VAT contained in invoiced advance payments (section 14 (5) UStG, UStAE 14.8 (7)).
      // With one rate it is derived; with several the user has to state it (PDF-1).
      var t0 = model.totals;
      if (t0.prepaid && !t0.prepaidVat && t0.vat.some(function (v) { return v.category === 'S'; })) {
        issues.push(issue('warning', 'PREPAID_VAT_UNKNOWN', 'totals.prepaid', cells['totals.prepaid']));
      }
    }
    return { model: model, issues: issues };
  }

  // Convenience: values -> {model, issues, ok}
  function checkSheet(values, opts) {
    var parsed = parseSheet(values);
    if (!parsed.foundMarker && Object.keys(parsed.raw).length === 0) {
      return { ok: false, model: null, issues: [issue('error', 'NO_TEMPLATE', null, null)] };
    }
    var res = buildModel(parsed, opts);
    res.ok = !res.issues.some(function (i) { return i.level === 'error'; });
    return res;
  }

  // ---------------------------------------------------------------------------
  // EN 16931 totals.
  // ---------------------------------------------------------------------------
  function groupKey(category, rate) { return category + '|' + (rate === null ? '' : normalizeDecimal(rate)); }

  function computeTotals(model) {
    var groups = {};
    var order = [];
    function group(category, rate) {
      var k = groupKey(category, rate);
      if (!groups[k]) {
        groups[k] = { category: category, rate: rate === null ? null : normalizeDecimal(rate), lineNet: '0.00' };
        order.push(k);
      }
      return groups[k];
    }
    var lines = model.lines.map(function (l) {
      var gross = round2(mul(l.quantity, l.price));
      var allowance = l.discount ? round2(percentOf(gross, l.discount)) : null;
      var net = allowance ? sub(gross, allowance) : gross;
      var g = group(l.category, l.rate);
      g.lineNet = add(g.lineNet, net);
      return { n: l.n, gross: round2(gross), allowance: allowance, net: round2(net) };
    });
    var lineTotal = round2(sum(lines.map(function (x) { return x.net; })));
    var allowances = [];
    if (model.discountPercent) {
      order.forEach(function (k) {
        var g = groups[k];
        var amount = round2(percentOf(g.lineNet, model.discountPercent));
        if (!isZero(amount)) {
          allowances.push({ category: g.category, rate: g.rate, base: round2(g.lineNet), percent: model.discountPercent,
            amount: amount, reason: model.discountReason });
          g.allowance = amount;
        }
      });
    }
    var allowanceTotal = round2(sum(allowances.map(function (a) { return a.amount; })));
    var chargeTotal = '0.00';
    var taxExclusive = round2(sub(add(lineTotal, chargeTotal), allowanceTotal));
    var vat = order.map(function (k) {
      var g = groups[k];
      var taxable = round2(sub(g.lineNet, g.allowance || '0'));
      var rate = g.rate;
      var tax = rate === null ? '0.00' : round2(percentOf(taxable, rate));
      var cat = VAT_CATEGORIES[g.category];
      var out = { category: g.category, rate: rate, taxable: taxable, tax: tax };
      if (g.category === 'E') { out.reason = model.exemptionReason; }
      else if (cat.code) {
        out.reasonCode = cat.code;
        out.reason = (g.category === 'O' && model.exemptionReason) ? model.exemptionReason : cat.textDe;
      }
      return out;
    });
    var taxTotal = round2(sum(vat.map(function (v) { return v.tax; })));
    var taxInclusive = round2(add(taxExclusive, taxTotal));
    var prepaid = model.prepaid ? round2(model.prepaid) : null;
    var payable = round2(sub(taxInclusive, prepaid || '0'));
    // VAT contained in the amount already paid, when the invoice has exactly one VAT rate (PDF-1).
    var prepaidVat = null;
    if (prepaid && vat.length === 1 && vat[0].category === 'S') {
      prepaidVat = { rate: vat[0].rate, amount: divRound(mul(prepaid, vat[0].rate), add('100', vat[0].rate), 2) };
    }
    var skonto = null;
    if (model.payment && model.payment.skontoPercent) {
      // The XML states the percent with two decimals, so the amount uses that percent too (INV-9).
      var skPercent = round2(model.payment.skontoPercent);
      var skAmount = round2(percentOf(payable, skPercent));
      skonto = { percent: skPercent, days: model.payment.skontoDays,
        amount: skAmount, payableAfter: round2(sub(payable, skAmount)) };
    }
    return { lines: lines, lineTotal: lineTotal, allowances: allowances, allowanceTotal: allowanceTotal,
      chargeTotal: chargeTotal, taxExclusive: taxExclusive, vat: vat, taxTotal: taxTotal,
      taxInclusive: taxInclusive, prepaid: prepaid, prepaidVat: prepaidVat, payable: payable, skonto: skonto };
  }

  // ---------------------------------------------------------------------------
  // Model -> @e-invoice-eu/core input ("ubl:Invoice" shaped JSON).
  // ---------------------------------------------------------------------------
  // XRechnung structured cash discount line (BR-DE-18), as the official schematron checks it.
  var SKONTO_LINE = /^#SKONTO#TAGE=[0-9]+#PROZENT=[0-9]+\.[0-9]{2}(#BASISBETRAG=-?[0-9]+\.[0-9]{2})?#$/;
  // Splits payment terms into plain lines and lines starting with # (trimmed).
  function splitTerms(terms) {
    var out = { plain: [], skonto: [] };
    String(terms || '').split(/\r?\n/).forEach(function (line) {
      var x = line.trim();
      if (x.charAt(0) === '#') { out.skonto.push(x); } else { out.plain.push(line); }
    });
    return out;
  }
  function paymentTermsText(model) {
    var parts = splitTerms(model.payment.terms);
    var t = model.totals;
    if (t.skonto) { parts.skonto.push('#SKONTO#TAGE=' + normalizeDecimal(t.skonto.days) + '#PROZENT=' + t.skonto.percent + '#'); }
    var s = model.payment.terms ? parts.plain.join('\n') : '';
    // Skonto lines go last, and the last one too ends with a line feed (BR-DE-18).
    if (parts.skonto.length) { s = (s ? s + '\n' : '') + parts.skonto.join('\n') + '\n'; }
    return s === '' ? null : s;
  }

  function toInvoiceData(model) {
    var t = model.totals;
    if (!t) { throw new Error('Model has no totals. Run buildModel without errors first.'); }
    var cur = model.currency;
    var isO = model.categories.indexOf('O') >= 0;
    function amt(v) { return v; }
    var inv = {};
    inv['cbc:ID'] = model.number;
    inv['cbc:IssueDate'] = model.issueDate;
    if (model.dueDate) { inv['cbc:DueDate'] = model.dueDate; }
    inv['cbc:InvoiceTypeCode'] = model.type;
    var notes = [];
    if (model.note) { notes.push(model.note); }
    if (t.prepaidVat) {
      notes.push('Bereits bezahlt: ' + formatMoney(t.prepaid, 'de') + ' EUR, darin enthaltene Umsatzsteuer ' +
        formatDecimal(t.prepaidVat.rate, 'de', 0, 2) + ' %: ' + formatMoney(t.prepaidVat.amount, 'de') + ' EUR.');
    }
    if (notes.length) { inv['cbc:Note'] = notes; }
    inv['cbc:DocumentCurrencyCode'] = cur;
    inv['cbc:BuyerReference'] = model.buyerReference;
    if (model.periodStart && model.periodEnd) {
      inv['cac:InvoicePeriod'] = { 'cbc:StartDate': model.periodStart, 'cbc:EndDate': model.periodEnd };
    }
    if (model.orderReference) { inv['cac:OrderReference'] = { 'cbc:ID': model.orderReference }; }
    if (model.precedingInvoice) {
      var ref = { 'cbc:ID': model.precedingInvoice };
      if (model.precedingInvoiceDate) { ref['cbc:IssueDate'] = model.precedingInvoiceDate; }
      inv['cac:BillingReference'] = [{ 'cac:InvoiceDocumentReference': ref }];
    }

    var s = model.seller;
    var sellerTax = [];
    if (s.vatId && !isO) { sellerTax.push({ 'cbc:CompanyID': s.vatId, 'cac:TaxScheme': { 'cbc:ID': 'VAT' } }); }
    if (s.taxNumber) { sellerTax.push({ 'cbc:CompanyID': s.taxNumber, 'cac:TaxScheme': { 'cbc:ID': 'FC' } }); }
    var sellerParty = {
      'cbc:EndpointID': s.email,
      'cbc:EndpointID@schemeID': 'EM',
      'cac:PostalAddress': {
        'cbc:StreetName': s.street, 'cbc:CityName': s.city, 'cbc:PostalZone': s.zip,
        'cac:Country': { 'cbc:IdentificationCode': s.country }
      },
      'cac:PartyLegalEntity': { 'cbc:RegistrationName': s.name },
      'cac:Contact': { 'cbc:Name': s.contactName, 'cbc:Telephone': s.phone, 'cbc:ElectronicMail': s.email }
    };
    if (sellerTax.length) { sellerParty['cac:PartyTaxScheme'] = sellerTax; }
    // BR-CO-26 needs a seller identifier (BT-29), registration (BT-30) or VAT ID (BT-31).
    // Small businesses often have only a tax number, so it doubles as the seller identifier.
    if ((!s.vatId || isO) && !s.registration && s.taxNumber) {
      sellerParty['cac:PartyIdentification'] = [{ 'cbc:ID': s.taxNumber }];
    }
    if (s.registration) { sellerParty['cac:PartyLegalEntity']['cbc:CompanyID'] = s.registration; }
    inv['cac:AccountingSupplierParty'] = { 'cac:Party': sellerParty };

    var b = model.buyer;
    var buyerParty = {
      'cbc:EndpointID': model.leitwegId ? model.buyerReference : b.email,
      'cbc:EndpointID@schemeID': model.leitwegId ? '0204' : 'EM',
      'cac:PostalAddress': {
        'cbc:StreetName': b.street, 'cbc:CityName': b.city, 'cbc:PostalZone': b.zip,
        'cac:Country': { 'cbc:IdentificationCode': b.country }
      },
      'cac:PartyLegalEntity': { 'cbc:RegistrationName': b.name }
    };
    if (b.customerNumber) { buyerParty['cac:PartyIdentification'] = { 'cbc:ID': b.customerNumber }; }
    if (b.vatId && !isO) { buyerParty['cac:PartyTaxScheme'] = { 'cbc:CompanyID': b.vatId, 'cac:TaxScheme': { 'cbc:ID': 'VAT' } }; }
    if (b.contactName) { buyerParty['cac:Contact'] = { 'cbc:Name': b.contactName }; }
    inv['cac:AccountingCustomerParty'] = { 'cac:Party': buyerParty };

    var delivery = {};
    if (model.deliveryDate) { delivery['cbc:ActualDeliveryDate'] = model.deliveryDate; }
    if (model.deliverTo) {
      var dt = model.deliverTo;
      delivery['cac:DeliveryLocation'] = { 'cac:Address': {
        'cbc:StreetName': dt.street, 'cbc:CityName': dt.city, 'cbc:PostalZone': dt.zip,
        'cac:Country': { 'cbc:IdentificationCode': dt.country } } };
      delivery['cac:DeliveryParty'] = { 'cac:PartyName': { 'cbc:Name': dt.name } };
    }
    if (Object.keys(delivery).length) { inv['cac:Delivery'] = delivery; }

    var account = { 'cbc:ID': s.iban };
    if (s.accountHolder) { account['cbc:Name'] = s.accountHolder; }
    if (s.bic) { account['cac:FinancialInstitutionBranch'] = { 'cbc:ID': s.bic }; }
    inv['cac:PaymentMeans'] = [{
      'cbc:PaymentMeansCode': '58',
      'cbc:PaymentMeansCode@name': 'SEPA credit transfer',
      'cbc:PaymentID': model.payment.reference,
      'cac:PayeeFinancialAccount': account
    }];
    var terms = paymentTermsText(model);
    if (terms) { inv['cac:PaymentTerms'] = { 'cbc:Note': terms }; }

    if (t.allowances.length) {
      inv['cac:AllowanceCharge'] = t.allowances.map(function (a) {
        var cat = { 'cbc:ID': a.category, 'cac:TaxScheme': { 'cbc:ID': 'VAT' } };
        if (a.rate !== null) { cat['cbc:Percent'] = a.rate; }
        return {
          'cbc:ChargeIndicator': 'false',
          'cbc:AllowanceChargeReasonCode': '95',
          'cbc:AllowanceChargeReason': a.reason,
          'cbc:MultiplierFactorNumeric': normalizeDecimal(a.percent),
          'cbc:Amount': amt(a.amount), 'cbc:Amount@currencyID': cur,
          'cbc:BaseAmount': amt(a.base), 'cbc:BaseAmount@currencyID': cur,
          'cac:TaxCategory': cat
        };
      });
    }

    inv['cac:TaxTotal'] = [{
      'cbc:TaxAmount': amt(t.taxTotal), 'cbc:TaxAmount@currencyID': cur,
      'cac:TaxSubtotal': t.vat.map(function (v) {
        var cat = { 'cbc:ID': v.category };
        // XRechnung (BR-DE-14) wants a rate in every VAT breakdown; for O it is 0.
        cat['cbc:Percent'] = v.rate !== null ? v.rate : '0';
        if (v.reasonCode) { cat['cbc:TaxExemptionReasonCode'] = v.reasonCode; }
        if (v.reason) { cat['cbc:TaxExemptionReason'] = v.reason; }
        cat['cac:TaxScheme'] = { 'cbc:ID': 'VAT' };
        return {
          'cbc:TaxableAmount': amt(v.taxable), 'cbc:TaxableAmount@currencyID': cur,
          'cbc:TaxAmount': amt(v.tax), 'cbc:TaxAmount@currencyID': cur,
          'cac:TaxCategory': cat
        };
      })
    }];

    var lmt = {
      'cbc:LineExtensionAmount': amt(t.lineTotal), 'cbc:LineExtensionAmount@currencyID': cur,
      'cbc:TaxExclusiveAmount': amt(t.taxExclusive), 'cbc:TaxExclusiveAmount@currencyID': cur,
      'cbc:TaxInclusiveAmount': amt(t.taxInclusive), 'cbc:TaxInclusiveAmount@currencyID': cur
    };
    if (t.allowances.length) { lmt['cbc:AllowanceTotalAmount'] = amt(t.allowanceTotal); lmt['cbc:AllowanceTotalAmount@currencyID'] = cur; }
    if (t.prepaid) { lmt['cbc:PrepaidAmount'] = amt(t.prepaid); lmt['cbc:PrepaidAmount@currencyID'] = cur; }
    lmt['cbc:PayableAmount'] = amt(t.payable);
    lmt['cbc:PayableAmount@currencyID'] = cur;
    inv['cac:LegalMonetaryTotal'] = lmt;

    inv['cac:InvoiceLine'] = model.lines.map(function (l, idx) {
      var lt = t.lines[idx];
      var cat = { 'cbc:ID': l.category, 'cac:TaxScheme': { 'cbc:ID': 'VAT' } };
      if (l.rate !== null) { cat['cbc:Percent'] = normalizeDecimal(l.rate); }
      var item = { 'cbc:Name': l.description, 'cac:ClassifiedTaxCategory': cat };
      if (l.itemNo) { item['cac:SellersItemIdentification'] = { 'cbc:ID': l.itemNo }; }
      var out = {
        'cbc:ID': String(l.n),
        'cbc:InvoicedQuantity': normalizeDecimal(l.quantity),
        'cbc:InvoicedQuantity@unitCode': l.unit,
        'cbc:LineExtensionAmount': amt(lt.net), 'cbc:LineExtensionAmount@currencyID': cur
      };
      if (lt.allowance) {
        out['cac:AllowanceCharge'] = [{
          'cbc:ChargeIndicator': 'false',
          'cbc:AllowanceChargeReasonCode': '95',
          'cbc:AllowanceChargeReason': 'Rabatt',
          'cbc:MultiplierFactorNumeric': normalizeDecimal(l.discount),
          'cbc:Amount': amt(lt.allowance), 'cbc:Amount@currencyID': cur,
          'cbc:BaseAmount': amt(lt.gross), 'cbc:BaseAmount@currencyID': cur
        }];
      }
      out['cac:Item'] = item;
      out['cac:Price'] = { 'cbc:PriceAmount': normalizeDecimal(l.price), 'cbc:PriceAmount@currencyID': cur };
      return out;
    });
    return { 'ubl:Invoice': inv };
  }

  // Formats for the two output files. The PDF uses the XRechnung profile when a valid Leitweg-ID is present.
  function outputFormats(model) {
    return { xml: 'XRECHNUNG-CII', pdf: model.leitwegId ? 'Factur-X-XRechnung' : 'Factur-X-EN16931' };
  }
  function safeFileName(s) {
    var base = text(s).replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '-').replace(/\s+/g, '_').replace(/^[.-]+/, '');
    return base.slice(0, 80) || 'Rechnung';
  }
  function outputFileNames(model) {
    var base = safeFileName(model.number);
    var kind = model.type === '381' ? 'Rechnungskorrektur' : 'Rechnung';
    return { xml: kind + '_' + base + '_XRechnung.xml', pdf: kind + '_' + base + '_ZUGFeRD.pdf' };
  }

  // ---------------------------------------------------------------------------
  // Formatting for display (sidebar and PDF).
  // ---------------------------------------------------------------------------
  function formatDecimal(dec, lang, minFrac, maxFrac) {
    if (dec === null || dec === undefined) { return ''; }
    var s = dToString(dRound(dParse(dec), maxFrac));
    var neg = s.charAt(0) === '-';
    if (neg) { s = s.slice(1); }
    var parts = s.split('.');
    var intPart = parts[0];
    var frac = parts[1] || '';
    while (frac.length > minFrac && frac.charAt(frac.length - 1) === '0') { frac = frac.slice(0, -1); }
    while (frac.length < minFrac) { frac += '0'; }
    var groupSep = lang === 'en' ? ',' : '.';
    var decSep = lang === 'en' ? '.' : ',';
    intPart = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, groupSep);
    return (neg ? '-' : '') + intPart + (frac ? decSep + frac : '');
  }
  function formatMoney(dec, lang) { return formatDecimal(dec, lang, 2, 2); }
  var MONTHS_EN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  function formatDate(ymd, lang) {
    if (!ymd) { return ''; }
    var p = ymd.split('-');
    if (lang === 'en') { return (+p[2]) + ' ' + MONTHS_EN[+p[1] - 1] + ' ' + p[0]; }
    return p[2] + '.' + p[1] + '.' + p[0];
  }

  // ---------------------------------------------------------------------------
  // Licence handling (Polar licence keys) and the free allowance.
  // ---------------------------------------------------------------------------
  // Interprets the answer of POST /v1/customer-portal/license-keys/validate.
  // Returns {status: 'valid'|'invalid'|'expired'|'error', checkedAt, expiresAt, detail}
  function interpretLicenceResponse(httpStatus, body, nowIso) {
    var res = { checkedAt: nowIso, expiresAt: null, detail: '' };
    if (httpStatus === 0 || httpStatus === null || httpStatus === undefined) {
      res.status = 'error'; res.detail = 'network'; return res;
    }
    if (httpStatus === 200 && body && typeof body === 'object') {
      res.expiresAt = body.expires_at || null;
      if (body.status !== 'granted') { res.status = 'invalid'; res.detail = String(body.status || 'unknown'); return res; }
      if (res.expiresAt && Date.parse(res.expiresAt) <= Date.parse(nowIso)) { res.status = 'expired'; return res; }
      res.status = 'valid';
      return res;
    }
    if (httpStatus === 404 || httpStatus === 422 || httpStatus === 400) {
      res.status = 'invalid';
      res.detail = body && body.detail ? (typeof body.detail === 'string' ? body.detail : 'validation') : String(httpStatus);
      return res;
    }
    res.status = 'error';
    res.detail = 'http ' + httpStatus;
    return res;
  }
  var LICENCE_FIELDS = ['licenceStatus', 'licenceCheckedAt', 'licenceLastOkAt', 'licenceExpiresAt', 'licenceLastError',
    'licenceLastErrorAt'];
  // Merges a check result into the stored account state. Pure.
  function applyLicenceCheck(state, key, result) {
    var next = {};
    Object.keys(state || {}).forEach(function (k) { next[k] = state[k]; });
    if (!state || state.licenceKey !== key) {
      // A different key starts from nothing and never inherits the state of the previous key.
      LICENCE_FIELDS.forEach(function (k) { delete next[k]; });
    }
    next.licenceKey = key;
    if (result.status === 'error') {
      // A failed check (network, 429, 5xx) is no check: the last result and its time stay, so the next action asks
      // Polar again after licenceRetryMinutes. Meanwhile the grace period decides.
      next.licenceStatus = next.licenceStatus || 'unknown';
      next.licenceLastError = result.detail;
      next.licenceLastErrorAt = result.checkedAt;
      return next;
    }
    next.licenceCheckedAt = result.checkedAt;
    delete next.licenceLastError;
    delete next.licenceLastErrorAt;
    if (result.status === 'valid') {
      next.licenceStatus = 'valid';
      next.licenceLastOkAt = result.checkedAt;
      next.licenceExpiresAt = result.expiresAt;
    } else {
      next.licenceStatus = result.status;
      next.licenceExpiresAt = result.expiresAt;
    }
    return next;
  }
  // Decides whether one more e-invoice may be created. Pure.
  // state: {freeUsed, licenceKey, licenceStatus, licenceCheckedAt, licenceLastOkAt, licenceExpiresAt, licenceLastErrorAt}
  // Returns {allowed, mode: 'licence'|'free'|'blocked', freeLeft, needsCheck, reason, licence}, where licence is
  // 'none' (no key), 'active', 'unverified' (not confirmed by Polar lately), 'invalid' or 'expired'.
  function decideAccess(state, nowIso, cfg) {
    cfg = cfg || CONFIG;
    state = state || {};
    var now = Date.parse(nowIso);
    var freeUsed = state.freeUsed || 0;
    var freeLeft = Math.max(0, cfg.freeInvoices - freeUsed);
    var out = { freeLeft: freeLeft, needsCheck: false, licence: 'none' };
    if (state.licenceKey) {
      var checked = state.licenceCheckedAt ? Date.parse(state.licenceCheckedAt) : 0;
      var failed = state.licenceLastErrorAt ? Date.parse(state.licenceLastErrorAt) : 0;
      out.needsCheck = !checked || (now - checked) > cfg.licenceRecheckHours * 3600000;
      // After a failed check ask again on the next action, but not more often than every few minutes.
      if (out.needsCheck && failed && (now - failed) < cfg.licenceRetryMinutes * 60000) { out.needsCheck = false; }
      var lastOk = state.licenceLastOkAt ? Date.parse(state.licenceLastOkAt) : 0;
      var expired = !!state.licenceExpiresAt && Date.parse(state.licenceExpiresAt) <= now;
      var withinGrace = lastOk && (now - lastOk) <= cfg.licenceGraceDays * 86400000;
      if (state.licenceStatus === 'invalid') { out.licence = 'invalid'; }
      else if (state.licenceStatus === 'expired' || expired) { out.licence = 'expired'; }
      else if (state.licenceStatus === 'valid' && withinGrace) { out.licence = 'active'; }
      else { out.licence = 'unverified'; }
      if (out.licence === 'active') { out.allowed = true; out.mode = 'licence'; return out; }
    }
    if (freeLeft > 0) { out.allowed = true; out.mode = 'free'; return out; }
    out.allowed = false;
    out.mode = 'blocked';
    out.reason = out.licence === 'none' ? 'free-used-up' : 'licence-' + out.licence;
    return out;
  }
  // Cleans a pasted licence key: removes spaces, line breaks and invisible characters inside it and picks an SF- key
  // out of surrounding text such as "Lizenzschlüssel: SF-..." (LIC-5).
  function normalizeLicenceKey(v) {
    var s = String(v === null || v === undefined ? '' : v).replace(/[\s\u200B-\u200D\u2060\uFEFF]/g, '');
    var m = /SF-[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}/i.exec(s);
    return m ? m[0] : s;
  }
  var LICENCE_KEY_MAX = 100;
  function maskKey(key) {
    var s = text(key);
    if (s.length <= 8) { return s ? '****' : ''; }
    return s.slice(0, 4) + '...' + s.slice(-4);
  }

  // ---------------------------------------------------------------------------
  // Example data for "Create template with example" and for the screenshots.
  // ---------------------------------------------------------------------------
  function exampleInvoice() {
    return {
      fields: {
        'doc.type': '380 Rechnung / Invoice',
        'doc.number': 'RE-2026-0042',
        'doc.date': '2026-10-01',
        'doc.deliveryDate': '2026-09-30',
        'doc.dueDate': '2026-10-15',
        'doc.buyerReference': 'Bestellung 4711',
        'doc.orderReference': '4711',
        'seller.name': 'Muster Webdesign GmbH',
        'seller.street': 'Hauptstraße 12',
        'seller.zip': '10115',
        'seller.city': 'Berlin',
        'seller.country': 'DE',
        'seller.vatId': 'DE123456789',
        'seller.taxNumber': '30/123/45678',
        'seller.registration': 'HRB 123456 B',
        'seller.contactName': 'Anna Muster',
        'seller.phone': '+49 30 1234567',
        'seller.email': 'rechnung@muster-webdesign.example',
        'seller.iban': 'DE02120300000000202051',
        'seller.bic': 'BYLADEM1001',
        'seller.accountHolder': 'Muster Webdesign GmbH',
        'buyer.name': 'Bäckerei Schmidt KG',
        'buyer.contactName': 'Jonas Schmidt',
        'buyer.street': 'Marktplatz 3',
        'buyer.zip': '04109',
        'buyer.city': 'Leipzig',
        'buyer.country': 'DE',
        'buyer.vatId': 'DE987654321',
        'buyer.email': 'buchhaltung@baeckerei-schmidt.example',
        'buyer.customerNumber': 'K-1007',
        'payment.terms': 'Zahlbar innerhalb von 14 Tagen ohne Abzug.'
      },
      lines: [
        { description: 'Webseite Relaunch, Konzept und Gestaltung', quantity: 16, unit: 'Stunde (HUR)', price: 95,
          category: 'S Regelsatz / standard', rate: 19 },
        { description: 'Hosting Paket Business, 12 Monate', quantity: 12, unit: 'Monat (MON)', price: 14.9,
          category: 'S Regelsatz / standard', rate: 19 },
        { description: 'Fachbuch "Gutes Webdesign"', quantity: 2, unit: 'Stück (H87)', price: 39.9, discount: 10,
          category: 'S Regelsatz / standard', rate: 7 }
      ]
    };
  }

  return {
    VERSION: VERSION,
    CONFIG: CONFIG,
    configure: configure,
    configPlaceholders: configPlaceholders,
    TEMPLATE_MARKER: TEMPLATE_MARKER,
    TEMPLATE_SHEET_NAME: TEMPLATE_SHEET_NAME,
    SECTIONS: SECTIONS,
    LINE_COLUMNS: LINE_COLUMNS,
    LINE_TEXT_COLUMNS: LINE_TEXT_COLUMNS,
    isTextField: isTextField,
    VAT_CATEGORIES: VAT_CATEGORIES,
    UNITS: UNITS,
    UNIT_CODES: UNIT_CODES,
    DOC_TYPES: DOC_TYPES,
    MESSAGES: MESSAGES,
    // decimals
    add: add, sub: sub, mul: mul, round2: round2, percentOf: percentOf, cmp: cmp, sum: sum,
    normalizeDecimal: normalizeDecimal, numberToDecimalString: numberToDecimalString,
    // parsing and validation helpers
    parseNumber: parseNumber, parseDate: parseDate, isValidIban: isValidIban, normalizeIban: normalizeIban,
    isPlausibleVatId: isPlausibleVatId, isValidLeitwegId: isValidLeitwegId, leitwegShape: leitwegShape,
    normalizeCountry: normalizeCountry, normalizeUnit: normalizeUnit, unitLabel: unitLabel,
    normalizeCategory: normalizeCategory, normalizeDocType: normalizeDocType, isValidEmail: isValidEmail,
    // template
    buildTemplatePlan: buildTemplatePlan, fillTemplate: fillTemplate, docTypeChoices: docTypeChoices,
    categoryChoices: categoryChoices, unitChoices: unitChoices, fieldLabel: fieldLabel, a1: a1,
    // sheet -> invoice
    parseSheet: parseSheet, buildModel: buildModel, checkSheet: checkSheet, computeTotals: computeTotals,
    toInvoiceData: toInvoiceData, outputFormats: outputFormats, outputFileNames: outputFileNames,
    formatMessage: formatMessage, formatMoney: formatMoney, formatDecimal: formatDecimal, formatDate: formatDate,
    paymentTermsText: paymentTermsText, SKONTO_LINE: SKONTO_LINE, splitTerms: splitTerms,
    // licence
    interpretLicenceResponse: interpretLicenceResponse, applyLicenceCheck: applyLicenceCheck,
    decideAccess: decideAccess, maskKey: maskKey, LICENCE_FIELDS: LICENCE_FIELDS, normalizeLicenceKey: normalizeLicenceKey,
    LICENCE_KEY_MAX: LICENCE_KEY_MAX,
    exampleInvoice: exampleInvoice
  };
}));

EInvoiceCore.configure({"appName":"SheetFaktura","polar":{"apiBase":"https://api.polar.sh","organizationId":"00eedf7d-4d83-4bc5-9e39-5acfb6048f97","benefitId":"b2b868d4-929c-43c0-b3f8-5f74f72a5dc1","checkoutMonthly":"https://buy.polar.sh/polar_cl_YuvQl4FoqRVCzLeQuGTKuLV5mJTPJplnoYMmq41KzAB","checkoutYearly":"https://buy.polar.sh/polar_cl_D97vW9iyNG47DrXJIdZMiX7JUPSR9EtlwdG852vKUnj","customerPortal":"https://polar.sh/zemia/portal"},"links":{"help":"https://sheetfaktura.de/hilfe.html","privacy":"https://sheetfaktura.de/datenschutz.html","terms":"https://sheetfaktura.de/agb.html"}});
/*
 * Draws the human readable invoice PDF (A4) with pdf-lib and an embedded TrueType font.
 * The result is later turned into a PDF/A-3 ZUGFeRD file by @e-invoice-eu/core.
 * Works in the browser and in Node: pdf-lib, fontkit and the font bytes are passed in.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) {
    module.exports = api;
  } else {
    root.EInvoicePdf = api;
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var LABELS = {
    de: {
      title: { '380': 'Rechnung', '381': 'Rechnungskorrektur', '326': 'Teilrechnung' },
      number: 'Rechnungsnummer', date: 'Rechnungsdatum', deliveryDate: 'Leistungsdatum', period: 'Leistungszeitraum',
      dueDate: 'Fällig am', customerNumber: 'Kundennummer', buyerReference: 'Ihre Referenz', leitweg: 'Leitweg-ID',
      orderReference: 'Ihre Bestellung', preceding: 'Bezug auf Rechnung', buyerVatId: 'Ihre USt-IdNr.',
      pos: 'Pos.', description: 'Beschreibung', quantity: 'Menge', unit: 'Einheit', price: 'Einzelpreis',
      vat: 'USt', amount: 'Betrag', discountLine: 'abzüglich {p} % Rabatt: {a}', itemNo: 'Art.-Nr. {x}',
      lineTotal: 'Summe Positionen netto', taxExclusive: 'Summe netto', vatLine: 'USt {r} % auf {b}',
      zeroLine: '{c} auf {b}', grandTotal: 'Gesamtbetrag', creditTotal: 'Betrag der Korrektur', prepaid: 'Bereits bezahlt',
      prepaidVat: 'darin enthaltene USt {r} %',
      payable: 'Zahlbetrag', bank: 'Bankverbindung', holder: 'Kontoinhaber', reference: 'Verwendungszweck',
      skonto: 'Bei Zahlung innerhalb von {d} Tagen {p} % Skonto ({a} EUR). Zahlbetrag dann {b} EUR.',
      skontoLine: 'Bei Zahlung innerhalb von {d} Tagen {p} % Skonto.', skontoBase: 'Bei Zahlung innerhalb von {d} Tagen {p} % Skonto auf {b} EUR.',
      page: 'Seite {n} von {m}', vatId: 'USt-IdNr.', taxNumber: 'Steuernr.', registration: 'Register',
      phone: 'Tel.', email: 'E-Mail', to: 'bis', currency: 'EUR', continued: 'Fortsetzung',
      creditNoPayment: 'Diese Rechnungskorrektur ist keine Zahlungsaufforderung.',
      eInvoice: 'Diese PDF-Datei enthält die Rechnungsdaten als E-Rechnung (ZUGFeRD, Profil {p}).'
    },
    en: {
      title: { '380': 'Invoice', '381': 'Credit note', '326': 'Partial invoice' },
      number: 'Invoice number', date: 'Invoice date', deliveryDate: 'Delivery date', period: 'Service period',
      dueDate: 'Due date', customerNumber: 'Customer number', buyerReference: 'Your reference', leitweg: 'Leitweg-ID',
      orderReference: 'Your order', preceding: 'Corrects invoice', buyerVatId: 'Your VAT ID',
      pos: 'No.', description: 'Description', quantity: 'Qty', unit: 'Unit', price: 'Unit price',
      vat: 'VAT', amount: 'Amount', discountLine: 'less {p} % discount: {a}', itemNo: 'Item no. {x}',
      lineTotal: 'Items net total', taxExclusive: 'Net total', vatLine: 'VAT {r} % on {b}',
      zeroLine: '{c} on {b}', grandTotal: 'Total', creditTotal: 'Credit total', prepaid: 'Already paid',
      prepaidVat: 'including VAT {r} %',
      payable: 'Amount due', bank: 'Bank details', holder: 'Account holder', reference: 'Payment reference',
      skonto: 'If paid within {d} days, {p} % cash discount ({a} EUR). Amount due then {b} EUR.',
      skontoLine: 'If paid within {d} days, {p} % cash discount.', skontoBase: 'If paid within {d} days, {p} % cash discount on {b} EUR.',
      page: 'Page {n} of {m}', vatId: 'VAT ID', taxNumber: 'Tax no.', registration: 'Registration',
      phone: 'Phone', email: 'Email', to: 'to', currency: 'EUR', continued: 'continued',
      creditNoPayment: 'This credit note is not a request for payment.',
      eInvoice: 'This PDF file contains the invoice data as an e-invoice (ZUGFeRD, profile {p}).'
    }
  };

  // The template's German default texts, shown in English on English PDFs (PDF-12). Texts the user typed stay as typed.
  var DEFAULTS_EN = {
    'Rabatt': 'Discount',
    'Zahlbar innerhalb von 14 Tagen ohne Abzug.': 'Payable within 14 days without deduction.'
  };

  // Readable stand-ins for common symbols the embedded font lacks (PDF-6). Each stand-in is in the font.
  var FALLBACKS = {
    0x2190: '<-', 0x2192: '->', 0x2194: '<->', 0x21d2: '=>', 0x2248: '~', 0x2260: '!=', 0x2264: '<=', 0x2265: '>=',
    0x2103: '°C', 0x2300: 'Ø'
  };

  function fill(tpl, params) {
    return tpl.replace(/\{(\w+)\}/g, function (_, k) { return params[k] !== undefined ? String(params[k]) : ''; });
  }

  // Replaces characters the embedded font cannot show. PDF/A forbids drawing the .notdef glyph.
  // Text pasted from web pages, Word or macOS carries decomposed accents, special spaces and invisible characters:
  // accents are composed (NFC), special spaces become plain spaces, invisible characters are dropped (PDF-6).
  // With keepBreaks, line feeds typed in a cell stay for wrapText (PDF-13); text drawn as one line never has one.
  function makeSanitizer(charSet) {
    var supported = {};
    for (var k = 0; k < charSet.length; k++) { supported[charSet[k]] = true; }
    return function (s, keepBreaks) {
      var str = String(s === null || s === undefined ? '' : s);
      if (typeof str.normalize === 'function') { str = str.normalize('NFC'); }
      str = str.replace(/\r\n?|[\u0085\u2028\u2029]/g, '\n')
        .replace(/[\u00ad\u180e\u200b-\u200f\u202a-\u202e\u2060-\u2064\u2066-\u206f\ufeff]/g, '')
        .replace(/[\t\u00a0\u1680\u2000-\u200a\u202f\u205f\u3000]/g, ' ');
      if (!keepBreaks) { str = str.replace(/\n+/g, ' '); }
      var out = '';
      for (var i = 0; i < str.length; i++) {
        var cp = str.codePointAt(i);
        if (cp > 0xffff) { i++; }
        if (cp === 10) { out += '\n'; continue; }
        if (cp < 32 || (cp >= 0x7f && cp < 0xa0)) { continue; }
        out += supported[cp] ? String.fromCodePoint(cp) : (FALLBACKS[cp] || '?');
      }
      return out;
    };
  }

  // Longest start of a word that fits maxWidth (at least one character), found by doubling and then halving, so a
  // long run without spaces costs a few measurements per line (PDF-5: shortening one character at a time took
  // seconds to minutes). Breaks after , ; / or - when one is in the second half of that start.
  function fitLength(word, width, maxWidth) {
    var lo = 1;
    var hi = 2;
    while (hi <= word.length && width(word.slice(0, hi)) <= maxWidth) { lo = hi; hi *= 2; }
    if (hi > word.length) {
      if (width(word) <= maxWidth) { return word.length; }
      hi = word.length;
    }
    while (hi - lo > 1) {
      var mid = (lo + hi) >> 1;
      if (width(word.slice(0, mid)) <= maxWidth) { lo = mid; } else { hi = mid; }
    }
    for (var k = lo; k > lo / 2; k--) {
      if (/[,;\/\-]/.test(word.charAt(k - 1))) { return k; }
    }
    return lo;
  }

  // Splits text into lines no wider than maxWidth. A line feed starts a new line; a blank line between paragraphs
  // is kept once (PDF-13).
  function wrapText(textValue, font, size, maxWidth) {
    var width = function (s) { return font.widthOfTextAtSize(s, size); };
    var lines = [];
    String(textValue).split('\n').forEach(function (para) {
      var words = para.split(/\s+/).filter(function (w) { return w.length > 0; });
      if (!words.length) {
        if (lines.length && lines[lines.length - 1] !== '') { lines.push(''); }
        return;
      }
      var current = '';
      words.forEach(function (word) {
        var candidate = current ? current + ' ' + word : word;
        if (width(candidate) <= maxWidth) { current = candidate; return; }
        if (current) { lines.push(current); current = ''; }
        // Hard split a word that is wider than the column.
        while (word) {
          var cut = fitLength(word, width, maxWidth);
          if (cut >= word.length) { break; }
          lines.push(word.slice(0, cut));
          word = word.slice(cut);
        }
        current = word;
      });
      if (current) { lines.push(current); }
    });
    while (lines.length && lines[lines.length - 1] === '') { lines.pop(); }
    return lines.length ? lines : [''];
  }

  // model: from EInvoiceCore.buildModel (with totals). opts: {PDFLib, fontkit, fonts: {regular, bold}, lang, now, core, profileName}
  function renderInvoicePdf(model, opts) {
    var PDFLib = opts.PDFLib;
    var core = opts.core;
    var lang = opts.lang === 'en' ? 'en' : 'de';
    var L = LABELS[lang];
    var t = model.totals;
    var now = opts.now || new Date();
    var rgb = PDFLib.rgb;
    var INK = rgb(0.11, 0.12, 0.14);
    var GREY = rgb(0.42, 0.43, 0.46);
    var ACCENT = rgb(0.10, 0.33, 0.55);
    var RULE = rgb(0.78, 0.80, 0.83);
    var money = function (v) { return core.formatMoney(v, lang); };
    // Quantities and unit prices with all their decimals, as in the XML (PDF-8).
    var qty = function (v) { return core.formatDecimal(v, lang, 0, 15); };
    var unitPrice = function (v) { return core.formatDecimal(v, lang, 2, 15); };
    var pct = function (v) { return core.formatDecimal(v, lang, 0, 2); };
    var date = function (v) { return core.formatDate(v, lang); };
    var own = function (v) { return lang === 'en' && DEFAULTS_EN[v] ? DEFAULTS_EN[v] : v; };
    var abs = function (v) { return String(v).replace(/^-/, ''); };

    return PDFLib.PDFDocument.create().then(function (doc) {
      doc.registerFontkit(opts.fontkit);
      var charSet = opts.fontkit.create(opts.fonts.regular).characterSet;
      var clean = makeSanitizer(charSet);
      return Promise.all([
        doc.embedFont(opts.fonts.regular, { subset: false }),
        doc.embedFont(opts.fonts.bold, { subset: false })
      ]).then(function (fonts) {
        var regular = fonts[0];
        var bold = fonts[1];
        var W = 595.28;
        var H = 841.89;
        var left = 62;
        var right = W - 50;
        var pages = [];
        var page;
        var y;
        var s = model.seller;
        var b = model.buyer;
        var kind = L.title[model.type] || L.title['380'];
        var title = kind + ' ' + model.number;

        function text(str, x, yy, size, font, color) {
          var c = clean(str);
          if (c) { page.drawText(c, { x: x, y: yy, size: size, font: font || regular, color: color || INK }); }
        }
        function textRight(str, xRight, yy, size, font, color) {
          var c = clean(str);
          var f = font || regular;
          if (c) { page.drawText(c, { x: xRight - f.widthOfTextAtSize(c, size), y: yy, size: size, font: f, color: color || INK }); }
        }
        function rule(x1, x2, yy, color, thickness) {
          page.drawLine({ start: { x: x1, y: yy }, end: { x: x2, y: yy }, thickness: thickness || 0.6, color: color || RULE });
        }
        function wrapped(str, font, size, width) { return wrapText(clean(str, true), font, size, width); }
        function fits(str, font, size, width) { return font.widthOfTextAtSize(clean(str), size) <= width; }
        // Font size at which a number fits its column: the normal size, or smaller down to 5 pt (PDF-8).
        function fitSize(str, size, width) {
          var w = regular.widthOfTextAtSize(clean(str), size);
          return w <= width ? size : Math.max(5, Math.floor(size * width / w * 10) / 10);
        }

        // Footer lines are known first: their number sets how far down the content may go (PDF-2).
        var FOOT = 7.5;
        var footWidth = right - left - regular.widthOfTextAtSize(fill(L.page, { n: 99, m: 99 }), FOOT) - 12;
        var ids = [];
        if (s.vatId && model.categories.indexOf('O') < 0) { ids.push(L.vatId + ' ' + s.vatId); }
        if (s.taxNumber) { ids.push(L.taxNumber + ' ' + s.taxNumber); }
        if (s.registration) { ids.push(L.registration + ' ' + s.registration); }
        var footLines = wrapped([s.name, s.street, s.zip + ' ' + s.city].join(', '), regular, FOOT, footWidth);
        if (ids.length) { footLines = footLines.concat(wrapped(ids.join(', '), regular, FOOT, footWidth)); }
        var footTop = Math.max(50, 20 + (footLines.length - 1) * 10);
        var bottom = footTop + 42;
        var capacity = H - 100 - bottom;

        function newPage() {
          page = doc.addPage([W, H]);
          pages.push(page);
          y = H - 56;
          if (pages.length > 1) {
            // Continuation pages name the invoice, so a page that got separated can be matched (PDF-13).
            wrapped(title + ' (' + L.continued + ')', bold, 9, right - left).forEach(function (ln) { text(ln, left, y, 9, bold, GREY); y -= 11; });
            y -= 14;
          }
          return page;
        }

        newPage();

        // Seller header. Long names and addresses wrap instead of running off the page (PDF-2).
        wrapped(s.name, bold, 15, right - left).forEach(function (ln, i) {
          if (i) { y -= 17; }
          text(ln, left, y, 15, bold, ACCENT);
        });
        y -= 15;
        wrapped([s.street, s.zip + ' ' + s.city, s.country !== 'DE' ? s.country : ''].filter(Boolean).join(', '), regular, 9, right - left)
          .forEach(function (ln) { text(ln, left, y, 9, regular, GREY); y -= 11; });
        var contact = [L.phone + ' ' + s.phone, L.email + ' ' + s.email];
        var contactLines = fits(contact.join('   '), regular, 9, right - left) ? [contact.join('   ')]
          : wrapped(contact[0], regular, 9, right - left).concat(wrapped(contact[1], regular, 9, right - left));
        contactLines.forEach(function (ln, i) {
          if (i) { y -= 11; }
          text(ln, left, y, 9, regular, GREY);
        });
        rule(left, right, y - 8, ACCENT, 1.2);

        // Recipient block (left) and invoice facts (right).
        var blockTop = Math.min(H - 150, y - 68);
        y = blockTop;
        var sender = wrapped(s.name + ', ' + s.street + ', ' + s.zip + ' ' + s.city, regular, 7, right - left);
        sender.forEach(function (ln, i) { text(ln, left, y + 14 + (sender.length - 1 - i) * 8, 7, regular, GREY); });
        var addr = [b.name, b.contactName, b.street, b.zip + ' ' + b.city, b.country !== 'DE' ? b.country : ''];
        addr.filter(Boolean).forEach(function (line) {
          wrapped(line, regular, 10.5, 230).forEach(function (part) { text(part, left, y, 10.5); y -= 13.5; });
        });
        var facts = [];
        facts.push([L.number, model.number]);
        facts.push([L.date, date(model.issueDate)]);
        if (model.deliveryDate) { facts.push([L.deliveryDate, date(model.deliveryDate)]); }
        if (model.periodStart && model.periodEnd) { facts.push([L.period, date(model.periodStart) + ' ' + L.to + ' ' + date(model.periodEnd)]); }
        if (model.dueDate) { facts.push([L.dueDate, date(model.dueDate)]); }
        if (b.customerNumber) { facts.push([L.customerNumber, b.customerNumber]); }
        facts.push([model.leitwegId ? L.leitweg : L.buyerReference, model.buyerReference]);
        if (model.orderReference) { facts.push([L.orderReference, model.orderReference]); }
        if (model.precedingInvoice) {
          facts.push([L.preceding, model.precedingInvoice + (model.precedingInvoiceDate ? ' (' + date(model.precedingInvoiceDate) + ')' : '')]);
        }
        if (b.vatId && model.categories.indexOf('O') < 0) { facts.push([L.buyerVatId, b.vatId]); }
        var fy = blockTop;
        var factLabelX = 330;
        facts.forEach(function (f) {
          text(f[0], factLabelX, fy, 8.5, regular, GREY);
          var lines = wrapped(f[1], regular, 9, right - 430);
          lines.forEach(function (ln, i) { textRight(ln, right, fy - i * 11, 9); });
          fy -= 12.5 + (lines.length - 1) * 11;
        });
        y = Math.min(y, fy) - 26;

        // Starts a new page when less than height is left above the footer. withHeader repeats the table header.
        function ensureSpace(height, withHeader) {
          if (y - height < bottom) {
            newPage();
            if (withHeader) { tableHeader(); }
            return true;
          }
          return false;
        }
        // Draws lines one below the other and goes on on a new page at the footer (PDF-4).
        function drawLines(lines, x, size, font, color, step, withHeader) {
          lines.forEach(function (ln) {
            if (y < bottom) {
              newPage();
              if (withHeader) { tableHeader(); }
            }
            text(ln, x, y, size, font, color);
            y -= step;
          });
        }

        // Title.
        wrapped(title, bold, 16, right - left).forEach(function (ln, i) {
          if (i) { y -= 19; }
          text(ln, left, y, 16, bold);
        });
        y -= 20;
        if (model.note) {
          drawLines(wrapped(model.note, regular, 9.5, right - left), left, 9.5, regular, INK, 12.5, false);
          y -= 4;
        }

        // Line items table.
        var col = {
          pos: left, desc: left + 24, qtyRight: left + 292, unit: left + 298, priceRight: right - 94,
          vatRight: right - 62, amountRight: right
        };
        var descWidth = col.qtyRight - 50 - col.desc;
        function tableHeader() {
          rule(left, right, y + 11, RULE, 0.6);
          text(L.pos, col.pos, y, 8.5, bold, GREY);
          text(L.description, col.desc, y, 8.5, bold, GREY);
          textRight(L.quantity, col.qtyRight, y, 8.5, bold, GREY);
          text(L.unit, col.unit, y, 8.5, bold, GREY);
          textRight(L.price, col.priceRight, y, 8.5, bold, GREY);
          textRight(L.vat, col.vatRight, y, 8.5, bold, GREY);
          textRight(L.amount, col.amountRight, y, 8.5, bold, GREY);
          y -= 6;
          rule(left, right, y, RULE, 0.6);
          y -= 13;
        }
        // The header never stands alone at the foot of a page.
        ensureSpace(40, false);
        tableHeader();
        model.lines.forEach(function (l, idx) {
          var lt = t.lines[idx];
          var descLines = wrapped(l.description, regular, 9.5, descWidth);
          var extra = [];
          if (l.itemNo) { extra.push(fill(L.itemNo, { x: l.itemNo })); }
          // "abzüglich" already says the discount is deducted, also on returned items with a negative amount (PDF-11).
          if (lt.allowance) { extra.push(fill(L.discountLine, { p: pct(l.discount), a: money(abs(lt.allowance)) })); }
          var extraLines = [];
          extra.forEach(function (e) { extraLines = extraLines.concat(wrapped(e, regular, 8, descWidth)); });
          var height = descLines.length * 12 + extraLines.length * 10 + 6;
          // A row stays on one page when it can; a row taller than a page goes on over the next pages (PDF-4).
          ensureSpace(height <= capacity ? height : 18, true);
          text(String(l.n), col.pos, y, 9.5);
          var q = qty(l.quantity);
          textRight(q, col.qtyRight, y, fitSize(q, 9.5, col.qtyRight - col.desc - descWidth - 4));
          var unitText = core.unitLabel(l.unit, lang, true);
          text(unitText, col.unit, y, 9.5);
          var p = unitPrice(l.price);
          textRight(p, col.priceRight, y, fitSize(p, 9.5, col.priceRight - col.unit - regular.widthOfTextAtSize(clean(unitText), 9.5) - 4));
          // Lines that are not standard rated show their VAT category code, which the VAT summary names (PDF-10).
          textRight(l.category === 'S' ? pct(l.rate) + ' %' : l.category, col.vatRight, y, 9.5);
          var amount = money(lt.net);
          textRight(amount, col.amountRight, y, fitSize(amount, 9.5, col.amountRight - col.vatRight - 4));
          drawLines(descLines, col.desc, 9.5, regular, INK, 12, true);
          y += 1;
          drawLines(extraLines, col.desc, 8, regular, GREY, 10, true);
          y -= 7;
        });
        rule(left, right, y + 8, RULE, 0.6);
        y -= 10;

        // Totals.
        var rows = [];
        rows.push([L.lineTotal, money(t.lineTotal), false]);
        t.allowances.forEach(function (a) {
          var which = a.category === 'S' ? pct(a.rate) + ' %' : core.VAT_CATEGORIES[a.category][lang];
          var label = own(a.reason) + ' ' + pct(a.percent) + ' %' + (t.vat.length > 1 ? ' (' + which + ')' : '');
          rows.push([label, '-' + money(a.amount), false]);
        });
        if (t.allowances.length) { rows.push([L.taxExclusive, money(t.taxExclusive), false]); }
        t.vat.forEach(function (v) {
          if (v.category === 'S') {
            rows.push([fill(L.vatLine, { r: pct(v.rate), b: money(v.taxable) }), money(v.tax), false]);
          } else {
            rows.push([fill(L.zeroLine, { c: core.VAT_CATEGORIES[v.category][lang] + ' (' + v.category + ')', b: money(v.taxable) }), money(v.tax), false]);
          }
        });
        rows.push([model.type === '381' ? L.creditTotal : L.grandTotal, money(t.taxInclusive) + ' ' + L.currency, true]);
        if (t.prepaid) {
          rows.push([L.prepaid, '-' + money(t.prepaid), false]);
          if (t.prepaidVat) { rows.push([fill(L.prepaidVat, { r: pct(t.prepaidVat.rate) }), money(t.prepaidVat.amount), false]); }
          rows.push([L.payable, money(t.payable) + ' ' + L.currency, true]);
        }
        var labelX = left + 200;
        // A long label (a long discount reason) wraps left of its amount instead of running over it (PDF-3).
        var moreLines = 0;
        rows.forEach(function (r) {
          var size = r[2] ? 10 : 9.5;
          var font = r[2] ? bold : regular;
          r.push(wrapped(r[0], font, size, right - labelX - font.widthOfTextAtSize(clean(r[1]), size) - 10));
          moreLines += r[3].length - 1;
        });
        ensureSpace(rows.length * 14 + 10 + moreLines * 11.5, false);
        rows.forEach(function (r) {
          var size = r[2] ? 10 : 9.5;
          var font = r[2] ? bold : regular;
          if (r[2]) { rule(labelX, right, y + 10, INK, 0.8); }
          r[3].forEach(function (ln, k) { text(ln, labelX, y - k * 11.5, size, font); });
          textRight(r[1], right, y, size, font);
          y -= (r[2] ? 16 : 13.5) + (r[3].length - 1) * 11.5;
        });
        y -= 8;

        // VAT notes, payment terms, bank details.
        var paras = [];
        t.vat.forEach(function (v) {
          if (v.category === 'S' || v.category === 'Z') { return; }
          var cat = core.VAT_CATEGORIES[v.category];
          var reason = v.category === 'E' ? v.reason : (v.category === 'O' && model.exemptionReason ? model.exemptionReason : (lang === 'en' ? cat.textEn : cat.textDe));
          reason = String(reason || '').trim();
          // A reason typed as a sentence keeps its own full stop (PDF-11).
          paras.push({ text: /[.!?]$/.test(reason) ? reason : reason + '.', bold: true });
        });
        if (model.payment.terms) {
          // Skonto lines in the XRechnung syntax are shown in words (INV-3).
          var termLines = String(model.payment.terms).split(/\r?\n/).map(function (line) {
            var m = /^#SKONTO#TAGE=([0-9]+)#PROZENT=([0-9]+\.[0-9]{2})(?:#BASISBETRAG=(-?[0-9]+\.[0-9]{2}))?#$/.exec(line.trim());
            if (!m) { return own(line.trim()); }
            return fill(m[3] ? L.skontoBase : L.skontoLine, { d: m[1], p: pct(m[2]), b: m[3] ? money(m[3]) : '' });
          });
          paras.push({ text: termLines.join('\n') });
        }
        if (t.skonto) {
          paras.push({ text: fill(L.skonto, { d: t.skonto.days, p: pct(t.skonto.percent), a: money(t.skonto.amount), b: money(t.skonto.payableAfter) }) });
        }
        if (model.type === '381') {
          // A credit note asks for no payment: no bank details and no payment reference (PDF-7).
          paras.push({ text: L.creditNoPayment });
        } else {
          var bank = L.bank + ': ' + [s.accountHolder ? L.holder + ' ' + s.accountHolder : '', 'IBAN ' + formatIban(s.iban),
            s.bic ? 'BIC ' + s.bic : ''].filter(Boolean).join(', ');
          paras.push({ text: bank });
          paras.push({ text: L.reference + ': ' + model.payment.reference });
        }
        paras.push({ text: fill(L.eInvoice, { p: opts.profileName || 'EN 16931' }), grey: true });
        paras.forEach(function (p) {
          var font = p.bold ? bold : regular;
          var lines = wrapped(p.text, font, 9, right - left);
          if (lines.length * 12 + 4 <= capacity) { ensureSpace(lines.length * 12 + 4, false); }
          drawLines(lines, left, 9, font, p.grey ? GREY : INK, 12, false);
          y -= 4;
        });

        // Footer on every page.
        pages.forEach(function (pg, i) {
          page = pg;
          rule(left, right, footTop + 12, RULE, 0.6);
          footLines.forEach(function (ln, k) { text(ln, left, footTop - k * 10, FOOT, regular, GREY); });
          textRight(fill(L.page, { n: i + 1, m: pages.length }), right, footTop, FOOT, regular, GREY);
        });

        doc.setTitle(clean(title));
        doc.setAuthor(clean(s.name));
        doc.setSubject(clean(title + ', ' + b.name));
        doc.setCreator(opts.appName || 'E-Rechnung');
        doc.setProducer(opts.appName || 'E-Rechnung');
        doc.setLanguage(lang === 'en' ? 'en-GB' : 'de-DE');
        doc.setCreationDate(now);
        doc.setModificationDate(now);
        return doc.save({ useObjectStreams: false });
      });
    });
  }

  function formatIban(iban) { return String(iban).replace(/(.{4})/g, '$1 ').trim(); }

  return { renderInvoicePdf: renderInvoicePdf, LABELS: LABELS, wrapText: wrapText, makeSanitizer: makeSanitizer };
}));

/*
 * Creates the two output files from a checked model:
 *   1. XRechnung 3.0 (CII syntax) XML
 *   2. ZUGFeRD / Factur-X PDF/A-3 (EN 16931 profile, or XRechnung profile when a Leitweg-ID is present)
 * Runs the same way in the browser (sidebar) and in Node (tests). All libraries are passed in.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) {
    module.exports = api;
  } else {
    root.EInvoiceGenerate = api;
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var PROFILE_NAMES = { 'Factur-X-EN16931': 'EN 16931', 'Factur-X-XRechnung': 'XRechnung' };

  function quietLogger(sink) {
    return {
      log: function () {},
      info: function () {},
      debug: function () {},
      warn: function (m) { if (sink) { sink.push(String(m)); } },
      error: function (m) { if (sink) { sink.push(String(m)); } }
    };
  }

  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  // CII requires ram:ApplicableHeaderTradeDelivery (it may be empty). The library leaves it out
  // when an invoice has a service period but no delivery date, which the CII schema rejects.
  // This hook runs on the CII object before it is written, for the XML and for the PDF.
  function ensureHeaderTradeDelivery(cii) {
    var root = cii && cii['rsm:CrossIndustryInvoice'];
    var tx = root && root['rsm:SupplyChainTradeTransaction'];
    if (tx && !tx['ram:ApplicableHeaderTradeDelivery']) {
      var rebuilt = {};
      Object.keys(tx).forEach(function (k) {
        if (k === 'ram:ApplicableHeaderTradeSettlement') { rebuilt['ram:ApplicableHeaderTradeDelivery'] = {}; }
        rebuilt[k] = tx[k];
      });
      root['rsm:SupplyChainTradeTransaction'] = rebuilt;
    }
    return Promise.resolve();
  }

  // The library's browser bundle calls this.getCrypto(require) when it builds the PDF/A file.
  // In a browser the bare name "require" is undefined and throws before getCrypto can pick
  // window.crypto.subtle. A stub that exists only while we generate lets it fall through to
  // the Web Crypto API. Node is unaffected (the bundle sees its own module require there).
  function withRequireShim(run) {
    var isBrowser = typeof window !== 'undefined' && !!window.document;
    if (!isBrowser || typeof window.require !== 'undefined') { return run(); }
    window.require = function () { throw new Error('require is not available in the browser'); };
    var cleanup = function () { try { delete window.require; } catch (e) { window.require = undefined; } };
    return run().then(function (r) { cleanup(); return r; }, function (e) { cleanup(); throw e; });
  }

  // libs: {core, pdf, eInvoiceEU, PDFLib, fontkit}; fonts: {regular, bold} as Uint8Array
  // options: {lang: 'de'|'en', now: Date, appName, skipPdf: bool, invoiceDataOverride: fn(data)}
  function generateFiles(model, libs, fonts, options) {
    return withRequireShim(function () { return generateFilesInner(model, libs, fonts, options); });
  }

  function generateFilesInner(model, libs, fonts, options) {
    options = options || {};
    var core = libs.core;
    var warnings = [];
    var service = new libs.eInvoiceEU.InvoiceService(quietLogger(warnings));
    var formats = core.outputFormats(model);
    var names = core.outputFileNames(model);
    var data = core.toInvoiceData(model);
    if (options.invoiceDataOverride) { data = options.invoiceDataOverride(clone(data)) || data; }
    var lang = options.lang === 'en' ? 'en' : 'de';
    return service.generate(clone(data), { format: formats.xml, lang: 'de-de', noWarnings: true,
      postProcessor: ensureHeaderTradeDelivery }).then(function (xml) {
      if (options.skipPdf) {
        return { xml: xml, pdf: null, names: names, formats: formats, warnings: warnings, invoiceData: data };
      }
      return libs.pdf.renderInvoicePdf(model, {
        PDFLib: libs.PDFLib, fontkit: libs.fontkit, fonts: fonts, lang: lang, now: options.now, core: core,
        appName: options.appName || core.CONFIG.appName, profileName: PROFILE_NAMES[formats.pdf]
      }).then(function (visual) {
        return service.generate(clone(data), {
          format: formats.pdf,
          lang: lang === 'en' ? 'en-us' : 'de-de',
          noWarnings: true,
          postProcessor: ensureHeaderTradeDelivery,
          pdf: { buffer: visual, filename: names.pdf, mimetype: 'application/pdf' }
        });
      }).then(function (pdf) {
        return { xml: xml, pdf: pdf, names: names, formats: formats, warnings: warnings, invoiceData: data };
      });
    });
  }

  return { generateFiles: generateFiles, PROFILE_NAMES: PROFILE_NAMES, ensureHeaderTradeDelivery: ensureHeaderTradeDelivery };
}));

/*
 * Reads an Excel file (.xlsx) for "SheetFaktura für Excel", in the browser and in the Node tests.
 * No spreadsheet library: fflate unzips the file, a small XML scanner reads the parts it needs.
 *
 * Handled: shared and inline strings (rich text runs, phonetic runs left out), numbers, booleans,
 * error values (#N/A ...), formula cells through their cached values, ISO dates (t="d"), serial
 * dates told apart from numbers by the cell's number format, the 1904 date system, sparse rows and
 * cells without "r" references. Refused with a code: empty, too large, encrypted (password),
 * old .xls, OpenDocument, Numbers, .xlsb, other ZIP or non-ZIP files, damaged files.
 *
 * sheetValues() returns rows in the shape the add-on's server hands to core.checkSheet
 * (src/server/Code.js, getInvoiceSheetData): 10 columns, "" for empty cells, dates as
 * "YYYY-MM-DD", and numbers in text fields as the cell shows them (postcode 4109 formatted
 * 00000 becomes "04109").
 *
 * Syntax stays at ES2017 like the rest of src/.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) {
    module.exports = api;
  } else {
    root.SheetFakturaXlsx = api;
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var LIMITS = {
    fileBytes: 5 * 1024 * 1024, // the file as picked
    partBytes: 40 * 1024 * 1024, // one unpacked XML part
    totalBytes: 80 * 1024 * 1024, // all unpacked parts together
    rows: 20000 // rows of the invoice sheet
  };
  var READ_COLUMNS = 10; // A to J, like the add-on
  var MARKER_SCAN_ROWS = 50; // where the template marker or the keys are looked for, like the add-on
  var STOP = {};

  function fail(code, detail) {
    var e = new Error(code + (detail ? ': ' + detail : ''));
    e.code = code;
    e.isXlsxError = true;
    return e;
  }

  // ---------------------------------------------------------------------------
  // File type
  // ---------------------------------------------------------------------------
  function startsWith(bytes, sig) {
    if (bytes.length < sig.length) { return false; }
    for (var i = 0; i < sig.length; i++) { if (bytes[i] !== sig[i]) { return false; } }
    return true;
  }
  // An ASCII name stored as UTF-16LE, as in the directory of a compound file.
  function containsUtf16(bytes, ascii) {
    var first = ascii.charCodeAt(0);
    var len = ascii.length * 2;
    for (var i = 0; i + len <= bytes.length; i++) {
      if (bytes[i] !== first || bytes[i + 1] !== 0) { continue; }
      var k = 1;
      for (; k < ascii.length; k++) {
        if (bytes[i + 2 * k] !== ascii.charCodeAt(k) || bytes[i + 2 * k + 1] !== 0) { break; }
      }
      if (k === ascii.length) { return true; }
    }
    return false;
  }
  var ZIP = [0x50, 0x4b, 0x03, 0x04];
  var ZIP_EMPTY = [0x50, 0x4b, 0x05, 0x06];
  var CFB = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  // 'zip', 'encrypted' (an Excel file with a password to open is an encrypted package inside a compound
  // file), 'cfb' (old .xls and other Office 97-2003 files) or 'unknown'.
  function sniff(bytes) {
    if (startsWith(bytes, ZIP) || startsWith(bytes, ZIP_EMPTY)) { return 'zip'; }
    if (startsWith(bytes, CFB)) {
      return (containsUtf16(bytes, 'EncryptedPackage') || containsUtf16(bytes, 'EncryptionInfo')) ? 'encrypted' : 'cfb';
    }
    return 'unknown';
  }

  // ---------------------------------------------------------------------------
  // Text and XML
  // ---------------------------------------------------------------------------
  function decodeText(bytes) {
    var enc = 'utf-8';
    var start = 0;
    if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) { enc = 'utf-16le'; start = 2; }
    else if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) { enc = 'utf-16be'; start = 2; }
    else if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) { start = 3; }
    if (typeof TextDecoder === 'function') {
      return new TextDecoder(enc).decode(start ? bytes.subarray(start) : bytes);
    }
    throw fail('CORRUPT', 'no TextDecoder');
  }
  var NAMED = { lt: '<', gt: '>', amp: '&', quot: '"', apos: '\'' };
  function codePoint(cp) {
    try { return String.fromCodePoint(cp); } catch (e) { return '�'; }
  }
  function decodeEntities(s) {
    if (s.indexOf('&') < 0) { return s; }
    return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|lt|gt|amp|quot|apos);/g, function (m, e) {
      if (e.charAt(0) === '#') { return codePoint(e.charAt(1) === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)); }
      return NAMED[e];
    });
  }
  // Office escapes characters XML cannot hold as _xHHHH_ (and a literal "_x" as _x005F_x).
  function decodeOfficeEscapes(s) {
    if (s.indexOf('_x') < 0) { return s; }
    return s.replace(/_x([0-9A-Fa-f]{4})_/g, function (m, h) { return String.fromCharCode(parseInt(h, 16)); });
  }
  function localName(n) {
    var i = n.indexOf(':');
    return i < 0 ? n : n.slice(i + 1);
  }
  function parseAttrs(s) {
    var out = {};
    var re = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    var m;
    while ((m = re.exec(s))) { out[localName(m[1])] = decodeEntities(m[2] !== undefined ? m[2] : m[3]); }
    return out;
  }
  // Calls h.open(name, attrs, selfClosing), h.close(name) and h.text(raw) with local names (namespace
  // prefixes dropped). A handler that returns STOP ends the scan. Raw text still holds its entities.
  function scanXml(s, h) {
    var i = 0;
    var n = s.length;
    while (i < n) {
      var lt = s.indexOf('<', i);
      if (lt < 0) {
        if (h.text) { h.text(s.slice(i)); }
        return;
      }
      if (lt > i && h.text && h.text(s.slice(i, lt)) === STOP) { return; }
      var c = s.charAt(lt + 1);
      var end;
      if (c === '?') {
        end = s.indexOf('?>', lt + 2);
        if (end < 0) { throw fail('CORRUPT', 'xml'); }
        i = end + 2;
        continue;
      }
      if (c === '!') {
        if (s.substr(lt, 4) === '<!--') {
          end = s.indexOf('-->', lt + 4);
          if (end < 0) { throw fail('CORRUPT', 'xml'); }
          i = end + 3;
          continue;
        }
        if (s.substr(lt, 9) === '<![CDATA[') {
          end = s.indexOf(']]>', lt + 9);
          if (end < 0) { throw fail('CORRUPT', 'xml'); }
          // CDATA is literal: protect its ampersands from the entity decoding of the caller.
          if (h.text && h.text(s.slice(lt + 9, end).replace(/&/g, '&amp;')) === STOP) { return; }
          i = end + 3;
          continue;
        }
        // Office parts never carry a document type; refusing it rules out entity tricks.
        throw fail('CORRUPT', 'doctype');
      }
      var gt = s.indexOf('>', lt + 1);
      if (gt < 0) { throw fail('CORRUPT', 'xml'); }
      var q1 = s.indexOf('"', lt + 1);
      var q2 = s.indexOf('\'', lt + 1);
      if ((q1 >= 0 && q1 < gt) || (q2 >= 0 && q2 < gt)) {
        // A ">" may stand inside a quoted attribute value: find the end outside quotes.
        var quote = 0;
        for (gt = lt + 1; gt < n; gt++) {
          var ch = s.charCodeAt(gt);
          if (quote) { if (ch === quote) { quote = 0; } } else if (ch === 34 || ch === 39) { quote = ch; } else if (ch === 62) { break; }
        }
        if (gt >= n) { throw fail('CORRUPT', 'xml'); }
      }
      var inner = s.slice(lt + 1, gt);
      i = gt + 1;
      if (inner.charAt(0) === '/') {
        if (h.close && h.close(localName(inner.slice(1).trim())) === STOP) { return; }
        continue;
      }
      var selfClosing = inner.charAt(inner.length - 1) === '/';
      if (selfClosing) { inner = inner.slice(0, -1); }
      var sp = inner.search(/\s/);
      var name = localName(sp < 0 ? inner : inner.slice(0, sp));
      var attrs = sp < 0 ? {} : parseAttrs(inner.slice(sp));
      if (h.open && h.open(name, attrs, selfClosing) === STOP) { return; }
      if (selfClosing && h.close && h.close(name) === STOP) { return; }
    }
  }

  // ---------------------------------------------------------------------------
  // Package parts
  // ---------------------------------------------------------------------------
  function dirOf(path) { var i = path.lastIndexOf('/'); return i < 0 ? '' : path.slice(0, i + 1); }
  function resolvePath(baseDir, target) {
    var t = String(target || '').replace(/\\/g, '/');
    var parts = (t.charAt(0) === '/' ? t.slice(1) : baseDir + t).split('/');
    var out = [];
    parts.forEach(function (p) {
      if (p === '..') { out.pop(); } else if (p !== '.' && p !== '') { out.push(p); }
    });
    return out.join('/');
  }
  function relsPathFor(partPath) {
    var d = dirOf(partPath);
    return d + '_rels/' + partPath.slice(d.length) + '.rels';
  }
  function parseRels(xml) {
    var rels = [];
    scanXml(xml, {
      open: function (name, a) {
        if (name === 'Relationship') {
          rels.push({ id: a.Id || a.id || '', type: a.Type || a.type || '', target: a.Target || a.target || '', external: a.TargetMode === 'External' });
        }
      }
    });
    return rels;
  }
  function relOfType(rels, suffix) {
    for (var i = 0; i < rels.length; i++) {
      if (!rels[i].external && rels[i].type.slice(-suffix.length) === suffix) { return rels[i]; }
    }
    return null;
  }

  // Unpacks the named parts (exact or case-insensitive match), checking the declared sizes first.
  function unpacker(bytes, fflate) {
    var listing = {};
    var lower = {};
    try {
      fflate.unzipSync(bytes, { filter: function (f) { listing[f.name] = f; return false; } });
    } catch (e) {
      throw fail('CORRUPT', e && e.message);
    }
    Object.keys(listing).forEach(function (name) { lower[name.toLowerCase()] = name; });
    var unpackedTotal = 0;
    function find(path) {
      if (listing[path]) { return path; }
      return lower[path.toLowerCase()] || null;
    }
    function get(paths) {
      var wanted = {};
      paths.forEach(function (p) { var real = find(p); if (real) { wanted[real] = p; } });
      var names = Object.keys(wanted);
      if (!names.length) { return {}; }
      names.forEach(function (name) {
        var size = listing[name].originalSize;
        if (size > LIMITS.partBytes) { throw fail('TOO_LARGE', name); }
        unpackedTotal += size;
        if (unpackedTotal > LIMITS.totalBytes) { throw fail('TOO_LARGE', 'unpacked'); }
      });
      var files;
      try {
        files = fflate.unzipSync(bytes, { filter: function (f) { return !!wanted[f.name]; } });
      } catch (e) {
        throw fail('CORRUPT', e && e.message);
      }
      var out = {};
      names.forEach(function (name) {
        if (!files[name]) { throw fail('CORRUPT', 'missing ' + name); }
        out[wanted[name]] = decodeText(files[name]);
      });
      return out;
    }
    return { names: Object.keys(listing), find: find, get: get };
  }

  // ---------------------------------------------------------------------------
  // Workbook, shared strings, styles
  // ---------------------------------------------------------------------------
  function parseWorkbookXml(xml) {
    var wb = { sheets: [], activeTab: 0, date1904: false };
    scanXml(xml, {
      open: function (name, a) {
        if (name === 'sheet') {
          wb.sheets.push({ name: a.name || '', relId: a.id || '', state: a.state || 'visible' });
        } else if (name === 'workbookPr') {
          wb.date1904 = a.date1904 === '1' || a.date1904 === 'true';
        } else if (name === 'workbookView' && a.activeTab !== undefined) {
          wb.activeTab = parseInt(a.activeTab, 10) || 0;
        }
      }
    });
    return wb;
  }
  function parseSharedStrings(xml) {
    var list = [];
    var cur = null;
    var inT = false;
    var phonetic = 0;
    var buf = '';
    scanXml(xml, {
      open: function (name, a, selfClosing) {
        if (name === 'si') { cur = ''; } else if (name === 'rPh') { phonetic += 1; } else if (name === 't' && cur !== null && !phonetic && !selfClosing) { inT = true; buf = ''; }
      },
      text: function (raw) { if (inT) { buf += raw; } },
      close: function (name) {
        if (name === 't' && inT) { inT = false; cur += decodeOfficeEscapes(decodeEntities(buf)); } else if (name === 'rPh') { phonetic -= 1; } else if (name === 'si') { list.push(cur === null ? '' : cur); cur = null; }
      }
    });
    return list;
  }
  var BUILTIN_FORMATS = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00',
    12: '# ?/?', 13: '# ??/??', 14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@' };
  // Built-in ids that are dates or times; 27 to 36 and 50 to 58 are the East Asian date formats.
  function isBuiltinDateId(id) {
    return (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58);
  }
  function cleanFormat(code) {
    return String(code || '').replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/_./g, '').replace(/\*./g, '');
  }
  function isDateFormatCode(code) {
    var s = cleanFormat(code);
    if (/\[(h+|m+|s+)\]/i.test(s)) { return true; } // elapsed time
    s = s.replace(/\[[^\]]*\]/g, '').replace(/general/gi, ''); // colours, conditions, locale; the word General
    return /[dmyhs]/i.test(s);
  }
  // "00000": a number shown with leading zeros. Returns the width or 0.
  function zeroPadWidth(code) {
    var first = cleanFormat(code).replace(/\[[^\]]*\]/g, '').split(';')[0].trim();
    return /^0+$/.test(first) ? first.length : 0;
  }
  function parseStyles(xml) {
    var custom = {};
    var xfs = [];
    var inCellXfs = false;
    scanXml(xml, {
      open: function (name, a) {
        if (name === 'numFmt' && a.numFmtId !== undefined) { custom[parseInt(a.numFmtId, 10)] = a.formatCode || ''; }
        else if (name === 'cellXfs') { inCellXfs = true; }
        else if (name === 'xf' && inCellXfs) { xfs.push(parseInt(a.numFmtId || '0', 10) || 0); }
      },
      close: function (name) { if (name === 'cellXfs') { inCellXfs = false; } }
    });
    var cache = {};
    return function formatOf(styleIndex) {
      var id = xfs[styleIndex] || 0;
      if (!cache[id]) {
        var code = custom[id] !== undefined ? custom[id] : (BUILTIN_FORMATS[id] || 'General');
        cache[id] = { id: id, code: code, isDate: custom[id] !== undefined ? isDateFormatCode(code) : isBuiltinDateId(id), zeroPad: zeroPadWidth(code) };
      }
      return cache[id];
    };
  }

  // ---------------------------------------------------------------------------
  // Worksheets
  // ---------------------------------------------------------------------------
  function colIndex(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) { n = n * 26 + (letters.charCodeAt(i) - 64); }
    return n;
  }
  // Raw cells of a worksheet: rows[r][c] = {t, v, s} for columns 1 to READ_COLUMNS. opts.maxRow stops early.
  function parseWorksheet(xml, opts) {
    opts = opts || {};
    var rows = {};
    var lastRow = 0;
    var row = 0;
    var col = 0;
    var cell = null;
    var buf = null;
    var inIs = false;
    var inT = false;
    var phonetic = 0;
    var isBuf = '';
    var tooMany = false;
    scanXml(xml, {
      open: function (name, a, selfClosing) {
        if (name === 'row') {
          row = a.r ? parseInt(a.r, 10) : row + 1;
          col = 0;
          if (opts.maxRow && row > opts.maxRow) { return STOP; }
          if (row > LIMITS.rows) { tooMany = true; return STOP; }
        } else if (name === 'c') {
          var ref = a.r ? /^([A-Z]+)(\d+)$/.exec(String(a.r).toUpperCase()) : null;
          if (ref) {
            col = colIndex(ref[1]);
            var refRow = parseInt(ref[2], 10);
            if (refRow !== row) {
              row = refRow;
              if (opts.maxRow && row > opts.maxRow) { return STOP; }
              if (row > LIMITS.rows) { tooMany = true; return STOP; }
            }
          } else {
            col += 1;
          }
          cell = { t: a.t || 'n', s: a.s ? parseInt(a.s, 10) || 0 : 0, v: null, is: null };
          if (selfClosing) { cell = null; }
        } else if (cell && name === 'v') {
          buf = '';
        } else if (cell && name === 'is') {
          inIs = true;
          isBuf = '';
        } else if (inIs && name === 'rPh') {
          phonetic += 1;
        } else if (inIs && name === 't' && !phonetic) {
          inT = true;
        }
        return undefined;
      },
      text: function (raw) {
        if (buf !== null) { buf += raw; } else if (inT) { isBuf += raw; }
      },
      close: function (name) {
        if (name === 'v' && buf !== null) {
          if (cell) { cell.v = decodeEntities(buf); }
          buf = null;
        } else if (name === 't' && inT) {
          inT = false;
        } else if (name === 'rPh' && inIs) {
          phonetic -= 1;
        } else if (name === 'is' && inIs) {
          inIs = false;
          if (cell) { cell.is = decodeOfficeEscapes(decodeEntities(isBuf)); }
        } else if (name === 'c' && cell) {
          if (col >= 1 && col <= READ_COLUMNS) {
            (rows[row] = rows[row] || [])[col] = cell;
            if (row > lastRow) { lastRow = row; }
          }
          cell = null;
        } else if (name === 'sheetData') {
          return STOP;
        }
        return undefined;
      }
    });
    if (tooMany) { throw fail('TOO_MANY_ROWS'); }
    return { rows: rows, lastRow: lastRow };
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function serialToYmd(serial, date1904) {
    var days = Math.floor(serial) + (date1904 ? 1462 : 0);
    // Day 0 is 30 December 1899; right from 1 March 1900 on (Excel counts a 29 February 1900 that never was).
    var d = new Date(Date.UTC(1899, 11, 30) + days * 86400000);
    if (isNaN(d.getTime())) { return String(serial); }
    var y = d.getUTCFullYear();
    return (y < 1000 ? ('000' + y).slice(-4) : String(y)) + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
  }
  // A raw cell as the add-on's server would pass it. textColumn: the template reads this cell as text.
  function cellValue(cell, wb, textColumn) {
    if (!cell) { return ''; }
    var v = cell.v;
    switch (cell.t) {
      case 's': {
        var idx = parseInt(v, 10);
        var s = wb.shared[idx];
        return s === undefined ? '' : s;
      }
      case 'inlineStr': return cell.is !== null ? cell.is : (v === null ? '' : decodeOfficeEscapes(v));
      case 'str': return v === null ? '' : decodeOfficeEscapes(v);
      case 'b': return v === null ? '' : (v === '1' || v === 'true');
      case 'e': return v === null ? '' : v;
      case 'd': {
        if (v === null) { return ''; }
        var m = /^(\d{4}-\d{2}-\d{2})/.exec(v);
        return m ? m[1] : v;
      }
      default: {
        if (v === null || v.trim() === '') { return ''; }
        var num = Number(v);
        if (!isFinite(num)) { return v; }
        var fmt = wb.formatOf(cell.s);
        if (fmt.isDate) { return serialToYmd(num, wb.date1904); }
        if (textColumn && fmt.zeroPad && num >= 0 && num === Math.floor(num)) {
          var digits = String(num);
          while (digits.length < fmt.zeroPad) { digits = '0' + digits; }
          return digits;
        }
        return num;
      }
    }
  }

  // ---------------------------------------------------------------------------
  // Public API
  // ---------------------------------------------------------------------------
  // bytes: Uint8Array of the file. Returns a workbook handle; worksheets are parsed on demand.
  function readWorkbook(bytes, fflate) {
    if (!bytes || !bytes.length) { throw fail('EMPTY'); }
    if (bytes.length > LIMITS.fileBytes) { throw fail('TOO_LARGE'); }
    var kind = sniff(bytes);
    if (kind === 'encrypted') { throw fail('ENCRYPTED'); }
    if (kind === 'cfb') { throw fail('OLD_FORMAT'); }
    if (kind !== 'zip') { throw fail('NOT_XLSX'); }
    var pkg = unpacker(bytes, fflate);
    var names = pkg.names;
    var has = function (p) { return !!pkg.find(p); };
    if (has('mimetype') && has('content.xml')) { throw fail('ODS'); }
    if (names.some(function (n) { return /^Index\/.*\.iwa$/i.test(n) || /^Index\.zip$/i.test(n); })) { throw fail('NUMBERS'); }
    if (has('xl/workbook.bin')) { throw fail('XLSB'); }
    // A package with spreadsheet parts that cannot be read is damaged, not some other kind of file.
    var spreadsheetParts = names.some(function (n) { return /^xl\//i.test(n); });
    var notReadable = function () { return fail(spreadsheetParts ? 'CORRUPT' : 'NOT_XLSX'); };
    if (!has('[Content_Types].xml') || !has('_rels/.rels')) { throw notReadable(); }

    var root = parseRels(pkg.get(['_rels/.rels'])['_rels/.rels']);
    var officeDoc = relOfType(root, '/officeDocument');
    if (!officeDoc) { throw notReadable(); }
    var wbPath = resolvePath('', officeDoc.target);
    if (!/\.xml$/i.test(wbPath) || !has(wbPath)) { throw notReadable(); }
    var wbRelsPath = relsPathFor(wbPath);
    var first = pkg.get([wbPath, wbRelsPath]);
    if (first[wbPath] === undefined) { throw notReadable(); }
    var info = parseWorkbookXml(first[wbPath]);
    if (!/<(\w+:)?sheets[\s>]/.test(first[wbPath])) { throw notReadable(); }
    var rels = first[wbRelsPath] !== undefined ? parseRels(first[wbRelsPath]) : [];
    var base = dirOf(wbPath);
    var relById = {};
    rels.forEach(function (r) { relById[r.id] = r; });

    var sharedRel = relOfType(rels, '/sharedStrings');
    var stylesRel = relOfType(rels, '/styles');
    var sharedPath = sharedRel ? resolvePath(base, sharedRel.target) : null;
    var stylesPath = stylesRel ? resolvePath(base, stylesRel.target) : null;
    var second = pkg.get([sharedPath, stylesPath].filter(Boolean));

    var wb = {
      date1904: info.date1904,
      activeTab: info.activeTab,
      shared: sharedPath && second[sharedPath] !== undefined ? parseSharedStrings(second[sharedPath]) : [],
      formatOf: stylesPath && second[stylesPath] !== undefined ? parseStyles(second[stylesPath]) : parseStyles(''),
      sheets: [],
      _pkg: pkg,
      _xml: {},
      _parsed: {}
    };
    info.sheets.forEach(function (s, i) {
      var rel = relById[s.relId];
      var isWorksheet = rel && /\/worksheet$/.test(rel.type);
      wb.sheets.push({ index: i, name: s.name, hidden: s.state !== 'visible', path: isWorksheet ? resolvePath(base, rel.target) : null });
    });
    if (!wb.sheets.some(function (s) { return s.path; })) { throw notReadable(); }
    return wb;
  }

  function sheetXml(wb, index) {
    var s = wb.sheets[index];
    if (!s || !s.path) { return null; }
    if (wb._xml[index] === undefined) {
      var got = wb._pkg.get([s.path]);
      if (got[s.path] === undefined) { throw fail('CORRUPT', 'missing ' + s.path); }
      wb._xml[index] = got[s.path];
    }
    return wb._xml[index];
  }
  function parsedSheet(wb, index) {
    if (!wb._parsed[index]) {
      var xml = sheetXml(wb, index);
      wb._parsed[index] = xml === null ? { rows: {}, lastRow: 0 } : parseWorksheet(xml);
    }
    return wb._parsed[index];
  }

  // Invoice sheets: the template marker or template keys in column A of the first rows (the add-on's rule).
  function findInvoiceSheets(wb, core) {
    var marker = core.TEMPLATE_MARKER;
    var found = [];
    wb.sheets.forEach(function (s) {
      if (!s.path) { return; }
      var head = wb._parsed[s.index] || parseWorksheet(sheetXml(wb, s.index), { maxRow: MARKER_SCAN_ROWS });
      for (var r = 1; r <= Math.min(head.lastRow, MARKER_SCAN_ROWS); r++) {
        var cells = head.rows[r];
        var a = cells ? cellValue(cells[1], wb, true) : '';
        var key = typeof a === 'string' ? a.trim() : '';
        if (key === marker || key === '#lines' || key.indexOf('#section:') === 0) {
          found.push({ index: s.index, name: s.name, hidden: s.hidden });
          return;
        }
      }
    });
    return found;
  }
  // The sheet open when the file was saved, if it is an invoice sheet; else the first visible one.
  function defaultInvoiceSheet(wb, found) {
    if (!found.length) { return null; }
    for (var i = 0; i < found.length; i++) { if (found[i].index === wb.activeTab) { return found[i].index; } }
    for (var k = 0; k < found.length; k++) { if (!found[k].hidden) { return found[k].index; } }
    return found[0].index;
  }

  // Rows 1 to the last used row, 10 columns, the way src/server/Code.js reads a sheet.
  function sheetValues(wb, index, core) {
    var sheet = parsedSheet(wb, index);
    var values = [];
    var inLines = false;
    var lastUsed = 0;
    for (var r = 1; r <= sheet.lastRow; r++) {
      var cells = sheet.rows[r] || [];
      var keyValue = cellValue(cells[1], wb, true);
      var key = String(keyValue);
      if (key === '#lines') { inLines = true; } else if (key === '#lines-end') { inLines = false; }
      var textCols = key === '#lines' ? [] : (inLines ? core.LINE_TEXT_COLUMNS : (core.isTextField(key) ? [3] : []));
      var row = [];
      for (var c = 1; c <= READ_COLUMNS; c++) {
        var v = c === 1 ? keyValue : cellValue(cells[c], wb, textCols.indexOf(c) >= 0);
        row.push(v);
        if (v !== '' && v !== null && v !== undefined) { lastUsed = r; }
      }
      values.push(row);
    }
    values.length = lastUsed;
    return values;
  }

  return {
    LIMITS: LIMITS,
    READ_COLUMNS: READ_COLUMNS,
    readWorkbook: readWorkbook,
    findInvoiceSheets: findInvoiceSheets,
    defaultInvoiceSheet: defaultInvoiceSheet,
    sheetValues: sheetValues,
    // exposed for the tests
    sniff: sniff,
    scanXml: scanXml,
    isDateFormatCode: isDateFormatCode,
    zeroPadWidth: zeroPadWidth,
    serialToYmd: serialToYmd
  };
}));

/*
 * "SheetFaktura für Excel": the tool on sheetfaktura.de/excel.html.
 * Reads the picked .xlsx in the browser (SheetFakturaXlsx), runs the add-on's checks (EInvoiceCore),
 * makes the XRechnung XML and the ZUGFeRD PDF with the add-on's code (EInvoiceGenerate, EInvoicePdf) and
 * offers both as downloads. Nothing is uploaded.
 * Free allowance and licence follow the add-on (src/sidebar/sidebar.js, src/server/Code.js); the state lives in
 * this browser's localStorage instead of the user's Google account. Every storage access is guarded: when the
 * browser refuses storage, the state lasts until the page is closed.
 * Injected by scripts/build-web.js before this file: SF_WEB (library files with SRI hashes), EInvoiceCore
 * (configured from deploy.config.json), EInvoicePdf, EInvoiceGenerate, SheetFakturaXlsx. fflate loads with the page.
 */
(function () {
  'use strict';
  var core = window.EInvoiceCore;
  var X = window.SheetFakturaXlsx;
  var CFG = core.CONFIG;
  var WEB = window.SF_WEB || { libs: [] };

  // ---------------------------------------------------------------------------
  // Texts. German page texts are in the HTML; PAGE_EN holds their English versions by data-i18n key.
  // ---------------------------------------------------------------------------
  var STR = {
    de: {
      title: 'E-Rechnung aus Excel erstellen: XRechnung und ZUGFeRD im Browser | SheetFaktura',
      dropTitle: 'Excel-Datei hierher ziehen', dropOr: 'oder', pick: 'Datei auswählen',
      dropHint: 'Nur .xlsx, höchstens 5 MB. Die Datei bleibt auf Ihrem Gerät.',
      reading: 'Datei wird gelesen ...', fileInfo: 'Datei: {file}', sheetInfo: 'Blatt: {sheet}',
      sheetChoose: 'Die Datei hat mehrere Rechnungsblätter. Bitte das richtige wählen:',
      checkOk: 'Alles in Ordnung.', checkErrors: '{n} Problem(e). Bitte in Excel korrigieren, die Datei speichern und hier erneut auswählen:',
      checkWarnings: 'Hinweise:', cell: 'Zelle {c}',
      lines: 'Positionen', net: 'Summe netto', vat: 'Umsatzsteuer', gross: 'Gesamtbetrag', payable: 'Zahlbetrag',
      formats: 'Ausgabe: XRechnung 3.0 (XML) und ZUGFeRD, Profil {p} (PDF)',
      pdfLang: 'Sprache der PDF:', createBtn: 'E-Rechnung erstellen', creating: 'E-Rechnung wird erstellt ...',
      loadingLibs: 'Programmteile werden geladen ...', created: 'Fertig. Bitte beide Dateien speichern:',
      dlXml: 'XRechnung speichern (XML)', dlPdf: 'ZUGFeRD speichern (PDF)', again: 'Andere oder geänderte Datei auswählen',
      dropAgain: 'Datei in Excel geändert? Hierher ziehen oder neu auswählen.',
      storageOff: 'Ihr Browser erlaubt dieser Seite keine Speicherung. Zähler und Lizenzschlüssel gelten deshalb nur, bis Sie die Seite schließen.',
      toTemplate: 'Zur Vorlage',
      errEmpty: 'Die Datei ist leer.',
      errTooLarge: 'Die Datei ist größer als 5 MB. Bitte das Rechnungsblatt in einer eigenen Datei speichern, zum Beispiel in einer neuen Datei aus der Vorlage.',
      errEncrypted: 'Die Datei ist mit einem Kennwort geschützt. Bitte in Excel den Kennwortschutz entfernen (Datei, Informationen, Arbeitsmappe schützen, Mit Kennwort verschlüsseln, Kennwort löschen), speichern und erneut auswählen.',
      errOld: 'Das ist eine ältere Excel-Datei (.xls) oder ein anderes Format. Bitte in Excel mit Datei, Speichern unter als „Excel-Arbeitsmappe (.xlsx)“ speichern und erneut auswählen.',
      errOds: 'Das ist eine OpenDocument-Datei (.ods). Bitte in LibreOffice mit Datei, Speichern unter als „Excel 2007-365 (.xlsx)“ speichern und erneut auswählen.',
      errNumbers: 'Das ist eine Numbers-Datei. Bitte in Numbers mit Ablage, Exportieren, Excel als .xlsx exportieren und erneut auswählen.',
      errXlsb: 'Das ist eine Excel-Binärdatei (.xlsb). Bitte als „Excel-Arbeitsmappe (.xlsx)“ speichern und erneut auswählen.',
      errNotXlsx: 'Diese Datei ist keine Excel-Datei im Format .xlsx.',
      errCorrupt: 'Die Datei lässt sich nicht lesen. Vielleicht ist sie beschädigt oder unvollständig. Bitte in Excel öffnen, erneut als .xlsx speichern und noch einmal auswählen.',
      errTooManyRows: 'Das Rechnungsblatt hat mehr als 20.000 Zeilen. Bitte die überzähligen Zeilen löschen.',
      errNoTemplate: 'Diese Datei ist keine SheetFaktura-Vorlage. Bitte die Vorlage herunterladen, Ihre Angaben in die gelben Felder übertragen und die Datei hier auswählen.',
      errGeneric: 'Etwas ist schiefgelaufen. Bitte erneut versuchen.',
      errLibs: 'Programmteile konnten nicht geladen werden. Bitte die Internetverbindung prüfen und erneut versuchen. Auch ein Werbeblocker oder eine Firewall kann das Laden verhindern.',
      fixFirst: 'Bitte zuerst die Probleme beheben, die Datei speichern und hier erneut auswählen.',
      noFile: 'Bitte zuerst eine Excel-Datei auswählen.',
      licTitle: 'Lizenz', freeLeft: 'Kostenlos: noch {n} von {m} E-Rechnungen', freeUsedUp: 'Die {m} kostenlosen E-Rechnungen sind verbraucht.',
      freeWhere: 'Die kostenlosen E-Rechnungen werden in diesem Browser gezählt.',
      price: '7,90 EUR im Monat (monatlich kündbar) oder 69 EUR im Jahr, jeweils zzgl. USt. Keine Begrenzung der Anzahl. Der Schlüssel gilt hier und im Add-on für Google Sheets.',
      buyMonthly: 'Monatlich', buyYearly: 'Jährlich', keyPlaceholder: 'Lizenzschlüssel', keyLabel: 'Lizenzschlüssel', activate: 'Aktivieren',
      checkingKey: 'Schlüssel wird geprüft ...', licActive: 'Lizenz aktiv', licKey: 'Schlüssel {k}, zuletzt geprüft am {d}.',
      licChecking: 'Lizenz wird geprüft ...', licStored: 'Schlüssel {k}', licUnverifiedPill: 'Lizenz nicht bestätigt',
      licUnverified: 'Die Lizenz konnte gerade nicht bei unserem Zahlungsanbieter Polar (api.polar.sh) geprüft werden. Vielleicht blockiert ein Werbeblocker oder eine Firewall die Adresse. Bitte in ein paar Minuten auf „Erneut prüfen“ klicken.',
      licGrace: 'Die letzte Prüfung bei Polar war nicht möglich. Die Lizenz gilt bis zu 7 Tage nach der letzten erfolgreichen Prüfung weiter.',
      licRetry: 'Erneut prüfen', licInvalid: 'Dieser Lizenzschlüssel ist ungültig, gesperrt oder abgelaufen.',
      licManage: 'Abo verwalten', licRemove: 'Schlüssel entfernen', licNotSetUp: 'Die Lizenzprüfung ist noch nicht eingerichtet.',
      licKeyMissing: 'Bitte Lizenzschlüssel eingeben.', licRemoveConfirm: 'Schlüssel wirklich entfernen?', licRemoveYes: 'Ja, entfernen',
      licRemoveNo: 'Abbrechen', paywall: 'Für weitere E-Rechnungen bitte eine Lizenz kaufen oder Ihren Schlüssel eingeben.',
      loading: 'Wird geladen ...'
    },
    en: {
      title: 'Create e-invoices from Excel: XRechnung and ZUGFeRD in the browser | SheetFaktura',
      dropTitle: 'Drag the Excel file here', dropOr: 'or', pick: 'Choose file',
      dropHint: 'Only .xlsx, at most 5 MB. The file stays on your device.',
      reading: 'Reading the file ...', fileInfo: 'File: {file}', sheetInfo: 'Sheet: {sheet}',
      sheetChoose: 'The file has several invoice sheets. Please pick the right one:',
      checkOk: 'All good.', checkErrors: '{n} problem(s). Please fix them in Excel, save the file and pick it here again:',
      checkWarnings: 'Notes:', cell: 'Cell {c}',
      lines: 'Line items', net: 'Net total', vat: 'VAT', gross: 'Total', payable: 'Amount due',
      formats: 'Output: XRechnung 3.0 (XML) and ZUGFeRD, profile {p} (PDF)',
      pdfLang: 'PDF language:', createBtn: 'Create e-invoice', creating: 'Creating the e-invoice ...',
      loadingLibs: 'Loading components ...', created: 'Done. Please save both files:',
      dlXml: 'Save XRechnung (XML)', dlPdf: 'Save ZUGFeRD (PDF)', again: 'Pick another or changed file',
      dropAgain: 'Changed the file in Excel? Drag it here or pick it again.',
      storageOff: 'Your browser does not let this page store data. The counter and the licence key therefore only last until you close the page.',
      toTemplate: 'Go to the template',
      errEmpty: 'The file is empty.',
      errTooLarge: 'The file is larger than 5 MB. Please save the invoice sheet in a file of its own, for example a new file from the template.',
      errEncrypted: 'The file is protected with a password. Please remove the password in Excel (File, Info, Protect Workbook, Encrypt with Password, clear the password), save and pick it again.',
      errOld: 'This is an older Excel file (.xls) or another format. Please save it in Excel with File, Save As as "Excel Workbook (.xlsx)" and pick it again.',
      errOds: 'This is an OpenDocument file (.ods). Please save it in LibreOffice with File, Save As as "Excel 2007-365 (.xlsx)" and pick it again.',
      errNumbers: 'This is a Numbers file. Please export it in Numbers with File, Export To, Excel as .xlsx and pick it again.',
      errXlsb: 'This is an Excel binary workbook (.xlsb). Please save it as "Excel Workbook (.xlsx)" and pick it again.',
      errNotXlsx: 'This file is not an Excel file in .xlsx format.',
      errCorrupt: 'The file cannot be read. It may be damaged or incomplete. Please open it in Excel, save it again as .xlsx and pick it again.',
      errTooManyRows: 'The invoice sheet has more than 20,000 rows. Please delete the extra rows.',
      errNoTemplate: 'This file is not a SheetFaktura template. Please download the template, copy your data into the yellow fields and pick the file here.',
      errGeneric: 'Something went wrong. Please try again.',
      errLibs: 'Components could not be loaded. Please check the internet connection and try again. An ad blocker or a firewall can also block loading.',
      fixFirst: 'Please fix the problems first, save the file and pick it here again.',
      noFile: 'Please pick an Excel file first.',
      licTitle: 'Licence', freeLeft: 'Free: {n} of {m} e-invoices left', freeUsedUp: 'The {m} free e-invoices are used up.',
      freeWhere: 'The free e-invoices are counted in this browser.',
      price: '7.90 EUR a month (cancel monthly) or 69 EUR a year, plus VAT. No limit on the number of invoices. The key works here and in the Google Sheets add-on.',
      buyMonthly: 'Monthly', buyYearly: 'Yearly', keyPlaceholder: 'Licence key', keyLabel: 'Licence key', activate: 'Activate',
      checkingKey: 'Checking the key ...', licActive: 'Licence active', licKey: 'Key {k}, last checked on {d}.',
      licChecking: 'Checking the licence ...', licStored: 'Key {k}', licUnverifiedPill: 'Licence not confirmed',
      licUnverified: 'The licence could not be checked with our payment provider Polar (api.polar.sh) right now. An ad blocker or a firewall may block the address. Please click "Check again" in a few minutes.',
      licGrace: 'The last check with Polar was not possible. The licence stays valid for up to 7 days after the last successful check.',
      licRetry: 'Check again', licInvalid: 'This licence key is invalid, revoked or expired.',
      licManage: 'Manage subscription', licRemove: 'Remove key', licNotSetUp: 'The licence check is not set up yet.',
      licKeyMissing: 'Please enter the licence key.', licRemoveConfirm: 'Really remove the key?', licRemoveYes: 'Yes, remove',
      licRemoveNo: 'Cancel', paywall: 'To create more e-invoices, please buy a licence or enter your key.',
      loading: 'Loading ...'
    }
  };
  // Error codes of SheetFakturaXlsx and their texts. Every one of them offers the template.
  var FILE_ERRORS = { EMPTY: 'errEmpty', TOO_LARGE: 'errTooLarge', ENCRYPTED: 'errEncrypted', OLD_FORMAT: 'errOld', ODS: 'errOds',
    NUMBERS: 'errNumbers', XLSB: 'errXlsb', NOT_XLSX: 'errNotXlsx', CORRUPT: 'errCorrupt', TOO_MANY_ROWS: 'errTooManyRows',
    NO_TEMPLATE: 'errNoTemplate' };
  // Core messages that name buttons of the add-on get web wording.
  core.MESSAGES.WEB_REQUIRED_ROW_MISSING = {
    de: 'Die Zeile "{label}" fehlt im Blatt. Bitte die Löschung rückgängig machen oder die Vorlage neu herunterladen.',
    en: 'The row "{label}" is missing from the sheet. Please undo the deletion or download the template again.'
  };
  core.MESSAGES.WEB_NO_TEMPLATE = { de: STR.de.errNoTemplate, en: STR.en.errNoTemplate };
  var WEB_CODES = { REQUIRED_ROW_MISSING: 'WEB_REQUIRED_ROW_MISSING', NO_TEMPLATE: 'WEB_NO_TEMPLATE' };

  // ---------------------------------------------------------------------------
  // Storage (localStorage, every access guarded)
  // ---------------------------------------------------------------------------
  var STATE_KEY = 'sheetfaktura.web.account.v1';
  var LANG_KEY = 'sheetfaktura.lang';
  var memory = { account: null, lang: null };
  var storageOk = true;
  function storageGet(key) {
    try {
      return window.localStorage.getItem(key);
    } catch (e) {
      storageOk = false;
      return null;
    }
  }
  function storageSet(key, value) {
    try {
      window.localStorage.setItem(key, value);
      return true;
    } catch (e) {
      storageOk = false;
      return false;
    }
  }
  function readState() {
    var s = null;
    if (storageOk) {
      var raw = storageGet(STATE_KEY);
      if (raw) {
        try { s = JSON.parse(raw); } catch (e) { s = null; }
      }
    }
    if (!storageOk && memory.account) { s = memory.account; }
    if (!s || typeof s !== 'object') { s = { freeUsed: 0 }; }
    if (typeof s.freeUsed !== 'number' || !(s.freeUsed >= 0)) { s.freeUsed = 0; }
    return s;
  }
  function writeState(s) {
    memory.account = s;
    storageSet(STATE_KEY, JSON.stringify(s));
    return s;
  }
  // The same rules as saveLicenceCheck, removeLicence and recordGeneration in src/server/Code.js.
  function saveLicenceCheck(key, result) {
    var clean = core.normalizeLicenceKey(key);
    if (!clean || clean.length > core.LICENCE_KEY_MAX) { throw new Error('Invalid licence key'); }
    var allowed = { valid: true, invalid: true, expired: true, error: true };
    if (!result || !allowed[result.status]) { throw new Error('Invalid licence check result'); }
    return writeState(core.applyLicenceCheck(readState(), clean, {
      status: result.status,
      checkedAt: String(result.checkedAt || new Date().toISOString()),
      expiresAt: result.expiresAt || null,
      detail: String(result.detail || '').slice(0, 200)
    }));
  }
  function removeLicence() {
    var s = readState();
    ['licenceKey'].concat(core.LICENCE_FIELDS).forEach(function (k) { delete s[k]; });
    return writeState(s);
  }
  function recordGeneration(mode) {
    var s = readState();
    if (mode !== 'licence') { s.freeUsed = (s.freeUsed || 0) + 1; }
    return writeState(s);
  }

  // ---------------------------------------------------------------------------
  // State and helpers
  // ---------------------------------------------------------------------------
  // German unless English was chosen (?lang=en or the switch). The page does not follow the browser language,
  // so search engines always see the German page.
  function initialLang() {
    var m = /[?&]lang=(de|en)\b/.exec(window.location.search || '');
    if (m) { return m[1]; }
    var saved = storageGet(LANG_KEY);
    return saved === 'en' ? 'en' : 'de';
  }
  var state = { lang: initialLang(), file: null, wb: null, sheets: [], sheetIndex: null, check: null, fileError: null, files: null,
    busy: false, notice: null, progress: null, account: readState(), licenceMsg: null, licenceChecking: false, keyDraft: null,
    confirmRemove: false, dragOver: false };

  function t(key, params) {
    var s = (STR[state.lang] && STR[state.lang][key]) || STR.de[key] || key;
    return s.replace(/\{(\w+)\}/g, function (_, k) { return params && params[k] !== undefined ? String(params[k]) : ''; });
  }
  function esc(s) {
    return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', '\'': '&#39;' }[c];
    });
  }
  function $(id) { return document.getElementById(id); }
  function nowIso() { return new Date().toISOString(); }
  function localYmd(d) {
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  }
  // No server: technical details go to the console only.
  function logError(err) {
    try { window.console.error('[SheetFaktura]', err && err.stack ? err.stack : err); } catch (e) { /* ignore */ }
  }

  // ---------------------------------------------------------------------------
  // Libraries: same-origin files pinned with SRI hashes, loaded on demand.
  // ---------------------------------------------------------------------------
  var libsPromise = null;
  function loadScript(lib) {
    return new Promise(function (resolve, reject) {
      if (window[lib.global]) { resolve(); return; }
      var s = document.createElement('script');
      var timer = window.setTimeout(function () { fail(new Error('Timed out loading ' + lib.src)); }, CFG.libTimeoutMs);
      function fail(err) {
        window.clearTimeout(timer);
        s.onload = null;
        s.onerror = null;
        if (s.parentNode) { s.parentNode.removeChild(s); }
        reject(err);
      }
      s.src = lib.src;
      s.integrity = lib.integrity;
      s.crossOrigin = 'anonymous';
      s.async = false;
      s.onload = function () {
        window.clearTimeout(timer);
        if (window[lib.global]) { resolve(); } else { fail(new Error(lib.global + ' missing')); }
      };
      s.onerror = function () { fail(new Error('Could not load ' + lib.src)); };
      document.head.appendChild(s);
    });
  }
  function loadLibs() {
    if (libsPromise) { return libsPromise; }
    libsPromise = (WEB.libs || []).reduce(function (p, lib) {
      return p.then(function () { return loadScript(lib); });
    }, Promise.resolve()).then(function () {
      return { core: core, pdf: window.EInvoicePdf, generate: window.EInvoiceGenerate, eInvoiceEU: window.eInvoiceEU,
        PDFLib: window.PDFLib, fontkit: window.fontkit };
    });
    libsPromise.catch(function () { libsPromise = null; });
    return libsPromise;
  }
  var fontCache = null;
  function fonts() {
    if (!fontCache) {
      var dec = function (b64) {
        var bin = window.atob(b64);
        var out = new Uint8Array(bin.length);
        for (var i = 0; i < bin.length; i++) { out[i] = bin.charCodeAt(i); }
        return out;
      };
      fontCache = { regular: dec(window.EINVOICE_FONTS.regular), bold: dec(window.EINVOICE_FONTS.bold) };
    }
    return fontCache;
  }

  // ---------------------------------------------------------------------------
  // Licence (Polar), as in the add-on's sidebar
  // ---------------------------------------------------------------------------
  function licenceConfigured() { return core.configPlaceholders().indexOf('polar.organizationId') < 0; }
  function polarValidate(key) {
    var p = CFG.polar;
    var body = { key: key, organization_id: p.organizationId };
    if (p.benefitId) { body.benefit_id = p.benefitId; }
    // A request that hangs or is blocked counts as a network error: the grace period decides (LIC-3).
    var ctrl = typeof AbortController === 'function' ? new AbortController() : null;
    var timer = ctrl ? window.setTimeout(function () { ctrl.abort(); }, CFG.polarTimeoutMs) : null;
    var done = function (result) { if (timer) { window.clearTimeout(timer); } return result; };
    var request;
    try {
      request = window.fetch(p.apiBase + '/v1/customer-portal/license-keys/validate', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
        signal: ctrl ? ctrl.signal : undefined, credentials: 'omit', cache: 'no-store'
      });
    } catch (e) {
      request = Promise.reject(e);
    }
    return request.then(function (res) {
      return res.json().catch(function () { return null; }).then(function (b) { return done(core.interpretLicenceResponse(res.status, b, nowIso())); });
    }, function () {
      return done(core.interpretLicenceResponse(0, null, nowIso()));
    });
  }
  function licenceCheckDue() {
    var a = state.account || {};
    return !!a.licenceKey && licenceConfigured() && core.decideAccess(a, nowIso()).needsCheck;
  }
  var licenceCheck = null;
  function refreshLicenceIfNeeded() {
    if (licenceCheck) { return licenceCheck; }
    state.account = readState();
    if (!licenceCheckDue()) { return Promise.resolve(); }
    var key = state.account.licenceKey;
    licenceCheck = polarValidate(key).then(function (result) {
      state.account = saveLicenceCheck(key, result);
      licenceCheck = null;
    }, function (err) { licenceCheck = null; throw err; });
    return licenceCheck;
  }

  // ---------------------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------------------
  // The page texts change only when the language does. The German ones come from the HTML and are kept for
  // switching back; the English ones are SF_PAGE_EN (fixed texts of the build, no user data).
  var pageDe = null;
  var pageLang = 'de';
  function renderPage() {
    var i18n = document.querySelectorAll('[data-i18n]');
    var links = document.querySelectorAll('[data-href-en]');
    if (!pageDe) {
      pageDe = { title: document.title, texts: [], hrefs: [] };
      Array.prototype.forEach.call(i18n, function (el, i) { pageDe.texts[i] = el.innerHTML; });
      Array.prototype.forEach.call(links, function (el, i) { pageDe.hrefs[i] = el.getAttribute('href'); });
    }
    if (state.lang === pageLang) { return; }
    var en = state.lang === 'en';
    var english = window.SF_PAGE_EN || {};
    Array.prototype.forEach.call(i18n, function (el, i) {
      var key = el.getAttribute('data-i18n');
      el.innerHTML = en && english[key] !== undefined ? english[key] : pageDe.texts[i];
    });
    Array.prototype.forEach.call(links, function (el, i) {
      el.setAttribute('href', en ? el.getAttribute('data-href-en') : pageDe.hrefs[i]);
    });
    document.title = en ? STR.en.title : pageDe.title;
    document.documentElement.lang = state.lang;
    Array.prototype.forEach.call(document.querySelectorAll('.sf-lang button'), function (b) {
      b.setAttribute('aria-pressed', b.getAttribute('data-lang') === state.lang ? 'true' : 'false');
    });
    pageLang = state.lang;
  }
  function setNotice(kind, key, params) {
    state.notice = key ? { kind: kind, key: key, params: params || null } : null;
    renderNotice();
    var n = $('sf-notice');
    if (key && n && n.scrollIntoView) { n.scrollIntoView({ block: 'nearest' }); }
  }
  function renderNotice() {
    var n = $('sf-notice');
    var x = state.notice;
    if (!x) { n.hidden = true; n.textContent = ''; return; }
    n.hidden = false;
    n.className = 'sf-notice ' + x.kind;
    n.textContent = t(x.key, x.params);
  }
  // Once a file is read the drop zone shrinks to a bar for picking the changed file again.
  function renderDrop() {
    var loaded = !!state.file;
    $('sf-drop-title').textContent = t(loaded ? 'dropAgain' : 'dropTitle');
    $('sf-drop-or').textContent = t('dropOr');
    $('sf-pick').textContent = t(loaded ? 'again' : 'pick');
    $('sf-pick').className = 'sf-btn' + (loaded ? ' sf-secondary' : '');
    $('sf-drop-hint').textContent = t('dropHint');
    $('sf-pick').disabled = state.busy;
    $('sf-drop').classList.toggle('sf-over', state.dragOver);
    $('sf-drop').classList.toggle('sf-compact', loaded);
  }
  function renderFile() {
    var box = $('sf-file-info');
    if (!state.file) { box.innerHTML = ''; return; }
    var html = '<span>' + esc(t('fileInfo', { file: state.file.name })) + '</span>';
    if (state.sheets.length === 1) {
      html += '<span>' + esc(t('sheetInfo', { sheet: state.sheets[0].name })) + '</span>';
    } else if (state.sheets.length > 1) {
      html += '<label class="sf-sheet"><span>' + esc(t('sheetChoose')) + '</span> <select id="sf-sheet">' +
        state.sheets.map(function (s) {
          return '<option value="' + s.index + '"' + (s.index === state.sheetIndex ? ' selected' : '') + '>' + esc(s.name) + '</option>';
        }).join('') + '</select></label>';
    }
    box.innerHTML = html;
  }
  function renderCheck() {
    var box = $('sf-check');
    if (state.progress === 'reading') {
      box.innerHTML = '<div class="sf-result"><span class="sf-spinner" aria-hidden="true"></span>' + esc(t('reading')) + '</div>';
      return;
    }
    if (state.fileError) {
      box.innerHTML = '<div class="sf-result sf-bad">' + esc(t(state.fileError)) + ' <a href="#vorlage">' + esc(t('toTemplate')) + '</a></div>';
      return;
    }
    var c = state.check;
    if (!c) { box.innerHTML = ''; return; }
    var errors = c.issues.filter(function (i) { return i.level === 'error'; });
    var warnings = c.issues.filter(function (i) { return i.level === 'warning'; });
    var item = function (i) {
      var code = WEB_CODES[i.code];
      var text = core.formatMessage(code ? Object.assign({}, i, { code: code }) : i, state.lang);
      var cell = i.cell ? '<span class="sf-cell">' + esc(t('cell', { c: i.cell })) + '</span>' : '';
      return '<li class="' + i.level + '">' + cell + esc(text) + '</li>';
    };
    var html = '';
    if (errors.length) {
      html += '<div class="sf-result sf-bad">' + esc(t('checkErrors', { n: errors.length })) + '</div><ul class="sf-issues">' + errors.map(item).join('') + '</ul>';
    } else {
      html += '<div class="sf-result sf-ok">' + esc(t('checkOk')) + '</div>';
    }
    if (warnings.length) {
      html += '<div class="sf-result">' + esc(t('checkWarnings')) + '</div><ul class="sf-issues">' + warnings.map(item).join('') + '</ul>';
    }
    if (!errors.length && c.model && c.model.totals) {
      var tt = c.model.totals;
      var m = function (v) { return core.formatMoney(v, state.lang); };
      html += '<dl class="sf-totals"><dt>' + esc(t('lines')) + '</dt><dd>' + c.model.lines.length + '</dd>' +
        '<dt>' + esc(t('net')) + '</dt><dd>' + m(tt.taxExclusive) + ' EUR</dd>' +
        '<dt>' + esc(t('vat')) + '</dt><dd>' + m(tt.taxTotal) + ' EUR</dd>' +
        '<dt class="sf-grand">' + esc(t('gross')) + '</dt><dd class="sf-grand">' + m(tt.taxInclusive) + ' EUR</dd>' +
        (tt.prepaid ? '<dt class="sf-grand">' + esc(t('payable')) + '</dt><dd class="sf-grand">' + m(tt.payable) + ' EUR</dd>' : '') +
        '</dl><p class="sf-small">' + esc(t('formats', { p: c.model.leitwegId ? 'XRechnung' : 'EN 16931' })) + '</p>';
    }
    box.innerHTML = html;
  }
  function renderCreate() {
    var box = $('sf-create-box');
    var ready = !!(state.check && state.check.ok && !state.fileError);
    box.hidden = !ready;
    $('sf-pdflang-label').textContent = t('pdfLang');
    $('sf-create').textContent = t('createBtn');
    $('sf-create').disabled = state.busy;
    var res = $('sf-result');
    if (state.progress && state.progress !== 'reading') {
      res.innerHTML = '<div class="sf-result"><span class="sf-spinner" aria-hidden="true"></span>' + esc(t(state.progress)) + '</div>';
    } else if (state.files) {
      res.innerHTML = '<div class="sf-result sf-ok">' + esc(t('created')) + '</div><div class="sf-downloads">' +
        '<a class="sf-btn" id="sf-dl-xml" href="' + state.files.xmlUrl + '" download="' + esc(state.files.names.xml) + '">' + esc(t('dlXml')) + '</a>' +
        '<a class="sf-btn" id="sf-dl-pdf" href="' + state.files.pdfUrl + '" download="' + esc(state.files.names.pdf) + '">' + esc(t('dlPdf')) + '</a></div>';
    } else {
      res.innerHTML = '';
    }
  }
  function keyForm(a, buttonLabel) {
    var value = state.keyDraft !== null ? state.keyDraft : (a.licenceKey || '');
    return '<form class="sf-key" id="sf-lic-form"><label class="sf-visually-hidden" for="sf-lic-key">' + esc(t('keyLabel')) + '</label>' +
      '<input id="sf-lic-key" autocomplete="off" spellcheck="false" placeholder="' + esc(t('keyPlaceholder')) + '"' +
      (value ? ' value="' + esc(value) + '"' : '') + '><button type="submit" class="sf-btn sf-secondary"' + (state.busy ? ' disabled' : '') + '>' +
      esc(buttonLabel) + '</button></form>';
  }
  function renderLicence() {
    $('sf-lic-title').textContent = t('licTitle');
    var box = $('sf-licence');
    var a = state.account || { freeUsed: 0 };
    var access = core.decideAccess(a, nowIso());
    var html = '';
    var links = function (withRemove) {
      var remove = !withRemove ? '' : (state.confirmRemove
        ? '<span class="sf-confirm">' + esc(t('licRemoveConfirm')) + ' <a href="#" id="sf-lic-remove-yes">' + esc(t('licRemoveYes')) + '</a> ' +
          '<a href="#" id="sf-lic-remove-no">' + esc(t('licRemoveNo')) + '</a></span>'
        : '<a href="#" id="sf-lic-remove">' + esc(t('licRemove')) + '</a>');
      return '<p class="sf-links"><a href="' + esc(CFG.polar.customerPortal) + '" target="_blank" rel="noopener">' + esc(t('licManage')) + '</a> ' + remove + '</p>';
    };
    if (access.licence === 'active') {
      html += '<span class="sf-pill sf-active">' + esc(t('licActive')) + '</span>' +
        '<p class="sf-small">' + esc(t('licKey', { k: core.maskKey(a.licenceKey), d: core.formatDate(localYmd(new Date(a.licenceLastOkAt)), state.lang) })) + '</p>';
      if (a.licenceLastError) { html += '<p class="sf-small">' + esc(t('licGrace')) + '</p>'; }
      html += links(true);
    } else if (a.licenceKey && state.licenceChecking) {
      html += '<span class="sf-pill sf-neutral">' + esc(t('licChecking')) + '</span>' +
        '<p class="sf-small">' + esc(t('licStored', { k: core.maskKey(a.licenceKey) })) + '</p>' + links(false);
    } else if (access.licence === 'unverified') {
      html += '<span class="sf-pill sf-neutral">' + esc(t('licUnverifiedPill')) + '</span>' +
        '<p class="sf-warn">' + esc(t('licUnverified')) + '</p>' + keyForm(a, t('licRetry')) + links(true);
    } else {
      html += access.freeLeft > 0
        ? '<span class="sf-pill sf-free">' + esc(t('freeLeft', { n: access.freeLeft, m: CFG.freeInvoices })) + '</span>'
        : '<span class="sf-pill sf-blocked">' + esc(t('freeUsedUp', { m: CFG.freeInvoices })) + '</span>';
      html += '<p class="sf-small">' + esc(t('freeWhere')) + '</p>';
      if (a.licenceKey) { html += '<p class="sf-warn">' + esc(t('licInvalid')) + '</p>'; }
      html += '<p class="sf-small">' + esc(t('price')) + '</p>' +
        '<p class="sf-buy"><a class="sf-btn" href="' + esc(CFG.polar.checkoutMonthly) + '" target="_blank" rel="noopener">' + esc(t('buyMonthly')) + '</a>' +
        '<a class="sf-btn" href="' + esc(CFG.polar.checkoutYearly) + '" target="_blank" rel="noopener">' + esc(t('buyYearly')) + '</a></p>' +
        keyForm(a, t('activate'));
      if (a.licenceKey) { html += links(true); }
    }
    if (state.licenceMsg) {
      html += '<p class="' + (state.licenceMsg.kind === 'error' ? 'sf-warn' : 'sf-small') + '">' + esc(t(state.licenceMsg.key)) + '</p>';
    }
    if (!storageOk) { html += '<p class="sf-small">' + esc(t('storageOff')) + '</p>'; }
    box.innerHTML = html;
  }
  function renderAll() {
    renderPage();
    renderNotice();
    renderDrop();
    renderFile();
    renderCheck();
    renderCreate();
    renderLicence();
  }
  function setBusy(busy) {
    state.busy = busy;
    renderDrop();
    renderCreate();
    Array.prototype.forEach.call(document.querySelectorAll('#sf-licence button'), function (b) { b.disabled = busy; });
  }
  function setProgress(key) {
    state.progress = key;
    renderCheck();
    renderCreate();
  }

  // ---------------------------------------------------------------------------
  // Actions
  // ---------------------------------------------------------------------------
  function revokeFiles() {
    if (state.files) {
      try { URL.revokeObjectURL(state.files.xmlUrl); URL.revokeObjectURL(state.files.pdfUrl); } catch (e) { /* ignore */ }
    }
    state.files = null;
  }
  function readBytes(file) {
    if (file.arrayBuffer) { return file.arrayBuffer().then(function (b) { return new Uint8Array(b); }); }
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(new Uint8Array(fr.result)); };
      fr.onerror = function () { reject(fr.error || new Error('read failed')); };
      fr.readAsArrayBuffer(file);
    });
  }
  function runCheck() {
    var values = X.sheetValues(state.wb, state.sheetIndex, core);
    state.check = core.checkSheet(values, { today: localYmd(new Date()) });
  }
  function onFile(file) {
    if (!file || state.busy) { return; }
    setNotice(null);
    revokeFiles();
    state.file = { name: file.name, size: file.size };
    state.wb = null;
    state.sheets = [];
    state.sheetIndex = null;
    state.check = null;
    state.fileError = null;
    if (file.size > X.LIMITS.fileBytes) {
      state.fileError = 'errTooLarge';
      renderAll();
      return;
    }
    setBusy(true);
    setProgress('reading');
    renderFile();
    readBytes(file).then(function (bytes) {
      if (!window.fflate) { throw { key: 'errLibs' }; }
      var wb = X.readWorkbook(bytes, window.fflate);
      var sheets = X.findInvoiceSheets(wb, core);
      if (!sheets.length) { throw { code: 'NO_TEMPLATE' }; }
      state.wb = wb;
      state.sheets = sheets;
      state.sheetIndex = X.defaultInvoiceSheet(wb, sheets);
      runCheck();
      // The components for Create load in the background meanwhile.
      if (state.check.ok) { loadLibs().catch(logError); }
    }).catch(function (err) {
      if (err && err.key) {
        state.fileError = err.key;
      } else if (err && FILE_ERRORS[err.code]) {
        state.fileError = FILE_ERRORS[err.code];
      } else {
        logError(err);
        state.fileError = 'errCorrupt';
      }
      state.check = null;
    }).then(function () {
      state.busy = false;
      state.progress = null;
      renderAll();
      var target = $('sf-check');
      if (target && target.scrollIntoView) { target.scrollIntoView({ block: 'nearest' }); }
    });
  }
  function onSheet(index) {
    if (state.busy || !state.wb) { return; }
    revokeFiles();
    state.sheetIndex = index;
    try { runCheck(); } catch (err) {
      state.check = null;
      state.fileError = err && FILE_ERRORS[err.code] ? FILE_ERRORS[err.code] : 'errCorrupt';
    }
    renderAll();
  }
  function pdfLang() {
    var el = document.querySelector('input[name="sf-pdflang"]:checked');
    return el ? el.value : 'de';
  }
  function onCreate() {
    if (state.busy) { return; }
    setNotice(null);
    if (!state.check) { setNotice('error', 'noFile'); return; }
    if (!state.check.ok) { setNotice('error', 'fixFirst'); return; }
    setBusy(true);
    revokeFiles();
    setProgress('creating');
    var access;
    refreshLicenceIfNeeded().then(function () {
      state.account = readState();
      access = core.decideAccess(state.account, nowIso());
      if (!access.allowed) {
        renderLicence();
        // A paying customer whose key could not be checked is not asked to buy (LIC-1, LIC-4).
        var why = { unverified: 'licUnverified', invalid: 'licInvalid', expired: 'licInvalid' }[access.licence] || 'paywall';
        throw { key: why };
      }
      setProgress('loadingLibs');
      return loadLibs().catch(function (err) { logError(err); throw { key: 'errLibs' }; });
    }).then(function (libs) {
      setProgress('creating');
      return libs.generate.generateFiles(state.check.model, libs, fonts(), { lang: pdfLang(), now: new Date(), appName: CFG.appName });
    }).then(function (out) {
      state.files = {
        names: out.names,
        xmlUrl: URL.createObjectURL(new Blob([out.xml], { type: 'application/xml' })),
        pdfUrl: URL.createObjectURL(new Blob([out.pdf], { type: 'application/pdf' }))
      };
      // Counted once both files exist; a failing storage never turns the result into an error (LIC-7).
      try { state.account = recordGeneration(access.mode); } catch (e) { logError(e); }
      state.busy = false;
      state.progress = null;
      renderAll();
    }).catch(function (err) {
      state.busy = false;
      state.progress = null;
      renderAll();
      if (err && err.key) { setNotice('error', err.key); } else { logError(err); setNotice('error', 'errGeneric'); }
    });
  }
  function onActivate(rawKey) {
    if (state.busy) { return; }
    var key = core.normalizeLicenceKey(rawKey);
    var msg = function (kind, k) { state.licenceMsg = { kind: kind, key: k }; renderLicence(); };
    if (!key) { msg('error', 'licKeyMissing'); return; }
    if (key.length > core.LICENCE_KEY_MAX) { msg('error', 'licInvalid'); return; }
    if (!licenceConfigured()) { msg('error', 'licNotSetUp'); return; }
    state.keyDraft = key;
    setBusy(true);
    msg('info', 'checkingKey');
    polarValidate(key).then(function (result) {
      state.account = saveLicenceCheck(key, result);
      state.licenceMsg = null;
      state.keyDraft = null;
      if (result.status === 'valid') { setNotice(null); }
    }).then(function () { setBusy(false); renderLicence(); }, function (err) {
      logError(err);
      setBusy(false);
      msg('error', 'errGeneric');
    });
  }
  function onLang(lang) {
    state.lang = lang === 'en' ? 'en' : 'de';
    if (!storageSet(LANG_KEY, state.lang)) { memory.lang = state.lang; }
    renderAll();
  }

  // ---------------------------------------------------------------------------
  // Wiring
  // ---------------------------------------------------------------------------
  var input = $('sf-file');
  input.addEventListener('change', function () {
    var f = input.files && input.files[0];
    // Reset, so picking the same file again after a change in Excel reads it again.
    input.value = '';
    onFile(f);
  });
  $('sf-pick').addEventListener('click', function () { if (!state.busy) { input.click(); } });
  var drop = $('sf-drop');
  ['dragenter', 'dragover'].forEach(function (type) {
    drop.addEventListener(type, function (ev) {
      ev.preventDefault();
      if (ev.dataTransfer) { ev.dataTransfer.dropEffect = 'copy'; }
      if (!state.dragOver) { state.dragOver = true; renderDrop(); }
    });
  });
  ['dragleave', 'dragend'].forEach(function (type) {
    drop.addEventListener(type, function () { state.dragOver = false; renderDrop(); });
  });
  drop.addEventListener('drop', function (ev) {
    ev.preventDefault();
    state.dragOver = false;
    renderDrop();
    var f = ev.dataTransfer && ev.dataTransfer.files && ev.dataTransfer.files[0];
    onFile(f);
  });
  // A file dropped next to the drop zone must not make the browser open it and leave the page.
  window.addEventListener('dragover', function (ev) { ev.preventDefault(); });
  window.addEventListener('drop', function (ev) { ev.preventDefault(); });
  $('sf-create').addEventListener('click', onCreate);
  document.addEventListener('change', function (ev) {
    if (ev.target && ev.target.id === 'sf-sheet') { onSheet(parseInt(ev.target.value, 10)); }
  });
  document.addEventListener('click', function (ev) {
    var el = ev.target;
    if (!el || !el.closest) { return; }
    var langBtn = el.closest('.sf-lang button');
    if (langBtn) { onLang(langBtn.getAttribute('data-lang')); return; }
    if (el.id === 'sf-lic-remove' || el.id === 'sf-lic-remove-no') {
      ev.preventDefault();
      state.confirmRemove = el.id === 'sf-lic-remove';
      renderLicence();
      return;
    }
    if (el.id === 'sf-lic-remove-yes') {
      ev.preventDefault();
      if (state.busy) { return; }
      state.account = removeLicence();
      state.licenceMsg = null;
      state.keyDraft = null;
      state.confirmRemove = false;
      renderLicence();
    }
  });
  document.addEventListener('input', function (ev) {
    if (ev.target && ev.target.id === 'sf-lic-key') { state.keyDraft = ev.target.value; }
  });
  document.addEventListener('submit', function (ev) {
    if (ev.target && ev.target.id === 'sf-lic-form') { ev.preventDefault(); onActivate($('sf-lic-key').value); }
  });
  // Another tab of this page changed the counter or the key.
  window.addEventListener('storage', function (ev) {
    if (ev.key === STATE_KEY) { state.account = readState(); renderLicence(); }
  });

  document.documentElement.classList.add('sf-js');
  renderAll();
  // A stored key is checked when the page opens, if a check is due (LIC-2); meanwhile the card says so.
  state.licenceChecking = licenceCheckDue();
  if (state.licenceChecking) {
    renderLicence();
    refreshLicenceIfNeeded().catch(logError).then(function () { state.licenceChecking = false; renderLicence(); });
  }

  window.SheetFakturaWeb = { state: state, loadLibs: loadLibs, texts: STR };
}());
