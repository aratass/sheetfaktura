/* SheetFaktura funnel counter, browser side (same pattern as the DAS step counter).
   On the Excel tool page it sends one small anonymous event at these moments:
     open   the page was opened
     check  a file was checked: tier ok (no errors), err (errors to fix) or bad (not a readable template)
     make   an e-invoice was made: tier free or lic (licence active); each new pair of files counts once
     limit  the free e-invoices are used up (shown in the licence box)
     buy    a Polar buy link was clicked: tier 12m, 7d, 1m or other
   Sent: event, tier, a source label and a campaign label. Source: utm_source of the address; else a short name of
   the referring site (google, bing, ...), "site" for our own pages, "direct" without a referrer. Campaign:
   utm_campaign; else the referring page of our own site; else "none". Nothing else: no identifier, no file content.
   Nothing is stored in the browser: no cookies, no localStorage. Apart from make, each event and tier counts once
   per page view. The endpoint comes from <meta name="sf-count" content="https://.../e">; when it is empty or
   missing, nothing is sent. Counting never breaks the page: every error is swallowed. */
(function () {
  'use strict';
  var meta = document.querySelector('meta[name="sf-count"]');
  var endpoint = meta ? (meta.getAttribute('content') || '').trim() : '';
  if (!/^https:\/\/[^\s]+$/.test(endpoint) && !/^http:\/\/(127\.0\.0\.1|localhost)(:\d+)?\/[^\s]*$/.test(endpoint)) { return; }

  var LABEL_MAX = 40;
  var PLANS = { TK6HjmcBGljvxrKZ6xyxjCP2NKW0JlsbpFF3X00eUgI: '12m', '87xpVz5RjJ5mCE9tWnmil9T5j088uVucUVFPA3TlUt4': '7d', YuvQl4FoqRVCzLeQuGTKuLV5mJTPJplnoYMmq41KzAB: '1m' };
  var SITES = [
    [/^workspace\.google\.com$/, 'marketplace'],
    [/(^|\.)google\.[a-z.]+$/, 'google'],
    [/googlequicksearchbox/, 'google'],
    [/(^|\.)bing\.com$/, 'bing'],
    [/(^|\.)duckduckgo\.com$/, 'duckduckgo'],
    [/(^|\.)ecosia\.org$/, 'ecosia'],
    [/(^|\.)yahoo\.[a-z.]+$/, 'yahoo'],
    [/(^|\.)qwant\.com$/, 'qwant'],
    [/(^|\.)startpage\.com$/, 'startpage'],
    [/(^|\.)ihk-muenchen\.de$/, 'ihk'],
    [/(^|\.)omr\.com$/, 'omr'],
    [/(^|\.)(chatgpt\.com|openai\.com)$/, 'chatgpt'],
    [/(^|\.)perplexity\.ai$/, 'perplexity'],
    [/(^|\.)facebook\.com$/, 'facebook'],
    [/(^|\.)linkedin\.com$/, 'linkedin']
  ];

  function clean(value) {
    if (typeof value !== 'string') { return 'none'; }
    var v = value;
    try { v = v.normalize('NFKD').replace(/[̀-ͯ]/g, ''); } catch (e) { /* old browser */ }
    v = v.toLowerCase().replace(/[^a-z0-9._-]+/g, '_').replace(/^[._-]+|[._-]+$/g, '').slice(0, LABEL_MAX).replace(/[._-]+$/, '');
    return v || 'none';
  }

  function labels() {
    var params;
    try { params = new URLSearchParams(location.search); } catch (e) { params = null; }
    var utmSource = params && params.get('utm_source');
    var utmCampaign = params && params.get('utm_campaign');
    var source = 'direct';
    var campaign = 'none';
    var ref = document.referrer || '';
    if (ref) {
      try {
        var u = new URL(ref);
        var host = (u.hostname || u.host || '').toLowerCase().replace(/^www\./, '');
        if (host && host === location.hostname.toLowerCase().replace(/^www\./, '')) {
          source = 'site';
          var page = u.pathname.split('/').pop().replace(/\.html$/, '');
          campaign = page || 'index';
        } else {
          source = 'other';
          for (var i = 0; i < SITES.length; i++) {
            if (SITES[i][0].test(host || ref)) { source = SITES[i][1]; break; }
          }
          if (source === 'other' && host) { source = host; }
        }
      } catch (e) {
        source = 'other';
      }
    }
    if (utmSource) { source = utmSource; }
    if (utmCampaign) { campaign = utmCampaign; }
    return { s: clean(source), c: clean(campaign) };
  }

  var lab = labels();
  var sent = {};

  function send(event, tier, key) {
    try {
      var k = key || (event + ':' + (tier || ''));
      if (sent[k]) { return; }
      sent[k] = true;
      var body = JSON.stringify({ e: event, t: tier || '', s: lab.s, c: lab.c });
      if (navigator.sendBeacon && navigator.sendBeacon(endpoint, body)) { return; }
      if (window.fetch) {
        window.fetch(endpoint, { method: 'POST', body: body, keepalive: true, mode: 'no-cors', credentials: 'omit',
          headers: { 'Content-Type': 'text/plain;charset=UTF-8' } }).catch(function () {});
      }
    } catch (e) {
      // counting must never break the page
    }
  }

  function byId(id) { return document.getElementById(id); }

  function onCheck() {
    var box = byId('sf-check');
    if (!box) { return; }
    if (box.querySelector('.sf-ok')) { send('check', 'ok'); }
    else if (box.querySelector('ul.sf-issues li.error')) { send('check', 'err'); }
    else if (box.querySelector('.sf-bad')) { send('check', 'bad'); }
  }

  function onResult() {
    var box = byId('sf-result');
    var xml = box && box.querySelector('#sf-dl-xml');
    if (!xml) { return; }
    var href = xml.getAttribute('href') || '';
    window.setTimeout(function () {
      var lic = byId('sf-licence');
      var tier = lic && lic.querySelector('.sf-pill.sf-active') ? 'lic' : 'free';
      send('make', tier, 'make:' + href);
    }, 0);
  }

  function onLicence() {
    var lic = byId('sf-licence');
    if (lic && lic.querySelector('.sf-pill.sf-blocked')) { send('limit', ''); }
  }

  function watch(id, fn) {
    var el = byId(id);
    if (!el || !window.MutationObserver) { return; }
    new MutationObserver(fn).observe(el, { childList: true, subtree: true });
    fn();
  }

  function onClick(ev) {
    try {
      if (ev.type === 'auxclick' && ev.button !== 1) { return; }
      var t = ev.target;
      var a = t && t.closest ? t.closest('a[href*="buy.polar.sh/"]') : null;
      if (!a) { return; }
      var m = /polar_cl_([A-Za-z0-9]+)/.exec(a.getAttribute('href') || '');
      send('buy', (m && PLANS[m[1]]) || 'other');
    } catch (e) {
      // ignore
    }
  }

  function start() {
    send('open', '');
    watch('sf-check', onCheck);
    watch('sf-result', onResult);
    watch('sf-licence', onLicence);
    document.addEventListener('click', onClick, true);
    document.addEventListener('auxclick', onClick, true);
  }

  if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', start); } else { start(); }
})();
