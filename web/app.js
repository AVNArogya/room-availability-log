(function () {
  'use strict';

  var CFG = window.APP_CONFIG || {};
  var DEMO = !CFG.API_URL;
  var LS = { queue: 'ral.queue', cache: 'ral.entries', lists: 'ral.lists', rates: 'ral.rates', demo: 'ral.demo', user: 'ral.enteredBy' };
  var LOST = ['Lost', 'Went elsewhere', 'Cancelled'];
  var BOOKED = ['Booked', 'Admitted'];

  var DEFAULT_LISTS = {
    newOrReturning: ['New', 'Returning'],
    source: ['Incoming call - Google', 'Incoming call - Website', 'Direct counseling', 'Zoho Bookings',
             'Website', 'Google Ad words', 'Referral', 'Walk-in'],
    doctor: ['Dr. Geethu', 'Dr. Sridevi', 'Dr. Chippi', 'Dr. Biju', 'Dr. Ramesh', 'Dr. Padmavathy', 'Dr. Jyothish'],
    roomCategory: ['Economy', 'Executive AC', 'Classic Villa'],
    status: ['Waiting', 'Booked', 'Admitted', 'Lost'],
    reasonNotAdmitted: ['Room not available', 'Waiting time too long', 'Cost', 'Postponed by patient',
                        'Went elsewhere', 'Not reachable', 'Other'],
    enteredBy: []
  };

  var $ = function (s, root) { return (root || document).querySelector(s); };
  var $$ = function (s, root) { return Array.prototype.slice.call((root || document).querySelectorAll(s)); };
  var form = $('#entryForm');
  var state = { entries: load(LS.cache, []), lists: mergeLists(load(LS.lists, {})), rates: load(LS.rates, {}), page: 1 };
  var PAGE_SIZE = 10;

  // Sheet lists override the defaults one list at a time; a missing or empty column keeps the default.
  function mergeLists(fromSheet) {
    var out = Object.assign({}, DEFAULT_LISTS);
    Object.keys(fromSheet || {}).forEach(function (k) {
      if (Array.isArray(fromSheet[k]) && fromSheet[k].length) out[k] = fromSheet[k];
    });
    out.status = DEFAULT_LISTS.status;
    if (!(fromSheet && fromSheet.offeredRoom && fromSheet.offeredRoom.length)) {
      out.offeredRoom = out.roomCategory.concat(['Not offered']);
    }
    return out;
  }

  /* ---------- storage helpers ---------- */
  function load(k, fallback) {
    try { var v = JSON.parse(localStorage.getItem(k)); return v == null ? fallback : v; } catch (e) { return fallback; }
  }
  function save(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* storage unavailable */ } }

  /* ---------- API ---------- */
  function apiGet(action, extra) {
    if (DEMO) return Promise.resolve(demoGet(action));
    var url = CFG.API_URL + '?action=' + action + (extra || '') + '&key=' + encodeURIComponent(CFG.APP_KEY || '');
    return fetch(url).then(unwrap);
  }
  function apiPost(body) {
    if (DEMO) return Promise.resolve(demoPost(body));
    body.key = CFG.APP_KEY || '';
    // text/plain keeps this a "simple" request so Apps Script doesn't need CORS preflight.
    return fetch(CFG.API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(body) })
      .then(unwrap);
  }
  function unwrap(res) {
    return res.json().then(function (j) {
      if (!j.ok) { var err = new Error(j.error || 'Server error'); err.fatal = true; throw err; }
      return j.data;
    });
  }

  /* ---------- demo backend (localStorage) ---------- */
  function demoGet(action) {
    if (action === 'config') return { lists: DEFAULT_LISTS, rates: load(LS.rates, {}) };
    return { entries: load(LS.demo, []).slice().reverse() };
  }
  function demoPost(body) {
    var rows = load(LS.demo, []);
    var e = Object.assign({}, body.entry);
    var now = new Date().toISOString();
    if (body.action === 'add') {
      var dup = rows.filter(function (r) { return r.id === e.clientId; })[0];
      if (dup) return { entry: dup };
      e.id = e.clientId; e.createdAt = now;
    }
    e.updatedAt = now;
    e = computeDerived(e);
    if (body.action === 'add') rows.push(e);
    else rows = rows.map(function (r) { return r.id === e.id ? Object.assign(r, e) : r; });
    save(LS.demo, rows);
    return { entry: e };
  }

  /* ---------- derived fields (mirrors backend) ---------- */
  function daysBetween(a, b) {
    if (!a || !b) return '';
    return Math.round((Date.parse(b) - Date.parse(a)) / 86400000);
  }
  // Mirrors revenueLoss_ in backend/Code.gs; the sheet's value is the one that counts.
  function revenueLoss(e) {
    var rates = state.rates || {};
    var rateOf = function (c) { return rates[String(c || '').trim().toLowerCase()]; };
    var stay = parseFloat(e.plannedStay) || 0;
    var asked = rateOf(e.categoryAsked) || 0;
    if (LOST.indexOf(e.status) !== -1) return { days: stay, amount: stay * asked };
    var days = 0, amount = 0;
    var delay = daysBetween(e.dateWanted, e.admissionDate);
    if (delay > 0) { days = delay; amount += delay * asked; }
    if (BOOKED.indexOf(e.status) !== -1) {
      var given = rateOf(e.categoryAdmitted);
      if (given === undefined) given = rateOf(e.offeredRoom);
      if (given !== undefined && given < asked) amount += (asked - given) * stay;
    }
    return { days: days, amount: amount };
  }
  function computeDerived(e) {
    e.waitDays = daysBetween(e.dateWanted, e.earliestDate);
    var loss = revenueLoss(e);
    e.patientDaysLost = loss.days;
    e.revenueLost = loss.amount;
    e.month = e.enquiryDate ? new Date(e.enquiryDate + 'T00:00:00')
      .toLocaleDateString('en-GB', { month: 'short', year: 'numeric' }) : '';
    return e;
  }

  /* ---------- offline queue ---------- */
  function queue() { return load(LS.queue, []); }
  function enqueue(item) { var q = queue(); q.push(item); save(LS.queue, q); renderPending(); }

  var flushing = false;
  function flush() {
    if (flushing || !navigator.onLine) return Promise.resolve();
    var q = queue();
    if (!q.length) return Promise.resolve();
    flushing = true;
    var item = q[0];
    return apiPost(item).then(function (data) {
      save(LS.queue, queue().slice(1));
      upsertLocal(data.entry);
      flushing = false;
      return flush();
    }, function (err) {
      flushing = false;
      if (err.fatal) { // server rejected it: drop it so it doesn't block the queue
        save(LS.queue, queue().slice(1));
        toast('Could not save ' + (item.entry.patientName || 'entry') + ': ' + err.message);
        return flush();
      }
    }).then(renderPending);
  }

  /* ---------- lists / dropdowns ---------- */
  function applyLists() {
    var L = state.lists;
    $$('select[data-list]').forEach(function (sel) {
      var cur = sel.value;
      var opts = (L[sel.dataset.list] || []).slice();
      var html = '<option value="">' + (sel.required ? 'Select…' : '—') + '</option>';
      html += opts.map(function (o) { return '<option>' + esc(o) + '</option>'; }).join('');
      sel.innerHTML = html;
      if (cur) setSelect(sel, cur);
    });
    ['source', 'doctor', 'reasonNotAdmitted', 'enteredBy'].forEach(function (k) {
      $('#dl-' + k).innerHTML = (L[k] || []).map(function (o) { return '<option value="' + esc(o) + '">'; }).join('');
    });
    $('#statusFilter').innerHTML = '<option value="">All statuses</option>' +
      (L.status || []).map(function (o) { return '<option>' + esc(o) + '</option>'; }).join('');
  }
  // Keep values that aren't in the list (e.g. older free-text data) selectable.
  function setSelect(sel, value) {
    if (value && !$$('option', sel).some(function (o) { return o.value === value; })) {
      var o = document.createElement('option'); o.textContent = value; sel.appendChild(o);
    }
    sel.value = value || '';
  }

  /* ---------- form ---------- */
  function today() {
    var d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
    return d.toISOString().slice(0, 10);
  }
  function resetForm() {
    form.reset();
    form.elements.id.value = '';
    form.elements.enquiryDate.value = today();
    form.elements.status.value = 'Waiting';
    form.elements.enteredBy.value = load(LS.user, '');
    $('#formTitle').textContent = 'New entry';
    $('#saveBtn').textContent = 'Save entry';
    $('#cancelEdit').hidden = true;
    $$('.invalid', form).forEach(function (el) { el.classList.remove('invalid'); });
    updateCalc();
  }
  function readForm() {
    var e = {};
    $$('input[name], select[name], textarea[name]', form).forEach(function (el) { e[el.name] = el.value.trim(); });
    var digits = e.phone.replace(/\D/g, '');
    e.phone = digits ? $('#countryCode').value + ' ' + digits : '';
    return e;
  }
  function fillForm(e) {
    resetForm();
    $$('input[name], select[name], textarea[name]', form).forEach(function (el) {
      var v = e[el.name] == null ? '' : String(e[el.name]);
      if (el.type === 'date') v = v.slice(0, 10);
      if (el.tagName === 'SELECT') setSelect(el, v); else el.value = v;
    });
    // Stored as "+91 9876543210"; older rows may have no code, which means India.
    var m = String(e.phone || '').match(/^(\+\d{1,4})\s*(.*)$/);
    var code = m && $$('#countryCode option').some(function (o) { return o.value === m[1]; }) ? m[1] : '+91';
    $('#countryCode').value = code;
    form.elements.phone.value = (m && code === m[1] ? m[2] : String(e.phone || '')).replace(/\D/g, '');
    $('#formTitle').textContent = 'Edit: ' + (e.patientName || 'entry');
    $('#saveBtn').textContent = 'Update entry';
    $('#cancelEdit').hidden = false;
    updateCalc();
  }
  function updateCalc() {
    var e = readForm();
    applyStatusFields(e.status);
    // Hidden fields are cleared on save, so leave them out of the estimate too.
    $$('[data-show][hidden] [name]', form).forEach(function (el) { e[el.name] = ''; });
    var w = daysBetween(e.dateWanted, e.earliestDate);
    $('#waitOut').textContent = w === '' ? '—' : w + (w === 1 ? ' day' : ' days');
    var loss = revenueLoss(e);
    $('#daysLostOut').textContent = loss.days;
    $('#revenueOut').textContent = '₹' + loss.amount.toLocaleString('en-IN');
    // UHID is only mandatory for returning patients; new patients aren't registered yet.
    var returning = /^return/i.test(e.newOrReturning);
    form.elements.uhid.required = returning;
    $('#uhidReq').hidden = !returning;
    $('#uhidHint').hidden = returning;
  }
  // Show only the Outcome fields that belong to the chosen status (see data-show in index.html).
  function applyStatusFields(status) {
    var st = String(status || '').toLowerCase();
    $$('[data-show]', form).forEach(function (label) {
      var on = label.dataset.show.split(' ').indexOf(st) !== -1;
      label.hidden = !on;
      var field = $('input, select, textarea', label);
      field.disabled = !on;
      if (label.hasAttribute('data-require')) field.required = on;
    });
  }
  function validate() {
    var bad = $$('[required]:not(:disabled)', form).filter(function (el) { return !el.value.trim(); });
    $$('.invalid', form).forEach(function (el) { el.classList.remove('invalid'); });
    bad.forEach(function (el) { el.classList.add('invalid'); });
    var phoneEl = form.elements.phone;
    var phone = phoneEl.value.replace(/[\s-]/g, '');
    var india = $('#countryCode').value === '+91';
    var phoneMsg = '';
    if (phone && india && !/^[6-9]\d{9}$/.test(phone)) phoneMsg = 'Indian mobile number must be 10 digits';
    else if (phone && !india && !/^\d{5,14}$/.test(phone)) phoneMsg = 'Phone number should be 5 to 14 digits';
    if (phoneMsg) { phoneEl.classList.add('invalid'); if (bad.indexOf(phoneEl) === -1) bad.push(phoneEl); }
    var stay = form.elements.plannedStay;
    if (stay.value && !(Number(stay.value) >= 1)) { stay.classList.add('invalid'); bad.push(stay); }
    if (bad.length) { bad[0].focus(); toast(phoneMsg && bad[0] === phoneEl ? phoneMsg : 'Please fill the highlighted fields'); }
    return !bad.length;
  }

  form.addEventListener('submit', function (ev) {
    ev.preventDefault();
    if (!validate()) return;
    var e = readForm();
    $$('[data-show][hidden] [name]', form).forEach(function (el) { e[el.name] = ''; });
    save(LS.user, e.enteredBy);
    var isEdit = !!e.id;
    if (!isEdit) { e.clientId = uid(); delete e.id; }
    var item = { action: isEdit ? 'update' : 'add', entry: e };

    // Optimistic local copy so the list shows it immediately, even offline.
    var local = computeDerived(Object.assign({}, e, { id: e.id || e.clientId, _queued: true }));
    upsertLocal(local);
    enqueue(item);

    var btn = $('#saveBtn'); btn.disabled = true;
    flush().then(function () {
      btn.disabled = false;
      var key = e.clientId || e.id;
      var stillQueued = queue().some(function (q) { return (q.entry.clientId || q.entry.id) === key; });
      toast(stillQueued ? 'Saved offline — will sync when online' : (isEdit ? 'Entry updated' : DEMO ? 'Entry saved (demo)' : 'Entry saved to Google Sheet'));
      resetForm();
      if (isEdit) show('list');
    });
  });
  $('#resetBtn').addEventListener('click', resetForm);
  form.addEventListener('input', updateCalc);
  form.addEventListener('change', updateCalc);
  $('#cancelEdit').addEventListener('click', function () { resetForm(); show('list'); });

  /* ---------- list ---------- */
  function upsertLocal(entry) {
    var found = false;
    state.entries = state.entries.map(function (x) {
      if (x.id === entry.id) { found = true; return entry; }
      return x;
    });
    if (!found) state.entries.unshift(entry);
    save(LS.cache, state.entries);
    renderList();
  }
  function refresh() {
    setSync();
    return Promise.all([apiGet('config'), apiGet('list', '&limit=5000')]).then(function (r) {
      if (r[0] && r[0].lists) { save(LS.lists, r[0].lists); state.lists = mergeLists(r[0].lists); applyLists(); }
      if (r[0] && r[0].rates) { state.rates = r[0].rates; save(LS.rates, state.rates); updateCalc(); }
      var queuedIds = queue().map(function (q) { return q.entry.id || q.entry.clientId; });
      var queued = state.entries.filter(function (x) { return queuedIds.indexOf(x.id) !== -1; });
      state.entries = queued.concat(r[1].entries.filter(function (x) { return queuedIds.indexOf(x.id) === -1; }));
      save(LS.cache, state.entries);
      renderList();
    }).catch(function (err) {
      toast(err.fatal ? err.message : 'Offline — showing saved entries');
    });
  }
  function renderList() {
    var q = $('#search').value.trim().toLowerCase();
    var st = $('#statusFilter').value;
    var rows = state.entries.filter(function (e) {
      if (st && e.status !== st) return false;
      if (!q) return true;
      return [e.patientName, e.phone, e.uhid, e.city, e.condition, e.doctor, e.categoryAsked]
        .join(' ').toLowerCase().indexOf(q) !== -1;
    });
    var num = function (v) { return v === '' || v == null ? '—' : esc(v); };
    var pages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
    state.page = Math.min(Math.max(1, state.page), pages);
    var first = (state.page - 1) * PAGE_SIZE;
    $('#entries').innerHTML = rows.slice(first, first + PAGE_SIZE).map(function (e) {
      var cls = LOST.indexOf(e.status) !== -1 ? 'lost' : (e.status === 'Waiting' ? 'waiting' : '');
      var revenue = Number(e.revenueLost) ? '₹' + Number(e.revenueLost).toLocaleString('en-IN') : '—';
      return '<tr class="entry" data-id="' + esc(e.id) + '" tabindex="0">' +
        '<td class="name">' + esc(e.patientName) + (e.uhid ? ' <span class="muted">' + esc(e.uhid) + '</span>' : '') + '</td>' +
        '<td>' + fmtDate(e.enquiryDate) + '</td>' +
        '<td>' + esc(e.phone || '—') + '</td>' +
        '<td>' + esc(e.categoryAsked || '—') + '</td>' +
        '<td>' + esc(e.offeredRoom || '—') + '</td>' +
        '<td>' + fmtDate(e.dateWanted) + '</td>' +
        '<td class="num">' + num(e.waitDays) + '</td>' +
        '<td>' + (e._queued ? '<span class="badge queued">Not synced</span>' : '<span class="badge ' + cls + '">' + esc(e.status || '—') + '</span>') + '</td>' +
        '<td class="num">' + num(e.patientDaysLost) + '</td>' +
        '<td class="num">' + revenue + '</td>' +
        '<td class="muted">' + esc(e.enteredBy || '—') + '</td>' +
        '</tr>';
    }).join('');
    $('#empty').hidden = rows.length > 0;
    renderPager(rows.length, pages);
    renderStats(rows);
  }
  // "Showing 11–20 of 57" plus Prev / page numbers / Next. Long ranges collapse to 1 … 4 5 6 … 12.
  function renderPager(total, pages) {
    var pager = $('#pager');
    pager.hidden = total === 0;
    if (!total) return;
    var p = state.page;
    var from = (p - 1) * PAGE_SIZE + 1, to = Math.min(p * PAGE_SIZE, total);
    var nums = [];
    for (var i = 1; i <= pages; i++) {
      if (i === 1 || i === pages || Math.abs(i - p) <= 1) nums.push(i);
      else if (nums[nums.length - 1] !== '…') nums.push('…');
    }
    var btn = function (label, page, opts) {
      opts = opts || {};
      return '<button type="button" class="pg' + (opts.current ? ' current' : '') + '" data-page="' + page + '"' +
        (opts.disabled ? ' disabled' : '') + (opts.current ? ' aria-current="page"' : '') +
        (opts.aria ? ' aria-label="' + opts.aria + '"' : '') + '>' + label + '</button>';
    };
    pager.innerHTML =
      '<span class="pg-info">Showing ' + from + '–' + to + ' of ' + total + '</span>' +
      '<span class="pg-buttons">' +
      btn('‹<span class="pg-word"> Prev</span>', p - 1, { disabled: p === 1, aria: 'Previous page' }) +
      nums.map(function (n) { return n === '…' ? '<span class="pg-gap">…</span>' : btn(n, n, { current: n === p, aria: 'Page ' + n }); }).join('') +
      btn('<span class="pg-word">Next </span>›', p + 1, { disabled: p === pages, aria: 'Next page' }) +
      '</span>';
  }
  function renderStats(rows) {
    var waiting = rows.filter(function (e) { return e.status === 'Waiting'; }).length;
    var lost = rows.filter(function (e) { return LOST.indexOf(e.status) !== -1; }).length;
    var waits = rows.map(function (e) { return Number(e.waitDays); }).filter(function (n) { return !isNaN(n) && n > 0; });
    var avg = waits.length ? (waits.reduce(function (a, b) { return a + b; }, 0) / waits.length).toFixed(1) : '—';
    var daysLost = rows.reduce(function (a, e) { return a + (Number(e.patientDaysLost) || 0); }, 0);
    $('#stats').innerHTML = [
      [rows.length, 'Entries'], [waiting, 'Waiting'], [lost, 'Lost'], [avg, 'Avg wait (days)'], [daysLost, 'Patient-days lost']
    ].map(function (s) { return '<div class="stat"><b>' + s[0] + '</b><span>' + s[1] + '</span></div>'; }).join('');
  }
  function renderPending() {
    var n = queue().length;
    var el = $('#pending');
    el.hidden = !n;
    el.textContent = n + ' entr' + (n === 1 ? 'y' : 'ies') + ' waiting to sync to Google Sheet.';
    setSync();
  }
  function openEntry(id) {
    var e = state.entries.filter(function (x) { return x.id === id; })[0];
    if (!e) return;
    fillForm(e);
    show('form');
    window.scrollTo(0, 0);
  }
  $('#entries').addEventListener('click', function (ev) {
    var li = ev.target.closest('.entry'); if (li) openEntry(li.dataset.id);
  });
  $('#entries').addEventListener('keydown', function (ev) {
    if (ev.key === 'Enter' && ev.target.classList.contains('entry')) openEntry(ev.target.dataset.id);
  });
  $('#search').addEventListener('input', function () { state.page = 1; renderList(); });
  $('#statusFilter').addEventListener('change', function () { state.page = 1; renderList(); });
  $('#pager').addEventListener('click', function (ev) {
    var b = ev.target.closest('button[data-page]');
    if (!b || b.disabled) return;
    state.page = Number(b.dataset.page);
    renderList();
    $('.table-wrap').scrollTop = 0;
  });
  $('#refreshBtn').addEventListener('click', function () { flush().then(refresh); });

  /* ---------- navigation & status ---------- */
  function show(view) {
    $$('.view').forEach(function (v) { v.classList.toggle('active', v.id === 'view-' + view); });
    $$('.tabs button').forEach(function (b) { b.classList.toggle('active', b.dataset.view === view); });
    if (view === 'list') renderList();
  }
  $$('.tabs button').forEach(function (b) { b.addEventListener('click', function () { show(b.dataset.view); }); });

  function setSync() {
    var el = $('#sync');
    var online = navigator.onLine;
    var n = queue().length;
    el.className = 'sync ' + (online ? 'online' : 'offline');
    $('#syncText').textContent = DEMO ? 'Demo' : !online ? 'Offline' : n ? n + ' pending' : 'Synced';
  }
  window.addEventListener('online', function () { flush().then(refresh); });
  window.addEventListener('offline', setSync);

  var toastTimer;
  function toast(msg) {
    var t = $('#toast'); t.textContent = msg; t.classList.add('show');
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.classList.remove('show'); }, 3000);
  }

  /* ---------- utils ---------- */
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function fmtDate(s) {
    if (!s) return '—';
    var d = new Date(String(s).slice(0, 10) + 'T00:00:00');
    return isNaN(d) ? esc(s) : d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  }
  function uid() {
    return (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
  }

  /* ---------- boot ---------- */
  $('#demoBanner').hidden = !DEMO;
  applyLists();
  resetForm();
  renderList();
  renderPending();
  flush().then(refresh);

  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').catch(function () {});
  }
})();
