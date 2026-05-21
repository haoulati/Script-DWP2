// ==UserScript==
// @name         DWP2 - At Station
// @namespace    DWP2
// @version      1.5
// @description  Affiche les packages Pending Stow par allée (A et B uniquement) pour DWP2
// @author       haoulati@
// @match        https://logistics.amazon.co.uk/station/dashboard/*
// @updateURL    https://raw.githubusercontent.com/haoulati/Script-DWP2/main/DWP2_AtStation.user.js
// @downloadURL  https://raw.githubusercontent.com/haoulati/Script-DWP2/main/DWP2_AtStation.user.js
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @run-at       document-end
// ==/UserScript==

(function() {
'use strict';

var motherStation = 'DWP2';
var fingers = ['A', 'B'];

// ── STYLES ───────────────────────────────────────────────────────────────────
GM_addStyle(`
  #_as_overlay {
    position: fixed; inset: 0; z-index: 2147483647;
    background: #f4f6f9; display: none; flex-direction: column;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  }
  #_as_header {
    background: #1a253c; color: #fff; padding: 10px 20px;
    display: flex; align-items: center; gap: 15px; flex-shrink: 0;
    box-shadow: 0 2px 8px rgba(0,0,0,.3);
  }
  #_as_header h1 { margin: 0; font-size: 1.1rem; flex-grow: 1; }
  #_as_close {
    background: #ff3b30; color: #fff; border: none; border-radius: 50%;
    width: 28px; height: 28px; font-size: 1.2rem; font-weight: 700;
    cursor: pointer; display: flex; align-items: center; justify-content: center;
  }
  #_as_filters {
    padding: 10px 20px; background: #fff; border-bottom: 1px solid #eef1f5;
    display: flex; align-items: center; gap: 20px; flex-shrink: 0;
  }
  #_as_filters label { display: flex; align-items: center; gap: 6px; font-size: .9rem; cursor: pointer; }
  #_as_status { font-size: .85rem; color: #6c757d; margin-left: auto; }
  #_as_body {
    flex-grow: 1; overflow-y: auto; padding: 16px 20px;
    display: flex; gap: 16px; align-items: flex-start;
  }
  .as-table-wrap { flex: 1; background: #fff; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,.08); overflow: hidden; }
  .as-table-wrap table { width: 100%; border-collapse: collapse; font-size: 13px; }
  .as-finger { text-align: center; font-size: 18px; font-weight: bold; background: #1a253c; color: #fff; padding: 10px; }
  .as-table-wrap thead tr:nth-child(2) th { background: #f8f9fa; padding: 8px 6px; border-bottom: 2px solid #eef1f5; font-size: 12px; white-space: nowrap; }
  .as-table-wrap tbody td { padding: 6px; border-bottom: 1px solid #eef1f5; white-space: nowrap; font-size: 12px; }
  .as-table-wrap tbody tr:hover { background: #f0f4ff; }
  .as-table-wrap tbody td:nth-child(7) { font-size: 11px; }
  .as-table-wrap a { color: #4a90e2; text-decoration: none; }
  .as-table-wrap a:hover { text-decoration: underline; }
  .as-count { font-size: 12px; font-weight: normal; color: #aaa; margin-left: 8px; }
  #_as_bubble {
    position: fixed; bottom: 140px; right: 20px;
    width: 52px; height: 52px; border-radius: 50%;
    background: #e67e22; color: #fff; font-size: 22px;
    border: none; cursor: pointer; z-index: 99995;
    box-shadow: 0 4px 12px rgba(0,0,0,.3);
    display: flex; align-items: center; justify-content: center;
  }
  .as-minutes-high { color: #ff3b30; font-weight: bold; }
  .as-minutes-med  { color: #ff9500; font-weight: bold; }
`);

// ── BUILD UI ──────────────────────────────────────────────────────────────────
function buildUI() {
  if (document.getElementById('_as_overlay')) return;

  var bubble = document.createElement('button');
  bubble.id = '_as_bubble';
  bubble.title = 'At Station - Pending Stow';
  bubble.innerHTML = '&#128205;';
  bubble.addEventListener('click', toggleOverlay);
  document.body.appendChild(bubble);

  var overlay = document.createElement('div');
  overlay.id = '_as_overlay';
  overlay.innerHTML = `
    <div id="_as_header">
      <h1>&#128205; At Station — Pending Stow (DWP2)</h1>
      <label style="color:#fff;font-size:.85rem;display:flex;align-items:center;gap:6px;">
        <input type="checkbox" id="_as_xpt" checked> XPT
      </label>
      <label style="color:#fff;font-size:.85rem;display:flex;align-items:center;gap:6px;">
        <input type="checkbox" id="_as_dwp2" checked> DWP2
      </label>
      <span id="_as_status" style="color:#aaa;font-size:.82rem;">Chargement...</span>
      <button id="_as_refresh" style="padding:4px 12px;border-radius:5px;border:none;background:#4a90e2;color:#fff;cursor:pointer;font-size:.82rem;">&#8635; Refresh</button>
      <button id="_as_close">&#215;</button>
    </div>
    <div id="_as_body">
      ${fingers.map(c => `
        <div class="as-table-wrap" id="_as_wrap_${c}">
          <table>
            <thead>
              <tr><th colspan="9" class="as-finger">Allée ${c} <span class="as-count" id="_as_count_${c}"></span></th></tr>
              <tr>
                <th>Tracking ID</th>
                <th>Minutes</th>
                <th>Associate</th>
                <th>Location</th>
                <th>Sort Zone</th>
                <th>Node</th>
                <th>Taille (LxWxH)</th>
                <th>Poids</th>
                <th>Reason</th>
              </tr>
            </thead>
            <tbody id="_as_tbody_${c}"></tbody>
          </table>
        </div>`).join('')}
    </div>`;
  document.body.appendChild(overlay);

  document.getElementById('_as_close').addEventListener('click', toggleOverlay);
  document.getElementById('_as_refresh').addEventListener('click', function() {
    // Force refresh : reset le flag loading et relance immédiatement
    _isLoading = false;
    if (_refreshTimer) clearTimeout(_refreshTimer);
    setStatus('Actualisation...');
    getAtStation();
  });
  document.getElementById('_as_xpt').addEventListener('change', applyFilters);
  document.getElementById('_as_dwp2').addEventListener('change', applyFilters);
}

function toggleOverlay() {
  var o = document.getElementById('_as_overlay');
  if (!o) return;
  if (o.style.display === 'flex') {
    o.style.display = 'none';
  } else {
    o.style.display = 'flex';
    getAtStation();
  }
}

function applyFilters() {
  var showXPT  = document.getElementById('_as_xpt').checked;
  var showDWP2 = document.getElementById('_as_dwp2').checked;
  fingers.forEach(function(c) {
    var tbody = document.getElementById('_as_tbody_' + c);
    if (!tbody) return;
    tbody.querySelectorAll('tr').forEach(function(row) {
      var isXPT  = row.classList.contains('XPT');
      var isDWP2 = row.classList.contains('DWP2');
      row.style.display = (isXPT && showXPT) || (isDWP2 && showDWP2) ? '' : 'none';
    });
    updateCount(c);
  });
}

function updateCount(cluster) {
  var tbody = document.getElementById('_as_tbody_' + cluster);
  var countEl = document.getElementById('_as_count_' + cluster);
  if (!tbody || !countEl) return;
  var visible = tbody.querySelectorAll('tr:not([style*="display: none"])').length;
  countEl.textContent = '(' + visible + ' packages)';
}

// ── STEAL API KEY from SCC native requests ────────────────────────────────────
var _apiKey = 'scc-boson-api-k8x7m2n4p9q1r3s5:1776808103761:SMszXG+e4hj9HfXeAXe5012G+BVpAzu7tYxc6+N0siA=';
(function(){
  try {
    var origSetHeader = XMLHttpRequest.prototype.setRequestHeader;
    XMLHttpRequest.prototype.setRequestHeader = function(k, v){
      if (k && k.toLowerCase() === 'x-api-usage-key' && v) _apiKey = v;
      return origSetHeader.apply(this, arguments);
    };
  } catch(e) {}
  try {
    var origFetch = window.fetch;
    window.fetch = function(url, opts){
      if (opts && opts.headers) {
        var h = opts.headers;
        var key = null;
        if (h instanceof Headers) key = h.get('x-api-usage-key') || h.get('X-Api-Usage-Key');
        else if (typeof h === 'object') key = h['X-Api-Usage-Key'] || h['x-api-usage-key'];
        if (key) _apiKey = key;
      }
      return origFetch.apply(this, arguments);
    };
  } catch(e) {}
})();

// ── MEMORY : garde le temps réel basé sur le premier scan ────────────────────
// Nouvelle approche : on stocke le timestamp de première apparition du colis.
// Les minutes affichées = (maintenant - première apparition) en minutes.
// La colonne Reason affiche la dernière intervention/problème détecté sur le colis.
var _memory = {}; // { tid: { firstSeen, lastSeenTime, cluster, sc, location, bin, associate, node, size, weight, statusHistory, reason } }

// ── TIME RANGE ────────────────────────────────────────────────────────────────
function getTimeRange() {
  var now = new Date();
  var h = now.getHours();
  var cutoff = 16;
  var start, end;
  if (h >= cutoff) {
    start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), cutoff, 0, 0);
    end   = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, cutoff, 0, 0);
  } else {
    start = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, cutoff, 0, 0);
    end   = new Date(now.getFullYear(), now.getMonth(), now.getDate(), cutoff, 0, 0);
  }
  return { startTime: Math.floor(start.getTime() / 1000), endTime: Math.floor(end.getTime() / 1000) };
}

// ── FETCH DATA ────────────────────────────────────────────────────────────────
var _refreshTimer = null;
var _atStation = {};
var _isLoading = false; // empêche les appels en doublon
var REFRESH_INTERVAL = 20 * 1000; // 20 secondes

function getAtStation() {
  if (_refreshTimer) clearTimeout(_refreshTimer);
  if (_isLoading) return; // déjà en cours, ne pas re-fetch
  _isLoading = true;
  setStatus('Chargement...');

  var range = getTimeRange();
  var requestBody = {
    resourcePath: '/os/getDrillDownData',
    httpMethod: 'post',
    processName: 'oculus',
    requestBody: {
      nodeId: motherStation,
      packageStatusMap: {
        'Inducted': [],
        'Stow Problem Solve': [],
        'Stow Buffered': []
      },
      filters: [
        {
          '__type': 'TermFilter:http://internal.amazon.com/coral/com.amazon.oculusservice.model.filter/',
          filterMap: { CYCLE: ['CYCLE_1'], SHIPMENT_TYPE: ['Delivery'] }
        },
        {
          '__type': 'RangeFilter:http://internal.amazon.com/coral/com.amazon.oculusservice.model.filter/',
          filterMap: {}
        }
      ],
      lastUpdatedRange: { startTime: range.startTime, endTime: range.endTime },
      size: 60000,
      startingIndex: 0
    }
  };

  GM_xmlhttpRequest({
    method: 'POST',
    url: 'https://logistics.amazon.co.uk/station/proxyapigateway/data',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'X-Requested-With': 'XMLHttpRequest',
      'X-Api-Usage-Key': _apiKey
    },
    data: JSON.stringify(requestBody),
    withCredentials: true,
    onload: function(response) {
      try {
        if (!response.responseText || response.responseText.trim() === '') {
          setStatus('Réponse vide (HTTP ' + response.status + '). Vérifie ta connexion SCC.');
          _isLoading = false;
          _refreshTimer = setTimeout(getAtStation, REFRESH_INTERVAL);
          return;
        }
        var js = JSON.parse(response.responseText);
        if (js.packageResultList) {
          processPackages(js.packageResultList);
        } else {
          setStatus('Aucun package trouvé (HTTP ' + response.status + ').');
        }
      } catch(e) {
        setStatus('Erreur HTTP ' + response.status + ' — ' + response.responseText.substring(0, 80));
      }
      _isLoading = false;
      _refreshTimer = setTimeout(getAtStation, REFRESH_INTERVAL);
    },
    onerror: function() {
      setStatus('Erreur réseau. Retry dans 20s.');
      _isLoading = false;
      _refreshTimer = setTimeout(getAtStation, REFRESH_INTERVAL);
    }
  });
}

function processPackages(packages) {
  var now = Date.now();

  // Reset display data
  _atStation = {};
  fingers.forEach(function(c) { _atStation[c] = {}; });

  // Set of TIDs currently returned by API
  var currentTids = new Set(packages.map(function(p) { return p.trackingId; }));

  packages.forEach(function(pkg) {
    // ✅ Extraire la raison/intervention depuis le champ "reason" de l'API
    // Valeurs connues : "SUCCESSFUL" (normal), "DAMAGED", "WRONG_AISLE", "REPACKED", etc.
    var rawReason = (pkg.reason || '').trim();
    var intervention = '';
    if (rawReason && rawReason !== 'SUCCESSFUL') {
      intervention = rawReason.replace(/_/g, ' ');
      // Capitaliser : "WRONG AISLE" → "Wrong Aisle"
      intervention = intervention.replace(/\b\w+/g, function(w) {
        return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
      });
    }

    var cluster = pkg.cluster;
    if (!fingers.includes(cluster)) return;

    var tid = pkg.trackingId;
    var sc = pkg.stationCode === motherStation ? 'DWP2' : 'XPT';

    if (!_memory[tid]) {
      // ✅ Première fois qu'on voit ce colis → on note le timestamp actuel
      // On utilise minutesInState de l'API pour estimer quand il est arrivé
      var apiMinutes = parseInt(pkg.minutesInState) || 0;
      _memory[tid] = {
        firstSeen:    now - (apiMinutes * 60 * 1000), // estimation du moment d'arrivée
        lastSeenTime: now,
        lastApiStatus: pkg.state || '',
        cluster:      cluster,
        sc:           sc,
        location:     pkg.lastScanLocation || pkg.location || '',
        bin:          pkg.sortZone || '',
        associate:    pkg.lastScanBy || '',
        node:         pkg.stationCode || '',
        size:         (pkg.length||'?')+' x '+(pkg.width||'?')+' x '+(pkg.height||'?'),
        weight:       pkg.weight || '',
        statusChanges: 0,
        reason:       intervention,
        statusHistory: [pkg.state || '']
      };
      // Si le colis arrive déjà avec un statut "problème", noter la raison
      if (pkg.state === 'PROBLEM_SOLVE' && !intervention) {
        _memory[tid].reason = 'Problem Solve';
      }
    } else {
      var mem = _memory[tid];
      mem.lastSeenTime = now;

      // ✅ Détection de changement de statut RÉEL :
      // On compare le packageStatus actuel au précédent
      var currentStatus = pkg.state || '';
      if (mem.lastApiStatus && currentStatus && currentStatus !== mem.lastApiStatus) {
        mem.statusChanges++;
        // Garder l'historique des transitions
        if (!mem.statusHistory.includes(currentStatus)) {
          mem.statusHistory.push(currentStatus);
        }
        // Si on a une intervention spécifique, l'utiliser comme reason
        if (intervention) {
          mem.reason = intervention;
        } else {
          mem.reason = mem.lastApiStatus.replace(/_/g, ' ').toLowerCase()
            + ' → '
            + currentStatus.replace(/_/g, ' ').toLowerCase();
        }
      } else if (intervention && intervention !== mem.reason) {
        // ✅ Même si le statut n'a pas changé, si une intervention apparaît, on la note
        mem.reason = intervention;
      }
      mem.lastApiStatus = currentStatus;

      // Mise à jour des infos
      mem.location  = pkg.lastScanLocation || pkg.location || mem.location;
      mem.bin       = pkg.sortZone || mem.bin;
      mem.associate = pkg.lastScanBy || mem.associate;
      mem.node      = pkg.stationCode || mem.node;
      mem.cluster   = cluster;
      mem.sc        = sc;
    }

    var mem = _memory[tid];
    // ✅ Minutes = temps écoulé depuis la première apparition (stable, pas de fluctuation)
    var realMinutes = Math.round((now - mem.firstSeen) / 60000);

    _atStation[cluster][tid] = {
      location:            mem.location,
      bin:                 mem.bin,
      lastScanAssociateId: mem.associate,
      minutesInState:      realMinutes,
      node:                mem.node,
      size:                mem.size,
      weight:              mem.weight,
      sc:                  mem.sc,
      changedStatus:       mem.statusChanges > 0 || mem.reason !== '',
      reason:              mem.reason
    };
  });

  // ✅ Colis disparus de l'API → probablement stowés → on les retire
  // On ne les garde plus artificiellement (source de bugs)
  Object.keys(_memory).forEach(function(tid) {
    if (currentTids.has(tid)) return;

    var mem = _memory[tid];
    var minutesSinceLastSeen = Math.round((now - mem.lastSeenTime) / 60000);

    // Si le colis n'est plus dans l'API depuis plus de 1 refresh (>4 min), on le supprime
    if (minutesSinceLastSeen > 4) {
      delete _memory[tid];
    } else {
      // Garder 1 refresh de grâce (en cas de fluctuation API)
      var cluster = mem.cluster;
      if (!_atStation[cluster]) _atStation[cluster] = {};
      var realMinutes = Math.round((now - mem.firstSeen) / 60000);
      _atStation[cluster][tid] = {
        location:            mem.location,
        bin:                 mem.bin,
        lastScanAssociateId: mem.associate,
        minutesInState:      realMinutes,
        node:                mem.node,
        size:                mem.size,
        weight:              mem.weight,
        sc:                  mem.sc,
        changedStatus:       mem.statusChanges > 0 || mem.reason !== '',
        reason:              mem.reason || 'Disparu de l\'API'
      };
    }
  });

  renderTables();
  setStatus('Mis à jour : ' + new Date().toLocaleTimeString('fr-FR'));
}

function renderTables() {
  fingers.forEach(function(cluster) {
    var tbody = document.getElementById('_as_tbody_' + cluster);
    var wrap  = document.getElementById('_as_wrap_' + cluster);
    if (!tbody || !wrap) return;

    var data = _atStation[cluster] || {};
    var keys = Object.keys(data);

    if (keys.length === 0) {
      wrap.style.display = 'none';
      return;
    }
    wrap.style.display = 'block';

    // Remove packages no longer present
    tbody.querySelectorAll('tr').forEach(function(row) {
      if (!data[row.dataset.tid]) row.remove();
    });

    // Update or add rows
    keys.forEach(function(tid) {
      var d = data[tid];
      var existing = tbody.querySelector('tr[data-tid="' + tid + '"]');
      var minClass = d.minutesInState >= 60 ? 'as-minutes-high' : d.minutesInState >= 30 ? 'as-minutes-med' : '';

      if (existing) {
        existing.cells[1].textContent = d.minutesInState;
        existing.cells[1].className = minClass;
        existing.cells[2].textContent = d.lastScanAssociateId;
        existing.cells[3].textContent = d.location;
        existing.cells[4].textContent = d.bin;
        existing.cells[5].textContent = d.node;
        existing.cells[6].textContent = d.size;
        existing.cells[7].textContent = d.weight;
        if (d.changedStatus) {
          existing.cells[8].textContent = d.reason;
          existing.style.background = '#fff3cd';
        } else {
          existing.cells[8].textContent = '';
          existing.style.background = '';
        }
      } else {
        var tr = document.createElement('tr');
        tr.dataset.tid = tid;
        tr.className = d.sc;
        if (d.changedStatus) tr.style.background = '#fff3cd';
        tr.innerHTML = `
          <td><a href="https://logistics.amazon.co.uk/station/dashboard/search?shareableLink=detailPage%2F${tid}" target="_blank">${tid}</a></td>
          <td class="${minClass}">${d.minutesInState}</td>
          <td>${d.lastScanAssociateId}</td>
          <td>${d.location}</td>
          <td>${d.bin}</td>
          <td>${d.node}</td>
          <td>${d.size}</td>
          <td>${d.weight}</td>
          <td style="color:#e67e22;font-weight:600;font-size:11px;">${d.changedStatus ? d.reason : ''}</td>`;
        tbody.appendChild(tr);
      }
    });

    // Sort by minutes descending
    var rows = Array.from(tbody.querySelectorAll('tr'));
    rows.sort(function(a, b) {
      return parseInt(b.cells[1].textContent) - parseInt(a.cells[1].textContent);
    });
    rows.forEach(function(r) { tbody.appendChild(r); });

    updateCount(cluster);
  });

  applyFilters();
}

function setStatus(msg) {
  var el = document.getElementById('_as_status');
  if (el) el.textContent = msg;
}

// ── INIT ──────────────────────────────────────────────────────────────────────
function init() {
  if (document.body) {
    buildUI();
  } else {
    setTimeout(init, 500);
  }
}

init();

})();
