/* Katapult NESC QC Map - client-only app (no backend)
   - Parses Katapult Job JSON in a Web Worker (inline worker-js)
   - Runs QC rule checks in the main thread (re-runs instantly when Rules change)
   - Visualizes poles + midspan measurement points on a Leaflet map
*/

(() => {
  "use strict";

  // ────────────────────────────────────────────────────────────────────────────
  //  Defaults
  // ────────────────────────────────────────────────────────────────────────────

  const DEFAULT_RULES = {
    pole: {
      minLowestCommAttachIn: 16 * 12,      // 16' 0"
      commSepDiffIn: 12,                   // 12" between different comm companies
      commSepSameIn: 4,                    // 4" minimum between same comm company (unless exact same height)
      commToPowerSepIn: 40,                // 40" pole clearance comm-to-lowest-power (excluding streetlight drip loop)
      adssCommToPowerSepIn: 30,            // ADSS comms: 30" min to nearest power (secondary/neutral/secondary drip loop)
      commToStreetLightSepIn: 12,          // 12" comm-to-streetlight equipment
      movedHoleBufferIn: 4,                // 4" keep-out around moved-from holes
      enforceAdssHighest: true,
      enforceEquipmentMove: true,
      enforcePowerOrder: true,
      warnMissingPoleIdentifiers: true,
    },
    midspan: {
      minCommDefaultIn: 15 * 12 + 6,       // 15' 6"
      minCommPedestrianIn: 9 * 12 + 6,     // 9' 6"
      minCommHighwayIn: 18 * 12,           // 18' 0"
      minCommFarmIn: 18 * 12,              // 18' 0"
      minCommRailIn: 23 * 12 + 6,          // 23' 6"
      commSepIn: 4,                        // 4" between comms (NESC-style baseline)
      commToPowerSepIn: 30,                // 30" midspan clearance comm-to-lowest-power
      adssCommToPowerSepIn: 12,            // ADSS midspan comm-to-power min (inches)
      installingCompany: "",
      installingCompanyCommSepIn: 4,       // per-utility stricter standard for installing company
      enforceAdssHighest: true,
      warnMissingRowType: true,
    },
  };

  const STORAGE_KEY = "katapultQcRules.v1";

  // ────────────────────────────────────────────────────────────────────────────
  //  DOM helpers
  // ────────────────────────────────────────────────────────────────────────────

  const $ = (id) => document.getElementById(id);

  const els = {
    fileInput: $("fileInput"),
    btnPreview: $("btnPreview"),
    btnRunQc: $("btnRunQc"),
    btnReset: $("btnReset"),
    btnResetRules: $("btnResetRules"),
    btnZoomAll: $("btnZoomAll"),

    jobName: $("jobName"),
    summaryPoles: $("summaryPoles"),
    summaryMidspans: $("summaryMidspans"),
    summaryIssues: $("summaryIssues"),

    progressBar: $("progressBar"),
    progressLabel: $("progressLabel"),
    logBox: $("logBox"),

    // Map toggles / filters
    togglePoles: $("togglePoles"),
    toggleMidspans: $("toggleMidspans"),
    toggleSpans: $("toggleSpans"),
    toggleScidLabels: $("toggleScidLabels"),
    filterPass: $("filterPass"),
    filterWarn: $("filterWarn"),
    filterFail: $("filterFail"),
    searchPole: $("searchPole"),

    detailsPanel: $("detailsPanel"),
    btnCloseDetails: $("btnCloseDetails"),
    details: $("details"),

    // Rules inputs
    rulePoleMinAttach: $("rulePoleMinAttach"),
    rulePoleCommSepDiff: $("rulePoleCommSepDiff"),
    rulePoleCommSepSame: $("rulePoleCommSepSame"),
    rulePoleCommToPower: $("rulePoleCommToPower"),
    rulePoleAdssCommToPower: $("rulePoleAdssCommToPower"),
    rulePoleCommToStreet: $("rulePoleCommToStreet"),
    rulePoleHoleBuffer: $("rulePoleHoleBuffer"),
    rulePoleEnforceAdss: $("rulePoleEnforceAdss"),
    rulePoleEnforceEquipMove: $("rulePoleEnforceEquipMove"),
    rulePoleEnforcePowerOrder: $("rulePoleEnforcePowerOrder"),
    rulePoleWarnMissingIds: $("rulePoleWarnMissingIds"),

    ruleMidMinDefault: $("ruleMidMinDefault"),
    ruleMidMinPed: $("ruleMidMinPed"),
    ruleMidMinHwy: $("ruleMidMinHwy"),
    ruleMidMinFarm: $("ruleMidMinFarm"),
    ruleMidCommSep: $("ruleMidCommSep"),
    ruleMidCommToPower: $("ruleMidCommToPower"),
    ruleMidAdssCommToPower: $("ruleMidAdssCommToPower"),
    ruleInstallingCompany: $("ruleInstallingCompany"),
    ruleMidCommSepInstall: $("ruleMidCommSepInstall"),
    ruleMidEnforceAdss: $("ruleMidEnforceAdss"),
    ruleMidWarnMissingRow: $("ruleMidWarnMissingRow"),

    // Issues
    issuesFail: $("issuesFail"),
    issuesWarn: $("issuesWarn"),
    issuesSearch: $("issuesSearch"),
    btnExportIssues: $("btnExportIssues"),
    issuesTbody: $("issuesTbody"),
  };

  // ────────────────────────────────────────────────────────────────────────────
  //  State
  // ────────────────────────────────────────────────────────────────────────────

  let fileBuffer = null;
  let fileName = "";

  let worker = null;
  let workerUrl = null;

  let previewPoles = [];        // preview payload poles
  let model = null;             // full payload: poles + attachments + midspanPoints

  let rules = loadRules();
  let qcResults = null;

  // Map state
  let map = null;
  let poleCluster = null;
  let midspanCluster = null;
  let spanLayer = null;
  let scidLabelLayer = null;

  const poleMarkers = new Map();      // poleId -> marker
  const midspanMarkers = new Map();   // midspanId -> marker
  const spanPolylines = [];           // L.Polyline[]
  const scidLabelMarkers = new Map(); // poleId -> label marker

  // ────────────────────────────────────────────────────────────────────────────
  //  Init
  // ────────────────────────────────────────────────────────────────────────────

  function init() {
    initTabs();
    initMap();
    initEvents();

    // Maximize initial map viewport area.
    maximizeMapHeight();
    if (map) map.invalidateSize();

    // Keep map sized to available viewport space.
    window.addEventListener("resize", () => {
      maximizeMapHeight();
      if (map) map.invalidateSize();
    });

    applyRulesToUi();
    updateUiEnabled(false);
    setProgress(0, "—");
    logLine("Ready.");
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Tabs
  // ────────────────────────────────────────────────────────────────────────────

  function initTabs() {
    const tabButtons = Array.from(document.querySelectorAll(".tab-btn"));
    tabButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        tabButtons.forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");

        const tab = btn.getAttribute("data-tab");
        Array.from(document.querySelectorAll(".tab-panel")).forEach((p) => p.classList.remove("active"));
        const panel = document.getElementById(`tab-${tab}`);
        if (panel) panel.classList.add("active");

        // Leaflet needs invalidateSize when its container becomes visible
        if (tab === "map" && map) {
          setTimeout(() => {
            maximizeMapHeight();
            map.invalidateSize();
          }, 50);
        }
      });
    });
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Worker
  // ────────────────────────────────────────────────────────────────────────────

  function createInlineWorker() {
    const script = document.getElementById("worker-js");
    if (!script) throw new Error("Inline worker script not found.");
    const js = script.textContent || "";
    const blob = new Blob([js], { type: "application/javascript" });
    const url = URL.createObjectURL(blob);
    const w = new Worker(url);
    return { w, url };
  }

  function ensureWorker() {
    if (worker) return worker;
    const { w, url } = createInlineWorker();
    worker = w;
    workerUrl = url;

    worker.onmessage = (ev) => {
      const msg = ev.data || {};
      if (msg.type === "progress") {
        setProgress(msg.pct || 0, msg.label || "");
      } else if (msg.type === "log") {
        logLine(msg.message || "");
      } else if (msg.type === "error") {
        logLine(`ERROR: ${msg.message || "Unknown error"}`);
        setProgress(0, "Error");
      } else if (msg.type === "preview") {
        onPreview(msg.payload);
      } else if (msg.type === "done") {
        onModel(msg.payload);
      }
    };

    return worker;
  }

  function disposeWorker() {
    try {
      if (worker) worker.terminate();
    } catch (_) {}
    worker = null;
    if (workerUrl) URL.revokeObjectURL(workerUrl);
    workerUrl = null;
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Map
  // ────────────────────────────────────────────────────────────────────────────

  function initMap() {
    map = L.map("map", {
      preferCanvas: true,
      zoomControl: false,
    }).setView([39.5, -98.35], 4);

    // Panes (z-index ordering)
    map.createPane("paneSpans");
    map.getPane("paneSpans").style.zIndex = 310;

    map.createPane("paneMidspans");
    map.getPane("paneMidspans").style.zIndex = 410;

    map.createPane("panePoles");
    map.getPane("panePoles").style.zIndex = 420;

    map.createPane("paneLabels");
    map.getPane("paneLabels").style.zIndex = 430;

    // Basemaps
    const baseDark = L.tileLayer("https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png", {
      maxZoom: 20,
      attribution: "&copy; OpenStreetMap contributors &copy; CARTO",
    });
    const baseLight = L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
      maxZoom: 20,
      attribution: "&copy; OpenStreetMap contributors &copy; CARTO",
    });
    const baseImagery = L.tileLayer("https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}", {
      maxZoom: 20,
      attribution: "Tiles &copy; Esri",
    });

    baseDark.addTo(map);

    L.control.layers(
      {
        "Dark": baseDark,
        "Light": baseLight,
        "Imagery": baseImagery,
      },
      {},
      { position: "topright", collapsed: true }
    ).addTo(map);

    L.control.zoom({ position: "topright" }).addTo(map);
    L.control.scale({ position: "bottomright", imperial: true, metric: false, maxWidth: 140 }).addTo(map);

    // NOTE: marker clustering is intentionally disabled.
    // This QC workflow requires every pole and midspan point to remain visible.
    poleCluster = L.layerGroup();
    midspanCluster = L.layerGroup();
    spanLayer = L.layerGroup();
    scidLabelLayer = L.layerGroup();

    map.addLayer(poleCluster);
    map.addLayer(midspanCluster);

    // Default: span lines on (user can toggle off)
    if (els.toggleSpans && els.toggleSpans.checked) map.addLayer(spanLayer);

    // Default: SCID labels off
    if (els.toggleScidLabels && els.toggleScidLabels.checked) map.addLayer(scidLabelLayer);
  }

  function maximizeMapHeight() {
    // Expand the map to use as much of the viewport as possible, without requiring a full-screen mode.
    // This makes the QC map feel more like a dedicated workspace.
    const el = document.getElementById("map");
    if (!el) return;
    // If the map tab is hidden, measurements are unreliable.
    if (el.offsetParent === null) return;

    const rect = el.getBoundingClientRect();
    // Keep a small gutter so the map doesn't visually collide with the page edge.
    const bottomPad = 10;
    const base = Math.max(0, Math.floor(window.innerHeight - rect.top - bottomPad));

    // User request: make the map workspace ~2× taller than the prior build.
    // This gives more vertical room for the in-map details panel (less scrolling) and a larger canvas.
    const h = Math.max(720, Math.floor(base * 2));
    el.style.height = `${h}px`;
  }

  function makeIcon(status, shape, hasOrderHalo = false) {
    const st = status || "unknown";
    const sh = shape || "pole";

    // Midspans: explicit 3-sided marker (SVG) to avoid browser-specific CSS triangle quirks.
    let html;
    let iconSize = [16, 16];
    let iconAnchor = [8, 8];

    if (sh === "midspan") {
      html = `
        <svg class="marker-tri marker-tri--${st}" viewBox="0 0 20 18" aria-hidden="true">
          <path d="M10 1 L19 17 H1 Z"></path>
        </svg>
      `;
      iconSize = [16, 14];
      iconAnchor = [8, 7];
    } else {
      html = `<div class="marker-dot marker-dot--${sh} marker-dot--${st}"></div>`;
    }

    return L.divIcon({
      className: hasOrderHalo ? "marker-icon marker-icon--halo" : "marker-icon",
      html,
      iconSize,
      iconAnchor,
    });
  }

  function hashStringToInt(s) {
    // Deterministic small hash for stable label offsets.
    const str = String(s || "");
    let h = 2166136261;
    for (let i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    return h >>> 0;
  }

  function computeScidLabelOffset(poleId, scidText) {
    // Provide a small, deterministic offset so labels don't sit directly on top of the pole marker.
    // We also vary the direction to reduce local overlaps when poles are dense.
    const h = hashStringToInt(`${poleId}|${scidText}`);
    const slot = h % 8;
    const base = [
      [18, -18],
      [22, -4],
      [18, 14],
      [0, 18],
      [-18, 14],
      [-22, -4],
      [-18, -18],
      [0, -22],
    ][slot];
    const jx = ((h >>> 3) % 7) - 3; // -3..+3
    const jy = ((h >>> 6) % 7) - 3;
    return { dx: base[0] + jx, dy: base[1] + jy };
  }

  function makeScidLabelIcon(text, dx, dy) {
    const safe = escapeHtml(String(text || "").trim());
    const html = `<div class="scid-label" style="--dx:${dx}px; --dy:${dy}px">${safe}</div>`;
    return L.divIcon({
      className: "scid-label-icon",
      html,
      iconSize: [0, 0],
      iconAnchor: [0, 0],
    });
  }

  function clearMapData() {
    poleCluster.clearLayers();
    midspanCluster.clearLayers();
    spanLayer.clearLayers();
    scidLabelLayer.clearLayers();
    spanPolylines.length = 0;
    poleMarkers.clear();
    midspanMarkers.clear();
    scidLabelMarkers.clear();
  }

  function renderPreviewPoles(poles) {
    clearMapData();

    if (!Array.isArray(poles) || !poles.length) return;

    let bounds = null;

    for (const p of poles) {
      const lat = typeof p.lat === "number" ? p.lat : null;
      const lon = typeof p.lon === "number" ? p.lon : null;
      if (lat == null || lon == null) continue;

      const marker = L.marker([lat, lon], {
        icon: makeIcon("unknown", "pole"),
        title: p.displayName || "",
        pane: "panePoles",
        zIndexOffset: 600,
      });

      marker.on("click", (ev) => selectEntity({ type: "pole", id: p.poleId, latlng: ev && ev.latlng ? ev.latlng : null }));

      poleMarkers.set(String(p.poleId), marker);
      poleCluster.addLayer(marker);

      // Optional SCID label marker (non-interactive).
      const scidText = String(p.scid || "").trim();
      if (scidText) {
        const { dx, dy } = computeScidLabelOffset(String(p.poleId), scidText);
        const lbl = L.marker([lat, lon], {
          icon: makeScidLabelIcon(scidText, dx, dy),
          interactive: false,
          pane: "paneLabels",
          zIndexOffset: 650,
        });
        scidLabelMarkers.set(String(p.poleId), lbl);
      }

      const ll = L.latLng(lat, lon);
      bounds = bounds ? bounds.extend(ll) : L.latLngBounds(ll, ll);
    }

    if (bounds) map.fitBounds(bounds.pad(0.10));
  }

  function renderModelMidspans(midspanPoints) {
    if (!Array.isArray(midspanPoints)) return;

    for (const ms of midspanPoints) {
      const lat = typeof ms.lat === "number" ? ms.lat : null;
      const lon = typeof ms.lon === "number" ? ms.lon : null;
      if (lat == null || lon == null) continue;

      const marker = L.marker([lat, lon], {
        icon: makeIcon("unknown", "midspan"),
        title: `Midspan (${ms.rowTypeRaw || ms.rowType || "default"})`,
        pane: "paneMidspans",
        zIndexOffset: 300,
      });

      marker.on("click", (ev) => selectEntity({ type: "midspan", id: ms.midspanId, latlng: ev && ev.latlng ? ev.latlng : null }));

      midspanMarkers.set(String(ms.midspanId), marker);
      midspanCluster.addLayer(marker);
    }
  }

  function renderSpanLines(spans) {
    spanLayer.clearLayers();
    spanPolylines.length = 0;

    if (!Array.isArray(spans)) return;

    for (const s of spans) {
      if (s.aLat == null || s.aLon == null || s.bLat == null || s.bLon == null) continue;
      const dashed = !(s.aIsPole && s.bIsPole);
      const poly = L.polyline(
        [[s.aLat, s.aLon], [s.bLat, s.bLon]],
        {
          pane: "paneSpans",
          weight: dashed ? 1.75 : 2.25,
          opacity: dashed ? 0.22 : 0.32,
          color: dashed ? "rgba(148, 163, 184, 0.55)" : "rgba(56, 189, 248, 0.55)",
          dashArray: dashed ? "6 8" : null,
          lineCap: "round",
          lineJoin: "round",
        }
      );
      spanPolylines.push(poly);
      spanLayer.addLayer(poly);
    }
  }

  function zoomToAll() {
    const latlngs = [];
    for (const m of poleMarkers.values()) latlngs.push(m.getLatLng());
    for (const m of midspanMarkers.values()) latlngs.push(m.getLatLng());
    if (!latlngs.length) return;
    const b = L.latLngBounds(latlngs);
    map.fitBounds(b.pad(0.12));
  }

  function allowedStatusSet() {
    const s = new Set();
    if (els.filterPass.checked) s.add("pass");
    if (els.filterWarn.checked) s.add("warn");
    if (els.filterFail.checked) s.add("fail");
    // When no QC has run, markers are "unknown" — show them unless all filters are unchecked.
    if (s.size === 0) return new Set(["pass", "warn", "fail", "unknown"]);
    return s;
  }

  function refreshLayers() {
    const allowed = allowedStatusSet();

    poleCluster.clearLayers();
    if (els.togglePoles.checked) {
      for (const [poleId, marker] of poleMarkers.entries()) {
        const st = qcResults && qcResults.poles && qcResults.poles[poleId] ? qcResults.poles[poleId].status : "unknown";
        if (allowed.has(st)) poleCluster.addLayer(marker);
      }
    }

    // SCID labels follow pole visibility + status filters
    scidLabelLayer.clearLayers();
    if (els.toggleScidLabels && els.toggleScidLabels.checked && els.togglePoles.checked) {
      for (const [poleId, marker] of poleMarkers.entries()) {
        const st = qcResults && qcResults.poles && qcResults.poles[poleId] ? qcResults.poles[poleId].status : "unknown";
        if (!allowed.has(st)) continue;
        const lbl = scidLabelMarkers.get(poleId);
        if (lbl) scidLabelLayer.addLayer(lbl);
      }
    }

    midspanCluster.clearLayers();
    if (els.toggleMidspans.checked) {
      for (const [midId, marker] of midspanMarkers.entries()) {
        const st = qcResults && qcResults.midspans && qcResults.midspans[midId] ? qcResults.midspans[midId].status : "unknown";
        if (allowed.has(st)) midspanCluster.addLayer(marker);
      }
    }

    // Spans layer
    if (els.toggleSpans.checked) {
      if (!map.hasLayer(spanLayer)) map.addLayer(spanLayer);
    } else {
      if (map.hasLayer(spanLayer)) map.removeLayer(spanLayer);
    }

    // SCID labels layer
    if (els.toggleScidLabels && els.toggleScidLabels.checked) {
      if (!map.hasLayer(scidLabelLayer)) map.addLayer(scidLabelLayer);
    } else {
      if (map.hasLayer(scidLabelLayer)) map.removeLayer(scidLabelLayer);
    }
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Events
  // ────────────────────────────────────────────────────────────────────────────

  function initEvents() {
    els.fileInput.addEventListener("change", async (e) => {
      const f = e.target.files && e.target.files[0] ? e.target.files[0] : null;
      if (!f) return;

      fileName = f.name || "job.json";
      fileBuffer = await f.arrayBuffer();

      // Reset state
      previewPoles = [];
      model = null;
      qcResults = null;
      clearDetails();
      clearIssuesTable();

      setJobName("—");
      setSummaryCounts(null);
      setProgress(0, "Ready to preview…");
      logLine(`Loaded file: ${fileName}`);

      updateUiEnabled(true);

      // Auto preview for convenience
      runPreview();
    });

    els.btnPreview.addEventListener("click", () => runPreview());
    els.btnRunQc.addEventListener("click", () => runQc());
    els.btnReset.addEventListener("click", () => resetAll());
    els.btnResetRules.addEventListener("click", () => {
      rules = structuredClone(DEFAULT_RULES);
      saveRules(rules);
      applyRulesToUi();
      if (model) recomputeQc();
    });

    els.btnZoomAll.addEventListener("click", () => zoomToAll());

    // Map toggles/filters
    [els.togglePoles, els.toggleMidspans, els.toggleSpans, els.toggleScidLabels, els.filterPass, els.filterWarn, els.filterFail].forEach((el) => {
      el.addEventListener("change", () => refreshLayers());
    });

    // In-map details panel controls
    if (els.btnCloseDetails) {
      els.btnCloseDetails.addEventListener("click", () => closeDetailsPanel());
    }
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") closeDetailsPanel();
    });

    // Search
    els.searchPole.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        searchAndFocusPole(els.searchPole.value || "");
      }
    });

    // Rules change handlers
    const ruleInputs = [
      els.rulePoleMinAttach,
      els.rulePoleCommSepDiff,
      els.rulePoleCommSepSame,
      els.rulePoleCommToPower,
      els.rulePoleAdssCommToPower,
      els.rulePoleCommToStreet,
      els.rulePoleHoleBuffer,
      els.rulePoleEnforceAdss,
      els.rulePoleEnforceEquipMove,
      els.rulePoleEnforcePowerOrder,
      els.rulePoleWarnMissingIds,

      els.ruleMidMinDefault,
      els.ruleMidMinPed,
      els.ruleMidMinHwy,
      els.ruleMidMinFarm,
      els.ruleMidCommSep,
      els.ruleMidCommToPower,
      els.ruleMidAdssCommToPower,
      els.ruleInstallingCompany,
      els.ruleMidCommSepInstall,
      els.ruleMidEnforceAdss,
      els.ruleMidWarnMissingRow,
    ];

    ruleInputs.forEach((el) => {
      if (!el) return;
      el.addEventListener("change", () => {
        rules = readRulesFromUi();
        saveRules(rules);
        if (model) recomputeQc();
      });
    });

    // Issues filtering
    [els.issuesFail, els.issuesWarn, els.issuesSearch].forEach((el) => {
      if (!el) return;
      el.addEventListener("input", () => renderIssuesTable());
      el.addEventListener("change", () => renderIssuesTable());
    });

    els.btnExportIssues.addEventListener("click", () => exportIssuesCsv());
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Actions: Preview / Parse
  // ────────────────────────────────────────────────────────────────────────────

  function runPreview() {
    if (!fileBuffer) return;
    disposeWorker();
    ensureWorker();

    setProgress(0, "Previewing poles…");
    logLine("Starting preview parse…");

    worker.postMessage({
      type: "preview",
      buffer: fileBuffer,
      options: {
        nodeColorAttribute: "company",
      },
    });
  }

  function runQc() {
    if (!fileBuffer) return;

    // Prevent duplicate midspan markers / span lines if QC is run multiple times.
    try {
      midspanCluster && midspanCluster.clearLayers();
      midspanMarkers.clear();
      spanLayer && spanLayer.clearLayers();
      spanPolylines.length = 0;
    } catch (_) {}

    disposeWorker();
    ensureWorker();

    setProgress(0, "Parsing attachments & midspans…");
    logLine("Starting full parse…");

    worker.postMessage({
      type: "start",
      buffer: fileBuffer,
      options: {
        includeMidspans: true,
        includeEquipment: true,
        includeGuys: true,
        nodeColorAttribute: "company",
      },
    });
  }

  function resetAll() {
    fileBuffer = null;
    fileName = "";
    previewPoles = [];
    model = null;
    qcResults = null;

    disposeWorker();
    clearMapData();
    clearDetails();
    clearIssuesTable();
    setProgress(0, "—");
    setJobName("—");
    setSummaryCounts(null);
    els.fileInput.value = "";
    updateUiEnabled(false);
    logLine("Reset complete.");
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Worker callbacks
  // ────────────────────────────────────────────────────────────────────────────

  function onPreview(payload) {
    if (!payload) return;
    previewPoles = Array.isArray(payload.poles) ? payload.poles : [];
    setJobName(payload.jobName || "—");
    els.summaryPoles.textContent = String(previewPoles.length || 0);
    logLine(`Preview poles: ${previewPoles.length}`);

    renderPreviewPoles(previewPoles);
    // Optional in preview mode: show span lines immediately.
    if (Array.isArray(payload.spans) && payload.spans.length) {
      renderSpanLines(payload.spans);
    }
    refreshLayers();

    els.btnRunQc.disabled = false;
    els.btnZoomAll.disabled = false;
    els.btnReset.disabled = false;
    els.btnResetRules.disabled = false;
    els.btnExportIssues.disabled = true;

    setProgress(100, "Preview ready.");
  }

  function onModel(payload) {
    if (!payload) return;
    model = payload;

    // If pole markers are not present (e.g., user skipped preview), render poles from the model.
    if (poleMarkers.size === 0 && Array.isArray(model.poles) && model.poles.length) {
      renderPreviewPoles(model.poles);
    }

    logLine(`Model poles: ${Array.isArray(model.poles) ? model.poles.length : 0}`);
    logLine(`Model midspans: ${Array.isArray(model.midspanPoints) ? model.midspanPoints.length : 0}`);

    // Populate installing company dropdown
    populateCompanyDropdown(model.companies || []);

    // Add midspan markers + span lines
    renderModelMidspans(model.midspanPoints || []);
    renderSpanLines(model.spans || []);

    recomputeQc();

    els.btnExportIssues.disabled = false;
    els.btnZoomAll.disabled = false;
    setProgress(100, "QC ready.");
  }

  function populateCompanyDropdown(companies) {
    const list = Array.isArray(companies) ? companies.filter(Boolean) : [];
    const sel = els.ruleInstallingCompany;
    sel.innerHTML = "";

    const optNone = document.createElement("option");
    optNone.value = "";
    optNone.textContent = "— (not set)";
    sel.appendChild(optNone);

    for (const c of list) {
      const opt = document.createElement("option");
      opt.value = c;
      opt.textContent = c;
      sel.appendChild(opt);
    }

    // Keep existing selection if possible
    if (rules.midspan.installingCompany && list.includes(rules.midspan.installingCompany)) {
      sel.value = rules.midspan.installingCompany;
    } else {
      sel.value = "";
      rules.midspan.installingCompany = "";
      saveRules(rules);
    }
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  QC (core)
  // ────────────────────────────────────────────────────────────────────────────

  function recomputeQc() {
    if (!model) return;

    qcResults = runQcEngine(model, rules);
    updateMarkerStyles();
    refreshLayers();
    renderIssuesTable();
    setSummaryCounts(qcResults.summary);

    logLine(`QC complete. Issues: ${qcResults.issues.length}`);
  }

  function updateMarkerStyles() {
    // Poles
    for (const p of (model.poles || [])) {
      const poleId = String(p.poleId);
      const marker = poleMarkers.get(poleId);
      if (!marker) continue;

      const res = qcResults && qcResults.poles ? qcResults.poles[poleId] : null;
      const st = res && res.status ? res.status : "unknown";
      const halo = !!(res && res.hasCommOrderIssue);
      marker.setIcon(makeIcon(st, "pole", halo));

      const popupHtml = buildPolePopupHtml(poleId);
      marker.bindPopup(popupHtml, { maxWidth: 420 });
    }

    // Midspans
    for (const ms of (model.midspanPoints || [])) {
      const midId = String(ms.midspanId);
      const marker = midspanMarkers.get(midId);
      if (!marker) continue;

      const res = qcResults && qcResults.midspans ? qcResults.midspans[midId] : null;
      const st = res && res.status ? res.status : "unknown";
      const halo = !!(res && res.hasCommOrderIssue);
      marker.setIcon(makeIcon(st, "midspan", halo));

      const popupHtml = buildMidspanPopupHtml(midId);
      marker.bindPopup(popupHtml, { maxWidth: 420 });
    }
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Classification
  // ────────────────────────────────────────────────────────────────────────────

  function normalizeStr(s) {
    return String(s || "").toLowerCase();
  }

  function classify(item) {
    // item may be pole attachment or midspan measure
    const owner = String(item.owner || "").trim();
    const category = String(item.category || "Wire").trim();
    const text = [
      item.label,
      item.traceType,
      item.cableType,
      item.name,
      item.traceLabel,
      owner,
      category,
    ].filter(Boolean).join(" ");
    const s = normalizeStr(text);

    const isAdss = s.includes("adss");

    const isDownGuy = s.includes("down guy") || s.includes("down-guy") || s.includes("downguy") || s.includes("down_guy");

    // Riser equipment: allowed to move per the QC logic (see equipment movement rule).
    const isRiser = s.includes("riser");

    const isDripLoop = s.includes("drip loop") || s.includes("driploop");

    // Street light detection
    // IMPORTANT:
    // - "Street Light Feed" is a power utility wire (supply), NOT a communications facility.
    // - It also should NOT be treated as "street light equipment" for the 12" comm-to-streetlight
    //   equipment separation rule.
    const isStreetLightFeed = /\bstreet\s*light\s*feed\b/.test(s) || /\bstreetlight\s*feed\b/.test(s);
    const isStreetLight = (s.includes("street light") || s.includes("streetlight") || s.includes("luminaire")) && !isStreetLightFeed;
    const isStreetLightDripLoop = isDripLoop && isStreetLight;

    // Communications detection (used both for comm classification and to avoid misclassifying
    // communications "service drops" as secondary power service).
    const commWords = ["communication", "comm", "catv", "fiber", "telephone", "tel", "coax", "cable", "adss", "drop"];
    const looksComm = commWords.some(w => s.includes(w));

    // Power detection
    const isPrimary = /\bprimary\b/.test(s) || s.includes("transmission");
    const isNeutral = /\bneutral\b/.test(s);
    const isSecondary = /\bsecondary\b/.test(s) || s.includes("triplex") || (s.includes("service") && !looksComm && !s.includes("communication"));

    let kind = "other";

    if (category.toLowerCase() === "equipment") {
      if (isStreetLight) kind = "streetlight";
      else if (isDripLoop) kind = isStreetLightDripLoop ? "streetlight_drip_loop" : "power_drip_loop";
      else if (isRiser) kind = "riser";
      else kind = "equipment";
    } else if (category.toLowerCase() === "wire") {
      if (isPrimary) kind = "power_primary";
      else if (isNeutral) kind = "power_neutral";
      else if (isSecondary || isStreetLightFeed) kind = "power_secondary";
      else if (s.includes("power") || s.includes("electric") || s.includes("supply")) kind = "power_other";
      else if (looksComm) kind = "comm";
      else kind = "other";
    } else if (category.toLowerCase() === "guy") {
      kind = "guy";
    }

    // Comm drops: exempt from the pole-hole buffer rule.
    // Heuristic:
    // - requires a standalone "drop" token
    // - requires at least one *other* comm indicator (fiber/catv/cable/etc.) so that
    //   power "service drops" are not accidentally treated as comm drops.
    const isCommDrop = /\bdrop\b/.test(s) && ["catv", "fiber", "telephone", "tel", "coax", "cable", "communication", "comm", "adss"].some(w => s.includes(w));

    return {
      owner,
      kind,
      isAdss,
      isDownGuy,
      isRiser,
      isCommDrop,
      isDripLoop,
      isStreetLight,
      isStreetLightDripLoop,
      isStreetLightFeed,
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  QC engine
  // ────────────────────────────────────────────────────────────────────────────

  function runQcEngine(model, rules) {
    const poles = model.poles || [];
    const midspans = model.midspanPoints || [];

    const poleResults = {};
    const midResults = {};
    const issues = [];

    for (const pole of poles) {
      const res = evaluatePole(pole, rules);
      poleResults[String(pole.poleId)] = res;
      for (const iss of res.issues) issues.push(iss);
    }

    for (const ms of midspans) {
      const res = evaluateMidspan(ms, rules);
      midResults[String(ms.midspanId)] = res;
      for (const iss of res.issues) issues.push(iss);
    }

    // Span comm arrangement check:
    // Ensure comm ordering in midspans remains consistent with the ordering observed on poles.
    // This also flags endpoint poles when the ordering reverses across a span.
    const orderIssues = computeCommOrderIssues(model);
    for (const iss of orderIssues) {
      issues.push(iss);
      if (iss.entityType === "pole") {
        const r = poleResults[String(iss.entityId)];
        if (r && Array.isArray(r.issues)) r.issues.push(iss);
      } else if (iss.entityType === "midspan") {
        const r = midResults[String(iss.entityId)];
        if (r && Array.isArray(r.issues)) r.issues.push(iss);
      }
    }

    // Recompute statuses now that span-level issues may have been added.
    for (const r of Object.values(poleResults)) {
      r.status = deriveStatus(r.issues || []);
      r.hasCommOrderIssue = !!(r.issues || []).some((i) => String(i.ruleCode || "").startsWith("ORDER.COMM"));
    }
    for (const r of Object.values(midResults)) {
      r.status = deriveStatus(r.issues || []);
      r.hasCommOrderIssue = !!(r.issues || []).some((i) => String(i.ruleCode || "").startsWith("ORDER.COMM"));
    }

    const summary = summarizeResults(poleResults, midResults, issues);

    return {
      poles: poleResults,
      midspans: midResults,
      issues,
      summary,
    };
  }

  function summarizeResults(poleResults, midResults, issues) {
    const poles = { pass: 0, warn: 0, fail: 0, unknown: 0 };
    const mids = { pass: 0, warn: 0, fail: 0, unknown: 0 };

    for (const v of Object.values(poleResults || {})) poles[v.status || "unknown"] = (poles[v.status || "unknown"] || 0) + 1;
    for (const v of Object.values(midResults || {})) mids[v.status || "unknown"] = (mids[v.status || "unknown"] || 0) + 1;

    const issueCounts = { warn: 0, fail: 0 };
    for (const i of issues || []) {
      if (i.severity === "WARN") issueCounts.warn++;
      if (i.severity === "FAIL") issueCounts.fail++;
    }

    return { poles, midspans: mids, issues: issueCounts };
  }

  function deriveStatus(issues) {
    let hasFail = false;
    let hasWarn = false;
    for (const i of issues) {
      if (i.severity === "FAIL") hasFail = true;
      if (i.severity === "WARN") hasWarn = true;
    }
    if (hasFail) return "fail";
    if (hasWarn) return "warn";
    return "pass";
  }

  function issue(severity, entityType, entityId, entityName, ruleCode, message, context = {}) {
    return {
      severity, // "FAIL" | "WARN"
      entityType, // "pole" | "midspan"
      entityId,
      entityName,
      ruleCode,
      message,
      context,
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Span comm arrangement / ordering consistency
  // ────────────────────────────────────────────────────────────────────────────

  function normalizeOwnerKey(owner) {
    return String(owner || "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "")
      .trim();
  }

  function poleDisplayName(pole, fallbackId = "") {
    if (!pole) return fallbackId || "";
    return pole.displayName || pole.scid || pole.poleTag || String(pole.poleId || fallbackId || "");
  }

  function groupCommsOnPole(pole) {
    // Returns Map(ownerKey -> { owner, heightIn, ids: [] })
    const groups = new Map();
    const attsRaw = Array.isArray(pole && pole.attachments) ? pole.attachments : [];

    for (const a of attsRaw) {
      if (!a) continue;
      const h = (a.proposedIn != null ? a.proposedIn : a.existingIn);
      if (h == null) continue;

      const cls = classify(a);
      if (cls.kind !== "comm") continue;
      if (cls.isCommDrop) continue;

      const owner = String(cls.owner || "").trim();
      const key = normalizeOwnerKey(owner);
      if (!key) continue;

      const id = a.id != null ? String(a.id) : "";

      const prev = groups.get(key);
      if (!prev) {
        groups.set(key, { owner, heightIn: Number(h), ids: id ? [id] : [] });
        continue;
      }

      if (Number(h) > prev.heightIn) {
        prev.heightIn = Number(h);
        prev.ids = id ? [id] : [];
      } else if (Number(h) === prev.heightIn && id) {
        prev.ids.push(id);
      }
    }

    // Deduplicate IDs
    for (const g of groups.values()) g.ids = Array.from(new Set(g.ids || []));
    return groups;
  }

  function groupCommsOnMidspan(ms) {
    // Returns Map(ownerKey -> { owner, heightIn, ids: [] })
    const groups = new Map();
    const measuresRaw = Array.isArray(ms && ms.measures) ? ms.measures : [];

    for (const m0 of measuresRaw) {
      if (!m0) continue;
      const m = { ...m0, category: "Wire" };
      const h = (m.proposedIn != null ? m.proposedIn : m.existingIn);
      if (h == null) continue;

      const cls = classify(m);
      if (cls.kind !== "comm") continue;

      const owner = String(cls.owner || "").trim();
      const key = normalizeOwnerKey(owner);
      if (!key) continue;

      const id = m.id != null
        ? String(m.id)
        : `${String(m.traceId || "")}|${String(m.wireId || "")}|${String(h)}`;

      const prev = groups.get(key);
      if (!prev) {
        groups.set(key, { owner, heightIn: Number(h), ids: id ? [id] : [] });
        continue;
      }

      if (Number(h) > prev.heightIn) {
        prev.heightIn = Number(h);
        prev.ids = id ? [id] : [];
      } else if (Number(h) === prev.heightIn && id) {
        prev.ids.push(id);
      }
    }

    // Deduplicate IDs
    for (const g of groups.values()) g.ids = Array.from(new Set(g.ids || []));
    return groups;
  }

  function aboveOwnerKey(groups, k1, k2) {
    if (!groups) return "";
    const a = groups.get(k1);
    const b = groups.get(k2);
    if (!a || !b) return "";
    if (a.heightIn == null || b.heightIn == null) return "";
    if (Number(a.heightIn) === Number(b.heightIn)) return ""; // ambiguous
    return Number(a.heightIn) > Number(b.heightIn) ? k1 : k2;
  }

  function idsForPair(groups, k1, k2) {
    if (!groups) return [];
    const a = groups.get(k1);
    const b = groups.get(k2);
    const out = [];
    for (const id of (a && a.ids ? a.ids : [])) out.push(String(id));
    for (const id of (b && b.ids ? b.ids : [])) out.push(String(id));
    return Array.from(new Set(out.filter(Boolean)));
  }

  function computeCommOrderIssues(model) {
    const out = [];
    const dedupePole = new Set();
    const dedupeMid = new Set();

    const poles = Array.isArray(model && model.poles) ? model.poles : [];
    const mids = Array.isArray(model && model.midspanPoints) ? model.midspanPoints : [];

    const poleById = new Map();
    for (const p of poles) poleById.set(String(p.poleId), p);

    const midById = new Map();
    for (const m of mids) midById.set(String(m.midspanId), m);

    // Precompute comm grouping for each node.
    const poleComms = new Map();
    for (const p of poles) {
      poleComms.set(String(p.poleId), groupCommsOnPole(p));
    }

    const midComms = new Map();
    for (const m of mids) {
      midComms.set(String(m.midspanId), groupCommsOnMidspan(m));
    }

    // Build connection index: connectionId -> { aPoleId, bPoleId, midspanIds: [] }
    const connIndex = new Map();

    const spans = Array.isArray(model && model.spans) ? model.spans : [];
    for (const s of spans) {
      if (!s) continue;
      const cid = s.connectionId != null ? String(s.connectionId) : "";
      if (!cid) continue;
      const entry = connIndex.get(cid) || { connectionId: cid, aPoleId: "", bPoleId: "", midspanIds: [] };
      if (s.aIsPole && s.aNodeId != null) entry.aPoleId = String(s.aNodeId);
      if (s.bIsPole && s.bNodeId != null) entry.bPoleId = String(s.bNodeId);
      connIndex.set(cid, entry);
    }

    for (const m of mids) {
      const cid = m && m.connectionId != null ? String(m.connectionId) : "";
      if (!cid) continue;
      const entry = connIndex.get(cid) || { connectionId: cid, aPoleId: "", bPoleId: "", midspanIds: [] };
      if (!entry.aPoleId && m.aPoleId) entry.aPoleId = String(m.aPoleId);
      if (!entry.bPoleId && m.bPoleId) entry.bPoleId = String(m.bPoleId);
      entry.midspanIds.push(String(m.midspanId));
      connIndex.set(cid, entry);
    }

    const addPoleIssue = (poleId, poleName, key, iss) => {
      if (!poleId) return;
      if (dedupePole.has(key)) return;
      dedupePole.add(key);
      out.push(iss);
    };

    const addMidIssue = (midId, key, iss) => {
      if (!midId) return;
      if (dedupeMid.has(key)) return;
      dedupeMid.add(key);
      out.push(iss);
    };

    for (const entry of connIndex.values()) {
      const cid = String(entry.connectionId || "");
      const aId = entry.aPoleId ? String(entry.aPoleId) : "";
      const bId = entry.bPoleId ? String(entry.bPoleId) : "";

      const aPole = aId ? poleById.get(aId) : null;
      const bPole = bId ? poleById.get(bId) : null;
      const aName = poleDisplayName(aPole, aId);
      const bName = poleDisplayName(bPole, bId);

      const aGroups = aId ? poleComms.get(aId) : null;
      const bGroups = bId ? poleComms.get(bId) : null;

      // 1) Endpoint pole-to-pole ordering must be consistent for common comm owners.
      if (aGroups && bGroups && aGroups.size >= 2 && bGroups.size >= 2) {
        const commonKeys = Array.from(aGroups.keys()).filter((k) => bGroups.has(k));
        if (commonKeys.length >= 2) {
          for (const [k1, k2] of pairwise(commonKeys)) {
            const aboveA = aboveOwnerKey(aGroups, k1, k2);
            const aboveB = aboveOwnerKey(bGroups, k1, k2);
            if (!aboveA || !aboveB) continue;
            if (aboveA === aboveB) continue;

            const pairKey = [k1, k2].sort().join("|");

            const aAboveName = (aGroups.get(aboveA) && aGroups.get(aboveA).owner) ? aGroups.get(aboveA).owner : aboveA;
            const aBelowKey = (aboveA === k1 ? k2 : k1);
            const aBelowName = (aGroups.get(aBelowKey) && aGroups.get(aBelowKey).owner) ? aGroups.get(aBelowKey).owner : aBelowKey;

            const bAboveName = (bGroups.get(aboveB) && bGroups.get(aboveB).owner) ? bGroups.get(aboveB).owner : aboveB;
            const bBelowKey = (aboveB === k1 ? k2 : k1);
            const bBelowName = (bGroups.get(bBelowKey) && bGroups.get(bBelowKey).owner) ? bGroups.get(bBelowKey).owner : bBelowKey;

            const aIds = idsForPair(aGroups, k1, k2);
            const bIds = idsForPair(bGroups, k1, k2);

            addPoleIssue(
              aId,
              aName,
              `pole|${aId}|ORDER.COMM.ENDPOINTS|${cid}|${pairKey}`,
              issue(
                "FAIL",
                "pole",
                aId,
                aName,
                "ORDER.COMM.ENDPOINTS",
                `Comm order is reversed across span ${cid} between this pole and ${bName}: here ${aAboveName} is above ${aBelowName}, but at ${bName} ${bAboveName} is above ${bBelowName}.`,
                { connectionId: cid, otherPoleId: bId, otherPoleName: bName, owners: [aAboveName, aBelowName], attachmentIds: aIds }
              )
            );

            addPoleIssue(
              bId,
              bName,
              `pole|${bId}|ORDER.COMM.ENDPOINTS|${cid}|${pairKey}`,
              issue(
                "FAIL",
                "pole",
                bId,
                bName,
                "ORDER.COMM.ENDPOINTS",
                `Comm order is reversed across span ${cid} between this pole and ${aName}: here ${bAboveName} is above ${bBelowName}, but at ${aName} ${aAboveName} is above ${aBelowName}.`,
                { connectionId: cid, otherPoleId: aId, otherPoleName: aName, owners: [bAboveName, bBelowName], attachmentIds: bIds }
              )
            );
          }
        }
      }

      // 2) Each midspan photo must preserve the same comm owner ordering as the pole(s).
      const midspanIds = Array.isArray(entry.midspanIds) ? entry.midspanIds : [];
      for (const midId0 of midspanIds) {
        const midId = String(midId0 || "");
        if (!midId) continue;
        const ms = midById.get(midId);
        if (!ms) continue;

        const msGroups = midComms.get(midId);
        if (!msGroups || msGroups.size < 2) continue;

        const msKeys = Array.from(msGroups.keys());
        for (const [k1, k2] of pairwise(msKeys)) {
          const aboveMS = aboveOwnerKey(msGroups, k1, k2);
          if (!aboveMS) continue;

          const aboveA = aboveOwnerKey(aGroups, k1, k2);
          const aboveB = aboveOwnerKey(bGroups, k1, k2);

          // Skip if we have no reference order from either pole.
          if (!aboveA && !aboveB) continue;

          const pairKey = [k1, k2].sort().join("|");

          // If both poles define the pair and they disagree, flag the midspan as an endpoint-order conflict.
          if (aboveA && aboveB && aboveA !== aboveB) {
            const aAboveName = (aGroups.get(aboveA) && aGroups.get(aboveA).owner) ? aGroups.get(aboveA).owner : aboveA;
            const aBelowKey = (aboveA === k1 ? k2 : k1);
            const aBelowName = (aGroups.get(aBelowKey) && aGroups.get(aBelowKey).owner) ? aGroups.get(aBelowKey).owner : aBelowKey;

            const bAboveName = (bGroups.get(aboveB) && bGroups.get(aboveB).owner) ? bGroups.get(aboveB).owner : aboveB;
            const bBelowKey = (aboveB === k1 ? k2 : k1);
            const bBelowName = (bGroups.get(bBelowKey) && bGroups.get(bBelowKey).owner) ? bGroups.get(bBelowKey).owner : bBelowKey;

            addMidIssue(
              midId,
              `midspan|${midId}|ORDER.COMM.ENDPOINTS|${cid}|${pairKey}`,
              issue(
                "FAIL",
                "midspan",
                midId,
                `Midspan ${midId}`,
                "ORDER.COMM.ENDPOINTS",
                `Endpoint poles disagree on comm order across span ${cid} for ${aAboveName}/${aBelowName}. Pole ${aName}: ${aAboveName} above ${aBelowName}; Pole ${bName}: ${bAboveName} above ${bBelowName}.`,
                { connectionId: cid, aPoleId: aId, bPoleId: bId, aPoleName: aName, bPoleName: bName, owners: [aAboveName, aBelowName], measureIds: idsForPair(msGroups, k1, k2) }
              )
            );
            continue;
          }

          // Reference order: prefer both poles, otherwise whichever pole defines the pair.
          let refAbove = "";
          let refFrom = ""; // "both" | "a" | "b"
          if (aboveA && aboveB && aboveA === aboveB) {
            refAbove = aboveA;
            refFrom = "both";
          } else if (aboveA) {
            refAbove = aboveA;
            refFrom = "a";
          } else if (aboveB) {
            refAbove = aboveB;
            refFrom = "b";
          }
          if (!refAbove) continue;

          if (aboveMS === refAbove) continue;

          const refGroups = (refFrom === "b" ? bGroups : aGroups) || aGroups || bGroups;
          const refAboveName = (refGroups && refGroups.get(refAbove) && refGroups.get(refAbove).owner) ? refGroups.get(refAbove).owner : refAbove;
          const refBelowKey = (refAbove === k1 ? k2 : k1);
          const refBelowName = (refGroups && refGroups.get(refBelowKey) && refGroups.get(refBelowKey).owner) ? refGroups.get(refBelowKey).owner : refBelowKey;

          const msAboveName = (msGroups.get(aboveMS) && msGroups.get(aboveMS).owner) ? msGroups.get(aboveMS).owner : aboveMS;
          const msBelowKey = (aboveMS === k1 ? k2 : k1);
          const msBelowName = (msGroups.get(msBelowKey) && msGroups.get(msBelowKey).owner) ? msGroups.get(msBelowKey).owner : msBelowKey;

          const severity = (refFrom === "both") ? "FAIL" : "WARN";
          const refNote = (refFrom === "both")
            ? "(both poles agree)"
            : (refFrom === "a")
              ? `(based on pole ${aName} only)`
              : `(based on pole ${bName} only)`;

          addMidIssue(
            midId,
            `midspan|${midId}|ORDER.COMM.MIDSPAN|${cid}|${pairKey}`,
            issue(
              severity,
              "midspan",
              midId,
              `Midspan ${midId}`,
              "ORDER.COMM.MIDSPAN",
              `Comm order mismatch in span ${cid}: midspan shows ${msAboveName} above ${msBelowName}, but pole order requires ${refAboveName} above ${refBelowName} ${refNote}.`,
              { connectionId: cid, aPoleId: aId, bPoleId: bId, aPoleName: aName, bPoleName: bName, owners: [refAboveName, refBelowName], measureIds: idsForPair(msGroups, k1, k2) }
            )
          );

          // Also flag the involved pole(s) so the map can quickly highlight the affected span endpoints.
          if ((refFrom === "both" || refFrom === "a") && aId && aGroups && aGroups.has(k1) && aGroups.has(k2)) {
            addPoleIssue(
              aId,
              aName,
              `pole|${aId}|ORDER.COMM.MIDSPAN|${cid}|${pairKey}`,
              issue(
                severity,
                "pole",
                aId,
                aName,
                "ORDER.COMM.MIDSPAN",
                `Comm order mismatch on span ${cid} (midspan ${midId}): expected ${refAboveName} above ${refBelowName}, but midspan shows ${msAboveName} above ${msBelowName}.`,
                { connectionId: cid, midspanId: midId, otherPoleId: bId, otherPoleName: bName, owners: [refAboveName, refBelowName], attachmentIds: idsForPair(aGroups, k1, k2) }
              )
            );
          }

          if ((refFrom === "both" || refFrom === "b") && bId && bGroups && bGroups.has(k1) && bGroups.has(k2)) {
            addPoleIssue(
              bId,
              bName,
              `pole|${bId}|ORDER.COMM.MIDSPAN|${cid}|${pairKey}`,
              issue(
                severity,
                "pole",
                bId,
                bName,
                "ORDER.COMM.MIDSPAN",
                `Comm order mismatch on span ${cid} (midspan ${midId}): expected ${refAboveName} above ${refBelowName}, but midspan shows ${msAboveName} above ${msBelowName}.`,
                { connectionId: cid, midspanId: midId, otherPoleId: aId, otherPoleName: aName, owners: [refAboveName, refBelowName], attachmentIds: idsForPair(bGroups, k1, k2) }
              )
            );
          }
        }
      }
    }

    return out;
  }

  function evaluatePole(pole, rules) {
    const issues = [];
    const attsRaw = Array.isArray(pole.attachments) ? pole.attachments : [];

    const atts = attsRaw
      .map(a => ({ ...a, _cls: classify(a) }))
      .filter(a => a && (a.proposedIn != null || a.existingIn != null));

    const poleName = pole.displayName || pole.scid || pole.poleTag || pole.poleId;

    // Warn if missing pole identifiers (Pole Spec OR Pole Tag OR SCID)
    if (rules.pole.warnMissingPoleIdentifiers) {
      const hasSpec = !!String(pole.poleSpec || pole.proposedPoleSpec || pole.poleHeightClass || "").trim();
      const hasTag = !!String(pole.poleTag || "").trim();
      const hasScid = !!String(pole.scid || "").trim();
      if (!hasSpec && !hasTag && !hasScid) {
        issues.push(issue("WARN", "pole", String(pole.poleId), poleName, "POLE.MISSING_ID",
          "Pole is missing Pole Spec, Pole Tag, and SCID (at least one identifier is recommended)."));
      }
    }

    // Use proposed heights for primary QC checks; fallback to existing if proposed missing.
    const getH = (a) => (a.proposedIn != null ? a.proposedIn : a.existingIn);

    // Helper: capture which attachment rows are implicated by a given rule violation.
    // This enables row-level highlighting in the Details panel.
    const attachmentIdsOf = (...items) => {
      const flat = items.flat().filter(Boolean);
      const out = [];
      for (const it of flat) {
        const id = it && it.id != null ? String(it.id) : "";
        if (id) out.push(id);
      }
      return Array.from(new Set(out));
    };

    // COMM wires (exclude drip loops; drip loops are equipment)
    const comm = atts.filter(a => a._cls.kind === "comm" && a.proposedIn != null);
    const commUnique = uniqueByKey(comm, (a) => `${a._cls.owner}|${a.proposedIn}`);

    // LOWEST COMM attachment
    if (comm.length) {
      const minComm = Math.min(...comm.map(a => a.proposedIn));
      if (minComm < rules.pole.minLowestCommAttachIn) {
        const offenders = comm.filter(a => a.proposedIn === minComm);
        issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.MIN_COMM",
          `Lowest comm attachment is ${fmtFtIn(minComm)}, below the minimum ${fmtFtIn(rules.pole.minLowestCommAttachIn)}.` ,
          { minComm, minRequired: rules.pole.minLowestCommAttachIn, attachmentIds: attachmentIdsOf(offenders) }));
      }
    }

    // Comms separation on pole: 12" between different comms; same comm: 0 or >= 4"
    if (commUnique.length > 1) {
      const pairs = pairwise(commUnique);
      for (const [a, b] of pairs) {
        const dh = Math.abs(a.proposedIn - b.proposedIn);
        const sameOwner = (a._cls.owner && b._cls.owner && a._cls.owner === b._cls.owner);

        if (sameOwner) {
          if (dh !== 0 && dh < rules.pole.commSepSameIn) {
            issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.COMM_SEP_SAME",
              `${a._cls.owner}: comm separation ${dh}" is below ${rules.pole.commSepSameIn}" (same company; exact same height is allowed).`,
              { owner: a._cls.owner, h1: a.proposedIn, h2: b.proposedIn, dh, attachmentIds: attachmentIdsOf(a, b) }));
          }
        } else {
          if (dh < rules.pole.commSepDiffIn) {
            issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.COMM_SEP_DIFF",
              `Comm separation ${dh}" between "${a._cls.owner || "Unknown"}" and "${b._cls.owner || "Unknown"}" is below ${rules.pole.commSepDiffIn}".`,
              { ownerA: a._cls.owner, ownerB: b._cls.owner, h1: a.proposedIn, h2: b.proposedIn, dh, attachmentIds: attachmentIdsOf(a, b) }));
          }
        }
      }
    }

    // ADSS must be the highest comm
    if (rules.pole.enforceAdssHighest && comm.length > 1) {
      const commSorted = comm.slice().sort((x, y) => (y.proposedIn - x.proposedIn));
      const top = commSorted[0];
      const hasAdss = comm.some(a => a._cls.isAdss);
      if (hasAdss && !top._cls.isAdss) {
        const highestAdss = commSorted.find(a => a._cls.isAdss);
        issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.ADSS_TOP",
          `Highest comm (${fmtFtIn(top.proposedIn)}) is not ADSS, but at least one ADSS comm exists on the pole.`,
          { topHeight: top.proposedIn, attachmentIds: attachmentIdsOf(top, highestAdss) }));
      }
    }

    // Comms-to-power separation (pole)
    // - Default: 40" from lowest power (including drip loops), excluding street light drip loops.
    // - ADSS comms: may be closer (default 12") — applies to ADSS facilities only.
    // - Risers:
    //   • risers are allowed to move (handled in equipment movement rule)
    //   • comm-owned risers are checked against the 40" comm-to-power separation
    //   • power-owned risers are ignored for comm-to-power separation
    //
    // NOTE: We evaluate each comm facility independently so that non-ADSS comms remain subject to
    // the 40" rule even when an ADSS facility is permitted closer to power.
    {
      const powerWires = atts.filter(a => a._cls.kind.startsWith("power_") && a.proposedIn != null);
      const dripLoops = atts.filter(a => a._cls.isDripLoop && a.proposedIn != null && !a._cls.isStreetLightDripLoop);
      const powerCandidates = [...powerWires, ...dripLoops];

      if (powerCandidates.length) {
        const powerSorted = powerCandidates.slice().sort((a, b) => getH(a) - getH(b));
        const lowPowerAtt = powerSorted[0];
        const lowPower = getH(lowPowerAtt);

        const normalizeOwnerKey = (v) => String(v || "").toLowerCase().replace(/[^a-z0-9]/g, "");
        const looksLikePowerCompany = (v) => {
          const s = String(v || "").toLowerCase();
          return s.includes("electric") || s.includes("power") || s.includes("energy") || s.includes("utility") || s.includes("utilit") || s.includes("coop") || s.includes("co-op");
        };
        const likelyPowerOwnerKey = (() => {
          const counts = new Map();
          for (const p of powerWires) {
            const k = normalizeOwnerKey(p._cls.owner);
            if (!k) continue;
            counts.set(k, (counts.get(k) || 0) + 1);
          }
          if (counts.size) {
            let bestK = "";
            let bestN = -1;
            for (const [k, n] of counts.entries()) {
              if (n > bestN) { bestN = n; bestK = k; }
            }
            return bestK;
          }
          const pk = normalizeOwnerKey(pole.poleOwner);
          return pk || "";
        })();

        // Identify risers; include them in comm-to-power only when they appear to be comm-owned.
        const risers = atts.filter(a => a._cls.isRiser && a.proposedIn != null);
        const commOwnedRisers = [];
        const unknownOwnerRisers = [];

        for (const r of risers) {
          const ownerKey = normalizeOwnerKey(r._cls.owner);
          if (!ownerKey) {
            unknownOwnerRisers.push(r);
            continue;
          }
          if (likelyPowerOwnerKey && ownerKey === likelyPowerOwnerKey) {
            continue; // power-owned riser
          }
          if (!likelyPowerOwnerKey && looksLikePowerCompany(r._cls.owner)) {
            continue; // best-effort power-owned riser
          }
          commOwnedRisers.push(r);
        }

        const commFacilities = [...comm, ...commOwnedRisers, ...unknownOwnerRisers];

        for (const c of commFacilities) {
          const h = c.proposedIn;
          if (h == null) continue;

          const sep = lowPower - h;
          const isAdss = !!(c._cls && c._cls.isAdss);
          const isRiser = !!(c._cls && c._cls.isRiser);
          const minReq = isAdss ? rules.pole.adssCommToPowerSepIn : rules.pole.commToPowerSepIn;

          if (sep < minReq) {
            const ownerKey = normalizeOwnerKey(c._cls.owner);
            const unknownOwner = isRiser && !ownerKey;
            const severity = unknownOwner ? "WARN" : "FAIL";
            const facilityLabel = isRiser ? "Riser" : (isAdss ? "ADSS comm" : "Comm");
            const ownerLabel = c._cls.owner ? ` (${c._cls.owner})` : "";
            const suffix = isAdss ? " (ADSS allowance)" : "";
            const note = unknownOwner ? " Owner unknown; if this is power-owned, this check may not apply." : "";

            issues.push(issue(severity, "pole", String(pole.poleId), poleName, "POLE.COMM_TO_POWER",
              `${facilityLabel}${ownerLabel} at ${fmtFtIn(h)} is ${sep}" below lowest power at ${fmtFtIn(lowPower)}, below ${minReq}"${suffix}.${note}`,
              { highComm: h, lowPower, sep, minReq, isAdssTop: isAdss, attachmentIds: attachmentIdsOf(c, lowPowerAtt) }));
          }
        }
      }
    }

    // Comms-to-streetlight separation (12")
    if (comm.length) {
      const streetEq = atts.filter(a => a._cls.isStreetLight && a.proposedIn != null);
      if (streetEq.length) {
        for (const c of comm) {
          for (const s of streetEq) {
            const dh = Math.abs(c.proposedIn - s.proposedIn);
            if (dh < rules.pole.commToStreetLightSepIn) {
              issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.COMM_TO_STREETLIGHT",
                `Comm at ${fmtFtIn(c.proposedIn)} is ${dh}" from street light equipment at ${fmtFtIn(s.proposedIn)} (min ${rules.pole.commToStreetLightSepIn}").`,
                { commH: c.proposedIn, streetH: s.proposedIn, dh, attachmentIds: attachmentIdsOf(c, s) }));
            }
          }
        }
      }
    }

    // Moved-hole / keep-out buffer rule (skip on pole replacements)
    //
    // Interpretation implemented here (based on your description):
    // - This rule applies ONLY when the pole is NOT being replaced.
    // - Drip loops are excluded (not actual pole attachments).
    // - Any moved or new attachment's PROPOSED height must not be within the buffer distance of:
    //   (a) any moved-from hole height (existing height of a moved attachment), OR
    //   (b) any other attachment height, OR
    //   (c) another moved/new attachment height.
    // - Attaching at the exact same height as an existing hole is allowed (reuse hole).
    if (!pole.poleReplacement && rules.pole.movedHoleBufferIn > 0) {
      const buffer = rules.pole.movedHoleBufferIn;

      // Comm drops are immune to the moved-hole/keep-out buffer rule.
      // They should not be checked *as* moved/new attachments, and they should not
      // block other attachments that move around them.
      const nonDrip = atts.filter(a => !a._cls.isDripLoop && a.proposedIn != null && !a._cls.isCommDrop);

      const stationary = nonDrip.filter(a => !(a.isMoved || a.isNew));
      const moved = nonDrip.filter(a => a.isMoved && a.existingIn != null);
      const movedOrNew = nonDrip.filter(a => (a.isMoved || a.isNew));

      const movedFromHeights = moved.map(a => a.existingIn);

      // (a) vs moved-from holes
      if (movedFromHeights.length && movedOrNew.length) {
        for (const a of movedOrNew) {
          for (const holeH of movedFromHeights) {
            const dh = Math.abs(a.proposedIn - holeH);
            if (dh !== 0 && dh < buffer) {
              const holeSources = moved.filter(m => m.existingIn === holeH);
              issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.HOLE_BUFFER",
                `Proposed attachment at ${fmtFtIn(a.proposedIn)} is ${dh}" from a moved-from hole at ${fmtFtIn(holeH)} (min ${buffer}").`,
                { proposed: a.proposedIn, holeH, dh, buffer, attachmentIds: attachmentIdsOf(a, holeSources) }));
            }
          }
        }
      }

      // (b) vs stationary attachments
      if (stationary.length && movedOrNew.length) {
        for (const a of movedOrNew) {
          for (const s of stationary) {
            const dh = Math.abs(a.proposedIn - s.proposedIn);
            if (dh !== 0 && dh < buffer) {
              issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.HOLE_BUFFER",
                `Proposed attachment at ${fmtFtIn(a.proposedIn)} is ${dh}" from an existing attachment at ${fmtFtIn(s.proposedIn)} (min ${buffer}").`,
                { proposed: a.proposedIn, other: s.proposedIn, dh, buffer, attachmentIds: attachmentIdsOf(a, s) }));
            }
          }
        }
      }

      // (c) moved/new vs moved/new
      if (movedOrNew.length > 1) {
        for (const [a, b] of pairwise(movedOrNew)) {
          const dh = Math.abs(a.proposedIn - b.proposedIn);
          if (dh !== 0 && dh < buffer) {
            issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.HOLE_BUFFER",
              `Two proposed moved/new attachments are ${dh}" apart at ${fmtFtIn(a.proposedIn)} and ${fmtFtIn(b.proposedIn)} (min ${buffer}").`,
              { h1: a.proposedIn, h2: b.proposedIn, dh, buffer, attachmentIds: attachmentIdsOf(a, b) }));
          }
        }
      }
    }

    // Restrict equipment movement
    // - Street lights are allowed to move
    // - Down guys are immune
    // - Risers are allowed to move (immune)
    // - If a pole is replaced AND upgraded to a taller pole, do not warn on equipment moves
    //   (power ordering is still enforced separately)
    if (rules.pole.enforceEquipmentMove) {
      const replacementIsTaller = !!pole.poleReplacementIsTaller;
      const movedEquip = atts.filter(a =>
        String(a.category || "").toLowerCase() === "equipment" &&
        a.isMoved &&
        !a._cls.isDripLoop &&
        !a._cls.isDownGuy &&
        !a._cls.isRiser
      );

      for (const e of movedEquip) {
        if (e._cls.isStreetLight) continue; // allowed
        if (replacementIsTaller) continue;  // allowed (no warning)

        if (pole.poleReplacement) {
          issues.push(issue("WARN", "pole", String(pole.poleId), poleName, "POLE.EQUIP_MOVE",
            `Equipment move detected (${e.label || "equipment"}). Pole is marked as replacement; review per utility standard.`,
            { label: e.label || "", attachmentIds: attachmentIdsOf(e) }));
        } else {
          issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.EQUIP_MOVE",
            `Equipment move detected (${e.label || "equipment"}). Equipment generally cannot be moved unless explicitly allowed.`,
            { label: e.label || "", attachmentIds: attachmentIdsOf(e) }));
        }
      }
    }

    // Power wire movement ordering (neutral vs secondary)
    if (rules.pole.enforcePowerOrder) {
      const neutrals = atts.filter(a => a._cls.kind === "power_neutral" && a.existingIn != null && a.proposedIn != null && !a.isNew);
      const secondaries = atts.filter(a => a._cls.kind === "power_secondary" && a.existingIn != null && a.proposedIn != null && !a.isNew);

      if (neutrals.length && secondaries.length) {
        const neutralMaxExisting = Math.max(...neutrals.map(a => a.existingIn));
        const secondaryMaxExisting = Math.max(...secondaries.map(a => a.existingIn));
        const neutralMaxProposed = Math.max(...neutrals.map(a => a.proposedIn));
        const secondaryMaxProposed = Math.max(...secondaries.map(a => a.proposedIn));
        const powerAttIds = attachmentIdsOf(neutrals, secondaries);

        if (secondaryMaxExisting > neutralMaxExisting) {
          // secondary was above neutral; do not move neutral above secondary
          if (neutralMaxProposed > secondaryMaxProposed) {
            issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.POWER_ORDER",
              `Power ordering violation: secondary was above neutral (existing), but proposed neutral is above proposed secondary.`,
              { neutralMaxExisting, secondaryMaxExisting, neutralMaxProposed, secondaryMaxProposed, attachmentIds: powerAttIds }));
          }
        } else if (neutralMaxExisting > secondaryMaxExisting) {
          // neutral was above secondary; do not move secondary above neutral
          if (secondaryMaxProposed > neutralMaxProposed) {
            issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.POWER_ORDER",
              `Power ordering violation: neutral was above secondary (existing), but proposed secondary is above proposed neutral.`,
              { neutralMaxExisting, secondaryMaxExisting, neutralMaxProposed, secondaryMaxProposed, attachmentIds: powerAttIds }));
          }
        }
      }
    }

    return {
      status: deriveStatus(issues),
      issues,
    };
  }

  function requiredMidspanMinComm(rules, rowType) {
    const t = String(rowType || "default");
    if (t === "pedestrian") return rules.midspan.minCommPedestrianIn;
    if (t === "highway") return rules.midspan.minCommHighwayIn;
    if (t === "farm") return rules.midspan.minCommFarmIn;
    if (t === "rail") return rules.midspan.minCommRailIn;
    return rules.midspan.minCommDefaultIn;
  }

  function evaluateMidspan(ms, rules) {
    const issues = [];
    const msName = `Midspan ${ms.midspanId}`;

    const measuresRaw = Array.isArray(ms.measures) ? ms.measures : [];
    // NOTE:
    // Midspan moves can be interpolated, yielding floating-point inches. QC comparisons are
    // performed at whole-inch resolution to match field measurement and rule definitions.
    const roundIn = (v) => {
      const n = Number(v);
      return Number.isFinite(n) ? Math.round(n) : null;
    };

    const measures = measuresRaw
      .map(m => {
        const ex = roundIn(m && m.existingIn);
        const pr = roundIn(m && m.proposedIn);
        return {
          ...m,
          existingIn: (ex != null ? ex : (m && m.existingIn != null ? m.existingIn : null)),
          proposedIn: (pr != null ? pr : (m && m.proposedIn != null ? m.proposedIn : null)),
          category: "Wire",
          _cls: classify(m),
        };
      })
      .filter(m => m && m.proposedIn != null);

    // For row-level highlighting in the Details panel.
    const measureIdOf = (m) => {
      if (!m) return "";
      if (m.id != null) return String(m.id);
      const trace = m.traceId != null ? String(m.traceId) : "";
      const wire = m.wireId != null ? String(m.wireId) : "";
      const h = m.proposedIn != null ? String(m.proposedIn) : "";
      return `${trace}|${wire}|${h}`;
    };
    const measureIdsOf = (...items) => {
      const flat = items.flat().filter(Boolean);
      const out = [];
      for (const it of flat) {
        const id = measureIdOf(it);
        if (id) out.push(id);
      }
      return Array.from(new Set(out));
    };

    // Warn on missing ROW type
    const rowTypeRaw = String(ms.rowTypeRaw || "").trim();
    if (rules.midspan.warnMissingRowType && !rowTypeRaw) {
      issues.push(issue("WARN", "midspan", String(ms.midspanId), msName, "MIDSPAN.MISSING_ROW",
        "Midspan ROW type is missing; defaulting to “Default” (15' 6\")."));
    }


    const rowTypeNorm = ms.rowType || "default";
    const rowLabel = String(rowTypeRaw || rowTypeNorm || "default").trim() || (rowTypeNorm || "default");
    const reqComm = requiredMidspanMinComm(rules, rowTypeNorm);

    // Lowest comm height rule
    const comm = measures.filter(m => m._cls.kind === "comm");
    const power = measures.filter(m => m._cls.kind.startsWith("power_") || m._cls.kind === "power_other");

    if (comm.length) {
      const minComm = Math.min(...comm.map(m => m.proposedIn));
      if (minComm < reqComm) {
        const offenders = comm.filter(m => m.proposedIn === minComm);
        issues.push(issue("FAIL", "midspan", String(ms.midspanId), msName, "MIDSPAN.MIN_COMM",
          `Lowest comm at ${fmtFtIn(minComm)} is below the minimum ${fmtFtIn(reqComm)} for ROW type "${rowLabel}".`,
          { minComm, req: reqComm, rowType: rowTypeNorm, rowTypeRaw: rowTypeRaw || "", measureIds: measureIdsOf(offenders) }));
      }
    }

    // Power-only minimum clearance rule:
    // If a midspan photo contains ONLY power wires (secondary/neutral/primary/etc.) and NO comms,
    // the minimum proposed height must be 1' 0" above the comm minimum for that ROW type.
    if (!comm.length && power.length) {
      const minPower = Math.min(...power.map(m => m.proposedIn));
      const reqPower = reqComm + 12;
      if (minPower < reqPower) {
        const offenders = power.filter(m => m.proposedIn === minPower);
        issues.push(issue("FAIL", "midspan", String(ms.midspanId), msName, "MIDSPAN.MIN_POWER_ONLY",
          `Lowest power at ${fmtFtIn(minPower)} is below the minimum ${fmtFtIn(reqPower)} (1' 0\" above comm minimum) for ROW type "${rowLabel}".`,
          { minPower, reqPower, reqComm, rowType: rowTypeNorm, rowTypeRaw: rowTypeRaw || "", measureIds: measureIdsOf(offenders) }));
      }
    }

    // Midspan comm separation (default 4", optionally stricter for installing company)
    if (comm.length > 1) {
      const installCo = String(rules.midspan.installingCompany || "").trim();
      const baseMin = rules.midspan.commSepIn;
      const installMin = Math.max(baseMin, rules.midspan.installingCompanyCommSepIn || baseMin);

      const commUnique = uniqueByKey(comm, (a) => `${a.owner}|${a.proposedIn}`);
      for (const [a, b] of pairwise(commUnique)) {
        const dh = Math.abs(a.proposedIn - b.proposedIn);
        const minReq = (installCo && (a.owner === installCo || b.owner === installCo)) ? installMin : baseMin;
        if (dh !== 0 && dh < minReq) {
          issues.push(issue("FAIL", "midspan", String(ms.midspanId), msName, "MIDSPAN.COMM_SEP",
            `Comm separation ${dh}" between "${a.owner || "Unknown"}" and "${b.owner || "Unknown"}" is below ${minReq}".`,
            { dh, minReq, ownerA: a.owner, ownerB: b.owner, h1: a.proposedIn, h2: b.proposedIn, measureIds: measureIdsOf(a, b) }));
        }
      }
    }

    // Midspan comm-to-power separation
    // - Default: 30" from lowest power to highest comm
    // - ADSS: may be closer (default 12")
    if (comm.length) {
      const power = measures.filter(m => m._cls.kind.startsWith("power_") || m._cls.kind === "power_other");
      if (power.length) {
        const topComm = comm.slice().sort((a, b) => b.proposedIn - a.proposedIn)[0];
        const lowPowerAtt = power.slice().sort((a, b) => a.proposedIn - b.proposedIn)[0];
        const highComm = topComm.proposedIn;
        const lowPower = lowPowerAtt.proposedIn;
        const sep = lowPower - highComm;

        const isAdss = !!(topComm._cls && topComm._cls.isAdss);
        const minReq = isAdss ? (rules.midspan.adssCommToPowerSepIn ?? rules.midspan.commToPowerSepIn) : rules.midspan.commToPowerSepIn;
        const suffix = isAdss ? " (ADSS allowance)" : "";

        if (sep < minReq) {
          issues.push(issue("FAIL", "midspan", String(ms.midspanId), msName, "MIDSPAN.COMM_TO_POWER",
            `Comm-to-power separation is ${sep}" (${fmtFtIn(highComm)} to ${fmtFtIn(lowPower)}), below ${minReq}"${suffix}.`,
            { highComm, lowPower, sep, minReq, measureIds: measureIdsOf(topComm, lowPowerAtt) }));
        }
      }
    }

    // ADSS highest comm
    if (rules.midspan.enforceAdssHighest && comm.length > 1) {
      const hasAdss = comm.some(c => c._cls.isAdss);
      if (hasAdss) {
        const sorted = comm.slice().sort((x, y) => y.proposedIn - x.proposedIn);
        const top = sorted[0];
        if (!top._cls.isAdss) {
          const highestAdss = sorted.find(c => c._cls.isAdss);
          issues.push(issue("FAIL", "midspan", String(ms.midspanId), msName, "MIDSPAN.ADSS_TOP",
            `Highest comm (${fmtFtIn(top.proposedIn)}) is not ADSS, but at least one ADSS comm exists in this midspan photo.`,
            { topHeight: top.proposedIn, measureIds: measureIdsOf(top, highestAdss) }));
        }
      }
    }

    return {
      status: deriveStatus(issues),
      issues,
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Popups + Details
  // ────────────────────────────────────────────────────────────────────────────

  function buildPolePopupHtml(poleId) {
    const p = (model.poles || []).find(x => String(x.poleId) === String(poleId));
    const res = qcResults && qcResults.poles ? qcResults.poles[String(poleId)] : null;
    if (!p) return "<div class='muted'>Pole not found.</div>";

    const name = escapeHtml(p.displayName || p.scid || p.poleTag || p.poleId);
    const st = res ? res.status : "unknown";
    const badge = `<span class="badge badge--${st}">${st.toUpperCase()}</span>`;

    const issues = res ? res.issues : [];
    const top = issues.slice(0, 5).map(i => `<li><strong>${escapeHtml(i.ruleCode)}</strong>: ${escapeHtml(i.message)}</li>`).join("");

    return `
      <div class="popup">
        <div class="popup-title">${name} ${badge}</div>
        ${issues.length ? `<ul class="popup-list">${top}</ul>` : `<div class="muted">No issues.</div>`}
        ${issues.length > 5 ? `<div class="muted">+${issues.length - 5} more…</div>` : ""}
      </div>
    `;
  }

  function buildMidspanPopupHtml(midId) {
    const ms = (model.midspanPoints || []).find(x => String(x.midspanId) === String(midId));
    const res = qcResults && qcResults.midspans ? qcResults.midspans[String(midId)] : null;
    if (!ms) return "<div class='muted'>Midspan not found.</div>";

    const st = res ? res.status : "unknown";
    const badge = `<span class="badge badge--${st}">${st.toUpperCase()}</span>`;
    const row = escapeHtml(ms.rowTypeRaw || ms.rowType || "default");

    const issues = res ? res.issues : [];
    const top = issues.slice(0, 5).map(i => `<li><strong>${escapeHtml(i.ruleCode)}</strong>: ${escapeHtml(i.message)}</li>`).join("");

    return `
      <div class="popup">
        <div class="popup-title">Midspan (${row}) ${badge}</div>
        ${issues.length ? `<ul class="popup-list">${top}</ul>` : `<div class="muted">No issues.</div>`}
        ${issues.length > 5 ? `<div class="muted">+${issues.length - 5} more…</div>` : ""}
      </div>
    `;
  }

  function clearDetails() {
    els.details.innerHTML = `<div class="muted">Load a job and click a pole or midspan marker to view violations and measured heights.</div>`;
    closeDetailsPanel();
  }

  function openDetailsPanel(anchorLatLng) {
    if (!els.detailsPanel) return;

    // Default slide-in fallback.
    els.detailsPanel.style.setProperty("--enter-x", "18px");
    els.detailsPanel.style.setProperty("--enter-y", "0px");
    els.detailsPanel.style.setProperty("--origin-x", "80%");
    els.detailsPanel.style.setProperty("--origin-y", "12%");

    // Animate the panel from the clicked marker location for a more "map-native" feel.
    // This is computed in map container pixels, so it remains stable across zoom levels.
    if (anchorLatLng && map) {
      const mapEl = document.getElementById("map");
      if (mapEl) {
        try {
          const pt = map.latLngToContainerPoint(anchorLatLng);
          const panel = els.detailsPanel;
          const cs = getComputedStyle(panel);
          const topPx = parseFloat(cs.top) || 12;
          const rightPx = parseFloat(cs.right) || 12;

          // Panel dimensions are stable even when closed (opacity 0).
          const panelW = panel.offsetWidth || 520;
          const panelH = panel.offsetHeight || 520;

          // Compute the panel's final (open) box relative to the map container.
          const finalLeft = mapEl.clientWidth - rightPx - panelW;
          const finalTop = topPx;
          const finalCx = finalLeft + panelW / 2;
          const finalCy = finalTop + panelH / 2;

          const dx = pt.x - finalCx;
          const dy = pt.y - finalCy;

          panel.style.setProperty("--enter-x", `${dx}px`);
          panel.style.setProperty("--enter-y", `${dy}px`);

          const ox = Math.max(0, Math.min(100, ((pt.x - finalLeft) / panelW) * 100));
          const oy = Math.max(0, Math.min(100, ((pt.y - finalTop) / panelH) * 100));
          panel.style.setProperty("--origin-x", `${ox}%`);
          panel.style.setProperty("--origin-y", `${oy}%`);
        } catch (_) {
          // Non-fatal: fall back to slide-in.
        }
      }
    }

    els.detailsPanel.classList.add("is-open");
    els.detailsPanel.setAttribute("aria-hidden", "false");
  }

  function closeDetailsPanel() {
    if (!els.detailsPanel) return;
    els.detailsPanel.classList.remove("is-open");
    els.detailsPanel.setAttribute("aria-hidden", "true");
  }

  function selectEntity(sel) {
    if (!sel || !model) return;

    // Render first so the in-map panel can size itself to the content before the
    // open animation computes its final box.
    if (sel.type === "pole") {
      const pole = (model.poles || []).find(p => String(p.poleId) === String(sel.id));
      if (!pole) return;
      renderPoleDetails(pole);
    } else if (sel.type === "midspan") {
      const ms = (model.midspanPoints || []).find(m => String(m.midspanId) === String(sel.id));
      if (!ms) return;
      renderMidspanDetails(ms);
    }

    openDetailsPanel(sel.latlng || null);
  }

  function renderPoleDetails(pole) {
    const poleId = String(pole.poleId);
    const res = qcResults && qcResults.poles ? qcResults.poles[poleId] : null;
    const st = res ? res.status : "unknown";
    const name = escapeHtml(pole.displayName || pole.scid || pole.poleTag || poleId);

    const issues = res ? res.issues : [];

    // Attachment row highlighting: any attachment implicated in a FAIL is marked red.
    // (WARN rows are marked amber.)
    const failAttIds = new Set();
    const warnAttIds = new Set();
    for (const iss of issues) {
      const ids = iss && iss.context && Array.isArray(iss.context.attachmentIds) ? iss.context.attachmentIds : [];
      for (const id of ids) {
        const sid = String(id || "");
        if (!sid) continue;
        if (iss.severity === "FAIL") failAttIds.add(sid);
        if (iss.severity === "WARN") warnAttIds.add(sid);
      }
    }
    const issuesHtml = issues.length
      ? `<ul class="issue-list">${issues.map(i => `<li class="issue issue--${i.severity === "FAIL" ? "fail" : "warn"}">
          <div class="issue-code">${escapeHtml(i.ruleCode)}</div>
          <div class="issue-msg">${escapeHtml(i.message)}</div>
        </li>`).join("")}</ul>`
      : `<div class="muted">No issues.</div>`;

    const atts = (pole.attachments || []).slice()
      .filter(a => a && (a.proposedIn != null || a.existingIn != null))
      .sort((a, b) => ((b.proposedIn ?? b.existingIn ?? 0) - (a.proposedIn ?? a.existingIn ?? 0)));

    const rows = atts.map(a => {
      const h = (a.proposedIn != null ? a.proposedIn : a.existingIn);
      const cls = classify(a);
      const aId = String(a.id || "");
      const rowCls = failAttIds.has(aId) ? "detail-row--fail" : (warnAttIds.has(aId) ? "detail-row--warn" : "");
      return `
        <tr class="${rowCls}">
          <td>${escapeHtml(fmtFtIn(h))}</td>
          <td>${escapeHtml(String(a.category || ""))}</td>
          <td>${escapeHtml(cls.owner || "")}</td>
          <td>${escapeHtml(String(a.label || ""))}</td>
          <td class="muted">${escapeHtml(String(a.existingHeight || ""))}</td>
          <td class="muted">${escapeHtml(String(a.proposedHeight || ""))}</td>
        </tr>
      `;
    }).join("");

    els.details.innerHTML = `
      <div class="details-header">
        <div>
          <div class="details-title">${name}</div>
          <div class="muted">Pole ID: ${escapeHtml(poleId)}${pole.poleReplacement ? " • Pole replacement" : ""}</div>
        </div>
        <div><span class="badge badge--${st}">${st.toUpperCase()}</span></div>
      </div>

      <h3 class="details-subtitle">Issues</h3>
      ${issuesHtml}

      <h3 class="details-subtitle">Attachments (sorted by height)</h3>
      <div class="table-wrap">
        <table class="table table--small">
          <thead>
            <tr>
              <th>Height</th>
              <th>Category</th>
              <th>Owner</th>
              <th>Label</th>
              <th>Existing</th>
              <th>Proposed</th>
            </tr>
          </thead>
          <tbody>${rows || `<tr><td colspan="6" class="muted">No attachment rows.</td></tr>`}</tbody>
        </table>
      </div>
    `;
  }

  function renderMidspanDetails(ms) {
    const midId = String(ms.midspanId);
    const res = qcResults && qcResults.midspans ? qcResults.midspans[midId] : null;
    const st = res ? res.status : "unknown";

    const issues = res ? res.issues : [];

    // Row-level highlighting: any measured wire implicated in a FAIL is marked red
    // (WARN rows are amber). This mirrors pole attachment highlighting.
    const failMeasIds = new Set();
    const warnMeasIds = new Set();
    for (const iss of issues) {
      const ids = iss && iss.context && Array.isArray(iss.context.measureIds) ? iss.context.measureIds : [];
      for (const id of ids) {
        const sid = String(id || "");
        if (!sid) continue;
        if (iss.severity === "FAIL") failMeasIds.add(sid);
        if (iss.severity === "WARN") warnMeasIds.add(sid);
      }
    }
    const issuesHtml = issues.length
      ? `<ul class="issue-list">${issues.map(i => `<li class="issue issue--${i.severity === "FAIL" ? "fail" : "warn"}">
          <div class="issue-code">${escapeHtml(i.ruleCode)}</div>
          <div class="issue-msg">${escapeHtml(i.message)}</div>
        </li>`).join("")}</ul>`
      : `<div class="muted">No issues.</div>`;

    const measures = (ms.measures || []).slice()
      .filter(m => m && m.proposedIn != null)
      .sort((a, b) => (b.proposedIn - a.proposedIn));

    const measureIdOf = (m) => {
      if (!m) return "";
      if (m.id != null) return String(m.id);
      const trace = m.traceId != null ? String(m.traceId) : "";
      const wire = m.wireId != null ? String(m.wireId) : "";
      const h = m.proposedIn != null ? String(m.proposedIn) : "";
      return `${trace}|${wire}|${h}`;
    };

    const rows = measures.map(m => {
      const cls = classify(m);
      const mId = measureIdOf(m);
      const rowCls = failMeasIds.has(mId) ? "detail-row--fail" : (warnMeasIds.has(mId) ? "detail-row--warn" : "");
      return `
        <tr class="${rowCls}">
          <td>${escapeHtml(fmtFtIn(m.proposedIn))}</td>
          <td>${escapeHtml(cls.owner || "")}</td>
          <td>${escapeHtml(String(m.label || ""))}</td>
          <td class="muted">${escapeHtml(String(m.existingHeight || ""))}</td>
          <td class="muted">${escapeHtml(String(m.proposedHeight || ""))}</td>
        </tr>
      `;
    }).join("");

    els.details.innerHTML = `
      <div class="details-header">
        <div>
          <div class="details-title">Midspan</div>
          <div class="muted">ROW: ${escapeHtml(ms.rowTypeRaw || ms.rowType || "default")} • ID: ${escapeHtml(midId)}</div>
        </div>
        <div><span class="badge badge--${st}">${st.toUpperCase()}</span></div>
      </div>

      <h3 class="details-subtitle">Issues</h3>
      ${issuesHtml}

      <h3 class="details-subtitle">Measured Wires (sorted by height)</h3>
      <div class="table-wrap">
        <table class="table table--small">
          <thead>
            <tr>
              <th>Proposed Height</th>
              <th>Owner</th>
              <th>Label</th>
              <th>Existing</th>
              <th>Proposed</th>
            </tr>
          </thead>
          <tbody>${rows || `<tr><td colspan="5" class="muted">No midspan measures.</td></tr>`}</tbody>
        </table>
      </div>
    `;
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Issues table
  // ────────────────────────────────────────────────────────────────────────────

  function clearIssuesTable() {
    els.issuesTbody.innerHTML = `<tr><td colspan="5" class="muted">No issues yet.</td></tr>`;
  }

  function renderIssuesTable() {
    if (!qcResults) return clearIssuesTable();

    const wantFail = els.issuesFail.checked;
    const wantWarn = els.issuesWarn.checked;
    const q = normalizeStr(els.issuesSearch.value || "").trim();

    const rows = [];
    for (const iss of qcResults.issues || []) {
      if (iss.severity === "FAIL" && !wantFail) continue;
      if (iss.severity === "WARN" && !wantWarn) continue;

      const hay = normalizeStr(`${iss.severity} ${iss.entityType} ${iss.entityName} ${iss.ruleCode} ${iss.message}`);
      if (q && !hay.includes(q)) continue;

      rows.push(iss);
    }

    if (!rows.length) {
      els.issuesTbody.innerHTML = `<tr><td colspan="5" class="muted">No matching issues.</td></tr>`;
      return;
    }

    els.issuesTbody.innerHTML = rows.map((iss) => {
      const sev = iss.severity === "FAIL" ? "FAIL" : "WARN";
      const sevClass = iss.severity === "FAIL" ? "fail" : "warn";
      const name = escapeHtml(iss.entityType === "pole" ? (iss.entityName || iss.entityId) : (iss.entityName || "Midspan"));
      const msg = escapeHtml(iss.message);

      return `
        <tr class="issue-row issue-row--${sevClass}" data-entity-type="${escapeHtml(iss.entityType)}" data-entity-id="${escapeHtml(iss.entityId)}">
          <td><span class="pill pill--${sevClass}">${sev}</span></td>
          <td>${escapeHtml(iss.entityType)}</td>
          <td>${name}</td>
          <td><code>${escapeHtml(iss.ruleCode)}</code></td>
          <td>${msg}</td>
        </tr>
      `;
    }).join("");

    // Click to focus
    Array.from(els.issuesTbody.querySelectorAll("tr.issue-row")).forEach((tr) => {
      tr.addEventListener("click", () => {
        const t = tr.getAttribute("data-entity-type");
        const id = tr.getAttribute("data-entity-id");
        focusEntity(t, id);
      });
    });

  }

  function focusEntity(type, id) {
    if (!type || !id) return;

    // switch to map tab
    const btn = document.querySelector('.tab-btn[data-tab="map"]');
    if (btn) btn.click();

    if (type === "pole") {
      const marker = poleMarkers.get(String(id));
      if (marker) {
        map.setView(marker.getLatLng(), Math.max(map.getZoom(), 18));
        marker.openPopup();
        selectEntity({ type: "pole", id });
      }
    } else if (type === "midspan") {
      const marker = midspanMarkers.get(String(id));
      if (marker) {
        map.setView(marker.getLatLng(), Math.max(map.getZoom(), 18));
        marker.openPopup();
        selectEntity({ type: "midspan", id });
      }
    }
  }

  function exportIssuesCsv() {
    if (!qcResults || !qcResults.issues || !qcResults.issues.length) return;

    const cols = ["severity", "entityType", "entityId", "entityName", "ruleCode", "message"];
    const lines = [cols.join(",")];

    for (const i of qcResults.issues) {
      const row = [
        i.severity,
        i.entityType,
        i.entityId,
        i.entityName || "",
        i.ruleCode,
        i.message,
      ].map(csvEscape).join(",");
      lines.push(row);
    }

    const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `katapult_qc_issues_${safeFilename(model.jobName || "job")}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();

    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Rules UI
  // ────────────────────────────────────────────────────────────────────────────

  function applyRulesToUi() {
    // Pole
    els.rulePoleMinAttach.value = fmtFtIn(rules.pole.minLowestCommAttachIn);
    els.rulePoleCommSepDiff.value = String(rules.pole.commSepDiffIn);
    els.rulePoleCommSepSame.value = String(rules.pole.commSepSameIn);
    els.rulePoleCommToPower.value = String(rules.pole.commToPowerSepIn);
    els.rulePoleAdssCommToPower.value = String(rules.pole.adssCommToPowerSepIn);
    els.rulePoleCommToStreet.value = String(rules.pole.commToStreetLightSepIn);
    els.rulePoleHoleBuffer.value = String(rules.pole.movedHoleBufferIn);

    els.rulePoleEnforceAdss.checked = !!rules.pole.enforceAdssHighest;
    els.rulePoleEnforceEquipMove.checked = !!rules.pole.enforceEquipmentMove;
    els.rulePoleEnforcePowerOrder.checked = !!rules.pole.enforcePowerOrder;
    els.rulePoleWarnMissingIds.checked = !!rules.pole.warnMissingPoleIdentifiers;

    // Midspan
    els.ruleMidMinDefault.value = fmtFtIn(rules.midspan.minCommDefaultIn);
    els.ruleMidMinPed.value = fmtFtIn(rules.midspan.minCommPedestrianIn);
    els.ruleMidMinHwy.value = fmtFtIn(rules.midspan.minCommHighwayIn);
    els.ruleMidMinFarm.value = fmtFtIn(rules.midspan.minCommFarmIn);

    els.ruleMidCommSep.value = String(rules.midspan.commSepIn);
    els.ruleMidCommToPower.value = String(rules.midspan.commToPowerSepIn);
    if (els.ruleMidAdssCommToPower) els.ruleMidAdssCommToPower.value = String(rules.midspan.adssCommToPowerSepIn);

    els.ruleMidCommSepInstall.value = String(rules.midspan.installingCompanyCommSepIn);
    els.ruleMidEnforceAdss.checked = !!rules.midspan.enforceAdssHighest;
    els.ruleMidWarnMissingRow.checked = !!rules.midspan.warnMissingRowType;

    if (els.ruleInstallingCompany) {
      els.ruleInstallingCompany.value = rules.midspan.installingCompany || "";
    }
  }

  function readRulesFromUi() {
    const r = structuredClone(DEFAULT_RULES);

    // Pole
    r.pole.minLowestCommAttachIn = parseFtIn(els.rulePoleMinAttach.value, DEFAULT_RULES.pole.minLowestCommAttachIn);
    r.pole.commSepDiffIn = parseIntSafe(els.rulePoleCommSepDiff.value, DEFAULT_RULES.pole.commSepDiffIn);
    r.pole.commSepSameIn = parseIntSafe(els.rulePoleCommSepSame.value, DEFAULT_RULES.pole.commSepSameIn);
    r.pole.commToPowerSepIn = parseIntSafe(els.rulePoleCommToPower.value, DEFAULT_RULES.pole.commToPowerSepIn);
    r.pole.adssCommToPowerSepIn = parseIntSafe(els.rulePoleAdssCommToPower.value, DEFAULT_RULES.pole.adssCommToPowerSepIn);
    r.pole.commToStreetLightSepIn = parseIntSafe(els.rulePoleCommToStreet.value, DEFAULT_RULES.pole.commToStreetLightSepIn);
    r.pole.movedHoleBufferIn = parseIntSafe(els.rulePoleHoleBuffer.value, DEFAULT_RULES.pole.movedHoleBufferIn);

    r.pole.enforceAdssHighest = !!els.rulePoleEnforceAdss.checked;
    r.pole.enforceEquipmentMove = !!els.rulePoleEnforceEquipMove.checked;
    r.pole.enforcePowerOrder = !!els.rulePoleEnforcePowerOrder.checked;
    r.pole.warnMissingPoleIdentifiers = !!els.rulePoleWarnMissingIds.checked;

    // Midspan
    r.midspan.minCommDefaultIn = parseFtIn(els.ruleMidMinDefault.value, DEFAULT_RULES.midspan.minCommDefaultIn);
    r.midspan.minCommPedestrianIn = parseFtIn(els.ruleMidMinPed.value, DEFAULT_RULES.midspan.minCommPedestrianIn);
    r.midspan.minCommHighwayIn = parseFtIn(els.ruleMidMinHwy.value, DEFAULT_RULES.midspan.minCommHighwayIn);
    r.midspan.minCommFarmIn = parseFtIn(els.ruleMidMinFarm.value, DEFAULT_RULES.midspan.minCommFarmIn);

    r.midspan.commSepIn = parseIntSafe(els.ruleMidCommSep.value, DEFAULT_RULES.midspan.commSepIn);
    r.midspan.commToPowerSepIn = parseIntSafe(els.ruleMidCommToPower.value, DEFAULT_RULES.midspan.commToPowerSepIn);
    if (els.ruleMidAdssCommToPower) r.midspan.adssCommToPowerSepIn = parseIntSafe(els.ruleMidAdssCommToPower.value, DEFAULT_RULES.midspan.adssCommToPowerSepIn);

    r.midspan.installingCompany = String(els.ruleInstallingCompany.value || "").trim();
    r.midspan.installingCompanyCommSepIn = parseIntSafe(els.ruleMidCommSepInstall.value, DEFAULT_RULES.midspan.installingCompanyCommSepIn);

    r.midspan.enforceAdssHighest = !!els.ruleMidEnforceAdss.checked;
    r.midspan.warnMissingRowType = !!els.ruleMidWarnMissingRow.checked;

    return r;
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Search
  // ────────────────────────────────────────────────────────────────────────────

  function searchAndFocusPole(query) {
    const q = normalizeStr(query).trim();
    if (!q) return;

    const list = (previewPoles && previewPoles.length) ? previewPoles : (model && model.poles ? model.poles : []);
    if (!list.length) return;

    let found = null;
    for (const p of list) {
      const hay = normalizeStr(`${p.displayName || ""} ${p.scid || ""} ${p.poleTag || ""}`);
      if (hay.includes(q)) {
        found = p;
        break;
      }
    }
    if (!found) {
      logLine(`Search: no match for "${query}".`);
      return;
    }

    const marker = poleMarkers.get(String(found.poleId));
    if (marker) {
      map.setView(marker.getLatLng(), Math.max(map.getZoom(), 18));
      marker.openPopup();
      selectEntity({ type: "pole", id: found.poleId });
    }
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  UI helpers
  // ────────────────────────────────────────────────────────────────────────────

  function updateUiEnabled(hasFile) {
    els.btnPreview.disabled = !hasFile;
    els.btnRunQc.disabled = !hasFile;
    els.btnReset.disabled = !hasFile;
    els.btnResetRules.disabled = !hasFile;
    els.btnZoomAll.disabled = !hasFile;

    els.btnExportIssues.disabled = true;
  }

  function setProgress(pct, label) {
    const p = Math.max(0, Math.min(100, Number(pct) || 0));
    els.progressBar.style.width = `${p}%`;
    els.progressLabel.textContent = label || "";
  }

  function logLine(s) {
    const line = String(s || "");
    els.logBox.textContent += (els.logBox.textContent ? "\n" : "") + line;
    els.logBox.scrollTop = els.logBox.scrollHeight;
  }

  function setJobName(name) {
    els.jobName.textContent = String(name || "—");
  }

  function setSummaryCounts(summary) {
    if (!summary) {
      els.summaryMidspans.textContent = "—";
      els.summaryIssues.textContent = "—";
      return;
    }
    const poleTotals = summary.poles || {};
    const midTotals = summary.midspans || {};
    const issueTotals = summary.issues || {};

    els.summaryPoles.textContent = String((poleTotals.pass || 0) + (poleTotals.warn || 0) + (poleTotals.fail || 0) + (poleTotals.unknown || 0));
    els.summaryMidspans.textContent = String((midTotals.pass || 0) + (midTotals.warn || 0) + (midTotals.fail || 0) + (midTotals.unknown || 0));
    els.summaryIssues.textContent = String((issueTotals.fail || 0) + (issueTotals.warn || 0));
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Utilities
  // ────────────────────────────────────────────────────────────────────────────

  function parseIntSafe(v, fallback) {
    const n = parseInt(String(v || "").trim(), 10);
    return Number.isFinite(n) ? n : fallback;
  }

  function parseFtIn(text, fallbackIn) {
    const s = String(text || "").trim();
    if (!s) return fallbackIn;

    // Accept forms like:
    //  - 16' 0"
    //  - 16'0"
    //  - 16
    //  - 192" (inches)
    const m = s.match(/^\s*(\d+)\s*(?:'\s*(\d+)\s*(?:\"|in)?\s*)?$/i);
    if (m) {
      const ft = parseInt(m[1], 10);
      const inch = m[2] != null ? parseInt(m[2], 10) : 0;
      if (Number.isFinite(ft) && Number.isFinite(inch)) return ft * 12 + inch;
    }

    const m2 = s.match(/^\s*(\d+)\s*(?:\"|in)\s*$/i);
    if (m2) {
      const inch = parseInt(m2[1], 10);
      if (Number.isFinite(inch)) return inch;
    }

    return fallbackIn;
  }

  function fmtFtIn(inches) {
    const v = Number(inches);
    if (!Number.isFinite(v)) return "";
    const ft = Math.floor(v / 12);
    const inch = Math.round(v - ft * 12);
    return `${ft}' ${inch}"`;
  }

  function uniqueByKey(arr, keyFn) {
    const out = [];
    const seen = new Set();
    for (const a of arr || []) {
      const k = keyFn(a);
      if (seen.has(k)) continue;
      seen.add(k);
      out.push(a);
    }
    return out;
  }

  function pairwise(arr) {
    const out = [];
    for (let i = 0; i < arr.length; i++) {
      for (let j = i + 1; j < arr.length; j++) out.push([arr[i], arr[j]]);
    }
    return out;
  }

  function escapeHtml(s) {
    return String(s || "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll("\"", "&quot;")
      .replaceAll("'", "&#039;");
  }

  function csvEscape(v) {
    const s = String(v ?? "");
    const needs = /[",\n\r]/.test(s);
    const t = s.replaceAll('"', '""');
    return needs ? `"${t}"` : t;
  }

  function safeFilename(name) {
    return String(name || "job").replace(/[^a-z0-9\-_]+/gi, "_").slice(0, 60);
  }

  function loadRules() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return structuredClone(DEFAULT_RULES);
      const parsed = JSON.parse(raw);
      // Shallow merge for safety
      return {
        pole: { ...DEFAULT_RULES.pole, ...(parsed.pole || {}) },
        midspan: { ...DEFAULT_RULES.midspan, ...(parsed.midspan || {}) },
      };
    } catch (_) {
      return structuredClone(DEFAULT_RULES);
    }
  }

  function saveRules(rules) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(rules));
    } catch (_) {}
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  Start
  // ────────────────────────────────────────────────────────────────────────────

  document.addEventListener("DOMContentLoaded", init);
})();
