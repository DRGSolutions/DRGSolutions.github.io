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
      enforceNeutralSecondaryBelowTransformer: false,
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
    btnSaveRulesFile: $("btnSaveRulesFile"),
    btnLoadRulesFile: $("btnLoadRulesFile"),
    rulesFileInput: $("rulesFileInput"),
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
    // View mode (2D/3D)
    view2d: $("view2d"),
    view3d: $("view3d"),
    map3d: $("map3d"),
    basemap3d: $("basemap3d"),
    filterPass: $("filterPass"),
    filterWarn: $("filterWarn"),
    filterFail: $("filterFail"),
	    searchScid: $("searchScid"),
	    searchPoleTag: $("searchPoleTag"),

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
    rulePoleNeutralSecondaryBelowTransformer: $("rulePoleNeutralSecondaryBelowTransformer"),
    rulePoleWarnMissingIds: $("rulePoleWarnMissingIds"),

    ruleMidMinDefault: $("ruleMidMinDefault"),
    ruleMidMinPed: $("ruleMidMinPed"),
    ruleMidMinHwy: $("ruleMidMinHwy"),
    ruleMidMinFarm: $("ruleMidMinFarm"),
    ruleMidMinRail: $("ruleMidMinRail"),
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
  let baseLayers = null;
  let currentBasemap = "Dark";

  let poleCluster = null;
  let midspanCluster = null;
  let spanLayer = null;
  let scidLabelLayer = null;

  const poleMarkers = new Map();      // poleId -> marker
  const midspanMarkers = new Map();   // midspanId -> marker
  const spanPolylines = [];           // L.Polyline[]
  const scidLabelMarkers = new Map(); // poleId -> label marker

  // 3D view state (visualization only; QC logic remains unchanged)
  let viewMode = "2d";               // "2d" | "3d"
  let threeState = null;             // lazily initialized
  let threeDirty = false;            // rebuild needed after QC/model changes

  // ────────────────────────────────────────────────────────────────────────────
  //  Init
  // ────────────────────────────────────────────────────────────────────────────

  function init() {
    initTabs();
    initMap();
    initEvents();

    // Maximize initial map viewport area.
    maximizeMapHeight();
    if (viewMode === "2d" && map) map.invalidateSize();
    if (viewMode === "3d") resizeThreeRenderer();

    // Keep map sized to available viewport space.
    window.addEventListener("resize", () => {
      maximizeMapHeight();
      if (viewMode === "2d" && map) map.invalidateSize();
      if (viewMode === "3d") resizeThreeRenderer();
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

        // Leaflet/3D canvas need a resize pass when the map container becomes visible
        if (tab === "map") {
          setTimeout(() => {
            maximizeMapHeight();
            if (viewMode === "2d" && map) map.invalidateSize();
            if (viewMode === "3d") resizeThreeRenderer();
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

    // Track selected basemap for 3D ground rendering (visualization only).
    baseLayers = { Dark: baseDark, Light: baseLight, Imagery: baseImagery };
    currentBasemap = "Dark";
    if (els.basemap3d) els.basemap3d.value = currentBasemap;

    map.on("baselayerchange", (e) => {
      try {
        if (!e || !e.layer) return;
        if (e.layer === baseDark) currentBasemap = "Dark";
        else if (e.layer === baseLight) currentBasemap = "Light";
        else if (e.layer === baseImagery) currentBasemap = "Imagery";
        if (els.basemap3d) els.basemap3d.value = currentBasemap;

        // Mark 3D scene dirty so the ground texture matches the selected basemap.
        threeDirty = true;
        if (viewMode === "3d" && threeState) {
          try { rebuildThreeScene(); } catch (_) {}
        }
      } catch (_) {}
    });

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

  function setBasemap(name) {
    // Applies to 2D Leaflet and 3D ground (visualization only).
    if (!map || !baseLayers) return;
    const key = (name && baseLayers[name]) ? String(name) : "Dark";

    try {
      for (const k of Object.keys(baseLayers)) {
        const lyr = baseLayers[k];
        if (lyr && map.hasLayer(lyr)) map.removeLayer(lyr);
      }
    } catch (_) {}

    try { baseLayers[key].addTo(map); } catch (_) {}
    currentBasemap = key;
    if (els.basemap3d) els.basemap3d.value = currentBasemap;

    // Mark 3D scene dirty so the ground updates to the selected basemap.
    threeDirty = true;
    if (viewMode === "3d" && threeState) {
      try { rebuildThreeScene(); } catch (_) {}
    }
  }

  function maximizeMapHeight() {
    // Expand the map to use as much of the viewport as possible, without requiring a full-screen mode.
    // This makes the QC map feel more like a dedicated workspace.
    const primaryId = (viewMode === "3d") ? "map3d" : "map";
    const el = document.getElementById(primaryId);
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
    // Apply to both 2D and 3D containers so switching views does not cause a jump.
    for (const id of ["map", "map3d"]) {
      const e2 = document.getElementById(id);
      if (e2) e2.style.height = `${h}px`;
    }
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

    // 3D visibility mirrors the same controls/filters.
    refreshThreeVisibility();
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  3D View (visualization only)
  // ────────────────────────────────────────────────────────────────────────────

  // IMPORTANT:
  // Some Three.js CDN distributions (including certain unpkg/jsDelivr routes)
  // may not provide the legacy non-module OrbitControls at the expected path.
  // When that happens, THREE.OrbitControls is undefined and the 3D view would
  // appear "blank". To make the 3D mode robust, we use OrbitControls if present,
  // otherwise fall back to a small built-in orbit control implementation.

  function makeOrbitControls(camera, domElement) {
    // Stable orbit controls (visualization only; does not touch QC logic/results):
    // - Left-drag: pan/translate
    // - Right-drag: orbit/rotate around target (keeps the scene in view)
    // - Wheel: zoom
    return new OrbitControlsLite(camera, domElement);
  }

  function OrbitControlsLite(camera, domElement) {
    // Minimal OrbitControls-style camera controller (rotate / pan / zoom).
    // This is purely a visualization aid and does not touch QC logic.
    const scope = this;
    this.object = camera;
    this.domElement = domElement;

    this.target = new THREE.Vector3();

    // Prevent drag navigation from firing selection clicks.
    this.suppressClick = false;

    this.enableDamping = true;
    this.dampingFactor = 0.06;
    this.screenSpacePanning = true;

    // Allow closer inspections (visualization only).
    // Keeping this small also reduces the feeling of a hard stop when zooming in.
    this.minDistance = 0.001;
    this.maxDistance = 250000;
    this.maxPolarAngle = Math.PI;

    // Speed tuning (visualization only): keep consistent behavior regardless of zoom.
    // (No distance-aware scaling; the same input always produces the same response.)
    this.rotateSpeed = 0.00048;
    // Zoom (dolly) uses a constant world-units step per wheel "notch" (ft).
    // This avoids the "more zoomed = smaller zoom step" feel in perspective views.
    this.zoomStepFt = 120;

    this.panSpeed = 0.10;
	    // NOTE: We intentionally do not implement any automatic "focus near pole" behavior.
	    // Pivot snapping / cursor focus caused visible camera jumps in real-world use.


    // Keyboard translation (6DOF-style navigation).
    // - WASD / Arrow keys: strafe + forward/back (horizontal)
    // - Q/E (and F/R as alternates): down/up (world vertical)
    // - Hold Shift to boost; hold Ctrl to slow.
    this.enableKeys = true;
    this.keyMoveSpeed = 130; // ft per second

    const STATE = { NONE: 0, ROTATE: 1, PAN: 2 };
    let state = STATE.NONE;
    let activePointerId = null;
    let lastX = 0;
    let lastY = 0;
    let movedPx = 0;
    let hoverX = 0;
    let hoverY = 0;
    let zoomFromWheel = false;


    const spherical = new THREE.Spherical();
    const sphericalDelta = new THREE.Spherical(1, 0, 0); // radius unused
    sphericalDelta.theta = 0;
    sphericalDelta.phi = 0;
    const panOffset = new THREE.Vector3();
    let zoomDelta = 0;

    const EPS = 1e-6;

    const v3tmp = new THREE.Vector3();
    const v3tmp2 = new THREE.Vector3();



    // 6DOF-style keyboard navigation (visualization only).
    const keys = Object.create(null);
    const vKeyMove = new THREE.Vector3();
    const vKeyWorld = new THREE.Vector3();
    const WORLD_UP = new THREE.Vector3(0, 1, 0);

    // Cursor-to-world helpers: used for "zoom toward cursor" so the camera does not
    // become center-map biased and can continue to zoom in near outlying poles.
    const vNdc = new THREE.Vector3();
    const vRayDir = new THREE.Vector3();
    const vCursorHit = new THREE.Vector3();
    const vPlaneN = new THREE.Vector3();

    function isTypingTarget(evt) {
      // Only treat true text-entry controls as typing targets.
      // This avoids "WASDQE doesn't work" reports when focus is on a checkbox/radio.
      const el = evt && evt.target;
      if (!el) return false;
      const tag = String(el.tagName || "").toUpperCase();
      if (tag === "TEXTAREA" || !!el.isContentEditable) return true;
      if (tag === "SELECT") return true;
      if (tag === "INPUT") {
        const type = String(el.type || "text").toLowerCase();
        return type === "text" || type === "search" || type === "email" || type === "number" ||
          type === "password" || type === "tel" || type === "url";
      }
      return false;
    }

    function is3dActive() {
      try { return !!(document.body && document.body.classList && document.body.classList.contains("is-3d")); } catch (_) {}
      return true;
    }

    function shouldCaptureKey(code) {
      return code === "KeyW" || code === "KeyA" || code === "KeyS" || code === "KeyD" ||
        code === "KeyQ" || code === "KeyE" || code === "KeyR" || code === "KeyF" ||
        code === "ArrowUp" || code === "ArrowDown" || code === "ArrowLeft" || code === "ArrowRight" ||
        code === "PageUp" || code === "PageDown" ||
        code === "Equal" || code === "Minus" || code === "NumpadAdd" || code === "NumpadSubtract" ||
        code === "ShiftLeft" || code === "ShiftRight" || code === "ControlLeft" || code === "ControlRight";
    }

    function onKeyDown(evt) {
      if (!scope.enableKeys) return;
      if (!is3dActive()) return;
      if (isTypingTarget(evt)) return;

      const code = String(evt.code || "");
      if (!code) return;

      keys[code] = true;

      // Keyboard zoom for users without mouse wheels.
      if (code === "Equal" || code === "NumpadAdd") {
        zoomDelta -= 120;
      } else if (code === "Minus" || code === "NumpadSubtract") {
        zoomDelta += 120;
      }

      if (shouldCaptureKey(code)) {
        evt.preventDefault();
      }
    }

    function onKeyUp(evt) {
      if (!scope.enableKeys) return;
      if (!is3dActive()) return;
      if (isTypingTarget(evt)) return;

      const code = String(evt.code || "");
      if (!code) return;

      keys[code] = false;

      if (shouldCaptureKey(code)) {
        evt.preventDefault();
      }
    }

    document.addEventListener("keydown", onKeyDown, true);
    document.addEventListener("keyup", onKeyUp, true);

    function updateSphericalFromCamera() {
      const offset = v3tmp.copy(camera.position).sub(scope.target);
      spherical.setFromVector3(offset);
      if (!Number.isFinite(spherical.radius) || spherical.radius <= 0) spherical.radius = 800;
    }

    updateSphericalFromCamera();
    function getZoomScaleFromWheel(deltaY) {
      // Normalize wheel delta to keep zoom feel consistent across devices.
      const dy = Number(deltaY || 0);
      if (!isFinite(dy) || dy === 0) return 0;
	      // Some devices (especially trackpads) can emit very large deltas in bursts;
	      // clamp to avoid sudden, disorienting jumps.
	      return Math.max(-360, Math.min(360, dy));
    }

    function getCursorHitOnTargetPlane(out) {
      // Intersect the mouse cursor ray with the plane that passes through
      // the current orbit target and is perpendicular to the camera view direction.
      // This keeps zoom behavior stable and prevents "center-map" focal drift.
      try {
        const rect = scope.domElement.getBoundingClientRect();
        if (!rect || rect.width <= 0 || rect.height <= 0) return false;

        const ndcX = ((hoverX - rect.left) / rect.width) * 2 - 1;
        const ndcY = -((hoverY - rect.top) / rect.height) * 2 + 1;

        vNdc.set(ndcX, ndcY, 0.5).unproject(camera);
        vRayDir.copy(vNdc).sub(camera.position).normalize();

        camera.getWorldDirection(vPlaneN);
        const denom = vRayDir.dot(vPlaneN);
        if (!Number.isFinite(denom) || Math.abs(denom) < 1e-6) return false;

        const t = vCursorHit.copy(scope.target).sub(camera.position).dot(vPlaneN) / denom;
        if (!Number.isFinite(t) || t <= 0) return false;

        out.copy(camera.position).addScaledVector(vRayDir, t);
        return true;
      } catch (_) {
        return false;
      }
    }

    function pan(dx, dy) {
      const element = scope.domElement;
      const w = element.clientWidth || 1;
      const h = element.clientHeight || 1;

	      // Ensure camera matrices are current so pan vectors don't "snap" or jitter
	      // due to stale transforms during rapid interactions.
	      try { camera.updateMatrixWorld(true); } catch (_) {}

      if (camera.isPerspectiveCamera) {
        // Constant world-units-per-pixel pan so navigation feels the same
        // regardless of zoom level (visualization only).
        const panX = dx * scope.panSpeed;
        const panY = dy * scope.panSpeed;

        // camera matrix columns: 0 = right, 1 = up
        // Keep mouse panning in the X/Z plane so zoom does not become increasingly
        // restrictive due to target Y drift during long navigation sessions.
	        const right = v3tmp.setFromMatrixColumn(camera.matrixWorld, 0);
        right.y = 0;
        if (right.lengthSq() < EPS) right.set(1, 0, 0);
        right.normalize().multiplyScalar(-panX);

	        const up = v3tmp2.setFromMatrixColumn(camera.matrixWorld, 1);
        up.y = 0;
        // When the camera is perfectly level, "up" projects to ~0 in X/Z.
        // Fall back to forward projected onto the ground plane.
        if (up.lengthSq() < EPS) {
          try { camera.getWorldDirection(up); } catch (_) { up.set(0, 0, 1); }
          up.y = 0;
        }
        if (up.lengthSq() < EPS) up.set(0, 0, 1);
        up.normalize().multiplyScalar(panY);

        panOffset.add(right).add(up);
      } else if (camera.isOrthographicCamera) {
        const panX = (dx * (camera.right - camera.left)) / w * scope.panSpeed;
        const panY = (dy * (camera.top - camera.bottom)) / h * scope.panSpeed;
	        const right = v3tmp.setFromMatrixColumn(camera.matrixWorld, 0);
        right.y = 0;
        if (right.lengthSq() < EPS) right.set(1, 0, 0);
        right.normalize().multiplyScalar(-panX);

	        const up = v3tmp2.setFromMatrixColumn(camera.matrixWorld, 1);
        up.y = 0;
        if (up.lengthSq() < EPS) {
          try { camera.getWorldDirection(up); } catch (_) { up.set(0, 0, 1); }
          up.y = 0;
        }
        if (up.lengthSq() < EPS) up.set(0, 0, 1);
        up.normalize().multiplyScalar(panY);

        panOffset.add(right).add(up);
      }
    }


function onPointerDown(e) {
      if (activePointerId != null) return;
      activePointerId = e.pointerId;
      try { scope.domElement.setPointerCapture(activePointerId); } catch (_) {}

      lastX = e.clientX;
      lastY = e.clientY;
      hoverX = e.clientX;
      hoverY = e.clientY;

      movedPx = 0;
      scope.suppressClick = false;

      if (e.button === 0 || e.button === 1) {
        state = STATE.PAN;
      } else {
        state = STATE.ROTATE;
      }
      e.preventDefault();
    }

    function onPointerMove(e) {
      hoverX = e.clientX;
      hoverY = e.clientY;
      if (activePointerId == null || e.pointerId !== activePointerId) return;
      const dx = e.clientX - lastX;
      const dy = e.clientY - lastY;
      lastX = e.clientX;
      lastY = e.clientY;

      movedPx += Math.abs(dx) + Math.abs(dy);

      if (state === STATE.ROTATE) {
        // Keep rotate behavior consistent regardless of zoom level.
        const sp = scope.rotateSpeed;
        sphericalDelta.theta -= dx * sp;
        sphericalDelta.phi -= dy * sp;
      } else if (state === STATE.PAN) {
        pan(dx, dy);
      }
      e.preventDefault();
    }

    function onPointerUp(e) {
      if (activePointerId == null || e.pointerId !== activePointerId) return;
      try { scope.domElement.releasePointerCapture(activePointerId); } catch (_) {}
      if (movedPx > 3) scope.suppressClick = true;
      activePointerId = null;
      state = STATE.NONE;

      e.preventDefault();
    }

    function onWheel(e) {
      // Track cursor position so zoom can bias toward the user's current focus.
      hoverX = e.clientX;
      hoverY = e.clientY;
      zoomFromWheel = true;
      zoomDelta += getZoomScaleFromWheel(e.deltaY);
      e.preventDefault();
    }

    function onContextMenu(e) {
      // Allow right-drag panning without the browser context menu.
      e.preventDefault();
    }

    domElement.addEventListener("pointerdown", onPointerDown, { passive: false });
    domElement.addEventListener("pointermove", onPointerMove, { passive: false });
    domElement.addEventListener("pointerup", onPointerUp, { passive: false });
    domElement.addEventListener("pointercancel", onPointerUp, { passive: false });
    domElement.addEventListener("wheel", onWheel, { passive: false });
    domElement.addEventListener("contextmenu", onContextMenu);

    this.update = function update(dt) {
      const t = (typeof dt === "number" && isFinite(dt)) ? dt : 0.016;

      // Re-sync spherical from the current camera position (handles external camera moves).
      updateSphericalFromCamera();

      spherical.theta += sphericalDelta.theta;
      spherical.phi += sphericalDelta.phi;

      // Clamp polar angle.
      const maxPhi = Math.min(scope.maxPolarAngle, Math.PI - EPS);
      spherical.phi = Math.max(EPS, Math.min(maxPhi, spherical.phi));

      if (zoomDelta !== 0) {
        // "Infinite" dolly zoom: constant world-units step per notch so zooming
        // feels the same no matter how close/far you already are (visualization only).
        const notches = zoomDelta / 120;
        const stepFt = Math.max(0.1, Number(scope.zoomStepFt != null ? scope.zoomStepFt : 120));
        const boost = (keys.ShiftLeft || keys.ShiftRight) ? 4 : 1;
        const slow = (keys.ControlLeft || keys.ControlRight) ? 0.25 : 1;
        const deltaR = notches * stepFt * boost * slow;
        if (Number.isFinite(deltaR) && deltaR !== 0) {
          // Keep zoom behavior usable at the outskirts of the map by gently
          // sliding the orbit target toward the cursor position (at the current
          // target depth) while zooming in.
          if (deltaR < 0 && zoomFromWheel && getCursorHitOnTargetPlane(vCursorHit)) {
            const rel = Math.min(1, Math.abs(deltaR) / Math.max(1, spherical.radius));
            const alpha = Math.min(0.55, Math.max(0.12, rel * 1.6));
            scope.target.add(v3tmp2.copy(vCursorHit).sub(scope.target).multiplyScalar(alpha));
          }
          spherical.radius += deltaR;
        }
      }
      spherical.radius = Math.max(scope.minDistance, Math.min(scope.maxDistance, spherical.radius));

      scope.target.add(panOffset);

      const newPos = v3tmp.setFromSpherical(spherical).add(scope.target);
      camera.position.copy(newPos);
      camera.lookAt(scope.target);

      // Keyboard translation (6DOF-style): move both camera and target so framing stays stable.
      if (scope.enableKeys) {
        const boost = (keys.ShiftLeft || keys.ShiftRight) ? 4 : 1;
        const slow = (keys.ControlLeft || keys.ControlRight) ? 0.25 : 1;
        const speed = Number(scope.keyMoveSpeed || 0) * boost * slow;
        const amt = speed * t;

        if (amt > 0) {
          vKeyMove.set(0, 0, 0);

          if (keys.KeyW || keys.ArrowUp) vKeyMove.z -= 1;
          if (keys.KeyS || keys.ArrowDown) vKeyMove.z += 1;
          if (keys.KeyA || keys.ArrowLeft) vKeyMove.x -= 1;
          if (keys.KeyD || keys.ArrowRight) vKeyMove.x += 1;

          // Q/E per user request (with R/F and PgUp/PgDn as alternates): vertical movement.
          let upDown = 0;
          if (keys.KeyE || keys.KeyR || keys.PageUp) upDown += 1;
          if (keys.KeyQ || keys.KeyF || keys.PageDown) upDown -= 1;

          // Horizontal move (camera-relative, constrained to ground plane).
          if (vKeyMove.x !== 0 || vKeyMove.z !== 0) {
            vKeyWorld.set(vKeyMove.x, 0, vKeyMove.z);
            vKeyWorld.normalize();
            vKeyWorld.applyQuaternion(camera.quaternion);
            vKeyWorld.y = 0;
            if (vKeyWorld.lengthSq() > 0) {
              vKeyWorld.normalize();
              vKeyWorld.multiplyScalar(amt);
              camera.position.add(vKeyWorld);
              scope.target.add(vKeyWorld);
            }
          }

          // Vertical move (world-up).
          if (upDown !== 0) {
            const dy = upDown * amt;
            camera.position.addScaledVector(WORLD_UP, dy);
            scope.target.addScaledVector(WORLD_UP, dy);
          }
        }
      }


      // Allow navigation above/below the basemap plane (per user request).

      if (scope.enableDamping) {
        sphericalDelta.theta *= (1 - scope.dampingFactor);
        sphericalDelta.phi *= (1 - scope.dampingFactor);
        panOffset.multiplyScalar(1 - scope.dampingFactor);
      } else {
        sphericalDelta.theta = 0;
        sphericalDelta.phi = 0;
        panOffset.set(0, 0, 0);
      }

      zoomDelta = 0;
      zoomFromWheel = false;
    };

    this.dispose = function dispose() {
      try { document.removeEventListener("keydown", onKeyDown, true); } catch (_) {}
      try { document.removeEventListener("keyup", onKeyUp, true); } catch (_) {}
      try { domElement.removeEventListener("pointerdown", onPointerDown); } catch (_) {}
      try { domElement.removeEventListener("pointermove", onPointerMove); } catch (_) {}
      try { domElement.removeEventListener("pointerup", onPointerUp); } catch (_) {}
      try { domElement.removeEventListener("pointercancel", onPointerUp); } catch (_) {}
      try { domElement.removeEventListener("wheel", onWheel); } catch (_) {}
      try { domElement.removeEventListener("contextmenu", onContextMenu); } catch (_) {}
    };
  }


  function FreeCamControlsLite(camera, domElement) {
    // Minimal 6DOF camera controller (fly / FPS-style).
    // - WASD: move forward/back/strafe
    // - R/F: up/down
    // - Q/E: roll
    // - Drag (LMB): look (yaw/pitch)
    // - Drag (RMB or MMB): pan/strafe (camera-relative)
    // - Wheel: dolly (zoom)
    // This is visualization-only and does not modify QC logic.

    const scope = this;
    this.object = camera;
    this.domElement = domElement;

    this.target = new THREE.Vector3();

    // Prevent drag navigation from firing selection clicks.
    this.suppressClick = false;

    this.lookSpeed = 0.0016;      // radians per pixel
    this.panSpeed = 0.35;         // ft per pixel (scaled by camera height)
    this.dollySpeed = 0.35;       // ft per wheel delta unit
    this.moveSpeed = 280;         // ft per second (tuned a bit faster)
    this.rollSpeed = 1.00;        // rad/sec

    // Click suppression so drag navigation does not accidentally select poles/midspans.
    this.suppressClick = false;

    const keys = Object.create(null);
    let dragging = false;
    let dragMode = "look"; // "look" | "pan"
    let activePointerId = null;
    let lastX = 0;
    let lastY = 0;
    let movedPx = 0;

    const v3tmp = new THREE.Vector3();
    const v3tmp2 = new THREE.Vector3();

    function onKeyDown(e) {
      keys[e.code] = true;
    }
    function onKeyUp(e) {
      keys[e.code] = false;
    }

    function onPointerDown(e) {
      // Mouse buttons (inverted per user request):
      // - Left button: pan/translate
      // - Right button: look (yaw/pitch)
      // - Middle button: pan/translate
      if (e.pointerType === "mouse") {
        if (e.button === 0 || e.button === 1) dragMode = "pan";
        else dragMode = "look";
      } else {
        dragMode = "look";
      }

      dragging = true;
      activePointerId = e.pointerId;
      lastX = e.clientX;
      lastY = e.clientY;
      movedPx = 0;

      try { scope.domElement.setPointerCapture(activePointerId); } catch (_) {}
      e.preventDefault();
    }

    function clampPitch() {
      // Avoid gimbal lock; keep pitch just under +/-90 degrees.
      const limit = Math.PI / 2 - 0.01;
      camera.rotation.x = Math.max(-limit, Math.min(limit, camera.rotation.x));
    }

    function onPointerMove(e) {
      if (!dragging || e.pointerId !== activePointerId) return;

      const dx = e.clientX - lastX;
      const dy = e.clientY - lastY;
      lastX = e.clientX;
      lastY = e.clientY;

      movedPx += Math.abs(dx) + Math.abs(dy);

      camera.rotation.order = "YXZ";

      if (dragMode === "look") {
        // Natural directions: drag right -> look right; drag up -> look up.
        camera.rotation.y -= dx * scope.lookSpeed;
        camera.rotation.x -= dy * scope.lookSpeed;
        clampPitch();
      } else {
        // Pan/strafe relative to camera.
        const heightScale = Math.max(0.25, Math.min(8, (camera.position.y + 80) / 480));
        const pan = scope.panSpeed * heightScale;

        const right = v3tmp.setFromMatrixColumn(camera.matrix, 0).normalize();
        const up = v3tmp2.setFromMatrixColumn(camera.matrix, 1).normalize();

        camera.position.add(right.multiplyScalar(-dx * pan));
        camera.position.add(up.multiplyScalar(dy * pan));
      }

      e.preventDefault();
    }

    function onPointerUp(e) {
      if (!dragging || e.pointerId !== activePointerId) return;
      dragging = false;
      try { scope.domElement.releasePointerCapture(activePointerId); } catch (_) {}
      activePointerId = null;

      if (movedPx > 3) scope.suppressClick = true;
      e.preventDefault();
    }

    function onWheel(e) {
      // Wheel zoom: dolly along view direction.
      const deltaY = Number(e.deltaY || 0);
      if (!isFinite(deltaY) || deltaY === 0) return;

      const dir = v3tmp.set(0, 0, -1).applyQuaternion(camera.quaternion).normalize();
      const heightScale = Math.max(0.6, Math.min(10, (camera.position.y + 120) / 520));
      const amt = deltaY * scope.dollySpeed * heightScale;

      // Natural: wheel down (positive deltaY) -> zoom out (move backward)
      camera.position.add(dir.multiplyScalar(-amt));
      scope.suppressClick = true;

      e.preventDefault();
    }

    function onContextMenu(e) {
      e.preventDefault();
    }

    window.addEventListener("keydown", onKeyDown);
    window.addEventListener("keyup", onKeyUp);
    scope.domElement.addEventListener("pointerdown", onPointerDown, { passive: false });
    scope.domElement.addEventListener("pointermove", onPointerMove, { passive: false });
    scope.domElement.addEventListener("pointerup", onPointerUp, { passive: false });
    scope.domElement.addEventListener("pointercancel", onPointerUp, { passive: false });
    scope.domElement.addEventListener("wheel", onWheel, { passive: false });
    scope.domElement.addEventListener("contextmenu", onContextMenu, { passive: false });

    this.update = function update(dt) {
      const t = (typeof dt === "number" && isFinite(dt)) ? dt : 0.016;

      // Translation
      const boost = (keys.ShiftLeft || keys.ShiftRight) ? 4 : 1;
      const slow = (keys.ControlLeft || keys.ControlRight) ? 0.25 : 1;
      const speed = scope.moveSpeed * boost * slow;

      const move = v3tmp.set(0, 0, 0);
      if (keys.KeyW || keys.ArrowUp) move.z -= 1;
      if (keys.KeyS || keys.ArrowDown) move.z += 1;
      if (keys.KeyA || keys.ArrowLeft) move.x -= 1;
      if (keys.KeyD || keys.ArrowRight) move.x += 1;
      if (keys.KeyR || keys.PageUp) move.y += 1;
      if (keys.KeyF || keys.PageDown) move.y -= 1;

      if (move.lengthSq() > 0) {
        move.normalize();
        move.applyQuaternion(camera.quaternion);
        camera.position.add(move.multiplyScalar(speed * t));
      }

      // Roll
      let rollDir = 0;
      if (keys.KeyQ) rollDir += 1;
      if (keys.KeyE) rollDir -= 1;
      if (rollDir !== 0) {
        camera.rotation.order = "YXZ";
        camera.rotation.z += rollDir * scope.rollSpeed * t;
      }
    };

    this.dispose = function dispose() {
      try { window.removeEventListener("keydown", onKeyDown); } catch (_) {}
      try { window.removeEventListener("keyup", onKeyUp); } catch (_) {}
      try { scope.domElement.removeEventListener("pointerdown", onPointerDown); } catch (_) {}
      try { scope.domElement.removeEventListener("pointermove", onPointerMove); } catch (_) {}
      try { scope.domElement.removeEventListener("pointerup", onPointerUp); } catch (_) {}
      try { scope.domElement.removeEventListener("pointercancel", onPointerUp); } catch (_) {}
      try { scope.domElement.removeEventListener("wheel", onWheel); } catch (_) {}
      try { scope.domElement.removeEventListener("contextmenu", onContextMenu); } catch (_) {}
    };
  }

  function setViewMode(mode) {
    const next = (mode === "3d") ? "3d" : "2d";
    if (next === viewMode) return;
    viewMode = next;

    // Toggle UI state for 3D-only controls (visualization only).
    try { document.body.classList.toggle("is-3d", viewMode === "3d"); } catch (_) {}
    try { if (els.basemap3d) els.basemap3d.value = currentBasemap || "Dark"; } catch (_) {}

    const mapEl = document.getElementById("map");
    const threeEl = els.map3d || document.getElementById("map3d");

    if (viewMode === "3d") {
      if (mapEl) mapEl.classList.add("is-hidden");
      if (threeEl) threeEl.classList.remove("is-hidden");

      const st = ensureThree();
      maximizeMapHeight();
      resizeThreeRenderer();

      // If 3D could not initialize (CDN blocked, OrbitControls missing, WebGL unavailable),
      // immediately fall back to 2D so the user never ends up with a blank panel.
      if (!st) {
        if (threeEl) threeEl.classList.add("is-hidden");
        if (mapEl) mapEl.classList.remove("is-hidden");
        viewMode = "2d";
        if (els.view2d) els.view2d.checked = true;
        logLine("3D view could not initialize in this browser/session. Staying in 2D.");
        if (map) setTimeout(() => map.invalidateSize(), 60);
        return;
      }

      if (threeDirty) rebuildThreeScene();
      startThreeLoop();
      refreshThreeVisibility();

      // Ensure keyboard navigation is immediately available in 3D.
      try {
        if (threeState && threeState.renderer && threeState.renderer.domElement && typeof threeState.renderer.domElement.focus === "function") {
          threeState.renderer.domElement.focus();
        }
      } catch (_) {}
    } else {
      if (threeEl) threeEl.classList.add("is-hidden");
      if (mapEl) mapEl.classList.remove("is-hidden");

      stopThreeLoop();
      maximizeMapHeight();
      if (map) {
        // Leaflet needs invalidateSize after container becomes visible again.
        setTimeout(() => map.invalidateSize(), 60);
      }
    }
  }

  function ensureThree() {
    if (threeState) return threeState;
    const container = els.map3d || document.getElementById("map3d");
    if (!container) return null;

    if (!window.THREE || !THREE.WebGLRenderer) {
      logLine("3D view unavailable: Three.js failed to load.");
      return null;
    }

    try {
      // Clear anything previously rendered into the container.
      container.innerHTML = "";

      const renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
      renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));
      renderer.setSize(container.clientWidth || 800, container.clientHeight || 600, false);
      // Match the UI's sleek dark theme.
      renderer.setClearColor(0x050a16, 1);

      // Ensure textures render with correct brightness/contrast (visualization only).
      try {
        if ("outputColorSpace" in renderer) renderer.outputColorSpace = THREE.SRGBColorSpace;
        else if ("outputEncoding" in renderer) renderer.outputEncoding = THREE.sRGBEncoding;
      } catch (_) {}

      container.appendChild(renderer.domElement);

      // Make the 3D canvas focusable so WASD/QE controls work reliably.
      try {
        renderer.domElement.tabIndex = 0;
        renderer.domElement.style.outline = "none";
        renderer.domElement.addEventListener("pointerdown", () => {
          try { renderer.domElement.focus(); } catch (_) {}
        }, { passive: true });
      } catch (_) {}

      // 3D label overlay (SCIDs). Rendered as HTML so it stays crisp.
      const labelLayer = document.createElement("div");
      labelLayer.className = "three-label-layer";
      container.appendChild(labelLayer);

      const scene = new THREE.Scene();
      scene.fog = new THREE.Fog(0x050a16, 0, 20000);

      const camera = new THREE.PerspectiveCamera(
        60,
        (container.clientWidth || 800) / (container.clientHeight || 600),
        0.1,
        500000
      );
      camera.position.set(0, 650, 950);

      // Orbit controls: pan + zoom + rotate.
      // Use built-in OrbitControls when available; otherwise, use the lightweight fallback.
      const controls = makeOrbitControls(camera, renderer.domElement);
      // Ensure default tuning is applied consistently across either controls implementation.
      controls.enableDamping = true;
      controls.dampingFactor = 0.06;
      controls.screenSpacePanning = true;
      controls.minDistance = 0.001;
      controls.maxDistance = 250000;
      controls.maxPolarAngle = Math.PI;

      const ambient = new THREE.AmbientLight(0xffffff, 0.72);
      const dir = new THREE.DirectionalLight(0xffffff, 0.85);
      dir.position.set(600, 1200, 450);
      scene.add(ambient);
      scene.add(dir);

      const grid = new THREE.GridHelper(4000, 80, 0x334155, 0x1e293b);
      grid.material.opacity = 0.22;
      grid.material.transparent = true;
      scene.add(grid);

      const root = new THREE.Group();
      scene.add(root);

      const raycaster = new THREE.Raycaster();
      const pointer = new THREE.Vector2();

      const state = {
        container,
        labelLayer,
        renderer,
        scene,
        camera,
        controls,
        root,
        grid,
        raycaster,
        pointer,
        animationId: null,

        // Built per-model
        origin: null,
        bounds: null,
        groups: {
          ground: null,
          spans: null,
          guying: null,
          poles: null,
          equipment: null,
          midspans: null,
          issues: null,
        },
        poleGroupsById: new Map(),
        midGroupsById: new Map(),
        scidLabels3d: new Map(),
        pickables: [],

        // Node anchors (poles + midspans): used to keep the camera pivot pole-oriented.
        nodeAnchors: [],
        // Pole-only anchors for the "double click to focus pole" feature.
        poleAnchors: [],
        pivotSnapFt: 250,

        // Visualization-only: materials that receive a subtle pulse (installing company highlight).
        installPulseMaterials: [],
      };

      // Picking for Details panel
      renderer.domElement.addEventListener("click", (ev) => {
        try {
          if (!threeState) return;
          // Do not treat a navigation drag as a selection click.
          if (threeState.controls && threeState.controls.suppressClick) {
            threeState.controls.suppressClick = false;
            return;
          }
          const rect = renderer.domElement.getBoundingClientRect();
          const x = ((ev.clientX - rect.left) / rect.width) * 2 - 1;
          const y = -(((ev.clientY - rect.top) / rect.height) * 2 - 1);
          pointer.set(x, y);
          raycaster.setFromCamera(pointer, camera);
          const hits = raycaster.intersectObjects(state.pickables, true);
          if (!hits || !hits.length) return;
          const hit = hits[0] && hits[0].object ? hits[0].object : null;
          if (!hit) return;

          let o = hit;
          while (o && o.parent && !o.userData) o = o.parent;
          const ud = o && o.userData ? o.userData : (hit.userData || {});
          if (ud && ud.entityType && ud.entityId != null) {
            selectEntity({ type: String(ud.entityType), id: String(ud.entityId), latlng: null });
          }
        } catch (_) {}
      });

      threeState = state;
      return threeState;
    } catch (err) {
      // If WebGL initialization fails for any reason, keep the app usable in 2D.
      try { logLine(`3D view unavailable: ${err && err.message ? err.message : String(err)}`); } catch (_) {}
      try { container.innerHTML = ""; } catch (_) {}
      threeState = null;
      return null;
    }
  }

  function disposeThreeObject(obj) {
    if (!obj) return;
    obj.traverse((o) => {
      if (o.geometry && typeof o.geometry.dispose === "function") {
        try { o.geometry.dispose(); } catch (_) {}
      }
      if (o.material) {
        const mats = Array.isArray(o.material) ? o.material : [o.material];
        for (const m of mats) {
          if (m && typeof m.dispose === "function") {
            try { m.dispose(); } catch (_) {}
          }
        }
      }
    });
  }

  function rebuildThreeScene() {
    if (!threeState) return;

    // Clear prior world
    if (threeState.root) {
      disposeThreeObject(threeState.root);
      threeState.root.clear();
    }
    threeState.pickables = [];
    threeState.nodeAnchors = [];
    threeState.poleAnchors = [];
    threeState.poleGroupsById = new Map();
    threeState.midGroupsById = new Map();
    threeState.installPulseMaterials = [];
    threeState.wirePickIndex = new Map();
    threeState.guyPickIndex = new Map();

    // Reset 3D SCID label overlay (visualization only).
    try { if (threeState.labelLayer) threeState.labelLayer.innerHTML = ""; } catch (_) {}
    threeState.scidLabels3d = new Map();

    // Camera focus anchor (visualization only): used to keep the initial 3D
    // framing centered on the first node in the model.
    threeState.firstNode = null;

    threeState.groups.ground = new THREE.Group();
    threeState.groups.spans = new THREE.Group();
    threeState.groups.guying = new THREE.Group();
    threeState.groups.poles = new THREE.Group();
    // Equipment meshes are stored inside each pole group so they respect status filtering.
    threeState.groups.equipment = null;
    threeState.groups.midspans = new THREE.Group();
    threeState.groups.issues = new THREE.Group();
    threeState.root.add(threeState.groups.ground);
    threeState.root.add(threeState.groups.spans);
    threeState.root.add(threeState.groups.guying);
    threeState.root.add(threeState.groups.poles);
    threeState.root.add(threeState.groups.midspans);
    threeState.root.add(threeState.groups.issues);

    if (!model) {
      threeDirty = false;
      return;
    }

    const points = [];
    for (const p of (model.poles || [])) {
      if (typeof p.lat === "number" && typeof p.lon === "number") points.push([p.lat, p.lon]);
    }
    for (const m of (model.midspanPoints || [])) {
      if (typeof m.lat === "number" && typeof m.lon === "number") points.push([m.lat, m.lon]);
    }
    if (!points.length) {
      threeDirty = false;
      return;
    }

    // Use mean lat/lon as projection origin.
    const lat0 = points.reduce((a, b) => a + b[0], 0) / points.length;
    const lon0 = points.reduce((a, b) => a + b[1], 0) / points.length;
    threeState.origin = { lat0, lon0 };

    const FT_PER_M = 3.28084;
    const R = 6378137;
    const toRad = (d) => (d * Math.PI) / 180;

    // NOTE (visualization only): Use Web Mercator for the 3D X/Z ground projection.
    // This matches Leaflet + slippy-map tile math and prevents tile-edge misalignment
    // (roads/labels should stitch correctly between tiles).
    const mercatorXY = (lat, lon) => {
      const phi = toRad(lat);
      const lam = toRad(lon);
      const x = R * lam;
      const y = R * Math.log(Math.tan(Math.PI / 4 + phi / 2));
      return { x, y };
    };
    const originM = mercatorXY(lat0, lon0);
    // Visualization-only: mirror the 3D X axis once (east/west) to match the expected orientation.
    // This affects ONLY the 3D rendering (basemap + 3D overlays), not QC logic.
    const MIRROR_X = true;
    const latLonToXZFt = (lat, lon) => {
      const m = mercatorXY(lat, lon);
      return {
        x: (m.x - originM.x) * FT_PER_M * (MIRROR_X ? -1 : 1),
        z: (m.y - originM.y) * FT_PER_M,
      };
    };

    const statusColor = (st) => {
      if (st === "pass") return 0x22c55e;
      if (st === "warn") return 0xf59e0b;
      if (st === "fail") return 0xef4444;
      return 0x94a3b8;
    };

    // Installing Company highlight (3D only): green dull flashing overlay on the wire + a bead on the pole.
    // NOTE: the rule selection lives under rules.midspan.installingCompany.
    const installingCompanyKey = (rules && rules.midspan && rules.midspan.installingCompany)
      ? normalizeOwnerKey(rules.midspan.installingCompany)
      : "";

    const INSTALL_PULSE_COLOR = 0x22c55e;
    const installPulseWireMat = installingCompanyKey ? new THREE.MeshStandardMaterial({
      color: INSTALL_PULSE_COLOR,
      emissive: INSTALL_PULSE_COLOR,
      emissiveIntensity: 0.10,
      roughness: 0.42,
      metalness: 0.02,
      transparent: true,
      opacity: 0.34,
      depthWrite: false,
      blending: THREE.AdditiveBlending,
    }) : null;

    const installPulseBeadMat = installingCompanyKey ? new THREE.MeshStandardMaterial({
      color: INSTALL_PULSE_COLOR,
      emissive: INSTALL_PULSE_COLOR,
      emissiveIntensity: 0.16,
      roughness: 0.36,
      metalness: 0.02,
      transparent: true,
      opacity: 0.52,
      depthWrite: false,
      blending: THREE.AdditiveBlending,
    }) : null;

    if (installPulseWireMat) {
      // Make the installing-company highlight OBVIOUS (requested): opacity pulses 25% → 100%.
      installPulseWireMat.userData.pulse = { baseOpacity: 0.25, ampOpacity: 0.75, baseEmissiveIntensity: 0.06, ampEmissiveIntensity: 0.12 };
      threeState.installPulseMaterials.push(installPulseWireMat);
    }
    if (installPulseBeadMat) {
      // Make the installing-company pole bead OBVIOUS (requested): opacity pulses 25% → 100%.
      installPulseBeadMat.userData.pulse = { baseOpacity: 0.25, ampOpacity: 0.75, baseEmissiveIntensity: 0.10, ampEmissiveIntensity: 0.16 };
      threeState.installPulseMaterials.push(installPulseBeadMat);
    }



    // 3D ground basemap (streets / light / imagery) using slippy-map tiles.
    // Visualization only: does not affect QC logic or results.
    (function buildGroundTiles() {
      try {
        const groundGroup = threeState.groups && threeState.groups.ground ? threeState.groups.ground : null;
        if (!groundGroup) return;

        // Hide grid when we have a basemap for better legibility.
        try { if (threeState.grid) threeState.grid.visible = false; } catch (_) {}

        // Compute bounds from all known points.
        let minLat = Infinity, maxLat = -Infinity, minLon = Infinity, maxLon = -Infinity;
        for (const pt of points) {
          const lat = Number(pt[0]);
          const lon = Number(pt[1]);
          if (!isFinite(lat) || !isFinite(lon)) continue;
          if (lat < minLat) minLat = lat;
          if (lat > maxLat) maxLat = lat;
          if (lon < minLon) minLon = lon;
          if (lon > maxLon) maxLon = lon;
        }
        if (!isFinite(minLat) || !isFinite(minLon)) return;

        // Add a modest padding so the basemap extends beyond the outermost poles.
        const padLat = Math.max(0.0015, (maxLat - minLat) * 0.10);
        const padLon = Math.max(0.0015, (maxLon - minLon) * 0.10);
        minLat -= padLat;
        maxLat += padLat;
        minLon -= padLon;
        maxLon += padLon;

        const toRad2 = (d) => (d * Math.PI) / 180;
        const toDeg2 = (r) => (r * 180) / Math.PI;

        const lonToTileX = (lon, z) => {
          const n = 2 ** z;
          return Math.floor(((lon + 180) / 360) * n);
        };
        const latToTileY = (lat, z) => {
          const n = 2 ** z;
          const φ = toRad2(lat);
          const y = (1 - Math.log(Math.tan(φ) + 1 / Math.cos(φ)) / Math.PI) / 2;
          return Math.floor(y * n);
        };
        const tileXToLon = (x, z) => {
          const n = 2 ** z;
          return (x / n) * 360 - 180;
        };
        const tileYToLat = (y, z) => {
          const n = 2 ** z;
          const a = Math.PI * (1 - (2 * y) / n);
          return toDeg2(Math.atan(Math.sinh(a)));
        };

        const tileBoundsForZoom = (z) => {
          const x1 = lonToTileX(minLon, z);
          const x2 = lonToTileX(maxLon, z);
          const y1 = latToTileY(maxLat, z);
          const y2 = latToTileY(minLat, z);
          const xMin = Math.min(x1, x2);
          const xMax = Math.max(x1, x2);
          const yMin = Math.min(y1, y2);
          const yMax = Math.max(y1, y2);
          const count = (xMax - xMin + 1) * (yMax - yMin + 1);
          return { xMin, xMax, yMin, yMax, count };
        };

        // Choose the highest zoom level that does not require too many tiles.
        const MAX_TILES = 90;
        let zChosen = 17;
        let tb = null;
        for (const z of [19, 18, 17, 16, 15, 14, 13, 12]) {
          const t = tileBoundsForZoom(z);
          if (t.count <= MAX_TILES) {
            zChosen = z;
            tb = t;
            break;
          }
          tb = t;
        }
        if (!tb) tb = tileBoundsForZoom(zChosen);

        const basemap = (currentBasemap || "Dark");
        const tileUrl = (x, y, z) => {
          if (basemap === "Imagery") {
            return `https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/${z}/${y}/${x}`;
          }
          const sub = (x + y) % 4;
          const s = ["a", "b", "c", "d"][sub] || "a";
          const style = (basemap === "Light") ? "light_all" : "dark_all";
          return `https://${s}.basemaps.cartocdn.com/${style}/${z}/${x}/${y}.png`;
        };

        const texLoader = new THREE.TextureLoader();
        try { texLoader.setCrossOrigin("anonymous"); } catch (_) { texLoader.crossOrigin = "anonymous"; }

        // Build tile planes.
        const yPlane = 0.02;
        for (let x = tb.xMin; x <= tb.xMax; x++) {
          for (let y = tb.yMin; y <= tb.yMax; y++) {
            const lonL = tileXToLon(x, zChosen);
            const lonR = tileXToLon(x + 1, zChosen);
            const latT = tileYToLat(y, zChosen);
            const latB = tileYToLat(y + 1, zChosen);

            const xLeft = latLonToXZFt(lat0, lonL).x;
            const xRight = latLonToXZFt(lat0, lonR).x;
            const zTop = latLonToXZFt(latT, lon0).z;
            const zBottom = latLonToXZFt(latB, lon0).z;

            const w = Math.abs(xRight - xLeft);
            const d = Math.abs(zBottom - zTop);
            if (!isFinite(w) || !isFinite(d) || w <= 0 || d <= 0) continue;

            const geo = new THREE.PlaneGeometry(w, d, 1, 1);
            geo.rotateX(-Math.PI / 2);


            // If we mirrored the world X axis, flip U on the front-face UVs so
            // the slippy-map imagery still stitches correctly across tile edges.
            if (MIRROR_X) {
              try {
                const uvF = geo.getAttribute("uv");
                if (uvF && uvF.count) {
                  for (let i = 0; i < uvF.count; i++) {
                    uvF.setX(i, 1 - uvF.getX(i));
                  }
                  uvF.needsUpdate = true;
                }
              } catch (_) {}
            }

            // IMPORTANT (visualization only): When the camera ends up viewing the basemap
            // from the reverse side, a single textured plane will appear optically mirrored
            // (east/west swapped). To make the basemap robust from either side without
            // changing QC logic, render a second plane for the backface with U-flipped UVs.
            const geoBack = geo.clone();
            try {
              const uv = geoBack.getAttribute("uv");
              if (uv && uv.count) {
                for (let i = 0; i < uv.count; i++) {
                  uv.setX(i, 1 - uv.getX(i));
                }
                uv.needsUpdate = true;
              }
            } catch (_) {}

            const matFront = new THREE.MeshBasicMaterial({
              // Do not tint tiles (keeps streets/labels legible across basemaps)
              color: 0xffffff,
              transparent: true,
              opacity: 1.0,
              side: THREE.FrontSide,
              fog: false,
            });

            const matBack = new THREE.MeshBasicMaterial({
              color: 0xffffff,
              transparent: true,
              opacity: 1.0,
              side: THREE.BackSide,
              fog: false,
            });

            const mesh = new THREE.Mesh(geo, matFront);
            const meshBack = new THREE.Mesh(geoBack, matBack);
            const cx = (xLeft + xRight) / 2;
            const cz = (zTop + zBottom) / 2;
            mesh.position.set(cx, yPlane, cz);
            meshBack.position.set(cx, yPlane, cz);
            mesh.renderOrder = -50;
            meshBack.renderOrder = -50;
            groundGroup.add(mesh);
            groundGroup.add(meshBack);

            // Load tile texture asynchronously.
            const url = tileUrl(x, y, zChosen);
            texLoader.load(
              url,
              (tex) => {
                try {
                  // three@0.160 supports colorSpace. Keep it safe across versions.
                  if (tex && "colorSpace" in tex) {
                    tex.colorSpace = THREE.SRGBColorSpace;
                  } else if (tex && "encoding" in tex) {
                    tex.encoding = THREE.sRGBEncoding;
                  }

                  // Slippy-map tiles are north-up. The ground-plane UVs after rotation
                  // are inverted relative to the default TextureLoader flip behavior,
                  // which causes adjacent tiles to fail to stitch correctly.
                  tex.flipY = false;

                  // Prevent visible seams between adjacent tiles by avoiding
                  // mipmapped sampling bleeding across tile edges.
                  try {
                    tex.wrapS = THREE.ClampToEdgeWrapping;
                    tex.wrapT = THREE.ClampToEdgeWrapping;
                    tex.minFilter = THREE.LinearFilter;
                    tex.magFilter = THREE.LinearFilter;
                    tex.generateMipmaps = false;
                  } catch (_) {}

                  tex.anisotropy = Math.min(8, threeState.renderer.capabilities.getMaxAnisotropy());
                  tex.needsUpdate = true;
                  matFront.map = tex;
                  matBack.map = tex;
                  matFront.needsUpdate = true;
                  matBack.needsUpdate = true;
                } catch (_) {}
              },
              undefined,
              () => {
                // Leave the solid fill if tile load fails.
              }
            );
          }
        }
      } catch (_) {}
    })();

    // Build quick indexes
    const spanByConnection = new Map();
    for (const s of (model.spans || [])) {
      if (s && s.connectionId != null) spanByConnection.set(String(s.connectionId), s);
    }

    const poleById = new Map();
    for (const p of (model.poles || [])) poleById.set(String(p.poleId), p);

    const poleWireHeightByPoleTrace = new Map();
    for (const p of (model.poles || [])) {
      const poleId = String(p.poleId);
      const atts = Array.isArray(p.attachments) ? p.attachments : [];
      for (const a of atts) {
        if (!a) continue;
        if (a.traceId == null) continue;
        const hIn = (a.proposedIn != null ? a.proposedIn : a.existingIn);
        if (hIn == null) continue;
        const key = `${poleId}|${String(a.traceId)}`;
        const prev = poleWireHeightByPoleTrace.get(key);
        if (!prev || Number(hIn) > Number(prev)) poleWireHeightByPoleTrace.set(key, Number(hIn));
      }
    }

    const attachmentById = new Map();
    for (const p of (model.poles || [])) {
      for (const a of (p.attachments || [])) {
        if (a && a.id != null) attachmentById.set(String(a.id), { poleId: String(p.poleId), att: a });
      }
    }

    const measureById = new Map();
    for (const ms of (model.midspanPoints || [])) {
      for (const m of (ms.measures || [])) {
        const id = (m && m.id != null) ? String(m.id) : "";
        if (id) measureById.set(id, { midspanId: String(ms.midspanId), measure: m });
      }
    }

    const issueSets = (() => {
      const failAttachmentIds = new Set();
      const failMeasureIds = new Set();
      const warnAttachmentIds = new Set();
      const warnMeasureIds = new Set();
      const orderAttachmentIds = new Set();
      const orderMeasureIds = new Set();

      const all = qcResults && Array.isArray(qcResults.issues) ? qcResults.issues : [];
      for (const iss of all) {
        const rc = String(iss && iss.ruleCode || "");
        const ctx = iss && iss.context ? iss.context : {};
        const aIds = Array.isArray(ctx.attachmentIds) ? ctx.attachmentIds : [];
        const mIds = Array.isArray(ctx.measureIds) ? ctx.measureIds : [];

        if (iss && iss.severity === "FAIL") {
          for (const id of aIds) if (id != null) failAttachmentIds.add(String(id));
          for (const id of mIds) if (id != null) failMeasureIds.add(String(id));
        }
        if (iss && iss.severity === "WARN") {
          for (const id of aIds) if (id != null) warnAttachmentIds.add(String(id));
          for (const id of mIds) if (id != null) warnMeasureIds.add(String(id));
        }
        if (rc.startsWith("ORDER.COMM")) {
          for (const id of aIds) if (id != null) orderAttachmentIds.add(String(id));
          for (const id of mIds) if (id != null) orderMeasureIds.add(String(id));
        }
      }
      return { failAttachmentIds, failMeasureIds, warnAttachmentIds, warnMeasureIds, orderAttachmentIds, orderMeasureIds };
    })();

    // --- 3D helpers (visualization only) ---
    const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));

    // Deterministic hash for stable "default" equipment bearings when a job does not
    // provide explicit direction data.
    const hashU32 = (s) => {
      const str = String(s || "");
      let h = 2166136261;
      for (let i = 0; i < str.length; i++) {
        h ^= str.charCodeAt(i);
        h = Math.imul(h, 16777619);
      }
      return (h >>> 0);
    };

    const parseBearingDeg = (text) => {
      const t = String(text || "");
      // Common patterns: "bearing 123", "@123°", "123 deg".
      const m = t.match(/(?:bearing\s*[:=]?\s*|@\s*)(\d{1,3}(?:\.\d+)?)/i) || t.match(/(\d{1,3}(?:\.\d+)?)\s*(?:deg|°)/i);
      if (!m) return null;
      const v = Number(m[1]);
      if (!Number.isFinite(v)) return null;
      return ((v % 360) + 360) % 360;
    };

    const bearingVecXZ = (bearingDeg) => {
      const rad = toRad(bearingDeg);
      // Bearing: 0=north (+z), 90=east (-x due to MIRROR_X), 180=south (-z), 270=west (+x)
      return new THREE.Vector3(-Math.sin(rad), 0, Math.cos(rad));
    };

    const parseFirstFeet = (text) => {
      const t = String(text || "");
      const m = t.match(/(\d+(?:\.\d+)?)\s*(?:ft|')/i);
      const v = m ? Number(m[1]) : NaN;
      return Number.isFinite(v) ? v : null;
    };

    const cylinderBetween = (p1, p2, radius, material) => {
      const dir = new THREE.Vector3().subVectors(p2, p1);
      const len = dir.length();
      if (!Number.isFinite(len) || len <= 0.001) return null;
      const geom = new THREE.CylinderGeometry(radius, radius, len, 10, 1);
      const mesh = new THREE.Mesh(geom, material);
      const mid = new THREE.Vector3().addVectors(p1, p2).multiplyScalar(0.5);
      mesh.position.copy(mid);
      mesh.quaternion.setFromUnitVectors(new THREE.Vector3(0, 1, 0), dir.normalize());
      return mesh;
    };

    // Span lines (ground reference)
    {
      const matPolePole = new THREE.LineBasicMaterial({ color: 0x38bdf8, transparent: true, opacity: 0.32 });
      const matOther = new THREE.LineDashedMaterial({ color: 0x94a3b8, transparent: true, opacity: 0.22, dashSize: 28, gapSize: 22 });

      for (const s of (model.spans || [])) {
        if (!s) continue;
        if (s.aLat == null || s.aLon == null || s.bLat == null || s.bLon == null) continue;
        const a = latLonToXZFt(Number(s.aLat), Number(s.aLon));
        const b = latLonToXZFt(Number(s.bLat), Number(s.bLon));
        const g = new THREE.BufferGeometry().setFromPoints([
          new THREE.Vector3(a.x, 0.15, a.z),
          new THREE.Vector3(b.x, 0.15, b.z),
        ]);
        const dashed = !(s.aIsPole && s.bIsPole);
        const line = new THREE.Line(g, dashed ? matOther : matPolePole);
        if (dashed && line.computeLineDistances) line.computeLineDistances();
        threeState.groups.spans.add(line);
      }
    }

    // Guying (existing + proposed) / messengers (3D)
    // Rendered as gray linework (requested). Down guys are drawn from the pole to
    // the ground anchor. Pole-to-pole guying is drawn between poles.
    {
      const guys = Array.isArray(model.guyLines) ? model.guyLines : [];
      if (guys.length && threeState.groups.guying) {
        const guyMatExisting = new THREE.MeshStandardMaterial({
          color: 0x9ca3af,
          emissive: 0x9ca3af,
          emissiveIntensity: 0.05,
          roughness: 0.60,
          metalness: 0.02,
          transparent: true,
          opacity: 0.46,
          depthWrite: false,
        });

        // Proposed guying/downguying: gray linework.
        // Visual rule (3D): ONLY the selected installing company's proposed items should flash.
        const guyMatProposedStatic = new THREE.MeshStandardMaterial({
          color: 0x9ca3af,
          emissive: 0x9ca3af,
          emissiveIntensity: 0.08,
          roughness: 0.52,
          metalness: 0.02,
          transparent: true,
          opacity: 0.72,
          depthWrite: false,
        });

        const guyMatProposedPulse = new THREE.MeshStandardMaterial({
          color: 0x9ca3af,
          emissive: 0x9ca3af,
          emissiveIntensity: 0.10,
          roughness: 0.52,
          metalness: 0.02,
          transparent: true,
          opacity: 0.72,
          depthWrite: false,
        });

        // Flash linework ONLY for the selected installing company (3D only).
        // Make the flash OBVIOUS (requested): opacity pulses 25% → 100%.
        if (installingCompanyKey) {
          guyMatProposedPulse.userData.pulse = { baseOpacity: 0.25, ampOpacity: 0.75, baseEmissiveIntensity: 0.06, ampEmissiveIntensity: 0.10 };
          threeState.installPulseMaterials.push(guyMatProposedPulse);
        }

        for (const g0 of guys) {
          if (!g0) continue;
          const poleId = (g0.poleId != null) ? String(g0.poleId) : "";
          if (!poleId) continue;
          const pole = poleById.get(poleId);
          if (!pole || pole.lat == null || pole.lon == null) continue;

          const poleXZ = latLonToXZFt(Number(pole.lat), Number(pole.lon));
          const anchorId = (g0.anchorId != null) ? String(g0.anchorId) : "";
          const anchorType = String(g0.anchorType || "").toLowerCase();
          const isPoleToPole = anchorType === "pole";

          let anchorXZ = null;
          if (g0.anchorLat != null && g0.anchorLon != null) {
            anchorXZ = latLonToXZFt(Number(g0.anchorLat), Number(g0.anchorLon));
          } else if (anchorId && poleById.has(anchorId)) {
            const ap = poleById.get(anchorId);
            if (ap && ap.lat != null && ap.lon != null) anchorXZ = latLonToXZFt(Number(ap.lat), Number(ap.lon));
          }

          if (!anchorXZ) {
            // Fallback: deterministic offset so the downguys are still visible.
            const b = hashU32(`${poleId}|${anchorId}|${g0.traceId || ""}`) % 360;
            const dir = bearingVecXZ(b);
            const dFt = 28;
            anchorXZ = { x: poleXZ.x + dir.x * dFt, z: poleXZ.z + dir.z * dFt };
          }

          const exIn = (g0.existingIn != null && Number.isFinite(Number(g0.existingIn))) ? Number(g0.existingIn) : null;
          const prIn = (g0.proposedIn != null && Number.isFinite(Number(g0.proposedIn))) ? Number(g0.proposedIn) : null;

          const guyOwner = (g0.owner != null) ? String(g0.owner) : "";
          const guyOwnerKey = normalizeOwnerKey(guyOwner);
          const isInstallGuy = !!(installingCompanyKey && guyOwnerKey === installingCompanyKey);

          // Proposed detection for guying/downguys:
          // - Katapult may mark proposed facilities at the trace level (trace.proposed).
          // - Some datasets also include an explicit guying_type label containing "proposed".
          // We treat either as proposed so new/proposed downguying reliably flashes in 3D.
          const guyingTypeLower = String(g0.guyingType || "").toLowerCase();
          const isProposedTagged = !!(g0 && g0.traceProposed) || guyingTypeLower.includes("proposed");

          const hasExisting = (exIn != null) && !isProposedTagged;
          const hasProposed = (prIn != null) && (isProposedTagged || !hasExisting || (exIn != null && Math.abs(prIn - exIn) >= 1));
          if (!hasExisting && !hasProposed) continue;

          const yExisting = hasExisting ? (exIn / 12) : null;
          const yProposed = hasProposed ? (prIn / 12) : null;

          // Only highlight/flash for the selected installing company when the facility itself is
          // explicitly proposed (NOT merely because a "proposed height" exists).
          const proposedMat = (isInstallGuy && isProposedTagged) ? guyMatProposedPulse : guyMatProposedStatic;
          const isProposedFlashing = !!(isInstallGuy && isProposedTagged && yProposed != null);

          const groundY = 0.15;
          const endYExisting = isPoleToPole ? (yExisting != null ? yExisting : groundY) : groundY;
          const endYProposed = isPoleToPole ? (yProposed != null ? yProposed : groundY) : groundY;

          const entType = isPoleToPole ? "guy" : "downguy";
          const traceId = (g0.traceId != null) ? String(g0.traceId) : "";
          const exKey = (exIn != null) ? String(Math.round(exIn)) : "na";
          const prKey = (prIn != null) ? String(Math.round(prIn)) : "na";
          const baseKey = `${poleId}|${anchorType}|${anchorId}|${traceId}|${exKey}|${prKey}`;

          if (yExisting != null) {
            const p1 = new THREE.Vector3(poleXZ.x, yExisting, poleXZ.z);
            const p2 = new THREE.Vector3(anchorXZ.x, endYExisting, anchorXZ.z);
            const seg = cylinderBetween(p1, p2, 0.26, guyMatExisting);
            if (seg) {
              const segId = `${baseKey}|existing`;
              seg.userData = { entityType: entType, entityId: segId };
              threeState.guyPickIndex.set(segId, {
                type: entType,
                variant: "existing",
                poleId,
                anchorType,
                anchorId,
                traceId,
                owner: guyOwner,
                isFlashingProposed: false,
                existingIn: exIn,
                proposedIn: prIn,
                isPoleToPole,
              });
              threeState.groups.guying.add(seg);
              threeState.pickables.push(seg);
            }
          }
          if (yProposed != null) {
            const p1 = new THREE.Vector3(poleXZ.x, yProposed, poleXZ.z);
            const p2 = new THREE.Vector3(anchorXZ.x, endYProposed, anchorXZ.z);
            const seg = cylinderBetween(p1, p2, 0.30, proposedMat);
            if (seg) {
              const segId = `${baseKey}|proposed`;
              seg.userData = { entityType: entType, entityId: segId };
              threeState.guyPickIndex.set(segId, {
                type: entType,
                variant: "proposed",
                poleId,
                anchorType,
                anchorId,
                traceId,
                owner: guyOwner,
                isFlashingProposed: !!isProposedFlashing,
                existingIn: exIn,
                proposedIn: prIn,
                isPoleToPole,
              });
              threeState.groups.guying.add(seg);
              threeState.pickables.push(seg);
            }
          }
        }
      }
    }

    // Poles
    {
      const poleRadiusBottom = 0.55;
      const poleRadiusTop = 0.38;

      // Equipment (3D): transformer cans + streetlights.
      // Purple @ ~50% opacity (requested).
      const equipmentColor = 0x8b5cf6;
      const equipmentMat = new THREE.MeshStandardMaterial({
        color: equipmentColor,
        emissive: equipmentColor,
        emissiveIntensity: 0.16,
        roughness: 0.32,
        metalness: 0.08,
        transparent: true,
        opacity: 0.50,
      });
      equipmentMat.depthWrite = false;

      const equipmentPanelMat = new THREE.MeshBasicMaterial({
        color: equipmentColor,
        transparent: true,
        opacity: 0.18,
        blending: THREE.AdditiveBlending,
        side: THREE.DoubleSide,
        depthWrite: false,
      });

      const parsePoleHeightFt = (p) => {
        const raw = String(p.proposedPoleSpec || p.poleSpec || p.poleHeightClass || "").trim();
        const m = raw.match(/(\d+(?:\.\d+)?)/);
        if (m) {
          const n = Number(m[1]);
          if (Number.isFinite(n) && n > 10 && n < 200) return n;
        }
        let maxIn = 0;
        for (const a of (p.attachments || [])) {
          const h = (a && (a.proposedIn != null ? a.proposedIn : a.existingIn));
          if (h != null && Number(h) > maxIn) maxIn = Number(h);
        }
        if (maxIn > 0) return (maxIn / 12) + 8;
        return 40;
      };

      for (const p of (model.poles || [])) {
        if (!p || typeof p.lat !== "number" || typeof p.lon !== "number") continue;
        const pos = latLonToXZFt(p.lat, p.lon);

        // Preserve a deterministic focus point for 3D framing: the first pole in the
        // dataset with valid coordinates.
        if (!threeState.firstNode) {
          threeState.firstNode = new THREE.Vector3(pos.x, 0.15, pos.z);
        }
        const poleId = String(p.poleId);

        // Anchor for camera pivot snapping (poles + midspans).
        // NOTE: midspans should *not* be added to poleAnchors; poleAnchors are reserved for
        // pole-only interactions like double-click "focus nearest pole".
        if (threeState.nodeAnchors) threeState.nodeAnchors.push({ x: pos.x, z: pos.z });

        const res = qcResults && qcResults.poles ? qcResults.poles[poleId] : null;
        const st = res && res.status ? res.status : "unknown";
        const hasOrder = !!(res && res.hasCommOrderIssue);

        const hFt = parsePoleHeightFt(p);
        const geom = new THREE.CylinderGeometry(poleRadiusTop, poleRadiusBottom, hFt, 8, 1);
        const cSt = statusColor(st);
        const emiss = (st === "fail") ? 0.32 : (st === "warn") ? 0.24 : (st === "pass") ? 0.12 : 0.10;
        const mat = new THREE.MeshStandardMaterial({
          color: cSt,
          emissive: cSt,
          emissiveIntensity: emiss,
          roughness: 0.63,
          metalness: 0.06,
        });
        const mesh = new THREE.Mesh(geom, mat);
        mesh.position.set(pos.x, hFt / 2, pos.z);
        mesh.userData = { entityType: "pole", entityId: poleId };

        const group = new THREE.Group();
        group.add(mesh);

        // Equipment (3D): transformers + streetlights.
        // Uses equipment rows already present in pole attachments (extracted from the job JSON).
        try {
          const equipAtts = (p.attachments || []).filter(a => a && String(a.category || "").toLowerCase() === "equipment");
          if (equipAtts.length) {
            const transformersByTrace = new Map();
            const streetLightsByTrace = new Map();

            for (const a of equipAtts) {
              const hIn = (a.proposedIn != null) ? a.proposedIn : a.existingIn;
              if (hIn == null || !Number.isFinite(Number(hIn))) continue;
              const hFtEq = Number(hIn) / 12;

              const clsEq = classify(a);
              if (!clsEq.isTransformer && !clsEq.isStreetLight) continue;

              const traceId = (a.traceId != null) ? String(a.traceId) : (a.id != null ? `equip_${a.id}` : `equip_${Math.random()}`);
              const pack = {
                hFt: hFtEq,
                label: String(a.label || ""),
                traceLabel: String(a.traceLabel || ""),
                comments: String(a.comments || ""),
              };

              const bucket = clsEq.isTransformer ? transformersByTrace : streetLightsByTrace;
              if (!bucket.has(traceId)) bucket.set(traceId, []);
              bucket.get(traceId).push(pack);
            }

            const pickBearingDeg = (seedText, fallbackSeed) => {
              const b = parseBearingDeg(seedText);
              if (b != null) return b;
              return (hashU32(fallbackSeed) % 360);
            };

            // Transformer cans
            for (const [traceId, pts] of transformersByTrace.entries()) {
              const hs = pts.map(p => p.hFt).filter(h => Number.isFinite(h));
              if (!hs.length) continue;
              let topFt = Math.max(...hs);
              let bottomFt = Math.min(...hs);
              if (!Number.isFinite(topFt) || !Number.isFinite(bottomFt)) continue;

              // If only one point is provided, assume a reasonable can height.
              if ((topFt - bottomFt) < 1.25) bottomFt = topFt - 4.0;
              bottomFt = Math.max(2.0, bottomFt);
              const canH = Math.max(2.5, topFt - bottomFt);
              const canR = 1.25;

              const seedText = pts.map(p => p.label || p.traceLabel || p.comments).join(" ");
              const bDeg = pickBearingDeg(seedText, `${poleId}|${traceId}|transformer`);
              const dir = bearingVecXZ(bDeg).normalize();
              const offset = poleRadiusBottom + canR + 0.25;

              const canCenter = new THREE.Vector3(
                pos.x + dir.x * offset,
                (topFt + bottomFt) / 2,
                pos.z + dir.z * offset
              );
              const canG = new THREE.CylinderGeometry(canR, canR, canH, 16, 1);
              const can = new THREE.Mesh(canG, equipmentMat);
              can.position.copy(canCenter);
              group.add(can);

              // Top + bottom brackets
              for (const y of [topFt, bottomFt]) {
                const polePt = new THREE.Vector3(pos.x, y, pos.z).addScaledVector(dir, poleRadiusBottom * 0.95);
                const canPt = new THREE.Vector3(canCenter.x, y, canCenter.z).addScaledVector(dir, -canR * 0.95);
                const br = cylinderBetween(polePt, canPt, 0.11, equipmentMat);
                if (br) group.add(br);
              }

              // Direction tick (keeps bearing readable without the large exterior disk).
              const tickStart = canCenter.clone().addScaledVector(dir, canR * 0.98);
              const tickEnd = tickStart.clone().addScaledVector(dir, 1.15);
              const tick = cylinderBetween(tickStart, tickEnd, 0.07, equipmentMat);
              if (tick) group.add(tick);

              const tipG = new THREE.SphereGeometry(0.18, 12, 10);
              const tip = new THREE.Mesh(tipG, equipmentMat);
              tip.position.copy(tickEnd);
              group.add(tip);
            }

            // Streetlights
            for (const [traceId, pts] of streetLightsByTrace.entries()) {
              const hs = pts.map(p => p.hFt).filter(h => Number.isFinite(h));
              if (!hs.length) continue;
              let topFt = Math.max(...hs);
              let bottomFt = Math.min(...hs);
              if (!Number.isFinite(topFt) || !Number.isFinite(bottomFt)) continue;
              if ((topFt - bottomFt) < 0.60) bottomFt = topFt - 1.0;
              bottomFt = Math.max(1.5, bottomFt);

              const seedText = pts.map(p => p.label || p.traceLabel || p.comments).join(" ");
              const bDeg = pickBearingDeg(seedText, `${poleId}|${traceId}|streetlight`);
              const dir = bearingVecXZ(bDeg).normalize();

              // Bracket/arm length (feet). We parse the first "X ft" value from the label.
              let armFt = parseFirstFeet(seedText);
              if (!Number.isFinite(armFt)) armFt = 5;
              armFt = clamp(armFt, 2.5, 25);

              const armEndXZ = new THREE.Vector3(pos.x, 0, pos.z).addScaledVector(dir, poleRadiusBottom + armFt);

              const poleTop = new THREE.Vector3(pos.x, topFt, pos.z).addScaledVector(dir, poleRadiusBottom * 0.95);
              const poleBot = new THREE.Vector3(pos.x, bottomFt, pos.z).addScaledVector(dir, poleRadiusBottom * 0.95);
              const endTop = new THREE.Vector3(armEndXZ.x, topFt, armEndXZ.z);
              const endBot = new THREE.Vector3(armEndXZ.x, bottomFt, armEndXZ.z);

              const topArm = cylinderBetween(poleTop, endTop, 0.12, equipmentMat);
              const botArm = cylinderBetween(poleBot, endBot, 0.12, equipmentMat);
              const tie = cylinderBetween(endBot, endTop, 0.10, equipmentMat);
              if (topArm) group.add(topArm);
              if (botArm) group.add(botArm);
              if (tie) group.add(tie);

              // Lamp head (simple + readable)
              const headG = new THREE.ConeGeometry(0.55, 1.15, 18);
              const head = new THREE.Mesh(headG, equipmentMat);
              head.position.set(armEndXZ.x, bottomFt - 0.15, armEndXZ.z);
              head.rotation.x = Math.PI;
              group.add(head);

              const lensG = new THREE.CircleGeometry(0.70, 18);
              const lens = new THREE.Mesh(lensG, equipmentPanelMat);
              lens.position.set(armEndXZ.x, bottomFt - 0.58, armEndXZ.z);
              lens.rotation.x = Math.PI / 2;
              group.add(lens);
            }
          }
        } catch (_) {}

        // Status beacon (3D): makes FAIL/WARN poles visually unmistakable.
        // Futuristic, high-contrast indicators that stay readable at wide extents.
        if (st === "fail" || st === "warn") {
          const beaconOpacity = (st === "fail") ? 0.42 : 0.34;
          const beamOpacity = (st === "fail") ? 0.12 : 0.10;

          // Wider + thicker ground halo for clearer PASS/FAIL/WARN identification in 3D.
          const haloG = new THREE.TorusGeometry(5.25, 0.36, 12, 64);
          const haloM = new THREE.MeshBasicMaterial({
            color: cSt,
            transparent: true,
            opacity: beaconOpacity,
            blending: THREE.AdditiveBlending,
            depthWrite: false,
          });
          const halo = new THREE.Mesh(haloG, haloM);
          halo.rotation.x = Math.PI / 2;
          halo.position.set(pos.x, 0.85, pos.z);
          group.add(halo);

          const beamH = Math.min(120, hFt + 28);
          // Slightly wider "holographic" beacon column for better readability at a glance.
          const beamG = new THREE.CylinderGeometry(2.10, 2.10, beamH, 16, 1, true);
          const beamM = new THREE.MeshBasicMaterial({
            color: cSt,
            transparent: true,
            opacity: beamOpacity,
            blending: THREE.AdditiveBlending,
            side: THREE.DoubleSide,
            depthWrite: false,
          });
          const beam = new THREE.Mesh(beamG, beamM);
          beam.position.set(pos.x, beamH / 2, pos.z);
          group.add(beam);
        }

	        // Yellow ring around nodes with comm-order issues (keeps parity with 2D halo behavior).
        // Per user request: circle only (no yellow hologram/beam).
        if (hasOrder) {
          // Thick, high-contrast ring around comm arrangement issues.
          const ringG = new THREE.TorusGeometry(4.25, 0.48, 14, 64);
          const ringM = new THREE.MeshBasicMaterial({
            color: 0xfacc15,
            transparent: true,
            opacity: 0.82,
            blending: THREE.AdditiveBlending,
            depthWrite: false,
          });
          const ring = new THREE.Mesh(ringG, ringM);
          ring.rotation.x = Math.PI / 2;
          ring.position.set(pos.x, 1.35, pos.z);
          group.add(ring);
        }


        // Highlight attachments implicated by FAIL/WARN only.
        // (Comm arrangement issues are rendered via the yellow node ring, not beads.)
        for (const a of (p.attachments || [])) {
          if (!a || a.id == null) continue;
          const id = String(a.id);
          const hIn = (a.proposedIn != null ? a.proposedIn : a.existingIn);
          if (hIn == null) continue;

          const isFail = issueSets.failAttachmentIds.has(id);
          const isWarn = issueSets.warnAttachmentIds.has(id);
          if (!isFail && !isWarn) continue;

          const cls = classify(a);
          const isPower = typeof cls.kind === "string" && (cls.kind.startsWith("power_") || cls.kind === "power_other");
          const sev = isFail ? "fail" : "warn";
          const color = (sev === "fail")
            ? (isPower ? 0xb91c1c : 0xf87171)
            : (isPower ? 0xd97706 : 0xfbbf24);

          const y = Number(hIn) / 12;
          const sphG = new THREE.SphereGeometry(0.78, 14, 10);
          const sphM = new THREE.MeshStandardMaterial({
            color,
            emissive: color,
            emissiveIntensity: (sev === "fail") ? 0.32 : 0.22,
            roughness: 0.35,
          });
          const sph = new THREE.Mesh(sphG, sphM);
          sph.position.set(pos.x, y, pos.z);
          group.add(sph);

        }

        // Installing Company marker bead (3D only): green dull flash.
        if (installPulseBeadMat && installingCompanyKey) {
          const seenY = new Set();
          for (const a of (p.attachments || [])) {
            if (!a) continue;
            if (String(a.category || "").toLowerCase() !== "wire") continue;
            const hIn = (a.proposedIn != null ? a.proposedIn : a.existingIn);
            if (hIn == null) continue;
            const id = (a.id != null) ? String(a.id) : "";
            if (id && (issueSets.failAttachmentIds.has(id) || issueSets.warnAttachmentIds.has(id))) continue;
            const cls = classify(a);
            if (normalizeOwnerKey(cls.owner) !== installingCompanyKey) continue;
            const y = Number(hIn) / 12;
            const yKey = String(Math.round(y * 4) / 4);
            if (seenY.has(yKey)) continue;
            seenY.add(yKey);

            const bG = new THREE.SphereGeometry(0.72, 14, 10);
            const bead = new THREE.Mesh(bG, installPulseBeadMat);
            bead.position.set(pos.x, y, pos.z);
            group.add(bead);
          }
        }

        group.position.set(0, 0, 0);
        threeState.groups.poles.add(group);
        threeState.poleGroupsById.set(poleId, group);
        threeState.pickables.push(mesh);


        // 3D SCID label (rendered via HTML overlay so it stays crisp).
        try {
          const scid = (p.scid != null) ? String(p.scid).trim() : "";
          if (scid && threeState.labelLayer) {
            const lbl = document.createElement("div");
            const cls = ["three-scid-label"];
            if (st === "fail") cls.push("three-scid-label--fail");
            else if (st === "warn") cls.push("three-scid-label--warn");
            else if (st === "pass") cls.push("three-scid-label--pass");
            if (hasOrder) cls.push("three-scid-label--order");
            lbl.className = cls.join(" ");
            lbl.textContent = scid;
            lbl.title = `${scid} • ${st.toUpperCase()}${hasOrder ? " • COMM ORDER" : ""}`;
            threeState.labelLayer.appendChild(lbl);
            const worldPos = new THREE.Vector3(pos.x, hFt + 2.6, pos.z);
            threeState.scidLabels3d.set(poleId, { el: lbl, poleGroup: group, worldPos });
          }
        } catch (_) {}
      }
    }

    // Midspans + measured wires
    {
      const isPowerKind = (k) => {
        return typeof k === "string" && (k.startsWith("power_") || k === "power_other");
      };

      // Wire colors: warnings orange, failures red.
      // Power (pass) is rendered in futuristic purple (unless in violation).
      const wirePalette = {
        pass: { power: 0x8b5cf6, comm: 0x4ade80, other: 0x4ade80 },
        warn: { power: 0xd97706, comm: 0xfbbf24, other: 0xfbbf24 },
        fail: { power: 0xb91c1c, comm: 0xf87171, other: 0xf87171 },
        unknown: { power: 0x64748b, comm: 0x94a3b8, other: 0x94a3b8 },
      };

      const wireColor = (sev, kind) => {
        const s = wirePalette[sev] ? sev : "unknown";
        const k = (kind === "power" || kind === "comm" || kind === "other") ? kind : "other";
        return wirePalette[s][k] || wirePalette.unknown.other;
      };

      const makeWire = (pts, color, radius, withOrderHalo, sev, pulseMat, proposedPulse) => {
        const curve = new THREE.CatmullRomCurve3(pts);
        const geom = new THREE.TubeGeometry(curve, 10, radius, 6, false);

        const emiss = (sev === "fail") ? 0.34 : (sev === "warn") ? 0.26 : 0.16;
        const baseMat = new THREE.MeshStandardMaterial({
          color,
          emissive: color,
          emissiveIntensity: emiss,
          roughness: 0.36,
          metalness: 0.03,
        });
        // Proposed items (3D): flash the wire's native color (requested).
        if (proposedPulse) {
          const baseEI = emiss * 0.55;
          const ampEI = emiss * 0.85;
          baseMat.emissiveIntensity = baseEI;
          // Make the flash OBVIOUS (requested): opacity pulses 25% → 100%.
          baseMat.transparent = true;
          baseMat.opacity = 1.0;
          baseMat.depthWrite = false;
          baseMat.userData.pulse = { baseOpacity: 0.25, ampOpacity: 0.75, baseEmissiveIntensity: baseEI, ampEmissiveIntensity: ampEI };
          threeState.installPulseMaterials.push(baseMat);
        }

        const inner = new THREE.Mesh(geom, baseMat);

        const g = new THREE.Group();
        g.add(inner);

        // Installing-company highlighting: keep the wire's native color, add a pulsing green overlay.
        if (pulseMat) {
          const outerGeom = new THREE.TubeGeometry(curve, 10, radius + 0.12, 8, false);
          const outer = new THREE.Mesh(outerGeom, pulseMat);
          outer.renderOrder = 6;
          g.add(outer);
        }

        // Comm arrangement issues are shown by the yellow ring/triangle on the pole/midspan marker.
        // No additional yellow "hologram" is drawn around the wire itself.
        return g;
      };

      for (const ms of (model.midspanPoints || [])) {
        if (!ms || typeof ms.lat !== "number" || typeof ms.lon !== "number") continue;
        const msId = String(ms.midspanId);
        const pos = latLonToXZFt(ms.lat, ms.lon);

        // Anchor for camera pivot snapping (poles + midspans).
        if (threeState.nodeAnchors) threeState.nodeAnchors.push({ x: pos.x, z: pos.z });

        // Fallback focus point when there are no poles in the model.
        if (!threeState.firstNode) {
          threeState.firstNode = new THREE.Vector3(pos.x, 0.15, pos.z);
        }

        const res = qcResults && qcResults.midspans ? qcResults.midspans[msId] : null;
        const st = res && res.status ? res.status : "unknown";
        const hasOrder = !!(res && res.hasCommOrderIssue);

        const g = new THREE.Group();

        // Midspan marker
        const coneG = new THREE.ConeGeometry(1.15, 3.0, 8);
        const coneM = new THREE.MeshStandardMaterial({ color: statusColor(st), roughness: 0.55, metalness: 0.06 });
        const cone = new THREE.Mesh(coneG, coneM);
        cone.position.set(pos.x, 1.6, pos.z);
        cone.userData = { entityType: "midspan", entityId: msId };
        g.add(cone);
        threeState.pickables.push(cone);

	        if (hasOrder) {
	          // Comm arrangement issues (midspan): use a yellow TRIANGLE ring so it
	          // is immediately distinguishable from pole comm-order circles.
	          const triR = 4.75;
	          const triY = 1.05;
	          const a0 = Math.PI / 2;
	          const tp1 = new THREE.Vector3(pos.x + triR * Math.cos(a0), triY, pos.z + triR * Math.sin(a0));
	          const tp2 = new THREE.Vector3(pos.x + triR * Math.cos(a0 + (2 * Math.PI / 3)), triY, pos.z + triR * Math.sin(a0 + (2 * Math.PI / 3)));
	          const tp3 = new THREE.Vector3(pos.x + triR * Math.cos(a0 + (4 * Math.PI / 3)), triY, pos.z + triR * Math.sin(a0 + (4 * Math.PI / 3)));
	          const triPath = new THREE.CurvePath();
	          triPath.add(new THREE.LineCurve3(tp1, tp2));
	          triPath.add(new THREE.LineCurve3(tp2, tp3));
	          triPath.add(new THREE.LineCurve3(tp3, tp1));
	          const ringG = new THREE.TubeGeometry(triPath, 60, 0.48, 10, false);
	          const ringM = new THREE.MeshBasicMaterial({
	            color: 0xfacc15,
	            transparent: true,
	            opacity: 0.82,
	            blending: THREE.AdditiveBlending,
	            depthWrite: false,
	          });
	          const ring = new THREE.Mesh(ringG, ringM);
	          g.add(ring);
        }

        const conn = (ms.connectionId != null) ? String(ms.connectionId) : "";
        const span = conn ? spanByConnection.get(conn) : null;
        const aPos = (span && span.aLat != null && span.aLon != null) ? latLonToXZFt(Number(span.aLat), Number(span.aLon)) : null;
        const bPos = (span && span.bLat != null && span.bLon != null) ? latLonToXZFt(Number(span.bLat), Number(span.bLon)) : null;
        const aPoleId = (span && span.aIsPole && span.aNodeId != null) ? String(span.aNodeId) : "";
        const bPoleId = (span && span.bIsPole && span.bNodeId != null) ? String(span.bNodeId) : "";

        const measures = Array.isArray(ms.measures) ? ms.measures : [];
        let maxWireFt = 0;
        for (const m0 of measures) {
          if (!m0) continue;
          const mId = (m0.id != null) ? String(m0.id) : "";
          const exIn = (m0.existingIn != null && Number.isFinite(Number(m0.existingIn))) ? Number(m0.existingIn) : null;
          const prIn = (m0.proposedIn != null && Number.isFinite(Number(m0.proposedIn))) ? Number(m0.proposedIn) : null;
          const hIn = (prIn != null ? prIn : exIn);
          if (hIn == null) continue;

          const traceId = (m0.traceId != null) ? String(m0.traceId) : "";
          const hMid = Number(hIn) / 12;
          if (Number.isFinite(hMid) && hMid > maxWireFt) maxWireFt = hMid;

          const aHIn = (aPoleId && traceId) ? poleWireHeightByPoleTrace.get(`${aPoleId}|${traceId}`) : null;
          const bHIn = (bPoleId && traceId) ? poleWireHeightByPoleTrace.get(`${bPoleId}|${traceId}`) : null;
          const hA = (aHIn != null ? Number(aHIn) / 12 : hMid);
          const hB = (bHIn != null ? Number(bHIn) / 12 : hMid);

          if (!aPos || !bPos) continue;

          const p1 = new THREE.Vector3(aPos.x, hA, aPos.z);
          const p2 = new THREE.Vector3(pos.x, hMid, pos.z);
          const p3 = new THREE.Vector3(bPos.x, hB, bPos.z);          const cls = classify({ ...m0, category: "Wire" });
          const isPower = isPowerKind(cls.kind);
          const kindKey = isPower ? "power" : (cls.kind === "comm" ? "comm" : "other");

          const hasFail = mId ? issueSets.failMeasureIds.has(mId) : false;
          const hasWarn = mId ? issueSets.warnMeasureIds.has(mId) : false;
          const hasOrderWire = mId ? issueSets.orderMeasureIds.has(mId) : false;

          const sev = hasFail ? "fail" : hasWarn ? "warn" : "pass";
          const withOrderHalo = hasOrderWire || (hasOrder && cls.kind === "comm");

          const ownerKey = normalizeOwnerKey(cls.owner);

          // A wire is considered "proposed" ONLY when Katapult flags the trace as proposed.
          // IMPORTANT: "proposed height" (existing vs proposedIn) is NOT used to determine
          // whether the facility itself is proposed.
          const isProposedWire = !!(m0 && m0.traceProposed);

          // Visual rule (3D): ONLY the selected installing company's proposed items should flash.
          const isProposedForInstall = !!(
            installingCompanyKey &&
            ownerKey === installingCompanyKey &&
            isProposedWire
          );

          const isInstallPulse = !!(
            installPulseWireMat &&
            installingCompanyKey &&
            sev === "pass" &&
            ownerKey === installingCompanyKey
          );

          const isMessenger = normalizeStr(`${m0.label || ""} ${m0.name || ""} ${m0.traceLabel || ""} ${m0.traceType || ""} ${m0.cableType || ""}`).includes("messenger");

          const color = (isMessenger && sev === "pass") ? 0x94a3b8 : wireColor(sev, kindKey);
          const radius = isPower ? 0.26 : 0.21;

          const wirePickId = mId || `${msId}|${traceId}|${String(m0.wireId || "")}|${Math.round(Number(hIn) || 0)}`;
          threeState.wirePickIndex.set(wirePickId, {
            type: "wire",
            measureId: mId || null,
            midspanId: msId,
            connectionId: conn || null,
            aPoleId: aPoleId || null,
            bPoleId: bPoleId || null,
            traceId: traceId || null,
            wireId: (m0.wireId != null) ? String(m0.wireId) : null,
            owner: cls.owner || null,
            label: String(m0.label || m0.name || cls.label || "").trim() || null,
            kind: cls.kind || null,
            existingIn: exIn,
            proposedIn: prIn,
            drawnIn: Number(hIn),
            isProposed: !!isProposedWire,
            isFlashingProposed: !!isProposedForInstall,
            severity: sev,
          });

          const tube = makeWire([p1, p2, p3], color, radius, withOrderHalo, sev, isInstallPulse ? installPulseWireMat : null, isProposedForInstall);
          // Make each wire selectable in 3D.
          try {
            tube.traverse((o) => {
              if (o && o.isMesh) o.userData = { entityType: "wire", entityId: wirePickId };
            });
            threeState.pickables.push(tube);
          } catch (_) {}
          g.add(tube);

          // Beads should only indicate actual WARN/FAIL locations (not comm arrangement).
          if (hasFail || hasWarn) {
            const dotG = new THREE.SphereGeometry(0.54, 14, 10);
            const dotM = new THREE.MeshStandardMaterial({
              color,
              emissive: color,
              emissiveIntensity: sev === "fail" ? 0.30 : sev === "warn" ? 0.22 : 0.14,
              roughness: 0.35,
            });
            const dot = new THREE.Mesh(dotG, dotM);
            dot.position.copy(p2);
            g.add(dot);
          }
        }

        // Status beacon for midspans (3D): mirrors pole readability.
        // Wider + thicker halo + larger "holographic" column per user request.
	        if (st === "fail" || st === "warn") {
          const cSt = statusColor(st);
          const beaconOpacity = (st === "fail") ? 0.42 : 0.34;
          const beamOpacity = (st === "fail") ? 0.12 : 0.10;

	          // Midspan hologram: use a triangular "halo" to distinguish midspans from poles.
	          const triR = 5.25;
	          const triY = 0.85;
	          const a0 = Math.PI / 2;
	          const p1 = new THREE.Vector3(pos.x + triR * Math.cos(a0), triY, pos.z + triR * Math.sin(a0));
	          const p2 = new THREE.Vector3(pos.x + triR * Math.cos(a0 + (2 * Math.PI / 3)), triY, pos.z + triR * Math.sin(a0 + (2 * Math.PI / 3)));
	          const p3 = new THREE.Vector3(pos.x + triR * Math.cos(a0 + (4 * Math.PI / 3)), triY, pos.z + triR * Math.sin(a0 + (4 * Math.PI / 3)));
	          const triPath = new THREE.CurvePath();
	          triPath.add(new THREE.LineCurve3(p1, p2));
	          triPath.add(new THREE.LineCurve3(p2, p3));
	          triPath.add(new THREE.LineCurve3(p3, p1));
	          const haloG = new THREE.TubeGeometry(triPath, 54, 0.36, 10, false);
          const haloM = new THREE.MeshBasicMaterial({
            color: cSt,
            transparent: true,
            opacity: beaconOpacity,
            blending: THREE.AdditiveBlending,
            depthWrite: false,
          });
          const halo = new THREE.Mesh(haloG, haloM);
	          // TubeGeometry path is already in the ground plane at triY.
          g.add(halo);

          const beamH = Math.min(120, Math.max(60, (Number.isFinite(maxWireFt) ? maxWireFt : 0) + 18));
	          // Triangular holographic column for midspans (easy pole vs midspan differentiation).
	          const beamG = new THREE.CylinderGeometry(2.10, 2.10, beamH, 3, 1, true);
          const beamM = new THREE.MeshBasicMaterial({
            color: cSt,
            transparent: true,
            opacity: beamOpacity,
            blending: THREE.AdditiveBlending,
            side: THREE.DoubleSide,
            depthWrite: false,
          });
          const beam = new THREE.Mesh(beamG, beamM);
          beam.position.set(pos.x, beamH / 2, pos.z);
          g.add(beam);
        }

	        // Comm arrangement issues: circle-only marker is rendered at the midspan node (no yellow hologram/beam).

        threeState.groups.midspans.add(g);
        threeState.midGroupsById.set(msId, g);
      }
    }

    // Endpoint-order reversal visualization (ORDER.COMM.ENDPOINTS)
    {
      const order = new Map();
      const issuesAll = qcResults && Array.isArray(qcResults.issues) ? qcResults.issues : [];
      for (const iss of issuesAll) {
        if (!iss || iss.entityType !== "pole") continue;
        if (String(iss.ruleCode || "") !== "ORDER.COMM.ENDPOINTS") continue;
        const poleId = String(iss.entityId || "");
        const ctx = iss.context || {};
        const otherPoleId = String(ctx.otherPoleId || "");
        const cid = String(ctx.connectionId || "");
        if (!poleId || !otherPoleId || !cid) continue;
        const key = `${cid}|${[poleId, otherPoleId].sort().join("|")}`;
        const e = order.get(key) || { cid, a: null, b: null };
        const rec = { poleId, otherPoleId, attachmentIds: Array.isArray(ctx.attachmentIds) ? ctx.attachmentIds.map(String) : [] };
        // assign deterministically
        if (!e.a || e.a.poleId === poleId) e.a = rec; else e.b = rec;
        order.set(key, e);
      }

      const normalizeOwner = (v) => String(v || "").toLowerCase().replace(/[^a-z0-9]/g, "");

      const makeTubeSpan = (pA, pB) => {
        const mid = pA.clone().add(pB).multiplyScalar(0.5);
        // tiny sag to emphasize crossing
        mid.y = (pA.y + pB.y) / 2 - 0.8;
        const curve = new THREE.CatmullRomCurve3([pA, mid, pB]);
        const geom = new THREE.TubeGeometry(curve, 20, 0.24, 6, false);
        const mat = new THREE.MeshStandardMaterial({ color: 0xef4444, emissive: 0xef4444, emissiveIntensity: 0.22, roughness: 0.38, metalness: 0.05 });
        const inner = new THREE.Mesh(geom, mat);

        const outerG = new THREE.TubeGeometry(curve, 20, 0.46, 7, false);
        const outerM = new THREE.MeshBasicMaterial({ color: 0xfacc15, transparent: true, opacity: 0.14, blending: THREE.AdditiveBlending, depthWrite: false });
        const outer = new THREE.Mesh(outerG, outerM);

        const g = new THREE.Group();
        g.add(inner);
        g.add(outer);
        return g;
      };

      for (const e of order.values()) {
        if (!e || !e.a || !e.b) continue;
        const poleA = String(e.a.poleId);
        const poleB = String(e.b.poleId);
        if (!poleA || !poleB) continue;

        const pA = poleById.get(poleA);
        const pB = poleById.get(poleB);
        if (!pA || !pB) continue;
        if (typeof pA.lat !== "number" || typeof pA.lon !== "number" || typeof pB.lat !== "number" || typeof pB.lon !== "number") continue;

        const aPos = latLonToXZFt(pA.lat, pA.lon);
        const bPos = latLonToXZFt(pB.lat, pB.lon);

        const ownerHeightsA = new Map();
        for (const id of (e.a.attachmentIds || [])) {
          const rec = attachmentById.get(String(id));
          if (!rec || !rec.att) continue;
          const cls = classify(rec.att);
          const ok = normalizeOwner(cls.owner);
          const hIn = (rec.att.proposedIn != null ? rec.att.proposedIn : rec.att.existingIn);
          if (!ok || hIn == null) continue;
          ownerHeightsA.set(ok, { owner: cls.owner || ok, y: Number(hIn) / 12 });
        }

        const ownerHeightsB = new Map();
        for (const id of (e.b.attachmentIds || [])) {
          const rec = attachmentById.get(String(id));
          if (!rec || !rec.att) continue;
          const cls = classify(rec.att);
          const ok = normalizeOwner(cls.owner);
          const hIn = (rec.att.proposedIn != null ? rec.att.proposedIn : rec.att.existingIn);
          if (!ok || hIn == null) continue;
          ownerHeightsB.set(ok, { owner: cls.owner || ok, y: Number(hIn) / 12 });
        }

        for (const ok of ownerHeightsA.keys()) {
          if (!ownerHeightsB.has(ok)) continue;
          const a = ownerHeightsA.get(ok);
          const b = ownerHeightsB.get(ok);
          const p1 = new THREE.Vector3(aPos.x, a.y, aPos.z);
          const p2 = new THREE.Vector3(bPos.x, b.y, bPos.z);
          const t = makeTubeSpan(p1, p2);
          threeState.groups.issues.add(t);
        }
      }
    }

    // Compute bounds (for Zoom to All)
    {
      const box = new THREE.Box3();
      box.setFromObject(threeState.root);
      threeState.bounds = box;
    }

    threeDirty = false;
    threeZoomToAll();
  }

  function threeZoomToAll() {
    if (!threeState || !threeState.bounds) return;
    const box = threeState.bounds;
    if (!box || box.isEmpty()) return;

    const center = new THREE.Vector3();
    const size = new THREE.Vector3();
    box.getCenter(center);
    box.getSize(size);

    const maxDim = Math.max(size.x, size.y, size.z, 1);

    // Frame the entire job extents in view regardless of where the job sits on the map.
    // Use a bounding-sphere fit against both vertical + horizontal FOV.
    let dist = (size.length() * 0.5);
    try {
      const cam = threeState.camera;
      const aspect = (cam && cam.aspect) ? cam.aspect : 1;
      const vFov = (cam && cam.fov) ? THREE.MathUtils.degToRad(cam.fov) : (Math.PI / 3);
      const hFov = 2 * Math.atan(Math.tan(vFov / 2) * aspect);
      const r = size.length() * 0.5;
      const distV = r / Math.max(1e-6, Math.sin(vFov / 2));
      const distH = r / Math.max(1e-6, Math.sin(hFov / 2));
      dist = Math.max(distV, distH);
      if (!isFinite(dist) || dist <= 0) dist = maxDim;
      dist = Math.max(520, dist * 1.15);
    } catch (_) {
      dist = Math.max(520, maxDim * 0.75 + 400);
    }

    // Visualization-only: keep zoom feel consistent across job sizes.
    try {
      if (threeState && threeState.controls) {
        threeState.controls.zoomStepFt = Math.max(6, Math.min(260, maxDim * 0.02));
      }
    } catch (_) {}

    // Center the 3D camera on the job's overall extents.
    const focus = center;
    const dir = new THREE.Vector3(1, 0.62, 1).normalize();

    threeState.controls.target.copy(focus);
    threeState.camera.position.copy(focus).addScaledVector(dir, dist);
    // FreeCamControlsLite does not automatically aim the camera at the target.
    threeState.camera.lookAt(focus);
    threeState.camera.near = Math.max(0.1, dist / 2000);
    threeState.camera.far = Math.max(50000, dist * 6);
    threeState.camera.updateProjectionMatrix();
    threeState.controls.update();
  }

  // Keep camera clipping stable across both wide "Zoom to All" views and close-up
  // inspections. Without this, large initial far/near settings can make close-up
  // navigation feel increasingly "restricted" due to near-plane clipping.
  function tuneThreeCameraClipping() {
    if (!threeState || !threeState.camera || !threeState.controls) return;
    const cam = threeState.camera;
    const tgt = threeState.controls.target;
    if (!tgt || !tgt.isVector3) return;

    const d = cam.position.distanceTo(tgt);
    if (!Number.isFinite(d) || d <= 0) return;

    // Dynamic near plane: scales with distance but is capped to avoid clipping the basemap.
    const desiredNear = Math.max(0.02, Math.min(75, d / 2000));
    if (Math.abs(desiredNear - cam.near) > Math.max(0.02, cam.near * 0.15)) {
      cam.near = desiredNear;
      cam.updateProjectionMatrix();
    }
  }

  // Subtle pulse for "Installing Company" highlighting in 3D.
  // This is visualization-only and does NOT affect QC logic.
  function updateThreeInstallPulse(nowMs) {
    if (!threeState) return;
    const mats = Array.isArray(threeState.installPulseMaterials) ? threeState.installPulseMaterials : [];
    if (!mats.length) return;
    const t = (Number(nowMs) || 0) * 0.001;
    // Dull, non-distracting flash.
    const wave = 0.5 + 0.5 * Math.sin(t * 2.2);

    for (const m of mats) {
      if (!m || !m.userData || !m.userData.pulse) continue;
      const p = m.userData.pulse;
      if (typeof p.baseEmissiveIntensity === "number" && typeof p.ampEmissiveIntensity === "number") {
        m.emissiveIntensity = p.baseEmissiveIntensity + p.ampEmissiveIntensity * wave;
      }
      if (m.transparent && typeof p.baseOpacity === "number" && typeof p.ampOpacity === "number") {
        m.opacity = p.baseOpacity + p.ampOpacity * wave;
      }
    }
  }

  function startThreeLoop() {
    if (!threeState || threeState.animationId) return;
    let last = performance.now();
    const tick = (now) => {
      if (!threeState) return;
      threeState.animationId = requestAnimationFrame(tick);

      const dt = Math.min(0.05, Math.max(0.001, (now - last) / 1000));
      last = now;

      // Camera controls (visualization only).
      try {
        if (threeState.controls && typeof threeState.controls.update === "function") {
          threeState.controls.update(dt);
        }
      } catch (_) {}

      // Dynamic clipping keeps close-up inspection usable even on very large job extents.
      try { tuneThreeCameraClipping(); } catch (_) {}

      // 3D SCID labels are HTML overlays that must be re-projected each frame.
      try { updateThreeScidLabels(); } catch (_) {}

      // Installing-company visual pulse (green dull flash).
      try { updateThreeInstallPulse(now); } catch (_) {}

      threeState.renderer.render(threeState.scene, threeState.camera);
    };
    threeState.animationId = requestAnimationFrame(tick);
  }

  function stopThreeLoop() {
    if (!threeState || !threeState.animationId) return;
    try { cancelAnimationFrame(threeState.animationId); } catch (_) {}
    threeState.animationId = null;
  }

  function disposeThree() {
    if (!threeState) return;
    try { stopThreeLoop(); } catch (_) {}
    try {
      if (threeState.controls && typeof threeState.controls.dispose === "function") {
        threeState.controls.dispose();
      }
    } catch (_) {}
    try {
      if (threeState.root) disposeThreeObject(threeState.root);
    } catch (_) {}
    try {
      if (threeState.renderer && typeof threeState.renderer.dispose === "function") threeState.renderer.dispose();
    } catch (_) {}
    try {
      if (threeState.container) threeState.container.innerHTML = "";
    } catch (_) {}
    threeState = null;
    threeDirty = false;
  }

  function resizeThreeRenderer() {
    if (!threeState) return;
    const c = threeState.container;
    if (!c) return;
    const w = c.clientWidth;
    const h = c.clientHeight;
    if (!w || !h) return;
    threeState.renderer.setSize(w, h, false);
    threeState.camera.aspect = w / h;
    threeState.camera.updateProjectionMatrix();
  }

  function refreshThreeVisibility() {
    if (!threeState || !threeState.groups) return;
    const allowed = allowedStatusSet();

    const showPoles = !!(els.togglePoles && els.togglePoles.checked);
    const showMidspans = !!(els.toggleMidspans && els.toggleMidspans.checked);
    const showSpans = !!(els.toggleSpans && els.toggleSpans.checked);

    if (threeState.groups.spans) threeState.groups.spans.visible = showSpans;
    if (threeState.groups.guying) threeState.groups.guying.visible = showSpans;
    if (threeState.groups.poles) threeState.groups.poles.visible = showPoles;
    if (threeState.groups.midspans) threeState.groups.midspans.visible = showMidspans;
    if (threeState.groups.issues) threeState.groups.issues.visible = true;

    // Status filters apply to pole/midspan groups.
    if (showPoles) {
      for (const [poleId, g] of threeState.poleGroupsById.entries()) {
        const st = qcResults && qcResults.poles && qcResults.poles[poleId] ? qcResults.poles[poleId].status : "unknown";
        g.visible = allowed.has(st);
      }
    }

    if (showMidspans) {
      for (const [midId, g] of threeState.midGroupsById.entries()) {
        const st = qcResults && qcResults.midspans && qcResults.midspans[midId] ? qcResults.midspans[midId].status : "unknown";
        g.visible = allowed.has(st);
      }
    }
  }

  function updateThreeScidLabels() {
    if (!threeState || !threeState.labelLayer || !threeState.scidLabels3d) return;
    if (!window.THREE || !threeState.camera) return;

    const show = !!(els.toggleScidLabels && els.toggleScidLabels.checked) && (viewMode === "3d");
    const c = threeState.container;
    const w = c ? c.clientWidth : 0;
    const h = c ? c.clientHeight : 0;
    if (!w || !h) return;

    const cam = threeState.camera;
    const tmp = new THREE.Vector3();

    for (const [poleId, rec] of threeState.scidLabels3d.entries()) {
      if (!rec || !rec.el) continue;
      const el = rec.el;
      const poleGroup = rec.poleGroup;

      if (!show || !poleGroup || !poleGroup.visible) {
        el.style.display = "none";
        continue;
      }

      tmp.copy(rec.worldPos);
      tmp.project(cam);

      // Behind camera / clipped
      if (tmp.z < -1 || tmp.z > 1) {
        el.style.display = "none";
        continue;
      }

      const x = (tmp.x * 0.5 + 0.5) * w;
      const y = (-tmp.y * 0.5 + 0.5) * h;

      // Off screen padding
      if (x < -80 || x > w + 80 || y < -80 || y > h + 80) {
        el.style.display = "none";
        continue;
      }

      el.style.display = "block";
      el.style.transform = `translate(-50%, -100%) translate(${x}px, ${y}px)`;
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

    // Export/Import rule sets (local file).
    if (els.btnSaveRulesFile) {
      els.btnSaveRulesFile.addEventListener("click", () => exportRulesFile());
    }
    if (els.btnLoadRulesFile) {
      els.btnLoadRulesFile.addEventListener("click", () => {
        if (els.rulesFileInput) els.rulesFileInput.click();
      });
    }
    if (els.rulesFileInput) {
      els.rulesFileInput.addEventListener("change", async (e) => {
        const f = e && e.target && e.target.files && e.target.files[0] ? e.target.files[0] : null;
        // Allow picking the same file repeatedly.
        try { e.target.value = ""; } catch (_) {}
        if (!f) return;
        await importRulesFile(f);
      });
    }

    els.btnResetRules.addEventListener("click", () => {
      rules = structuredClone(DEFAULT_RULES);
      saveRules(rules);
      applyRulesToUi();
      if (model) recomputeQc();
    });

    els.btnZoomAll.addEventListener("click", () => {
      if (viewMode === "3d") {
        threeZoomToAll();
      } else {
        zoomToAll();
      }
    });

    // 2D/3D view toggle
    const handleView = () => {
      const want3d = !!(els.view3d && els.view3d.checked);
      setViewMode(want3d ? "3d" : "2d");
    };
    if (els.view2d) els.view2d.addEventListener("change", handleView);
    if (els.view3d) els.view3d.addEventListener("change", handleView);


    // 3D basemap selector (visualization only).
    if (els.basemap3d) {
      els.basemap3d.addEventListener("change", () => {
        setBasemap(els.basemap3d.value || "Dark");
      });
    }

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
	    if (els.searchScid) {
	      els.searchScid.addEventListener("keydown", (e) => {
	        if (e.key === "Enter") {
	          e.preventDefault();
	          searchAndFocusPoleByScid(els.searchScid.value || "");
	        }
	      });
	    }
	    if (els.searchPoleTag) {
	      els.searchPoleTag.addEventListener("keydown", (e) => {
	        if (e.key === "Enter") {
	          e.preventDefault();
	          searchAndFocusPoleByTag(els.searchPoleTag.value || "");
	        }
	      });
	    }

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
      els.rulePoleNeutralSecondaryBelowTransformer,
      els.rulePoleWarnMissingIds,

      els.ruleMidMinDefault,
      els.ruleMidMinPed,
      els.ruleMidMinHwy,
      els.ruleMidMinFarm,
      els.ruleMidMinRail,
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
    disposeThree();
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
    // Mark 3D scene dirty (visualization only).
    threeDirty = true;
    updateMarkerStyles();
    refreshLayers();
    renderIssuesTable();
    setSummaryCounts(qcResults.summary);

    // If the user is currently in 3D view, rebuild to reflect updated statuses/issues.
    if (viewMode === "3d" && threeState) {
      try {
        rebuildThreeScene();
      } catch (_) {}
    }

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

    // Transformer equipment (utility): used for optional neutral/secondary ordering checks.
    const isTransformer = /\btransformer\b/.test(s) || /\bxfmr\b/.test(s) || /\bxfrmr\b/.test(s) || /\bxformer\b/.test(s);

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
      isTransformer,
    };
  }

  // ────────────────────────────────────────────────────────────────────────────
  //  QC engine
  // ────────────────────────────────────────────────────────────────────────────

  // Deduplicate issues to prevent repeated counting/rendering when the same
  // condition is detected multiple times (e.g., duplicate tagged rows).
  function issueDedupeKey(iss) {
    const sev = String(iss && iss.severity || "");
    const et = String(iss && iss.entityType || "");
    const eid = String(iss && iss.entityId || "");
    const rc = String(iss && iss.ruleCode || "");
    const msg = String(iss && iss.message || "");
    const ctx = (iss && iss.context) ? iss.context : {};
    const aIds = Array.isArray(ctx.attachmentIds)
      ? Array.from(new Set(ctx.attachmentIds.map((v) => String(v)))).sort()
      : [];
    const mIds = Array.isArray(ctx.measureIds)
      ? Array.from(new Set(ctx.measureIds.map((v) => String(v)))).sort()
      : [];
    return `${sev}|${et}|${eid}|${rc}|${msg}|a:${aIds.join(",")}|m:${mIds.join(",")}`;
  }

  function dedupeIssuesList(list) {
    const out = [];
    const seen = new Set();
    for (const iss of (list || [])) {
      const k = issueDedupeKey(iss);
      if (seen.has(k)) continue;
      seen.add(k);
      out.push(iss);
    }
    return out;
  }

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

    // Ensure we do not repeat issues already found.
    for (const r of Object.values(poleResults)) {
      if (r && Array.isArray(r.issues)) r.issues = dedupeIssuesList(r.issues);
    }
    for (const r of Object.values(midResults)) {
      if (r && Array.isArray(r.issues)) r.issues = dedupeIssuesList(r.issues);
    }
    const issuesDeduped = dedupeIssuesList(issues);

    // Recompute statuses now that span-level issues may have been added.
    for (const r of Object.values(poleResults)) {
      r.status = deriveStatus(r.issues || []);
      r.hasCommOrderIssue = !!(r.issues || []).some((i) => String(i.ruleCode || "").startsWith("ORDER.COMM"));
    }
    for (const r of Object.values(midResults)) {
      r.status = deriveStatus(r.issues || []);
      r.hasCommOrderIssue = !!(r.issues || []).some((i) => String(i.ruleCode || "").startsWith("ORDER.COMM"));
    }

    const summary = summarizeResults(poleResults, midResults, issuesDeduped);

    return {
      poles: poleResults,
      midspans: midResults,
      issues: issuesDeduped,
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

      // If an attachment is moved/installed INTO an existing hole (including a hole
      // vacated by another moved attachment), it is a hole reuse event and should
      // NOT be treated as drilling a new hole for the keep-out/buffer rule.
      const existingHoleHeights = new Set();
      for (const s of stationary) {
        if (s && s.proposedIn != null) existingHoleHeights.add(Number(s.proposedIn));
      }
      for (const h of movedFromHeights) {
        if (h != null) existingHoleHeights.add(Number(h));
      }
      const movedOrNewNewHoles = movedOrNew.filter(a => !existingHoleHeights.has(Number(a.proposedIn)));

      // (a) vs moved-from holes
      if (movedFromHeights.length && movedOrNewNewHoles.length) {
        for (const a of movedOrNewNewHoles) {
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
      if (stationary.length && movedOrNewNewHoles.length) {
        for (const a of movedOrNewNewHoles) {
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
      if (movedOrNewNewHoles.length > 1) {
        for (const [a, b] of pairwise(movedOrNewNewHoles)) {
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

    // Neutral and Secondaries must be below transformers (optional; disabled by default).
    // IMPORTANT: When this rule is enabled, "below transformer" means below *all* transformer
    // points — i.e., below both the top and the bottom measurement points when present.
    if (rules.pole.enforceNeutralSecondaryBelowTransformer) {
      const transformers = atts.filter(a => String(a.category || "").toLowerCase() === "equipment" && a._cls.isTransformer && getH(a) != null);
      if (transformers.length) {
        const tHeights = transformers.map(getH).filter((v) => v != null);
        if (tHeights.length) {
          const tBottom = Math.min(...tHeights);
          const tTop = Math.max(...tHeights);

          const nsWires = atts.filter(a => (a._cls.kind === "power_neutral" || a._cls.kind === "power_secondary") && getH(a) != null);
          for (const w of nsWires) {
            const h = getH(w);
            if (h == null) continue;
            if (h >= tBottom) {
              const tIds = attachmentIdsOf(transformers);
              const wIds = attachmentIdsOf(w);
              const attIds = Array.from(new Set([...wIds, ...tIds]));

              const wLabel = (w.label || w.traceLabel || w.name || w.traceType || w.cableType || "wire");
              const tLabel = transformers[0] && (transformers[0].label || transformers[0].traceLabel || transformers[0].name || transformers[0].comments) ? (transformers[0].label || transformers[0].traceLabel || transformers[0].name || transformers[0].comments) : "transformer";

              const spanNote = (tTop !== tBottom) ? ` (top ${fmtFtIn(tTop)}, bottom ${fmtFtIn(tBottom)})` : "";
              issues.push(issue("FAIL", "pole", String(pole.poleId), poleName, "POLE.NS_BELOW_XFMR",
                `Power ordering violation: ${wLabel} at ${fmtFtIn(h)} must be below all parts of the ${tLabel}${spanNote}.`,
                { wireHeight: h, transformerTop: tTop, transformerBottom: tBottom, attachmentIds: attIds }));
            }
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
    let reqComm = requiredMidspanMinComm(rules, rowTypeNorm);

    // Driveway special-case (same as parking lots):
    // - "driveway" / "drive way" => 15' 6" minimum (including "commercial driveway")
    //
    // This only changes the ROW minimum used for midspan clearance checks.
    const rowText = String(rowTypeRaw || rowTypeNorm || "").toLowerCase();
    if (/drive\s*way/.test(rowText)) {
      // Force driveway to the parking-lot style minimum.
      reqComm = rules.midspan.minCommDefaultIn;
    }


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
        const ownerA = String(a.owner || "").trim();
        const ownerB = String(b.owner || "").trim();
        const ownersDiffer = (ownerA && ownerB && ownerA !== ownerB);
        const minReq = (installCo && (ownerA === installCo || ownerB === installCo)) ? installMin : baseMin;
        // NOTE: dh===0 is permitted ONLY when the two comm attachments belong to the same company.
        // If two DIFFERENT comm companies are at the exact same height, it must be flagged.
        if ((dh === 0 && ownersDiffer) || (dh !== 0 && dh < minReq)) {
          issues.push(issue("FAIL", "midspan", String(ms.midspanId), msName, "MIDSPAN.COMM_SEP",
            `Comm separation ${dh}" between "${ownerA || "Unknown"}" and "${ownerB || "Unknown"}" is below ${minReq}".`,
            { dh, minReq, ownerA: ownerA || a.owner, ownerB: ownerB || b.owner, h1: a.proposedIn, h2: b.proposedIn, measureIds: measureIdsOf(a, b) }));
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
    } else if (sel.type === "wire") {
      renderWireDetails(sel.id);
    } else if (sel.type === "guy" || sel.type === "downguy") {
      renderGuyDetails(sel.type, sel.id);
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

  function renderWireDetails(wirePickId) {
    const id = String(wirePickId || "");
    const info = (threeState && threeState.wirePickIndex && threeState.wirePickIndex.get) ? threeState.wirePickIndex.get(id) : null;
    if (!info) {
      els.details.innerHTML = `<div class="muted">Wire details not available.</div>`;
      return;
    }

    const st = String(info.severity || "unknown").toLowerCase();
    const title = escapeHtml(info.label || "Wire");
    const owner = escapeHtml(info.owner || "");
    const kind = escapeHtml(info.kind || "");
    const measureId = escapeHtml(info.measureId || "");
    const traceId = escapeHtml(info.traceId || "");
    const connId = escapeHtml(info.connectionId || "");
    const aPoleId = escapeHtml(info.aPoleId || "");
    const bPoleId = escapeHtml(info.bPoleId || "");

    const ex = (info.existingIn != null) ? fmtFtIn(info.existingIn) : "";
    const pr = (info.proposedIn != null) ? fmtFtIn(info.proposedIn) : "";
    const drawn = (info.drawnIn != null) ? fmtFtIn(info.drawnIn) : "";

    // Inline issue list for this specific measured wire (if the job provides a stable measure id).
    let issuesHtml = `<div class="muted">No issues.</div>`;
    try {
      const all = qcResults && Array.isArray(qcResults.issues) ? qcResults.issues : [];
      const midIssues = [];
      if (info.measureId) {
        for (const iss of all) {
          const ctx = iss && iss.context ? iss.context : {};
          const mIds = Array.isArray(ctx.measureIds) ? ctx.measureIds : [];
          if (mIds.map(String).includes(String(info.measureId))) midIssues.push(iss);
        }
      }
      if (midIssues.length) {
        issuesHtml = `<ul class="issue-list">${midIssues.map(i => `<li class="issue issue--${i.severity === "FAIL" ? "fail" : "warn"}">
            <div class="issue-code">${escapeHtml(i.ruleCode)}</div>
            <div class="issue-msg">${escapeHtml(i.message)}</div>
          </li>`).join("")}</ul>`;
      }
    } catch (_) {}

    els.details.innerHTML = `
      <div class="details-header">
        <div>
          <div class="details-title">${title}</div>
          <div class="muted">Wire${kind ? ` • ${kind}` : ""}</div>
        </div>
        <div><span class="badge badge--${st}">${escapeHtml(st.toUpperCase())}</span></div>
      </div>

      <h3 class="details-subtitle">Summary</h3>
      <div class="table-wrap">
        <table class="table table--small">
          <tbody>
            <tr><th>Owner</th><td>${owner || `<span class="muted">—</span>`}</td></tr>
            <tr><th>Midspan ID</th><td>${escapeHtml(info.midspanId || "")}</td></tr>
            <tr><th>Connection ID</th><td>${connId || `<span class="muted">—</span>`}</td></tr>
            <tr><th>Endpoints</th><td>${aPoleId && bPoleId ? `${aPoleId} ↔ ${bPoleId}` : `<span class="muted">—</span>`}</td></tr>
            <tr><th>Trace ID</th><td>${traceId || `<span class="muted">—</span>`}</td></tr>
            <tr><th>Measure ID</th><td>${measureId || `<span class="muted">—</span>`}</td></tr>
          </tbody>
        </table>
      </div>

      <h3 class="details-subtitle">Heights</h3>
      <div class="table-wrap">
        <table class="table table--small">
          <thead>
            <tr>
              <th>Existing</th>
              <th>Proposed</th>
              <th>Shown in 3D</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>${escapeHtml(ex || "—")}</td>
              <td>${escapeHtml(pr || "—")}</td>
              <td>${escapeHtml(drawn || "—")}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <h3 class="details-subtitle">Issues</h3>
      ${issuesHtml}
    `;
  }

  function renderGuyDetails(kind, guyPickId) {
    const id = String(guyPickId || "");
    const info = (threeState && threeState.guyPickIndex && threeState.guyPickIndex.get) ? threeState.guyPickIndex.get(id) : null;
    if (!info) {
      els.details.innerHTML = `<div class="muted">Guying details not available.</div>`;
      return;
    }

    const title = kind === "downguy" ? "Downguy" : "Guying";
    const ex = (info.existingIn != null) ? fmtFtIn(info.existingIn) : "";
    const pr = (info.proposedIn != null) ? fmtFtIn(info.proposedIn) : "";

    els.details.innerHTML = `
      <div class="details-header">
        <div>
          <div class="details-title">${escapeHtml(title)}</div>
          <div class="muted">Pole ID: ${escapeHtml(info.poleId || "")}${info.anchorId ? ` • Anchor: ${escapeHtml(info.anchorId)}` : ""}</div>
        </div>
        <div><span class="badge badge--unknown">${escapeHtml(String(info.variant || "").toUpperCase() || "LINE")}</span></div>
      </div>

      <h3 class="details-subtitle">Summary</h3>
      <div class="table-wrap">
        <table class="table table--small">
          <tbody>
            <tr><th>Type</th><td>${escapeHtml(title)}</td></tr>
            <tr><th>Variant</th><td>${escapeHtml(String(info.variant || "").toUpperCase())}</td></tr>
            <tr><th>Anchor type</th><td>${escapeHtml(String(info.anchorType || ""))}</td></tr>
            <tr><th>Trace ID</th><td>${escapeHtml(String(info.traceId || "")) || `<span class="muted">—</span>`}</td></tr>
          </tbody>
        </table>
      </div>

      <h3 class="details-subtitle">Heights</h3>
      <div class="table-wrap">
        <table class="table table--small">
          <thead>
            <tr>
              <th>Existing</th>
              <th>Proposed</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>${escapeHtml(ex || "—")}</td>
              <td>${escapeHtml(pr || "—")}</td>
            </tr>
          </tbody>
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

    // Midspans in the Issues tab should be labeled by their connected pole SCID(s)
    // rather than only the raw JSON midspan ID.
    const poleById = new Map();
    for (const p of (model && Array.isArray(model.poles) ? model.poles : [])) {
      if (!p) continue;
      if (p.poleId == null) continue;
      poleById.set(String(p.poleId), p);
    }

    const midById = new Map();
    for (const ms of (model && Array.isArray(model.midspanPoints) ? model.midspanPoints : [])) {
      if (!ms) continue;
      if (ms.midspanId == null) continue;
      midById.set(String(ms.midspanId), ms);
    }

    function poleScidByPoleId(poleId) {
      const key = String(poleId || "");
      const p = poleById.get(key);
      return String((p && (p.scid || p.displayName || p.poleTag)) || key);
    }

    function midspanEntityLabel(issue) {
      if (!issue) return "Midspan";
      const ms = midById.get(String(issue.entityId || ""));
      if (!ms) return String(issue.entityName || issue.entityId || "Midspan");

      const aPoleId = String(ms.aPoleId || "");
      const bPoleId = String(ms.bPoleId || "");

      const scids = [];
      if (aPoleId) scids.push(poleScidByPoleId(aPoleId));
      if (bPoleId && bPoleId !== aPoleId) scids.push(poleScidByPoleId(bPoleId));

      if (!scids.length) return String(issue.entityName || issue.entityId || "Midspan");
      if (scids.length === 1) return `Midspan ${scids[0]}`;
      return `Midspan ${scids[0]} - ${scids[1]}`;
    }

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
      const displayName = (iss.entityType === "pole")
        ? (iss.entityName || iss.entityId)
        : (iss.entityType === "midspan" ? midspanEntityLabel(iss) : (iss.entityName || "Midspan"));
      const name = escapeHtml(displayName);
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

    // CSV export requirements:
    // - Remove the "entityId" and "ruleCode" columns.
    // - If entityType is "midspan", set entityName to: "midspan <SCID-A> - <SCID-B>".
    //   If the midspan terminates at a reference point, use only the originating pole SCID.
    const cols = ["severity", "entityType", "entityName", "message"];
    const lines = [cols.join(",")];

    const poleById = new Map();
    for (const p of (model && Array.isArray(model.poles) ? model.poles : [])) {
      if (!p) continue;
      if (p.poleId == null) continue;
      poleById.set(String(p.poleId), p);
    }

    const midById = new Map();
    for (const ms of (model && Array.isArray(model.midspanPoints) ? model.midspanPoints : [])) {
      if (!ms) continue;
      if (ms.midspanId == null) continue;
      midById.set(String(ms.midspanId), ms);
    }

    function poleScidByPoleId(poleId) {
      const key = String(poleId || "");
      const p = poleById.get(key);
      return String((p && (p.scid || p.displayName || p.poleTag)) || key);
    }

    function csvEntityNameForIssue(i) {
      if (!i) return "";
      if (i.entityType !== "midspan") return i.entityName || "";

      const ms = midById.get(String(i.entityId || ""));
      if (!ms) return i.entityName || "midspan";

      const aPoleId = String(ms.aPoleId || "");
      const bPoleId = String(ms.bPoleId || "");

      const scids = [];
      if (aPoleId) scids.push(poleScidByPoleId(aPoleId));
      if (bPoleId && bPoleId !== aPoleId) scids.push(poleScidByPoleId(bPoleId));

      if (!scids.length) return "midspan";
      if (scids.length === 1) return `midspan ${scids[0]}`;
      return `midspan ${scids[0]} - ${scids[1]}`;
    }

    for (const i of qcResults.issues) {
      const row = [
        i.severity,
        i.entityType,
        csvEntityNameForIssue(i),
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
  //  Rules: Export / Import (local file)
  // ────────────────────────────────────────────────────────────────────────────

  function exportRulesFile() {
    try {
      const now = new Date();
      const pad2 = (n) => String(n).padStart(2, "0");
      const stamp = `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}_${pad2(now.getHours())}${pad2(now.getMinutes())}${pad2(now.getSeconds())}`;
      const base = safeFilename((model && model.jobName) ? model.jobName : (fileName || "rules"));

      const payload = {
        schema: "katapultQcRules",
        schemaVersion: 1,
        exportedAt: now.toISOString(),
        rules,
      };

      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
      const url = URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      a.download = `katapult_qc_rules_${base}_${stamp}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();

      setTimeout(() => URL.revokeObjectURL(url), 1000);
      logLine(`Saved rules to file: katapult_qc_rules_${base}_${stamp}.json`);
    } catch (err) {
      logLine(`Failed to save rules file: ${err && err.message ? err.message : String(err)}`);
    }
  }

  function coerceImportedRules(obj) {
    const src = (obj && typeof obj === "object" && obj.rules && typeof obj.rules === "object") ? obj.rules : obj;
    if (!src || typeof src !== "object") return structuredClone(DEFAULT_RULES);

    const merged = {
      pole: { ...DEFAULT_RULES.pole, ...(src.pole && typeof src.pole === "object" ? src.pole : {}) },
      midspan: { ...DEFAULT_RULES.midspan, ...(src.midspan && typeof src.midspan === "object" ? src.midspan : {}) },
    };

    const numOr = (v, fb) => {
      const n = Number(v);
      return Number.isFinite(n) ? n : fb;
    };

    // Pole numeric fields
    merged.pole.minLowestCommAttachIn = numOr(merged.pole.minLowestCommAttachIn, DEFAULT_RULES.pole.minLowestCommAttachIn);
    merged.pole.commSepDiffIn = numOr(merged.pole.commSepDiffIn, DEFAULT_RULES.pole.commSepDiffIn);
    merged.pole.commSepSameIn = numOr(merged.pole.commSepSameIn, DEFAULT_RULES.pole.commSepSameIn);
    merged.pole.commToPowerSepIn = numOr(merged.pole.commToPowerSepIn, DEFAULT_RULES.pole.commToPowerSepIn);
    merged.pole.adssCommToPowerSepIn = numOr(merged.pole.adssCommToPowerSepIn, DEFAULT_RULES.pole.adssCommToPowerSepIn);
    merged.pole.commToStreetLightSepIn = numOr(merged.pole.commToStreetLightSepIn, DEFAULT_RULES.pole.commToStreetLightSepIn);
    merged.pole.movedHoleBufferIn = numOr(merged.pole.movedHoleBufferIn, DEFAULT_RULES.pole.movedHoleBufferIn);

    // Pole boolean fields
    merged.pole.enforceAdssHighest = !!merged.pole.enforceAdssHighest;
    merged.pole.enforceEquipmentMove = !!merged.pole.enforceEquipmentMove;
    merged.pole.enforcePowerOrder = !!merged.pole.enforcePowerOrder;
    merged.pole.enforceNeutralSecondaryBelowTransformer = !!merged.pole.enforceNeutralSecondaryBelowTransformer;
    merged.pole.warnMissingPoleIdentifiers = !!merged.pole.warnMissingPoleIdentifiers;

    // Midspan numeric fields
    merged.midspan.minCommDefaultIn = numOr(merged.midspan.minCommDefaultIn, DEFAULT_RULES.midspan.minCommDefaultIn);
    merged.midspan.minCommPedestrianIn = numOr(merged.midspan.minCommPedestrianIn, DEFAULT_RULES.midspan.minCommPedestrianIn);
    merged.midspan.minCommHighwayIn = numOr(merged.midspan.minCommHighwayIn, DEFAULT_RULES.midspan.minCommHighwayIn);
    merged.midspan.minCommFarmIn = numOr(merged.midspan.minCommFarmIn, DEFAULT_RULES.midspan.minCommFarmIn);
    merged.midspan.minCommRailIn = numOr(merged.midspan.minCommRailIn, DEFAULT_RULES.midspan.minCommRailIn);
    merged.midspan.commSepIn = numOr(merged.midspan.commSepIn, DEFAULT_RULES.midspan.commSepIn);
    merged.midspan.commToPowerSepIn = numOr(merged.midspan.commToPowerSepIn, DEFAULT_RULES.midspan.commToPowerSepIn);
    merged.midspan.adssCommToPowerSepIn = numOr(merged.midspan.adssCommToPowerSepIn, DEFAULT_RULES.midspan.adssCommToPowerSepIn);
    merged.midspan.installingCompanyCommSepIn = numOr(merged.midspan.installingCompanyCommSepIn, DEFAULT_RULES.midspan.installingCompanyCommSepIn);

    // Midspan other fields
    merged.midspan.installingCompany = String(merged.midspan.installingCompany || "").trim();
    merged.midspan.enforceAdssHighest = !!merged.midspan.enforceAdssHighest;
    merged.midspan.warnMissingRowType = !!merged.midspan.warnMissingRowType;

    return merged;
  }

  async function importRulesFile(file) {
    try {
      const name = file && file.name ? String(file.name) : "(rules.json)";
      const text = await file.text();
      const parsed = JSON.parse(text);
      const nextRules = coerceImportedRules(parsed);

      rules = nextRules;
      saveRules(rules);
      applyRulesToUi();
      if (model) recomputeQc();

      logLine(`Loaded rules from file: ${name}`);
    } catch (err) {
      logLine(`Failed to load rules file: ${err && err.message ? err.message : String(err)}`);
    }
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
    if (els.rulePoleNeutralSecondaryBelowTransformer) {
      els.rulePoleNeutralSecondaryBelowTransformer.checked = !!rules.pole.enforceNeutralSecondaryBelowTransformer;
    }
    els.rulePoleWarnMissingIds.checked = !!rules.pole.warnMissingPoleIdentifiers;

    // Midspan
    els.ruleMidMinDefault.value = fmtFtIn(rules.midspan.minCommDefaultIn);
    els.ruleMidMinPed.value = fmtFtIn(rules.midspan.minCommPedestrianIn);
    els.ruleMidMinHwy.value = fmtFtIn(rules.midspan.minCommHighwayIn);
    els.ruleMidMinFarm.value = fmtFtIn(rules.midspan.minCommFarmIn);
    if (els.ruleMidMinRail) els.ruleMidMinRail.value = fmtFtIn(rules.midspan.minCommRailIn);

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
    if (els.rulePoleNeutralSecondaryBelowTransformer) {
      r.pole.enforceNeutralSecondaryBelowTransformer = !!els.rulePoleNeutralSecondaryBelowTransformer.checked;
    }
    r.pole.warnMissingPoleIdentifiers = !!els.rulePoleWarnMissingIds.checked;

    // Midspan
    r.midspan.minCommDefaultIn = parseFtIn(els.ruleMidMinDefault.value, DEFAULT_RULES.midspan.minCommDefaultIn);
    r.midspan.minCommPedestrianIn = parseFtIn(els.ruleMidMinPed.value, DEFAULT_RULES.midspan.minCommPedestrianIn);
    r.midspan.minCommHighwayIn = parseFtIn(els.ruleMidMinHwy.value, DEFAULT_RULES.midspan.minCommHighwayIn);
    r.midspan.minCommFarmIn = parseFtIn(els.ruleMidMinFarm.value, DEFAULT_RULES.midspan.minCommFarmIn);
    if (els.ruleMidMinRail) r.midspan.minCommRailIn = parseFtIn(els.ruleMidMinRail.value, DEFAULT_RULES.midspan.minCommRailIn);

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

	function normalizeKey(s) {
	  // Normalize user input / identifiers so searches are resilient to spacing, dashes, etc.
	  // (visualization-only; no impact to QC logic)
	  return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
	}

	function threeFocusOnGroup(group) {
	  if (!threeState || !threeState.camera || !threeState.controls || !group) return false;
	  try {
	    const box = new THREE.Box3().setFromObject(group);
	    if (!box || box.isEmpty()) return false;
	    const center = new THREE.Vector3();
	    box.getCenter(center);
	    const cam = threeState.camera;
	    const ctl = threeState.controls;
	
	    const offset = new THREE.Vector3().copy(cam.position).sub(ctl.target);
	    ctl.target.copy(center);
	    cam.position.copy(center).add(offset);
	    cam.lookAt(ctl.target);
	    if (typeof ctl.update === "function") ctl.update(0);
	    return true;
	  } catch (_) {
	    return false;
	  }
	}

	function focusPoleId(poleId) {
	  const id = String(poleId || "");
	  if (!id) return;

	  // Prefer focusing the current view mode.
	  if (viewMode === "3d" && threeState && threeState.poleGroupsById) {
	    const g = threeState.poleGroupsById.get(id);
	    if (g && threeFocusOnGroup(g)) {
	      selectEntity({ type: "pole", id, latlng: null });
	      return;
	    }
	  }

	  const marker = poleMarkers.get(id);
	  if (marker) {
	    map.setView(marker.getLatLng(), Math.max(map.getZoom(), 18));
	    marker.openPopup();
	    selectEntity({ type: "pole", id });
	  }
	}

	function searchAndFocusPoleByScid(query) {
	  const raw = String(query || "").trim();
	  if (!raw) return;

	  const qKey0 = normalizeKey(raw);
	  const qKey = qKey0.replace(/^scid/, "");
	  if (!qKey) return;

	  const list = (previewPoles && previewPoles.length) ? previewPoles : (model && model.poles ? model.poles : []);
	  if (!list.length) return;

	  let found = null;
	  // Prefer exact match, fall back to partial contains.
	  for (const p of list) {
	    const v = normalizeKey(p && p.scid != null ? p.scid : "");
	    if (!v) continue;
	    if (v === qKey) { found = p; break; }
	  }
	  if (!found) {
	    for (const p of list) {
	      const v = normalizeKey(p && p.scid != null ? p.scid : "");
	      if (!v) continue;
	      if (v.includes(qKey)) { found = p; break; }
	    }
	  }

	  if (!found) {
	    logLine(`SCID search: no match for "${raw}".`);
	    return;
	  }
	  focusPoleId(found.poleId);
	}

	function searchAndFocusPoleByTag(query) {
	  const raw = String(query || "").trim();
	  if (!raw) return;

	  const qKey0 = normalizeKey(raw);
	  const qKey = qKey0.replace(/^tag/, "");
	  if (!qKey) return;

	  const list = (previewPoles && previewPoles.length) ? previewPoles : (model && model.poles ? model.poles : []);
	  if (!list.length) return;

	  let found = null;
	  // Prefer exact match, fall back to partial contains.
	  for (const p of list) {
	    const v = normalizeKey(p && p.poleTag != null ? p.poleTag : "");
	    if (!v) continue;
	    if (v === qKey) { found = p; break; }
	  }
	  if (!found) {
	    for (const p of list) {
	      const v = normalizeKey(p && p.poleTag != null ? p.poleTag : "");
	      if (!v) continue;
	      if (v.includes(qKey)) { found = p; break; }
	    }
	  }

	  if (!found) {
	    logLine(`Pole Tag search: no match for "${raw}".`);
	    return;
	  }
	  focusPoleId(found.poleId);
	}

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
