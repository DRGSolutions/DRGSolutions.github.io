/* Katapult JSON → Workbook Generator
   - Static, client-only app (safe to host on the internet)
   - Uses a Web Worker for parsing/normalization, then ExcelJS to write XLSX.
   - Optional polygon selection to generate a workbook for a subset of poles.
*/

const $ = (sel) => document.querySelector(sel);

const jsonFile = $("#jsonFile");
const btnGenerate = $("#btnGenerate");
const logEl = $("#log");
const progressBar = $("#progressBar");
const progressLabel = $("#progressLabel");
const downloadLink = $("#downloadLink");
const optNodeColorAttr = $("#optNodeColorAttr");

const btnSelectAll = $("#btnSelectAll");
const btnClearSelection = $("#btnClearSelection");
const selCountEl = $("#selCount");
const selHintEl = $("#selHint");

// Hidden/internal defaults (we keep the converter behavior stable by default)
const DEFAULTS = {
  includeMidspans: true,
  includeGuys: true,
  includeEquipment: true,
  mergePoleCols: true,
  moveUnit: "in", // safest default for Katapult exports
};


function createInlineWorker() {
  // Works on HTTPS origins and also when opened via file:// (origin "null").
  const el = document.getElementById('worker-js');
  if (!el) throw new Error('Inline worker code not found (missing #worker-js).');
  const code = el.textContent || '';
  const blob = new Blob([code], { type: 'text/javascript' });
  const url = URL.createObjectURL(blob);
  const w = new Worker(url);
  // We can revoke the URL after the worker starts; keep a reference just in case.
  w.__blobUrl = url;
  return w;
}
let selectedFile = null;

// Map + selection state
let map = null;
let poleMarkers = null; // L.LayerGroup
let poleLabels = null; // L.LayerGroup (SCID labels)
let drawControl = null;
let drawnItems = null; // L.FeatureGroup
let previewPoles = []; // [{ poleId, scid, poleTag, displayName, lat, lon, nodeColorHex }]
let selectedPoleIds = new Set();

function setSelectionCount(n) {
  if (selCountEl) selCountEl.textContent = String(n);
}

function setSelectionHint(text) {
  if (selHintEl) selHintEl.textContent = text;
}

function clearSelection() {
  selectedPoleIds = new Set();
  updateMarkerStyles();
  setSelectionCount(0);
  setSelectionHint("Draw a polygon on the map to select poles (or select all).");
  btnGenerate.disabled = true;
}

function selectAllPoles() {
  try { if (drawnItems) drawnItems.clearLayers(); } catch (_) {}
  selectedPoleIds = new Set(previewPoles.map(p => p.poleId));
  updateMarkerStyles();
  setSelectionCount(selectedPoleIds.size);
  setSelectionHint("All poles selected.");
  btnGenerate.disabled = selectedPoleIds.size === 0;
}

function updateMarkerStyles() {
  if (!poleMarkers) return;
  poleMarkers.eachLayer((layer) => {
    if (!layer || !layer.__pole) return;
    const pole = layer.__pole;
    const selected = selectedPoleIds.has(pole.poleId);
    const fill = pole.nodeColorHex || "#ffffff";
    layer.setStyle({
      fillColor: fill,
      fillOpacity: selected ? 0.95 : 0.55,
      opacity: selected ? 1 : 0.55,
      weight: selected ? 2 : 1,
      radius: selected ? 7 : 5,
    });
  });
}

function ensureMap() {
  if (map) return map;
  const el = document.getElementById("map");
  if (!el) return null;

  map = L.map(el, {
    zoomControl: true,
    attributionControl: true,
  });

  // Basemap: OpenStreetMap tiles
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 20,
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);

  poleMarkers = L.layerGroup().addTo(map);
  poleLabels = L.layerGroup().addTo(map);
  drawnItems = new L.FeatureGroup().addTo(map);

  drawControl = new L.Control.Draw({
    edit: {
      featureGroup: drawnItems,
      remove: true,
    },
    draw: {
      polygon: {
        allowIntersection: false,
        showArea: true,
      },
      rectangle: true,
      circle: false,
      circlemarker: false,
      marker: false,
      polyline: false,
    },
  });
  map.addControl(drawControl);

  map.on(L.Draw.Event.CREATED, (e) => {
    const layer = e.layer;
    drawnItems.addLayer(layer);
    applyPolygonSelection();
  });

  map.on(L.Draw.Event.EDITED, () => {
    applyPolygonSelection(true);
  });

  map.on(L.Draw.Event.DELETED, () => {
    applyPolygonSelection(true);
  });

  return map;
}

function pointInPolygon(pointLngLat, polygonLngLat) {
  // Ray-casting algorithm. pointLngLat = [lng, lat]
  const x = pointLngLat[0];
  const y = pointLngLat[1];
  let inside = false;
  for (let i = 0, j = polygonLngLat.length - 1; i < polygonLngLat.length; j = i++) {
    const xi = polygonLngLat[i][0], yi = polygonLngLat[i][1];
    const xj = polygonLngLat[j][0], yj = polygonLngLat[j][1];
    const intersect = ((yi > y) !== (yj > y)) &&
      (x < (xj - xi) * (y - yi) / ((yj - yi) || 1e-12) + xi);
    if (intersect) inside = !inside;
  }
  return inside;
}

function polygonLatLngsToLngLat(poly) {
  // Leaflet polygon latlngs can be nested. Normalize to first ring.
  const latlngs = poly.getLatLngs();
  const ring = Array.isArray(latlngs[0]) ? latlngs[0] : latlngs;
  return ring.map(ll => [ll.lng, ll.lat]);
}

function applyPolygonSelection(recompute = false) {
  if (!drawnItems) return;
  const polys = [];
  drawnItems.eachLayer((layer) => {
    if (layer && typeof layer.getLatLngs === "function") {
      polys.push(polygonLatLngsToLngLat(layer));
    }
  });

  if (!polys.length) {
    if (recompute) clearSelection();
    return;
  }

  // Union selection across all polygons.
  const next = new Set();
  for (const p of previewPoles) {
    if (p.lat == null || p.lon == null) continue;
    const pt = [p.lon, p.lat];
    for (const poly of polys) {
      if (pointInPolygon(pt, poly)) {
        next.add(p.poleId);
        break;
      }
    }
  }

  selectedPoleIds = next;
  updateMarkerStyles();
  setSelectionCount(selectedPoleIds.size);
  setSelectionHint(selectedPoleIds.size ? "Selection updated." : "No poles selected in the drawn area.");
  btnGenerate.disabled = selectedPoleIds.size === 0;
}

function log(msg) {
  const ts = new Date().toLocaleTimeString();
  logEl.textContent += `[${ts}] ${msg}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}

function setProgress(pct, label) {
  progressBar.style.width = `${Math.max(0, Math.min(100, pct))}%`;
  if (label) progressLabel.textContent = label;
}

jsonFile.addEventListener("change", () => {
  selectedFile = jsonFile.files && jsonFile.files[0] ? jsonFile.files[0] : null;
  btnGenerate.disabled = true;
  downloadLink.classList.add("hidden");
  downloadLink.removeAttribute("href");
  downloadLink.textContent = "";
  logEl.textContent = "";
  if (selectedFile) {
    log(`Selected: ${selectedFile.name} (${(selectedFile.size/1024/1024).toFixed(1)} MB)`);
    setProgress(0, "Loading map…");
    loadPreviewAndMap(selectedFile).catch((err) => {
      console.error(err);
      log(`ERROR: ${err && err.message ? err.message : String(err)}`);
      setProgress(0, "Error.");
      alert(`Failed to read JSON for mapping: ${err && err.message ? err.message : String(err)}`);
    });
  }
});

// If the user changes the tab color attribute after loading, re-index
// pole colors so the map preview and Excel tabs stay aligned.
let _colorAttrTimer = null;
optNodeColorAttr?.addEventListener("input", () => {
  if (!selectedFile) return;
  clearTimeout(_colorAttrTimer);
  _colorAttrTimer = setTimeout(() => {
    loadPreviewAndMap(selectedFile).catch((err) => {
      console.error(err);
      log(`ERROR: ${err && err.message ? err.message : String(err)}`);
      setProgress(0, "Error.");
    });
  }, 350);
});

btnSelectAll?.addEventListener("click", () => {
  if (!previewPoles.length) return;
  selectAllPoles();
});

btnClearSelection?.addEventListener("click", () => {
  if (!previewPoles.length) return;
  try { if (drawnItems) drawnItems.clearLayers(); } catch (_) {}
  clearSelection();
});

async function loadPreviewAndMap(file) {
  // Reset state
  previewPoles = [];
  clearSelection();
  btnSelectAll.disabled = true;
  btnClearSelection.disabled = true;

  const m = ensureMap();
  if (!m) throw new Error("Map container not found.");
  if (poleMarkers) poleMarkers.clearLayers();
  if (poleLabels) poleLabels.clearLayers();
  if (drawnItems) drawnItems.clearLayers();

  setProgress(2, "Reading file…");
  const buf = await file.arrayBuffer();

  setProgress(6, "Indexing poles for map…");
  log("Starting worker preview parse (map indexing)…");

  const worker = createInlineWorker();
  const opts = {
    nodeColorAttribute: (optNodeColorAttr && optNodeColorAttr.value ? optNodeColorAttr.value.trim() : "") || null,
    moveUnit: DEFAULTS.moveUnit,
  };

  const preview = await new Promise((resolve, reject) => {
    worker.onmessage = (ev) => {
      const msg = ev.data || {};
      if (msg.type === "log") log(msg.message);
      if (msg.type === "progress") setProgress(msg.pct, msg.label || undefined);
      if (msg.type === "preview") resolve(msg.payload);
      if (msg.type === "error") reject(new Error(msg.message || "Worker error"));
    };
    worker.onerror = (err) => reject(err);
    // Do NOT transfer ownership of buf; keep it usable in the main thread.
    worker.postMessage({ type: "preview", buffer: buf, options: opts });
  });

  worker.terminate();
  try { if (worker.__blobUrl) URL.revokeObjectURL(worker.__blobUrl); } catch (_) {}

  previewPoles = Array.isArray(preview && preview.poles) ? preview.poles : [];
  if (!previewPoles.length) {
    setProgress(0, "No poles found.");
    setSelectionHint("No poles found in this JSON.");
    return;
  }

  // Add markers
  for (const p of previewPoles) {
    if (p.lat == null || p.lon == null) continue;

    // Always-visible SCID label so users can confidently pick poles.
    // This is intentionally non-interactive so clicks still land on the pole marker.
    if (poleLabels) {
      const scidLabel = (p.scid != null && String(p.scid).trim() !== "") ? String(p.scid).trim() : "—";
      const labelMarker = L.marker([p.lat, p.lon], {
        interactive: false,
        keyboard: false,
        icon: L.divIcon({
          className: "scid-label-icon",
          html: `<div class="scid-label-text">${escapeHtml(scidLabel)}</div>`,
        }),
      });
      labelMarker.addTo(poleLabels);
    }

    const fill = p.nodeColorHex || "#ffffff";
    const marker = L.circleMarker([p.lat, p.lon], {
      radius: 5,
      weight: 1,
      color: "#111827",
      fillColor: fill,
      fillOpacity: 0.55,
    });
    marker.__pole = p;
    marker.bindTooltip(`${escapeHtml(p.displayName || "Pole")}`, { direction: "top", opacity: 0.95 });
    marker.on("click", () => {
      // Toggle selection on click.
      if (selectedPoleIds.has(p.poleId)) selectedPoleIds.delete(p.poleId);
      else selectedPoleIds.add(p.poleId);
      updateMarkerStyles();
      setSelectionCount(selectedPoleIds.size);
      btnGenerate.disabled = selectedPoleIds.size === 0;
      setSelectionHint("Tip: draw polygons for bulk selection; click markers to toggle individual poles.");
    });
    marker.addTo(poleMarkers);
  }

  // Fit view
  const latlngs = previewPoles
    .filter(p => p.lat != null && p.lon != null)
    .map(p => L.latLng(p.lat, p.lon));
  if (latlngs.length) {
    const bounds = L.latLngBounds(latlngs);
    m.fitBounds(bounds.pad(0.15));
  } else {
    m.setView([0, 0], 2);
  }

  btnSelectAll.disabled = false;
  btnClearSelection.disabled = false;
  setProgress(0, "Ready.");
  setSelectionHint("Draw a polygon on the map to select poles (or select all)." );
}

btnGenerate.addEventListener("click", async () => {
  if (!selectedFile) return;
  if (!selectedPoleIds || selectedPoleIds.size === 0) {
    alert("Please select at least one pole (draw a polygon or select all).");
    return;
  }

  try {
    btnGenerate.disabled = true;
    downloadLink.classList.add("hidden");

    setProgress(2, "Reading file…");
    log("Reading JSON file into memory…");
    const buf = await selectedFile.arrayBuffer();

    setProgress(6, "Parsing + normalizing…");
    log("Starting Web Worker parse/normalize…");

    const worker = createInlineWorker();
    const opts = {
      includeMidspans: DEFAULTS.includeMidspans,
      includeGuys: DEFAULTS.includeGuys,
      includeEquipment: DEFAULTS.includeEquipment,
      mergePoleCols: DEFAULTS.mergePoleCols,
      nodeColorAttribute: (optNodeColorAttr && optNodeColorAttr.value ? optNodeColorAttr.value.trim() : "") || null,
      moveUnit: DEFAULTS.moveUnit,
      selectedPoleIds: Array.from(selectedPoleIds),
    };

    const normalized = await new Promise((resolve, reject) => {
      worker.onmessage = (ev) => {
        const msg = ev.data || {};
        if (msg.type === "log") log(msg.message);
        if (msg.type === "progress") setProgress(msg.pct, msg.label || undefined);
        if (msg.type === "done") resolve(msg.payload);
        if (msg.type === "error") reject(new Error(msg.message || "Worker error"));
      };
      worker.onerror = (err) => reject(err);
      worker.postMessage({ type: "start", buffer: buf, options: opts }, [buf]);
    });

    worker.terminate();
    try { if (worker.__blobUrl) URL.revokeObjectURL(worker.__blobUrl); } catch (_) {}

    setProgress(65, "Building Excel workbook…");
    log("Building XLSX using ExcelJS…");

    const wb = new ExcelJS.Workbook();
    wb.creator = "Katapult JSON → Workbook Generator (browser)";
    wb.created = new Date();

    // Cover sheet: "Make Ready App Info"
    const cover = wb.addWorksheet("Make Ready App Info");
    buildCoverSheet(cover, normalized);

    // Optional GPS Points sheet (useful if you want to map later)
    const pts = wb.addWorksheet("GPS Points");
    buildGpsPointsSheet(pts, normalized);

    // Pole sheets
    const poleCount = normalized.poles.length;
    for (let i = 0; i < poleCount; i++) {
      const pole = normalized.poles[i];
      setProgress(65 + Math.floor((i / Math.max(1, poleCount)) * 30), `Writing pole sheets (${i+1}/${poleCount})…`);
      const ws = wb.addWorksheet(pole.sheetName);
      // Match Katapult node color (when available) for the Excel tab color.
      if (pole.nodeColorHex) {
        const argb = hexToArgb(pole.nodeColorHex);
        if (argb) ws.properties.tabColor = { argb };
      }
      buildPoleSheet(ws, pole, { mergePoleCols: !!DEFAULTS.mergePoleCols });
    }

    setProgress(96, "Finalizing workbook…");
    const outName = normalized.suggestedFilename;
    const xlsxBuf = await wb.xlsx.writeBuffer();
    const blob = new Blob([xlsxBuf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);

    // Auto-download
    downloadLink.href = url;
    downloadLink.download = outName;
    downloadLink.textContent = `Download: ${outName}`;
    downloadLink.classList.remove("hidden");

    // Trigger automatic download (most browsers allow this after a user-initiated click).
    try {
      downloadLink.click();
    } catch (_) {
      // If blocked, the visible link remains as a fallback.
    }

    // Best-effort cleanup.
    setTimeout(() => {
      try { URL.revokeObjectURL(url); } catch (_) {}
    }, 60_000);

    setProgress(100, "Done.");
    log("Done.");
  } catch (err) {
    console.error(err);
    log(`ERROR: ${err && err.message ? err.message : String(err)}`);
    setProgress(0, "Error.");
    alert(`Failed to generate workbook: ${err && err.message ? err.message : String(err)}`);
  } finally {
    btnGenerate.disabled = !selectedFile || !selectedPoleIds || selectedPoleIds.size === 0;
  }
});

function buildCoverSheet(ws, normalized) {
  ws.getColumn(1).width = 24;
  ws.getColumn(2).width = 44;

  ws.mergeCells("A1:D1");
  ws.getCell("A1").value = "Make Ready Application Information";
  ws.getCell("A1").font = { bold: true, size: 14 };
  ws.getCell("A1").alignment = { horizontal: "center" };

  // Keep this sheet minimal and utility-agnostic.
  // Requested removals: Job Owner, Job Creator, Last Upload, Total # of Midspans,
  // Company (best guess), Location (best guess).
  const rows = [
    ["Job Name:", normalized.jobName || ""],
    ["Date Created:", normalized.dateCreated || ""],
    ["Total # of Poles:", normalized.poles.length],
  ];

  let r = 2;
  for (const [k, v] of rows) {
    ws.getCell(r, 1).value = k;
    ws.getCell(r, 2).value = v;
    ws.getCell(r, 1).font = { bold: true };
    ws.getCell(r, 1).alignment = { vertical: "middle" };
    ws.getCell(r, 2).alignment = { vertical: "middle" };
    r++;
  }

  // Add a light border around A1:B{r-1}
  for (let row = 1; row <= r - 1; row++) {
    for (let col = 1; col <= 2; col++) {
      ws.getCell(row, col).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    }
  }
}

function escapeHtml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function buildGpsPointsSheet(ws, normalized) {
  ws.columns = [
    { header: "Sheet", key: "sheet", width: 32 },
    { header: "Pole Tag", key: "tag", width: 16 },
    { header: "SCID", key: "scid", width: 10 },
    { header: "Latitude", key: "lat", width: 14 },
    { header: "Longitude", key: "lon", width: 14 },
  ];
  ws.getRow(1).font = { bold: true };

  for (const p of normalized.poles) {
    ws.addRow({
      sheet: p.sheetName,
      tag: p.poleTag,
      scid: p.scid,
      lat: p.lat,
      lon: p.lon,
    });
  }
}

function buildPoleSheet(ws, pole, { mergePoleCols }) {
  // Mimic the general layout produced by the legacy converter:
  // Row 1: static headers + per-direction group headers
  // Row 2: per-direction subheaders (wire label / existing / proposed)
  // Rows 3..N: data

  const staticCols = [
    "SCID , Tag",
    "Pole Owner",
    "Owner",
    "Existing Height",
    "Proposed Height",
    "Pole Height & Class",
    "Pole Replacement Needed (PLA)",
    "Comments (Group OR Equipment)",
    "Make-Ready Notes",
  ];

  const directions = pole.directions || [];
  // Number of midspan measurement points can vary per direction.
  // Worker emits pole.midspanSlotsByDir[deg] so we can render additional Existing/Proposed columns.
  const slotsByDir = { ...(pole.midspanSlotsByDir || {}) };

  // Fallback: derive slot counts from row data if not provided.
  if (!pole.midspanSlotsByDir) {
    for (const deg of directions) slotsByDir[deg] = 1;
    for (const row of pole.rows || []) {
      const exMap = row.midspansExisting || {};
      const prMap = row.midspansProposed || {};
      for (const deg of directions) {
        const exRaw = exMap[deg];
        const prRaw = prMap[deg];
        const exArr = Array.isArray(exRaw) ? exRaw : (exRaw ? [exRaw] : []);
        const prArr = Array.isArray(prRaw) ? prRaw : (prRaw ? [prRaw] : []);
        const n = Math.max(exArr.length, prArr.length, 1);
        if (!slotsByDir[deg] || n > slotsByDir[deg]) slotsByDir[deg] = n;
      }
    }
  }

  // Header rows
  const header1 = [...staticCols];
  const header2 = new Array(staticCols.length).fill("");

  for (const deg of directions) {
    const slots = Math.max(1, parseInt(slotsByDir[deg] || 1, 10) || 1);
    const groupWidth = 1 + (2 * slots);

    // Row 1: single group label (merged across the direction's columns)
    header1.push(`Pole → ${cardinalFromDegrees(deg)} ${deg} °`);
    for (let i = 1; i < groupWidth; i++) header1.push("");

    // Row 2: subheaders
    header2.push("Wire/Group");
    for (let i = 1; i <= slots; i++) {
      const suf = slots > 1 ? ` ${i}` : "";
      header2.push(`Existing Midspan${suf}`, `Proposed Midspan${suf}`);
    }
  }

  ws.addRow(header1);
  ws.addRow(header2);

  // Merge each direction group header across its full width
  let groupStart = staticCols.length + 1;
  for (const deg of directions) {
    const slots = Math.max(1, parseInt(slotsByDir[deg] || 1, 10) || 1);
    const groupWidth = 1 + (2 * slots);
    const groupEnd = groupStart + groupWidth - 1;
    ws.mergeCells(1, groupStart, 1, groupEnd);
    groupStart = groupEnd + 1;
  }

  ws.getRow(1).font = { bold: true };
  ws.getRow(1).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  ws.getRow(2).font = { bold: true };
  ws.getRow(2).alignment = { horizontal: "center", vertical: "middle", wrapText: true };

  // Column widths
  const widths = [28, 12, 22, 16, 16, 18, 22, 30, 44];
  for (let i = 0; i < staticCols.length; i++) {
    ws.getColumn(i + 1).width = widths[i] || 18;
  }

  // Midspan columns: label + (Existing/Proposed) x N
  let colIdx = staticCols.length + 1;
  for (const deg of directions) {
    const slots = Math.max(1, parseInt(slotsByDir[deg] || 1, 10) || 1);
    ws.getColumn(colIdx).width = 22; // label
    colIdx += 1;
    for (let i = 0; i < slots; i++) {
      ws.getColumn(colIdx).width = 16;     // existing
      ws.getColumn(colIdx + 1).width = 16; // proposed
      colIdx += 2;
    }
  }

  // Data rows
  for (const row of pole.rows) {
    const base = [
      pole.displayName,
      pole.poleOwnerAbbrev || "",
      row.owner || "",
      row.existingHeight || "",
      row.proposedHeight || "",
      pole.poleHeightClass || "",
      row.pla || "NO",
      row.comments || "",
      row.makeReadyNotes || "",
    ];

    const piv = [];
    for (const deg of directions) {
      const slots = Math.max(1, parseInt(slotsByDir[deg] || 1, 10) || 1);

      const exRaw = row.midspansExisting && row.midspansExisting[deg] ? row.midspansExisting[deg] : null;
      const prRaw = row.midspansProposed && row.midspansProposed[deg] ? row.midspansProposed[deg] : null;
      const exArr = Array.isArray(exRaw) ? exRaw : (exRaw ? [exRaw] : []);
      const prArr = Array.isArray(prRaw) ? prRaw : (prRaw ? [prRaw] : []);

      const hasAny = exArr.some(v => v && String(v).trim() !== "") || prArr.some(v => v && String(v).trim() !== "");
      if (hasAny) {
        piv.push(row.comments || "");
        for (let i = 0; i < slots; i++) {
          piv.push(exArr[i] || "");
          piv.push(prArr[i] || "");
        }
      } else {
        piv.push("");
        for (let i = 0; i < slots; i++) {
          piv.push("", "");
        }
      }
    }

    ws.addRow([...base, ...piv]);
  }

  const lastRow = ws.rowCount;

  // Merge SCID , Tag and Pole Owner down the table (optional)
  if (mergePoleCols && lastRow >= 3) {
    ws.mergeCells(3, 1, lastRow, 1);
    ws.mergeCells(3, 2, lastRow, 2);

    // Also merge "Pole Height & Class" and "Pole Replacement Needed (PLA)" like your script
    ws.mergeCells(3, 6, lastRow, 6);
    ws.mergeCells(3, 7, lastRow, 7);
  }

  // Borders + alignment
  for (let r = 1; r <= lastRow; r++) {
    const row = ws.getRow(r);
    row.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    for (let c = 1; c <= ws.columnCount; c++) {
      const cell = ws.getCell(r, c);
      cell.border = {
        top: { style: r === 1 ? "thick" : "thin" },
        left: { style: r === 1 ? "thick" : "thin" },
        bottom: { style: r === 1 ? "thick" : "thin" },
        right: { style: r === 1 ? "thick" : "thin" },
      };
    }
  }

  // Put sheet display name into A3 (this matches your Excel-merging convention)
  if (lastRow >= 3) {
    ws.getCell(3, 1).value = pole.displayName;
  }
}

function cardinalFromDegrees(deg) {
  const dirs = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"];
  const idx = Math.floor(((deg % 360) + 22.5) / 45) % 8;
  return dirs[idx];
}
function hexToArgb(hex) {
  if (!hex) return null;
  let s = String(hex).trim();
  if (!s) return null;
  if (s.startsWith("#")) s = s.slice(1);
  if (s.length === 3) s = s.split("").map(ch => ch + ch).join("");
  if (s.length !== 6) return null;
  return ("FF" + s).toUpperCase();
}


