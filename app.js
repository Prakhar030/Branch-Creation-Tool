/*
 * Branch Creation Tool
 * -----------------------
 * Reads an EXISTING, already-final Seller Clustering Excel export (Cluster ID,
 * GLID, Latitude, Longitude, Outstanding Amount, Renewal Date, Eligibility)
 * and lets a Regional/Sales Manager draw polygons ("Branches") that group
 * the EXISTING Clusters together. It never creates, splits, recalculates, or
 * renumbers Clusters, and never re-runs any clustering algorithm - Clusters
 * are read-only input here.
 */

// ------------------------------------------------------------------
// Column resolution
// ------------------------------------------------------------------

const REQUIRED_ALIASES = {
  GLID: ["glid"],
  LAT: ["latitude"],
  LON: ["longitude"],
  OUTSTANDING: ["outstanding amount"],
  RENEWAL: ["renewal date"],
  ELIGIBILITY: ["eligibility", "business eligibility status"],
  CLUSTER: ["cluster id"],
};

const OPTIONAL_SELLER_ALIASES = {
  CENTROID_LAT: ["cluster centroid latitude", "centroid latitude"],
  CENTROID_LON: ["cluster centroid longitude", "centroid longitude"],
};

const SUMMARY_ALIASES = {
  CLUSTER: ["cluster id"],
  CENTROID_LAT: ["centroid latitude"],
  CENTROID_LON: ["centroid longitude"],
  AVG_DIST: ["average distance"],
  SPREAD: ["geographic spread"],
  MAX_DIST: ["maximum distance"],
  MIN_DIST: ["minimum distance"],
};

function resolveColumn(headers, aliases) {
  const lowered = headers.map((h) => ({ orig: h, low: String(h).trim().toLowerCase() }));
  for (const alias of aliases) {
    const hit = lowered.find((h) => h.low === alias);
    if (hit) return hit.orig;
  }
  for (const alias of aliases) {
    const hit = lowered.find((h) => h.low.includes(alias));
    if (hit) return hit.orig;
  }
  return null;
}

function resolveAllRequired(headers) {
  const map = {};
  const missing = [];
  for (const [key, aliases] of Object.entries(REQUIRED_ALIASES)) {
    const col = resolveColumn(headers, aliases);
    if (!col) missing.push(key);
    else map[key] = col;
  }
  for (const [key, aliases] of Object.entries(OPTIONAL_SELLER_ALIASES)) {
    const col = resolveColumn(headers, aliases);
    if (col) map[key] = col;
  }
  return { ok: missing.length === 0, map, missing };
}

function resolveSummaryColumns(headers) {
  const map = {};
  for (const [key, aliases] of Object.entries(SUMMARY_ALIASES)) {
    const col = resolveColumn(headers, aliases);
    if (col) map[key] = col;
  }
  return map;
}

function parseWorkbook(workbook) {
  const sheetNames = workbook.SheetNames;
  const ordered = [...sheetNames].sort((a, b) => (/seller/i.test(a) ? 0 : 1) - (/seller/i.test(b) ? 0 : 1));

  let sellerRows = null;
  let columns = null;
  let sheetName = null;

  for (const name of ordered) {
    const ws = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
    if (!rows.length) continue;
    const headers = Object.keys(rows[0]);
    const res = resolveAllRequired(headers);
    if (res.ok) {
      sellerRows = rows;
      columns = res.map;
      sheetName = name;
      break;
    }
  }

  if (!sellerRows) {
    return {
      error:
        "Could not find a sheet containing all required columns: GLID, Latitude, Longitude, " +
        "Outstanding Amount, Renewal Date, Eligibility, Cluster ID.",
    };
  }

  let summaryRows = null;
  let summaryCols = null;
  for (const name of sheetNames) {
    if (/cluster summary/i.test(name)) {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: null, raw: true });
      if (rows.length) {
        const headers = Object.keys(rows[0]);
        const res2 = resolveSummaryColumns(headers);
        if (res2.CLUSTER) {
          summaryRows = rows;
          summaryCols = res2;
        }
      }
      break;
    }
  }

  return { sellerRows, columns, sheetName, summaryRows, summaryCols };
}

// ------------------------------------------------------------------
// Validation + seller construction
// ------------------------------------------------------------------

function buildSellers(sellerRows, cols) {
  const errors = [];
  const warnings = [];
  const seen = new Map(); // glid -> clusterId (first occurrence)
  const sellers = [];

  sellerRows.forEach((row, idx) => {
    const rowNum = idx + 2;
    const glidRaw = row[cols.GLID];

    if (glidRaw === null || glidRaw === undefined || String(glidRaw).trim() === "") {
      errors.push(`Row ${rowNum}: missing GLID - row skipped`);
      return;
    }
    const glid = String(glidRaw).trim();

    const latRaw = row[cols.LAT];
    const lonRaw = row[cols.LON];
    const lat = latRaw === null || latRaw === undefined || latRaw === "" ? NaN : Number(latRaw);
    const lon = lonRaw === null || lonRaw === undefined || lonRaw === "" ? NaN : Number(lonRaw);
    const validCoords = Number.isFinite(lat) && Number.isFinite(lon) && lat >= -90 && lat <= 90 && lon >= -180 && lon <= 180;
    if (!validCoords) {
      warnings.push(`GLID ${glid} (row ${rowNum}): invalid/missing Latitude or Longitude - excluded from map & Branch assignment`);
    }

    const clusterRaw = row[cols.CLUSTER];
    const clusterId = clusterRaw === null || clusterRaw === undefined || String(clusterRaw).trim() === "" ? null : String(clusterRaw).trim();
    if (validCoords && clusterId === null) {
      warnings.push(`GLID ${glid} (row ${rowNum}): has valid coordinates but no Cluster ID - cannot be assigned to a Branch`);
    }

    const eligRaw = row[cols.ELIGIBILITY];
    const eligStr = eligRaw === null || eligRaw === undefined ? "" : String(eligRaw).trim();
    const eligible = /^eligible$/i.test(eligStr);
    if (eligStr && !/^eligible$/i.test(eligStr) && !/^not eligible$/i.test(eligStr)) {
      warnings.push(`GLID ${glid} (row ${rowNum}): unrecognised Eligibility value "${eligStr}" - treated as Not Eligible`);
    }

    if (seen.has(glid)) {
      const prevClusterId = seen.get(glid);
      if (prevClusterId !== clusterId) {
        errors.push(
          `GLID ${glid} (row ${rowNum}): appears more than once with different Cluster IDs (${prevClusterId} vs ${clusterId}) - ` +
            `a seller cannot belong to multiple Clusters; first occurrence kept`
        );
      } else {
        warnings.push(`GLID ${glid} (row ${rowNum}): duplicate row - ignored`);
      }
      return;
    }
    seen.set(glid, clusterId);

    sellers.push({
      glid,
      lat,
      lon,
      validCoords,
      outstanding: cols.OUTSTANDING ? row[cols.OUTSTANDING] : null,
      renewal: cols.RENEWAL ? row[cols.RENEWAL] : null,
      eligibilityRaw: eligStr,
      eligible,
      clusterId,
      centroidLat: cols.CENTROID_LAT && row[cols.CENTROID_LAT] != null ? Number(row[cols.CENTROID_LAT]) : null,
      centroidLon: cols.CENTROID_LON && row[cols.CENTROID_LON] != null ? Number(row[cols.CENTROID_LON]) : null,
    });
  });

  return { sellers, errors, warnings };
}

// ------------------------------------------------------------------
// Cluster construction + geometry
// ------------------------------------------------------------------

function buildClusterGeometry(c) {
  const points = c.sellers.map((s) => turf.point([s.lon, s.lat]));
  let poly = null;
  if (points.length >= 3) {
    try {
      poly = turf.convex(turf.featureCollection(points));
    } catch (e) {
      poly = null;
    }
  }
  if (!poly) {
    const center = turf.point([c.centroid[1], c.centroid[0]]);
    poly = turf.buffer(center, 0.1, { units: "kilometers", steps: 16 });
  }
  c.turfPoly = poly;
  c.area = turf.area(poly);
  c.bbox = turf.bbox(poly);
  c.centroidPoint = turf.point([c.centroid[1], c.centroid[0]]);
  c.hullLatLngs = poly.geometry.coordinates[0].map((pt) => [pt[1], pt[0]]);
}

function buildClusters(sellers, summaryRows, summaryCols) {
  const map = new Map();
  for (const s of sellers) {
    if (!s.validCoords || !s.clusterId) continue;
    if (!map.has(s.clusterId)) {
      map.set(s.clusterId, { id: s.clusterId, sellers: [], eligibleCount: 0, ineligibleCount: 0 });
    }
    const c = map.get(s.clusterId);
    c.sellers.push(s);
    if (s.eligible) c.eligibleCount++;
    else c.ineligibleCount++;
  }

  const summaryByCluster = {};
  if (summaryRows && summaryCols && summaryCols.CLUSTER) {
    for (const row of summaryRows) {
      const cid = row[summaryCols.CLUSTER];
      if (cid === null || cid === undefined) continue;
      summaryByCluster[String(cid).trim()] = row;
    }
  }

  for (const [cid, c] of map) {
    const firstWithCentroid = c.sellers.find((s) => s.centroidLat != null && s.centroidLon != null);
    const summaryRow = summaryByCluster[cid];
    if (firstWithCentroid) {
      c.centroid = [firstWithCentroid.centroidLat, firstWithCentroid.centroidLon];
    } else if (summaryRow && summaryCols.CENTROID_LAT && summaryRow[summaryCols.CENTROID_LAT] != null) {
      c.centroid = [Number(summaryRow[summaryCols.CENTROID_LAT]), Number(summaryRow[summaryCols.CENTROID_LON])];
    } else {
      const meanLat = c.sellers.reduce((a, s) => a + s.lat, 0) / c.sellers.length;
      const meanLon = c.sellers.reduce((a, s) => a + s.lon, 0) / c.sellers.length;
      c.centroid = [meanLat, meanLon];
    }

    c.extra = null;
    if (summaryRow) {
      c.extra = {
        avgDist: summaryCols.AVG_DIST ? summaryRow[summaryCols.AVG_DIST] : null,
        spread: summaryCols.SPREAD ? summaryRow[summaryCols.SPREAD] : null,
        maxDist: summaryCols.MAX_DIST ? summaryRow[summaryCols.MAX_DIST] : null,
        minDist: summaryCols.MIN_DIST ? summaryRow[summaryCols.MIN_DIST] : null,
      };
    }

    buildClusterGeometry(c);
  }

  return map;
}

// ------------------------------------------------------------------
// Global state
// ------------------------------------------------------------------

const STATE = {
  sellers: [],
  clusters: new Map(),
  rawSummaryRows: null,
  rawSummaryCols: null,
  branches: [],
  overrides: {}, // clusterId -> branchId | 'UNMAPPED'
  assignments: {}, // clusterId -> {status, branchId?, candidates?, manual?}
  assignmentRule: "C",
  selectedBranchId: null,
  sourceFileName: "",
  branchCounter: 1,
  map: null,
};

const PALETTE = [
  "#e6194b", "#3cb44b", "#4363d8", "#f58231", "#911eb4",
  "#46b3b3", "#f032e6", "#9a9a00", "#008080", "#9a6324",
  "#800000", "#5b6ee1", "#c9a300", "#a9327a", "#2b6d3c",
];
function nextColor(idx) {
  return PALETTE[idx % PALETTE.length];
}

// ------------------------------------------------------------------
// Helpers
// ------------------------------------------------------------------

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));
}
function cssId(cid) {
  return String(cid).replace(/[^a-zA-Z0-9_-]/g, "_");
}
function formatDateVal(v) {
  if (v === null || v === undefined || v === "") return "N/A";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v);
}
function bboxIntersects(a, b) {
  return a[0] <= b[2] && a[2] >= b[0] && a[1] <= b[3] && a[3] >= b[1];
}

// ------------------------------------------------------------------
// Modal helpers
// ------------------------------------------------------------------

const modalOverlay = document.getElementById("modalOverlay");
const modalBox = document.getElementById("modalBox");

function openModal(html) {
  modalBox.innerHTML = html;
  modalOverlay.hidden = false;
}
function closeModal() {
  modalOverlay.hidden = true;
  modalBox.innerHTML = "";
}
function alertModal(msg) {
  const html = escapeHtml(msg).replace(/\n/g, "<br/>");
  openModal(`<div>${html}</div><div class="modal-actions"><button id="amOk" class="primary">OK</button></div>`);
  document.getElementById("amOk").onclick = closeModal;
}
function confirmModal(msg, onConfirm) {
  openModal(
    `<p>${escapeHtml(msg)}</p><div class="modal-actions"><button id="cmCancel">Cancel</button><button id="cmOk" class="danger">Confirm</button></div>`
  );
  document.getElementById("cmCancel").onclick = closeModal;
  document.getElementById("cmOk").onclick = () => {
    closeModal();
    onConfirm();
  };
}
function promptBranchName(defaultName, onConfirm) {
  openModal(`
    <h3>Name this Branch</h3>
    <input type="text" id="branchNameInput" placeholder="e.g. South Delhi" value="${escapeHtml(defaultName || "")}" />
    <div class="modal-actions">
      <button id="branchNameCancel">Discard polygon</button>
      <button id="branchNameOk" class="primary">Create Branch</button>
    </div>
  `);
  const input = document.getElementById("branchNameInput");
  input.focus();
  input.select();
  document.getElementById("branchNameCancel").onclick = () => {
    closeModal();
    onConfirm(null);
  };
  document.getElementById("branchNameOk").onclick = () => {
    const val = input.value.trim();
    closeModal();
    onConfirm(val || null);
  };
}
function promptRename(t) {
  openModal(`
    <h3>Rename Branch</h3>
    <input type="text" id="renameInput" value="${escapeHtml(t.name)}" />
    <div class="modal-actions">
      <button id="renameCancel">Cancel</button>
      <button id="renameOk" class="primary">Save</button>
    </div>
  `);
  const input = document.getElementById("renameInput");
  input.focus();
  input.select();
  document.getElementById("renameCancel").onclick = closeModal;
  document.getElementById("renameOk").onclick = () => {
    const val = input.value.trim();
    if (val) {
      t.name = val;
      if (t.layer.getTooltip()) t.layer.setTooltipContent(val);
      else t.layer.bindTooltip(val, { permanent: true, direction: "center", className: "branch-label" });
      renderAll();
    }
    closeModal();
  };
}

// ------------------------------------------------------------------
// Load screen
// ------------------------------------------------------------------

const loadStatusEl = document.getElementById("loadStatus");

function showLoadStatus(msg, cls) {
  loadStatusEl.textContent = msg;
  loadStatusEl.className = "load-status" + (cls ? " " + cls : "");
}

/*
 * Validation errors/warnings are never silently dropped, but they no longer
 * gate a separate pre-map screen - they're just recorded, and surfaced via
 * the small "data notes" badge in the map screen's top bar (see
 * updateWarningsBadge/dataWarningsBtn below) so the tool can go straight to
 * the map.
 */
function showValidation(errors, warnings, sellerCount) {
  STATE.lastErrors = errors;
  STATE.lastWarnings = warnings;
  STATE.lastSellerCount = sellerCount;
  updateWarningsBadge();
}

function updateWarningsBadge() {
  const btn = document.getElementById("dataWarningsBtn");
  if (!btn) return;
  const total = (STATE.lastErrors || []).length + (STATE.lastWarnings || []).length;
  if (total === 0) {
    btn.hidden = true;
    return;
  }
  btn.hidden = false;
  btn.textContent = `⚠ ${total} data note${total === 1 ? "" : "s"}`;
}

document.getElementById("dataWarningsBtn").addEventListener("click", () => {
  const errors = STATE.lastErrors || [];
  const warnings = STATE.lastWarnings || [];
  const errHtml = errors.length
    ? errors.slice(0, 500).map((e) => `<div>${escapeHtml(e)}</div>`).join("")
    : "<div>No blocking errors.</div>";
  const warnHtml = warnings.length
    ? warnings.slice(0, 500).map((w) => `<div>${escapeHtml(w)}</div>`).join("")
    : "<div>No warnings.</div>";
  openModal(`
    <h3>Data notes</h3>
    <p>${STATE.lastSellerCount || 0} sellers read &middot; ${errors.length} error(s) &middot; ${warnings.length} warning(s)</p>
    <div class="validation-list validation-errors">${errHtml}</div>
    <div class="validation-list validation-warnings">${warnHtml}</div>
    <div class="modal-actions"><button id="warnModalClose" class="primary">Close</button></div>
  `);
  document.getElementById("warnModalClose").onclick = closeModal;
});

const DEFAULT_INPUT_FILENAME = "input.xlsx";

/*
 * Shared by both the manual file picker and the automatic default-file
 * load below - parses+validates a workbook already read into memory and
 * updates the load screen. Returns true on success.
 */
function processWorkbookData(arrayBufferOrUint8, filename) {
  STATE.sourceFileName = filename;
  try {
    const data = arrayBufferOrUint8 instanceof Uint8Array ? arrayBufferOrUint8 : new Uint8Array(arrayBufferOrUint8);
    const workbook = XLSX.read(data, { type: "array", cellDates: true });
    const parsed = parseWorkbook(workbook);
    if (parsed.error) {
      showLoadStatus(parsed.error, "error");
      return false;
    }
    const built = buildSellers(parsed.sellerRows, parsed.columns);
    STATE.sellers = built.sellers;
    STATE.rawSummaryRows = parsed.summaryRows;
    STATE.rawSummaryCols = parsed.summaryCols;
    STATE.rawWorkbookBytes = data; // kept so "Export Standalone Copy" can re-embed the exact same file

    const plottable = STATE.sellers.filter((s) => s.validCoords && s.clusterId).length;
    showLoadStatus(`Loaded sheet "${parsed.sheetName}" - ${plottable} sellers ready to map.`, "ok");
    showValidation(built.errors, built.warnings, STATE.sellers.length);
    return true;
  } catch (err) {
    showLoadStatus("Failed to read file: " + err.message, "error");
    return false;
  }
}

/*
 * Loads "input.xlsx" from next to index.html - the only way this tool takes
 * in data (besides a packaged standalone copy's embedded data.js). Requires
 * the page to be served over http/https (i.e. started via start.bat, not by
 * double-clicking index.html directly - browsers block fetching sibling
 * files from a plain file:// page).
 */
async function tryLoadDefaultInputFile() {
  try {
    const resp = await fetch(DEFAULT_INPUT_FILENAME, { cache: "no-store" });
    if (!resp.ok) {
      showLoadStatus(`Could not find "${DEFAULT_INPUT_FILENAME}" in this folder (HTTP ${resp.status}).`, "error");
      return;
    }
    const buf = await resp.arrayBuffer();
    const ok = processWorkbookData(buf, DEFAULT_INPUT_FILENAME);
    if (ok) proceedToMap();
  } catch (err) {
    showLoadStatus(
      `Could not load "${DEFAULT_INPUT_FILENAME}" automatically. Make sure it exists in this folder and that ` +
        "you started the tool via start.bat (not by double-clicking index.html directly).",
      "error"
    );
  }
}

function base64FromUint8Array(bytes) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function uint8ArrayFromBase64(b64) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/*
 * A "standalone copy" is this same tool with the seller/Cluster workbook
 * baked directly into a small data.js file (as base64), instead of relying
 * on a separate input.xlsx or a file picker. Since data.js is loaded via a
 * normal <script src>, it works when the recipient just double-clicks
 * index.html - no Python, no local server, no Excel file to find, nothing
 * to install. See the README's "Sharing this tool" section.
 */
document.getElementById("exportStandaloneBtn").addEventListener("click", () => {
  if (!STATE.rawWorkbookBytes) {
    alertModal("No data loaded yet.");
    return;
  }
  const b64 = base64FromUint8Array(STATE.rawWorkbookBytes);
  const content = `// Generated by "Export Standalone Copy" - embeds the clustered seller workbook\n` +
    `// so this tool can run standalone with no Excel file, Python, or server needed.\n` +
    `window.EMBEDDED_WORKBOOK_BASE64 = "${b64}";\n`;
  const blob = new Blob([content], { type: "text/javascript" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "data.js";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  alertModal(
    'data.js downloaded. Put it in a copy of this folder (next to index.html, app.js, styles.css), ' +
      "zip that folder, and send it - see the README's \"Sharing this tool\" section for the full steps."
  );
});

/*
 * If this page was packaged as a standalone copy (data.js present, defining
 * window.EMBEDDED_WORKBOOK_BASE64), skip the load screen entirely - decode
 * and go straight to the map. Otherwise fall back to the normal input.xlsx
 * auto-load / manual picker flow.
 */
function loadEmbeddedWorkbook(base64) {
  try {
    const bytes = uint8ArrayFromBase64(base64);
    const ok = processWorkbookData(bytes, "input.xlsx");
    if (ok) proceedToMap();
  } catch (err) {
    showLoadStatus("Could not read the embedded data: " + err.message, "error");
  }
}

document.getElementById("loadProjectBtn").addEventListener("click", () => {
  document.getElementById("projectFileInput").click();
});

document.getElementById("projectFileInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const proj = JSON.parse(evt.target.result);
      restoreProject(proj);
    } catch (err) {
      alertModal("Could not read project file: " + err.message);
    }
  };
  reader.readAsText(file);
});

function proceedToMap() {
  STATE.clusters = buildClusters(STATE.sellers, STATE.rawSummaryRows, STATE.rawSummaryCols);
  STATE.branches = [];
  STATE.overrides = {};
  STATE.selectedBranchId = null;
  STATE.branchCounter = 1;

  document.getElementById("loadScreen").hidden = true;
  document.getElementById("mainScreen").hidden = false;
  document.getElementById("sourceFileLabel").textContent = STATE.sourceFileName;

  if (!STATE.map) initMap();
  renderSellerMarkers();
  renderClusterLayers();
  fitMapToData();
  recomputeAssignments();
}

// ------------------------------------------------------------------
// Map
// ------------------------------------------------------------------

function initMap() {
  STATE.map = L.map("map", { preferCanvas: true });

  // Base map: Voyager (roads, localities, landmarks) is the default - a
  // richer free/no-token basemap than the previous minimal "Positron"
  // style, from the same CARTO/OpenStreetMap source already in use.
  // OSM Standard and the old Positron style remain selectable.
  STATE.baseLayers = {
    voyager: L.tileLayer("https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png", {
      attribution: "&copy; OpenStreetMap contributors &copy; CARTO",
      subdomains: "abcd",
      maxZoom: 20,
    }),
    osm: L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "&copy; OpenStreetMap contributors",
      subdomains: "abc",
      maxZoom: 19,
    }),
    positron: L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
      attribution: "&copy; OpenStreetMap contributors &copy; CARTO",
      subdomains: "abcd",
      maxZoom: 19,
    }),
  };
  STATE.currentBaseLayer = STATE.baseLayers.voyager.addTo(STATE.map);
  document.getElementById("baseMapSelect").addEventListener("change", (e) => {
    STATE.map.removeLayer(STATE.currentBaseLayer);
    STATE.currentBaseLayer = STATE.baseLayers[e.target.value].addTo(STATE.map);
  });

  STATE.boundaryLayerGroup = L.layerGroup().addTo(STATE.map);
  STATE.centroidLayerGroup = L.layerGroup();
  STATE.eligibleCluster = L.markerClusterGroup({ disableClusteringAtZoom: 16, chunkedLoading: true, maxClusterRadius: 60 });
  STATE.ineligibleCluster = L.markerClusterGroup({ disableClusteringAtZoom: 16, chunkedLoading: true, maxClusterRadius: 60 });
  STATE.metroLayerGroup = L.layerGroup();
  STATE.metroFetchState = "idle";
  STATE.branchDrawnItems = L.featureGroup().addTo(STATE.map);
  STATE.clusterPolygons = new Map();
  STATE.clusterCentroidMarkers = new Map();

  document.getElementById("layerBoundaries").addEventListener("change", (e) => toggleLayer(STATE.boundaryLayerGroup, e.target.checked));
  document.getElementById("layerEligible").addEventListener("change", (e) => toggleLayer(STATE.eligibleCluster, e.target.checked));
  document.getElementById("layerIneligible").addEventListener("change", (e) => toggleLayer(STATE.ineligibleCluster, e.target.checked));
  document.getElementById("layerCentroids").addEventListener("change", (e) => toggleLayer(STATE.centroidLayerGroup, e.target.checked));
  document.getElementById("layerBranches").addEventListener("change", (e) => toggleLayer(STATE.branchDrawnItems, e.target.checked));
  document.getElementById("layerMetro").addEventListener("change", (e) => {
    if (e.target.checked) {
      toggleLayer(STATE.metroLayerGroup, true);
      ensureMetroDataLoaded();
    } else {
      toggleLayer(STATE.metroLayerGroup, false);
    }
  });

  STATE.map.on(L.Draw.Event.CREATED, (e) => {
    if (e.layerType !== "polygon") return;
    const layer = e.layer;
    promptBranchName("", (name) => {
      if (!name) return; // discarded
      createBranch(name, layer);
    });
  });

  // Drawing UX: show a clear instruction + live mouse coordinates only
  // while a Branch polygon is actively being drawn.
  STATE.map.on("draw:drawstart", () => {
    document.getElementById("drawHint").hidden = false;
    document.getElementById("coordIndicator").hidden = false;
  });
  STATE.map.on("draw:drawstop", () => {
    document.getElementById("drawHint").hidden = true;
    document.getElementById("coordIndicator").hidden = true;
    activeDrawHandler = null;
  });
  STATE.map.on("mousemove", (e) => {
    const el = document.getElementById("coordText");
    if (el) el.textContent = `${e.latlng.lat.toFixed(5)}, ${e.latlng.lng.toFixed(5)}`;
  });

  STATE.map.on("popupopen", (e) => {
    const layer = e.popup._source;
    if (!layer) return;
    if (layer._tctClusterId) {
      e.popup.setContent(clusterPopupHtml(layer._tctClusterId));
      wireClusterPopupHandlers(layer._tctClusterId);
      return;
    }
    if (layer.branchId) {
      const t = STATE.branches.find((t) => t.id === layer.branchId);
      if (t) e.popup.setContent(branchPopupHtml(t));
    }
  });
}

function toggleLayer(layer, show) {
  if (show) layer.addTo(STATE.map);
  else STATE.map.removeLayer(layer);
}

function fitMapToData() {
  const coords = STATE.sellers.filter((s) => s.validCoords).map((s) => [s.lat, s.lon]);
  if (coords.length) STATE.map.fitBounds(coords);
  else STATE.map.setView([28.6139, 77.209], 10);
}

function sellerPopupHtml(s) {
  return [
    `<b>GLID:</b> ${escapeHtml(s.glid)}`,
    `<b>Cluster ID:</b> ${escapeHtml(s.clusterId || "N/A")}`,
    `<b>Eligibility:</b> ${escapeHtml(s.eligibilityRaw || (s.eligible ? "Eligible" : "Not Eligible"))}`,
    `<b>Outstanding Amount:</b> ${s.outstanding != null ? s.outstanding : "N/A"}`,
    `<b>Renewal Date:</b> ${formatDateVal(s.renewal)}`,
    `<b>Latitude:</b> ${s.lat.toFixed(5)}`,
    `<b>Longitude:</b> ${s.lon.toFixed(5)}`,
  ].join("<br/>");
}

function renderSellerMarkers() {
  STATE.eligibleCluster.clearLayers();
  STATE.ineligibleCluster.clearLayers();
  for (const s of STATE.sellers) {
    if (!s.validCoords) continue;
    const color = s.eligible ? "#1f78b4" : "#808080";
    const marker = L.circleMarker([s.lat, s.lon], {
      radius: 3,
      color,
      weight: 1,
      fillColor: color,
      fillOpacity: s.eligible ? 0.85 : 0.6,
    });
    marker.bindPopup(sellerPopupHtml(s), { maxWidth: 270 });
    (s.eligible ? STATE.eligibleCluster : STATE.ineligibleCluster).addLayer(marker);
  }
}

function starIcon() {
  return L.divIcon({ html: "&#9733;", className: "star-icon", iconSize: [16, 16], iconAnchor: [8, 8] });
}

function metroIcon() {
  return L.divIcon({ html: "M", className: "metro-icon", iconSize: [16, 16], iconAnchor: [8, 8] });
}

function computeDataBbox() {
  const valid = STATE.sellers.filter((s) => s.validCoords);
  if (!valid.length) return null;
  let south = 90, north = -90, west = 180, east = -180;
  for (const s of valid) {
    if (s.lat < south) south = s.lat;
    if (s.lat > north) north = s.lat;
    if (s.lon < west) west = s.lon;
    if (s.lon > east) east = s.lon;
  }
  const margin = 0.05; // ~5km, so nearby stations just outside the data extent still show
  return { south: south - margin, west: west - margin, north: north + margin, east: east + margin };
}

function updateMetroStatus(msg) {
  const el = document.getElementById("metroStatus");
  if (el) el.textContent = msg;
}

function renderMetroStations(elements) {
  STATE.metroLayerGroup.clearLayers();
  for (const el of elements) {
    if (el.lat == null || el.lon == null) continue;
    const tags = el.tags || {};
    const name = tags.name || tags["name:en"] || "Metro Station";
    const network = tags.network || "";
    const line = tags.line || tags.route || "";
    const marker = L.marker([el.lat, el.lon], { icon: metroIcon() });
    const details = [`<b>${escapeHtml(name)}</b>`];
    if (network) details.push(`Network: ${escapeHtml(network)}`);
    if (line) details.push(`Line: ${escapeHtml(line)}`);
    marker.bindPopup(details.join("<br/>"), { maxWidth: 220 });
    marker.bindTooltip(name, { direction: "top", sticky: true });
    STATE.metroLayerGroup.addLayer(marker);
  }
}

/*
 * Metro Stations layer: fetched lazily (only when the layer is first
 * switched on) from the public Overpass API (https://overpass-api.de),
 * which serves OpenStreetMap data - free, no API key/token required.
 * Station name/network/line come straight from whatever OpenStreetMap
 * has tagged; nothing is fabricated, and if the request fails (e.g. no
 * network access, or the public Overpass instance is rate-limiting) the
 * layer simply stays empty with a status message - Branch drawing and
 * every other feature keep working normally either way.
 */
async function ensureMetroDataLoaded() {
  if (STATE.metroFetchState === "done" || STATE.metroFetchState === "loading") return;
  const bbox = computeDataBbox();
  if (!bbox) return;

  STATE.metroFetchState = "loading";
  updateMetroStatus("Loading metro stations from OpenStreetMap (Overpass API)...");

  const query =
    `[out:json][timeout:25];` +
    `(node["railway"="station"]["station"="subway"](${bbox.south},${bbox.west},${bbox.north},${bbox.east}););` +
    `out body;`;

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 15000);
  try {
    const resp = await fetch("https://overpass-api.de/api/interpreter", {
      method: "POST",
      body: "data=" + encodeURIComponent(query),
      signal: controller.signal,
    });
    clearTimeout(timeoutId);
    if (!resp.ok) throw new Error("Overpass API returned HTTP " + resp.status);
    const data = await resp.json();
    const elements = data.elements || [];
    renderMetroStations(elements);
    STATE.metroFetchState = "done";
    updateMetroStatus(
      elements.length
        ? `${elements.length} metro station(s) loaded (OpenStreetMap / Overpass API).`
        : "No metro stations found in OpenStreetMap for this area."
    );
  } catch (err) {
    clearTimeout(timeoutId);
    STATE.metroFetchState = "error";
    updateMetroStatus("Metro station data unavailable right now (network or Overpass API issue). Everything else is unaffected.");
  }
}

function renderClusterLayers() {
  STATE.boundaryLayerGroup.clearLayers();
  STATE.centroidLayerGroup.clearLayers();
  STATE.clusterPolygons.clear();
  STATE.clusterCentroidMarkers.clear();

  for (const [cid, c] of STATE.clusters) {
    // Solid blue, medium-thickness outline for EVERY Cluster boundary,
    // always - this is the fixed geographic reference layer and must stay
    // visually distinct from Branch polygons (see restyleClusters()).
    const poly = L.polygon(c.hullLatLngs, { color: "#1f78b4", weight: 2, fillColor: "#1f78b4", fillOpacity: 0.05 });
    poly._tctClusterId = cid;
    poly.bindTooltip(
      `Cluster ${escapeHtml(cid)} &mdash; ${c.sellers.length} sellers (${c.eligibleCount} elig. / ${c.ineligibleCount} inelig.)`,
      { sticky: true, direction: "top" }
    );
    poly.bindPopup("", { maxWidth: 280 });
    poly.addTo(STATE.boundaryLayerGroup);
    STATE.clusterPolygons.set(cid, poly);

    const marker = L.marker(c.centroid, { icon: starIcon() });
    marker._tctClusterId = cid;
    marker.bindPopup("", { maxWidth: 280 });
    marker.addTo(STATE.centroidLayerGroup);
    STATE.clusterCentroidMarkers.set(cid, marker);
  }
}

function clusterPopupHtml(cid) {
  const c = STATE.clusters.get(cid);
  const a = STATE.assignments[cid] || { status: "UNMAPPED" };
  let statusLine;
  if (a.status === "MAPPED") {
    const t = STATE.branches.find((t) => t.id === a.branchId);
    statusLine = `<b>Branch:</b> ${escapeHtml(t ? t.name : "Unknown")}${a.manual ? " (manual)" : ""}`;
  } else if (a.status === "CONFLICT") {
    const names = a.candidates.map((id) => STATE.branches.find((t) => t.id === id)).filter(Boolean).map((t) => t.name);
    statusLine = `<b>Branch:</b> <span style="color:#c0392b">CONFLICT</span> (${escapeHtml(names.join(", "))})`;
  } else {
    statusLine = `<b>Branch:</b> UNMAPPED`;
  }

  const idSuffix = cssId(cid);
  const options = STATE.branches.map((t) => `<option value="${escapeHtml(t.id)}">${escapeHtml(t.name)}</option>`).join("");
  const assignBlock =
    STATE.branches.length > 0
      ? `<select id="assignSel_${idSuffix}">${options}</select><button id="assignBtn_${idSuffix}">Assign to selected Branch</button>`
      : `<div style="color:#777">No Branches yet.</div>`;
  const unassignBlock =
    a.status !== "UNMAPPED" ? `<button id="unassignBtn_${idSuffix}">Unassign (mark Unmapped)</button>` : "";

  return `
    <div>
      <b>Cluster ID:</b> ${escapeHtml(cid)}<br/>
      ${statusLine}<br/>
      <b>Total Sellers:</b> ${c.sellers.length}<br/>
      <b>Eligible Sellers:</b> ${c.eligibleCount}<br/>
      <b>Ineligible Sellers:</b> ${c.ineligibleCount}
      <div class="popup-actions">
        ${assignBlock}
        ${unassignBlock}
      </div>
    </div>
  `;
}

function branchPopupHtml(t) {
  const totals = branchTotals(t);
  return [
    `<b>Branch:</b> ${escapeHtml(t.name)}`,
    `<b>Clusters:</b> ${t.clusterIds.length}`,
    `<b>Sellers:</b> ${totals.sellers}`,
    `<b>Eligible Sellers:</b> ${totals.eligible}`,
    `<b>Ineligible Sellers:</b> ${totals.ineligible}`,
  ].join("<br/>");
}

function reopenPopupContent(cid) {
  const poly = STATE.clusterPolygons.get(cid);
  if (poly && poly.isPopupOpen()) {
    poly.getPopup().setContent(clusterPopupHtml(cid));
    wireClusterPopupHandlers(cid);
  }
  const marker = STATE.clusterCentroidMarkers.get(cid);
  if (marker && marker.isPopupOpen()) {
    marker.getPopup().setContent(clusterPopupHtml(cid));
    wireClusterPopupHandlers(cid);
  }
}

function wireClusterPopupHandlers(cid) {
  const idSuffix = cssId(cid);
  const assignBtn = document.getElementById("assignBtn_" + idSuffix);
  const unassignBtn = document.getElementById("unassignBtn_" + idSuffix);
  if (assignBtn) {
    assignBtn.onclick = () => {
      const sel = document.getElementById("assignSel_" + idSuffix);
      if (sel && sel.value) {
        STATE.overrides[cid] = sel.value;
        recomputeAssignments();
        reopenPopupContent(cid);
      }
    };
  }
  if (unassignBtn) {
    unassignBtn.onclick = () => {
      STATE.overrides[cid] = "UNMAPPED";
      recomputeAssignments();
      reopenPopupContent(cid);
    };
  }
}

function restyleClusters() {
  // The Cluster boundary OUTLINE is always solid blue, medium thickness -
  // Clusters are the fixed geographic reference layer and must stay
  // visually distinct from Branch polygons at all times (see index.html
  // change notes). Only the fill tint communicates assignment status.
  for (const [cid, poly] of STATE.clusterPolygons) {
    const a = STATE.assignments[cid] || { status: "UNMAPPED" };
    let fillColor = "#1f78b4";
    let fillOpacity = 0.05;
    if (a.status === "MAPPED") {
      const t = STATE.branches.find((t) => t.id === a.branchId);
      fillColor = t ? t.color : "#1f78b4";
      fillOpacity = 0.25;
    } else if (a.status === "CONFLICT") {
      fillColor = "#e6194b";
      fillOpacity = 0.35;
    }
    poly.setStyle({ color: "#1f78b4", weight: 2, fillColor, fillOpacity, dashArray: null });
  }
}

// ------------------------------------------------------------------
// Branch management
// ------------------------------------------------------------------

function updateBranchGeometry(t) {
  const geojson = t.layer.toGeoJSON();
  t.turfPoly = geojson;
  t.bbox = turf.bbox(geojson);
}

function createBranch(name, layer) {
  const id = "B" + STATE.branchCounter++;
  const color = nextColor(STATE.branches.length);
  // Completed Branch: solid outline in the Branch's own colour +
  // a clearly visible transparent fill - distinct from both the orange
  // "being drawn" style and the fixed blue Cluster boundary outline.
  layer.setStyle({ color, weight: 3, dashArray: null, fillColor: color, fillOpacity: 0.12 });
  layer.bindTooltip(name, { permanent: true, direction: "center", className: "branch-label" });
  layer.bindPopup("", { maxWidth: 260 });
  layer.branchId = id;
  STATE.branchDrawnItems.addLayer(layer);

  const t = { id, name, color, layer, clusterIds: [] };
  updateBranchGeometry(t);
  STATE.branches.push(t);
  STATE.selectedBranchId = id;
  afterDrawChange();
}

function afterDrawChange() {
  recomputeAssignments();
  const conflicts = Object.entries(STATE.assignments).filter(([, a]) => a.status === "CONFLICT");
  if (conflicts.length > 0) {
    const ids = conflicts.map(([cid]) => cid).join(", ");
    alertModal(
      `Some Clusters fall into more than one Branch:\n${ids}\n\nResolve these in the "Conflicts" panel (or a Cluster's map popup) before exporting.`
    );
  }
}

let activeEditHandler = null;
let activeEditBranch = null;
let activeDrawHandler = null;

document.getElementById("drawBranchBtn").addEventListener("click", () => {
  if (activeEditHandler) {
    alertModal("Finish the current edit first (Save or Cancel).");
    return;
  }
  if (activeDrawHandler) {
    alertModal("Finish or cancel the Branch you're currently drawing first.");
    return;
  }
  // "Being drawn" style: darker/thicker outline + light fill, clearly
  // different from both the blue Cluster boundary and a completed Branch.
  activeDrawHandler = new L.Draw.Polygon(STATE.map, {
    shapeOptions: { color: "#d35400", weight: 4, fillColor: "#d35400", fillOpacity: 0.15 },
    showArea: false,
  });
  activeDrawHandler.enable();
});

document.getElementById("cancelDrawBtn").addEventListener("click", () => {
  if (activeDrawHandler) activeDrawHandler.disable();
});

document.getElementById("editBranchBtn").addEventListener("click", () => {
  const btn = document.getElementById("editBranchBtn");
  if (activeDrawHandler) {
    alertModal("Finish or cancel the Branch you're currently drawing first.");
    return;
  }
  if (activeEditHandler) {
    activeEditHandler.save();
    activeEditHandler.disable();
    activeEditHandler = null;
    updateBranchGeometry(activeEditBranch);
    activeEditBranch = null;
    btn.textContent = "Edit Branch";
    btn.classList.remove("active-mode");
    afterDrawChange();
    return;
  }
  if (!STATE.selectedBranchId) {
    alertModal("Select a Branch from the list first, then click Edit Branch.");
    return;
  }
  const t = STATE.branches.find((t) => t.id === STATE.selectedBranchId);
  activeEditBranch = t;
  activeEditHandler = new L.EditToolbar.Edit(STATE.map, { featureGroup: L.featureGroup([t.layer]) });
  activeEditHandler.enable();
  btn.textContent = "Save Branch Edits";
  btn.classList.add("active-mode");
});

document.getElementById("deleteBranchBtn").addEventListener("click", () => {
  if (activeDrawHandler) {
    alertModal("Finish or cancel the Branch you're currently drawing first.");
    return;
  }
  if (!STATE.selectedBranchId) {
    alertModal("Select a Branch from the list first, then click Delete Branch.");
    return;
  }
  const t = STATE.branches.find((t) => t.id === STATE.selectedBranchId);
  confirmModal(`Delete Branch "${t.name}"? Its Clusters will become Unmapped unless individually re-assigned.`, () => {
    STATE.branchDrawnItems.removeLayer(t.layer);
    STATE.branches = STATE.branches.filter((x) => x.id !== t.id);
    for (const cid of Object.keys(STATE.overrides)) {
      if (STATE.overrides[cid] === t.id) delete STATE.overrides[cid];
    }
    STATE.selectedBranchId = null;
    recomputeAssignments();
  });
});

document.getElementById("resetBtn").addEventListener("click", () => {
  confirmModal(
    "Reset Branch Mapping? This removes all Branches and manual assignments created in this tool. The underlying Cluster data is not affected.",
    () => {
      for (const t of STATE.branches) STATE.branchDrawnItems.removeLayer(t.layer);
      STATE.branches = [];
      STATE.overrides = {};
      STATE.selectedBranchId = null;
      recomputeAssignments();
    }
  );
});

document.getElementById("assignmentRule").addEventListener("change", (e) => {
  STATE.assignmentRule = e.target.value;
  recomputeAssignments();
});

// ------------------------------------------------------------------
// Assignment computation
// ------------------------------------------------------------------

function majorityAreaInside(cluster, branch) {
  try {
    const inter = turf.intersect(cluster.turfPoly, branch.turfPoly);
    if (!inter) return false;
    return turf.area(inter) / cluster.area > 0.5;
  } catch (e) {
    return false;
  }
}

function majoritySellersInside(cluster, branch) {
  let count = 0;
  for (const s of cluster.sellers) {
    if (turf.booleanPointInPolygon([s.lon, s.lat], branch.turfPoly)) count++;
  }
  return count / cluster.sellers.length > 0.5;
}

function recomputeAssignments() {
  const rule = STATE.assignmentRule;
  const newAssignments = {};

  for (const [cid, cluster] of STATE.clusters) {
    if (Object.prototype.hasOwnProperty.call(STATE.overrides, cid)) {
      const ov = STATE.overrides[cid];
      newAssignments[cid] = ov === "UNMAPPED" ? { status: "UNMAPPED" } : { status: "MAPPED", branchId: ov, manual: true };
      continue;
    }

    const candidates = [];
    for (const t of STATE.branches) {
      if (!bboxIntersects(cluster.bbox, t.bbox)) continue;
      let matched = false;
      if (rule === "A") matched = turf.booleanPointInPolygon(cluster.centroidPoint, t.turfPoly);
      else if (rule === "B") matched = majorityAreaInside(cluster, t);
      else matched = majoritySellersInside(cluster, t);
      if (matched) candidates.push(t.id);
    }

    if (candidates.length === 0) newAssignments[cid] = { status: "UNMAPPED" };
    else if (candidates.length === 1) newAssignments[cid] = { status: "MAPPED", branchId: candidates[0] };
    else newAssignments[cid] = { status: "CONFLICT", candidates };
  }

  STATE.assignments = newAssignments;
  for (const t of STATE.branches) t.clusterIds = [];
  for (const [cid, a] of Object.entries(newAssignments)) {
    if (a.status === "MAPPED") {
      const t = STATE.branches.find((t) => t.id === a.branchId);
      if (t) t.clusterIds.push(cid);
    }
  }

  renderAll();
}

// ------------------------------------------------------------------
// Side panel rendering
// ------------------------------------------------------------------

function branchTotals(t) {
  let sellers = 0, eligible = 0, ineligible = 0;
  for (const cid of t.clusterIds) {
    const c = STATE.clusters.get(cid);
    sellers += c.sellers.length;
    eligible += c.eligibleCount;
    ineligible += c.ineligibleCount;
  }
  return { sellers, eligible, ineligible };
}

function renderOverview() {
  const totalClusters = STATE.clusters.size;
  const mapped = Object.values(STATE.assignments).filter((a) => a.status === "MAPPED").length;
  const unmapped = Object.values(STATE.assignments).filter((a) => a.status === "UNMAPPED").length;
  const conflicts = Object.values(STATE.assignments).filter((a) => a.status === "CONFLICT").length;
  const totalSellers = STATE.sellers.filter((s) => s.validCoords && s.clusterId).length;

  document.getElementById("overviewStats").innerHTML = `
    <div class="stat"><b>${totalClusters}</b>Total Clusters</div>
    <div class="stat"><b>${mapped}</b>Mapped</div>
    <div class="stat"><b>${unmapped}</b>Unmapped</div>
    <div class="stat"><b>${conflicts}</b>Conflicts</div>
    <div class="stat"><b>${STATE.branches.length}</b>Branches</div>
    <div class="stat"><b>${totalSellers}</b>Sellers (clustered)</div>
  `;
}

function renderBranchList() {
  const el = document.getElementById("branchList");
  el.innerHTML = "";
  if (STATE.branches.length === 0) {
    el.innerHTML = '<div class="list-row">No Branches yet - use "Draw Branch".</div>';
    return;
  }
  for (const t of STATE.branches) {
    const totals = branchTotals(t);
    const row = document.createElement("div");
    row.className = "list-row" + (STATE.selectedBranchId === t.id ? " selected" : "");
    row.innerHTML = `
      <span><span class="swatch" style="background:${t.color}"></span>${escapeHtml(t.name)}</span>
      <span class="count">${t.clusterIds.length} clusters &middot; ${totals.sellers} sellers</span>
    `;
    row.title = "Click to select. Double-click to rename.";
    row.onclick = () => {
      STATE.selectedBranchId = t.id;
      STATE.map.fitBounds(t.layer.getBounds(), { maxZoom: 14 });
      renderBranchList();
    };
    row.ondblclick = () => promptRename(t);
    el.appendChild(row);
  }
}

function renderUnmappedList() {
  const el = document.getElementById("unmappedList");
  const unmapped = [...STATE.clusters.keys()].filter((cid) => (STATE.assignments[cid] || {}).status === "UNMAPPED");
  el.innerHTML = "";
  if (unmapped.length === 0) {
    el.innerHTML = '<div class="list-row">None &mdash; every Cluster is mapped.</div>';
    return;
  }
  for (const cid of unmapped) {
    const c = STATE.clusters.get(cid);
    const row = document.createElement("div");
    row.className = "list-row";
    row.innerHTML = `<span>${escapeHtml(cid)}</span><span class="count">${c.sellers.length} sellers &middot; ${c.eligibleCount} elig. / ${c.ineligibleCount} inelig.</span>`;
    row.onclick = () => {
      document.getElementById("layerBoundaries").checked = true;
      toggleLayer(STATE.boundaryLayerGroup, true);
      const poly = STATE.clusterPolygons.get(cid);
      STATE.map.fitBounds(poly.getBounds(), { maxZoom: 15 });
      poly.openPopup();
    };
    el.appendChild(row);
  }
}

function renderConflictList() {
  const section = document.getElementById("conflictSection");
  const el = document.getElementById("conflictList");
  const conflicts = Object.entries(STATE.assignments).filter(([, a]) => a.status === "CONFLICT");
  if (conflicts.length === 0) {
    section.hidden = true;
    el.innerHTML = "";
    return;
  }
  section.hidden = false;
  el.innerHTML = "";
  for (const [cid, a] of conflicts) {
    const names = a.candidates.map((id) => STATE.branches.find((t) => t.id === id)).filter(Boolean);
    const row = document.createElement("div");
    row.className = "list-row";
    row.style.display = "block";
    const btns = names
      .map((t) => `<button data-cid="${escapeHtml(cid)}" data-tid="${escapeHtml(t.id)}" class="resolveBtn">${escapeHtml(t.name)}</button>`)
      .join(" ");
    row.innerHTML = `
      <div><b>${escapeHtml(cid)}</b> is in ${names.length} Branches</div>
      <div class="popup-actions">${btns} <button data-cid="${escapeHtml(cid)}" data-tid="UNMAPPED" class="resolveBtn">Leave Unmapped</button></div>
    `;
    el.appendChild(row);
  }
  el.querySelectorAll(".resolveBtn").forEach((btn) => {
    btn.onclick = () => {
      STATE.overrides[btn.dataset.cid] = btn.dataset.tid;
      recomputeAssignments();
    };
  });
}

function renderSummaryTable() {
  const tbody = document.querySelector("#summaryTable tbody");
  tbody.innerHTML = "";
  for (const t of STATE.branches) {
    const totals = branchTotals(t);
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${escapeHtml(t.name)}</td><td>${t.clusterIds.length}</td><td>${totals.sellers}</td><td>${totals.eligible}</td><td>${totals.ineligible}</td>`;
    tbody.appendChild(tr);
  }
}

function renderAll() {
  restyleClusters();
  renderOverview();
  renderBranchList();
  renderUnmappedList();
  renderConflictList();
  renderSummaryTable();
}

// ------------------------------------------------------------------
// Save / Load project
// ------------------------------------------------------------------

document.getElementById("saveProjectBtn").addEventListener("click", () => {
  const project = {
    version: 1,
    sourceFileName: STATE.sourceFileName,
    assignmentRule: STATE.assignmentRule,
    branchCounter: STATE.branchCounter,
    branches: STATE.branches.map((t) => ({ id: t.id, name: t.name, color: t.color, geojson: t.layer.toGeoJSON() })),
    overrides: STATE.overrides,
  };
  const blob = new Blob([JSON.stringify(project, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  a.href = url;
  a.download = `branch_project_${stamp}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

function restoreProject(proj) {
  STATE.assignmentRule = proj.assignmentRule || "C";
  document.getElementById("assignmentRule").value = STATE.assignmentRule;
  STATE.branchCounter = proj.branchCounter || (proj.branches ? proj.branches.length + 1 : 1);

  for (const pt of proj.branches || []) {
    const layer = L.geoJSON(pt.geojson).getLayers()[0];
    layer.setStyle({ color: pt.color, weight: 3, dashArray: null, fillColor: pt.color, fillOpacity: 0.12 });
    layer.bindTooltip(pt.name, { permanent: true, direction: "center", className: "branch-label" });
    layer.bindPopup("", { maxWidth: 260 });
    layer.branchId = pt.id;
    STATE.branchDrawnItems.addLayer(layer);
    const t = { id: pt.id, name: pt.name, color: pt.color, layer, clusterIds: [] };
    updateBranchGeometry(t);
    STATE.branches.push(t);
  }
  STATE.overrides = proj.overrides || {};
  recomputeAssignments();
}

// ------------------------------------------------------------------
// Excel export
// ------------------------------------------------------------------

document.getElementById("downloadExcelBtn").addEventListener("click", () => {
  const unmappedCount = Object.values(STATE.assignments).filter((a) => a.status === "UNMAPPED").length;
  const conflictCount = Object.values(STATE.assignments).filter((a) => a.status === "CONFLICT").length;

  if (conflictCount > 0) {
    alertModal(`${conflictCount} Cluster(s) fall into more than one Branch. Resolve these conflicts before exporting.`);
    return;
  }
  if (unmappedCount > 0) {
    confirmModal(`${unmappedCount} Cluster(s) are still Unmapped. Export the mapping anyway?`, doExport);
  } else {
    doExport();
  }
});

function doExport() {
  const clusterBranchName = {};
  for (const [cid, a] of Object.entries(STATE.assignments)) {
    clusterBranchName[cid] = a.status === "MAPPED" ? STATE.branches.find((t) => t.id === a.branchId).name : "UNMAPPED";
  }

  const sheet1 = STATE.sellers.map((s) => ({
    GLID: s.glid,
    Latitude: s.lat,
    Longitude: s.lon,
    "Outstanding Amount": s.outstanding,
    "Renewal Date": s.renewal,
    Eligibility: s.eligibilityRaw || (s.eligible ? "Eligible" : "Not Eligible"),
    "Cluster ID": s.clusterId || "",
    "Branch": s.clusterId ? clusterBranchName[s.clusterId] || "UNMAPPED" : "",
  }));

  const sheet2 = [...STATE.clusters.values()].map((c) => {
    const row = {
      "Cluster ID": c.id,
      "Branch": clusterBranchName[c.id] || "UNMAPPED",
      "Total Sellers": c.sellers.length,
      "Eligible Sellers": c.eligibleCount,
      "Ineligible Sellers": c.ineligibleCount,
      "Cluster Centroid Latitude": c.centroid[0],
      "Cluster Centroid Longitude": c.centroid[1],
    };
    if (c.extra) {
      if (c.extra.avgDist != null) row["Average Distance (km)"] = c.extra.avgDist;
      if (c.extra.spread != null) row["Geographic Spread (km)"] = c.extra.spread;
      if (c.extra.maxDist != null) row["Maximum Distance (km)"] = c.extra.maxDist;
      if (c.extra.minDist != null) row["Minimum Distance (km)"] = c.extra.minDist;
    }
    return row;
  });

  const sheet3 = STATE.branches.map((t) => {
    const totals = branchTotals(t);
    return {
      "Branch": t.name,
      "Number of Clusters": t.clusterIds.length,
      "Total Sellers": totals.sellers,
      "Eligible Sellers": totals.eligible,
      "Ineligible Sellers": totals.ineligible,
    };
  });

  const sheet4 = [...STATE.clusters.values()]
    .filter((c) => (STATE.assignments[c.id] || {}).status === "UNMAPPED")
    .map((c) => ({
      "Cluster ID": c.id,
      "Total Sellers": c.sellers.length,
      "Eligible Sellers": c.eligibleCount,
      "Ineligible Sellers": c.ineligibleCount,
    }));

  const ruleLabel = { A: "Centroid inside polygon", B: "Majority of cluster area inside polygon", C: "Majority of sellers inside polygon" }[
    STATE.assignmentRule
  ];
  const sheet5 = [
    { Parameter: "Input file", Value: STATE.sourceFileName },
    { Parameter: "Number of sellers", Value: STATE.sellers.length },
    { Parameter: "Number of clusters", Value: STATE.clusters.size },
    { Parameter: "Number of branches", Value: STATE.branches.length },
    { Parameter: "Mapped clusters", Value: [...STATE.clusters.keys()].filter((cid) => (STATE.assignments[cid] || {}).status === "MAPPED").length },
    { Parameter: "Unmapped clusters", Value: [...STATE.clusters.keys()].filter((cid) => (STATE.assignments[cid] || {}).status === "UNMAPPED").length },
    { Parameter: "Cluster assignment method", Value: ruleLabel },
    { Parameter: "Date/time of export", Value: new Date().toString() },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet1), "Seller Mapping");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet2), "Cluster Mapping");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet3), "Branch Summary");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet4), "Unmapped Clusters");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet5), "Run Configuration");

  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  XLSX.writeFile(wb, `branch_mapping_${stamp}.xlsx`);
}

// ------------------------------------------------------------------
// Startup
// ------------------------------------------------------------------

if (typeof window.EMBEDDED_WORKBOOK_BASE64 === "string" && window.EMBEDDED_WORKBOOK_BASE64) {
  loadEmbeddedWorkbook(window.EMBEDDED_WORKBOOK_BASE64);
} else {
  tryLoadDefaultInputFile();
}
