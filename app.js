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
  // Current/serviced Branch for this GLID, as supplied directly in the
  // seller input - this is the source of truth for seller-branch identity.
  // Never inferred from coordinates, nearest office, pincode, Cluster, or
  // any drawn polygon.
  BRANCH: ["branch"],
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

// Branch Master (office locations) and Branch-Pincode Mapping (current
// service areas) - two new, independent reference inputs. Neither touches
// Cluster data or the existing Branch-drawing/assignment logic in any way.
const BRANCH_MASTER_ALIASES = {
  BRANCH: ["branch"],
  OFFICE_LAT: ["branch office latitude", "office latitude"],
  OFFICE_LON: ["branch office longitude", "office longitude"],
};
const BRANCH_MASTER_OPTIONAL_ALIASES = {
  ADDRESS: ["address", "office address"],
};
const BRANCH_PINCODE_ALIASES = {
  PINCODE: ["pincode"],
  BRANCH: ["location", "branch"],
};

/*
 * Branch names differ between Branch Master ("Daryaganj") and the
 * Branch-Pincode Mapping ("NCR/Delhi-Daryaganj 1") - they are not the same
 * string. Per instruction this tool never guesses that pairing
 * algorithmically; this exact 10-pair list was confirmed by the business
 * user on 2026-09-21. Update it here if branch names in either source file
 * change - any name on either side that isn't an exact match AND isn't in
 * this list is flagged as unmatched, never silently paired.
 */
const CONFIRMED_BRANCH_NAME_PAIRS = [
  ["Daryaganj", "NCR/Delhi-Daryaganj 1"],
  ["Peeragarhi", "NCR/Delhi-Peeragarhi 2"],
  ["Rohini-2", "NCR/Delhi-Rohini CSD"],
  ["Okhla", "NCR/Delhi-Okhla"],
  ["Sector-63,Noida", "NCR/Noida Sec-63"],
  ["Ghaziabad", "NCR/Ghaziabad"],
  ["Faridabad", "NCR/Faridabad CSD"],
  ["Gurgaon", "NCR/Gurgaon"],
  ["Rajendra Place", "NCR/Delhi-Rajendra Place"],
  ["Shahdara", "NCR/Delhi-Shahdara"],
];

function normalizeBranchKey(s) {
  return String(s || "").trim().toLowerCase();
}

const MASTER_TO_MAPPING_NAME = new Map(CONFIRMED_BRANCH_NAME_PAIRS.map(([m, p]) => [normalizeBranchKey(m), p]));
const MAPPING_TO_MASTER_NAME = new Map(CONFIRMED_BRANCH_NAME_PAIRS.map(([m, p]) => [normalizeBranchKey(p), m]));

/*
 * Central branch colour/shape configuration - the single source of truth
 * for "current Branch" colour-coding across the seller/GLID markers, the
 * Current Branch Service Area polygons, the legend, and popups. Keyed by
 * the Branch name as it appears in the seller input's Branch column /
 * Branch-Pincode Mapping (e.g. "NCR/Delhi-Daryaganj 1"). To change a
 * colour, edit it here only - nowhere else references a hex value.
 * "shape" only picks the fallback marker glyph (pin/circle/diamond/square);
 * colour, not shape, is what carries meaning, per instruction.
 */
const BRANCH_COLOR_CONFIG = {
  "NCR/Faridabad CSD": { color: "#e91e63", label: "Pink", shape: "pin" },
  "NCR/Delhi-Peeragarhi 2": { color: "#1e88e5", label: "Blue", shape: "square" },
  "NCR/Delhi-Daryaganj 1": { color: "#2e7d32", label: "Green", shape: "diamond" },
  "NCR/Delhi-Rohini CSD": { color: "#8e24aa", label: "Purple", shape: "pin" },
  "NCR/Gurgaon": { color: "#e53935", label: "Red", shape: "pin" },
  "NCR/Delhi-Shahdara": { color: "#5e35b1", label: "Violet", shape: "diamond" },
  "NCR/Delhi-Rajendra Place": { color: "#ad1457", label: "Magenta", shape: "pin" },
  "NCR/Ghaziabad": { color: "#7cb342", label: "Light Green", shape: "pin" },
  "NCR/Noida Sec-63": { color: "#fb8c00", label: "Orange", shape: "pin" },
  "NCR/Delhi-Okhla": { color: "#3949ab", label: "Indigo", shape: "circle" },
  "NCR/Meerut": { color: "#f9a825", label: "Yellow", shape: "pin" },
  "NCR/Panipat": { color: "#6d4c41", label: "Brown", shape: "pin" },
  "NCR/Delhi-Daryaganj 2": { color: "#ff8f00", label: "Amber", shape: "pin" },
};
const OTHER_BRANCH_KEY = "Other / No data";
const OTHER_BRANCH_ENTRY = { color: "#9e9e9e", label: "Grey", shape: "pin" };

const BRANCH_COLOR_MAP = new Map(
  Object.entries(BRANCH_COLOR_CONFIG).map(([name, entry]) => [normalizeBranchKey(name), { name, ...entry }])
);

/*
 * Resolves a Branch name (or null/blank, meaning "no Branch on this
 * record") to its configured colour/shape/display-name. A Branch name
 * that IS present but has no entry in BRANCH_COLOR_CONFIG is never
 * dropped - it renders in the same grey as "Other / No data" and is
 * surfaced in the validation summary as "without a configured colour",
 * so nothing on the map silently loses its identity.
 */
function getBranchColorEntry(branchName) {
  if (!branchName) return { name: OTHER_BRANCH_KEY, ...OTHER_BRANCH_ENTRY, configured: true };
  const hit = BRANCH_COLOR_MAP.get(normalizeBranchKey(branchName));
  if (hit) return { ...hit, configured: true };
  return { name: branchName, ...OTHER_BRANCH_ENTRY, configured: false };
}

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
  let missingBranchCount = 0;

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

    const branchRaw = cols.BRANCH ? row[cols.BRANCH] : null;
    const branch = branchRaw === null || branchRaw === undefined || String(branchRaw).trim() === "" ? null : String(branchRaw).trim();
    if (!branch) missingBranchCount++;

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
      branch, // current/serviced Branch for this GLID, straight from the input - null means "Other / No data"
      centroidLat: cols.CENTROID_LAT && row[cols.CENTROID_LAT] != null ? Number(row[cols.CENTROID_LAT]) : null,
      centroidLon: cols.CENTROID_LON && row[cols.CENTROID_LON] != null ? Number(row[cols.CENTROID_LON]) : null,
    });
  });

  if (!cols.BRANCH) {
    warnings.push('Seller input has no "Branch" column - every seller is shown as "Other / No data".');
  } else if (missingBranchCount > 0) {
    warnings.push(`${missingBranchCount} seller(s) have no Branch value in the input - shown as "Other / No data".`);
  }

  return { sellers, errors, warnings };
}

/*
 * "Total GLIDs / with Branch / without Branch / total (current) Branches /
 * Branches without a configured colour" - the validation-summary numbers
 * required whenever the seller-Branch colouring feature is in use. Purely
 * informational: never drops or reassigns a GLID or a Branch.
 */
function computeBranchValidationSummary(sellers) {
  const totalGlids = sellers.length;
  let withBranch = 0;
  const distinctBranches = new Set();
  for (const s of sellers) {
    if (s.branch) {
      withBranch++;
      distinctBranches.add(s.branch);
    }
  }
  const branchesWithoutColor = [...distinctBranches].filter((b) => !BRANCH_COLOR_MAP.has(normalizeBranchKey(b)));
  return {
    totalGlids,
    withBranch,
    withoutBranch: totalGlids - withBranch,
    totalBranches: distinctBranches.size,
    branchesWithoutColor,
  };
}

// ------------------------------------------------------------------
// Branch Master + Branch-Pincode Mapping parsing
// ------------------------------------------------------------------

function resolveBranchMasterColumns(headers) {
  const map = {};
  const missing = [];
  for (const [key, aliases] of Object.entries(BRANCH_MASTER_ALIASES)) {
    const col = resolveColumn(headers, aliases);
    if (!col) missing.push(key);
    else map[key] = col;
  }
  for (const [key, aliases] of Object.entries(BRANCH_MASTER_OPTIONAL_ALIASES)) {
    const col = resolveColumn(headers, aliases);
    if (col) map[key] = col;
  }
  return { ok: missing.length === 0, map, missing };
}

function resolveBranchPincodeColumns(headers) {
  const map = {};
  const missing = [];
  for (const [key, aliases] of Object.entries(BRANCH_PINCODE_ALIASES)) {
    const col = resolveColumn(headers, aliases);
    if (!col) missing.push(key);
    else map[key] = col;
  }
  return { ok: missing.length === 0, map, missing };
}

function parseBranchMasterWorkbook(workbook) {
  for (const name of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: null, raw: true });
    if (!rows.length) continue;
    const res = resolveBranchMasterColumns(Object.keys(rows[0]));
    if (res.ok) return { rows, columns: res.map, sheetName: name };
  }
  return {
    error: "Could not find a sheet in " + BRANCH_MASTER_FILENAME + " containing required columns: Branch, Branch Office Latitude, Branch Office Longitude.",
  };
}

function buildBranchMasterRecords(rows, cols) {
  const errors = [];
  const warnings = [];
  const records = [];
  const seen = new Set();

  rows.forEach((row, idx) => {
    const rowNum = idx + 2;
    const nameRaw = row[cols.BRANCH];
    if (nameRaw === null || nameRaw === undefined || String(nameRaw).trim() === "") {
      errors.push(`Branch Master row ${rowNum}: missing Branch name - row skipped`);
      return;
    }
    const name = String(nameRaw).trim();
    if (seen.has(normalizeBranchKey(name))) {
      warnings.push(`Branch Master: duplicate Branch "${name}" (row ${rowNum}) - first occurrence kept`);
      return;
    }
    seen.add(normalizeBranchKey(name));

    const latRaw = row[cols.OFFICE_LAT];
    const lonRaw = row[cols.OFFICE_LON];
    const lat = latRaw === null || latRaw === undefined || latRaw === "" ? NaN : Number(latRaw);
    const lon = lonRaw === null || lonRaw === undefined || lonRaw === "" ? NaN : Number(lonRaw);
    const hasOfficeCoords = Number.isFinite(lat) && Number.isFinite(lon) && lat >= -90 && lat <= 90 && lon >= -180 && lon <= 180;
    if (!hasOfficeCoords) {
      warnings.push(`Branch "${name}" (row ${rowNum}): missing/invalid office coordinates - no office marker will be shown for it`);
    }

    records.push({
      name,
      address: cols.ADDRESS ? row[cols.ADDRESS] : null,
      officeLat: hasOfficeCoords ? lat : null,
      officeLon: hasOfficeCoords ? lon : null,
      hasOfficeCoords,
    });
  });

  return { records, errors, warnings };
}

function parseBranchPincodeWorkbook(workbook) {
  for (const name of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: null, raw: true });
    if (!rows.length) continue;
    const res = resolveBranchPincodeColumns(Object.keys(rows[0]));
    if (res.ok) return { rows, columns: res.map, sheetName: name };
  }
  return {
    error: "Could not find a sheet in " + BRANCH_PINCODE_FILENAME + " containing required columns: Branch (or Location), Pincode.",
  };
}

function buildBranchPincodeRows(rows, cols) {
  const errors = [];
  const warnings = [];
  const out = [];
  const seenCombo = new Set();

  rows.forEach((row, idx) => {
    const rowNum = idx + 2;
    const pinRaw = row[cols.PINCODE];
    const branchRaw = row[cols.BRANCH];
    const pincode = pinRaw === null || pinRaw === undefined || String(pinRaw).trim() === "" ? null : String(pinRaw).trim();
    const branchName = branchRaw === null || branchRaw === undefined || String(branchRaw).trim() === "" ? null : String(branchRaw).trim();
    if (!pincode || !branchName) {
      errors.push(`Branch-Pincode Mapping row ${rowNum}: missing ${!pincode ? "Pincode" : "Branch"} - row skipped`);
      return;
    }
    const comboKey = normalizeBranchKey(branchName) + "|" + pincode;
    if (seenCombo.has(comboKey)) {
      warnings.push(`Branch-Pincode Mapping: duplicate Branch+Pincode "${branchName}" / ${pincode} (row ${rowNum}) - ignored`);
      return;
    }
    seenCombo.add(comboKey);
    out.push({ pincode, branchName });
  });

  return { rows: out, errors, warnings };
}

/*
 * Joins Branch Master records to Branch-Pincode Mapping rows (exact
 * normalized name match first, then the CONFIRMED_BRANCH_NAME_PAIRS list -
 * never a fuzzy/algorithmic guess), then attaches each branch's pincode
 * geometries (source of truth: pincode_boundaries.geojson - never
 * generated/approximated here) and unions them into that branch's Current
 * Branch Service Area. Never touches Cluster data, never assigns Clusters.
 */
function buildOrgBranches(masterRecords, pincodeRows, pincodeGeometryByCode) {
  const errors = [];
  const warnings = [];
  const branches = new Map(); // key: normalized Branch Master name

  for (const rec of masterRecords) {
    branches.set(normalizeBranchKey(rec.name), {
      masterName: rec.name,
      address: rec.address,
      officeLat: rec.officeLat,
      officeLon: rec.officeLon,
      hasOfficeCoords: rec.hasOfficeCoords,
      mappingName: MASTER_TO_MAPPING_NAME.get(normalizeBranchKey(rec.name)) || null,
      pincodes: [],
      pincodesWithGeometry: [],
      pincodesMissingGeometry: [],
      servicePoly: null,
      boundaryStatus: "INCOMPLETE",
    });
  }

  const byMappingName = new Map();
  for (const b of branches.values()) {
    byMappingName.set(normalizeBranchKey(b.masterName), b); // exact-name match, in case both files already agree
    if (b.mappingName) byMappingName.set(normalizeBranchKey(b.mappingName), b);
  }

  for (const row of pincodeRows) {
    const key = normalizeBranchKey(row.branchName);
    let branch = byMappingName.get(key);
    if (!branch) {
      const masterName = MAPPING_TO_MASTER_NAME.get(key);
      if (masterName) branch = branches.get(normalizeBranchKey(masterName));
    }
    if (!branch) {
      warnings.push(
        `Branch "${row.branchName}" (Branch-Pincode Mapping, pincode ${row.pincode}) has no matching Branch Master record - flagged, office location unavailable for it`
      );
      continue;
    }
    branch.pincodes.push(row.pincode);
  }

  for (const b of branches.values()) {
    const polys = [];
    for (const pin of b.pincodes) {
      const feat = pincodeGeometryByCode.get(pin);
      if (feat) {
        b.pincodesWithGeometry.push(pin);
        polys.push(feat);
      } else {
        b.pincodesMissingGeometry.push(pin);
      }
    }
    if (b.pincodesMissingGeometry.length > 0) {
      warnings.push(`Branch "${b.masterName}": missing pincode geometry for ${b.pincodesMissingGeometry.join(", ")}`);
    }
    if (polys.length > 0) {
      let union = null;
      for (const p of polys) {
        try {
          union = union ? turf.union(union, p) : p;
        } catch (e) {
          // Skip a polygon that fails to union rather than losing the whole service area.
        }
      }
      b.servicePoly = union;
    }
    b.boundaryStatus = b.pincodes.length > 0 && b.pincodesMissingGeometry.length === 0 ? "COMPLETE" : "INCOMPLETE";
  }

  return { branches, errors, warnings };
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
  // Branch Master / Branch-Pincode Mapping / pincode boundary data - entirely
  // separate from Cluster data and from the RM-drawn Branch polygons above.
  orgBranches: new Map(),
  pincodeGeometryByCode: new Map(),
  rawBranchMasterRecords: [],
  rawBranchPincodeRows: [],
  // Raw bytes/text of the 3 org-data inputs, kept only so "Export Standalone
  // Copy" can re-embed exactly what was loaded (network fetch or a prior
  // standalone copy's embedded data) - never re-parsed from here directly.
  rawBranchMasterBytes: null,
  rawBranchPincodeBytes: null,
  rawPincodeGeojsonText: null,
  lastErrors: [],
  lastWarnings: [],
  lastSellerCount: 0,
  branchValidationSummary: null,
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
 * Validation errors/warnings (from the seller workbook, Branch Master,
 * Branch-Pincode Mapping, and pincode-geometry matching) are never silently
 * dropped, but they no longer gate a separate pre-map screen - they
 * accumulate here and are surfaced via the small "data notes" badge in the
 * map screen's top bar (see updateWarningsBadge/dataWarningsBtn below) so
 * the tool can go straight to the map.
 */
function showValidation(errors, warnings, sellerCount) {
  STATE.lastErrors = STATE.lastErrors.concat(errors);
  STATE.lastWarnings = STATE.lastWarnings.concat(warnings);
  if (sellerCount != null) STATE.lastSellerCount = sellerCount;
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
  const bvs = STATE.branchValidationSummary;
  const bvsHtml = bvs
    ? `
    <div class="validation-summary">
      <div><b>Total GLIDs:</b> ${bvs.totalGlids}</div>
      <div><b>GLIDs with Branch:</b> ${bvs.withBranch}</div>
      <div><b>GLIDs without Branch (Other / No data):</b> ${bvs.withoutBranch}</div>
      <div><b>Total current Branches (from input):</b> ${bvs.totalBranches}</div>
      <div><b>Branches without configured colour:</b> ${bvs.branchesWithoutColor.length}</div>
    </div>`
    : "";
  openModal(`
    <h3>Data notes</h3>
    <p>${STATE.lastSellerCount || 0} sellers read &middot; ${errors.length} error(s) &middot; ${warnings.length} warning(s)</p>
    ${bvsHtml}
    <div class="validation-list validation-errors">${errHtml}</div>
    <div class="validation-list validation-warnings">${warnHtml}</div>
    <div class="modal-actions"><button id="warnModalClose" class="primary">Close</button></div>
  `);
  document.getElementById("warnModalClose").onclick = closeModal;
});

const DEFAULT_INPUT_FILENAME = "input.xlsx";
const BRANCH_MASTER_FILENAME = "branch_master.xlsx";
const BRANCH_PINCODE_FILENAME = "branch_pincodes.xlsx";
const PINCODE_BOUNDARY_FILENAME = "pincode_boundaries.geojson";

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

    STATE.branchValidationSummary = computeBranchValidationSummary(STATE.sellers);
    const extraWarnings = STATE.branchValidationSummary.branchesWithoutColor.length
      ? [`Branch(es) without a configured colour - shown in grey: ${STATE.branchValidationSummary.branchesWithoutColor.join(", ")}`]
      : [];
    showValidation(built.errors, built.warnings.concat(extraWarnings), STATE.sellers.length);
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
  STATE.lastErrors = [];
  STATE.lastWarnings = [];
  STATE.lastSellerCount = 0;
  try {
    const [ok] = await Promise.all([
      (async () => {
        const resp = await fetch(DEFAULT_INPUT_FILENAME, { cache: "no-store" });
        if (!resp.ok) {
          showLoadStatus(`Could not find "${DEFAULT_INPUT_FILENAME}" in this folder (HTTP ${resp.status}).`, "error");
          return false;
        }
        const buf = await resp.arrayBuffer();
        return processWorkbookData(buf, DEFAULT_INPUT_FILENAME);
      })(),
      fetchOrgBranchData(),
    ]);
    if (ok) proceedToMap();
  } catch (err) {
    showLoadStatus(
      `Could not load "${DEFAULT_INPUT_FILENAME}" automatically. Make sure it exists in this folder and that ` +
        "you started the tool via start.bat (not by double-clicking index.html directly).",
      "error"
    );
  }
}

/*
 * Shared tail of fetchOrgBranchData() and loadEmbeddedOrgBranchData(): both
 * end up with the same three record lists (Branch Master records,
 * Branch-Pincode Mapping rows, pincode->geometry map) regardless of
 * whether they came from a network fetch or a standalone copy's embedded
 * data - this builds STATE.orgBranches from them and surfaces every issue
 * via the non-blocking "data notes" badge.
 */
function finalizeOrgBranches(masterRecords, pincodeRows, geometryByCode, errors, warnings) {
  const orgResult = buildOrgBranches(masterRecords, pincodeRows, geometryByCode);
  STATE.orgBranches = orgResult.branches;
  STATE.pincodeGeometryByCode = geometryByCode;
  STATE.rawBranchMasterRecords = masterRecords;
  STATE.rawBranchPincodeRows = pincodeRows;
  errors.push(...orgResult.errors);
  warnings.push(...orgResult.warnings);

  const uncoloredOrgBranches = [...orgResult.branches.values()]
    .filter((b) => b.mappingName && !BRANCH_COLOR_MAP.has(normalizeBranchKey(b.mappingName)))
    .map((b) => b.mappingName);
  if (uncoloredOrgBranches.length) {
    warnings.push(`Current Branch Service Area(s) without a configured colour - shown in grey: ${uncoloredOrgBranches.join(", ")}`);
  }

  showValidation(errors, warnings, null);
}

/*
 * Branch Master, Branch-Pincode Mapping and the pincode boundary dataset are
 * all OPTIONAL, best-effort inputs layered on top of the core seller/Cluster
 * workflow above: a missing/unreadable file here never blocks the map -
 * it's recorded as a validation note and the corresponding layer (Branch
 * Offices / Current Branch Service Areas) is simply left empty. The raw
 * bytes/text of each successfully-fetched file are kept on STATE purely so
 * "Export Standalone Copy" can re-embed exactly what this session loaded.
 */
async function fetchOrgBranchData() {
  const errors = [];
  const warnings = [];
  let masterRecords = [];
  let pincodeRows = [];
  const geometryByCode = new Map();

  try {
    const resp = await fetch(BRANCH_MASTER_FILENAME, { cache: "no-store" });
    if (resp.ok) {
      const bytes = new Uint8Array(await resp.arrayBuffer());
      STATE.rawBranchMasterBytes = bytes;
      const wb = XLSX.read(bytes, { type: "array", cellDates: true });
      const parsed = parseBranchMasterWorkbook(wb);
      if (parsed.error) {
        warnings.push(parsed.error);
      } else {
        const built = buildBranchMasterRecords(parsed.rows, parsed.columns);
        masterRecords = built.records;
        errors.push(...built.errors);
        warnings.push(...built.warnings);
      }
    } else {
      warnings.push(`${BRANCH_MASTER_FILENAME} not found in this folder - Branch Offices layer will be empty.`);
    }
  } catch (e) {
    warnings.push(`${BRANCH_MASTER_FILENAME} could not be loaded - Branch Offices layer will be empty.`);
  }

  try {
    const resp = await fetch(BRANCH_PINCODE_FILENAME, { cache: "no-store" });
    if (resp.ok) {
      const bytes = new Uint8Array(await resp.arrayBuffer());
      STATE.rawBranchPincodeBytes = bytes;
      const wb = XLSX.read(bytes, { type: "array", cellDates: true });
      const parsed = parseBranchPincodeWorkbook(wb);
      if (parsed.error) {
        warnings.push(parsed.error);
      } else {
        const built = buildBranchPincodeRows(parsed.rows, parsed.columns);
        pincodeRows = built.rows;
        errors.push(...built.errors);
        warnings.push(...built.warnings);
      }
    } else {
      warnings.push(`${BRANCH_PINCODE_FILENAME} not found in this folder - Current Branch Service Areas layer will be empty.`);
    }
  } catch (e) {
    warnings.push(`${BRANCH_PINCODE_FILENAME} could not be loaded - Current Branch Service Areas layer will be empty.`);
  }

  try {
    const resp = await fetch(PINCODE_BOUNDARY_FILENAME, { cache: "no-store" });
    if (resp.ok) {
      const text = await resp.text();
      STATE.rawPincodeGeojsonText = text;
      const geo = JSON.parse(text);
      for (const feat of geo.features || []) {
        const pin = feat && feat.properties ? String(feat.properties.Pincode || feat.properties.pincode || "").trim() : "";
        if (pin) geometryByCode.set(pin, feat);
      }
    } else {
      warnings.push(`${PINCODE_BOUNDARY_FILENAME} not found in this folder - Current Branch Service Areas cannot be drawn.`);
    }
  } catch (e) {
    warnings.push(`${PINCODE_BOUNDARY_FILENAME} could not be loaded - Current Branch Service Areas cannot be drawn.`);
  }

  finalizeOrgBranches(masterRecords, pincodeRows, geometryByCode, errors, warnings);
}

/*
 * Standalone-copy counterpart to fetchOrgBranchData(): decodes Branch
 * Master / Branch-Pincode Mapping / pincode boundary data from a packaged
 * data.js's embedded base64 instead of fetching sibling files. Any of the
 * three not embedded (e.g. an older standalone copy made before this
 * feature) is simply treated as "not available", same as a missing file
 * over fetch - never blocks the map.
 */
function loadEmbeddedOrgBranchData(masterB64, pincodeB64, geojsonB64) {
  const errors = [];
  const warnings = [];
  let masterRecords = [];
  let pincodeRows = [];
  const geometryByCode = new Map();

  if (masterB64) {
    try {
      const bytes = uint8ArrayFromBase64(masterB64);
      STATE.rawBranchMasterBytes = bytes;
      const wb = XLSX.read(bytes, { type: "array", cellDates: true });
      const parsed = parseBranchMasterWorkbook(wb);
      if (parsed.error) {
        warnings.push(parsed.error);
      } else {
        const built = buildBranchMasterRecords(parsed.rows, parsed.columns);
        masterRecords = built.records;
        errors.push(...built.errors);
        warnings.push(...built.warnings);
      }
    } catch (e) {
      warnings.push(`Embedded ${BRANCH_MASTER_FILENAME} data could not be read - Branch Offices layer will be empty.`);
    }
  } else {
    warnings.push(`${BRANCH_MASTER_FILENAME} was not included in this standalone copy - Branch Offices layer will be empty.`);
  }

  if (pincodeB64) {
    try {
      const bytes = uint8ArrayFromBase64(pincodeB64);
      STATE.rawBranchPincodeBytes = bytes;
      const wb = XLSX.read(bytes, { type: "array", cellDates: true });
      const parsed = parseBranchPincodeWorkbook(wb);
      if (parsed.error) {
        warnings.push(parsed.error);
      } else {
        const built = buildBranchPincodeRows(parsed.rows, parsed.columns);
        pincodeRows = built.rows;
        errors.push(...built.errors);
        warnings.push(...built.warnings);
      }
    } catch (e) {
      warnings.push(`Embedded ${BRANCH_PINCODE_FILENAME} data could not be read - Current Branch Service Areas layer will be empty.`);
    }
  } else {
    warnings.push(`${BRANCH_PINCODE_FILENAME} was not included in this standalone copy - Current Branch Service Areas layer will be empty.`);
  }

  if (geojsonB64) {
    try {
      const text = new TextDecoder("utf-8").decode(uint8ArrayFromBase64(geojsonB64));
      STATE.rawPincodeGeojsonText = text;
      const geo = JSON.parse(text);
      for (const feat of geo.features || []) {
        const pin = feat && feat.properties ? String(feat.properties.Pincode || feat.properties.pincode || "").trim() : "";
        if (pin) geometryByCode.set(pin, feat);
      }
    } catch (e) {
      warnings.push(`Embedded ${PINCODE_BOUNDARY_FILENAME} data could not be read - Current Branch Service Areas cannot be drawn.`);
    }
  } else {
    warnings.push(`${PINCODE_BOUNDARY_FILENAME} was not included in this standalone copy - Current Branch Service Areas cannot be drawn.`);
  }

  finalizeOrgBranches(masterRecords, pincodeRows, geometryByCode, errors, warnings);
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

function textToBase64(text) {
  return base64FromUint8Array(new TextEncoder().encode(text));
}

/*
 * A "standalone copy" is this same tool with every currently-loaded input -
 * the seller/Cluster workbook, and (if present) Branch Master, the
 * Branch-Pincode Mapping, and the pincode boundary geojson - baked directly
 * into a small data.js file (as base64), instead of relying on separate
 * sibling files or a file picker. Since data.js is loaded via a normal
 * <script src>, it works when the recipient just double-clicks index.html -
 * no Python, no local server, no Excel/geojson files to find, nothing to
 * install. Whichever of the 3 optional inputs weren't loaded in this
 * session are simply left out (same as if the file were missing on a
 * normal run). See the README's "Sharing this tool" section.
 */
document.getElementById("exportStandaloneBtn").addEventListener("click", () => {
  if (!STATE.rawWorkbookBytes) {
    alertModal("No data loaded yet.");
    return;
  }
  const lines = [
    '// Generated by "Export Standalone Copy" - embeds every input this tool',
    "// had loaded, so it can run standalone with no Excel/geojson files,",
    "// Python, or server needed.",
    `window.EMBEDDED_WORKBOOK_BASE64 = "${base64FromUint8Array(STATE.rawWorkbookBytes)}";`,
  ];
  if (STATE.rawBranchMasterBytes) {
    lines.push(`window.EMBEDDED_BRANCH_MASTER_BASE64 = "${base64FromUint8Array(STATE.rawBranchMasterBytes)}";`);
  }
  if (STATE.rawBranchPincodeBytes) {
    lines.push(`window.EMBEDDED_BRANCH_PINCODE_BASE64 = "${base64FromUint8Array(STATE.rawBranchPincodeBytes)}";`);
  }
  if (STATE.rawPincodeGeojsonText) {
    lines.push(`window.EMBEDDED_PINCODE_BOUNDARY_BASE64 = "${textToBase64(STATE.rawPincodeGeojsonText)}";`);
  }
  const content = lines.join("\n") + "\n";
  const blob = new Blob([content], { type: "text/javascript" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "data.js";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  const includedExtras = [
    STATE.rawBranchMasterBytes ? "Branch Master" : null,
    STATE.rawBranchPincodeBytes ? "Branch-Pincode Mapping" : null,
    STATE.rawPincodeGeojsonText ? "pincode boundaries" : null,
  ].filter(Boolean);
  const extrasNote = includedExtras.length
    ? ` Also embedded: ${includedExtras.join(", ")}.`
    : " (Branch Master / Branch-Pincode Mapping / pincode boundaries were not loaded in this session, so Branch Offices and Current Branch Service Areas will be empty in the shared copy.)";
  alertModal(
    'data.js downloaded. Put it in a copy of this folder (next to index.html, app.js, styles.css), ' +
      "zip that folder, and send it - see the README's \"Sharing this tool\" section for the full steps." +
      extrasNote
  );
});

/*
 * If this page was packaged as a standalone copy (data.js present, defining
 * window.EMBEDDED_WORKBOOK_BASE64), skip the load screen entirely - decode
 * and go straight to the map. Otherwise fall back to the normal input.xlsx
 * auto-load / manual picker flow.
 */
async function loadEmbeddedWorkbook(base64) {
  STATE.lastErrors = [];
  STATE.lastWarnings = [];
  STATE.lastSellerCount = 0;
  try {
    const bytes = uint8ArrayFromBase64(base64);
    const ok = processWorkbookData(bytes, "input.xlsx");
    // Branch Master / Branch-Pincode Mapping / pincode boundaries: use this
    // standalone copy's own embedded data.js globals if "Export Standalone
    // Copy" included them; otherwise best-effort fall back to fetching
    // sibling files (harmless no-ops if this copy has none of those files).
    if (window.EMBEDDED_BRANCH_MASTER_BASE64 || window.EMBEDDED_BRANCH_PINCODE_BASE64 || window.EMBEDDED_PINCODE_BOUNDARY_BASE64) {
      loadEmbeddedOrgBranchData(
        window.EMBEDDED_BRANCH_MASTER_BASE64 || null,
        window.EMBEDDED_BRANCH_PINCODE_BASE64 || null,
        window.EMBEDDED_PINCODE_BOUNDARY_BASE64 || null
      );
    } else {
      await fetchOrgBranchData();
    }
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
  renderOrgBranchLayers();
  renderOrgBranchList();
  fitMapToData();
  recomputeAssignments();
}

// ------------------------------------------------------------------
// Map
// ------------------------------------------------------------------

function initMap() {
  STATE.map = L.map("map", { preferCanvas: true });

  // So plain L.marker() (used for Seller Locations below) resolves Leaflet's
  // own default pin icon correctly regardless of how the page was loaded -
  // the closest available "branch/location symbol" without a custom asset.
  L.Icon.Default.imagePath = "https://unpkg.com/leaflet@1.9.4/dist/images/";

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

  // Cluster Boundaries: available/toggleable, but OFF at initial load.
  STATE.boundaryLayerGroup = L.layerGroup();
  STATE.centroidLayerGroup = L.layerGroup();
  // Seller Locations: ONE layer/style for every seller (see renderSellerMarkers) -
  // eligibility is no longer conveyed by marker colour, only by popup text.
  STATE.sellerLayer = L.markerClusterGroup({ disableClusteringAtZoom: 16, chunkedLoading: true, maxClusterRadius: 60 }).addTo(STATE.map);
  STATE.metroLayerGroup = L.layerGroup();
  STATE.metroFetchState = "idle";
  STATE.branchDrawnItems = L.featureGroup().addTo(STATE.map);
  STATE.clusterPolygons = new Map();
  STATE.clusterCentroidMarkers = new Map();
  // Branch Offices and Current Branch Service Areas (= "current
  // territories", now colour-coded per BRANCH_COLOR_CONFIG) are both ON by
  // default.
  STATE.officeLayer = L.layerGroup().addTo(STATE.map);
  STATE.serviceAreaLayer = L.layerGroup().addTo(STATE.map);
  STATE.orgBranchOfficeMarkers = new Map();
  STATE.orgBranchServiceAreaLayers = new Map();

  document.getElementById("layerBoundaries").addEventListener("change", (e) => toggleLayer(STATE.boundaryLayerGroup, e.target.checked));
  document.getElementById("layerSellers").addEventListener("change", (e) => toggleLayer(STATE.sellerLayer, e.target.checked));
  document.getElementById("layerBranchOffices").addEventListener("change", (e) => toggleLayer(STATE.officeLayer, e.target.checked));
  document.getElementById("layerServiceArea").addEventListener("change", (e) => toggleLayer(STATE.serviceAreaLayer, e.target.checked));
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
    if (layer._orgBranchKey) {
      e.popup.setContent(orgBranchPopupHtml(layer._orgBranchKey));
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
  const entry = getBranchColorEntry(s.branch);
  return [
    `<b>GLID:</b> ${escapeHtml(s.glid)}`,
    `<b>Branch:</b> <span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:${entry.color};margin-right:4px;"></span>${escapeHtml(s.branch || OTHER_BRANCH_KEY)}`,
    `<b>Cluster ID:</b> ${escapeHtml(s.clusterId || "N/A")}`,
    `<b>Eligibility:</b> ${escapeHtml(s.eligibilityRaw || (s.eligible ? "Eligible" : "Not Eligible"))}`,
    `<b>Outstanding Amount:</b> ${s.outstanding != null ? s.outstanding : "N/A"}`,
    `<b>Renewal Date:</b> ${formatDateVal(s.renewal)}`,
    `<b>Latitude:</b> ${s.lat.toFixed(5)}`,
    `<b>Longitude:</b> ${s.lon.toFixed(5)}`,
  ].join("<br/>");
}

/*
 * Google My Maps-style branch markers, built as inline SVG (no external
 * icon assets/URLs are reachable from this offline tool, so per
 * instruction this is the documented fallback: a consistent pin/circle/
 * diamond/square shape in the seller's Branch colour - colour, not exact
 * icon artwork, is what carries the meaning). Eligibility/Cluster
 * membership are unaffected; they remain popup-only text.
 */
function svgForBranchShape(shape, color) {
  switch (shape) {
    case "circle":
      return `<svg width="20" height="20" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9" fill="${color}" stroke="#fff" stroke-width="2"/></svg>`;
    case "diamond":
      return `<svg width="20" height="20" viewBox="0 0 24 24"><rect x="5" y="5" width="14" height="14" fill="${color}" stroke="#fff" stroke-width="1.5" transform="rotate(45 12 12)"/></svg>`;
    case "square":
      return `<svg width="20" height="20" viewBox="0 0 24 24"><rect x="3" y="3" width="18" height="18" rx="3" fill="${color}" stroke="#fff" stroke-width="1.5"/></svg>`;
    default: // "pin"
      return `<svg width="20" height="26" viewBox="0 0 24 32"><path d="M12 0C6.5 0 2 4.5 2 10c0 7.5 10 21 10 21s10-13.5 10-21c0-5.5-4.5-10-10-10z" fill="${color}" stroke="#fff" stroke-width="1.5"/><circle cx="12" cy="10" r="4" fill="#fff"/></svg>`;
  }
}

const sellerIconCache = new Map();
function sellerIcon(branchName) {
  const entry = getBranchColorEntry(branchName);
  const cacheKey = entry.shape + "|" + entry.color;
  if (sellerIconCache.has(cacheKey)) return sellerIconCache.get(cacheKey);
  const isPin = !entry.shape || entry.shape === "pin";
  const icon = L.divIcon({
    html: svgForBranchShape(entry.shape, entry.color),
    className: "seller-marker-icon",
    iconSize: isPin ? [20, 26] : [20, 20],
    iconAnchor: isPin ? [10, 26] : [10, 10],
    popupAnchor: [0, isPin ? -24 : -10],
  });
  sellerIconCache.set(cacheKey, icon);
  return icon;
}

/*
 * Seller Locations: one Branch-coloured/shaped marker per seller, per
 * BRANCH_COLOR_CONFIG. Eligible/ineligible is purely informational text in
 * the popup, never a marker colour/shape - the underlying eligibility data
 * itself (and Cluster membership) is untouched; only the visual glyph
 * changed, from a plain Leaflet pin to a Branch-specific one.
 */
function renderSellerMarkers() {
  STATE.sellerLayer.clearLayers();
  for (const s of STATE.sellers) {
    if (!s.validCoords) continue;
    const marker = L.marker([s.lat, s.lon], { icon: sellerIcon(s.branch) });
    marker.bindPopup(sellerPopupHtml(s), { maxWidth: 270 });
    STATE.sellerLayer.addLayer(marker);
  }
}

function starIcon() {
  return L.divIcon({ html: "&#9733;", className: "star-icon", iconSize: [16, 16], iconAnchor: [8, 8] });
}

function metroIcon() {
  return L.divIcon({ html: "M", className: "metro-icon", iconSize: [16, 16], iconAnchor: [8, 8] });
}

/*
 * Branch Office marker: deliberately larger/shape-distinct from every
 * seller/GLID marker (a building glyph in a rounded square, vs. seller
 * pins/circles/diamonds), so the two marker types are never confused - but
 * tinted with the same Branch colour used everywhere else for that Branch,
 * per the "use the same Branch colour consistently" instruction. Falls
 * back to the same grey as "Other / No data" if this Branch's mapping
 * name has no configured colour (flagged in validation, never dropped).
 */
function officeIcon(color) {
  const bg = color || OTHER_BRANCH_ENTRY.color;
  return L.divIcon({
    html: `<div style="background:${bg};width:22px;height:22px;border-radius:4px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:13px;border:2px solid #fff;box-shadow:0 0 3px rgba(0,0,0,0.6);">&#127970;</div>`,
    className: "office-icon-wrap",
    iconSize: [22, 22],
    iconAnchor: [11, 22],
  });
}

function orgBranchStatusBadgeHtml(status) {
  return status === "COMPLETE"
    ? '<span class="status-badge complete">COMPLETE</span>'
    : '<span class="status-badge incomplete">INCOMPLETE &mdash; Missing Pincode Geometry</span>';
}

function orgBranchPopupHtml(key) {
  const b = STATE.orgBranches.get(key);
  if (!b) return "";
  const entry = getBranchColorEntry(b.mappingName);
  return [
    `<b>Branch:</b> <span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:${entry.color};margin-right:4px;"></span>${escapeHtml(b.masterName)}`,
    b.address ? `<b>Address:</b> ${escapeHtml(String(b.address))}` : null,
    b.hasOfficeCoords ? `<b>Office Latitude:</b> ${b.officeLat}` : `<b>Office Latitude:</b> N/A`,
    b.hasOfficeCoords ? `<b>Office Longitude:</b> ${b.officeLon}` : `<b>Office Longitude:</b> N/A`,
    `<b>Pincodes Served:</b> ${b.pincodes.length}`,
    `<b>Boundary Status:</b> ${orgBranchStatusBadgeHtml(b.boundaryStatus)}`,
  ]
    .filter(Boolean)
    .join("<br/>");
}

/*
 * Renders the Branch Offices (marker) and Current Branch Service Areas
 * (polygon) layers from STATE.orgBranches. Purely a reference/visualization
 * layer pair - never touches Cluster data, never participates in the
 * Branch-drawing/cluster-assignment logic below.
 */
function renderOrgBranchLayers() {
  STATE.officeLayer.clearLayers();
  STATE.serviceAreaLayer.clearLayers();
  STATE.orgBranchOfficeMarkers.clear();
  STATE.orgBranchServiceAreaLayers.clear();

  for (const [key, b] of STATE.orgBranches) {
    const colorEntry = getBranchColorEntry(b.mappingName);

    if (b.hasOfficeCoords) {
      const marker = L.marker([b.officeLat, b.officeLon], { icon: officeIcon(colorEntry.color) });
      marker._orgBranchKey = key;
      marker.bindPopup("", { maxWidth: 260 });
      marker.bindTooltip(b.masterName, { direction: "top" });
      marker.addTo(STATE.officeLayer);
      STATE.orgBranchOfficeMarkers.set(key, marker);
    }

    if (b.servicePoly) {
      // Current Branch Service Area ("current territory"): solid, distinct,
      // persistent colour per Branch (see BRANCH_COLOR_CONFIG) - same
      // colour as that Branch's legend entry, seller markers and popups.
      // Reasonable fill transparency so Clusters/sellers/roads underneath
      // stay visible. Falls back to grey (flagged in validation) if this
      // Branch's mapping name has no configured colour - never hidden.
      const layer = L.geoJSON(b.servicePoly, {
        style: { color: colorEntry.color, weight: 2, fillColor: colorEntry.color, fillOpacity: 0.25 },
        onEachFeature: (feature, lyr) => {
          lyr._orgBranchKey = key;
        },
      });
      layer.bindPopup("", { maxWidth: 260 });
      layer.addTo(STATE.serviceAreaLayer);
      STATE.orgBranchServiceAreaLayers.set(key, layer);
    }
  }

  renderBranchLegend();
}

function renderOrgBranchList() {
  const el = document.getElementById("orgBranchList");
  if (!el) return;
  el.innerHTML = "";
  if (!STATE.orgBranches || STATE.orgBranches.size === 0) {
    el.innerHTML = '<div class="list-row">No Branch Master data loaded.</div>';
    return;
  }
  for (const [key, b] of STATE.orgBranches) {
    const entry = getBranchColorEntry(b.mappingName);
    const row = document.createElement("div");
    row.className = "list-row";
    row.innerHTML = `<span><span class="swatch" style="background:${entry.color}"></span>${escapeHtml(b.masterName)}</span><span class="count">${b.pincodes.length} pincodes ${orgBranchStatusBadgeHtml(b.boundaryStatus)}</span>`;
    row.onclick = () => showOrgBranchDetail(key);
    el.appendChild(row);
  }
}

/*
 * Dynamic map legend for the colour-coded "current Branches / territories"
 * - built entirely from BRANCH_COLOR_CONFIG plus whichever Branch names are
 * actually present in the loaded data (seller Branch field and/or Branch
 * Master / Branch-Pincode Mapping), so there is exactly one place that
 * defines these colours; nothing is hardcoded a second time here.
 */
function renderBranchLegend() {
  const el = document.getElementById("branchLegendList");
  if (!el) return;
  el.innerHTML = "";

  const present = new Set();
  for (const s of STATE.sellers) if (s.branch) present.add(s.branch);
  for (const b of STATE.orgBranches.values()) if (b.mappingName) present.add(b.mappingName);
  const hasOther = STATE.sellers.some((s) => !s.branch);

  const rows = [];
  for (const name of Object.keys(BRANCH_COLOR_CONFIG)) {
    if (present.has(name)) rows.push(getBranchColorEntry(name));
  }
  for (const name of present) {
    if (!BRANCH_COLOR_MAP.has(normalizeBranchKey(name))) rows.push(getBranchColorEntry(name));
  }
  if (hasOther) rows.push(getBranchColorEntry(null));

  if (!rows.length) {
    el.innerHTML = '<div class="legend-row">No current Branch data loaded.</div>';
    return;
  }
  for (const entry of rows) {
    const row = document.createElement("div");
    row.className = "legend-row";
    row.innerHTML = `<span class="legend-swatch" style="background:${entry.color}"></span><span>${escapeHtml(entry.name)}</span>`;
    el.appendChild(row);
  }
}

function showOrgBranchDetail(key) {
  const b = STATE.orgBranches.get(key);
  if (!b) return;

  const officeMarker = STATE.orgBranchOfficeMarkers.get(key);
  const areaLayer = STATE.orgBranchServiceAreaLayers.get(key);
  if (officeMarker) {
    document.getElementById("layerBranchOffices").checked = true;
    toggleLayer(STATE.officeLayer, true);
    STATE.map.setView(officeMarker.getLatLng(), 13);
  } else if (areaLayer) {
    document.getElementById("layerServiceArea").checked = true;
    toggleLayer(STATE.serviceAreaLayer, true);
    STATE.map.fitBounds(areaLayer.getBounds());
  }

  const pincodeList = b.pincodes.length ? escapeHtml(b.pincodes.join(", ")) : "(none assigned)";
  const missingBlock = b.pincodesMissingGeometry.length
    ? `<p><b>Missing geometry for:</b> ${escapeHtml(b.pincodesMissingGeometry.join(", "))}</p>`
    : "";
  openModal(`
    <h3>${escapeHtml(b.masterName)}</h3>
    <p>
      <b>Office Address:</b> ${b.address ? escapeHtml(String(b.address)) : "N/A"}<br/>
      <b>Office Latitude:</b> ${b.hasOfficeCoords ? b.officeLat : "N/A"}<br/>
      <b>Office Longitude:</b> ${b.hasOfficeCoords ? b.officeLon : "N/A"}<br/>
      <b>Number of Pincodes Served:</b> ${b.pincodes.length}<br/>
      <b>Boundary Data Status:</b> ${orgBranchStatusBadgeHtml(b.boundaryStatus)}
    </p>
    <p><b>Pincodes Served:</b> ${pincodeList}</p>
    ${missingBlock}
    <div class="modal-actions"><button id="orgBranchModalClose" class="primary">Close</button></div>
  `);
  document.getElementById("orgBranchModalClose").onclick = closeModal;
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
    "Current Branch (from Seller Data)": s.branch || OTHER_BRANCH_KEY,
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

  // Reference sheets for the Branch Master / Branch-Pincode Mapping inputs -
  // passthrough of the source data plus a computed boundary-status summary.
  // None of this feeds back into the seller/Cluster mapping sheets above.
  const sheet6 = STATE.rawBranchMasterRecords.map((r) => ({
    Branch: r.name,
    "Office Address": r.address,
    "Office Latitude": r.officeLat,
    "Office Longitude": r.officeLon,
    "Has Office Coordinates": r.hasOfficeCoords ? "Yes" : "No",
  }));

  const sheet7 = STATE.rawBranchPincodeRows.map((r) => ({
    Branch: r.branchName,
    Pincode: r.pincode,
  }));

  const sheet8 = [...STATE.orgBranches.values()].map((b) => ({
    Branch: b.masterName,
    "Total Pincodes": b.pincodes.length,
    "Pincodes With Geometry": b.pincodesWithGeometry.length,
    "Pincodes Missing Geometry": b.pincodesMissingGeometry.length,
    "Boundary Status": b.boundaryStatus,
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet1), "Seller Mapping");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet2), "Cluster Mapping");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet3), "Branch Summary");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet4), "Unmapped Clusters");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet5), "Run Configuration");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet6), "Branch Master");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet7), "Branch-Pincode Mapping");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sheet8), "Branch Service Area Status");

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
