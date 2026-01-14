/**
 * ==========================================
 * TROOP COOKIE APP – BACKEND API (Code.gs)
 * Container-bound to the Cookie Season 2026 workbook
 * ==========================================
 *
 * Frontend calls: google.script.run.api({ action, payload })
 * Response shape: { ok: true|false, data: any, error: string|null }
 *
 * Routing:
 *  - Admin: /exec (default)
 *  - Parent: /exec?page=parent&t=<Parent_Token>
 *
 * Notes:
 *  - NO hard-coded column numbers. All reads/writes are header-based.
 *  - Idempotent posting (requests, booth sales, booth credits) using Posted_* fields.
 *  - Uses LockService for post actions to prevent double clicks / race conditions.
 */

/* ===========================
   SHEET NAMES (verified)
=========================== */
const SHEETS = {
  CONFIG: "Config",

  PARENTS_GIRLS: "Parents_Girls",

  INVENTORY_MOVES: "Inventory_Movements",
  INVENTORY_SUMMARY: "Inventory_Summary",

  REQUESTS: "Requests_Log",

  PAYMENTS: "Payments",
  PAYMENT_APPS: "Payment_Applications_Log",
  PARENT_SUMMARY: "Parent_Summary",

  BOOTHS: "Booths",
  BOOTH_SIGNUPS: "Booth_Signups",
  BOOTH_COUNTS: "Booth_Inventory_Counts",
  BOOTH_CASH: "Booth_Cash_Tracker",
  BOOTH_SALES: "Booth_Sales_Log",
  BOOTH_ALLOC: "Booth_Allocations",

  GIRL_BOOTH_CREDITS: "Girl_Booth_Credits_Log",
  GIRL_TOTAL_CREDITS: "Girl_Total_Credits_Log",
  GIRL_ALL_CREDITS_VIEW: "Girl_All_Credits_View",

  TROOP_DASHBOARD: "Troop_Dashboard",
  GIRL_CREDIT_SUMMARY: "Girl_Credit_Summary"
};

/* ===========================
   DOMAIN CONSTANTS
=========================== */
const MOVE_TYPES = {
  TROOP_TO_PARENT: "TROOP_TO_PARENT",
  BOOTH_SALE: "BOOTH_SALE"
};
const ENTITIES = {
  TROOP: "Troop"
};

/* ===========================
   WEB APP ROUTER
=========================== */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const page = String(params.page || "").toLowerCase();
  const mode = page === "parent" ? "parent" : "admin";

  const t = HtmlService.createTemplateFromFile("Index");
  t.__query = params;
  t.__mode = mode;

  return t.evaluate()
    .setTitle("Troop Cookie Management")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Include helper for HTML templates (if you use it in index.html)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ===========================
   SINGLE ENTRYPOINT API
=========================== */
function api(req) {
  try {
    const action = (req && req.action) ? String(req.action) : "";
    const payload = (req && req.payload) ? req.payload : {};

    // Small helpers (defined inside so we don't depend on other helper functions)
    const ok = (data) => ({ ok: true, data });
    const fail = (error) => ({ ok: false, error: String(error || "Unknown error") });

    if (!action) return fail("Missing action");

    switch (action) {
      // --- Auth ---
      case "auth.whoAmI":
        // your existing function should return an object
        return ok(auth_whoAmI_());

      // --- Admin ---
      case "admin.bootstrap":
        return ok(admin_bootstrap_());

      // --- Parent ---
      case "parent.bootstrap":
        // IMPORTANT: must RETURN
        return ok(parent_bootstrap_(payload));

      // --- Parent creates request ---
      case "requests.parent.create":
        return ok(requests_parent_create_(payload));

      default:
        return fail("Unknown action: " + action);
    }
  } catch (err) {
    // Always return a response object so the client never sees "no response"
    return { ok: false, error: (err && err.stack) ? String(err.stack) : String(err) };
  }
}

function apiV2(req) {
  try {
    const action = (req && req.action) ? String(req.action) : "";
    const payload = (req && req.payload) ? req.payload : {};

    const ok = (data) => ({ ok: true, data });
    const fail = (error) => ({ ok: false, error: String(error || "Unknown error") });

    if (!action) return fail("Missing action");

    switch (action) {
      case "auth.whoAmI":
        return ok(auth_whoAmI_());

      case "admin.bootstrap":
        return ok(admin_bootstrap_());

      case "parent.bootstrap":
        return ok(parent_bootstrap_(payload));

      case "requests.parent.create":
        return ok(requests_parent_create_(payload));

      default:
        return fail("Unknown action: " + action);
    }
  } catch (err) {
    return { ok: false, error: (err && err.stack) ? String(err.stack) : String(err) };
  }
}

/* ===========================
   AUTH / BOOTSTRAP
=========================== */

function auth_whoAmI_() {
  const email = getUserEmail_();
  const isAdmin = isAdminEmail_(email);
  return { role: isAdmin ? "admin" : "guest", email: email || "" };
}

function admin_bootstrap_() {
  requireAdmin_();

  // Lightweight bootstrap: config + key tables that are formula-driven
  const config = getConfigMap_();
  const boxPrice = getBoxPrice_(config);

  // KPIs (prefer formula sheets; reads are cheap)
  const invSummary = readSheetObjects_(SHEETS.INVENTORY_SUMMARY);
  const parentSummary = readSheetObjects_(SHEETS.PARENT_SUMMARY);

  // Counts from Requests_Log and booth posting status derived from Booth_Inventory_Counts
  const reqRows = readSheetObjects_(SHEETS.REQUESTS);
  const requestsPending = reqRows.filter(r => lc_(r.Status) === "pending").length;
  const requestsApproved = reqRows.filter(r => lc_(r.Status) === "approved").length;

  const boothCounts = readSheetObjects_(SHEETS.BOOTH_COUNTS);
  const boothsNeedingPost = boothCounts.filter(r => !truthy_(r.Posted_Sold) && n_(r.Boxes_Sold) > 0).length;

  return {
    config,
    boxPrice,
    inventorySummary: invSummary,
    parentSummary,
    kpis: {
      requestsPending,
      requestsApproved,
      boothsNeedingPost
    }
  };
}

function parent_bootstrap_(payload) {
  const token = String(payload.token || "").trim();
  const scope = requireParentToken_(token);

  // Parent summary + requests + credits, scoped to Parent_Name and their girls
  const parentSummary = readSheetObjects_(SHEETS.PARENT_SUMMARY)
    .find(r => eq_(r.Parent_Name, scope.parentName)) || null;

  const requests = readSheetObjects_(SHEETS.REQUESTS)
    .filter(r => eq_(r.Parent_Name, scope.parentName));

  // Girl credits view: filter to girls in this family
  const girlsSet = new Set(scope.girls.map(g => String(g.Girl_Name || "").trim()));
  const allCredits = readSheetObjects_(SHEETS.GIRL_ALL_CREDITS_VIEW)
    .filter(r => girlsSet.has(String(r.Girl_Name || "").trim()));

  return {
    parent: { parentName: scope.parentName, parentEmail: scope.parentEmail, token },
    girls: scope.girls,
    summary: parentSummary,
    requests,
    credits: allCredits
  };
}

/* ===========================
   PARENTS & GIRLS
=========================== */

function parents_list_() {
  requireAdmin_();

  const pg = readSheetObjects_(SHEETS.PARENTS_GIRLS);
  const summary = readSheetObjects_(SHEETS.PARENT_SUMMARY);

  // unique parents
  const map = new Map();
  for (const r of pg) {
    const parentName = String(r.Parent_Name || "").trim();
    if (!parentName) continue;
    if (!map.has(parentName)) {
      map.set(parentName, {
        Parent_Name: parentName,
        Parent_Email: String(r.Parent_Email || "").trim(),
        Parent_Token: String(r.Parent_Token || "").trim(),
        Girls_Count: 0
      });
    }
    map.get(parentName).Girls_Count += 1;
  }

  // attach balance
  const sumByParent = new Map(summary.map(r => [String(r.Parent_Name || "").trim(), r]));
  const out = Array.from(map.values()).map(p => {
    const s = sumByParent.get(p.Parent_Name);
    return Object.assign({}, p, {
      Total_Credited_$: s ? s["Total_Credited_$"] : "",
      Total_Paid_$: s ? s["Total_Paid_$"] : "",
      Balance_Due_$: s ? s["Balance_Due_$"] : "",
      Status: s ? s["Status"] : ""
    });
  });

  // sort alpha
  out.sort((a, b) => a.Parent_Name.localeCompare(b.Parent_Name));
  return out;
}

function parents_detail_(payload) {
  requireAdmin_();
  const parentName = String(payload.parentName || "").trim();
  if (!parentName) throw new Error("parentName is required");

  const pg = readSheetObjects_(SHEETS.PARENTS_GIRLS).filter(r => eq_(r.Parent_Name, parentName));
  const summary = readSheetObjects_(SHEETS.PARENT_SUMMARY).find(r => eq_(r.Parent_Name, parentName)) || null;
  const requests = readSheetObjects_(SHEETS.REQUESTS).filter(r => eq_(r.Parent_Name, parentName));
  const payments = readSheetObjects_(SHEETS.PAYMENTS).filter(r => eq_(r.Parent_Name, parentName));

  const token = (pg[0] && pg[0].Parent_Token) ? String(pg[0].Parent_Token).trim() : "";
  return {
    parent: { Parent_Name: parentName, Parent_Email: (pg[0] ? pg[0].Parent_Email : "") || "", Parent_Token: token },
    girls: pg,
    summary,
    requests: requests.slice(-50),
    payments: payments.slice(-50),
    parentLink: token ? buildParentLink_(token) : ""
  };
}

function parents_regenToken_(payload) {
  requireAdmin_();
  const parentName = String(payload.parentName || "").trim();
  if (!parentName) throw new Error("parentName is required");

  const sheet = getSheet_(SHEETS.PARENTS_GIRLS);
  const { headers, map, values } = readSheet_(sheet);

  const colParent = needCol_(map, "parent_name");
  const colToken = needCol_(map, "parent_token");

  const newToken = makeToken_();
  let updated = 0;

  for (let i = 0; i < values.length; i++) {
    const rowIdx = i + 2;
    const p = String(values[i][colParent - 1] || "").trim();
    if (!eq_(p, parentName)) continue;
    sheet.getRange(rowIdx, colToken).setValue(newToken);
    updated++;
  }

  return { parentName, token: newToken, updatedRows: updated, parentLink: buildParentLink_(newToken) };
}

function girls_add_(payload) {
  requireAdmin_();
  const girlName = String(payload.girlName || "").trim();
  const parentName = String(payload.parentName || "").trim();
  const parentEmail = String(payload.parentEmail || "").trim();
  const active = payload.active === false ? false : true;

  if (!girlName || !parentName) throw new Error("girlName and parentName are required");

  // Reuse existing parent token if present, else generate
  const existing = getParentScopeByName_(parentName);
  const token = existing && existing.token ? existing.token : makeToken_();

  const sheet = getSheet_(SHEETS.PARENTS_GIRLS);
  appendByHeader_(sheet, {
    Girl_Name: girlName,
    Parent_Name: parentName,
    Parent_Email: parentEmail,
    Active: active ? "TRUE" : "FALSE",
    Parent_Token: token
  });

  return { girlName, parentName, parentEmail, active, token, parentLink: buildParentLink_(token) };
}

function girls_setActive_(payload) {
  requireAdmin_();
  const girlName = String(payload.girlName || "").trim();
  const parentName = payload.parentName ? String(payload.parentName || "").trim() : "";
  const active = payload.active === true;

  if (!girlName) throw new Error("girlName is required");

  const sheet = getSheet_(SHEETS.PARENTS_GIRLS);
  const { map, values } = readSheet_(sheet);
  const colGirl = needCol_(map, "girl_name");
  const colParent = map["parent_name"] || 0;
  const colActive = needCol_(map, "active");

  let updated = 0;
  for (let i = 0; i < values.length; i++) {
    const rowIdx = i + 2;
    const g = String(values[i][colGirl - 1] || "").trim();
    if (!eq_(g, girlName)) continue;
    if (parentName && colParent) {
      const p = String(values[i][colParent - 1] || "").trim();
      if (!eq_(p, parentName)) continue;
    }
    sheet.getRange(rowIdx, colActive).setValue(active ? "TRUE" : "FALSE");
    updated++;
  }

  return { girlName, parentName, active, updatedRows: updated };
}

/* ===========================
   INVENTORY
=========================== */

function inventory_summary_() {
  requireAdmin_();
  return readSheetObjects_(SHEETS.INVENTORY_SUMMARY);
}

function inventory_movements_list_(payload) {
  requireAdmin_();
  const filters = payload || {};
  const moveType = String(filters.moveType || "").trim();
  const parentName = String(filters.parentName || "").trim();
  const cookieType = String(filters.cookieType || "").trim();
  const fromDate = filters.fromDate ? new Date(filters.fromDate) : null;
  const toDate = filters.toDate ? new Date(filters.toDate) : null;

  const rows = readSheetObjects_(SHEETS.INVENTORY_MOVES);
  return rows.filter(r => {
    if (moveType && !eq_(r.Move_Type, moveType)) return false;
    if (parentName && !eq_(r.Parent_Name, parentName)) return false;
    if (cookieType && !eq_(r.Cookie_Type, cookieType)) return false;

    if (fromDate || toDate) {
      const d = toDate_(r.Move_Date);
      if (!d) return false;
      if (fromDate && d < fromDate) return false;
      if (toDate && d > toDate) return false;
    }
    return true;
  });
}

function inventory_movements_add_(payload) {
  requireAdmin_();
  const p = payload || {};
  const now = new Date();
  const email = getUserEmail_();

  const move = {
    Move_ID: makeId_("MOVE"),
    Move_Date: p.Move_Date || now,
    Move_Type: String(p.Move_Type || "").trim(),
    From_Entity: String(p.From_Entity || "").trim(),
    To_Entity: String(p.To_Entity || "").trim(),
    Parent_Name: String(p.Parent_Name || "").trim(),
    Girl_Name: String(p.Girl_Name || "").trim(),
    Cookie_Type: String(p.Cookie_Type || "").trim(),
    Boxes: n_(p.Boxes),
    Notes: String(p.Notes || "").trim(),
    Entered_By: email || "",
    Entered_At: now
  };

  if (!move.Move_Type || !move.Cookie_Type || !move.Boxes) {
    throw new Error("Move_Type, Cookie_Type, and Boxes are required");
  }
  if (move.Boxes <= 0) throw new Error("Boxes must be a positive number");

  const sheet = getSheet_(SHEETS.INVENTORY_MOVES);
  const rowIndex = appendByHeader_(sheet, move);
  return Object.assign({ _row: rowIndex }, move);
}

/* ===========================
   REQUESTS
=========================== */

function requests_parent_create_(payload) {
  const token = String(payload.token || "").trim();
  const scope = requireParentToken_(token);

  const cookieType = String(payload.Cookie_Type || payload.cookieType || "").trim();
  const boxesRequested = n_(payload.Boxes_Requested || payload.boxesRequested);
  const notes = String(payload.Parent_Notes || payload.notes || "").trim();

  if (!cookieType || !boxesRequested) throw new Error("Cookie_Type and Boxes_Requested are required");
  if (boxesRequested <= 0) throw new Error("Boxes_Requested must be positive");

  const now = new Date();

  const rowObj = {
    Timestamp: now,
    Parent_Name: scope.parentName,
    Cookie_Type: cookieType,
    Boxes_Requested: boxesRequested,
    Parent_Notes: notes,
    Status: "Pending"
    // other columns left to formulas / admin workflow
  };

  const sheet = getSheet_(SHEETS.REQUESTS);
  const rowIndex = appendByHeader_(sheet, rowObj);
  return Object.assign({ _row: rowIndex }, rowObj);
}

function requests_admin_list_(payload) {
  requireAdmin_();
  const f = payload || {};
  const status = String(f.status || "").trim();
  const parentName = String(f.parentName || "").trim();
  const fromDate = f.fromDate ? new Date(f.fromDate) : null;
  const toDate = f.toDate ? new Date(f.toDate) : null;
  const unpostedOnly = !!f.unpostedOnly;

  const rows = readSheetObjects_(SHEETS.REQUESTS);
  return rows.filter(r => {
    if (status && !eq_(r.Status, status)) return false;
    if (parentName && !eq_(r.Parent_Name, parentName)) return false;

    if (fromDate || toDate) {
      const d = toDate_(r.Timestamp);
      if (!d) return false;
      if (fromDate && d < fromDate) return false;
      if (toDate && d > toDate) return false;
    }

    if (unpostedOnly) {
      // "posted" = Posted_At present OR Posted_Delivered_Boxes present
      if (truthy_(r.Posted_At) || n_(r.Posted_Delivered_Boxes) > 0) return false;
    }
    return true;
  });
}

function requests_admin_approve_(payload) {
  requireAdmin_();
  const row = toRowIndex_(payload.requestRowId || payload._row || payload.row);
  const boxesApproved = n_(payload.Boxes_Approved || payload.boxesApproved);
  const notes = String(payload.Approver_Notes || payload.notes || "").trim();
  if (!row) throw new Error("requestRowId (_row) is required");
  if (boxesApproved <= 0) throw new Error("Boxes_Approved must be positive");

  const sheet = getSheet_(SHEETS.REQUESTS);
  const now = new Date();

  updateByHeader_(sheet, row, {
    Status: "Approved",
    Boxes_Approved: boxesApproved,
    Approver_Notes: notes,
    Approved_Date: now
  });

  return { _row: row, Status: "Approved", Boxes_Approved: boxesApproved, Approver_Notes: notes, Approved_Date: now };
}

function requests_admin_deliverAndPost_(payload) {
  requireAdmin_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const row = toRowIndex_(payload.requestRowId || payload._row || payload.row);
    const boxesDelivered = n_(payload.Boxes_Delivered || payload.boxesDelivered);
    const notes = String(payload.Notes || payload.notes || "").trim();
    if (!row) throw new Error("requestRowId (_row) is required");
    if (boxesDelivered <= 0) throw new Error("Boxes_Delivered must be positive");

    const reqSheet = getSheet_(SHEETS.REQUESTS);
    const reqMap = getHeaderMap_(reqSheet);

    const colPostedAt = reqMap["posted_at"];
    const colPostedBoxes = reqMap["posted_delivered_boxes"];

    // Check idempotency
    if (colPostedAt) {
      const already = reqSheet.getRange(row, colPostedAt).getValue();
      if (truthy_(already)) {
        return { _row: row, alreadyPosted: true };
      }
    }
    if (colPostedBoxes) {
      const alreadyBoxes = n_(reqSheet.getRange(row, colPostedBoxes).getValue());
      if (alreadyBoxes > 0) {
        return { _row: row, alreadyPosted: true };
      }
    }

    // Read request row minimal fields
    const colParent = needCol_(reqMap, "parent_name");
    const colCookie = needCol_(reqMap, "cookie_type");

    const parentName = String(reqSheet.getRange(row, colParent).getValue() || "").trim();
    const cookieType = String(reqSheet.getRange(row, colCookie).getValue() || "").trim();
    if (!parentName || !cookieType) throw new Error("Request row missing Parent_Name or Cookie_Type");

    const now = new Date();

    // Update request row
    updateByHeader_(reqSheet, row, {
      Status: "Delivered",
      Boxes_Delivered: boxesDelivered,
      Delivered_Date: now,
      Posted_At: now,
      Posted_Delivered_Boxes: boxesDelivered
    });

    // Append inventory movement
    const move = {
      Move_ID: makeId_("MOVE"),
      Move_Date: now,
      Move_Type: MOVE_TYPES.TROOP_TO_PARENT,
      From_Entity: ENTITIES.TROOP,
      To_Entity: parentName,
      Parent_Name: parentName,
      Girl_Name: "",
      Cookie_Type: cookieType,
      Boxes: boxesDelivered,
      Notes: notes ? `Request delivery: ${notes}` : "Request delivery",
      Entered_By: getUserEmail_() || "",
      Entered_At: now
    };
    const moveSheet = getSheet_(SHEETS.INVENTORY_MOVES);
    const moveRow = appendByHeader_(moveSheet, move);

    return { request: { _row: row, parentName, cookieType, boxesDelivered }, movement: Object.assign({ _row: moveRow }, move) };
  } finally {
    lock.releaseLock();
  }
}

/* ===========================
   PAYMENTS
=========================== */

function payments_list_(payload) {
  requireAdmin_();
  const f = payload || {};
  const parentName = String(f.parentName || "").trim();
  const method = String(f.method || "").trim();
  const fromDate = f.fromDate ? new Date(f.fromDate) : null;
  const toDate = f.toDate ? new Date(f.toDate) : null;

  const rows = readSheetObjects_(SHEETS.PAYMENTS);
  return rows.filter(r => {
    if (parentName && !eq_(r.Parent_Name, parentName)) return false;
    if (method && !eq_(r.Method, method)) return false;
    if (fromDate || toDate) {
      const d = toDate_(r.Timestamp);
      if (!d) return false;
      if (fromDate && d < fromDate) return false;
      if (toDate && d > toDate) return false;
    }
    return true;
  });
}

function payments_add_(payload) {
  requireAdmin_();
  const p = payload || {};
  const parentName = String(p.Parent_Name || p.parentName || "").trim();
  const amount = n_(p.Amount || p.amount);
  const method = String(p.Method || p.method || "").trim();
  const notes = String(p.Notes || p.notes || "").trim();

  if (!parentName || !amount || !method) throw new Error("Parent_Name, Amount, and Method are required");
  if (amount <= 0) throw new Error("Amount must be positive");

  const now = new Date();

  const rowObj = {
    Timestamp: now,
    Parent_Name: parentName,
    Amount: amount,
    Method: method,
    Notes: notes,
    Posted_At: now,
    Posted_Amount: amount
  };

  const sheet = getSheet_(SHEETS.PAYMENTS);
  const row = appendByHeader_(sheet, rowObj);
  return Object.assign({ _row: row }, rowObj);
}

function payments_apply_(payload) {
  requireAdmin_();
  const paymentRow = toRowIndex_(payload.paymentRow || payload.Payment_Row || payload.payment_row);
  const apps = payload.applications || [];
  if (!paymentRow) throw new Error("paymentRow is required");
  if (!Array.isArray(apps) || apps.length === 0) throw new Error("applications[] is required");

  const paySheet = getSheet_(SHEETS.PAYMENTS);
  const payMap = getHeaderMap_(paySheet);
  const colParent = needCol_(payMap, "parent_name");
  const parentName = String(paySheet.getRange(paymentRow, colParent).getValue() || "").trim();
  if (!parentName) throw new Error("Payments row missing Parent_Name");

  const now = new Date();
  const out = [];

  const appSheet = getSheet_(SHEETS.PAYMENT_APPS);
  for (const a of apps) {
    const girlName = String(a.Girl_Name || a.girlName || "").trim();
    const amt = n_(a.Amount_Applied || a.amount);
    if (!girlName || amt <= 0) continue;
    const rowObj = {
      Apply_ID: makeId_("APPLY"),
      Payment_Row: paymentRow,
      Applied_At: now,
      Parent_Name: parentName,
      Girl_Name: girlName,
      Amount_Applied: amt
    };
    const row = appendByHeader_(appSheet, rowObj);
    out.push(Object.assign({ _row: row }, rowObj));
  }

  return out;
}

function parents_balances_() {
  requireAdmin_();
  return readSheetObjects_(SHEETS.PARENT_SUMMARY);
}

/* ===========================
   BOOTHS
=========================== */

function booths_list_(payload) {
  // Admin can see all. Parent token can see upcoming (or all if you prefer).
  const token = payload && payload.token ? String(payload.token || "").trim() : "";
  if (token) requireParentToken_(token); // validates only

  const f = payload || {};
  const fromDate = f.fromDate ? new Date(f.fromDate) : null;
  const toDate = f.toDate ? new Date(f.toDate) : null;
  const status = String(f.status || "").trim();

  const rows = readSheetObjects_(SHEETS.BOOTHS);
  return rows.filter(r => {
    if (status && !eq_(r.Status, status) && r.Status !== undefined) return false; // only if Status exists
    if (fromDate || toDate) {
      const d = toDate_(r.Booth_Date);
      if (!d) return false;
      if (fromDate && d < fromDate) return false;
      if (toDate && d > toDate) return false;
    }
    return true;
  });
}

function booths_create_(payload) {
  requireAdmin_();
  const p = payload || {};
  const boothDate = p.Booth_Date || p.boothDate;
  const location = String(p.Location || p.location || "").trim();
  const startTime = p.Start_Time || p.startTime || "";
  const endTime = p.End_Time || p.endTime || "";

  if (!boothDate || !location) throw new Error("Booth_Date and Location are required");

  const rowObj = {
    Booth_ID: makeId_("BOOTH"),
    Booth_Date: boothDate,
    Location: location,
    Start_Time: startTime,
    End_Time: endTime
    // Booth_Display is formula-driven; we do not set it
  };

  const sheet = getSheet_(SHEETS.BOOTHS);
  const row = appendByHeader_(sheet, rowObj);
  return Object.assign({ _row: row }, rowObj);
}

function booths_update_(payload) {
  requireAdmin_();
  const boothId = String(payload.Booth_ID || payload.boothId || "").trim();
  const patch = payload.patch || {};
  if (!boothId) throw new Error("Booth_ID is required");

  const sheet = getSheet_(SHEETS.BOOTHS);
  const row = findFirstRow_(sheet, { Booth_ID: boothId });
  if (!row) throw new Error(`Booth not found: ${boothId}`);

  // Only allow updating specific columns (safe)
  const allowed = {};
  ["Booth_Date", "Location", "Start_Time", "End_Time"].forEach(k => {
    if (patch[k] !== undefined) allowed[k] = patch[k];
  });

  updateByHeader_(sheet, row, allowed);
  return Object.assign({ _row: row, Booth_ID: boothId }, allowed);
}

/* ---- Signups ---- */

function booths_signups_list_(payload) {
  requireAdmin_();
  const boothId = String(payload.Booth_ID || payload.boothId || "").trim();
  if (!boothId) throw new Error("Booth_ID is required");
  const rows = readSheetObjects_(SHEETS.BOOTH_SIGNUPS);
  return rows.filter(r => eq_(r.Booth_ID, boothId));
}

function booths_signups_add_(payload) {
  // Either parent token OR admin
  const boothId = String(payload.Booth_ID || payload.boothId || "").trim();
  const girlName = String(payload.Girl_Name || payload.girlName || "").trim();
  if (!boothId || !girlName) throw new Error("Booth_ID and Girl_Name are required");

  const token = payload.token ? String(payload.token || "").trim() : "";
  if (token) {
    // validate girl belongs to this parent token
    const scope = requireParentToken_(token);
    const ok = scope.girls.some(g => eq_(g.Girl_Name, girlName));
    if (!ok) throw new Error("Girl_Name is not associated with this parent token");
  } else {
    requireAdmin_();
  }

  // Get booth details to copy into signup row
  const boothSheet = getSheet_(SHEETS.BOOTHS);
  const boothRow = findFirstRow_(boothSheet, { Booth_ID: boothId });
  if (!boothRow) throw new Error(`Booth not found: ${boothId}`);

  const booth = readRowObject_(boothSheet, boothRow);
  const signupKey = makeSignupKey_(girlName, boothId, booth.Booth_Date, booth.Start_Time, booth.End_Time);

  // De-dupe
  const signupSheet = getSheet_(SHEETS.BOOTH_SIGNUPS);
  const existing = readSheetObjects_(SHEETS.BOOTH_SIGNUPS).some(r => eq_(r.Signup_Key, signupKey));
  if (existing) return { alreadyExists: true, Signup_Key: signupKey };

  const rowObj = {
    Booth_ID: boothId,
    Booth_Date: booth.Booth_Date,
    Location: booth.Location,
    Start_Time: booth.Start_Time,
    End_Time: booth.End_Time,
    Girl_Name: girlName,
    Signup_Key: signupKey
  };

  const row = appendByHeader_(signupSheet, rowObj);
  return Object.assign({ _row: row }, rowObj);
}

function booths_signups_remove_(payload) {
  requireAdmin_();
  const key = String(payload.Signup_Key || payload.signupKey || "").trim();
  if (!key) throw new Error("Signup_Key is required");

  const sheet = getSheet_(SHEETS.BOOTH_SIGNUPS);
  const row = findFirstRow_(sheet, { Signup_Key: key });
  if (!row) return { ok: true, removed: false };

  sheet.deleteRow(row);
  return { ok: true, removed: true, Signup_Key: key };
}

/* ---- Counts / Sales posting ---- */

function booths_counts_get_(payload) {
  requireAdmin_();
  const boothId = String(payload.Booth_ID || payload.boothId || "").trim();
  if (!boothId) throw new Error("Booth_ID is required");
  const rows = readSheetObjects_(SHEETS.BOOTH_COUNTS);
  return rows.filter(r => eq_(r.Booth_ID, boothId));
}

function booths_counts_save_(payload) {
  requireAdmin_();
  const boothId = String(payload.Booth_ID || payload.boothId || "").trim();
  const rows = payload.rows || [];
  if (!boothId) throw new Error("Booth_ID is required");
  if (!Array.isArray(rows) || rows.length === 0) throw new Error("rows[] is required");

  const sheet = getSheet_(SHEETS.BOOTH_COUNTS);

  // Need booth metadata for missing columns (date/location)
  const boothSheet = getSheet_(SHEETS.BOOTHS);
  const boothRow = findFirstRow_(boothSheet, { Booth_ID: boothId });
  if (!boothRow) throw new Error(`Booth not found: ${boothId}`);
  const booth = readRowObject_(boothSheet, boothRow);

  const out = [];
  for (const r of rows) {
    const cookieType = String(r.Cookie_Type || r.cookieType || "").trim();
    if (!cookieType) continue;

    // Find existing count row by Booth_ID + Cookie_Type
    const existingRow = findFirstRow_(sheet, { Booth_ID: boothId, Cookie_Type: cookieType });
    const patch = {
      Booth_ID: boothId,
      Booth_Date: booth.Booth_Date,
      Location: booth.Location,
      Cookie_Type: cookieType,
      Count_Before: (r.Count_Before !== undefined ? r.Count_Before : r.countBefore),
      Count_After: (r.Count_After !== undefined ? r.Count_After : r.countAfter),
      Status: (r.Status !== undefined ? r.Status : r.status),
      Notes: (r.Notes !== undefined ? r.Notes : r.notes)
      // Boxes_Sold is formula-driven in many setups; if not, we can set if provided
    };

    if (existingRow) {
      updateByHeader_(sheet, existingRow, patch);
      out.push(Object.assign({ _row: existingRow }, patch));
    } else {
      const newRow = appendByHeader_(sheet, patch);
      out.push(Object.assign({ _row: newRow }, patch));
    }
  }
  return out;
}

function booths_counts_finalizeAndPostSales_(payload) {
  requireAdmin_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const boothId = String(payload.Booth_ID || payload.boothId || "").trim();
    if (!boothId) throw new Error("Booth_ID is required");

    const countsSheet = getSheet_(SHEETS.BOOTH_COUNTS);
    const salesSheet = getSheet_(SHEETS.BOOTH_SALES);
    const movesSheet = getSheet_(SHEETS.INVENTORY_MOVES);

    const boothSheet = getSheet_(SHEETS.BOOTHS);
    const boothRow = findFirstRow_(boothSheet, { Booth_ID: boothId });
    if (!boothRow) throw new Error(`Booth not found: ${boothId}`);
    const booth = readRowObject_(boothSheet, boothRow);

    const counts = readSheetObjects_(SHEETS.BOOTH_COUNTS).filter(r => eq_(r.Booth_ID, boothId));
    if (!counts.length) throw new Error("No Booth_Inventory_Counts rows found for this booth.");

    const now = new Date();
    const email = getUserEmail_();

    const posted = [];
    const invMoves = [];

    for (const r of counts) {
      const sold = n_(r.Boxes_Sold);
      const cookieType = String(r.Cookie_Type || "").trim();
      if (!cookieType || sold <= 0) continue;

      // Idempotency check at row level
      if (truthy_(r.Posted_Sold)) continue;

      // Append to Booth_Sales_Log
      const saleRowObj = {
        Booth_ID: boothId,
        Booth_Date: booth.Booth_Date,
        Location: booth.Location,
        Cookie_Type: cookieType,
        Boxes_Sold: sold,
        Notes: "Posted from Booth_Inventory_Counts",
        Posted_At: now,
        Posted_Boxes: sold
      };
      const saleRow = appendByHeader_(salesSheet, saleRowObj);
      posted.push(Object.assign({ _row: saleRow }, saleRowObj));

      // Inventory movement reduction
      const move = {
        Move_ID: makeId_("MOVE"),
        Move_Date: now,
        Move_Type: MOVE_TYPES.BOOTH_SALE,
        From_Entity: ENTITIES.TROOP,
        To_Entity: `Booth:${boothId}`,
        Parent_Name: "",
        Girl_Name: "",
        Cookie_Type: cookieType,
        Boxes: sold,
        Notes: `Booth sale posted (${booth.Location})`,
        Entered_By: email || "",
        Entered_At: now
      };
      const moveRow = appendByHeader_(movesSheet, move);
      invMoves.push(Object.assign({ _row: moveRow }, move));

      // Mark Posted on the count row
      const rowIndex = findFirstRow_(countsSheet, { Booth_ID: boothId, Cookie_Type: cookieType });
      if (rowIndex) {
        updateByHeader_(countsSheet, rowIndex, { Posted_Sold: "TRUE", Posted_At: now });
      }
    }

    return { salesRows: posted, inventoryMovements: invMoves };
  } finally {
    lock.releaseLock();
  }
}

/* ---- Cash ---- */

function booths_cash_get_(payload) {
  requireAdmin_();
  const boothId = String(payload.Booth_ID || payload.boothId || "").trim();
  if (!boothId) throw new Error("Booth_ID is required");
  const rows = readSheetObjects_(SHEETS.BOOTH_CASH).filter(r => eq_(r.Booth_ID, boothId));
  return rows[0] || null;
}

function booths_cash_save_(payload) {
  requireAdmin_();
  const boothId = String(payload.Booth_ID || payload.boothId || "").trim();
  if (!boothId) throw new Error("Booth_ID is required");

  const patch = {};
  ["Cash_Before", "Cash_After", "Digital_Collected", "Military_Donation_$"].forEach(k => {
    if (payload[k] !== undefined) patch[k] = payload[k];
  });

  const sheet = getSheet_(SHEETS.BOOTH_CASH);
  const row = findFirstRow_(sheet, { Booth_ID: boothId });

  if (row) {
    updateByHeader_(sheet, row, patch);
    return Object.assign({ _row: row, Booth_ID: boothId }, patch);
  }

  // Create row if missing
  const boothSheet = getSheet_(SHEETS.BOOTHS);
  const boothRow = findFirstRow_(boothSheet, { Booth_ID: boothId });
  if (!boothRow) throw new Error(`Booth not found: ${boothId}`);
  const booth = readRowObject_(boothSheet, boothRow);

  const rowObj = Object.assign({
    Booth_ID: boothId,
    Booth_Date: booth.Booth_Date,
    Booth_Location: booth.Location
  }, patch);

  const newRow = appendByHeader_(sheet, rowObj);
  return Object.assign({ _row: newRow }, rowObj);
}

/* ---- Allocations / Credits ---- */

function booths_allocations_get_(payload) {
  requireAdmin_();
  const boothId = String(payload.Booth_ID || payload.boothId || "").trim();
  if (!boothId) throw new Error("Booth_ID is required");
  return readSheetObjects_(SHEETS.BOOTH_ALLOC).filter(r => eq_(r.Booth_ID, boothId));
}

function booths_allocations_save_(payload) {
  requireAdmin_();
  const boothId = String(payload.Booth_ID || payload.boothId || "").trim();
  const allocations = payload.allocations || [];
  if (!boothId) throw new Error("Booth_ID is required");
  if (!Array.isArray(allocations) || allocations.length === 0) throw new Error("allocations[] is required");

  const boothSheet = getSheet_(SHEETS.BOOTHS);
  const boothRow = findFirstRow_(boothSheet, { Booth_ID: boothId });
  if (!boothRow) throw new Error(`Booth not found: ${boothId}`);
  const booth = readRowObject_(boothSheet, boothRow);

  const sheet = getSheet_(SHEETS.BOOTH_ALLOC);
  const out = [];

  for (const a of allocations) {
    const cookieType = String(a.Cookie_Type || a.cookieType || "").trim();
    const girlName = String(a.Girl_Name || a.girlName || "").trim();
    const boxes = n_(a.Boxes_Allocated || a.boxesAllocated);
    if (!cookieType || !girlName || boxes <= 0) continue;

    // Use existing Allocation_ID if provided; else upsert by booth+cookie+girl
    const allocId = String(a.Allocation_ID || a.allocationId || "").trim();
    let rowIndex = 0;
    if (allocId) {
      rowIndex = findFirstRow_(sheet, { Allocation_ID: allocId });
    } else {
      rowIndex = findFirstRow_(sheet, { Booth_ID: boothId, Cookie_Type: cookieType, Girl_Name: girlName });
    }

    const rowObj = {
      Allocation_ID: allocId || makeId_("ALLOC"),
      Booth_ID: boothId,
      Booth_Date: booth.Booth_Date,
      Location: booth.Location,
      Cookie_Type: cookieType,
      Girl_Name: girlName,
      Boxes_Allocated: boxes,
      Allocation_Source: String(a.Allocation_Source || a.allocationSource || "Booth").trim(),
      Status: String(a.Status || a.status || "Draft").trim(),
      Notes: String(a.Notes || a.notes || "").trim()
      // Sales_Row optional; leave untouched unless provided
    };
    if (a.Sales_Row !== undefined) rowObj.Sales_Row = a.Sales_Row;

    if (rowIndex) {
      updateByHeader_(sheet, rowIndex, rowObj);
      out.push(Object.assign({ _row: rowIndex }, rowObj));
    } else {
      const newRow = appendByHeader_(sheet, rowObj);
      out.push(Object.assign({ _row: newRow }, rowObj));
    }
  }

  return out;
}

function booths_allocations_postCredits_(payload) {
  requireAdmin_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const boothId = String(payload.Booth_ID || payload.boothId || "").trim();
    if (!boothId) throw new Error("Booth_ID is required");

    const config = getConfigMap_();
    const boxPrice = getBoxPrice_(config);

    // Validate totals: allocations must match sold per cookie
    const soldRows = readSheetObjects_(SHEETS.BOOTH_SALES).filter(r => eq_(r.Booth_ID, boothId));
    const soldByCookie = sumBy_(soldRows, "Cookie_Type", "Boxes_Sold");

    // If no Booth_Sales_Log rows exist yet, fall back to Booth_Inventory_Counts Boxes_Sold
    if (Object.keys(soldByCookie).length === 0) {
      const counts = readSheetObjects_(SHEETS.BOOTH_COUNTS).filter(r => eq_(r.Booth_ID, boothId));
      const countsByCookie = sumBy_(counts, "Cookie_Type", "Boxes_Sold");
      Object.assign(soldByCookie, countsByCookie);
    }

    const allocSheet = getSheet_(SHEETS.BOOTH_ALLOC);
    const allocations = readSheetObjects_(SHEETS.BOOTH_ALLOC).filter(r => eq_(r.Booth_ID, boothId));

    // Only post allocations not yet posted
    const toPost = allocations.filter(r => !truthy_(r.Posted_At) && lc_(r.Status) !== "posted");

    if (!toPost.length) return { ok: true, message: "Nothing to post (all allocations already posted)." };

    // Validate sums
    const allocByCookie = sumBy_(toPost, "Cookie_Type", "Boxes_Allocated");
    for (const cookieType of Object.keys(soldByCookie)) {
      const sold = n_(soldByCookie[cookieType]);
      const alloc = n_(allocByCookie[cookieType]);
      if (sold !== alloc) {
        throw new Error(`Allocation mismatch for ${cookieType}: sold=${sold} allocated=${alloc}`);
      }
    }

    // Need booth metadata
    const boothSheet = getSheet_(SHEETS.BOOTHS);
    const boothRow = findFirstRow_(boothSheet, { Booth_ID: boothId });
    if (!boothRow) throw new Error(`Booth not found: ${boothId}`);
    const booth = readRowObject_(boothSheet, boothRow);

    const now = new Date();
    const email = getUserEmail_();

    // Log credits
    const boothCreditsSheet = getSheet_(SHEETS.GIRL_BOOTH_CREDITS);
    const totalCreditsSheet = getSheet_(SHEETS.GIRL_TOTAL_CREDITS);

    // Prepare Dollar_Value formula template if present
    const totalMap = getHeaderMap_(totalCreditsSheet);
    const colDollar = totalMap["dollar_value"] || 0;
    const dollarFormulaR1C1 = (colDollar ? totalCreditsSheet.getRange(2, colDollar).getFormulaR1C1() : "");

    const boothCreditsOut = [];
    const totalCreditsOut = [];

    for (const r of toPost) {
      const allocId = String(r.Allocation_ID || "").trim();
      const girlName = String(r.Girl_Name || "").trim();
      const cookieType = String(r.Cookie_Type || "").trim();
      const boxes = n_(r.Boxes_Allocated);
      if (!girlName || !cookieType || boxes <= 0) continue;

      // Booth credits log row
      const boothCreditRow = {
        Timestamp: now,
        Girl_Name: girlName,
        Cookie_Type: cookieType,
        Boxes_Credited: boxes,
        Credit_Source: "BOOTH_SALE",
        Booth_ID: boothId,
        Booth_Date: booth.Booth_Date,
        Location: booth.Location,
        Allocation_ID: allocId,
        Notes: String(r.Notes || "").trim(),
        Entered_By: email || "",
        Entered_At: now
      };
      const bcRow = appendByHeader_(boothCreditsSheet, boothCreditRow);
      boothCreditsOut.push(Object.assign({ _row: bcRow }, boothCreditRow));

      // Total credits log row (Dollar_Value handled by formula template)
      const totalRowObj = {
        Timestamp: now,
        Girl_Name: girlName,
        Parent_Name: "", // workbook formulas/views can derive parent; leaving blank is safer than guessing
        Cookie_Type: cookieType,
        Boxes_Credited: boxes,
        Dollar_Value: "", // leave blank; restore formula if available
        Credit_Source: "BOOTH_SALE",
        Notes: `Booth ${boothId} - ${booth.Location}`,
        Entered_By: email || "",
        Entered_At: now
      };
      const tcRow = appendByHeader_(totalCreditsSheet, totalRowObj);

      // Restore Dollar_Value formula if we have a template
      if (colDollar && dollarFormulaR1C1) {
        totalCreditsSheet.getRange(tcRow, colDollar).setFormulaR1C1(dollarFormulaR1C1);
      }

      totalCreditsOut.push(Object.assign({ _row: tcRow }, totalRowObj));
    }

    // Mark allocations posted
    for (const r of toPost) {
      const rowIndex = findFirstRow_(allocSheet, { Allocation_ID: r.Allocation_ID });
      if (rowIndex) updateByHeader_(allocSheet, rowIndex, { Posted_At: now, Status: "Posted" });
    }

    return { boothCredits: boothCreditsOut, totalCredits: totalCreditsOut, boxPrice };
  } finally {
    lock.releaseLock();
  }
}

/* ===========================
   HELPERS: AUTH / SCOPE
=========================== */

function requireAdmin_() {
  const email = getUserEmail_();
  if (!isAdminEmail_(email)) throw new Error("Unauthorized: admin access required.");
  return true;
}

function isAdminEmail_(email) {
  const e = String(email || "").trim().toLowerCase();
  if (!e) return false;

  // Try Config allowlist first
  const cfg = getConfigMap_();
  const keysToTry = ["admin_emails", "cookie_mom_emails", "admins", "admin_email"];
  for (const k of keysToTry) {
    const v = String(cfg[k] || "").trim();
    if (v) {
      const list = v.split(/[,;\n]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
      if (list.includes(e)) return true;
    }
  }

  // Fallback: allow spreadsheet owner and current user (owner usually)
  try {
    const owner = SpreadsheetApp.getActive().getOwner();
    const ownerEmail = owner && owner.getEmail ? String(owner.getEmail() || "").trim().toLowerCase() : "";
    if (ownerEmail && ownerEmail === e) return true;
  } catch (_) { /* ignore */ }

  // If no allowlist set, safest behavior is: only allow active user if they are the effective owner/editor.
  // (In practice, you will add your email(s) to Config to avoid surprises.)
  return false;
}

function requireParentToken_(token) {
  if (!token) throw new Error("Missing token");

  const pg = readSheetObjects_(SHEETS.PARENTS_GIRLS);
  const rows = pg.filter(r => String(r.Parent_Token || "").trim() === token);
  if (!rows.length) throw new Error("Invalid token");

  // Parent scope derived from the first match
  const parentName = String(rows[0].Parent_Name || "").trim();
  const parentEmail = String(rows[0].Parent_Email || "").trim();
  const girls = rows.map(r => ({
    Girl_Name: r.Girl_Name,
    Parent_Name: r.Parent_Name,
    Parent_Email: r.Parent_Email,
    Active: r.Active,
    Parent_Token: r.Parent_Token
  }));

  return { parentName, parentEmail, girls };
}

function getParentScopeByName_(parentName) {
  const pg = readSheetObjects_(SHEETS.PARENTS_GIRLS).filter(r => eq_(r.Parent_Name, parentName));
  if (!pg.length) return null;
  const token = String(pg[0].Parent_Token || "").trim();
  return { parentName, parentEmail: String(pg[0].Parent_Email || "").trim(), token };
}

/* ===========================
   HELPERS: CONFIG
=========================== */

function getConfigMap_() {
  const sheet = getSheet_(SHEETS.CONFIG);
  const { map, values } = readSheet_(sheet);

  // Accept "Key"/"Value" header names case-insensitively
  const colKey = map["key"] || 1;
  const colVal = map["value"] || 2;

  const out = {};
  for (const row of values) {
    const k = String(row[colKey - 1] || "").trim();
    if (!k) continue;
    out[lc_(k)] = row[colVal - 1];
  }
  return out;
}

function getBoxPrice_(cfg) {
  // Try a few likely keys, fallback to 6 (common) but prefers config
  const keys = ["box_price", "price_per_box", "cookie_price", "price"];
  for (const k of keys) {
    const v = cfg[k];
    const n = Number(v);
    if (isFinite(n) && n > 0) return n;
  }
  return 6;
}

/* ===========================
   HELPERS: SHEETS (header-based)
=========================== */

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Missing sheet: ${name}`);
  return sh;
}

function readSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = (lastCol > 0)
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim())
    : [];
  const map = {};
  headers.forEach((h, i) => {
    const key = norm_(h);
    if (key) map[key] = i + 1; // 1-based col index
  });
  const values = (lastRow >= 2 && lastCol > 0)
    ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
    : [];
  return { headers, map, values };
}

function getHeaderMap_(sheet) {
  return readSheet_(sheet).map;
}

function needCol_(map, key) {
  const k = norm_(key);
  const col = map[k];
  if (!col) throw new Error(`Missing required column: ${key}`);
  return col;
}

function readSheetObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const { headers, values } = readSheet_(sheet);

  const objs = [];
  for (let r = 0; r < values.length; r++) {
    const o = {};
    for (let c = 0; c < headers.length; c++) {
      o[headers[c]] = values[r][c];
    }
    // include row index for convenience
    o._row = r + 2;
    objs.push(o);
  }
  return objs;
}

function readRowObject_(sheet, rowIndex) {
  const { headers } = readSheet_(sheet);
  if (rowIndex < 2) throw new Error("rowIndex must be >= 2");
  const rowVals = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const o = {};
  for (let c = 0; c < headers.length; c++) o[headers[c]] = rowVals[c];
  o._row = rowIndex;
  return o;
}

function appendByHeader_(sheet, obj) {
  const { headers, map } = readSheet_(sheet);
  const row = nextEmptyRowByKeyCol_(sheet, 1); // scan col A for true last row

  const out = headers.map(h => {
    const key = norm_(h);
    // match either exact header property or normalized fallback
    if (Object.prototype.hasOwnProperty.call(obj, h)) return obj[h];
    for (const k in obj) {
      if (norm_(k) === key) return obj[k];
    }
    return "";
  });

  sheet.getRange(row, 1, 1, out.length).setValues([out]);
  return row;
}

function updateByHeader_(sheet, rowIndex, patch) {
  const { headers, map } = readSheet_(sheet);
  const rowVals = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  for (const k in patch) {
    if (!Object.prototype.hasOwnProperty.call(patch, k)) continue;
    const col = map[norm_(k)];
    if (!col) continue; // ignore unknown keys
    rowVals[col - 1] = patch[k];
  }

  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowVals]);
}

function findFirstRow_(sheet, criteria) {
  const { headers, map, values } = readSheet_(sheet);
  const critKeys = Object.keys(criteria || {});
  if (!critKeys.length) return 0;

  const crit = critKeys.map(k => ({ col: map[norm_(k)], val: criteria[k] }));
  if (crit.some(x => !x.col)) return 0;

  for (let i = 0; i < values.length; i++) {
    let ok = true;
    for (const c of crit) {
      const cell = values[i][c.col - 1];
      if (!eq_(cell, c.val)) { ok = false; break; }
    }
    if (ok) return i + 2;
  }
  return 0;
}

// Next empty row in a "key" column (prevents jumping past dragged formulas)
function nextEmptyRowByKeyCol_(sheet, keyColIndex) {
  const last = sheet.getLastRow();
  if (last < 2) return 2;
  const values = sheet.getRange(2, keyColIndex, last - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || "").trim() !== "") return i + 3;
  }
  return 2;
}

/* ===========================
   HELPERS: UTIL
=========================== */

function ok_(data) {
  return { ok: true, data: data === undefined ? null : data, error: null };
}

function err_(msg) {
  return { ok: false, data: null, error: String(msg || "Unknown error") };
}

function getUserEmail_() {
  // In consumer accounts, this can be blank depending on deployment access settings.
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "";
  } catch (_) {
    return "";
  }
}

function buildParentLink_(token) {
  // The frontend can build this too, but this is handy for admin screens.
  // We can't reliably know deployment URL here without Properties, so return query fragment:
  return `?page=parent&t=${encodeURIComponent(token)}`;
}

function makeId_(prefix) {
  const p = String(prefix || "ID").toUpperCase();
  return `${p}_${Utilities.getUuid().replace(/-/g, "").slice(0, 10)}`;
}

function makeToken_() {
  return Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
}

function makeSignupKey_(girlName, boothId, boothDate, startTime, endTime) {
  return [
    String(girlName || "").trim().toLowerCase(),
    String(boothId || "").trim().toLowerCase(),
    String(boothDate || "").trim().toLowerCase(),
    String(startTime || "").trim().toLowerCase(),
    String(endTime || "").trim().toLowerCase()
  ].join("|");
}

function sumBy_(rows, keyCol, valCol) {
  const out = {};
  for (const r of rows || []) {
    const k = String(r[keyCol] || "").trim();
    if (!k) continue;
    out[k] = (out[k] || 0) + n_(r[valCol]);
  }
  return out;
}

function toRowIndex_(v) {
  const n = Number(v);
  return isFinite(n) && n >= 2 ? Math.floor(n) : 0;
}

function toDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isFinite(d.getTime()) ? d : null;
}

function lc_(v) {
  return String(v || "").trim().toLowerCase();
}

function norm_(h) {
  return String(h || "").trim().toLowerCase().replace(/\s+/g, "_");
}

function eq_(a, b) {
  return String(a || "").trim() === String(b || "").trim();
}

function n_(v) {
  if (v === "" || v === null || v === undefined) return 0;
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function truthy_(v) {
  if (v === true) return true;
  const s = String(v || "").trim().toLowerCase();
  return s === "true" || s === "yes" || s === "y" || s === "1";
}
