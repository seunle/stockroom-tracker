/**
 * STOCKROOM TRACKER — Google Sheets backend (matches your sheet exactly)
 *
 * Your sheet's columns, in order (row 1 headers):
 *   ID | Location | Photo | Name | Type | Size | Active |
 *   Min Count Tangle | Min Count Parkway | Min Count Bears |
 *   Display Tangle | Display Parkway | Display Bears | Price |
 *   Storage Tangle | Storage Parkway | Storage Bears |
 *   Resupply Tangle | Resupply Parkway | Resupply Bears |
 *   On Hold Tangle | On Hold Parkway | On Hold Bears |
 *   Order Tangle | Order Parkway | Order Bears
 *
 * One row = one size of one item. E.g. "Lean Meat / Adult / S" and
 * "Lean Meat / Adult / M" are two separate rows that the app groups
 * together on screen because they share the same Name + Type + Location.
 *
 * PER-LOCATION DATA: Min Count, Display, Storage, Resupply, On Hold, and
 * Order each have their own column per location, so the same item+size
 * row can hold completely independent numbers for Tangle, Parkway, and
 * Bears — someone working at one location never overwrites another
 * location's counts, and each location can require a different minimum
 * quantity for the same size. The app always writes to whichever
 * location's column matches what's currently selected in its top bar (or,
 * in the New/Edit item form, whichever location tab is being edited).
 *
 * PRICE: a single shared value per row (not per-location), used only by
 * Hoodie-family items (Hoodie, Crewneck, Y Hoodie Crewneck, K Hoodie
 * Crewneck). The same price applies to every size of a given item —
 * entered once in the New/Edit item form and written to every size row
 * created for that item. Everything else leaves this at 0 (unused). This
 * column used to be called "Restock" and was unused.
 *
 * STORAGE: Storage Tangle, Storage Parkway, and Storage Bears each hold a
 * 1/0 flag — whether that size is currently eligible to show in the app's
 * Storage tab for that location. Set to 1 when a below-minimum size is
 * explicitly reviewed in the Display popup, cleared back to 0 once it's
 * resolved (resupplied, put On Hold, etc.). This is real sheet data, not
 * anything device-local, so every device sees the same Storage tab for a
 * given location.
 *
 * ON HOLD: each location has its own column. On Hold Tangle, On Hold
 * Parkway, and On Hold Bears are read and written independently. A
 * location's On Hold is the shortfall between that location's required
 * resupply quantity and the quantity entered in that location's Resupply
 * popup, capped at 6.
 *
 * DISPLAY ORDER: Order Tangle, Order Parkway, and Order Bears each hold a
 * per-location sort number, set via drag-and-drop reordering in the app's
 * Edit tab. Each location has its own independent order — reordering at
 * Bears never touches Tangle's or Parkway's numbers. Blank/empty means
 * "not manually ordered yet" — those items just keep their natural sheet
 * row order, sorted after anything that HAS been manually positioned.
 *
 * DIRECT vs INDIRECT SYSTEM: every Type uses one of two fundamentally
 * different workflows. Indirect is the original one — Min Count, Display,
 * Storage, the whole "check the floor, flag it below minimum, then
 * resupply" pipeline. Direct skips all of that: a Price instead of
 * per-location minimums, and requests go straight from the app's Request
 * tab to that location's Resupply (or On Hold, if that location already
 * has some) with no other side effects — Display is never touched for a
 * Direct-system item, deliberately, since the sheet's Display columns
 * aren't used for them at all.
 *
 * Which system a Type uses is no longer fixed by its name. HOODIE_TYPES
 * (Hoodie, Crewneck, Y Hoodie Crewneck, K Hoodie Crewneck) below is only
 * each Type's DEFAULT — Direct for those four, Indirect for everything
 * else — and can be overridden per-Type via the in-app toggle (Display/
 * Request tab), which writes to the separate "Type Systems" sheet tab (see
 * getTypeSystemsSheet_/effectiveSystemForType_). A Type with no row there
 * just uses its default. Direct-system item rows always get Min Count
 * forced to 1 (not 0) on save — irrelevant while Direct (Active, set via
 * the size checkboxes, governs availability instead), but a meaningful
 * starting minimum instead of the 0 sentinel if that Type ever gets
 * switched to Indirect later.
 *
 * PHOTOS: uploaded to a Google Drive folder this script creates, and the
 * Photo cell just holds that file's link as plain text.
 *
 * LOCATION CODES: single letters combined, e.g. "T" = Tangle only,
 * "TPB" = all three. T = Tangle, P = Parkway, B = Bears.
 *
 * SETUP:
 * 1. Open your existing sheet (the one with your columns already in it).
 * 2. Extensions -> Apps Script. Delete any starter code, paste this whole file in.
 * 3. Deploy -> New deployment -> gear icon -> "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 4. Deploy, then authorize permissions (this needs Drive access too, for photos).
 * 5. Copy the Web app URL into the app's Settings screen.
 *
 * After editing this script again later: Deploy -> Manage deployments ->
 * pencil icon -> New version -> Deploy (keeps the same URL).
 *
 * IF UPGRADING FROM AN OLDER SHEET: the column POSITIONS must match the
 * HEADERS array below exactly, since col_() locates columns by position in
 * that array, not by re-reading your sheet's actual header text. Rename
 * your sheet's "Restock" column (column N) to "Price" — same position,
 * new name and new purpose.
 */

const SHEET_NAME = 'Inventory'; // change this if your tab is named differently
const HEADERS = [
  'ID','Location','Photo','Name','Type','Size','Active',
  'Min Count Tangle','Min Count Parkway','Min Count Bears',
  'Display Tangle','Display Parkway','Display Bears','Price',
  'Storage Tangle','Storage Parkway','Storage Bears',
  'Resupply Tangle','Resupply Parkway','Resupply Bears',
  'On Hold Tangle','On Hold Parkway','On Hold Bears',
  'Order Tangle','Order Parkway','Order Bears',
  'Checked Tangle','Checked Parkway','Checked Bears'
];
const DRIVE_FOLDER_NAME = 'Stockroom Tracker Photos';
const LOCATION_CODES = { Tangle: 'T', Parkway: 'P', Bears: 'B' };
const HOODIE_TYPES = ['Hoodie', 'Crewneck', 'Y Hoodie Crewneck', 'K Hoodie Crewneck']; // every FIXED_TYPES value that gets the alternate Price/Request workflow — treated identically, each still its own separate category
// Maps a location code to its own Min Count / Display / Storage / Resupply column.
const MIN_COUNT_COL_BY_CODE = { T: 'Min Count Tangle', P: 'Min Count Parkway', B: 'Min Count Bears' };
const DISPLAY_COL_BY_CODE  = { T: 'Display Tangle',  P: 'Display Parkway',  B: 'Display Bears' };
const STORAGE_COL_BY_CODE  = { T: 'Storage Tangle',  P: 'Storage Parkway',  B: 'Storage Bears' };
const RESUPPLY_COL_BY_CODE = { T: 'Resupply Tangle', P: 'Resupply Parkway', B: 'Resupply Bears' };
const ON_HOLD_COL_BY_CODE  = { T: 'On Hold Tangle', P: 'On Hold Parkway', B: 'On Hold Bears' };
const ORDER_COL_BY_CODE    = { T: 'Order Tangle', P: 'Order Parkway', B: 'Order Bears' };
// "Checked" per location — a real sheet column now (not client-only
// localStorage), so it survives an app restart and is shared across every
// device instead of living in just the one browser that checked it. Only
// ever cleared by an explicit uncheck (setChecked) or a Refresh Session
// press scoped to one Type+location (clearCheckedForType) — see doPost.
const CHECKED_COL_BY_CODE  = { T: 'Checked Tangle', P: 'Checked Parkway', B: 'Checked Bears' };

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

// DEVICE APPROVAL — a separate sheet tab, "Devices", acts as an allowlist.
// Every device that opens the app generates its own random token (stored
// in that browser's localStorage) and sends it with every request. A
// device is only approved once its token appears in this tab with
// Approved = 1. To approve a new device: open it once (it'll show its own
// token on a "not approved" screen), then add a row here with a name you
// choose, that exact token, and 1 in Approved. To revoke a device, delete
// its row or set Approved to 0 — takes effect on its very next request, no
// redeploy needed.
const DEVICES_SHEET_NAME = 'Devices';
function getDevicesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DEVICES_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(DEVICES_SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(['Device Name', 'Token', 'Approved']);
  return sheet;
}
function isTokenApproved_(token) {
  if (!token) return false;
  const sheet = getDevicesSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]) === String(token) && Number(data[i][2]) === 1) return true;
  }
  return false;
}

// TYPE SYSTEMS — a separate sheet tab, "Type Systems", holds the per-Type
// Direct/Indirect override set via the in-app toggle (Display/Request
// tab). Only holds a row for a Type once it's been explicitly toggled away
// from its own default — a Type with no row here just uses its default
// (Direct for HOODIE_TYPES, Indirect for everything else). Shared across
// every device on purpose: the two systems interpret the exact same row
// data completely differently, so devices disagreeing here would corrupt
// how data gets read, not just look different.
const TYPE_SYSTEMS_SHEET_NAME = 'Type Systems';
function getTypeSystemsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TYPE_SYSTEMS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(TYPE_SYSTEMS_SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(['Type', 'System']);
  return sheet;
}
// Returns { TypeName: 'direct'|'indirect', ... } — only explicit overrides,
// not every FIXED_TYPES value (the client already knows each type's
// default and only needs to know what's been overridden).
function readTypeSystemOverrides_() {
  const sheet = getTypeSystemsSheet_();
  const lastRow = sheet.getLastRow();
  const overrides = {};
  if (lastRow < 2) return overrides;
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    const type = data[i][0];
    const system = data[i][1];
    if (type && (system === 'direct' || system === 'indirect')) overrides[type] = system;
  }
  return overrides;
}
function effectiveSystemForType_(type, overrides) {
  if (overrides && overrides[type]) return overrides[type];
  return HOODIE_TYPES.indexOf(type) !== -1 ? 'direct' : 'indirect';
}

function getDriveFolder_() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function col_(name) { return HEADERS.indexOf(name) + 1; }

// Validates a Photo cell's value: it's just a plain text URL, so this is a
// direct check with no async/processing-lag concerns. Takes the value
// itself (already pulled via the bulk range read in readRows_) rather than
// re-fetching the cell — re-fetching each row's Photo cell individually
// was the classic Apps Script anti-pattern (one Sheets-service call per
// row on top of the bulk read that already has this data), and at a few
// hundred+ rows that overhead is the difference between doGet() feeling
// instant and feeling sluggish.
function isPhotoUrl_(value) {
  return (typeof value === 'string' && value.indexOf('http') === 0) ? value : '';
}

function nextId_(sheet) {
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const n = Number(data[i][0]);
    if (!isNaN(n) && n > max) max = n;
  }
  return max + 1;
}

function readRows_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = HEADERS.length;
  if (lastRow < 2) return [];
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const values = range.getValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[col_('ID')-1] === '' || row[col_('ID')-1] === null) continue; // skip blank rows
    rows.push({
      ID: row[col_('ID')-1],
      Location: row[col_('Location')-1] || '',
      Photo: isPhotoUrl_(row[col_('Photo')-1]),
      Name: row[col_('Name')-1] || '',
      Type: row[col_('Type')-1] || '',
      Size: row[col_('Size')-1] || '',
      Active: row[col_('Active')-1] === '' ? 1 : Number(row[col_('Active')-1]),
      MinCountTangle: Number(row[col_('Min Count Tangle')-1] || 0),
      MinCountParkway: Number(row[col_('Min Count Parkway')-1] || 0),
      MinCountBears: Number(row[col_('Min Count Bears')-1] || 0),
      DisplayTangle: Number(row[col_('Display Tangle')-1] || 0),
      DisplayParkway: Number(row[col_('Display Parkway')-1] || 0),
      DisplayBears: Number(row[col_('Display Bears')-1] || 0),
      Price: Number(row[col_('Price')-1] || 0),
      StorageTangle: Number(row[col_('Storage Tangle')-1] || 0),
      StorageParkway: Number(row[col_('Storage Parkway')-1] || 0),
      StorageBears: Number(row[col_('Storage Bears')-1] || 0),
      ResupplyTangle: Number(row[col_('Resupply Tangle')-1] || 0),
      ResupplyParkway: Number(row[col_('Resupply Parkway')-1] || 0),
      ResupplyBears: Number(row[col_('Resupply Bears')-1] || 0),
      OnHoldTangle: Number(row[col_('On Hold Tangle')-1] || 0),
      OnHoldParkway: Number(row[col_('On Hold Parkway')-1] || 0),
      OnHoldBears: Number(row[col_('On Hold Bears')-1] || 0),
      // null = not manually ordered yet for that location (distinct from an
      // explicit order of 0) — the client falls back to natural row order.
      OrderTangle: row[col_('Order Tangle')-1] === '' ? null : Number(row[col_('Order Tangle')-1]),
      OrderParkway: row[col_('Order Parkway')-1] === '' ? null : Number(row[col_('Order Parkway')-1]),
      OrderBears: row[col_('Order Bears')-1] === '' ? null : Number(row[col_('Order Bears')-1]),
      CheckedTangle: Number(row[col_('Checked Tangle')-1] || 0),
      CheckedParkway: Number(row[col_('Checked Parkway')-1] || 0),
      CheckedBears: Number(row[col_('Checked Bears')-1] || 0)
    });
  }
  return rows;
}

function doGet(e) {
  try {
    const token = e && e.parameter ? e.parameter.token : null;
    if (!isTokenApproved_(token)) {
      return jsonOut_({ ok:false, unauthorized:true, token: token || '', error:'Device not approved' });
    }
    return jsonOut_({ ok:true, items: readRows_(), typeSystems: readTypeSystemOverrides_() });
  } catch (err) {
    return jsonOut_({ ok:false, error:String(err) });
  }
}

function findSheetRow_(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-indexed sheet row
  }
  return -1;
}

// Finds a row by its Name+Type+Location+Size combination (used by saveItem
// to decide whether a given size already exists for this item).
function findSizeRow_(sheet, name, type, location, size) {
  const data = sheet.getDataRange().getValues();
  const nameC = col_('Name')-1, typeC = col_('Type')-1, locC = col_('Location')-1, sizeC = col_('Size')-1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameC] === name && data[i][typeC] === type &&
        data[i][locC] === location && data[i][sizeC] === size) {
      return i + 1;
    }
  }
  return -1;
}

// Every existing row for a given Name+Type+Location, keyed by Size,
// regardless of Active state. Used by Hoodie's toggle-based size
// selection: any existing size NOT in the newly-submitted selection needs
// to be deactivated (never deleted), which requires knowing about rows
// the client isn't currently trying to create or update.
function findAllSizeRowsForItem_(sheet, name, type, location) {
  const data = sheet.getDataRange().getValues();
  const nameC = col_('Name')-1, typeC = col_('Type')-1, locC = col_('Location')-1, sizeC = col_('Size')-1;
  const map = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameC] === name && data[i][typeC] === type && data[i][locC] === location) {
      map[data[i][sizeC]] = i + 1;
    }
  }
  return map;
}

// Uploads a photo to Drive once and returns its durable, public link.
function uploadPhotoOnce_(base64Data, filenameHint) {
  const matches = base64Data.match(/^data:(image\/[a-zA-Z]+);base64,(.*)$/);
  const mime = matches ? matches[1] : 'image/jpeg';
  const raw = matches ? matches[2] : base64Data;
  const bytes = Utilities.base64Decode(raw);
  const blob = Utilities.newBlob(bytes, mime, (filenameHint || 'photo') + '.jpg');
  const folder = getDriveFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?sz=w1000&id=' + file.getId();
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (!isTokenApproved_(body.token)) {
      return jsonOut_({ ok:false, unauthorized:true, token: body.token || '', error:'Device not approved' });
    }
    const action = body.action;
    const sheet = getSheet_();
    let result = { ok:true };

    if (action === 'saveItem') {
      // Handles BOTH creating a new item and editing an existing one, for
      // BOTH regular items and Hoodies (a completely different payload
      // shape — see the branch below).
      // Regular item body: { matchIds, name, type, location,
      //          sizes: { XS: {T:1,P:2,B:0}, S: {T:0,P:0,B:3}, ... }
      //          (= Min Count per size, per location), photoBase64 (optional) }
      // Hoodie body: { matchIds, name, type:'Hoodie', location,
      //          hoodiePrice: 29.99, hoodieSizes: ['S','M','L'],
      //          photoBase64 (optional) }
      const matchIds = body.matchIds || [];
      const isNewItem = matchIds.length === 0;
      // Which workflow this SAVE uses is the Type's CURRENT effective
      // system (respecting any Direct/Indirect override from the in-app
      // toggle), not just raw HOODIE_TYPES membership — e.g. saving a
      // "Men" item after Men has been toggled to Direct must use the
      // Price/active-size-selection shape below, not the Min Count grid.
      const isHoodie = effectiveSystemForType_(body.type, readTypeSystemOverrides_()) === 'direct';

      const photoUrl = body.photoBase64 ? uploadPhotoOnce_(body.photoBase64, body.name) : null;

      // Step 1: rename/retype/relocate/re-photo every existing row in this group.
      if (!isNewItem) {
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (matchIds.indexOf(data[i][0]) !== -1) {
            const row = i + 1;
            sheet.getRange(row, col_('Name')).setValue(body.name);
            sheet.getRange(row, col_('Type')).setValue(body.type);
            sheet.getRange(row, col_('Location')).setValue(body.location);
            if (photoUrl) sheet.getRange(row, col_('Photo')).setValue(photoUrl);
          }
        }
      }

      const newIds = [];

      if (isHoodie) {
        // Hoodie path: one shared Price (not per-size, not per-location)
        // and a simple list of which sizes are selected — no Min Count,
        // no Display, no Storage involvement at all. Any size that WAS
        // active but is no longer in hoodieSizes gets deactivated (never
        // deleted), matching the "just deactivate it" instruction.
        const price = Number(body.hoodiePrice) || 0;
        const selectedSizes = body.hoodieSizes || [];
        const existingSizeRows = findAllSizeRowsForItem_(sheet, body.name, body.type, body.location);
        selectedSizes.forEach(size => {
          const foundRow = existingSizeRows[size];
          if (foundRow > 0) {
            sheet.getRange(foundRow, col_('Active')).setValue(1);
            sheet.getRange(foundRow, col_('Price')).setValue(price);
            if (photoUrl) sheet.getRange(foundRow, col_('Photo')).setValue(photoUrl);
            // Direct-system items don't use Min Count for anything while
            // Direct — availability is governed by Active (the size
            // checkboxes above) instead. It's force-set to 1 (never left
            // at 0) purely so that IF this type ever gets toggled to
            // Indirect later, the item starts with a real, meaningful
            // minimum instead of the sentinel 0 that now triggers the
            // Display popup's "no Min Count set" warning.
            sheet.getRange(foundRow, col_('Min Count Tangle')).setValue(1);
            sheet.getRange(foundRow, col_('Min Count Parkway')).setValue(1);
            sheet.getRange(foundRow, col_('Min Count Bears')).setValue(1);
          } else {
            const id = nextId_(sheet);
            const row = [];
            row[col_('ID')-1] = id;
            row[col_('Location')-1] = body.location || '';
            row[col_('Photo')-1] = photoUrl || '';
            row[col_('Name')-1] = body.name || '';
            row[col_('Type')-1] = body.type || '';
            row[col_('Size')-1] = size;
            row[col_('Active')-1] = 1;
            row[col_('Min Count Tangle')-1] = 1;
            row[col_('Min Count Parkway')-1] = 1;
            row[col_('Min Count Bears')-1] = 1;
            row[col_('Display Tangle')-1] = 0;
            row[col_('Display Parkway')-1] = 0;
            row[col_('Display Bears')-1] = 0;
            row[col_('Price')-1] = price;
            row[col_('Storage Tangle')-1] = 0;
            row[col_('Storage Parkway')-1] = 0;
            row[col_('Storage Bears')-1] = 0;
            row[col_('Resupply Tangle')-1] = 0;
            row[col_('Resupply Parkway')-1] = 0;
            row[col_('Resupply Bears')-1] = 0;
            row[col_('On Hold Tangle')-1] = 0;
            row[col_('On Hold Parkway')-1] = 0;
            row[col_('On Hold Bears')-1] = 0;
            row[col_('Order Tangle')-1] = '';
            row[col_('Order Parkway')-1] = '';
            row[col_('Order Bears')-1] = '';
            row[col_('Checked Tangle')-1] = 0;
            row[col_('Checked Parkway')-1] = 0;
            row[col_('Checked Bears')-1] = 0;
            sheet.appendRow(row);
            newIds.push(id);
          }
        });
        Object.keys(existingSizeRows).forEach(size => {
          if (selectedSizes.indexOf(size) === -1) {
            sheet.getRange(existingSizeRows[size], col_('Active')).setValue(0);
          }
        });

      } else {
        // Regular item path (unchanged): for each size, update its
        // per-location Min Count values if the row already exists, or
        // create a new row if it doesn't (covers both "brand new item"
        // and "adding a size that wasn't there before"). A size with
        // EVERY location's minimum at 0 is auto-deactivated (not
        // tracked); raising ANY location's minimum back above 0 later
        // auto-reactivates it. A brand-new size starts each location's
        // Display stocked to that SAME location's own minimum — only the
        // locations actually named in this item's Location field are ever
        // shown/edited by the app, so the others just sit unused.
        const sizes = body.sizes || {};
        Object.keys(sizes).forEach(size => {
          const valuesByLoc = sizes[size] || {};
          const vT = Number(valuesByLoc.T) || 0;
          const vP = Number(valuesByLoc.P) || 0;
          const vB = Number(valuesByLoc.B) || 0;
          const activeFlag = (vT > 0 || vP > 0 || vB > 0) ? 1 : 0;
          const foundRow = findSizeRow_(sheet, body.name, body.type, body.location, size);
          if (foundRow > 0) {
            sheet.getRange(foundRow, col_('Min Count Tangle')).setValue(vT);
            sheet.getRange(foundRow, col_('Min Count Parkway')).setValue(vP);
            sheet.getRange(foundRow, col_('Min Count Bears')).setValue(vB);
            sheet.getRange(foundRow, col_('Active')).setValue(activeFlag);
            if (photoUrl) sheet.getRange(foundRow, col_('Photo')).setValue(photoUrl);
          } else {
            const id = nextId_(sheet);
            const row = [];
            row[col_('ID')-1] = id;
            row[col_('Location')-1] = body.location || '';
            row[col_('Photo')-1] = photoUrl || '';
            row[col_('Name')-1] = body.name || '';
            row[col_('Type')-1] = body.type || '';
            row[col_('Size')-1] = size;
            row[col_('Active')-1] = activeFlag;
            row[col_('Min Count Tangle')-1] = vT;
            row[col_('Min Count Parkway')-1] = vP;
            row[col_('Min Count Bears')-1] = vB;
            row[col_('Display Tangle')-1] = vT;
            row[col_('Display Parkway')-1] = vP;
            row[col_('Display Bears')-1] = vB;
            row[col_('Price')-1] = 0;
            row[col_('Storage Tangle')-1] = 0;
            row[col_('Storage Parkway')-1] = 0;
            row[col_('Storage Bears')-1] = 0;
            row[col_('Resupply Tangle')-1] = 0;
            row[col_('Resupply Parkway')-1] = 0;
            row[col_('Resupply Bears')-1] = 0;
            row[col_('On Hold Tangle')-1] = 0;
            row[col_('On Hold Parkway')-1] = 0;
            row[col_('On Hold Bears')-1] = 0;
            row[col_('Order Tangle')-1] = '';
            row[col_('Order Parkway')-1] = '';
            row[col_('Order Bears')-1] = '';
            row[col_('Checked Tangle')-1] = 0;
            row[col_('Checked Parkway')-1] = 0;
            row[col_('Checked Bears')-1] = 0;
            sheet.appendRow(row);
            newIds.push(id);
          }
        });
      }

      result.newIds = newIds;
      if (photoUrl) result.photoUrl = photoUrl;

    } else if (action === 'setDisplayCount') {
      // The +/- steppers in the Display tab's popup write here. Always
      // scoped to one location — body.location must be 'T'|'P'|'B'.
      const col = DISPLAY_COL_BY_CODE[body.location];
      if (!col) throw new Error('A specific location (not "All") is required to set a display count.');
      const row = findSheetRow_(sheet, body.id);
      if (row < 0) throw new Error('Row not found');
      sheet.getRange(row, col_(col)).setValue(Math.max(0, Number(body.value)));

    } else if (action === 'setStorageFlag') {
      // Storage-tab eligibility, set from the Display popup in the app.
      // A 1/0 flag per location, stored as real sheet data (not on any one
      // device) so every device sees the same Storage tab for a location —
      // previously this lived only in the browser that did the check.
      const col = STORAGE_COL_BY_CODE[body.location];
      if (!col) throw new Error('A specific location (not "All") is required to set the Storage flag.');
      const row = findSheetRow_(sheet, body.id);
      if (row < 0) throw new Error('Row not found');
      sheet.getRange(row, col_(col)).setValue(body.active ? 1 : 0);

    } else if (action === 'commitDisplayReview') {
      // Everything from ONE Display-popup "Done"/"Checked" press, in a
      // SINGLE round trip. This used to be many separate setDisplayCount /
      // setStorageFlag / On Hold / Resupply calls sent one at a time —
      // reviewing a 7-size item could mean 8-10+ sequential network calls,
      // each with its own fixed Apps Script overhead, which is what
      // actually caused the visible lag (not sheet size — a bulk read/
      // write here costs about the same at 700 rows as at 7,000).
      // body: { location: 'T'|'P'|'B', updates: [{
      //   id, displayValue (optional), storageActive (0/1, optional),
      //   onHold (optional), resupplyIncrement (optional), checked (0/1, optional)
      // }, ...] }
      const displayCol = DISPLAY_COL_BY_CODE[body.location];
      const storageCol = STORAGE_COL_BY_CODE[body.location];
      const onHoldCol = ON_HOLD_COL_BY_CODE[body.location];
      const resupplyCol = RESUPPLY_COL_BY_CODE[body.location];
      const checkedCol = CHECKED_COL_BY_CODE[body.location];
      if (!displayCol || !storageCol || !onHoldCol || !resupplyCol) throw new Error('A specific location (not "All") is required.');
      const reviewUpdates = body.updates || [];
      const reviewData = sheet.getDataRange().getValues();
      const reviewRowById = {};
      for (let i = 1; i < reviewData.length; i++) { reviewRowById[String(reviewData[i][0])] = i + 1; }
      reviewUpdates.forEach(function(u) {
        const row = reviewRowById[String(u.id)];
        if (!row) return;
        if (u.displayValue !== undefined) {
          sheet.getRange(row, col_(displayCol)).setValue(Math.max(0, Number(u.displayValue) || 0));
        }
        if (u.storageActive !== undefined) {
          sheet.getRange(row, col_(storageCol)).setValue(u.storageActive ? 1 : 0);
        }
        if (u.onHold !== undefined) {
          sheet.getRange(row, col_(onHoldCol)).setValue(Math.min(6, Math.max(0, Number(u.onHold) || 0)));
        }
        if (u.resupplyIncrement) {
          const current = Number(sheet.getRange(row, col_(resupplyCol)).getValue() || 0);
          sheet.getRange(row, col_(resupplyCol)).setValue(Math.min(6, Math.max(0, current + Number(u.resupplyIncrement))));
        }
        if (u.checked !== undefined) {
          sheet.getRange(row, col_(checkedCol)).setValue(u.checked ? 1 : 0);
        }
      });

    } else if (action === 'requestResupply') {
      // Storage tab's Request/Resupply popup. Reads and bumps Display for
      // THE LOCATION THE APP IS CURRENTLY SHOWING (body.location: 'T'|'P'|'B'),
      // and writes the requested quantity into that same location's Resupply
      // column — this OVERWRITES any prior value in that specific column
      // (never adds to it), so re-requesting from the same location always
      // reflects only the latest request. Other locations' Display and
      // Resupply columns for this row are completely untouched.
      //
      // Display bump uses a diminishing-returns rule, with a floor of 2 on
      // the full-credit zone (storage is meant to hold a buffer of at least
      // 2): requesting up to max(Need, 2) gets the full Need added; every
      // unit requested beyond that reduces the credited increase by one,
      // down to zero.
      const displayCol = DISPLAY_COL_BY_CODE[body.location];
      const resupplyCol = RESUPPLY_COL_BY_CODE[body.location];
      if (!displayCol || !resupplyCol) throw new Error('A specific location (not "All") is required to request a resupply.');
      const row = findSheetRow_(sheet, body.id);
      if (row < 0) throw new Error('Row not found');
      const minCountCol = MIN_COUNT_COL_BY_CODE[body.location];
      if (!minCountCol) throw new Error('A specific location (not "All") is required to request a resupply.');
      const minCount = Number(sheet.getRange(row, col_(minCountCol)).getValue() || 0);
      const display = Number(sheet.getRange(row, col_(displayCol)).getValue() || 0);
      const onHoldCol = ON_HOLD_COL_BY_CODE[body.location];
      if (!onHoldCol) throw new Error('A specific location (not "All") is required for On Hold.');
      const onHold = Number(sheet.getRange(row, col_(onHoldCol)).getValue() || 0);
      // Storage's remaining need accounts for units already placed On Hold.
      // Those units should not be requested again.
      const need = Math.max(0, minCount - display - onHold);
      const reqQty = Math.min(6, Math.max(0, Number(body.requestQty) || 0)); // Resupply is capped at 6 per location
      const threshold = Math.max(need, 2);
      const increase = Math.max(0, need - Math.max(0, reqQty - threshold));
      const newDisplay = display + increase;
      sheet.getRange(row, col_(displayCol)).setValue(newDisplay);
      sheet.getRange(row, col_(resupplyCol)).setValue(reqQty);
      result.newDisplay = newDisplay;
      result.increase = increase;

    } else if (action === 'setResupplyForLocation') {
      // Resupply tab's own popup — directly edits one location's resupply
      // number (e.g. to mark a request fulfilled by setting it back to 0).
      // Also used by the On Hold resolve popup, one location at a time.
      //
      // On Hold is location-specific. Completing a Tangle/Parkway/Bears
      // resupply only changes that same location's Resupply and On Hold
      // columns; the other locations are left untouched.
      const resupplyCol = RESUPPLY_COL_BY_CODE[body.location];
      const onHoldCol = ON_HOLD_COL_BY_CODE[body.location];
      const displayCol = DISPLAY_COL_BY_CODE[body.location];
      if (!resupplyCol || !onHoldCol || !displayCol) throw new Error('Unknown location: ' + body.location);
      const row = findSheetRow_(sheet, body.id);
      if (row < 0) throw new Error('Row not found');
      // Completing the Resupply popup closes out the active location request.
      // Any shortfall is written only to that location's On Hold column,
      // capped at 6.
      sheet.getRange(row, col_(resupplyCol)).setValue(0);
      if (body.onHold !== undefined) {
        sheet.getRange(row, col_(onHoldCol)).setValue(
          Math.min(6, Math.max(0, Number(body.onHold) || 0))
        );
      }
      // The resupplied/stocked quantity is credited onto that location's
      // Display count too, capped at that location's own Min Count — the
      // client computes and sends the already-capped final value.
      if (body.displayValue !== undefined) {
        sheet.getRange(row, col_(displayCol)).setValue(Math.max(0, Number(body.displayValue) || 0));
      }

    } else if (action === 'requestHoodieSizes') {
      // Request tab's popup for Hoodie items. Adds the entered quantity
      // for each size DIRECTLY onto that location's Resupply — or that
      // location's On Hold instead, if it already has some — with none of
      // the diminishing-returns math or Display-count changes the regular
      // Storage/Display flow uses. Hoodie items skip that whole pipeline;
      // this is deliberately just a straight add.
      // body: { location: 'T'|'P'|'B', requests: [{id, qty}, ...], checkedIds: [id, ...] }
      const resupplyCol = RESUPPLY_COL_BY_CODE[body.location];
      const onHoldCol = ON_HOLD_COL_BY_CODE[body.location];
      const checkedCol = CHECKED_COL_BY_CODE[body.location];
      if (!resupplyCol || !onHoldCol) throw new Error('A specific location (not "All") is required to request.');
      const requests = body.requests || [];
      const reqData = sheet.getDataRange().getValues();
      const reqRowById = {};
      for (let i = 1; i < reqData.length; i++) { reqRowById[String(reqData[i][0])] = i + 1; }
      requests.forEach(function(r) {
        const row = reqRowById[String(r.id)];
        if (!row) return;
        const qty = Math.max(0, Number(r.qty) || 0);
        if (qty <= 0) return;
        const currentOnHold = Number(sheet.getRange(row, col_(onHoldCol)).getValue() || 0);
        if (currentOnHold > 0) {
          sheet.getRange(row, col_(onHoldCol)).setValue(Math.min(6, currentOnHold + qty));
        } else {
          const currentResupply = Number(sheet.getRange(row, col_(resupplyCol)).getValue() || 0);
          sheet.getRange(row, col_(resupplyCol)).setValue(Math.min(6, currentResupply + qty)); // Resupply is capped at 6 per location
        }
      });
      // Pressing Request always marks the WHOLE item checked for this
      // location — every size in the popup, not just the ones actually
      // requested (a size left at 0 was still reviewed).
      if (checkedCol) {
        (body.checkedIds || []).forEach(function(id) {
          const row = reqRowById[String(id)];
          if (row) sheet.getRange(row, col_(checkedCol)).setValue(1);
        });
      }

    } else if (action === 'setChecked') {
      // The standalone checkmark toggle in the Display/Request popups.
      // body: { ids: [id, ...], location: 'T'|'P'|'B', value: 0|1 } — sets
      // that ONE location's Checked column for every row in the group
      // (every size shares one checked state), so any device looking at
      // that same location sees the same checked/unchecked status.
      const checkedCol = CHECKED_COL_BY_CODE[body.location];
      if (!checkedCol) throw new Error('A specific location (not "All") is required to set Checked.');
      const ids = body.ids || [];
      const setData = sheet.getDataRange().getValues();
      const setRowById = {};
      for (let i = 1; i < setData.length; i++) { setRowById[String(setData[i][0])] = i + 1; }
      ids.forEach(function(id) {
        const row = setRowById[String(id)];
        if (row) sheet.getRange(row, col_(checkedCol)).setValue(body.value ? 1 : 0);
      });

    } else if (action === 'clearCheckedForType') {
      // Refresh Session button — deliberately scoped to just the ONE Type
      // (sidebar tab) and ONE location the user was looking at, not a
      // global clear. body: { type, location: 'T'|'P'|'B' }. Everything
      // else (other Types, other locations, even this same Type at a
      // different location) keeps whatever Checked state it already had.
      const checkedCol = CHECKED_COL_BY_CODE[body.location];
      if (!checkedCol) throw new Error('A specific location (not "All") is required to clear Checked.');
      const clearData = sheet.getDataRange().getValues();
      const typeColIdx = col_('Type') - 1;
      for (let i = 1; i < clearData.length; i++) {
        if (clearData[i][typeColIdx] === body.type) {
          sheet.getRange(i + 1, col_(checkedCol)).setValue(0);
        }
      }

    } else if (action === 'incrementResupplyForLocation') {
      // Display popup redirect: if a location already has an outstanding
      // Resupply request (Resupply column > 0) and a new Display-driven
      // shortfall shows up, that shortfall is added onto the EXISTING
      // Resupply number for that location instead of ever surfacing in
      // Storage. Deliberately separate from setResupplyForLocation above —
      // that action always closes a request out to 0 (by design, for the
      // Resupply tab's own "mark fulfilled" flow); this one only ever adds
      // to whatever's already there, computed from the sheet's own current
      // value (not a value the client precomputed) so two near-simultaneous
      // requests can't stomp on each other. Touches nothing else — no
      // Display, no On Hold.
      const resupplyCol = RESUPPLY_COL_BY_CODE[body.location];
      if (!resupplyCol) throw new Error('Unknown location: ' + body.location);
      const row = findSheetRow_(sheet, body.id);
      if (row < 0) throw new Error('Row not found');
      const current = Number(sheet.getRange(row, col_(resupplyCol)).getValue() || 0);
      const newValue = Math.min(6, Math.max(0, current + (Number(body.amount) || 0))); // Resupply is capped at 6 per location
      sheet.getRange(row, col_(resupplyCol)).setValue(newValue);
      result.newResupply = newValue;

    } else if (action === 'setActive') {
      const row = findSheetRow_(sheet, body.id);
      if (row < 0) throw new Error('Row not found');
      sheet.getRange(row, col_('Active')).setValue(body.active ? 1 : 0);

    } else if (action === 'setTypeSystem') {
      // The Direct/Indirect toggle on the Display/Request tab. body: {
      // type, system: 'direct'|'indirect'|null }. Upserts a row in the Type
      // Systems sheet — this is deliberately a SEPARATE sheet/tab from the
      // main Inventory one (see getTypeSystemsSheet_), since it's a
      // property of the Type itself, not of any individual item row.
      // system: null clears the override entirely (deletes the row) —
      // used when the app resets a Type back to its default the moment
      // you navigate away from it (see clearTypeSystemOverride).
      const typeSheet = getTypeSystemsSheet_();
      const lastRow = typeSheet.getLastRow();
      let foundRow = -1;
      if (lastRow >= 2) {
        const typeData = typeSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < typeData.length; i++) {
          if (typeData[i][0] === body.type) { foundRow = i + 2; break; }
        }
      }
      if (body.system === null || body.system === undefined) {
        if (foundRow > 0) typeSheet.deleteRow(foundRow);
      } else {
        if (body.system !== 'direct' && body.system !== 'indirect') throw new Error('system must be "direct", "indirect", or null');
        if (foundRow > 0) {
          typeSheet.getRange(foundRow, 2).setValue(body.system);
        } else {
          typeSheet.appendRow([body.type, body.system]);
        }
      }

    } else if (action === 'setGroupOrder') {
      // Display-tab custom ordering, set via drag-and-drop in the Edit
      // tab. body: { location: 'T'|'P'|'B', updates: [{id, order}, ...] }
      // — one update per ROW (a multi-size item sends the same order
      // value for every one of its size rows). One combined sheet read
      // builds an ID -> row-number map, then every update in the batch is
      // applied as an individual cell write — much cheaper than
      // re-scanning the whole sheet once per update (the same anti-pattern
      // already fixed in readRows_), especially since a single drag commit
      // can touch dozens of rows in one go.
      const orderCol = ORDER_COL_BY_CODE[body.location];
      if (!orderCol) throw new Error('A specific location (not "All") is required to set display order.');
      const updates = body.updates || [];
      const data = sheet.getDataRange().getValues();
      const rowById = {};
      for (let i = 1; i < data.length; i++) {
        rowById[String(data[i][0])] = i + 1; // 1-indexed sheet row
      }
      updates.forEach(function(u) {
        const row = rowById[String(u.id)];
        if (row) sheet.getRange(row, col_(orderCol)).setValue(Number(u.order) || 0);
      });

    } else if (action === 'deleteRow') {
      const row = findSheetRow_(sheet, body.id);
      if (row > 0) sheet.deleteRow(row);

    } else if (action === 'restoreRow') {
      // Used by Undo — recreates or overwrites a row with a prior full
      // snapshot. The client always sends the row's own field name "ID"
      // (uppercase, matching the rest of the schema), never a separate
      // lowercase "id".
      let row = findSheetRow_(sheet, body.ID);
      if (row < 0) {
        sheet.appendRow([body.ID,'','','','','',1, 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, '','','', 0,0,0]);
        row = sheet.getLastRow();
      }
      sheet.getRange(row, col_('Location')).setValue(body.Location||'');
      sheet.getRange(row, col_('Name')).setValue(body.Name||'');
      sheet.getRange(row, col_('Type')).setValue(body.Type||'');
      sheet.getRange(row, col_('Size')).setValue(body.Size||'');
      sheet.getRange(row, col_('Active')).setValue(body.Active===undefined?1:body.Active);
      sheet.getRange(row, col_('Min Count Tangle')).setValue(body.MinCountTangle||0);
      sheet.getRange(row, col_('Min Count Parkway')).setValue(body.MinCountParkway||0);
      sheet.getRange(row, col_('Min Count Bears')).setValue(body.MinCountBears||0);
      sheet.getRange(row, col_('Display Tangle')).setValue(body.DisplayTangle||0);
      sheet.getRange(row, col_('Display Parkway')).setValue(body.DisplayParkway||0);
      sheet.getRange(row, col_('Display Bears')).setValue(body.DisplayBears||0);
      sheet.getRange(row, col_('Price')).setValue(body.Price||0);
      sheet.getRange(row, col_('Storage Tangle')).setValue(body.StorageTangle||0);
      sheet.getRange(row, col_('Storage Parkway')).setValue(body.StorageParkway||0);
      sheet.getRange(row, col_('Storage Bears')).setValue(body.StorageBears||0);
      sheet.getRange(row, col_('Resupply Tangle')).setValue(body.ResupplyTangle||0);
      sheet.getRange(row, col_('Resupply Parkway')).setValue(body.ResupplyParkway||0);
      sheet.getRange(row, col_('Resupply Bears')).setValue(body.ResupplyBears||0);
      sheet.getRange(row, col_('On Hold Tangle')).setValue(body.OnHoldTangle||0);
      sheet.getRange(row, col_('On Hold Parkway')).setValue(body.OnHoldParkway||0);
      sheet.getRange(row, col_('On Hold Bears')).setValue(body.OnHoldBears||0);
      // Order fields can be legitimately null ("never manually ordered") —
      // that must round-trip back to a blank cell, not 0, or Undo would
      // accidentally give the row an explicit order it never had.
      sheet.getRange(row, col_('Order Tangle')).setValue(body.OrderTangle===null||body.OrderTangle===undefined ? '' : body.OrderTangle);
      sheet.getRange(row, col_('Order Parkway')).setValue(body.OrderParkway===null||body.OrderParkway===undefined ? '' : body.OrderParkway);
      sheet.getRange(row, col_('Order Bears')).setValue(body.OrderBears===null||body.OrderBears===undefined ? '' : body.OrderBears);
      sheet.getRange(row, col_('Checked Tangle')).setValue(body.CheckedTangle||0);
      sheet.getRange(row, col_('Checked Parkway')).setValue(body.CheckedParkway||0);
      sheet.getRange(row, col_('Checked Bears')).setValue(body.CheckedBears||0);

    } else {
      result = { ok:false, error:'Unknown action: ' + action };
    }

    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ ok:false, error:String(err) });
  }
}
