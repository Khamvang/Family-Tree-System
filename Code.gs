/**
 * Family Tree Management System — Backend (Google Apps Script)
 *
 * Sheet used: "Members" (auto-created on first run if missing)
 * Columns: ID | FirstName | LastName | Gender | BirthDate | DeathDate |
 *          PhotoURL | FatherID | MotherID | SpouseID | ChildrenCount | Gen | Notes | DateAdded
 */

const SHEET_NAME = 'Members';
const HEADERS = ['ID',
                 'LaoFirstName', 'LaoLastName',
                 'EnglishFirstName', 'EnglishLastName',
                 'HmongFirstName', 'HmongLastName',
                 'Gender', 'Status', 'BirthDate', 'DeathDate',
                 'PhotoURL', 'FatherID', 'MotherID', 'SpouseID', 'ChildrenCount', 'Gen', 'Notes',
                 'Country', 'Province', 'City', 'Village', 'HouseLocationURL', 'HouseImages',
                 'DateAdded', 'DateUpdated',
                 // New columns appended at the end (rather than interleaved) so that any
                 // already-deployed sheet's existing physical column order stays intact —
                 // ensureMembersHeaders_() appends these same new columns to the end of the
                 // real sheet the first time it runs, keeping HEADERS order === sheet order.
                 'BirthCountry', 'BirthProvince', 'BirthCity', 'BirthVillage', 'BirthRemark', 'BirthAttachments',
                 'DeathCountry', 'DeathProvince', 'DeathCity', 'DeathVillage', 'DeathRemark', 'DeathAttachments'];

// --- Sub-sheets for repeating record types (Address / Education / Work history) ---
const ADDRESS_HISTORY_SHEET = 'AddressHistory';
const ADDRESS_HISTORY_HEADERS = ['RecordID', 'MemberID', 'FromDate', 'ToDate',
                                  'Country', 'Province', 'City', 'Village', 'Remark', 'Attachments',
                                  'DateAdded', 'DateUpdated'];

const EDUCATION_SHEET = 'EducationHistory';
const EDUCATION_HEADERS = ['RecordID', 'MemberID', 'FromDate', 'ToDate',
                            'SchoolLevel', 'SchoolName', 'Major', 'GPA', 'Remark', 'Attachments',
                            'DateAdded', 'DateUpdated'];

const WORK_SHEET = 'WorkExperience';
const WORK_HEADERS = ['RecordID', 'MemberID', 'FromDate', 'ToDate',
                       'CompanyName', 'Position', 'Remark', 'Attachments',
                       'DateAdded', 'DateUpdated'];

// Whitelist of sub-sheets that generic history-record functions are allowed to touch.
const HISTORY_SHEETS_ = {
  AddressHistory: ADDRESS_HISTORY_HEADERS,
  EducationHistory: EDUCATION_HEADERS,
  WorkExperience: WORK_HEADERS
};

/**
 * Web app entry point. The whole UI (sidebar + both panels) lives in Index.html now,
 * so there's no more page= routing — this always serves that one page.
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Hmoob Keeb Kwm')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Lets templates pull in shared CSS/JS via <?!= include('Styles'); ?> */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Gets (or creates) the Members sheet with headers already set up. */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    // In-place renames (preserve the column's position and existing data,
    // unlike ensureMembersHeaders_ below which only ever *appends* new
    // columns) — safe to run on every load, since each is a no-op once the
    // header has already been renamed on a previous run.
    renameHeaderIfPresent_(sheet, 'PajLevel', 'Gen');
    const childrenCountJustRenamed = renameHeaderIfPresent_(sheet, 'Generation', 'ChildrenCount');
    ensureMembersHeaders_(sheet);
    if (childrenCountJustRenamed) recalcAllChildrenCounts_(sheet);
  }
  return sheet;
}

/**
 * Renames an existing column header in place (its data underneath is
 * untouched, since only the row-1 header cell changes) — used when a field's
 * purpose or display name changes, so old data isn't orphaned under an
 * abandoned column. No-ops if oldName isn't present, or newName already is.
 */
function renameHeaderIfPresent_(sheet, oldName, newName) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return false;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf(oldName);
  if (idx > -1 && headers.indexOf(newName) === -1) {
    sheet.getRange(1, idx + 1).setValue(newName);
    return true;
  }
  return false;
}

/**
 * One-time-per-column migration: if this is a sheet that was already live before
 * the Birth/Death detail columns existed, this appends any headers from HEADERS
 * that aren't in the sheet yet — as new columns at the end, never touching or
 * reordering existing columns or data. Safe to run on every load; it's a no-op
 * once the columns already exist.
 */
function ensureMembersHeaders_(sheet) {
  const lastCol = sheet.getLastColumn();
  const existingHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const missing = HEADERS.filter(h => existingHeaders.indexOf(h) === -1);
  if (missing.length) {
    sheet.getRange(1, existingHeaders.length + 1, 1, missing.length).setValues([missing]);
  }
}

/** Generic "get or create a sub-sheet with these headers" — same pattern as getSheet_(), for the history sheets. */
function getOrCreateSubSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  } else {
    ensureSheetHeaders_(sheet, headers);
  }
  return sheet;
}

/** Generic version of ensureMembersHeaders_() — appends any headers missing from an existing sheet, without touching or reordering existing columns/data. */
function ensureSheetHeaders_(sheet, headers) {
  const lastCol = sheet.getLastColumn();
  const existingHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const missing = headers.filter(h => existingHeaders.indexOf(h) === -1);
  if (missing.length) {
    sheet.getRange(1, existingHeaders.length + 1, 1, missing.length).setValues([missing]);
  }
}

/** Returns every member as an array of plain objects (used by the tree view). */
function getFamilyData() {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row[0]) // skip blank rows
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        // Sheets returns date/time cells as real Date objects, but Date objects
        // don't reliably survive the trip back to the browser via google.script.run
        // (they can silently turn the whole response into null). Converting to a
        // plain string here means nothing downstream ever touches a raw Date.
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone() || 'Etc/UTC', "yyyy-MM-dd'T'HH:mm:ss");
        }
        obj[h] = val;
      });
      return obj;
    });
}

/** Primary label shown in lists/dropdowns/tree: prefer English, then Lao, then Hmong. */
function displayName_(m) {
  const english = `${m.EnglishFirstName || ''} ${m.EnglishLastName || ''}`.trim();
  const lao = `${m.LaoFirstName || ''} ${m.LaoLastName || ''}`.trim();
  const hmong = `${m.HmongFirstName || ''} ${m.HmongLastName || ''}`.trim();
  return english || lao || hmong || '(unnamed)';
}

/** Lowercased blob of every name field (all languages), used for matching search queries. */
function searchableName_(m) {
  return [m.EnglishFirstName, m.EnglishLastName, m.LaoFirstName, m.LaoLastName, m.HmongFirstName, m.HmongLastName]
    .filter(Boolean).join(' ').toLowerCase();
}

/** Resolves a member's father, mother, spouse(s), and children into simple {id, name, birthYear, gender} objects — used to compile the printable CV. */
function getFamilySummary(memberId) {
  const all = getFamilyData();
  const byId = {};
  all.forEach(m => byId[m.ID] = m);
  const target = String(memberId).trim();
  const person = all.find(m => String(m.ID).trim() === target);
  if (!person) return null;

  function toSummary(m) {
    if (!m) return null;
    return {
      id: m.ID,
      name: displayName_(m),
      birthYear: m.BirthDate ? new Date(m.BirthDate).getFullYear() : '',
      gender: m.Gender || ''
    };
  }

  const father = person.FatherID ? byId[String(person.FatherID).trim()] : null;
  const mother = person.MotherID ? byId[String(person.MotherID).trim()] : null;
  const spouseIds = String(person.SpouseID || '').split(',').map(s => s.trim()).filter(Boolean);
  const spouses = spouseIds.map(id => byId[id]).filter(Boolean).map(toSummary);
  const children = all
    .filter(m => String(m.FatherID).trim() === target || String(m.MotherID).trim() === target)
    .map(toSummary);

  return {
    father: toSummary(father),
    mother: toSummary(mother),
    spouses: spouses,
    children: children
  };
}

/**
 * Everything the printable CV needs, fetched in a single server round-trip
 * instead of five separate ones (member + family used to each call
 * getFamilyData() independently, reading the whole Members sheet twice for
 * no reason; the three history sheets are also bundled into this one response).
 */
function getMemberCVData(memberId) {
  const target = String(memberId).trim();
  const all = getFamilyData(); // one read of the Members sheet, reused for both the member and their family
  const byId = {};
  all.forEach(m => byId[m.ID] = m);
  const member = byId[target] || null;
  if (!member) return null;

  function toSummary(m) {
    if (!m) return null;
    return {
      id: m.ID,
      name: displayName_(m),
      birthYear: m.BirthDate ? new Date(m.BirthDate).getFullYear() : '',
      gender: m.Gender || ''
    };
  }

  const father = member.FatherID ? byId[String(member.FatherID).trim()] : null;
  const mother = member.MotherID ? byId[String(member.MotherID).trim()] : null;
  const spouseIds = String(member.SpouseID || '').split(',').map(s => s.trim()).filter(Boolean);
  const spouses = spouseIds.map(id => byId[id]).filter(Boolean).map(toSummary);
  const children = all
    .filter(m => String(m.FatherID).trim() === target || String(m.MotherID).trim() === target)
    .map(toSummary);

  return {
    member: member,
    family: {
      father: toSummary(father),
      mother: toSummary(mother),
      spouses: spouses,
      children: children
    },
    addresses: getRecordsForMember_(ADDRESS_HISTORY_SHEET, ADDRESS_HISTORY_HEADERS, target),
    education: getRecordsForMember_(EDUCATION_SHEET, EDUCATION_HEADERS, target),
    work: getRecordsForMember_(WORK_SHEET, WORK_HEADERS, target)
  };
}

/** Lightweight lookup for autocomplete fields: [{id, names, birthYear, father, mother, ...}] */
function searchMembers(query) {
  const q = (query || '').toString().toLowerCase().trim();
  const all = getFamilyData();
  const byId = {};
  all.forEach(m => byId[m.ID] = m);

  const matches = q ? all.filter(m => searchableName_(m).includes(q)) : all;

  const nameSet_ = (m) => ({
    en: `${m.EnglishFirstName || ''} ${m.EnglishLastName || ''}`.trim(),
    lo: `${m.LaoFirstName || ''} ${m.LaoLastName || ''}`.trim(),
    hmn: `${m.HmongFirstName || ''} ${m.HmongLastName || ''}`.trim()
  });

  return matches.slice(0, 25).map(m => {
    const father = byId[m.FatherID];
    const mother = byId[m.MotherID];
    return {
      id: m.ID,
      name: displayName_(m), // kept for any older caller that just wants a single fallback string
      names: nameSet_(m),
      fatherNames: father ? nameSet_(father) : null,
      motherNames: mother ? nameSet_(mother) : null,
      birthYear: m.BirthDate ? new Date(m.BirthDate).getFullYear() : '',
      country: m.Country || '',
      province: m.Province || '',
      city: m.City || '',
      village: m.Village || ''
    };
  });
}

/** Fetches one member by ID (used to prefill the form in edit mode). */
function getMemberById(id) {
  const target = String(id).trim();
  const all = getFamilyData();
  return all.find(m => String(m.ID).trim() === target) || null;
}

/** Gets (or creates) a subfolder with the given name inside a parent folder. */
function getOrCreateSubfolder_(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

/**
 * Gets (or creates) the per-member image folder:
 * <folder containing this Sheet>/Family Tree - House Images/<ID> - <name>/
 */
function getOrCreateMemberImageFolder_(memberId, memberLabel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  const rootFolder = getOrCreateSubfolder_(parentFolder, 'Family Tree - House Images');
  const safeName = (memberId + ' - ' + (memberLabel || 'Unnamed')).replace(/[\\/:*?"<>|]/g, '-');
  return getOrCreateSubfolder_(rootFolder, safeName);
}

/**
 * Uploads one file to that member's Drive folder, sets it viewable by link,
 * and immediately writes the resulting URL into the given field on that
 * member's row (so it's saved right away, not just held in the form).
 *   - fieldName 'PhotoURL' replaces any existing value (one profile photo).
 *   - fieldName 'HouseImages' appends onto the existing list (multiple files).
 * fileData: base64-encoded file content (no "data:...;base64," prefix).
 */
function uploadMemberFile(memberId, memberLabel, fileName, mimeType, fileData, fieldName) {
  if (!memberId) {
    return { success: false, error: 'This person needs to be saved first before files can be attached.' };
  }

  const folder = getOrCreateMemberImageFolder_(memberId, memberLabel);
  const bytes = Utilities.base64Decode(fileData);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';

  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('ID');
  const fieldCol = headers.indexOf(fieldName);
  const target = String(memberId).trim();

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]).trim() === target) {
      let updated;
      if (fieldName === 'PhotoURL') {
        updated = url; // single profile photo — replaces whatever was there
      } else {
        const existing = values[r][fieldCol] || '';
        updated = existing ? existing + '\n' + url : url;
      }
      sheet.getRange(r + 1, fieldCol + 1).setValue(updated);
      return { success: true, url: url, allUrls: updated };
    }
  }
  return { success: false, url: url, error: 'Uploaded to Drive, but could not find this member\'s row to attach it to.' };
}

/**
 * Removes a file reference from a member's field (PhotoURL clears entirely;
 * HouseImages removes just the matching line) and trashes the actual Drive
 * file so it doesn't linger unused.
 */
function removeMemberFile(memberId, fieldName, url) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('ID');
  const fieldCol = headers.indexOf(fieldName);
  const target = String(memberId).trim();

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]).trim() === target) {
      let updated;
      if (fieldName === 'PhotoURL') {
        updated = '';
      } else {
        const existing = String(values[r][fieldCol] || '');
        updated = existing.split('\n').map(s => s.trim()).filter(line => line && line !== url).join('\n');
      }
      sheet.getRange(r + 1, fieldCol + 1).setValue(updated);

      try {
        const match = url && (url.match(/\/d\/([a-zA-Z0-9_-]{10,})/) || url.match(/[?&]id=([a-zA-Z0-9_-]{10,})/));
        if (match) DriveApp.getFileById(match[1]).setTrashed(true);
      } catch (e) {
        // Ignore Drive cleanup failures — removing the sheet reference is what matters most.
      }

      return { success: true, allUrls: updated };
    }
  }
  return { success: false, error: 'Could not find this member\'s row.' };
}

/**
 * Reads the "Laos_Address_Data" sheet (columns: province, city, village — any
 * extra columns like pvd_id are ignored) for the cascading address dropdowns.
 * Returns [] if that sheet doesn't exist yet.
 */
function getLaosAddressData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Laos_Address_Data');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const provinceCol = headers.indexOf('province');
  const cityCol = headers.indexOf('city');
  const villageCol = headers.indexOf('village');
  if (provinceCol === -1 || cityCol === -1 || villageCol === -1) return [];

  return data.slice(1)
    .filter(row => row[provinceCol])
    .map(row => ({
      province: String(row[provinceCol] || '').trim(),
      city: String(row[cityCol] || '').trim(),
      village: String(row[villageCol] || '').trim()
    }));
}

/** Generates the next sequential ID, e.g. M0001, M0002 ... */
function generateId_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'M0001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
  const nums = ids.map(id => parseInt(String(id).replace(/\D/g, ''), 10)).filter(n => !isNaN(n));
  const max = nums.length ? Math.max(...nums) : 0;
  return 'M' + String(max + 1).padStart(4, '0');
}

/**
 * ChildrenCount = how many members currently list this person as their
 * Father or Mother. Recomputed and written for one member at a time,
 * triggered whenever someone's Father/Mother links change (see saveMember) —
 * not calculated at the member's own save time, since a person's children
 * count depends on other rows, not their own.
 */
function recalcChildrenCount_(memberId, cached) {
  if (!memberId) return;
  const sheet = (cached && cached.sheet) || getSheet_();
  const values = (cached && cached.values) || sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('ID');
  const fatherCol = headers.indexOf('FatherID');
  const motherCol = headers.indexOf('MotherID');
  const countCol = headers.indexOf('ChildrenCount');
  if (idCol === -1 || countCol === -1) return;

  const target = String(memberId).trim();
  let count = 0;
  let targetRow = -1;
  for (let r = 1; r < values.length; r++) {
    const fid = String(values[r][fatherCol] || '').trim();
    const mid = String(values[r][motherCol] || '').trim();
    if (fid === target || mid === target) count++;
    if (String(values[r][idCol]).trim() === target) targetRow = r;
  }
  if (targetRow === -1) return;
  sheet.getRange(targetRow + 1, countCol + 1).setValue(count);
  values[targetRow][countCol] = count; // keep the in-memory snapshot in sync too, for any caller reusing it after this
}

/** One-time bulk version of recalcChildrenCount_ for every row at once — used right after the Generation→ChildrenCount migration, since the old column's stored values (generation-distance numbers) are meaningless for this new purpose. */
function recalcAllChildrenCounts_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('ID');
  const fatherCol = headers.indexOf('FatherID');
  const motherCol = headers.indexOf('MotherID');
  const countCol = headers.indexOf('ChildrenCount');
  if (idCol === -1 || countCol === -1) return;

  const countById = {};
  for (let r = 1; r < values.length; r++) {
    const fid = String(values[r][fatherCol] || '').trim();
    const mid = String(values[r][motherCol] || '').trim();
    if (fid) countById[fid] = (countById[fid] || 0) + 1;
    if (mid) countById[mid] = (countById[mid] || 0) + 1;
  }
  const countColumn = [];
  for (let r = 1; r < values.length; r++) {
    const id = String(values[r][idCol] || '').trim();
    countColumn.push([id ? (countById[id] || 0) : values[r][countCol]]);
  }
  if (countColumn.length) {
    sheet.getRange(2, countCol + 1, countColumn.length, 1).setValues(countColumn);
  }
}

/**
 * Creates a new member, or updates an existing one if data.ID matches a row.
 * data: {ID?, FirstName, LastName, Gender, BirthDate, DeathDate, PhotoURL,
 *        FatherID, MotherID, SpouseID, Notes}
 *
 * IMPORTANT: this always writes according to the sheet's *actual* current
 * column order (read fresh from row 1), not the order fields happen to be
 * listed in the HEADERS constant. That's what makes it safe for HEADERS to
 * gain new fields anywhere (not just appended at the end) across updates —
 * ensureMembersHeaders_() only needs to guarantee every column *exists*
 * somewhere, never that it's in any particular position.
 */
function saveMember(data) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0]; // the sheet's real, current column order — authoritative

  // These are no longer submitted by the form directly — they're maintained by
  // recalcCurrentAddress_() whenever an Address History record is saved, or by
  // recalcChildrenCount_()/recalcAllChildrenCounts_() whenever a child's own
  // Father/Mother links change — so a normal member save must leave whatever
  // is already in these cells untouched.
  const SERVER_MANAGED_FIELDS_ = ['Country', 'Province', 'City', 'Village', 'ChildrenCount'];

  if (data.ID) {
    const target = String(data.ID).trim();
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][0]).trim() === target) {
        const rowNum = r + 1;
        const fatherCol = headers.indexOf('FatherID');
        const motherCol = headers.indexOf('MotherID');
        const oldFatherId = fatherCol > -1 ? String(values[r][fatherCol] || '').trim() : '';
        const oldMotherId = motherCol > -1 ? String(values[r][motherCol] || '').trim() : '';

        const rowValues = headers.map((h, i) => {
          if (h === 'ID') return data.ID;
          if (h === 'DateAdded') return values[r][i]; // preserve original
          if (h === 'DateUpdated') return new Date(); // always refreshed on save
          if (SERVER_MANAGED_FIELDS_.indexOf(h) > -1) return values[r][i]; // preserve
          return data[h] !== undefined ? data[h] : '';
        });
        sheet.getRange(rowNum, 1, 1, headers.length).setValues([rowValues]);
        values[r] = rowValues; // keep the in-memory snapshot in sync with what was just written
        const cached = { sheet: sheet, values: values };
        if (data.Gen) cascadeGen_(data.ID, cached);

        // Recompute children-count for every parent involved — the new ones
        // (who just gained a child) and, if a parent link changed, the old
        // ones too (who just lost one).
        const newFatherId = String(data.FatherID || '').trim();
        const newMotherId = String(data.MotherID || '').trim();
        const parentsToRecalc = new Set([oldFatherId, oldMotherId, newFatherId, newMotherId].filter(Boolean));
        parentsToRecalc.forEach(pid => recalcChildrenCount_(pid, cached));

        return { success: true, id: data.ID, mode: 'updated' };
      }
    }
  }

  const newId = generateId_();
  const rowValues = headers.map(h => {
    if (h === 'ID') return newId;
    if (h === 'DateAdded') return new Date();
    if (h === 'DateUpdated') return new Date();
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(rowValues);
  values.push(rowValues); // keep the in-memory snapshot in sync with the row just appended
  // A brand-new member can't have any children pointing at them yet, so no
  // Gen cascade is needed — but they might already list a Father/Mother
  // (their own children count doesn't change; their PARENTS' does).
  const cached = { sheet: sheet, values: values };
  [data.FatherID, data.MotherID].filter(Boolean).forEach(pid => recalcChildrenCount_(pid, cached));
  return { success: true, id: newId, mode: 'created' };
}

/**
 * Recomputes the Paj (respect-level) number for every relative connected to a
 * member, in every direction, whenever that member's own level changes:
 *   - down to children, grandchildren, etc.        (each generation: +1)
 *   - up to parents, grandparents, etc.             (each generation: -1)
 *   - across to siblings (via a shared parent)      (same level as them)
 *   - across to spouse(s)                           (same level as them)
 * This is what makes fixing one person's level automatically ripple through
 * their *entire* connected family instead of needing every relative corrected
 * by hand. Always overrides whatever level a relative had before, since level
 * is purely a function of relationship distance from whichever person was
 * just corrected. If two different paths through the family disagree (e.g.
 * inconsistent data), whichever one this traversal reaches a person by first
 * wins — later, conflicting paths are left alone rather than fought over.
 */
function cascadeGen_(rootMemberId, cached) {
  const sheet = (cached && cached.sheet) || getSheet_();
  const values = (cached && cached.values) || sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('ID');
  const fatherCol = headers.indexOf('FatherID');
  const motherCol = headers.indexOf('MotherID');
  const spouseCol = headers.indexOf('SpouseID');
  const genCol = headers.indexOf('Gen');
  if (idCol === -1 || genCol === -1) return;

  const rowById = {};
  for (let r = 1; r < values.length; r++) {
    const id = String(values[r][idCol] || '').trim();
    if (id) rowById[id] = r;
  }

  // Build a weighted adjacency list covering every relationship type at once:
  //   parent → child : +1     child → parent : -1     spouse ↔ spouse : 0
  const neighbors = {};
  function addEdge(a, b, delta) {
    if (!a || !b || a === b) return;
    (neighbors[a] = neighbors[a] || []).push({ id: b, delta: delta });
  }
  for (let r = 1; r < values.length; r++) {
    const id = String(values[r][idCol] || '').trim();
    if (!id) continue;
    const fid = fatherCol > -1 ? String(values[r][fatherCol] || '').trim() : '';
    const mid = motherCol > -1 ? String(values[r][motherCol] || '').trim() : '';
    if (fid) { addEdge(id, fid, -1); addEdge(fid, id, 1); }
    if (mid) { addEdge(id, mid, -1); addEdge(mid, id, 1); }
    if (spouseCol > -1) {
      String(values[r][spouseCol] || '').split(',').map(s => s.trim()).filter(Boolean).forEach(sid => {
        addEdge(id, sid, 0);
        addEdge(sid, id, 0);
      });
    }
  }

  const target = String(rootMemberId).trim();
  if (rowById[target] === undefined) return;
  const rootLevel = Number(values[rowById[target]][genCol]) || 0;
  if (!rootLevel) return;

  // Breadth-first outward from the corrected person, in every direction at once.
  const levels = {};
  levels[target] = rootLevel;
  const queue = [target];
  while (queue.length) {
    const cur = queue.shift();
    const curLevel = levels[cur];
    (neighbors[cur] || []).forEach(edge => {
      if (levels[edge.id] !== undefined) return; // already resolved via a shorter path
      const newLevel = curLevel + edge.delta;
      if (newLevel < 1) return; // never assign a non-positive level to an ancestor beyond the top
      levels[edge.id] = newLevel;
      queue.push(edge.id);
    });
  }

  // Write back only what actually changed — batched into a single range write
  // instead of one setValue() call per row, since correcting one person's
  // level can ripple through dozens of relatives at once, and each
  // individual Sheets API call has its own fixed latency overhead.
  let anyChanged = false;
  const genColumn = values.slice(1).map(row => [row[genCol]]);
  Object.keys(levels).forEach(id => {
    if (id === target) return;
    const r = rowById[id];
    if (r === undefined) return;
    const current = Number(values[r][genCol]) || 0;
    if (current !== levels[id]) {
      genColumn[r - 1] = [levels[id]];
      values[r][genCol] = levels[id]; // keep the in-memory snapshot in sync too, for any caller reusing it after this
      anyChanged = true;
    }
  });
  if (anyChanged) {
    sheet.getRange(2, genCol + 1, genColumn.length, 1).setValues(genColumn);
  }
}

/** Converts a raw sheet row into a plain object keyed by header, normalizing Date cells to strings (see getFamilyData for why). */
function rowToObj_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    let val = row[i];
    if (val instanceof Date) {
      val = Utilities.formatDate(val, Session.getScriptTimeZone() || 'Etc/UTC', "yyyy-MM-dd'T'HH:mm:ss");
    }
    obj[h] = val;
  });
  return obj;
}

/** Generates the next sequential RecordID for a history sub-sheet, e.g. A0001, E0002, W0003. */
function generateRecordId_(sheet, prefix) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return prefix + '0001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
  const nums = ids.map(id => parseInt(String(id).replace(/\D/g, ''), 10)).filter(n => !isNaN(n));
  const max = nums.length ? Math.max(...nums) : 0;
  return prefix + String(max + 1).padStart(4, '0');
}

/** All records in a history sub-sheet belonging to one member. */
function getRecordsForMember_(sheetName, headers, memberId) {
  const sheet = getOrCreateSubSheet_(sheetName, headers);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const hdrs = data[0];
  const midCol = hdrs.indexOf('MemberID');
  const target = String(memberId).trim();
  return data.slice(1)
    .filter(row => row[0] && String(row[midCol]).trim() === target)
    .map(row => rowToObj_(hdrs, row));
}

/**
 * Creates a new history record, or updates an existing one if data.RecordID
 * matches a row. Like saveMember(), this writes according to the sheet's
 * actual current column order (read fresh from row 1), not the order fields
 * happen to be listed in the passed-in headers constant — safe for future
 * field additions anywhere in that constant, not just at the end.
 */
function saveRecord_(sheetName, headersConst, prefix, data, idField, preserveFields) {
  idField = idField || 'RecordID';
  preserveFields = preserveFields || [];
  const sheet = getOrCreateSubSheet_(sheetName, headersConst);
  const values = sheet.getDataRange().getValues();
  const hdrs = values[0]; // the sheet's real, current column order — authoritative
  const userEmail = currentUserEmail_();

  if (data[idField]) {
    const target = String(data[idField]).trim();
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][0]).trim() === target) {
        const rowValues = hdrs.map((h, i) => {
          if (h === idField) return data[idField];
          if (h === 'DateAdded') return values[r][i]; // preserve original
          if (h === 'DateUpdated') return new Date();
          if (h === 'CreatedBy') return values[r][i]; // preserve the original creator, never overwritten by later edits
          if (h === 'UpdatedBy') return userEmail || values[r][i]; // record it if known this time; otherwise leave whatever was already there
          if (preserveFields.indexOf(h) > -1) return values[r][i]; // server-managed — never overwrite from the client payload
          return data[h] !== undefined ? data[h] : '';
        });
        sheet.getRange(r + 1, 1, 1, hdrs.length).setValues([rowValues]);
        return { success: true, id: data[idField], mode: 'updated' };
      }
    }
  }

  const newId = generateRecordId_(sheet, prefix);
  const rowValues = hdrs.map(h => {
    if (h === idField) return newId;
    if (h === 'DateAdded') return new Date();
    if (h === 'DateUpdated') return new Date();
    if (h === 'CreatedBy') return userEmail;
    if (h === 'UpdatedBy') return userEmail;
    if (preserveFields.indexOf(h) > -1) return ''; // brand-new row — nothing to preserve yet
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(rowValues);
  return { success: true, id: newId, mode: 'created' };
}

/**
 * The current visitor's email, only when it's actually determinable — this
 * depends on the web app's "Execute as" deployment setting and whether the
 * visitor is signed into a Google account the script can identify. Returns
 * '' rather than throwing when it can't be determined (e.g. anonymous
 * access), so callers can simply skip recording it instead of requiring it.
 */
function currentUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

/** Deletes one history record row by RecordID. */
function deleteRecord_(sheetName, headers, recordId) {
  const sheet = getOrCreateSubSheet_(sheetName, headers);
  const values = sheet.getDataRange().getValues();
  const target = String(recordId).trim();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === target) {
      sheet.deleteRow(r + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Could not find that record.' };
}

// ---------- Address History ----------

function getAddressHistory(memberId) {
  return getRecordsForMember_(ADDRESS_HISTORY_SHEET, ADDRESS_HISTORY_HEADERS, memberId);
}

/** Saves an address-history record, then recomputes the member's cached "current address" shown everywhere else in the app. */
function saveAddressRecord(data) {
  const res = saveRecord_(ADDRESS_HISTORY_SHEET, ADDRESS_HISTORY_HEADERS, 'A', data);
  recalcCurrentAddress_(data.MemberID);
  return res;
}

function deleteAddressRecord(recordId, memberId) {
  const res = deleteRecord_(ADDRESS_HISTORY_SHEET, ADDRESS_HISTORY_HEADERS, recordId);
  recalcCurrentAddress_(memberId);
  return res;
}

/**
 * Recomputes Members.Country/Province/City/Village (the "current address" shown
 * in search results, the tree, etc.) from that member's Address History:
 *   - Prefer whichever record has no ToDate (still living there); if several,
 *     the one with the latest FromDate wins.
 *   - If every record has a ToDate (no longer anywhere current on file),
 *     fall back to the one with the latest ToDate.
 */
function recalcCurrentAddress_(memberId) {
  if (!memberId) return;
  const records = getRecordsForMember_(ADDRESS_HISTORY_SHEET, ADDRESS_HISTORY_HEADERS, memberId);
  if (!records.length) return;

  const ongoing = records.filter(r => !r.ToDate);
  let best;
  if (ongoing.length) {
    best = ongoing.reduce((a, b) => (String(b.FromDate || '') > String(a.FromDate || '') ? b : a));
  } else {
    best = records.reduce((a, b) => (String(b.ToDate || '') > String(a.ToDate || '') ? b : a));
  }

  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('ID');
  const target = String(memberId).trim();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]).trim() === target) {
      const rowValues = values[r].slice();
      ['Country', 'Province', 'City', 'Village'].forEach(f => {
        const col = headers.indexOf(f);
        if (col > -1) rowValues[col] = best[f] || '';
      });
      sheet.getRange(r + 1, 1, 1, headers.length).setValues([rowValues]);
      return;
    }
  }
}

// ---------- Education History ----------

function getEducationHistory(memberId) {
  return getRecordsForMember_(EDUCATION_SHEET, EDUCATION_HEADERS, memberId);
}

function saveEducationRecord(data) {
  return saveRecord_(EDUCATION_SHEET, EDUCATION_HEADERS, 'E', data);
}

function deleteEducationRecord(recordId) {
  return deleteRecord_(EDUCATION_SHEET, EDUCATION_HEADERS, recordId);
}

/** Distinct school names already on file, for the "search or add as new" autocomplete — no separate reference sheet needed. */
/** Both distinct School Names and Majors already on file, in a single sheet read instead of two separate ones. */
function getDistinctSchoolNamesAndMajors() {
  const sheet = getOrCreateSubSheet_(EDUCATION_SHEET, EDUCATION_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { schoolNames: [], majors: [] };
  const headers = data[0];
  const schoolCol = headers.indexOf('SchoolName');
  const majorCol = headers.indexOf('Major');
  const schoolNames = new Set();
  const majors = new Set();
  data.slice(1).forEach(r => {
    const sn = String(r[schoolCol] || '').trim();
    const mj = String(r[majorCol] || '').trim();
    if (sn) schoolNames.add(sn);
    if (mj) majors.add(mj);
  });
  return {
    schoolNames: Array.from(schoolNames).sort(),
    majors: Array.from(majors).sort()
  };
}

function getDistinctSchoolNames() {
  const sheet = getOrCreateSubSheet_(EDUCATION_SHEET, EDUCATION_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const col = data[0].indexOf('SchoolName');
  return Array.from(new Set(data.slice(1).map(r => String(r[col] || '').trim()).filter(Boolean))).sort();
}

/** Distinct majors already on file, same idea as getDistinctSchoolNames(). */
function getDistinctMajors() {
  const sheet = getOrCreateSubSheet_(EDUCATION_SHEET, EDUCATION_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const col = data[0].indexOf('Major');
  return Array.from(new Set(data.slice(1).map(r => String(r[col] || '').trim()).filter(Boolean))).sort();
}

// ---------- Work Experience ----------

function getWorkHistory(memberId) {
  return getRecordsForMember_(WORK_SHEET, WORK_HEADERS, memberId);
}

function saveWorkRecord(data) {
  return saveRecord_(WORK_SHEET, WORK_HEADERS, 'W', data);
}

function deleteWorkRecord(recordId) {
  return deleteRecord_(WORK_SHEET, WORK_HEADERS, recordId);
}

// ---------- Attachments for history records (Address / Education / Work) ----------
// Birth/Death attachments don't need new functions — they're plain fields on the
// Members row, so the existing uploadMemberFile()/removeMemberFile() (which already
// take an arbitrary fieldName) work for 'BirthAttachments'/'DeathAttachments' as-is.

/** Uploads one file into a history record's Drive folder (same per-member folder as house images) and appends the URL onto that record's Attachments cell. */
function uploadHistoryFile(sheetName, recordId, memberId, memberLabel, fileName, mimeType, fileData) {
  const headers = HISTORY_SHEETS_[sheetName];
  if (!headers) return { success: false, error: 'Unknown record type: ' + sheetName };
  if (!recordId) return { success: false, error: 'This record needs to be saved first before files can be attached.' };

  const folder = getOrCreateMemberImageFolder_(memberId, memberLabel);
  const bytes = Utilities.base64Decode(fileData);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';

  const sheet = getOrCreateSubSheet_(sheetName, headers);
  const values = sheet.getDataRange().getValues();
  const hdrs = values[0];
  const attCol = hdrs.indexOf('Attachments');
  const target = String(recordId).trim();

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === target) {
      const existing = values[r][attCol] || '';
      const updated = existing ? existing + '\n' + url : url;
      sheet.getRange(r + 1, attCol + 1).setValue(updated);
      return { success: true, url: url, allUrls: updated };
    }
  }
  return { success: false, url: url, error: 'Uploaded to Drive, but could not find that record\'s row to attach it to.' };
}

/** Removes one attachment URL from a history record and trashes the Drive file, mirroring removeMemberFile(). */
function removeHistoryFile(sheetName, recordId, url) {
  const headers = HISTORY_SHEETS_[sheetName];
  if (!headers) return { success: false, error: 'Unknown record type: ' + sheetName };

  const sheet = getOrCreateSubSheet_(sheetName, headers);
  const values = sheet.getDataRange().getValues();
  const hdrs = values[0];
  const attCol = hdrs.indexOf('Attachments');
  const target = String(recordId).trim();

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === target) {
      const existing = String(values[r][attCol] || '');
      const updated = existing.split('\n').map(s => s.trim()).filter(line => line && line !== url).join('\n');
      sheet.getRange(r + 1, attCol + 1).setValue(updated);

      try {
        const match = url && (url.match(/\/d\/([a-zA-Z0-9_-]{10,})/) || url.match(/[?&]id=([a-zA-Z0-9_-]{10,})/));
        if (match) DriveApp.getFileById(match[1]).setTrashed(true);
      } catch (e) {
        // Ignore Drive cleanup failures — removing the sheet reference is what matters most.
      }

      return { success: true, allUrls: updated };
    }
  }
  return { success: false, error: 'Could not find that record.' };
}

// =========================================================================
// KNOWLEDGE BASE MODULE — Hmong cultural customs/rituals reference library.
// Entirely separate sheets and Drive folder from the family tree data above;
// nothing here reads or writes Members/AddressHistory/EducationHistory/
// WorkExperience. Items nest to unlimited depth via ParentItemID (same idea
// as FatherID/MotherID elsewhere — the client reconstructs the tree from a
// flat list rather than the sheet storing hierarchy directly).
// =========================================================================

const KB_CATEGORIES_SHEET = 'KB_Categories';
// 'IconURL' is appended at the END (not slotted in next to the names) for the same
// reason as the Members Birth/Death columns — ensureSheetHeaders_() only ever appends,
// so an already-deployed sheet's existing column order stays intact.
// 'IconURL' is the single record of a category's icon. A separate Drive-file-ID
// column was tried and dropped: the URL is built by uploadKBCategoryIcon() in a
// fixed format that always embeds the ID, so driveFileIdFromUrl_() recovers it
// whenever a file has to be trashed — a second column would have been derived
// data that could only ever drift out of sync with this one.
const KB_CATEGORIES_HEADERS = ['CategoryID', 'Name_En', 'Name_Lo', 'Name_Hmn', 'OrderIndex',
                                'DateAdded', 'DateUpdated', 'CreatedBy', 'UpdatedBy',
                                'IconURL'];

const KB_ITEMS_SHEET = 'KB_Items';
const KB_ITEMS_HEADERS = ['ItemID', 'CategoryID', 'ParentItemID',
                           'Title_En', 'Title_Lo', 'Title_Hmn',
                           'Description_En', 'Description_Lo', 'Description_Hmn',
                           'OrderIndex',
                           'DateAdded', 'DateUpdated', 'CreatedBy', 'UpdatedBy'];

const KB_MEDIA_SHEET = 'KB_Media';
const KB_MEDIA_HEADERS = ['MediaID', 'ItemID', 'MediaType', 'DriveFileId', 'FileName', 'UploadedDate',
                           'CreatedBy', 'UpdatedBy'];

/** The two starting categories, seeded once if the sheet is completely empty. Name_En/Name_Lo default to the same Hmong text — edit them via the Add/Edit UI (Phase 3) or directly in the sheet once real translations are available. */
function seedKBCategoriesIfEmpty_(sheet) {
  if (sheet.getLastRow() >= 2) return; // already has data — never overwrite
  [
    ['C0001', 'Kab Lij Kev Cai Hmoob', 'Kab Lij Kev Cai Hmoob', 'Kab Lij Kev Cai Hmoob', 1],
    ['C0002', 'Txuj Ci Khaw Koob', 'Txuj Ci Khaw Koob', 'Txuj Ci Khaw Koob', 2]
  ].forEach(row => sheet.appendRow(row));
}

/** Every KB category, sorted by OrderIndex. */
function getKBCategories() {
  const sheet = getOrCreateSubSheet_(KB_CATEGORIES_SHEET, KB_CATEGORIES_HEADERS);
  seedKBCategoriesIfEmpty_(sheet);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row[0])
    .map(row => rowToObj_(headers, row))
    .sort((a, b) => (Number(a.OrderIndex) || 0) - (Number(b.OrderIndex) || 0));
}

/**
 * 'IconURL' is listed as a preserved field because the icon is managed only by
 * uploadKBCategoryIcon()/removeKBCategoryIcon() — the category form never sends
 * it, and saveRecord_ blanks any column the payload omits, so without this an
 * ordinary name edit would silently wipe the category's icon.
 */
function saveKBCategory(data) {
  return saveRecord_(KB_CATEGORIES_SHEET, KB_CATEGORIES_HEADERS, 'C', data, 'CategoryID', ['IconURL']);
}

/** Shared Drive folder for every category icon: Knowledge Base Media/Category Icons/ */
function getOrCreateKBCategoryIconFolder_() {
  return getOrCreateSubfolder_(getOrCreateKBMediaFolder_(), 'Category Icons');
}

/** Pulls the Drive file ID out of either share-link format, or '' if it isn't one. */
function driveFileIdFromUrl_(url) {
  if (!url) return '';
  const s = String(url);
  const m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/) || s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  return m ? m[1] : '';
}

/** Moves a Drive file to the trash by ID, ignoring failures (a stale reference isn't worth blocking on). */
function trashDriveFileById_(fileId) {
  if (!fileId) return;
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {
    // Ignore Drive cleanup failures — clearing the sheet reference is what matters most.
  }
}

/** Same, for callers that only have the share URL. */
function trashDriveFileFromUrl_(url) {
  trashDriveFileById_(driveFileIdFromUrl_(url));
}

/**
 * Finds a category's row (1-based sheet row) plus its IconURL column index and
 * current value. Shared by the icon upload/remove pair and by deleteKBCategory's
 * cleanup. `currentFileId` is derived from the URL rather than stored separately —
 * uploadKBCategoryIcon builds that URL itself, so the ID is always recoverable
 * from it, and keeping one copy means the two can never drift apart.
 */
function findKBCategoryRow_(categoryId) {
  const sheet = getOrCreateSubSheet_(KB_CATEGORIES_SHEET, KB_CATEGORIES_HEADERS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const urlCol = headers.indexOf('IconURL');
  const target = String(categoryId).trim();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === target) {
      const currentUrl = urlCol > -1 ? String(values[r][urlCol] || '').trim() : '';
      return {
        sheet: sheet,
        row: r + 1,
        urlCol: urlCol,
        currentUrl: currentUrl,
        currentFileId: driveFileIdFromUrl_(currentUrl)
      };
    }
  }
  return null;
}

/**
 * Uploads a category's icon image (already square-cropped client-side) and
 * writes the resulting share link onto that category's row. Replacing an icon
 * trashes the previous file, so swapping icons repeatedly doesn't leave Drive
 * littered. fileData: base64-encoded content, no "data:...;base64," prefix.
 */
function uploadKBCategoryIcon(categoryId, categoryLabel, fileName, mimeType, fileData) {
  if (!categoryId) {
    return { success: false, error: 'This category needs to be saved first before an icon can be attached.' };
  }
  const target = findKBCategoryRow_(categoryId);
  if (!target) return { success: false, error: 'Could not find that category.' };
  if (target.urlCol === -1) return { success: false, error: 'The IconURL column is missing from the KB_Categories sheet.' };

  const folder = getOrCreateKBCategoryIconFolder_();
  const safeName = (categoryId + ' - ' + (categoryLabel || 'icon')).replace(/[\\/:*?"<>|]/g, '-');
  const bytes = Utilities.base64Decode(fileData);
  const blob = Utilities.newBlob(bytes, mimeType, safeName + '.jpg');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();
  const url = 'https://drive.google.com/file/d/' + fileId + '/view?usp=sharing';

  const previousFileId = target.currentFileId;
  target.sheet.getRange(target.row, target.urlCol + 1).setValue(url);
  if (previousFileId && previousFileId !== fileId) trashDriveFileById_(previousFileId);

  return { success: true, url: url };
}

/** Clears a category's icon and trashes the underlying Drive file. */
function removeKBCategoryIcon(categoryId) {
  const target = findKBCategoryRow_(categoryId);
  if (!target) return { success: false, error: 'Could not find that category.' };
  if (target.urlCol === -1) return { success: true }; // nothing to clear

  target.sheet.getRange(target.row, target.urlCol + 1).setValue('');
  trashDriveFileById_(target.currentFileId);
  return { success: true };
}

/**
 * Blocks deletion if the category still has items in it, rather than silently
 * orphaning them. The refusal carries a machine-readable `reason` plus the item
 * count so the client can phrase the warning in whichever language is currently
 * selected; `error` stays as an English fallback for any caller that doesn't.
 */
function deleteKBCategory(categoryId) {
  const items = getKBItemsByCategory(categoryId);
  if (items.length) {
    return {
      success: false,
      reason: 'has_items',
      count: items.length,
      error: 'This category still has ' + items.length + ' item(s) in it. Move or delete them first.'
    };
  }
  // Trash the icon before the row goes, otherwise its Drive file is orphaned
  // with nothing left pointing at it.
  const target = findKBCategoryRow_(categoryId);
  if (target) trashDriveFileById_(target.currentFileId);
  return deleteRecord_(KB_CATEGORIES_SHEET, KB_CATEGORIES_HEADERS, categoryId);
}

/** Every item in one category (flat list — the client builds the nested tree from ParentItemID), sorted by OrderIndex. */
function getKBItemsByCategory(categoryId) {
  const sheet = getOrCreateSubSheet_(KB_ITEMS_SHEET, KB_ITEMS_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const catCol = headers.indexOf('CategoryID');
  const target = String(categoryId).trim();
  return data.slice(1)
    .filter(row => row[0] && String(row[catCol]).trim() === target)
    .map(row => rowToObj_(headers, row))
    .sort((a, b) => (Number(a.OrderIndex) || 0) - (Number(b.OrderIndex) || 0));
}

function getKBItem(itemId) {
  const sheet = getOrCreateSubSheet_(KB_ITEMS_SHEET, KB_ITEMS_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  const target = String(itemId).trim();
  const row = data.slice(1).find(r => r[0] && String(r[0]).trim() === target);
  return row ? rowToObj_(headers, row) : null;
}

/** Direct children (one level) of an item — used by deleteKBItem's safety check. */
function getKBChildren_(itemId) {
  const sheet = getOrCreateSubSheet_(KB_ITEMS_SHEET, KB_ITEMS_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const parentCol = headers.indexOf('ParentItemID');
  const target = String(itemId).trim();
  return data.slice(1).filter(row => row[0] && String(row[parentCol] || '').trim() === target);
}

/**
 * Saves a KB item (topic). A minimal guard against the most obvious cycle
 * (an item listing itself as its own parent) — full ancestor-chain cycle
 * checking belongs in the Add/Edit UI (Phase 3), where re-parenting is
 * actually exposed to the user.
 */
function saveKBItem(data) {
  if (data.ItemID && data.ParentItemID && String(data.ItemID).trim() === String(data.ParentItemID).trim()) {
    return { success: false, error: 'An item cannot be its own parent.' };
  }
  return saveRecord_(KB_ITEMS_SHEET, KB_ITEMS_HEADERS, 'I', data, 'ItemID');
}

/** Blocks deletion if the item still has sub-items, and cleans up (trashes) all of its own media first. */
function deleteKBItem(itemId) {
  const children = getKBChildren_(itemId);
  if (children.length) {
    return { success: false, error: 'This item has ' + children.length + ' sub-item(s). Delete or move them first.' };
  }
  getKBMediaForItem(itemId).forEach(m => {
    try {
      if (m.DriveFileId) DriveApp.getFileById(m.DriveFileId).setTrashed(true);
    } catch (e) {
      // Ignore Drive cleanup failures — removing the sheet rows is what matters most.
    }
  });
  deleteAllKBMediaRowsForItem_(itemId);
  return deleteRecord_(KB_ITEMS_SHEET, KB_ITEMS_HEADERS, itemId);
}

/** Every media row attached to one item. */
function getKBMediaForItem(itemId) {
  const sheet = getOrCreateSubSheet_(KB_MEDIA_SHEET, KB_MEDIA_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const itemCol = headers.indexOf('ItemID');
  const target = String(itemId).trim();
  return data.slice(1)
    .filter(row => row[0] && String(row[itemCol]).trim() === target)
    .map(row => rowToObj_(headers, row));
}

/** Deletes every KB_Media row for one item (sheet rows only — Drive cleanup happens separately in deleteKBItem/deleteKBMedia). */
function deleteAllKBMediaRowsForItem_(itemId) {
  const sheet = getOrCreateSubSheet_(KB_MEDIA_SHEET, KB_MEDIA_HEADERS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const itemCol = headers.indexOf('ItemID');
  const target = String(itemId).trim();
  for (let r = values.length - 1; r >= 1; r--) { // bottom-up so row indices don't shift under us
    if (String(values[r][itemCol]).trim() === target) sheet.deleteRow(r + 1);
  }
}

/** Root Drive folder for all Knowledge Base media, alongside the spreadsheet itself (same convention as getOrCreateMemberImageFolder_). */
function getOrCreateKBMediaFolder_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  return getOrCreateSubfolder_(parentFolder, 'Knowledge Base Media');
}

/** Per-item Drive folder: Knowledge Base Media/<ItemID> - <title>/ */
function getOrCreateKBItemFolder_(itemId, itemLabel) {
  const root = getOrCreateKBMediaFolder_();
  const safeName = (itemId + ' - ' + (itemLabel || 'Untitled')).replace(/[\\/:*?"<>|]/g, '-');
  return getOrCreateSubfolder_(root, safeName);
}

/** Guesses the KB MediaType bucket (image/video/audio/pdf/docx) from a MIME type, for when the client doesn't supply one explicitly. */
function guessKBMediaType_(mimeType) {
  mimeType = mimeType || '';
  if (mimeType.indexOf('image/') === 0) return 'image';
  if (mimeType.indexOf('video/') === 0) return 'video';
  if (mimeType.indexOf('audio/') === 0) return 'audio';
  if (mimeType.indexOf('pdf') > -1) return 'pdf';
  return 'docx';
}

/**
 * Uploads one media file (image/video/audio/PDF/docx) to the item's Drive
 * folder and logs it as a new KB_Media row. Mirrors uploadHistoryFile()'s
 * shape exactly, for the same reasons (consistency, and so the client-side
 * dropzone code added in Phase 4 can reuse the same calling convention).
 */
function uploadKBMedia(itemId, itemLabel, fileName, mimeType, fileData, mediaType) {
  if (!itemId) {
    return { success: false, error: 'This item needs to be saved first before files can be attached.' };
  }
  const folder = getOrCreateKBItemFolder_(itemId, itemLabel);
  const bytes = Utilities.base64Decode(fileData);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const sheet = getOrCreateSubSheet_(KB_MEDIA_SHEET, KB_MEDIA_HEADERS);
  const newId = generateRecordId_(sheet, 'M');
  const userEmail = currentUserEmail_();
  const rowValues = KB_MEDIA_HEADERS.map(h => {
    if (h === 'MediaID') return newId;
    if (h === 'ItemID') return itemId;
    if (h === 'MediaType') return mediaType || guessKBMediaType_(mimeType);
    if (h === 'DriveFileId') return file.getId();
    if (h === 'FileName') return fileName;
    if (h === 'UploadedDate') return new Date();
    if (h === 'CreatedBy') return userEmail;
    if (h === 'UpdatedBy') return userEmail;
    return '';
  });
  sheet.appendRow(rowValues);
  return { success: true, mediaId: newId, driveFileId: file.getId(), fileName: fileName };
}

/** Removes one media row and trashes its underlying Drive file. */
function deleteKBMedia(mediaId) {
  const sheet = getOrCreateSubSheet_(KB_MEDIA_SHEET, KB_MEDIA_HEADERS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('MediaID');
  const fileCol = headers.indexOf('DriveFileId');
  const target = String(mediaId).trim();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]).trim() === target) {
      try {
        const fileId = values[r][fileCol];
        if (fileId) DriveApp.getFileById(fileId).setTrashed(true);
      } catch (e) {
        // Ignore Drive cleanup failures — removing the sheet row is what matters most.
      }
      sheet.deleteRow(r + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Could not find that file.' };
}