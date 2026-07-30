/**
 * Family Tree Management System — Backend (Google Apps Script)
 *
 * Sheet used: "Members" (auto-created on first run if missing)
 * Columns: ID | FirstName | LastName | Gender | BirthDate | DeathDate |
 *          PhotoURL | FatherID | MotherID | SpouseID | Generation | Notes | DateAdded
 */

const SHEET_NAME = 'Members';
const HEADERS = ['ID',
                 'LaoFirstName', 'LaoLastName',
                 'EnglishFirstName', 'EnglishLastName',
                 'HmongFirstName', 'HmongLastName',
                 'Gender', 'BirthDate', 'DeathDate',
                 'PhotoURL', 'FatherID', 'MotherID', 'SpouseID', 'Generation', 'Notes',
                 'Country', 'Province', 'City', 'Village', 'HouseLocationURL', 'HouseImages',
                 'DateAdded', 'DateUpdated'];

/**
 * Web app entry point. The whole UI (sidebar + both panels) lives in Index.html now,
 * so there's no more page= routing — this always serves that one page.
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Family Tree System')
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
  }
  return sheet;
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

/** Lightweight lookup for autocomplete fields: [{id, name, birthYear, fatherName, motherName}] */
function searchMembers(query) {
  const q = (query || '').toString().toLowerCase().trim();
  const all = getFamilyData();
  const byId = {};
  all.forEach(m => byId[m.ID] = m);

  const matches = q ? all.filter(m => searchableName_(m).includes(q)) : all;

  return matches.slice(0, 25).map(m => {
    const father = byId[m.FatherID];
    const mother = byId[m.MotherID];
    return {
      id: m.ID,
      name: displayName_(m),
      birthYear: m.BirthDate ? new Date(m.BirthDate).getFullYear() : '',
      fatherName: father ? displayName_(father) : '',
      motherName: mother ? displayName_(mother) : '',
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
 * Generation = 1 + max(parents' generation). Root ancestors (no parents on file) = 0.
 * This is all that's needed for cousins to appear correctly in the tree — cousins simply
 * fall out naturally once every child correctly points at FatherID/MotherID.
 */
function calculateGeneration_(fatherId, motherId) {
  const all = getFamilyData();
  const fatherTarget = String(fatherId || '').trim();
  const motherTarget = String(motherId || '').trim();
  const father = fatherTarget ? all.find(m => String(m.ID).trim() === fatherTarget) : null;
  const mother = motherTarget ? all.find(m => String(m.ID).trim() === motherTarget) : null;
  const gens = [father, mother].filter(Boolean).map(p => Number(p.Generation) || 0);
  return gens.length ? Math.max(...gens) + 1 : 0;
}

/**
 * Creates a new member, or updates an existing one if data.ID matches a row.
 * data: {ID?, FirstName, LastName, Gender, BirthDate, DeathDate, PhotoURL,
 *        FatherID, MotherID, SpouseID, Notes}
 */
function saveMember(data) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const generation = calculateGeneration_(data.FatherID, data.MotherID);

  if (data.ID) {
    const target = String(data.ID).trim();
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][0]).trim() === target) {
        const rowNum = r + 1;
        const rowValues = HEADERS.map(h => {
          if (h === 'ID') return data.ID;
          if (h === 'Generation') return generation;
          if (h === 'DateAdded') return values[r][headers.indexOf('DateAdded')]; // preserve original
          if (h === 'DateUpdated') return new Date(); // always refreshed on save
          return data[h] !== undefined ? data[h] : '';
        });
        sheet.getRange(rowNum, 1, 1, HEADERS.length).setValues([rowValues]);
        return { success: true, id: data.ID, mode: 'updated' };
      }
    }
  }

  const newId = generateId_();
  const rowValues = HEADERS.map(h => {
    if (h === 'ID') return newId;
    if (h === 'Generation') return generation;
    if (h === 'DateAdded') return new Date();
    if (h === 'DateUpdated') return new Date();
    return data[h] !== undefined ? data[h] : '';
  });
  sheet.appendRow(rowValues);
  return { success: true, id: newId, mode: 'created' };
}
