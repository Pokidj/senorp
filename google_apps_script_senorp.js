const SHEETS = {
  inventory: "Inventario",
  bolos: "Bolos",
  settings: "Configuracion",
};

function doGet() {
  return jsonResponse({
    apiVersion: 3,
    capabilities: ["inventory", "bolos", "technicians", "technician-specialties-v2", "riders"],
    inventory: readInventory(),
    bolos: readBolos(),
    tecnicos: readTechnicians(),
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (payload.action === "saveInventory") {
      writeInventory(payload.inventory || {});
      return jsonResponse({ ok: true });
    }

    if (payload.action === "saveBolo") {
      saveBolo(payload.bolo);
      return jsonResponse({ ok: true });
    }

    if (payload.action === "saveTechnicians") {
      writeTechnicians(payload.tecnicos || payload.technicians || []);
      return jsonResponse({ ok: true, tecnicos: readTechnicians() });
    }

    if (payload.action === "uploadRider") {
      return jsonResponse({ ok: true, rider: uploadRider(payload.rider, payload.boloId) });
    }

    if (payload.action === "deleteBolo") {
      deleteBolo(payload.id);
      return jsonResponse({ ok: true });
    }

    return jsonResponse({ ok: false, error: "Unknown action" });
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return jsonResponse({
      ok: false,
      error: error && error.message ? error.message : String(error || "Error interno")
    });
  } finally {
    lock.releaseLock();
  }
}

function readInventory() {
  const sheet = getSheet(SHEETS.inventory, ["categoria", "nombre", "modelo", "stock", "estado"]);
  const values = readSheetValues(sheet, 5);
  const rows = values.slice(1);
  const inventory = { sonido: [], iluminacion: [], otros: [] };

  rows.forEach(([categoria, nombre, modelo, stock, estado]) => {
    categoria = normalizeCategory(categoria);
    nombre = String(nombre || "").trim();
    modelo = String(modelo || "Sin modelo").trim();
    if (!nombre) return;

    upsertInventoryModel(inventory, categoria, nombre, {
      nombre: modelo,
      stock: Number(stock) || 0,
      estado: normalizeEstado(estado),
    });
  });

  return inventory;
}

function writeInventory(inventory) {
  const sheet = getSheet(SHEETS.inventory, ["categoria", "nombre", "modelo", "stock", "estado"]);
  inventory = normalizeInventory(inventory);
  const rows = [["categoria", "nombre", "modelo", "stock", "estado"]];

  ["sonido", "iluminacion", "otros"].forEach(categoria => {
    (inventory[categoria] || []).forEach(item => {
      (item.modelos || []).forEach(modelo => {
        rows.push([
          categoria,
          item.nombre || "",
          modelo.nombre || "Sin modelo",
          Number(modelo.stock) || 0,
          normalizeEstado(modelo.estado),
        ]);
      });
    });
  });

  replaceSheetRows(sheet, rows);
}

function normalizeInventory(inventory) {
  const normalized = { sonido: [], iluminacion: [], otros: [] };
  ["sonido", "iluminacion", "otros"].forEach(categoria => {
    (inventory[categoria] || []).forEach(item => {
      const material = cleanName(item.nombre, "Nuevo material");
      (item.modelos || []).forEach(modelo => {
        upsertInventoryModel(normalized, categoria, material, {
          nombre: cleanName(modelo.nombre || modelo.modelo, "Sin modelo"),
          stock: Number(modelo.stock) || 0,
          estado: normalizeEstado(modelo.estado),
        });
      });
    });
  });
  return normalized;
}

function upsertInventoryModel(target, categoria, materialName, modelo) {
  const materialKey = String(materialName).trim().toLowerCase();
  const modelKey = String(modelo.nombre).trim().toLowerCase();
  let item = target[categoria].find(entry => String(entry.nombre).trim().toLowerCase() === materialKey);
  if (!item) {
    item = { nombre: materialName, modelos: [] };
    target[categoria].push(item);
  } else {
    item.nombre = materialName;
  }
  let modelItem = item.modelos.find(entry => String(entry.nombre).trim().toLowerCase() === modelKey);
  if (!modelItem) {
    item.modelos.push({ nombre: modelo.nombre, stock: modelo.stock, estado: modelo.estado });
  } else {
    modelItem.nombre = modelo.nombre;
    modelItem.stock = modelo.stock;
    modelItem.estado = modelo.estado;
  }
}

function cleanName(value, fallback) {
  return String(value || "").trim() || fallback;
}

function readBolos() {
  const sheet = getSheet(SHEETS.bolos, ["json"]);
  const values = readSheetValues(sheet, 1);
  return values.slice(1)
    .map(row => {
      try {
        return JSON.parse(row[0]);
      } catch (error) {
        return null;
      }
    })
    .filter(Boolean);
}

function saveBolo(bolo) {
  if (!bolo || !bolo.id) throw new Error("Missing bolo");
  const sheet = getSheet(SHEETS.bolos, ["json"]);
  const values = readSheetValues(sheet, 1);
  const rowIndex = values.findIndex((row, index) => index > 0 && safeJsonId(row[0]) === bolo.id);
  const json = JSON.stringify(bolo);

  if (rowIndex >= 0) {
    sheet.getRange(rowIndex + 1, 1).setValue(json);
  } else {
    sheet.appendRow([json]);
  }
  compactSheetRows(sheet);
}

function deleteBolo(id) {
  const sheet = getSheet(SHEETS.bolos, ["json"]);
  const values = readSheetValues(sheet, 1);
  for (let index = values.length - 1; index >= 1; index--) {
    if (safeJsonId(values[index][0]) === id) {
      sheet.deleteRow(index + 1);
    }
  }
  compactSheetRows(sheet);
}

function readTechnicians() {
  const sheet = getSheet(SHEETS.settings, ["clave", "json"]);
  const values = readSheetValues(sheet, 2);
  const row = values.slice(1).find(entry => String(entry[0] || "").trim() === "tecnicos");
  if (!row) return [];
  try {
    return normalizeTechnicians(JSON.parse(row[1] || "[]"));
  } catch (error) {
    return [];
  }
}

function writeTechnicians(values) {
  const sheet = getSheet(SHEETS.settings, ["clave", "json"]);
  const rows = readSheetValues(sheet, 2);
  const rowIndex = rows.findIndex((row, index) => index > 0 && String(row[0] || "").trim() === "tecnicos");
  const json = JSON.stringify(normalizeTechnicians(values));
  if (rowIndex >= 0) {
    sheet.getRange(rowIndex + 1, 1, 1, 2).setValues([["tecnicos", json]]);
  } else {
    sheet.appendRow(["tecnicos", json]);
  }
  compactSheetRows(sheet);
}

function normalizeTechnicians(values) {
  const seen = {};
  return (Array.isArray(values) ? values : [])
    .map(value => {
      const source = value && typeof value === "object" ? value : { nombre: value };
      const nombre = String(source.nombre || source.name || "").trim();
      const rawSpecialty = String(source.especialidad || source.specialty || "")
        .trim()
        .toUpperCase()
        .replace(/\s+/g, "")
        .replace(/[ÓÒÖÔÕ]/g, "O");
      const especialidad = rawSpecialty === "SONIDO"
        ? "SONIDO"
        : (rawSpecialty === "ILUMINACION"
          ? "ILUMINACION"
          : (rawSpecialty === "SONIDO+ILUMINACION" || rawSpecialty === "SONIDOILUMINACION"
            ? "SONIDO+ILUMINACION"
            : "AUXILIAR"));
      const rawColor = String(source.color || "").trim();
      const color = /^#[0-9a-f]{6}$/i.test(rawColor) ? rawColor.toLowerCase() : "#e0001b";
      return { nombre, especialidad, color };
    })
    .filter(technician => {
      const key = technician.nombre.toLowerCase();
      if (!technician.nombre || seen[key]) return false;
      seen[key] = true;
      return true;
    })
    .sort((a, b) => a.nombre.localeCompare(b.nombre));
}

function uploadRider(rider, boloId) {
  if (!rider || !rider.dataUrl) throw new Error("Missing rider");
  const match = String(rider.dataUrl).match(/^data:application\/pdf;base64,(.+)$/);
  if (!match) throw new Error("Invalid rider PDF");
  const bytes = Utilities.base64Decode(match[1]);
  const safeName = String(rider.name || "rider-tecnico.pdf").replace(/[\\/:*?"<>|]/g, "-");
  const fileName = `${boloId || Date.now()}-${safeName}`;
  const folder = getOrCreateFolder("SENORP Riders");
  const blob = Utilities.newBlob(bytes, "application/pdf", fileName);
  const file = folder.createFile(blob);
  let shared = true;
  let sharingWarning = "";
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    shared = false;
    sharingWarning = "El archivo se ha subido, pero la cuenta de Google no permite compartirlo públicamente.";
    console.warn(sharingWarning + " " + (error && error.message ? error.message : error));
  }
  return {
    name: safeName,
    type: "application/pdf",
    size: Number(rider.size) || bytes.length,
    fileId: file.getId(),
    url: file.getUrl(),
    downloadUrl: shared ? `https://drive.google.com/uc?export=download&id=${file.getId()}` : file.getUrl(),
    shared,
    sharingWarning
  };
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function safeJsonId(value) {
  try {
    return JSON.parse(value).id;
  } catch (error) {
    return null;
  }
}

function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}

function readSheetValues(sheet, columns) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  return sheet.getRange(1, 1, lastRow, columns).getValues();
}

function replaceSheetRows(sheet, rows) {
  const requiredRows = Math.max(rows.length, 1);
  const requiredCols = rows[0].length;

  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  sheet.clear();
  sheet.getRange(1, 1, rows.length, requiredCols).setValues(rows);

  const extraRows = sheet.getMaxRows() - requiredRows;
  if (extraRows > 0) {
    sheet.deleteRows(requiredRows + 1, extraRows);
  }
}

function compactSheetRows(sheet) {
  const columns = Math.max(sheet.getLastColumn(), 1);
  const values = readSheetValues(sheet, columns);
  const compacted = values.filter((row, index) => index === 0 || row.some(cell => String(cell || "").trim() !== ""));
  replaceSheetRows(sheet, compacted.length ? compacted : [["json"]]);
}

function normalizeCategory(value) {
  const clean = String(value || "").trim().toLowerCase();
  if (clean.indexOf("ilu") === 0) return "iluminacion";
  if (clean.indexOf("son") === 0) return "sonido";
  return "otros";
}

function normalizeEstado(value) {
  const clean = String(value || "").trim().toLowerCase();
  if (clean.indexOf("rev") === 0) return "revision";
  if (clean.indexOf("baj") === 0) return "baja";
  return "ok";
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
