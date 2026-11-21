// ============================================================================
// GOOGLE APPS SCRIPT PARA AVON AMOR — CRM RECLUTAMIENTO (ALAJUELA)
// ============================================================================
// Cómo instalar en 1 minuto:
// 1. Abre tu hoja "AVON AMOR — CRM y Programador WhatsApp (Alajuela)"
// 2. Ve al menú: Extensiones -> Apps Script
// 3. Borra todo lo que aparece y pega este código completo. Guarda (Ctrl+S).
// 4. Clic en "Implementar" -> "Nueva implementación" -> Tipo: "Aplicación web"
//    - Ejecutar como: "Yo"
//    - Quién tiene acceso: "Cualquier persona"
// 5. Autoriza los permisos y copia la "URL de la aplicación web".
// ============================================================================

const SHEET_NAME = "CRM_Reclutamiento";

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return jsonResponse({ ok: true, data: [] });
    }
    const values = sheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
    const data = [];
    for (let i = 0; i < values.length; i++) {
      const r = values[i];
      if (!r[0] && !r[1]) continue;
      data.push({
        row: i + 2,
        name: r[0] || "",
        phone: r[1] || "",
        email: r[2] || "",
        zone: r[3] || "Coyol",
        profile: r[4] || "Ama de casa / Mamá",
        source: r[5] || "Formulario Web",
        stageLabel: r[6] || "Prospecto nuevo",
        called: (r[7] === "Sí" || r[7] === "Si") ? "Si" : "No",
        lastCall: r[8] || "",
        cita: (r[9] === "Sí" || r[9] === "Si") ? "Si" : (r[9] === "No" ? "No" : "Pendiente"),
        citaDate: r[10] || "",
        nextFollowup: r[11] || "",
        notes: r[12] || ""
      });
    }
    return jsonResponse({ ok: true, data });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const action = body.action || "add";
    const sheet = getSheet();

    if (action === "add") {
      const p = body.prospect || body;
      const newRow = Math.max(2, sheet.getLastRow() + 1);
      const calledTxt = (p.called === "Si" || p.called === "Sí") ? "Sí" : "No";
      const citaTxt = (p.cita === "Si" || p.cita === "Sí") ? "Sí" : (p.cita === "No" ? "No" : "Pendiente");
      const rowValues = [[
        p.name || "",
        p.phone || "",
        p.email || "",
        p.zone || "Coyol",
        p.profile || "Ama de casa / Mamá",
        p.source || "Formulario Web",
        p.stageLabel || "Prospecto nuevo",
        calledTxt,
        p.lastCall || "",
        citaTxt,
        p.citaDate || "",
        p.nextFollowup || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        p.notes || ""
      ]];
      sheet.getRange(newRow, 1, 1, 13).setValues(rowValues);
      sheet.getRange(newRow, 14).setFormula(
        `=IF(B${newRow}="","",HYPERLINK("https://wa.me/506"&REGEXREPLACE(TO_TEXT(B${newRow}),"[^0-9]",""),"Abrir WhatsApp"))`
      );
      return jsonResponse({ ok: true, row: newRow });
    }

    if (action === "update") {
      const p = body.prospect || {};
      const rowNum = Number(p.row);
      if (!rowNum || rowNum < 2) {
        return jsonResponse({ ok: false, error: "Número de fila inválido" });
      }
      const calledTxt = (p.called === "Si" || p.called === "Sí") ? "Sí" : "No";
      const citaTxt = (p.cita === "Si" || p.cita === "Sí") ? "Sí" : (p.cita === "No" ? "No" : "Pendiente");
      const rowValues = [[
        p.name || "",
        p.phone || "",
        p.email || "",
        p.zone || "Coyol",
        p.profile || "Ama de casa / Mamá",
        p.source || "Formulario Web",
        p.stageLabel || "Prospecto nuevo",
        calledTxt,
        p.lastCall || "",
        citaTxt,
        p.citaDate || "",
        p.nextFollowup || "",
        p.notes || ""
      ]];
      sheet.getRange(rowNum, 1, 1, 13).setValues(rowValues);
      return jsonResponse({ ok: true, row: rowNum });
    }

    if (action === "delete") {
      const rowNum = Number(body.row);
      if (!rowNum || rowNum < 2) {
        return jsonResponse({ ok: false, error: "Número de fila inválido" });
      }
      sheet.deleteRow(rowNum);
      return jsonResponse({ ok: true });
    }

    return jsonResponse({ ok: false, error: "Acción desconocida" });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}
