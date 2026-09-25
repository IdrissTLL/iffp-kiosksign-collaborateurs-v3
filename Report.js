/*******************************************************
 * KIOSKSIGN Collaborateurs — Récap mensuel + emails
 * Regroupé par JOUR (1 ligne / date)
 * UI-safe (pas de getUi() en contexte trigger)
 *******************************************************/

const KIOSK_COLLAB_CONFIG = {
  SOURCE_SHEET_NAME: "Présences",
  /** true = tous les récaps partent vers TEST_EMAIL (aucun envoi aux collaborateurs) */
  TEST_MODE: false,
  TEST_EMAIL: "idriss+testkiosksign@iffp.school",
  SENDER_NAME: "KioskSign IFFP - NOREPLY",
  REPLY_TO: "noreply@iffp.school",
  FROM: "noreply@iffp.school"
};

/* ================== REPORTING (ONGLETS) ================== */

/**
 * Exécution manuelle (UI) : mois courant + alert.
 */
function genererRecapCollaborateursMoisCourant() {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const now = new Date();
  const mm = Utilities.formatDate(now, tz, "MM");
  const yy = Utilities.formatDate(now, tz, "yy");
  genererRecapCollaborateursDansOnglet_(mm + "/" + yy, true);
}

/**
 * Exécution manuelle (UI) : mois précédent + alert.
 */
function genererRecapCollaborateursMoisPrecedent() {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const mm = Utilities.formatDate(prev, tz, "MM");
  const yy = Utilities.formatDate(prev, tz, "yy");
  genererRecapCollaborateursDansOnglet_(mm + "/" + yy, true);
}

/**
 * Exécution manuelle (UI) : prompt mois.
 */
function genererRecapCollaborateursMoisPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt("Mois à générer", "Entre le mois au format MM/YY (ex: 02/26)", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const monthStr = (res.getResponseText() || "").trim();
  if (!/^\d{2}\/\d{2}$/.test(monthStr)) {
    ui.alert("Format invalide. Exemple attendu : 02/26");
    return;
  }
  genererRecapCollaborateursDansOnglet_(monthStr, true);
}

/**
 * Exécution automatique (trigger) : mois courant, sans UI.
 */
function TRIGGER_genererRecapCollaborateursMoisCourant() {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const now = new Date();
  const mm = Utilities.formatDate(now, tz, "MM");
  const yy = Utilities.formatDate(now, tz, "yy");
  genererRecapCollaborateursDansOnglet_(mm + "/" + yy, false);
}

/**
 * Coeur : crée/MAJ l’onglet "MM/YY" avec blocs par collaborateur,
 * regroupés par JOUR. UI-safe.
 *
 * @param {string} monthStr "MM/YY"
 * @param {boolean} notifyUi si true => ui.alert en fin (uniquement contexte UI)
 */
function genererRecapCollaborateursDansOnglet_(monthStr, notifyUi) {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  const src = ss.getSheetByName(KIOSK_COLLAB_CONFIG.SOURCE_SHEET_NAME);
  if (!src) throw new Error('Feuille source introuvable : "' + KIOSK_COLLAB_CONFIG.SOURCE_SHEET_NAME + '"');

  const { start, end } = parseMonthMMYY_(monthStr);

  const data = src.getDataRange().getValues();
  if (data.length < 2) {
    if (notifyUi) SpreadsheetApp.getUi().alert('La feuille "' + KIOSK_COLLAB_CONFIG.SOURCE_SHEET_NAME + '" ne contient aucune donnée.');
    return;
  }

  const header = data[0].map(String);
  const idx = indexHeadersCollaborateursReporting_(header);

  const byCollab = {}; // collab -> { days: {dateStr: bucket} }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const d = coerceDate_(row[idx["Date"]], tz);
    if (!d) continue;
    if (d < start || d >= end) continue;

    const name = safe_(row[idx["Nom Complet"]]).trim();
    if (!name) continue;

    const dateStr = Utilities.formatDate(d, tz, "dd/MM/yyyy");

    const hValid = parseNumberFR_(row[idx["Nb H Valid"]]);
    const hRaw = parseNumberFR_(row[idx["Nb heures"]]);
    const nbHeures = (isFinite(hValid) && hValid > 0) ? hValid : hRaw;

    const hArr = formatTimeHHMM_(row[idx["Heure d'arrivée"]], tz);
    const hDep = formatTimeHHMM_(row[idx["Heure de départ"]], tz);
    const line = buildDetailLine_(hArr, hDep, nbHeures);

    const ipLogin = safe_(row[idx["Adresse IP Login"]]).trim();
    const ipLogout = safe_(row[idx["Adresse IP Logout"]]).trim();
    const ip = ipLogin || ipLogout || "";

    const rem = safe_(row[idx["Remarques"]]).trim();

    if (!byCollab[name]) byCollab[name] = { days: {} };
    if (!byCollab[name].days[dateStr]) {
      byCollab[name].days[dateStr] = {
        dateObj: d,
        details: [],
        totalDay: 0,
        ips: new Set(),
        remarks: new Set()
      };
    }

    const bucket = byCollab[name].days[dateStr];
    const safeHours = (isFinite(nbHeures) ? nbHeures : 0);
    bucket.totalDay += safeHours;

    if (line) bucket.details.push(line);
    if (ip) bucket.ips.add(ip);
    if (rem) bucket.remarks.add(rem);
  }

  const targetName = monthStr;
  let sh = ss.getSheetByName(targetName);
  if (!sh) sh = ss.insertSheet(targetName);

  sh.clear({ contentsOnly: false });
  sh.setHiddenGridlines(false);
  sh.setFrozenRows(0);

  sh.setColumnWidths(1, 1, 120);
  sh.setColumnWidths(2, 1, 320);
  sh.setColumnWidths(3, 1, 100);
  sh.setColumnWidths(4, 1, 150);
  sh.setColumnWidths(5, 1, 450);

  const collabs = Object.keys(byCollab).sort((a, b) => a.localeCompare(b, "fr", { sensitivity: "base" }));

  if (collabs.length === 0) {
    sh.getRange(1, 1).setValue("Aucune donnée pour " + monthStr + ' (source : ' + KIOSK_COLLAB_CONFIG.SOURCE_SHEET_NAME + ")");
    sh.getRange(1, 1).setFontWeight("bold");
    if (notifyUi) SpreadsheetApp.getUi().alert("Aucune donnée trouvée pour " + monthStr);
    return;
  }

  const out = [];
  const styleMap = [];

  collabs.forEach((name, k) => {
    const days = byCollab[name].days;
    const dateKeys = Object.keys(days).sort((a, b) => days[a].dateObj.getTime() - days[b].dateObj.getTime());

    out.push([name, "", "", "", ""]);
    styleMap.push({ type: "collab" });

    out.push(["Date", "Détails (shifts)", "Nb Heures J", "IP", "Remarques"]);
    styleMap.push({ type: "blockHeader" });

    let totalMonth = 0;

    dateKeys.forEach(ds => {
      const b = days[ds];
      totalMonth += (isFinite(b.totalDay) ? b.totalDay : 0);

      const detailsJoined = (b.details && b.details.length) ? b.details.join("\n") : "";
      const ipJoined = Array.from(b.ips).join(" / ");
      const remJoined = Array.from(b.remarks).join(" | ");

      out.push([ds, detailsJoined, formatHours_(b.totalDay), ipJoined, remJoined]);
      styleMap.push({ type: "row" });
    });

    out.push(["Total Mois", "", formatHours_(totalMonth), "", ""]);
    styleMap.push({ type: "total" });

    if (k < collabs.length - 1) {
      out.push(["", "", "", "", ""]);
      styleMap.push({ type: "spacer" });
    }
  });

  const range = sh.getRange(1, 1, out.length, 5);
  range.setValues(out);

  sh.getRange(1, 2, out.length, 1).setWrap(true);

  applyStylesCollaborateurs_(sh, styleMap);

  const full = sh.getRange(1, 1, out.length, 5);
  full.setBorder(true, true, true, true, true, true);

  for (let i = 0; i < styleMap.length; i++) {
    if (styleMap[i].type === "spacer") {
      sh.getRange(i + 1, 1, 1, 5).setBorder(false, false, false, false, false, false);
    }
  }

  sh.getRange(1, 7).setValue("Source: " + KIOSK_COLLAB_CONFIG.SOURCE_SHEET_NAME);
  sh.getRange(2, 7).setValue(
    "Période: " + Utilities.formatDate(start, tz, "dd/MM/yyyy") +
    " → " + Utilities.formatDate(new Date(end.getTime() - 1), tz, "dd/MM/yyyy")
  );
  sh.getRange(1, 7, 2, 1).setFontColor("#666666").setFontSize(9);

  if (notifyUi) SpreadsheetApp.getUi().alert("Récap collaborateurs (regroupé par jour) généré / mis à jour : " + targetName);
}

/* ================== STYLES ================== */

function applyStylesCollaborateurs_(sheet, styleMap) {
  for (let i = 0; i < styleMap.length; i++) {
    const row = i + 1;
    const t = styleMap[i].type;
    const r = sheet.getRange(row, 1, 1, 5);

    if (t === "collab") {
      r.setFontWeight("bold").setFontSize(12).setBackground("#F3F6FF");
    } else if (t === "blockHeader") {
      r.setFontWeight("bold").setBackground("#E9EEF9");
    } else if (t === "total") {
      r.setFontWeight("bold").setBackground("#F7F7F7");
    } else if (t === "spacer") {
      r.setBackground(null);
    } else {
      r.setFontWeight("normal");
    }
  }

  const lastRow = styleMap.length;
  sheet.getRange(1, 1, lastRow, 1).setHorizontalAlignment("left");
  sheet.getRange(1, 2, lastRow, 1).setHorizontalAlignment("left");
  sheet.getRange(1, 3, lastRow, 1).setHorizontalAlignment("center");
  sheet.getRange(1, 4, lastRow, 1).setHorizontalAlignment("center");
  sheet.getRange(1, 5, lastRow, 1).setHorizontalAlignment("left");
}

/* ================== HEADER INDEXERS ================== */

function indexHeadersCollaborateursReporting_(headerRow) {
  const need = [
    "Date",
    "Nom Complet",
    "Heure d'arrivée",
    "Heure de départ",
    "Nb heures",
    "Nb H Valid",
    "Adresse IP Login",
    "Adresse IP Logout",
    "Remarques"
  ];
  return buildHeaderIndexMap_(headerRow, need, KIOSK_COLLAB_CONFIG.SOURCE_SHEET_NAME);
}

function buildHeaderIndexMap_(headerRow, need, sheetName) {
  const normalized = headerRow.map(h => normalizeHeader_(h));
  const map = {};
  need.forEach(n => {
    const i = normalized.indexOf(normalizeHeader_(n));
    if (i === -1) throw new Error("Colonne introuvable dans '" + sheetName + "' : " + n);
    map[n] = i;
  });
  return map;
}

/* ================== HELPERS ================== */

function normalizeHeader_(s) {
  return String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function safe_(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}

function coerceDate_(v, tz) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  const s = safe_(v).trim();
  if (!s) return null;

  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));

  const d = new Date(s);
  if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return null;
}

function parseNumberFR_(v) {
  if (typeof v === "number") return v;
  const s = safe_(v).trim();
  if (!s) return 0;
  const cleaned = s.replace(/\s/g, "").replace(",", ".");
  const n = parseFloat(cleaned);
  return isFinite(n) ? n : 0;
}

function parseMonthMMYY_(monthStr) {
  const m = monthStr.match(/^(\d{2})\/(\d{2})$/);
  if (!m) throw new Error("Format monthStr invalide : " + monthStr);

  const mm = parseInt(m[1], 10);
  const yy = parseInt(m[2], 10);
  if (mm < 1 || mm > 12) throw new Error("Mois invalide : " + monthStr);

  const fullYear = 2000 + yy;
  return { start: new Date(fullYear, mm - 1, 1), end: new Date(fullYear, mm, 1) };
}

function formatHours_(n) {
  if (!isFinite(n)) n = 0;
  const rounded = Math.round(n * 100) / 100;
  let s = String(rounded);
  if (s.indexOf(".") >= 0) s = s.replace(/0+$/, "").replace(/\.$/, "");
  return s.replace(".", ",");
}

function formatTimeHHMM_(v, tz) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, tz, "HH:mm");
  if (typeof v === "number" && isFinite(v)) {
    const totalMinutes = Math.round(v * 24 * 60);
    const hh = String(Math.floor(totalMinutes / 60)).padStart(2, "0");
    const mm = String(totalMinutes % 60).padStart(2, "0");
    return hh + ":" + mm;
  }
  const s = safe_(v).trim();
  if (!s) return "";
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return String(m[1]).padStart(2, "0") + ":" + m[2];
  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, "HH:mm");
  return s;
}

function buildDetailLine_(hArr, hDep, nbHeures) {
  const hasTimes = (hArr || hDep);
  const hTxt = formatHours_(nbHeures);
  if (!hasTimes && (!isFinite(nbHeures) || nbHeures === 0)) return "";
  const left = (hArr || "??:??");
  const right = (hDep || "??:??");
  return left + " - " + right + " => " + hTxt + " H";
}

function startOfWeekMonday_(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const day = x.getDay();
  const diff = (day === 0 ? -6 : 1 - day);
  x.setDate(x.getDate() + diff);
  x.setHours(0, 0, 0, 0);
  return x;
}

function isLastDayOfMonth_(d) {
  const last = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  return d.getDate() === last;
}

function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

/* ================== ENVOI EMAIL RÉCAP HEURES ================== */

/**
 * Index headers pour l'envoi email (Email obligatoire).
 */
function indexHeadersCollaborateursEmailReporting_(headerRow) {
  const need = [
    "Date",
    "Email",
    "Nom Complet",
    "Prénom",
    "Nom",
    "Heure d'arrivée",
    "Heure de départ",
    "Nb heures",
    "Nb H Valid",
    "Remarques"
  ];
  const normalized = headerRow.map(h => normalizeHeader_(h));
  const map = {};
  need.forEach(n => {
    const i = normalized.indexOf(normalizeHeader_(n));
    // Prénom / Nom / Remarques optionnels
    if (i === -1 && (n === "Prénom" || n === "Nom" || n === "Remarques")) {
      map[n] = -1;
      return;
    }
    if (i === -1) throw new Error("Colonne introuvable dans '" + KIOSK_COLLAB_CONFIG.SOURCE_SHEET_NAME + "' : " + n);
    map[n] = i;
  });
  return map;
}

/**
 * Agrège les présences par email sur [start, end[.
 * @returns {Array<{email:string, name:string, days:Array, totalPeriod:number}>}
 */
function collectCollaboratorRecapsForPeriod_(start, end) {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const src = ss.getSheetByName(KIOSK_COLLAB_CONFIG.SOURCE_SHEET_NAME);
  if (!src) throw new Error('Feuille source introuvable : "' + KIOSK_COLLAB_CONFIG.SOURCE_SHEET_NAME + '"');

  const data = src.getDataRange().getValues();
  if (data.length < 2) return [];

  const idx = indexHeadersCollaborateursEmailReporting_(data[0].map(String));
  const byEmail = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const d = coerceDate_(row[idx["Date"]], tz);
    if (!d || d < start || d >= end) continue;

    const email = safe_(row[idx["Email"]]).trim().toLowerCase();
    if (!email || email.indexOf("@") === -1) continue;

    const prenom = idx["Prénom"] >= 0 ? safe_(row[idx["Prénom"]]).trim() : "";
    const nom = idx["Nom"] >= 0 ? safe_(row[idx["Nom"]]).trim() : "";
    const nomComplet = safe_(row[idx["Nom Complet"]]).trim();
    const displayName = (prenom || nom) ? (prenom + " " + nom).trim() : (nomComplet || email);

    const hValid = parseNumberFR_(row[idx["Nb H Valid"]]);
    const hRaw = parseNumberFR_(row[idx["Nb heures"]]);
    const nbHeures = (isFinite(hValid) && hValid > 0) ? hValid : hRaw;

    const hArr = formatTimeHHMM_(row[idx["Heure d'arrivée"]], tz);
    const hDep = formatTimeHHMM_(row[idx["Heure de départ"]], tz);
    const line = buildDetailLine_(hArr, hDep, nbHeures);
    const rem = idx["Remarques"] >= 0 ? safe_(row[idx["Remarques"]]).trim() : "";
    const dateStr = Utilities.formatDate(d, tz, "dd/MM/yyyy");

    if (!byEmail[email]) {
      byEmail[email] = { email: email, name: displayName, days: {} };
    } else if (!byEmail[email].name || byEmail[email].name === email) {
      byEmail[email].name = displayName;
    }

    if (!byEmail[email].days[dateStr]) {
      byEmail[email].days[dateStr] = {
        dateObj: d,
        details: [],
        totalDay: 0,
        remarks: []
      };
    }

    const bucket = byEmail[email].days[dateStr];
    const safeHours = isFinite(nbHeures) ? nbHeures : 0;
    bucket.totalDay += safeHours;
    if (line) bucket.details.push(line);
    if (rem && bucket.remarks.indexOf(rem) === -1) bucket.remarks.push(rem);
  }

  return Object.keys(byEmail).sort().map(email => {
    const collab = byEmail[email];
    const dateKeys = Object.keys(collab.days).sort(
      (a, b) => collab.days[a].dateObj.getTime() - collab.days[b].dateObj.getTime()
    );
    let totalPeriod = 0;
    const days = dateKeys.map(ds => {
      const b = collab.days[ds];
      totalPeriod += isFinite(b.totalDay) ? b.totalDay : 0;
      return {
        dateStr: ds,
        details: b.details,
        totalDay: b.totalDay,
        remarks: b.remarks
      };
    });
    return { email: collab.email, name: collab.name, days: days, totalPeriod: totalPeriod };
  });
}

function buildRecapEmailHtml_(collab, periodLabel, type, destNote) {
  const rows = collab.days.map(day => {
    const details = day.details.length
      ? day.details.map(x => escapeHtml_(x)).join("<br>")
      : "—";
    const rem = day.remarks.length
      ? day.remarks.map(x => escapeHtml_(x)).join(" | ")
      : "";
    return (
      "<tr>" +
      '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;white-space:nowrap;">' + escapeHtml_(day.dateStr) + "</td>" +
      '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">' + details +
        (rem ? '<div style="margin-top:4px;color:#6b7280;font-size:12px;">' + rem + "</div>" : "") +
      "</td>" +
      '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;text-align:center;">' + escapeHtml_(formatHours_(day.totalDay)) + " H</td>" +
      "</tr>"
    );
  }).join("");

  const typeLabel = type === "MENSUEL" ? "mensuel" : "hebdomadaire";
  const testBanner = destNote
    ? '<p style="margin:0 0 14px 0;padding:10px 12px;background:#FEF3C7;border:1px solid #F59E0B;border-radius:6px;color:#92400E;font-size:13px;">' +
      "<strong>MODE TEST</strong> — destinataire réel prévu : " + escapeHtml_(destNote) +
      "</p>"
    : "";

  return (
    '<div style="font-family:Arial,sans-serif;color:#111827;line-height:1.45;max-width:720px;">' +
      testBanner +
      "<p style=\"margin:0 0 10px 0;\">Bonjour " + escapeHtml_(collab.name) + ",</p>" +
      "<p style=\"margin:0 0 14px 0;\">Voici votre récapitulatif <strong>" + escapeHtml_(typeLabel) +
        "</strong> des heures de travail — <strong>" + escapeHtml_(periodLabel) + "</strong>.</p>" +
      '<table style="border-collapse:collapse;width:100%;font-size:14px;margin:0 0 14px 0;">' +
        "<thead><tr>" +
          '<th style="text-align:left;padding:8px 10px;background:#EEF2FF;border-bottom:2px solid #c7d2fe;">Date</th>' +
          '<th style="text-align:left;padding:8px 10px;background:#EEF2FF;border-bottom:2px solid #c7d2fe;">Détails (shifts)</th>' +
          '<th style="text-align:center;padding:8px 10px;background:#EEF2FF;border-bottom:2px solid #c7d2fe;">Heures</th>' +
        "</tr></thead>" +
        "<tbody>" + rows + "</tbody>" +
        "<tfoot><tr>" +
          '<td colspan="2" style="padding:10px;font-weight:bold;background:#F9FAFB;">Total période</td>' +
          '<td style="padding:10px;font-weight:bold;text-align:center;background:#F9FAFB;">' +
            escapeHtml_(formatHours_(collab.totalPeriod)) + " H</td>" +
        "</tr></tfoot>" +
      "</table>" +
      '<p style="margin:0;color:#6b7280;font-size:12px;">Message automatique KIOSKSIGN — ne pas répondre</p>' +
    "</div>"
  );
}

function buildRecapEmailText_(collab, periodLabel, type, destNote) {
  const typeLabel = type === "MENSUEL" ? "mensuel" : "hebdomadaire";
  const lines = [
    "Bonjour " + collab.name + ",",
    "",
    "Voici votre récapitulatif " + typeLabel + " des heures de travail — " + periodLabel + ".",
    ""
  ];
  if (destNote) {
    lines.push("[MODE TEST] Destinataire réel prévu : " + destNote);
    lines.push("");
  }
  collab.days.forEach(day => {
    lines.push(day.dateStr + " — " + formatHours_(day.totalDay) + " H");
    day.details.forEach(d => lines.push("  • " + d));
    if (day.remarks.length) lines.push("  Remarques : " + day.remarks.join(" | "));
  });
  lines.push("");
  lines.push("Total période : " + formatHours_(collab.totalPeriod) + " H");
  lines.push("");
  lines.push("Message automatique KIOSKSIGN");
  return lines.join("\n");
}

/**
 * Moteur d'envoi : 1 email par collaborateur ayant des présences sur la période.
 * En TEST_MODE, tous les mails partent vers TEST_EMAIL (serge@...).
 *
 * @param {Date} start début inclus
 * @param {Date} end fin exclusive
 * @param {string} periodLabel libellé affiché
 * @param {string} type "HEBDO" | "MENSUEL"
 * @returns {{sent:number, skipped:number, mode:string}}
 */
function sendRecapToAllCollaboratorsForPeriod_(start, end, periodLabel, type) {
  const collabs = collectCollaboratorRecapsForPeriod_(start, end);
  const testMode = !!KIOSK_COLLAB_CONFIG.TEST_MODE;
  const testEmail = (KIOSK_COLLAB_CONFIG.TEST_EMAIL || "").trim().toLowerCase();

  if (testMode && !testEmail) {
    throw new Error("TEST_MODE activé mais TEST_EMAIL vide.");
  }

  let sent = 0;
  let skipped = 0;

  for (const collab of collabs) {
    if (!collab.days.length) {
      skipped++;
      continue;
    }

    const to = testMode ? testEmail : collab.email;
    const destNote = testMode ? (collab.name + " <" + collab.email + ">") : "";
    const prefix = testMode ? "[TEST] " : "";
    const subject =
      prefix + "[KIOSKSIGN] " + periodLabel + " — " + collab.name +
      (testMode ? " (pour " + collab.email + ")" : "");

    const html = buildRecapEmailHtml_(collab, periodLabel, type, destNote);
    const text = buildRecapEmailText_(collab, periodLabel, type, destNote);

    const options = {
      htmlBody: html,
      name: KIOSK_COLLAB_CONFIG.SENDER_NAME,
      replyTo: KIOSK_COLLAB_CONFIG.REPLY_TO
    };
    // from: uniquement si alias configuré côté Gmail du compte exécutant
    if (KIOSK_COLLAB_CONFIG.FROM) {
      options.from = KIOSK_COLLAB_CONFIG.FROM;
    }

    try {
      MailApp.sendEmail(to, subject, text, options);
      sent++;
    } catch (err) {
      // Retry sans from si l'alias n'est pas autorisé
      if (options.from) {
        delete options.from;
        MailApp.sendEmail(to, subject, text, options);
        sent++;
      } else {
        Logger.log("Échec envoi récap à " + to + " (prévu " + collab.email + ") : " + err);
        skipped++;
      }
    }
  }

  const summary =
    "Récap " + type + " — " + periodLabel +
    " | collabs=" + collabs.length +
    " | envoyés=" + sent +
    " | ignorés=" + skipped +
    " | mode=" + (testMode ? "TEST→" + testEmail : "PROD");
  Logger.log(summary);
  return { sent: sent, skipped: skipped, mode: testMode ? "TEST" : "PROD", total: collabs.length };
}

/** Envoi hebdo : lundi → samedi (exclus), typiquement lancé le vendredi. */
function sendWeeklyRecapCollaborators_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const now = new Date();
  const start = startOfWeekMonday_(now);
  const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 5); // samedi 00:00
  const startLabel = Utilities.formatDate(start, tz, "dd/MM/yyyy");
  const endLabel = Utilities.formatDate(new Date(end.getTime() - 1), tz, "dd/MM/yyyy");
  const periodLabel = "Récap hebdo " + startLabel + " → " + endLabel;
  return sendRecapToAllCollaboratorsForPeriod_(start, end, periodLabel, "HEBDO");
}

/** Trigger vendredi : envoie le récap hebdo, puis le mensuel si dernier vendredi. */
function sendFridayRecapCollaboratorsTrigger() {
  const weekly = sendWeeklyRecapCollaborators_();
  Logger.log("Hebdo terminé : " + JSON.stringify(weekly));
  sendMonthlyRecapCollaboratorsIfLastFriday();
}

/**
 * Handler mensuel : uniquement le DERNIER vendredi du mois.
 */
function sendMonthlyRecapCollaboratorsIfLastFriday() {
  const now = new Date();
  if (!isLastFridayOfMonth_(now)) {
    Logger.log("Pas le dernier vendredi du mois — récap mensuel ignoré.");
    return null;
  }
  return sendMonthlyRecapCollaboratorsForCurrentMonth_();
}

/** Envoi mensuel mois courant (sans condition de date). */
function sendMonthlyRecapCollaboratorsForCurrentMonth_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const end = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const monthStr = Utilities.formatDate(now, tz, "MM/yy");
  const periodLabel = "Récap mensuel " + monthStr;
  return sendRecapToAllCollaboratorsForPeriod_(start, end, periodLabel, "MENSUEL");
}

/**
 * Tests manuels menu / éditeur.
 * Forcent toujours l'envoi vers TEST_EMAIL (même si TEST_MODE=false en prod).
 */
function TEST_sendWeeklyRecapCollaborators() {
  return runRecapTest_(sendWeeklyRecapCollaborators_, "hebdo");
}

function TEST_sendMonthlyRecapCollaborators() {
  return runRecapTest_(sendMonthlyRecapCollaboratorsForCurrentMonth_, "mensuel");
}

function runRecapTest_(fn, label) {
  const prev = KIOSK_COLLAB_CONFIG.TEST_MODE;
  KIOSK_COLLAB_CONFIG.TEST_MODE = true;
  let result;
  try {
    result = fn();
  } finally {
    KIOSK_COLLAB_CONFIG.TEST_MODE = prev;
  }
  try {
    SpreadsheetApp.getUi().alert(
      "Test " + label + " terminé.\n" +
      "Mode forcé : TEST → " + KIOSK_COLLAB_CONFIG.TEST_EMAIL + "\n" +
      "Collaborateurs : " + result.total + "\n" +
      "Envoyés : " + result.sent + "\n" +
      "Ignorés : " + result.skipped + "\n" +
      "(TEST_MODE prod reste : " + prev + ")"
    );
  } catch (e) {
    Logger.log(JSON.stringify(result));
  }
  return result;
}

/**
 * Installe les triggers :
 * 1) logout 13h35 / 19h35 tous les jours
 * 2) récap email vendredi ~19h (hebdo + mensuel si dernier vendredi)
 */
function installerTriggersLogoutCollaborateurs() {
  const HANDLERS = [
    "logoutAllForgottenShifts_1335",
    "logoutAllForgottenShifts_1935",
    "sendFridayRecapCollaboratorsTrigger"
  ];

  ScriptApp.getProjectTriggers().forEach(t => {
    if (HANDLERS.includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("logoutAllForgottenShifts_1335")
    .timeBased()
    .everyDays(1)
    .atHour(13)
    .nearMinute(35)
    .create();

  ScriptApp.newTrigger("logoutAllForgottenShifts_1935")
    .timeBased()
    .everyDays(1)
    .atHour(19)
    .nearMinute(35)
    .create();

  // Vendredi ~19h : hebdo + mensuel (si dernier vendredi)
  ScriptApp.newTrigger("sendFridayRecapCollaboratorsTrigger")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(19)
    .nearMinute(0)
    .create();

  const mode = KIOSK_COLLAB_CONFIG.TEST_MODE
    ? ("MODE TEST → " + KIOSK_COLLAB_CONFIG.TEST_EMAIL)
    : "MODE PROD (emails collaborateurs)";

  try {
    SpreadsheetApp.getUi().alert(
      "Triggers installés :\n" +
      "• Logout collab : 13h35 et 19h35 (tous les jours)\n" +
      "• Récap heures email : vendredi ~19h (hebdo + mensuel si dernier vendredi)\n\n" +
      mode
    );
  } catch (e) {}
}

/**
 * Retourne true si la date est le dernier vendredi du mois.
 */
function isLastFridayOfMonth_(d) {
  // vendredi = 5 (0=dim,1=lun,...5=ven)
  if (d.getDay() !== 5) return false;

  const y = d.getFullYear();
  const m = d.getMonth();

  // dernier jour du mois
  const lastDay = new Date(y, m + 1, 0);

  // reculer jusqu'au dernier vendredi
  const x = new Date(lastDay.getFullYear(), lastDay.getMonth(), lastDay.getDate());
  while (x.getDay() !== 5) {
    x.setDate(x.getDate() - 1);
  }

  return (
    d.getFullYear() === x.getFullYear() &&
    d.getMonth() === x.getMonth() &&
    d.getDate() === x.getDate()
  );
}