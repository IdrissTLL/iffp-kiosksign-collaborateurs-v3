/**
 * Transforme une valeur de cellule en objet Date (JJ/MM/YYYY ou Date natif).
 * @param {*} raw – valeur brute de la cellule.
 * @returns {Date|null}
 */
function parseDateCell(raw) {
  if (raw instanceof Date) {
    return new Date(raw.getFullYear(), raw.getMonth(), raw.getDate());
  }
  if (typeof raw === 'string') {
    const parts = raw.trim().split('/');
    if (parts.length === 3) {
      const [jj, mm, aaaa] = parts.map(n => parseInt(n, 10));
      if (![jj, mm, aaaa].some(isNaN)) {
        return new Date(aaaa, mm - 1, jj);
      }
    }
  }
  return null;
}

/**
 * Formate une valeur de cellule en chaîne JJ/MM/YYYY.
 * @param {*} raw – valeur brute ou Date.
 * @returns {string|null}
 */
function formatDateCell(raw) {
  const d = parseDateCell(raw);
  return d
    ? Utilities.formatDate(d, "Europe/Paris", "dd/MM/yyyy")
    : null;
}

/**
 * Formate une valeur de cellule en chaîne HH:MM (Date ou string HH:MM).
 * @param {*} raw – valeur brute ou Date.
 * @returns {string|null}
 */
function formatTimeCell(raw) {
  let h, m;
  if (raw instanceof Date) {
    h = raw.getHours(); m = raw.getMinutes();
  } else if (typeof raw === 'string') {
    const parts = raw.trim().match(/^(\d{1,2}):(\d{2})$/);
    if (!parts) return null;
    h = parseInt(parts[1], 10); m = parseInt(parts[2], 10);
  } else {
    return null;
  }
  const hh = ('0' + h).slice(-2);
  const mm = ('0' + m).slice(-2);
  return `${hh}:${mm}`;
}

/**
 * Vérifie pour aujourd’hui (JJ/MM/YYYY) que chaque personne planifiée au matin
 * s’est connectée entre 07:00 et 11:00 et dont le shift est dépassé de ≥15 min.
 * Ne traite que les lignes à la date du jour.
 * En cas d’absence, appelle notifShiftAbsent(nomComplet, 'matin') et marque " / notif KS matin".
 */
function verifShiftAbsentMatin() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const today    = new Date();
  const todayStr = Utilities.formatDate(today, "Europe/Paris", "dd/MM/yyyy");
  const now      = new Date();
  Logger.log(`🔍 Début verifShiftAbsentMatin – date=${todayStr} heure=${Utilities.formatDate(now,"Europe/Paris","HH:mm")}`);

  // 1️⃣ Charger le planning et ne garder que les lignes du jour
  const planName = "Planning " + Utilities.formatDate(today, "Europe/Paris", "MM/yy");
  const planning = ss.getSheetByName(planName);
  if (!planning) return Logger.log(`❌ "${planName}" introuvable`);

  const [hdrP, ...rowsP] = planning.getDataRange().getValues();
  const idxP = {
    date       : hdrP.indexOf("Date"),
    nomFull    : hdrP.indexOf("Nom Complet"),
    shiftDeb   : hdrP.indexOf("Shift Matin Début"),
    remarques  : hdrP.indexOf("Remarques")
  };
  if (Object.values(idxP).some(i => i < 0)) return Logger.log("❌ Colonnes manquantes planning");

  const todayPlans = rowsP
    .map((row, i) => ({ row, rowNum: i + 2 }))
    .filter(({ row, rowNum }) => {
      const fmtDate = formatDateCell(row[idxP.date]);
      if (fmtDate !== todayStr) return false;
      if (row[idxP.date] !== fmtDate) {
        planning.getRange(rowNum, idxP.date + 1).setValue(fmtDate);
      }
      return true;
    });

  // 2️⃣ Charger les présences du jour entre 07:00 et 11:00
  const presSheet = ss.getSheetByName("Présences");
  if (!presSheet) return Logger.log("❌ 'Présences' introuvable");
  const [hdrR, ...rowsR] = presSheet.getDataRange().getValues();
  const idxR = {
    date    : hdrR.indexOf("Date"),
    nomFull : hdrR.indexOf("Nom Complet"),
    arrivee : hdrR.indexOf("Heure d'arrivée")
  };
  if (Object.values(idxR).some(i => i < 0)) return Logger.log("❌ Colonnes manquantes présences");

  const presToday = rowsR
    .map((row, i) => {
      const rowNum  = i + 2;
      let fmtDate   = formatDateCell(row[idxR.date]);
      if (fmtDate === todayStr && row[idxR.date] !== fmtDate) {
        presSheet.getRange(rowNum, idxR.date + 1).setValue(fmtDate);
      }
      const fmtTime  = formatTimeCell(row[idxR.arrivee]);
      return { name: row[idxR.nomFull].toString().trim(), fmtDate, fmtTime };
    })
    .filter(e =>
      e.fmtDate === todayStr &&
      e.fmtTime >= "07:00" &&
      e.fmtTime <= "11:00"
    )
    .map(e => e.name);

  // 3️⃣ Vérifier chaque planning du matin
  todayPlans.forEach(({ row, rowNum }) => {
    const name    = row[idxP.nomFull].toString().trim();
    const rawDeb  = row[idxP.shiftDeb];
    const fmtDeb  = formatTimeCell(rawDeb);
    const remRaw  = row[idxP.remarques] || "";

    Logger.log(`Planning matin ligne ${rowNum}: ${name}, shiftDeb=${fmtDeb}, rem='${remRaw}'`);
    if (!name || !fmtDeb) return;
    if (/(CONG(E|É)S?|FORMATION|ABSENT)/i.test(remRaw)) return;
    if (remRaw.toUpperCase().includes("NOTIF KS MATIN")) return;

    // Ne notifier que si maintenant ≥ shiftDeb + 15min
    const [h,m] = fmtDeb.split(':').map(Number);
    const shiftDT = new Date(today.getFullYear(), today.getMonth(), today.getDate(), h, m);
    const threshold = new Date(shiftDT.getTime() + 15*60000);
    if (now < threshold) {
      Logger.log(`⏭️ ${name} commence à ${fmtDeb}, <15 min écoulées → pas de notif`);
      return;
    }

    const present = presToday.includes(name);
    if (!present) {
      notifShiftAbsent(name, 'matin');
      const newRem = remRaw + (remRaw ? " / notif KS matin" : "notif KS matin");
      planning.getRange(rowNum, idxP.remarques + 1).setValue(newRem);
    }
  });

  Logger.log("✅ Fin verifShiftAbsentMatin");
}

/**
 * Vérifie pour aujourd’hui que chaque personne planifiée à l'après-midi
 * s’est connectée après le début de son shift + 15 min.
 * En cas d’absence, appelle notifShiftAbsent(nomComplet, 'aprem')
 * et marque " / Notif KS Aprem".
 */
function verifShiftAbsentApresMidi() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const today    = new Date();
  const todayStr = Utilities.formatDate(today, "Europe/Paris", "dd/MM/yyyy");
  const now      = new Date();
  Logger.log(`🔍 Début verifShiftAbsentApresMidi – date=${todayStr} heure=${Utilities.formatDate(now,"Europe/Paris","HH:mm")}`);

  const planName = "Planning " + Utilities.formatDate(today, "Europe/Paris", "MM/yy");
  const planning = ss.getSheetByName(planName);
  if (!planning) return Logger.log(`❌ "${planName}" introuvable`);

  const [hdrP, ...rowsP] = planning.getDataRange().getValues();
  const idxP = {
    date       : hdrP.indexOf("Date"),
    nomFull    : hdrP.indexOf("Nom Complet"),
    shiftDeb   : hdrP.indexOf("Shift Après-midi Début"),
    shiftFin   : hdrP.indexOf("Shift Après-midi Fin"),
    remarques  : hdrP.indexOf("Remarques")
  };
  if (Object.values(idxP).some(i => i < 0)) return Logger.log("❌ Colonnes manquantes planning");

  const todayPlans = rowsP
    .map((row, i) => ({ row, rowNum: i + 2 }))
    .filter(({ row, rowNum }) => {
      const fmtDate = formatDateCell(row[idxP.date]);
      if (fmtDate !== todayStr) return false;
      if (row[idxP.date] !== fmtDate) {
        planning.getRange(rowNum, idxP.date + 1).setValue(fmtDate);
      }
      // reformater heures
      ['shiftDeb','shiftFin'].forEach(key => {
        const raw = row[idxP[key]];
        const fmt = formatTimeCell(raw);
        if (fmt && raw !== fmt) {
          planning.getRange(rowNum, idxP[key] + 1).setValue(fmt);
        }
      });
      return true;
    });

  const presSheet = ss.getSheetByName("Présences");
  if (!presSheet) return Logger.log("❌ 'Présences' introuvable");
  const [hdrR, ...rowsR] = presSheet.getDataRange().getValues();
  const idxR = {
    date    : hdrR.indexOf("Date"),
    nomFull : hdrR.indexOf("Nom Complet"),
    arrivee : hdrR.indexOf("Heure d'arrivée")
  };
  if (Object.values(idxR).some(i => i < 0)) return Logger.log("❌ Colonnes manquantes présences");

  const presToday = rowsR
    .map((row, i) => {
      const rowNum  = i + 2;
      const fmtDate = formatDateCell(row[idxR.date]);
      if (fmtDate === todayStr && row[idxR.date] !== fmtDate) {
        presSheet.getRange(rowNum, idxR.date + 1).setValue(fmtDate);
      }
      const fmtTime  = formatTimeCell(row[idxR.arrivee]);
      return { name: row[idxR.nomFull].toString().trim(), fmtDate, fmtTime };
    })
    .filter(e => e.fmtDate === todayStr)
    .map(e => e);

  todayPlans.forEach(({ row, rowNum }) => {
    const name    = row[idxP.nomFull].toString().trim();
    const fmtDeb  = formatTimeCell(row[idxP.shiftDeb]);
    const fmtFin  = formatTimeCell(row[idxP.shiftFin]);
    const remRaw  = row[idxP.remarques] || "";

    Logger.log(`Planning aprem ligne ${rowNum}: ${name}, shiftDeb=${fmtDeb}, rem='${remRaw}'`);
    if (!name || !fmtDeb) return;
    if (/(CONG(E|É)S?|FORMATION|ABSENT)/i.test(remRaw)) return;
    if (remRaw.toUpperCase().includes("NOTIF KS APREM")) return;

    const [h,m] = fmtDeb.split(':').map(Number);
    const shiftDT = new Date(today.getFullYear(), today.getMonth(), today.getDate(), h, m);
    const threshold = new Date(shiftDT.getTime() + 15*60000);
    if (now < threshold) {
      Logger.log(`⏭️ ${name} commence aprem à ${fmtDeb}, <15 min écoulées → pas de notif`);
      return;
    }

    const present = presToday.some(e => e.name === name && e.fmtTime >= fmtDeb && e.fmtTime <= fmtFin);
    if (!present) {
      notifShiftAbsent(name, 'aprem');
      const newRem = remRaw + (remRaw ? " / Notif KS Aprem" : "Notif KS Aprem");
      planning.getRange(rowNum, idxP.remarques + 1).setValue(newRem);
    }
  });

  Logger.log("✅ Fin verifShiftAbsentApresMidi");
}

/**
 * Supprime tous les triggers verifShiftAbsentMatin et verifShiftAbsentApresMidi.
 */
function deleteShiftTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'verifShiftAbsentMatin' || fn === 'verifShiftAbsentApresMidi') {
      ScriptApp.deleteTrigger(t);
      Logger.log(`🗑️ Trigger supprimé pour ${fn}`);
    }
  });
}

/**
 * Crée des déclencheurs horaires pour :
 *  • verifShiftAbsentMatin à 08:32, 09:02, 09:32, 10:02, 10:32, 11:02
 *  • verifShiftAbsentApresMidi à 12:32, 13:02, 13:32, 14:02, 14:32, 15:02, 15:32
 */
function createShiftTriggers() {
  deleteShiftTriggers();

  // matin
  [
    {h:8, m:22}, {h:8, m:52}, {h:9, m:22}, {h:9, m:52},
    {h:10, m:22}, {h:10, m:52}
  ].forEach(({h,m}) => {
    ScriptApp.newTrigger('verifShiftAbsentMatin')
      .timeBased()
      .everyDays(1)
      .atHour(h)
      .nearMinute(m)
      .create();
    Logger.log(`⏰ verifShiftAbsentMatin à ${h}:${('0'+m).slice(-2)}`);
  });

  // après-midi
  [
    {h:12, m:52}, {h:13, m:22}, {h:13, m:52},
    {h:14, m:22}, {h:14, m:52}, {h:15, m:02},
    {h:15, m:32}
  ].forEach(({h,m}) => {
    ScriptApp.newTrigger('verifShiftAbsentApresMidi')
      .timeBased()
      .everyDays(1)
      .atHour(h)
      .nearMinute(m)
      .create();
    Logger.log(`⏰ verifShiftAbsentApresMidi à ${h}:${('0'+m).slice(-2)}`);
  });

  Logger.log("✅ Triggers matin/aprem créés.");
}

/**
 * Envoie une notification à un collaborateur pour le shift manquant.
 * @param {string} nomCompletFull – "Nom Complet" du planning.
 * @param {'matin'|'aprem'} shiftType – pour adapter horaire et message.
 */
function notifShiftAbsent(nomCompletFull, shiftType) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const today    = new Date();
  const todayStr = Utilities.formatDate(today, "Europe/Paris", "dd/MM/yyyy");

  const surname = nomCompletFull.trim().split(/\s+/).pop().toUpperCase();
  const webhook = chercherWebhookPersonnel(surname);
  if (!webhook) {
    return Logger.log(`❌ Pas de webhook pour '${surname}'`);
  }

  const planName = "Planning " + Utilities.formatDate(today, "Europe/Paris", "MM/yy");
  const planning = ss.getSheetByName(planName);
  if (!planning) return Logger.log(`❌ "${planName}" introuvable`);

  const [hdr, ...rows] = planning.getDataRange().getValues();
  const idx = {
    date     : hdr.indexOf("Date"),
    nomFull  : hdr.indexOf("Nom Complet"),
    matinDeb : hdr.indexOf("Shift Matin Début"),
    matinFin : hdr.indexOf("Shift Matin Fin"),
    apremDeb : hdr.indexOf("Shift Après-midi Début"),
    apremFin : hdr.indexOf("Shift Après-midi Fin")
  };

  let rowData = null;
  rows.some(row => {
    const fmtDate = Utilities.formatDate(new Date(row[idx.date]), "Europe/Paris", "dd/MM/yyyy");
    const nm      = row[idx.nomFull].toString().trim().split(/\s+/).pop().toUpperCase();
    if (fmtDate === todayStr && nm === surname) {
      rowData = row;
      return true;
    }
  });
  if (!rowData) return Logger.log(`⚠️ Pas de planning pour '${nomCompletFull}'`);

  const fmtTime = v => v instanceof Date
    ? Utilities.formatDate(v, "Europe/Paris", "HH:mm")
    : (typeof v==='string' && v.match(/^\d{1,2}:\d{2}$/) ? v : "N/A");

  let start, end, label;
  if (shiftType === 'matin') {
    start = fmtTime(rowData[idx.matinDeb]);
    end   = fmtTime(rowData[idx.matinFin]);
    label = "matin";
  } else {
    start = fmtTime(rowData[idx.apremDeb]);
    end   = fmtTime(rowData[idx.apremFin]);
    label = "après-midi";
  }

  const message =
    `🔔👋 Bonjour ${nomCompletFull},\n` +
    `Aujourd’hui (*${todayStr}*), vous étiez planifié(e) de *${start}* à *${end}*.\n` +
    `Aucune connexion n'a été enregistrée sur ce créneau.\n` +
    `Merci de vous connecter ou de prévenir en cas d’imprévu.\n` +
    `— L'équipe IFFP`;

  envoyerMessageWebhook(webhook, message);
  Logger.log(`✅ notifShiftAbsent('${surname}','${shiftType}') envoyé`);
}



function diagnosticCollaborateur() {
  try {
    // 🔧 À MODIFIER si tu veux tester un email précis :
    var emailTest = "nor@iffp.school"; // ← mets "" (vide) pour tester l'utilisateur connecté

    const fichierDestinationId = "1Dn5HfoYTegM9nJR4LKfrqOJ65CgJ_1crtiNFFUd2LMU";
    const feuilleNom = "liste collaborateurs";

    // 🧩 Normalisation des emails (enlève espaces invisibles, majuscules, etc.)
    const clean = s => (s || "")
      .toString()
      .replace(/\u00A0/g, " ")
      .replace(/[\u200B-\u200D\uFEFF]/g, "")
      .trim()
      .toLowerCase();

    // 📧 Email à rechercher
    var emailRecherche = emailTest && emailTest.trim() !== "" 
      ? clean(emailTest) 
      : clean(Session.getActiveUser().getEmail());
    Logger.log("🔍 Diagnostic pour : " + emailRecherche);

    // 📂 Ouvrir le fichier et la feuille
    const fichier = SpreadsheetApp.openById(fichierDestinationId);
    const feuille = fichier.getSheetByName(feuilleNom);
    if (!feuille) throw new Error('Feuille "' + feuilleNom + '" introuvable.');

    const data = feuille.getDataRange().getValues();
    const headers = data[0].map(h => (h + "").trim());
    const colEmail = headers.findIndex(h => h.toLowerCase().replace(/\s|-/g, "") === "email");
    if (colEmail === -1) throw new Error("Colonne 'Email' introuvable.");

    let trouve = false;

    // 🔎 Recherche exacte
    for (let i = 1; i < data.length; i++) {
      const brut = data[i][colEmail];
      const emailCell = clean(brut);
      if (emailCell === emailRecherche) {
        Logger.log("✅ Email trouvé ligne " + (i + 1) + " : " + brut);
        Logger.log("🧾 Ligne complète : " + JSON.stringify(data[i]));
        const lastChar = (brut + "").charCodeAt((brut + "").length - 1);
        Logger.log("🔡 Code du dernier caractère : " + lastChar);
        trouve = true;
        break;
      }
    }

    // 🧐 Recherche approchante si rien trouvé
    if (!trouve) {
      Logger.log("❌ Aucune correspondance exacte trouvée.");
      Logger.log("🔎 Emails approchants (même partie avant @) :");
      const local = emailRecherche.split("@")[0];
      for (let i = 1; i < data.length; i++) {
        const cell = clean(data[i][colEmail]);
        if (cell.includes(local)) {
          Logger.log("→ Ligne " + (i + 1) + " : " + data[i][colEmail]);
        }
      }
    }

  } catch (err) {
    Logger.log("🚨 Erreur diagnosticCollaborateur : " + err.message);
  }
}