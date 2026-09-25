

/**
 * Nettoie et complète les heures de présence dans la feuille "Présences",
 * en ne traitant que les lignes dont la date est dans les derniers `daysRange` jours.
 * @param {number} daysRange 0 = aujourd'hui uniquement, 10 = aujourd'hui et les 9 jours précédents
 */
function nettoyerPresences() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Présences");
  if (!feuille) {
    Logger.log("❌ Feuille 'Présences' introuvable.");
    return;
  }

  // ─── CONFIGURATION ───────────────────────────────────────────────────────────
  const daysRange = 0; // 0 = aujourd’hui, 10 = depuis 10 jours (inclus)
  const today     = new Date();
  // On ne garde que la date (minuit)
  const todayMid  = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const msPerDay  = 24 * 60 * 60 * 1000;

  // ─── CHARGEMENT DES DONNÉES ────────────────────────────────────────────────────
  const data    = feuille.getDataRange().getValues();
  const headers = data[0];
  const rows    = data.slice(1);
  const nRows   = rows.length;

  // Index des colonnes par nom
  const idx = {
    date       : headers.indexOf("Date"),
    debut      : headers.indexOf("Heure d'arrivée"),
    fin        : headers.indexOf("Heure de départ"),
    nbH        : headers.indexOf("Nb heures"),
    nomComplet : headers.indexOf("Nom Complet")
  };

  // Ajouter la colonne "Nb H Valid" si manquante
  let idxNbHValid = headers.indexOf("Nb H Valid");
  if (idxNbHValid === -1) {
    idxNbHValid = headers.length;
    feuille.getRange(1, idxNbHValid + 1).setValue("Nb H Valid");
  }

  // Vérifier qu'on a bien les colonnes indispensables
  if ([idx.date, idx.debut, idx.fin, idx.nbH, idx.nomComplet].includes(-1)) {
    Logger.log("❌ Colonnes manquantes (Date, Heure d'arrivée, Heure de départ, Nb heures ou Nom Complet).");
    return;
  }

  const doublons = {};

  // ─── TRAITEMENT LIGNE PAR LIGNE ────────────────────────────────────────────────
  rows.forEach((ligne, i) => {
    const rowNum = i + 2;

    // ── 1. Lire et normaliser la date
    const rawDate = ligne[idx.date];
    let dateObj;
    if (rawDate instanceof Date) {
      // Si la cellule est de type Date
      dateObj = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate());
    } else if (typeof rawDate === "string") {
      // Si c’est une chaîne "JJ/MM/YYYY"
      const [jj, mm, aaaa] = rawDate.split("/");
      dateObj = new Date(parseInt(aaaa,10), parseInt(mm,10)-1, parseInt(jj,10));
    } else {
      return; // pas de date valide → passer
    }

    // ── 2. Filtrer par plage de dates
    const diffDays = Math.floor((todayMid - dateObj) / msPerDay);
    if (diffDays < 0 || diffDays > daysRange) {
      return; // hors de la plage définie → on ne traite pas cette ligne
    }

    // ── 3. Récupérer les autres champs
    const nomComplet = (ligne[idx.nomComplet] || "")
      .toString().trim().replace(/\s+/g, " ");
    const debutRaw   = ligne[idx.debut];
    const finRaw     = ligne[idx.fin];
    let   nbH        = ligne[idx.nbH];
    const rangeNbH   = feuille.getRange(rowNum, idx.nbH + 1);
    const rangeNbHValid = feuille.getRange(rowNum, idxNbHValid + 1);

    if (!nomComplet) {
      return; // sans nom, rien à faire
    }

    // ── 4. Gestion des doublons par date
    const key = `${nomComplet}_${dateObj.toDateString()}`;
    doublons[key] = (doublons[key] || 0) + 1;

    // ── 5. Calcul ou lecture des heures
    let heures = 0;
    if (!nbH || nbH === "") {
      // Si la colonne Nb heures est vide, on la calcule
      if (debutRaw && finRaw) {
        try {
          const h1 = typeof debutRaw === "string"
            ? parseHeureToDate(debutRaw)
            : new Date(debutRaw);
          const h2 = typeof finRaw   === "string"
            ? parseHeureToDate(finRaw)
            : new Date(finRaw);
          const loginMin  = h1.getHours() * 60 + h1.getMinutes();
          const logoutMin = h2.getHours() * 60 + h2.getMinutes();
          let diff = logoutMin - loginMin;
          diff = Math.round(diff / 5) * 5;           // arrondi au multiple de 5 minutes
          heures = diff / 60;
          heures = Math.round(heures * 100) / 100;  // 2 décimales
          nbH = heures.toFixed(2).replace(".", ",");
          rangeNbH.setValue(nbH);
        } catch (e) {
          Logger.log(`⛔ Erreur calcul heures ligne ${rowNum} : ${e}`);
          rangeNbH.setValue("#ERREUR");
          return;
        }
      }
    } else {
      // Si déjà renseigné, on parse
      heures = parseFloat(nbH.toString().replace(",", "."));
    }

    // ── 6. Calcul de "Nb H Valid" (heures - 0.25h)
    let nbHValid = Math.max(heures - 0.25, 0);
    nbHValid = Math.round(nbHValid * 100) / 100;
    rangeNbHValid.setValue(nbHValid.toFixed(2).replace(".", ","));

    // ── 7. Mise en forme conditionnelle
    if (heures < 2) {
      rangeNbH.setBackground("#f4cccc");
      rangeNbHValid.setBackground("#f4cccc");
    } else {
      rangeNbH.setBackground(null);
      rangeNbHValid.setBackground(null);
    }
  });

  // ─── MARQUAGE DES DOUBLONS (>2 entrées par jour) ─────────────────────────────
  Object.keys(doublons).forEach(key => {
    if (doublons[key] > 2) {
      const [nomDoublon, dateStr] = key.split("_");
      rows.forEach((ligne, i) => {
        const rowNum  = i + 2;
        const nomL    = (ligne[idx.nomComplet] || "")
          .toString().trim().replace(/\s+/g, " ");
        // Reconstruire date pour comparaison
        let dateL = ligne[idx.date];
        if (!(dateL instanceof Date)) {
          const [jj, mm, aaaa] = dateL.split("/");
          dateL = new Date(parseInt(aaaa,10), parseInt(mm,10)-1, parseInt(jj,10));
        }
        if (
          nomL === nomDoublon &&
          dateL.toDateString() === dateStr &&
          Math.floor((todayMid - new Date(dateL.getFullYear(), dateL.getMonth(), dateL.getDate()))/msPerDay) <= daysRange
        ) {
          feuille.getRange(rowNum, idx.nomComplet + 1).setBackground("#f4cccc");
        }
      });
    }
  });

  // ─── FORMAT DES CELLULES NUMÉRIQUES ───────────────────────────────────────────
  feuille.getRange(2, idx.nbH + 1, nRows).setNumberFormat("0.00");
  feuille.getRange(2, idxNbHValid + 1, nRows).setNumberFormat("0.00");

  Logger.log(`✅ Nettoyage des heures terminé (plage ${daysRange} jour(s)).`);
}

function recapRH() {
  const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAAAyyCjA8U/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=QDRXYj9w4NvrJmAGTTmLVZroKeiaIX1zbSMuRwETphE"; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPresences = ss.getSheetByName("Présences");
  const sheetPlanning = ss.getSheetByName("Planning 07/25");

  const dataPresences = sheetPresences.getDataRange().getValues();
  const dataPlanning = sheetPlanning.getDataRange().getValues();

  const idxNomP = dataPresences[0].indexOf("Nom Complet");
  const idxDateP = dataPresences[0].indexOf("Date");

  const idxDatePlan = dataPlanning[0].indexOf("Date");
  const idxNomPlan = dataPlanning[0].indexOf("Nom Complet");
  const idxMatinDebut = dataPlanning[0].indexOf("Shift Matin Début");
  const idxApremDebut = dataPlanning[0].indexOf("Shift Après-midi Début");
  const idxRemarques = dataPlanning[0].indexOf("Remarques");

  const today = new Date();
  const hier = new Date(today);
  hier.setDate(today.getDate() - 1);
  const hierStr = Utilities.formatDate(hier, "Europe/Paris", "dd/MM/yyyy");

  const prevusHier = new Map(); // nom => remarque

  // 📌 1. Qui était prévu hier (et leur remarque éventuelle)
  for (let i = 1; i < dataPlanning.length; i++) {
    const dateCell = dataPlanning[i][idxDatePlan];
    const dateStr = Utilities.formatDate(new Date(dateCell), "Europe/Paris", "dd/MM/yyyy");
    if (dateStr !== hierStr) continue;

    const nom = dataPlanning[i][idxNomPlan]?.toString().trim().toUpperCase();
    const shiftMatin = dataPlanning[i][idxMatinDebut];
    const shiftAprem = dataPlanning[i][idxApremDebut];
    const remarque = dataPlanning[i][idxRemarques]?.toString().trim().toLowerCase() || "";

    if (nom && (shiftMatin || shiftAprem)) {
      prevusHier.set(nom, remarque);
    }
  }

  // 📌 2. Qui a pointé hier
  const presentsHier = new Set();
  for (let i = 1; i < dataPresences.length; i++) {
    const nom = dataPresences[i][idxNomP]?.toString().trim().toUpperCase();
    const date = new Date(dataPresences[i][idxDateP]);
    const dateStr = Utilities.formatDate(date, "Europe/Paris", "dd/MM/yyyy");
    if (dateStr === hierStr && nom) {
      presentsHier.add(nom);
    }
  }

  // 📌 3. Classer les cas : absents / formation / congé révision
  const absents = [];
  const formations = [];
  const revisions = [];

  for (const [nom, remarque] of prevusHier.entries()) {
    if (presentsHier.has(nom)) continue;

    if (remarque.includes("formation")) {
      formations.push(nom);
    } else if (remarque.includes("congé révision")) {
      revisions.push(nom);
    } else {
      absents.push(nom);
    }
  }

  // 📌 4. Construire le message
  const dateStr = Utilities.formatDate(today, "Europe/Paris", "dd/MM/yyyy");
  let message = `📊 *RÉCAP RH - ${dateStr}*\n\n`;

  if (absents.length > 0) {
    message += `📍 *Liste des absents hier :*\n\n`;
    absents.forEach(nom => message += `- ${nom}\n`);
    message += `\n`;
  }

  if (formations.length > 0) {
    message += `🎓 *En formation hier :*\n\n`;
    formations.forEach(nom => message += `- ${nom}\n`);
    message += `\n`;
  }

  if (revisions.length > 0) {
    message += `📘 *En congé révision hier :*\n\n`;
    revisions.forEach(nom => message += `- ${nom}\n`);
    message += `\n`;
  }

  if (absents.length === 0 && formations.length === 0 && revisions.length === 0) {
    message += `✅ Aucun absent hier à signaler.`;
  }

  // 📤 5. Envoyer le message
  UrlFetchApp.fetch(WEBHOOK_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ text: message })
  });

  Logger.log(`✅ Message envoyé. Absents : ${absents.length} | Formations : ${formations.length} | Révisions : ${revisions.length}`);
}

function recapHebdoRH() {
  const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAAAyyCjA8U/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=QDRXYj9w4NvrJmAGTTmLVZroKeiaIX1zbSMuRwETphE";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPresences = ss.getSheetByName("Présences");
  const sheetPlanning = ss.getSheetByName("Planning 07/25");

  const data = sheetPresences.getDataRange().getValues();
  const planning = sheetPlanning.getDataRange().getValues();

  const headers = data[0];
  const idxDate = headers.indexOf("Date");
  const idxNom = headers.indexOf("Nom Complet");
  const idxHeures = headers.indexOf("Nb heures");

  const idxDatePlan = planning[0].indexOf("Date");
  const idxNomPlan = planning[0].indexOf("Nom Complet");
  const idxMatin = planning[0].indexOf("Shift Matin Début");
  const idxAprem = planning[0].indexOf("Shift Après-midi Début");

  const today = new Date();
  const jour = today.getDay(); // 1 = lundi, 5 = vendredi
  if (jour !== 5) {
    Logger.log("⏳ Aujourd’hui n’est pas vendredi. Pas de récap envoyé.");
    return;
  }

  const vendredi = new Date(today);
  const lundi = new Date(today);
  lundi.setDate(vendredi.getDate() - 4); // lundi de la semaine en cours

  // 📌 Obtenir les personnes prévues au planning du lundi au vendredi
  const personnesPrevues = new Set();
  for (let i = 1; i < planning.length; i++) {
    const datePlan = new Date(planning[i][idxDatePlan]);
    if (datePlan < lundi || datePlan > vendredi) continue;

    const nom = planning[i][idxNomPlan]?.toString().trim().toUpperCase();
    const matin = planning[i][idxMatin];
    const aprem = planning[i][idxAprem];

    if (nom && (matin || aprem)) {
      personnesPrevues.add(nom);
    }
  }

  const recap = {};

  for (let i = 1; i < data.length; i++) {
    const nom = data[i][idxNom]?.toString().trim().toUpperCase();
    const date = new Date(data[i][idxDate]);
    const heures = parseFloat(data[i][idxHeures]?.toString().replace(",", "."));

    if (!nom || isNaN(heures)) continue;
    if (!personnesPrevues.has(nom)) continue;
    if (date >= lundi && date <= vendredi) {
      if (!recap[nom]) recap[nom] = 0;
      recap[nom] += heures;
    }
  }

  if (Object.keys(recap).length === 0) {
    Logger.log("⚠️ Aucun récap hebdo à envoyer.");
    return;
  }

  let message = `📈 *RÉCAP SEMAINE* du ${Utilities.formatDate(lundi, "Europe/Paris", "dd/MM")} au ${Utilities.formatDate(vendredi, "Europe/Paris", "dd/MM/yyyy")}*\n\n`;

  for (const [nom, total] of Object.entries(recap)) {
    message += `👤 *${nom}* : ⏱️ ${total.toFixed(2)} heures\n`;
  }

  message += `\n🔁 Bon suivi !`;

  UrlFetchApp.fetch(WEBHOOK_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ text: message })
  });

  Logger.log("✅ Message hebdomadaire (lundi–vendredi) envoyé.");
}

function recapPersoFinSemaine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPresences = ss.getSheetByName("Présences");
  const sheetPlanning = ss.getSheetByName("Planning 07/25");
  const sheetConfig = ss.getSheetByName("config");

  const dataPresences = sheetPresences.getDataRange().getValues();
  const dataPlanning = sheetPlanning.getDataRange().getValues();
  const config = sheetConfig.getDataRange().getValues();

  const idxDateP = dataPresences[0].indexOf("Date");
  const idxNomP = dataPresences[0].indexOf("Nom Complet");
  const idxHeures = dataPresences[0].indexOf("Nb heures");

  const idxDatePlan = dataPlanning[0].indexOf("Date");
  const idxNomPlan = dataPlanning[0].indexOf("Nom Complet");
  const idxMatin = dataPlanning[0].indexOf("Shift Matin Début");
  const idxAprem = dataPlanning[0].indexOf("Shift Après-midi Début");

  const WEBHOOK = {};
  const personnesCiblees = [];

  for (let i = 2; i < config.length; i++) {
    const nom = config[i][0]?.toString().trim().toUpperCase();
    const webhook = config[i][1]?.toString().trim();
    if (nom && webhook) {
      WEBHOOK[nom] = webhook;
      personnesCiblees.push(nom);
    }
  }

  const today = new Date();
  const lundi = new Date(today);
  lundi.setDate(today.getDate() - today.getDay() + 1);
  const vendredi = new Date(lundi);
  vendredi.setDate(lundi.getDate() + 4);

  const planifies = new Set();

  for (let i = 1; i < dataPlanning.length; i++) {
    const row = dataPlanning[i];
    const date = new Date(row[idxDatePlan]);
    if (date < lundi || date > vendredi) continue;

    const jour = date.getDay();
    if (jour === 0 || jour === 6) continue;

    const nom = row[idxNomPlan]?.toString().trim().toUpperCase();
    const matin = row[idxMatin];
    const aprem = row[idxAprem];

    if (nom && (matin || aprem)) {
      planifies.add(`${Utilities.formatDate(date, "Europe/Paris", "yyyy-MM-dd")}|${nom}`);
    }
  }

  const recap = {};

  for (let i = 1; i < dataPresences.length; i++) {
    const nom = dataPresences[i][idxNomP]?.toString().trim().toUpperCase();
    const date = new Date(dataPresences[i][idxDateP]);
    const heures = parseFloat(dataPresences[i][idxHeures]);

    if (!nom || isNaN(heures)) continue;
    if (!personnesCiblees.includes(nom)) continue;

    const key = `${Utilities.formatDate(date, "Europe/Paris", "yyyy-MM-dd")}|${nom}`;
    if (planifies.has(key)) {
      recap[nom] = (recap[nom] || 0) + heures;
    }
  }

  for (const [nom, total] of Object.entries(recap)) {
    const webhook = WEBHOOK[nom];
    if (webhook) {
      const message = `📊 *RÉCAP DE TA SEMAINE - Du ${Utilities.formatDate(lundi, "Europe/Paris", "dd/MM")} au ${Utilities.formatDate(vendredi, "Europe/Paris", "dd/MM/yyyy")}*\n\n` +
                      `👤 *${nom}*, tu as effectué ⏱️ ${total.toFixed(2)} heures.\n\n` +
                      `✅ Bon travail !`;
      UrlFetchApp.fetch(webhook, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ text: message })
      });
    }
  }
}

function recapPersoMois() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const presences = ss.getSheetByName("Présences").getDataRange().getValues();
  const planning = ss.getSheetByName("Planning 07/25").getDataRange().getValues();
  const config = ss.getSheetByName("config").getDataRange().getValues();

  const idxDateP = presences[0].indexOf("Date");
  const idxNomP = presences[0].indexOf("Nom Complet");
  const idxHeures = presences[0].indexOf("Nb heures");

  const idxDatePlan = planning[0].indexOf("Date");
  const idxNomPlan = planning[0].indexOf("Nom Complet");
  const idxMatin = planning[0].indexOf("Shift Matin Début");
  const idxAprem = planning[0].indexOf("Shift Après-midi Début");

  const today = new Date();
  const mois = today.getMonth();
  const annee = today.getFullYear();
  const debut = new Date(annee, mois, 1);
  const fin = new Date(annee, mois + 1, 0);

  const moisFr = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
  const moisStr = `${moisFr[mois]} ${annee}`;

  const WEBHOOK = {};
  const personnesCiblees = [];

  for (let i = 2; i < config.length; i++) {
    const nom = config[i][0]?.toString().trim().toUpperCase();
    const webhook = config[i][1]?.toString().trim();
    if (nom && webhook) {
      WEBHOOK[nom] = webhook;
      personnesCiblees.push(nom);
    }
  }

  const planifies = new Set();
  for (let i = 1; i < planning.length; i++) {
    const row = planning[i];
    const date = new Date(row[idxDatePlan]);
    const jour = date.getDay();
    if (date < debut || date > fin || jour === 0 || jour === 6) continue;

    const nom = row[idxNomPlan]?.toString().trim().toUpperCase();
    const matin = row[idxMatin];
    const aprem = row[idxAprem];
    if (nom && (matin || aprem)) {
      planifies.add(`${Utilities.formatDate(date, "Europe/Paris", "yyyy-MM-dd")}|${nom}`);
    }
  }

  const recap = {};
  for (let i = 1; i < presences.length; i++) {
    const nom = presences[i][idxNomP]?.toString().trim().toUpperCase();
    const date = new Date(presences[i][idxDateP]);
    const heures = parseFloat(presences[i][idxHeures]);

    if (!nom || isNaN(heures)) continue;
    if (!personnesCiblees.includes(nom)) continue;

    const key = `${Utilities.formatDate(date, "Europe/Paris", "yyyy-MM-dd")}|${nom}`;
    if (planifies.has(key)) {
      recap[nom] = (recap[nom] || 0) + heures;
    }
  }

  for (const [nom, total] of Object.entries(recap)) {
    const webhook = WEBHOOK[nom];
    if (!webhook) continue;
    const message = `📊 *RÉCAP DE TON MOIS - ${moisStr}*\n\n` +
                    `👤 *${nom}*, tu as effectué ⏱️ ${total.toFixed(2)} heures ce mois-ci.\n\n` +
                    `✅ Bravo pour ton implication !`;
    UrlFetchApp.fetch(webhook, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ text: message })
    });
  }
}

function recapGlobalVeille() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPresences = ss.getSheetByName("Présences");
  const sheetPlanning = ss.getSheetByName("Planning 07/25");
  const config = ss.getSheetByName("config");

  const dataPresences = sheetPresences.getDataRange().getValues();
  const dataPlanning = sheetPlanning.getDataRange().getValues();
  const configData = config.getDataRange().getValues();

  const idxNom = dataPresences[0].indexOf("Nom Complet");
  const idxDate = dataPresences[0].indexOf("Date");
  const idxHeures = dataPresences[0].indexOf("Nb heures");

  const idxDatePlan = dataPlanning[0].indexOf("Date");
  const idxNomPlan = dataPlanning[0].indexOf("Nom Complet");
  const idxMatin = dataPlanning[0].indexOf("Shift Matin Début");
  const idxAprem = dataPlanning[0].indexOf("Shift Après-midi Début");

  const WEBHOOK = {};
  for (let i = 2; i < configData.length; i++) {
    const nom = configData[i][0]?.toString().trim().toUpperCase();
    const webhook = configData[i][1]?.toString().trim();
    if (nom && webhook) WEBHOOK[nom] = webhook;
  }

  const recap = {};
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const jour = yesterday.getDay();
  if (jour === 0 || jour === 6) return;

  const planifies = new Set();
  for (let i = 1; i < dataPlanning.length; i++) {
    const row = dataPlanning[i];
    const date = new Date(row[idxDatePlan]);
    if (date.toDateString() !== yesterday.toDateString()) continue;

    const nom = row[idxNomPlan]?.toString().trim().toUpperCase();
    const matin = row[idxMatin];
    const aprem = row[idxAprem];
    if (nom && (matin || aprem)) {
      planifies.add(nom);
    }
  }

  for (let i = 1; i < dataPresences.length; i++) {
    const nom = dataPresences[i][idxNom]?.toString().trim().toUpperCase();
    const date = new Date(dataPresences[i][idxDate]);
    const heures = parseFloat(dataPresences[i][idxHeures]?.toString().replace(",", "."));
    if (!nom || isNaN(heures)) continue;
    if (planifies.has(nom) && date.toDateString() === yesterday.toDateString()) {
      recap[nom] = (recap[nom] || 0) + heures;
    }
  }

  for (const [nom, total] of Object.entries(recap)) {
    const webhook = WEBHOOK[nom];
    if (!webhook) continue;
    const msg = `📊 *RÉCAP D'HIER - ${Utilities.formatDate(yesterday, "Europe/Paris", "dd/MM/yyyy")}*\n\n` +
                `👤 *${nom}* : ⏱️ ${total.toFixed(2)} h`;
    UrlFetchApp.fetch(webhook, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ text: msg })
    });
  }
}

/**
 * Récap global semaine - envoie à chacun son total de la semaine
 */
function recapGlobalSemaine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPresences = ss.getSheetByName("Présences");
  const sheetPlanning = ss.getSheetByName("Planning 07/25");
  const config = ss.getSheetByName("config");

  const dataPresences = sheetPresences.getDataRange().getValues();
  const dataPlanning = sheetPlanning.getDataRange().getValues();
  const configData = config.getDataRange().getValues();

  const idxDateP = dataPresences[0].indexOf("Date");
  const idxNomP = dataPresences[0].indexOf("Nom Complet");
  const idxHeures = dataPresences[0].indexOf("Nb heures");

  const idxDatePlan = dataPlanning[0].indexOf("Date");
  const idxNomPlan = dataPlanning[0].indexOf("Nom Complet");

  const WEBHOOK = {};
  for (let i = 2; i < configData.length; i++) {
    const nom = configData[i][0]?.toString().trim().toUpperCase();
    const webhook = configData[i][1]?.toString().trim();
    if (nom && webhook) WEBHOOK[nom] = webhook;
  }

  const today = new Date();
  const lundi = new Date(today);
  lundi.setDate(today.getDate() - today.getDay() + 1);
  const vendredi = new Date(lundi);
  vendredi.setDate(lundi.getDate() + 4);

  const prevus = new Set();
  for (let i = 1; i < dataPlanning.length; i++) {
    const nom = dataPlanning[i][idxNomPlan]?.toString().trim().toUpperCase();
    const date = new Date(dataPlanning[i][idxDatePlan]);
    const jour = date.getDay();
    if (!nom || jour === 0 || jour === 6) continue;
    if (date >= lundi && date <= vendredi) {
      prevus.add(nom);
    }
  }

  const recap = {};
  for (let i = 1; i < dataPresences.length; i++) {
    const nom = dataPresences[i][idxNomP]?.toString().trim().toUpperCase();
    const date = new Date(dataPresences[i][idxDateP]);
    const jour = date.getDay();
    const heures = parseFloat(dataPresences[i][idxHeures]);

    if (!nom || isNaN(heures)) continue;
    if (jour === 0 || jour === 6) continue;
    if (date >= lundi && date <= vendredi && prevus.has(nom)) {
      recap[nom] = (recap[nom] || 0) + heures;
    }
  }

  for (const [nom, total] of Object.entries(recap)) {
    const webhook = WEBHOOK[nom];
    if (!webhook) continue;
    const msg = `📆 *RÉCAP DE TA SEMAINE EN COURS - Jusqu'au ${Utilities.formatDate(today, "Europe/Paris", "dd/MM/yyyy")}*\n\n` +
                `👤 *${nom}* : ⏱️ ${total.toFixed(2)} heures.`;
    UrlFetchApp.fetch(webhook, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ text: msg })
    });
    Logger.log(`✅ Message envoyé à ${nom}`);
  }
}

/**
 * Récap global mois - envoie à chacun son total du mois
 */
function recapGlobalMois() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPresences = ss.getSheetByName("Présences");
  const sheetPlanning = ss.getSheetByName("Planning 07/25");
  const config = ss.getSheetByName("config");

  const dataPresences = sheetPresences.getDataRange().getValues();
  const dataPlanning = sheetPlanning.getDataRange().getValues();
  const configData = config.getDataRange().getValues();

  const idxDateP = dataPresences[0].indexOf("Date");
  const idxNomP = dataPresences[0].indexOf("Nom Complet");
  const idxHeures = dataPresences[0].indexOf("Nb heures");

  const idxDatePlan = dataPlanning[0].indexOf("Date");
  const idxNomPlan = dataPlanning[0].indexOf("Nom Complet");

  const WEBHOOK = {};
  for (let i = 2; i < configData.length; i++) {
    const nom = configData[i][0]?.toString().trim().toUpperCase();
    const webhook = configData[i][1]?.toString().trim();
    if (nom && webhook) WEBHOOK[nom] = webhook;
  }

  const today = new Date();
  const debutMois = new Date(today.getFullYear(), today.getMonth(), 1);
  const finMois = new Date(today.getFullYear(), today.getMonth() + 1, 0);

  const moisFrancais = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
  const moisStr = `${moisFrancais[today.getMonth()]} ${today.getFullYear()}`;

  const prevus = new Set();
  for (let i = 1; i < dataPlanning.length; i++) {
    const nom = dataPlanning[i][idxNomPlan]?.toString().trim().toUpperCase();
    const date = new Date(dataPlanning[i][idxDatePlan]);
    const jour = date.getDay();
    if (!nom || jour === 0 || jour === 6) continue;
    if (date >= debutMois && date <= finMois) {
      prevus.add(nom);
    }
  }

  const recap = {};
  for (let i = 1; i < dataPresences.length; i++) {
    const nom = dataPresences[i][idxNomP]?.toString().trim().toUpperCase();
    const date = new Date(dataPresences[i][idxDateP]);
    const jour = date.getDay();
    const heures = parseFloat(dataPresences[i][idxHeures]);

    if (!nom || isNaN(heures)) continue;
    if (jour === 0 || jour === 6) continue;
    if (date >= debutMois && date <= finMois && prevus.has(nom)) {
      recap[nom] = (recap[nom] || 0) + heures;
    }
  }

  for (const [nom, total] of Object.entries(recap)) {
    const webhook = WEBHOOK[nom];
    if (!webhook) continue;
    const msg = `📊 *RÉCAP DU MOIS - ${moisStr}*\n\n` +
                `👤 *${nom}* : ⏱️ ${total.toFixed(2)} heures.`;
    UrlFetchApp.fetch(webhook, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ text: msg })
    });
    Logger.log(`✅ Message envoyé à ${nom}`);
  }
}

function envoyerTotalJournalierComptable() {
  const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAAAeK9bLcs/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=tHWDncoWNnqAMx6IfaryjA6M19PWVqRCEkj0RCWaG-k";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPresences = ss.getSheetByName("Présences");
  const sheetPlanning = ss.getSheetByName("Planning 07/25");

  const dataPresences = sheetPresences.getDataRange().getValues();
  const dataPlanning = sheetPlanning.getDataRange().getValues();

  const idxDateP = dataPresences[0].indexOf("Date");
  const idxNomP = dataPresences[0].indexOf("Nom Complet");
  const idxHeures = dataPresences[0].indexOf("Nb heures");

  const idxDatePlan = dataPlanning[0].indexOf("Date");
  const idxNomPlan = dataPlanning[0].indexOf("Nom Complet");

  const today = new Date();
  const dateStr = Utilities.formatDate(today, "Europe/Paris", "dd/MM/yyyy");

  const prevus = new Set();
  for (let i = 1; i < dataPlanning.length; i++) {
    const date = new Date(dataPlanning[i][idxDatePlan]);
    const nom = dataPlanning[i][idxNomPlan]?.toString().trim().toUpperCase();
    const jour = date.getDay();
    const planStr = Utilities.formatDate(date, "Europe/Paris", "dd/MM/yyyy");
    if (planStr === dateStr && nom && jour !== 0 && jour !== 6) {
      prevus.add(nom);
    }
  }

  const recap = {};
  let totalGeneral = 0;

  for (let i = 1; i < dataPresences.length; i++) {
    const date = new Date(dataPresences[i][idxDateP]);
    const nom = dataPresences[i][idxNomP]?.toString().trim().toUpperCase();
    let heures = parseFloat(dataPresences[i][idxHeures]);
    const jour = date.getDay();

    const presStr = Utilities.formatDate(date, "Europe/Paris", "dd/MM/yyyy");
    if (presStr !== dateStr || !nom || isNaN(heures) || jour === 0 || jour === 6) continue;
    if (!prevus.has(nom)) continue;

    recap[nom] = (recap[nom] || 0) + heures;
    totalGeneral += heures;
  }

  let message = `📘 *Total des heures travaillées - ${dateStr}*\n\n`;
  for (const nom in recap) {
    message += `👤 ${nom} : ${recap[nom].toFixed(2)} h\n`;
  }
  message += `\n📊 *Total général* : ${totalGeneral.toFixed(2)} h`;

  UrlFetchApp.fetch(WEBHOOK_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ text: message })
  });

  Logger.log("✅ Message total journalier comptable envoyé au webhook URL uniquement.");
}

function createTriggers() {
  // Supprimer les anciens triggers si besoin (optionnel)
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    ScriptApp.deleteTrigger(trigger);
  }

  ScriptApp.newTrigger('recapRH')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();

  ScriptApp.newTrigger('recapHebdoRH')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();

ScriptApp.newTrigger('recapPersoFinSemaine')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(19)
    .create();

ScriptApp.newTrigger('recapPersoMois')
    .timeBased()
    .onMonthDay(1)
    .atHour(10)
    .create();

ScriptApp.newTrigger('recapGlobalVeille')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();

ScriptApp.newTrigger('recapGlobalSemaine')
  .timeBased()
  .everyWeeks(1)
  .onWeekDay(ScriptApp.WeekDay.FRIDAY)
  .atHour(19)
  .create();

ScriptApp.newTrigger('recapGlobalMois')
    .timeBased()
    .nearMinute(0)
    .atHour(19)
    .onMonthDay(getLastDayOfMonth(new Date()))
    .create();

ScriptApp.newTrigger('envoyerTotalJournalierComptable')
    .timeBased()
    .everyDays(1)
    .atHour(19)
    .create();

function getLastDayOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
}
}

/**
 * Récupère tous les triggers du projet et les loggue dans le journal
 */
function logAllTriggers() {
  // Récupérer la liste des triggers installés sur ce projet
  const triggers = ScriptApp.getProjectTriggers();
  
  if (triggers.length === 0) {
    Logger.log("⚠️ Aucun déclencheur trouvé dans ce projet.");
    return;
  }
  
  // Parcourir et logguer pour chaque trigger :
  triggers.forEach(trigger => {
    // getHandlerFunction() : nom de la fonction appelée
    // getTriggerSource()   : origine du trigger (TIME, SPREADSHEETS, etc.)
    // getEventType()       : type d’événement (ON_OPEN, ON_EDIT, etc.)
    // getUniqueId()        : identifiant unique du trigger
    Logger.log(
      `• Fonction : ${trigger.getHandlerFunction()}  |  ` +
      `Source : ${trigger.getTriggerSource()}  |  ` +
      `Événement : ${trigger.getEventType()}  |  ` +
      `ID : ${trigger.getUniqueId()}`
    );
  });
}

function logInstallableTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    Logger.log("⚠️ Aucun trigger installable trouvé.");
    return;
  }
  triggers.forEach(t => {
    Logger.log(
      `• Fonction : ${t.getHandlerFunction()} | ` +
      `Source : ${t.getTriggerSource()} | ` +
      `Événement : ${t.getEventType()}`
    );
  });
}

function logContainerTriggers() {
  // Récupère les triggers que VOUS avez installés pour cette feuille
  const container = SpreadsheetApp.getActive();
  const triggers  = ScriptApp.getUserTriggers(container);
  if (triggers.length === 0) {
    Logger.log("⚠️ Aucun trigger installable lié à cette feuille.");
    return;
  }
  triggers.forEach(t => {
    Logger.log(
      `• Fonction : ${t.getHandlerFunction()} | ` +
      `Source : ${t.getTriggerSource()} | ` +
      `Événement : ${t.getEventType()}`
    );
  });
}

/**
 * Affiche dans le journal :
 *  • Tous les triggers “installables” du projet
 *  • Tous les triggers associés à cette feuille de calcul (container-bound)
 */
function logAllTriggersDetailed() {
  // 1️⃣ Triggers du projet
  const projectTriggers = ScriptApp.getProjectTriggers();
  Logger.log("=== Triggers du projet Apps Script ===");
  if (projectTriggers.length === 0) {
    Logger.log("⚠️ Aucun trigger installable trouvé dans le projet.");
  } else {
    projectTriggers.forEach(t => {
      Logger.log(
        `Fonction : ${t.getHandlerFunction()}  |  ` +
        `Source : ${t.getTriggerSource()}  |  ` +
        `Événement : ${t.getEventType()}`
      );
    });
  }

  // 2️⃣ Triggers “container-bound” (sur la feuille active)
  const container = SpreadsheetApp.getActiveSpreadsheet();
  const userTriggers = ScriptApp.getUserTriggers(container);
  Logger.log("=== Triggers container-bound (cette feuille) ===");
  if (userTriggers.length === 0) {
    Logger.log("⚠️ Aucun trigger installé sur cette feuille par un autre utilisateur ou add-on.");
  } else {
    userTriggers.forEach(t => {
      Logger.log(
        `Fonction : ${t.getHandlerFunction()}  |  ` +
        `Source : ${t.getTriggerSource()}  |  ` +
        `Événement : ${t.getEventType()}`
      );
    });
  }
}

function deleteTriggersForMyFunction() {
  const functionName = 'verifierPlanningEtNotifsAvecHeure'; // nom de la fonction dont on veut supprimer les triggers
  const allTriggers = ScriptApp.getProjectTriggers();
  let count = 0;
  
  allTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });
  
  Logger.log(`Supprimé ${count} déclencheur(s) pour la fonction "${functionName}".`);
}

function genererReportingMoisPrecedent() {
  const today = new Date();

  // Mois précédent
  let mois = today.getMonth(); // Janvier = 0
  let annee = today.getFullYear();

  if (mois === 0) {
    mois = 12;
    annee -= 1;
  }

  genererReportingMensuelKiosksign(mois, annee);
}

function menuGenererReportingMensuel() {
  const ui = SpreadsheetApp.getUi();
  const today = new Date();

  const mois = Number(
    ui.prompt("Mois", "Entrez le mois (1-12)", ui.ButtonSet.OK_CANCEL).getResponseText()
  );
  const annee = Number(
    ui.prompt("Année", "Entrez l’année (ex: 2025)", ui.ButtonSet.OK_CANCEL).getResponseText()
  );

  if (!mois || !annee) {
    ui.alert("❌ Mois ou année invalide.");
    return;
  }

  genererReportingMensuelKiosksign(mois, annee);
}


/*
function genererReportingMensuelKiosksign(mois, annee) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const mm = Utilities.formatString("%02d", mois);
  const yy = annee.toString().slice(-2);

  const planSheetName = `Planning ${mm}/${yy}`;
  const planSheet = ss.getSheetByName(planSheetName);

  // ⚠️ Drapeau absence planning
  const planningAbsent = !planSheet;

  if (planningAbsent) {
    Logger.log(`⚠️ Feuille "${planSheetName}" absente : reporting généré sans données planning.`);
  }

  // Charger les données (le loader gérera le cas sans planning)
  const data = chargerDonneesMensuelles_(mois, annee, planningAbsent);

  creerOuMaj_RECAP_EVP_(data, mm, yy);
  creerOuMaj_COMPARATIF_(data, mm, yy);
  creerOuMaj_ANOMALIES_(data, mm, yy);

  // Message utilisateur clair
  const msg = planningAbsent
    ? `⚠️ Reporting Kiosksign ${mm}${yy} généré SANS planning.\nLes écarts planning ne sont pas calculés.`
    : `✅ Reporting Kiosksign ${mm}${yy} généré avec succès.`;

  SpreadsheetApp.getUi().alert(msg);
}
*/

function chargerDonneesMensuelles_(mois, annee) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const mm = Utilities.formatString("%02d", mois);
  const yy = annee.toString().slice(-2);

  const debut = new Date(annee, mois - 1, 1);
  const fin = new Date(annee, mois, 0);

  // =========================
  // JOURS OUVRÉS (lun-ven)
  // =========================
  let joursOuvres = 0;
  for (let d = new Date(debut); d <= fin; d.setDate(d.getDate() + 1)) {
    if (d.getDay() !== 0 && d.getDay() !== 6) joursOuvres++;
  }

  // =========================
  // COLLABORATEURS (HRMS)
  // =========================
  const HR_FILE_ID = "1Dn5HfoYTegM9nJR4LKfrqOJ65CgJ_1crtiNFFUd2LMU";
  const hrSS = SpreadsheetApp.openById(HR_FILE_ID);
  const hrSheet = hrSS.getSheetByName("liste collaborateurs");
  if (!hrSheet) throw new Error("❌ Feuille HRMS introuvable.");

  const hrValues = hrSheet.getDataRange().getValues();
  const hrHead = hrValues[0];
  const hrData = hrValues.slice(1);

  const idx = {
    nom: hrHead.indexOf("Nom Complet"),
    email: hrHead.indexOf("Email"),
    statut: hrHead.indexOf("Statut"),
    entree: hrHead.indexOf("Date entrée"),
    sortie: hrHead.indexOf("Date Sortie"),
    actif: hrHead.indexOf("Actif/Inactif")
  };

  const collaborateurs = {};

  hrData.forEach(r => {
    const nom = (r[idx.nom] || "").toString().trim().replace(/\s+/g, " ");
    if (!nom) return;

    const actif = (r[idx.actif] || "").toString().toLowerCase() === "actif";
    const entree = r[idx.entree] ? new Date(r[idx.entree]) : new Date("2000-01-01");
    const sortie = r[idx.sortie] ? new Date(r[idx.sortie]) : new Date("2100-01-01");

    if (!actif || entree > fin || sortie < debut) return;

    const statut = (r[idx.statut] || "").toLowerCase();
    const groupe =
      statut.includes("indépend") ? "independants" :
      statut.includes("stag") ? "stagiaires" :
      statut.includes("apprenti") || statut.includes("altern") ? "alternants" :
      "salaries";

    const objectifHebdo = (groupe === "stagiaires" || groupe === "alternants") ? 35 : 39;

    collaborateurs[nom] = {
      nom,
      email: r[idx.email],
      statut,
      groupe,
      objectifHebdo,
      jours: {},
      anomalies: []
    };
  });

  // =========================
  // PRÉSENCES
  // =========================
  const presSheet = ss.getSheetByName("Présences");
  const presValues = presSheet.getDataRange().getValues();
  const presHead = presValues[0];
  const presData = presValues.slice(1);

  const p = {
    date: presHead.indexOf("Date"),
    nom: presHead.indexOf("Nom Complet"),
    heures: presHead.indexOf("Nb heures"),
    sign: presHead.indexOf("Signature URL")
  };

  presData.forEach(r => {
    const rawDate = r[p.date];
    if (!rawDate) return;

    const date = new Date(rawDate);
    if (date < debut || date > fin) return;

    const nom = (r[p.nom] || "").toString().trim();
    if (!collaborateurs[nom]) return;

    const d = date.getDate();
    if (!collaborateurs[nom].jours[d]) {
      collaborateurs[nom].jours[d] = { heures: 0, shifts: 0, signature: false };
    }
    const jour = collaborateurs[nom].jours[d];

    jour.shifts++;
    jour.heures += Number((r[p.heures] || 0).toString().replace(",", ".")) || 0;
    if (r[p.sign]) jour.signature = true;
  });

  return { collaborateurs, joursOuvres, debut, fin };
}

function creerOuMaj_RECAP_EVP_(data, mm, yy) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomFeuille = `RECAP EVP Salariés ${mm}${yy}`;

  let sh = ss.getSheetByName(nomFeuille);
  if (!sh) sh = ss.insertSheet(nomFeuille);
  sh.clear();

  const collaborateurs = data.collaborateurs;
  const joursOuvres = data.joursOuvres;

  /* =========================
     CONSTANTES STRUCTURE
  ========================= */

  const NB_JOURS_MAX = 31;
  const COL_NOM = 1;
  const COL_JOURS_START = 2;
  const COL_TTM = COL_JOURS_START + NB_JOURS_MAX;
  const COL_JOM = COL_TTM + 1;
  const COL_DELTA = COL_TTM + 2;
  const NB_COLS_TOTAL = COL_DELTA;

  // Sécurisation du nombre de colonnes
  if (sh.getMaxColumns() < NB_COLS_TOTAL) {
    sh.insertColumnsAfter(
      sh.getMaxColumns(),
      NB_COLS_TOTAL - sh.getMaxColumns()
    );
  }

  /* =========================
     COULEURS
  ========================= */

  const COLORS = {
    header: "#d9e1f2",
    weekend: "#bfbfbf",
    orange: "#ffe0b2",
    redLight: "#f8d7da",
    redDark: "#7a1f1f",
    recap: "#fff2cc"
  };

  let row = 1;

  /* =========================
     BANDEAU HAUT
  ========================= */

  sh.getRange(row, 1).setValue("Mois").setFontWeight("bold").setBackground(COLORS.recap);
  sh.getRange(row, 2).setValue(`${mm}/${yy}`).setBackground(COLORS.recap);

  row++;
  sh.getRange(row, 1).setValue("Nb jours ouvrés").setFontWeight("bold");
  sh.getRange(row, 2).setValue(joursOuvres);

  row++;
  sh.getRange(row, 1).setValue("Nb heures ouvrées (35h)").setFontWeight("bold");
  sh.getRange(row, 2).setValue(joursOuvres * 7);

  row++;
  sh.getRange(row, 1).setValue("Nb heures ouvrées (39h)").setFontWeight("bold");
  sh.getRange(row, 2).setValue(joursOuvres * 7.8);

  /* =========================
     LÉGENDE ABSENCES
  ========================= */

  const legend = [
    ["AM", "Arrêt maladie"],
    ["AA", "Absence autorisée"],
    ["AI", "Absence injustifiée"],
    ["DFC", "Début / Fin contrat"],
    ["CP", "Congé payé"],
    ["CSS", "Congé sans solde"],
    ["REC", "Récupération"]
  ];

  sh.getRange(1, 6).setValue("Légende Absences").setFontWeight("bold");
  sh.getRange(2, 6, legend.length, 2).setValues(legend);
  sh.getRange(2, 6, legend.length, 2).setBorder(true, true, true, true, true, true);

  /* =========================
     SECTIONS EVP
  ========================= */

  row = 8;

  const sections = [
    { label: "Salariés", key: "salaries" },
    { label: "Indépendants", key: "independants" },
    { label: "Apprentis / Alternants", key: "alternants" },
    { label: "Stagiaires", key: "stagiaires" }
  ];

  sections.forEach(section => {

    // Titre section
    sh.getRange(row, 1).setValue(section.label)
      .setFontWeight("bold")
      .setBackground(COLORS.header);
    row++;

    // En-têtes
    sh.getRange(row, COL_NOM).setValue("Collaborateur").setFontWeight("bold");

    for (let d = 1; d <= NB_JOURS_MAX; d++) {
      sh.getRange(row, COL_JOURS_START + d - 1)
        .setValue(d)
        .setHorizontalAlignment("center")
        .setFontWeight("bold");
    }

    sh.getRange(row, COL_TTM).setValue("TT M").setFontWeight("bold");
    sh.getRange(row, COL_JOM).setValue("JO M").setFontWeight("bold");
    sh.getRange(row, COL_DELTA).setValue("Delta").setFontWeight("bold");

    row++;
    const debutBloc = row;

    Object.values(collaborateurs)
      .filter(c => c.groupe === section.key)
      .sort((a, b) => a.nom.localeCompare(b.nom))
      .forEach(c => {

        sh.getRange(row, COL_NOM).setValue(c.nom);

        const objectifMois = (c.objectifHebdo * joursOuvres) / 5;

        for (let d = 1; d <= NB_JOURS_MAX; d++) {
          const cell = sh.getRange(row, COL_JOURS_START + d - 1);
          const j = c.jours[d];
          if (!j) continue;

          // + de 2 shifts
          if (j.shifts > 2) {
            cell.setBackground(COLORS.orange)
                .setFontColor(COLORS.redDark)
                .setComment("Plus de 2 shifts enregistrés");
          }

          // absence signature
          if ((j.planE || j.planF) && !j.signature) {
            cell.setValue("AI").setBackground(COLORS.redLight);
          } else if (j.signature) {
            cell.setValue(j.heures || "");
          }
        }

        // Formules
        sh.getRange(row, COL_TTM)
          .setFormulaR1C1(`=SUM(RC${COL_JOURS_START}:RC${COL_TTM - 1})`);

        sh.getRange(row, COL_JOM)
          .setFormulaR1C1(`=COUNTIF(RC${COL_JOURS_START}:RC${COL_TTM - 1},">0")`);

        sh.getRange(row, COL_DELTA)
          .setFormula(`=${objectifMois}-INDIRECT("R"&ROW()&"C"&${COL_TTM},FALSE)`);

        row++;
      });

    // Ligne total
    sh.getRange(row, COL_NOM).setValue("Total").setFontWeight("bold");
    for (let c = COL_JOURS_START; c <= COL_DELTA; c++) {
      sh.getRange(row, c)
        .setFormulaR1C1(`=SUM(R${debutBloc}C${c}:R${row - 1}C${c})`);
    }

    row += 2;
  });

  sh.autoResizeColumns(1, NB_COLS_TOTAL);
}

/*
function creerOuMaj_COMPARATIF_(data, mm, yy) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomFeuille = `DETAIL COMPARATIF PLANNING ${mm}${yy}`;

  let sh = ss.getSheetByName(nomFeuille);
  if (!sh) sh = ss.insertSheet(nomFeuille);
  sh.clear();

  const headers = [
    "Collaborateur",
    "H plan E",
    "H plan F",
    "H Tot Plan",
    "H Réalisées",
    "H Absence",
    "Écart présence",
    "Écart MOIS"
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

  let row = 2;

  Object.values(data.collaborateurs)
    .sort((a, b) => a.nom.localeCompare(b.nom))
    .forEach(c => {

      let planE = 0, planF = 0, real = 0;

      Object.values(c.jours).forEach(j => {
        planE += j.planE || 0;
        planF += j.planF || 0;
        if (j.signature) real += j.heures || 0;
      });

      const objectif = c.objectifHebdo * data.joursOuvres / 5;

      sh.getRange(row, 1).setValue(c.nom);
      sh.getRange(row, 2).setValue(planE);
      sh.getRange(row, 3).setValue(planF);
      sh.getRange(row, 4).setFormulaR1C1("=RC[-2]+RC[-1]");
      sh.getRange(row, 5).setValue(real);
      sh.getRange(row, 6).setFormulaR1C1("=RC[-2]-RC[-1]");
      sh.getRange(row, 7).setFormulaR1C1("=RC[-2]-RC[-3]");
      sh.getRange(row, 8).setFormula(`=${objectif}-E${row}`);

      row++;
    });

  sh.autoResizeColumns(1, headers.length);
}


function creerOuMaj_ANOMALIES_(data, mm, yy) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomFeuille = `Anomalies Kiosksign ${mm}${yy}`;

  let sh = ss.getSheetByName(nomFeuille);
  if (!sh) sh = ss.insertSheet(nomFeuille);
  sh.clear();

  const headers = ["Collaborateur", "Email", "Statut", "Date", "Problème", "Détail"];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

  let row = 2;

  const ordre = { salaries: 1, independants: 2, alternants: 3, stagiaires: 4 };

  Object.values(data.collaborateurs)
    .sort((a, b) => ordre[a.groupe] - ordre[b.groupe] || a.nom.localeCompare(b.nom))
    .forEach(c => {
      c.anomalies.forEach(a => {
        sh.getRange(row, 1).setValue(c.nom);
        sh.getRange(row, 2).setValue(c.email);
        sh.getRange(row, 3).setValue(c.groupe);
        sh.getRange(row, 4).setValue(Utilities.formatDate(a.date, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy"));
        sh.getRange(row, 5).setValue(a.type);
        sh.getRange(row, 6).setValue(a.detail);
        row++;
      });
    });

  sh.autoResizeColumns(1, headers.length);
}
*/

function getISOWeek_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function clonerModeleRH_(mm, yy) {

  const TEMPLATE_FILE_ID = "1q4RHeaaWwZFsfzqp6bxnDblM3FuR8qdT3t5mADZ2gzg";
  const TEMPLATE_SHEET_NAME = "TEST EVP Janvier 26";

  const ssTarget = SpreadsheetApp.getActiveSpreadsheet();
  const ssTemplate = SpreadsheetApp.openById(TEMPLATE_FILE_ID);
  const shTemplate = ssTemplate.getSheetByName(TEMPLATE_SHEET_NAME);
  if (!shTemplate) throw new Error("❌ Template RH introuvable.");

  const name = `RECAP EVP Salariés ${mm}${yy}`;
  const existing = ssTarget.getSheetByName(name);
  if (existing) ssTarget.deleteSheet(existing);

  return shTemplate.copyTo(ssTarget).setName(name);
}

function creerOuMaj_RECAP_EVP_depuisTemplate_(data, mm, yy) {

  const sh = clonerModeleRH_(mm, yy);

  // Bandeau haut
  sh.getRange("B2").setValue(`${mm}/${yy}`);
  sh.getRange("B3").setValue(data.joursOuvres);
  sh.getRange("B4").setValue(data.joursOuvres * 7);
  sh.getRange("B5").setValue(data.joursOuvres * 7.8);

  // Sections EVP
  remplirSectionEVP_(sh, "Salariés", data, "salaries");
  remplirSectionEVP_(sh, "Indépendants", data, "independants");
  remplirSectionEVP_(sh, "Apprentis / Alternants", data, "alternants");
  remplirSectionEVP_(sh, "Stagiaires", data, "stagiaires");
}

function remplirSectionEVP_(sh, libelleSection, data, groupeKey) {

  const collaborateurs = Object.values(data.collaborateurs)
    .filter(c => c.groupe === groupeKey)
    .sort((a, b) => a.nom.localeCompare(b.nom));

  const finder = sh.createTextFinder(libelleSection).findNext();
  if (!finder) {
    Logger.log(`⚠️ Section "${libelleSection}" introuvable`);
    return;
  }

  // Ligne du titre de section
  const sectionRow = finder.getRow();

  // Ligne d’en-tête jours = juste en dessous
  const headerRow = sectionRow + 1;

  // Trouver la ligne "Total"
  const totalFinder = sh.createTextFinder("Total")
    .matchEntireCell(true)
    .findAll()
    .find(c => c.getRow() > headerRow);

  if (!totalFinder) {
    throw new Error(`❌ Ligne Total introuvable pour ${libelleSection}`);
  }

  const totalRow = totalFinder.getRow();
  const firstDataRow = headerRow + 1;

  // Nombre de lignes disponibles
  const lignesExistantes = totalRow - firstDataRow;
  const lignesNecessaires = collaborateurs.length;

  // Ajouter des lignes si nécessaire
  if (lignesNecessaires > lignesExistantes) {
    const nbToInsert = lignesNecessaires - lignesExistantes;
    sh.insertRowsBefore(totalRow, nbToInsert);

    // Copier le style de la première ligne collaborateur
    const sourceRange = sh.getRange(firstDataRow, 1, 1, sh.getLastColumn());
    const targetRange = sh.getRange(firstDataRow, 1, nbToInsert + 1, sh.getLastColumn());
    sourceRange.copyTo(targetRange, { contentsOnly: false });
  }

  // Nettoyage lignes en trop
  if (lignesNecessaires < lignesExistantes) {
    sh.getRange(firstDataRow + lignesNecessaires, 1, lignesExistantes - lignesNecessaires, sh.getLastColumn())
      .clearContent();
  }

  // Remplissage
  collaborateurs.forEach((c, idx) => {
    const row = firstDataRow + idx;
    sh.getRange(row, 1).setValue(c.nom);

    for (let d = 1; d <= 31; d++) {
      const cell = sh.getRange(row, d + 1);
      const j = c.jours[d];
      if (!j) continue;

      if (j.shifts > 2) {
        cell.setBackground("#ffe0b2")
            .setFontColor("#7a1f1f")
            .setNote("Plus de 2 shifts enregistrés");
      }

      if ((j.planE || j.planF) && !j.signature) {
        cell.setValue("AI").setBackground("#f8d7da");
      } else if (j.signature) {
        cell.setValue(j.heures);
      }
    }
  });
}

function genererReportingMensuelKiosksign(mois, annee) {
  const mm = Utilities.formatString("%02d", mois);
  const yy = annee.toString().slice(-2);
  const data = chargerDonneesMensuelles_(mois, annee);
  creerOuMaj_RECAP_EVP_depuisTemplate_(data, mm, yy);
  SpreadsheetApp.getUi().alert(`Reporting RH ${mm}/${yy} généré.`);
}

function genererReportMoisPrecedent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 📅 Calcul du mois précédent
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);

  const mm = Utilities.formatString("%02d", prev.getMonth() + 1);
  const yy = prev.getFullYear().toString().slice(-2);
  const monthKey = `${mm}/${yy}`;

  const reportingName = "Reporting " + monthKey;
  const planningName = "Planning " + monthKey;

  // 🗓️ Import du planning si absent
  if (!ss.getSheetByName(planningName)) {
    importPlanning(
      parseInt(mm, 10),
      prev.getFullYear(),
      monthKey,
      false
    );
  }

  // 📊 Création ou réinitialisation du reporting
  let reportingSheet = ss.getSheetByName(reportingName);
  if (!reportingSheet) {
    reportingSheet = ss.insertSheet(reportingName);
  } else {
    reportingSheet.clear();
  }

  // ⚙️ Génération du reporting
  generateMonthlyReport(reportingSheet, monthKey);

  // 📐 Placement des feuilles (cohérence avec ton standard)
  const sheets = ss.getSheets();
  ss.setActiveSheet(reportingSheet);
  ss.moveActiveSheet(sheets.length - 1);

  const planningSheet = ss.getSheetByName(planningName);
  if (planningSheet) {
    ss.setActiveSheet(planningSheet);
    ss.moveActiveSheet(ss.getSheets().length);
  }

  Logger.log(`✅ Reporting du mois précédent généré avec succès (${monthKey}).`);
}

function buildDailyRecapTable_(sheet, startRow, collaborateurs, realisations, startDate, endDate) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  // 🔹 Construire les semaines ISO avec jours ouvrés
  const weeksMap = {};
  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    const day = d.getDay(); // 1=lundi ... 5=vendredi
    if (day < 1 || day > 5) continue;

    const week = getWeekNumber(d);
    if (!weeksMap[week]) weeksMap[week] = { days: [] };
    weeksMap[week].days.push(new Date(d));
  }

  const weekNumbers = Object.keys(weeksMap).map(Number).sort((a, b) => a - b);

  // 🧾 EN-TÊTE
  const header = ["Collaborateur"];
  weekNumbers.forEach(w => {
    header.push("Lun", "Mar", "Mer", "Jeu", "Ven", `Total S${w}`);
  });
  header.push("Total Mois");

  sheet.getRange(startRow, 1, 1, header.length)
    .setValues([header])
    .setFontWeight("bold")
    .setBackground("#e8f0fe");

  let row = startRow + 1;

  // 📊 Accumulateurs TOTAL JOUR
  const totalJour = Array(header.length - 1).fill(0); // sans colonne "Collaborateur"

  // 🧑‍💼 LIGNES COLLABORATEURS
  for (const nom in collaborateurs) {
    const line = [nom];
    let totalMois = 0;
    let colIndex = 0;

    weekNumbers.forEach(w => {
      let totalWeek = 0;
      const days = weeksMap[w].days;

      const dayMap = {};
      days.forEach(d => {
        const key = Utilities.formatDate(d, "Europe/Paris", "yyyy-MM-dd");
        dayMap[d.getDay()] = realisations[nom]?.[key] || 0;
      });

      for (let dow = 1; dow <= 5; dow++) {
        const h = roundToQuarterHour(dayMap[dow] || 0);
        line.push(h);
        totalWeek += h;
        totalMois += h;

        totalJour[colIndex] += h;
        colIndex++;
      }

      line.push(roundToQuarterHour(totalWeek));
      totalJour[colIndex] += totalWeek;
      colIndex++;
    });

    line.push(roundToQuarterHour(totalMois));
    totalJour[colIndex] += totalMois;

    sheet.getRange(row++, 1, 1, header.length).setValues([line]);
  }

  // ➕ LIGNE TOTAL JOUR
  const totalLine = ["TOTAL JOUR", ...totalJour.map(v => roundToQuarterHour(v))];
  sheet.getRange(row, 1, 1, header.length)
    .setValues([totalLine])
    .setFontWeight("bold")
    .setBackground("#f3f3f3");

  const firstDataRow = startRow + 1;
  const lastDataRow = row;

  // 🎨 MISE EN COULEUR DES CELLULES (hors colonne nom)
  for (let c = 2; c <= header.length; c++) {
    const range = sheet.getRange(firstDataRow, c, lastDataRow - firstDataRow + 1, 1);
    const values = range.getValues();
    const backgrounds = [];

    for (let r = 0; r < values.length; r++) {
      const v = values[r][0];
      let color = "#ffffff";

      if (v === 0) color = "#eeeeee";              // 0h
      else if (v > 0 && v < 7) color = "#ffe5cc";  // partiel
      else if (v === 7) color = "#e6f4ea";         // normal
      else if (v > 7) color = "#fce8e6";           // anomalie

      backgrounds.push([color]);
    }

    range.setBackgrounds(backgrounds);
  }

  // 🧱 Bordures
  sheet.getRange(startRow, 1, lastDataRow - startRow + 1, header.length)
    .setBorder(true, true, true, true, true, true);

  sheet.autoResizeColumns(1, header.length);

  return row + 2;
}

function buildReport(sheet, monthKey, startDate, endDate, collaborateurs, planEnt, planForm, real, anomalies, joursEnt, joursForm, joursPres, emails, planByWeek, realByWeek) {
  sheet.clear();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const [month, yearShort] = monthKey.split("/");
  const year = 2000 + parseInt(yearShort);
  const joursOuvres = getBusinessDays(startDate, endDate);
  const heuresOuvrees = joursOuvres * 7;
  const weeks = getWeeksInMonth(startDate, endDate);

  let row = 1;
  sheet.getRange(row++, 1).setValue("Synthèse PRÉSENTIEL - Mois " + monthKey).setFontWeight("bold").setFontSize(14);
  sheet.getRange(row++, 1).setValue(`Jours ouvrés : ${joursOuvres} / Heures ouvrées : ${heuresOuvrees}`);

  row += 1;

  // 🧑‍💼 STATUT PAR COLLABORATEUR
  const header = ["Collaborateur", "Email", "Statut", "H plan E", "H plan F", "H Tot Plan", "H Réal", "H Abs", "Ecart présence", "Ecart MOIS"];
  sheet.getRange(row, 1, 1, header.length).setValues([header]).setFontWeight("bold").setBackground("#dfefff");
  row++;

  const lignes = [];

  for (const nom in collaborateurs) {
    const stat = collaborateurs[nom].statut || "";
    const email = emails[nom] || "";
    const joursEntSet = joursEnt[nom] || new Set();
    const joursFormSet = joursForm[nom] || new Set();
    const joursPresSet = joursPres[nom] || new Set();

    const hPlanEnt = Object.values(planEnt[nom] || {}).reduce((a, b) => a + b, 0);
    const hPlanForm = Object.values(planForm[nom] || {}).reduce((a, b) => a + b, 0);
    const hTotPlan = hPlanEnt + hPlanForm;
    const hReel = Object.values(real[nom] || {}).reduce((a, b) => a + b, 0);

    if (hTotPlan === 0 && hReel === 0) continue;

    const joursPlan = new Set([...joursEntSet, ...joursFormSet]);
    const hAbs = Math.max(hTotPlan - hReel, 0);
    const ecartPresence = roundToQuarterHour(hTotPlan - (hReel + hPlanForm));
    const ecartHeures = roundToQuarterHour(heuresOuvrees - (hReel + hPlanForm));

    lignes.push([nom, email, stat, hPlanEnt, hPlanForm, hTotPlan, hReel, hAbs, ecartPresence, ecartHeures]);
  }

  if (lignes.length) {
    sheet.getRange(row, 1, lignes.length, header.length).setValues(lignes);
    sheet.getRange(row - 1, 1, lignes.length + 1, header.length).setBorder(true, true, true, true, true, true);
    row += lignes.length + 2;
  }

  // 📊 DÉTAIL COLLABORATEUR PAR SEMAINE
  sheet.getRange(row++, 1).setValue("Détail heures réalisées par semaine").setFontWeight("bold");
  const hWeeks = ["Collaborateur", ...weeks.map(w => `Semaine ${w}`), "Total"];
  sheet.getRange(row, 1, 1, hWeeks.length).setValues([hWeeks]).setFontWeight("bold").setBackground("#f0f0f0");
  row++;

  for (const nom in realByWeek) {
    const data = weeks.map(w => roundToQuarterHour(realByWeek[nom][w] || 0));
    const total = data.reduce((a, b) => a + b, 0);
    sheet.getRange(row++, 1, 1, hWeeks.length).setValues([[nom, ...data, total]]);
  }

  row += 2;


  // 📅 RÉCAP QUOTIDIEN PAR COLLABORATEUR
  sheet.getRange(row++, 1)
    .setValue("Récapitulatif quotidien des heures réalisées (Lundi → Vendredi)")
    .setFontWeight("bold");

  row = buildDailyRecapTable_(
    sheet,
    row,
    collaborateurs,
    real,
    startDate,
    endDate
  );

  // 📈 GLOBAL PAR SEMAINE
  sheet.getRange(row++, 1).setValue("Synthèse globale par semaine").setFontWeight("bold");
  const head = ["Type", ...weeks.map(w => `Semaine ${w}`), "Total"];
  const types = ["H plan E", "H plan F", "H Tot Plan", "H Réal", "H Abs"];
  const dataGlobal = types.map(type => {
    const row = [type];
    let total = 0;
    for (const wk of weeks) {
      let val = 0;
      for (const nom in planByWeek) {
        if (!planByWeek[nom][wk]) continue;
        if (type === "H plan E") val += planByWeek[nom][wk].entreprise || 0;
        if (type === "H plan F") val += planByWeek[nom][wk].formation || 0;
        if (type === "H Tot Plan") val += (planByWeek[nom][wk].entreprise || 0) + (planByWeek[nom][wk].formation || 0);
      }
      if (type === "H Réal") {
        for (const nom in realByWeek) {
          val += realByWeek[nom]?.[wk] || 0;
        }
      }
      if (type === "H Abs") {
        for (const nom in planByWeek) {
          if (!planByWeek[nom][wk]) continue;
          const totalPlan = (planByWeek[nom][wk].entreprise || 0) + (planByWeek[nom][wk].formation || 0);
          const real = realByWeek[nom]?.[wk] || 0;
          val += Math.max(totalPlan - real, 0);
        }
      }
      row.push(roundToQuarterHour(val));
      total += val;
    }
    row.push(roundToQuarterHour(total));
    return row;
  });

  sheet.getRange(row, 1, dataGlobal.length + 1, head.length).setValues([head, ...dataGlobal]);
  sheet.getRange(row, 1, dataGlobal.length + 1, head.length).setBorder(true, true, true, true, true, true);
  row += dataGlobal.length + 2;

  // 🔍 ANOMALIES
  if (anomalies.length) {
    sheet.getRange(row++, 1).setValue("Liste des anomalies").setFontWeight("bold").setBackground("#ffe5e5");
    const ah = ["Collaborateur", "Email", "Statut", "Date", "Problème", "Détail"];
    sheet.getRange(row, 1, 1, ah.length).setValues([ah]).setFontWeight("bold");
    row++; // ✅ Correction ici !
    sheet.getRange(row, 1, anomalies.length, ah.length).setValues(anomalies);
    sheet.getRange(row - 1, 1, anomalies.length + 1, ah.length).setBorder(true, true, true, true, true, true);
    row += anomalies.length + 2;
  }

  sheet.autoResizeColumns(1, 15);
}

function getBusinessDays(start, end) {
  let days = 0;
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    const day = d.getDay();
    if (day !== 0 && day !== 6) days++;
  }
  return days;
}

function getWeeksInMonth(start, end) {
  const weeks = new Set();
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    weeks.add(getWeekNumber(d));
  }
  return Array.from(weeks).sort((a, b) => a - b);
}

function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function roundToQuarterHour(val) {
  return Math.round(val * 4) / 4;
}


function parseHeureToDate(hhmm) {
  const parts = hhmm.split(":");
  if (parts.length !== 2 && parts.length !== 3) throw new Error("Format heure invalide : " + hhmm);
  const h = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  if (isNaN(h) || isNaN(m)) throw new Error("Heure ou minute invalide : " + hhmm);
  return new Date(1970, 0, 1, h, m, 0);
}

function afficherToutesLesFeuilles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuilles = ss.getSheets();
  feuilles.forEach(feuille => feuille.showSheet());
  SpreadsheetApp.getUi().alert("✅ Toutes les feuilles sont à nouveau visibles !");
}

function nettoyerFeuilleFormulaire() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 🧹 Supprimer la feuille "Formulaire" si elle existe
  const feuilleFormulaire = ss.getSheetByName("Formulaire");
  if (feuilleFormulaire) {
    ss.deleteSheet(feuilleFormulaire);
    Logger.log("✅ Feuille 'Formulaire' supprimée.");
  } else {
    Logger.log("ℹ️ Aucune feuille 'Formulaire' trouvée.");
  }

  // 🧽 Vider la cellule A1 de toutes les feuilles (au cas où le lien aurait été mis ailleurs)
  const feuilles = ss.getSheets();
  feuilles.forEach(f => f.getRange("A1").clearContent());

  Logger.log("✅ Cellule A1 vidée dans toutes les feuilles.");
}



function genererRecapQuotidienDansNouvelleFeuille(monthKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const presencesSheet = ss.getSheetByName("Présences");
  if (!presencesSheet) {
    throw new Error("❌ Feuille 'Présences' introuvable");
  }

  const tz = ss.getSpreadsheetTimeZone();
  const [mm, yy] = monthKey.split("/");
  const year = 2000 + parseInt(yy, 10);
  const month = parseInt(mm, 10) - 1;

  const startDate = new Date(year, month, 1);
  const endDate = new Date(year, month + 1, 0);

  // 🧾 Préparer / recréer la feuille
  const sheetName = `Récap Quotidien ${monthKey}`;
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  // 📥 Lecture Présences
  const data = presencesSheet.getDataRange().getValues();
  const headers = data.shift();

  const idxDate = headers.indexOf("Date");
  const idxNom = headers.indexOf("Nom Complet");
  const idxHeures = headers.indexOf("Nb heures");

  if (idxDate === -1 || idxNom === -1 || idxHeures === -1) {
    throw new Error("❌ Colonnes requises manquantes dans 'Présences'");
  }

  // 🔹 Structure : nom → semaine → jour (1–5) → heures
  const map = {};
  const weeksSet = new Set();

  data.forEach(row => {
    const rawDate = row[idxDate];
    if (!rawDate) return;

    const d = new Date(rawDate);
    if (d < startDate || d > endDate) return;

    const dow = d.getDay(); // 1=lundi … 5=vendredi
    if (dow < 1 || dow > 5) return;

    const nom = (row[idxNom] || "").toString().trim();
    if (!nom) return;

    const h = Math.round((parseFloat(row[idxHeures]) || 0) * 4) / 4;
    const week = getWeekNumber(d);

    weeksSet.add(week);

    if (!map[nom]) map[nom] = {};
    if (!map[nom][week]) map[nom][week] = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };

    map[nom][week][dow] += h;
  });

  const weeks = Array.from(weeksSet).sort((a, b) => a - b);

  // 🧾 TITRE
  sheet.getRange(1, 1).setValue(`Récapitulatif quotidien – ${monthKey}`).setFontWeight("bold");

  // 🧾 EN-TÊTE
  const header = ["Collaborateur"];
  weeks.forEach(w => {
    header.push("Lun", "Mar", "Mer", "Jeu", "Ven", `Total S${w}`);
  });
  header.push("Total Mois");

  sheet.getRange(3, 1, 1, header.length)
    .setValues([header])
    .setFontWeight("bold")
    .setBackground("#e8f0fe");

  let row = 4;
  const totalJour = Array(header.length - 1).fill(0);

  // 🧑‍💼 LIGNES COLLABORATEURS
  Object.keys(map).sort().forEach(nom => {
    const line = [nom];
    let totalMois = 0;
    let col = 0;

    weeks.forEach(w => {
      let totalSemaine = 0;
      for (let dow = 1; dow <= 5; dow++) {
        const h = Math.round((map[nom][w]?.[dow] || 0) * 4) / 4;
        line.push(h);
        totalSemaine += h;
        totalMois += h;
        totalJour[col++] += h;
      }
      line.push(totalSemaine);
      totalJour[col++] += totalSemaine;
    });

    line.push(totalMois);
    totalJour[col] += totalMois;

    sheet.getRange(row++, 1, 1, header.length).setValues([line]);
  });

  // ➕ TOTAL JOUR
  const totalLine = ["TOTAL JOUR", ...totalJour];
  sheet.getRange(row, 1, 1, header.length)
    .setValues([totalLine])
    .setFontWeight("bold")
    .setBackground("#f3f3f3");

  // 🎨 COULEURS
  const firstDataRow = 4;
  const lastDataRow = row;
  for (let c = 2; c <= header.length; c++) {
    const range = sheet.getRange(firstDataRow, c, lastDataRow - firstDataRow + 1, 1);
    const values = range.getValues();
    const colors = values.map(([v]) => {
      if (v === 0) return ["#eeeeee"];
      if (v > 0 && v < 7) return ["#ffe5cc"];
      if (v === 7) return ["#e6f4ea"];
      if (v > 7) return ["#fce8e6"];
      return ["#ffffff"];
    });
    range.setBackgrounds(colors);
  }

  sheet.autoResizeColumns(1, header.length);
}

function TEST_recap_quotidien_janvier_2026() {
  genererRecapQuotidienDansNouvelleFeuille("01/26");
}


function recap_quotidien_mois_precedent() {
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);

  const mm = Utilities.formatString("%02d", prev.getMonth() + 1);
  const yy = prev.getFullYear().toString().slice(-2);

  const monthKey = `${mm}/${yy}`;

  genererRecapQuotidienDansNouvelleFeuille(monthKey);
}