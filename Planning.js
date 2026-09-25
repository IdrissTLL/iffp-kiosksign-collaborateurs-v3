function processMonthlyReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const presencesSheet = ss.getSheetByName("Présences");
  if (!presencesSheet) throw new Error("Feuille 'Présences' introuvable.");

  const presencesData = presencesSheet.getDataRange().getValues();
  const headers = presencesData.shift();
  const dateIdx = headers.indexOf("Date");
  if (dateIdx === -1) throw new Error("Colonne 'Date' absente dans la feuille 'Présences'.");

  const months = {};
  presencesData.forEach(row => {
    const rawDate = row[dateIdx];
    if (!rawDate || rawDate === "") return;
    const date = new Date(rawDate);
    const mm = Utilities.formatString("%02d", date.getMonth() + 1);
    const yy = date.getFullYear().toString().slice(-2);
    const key = `${mm}/${yy}`;
    if (!months[key]) months[key] = [];
    months[key].push(date);
  });

  Object.keys(months).forEach(monthKey => {
    const reportingName = "Reporting " + monthKey;
    const planningName = "Planning " + monthKey;
    if (!ss.getSheetByName(reportingName)) {
      const planningMonth = parseInt(monthKey.split("/")[0]);
      const year = 2000 + parseInt(monthKey.split("/")[1]);
      importPlanning(planningMonth, year, monthKey, false);
      const reportingSheet = ss.insertSheet(reportingName);
      generateMonthlyReport(reportingSheet, monthKey);
      const sheets = ss.getSheets();
      const planningSheet = ss.getSheetByName(planningName);
      ss.setActiveSheet(reportingSheet);
      ss.moveActiveSheet(sheets.length - 1); // reporting à droite
      if (planningSheet) {
        ss.setActiveSheet(planningSheet);
        ss.moveActiveSheet(ss.getSheets().length); // planning tout à droite
      }
    }
  });
}

function importPlanning(month, year, monthKey, shouldOverwrite) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planningSheetName = "Planning " + monthKey;
  const planningFileId = "1jA5xx-oNB4WJyoCcNH1M6Csj0-FUs3TjPQXuXbjkd0w";
  const sourceFile = SpreadsheetApp.openById(planningFileId);
  const sourceSheetName = "Planning " + Utilities.formatString("%02d", month) + "/" + year.toString().slice(-2);
  const sourceSheet = sourceFile.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    Logger.log("Feuille planning source introuvable : " + sourceSheetName);
    return;
  }

  let targetSheet = ss.getSheetByName(planningSheetName);
  if (!targetSheet) {
    targetSheet = sourceSheet.copyTo(ss).setName(planningSheetName);
    Logger.log("Feuille planning copiée : " + planningSheetName);
  } else if (shouldOverwrite) {
    const data = sourceSheet.getDataRange().getValues();
    targetSheet.clear();
    targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    Logger.log("Feuille planning écrasée : " + planningSheetName);
  } else {
    const existingData = targetSheet.getDataRange().getValues();
    const existingDates = existingData.map(row => row[0]?.toString());
    const newData = sourceSheet.getDataRange().getValues().slice(1);
    newData.forEach(row => {
      if (!existingDates.includes(row[0]?.toString())) {
        targetSheet.appendRow(row);
      }
    });
    Logger.log("Planning mis à jour (ajout) : " + planningSheetName);
  }

  ss.setActiveSheet(targetSheet);
  ss.moveActiveSheet(ss.getSheets().length); // planning toujours en dernier
}

function generateMonthlyReport(sheet, monthKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const planningSheet = ss.getSheetByName("Planning " + monthKey);
  const presencesSheet = ss.getSheetByName("Présences");

  // 🗂️ Utilisation de la feuille centralisée
  const fichierCentralId = "1dQEgyFZ2mmoT0PzFK1gc-KP_PVL75K9hekE0NFnos4s";
  const fichierCentral = SpreadsheetApp.openById(fichierCentralId);
  const employesSheet = fichierCentral.getSheetByName("liste collaborateurs");

  const officialIP = "83.118.196.38";
  const tz = ss.getSpreadsheetTimeZone();

  const [month, yearShort] = monthKey.split("/");
  const year = 2000 + parseInt(yearShort);
  const startDate = new Date(year, parseInt(month) - 1, 1);
  const endDate = new Date(year, parseInt(month), 0);

  const employesData = employesSheet.getDataRange().getValues();
  const enteteEmployes = employesData[0];
  const idxNom = enteteEmployes.indexOf("Nom Complet");
  const idxStatut = enteteEmployes.indexOf("Statut");
  const idxEntree = enteteEmployes.indexOf("Date entrée");
  const idxSortie = enteteEmployes.indexOf("Date Sortie");

  const collaborateursActifs = {};
  for (let i = 1; i < employesData.length; i++) {
    let nom = (employesData[i][idxNom] || "").toString().trim();
    nom = nom.replace(/\s+/g, " ");
    const statut = employesData[i][idxStatut] || "";
    const entree = new Date(employesData[i][idxEntree]);
    const sortie = employesData[i][idxSortie] ? new Date(employesData[i][idxSortie]) : new Date("2100-01-01");
    if (entree <= endDate && sortie >= startDate) {
      collaborateursActifs[nom] = { statut, actif: true };
    }
  }

  const presencesData = presencesSheet.getDataRange().getValues();
  const entetePresences = presencesData[0];
  const dataPresences = presencesData.slice(1);

  const pIdx = {
    date: entetePresences.indexOf("Date"),
    nom: entetePresences.indexOf("Nom Complet"),
    email: entetePresences.indexOf("Email"),
    heures: entetePresences.indexOf("Nb heures"),
    signature: entetePresences.indexOf("Signature URL"),
    ipLogin: entetePresences.indexOf("Adresse IP Login"),
    ipLogout: entetePresences.indexOf("Adresse IP Logout")
  };

  const realisations = {};
  const anomalies = [];
  const joursPresence = {};
  const emails = {};
  const realByWeek = {};

  for (const row of dataPresences) {
    const rawDate = row[pIdx.date];
    if (!rawDate) continue;
    const date = new Date(rawDate);
    if (date < startDate || date > endDate) continue;

    let nom = (row[pIdx.nom] || "").toString().trim().replace(/\s+/g, " ");
    const email = (row[pIdx.email] || "").toString().trim();
    const heures = roundToQuarterHour(parseFloat(row[pIdx.heures]) || 0);
    const dateKey = date.toDateString();
    const week = getWeekNumber(date);

    emails[nom] = email;

    if (!joursPresence[nom]) joursPresence[nom] = new Set();
    joursPresence[nom].add(dateKey);

    if (!realisations[nom]) realisations[nom] = {};
    if (!realByWeek[nom]) realByWeek[nom] = {};
    if (!realByWeek[nom][week]) realByWeek[nom][week] = 0;

    realisations[nom][dateKey] = (realisations[nom][dateKey] || 0) + heures;
    realByWeek[nom][week] += heures;

    const ipLogin = (row[pIdx.ipLogin] || "").trim();
    const ipLogout = (row[pIdx.ipLogout] || "").trim();
    const signature = (row[pIdx.signature] || "").trim();

    if ((ipLogin && ipLogin !== officialIP) || (ipLogout && ipLogout !== officialIP)) {
      anomalies.push([nom, email, collaborateursActifs[nom]?.statut || "", dateKey, "IP différente", `Login: ${ipLogin} / Logout: ${ipLogout}`]);
    }

    if (!signature) {
      anomalies.push([nom, email, collaborateursActifs[nom]?.statut || "", dateKey, "Signature manquante", "Aucune signature"]);
    }

    if (!collaborateursActifs[nom]) {
      Logger.log("NOM PRÉSENT MAIS NON DÉTECTÉ COMME ACTIF : " + nom);
      collaborateursActifs[nom] = { statut: "", actif: true };
    }
  }

  const planningData = planningSheet ? planningSheet.getDataRange().getValues().slice(1) : [];
  const entetePlanning = planningSheet ? planningSheet.getDataRange().getValues()[0] : [];
  const idxDateP = entetePlanning.indexOf("Date");
  const idxNomP = entetePlanning.indexOf("Nom Complet");
  const idxHeures = entetePlanning.indexOf("Nb H plan.");
  const idxRemarques = entetePlanning.indexOf("Remarques");

  const planifieEntreprise = {};
  const planifieFormation = {};
  const joursEntreprise = {};
  const joursFormation = {};
  const planByWeek = {};

  for (const row of planningData) {
    const date = new Date(row[idxDateP]);
    if (date < startDate || date > endDate) continue;

    let nom = (row[idxNomP] || "").toString().trim().replace(/\s+/g, " ");
    const heuresBrutes = row[idxHeures];
    const remarques = (row[idxRemarques] || "").toString().toLowerCase();
    const week = getWeekNumber(date);
    const dateKey = date.toDateString();

    if (!collaborateursActifs[nom]) continue;

    let heures = parseFloat(heuresBrutes);
    if (isNaN(heures)) heures = 0;
    if (heures === 0 && remarques.includes("formation")) heures = 7.0;
    heures = roundToQuarterHour(heures);

    const isFormation = remarques.includes("formation");
    const target = isFormation ? planifieFormation : planifieEntreprise;
    const jours = isFormation ? joursFormation : joursEntreprise;

    if (!target[nom]) target[nom] = {};
    if (!planByWeek[nom]) planByWeek[nom] = {};
    if (!planByWeek[nom][week]) planByWeek[nom][week] = { entreprise: 0, formation: 0 };

    target[nom][dateKey] = (target[nom][dateKey] || 0) + heures;
    if (isFormation) {
      planByWeek[nom][week].formation += heures;
    } else {
      planByWeek[nom][week].entreprise += heures;
    }

    if (!jours[nom]) jours[nom] = new Set();
    jours[nom].add(dateKey);
  }

  buildReport(
    sheet, monthKey, startDate, endDate, collaborateursActifs,
    planifieEntreprise, planifieFormation, realisations, anomalies,
    joursEntreprise, joursFormation, joursPresence, emails,
    planByWeek, realByWeek
  );
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

// 🔍 Comparaison des collaborateurs actifs entre 2 feuilles "Employés"
function comparerCollaborateursActifs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const local = ss.getSheetByName("Employés");
  const sourceId = "1jA5xx-oNB4WJyoCcNH1M6Csj0-FUs3TjPQXuXbjkd0w";
  const source = SpreadsheetApp.openById(sourceId).getSheetByName("Employés");
  const getActifs = (sheet) => {
    const data = sheet.getDataRange().getValues();
    const idxNom = data[0].indexOf("Nom Complet");
    const idxActif = data[0].indexOf("Actif");
    const set = new Set();
    for (let i = 1; i < data.length; i++) {
      const actif = data[i][idxActif];
      if (actif === "Oui" || actif === "oui" || actif === true) {
        set.add((data[i][idxNom] || "").toString().trim());
      }
    }
    return set;
  };

  const localSet = getActifs(local);
  const sourceSet = getActifs(source);
  const onlyInLocal = [...localSet].filter(x => !sourceSet.has(x));
  const onlyInSource = [...sourceSet].filter(x => !localSet.has(x));

  Logger.log("🧾 Différences de collaborateurs ACTIFS :");
  Logger.log("✅ Présents uniquement en local : " + onlyInLocal.join(", "));
  Logger.log("🕵️ Présents uniquement dans la source : " + onlyInSource.join(", "));
}

function nettoyerPresences() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Présences");
  if (!feuille) {
    Logger.log("❌ Feuille 'Présences' introuvable.");
    return;
  }

  const data = feuille.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const nRows = rows.length;

  const idx = {
    date: headers.indexOf("Date"),
    debut: headers.indexOf("Heure d'arrivée"),
    fin: headers.indexOf("Heure de départ"),
    nbH: headers.indexOf("Nb heures"),
    nomComplet: headers.indexOf("Nom Complet")
  };

  // Ajout de la colonne Nb H Valid si manquante
  let idxNbHValid = headers.indexOf("Nb H Valid");
  if (idxNbHValid === -1) {
    idxNbHValid = headers.length;
    feuille.getRange(1, idxNbHValid + 1).setValue("Nb H Valid");
  }

  const doublons = {};

  for (let i = 0; i < nRows; i++) {
    const ligne = rows[i];
    const nom = (ligne[idx.nomComplet] || "").toString().trim().replace(/\s+/g, " ");
    const date = ligne[idx.date];
    const debut = ligne[idx.debut];
    const fin = ligne[idx.fin];
    const key = `${nom}_${new Date(date).toDateString()}`;
    doublons[key] = (doublons[key] || 0) + 1;

    const rangeNbH = feuille.getRange(i + 2, idx.nbH + 1);
    const rangeNbHValid = feuille.getRange(i + 2, idxNbHValid + 1);
    const rangeNom = feuille.getRange(i + 2, idx.nomComplet + 1);

    let nbH = ligne[idx.nbH];
    let heures = 0;

    // Calcul des heures si cellule vide
    if (!nbH || nbH === "") {
      if (debut && fin) {
        try {
          let h1 = typeof debut === 'string' ? parseHeureToDate(debut) : new Date(debut);
          let h2 = typeof fin === 'string' ? parseHeureToDate(fin) : new Date(fin);
          let loginMinutes = h1.getHours() * 60 + h1.getMinutes();
          let logoutMinutes = h2.getHours() * 60 + h2.getMinutes();
          let diff = logoutMinutes - loginMinutes;
          diff = Math.round(diff / 5) * 5; // arrondi aux 5 minutes
          heures = diff / 60;
          heures = Math.round(heures * 100) / 100;
          nbH = heures.toFixed(2).replace(".", ",");
          feuille.getRange(i + 2, idx.nbH + 1).setValue(nbH);
        } catch (e) {
          Logger.log(`⛔ Erreur ligne ${i + 2} : ${e}`);
          feuille.getRange(i + 2, idx.nbH + 1).setValue("#ERREUR");
          continue;
        }
      }
    } else {
      heures = parseFloat(nbH.toString().replace(",", "."));
    }

    // Calcul Nb H Valid (heures - 0.25)
    let nbHValid = Math.max(heures - 0.25, 0);
    nbHValid = Math.round(nbHValid * 100) / 100;
    const nbHValidStr = nbHValid.toFixed(2).replace(".", ",");
    rangeNbHValid.setValue(nbHValidStr);

    // Mise en forme rouge si < 2h
    if (heures < 2) {
      rangeNbH.setBackground("#f4cccc");
      rangeNbHValid.setBackground("#f4cccc");
    } else {
      rangeNbH.setBackground(null);
      rangeNbHValid.setBackground(null);
    }
  }

  // Mise en rouge des noms avec plus de 2 entrées le même jour
  Object.keys(doublons).forEach(key => {
    if (doublons[key] > 2) {
      const [nom, dateStr] = key.split("_");
      for (let i = 0; i < nRows; i++) {
        const ligne = rows[i];
        const nomL = (ligne[idx.nomComplet] || "").toString().trim().replace(/\s+/g, " ");
        const dateL = new Date(ligne[idx.date]).toDateString();
        if (nomL === nom && dateL === dateStr) {
          feuille.getRange(i + 2, idx.nomComplet + 1).setBackground("#f4cccc");
        }
      }
    }
  });

  // Format des heures → décimal
  feuille.getRange(2, idx.nbH + 1, nRows).setNumberFormat("0.00");
  feuille.getRange(2, idxNbHValid + 1, nRows).setNumberFormat("0.00");

  Logger.log("✅ Nettoyage des heures terminé !");
}

function parseHeureToDate(hhmm) {
  const parts = hhmm.split(":");
  if (parts.length !== 2 && parts.length !== 3) throw new Error("Format heure invalide : " + hhmm);
  const h = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  if (isNaN(h) || isNaN(m)) throw new Error("Heure ou minute invalide : " + hhmm);
  return new Date(1970, 0, 1, h, m, 0);
}

function envoyerRappelsConnexions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const presencesSheet = ss.getSheetByName("Présences");
  const employesSheet = ss.getSheetByName("Employés");
  const planningSheet = ss.getSheetByName("Planning 07/25");
  if (!presencesSheet || !employesSheet || !planningSheet) return;

  const dataPres = presencesSheet.getDataRange().getValues();
  const entetesPres = dataPres[0];
  const pres = dataPres.slice(1);
  const idxNomComplet = entetesPres.indexOf("Nom Complet");
  const idxDate = entetesPres.indexOf("Date");
  const idxHeureArrivee = entetesPres.indexOf("Heure d'arrivée");
  const idxHeureDepart = entetesPres.indexOf("Heure de départ");

  const employes = employesSheet.getDataRange().getValues();
  const entetesEmp = employes[0];
  const idxNomEmp = entetesEmp.indexOf("Nom Complet");
  const idxNom = entetesEmp.indexOf("Nom");
  const idxCiv = entetesEmp.indexOf("Civ");

  const planningData = planningSheet.getDataRange().getValues();
  const planningHeaders = planningData[0];
  const idxPlanningDate = planningHeaders.indexOf("Date");
  const idxPlanningNom = planningHeaders.indexOf("Nom Complet");
  const idxPlanningHeureDebut = planningHeaders.indexOf("Heure Début");

  const today = new Date();
  const todayStr = today.toDateString();
  const tz = ss.getSpreadsheetTimeZone();
  const currentTimeStr = Utilities.formatDate(today, tz, "HH:mm");

  const dejaPresent = {};
  pres.forEach(row => {
    const nom = (row[idxNomComplet] || "").toString().trim();
    const dateStr = new Date(row[idxDate]).toDateString();
    if (dateStr !== todayStr) return;

    const hArr = row[idxHeureArrivee] ? Utilities.formatDate(new Date(row[idxHeureArrivee]), tz, "HH:mm") : "";
    const hDep = row[idxHeureDepart] ? Utilities.formatDate(new Date(row[idxHeureDepart]), tz, "HH:mm") : "";

    if (!dejaPresent[nom]) dejaPresent[nom] = { connexions: [], deconnexions: [] };
    if (hArr) dejaPresent[nom].connexions.push(hArr);
    if (hDep) dejaPresent[nom].deconnexions.push(hDep);
  });

  const horairesParHeureDebut = {
    "08:30": ["08:35", "12:05", "13:35", "17:35"],
    "09:00": ["09:05", "12:35", "14:05", "18:05"],
    "09:30": ["09:35", "13:05", "14:35", "18:35"],
    "10:00": ["10:05", "13:35", "15:05", "19:05"]
  };

  const messages = [];

  employes.slice(1).forEach(row => {
    const nomComplet = (row[idxNomEmp] || "").toString().trim();
    const genre = (row[idxCiv] || "M.").toString().trim();
    const nom = (row[idxNom] || "").toString().trim();

    const journal = dejaPresent[nomComplet] || { connexions: [], deconnexions: [] };

    // 🎯 Chercher l'heure de début planifiée
    const lignePlanning = planningData.find(r =>
      new Date(r[idxPlanningDate]).toDateString() === todayStr &&
      (r[idxPlanningNom] || "").toString().trim() === nomComplet
    );

    if (!lignePlanning || !lignePlanning[idxPlanningHeureDebut]) return;

    const heureDebut = Utilities.formatDate(new Date(todayStr + " " + lignePlanning[idxPlanningHeureDebut]), tz, "HH:mm");
    const horaires = horairesParHeureDebut[heureDebut];
    if (!horaires) return;

    horaires.forEach(h => {
      if (h < currentTimeStr) {
        const type = horaires.indexOf(h) % 2 === 0 ? "connexion" : "déconnexion";
        if (!journal[type + "s"].includes(h)) {
          messages.push({ nomComplet, genre, nom, heure: h, type });
        }
      }
    });
  });

  const webhookUrl = "https://chat.googleapis.com/v1/spaces/qWoOH8AAAAE/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=Qbu2gwyXHfWODbjMvIqwTlWxXsBkBCrUPqV7aMwCmSQ"; // ⛔️ À remplacer par le vrai webhook
  messages.forEach(msg => {
    const texte = `Bonjour ${msg.genre} ${msg.nom}, vous n'avez pas effectué votre ${msg.type} prévue à ${msg.heure}.\nPensez à vous ${msg.type === "connexion" ? "connecter" : "déconnecter"}.\nMerci,\nL'équipe IFFP.`;
    const payload = JSON.stringify({ text: texte });
    UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload
    });
  });

  Logger.log("📢 Rappels envoyés : " + messages.length);
}

// 📅 Fonction déclencheur automatique (à utiliser dans le menu Triggers)
function creerDeclencheurRappels() {
  ScriptApp.newTrigger("envoyerRappelsConnexions")
    .timeBased()
    .everyMinutes(30) // toutes les 30 minutes
    .create();
}

// ⛔️ Fonction pour désactiver tous les déclencheurs liés à ce script
function supprimerTousLesDeclencheurs() {
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "envoyerRappelsConnexions") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
