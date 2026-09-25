function onOpen() {
  afficherSeulementMoisEnCours();

  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 Kiosksign Reporting")
    .addItem('🧮 Lancer le reporting mensuel', 'processMonthlyReports')
    .addItem('🔔 Lancer le RECAP EVP mensuel', 'genererReportingMoisPrecedent')
    .addItem('🔔 Lancer le Report M-1', 'genererReportMoisPrecedent')
    .addItem('🔔 Générer Report Mensuel XX/XX', 'genererRecapCollaborateursMoisPrompt')
    .addSeparator()
    .addItem('📧 TEST envoi récap hebdo (→ serge)', 'TEST_sendWeeklyRecapCollaborators')
    .addItem('📧 TEST envoi récap mensuel (→ serge)', 'TEST_sendMonthlyRecapCollaborators')
    .addItem('⏰ Installer triggers logout + récap email', 'installerTriggersLogoutCollaborateurs')
    .addSeparator()
    .addItem('📤 Importer un planning manuellement', 'importPlanningDialogue')
    .addSeparator()
    .addItem('🔍 Comparer collaborateurs actifs', 'comparerCollaborateursActifs')
    .addSeparator()
    .addItem('Mise à jour Noms ID', 'majNomCompletEtUserID')
    .addSeparator()
    .addItem('📊 Récap S-1 RH', 'recapHebdoRH')
    .addSeparator()
    .addItem('📬 Récap perso + planning S+1', 'recapPersoFinSemaine')
    .addSeparator()
    .addItem('📈 Récap Global (veille, semaine, mois)', 'recapGlobalAll')
    .addSeparator()
    .addItem('📤 Envoyer total heures journalières', 'envoyerTotalJournalierComptable')
    .addSeparator()
    .addItem("🧹 Nettoyer heures (Présences)", "nettoyerPresences")
    .addItem("↕️ Grouper présences par jour", "regrouperToutesLesPresencesParJour")
    .addItem("↕️ Grouper présences semaine dernière", "regrouperSemaineDerniere")
    .addSeparator()
    .addItem("🔁 Compléter infos + recalcul", "completerInfosPresences")
    .addItem("🔄 Recalculer toutes les heures", "recalculerHeuresPresence")
    .addToUi();

  ui.createMenu("📣 Notifications")
    .addItem("🔁 Envoyer les notifications", "traiterPresences")
    .addItem("⚠️ Vérifier les webhooks manquants", "verifierWebhooksManquants")
    .addToUi();

  ui.createMenu("📋 Émargement")
    .addItem("📋 Ouvrir le formulaire", "ouvrirFormulaireEmargement")
    .addToUi();
}

function afficherSeulementMoisEnCours() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const mois = Utilities.formatString("%02d", now.getMonth() + 1);
  const annee = now.getFullYear().toString().slice(-2);
  const moisActuel = `${mois}/${annee}`;
  
  const feuilles = ss.getSheets();

  feuilles.forEach(feuille => {
    const nom = feuille.getName();

    const toujoursVisibles = (
      nom === "Présences" ||
      nom === "Employés" ||
      nom.startsWith("Présences")
    );

    const estMoisEnCours = (
      nom.includes(moisActuel) &&
      (nom.startsWith("Reporting ") || nom.startsWith("Planning "))
    );

    const estAncienMois = (
      (nom.startsWith("Reporting ") || nom.startsWith("Planning ")) &&
      !nom.includes(moisActuel)
    );

    if (toujoursVisibles || estMoisEnCours) {
      feuille.showSheet();
    } else if (estAncienMois) {
      feuille.hideSheet();
    } else {
      feuille.showSheet(); // autres feuilles persos restent visibles
    }
  });
}

function ouvrirFormulaireEmargement() {
  const url = "https://script.google.com/a/macros/iffp.school/s/AKfycbzIeTeCMr1ltS51Qnrv_CvQ6bYeBklCkPffXSyNekW5/dev";
  const html = HtmlService.createHtmlOutput(`
    <script>
      window.open("${url}", "_blank");
      google.script.host.close();
    </script>
  `);
  SpreadsheetApp.getUi().showModalDialog(html, "Redirection vers le formulaire...");
}


function testLogUserIn() {
  Logger.log("testLogUserIn() démarré");
  var testIP = "127.0.0.1";
  var testEmail = "test@example.com";
  var testFirstName = "TestPrenom";
  var testLastName = "TestNom";
  var testShift = "Après-midi";
  var testLocation = "Nanterre Pesaro";
  var testRemarks = "Aucune";

  var result = logUser('in', '', testIP, testEmail, testFirstName, testLastName, testShift, testLocation, testRemarks);
  Logger.log("Résultat testLogUserIn() : " + JSON.stringify(result));
}

function testLogUserOut() {
  Logger.log("testLogUserOut() démarré");
  var testIP = "127.0.0.1";
  var testEmail = "test@example.com";
  var testFirstName = "TestPrenom";
  var testLastName = "TestNom";
  var testShift = "Après-midi";
  var testLocation = "Nanterre Pesaro";
  var testRemarks = "Aucune";

  var dummySignature = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUA";

  var result = logUser('out', dummySignature, testIP, testEmail, testFirstName, testLastName, testShift, testLocation, testRemarks);
  Logger.log("Résultat testLogUserOut() : " + JSON.stringify(result));
}

function testLien() {
  const name = SpreadsheetApp.getActiveSpreadsheet().getName();
  SpreadsheetApp.getUi().alert("Ce script est bien lié à : " + name);
}

function importPlanningDialogue() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Importer un planning",
    "Saisis le mois au format MM/YY suivi de 'oui' ou 'non' pour écraser les données (ex: 04/25 oui)", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const match = input.match(/^(\d{2})\/(\d{2})\s+(oui|non)$/i);
  if (!match) {
    ui.alert("Format invalide. Réessaie avec : 04/25 oui");
    return;
  }

  const month = parseInt(match[1]);
  const year = 2000 + parseInt(match[2]);
  const shouldOverwrite = match[3].toLowerCase() === "oui";
  const monthKey = match[1] + "/" + match[2];

  importPlanning(month, year, monthKey, shouldOverwrite);
}

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

  const centralSheetId = "1jA5xx-oNB4WJyoCcNH1M6Csj0-FUs3TjPQXuXbjkd0w"; // Fichier centralisé
  const collaborateursSheet = SpreadsheetApp.openById(centralSheetId).getSheetByName("liste collaborateurs");

  const officialIP = "83.118.196.38";
  const tz = ss.getSpreadsheetTimeZone();

  const [month, yearShort] = monthKey.split("/");
  const year = 2000 + parseInt(yearShort);
  const startDate = new Date(year, parseInt(month) - 1, 1);
  const endDate = new Date(year, parseInt(month), 0);

  const collaborateursData = collaborateursSheet.getDataRange().getValues();
  const enteteCollaborateurs = collaborateursData[0];
  const idxNom = enteteCollaborateurs.indexOf("Nom Complet");
  const idxStatut = enteteCollaborateurs.indexOf("Statut");
  const idxEntree = enteteCollaborateurs.indexOf("Date entrée");
  const idxSortie = enteteCollaborateurs.indexOf("Date Sortie");
  const idxActif = enteteCollaborateurs.findIndex(h => h.toString().toLowerCase().replace(/\s/g, "") === "actif/inactif");

  const collaborateursActifs = {};
  for (let i = 1; i < collaborateursData.length; i++) {
    let nom = (collaborateursData[i][idxNom] || "").toString().trim().replace(/\s+/g, " ");
    const statut = collaborateursData[i][idxStatut] || "";
    const entree = new Date(collaborateursData[i][idxEntree]);
    const sortie = collaborateursData[i][idxSortie] ? new Date(collaborateursData[i][idxSortie]) : new Date("2100-01-01");
    const actif = (collaborateursData[i][idxActif] || "").toString().toLowerCase() === "actif";

    if (entree <= endDate && sortie >= startDate && actif) {
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
    const dateKey = Utilities.formatDate(date, "Europe/Paris", "yyyy-MM-dd");
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
    const dateKey = Utilities.formatDate(date, "Europe/Paris", "yyyy-MM-dd");

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




