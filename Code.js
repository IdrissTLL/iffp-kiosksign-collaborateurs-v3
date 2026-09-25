/**
 * Kiosksign V3 (Collaborateurs) — Version corrigée
 * Application d'émargement pour IFFP.
 *
 * Fonctionnalités :
 * - La feuille "Présences" est organisée avec les colonnes (dans cet ordre) :
 *     Date | Shift | Localisation | Email | Prénom | Nom | Nom Complet | User ID | Heure d'arrivée | Heure de départ | Nb heures | Signature URL | Horodatage signature | Adresse IP Login | Adresse IP Logout | Remarques | Msg Notif
 * - La connexion (LOGIN) est refusée si la localisation est "Nanterre Pesaro" et l'IP est différente de 83.118.196.38.
 * - Pour toute autre localisation et IP ≠ attendue, une remarque doit être obligatoirement renseignée.
 * - Plusieurs connexions/déconnexions sont possibles par jour pour le même utilisateur (chaque LOGIN ajoute une ligne, chaque LOGOUT complète la dernière ouverte).
 */

/* ---------- Helpers ---------- */

/* // ==== Fonction principale appelée automatiquement sur modification ==== 
function onEdit(e) {
  var feuille = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (feuille.getName() === "Présences") {
    nettoyerPresences();
  }
}
*/

function getHeaderIndex(sheet, columnName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim() === columnName) return i + 1;
  }
  return -1;
}

function formatCellDate(cell) {
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, "Europe/Paris", "dd/MM/yyyy").trim();
  }
  return cell.toString().trim();
}

function roundTimeTo5Minutes(timeStr) {
  var parts = timeStr.split(":");
  var hour = parseInt(parts[0], 10);
  var minute = parseInt(parts[1], 10);
  var roundedMinute = Math.round(minute / 5) * 5;
  if (roundedMinute === 60) {
    hour += 1;
    roundedMinute = 0;
  }
  var hStr = hour < 10 ? "0" + hour : "" + hour;
  var mStr = roundedMinute < 10 ? "0" + roundedMinute : "" + roundedMinute;
  return hStr + ":" + mStr;
}

function timeStrToMinutes(timeStr) {
  var parts = timeStr.split(":");
  var hour = parseInt(parts[0], 10);
  var minute = parseInt(parts[1], 10);
  return hour * 60 + minute;
}

function formatTimeStr(timeStr) {
  var parts = timeStr.split(":");
  return parts[0] + "h" + parts[1];
}

function parseTime(cell) {
  if (!cell) return "00:00:00";
  if (cell instanceof Date) {
    return Utilities.formatDate(cell, "Europe/Paris", "HH:mm:ss").trim();
  }
  var str = cell.toString().trim();
  return str === "" ? "00:00:00" : str;
}

function findHeaderIndex(headers, searchName) {
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim().toLowerCase() === searchName.toLowerCase()) {
      return i;
    }
  }
  return -1;
}

/* ---------- Fonctions Serveur ---------- */

function doGet(e) {
  return showUI();
}

function ensureSheetAndHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Présences");
  if (!sheet) {
    sheet = ss.insertSheet("Présences");
  }
  var expectedHeaders = [
    "Date", "Shift", "Localisation", "Email", "Prénom", "Nom", "Nom Complet", "User ID", "Heure d'arrivée",
    "Heure de départ", "Nb heures", "Signature URL", "Horodatage signature", "Adresse IP Login", "Adresse IP Logout", "Remarques", "Msg Notif"
  ];
  var lastCol = sheet.getLastColumn();
  if (lastCol < expectedHeaders.length) {
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
  } else {
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers[0] !== expectedHeaders[0]) {
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    }
    if (headers[headers.length - 1] !== "Msg Notif") {
      sheet.insertColumnAfter(headers.length);
      sheet.getRange(1, headers.length + 1).setValue("Msg Notif");
    }
  }
  return sheet;
}

function getSignatureFolder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var parents = file.getParents();
  var parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var folderName = "Sign emargemt " + Utilities.formatDate(new Date(), "Europe/Paris", "MMyy");
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()){
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

function debugRechercheEmail(motCle) {
  const fichierDestinationId = "1Dn5HfoYTegM9nJR4LKfrqOJ65CgJ_1crtiNFFUd2LMU";
  const feuille = SpreadsheetApp.openById(fichierDestinationId).getSheetByName("liste collaborateurs");
  const donnees = feuille.getDataRange().getValues();
  const headers = donnees[0];
  const colEmail = headers.indexOf("Email");

  Logger.log("Colonne Email trouvée à l'index : " + colEmail);

  let trouveAuMoinsUn = false;
  for (let i = 1; i < donnees.length; i++) {
    const raw = donnees[i][colEmail];
    const str = (raw || "").toString();
    if (str.toLowerCase().includes(motCle.toLowerCase())) {
      trouveAuMoinsUn = true;
      Logger.log(`Ligne ${i + 1} : "${str}" (longueur ${str.length})`);
      Logger.log("Codes caractères : " + Array.from(str).map(c => c.charCodeAt(0)).join(","));
    }
  }
  if (!trouveAuMoinsUn) {
    Logger.log(`Aucune ligne ne contient "${motCle}" dans la colonne Email.`);
  }
}

function testDebugDeux() {
  debugRechercheEmail("tenzin");
  debugRechercheEmail("arley");
}

// ✅ Fonction pour récupérer les infos depuis la feuille "liste collaborateurs"
function getEmployeeNames(userEmail) {
  const fichierDestinationId = "1Dn5HfoYTegM9nJR4LKfrqOJ65CgJ_1crtiNFFUd2LMU";
  const feuille = SpreadsheetApp.openById(fichierDestinationId).getSheetByName("liste collaborateurs");
  if (!feuille) return null;

  const donnees = feuille.getDataRange().getValues();
  const headers = donnees[0];
  const colEmail = headers.indexOf("Email");
  const colPrenom = headers.indexOf("Prénom");
  const colNom = headers.indexOf("Nom");
  const colNomComplet = headers.indexOf("Nom Complet");
  const colUserID = headers.indexOf("User ID");

  if (colEmail === -1) return null;

  for (let i = 1; i < donnees.length; i++) {
    const ligne = donnees[i];
    const email = (ligne[colEmail] || "").toLowerCase().trim();
    if (email === userEmail.toLowerCase().trim()) {
      return {
        prenom: ligne[colPrenom] || "",
        nom: ligne[colNom] || "",
        nomComplet: ligne[colNomComplet] || ((ligne[colPrenom] || "") + " " + (ligne[colNom] || "")).trim(),
        userID: ligne[colUserID] || ""
      };
    }
  }

  return null;
}

function getEmployeeAdditionalInfo(userEmail) {
  const fichierDestinationId = "1Dn5HfoYTegM9nJR4LKfrqOJ65CgJ_1crtiNFFUd2LMU";
  const feuille = SpreadsheetApp.openById(fichierDestinationId).getSheetByName("liste collaborateurs");
  if (!feuille) return null;

  const donnees = feuille.getDataRange().getValues();
  const headers = donnees[0];

  const colEmail = headers.indexOf("Email");
  const colPrenom = headers.indexOf("Prénom");
  const colNom = headers.indexOf("Nom");
  const colNomComplet = headers.indexOf("Nom Complet");
  const colUserID = headers.indexOf("User ID");
  const colTelephone = headers.indexOf("Téléphone");
  const colStatut = headers.indexOf("Statut");
  const colService = headers.indexOf("Service");

  if (colEmail === -1) return null;

  for (let i = 1; i < donnees.length; i++) {
    const ligne = donnees[i];
    const email = (ligne[colEmail] || "").toLowerCase().trim();
    if (email === userEmail.toLowerCase().trim()) {
      const prenom = ligne[colPrenom] || "";
      const nom = ligne[colNom] || "";
      const nomComplet = ligne[colNomComplet] || `${prenom} ${nom}`.trim();
      return {
        prenom,
        nom,
        nomComplet,
        userID: ligne[colUserID] || "",
        telephone: ligne[colTelephone] || "",
        statut: ligne[colStatut] || "",
        service: ligne[colService] || ""
      };
    }
  }

  return null;
}

function getEmployeeLifecycleInfo_(userEmail) {
  const fichierDestinationId = "1Dn5HfoYTegM9nJR4LKfrqOJ65CgJ_1crtiNFFUd2LMU";
  const feuille = SpreadsheetApp.openById(fichierDestinationId).getSheetByName("liste collaborateurs");
  if (!feuille) return null;

  const donnees = feuille.getDataRange().getValues();
  const headers = donnees[0] || [];
  const colEmail = headers.indexOf("Email");
  const colStatut = headers.indexOf("Statut");
  const colEntree = headers.indexOf("Date entrée");
  const colSortie = headers.indexOf("Date Sortie");

  if (colEmail === -1) return null;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 1; i < donnees.length; i++) {
    const ligne = donnees[i];
    const email = (ligne[colEmail] || "").toString().toLowerCase().trim();
    if (email !== userEmail.toLowerCase().trim()) continue;

    const statut = (ligne[colStatut] || "").toString().trim();
    const statutNorm = statut.toLowerCase();
    const entree = ligne[colEntree] ? new Date(ligne[colEntree]) : new Date("2000-01-01");
    const sortie = ligne[colSortie] ? new Date(ligne[colSortie]) : new Date("2100-01-01");

    entree.setHours(0, 0, 0, 0);
    sortie.setHours(0, 0, 0, 0);

    const groupe =
      statutNorm.includes("indépend") || statutNorm.includes("indep") ? "independants" :
      statutNorm.includes("stag") ? "stagiaires" :
      (statutNorm.includes("apprenti") || statutNorm.includes("altern")) ? "alternants" :
      "salaries";

    return {
      statut,
      groupe,
      actif: entree <= today && sortie >= today
    };
  }

  return null;
}

function buildLogoutReminderMessage_(displayName, todayStr, logoutTimeStr, remarkTag) {
  const motif = remarkTag || "rappel de déconnexion";
  return (
    `Bonjour ${displayName},\n\n` +
    `Vous semblez toujours connecté(e) sur KIOSKSIGN aujourd'hui (${todayStr}). ` +
    `Un rappel vous est envoyé à ${logoutTimeStr} : ${motif}.\n\n` +
    `Merci de vous déconnecter manuellement si votre shift est terminé afin d'éviter les incohérences de suivi horaire.\n\n` +
    `Message automatique KIOSKSIGN`
  );
}

function buildLogoutReminderWebhookMessage_(displayName, todayStr, logoutTimeStr, remarkTag) {
  return (
    `⏰ *Rappel KIOSKSIGN*\n` +
    `${displayName}, vous semblez encore connecté(e) le *${todayStr}*.\n` +
    `Rappel envoyé à *${logoutTimeStr}* : ${remarkTag}.\n` +
    `Merci de vous déconnecter manuellement si votre shift est terminé.`
  );
}

function showUI() {
  var sheet = ensureSheetAndHeaders();
  var user = Session.getActiveUser();
  var email = user.getEmail().trim().toLowerCase();

  var names = getEmployeeNames(email);
  if (!names) {
    names = { prenom: "", nom: "", nomComplet: "" };
  }
  var firstName = names.prenom;
  var lastName = names.nom;

  var date = new Date();
  var formattedDate = Utilities.formatDate(date, "Europe/Paris", "dd/MM/yyyy").trim();
  var time = Utilities.formatDate(date, "Europe/Paris", "HH:mm:ss").trim();

  var sheetData = sheet.getDataRange().getValues();
  var idxEmail = getHeaderIndex(sheet, "Email") - 1;
  var idxDate = getHeaderIndex(sheet, "Date") - 1;
  var idxHeureArrivee = getHeaderIndex(sheet, "Heure d'arrivée") - 1;
  var idxHeureDepart = getHeaderIndex(sheet, "Heure de départ") - 1;

  var lastOpenEntry = null;
  for (var i = sheetData.length - 1; i >= 1; i--) {
    var row = sheetData[i];
    var rowEmail = row[idxEmail].toString().trim().toLowerCase();
    var rowDate = formatCellDate(row[idxDate]);
    if (rowEmail === email && rowDate === formattedDate && (!row[idxHeureDepart] || row[idxHeureDepart].toString().trim() === "")) {
      lastOpenEntry = row;
      break;
    }
  }

  var loginTime = lastOpenEntry ? roundTimeTo5Minutes(parseTime(lastOpenEntry[idxHeureArrivee])) : "";
  var signatureURL = "";
  var logoutTime = "";
  var location = "Nanterre Pesaro";
  var idxSignatureURL = getHeaderIndex(sheet, "Signature URL") - 1;
  var idxLocation = getHeaderIndex(sheet, "Localisation") - 1;
  var idxRemarks = getHeaderIndex(sheet, "Remarques") - 1;
  var idxIPLogout = getHeaderIndex(sheet, "Adresse IP Logout") - 1;
  var ip = "";
  if (lastOpenEntry) {
    signatureURL = lastOpenEntry[idxSignatureURL] ? lastOpenEntry[idxSignatureURL].toString().trim() : "";
    location = lastOpenEntry[idxLocation] ? lastOpenEntry[idxLocation].toString().trim() : "Nanterre Pesaro";
    ip = lastOpenEntry[idxIPLogout] ? lastOpenEntry[idxIPLogout].toString().trim() : "";
  }

  var template = HtmlService.createTemplateFromFile("FormHTML");
  template.userEmail = email;
  template.firstName = firstName;
  template.lastName = lastName;
  template.formattedDate = formattedDate;
  template.time = time;
  template.existingEntry = lastOpenEntry;
  template.loginTime = loginTime;
  template.signatureURL = signatureURL;
  template.logoutTime = logoutTime;
  template.location = location;
  template.remarks = lastOpenEntry ? (lastOpenEntry[idxRemarks] || "") : "";
  template.ip = ip;
  return template.evaluate().setSandboxMode(HtmlService.SandboxMode.NATIVE);
}

function logUser(action, signature, ipAddress, userEmail, firstName, lastName, shift, location, remarks) {
  var sheet = ensureSheetAndHeaders();
  var date = new Date();
  var time = Utilities.formatDate(date, "Europe/Paris", "HH:mm:ss").trim();
  var formattedDate = Utilities.formatDate(date, "Europe/Paris", "dd/MM/yyyy").trim();

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  function getIndex(name) {
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].toString().trim() === name) return i;
    }
    return -1;
  }

  var idxDate = getIndex("Date");
  var idxShift = getIndex("Shift");
  var idxLocalisation = getIndex("Localisation");
  var idxEmail = getIndex("Email");
  var idxPrenom = getIndex("Prénom");
  var idxNom = getIndex("Nom");
  var idxNomComplet = getIndex("Nom Complet");
  var idxUserID = getIndex("User ID");
  var idxHeureArrivee = getIndex("Heure d'arrivée");
  var idxHeureDepart = getIndex("Heure de départ");
  var idxNbHeures = getIndex("Nb heures");
  var idxSignatureURL = getIndex("Signature URL");
  var idxHorodatageSignature = getIndex("Horodatage signature");
  var idxIPLogin = getIndex("Adresse IP Login");
  var idxIPLogout = getIndex("Adresse IP Logout");
  var idxRemarques = getIndex("Remarques");

  const expectedIP = "193.248.49.107";
  var data = sheet.getDataRange().getValues();

  // --- Connexion (Login) ---
  if (action === 'in') {

    // Vérification IP IFFP
    if (location === "Nanterre Pesaro" && ipAddress !== expectedIP) {
      return {
        message: "Connexion refusée : vous n'êtes pas connecté au WIFI IFFP (Nanterre Pesaro).",
        loggedIn: false,
        confirmation: { userEmail, firstName, lastName, shift: "", loginTime: "", date: formattedDate, ip: ipAddress, remarks, location }
      };
    }

    // Hors site sans remarque
    if (location !== "Nanterre Pesaro" && ipAddress !== expectedIP && (!remarks || remarks.trim() === "")) {
      return {
        message: "Merci de saisir une remarque obligatoire pour justifier votre connexion hors site.",
        loggedIn: false,
        confirmation: { userEmail, firstName, lastName, shift: "", loginTime: "", date: formattedDate, ip: ipAddress, remarks, location }
      };
    }

    // Ajout de la nouvelle ligne
    var roundedTime = roundTimeTo5Minutes(time);
    sheet.appendRow([
      formattedDate, "", location, userEmail, firstName, lastName, "", "", roundedTime,
      "", "", "", "", ipAddress, "", remarks, ""
    ]);
    SpreadsheetApp.flush();

    // Compléter les infos Nom Complet et UserID
    var lastRow = sheet.getLastRow();
    var currentValues = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    var nomCompletActuel = currentValues[idxNomComplet];
    var userIDActuel = currentValues[idxUserID];

    if (!nomCompletActuel || !userIDActuel) {
      var additionalInfo = getEmployeeAdditionalInfo(userEmail);
      if (additionalInfo) {
        if ((!nomCompletActuel || nomCompletActuel.toString().trim() === "") && additionalInfo.nomComplet) {
          sheet.getRange(lastRow, idxNomComplet + 1).setValue(additionalInfo.nomComplet);
        }
        if ((!userIDActuel || userIDActuel.toString().trim() === "") && additionalInfo.userID) {
          sheet.getRange(lastRow, idxUserID + 1).setValue(additionalInfo.userID);
        }
      }
    }

    return {
      message: 'Connexion enregistrée.',
      loggedIn: true,
      confirmation: {
        userEmail, firstName, lastName, shift: "",
        loginTime: formatTimeStr(roundedTime),
        date: formattedDate, ip: ipAddress, remarks, location
      }
    };

  // --- Déconnexion (Logout) ---
  } else if (action === 'out') {
    var lastEntryRow = -1;
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      if (
        row[idxEmail].toString().trim().toLowerCase() === userEmail.trim().toLowerCase() &&
        formatCellDate(row[idxDate]) === formattedDate &&
        (!row[idxHeureDepart] || row[idxHeureDepart].toString().trim() === "")
      ) {
        lastEntryRow = i;
        break;
      }
    }

    if (lastEntryRow === -1) {
      return { message: 'Aucune connexion à clôturer trouvée.', loggedIn: false };
    }

    var loginTimeStr = parseTime(data[lastEntryRow][idxHeureArrivee]);
    var roundedLoginTime = roundTimeTo5Minutes(loginTimeStr);
    var roundedLogoutTime = roundTimeTo5Minutes(time);
    var loginMinutes = timeStrToMinutes(roundedLoginTime);
    var logoutMinutes = timeStrToMinutes(roundedLogoutTime);
    let diffHours = (logoutMinutes - loginMinutes) / 60;
    let nbHours = Math.round(diffHours / 0.05) * 0.05;
    nbHours = nbHours.toFixed(2).replace(".", ",");

    const formattedTimestamp = Utilities.formatDate(date, "Europe/Paris", "dd-MM-yyyy HH'h'mm");

    sheet.getRange(lastEntryRow + 1, idxHeureDepart + 1).setValue(roundedLogoutTime);
    sheet.getRange(lastEntryRow + 1, idxNbHeures + 1).setValue(nbHours);

    var folder = getSignatureFolder();
    if (!folder) return { message: 'Erreur : Impossible de créer ou de trouver le dossier de signatures.', loggedIn: false };

    var fileName = formattedTimestamp + " - " + lastName.toUpperCase() + " - sign emargemt.png";
    var blob = Utilities.newBlob(Utilities.base64Decode(signature.split(",")[1]), "image/png", fileName);
    var file = folder.createFile(blob);

    sheet.getRange(lastEntryRow + 1, idxSignatureURL + 1).setValue(file.getUrl());
    sheet.getRange(lastEntryRow + 1, idxHorodatageSignature + 1).setValue(formattedTimestamp);
    sheet.getRange(lastEntryRow + 1, idxIPLogout + 1).setValue(ipAddress);
    sheet.getRange(lastEntryRow + 1, idxLocalisation + 1).setValue(location);
    sheet.getRange(lastEntryRow + 1, idxRemarques + 1).setValue(remarks);

    return {
      message: 'Déconnexion enregistrée.',
      loggedIn: false,
      confirmation: {
        userEmail, firstName, lastName, shift: "", location, date: formattedDate,
        loginTime: formatTimeStr(roundedLoginTime),
        logoutTime: formatTimeStr(roundedLogoutTime),
        ip: ipAddress, remarks
      }
    };
  }
}

function filtrerMoisActuel() {
  var sheet = ensureSheetAndHeaders();
  var data = sheet.getDataRange().getValues();
  var idxDate = getHeaderIndex(sheet, "Date") - 1;

  if (idxDate < 0) {
    Logger.log("❌ Colonne 'Date' introuvable.");
    return;
  }

  var today = new Date();
  var moisActuel = today.getMonth(); // 0 = janvier
  var anneeActuelle = today.getFullYear();

  var nouvellesDonnees = [data[0]]; // Garder les en-têtes

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateCell = row[idxDate];
    var parsedDate = normalizeDate(dateCell);

    if (
      parsedDate &&
      parsedDate.getMonth() === moisActuel &&
      parsedDate.getFullYear() === anneeActuelle
    ) {
      nouvellesDonnees.push(row);
    }
  }

  if (nouvellesDonnees.length > 1) {
    sheet.clearContents();
    sheet.getRange(1, 1, nouvellesDonnees.length, nouvellesDonnees[0].length).setValues(nouvellesDonnees);
    Logger.log("✅ Données filtrées : " + (nouvellesDonnees.length - 1) + " lignes gardées pour " + (moisActuel + 1) + "/" + anneeActuelle);
  } else {
    Logger.log("⚠️ Aucune donnée pour le mois actuel (" + (moisActuel + 1) + "/" + anneeActuelle + ")");
  }
}

function normalizeDate(cellValue) {
  if (cellValue instanceof Date) {
    return cellValue;
  }

  if (typeof cellValue === "string") {
    cellValue = cellValue.trim();

    var parts = cellValue.split("/");
    if (parts.length === 3) {
      var day = parseInt(parts[0], 10);
      var month = parseInt(parts[1], 10) - 1;
      var year = parseInt(parts[2], 10);

      if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
        return new Date(year, month, day);
      }
    }
  }

  return null; // Pas une date valide
}
/**
 * ✅ Met à jour "Nom Complet" et "User ID" dans "Présences"
 * à partir de la feuille "liste collaborateurs" du fichier externe.
 * + Nettoyage emails et optimisation massive (1 seul setValues)
 */
function majNomCompletEtUserID() {
  const start = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuillePresences = ss.getSheetByName("Présences");
  const fichierEmployes = SpreadsheetApp.openById("1Dn5HfoYTegM9nJR4LKfrqOJ65CgJ_1crtiNFFUd2LMU");
  const feuilleEmployes = fichierEmployes.getSheetByName("liste collaborateurs");

  if (!feuillePresences || !feuilleEmployes) {
    Logger.log("❌ Feuille manquante (Présences ou liste collaborateurs)");
    return;
  }

  Logger.log(`📘 Connexion OK à "${fichierEmployes.getName()}"`);
  const dataPresences = feuillePresences.getDataRange().getValues();
  const dataEmployes = feuilleEmployes.getDataRange().getValues();

  const headersPresences = dataPresences[0];
  const headersEmployes = dataEmployes[0];

  const idxEmailPres = headersPresences.indexOf("Email");
  const idxNomCompletPres = headersPresences.indexOf("Nom Complet");
  const idxUserIDPres = headersPresences.indexOf("User ID");

  const idxEmailEmp = headersEmployes.indexOf("Email");
  const idxNomEmp = headersEmployes.indexOf("Nom");
  const idxPrenomEmp = headersEmployes.indexOf("Prénom");
  const idxUserIDEmp = headersEmployes.indexOf("User ID");
  const idxNomCompletEmp = headersEmployes.indexOf("Nom Complet");

  if (
    [idxEmailPres, idxNomCompletPres, idxUserIDPres,
     idxEmailEmp, idxNomEmp, idxPrenomEmp, idxUserIDEmp].includes(-1)
  ) {
    Logger.log("❌ Colonnes manquantes. Vérifie les noms exacts dans les deux feuilles.");
    return;
  }

  // === 1️⃣ Dictionnaire rapide {email : {nomComplet, userID}} ===
  const employesMap = {};
  for (let i = 1; i < dataEmployes.length; i++) {
    const row = dataEmployes[i];
    const email = (row[idxEmailEmp] || "").toString().toLowerCase().trim();
    if (!email) continue;
    const nomComplet =
      row[idxNomCompletEmp] ||
      `${(row[idxPrenomEmp] || "").trim()} ${(row[idxNomEmp] || "").trim()}`.trim();
    employesMap[email] = {
      nomComplet,
      userID: row[idxUserIDEmp] || ""
    };
  }

  // === 2️⃣ Préparation de la mise à jour ===
  let majCount = 0;
  const newData = [...dataPresences];
  for (let i = 1; i < dataPresences.length; i++) {
    const emailPres = (dataPresences[i][idxEmailPres] || "").toString().toLowerCase().trim();
    if (!emailPres) continue;
    const emp = employesMap[emailPres];
    if (emp) {
      const oldNomComplet = dataPresences[i][idxNomCompletPres];
      const oldUserID = dataPresences[i][idxUserIDPres];
      if (emp.nomComplet !== oldNomComplet || emp.userID !== oldUserID) {
        newData[i][idxNomCompletPres] = emp.nomComplet;
        newData[i][idxUserIDPres] = emp.userID;
        majCount++;
      }
    }
  }

  // === 3️⃣ Mise à jour groupée ===
  feuillePresences.getRange(1, 1, newData.length, newData[0].length).setValues(newData);

  const duration = ((new Date() - start) / 1000).toFixed(2);
  Logger.log(`✅ ${majCount} lignes mises à jour depuis "${feuilleEmployes.getName()}" (${duration}s)`);
}


/**
 * 🧪 TEST — Vérifie la connexion et la recherche dans la feuille externe "liste collaborateurs"
 * - Vérifie l’accès au fichier
 * - Vérifie la présence des colonnes : ID, Nom Complet, Email, Nom, Prénom
 * - Recherche un utilisateur par email (par défaut, l'utilisateur connecté)
 */
function testConnexionCollaborateursExterne(emailTest) {
  try {
    // === CONFIG ===
    const fichierDestinationId = "1Dn5HfoYTegM9nJR4LKfrqOJ65CgJ_1crtiNFFUd2LMU";
    const nomFeuille = "liste collaborateurs";

    // === 1️⃣ OUVERTURE DU FICHIER EXTERNE ===
    const fichier = SpreadsheetApp.openById(fichierDestinationId);
    const feuille = fichier.getSheetByName(nomFeuille);
    if (!feuille) {
      Logger.log(`❌ Feuille "${nomFeuille}" introuvable dans ${fichier.getName()}`);
      return;
    }

    Logger.log(`📘 Connexion OK au fichier : ${fichier.getName()}`);
    Logger.log(`🔗 Lien : https://docs.google.com/spreadsheets/d/${fichierDestinationId}/edit`);

    // === 2️⃣ CHARGEMENT DES DONNÉES ===
    const data = feuille.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());

    // === 3️⃣ CONTRÔLE DES COLONNES ESSENTIELLES ===
    const champsRequis = ["Email", "Nom Complet", "User ID", "Nom", "Prénom"];
    const manquants = champsRequis.filter(c => !headers.includes(c));
    if (manquants.length > 0) {
      Logger.log(`⚠️ Colonnes manquantes : ${manquants.join(", ")}`);
    } else {
      Logger.log(`✅ Colonnes détectées : ${champsRequis.join(", ")}`);
    }

    // === 4️⃣ RECHERCHE PAR EMAIL ===
    const colEmail = headers.indexOf("Email");
    const colNomComplet = headers.indexOf("Nom Complet");
    const colPrenom = headers.indexOf("Prénom");
    const colNom = headers.indexOf("Nom");
    const colID = headers.indexOf("User ID");

    const email = (emailTest || Session.getActiveUser().getEmail() || "").toLowerCase().trim();
    if (!email) {
      Logger.log("⚠️ Aucun email fourni ni utilisateur détecté.");
      return;
    }

    let trouve = null;
    for (let i = 1; i < data.length; i++) {
      const ligne = data[i];
      const mail = (ligne[colEmail] || "").toString().toLowerCase().trim();
      if (mail === email) {
        trouve = {
          email: mail,
          nom: ligne[colNom] || "",
          prenom: ligne[colPrenom] || "",
          nomComplet: ligne[colNomComplet] || `${ligne[colPrenom]} ${ligne[colNom]}`.trim(),
          userID: ligne[colID] || ""
        };
        break;
      }
    }

    if (trouve) {
      Logger.log(`✅ Utilisateur trouvé :`);
      Logger.log(`  👤 Nom complet : ${trouve.nomComplet}`);
      Logger.log(`  🪪 User ID : ${trouve.userID}`);
      Logger.log(`  ✉️ Email : ${trouve.email}`);
      Logger.log(`  🧍‍♂️ Prénom : ${trouve.prenom} | Nom : ${trouve.nom}`);
    } else {
      Logger.log(`❌ Aucun utilisateur trouvé pour ${email}`);
    }

  } catch (err) {
    Logger.log(`❌ Erreur testConnexionCollaborateursExterne : ${err.message}`);
  }
}

/**
 * 🔁 Complète les infos manquantes dans "Présences"
 * (Prénom, Nom, Nom Complet, User ID)
 * et recalcule les heures pour les lignes complètes.
 */
function completerInfosPresences() {
  try {
    const fichierDestinationId = "1Dn5HfoYTegM9nJR4LKfrqOJ65CgJ_1crtiNFFUd2LMU";
    const feuilleEmployes = SpreadsheetApp.openById(fichierDestinationId).getSheetByName("liste collaborateurs");
    const feuillePresences = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Présences");
    if (!feuilleEmployes || !feuillePresences) throw new Error("❌ Feuille manquante (Présences ou liste collaborateurs).");

    const normalize = s => (s || "").toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\s\-_]+/g, "");
    const findCol = (headers, possibles) => {
      const norm = headers.map(normalize);
      for (let alias of possibles) {
        const idx = norm.indexOf(normalize(alias));
        if (idx !== -1) return idx;
      }
      return -1;
    };

    // === CHARGEMENT HRMS ===
    const dataEmp = feuilleEmployes.getDataRange().getValues();
    const headersEmp = dataEmp[0];
    const idxEmailEmp = findCol(headersEmp, ["email", "mail"]);
    const idxPrenomEmp = findCol(headersEmp, ["prenom", "prénom"]);
    const idxNomEmp = findCol(headersEmp, ["nom"]);
    const idxNomCompletEmp = findCol(headersEmp, ["nom complet", "nomcomplet"]);
    const idxUserIDEmp = findCol(headersEmp, ["userid", "id utilisateur", "id"]);

    // === CHARGEMENT Présences ===
    const dataPres = feuillePresences.getDataRange().getValues();
    const headersPres = dataPres[0];
    const idxEmailPres = findCol(headersPres, ["email", "mail"]);
    const idxPrenomPres = findCol(headersPres, ["prenom", "prénom"]);
    const idxNomPres = findCol(headersPres, ["nom"]);
    const idxNomCompletPres = findCol(headersPres, ["nom complet", "nomcomplet"]);
    const idxUserIDPres = findCol(headersPres, ["userid", "user id"]);
    const idxHeureArrivee = findCol(headersPres, ["heure d'arrivée"]);
    const idxHeureDepart = findCol(headersPres, ["heure de départ"]);
    const idxNbHeures = findCol(headersPres, ["nb heures"]);
    const idxNbHValid = findCol(headersPres, ["nb h valid"]);

    if (idxEmailEmp === -1 || idxEmailPres === -1) throw new Error("⚠️ Colonnes Email introuvables dans HRMS ou Présences.");

    // === Dictionnaire HRMS ===
    const employesMap = {};
    for (let i = 1; i < dataEmp.length; i++) {
      const row = dataEmp[i];
      const email = (row[idxEmailEmp] || "").toString().trim().toLowerCase();
      if (!email) continue;
      const prenom = row[idxPrenomEmp] || "";
      const nom = row[idxNomEmp] || "";
      const nomComplet = row[idxNomCompletEmp] || `${prenom} ${nom}`.trim();
      const userID = idxUserIDEmp !== -1 ? (row[idxUserIDEmp] || "") : "";
      employesMap[email] = { prenom, nom, nomComplet, userID };
    }

    // === Mise à jour + recalcul ===
    let majCount = 0;
    let recalculCount = 0;
    const newData = [...dataPres];

    for (let i = 1; i < dataPres.length; i++) {
      const email = (dataPres[i][idxEmailPres] || "").toString().trim().toLowerCase();
      if (!email) continue;
      const emp = employesMap[email];
      if (!emp) continue;

      let maj = false;
      if (idxPrenomPres !== -1 && (!dataPres[i][idxPrenomPres] || dataPres[i][idxPrenomPres].toString().trim() === "")) {
        newData[i][idxPrenomPres] = emp.prenom;
        maj = true;
      }
      if (idxNomPres !== -1 && (!dataPres[i][idxNomPres] || dataPres[i][idxNomPres].toString().trim() === "")) {
        newData[i][idxNomPres] = emp.nom;
        maj = true;
      }
      if (idxNomCompletPres !== -1 && (!dataPres[i][idxNomCompletPres] || dataPres[i][idxNomCompletPres].toString().trim() === "")) {
        newData[i][idxNomCompletPres] = emp.nomComplet;
        maj = true;
      }
      if (idxUserIDPres !== -1 && (!dataPres[i][idxUserIDPres] || dataPres[i][idxUserIDPres].toString().trim() === "")) {
        newData[i][idxUserIDPres] = emp.userID;
        maj = true;
      }

      // --- Recalcul des heures si possible ---
      const hArr = dataPres[i][idxHeureArrivee];
      const hDep = dataPres[i][idxHeureDepart];
      if (hArr && hDep) {
        const { nbHeures, nbHValid } = calculHeures(hArr, hDep);
        if (nbHeures && nbHValid) {
          newData[i][idxNbHeures] = nbHeures;
          newData[i][idxNbHValid] = nbHValid;
          recalculCount++;
        }
      }

      if (maj) majCount++;
    }

    feuillePresences.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
    Logger.log(`✅ ${majCount} lignes complétées + ${recalculCount} lignes recalculées.`);

  } catch (err) {
    Logger.log("🚨 Erreur completerInfosPresences : " + err.message);
  }
}

/**
 * 🧮 Calcule Nb heures et Nb H Valid à partir des heures données
 */
function calculHeures(hArr, hDep) {
  try {
    const arrivee = convertToDate(hArr);
    const depart = convertToDate(hDep);
    if (!arrivee || !depart) return {};

    const diffHeures = (depart - arrivee) / (1000 * 60 * 60);
    if (diffHeures <= 0 || diffHeures > 15) return {};

    const nbHeures = Math.round(diffHeures * 4) / 4;
    const nbHValid = Math.max(nbHeures - 0.25, 0);

    return {
      nbHeures: nbHeures.toFixed(2).replace(".", ","),
      nbHValid: nbHValid.toFixed(2).replace(".", ",")
    };
  } catch {
    return {};
  }
}

/** Convertit une valeur texte ou Date en objet Date */
function convertToDate(value) {
  if (value instanceof Date) return value;
  const str = value.toString().trim().replace("h", ":");
  const [h, m] = str.split(":").map(v => parseInt(v, 10) || 0);
  const d = new Date();
  d.setHours(h);
  d.setMinutes(m);
  d.setSeconds(0);
  return d;
}

/**
 * 🔄 Recalcule les heures dans la feuille "Présences"
 * - "Nb heures" = (Heure de départ - Heure d'arrivée), arrondi au 0.25 le plus proche
 * - "Nb H Valid" = "Nb heures" - 0.25
 */
function recalculerHeuresPresence() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Présences");
    if (!sheet) throw new Error("Feuille 'Présences' introuvable.");

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim().toLowerCase());

    const idxArrivee = headers.indexOf("heure d'arrivée");
    const idxDepart = headers.indexOf("heure de départ");
    const idxNbHeures = headers.indexOf("nb heures");
    const idxNbHValid = headers.indexOf("nb h valid");

    if ([idxArrivee, idxDepart, idxNbHeures, idxNbHValid].includes(-1)) {
      throw new Error("Certaines colonnes sont introuvables (Heure d'arrivée, Heure de départ, Nb heures, Nb H Valid).");
    }

    let majCount = 0;
    const newData = [...data];

    for (let i = 1; i < data.length; i++) {
      const hArr = data[i][idxArrivee];
      const hDep = data[i][idxDepart];
      if (!hArr || !hDep) continue;

      const heureArrivee = convertToDate(hArr);
      const heureDepart = convertToDate(hDep);
      if (!heureArrivee || !heureDepart) continue;

      const diffHeures = (heureDepart - heureArrivee) / (1000 * 60 * 60);
      if (diffHeures <= 0 || diffHeures > 15) continue; // ignore les incohérences

      // ⏱ Arrondi au quart d’heure
      const nbHeures = Math.round(diffHeures * 4) / 4;
      const nbHValid = Math.max(nbHeures - 0.25, 0); // jamais négatif

      const nbHeuresStr = nbHeures.toFixed(2).replace(".", ",");
      const nbHValidStr = nbHValid.toFixed(2).replace(".", ",");

      // Vérifie si une mise à jour est nécessaire
      const currentNb = data[i][idxNbHeures];
      const currentValid = data[i][idxNbHValid];

      if (currentNb !== nbHeuresStr || currentValid !== nbHValidStr) {
        newData[i][idxNbHeures] = nbHeuresStr;
        newData[i][idxNbHValid] = nbHValidStr;
        majCount++;
      }
    }

    // ✍️ Mise à jour globale
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
    Logger.log(`✅ ${majCount} lignes recalculées (Nb heures + Nb H Valid).`);

  } catch (err) {
    Logger.log("🚨 Erreur recalculerHeuresPresence : " + err.message);
  }
}

const LOGOUT_COLLAB_CONFIG = {
  SOURCE_SHEET_NAME: "Présences",
  SENDER_NAME: "KioskSign IFFP - NOREPLY",
  REPLY_TO: "noreply@iffp.school",
  FROM: "noreply@iffp.school"
};

// Triggers -> wrappers
function logoutAllForgottenShifts_1335() {
  logoutAllForgottenShiftsCollaborateurs_(
    "13:35",
    "Déconnexion automatique à la pause déjeuner",
    { maxLoginTime: "12:59" }
  );
}

function logoutAllForgottenShifts_1935() {
  logoutAllForgottenShiftsCollaborateurs_("19:35", "Déconnexion automatique de fin de journée");
}

/**
 * Déconnecte automatiquement les shifts non clôturés du JOUR
 * à une heure cible (13:35 ou 19:35) + email de prévention.
 *
 * Règle : si Date = aujourd'hui ET Heure de départ vide => on fixe Heure de départ = logoutTimeStr
 */
/**
 * Déconnecte automatiquement les shifts non clôturés du JOUR
 * à une heure cible (13:35 ou 19:35) + email de prévention.
 *
 * Règle : si Date = aujourd'hui ET Heure de départ vide => on fixe Heure de départ = logoutTimeStr
 * + écrit "LOGOUT AUTOMATIQUE Bot Kiosksign" dans Remarques (avec tag)
 * + peut filtrer les lignes selon l'heure de connexion (ex: ne fermer que les shifts du matin)
 */
function logoutAllForgottenShiftsCollaborateurs_(logoutTimeStr, remarkTag, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOGOUT_COLLAB_CONFIG.SOURCE_SHEET_NAME);
  if (!sheet) throw new Error('Feuille introuvable : "' + LOGOUT_COLLAB_CONFIG.SOURCE_SHEET_NAME + '"');

  const now = new Date();

  // Pas le week-end
  const day = now.getDay();
  if (day === 0 || day === 6) return;

  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(now, tz, "dd/MM/yyyy");

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  // Colonnes via en-têtes (robuste)
  const colDate        = getHeaderIndex(sheet, "Date");            // 1-based
  const colEmail       = getHeaderIndex(sheet, "Email");
  const colPrenom      = getHeaderIndex(sheet, "Prénom");
  const colNom         = getHeaderIndex(sheet, "Nom");
  const colNomComplet  = getHeaderIndex(sheet, "Nom Complet");
  const colHeureArrivee = getHeaderIndex(sheet, "Heure d'arrivée");
  const colHeureDepart = getHeaderIndex(sheet, "Heure de départ");
  const colRemarques   = getHeaderIndex(sheet, "Remarques");

  if ([colDate, colEmail, colHeureArrivee, colHeureDepart, colRemarques].includes(-1)) {
    throw new Error("Colonnes requises introuvables (Date / Email / Heure d'arrivée / Heure de départ / Remarques).");
  }

  const maxLoginMinutes = options && options.maxLoginTime
    ? timeStrToMinutes(options.maxLoginTime + ":00")
    : null;

  // Anti-spam : 1 mail max / personne / jour / créneau
  const props = PropertiesService.getScriptProperties();
  const employeeInfoCache = {};

  // On batch les écritures pour éviter setValue dans la boucle
  const updates = []; // {r, c, v}
  const emailsToSend = []; // {email, subject, text, html, key}
  const reminderWebhooksToSend = []; // {webhook, message, key}

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // ✅ Date normalisée (gère Date objet OU string)
    const rowDateStr = formatCellDate(row[colDate - 1]);
    if (rowDateStr !== todayStr) continue;

    const rawDepart = row[colHeureDepart - 1];
    const departStr = (rawDepart === null || rawDepart === undefined) ? "" : rawDepart.toString().trim();

    // ✅ Non clôturé = vide ou espaces
    if (departStr !== "") continue;

    const loginTimeStr = parseTime(row[colHeureArrivee - 1]);
    if (maxLoginMinutes !== null) {
      const loginMinutes = timeStrToMinutes(loginTimeStr);
      if (loginMinutes > maxLoginMinutes) continue;
    }

    const email = (row[colEmail - 1] || "").toString().trim();
    const prenom = colPrenom !== -1 ? (row[colPrenom - 1] || "").toString().trim() : "";
    const nom = colNom !== -1 ? (row[colNom - 1] || "").toString().trim() : "";
    const nomComplet = colNomComplet !== -1 ? (row[colNomComplet - 1] || "").toString().trim() : "";
    const displayName = (prenom || nom) ? `${prenom} ${nom}`.trim() : (nomComplet || "Bonjour");

    const employeeInfo = email
      ? (employeeInfoCache[email] !== undefined
          ? employeeInfoCache[email]
          : (employeeInfoCache[email] = getEmployeeLifecycleInfo_(email)))
      : null;
    const isActiveEmployee = employeeInfo && employeeInfo.actif && employeeInfo.groupe === "salaries";

    if (isActiveEmployee) {
      const reminderKey = `AUTO_LOGOUT_REMINDER|${email}|${todayStr}|${logoutTimeStr}`;
      if (props.getProperty(reminderKey)) continue;

      const webhookPerso = typeof chercherWebhookPersonnel === "function"
        ? chercherWebhookPersonnel(nom, email)
        : null;
      const reminderText = buildLogoutReminderMessage_(displayName, todayStr, logoutTimeStr, remarkTag);

      if (webhookPerso) {
        reminderWebhooksToSend.push({
          webhook: webhookPerso,
          message: buildLogoutReminderWebhookMessage_(displayName, todayStr, logoutTimeStr, remarkTag),
          key: reminderKey
        });
      } else if (email) {
        const reminderHtml = `
          <div style="font-family:Arial,sans-serif;color:#111827;line-height:1.45;">
            <p style="margin:0 0 10px 0;">Bonjour ${escapeHtml_(displayName)},</p>
            <p style="margin:0 0 10px 0;">
              Vous semblez toujours connecté(e) sur <strong>KIOSKSIGN</strong> aujourd'hui (<strong>${escapeHtml_(todayStr)}</strong>).
              Un rappel vous est envoyé à <strong>${escapeHtml_(logoutTimeStr)}</strong> : ${escapeHtml_(remarkTag)}.
            </p>
            <p style="margin:0 0 10px 0;">
              Merci de vous déconnecter manuellement si votre shift est terminé afin d'éviter les incohérences de suivi horaire.
            </p>
            <p style="margin:0; color:#6b7280; font-size:12px;">Message automatique KIOSKSIGN</p>
          </div>
        `;

        emailsToSend.push({
          email,
          subject: "[KIOSKSIGN] Rappel de déconnexion",
          text: reminderText,
          html: reminderHtml,
          key: reminderKey
        });
      }

      continue;
    }

    const remarquesExistantes = (row[colRemarques - 1] || "").toString().trim();
    const newRemark =
      (remarquesExistantes ? (remarquesExistantes + " | ") : "") +
      "LOGOUT AUTOMATIQUE Bot Kiosksign" +
      (remarkTag ? (" — " + remarkTag) : "") +
      " (" + logoutTimeStr + ")";

    // Écrit Heure de départ
    updates.push({ r: i + 1, c: colHeureDepart, v: logoutTimeStr });
    // Écrit Remarques
    updates.push({ r: i + 1, c: colRemarques, v: newRemark });

    // Email
    if (email) {
      const key = `AUTO_LOGOUT_SENT|${email}|${todayStr}|${logoutTimeStr}`;
      if (!props.getProperty(key)) {
        const subject = "[KIOSKSIGN] Déconnexion automatique — oubli de déconnexion";

        const html = `
          <div style="font-family:Arial,sans-serif;color:#111827;line-height:1.45;">
            <p style="margin:0 0 10px 0;">Bonjour ${escapeHtml_(displayName)},</p>

            <p style="margin:0 0 10px 0;">
              Nous avons constaté un <strong>oubli de déconnexion</strong> sur KIOSKSIGN aujourd’hui (<strong>${escapeHtml_(todayStr)}</strong>).
              Pour garantir la fiabilité du suivi horaire, nous avons procédé à une <strong>déconnexion automatique</strong> à <strong>${escapeHtml_(logoutTimeStr)}</strong>.
            </p>

            <p style="margin:0 0 10px 0;">
              Cet oubli <strong>perturbe le bon fonctionnement du suivi horaire</strong> et peut générer des incohérences (administratives, reporting, validation des heures).
              Nous le faisons cette fois-ci, mais merci d’être vigilant(e) afin que cela ne se reproduise pas.
            </p>

            <p style="margin:0 0 10px 0;">
              En cas de répétition, cela peut entraîner une <strong>situation administrative</strong> ou des conséquences sur le suivi horaire, et conduire à un <strong>entretien avec la direction</strong>.
            </p>

            <p style="margin:0;">
              Merci pour votre attention.<br>
              <span style="color:#6b7280;font-size:12px;">Message automatique KIOSKSIGN</span>
            </p>
          </div>
        `;

        const text = [
          `Bonjour ${displayName},`,
          ``,
          `Nous avons constaté un oubli de déconnexion sur KIOSKSIGN aujourd’hui (${todayStr}).`,
          `Pour garantir la fiabilité du suivi horaire, nous avons procédé à une déconnexion automatique à ${logoutTimeStr}.`,
          ``,
          `Cet oubli perturbe le bon fonctionnement du suivi horaire et peut générer des incohérences (administratives, reporting, validation des heures).`,
          `Nous le faisons cette fois-ci, mais merci d’être vigilant(e) afin que cela ne se reproduise pas.`,
          ``,
          `En cas de répétition, cela peut entraîner une situation administrative ou des conséquences sur le suivi horaire, et conduire à un entretien avec la direction.`,
          ``,
          `Message automatique KIOSKSIGN`
        ].join("\n");

        emailsToSend.push({ email, subject, text, html, key });
      }
    }
  }

  // Applique les updates en une fois
  for (const u of updates) {
    sheet.getRange(u.r, u.c).setValue(u.v);
  }
  SpreadsheetApp.flush();

  // Envoie les mails après écriture
  for (const m of emailsToSend) {
    MailApp.sendEmail(m.email, m.subject, m.text, {
      htmlBody: m.html,
      name: LOGOUT_COLLAB_CONFIG.SENDER_NAME,
      replyTo: LOGOUT_COLLAB_CONFIG.REPLY_TO,
      from: LOGOUT_COLLAB_CONFIG.FROM
    });
    props.setProperty(m.key, "1");
  }

  for (const m of reminderWebhooksToSend) {
    envoyerMessageWebhook(m.webhook, m.message);
    props.setProperty(m.key, "1");
  }
}

function escapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function ouvrirFormulaireEmargement() {
  const url = "https://script.google.com/a/macros/iffp.school/s/AKfycbzIeTeCMr1ltS51Qnrv_CvQ6bYeBklCkPffXSyNekW5/dev";
  const html = HtmlService.createHtmlOutput(`
    <html><head><script>
      window.open("${url}", "_blank");
      google.script.host.close();
    </script></head><body></body></html>
  `);
  SpreadsheetApp.getUi().showModalDialog(html, "Ouverture du formulaire...");
}


function testLogUserIn() {
  Logger.log("testLogUserIn() démarré");
  var testIP = "127.0.0.1";
  var testEmail = "test@example.com";
  var testFirstName = "TestPrenom";
  var testLastName = "TestNom";
  var testShift = "Après-midi"; // ou "Matin"
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
  var testShift = "Après-midi"; // Doit correspondre au shift du LOGIN.
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
