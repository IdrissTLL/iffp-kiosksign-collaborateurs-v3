/**
 * Parcourt la feuille "Présences" et envoie des notifications via webhook
 * uniquement pour les lignes dont la date correspond à aujourd’hui (format JJ/MM/YYYY).
 */
function traiterPresences() {
  // Récupérer le classeur et les feuilles nécessaires
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const feuille     = ss.getSheetByName("Présences");
  const configSheet = ss.getSheetByName("config");
  if (!feuille || !configSheet) return;

  // Vérifier si les notifications sont activées dans la config
  const activerNotif = chercherValeurDansConfig("activer Notifications");
  if (activerNotif !== "OUI") {
    Logger.log("Notifications désactivées.");
    return;
  }

  // Préparer la date d’aujourd’hui au format JJ/MM/YYYY
  const today    = new Date();
  const todayStr = Utilities.formatDate(today, "Europe/Paris", "dd/MM/yyyy");

  // Lire les en-têtes (ligne 1) et déterminer les index de colonnes par nom
  const headers = feuille
    .getRange(1, 1, 1, feuille.getLastColumn())
    .getValues()[0];
  const getCol = (titre) => headers.indexOf(titre);

  const colNom       = getCol("Nom");
  const colPrenom    = getCol("Prénom");
  const colEmail     = getCol("Email");
  const colDate      = getCol("Date");
  const colArrivee   = getCol("Heure d'arrivée");
  const colDepart    = getCol("Heure de départ");
  const colIPIn      = getCol("Adresse IP Login");
  const colIPOut     = getCol("Adresse IP Logout");
  const colLoc       = getCol("Localisation");
  const colNbHeures  = getCol("Nb heures");
  const colNbHValid  = getCol("Nb H Valid");   // nouveau champ demandé
  const colNotif     = getCol("Msg Notif");    // toujours par nom

  // Vérifier que les colonnes critiques existent
  if ([colNom, colPrenom, colEmail, colDate, colArrivee, colNotif, colNbHValid].includes(-1)) {
    Logger.log("Colonnes manquantes (Nom, Prénom, Email, Date, Heure d'arrivée, Msg Notif ou Nb H Valid).");
    return;
  }

  // Nombre de lignes à traiter (hors en-tête)
  const nbLignes = feuille.getLastRow() - 1;
  if (nbLignes < 1) {
    Logger.log("Aucune donnée à traiter.");
    return;
  }

  // Charger toutes les données sous les en-têtes
  const data = feuille
    .getRange(2, 1, nbLignes, feuille.getLastColumn())
    .getValues();

  // Parcourir chaque ligne de données
  data.forEach((ligne, idx) => {
    const rowNum    = idx + 2;               // numéro de ligne dans la feuille
    const nom       = ligne[colNom];
    const prenom    = ligne[colPrenom];
    const dateRaw   = ligne[colDate];
    const arriveeRaw= ligne[colArrivee];
    const departRaw = ligne[colDepart];
    const ipIn      = ligne[colIPIn];
    const ipOut     = ligne[colIPOut];
    const loc       = ligne[colLoc];
    const nbHeures  = ligne[colNbHeures];
    const nbHValid  = ligne[colNbHValid];
    let notif       = (ligne[colNotif] || "").toString().trim();

    // Ne traiter que si la date est un Date et correspond à aujourd’hui (JJ/MM/YYYY)
    if (!(dateRaw instanceof Date) ||
        Utilities.formatDate(dateRaw, "Europe/Paris", "dd/MM/yyyy") !== todayStr) {
      return; // passer à la ligne suivante
    }

    // Si nom ou prénom manquant, ignorer
    if (!nom || !prenom) return;

    // Journaliser la présence ou absence de webhook
    const webhookPerso = chercherWebhookPersonnel(nom, ligne[colEmail]);
    Logger.log(
      webhookPerso
        ? `✅ Webhook trouvé pour ${prenom} ${nom} : ${webhookPerso}`
        : `❌ Pas de webhook pour ${prenom} ${nom}`
    );

    // Mettre en forme date et heures pour l’affichage
    const date    = formatDateFR(dateRaw);
    const arrivee = formatHeure(arriveeRaw);
    const depart  = departRaw ? formatHeure(departRaw) : "";

    let message   = "";
    let typeNotif = "";

    // Cas : connexion non encore notifiée
    if (notif === "" && arriveeRaw) {
      message   =
        `📌📆✅ *${prenom}*, vous vous êtes connecté(e) le *${date}* à *${arrivee}* sur *${loc}* (${ipIn}).\n` +
        `Respectez bien les horaires planifiés. Merci et excellente journée — L'équipe IFFP — 🤝✨`;
      typeNotif = "OUI";
    }
    // Cas : déconnexion non encore notifiée
    else if (notif === "OUI" && departRaw) {
      message   =
        `📌📆✍🏼 *${prenom}*, vous vous êtes *déconnecté(e)* le *${date}* à *${depart}*.\n` +
        `Connexion à *${arrivee}* (${ipIn}) // Déconnexion à *${depart}* (${ipOut}) ` +
        `soit *${nbHeures} H* (validées : ${nbHValid}).\n` +
        `Merci pour votre présence et votre implication — L'équipe IFFP — 🙏🏼✨`;
      typeNotif = "OUI OUI";
    }
    // Sinon, on ne fait rien (déjà notifié ou pas d’événement pertinent)
    else {
      return;
    }

    // Référence à la cellule "Msg Notif"
    const celluleNotif   = feuille.getRange(rowNum, colNotif + 1);
    const valeurActuelle = celluleNotif.getValue().toString().trim();

    // Si déjà notifié de ce type, on ne renvoie pas
    if (valeurActuelle === typeNotif) {
      Logger.log(`ℹ️ Déjà notifié (${typeNotif}) – ligne ${rowNum}`);
      return;
    }

    // Envoi de la notification si possible
    if (message && webhookPerso) {
      envoyerMessageWebhook(webhookPerso, message);
      celluleNotif.setValue(typeNotif);
      Logger.log(`✅ Notification envoyée pour ${prenom} ${nom} (type : ${typeNotif})`);
    }
    // En l’absence de webhook personnel, alerte au webhook Kiosksign
    else if (!webhookPerso && message) {
      const webhookKiosksign = chercherValeurDansConfig("webhook Kiosksign");
      if (webhookKiosksign) {
        const msgErr =
          `⚠️ Webhook perso manquant pour ${prenom} ${nom}. Merci de le configurer.`;
        envoyerMessageWebhook(webhookKiosksign, msgErr);
        Logger.log(`❌ Erreur : webhook perso manquant pour ${prenom} ${nom}`);
      }
    }
  });
}

/**
 * Recherche et renvoie le webhook personnel associé à un collaborateur.
 * @param {string} nom  – Nom du collaborateur à rechercher (inscrit dans config).
 * @param {string=} email – Email du collaborateur pour une recherche plus fiable.
 * @returns {string|null} – URL du webhook si trouvée et non vide, sinon null.
 */
function chercherWebhookPersonnel(nom, email) {
  // 1️⃣ Vérifier que le paramètre 'nom' est fourni
  if (!nom && !email) {
    return null; // Pas de nom → on ne peut rien chercher
  }

  // 2️⃣ Accéder à la feuille "config" du classeur actif
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("config");
  if (!configSheet) {
    return null; // Pas de feuille config → on arrête
  }

  // 3️⃣ Récupérer toutes les données de la feuille en un tableau 2D
  const data = configSheet.getDataRange().getValues();
  // Doit comporter au moins 4 lignes : 
  // ligne 1 = en-têtes, lignes 2–3 = infos, ligne 4+ = données effectives
  if (data.length < 4) {
    return null;
  }

  // 4️⃣ Identifier la ligne des en-têtes (ligne 1, index 0)
  const normalize = value => (value || "")
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[\s'’-]+/g, "")
    .trim()
    .toUpperCase();

  const cleanEmail = value => (value || "")
    .toString()
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .trim()
    .toLowerCase();

  const normalizedHeaders = data[0].map(h => normalize(h)); // ex. ["ID","NOM","EMAIL","WEBHOOK",...]
  const colNom = normalizedHeaders.indexOf("NOM");
  const colEmail = normalizedHeaders.indexOf("EMAIL");
  const colWebhook = normalizedHeaders.indexOf("WEBHOOK");
  if (colWebhook === -1 || (colNom === -1 && colEmail === -1)) {
    Logger.log("Colonnes webhook introuvables dans la feuille config.");
    return null;
  }

  const targetName = normalize(nom);
  const targetEmail = cleanEmail(email);

  // Recherche prioritaire par email : plus stable que le NOM.
  if (targetEmail) {
    for (let i = 1; i < data.length; i++) {
      const ligne = data[i];
      const emailConfig = colEmail !== -1
        ? cleanEmail(ligne[colEmail])
        : (colNom !== -1 && String(ligne[colNom] || "").includes("@")
            ? cleanEmail(ligne[colNom])
            : "");
      if (emailConfig && emailConfig === targetEmail) {
        const webhook = ligne[colWebhook];
        return webhook && webhook.toString().trim() !== ""
          ? webhook.toString().trim()
          : null;
      }
    }
  }

  // Fallback par nom si l'email n'est pas disponible ou ne correspond pas.
  for (let i = 1; i < data.length; i++) {
    const ligne    = data[i];
    const nomConfig = ligne[colNom];
    if (
      nomConfig &&
      normalize(nomConfig) === targetName
    ) {
      // Trouvé la bonne ligne → récupérer le webhook et le retourner s'il n'est pas vide
      const webhook = ligne[colWebhook];
      return webhook && webhook.toString().trim() !== ""
        ? webhook.toString().trim()
        : null;
    }
  }

  // 6️⃣ Aucun correspondance trouvée
  Logger.log(`Aucun webhook personnel trouvé pour "${nom || targetEmail}"`);
  return null;
}

// ==== Fonction pour formater la date ====
function formatDateFR(dateInput) {
  try {
    const d = new Date(dateInput);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
  } catch (e) {
    return dateInput;
  }
}
// ==== Fonction pour formater l'heure ====
function formatHeure(input) {
  try {
    const d = new Date(input);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm");
  } catch (e) {
    return input;
  }
}
// ==== Fonction pour chercher une valeur dans la feuille config (clé/valeur classique) ====
function chercherValeurDansConfig(cle) {
  if (!cle || typeof cle !== "string") {
    Logger.log("Clé invalide ou non définie : " + cle);
    return null;
  }
  const cleanKey = value => (value || "")
    .toString()
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .trim()
    .toLowerCase();

  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("config");
  if (!configSheet) {
    Logger.log("Feuille config introuvable !");
    return null;
  }
  // Les deux premières lignes sont "activer Notifications" et "Test Notifications"
  // On scanne lignes 1 et 2 (index 0 et 1)
  const data = configSheet.getDataRange().getValues();
  for (let i = 0; i < 2 && i < data.length; i++) {
    const key = data[i][0];
    if (key && cleanKey(key) === cleanKey(cle)) {
      const valeur = data[i][1] ? String(data[i][1]).trim() : "";
      Logger.log(`Clé trouvée: '${cle}', valeur: '${valeur}'`);
      return valeur;
    }
  }
  // Pour les webhooks spéciaux type "webhook Kiosksign", on peut aussi regarder tout le tableau
  for (let i = 2; i < data.length; i++) {
    const key = data[i][0];
    if (key && cleanKey(key) === cleanKey(cle)) {
      const valeur = data[i][1] ? String(data[i][1]).trim() : "";
      Logger.log(`Clé trouvée (webhook): '${cle}', valeur: '${valeur}'`);
      return valeur;
    }
  }
  Logger.log(`Clé non trouvée dans config : ${cle}`);
  return null;
}
// ==== Fonction pour envoyer un message via webhook ====
function envoyerMessageWebhook(webhookUrl, message) {
  if (!webhookUrl || webhookUrl.trim() === "") {
    Logger.log("Webhook vide. Message non envoyé.");
    return;
  }
  const payload = { text: message };
  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    Logger.log(`Message envoyé (${response.getResponseCode()}): ${response.getContentText()}`);
  } catch (error) {
    Logger.log(`Erreur d'envoi webhook : ${error.message}`);
  }
}
// ==== Fonction de test pour simuler une ligne de présence ====
function testerNotification() {
  const ligneTest = {
    nom: "DUPONT",
    prenom: "Claire",
    date: new Date(),
    arrivee: new Date("2025-06-24T08:30:00"),
    depart: new Date("2025-06-24T17:15:00"),
    ipIn: "192.168.1.10",
    ipOut: "192.168.1.99",
    loc: "Bureau Paris",
    nbHeures: 8.75,
    notif: "OUI" // mettre "" pour tester la connexion, "OUI" pour tester la déconnexion
  };
  const webhookPerso = chercherWebhookPersonnel(ligneTest.nom);
  const date = formatDateFR(ligneTest.date);
  const arrivee = formatHeure(ligneTest.arrivee);
  const depart = ligneTest.depart ? formatHeure(ligneTest.depart) : "";
  const heureNotif = formatHeure(new Date());
  let message = "";
  if (!ligneTest.notif) {
    message = `📌 📆 ✅ *${ligneTest.prenom}*, vous vous êtes *connecté(e)* le *${date}* à *${arrivee}* sur *${ligneTest.loc}* (${ligneTest.ipIn}).\nRespectez bien vos heures de planification. Merci — L'équipe IFFP — 🤝✨`;
  } else if (ligneTest.notif === "OUI" && ligneTest.depart) {
    message = `📌 📆 ✍🏼 *${ligneTest.prenom}* s'est *déconnecté(e)* le *${date}* à *${heureNotif}*. ` +
              `Connexion à *${arrivee}* (${ligneTest.ipIn})  //  Déconnexion à *${depart}* (${ligneTest.ipOut}), ` +
              `soit *${ligneTest.nbHeures}* heures de travail enregistrées.\n` +
              `Merci pour votre présence et votre implication. — L'équipe IFFP — 🙏✨`;
  }
  if (message && webhookPerso) {
    envoyerMessageWebhook(webhookPerso, message);
  } else {
    const webhookKiosksign = chercherValeurDansConfig("webhook Kiosksign");
    if (webhookKiosksign) {
      const messageErreur = `⚠️ Webhook de test manquant pour ${ligneTest.prenom} ${ligneTest.nom}. Merci de le configurer dans la feuille config colonne NOM/Webhook.`;
      envoyerMessageWebhook(webhookKiosksign, messageErreur);
    }
  }
}
