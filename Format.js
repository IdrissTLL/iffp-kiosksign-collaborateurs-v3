function regrouperPresencesJour() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Présences");
  const lastRow = feuille.getLastRow();
  const fuseau = Session.getScriptTimeZone();

  Logger.log("➡️ Début du regroupement pour la date d'hier. Total lignes : " + lastRow);

  if (lastRow < 3) {
    Logger.log("❌ Pas assez de lignes pour travailler.");
    return;
  }

  // Obtenir la date d'hier au format JJ/MM/YYYY
  const hier = new Date();
  hier.setDate(hier.getDate() - 1);
  const dateHierTexte = Utilities.formatDate(hier, fuseau, "dd/MM/yyyy");

  Logger.log("🔍 Recherche des lignes correspondant à : " + dateHierTexte);

  const données = feuille.getRange(2, 1, lastRow - 1).getValues(); // Colonne A, depuis ligne 2
  const lignesCiblées = [];

  for (let i = 0; i < données.length; i++) {
    const ligneIndex = i + 2;
    let cellule = données[i][0];

    if (cellule instanceof Date) {
      const texte = Utilities.formatDate(cellule, fuseau, "dd/MM/yyyy");
      if (texte === dateHierTexte) {
        lignesCiblées.push(ligneIndex);
        Logger.log(`  ✅ Ligne ${ligneIndex} correspond à ${dateHierTexte}`);
      }
    } else if (typeof cellule === 'string') {
      const clean = cellule.trim();
      if (clean === dateHierTexte) {
        lignesCiblées.push(ligneIndex);
        Logger.log(`  ✅ Ligne ${ligneIndex} correspond à ${dateHierTexte} (texte brut)`);
      }
    }
  }

  if (lignesCiblées.length <= 1) {
    Logger.log("⛔️ Moins de 2 lignes trouvées pour regrouper. Fin du script.");
    return;
  }

  const debut = lignesCiblées[1]; // on laisse la première ligne visible
  const fin = lignesCiblées[lignesCiblées.length - 1];
  const nbLignes = fin - debut + 1;

  // 🔍 Vérifier uniquement si la première ligne à regrouper est déjà dans un groupe
  const ligneTest = debut;
  const profondeur = feuille.getRowGroupDepth(ligneTest);

  if (profondeur > 0) {
    Logger.log(`⏩ La ligne ${ligneTest} est déjà regroupée. Aucune action.`);
  } else {
    Logger.log(`📦 Regroupement des lignes ${debut} à ${fin} en cours...`);
    try {
      feuille.getRange(debut, 1, nbLignes).shiftRowGroupDepth(1);
      Logger.log("✅ Groupe créé avec succès.");
    } catch (e) {
      Logger.log(`❌ Erreur lors du regroupement : ${e.message}`);
    }
  }

  // ✅ Appliquer une bordure supérieure épaisse sur la première ligne visible (colonnes A à Q)
  const ligneVisible = lignesCiblées[0];
  try {
    feuille.getRange(ligneVisible, 1, 1, 17).setBorder(
      true, false, false, false, false, false,
      null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
    Logger.log(`🎨 Bordure supérieure épaisse appliquée sur la ligne ${ligneVisible}, colonnes A à Q`);
  } catch (e) {
    Logger.log(`❌ Erreur lors de l'application de la bordure : ${e.message}`);
  }

  feuille.collapseAllRowGroups();
  Logger.log("📉 Tous les groupes repliés. ✅ Script terminé.");
}

function regrouperToutesLesPresencesParJour() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Présences");
  const lastRow = feuille.getLastRow();
  const fuseau = Session.getScriptTimeZone();

  Logger.log("➡️ Démarrage du regroupement de tous les jours. Total lignes : " + lastRow);

  if (lastRow < 3) {
    Logger.log("❌ Trop peu de lignes.");
    return;
  }

  const données = feuille.getRange(2, 1, lastRow - 1).getValues(); // Colonne A uniquement
  const datesMap = new Map();

  // 🔍 Étape 1 : Collecter toutes les dates présentes (format JJ/MM/YYYY)
  for (let i = 0; i < données.length; i++) {
    const ligneIndex = i + 2;
    const cellule = données[i][0];

    let dateFormatee = null;

    if (cellule instanceof Date) {
      dateFormatee = Utilities.formatDate(cellule, fuseau, "dd/MM/yyyy");
    } else if (typeof cellule === 'string') {
      const clean = cellule.trim();
      const dateParts = clean.split('/');
      if (dateParts.length === 3) {
        dateFormatee = clean;
      }
    }

    if (dateFormatee) {
      if (!datesMap.has(dateFormatee)) {
        datesMap.set(dateFormatee, []);
      }
      datesMap.get(dateFormatee).push(ligneIndex);
    }
  }

  Logger.log("📆 Dates trouvées : " + [...datesMap.keys()].join(", "));

  // 🔁 Étape 2 : Boucle sur chaque date et appliquer le regroupement
  for (const [dateTexte, lignesCiblées] of datesMap.entries()) {
    Logger.log(`\n🔄 Traitement du ${dateTexte} (${lignesCiblées.length} lignes)`);

    if (lignesCiblées.length <= 1) {
      Logger.log(`⛔️ Moins de 2 lignes pour ${dateTexte}, on passe.`);
      continue;
    }

    const debut = lignesCiblées[1]; // on laisse la première ligne visible
    const fin = lignesCiblées[lignesCiblées.length - 1];
    const nbLignes = fin - debut + 1;

    // 🔍 Vérification si déjà groupé
    const profondeur = feuille.getRowGroupDepth(debut);
    if (profondeur > 0) {
      Logger.log(`⏩ Lignes ${debut} à ${fin} déjà regroupées. Pas de duplication.`);
    } else {
      try {
        feuille.getRange(debut, 1, nbLignes).shiftRowGroupDepth(1);
        Logger.log(`✅ Groupe créé : lignes ${debut} à ${fin}`);
      } catch (e) {
        Logger.log(`❌ Erreur lors du regroupement pour ${dateTexte} : ${e.message}`);
      }
    }

    // 🎨 Appliquer la bordure épaisse sur la ligne visible
    const ligneVisible = lignesCiblées[0];
    try {
      feuille.getRange(ligneVisible, 1, 1, 17).setBorder(
        true, false, false, false, false, false,
        null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
      Logger.log(`🎨 Bordure supérieure appliquée ligne ${ligneVisible}, A à Q`);
    } catch (e) {
      Logger.log(`❌ Erreur lors de la bordure sur ${dateTexte} : ${e.message}`);
    }
  }

  // 📉 Repli global
  feuille.collapseAllRowGroups();
  Logger.log("✅ Tous les groupes repliés. Script terminé.");
}

function regrouperSemaineDerniere() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Présences");
  const lastRow = feuille.getLastRow();
  const fuseau = Session.getScriptTimeZone();

  Logger.log("➡️ Démarrage du regroupement de la semaine dernière");

  // 📅 Trouver les bornes de la semaine dernière (lundi → dimanche)
  const today = new Date();
  const jourSemaine = today.getDay(); // 0=dim, 1=lun, ..., 6=sam
  const decalage = jourSemaine === 0 ? 7 : jourSemaine;

  const lundi = new Date(today);
  lundi.setDate(today.getDate() - decalage - 6);
  const dimanche = new Date(lundi);
  dimanche.setDate(lundi.getDate() + 6);

  Logger.log(`📆 Plage ciblée : du ${Utilities.formatDate(lundi, fuseau, "dd/MM/yyyy")} au ${Utilities.formatDate(dimanche, fuseau, "dd/MM/yyyy")}`);

  // 🧾 Lire toutes les dates de la colonne A
  const valeurs = feuille.getRange(2, 1, lastRow - 1).getValues();
  const lignesSemaine = [];

  for (let i = 0; i < valeurs.length; i++) {
    const ligne = i + 2;
    const valeur = valeurs[i][0];

    if (valeur instanceof Date) {
      const date = new Date(valeur.getFullYear(), valeur.getMonth(), valeur.getDate());

      if (date >= lundi && date <= dimanche) {
        const profondeur = feuille.getRowGroupDepth(ligne);

        if (profondeur >= 1) { // ✅ Ligne déjà groupée par jour
          lignesSemaine.push(ligne);
        }
      }
    }
  }

  if (lignesSemaine.length === 0) {
    Logger.log("⛔️ Aucune ligne trouvée déjà groupée pour cette semaine.");
    return;
  }

  const ligneDebut = Math.min(...lignesSemaine);
  const ligneFin = Math.max(...lignesSemaine);
  const nbLignes = ligneFin - ligneDebut + 1;

  Logger.log(`📚 Lignes à regrouper pour la semaine : ${ligneDebut} à ${ligneFin}`);

  // 🔍 Vérifier si déjà groupé à un niveau supérieur
  const profondeurDebut = feuille.getRowGroupDepth(ligneDebut);
  const profondeurFin = feuille.getRowGroupDepth(ligneFin);

  if (profondeurDebut >= 2 && profondeurFin >= 2) {
    Logger.log("⏩ Groupe de semaine déjà existant. Fin.");
    return;
  }

  try {
    feuille.getRange(ligneDebut, 1, nbLignes).shiftRowGroupDepth(1); // Ajout du groupement de niveau supérieur
    Logger.log("✅ Groupe de semaine créé avec succès.");
    feuille.collapseAllRowGroups();
    Logger.log("📉 Tous les groupes repliés.");
  } catch (e) {
    Logger.log(`❌ Erreur lors du regroupement de la semaine : ${e.message}`);
  }
}

function regrouperPresencesMoisDernier() {
  const feuille = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Présences");
  const lastRow = feuille.getLastRow();
  const fuseau = Session.getScriptTimeZone();

  Logger.log("➡️ Démarrage du regroupement du mois précédent. Total lignes : " + lastRow);

  if (lastRow < 3) {
    Logger.log("❌ Pas assez de lignes à traiter.");
    return;
  }

  // 📅 Déterminer le mois précédent
  const today = new Date();
  const premierDuMois = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const dernierDuMois = new Date(today.getFullYear(), today.getMonth(), 0); // dernier jour mois précédent

  Logger.log(`📆 Mois précédent : du ${Utilities.formatDate(premierDuMois, fuseau, "dd/MM/yyyy")} au ${Utilities.formatDate(dernierDuMois, fuseau, "dd/MM/yyyy")}`);

  // 🔍 Lecture de toutes les dates en colonne A
  const valeurs = feuille.getRange(2, 1, lastRow - 1).getValues();
  const lignesCiblées = [];

  for (let i = 0; i < valeurs.length; i++) {
    const ligne = i + 2;
    const cellule = valeurs[i][0];

    if (cellule instanceof Date) {
      const date = new Date(cellule.getFullYear(), cellule.getMonth(), cellule.getDate());

      if (date >= premierDuMois && date <= dernierDuMois) {
        const profondeur = feuille.getRowGroupDepth(ligne);
        if (profondeur >= 2) { // ✅ Groupée par jour + semaine
          lignesCiblées.push(ligne);
        }
      }
    }
  }

  if (lignesCiblées.length === 0) {
    Logger.log("⛔️ Aucune ligne groupée jour+semaine dans le mois précédent.");
    return;
  }

  const debut = Math.min(...lignesCiblées);
  const fin = Math.max(...lignesCiblées);
  const nbLignes = fin - debut + 1;

  Logger.log(`📚 Lignes du mois à regrouper : ${debut} à ${fin}`);

  // ⚠️ Vérifier si déjà groupées au niveau 3
  const profondeurDebut = feuille.getRowGroupDepth(debut);
  const profondeurFin = feuille.getRowGroupDepth(fin);

  if (profondeurDebut >= 3 && profondeurFin >= 3) {
    Logger.log("⏩ Groupe du mois déjà existant. Fin du script.");
    return;
  }

  try {
    feuille.getRange(debut, 1, nbLignes).shiftRowGroupDepth(1); // ✅ Regroupement de niveau 3
    Logger.log(`✅ Groupe du mois précédent créé de la ligne ${debut} à ${fin}`);
    feuille.collapseAllRowGroups();
    Logger.log("📉 Tous les groupes repliés. Terminé.");
  } catch (e) {
    Logger.log(`❌ Erreur lors du regroupement du mois : ${e.message}`);
  }
}
