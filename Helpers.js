/**
 * ⏰ Crée un trigger automatique pour exécuter completerInfosPresences chaque soir à 23h00
 */
function creerTriggerAutoCompleterInfos() {
  // Supprime les anciens triggers pour éviter les doublons
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let t of allTriggers) {
    if (t.getHandlerFunction() === "completerInfosPresences") {
      ScriptApp.deleteTrigger(t);
    }
  }

  // Crée un nouveau trigger journalier à 23h
  ScriptApp.newTrigger("completerInfosPresences")
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();

  Logger.log("✅ Trigger 'completerInfosPresences' créé avec succès (exécution quotidienne à 23h00).");
}

function convertToBase100(hoursDecimal) {
  if (isNaN(hoursDecimal)) return 0;
  const heures = Math.floor(hoursDecimal);
  const minutes = (hoursDecimal - heures) * 60;
  const base100 = heures + (minutes / 60) * 100 / 100;
  return parseFloat(base100.toFixed(2));
}
