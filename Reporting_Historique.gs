/**
 * Ce script crée un reporting hebdomadaire des statistiques des agents.
 * Il récupère les données de la feuille 'DASHBOARD' et les ajoute
 * à la feuille 'Historique V2' avec la date du jour.
 */
function creerReportingHebdomadaire() {
  // --- Configuration ---
  const nomFeuilleSource = 'DASHBOARD';
  const nomFeuilleCible = 'Historique V2';
  
  // --- Accès au classeur et aux feuilles ---
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName(nomFeuilleSource);
  const destinationSheet = spreadsheet.getSheetByName(nomFeuilleCible);

  // Vérification si les feuilles existent
  if (!sourceSheet) {
    Logger.log(`Erreur : La feuille source "${nomFeuilleSource}" est introuvable.`);
    return;
  }
  if (!destinationSheet) {
    Logger.log(`Erreur : La feuille cible "${nomFeuilleCible}" est introuvable.`);
    return;
  }

  // --- 1. Récupération des données de la feuille source ---
  const agents = sourceSheet.getRange('B6:B19').getValues();
  const picking = sourceSheet.getRange('D6:D19').getValues();
  const doubleEcoute = sourceSheet.getRange('E6:E19').getValues();
  const noteQualite = sourceSheet.getRange('F6:F19').getValues();
  const wqNette = sourceSheet.getRange('O6:O19').getValues();
  
  const dateActuelle = new Date();
  
  // --- 2. Préparation des données pour l'écriture ---
  const donneesAecrire = [];
  
  for (let i = 0; i < agents.length; i++) {
    if (agents[i][0] !== "") {
      const nouvelleLigne = [
        dateActuelle,
        agents[i][0],
        picking[i][0],
        doubleEcoute[i][0],
        noteQualite[i][0],
        wqNette[i][0]
      ];
      donneesAecrire.push(nouvelleLigne);
    }
  }

  if (donneesAecrire.length === 0) {
    Logger.log("Avertissement : Aucun agent trouvé dans la plage B6:B19. Le script n'a rien à exporter.");
    return;
  }

  // --- 3. Écriture des données dans la feuille cible ---
  const premiereLigneVide = destinationSheet.getLastRow() + 1;
  
  destinationSheet.getRange(
    premiereLigneVide,
    1,
    donneesAecrire.length,
    donneesAecrire[0].length
  ).setValues(donneesAecrire);
  
  Logger.log(`Reporting créé avec succès. ${donneesAecrire.length} lignes ajoutées.`);
}
