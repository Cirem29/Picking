// --- CONFIGURATION GLOBALE ---
// Ces variables stockent les noms des feuilles et les paramètres BigQuery.

// ID de ton projet Google Cloud où se trouve BigQuery
var BIGQUERY_PROJECT_ID = 'helpdesk-data-xa3k'; 

// Noms des feuilles de calcul
var WQ_TICKETS_SHEET_NAME = "WQ"; // Feuille cible pour les tickets WQ extraits.

////////////////////////////////////////////////////////////////////////////////////////////////////
// --- FONCTION PRINCIPALE : MISE À JOUR DES DONNÉES WQ DEPUIS BIGQUERY ---
////////////////////////////////////////////////////////////////////////////////////////////////////

/**
 * Exécute une requête SQL BigQuery pour récupérer les tickets WQ et les insère dans la feuille "WQ".
 */
function updateWQDataFromBigQuery() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wqTicketsSheet = ss.getSheetByName(WQ_TICKETS_SHEET_NAME);

  // Crée la feuille "WQ" si elle n'existe pas.
  if (!wqTicketsSheet) {
    wqTicketsSheet = ss.insertSheet(WQ_TICKETS_SHEET_NAME);
    Logger.log("Feuille '" + WQ_TICKETS_SHEET_NAME + "' créée."); 
  }

  // Nettoie tout le contenu existant de la feuille "WQ".
  wqTicketsSheet.clear();
  Logger.log("Feuille '" + WQ_TICKETS_SHEET_NAME + "' nettoyée.");

  // Définit les en-têtes de colonnes pour la feuille "WQ".
  // L'ordre DOIT correspondre à l'ordre des colonnes dans ta requête SQL BigQuery.
  var headers = [
    "ID",
    "Ems Creation Time",
    "Log by Agent",
    "Impacted Service",
    "Impacted Category",
    "Impacted Subcategory",
    "WQ Code",          
    "WQ détail"         
  ];

  // Écrit les en-têtes dans la première ligne de la feuille "WQ".
  wqTicketsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  wqTicketsSheet.setFrozenRows(1); // Gèle la première ligne (les en-têtes).
  Logger.log("En-têtes ajoutés à la feuille '" + WQ_TICKETS_SHEET_NAME + "'.");

  var sqlQuery = `
    SELECT
      Id,
      EmsCreationTime,
      LogAgent_c_Name,
      RegisteredForActualService_DisplayLabel,
      Category_DisplayLabel,
      Subcategory_c_DisplayLabel,
      WrongQualificationCode_c,
      WrongQualification_c
    FROM
      \`helpdesk-data-xa3k.SMAX_OPENTEXT.REQUEST\`
    WHERE
      LogGroup_c_Name = 'CC-BRANDS'
      AND WrongQualificationBool_c = 'true'
      AND EXTRACT(YEAR FROM EmsCreationTime) = 2025;
  `;

  var queryRequest = {
    query: sqlQuery,
    useLegacySql: false 
  };

  var queryResults;
  var jobId;
  var jobLocation;

  try {
    queryResults = BigQuery.Jobs.query(queryRequest, BIGQUERY_PROJECT_ID);
    jobId = queryResults.jobReference.jobId;
    jobLocation = queryResults.jobReference.location; 
    Logger.log("Requête BigQuery exécutée. Job ID: " + jobId + ", Location: " + jobLocation);

  } catch (e) {
    Logger.log("Erreur critique lors du Lancement de la requête BigQuery : " + e.message);
    return; 
  }

  // Boucle pour attendre que la tâche soit complète et récupérer les résultats
  var sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs); 
    try {
      queryResults = BigQuery.Jobs.getQueryResults(BIGQUERY_PROJECT_ID, jobId, {location: jobLocation}); 
      Logger.log("Attente de la complétion de la tâche BigQuery... Job ID: " + jobId + ", Location: " + jobLocation);
    } catch (e) {
      Logger.log("Erreur lors de la récupération du statut/résultats du job BigQuery (" + jobId + ") à la localisation " + jobLocation + ": " + e.message);
      return; 
    }
  }
  
  Logger.log("Tâche BigQuery terminée.");

  // Récupère les lignes de résultats.
  var rows = queryResults.rows;
  var allData = [];

  if (rows) {
    Logger.log("Nombre de lignes récupérées : " + rows.length);
    // Parcours chaque ligne de résultats.
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var rowData = [];
      // Parcours chaque champ (colonne) dans la ligne.
      for (var j = 0; j < row.f.length; j++) {
        var field = row.f[j];
        
        // Formate EmsCreationTime comme un objet Date JavaScript
        if (headers[j] === "Ems Creation Time") { 
            var secondsSinceEpoch = parseFloat(field.v); 
            var millisecondsSinceEpoch = secondsSinceEpoch * 1000; 
            var date = new Date(millisecondsSinceEpoch); 
            
            if (date instanceof Date && !isNaN(date.getTime())) {
                 rowData.push(date); 
            } else {
                 rowData.push(null); 
            }
        } 
        else {
          rowData.push(field.v); 
        }
      }
      allData.push(rowData);
    }
  } else {
    Logger.log("Aucune ligne de résultats trouvée.");
  }

  // Écrit toutes les données collectées dans la feuille "WQ" en une seule opération.
  if (allData.length > 0) {
    wqTicketsSheet.getRange(2, 1, allData.length, allData[0].length).setValues(allData);
    Logger.log("Données insérées dans la feuille '" + WQ_TICKETS_SHEET_NAME + "'.");
  } else {
    Logger.log("Aucune donnée à insérer dans la feuille '" + WQ_TICKETS_SHEET_NAME + "'.");
  }

  // Ajuste automatiquement la largeur des colonnes pour qu'elles s'adaptent au contenu.
  wqTicketsSheet.autoResizeColumns(1, headers.length);
  Logger.log("Feuille '" + WQ_TICKETS_SHEET_NAME + "' mise à jour avec les données BigQuery.");
}
