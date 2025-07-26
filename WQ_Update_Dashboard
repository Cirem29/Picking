// --- NOUVELLE FONCTION : METTRE À JOUR LE TABLEAU WQ VALIDÉES SUR LE DASHBOARD ---
function mettreAJourDashboardWqValidees() {
  // var ui = SpreadsheetApp.getUi(); // Non nécessaire car plus de ui.alert
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var dashboardSheet = ss.getSheetByName("DASHBOARD");
  var listAgentsSheet = ss.getSheetByName("Liste Agents");

  if (!dashboardSheet) {
    // ui.alert("Erreur", "La feuille 'DASHBOARD' est introuvable. Veuillez la nommer 'DASHBOARD'.", ui.ButtonSet.OK);
    Logger.log("Erreur: La feuille 'DASHBOARD' est introuvable pour la mise à jour WQ Validées.");
    return;
  }
  if (!listAgentsSheet) {
    // ui.alert("Erreur", "La feuille 'Liste Agents' est introuvable. Veuillez vous assurer qu'elle existe et est nommée 'Liste Agents'.", ui.ButtonSet.OK);
    Logger.log("Erreur: La feuille 'Liste Agents' est introuvable pour la mise à jour WQ Validées.");
    return;
  }

  // --- 1. Définir la période de 2 mois en cours (fixe : Jan-Fév, Mar-Avr, etc.) ---
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth(); // 0 pour Janvier, 1 pour Février...

  var periodIndex = Math.floor(currentMonth / 2); // 0 pour Jan/Fev, 1 pour Mar/Avr...
  var startMonthOfPeriod = periodIndex * 2; // Mois de début (0-indexed) de la période
  var endMonthOfPeriod = startMonthOfPeriod + 1; // Mois de fin (0-indexed) de la période

  var startDate = new Date(currentYear, startMonthOfPeriod, 1); // Premier jour de la période
  var endDate = new Date(currentYear, endMonthOfPeriod + 1, 0); // Dernier jour de la période (le 0ème jour du mois suivant)

  // Pour s'assurer que l'heure est minuit (début de journée) pour la comparaison
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(23, 59, 59, 999); 

  Logger.log("Période de 2 mois en cours pour WQ Validées: du %s au %s", startDate.toLocaleDateString('fr-FR'), endDate.toLocaleDateString('fr-FR'));

  // --- 2. Lire la liste des agents depuis le Dashboard (colonne N6:N19) ---
  var dashboardAgentsRange = dashboardSheet.getRange("N6:N19"); // Colonne N pour le nom de l'agent
  var dashboardAgentNames = dashboardAgentsRange.getValues(); 

  // --- 3. Parcourir chaque agent pour calculer ses WQ validées ---
  var agentsToUpdate = []; 
  var errors = []; 

  for (var r = 0; r < dashboardAgentNames.length; r++) {
    var agentFullNameDashboard = dashboardAgentNames[r][0]; 

    if (!agentFullNameDashboard || agentFullNameDashboard.toString().trim() === "") {
        continue; 
    }
    
    var agentSheetName = "";
    var listAgentData = listAgentsSheet.getRange(2, 1, listAgentsSheet.getLastRow() - 1, 4).getValues(); 
    for (var j = 0; j < listAgentData.length; j++) {
        if (listAgentData[j][3] && listAgentData[j][3].toLowerCase() === agentFullNameDashboard.toLowerCase()) {
            agentSheetName = listAgentData[j][0]; 
            break;
        }
    }

    if (!agentSheetName) {
        Logger.log("Avertissement: Feuille d'agent centrale introuvable pour '%s' (calcul WQ validées). Skip WQ count.", agentFullNameDashboard);
        continue;
    }

    Logger.log("Calcul des WQ validées pour l'agent: %s (Feuille: %s)", agentFullNameDashboard, agentSheetName);

    try { 
      var agentCentralSheet = ss.getSheetByName(agentSheetName);
      if (!agentCentralSheet) {
        errors.push("Erreur: Feuille d'agent '" + agentSheetName + "' introuvable pour le calcul WQ.");
        continue;
      }

      // --- NOUVEAU : Définir la plage de données WQ de manière plus précise ---
      // On cherche la dernière ligne non vide dans la colonne S (première colonne du bloc WQ)
      var lastWqRowInAgentSheet = agentCentralSheet.getRange("S:S").getValues().filter(String).length; 
      var wqDataStartRow = 3; // Les données WQ commencent à la ligne 3 (après les en-têtes)
      
      // Si lastWqRowInAgentSheet est inférieur à wqDataStartRow (ex: 0, 1 ou 2), cela signifie qu'il n'y a pas de données WQ
      if (lastWqRowInAgentSheet < wqDataStartRow) {
          Logger.log("Avertissement: Aucune donnée WQ réelle pour l'agent '%s' (colonne S vide ou seulement en-têtes). Skip WQ counting.", agentFullNameDashboard);
          agentsToUpdate.push({ rowInDashboard: r + 6, count: 0 }); // Compter 0 pour cet agent
          continue; // Passer à l'agent suivant
      }

      var agentWqDataRange = agentCentralSheet.getRange("S" + wqDataStartRow + ":Z" + lastWqRowInAgentSheet); 
      var agentWqData = agentWqDataRange.getValues();
      Logger.log("Lecture de la plage WQ: '%s!%s' pour l'agent '%s'. Nombre de lignes lues: %s", 
                 agentSheetName, agentWqDataRange.getA1Notation(), agentFullNameDashboard, agentWqData.length);
      
      var valideWqCount = 0;
      var agentWqColIdx = {
          ID: 0, 
          CreatedTime: 1, 
          ValidationSupervision: 5 
      };

      for (var k = 0; k < agentWqData.length; k++) {
        var wqRow = agentWqData[k];

        // --- Vérification : Ignorer les lignes sans ID de WQ (lignes vides dans le tableau lu) ---
        if (!wqRow[agentWqColIdx.ID] || wqRow[agentWqColIdx.ID].toString().trim() === "") {
            continue; 
        }

        var rawWqDateValue = wqRow[agentWqColIdx.CreatedTime]; 
        var validationStatus = wqRow[agentWqColIdx.ValidationSupervision]; 
        var wqDate = null; 

        // Tenter de convertir la date de manière robuste
        if (rawWqDateValue instanceof Date && !isNaN(rawWqDateValue.getTime())) {
            wqDate = rawWqDateValue;
        } else if (typeof rawWqDateValue === 'number' && !isNaN(rawWqDateValue)) {
            wqDate = convertSheetsSerialToJsDate(rawWqDateValue); 
        } else if (typeof rawWqDateValue === 'string' && rawWqDateValue.trim() !== '') {
            try {
                var parsedDate = new Date(rawWqDateValue);
                if (!isNaN(parsedDate.getTime())) { 
                    wqDate = parsedDate;
                } else {
                    Logger.log("Avertissement: Échec d'analyse de la date (chaîne non valide) pour '%s', ligne %s (ID WQ: %s): '%s'", agentFullNameDashboard, k + 3, wqRow[agentWqColIdx.ID], rawWqDateValue);
                }
            } catch (e) {
                Logger.log("Avertissement: Erreur lors de l'analyse de la date (chaîne) pour '%s', ligne %s (ID WQ: %s): '%s' - %s", agentFullNameDashboard, k + 3, wqRow[agentWqColIdx.ID], rawWqDateValue, e.message);
            }
        }
        
        if (wqDate instanceof Date && !isNaN(wqDate.getTime())) {
            var isWqValidated = (validationStatus === true); 

            // --- NOUVEAUX LOGS DE DÉBOGAGE DANS LA CONDITION PRINCIPALE ---
            Logger.log("DEBUG WQ Count: Agent '%s', WQ ID '%s' (Ligne Feuille: %s)", agentFullNameDashboard, wqRow[agentWqColIdx.ID], k + 3);
            Logger.log("    Status: '%s' (doit être TRUE). Date WQ: '%s' (Heure: %s)", 
                       validationStatus, wqDate.toLocaleDateString('fr-FR'), wqDate.toTimeString());
            Logger.log("    Période: du '%s' au '%s'", startDate.toLocaleDateString('fr-FR'), endDate.toLocaleDateString('fr-FR'));
            Logger.log("    Conditions: Status Valide = %s, Date >= Start = %s, Date <= End = %s",
                       isWqValidated,
                       (wqDate >= startDate),
                       (wqDate <= endDate));
            // --- FIN NOUVEAUX LOGS ---

            if (isWqValidated && 
                wqDate >= startDate && wqDate <= endDate) {
                valideWqCount++;
                Logger.log("    -> WQ ID '%s' VALIDÉE et COMPTÉE.", wqRow[agentWqColIdx.ID]);
            } else {
                Logger.log("    -> WQ ID '%s' NON COMPTÉE. (Raison: Statut ou Date hors période).", wqRow[agentWqColIdx.ID]);
            }
        } else {
            Logger.log("Avertissement: Date WQ invalide (format non reconnu) pour '%s', ligne %s (ID WQ: %s): '%s'", agentFullNameDashboard, k + 3, wqRow[agentWqColIdx.ID], rawWqDateValue);
        }
      } 
      
      agentsToUpdate.push({ rowInDashboard: r + 6, count: valideWqCount }); 
      Logger.log("Agent '%s' : %s WQ validées trouvées pour la période actuelle.", agentFullNameDashboard, valideWqCount);

    } catch (e) { 
      errors.push("Erreur lors du calcul des WQ validées pour l'agent '" + agentFullNameDashboard + "': " + e.message);
      Logger.log("Erreur inattendue pour l'agent '%s' lors du calcul WQ validées: %s", agentFullNameDashboard, e.message);
    } 
  } 

  var updatedCount = 0;
  for (var i = 0; i < agentsToUpdate.length; i++) {
    var rowData = agentsToUpdate[i];
    dashboardSheet.getRange(rowData.rowInDashboard, 15).setValue(rowData.count); 
    updatedCount++;
  }
  
  var msg = updatedCount + " agents mis à jour pour les WQ validées sur le Dashboard.";
  if (errors.length > 0) {
    msg += "\n\nErreurs rencontrées :\n" + errors.join("\n");
  }
  Logger.log("Mise à jour WQ Validées terminée: " + msg); 
}

// --- UTILITAIRE : Fonction pour convertir un numéro de série Google Sheets en Date JavaScript ---
function convertSheetsSerialToJsDate(serial) {
  var MS_PER_DAY = 24 * 60 * 60 * 1000; 
  var excelEpoch = new Date(1899, 11, 30).getTime(); 
  var date = new Date(excelEpoch + serial * MS_PER_DAY);
  
  if (serial >= 60) { 
      date.setTime(date.getTime() - MS_PER_DAY); 
  }
  
  return date;
}
