// --- FONCTION PRINCIPALE POUR ENREGISTRER UNE ÉVALUATION ---
function enregistrerEvaluation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pickingSheet = ss.getSheetByName("Picking");
  if (!pickingSheet) {
    Browser.msgBox("Erreur", "La feuille 'Picking' est introuvable. Veuillez vous assurer qu'elle existe et est nommée 'Picking'.", Browser.Buttons.OK);
    return;
  }

  var ui = SpreadsheetApp.getUi();

  // --- 1. Demander le nom de l'agent ---
  var listAgentsSheet = ss.getSheetByName("Liste Agents");
  if (!listAgentsSheet) {
    Browser.msgBox("Erreur", "La feuille 'Liste Agents' est introuvable. Nécessaire pour la validation de l'agent.", Browser.Buttons.OK);
    return;
  }
  var agentData = listAgentsSheet.getRange(2, 1, listAgentsSheet.getLastRow() - 1, 4).getValues(); 

  var sheetNamesForPrompt = agentData.map(row => row[0]); 
  var agentPrompt = ui.prompt(
    "Enregistrer Évaluation",
    "Pour quel agent (nom de la feuille, ex: Test) souhaitez-vous enregistrer cette évaluation ?\n\nFeuilles existantes : " + sheetNamesForPrompt.join(", "),
    ui.ButtonSet.OK_CANCEL
  );

  if (agentPrompt.getSelectedButton() === ui.Button.CANCEL || !agentPrompt.getResponseText()) {
    return; 
  }
  var agentSheetNameInput = agentPrompt.getResponseText().trim(); 

  var agentFullNameForDashboard = ""; 
  var foundAgentDataRow = null; 

  for (var i = 0; i < agentData.length; i++) {
    if (agentData[i][0] && agentData[i][0].toLowerCase() === agentSheetNameInput.toLowerCase()) {
      agentFullNameForDashboard = agentData[i][3]; 
      foundAgentDataRow = agentData[i]; 
      break;
    }
  }

  if (!agentFullNameForDashboard) {
    Browser.msgBox("Erreur", "L'agent '" + agentSheetNameInput + "' n'est pas listé ou n'a pas de nom complet dans la feuille 'Liste Agents'.", Browser.Buttons.OK);
    return;
  }

  var targetSheet = ss.getSheetByName(agentSheetNameInput); 
  if (!targetSheet) {
    Browser.msgBox("Erreur", "La feuille pour l'agent '" + agentSheetNameInput + "' n'existe pas. Veuillez la créer d'abord ou vérifier l'orthographe.", Browser.Buttons.OK);
    return;
  }

  // --- 2. Demander le numéro de la semaine ---
  var weekPrompt = ui.prompt(
    "Numéro de la semaine",
    "Quel est le numéro de la semaine pour cette évaluation ?",
    ui.ButtonSet.OK_CANCEL
  );

  if (weekPrompt.getSelectedButton() === ui.Button.CANCEL || !weekPrompt.getResponseText()) {
    return; 
  }
  var weekNumberInput = weekPrompt.getResponseText().trim();

  if (isNaN(weekNumberInput) || weekNumberInput.length === 0) {
      Browser.msgBox("Erreur", "Le numéro de la semaine doit être un nombre valide.", Browser.Buttons.OK);
      return;
  }
  var weekNumberInt = parseInt(weekNumberInput); 

  // --- Écrire UNIQUEMENT le numéro de la semaine sur le template avant la copie ---
  pickingSheet.getRange("H1").setValue(weekNumberInt);


  // --- Définition de la plage du template à copier (A1:N18) ---
  var templateToCopyRange = pickingSheet.getRange("A1:N18"); 
  var sourceNumRows = templateToCopyRange.getNumRows(); 
  var sourceNumCols = templateToCopyRange.getNumColumns(); 

  // --- Trouver la première ligne vide dans la feuille de l'agent cible pour le TEMPLATE ---
  var lastRowTargetForTemplate = targetSheet.getLastRow();
  // CHANGEMENT : Coller directement à la suite du dernier contenu (+1)
  var startRowForTemplatePaste = lastRowTargetForTemplate + 1; 

  // --- DÉFINIR LA COLONNE DE DÉPART POUR LE COLLAGE DU TEMPLATE : Colonne E (5ème colonne) ---
  var startColumnForTemplatePaste = 5; 

  // --- Définir la plage de destination du template (cellule supérieure gauche) ---
  var pasteTopLeftCellForTemplate = targetSheet.getRange(startRowForTemplatePaste, startColumnForTemplatePaste); 
  
  // --- VÉRIFIER ET AJOUTER DES COLONNES SI NÉCESSAIRE ---
  var requiredEndColumnForMainTemplate = startColumnForTemplatePaste + sourceNumCols - 1; 
  var absoluteRequiredMaxColumn = Math.max(requiredEndColumnForMainTemplate, 26); 

  var maxColumnsInTargetSheetBeforeInsert = targetSheet.getMaxColumns();

  if (absoluteRequiredMaxColumn > maxColumnsInTargetSheetBeforeInsert) {
    var columnsToAdd = absoluteRequiredMaxColumn - maxColumnsInTargetSheetBeforeInsert;
    targetSheet.insertColumns(maxColumnsInTargetSheetBeforeInsert + 1, columnsToAdd);
  }

  // --- VÉRIFIER ET AJOUTER DES LIGNES SI NÉCESSAIRE ---
  var requiredEndRowForPaste = startRowForTemplatePaste + sourceNumRows - 1; 
  var maxRowsInTargetSheet = targetSheet.getMaxRows();

  if (requiredEndRowForPaste > maxRowsInTargetSheet) {
    var rowsToAdd = requiredEndRowForPaste - maxRowsInTargetSheet;
    if (maxRowsInTargetSheet === 0) { 
        targetSheet.insertRows(1, rowsToAdd); 
    } else {
        targetSheet.insertRowsAfter(maxRowsInTargetSheet, rowsToAdd); 
    }
  }


  templateToCopyRange.copyTo(pasteTopLeftCellForTemplate, {contentsOnly: false, formatOnly: false});


  // --- REMPLIR LA PARTIE RÉSUMÉ SUR LA FEUILLE DE L'AGENT (dans le fichier centralisé) ---

  // Calcul des coordonnées réelles des données du template collé sur la feuille de l'agent
  var templateOffsetRow = startRowForTemplatePaste - 1; 
  var templateOffsetCol = startColumnForTemplatePaste - 1; 

  // Récupérer les valeurs des cellules du template qui vient d'être copié
  var semaineValue = pickingSheet.getRange("H1").getValue(); 
  var noteGlobaleValue = targetSheet.getRange(templateOffsetRow + 8, templateOffsetCol + 14).getValue(); 
  var noteDoubleEcouteValue = targetSheet.getRange(templateOffsetRow + 12, templateOffsetCol + 14).getValue(); 

  var axesProgresValue = targetSheet.getRange(templateOffsetRow + 15, templateOffsetCol + 1).getValue(); 
  var pointsPositifsValue = targetSheet.getRange(templateOffsetRow + 17, templateOffsetCol + 1).getValue(); 
  var feedbackAgentsValue = targetSheet.getRange(templateOffsetRow + 15, templateOffsetCol + 14).getValue(); 
  var planActionValue = targetSheet.getRange(templateOffsetRow + 17, templateOffsetCol + 14).getValue(); 


  // Trouver la PREMIÈRE LIGNE VIDE dans la zone de résumé (colonnes A:D)
  var lastRowInResumeColumn = targetSheet.getRange("A:A").getValues().filter(String).length;
  var resumeDataStartRow = 3; 
  var nextResumeRow = Math.max(resumeDataStartRow, lastRowInResumeColumn + 1); 


  // Remplir les colonnes A:D à la suite
  targetSheet.getRange(nextResumeRow, 1).setValue(semaineValue); 
  targetSheet.getRange(nextResumeRow, 2).setValue(noteGlobaleValue); 
  targetSheet.getRange(nextResumeRow, 3).setValue(noteDoubleEcouteValue); 
  
  var commentaireSynthese = 
    "Axe(s) de progrès / indispensable sur le prochain suivi qualité : \n" + (axesProgresValue || "N/A") + "\n\n" +
    "Point(s) positif(s) : \n" + (pointsPositifsValue || "N/A") + "\n\n" +
    "Feedback Agent : \n" + (feedbackAgentsValue || "N/A") + "\n\n" +
    "Plan d'action : \n" + (planActionValue || "N/A");

  targetSheet.getRange(nextResumeRow, 4).setValue(commentaireSynthese); 
  targetSheet.getRange(nextResumeRow, 4).setWrap(true); 
  targetSheet.setRowHeight(nextResumeRow, 100); 


  // --- MISE À JOUR DU DASHBOARD (LIGNES 6-20) ---
  var dashboardSheet = ss.getSheetByName("DASHBOARD"); 
  if (!dashboardSheet) {
    return { success: false, message: "Erreur Dashboard: La feuille 'DASHBOARD' est introuvable." };
  } else {
    var year = 2025; 
    var dateOfWeek = getMondayOfWeek(year, weekNumberInt); 
    
    var dashboardAgentListRange = dashboardSheet.getRange("B6:B20"); 
    var dashboardAgentNames = dashboardAgentListRange.getValues();
    var dashboardAgentRow = -1; 

    for (var r = 0; r < dashboardAgentNames.length; r++) {
      if (dashboardAgentNames[r][0] && agentFullNameForDashboard && dashboardAgentNames[r][0].toString().trim().toLowerCase() === agentFullNameForDashboard.toLowerCase()) {
        dashboardAgentRow = r; 
        break;
      }
    }

    if (dashboardAgentRow === -1) {
      return { success: false, message: "Avertissement Dashboard: L'agent '" + agentFullNameForDashboard + "' n'a pas été trouvé dans le Dashboard principal." };
    } else {
      var targetDashboardRow = 6 + dashboardAgentRow;

      dashboardSheet.getRange(targetDashboardRow, 3).setValue(weekNumberInt); 
      dashboardSheet.getRange(targetDashboardRow, 4).setValue(noteGlobaleValue); 
      dashboardSheet.getRange(targetDashboardRow, 5).setValue(noteDoubleEcouteValue); 
    }
  }

  // --- Réinitialiser la feuille "Picking" (effacer uniquement les zones de saisie) ---
  pickingSheet.getRange("A3:N7").clearContent(); 
  pickingSheet.getRange("H1").clearContent();   
  pickingSheet.getRange("N12").clearContent(); 
  
  pickingSheet.getRange("A13").clearContent();  
  pickingSheet.getRange("A15").clearContent();  
  pickingSheet.getRange("A17").clearContent();  
  pickingSheet.getRange("N15").clearContent();  
  pickingSheet.getRange("N17").clearContent();  
  
  // Renvoyer un objet de succès pour le HTML
  return { success: true, message: "Évaluation enregistrée pour '" + agentSheetNameInput + "' (Semaine " + weekNumberInt + ")." };

}

// --- UTILITAIRE : Fonction pour calculer le lundi d'une semaine donnée pour une année donnée ---
function getMondayOfWeek(year, week) {
  var date = new Date(year, 0, 1); 
  var dayOfWeek = date.getDay(); 
  var daysToMonday = 0;

  if (dayOfWeek === 0) { 
    daysToMonday = 1;
  } else if (dayOfWeek > 1) { 
    daysToMonday = 8 - dayOfWeek;
  }

  date.setDate(date.getDate() + daysToMonday); 
  date.setDate(date.getDate() + (week - 1) * 7); 
  
  return date; 
}

// --- FONCTION DE MISE À JOUR DES FICHIERS INDIVIDUELS ---
function mettreAJourFichiersAgents() {
  var mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var listAgentsSheet = mainSpreadsheet.getSheetByName("Liste Agents");
  var hiddenTemplateFormatSheetName = "_AgentSheetFormatTemplate"; // Nom de la feuille masquée dans le fichier de l'agent

  if (!listAgentsSheet) {
    Logger.log("Erreur: La feuille 'Liste Agents' est introuvable pour la mise à jour des fichiers agents.");
    return;
  }

  var driveFolderId = "1mVOeNt4MfUZVmq0gaFMSBH4RYRc_xAuJ"; 
  var agentDataRange = listAgentsSheet.getRange(2, 1, listAgentsSheet.getLastRow() - 1, 4); 
  var agentData = agentDataRange.getValues(); 

  var updatedAgentsCount = 0;
  var errors = [];

  for (var i = 0; i < agentData.length; i++) { 
    var row = agentData[i];
    var centralSheetName = row[0]; 
    var agentFirstNameForFileName = row[1]; 
    var agentFileId = row[2];      
    var agentFullNameFromList = row[3]; 

    if (!centralSheetName || !agentFirstNameForFileName || !agentFullNameFromList) { 
      Logger.log("Skipping row in Liste Agents due to missing essential data: %s", row); 
      continue; 
    }

    Logger.log("--- Traitement de l'agent: %s (Central Sheet: %s, File ID: %s) ---", agentFullNameFromList, centralSheetName, agentFileId); 

    try {
      var centralAgentSheet = mainSpreadsheet.getSheetByName(centralSheetName);
      if (!centralAgentSheet) {
        errors.push("Erreur: La feuille '" + centralSheetName + "' n'existe pas dans le fichier centralisé. (Agent: " + agentFullNameFromList + ")");
        Logger.log("Erreur: Feuille centrale '%s' manquante pour l'agent %s", centralSheetName, agentFullNameFromList); 
        continue;
      }

      var lastRowCentral = centralAgentSheet.getLastRow();
      Logger.log("lastRowCentral pour '%s': %s", centralSheetName, lastRowCentral); 
      
      if (lastRowCentral < 1) { 
        errors.push("Avertissement: Aucune donnée d'évaluation trouvée pour l'agent '" + agentFullNameFromList + "' dans le fichier centralisé. (Feuille " + centralSheetName + ")");
        Logger.log("Avertissement: Pas de données source (lastRowCentral < 1) pour l'agent '%s'", agentFullNameFromList); 
      }

      var dataToCopySourceRange = centralAgentSheet.getRange("A1:Z" + lastRowCentral); 
      Logger.log("Plage source à copier: %s!%s", centralAgentSheet.getName(), dataToCopySourceRange.getA1Notation()); 

      var valuesToCopy = dataToCopySourceRange.getValues();
      Logger.log("Nombre de lignes de données à copier (valuesToCopy.length): %s", valuesToCopy.length); 
      
      if (valuesToCopy.length === 0 || (valuesToCopy.length > 0 && valuesToCopy[0].length === 0)) {
        errors.push("Avertissement: La plage source de l'agent '" + agentFullNameFromList + "' est vide ou n'a pas de colonnes. Rien à copier.");
        Logger.log("Avertissement: 'valuesToCopy' est vide ou sans colonnes pour l'agent '%s'. Pas de données à coller.", agentFullNameFromList); 
      } else {
        Logger.log("Premières valeurs à copier (extrait 5x5) pour '%s': %s", agentFullNameFromList, JSON.stringify(valuesToCopy.slice(0, Math.min(valuesToCopy.length, 5)).map(function(row) { return row.slice(0, Math.min(row.length, 5)); }))); 
        var formulasToCopyForLog = dataToCopySourceRange.getFormulas();
        Logger.log("Premières formules à copier (extrait 5x5) pour '%s': %s", agentFullNameFromList, JSON.stringify(formulasToCopyForLog.slice(0, Math.min(formulasToCopyForLog.length, 5)).map(function(row) { return row.slice(0, Math.min(row.length, 5)); }))); 
      }

      // --- DÉBUT DU NOUVEAU FLUX DE MISE EN FORME PAR TEMPLATE ---
      var agentFile; // Déclarer agentFile ici pour qu'il soit accessible
      
      if (agentFileId) {
        try {
          agentFile = SpreadsheetApp.openById(agentFileId);
        } catch (e) {
          errors.push("Erreur: Fichier agent (ID: " + agentFileId + ") introuvable pour l'agent '" + agentFullNameFromList + "'. Vérifiez l'ID ou les permissions. Erreur: " + e.message);
          Logger.log("Erreur: Impossible d'ouvrir le fichier par ID '%s' pour '%s': %s", agentFileId, agentFullNameFromList, e.message);
          continue; // Passe à l'agent suivant si l'ouverture par ID échoue
        }
      } else {
        // Fallback à la recherche par nom si l'ID est vide/manquant
        var files = DriveApp.getFolderById(driveFolderId).getFilesByName("Suivi Qualité 2025 - " + agentFirstNameForFileName); 
        if (files.hasNext()) {
          agentFile = SpreadsheetApp.open(files.next());
        } else {
          errors.push("Erreur: Fichier 'Suivi Qualité 2025 - " + agentFirstNameForFileName + "' introuvable pour l'agent '" + agentFullNameFromList + "'. Vérifiez le nom ou l'ID.");
          Logger.log("Erreur: Fichier agent introuvable par nom pour '%s'", agentFullNameFromList); 
          continue; // Passe à l'agent suivant si le fichier n'est pas trouvé par nom
        }
      }

      // À ce point, agentFile DOIT être défini, sinon un 'continue' a déjà été exécuté.
      // Ajout d'une vérification de sécurité supplémentaire (bien que normalement inutile avec les continues ci-dessus)
      if (!agentFile) {
        errors.push("Erreur interne: La référence au fichier agent est nulle après les tentatives d'ouverture pour '" + agentFullNameFromList + "'.");
        Logger.log("Erreur interne: agentFile est null après les tentatives d'ouverture pour '%s'.", agentFullNameFromList);
        continue;
      }

      var agentTargetSheet = agentFile.getSheetByName("Suivi Qualité");
      if (!agentTargetSheet) {
        errors.push("Erreur: Feuille 'Suivi Qualité' introuvable dans le fichier de l'agent '" + agentFullNameFromList + "'.");
        Logger.log("Erreur: Feuille cible 'Suivi Qualité' manquante dans le fichier de l'agent '%s'", agentFullNameFromList); 
        continue;
      }
      Logger.log("Fichier et feuille cible trouvés pour l'agent '%s'.", agentFullNameFromList); 

      var hiddenFormatTemplateSheet = agentFile.getSheetByName(hiddenTemplateFormatSheetName);
      if (!hiddenFormatTemplateSheet) {
          errors.push("Erreur: La feuille de formatage masquée '" + hiddenTemplateFormatSheetName + "' est introuvable dans le fichier de l'agent '" + agentFullNameFromList + "'. La mise en forme sera basique.");
          Logger.log("Erreur: Feuille de formatage masquée manquante pour '%s'.", agentFullNameFromList);
          // Si le template de format est manquant, on ne peut pas copier les formats. On peut choisir de :
          // 1. Sortir et forcer la correction (comme avant)
          // 2. Continuer mais avec un formatage basique (moins bon visuellement)
          // Pour l'instant, on sort pour que l'erreur soit corrigée (template manquant).
          return; 
      }
      Logger.log("Feuille de formatage masquée trouvée pour l'agent '%s'.", agentFullNameFromList);

      // 1. Nettoyer complètement la feuille visible "Suivi Qualité" (formats, contenu, MFC, tout)
      agentTargetSheet.clear(); 
      Logger.log("Feuille 'Suivi Qualité' de l'agent '%s' complètement nettoyée.", agentFullNameFromList);

      // 2. Copier le contenu et les formats (y compris MFC) de la feuille masquée vers la feuille visible.
      var formatSourceRange = hiddenFormatTemplateSheet.getRange("A1:Z" + hiddenFormatTemplateSheet.getLastRow()); 
      formatSourceRange.copyTo(agentTargetSheet.getRange("A1"), {contentsOnly: false, formatOnly: false});
      Logger.log("Format et structure copiés depuis '%s' vers 'Suivi Qualité' pour l'agent '%s'.", hiddenTemplateFormatSheetName, agentFullNameFromList);

      // 3. Déposer les données réelles par-dessus le formatage (ce sont les valuesToCopy lues du centralAgentSheet)
      if (valuesToCopy.length > 0 && valuesToCopy[0].length > 0) { 
        var numRows = valuesToCopy.length;
        var numCols = valuesToCopy[0].length; 
        var destinationRange = agentTargetSheet.getRange(1, 1, numRows, numCols);
        
        destinationRange.setValues(valuesToCopy);
        Logger.log("Données réelles copiées pour l'agent '%s'.", agentFullNameFromList);
      } else {
        Logger.log("Avertissement: Aucune donnée réelle à coller pour l'agent '%s'. Le fichier cible a été formaté mais reste vide.", agentFullNameFromList);
      }
      // --- FIN DU NOUVEAU FLUX DE MISE EN FORME PAR TEMPLATE ---

      updatedAgentsCount++;

    } catch (e) {
      errors.push("Erreur pour l'agent '" + agentFullNameFromList + "': " + e.message);
      Logger.log("Erreur inattendue pour l'agent '%s': %s", agentFullNameFromList, e.message); 
    }
  }

  var msg = updatedAgentsCount + " fichiers d'agents mis à jour avec succès.";
  if (errors.length > 0) {
    msg += "\n\nErreurs rencontrées :\n" + errors.join("\n");
  }
  Logger.log("Mise à jour terminée: " + msg); 
}
