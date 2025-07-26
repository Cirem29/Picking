/**
 * Ce fichier contient les fonctions de maintenance des agents (ajout, suppression).
 *
 * NOTE IMPORTANTE POUR LA MAINTENANCE FUTURE :
 * - Les plages de lignes et colonnes sont définies de manière explicite (ex: B6:B19 pour Dashboard).
 * - Les plages du template Stats Agents sont définies (A1:D32 dans 'Stats agents Template').
 * Si l'équipe augmente ou si la structure des tableaux change, ces plages devront être ajustées manuellement dans le code.
 * - Le placeholder pour le nom de la feuille dans les formules INDIRECT est 'AGENT_SHEET_NAME'.
 */

// --- FONCTION POUR AFFICHER LA BOÎTE DE DIALOGUE D'AJOUT D'AGENT ---
function showAddAgentDialog() {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Dialogue_Ajout_Agent')
      .setWidth(450)
      .setHeight(350);
  ui.showModalDialog(htmlOutput, 'Ajouter un Agent');
}

// --- FONCTION POUR FERMER LES DIALOGUES HTML (utilisée par Dialogue_Ajout_Agent.html et potentiellement d'autres) ---
function closeDialog() {
  google.script.host.close();
}

// --- FONCTION CÔTÉ SERVEUR POUR TRAITER LE FORMULAIRE D'AJOUT D'AGENT ---
function processAddAgentForm(agentFullName, agentSheetName, agentFirstNameForFile, agentEmail) { 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listAgentsSheet = ss.getSheetByName("Liste Agents");
  var templateSheet = ss.getSheetByName("Template"); // Feuille modèle pour dupliquer les feuilles centrales d'agent
  var dashboardSheet = ss.getSheetByName("DASHBOARD"); // Accès à la feuille Dashboard
  var statsAgentsSheet = ss.getSheetByName("Stats agents"); // Accès à la feuille Stats agents
  var statsAgentsTemplateSheet = ss.getSheetByName("Stats agents Template"); // Accès au template pour Stats agents

  var driveFolderId = "1mVOeNt4MfUZVmq0gaFMSBH4RYRc_xAuJ"; // Votre ID de dossier Drive (CORRIGÉ !)
  var driveTemplateFileId = "1xGQaVaY7a2y1I9LZc2rYTNhZLOheZRJNk4yEHLG7Uds"; // ID du fichier template Drive

  // --- Vérifications préliminaires ---
  if (!listAgentsSheet) { return { success: false, message: "Erreur: La feuille 'Liste Agents' est introuvable." }; }
  if (!templateSheet) { return { success: false, message: "Erreur: La feuille modèle 'Template' est introuvable dans ce fichier." }; }
  if (!dashboardSheet) { return { success: false, message: "Erreur: La feuille 'DASHBOARD' est introuvable. Veuillez la nommer 'DASHBOARD'." }; }
  if (!statsAgentsSheet) { return { success: false, message: "Erreur: La feuille 'Stats agents' est introuvable." }; }
  if (!statsAgentsTemplateSheet) { return { success: false, message: "Erreur: La feuille 'Stats agents Template' est introuvable." }; }

  // Vérifier si la feuille de l'agent existe déjà dans le classeur central
  if (ss.getSheetByName(agentSheetName)) {
    return { success: false, message: "Erreur: Une feuille nommée '" + agentSheetName + "' existe déjà dans ce classeur." };
  }

  // Vérifier si l'agent existe déjà dans la 'Liste Agents' par son nom complet ou son e-mail
  var listAgentData = listAgentsSheet.getRange("D:E").getValues(); // Lire colonnes D (Prénom NOM) et E (Email)
  for (var i = 0; i < listAgentData.length; i++) {
    if (listAgentData[i][0] && listAgentData[i][0].toLowerCase() === agentFullName.toLowerCase()) {
      return { success: false, message: "Erreur: L'agent '" + agentFullName + "' existe déjà dans la feuille 'Liste Agents'." };
    }
    if (listAgentData[i][1] && listAgentData[i][1].toLowerCase() === agentEmail.toLowerCase()) {
        return { success: false, message: "Erreur: L'e-mail '" + agentEmail + "' est déjà associé à un agent dans 'Liste Agents'." };
    }
  }

  try {
    // --- 1. Dupliquer la feuille modèle dans le classeur central ---
    var newAgentSheet = templateSheet.copyTo(ss);
    newAgentSheet.setName(agentSheetName);
    Logger.log("Feuille centrale '%s' créée.", agentSheetName);

    // --- 2. Créer le fichier de suivi individuel sur Google Drive ---
    var templateFile = DriveApp.getFileById(driveTemplateFileId); // Fichier modèle Drive (votre "Template suivi")
    
    var newAgentDriveFile = templateFile.makeCopy("Suivi Qualité 2025 - " + agentFirstNameForFile, DriveApp.getFolderById(driveFolderId));
    var newAgentFileId = newAgentDriveFile.getId();
    Logger.log("Fichier Drive '%s' créé avec ID: %s", newAgentDriveFile.getName(), newAgentFileId);

    // Le template de formatage masqué doit déjà être dans "Template suivi", donc il est automatiquement copié avec le nouveau fichier.

    // --- 3. Mettre à jour la feuille 'Liste Agents' ---
    var lastRowListAgents = listAgentsSheet.getLastRow();
    listAgentsSheet.getRange(lastRowListAgents + 1, 1).setValue(agentSheetName); // Colonne A (Nom de la feuille centrale)
    listAgentsSheet.getRange(lastRowListAgents + 1, 2).setValue(agentFirstNameForFile); // Colonne B (Prénom pour nom du fichier)
    listAgentsSheet.getRange(lastRowListAgents + 1, 3).setValue(newAgentFileId); // Colonne C (ID du fichier Drive)
    listAgentsSheet.getRange(lastRowListAgents + 1, 4).setValue(agentFullName); // Colonne D (Prénom NOM complet)
    listAgentsSheet.getRange(lastRowListAgents + 1, 5).setValue(agentEmail); // Colonne E (Email)
    Logger.log("Ligne ajoutée à 'Liste Agents' pour '%s' avec e-mail '%s'.", agentFullName, agentEmail);

    // --- 4. Ajouter l'agent au Dashboard (colonne B, lignes 6-19) ---
    // PLAGE À VÉRIFIER POUR LA MAINTENANCE FUTURE : B6:B19 du DASHBOARD
    var dashboardAgentListRange = dashboardSheet.getRange("B6:B19");
    var dashboardAgentNames = dashboardAgentListRange.getValues();
    var firstEmptyRowInDashboard = -1;

    for (var r = 0; r < dashboardAgentNames.length; r++) {
        if (!dashboardAgentNames[r][0] || dashboardAgentNames[r][0].toString().trim() === "") { // Si la cellule est vide
            firstEmptyRowInDashboard = r; // Index de la ligne vide dans la plage (0-based)
            break;
        }
    }

    if (firstEmptyRowInDashboard !== -1) {
        var targetDashboardRow = 6 + firstEmptyRowInDashboard; // La ligne réelle sur le Dashboard
        dashboardSheet.getRange(targetDashboardRow, 2).setValue(agentFullName); // Écrit le nom complet en colonne B
        Logger.log("Agent '%s' ajouté au Dashboard à la ligne %s.", agentFullName, targetDashboardRow);
    } else {
        Logger.log("Avertissement: Toutes les lignes du Dashboard B6:B19 sont occupées. L'agent '%s' n'a pas été ajouté au Dashboard.", agentFullName);
    }

    // --- 5. Ajouter le bloc de l'agent à 'Stats agents' (selon 'Stats agents Template') ---
    // PLAGE DU TEMPLATE STATS AGENTS : A1:D32 sur la feuille 'Stats agents Template'
    var statsTemplateRange = statsAgentsTemplateSheet.getRange("A1:D32");
    var statsTemplateValues = statsTemplateRange.getValues(); 
    var statsTemplateFormulasRaw = statsTemplateRange.getFormulas(); 
    var statsTemplateBackgrounds = statsTemplateRange.getBackgrounds(); 
    var statsTemplateFontColors = statsTemplateRange.getFontColors();
    var statsTemplateFontWeights = statsTemplateRange.getFontWeights();
    var statsTemplateDataValidations = statsTemplateRange.getDataValidations(); 
    var statsTemplateTextStyles = statsTemplateRange.getTextStyles(); 
    var statsTemplateMergedRanges = statsTemplateRange.getMergedRanges(); 

    // Trouver la prochaine colonne disponible dans 'Stats agents'
    var lastColumnInStats = statsAgentsSheet.getLastColumn();
    var targetStatsStartCol = lastColumnInStats + 1;
    // Si la feuille est presque vide ou le dernier contenu est avant la colonne P (16), on commence à P.
    targetStatsStartCol = Math.max(targetStatsStartCol, 16); 

    var numRowsTemplate = statsTemplateRange.getNumRows(); // 32
    var numColsTemplate = statsTemplateRange.getNumColumns(); // 4

    var statsAgentTargetRange = statsAgentsSheet.getRange(
        1, targetStatsStartCol, 
        numRowsTemplate, 
        numColsTemplate
    );

    // 5.1 Effacer la zone cible avant de coller pour éviter les résidus (seulement le contenu/formatage dans la nouvelle zone)
    // Ne pas utiliser breakApartMergedCells() sur Range. La création des nouvelles fusions gérera les superpositions.
    statsAgentsSheet.getRange(1, targetStatsStartCol, statsAgentsSheet.getLastRow(), numColsTemplate).clearContent();
    statsAgentsSheet.getRange(1, targetStatsStartCol, statsAgentsSheet.getLastRow(), numColsTemplate).clearFormat();
    statsAgentsSheet.getRange(1, targetStatsStartCol, statsAgentsSheet.getLastRow(), numColsTemplate).clearDataValidations();
    
    // 5.2 Coller toutes les propriétés de format et les valeurs non-formules
    statsAgentTargetRange.setBackgrounds(statsTemplateBackgrounds);
    statsAgentTargetRange.setFontColors(statsTemplateFontColors);
    statsAgentTargetRange.setFontWeights(statsTemplateFontWeights);
    statsAgentTargetRange.setDataValidations(statsTemplateDataValidations);
    statsAgentTargetRange.setTextStyles(statsTemplateTextStyles); 
    statsAgentTargetRange.setValues(statsTemplateValues); // Coller les valeurs (texte des en-têtes comme "Pickings", "Total moyenne")
    Logger.log("Formats et valeurs de base copiés pour le bloc Stats agents de '%s'.", agentFullName);

    // 5.3 Recréer les fusions de cellules
    for (var i = 0; i < statsTemplateMergedRanges.length; i++) {
        var mergedRange = statsTemplateMergedRanges[i];
        var startRow = mergedRange.getRow(); 
        var startCol = mergedRange.getColumn(); 
        var numRows = mergedRange.getNumRows();
        var numCols = mergedRange.getNumColumns();
        
        var newMergeStartCol = targetStatsStartCol + (startCol - 1); // Ajuster la colonne de début de la fusion
        
        statsAgentsSheet.getRange(startRow, newMergeStartCol, numRows, numCols).merge();
    }
    Logger.log("Fusions de cellules recréées pour le bloc Stats agents de '%s'.", agentFullName);


    // 5.4 Mettre à jour la cellule A1 (du template) avec le nom complet de l'agent
    // Cette cellule A1 est maintenant une cellule fusionnée
    statsAgentsSheet.getRange(1, targetStatsStartCol).setValue(agentFullName);
    Logger.log("Nom d'agent '%s' placé dans l'en-tête du bloc Stats agents.", agentFullName);


    // 5.5 Écrire les formules modifiées une par une
    var agentSheetNamePlaceholder = "AGENT_SHEET_NAME"; // Le placeholder dans les formules de votre template

    // On parcourt la plage A3:C30 (lignes 3 à 30, colonnes A à C) du template
    for (var r = 3; r <= 30; r++) { // Lignes 3 à 30
      for (var c = 1; c <= 3; c++) { // Colonnes 1 (A), 2 (B), 3 (C) du template
        var templateFormula = statsTemplateFormulasRaw[r - 1][c - 1]; // Récupère la formule brute du tableau 0-based
        
        if (templateFormula && typeof templateFormula === 'string' && templateFormula.startsWith('=')) { 
          var modifiedFormula = templateFormula.replace(agentSheetNamePlaceholder, agentSheetName);
          statsAgentsSheet.getRange(r, targetStatsStartCol + (c - 1)).setFormula(modifiedFormula);
        }
      }
    }
    Logger.log("Formules INDIRECT et MOYENNE copiées et ajustées pour le bloc Stats agents de '%s'.", agentFullName);
    
    // Assurer que la colonne séparatrice est vide et non formatée (colonne D du template)
    var separatorColumn = targetStatsStartCol + 3; 
    statsAgentsSheet.getRange(1, separatorColumn, numRowsTemplate, 1).clearContent();
    statsAgentsSheet.getRange(1, separatorColumn, numRowsTemplate, 1).clearFormat();

    Logger.log("Bloc 'Stats agents' ajouté pour l'agent '%s' à partir de la colonne %s.", agentFullName, targetStatsStartCol);

    // --- 6. Partager le nouveau fichier avec l'agent ---
    newAgentDriveFile.addCommenter(agentEmail); 
    Logger.log("Fichier '%s' partagé avec '%s' comme commentateur.", newAgentDriveFile.getName(), agentEmail);

    return { success: true, message: "Agent '" + agentFullName + "' ajouté et fichier partagé avec succès !" };

  } catch (e) {
    Logger.log("Erreur lors de l'ajout de l'agent %s: %s", agentFullName, e.message);
    return { success: false, message: "Une erreur est survenue lors de l'ajout de l'agent: " + e.message };
  }
}

// --- FONCTION POUR AFFICHER LA BOÎTE DE DIALOGUE DE SUPPRESSION D'AGENT ---
function showRemoveAgentDialog() {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Dialogue_Retrait_Agent')
      .setWidth(450)
      .setHeight(250); 
  ui.showModalDialog(htmlOutput, 'Supprimer un Agent');
}

// --- FONCTION CÔTÉ SERVEUR POUR OBTENIR LA LISTE DES AGENTS POUR LA LISTE DÉROULANTE ---
function getAgentsListForRemoveDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listAgentsSheet = ss.getSheetByName("Liste Agents");
  var agents = [];

  if (!listAgentsSheet) {
    Logger.log("Feuille 'Liste Agents' introuvable.");
    return agents; 
  }

  // Lire la colonne D (Prénom NOM complet) à partir de la ligne 2
  var agentData = listAgentsSheet.getRange(2, 1, listAgentsSheet.getLastRow() - 1, 4).getValues(); 

  for (var i = 0; i < agentData.length; i++) {
    var agentSheetName = agentData[i][0]; // Colonne A (Nom de la feuille)
    var agentFullName = agentData[i][3]; // Colonne D (Prénom NOM complet)
    if (agentFullName) {
      agents.push({ fullName: agentFullName, sheetName: agentSheetName });
    }
  }
  return agents;
}

// --- FONCTION CÔTÉ SERVEUR POUR TRAITER LA SUPPRESSION DE L'AGENT ---
function processRemoveAgentForm(agentFullName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listAgentsSheet = ss.getSheetByName("Liste Agents");
  var dashboardSheet = ss.getSheetByName("DASHBOARD");
  var statsAgentsSheet = ss.getSheetByName("Stats agents"); // Accès à la feuille Stats agents

  // --- Vérifications préliminaires ---
  if (!listAgentsSheet) { return { success: false, message: "Erreur: La feuille 'Liste Agents' est introuvable." }; }
  if (!dashboardSheet) { return { success: false, message: "Erreur: La feuille 'DASHBOARD' est introuvable." }; }
  if (!statsAgentsSheet) { return { success: false, message: "Erreur: La feuille 'Stats agents' est introuvable." }; }

  var agentRowInListAgents = -1;
  var agentSheetName = ""; 
  
  // Chercher l'agent dans la feuille 'Liste Agents'
  var listAgentData = listAgentsSheet.getRange("D:D").getValues(); 
  for (var r = 0; r < listAgentData.length; r++) {
    if (listAgentData[r][0] && listAgentData[r][0].toLowerCase() === agentFullName.toLowerCase()) {
      agentRowInListAgents = r + 1; 
      agentSheetName = listAgentsSheet.getRange(r + 1, 1).getValue(); 
      break;
    }
  }

  if (agentRowInListAgents === -1) {
    return { success: false, message: "Erreur: L'agent '" + agentFullName + "' n'a pas été trouvé dans la feuille 'Liste Agents'." };
  }

  try {
    // --- 1. Masquer la feuille de l'agent dans le classeur central ---
    var agentCentralSheet = ss.getSheetByName(agentSheetName);
    if (agentCentralSheet) {
      agentCentralSheet.hideSheet();
      Logger.log("Feuille centrale '%s' masquée.", agentSheetName);
    } else {
      Logger.log("Avertissement: Feuille centrale '%s' non trouvée pour masquage. Peut-être déjà supprimée.", agentSheetName);
    }

    // --- 2. Supprimer la ligne de l'agent dans 'Liste Agents' ---
    listAgentsSheet.deleteRow(agentRowInListAgents);
    Logger.log("Ligne de l'agent '%s' supprimée de 'Liste Agents' (ligne %s).", agentFullName, agentRowInListAgents);

    // --- 3. Effacer le nom de l'agent du Dashboard (colonne B, lignes 6-19) ---
    // PLAGE À VÉRIFIER POUR LA MAINTENANCE FUTURE : B6:B19 du DASHBOARD
    var dashboardAgentListRange = dashboardSheet.getRange("B6:B19");
    var dashboardAgentNames = dashboardAgentListRange.getValues();
    
    for (var r = 0; r < dashboardAgentNames.length; r++) {
        if (dashboardAgentNames[r][0] && dashboardAgentNames[r][0].toString().trim().toLowerCase() === agentFullName.toLowerCase()) {
            var targetDashboardRow = 6 + r; 
            // Effacer le nom de l'agent et les données associées (C à E) de cette ligne
            dashboardSheet.getRange(targetDashboardRow, 2, 1, 4).clearContent(); // <<< CHANGEMENT : De colonne B à E (4 colonnes)
            // Optionnel: clearFormat() si besoin
            Logger.log("Agent '%s' effacé du Dashboard à la ligne %s.", agentFullName, targetDashboardRow);
            break; 
        }
    }

    // --- 4. NE PAS EFFACER LE BLOC DE COLONNES DE L'AGENT DANS 'Stats agents' ---
    Logger.log("Bloc de l'agent '%s' dans 'Stats agents' conservé (pas d'effacement).", agentFullName);

    return { success: true, message: "Agent '" + agentFullName + "' supprimé avec succès !" };

  } catch (e) {
    Logger.log("Erreur lors de la suppression de l'agent %s: %s", agentFullName, e.message);
    return { success: false, message: "Une erreur est survenue lors de la suppression de l'agent: " + e.message };
  }
}
