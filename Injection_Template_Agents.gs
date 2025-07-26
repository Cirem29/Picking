/**
 * Script temporaire pour injecter et masquer la feuille de format dans les fichiers individuels d'agents existants.
 * À exécuter UNE SEULE FOIS, puis à supprimer (ou commenter).
 */
function injectFormatTemplateIntoExistingAgentFiles() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listAgentsSheet = ss.getSheetByName("Liste Agents");
  var masterFormatTemplateSheet = ss.getSheetByName("Template_Agent_Format"); // Le nom de votre feuille de format dans le fichier centralisé

  if (!listAgentsSheet) {
    ui.alert("Erreur", "La feuille 'Liste Agents' est introuvable.", ui.ButtonSet.OK);
    return;
  }
  if (!masterFormatTemplateSheet) {
    ui.alert("Erreur", "La feuille 'Template_Agent_Format' est introuvable. Assurez-vous qu'elle existe et est nommée 'Template_Agent_Format'.", ui.ButtonSet.OK);
    return;
  }

  var driveFolderId = "1mVOeNt4MfUZVmq0gaFMSBH4RYRc_xAuJ"; // Votre ID de dossier Drive

  var agentData = listAgentsSheet.getRange(2, 1, listAgentsSheet.getLastRow() - 1, 4).getValues(); 

  var injectedCount = 0;
  var errors = [];
  var hiddenSheetNameInAgentFile = "_AgentSheetFormatTemplate"; // Nom de la feuille masquée dans le fichier de l'agent

  for (var i = 0; i < agentData.length; i++) {
        var row = agentData[i];
        var agentFirstNameForFileName = row[1]; 
        var agentFileId = row[2];      
        var agentFullNameFromList = row[3]; 

        if (!agentFileId || !agentFirstNameForFileName) {
          Logger.log("Skipping row in Liste Agents due to missing File ID or first name: %s", row);
          continue;
        }

        Logger.log("Traitement du fichier de l'agent: %s (ID: %s)", agentFullNameFromList, agentFileId);

        try {
          var agentFile = SpreadsheetApp.openById(agentFileId);
          var agentFormatSheet = agentFile.getSheetByName(hiddenSheetNameInAgentFile);

          if (!agentFormatSheet) { // Si la feuille n'existe pas, on la copie
            // --- DÉBUT DE LA CORRECTION ---
            var newlyCopiedSheet = masterFormatTemplateSheet.copyTo(agentFile); // CopyTo retourne l'objet Sheet directement
            newlyCopiedSheet.setName(hiddenSheetNameInAgentFile); // Renomme la feuille en utilisant l'objet direct
            newlyCopiedSheet.hideSheet(); // Masque la feuille
            // --- FIN DE LA CORRECTION ---
            
            injectedCount++;
            Logger.log("Feuille '%s' injectée et masquée dans le fichier de '%s'.", hiddenSheetNameInAgentFile, agentFullNameFromList);
          } else {
            Logger.log("Feuille '%s' déjà présente dans le fichier de '%s'. Ignoré.", hiddenSheetNameInAgentFile, agentFullNameFromList);
          }

        } catch (e) {
          errors.push("Erreur pour l'agent '" + agentFullNameFromList + "' (ID: " + agentFileId + "): " + e.message);
          Logger.log("Erreur inattendue lors de l'injection pour l'agent '%s': %s", agentFullNameFromList, e.message);
        }
      }

  var msg = injectedCount + " fichiers d'agents mis à jour avec le template de format.";
  if (errors.length > 0) {
    msg += "\n\nErreurs rencontrées :\n" + errors.join("\n");
  }
  ui.alert("Injection terminée", msg, ui.ButtonSet.OK);
}
