function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Scripts")
    .addItem("Dashboard", "gestionHebdomadaire")
    .addItem("Impression historique", "insererLigneEtExecuterScript")
    .addItem("Format date", "testTraitementDatesSeul")
    .addItem("Reporting Stats agents", "creerReportingHebdomadaire")
    .addSeparator() 
    .addItem('Enregistrer Évaluation', 'enregistrerEvaluation')
    .addItem('Mettre à jour fichiers Agents', 'mettreAJourFichiersAgents')
    .addSeparator() 
    .addItem('Ajouter un nouvel Agent', 'showAddAgentDialog')
    .addItem('Supprimer un Agent', 'showRemoveAgentDialog')
    .addSeparator() 
    .addItem('Mettre à jour Dashboard WQ Validées', 'mettreAJourDashboardWqValidees')
    .addToUi();
}
