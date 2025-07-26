function gestionHebdomadaire() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rangeData = sheet.getDataRange();
  const lastRow = rangeData.getLastRow();

  // Étape 1 : Insérer une nouvelle ligne pour la nouvelle semaine
  const newRow = 23; // Par exemple, la ligne 23 est le début des données hebdomadaires
  sheet.insertRowBefore(newRow);

  // Copier les formats et titres de la ligne précédente
  const previousRow = sheet.getRange(newRow + 1, 1, 1, sheet.getLastColumn());
  previousRow.copyTo(sheet.getRange(newRow, 1, 1, sheet.getLastColumn()), { contentsOnly: false });

  // Étape 2 : Sauvegarder et restaurer la formule de C23 : =TEXTJOIN(", "; VRAI; FILTER($B$6:$B$20; $A$6:$A$20 = VRAI))
  const cellC26 = sheet.getRange("C23");
  const formulaC26 = cellC26.getFormula(); // Sauvegarde la formule de C23
  const cellB26 = sheet.getRange("B23");
  const formulaB26 = cellB26.getFormula(); // Sauvegarde la formule de B23
  const donneesvalidation = sheet.getRange("A6:A19"); // Cellule où "imprimer" le texte brut (ex. E1)
  
  // Effacer uniquement le contenu de la ligne 23 sauf les formules et la cellule B24
  const rangeToClear = sheet.getRange(23, 1, 1, sheet.getLastColumn());
  const values = rangeToClear.getValues()[0];
  for (let i = 0; i < values.length; i++) {
    if (i !== 1) { // Exclure la colonne B (index 1)
      rangeToClear.getCell(1, i + 1).clearContent();
    }
  }
  
  // Restaurer les formules dans C23
  cellC26.setFormula(formulaC26); // Rétablit la formule dans C26

  // Étape 3 : Imprimer le texte brut dans une autre cellule (par exemple, E1)
  imprimerTexteBrut(); // Appelle la fonction définie ci-dessous

  // Étape 4 : Optionnel, recalculer les données (ex. pour la moyenne ou autres)
  SpreadsheetApp.flush();
  donneesvalidation.clearContent();
}


function imprimerTexteBrut() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Feuille active
  const celluleSource = sheet.getRange("B23:F23"); // Cellule avec le TEXTJOIN 
  const celluleDestination = sheet.getRange("B24:F24"); // Cellule où "imprimer" le texte brut 
  
    // Sauvegarde explicite des valeurs actuelles
  const texteBrut = celluleSource.getDisplayValues(); // Utilisez getDisplayValues pour le texte brut
  celluleDestination.setValues(texteBrut); // Écrit les valeurs brutes dans la cellule destination
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const editedCell = e.range;

  // Vérifier si la modification est dans la cellule C23
  if (editedCell.getA1Notation() == 'C23') {
    const value = editedCell.getValue();
    
    // Déplacer le contenu actuel de C23 vers C24
    const currentC24Value = sheet.getRange('C24').getValue();
    sheet.getRange('C24').setValue(value);
    
    // Optionnel : Déplacer les valeurs suivantes vers la ligne suivante (C25, C26, etc.)
    let nextRow = 25;
    while (sheet.getRange('C' + nextRow).getValue() !== "") {
      const nextValue = sheet.getRange('C' + nextRow).getValue();
      sheet.getRange('C' + (nextRow + 1)).setValue(nextValue);
      sheet.getRange('C' + nextRow).setValue("");
      nextRow++;
    }
  }
}
