---
layout: slide
title: "Exporter chaque onglet en PDF"
---
Découvrez ce script Google Apps Script pour exporter chaque onglet d'un classeur en PDF avec un nom de fichier personnalisé :

```javascript
/*******************************************************
 *  EXPORTER CHAQUE ONGLETS EN PDF
 *  Format de fichier : [B9] - [C1] - 16 oct.pdf
 *  Menu : Exporter → En PDF (par onglet)
 *******************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Exporter")
    .addItem("En PDF (par onglet)", "exportSheetsToPDFs")
    .addToUi();
}

function exportSheetsToPDFs() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const folderName = spreadsheet.getName() + " - PDFs";

  // ✅ Mode silencieux : si pas d'UI disponible (exécuté depuis Apps Script)
  let overwrite = true;
  try {
    const ui = SpreadsheetApp.getUi();
    const overwriteResponse = ui.alert(
      "Export PDF",
      "Voulez-vous écraser les fichiers PDF existants ?",
      ui.ButtonSet.YES_NO
    );
    overwrite = overwriteResponse == ui.Button.YES;
  } catch (e) {
    Logger.log("⚠️ Aucune interface UI disponible — mode silencieux activé (écrasement auto).");
  }

  // 📁 Vérifier si le dossier existe ou le créer
  let folder;
  const folders = DriveApp.getFoldersByName(folderName);
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  // 🗓️ Date du jour : ex. "16 oct"
  const now = new Date();
  const day = now.getDate();
  const monthNames = ["jan", "fev", "mar", "avr", "mai", "jun", "jul", "aou", "sep", "oct", "nov", "dec"];
  const month = monthNames[now.getMonth()];
  const dateSuffix = `${day} ${month}`;

  // 🔁 Parcourir les onglets
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const c1Value = (sheet.getRange("C1").getValue() || "").toString().trim();
    const b9Value = (sheet.getRange("B9").getValue() || "").toString().trim();

    // 🚫 Ignorer les onglets non numériques ou C1 = "cie"
    if (!/^[0-9]+$/.test(sheetName) || c1Value.toLowerCase() === "cie") {
      Logger.log(`Ignoré : ${sheetName}`);
      return;
    }

    // 📄 Nom du PDF : [B9] - [C1] - 16 oct.pdf
    const pdfName = `${b9Value || "SansTitre"} - ${c1Value || sheetName} - ${dateSuffix}`;

    // 🔍 Vérifier si le fichier existe déjà
    const existingFiles = folder.getFilesByName(pdfName + ".pdf");
    if (existingFiles.hasNext()) {
      if (overwrite) {
        while (existingFiles.hasNext()) {
          existingFiles.next().setTrashed(true);
        }
        Logger.log(`♻️ Fichier existant écrasé : ${pdfName}.pdf`);
      } else {
        Logger.log(`⏭️ Fichier existant ignoré : ${pdfName}.pdf`);
        return;
      }
    }

    // 🧾 URL d’export PDF
    const url =
      "https://docs.google.com/spreadsheets/d/" +
      spreadsheet.getId() +
      "/export?format=pdf&" +
      "size=A4&" +
      "portrait=false&" + // paysage
      "fitw=true&" +
      "sheetnames=false&" +
      "printtitle=false&" +
      "pagenumbers=false&" +
      "gridlines=true&" +
      "fzr=false&" +
      "gid=" +
      sheet.getSheetId();

    const token = ScriptApp.getOAuthToken();
    let attempt = 0;
    const maxAttempts = 5;
    let success = false;

    while (attempt < maxAttempts && !success) {
      attempt++;
      Logger.log(`🕒 Essai ${attempt} pour exporter ${sheetName}...`);
      try {
        const response = UrlFetchApp.fetch(url, {
          headers: { Authorization: "Bearer " + token },
          muteHttpExceptions: true,
        });

        const code = response.getResponseCode();
        if (code === 200) {
          folder.createFile(response.getBlob().setName(pdfName + ".pdf"));
          Logger.log(`✅ Export réussi : ${sheetName} → ${pdfName}.pdf`);
          success = true;
        } else if (code === 429) {
          Logger.log(`⚠️ Erreur 429 (trop de requêtes) → pause 2s`);
          Utilities.sleep(2000);
        } else {
          Logger.log(`❌ Erreur HTTP ${code} : ${response.getContentText()}`);
          success = true; // abandon pour autres erreurs
        }
      } catch (e) {
        Logger.log(`❌ Exception : ${e}`);
        Utilities.sleep(2000);
      }
    }

    if (!success) {
      Logger.log(`⚠️ Échec après ${maxAttempts} essais pour ${sheetName}.`);
    }

    // Petite pause entre les onglets
    Utilities.sleep(2000);
  });

  try {
    SpreadsheetApp.getUi().alert("✅ Export terminé.\nPDFs créés dans le dossier : " + folder.getName());
  } catch (e) {
    Logger.log("✅ Export terminé (mode silencieux, pas d'alerte UI affichée).");
  }
}
```

Appuyez sur **S** pour afficher cette slide, puis sur la flèche droite pour revenir à la présentation.
