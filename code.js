function setupTrigger() {
    const triggers = ScriptApp.getProjectTriggers();
    const triggerExists = triggers.some(trigger => trigger.getHandlerFunction() === "onOpen");

    if (!triggerExists) {
        ScriptApp.newTrigger("onOpen")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onOpen()
            .create();
        Logger.log("onOpen Trigger wurde erfolgreich erstellt!");
    }
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();

    // Submenü für globale Funktionen (Rechnungen importieren)
    const rechnungenMenu = ui.createMenu("🧾 Rechnungen");
    rechnungenMenu.addItem("📥 Rechnungen importieren", "importDriveFiles");

    // Submenü für Funktionen, die nur für "Einnahmen" und "Ausgaben" gelten
    const einnahmenAusgabenMenu = ui.createMenu("💰 Einnahmen/Ausgaben");
    einnahmenAusgabenMenu.addItem("🔀 Tabelle sortieren", "sortCurrentSheetByColumn");
    einnahmenAusgabenMenu.addItem("🔄 Aktualisieren (Formeln & Formatierung)", "updateSheetOnCurrentSheet");

    // Submenü für GUV
    const guvMenu = ui.createMenu("📊 GUV");
    guvMenu.addItem("✅ GUV berechnen", "calculateGUV");

    // Hauptmenü "Buchhaltung" mit den drei Submenüs
    ui.createMenu("📂 Buchhaltung")
        .addSubMenu(rechnungenMenu)
        .addSubMenu(einnahmenAusgabenMenu)
        .addSubMenu(guvMenu)
        .addToUi();
}

/* ----- Globale Funktionen für den Rechnungsimport ----- */
function importDriveFiles() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEinnahmen = ss.getSheetByName("Rechnungen Einnahmen");
    const sheetAusgaben = ss.getSheetByName("Rechnungen Ausgaben");
    const einnahmenTab = ss.getSheetByName("Einnahmen");
    const ausgabenTab = ss.getSheetByName("Ausgaben");
    let historyTab = ss.getSheetByName("Änderungshistorie");

    // Änderungshistorie erstellen, falls nicht vorhanden
    if (!historyTab) {
        historyTab = ss.insertSheet("Änderungshistorie");
        historyTab.appendRow(["Datum", "Rechnungstyp", "Dateiname", "Link zur Datei"]);
    }

    // Überschriften in den Rechnungs-Sheets setzen, falls leer
    if (sheetEinnahmen.getLastRow() === 0) {
        sheetEinnahmen.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
    }
    if (sheetAusgaben.getLastRow() === 0) {
        sheetAusgaben.appendRow(["Dateiname", "Link zur Datei", "Rechnungsnummer"]);
    }

    // Hauptordner ermitteln (Ordner, in dem sich das aktuelle Sheet befindet)
    const fileId = ss.getId();
    const file = DriveApp.getFileById(fileId);
    const parentIterator = file.getParents();
    if (!parentIterator.hasNext()) {
        Logger.log("Kein übergeordneter Ordner gefunden.");
        return;
    }
    const parentFolder = parentIterator.next();

    // Unterordner "Einnahmen" und "Ausgaben" suchen
    const einnahmenOrdner = getFolderByName(parentFolder, "Einnahmen");
    const ausgabenOrdner = getFolderByName(parentFolder, "Ausgaben");

    if (einnahmenOrdner) {
        // Import in Rechnungen-Eingangsblatt und Hauptübersicht "Einnahmen"
        importFilesFromFolder(einnahmenOrdner, sheetEinnahmen, einnahmenTab, "Einnahme", historyTab);
        // Spaltenbreite in Rechnungen-Eingang anpassen
        sheetEinnahmen.autoResizeColumns(1, sheetEinnahmen.getLastColumn());
        if (einnahmenTab.getLastRow() > 1) {
            updateFormulasOnSheet(einnahmenTab);
            applyFormatting(einnahmenTab);
            einnahmenTab.autoResizeColumns(1, einnahmenTab.getLastColumn());
        }
    } else {
        Logger.log("Fehler: Der 'Einnahmen'-Ordner wurde nicht gefunden.");
    }

    if (ausgabenOrdner) {
        // Import in Rechnungen-Ausgangsblatt und Hauptübersicht "Ausgaben"
        importFilesFromFolder(ausgabenOrdner, sheetAusgaben, ausgabenTab, "Ausgabe", historyTab);
        // Spaltenbreite in Rechnungen-Ausgang anpassen
        sheetAusgaben.autoResizeColumns(1, sheetAusgaben.getLastColumn());
        if (ausgabenTab.getLastRow() > 1) {
            updateFormulasOnSheet(ausgabenTab);
            applyFormatting(ausgabenTab);
            ausgabenTab.autoResizeColumns(1, ausgabenTab.getLastColumn());
        }
    } else {
        Logger.log("Fehler: Der 'Ausgaben'-Ordner wurde nicht gefunden.");
    }
}

function importFilesFromFolder(folder, importSheet, mainSheet, type, historyTab) {
    const files = folder.getFiles();

    // Vorhandene Dateinamen in Hauptblatt und Import-Sheet ermitteln:
    const existingMainData = mainSheet.getDataRange().getValues();
    const existingMainFiles = new Set(existingMainData.slice(1).map(row => row[1])); // Spalte B: Dateiname

    const existingImportData = importSheet.getDataRange().getValues();
    const existingImportFiles = new Set(existingImportData.slice(1).map(row => row[0])); // Spalte 1

    const newImportRows = [];
    const newMainRows = [];
    const newHistoryRows = [];

    // Startindex für neue Zeilen im Hauptblatt:
    const mainStartRow = mainSheet.getLastRow() + 1;

    while (files.hasNext()) {
        const file = files.next();
        const fileNameWithExtension = file.getName();
        const fileName = fileNameWithExtension.replace(/\.[^/.]+$/, ""); // Dateiendung entfernen
        const fileUrl = file.getUrl();
        const timestamp = new Date();
        let wasImported = false;

        // Nur hinzufügen, wenn Datei im Hauptblatt noch nicht vorhanden ist:
        if (!existingMainFiles.has(fileName)) {
            const currentRow = mainStartRow + newMainRows.length;
            newMainRows.push([
                timestamp,               // Spalte A: Datum
                fileName,                // Spalte B: Dateiname
                "",                      // Spalte C: Kategorie (neu)
                "",                      // Spalte D: Kunde (ursprünglich Spalte C)
                "",                      // Spalte E: Nettobetrag (ursprünglich Spalte D)
                "",                      // Spalte F: Prozentsatz (ursprünglich Spalte E, als Prozent formatiert)
                `=E${currentRow}*F${currentRow}`, // Spalte G: MwSt.-Betrag (berechnet aus E und F)
                `=E${currentRow}+G${currentRow}`, // Spalte H: Bruttobetrag (E plus MwSt)
                "",                      // Spalte I: (leer, wie zuvor)
                `=E${currentRow}-(I${currentRow}-G${currentRow})`, // Spalte J: Restbetrag Netto
                `=IF(A${currentRow}=""; ""; ROUNDUP(MONTH(A${currentRow})/3;0))`, // Spalte K: Quartal
                `=IF(I${currentRow}>=H${currentRow}; "Bezahlt"; "Offen")`, // Spalte L: Zahlungsstatus
                "",                      // Spalte M: (leer)
                fileName,                // Spalte N: Dateiname (wiederholt)
                fileUrl,                 // Spalte O: Link zur Datei
                timestamp                // Spalte P: Letzte Aktualisierung
            ]);
            existingMainFiles.add(fileName);
            wasImported = true;
        }

        // Eintrag im Import-Sheet hinzufügen, wenn noch nicht vorhanden:
        if (!existingImportFiles.has(fileName)) {
            newImportRows.push([fileName, fileUrl, fileName]);
            existingImportFiles.add(fileName);
            wasImported = true;
        }

        // Historieneintrag nur, wenn die Datei tatsächlich neu importiert wurde:
        if (wasImported) {
            newHistoryRows.push([timestamp, type, fileName, fileUrl]);
        }
    }

    if (newImportRows.length > 0) {
        const importStartRow = importSheet.getLastRow() + 1;
        importSheet.getRange(importStartRow, 1, newImportRows.length, newImportRows[0].length).setValues(newImportRows);
    }

    if (newMainRows.length > 0) {
        mainSheet.getRange(mainStartRow, 1, newMainRows.length, newMainRows[0].length).setValues(newMainRows);
    }

    if (newHistoryRows.length > 0) {
        const historyStartRow = historyTab.getLastRow() + 1;
        historyTab.getRange(historyStartRow, 1, newHistoryRows.length, newHistoryRows[0].length).setValues(newHistoryRows);
    }
}

// Hilfsfunktion: Sucht einen Unterordner anhand des Namens
function getFolderByName(parentFolder, folderName) {
    const folderIter = parentFolder.getFoldersByName(folderName);
    return folderIter.hasNext() ? folderIter.next() : null;
}

/* ----- Funktionen ausschließlich für die Tabellen "Einnahmen" und "Ausgaben" ----- */
function sortCurrentSheetByColumn() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const name = sheet.getName();
    if (name !== "Einnahmen" && name !== "Ausgaben") {
        SpreadsheetApp.getUi().alert("Diese Funktion ist nur für die Blätter 'Einnahmen' und 'Ausgaben' verfügbar.");
        return;
    }
    sortSheetByColumn(sheet, 2);
}

function updateSheetOnCurrentSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const name = sheet.getName();
    if (name !== "Einnahmen" && name !== "Ausgaben") {
        SpreadsheetApp.getUi().alert("Diese Funktion ist nur für die Blätter 'Einnahmen' und 'Ausgaben' verfügbar.");
        return;
    }
    updateFormulasOnSheet(sheet);
    applyFormatting(sheet);
    sheet.autoResizeColumns(1, sheet.getLastColumn());
}

function sortSheetByColumn(sheet, column) {
    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length <= 1) {
        Logger.log("Keine Daten zum Sortieren.");
        return;
    }
    const header = values.shift();
    values.sort((a, b) => {
        const valA = a[column - 1];
        const valB = b[column - 1];
        if (typeof valA === "number" && typeof valB === "number") {
            return valA - valB;
        }
        return String(valA).localeCompare(String(valB));
    });
    sheet.clearContents();
    sheet.appendRow(header);
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
    updateFormulasOnSheet(sheet);
}

function updateFormulasOnSheet(sheet) {
    const lastRow = sheet.getLastRow();
    const numRows = lastRow - 1;
    if (numRows < 1) return;

    const formulas6 = [];
    const formulas7 = [];
    const formulas9 = [];
    const formulas10 = [];
    const formulas11 = [];

    for (let i = 2; i <= lastRow; i++) {
        formulas6.push([`=E${i}*F${i}`]); // MwSt.-Betrag (Spalte G)
        formulas7.push([`=E${i}+G${i}`]); // Bruttobetrag (Spalte H)
        formulas9.push([`=E${i}-(I${i}-G${i})`]); // Restbetrag Netto (Spalte J)
        formulas10.push([`=IF(A${i}=""; ""; ROUNDUP(MONTH(A${i})/3;0))`]); // Quartal (Spalte K)
        formulas11.push([`=IF(I${i}>=H${i}; "Bezahlt"; "Offen")`]); // Zahlungsstatus (Spalte L)
    }

    sheet.getRange(2, 7, numRows, 1).setFormulas(formulas6);
    sheet.getRange(2, 8, numRows, 1).setFormulas(formulas7);
    sheet.getRange(2, 10, numRows, 1).setFormulas(formulas9);
    sheet.getRange(2, 11, numRows, 1).setFormulas(formulas10);
    sheet.getRange(2, 12, numRows, 1).setFormulas(formulas11);
}

function applyFormatting(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        // Datum in Spalte A
        sheet.getRange(`A2:A${lastRow}`).setNumberFormat("DD.MM.YYYY");
        // Letzte Aktualisierung in Spalte P
        sheet.getRange(`P2:P${lastRow}`).setNumberFormat("DD.MM.YYYY");
        // Nettobetrag in Spalte E als Währung
        sheet.getRange(`E2:E${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");
        // MwSt, Bruttobetrag und Restbetrag in den Spalten G, H und J als Währung
        sheet.getRange(`G2:G${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");
        sheet.getRange(`H2:H${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");
        sheet.getRange(`J2:J${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");
        // Prozentsatz in Spalte F
        sheet.getRange(`F2:F${lastRow}`).setNumberFormat("0.00%");
    }
}

/* ----- Neue Funktion: GUV-Berechnung ----- */
function calculateGUV() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const einnahmenSheet = ss.getSheetByName("Einnahmen");
    const ausgabenSheet = ss.getSheetByName("Ausgaben");

    if (!einnahmenSheet || !ausgabenSheet) {
        SpreadsheetApp.getUi().alert("Eines der Blätter 'Einnahmen' oder 'Ausgaben' wurde nicht gefunden.");
        return;
    }

    const einnahmenData = einnahmenSheet.getDataRange().getValues().slice(1);
    const ausgabenData = ausgabenSheet.getDataRange().getValues().slice(1);

    let fehlendeDaten = [];

    let guvData = {};
    for (let m = 1; m <= 12; m++) {
        guvData[m] = { einnahmen: 0, einnahmenOffen: 0, ausgaben: 0, ausgabenOffen: 0,
            umsatzsteuer: 0, vorsteuer: 0, ustZahlung: 0, ergebnis: 0 };
    }

    einnahmenData.forEach((row, index) => {
        let date = row[0];
        if (!(date instanceof Date)) {
            fehlendeDaten.push(`Einnahmen - Zeile ${index + 2}`);
            return;
        }
        let monat = date.getMonth() + 1;
        let netto = parseFloat(row[4]) || 0;
        let mwst = parseFloat(row[6]) || 0;
        let bezahlt = parseFloat(row[8]) || 0;

        guvData[monat].einnahmen += bezahlt;
        guvData[monat].umsatzsteuer += bezahlt > 0 ? mwst : 0;
        guvData[monat].einnahmenOffen += netto - bezahlt;
    });

    ausgabenData.forEach((row, index) => {
        let date = row[0];
        if (!(date instanceof Date)) {
            fehlendeDaten.push(`Ausgaben - Zeile ${index + 2}`);
            return;
        }
        let monat = date.getMonth() + 1;
        let netto = parseFloat(row[4]) || 0;
        let mwst = parseFloat(row[6]) || 0;
        let bezahlt = parseFloat(row[8]) || 0;

        guvData[monat].ausgaben += bezahlt;
        guvData[monat].vorsteuer += bezahlt > 0 ? mwst : 0;
        guvData[monat].ausgabenOffen += netto - bezahlt;
    });

    if (fehlendeDaten.length > 0) {
        SpreadsheetApp.getUi().alert("Fehler: Es gibt Einträge ohne Datum! Bitte überprüfe folgende Zeilen:\n" + fehlendeDaten.join("\n"));
        return;
    }

    let guvSheet = ss.getSheetByName("GUV");
    if (!guvSheet) {
        guvSheet = ss.insertSheet("GUV");
    } else {
        guvSheet.clearContents();
    }

    guvSheet.appendRow(["Zeitraum", "Einnahmen", "Offene Forderungen", "Ausgaben", "Offene Verbindlichkeiten",
        "Umsatzsteuer", "Vorsteuer", "USt-Zahlung", "Ergebnis"]);

    let quartalsDaten = {1: {}, 2: {}, 3: {}, 4: {}};
    for (let q = 1; q <= 4; q++) {
        quartalsDaten[q] = { einnahmen: 0, einnahmenOffen: 0, ausgaben: 0, ausgabenOffen: 0,
            umsatzsteuer: 0, vorsteuer: 0, ustZahlung: 0, ergebnis: 0 };
    }

    for (let m = 1; m <= 12; m++) {
        let q = Math.ceil(m / 3);
        let data = guvData[m];
        quartalsDaten[q].einnahmen += data.einnahmen;
        quartalsDaten[q].einnahmenOffen += data.einnahmenOffen;
        quartalsDaten[q].ausgaben += data.ausgaben;
        quartalsDaten[q].ausgabenOffen += data.ausgabenOffen;
        quartalsDaten[q].umsatzsteuer += data.umsatzsteuer;
        quartalsDaten[q].vorsteuer += data.vorsteuer;
        quartalsDaten[q].ustZahlung += data.umsatzsteuer - data.vorsteuer;
        quartalsDaten[q].ergebnis += data.einnahmen - data.ausgaben;

        guvSheet.appendRow([`Monat ${m}`, data.einnahmen, data.einnahmenOffen, data.ausgaben, data.ausgabenOffen,
            data.umsatzsteuer, data.vorsteuer, data.umsatzsteuer - data.vorsteuer, data.einnahmen - data.ausgaben]);
    }

    for (let q = 1; q <= 4; q++) {
        let data = quartalsDaten[q];
        guvSheet.appendRow([`Quartal ${q}`, data.einnahmen, data.einnahmenOffen, data.ausgaben, data.ausgabenOffen,
            data.umsatzsteuer, data.vorsteuer, data.umsatzsteuer - data.vorsteuer, data.ergebnis]);
    }

    // Gesamtjahr berechnen
    let gesamtjahr = [0,0,0,0,0,0,0,0];
    for (let q = 1; q <= 4; q++) {
        let data = quartalsDaten[q];
        gesamtjahr[0] += data.einnahmen;
        gesamtjahr[1] += data.einnahmenOffen;
        gesamtjahr[2] += data.ausgaben;
        gesamtjahr[3] += data.ausgabenOffen;
        gesamtjahr[4] += data.umsatzsteuer;
        gesamtjahr[5] += data.vorsteuer;
        gesamtjahr[6] += data.umsatzsteuer - data.vorsteuer;
        gesamtjahr[7] += data.ergebnis;
    }
    guvSheet.appendRow(["Gesamtjahr", "", "", "", "", "", "", "", gesamtjahr[7]]);

    let lastRow = guvSheet.getLastRow();
    guvSheet.getRange(`B2:I${lastRow}`).setNumberFormat("#,##0.00 €");

    // Automatisch Spaltenbreiten anpassen (Fit-to-Content)
    guvSheet.autoResizeColumns(1, guvSheet.getLastColumn());

    SpreadsheetApp.getUi().alert("GUV-Berechnung abgeschlossen und aktualisiert.");
}
