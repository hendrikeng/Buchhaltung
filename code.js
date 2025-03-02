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
    ui.createMenu("📂 Buchhaltung")
        .addItem("📥 Rechnungen importieren", "importDriveFiles")
        .addItem("🔄 Aktualisieren (Formeln & Formatierung)", "updateSheetOnCurrentSheet")
        .addItem("📊 GUV berechnen", "calculateGUV")
        .addItem("📈 BWA berechnen", "calculateBWA")
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
                "",                      // Spalte D: Kunde
                "",                      // Spalte E: Nettobetrag
                "",                      // Spalte F: Prozentsatz (als Prozent formatiert)
                `=E${currentRow}*F${currentRow}`, // Spalte G: MwSt.-Betrag
                `=E${currentRow}+G${currentRow}`, // Spalte H: Bruttobetrag
                "",                      // Spalte I: (wird als Währung formatiert)
                `=E${currentRow}-(I${currentRow}-G${currentRow})`, // Spalte J: Restbetrag Netto
                `=IF(A${currentRow}=""; ""; ROUNDUP(MONTH(A${currentRow})/3;0))`, // Spalte K: Quartal
                `=IF(I${currentRow}>=H${currentRow}; "Bezahlt"; "Offen")`, // Spalte L: Zahlungsstatus
                "",                      // Spalte M: (wird als Datum formatiert)
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
        formulas11.push([`=IF(OR(I${i}=""; I${i}=0); "Offen"; IF(I${i}>=H${i}; "Bezahlt"; "Offen"))`]); // Zahlungsstatus (Spalte L)
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
        // Spalte A (Datum)
        sheet.getRange(`A2:A${lastRow}`).setNumberFormat("DD.MM.YYYY");

        // Spalte M (z.B. weiterer Datumswert)
        sheet.getRange(`M2:M${lastRow}`).setNumberFormat("DD.MM.YYYY");

        // Spalte I als Währung (Accounting)
        sheet.getRange(`I2:I${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");

        // Spalte P (Letzte Aktualisierung, Datum)
        sheet.getRange(`P2:P${lastRow}`).setNumberFormat("DD.MM.YYYY");

        // Spalte E (Nettobetrag) als Währung
        sheet.getRange(`E2:E${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");

        // MwSt, Bruttobetrag und Restbetrag in G, H und J als Währung
        sheet.getRange(`G2:G${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");
        sheet.getRange(`H2:H${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");
        sheet.getRange(`J2:J${lastRow}`).setNumberFormat("€#,##0.00;€-#,##0.00");

        // Spalte F (Prozentsatz)
        sheet.getRange(`F2:F${lastRow}`).setNumberFormat("0.00%");
    }
}

/**
 * calculateGUV (Ist-Besteuerung) – optimierte Version mit MwSt.-Differenzierung
 *
 * 📌 Zweck:
 *  - Erstellt eine Gewinn- und Verlustrechnung (GuV) nach dem Prinzip der Ist-Besteuerung.
 *  - Berücksichtigt Einnahmen und Ausgaben nur, wenn sie tatsächlich bezahlt wurden.
 *  - Offene Posten (kein Zahlungsdatum) erscheinen nur in der Jahresübersicht,
 *    nicht in den Monats- oder Quartalswerten.
 *
 * ✅ Verbesserungen & Fixes:
 *  - Dynamische Berechnung des MwSt.-Satzes (0%, 7%, 19%) mit separaten Spalten.
 *  - Validierung des Zahlungsdatums (keine zukünftigen Zahlungen erlaubt).
 *  - Gutschriften & Erstattungen (negative Werte korrekt berücksichtigt).
 *  - Korrekte Berechnung des Gewinns (ohne Umsatzsteuerabzug).
 *  - Fix für doppelte Addition von offenen Forderungen/Verbindlichkeiten in der Jahresübersicht.
 *
 * 📂 Benötigte Tabellenblätter:
 *  - "Einnahmen" (benötigte Spalten: Netto-Betrag, MwSt.-Satz, Bezahlter Brutto-Betrag, Zahlungsdatum).
 *  - "Ausgaben" (benötigte Spalten wie bei Einnahmen).
 *  - "GUV" (wird automatisch erzeugt oder aktualisiert).
 */
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

    let guvData = {},
        quartalsDaten = {},
        fehlendeDaten = [];

    // Offene Beträge als Objekt, damit wir sie innerhalb von processRow verändern können.
    let offeneTotals = { einnahmen: 0, ausgaben: 0 };

    // Initialisiere Monatsdaten
    for (let m = 1; m <= 12; m++) {
        guvData[m] = {
            einnahmen: 0, einnahmenOffen: 0,
            ausgaben: 0, ausgabenOffen: 0,
            ust_0: 0, ust_7: 0, ust_19: 0,
            vst_0: 0, vst_7: 0, vst_19: 0,
            ustZahlung: 0, ergebnis: 0
        };
    }
    // Initialisiere Quartalsdaten als Kopie der Monatsstruktur (z.B. von Monat 1)
    for (let q = 1; q <= 4; q++) {
        quartalsDaten[q] = JSON.parse(JSON.stringify(guvData[1]));
    }

    // Verarbeitung der Einnahmen und Ausgaben (offeneTotals wird mitgegeben)
    einnahmenData.forEach((row, index) => processGuVRow(row, index, guvData, quartalsDaten, true, fehlendeDaten, offeneTotals));
    ausgabenData.forEach((row, index) => processGuVRow(row, index, guvData, quartalsDaten, false, fehlendeDaten, offeneTotals));

    // Falls Fehler (Bezahlung ohne Datum oder zukünftiges Datum) vorhanden sind, abbrechen
    if (fehlendeDaten.length > 0) {
        SpreadsheetApp.getUi().alert("Fehlende oder fehlerhafte Zahlungsdaten gefunden:\n" + fehlendeDaten.join("\n"));
        return;
    }

    // Aufbau des GUV-Sheets
    let guvSheet = ss.getSheetByName("GUV") || ss.insertSheet("GUV");
    guvSheet.clearContents();
    guvSheet.appendRow(["Zeitraum", "Einnahmen (netto)", "Offene Forderungen", "Ausgaben (netto)", "Offene Verbindlichkeiten",
        "USt 0%", "USt 7%", "USt 19%", "VSt 0%", "VSt 7%", "VSt 19%", "USt-Zahlung", "Ergebnis"]);

    for (let m = 1; m <= 12; m++) {
        appendRowToSheet(guvSheet, `Monat ${m}`, guvData[m]);
        if (m % 3 === 0) {
            let q = Math.ceil(m / 3);
            appendRowToSheet(guvSheet, `Quartal ${q}`, quartalsDaten[q]);
        }
    }

    let jahresSumme = calculateYearlySum(guvData);
    // Offene Beträge aus Zeilen ohne Datum hinzufügen
    jahresSumme.einnahmenOffen += offeneTotals.einnahmen;
    jahresSumme.ausgabenOffen += offeneTotals.ausgaben;

    appendRowToSheet(guvSheet, "Gesamtjahr", jahresSumme);

    guvSheet.getRange(`B2:M${guvSheet.getLastRow()}`).setNumberFormat("#,##0.00€");
    guvSheet.autoResizeColumns(1, guvSheet.getLastColumn());
    SpreadsheetApp.getUi().alert("GUV-Berechnung abgeschlossen und aktualisiert.");
}

/**
 * calculateGUV (Ist-Besteuerung) – optimierte Version mit MwSt.-Differenzierung
 *
 * 📌 Zweck:
 *  - Erstellt eine Gewinn- und Verlustrechnung (GuV) nach dem Prinzip der Ist-Besteuerung.
 *  - Berücksichtigt Einnahmen und Ausgaben nur, wenn sie tatsächlich bezahlt wurden.
 *  - Offene Posten (kein Zahlungsdatum) erscheinen nur in der Jahresübersicht,
 *    nicht in den Monats- oder Quartalswerten.
 *
 * ✅ Verbesserungen & Fixes:
 *  - Dynamische Berechnung des MwSt.-Satzes (0%, 7%, 19%) mit separaten Spalten.
 *  - Validierung des Zahlungsdatums (keine zukünftigen Zahlungen erlaubt).
 *  - Gutschriften & Erstattungen (negative Werte korrekt berücksichtigt).
 *  - Korrekte Berechnung des Gewinns (ohne Umsatzsteuerabzug).
 *  - Fix für doppelte Addition von offenen Forderungen/Verbindlichkeiten in der Jahresübersicht.
 *
 * 📂 Benötigte Tabellenblätter:
 *  - "Einnahmen" (benötigte Spalten: Netto-Betrag, MwSt.-Satz, Bezahlter Brutto-Betrag, Zahlungsdatum).
 *  - "Ausgaben" (benötigte Spalten wie bei Einnahmen).
 *  - "GUV" (wird automatisch erzeugt oder aktualisiert).
 */
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

    let guvData = {},
        quartalsDaten = {},
        fehlendeDaten = [];

    // Offene Beträge als Objekt, damit wir sie innerhalb von processRow verändern können.
    let offeneTotals = { einnahmen: 0, ausgaben: 0 };

    // Initialisiere Monatsdaten
    for (let m = 1; m <= 12; m++) {
        guvData[m] = {
            einnahmen: 0, einnahmenOffen: 0,
            ausgaben: 0, ausgabenOffen: 0,
            ust_0: 0, ust_7: 0, ust_19: 0,
            vst_0: 0, vst_7: 0, vst_19: 0,
            ustZahlung: 0, ergebnis: 0
        };
    }
    // Initialisiere Quartalsdaten als Kopie der Monatsstruktur (z.B. von Monat 1)
    for (let q = 1; q <= 4; q++) {
        quartalsDaten[q] = JSON.parse(JSON.stringify(guvData[1]));
    }

    // Verarbeitung der Einnahmen und Ausgaben (offeneTotals wird mitgegeben)
    einnahmenData.forEach((row, index) => processGuVRow(row, index, guvData, quartalsDaten, true, fehlendeDaten, offeneTotals));
    ausgabenData.forEach((row, index) => processGuVRow(row, index, guvData, quartalsDaten, false, fehlendeDaten, offeneTotals));

    // Falls Fehler (Bezahlung mit ungültigem Datum oder zukünftiges Datum) vorhanden sind, abbrechen
    if (fehlendeDaten.length > 0) {
        SpreadsheetApp.getUi().alert("Fehlende oder fehlerhafte Zahlungsdaten gefunden:\n" + fehlendeDaten.join("\n"));
        return;
    }

    // Aufbau des GUV-Sheets
    let guvSheet = ss.getSheetByName("GUV") || ss.insertSheet("GUV");
    guvSheet.clearContents();
    guvSheet.appendRow(["Zeitraum", "Einnahmen (netto)", "Offene Forderungen", "Ausgaben (netto)", "Offene Verbindlichkeiten",
        "USt 0%", "USt 7%", "USt 19%", "VSt 0%", "VSt 7%", "VSt 19%", "USt-Zahlung", "Ergebnis"]);

    for (let m = 1; m <= 12; m++) {
        appendRowToSheet(guvSheet, `Monat ${m}`, guvData[m]);
        if (m % 3 === 0) {
            let q = Math.ceil(m / 3);
            appendRowToSheet(guvSheet, `Quartal ${q}`, quartalsDaten[q]);
        }
    }

    let jahresSumme = calculateYearlySum(guvData);
    // Offene Beträge aus Zeilen ohne Datum hinzufügen
    jahresSumme.einnahmenOffen += offeneTotals.einnahmen;
    jahresSumme.ausgabenOffen += offeneTotals.ausgaben;

    appendRowToSheet(guvSheet, "Gesamtjahr", jahresSumme);

    guvSheet.getRange(`B2:M${guvSheet.getLastRow()}`).setNumberFormat("#,##0.00€");
    guvSheet.autoResizeColumns(1, guvSheet.getLastColumn());
    SpreadsheetApp.getUi().alert("GUV-Berechnung abgeschlossen und aktualisiert.");
}

/**
 * Verarbeitet eine einzelne Zeile (Einnahmen oder Ausgaben) und fügt die Werte den
 * entsprechenden Monats- und Quartalsdaten hinzu.
 *
 * Neben den bereits vorhandenen Prüfungen wird hier auch der Fall behandelt,
 * dass kein Zahlungsdatum gesetzt wurde. In diesem Fall wird der Nettobetrag
 * zu den offenen Beträgen (je nach Kategorie) addiert.
 */
function processGuVRow(row, index, guvData, quartalsDaten, isIncome, fehlendeDaten, offeneTotals) {
    let zahlungsDatum = parseDate(row[12]);
    let netto = parseCurrency(row[4]);
    let mwstRate = parseMwstRate(row[5]);
    let bezahltBrutto = parseCurrency(row[8]);
    let factor = (mwstRate > 0) ? 1 + (mwstRate / 100) : 1;
    let bezahltNetto = bezahltBrutto / factor;
    let bezahltMwst = bezahltBrutto - bezahltNetto;
    let restbetrag = Math.max(0, netto - bezahltNetto);

    // Falls eine Zahlung erfolgt ist, aber kein Datum gesetzt wurde -> Fehler
    if (bezahltBrutto !== 0 && !zahlungsDatum) {
        fehlendeDaten.push(`${isIncome ? "Einnahmen" : "Ausgaben"} - Zeile ${index + 2}: Fehlendes Zahlungsdatum trotz Zahlung!`);
        return;
    }

    // Falls Zahlungsdatum in der Zukunft liegt -> Fehler
    if (zahlungsDatum && zahlungsDatum > new Date()) {
        fehlendeDaten.push(`${isIncome ? "Einnahmen" : "Ausgaben"} - Zeile ${index + 2}: Zahlungsdatum in der Zukunft!`);
        return;
    }

    let monat = zahlungsDatum ? zahlungsDatum.getMonth() + 1 : null;
    let quartal = zahlungsDatum ? Math.ceil(monat / 3) : null;
    let category = isIncome ? "ust" : "vst";

    // Falls eine Rechnung offen ist -> Zu offenen Forderungen/Verbindlichkeiten hinzufügen
    if (!zahlungsDatum || restbetrag > 0) {
        if (isIncome) {
            offeneTotals.einnahmen += restbetrag;
        } else {
            offeneTotals.ausgaben += restbetrag;
        }
    }

    if (!zahlungsDatum) return; // Offene Rechnungen werden nicht für Umsatzsteuer berechnet

    // Sicherstellen, dass die Datenstrukturen existieren
    if (!guvData[monat]) {
        guvData[monat] = createEmptyGuVObject();
    }
    if (!quartalsDaten[quartal]) {
        quartalsDaten[quartal] = createEmptyGuVObject();
    }

    guvData[monat][`${category}_${mwstRate}`] += bezahltMwst;
    quartalsDaten[quartal][`${category}_${mwstRate}`] += bezahltMwst;

    guvData[monat][isIncome ? "einnahmen" : "ausgaben"] += bezahltNetto;
    quartalsDaten[quartal][isIncome ? "einnahmen" : "ausgaben"] += bezahltNetto;
}

function createEmptyGuVObject() {
    return {
        einnahmen: 0, einnahmenOffen: 0,
        ausgaben: 0, ausgabenOffen: 0,
        ust_0: 0, ust_7: 0, ust_19: 0,
        vst_0: 0, vst_7: 0, vst_19: 0,
        ustZahlung: 0, ergebnis: 0
    };
}

/**
 * Berechnet die Jahressumme über alle Monatswerte.
 */
function calculateYearlySum(guvData) {
    let sum = JSON.parse(JSON.stringify(guvData[1]));
    for (let m = 2; m <= 12; m++) {
        Object.keys(sum).forEach(k => sum[k] += guvData[m][k]);
    }
    return sum;
}

/**
 * Hängt eine Zeile mit den Daten zum angegebenen Label an das Sheet an.
 */
function appendRowToSheet(sheet, label, data) {
    sheet.appendRow([
        label,
        data.einnahmen,
        data.einnahmenOffen,
        data.ausgaben,
        data.ausgabenOffen,
        data.ust_0,
        data.ust_7,
        data.ust_19,
        data.vst_0,
        data.vst_7,
        data.vst_19,
        data.ust_19 - data.vst_19,
        data.einnahmen - data.ausgaben
    ]);
}

/**
 * parseDate:
 * - Erzeugt ein Date-Objekt aus einem Wert.
 * - Gibt null zurück, falls der Wert kein gültiges Datum darstellt.
 */
function parseDate(value) {
    let d = (typeof value === "string") ? new Date(value) : (value instanceof Date ? value : null);
    // Wenn das Datum ungültig ist (getTime() => NaN), gib null zurück
    if (d && isNaN(d.getTime())) {
        return null;
    }
    return d;
}

/**
 * parseCurrency:
 * - Wandelt einen Währungs-/Zahlenstring in eine float-Variable um
 *   (z. B. "1.234,56" → 1234.56).
 */
function parseCurrency(value) {
    return parseFloat(
        value
            .toString()
            .replace(/[^\d,.-]/g, "")
            .replace(",", ".")
    ) || 0;
}

/**
 * parseMwstRate:
 * - Konvertiert einen MwSt.-Wert in Prozent (0, 7, 19, etc.).
 * - Beispiel: 0,07 => 7, "19%" => 19, "0,07" => 7, usw.
 */
function parseMwstRate(value) {
    let rate = (typeof value === "number")
        ? (value < 1 ? value * 100 : value)
        : parseFloat(value.toString().replace("%", "").replace(",", "."));
    return isNaN(rate) ? 19 : rate;
}

/**
 * Liefert das BWA-Datenobjekt für einen einzelnen Monat,
 * das alle benötigten Felder enthält.
 */
function createEmptyBwaMonth() {
    return {
        // Einnahmen-Kategorien
        dienstleistungen: 0,
        produkte: 0,
        sonstigeEinnahmen: 0,
        gesamtEinnahmen: 0,

        // Ausgaben-Kategorien
        betriebskosten: 0,
        marketing: 0,
        reisen: 0,
        einkauf: 0,
        personal: 0,
        gesamtAusgaben: 0,

        // Offene Posten
        offeneForderungen: 0,
        offeneVerbindlichkeiten: 0,

        // Ergebnis- und Liquiditätskennzahlen
        rohertrag: 0,
        betriebsergebnis: 0,
        ergebnisVorSteuern: 0,
        ergebnisNachSteuern: 0,

        // Bank-/Liquiditätsstand (wird als Endsaldo pro Monat gesetzt)
        liquiditaet: 0
    };
}

/**
 * categoryMapping:
 * Weist die Textkategorien den Feldern im BWA-Objekt zu.
 * Falls unbekannt, weicht man auf "sonstigeEinnahmen" (bei Einnahmen)
 * oder "betriebskosten" (bei Ausgaben) aus.
 */
function getBwaCategory(category, isIncome, rowIndex, fehlendeKategorien) {
    const map = {
        "Dienstleistungen": "dienstleistungen",
        "Produkte & Waren": "produkte",
        "Sonstige Einnahmen": "sonstigeEinnahmen",
        "Betriebskosten": "betriebskosten",
        "Marketing & Werbung": "marketing",
        "Reisen & Mobilität": "reisen",
        "Wareneinkauf & Dienstleistungen": "einkauf",
        "Personal & Gehälter": "personal"
    };

    if (!category || !map[category]) {
        let fallback = isIncome ? "sonstigeEinnahmen" : "betriebskosten";
        fehlendeKategorien.push(
            `Zeile ${rowIndex}: Unbekannte Kategorie "${category || "N/A"}" → Verwende "${fallback}"`
        );
        return fallback;
    }
    return map[category];
}

/**
 * processBwaRow:
 * Verarbeitet eine einzelne Zeile aus Einnahmen/Ausgaben.
 * - Prüft das Zahlungsdatum, ggf. Fehlermeldung bei ungültigen Daten.
 * - Spalte [8] enthält den bezahlten Bruttobetrag; wir ermitteln daraus den Nettoanteil.
 * - Offene Posten (bei Teilzahlungen oder fehlendem Datum) werden separat erfasst.
 */
function processBwaRow(row, index, bwaData, isIncome, fehlendeDaten, fehlendeKategorien) {
    // Annahme:
    // row[2] = Kategorie
    // row[4] = Rechnungs-Netto (z.B. Gesamtnetto)
    // row[5] = MwSt.-Satz (z.B. 0, 7, 19)
    // row[8] = Bezahlter Bruttobetrag
    // row[12] = Zahlungsdatum
    const category = row[2];
    const nettoRechnung = parseCurrency(row[4]);
    const mwstRate = parseMwstRate(row[5]);
    const bezahltBrutto = parseCurrency(row[8]);
    const zahlungsDatum = parseDate(row[12]);

    const factor = mwstRate > 0 ? 1 + (mwstRate / 100) : 1;
    const bezahltNetto = bezahltBrutto / factor;

    // Datum-Checks
    if (bezahltBrutto !== 0 && !zahlungsDatum) {
        fehlendeDaten.push(
            `${isIncome ? "Einnahmen" : "Ausgaben"} - Zeile ${index + 2}: Zahlung ohne gültiges Zahlungsdatum!`
        );
        return;
    }
    if (zahlungsDatum && zahlungsDatum > new Date()) {
        fehlendeDaten.push(
            `${isIncome ? "Einnahmen" : "Ausgaben"} - Zeile ${index + 2}: Zahlungsdatum liegt in der Zukunft!`
        );
        return;
    }
    if (!zahlungsDatum) {
        if (isIncome) {
            bwaData.offeneForderungen += nettoRechnung;
        } else {
            bwaData.offeneVerbindlichkeiten += nettoRechnung;
        }
        return;
    }

    const monat = zahlungsDatum.getMonth() + 1;
    const mappedCat = getBwaCategory(category, isIncome, index + 2, fehlendeKategorien);

    const restOffen = Math.max(0, nettoRechnung - bezahltNetto);
    if (restOffen > 0) {
        if (isIncome) {
            bwaData.offeneForderungen += restOffen;
        } else {
            bwaData.offeneVerbindlichkeiten += restOffen;
        }
    }

    if (!bwaData.monats[monat]) {
        bwaData.monats[monat] = createEmptyBwaMonth();
    }
    let monthObj = bwaData.monats[monat];
    if (isIncome) {
        monthObj[mappedCat] += bezahltNetto;
        monthObj.gesamtEinnahmen += bezahltNetto;
    } else {
        monthObj[mappedCat] += bezahltNetto;
        monthObj.gesamtAusgaben += bezahltNetto;
    }
}

/**
 * calculateBWA:
 * Liest Daten aus "Einnahmen", "Ausgaben" und "Bankbewegungen" und erstellt
 * ein "BWA"-Sheet mit Monats-, Quartals- und Jahreswerten in folgender Reihenfolge:
 *
 * Monat 1
 * Monat 2
 * Monat 3
 * Quartal 1 (Zusammenfassung von Monat 1–3)
 * Monat 4
 * Monat 5
 * Monat 6
 * Quartal 2
 * …
 * Monat 10
 * Monat 11
 * Monat 12
 * Quartal 4
 * Gesamtjahr
 *
 * In Spalte [8] wird der bezahlte Bruttobetrag erwartet, der via MwSt.-Satz
 * in einen Nettoanteil umgerechnet wird.
 * Die Liquidität (Bankbestand) wird als Endsaldo der letzten Buchung des Monats übernommen.
 * Fehlt in einem Monat ein Bankwert, wird der letzte bekannte Wert (Carry Forward) übernommen.
 */
function calculateBWA() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const einnahmenSheet = ss.getSheetByName("Einnahmen");
    const ausgabenSheet = ss.getSheetByName("Ausgaben");
    const bankSheet = ss.getSheetByName("Bankbewegungen");

    if (!einnahmenSheet || !ausgabenSheet || !bankSheet) {
        SpreadsheetApp.getUi().alert("Fehlende Tabellenblätter! Bitte prüfe 'Einnahmen', 'Ausgaben' und 'Bankbewegungen'.");
        return;
    }

    const einnahmenData = einnahmenSheet.getDataRange().getValues().slice(1);
    const ausgabenData = ausgabenSheet.getDataRange().getValues().slice(1);
    const bankData = bankSheet.getDataRange().getValues().slice(1);

    let bwaData = {
        monats: {},
        offeneForderungen: 0,
        offeneVerbindlichkeiten: 0
    };

    let fehlendeDaten = [];
    let fehlendeKategorien = [];

    // Einnahmen und Ausgaben verarbeiten
    einnahmenData.forEach((row, index) => {
        processBwaRow(row, index, bwaData, true, fehlendeDaten, fehlendeKategorien);
    });
    ausgabenData.forEach((row, index) => {
        processBwaRow(row, index, bwaData, false, fehlendeDaten, fehlendeKategorien);
    });

    // Bankbewegungen verarbeiten: Spalte 0 = Buchungsdatum, Spalte 3 = Saldo
    bankData.forEach((row) => {
        const buchungsDatum = parseDate(row[0]);
        if (!buchungsDatum || buchungsDatum > new Date()) return;
        const saldo = parseCurrency(row[3]);
        const monat = buchungsDatum.getMonth() + 1;
        if (!bwaData.monats[monat]) {
            bwaData.monats[monat] = createEmptyBwaMonth();
        }
        // Nur den Saldo der zuletzt datierten Buchung im Monat übernehmen:
        if (!bwaData.monats[monat].__lastBankDate || buchungsDatum > bwaData.monats[monat].__lastBankDate) {
            bwaData.monats[monat].liquiditaet = saldo;
            bwaData.monats[monat].__lastBankDate = buchungsDatum;
        }
    });
    // Entferne das Hilfsfeld __lastBankDate
    for (let m in bwaData.monats) {
        delete bwaData.monats[m].__lastBankDate;
    }

    // Carry-Forward: Wenn ein Monat keinen Bankwert hat (liquiditaet 0), übernehme den letzten bekannten Saldo.
    let lastKnownLiquid = 0;
    for (let m = 1; m <= 12; m++) {
        if (!bwaData.monats[m]) {
            bwaData.monats[m] = createEmptyBwaMonth();
        }
        if (bwaData.monats[m].liquiditaet === 0) {
            // Falls m > 1, setze den Wert vom Vormonat
            bwaData.monats[m].liquiditaet = lastKnownLiquid;
        } else {
            lastKnownLiquid = bwaData.monats[m].liquiditaet;
        }
    }

    if (fehlendeDaten.length > 0) {
        SpreadsheetApp.getUi().alert("Fehlerhafte Datensätze:\n" + fehlendeDaten.join("\n"));
        return;
    }

    // Monatsweise Kennzahlen berechnen
    for (let m = 1; m <= 12; m++) {
        let moObj = bwaData.monats[m];
        if (!moObj) {
            bwaData.monats[m] = createEmptyBwaMonth();
            moObj = bwaData.monats[m];
        }
        // Beispielberechnung: Rohertrag = gesamtEinnahmen - Einkauf
        moObj.rohertrag = moObj.gesamtEinnahmen - moObj.einkauf;
        // Betriebsergebnis = Rohertrag - (alle anderen Ausgaben außer Einkauf)
        let sonstigeAusgaben = moObj.gesamtAusgaben - moObj.einkauf;
        moObj.betriebsergebnis = moObj.rohertrag - sonstigeAusgaben;
        moObj.ergebnisVorSteuern = moObj.betriebsergebnis;
        moObj.ergebnisNachSteuern = moObj.ergebnisVorSteuern;
    }

    // --- Aggregation und Ausgabe ---
    let bwaSheet = ss.getSheetByName("BWA") || ss.insertSheet("BWA");
    bwaSheet.clearContents();

    // Kopfzeile
    bwaSheet.appendRow([
        "Zeitraum",
        "Dienstleistungen", "Produkte", "Sonst. Einnahmen", "Gesamt-Einnahmen",
        "Betriebskosten", "Marketing", "Reisen", "Einkauf", "Personal", "Gesamt-Ausgaben",
        "Offene Forderungen", "Offene Verbindlichkeiten",
        "Rohertrag", "Betriebsergebnis", "Ergebnis vor Steuern", "Ergebnis nach Steuern",
        "Liquidität"
    ]);

    // Hilfsfunktion, um ein Monatsobjekt in eine Zeile zu packen
    function toRowData(moObj, label) {
        return [
            label,
            moObj.dienstleistungen,
            moObj.produkte,
            moObj.sonstigeEinnahmen,
            moObj.gesamtEinnahmen,
            moObj.betriebskosten,
            moObj.marketing,
            moObj.reisen,
            moObj.einkauf,
            moObj.personal,
            moObj.gesamtAusgaben,
            moObj.offeneForderungen,
            moObj.offeneVerbindlichkeiten,
            moObj.rohertrag,
            moObj.betriebsergebnis,
            moObj.ergebnisVorSteuern,
            moObj.ergebnisNachSteuern,
            moObj.liquiditaet
        ];
    }

    // Quartals- und Jahresaggregation vorberechnen:
    let quartalsSum = {
        1: createEmptyBwaMonth(),
        2: createEmptyBwaMonth(),
        3: createEmptyBwaMonth(),
        4: createEmptyBwaMonth()
    };
    let jahresSum = createEmptyBwaMonth();
    jahresSum.offeneForderungen = bwaData.offeneForderungen;
    jahresSum.offeneVerbindlichkeiten = bwaData.offeneVerbindlichkeiten;

    // Ausgabe in der Reihenfolge: Monat 1, Monat 2, Monat 3, Quartal 1, Monat 4, ...
    for (let m = 1; m <= 12; m++) {
        let moObj = bwaData.monats[m] || createEmptyBwaMonth();
        bwaSheet.appendRow(toRowData(moObj, `Monat ${m}`));

        let q = Math.ceil(m / 3);
        // Für Quartalsaggregation: Summiere alle Felder außer liquiditaet
        for (let key in moObj) {
            if (key !== "liquiditaet" && typeof moObj[key] === "number") {
                quartalsSum[q][key] += moObj[key];
            }
        }
        // Liquidität im Quartal: Verwende den Wert des letzten Monats im Quartal
        if (m % 3 === 0) {
            quartalsSum[q]["liquiditaet"] = moObj.liquiditaet;
        }
        // Für Jahresaggregation (außer liquiditaet)
        for (let key in moObj) {
            if (key !== "liquiditaet" && typeof moObj[key] === "number") {
                jahresSum[key] += moObj[key];
            }
        }
    }
    // Für die Jahreszeile: Liquidität = Wert aus Monat 12
    jahresSum.liquiditaet = bwaData.monats[12] ? bwaData.monats[12].liquiditaet : 0;

    // Füge Quartalszeilen an der richtigen Stelle ein:
    // Zuerst Quartal 1 (nach Monat 3), dann Quartal 2 (nach Monat 6), usw.
    // Wir erstellen eine Hilfsvariable für das Endergebnis:
    let finalOutput = [];
    // Kopfzeile bleibt bereits
    // Quartale hinzufügen:
    for (let q = 1; q <= 4; q++) {
        finalOutput.push(toRowData(quartalsSum[q], `Quartal ${q}`));
    }
    // Jetzt: Hänge die Quartalszeilen nach den entsprechenden Monatsgruppen an.
    // Hier bauen wir das finale Ausgabe-Array, indem wir die Zeilen in folgender Reihenfolge zusammenführen:
    // (Monat1, Monat2, Monat3, Quartal1, Monat4, Monat5, Monat6, Quartal2, …, Monat10, Monat11, Monat12, Quartal4, Gesamtjahr)
    let outputRows = [];
    for (let q = 1; q <= 4; q++) {
        let startMonth = (q - 1) * 3 + 1;
        for (let m = startMonth; m < startMonth + 3; m++) {
            outputRows.push(toRowData(bwaData.monats[m] || createEmptyBwaMonth(), `Monat ${m}`));
        }
        outputRows.push(toRowData(quartalsSum[q], `Quartal ${q}`));
    }
    // Füge die Jahreszeile hinzu:
    outputRows.push(toRowData(jahresSum, "Gesamtjahr"));

    // Leere das Blatt und füge Kopfzeile und alle Zeilen ein:
    bwaSheet.clearContents();
    bwaSheet.appendRow([
        "Zeitraum",
        "Dienstleistungen", "Produkte", "Sonst. Einnahmen", "Gesamt-Einnahmen",
        "Betriebskosten", "Marketing", "Reisen", "Einkauf", "Personal", "Gesamt-Ausgaben",
        "Offene Forderungen", "Offene Verbindlichkeiten",
        "Rohertrag", "Betriebsergebnis", "Ergebnis vor Steuern", "Ergebnis nach Steuern",
        "Liquidität"
    ]);
    outputRows.forEach(row => {
        bwaSheet.appendRow(row);
    });

    // Formatierung
    const lastRow = bwaSheet.getLastRow();
    if (lastRow > 1) {
        bwaSheet.getRange(`B2:R${lastRow}`).setNumberFormat("#,##0.00€");
        bwaSheet.autoResizeColumns(1, bwaSheet.getLastColumn());
    }

    if (fehlendeKategorien.length > 0) {
        SpreadsheetApp.getUi().alert(
            "Achtung! Unbekannte Kategorien gefunden:\n" +
            fehlendeKategorien.join("\n") +
            "\n\n→ Es wurde 'sonstigeEinnahmen' bzw. 'betriebskosten' verwendet."
        );
    }

    SpreadsheetApp.getUi().alert("BWA-Berechnung abgeschlossen und aktualisiert.");
}
