import config from "./config.js";

const calculateBilanz = () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Ermittelt Bankguthaben aus "Bankbewegungen" (Endsaldo)
    let bankGuthaben = "";
    const bankSheet = ss.getSheetByName("Bankbewegungen");
    if (bankSheet) {
        const lastRow = bankSheet.getLastRow();
        if (lastRow >= 1) {
            const label = bankSheet.getRange(lastRow, 2).getValue().toString().toLowerCase();
            if (label === "endsaldo") {
                bankGuthaben = bankSheet.getRange(lastRow, 4).getValue();
            }
        }
    }

    // Jahresüberschuss aus "BWA" (Letzte Zeile, sofern dort "Jahresüberschuss" vorkommt)
    let jahresUeberschuss = "";
    const bwaSheet = ss.getSheetByName("BWA");
    if (bwaSheet) {
        const data = bwaSheet.getDataRange().getValues();
        const lastRowData = data[data.length - 1];
        if (lastRowData[0].toString().toLowerCase().includes("jahresüberschuss")) {
            jahresUeberschuss = lastRowData[lastRowData.length - 1];
        }
    }

    // Aufbau der Aktiva (2 Spalten: Bezeichnung, Wert)
    const aktiva = [
        ["Aktiva (Vermögenswerte)", ""],
        ["1.1 Anlagevermögen", ""],
        ["Sachanlagen", ""],
        ["Immaterielle Vermögenswerte", ""],
        ["Finanzanlagen", ""],
        ["Zwischensumme Anlagevermögen", "=SUM(B3:B5)"],
        ["", ""],
        ["1.2 Umlaufvermögen", ""],
        ["Bankguthaben", bankGuthaben],
        ["Kasse", ""],
        ["Forderungen aus L&L", ""],
        ["Vorräte", ""],
        ["Zwischensumme Umlaufvermögen", "=SUM(B9:B12)"],
        ["", ""],
        ["Gesamt Aktiva", "=B6+B13"]
    ];

    // Aufbau der Passiva (2 Spalten: Bezeichnung, Wert)
    const passiva = [
        ["Passiva (Finanzierung & Schulden)", ""],
        ["2.1 Eigenkapital", ""],
        ["Stammkapital", config.tax.stammkapital],
        ["Gewinn-/Verlustvortrag", ""],
        ["Jahresüberschuss/-fehlbetrag", jahresUeberschuss],
        ["Zwischensumme Eigenkapital", "=SUM(F3:F5)"],
        ["", ""],
        ["2.2 Verbindlichkeiten", ""],
        ["Bankdarlehen", ""],
        ["Gesellschafterdarlehen", ""],
        ["Verbindlichkeiten aus L&L", ""],
        ["Steuerrückstellungen", ""],
        ["Zwischensumme Verbindlichkeiten", "=SUM(F8:F11)"],
        ["", ""],
        ["Gesamt Passiva", "=F6+F13"]
    ];

    // Erstelle oder leere das Blatt "Bilanz"
    let bilanzSheet = ss.getSheetByName("Bilanz");
    if (!bilanzSheet) {
        bilanzSheet = ss.insertSheet("Bilanz");
    } else {
        bilanzSheet.clearContents();
    }

    // Schreibe Aktiva ab Zelle A1 und Passiva ab Zelle E1 (2 Spalten jeweils)
    bilanzSheet.getRange(1, 1, aktiva.length, 2).setValues(aktiva);
    bilanzSheet.getRange(1, 5, passiva.length, 2).setValues(passiva);

    // Spalten anpassen
    bilanzSheet.autoResizeColumns(1, 6);
    SpreadsheetApp.getUi().alert("Bilanz wurde erstellt!");
};

export default calculateBilanz;