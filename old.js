
/**
 * Konvertiert einen MwSt.-Wert in eine Zahl (Prozent).
 * Beispiel: "0,19" wird zu 19, "19%" wird zu 19.
 */
const parseMwstRate = value => {
    let rate = parseFloat(value.toString().replace("%", "").replace(",", "."));
    if (rate < 1) rate *= 100;
    return isNaN(rate) ? 19 : rate;
};


/**
 * Wandelt einen Wert in ein Date-Objekt um.
 */
const parseDate = value => {
    const d = (typeof value === "string")
        ? new Date(value)
        : (value instanceof Date ? value : null);
    return (d && !isNaN(d.getTime())) ? d : null;
};

/**
 * Wandelt einen Währungs-/Zahlenstring in einen Float um.
 */
const parseCurrency = value =>
    parseFloat(value.toString().replace(/[^\d,.-]/g, "").replace(",", ".")) || 0;

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
