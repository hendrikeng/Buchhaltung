// src/modules/refreshModule/matchingHandler.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import globalCache from '../../utils/cacheUtils.js';
import matchingUtils from './matchingUtils.js';
import formattingHandler from './formattingHandler.js';
import syncHandler from './syncHandler.js';

/**
 * Hauptfunktion: Führt das Matching von Bankbewegungen mit Rechnungen durch
 */
function performBankReferenceMatching(ss, sheet, firstDataRow, numDataRows, lastRow, columns, columnLetters, config) {
    // Referenzdaten laden für alle relevanten Sheets
    const sheetData = matchingUtils.loadReferenceData(ss, config);

    // Bankbewegungen Daten für Verarbeitung holen
    const bankData = sheet.getRange(firstDataRow, 1, numDataRows, columns.matchInfo).getDisplayValues();

    // Vorbereitung für Batch-Updates
    const matchResults = [];
    const kontoSollResults = [];
    const kontoHabenResults = [];
    const bankZuordnungen = {};

    // Kontoinformationen sammeln
    const { allowedKontoSoll, allowedGegenkonto } = collectAccountInfo(config);

    // Durchlaufe jede Bankbewegung und suche nach Übereinstimmungen
    processBankEntries(bankData, firstDataRow, lastRow, columns, sheetData,
        matchResults, kontoSollResults, kontoHabenResults,
        bankZuordnungen, allowedKontoSoll, allowedGegenkonto, config);

    // Batch-Updates ausführen für bessere Performance
    performBatchUpdates(sheet, firstDataRow, numDataRows, columns,
        matchResults, kontoSollResults, kontoHabenResults);

    // Formatiere die gesamten Zeilen basierend auf dem Match-Typ
    formattingHandler.formatMatchedRows(sheet, firstDataRow, matchResults, columns);

    // Bedingte Formatierung für Match-Spalte
    formattingHandler.setMatchColumnFormatting(sheet, columnLetters.matchInfo);

    // Setze farbliche Markierung in den anderen Sheets basierend auf Zahlungsstatus
    syncHandler.markPaidInvoices(ss, bankZuordnungen, config);
}

/**
 * Sammelt Kontoinformationen aus der Konfiguration
 */
function collectAccountInfo(config) {
    const allowedKontoSoll = new Set();
    const allowedGegenkonto = new Set();

    [
        config.einnahmen.kontoMapping,
        config.ausgaben.kontoMapping,
        config.eigenbelege.kontoMapping,
        config.gesellschafterkonto.kontoMapping,
        config.holdingTransfers.kontoMapping
    ].forEach(mapping => {
        Object.values(mapping).forEach(m => {
            if (m.soll) allowedKontoSoll.add(m.soll);
            if (m.gegen) allowedGegenkonto.add(m.gegen);
        });
    });

    return { allowedKontoSoll, allowedGegenkonto };
}

/**
 * Verarbeitet alle Bankeinträge und sucht nach Übereinstimmungen
 */
function processBankEntries(bankData, firstDataRow, lastRow, columns, sheetData,
                            matchResults, kontoSollResults, kontoHabenResults,
                            bankZuordnungen, allowedKontoSoll, allowedGegenkonto, config) {

    // Fallback-Konten
    const fallbackKontoSoll = allowedKontoSoll.size > 0 ? Array.from(allowedKontoSoll)[0] : "";
    const fallbackKontoHaben = allowedGegenkonto.size > 0 ? Array.from(allowedGegenkonto)[0] : "";

    // Durchlaufe jede Bankbewegung
    for (let i = 0; i < bankData.length; i++) {
        const rowIndex = i + firstDataRow;
        const row = bankData[i];

        // Prüfe, ob es sich um die Endsaldo-Zeile handelt
        if (isEndSaldoRow(rowIndex, lastRow, row, columns)) {
            addEmptyResults(matchResults, kontoSollResults, kontoHabenResults);
            continue;
        }

        const tranType = row[columns.transaktionstyp - 1]; // Einnahme/Ausgabe
        const refNumber = row[columns.referenz - 1];       // Referenznummer
        const betragValue = Math.abs(numberUtils.parseCurrency(row[columns.betrag - 1]));
        const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : "";

        // Match-Informationen
        let matchFound = false;

        // Matching-Logik basierend auf Transaktionstyp
        if (!stringUtils.isEmpty(refNumber)) {
            if (tranType === "Einnahme") {
                matchFound = processDocumentTypeMatching(sheetData.einnahmen, refNumber, betragValue, row,
                    columns, "einnahme", config.einnahmen.columns, config.einnahmen.kontoMapping,
                    matchResults, kontoSollResults, kontoHabenResults,
                    bankZuordnungen, category, config);
            } else if (tranType === "Ausgabe") {
                const docTypes = [
                    { data: sheetData.ausgaben, type: "ausgabe", cols: config.ausgaben.columns, mapping: config.ausgaben.kontoMapping },
                    { data: sheetData.eigenbelege, type: "eigenbeleg", cols: config.eigenbelege.columns, mapping: config.eigenbelege.kontoMapping },
                    { data: sheetData.gesellschafterkonto, type: "gesellschafterkonto", cols: config.gesellschafterkonto.columns, mapping: config.gesellschafterkonto.kontoMapping },
                    { data: sheetData.holdingTransfers, type: "holdingtransfer", cols: config.holdingTransfers.columns, mapping: config.holdingTransfers.kontoMapping }
                ];

                // Versuche alle Dokumenttypen der Reihe nach
                for (const docType of docTypes) {
                    matchFound = processDocumentTypeMatching(docType.data, refNumber, betragValue, row,
                        columns, docType.type, docType.cols, docType.mapping,
                        matchResults, kontoSollResults, kontoHabenResults,
                        bankZuordnungen, category, config);
                    if (matchFound) break;
                }

                // Falls keine Übereinstimmung, noch nach Gutschriften in Einnahmen suchen
                if (!matchFound) {
                    matchFound = processDocumentTypeMatching(sheetData.einnahmen, refNumber, betragValue, row,
                        columns, "gutschrift", config.einnahmen.columns, config.einnahmen.kontoMapping,
                        matchResults, kontoSollResults, kontoHabenResults,
                        bankZuordnungen, category, config, true);
                }
            }
        }

        // Wenn keine Übereinstimmung gefunden, default Konten basierend auf Kategorie setzen
        if (!matchFound) {
            setDefaultAccounts(category, tranType, config, fallbackKontoSoll, fallbackKontoHaben,
                matchResults, kontoSollResults, kontoHabenResults);
        }
    }
}

/**
 * Prüft, ob es sich um die Endsaldo-Zeile handelt
 */
function isEndSaldoRow(rowIndex, lastRow, row, columns) {
    const label = row[columns.buchungstext - 1] ? row[columns.buchungstext - 1].toString().trim().toLowerCase() : "";
    return rowIndex === lastRow && label === "endsaldo";
}

/**
 * Fügt leere Ergebnisse zu den Arrays hinzu
 */
function addEmptyResults(matchResults, kontoSollResults, kontoHabenResults) {
    matchResults.push([""]);
    kontoSollResults.push([""]);
    kontoHabenResults.push([""]);
}

/**
 * Generalisierte Funktion für das Matching eines Dokumenttyps
 */
function processDocumentTypeMatching(docData, refNumber, betragValue, row, columns, docType, docCols, kontoMapping,
                                     matchResults, kontoSollResults, kontoHabenResults, bankZuordnungen,
                                     category, config, isGutschrift = false) {

    const matchResult = matchingUtils.findMatch(refNumber, docData, isGutschrift ? null : betragValue);

    if (!matchResult) return false;

    // Match-Info erstellen basierend auf Dokumenttyp
    let matchInfo = createMatchInfo(matchResult, docType, betragValue);

    // Konten aus dem Mapping setzen
    let kontoSoll = "", kontoHaben = "";
    if (category && kontoMapping[category]) {
        const mapping = kontoMapping[category];
        kontoSoll = mapping.soll || "";
        kontoHaben = mapping.gegen || "";
    }

    // Ergebnisse zu den Arrays hinzufügen
    matchResults.push([matchInfo]);
    kontoSollResults.push([kontoSoll]);
    kontoHabenResults.push([kontoHaben]);

    // Für spätere Markierung merken
    const key = `${docType}#${matchResult.row}`;
    bankZuordnungen[key] = {
        typ: docType === "gutschrift" ? "einnahme" : docType,
        row: matchResult.row,
        bankDatum: row[columns.datum - 1],
        matchInfo: matchInfo,
        transTyp: docType === "gutschrift" ? "Gutschrift" : row[columns.transaktionstyp - 1]
    };

    return true;
}

/**
 * Erzeugt die Match-Info basierend auf dem Dokumenttyp
 */
function createMatchInfo(matchResult, docType, betragValue) {
    let matchStatus = "";

    if (matchResult.matchType) {
        if (matchResult.matchType === "Unsichere Zahlung" && matchResult.betragsDifferenz) {
            matchStatus = ` (${matchResult.matchType}, Diff: ${matchResult.betragsDifferenz}€)`;
        } else {
            matchStatus = ` (${matchResult.matchType})`;
        }
    }

    // Spezialbehandlung für Gutschriften
    if (docType === "gutschrift") {
        const gutschriftBetrag = Math.abs(matchResult.brutto);

        if (numberUtils.isApproximatelyEqual(betragValue, gutschriftBetrag, 0.01)) {
            matchStatus = " (Vollständige Gutschrift)";
        } else if (betragValue < gutschriftBetrag) {
            matchStatus = " (Teilgutschrift)";
        } else {
            matchStatus = " (Ungewöhnliche Gutschrift)";
        }

        return `Gutschrift zu Einnahme #${matchResult.row}${matchStatus}`;
    }

    // Standardbehandlung für andere Dokumenttypen
    const docTypeName = {
        "einnahme": "Einnahme",
        "ausgabe": "Ausgabe",
        "eigenbeleg": "Eigenbeleg",
        "gesellschafterkonto": "Gesellschafterkonto",
        "holdingtransfer": "Holding Transfer"
    }[docType] || docType;

    return `${docTypeName} #${matchResult.row}${matchStatus}`;
}

/**
 * Setzt Default-Konten basierend auf Kategorie und Transaktionstyp
 */
function setDefaultAccounts(category, tranType, config, fallbackKontoSoll, fallbackKontoHaben,
                            matchResults, kontoSollResults, kontoHabenResults) {

    let kontoSoll = fallbackKontoSoll;
    let kontoHaben = fallbackKontoHaben;

    if (category) {
        // Den richtigen Mapping-Typ basierend auf der Transaktionsart auswählen
        const mappingSource = tranType === "Einnahme" ?
            config.einnahmen.kontoMapping :
            config.ausgaben.kontoMapping;

        // Mapping für die Kategorie finden
        if (mappingSource && mappingSource[category]) {
            const mapping = mappingSource[category];
            kontoSoll = mapping.soll || fallbackKontoSoll;
            kontoHaben = mapping.gegen || fallbackKontoHaben;
        }
    }

    matchResults.push([""]);
    kontoSollResults.push([kontoSoll]);
    kontoHabenResults.push([kontoHaben]);
}

/**
 * Führt die Batch-Updates für die generierten Daten aus
 */
function performBatchUpdates(sheet, firstDataRow, numDataRows, columns,
                             matchResults, kontoSollResults, kontoHabenResults) {

    sheet.getRange(firstDataRow, columns.matchInfo, numDataRows, 1).setValues(matchResults);
    sheet.getRange(firstDataRow, columns.kontoSoll, numDataRows, 1).setValues(kontoSollResults);
    sheet.getRange(firstDataRow, columns.kontoHaben, numDataRows, 1).setValues(kontoHabenResults);
}

export default {
    performBankReferenceMatching
};