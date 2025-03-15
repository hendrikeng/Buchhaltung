// src/modules/refreshModule/matchingHandler.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import globalCache from '../../utils/cacheUtils.js';
import formattingHandler from './formattingHandler.js';
import syncHandler from './syncHandler.js';

/**
 * Führt das Matching von Bankbewegungen mit Rechnungen durch
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Sheet} sheet - Das Bankbewegungen-Sheet
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {number} numDataRows - Anzahl der Datenzeilen
 * @param {number} lastRow - Letzte Zeile
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} columnLetters - Buchstaben für die Spalten
 * @param {Object} config - Die Konfiguration
 */
function performBankReferenceMatching(ss, sheet, firstDataRow, numDataRows, lastRow, columns, columnLetters, config) {
    // Referenzdaten laden für alle relevanten Sheets
    const sheetData = loadReferenceData(ss, config);

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
 * Lädt Referenzdaten aus allen relevanten Sheets
 */
function loadReferenceData(ss, config) {
    const result = {};

    // Sheets und ihre Konfigurationen
    const sheets = {
        einnahmen: { name: 'Einnahmen', cols: config.einnahmen.columns },
        ausgaben: { name: 'Ausgaben', cols: config.ausgaben.columns },
        eigenbelege: { name: 'Eigenbelege', cols: config.eigenbelege.columns },
        gesellschafterkonto: { name: 'Gesellschafterkonto', cols: config.gesellschafterkonto.columns },
        holdingTransfers: { name: 'Holding Transfers', cols: config.holdingTransfers.columns },
    };

    // Für jedes Sheet Referenzdaten laden
    Object.entries(sheets).forEach(([key, config]) => {
        const sheet = ss.getSheetByName(config.name);
        if (sheet && sheet.getLastRow() > 1) {
            result[key] = loadSheetReferenceData(sheet, config.cols, key);
        } else {
            result[key] = {};
        }
    });

    return result;
}

/**
 * Lädt Referenzdaten aus einem einzelnen Sheet
 */
function loadSheetReferenceData(sheet, columns, sheetType) {
    // Cache prüfen
    if (globalCache.has('references', sheetType)) {
        return globalCache.get('references', sheetType);
    }

    const map = {};
    const numRows = sheet.getLastRow() - 1;

    if (numRows <= 0) return map;

    // Die benötigten Spalten bestimmen
    const columnsToLoad = Math.max(
        columns.rechnungsnummer,
        columns.nettobetrag,
        columns.mwstSatz || 0,
        columns.bezahlt || 0,
    );

    // Die relevanten Spalten laden
    const data = sheet.getRange(2, 1, numRows, columnsToLoad).getValues();

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const ref = row[columns.rechnungsnummer - 1];

        if (stringUtils.isEmpty(ref)) continue;

        // Entferne "G-" Prefix für den Key, falls vorhanden (für Gutschriften)
        let key = ref.toString().trim();
        const isGutschrift = key.startsWith('G-');

        if (isGutschrift) {
            key = key.substring(2);
        }

        const normalizedKey = stringUtils.normalizeText(key);

        // Werte extrahieren
        let betrag = 0;
        if (!stringUtils.isEmpty(row[columns.nettobetrag - 1])) {
            betrag = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
            betrag = Math.abs(betrag);
        }

        let mwstRate = 0;
        if (columns.mwstSatz && !stringUtils.isEmpty(row[columns.mwstSatz - 1])) {
            mwstRate = numberUtils.parseMwstRate(row[columns.mwstSatz - 1]);
        }

        let bezahlt = 0;
        if (columns.bezahlt && !stringUtils.isEmpty(row[columns.bezahlt - 1])) {
            bezahlt = numberUtils.parseCurrency(row[columns.bezahlt - 1]);
            bezahlt = Math.abs(bezahlt);
        }

        // Bruttobetrag berechnen
        const brutto = betrag * (1 + mwstRate/100);

        map[key] = {
            ref: ref.toString().trim(),
            normalizedRef: normalizedKey,
            row: i + 2, // 1-basiert + Überschrift
            betrag: betrag,
            mwstRate: mwstRate,
            brutto: brutto,
            bezahlt: bezahlt,
            offen: numberUtils.round(brutto - bezahlt, 2),
            isGutschrift: isGutschrift,
        };

        // Normalisierter Key als Alternative
        if (normalizedKey !== key && !map[normalizedKey]) {
            map[normalizedKey] = map[key];
        }
    }

    // Referenzdaten cachen
    globalCache.set('references', sheetType, map);

    return map;
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
        config.holdingTransfers.kontoMapping,
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
    const fallbackKontoSoll = allowedKontoSoll.size > 0 ? Array.from(allowedKontoSoll)[0] : '';
    const fallbackKontoHaben = allowedGegenkonto.size > 0 ? Array.from(allowedGegenkonto)[0] : '';

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
        const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : '';

        // Match-Informationen
        let matchFound = false;

        // Matching-Logik basierend auf Transaktionstyp
        if (!stringUtils.isEmpty(refNumber)) {
            if (tranType === 'Einnahme') {
                matchFound = processDocumentTypeMatching(sheetData.einnahmen, refNumber, betragValue, row,
                    columns, 'einnahme', config.einnahmen.columns, config.einnahmen.kontoMapping,
                    matchResults, kontoSollResults, kontoHabenResults,
                    bankZuordnungen, category, config);
            } else if (tranType === 'Ausgabe') {
                const docTypes = [
                    { data: sheetData.ausgaben, type: 'ausgabe', cols: config.ausgaben.columns, mapping: config.ausgaben.kontoMapping },
                    { data: sheetData.eigenbelege, type: 'eigenbeleg', cols: config.eigenbelege.columns, mapping: config.eigenbelege.kontoMapping },
                    { data: sheetData.gesellschafterkonto, type: 'gesellschafterkonto', cols: config.gesellschafterkonto.columns, mapping: config.gesellschafterkonto.kontoMapping },
                    { data: sheetData.holdingTransfers, type: 'holdingtransfer', cols: config.holdingTransfers.columns, mapping: config.holdingTransfers.kontoMapping },
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
                        columns, 'gutschrift', config.einnahmen.columns, config.einnahmen.kontoMapping,
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

// Weitere Hilfsfunktionen
function isEndSaldoRow(rowIndex, lastRow, row, columns) {
    const label = row[columns.buchungstext - 1] ? row[columns.buchungstext - 1].toString().trim().toLowerCase() : '';
    return rowIndex === lastRow && label === 'endsaldo';
}

function addEmptyResults(matchResults, kontoSollResults, kontoHabenResults) {
    matchResults.push(['']);
    kontoSollResults.push(['']);
    kontoHabenResults.push(['']);
}

function processDocumentTypeMatching(docData, refNumber, betragValue, row, columns, docType, docCols, kontoMapping,
    matchResults, kontoSollResults, kontoHabenResults, bankZuordnungen,
    category, config, isGutschrift = false) {
    const matchResult = findMatch(refNumber, docData, isGutschrift ? null : betragValue);
    if (!matchResult) return false;

    // Match-Info erstellen basierend auf Dokumenttyp
    const matchInfo = createMatchInfo(matchResult, docType, betragValue);

    // Konten aus dem Mapping setzen
    let kontoSoll = '', kontoHaben = '';
    if (category && kontoMapping[category]) {
        const mapping = kontoMapping[category];
        kontoSoll = mapping.soll || '';
        kontoHaben = mapping.gegen || '';
    }

    // Ergebnisse zu den Arrays hinzufügen
    matchResults.push([matchInfo]);
    kontoSollResults.push([kontoSoll]);
    kontoHabenResults.push([kontoHaben]);

    // Für spätere Markierung merken
    const key = `${docType}#${matchResult.row}`;
    bankZuordnungen[key] = {
        typ: docType === 'gutschrift' ? 'einnahme' : docType,
        row: matchResult.row,
        bankDatum: row[columns.datum - 1],
        matchInfo: matchInfo,
        transTyp: docType === 'gutschrift' ? 'Gutschrift' : row[columns.transaktionstyp - 1],
    };

    return true;
}

function findMatch(reference, refMap, betrag = null) {
    // Keine Daten
    if (!reference || !refMap) return null;

    // Normalisierte Suche vorbereiten
    const normalizedRef = stringUtils.normalizeText(reference);

    // Priorität der Matching-Strategien:
    // 1. Exakte Übereinstimmung
    // 2. Normalisierte Übereinstimmung
    // 3. Enthält-Beziehung (in beide Richtungen)

    // 1. Exakte Übereinstimmung
    if (refMap[reference]) {
        return evaluateMatch(refMap[reference], betrag);
    }

    // 2. Normalisierte Übereinstimmung
    if (normalizedRef && refMap[normalizedRef]) {
        return evaluateMatch(refMap[normalizedRef], betrag);
    }

    // 3. Teilweise Übereinstimmung (mit Cache für Performance)
    const candidateKeys = Object.keys(refMap);

    // Zuerst prüfen wir, ob die Referenz in einem Schlüssel enthalten ist
    for (const key of candidateKeys) {
        if (key.includes(reference) || reference.includes(key)) {
            return evaluateMatch(refMap[key], betrag);
        }
    }

    // Falls keine exakten Treffer, probieren wir es mit normalisierten Werten
    for (const key of candidateKeys) {
        const normalizedKey = stringUtils.normalizeText(key);
        if (normalizedKey.includes(normalizedRef) || normalizedRef.includes(normalizedKey)) {
            return evaluateMatch(refMap[key], betrag);
        }
    }

    return null;
}

function evaluateMatch(match, betrag = null) {
    if (!match) return null;

    // Behalte die ursprüngliche Übereinstimmung
    const result = { ...match };

    // Wenn kein Betrag zum Vergleich angegeben wurde
    if (betrag === null) {
        result.matchType = 'Referenz-Match';
        return result;
    }

    // Hole die Beträge aus der Übereinstimmung
    const matchBrutto = Math.abs(match.brutto);
    const matchBezahlt = Math.abs(match.bezahlt);

    // Toleranzwert für Betragsabweichungen (2 Cent)
    const tolerance = 0.02;

    // Fall 1: Betrag entspricht genau dem Bruttobetrag (Vollständige Zahlung)
    if (numberUtils.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
        result.matchType = 'Vollständige Zahlung';
        return result;
    }

    // Fall 2: Position ist bereits vollständig bezahlt
    if (numberUtils.isApproximatelyEqual(matchBezahlt, matchBrutto, tolerance) && matchBezahlt > 0) {
        result.matchType = 'Vollständige Zahlung';
        return result;
    }

    // Fall 3: Teilzahlung
    if (betrag < matchBrutto && (matchBrutto - betrag) > (matchBrutto * 0.1)) {
        result.matchType = 'Teilzahlung';
        return result;
    }

    // Fall 4: Betrag größer, aber nahe dem Bruttobetrag
    if (betrag > matchBrutto && numberUtils.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
        result.matchType = 'Vollständige Zahlung';
        return result;
    }

    // Fall 5: Bei allen anderen Fällen
    result.matchType = 'Unsichere Zahlung';
    result.betragsDifferenz = numberUtils.round(Math.abs(betrag - matchBrutto), 2);
    return result;
}

function createMatchInfo(matchResult, docType, betragValue) {
    let matchStatus = '';

    if (matchResult.matchType) {
        if (matchResult.matchType === 'Unsichere Zahlung' && matchResult.betragsDifferenz) {
            matchStatus = ` (${matchResult.matchType}, Diff: ${matchResult.betragsDifferenz}€)`;
        } else {
            matchStatus = ` (${matchResult.matchType})`;
        }
    }

    // Spezialbehandlung für Gutschriften
    if (docType === 'gutschrift') {
        const gutschriftBetrag = Math.abs(matchResult.brutto);

        if (numberUtils.isApproximatelyEqual(betragValue, gutschriftBetrag, 0.01)) {
            matchStatus = ' (Vollständige Gutschrift)';
        } else if (betragValue < gutschriftBetrag) {
            matchStatus = ' (Teilgutschrift)';
        } else {
            matchStatus = ' (Ungewöhnliche Gutschrift)';
        }

        return `Gutschrift zu Einnahme #${matchResult.row}${matchStatus}`;
    }

    // Standardbehandlung für andere Dokumenttypen
    const docTypeName = {
        'einnahme': 'Einnahme',
        'ausgabe': 'Ausgabe',
        'eigenbeleg': 'Eigenbeleg',
        'gesellschafterkonto': 'Gesellschafterkonto',
        'holdingtransfer': 'Holding Transfer',
    }[docType] || docType;

    return `${docTypeName} #${matchResult.row}${matchStatus}`;
}

function setDefaultAccounts(category, tranType, config, fallbackKontoSoll, fallbackKontoHaben,
    matchResults, kontoSollResults, kontoHabenResults) {
    let kontoSoll = fallbackKontoSoll;
    let kontoHaben = fallbackKontoHaben;

    if (category) {
        // Den richtigen Mapping-Typ basierend auf der Transaktionsart auswählen
        const mappingSource = tranType === 'Einnahme' ?
            config.einnahmen.kontoMapping :
            config.ausgaben.kontoMapping;

        // Mapping für die Kategorie finden
        if (mappingSource && mappingSource[category]) {
            const mapping = mappingSource[category];
            kontoSoll = mapping.soll || fallbackKontoSoll;
            kontoHaben = mapping.gegen || fallbackKontoHaben;
        }
    }

    matchResults.push(['']);
    kontoSollResults.push([kontoSoll]);
    kontoHabenResults.push([kontoHaben]);
}

function performBatchUpdates(sheet, firstDataRow, numDataRows, columns,
    matchResults, kontoSollResults, kontoHabenResults) {
    sheet.getRange(firstDataRow, columns.matchInfo, numDataRows, 1).setValues(matchResults);
    sheet.getRange(firstDataRow, columns.kontoSoll, numDataRows, 1).setValues(kontoSollResults);
    sheet.getRange(firstDataRow, columns.kontoHaben, numDataRows, 1).setValues(kontoHabenResults);
}

export default {
    performBankReferenceMatching,
};