// src/modules/bankReconciliationModule/matchingHandler.js
import stringUtils from '../../utils/stringUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import globalCache from '../../utils/cacheUtils.js';
import dateUtils from '../../utils/dateUtils.js';

/**
 * Führt das Matching von Bankbewegungen mit Rechnungen durch
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Sheet} sheet - Das Bankbewegungen-Sheet
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {number} numDataRows - Anzahl der Datenzeilen
 * @param {number} lastRow - Letzte Zeile
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Ergebnis des Matchings mit matchInfo und bankZuordnungen
 */
function performBankReferenceMatching(ss, sheet, firstDataRow, numDataRows, lastRow, columns, config) {
    // Optimierung: Laden und Verarbeiten der Daten in einer Operation
    console.log('Starting bank reference matching...');

    // 1. Referenzdaten in einem Batch laden - dies vermeidet wiederholte Sheet-Zugriffe
    const sheetData = loadReferenceData(ss, config);

    // 2. Bankdaten in einem Batch laden
    const bankData = sheet.getRange(firstDataRow, 1, numDataRows, sheet.getLastColumn()).getValues();

    // 3. Matching vorbereiten
    const results = {
        matchInfo: [],
        kontoSollResults: [],
        kontoHabenResults: [],
        categoryResults: [],
        bankZuordnungen: {},
    };

    // 4. Kontoinformationen sammeln (einmalig)
    const accountInfo = collectAccountInfo(config);

    // 5. Matching in einem Durchgang durchführen
    processBankEntries(
        bankData,
        firstDataRow,
        lastRow,
        columns,
        sheetData,
        results,
        accountInfo,
        config,
    );

    // 6. Batch-Updates durchführen
    performBatchUpdates(
        sheet,
        firstDataRow,
        numDataRows,
        columns,
        results.matchInfo,
        results.kontoSollResults,
        results.kontoHabenResults,
        results.categoryResults,
    );

    return {
        matchInfo: results.matchInfo,
        bankZuordnungen: results.bankZuordnungen,
    };
}

/**
 * Lädt Referenzdaten aus allen relevanten Sheets in einem Batch
 * @param {Spreadsheet} ss - Das Spreadsheet
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Die geladenen Referenzdaten
 */
function loadReferenceData(ss, config) {
    const result = {};

    // Optimierung: Sheets und Konfigurationen in einem Array definieren
    const sheetConfig = [
        { key: 'einnahmen', name: 'Einnahmen', cols: config.einnahmen.columns },
        { key: 'ausgaben', name: 'Ausgaben', cols: config.ausgaben.columns },
        { key: 'eigenbelege', name: 'Eigenbelege', cols: config.eigenbelege.columns },
        { key: 'gesellschafterkonto', name: 'Gesellschafterkonto', cols: config.gesellschafterkonto.columns },
        { key: 'holdingTransfers', name: 'Holding Transfers', cols: config.holdingTransfers.columns },
    ];

    // Für jedes Sheet Referenzdaten laden - parallelisiert
    sheetConfig.forEach(cfg => {
        // Prüfe zuerst den Cache
        if (globalCache.has('references', cfg.key)) {
            result[cfg.key] = globalCache.get('references', cfg.key);
            return;
        }

        const sheet = ss.getSheetByName(cfg.name);
        if (sheet && sheet.getLastRow() > 1) {
            result[cfg.key] = loadSheetReferenceData(sheet, cfg.cols, cfg.key);
        } else {
            result[cfg.key] = {};
        }
    });

    return result;
}

/**
 * Lädt Referenzdaten aus einem einzelnen Sheet mit optimierter Verarbeitung
 * @param {Sheet} sheet - Das Sheet
 * @param {Object} columns - Die Spaltenkonfiguration
 * @param {string} sheetType - Der Sheet-Typ
 * @returns {Object} Die geladenen Referenzdaten
 */
function loadSheetReferenceData(sheet, columns, sheetType) {
    // Cache prüfen
    if (globalCache.has('references', sheetType)) {
        return globalCache.get('references', sheetType);
    }

    const map = {};
    const numRows = sheet.getLastRow() - 1;

    if (numRows <= 0) return map;

    // Optimierung: Nur benötigte Spalten laden
    let requiredCols = [];

    // Für Gesellschafterkonto und Holding Transfers spezielle Konfiguration
    if (sheetType === 'gesellschafterkonto' || sheetType === 'holdingTransfers') {
        requiredCols = [
            columns.referenz,         // Statt rechnungsnummer
            columns.betrag,           // Betrag für Gesellschafterkonto/HoldingTransfers
            columns.kategorie,
            columns.buchungskonto,
            columns.bezahlt || 0,      // Falls vorhanden
        ].filter(col => col !== undefined);
    } else {
        requiredCols = [
            columns.rechnungsnummer,
            columns.nettobetrag,
            columns.mwstSatz,
            columns.bezahlt,
            columns.kategorie,
            columns.buchungskonto,
        ].filter(col => col !== undefined);
    }

    const maxCol = Math.max(...requiredCols);

    // Daten in einem Batch laden
    const data = sheet.getRange(2, 1, numRows, maxCol).getValues();

    // Effizienter Aufbau der Referenzdaten
    for (let i = 0; i < data.length; i++) {
        const row = data[i];

        // Referenz basierend auf dem Sheettyp extrahieren
        let ref;
        if (sheetType === 'gesellschafterkonto' || sheetType === 'holdingTransfers') {
            ref = row[columns.referenz - 1]; // Für Gesellschafterkonto und Holding Transfers verwende referenz
        } else {
            ref = row[columns.rechnungsnummer - 1]; // Für andere Sheets verwende rechnungsnummer
        }

        if (stringUtils.isEmpty(ref)) continue;

        // Key immer ohne G-Prefix
        const key = String(ref).trim();

        // Werte vorab extrahieren und berechnen
        let betrag;
        if (sheetType === 'gesellschafterkonto' || sheetType === 'holdingTransfers') {
            betrag = numberUtils.parseCurrency(row[columns.betrag - 1] || 0);
        } else {
            betrag = numberUtils.parseCurrency(row[columns.nettobetrag - 1] || 0);
        }

        const isGutschrift = betrag < 0;
        const normalizedKey = stringUtils.normalizeText(key);

        // MwSt-Rate je nach Sheettyp
        let mwstRate = 0;
        if (sheetType !== 'gesellschafterkonto' && columns.mwstSatz) {
            mwstRate = numberUtils.parseMwstRate(row[columns.mwstSatz - 1] || 0);
        }

        // Bezahlt-Betrag
        let bezahlt = 0;
        if (columns.bezahlt) {
            bezahlt = Math.abs(numberUtils.parseCurrency(row[columns.bezahlt - 1] || 0));
        }

        const kategorie = columns.kategorie && row[columns.kategorie - 1] ?
            row[columns.kategorie - 1].toString().trim() : '';
        const buchungskonto = columns.buchungskonto && row[columns.buchungskonto - 1] ?
            row[columns.buchungskonto - 1].toString().trim() : '';

        // Bruttobetrag einmal berechnen (für Gesellschafterkonto ist Brutto = Betrag)
        let bruttoAbs;
        if (sheetType === 'gesellschafterkonto') {
            bruttoAbs = Math.abs(betrag);
        } else {
            bruttoAbs = Math.abs(betrag) * (1 + mwstRate/100);
        }

        // Optimiertes Datenformat mit allen benötigten Informationen
        map[key] = {
            ref: key,
            normalizedRef: normalizedKey,
            row: i + 2, // 1-basiert + Überschrift
            betrag: betrag,
            mwstRate: mwstRate,
            brutto: bruttoAbs * (isGutschrift ? -1 : 1), // Preserve sign for display
            bezahlt: bezahlt,
            offen: numberUtils.round(bruttoAbs - bezahlt, 2) * (isGutschrift ? -1 : 1), // Preserve sign
            isGutschrift: isGutschrift,
            kategorie: kategorie,
            buchungskonto: buchungskonto,
            originalData: {
                kategorie: kategorie,
                buchungskonto: buchungskonto,
                betrag: betrag, // Keep the original value with sign
                rowData: row, // Komplette Zeile speichern
                sheetType: sheetType, // Typ des Sheets speichern
            },
        };

        // Normalisierter Key als Alternative für besseres Matching
        if (normalizedKey !== key && !map[normalizedKey]) {
            map[normalizedKey] = map[key];
        }
    }

    // Referenzdaten cachen
    globalCache.set('references', sheetType, map);
    return map;
}

/**
 * Sammelt Kontoinformationen aus der Konfiguration (optimiert)
 * @param {Object} config - Die Konfiguration
 * @returns {Object} Gesammelte Kontoinformationen
 */
function collectAccountInfo(config) {
    const allowedKontoSoll = new Set();
    const allowedGegenkonto = new Set();

    // Process all sheet types
    const sheetTypes = ['einnahmen', 'ausgaben', 'eigenbelege', 'gesellschafterkonto', 'holdingTransfers'];

    sheetTypes.forEach(sheetType => {
        // Skip if not in config
        if (!config[sheetType] || !config[sheetType].categories) return;

        // Get all categories for this sheet type
        const categories = config[sheetType].categories;

        // Extract accounts from each category
        Object.keys(categories).forEach(categoryKey => {
            const category = categories[categoryKey];
            if (category && category.kontoMapping) {
                if (category.kontoMapping.soll) allowedKontoSoll.add(category.kontoMapping.soll);
                if (category.kontoMapping.gegen) allowedGegenkonto.add(category.kontoMapping.gegen);
            }
        });
    });

    // Fallback accounts
    const fallbackKontoSoll = allowedKontoSoll.size > 0 ? Array.from(allowedKontoSoll)[0] : '';
    const fallbackKontoHaben = allowedGegenkonto.size > 0 ? Array.from(allowedGegenkonto)[0] : '';

    return {
        allowedKontoSoll,
        allowedGegenkonto,
        fallbackKontoSoll,
        fallbackKontoHaben,
    };
}

/**
 * Verarbeitet Bankeinträge und sucht nach Übereinstimmungen
 * @param {Array} bankData - Die Bankdaten
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {number} lastRow - Letzte Zeile
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Object} sheetData - Referenzdaten
 * @param {Object} results - Ergebnisobjekt
 * @param {Object} accountInfo - Kontoinformationen
 * @param {Object} config - Die Konfiguration
 */
// src/modules/bankReconciliationModule/matchingHandler.js - modified processBankEntries function

function processBankEntries(bankData, firstDataRow, lastRow, columns, sheetData,
    results, accountInfo, config) {

    // Destrukturiere die accountInfo für bessere Lesbarkeit
    const { fallbackKontoSoll, fallbackKontoHaben } = accountInfo;

    // Durchlaufe jede Bankbewegung einmal
    for (let i = 0; i < bankData.length; i++) {
        const rowIndex = i + firstDataRow;
        const row = bankData[i];

        // Prüfe, ob es sich um die Endsaldo-Zeile handelt
        if (isEndSaldoRow(rowIndex, lastRow, row, columns)) {
            addEmptyResults(results);
            continue;
        }

        // Extrahiere alle benötigten Informationen aus der Zeile
        const tranType = row[columns.transaktionstyp - 1]; // Einnahme/Ausgabe
        const refNumber = row[columns.referenz - 1];       // Referenznummer
        const betragValue = Math.abs(numberUtils.parseCurrency(row[columns.betrag - 1]));
        const betragOriginal = numberUtils.parseCurrency(row[columns.betrag - 1]); // Original with sign
        const category = row[columns.kategorie - 1] ? row[columns.kategorie - 1].toString().trim() : '';

        // Flag für gefundene Übereinstimmung
        let matchFound = false;

        // Nur Matching durchführen, wenn eine Referenznummer vorhanden ist
        if (!stringUtils.isEmpty(refNumber)) {
            // Einnahmen können auch Gesellschafterkonto-Einlagen oder Holding-Transfers sein
            if (tranType === 'Einnahme' && !matchFound) {
                // Zuerst auf Einnahmen prüfen
                matchFound = processDocumentTypeMatching(
                    sheetData.einnahmen, refNumber, betragValue, row,
                    columns, 'einnahme', config.einnahmen.columns, config.einnahmen.kontoMapping,
                    results, category, config, firstDataRow);

                // Wenn keine Einnahme gefunden, auf Gesellschafterkonto mit positivem Betrag prüfen
                if (!matchFound && sheetData.gesellschafterkonto) {
                    matchFound = processGesellschafterkontoMatching(
                        sheetData.gesellschafterkonto, refNumber, betragValue, row,
                        columns, true, // isPositive = true für Einlagen
                        config, results, firstDataRow);
                }

                // Wenn immer noch kein Match, auf Holding Transfers mit positivem Betrag prüfen
                if (!matchFound && sheetData.holdingTransfers) {
                    matchFound = processHoldingTransferMatching(
                        sheetData.holdingTransfers, refNumber, betragValue, row,
                        columns, true, // isPositive = true für Kapitalrückführung
                        config, results, firstDataRow);
                }
            }

            // Ausgaben können auch Gesellschafterkonto-Rückzahlungen oder negative Holding-Transfers sein
            else if (tranType === 'Ausgabe' && !matchFound) {
                // Für Ausgaben mit negativem Betrag zuerst auf Gutschriften prüfen
                if (betragOriginal < 0) {
                    matchFound = processDocumentTypeMatching(
                        sheetData.einnahmen, refNumber, betragValue, row,
                        columns, 'gutschrift', config.einnahmen.columns, config.einnahmen.kontoMapping,
                        results, category, config, firstDataRow, true);

                    // Wenn keine Gutschrift gefunden, auf Gesellschafterkonto mit negativem Betrag prüfen
                    if (!matchFound && sheetData.gesellschafterkonto) {
                        matchFound = processGesellschafterkontoMatching(
                            sheetData.gesellschafterkonto, refNumber, betragValue, row,
                            columns, false, // isPositive = false für Rückzahlungen
                            config, results, firstDataRow);
                    }

                    // Wenn immer noch kein Match, auf Holding Transfer mit negativem Betrag prüfen
                    if (!matchFound && sheetData.holdingTransfers) {
                        matchFound = processHoldingTransferMatching(
                            sheetData.holdingTransfers, refNumber, betragValue, row,
                            columns, false, // isPositive = false für Gewinnübertrag
                            config, results, firstDataRow);
                    }
                }

                // Wenn noch keine Übereinstimmung gefunden wurde, normale Ausgabentypen prüfen
                if (!matchFound) {
                    const docTypes = [
                        { data: sheetData.ausgaben, type: 'ausgabe', cols: config.ausgaben.columns, mapping: config.ausgaben.kontoMapping },
                        { data: sheetData.eigenbelege, type: 'eigenbeleg', cols: config.eigenbelege.columns, mapping: config.eigenbelege.kontoMapping },
                    ];

                    // Optimiert: Verwende some() für frühzeitigen Abbruch bei Treffer
                    matchFound = docTypes.some(docType =>
                        processDocumentTypeMatching(
                            docType.data, refNumber, betragValue, row,
                            columns, docType.type, docType.cols, docType.mapping,
                            results, category, config, firstDataRow),
                    );
                }
            }
        }

        // Wenn keine Übereinstimmung gefunden, default Konten basierend auf Kategorie setzen
        if (!matchFound) {
            setDefaultAccounts(
                category, tranType, config, fallbackKontoSoll, fallbackKontoHaben,
                results,
            );
        }
    }
}

/**
 * Prüft, ob es sich um die Endsaldo-Zeile handelt
 * @param {number} rowIndex - Zeilenindex
 * @param {number} lastRow - Letzte Zeile
 * @param {Array} row - Zeilendaten
 * @param {Object} columns - Spaltenkonfiguration
 * @returns {boolean} true wenn es die Endsaldo-Zeile ist
 */
function isEndSaldoRow(rowIndex, lastRow, row, columns) {
    const label = row[columns.buchungstext - 1] ? row[columns.buchungstext - 1].toString().trim().toLowerCase() : '';
    return rowIndex === lastRow && label === 'endsaldo';
}

/**
 * Fügt leere Ergebnisse zu den Ergebnis-Arrays hinzu
 * @param {Object} results - Das Ergebnisobjekt
 */
function addEmptyResults(results) {
    results.matchInfo.push(['']);
    results.kontoSollResults.push(['']);
    results.kontoHabenResults.push(['']);
    results.categoryResults.push(['']);
}

/**
 * Optimierte Version des Document-Type-Matchings
 * @param {Object} docData - Referenzdaten
 * @param {string} refNumber - Referenznummer
 * @param {number} betragValue - Betrag
 * @param {Array} row - Zeilendaten
 * @param {Object} columns - Spaltenkonfiguration
 * @param {string} docType - Dokumenttyp
 * @param {Object} docCols - Dokumentspaltenkonfiguration
 * @param {Object} kontoMapping - Kontomapping
 * @param {Object} results - Ergebnisobjekt
 * @param {string} category - Kategorie
 * @param {Object} config - Konfiguration
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {boolean} isGutschrift - Ist eine Gutschrift
 * @returns {boolean} true wenn eine Übereinstimmung gefunden wurde
 */
function processDocumentTypeMatching(docData, refNumber, betragValue, row, columns, docType, docCols, kontoMapping,
    results, category, config, firstDataRow, isGutschrift = false, isGesellschafterEinlage = false) {

    // Spezialbehandlung für Gutschriften und Gesellschafterkonto mit negativen Beträgen
    if (isGutschrift && docType !== 'einnahme' && docType !== 'gutschrift' && docType !== 'gesellschafterkonto') {
        return false;
    }

    // Match finding - no changes needed
    const matchResult = findMatch(refNumber, docData, isGutschrift || isGesellschafterEinlage ? null : betragValue);
    if (!matchResult) return false;

    // Prüfen auf Gutschrift oder Gesellschafterkonto-Rückzahlung (negativ)
    if (isGutschrift) {
        // Bei Gutschriften oder Gesellschafterkonto-Rückzahlungen müssen die Beträge negativ sein
        if (!matchResult.originalData || matchResult.originalData.betrag >= 0) {
            return false;
        }
    }

    // Prüfen auf Gesellschafterkonto-Einlage (positiv)
    if (isGesellschafterEinlage) {
        // Bei Gesellschafterkonto-Einlagen müssen die Beträge positiv sein
        if (!matchResult.originalData || matchResult.originalData.betrag <= 0) {
            return false;
        }
    }

    // Validate reference match
    if (matchResult.ref && !isGoodReferenceMatch(refNumber, matchResult.ref)) {
        return false;
    }

    // Create match info
    const matchInfoText = createMatchInfo(matchResult, docType, betragValue, isGutschrift, isGesellschafterEinlage);

    // Get category and konto from the document
    let sourceCategory = '';
    let sourceKonto = '';

    if (matchResult.originalData) {
        sourceCategory = matchResult.originalData.kategorie || '';
        sourceKonto = matchResult.originalData.buchungskonto || '';
    }

    // Set konten from mapping
    let kontoSoll = '', kontoHaben = '';

    // If source category is available, try to get mapping from appropriate category
    if (sourceCategory && matchResult.originalData && matchResult.originalData.sheetType) {
        // Determine the correct config key based on sheetType
        const configKey = matchResult.originalData.sheetType;

        // Safely access the category and its kontoMapping
        const categoryConfig = config[configKey]?.categories?.[sourceCategory];
        if (categoryConfig && categoryConfig.kontoMapping) {
            kontoSoll = categoryConfig.kontoMapping.soll || '';
            kontoHaben = categoryConfig.kontoMapping.gegen || '';
        }
    }

    // If we couldn't get mapping from source, try from bank category
    if ((!kontoSoll || !kontoHaben) && category) {
        // Determine the config key based on docType
        let configKey;
        if (docType === 'einnahme' || docType === 'gutschrift') {
            configKey = 'einnahmen';
        } else if (docType === 'ausgabe') {
            configKey = 'ausgaben';
        } else if (docType === 'eigenbeleg') {
            configKey = 'eigenbelege';
        } else if (docType === 'gesellschafterkonto' || docType === 'gesellschafterkontoEinlage') {
            configKey = 'gesellschafterkonto';
        } else if (docType === 'holdingtransfer') {
            configKey = 'holdingTransfers';
        }

        // Safely access the category config
        if (configKey) {
            const categoryConfig = config[configKey]?.categories?.[category];
            if (categoryConfig && categoryConfig.kontoMapping) {
                kontoSoll = kontoSoll || categoryConfig.kontoMapping.soll || '';
                kontoHaben = kontoHaben || categoryConfig.kontoMapping.gegen || '';
            }
        }
    }

    // Add results to arrays
    results.matchInfo.push([matchInfoText]);
    results.kontoSollResults.push([kontoSoll]);
    results.kontoHabenResults.push([kontoHaben]);
    results.categoryResults.push([category]);

    // Generate key for bank matches
    const key = `${docType}#${matchResult.row}`;

    // Detect if this is a gutschrift or Gesellschafterkonto with negative amount
    const isActualGutschrift = docType === 'gutschrift' ||
        (matchResult.originalData && matchResult.originalData.betrag < 0);

    // Detect if this is a Gesellschafterkonto entry with positive amount (Einlage)
    const isGesellschafterkontoEinlage = docType === 'gesellschafterkontoEinlage' ||
        (docType === 'gesellschafterkonto' && matchResult.originalData && matchResult.originalData.betrag > 0);

    // Store information for later marking
    results.bankZuordnungen[key] = {
        typ: isActualGutschrift ? 'gutschrift' :
            isGesellschafterkontoEinlage ? 'gesellschafterkontoEinlage' :
                (docType === 'gutschrift' ? 'einnahme' : docType),
        row: matchResult.row,
        bankRow: firstDataRow + results.matchInfo.length - 1,
        bankDatum: row[columns.datum - 1],
        matchInfo: matchInfoText,
        transTyp: isActualGutschrift ? 'Gutschrift' :
            isGesellschafterkontoEinlage ? 'Einlage' :
                row[columns.transaktionstyp - 1],
        category: category,
        sourceCategory: sourceCategory,
        kontoSoll: kontoSoll,
        kontoHaben: kontoHaben,
        isGutschrift: isActualGutschrift,
        isGesellschafterkontoEinlage: isGesellschafterkontoEinlage,
        originalRef: matchResult.ref,
    };

    return true;
}

/**
 * Sucht eine Übereinstimmung in der Referenz-Map mit optimierten Suchstrategien
 * @param {string} reference - Die Referenz
 * @param {Object} refMap - Die Referenzdaten
 * @param {number} betrag - Der Betrag
 * @returns {Object|null} Die gefundene Übereinstimmung oder null
 */
function findMatch(reference, refMap, betrag = null) {
    // Keine Daten oder keine Referenz
    if (!reference || !refMap) return null;

    // Ensure reference is a string
    const refString = String(reference).trim();
    if (!refString) return null;

    // Normalisierte Suche vorbereiten
    const normalizedRef = stringUtils.normalizeText(refString);
    if (!normalizedRef) return null;

    // Matching-Strategien in effizienter Reihenfolge
    let match = null;
    let matchedKey = null;

    // 1. Exakte Übereinstimmung (am häufigsten und schnellsten)
    if (refMap[refString]) {
        match = evaluateMatch(refMap[refString], betrag);
        if (match) matchedKey = refString;
    }

    // 2. Nur wenn keine exakte Übereinstimmung gefunden wurde: Normalisierte Übereinstimmung
    if (!match && normalizedRef && refMap[normalizedRef]) {
        match = evaluateMatch(refMap[normalizedRef], betrag);
        if (match) matchedKey = normalizedRef;
    }

    // 3. Nur wenn immer noch kein Match: Teilweise Übereinstimmung mit Optimierung
    if (!match) {
        // Optimierung: Zuerst prüfen wir Schlüssel, die die Referenz enthalten
        // Dies ist eine schnellere Strategie als alle Schlüssel zu prüfen
        const candidateKeys = Object.keys(refMap).filter(key =>
            typeof key === 'string' &&
            (key.includes(refString) || refString.includes(key)),
        );

        // Wenn wir potenzielle Kandidaten haben, prüfen wir diese
        for (const key of candidateKeys) {
            if (isGoodReferenceMatch(key, refString)) {
                match = evaluateMatch(refMap[key], betrag);
                if (match) {
                    matchedKey = key;
                    break;
                }
            }
        }
    }

    // Originaldata hinzufügen, wenn ein Match gefunden wurde
    if (match && matchedKey) {
        match.originalData = refMap[matchedKey].originalData;
        match.ref = refMap[matchedKey].ref; // Original-Referenz für Validierung
    }

    return match;
}

/**
 * Bewertet die Qualität einer Übereinstimmung basierend auf Beträgen
 * @param {Object} match - Die gefundene Übereinstimmung
 * @param {number} betrag - Der zu vergleichende Betrag
 * @returns {Object|null} Die bewertete Übereinstimmung oder null
 */
function evaluateMatch(match, betrag = null) {
    if (!match) return null;

    // Ergebnis mit ursprünglichen Daten initialisieren
    const result = { ...match };

    // Wenn kein Betrag zum Vergleich angegeben wurde
    if (betrag === null) {
        result.matchType = 'Referenz-Match';
        return result;
    }

    // Absolute Beträge für Vergleich verwenden
    const matchBrutto = Math.abs(match.brutto);
    const matchBezahlt = Math.abs(match.bezahlt);

    // Toleranzwert für Betragsabweichungen (2 Cent)
    const tolerance = 0.02;

    // Optimierter Vergleichsalgorithmus mit frühzeitigem Abbruch

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

    // Fall 3: Teilzahlung - nur wenn signifikant unter dem Bruttobetrag
    if (betrag < matchBrutto && (matchBrutto - betrag) > (matchBrutto * 0.1)) {
        result.matchType = 'Teilzahlung';
        return result;
    }

    // Fall 4: Betrag größer, aber nahe dem Bruttobetrag
    if (betrag > matchBrutto && numberUtils.isApproximatelyEqual(betrag, matchBrutto, tolerance)) {
        result.matchType = 'Vollständige Zahlung';
        return result;
    }

    // Fall 5: Alle anderen Fälle
    result.matchType = 'Unsichere Zahlung';
    result.betragsDifferenz = numberUtils.round(Math.abs(betrag - matchBrutto), 2);
    return result;
}

/**
 * Erstellt einen Informationstext für eine Match-Übereinstimmung
 * @param {Object} matchResult - Das Matching-Ergebnis
 * @param {string} docType - Der Dokumenttyp
 * @param {number} betragValue - Der Betrag
 * @returns {string} Der Informationstext
 */
function createMatchInfo(matchResult, docType, betragValue, isGutschrift = false, isGesellschafterEinlage = false) {
    let matchStatus = '';

    // Match-Status mit Differenzanzeige
    if (matchResult.matchType) {
        if (matchResult.matchType === 'Unsichere Zahlung' && matchResult.betragsDifferenz) {
            matchStatus = ` (${matchResult.matchType}, Diff: ${matchResult.betragsDifferenz}€)`;
        } else {
            matchStatus = ` (${matchResult.matchType})`;
        }
    }

    // Verbesserte Gutschrift-Erkennung
    const isActualGutschrift = docType === 'gutschrift' ||
        (matchResult.originalData && matchResult.originalData.betrag < 0 &&
            docType !== 'gesellschafterkontoEinlage' && docType !== 'holdingtransfer');

    // Gesellschafterkonto-Einlagen-Erkennung
    const isActualGesellschafterEinlage = docType === 'gesellschafterkontoEinlage' ||
        (docType === 'gesellschafterkonto' && matchResult.originalData && matchResult.originalData.betrag > 0);

    // Spezielle Behandlung für Gutschriften
    if (isActualGutschrift && docType !== 'gesellschafterkonto' && docType !== 'holdingtransfer') {
        const gutschriftBetrag = Math.abs(matchResult.originalData ?
            matchResult.originalData.betrag : matchResult.brutto);

        // Status für Gutschriften basierend auf Betrag
        if (numberUtils.isApproximatelyEqual(betragValue, gutschriftBetrag, 0.01)) {
            matchStatus = ' (Vollständige Gutschrift)';
        } else if (betragValue < gutschriftBetrag) {
            matchStatus = ' (Teilgutschrift)';
        } else {
            matchStatus = ' (Ungewöhnliche Gutschrift)';
        }

        return `Gutschrift zu Einnahme #${matchResult.row}${matchStatus}`;
    }

    // Spezielle Behandlung für Gesellschafterkonto
    if (docType === 'gesellschafterkonto' || docType === 'gesellschafterkontoEinlage') {
        const betrag = matchResult.originalData ? Math.abs(matchResult.originalData.betrag) : betragValue;

        // Status basierend auf Betrag und Vorzeichen
        if (isActualGutschrift) {
            // Gesellschafterkonto mit negativem Betrag (Rückzahlung)
            if (numberUtils.isApproximatelyEqual(betragValue, betrag, 0.01)) {
                matchStatus = ' (Vollständige Rückzahlung)';
            } else if (betragValue < betrag) {
                matchStatus = ' (Teilrückzahlung)';
            } else {
                matchStatus = ' (Ungewöhnliche Rückzahlung)';
            }
            return `Rückzahlung zu Gesellschafterkonto #${matchResult.row}${matchStatus}`;
        } else if (isActualGesellschafterEinlage || isGesellschafterEinlage) {
            // Gesellschafterkonto mit positivem Betrag (Einlage)
            if (numberUtils.isApproximatelyEqual(betragValue, betrag, 0.01)) {
                matchStatus = ' (Vollständige Einlage)';
            } else if (betragValue < betrag) {
                matchStatus = ' (Teileinlage)';
            } else {
                matchStatus = ' (Ungewöhnliche Einlage)';
            }
            return `Einlage zu Gesellschafterkonto #${matchResult.row}${matchStatus}`;
        }
    }

    // Spezielle Behandlung für Holding Transfers
    if (docType === 'holdingtransfer') {
        const betrag = matchResult.originalData ? Math.abs(matchResult.originalData.betrag) : betragValue;

        // Kategorie und isHolding extrahieren
        const kategorie = matchResult.originalData?.kategorie || '';
        const isHolding = matchResult.isHolding || false;

        // Status basierend auf Kategorie, isHolding und Betrag
        if (kategorie === 'Gewinnübertrag') {
            // Gewinnübertrag: Bei Holding GmbH Einnahme, bei operativer GmbH Ausgabe
            const perspective = isHolding ? 'Einnahme' : 'Ausgabe';
            if (numberUtils.isApproximatelyEqual(betragValue, betrag, 0.01)) {
                matchStatus = ' (Vollständige Übertragung)';
            } else if (betragValue < betrag) {
                matchStatus = ' (Teilübertragung)';
            } else {
                matchStatus = ' (Ungewöhnliche Übertragung)';
            }
            return `Gewinnübertrag ${perspective} #${matchResult.row}${matchStatus}`;
        }
        else if (kategorie === 'Kapitalrückführung') {
            // Kapitalrückführung: Bei Holding GmbH Ausgabe, bei operativer GmbH Einnahme
            const perspective = isHolding ? 'Ausgabe' : 'Einnahme';
            if (numberUtils.isApproximatelyEqual(betragValue, betrag, 0.01)) {
                matchStatus = ' (Vollständige Rückführung)';
            } else if (betragValue < betrag) {
                matchStatus = ' (Teilrückführung)';
            } else {
                matchStatus = ' (Ungewöhnliche Rückführung)';
            }
            return `Kapitalrückführung ${perspective} #${matchResult.row}${matchStatus}`;
        }
        else {
            // Andere Holding Transfer Typen
            if (numberUtils.isApproximatelyEqual(betragValue, betrag, 0.01)) {
                matchStatus = ' (Vollständig)';
            } else if (betragValue < betrag) {
                matchStatus = ' (Teilweise)';
            } else {
                matchStatus = ' (Ungewöhnlich)';
            }
            return `Holding Transfer ${kategorie} #${matchResult.row}${matchStatus}`;
        }
    }

    // Konstanten für Dokumenttypen verbessern die Lesbarkeit und Performance
    const docTypeNames = {
        'einnahme': 'Einnahme',
        'ausgabe': 'Ausgabe',
        'eigenbeleg': 'Eigenbeleg',
        'gesellschafterkonto': 'Gesellschafterkonto',
        'holdingtransfer': 'Holding Transfer',
    };

    return `${docTypeNames[docType] || docType} #${matchResult.row}${matchStatus}`;
}

/**
 * Setzt Standard-Konten basierend auf der Kategorie und Transaktionstyp
 * @param {string} category - Die Kategorie
 * @param {string} tranType - Der Transaktionstyp
 * @param {Object} config - Die Konfiguration
 * @param {string} fallbackKontoSoll - Das Fallback-Konto Soll
 * @param {string} fallbackKontoHaben - Das Fallback-Konto Haben
 * @param {Object} results - Das Ergebnisobjekt
 */
function setDefaultAccounts(category, tranType, config, fallbackKontoSoll, fallbackKontoHaben, results) {
    let kontoSoll = fallbackKontoSoll;
    let kontoHaben = fallbackKontoHaben;

    if (category) {
        // Select configuration source based on transaction type
        const configKey = tranType === 'Einnahme' ? 'einnahmen' : 'ausgaben';

        // Safely access the category configuration
        if (config[configKey] &&
            config[configKey].categories &&
            config[configKey].categories[category] &&
            config[configKey].categories[category].kontoMapping) {

            const mapping = config[configKey].categories[category].kontoMapping;
            kontoSoll = mapping.soll || fallbackKontoSoll;
            kontoHaben = mapping.gegen || fallbackKontoHaben;
        }
    }

    // Add results
    results.matchInfo.push(['']);
    results.kontoSollResults.push([kontoSoll]);
    results.kontoHabenResults.push([kontoHaben]);
    results.categoryResults.push([category]);
}

/**
 * Führt Batch-Updates für alle betroffenen Spalten aus
 * @param {Sheet} sheet - Das Sheet
 * @param {number} firstDataRow - Erste Datenzeile
 * @param {number} numDataRows - Anzahl der Datenzeilen
 * @param {Object} columns - Spaltenkonfiguration
 * @param {Array} matchInfo - Match-Informationen
 * @param {Array} kontoSollResults - Konto-Soll-Ergebnisse
 * @param {Array} kontoHabenResults - Konto-Haben-Ergebnisse
 * @param {Array} categoryResults - Kategorie-Ergebnisse
 */
function performBatchUpdates(sheet, firstDataRow, numDataRows, columns,
    matchInfo, kontoSollResults, kontoHabenResults, categoryResults) {

    // Batch-Updates für alle Spalten in separaten API-Calls
    if (columns.matchInfo) {
        sheet.getRange(firstDataRow, columns.matchInfo, numDataRows, 1).setValues(matchInfo);
    }

    if (columns.kontoSoll) {
        sheet.getRange(firstDataRow, columns.kontoSoll, numDataRows, 1).setValues(kontoSollResults);
    }

    if (columns.kontoHaben) {
        sheet.getRange(firstDataRow, columns.kontoHaben, numDataRows, 1).setValues(kontoHabenResults);
    }

    // Kategorien aktualisieren, wenn Kategorie-Spalte existiert
    if (columns.kategorie && categoryResults.length === numDataRows) {
        sheet.getRange(firstDataRow, columns.kategorie, numDataRows, 1).setValues(categoryResults);
    }
}

/**
 * Prüft, ob zwei Referenznummern gut genug übereinstimmen
 * @param {string} ref1 - Erste Referenz
 * @param {string} ref2 - Zweite Referenz
 * @returns {boolean} true wenn die Referenzen übereinstimmen
 */
function isGoodReferenceMatch(ref1, ref2) {
    // Leere Werte behandeln
    if (!ref1 || !ref2) return false;

    // Sicherstellen, dass beide Referenzen Strings sind
    const ref1Str = String(ref1);
    const ref2Str = String(ref2);

    // Frühe Rückkehr bei exakter Übereinstimmung
    if (ref1Str === ref2Str) return true;

    // Schnelle Prüfung, ob eine Referenz die andere enthält
    if (ref1Str.includes(ref2Str) || ref2Str.includes(ref1Str)) {
        // Bei kurzen Referenzen (weniger als 5 Zeichen) strengere Prüfung
        const shorter = ref1Str.length <= ref2Str.length ? ref1Str : ref2Str;
        if (shorter.length < 5) {
            return ref1Str.startsWith(ref2Str) || ref2Str.startsWith(ref1Str);
        }
        return true;
    }

    // Keine gute Übereinstimmung
    return false;
}

/**
 * Verarbeitet spezifisch Gesellschafterkonto-Matches
 * @param {Object} gesellschafterData - Referenzdaten für Gesellschafterkonto
 * @param {string} refNumber - Referenznummer
 * @param {number} betragValue - Betrag aus Bankbewegung (absoluter Wert)
 * @param {Array} row - Zeilendaten
 * @param {Object} columns - Spaltenkonfiguration
 * @param {boolean} isPositive - True für Einlagen (positive Beträge), False für Rückzahlungen (negative)
 * @param {Object} config - Die Konfiguration
 * @param {Object} results - Ergebnisobjekt
 * @param {number} firstDataRow - Erste Datenzeile
 * @returns {boolean} true wenn eine Übereinstimmung gefunden wurde
 */
function processGesellschafterkontoMatching(gesellschafterData, refNumber, betragValue, row, columns, isPositive, config, results, firstDataRow) {
    // Match finding für Gesellschafterkonto
    const matchResult = findGesellschafterMatch(gesellschafterData, refNumber, betragValue, isPositive);
    if (!matchResult) return false;

    // Validate reference match
    if (matchResult.ref && !isGoodReferenceMatch(refNumber, matchResult.ref)) {
        return false;
    }

    // Create match info mit korrekter Bezeichnung für Gesellschafterkonto
    const matchType = isPositive ? 'Einlage' : 'Rückzahlung';
    const matchInfoText = `Gesellschafterkonto ${matchType} #${matchResult.row} (${matchResult.matchType || ''})`;

    // Get category from the document
    let sourceCategory = '';
    if (matchResult.originalData) {
        sourceCategory = matchResult.originalData.kategorie || '';
    }

    // Kategorie aus Bankbewegung oder fallback
    const category = row[columns.kategorie - 1] ?
        row[columns.kategorie - 1].toString().trim() : sourceCategory;

    // Set konten from mapping
    let kontoSoll = '', kontoHaben = '';

    // Mapping aus Gesellschafterkonto-Kategorie verwenden
    if (sourceCategory) {
        const categoryConfig = config.gesellschafterkonto?.categories?.[sourceCategory];
        if (categoryConfig && categoryConfig.kontoMapping) {
            kontoSoll = categoryConfig.kontoMapping.soll || '';
            kontoHaben = categoryConfig.kontoMapping.gegen || '';
        }
    }

    // Add results to arrays
    results.matchInfo.push([matchInfoText]);
    results.kontoSollResults.push([kontoSoll]);
    results.kontoHabenResults.push([kontoHaben]);
    results.categoryResults.push([category]);

    // Generate key for bank matches
    const key = `gesellschafterkonto#${matchResult.row}`;

    // Store information for later marking
    results.bankZuordnungen[key] = {
        typ: 'gesellschafterkonto',
        row: matchResult.row,
        bankRow: firstDataRow + results.matchInfo.length - 1,
        bankDatum: row[columns.datum - 1],
        matchInfo: matchInfoText,
        transTyp: isPositive ? 'Einlage' : 'Rückzahlung',
        category: category,
        sourceCategory: sourceCategory,
        kontoSoll: kontoSoll,
        kontoHaben: kontoHaben,
        isPositive: isPositive,
        originalRef: matchResult.ref,
    };

    return true;
}

/**
 * Sucht spezifisch nach Gesellschafterkonto-Übereinstimmungen
 * @param {Object} refMap - Referenzdaten für Gesellschafterkonto
 * @param {string} reference - Referenznummer
 * @param {number} betrag - Betrag (absoluter Wert)
 * @param {boolean} isPositive - True für Einlagen, False für Rückzahlungen
 * @returns {Object|null} Matchergebnis oder null
 */
function findGesellschafterMatch(refMap, reference, betrag, isPositive) {
    // Keine Daten oder keine Referenz
    if (!reference || !refMap) return null;

    // Ensure reference is a string
    const refString = String(reference).trim();
    if (!refString) return null;

    // Normalisierte Suche vorbereiten
    const normalizedRef = stringUtils.normalizeText(refString);
    if (!normalizedRef) return null;

    // Potenzielle Treffer sammeln
    const candidates = [];

    // Durchsuche alle Einträge im refMap
    for (const key in refMap) {
        const entry = refMap[key];
        // Prüfe, ob Vorzeichen passt
        const hasCorrectSign = isPositive ? (entry.betrag > 0) : (entry.betrag < 0);

        if (!hasCorrectSign) continue;

        // Prüfe Referenz auf Übereinstimmung
        if (key === refString || key === normalizedRef ||
            isGoodReferenceMatch(key, refString)) {

            // Betrag prüfen
            const entryBetrag = Math.abs(entry.betrag);
            if (numberUtils.isApproximatelyEqual(betrag, entryBetrag, 0.02)) {
                candidates.push({...entry, matchType: 'Vollständige Zahlung'});
            } else if (Math.abs(betrag - entryBetrag) <= 0.1 * entryBetrag) {
                candidates.push({...entry, matchType: 'Annähernde Zahlung'});
            }
        }
    }

    // Besten Kandidaten zurückgeben (wenn vorhanden)
    if (candidates.length > 0) {
        // Sortiere nach Matchqualität, bevorzuge "Vollständige Zahlung"
        candidates.sort((a, b) => {
            if (a.matchType === 'Vollständige Zahlung' && b.matchType !== 'Vollständige Zahlung') return -1;
            if (a.matchType !== 'Vollständige Zahlung' && b.matchType === 'Vollständige Zahlung') return 1;
            return 0;
        });
        return candidates[0];
    }

    return null;
}

/**
 * Verarbeitet spezifisch Holding Transfer-Matches
 * @param {Object} holdingData - Referenzdaten für Holding Transfers
 * @param {string} refNumber - Referenznummer
 * @param {number} betragValue - Betrag aus Bankbewegung (absoluter Wert)
 * @param {Array} row - Zeilendaten
 * @param {Object} columns - Spaltenkonfiguration
 * @param {boolean} isPositive - Betrag positiv oder negativ in der Bankbewegung
 * @param {Object} config - Die Konfiguration
 * @param {Object} results - Ergebnisobjekt
 * @param {number} firstDataRow - Erste Datenzeile
 * @returns {boolean} true wenn eine Übereinstimmung gefunden wurde
 */
function processHoldingTransferMatching(holdingData, refNumber, betragValue, row, columns, isPositive, config, results, firstDataRow) {
    // Prüfe, ob wir eine Holding GmbH oder eine operative GmbH haben
    const isHolding = config.tax.isHolding || false;

    // Match finding für Holding Transfers - Beachte, dass die erwartete Vorzeichenrichtung
    // vom isHolding-Flag abhängt
    const matchResult = findHoldingTransferMatch(holdingData, refNumber, betragValue, isPositive, isHolding);
    if (!matchResult) return false;

    // Validate reference match
    if (matchResult.ref && !isGoodReferenceMatch(refNumber, matchResult.ref)) {
        return false;
    }

    // Kategorie aus dem Originaldokument extrahieren
    const kategorie = matchResult.originalData?.kategorie || '';

    // Match-Typ basierend auf Kategorie und Buchungsrichtung bestimmen
    let matchType = 'Transfer';
    if (kategorie === 'Gewinnübertrag') {
        matchType = 'Gewinnübertrag';
    } else if (kategorie === 'Kapitalrückführung') {
        matchType = 'Kapitalrückführung';
    }

    // Match-Info mit Berücksichtigung der Perspektive (Holding oder Operative)
    const matchInfoText = `Holding Transfer ${matchType} #${matchResult.row} (${matchResult.matchType || ''})`;

    // Get category from the document
    let sourceCategory = '';
    if (matchResult.originalData) {
        sourceCategory = matchResult.originalData.kategorie || '';
    }

    // Kategorie aus Bankbewegung oder fallback
    const category = row[columns.kategorie - 1] ?
        row[columns.kategorie - 1].toString().trim() : sourceCategory;

    // Set konten from mapping
    let kontoSoll = '', kontoHaben = '';

    // Mapping aus Holding Transfer-Kategorie verwenden
    if (sourceCategory) {
        const categoryConfig = config.holdingTransfers?.categories?.[sourceCategory];
        if (categoryConfig && categoryConfig.kontoMapping) {
            kontoSoll = categoryConfig.kontoMapping.soll || '';
            kontoHaben = categoryConfig.kontoMapping.gegen || '';
        }
    }

    // Add results to arrays
    results.matchInfo.push([matchInfoText]);
    results.kontoSollResults.push([kontoSoll]);
    results.kontoHabenResults.push([kontoHaben]);
    results.categoryResults.push([category]);

    // Generate key for bank matches
    const key = `holdingtransfer#${matchResult.row}`;

    // Store information for later marking
    results.bankZuordnungen[key] = {
        typ: 'holdingtransfer',
        row: matchResult.row,
        bankRow: firstDataRow + results.matchInfo.length - 1,
        bankDatum: row[columns.datum - 1],
        matchInfo: matchInfoText,
        transTyp: matchType,
        category: category,
        sourceCategory: sourceCategory,
        kontoSoll: kontoSoll,
        kontoHaben: kontoHaben,
        isPositive: isPositive,
        isHolding: isHolding,
        originalRef: matchResult.ref,
        originalCategory: kategorie,
    };

    return true;
}

/**
 * Sucht spezifisch nach Holding Transfer-Übereinstimmungen
 * @param {Object} refMap - Referenzdaten für Holding Transfers
 * @param {string} reference - Referenznummer
 * @param {number} betrag - Betrag (absoluter Wert)
 * @param {boolean} isPositive - Ist der Bankbetrag positiv
 * @param {boolean} isHolding - Ist es eine Holding GmbH
 * @returns {Object|null} Matchergebnis oder null
 */
function findHoldingTransferMatch(refMap, reference, betrag, isPositive, isHolding) {
    // Keine Daten oder keine Referenz
    if (!reference || !refMap) return null;

    // Ensure reference is a string
    const refString = String(reference).trim();
    if (!refString) return null;

    // Normalisierte Suche vorbereiten
    const normalizedRef = stringUtils.normalizeText(refString);
    if (!normalizedRef) return null;

    // Potenzielle Treffer sammeln
    const candidates = [];

    // Durchsuche alle Einträge im refMap
    for (const key in refMap) {
        const entry = refMap[key];
        const category = entry.originalData?.kategorie;

        // Prüfe, ob Vorzeichen passt basierend auf Kategorie und isHolding
        let hasCorrectSign = false;

        if (category === 'Gewinnübertrag') {
            // Für Gewinnübertrag:
            // - Bei operativer GmbH (isHolding=false): Betrag negativ, Bankbuchung negativ
            // - Bei Holding GmbH (isHolding=true): Betrag positiv, Bankbuchung positiv
            hasCorrectSign = (isHolding && entry.betrag > 0 && isPositive) ||
                (!isHolding && entry.betrag < 0 && !isPositive);
        }
        else if (category === 'Kapitalrückführung') {
            // Für Kapitalrückführung:
            // - Bei operativer GmbH: Betrag positiv, Bankbuchung positiv
            // - Bei Holding GmbH: Betrag negativ, Bankbuchung negativ
            hasCorrectSign = (isHolding && entry.betrag < 0 && !isPositive) ||
                (!isHolding && entry.betrag > 0 && isPositive);
        }
        else {
            // Für andere Kategorien: prüfe basierend auf Vorzeichen
            hasCorrectSign = (entry.betrag > 0 && isPositive) || (entry.betrag < 0 && !isPositive);
        }

        if (!hasCorrectSign) continue;

        // Prüfe Referenz auf Übereinstimmung
        if (key === refString || key === normalizedRef ||
            isGoodReferenceMatch(key, refString)) {

            // Betrag prüfen
            const entryBetrag = Math.abs(entry.betrag);
            if (numberUtils.isApproximatelyEqual(betrag, entryBetrag, 0.02)) {
                candidates.push({...entry, matchType: 'Vollständige Zahlung'});
            } else if (Math.abs(betrag - entryBetrag) <= 0.1 * entryBetrag) {
                candidates.push({...entry, matchType: 'Annähernde Zahlung'});
            }
        }
    }

    // Besten Kandidaten zurückgeben (wenn vorhanden)
    if (candidates.length > 0) {
        // Sortiere nach Matchqualität, bevorzuge "Vollständige Zahlung"
        candidates.sort((a, b) => {
            if (a.matchType === 'Vollständige Zahlung' && b.matchType !== 'Vollständige Zahlung') return -1;
            if (a.matchType !== 'Vollständige Zahlung' && b.matchType === 'Vollständige Zahlung') return 1;
            return 0;
        });
        return candidates[0];
    }

    return null;
}

export default {
    performBankReferenceMatching,
    isGoodReferenceMatch,
};