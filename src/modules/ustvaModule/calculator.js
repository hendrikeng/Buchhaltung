// modules/ustvaModule/calculator.js
import dateUtils from '../../utils/dateUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import stringUtils from '../../utils/stringUtils.js';
import dataModel from './dataModel.js';

// Cache für häufig verwendete Steuerberechnungen
const taxCalculationCache = new Map();

/**
 * Verarbeitet eine Zeile aus den Einnahmen/Ausgaben/Eigenbelegen für die UStVA
 * Optimierte Version mit effizienter Kategorisierung und Berechnungslogik
 *
 * @param {Array} row - Datenzeile aus einem Sheet
 * @param {Object} data - UStVA-Datenobjekt nach Monaten
 * @param {boolean} isIncome - Handelt es sich um Einnahmen (true) oder Ausgaben (false)
 * @param {boolean} isEigen - Handelt es sich um Eigenbelege (true oder false)
 * @param {Object} config - Die Konfiguration
 */
function processUStVARow(row, data, isIncome, isEigen = false, config) {
    try {
        // Sheet-Typ bestimmen
        const sheetType = isIncome ? 'einnahmen' : isEigen ? 'eigenbelege' : 'ausgaben';
        const columns = config[sheetType].columns;

        // Zahlungsdatum prüfen (nur abgeschlossene Zahlungen)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Cache-Schlüssel für Datum und Monat
        const dateCacheKey = `date_${row[columns.zahlungsdatum - 1]}`;

        // Monat und Jahr prüfen (nur relevantes Geschäftsjahr)
        const month = dateUtils.getMonthFromRow(row, sheetType, config);
        if (!month || month < 1 || month > 12) return;

        // Beträge aus der Zeile extrahieren
        const netto = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        const restNetto = numberUtils.parseCurrency(row[columns.restbetragNetto - 1]) || 0;
        const gezahlt = netto - restNetto; // Tatsächlich gezahlter/erhaltener Betrag

        // Falls kein Betrag gezahlt wurde, nichts zu verarbeiten
        if (numberUtils.isApproximatelyEqual(gezahlt, 0)) return;

        // NEU: Ausland-Information auslesen
        const auslandWert = columns.ausland && row[columns.ausland - 1] ?
            row[columns.ausland - 1].toString().trim() : 'Inland';
        const isEuAusland = auslandWert === 'EU-Ausland';
        const isNichtEuAusland = auslandWert === 'Nicht-EU-Ausland';

        // MwSt-Rate und Steuer berechnen mit Cache
        const mwstStr = row[columns.mwstSatz - 1]?.toString() || '';
        const cacheKey = `${mwstStr}_${gezahlt}`;

        let mwstRate, roundedRate, tax;

        if (taxCalculationCache.has(cacheKey)) {
            const cached = taxCalculationCache.get(cacheKey);
            mwstRate = cached.mwstRate;
            roundedRate = cached.roundedRate;
            tax = cached.tax;
        } else {
            mwstRate = numberUtils.parseMwstRate(mwstStr, config.tax.defaultMwst);
            roundedRate = Math.round(mwstRate);
            tax = numberUtils.round(gezahlt * (mwstRate / 100), 2);

            // Cache-Ergebnis für künftige Berechnungen
            taxCalculationCache.set(cacheKey, { mwstRate, roundedRate, tax });
        }

        // Kategorie ermitteln
        const category = row[columns.kategorie - 1]?.toString().trim() || '';

        // Verarbeitung basierend auf Zeilentyp und Ausland-Status durch optimierte Handler
        if (isIncome) {
            // Bei Einnahmen UStVA-Behandlung nach Ausland-Status
            if (isEuAusland) {
                // Innergemeinschaftliche Lieferungen sind steuerfrei
                data[month].innergemeinschaftliche_lieferungen += gezahlt;
            } else if (isNichtEuAusland) {
                // Ausfuhrlieferungen in Drittländer sind steuerfrei
                data[month].steuerfreie_ausland_einnahmen += gezahlt;
            } else {
                // Normale Verarbeitung für Inland
                processIncomeRow(data, month, gezahlt, tax, roundedRate, category, config);
            }
        } else if (isEigen) {
            // Eigenbelege normal verarbeiten
            processEigenRow(data, month, gezahlt, tax, roundedRate, category, config);
        } else {
            // Bei Ausgaben UStVA-Behandlung nach Ausland-Status
            if (isEuAusland) {
                // Innergemeinschaftliche Erwerbe: Erwerbsteuer + Vorsteuer bei B2B
                data[month].innergemeinschaftliche_erwerbe += gezahlt;

                // Erwerbsteuer nach MwSt-Satz - diese ist sowohl Zahlungspflicht als auch Vorsteuerabzug
                if (roundedRate === 7) {
                    data[month].erwerbsteuer_7 += tax;
                    data[month].vst_7 += tax; // Gleichzeitig Vorsteuerabzug
                } else if (roundedRate === 19) {
                    data[month].erwerbsteuer_19 += tax;
                    data[month].vst_19 += tax; // Gleichzeitig Vorsteuerabzug
                }
            } else if (isNichtEuAusland) {
                // Einfuhren aus Drittländern
                data[month].steuerfreie_ausland_ausgaben += gezahlt;
            } else {
                // Normale Verarbeitung für Inland
                processExpenseRow(data, month, gezahlt, tax, roundedRate, category, config);
            }
        }
    } catch (e) {
        console.error('Fehler bei der Verarbeitung einer UStVA-Zeile:', e);
    }
}

/**
 * Verarbeitet eine Einnahmenzeile für die UStVA mit optimierter Kategorisierung
 * @param {Object} data - UStVA-Datenobjekt nach Monaten
 * @param {number} month - Monat (1-12)
 * @param {number} gezahlt - Gezahlter Betrag
 * @param {number} tax - Berechnete Steuer
 * @param {number} roundedRate - Gerundeter Steuersatz
 * @param {string} category - Kategorie
 * @param {Object} config - Die Konfiguration
 */
function processIncomeRow(data, month, gezahlt, tax, roundedRate, category, config) {
    // Kategorie-Konfiguration mit Fallback
    const catCfg = config.einnahmen.categories[category] ?? {};
    const taxType = catCfg.taxType ?? 'steuerpflichtig';

    // Optimierung: Direktere Zuordnung basierend auf Steuertyp
    if (taxType === 'steuerfrei_inland') {
        // Steuerfreie Einnahmen im Inland
        data[month].steuerfreie_inland_einnahmen += gezahlt;
    } else if (taxType === 'steuerfrei_ausland' || !roundedRate) {
        // Steuerfreie Einnahmen aus dem Ausland oder ohne MwSt
        data[month].steuerfreie_ausland_einnahmen += gezahlt;
    } else {
        // Steuerpflichtige Einnahmen
        data[month].steuerpflichtige_einnahmen += gezahlt;

        // USt nach validem Steuersatz addieren (7% oder 19%)
        if (roundedRate === 7 || roundedRate === 19) {
            data[month][`ust_${roundedRate}`] += tax;
        }
    }
}

/**
 * Verarbeitet eine Eigenbeleg-Zeile für die UStVA mit optimierter Kategorisierung
 * @param {Object} data - UStVA-Datenobjekt nach Monaten
 * @param {number} month - Monat (1-12)
 * @param {number} gezahlt - Gezahlter Betrag
 * @param {number} tax - Berechnete Steuer
 * @param {number} roundedRate - Gerundeter Steuersatz
 * @param {string} category - Kategorie
 * @param {Object} config - Die Konfiguration
 */
function processEigenRow(data, month, gezahlt, tax, roundedRate, category, config) {
    // Kategorie-Konfiguration mit Fallback
    const eigenCfg = config.eigenbelege.categories[category] ?? {};
    const taxType = eigenCfg.taxType ?? 'steuerpflichtig';
    const besonderheit = eigenCfg.besonderheit;

    // Optimierung: Effizientere Verzweigungslogik
    if (taxType === 'steuerfrei') {
        // Steuerfreie Eigenbelege
        data[month].eigenbelege_steuerfrei += gezahlt;
    } else if (taxType === 'eigenbeleg' && besonderheit === 'bewirtung') {
        // Bewirtungsbelege (nur 70% der Vorsteuer absetzbar)
        data[month].eigenbelege_steuerpflichtig += gezahlt;

        if (roundedRate === 7 || roundedRate === 19) {
            // Optimierung: Nur einmal berechnen
            const vst70 = numberUtils.round(tax * 0.7, 2);
            const vst30 = tax - vst70; // Exakter als tax * 0.3 wegen Rundungsfehlern

            data[month][`vst_${roundedRate}`] += vst70;
            data[month].nicht_abzugsfaehige_vst += vst30;
        }
    } else {
        // Normale steuerpflichtige Eigenbelege
        data[month].eigenbelege_steuerpflichtig += gezahlt;

        if (roundedRate === 7 || roundedRate === 19) {
            data[month][`vst_${roundedRate}`] += tax;
        }
    }
}

/**
 * Verarbeitet eine Ausgabenzeile für die UStVA mit optimierter Kategorisierung
 * @param {Object} data - UStVA-Datenobjekt nach Monaten
 * @param {number} month - Monat (1-12)
 * @param {number} gezahlt - Gezahlter Betrag
 * @param {number} tax - Berechnete Steuer
 * @param {number} roundedRate - Gerundeter Steuersatz
 * @param {string} category - Kategorie
 * @param {Object} config - Die Konfiguration
 * @param {boolean} isEuAusland - True, wenn Ausland-Status EU-Ausland ist
 */
function processExpenseRow(data, month, gezahlt, tax, roundedRate, category, config, isEuAusland = false) {
    // Kategorie-Konfiguration mit Fallback
    const catCfg = config.ausgaben.categories[category] ?? {};
    const taxType = catCfg.taxType ?? 'steuerpflichtig';
    const besonderheit = catCfg.besonderheit;

    // Optimierung: Direkte Zuordnung mit Map-ähnlichem Pattern
    const typeHandlers = {
        'steuerfrei_inland': () => {
            data[month].steuerfreie_inland_ausgaben += gezahlt;
        },
        'steuerfrei_ausland': () => {
            data[month].steuerfreie_ausland_ausgaben += gezahlt;
        },
        'steuerpflichtig': () => {
            data[month].steuerpflichtige_ausgaben += gezahlt;

            // Besondere Behandlung für Bewirtung
            if (besonderheit === 'bewirtung' && (roundedRate === 7 || roundedRate === 19)) {
                // Bei Bewirtung nur 70% VSt abziehbar
                const vst70 = numberUtils.round(tax * 0.7, 2);
                const vst30 = tax - vst70;

                data[month][`vst_${roundedRate}`] += vst70;
                data[month].nicht_abzugsfaehige_vst += vst30;
            }
            // Für EU-Ausland brauchen wir ggf. spezielle Behandlung wegen Reverse Charge
            else if (!isEuAusland && (roundedRate === 7 || roundedRate === 19)) {
                data[month][`vst_${roundedRate}`] += tax;
            }
        },
    };

    // Handler ausführen oder Default, wenn nicht gefunden
    (typeHandlers[taxType] || typeHandlers['steuerpflichtig'])();
}

/**
 * Aggregiert UStVA-Daten für einen Zeitraum (z.B. Quartal oder Jahr)
 * Optimierte Version mit effizientem Datenzugriff
 *
 * @param {Object} data - UStVA-Datenobjekt nach Monaten
 * @param {number} start - Startmonat (1-12)
 * @param {number} end - Endmonat (1-12)
 * @returns {Object} Aggregiertes UStVA-Datenobjekt
 */
function aggregateUStVA(data, start, end) {
    const sum = dataModel.createEmptyUStVA();

    // Optimierung: Sammle alle relevanten Monatsobjekte vor der Verarbeitung
    const relevantMonths = [];
    for (let m = start; m <= end; m++) {
        if (data[m]) {
            relevantMonths.push(data[m]);
        }
    }

    // Optimierung: Verarbeite alle Felder einmal statt verschachtelte Schleife
    Object.keys(sum).forEach(key => {
        sum[key] = relevantMonths.reduce((total, month) => total + (month[key] || 0), 0);
    });

    return sum;
}

export default {
    processUStVARow,
    aggregateUStVA,
};