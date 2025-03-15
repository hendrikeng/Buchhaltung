// modules/ustvaModule/calculator.js
import dateUtils from '../../utils/dateUtils.js';
import numberUtils from '../../utils/numberUtils.js';
import stringUtils from '../../utils/stringUtils.js';
import dataModel from './dataModel.js';

/**
 * Verarbeitet eine Zeile aus den Einnahmen/Ausgaben/Eigenbelegen für die UStVA
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
        const sheetType = isIncome ? "einnahmen" : isEigen ? "eigenbelege" : "ausgaben";
        const columns = config[sheetType].columns;

        // Zahlungsdatum prüfen (nur abgeschlossene Zahlungen)
        const paymentDate = dateUtils.parseDate(row[columns.zahlungsdatum - 1]);
        if (!paymentDate || paymentDate > new Date()) return;

        // Monat und Jahr prüfen (nur relevantes Geschäftsjahr)
        const month = dateUtils.getMonthFromRow(row, sheetType, config);
        if (!month || month < 1 || month > 12) return;

        // Beträge aus der Zeile extrahieren
        const netto = numberUtils.parseCurrency(row[columns.nettobetrag - 1]);
        const restNetto = numberUtils.parseCurrency(row[columns.restbetragNetto - 1]) || 0; // Steuerbemessungsgrundlage für Teilzahlungen
        const gezahlt = netto - restNetto; // Tatsächlich gezahlter/erhaltener Betrag

        // Falls kein Betrag gezahlt wurde, nichts zu verarbeiten
        if (numberUtils.isApproximatelyEqual(gezahlt, 0)) return;

        // MwSt-Satz normalisieren
        const mwstRate = numberUtils.parseMwstRate(row[columns.mwstSatz - 1], config.tax.defaultMwst);
        const roundedRate = Math.round(mwstRate);

        // Steuer berechnen
        const tax = numberUtils.round(gezahlt * (mwstRate / 100), 2);

        // Kategorie ermitteln
        const category = row[columns.kategorie - 1]?.toString().trim() || "";

        // Verarbeitung basierend auf Zeilentyp
        if (isIncome) {
            processIncomeRow(data, month, gezahlt, tax, roundedRate, category, config);
        } else if (isEigen) {
            processEigenRow(data, month, gezahlt, tax, roundedRate, category, config);
        } else {
            processExpenseRow(data, month, gezahlt, tax, roundedRate, category, config);
        }
    } catch (e) {
        console.error("Fehler bei der Verarbeitung einer UStVA-Zeile:", e);
    }
}

/**
 * Verarbeitet eine Einnahmenzeile für die UStVA
 * @param {Object} data - UStVA-Datenobjekt nach Monaten
 * @param {number} month - Monat (1-12)
 * @param {number} gezahlt - Gezahlter Betrag
 * @param {number} tax - Berechnete Steuer
 * @param {number} roundedRate - Gerundeter Steuersatz
 * @param {string} category - Kategorie
 * @param {Object} config - Die Konfiguration
 */
function processIncomeRow(data, month, gezahlt, tax, roundedRate, category, config) {
    // EINNAHMEN
    const catCfg = config.einnahmen.categories[category] ?? {};
    const taxType = catCfg.taxType ?? "steuerpflichtig";

    if (taxType === "steuerfrei_inland") {
        // Steuerfreie Einnahmen im Inland (z.B. Vermietung)
        data[month].steuerfreie_inland_einnahmen += gezahlt;
    } else if (taxType === "steuerfrei_ausland" || !roundedRate) {
        // Steuerfreie Einnahmen aus dem Ausland oder ohne MwSt
        data[month].steuerfreie_ausland_einnahmen += gezahlt;
    } else {
        // Steuerpflichtige Einnahmen
        data[month].steuerpflichtige_einnahmen += gezahlt;

        // USt nach Steuersatz addieren
        // Wir prüfen ob es ein gültiger Steuersatz ist
        if (roundedRate === 7 || roundedRate === 19) {
            data[month][`ust_${roundedRate}`] += tax;
        }
    }
}

/**
 * Verarbeitet eine Eigenbeleg-Zeile für die UStVA
 * @param {Object} data - UStVA-Datenobjekt nach Monaten
 * @param {number} month - Monat (1-12)
 * @param {number} gezahlt - Gezahlter Betrag
 * @param {number} tax - Berechnete Steuer
 * @param {number} roundedRate - Gerundeter Steuersatz
 * @param {string} category - Kategorie
 * @param {Object} config - Die Konfiguration
 */
function processEigenRow(data, month, gezahlt, tax, roundedRate, category, config) {
    // EIGENBELEGE
    const eigenCfg = config.eigenbelege.categories[category] ?? {};
    const taxType = eigenCfg.taxType ?? "steuerpflichtig";

    if (taxType === "steuerfrei") {
        // Steuerfreie Eigenbelege
        data[month].eigenbelege_steuerfrei += gezahlt;
    } else if (taxType === "eigenbeleg" && eigenCfg.besonderheit === "bewirtung") {
        // Bewirtungsbelege (nur 70% der Vorsteuer absetzbar)
        data[month].eigenbelege_steuerpflichtig += gezahlt;

        if (roundedRate === 7 || roundedRate === 19) {
            const vst70 = numberUtils.round(tax * 0.7, 2); // 70% absetzbare Vorsteuer
            const vst30 = numberUtils.round(tax * 0.3, 2); // 30% nicht absetzbar
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
 * Verarbeitet eine Ausgabenzeile für die UStVA
 * @param {Object} data - UStVA-Datenobjekt nach Monaten
 * @param {number} month - Monat (1-12)
 * @param {number} gezahlt - Gezahlter Betrag
 * @param {number} tax - Berechnete Steuer
 * @param {number} roundedRate - Gerundeter Steuersatz
 * @param {string} category - Kategorie
 * @param {Object} config - Die Konfiguration
 */
function processExpenseRow(data, month, gezahlt, tax, roundedRate, category, config) {
    // AUSGABEN
    const catCfg = config.ausgaben.categories[category] ?? {};
    const taxType = catCfg.taxType ?? "steuerpflichtig";

    if (taxType === "steuerfrei_inland") {
        // Steuerfreie Ausgaben im Inland
        data[month].steuerfreie_inland_ausgaben += gezahlt;
    } else if (taxType === "steuerfrei_ausland") {
        // Steuerfreie Ausgaben im Ausland
        data[month].steuerfreie_ausland_ausgaben += gezahlt;
    } else {
        // Steuerpflichtige Ausgaben
        data[month].steuerpflichtige_ausgaben += gezahlt;

        if (roundedRate === 7 || roundedRate === 19) {
            data[month][`vst_${roundedRate}`] += tax;
        }
    }
}

/**
 * Aggregiert UStVA-Daten für einen Zeitraum (z.B. Quartal oder Jahr)
 *
 * @param {Object} data - UStVA-Datenobjekt nach Monaten
 * @param {number} start - Startmonat (1-12)
 * @param {number} end - Endmonat (1-12)
 * @returns {Object} Aggregiertes UStVA-Datenobjekt
 */
function aggregateUStVA(data, start, end) {
    const sum = dataModel.createEmptyUStVA();

    for (let m = start; m <= end; m++) {
        if (!data[m]) continue; // Überspringe, falls keine Daten für den Monat

        for (const key in sum) {
            sum[key] += data[m][key] || 0;
        }
    }

    return sum;
}

export default {
    processUStVARow,
    aggregateUStVA
};