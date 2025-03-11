// file: src/uStVACalculator.js
import Helpers from "./helpers.js";
import config from "./config.js";
import Validator from "./validator.js";

/**
 * Modul zur Berechnung der Umsatzsteuervoranmeldung (UStVA)
 * Unterstützt die Berechnung nach SKR04 für monatliche und quartalsweise Auswertungen
 */
const UStVACalculator = (() => {
    /**
     * Erstellt ein leeres UStVA-Datenobjekt mit Nullwerten
     * @returns {Object} Leere UStVA-Datenstruktur
     */
    const createEmptyUStVA = () => ({
        steuerpflichtige_einnahmen: 0,
        steuerfreie_inland_einnahmen: 0,
        steuerfreie_ausland_einnahmen: 0,
        steuerpflichtige_ausgaben: 0,
        steuerfreie_inland_ausgaben: 0,
        steuerfreie_ausland_ausgaben: 0,
        eigenbelege_steuerpflichtig: 0,
        eigenbelege_steuerfrei: 0,
        nicht_abzugsfaehige_vst: 0,
        ust_7: 0,
        ust_19: 0,
        vst_7: 0,
        vst_19: 0
    });

    /**
     * Verarbeitet eine Zeile aus den Einnahmen/Ausgaben/Eigenbelegen für die UStVA
     *
     * @param {Array} row - Datenzeile aus einem Sheet
     * @param {Object} data - UStVA-Datenobjekt nach Monaten
     * @param {boolean} isIncome - Handelt es sich um Einnahmen (true) oder Ausgaben (false)
     * @param {boolean} isEigen - Handelt es sich um Eigenbelege (true oder false)
     */
    const processUStVARow = (row, data, isIncome, isEigen = false) => {
        try {
            // Zahlungsdatum prüfen (nur abgeschlossene Zahlungen)
            const paymentDate = Helpers.parseDate(row[13]);
            if (!paymentDate || paymentDate > new Date()) return;

            // Monat und Jahr prüfen (nur relevantes Geschäftsjahr)
            const month = Helpers.getMonthFromRow(row);
            if (!month || month < 1 || month > 12) return;

            // Beträge aus der Zeile extrahieren
            const netto = Helpers.parseCurrency(row[4]);
            const restNetto = Helpers.parseCurrency(row[10]) || 0; // Steuerbemessungsgrundlage für Teilzahlungen
            const gezahlt = netto - restNetto; // Tatsächlich gezahlter/erhaltener Betrag

            // Falls kein Betrag gezahlt wurde, nichts zu verarbeiten
            if (gezahlt === 0) return;

            // MwSt-Satz normalisieren
            const mwstRate = Helpers.parseMwstRate(row[5]);
            const roundedRate = Math.round(mwstRate);

            // Steuer berechnen
            const tax = gezahlt * (mwstRate / 100);

            // Kategorie ermitteln
            const category = row[2]?.toString().trim() || "";

            // Je nach Typ (Einnahme/Ausgabe/Eigenbeleg) unterschiedlich verarbeiten
            if (isIncome) {
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
                    } else {
                        console.warn(`Unbekannter MwSt-Satz für Einnahme: ${roundedRate}%`);
                    }
                }
            } else if (isEigen) {
                // EIGENBELEGE
                const eigenCfg = config.eigenbelege.mapping[category] ?? {};
                const taxType = eigenCfg.taxType ?? "steuerpflichtig";

                if (taxType === "steuerfrei") {
                    // Steuerfreie Eigenbelege
                    data[month].eigenbelege_steuerfrei += gezahlt;
                } else if (taxType === "eigenbeleg" && eigenCfg.besonderheit === "bewirtung") {
                    // Bewirtungsbelege (nur 70% der Vorsteuer absetzbar)
                    data[month].eigenbelege_steuerpflichtig += gezahlt;

                    if (roundedRate === 7 || roundedRate === 19) {
                        data[month][`vst_${roundedRate}`] += tax * 0.7; // 70% absetzbare Vorsteuer
                        data[month].nicht_abzugsfaehige_vst += tax * 0.3; // 30% nicht absetzbar
                    } else {
                        console.warn(`Unbekannter MwSt-Satz für Bewirtung: ${roundedRate}%`);
                    }
                } else {
                    // Normale steuerpflichtige Eigenbelege
                    data[month].eigenbelege_steuerpflichtig += gezahlt;

                    if (roundedRate === 7 || roundedRate === 19) {
                        data[month][`vst_${roundedRate}`] += tax;
                    } else {
                        console.warn(`Unbekannter MwSt-Satz für Eigenbeleg: ${roundedRate}%`);
                    }
                }
            } else {
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
                    } else if (roundedRate !== 0) {
                        // 0% ist kein Fehler, daher nur für andere Sätze warnen
                        console.warn(`Unbekannter MwSt-Satz für Ausgabe: ${roundedRate}%`);
                    }
                }
            }
        } catch (e) {
            console.error("Fehler bei der Verarbeitung einer UStVA-Zeile:", e);
        }
    };

    /**
     * Formatiert eine UStVA-Datenzeile für die Ausgabe
     *
     * @param {string} label - Bezeichnung der Zeile (z.B. Monat oder Quartal)
     * @param {Object} d - UStVA-Datenobjekt für den Zeitraum
     * @returns {Array} Formatierte Zeile für die Ausgabe
     */
    const formatUStVARow = (label, d) => {
        // Berechnung der USt-Zahlung: USt minus VSt (abzüglich nicht abzugsfähiger VSt)
        const ustZahlung = (d.ust_7 + d.ust_19) - ((d.vst_7 + d.vst_19) - d.nicht_abzugsfaehige_vst);

        // Berechnung des Gesamtergebnisses: Einnahmen minus Ausgaben (ohne Steueranteil)
        const ergebnis = (d.steuerpflichtige_einnahmen + d.steuerfreie_inland_einnahmen + d.steuerfreie_ausland_einnahmen) -
            (d.steuerpflichtige_ausgaben + d.steuerfreie_inland_ausgaben + d.steuerfreie_ausland_ausgaben +
                d.eigenbelege_steuerpflichtig + d.eigenbelege_steuerfrei);

        // Formatierte Zeile zurückgeben
        return [
            label,
            d.steuerpflichtige_einnahmen,
            d.steuerfreie_inland_einnahmen,
            d.steuerfreie_ausland_einnahmen,
            d.steuerpflichtige_ausgaben,
            d.steuerfreie_inland_ausgaben,
            d.steuerfreie_ausland_ausgaben,
            d.eigenbelege_steuerpflichtig,
            d.eigenbelege_steuerfrei,
            d.nicht_abzugsfaehige_vst,
            d.ust_7,
            d.ust_19,
            d.vst_7,
            d.vst_19,
            ustZahlung,
            ergebnis
        ];
    };

    /**
     * Aggregiert UStVA-Daten für einen Zeitraum (z.B. Quartal oder Jahr)
     *
     * @param {Object} data - UStVA-Datenobjekt nach Monaten
     * @param {number} start - Startmonat (1-12)
     * @param {number} end - Endmonat (1-12)
     * @returns {Object} Aggregiertes UStVA-Datenobjekt
     */
    const aggregateUStVA = (data, start, end) => {
        const sum = createEmptyUStVA();

        for (let m = start; m <= end; m++) {
            if (!data[m]) continue; // Überspringe, falls keine Daten für den Monat

            for (const key in sum) {
                sum[key] += data[m][key] || 0;
            }
        }

        return sum;
    };

    /**
     * Hauptfunktion zur Berechnung der UStVA
     * Sammelt Daten aus allen relevanten Sheets und erstellt ein UStVA-Sheet
     */
    const calculateUStVA = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const ui = SpreadsheetApp.getUi();

            // Benötigte Sheets abrufen
            const revenueSheet = ss.getSheetByName("Einnahmen");
            const expenseSheet = ss.getSheetByName("Ausgaben");
            const eigenSheet = ss.getSheetByName("Eigenbelege");

            // Prüfen, ob alle benötigten Sheets vorhanden sind
            if (!revenueSheet || !expenseSheet) {
                ui.alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben' wurde nicht gefunden.");
                return;
            }

            // Sheets validieren
            if (!Validator.validateAllSheets(revenueSheet, expenseSheet)) {
                ui.alert("Die UStVA-Berechnung wurde abgebrochen, da Fehler in den Daten gefunden wurden.");
                return;
            }

            // Daten aus den Sheets laden
            const revenueData = revenueSheet.getDataRange().getValues();
            const expenseData = expenseSheet.getDataRange().getValues();
            const eigenData = eigenSheet ? eigenSheet.getDataRange().getValues() : [];

            // Leere UStVA-Datenstruktur für alle Monate erstellen
            const ustvaData = Object.fromEntries(
                Array.from({length: 12}, (_, i) => [i + 1, createEmptyUStVA()])
            );

            // Helfer-Funktion zum Verarbeiten von Datenzeilen
            const processRows = (data, isIncome, isEigen = false) => {
                data.slice(1).forEach(row => { // Ab Zeile 2 (nach Header)
                    processUStVARow(row, ustvaData, isIncome, isEigen);
                });
            };

            // Daten verarbeiten
            processRows(revenueData, true);         // Einnahmen
            processRows(expenseData, false);        // Ausgaben
            if (eigenData.length) {
                processRows(eigenData, false, true); // Eigenbelege
            }

            // Ausgabe-Header erstellen
            const outputRows = [
                [
                    "Zeitraum",
                    "Steuerpflichtige Einnahmen",
                    "Steuerfreie Inland-Einnahmen",
                    "Steuerfreie Ausland-Einnahmen",
                    "Steuerpflichtige Ausgaben",
                    "Steuerfreie Inland-Ausgaben",
                    "Steuerfreie Ausland-Ausgaben",
                    "Eigenbelege steuerpflichtig",
                    "Eigenbelege steuerfrei",
                    "Nicht abzugsfähige VSt (Bewirtung)",
                    "USt 7%",
                    "USt 19%",
                    "VSt 7%",
                    "VSt 19%",
                    "USt-Zahlung",
                    "Ergebnis"
                ]
            ];

            // Monatliche Daten ausgeben
            config.common.months.forEach((name, i) => {
                const month = i + 1;
                outputRows.push(formatUStVARow(name, ustvaData[month]));

                // Nach jedem Quartal eine Zusammenfassung einfügen
                if (month % 3 === 0) {
                    const quartalsNummer = month / 3;
                    const quartalsStart = month - 2;
                    outputRows.push(formatUStVARow(
                        `Quartal ${quartalsNummer}`,
                        aggregateUStVA(ustvaData, quartalsStart, month)
                    ));
                }
            });

            // Jahresergebnis anfügen
            outputRows.push(formatUStVARow("Gesamtjahr", aggregateUStVA(ustvaData, 1, 12)));

            // UStVA-Sheet erstellen oder aktualisieren
            const ustvaSheet = ss.getSheetByName("UStVA") || ss.insertSheet("UStVA");
            ustvaSheet.clearContents();

            // Daten in das Sheet schreiben
            const dataRange = ustvaSheet.getRange(1, 1, outputRows.length, outputRows[0].length);
            dataRange.setValues(outputRows);

            // Header formatieren
            ustvaSheet.getRange(1, 1, 1, outputRows[0].length).setFontWeight("bold");

            // Quartalszellen formatieren
            for (let i = 0; i < 4; i++) {
                const row = 3 * (i + 1) + 1 + i; // Position der Quartalszeile
                ustvaSheet.getRange(row, 1, 1, outputRows[0].length).setBackground("#e6f2ff");
            }

            // Jahreszeile formatieren
            ustvaSheet.getRange(outputRows.length, 1, 1, outputRows[0].length)
                .setBackground("#d9e6f2")
                .setFontWeight("bold");

            // Zahlenformate anwenden
            // Währungsformat für Beträge (Spalten 2-16)
            ustvaSheet.getRange(2, 2, outputRows.length - 1, 15).setNumberFormat("#,##0.00 €");

            // Spaltenbreiten anpassen
            ustvaSheet.autoResizeColumns(1, outputRows[0].length);

            // UStVA-Sheet aktivieren
            ss.setActiveSheet(ustvaSheet);

            ui.alert("UStVA wurde erfolgreich aktualisiert!");
        } catch (e) {
            console.error("Fehler bei der UStVA-Berechnung:", e);
            SpreadsheetApp.getUi().alert("Fehler bei der UStVA-Berechnung: " + e.toString());
        }
    };

    // Öffentliche API des Moduls
    return {
        calculateUStVA,
        // Für Testzwecke könnten hier weitere Funktionen exportiert werden
        _internal: {
            createEmptyUStVA,
            processUStVARow,
            formatUStVARow,
            aggregateUStVA
        }
    };
})();

export default UStVACalculator;