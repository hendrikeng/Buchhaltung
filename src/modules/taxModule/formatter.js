// src/modules/taxModule/formatter.js
import numberUtils from '../../utils/numberUtils.js';
import sheetUtils from '../../utils/sheetUtils.js';

/**
 * Modul zur Formatierung und Ausgabe von Steuerberechnungen
 */
const formatter = {
    /**
     * Erstellt einen Steuerberichts als Sheet mit formatierten Tabellen
     * @param {Spreadsheet} ss - Das Spreadsheet
     * @param {Object} taxData - Die berechneten Steuerdaten
     * @param {Object} bilanzData - Die Bilanzdaten
     * @param {Object} config - Die Konfiguration
     * @returns {boolean} - true bei Erfolg, false bei Fehler
     */
    generateTaxReport(ss, taxData, bilanzData, config) {
        try {
            console.log('Generating tax report...');
            const gewinnVorSteuern = bilanzData?.passiva?.jahresueberschuss || 0;
            const year = config.tax.year || new Date().getFullYear();
            const isHolding = config.tax.isHolding;

            // Steuerberichts-Sheet erstellen oder aktualisieren
            const taxSheet = sheetUtils.getOrCreateSheet(ss, 'Steuerberechnung');
            taxSheet.clearContents();

            // Arrays für Hauptabschnitte des Berichts
            const headerSection = [
                [`Steuerberechnung ${year}`, ''],
                [`${isHolding ? 'Holding GmbH' : 'Operative GmbH'}`, ''],
                ['', ''],
            ];

            // Gewinn- und Verlustrechnung (als Übersicht)
            const gewinnSection = [
                ['Gewinn- und Verlustrechnung (Überblick)', ''],
                ['Gewinn vor Steuern', numberUtils.formatCurrency(gewinnVorSteuern)],
                ['Körperschaftsteuer', numberUtils.formatCurrency(taxData.koerperschaftsteuer || 0)],
                ['Solidaritätszuschlag', numberUtils.formatCurrency(taxData.solidaritaetszuschlag || 0)],
            ];

            if (!isHolding) {
                // Gewerbesteuer nur für operative GmbH
                gewinnSection.push(['Gewerbesteuer', numberUtils.formatCurrency(taxData.gewerbesteuer || 0)]);
            } else {
                // Beteiligungserträge nur für Holding
                if (taxData.dividendensteuer) {
                    gewinnSection.push(['Steuer auf Beteiligungserträge (5%)',
                        numberUtils.formatCurrency(taxData.dividendensteuer.gesamtsteuer || 0)]);
                }
            }

            gewinnSection.push(
                ['Gesamtsteuerbelastung', numberUtils.formatCurrency(taxData.gesamtsteuerbelastung || 0)],
                ['Effektiver Steuersatz', `${numberUtils.round(taxData.effektiverSteuersatz || 0, 2)}%`],
                ['Gewinn nach Steuern', numberUtils.formatCurrency(gewinnVorSteuern - (taxData.gesamtsteuerbelastung || 0))],
                ['', ''],
            );

            // Abschnitt für Körperschaftsteuer
            const kstSection = this._formatKoerperschaftsteuerSection(taxData);

            // Abschnitt für Gewerbesteuer (nur für operative GmbH)
            const gewstSection = !isHolding && taxData.gewerbesteuer ?
                this._formatGewerbesteuerSection(taxData) : [];

            // Abschnitt für Dividendensteuer (nur für Holding)
            const divSection = isHolding && taxData.dividendensteuer ?
                this._formatDividendensteuerSection(taxData.dividendensteuer) : [];

            // Abschnitt für Optimierungsszenarien
            const optSection = taxData.optimierung ?
                this._formatOptimierungSection(taxData.optimierung, config) : [];

            // Alle Abschnitte zusammenführen
            const allSections = [
                ...headerSection,
                ...gewinnSection,
                ...kstSection,
                ...gewstSection,
                ...divSection,
                ...optSection,
            ];

            // Daten in einem Batch-API-Call schreiben
            taxSheet.getRange(1, 1, allSections.length, 2).setValues(allSections);

            // Formatierungen in Batches anwenden
            this._applyTaxReportFormatting(taxSheet, allSections.length);

            // Sheet aktivieren
            ss.setActiveSheet(taxSheet);
            console.log('Tax report generated successfully');

            return true;
        } catch (e) {
            console.error('Fehler bei der Generierung des Steuerberichts:', e);
            return false;
        }
    },

    /**
     * Formatiert den Abschnitt für Körperschaftsteuer
     * @param {Object} taxData - Die Steuerdaten
     * @returns {Array} - Formatierter Abschnitt als 2D-Array
     */
    _formatKoerperschaftsteuerSection(taxData) {
        const kstDetails = taxData.koerperschaftsteuerDetails || {};

        return [
            ['Körperschaftsteuer-Berechnung', ''],
            ['Bemessungsgrundlage', numberUtils.formatCurrency(kstDetails.bemessungsgrundlage || 0)],
            ['Körperschaftsteuersatz', `${(kstDetails.kstSatz || 0) * 100}%`],
            ['Körperschaftsteuer', numberUtils.formatCurrency(taxData.koerperschaftsteuer || 0)],
            ['Solidaritätszuschlag', numberUtils.formatCurrency(taxData.solidaritaetszuschlag || 0)],
            ['Gesamte Körperschaftsteuerbelastung',
                numberUtils.formatCurrency((taxData.koerperschaftsteuer || 0) + (taxData.solidaritaetszuschlag || 0))],
            ['', ''],
        ];
    },

    /**
     * Formatiert den Abschnitt für Gewerbesteuer
     * @param {Object} taxData - Die Steuerdaten
     * @returns {Array} - Formatierter Abschnitt als 2D-Array
     */
    _formatGewerbesteuerSection(taxData) {
        const gewstDetails = taxData.gewerbesteuerDetails || {};
        const gewinnVorSteuern = gewstDetails.gewinnVorSteuern || 0;

        const section = [
            ['Gewerbesteuer-Berechnung', ''],
            ['Gewinn vor Steuern', numberUtils.formatCurrency(gewinnVorSteuern)],
        ];

        // Hinzurechnungen
        if (gewstDetails.hinzurechnungen) {
            section.push(['Hinzurechnungen:', '']);
            if (gewstDetails.hinzurechnungen.miete) {
                section.push(['- 25% der Miete', numberUtils.formatCurrency(gewstDetails.hinzurechnungen.miete)]);
            }
            if (gewstDetails.hinzurechnungen.zinsen) {
                section.push(['- 25% der Zinsen', numberUtils.formatCurrency(gewstDetails.hinzurechnungen.zinsen)]);
            }
            if (gewstDetails.hinzurechnungen.leasing) {
                section.push(['- 20% der Leasingkosten', numberUtils.formatCurrency(gewstDetails.hinzurechnungen.leasing)]);
            }
            section.push(['Summe Hinzurechnungen', numberUtils.formatCurrency(gewstDetails.hinzurechnungen.summe)]);
        }

        // Kürzungen
        if (gewstDetails.kuerzungen) {
            section.push(['Kürzungen:', '']);
            if (gewstDetails.kuerzungen.grundbesitz) {
                section.push(['- 1,2% des Einheitswerts bei Grundbesitz',
                    numberUtils.formatCurrency(gewstDetails.kuerzungen.grundbesitz)]);
            }
            section.push(['Summe Kürzungen', numberUtils.formatCurrency(gewstDetails.kuerzungen.summe)]);
        }

        // Weitere Berechnung
        section.push(
            ['Gewerbeertrag', numberUtils.formatCurrency(gewstDetails.gewerbeertrag || 0)],
            ['Freibetrag', numberUtils.formatCurrency(gewstDetails.freibetrag || 0)],
            ['Gewerbeertrag nach Freibetrag', numberUtils.formatCurrency(gewstDetails.gewerbeertragNachFreibetrag || 0)],
            ['Steuermesszahl', `${(gewstDetails.steuermesszahl || 0) * 100}%`],
            ['Steuermessbetrag', numberUtils.formatCurrency(gewstDetails.messbetrag || 0)],
            ['Hebesatz', `${(gewstDetails.hebesatz || 0) * 100}%`],
            ['Gewerbesteuer', numberUtils.formatCurrency(taxData.gewerbesteuer || 0)],
            ['', ''],
        );

        return section;
    },

    /**
     * Formatiert den Abschnitt für Dividendensteuer (Holding)
     * @param {Object} dividendenData - Die Dividendensteuerdaten
     * @returns {Array} - Formatierter Abschnitt als 2D-Array
     */
    _formatDividendensteuerSection(dividendenData) {
        const section = [
            ['Besteuerung von Beteiligungserträgen (Holding)', ''],
            ['Beteiligungserträge insgesamt', numberUtils.formatCurrency(dividendenData.beteiligungsertraege || 0)],
            ['Steuerfreier Anteil (95%)', numberUtils.formatCurrency(dividendenData.steuerfreierAnteil || 0)],
            ['Steuerpflichtiger Anteil (5%)', numberUtils.formatCurrency(dividendenData.steuerpflichtigerAnteil || 0)],
            ['Körperschaftsteuer auf 5% (15%)', numberUtils.formatCurrency(dividendenData.koerperschaftsteuer || 0)],
            ['Solidaritätszuschlag (5,5%)', numberUtils.formatCurrency(dividendenData.solidaritaetszuschlag || 0)],
            ['Gesamtsteuer auf Beteiligungserträge', numberUtils.formatCurrency(dividendenData.gesamtsteuer || 0)],
            ['Effektiver Steuersatz auf Beteiligungserträge',
                `${numberUtils.round(dividendenData.effektiverSteuersatz || 0, 2)}%`],
            ['', ''],
        ];

        return section;
    },

    /**
     * Formatiert den Abschnitt für Optimierungsszenarien
     * @param {Object} optimierungData - Die Optimierungsdaten
     * @param {Object} config - Die Konfiguration
     * @returns {Array} - Formatierter Abschnitt als 2D-Array
     */
    _formatOptimierungSection(optimierungData, config) {
        const section = [
            ['Steueroptimierung - Handlungsempfehlungen', ''],
        ];

        // Thesaurierung vs. Ausschüttung
        if (optimierungData.thesaurierungVsAusschuettung) {
            section.push(
                ['1. Thesaurierung vs. Ausschüttung', ''],
                ['Bei Thesaurierung (Gewinn in GmbH belassen):', ''],
                ['Unternehmenssteuern',
                    numberUtils.formatCurrency(optimierungData.thesaurierungVsAusschuettung.thesaurierung.unternehmenssteuern)],
                ['Effektiver Steuersatz',
                    `${numberUtils.round(optimierungData.thesaurierungVsAusschuettung.thesaurierung.effektiverSteuersatz, 2)}%`],
                ['Bei Ausschüttung (mit 42% ESt-Satz):', ''],
                ['Gesamtsteuerbelastung',
                    numberUtils.formatCurrency(optimierungData.thesaurierungVsAusschuettung.ausschuettungsSzenarien[2].gesamtsteuer)],
                ['Effektiver Steuersatz',
                    `${numberUtils.round(optimierungData.thesaurierungVsAusschuettung.ausschuettungsSzenarien[2].effektiverSteuersatz, 2)}%`],
                ['Empfehlung:', optimierungData.thesaurierungVsAusschuettung.empfehlung || ''],
                ['', ''],
            );
        }

        // Holding vs. direkte Ausschüttung
        if (optimierungData.holdingVsDirekt) {
            section.push(
                ['2. Holding-Struktur vs. direkte Ausschüttung', ''],
                ['Bei direkter Ausschüttung aus operativer GmbH:', ''],
                ['Gesamtsteuerbelastung',
                    numberUtils.formatCurrency(optimierungData.holdingVsDirekt.direkteAusschuettung.gesamtsteuer)],
                ['Effektiver Steuersatz',
                    `${numberUtils.round(optimierungData.holdingVsDirekt.direkteAusschuettung.effektiverSteuersatz, 2)}%`],
                ['Bei Ausschüttung über Holding-Struktur:', ''],
                ['Gesamtsteuerbelastung',
                    numberUtils.formatCurrency(optimierungData.holdingVsDirekt.ausschuettungUeberHolding.gesamtsteuer)],
                ['Effektiver Steuersatz',
                    `${numberUtils.round(optimierungData.holdingVsDirekt.ausschuettungUeberHolding.effektiverSteuersatz, 2)}%`],
                ['Steuerersparnis durch Holding',
                    numberUtils.formatCurrency(optimierungData.holdingVsDirekt.steuerersparnis)],
                ['Empfehlung:', optimierungData.holdingVsDirekt.empfehlung || ''],
                ['', ''],
            );
        }

        // Gesamtempfehlung
        if (optimierungData.simulation && optimierungData.simulation.empfehlung) {
            section.push(
                ['Gesamtempfehlung zur Steueroptimierung:', ''],
                [optimierungData.simulation.empfehlung, ''],
                ['', ''],
            );
        }

        return section;
    },

    /**
     * Wendet Formatierungen auf den Steuerbericht an
     * @param {Sheet} sheet - Das Sheet
     * @param {number} rowCount - Anzahl der Zeilen
     */
    _applyTaxReportFormatting(sheet, rowCount) {
        // Optimierung: Formatierungen in Batches anwenden

        // 1. Überschriften und Abschnittstitel formatieren
        const headerRanges = [
            sheet.getRange('A1:B1'),  // Haupttitel
            sheet.getRange('A2:B2'),  // Untertitel
            sheet.getRange('A4:B4'),  // G&V Überschrift
            sheet.getRange('A5:A14'), // Linke Spalte (Beschriftungen)
        ];

        // Formatierungsbatches anwenden
        headerRanges.forEach(range => {
            range.setFontWeight('bold');
        });

        // 2. Hintergrundfarben für Abschnitte
        const backgroundRanges = {
            title: { range: sheet.getRange('A1:B2'), color: '#D9EAD3' },   // Hellgrün für Titel
            gv: { range: sheet.getRange('A4:B5'), color: '#F4F7F9' },      // Hellgrau für G&V
            kst: { range: sheet.getRange('A15:B15'), color: '#F4F7F9' },   // Hellgrau für KSt
            gewst: { range: sheet.getRange('A23:B23'), color: '#F4F7F9' }, // Hellgrau für GewSt
            div: { range: sheet.getRange('A35:B35'), color: '#F4F7F9' },   // Hellgrau für Dividenden
            opt: { range: sheet.getRange('A44:B44'), color: '#F4F7F9' },    // Hellgrau für Optimierung
        };

        // Hintergrundfarben anwenden
        Object.values(backgroundRanges).forEach(batch => {
            try {
                batch.range.setBackground(batch.color);
            } catch (e) {
                // Ignoriere Fehler, falls Range nicht existiert
            }
        });

        // 3. Währungsformat für alle Geldbetrag-Zellen
        try {
            sheet.getRange('B5:B' + rowCount).setNumberFormat('#,##0.00 €');
        } catch (e) {
            // Ignoriere Fehler bei leeren Bereichen
        }

        // 4. Spaltenbreiten anpassen
        sheet.setColumnWidth(1, 300); // Breite für Beschreibungen
        sheet.setColumnWidth(2, 150); // Breite für Werte
    },
};

export default formatter;