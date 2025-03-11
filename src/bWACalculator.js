// file: src/bWACalculator.js
import Helpers from "./helpers.js";
import config from "./config.js";
import Validator from "./validator.js";

/**
 * Modul zur Berechnung der Betriebswirtschaftlichen Auswertung (BWA)
 */
const BWACalculator = (() => {
    /**
     * Erstellt ein leeres BWA-Datenobjekt mit Nullwerten
     * @returns {Object} Leere BWA-Datenstruktur
     */
    const createEmptyBWA = () => ({
        // Gruppe 1: Betriebserlöse (Einnahmen)
        umsatzerloese: 0,
        provisionserloese: 0,
        steuerfreieInlandEinnahmen: 0,
        steuerfreieAuslandEinnahmen: 0,
        sonstigeErtraege: 0,
        vermietung: 0,
        zuschuesse: 0,
        waehrungsgewinne: 0,
        anlagenabgaenge: 0,
        gesamtErloese: 0,

        // Gruppe 2: Materialaufwand & Wareneinsatz
        wareneinsatz: 0,
        fremdleistungen: 0,
        rohHilfsBetriebsstoffe: 0,
        gesamtWareneinsatz: 0,

        // Gruppe 3: Betriebsausgaben (Sachkosten)
        bruttoLoehne: 0,
        sozialeAbgaben: 0,
        sonstigePersonalkosten: 0,
        werbungMarketing: 0,
        reisekosten: 0,
        versicherungen: 0,
        telefonInternet: 0,
        buerokosten: 0,
        fortbildungskosten: 0,
        kfzKosten: 0,
        sonstigeAufwendungen: 0,
        gesamtBetriebsausgaben: 0,

        // Gruppe 4: Abschreibungen & Zinsen
        abschreibungenMaschinen: 0,
        abschreibungenBueromaterial: 0,
        abschreibungenImmateriell: 0,
        zinsenBank: 0,
        zinsenGesellschafter: 0,
        leasingkosten: 0,
        gesamtAbschreibungenZinsen: 0,

        // Gruppe 5: Besondere Posten (Kapitalbewegungen)
        eigenkapitalveraenderungen: 0,
        gesellschafterdarlehen: 0,
        ausschuettungen: 0,
        gesamtBesonderePosten: 0,

        // Gruppe 6: Rückstellungen
        steuerrueckstellungen: 0,
        rueckstellungenSonstige: 0,
        gesamtRueckstellungenTransfers: 0,

        // Gruppe 7: EBIT
        ebit: 0,

        // Gruppe 8: Steuern & Vorsteuer
        umsatzsteuer: 0,
        vorsteuer: 0,
        nichtAbzugsfaehigeVSt: 0,
        koerperschaftsteuer: 0,
        solidaritaetszuschlag: 0,
        gewerbesteuer: 0,
        gewerbesteuerRueckstellungen: 0,
        sonstigeSteuerrueckstellungen: 0,
        steuerlast: 0,

        // Gruppe 9: Jahresüberschuss/-fehlbetrag
        gewinnNachSteuern: 0,

        // Eigenbelege (zur Aggregation)
        eigenbelegeSteuerfrei: 0,
        eigenbelegeSteuerpflichtig: 0
    });

    /**
     * Verarbeitet Einnahmen und ordnet sie den BWA-Kategorien zu
     *
     * @param {Array} row - Zeile aus dem Einnahmen-Sheet
     * @param {Object} bwaData - BWA-Datenstruktur
     */
    const processRevenue = (row, bwaData) => {
        try {
            const m = Helpers.getMonthFromRow(row);
            if (!m) return;

            const amount = Helpers.parseCurrency(row[4]);
            const category = row[2]?.toString().trim() || "";

            if (["Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) return;

            // Spezielle Kategorien direkt zuordnen
            if (category === "Sonstige betriebliche Erträge") return void (bwaData[m].sonstigeErtraege += amount);
            if (category === "Erträge aus Vermietung/Verpachtung") return void (bwaData[m].vermietung += amount);
            if (category === "Erträge aus Zuschüssen") return void (bwaData[m].zuschuesse += amount);
            if (category === "Erträge aus Währungsgewinnen") return void (bwaData[m].waehrungsgewinne += amount);
            if (category === "Erträge aus Anlagenabgängen") return void (bwaData[m].anlagenabgaenge += amount);

            // BWA-Mapping aus Konfiguration verwenden
            const mapping = config.einnahmen.bwaMapping[category];
            if (["umsatzerloese", "provisionserloese"].includes(mapping)) {
                bwaData[m][mapping] += amount;
            } else if (Helpers.parseMwstRate(row[5]) === 0) {
                bwaData[m].steuerfreieInlandEinnahmen += amount;
            } else {
                bwaData[m].umsatzerloese += amount;
            }
        } catch (e) {
            console.error("Fehler bei der Verarbeitung einer Einnahme:", e);
        }
    };

    /**
     * Verarbeitet Ausgaben und ordnet sie den BWA-Kategorien zu
     *
     * @param {Array} row - Zeile aus dem Ausgaben-Sheet
     * @param {Object} bwaData - BWA-Datenstruktur
     */
    const processExpense = (row, bwaData) => {
        try {
            const m = Helpers.getMonthFromRow(row);
            if (!m) return;

            const amount = Helpers.parseCurrency(row[4]);
            const category = row[2]?.toString().trim() || "";

            // Nicht-betriebliche Positionen ignorieren
            if (["Privatentnahme", "Privateinlage", "Holding Transfers",
                "Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) return;

            // Spezielle Kategorien direkt zuordnen
            if (category === "Bruttolöhne & Gehälter") return void (bwaData[m].bruttoLoehne += amount);
            if (category === "Soziale Abgaben & Arbeitgeberanteile") return void (bwaData[m].sozialeAbgaben += amount);
            if (category === "Sonstige Personalkosten") return void (bwaData[m].sonstigePersonalkosten += amount);
            if (category === "Gewerbesteuerrückstellungen") return void (bwaData[m].gewerbesteuerRueckstellungen += amount);
            if (category === "Telefon & Internet") return void (bwaData[m].telefonInternet += amount);
            if (category === "Bürokosten") return void (bwaData[m].buerokosten += amount);
            if (category === "Fortbildungskosten") return void (bwaData[m].fortbildungskosten += amount);

            // BWA-Mapping aus Konfiguration verwenden
            const mapping = config.ausgaben.bwaMapping[category];
            switch (mapping) {
                case "wareneinsatz":
                    bwaData[m].wareneinsatz += amount;
                    break;
                case "fremdleistungen":
                    bwaData[m].fremdleistungen += amount;
                    break;
                case "rohHilfsBetriebsstoffe":
                    bwaData[m].rohHilfsBetriebsstoffe += amount;
                    break;
                case "werbungMarketing":
                    bwaData[m].werbungMarketing += amount;
                    break;
                case "reisekosten":
                    bwaData[m].reisekosten += amount;
                    break;
                case "versicherungen":
                    bwaData[m].versicherungen += amount;
                    break;
                case "kfzKosten":
                    bwaData[m].kfzKosten += amount;
                    break;
                case "abschreibungenMaschinen":
                    bwaData[m].abschreibungenMaschinen += amount;
                    break;
                case "abschreibungenBueromaterial":
                    bwaData[m].abschreibungenBueromaterial += amount;
                    break;
                case "abschreibungenImmateriell":
                    bwaData[m].abschreibungenImmateriell += amount;
                    break;
                case "zinsenBank":
                    bwaData[m].zinsenBank += amount;
                    break;
                case "zinsenGesellschafter":
                    bwaData[m].zinsenGesellschafter += amount;
                    break;
                case "leasingkosten":
                    bwaData[m].leasingkosten += amount;
                    break;
                case "koerperschaftsteuer":
                    bwaData[m].koerperschaftsteuer += amount;
                    break;
                case "solidaritaetszuschlag":
                    bwaData[m].solidaritaetszuschlag += amount;
                    break;
                case "gewerbesteuer":
                    bwaData[m].gewerbesteuer += amount;
                    break;
                default:
                    bwaData[m].sonstigeAufwendungen += amount;
            }
        } catch (e) {
            console.error("Fehler bei der Verarbeitung einer Ausgabe:", e);
        }
    };

    /**
     * Verarbeitet Eigenbelege und ordnet sie den BWA-Kategorien zu
     *
     * @param {Array} row - Zeile aus dem Eigenbelege-Sheet
     * @param {Object} bwaData - BWA-Datenstruktur
     */
    const processEigen = (row, bwaData) => {
        try {
            const m = Helpers.getMonthFromRow(row);
            if (!m) return;

            const amount = Helpers.parseCurrency(row[4]);
            const category = row[2]?.toString().trim() || "";
            const eigenCfg = config.eigenbelege.mapping[category] ?? {};
            const taxType = eigenCfg.taxType ?? "steuerpflichtig";

            if (taxType === "steuerfrei") {
                bwaData[m].eigenbelegeSteuerfrei += amount;
            } else {
                bwaData[m].eigenbelegeSteuerpflichtig += amount;
            }
        } catch (e) {
            console.error("Fehler bei der Verarbeitung eines Eigenbelegs:", e);
        }
    };

    /**
     * Sammelt alle BWA-Daten aus den verschiedenen Sheets
     *
     * @returns {Object|null} BWA-Datenstruktur oder null bei Fehler
     */
    const aggregateBWAData = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const revenueSheet = ss.getSheetByName("Einnahmen");
            const expenseSheet = ss.getSheetByName("Ausgaben");
            const eigenSheet = ss.getSheetByName("Eigenbelege");

            if (!revenueSheet || !expenseSheet) {
                SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
                return null;
            }

            // BWA-Daten für alle Monate initialisieren
            const bwaData = Object.fromEntries(Array.from({length: 12}, (_, i) => [i + 1, createEmptyBWA()]));

            // Daten aus den Sheets verarbeiten
            revenueSheet.getDataRange().getValues().slice(1).forEach(processRevenue);
            expenseSheet.getDataRange().getValues().slice(1).forEach(processExpense);
            if (eigenSheet) eigenSheet.getDataRange().getValues().slice(1).forEach(processEigen);

            // Gruppensummen und weitere Berechnungen
            for (let m = 1; m <= 12; m++) {
                const d = bwaData[m];

                // Erlöse
                d.gesamtErloese = d.umsatzerloese + d.provisionserloese + d.steuerfreieInlandEinnahmen +
                    d.steuerfreieAuslandEinnahmen + d.sonstigeErtraege + d.vermietung +
                    d.zuschuesse + d.waehrungsgewinne + d.anlagenabgaenge;

                // Materialkosten
                d.gesamtWareneinsatz = d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe;

                // Betriebsausgaben
                d.gesamtBetriebsausgaben = d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten +
                    d.werbungMarketing + d.reisekosten + d.versicherungen + d.telefonInternet +
                    d.buerokosten + d.fortbildungskosten + d.kfzKosten + d.sonstigeAufwendungen;

                // Abschreibungen & Zinsen
                d.gesamtAbschreibungenZinsen = d.abschreibungenMaschinen + d.abschreibungenBueromaterial +
                    d.abschreibungenImmateriell + d.zinsenBank + d.zinsenGesellschafter +
                    d.leasingkosten;

                // Besondere Posten
                d.gesamtBesonderePosten = d.eigenkapitalveraenderungen + d.gesellschafterdarlehen + d.ausschuettungen;

                // Rückstellungen
                d.gesamtRueckstellungenTransfers = d.steuerrueckstellungen + d.rueckstellungenSonstige;

                // EBIT
                d.ebit = d.gesamtErloese - (d.gesamtWareneinsatz + d.gesamtBetriebsausgaben +
                    d.gesamtAbschreibungenZinsen + d.gesamtBesonderePosten);

                // Steuern berechnen
                const taxConfig = config.tax.isHolding ? config.tax.holding : config.tax.operative;

                // Für Holdings gelten spezielle Steuersätze wegen Beteiligungsprivileg
                const steuerfaktor = config.tax.isHolding
                    ? taxConfig.gewinnUebertragSteuerpflichtig / 100
                    : 1;

                d.gewerbesteuer = d.ebit * (taxConfig.gewerbesteuer / 100) * steuerfaktor;
                d.koerperschaftsteuer = d.ebit * (taxConfig.koerperschaftsteuer / 100) * steuerfaktor;
                d.solidaritaetszuschlag = d.koerperschaftsteuer * (taxConfig.solidaritaetszuschlag / 100);

                // Gesamte Steuerlast
                d.steuerlast = d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer;

                // Gewinn nach Steuern
                d.gewinnNachSteuern = d.ebit - d.steuerlast;
            }

            return bwaData;
        } catch (e) {
            console.error("Fehler bei der Aggregation der BWA-Daten:", e);
            SpreadsheetApp.getUi().alert("Fehler bei der BWA-Berechnung: " + e.toString());
            return null;
        }
    };

    /**
     * Erstellt den Header für die BWA mit Monats- und Quartalsspalten
     * @returns {Array} Header-Zeile
     */
    const buildHeaderRow = () => {
        const headers = ["Kategorie"];
        for (let q = 0; q < 4; q++) {
            for (let m = q * 3; m < q * 3 + 3; m++) {
                headers.push(`${config.common.months[m]} (€)`);
            }
            headers.push(`Q${q + 1} (€)`);
        }
        headers.push("Jahr (€)");
        return headers;
    };

    /**
     * Erstellt eine Ausgabezeile für eine Position
     * @param {Object} pos - Position mit Label und Wert-Funktion
     * @param {Object} bwaData - BWA-Daten
     * @returns {Array} Formatierte Zeile
     */
    const buildOutputRow = (pos, bwaData) => {
        const monthly = [];
        let yearly = 0;

        // Monatswerte berechnen
        for (let m = 1; m <= 12; m++) {
            const val = pos.get(bwaData[m]) || 0;
            monthly.push(val);
            yearly += val;
        }

        // Quartalswerte berechnen
        const quarters = [0, 0, 0, 0];
        for (let i = 0; i < 12; i++) {
            quarters[Math.floor(i / 3)] += monthly[i];
        }

        // Zeile zusammenstellen
        return [pos.label,
            ...monthly.slice(0, 3), quarters[0],
            ...monthly.slice(3, 6), quarters[1],
            ...monthly.slice(6, 9), quarters[2],
            ...monthly.slice(9, 12), quarters[3],
            yearly];
    };

    /**
     * Hauptfunktion zur Berechnung der BWA
     */
    const calculateBWA = () => {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const bwaData = aggregateBWAData();
            if (!bwaData) return;

            // Positionen definieren
            const positions = [
                {label: "Erlöse aus Lieferungen und Leistungen", get: d => d.umsatzerloese || 0},
                {label: "Provisionserlöse", get: d => d.provisionserloese || 0},
                {label: "Steuerfreie Inland-Einnahmen", get: d => d.steuerfreieInlandEinnahmen || 0},
                {label: "Steuerfreie Ausland-Einnahmen", get: d => d.steuerfreieAuslandEinnahmen || 0},
                {label: "Sonstige betriebliche Erträge", get: d => d.sonstigeErtraege || 0},
                {label: "Erträge aus Vermietung/Verpachtung", get: d => d.vermietung || 0},
                {label: "Erträge aus Zuschüssen", get: d => d.zuschuesse || 0},
                {label: "Erträge aus Währungsgewinnen", get: d => d.waehrungsgewinne || 0},
                {label: "Erträge aus Anlagenabgängen", get: d => d.anlagenabgaenge || 0},
                {label: "Betriebserlöse", get: d => d.gesamtErloese || 0},
                {label: "Wareneinsatz", get: d => d.wareneinsatz || 0},
                {label: "Bezogene Leistungen", get: d => d.fremdleistungen || 0},
                {label: "Roh-, Hilfs- & Betriebsstoffe", get: d => d.rohHilfsBetriebsstoffe || 0},
                {label: "Gesamtkosten Material & Fremdleistungen", get: d => d.gesamtWareneinsatz || 0},
                {label: "Bruttolöhne & Gehälter", get: d => d.bruttoLoehne || 0},
                {label: "Soziale Abgaben & Arbeitgeberanteile", get: d => d.sozialeAbgaben || 0},
                {label: "Sonstige Personalkosten", get: d => d.sonstigePersonalkosten || 0},
                {label: "Werbung & Marketing", get: d => d.werbungMarketing || 0},
                {label: "Reisekosten", get: d => d.reisekosten || 0},
                {label: "Versicherungen", get: d => d.versicherungen || 0},
                {label: "Telefon & Internet", get: d => d.telefonInternet || 0},
                {label: "Bürokosten", get: d => d.buerokosten || 0},
                {label: "Fortbildungskosten", get: d => d.fortbildungskosten || 0},
                {label: "Kfz-Kosten", get: d => d.kfzKosten || 0},
                {label: "Sonstige betriebliche Aufwendungen", get: d => d.sonstigeAufwendungen || 0},
                {label: "Abschreibungen Maschinen", get: d => d.abschreibungenMaschinen || 0},
                {label: "Abschreibungen Büroausstattung", get: d => d.abschreibungenBueromaterial || 0},
                {label: "Abschreibungen immaterielle Wirtschaftsgüter", get: d => d.abschreibungenImmateriell || 0},
                {label: "Zinsen auf Bankdarlehen", get: d => d.zinsenBank || 0},
                {label: "Zinsen auf Gesellschafterdarlehen", get: d => d.zinsenGesellschafter || 0},
                {label: "Leasingkosten", get: d => d.leasingkosten || 0},
                {label: "Gesamt Abschreibungen & Zinsen", get: d => d.gesamtAbschreibungenZinsen || 0},
                {label: "Eigenkapitalveränderungen", get: d => d.eigenkapitalveraenderungen || 0},
                {label: "Gesellschafterdarlehen", get: d => d.gesellschafterdarlehen || 0},
                {label: "Ausschüttungen an Gesellschafter", get: d => d.ausschuettungen || 0},
                {label: "Steuerrückstellungen", get: d => d.steuerrueckstellungen || 0},
                {label: "Rückstellungen sonstige", get: d => d.rueckstellungenSonstige || 0},
                {label: "Betriebsergebnis vor Steuern (EBIT)", get: d => d.ebit || 0},
                {label: "Umsatzsteuer (abzuführen)", get: d => d.umsatzsteuer || 0},
                {label: "Vorsteuer", get: d => d.vorsteuer || 0},
                {label: "Nicht abzugsfähige VSt (Bewirtung)", get: d => d.nichtAbzugsfaehigeVSt || 0},
                {label: "Körperschaftsteuer", get: d => d.koerperschaftsteuer || 0},
                {label: "Solidaritätszuschlag", get: d => d.solidaritaetszuschlag || 0},
                {label: "Gewerbesteuer", get: d => d.gewerbesteuer || 0},
                {label: "Gesamtsteueraufwand", get: d => d.steuerlast || 0},
                {label: "Jahresüberschuss/-fehlbetrag", get: d => d.gewinnNachSteuern || 0}
            ];

            // Header-Zeile erstellen
            const headerRow = buildHeaderRow();
            const outputRows = [headerRow];

            // Gruppenhierarchie für BWA
            const bwaGruppen = [
                {titel: "Betriebserlöse (Einnahmen)", count: 10},
                {titel: "Materialaufwand & Wareneinsatz", count: 4},
                {titel: "Betriebsausgaben (Sachkosten)", count: 11},
                {titel: "Abschreibungen & Zinsen", count: 7},
                {titel: "Besondere Posten", count: 3},
                {titel: "Rückstellungen", count: 2},
                {titel: "Betriebsergebnis vor Steuern (EBIT)", count: 1},
                {titel: "Steuern & Vorsteuer", count: 7},
                {titel: "Jahresüberschuss/-fehlbetrag", count: 1}
            ];

            // Ausgabe mit Gruppenhierarchie erstellen
            let posIndex = 0;
            for (let gruppenIndex = 0; gruppenIndex < bwaGruppen.length; gruppenIndex++) {
                const gruppe = bwaGruppen[gruppenIndex];

                // Gruppenüberschrift
                outputRows.push([
                    `${gruppenIndex + 1}. ${gruppe.titel}`,
                    ...Array(headerRow.length - 1).fill("")
                ]);

                // Gruppenpositionen
                for (let i = 0; i < gruppe.count; i++) {
                    outputRows.push(buildOutputRow(positions[posIndex++], bwaData));
                }

                // Leerzeile nach jeder Gruppe außer der letzten
                if (gruppenIndex < bwaGruppen.length - 1) {
                    outputRows.push(Array(headerRow.length).fill(""));
                }
            }

            // BWA-Sheet erstellen oder aktualisieren
            const bwaSheet = ss.getSheetByName("BWA") || ss.insertSheet("BWA");
            bwaSheet.clearContents();

            // Daten in das Sheet schreiben
            const dataRange = bwaSheet.getRange(1, 1, outputRows.length, outputRows[0].length);
            dataRange.setValues(outputRows);

            // Formatierungen anwenden
            // Header formatieren
            bwaSheet.getRange(1, 1, 1, headerRow.length).setFontWeight("bold").setBackground("#f3f3f3");

            // Gruppenüberschriften formatieren
            for (let i = 0, rowIndex = 2; i < bwaGruppen.length; i++) {
                bwaSheet.getRange(rowIndex, 1).setFontWeight("bold");
                rowIndex += bwaGruppen[i].count + 1; // +1 für die Leerzeile
            }

            // Währungsformat für alle Zahlenwerte
            bwaSheet.getRange(2, 2, outputRows.length - 1, headerRow.length - 1).setNumberFormat("#,##0.00 €");

            // Summen-Zeilen hervorheben
            const summenZeilen = [11, 15, 26, 33, 36, 38, 39, 46];
            summenZeilen.forEach(row => {
                bwaSheet.getRange(row, 1, 1, headerRow.length).setBackground("#e6f2ff");
            });

            // EBIT und Jahresüberschuss hervorheben
            bwaSheet.getRange(39, 1, 1, headerRow.length).setFontWeight("bold");
            bwaSheet.getRange(46, 1, 1, headerRow.length).setFontWeight("bold");

            // Spaltenbreiten anpassen
            bwaSheet.autoResizeColumns(1, headerRow.length);

            // Erfolgsbenachrichtigung
            SpreadsheetApp.getUi().alert("BWA wurde aktualisiert!");

            // BWA-Sheet aktivieren
            ss.setActiveSheet(bwaSheet);

        } catch (e) {
            console.error("Fehler bei der BWA-Berechnung:", e);
            SpreadsheetApp.getUi().alert("Fehler bei der BWA-Berechnung: " + e.toString());
        }
    };

    // Öffentliche API des Moduls
    return {calculateBWA};
})();

export default BWACalculator;