/**
 * bwaModule.js - Modul für die Betriebswirtschaftliche Auswertung (BWA)
 *
 * Erstellt die BWA (Betriebswirtschaftliche Auswertung) basierend auf
 * den Daten aus den Einnahmen, Ausgaben und Eigenbelegen.
 */

const BWAModule = (function() {
    // Private Variablen und Funktionen

    /**
     * Erstellt ein leeres BWA-Objekt mit allen relevanten Positionen
     * @returns {Object} Leeres BWA-Objekt
     */
    function createEmptyBWA() {
        return {
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
        };
    }

    /**
     * Ermittelt den Monat aus einer Zeile basierend auf dem Datum
     * @param {Array} row - Zeile mit Daten
     * @param {number} colIndex - Spaltenindex des Datums
     * @returns {number} Monat (1-12) oder 0 bei Fehler
     */
    function getMonthFromRow(row, colIndex = 0) {
        return HelperModule.getMonthFromRow(row, colIndex);
    }

    /**
     * Sammelt alle BWA-Daten aus den verschiedenen Tabellen
     * @returns {Object} BWA-Daten nach Monaten
     */
    function aggregateBWAData() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        // Sheets laden
        const revenueSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN);
        const expenseSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN);
        const eigenSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.EIGENBELEGE);

        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert(`Fehlendes Blatt: '${CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN}' oder '${CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN}'`);
            return null;
        }

        // BWA-Daten für jeden Monat initialisieren
        const bwaData = Object.fromEntries(
            Array.from({length: 12}, (_, i) => [i + 1, createEmptyBWA()])
        );

        // Einnahmen verarbeiten
        function processRevenue(row) {
            const m = getMonthFromRow(row, CONFIG.SYSTEM.EINNAHMEN_COLS.DATUM);
            if (!m) return;

            // Nur bezahlte Einnahmen berücksichtigen (Ist-Besteuerung)
            if (row[CONFIG.SYSTEM.EINNAHMEN_COLS.ZAHLUNGSSTATUS] !== 'Bezahlt') return;

            const amount = HelperModule.parseCurrency(row[CONFIG.SYSTEM.EINNAHMEN_COLS.NETTO]);
            if (amount === 0) return;

            const category = row[CONFIG.SYSTEM.EINNAHMEN_COLS.KATEGORIE]?.toString().trim() || "";

            // Gewinnvortrag/Verlustvortrag überspringen
            if (["Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) {
                return;
            }

            // Spezielle Kategorien
            if (category === "Sonstige betriebliche Erträge") {
                bwaData[m].sonstigeErtraege += amount;
                return;
            }
            if (category.includes("Vermietung") || category.includes("Verpachtung")) {
                bwaData[m].vermietung += amount;
                return;
            }
            if (category.includes("Zuschüsse") || category.includes("Zuschuss")) {
                bwaData[m].zuschuesse += amount;
                return;
            }
            if (category.includes("Währungsgewinn")) {
                bwaData[m].waehrungsgewinne += amount;
                return;
            }
            if (category.includes("Anlagenabgänge")) {
                bwaData[m].anlagenabgaenge += amount;
                return;
            }

            // BWA-Mapping für Kategorie ermitteln
            // Falls in der Konfiguration hinterlegt
            let mappingKey = "umsatzerloese"; // Standardwert

            if (CONFIG.EINNAHMEN && CONFIG.EINNAHMEN.bwaMapping && CONFIG.EINNAHMEN.bwaMapping[category]) {
                mappingKey = CONFIG.EINNAHMEN.bwaMapping[category];
            } else if (category.includes("Provision")) {
                mappingKey = "provisionserloese";
            } else if (category.includes("steuerfrei") && category.includes("Inland")) {
                mappingKey = "steuerfreieInlandEinnahmen";
            } else if (category.includes("steuerfrei") &&
                (category.includes("Ausland") ||
                    category.includes("Export") ||
                    category.includes("§4 Nr. 1"))) {
                mappingKey = "steuerfreieAuslandEinnahmen";
            }

            // Wert zur entsprechenden BWA-Position hinzufügen
            bwaData[m][mappingKey] += amount;
        }

        // Ausgaben verarbeiten
        function processExpense(row) {
            const m = getMonthFromRow(row, CONFIG.SYSTEM.AUSGABEN_COLS.DATUM);
            if (!m) return;

            // Nur bezahlte Ausgaben berücksichtigen (Ist-Besteuerung)
            if (row[CONFIG.SYSTEM.AUSGABEN_COLS.ZAHLUNGSSTATUS] !== 'Bezahlt') return;

            const amount = HelperModule.parseCurrency(row[CONFIG.SYSTEM.AUSGABEN_COLS.NETTO]);
            if (amount === 0) return;

            const category = row[CONFIG.SYSTEM.AUSGABEN_COLS.KATEGORIE]?.toString().trim() || "";

            // Private Entnahmen/Einlagen überspringen
            if (["Privatentnahme", "Privateinlage", "Holding Transfers",
                "Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) {
                return;
            }

            // Spezielle Kategorien
            if (category.includes("Bruttolöhne") || category.includes("Gehälter")) {
                bwaData[m].bruttoLoehne += amount;
                return;
            }
            if (category.includes("Soziale Abgaben") || category.includes("Arbeitgeberanteil")) {
                bwaData[m].sozialeAbgaben += amount;
                return;
            }
            if (category.includes("Personalkosten") && category.includes("Sonstige")) {
                bwaData[m].sonstigePersonalkosten += amount;
                return;
            }
            if (category.includes("Gewerbesteuerrückstellung")) {
                bwaData[m].gewerbesteuerRueckstellungen += amount;
                return;
            }
            if (category.includes("Telefon") || category.includes("Internet")) {
                bwaData[m].telefonInternet += amount;
                return;
            }
            if (category.includes("Bürokosten")) {
                bwaData[m].buerokosten += amount;
                return;
            }
            if (category.includes("Fortbildung")) {
                bwaData[m].fortbildungskosten += amount;
                return;
            }

            // BWA-Mapping für Kategorie ermitteln
            // Falls in der Konfiguration hinterlegt
            let mappingKey = "sonstigeAufwendungen"; // Standardwert

            if (CONFIG.AUSGABEN && CONFIG.AUSGABEN.bwaMapping && CONFIG.AUSGABEN.bwaMapping[category]) {
                mappingKey = CONFIG.AUSGABEN.bwaMapping[category];
            } else {
                // Verschiedene Kategorien erkennen
                const lowerCategory = category.toLowerCase();

                if (lowerCategory.includes("wareneinsatz")) {
                    mappingKey = "wareneinsatz";
                } else if (lowerCategory.includes("fremdleistung")) {
                    mappingKey = "fremdleistungen";
                } else if (lowerCategory.includes("material") || lowerCategory.includes("betriebsstoff")) {
                    mappingKey = "rohHilfsBetriebsstoffe";
                } else if (lowerCategory.includes("werbung") || lowerCategory.includes("marketing")) {
                    mappingKey = "werbungMarketing";
                } else if (lowerCategory.includes("reise")) {
                    mappingKey = "reisekosten";
                } else if (lowerCategory.includes("versicherung")) {
                    mappingKey = "versicherungen";
                } else if (lowerCategory.includes("kfz") || lowerCategory.includes("fahrzeug")) {
                    mappingKey = "kfzKosten";
                } else if (lowerCategory.includes("abschreibung") && lowerCategory.includes("maschine")) {
                    mappingKey = "abschreibungenMaschinen";
                } else if (lowerCategory.includes("abschreibung") && lowerCategory.includes("büro")) {
                    mappingKey = "abschreibungenBueromaterial";
                } else if (lowerCategory.includes("abschreibung") && lowerCategory.includes("immateriell")) {
                    mappingKey = "abschreibungenImmateriell";
                } else if (lowerCategory.includes("zins") && lowerCategory.includes("bank")) {
                    mappingKey = "zinsenBank";
                } else if (lowerCategory.includes("zins") && lowerCategory.includes("gesellschafter")) {
                    mappingKey = "zinsenGesellschafter";
                } else if (lowerCategory.includes("leasing")) {
                    mappingKey = "leasingkosten";
                }
            }

            // Wert zur entsprechenden BWA-Position hinzufügen
            bwaData[m][mappingKey] += amount;
        }

        // Eigenbelege verarbeiten
        function processEigenbeleg(row) {
            const m = getMonthFromRow(row, CONFIG.SYSTEM.EIGENBELEGE_COLS.DATUM);
            if (!m) return;

            // Nur bezahlte Eigenbelege berücksichtigen
            if (row[CONFIG.SYSTEM.EIGENBELEGE_COLS.ZAHLUNGSSTATUS] !== 'Bezahlt') return;

            const amount = HelperModule.parseCurrency(row[CONFIG.SYSTEM.EIGENBELEGE_COLS.NETTO]);
            if (amount === 0) return;

            const category = row[CONFIG.SYSTEM.EIGENBELEGE_COLS.KATEGORIE]?.toString().trim() || "";

            // Steuerlichen Status prüfen
            let isTaxFree = false;

            if (CONFIG.EIGENBELEGE && CONFIG.EIGENBELEGE.mapping && CONFIG.EIGENBELEGE.mapping[category]) {
                const taxType = CONFIG.EIGENBELEGE.mapping[category].taxType;
                isTaxFree = taxType === "steuerfrei";
            } else {
                // Fallback: Nach Schlüsselwörtern suchen
                isTaxFree = category.toLowerCase().includes("steuerfrei") ||
                    category.toLowerCase().includes("privat");
            }

            if (isTaxFree) {
                bwaData[m].eigenbelegeSteuerfrei += amount;
            } else {
                bwaData[m].eigenbelegeSteuerpflichtig += amount;

                // Je nach Kategorie in die entsprechende BWA-Position einsortieren
                const lowerCategory = category.toLowerCase();

                if (lowerCategory.includes("bewirtung")) {
                    bwaData[m].sonstigeAufwendungen += amount;
                } else if (lowerCategory.includes("reise")) {
                    bwaData[m].reisekosten += amount;
                } else if (lowerCategory.includes("büro")) {
                    bwaData[m].buerokosten += amount;
                } else {
                    bwaData[m].sonstigeAufwendungen += amount;
                }
            }
        }

        // Daten laden und verarbeiten
        if (revenueSheet) {
            revenueSheet.getDataRange().getValues().slice(1).forEach(processRevenue);
        }

        if (expenseSheet) {
            expenseSheet.getDataRange().getValues().slice(1).forEach(processExpense);
        }

        if (eigenSheet) {
            eigenSheet.getDataRange().getValues().slice(1).forEach(processEigenbeleg);
        }

        // Gruppensummen und weitere Berechnungen
        for (let m = 1; m <= 12; m++) {
            const d = bwaData[m];

            // Erlöse
            d.gesamtErloese = d.umsatzerloese + d.provisionserloese +
                d.steuerfreieInlandEinnahmen + d.steuerfreieAuslandEinnahmen +
                d.sonstigeErtraege + d.vermietung + d.zuschuesse +
                d.waehrungsgewinne + d.anlagenabgaenge;

            // Wareneinsatz
            d.gesamtWareneinsatz = d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe;

            // Betriebsausgaben
            d.gesamtBetriebsausgaben = d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten +
                d.werbungMarketing + d.reisekosten + d.versicherungen +
                d.telefonInternet + d.buerokosten + d.fortbildungskosten +
                d.kfzKosten + d.sonstigeAufwendungen;

            // Abschreibungen und Zinsen
            d.gesamtAbschreibungenZinsen = d.abschreibungenMaschinen + d.abschreibungenBueromaterial +
                d.abschreibungenImmateriell + d.zinsenBank +
                d.zinsenGesellschafter + d.leasingkosten;

            // Besondere Posten
            d.gesamtBesonderePosten = d.eigenkapitalveraenderungen + d.gesellschafterdarlehen +
                d.ausschuettungen;

            // Rückstellungen
            d.gesamtRueckstellungenTransfers = d.steuerrueckstellungen + d.rueckstellungenSonstige;

            // EBIT berechnen
            d.ebit = d.gesamtErloese - (
                d.gesamtWareneinsatz +
                d.gesamtBetriebsausgaben +
                d.gesamtAbschreibungenZinsen +
                d.gesamtBesonderePosten +
                d.gesamtRueckstellungenTransfers
            );

            // Steuern berechnen
            // Standardsätze aus Konfiguration oder Fallback-Werte
            const gewerbesteuerSatz = CONFIG.STEUER?.gewerbesteuer || 0.1645; // 16,45%
            const koerperschaftsteuerSatz = CONFIG.STEUER?.koerperschaftsteuer || 0.15; // 15%
            const soliSatz = CONFIG.STEUER?.solidaritaetszuschlag || 0.055; // 5,5% auf KSt

            d.gewerbesteuer = d.ebit * gewerbesteuerSatz;
            d.koerperschaftsteuer = d.ebit * koerperschaftsteuerSatz;
            d.solidaritaetszuschlag = d.koerperschaftsteuer * soliSatz;

            // Gesamtsteuerlast
            d.steuerlast = d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer;

            // Gewinn nach Steuern
            d.gewinnNachSteuern = d.ebit - d.steuerlast;
        }

        return bwaData;
    }

    /**
     * Erzeugt den Header für die BWA-Tabelle mit Monats- und Quartalsspalten
     * @returns {Array} Header-Zeile
     */
    function buildHeaderRow() {
        const headers = ["Kategorie"];
        const months = [
            "Januar", "Februar", "März", "April", "Mai", "Juni",
            "Juli", "August", "September", "Oktober", "November", "Dezember"
        ];

        // Für jedes Quartal: 3 Monate + Quartalssumme
        for (let q = 0; q < 4; q++) {
            // Drei Monate des Quartals
            for (let m = 0; m < 3; m++) {
                headers.push(`${months[q * 3 + m]} (€)`);
            }

            // Quartalssumme
            headers.push(`Q${q + 1} (€)`);
        }

        // Jahressumme
        headers.push("Jahr (€)");

        return headers;
    }

    // Öffentliche API
    return {
        /**
         * Berechnet die BWA und schreibt die Ergebnisse in das BWA-Sheet
         * @returns {boolean} true bei Erfolg, false bei Fehler
         */
        calculateBWA: function() {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();

                // BWA-Daten sammeln
                const bwaData = aggregateBWAData();
                if (!bwaData) return false;

                // BWA-Positionen definieren
                const positions = [
                    {label: "Erlöse aus Lieferungen und Leistungen", get: d => d.umsatzerloese || 0},
                    {label: "Provisionserlöse", get: d => d.provisionserloese || 0},
                    {label: "Steuerfreie Inlandserlöse", get: d => d.steuerfreieInlandEinnahmen || 0},
                    {label: "Steuerfreie Auslandserlöse", get: d => d.steuerfreieAuslandEinnahmen || 0},
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
                    {label: "Gesamte Betriebsausgaben", get: d => d.gesamtBetriebsausgaben || 0},
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
                    {label: "Gesamt besondere Posten", get: d => d.gesamtBesonderePosten || 0},
                    {label: "Steuerrückstellungen", get: d => d.steuerrueckstellungen || 0},
                    {label: "Rückstellungen sonstige", get: d => d.rueckstellungenSonstige || 0},
                    {label: "Gesamt Rückstellungen", get: d => d.gesamtRueckstellungenTransfers || 0},
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

                /**
                 * Erstellt eine Zeile mit Monats-, Quartals- und Jahreswerten für eine Position
                 * @param {Object} pos - Position mit label und get-Funktion
                 * @returns {Array} Zeile mit Werten
                 */
                function buildOutputRow(pos) {
                    const monthly = [];
                    let yearly = 0;

                    // Monatswerte sammeln
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

                    // Zeile zusammenbauen: Label + Monate + Quartale + Jahr
                    return [
                        pos.label,
                        // Q1: Jan, Feb, Mär + Q1-Summe
                        monthly[0], monthly[1], monthly[2], quarters[0],
                        // Q2: Apr, Mai, Jun + Q2-Summe
                        monthly[3], monthly[4], monthly[5], quarters[1],
                        // Q3: Jul, Aug, Sep + Q3-Summe
                        monthly[6], monthly[7], monthly[8], quarters[2],
                        // Q4: Okt, Nov, Dez + Q4-Summe
                        monthly[9], monthly[10], monthly[11], quarters[3],
                        // Jahressumme
                        yearly
                    ];
                }

                // Gruppierung der Positionen mit Überschriften
                const groups = [
                    {header: "1️⃣ Betriebserlöse (Einnahmen)", count: 10},
                    {header: "2️⃣ Materialaufwand & Wareneinsatz", count: 4},
                    {header: "3️⃣ Betriebsausgaben (Sachkosten)", count: 12},
                    {header: "4️⃣ Abschreibungen & Zinsen", count: 7},
                    {header: "5️⃣ Besondere Posten", count: 4},
                    {header: "6️⃣ Rückstellungen", count: 3},
                    {header: "7️⃣ Betriebsergebnis vor Steuern (EBIT)", count: 1},
                    {header: "8️⃣ Steuern & Vorsteuer", count: 8},
                    {header: "9️⃣ Jahresüberschuss/-fehlbetrag", count: 1}
                ];

                // Zeilen mit Gruppierung erstellen
                let posIndex = 0;
                for (const group of groups) {
                    // Gruppenüberschrift
                    outputRows.push([group.header, ...Array(headerRow.length - 1).fill("")]);

                    // Positionen der Gruppe
                    for (let i = 0; i < group.count; i++) {
                        if (posIndex < positions.length) {
                            outputRows.push(buildOutputRow(positions[posIndex++]));
                        }
                    }
                }

                // BWA-Sheet erstellen oder aktualisieren
                const bwaSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.BWA) ||
                    ss.insertSheet(CONFIG.SYSTEM.SHEET_NAMES.BWA);
                bwaSheet.clearContents();

                // Daten ins Sheet schreiben
                bwaSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);

                // Formatierung
                applyBWAFormatting(bwaSheet, outputRows.length, headerRow.length);

                // Erfolgsmeldung
                HelperModule.showToast("BWA wurde erfolgreich erstellt", "Erfolg");
                SpreadsheetApp.getUi().alert("BWA wurde aktualisiert!");

                return true;
            } catch (e) {
                Logger.log(`Fehler bei der BWA-Berechnung: ${e.message}`);
                SpreadsheetApp.getUi().alert("Fehler bei der BWA-Berechnung: " + e.message);
                return false;
            }
        },

        /**
         * Aktualisiert das BWA-Sheet
         * @returns {boolean} true bei Erfolg, false bei Fehler
         */
        refreshBWA: function() {
            return this.calculateBWA();
        },

        /**
         * Erstellt ein Diagramm für die BWA-Daten
         * @param {string} type - Diagrammtyp ('monthly', 'quarterly', 'yearly')
         * @param {Array} categories - Zu visualisierende Kategorien
         * @returns {Object} Erstelltes Diagramm oder Fehlerobjekt
         */
        createChart: function(type = 'monthly', categories = ['Betriebserlöse', 'Gesamtkosten', 'EBIT']) {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const bwaSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.BWA);

                if (!bwaSheet) {
                    throw new Error(`Sheet "${CONFIG.SYSTEM.SHEET_NAMES.BWA}" nicht gefunden`);
                }

                // BWA-Daten laden
                const data = bwaSheet.getDataRange().getValues();
                if (data.length <= 1) {
                    throw new Error("Keine BWA-Daten vorhanden");
                }

                // Zeilenindizes für die gewünschten Kategorien finden
                const categoryIndices = [];
                for (const category of categories) {
                    for (let i = 1; i < data.length; i++) {
                        if (data[i][0].includes(category)) {
                            categoryIndices.push(i);
                            break;
                        }
                    }
                }

                // Prüfen, ob alle Kategorien gefunden wurden
                if (categoryIndices.length !== categories.length) {
                    throw new Error("Nicht alle gewünschten Kategorien wurden gefunden");
                }

                // Spaltenindizes je nach Diagrammtyp
                let columnIndices = [];

                switch (type.toLowerCase()) {
                    case 'monthly':
                        // Alle Monatsspalten (1, 2, 3, 5, 6, 7, ...)
                        columnIndices = [1, 2, 3, 5, 6, 7, 9, 10, 11, 13, 14, 15];
                        break;
                    case 'quarterly':
                        // Alle Quartalsspalten (4, 8, 12, 16)
                        columnIndices = [4, 8, 12, 16];
                        break;
                    case 'yearly':
                        // Jahresspalte (17)
                        columnIndices = [17];
                        break;
                    default:
                        throw new Error(`Ungültiger Diagrammtyp: ${type}`);
                }

                // Diagramm erstellen
                const chartBuilder = bwaSheet.newChart()
                    .setChartType(SpreadsheetApp.ChartType.COLUMN)
                    .setPosition(data.length + 2, 1, 0, 0)
                    .setNumHeaders(1)
                    .setOption('title', `BWA ${type.charAt(0).toUpperCase() + type.slice(1)}`)
                    .setOption('legend', {position: 'top'})
                    .setOption('hAxis', {title: type === 'monthly' ? 'Monat' : (type === 'quarterly' ? 'Quartal' : 'Jahr')})
                    .setOption('vAxis', {title: 'Betrag (€)'});

                // X-Achse (Header-Zeile)
                const headerRow = data[0];
                const domainRange = bwaSheet.getRange(1, 1, 1, headerRow.length);
                chartBuilder.addRange(domainRange);

                // Datenreihen für jede Kategorie hinzufügen
                for (const rowIndex of categoryIndices) {
                    for (const colIndex of columnIndices) {
                        const range = bwaSheet.getRange(rowIndex + 1, colIndex + 1, 1, 1);
                        chartBuilder.addRange(range);
                    }
                }

                // Diagramm ins Sheet einfügen
                const chart = chartBuilder.build();
                bwaSheet.insertChart(chart);

                return {
                    success: true,
                    chart: chart
                };
            } catch (e) {
                Logger.log(`Fehler bei der Diagrammerstellung: ${e.message}`);
                return {
                    success: false,
                    error: e.message
                };
            }
        },

        /**
         * Exportiert die BWA-Daten als CSV
         * @returns {Object} Ergebnis des Exports
         */
        exportCSV: function() {
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const bwaSheet = ss.getSheetByName(CONFIG.SYSTEM.SHEET_NAMES.BWA);

                if (!bwaSheet) {
                    throw new Error(`Sheet "${CONFIG.SYSTEM.SHEET_NAMES.BWA}" nicht gefunden`);
                }

                // BWA-Daten laden
                const data = bwaSheet.getDataRange().getValues();
                if (data.length <= 1) {
                    throw new Error("Keine BWA-Daten vorhanden");
                }

                // CSV erstellen
                let csv = "";
                data.forEach(row => {
                    // Jede Zelle formatieren und mit Semicolon trennen
                    const csvRow = row.map(cell => {
                        // Zahlen mit 2 Dezimalstellen formatieren
                        if (typeof cell === 'number') {
                            return cell.toFixed(2);
                        }
                        // Strings in Anführungszeichen setzen, wenn sie Kommas oder Zeilenumbrüche enthalten
                        if (typeof cell === 'string' && (cell.includes(',') || cell.includes('\n') || cell.includes(';'))) {
                            return `"${cell.replace(/"/g, '""')}"`;
                        }
                        return cell;
                    }).join(';');

                    csv += csvRow + '\n';
                });

                // CSV anzeigen oder speichern
                const ui = SpreadsheetApp.getUi();
                ui.alert(
                    'BWA CSV-Export',
                    'Die BWA-Daten wurden erfolgreich als CSV exportiert.',
                    ui.ButtonSet.OK
                );

                return {
                    success: true,
                    csvData: csv
                };
            } catch (e) {
                Logger.log(`Fehler beim CSV-Export: ${e.message}`);
                SpreadsheetApp.getUi().alert("Fehler beim CSV-Export: " + e.message);
                return {
                    success: false,
                    error: e.message
                };
            }
        }
    };

    /**
     * Wendet Formatierung auf das BWA-Sheet an
     * @param {Object} sheet - Das BWA-Sheet
     * @param {number} numRows - Anzahl der Zeilen
     * @param {number} numCols - Anzahl der Spalten
     */
    function applyBWAFormatting(sheet, numRows, numCols) {
        // Spaltenbreiten anpassen
        sheet.autoResizeColumns(1, numCols);

        // Zahlenformat für Beträge
        sheet.getRange(2, 2, numRows - 1, numCols - 1).setNumberFormat("#,##0.00 €");

        // Header-Zeile formatieren
        sheet.getRange(1, 1, 1, numCols).setFontWeight("bold");
        sheet.getRange(1, 1, 1, numCols).setBackground("#f3f3f3");

        // Gruppenüberschriften formatieren
        for (let i = 1; i < numRows; i++) {
            const value = sheet.getRange(i + 1, 1).getValue();
            if (typeof value === 'string' && value.includes('️⃣')) {
                sheet.getRange(i + 1, 1, 1, numCols).setFontWeight('bold');
                sheet.getRange(i + 1, 1, 1, numCols).setBackground('#f3f3f3');
            }
        }

        // Quartalssummen und Jahressumme hervorheben
        for (let i = 4; i <= numCols; i += 4) {
            sheet.getRange(1, i, numRows, 1).setBackground("#e6f2ff");
        }
        sheet.getRange(1, numCols, numRows, 1).setBackground("#d9ead3");

        // Rahmen für bestimmte Zeilen
        // Hier könnten weitere Formatierungen hinzugefügt werden
    }
})();