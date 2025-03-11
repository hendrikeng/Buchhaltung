import Helpers from "./helpers.js";
import config from "./config.js";

const BWACalculator = (() => {
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

    // Hier nutzen wir die zentrale Helpers.getMonthFromRow-Funktion:
    const getMonthFromRow = row => {
        return Helpers.getMonthFromRow(row);
    };

    const aggregateBWAData = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const eigenSheet = ss.getSheetByName("Eigenbelege");
        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
            return null;
        }
        const bwaData = Object.fromEntries(Array.from({length: 12}, (_, i) => [i + 1, createEmptyBWA()]));

        const processRevenue = row => {
            const m = getMonthFromRow(row);
            if (!m) return;
            const amount = Helpers.parseCurrency(row[4]);
            const category = row[2]?.toString().trim() || "";
            if (["Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) return;
            if (category === "Sonstige betriebliche Erträge") return void (bwaData[m].sonstigeErtraege += amount);
            if (category === "Erträge aus Vermietung/Verpachtung") return void (bwaData[m].vermietung += amount);
            if (category === "Erträge aus Zuschüssen") return void (bwaData[m].zuschuesse += amount);
            if (category === "Erträge aus Währungsgewinnen") return void (bwaData[m].waehrungsgewinne += amount);
            if (category === "Erträge aus Anlagenabgängen") return void (bwaData[m].anlagenabgaenge += amount);
            const mapping = config.einnahmen.bwaMapping[category];
            if (["umsatzerloese", "provisionserloese"].includes(mapping)) {
                bwaData[m][mapping] += amount;
            } else if (Helpers.parseMwstRate(row[5]) === 0) {
                bwaData[m].steuerfreieInlandEinnahmen += amount;
            } else {
                bwaData[m].umsatzerloese += amount;
            }
        };

        const processExpense = row => {
            const m = getMonthFromRow(row);
            if (!m) return;
            const amount = Helpers.parseCurrency(row[4]);
            const category = row[2]?.toString().trim() || "";
            if (["Privatentnahme", "Privateinlage", "Holding Transfers",
                "Gewinnvortrag", "Verlustvortrag", "Gewinnvortrag/Verlustvortrag"].includes(category)) return;
            if (category === "Bruttolöhne & Gehälter") return void (bwaData[m].bruttoLoehne += amount);
            if (category === "Soziale Abgaben & Arbeitgeberanteile") return void (bwaData[m].sozialeAbgaben += amount);
            if (category === "Sonstige Personalkosten") return void (bwaData[m].sonstigePersonalkosten += amount);
            if (category === "Gewerbesteuerrückstellungen") return void (bwaData[m].gewerbesteuerRueckstellungen += amount);
            if (category === "Telefon & Internet") return void (bwaData[m].telefonInternet += amount);
            if (category === "Bürokosten") return void (bwaData[m].buerokosten += amount);
            if (category === "Fortbildungskosten") return void (bwaData[m].fortbildungskosten += amount);
            const mapping = config.ausgaben.bwaMapping[category];
            switch (mapping) {
                case "wareneinsatz":
                    bwaData[m].wareneinsatz += amount;
                    break;
                case "bezogeneLeistungen":
                    bwaData[m].fremdleistungen += amount;
                    break;
                case "rohHilfsBetriebsstoffe":
                    bwaData[m].rohHilfsBetriebsstoffe += amount;
                    break;
                case "betriebskosten":
                    bwaData[m].sonstigeAufwendungen += amount;
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
                case "zinsenBank":
                    bwaData[m].zinsenBank += amount;
                    break;
                case "zinsenGesellschafter":
                    bwaData[m].zinsenGesellschafter += amount;
                    break;
                case "leasingkosten":
                    bwaData[m].leasingkosten += amount;
                    break;
                default:
                    bwaData[m].sonstigeAufwendungen += amount;
            }
        };

        const processEigen = row => {
            const m = getMonthFromRow(row);
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
        };

        revenueSheet.getDataRange().getValues().slice(1).forEach(processRevenue);
        expenseSheet.getDataRange().getValues().slice(1).forEach(processExpense);
        if (eigenSheet) eigenSheet.getDataRange().getValues().slice(1).forEach(processEigen);

        // Gruppensummen und weitere Berechnungen
        for (let m = 1; m <= 12; m++) {
            const d = bwaData[m];
            d.gesamtErloese = d.umsatzerloese + d.provisionserloese + d.steuerfreieInlandEinnahmen + d.steuerfreieAuslandEinnahmen +
                d.sonstigeErtraege + d.vermietung + d.zuschuesse + d.waehrungsgewinne + d.anlagenabgaenge;
            d.gesamtWareneinsatz = d.wareneinsatz + d.fremdleistungen + d.rohHilfsBetriebsstoffe;
            d.gesamtBetriebsausgaben = d.bruttoLoehne + d.sozialeAbgaben + d.sonstigePersonalkosten +
                d.werbungMarketing + d.reisekosten + d.versicherungen + d.telefonInternet +
                d.buerokosten + d.fortbildungskosten + d.kfzKosten + d.sonstigeAufwendungen;
            d.gesamtAbschreibungenZinsen = d.abschreibungenMaschinen + d.abschreibungenBueromaterial +
                d.abschreibungenImmateriell + d.zinsenBank + d.zinsenGesellschafter + d.leasingkosten;
            d.gesamtBesonderePosten = d.eigenkapitalveraenderungen + d.gesellschafterdarlehen + d.ausschuettungen;
            d.gesamtRueckstellungenTransfers = d.steuerrueckstellungen + d.rueckstellungenSonstige;
            d.ebit = d.gesamtErloese - (d.gesamtWareneinsatz + d.gesamtBetriebsausgaben + d.gesamtAbschreibungenZinsen + d.gesamtBesonderePosten);
            d.gewerbesteuer = d.ebit * (config.tax.operative.gewerbesteuer / 100);
            if (config.tax.isHolding) {
                d.koerperschaftsteuer = d.ebit * (config.tax.holding.koerperschaftsteuer / 100) * (config.tax.holding.gewinnUebertragSteuerpflichtig / 100);
                d.solidaritaetszuschlag = d.ebit * (config.tax.holding.solidaritaetszuschlag / 100) * (config.tax.holding.gewinnUebertragSteuerpflichtig / 100);
            } else {
                d.koerperschaftsteuer = d.ebit * (config.tax.operative.koerperschaftsteuer / 100);
                d.solidaritaetszuschlag = d.ebit * (config.tax.operative.solidaritaetszuschlag / 100);
            }
            d.steuerlast = d.koerperschaftsteuer + d.solidaritaetszuschlag + d.gewerbesteuer;
            d.gewinnNachSteuern = d.ebit - d.steuerlast;
        }
        return bwaData;
    };

    // Erzeugt den Header (2 Spalten: Bezeichnung und Wert) mit Monats- und Quartalsspalten
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

    // Hier wird für jede Position eine Zeile mit Monats-, Quartals- und Jahreswerten aufgebaut:
    const calculateBWA = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const bwaData = aggregateBWAData();
        if (!bwaData) return;

        const positions = [
            {label: "Erlöse aus Lieferungen und Leistungen", get: d => d.umsatzerloese || 0},
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

        const headerRow = buildHeaderRow();
        const outputRows = [headerRow];

        // Für jede Position wird eine Zeile mit Monats-, Quartals- und Jahreswerten aufgebaut:
        const buildOutputRow = pos => {
            const monthly = [];
            let yearly = 0;
            for (let m = 1; m <= 12; m++) {
                const val = pos.get(bwaData[m]) || 0;
                monthly.push(val);
                yearly += val;
            }
            const quarters = [0, 0, 0, 0];
            for (let i = 0; i < 12; i++) {
                quarters[Math.floor(i / 3)] += monthly[i];
            }
            return [pos.label, ...monthly.slice(0, 3), quarters[0], ...monthly.slice(3, 6), quarters[1],
                ...monthly.slice(6, 9), quarters[2], ...monthly.slice(9, 12), quarters[3], yearly];
        };

        let posIndex = 0;
        for (const {header, count} of [
            {header: "1️⃣ Betriebserlöse (Einnahmen)", count: 7},
            {header: "2️⃣ Materialaufwand & Wareneinsatz", count: 4},
            {header: "3️⃣ Betriebsausgaben (Sachkosten)", count: 10},
            {header: "4️⃣ Abschreibungen & Zinsen", count: 7},
            {header: "5️⃣ Besondere Posten", count: 3},
            {header: "6️⃣ Rückstellungen", count: 2},
            {header: "7️⃣ Betriebsergebnis vor Steuern (EBIT)", count: 1},
            {header: "8️⃣ Steuern & Vorsteuer", count: 7},
            {header: "9️⃣ Jahresüberschuss/-fehlbetrag", count: 1}
        ]) {
            outputRows.push([header, ...Array(headerRow.length - 1).fill("")]);
            for (let i = 0; i < count; i++) {
                outputRows.push(buildOutputRow(positions[posIndex++]));
            }
        }

        const bwaSheet = ss.getSheetByName("BWA") || ss.insertSheet("BWA");
        bwaSheet.clearContents();
        bwaSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
        bwaSheet.autoResizeColumns(1, outputRows[0].length);
        SpreadsheetApp.getUi().alert("BWA wurde aktualisiert!");
    };

    return {calculateBWA};
})();

export default BWACalculator;