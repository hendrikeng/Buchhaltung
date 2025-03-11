import Helpers from "./helpers.js";
import config from "./config.js";
import Validator from "./validator.js";

const UStVACalculator = (() => {
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

    // Hier entfällt die eigene getMonthFromRow – stattdessen nutzen wir Helpers.getMonthFromRow:
    const processUStVARow = (row, data, isIncome, isEigen = false) => {
        const paymentDate = Helpers.parseDate(row[13]);
        if (!paymentDate || paymentDate > new Date()) return;
        const month = Helpers.getMonthFromRow(row);
        if (!month) return;
        const netto = Helpers.parseCurrency(row[4]);
        const restNetto = Helpers.parseCurrency(row[9]) || 0;
        const gezahlt = netto - restNetto; // Negative Werte bleiben erhalten
        const mwstRate = Helpers.parseMwstRate(row[5]);
        const roundedRate = Math.round(mwstRate);
        const tax = gezahlt * (mwstRate / 100);
        const category = row[2]?.toString().trim() || "";

        if (isIncome) {
            const catCfg = config.einnahmen.categories[category] ?? {};
            const taxType = catCfg.taxType ?? "steuerpflichtig";
            if (taxType === "steuerfrei_inland") {
                data[month].steuerfreie_inland_einnahmen += gezahlt;
            } else if (taxType === "steuerfrei_ausland" || !roundedRate) {
                data[month].steuerfreie_ausland_einnahmen += gezahlt;
            } else {
                data[month].steuerpflichtige_einnahmen += gezahlt;
                data[month][`ust_${roundedRate}`] += tax;
            }
        } else {
            if (isEigen) {
                const eigenCfg = config.eigenbelege.mapping[category] ?? {};
                const taxType = eigenCfg.taxType ?? "steuerpflichtig";
                if (taxType === "steuerfrei") {
                    data[month].eigenbelege_steuerfrei += gezahlt;
                } else if (taxType === "eigenbeleg" && eigenCfg.besonderheit === "bewirtung") {
                    data[month].eigenbelege_steuerpflichtig += gezahlt;
                    data[month][`vst_${roundedRate}`] += tax * 0.7;
                    data[month].nicht_abzugsfaehige_vst += tax * 0.3;
                } else {
                    data[month].eigenbelege_steuerpflichtig += gezahlt;
                    data[month][`vst_${roundedRate}`] += tax;
                }
            } else {
                const catCfg = config.ausgaben.categories[category] ?? {};
                const taxType = catCfg.taxType ?? "steuerpflichtig";
                if (taxType === "steuerfrei_inland") {
                    data[month].steuerfreie_inland_ausgaben += gezahlt;
                } else if (taxType === "steuerfrei_ausland") {
                    data[month].steuerfreie_ausland_ausgaben += gezahlt;
                } else {
                    data[month].steuerpflichtige_ausgaben += gezahlt;
                    data[month][`vst_${roundedRate}`] += tax;
                }
            }
        }
    };

    const formatUStVARow = (label, d) => {
        const ustZahlung = (d.ust_7 + d.ust_19) - ((d.vst_7 + d.vst_19) - d.nicht_abzugsfaehige_vst);
        const ergebnis = (d.steuerpflichtige_einnahmen + d.steuerfreie_inland_einnahmen + d.steuerfreie_ausland_einnahmen) -
            (d.steuerpflichtige_ausgaben + d.steuerfreie_inland_ausgaben + d.steuerfreie_ausland_ausgaben +
                d.eigenbelege_steuerpflichtig + d.eigenbelege_steuerfrei);
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

    const aggregateUStVA = (data, start, end) => {
        const sum = createEmptyUStVA();
        for (let m = start; m <= end; m++) {
            for (const key in sum) {
                sum[key] += data[m][key] || 0;
            }
        }
        return sum;
    };

    const calculateUStVA = () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const revenueSheet = ss.getSheetByName("Einnahmen");
        const expenseSheet = ss.getSheetByName("Ausgaben");
        const eigenSheet = ss.getSheetByName("Eigenbelege");
        if (!revenueSheet || !expenseSheet) {
            SpreadsheetApp.getUi().alert("Fehlendes Blatt: 'Einnahmen' oder 'Ausgaben'");
            return;
        }
        if (!Validator.validateAllSheets(revenueSheet, expenseSheet)) return;

        const revenueData = revenueSheet.getDataRange().getValues();
        const expenseData = expenseSheet.getDataRange().getValues();
        const eigenData = eigenSheet ? eigenSheet.getDataRange().getValues() : [];
        const ustvaData = Object.fromEntries(Array.from({length: 12}, (_, i) => [i + 1, createEmptyUStVA()]));

        const processRows = (data, isIncome, isEigen = false) =>
            data.slice(1).forEach(row => {
                const m = Helpers.getMonthFromRow(row);
                if (m) processUStVARow(row, ustvaData, isIncome, isEigen);
            });
        processRows(revenueData, true);
        processRows(expenseData, false);
        if (eigenData.length) processRows(eigenData, false, true);

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
        config.common.months.forEach((name, i) => {
            outputRows.push(formatUStVARow(name, ustvaData[i + 1]));
            if ((i + 1) % 3 === 0) {
                outputRows.push(formatUStVARow(`Quartal ${(i + 1) / 3}`, aggregateUStVA(ustvaData, i - 1, i + 1)));
            }
        });
        outputRows.push(formatUStVARow("Gesamtjahr", aggregateUStVA(ustvaData, 1, 12)));

        const ustvaSheet = ss.getSheetByName("UStVA") || ss.insertSheet("UStVA");
        ustvaSheet.clearContents();
        ustvaSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
        ustvaSheet.autoResizeColumns(1, outputRows[0].length);
        ss.setActiveSheet(ustvaSheet);
        SpreadsheetApp.getUi().alert("UStVA wurde aktualisiert!");
    };

    return {calculateUStVA};
})();

export default UStVACalculator;