// imports TEST
import ImportModule from "./importModule.js";
import RefreshModule from "./refreshModule.js";
import UStVACalculator from "./uStVACalculator.js";
import BWACalculator from "./bWACalculator.js";
// import calculateBilanz from "./calculateBilanz.js";

// =================== Globale Funktionen ===================
const onOpen = () => {
    SpreadsheetApp.getUi()
        .createMenu("📂 Buchhaltung")
        .addItem("📥 Dateien importieren", "importDriveFiles")
        .addItem("🔄 Refresh Active Sheet", "refreshSheet")
        .addItem("📊 UStVA berechnen", "calculateUStVA")
        .addItem("📈 BWA berechnen", "calculateBWA")
        .addItem("📝 Bilanz erstellen", "calculateBilanz")
        .addToUi();
};

const onEdit = e => {
    const {range} = e;
    const sheet = range.getSheet();
    const name = sheet.getName();
    const mapping = {
        "Einnahmen": 16,
        "Ausgaben": 16,
        "Eigenbelege": 16,
        "Bankbewegungen": 11,
        "Gesellschafterkonto": 12,
        "Holding Transfers": 6
    };
    if (!(name in mapping)) return;
    if (range.getRow() === 1) return;
    const headerLen = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].length;
    if (range.getColumn() > headerLen) return;
    if (range.getColumn() === mapping[name]) return;
    const rowValues = sheet.getRange(range.getRow(), 1, 1, headerLen).getValues()[0];
    if (rowValues.every(cell => cell === "")) return;
    const ts = new Date();
    sheet.getRange(range.getRow(), mapping[name]).setValue(ts);
};

const setupTrigger = () => {
    const triggers = ScriptApp.getProjectTriggers();
    if (!triggers.some(t => t.getHandlerFunction() === "onOpen"))
        SpreadsheetApp.newTrigger("onOpen")
            .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
            .onOpen()
            .create();
};

const refreshSheet = () => RefreshModule.refreshActiveSheet();
const calculateUStVA = () => {
    RefreshModule.refreshAllSheets();
    UStVACalculator.calculateUStVA();
};
const calculateBWA = () => {
    RefreshModule.refreshAllSheets();
    BWACalculator.calculateBWA();
};
const importDriveFiles = () => {
    ImportModule.importDriveFiles();
    RefreshModule.refreshAllSheets();
};

