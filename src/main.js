/**
 * main.js - Hauptdatei des Buchhaltungssystems
 *
 * Enthält die Initialisierungslogik und Menüfunktionen
 */

/**
 * onOpen wird automatisch beim Öffnen des Spreadsheets ausgeführt
 * Erstellt das benutzerdefinierte Menü
 */
function onOpen() {
    try {
        const ui = SpreadsheetApp.getUi();

        ui.createMenu('Buchhaltung')
            .addSubMenu(ui.createMenu('Daten')
                .addItem('📥 Dateien importieren', 'ImportModule.showImportDialog')
                .addItem('🔄 Aktuelles Sheet aktualisieren', 'refreshActiveSheet')
                .addItem('🔍 Bankbewegungen automatisch zuordnen', 'BankModule.matchBankbewegungToRechnung')
                .addSeparator()
                .addItem('📊 Alle Daten aktualisieren', 'refreshAllData'))
            .addSubMenu(ui.createMenu('Berichte')
                .addItem('📊 UStVA berechnen', 'UStVAModule.calculateUStVA')
                .addItem('📈 BWA berechnen', 'BWAModule.calculateBWA')
                .addItem('📝 Bilanz erstellen', 'BilanzModule.generateBilanz')
                .addSeparator()
                .addItem('📤 ELSTER-Export erstellen', 'UStVAModule.createElsterExport'))
            .addSubMenu(ui.createMenu('Einstellungen')
                .addItem('⚙️ Konfiguration', 'SettingsModule.showConfigDialog')
                .addItem('👤 Benutzereinstellungen', 'SettingsModule.showUserSettings')
                .addItem('🔧 System-Diagnose', 'SettingsModule.runSystemDiagnostics'))
            .addSubMenu(ui.createMenu('Hilfe')
                .addItem('📖 Anleitung', 'HelpModule.showHelp')
                .addItem('📅 Steuertermine', 'HelpModule.showTaxCalendar')
                .addItem('ℹ️ Über', 'HelpModule.showAbout'))
            .addToUi();

        Logger.log('Buchhaltungsmenü erfolgreich erstellt');
    } catch (e) {
        Logger.log('Fehler beim Erstellen des Menüs: ' + e.message);
    }
}

/**
 * Aktualisiert das aktuell ausgewählte Sheet
 */
function refreshActiveSheet() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const sheetName = sheet.getName();

    try {
        // Je nach Sheet-Typ die entsprechende Aktualisierungsfunktion aufrufen
        switch(sheetName) {
            case CONFIG.SYSTEM.SHEET_NAMES.EINNAHMEN:
                EinnahmenModule.refreshEinnahmen();
                break;
            case CONFIG.SYSTEM.SHEET_NAMES.AUSGABEN:
                AusgabenModule.refreshAusgaben();
                break;
            case CONFIG.SYSTEM.SHEET_NAMES.BANKBEWEGUNGEN:
                BankModule.refreshBankbewegungen();
                break;
            case CONFIG.SYSTEM.SHEET_NAMES.EIGENBELEGE:
                EigenbelagModule.refreshEigenbelege();
                break;
            case CONFIG.SYSTEM.SHEET_NAMES.USTVA:
                UStVAModule.refreshUStVA();
                break;
            case CONFIG.SYSTEM.SHEET_NAMES.BWA:
                BWAModule.refreshBWA();
                break;
            case CONFIG.SYSTEM.SHEET_NAMES.BILANZ:
                BilanzModule.refreshBilanz();
                break;
            default:
                SpreadsheetApp.getUi().alert(
                    'Aktualisierung nicht verfügbar',
                    `Für das Sheet "${sheetName}" ist keine spezifische Aktualisierungsfunktion definiert.`,
                    SpreadsheetApp.getUi().ButtonSet.OK
                );
        }

        // Benachrichtigung über erfolgreiche Aktualisierung
        showToast(`${sheetName} wurde erfolgreich aktualisiert`, 'Erfolg');
    } catch (e) {
        // Fehler protokollieren und Benutzer benachrichtigen
        Logger.log(`Fehler beim Aktualisieren von ${sheetName}: ${e.message}`);
        showToast(`Fehler beim Aktualisieren: ${e.message}`, 'Fehler');
    }
}

/**
 * Aktualisiert alle relevanten Sheets in der optimalen Reihenfolge
 */
function refreshAllData() {
    const ui = SpreadsheetApp.getUi();

    // Bestätigung vom Benutzer einholen
    const response = ui.alert(
        'Alle Daten aktualisieren?',
        'Dies kann je nach Datenmenge einige Zeit dauern. Fortfahren?',
        ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
        return;
    }

    try {
        // Startzeit für Performance-Messung
        const startTime = new Date().getTime();

        // Aktualisierung in der optimalen Reihenfolge
        showToast('Aktualisiere Basisblätter...', 'Info');

        // 1. Basis-Sheets aktualisieren
        EinnahmenModule.refreshEinnahmen();
        AusgabenModule.refreshAusgaben();
        EigenbelagModule.refreshEigenbelege();
        BankModule.refreshBankbewegungen();

        // 2. Abgeleitete Werte aktualisieren
        showToast('Aktualisiere Auswertungen...', 'Info');
        UStVAModule.refreshUStVA();
        BWAModule.refreshBWA();
        BilanzModule.refreshBilanz();

        // Gesamtdauer berechnen und anzeigen
        const endTime = new Date().getTime();
        const duration = (endTime - startTime) / 1000; // Sekunden

        showToast(`Alle Daten erfolgreich aktualisiert (${duration.toFixed(1)} Sekunden)`, 'Erfolg');
    } catch (e) {
        Logger.log(`Fehler bei der Gesamtaktualisierung: ${e.message}`);
        showToast(`Fehler bei der Aktualisierung: ${e.message}`, 'Fehler');
    }
}

/**
 * Zeigt eine kurze Benachrichtigung im Spreadsheet an
 * @param {string} message - Anzuzeigende Nachricht
 * @param {string} type - Typ der Nachricht (Info, Erfolg, Warnung, Fehler)
 */
function showToast(message, type = 'Info') {
    // Nur anzeigen, wenn Benachrichtigungen aktiviert sind
    if (!CONFIG.BENUTZER.BENACHRICHTIGUNG_ENABLED) return;

    // Farben/Icons je nach Meldungstyp anpassen
    let title;
    switch(type.toLowerCase()) {
        case 'erfolg':
            title = '✅ Erfolg';
            break;
        case 'warnung':
            title = '⚠️ Warnung';
            break;
        case 'fehler':
            title = '❌ Fehler';
            break;
        default:
            title = 'ℹ️ Info';
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 5);
}