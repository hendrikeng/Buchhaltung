/**
 * Zentrale Konfigurationsdatei für das Buchhaltungssystem
 *
 * Diese Datei enthält alle globalen Konfigurationsparameter, unterteilt in:
 * - SYSTEM: Konstante Parameter für Sheets, Spalten, etc.
 * - STEUER: Steuerliche Parameter, anpassbar bei Gesetzesänderungen
 * - BENUTZER: Anpassbare Benutzereinstellungen
 */

const CONFIG = {
    // Systemkonfiguration - technische Parameter
    SYSTEM: {
        // Namen der Sheets
        SHEET_NAMES: {
            EINNAHMEN: 'Einnahmen',
            AUSGABEN: 'Ausgaben',
            EIGENBELEGE: 'Eigenbelege',
            BANKBEWEGUNGEN: 'Bankbewegungen',
            GESELLSCHAFTERKONTO: 'Gesellschafterkonto',
            HOLDING_TRANSFERS: 'Holding Transfers',
            USTVA: 'UStVA',
            BWA: 'BWA',
            BILANZ: 'Bilanz',
            RECHNUNGEN_EINNAHMEN: 'Rechnungen Einnahmen',
            RECHNUNGEN_AUSGABEN: 'Rechnungen Ausgaben',
            AENDERUNGSHISTORIE: 'Änderungshistorie'
        },

        // Spaltenindices (0-basiert) für Einnahmen
        EINNAHMEN_COLS: {
            DATUM: 0,
            RECHNUNGSNUMMER: 1,
            KATEGORIE: 2,
            KUNDE: 3,
            NETTO: 4,
            MWST_SATZ: 5,
            MWST_BETRAG: 6,
            BRUTTO: 7,
            BEZAHLT: 8,
            RESTBETRAG: 9,
            QUARTAL: 10,
            ZAHLUNGSSTATUS: 11,
            ZAHLUNGSART: 12,
            ZAHLUNGSDATUM: 13,
            UEBERTRAG_AUS_JAHR: 14,
            LETZTE_AKTUALISIERUNG: 15,
            DATEINAME: 16,
            RECHNUNG_LINK: 17
        },

        // Spaltenindices für Ausgaben (identisch zu Einnahmen)
        AUSGABEN_COLS: {
            DATUM: 0,
            RECHNUNGSNUMMER: 1,
            KATEGORIE: 2,
            LIEFERANT: 3,
            NETTO: 4,
            MWST_SATZ: 5,
            MWST_BETRAG: 6,
            BRUTTO: 7,
            BEZAHLT: 8,
            RESTBETRAG: 9,
            QUARTAL: 10,
            ZAHLUNGSSTATUS: 11,
            ZAHLUNGSART: 12,
            ZAHLUNGSDATUM: 13,
            UEBERTRAG_AUS_JAHR: 14,
            LETZTE_AKTUALISIERUNG: 15,
            DATEINAME: 16,
            RECHNUNG_LINK: 17
        },

        // Spaltenindices für Eigenbelege
        EIGENBELEGE_COLS: {
            DATUM: 0,
            EMPFAENGER: 1,
            KATEGORIE: 2,
            NOTIZ: 3,
            NETTO: 4,
            MWST_SATZ: 5,
            MWST_BETRAG: 6,
            BRUTTO: 7,
            BEZAHLT: 8,
            RESTBETRAG: 9,
            QUARTAL: 10,
            ZAHLUNGSSTATUS: 11,
            ZAHLUNGSART: 12,
            ZAHLUNGSDATUM: 13,
            UEBERTRAG_AUS_JAHR: 14,
            LETZTE_AKTUALISIERUNG: 15,
            REFERENZ_BANKBEWEGUNG: 16
        },

        // Spaltenindices für Bankbewegungen
        BANKBEWEGUNGEN_COLS: {
            DATUM: 0,
            BUCHUNGSTEXT: 1,
            BETRAG: 2,
            SALDO: 3,
            TYP: 4,
            KATEGORIE: 5,
            KONTO_SOLL: 6,
            KONTO_HABEN: 7,
            REFERENZ_RECHNUNG: 8,
            NOTIZEN: 9,
            LETZTE_AKTUALISIERUNG: 10,
            ERFASST: 11
        },

        // Spaltenindices für Gesellschafterkonto
        GESELLSCHAFTERKONTO_COLS: {
            DATUM: 0,
            GESELLSCHAFTER: 1,
            ART: 2,
            VERWENDUNGSZWECK: 3,
            BETRAG: 4,
            BEZAHLT: 5,
            RESTBETRAG: 6,
            QUARTAL: 7,
            ZAHLUNGSSTATUS: 8,
            ZAHLUNGSART: 9,
            ZAHLUNGSDATUM: 10,
            LETZTE_AKTUALISIERUNG: 11,
            UEBERTRAG_AUS_JAHR: 12,
            REFERENZ_BANKBEWEGUNG: 13
        },

        // Spaltenindices für Holding Transfers
        HOLDING_TRANSFERS_COLS: {
            DATUM: 0,
            BETRAG: 1,
            ART: 2,
            BUCHUNGSTEXT: 3,
            STATUS: 4,
            LETZTE_AKTUALISIERUNG: 5,
            REFERENZ_BANKBEWEGUNG: 6
        },

        // Spaltenindices für UStVA
        USTVA_COLS: {
            ZEITRAUM: 0,
            STEUERPFLICHTIGE_EINNAHMEN: 1,
            STEUERFREIE_INLAND_EINNAHMEN: 2,
            STEUERFREIE_AUSLAND_EINNAHMEN: 3,
            STEUERPFLICHTIGE_AUSGABEN: 4,
            STEUERFREIE_INLAND_AUSGABEN: 5,
            STEUERFREIE_AUSLAND_AUSGABEN: 6,
            EIGENBELEGE_STEUERPFLICHTIG: 7,
            EIGENBELEGE_STEUERFREI: 8,
            NICHT_ABZUGSFAEHIGE_VST_BEWIRTUNG: 9,
            UST_7: 10,
            UST_19: 11,
            VST_7: 12,
            VST_19: 13,
            UST_ZAHLUNG: 14,
            ERGEBNIS: 15
        },

        // Spaltennamen für Dropdown-Menüs
        DROPDOWNS: {
            ZAHLUNGSSTATUS: ['Offen', 'Teilweise bezahlt', 'Bezahlt', 'Storniert'],
            ZAHLUNGSART: ['Überweisung', 'Lastschrift', 'Bar', 'PayPal', 'Kreditkarte', 'Sonstige'],
            GESELLSCHAFTER_ART: ['Darlehen', 'Rückzahlung', 'Ausschüttung', 'Kapitalerhöhung', 'Sonstiges'],
            HOLDING_TRANSFER_ART: ['Gewinnübertrag', 'Darlehen', 'Rückzahlung', 'Sonstiges']
        },

        // SKR04-konforme Kategorien für Einnahmen
        EINNAHMEN_KATEGORIEN: [
            'Erlöse aus Lieferungen und Leistungen',
            'Erlöse aus Vermietung und Verpachtung',
            'Erlöse steuerfrei §4 Nr. 8-28 UStG',
            'Erlöse innergemeinschaftliche Lieferungen §4 Nr. 1b UStG',
            'Erlöse Ausfuhrlieferungen §4 Nr. 1a UStG',
            'Sonstige betriebliche Erträge',
            'Zinserträge',
            'Erhaltene Anzahlungen'
        ],

        // SKR04-konforme Kategorien für Ausgaben
        AUSGABEN_KATEGORIEN: [
            'Wareneinkauf 19% VSt',
            'Wareneinkauf 7% VSt',
            'Wareneinkauf ohne VSt',
            'Fremdleistungen',
            'Personalkosten',
            'Raumkosten',
            'Fahrzeugkosten',
            'Werbe- und Reisekosten',
            'Kosten für Warenabgabe',
            'Post und Telekommunikation',
            'Bürobedarf',
            'Versicherungen/Beiträge',
            'Fortbildungskosten',
            'Beratungskosten',
            'Freiwillige soziale Aufwendungen',
            'Sonstige betriebliche Aufwendungen'
        ]
    },

    // Steuerliche Parameter (leicht anpassbar bei Gesetzesänderungen)
    STEUER: {
        // Mehrwertsteuersätze
        MEHRWERTSTEUER_STANDARD: 0.19,   // 19% Standardsatz
        MEHRWERTSTEUER_REDUZIERT: 0.07,  // 7% reduzierter Satz

        // Spezielle steuerliche Regelungen
        BEWIRTUNG_ABZUGSFAEHIG: 0.7,     // 70% der Vorsteuer abzugsfähig bei Bewirtung
        GEWINNUEBERTRAG_STEUERFREI: 0.95 // 95% steuerfrei bei Holding-Transfers
    },

    // Anpassbare Benutzereinstellungen
    BENUTZER: {
        FIRMA_NAMEN: ['Holding GmbH', 'Operative GmbH'],
        GESELLSCHAFTER: ['Gesellschafter 1', 'Gesellschafter 2'],
        AUTO_REFRESH: true,
        BENACHRICHTIGUNG_ENABLED: true,
        AKTUELLES_JAHR: new Date().getFullYear(), // Automatisch aktuelles Jahr
        WAEHRUNG: '€',                            // Währungssymbol

        // UStVA-Einstellungen
        USTVA: {
            MELDEZEITRAUM: 'monatlich',  // 'monatlich' oder 'quartalsweise'
            ELSTER_EXPORT_PFAD: ''       // Pfad für ELSTER-Export (falls implementiert)
        }
    }
};

/**
 * Funktion zum Abrufen von Konfigurationswerten
 * @param {string} section - Konfigurationsbereich (SYSTEM, STEUER, BENUTZER)
 * @param {string} key - Konfigurationsschlüssel
 * @param {string|number|boolean} defaultValue - Standardwert, falls Schlüssel nicht existiert
 * @returns {*} Der Konfigurationswert oder der Standardwert
 */
function getConfig(section, key, defaultValue = null) {
    try {
        if (CONFIG[section] && CONFIG[section][key] !== undefined) {
            return CONFIG[section][key];
        }
        return defaultValue;
    } catch (e) {
        console.error(`Fehler beim Abrufen der Konfiguration ${section}.${key}: ${e.message}`);
        return defaultValue;
    }
}

/**
 * Funktion zum Aktualisieren von Benutzereinstellungen
 * @param {string} key - Konfigurationsschlüssel
 * @param {*} value - Neuer Wert
 * @returns {boolean} true bei Erfolg, false bei Fehler
 */
function updateUserConfig(key, value) {
    try {
        if (CONFIG.BENUTZER[key] !== undefined) {
            CONFIG.BENUTZER[key] = value;
            return true;
        }
        return false;
    } catch (e) {
        console.error(`Fehler beim Aktualisieren der Benutzereinstellung ${key}: ${e.message}`);
        return false;
    }
}