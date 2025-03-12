# Projektübersicht: Buchhaltungs-App für Google Sheets

## Projekt-Zusammenfassung

Sie arbeiten an einer umfassenden Buchhaltungslösung für kleine Unternehmen, die auf Google Sheets und Google Apps Script basiert. Die Anwendung ist speziell für den deutschen Markt konzipiert und unterstützt sowohl operative GmbHs als auch Holding-Strukturen mit einem SKR04-konformen Kontenrahmen.

## Kernfunktionen

1. **Automatisierte Dokumentenverwaltung**
    - Import von Rechnungen aus Google Drive-Ordnern
    - Intelligente Erkennung von Datums- und Rechnungsinformationen aus Dateinamen
    - Protokollierung aller Dokumente in einer Änderungshistorie

2. **Steuerliche Auswertungen**
    - Automatische Berechnung der Umsatzsteuervoranmeldung (UStVA)
    - Unterstützung deutscher MwSt-Sätze (0%, 7%, 19%)
    - Berücksichtigung von steuerfreien Inland- und Auslandsumsätzen

3. **Betriebswirtschaftliche Auswertung (BWA)**
    - Monatliche und quartalsweise Aufbereitung wichtiger Kennzahlen
    - Differenzierte Darstellung von Umsatz, Kosten und Ertrag
    - Automatische Berechnung von EBIT und weiteren Finanzkennzahlen

4. **Bilanzerstellung**
    - Automatische Generierung einer SKR04-konformen Bilanz
    - Korrekte Darstellung von Aktiva und Passiva
    - Validierung der Bilanzidentität

5. **Bank-Abgleich**
    - Automatische Zuordnung von Bankbewegungen zu Rechnungen
    - Farbliche Kennzeichnung von bezahlten und offenen Posten
    - Tracking des Zahlungsstatus für Einnahmen und Ausgaben

6. **Validierung und Fehlerbehandlung**
    - Umfangreiche Validierung aller Eingaben
    - Detaillierte Fehlermeldungen bei Problemen
    - Robuste Fehlerbehandlung und Logging

## Architektur

Die Anwendung ist modular aufgebaut und besteht aus folgenden Hauptkomponenten:

1. **ImportModule**: Importiert und kategorisiert Dokumente aus Google Drive
2. **RefreshModule**: Aktualisiert Sheets und berechnet Formeln neu
3. **UStVACalculator**: Berechnet Umsatzsteuervoranmeldungen
4. **BWACalculator**: Erstellt betriebswirtschaftliche Auswertungen
5. **BilanzCalculator**: Generiert eine SKR04-konforme Bilanz
6. **Validator**: Validiert Eingaben und Berechnungen
7. **Helpers**: Bietet Hilfsfunktionen für Datum, Währung, etc.
8. **Config**: Zentrale Konfiguration mit Kontenplan und Steuereinstellungen

## Sheet-Struktur

Die Anwendung verwendet folgende Tabellenblätter:

- **Einnahmen**: Erfassung aller Einnahmen mit MwSt-Berechnung
- **Ausgaben**: Erfassung aller Ausgaben mit Vorsteuer-Berechnung
- **Eigenbelege**: Dokumentation von Belegen ohne formale Rechnung
- **Bankbewegungen**: Abgleich von Banktransaktionen mit Rechnungen
- **Gesellschafterkonto**: Verwaltung von Transaktionen mit Gesellschaftern
- **Holding Transfers**: Spezielle Buchungen in Holding-Strukturen
- **Änderungshistorie**: Protokollierung aller Importe und Änderungen
- **UStVA**: Generierte Umsatzsteuervoranmeldung mit monatlichen und quartalsweisen Daten
- **BWA**: Generierte betriebswirtschaftliche Auswertung
- **Bilanz**: Generierte Bilanz nach SKR04

## Weiterentwicklungspotenzial

Basierend auf dem vorhandenen Code könnten folgende Erweiterungen sinnvoll sein:

1. **Dashboard mit Kennzahlen**
    - Visualisierung wichtiger Finanzkennzahlen
    - Cashflow-Prognose basierend auf offenen Posten
    - Grafische Darstellung der Umsatz- und Gewinnentwicklung

2. **Automatisiertes Reporting**
    - Regelmäßige E-Mail-Berichte an Geschäftsführung
    - Export von Auswertungen als PDF oder XLSX
    - Vordefinierte Berichte für Steuerberater

3. **Erweitertes Dokumenten-Management**
    - OCR-Integration zur automatischen Datenextraktion aus Rechnungen
    - Verknüpfung mit Cloud-Speichern (Dropbox, OneDrive)
    - Automatische Kategorisierung von Dokumenten mit KI

4. **Mehrwährungsunterstützung**
    - Automatische Währungsumrechnung
    - Historische Wechselkurse für korrekte Buchhaltung
    - Bilanzierung in verschiedenen Währungen

5. **Compliance-Verbesserungen**
    - Integration aktueller Steuerrichtlinien
    - GoBD-konforme Datenspeicherung
    - Prüfroutinen für steuerliche Sonderregelungen

6. **API-Integration**
    - Anbindung an Banking-APIs für automatischen Kontoabgleich
    - Schnittstelle zu DATEV oder anderen Buchhaltungssystemen
    - Integration von Zahlungsdienstleistern für automatischen Abgleich

## Fazit

Die Buchhaltungs-App bietet bereits eine solide Basis für die Finanzverwaltung kleiner Unternehmen und ist speziell auf den deutschen Steuerkontext zugeschnitten. Mit modularem Aufbau und klaren Schnittstellen ist sie gut für zukünftige Erweiterungen gerüstet. Die Kombination aus automatisiertem Dokumentenimport, steuerkonformer Auswertung und Bilanzierung bietet einen erheblichen Mehrwert gegenüber herkömmlichen Tabellenkalkulationslösungen.