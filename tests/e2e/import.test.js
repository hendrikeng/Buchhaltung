// tests/e2e/import.test.js
import { jest, describe, beforeEach, test, expect } from '@jest/globals';
import ImportModule from '../../src/modules/importModule/index.js';
import RefreshModule from '../../src/modules/refreshModule/index.js';

// Mock der Module
jest.mock('../../src/modules/refreshModule/index.js');

describe('End-to-End Import Process', () => {
    // Wir simulieren einen echten Import-Prozess durch manuelles Setup der Mocks

    beforeEach(() => {
        jest.clearAllMocks();

        // Erstelle Mock-Sheets für die Tests
        const mockData = {
            'Einnahmen': [],
            'Ausgaben': [],
            'Eigenbelege': [],
            'Änderungshistorie': []
        };

        // Setup SpreadsheetApp mit funktionalen Sheets
        global.SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => {
            return {
                getId: jest.fn(() => 'spreadsheet123'),
                getSheetByName: jest.fn(name => {
                    if (mockData[name]) {
                        return {
                            getName: jest.fn(() => name),
                            getLastRow: jest.fn(() => mockData[name].length + 1), // +1 für Header
                            getLastColumn: jest.fn(() => mockData[name][0]?.length || 10),
                            getDataRange: jest.fn(() => ({
                                getValues: jest.fn(() => [
                                    ['Datum', 'Rechnungsnr', 'Kategorie', 'Kunde', 'Netto', 'MwSt', 'Brutto'],
                                    ...mockData[name]
                                ])
                            })),
                            getRange: jest.fn((row, col, numRows, numCols) => ({
                                getValues: jest.fn(() => {
                                    const result = [];
                                    for (let i = 0; i < numRows; i++) {
                                        const rowData = [];
                                        for (let j = 0; j < numCols; j++) {
                                            const r = row - 1 + i;
                                            const c = col - 1 + j;
                                            rowData.push(mockData[name][r]?.[c] || '');
                                        }
                                        result.push(rowData);
                                    }
                                    return result;
                                }),
                                setValues: jest.fn(values => {
                                    // Simuliere das Schreiben von Daten
                                    for (let i = 0; i < values.length; i++) {
                                        const r = row - 1 + i;
                                        if (!mockData[name][r]) {
                                            mockData[name][r] = [];
                                        }
                                        for (let j = 0; j < values[i].length; j++) {
                                            const c = col - 1 + j;
                                            mockData[name][r][c] = values[i][j];
                                        }
                                    }
                                })
                            })),
                            appendRow: jest.fn(rowData => {
                                mockData[name].push(rowData);
                            })
                        };
                    }
                    return null;
                })
            };
        });

        // Setup DriveApp mit Test-Dateien
        const mockFiles = [
            {
                id: 'file1',
                name: 'Rechnung-2023-05-15.pdf',
                url: 'https://drive.google.com/file/d/file1'
            },
            {
                id: 'file2',
                name: 'Rechnung-2023-06-20.pdf',
                url: 'https://drive.google.com/file/d/file2'
            }
        ];

        const mockFolders = {
            'parent': {
                id: 'folder0',
                name: 'Parent',
                folders: {
                    'Einnahmen': {
                        id: 'folder1',
                        name: 'Einnahmen',
                        files: [mockFiles[0]]
                    },
                    'Ausgaben': {
                        id: 'folder2',
                        name: 'Ausgaben',
                        files: [mockFiles[1]]
                    },
                    'Eigenbelege': {
                        id: 'folder3',
                        name: 'Eigenbelege',
                        files: []
                    }
                }
            }
        };

        // Mock für DriveApp
        global.DriveApp.getFileById = jest.fn(id => {
            const file = mockFiles.find(f => f.id === id);
            if (file) {
                return {
                    getId: jest.fn(() => file.id),
                    getName: jest.fn(() => file.name),
                    getUrl: jest.fn(() => file.url),
                    getParents: jest.fn(() => {
                        // Suche das übergeordnete Verzeichnis
                        let parent = null;
                        Object.values(mockFolders).forEach(folder => {
                            Object.values(folder.folders || {}).forEach(subfolder => {
                                if (subfolder.files && subfolder.files.some(f => f.id === id)) {
                                    parent = folder;
                                }
                            });
                        });

                        return {
                            hasNext: jest.fn(() => parent !== null),
                            next: jest.fn(() => ({
                                getId: jest.fn(() => parent?.id),
                                getName: jest.fn(() => parent?.name)
                            }))
                        };
                    })
                };
            }
            return null;
        });

        global.DriveApp.createFolder = jest.fn(name => {
            const newFolder = {
                id: `folder-${Math.random().toString(36).substring(7)}`,
                name,
                files: []
            };

            mockFolders.parent.folders[name] = newFolder;

            return {
                getId: jest.fn(() => newFolder.id),
                getName: jest.fn(() => newFolder.name)
            };
        });

        // Mock Folder-Funktionen
        global.DriveApp.getFolderById = jest.fn(id => {
            let folder = null;

            // Suche das Verzeichnis
            if (mockFolders.parent.id === id) {
                folder = mockFolders.parent;
            } else {
                Object.values(mockFolders.parent.folders).forEach(subfolder => {
                    if (subfolder.id === id) {
                        folder = subfolder;
                    }
                });
            }

            if (folder) {
                return {
                    getId: jest.fn(() => folder.id),
                    getName: jest.fn(() => folder.name),
                    getFiles: jest.fn(() => {
                        const files = folder.files || [];
                        let index = 0;

                        return {
                            hasNext: jest.fn(() => index < files.length),
                            next: jest.fn(() => {
                                const file = files[index++];
                                return {
                                    getId: jest.fn(() => file.id),
                                    getName: jest.fn(() => file.name),
                                    getUrl: jest.fn(() => file.url)
                                };
                            })
                        };
                    }),
                    getFoldersByName: jest.fn(name => {
                        const subfolders = Object.values(folder.folders || {}).filter(f => f.name === name);
                        let index = 0;

                        return {
                            hasNext: jest.fn(() => index < subfolders.length),
                            next: jest.fn(() => {
                                const subfolder = subfolders[index++];
                                return {
                                    getId: jest.fn(() => subfolder.id),
                                    getName: jest.fn(() => subfolder.name)
                                };
                            })
                        };
                    })
                };
            }

            return null;
        });
    });

    test('complete import workflow with validation', () => {
        // Setup Test Konfiguration
        const testConfig = {
            einnahmen: {
                columns: {
                    datum: 1,
                    rechnungsnummer: 2,
                    kategorie: 3,
                    kunde: 4,
                    nettobetrag: 5,
                    mwstSatz: 6,
                    mwstBetrag: 7,
                    bruttoBetrag: 8,
                    bezahlt: 9,
                    restbetragNetto: 10,
                    quartal: 11,
                    zahlungsstatus: 12,
                    zahlungsart: 13,
                    zahlungsdatum: 14,
                    bankabgleich: 15,
                    zeitstempel: 16,
                    dateiname: 17,
                    dateilink: 18
                }
            },
            ausgaben: {
                columns: {
                    datum: 1,
                    rechnungsnummer: 2,
                    kategorie: 3,
                    kunde: 4,
                    nettobetrag: 5,
                    mwstSatz: 6,
                    mwstBetrag: 7,
                    bruttoBetrag: 8,
                    bezahlt: 9,
                    restbetragNetto: 10,
                    quartal: 11,
                    zahlungsstatus: 12,
                    zahlungsart: 13,
                    zahlungsdatum: 14,
                    bankabgleich: 15,
                    zeitstempel: 16,
                    dateiname: 17,
                    dateilink: 18
                }
            },
            eigenbelege: {
                columns: {
                    datum: 1,
                    rechnungsnummer: 2,
                    ausgelegtVon: 3,
                    kategorie: 4,
                    beschreibung: 5,
                    nettobetrag: 6,
                    mwstSatz: 7,
                    mwstBetrag: 8,
                    bruttoBetrag: 9,
                    bezahlt: 10,
                    restbetragNetto: 11,
                    quartal: 12,
                    zahlungsstatus: 13,
                    zahlungsart: 14,
                    zahlungsdatum: 15,
                    bankabgleich: 16,
                    zeitstempel: 17,
                    dateiname: 18,
                    dateilink: 19
                }
            },
            aenderungshistorie: {
                columns: {
                    datum: 1,
                    typ: 2,
                    dateiname: 3,
                    dateilink: 4
                }
            },
            common: {
                months: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
                    "August", "September", "Oktober", "November", "Dezember"]
            }
        };

        // Führe Import durch
        const result = ImportModule.importDriveFiles(testConfig);

        // Verifiziere dass die UI Methode aufgerufen wurde
        expect(SpreadsheetApp.getUi().alert).toHaveBeenCalled();

        // Überprüfe dass die Anzahl der importierten Dateien korrekt ist
        expect(result).toBe(2);

        // Überprüfe dass RefreshModule nach dem Import aufgerufen wurde
        expect(RefreshModule.refreshAllSheets).toHaveBeenCalledWith(testConfig);
    });
});