// tests/modules/importModule/dataProcessor.test.js
import { jest, describe, beforeEach, it, expect } from '@jest/globals';
import dataProcessor from '../../src/modules/importModule/dataProcessor.js';
import dateUtils from '../../src/utils/dateUtils.js';
import sheetUtils from '../../src/utils/sheetUtils.js';

// Mocks für die abhängigen Module
jest.mock('../../../src/utils/dateUtils.js', () => ({
    extractDateFromFilename: jest.fn()
}));

jest.mock('../../../src/utils/sheetUtils.js', () => ({
    batchWriteToSheet: jest.fn()
}));

describe('dataProcessor', () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    describe('importFilesFromFolder', () => {
        it('should import files from folder and update sheets', () => {
            // Einrichten der Mocks
            const mockFiles = [
                { getName: () => 'Invoice-2021-01-15.pdf', getUrl: () => 'https://example.com/file1' },
                { getName: () => 'Receipt-2021-02-20.pdf', getUrl: () => 'https://example.com/file2' }
            ];

            let fileIndex = 0;
            const mockFolder = {
                getFiles: jest.fn(() => ({
                    hasNext: jest.fn(() => fileIndex < mockFiles.length),
                    next: jest.fn(() => mockFiles[fileIndex++])
                }))
            };

            const mockMainSheet = {
                getLastRow: jest.fn(() => 5),
                getLastColumn: jest.fn(() => 10)
            };

            const mockHistorySheet = {
                getLastRow: jest.fn(() => 2),
                getLastColumn: jest.fn(() => 4)
            };

            const mockExistingFiles = new Set();

            const mockConfig = {
                einnahmen: {
                    columns: {
                        datum: 1,
                        rechnungsnummer: 2,
                        zeitstempel: 16,
                        dateiname: 17,
                        dateilink: 18
                    }
                },
                aenderungshistorie: {
                    columns: {
                        datum: 1,
                        typ: 2,
                        dateiname: 3,
                        dateilink: 4
                    }
                }
            };

            // Mock dateUtils.extractDateFromFilename
            dateUtils.extractDateFromFilename.mockImplementation((filename) => {
                if (filename.includes('2021-01-15')) return new Date(2021, 0, 15);
                if (filename.includes('2021-02-20')) return new Date(2021, 1, 20);
                return null;
            });

            // Test ausführen
            const result = dataProcessor.importFilesFromFolder(
                mockFolder,
                mockMainSheet,
                'Einnahme',
                mockHistorySheet,
                mockExistingFiles,
                mockConfig
            );

            // Überprüfungen
            expect(result).toBe(2); // 2 Dateien importiert
            expect(mockFolder.getFiles).toHaveBeenCalled();
            expect(dateUtils.extractDateFromFilename).toHaveBeenCalledTimes(2);
            expect(sheetUtils.batchWriteToSheet).toHaveBeenCalledTimes(2); // Einmal für Main, einmal für History
            expect(mockExistingFiles.size).toBe(2);
        });

        it('should skip already imported files', () => {
            // Einrichten der Mocks
            const mockFiles = [
                { getName: () => 'Invoice-2021-01-15.pdf', getUrl: () => 'https://example.com/file1' },
                { getName: () => 'Receipt-2021-02-20.pdf', getUrl: () => 'https://example.com/file2' }
            ];

            let fileIndex = 0;
            const mockFolder = {
                getFiles: jest.fn(() => ({
                    hasNext: jest.fn(() => fileIndex < mockFiles.length),
                    next: jest.fn(() => mockFiles[fileIndex++])
                }))
            };

            const mockMainSheet = {
                getLastRow: jest.fn(() => 5),
                getLastColumn: jest.fn(() => 10)
            };

            const mockHistorySheet = {
                getLastRow: jest.fn(() => 2),
                getLastColumn: jest.fn(() => 4)
            };

            // Bereits importierte Dateien
            const mockExistingFiles = new Set(['Invoice-2021-01-15']);

            const mockConfig = {
                einnahmen: {
                    columns: {
                        datum: 1,
                        rechnungsnummer: 2,
                        zeitstempel: 16,
                        dateiname: 17,
                        dateilink: 18
                    }
                },
                aenderungshistorie: {
                    columns: {
                        datum: 1,
                        typ: 2,
                        dateiname: 3,
                        dateilink: 4
                    }
                }
            };

            // Mock dateUtils.extractDateFromFilename
            dateUtils.extractDateFromFilename.mockImplementation((filename) => {
                if (filename.includes('2021-02-20')) return new Date(2021, 1, 20);
                return null;
            });

            // Test ausführen
            const result = dataProcessor.importFilesFromFolder(
                mockFolder,
                mockMainSheet,
                'Einnahme',
                mockHistorySheet,
                mockExistingFiles,
                mockConfig
            );

            // Überprüfungen
            expect(result).toBe(1); // Nur 1 neue Datei importiert
            expect(dateUtils.extractDateFromFilename).toHaveBeenCalledTimes(1);
        });
    });
});