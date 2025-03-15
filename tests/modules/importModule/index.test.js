// tests/modules/importModule/index.test.js
import { jest, describe, beforeEach, it, expect } from '@jest/globals';
import ImportModule from '../../../src/modules/importModule/index.js';
import fileManager from '../../../src/modules/importModule/fileManager.js';
import dataProcessor from '../../src/modules/importModule/dataProcessor.js';
import historyTracker from '../../../src/modules/importModule/historyTracker.js';

// Mocks für die abhängigen Module
jest.mock('../../../src/modules/importModule/fileManager.js', () => ({
    getParentFolder: jest.fn(),
    findOrCreateFolder: jest.fn()
}));

jest.mock('../../../src/modules/importModule/dataProcessor.js', () => ({
    importFilesFromFolder: jest.fn()
}));

jest.mock('../../../src/modules/importModule/historyTracker.js', () => ({
    getOrCreateHistorySheet: jest.fn(),
    collectExistingFiles: jest.fn()
}));

describe('ImportModule', () => {
    beforeEach(() => {
        jest.clearAllMocks();

        // Setup für getActiveSpreadsheet
        SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ({
            getSheetByName: jest.fn(name => {
                if (['Einnahmen', 'Ausgaben', 'Eigenbelege'].includes(name)) {
                    return { name };
                }
                return null;
            })
        }));
    });

    describe('importDriveFiles', () => {
        it('should import files from all folders', () => {
            // Einrichten der Mocks
            const mockParentFolder = { name: 'Parent' };
            const mockRevFolder = { name: 'Einnahmen' };
            const mockExpFolder = { name: 'Ausgaben' };
            const mockRecFolder = { name: 'Eigenbelege' };
            const mockHistorySheet = { name: 'Änderungshistorie' };
            const mockExistingFiles = new Set();

            fileManager.getParentFolder.mockReturnValue(mockParentFolder);
            fileManager.findOrCreateFolder
                .mockReturnValueOnce(mockRevFolder)
                .mockReturnValueOnce(mockExpFolder)
                .mockReturnValueOnce(mockRecFolder);

            historyTracker.getOrCreateHistorySheet.mockReturnValue(mockHistorySheet);
            historyTracker.collectExistingFiles.mockReturnValue(mockExistingFiles);

            dataProcessor.importFilesFromFolder
                .mockReturnValueOnce(3) // 3 Einnahmen
                .mockReturnValueOnce(2) // 2 Ausgaben
                .mockReturnValueOnce(1); // 1 Eigenbeleg

            const mockConfig = {
                einnahmen: { columns: {} },
                ausgaben: { columns: {} },
                eigenbelege: { columns: {} }
            };

            // Test ausführen
            const result = ImportModule.importDriveFiles(mockConfig);

            // Überprüfungen
            expect(fileManager.getParentFolder).toHaveBeenCalled();
            expect(fileManager.findOrCreateFolder).toHaveBeenCalledTimes(3);
            expect(historyTracker.getOrCreateHistorySheet).toHaveBeenCalled();
            expect(historyTracker.collectExistingFiles).toHaveBeenCalled();
            expect(dataProcessor.importFilesFromFolder).toHaveBeenCalledTimes(3);
            expect(result).toBe(6); // Insgesamt 6 Dateien importiert
        });

        it('should handle case when parent folder is not found', () => {
            // Einrichten der Mocks
            fileManager.getParentFolder.mockReturnValue(null);

            const mockConfig = {};
            const ui = SpreadsheetApp.getUi();

            // Test ausführen
            const result = ImportModule.importDriveFiles(mockConfig);

            // Überprüfungen
            expect(fileManager.getParentFolder).toHaveBeenCalled();
            expect(ui.alert).toHaveBeenCalledWith("Fehler: Kein übergeordneter Ordner gefunden.");
            expect(result).toBe(0);
        });
    });
});