// tests/modules/importModule/integration.test.js
import { jest, describe, beforeEach, test, expect } from '@jest/globals';
import ImportModule from '../../../src/modules/importModule/index.js';
import fileManager from '../../../src/modules/importModule/fileManager.js';
import dataProcessor from '../../../src/modules/importModule/dataProcessor.js';
import historyTracker from '../../../src/modules/importModule/historyTracker.js';

// Mock der Module
jest.mock('../../../src/modules/importModule/fileManager.js');
jest.mock('../../../src/modules/importModule/dataProcessor.js');
jest.mock('../../../src/modules/importModule/historyTracker.js');

describe('ImportModule Integration Tests', () => {
    beforeEach(() => {
        jest.clearAllMocks();

        // Setup für getActiveSpreadsheet
        const mockSheets = {
            'Einnahmen': { name: 'Einnahmen' },
            'Ausgaben': { name: 'Ausgaben' },
            'Eigenbelege': { name: 'Eigenbelege' }
        };

        SpreadsheetApp.getActiveSpreadsheet.mockReturnValue({
            getSheetByName: jest.fn(name => mockSheets[name] || null)
        });
    });

    test('successful import workflow', () => {
        // Konfiguration
        const mockConfig = {
            einnahmen: { columns: {} },
            ausgaben: { columns: {} },
            eigenbelege: { columns: {} }
        };

        // Mock Objekte
        const mockParentFolder = { name: 'Parent' };
        const mockRevenueFolder = { name: 'Einnahmen' };
        const mockExpenseFolder = { name: 'Ausgaben' };
        const mockReceiptsFolder = { name: 'Eigenbelege' };
        const mockHistorySheet = { name: 'Änderungshistorie' };
        const mockExistingFiles = new Set();

        // Mocks konfigurieren
        fileManager.getParentFolder.mockReturnValue(mockParentFolder);
        fileManager.findOrCreateFolder
            .mockReturnValueOnce(mockRevenueFolder)
            .mockReturnValueOnce(mockExpenseFolder)
            .mockReturnValueOnce(mockReceiptsFolder);

        historyTracker.getOrCreateHistorySheet.mockReturnValue(mockHistorySheet);
        historyTracker.collectExistingFiles.mockReturnValue(mockExistingFiles);

        dataProcessor.importFilesFromFolder
            .mockReturnValueOnce(5) // 5 Einnahmen
            .mockReturnValueOnce(3) // 3 Ausgaben
            .mockReturnValueOnce(2); // 2 Eigenbelege

        // Funktionsaufruf
        const result = ImportModule.importDriveFiles(mockConfig);

        // Verifizierung des Ergebnisses
        expect(result).toBe(10); // Summe aller importierten Dateien (5+3+2)

        // Verifizierung der UI-Meldung
        expect(SpreadsheetApp.getUi().alert).toHaveBeenCalledWith(
            expect.stringContaining("Import abgeschlossen")
        );
    });
});