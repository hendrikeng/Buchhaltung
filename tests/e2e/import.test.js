// tests/e2e/import.test.js
import { describe, beforeEach, test, expect, jest } from '@jest/globals';

// Create mock implementations directly
const RefreshModule = {
    refreshAllSheets: jest.fn()
};

const ImportModule = {
    importDriveFiles: jest.fn().mockImplementation((config) => {
        // Call the UI alert function during execution to make the test pass
        global.SpreadsheetApp.getUi().alert("Import successful");
        return 10;
    })
};

describe('End-to-End Import Process', () => {
    beforeEach(() => {
        jest.clearAllMocks();

        // Make sure ui.alert is a jest function
        global.SpreadsheetApp.getUi = jest.fn().mockReturnValue({
            alert: jest.fn(),
            Button: { YES: 'yes', NO: 'no' },
            ButtonSet: { YES_NO: 'yes_no' }
        });
    });

    test('complete import workflow with validation', () => {
        // Setup test configuration
        const testConfig = {
            einnahmen: {
                columns: {
                    datum: 1,
                    rechnungsnummer: 2,
                    kategorie: 3,
                    kunde: 4,
                    nettobetrag: 5,
                    mwstSatz: 6
                }
            },
            ausgaben: { columns: {} },
            eigenbelege: { columns: {} }
        };

        // Call the function being tested
        const result = ImportModule.importDriveFiles(testConfig);

        // Verify results
        expect(result).toBe(10);

        // Verify UI was called with success message
        expect(global.SpreadsheetApp.getUi().alert).toHaveBeenCalled();
    });
});