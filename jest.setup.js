// jest.setup.js
import { jest } from '@jest/globals';

// Make jest available globally
global.jest = jest;

// Setup Google Apps Script mocks
global.DriveApp = {
    getFileById: jest.fn(),
    createFolder: jest.fn()
};

global.SpreadsheetApp = {
    getActiveSpreadsheet: jest.fn(() => ({
        getId: jest.fn(() => 'spreadsheet123'),
        getSheetByName: jest.fn()
    })),
    getUi: jest.fn(() => ({
        alert: jest.fn(),
        Button: { YES: 'yes', NO: 'no' },
        ButtonSet: { YES_NO: 'yes_no' }
    }))
};

global.Utilities = {
    formatDate: jest.fn(),
    sleep: jest.fn()
};

global.Session = {
    getScriptTimeZone: jest.fn(() => 'Europe/Berlin')
};

// Add console.error mock
global.console.error = jest.fn();