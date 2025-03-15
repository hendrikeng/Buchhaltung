// tests/modules/importModule/mocks/googleAppsMocks.js
export const setupGoogleAppsMocks = () => {
    // Mock für DriveApp
    global.DriveApp = {
        getFileById: () => jest.fn(),
        createFolder: () => jest.fn()
    };

    // Mock für SpreadsheetApp
    global.SpreadsheetApp = {
        getActiveSpreadsheet: () => ({
            getId: () => 'spreadsheet123',
            getSheetByName: () => jest.fn(),
            setActiveSheet: () => jest.fn()
        }),
        getUi: () => ({
            alert: () => jest.fn(),
            Button: { YES: 'yes', NO: 'no' },
            ButtonSet: { YES_NO: 'yes_no' }
        })
    };

    // Mock für Utilities
    global.Utilities = {
        formatDate: () => '01.01.2023', // Default test date
        sleep: () => jest.fn()
    };

    // Mock für Session
    global.Session = {
        getScriptTimeZone: () => 'Europe/Berlin'
    };

    // Mock für Logger
    global.console.error = jest.fn();
};