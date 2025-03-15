// jest.setup.js
import { jest } from '@jest/globals';
import { setupGoogleAppsMocks } from './tests/modules/importModule/mocks/googleAppsMocks.js';

// Make jest available globally
global.jest = jest;

// Setup Google Apps Script mocks before tests run
setupGoogleAppsMocks();