// eslint.config.js
import eslint from '@eslint/js';

export default [
    eslint.configs.recommended,
    {
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: 'module',
            globals: {
                // Google Apps Script globals
                SpreadsheetApp: 'readonly',
                DriveApp: 'readonly',
                Utilities: 'readonly',
                Session: 'readonly',
                ScriptApp: 'readonly',
                global: 'writable',
                console: 'readonly',
                // Browser and Node environments
                window: 'readonly',
                document: 'readonly',
                require: 'readonly',
                module: 'writable',
                process: 'readonly',
            },
        },
        files: ['**/*.js'],
        ignores: [
            'node_modules/**',
            'dist/**',
            'coverage/**',
        ],
        rules: {
            // Error prevention
            'no-unused-vars': ['warn', { argsIgnorePattern: '^_' }],
            'no-undef': 'error',
            'no-console': ['warn', { allow: ['error', 'warn', 'info'] }],

            // Style consistency
            'indent': ['warn', 4],
            'quotes': ['warn', 'single', { avoidEscape: true }],
            'semi': ['warn', 'always'],
            'comma-dangle': ['warn', 'always-multiline'],

            // Best practices
            'no-var': 'warn',
            'prefer-const': 'warn',
            'no-param-reassign': 'warn',
            'eqeqeq': ['error', 'always', { null: 'ignore' }],

            // Google Apps Script specific
            'no-restricted-globals': ['error',
                { name: 'setTimeout', message: 'Use Utilities.sleep() instead' },
                { name: 'setInterval', message: 'Not available in Google Apps Script' },
            ],
        },
    },
    {
        files: ['**/*.test.js'],
        languageOptions: {
            globals: {
                jest: true,
                describe: true,
                it: true,
                expect: true,
                beforeEach: true,
                afterEach: true,
            },
        },
    },
];