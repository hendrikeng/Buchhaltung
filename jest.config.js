// jest.config.js
export default {
    testEnvironment: 'node',
    transform: {},
    moduleNameMapper: {
        '^(\\.{1,2}/.*)\\.js$': '$1',
    },
    testMatch: ['**/tests/**/*.test.js'],
    setupFilesAfterEnv: ['./jest.setup.js'],
    collectCoverageFrom: [
        'src/modules/importModule/**/*.js',
        '!src/**/index.js'
    ]
};