// jest.config.js
export default {
    testEnvironment: 'node',
    transform: {},
    moduleNameMapper: {
        '^src/(.*)$': '<rootDir>/src/$1'
    },
    testMatch: ['**/tests/**/*.test.js'],
    setupFilesAfterEnv: ['./jest.setup.js']
};