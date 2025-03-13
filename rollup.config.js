// rollup.config.js
import resolve from '@rollup/plugin-node-resolve';
import rollupPluginGas from 'rollup-plugin-google-apps-script';
import terser from '@rollup/plugin-terser';

export default {
    input: 'src/code.js',
    output: {
        file: 'dist/code.js',
        format: 'iife',
        banner: '/* Bundled Code for Google Apps Script */',
    },
    plugins: [
        resolve(),
        rollupPluginGas(),
        terser({
                format: {
                    comments: false,
                },
                compress: {
                    drop_console: false,
                    drop_debugger: true,
                },
                mangle: {
                    reserved: ['global']
                }
            }
        )
    ]
};
