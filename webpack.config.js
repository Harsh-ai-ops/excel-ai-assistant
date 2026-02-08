const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, argv) => {
    const isDevelopment = argv.mode === 'development';

    return {
        entry: {
            taskpane: './src/taskpane/index.tsx',
        },

        output: {
            filename: '[name].bundle.js',
            path: path.resolve(__dirname, 'dist'),
            clean: true,
        },

        resolve: {
            extensions: ['.ts', '.tsx', '.js', '.jsx'],
        },

        module: {
            rules: [
                {
                    test: /\.tsx?$/,
                    use: 'ts-loader',
                    exclude: /node_modules/,
                },
                {
                    test: /\.css$/,
                    use: ['style-loader', 'css-loader'],
                },
            ],
        },

        plugins: [
            new HtmlWebpackPlugin({
                template: './src/taskpane/index.html',
                filename: 'taskpane.html',
                chunks: ['taskpane'],
            }),

            new CopyWebpackPlugin({
                patterns: [
                    { from: 'assets', to: 'assets' },
                    { from: 'manifest.xml', to: 'manifest.xml' },
                    { from: 'install.html', to: 'install.html' },
                ],
            }),
        ],

        devServer: {
            port: 3000,
            https: true, // Required for Office Add-ins
            hot: true,
            headers: {
                'Access-Control-Allow-Origin': '*',
            },
            static: {
                directory: path.join(__dirname, 'dist'),
            },
        },

        devtool: isDevelopment ? 'source-map' : false,
    };
};
