const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
    entry: './src/taskpane/taskpane.js',
    mode: 'development',
    devtool: 'source-map',
    resolve: {
        extensions: ['.js', '.html', '.css']
    },
    module: {
        rules: [
            {
                test: /\.css$/,
                use: ['style-loader', 'css-loader']
            }
        ]
    },
    plugins: [
        new HtmlWebpackPlugin({
            template: './src/taskpane/taskpane.html',
            filename: 'taskpane.html',
            inject: false
        }),
        new CopyWebpackPlugin({
            patterns: [
                {
                    from: './manifest.xml',
                    to: './manifest.xml'
                }
            ]
        })
    ],
    devServer: {
        static: {
            directory: path.join(__dirname, 'dist'),
        },
        compress: true,
        port: 3000,
        server: {
            type: 'https',
            options: {
                // Allow self-signed certificates
            }
        },
        allowedHosts: 'all',
        headers: {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, PATCH, OPTIONS',
            'Access-Control-Allow-Headers': 'X-Requested-With, content-type, Authorization'
        }
    },
    output: {
        filename: 'taskpane.js',
        path: path.resolve(__dirname, 'dist'),
        clean: true
    }
};