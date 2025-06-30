const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const webpack = require('webpack');

module.exports = (env, argv) => {
    const isProduction = argv.mode === 'production';
    
    const apiUrl = process.env.BASE_API_URL || 'https://backend-962119591036.europe-west1.run.app';
    
    return {
        entry: './src/taskpane/taskpane.js',
        mode: isProduction ? 'production' : 'development',
        devtool: isProduction ? 'source-map' : 'eval-source-map',
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
        new webpack.DefinePlugin({
            'process.env.BASE_API_URL': JSON.stringify(apiUrl)
        }),
        new HtmlWebpackPlugin({
            template: './src/taskpane/taskpane.html',
            filename: 'taskpane.html',
            inject: 'body'
        }),
        new CopyWebpackPlugin({
            patterns: [
                {
                    from: './manifest.xml',
                    to: './manifest.xml'
                },
                {
                    from: './assets',
                    to: './assets',
                    noErrorOnMissing: true
                }
            ]
        })
    ],
    devServer: {
        static: path.join(__dirname, 'dist'),
        port: 3000,
        open: true
    },
        output: {
            filename: 'taskpane.js',
            path: path.resolve(__dirname, 'dist'),
            publicPath: './',
            clean: true
        }
    };
};