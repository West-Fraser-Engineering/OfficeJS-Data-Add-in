const path = require('path');
const CustomFunctionsMetadataPlugin = require('custom-functions-metadata-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const HtmlBundlerPlugin = require('html-bundler-webpack-plugin');

module.exports = [
    {
        name: 'client',
        mode: 'production',
        optimization: {
            minimize: false
        },
        // entry: {
        //     functions: './src/functions/functions.ts'
        // },
        output: {
            path: path.resolve(__dirname, 'build'),
            clean: true,
            // filename: '[name].js',
        },
        resolve: {
            alias: {
                '@src': path.resolve(__dirname, 'src'), // Add this alias
            },
            extensions: ['.js', '.ts', '.tsx'],
            symlinks: false,
            fallback: {
                fs: require.resolve("browserify-fs"), // or 'empty' if you prefer an empty module
                crypto: require.resolve('crypto-browserify'),
                stream: require.resolve('stream-browserify'),
            }
        },
        module: {
            rules: [
                {
                    test: /\.ts$/,
                    exclude: /node_modules/,
                    include: path.resolve(__dirname, 'src'),
                    use: 'ts-loader',
                },
                {
                    test: /\.css$/i,
                    use: ['css-loader'],
                    include: path.resolve(__dirname, 'src'),
                    exclude: /node_modules/,
                },
                {
                    test: /\.css$/,
                    include: path.resolve(__dirname, 'src'),
                    use: ['style-loader', 'css-loader'],
                },
                {
                    test: /\.(png|svg|jpg|jpeg|gif)$/i,
                    type: 'asset/resource'
                },
                {
                    test: /\.wasm$/,
                    type: 'javascript/auto',
                },
                // {
                //     test: /\.html$/,
                //     include: path.resolve(__dirname, 'src'),
                //     loader: 'html-loader',
                // },
            ],
        },
        externals: {
            office: "office-js"
        },
        node: false,
        plugins: [
            // new CopyWebpackPlugin({
            //     patterns: [
            //         { from: '**/*.html', to: '', context: 'src/' },
            //         { from: '**/*.css', to: '', context: 'src/' },
            //     ]
            // }),
            new CustomFunctionsMetadataPlugin({
                output: 'functions.json',
                input: './src/functions/functions.ts'
            }),
            new HtmlBundlerPlugin({
                entry: 'src/',
                js: {
                    filename: 'js/[name].[contenthash:8].js'
                },
                css: {
                    filename: 'css/[name].[contenthash:8].css'
                }
            })
        ],
    },
    {
        name: 'test',
        mode: 'development',
        entry: {
            tests: './test/index.test.ts'
        },
        resolve: {
            alias: {
                '@src': path.resolve(__dirname, 'src'),
            },
            extensions: ['.js', '.ts', '.tsx'],
            symlinks: false,
        },
        module: {
            rules: [
                {
                    test: /\.tsx?$/,
                    use: 'ts-loader',
                    exclude: /node_modules/,
                },
            ],
        },
        output: {
            path: path.resolve(__dirname, 'test-build'),
            filename: 'tests.js',
        },
        node: false,
        externals: {
            office: "office-js"
        },
    }
];