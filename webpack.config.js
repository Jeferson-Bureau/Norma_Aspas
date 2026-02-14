const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

module.exports = async (env, options) => {
    const dev = options.mode === "development";
    const config = {
        devtool: "source-map",
        entry: {
            "apa-quote-formatter": "./apa-quote-formatter.ts",
        },
        output: {
            clean: true,
            filename: "[name].js", // This will likely produce taskpane.js, but the html expects apa-quote-formatter.js?
            // Wait, let's just use one entry for now named 'apa-quote-formatter'
        },
        resolve: {
            extensions: [".ts", ".tsx", ".html", ".js"],
        },
        module: {
            rules: [
                {
                    test: /\.ts$/,
                    exclude: /node_modules/,
                    use: {
                        loader: "ts-loader",
                    },
                },
                {
                    test: /\.(png|jpg|jpeg|gif|ico)$/,
                    type: "asset/resource",
                    generator: {
                        filename: "assets/[name][ext][query]",
                    },
                },
            ],
        },
        plugins: [
            new CopyWebpackPlugin({
                patterns: [
                    {
                        from: "manifest.xml",
                        to: "manifest.xml",
                        transform(content) {
                            if (dev) {
                                return content;
                            } else {
                                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
                            }
                        },
                    },
                    // Copy assets folder if it exists, otherwise we'll rely on the rule above or creating it
                ],
            }),
            new HtmlWebpackPlugin({
                filename: "taskpane.html",
                template: "./taskpane.html",
                chunks: ["apa-quote-formatter"],
            }),
            new HtmlWebpackPlugin({
                filename: "commands.html",
                template: "./commands.html",
                chunks: ["apa-quote-formatter"],
            }),
        ],
        devServer: {
            headers: {
                "Access-Control-Allow-Origin": "*",
            },
            server: {
                type: "https",
                options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await devCerts.getHttpsServerOptions(),
            },
            port: 3000,
            hot: false, // Office Add-ins often prefer full reload often
        },
    };

    // Fix entry name to match expected output
    config.entry = {
        "apa-quote-formatter": "./apa-quote-formatter.ts"
    };

    return config;
};
