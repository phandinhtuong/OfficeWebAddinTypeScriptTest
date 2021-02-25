var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var _this = this;
var devCerts = require("office-addin-dev-certs");
var CleanWebpackPlugin = require("clean-webpack-plugin");
var CopyWebpackPlugin = require("copy-webpack-plugin");
var CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
var HtmlWebpackPlugin = require("html-webpack-plugin");
var webpack = require("webpack");
var urlDev = "https://localhost:3000/";
var urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION
module.exports = function (env, options) { return __awaiter(_this, void 0, void 0, function () {
    var dev, buildType, config, _a;
    var _b, _c;
    return __generator(this, function (_d) {
        switch (_d.label) {
            case 0:
                dev = options.mode === "development";
                buildType = dev ? "dev" : "prod";
                _b = {
                    devtool: "source-map",
                    entry: {
                        functions: "./src/functions/functions.ts",
                        polyfill: "@babel/polyfill",
                        taskpane: "./src/taskpane/taskpane.ts",
                        commands: "./src/commands/commands.ts"
                    },
                    resolve: {
                        extensions: [".ts", ".tsx", ".html", ".js"]
                    },
                    module: {
                        rules: [
                            {
                                test: /\.ts$/,
                                exclude: /node_modules/,
                                use: "babel-loader"
                            },
                            {
                                test: /\.tsx?$/,
                                exclude: /node_modules/,
                                use: "ts-loader"
                            },
                            {
                                test: /\.html$/,
                                exclude: /node_modules/,
                                use: "html-loader"
                            },
                            {
                                test: /\.(png|jpg|jpeg|gif)$/,
                                loader: "file-loader",
                                options: {
                                    name: '[path][name].[ext]',
                                }
                            }
                        ]
                    },
                    plugins: [
                        new CleanWebpackPlugin({
                            cleanOnceBeforeBuildPatterns: dev ? [] : ["**/*"]
                        }),
                        new CustomFunctionsMetadataPlugin({
                            output: "functions.json",
                            input: "./src/functions/functions.ts"
                        }),
                        new HtmlWebpackPlugin({
                            filename: "functions.html",
                            template: "./src/functions/functions.html",
                            chunks: ["polyfill", "functions"]
                        }),
                        new HtmlWebpackPlugin({
                            filename: "taskpane.html",
                            template: "./src/taskpane/taskpane.html",
                            chunks: ["polyfill", "taskpane"]
                        }),
                        new CopyWebpackPlugin([
                            {
                                to: "taskpane.css",
                                from: "./src/taskpane/taskpane.css"
                            },
                            {
                                to: "[name]." + buildType + ".[ext]",
                                from: "manifest*.xml",
                                transform: function (content) {
                                    if (dev) {
                                        return content;
                                    }
                                    else {
                                        return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
                                    }
                                }
                            }
                        ]),
                        new HtmlWebpackPlugin({
                            filename: "commands.html",
                            template: "./src/commands/commands.html",
                            chunks: ["polyfill", "commands"]
                        })
                    ]
                };
                _c = {
                    headers: {
                        "Access-Control-Allow-Origin": "*"
                    }
                };
                if (!(options.https !== undefined)) return [3 /*break*/, 1];
                _a = options.https;
                return [3 /*break*/, 3];
            case 1: return [4 /*yield*/, devCerts.getHttpsServerOptions()];
            case 2:
                _a = _d.sent();
                _d.label = 3;
            case 3:
                config = (_b.devServer = (_c.https = _a,
                    _c.port = process.env.npm_package_config_dev_server_port || 3000,
                    _c),
                    _b);
                return [2 /*return*/, config];
        }
    });
}); };
//# sourceMappingURL=webpack.config.js.map