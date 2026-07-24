const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

module.exports = async (_env, argv) => {
  const serving = Boolean(argv.env?.WEBPACK_SERVE);
  const devServer = {
    host: "127.0.0.1",
    port: 3000,
    hot: true,
    headers: { "Access-Control-Allow-Origin": "*" },
  };

  if (serving) {
    const devCerts = require("office-addin-dev-certs");
    devServer.server = {
      type: "https",
      options: await devCerts.getHttpsServerOptions(),
    };
  }

  return {
    entry: { taskpane: "./src/taskpane/taskpane.ts" },
    devtool: argv.mode === "development" ? "source-map" : false,
    resolve: { extensions: [".ts", ".js"] },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: "ts-loader",
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          { from: "manifest.xml", to: "manifest.xml" },
        ],
      }),
    ],
    output: {
      clean: true,
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
    },
    devServer,
  };
};
