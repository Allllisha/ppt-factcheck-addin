/* eslint-disable no-undef */

const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

module.exports = {
  mode: "production",
  devtool: "source-map",
  entry: {
    polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
    taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
    commands: "./src/commands/commands.js",
  },
  output: {
    clean: true,
    path: __dirname + "/dist",
    publicPath: "/"
  },
  resolve: {
    extensions: [".html", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: {
          loader: "babel-loader",
        },
      },
      {
        test: /\.html$/,
        exclude: /node_modules/,
        use: "html-loader",
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
    new webpack.DefinePlugin({
      'process.env.JINA_API_TOKEN': JSON.stringify(process.env.JINA_API_TOKEN),
      'process.env.TAVILY_API_KEY': JSON.stringify(process.env.TAVILY_API_KEY),
      'process.env.GOOGLE_API_KEY': JSON.stringify(process.env.GOOGLE_API_KEY),
      'process.env.GOOGLE_SEARCH_ENGINE_ID': JSON.stringify(process.env.GOOGLE_SEARCH_ENGINE_ID),
    }),
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane/taskpane.html",
      chunks: ["polyfill", "taskpane"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "assets/*",
          to: "assets/[name][ext][query]",
        },
        {
          from: "src/taskpane/taskpane.css",
          to: "taskpane.css",
        },
      ],
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["polyfill", "commands"],
    }),
  ],
};