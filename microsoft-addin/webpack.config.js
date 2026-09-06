/**
 * Webpack Configuration for Gemini for Microsoft 365
 * 
 * @author Sathya AG, Principal Architect, Google
 */

/* eslint-disable no-undef */

const fs = require("fs");
const path = require("path");
const webpack = require("webpack");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://gemini-frontend-16933400417.us-central1.run.app/"; // Production Cloud Run deployment endpoint

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const manifestWifSrc = fs.existsSync(path.resolve(__dirname, "manifest-wif.xml"))
    ? "manifest-wif.xml"
    : (fs.existsSync(path.resolve(__dirname, "../manifest-wif.xml")) ? path.resolve(__dirname, "../manifest-wif.xml") : null);

  const manifestGsuiteSrc = fs.existsSync(path.resolve(__dirname, "manifest-gsuite.xml"))
    ? "manifest-gsuite.xml"
    : (fs.existsSync(path.resolve(__dirname, "../manifest-gsuite.xml")) ? path.resolve(__dirname, "../manifest-gsuite.xml") : null);

  const manifestDeployedSrc = fs.existsSync(path.resolve(__dirname, "manifest-deployed.xml"))
    ? "manifest-deployed.xml"
    : (fs.existsSync(path.resolve(__dirname, "../manifest-deployed.xml")) ? path.resolve(__dirname, "../manifest-deployed.xml") : null);

  const copyPatterns = [
    {
      from: "Dockerfile",
      to: "Dockerfile",
      toType: "file",
    },
    {
      from: "nginx.conf",
      to: "nginx.conf",
      toType: "file",
    },
    {
      from: "assets/*",
      to: "assets/[name][ext][query]",
    },
    {
      from: "architecture.html",
      to: "architecture.html",
      toType: "file",
      noErrorOnMissing: true,
    },
  ];

  if (manifestWifSrc) {
    copyPatterns.push({
      from: manifestWifSrc,
      to: "manifest-wif.xml",
      toType: "file",
      noErrorOnMissing: true,
      transform(content) {
        if (dev) {
          return content;
        } else {
          return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
        }
      },
    });
  }

  if (manifestGsuiteSrc) {
    copyPatterns.push({
      from: manifestGsuiteSrc,
      to: "manifest-gsuite.xml",
      toType: "file",
      noErrorOnMissing: true,
      transform(content) {
        if (dev) {
          return content;
        } else {
          return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
        }
      },
    });
  }

  if (manifestDeployedSrc) {
    copyPatterns.push({
      from: manifestDeployedSrc,
      to: "manifest-deployed.xml",
      toType: "file",
      noErrorOnMissing: true,
      transform(content) {
        if (dev) {
          return content;
        } else {
          return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
        }
      },
    });
  }

  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
    },
    output: {
      clean: true,
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
        "process.env.GEMINI_PROXY_URL": JSON.stringify(process.env.GEMINI_PROXY_URL || "")
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
        hash: true,
      }),
      new CopyWebpackPlugin({
        patterns: copyPatterns,
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "google-auth.html",
        template: "./src/auth/google-auth.html",
        chunks: [],
      }),
      new HtmlWebpackPlugin({
        filename: "google-callback.html",
        template: "./src/auth/google-callback.html",
        chunks: [],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
