const path = require('path');

module.exports = {
  entry: './js/chat-agent.js',
  output: {
    filename: 'bundle.js',
    path: path.resolve(__dirname, 'dist'),
  },
  mode: 'development',
  module: {
    rules: [
      {
        test: /\.m?js$/,
        exclude: /bower_components/, // Only exclude bower_components
        use: {
          loader: 'babel-loader',
          options: {
            presets: ['@babel/preset-env'],
          },
        },
      },
    ],
  },
  resolve: {
    fallback: {
      "fs": false,
      "path": false,
      "os": false,
    },
  },
};