'use strict'
const utils = require('./utils')
const config = require('../config')
const isProduction = process.env.NODE_ENV === 'production'
const sourceMapEnabled = isProduction
  ? config.build.productionSourceMap
  : config.dev.cssSourceMap

module.exports = {
  loaders: {
    // css: utils.cssLoaders({
    //   sourceMap: sourceMapEnabled,
    //   extract: isProduction
    // }),
    js: [{
      loader: 'babel-loader',
      options: {
        presets: [
          ["env", {
            modules: false,
            targets: {
              "browsers": ["> 1%", "last 2 versions", "not ie <= 8"]
            }
          }],
          "vue"
        ],
        plugins: [
          "transform-object-rest-spread",
          "transform-es2015-shorthand-properties",
        ]
      }
    }]
  },
  cssSourceMap: sourceMapEnabled,
  cacheBusting: config.dev.cacheBusting,
  transformToRequire: {
    video: ['src', 'poster'],
    source: 'src',
    img: 'src',
    image: 'xlink:href'
  }
}
