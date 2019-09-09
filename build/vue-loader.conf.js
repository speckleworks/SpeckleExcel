'use strict'
const utils = require('./utils')
const config = require('../config')
const babelConfig = require('./babel.config')
const isProduction = process.env.NODE_ENV === 'production'
const sourceMapEnabled = isProduction
  ? config.build.productionSourceMap
  : config.dev.cssSourceMap

module.exports = {
  loaders: {
    css: utils.cssLoaders({
      sourceMap: sourceMapEnabled,
      extract: isProduction
    }),
    js: [{
      loader: 'babel-loader',
      options: babelConfig,
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
