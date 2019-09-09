module.exports = {
  presets: [
    ["@babel/preset-env", {
      modules: false
    }],
  ],
  plugins: [
    "@vue/babel-plugin-transform-vue-jsx",
    "@babel/plugin-transform-runtime",
    "@babel/plugin-proposal-object-rest-spread",
    "@babel/plugin-transform-shorthand-properties",
    "@babel/plugin-transform-arrow-functions",
    "@babel/plugin-syntax-dynamic-import",
  ]
}
