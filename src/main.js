// The Vue build version to load with the `import` command
// (runtime-only or standalone) has been set in webpack.base.conf with an alias.
import Vue from 'vue'
import '../SpeckleUiApp/src/main'

Vue.config.productionTip = false

const Office = window.Office
Office.initialize = function () {
  return window.app
}
