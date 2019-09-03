import bindings from './ui-bindings'

window.UiBindings = bindings

import('../SpeckleUiApp/src/main')
.then(() => {
  const Office = window.Office
  Office.initialize = function () {
    return window.app
  }
})