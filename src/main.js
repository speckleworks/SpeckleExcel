import bindings from './ui-bindings'

function handleSelectionChange () {
  const Excel = window.Excel
  return Excel.run(function (context) {
    return context.sync(context)
  })
    .then(function (context) {
      const range = context.workbook.getSelectedRange()
      range.load(['rowIndex', 'rowCount'])

      return context.sync({context: context, range: range})
    })
    .then(function ({context, range}) {
      window.EventBus.$emit('update-selection-count', JSON.stringify({
        selectedObjectsCount: range.rowIndex === 0 ? range.rowCount - 1 : range.rowCount
      }))
      return context.sync()
    })
}

window.UiBindings = bindings

import('../SpeckleUiApp/src/main')
  .then(() => {
    const Office = window.Office
    Office.initialize = function () {
      handleSelectionChange()
      return window.app
    }

    const Excel = window.Excel
    Excel.run(function (context) {
      context.workbook.onSelectionChanged.add(handleSelectionChange)

      return context.sync()
    })
  })
