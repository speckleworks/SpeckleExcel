const Office = window.Office
const Excel = window.Excel

module.exports = {
  addSender (args) {
    this.myClients.push(JSON.parse(args))
    Office.context.document.settings.set('clients', this.myClients)
    Office.context.document.settings.saveAsync()
  },
  getObjectSelection () {
    return Excel.run(function (context) {
      const range = context.workbook.getSelectedRange()
      range.load(['rowIndex', 'rowCount', 'columnIndex', 'columnCount', 'worksheet/name'])

      return context.sync()
        .then(function () {
          // Get headers
          let headerRange = range.worksheet.getRangeByIndexes(0, range.columnIndex, 1, range.columnCount)
          headerRange.load('values')

          return context.sync()
            .then(function () {
              let headerValues = headerRange.values[0]

              let rIndex = range.rowIndex
              let rCount = range.rowCount
              if (rIndex === 0) {
                rIndex = 1
                rCount--
              }

              let selectionRange = range.worksheet.getRangeByIndexes(rIndex, range.columnIndex, rCount, range.columnCount)
              selectionRange.load('values')

              return context.sync()
                .then(function () {
                  let selectionValues = selectionRange.values

                  let allObjects = []

                  for (let i = 0; i < selectionValues.length; i++) {
                    let obj = {}

                    for (let j = 0; j < selectionValues[i].length; j++) {
                      let v = selectionValues[i][j]
                      if (v === '' || headerValues[j] === '') {
                        continue
                      }
                      obj[headerValues[j]] = v
                    }

                    if (Object.keys(obj).length === 0) {
                      continue
                    }

                    if (!headerValues.includes('applicationId')) {
                      obj.applicationId = 'excel/' + range.worksheet.name + '!' + (parseInt(rIndex) + i).toString()
                    }

                    allObjects.push(obj)
                  }
                  return context.sync(allObjects)
                })
            })
        })
    })
  },
  addSelectionToSender (args) {
    let client = JSON.parse(args)
    let index = this.myClients.findIndex(x => x._id === client._id)

    if (index < 0) {
      return
    }

    this.getObjectSelection()
      .then(res => {
        let objects = this.myClients[index].objects
          .filter(x => !res.map(o => o.applicationId).includes(x.applicationId))

        if (objects === undefined) {
          objects = []
        }

        res.forEach(o => {
          objects.push(o)
        })

        this.myClients[index].objects = objects

        window.EventBus.$emit('update-client', JSON.stringify({
          _id: client._id,
          objects: objects
        }))
      })
  },
  removeSelectionFromSender (args) {
    let client = JSON.parse(args)
    let index = this.myClients.findIndex(x => x._id === client._id)

    if (index < 0) {
      return
    }

    this.getObjectSelection()
      .then(res => {
        let objects = this.myClients[index].objects
          .filter(x => !res.map(o => o.applicationId).includes(x.applicationId))

        if (objects === undefined) {
          objects = []
        }

        this.myClients[index].objects = objects

        window.EventBus.$emit('update-client', JSON.stringify({
          _id: client._id,
          objects: objects
        }))
      })
  },
  updateSender (args) {
    throw new Error(args)
  }
}
