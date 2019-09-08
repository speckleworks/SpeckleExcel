const axios = require('axios')
const unflatten = require('flat').unflatten

const Office = window.Office
const Excel = window.Excel

function getObjects (objects) {
  return Excel.run(function (context) {
    let sheets = context.workbook.worksheets
    sheets.load('items/name')

    return context.sync({context: context, sheets: sheets})
  })
    .then(function ({context, sheets}) {
      let worksheetRanges = {}
      let goodObjects = []
      objects.forEach(o => {
        if (!worksheetRanges.hasOwnProperty(o.sheet)) {
          let sheetIndex = sheets.items.findIndex(x => x.name === o.sheet)
          if (sheetIndex > -1) {
            worksheetRanges[o.sheet] = sheets.items[sheetIndex].getUsedRange(true)
            worksheetRanges[o.sheet].load('values')
            goodObjects.push(o)
          }
        } else {
          goodObjects.push(o)
        }
      })

      return context.sync({context: context, worksheetRanges: worksheetRanges, goodObjects: goodObjects})
    })
    .then(function ({context, worksheetRanges, goodObjects}) {
      let convertedObjects = []

      goodObjects.forEach(o => {
        let convObj = {}
        let vals = worksheetRanges[o.sheet].values
        for (let i = 0; i < vals[o.row].length; i++) {
          let header = vals[0][i]
          let v = vals[o.row][i]

          if (v === '' || header === '' || header === '_id' || header === 'hash') {
            continue
          }

          try {
            convObj[header] = JSON.parse(v)
          } catch (ex) {
            convObj[header] = v
          }
        }

        if (Object.keys(convObj).length > 0) {
          if (!Object.keys(convObj).includes('applicationId')) {
            convObj.applicationId = 'excel/' + o.sheet + '!' + (o.row).toString()
          }

          convertedObjects.push(unflatten(convObj, {safe: true}))
        }
      })

      return context.sync(convertedObjects)
    })
}

module.exports = {
  addSender (args) {
    this.myClients.push(JSON.parse(args))
    Office.context.document.settings.set('clients', this.myClients)
    Office.context.document.settings.saveAsync()
  },
  getObjectSelection () {
    return Excel.run(function (context) {
      const range = context.workbook.getSelectedRange()
      range.load(['rowIndex', 'rowCount', 'worksheet/name'])

      return context.sync(range)
    })
      .then(function (range) {
        let objects = []
        for (let i = range.rowIndex; i < range.rowIndex + range.rowCount; i++) {
          if (i === 0) {
            continue
          }

          objects.push({
            sheet: range.worksheet.name,
            row: i
          })
        }

        return range.context.sync(objects)
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

        if (objects === undefined) {
          objects = []
        }

        res.forEach(o => {
          objects = objects.filter(x => !(x.sheet === o.sheet && x.row === o.row))
          if (objects === undefined) {
            objects = []
          }
          objects.push(o)
        })

        this.myClients[index].objects = objects

        window.EventBus.$emit('update-client', JSON.stringify({
          _id: client._id,
          objects: objects
        }))

        Office.context.document.settings.set('clients', this.myClients)
        Office.context.document.settings.saveAsync()
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

        if (objects === undefined) {
          objects = []
        }

        res.forEach(o => {
          objects = objects.filter(x => !(x.sheet === o.sheet && x.row === o.row))
        })

        this.myClients[index].objects = objects

        window.EventBus.$emit('update-client', JSON.stringify({
          _id: client._id,
          objects: objects
        }))
      })
  },
  updateSender (args) {
    let client = JSON.parse(args)
    let index = this.myClients.findIndex(x => x._id === client._id)

    if (index < 0) {
      return
    }

    window.EventBus.$emit('update-client', JSON.stringify({
      _id: client._id,
      loading: true,
      loadingBlurb: 'Converting objects...'
    }))

    let objects = this.myClients[index].objects

    if (objects === undefined) {
      objects = []
    }

    getObjects(objects)
      .then(res => {
        window.EventBus.$emit('update-client', JSON.stringify({
          _id: client._id,
          loading: true,
          loadingBlurb: 'Sending to stream...'
        }))

        axios.defaults.headers.common[ 'Authorization' ] = client.account.Token

        let promises = []
        // TODO: Do orchestration here
        promises.push(
          axios({
            method: 'POST',
            baseURL: client.account.RestApi,
            url: `objects`,
            data: res
          })
            .then(axRes => { return axRes.data.resources })
        )

        return Promise.all(promises)
      })
      .then(res => {
        let placeholders = []

        res.forEach(r => {
          r.forEach(p => {
            placeholders.push(p)
          })
        })

        return axios({
          method: 'PUT',
          baseURL: client.account.RestApi,
          url: `streams/${client.streamId}`,
          data: {name: client.name, objects: placeholders}
        })
      })
      .then(() => {
        window.EventBus.$emit('update-client', JSON.stringify({
          _id: client._id,
          loading: false,
          loadingBlurb: 'Done.'
        }))
      })
  }
}
