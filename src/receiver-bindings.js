const axios = require('axios')
const flatten = require('flat')
const convertNumToColumnLetter = require('./excel-helpers').convertNumToColumnLetter

function getObjects (baseUrl, objectIds) {
  return new Promise((resolve, reject) => {
    axios({
      method: 'POST',
      baseURL: baseUrl,
      url: 'objects/getbulk',
      data: objectIds
    })
      .then(res => resolve(res.data.resources.filter(x => !(x.type === 'String' && x.value === 'You do not have permissions to view this object'))))
      .catch(err => reject(err))
  })
}

function createExcelSheetStream (client, data) {
  let headers = []
  let counter = 1

  let flattenedObjects = []

  data.forEach(o => {
    o = flatten(o, {safe: true})
    flattenedObjects.push(o)

    Object.keys(o).forEach(k => {
      headers.push(k)
    })

    window.EventBus.$emit('update-client', JSON.stringify({
      _id: client._id,
      loading: true,
      isLoadingIndeterminate: false,
      loadingProgress: 100 * counter / data.length,
      loadingBlurb: `Flattening objects: ${counter} / ${data.length}`
    }))
    counter++
  })
  headers = headers
    .filter(function (item, i, ar) { return ar.indexOf(item) === i })
    .filter(h => !['private', 'canRead', 'canWrite', 'anonymousComments', 'comments', 'deleted', 'owner', '__v', 'createdAt', 'updatedAt'].includes(h))

  let arrayedData = []
  counter = 1
  flattenedObjects.forEach(o => {
    let newObj = []
    headers.forEach(h => {
      if (o.hasOwnProperty(h)) {
        let val = JSON.stringify(o[h])
        try {
          newObj.push(val.replace(/"/g, ''))
        } catch (ex) {
          newObj.push(val)
        }
      } else {
        newObj.push('')
      }
    })
    arrayedData.push(newObj)

    window.EventBus.$emit('update-client', JSON.stringify({
      _id: client._id,
      loading: true,
      isLoadingIndeterminate: false,
      loadingProgress: (100 * counter) / data.length,
      loadingBlurb: `Excel-ifying objects: ${counter} / ${data.length}`
    }))
    counter++
  })

  window.EventBus.$emit('update-client', JSON.stringify({
    _id: client._id,
    loading: true,
    isLoadingIndeterminate: true,
    loadingBlurb: `Writing sheet...`
  }))

  window.Excel.run(function (context) {
    let sheets = context.workbook.worksheets
    sheets.load('items/name')

    return context.sync({context: context, sheets: sheets})
  })
    .then(function ({context, sheets}) {
      let sheet = null
      let sheetName = client.streamId.substring(0, 30)
      if (sheets.items.findIndex(x => x.name === sheetName) < 0) {
        sheet = context.workbook.worksheets.add(sheetName)
      } else {
        let sheetIndex = sheets.items.findIndex(x => x.name === sheetName)
        sheet = sheets.items[sheetIndex]
        sheet.getRange().clear()
      }
      return context.sync({context: context, sheet: sheet})
    })
    .then(function ({context, sheet}) {
      if (arrayedData.length > 0) {
        let objectTable = sheet.tables.add(`A1:${convertNumToColumnLetter(headers.length)}1`)
        objectTable.style = 'TableStyleLight8'
        objectTable.getHeaderRowRange().values = [headers]

        objectTable.rows.add(null, arrayedData)
      }

      if (window.Office.context.requirements.isSetSupported('ExcelApi', '1.2')) {
        sheet.getUsedRange().format.autofitColumns()
        sheet.getUsedRange().format.autofitRows()
      }

      sheet.activate()
      return context.sync()
    })
    .then(function () {
      window.EventBus.$emit('update-client', JSON.stringify({
        _id: client._id,
        loading: false,
        isLoadingIndeterminate: true,
        loadingBlurb: `Done.`
      }))
    })
    .catch(err => {
      window.EventBus.$emit('update-client', JSON.stringify({
        _id: client._id,
        loading: false,
        isLoadingIndeterminate: true,
        loadingBlurb: `Unable to receive stream.`,
        errors: JSON.stringify(err)
      }))
    })
}

module.exports = {
  addReceiver (args) {
    this.myClients.push(JSON.parse(args))
    window.Office.context.document.settings.set('clients', this.myClients)
    window.Office.context.document.settings.saveAsync()
  },
  bakeReceiver (args) {
    let client = JSON.parse(args)
    let index = this.myClients.findIndex(x => x._id === client._id)

    if (index < 0) {
      return
    }

    window.EventBus.$emit('update-client', JSON.stringify({
      _id: client._id,
      loading: true,
      loadingBlurb: 'Getting stream from server...'
    }))

    axios.defaults.headers.common[ 'Authorization' ] = client.account.Token

    axios({
      method: 'GET',
      baseURL: client.account.RestApi,
      url: `streams/${client.streamId}?fields=objects,layers`
    })
      .then(res => {
        let ids = res.data.resource.objects.map(o => o._id)

        let promises = []

        let bucket = []
        let maxReq = 50 // magic number; maximum objects to request in a bucket

        for (let i = 0; i < ids.length; i++) {
          bucket.push(ids[i])
          if (i % maxReq === 0 && i !== 0) {
            promises.push(getObjects(client.account.RestApi, bucket.slice()))
            bucket = []
          }
        }

        if (bucket.length !== 0) {
          promises.push(getObjects(client.account.RestApi, bucket.slice()))
          bucket = []
        }

        return Promise.all(promises)
      })
      .then(res => {
        let objects = []

        res.forEach(arr => {
          arr.forEach(o => {
            objects.push(o)
          })
        })

        window.EventBus.$emit('update-client', JSON.stringify({
          _id: client._id,
          loading: true,
          loadingBlurb: 'Preparing to write sheet...',
          objects: objects
        }))

        createExcelSheetStream(client, objects)

        this.myClients[index].objects = objects.map(x => x._id)
        window.Office.context.document.settings.set('clients', this.myClients)
        window.Office.context.document.settings.saveAsync()
      })
      .catch(err => {
        window.EventBus.$emit('update-client', JSON.stringify({
          _id: client._id,
          loading: false,
          isLoadingIndeterminate: true,
          loadingBlurb: `Unable to receive stream.`,
          errors: JSON.stringify(err)
        }))
      })
  }
}
