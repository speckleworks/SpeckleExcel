const axios = require('axios')
const flatten = require('flat')
const convertNumToColumnLetter = require('./excel-helpers').convertNumToColumnLetter

const Office = window.Office
const Excel = window.Excel

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
  data.forEach(o => {
    o = flatten(o, {safe: true})
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
  data.forEach(o => {
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

  Excel.run(function (context) {
    let sheets = context.workbook.worksheets
    sheets.load('items/name')

    return context.sync()
      .then(function () {
        let sheet = null
        let sheetName = client.streamId.substring(0, 30)
        if (sheets.items.findIndex(x => x.name === sheetName) < 0) {
          sheet = context.workbook.worksheets.add(sheetName)
        } else {
          let sheetIndex = sheets.items.findIndex(x => x.name === sheetName)
          sheet = sheets.items[sheetIndex]
          sheet.getRange().clear()
        }
        return context.sync()
          .then(function () {
            let objectTable = sheet.tables.add(`A1:${convertNumToColumnLetter(headers.length)}1`)
            objectTable.getHeaderRowRange().values = [headers]

            objectTable.rows.add(null, arrayedData)

            if (Office.context.requirements.isSetSupported('ExcelApi', '1.2')) {
              sheet.getUsedRange().format.autofitColumns()
              sheet.getUsedRange().format.autofitRows()
            }

            sheet.activate()
            return context.sync()
              .then(function () {
                window.EventBus.$emit('update-client', JSON.stringify({
                  _id: client._id,
                  loading: false,
                  isLoadingIndeterminate: true,
                  loadingBlurb: `Done.`
                }))
              })
          })
      })
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
    Office.context.document.settings.set('clients', this.myClients)
    Office.context.document.settings.saveAsync()
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
        // TODO: Orchestrate this
        let ids = res.data.resource.objects.map(o => o._id)

        getObjects(client.account.RestApi, ids)
          .then(res => {
            window.EventBus.$emit('update-client', JSON.stringify({
              _id: client._id,
              loading: true,
              loadingBlurb: 'Preparing to write sheet...',
              objects: res
            }))

            createExcelSheetStream(client, res)

            this.myClients[index].objects = ids
            Office.context.document.settings.set('clients', this.myClients)
            Office.context.document.settings.saveAsync()
          })
          .catch(err => {
            console.log(err)
          })
      })
  }
}
