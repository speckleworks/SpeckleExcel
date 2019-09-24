const axios = require('axios')
const unflatten = require('flat').unflatten

function sendObjects (baseUrl, objects) {
  return new Promise((resolve, reject) => {
    axios({
      method: 'POST',
      baseURL: baseUrl,
      url: 'objects',
      data: objects
    })
      .then(res => resolve(res.data.resources))
      .catch(err => reject(err))
  })
}

function getObjects (objects) {
  return window.Excel.run(function (context) {
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

        if (!vals[o.row] || !vals[0]) {
          return
        }

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

          if (!Object.keys(convObj).includes('type')) {
            let typeMatch = Object.keys(convObj).findIndex(x => x.toLowerCase() === 'type')
            if (typeMatch > -1) {
              let type = Object.keys(convObj)[typeMatch]
              convObj.type = convObj[type]
              delete convObj[type]
            } else {
              convObj.type = 'Object'
            }
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
    this.addSelectionToSender(args)
    window.Office.context.document.settings.set('clients', this.myClients)
    window.Office.context.document.settings.saveAsync()
  },
  getObjectSelection () {
    return window.Excel.run(function (context) {
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

        window.Office.context.document.settings.set('clients', this.myClients)
        window.Office.context.document.settings.saveAsync()
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

        let bucket = []
        let maxReq = 100 // magic number; maximum objects to send in a bucket

        for (let i = 0; i < res.length; i++) {
          bucket.push(res[i])
          if (i % maxReq === 0 && i !== 0) {
            promises.push(sendObjects(client.account.RestApi, bucket.slice()))
            bucket = []
          }
        }

        if (bucket.length !== 0) {
          promises.push(sendObjects(client.account.RestApi, bucket.slice()))
          bucket = []
        }

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
      .catch(err => {
        window.EventBus.$emit('update-client', JSON.stringify({
          _id: client._id,
          loading: false,
          isLoadingIndeterminate: true,
          loadingBlurb: `Unable to send stream.`,
          errors: JSON.stringify(err)
        }))
      })
  }
}
