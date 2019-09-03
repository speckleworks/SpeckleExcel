const url = require('url')
const senderBindings = require('./sender-bindings')
const receiverBindings = require('./receiver-bindings')

const Office = window.Office
const Excel = window.Excel

module.exports = Object.assign({},
  {
    myClients: [],
    getApplicationHostName () {
      return 'Excel'
    },
    getFileName () {
      return 'MY FILE'
    },
    getDocumentId () {
      return 'TEST'
    },
    getDocumentLocation () {
      return 'COMP'
    },
    getFileClients () {
      this.myClients = Office.context.document.settings.get('clients')
      if (this.myClients === null || this.myClients === undefined) {
        this.myClients = []
      }
      return JSON.stringify(this.myClients)
    },
    removeClient (args) {
      let client = JSON.parse(args)
      let index = this.myClients.findIndex(x => x._id === client._id)
      if (index > -1) {
        this.myClients.splice(index, 1)
        Office.context.document.settings.set('clients', this.myClients)
        Office.context.document.settings.saveAsync()
      }
    },
    selectClientObjects (args) {
      let client = JSON.parse(args)

      Excel.run(function (context) {
        let sheets = context.workbook.worksheets
        sheets.load('items/name')

        return context.sync()
          .then(function () {
            let sheetIndex = sheets.items.findIndex(x => x.name === client.fullName)
            if (sheetIndex > -1) {
              sheets.items[sheetIndex].activate()
            }
          })
      })
    },
    showDev () {
      throw new Error('Not implemented')
    },
    showAccountsPopup () {
      let speckleServerUrl = 'https://hestia.speckle.works/api'
      let browserPath = url.resolve(speckleServerUrl.replace('api', ''), '/signin?redirectUrl=https://localhost:5050')

      window.open(browserPath, '_blank', 'toolbar=no,menubar=no,width=500,height=800')
    },
    getAccounts () {
      return JSON.stringify([
        {
          ServerName: 'Speckle Hestia',
          RestApi: 'https://hestia.speckle.works/api',
          Email: 'mishael.ebel.nuh@gmail.com',
          Token: 'XXX',
          IsDefault: 0
        }
      ])
    }
  },
  receiverBindings,
  senderBindings
)
