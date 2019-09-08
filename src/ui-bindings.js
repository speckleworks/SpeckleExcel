const senderBindings = require('./sender-bindings')
const receiverBindings = require('./receiver-bindings')
const accountsBindings = require('./accounts-bindings')

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

      // TODO: Figure out how to do this elegantly for senders
      if (client.type === 'receiver') {
        Excel.run(function (context) {
          let sheets = context.workbook.worksheets
          sheets.load('items/name')

          return context.sync()
            .then(function () {
              let sheetIndex = sheets.items.findIndex(x => x.name === client.streamId)
              if (sheetIndex > -1) {
                sheets.items[sheetIndex].activate()
              }
            })
        })
      }
    },
    showDev () {
      throw new Error('Not implemented')
    }
  },
  receiverBindings,
  senderBindings,
  accountsBindings
)
