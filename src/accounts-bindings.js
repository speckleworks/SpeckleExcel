const url = require('url')
const axios = require('axios')

var dialog

function processLoginToken (arg) {
  let message = decodeURI(arg.message)
  let loginToken = message.split('%3A%3A%3A')[0]
  let serverUrl = message.split('%3A%3A%3A')[1].replace(/%2F/g, '/').replace(/%3A/g, ':')

  let speckleServerName = 'Speckle Server'

  axios.get(url.resolve(serverUrl, '/api'),
    {
      headers: {
        'Authorization': loginToken
      }
    })
    .then(res => {
      speckleServerName = res.data.serverName

      return axios.get(url.resolve(serverUrl, '/api/accounts'),
        {
          headers: {
            'Authorization': loginToken
          }
        })
    })
    .then(res => {
      let account = {
        ServerName: speckleServerName,
        RestApi: url.resolve(serverUrl, '/api'),
        Email: res.data.resource.email,
        Token: res.data.resource.apitoken,
        AccountId: res.data.resource._id,
        IsDefault: false
      }

      let accounts = window.Office.context.document.settings.get('accounts')
      if (accounts === null || accounts === undefined) {
        accounts = []
      }

      let accIndex = accounts.findIndex(x => x.AccountId === account.AccountId)
      if (accIndex > -1) {
        accounts.splice(accIndex, 1)
      }

      if (accounts.length === 0) {
        account.IsDefault = true
      }

      accounts.push(account)

      window.Office.context.document.settings.set('accounts', accounts)
      window.Office.context.document.settings.saveAsync()

      window.Store.dispatch('getAccounts')
    })

  dialog.close()
}

module.exports = {
  showAccountsPopup () {
    let browserPath = url.resolve(window.location.origin, `login.html`)

    window.Office.context.ui.displayDialogAsync(browserPath, {height: 50, width: 30, displayInIframe: true},
      asyncResult => {
        dialog = asyncResult.value
        dialog.addEventHandler(window.Office.EventType.DialogMessageReceived, processLoginToken)
      })
  },
  getAccounts () {
    let accounts = window.Office.context.document.settings.get('accounts')
    if (accounts === null || accounts === undefined) {
      accounts = []
    }
    return JSON.stringify(accounts)
  }
}
