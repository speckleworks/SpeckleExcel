const url = require('url')
const axios = require('axios')

const Office = window.Office

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
        IsDefault: false
      }

      let accounts = Office.context.document.settings.get('accounts')
      if (accounts === null || accounts === undefined) {
        accounts = []
        account.IsDefault = true
      }
      accounts.push(account)
      Office.context.document.settings.set('accounts', accounts)
      Office.context.document.settings.saveAsync()

      window.location.reload()
    })

  dialog.close()
}

module.exports = {
  showAccountsPopup () {
    let browserPath = url.resolve(window.location.origin, `login.html`)

    Office.context.ui.displayDialogAsync(browserPath, {height: 80, width: 30, displayInIframe: true},
      asyncResult => {
        dialog = asyncResult.value
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processLoginToken)
      })
  },
  getAccounts () {
    let accounts = Office.context.document.settings.get('accounts')
    if (accounts === null || accounts === undefined) {
      accounts = []
    }
    return JSON.stringify(accounts)
  }
}
