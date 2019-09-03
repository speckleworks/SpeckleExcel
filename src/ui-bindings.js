module.exports = {
  myClients: [],
  addObjectsToSender(args) {

  },
  addReceiver(args) {
    this.myClients.push(JSON.parse(args))
  },
  addSender(args) {

  },
  bakeReceiver(args) {

  },
  getApplicationHostName() {
    return "Excel"
  },
  getFileName() {
    return "MY FILE"
  },
  getDocumentId() {
    return "TEST"
  },
  getDocumentLocation() {
    return "COMP"
  },
  getFileClients() {
    return JSON.stringify(this.myClients)
  },
  removeObjectsFromSender(args) {

  },
  removeClient(args) {
    let client = JSON.parse(args)
    let index = this.myClients.findIndex(x => x._id === client._id)
    if (index > -1)
    {
      this.myClients.splice(index, 1)
    }
  },
  addSelectionToSender(args) {
    
  },
  removeSelectionFromSender(args) {
    
  },
  updateSender(args) {

  },
  selectClientObjects(args) {

  },
  showDev() {
    
  },
  showAccountsPopup() {

  },
  getAccounts() {

  },
}