module.exports = {
  addSender (args) {
    this.myClients.push(JSON.parse(args))
  },
  addObjectsToSender (args) {
    throw new Error(args)
  },
  removeObjectsFromSender (args) {
    throw new Error(args)
  },
  addSelectionToSender (args) {
    throw new Error(args)
  },
  removeSelectionFromSender (args) {
    throw new Error(args)
  },
  updateSender (args) {
    throw new Error(args)
  }
}
