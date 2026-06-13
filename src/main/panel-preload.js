const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('panelApi', {
  getSnapshot: () => ipcRenderer.invoke('panel:get-snapshot'),
  sendCommand: (command) => ipcRenderer.send('panel:command', command),
  onStatusUpdated: (callback) => {
    ipcRenderer.on('panel:status-updated', (_event, snapshot) => callback(snapshot))
  },
  onEvent: (callback) => {
    ipcRenderer.on('panel:event', (_event, data) => callback(data))
  },
  onFindResult: (callback) => {
    ipcRenderer.on('panel:event', (_event, data) => {
      if (data && data.type === 'find-result') callback(data)
    })
  },
  onDownloadsUpdated: (callback) => {
    ipcRenderer.on('panel:downloads-updated', (_event, data) => callback(data))
  }
})
