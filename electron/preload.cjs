const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("desktopBridge", {
  chooseWorkbook: () => ipcRenderer.invoke("workbook:choose"),
  readWorkbook: (filePath) => ipcRenderer.invoke("workbook:read", filePath),
  saveTemplate: (payload) => ipcRenderer.invoke("template:save", payload),
  loadTemplate: () => ipcRenderer.invoke("template:load"),
  listPrinters: () => ipcRenderer.invoke("system:list-printers"),
  printLabels: (payload) => ipcRenderer.invoke("system:print", payload || {})
});
