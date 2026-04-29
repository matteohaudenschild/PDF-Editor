const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("desktopAPI", {
  getServiceBaseUrl: () => ipcRenderer.invoke("service:get-base-url"),
  openPdfDialog: () => ipcRenderer.invoke("dialog:open-pdf"),
  savePdfDialog: (suggestedName) => ipcRenderer.invoke("dialog:save-pdf", suggestedName),
  revealFile: (targetPath) => ipcRenderer.invoke("shell:reveal-file", targetPath),
});
