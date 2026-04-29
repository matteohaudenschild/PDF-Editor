const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("desktopAPI", {
  getServiceBaseUrl: () => ipcRenderer.invoke("service:get-base-url"),
  getAppVersion: () => ipcRenderer.invoke("app:get-version"),
  checkForUpdates: () => ipcRenderer.invoke("updates:check"),
  installUpdate: () => ipcRenderer.invoke("updates:install"),
  openPdfDialog: () => ipcRenderer.invoke("dialog:open-pdf"),
  savePdfDialog: (suggestedName) => ipcRenderer.invoke("dialog:save-pdf", suggestedName),
  revealFile: (targetPath) => ipcRenderer.invoke("shell:reveal-file", targetPath),
});
