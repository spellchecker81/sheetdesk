/* SheetDesk preload script.
 * Runs with contextIsolation enabled and nodeIntegration disabled.
 * Exposes a minimal, deliberate API surface to the renderer. */
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  /** OS platform string, e.g. 'darwin', 'win32', 'linux'. */
  platform: process.platform,

  /**
   * Show a Save dialog and write `data` to the chosen path.
   * Resolves to { path } on success, or null when the dialog is cancelled.
   * `opts` is optional: { title, filters } to override the dialog defaults.
   */
  saveFileAs: (defaultName, data, opts) =>
    ipcRenderer.invoke('save-file', defaultName, data, opts),

  /**
   * Write `data` to `path` with no dialog.
   * Resolves to { path } on success.
   */
  saveFileTo: (path, data) =>
    ipcRenderer.invoke('save-file-to', path, data),

  /**
   * Show an Open dialog and read the chosen .sheetdesk.json file.
   * Resolves to { path, data } on success, or null when cancelled.
   */
  openFile: () =>
    ipcRenderer.invoke('open-file'),

  /**
   * Subscribe to native menu actions. The callback receives a name string:
   * 'new', 'open', 'save', 'save-as', 'import-csv', 'export-csv', 'export-xlsx',
   * 'undo', 'redo', 'cut', 'copy', 'paste'.
   */
  onMenuAction: (cb) =>
    ipcRenderer.on('menu-action', (_event, name) => cb(name)),
});
