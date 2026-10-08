/* SheetDesk — Electron main process. */
const { app, BrowserWindow, Menu, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');

let mainWindow = null;

const FILE_FILTER = [
  { name: 'SheetDesk workbook', extensions: ['sheetdesk.json'] },
  { name: 'All files', extensions: ['*'] },
];

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    title: 'SheetDesk',
  });

  mainWindow.loadFile(path.join(__dirname, 'app', 'index.html'));

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

function sendMenuAction(name) {
  const win = BrowserWindow.getFocusedWindow() || mainWindow;
  if (win) win.webContents.send('menu-action', name);
}

const SHORTCUTS_TEXT = [
  'File',
  '  New workbook ............ Ctrl/Cmd+N',
  '  Open workbook ........... Ctrl/Cmd+O',
  '  Save .................... Ctrl/Cmd+S',
  'Editing',
  '  Undo .................... Ctrl/Cmd+Z',
  '  Redo .................... Ctrl/Cmd+Shift+Z',
  '  Cut / Copy / Paste ...... Ctrl/Cmd+X · C · V',
  '  Bold / Italic ........... Ctrl/Cmd+B · Ctrl/Cmd+I',
  '  Clear cells ............. Delete or Backspace',
  'Navigation',
  '  Move .................... Arrow keys',
  '  Extend selection ........ Shift+Arrows, Shift+Click, drag',
  '  Enter: move down ........ Enter (Shift+Enter: up)',
  '  Tab: move right ......... Tab (Shift+Tab: left)',
  '  Edit in cell ............ F2 or double-click',
  '  Commit / cancel edit .... Enter / Esc',
].join('\n');

function buildMenu() {
  const isMac = process.platform === 'darwin';
  const template = [
    {
      label: 'File',
      submenu: [
        { label: 'New', accelerator: 'CmdOrCtrl+N', click: () => sendMenuAction('new') },
        { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: () => sendMenuAction('open') },
        { type: 'separator' },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => sendMenuAction('save') },
        { label: 'Save As…', accelerator: 'CmdOrCtrl+Shift+S', click: () => sendMenuAction('save-as') },
        { type: 'separator' },
        { label: 'Import CSV…', click: () => sendMenuAction('import-csv') },
        { label: 'Export CSV…', click: () => sendMenuAction('export-csv') },
        { label: 'Export XLSX…', click: () => sendMenuAction('export-xlsx') },
        { type: 'separator' },
        isMac ? { role: 'close' } : { label: 'Quit', accelerator: 'CmdOrCtrl+Q', click: () => app.quit() },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        { label: 'Undo', accelerator: 'CmdOrCtrl+Z', click: () => sendMenuAction('undo') },
        { label: 'Redo', accelerator: 'CmdOrCtrl+Shift+Z', click: () => sendMenuAction('redo') },
        { type: 'separator' },
        { label: 'Cut', accelerator: 'CmdOrCtrl+X', click: () => sendMenuAction('cut') },
        { label: 'Copy', accelerator: 'CmdOrCtrl+C', click: () => sendMenuAction('copy') },
        { label: 'Paste', accelerator: 'CmdOrCtrl+V', click: () => sendMenuAction('paste') },
      ],
    },
    {
      label: 'View',
      submenu: [
        { role: 'reload' },
        { type: 'separator' },
        { role: 'zoomIn', label: 'Zoom In' },
        { role: 'zoomOut', label: 'Zoom Out' },
        { role: 'resetZoom', label: 'Reset Zoom' },
      ],
    },
    {
      label: 'Help',
      submenu: [
        {
          label: 'Keyboard Shortcuts',
          click: () => {
            const win = BrowserWindow.getFocusedWindow() || mainWindow;
            dialog.showMessageBox(win, {
              type: 'info',
              title: 'Keyboard Shortcuts',
              message: 'SheetDesk keyboard shortcuts',
              detail: SHORTCUTS_TEXT,
              buttons: ['Close'],
            });
          },
        },
        {
          label: 'About SheetDesk',
          click: () => {
            const win = BrowserWindow.getFocusedWindow() || mainWindow;
            let version = '1.0.0';
            try {
              version = require('./package.json').version || version;
            } catch (_) { /* keep default */ }
            dialog.showMessageBox(win, {
              type: 'info',
              title: 'About SheetDesk',
              message: `SheetDesk ${version}`,
              detail: 'A clean, offline-first desktop spreadsheet with an Excel-like formula engine.\nDynamic arrays, LET/LAMBDA, and 90+ functions. All data stays on your machine.',
              buttons: ['Close'],
            });
          },
        },
      ],
    },
  ];

  // macOS convention: app menu first.
  if (isMac) {
    template.unshift({
      label: app.name,
      submenu: [
        { role: 'about' },
        { type: 'separator' },
        { role: 'hide' },
        { role: 'hideOthers' },
        { role: 'unhide' },
        { type: 'separator' },
        { role: 'quit' },
      ],
    });
  }

  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

/* ------------------------------ IPC handlers ------------------------------ */

/** save-file: Save dialog + fs write. Used for .sheetdesk.json (and CSV export). */
ipcMain.handle('save-file', async (event, defaultName, data, opts) => {
  const win = BrowserWindow.fromWebContents(event.sender) || mainWindow;
  const o = opts || {};
  const { canceled, filePath } = await dialog.showSaveDialog(win, {
    title: o.title || 'Save workbook',
    defaultPath: defaultName || 'Untitled.sheetdesk.json',
    filters: o.filters || FILE_FILTER,
  });
  if (canceled || !filePath) return null;
  await fs.promises.writeFile(filePath, data, 'utf8');
  return { path: filePath };
});

/** save-file-to: silent write to a known path. */
ipcMain.handle('save-file-to', async (_event, filePath, data) => {
  if (!filePath || typeof filePath !== 'string') {
    throw new Error('save-file-to requires a file path');
  }
  await fs.promises.writeFile(filePath, data, 'utf8');
  return { path: filePath };
});

/** open-file: Open dialog + fs read for .sheetdesk.json files. */
ipcMain.handle('open-file', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender) || mainWindow;
  const { canceled, filePaths } = await dialog.showOpenDialog(win, {
    title: 'Open workbook',
    properties: ['openFile'],
    filters: FILE_FILTER,
  });
  if (canceled || !filePaths || filePaths.length === 0) return null;
  const filePath = filePaths[0];
  const data = await fs.promises.readFile(filePath, 'utf8');
  return { path: filePath, data };
});

/* --------------------------------- app ------------------------------------ */

app.whenReady().then(() => {
  buildMenu();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
