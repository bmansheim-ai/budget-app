const { app, BrowserWindow } = require('electron');
const path = require('path');

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 1300,
    height: 850,
    title: "Mansheim Budget Processor",
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  // Load the production build index.html file
  mainWindow.loadFile(path.join(__dirname, 'dist', 'index.html'));
  
  // Optional: Uncomment the line below to open DevTools by default
  // mainWindow.webContents.openDevTools();
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.on('window-all-closed', () => {
      app.quit();
    });
  }
});