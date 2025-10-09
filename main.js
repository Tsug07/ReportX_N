const { app, BrowserWindow } = require('electron');
const path = require('path');
const { fork } = require('child_process');

let mainWindow;
let serverProcess;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 900,
    height: 650,
    frame: false, // barra personalizada
    icon: path.join(__dirname, 'public', 'icon.ico'), // 👉 use o .ico aqui
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  mainWindow.loadURL("http://localhost:3000");
}


app.whenReady().then(() => {
  console.log("🚀 Iniciando servidor interno...");
  // Inicia o servidor Node.js
  serverProcess = fork(path.join(__dirname, 'server.js'));

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('quit', () => {
  if (serverProcess) {
    serverProcess.kill(); // encerra o servidor junto com o app
    console.log("🛑 Servidor encerrado.");
  }
});