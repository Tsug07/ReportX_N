const { app, BrowserWindow } = require('electron');
const path = require('path');
const { fork } = require('child_process');

let mainWindow;
let serverProcess;

// ✅ Detecta se está rodando empacotado (instalador) ou em dev
const isPackaged = app.isPackaged;

// ✅ Corrige caminhos de recursos
// process.resourcesPath aponta para /resources/ dentro do instalador
const basePath = isPackaged ? process.resourcesPath : __dirname;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 900,
    height: 650,
    frame: false,
    icon: path.join(basePath, 'public', 'icon.ico'), // garante que o ícone seja encontrado
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  if (isPackaged) {
    // ✅ No instalador: abre o index.html do build (sem localhost)
    mainWindow.loadFile(path.join(basePath, 'index.html'));
  } else {
    // ✅ Em desenvolvimento: carrega o servidor local
    mainWindow.loadURL('http://localhost:3000');
  }
}

app.whenReady().then(() => {
  console.log('🚀 Iniciando servidor interno...');
  
  // ✅ Corrige caminho do servidor.js para o modo empacotado
  const serverPath = path.join(basePath, 'server.js');
  serverProcess = fork(serverPath);

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
    serverProcess.kill();
    console.log('🛑 Servidor encerrado.');
  }
});
