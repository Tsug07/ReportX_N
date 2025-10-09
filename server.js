const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');

const app = express();
const port = 3000;

// ✅ Detecta se o app está empacotado (instalador)
const isPackaged = process.mainModule?.filename.indexOf('app.asar') !== -1;

// ✅ Corrige caminho base: usa process.resourcesPath dentro do instalador
const basePath = isPackaged ? process.resourcesPath : __dirname;
console.log('📦 BasePath:', basePath);

// ===============================
// 🔧 Configuração do multer (upload)
// ===============================
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, basePath); // ✅ salva os arquivos no caminho correto
  },
  filename: (req, file, cb) => {
    if (file.fieldname === 'certificado') {
      cb(null, file.originalname);
    } else if (file.fieldname === 'excel') {
      cb(null, 'empresas_filtradas.xlsx');
    }
  }
});

const upload = multer({ storage });

// Middleware
app.use(express.static(path.join(basePath, 'public')));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ===============================
// 🌐 Rota principal
// ===============================
app.get('/', (req, res) => {
  res.sendFile(path.join(basePath, 'index.html'));
});

// ===============================
// ⚙️ Leitura de configuração atual do RelaEcac.js
// ===============================
function lerConfiguracaoAtual() {
  const relaPath = path.join(basePath, 'RelaEcac.js'); // ✅ caminho corrigido
  try {
    const codigo = fs.readFileSync(relaPath, 'utf8');
    const certPath = codigo.match(/const CERT_PATH = .*?"([^"]+)"/)?.[1] || '';
    const certPassword = codigo.match(/const CERT_PASSWORD = "([^"]*)"/)?.[1] || '';
    const consumerKey = codigo.match(/const CONSUMER_KEY = "([^"]*)"/)?.[1] || '';
    const consumerSecret = codigo.match(/const CONSUMER_SECRET = "([^"]*)"/)?.[1] || '';
    const pdfDirMatch = codigo.match(/const pdfDir = path\.join\(__dirname,\s*"([^"]+)"\)/);
    const pdfDir = pdfDirMatch ? pdfDirMatch[1] : 'pdfs';
    return { certPath: path.basename(certPath), certPassword, consumerKey, consumerSecret, pdfDir };
  } catch {
    return { certPath: '', certPassword: '', consumerKey: '', consumerSecret: '', pdfDir: 'pdfs' };
  }
}

app.get('/config', (req, res) => {
  res.json(lerConfiguracaoAtual());
});

// ===============================
// 🚀 Upload e processamento
// ===============================
app.post(
  '/processar',
  upload.fields([
    { name: 'certificado', maxCount: 1 },
    { name: 'excel', maxCount: 1 },
  ]),
  (req, res) => {
    const { senha, pdfDir, selectedFolderPath } = req.body;

    if (!req.files.excel) {
      return res.status(400).json({ success: false, message: 'Por favor, selecione o arquivo Excel.' });
    }

    const excelPath = req.files.excel[0].path;
    let certificadoPath = req.files.certificado?.[0]?.path || null;
    let pastaPdf = selectedFolderPath || pdfDir;

    try {
      const relaPath = path.join(basePath, 'RelaEcac.js'); // ✅ caminho correto
      let codigo = fs.readFileSync(relaPath, 'utf8');

      if (certificadoPath) {
        codigo = codigo.replace(
          /const CERT_PATH = .*?;/,
          `const CERT_PATH = path.join(__dirname, "${path.basename(certificadoPath)}");`
        );
      }
      if (senha) {
        codigo = codigo.replace(/const CERT_PASSWORD = .*?;/, `const CERT_PASSWORD = "${senha}";`);
      }

      if (pastaPdf && pastaPdf.trim()) {
        const novoCaminhoCode = path.isAbsolute(pastaPdf)
          ? `const pdfDir = "${pastaPdf.replace(/\\/g, '\\\\')}";`
          : `const pdfDir = path.join(__dirname, "${pastaPdf}");`;
        codigo = codigo.replace(/const pdfDir = .*?;/, novoCaminhoCode);
      }

      fs.writeFileSync(relaPath, codigo);

      // ✅ Garante que a pasta de logs existe no basePath
      const logsDir = path.join(basePath, 'logs');
      if (!fs.existsSync(logsDir)) fs.mkdirSync(logsDir, { recursive: true });
      fs.writeFileSync(path.join(logsDir, 'success.log'), '');
      fs.writeFileSync(path.join(logsDir, 'errors.log'), '');

      // ✅ Executa RelaEcac.js usando caminho absoluto
      const cmd = `node "${path.join(basePath, 'RelaEcac.js')}"`;
      console.log('▶️ Executando:', cmd);
      const processo = exec(cmd, (error, stdout, stderr) => {
        if (error) {
          console.error('Erro na execução:', error);
          fs.appendFileSync(path.join(logsDir, 'errors.log'), `\n❌ Erro: ${error.message}\n`);
          return;
        }
        if (stderr) {
          console.error('Stderr:', stderr);
          fs.appendFileSync(path.join(logsDir, 'errors.log'), `\n⚠️ ${stderr}\n`);
        }
        fs.appendFileSync(path.join(logsDir, 'success.log'), '\n✅ Processamento concluído!\n');
        console.log(stdout);
      });

      res.json({ success: true, message: 'Processamento iniciado! Confira o console.' });
    } catch (error) {
      console.error('Erro ao processar:', error);
      res.status(500).json({ success: false, message: 'Erro interno: ' + error.message });
    }
  }
);

// ===============================
// 📊 Status dos logs
// ===============================
app.get('/status', (req, res) => {
  const logsDir = path.join(basePath, 'logs');
  try {
    const successLog = fs.existsSync(path.join(logsDir, 'success.log'))
      ? fs.readFileSync(path.join(logsDir, 'success.log'), 'utf8').split('\n').slice(-10).join('\n')
      : '';
    const errorLog = fs.existsSync(path.join(logsDir, 'errors.log'))
      ? fs.readFileSync(path.join(logsDir, 'errors.log'), 'utf8').split('\n').slice(-10).join('\n')
      : '';
    res.json({ success: successLog, errors: errorLog });
  } catch {
    res.status(500).json({ error: 'Erro ao ler logs' });
  }
});

// ===============================
// 🚀 Inicialização
// ===============================
app.listen(port, () => {
  console.log(`🌐 Servidor rodando em http://localhost:${port}`);
  console.log('📁 RelaEcac.js e arquivos de trabalho no basePath:', basePath);
});
