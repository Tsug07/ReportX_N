const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

const app = express();
const port = 3000;

console.log('📦 Diretório de trabalho:', __dirname);

// ===============================
// 🔧 Configuração do multer (upload)
// ===============================
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, __dirname);
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
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ✅ ADICIONE ESTA ROTA
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// ===============================
// ⚙️ Leitura de configuração atual do RelaEcac.js
// ===============================
function lerConfiguracaoAtual() {
  const relaPath = path.join(__dirname, 'RelaEcac.js');
  try {
    const codigo = fs.readFileSync(relaPath, 'utf8');
    const certPath = codigo.match(/const CERT_PATH = .*?"([^"]+)"/)?.[1] || '';
    const certPassword = codigo.match(/const CERT_PASSWORD = "([^"]*)"/)?.[1] || '';
    const consumerKey = codigo.match(/const CONSUMER_KEY = process\.env\.CONSUMER_KEY/)?.[0] || '';
    const consumerSecret = codigo.match(/const CONSUMER_SECRET = process\.env\.CONSUMER_SECRET/)?.[0] || '';
    const pdfDirMatch = codigo.match(/const pdfDir = path\.join\(__dirname,\s*"([^"]+)"\)/);
    const pdfDir = pdfDirMatch ? pdfDirMatch[1] : 'pdfs';
    return { certPath: path.basename(certPath), certPassword, pdfDir };
  } catch {
    return { certPath: '', certPassword: '', pdfDir: 'pdfs' };
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
    const logsDir = path.join(__dirname, 'logs');

    try {
      console.log('📝 Dados recebidos:', { senha: senha ? '***' : 'vazio', pdfDir, selectedFolderPath });

      const relaPath = path.join(__dirname, 'RelaEcac.js');
      let codigo = fs.readFileSync(relaPath, 'utf8');

      // ✅ Atualizar certificado (se enviado)
      if (req.files && req.files.certificado && req.files.certificado[0]) {
        const certFileName = req.files.certificado[0].filename;
        const certPathNew = path.join(__dirname, certFileName);
        
        codigo = codigo.replace(
          /const CERT_PATH = path\.join\(__dirname,\s*"[^"]+"\);/,
          `const CERT_PATH = path.join(__dirname, "${certFileName}");`
        );
        
        console.log('✅ Certificado atualizado:', certFileName);
      }

      // ✅ Atualizar senha do certificado
      if (senha) {
        codigo = codigo.replace(
          /const CERT_PASSWORD = "[^"]*";/,
          `const CERT_PASSWORD = "${senha}";`
        );
        console.log('✅ Senha do certificado atualizada');
      }

      // ✅ Atualizar pasta de PDFs
      const pastaPdfs = selectedFolderPath || pdfDir || 'pdfs';
      codigo = codigo.replace(
        /const pdfDir = path\.join\(__dirname,\s*"[^"]+"\);/,
        `const pdfDir = path.join(__dirname, "${pastaPdfs}");`
      );
      console.log('✅ Pasta de PDFs configurada:', pastaPdfs);

      // ✅ Salvar arquivo modificado
      fs.writeFileSync(relaPath, codigo);
      console.log('✅ RelaEcac.js atualizado com sucesso');

      // ✅ Garantir que pasta de logs existe e está limpa
      if (!fs.existsSync(logsDir)) {
        fs.mkdirSync(logsDir, { recursive: true });
      }
      
      // Limpar logs anteriores
      fs.writeFileSync(path.join(logsDir, 'success.log'), '');
      fs.writeFileSync(path.join(logsDir, 'errors.log'), '');
      console.log('✅ Logs limpos e prontos');

      // ✅ Executar RelaEcac.js
      console.log('▶️ Executando RelaEcac.js...');
      const processo = spawn('node', [path.join(__dirname, 'RelaEcac.js')], {
        cwd: __dirname,
        env: process.env
      });

      processo.on('error', (err) => {
        try {
          const errorMessage = `❌ Falha ao iniciar o processo: ${err.message}\n`;
          console.error(errorMessage);
          fs.appendFileSync(path.join(logsDir, 'errors.log'), errorMessage);
        } catch (e) {
          console.error('Falha ao escrever no log de erro:', e);
        }
      });

      processo.stdout.on('data', (data) => {
        try {
          const logMessage = data.toString();
          console.log('STDOUT:', logMessage);
          fs.appendFileSync(path.join(logsDir, 'success.log'), logMessage);
        } catch (e) {
          console.error('Falha ao escrever no log de sucesso:', e);
        }
      });

      processo.stderr.on('data', (data) => {
        try {
          const errorMessage = data.toString();
          console.error('STDERR:', errorMessage);
          fs.appendFileSync(path.join(logsDir, 'errors.log'), errorMessage);
        } catch (e) {
          console.error('Falha ao escrever no log de erro (stderr):', e);
        }
      });

      processo.on('close', (code) => {
        try {
          const finalMessage = `\n✅ Processamento concluído com código ${code}.\n`;
          console.log(finalMessage);
          fs.appendFileSync(path.join(logsDir, 'success.log'), finalMessage);
        } catch (e) {
          console.error('Falha ao escrever log de finalização:', e);
        }
      });

      res.json({ success: true, message: 'Processamento iniciado! Acompanhe os logs.' });
    } catch (error) {
      const errorMsg = 'Erro interno: ' + error.message;
      console.error('❌ Erro no /processar:', error);
      
      // Gravar erro no log
      try {
        fs.appendFileSync(path.join(logsDir, 'errors.log'), `[ERRO SERVIDOR] ${errorMsg}\n`);
      } catch (e) {
        console.error('Não foi possível gravar log de erro:', e);
      }
      
      res.status(500).json({ success: false, message: errorMsg });
    }
  }
);

// ===============================
// 📊 Status dos logs (COM CORS para desenvolvimento)
// ===============================
app.get('/status', (req, res) => {
  // Adicionar headers CORS
  res.header('Access-Control-Allow-Origin', '*');
  
  const logsDir = path.join(__dirname, 'logs');
  try {
    const successLogPath = path.join(logsDir, 'success.log');
    const errorLogPath = path.join(logsDir, 'errors.log');
    
    const successLog = fs.existsSync(successLogPath)
      ? fs.readFileSync(successLogPath, 'utf8')
      : 'Aguardando início...';
      
    const errorLog = fs.existsSync(errorLogPath)
      ? fs.readFileSync(errorLogPath, 'utf8')
      : 'Nenhum erro registrado.';
    
    console.log('📊 Status requisitado - Success:', successLog.length, 'chars, Errors:', errorLog.length, 'chars');
    
    res.json({ 
      success: successLog || 'Aguardando início...', 
      errors: errorLog || 'Nenhum erro registrado.'
    });
  } catch(e) {
    console.error('❌ Erro ao ler logs:', e);
    res.status(500).json({ 
      error: 'Erro ao ler logs: ' + e.message,
      success: 'Erro ao carregar logs',
      errors: 'Erro ao carregar logs'
    });
  }
});

// ===============================
// 🚀 Inicialização
// ===============================
app.listen(port, () => {
  console.log(`🌐 Servidor rodando em http://localhost:${port}`);
  console.log(`📁 Diretório de trabalho: ${__dirname}`);
  console.log(`📋 Acesse http://localhost:${port} para começar`);
});