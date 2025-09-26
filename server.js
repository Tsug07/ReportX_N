const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');

const app = express();
const port = 3000;

// Configuração do multer para upload de arquivos
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, __dirname);
  },
  filename: (req, file, cb) => {
    if (file.fieldname === 'certificado') {
      // Mantém o nome original do certificado
      cb(null, file.originalname);
    } else if (file.fieldname === 'excel') {
      cb(null, 'empresas_filtradas.xlsx');
    }
  }
});

const upload = multer({ storage: storage });

// Middleware
app.use(express.static('public'));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Rota principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Função para ler configurações atuais do RelaEcac.js
function lerConfiguracaoAtual() {
  try {
    const codigo = fs.readFileSync('RelaEcac.js', 'utf8');
    
    const certPath = codigo.match(/const CERT_PATH = .*?"([^"]+)"/)?.[1] || '';
    const certPassword = codigo.match(/const CERT_PASSWORD = "([^"]*)"/)?.[1] || '';
    const consumerKey = codigo.match(/const CONSUMER_KEY = "([^"]*)"/)?.[1] || '';
    const consumerSecret = codigo.match(/const CONSUMER_SECRET = "([^"]*)"/)?.[1] || '';
    
    // Ler pasta de PDFs atual
    const pdfDirMatch = codigo.match(/const pdfDir = path\.join\(__dirname,\s*"([^"]+)"\)/);
    const pdfDir = pdfDirMatch ? pdfDirMatch[1] : 'pdfs';
    
    return {
      certPath: path.basename(certPath),
      certPassword,
      consumerKey,
      consumerSecret,
      pdfDir
    };
  } catch (error) {
    return {
      certPath: '',
      certPassword: '',
      consumerKey: '',
      consumerSecret: '',
      pdfDir: 'pdfs'
    };
  }
}

// Rota para obter configurações atuais
app.get('/config', (req, res) => {
  const config = lerConfiguracaoAtual();
  res.json(config);
});

// Rota para upload e processamento
app.post('/processar', upload.fields([
  { name: 'certificado', maxCount: 1 },
  { name: 'excel', maxCount: 1 }
]), (req, res) => {
  const { senha, pdfDir, selectedFolderPath } = req.body;
  
  if (!req.files.excel) {
    return res.status(400).json({ 
      success: false, 
      message: 'Por favor, selecione o arquivo Excel.' 
    });
  }

  const excelPath = req.files.excel[0].path;
  let certificadoPath = null;
  
  // Se um novo certificado foi enviado, usar ele
  if (req.files.certificado && req.files.certificado[0]) {
    certificadoPath = req.files.certificado[0].path;
  }

  // Determinar qual pasta usar
  let pastaPdf = selectedFolderPath || pdfDir;

  // Atualizar o arquivo RelaEcac.js com os novos valores
  try {
    let codigo = fs.readFileSync('RelaEcac.js', 'utf8');
    
    // Só atualizar certificado se um novo foi enviado
    if (certificadoPath) {
      codigo = codigo.replace(
        /const CERT_PATH = .*?;/,
        `const CERT_PATH = path.join(__dirname, "${path.basename(certificadoPath)}");`
      );
    }
    
    // Só atualizar senha se foi fornecida
    if (senha) {
      codigo = codigo.replace(
        /const CERT_PASSWORD = .*?;/,
        `const CERT_PASSWORD = "${senha}";`
      );
    }

    // Atualizar pasta de PDFs se foi fornecida
    if (pastaPdf && pastaPdf.trim()) {
      // Para nomes de pasta simples (como selecionados pelo picker), criar na pasta do projeto
      // Para caminhos completos, usar como está
      let novoCaminhoCode;
      if (path.isAbsolute(pastaPdf)) {
        novoCaminhoCode = `const pdfDir = "${pastaPdf.replace(/\\/g, '\\\\')}";`;
      } else {
        novoCaminhoCode = `const pdfDir = path.join(__dirname, "${pastaPdf}");`;
      }
      
      codigo = codigo.replace(
        /const pdfDir = .*?;/,
        novoCaminhoCode
      );
      
      console.log(`📁 Pasta de PDFs atualizada para: ${pastaPdf}`);
    }

    fs.writeFileSync('RelaEcac.js', codigo);

    // Executar o script
    const processo = exec('node RelaEcac.js', (error, stdout, stderr) => {
      if (error) {
        console.error('Erro na execução:', error);
        return;
      }
      console.log('Saída:', stdout);
      if (stderr) console.error('Erro:', stderr);
    });

    // Enviar resposta imediata
    res.json({ 
      success: true, 
      message: 'Processamento iniciado! Acompanhe o progresso no console.' 
    });

  } catch (error) {
    console.error('Erro ao processar:', error);
    res.status(500).json({ 
      success: false, 
      message: 'Erro interno do servidor: ' + error.message 
    });
  }
});

// Rota para verificar status (logs)
app.get('/status', (req, res) => {
  try {
    const successLog = fs.existsSync('logs/success.log') 
      ? fs.readFileSync('logs/success.log', 'utf8').split('\n').slice(-10).join('\n')
      : '';
    
    const errorLog = fs.existsSync('logs/errors.log') 
      ? fs.readFileSync('logs/errors.log', 'utf8').split('\n').slice(-10).join('\n')
      : '';

    res.json({
      success: successLog,
      errors: errorLog
    });
  } catch (error) {
    res.status(500).json({ error: 'Erro ao ler logs' });
  }
});

// Iniciar servidor
app.listen(port, () => {
  console.log(`🌐 Servidor rodando em http://localhost:${port}`);
  console.log('📁 Certifique-se de que o arquivo RelaEcac.js está na mesma pasta');
});