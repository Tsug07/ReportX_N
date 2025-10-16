const { app, BrowserWindow, Notification } = require('electron');
const path = require('path');
const express = require('express');
const multer = require('multer');
const fs = require('fs');

let mainWindow;
let server;

// ✅ DETERMINAR CAMINHOS CORRETOS
const isDev = !app.isPackaged;
const basePath = isDev ? __dirname : path.join(process.resourcesPath, 'app.asar');

// ✅ FUNÇÃO PARA CONFIGURAR DIRETÓRIO DE TRABALHO
function setupWorkDirectory() {
  const workDir = app.getPath('userData');
  const logsDir = path.join(workDir, 'logs');
  const pdfsDir = path.join(workDir, 'pdfs');
  const jsonDir = path.join(workDir, 'json_responses');
  
  // Criar diretórios necessários
  [logsDir, pdfsDir, jsonDir].forEach(dir => {
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
  });

  // ✅ COPIAR .env
  const envSource = path.join(basePath, '.env');
  const envDest = path.join(workDir, '.env');
  
  if (fs.existsSync(envSource) && !fs.existsSync(envDest)) {
    fs.copyFileSync(envSource, envDest);
    console.log('✅ .env copiado para:', envDest);
  }

  // ✅ COPIAR RelaEcac.js (mantém para backup, mas não usaremos como spawn)
  const relaSource = path.join(basePath, 'RelaEcac.js');
  const relaDest = path.join(workDir, 'RelaEcac.js');
  
  if (fs.existsSync(relaSource)) {
    fs.copyFileSync(relaSource, relaDest);
    console.log('✅ RelaEcac.js copiado para:', relaDest);
  }

  console.log('✅ Diretórios configurados em:', workDir);
  return workDir;
}

// ✅ INICIAR SERVIDOR EXPRESS
function startServer() {
  return new Promise((resolve) => {
    console.log('🚀 Iniciando servidor Express...');
    
    const expressApp = express();
    const port = 3000;

    // ✅ DEFINIR workDir UMA VEZ NO ESCOPO DA FUNÇÃO
    const workDir = isDev ? __dirname : app.getPath('userData');

    // Configuração do multer
    const storage = multer.diskStorage({
      destination: (req, file, cb) => {
        cb(null, workDir);
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

    // Middlewares
    expressApp.use(express.static(path.join(basePath, 'public')));
    expressApp.use(express.json());
    expressApp.use(express.urlencoded({ extended: true }));

    // Rota principal
    expressApp.get('/', (req, res) => {
      res.sendFile(path.join(basePath, 'index.html'));
    });

    // Rota de configuração
    expressApp.get('/config', (req, res) => {
      res.json({ certPath: '', certPassword: '', pdfDir: 'pdfs' });
    });

    // ✅ ROTA DE PROCESSAMENTO - EXECUTA DIRETAMENTE SEM SPAWN
    expressApp.post('/processar', upload.fields([
      { name: 'certificado', maxCount: 1 },
      { name: 'excel', maxCount: 1 }
    ]), async (req, res) => {
      const { senha, pdfDir } = req.body;
      
      console.log('🔍 === DEBUG PATHS ===');
      console.log('📁 workDir:', workDir);
      console.log('📁 pdfDir (form):', pdfDir);
      console.log('📁 req.body completo:', req.body);
      console.log('=====================');

      // ✅ Responde imediatamente ao cliente
      res.json({ success: true, message: 'Processamento iniciado!' });

      // ✅ Executa o processamento em background
      try {
        // Passa o pdfDir para a função
        await executarProcessamento(workDir, pdfDir);
        
        // ✅ NOTIFICAÇÃO NATIVA QUANDO CONCLUIR
        if (Notification.isSupported()) {
          const notification = new Notification({
            title: '🎉 Processamento Concluído!',
            body: 'Todos os relatórios foram gerados com sucesso. Clique para visualizar.',
            icon: path.join(basePath, 'public', 'icon.ico'),
            urgency: 'normal'
          });
          
          notification.show();
          
          // Trazer janela para frente ao clicar na notificação
          notification.on('click', () => {
            if (mainWindow) {
              if (mainWindow.isMinimized()) mainWindow.restore();
              mainWindow.focus();
            }
          });
        }
        
        console.log('✅ Notificação de conclusão enviada');
        
      } catch (error) {
        console.error('❌ Erro no processamento:', error);
        const errorLogPath = path.join(workDir, 'logs', 'errors.log');
        const timestamp = new Date().toLocaleString('pt-BR');
        fs.appendFileSync(errorLogPath, `[${timestamp}] Erro geral: ${error.message}\n${error.stack}\n`);
        
        // ✅ NOTIFICAÇÃO DE ERRO
        if (Notification.isSupported()) {
          new Notification({
            title: '❌ Erro no Processamento',
            body: 'Ocorreu um erro. Verifique os logs para mais detalhes.',
            icon: path.join(basePath, 'public', 'icon.ico'),
            urgency: 'critical'
          }).show();
        }
      }
    });

    // Rota de status
    expressApp.get('/status', (req, res) => {
      const logsDir = path.join(workDir, 'logs');
      try {
        const successLog = fs.existsSync(path.join(logsDir, 'success.log'))
          ? fs.readFileSync(path.join(logsDir, 'success.log'), 'utf8')
          : 'Aguardando início...';
        const errorLog = fs.existsSync(path.join(logsDir, 'errors.log'))
          ? fs.readFileSync(path.join(logsDir, 'errors.log'), 'utf8')
          : 'Nenhum erro.';
        res.json({ success: successLog, errors: errorLog });
      } catch(e) {
        res.status(500).json({ error: 'Erro ao ler logs: ' + e.message });
      }
    });

    server = expressApp.listen(port, () => {
      console.log(`✅ Servidor rodando em http://localhost:${port}`);
      resolve();
    });
  });
}

// ✅ FUNÇÃO QUE EXECUTA O PROCESSAMENTO (adaptado do RelaEcac.js)
async function executarProcessamento(workDir, pdfDirCustom) {
  const axios = require('axios');
  const https = require('https');
  const pdfParse = require('pdf-parse');
  const XLSX = require('xlsx');
  const dotenv = require('dotenv');

  // Carregar variáveis de ambiente
  const envPath = path.join(workDir, '.env');
  if (fs.existsSync(envPath)) {
    dotenv.config({ path: envPath });
    console.log('✅ .env carregado');
  } else {
    console.warn('⚠️ .env não encontrado');
  }

  // Configurações
  const CERT_PATH = path.join(workDir, "102 - CANELLA & SANTOS CONTABILIDADE LTDA - SENHA 123456 - V 28.03.2026.pfx");
  const CERT_PASSWORD = "123456";
  const AUTH_URL = process.env.AUTH_URL || "https://autenticacao.sapi.serpro.gov.br/authenticate";
  const CONSUMER_KEY = process.env.CONSUMER_KEY;
  const CONSUMER_SECRET = process.env.CONSUMER_SECRET;
  const API_URL_APOIAR = process.env.API_URL_APOIAR || "https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Apoiar";
  const API_URL_EMITIR = process.env.API_URL_EMITIR || "https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Emitir";
  const CONTRATANTE_CNPJ = process.env.CONTRATANTE_CNPJ || "06310711000149";

  if (!CONSUMER_KEY || !CONSUMER_SECRET) {
    throw new Error("CONSUMER_KEY e CONSUMER_SECRET devem estar definidas no .env");
  }

  console.log("📋 Configurações carregadas");
  console.log(`📜 Certificado: ${path.basename(CERT_PATH)}`);
  console.log(`🏢 Contratante: ${CONTRATANTE_CNPJ}`);

  // ✅ Determinar pasta de PDFs (customizada ou padrão)
  let pdfDir;
  if (pdfDirCustom && pdfDirCustom.trim() !== '') {
    const caminhoLimpo = pdfDirCustom.trim();
    
    // Se for caminho absoluto (tem : ou começa com \ ou /)
    if (caminhoLimpo.includes(':') || caminhoLimpo.startsWith('\\') || caminhoLimpo.startsWith('/')) {
      pdfDir = caminhoLimpo;
      console.log('📁 Usando caminho ABSOLUTO:', pdfDir);
    } else {
      // Caminho relativo - cria dentro do workDir
      pdfDir = path.join(workDir, caminhoLimpo);
      console.log('📁 Usando caminho RELATIVO (dentro de AppData):', pdfDir);
    }
  } else {
    pdfDir = path.join(workDir, 'pdfs');
    console.log('📁 Usando pasta PADRÃO:', pdfDir);
  }

  // Criar pasta de PDFs se não existir
  try {
    if (!fs.existsSync(pdfDir)) {
      fs.mkdirSync(pdfDir, { recursive: true });
      console.log('✅ Pasta de PDFs criada:', pdfDir);
    } else {
      console.log('✅ Pasta de PDFs já existe:', pdfDir);
    }
  } catch (error) {
    console.error('❌ Erro ao criar pasta:', error.message);
    throw new Error(`Não foi possível criar a pasta: ${pdfDir}. Verifique as permissões.`);
  }

  // Diretórios de saída
  const jsonDir = path.join(workDir, 'json_responses');
  const logsDir = path.join(workDir, 'logs');
  const successLogPath = path.join(logsDir, 'success.log');
  const errorLogPath = path.join(logsDir, 'errors.log');

  // ✅ LIMPAR LOGS ANTIGOS NO INÍCIO
  if (fs.existsSync(successLogPath)) {
    fs.writeFileSync(successLogPath, ''); // Limpa o arquivo
  }
  if (fs.existsSync(errorLogPath)) {
    fs.writeFileSync(errorLogPath, ''); // Limpa o arquivo
  }
  console.log('🧹 Logs anteriores limpos');

  function writeLog(filePath, message) {
    const timestamp = new Date().toLocaleString('pt-BR');
    fs.appendFileSync(filePath, `[${timestamp}] ${message}\n`);
  }

  // Ler CNPJs do Excel
  function lerCNPJsDoExcel() {
    const excelPath = path.join(workDir, 'empresas_filtradas.xlsx');
    
    if (!fs.existsSync(excelPath)) {
      console.error('❌ Arquivo empresas_filtradas.xlsx não encontrado!');
      return [];
    }

    const workbook = XLSX.readFile(excelPath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const cnpjs = [];

    for (let row = 1; row <= range.e.r; row++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: 2 })];
      if (cell && cell.v) {
        cnpjs.push(String(cell.v).replace(/\D/g, ''));
      }
    }

    console.log('📂 CNPJs carregados:', cnpjs.length);
    return cnpjs;
  }

  // Autenticação
  async function getAuthTokens() {
    if (!fs.existsSync(CERT_PATH)) {
      throw new Error(`Certificado não encontrado: ${CERT_PATH}`);
    }

    const httpsAgent = new https.Agent({
      pfx: fs.readFileSync(CERT_PATH),
      passphrase: CERT_PASSWORD,
    });

    const authHeader = Buffer.from(`${CONSUMER_KEY}:${CONSUMER_SECRET}`).toString('base64');

    const response = await axios.post(
      AUTH_URL,
      new URLSearchParams({ grant_type: 'client_credentials' }),
      {
        headers: {
          Authorization: `Basic ${authHeader}`,
          'role-type': 'TERCEIROS',
          'content-type': 'application/x-www-form-urlencoded',
        },
        httpsAgent,
      }
    );

    console.log('✅ Tokens obtidos');
    return response.data;
  }

  // Solicitar protocolos
  async function solicitarProtocolos(tokens, contribuintes) {
    const protocolos = [];

    for (const cnpj of contribuintes) {
      const requestBody = {
        contratante: { numero: CONTRATANTE_CNPJ, tipo: 2 },
        autorPedidoDados: { numero: CONTRATANTE_CNPJ, tipo: 2 },
        contribuinte: { numero: cnpj, tipo: 2 },
        pedidoDados: {
          idSistema: 'SITFIS',
          idServico: 'SOLICITARPROTOCOLO91',
          versaoSistema: '2.0',
          dados: '',
        },
      };

      try {
        const response = await axios.post(API_URL_APOIAR, requestBody, {
          headers: {
            Authorization: `Bearer ${tokens.access_token}`,
            'Content-Type': 'application/json',
            jwt_token: tokens.jwt_token,
          },
        });

        const protocoloData = JSON.parse(response.data?.dados);
        console.log(`📌 Protocolo para ${cnpj}:`, protocoloData);
        protocolos.push({ cnpj, protocolo: protocoloData });
      } catch (error) {
        console.error(`❌ Erro ao solicitar protocolo para ${cnpj}:`, error.message);
        writeLog(errorLogPath, `Erro protocolo ${cnpj}: ${error.message}`);
      }

      await new Promise(resolve => setTimeout(resolve, 2000));
    }

    return protocolos;
  }

  // Extrair nome da empresa do PDF
  async function extractNomeEmpresaFromPDF(pdfBuffer) {
    try {
      const pdfData = await pdfParse(pdfBuffer);
      const match = pdfData.text.match(/CNPJ:\s*\d{2}\.\d{3}\.\d{3}\s*-\s*(.+)/i);
      return match && match[1] ? match[1].trim() : null;
    } catch {
      return null;
    }
  }

  function sanitizeFileName(fileName) {
    return fileName.replace(/[<>:"/\\|?*]+/g, '').trim();
  }

  // Emitir relatórios
  async function emitirRelatorios(protocolos, tokens) {
    for (const item of protocolos) {
      const requestBody = {
        contratante: { numero: CONTRATANTE_CNPJ, tipo: 2 },
        autorPedidoDados: { numero: CONTRATANTE_CNPJ, tipo: 2 },
        contribuinte: { numero: item.cnpj, tipo: 2 },
        pedidoDados: {
          idSistema: 'SITFIS',
          idServico: 'RELATORIOSITFIS92',
          versaoSistema: '2.0',
          dados: JSON.stringify({
            protocoloRelatorio: item.protocolo.protocoloRelatorio,
            tempoEspera: item.protocolo.tempoEspera || 60000,
          }),
        },
      };

      try {
        console.log(`📄 Emitindo relatório para ${item.cnpj}...`);
        
        const response = await axios.post(API_URL_EMITIR, requestBody, {
          headers: {
            Authorization: `Bearer ${tokens.access_token}`,
            'Content-Type': 'application/json',
            jwt_token: tokens.jwt_token,
          },
        });

        const responseData = response.data;
        const jsonFilePath = path.join(jsonDir, `${item.cnpj}.json`);
        fs.writeFileSync(jsonFilePath, JSON.stringify(responseData, null, 2));

        if (responseData.dados) {
          const dados = JSON.parse(responseData.dados);
          
          console.log(`📋 Dados recebidos para ${item.cnpj}:`, {
            temPDF: !!dados.pdf,
            mensagem: dados.mensagem || 'sem mensagem',
            status: dados.status || responseData.status
          });
          
          if (dados.pdf) {
            const pdfBuffer = Buffer.from(dados.pdf, 'base64');
            const nomeEmpresa = await extractNomeEmpresaFromPDF(pdfBuffer);
            const pdfFileName = sanitizeFileName(nomeEmpresa || item.cnpj) + '.pdf';
            const pdfFilePath = path.join(pdfDir, pdfFileName);
            fs.writeFileSync(pdfFilePath, pdfBuffer);

            console.log(`✅ PDF salvo: ${pdfFileName}`);
            writeLog(successLogPath, `PDF salvo: ${pdfFileName}`);
          } else {
            const motivoSemPDF = dados.mensagem || dados.descricao || 'Motivo não especificado';
            console.log(`⚠️ ${item.cnpj}: Sem PDF - ${motivoSemPDF}`);
            writeLog(errorLogPath, `Sem PDF para ${item.cnpj}: ${motivoSemPDF}\n${JSON.stringify(dados, null, 2)}`);
          }
        }
      } catch (error) {
        console.error(`❌ Erro ao emitir relatório para ${item.cnpj}:`, error.message);
        if (error.response?.data) {
          const errorJsonPath = path.join(jsonDir, `${item.cnpj}_ERROR.json`);
          fs.writeFileSync(errorJsonPath, JSON.stringify(error.response.data, null, 2));
        }
        writeLog(errorLogPath, `Erro relatório ${item.cnpj}: ${error.message}`);
      }

      await new Promise(resolve => setTimeout(resolve, 3000));
    }
  }

  // Execução principal
  console.log('🚀 Iniciando processamento...');
  
  const contribuintes = lerCNPJsDoExcel();
  if (contribuintes.length === 0) {
    throw new Error('Nenhum CNPJ encontrado no Excel');
  }

  const tokens = await getAuthTokens();
  const protocolos = await solicitarProtocolos(tokens, contribuintes);
  
  console.log(`📋 ${protocolos.length} protocolos válidos obtidos`);
  
  if (protocolos.length > 0) {
    await emitirRelatorios(protocolos, tokens);
  } else {
    console.log('⚠️ Nenhum protocolo válido para emitir relatórios');
  }

  console.log('🎉 Processamento concluído!');
  
  // ✅ REGISTRAR CONCLUSÃO NO LOG
  writeLog(successLogPath, '✅ Processamento concluído com sucesso!');
  writeLog(successLogPath, `📁 PDFs salvos em: ${pdfDir}`);
  writeLog(successLogPath, `📊 Total de protocolos processados: ${protocolos.length}`);
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 900,
    height: 650,
    frame: false,
    icon: path.join(basePath, 'public', 'icon.ico'),
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  mainWindow.loadURL('http://localhost:3000');
}

app.whenReady().then(async () => {
  try {
    setupWorkDirectory();
    await startServer();
    createWindow();
  } catch (error) {
    console.error('❌ Erro ao iniciar:', error);
    app.quit();
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (server) {
    server.close();
    console.log('🛑 Servidor encerrado.');
  }
  if (process.platform !== 'darwin') {
    app.quit();
  }
});