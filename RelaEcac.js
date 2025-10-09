const axios = require("axios");
const https = require("https");
const fs = require("fs");
const path = require("path");
const pdfParse = require("pdf-parse");
const XLSX = require("xlsx");

// Carregar variáveis de ambiente
require('dotenv').config();

// ==========================
// CONFIGURAÇÕES (usando .env)
// ==========================
const CERT_PATH = path.join(__dirname, "");
const CERT_PASSWORD = "";
const AUTH_URL = process.env.AUTH_URL || "https://autenticacao.sapi.serpro.gov.br/authenticate";
const CONSUMER_KEY = process.env.CONSUMER_KEY;
const CONSUMER_SECRET = process.env.CONSUMER_SECRET;

const API_URL_APOIAR = process.env.API_URL_APOIAR || "https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Apoiar";
const API_URL_EMITIR = process.env.API_URL_EMITIR || "https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Emitir";

const CONTRATANTE_CNPJ = process.env.CONTRATANTE_CNPJ || "";

// Verificar se as variáveis essenciais estão definidas
if (!CONSUMER_KEY || !CONSUMER_SECRET) {
  console.error("❌ Erro: CONSUMER_KEY e CONSUMER_SECRET devem estar definidas no arquivo .env");
  process.exit(1);
}

console.log("🔐 Configurações carregadas do arquivo .env");
console.log(`📜 Certificado: ${path.basename(CERT_PATH)}`);
console.log(`🏢 Contratante: ${CONTRATANTE_CNPJ}`);

// ==========================
// LEITURA DOS CONTRIBUINTES VIA EXCEL
// ==========================
function lerCNPJsDoExcel() {
  const excelPath = path.join(__dirname, "empresas_filtradas.xlsx");
  
  if (!fs.existsSync(excelPath)) {
    console.error("❌ Arquivo empresas_filtradas.xlsx não encontrado!");
    return [];
  }

  const workbook = XLSX.readFile(excelPath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  // Coluna B (índice 2)
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const cnpjs = [];

  for (let row = 1; row <= range.e.r; row++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: 2 })]; // coluna B = índice 2
    if (cell && cell.v) {
      cnpjs.push(String(cell.v).replace(/\D/g, "")); // mantém só números
    }
  }

  console.log("📂 CNPJs carregados do Excel:", cnpjs.length);
  return cnpjs;
}

const contribuintes = lerCNPJsDoExcel();

// ==========================
// PASTAS DE SAÍDA
// ==========================
const pdfDir = path.join(__dirname, "ReportX");
const jsonDir = path.join(__dirname, "json_responses");
const logsDir = path.join(__dirname, "logs");
if (!fs.existsSync(pdfDir)) fs.mkdirSync(pdfDir);
if (!fs.existsSync(jsonDir)) fs.mkdirSync(jsonDir);
if (!fs.existsSync(logsDir)) fs.mkdirSync(logsDir);

const successLogPath = path.join(logsDir, "success.log");
const errorLogPath = path.join(logsDir, "errors.log");

function writeLog(filePath, message) {
  const now = new Date();
  const timestamp = now.toLocaleString("pt-BR");
  fs.appendFileSync(filePath, `[${timestamp}] ${message}\n`);
}

// ==========================
// 1. AUTENTICAÇÃO
// ==========================
async function getAuthTokens() {
  try {
    if (!fs.existsSync(CERT_PATH)) {
      throw new Error(`Certificado não encontrado: ${CERT_PATH}`);
    }

    const httpsAgent = new https.Agent({
      pfx: fs.readFileSync(CERT_PATH),
      passphrase: CERT_PASSWORD,
    });

    const authHeader = Buffer.from(`${CONSUMER_KEY}:${CONSUMER_SECRET}`).toString("base64");

    const response = await axios.post(
      AUTH_URL,
      new URLSearchParams({ grant_type: "client_credentials" }),
      {
        headers: {
          Authorization: `Basic ${authHeader}`,
          "role-type": "TERCEIROS",
          "content-type": "application/x-www-form-urlencoded",
        },
        httpsAgent,
      }
    );

    console.log("✅ Tokens obtidos com sucesso");
    return response.data;
  } catch (error) {
    console.error("❌ Erro ao obter tokens:", error.response?.data || error.message);
    process.exit(1);
  }
}

// ==========================
// 2. SOLICITAR PROTOCOLOS
// ==========================
async function solicitarProtocolos(tokens) {
  const protocolos = [];

  for (const cnpj of contribuintes) {
    const requestBody = {
      contratante: { numero: CONTRATANTE_CNPJ, tipo: 2 },
      autorPedidoDados: { numero: CONTRATANTE_CNPJ, tipo: 2 },
      contribuinte: { numero: cnpj, tipo: 2 },
      pedidoDados: {
        idSistema: "SITFIS",
        idServico: "SOLICITARPROTOCOLO91",
        versaoSistema: "2.0",
        dados: "",
      },
    };

    try {
      const response = await axios.post(API_URL_APOIAR, requestBody, {
        headers: {
          Authorization: `Bearer ${tokens.access_token}`,
          "Content-Type": "application/json",
          jwt_token: tokens.jwt_token,
        },
      });

      // Parse do protocolo retornado
      const protocoloData = JSON.parse(response.data?.dados);
      console.log(`📌 Protocolo para ${cnpj}:`, protocoloData);
      protocolos.push({ cnpj, protocolo: protocoloData });
    } catch (error) {
      console.error(`❌ Erro ao solicitar protocolo para ${cnpj}:`, error.response?.data || error.message);
      writeLog(errorLogPath, `Erro protocolo ${cnpj}: ${error.message}`);
    }

    await new Promise((resolve) => setTimeout(resolve, 2000)); // delay opcional
  }

  return protocolos;
}

// ==========================
// 3. EMITIR RELATÓRIOS
// ==========================
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
  return fileName.replace(/[<>:"/\\|?*]+/g, "").trim();
}

async function emitirRelatorios(protocolos, tokens) {
  for (const item of protocolos) {
    const requestBody = {
      contratante: { numero: CONTRATANTE_CNPJ, tipo: 2 },
      autorPedidoDados: { numero: CONTRATANTE_CNPJ, tipo: 2 },
      contribuinte: { numero: item.cnpj, tipo: 2 },
      pedidoDados: {
        idSistema: "SITFIS",
        idServico: "RELATORIOSITFIS92",
        versaoSistema: "2.0",
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
          "Content-Type": "application/json",
          jwt_token: tokens.jwt_token,
        },
      });

      const responseData = response.data;
      const jsonFilePath = path.join(jsonDir, `${item.cnpj}.json`);
      fs.writeFileSync(jsonFilePath, JSON.stringify(responseData, null, 2));

      if (responseData.dados) {
        const dados = JSON.parse(responseData.dados);
        if (dados.pdf) {
          const pdfBuffer = Buffer.from(dados.pdf, "base64");
          const nomeEmpresa = await extractNomeEmpresaFromPDF(pdfBuffer);
          const pdfFileName = sanitizeFileName(nomeEmpresa || item.cnpj) + ".pdf";
          const pdfFilePath = path.join(pdfDir, pdfFileName);
          fs.writeFileSync(pdfFilePath, pdfBuffer);

          console.log(`✅ PDF salvo: ${pdfFileName}`);
          writeLog(successLogPath, `PDF salvo: ${pdfFileName}`);
        } else {
          console.log(`⚠️ Resposta para ${item.cnpj} não contém PDF`);
          writeLog(errorLogPath, `Sem PDF para ${item.cnpj}: ${JSON.stringify(dados)}`);
        }
      }
    } catch (error) {
      console.error(`❌ Erro ao emitir relatório para ${item.cnpj}:`, error.response?.data || error.message);
      writeLog(errorLogPath, `Erro relatório ${item.cnpj}: ${error.message}`);
    }

    // Adicionar delay entre requisições para evitar rate limiting
    await new Promise((resolve) => setTimeout(resolve, 3000));
  }
}

// ==========================
// EXECUÇÃO PRINCIPAL
// ==========================
(async () => {
  try {
    console.log("🚀 Iniciando processamento com configurações do .env...");
    
    const tokens = await getAuthTokens();
    const protocolos = await solicitarProtocolos(tokens);
    
    console.log(`📋 ${protocolos.length} protocolos válidos obtidos`);
    
    if (protocolos.length > 0) {
      await emitirRelatorios(protocolos, tokens);
    } else {
      console.log("⚠️ Nenhum protocolo válido para emitir relatórios");
    }

    console.log("🎉 Processamento concluído. Confira as pastas 'pdfs', 'json_responses' e 'logs'");
  } catch (error) {
    console.error("💥 Erro durante execução:", error);
    writeLog(errorLogPath, `Erro geral: ${error.message}`);
  }
})();