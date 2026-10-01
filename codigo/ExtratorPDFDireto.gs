/**
 * ============================================================
 * EXTRATOR PDF DIRETO — Módulo de teste para a GERADORA_APAC
 * ============================================================
 * Envia o PDF do laudo DIRETO para a API da Anthropic (sem OCR
 * intermediário, sem regex). O modelo "enxerga" a página e
 * devolve JSON estruturado, com sinalização de legibilidade
 * por campo e validação local de CNS/CPF (dígito verificador).
 *
 * As funções validarCNS() e validarCPF() ficam em Validacao.gs
 * (são usadas também pelo fluxo principal da Geradora).
 *
 * COMO USAR (teste em paralelo, sem mexer na Geradora atual):
 * 1. Crie uma pasta no Drive com 15–20 laudos que você JÁ
 *    processou e conferiu (o "gabarito").
 * 2. Cadastre o ID dessa pasta na propriedade do script
 *    FOLDER_TESTE_ID (Configurações do projeto ▸ Propriedades).
 * 3. Rode testeUnitario() primeiro com UM laudo.
 * 4. Depois rode testeComparativo(). Ela cria/preenche a aba
 *    "TESTE_PDF_DIRETO" com uma linha por laudo.
 * 5. Compare campo a campo com o que a Geradora atual extraiu.
 *    Só considere a troca se for igual ou melhor em TUDO.
 *
 * A chamada à IA passa por chamarIA_() (ver IA.gs), então o
 * consumo de tokens também entra na aba USO_IA.
 * ============================================================
 */

// ===================== CONFIGURAÇÃO =====================

const CONFIG_EXTRATOR = {
  MAX_TOKENS: 2000,
  ABA_RESULTADO: 'TESTE_PDF_DIRETO'
};

/** Pasta de laudos de teste (propriedade FOLDER_TESTE_ID). */
function pastaTeste_() {
  const id = lerConfig_('FOLDER_TESTE_ID', true);
  if (!id) {
    throw new Error(
      'Cadastre a propriedade FOLDER_TESTE_ID (ID de uma pasta do Drive com laudos de teste) ' +
      'em Configurações do projeto ▸ Propriedades do script.'
    );
  }
  return DriveApp.getFolderById(id);
}

// Campos que o modelo deve extrair do laudo.
const CAMPOS_LAUDO = [
  'nome_paciente',
  'nome_mae',
  'data_nascimento',      // DD/MM/AAAA
  'sexo',                 // M ou F
  'cns_paciente',         // 15 dígitos, sem espaços
  'cpf_paciente',         // 11 dígitos, sem pontuação (ou null se ausente)
  'endereco',
  'municipio',
  'cep',                  // 8 dígitos, sem hífen (ou null se ausente)
  'procedimento_solicitado',
  'codigo_sigtap',        // 10 dígitos, sem pontuação
  'cid10_principal',      // ex: M75.1
  'cid10_secundario',     // ou null
  'justificativa',        // resumo do quadro clínico
  'medico_solicitante',
  'cns_ou_crm_medico',
  'data_solicitacao'      // DD/MM/AAAA
];

// ===================== FUNÇÃO PRINCIPAL =====================

/**
 * Extrai os dados de um laudo em PDF enviando o arquivo direto
 * para a API. Retorna um objeto com:
 *   { arquivo, dados: {...}, alertas: [...], tokens: {...} }
 *
 * @param {string} fileId - ID do arquivo PDF no Drive
 */
function extrairDadosLaudoPDF(fileId) {
  const arquivo = DriveApp.getFileById(fileId);
  const blob = arquivo.getBlob();

  if (blob.getContentType() !== 'application/pdf') {
    throw new Error('O arquivo não é um PDF: ' + arquivo.getName());
  }

  const pdfBase64 = Utilities.base64Encode(blob.getBytes());
  const resposta = chamarClaudeComPDF_(pdfBase64);
  const dados = parsearJSON_(resposta.texto);

  return {
    arquivo: arquivo.getName(),
    dados: dados,
    alertas: alertasExtracao_(dados),
    tokens: resposta.tokens
  };
}

/** Validações locais (custo zero) + campos que o modelo marcou como duvidosos. */
function alertasExtracao_(dados) {
  const alertas = [];

  if (dados.cns_paciente && dados.cns_paciente.valor && !validarCNS(dados.cns_paciente.valor)) {
    alertas.push('CNS INVÁLIDO (' + motivoCNSInvalido_(dados.cns_paciente.valor) + '): ' +
                 dados.cns_paciente.valor);
  }
  if (dados.cpf_paciente && dados.cpf_paciente.valor && !validarCPF(dados.cpf_paciente.valor)) {
    alertas.push('CPF INVÁLIDO (não passa no dígito verificador): ' + dados.cpf_paciente.valor);
  }

  CAMPOS_LAUDO.forEach(function(campo) {
    if (dados[campo] && dados[campo].legivel === false) {
      alertas.push('CONFERIR MANUALMENTE (ilegível/duvidoso): ' + campo);
    }
  });
  return alertas;
}

// ===================== CHAMADA À IA =====================

function chamarClaudeComPDF_(pdfBase64) {
  const promptSistema =
    'Você é um extrator de dados de laudos médicos APAC do SUS brasileiro. ' +
    'Você recebe o PDF do laudo e devolve APENAS um objeto JSON válido, ' +
    'sem markdown, sem crase, sem texto antes ou depois.';

  const promptUsuario =
    'Extraia do laudo os seguintes campos: ' + CAMPOS_LAUDO.join(', ') + '.\n\n' +
    'REGRAS IMPORTANTES:\n' +
    '1. Para CADA campo, devolva um objeto: {"valor": "...", "legivel": true/false}.\n' +
    '   Use "legivel": false quando o dado estiver borrado, rasurado, ambíguo ' +
    'ou quando você tiver QUALQUER dúvida na leitura — principalmente em números.\n' +
    '2. Se o campo não existir no laudo, use {"valor": null, "legivel": true}.\n' +
    '3. NUNCA invente ou deduza dados. Transcreva exatamente o que está escrito.\n' +
    '4. Números (CNS, CPF, CEP, código SIGTAP) sem pontuação, espaços ou hífens.\n' +
    '5. Datas no formato DD/MM/AAAA.\n' +
    '6. Atenção redobrada na distinção entre 0/O, 1/l/I, 5/S, 8/B em números.\n' +
    '7. Responda SOMENTE o JSON.';

  const r = chamarIA_({
    funcao:    'EXTRATOR_TESTE',
    system:    promptSistema,
    maxTokens: CONFIG_EXTRATOR.MAX_TOKENS,
    content: [
      { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: pdfBase64 } },
      { type: 'text', text: promptUsuario }
    ]
  });

  return {
    texto: r.texto,
    tokens: { entrada: r.entrada, saida: r.saida }
  };
}

function parsearJSON_(texto) {
  // Remove cercas de markdown caso o modelo desobedeça
  const limpo = texto.replace(/```json/gi, '').replace(/```/g, '').trim();
  try {
    return JSON.parse(limpo);
  } catch (e) {
    throw new Error('A resposta não veio em JSON válido:\n' + texto);
  }
}

// ===================== TESTE COMPARATIVO =====================

/**
 * Roda a extração em todos os PDFs da pasta de teste e grava
 * os resultados na aba TESTE_PDF_DIRETO, uma linha por laudo,
 * para você comparar com o gabarito da Geradora atual.
 */
function testeComparativo() {
  const pasta = pastaTeste_();
  const arquivos = pasta.getFilesByType('application/pdf');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(CONFIG_EXTRATOR.ABA_RESULTADO);
  if (!aba) {
    const ativa = ss.getActiveSheet();
    aba = ss.insertSheet(CONFIG_EXTRATOR.ABA_RESULTADO);
    ss.setActiveSheet(ativa);
  }
  aba.clear();

  const cabecalho = ['arquivo'].concat(CAMPOS_LAUDO)
                    .concat(['alertas', 'tokens_entrada', 'tokens_saida']);
  aba.appendRow(cabecalho);
  aba.getRange(1, 1, 1, cabecalho.length).setFontWeight('bold');

  while (arquivos.hasNext()) {
    const arq = arquivos.next();
    try {
      const r = extrairDadosLaudoPDF(arq.getId());
      const linha = [r.arquivo];

      CAMPOS_LAUDO.forEach(function(campo) {
        const c = r.dados[campo];
        let valor = (c && c.valor !== null && c.valor !== undefined)
                    ? String(c.valor) : '';
        if (c && c.legivel === false) valor = '⚠ ' + valor;
        linha.push(valor);
      });

      linha.push(r.alertas.join(' | '));
      linha.push(r.tokens.entrada);
      linha.push(r.tokens.saida);
      aba.appendRow(linha);

      // Destaca em vermelho linhas com alerta
      if (r.alertas.length > 0) {
        aba.getRange(aba.getLastRow(), 1, 1, cabecalho.length)
           .setBackground('#FCE4E4');
      }
    } catch (e) {
      aba.appendRow([arq.getName(), 'ERRO: ' + e.message]);
      aba.getRange(aba.getLastRow(), 1, 1, 2).setBackground('#F4CCCC');
    }

    Utilities.sleep(1500); // pausa entre chamadas, evita rate limit
  }

  SpreadsheetApp.flush();
  Logger.log('Teste concluído. Confira a aba ' + CONFIG_EXTRATOR.ABA_RESULTADO);
}

/**
 * Teste rápido com UM único laudo. Cole o ID do PDF e rode.
 * O resultado sai no Registro de execução (Logger).
 */
function testeUnitario() {
  const FILE_ID = 'COLE_AQUI_O_ID_DE_UM_PDF';
  const r = extrairDadosLaudoPDF(FILE_ID);
  Logger.log(JSON.stringify(r, null, 2));
}

// ============ TESTE VIA UPLOAD (sem Drive) ============

// Abre o painel de teste — rode esta função ou use o menu Ferramentas
function abrirTesteUpload() {
  const html = HtmlService.createHtmlOutputFromFile('TesteUpload')
    .setWidth(500).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '🧪 Teste Extrator PDF Direto');
}

// Recebe o base64 do navegador e roda a extração completa
function testarExtracaoUpload(pdfBase64, nomeArquivo) {
  const resposta = chamarClaudeComPDF_(pdfBase64);
  const dados = parsearJSON_(resposta.texto);
  return { arquivo: nomeArquivo, dados: dados, alertas: alertasExtracao_(dados), tokens: resposta.tokens };
}
