/**
 * ============================================================
 * Atualizador.gs — 🔄 ATUALIZAR SISTEMA (código oficial no GitHub)
 * ============================================================
 * Menu 📋 APAC ▸ 🔧 Ferramentas ▸ 🔄 Atualizar sistema
 *
 * O código oficial fica no repositório público:
 *   github.com/tecnicoacslima-sys/geradora-apac
 * O arquivo versao.json diz qual é a versão oficial, o que mudou e
 * quais arquivos fazem parte dela.
 *
 * Ao atualizar:
 *   1. baixa o versao.json e os arquivos listados nele;
 *   2. confere cada arquivo (nome e tipo permitidos, conteúdo não vazio);
 *   3. guarda uma CÓPIA de todo o código atual no Drive
 *      (pasta APAC_EXTERNA ▸ BACKUP_CODIGO), para poder voltar;
 *   4. substitui SÓ os arquivos da lista; os outros ficam como estão.
 * NUNCA mexe nos dados das abas, nas propriedades (código de acesso
 * à IA, pastas, nome da unidade) nem nos gatilhos.
 *
 * Requisito (uma vez por pessoa): ligar a "API do Google Apps Script"
 * em https://script.google.com/home/usersettings
 * ============================================================
 */

const ATUALIZADOR = {
  VERSAO_LOCAL: '2.5.1',      // versão DESTE código; muda junto com o versao.json a cada publicação
  BASE_RAW:     'https://raw.githubusercontent.com/tecnicoacslima-sys/geradora-apac/main/',
  MANIFESTO:    'versao.json',
  PASTA_BACKUP: 'BACKUP_CODIGO',
  URL_CHAVE_API: 'https://script.google.com/home/usersettings',
  TIPOS:        ['SERVER_JS', 'HTML', 'JSON'],
  ESSENCIAIS:   ['Código', 'Atualizador', 'appsscript']   // sem eles a planilha não funcionaria
};

// ============================================================
// JANELA
// ============================================================

function abrirAtualizador() {
  const html = HtmlService.createHtmlOutputFromFile('DialogoAtualizador')
    .setWidth(500)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '🔄 Atualizar sistema');
}

/** Informações para a janela: versão local, versão oficial, novidades, se há backup. */
function infoAtualizacao() {
  const props = PropertiesService.getScriptProperties();
  const info  = {
    local:      ATUALIZADOR.VERSAO_LOCAL,
    remota:     '',
    data:       '',
    novidades:  [],
    haNova:     false,
    backup:     props.getProperty('BACKUP_CODIGO_VERSAO') || '',
    urlChave:   ATUALIZADOR.URL_CHAVE_API,
    erro:       ''
  };
  try {
    const m = lerManifesto_();
    info.remota    = m.versao;
    info.data      = m.data || '';
    info.novidades = m.novidades || [];
    info.haNova    = compararVersoes_(m.versao, ATUALIZADOR.VERSAO_LOCAL) > 0;
    if (info.haNova) props.setProperty('ATUALIZACAO_DISPONIVEL', m.versao);
    else             props.deleteProperty('ATUALIZACAO_DISPONIVEL');
  } catch (err) {
    info.erro = err.message;
  }
  return info;
}

// ============================================================
// LEITURA DO GITHUB
// ============================================================

function baixarTexto_(caminho) {
  const url = ATUALIZADOR.BASE_RAW + caminho + '?t=' + Date.now();
  const r   = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (r.getResponseCode() !== 200) {
    throw new Error('Não consegui baixar "' + caminho + '" do GitHub (código ' + r.getResponseCode() + ').');
  }
  return r.getContentText('UTF-8');
}

/** Lê e confere o versao.json. */
function lerManifesto_() {
  let m;
  try {
    m = JSON.parse(baixarTexto_(ATUALIZADOR.MANIFESTO));
  } catch (err) {
    throw new Error('Não consegui ler a versão oficial no GitHub. Confira a internet e tente de novo. (' + err.message + ')');
  }
  if (!m || !/^\d+(\.\d+)*$/.test(String(m.versao || '')) || !Array.isArray(m.arquivos) || !m.arquivos.length) {
    throw new Error('O arquivo de versão do GitHub está incompleto. Avise o suporte.');
  }
  m.arquivos.forEach(function(a) {
    if (!/^[A-Za-zÀ-ÿ0-9_]+$/.test(String(a.nome || ''))) throw new Error('Nome de arquivo não permitido na lista oficial: ' + a.nome);
    if (ATUALIZADOR.TIPOS.indexOf(a.tipo) === -1)          throw new Error('Tipo de arquivo não permitido: ' + a.tipo);
    if (a.tipo === 'JSON' && a.nome !== 'appsscript')       throw new Error('Só o appsscript pode ser do tipo JSON.');
    if (!/^[A-Za-z0-9_\-\/]+\.(gs|html|json)$/.test(String(a.caminho || ''))) throw new Error('Caminho não permitido: ' + a.caminho);
  });
  return m;
}

/** "2.10" > "2.9" → 1 · iguais → 0 · menor → -1 */
function compararVersoes_(a, b) {
  const pa = String(a).split('.').map(Number), pb = String(b).split('.').map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const x = pa[i] || 0, y = pb[i] || 0;
    if (x !== y) return x > y ? 1 : -1;
  }
  return 0;
}

// ============================================================
// LEITURA E GRAVAÇÃO DO CÓDIGO DESTA PLANILHA (API do Apps Script)
// ============================================================

function chamarApiScript_(metodo, corpo) {
  const url = 'https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId() + '/content';
  const opcoes = {
    method:             metodo,
    contentType:        'application/json',
    headers:            { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  if (corpo) opcoes.payload = JSON.stringify(corpo);

  const r     = UrlFetchApp.fetch(url, opcoes);
  const texto = r.getContentText();
  if (r.getResponseCode() === 200) return JSON.parse(texto);

  if (/Apps Script API|has not been used|not enabled|usersettings/i.test(texto)) {
    throw new Error(
      'Falta ligar a "API do Google Apps Script" para a sua conta. ' +
      'Abra ' + ATUALIZADOR.URL_CHAVE_API + ', ligue a chave "API do Google Apps Script" e tente de novo ' +
      '(pode levar alguns minutos para valer).'
    );
  }
  throw new Error('O Google recusou a operação (código ' + r.getResponseCode() + '): ' + texto.substring(0, 250));
}

/** Arquivos atuais do projeto: [{ name, type, source }] */
function codigoAtual_() {
  const c = chamarApiScript_('get');
  return (c.files || []).map(function(f) { return { name: f.name, type: f.type, source: f.source }; });
}

/** Guarda uma cópia do código atual no Drive e devolve o nome do arquivo. */
function guardarBackup_(arquivos) {
  let raiz;
  try { raiz = garantirPastaRaiz_(); } catch (err) { raiz = DriveApp.getRootFolder(); }
  const pasta = obterOuCriarPasta_(ATUALIZADOR.PASTA_BACKUP, raiz).pasta;

  const tz   = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const nome = 'codigo_v' + ATUALIZADOR.VERSAO_LOCAL + '_' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm') + '.json';
  const arq  = pasta.createFile(nome, JSON.stringify({ versao: ATUALIZADOR.VERSAO_LOCAL, files: arquivos }), 'application/json');

  const props = PropertiesService.getScriptProperties();
  props.setProperty('BACKUP_CODIGO_ID', arq.getId());
  props.setProperty('BACKUP_CODIGO_VERSAO', ATUALIZADOR.VERSAO_LOCAL);
  return nome;
}

// ============================================================
// ATUALIZAR E VOLTAR (chamados pela janela)
// ============================================================

function aplicarAtualizacao() {
  const trava = LockService.getScriptLock();
  if (!trava.tryLock(5000)) throw new Error('Já existe uma atualização em andamento. Aguarde.');
  try {
    // 1. Versão oficial e arquivos novos
    const m = lerManifesto_();
    if (compararVersoes_(m.versao, ATUALIZADOR.VERSAO_LOCAL) <= 0) {
      throw new Error('Esta planilha já está na versão ' + ATUALIZADOR.VERSAO_LOCAL + ' (a oficial é ' + m.versao + ').');
    }
    const novos = m.arquivos.map(function(a) {
      const fonte = baixarTexto_(a.caminho);
      if (!fonte || fonte.trim().length < 20) throw new Error('O arquivo "' + a.nome + '" veio vazio do GitHub. Nada foi alterado.');
      if (a.tipo === 'JSON') JSON.parse(fonte);       // appsscript.json precisa ser JSON válido
      return { name: a.nome, type: a.tipo, source: fonte };
    });

    // 2. Código atual + cópia de segurança
    const atuais = codigoAtual_();
    const backup = guardarBackup_(atuais);

    // 3. Junta: troca/acrescenta só os da lista; os outros ficam
    const chave  = function(f) { return f.type + '|' + f.name; };
    const mapa   = {};
    atuais.forEach(function(f) { mapa[chave(f)] = f; });
    novos.forEach(function(f)  { mapa[chave(f)] = f; });
    const final  = Object.keys(mapa).map(function(k) { return mapa[k]; });

    ATUALIZADOR.ESSENCIAIS.forEach(function(n) {
      if (!final.some(function(f) { return f.name === n; })) throw new Error('A atualização ficaria sem o arquivo "' + n + '". Nada foi alterado.');
    });

    // 4. Grava
    chamarApiScript_('put', { files: final });
    PropertiesService.getScriptProperties().deleteProperty('ATUALIZACAO_DISPONIVEL');

    return {
      ok:        true,
      de:        ATUALIZADOR.VERSAO_LOCAL,
      para:      m.versao,
      arquivos:  novos.map(function(f) { return f.name; }),
      backup:    backup
    };
  } finally {
    trava.releaseLock();
  }
}

/** Volta o código para a cópia guardada antes da última atualização. */
function voltarVersaoAnterior() {
  const props = PropertiesService.getScriptProperties();
  const id    = props.getProperty('BACKUP_CODIGO_ID');
  if (!id) throw new Error('Não há cópia de segurança guardada nesta planilha.');

  let dados;
  try {
    dados = JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8'));
  } catch (err) {
    throw new Error('Não consegui abrir a cópia de segurança no Drive (pasta ' + ATUALIZADOR.PASTA_BACKUP + ').');
  }
  if (!dados || !Array.isArray(dados.files) || !dados.files.length) throw new Error('A cópia de segurança está vazia.');

  chamarApiScript_('put', { files: dados.files });
  props.deleteProperty('BACKUP_CODIGO_ID');
  props.deleteProperty('BACKUP_CODIGO_VERSAO');
  return { ok: true, versao: dados.versao };
}

// ============================================================
// CONFERÊNCIA SEMANAL (gatilho) — só avisa, nunca atualiza sozinha
// ============================================================

function verificarAtualizacaoSemanal() {
  try {
    const m     = lerManifesto_();
    const props = PropertiesService.getScriptProperties();
    if (compararVersoes_(m.versao, ATUALIZADOR.VERSAO_LOCAL) > 0) props.setProperty('ATUALIZACAO_DISPONIVEL', m.versao);
    else props.deleteProperty('ATUALIZACAO_DISPONIVEL');
  } catch (err) {
    Logger.log('Verificação de atualização: ' + err);
  }
}
