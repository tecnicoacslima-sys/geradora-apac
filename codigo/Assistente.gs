/**
 * ============================================================
 * Assistente.gs — ASSISTENTE DE PRIMEIRA CONFIGURAÇÃO
 * ============================================================
 * Janela: DialogoAssistente.html.
 * Uma janela só, no lugar de "⚙️ Configurar sistema" + "🔑 Configurar
 * acesso à IA". A unidade informa nome, município, UF e o código de
 * acesso; o endereço do servidor de IA já vem preenchido.
 *
 * Ao clicar em Concluir, faz tudo de uma vez:
 *   1. grava os dados da unidade e o acesso à IA (propriedades do script)
 *   2. cria as pastas no Drive de quem está usando
 *   3. escreve o nome da unidade nos cabeçalhos (DADOS e PROTOCOLO_APAC)
 *   4. instala os gatilhos (buscas automáticas e painel ao abrir)
 *   5. testa a IA e mostra o resultado de cada passo
 *
 * Enquanto a planilha não estiver configurada, o menu 📋 APAC mostra
 * "🚀 Primeira configuração (comece aqui)" em primeiro lugar.
 * ============================================================
 */

const ASSISTENTE = {
  // Endereço do servidor de IA (Web App do PAINEL_CONSUMO_IA).
  // Se um dia mudar, troque só aqui.
  SERVIDOR_PADRAO: 'https://script.google.com/macros/s/AKfycbwwJnbX1Q3E_20qKbWER19A2juZB2RiqhcLhsvJBbznT25TapAwy7BaAJdv3D2nd89Brw/exec'
};

/** A planilha já foi configurada? (unidade, pastas e acesso à IA) */
function sistemaConfigurado_() {
  const p = PropertiesService.getScriptProperties();
  const temIA = (p.getProperty('PROXY_URL') && p.getProperty('PROXY_CODIGO')) || p.getProperty('CLAUDE_API_KEY');
  return !!(p.getProperty('UNIDADE_NOME') && p.getProperty('FOLDER_ENTRADA_ID') && temIA);
}

// Abre a janela (menu 🚀)
function abrirAssistente() {
  const html = HtmlService.createHtmlOutputFromFile('DialogoAssistente')
    .setWidth(480)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, '🚀 Configuração da Geradora de APAC');
}

/** Valores atuais, para preencher a janela (o código nunca volta para a tela). */
function dadosAssistente() {
  const p = PropertiesService.getScriptProperties();
  return {
    unidade:    p.getProperty('UNIDADE_NOME')   || '',
    municipio:  p.getProperty('MUNICIPIO_NOME') || '',
    uf:         p.getProperty('UF_SIGLA')       || '',
    url:        p.getProperty('PROXY_URL')      || ASSISTENTE.SERVIDOR_PADRAO,
    temCodigo:  !!p.getProperty('PROXY_CODIGO'),
    configurado: sistemaConfigurado_()
  };
}

/**
 * Chamado pelo botão Concluir.
 * d = { unidade, municipio, uf, codigo, url }
 * Devolve { tudoOk, passos: [{ tipo: 'ok'|'aviso'|'erro', texto }] }
 */
function concluirAssistente(d) {
  const props = PropertiesService.getScriptProperties();
  const limpar = function(t) { return String(t || '').replace(/\s+/g, ' ').trim(); };

  // ---- Conferência do que foi digitado ----
  const unidade   = limpar(d.unidade);
  const municipio = limpar(d.municipio);
  const uf        = limpar(d.uf).toUpperCase();
  let url         = limpar(d.url) || ASSISTENTE.SERVIDOR_PADRAO;
  let codigo      = String(d.codigo || '').replace(/\s+/g, '');

  if (!unidade)                 throw new Error('Digite o nome da unidade.');
  if (!municipio)               throw new Error('Digite o município.');
  if (!/^[A-Z]{2}$/.test(uf))   throw new Error('Digite a sigla do estado com 2 letras (ex.: MT).');
  if (!/^https:\/\/script\.google\.com\/(macros|a\/macros\/[^\/\s]+)\/s\/[^\/\s]+\/exec$/.test(url)) {
    throw new Error('O endereço do servidor não parece certo. Ele começa com https://script.google.com/ e termina em /exec.');
  }
  if (!codigo) codigo = props.getProperty('PROXY_CODIGO') || '';
  if (!codigo) throw new Error('Digite o código de acesso à IA que o suporte enviou.');

  // ---- 1. Grava os dados ----
  props.setProperties({
    UNIDADE_NOME:   unidade,
    MUNICIPIO_NOME: municipio,
    UF_SIGLA:       uf,
    PROXY_URL:      url,
    PROXY_CODIGO:   codigo
  });

  const passos = [];
  const ok    = function(t) { passos.push({ tipo: 'ok',    texto: t }); };
  const aviso = function(t) { passos.push({ tipo: 'aviso', texto: t }); };
  const erro  = function(t) { passos.push({ tipo: 'erro',  texto: t }); };

  ok('Dados da unidade gravados: ' + subtituloUnidade_(' · '));

  // ---- 2. Pastas no Drive ----
  try {
    const raiz = garantirPastaRaiz_();
    let criadas = 0;
    PASTAS_PADRAO.forEach(function(def) { if (garantirPasta_(def, raiz).criada) criadas++; });
    ok('Pastas no seu Drive, dentro de "' + PASTA_RAIZ_NOME + '"' +
       (criadas ? ' (' + criadas + ' criada(s) agora)' : ' (já existiam)'));
  } catch (err) {
    erro('Não consegui criar as pastas no Drive: ' + err.message);
  }

  // ---- 3. Cabeçalhos ----
  try {
    atualizarCabecalhos_();
    ok('Nome da unidade nos cabeçalhos das abas DADOS e PROTOCOLO_APAC');
  } catch (err) {
    aviso('Não consegui atualizar os cabeçalhos: ' + err.message);
  }

  // ---- 4. Gatilhos ----
  try {
    const inst = instalarGatilhos_();
    ok(inst.length ? 'Gatilhos instalados: ' + inst.join(' e ') : 'Gatilhos já estavam ativos');
  } catch (err) {
    erro('Não consegui instalar os gatilhos: ' + err.message);
  }

  // ---- Serviço Drive API (necessário para ler o Espelho CELK) ----
  if (typeof Drive === 'undefined') {
    aviso('Falta ativar o serviço "Drive API" (v2) no Apps Script. Sem ele o Espelho CELK não é lido. Avise o suporte.');
  }

  // ---- 5. Teste da IA ----
  try {
    chamarIA_({
      funcao:    'TESTE',
      system:    'Responda apenas com a palavra OK.',
      content:   [{ type: 'text', text: 'Teste de conexão.' }],
      maxTokens: 10
    });
    ok('Acesso à IA funcionando');
  } catch (err) {
    erro('A IA não respondeu: ' + err.message +
         ' Confira o código de acesso e clique em Concluir de novo.');
  }

  return {
    tudoOk: passos.every(function(p) { return p.tipo !== 'erro'; }),
    passos: passos
  };
}
