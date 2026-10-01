/**
 * MÓDULO SIGTAP — "Buscar dados da APAC" + preenchimento automático da DADOS
 * Arquivo: Code_SIGTAP.gs
 * Roda na própria planilha GERADORA_APAC_EXTERNA — sem openById externo.
 * Os campos da aba DADOS são acessados pelo nome (ver Layout_DADOS.gs).
 * A chamada à IA passa por chamarIA_() (ver IA.gs).
 *
 * A tabela oficial usada é sempre a aba SIGTAP_AAAAMM mais recente
 * (monte uma nova em 🔧 Ferramentas ▸ 📥 Atualizar tabela SIGTAP).
 *
 * Busca do procedimento na Referencia_SIGTAP: compara PALAVRAS INTEIRAS
 * (ex.: "EDA" não combina mais com "MEDALHA") e, entre vários parecidos,
 * escolhe o que aparece primeiro no nome do laudo.
 */

const CONFIG_SIGTAP = {
  ABA_REFERENCIA: 'Referencia_SIGTAP',
  // Usa sozinho a aba SIGTAP_AAAAMM mais recente da planilha.
  // Para atualizar: 🔧 Ferramentas ▸ 📥 Atualizar tabela SIGTAP (.zip). Nada a mudar aqui.
  get ABA_SIGTAP() { return abaSigtapMaisRecente_() || 'SIGTAP_202608'; }
};

/** Nome da aba SIGTAP_AAAAMM mais recente da planilha, ou '' se não houver nenhuma. */
function abaSigtapMaisRecente_() {
  const nomes = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return /^SIGTAP_\d{6}$/.test(n); })
    .sort();
  return nomes.length ? nomes[nomes.length - 1] : '';
}

/**
 * Código SIGTAP sempre com 10 dígitos, como texto.
 * Se a coluna da planilha virou número, o Google apaga o zero da frente
 * (0209010037 vira 209010037); aqui o zero volta.
 */
function codigoSigtap10_(valor) {
  const dig = String(valor === null || valor === undefined ? '' : valor).replace(/\D/g, '');
  return dig && dig.length < 10 ? ('0000000000' + dig).slice(-10) : dig;
}

/** Competência da tabela oficial, tirada do nome da aba (ex.: "202608"). */
function competenciaSigtap_() {
  return CONFIG_SIGTAP.ABA_SIGTAP.replace(/^SIGTAP_/, '');
}

function carregarTabelaSigtap() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG_SIGTAP.ABA_REFERENCIA);
  if (!aba) throw new Error('Aba "' + CONFIG_SIGTAP.ABA_REFERENCIA + '" não encontrada.');

  return aba.getDataRange().getValues().slice(1).map(l => ({
    procedimento: String(l[0] || '').trim(),
    codigo:       codigoSigtap10_(l[1]),
    cids:         String(l[2] || '').split(',').map(c => c.trim()).filter(Boolean),
    confianca:    String(l[3] || '').trim()
  })).filter(item => item.procedimento && item.codigo);
}

function carregarTabelaOficial() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG_SIGTAP.ABA_SIGTAP);
  if (!aba) return null;

  const mapa = new Map();
  aba.getDataRange().getValues().slice(1).forEach(l => {
    const co  = codigoSigtap10_(l[0]);
    const no  = String(l[1] || '').trim();
    const hab = String(l[2] || '').trim();
    if (co) mapa.set(co, { nome: no, habilitacoes: hab });
  });
  return mapa;
}

// ② Abre a janela "Buscar dados da APAC"
function abrirDialogoSigtap() {
  const html = HtmlService.createHtmlOutputFromFile('DialogoSigtap')
    .setWidth(480)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, '🤖 Buscar dados da APAC');
}

function conferirCodigoSigtap(arquivoBase64, nomeArquivo) {
  const extraido   = extrairDadosLaudo(arquivoBase64);
  const tabela     = carregarTabelaSigtap();
  const tabelaOfic = carregarTabelaOficial();

  return compararComReferencia(extraido, tabela, tabelaOfic);
}

/**
 * Grava na aba DADOS:
 *   Código SIGTAP · Nome do procedimento · CID-10 principal
 *   Descrição do diagnóstico · Observações · Data da solicitação
 *   Profissional e estabelecimento solicitantes da APAC externa
 */
function preencherGeradora(dados) {
  try {
    if (dados.codigo)
      gravarCampo_('proc_codigo', codigoSigtap10_(dados.codigo));
    if (dados.nomeProcedimento)
      gravarCampo_('proc_nome', dados.nomeProcedimento.toUpperCase());
    // Quantidade padrão 1, só se ainda estiver em branco (não sobrescreve o que foi digitado)
    if ((dados.codigo || dados.nomeProcedimento) && !String(textoCampo_('proc_qtd') || '').trim())
      gravarCampo_('proc_qtd', 1);
    if (dados.cid)
      gravarCampo_('cid_principal', dados.cid.toUpperCase());
    if (dados.descricaoDiagnostico)
      gravarCampo_('diag_descricao', dados.descricaoDiagnostico.toUpperCase());
    if (dados.observacoes)
      gravarCampo_('observacoes', dados.observacoes.toUpperCase());
    if (dados.dataSolicitacao)
      gravarCampo_('esf_data', dados.dataSolicitacao);
    if (dados.medicoSolicitante)
      gravarCampo_('ext_profissional', dados.medicoSolicitante.toUpperCase());
    if (dados.estabelecimentoSolicitante && dados.estabelecimentoSolicitante.trim() !== '')
      gravarCampo_('ext_estabelecimento', dados.estabelecimentoSolicitante.toUpperCase());

    SpreadsheetApp.flush();
    return { ok: true };

  } catch (err) {
    return { ok: false, msg: 'Erro ao preencher a aba DADOS: ' + err.message };
  }
}

function extrairDadosLaudo(arquivoBase64) {
  const systemPrompt =
    'Você é especialista em laudos APAC do SUS. Extraia exatamente os campos abaixo do PDF recebido.\n' +
    'Responda SOMENTE com JSON válido, sem texto antes ou depois, sem markdown:\n' +
    '{\n' +
    '  "paciente": "nome ou null",\n' +
    '  "procedimento": "nome exato do procedimento principal como consta no laudo",\n' +
    '  "codigo_no_laudo": "código de 10 dígitos ou null se ausente/ilegível",\n' +
    '  "cid": "código CID-10 principal ou null",\n' +
    '  "descricao_diagnostico": "texto do campo Descrição do Diagnóstico ou null",\n' +
    '  "estabelecimento_solicitante": "nome do estabelecimento solicitante ou null",\n' +
    '  "observacoes": "texto completo do campo Observações ou null",\n' +
    '  "data_solicitacao": "data no campo Data da Solicitação no formato DD/MM/AAAA ou null",\n' +
    '  "medico_solicitante": "nome completo do profissional solicitante ou null"\n' +
    '}';

  const r = chamarIA_({
    funcao:    'SIGTAP',
    system:    systemPrompt,
    maxTokens: 1200,
    content: [
      { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: arquivoBase64 }, title: 'Laudo APAC' },
      { type: 'text',     text: 'Extraia os dados do laudo conforme instruído.' }
    ]
  });

  if (r.parou === 'max_tokens')
    throw new Error('A resposta da IA foi cortada (laudo com texto muito longo). Tente novamente.');

  return interpretarJsonIA_(r.texto);
}

function normalizarTexto(s) {
  return String(s || '').trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

/**
 * Nome pronto para comparar por PALAVRAS: sem acento, sem pontuação,
 * espaços simples e um espaço em cada ponta (" EDA COM BIOPSIA ").
 * Assim " EDA " só combina com a palavra EDA, e não com "MEDALHA".
 */
function nomeParaComparar_(s) {
  return ' ' + normalizarTexto(s).replace(/[^A-Z0-9]+/g, ' ').trim() + ' ';
}

/**
 * Procura o procedimento do laudo na Referencia_SIGTAP.
 *   1. nome igual (ignorando acento, pontuação e espaços)
 *   2. nome contido no outro, por PALAVRAS INTEIRAS. Entre vários,
 *      fica o que aparece MAIS CEDO no nome do laudo (o procedimento
 *      principal costuma vir primeiro); empatando, o mais completo
 *      (mais longo); empatando de novo, o ✅ Confirmado.
 */
function buscarProcedimento(nomeLaudo, tabela) {
  const alvo = nomeParaComparar_(nomeLaudo);
  if (alvo.trim() === '') return { match: null, tipo: 'nenhuma' };

  let match = tabela.find(t => nomeParaComparar_(t.procedimento) === alvo);
  if (match) return { match, tipo: 'exata' };

  let melhor = null, melhorNota = null;
  tabela.forEach(t => {
    const ref = nomeParaComparar_(t.procedimento);
    if (ref.trim() === '') return;

    let posicao;
    if (alvo.indexOf(ref) !== -1)      posicao = alvo.indexOf(ref);   // nome da tabela dentro do laudo
    else if (ref.indexOf(alvo) !== -1) posicao = 0;                   // nome do laudo dentro da tabela
    else return;

    const nota = [posicao, -Math.min(ref.length, alvo.length), /CONFIRMADO/i.test(t.confianca || '') ? 0 : 1];
    if (!melhorNota || nota[0] < melhorNota[0] ||
        (nota[0] === melhorNota[0] && (nota[1] < melhorNota[1] ||
        (nota[1] === melhorNota[1] && nota[2] < melhorNota[2])))) {
      melhor = t; melhorNota = nota;
    }
  });
  if (melhor) return { match: melhor, tipo: 'parcial' };

  return { match: null, tipo: 'nenhuma' };
}

function compararComReferencia(extraido, tabela, tabelaOfic) {
  const busca       = buscarProcedimento(extraido.procedimento, tabela);
  const codigoLaudo = extraido.codigo_no_laudo
    ? codigoSigtap10_(extraido.codigo_no_laudo) : null;

  let codigoLaudoExisteOfic    = false;
  let nomeOficialDoCodigoLaudo = null;

  if (codigoLaudo && tabelaOfic) {
    const reg = tabelaOfic.get(codigoLaudo);
    if (reg) { codigoLaudoExisteOfic = true; nomeOficialDoCodigoLaudo = reg.nome; }
  }

  let status, mensagem, codigoSugerido = null, cidStatus = null, avisoCodigoInvalido = null;

  if (codigoLaudo && tabelaOfic && !codigoLaudoExisteOfic)
    avisoCodigoInvalido = '⚠️ O código ' + codigoLaudo +
      ' não existe na Tabela SIGTAP (competência ' + competenciaSigtap_() + '). Pode ser erro de digitação.';

  if (!busca.match) {
    status = 'nao_catalogado';
    if (codigoLaudo && codigoLaudoExisteOfic)
      mensagem = 'Procedimento não está na tabela local, mas o código ' + codigoLaudo +
                 ' existe no SIGTAP oficial como: "' + nomeOficialDoCodigoLaudo +
                 '". Considere adicionar à aba Referencia_SIGTAP após confirmar.';
    else if (codigoLaudo && tabelaOfic)
      mensagem = 'Procedimento não está na tabela local e o código ' + codigoLaudo +
                 ' não foi encontrado no SIGTAP oficial. Verifique em sigtap.datasus.gov.br.';
    else
      mensagem = 'Procedimento não encontrado na tabela de referência local. Consulte o SIGTAP oficial.';

  } else {
    codigoSugerido  = busca.match.codigo;
    const confLabel = busca.match.confianca || '';
    const prefConf  = confLabel.includes('REVISAR')
      ? '🔴 Este registro está marcado para revisão na tabela local. ' : '';
    const nomeRef   = busca.tipo === 'parcial'
      ? ' (parecido com "' + busca.match.procedimento + '" da tabela local)' : '';

    if (!codigoLaudo) {
      status   = 'codigo_ausente';
      mensagem = prefConf + 'O laudo não trouxe o código. Com base em "' +
        extraido.procedimento + '"' + nomeRef + ', o código sugerido é ' +
        codigoSugerido + '. ' +
        (tabelaOfic && tabelaOfic.has(codigoSugerido) ? 'Confirmado no SIGTAP oficial. ' : '') +
        'Confirme antes de lançar.';

    } else if (codigoLaudo === codigoSugerido) {
      status   = 'ok';
      mensagem = prefConf + 'Código confere. "' + extraido.procedimento + '" = ' + codigoLaudo +
        ', consistente com a referência local' +
        (codigoLaudoExisteOfic ? ' e com a Tabela SIGTAP oficial.' : '.');
      if (avisoCodigoInvalido) mensagem += ' ' + avisoCodigoInvalido;

    } else {
      status   = 'divergente';
      mensagem = prefConf + 'Código do laudo (' + codigoLaudo +
        ') diverge do código de referência (' + codigoSugerido + ') para "' +
        extraido.procedimento + '"' + nomeRef + '. ';
      if (tabelaOfic) {
        const regLaudo = tabelaOfic.get(codigoLaudo);
        const regRef   = tabelaOfic.get(codigoSugerido);
        mensagem += regLaudo
          ? 'O código do laudo corresponde a "' + regLaudo.nome + '" no SIGTAP. '
          : 'O código do laudo NÃO existe no SIGTAP oficial. ';
        if (regRef) mensagem += 'O código de referência corresponde a "' + regRef.nome + '". ';
      }
      mensagem += 'Verifique manualmente antes de lançar.';
    }

    if (extraido.cid && busca.match.cids && busca.match.cids.length > 0) {
      const cidNorm  = normalizarTexto(extraido.cid).replace(/\./g, '');
      const cidsRefN = busca.match.cids.map(c => normalizarTexto(c).replace(/\./g, ''));
      cidStatus = cidsRefN.some(c => c === cidNorm || cidNorm.startsWith(c) || c.startsWith(cidNorm))
        ? 'cid_conhecido' : 'cid_novo';
    }
  }

  return {
    paciente:                    extraido.paciente,
    procedimento_laudo:          extraido.procedimento,
    nome_oficial_sigtap:         nomeOficialDoCodigoLaudo,
    codigo_laudo:                codigoLaudo,
    codigo_referencia:           codigoSugerido,
    cid_laudo:                   extraido.cid,
    descricao_diagnostico:       extraido.descricao_diagnostico || null,
    observacoes:                 extraido.observacoes           || null,
    data_solicitacao:            extraido.data_solicitacao      || null,
    medico_solicitante:          extraido.medico_solicitante    || null,
    estabelecimento_solicitante: extraido.estabelecimento_solicitante || null,
    cid_status:                  cidStatus,
    status:                      status,
    mensagem:                    mensagem,
    tipo_correspondencia:        busca.tipo,
    confianca_referencia:        busca.match ? busca.match.confianca : null
  };
}

// ============================================================
// 🔍 BUSCAR PROCEDIMENTO SIGTAP (painel lateral)
// ============================================================

function abrirBuscaSIGTAP() {
  const html = HtmlService.createHtmlOutputFromFile('BuscaSIGTAP')
    .setTitle('Busca SIGTAP')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Retorna todos os registros da aba SIGTAP: [codigo, nome, numeroLinha]
function buscarTodosSIGTAP() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SIGTAP.ABA_SIGTAP);
  if (!aba) throw new Error('Aba "' + CONFIG_SIGTAP.ABA_SIGTAP + '" não encontrada.');

  const dados = aba.getDataRange().getValues();
  const resultado = [];

  for (let i = 0; i < dados.length; i++) {
    const nome   = String(dados[i][1]).trim();
    if (String(dados[i][0]).trim() === 'CO_PROCEDIMENTO' || nome === 'NO_PROCEDIMENTO') continue;
    const codigo = codigoSigtap10_(dados[i][0]);
    if (!codigo || !nome) continue;
    resultado.push([codigo, nome, i + 1]);
  }
  return resultado;
}

// Chamada ao clicar num resultado da busca
function navegarParaLinha(linha) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG_SIGTAP.ABA_SIGTAP);
  if (!aba) throw new Error('Aba "' + CONFIG_SIGTAP.ABA_SIGTAP + '" não encontrada.');

  ss.setActiveSheet(aba);
  aba.getRange(linha, 1, 1, Math.max(aba.getLastColumn(), 1)).activate();
}
