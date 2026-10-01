/**
 * ============================================================
 * SugestoesSIGTAP.gs — SUGESTÕES PARA A REFERENCIA_SIGTAP
 * ============================================================
 * Compara o que foi lançado no CHECK_LIST com a Referencia_SIGTAP
 * e monta a aba SUGESTOES_SIGTAP. Nada entra sozinho: você decide.
 *
 * Menu 📋 APAC ▸ 🔧 Ferramentas:
 *   💡 Sugestões para a Referência SIGTAP → monta/atualiza a aba
 *   ✅ Aplicar sugestões aprovadas        → leva as aprovadas para a Referência
 *
 * Tipos de sugestão:
 *   🆕 Procedimento novo  → o nome não existe na Referência
 *   🔄 Código diferente   → o nome existe, mas foi lançado com outro código
 *                           (pode ser mudança do regulador OU erro de digitação)
 *   ⚠️ Código não existe  → o código lançado não está na tabela SIGTAP
 *                           (quase sempre erro de digitação; não pode ser aprovado)
 *
 * Você pode CORRIGIR o "Código lançado" (coluna D) antes de aprovar:
 * é esse código que vai para a Referência.
 *
 * Regras de segurança:
 *   • só entra o que estiver marcado "✅ Aprovar";
 *   • "❌ Ignorar" fica lembrado e não volta como novidade;
 *   • registro "✅ Confirmado" na Referência NUNCA é trocado aqui;
 *   • roda sozinho todo dia 1º (gatilho instalado junto com os outros).
 * Não usa IA: não gasta tokens.
 * ============================================================
 */

const SUGESTOES = {
  ABA:       'SUGESTOES_SIGTAP',
  ABA_CHECK: 'CHECK_LIST',
  APROVAR:   '✅ Aprovar',
  IGNORAR:   '❌ Ignorar',
  TIPO_NOVO: '🆕 Procedimento novo',
  TIPO_DIF:  '🔄 Código diferente',
  TIPO_INEX: '⚠️ Código não existe no SIGTAP',
  CABECALHO: ['Decisão', 'Tipo', 'Procedimento (como no CHECK_LIST)', 'Código lançado',
              'Nome oficial do código lançado', 'Código na Referência',
              'Nome oficial do código da Referência', 'CIDs', 'Vezes', 'Última data']
};

// Colunas da aba SUGESTOES_SIGTAP (começando em 0)
const COL_SUG = { DECISAO: 0, TIPO: 1, NOME: 2, CODIGO: 3, OFIC: 4, REF: 5, OFIC_REF: 6, CIDS: 7, VEZES: 8, DATA: 9 };

function chaveSugestao_(nome, codigo) {
  return nomeParaComparar_(nome).trim() + '|' + codigoSigtap10_(codigo);
}

function cidLimpo_(v) {
  const c = String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  return /^[A-Z]\d{2,3}$/.test(c) ? c : '';
}

// ============================================================
// MONTAR AS SUGESTÕES
// ============================================================

/** Lê o CHECK_LIST e agrupa o que não bate com a Referência. */
function montarSugestoes_() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const abaCheck = ss.getSheetByName(SUGESTOES.ABA_CHECK);
  if (!abaCheck) throw new Error('Aba ' + SUGESTOES.ABA_CHECK + ' não encontrada.');

  const mapaRef = {};
  carregarTabelaSigtap().forEach(function(t) {
    const k = nomeParaComparar_(t.procedimento);
    if (!mapaRef[k]) mapaRef[k] = t;
  });
  const ofic = carregarTabelaOficial();   // pode ser null se não houver aba SIGTAP_

  const grupos = {};
  if (abaCheck.getLastRow() < 2) return { grupos: [], ofic: ofic };

  abaCheck.getRange(2, 1, abaCheck.getLastRow() - 1, 8).getValues().forEach(function(l) {
    const nome   = String(l[2] || '').replace(/\s+/g, ' ').trim();
    const codigo = codigoSigtap10_(l[4]);
    if (!nome || !codigo) return;

    const k   = nomeParaComparar_(nome);
    const ref = mapaRef[k];
    if (ref && ref.codigo === codigo) return;               // já está certo na Referência

    let tipo;
    if (ofic && !ofic.has(codigo)) tipo = SUGESTOES.TIPO_INEX;
    else                           tipo = ref ? SUGESTOES.TIPO_DIF : SUGESTOES.TIPO_NOVO;

    const chave = k.trim() + '|' + codigo;
    let g = grupos[chave];
    if (!g) {
      g = grupos[chave] = { nome: nome.toUpperCase(), codigo: codigo, tipo: tipo,
                            refCodigo: ref ? ref.codigo : '', cids: [], vezes: 0, data: null };
    }
    g.vezes++;
    const cid = cidLimpo_(l[3]);
    if (cid && g.cids.indexOf(cid) === -1) g.cids.push(cid);
    const d = l[7] instanceof Date ? l[7] : (l[6] instanceof Date ? l[6] : null);
    if (d && (!g.data || d > g.data)) g.data = d;
  });

  return { grupos: Object.keys(grupos).map(function(c) { return grupos[c]; }), ofic: ofic };
}

/**
 * Monta (ou atualiza) a aba SUGESTOES_SIGTAP, mantendo as decisões
 * que já estavam marcadas. Devolve um resumo com as contagens.
 */
function gerarSugestoesSigtap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  // Decisões já marcadas (principalmente os "Ignorar")
  let aba = ss.getSheetByName(SUGESTOES.ABA);
  const decisoes = {};
  if (aba && aba.getLastRow() > 1) {
    aba.getRange(2, 1, aba.getLastRow() - 1, 4).getValues().forEach(function(l) {
      if (l[COL_SUG.DECISAO]) decisoes[chaveSugestao_(l[COL_SUG.NOME], l[COL_SUG.CODIGO])] = l[COL_SUG.DECISAO];
    });
  }

  const r     = montarSugestoes_();
  const nomeOf = function(c) { const x = r.ofic && c ? r.ofic.get(c) : null; return x ? x.nome : ''; };
  const ordem = {};
  ordem[SUGESTOES.TIPO_NOVO] = 0; ordem[SUGESTOES.TIPO_DIF] = 1; ordem[SUGESTOES.TIPO_INEX] = 2;

  const linhas = r.grupos.map(function(g) {
    return [
      decisoes[chaveSugestao_(g.nome, g.codigo)] || '',
      g.tipo, g.nome, g.codigo, nomeOf(g.codigo),
      g.refCodigo, nomeOf(g.refCodigo), g.cids.join(','), g.vezes,
      g.data ? Utilities.formatDate(g.data, tz, 'dd/MM/yyyy') : ''
    ];
  });

  // Ordem: pendentes antes dos ignorados; depois novo → diferente → inexistente; mais vezes primeiro
  linhas.sort(function(a, b) {
    const ia = a[COL_SUG.DECISAO] === SUGESTOES.IGNORAR ? 1 : 0;
    const ib = b[COL_SUG.DECISAO] === SUGESTOES.IGNORAR ? 1 : 0;
    if (ia !== ib) return ia - ib;
    if (ordem[a[COL_SUG.TIPO]] !== ordem[b[COL_SUG.TIPO]]) return ordem[a[COL_SUG.TIPO]] - ordem[b[COL_SUG.TIPO]];
    return b[COL_SUG.VEZES] - a[COL_SUG.VEZES];
  });

  escreverAbaSugestoes_(linhas);

  const conta = function(tipo) {
    return linhas.filter(function(l) { return l[COL_SUG.TIPO] === tipo && l[COL_SUG.DECISAO] !== SUGESTOES.IGNORAR; }).length;
  };
  const resumo = {
    novos:      conta(SUGESTOES.TIPO_NOVO),
    diferentes: conta(SUGESTOES.TIPO_DIF),
    inexist:    conta(SUGESTOES.TIPO_INEX),
    ignorados:  linhas.filter(function(l) { return l[COL_SUG.DECISAO] === SUGESTOES.IGNORAR; }).length
  };
  resumo.pendentes = resumo.novos + resumo.diferentes;

  PropertiesService.getScriptProperties().setProperty('SUGESTOES_PENDENTES', String(resumo.pendentes));
  return resumo;
}

/** Recria o conteúdo da aba SUGESTOES_SIGTAP (cabeçalho, lista de decisão, cores). */
function escreverAbaSugestoes_(linhas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba  = ss.getSheetByName(SUGESTOES.ABA);
  if (!aba) {
    const ativa = ss.getActiveSheet();
    aba = ss.insertSheet(SUGESTOES.ABA);
    aba.setTabColor('#F39A3D');
    if (ativa) ss.setActiveSheet(ativa);
  }
  aba.clear();
  aba.getRange(1, 1, 1, SUGESTOES.CABECALHO.length).setValues([SUGESTOES.CABECALHO])
     .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E79').setWrap(true);
  aba.setFrozenRows(1);

  if (linhas.length) {
    const total = linhas.length;
    aba.getRange(2, COL_SUG.CODIGO + 1, total, 1).setNumberFormat('@');
    aba.getRange(2, COL_SUG.REF + 1,    total, 1).setNumberFormat('@');
    aba.getRange(2, 1, total, SUGESTOES.CABECALHO.length).setValues(linhas).setVerticalAlignment('middle');

    aba.getRange(2, 1, total, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList([SUGESTOES.APROVAR, SUGESTOES.IGNORAR], true)
        .setAllowInvalid(false).build()
    );

    const fundos = linhas.map(function(l) {
      const cor = l[COL_SUG.DECISAO] === SUGESTOES.IGNORAR ? '#EEEEEE'
                : l[COL_SUG.TIPO] === SUGESTOES.TIPO_INEX  ? '#FCE4E4'
                : l[COL_SUG.TIPO] === SUGESTOES.TIPO_DIF   ? '#FFF6E0' : '#E6F4EC';
      return SUGESTOES.CABECALHO.map(function() { return cor; });
    });
    aba.getRange(2, 1, total, SUGESTOES.CABECALHO.length).setBackgrounds(fundos);
  }

  [150, 210, 300, 110, 300, 120, 300, 160, 60, 95].forEach(function(w, i) { aba.setColumnWidth(i + 1, w); });
}

// ============================================================
// APLICAR AS APROVADAS
// ============================================================

/** Leva as linhas "✅ Aprovar" para a Referencia_SIGTAP. Devolve { feitos, recusados }. */
function aplicarSugestoesAprovadas_() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const tz  = ss.getSpreadsheetTimeZone();
  const aba = ss.getSheetByName(SUGESTOES.ABA);
  if (!aba || aba.getLastRow() < 2) throw new Error('Não há sugestões. Use antes 💡 Sugestões para a Referência SIGTAP.');

  const abaRef = ss.getSheetByName(CONFIG_SIGTAP.ABA_REFERENCIA);
  if (!abaRef) throw new Error('Aba ' + CONFIG_SIGTAP.ABA_REFERENCIA + ' não encontrada.');

  const ofic    = carregarTabelaOficial();
  const nomeOf  = function(c) { const x = ofic && c ? ofic.get(c) : null; return x ? x.nome : ''; };
  const hoje    = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  const nCols   = Math.min(7, abaRef.getMaxColumns());

  const refVals = abaRef.getDataRange().getValues();
  const idx = {};                                   // nome normalizado → nº da linha na planilha
  refVals.forEach(function(l, i) {
    if (i === 0 || !String(l[0] || '').trim()) return;
    const k = nomeParaComparar_(l[0]);
    if (!idx[k]) idx[k] = i + 1;
  });

  const linhas    = aba.getRange(2, 1, aba.getLastRow() - 1, SUGESTOES.CABECALHO.length).getValues();
  const restantes = [], novas = [], feitos = [], recusados = [], jaCertos = [];
  const nomesNovos = {};

  linhas.forEach(function(l) {
    if (l[COL_SUG.DECISAO] !== SUGESTOES.APROVAR) { restantes.push(l); return; }

    const nome   = String(l[COL_SUG.NOME]).trim();
    const codigo = codigoSigtap10_(l[COL_SUG.CODIGO]);
    const k      = nomeParaComparar_(nome);
    const recusar = function(motivo) { recusados.push(nome + ': ' + motivo); l[COL_SUG.DECISAO] = ''; restantes.push(l); };

    if (ofic && !ofic.has(codigo)) { recusar('o código ' + codigo + ' não existe no SIGTAP'); return; }

    // Procedimento novo
    if (!idx[k]) {
      if (nomesNovos[k]) { recusar('já foi aprovado outro código para este nome agora'); return; }
      nomesNovos[k] = true;
      novas.push([nome, codigo, String(l[COL_SUG.CIDS] || ''), '📋 Histórico', nomeOf(codigo),
                  'Incluído pelas sugestões do CHECK_LIST em ' + hoje, Number(l[COL_SUG.VEZES]) || 1].slice(0, nCols));
      feitos.push('🆕 ' + nome + ' → ' + codigo);
      return;
    }

    // Código diferente
    const linhaRef = idx[k];
    const atual    = refVals[linhaRef - 1];

    // Aprovado com o MESMO código que já está na Referência: nada a fazer
    if (codigoSigtap10_(atual[1]) === codigo) {
      jaCertos.push(nome);
      return;
    }

    if (/CONFIRMADO/i.test(String(atual[3] || ''))) {
      recusar('está ✅ Confirmado na Referência (linha ' + linhaRef + '). Se o regulador mudou mesmo, troque lá à mão');
      return;
    }

    const antigo = codigoSigtap10_(atual[1]);
    abaRef.getRange(linhaRef, 2).setNumberFormat('@').setValue(codigo);

    const cids = String(atual[2] || '').split(',').map(function(c) { return c.trim(); }).filter(String);
    String(l[COL_SUG.CIDS] || '').split(',').forEach(function(c) {
      c = c.trim(); if (c && cids.indexOf(c) === -1) cids.push(c);
    });
    abaRef.getRange(linhaRef, 3).setValue(cids.join(','));
    if (nCols >= 5) abaRef.getRange(linhaRef, 5).setValue(nomeOf(codigo));
    if (nCols >= 6) abaRef.getRange(linhaRef, 6).setValue('Código trocado de ' + antigo + ' para ' + codigo + ' pelas sugestões em ' + hoje);
    feitos.push('🔄 ' + nome + ': ' + antigo + ' → ' + codigo);
  });

  // Procedimentos novos no fim da Referência
  if (novas.length) {
    const inicio = abaRef.getLastRow() + 1;
    abaRef.getRange(inicio, 2, novas.length, 1).setNumberFormat('@');
    abaRef.getRange(inicio, 1, novas.length, nCols).setValues(novas);
  }

  // A aba de sugestões fica só com o que não foi aplicado
  escreverAbaSugestoes_(restantes);
  const pend = restantes.filter(function(l) {
    return l[COL_SUG.DECISAO] !== SUGESTOES.IGNORAR && l[COL_SUG.TIPO] !== SUGESTOES.TIPO_INEX;
  }).length;
  PropertiesService.getScriptProperties().setProperty('SUGESTOES_PENDENTES', String(pend));

  return { feitos: feitos, recusados: recusados, jaCertos: jaCertos };
}

// ============================================================
// MENU E ROTINA MENSAL
// ============================================================

// Menu 🔧 Ferramentas ▸ 💡 Sugestões para a Referência SIGTAP
function abrirSugestoesSigtap() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = gerarSugestoesSigtap_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setActiveSheet(ss.getSheetByName(SUGESTOES.ABA));
    ui.alert(
      '💡 Sugestões para a Referência SIGTAP',
      '🆕 Procedimentos novos: ' + r.novos + '\n' +
      '🔄 Código diferente: ' + r.diferentes + '\n' +
      '⚠️ Código que não existe no SIGTAP: ' + r.inexist + ' (erro de digitação no CHECK_LIST)\n' +
      (r.ignorados ? '❌ Já ignorados antes: ' + r.ignorados + '\n' : '') +
      '\nNa coluna Decisão, escolha ✅ Aprovar ou ❌ Ignorar.\n' +
      'Depois use 🔧 Ferramentas ▸ ✅ Aplicar sugestões aprovadas.\n\n' +
      'Dica: em "🔄 Código diferente", compare os dois nomes oficiais. ' +
      'Se o nome oficial do código lançado não tiver nada a ver com o procedimento, é erro de digitação: ignore.',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('❌ ' + err.message);
  }
}

// Menu 🔧 Ferramentas ▸ ✅ Aplicar sugestões aprovadas
function aplicarSugestoesSigtap() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = aplicarSugestoesAprovadas_();
    if (!r.feitos.length && !r.recusados.length && !r.jaCertos.length) {
      ui.alert('Nenhuma linha marcada como ✅ Aprovar na aba ' + SUGESTOES.ABA + '.');
      return;
    }
    ui.alert(
      '✅ Sugestões aplicadas',
      (r.feitos.length ? 'Levadas para a Referencia_SIGTAP (' + r.feitos.length + '):\n• ' + r.feitos.join('\n• ') + '\n\n' : '') +
      (r.recusados.length ? 'Não aplicadas (' + r.recusados.length + '):\n• ' + r.recusados.join('\n• ') + '\n\n' : '') +
      (r.jaCertos.length ? 'Já estavam certas na Referência, retiradas da lista (' + r.jaCertos.length + '): ' + r.jaCertos.join(', ') : ''),
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('❌ ' + err.message);
  }
}

/** Chamada pelo gatilho mensal (dia 1º). Sem janelas: só atualiza a aba e o aviso. */
function sugestoesSigtapMensal() {
  try {
    gerarSugestoesSigtap_();
  } catch (err) {
    Logger.log('Sugestões SIGTAP mensais: ' + err);
  }
}
