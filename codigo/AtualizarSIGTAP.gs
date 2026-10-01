/**
 * ============================================================
 * AtualizarSIGTAP.gs — MONTA A ABA SIGTAP_AAAAMM A PARTIR DO .ZIP
 * ============================================================
 * Menu 📋 APAC ▸ 🔧 Ferramentas ▸ 📥 Atualizar tabela SIGTAP (.zip)
 *
 * Você escolhe o arquivo "TabelaUnificada_AAAAMM_v....zip" baixado
 * do SIGTAP (sem descompactar). O script:
 *   1. abre o .zip e lê tb_procedimento.txt e rl_procedimento_habilitacao.txt
 *      (posições das colunas tiradas dos arquivos *_layout.txt do próprio pacote);
 *   2. cria a aba SIGTAP_AAAAMM (código, nome, habilitações, competência);
 *   3. compara com a tabela anterior (novos, que saíram);
 *   4. confere se todos os códigos da Referencia_SIGTAP ainda existem;
 *   5. se você marcar, apaga a(s) tabela(s) antiga(s).
 *
 * O Code_SIGTAP.gs usa sozinho a aba SIGTAP_ mais recente, então
 * não é preciso mexer em código nenhum a cada mês.
 * Não usa IA: não gasta tokens.
 * ============================================================
 */

const SIGTAP_IMPORT = {
  PREFIXO:     'SIGTAP_',
  MIN_PROCS:   3000,       // abaixo disso, o arquivo provavelmente está incompleto
  ARQ_PROC:    'tb_procedimento.txt',
  LAY_PROC:    'tb_procedimento_layout.txt',
  ARQ_HAB:     'rl_procedimento_habilitacao.txt',
  LAY_HAB:     'rl_procedimento_habilitacao_layout.txt',
  // Posições de reserva (usadas só se o pacote vier sem o arquivo de layout)
  PADRAO_PROC: { CO_PROCEDIMENTO: [1, 10], NO_PROCEDIMENTO: [11, 260], DT_COMPETENCIA: [331, 336] },
  PADRAO_HAB:  { CO_PROCEDIMENTO: [1, 10], CO_HABILITACAO: [11, 14] }
};

// ============================================================
// JANELA
// ============================================================

function abrirAtualizarSigtap() {
  const html = HtmlService.createHtmlOutputFromFile('DialogoAtualizarSigtap')
    .setWidth(480)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '📥 Atualizar tabela SIGTAP');
}

/** Qual tabela está em uso agora (mostrado no topo da janela). */
function infoTabelaSigtapAtual() {
  const nome = abaSigtapMaisRecente_();
  if (!nome) return { aba: '', linhas: 0 };
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  return { aba: nome, linhas: Math.max(aba.getLastRow() - 1, 0) };
}

// ============================================================
// LEITURA DO PACOTE (funções "puras": só texto entra e sai)
// ============================================================

/**
 * Lê um arquivo *_layout.txt ("Coluna,Tamanho,Inicio,Fim,Tipo")
 * e devolve { NOME_COLUNA: [inicio, fim] }. Sem layout, usa o padrão.
 */
function lerLayoutSigtap_(textoLayout, padrao) {
  const mapa = {};
  String(textoLayout || '').split(/\r?\n/).slice(1).forEach(function(l) {
    const p = l.split(',');
    if (p.length >= 4 && /^\d+$/.test(p[2].trim()) && /^\d+$/.test(p[3].trim())) {
      mapa[p[0].trim().toUpperCase()] = [parseInt(p[2], 10), parseInt(p[3], 10)];
    }
  });
  Object.keys(padrao).forEach(function(k) { if (!mapa[k]) mapa[k] = padrao[k]; });
  return mapa;
}

function cortar_(linha, pos) {
  return linha.substring(pos[0] - 1, pos[1]).trim();
}

/**
 * Monta as linhas da aba a partir dos textos do pacote.
 * Devolve { competencia, linhas: [[codigo, nome, habilitacoes, competencia], ...] }
 */
function montarLinhasSigtap_(txtProc, txtLayProc, txtHab, txtLayHab) {
  const lp = lerLayoutSigtap_(txtLayProc, SIGTAP_IMPORT.PADRAO_PROC);
  const lh = lerLayoutSigtap_(txtLayHab,  SIGTAP_IMPORT.PADRAO_HAB);

  // Habilitações por procedimento
  const hab = {};
  String(txtHab || '').split(/\r?\n/).forEach(function(l) {
    if (!l.trim()) return;
    const co = cortar_(l, lh.CO_PROCEDIMENTO);
    const h  = cortar_(l, lh.CO_HABILITACAO);
    if (!co || !h) return;
    if (!hab[co]) hab[co] = [];
    if (hab[co].indexOf(h) === -1) hab[co].push(h);
  });

  // Procedimentos
  const competencias = {};
  const vistos = {};
  const linhas = [];
  String(txtProc || '').split(/\r?\n/).forEach(function(l) {
    if (!l.trim()) return;
    const co   = cortar_(l, lp.CO_PROCEDIMENTO);
    const nome = cortar_(l, lp.NO_PROCEDIMENTO);
    const comp = cortar_(l, lp.DT_COMPETENCIA);
    if (!/^\d{10}$/.test(co) || !nome || vistos[co]) return;
    vistos[co] = true;
    if (/^\d{6}$/.test(comp)) competencias[comp] = true;
    linhas.push([co, nome, (hab[co] || []).sort().join(', '), comp]);
  });

  linhas.sort(function(a, b) { return a[0] < b[0] ? -1 : (a[0] > b[0] ? 1 : 0); });

  const lista = Object.keys(competencias);
  if (lista.length !== 1) {
    throw new Error('O arquivo traz ' + (lista.length ? 'mais de uma competência (' + lista.join(', ') + ')' : 'nenhuma competência') +
                    '. Confira se é a Tabela Unificada do SIGTAP.');
  }
  return { competencia: lista[0], linhas: linhas };
}

// ============================================================
// IMPORTAÇÃO (chamada pela janela)
// ============================================================

/**
 * base64: o .zip escolhido no computador.
 * apagarAntigas: true → apaga as outras abas SIGTAP_AAAAMM depois de criar a nova.
 */
function importarTabelaSigtap(base64, nomeArquivo, apagarAntigas) {
  // 1. Abre o .zip
  let arquivos;
  try {
    const zip = Utilities.newBlob(Utilities.base64Decode(base64), 'application/zip', nomeArquivo || 'sigtap.zip');
    arquivos = Utilities.unzip(zip);
  } catch (err) {
    throw new Error('Não consegui abrir o arquivo. Ele precisa ser o .zip da Tabela Unificada do SIGTAP ' +
                    '(TabelaUnificada_AAAAMM_v....zip), sem descompactar.');
  }

  const achar = function(nome) {
    return arquivos.filter(function(b) {
      return String(b.getName()).split('/').pop().toLowerCase() === nome;
    })[0] || null;
  };
  const texto = function(blob) { return blob ? blob.getDataAsString('ISO-8859-1') : ''; };

  const bProc = achar(SIGTAP_IMPORT.ARQ_PROC);
  if (!bProc) {
    throw new Error('O .zip não tem o arquivo ' + SIGTAP_IMPORT.ARQ_PROC +
                    '. Confira se é a Tabela Unificada do SIGTAP.');
  }

  // 2. Monta as linhas
  const r = montarLinhasSigtap_(
    texto(bProc), texto(achar(SIGTAP_IMPORT.LAY_PROC)),
    texto(achar(SIGTAP_IMPORT.ARQ_HAB)), texto(achar(SIGTAP_IMPORT.LAY_HAB))
  );
  if (r.linhas.length < SIGTAP_IMPORT.MIN_PROCS) {
    throw new Error('O arquivo trouxe só ' + r.linhas.length + ' procedimentos (o normal é mais de 5.000). ' +
                    'Ele pode estar incompleto. Baixe de novo.');
  }

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const nomeNova  = SIGTAP_IMPORT.PREFIXO + r.competencia;
  const anterior  = abaSigtapMaisRecente_();
  const ativa     = ss.getActiveSheet();

  // 3. Tabela anterior, para comparar (lida antes de mexer em qualquer aba)
  let antes = null;
  if (anterior && anterior !== nomeNova) {
    antes = {};
    const abaAnt = ss.getSheetByName(anterior);
    if (abaAnt.getLastRow() > 1) {
      abaAnt.getRange(2, 1, abaAnt.getLastRow() - 1, 1).getValues().forEach(function(l) {
        antes[codigoSigtap10_(l[0])] = true;
      });
    }
  }

  // 4. Cria a aba nova (se já existir com o mesmo nome, é substituída)
  const jaExistia = ss.getSheetByName(nomeNova);
  const posicao   = anterior ? ss.getSheetByName(anterior).getIndex() : ss.getNumSheets();
  if (jaExistia) ss.deleteSheet(jaExistia);
  const aba = ss.insertSheet(nomeNova, Math.min(posicao, ss.getNumSheets()));

  const total = r.linhas.length;
  aba.getRange(1, 1, total + 1, 4).setNumberFormat('@');          // tudo texto: o zero da frente não some
  aba.getRange(1, 1, 1, 4)
     .setValues([['CO_PROCEDIMENTO', 'NO_PROCEDIMENTO', 'HABILITACOES', 'COMPETENCIA']])
     .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F4E79');
  aba.getRange(2, 1, total, 4).setValues(r.linhas);
  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 110);
  aba.setColumnWidth(2, 620);
  aba.setColumnWidth(3, 260);
  aba.setColumnWidth(4, 90);
  const sobra = aba.getMaxRows() - (total + 1);
  if (sobra > 0) aba.deleteRows(total + 2, sobra);
  if (aba.getMaxColumns() > 4) aba.deleteColumns(5, aba.getMaxColumns() - 4);

  // 5. Comparações
  const novosCodigos = {};
  r.linhas.forEach(function(l) { novosCodigos[l[0]] = l[1]; });

  let novos = [], sairam = [];
  if (antes) {
    novos  = r.linhas.filter(function(l) { return !antes[l[0]]; }).map(function(l) { return l[0] + ' ' + l[1]; });
    sairam = Object.keys(antes).filter(function(c) { return c && !novosCodigos[c]; });
  }

  const refForaDaTabela = [];
  const abaRef = ss.getSheetByName(CONFIG_SIGTAP.ABA_REFERENCIA);
  if (abaRef && abaRef.getLastRow() > 1) {
    abaRef.getRange(2, 1, abaRef.getLastRow() - 1, 2).getValues().forEach(function(l, i) {
      const cod = codigoSigtap10_(l[1]);
      if (cod && !novosCodigos[cod]) {
        refForaDaTabela.push('linha ' + (i + 2) + ' · ' + String(l[0]).trim() + ' · ' + cod);
      }
    });
  }

  // 6. Apaga as antigas, se pedido
  const apagadas = [];
  if (apagarAntigas) {
    ss.getSheets().forEach(function(s) {
      const n = s.getName();
      if (n !== nomeNova && /^SIGTAP_\d{6}$/.test(n)) { ss.deleteSheet(s); apagadas.push(n); }
    });
  }

  if (ativa && ativa.getName() !== nomeNova && ss.getSheetByName(ativa.getName())) ss.setActiveSheet(ativa);
  SpreadsheetApp.flush();

  return {
    aba:           nomeNova,
    competencia:   r.competencia.substring(4, 6) + '/' + r.competencia.substring(0, 4),
    total:         total,
    substituida:   !!jaExistia,
    anterior:      anterior && anterior !== nomeNova ? anterior : '',
    novos:         novos.length,
    exemplosNovos: novos.slice(0, 8),
    sairam:        sairam,
    refFora:       refForaDaTabela,
    apagadas:      apagadas
  };
}
