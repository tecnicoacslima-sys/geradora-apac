/**
 * ============================================================
 * Modelo.gs — PREPARAR O MODELO PARA DISTRIBUIR
 * ============================================================
 * Rode SÓ na planilha modelo, antes de compartilhar com outra
 * unidade. NÃO fica no menu, de propósito: roda pelo editor.
 *
 * COMO USAR:
 *   1. Abra a planilha MODELO (o nome precisa ter "MODELO").
 *   2. Extensões ▸ Apps Script ▸ escolha a função
 *      prepararModeloParaDistribuir ▸ Executar.
 *   3. Volte para a aba da planilha: aparece uma pergunta de
 *      confirmação. Responda "Sim".
 *
 * O QUE FAZ:
 *   • apaga TODAS as propriedades do script (código de acesso à IA,
 *     chave própria, IDs das pastas, nome da unidade...);
 *   • apaga os gatilhos deste projeto;
 *   • esvazia PACIENTES, CHECK_LIST, CONFERENCIAS, USO_IA,
 *     PROTOCOLO_APAC (da linha 9 para baixo) e o formulário da DADOS;
 *   • apaga abas de teste e de sugestões (TESTE_PDF_DIRETO, DADOS_ANTIGA, SUGESTOES_SIGTAP);
 *   • volta os cabeçalhos para "Unidade de Saúde".
 *
 * NÃO MEXE em: MEDICO_SUS, CNES_UBS, Referencia_SIGTAP,
 * SIGTAP_AAAAMM, APAC (formulário do PDF) nem no código.
 * ============================================================
 */

const MODELO = {
  PALAVRA_OBRIGATORIA: 'MODELO',                    // trava: o nome da planilha precisa ter isto
  ABAS_ESVAZIAR: [                                  // [nome da aba, primeira linha a apagar]
    ['PACIENTES',      2],
    ['CHECK_LIST',     2],
    ['CONFERENCIAS',   2],
    ['USO_IA',         2],
    ['PROTOCOLO_APAC', 9]
  ],
  ABAS_APAGAR: ['TESTE_PDF_DIRETO', 'DADOS_ANTIGA', 'SUGESTOES_SIGTAP']
};

function prepararModeloParaDistribuir() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const nome = ss.getName();

  // Trava 1: só roda em planilha com "MODELO" no nome
  if (nome.toUpperCase().indexOf(MODELO.PALAVRA_OBRIGATORIA) === -1) {
    throw new Error(
      'Recusado: o nome desta planilha ("' + nome + '") não tem a palavra ' + MODELO.PALAVRA_OBRIGATORIA + '. ' +
      'Esta função apaga a configuração e os dados de pacientes, então só roda na planilha modelo. Nada foi alterado.'
    );
  }

  // Trava 2: confirmação na tela da planilha
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.alert(
    '🧹 Preparar modelo para distribuir',
    'Planilha: ' + nome + '\n\n' +
    'Isto vai APAGAR:\n' +
    '• as propriedades do script (código de acesso à IA, pastas, nome da unidade)\n' +
    '• os gatilhos\n' +
    '• os dados de PACIENTES, CHECK_LIST, CONFERENCIAS, USO_IA e PROTOCOLO_APAC\n' +
    '• o formulário da aba DADOS\n\n' +
    'MEDICO_SUS, CNES_UBS, Referencia_SIGTAP e a tabela SIGTAP ficam como estão.\n\n' +
    'Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resposta !== ui.Button.YES) {
    Logger.log('Cancelado. Nada foi alterado.');
    return;
  }

  const relatorio = [];

  // 1. Propriedades do script
  const props  = PropertiesService.getScriptProperties();
  const chaves = Object.keys(props.getProperties());
  props.deleteAllProperties();
  relatorio.push('🔑 ' + chaves.length + ' propriedade(s) apagada(s)' + (chaves.length ? ': ' + chaves.join(', ') : ''));

  // 2. Gatilhos
  const gatilhos = ScriptApp.getProjectTriggers();
  gatilhos.forEach(function(g) { ScriptApp.deleteTrigger(g); });
  relatorio.push('⏰ ' + gatilhos.length + ' gatilho(s) apagado(s)');

  // 3. Abas de dados
  MODELO.ABAS_ESVAZIAR.forEach(function(par) {
    const aba = ss.getSheetByName(par[0]);
    if (!aba) return;
    const linhas = aba.getMaxRows() - par[1] + 1;
    if (linhas > 0) aba.getRange(par[1], 1, linhas, aba.getMaxColumns()).clearContent();
    relatorio.push('🧽 ' + par[0] + ' esvaziada (da linha ' + par[1] + ' para baixo)');
  });

  // 4. Formulário da DADOS
  try {
    limparFormulario_();
    relatorio.push('🧽 Formulário da aba DADOS limpo');
  } catch (err) {
    relatorio.push('⚠️ Não consegui limpar o formulário da DADOS: ' + err.message);
  }

  // 5. Abas de teste
  MODELO.ABAS_APAGAR.forEach(function(n) {
    const aba = ss.getSheetByName(n);
    if (aba) { ss.deleteSheet(aba); relatorio.push('🗑️ Aba ' + n + ' apagada'); }
  });

  // 6. Cabeçalhos genéricos (sem propriedades, fica "Unidade de Saúde")
  atualizarCabecalhos_();
  relatorio.push('🏷️ Cabeçalhos das abas DADOS e PROTOCOLO_APAC com "Unidade de Saúde"');

  SpreadsheetApp.flush();
  Logger.log(relatorio.join('\n'));

  ui.alert(
    '✅ Modelo pronto',
    relatorio.join('\n') + '\n\n' +
    'Agora é só compartilhar. Na cópia, a unidade usa:\n' +
    '📋 APAC ▸ 🚀 Primeira configuração (comece aqui).',
    ui.ButtonSet.OK
  );
}
