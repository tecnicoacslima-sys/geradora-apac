/**
 * ============================================================
 * Código.gs — GERADORA APAC EXTERNA (v2.5 · modelo universal)
 * ============================================================
 * Todos os campos da aba DADOS são acessados pelo NOME
 * (ex.: campo_cns_paciente), definidos em Layout_DADOS.gs.
 * IDs de pastas, nome da unidade e chave da IA ficam nas
 * propriedades do script (ver Config.gs). Nada fixo aqui.
 *
 * v2.5 — 🔄 Atualizar sistema: baixa o código oficial do GitHub
 *        (ver Atualizador.gs); aviso ao abrir quando há versão nova;
 *        o painel confere os gatilhos sozinho a cada abertura.
 *
 * v2.4 — Sugestões para a Referencia_SIGTAP a partir do CHECK_LIST
 *        (ver SugestoesSIGTAP.gs): 2 itens no menu, gatilho do dia 1º
 *        e aviso ao abrir quando há sugestões pendentes.
 *
 * v2.3 — Menu mostra "🚀 Primeira configuração" enquanto a planilha
 *        não estiver configurada (ver Assistente.gs).
 *
 * v2.2 — Cartão SUS e CPF conferidos pelo dígito verificador
 *        (funções em Validacao.gs):
 *   • ao digitar o CNS do paciente e o documento do médico
 *   • ao ler o Espelho CELK (e escolhe o CNS certo se houver
 *     mais de um número parecido no PDF)
 *   • antes de gerar a APAC em PDF
 *   • ao salvar no Protocolo (médico com CNS errado não é cadastrado)
 * ============================================================
 */

function pedirPermissoes() {
  const doc = DocumentApp.create("__teste__");
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
}

// ── MENU ÚNICO, na ordem do trabalho ────────────────────────
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Planilha ainda não configurada (cópia nova)? Mostra o assistente em 1º lugar.
  let configurado = true;
  try { configurado = sistemaConfigurado_(); } catch (err) { configurado = false; }

  const menu = ui.createMenu('📋 APAC');
  if (!configurado) {
    menu.addItem('🚀 Primeira configuração (comece aqui)', 'abrirAssistente').addSeparator();
  }

  menu
    .addItem('🧭 Abrir painel',                'abrirPainelAPAC')
    .addSeparator()
    .addItem('①  📄 Selecionar Espelho CELK', 'abrirPainelSelecaoPDF')
    .addItem('②  🤖 Buscar dados da APAC',    'abrirDialogoSigtap')
    .addItem('③  🖨️ Gerar APAC em PDF',       'gerarPdfApacs')
    .addItem('④  💾 Salvar no Protocolo',     'salvarProcessoCompleto')
    .addItem('⑤  🗑️ Limpar formulário',       'limparFormularioAPAC')
    .addSeparator()
    .addItem('🧾 Conferir espelhos (SUS · CELK · docs)', 'abrirDialogoConferencia')
    .addItem('🔍 Buscar procedimento SIGTAP', 'abrirBuscaSIGTAPJanela')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Ferramentas')
      .addItem('🚀 Assistente de configuração',          'abrirAssistente')
      .addItem('🔄 Atualizar sistema',                    'abrirAtualizador')
      .addItem('✏️ Alterar dados da unidade',            'alterarDadosUnidade')
      .addItem('📊 Consumo da IA (tokens)',              'mostrarConsumoIA')
      .addItem('🔑 Configurar acesso à IA',              'configurarAcessoIA')
      .addItem('🔌 Testar acesso à IA',                  'testarAcessoIA')
      .addSeparator()
      .addItem('📥 Atualizar tabela SIGTAP (.zip)',      'abrirAtualizarSigtap')
      .addItem('💡 Sugestões para a Referência SIGTAP',  'abrirSugestoesSigtap')
      .addItem('✅ Aplicar sugestões aprovadas',          'aplicarSugestoesSigtap')
      .addItem('🪪 Conferir Cartões SUS cadastrados',    'conferirCartoesSUS')
      .addItem('🔢 Numerar médicos (executar 1x)',       'numerarMedicos')
      .addItem('🔁 Instalar gatilhos (buscas + painel)', 'instalarGatilhoSeNecessario')
      .addItem('🧪 Teste Extrator PDF Direto',           'abrirTesteUpload'))
    .addToUi();

  // Avisos: versão nova do sistema e sugestões pendentes para a Referencia_SIGTAP
  if (configurado) {
    try {
      const props  = PropertiesService.getScriptProperties();
      const avisos = [];
      const nova   = props.getProperty('ATUALIZACAO_DISPONIVEL');
      if (nova) avisos.push('🔄 Versão ' + nova + ' disponível: 🔧 Ferramentas ▸ Atualizar sistema.');
      const pend = parseInt(props.getProperty('SUGESTOES_PENDENTES') || '0', 10);
      if (pend > 0) avisos.push('💡 ' + pend + ' sugestão(ões) para a Referência SIGTAP (aba SUGESTOES_SIGTAP).');
      if (avisos.length) {
        SpreadsheetApp.getActiveSpreadsheet().toast(avisos.join('\n'), '📋 Geradora de APAC', 12);
      }
    } catch (err) { /* sem aviso */ }
  }

  if (!configurado) {
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Esta planilha ainda não foi configurada. Clique em 📋 APAC ▸ 🚀 Primeira configuração.',
        '👋 Bem-vindo à Geradora de APAC', 20
      );
    } catch (err) { /* sem aviso, o item do menu já basta */ }
  }
}
// ────────────────────────────────────────────────────────────

// ============================================================
// 🧭 PAINEL LATERAL
// ============================================================

/**
 * Abre o painel lateral com os botões ① a ⑤.
 * Também é chamado automaticamente ao abrir a planilha
 * (gatilho instalado por "Instalar gatilhos").
 */
function abrirPainelAPAC(e) {
  // Confere os gatilhos a cada abertura (instala os que vierem numa atualização nova)
  try { instalarGatilhos_(); } catch (err) { Logger.log('Gatilhos: ' + err); }
  abrirPainelNoModo_('inicio');
}

function abrirPainelNoModo_(modo) {
  const t = HtmlService.createTemplateFromFile('PainelAPAC');
  t.modo    = modo || 'inicio';
  t.unidade = subtituloUnidade_(' · ');
  const html = t.evaluate().setTitle('📋 APAC Externa');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Resumo do paciente mostrado no topo do painel
function obterResumoPainel() {
  try {
    return {
      ok:   true,
      nome: textoCampo_('nome_paciente'),
      cns:  textoCampo_('cns_paciente'),
      proc: textoCampo_('proc_nome'),
      data: textoCampo_('esf_data')
    };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

function gerarPdfApacPainel() {
  try {
    const r = gerarPdfApac_();
    return { ok: true, nome: r.nomeArquivo, url: r.url };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

function salvarNoProtocoloPainel() {
  try {
    return { ok: true, msg: salvarNoProtocolo_() };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

function limparFormularioPainel() {
  limparFormulario_();
  return { ok: true };
}

// ---- Buscar paciente já cadastrado na aba PACIENTES ----

function normalizarBusca_(t) {
  return String(t || '').toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ').trim();
}

/**
 * Procura na aba PACIENTES por nome (todas as palavras digitadas),
 * CNS ou CPF. Devolve no máximo 15 resultados.
 */
function buscarPacientesPainel(termo) {
  const texto = normalizarBusca_(termo);
  const digitos = String(termo || '').replace(/\D/g, '');
  if (texto.length < 3) return { ok: true, lista: [], total: 0 };

  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PACIENTES');
  if (!aba) return { ok: false, msg: 'Aba PACIENTES não encontrada.' };

  const tz     = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const dados  = aba.getDataRange().getValues();
  const buscaPorNumero = digitos.length >= 5 && digitos.length === texto.replace(/[\s.\-]/g, '').length;
  const palavras = texto.split(' ');
  const lista  = [];
  let total    = 0;

  for (let i = 1; i < dados.length; i++) {
    const p    = dados[i];
    const nome = normalizarBusca_(p[1]);
    if (!nome) continue;

    let achou;
    if (buscaPorNumero) {
      const cns = String(p[0]).replace(/\D/g, '');
      const cpf = String(p[9]).replace(/\D/g, '');
      achou = cns.indexOf(digitos) !== -1 || cpf.indexOf(digitos) !== -1;
    } else {
      achou = palavras.every(function(w) { return nome.indexOf(w) !== -1; });
    }
    if (!achou) continue;

    total++;
    if (lista.length < 15) {
      const nasc = p[2] instanceof Date ? Utilities.formatDate(p[2], tz, 'dd/MM/yyyy') : String(p[2] || '');
      lista.push({
        linha: i + 1,
        nome:  String(p[1]).toUpperCase().trim(),
        nasc:  nasc,
        cns:   String(p[0]).replace(/\D/g, '')
      });
    }
  }
  return { ok: true, lista: lista, total: total };
}

/**
 * Carrega no formulário o paciente da linha informada da aba PACIENTES
 * (inclusive o Cartão SUS, que antes você copiava e colava).
 */
function carregarPacientePainel(linha, nomeEsperado) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PACIENTES');
  if (!aba) return { ok: false, msg: 'Aba PACIENTES não encontrada.' };

  const p = aba.getRange(linha, 1, 1, 11).getValues()[0];
  if (normalizarBusca_(p[1]) !== normalizarBusca_(nomeEsperado)) {
    return { ok: false, msg: 'A lista de pacientes mudou. Faça a busca de novo.' };
  }

  gravarCampo_('cns_paciente', String(p[0]).replace(/\D/g, ''));
  preencherPaciente_(p);
  SpreadsheetApp.flush();
  return { ok: true, nome: String(p[1]).toUpperCase().trim() };
}

// Busca SIGTAP em janela solta (não fecha o painel lateral)
function abrirBuscaSIGTAPJanela() {
  const html = HtmlService.createHtmlOutputFromFile('BuscaSIGTAP')
    .setWidth(420)
    .setHeight(620);
  SpreadsheetApp.getUi().showModelessDialog(html, '🔍 Buscar procedimento SIGTAP');
}

// ── NUMERAÇÃO INICIAL DOS MÉDICOS (executar 1 vez pelo menu) ─
function numerarMedicos() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MEDICO_SUS");
  if (!aba) {
    SpreadsheetApp.getUi().alert("❌ Aba MEDICO_SUS não encontrada.");
    return;
  }

  aba.getRange("C1").setValue("Nº")
     .setFontWeight("bold").setHorizontalAlignment("center");

  const ultima = aba.getLastRow();
  let num = 0;
  for (let i = 2; i <= ultima; i++) {
    const nome = String(aba.getRange(i, 1).getValue()).toUpperCase().trim();
    if (!nome) continue;
    if (nome === "SELECIONE") {
      aba.getRange(i, 3).setValue("");
      continue;
    }
    num++;
    aba.getRange(i, 3).setValue(num)
       .setHorizontalAlignment("center").setFontWeight("bold")
       .setBorder(true, true, true, true, true, true);
  }

  SpreadsheetApp.getUi().alert(
    "✅ " + num + " profissionais numerados na coluna C da aba MEDICO_SUS.\n\n" +
    "Agora basta digitar o Nº no campo PROFISSIONAL SOLICITANTE da aba DADOS e apertar Enter."
  );
}

// Retorna o próximo número livre da coluna C de MEDICO_SUS
function obterProximoNumeroMedico(dadosMed) {
  let maior = 0;
  for (let i = 1; i < dadosMed.length; i++) {
    const n = parseInt(dadosMed[i][2], 10);
    if (!isNaN(n) && n > maior) maior = n;
  }
  return maior + 1;
}

// ============================================================
// GATILHO DE EDIÇÃO (instalável) — buscas automáticas
// ============================================================

function onEditAPAC(e) {
  if (!e || !e.range) return;
  const range = e.range;
  if (range.getSheet().getName().toUpperCase() !== "DADOS") return;

  // getValue() em vez de e.value: funciona tanto digitando quanto colando
  const valor = String(range.getValue() || "").trim();

  if      (ehCampo_(range, 'cnes_solicitante')) aoEditarCnes_(valor);
  else if (ehCampo_(range, 'esf_profissional')) aoEditarProfissional_(valor);
  else if (ehCampo_(range, 'esf_doc_numero'))   aoEditarDocProfissional_(valor);
  else if (ehCampo_(range, 'cns_paciente'))     aoEditarCnsPaciente_(e);
  else if (ehCampo_(range, 'cep'))              buscarEnderecoCEP(e);
}

// 1. BUSCA CNES
function aoEditarCnes_(valor) {
  const cnes = valor.replace(/\D/g, "");
  if (!cnes || /^0+$/.test(cnes)) return;

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const dados = ss.getSheetByName("CNES_UBS").getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const cnesLinha = String(dados[i][0]).replace(/\D/g, "").padStart(7, "0");
    if (cnesLinha === cnes.padStart(7, "0")) {
      gravarCampo_('estab_solicitante', dados[i][1]);
      ss.toast(String(dados[i][1]), "🏢 Estabelecimento encontrado", 4);
      return;
    }
  }

  campo_('estab_solicitante').clearContent();
  SpreadsheetApp.getUi().alert(
    "⚠️ UBS não cadastrada.\n\n" +
    "Digite o nome no campo NOME DO ESTABELECIMENTO. " +
    "Ao salvar no Protocolo, o sistema aprende esse CNES."
  );
  campo_('estab_solicitante').activate();
}

// 2. BUSCA MÉDICO — aceita Nº ou NOME
function aoEditarProfissional_(valor) {
  const entrada = valor.toUpperCase();
  if (!entrada || entrada === "SELECIONE") return;

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const dadosMed = ss.getSheetByName("MEDICO_SUS").getDataRange().getValues();

  // 2a. Digitou APENAS NÚMEROS → busca pelo Nº (coluna C)
  if (/^\d+$/.test(entrada)) {
    const numDigitado = parseInt(entrada, 10);

    for (let k = 1; k < dadosMed.length; k++) {
      if (parseInt(dadosMed[k][2], 10) === numDigitado) {
        const nomeMed = String(dadosMed[k][0]).toUpperCase().trim();
        gravarCampo_('esf_profissional', nomeMed);
        gravarCampo_('esf_doc_tipo', 'CNS');
        gravarCampo_('esf_doc_numero', String(dadosMed[k][1]).trim());
        ss.toast("Nº " + numDigitado + " → " + nomeMed, "👨‍⚕️ Profissional encontrado", 4);
        return;
      }
    }

    campo_('esf_profissional').clearContent();
    SpreadsheetApp.getUi().alert(
      "⚠️ O Nº " + numDigitado + " não está cadastrado na aba MEDICO_SUS.\n\n" +
      "Confira o número ou digite o NOME do profissional.\n" +
      "Se for profissional novo, digite o nome e depois o Cartão SUS no campo Nº DO DOCUMENTO."
    );
    campo_('esf_profissional').activate();
    return;
  }

  // 2b. Digitou NOME
  for (let j = 1; j < dadosMed.length; j++) {
    if (String(dadosMed[j][0]).toUpperCase().trim() === entrada) {
      gravarCampo_('esf_doc_tipo', 'CNS');
      gravarCampo_('esf_doc_numero', String(dadosMed[j][1]).trim());
      ss.toast(entrada, "👨‍⚕️ Profissional encontrado", 4);
      return;
    }
  }

  campo_('esf_doc_numero').clearContent();
  SpreadsheetApp.getUi().alert(
    "⚠️ Profissional novo!\n\n" +
    "Digite o Cartão SUS dele no campo Nº DO DOCUMENTO. O cadastro é feito automaticamente."
  );
  campo_('esf_doc_numero').activate();
}

/**
 * Motivo pelo qual o documento do profissional está errado,
 * conforme o tipo escolhido (CNS ou CPF). '' = está certo.
 */
function motivoDocProfissional_(tipo, numero) {
  const dig = soDigitos_(numero);
  if (String(tipo || '').toUpperCase() === 'CPF') {
    return validarCPF(dig) ? '' : 'o CPF ' + dig + ' tem dígito errado';
  }
  const motivo = motivoCNSInvalido_(dig);
  return motivo ? 'o Cartão SUS ' + dig + ' ' + motivo : '';
}

// 3. CADASTRO AUTOMÁTICO DO MÉDICO (só se o documento estiver certo)
function aoEditarDocProfissional_(valor) {
  const doc = soDigitos_(valor);
  if (!doc || /^0+$/.test(doc)) return;

  const motivo = motivoDocProfissional_(textoCampo_('esf_doc_tipo'), doc);
  if (motivo) {
    SpreadsheetApp.getUi().alert(
      "⚠️ Documento do profissional com erro\n\n" +
      "Não confere: " + motivo + ".\n\n" +
      "Confira o número no CNES ou no CADSUS e digite de novo.\n" +
      "Enquanto estiver errado, o profissional NÃO é cadastrado."
    );
    campo_('esf_doc_numero').activate();
    return;
  }

  const nome = textoCampo_('esf_profissional').toUpperCase();
  if (!nome || nome === "SELECIONE") return;

  const novoNum = cadastrarMedicoSeNovo_(nome, doc);
  if (novoNum) {
    SpreadsheetApp.getUi().alert("✅ Profissional " + nome + " cadastrado com sucesso!\nNº atribuído: " + novoNum);
  }
}

// Cadastra na MEDICO_SUS se ainda não existir. Retorna o Nº novo, ou 0 se já existia.
function cadastrarMedicoSeNovo_(nome, sus) {
  const aba   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MEDICO_SUS");
  const dados = aba.getDataRange().getValues();

  const jaExiste = dados.some(function(row) {
    return String(row[0]).toUpperCase().trim() === nome;
  });
  if (jaExiste) return 0;

  const novoNum   = obterProximoNumeroMedico(dados);
  const novaLinha = aba.getLastRow() + 1;
  aba.getRange(novaLinha, 1, 1, 3).setValues([[nome, sus, novoNum]]);
  aba.getRange(novaLinha, 1).setHorizontalAlignment("left");
  aba.getRange(novaLinha, 2, 1, 2).setHorizontalAlignment("center");
  aba.getRange(novaLinha, 1, 1, 3)
     .setBorder(true, true, true, true, true, true)
     .setFontWeight("bold");
  return novoNum;
}

// Cadastra na CNES_UBS se ainda não existir
function cadastrarCnesSeNovo_(cnes, nome) {
  const aba   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CNES_UBS");
  const dados = aba.getRange("A:A").getValues();
  const alvo  = cnes.padStart(7, "0");

  const jaExiste = dados.some(function(row) {
    return String(row[0]).replace(/\D/g, "").padStart(7, "0") === alvo;
  });
  if (jaExiste) return;

  const novaL = aba.getLastRow() + 1;
  aba.getRange(novaL, 1, 1, 2).setValues([[cnes, nome]]);
  aba.getRange(novaL, 1, 1, 2).setBorder(true, true, true, true, true, true).setFontWeight("bold");
  aba.getRange(novaL, 1).setHorizontalAlignment("center");
}

// ============================================================
// ④ SALVAR NO PROTOCOLO
// ============================================================

// Versão do menu
function salvarProcessoCompleto() {
  try {
    const msg = salvarNoProtocolo_();
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, "✅ Processo salvo", 5);
  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ " + err.message);
  }
}

// Faz o trabalho; lança erro se faltar algo
function salvarNoProtocolo_() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const abaCheck     = ss.getSheetByName("CHECK_LIST");
  const abaProtocolo = ss.getSheetByName("PROTOCOLO_APAC");

  if (!textoCampo_('cns_paciente')) {
    throw new Error("Informe o Cartão SUS do paciente antes de salvar.");
  }

  // APRENDIZADO CNES
  const cnesVal = textoCampo_('cnes_solicitante').replace(/\D/g, "");
  const ubsNome = textoCampo_('estab_solicitante').toUpperCase();
  if (cnesVal && !/^0+$/.test(cnesVal) && ubsNome && ubsNome !== "DIGITE O NOME AQUI") {
    cadastrarCnesSeNovo_(cnesVal, ubsNome);
  }

  // APRENDIZADO MÉDICO — só grava se o documento estiver certo
  const medNome = textoCampo_('esf_profissional').toUpperCase();
  const medSus  = soDigitos_(textoCampo_('esf_doc_numero'));
  if (medNome && medNome !== "SELECIONE" && medSus && !/^0+$/.test(medSus) &&
      !motivoDocProfissional_(textoCampo_('esf_doc_tipo'), medSus)) {
    cadastrarMedicoSeNovo_(medNome, medSus);
  }

  // COLETA DOS DADOS
  const nomePac  = textoCampo_('nome_paciente').toUpperCase();
  const sexoTxt  = textoCampo_('sexo').toUpperCase();
  const sexo     = sexoTxt === "MASCULINO" ? "M" : (sexoTxt === "FEMININO" ? "F" : "");

  const procedimento = textoCampo_('proc_nome').replace(/\s+/g, " ");
  const cid          = textoCampo_('cid_principal').toUpperCase();
  const codigo       = "'" + textoCampo_('proc_codigo').replace(/\s+/g, "");

  const medicoNome1  = textoCampo_('ext_profissional').toUpperCase();
  const medicoNome2  = textoCampo_('ext_estabelecimento').toUpperCase();
  const medicoConcat = medicoNome1 && medicoNome2 ? medicoNome1 + " - " + medicoNome2
                     : (medicoNome1 || medicoNome2);

  const dataSolic = formatarData_(campo_('esf_data'));
  const dataReceb = formatarData_(campo_('data_recebimento'));

  const linha = [nomePac, sexo, procedimento, cid, codigo, medicoConcat, dataSolic, dataReceb];

  // INSERE NOVA LINHA NO TOPO E GRAVA
  abaCheck.insertRowBefore(2);
  abaProtocolo.insertRowBefore(9);

  gravarLinhaProtocolo_(abaCheck.getRange("A2:H2"), linha);
  gravarLinhaProtocolo_(abaProtocolo.getRange("A9:H9"), linha);

  return nomePac + " gravado no CHECK_LIST e no PROTOCOLO_APAC.";
}

function gravarLinhaProtocolo_(range, linha) {
  range.setValues([linha]);
  range.setHorizontalAlignments([["left", "center", "left", "center", "center", "left", "center", "center"]]);
  aplicarEstiloGeral(range);
}

function formatarData_(range) {
  const v = range.getValue();
  if (v instanceof Date) {
    return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
  }
  return String(range.getDisplayValue() || "").trim();
}

function aplicarEstiloGeral(range) {
  range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID)
       .setFontWeight("bold").setVerticalAlignment("middle");
}

// ============================================================
// PACIENTE E ENDEREÇO
// ============================================================

/**
 * Valor da célula editada. Usa a célula (getValue) quando o evento
 * a traz; aceita também um evento simulado { value: '...' }.
 */
function valorDoEvento_(e) {
  if (e && e.range) return e.range.getValue();
  return e ? e.value : '';
}

/** Linha da aba PACIENTES com este CNS (só dígitos), ou null. */
function localizarPacientePorCns_(cartao) {
  const dados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PACIENTES").getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).replace(/\D/g, "") === cartao) return dados[i];
  }
  return null;
}

function buscarPacientePorCartaoSUS(e) {
  const cartao = soDigitos_(valorDoEvento_(e));
  if (!cartao) return;

  const p = localizarPacientePorCns_(cartao);
  if (p) {
    preencherPaciente_(p);
    SpreadsheetApp.getActiveSpreadsheet().toast(String(p[1]), "✅ Paciente carregado", 4);
    return;
  }
  SpreadsheetApp.getUi().alert("⚠️ Paciente não encontrado no banco de dados.");
}

/**
 * Digitou (ou colou) o Cartão SUS do paciente na aba DADOS.
 * Se o número estiver certo, busca normalmente.
 * Se estiver errado, avisa; e se mesmo assim ele existir na aba
 * PACIENTES (cadastro antigo com erro), carrega e pede para corrigir.
 */
function aoEditarCnsPaciente_(e) {
  const cartao = soDigitos_(valorDoEvento_(e));
  if (!cartao) return;

  const motivo = motivoCNSInvalido_(cartao);
  if (!motivo) {
    buscarPacientePorCartaoSUS(e);
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const p  = localizarPacientePorCns_(cartao);
  if (p) {
    preencherPaciente_(p);
    ui.alert(
      "⚠️ Paciente carregado, mas o Cartão SUS dele está com erro\n\n" +
      String(p[1]).toUpperCase().trim() + "\nCNS " + cartao + ": " + motivo + ".\n\n" +
      "Confira no CADSUS e corrija aqui no formulário e na aba PACIENTES."
    );
    return;
  }

  ui.alert(
    "⚠️ Cartão SUS inválido\n\n" +
    "O número " + cartao + " " + motivo + ".\n\n" +
    "Confira o número e digite de novo."
  );
  campo_('cns_paciente').activate();
}

/**
 * Colunas da aba PACIENTES:
 * 0 CNS · 1 Nome · 2 Nascimento · 3 Sexo · 4 Mãe · 5 Telefone
 * 6 CEP · 7 Número · 8 Complemento · 9 CPF · 10 Prontuário
 */
function preencherPaciente_(p) {
  gravarCampo_('nome_paciente',   p[1]);
  gravarCampo_('prontuario',      p[10]);
  gravarCampo_('data_nascimento', p[2]);

  // "Masculino"/"Feminino" (vindo do Espelho CELK) → MASCULINO/FEMININO.
  // A aba APAC transforma isso no "X" do quadradinho certo.
  const sexo = String(p[3]).trim().toUpperCase();
  gravarCampo_('sexo', (sexo === "MASCULINO" || sexo === "FEMININO") ? sexo : "");

  gravarCampo_('nome_mae', p[4]);

  const tel = String(p[5]).replace(/\D/g, "");
  if (tel.length >= 10) {
    gravarCampo_('ddd',      tel.substring(0, 2));
    gravarCampo_('telefone', tel.substring(2));
  } else {
    gravarCampo_('ddd',      "");
    gravarCampo_('telefone', tel);
  }

  const cep = String(p[6]).replace(/\D/g, "");
  if (cep) {
    const cep8 = cep.padStart(8, "0");
    gravarCampo_('cep', cep8);
    preencherEnderecoPorCEP(cep8);
  } else {
    gravarCampo_('cep', "");
  }

  gravarCampo_('numero',      p[7]);
  gravarCampo_('complemento', p[8]);
}

function buscarEnderecoCEP(e) {
  const cep = String(valorDoEvento_(e) || "").replace(/\D/g, "");
  if (cep.length !== 8) return;
  preencherEnderecoPorCEP(cep);
}

function preencherEnderecoPorCEP(cep) {
  try {
    const res  = UrlFetchApp.fetch("https://brasilapi.com.br/api/cep/v1/" + cep);
    const data = JSON.parse(res.getContentText());

    const logradouro = (data.street       || "").toUpperCase();
    const bairro     = (data.neighborhood || "").toUpperCase();
    const cidade     = (data.city         || "").toUpperCase();
    const uf         = (data.state        || "").toUpperCase();

    gravarCampo_('logradouro', logradouro);
    gravarCampo_('bairro',     bairro);
    gravarCampo_('municipio',  cidade);
    gravarCampo_('uf',         uf);

    if (cidade && uf) {
      try {
        const resIbge = UrlFetchApp.fetch(
          "https://servicodados.ibge.gov.br/api/v1/localidades/estados/" + uf + "/municipios"
        );
        const municipios = JSON.parse(resIbge.getContentText());
        const municipio  = municipios.find(function(m) {
          return m.nome.toUpperCase().trim() === cidade.trim();
        });
        if (municipio) gravarCampo_('cod_ibge', String(municipio.id));
      } catch (errIbge) {
        Logger.log("❌ Erro IBGE: " + errIbge);
      }
    }
  } catch (err) {
    Logger.log("❌ Erro CEP: " + err);
  }
}

function testarViaCEP() {
  preencherEnderecoPorCEP("78450314");
}

// ============================================================
// ① ESPELHO CELK — leitura do PDF (OCR + regex)
// ============================================================

/**
 * Interpreta o TEXTO que o OCR tirou do Espelho CELK e devolve os
 * dados do paciente. Função "pura" (não mexe em planilha nem Drive),
 * por isso dá para testar com qualquer texto.
 *
 * Layouts já tratados:
 *   • rótulo antigo "N°:"  e rótulo novo "N." (CELK v3.1.34x)
 *   • CNS antes ou depois da palavra "CNS" no texto do OCR
 *   • telefone com ou sem hífen · CEP com ou sem pontos e traço
 */
function interpretarTextoEspelho_(bodyText) {
  const textoCorrido = bodyText.replace(/[\r\n]+/g, " ");

  let prontuario = "";
  const matchPront = bodyText.match(/C[oó]digo:\s*(\d+)/i)
                  || bodyText.match(/(\d+)\s+Paciente:/i);
  if (matchPront) prontuario = matchPront[1];

  let nome = "";
  const matchNome = bodyText.match(/Paciente:\s*([^\n\r]+)/i);
  if (matchNome) {
    nome = matchNome[1].trim();
    nome = nome.split(/Situaç[aã]o:/i)[0].split(/Dados Pessoais/i)[0].trim().toUpperCase();
  }

  let sexo = "";
  const matchSexo = bodyText.match(/Sexo:\s*(Masculino|Feminino)/i);
  if (matchSexo) {
    const s = matchSexo[1].trim();
    sexo = s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
  }

  let nascimento = "";
  const matchData = bodyText.match(/Data de Nasc\.:?\s*(\d{2}\/\d{2}\/\d{4})/i)
                 || bodyText.match(/(\d{2}\/\d{2}\/\d{4})/);
  if (matchData) nascimento = matchData[1];

  let mae = "";
  const matchMae = bodyText.match(/Nome da M[ãa]e:\s*([^\n\r]+)/i);
  if (matchMae) mae = matchMae[1].trim().toUpperCase();

  // Telefone: hífen opcional (alguns laudos vêm "(65) 999413091")
  let telefone = "";
  const matchTel = bodyText.match(/Telefone:\s*(\(\d{2}\)\s*\d{4,5}-?\d{4})/i);
  if (matchTel) telefone = matchTel[1].trim();

  // CEP: aceita 78450266, 78450-266, 78.450-266, 78.450 - 266
  let cep = "";
  const matchCep = bodyText.match(/CEP:\s*(\d{2}\.?\d{3}\s*-?\s*\d{3})/i)
                || bodyText.match(/CEP:\s*(\d+[-]?\d*)/i);
  if (matchCep) cep = matchCep[1].replace(/\D/g, "").substring(0, 8);

  // NÚMERO DO ENDEREÇO — o CELK passou a usar "N." em vez de "N°:".
  // Como "N." também aparece no bloco do CNS, remove primeiro o trecho
  // do CNS e só então procura "N." no que sobrou.
  let numCasa = "";
  let matchNum = bodyText.match(/N[°º]:\s*(\d+)\s+(?:Bairro|CEP|UF)/i)
              || bodyText.match(/Bairro:.*?N[°º]:\s*(\d+)/is)
              || bodyText.match(/N[°º]:\s*(\d+)/i);
  if (!matchNum) {
    const textoSemCns = bodyText.replace(/N[°º.]\s*(?:\d[.\s]?){15}\s*D?/i, "");
    matchNum = textoSemCns.match(/N\.\s*(\d{1,6})\b/);
  }
  if (matchNum) numCasa = matchNum[1].trim();

  let complemento = "";
  const matchComp = bodyText.match(/Complemento:\s*([^\n\r]+)/i);
  if (matchComp) {
    complemento = matchComp[1].split(/Bairro:/i)[0].split(/Cidade:/i)[0].split(/N[°º]:/i)[0].trim();
  }

  let cpf = "";
  const matchCpf = bodyText.match(/CPF\s+N[uú]mero:\s*([\d.-]+)/i)
                || bodyText.match(/N[uú]mero:\s*(\d{3}\.\d{3}\.\d{3}-\d{2})/i);
  if (matchCpf) cpf = matchCpf[1].trim();

  // CNS — procura TODOS os "N." + número que dão exatamente 15 dígitos.
  // Se houver mais de um candidato, fica com o primeiro que passa no
  // dígito verificador; se nenhum passar, fica com o primeiro (e o aviso
  // de "Cartão SUS com erro" aparece no painel). O padrão antigo fica de reserva.
  let cns = "";
  const candidatos = [];
  const reCns = /N[°º.]\s*((?:\d[.\s]?){15})/g;
  let m;
  while ((m = reCns.exec(textoCorrido)) !== null) {
    const dig = m[1].replace(/[^\d]/g, "");
    if (dig.length === 15) candidatos.push(dig);
  }
  if (candidatos.length) {
    cns = candidatos.filter(validarCNS)[0] || candidatos[0];
  }
  if (!cns) {
    const matchCnsAntigo = textoCorrido.match(/CNS\s+N[°º°o]?\s*([\d.]+)/i)
                        || textoCorrido.match(/CNS.*?N[°º°o\.]\s*([\d.]+)/i)
                        || textoCorrido.match(/N[°º]\s*(7\d[\d.]{10,})/i);
    if (matchCnsAntigo) cns = matchCnsAntigo[1].replace(/[^\d]/g, "").trim();
  }

  return {
    cns: cns, nome: nome, nascimento: nascimento, sexo: sexo, mae: mae,
    telefone: telefone, cep: cep, numCasa: numCasa, complemento: complemento,
    cpf: cpf, prontuario: prontuario
  };
}

/** Lê um Espelho CELK em PDF (OCR do Google Drive) e devolve os dados do paciente. */
function lerEspelhoCelk_(fileId) {
  let tempFileId = null;
  try {
    const tempFile = Drive.Files.copy(
      { title: "__temp_ocr_formatado__", mimeType: "application/vnd.google-apps.document" },
      fileId, { ocr: true, ocrLanguage: "pt-BR" }
    );
    tempFileId = tempFile.id;

    const bodyText = DocumentApp.openById(tempFileId).getBody().getText();
    Drive.Files.remove(tempFileId);
    tempFileId = null;

    return interpretarTextoEspelho_(bodyText);

  } finally {
    if (tempFileId) { try { Drive.Files.remove(tempFileId); } catch (_) {} }
  }
}

function nomeValido_(nome) {
  return !!nome && nome.indexOf("CÓDIGO") === -1;
}

/**
 * Campos importantes que o leitor NÃO encontrou no espelho, ou que
 * encontrou com erro (Cartão SUS ou CPF com dígito errado).
 * Serve de alarme: se o CELK mudar o layout, ou se o OCR ler um
 * número errado, o aviso aparece na hora no painel.
 */
function camposFaltantes_(d) {
  const falta = [];
  if (!d.cns) {
    falta.push('Cartão SUS (CNS)');
  } else if (!validarCNS(d.cns)) {
    falta.push('Cartão SUS válido (o número lido, ' + d.cns + ', ' + motivoCNSInvalido_(d.cns) + ')');
  }
  if (d.cpf && !validarCPF(d.cpf)) {
    falta.push('CPF válido (o número lido, ' + d.cpf + ', tem dígito errado)');
  }
  if (!d.nascimento) falta.push('Data de nascimento');
  if (!d.sexo)       falta.push('Sexo');
  if (!d.mae)        falta.push('Nome da mãe');
  if (!d.telefone)   falta.push('Telefone');
  if (!d.cep)        falta.push('CEP');
  if (!d.numCasa)    falta.push('Nº do endereço');
  return falta;
}

/** Já existe na aba PACIENTES? Confere CPF ou CNS (só os dígitos). */
function pacienteJaCadastrado_(sheet, cpf, cns) {
  const cpfDig = String(cpf || '').replace(/\D/g, '');
  const cnsDig = String(cns || '').replace(/\D/g, '');
  const ultima = sheet.getLastRow();
  if (ultima <= 1 || (!cpfDig && !cnsDig)) return false;

  const linhas = sheet.getRange(2, 1, ultima - 1, 10).getValues();
  return linhas.some(function(row) {
    const cnsRow = String(row[0]).replace(/\D/g, '');
    const cpfRow = String(row[9]).replace(/\D/g, '');
    return (cnsDig && cnsRow === cnsDig) || (cpfDig && cpfRow === cpfDig);
  });
}

function gravarPacienteNaPlanilha_(sheet, d) {
  const proxLinha = sheet.getLastRow() + 1;

  // CNS, telefone, CEP e CPF como TEXTO, para nunca virarem número
  [1, 6, 7, 10].forEach(function(coluna) {
    sheet.getRange(proxLinha, coluna).setNumberFormat("@");
  });

  const pacote = [
    String(d.cns), String(d.nome), String(d.nascimento), String(d.sexo),
    String(d.mae), String(d.telefone), String(d.cep), String(d.numCasa),
    String(d.complemento), String(d.cpf), String(d.prontuario)
  ];
  sheet.getRange(proxLinha, 1, 1, 11).setValues([pacote]);
  sheet.getRange(proxLinha, 1, 1, 11).setHorizontalAlignment("center");
  sheet.getRange(proxLinha, 2).setHorizontalAlignment("left");
  sheet.getRange(proxLinha, 5).setHorizontalAlignment("left");
}

// ① Abre o painel já na tela de escolha do Espelho CELK
function abrirPainelSelecaoPDF() {
  abrirPainelNoModo_('espelho');
}

// Chamado pelo painel quando você clica em "Selecionar"
function processarPDFselecionado(fileId, nomeArquivo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PACIENTES");
  if (!sheet) return { ok: false, msg: "❌ Aba PACIENTES não encontrada." };

  try {
    const d = lerEspelhoCelk_(fileId);

    if (!nomeValido_(d.nome)) {
      return { ok: false, msg: "❌ Não foi possível identificar o paciente no Espelho CELK." };
    }

    if (pacienteJaCadastrado_(sheet, d.cpf, d.cns)) {
      DriveApp.getFileById(fileId).setTrashed(true);
      if (d.cns) {
        gravarCampo_('cns_paciente', d.cns);
        SpreadsheetApp.flush();
        buscarPacientePorCartaoSUS({ value: d.cns });
      }
      return { ok: true, duplicado: true, nome: d.nome };
    }

    gravarPacienteNaPlanilha_(sheet, d);
    DriveApp.getFileById(fileId).setTrashed(true);

    if (d.cns) {
      gravarCampo_('cns_paciente', d.cns);
      SpreadsheetApp.flush();
    }

    return { ok: true, duplicado: false, nome: d.nome, cns: d.cns, avisos: camposFaltantes_(d) };

  } catch (falha) {
    return { ok: false, msg: "❌ Erro ao processar: " + falha.message };
  }
}

function carregarPacienteNoFormulario(cns) {
  buscarPacientePorCartaoSUS({ value: cns });
}

function listarPDFsEntrada() {
  const folder    = pastaConfig_('FOLDER_ENTRADA_ID');
  const filesIter = folder.getFilesByType("application/pdf");
  const arquivos  = [];
  while (filesIter.hasNext()) {
    const f = filesIter.next();
    arquivos.push({ id: f.getId(), nome: f.getName() });
  }
  return arquivos;
}

/**
 * Versão antiga (processa o único PDF da pasta, sem painel).
 * Não está no menu; mantida por compatibilidade.
 */
function extrairDadosPDFparaPlanilhaV3() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ui    = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName("PACIENTES");

  if (!sheet) {
    ui.alert("❌ Erro: A aba 'PACIENTES' não foi encontrada.");
    return;
  }

  const folder     = pastaConfig_('FOLDER_ENTRADA_ID');
  const filesIter  = folder.getFilesByType("application/pdf");
  const listaFiles = [];
  while (filesIter.hasNext()) listaFiles.push(filesIter.next());

  if (listaFiles.length === 0) {
    ui.alert("⚠ Atenção", "Nenhum PDF encontrado na pasta de entrada.", ui.ButtonSet.OK);
    return;
  }
  if (listaFiles.length > 1) {
    ui.alert("⚠️ Atenção", "Há " + listaFiles.length + " PDFs na pasta de entrada.\nDeixe apenas 1 PDF por vez e tente novamente.", ui.ButtonSet.OK);
    return;
  }

  let importados = 0;
  let ultimoCns  = "";

  listaFiles.forEach(function(file) {
    try {
      const d = lerEspelhoCelk_(file.getId());
      if (!nomeValido_(d.nome)) {
        Logger.log("❌ Inconsistência crítica: " + file.getName());
        return;
      }
      if (pacienteJaCadastrado_(sheet, d.cpf, d.cns)) {
        Logger.log("⚠ Paciente já cadastrado (CPF " + d.cpf + " / CNS " + d.cns + "). Ignorando.");
        file.setTrashed(true);
        return;
      }
      gravarPacienteNaPlanilha_(sheet, d);
      importados++;
      ultimoCns = d.cns;
      file.setTrashed(true);
    } catch (falha) {
      Logger.log("❌ Erro em " + file.getName() + ": " + falha.message);
    }
  });

  if (ultimoCns) {
    gravarCampo_('cns_paciente', ultimoCns);
    SpreadsheetApp.flush();
  }

  if (importados > 0) {
    const resposta = ui.alert("✅ Sucesso!", importados + " paciente(s) importado(s).\n\nDeseja carregar os dados do paciente no formulário agora?", ui.ButtonSet.YES_NO);
    if (resposta === ui.Button.YES) buscarPacientePorCartaoSUS({ value: ultimoCns });
  } else {
    ui.alert("⚠ Atenção", "Nenhum dado novo extraído. Recarregue os PDFs na pasta de ENTRADA.", ui.ButtonSet.OK);
  }
}

// ============================================================
// ③ GERAR APAC EM PDF
// ============================================================

// Textos que parecem preenchidos, mas na prática são "vazio"
const PLACEHOLDERS_VAZIOS_ = ['SELECIONE', 'DIGITE O NOME AQUI'];

/** Texto do campo, ou '' se estiver em branco ou só com um texto-padrão ("SELECIONE"). */
function textoPreenchido_(campo) {
  const v = String(textoCampo_(campo) || '').trim();
  if (!v) return '';
  return PLACEHOLDERS_VAZIOS_.indexOf(v.toUpperCase()) !== -1 ? '' : v;
}

/**
 * Confere os campos que não podem ir em branco (ou errados) na APAC
 * e devolve a lista do que precisa de atenção (lista vazia = tudo certo).
 *
 * Para incluir ou tirar uma exigência, mexa só aqui.
 */
function camposObrigatoriosFaltando_() {
  const falta = [];
  const temNumero = function(v) { return /[1-9]/.test(soDigitos_(v)); };   // "0000000" conta como vazio
  const qtdOk     = function(campo) { return parseInt(textoPreenchido_(campo), 10) > 0; };

  if (!textoPreenchido_('estab_solicitante')) falta.push('Equipe (nome do estabelecimento)');
  if (!temNumero(textoPreenchido_('cnes_solicitante'))) falta.push('CNES');

  // Cartão SUS do paciente: tem que existir e passar no dígito verificador
  const cnsPac = soDigitos_(textoPreenchido_('cns_paciente'));
  if (!cnsPac) {
    falta.push('Cartão SUS do paciente');
  } else {
    const motivoPac = motivoCNSInvalido_(cnsPac);
    if (motivoPac) falta.push('Cartão SUS do paciente com erro (' + motivoPac + ')');
  }

  if (!textoPreenchido_('proc_codigo')) falta.push('Procedimento: código SIGTAP');
  if (!textoPreenchido_('proc_nome'))   falta.push('Procedimento: nome');
  if (!qtdOk('proc_qtd'))               falta.push('Quantidade do procedimento');

  if (!textoPreenchido_('cid_principal')) falta.push('CID-10 principal');

  if (!textoPreenchido_('esf_profissional')) falta.push('Médico (profissional solicitante)');

  // Documento do médico: CNS ou CPF, conforme o tipo escolhido
  const docTipo   = textoPreenchido_('esf_doc_tipo').toUpperCase();
  const docNumero = soDigitos_(textoPreenchido_('esf_doc_numero'));
  if (!temNumero(docNumero)) {
    falta.push(docTipo === 'CPF' ? 'CPF do médico' : 'Cartão SUS do médico');
  } else {
    const motivoMed = motivoDocProfissional_(docTipo, docNumero);
    if (motivoMed) falta.push('Documento do médico com erro (' + motivoMed + ')');
  }

  // Procedimento secundário preenchido sem quantidade
  for (let i = 1; i <= 5; i++) {
    const temProc = textoPreenchido_('sec' + i + '_codigo') || textoPreenchido_('sec' + i + '_nome');
    if (temProc && !qtdOk('sec' + i + '_qtd')) {
      falta.push('Quantidade do procedimento secundário ' + i);
    }
  }
  return falta;
}

// Chamado pelo painel antes de gerar: devolve o que está em branco ou errado
function verificarCamposPdfPainel() {
  try {
    return { ok: true, faltando: camposObrigatoriosFaltando_() };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}

// Versão do menu: confere os campos, gera e mostra a janelinha com o link
function gerarPdfApacs() {
  try {
    const ui      = SpreadsheetApp.getUi();
    const faltando = camposObrigatoriosFaltando_();
    if (faltando.length) {
      const resposta = ui.alert(
        '⚠️ Campos em branco ou com erro',
        'Estes campos da APAC estão vazios, incompletos ou com erro:\n\n• ' + faltando.join('\n• ') +
        '\n\nO PDF sairá assim. Gerar mesmo assim?',
        ui.ButtonSet.YES_NO
      );
      if (resposta !== ui.Button.YES) return;
    }

    const r = gerarPdfApac_();

    const html = HtmlService.createHtmlOutput(
      '<html><head><style>' +
      'body{font-family:Arial,sans-serif;padding:20px;text-align:center;}' +
      'h3{color:#2e7d32;}p{font-size:13px;color:#444;word-break:break-all;}' +
      '.btn{display:inline-block;margin:8px;padding:10px 20px;border-radius:6px;font-weight:bold;font-size:14px;cursor:pointer;border:none;}' +
      '.btn-abrir{background:#1F4E79;color:#fff;}.btn-fechar{background:#e0e0e0;color:#333;}' +
      '</style></head><body>' +
      '<h3>✅ APAC gerada com sucesso!</h3>' +
      '<p>' + r.nomeArquivo + '.pdf</p><br>' +
      '<button class="btn btn-abrir" onclick="window.open(\'' + r.url + '\',\'_blank\')">🖨️ Abrir / Imprimir PDF</button>' +
      '<button class="btn btn-fechar" onclick="google.script.host.close()">Fechar</button>' +
      '</body></html>'
    ).setWidth(400).setHeight(220);

    SpreadsheetApp.getUi().showModalDialog(html, "✅ APAC gerada");

  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ Erro ao gerar a APAC em PDF: " + err.message);
    Logger.log("Erro gerarPdfApacs: " + err);
  }
}

// Faz o trabalho e devolve { nomeArquivo, url }
function gerarPdfApac_() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const abaApac = ss.getSheetByName("APAC");
  if (!abaApac) throw new Error("Aba APAC não encontrada.");

  const nome = textoCampo_('nome_paciente').toUpperCase();
  if (!nome) throw new Error("Carregue um paciente antes de gerar a APAC.");

  const pasta = pastaConfig_('FOLDER_PDF_GERADO_ID');

  const especialidade = textoCampo_('proc_nome').toUpperCase();
  const data          = textoCampo_('esf_data').replace(/\//g, "-");
  const nomeArquivo   = (nome + "_" + especialidade + "_" + data).replace(/\s+/g, "_");

  abaApac.showSheet();
  SpreadsheetApp.flush();

  try {
    const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() +
                "/export?format=pdf&gid=" + abaApac.getSheetId() +
                "&size=A4&portrait=true&fitw=true" +
                "&sheetnames=false&printtitle=false" +
                "&pagenumbers=false&gridlines=false&fzr=false" +
                "&top_margin=0.2067&bottom_margin=0.2067" +
                "&left_margin=0.25&right_margin=0.25";

    const token    = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { "Authorization": "Bearer " + token } });
    const pdfBlob  = response.getBlob().setName(nomeArquivo + ".pdf");

    // O PDF fica privado: só quem tem acesso à pasta abre.
    const arquivo = pasta.createFile(pdfBlob);
    return { nomeArquivo: nomeArquivo, url: arquivo.getUrl() };

  } finally {
    abaApac.hideSheet();
    SpreadsheetApp.flush();
  }
}

// ============================================================
// ⑤ LIMPAR FORMULÁRIO
// ============================================================

// Versão do menu
function limparFormularioAPAC() {
  limparFormulario_();
  SpreadsheetApp.getActiveSpreadsheet().toast("Pronto para o próximo paciente!", "✨ Formulário limpo", 4);
}

function limparFormulario_() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const prefixo = LAYOUT_DADOS.PREFIXO_NOME;
  const manter  = ['data_recebimento'];   // fórmula =TODAY(), não apagar

  ss.getNamedRanges().forEach(function(n) {
    const nome = n.getName();
    if (nome.indexOf(prefixo) !== 0) return;
    if (manter.indexOf(nome.substring(prefixo.length)) !== -1) return;
    const rg = n.getRange();
    if (rg.getSheet().getName() !== "DADOS") return;
    rg.clearContent();
  });

  gravarCampo_('esf_doc_tipo', 'CNS');
  campo_('cns_paciente').activate();
}

// ============================================================
// GATILHOS
// ============================================================

/**
 * Instala (se faltar) os gatilhos do projeto e devolve a lista
 * do que foi instalado agora:
 *   • onEditAPAC            → buscas automáticas ao digitar (CNS, CEP, CNES, médico)
 *   • abrirPainelAPAC       → abre o painel lateral ao abrir a planilha
 *   • sugestoesSigtapMensal → todo dia 1º, monta as sugestões para a Referencia_SIGTAP
 *   • verificarAtualizacaoSemanal → toda segunda, só confere se há versão nova (não atualiza)
 */
function instalarGatilhos_() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const gatilhos = ScriptApp.getProjectTriggers();
  const existe   = function(nomeFuncao) {
    return gatilhos.some(function(g) { return g.getHandlerFunction() === nomeFuncao; });
  };

  const instalados = [];

  if (!existe("onEditAPAC")) {
    ScriptApp.newTrigger("onEditAPAC").forSpreadsheet(ss).onEdit().create();
    instalados.push("buscas automáticas");
  }
  if (!existe("abrirPainelAPAC")) {
    ScriptApp.newTrigger("abrirPainelAPAC").forSpreadsheet(ss).onOpen().create();
    instalados.push("painel ao abrir a planilha");
  }
  if (!existe("sugestoesSigtapMensal")) {
    ScriptApp.newTrigger("sugestoesSigtapMensal").timeBased().onMonthDay(1).atHour(7).create();
    instalados.push("sugestões SIGTAP todo dia 1º");
  }
  if (!existe("verificarAtualizacaoSemanal")) {
    ScriptApp.newTrigger("verificarAtualizacaoSemanal").timeBased()
             .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
    instalados.push("conferência semanal de atualização");
  }
  return instalados;
}

// Versão do menu
function instalarGatilhoSeNecessario() {
  const instalados = instalarGatilhos_();
  SpreadsheetApp.getUi().alert(
    instalados.length
      ? "✅ Gatilho(s) instalado(s): " + instalados.join(" e ") + "."
      : "✅ Todos os gatilhos já estão ativos!"
  );
}
