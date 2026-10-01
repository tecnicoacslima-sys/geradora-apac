/**
 * ============================================================
 * Validacao.gs — CONFERÊNCIA DE CARTÃO SUS (CNS) E CPF
 * ============================================================
 * Confere o dígito verificador sem gastar nada de IA.
 * Pega na hora um número digitado errado ou lido errado pelo
 * OCR do Espelho CELK (ex.: um 8 que virou 3).
 *
 * Usado por:
 *   • Código.gs  → ao digitar o CNS do paciente e do médico,
 *                  ao ler o Espelho CELK, antes de gerar o PDF
 *                  e ao salvar no Protocolo
 *   • ExtratorPDFDireto.gs → teste do extrator
 *   • Menu 🔧 Ferramentas ▸ 🪪 Conferir Cartões SUS cadastrados
 * ============================================================
 */

/** Só os números de um texto (ex.: "700.6069 8540-2967" → "700606985402967"). */
function soDigitos_(valor) {
  return String(valor === null || valor === undefined ? '' : valor).replace(/\D/g, '');
}

/**
 * Explica por que um CNS é inválido.
 * Devolve '' (texto vazio) quando o CNS está CERTO.
 *
 * Regra oficial: 15 dígitos; começa com 1 ou 2 (definitivo)
 * ou 7, 8 ou 9 (provisório); a soma de cada dígito multiplicado
 * pelos pesos 15, 14, 13 ... 1 tem que ser divisível por 11.
 */
function motivoCNSInvalido_(cns) {
  const s = soDigitos_(cns);
  if (!s) return 'está em branco';
  if (s.length !== 15) return 'tem ' + s.length + ' dígitos (o Cartão SUS tem 15)';
  if (!/^[12789]/.test(s)) return 'começa com ' + s.charAt(0) + ' (deve começar com 1, 2, 7, 8 ou 9)';

  let soma = 0;
  for (let i = 0; i < 15; i++) soma += parseInt(s.charAt(i), 10) * (15 - i);
  if (soma % 11 !== 0) return 'o dígito verificador não confere (provavelmente um número trocado)';
  return '';
}

/** true se o CNS está certo. */
function validarCNS(cns) {
  return motivoCNSInvalido_(cns) === '';
}

/** true se o CPF está certo (11 dígitos e os dois dígitos verificadores conferem). */
function validarCPF(cpf) {
  const s = soDigitos_(cpf);
  if (s.length !== 11) return false;
  if (/^(\d)\1{10}$/.test(s)) return false;   // 000.000.000-00, 111.111.111-11...

  for (let dv = 9; dv < 11; dv++) {
    let soma = 0;
    for (let i = 0; i < dv; i++) soma += parseInt(s.charAt(i), 10) * ((dv + 1) - i);
    let resto = (soma * 10) % 11;
    if (resto === 10) resto = 0;
    if (resto !== parseInt(s.charAt(dv), 10)) return false;
  }
  return true;
}

// ============================================================
// 🪪 CONFERIR CARTÕES SUS CADASTRADOS (menu 🔧 Ferramentas)
// ============================================================

/**
 * Passa por todos os médicos (MEDICO_SUS) e pacientes (PACIENTES)
 * e mostra quem está com Cartão SUS ou CPF inválido, com o número
 * da linha, para você corrigir. Não altera nada.
 */
function conferirCartoesSUS() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const problemas = [];
  let medicos = 0, pacientes = 0;

  const abaMed = ss.getSheetByName('MEDICO_SUS');
  if (abaMed && abaMed.getLastRow() >= 2) {
    abaMed.getRange(2, 1, abaMed.getLastRow() - 1, 2).getValues().forEach(function(l, i) {
      const nome = String(l[0] || '').trim();
      if (!nome || nome.toUpperCase() === 'SELECIONE') return;
      medicos++;
      const doc = soDigitos_(l[1]);
      if (doc.length === 11) {
        if (!validarCPF(doc)) problemas.push('MEDICO_SUS linha ' + (i + 2) + ' · ' + nome + ' · CPF ' + doc + ' com dígito errado');
        return;
      }
      const motivo = motivoCNSInvalido_(doc);
      if (motivo) problemas.push('MEDICO_SUS linha ' + (i + 2) + ' · ' + nome + ' · CNS ' + (doc || '—') + ' ' + motivo);
    });
  }

  const abaPac = ss.getSheetByName('PACIENTES');
  if (abaPac && abaPac.getLastRow() >= 2) {
    abaPac.getRange(2, 1, abaPac.getLastRow() - 1, 10).getValues().forEach(function(l, i) {
      const nome = String(l[1] || '').trim();
      if (!nome) return;
      pacientes++;
      const cns = soDigitos_(l[0]);
      const motivo = motivoCNSInvalido_(cns);
      if (motivo) problemas.push('PACIENTES linha ' + (i + 2) + ' · ' + nome + ' · CNS ' + (cns || '—') + ' ' + motivo);
      const cpf = soDigitos_(l[9]);
      if (cpf && !validarCPF(cpf)) problemas.push('PACIENTES linha ' + (i + 2) + ' · ' + nome + ' · CPF ' + cpf + ' com dígito errado');
    });
  }

  const resumo = 'Conferidos: ' + medicos + ' profissional(is) e ' + pacientes + ' paciente(s).';
  if (!problemas.length) {
    ui.alert('🪪 Cartões SUS', '✅ Nenhum problema encontrado.\n\n' + resumo, ui.ButtonSet.OK);
    return;
  }

  const LIMITE = 25;
  ui.alert(
    '🪪 Cartões SUS — ' + problemas.length + ' problema(s)',
    '• ' + problemas.slice(0, LIMITE).join('\n• ') +
    (problemas.length > LIMITE ? '\n\n… e mais ' + (problemas.length - LIMITE) + '.' : '') +
    '\n\n' + resumo + '\nConfira esses números no CADSUS ou no CNES e corrija na aba indicada.',
    ui.ButtonSet.OK
  );
}
