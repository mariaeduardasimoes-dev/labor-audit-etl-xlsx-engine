/**
 * DIRECT NATIVE .XLSX AUDIT AUTOMATION - 100% MATHEMATICAL CLOSING ACROSS ALL PLANTS
 * Governance Rule: 1 Lawsuit = 1 Primary Claim (Zero Double-Counting)
 */

const ABA_ORIGEM = 'Pedidos gerais';
const ABA_DESTINO = 'Relatorio_Automatico';
const UNIDADES_FIXAS = ['Plant_South', 'Plant_Southeast', 'Plant_North', 'Corporate_HQ'];

function getTargetXlsxId_() {
  return PropertiesService.getScriptProperties().getProperty('AUDIT_XLSX_FILE_ID') || 'YOUR_TARGET_XLSX_FILE_ID';
}

function atualizarPlanilhaXlsxOriginal() {
  const ARQUIVO_XLSX_ID = getTargetXlsxId_();
  const isDriveV3 = typeof Drive.Files.create === 'function';
  let tempFile;

  if (isDriveV3) {
    tempFile = Drive.Files.copy(
      { name: 'TEMP_CALCULO_TRABALHISTA', mimeType: 'application/vnd.google-apps.spreadsheet' },
      ARQUIVO_XLSX_ID,
      { supportsAllDrives: true }
    );
  } else {
    tempFile = Drive.Files.copy(
      { title: 'TEMP_CALCULO_TRABALHISTA', mimeType: 'application/vnd.google-apps.spreadsheet' },
      ARQUIVO_XLSX_ID,
      { convert: true, supportsAllDrives: true }
    );
  }

  const tempId = tempFile.id;
  try {
    const ss = SpreadsheetApp.openById(tempId);
    processarRelatorioNaMesmaPagina_(ss);
    SpreadsheetApp.flush();

    const urlExport = 'https://docs.google.com/spreadsheets/d/' + tempId + '/export?format=xlsx';
    const xlsxBlob = UrlFetchApp.fetch(urlExport, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
    }).getBlob();

    Drive.Files.update({}, ARQUIVO_XLSX_ID, xlsxBlob, { supportsAllDrives: true });
    Logger.log('✅ Relatorio_Automatico updated successfully closing 100% across all units!');
  } finally {
    DriveApp.getFileById(tempId).setTrashed(true);
  }
}

function normalizarUnidade_(texto) {
  const t = String(texto || '').trim().toLowerCase();
  if (!t) return '';
  if (t.includes('south') || t.includes('sul') || t.includes('unidade 1')) return 'Plant_South';
  if (t.includes('southeast') || t.includes('sudeste') || t.includes('unidade 2')) return 'Plant_Southeast';
  if (t.includes('north') || t.includes('norte') || t.includes('unidade 3')) return 'Plant_North';
  if (t.includes('hq') || t.includes('matriz') || t.includes('corporate')) return 'Corporate_HQ';
  return String(texto).trim();
}

function gerarBarraVisual_(percentual) {
  const p = Math.max(0, Math.min(1, Number(percentual) || 0));
  const blocosCheios = Math.round(p * 12);
  const blocosVazios = 12 - blocosCheios;
  return '█'.repeat(blocosCheios) + '░'.repeat(blocosVazios);
}

function classificarPedidoUnico_(resumoRaw, regrasPedidos) {
  let melhorPedido = 'Outros Pedidos Trabalhistas';
  let menorPosicao = Infinity;

  regrasPedidos.forEach(regra => {
    let textoAnalise = resumoRaw;
    if (regra.limparSemJustaCausa) {
      textoAnalise = textoAnalise.replace(/sem\s+justa\s+causa/gi, ' ');
    }
    const match = textoAnalise.match(regra.regex);
    if (match && match.index !== undefined && match.index < menorPosicao) {
      menorPosicao = match.index;
      melhorPedido = regra.nome;
    }
  });
  return melhorPedido;
}

function processarRelatorioNaMesmaPagina_(ss) {
  let sheetOrigem = ss.getSheetByName(ABA_ORIGEM);
  if (!sheetOrigem) {
    sheetOrigem = ss.getSheets().find(s =>
      s.getName().trim().toLowerCase().replace(/_/g, ' ').includes('pedidos gerais')
    );
  }
  if (!sheetOrigem) {
    throw new Error('Aba "Pedidos gerais" não encontrada!');
  }

  ['Relatorio_Proprios', 'Relatorio_Terceiros'].forEach(nomeAba => {
    const abaExtra = ss.getSheetByName(nomeAba);
    if (abaExtra) ss.deleteSheet(abaExtra);
  });

  const dados = sheetOrigem.getDataRange().getValues();
  if (dados.length < 2) return;

  const cabecalho = dados[0].map(c => String(c).trim().toLowerCase());
  const idxUnidade = cabecalho.findIndex(c => c.includes('centro de custo') || c.includes('unidade'));
  const idxTipo = cabecalho.findIndex(c => c.includes('processo - tipo') || c.includes('tipo'));
  const idxResumo = cabecalho.findIndex(c => c.includes('resumo'));

  const colUnidade = idxUnidade !== -1 ? idxUnidade : 1;
  const colTipo = idxTipo !== -1 ? idxTipo : 3;
  const colResumo = idxResumo !== -1 ? idxResumo : 4;

  const regrasPedidos = [
    {
      nome: 'Insalubridade / Periculosidade',
      regex: /insalubridade|periculosidade|ambiente\s+insalubre|agentes\s+nocivos|agentes\s+insalubres|agentes\s+perigosos/i,
      limparSemJustaCausa: false
    },
    {
      nome: 'Horas Extras',
      regex: /horas?\s+extras?|sobrejornada|extrapola[çc][ãa]o\s+de\s+jornada|intervalo\s+intrajornada|intervalo\s+interjornada|banco\s+de\s+horas|minutos\s+residuais|tempo\s+[àa]\s+disposi[çc][ãa]o/i,
      limparSemJustaCausa: false
    },
    {
      nome: 'Doença Ocupacional',
      regex: /doen[çc]a\s+ocupacional|doen[çc]a\s+profissional|doen[çc]a\s+do\s+trabalho|acidente\s+de\s+trabalho|acidente\s+t[íi]pico|les[ãa]o\s+por\s+esfor[çc]o|ler\/dort|epicondilite|t[úu]nel\s+do\s+carpo|nexo\s+causal|nexo\s+concausal|estabilidade\s+provis[óo]ria\s+acident[áa]ria|estabilidade\s+acident[áa]ria|redu[çc][ãa]o\s+da\s+capacidade\s+labor/i,
      limparSemJustaCausa: false
    },
    {
      nome: 'Reversão de Justa Causa',
      regex: /justa\s+causa|revers[ãa]o.*demiss[ãa]o|falta\s+grave\s+aplicada|des[íi]dia/i,
      limparSemJustaCausa: true
    },
    {
      nome: 'Rescisão Indireta',
      regex: /rescis[ãa]o\s+indireta/i,
      limparSemJustaCausa: false
    },
    {
      nome: 'Desvio de Função / Equiparação Salarial',
      regex: /desvio\s+de\s+fun[çc][ãa]o|ac[úu]mulo\s+de\s+fun[çc][ãa]o|equipara[çc][ãa]o\s+salarial|diferen[çc]as\s+salariais\s+por\s+equipara[çc][ãa]o|paradigma/i,
      limparSemJustaCausa: false
    },
    {
      nome: 'Assédio (Moral/Sexual)',
      regex: /ass[ée]dio\s+moral|ass[ée]dio\s+sexual|ass[ée]dio/i,
      limparSemJustaCausa: false
    }
  ];

  const todasCategoriasPedidos = regrasPedidos.map(r => r.nome).concat(['Outros Pedidos Trabalhistas']);

  // Sanitized Third-Party Vendor Dictionary (Pattern Matching)
  const regrasEmpresasTerceiras = [
    { nome: 'Vendor_TempStaffing_A', regex: /vendor\s*a|temp\s*staffing/i },
    { nome: 'Vendor_Facilities_B', regex: /vendor\s*b|facilities\s*corp/i },
    { nome: 'Vendor_Maintenance_C', regex: /vendor\s*c|maint\s*services/i },
    { nome: 'Vendor_Security_D', regex: /vendor\s*d|sec\s*group/i },
    { nome: 'Vendor_Logistics_E', regex: /vendor\s*e|cargo\s*logistics/i },
    { nome: 'Vendor_Transport_F', regex: /vendor\s*f|transport\s*sa/i }
  ];

  let totalProprios = 0;
  let totalTerceiros = 0;
  let totalPendentes = 0;
  let totalTerceirosComEmpresa = 0;

  const propriosPorUnidade = { 'Plant_South': 0, 'Plant_Southeast': 0, 'Plant_North': 0, 'Corporate_HQ': 0 };
  const terceirosPorUnidade = { 'Plant_South': 0, 'Plant_Southeast': 0, 'Plant_North': 0, 'Corporate_HQ': 0 };
  const contagemEmpresasTotal = {};
  const contagemPedidos = {};

  todasCategoriasPedidos.forEach(nomeCat => {
    contagemPedidos[nomeCat] = {
      total: 0,
      porUnidade: { 'Plant_South': 0, 'Plant_Southeast': 0, 'Plant_North': 0, 'Corporate_HQ': 0 }
    };
  });

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const unidadeNorm = normalizarUnidade_(linha[colUnidade]);
    const tipoRaw = String(linha[colTipo] || '').trim();
    const resumoRaw = String(linha[colResumo] || '').trim();

    if (!unidadeNorm || !tipoRaw || !resumoRaw || resumoRaw === '-') {
      totalPendentes++;
      continue;
    }

    if (/pr[óo]prio/i.test(tipoRaw)) {
      totalProprios++;
      propriosPorUnidade[unidadeNorm] = (propriosPorUnidade[unidadeNorm] || 0) + 1;
      const pedidoUnico = classificarPedidoUnico_(resumoRaw, regrasPedidos);
      contagemPedidos[pedidoUnico].total++;
      contagemPedidos[pedidoUnico].porUnidade[unidadeNorm] = (contagemPedidos[pedidoUnico].porUnidade[unidadeNorm] || 0) + 1;
    } else if (/terceir/i.test(tipoRaw)) {
      totalTerceiros++;
      terceirosPorUnidade[unidadeNorm] = (terceirosPorUnidade[unidadeNorm] || 0) + 1;
      for (let j = 0; j < regrasEmpresasTerceiras.length; j++) {
        if (regrasEmpresasTerceiras[j].regex.test(resumoRaw)) {
          const nomeEmp = regrasEmpresasTerceiras[j].nome;
          contagemEmpresasTotal[nomeEmp] = (contagemEmpresasTotal[nomeEmp] || 0) + 1;
          totalTerceirosComEmpresa++;
          break;
        }
      }
    } else {
      totalPendentes++;
    }
  }

  const totalGlobal = totalProprios + totalTerceiros;
  const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  let sheet = ss.getSheetByName(ABA_DESTINO);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_DESTINO);
  } else {
    if (sheet.getFilter()) sheet.getFilter().remove();
    sheet.clear();
  }

  sheet.setFrozenRows(0);
  sheet.setHiddenGridlines(true);

  const matriz = [];
  matriz.push(['⚖️ PAINEL EXECUTIVO DE AUDITORIA TRABALHISTA | LEGAL OPS', '', '', '', '', '']);
  matriz.push([
    '📅 Atualizado: ' + dataAtual,
    '✅ Total Identificados: ' + totalGlobal,
    '👔 Próprios: ' + totalProprios + ' (' + (totalGlobal ? ((totalProprios / totalGlobal) * 100).toFixed(1) : 0) + '%)',
    '🏢 Terceiros: ' + totalTerceiros + ' (' + (totalGlobal ? ((totalTerceiros / totalGlobal) * 100).toFixed(1) : 0) + '%)',
    '⏳ Ignorados (Sem Identificação): ' + totalPendentes,
    ''
  ]);
  matriz.push(['', '', '', '', '', '']);

  const linhaBannerProprios = matriz.length + 1;
  matriz.push(['👔 PARTE 1: ANÁLISE EXCLUSIVA DE EMPREGADOS PRÓPRIOS (' + totalProprios + ' PROCESSOS = 100%)', '', '', '', '', '']);

  const linhaCabResumoProp = matriz.length + 1;
  matriz.push([
    'Centro de Custo - Unidade (Próprios)', 'Qtd. Processos da Unidade', '% do Total (Soma = 100%)',
    'Pedido Principal (1 por Processo)', 'Qtd. Processos', '% sobre ' + totalProprios + ' (Soma = 100%)'
  ]);

  const linhasEsqProp = [
    ['Plant_South', propriosPorUnidade['Plant_South'] || 0, totalProprios ? (propriosPorUnidade['Plant_South'] || 0) / totalProprios : 0],
    ['Plant_Southeast', propriosPorUnidade['Plant_Southeast'] || 0, totalProprios ? (propriosPorUnidade['Plant_Southeast'] || 0) / totalProprios : 0],
    ['Plant_North', propriosPorUnidade['Plant_North'] || 0, totalProprios ? (propriosPorUnidade['Plant_North'] || 0) / totalProprios : 0],
    ['Corporate_HQ', propriosPorUnidade['Corporate_HQ'] || 0, totalProprios ? (propriosPorUnidade['Corporate_HQ'] || 0) / totalProprios : 0],
    ['TOTAL DE PRÓPRIOS', totalProprios, totalProprios ? 1 : 0],
    ['', '', ''], ['', '', ''], ['', '', ''], ['', '', '']
  ];

  const categoriasOrdenadas = todasCategoriasPedidos.slice().sort((a, b) => {
    if (a === 'Outros Pedidos Trabalhistas') return 1;
    if (b === 'Outros Pedidos Trabalhistas') return -1;
    return contagemPedidos[b].total - contagemPedidos[a].total;
  });

  const linhasDirProp = [];
  categoriasOrdenadas.forEach(nomeCat => {
    const totPed = contagemPedidos[nomeCat].total;
    const pctPed = totalProprios ? totPed / totalProprios : 0;
    linhasDirProp.push([nomeCat, totPed, pctPed]);
  });
  linhasDirProp.push(['TOTAL DE PROCESSOS PRÓPRIOS', totalProprios, totalProprios ? 1 : 0]);

  for (let idx = 0; idx < 9; idx++) {
    const esq = linhasEsqProp[idx];
    const dir = linhasDirProp[idx];
    matriz.push([esq[0], esq[1], esq[2], dir[0], dir[1], dir[2]]);
  }

  matriz.push(['', '', '', '', '', '']);
  const linhaTitFiltroProp = matriz.length + 1;
  matriz.push(['🔍 DETALHAMENTO DE PRÓPRIOS POR PEDIDO E UNIDADE (USE AS SETAS ▼ ABAIXO PARA FILTRAR)', '', '', '', '', '']);

  const linhaCabFiltroProp = matriz.length + 1;
  matriz.push([
    'Pedido Principal (1 por Processo)',
    'Centro de Custo - Unidade',
    'Qtd. de Processos',
    '% dentro da Unidade\n(Soma da Unidade = 100%)',
    '% no Total de Próprios\n(Soma das 4 Unidades = 100%)',
    'Termômetro Visual (na Unidade)'
  ]);

  const zebraProp = [];
  categoriasOrdenadas.forEach(nomeCat => {
    const d = contagemPedidos[nomeCat];
    UNIDADES_FIXAS.forEach(u => {
      const q = d.porUnidade[u] || 0;
      const baseUnid = propriosPorUnidade[u] || 0;
      const pctUnid = baseUnid ? q / baseUnid : 0;
      const pctGeral = totalProprios ? q / totalProprios : 0;
      matriz.push([nomeCat, u, q, pctUnid, pctGeral, gerarBarraVisual_(pctUnid)]);
      if (matriz.length % 2 === 0) zebraProp.push(matriz.length);
    });
  });

  const linhaFimFiltroProp = matriz.length;
  const rangeTotal = sheet.getRange(1, 1, matriz.length, 6);
  rangeTotal.setValues(matriz);
  rangeTotal.setFontFamily('Calibri').setFontSize(11).setVerticalAlignment('middle');
  sheet.setRowHeights(1, matriz.length, 32);

  sheet.getRange('A1:F1').merge().setBackground('#1B365D').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(15);
  sheet.getRange(linhaBannerProprios, 1, 1, 6).merge().setBackground('#1B365D').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13);
  sheet.getRange(linhaTitFiltroProp, 1, 1, 6).merge().setBackground('#1B365D').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12);
  sheet.getRange(linhaCabFiltroProp, 1, 1, 6).setBackground('#2F5597').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);

  const qtdLinhasFiltroProp = linhaFimFiltroProp - linhaCabFiltroProp;
  sheet.getRange(linhaCabFiltroProp, 1, qtdLinhasFiltroProp + 1, 6).createFilter();
}
