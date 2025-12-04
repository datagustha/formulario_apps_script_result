// CONFIGURAÇÕES
const CONFIG = {
  ID_PLANILHA: "1V4iGN14UpIQcwf3qKU0_Wbiy2exdW2WUmrYTniy0upA",
  ABA_PRINCIPAL: "Result",
  TIMEZONE: "America/Sao_Paulo"
};

// 🔥 ESTRUTURA DAS COLUNAS - CONFORME SUA PLANILHA ATUALIZADA
const COLUNAS = {
  RAZAO_SOCIAL: 0,
  NOME_FANTASIA: 1,
  CNPJ: 2,
  FORNECEDOR: 3,
  ULTIMA_ETAPA: 4,
  ETAPA: 5,
  OBSERVACAO: 6,
  CONTRATO_ENVIADO: 7,
  CONTRATO_ASSINADO: 8,
  ATIVACAO: 9,
  LINK: 10,
  MENSALIDADE: 11,
  MENSALIDADE_SIM: 12,
  MDR: 13,
  TIS: 14,
  REBATE: 15,
  ADESAO: 16,
  SITUACAO: 17
};

// 🔥 CONFIGURAÇÕES DOS WAITLABELS
const WAITLABELS_CONFIG = {
  WAITLABELS: ['Sim_Facilita', 'Result', 'Set_9', 'Doktorbank', 'Dr_Parcela'],
  WAITLABEL_PADRAO: 'Sim_Facilita',
  CORES: {
    'Sim_Facilita': '#7E3E9A',
    'Result': '#2EBE76',
    'Set_9': '#0682c5',
    'Doktorbank': '#E61B72',
    'Dr_Parcela': '#696969'
  }
};

function corrigirDadosExistentes() {
  const waitlabel = 'Sim_Facilita';
  const sheet = getSheetByName(waitlabel);
  const ultimaLinha = sheet.getLastRow();
  
  console.log("🔄 Corrigindo dados existentes...");
  
  if (ultimaLinha < 2) return { success: false, message: "Nenhum dado" };
  
  let correcoes = 0;
  
  for (let linha = 2; linha <= ultimaLinha; linha++) {
    const valorMDR = sheet.getRange(linha, COLUNAS.MDR + 1).getValue();
    const valorTIS = sheet.getRange(linha, COLUNAS.TIS + 1).getValue();
    const valorRebate = sheet.getRange(linha, COLUNAS.REBATE + 1).getValue();
    
    // 🔥 CORREÇÃO: Se o valor foi dividido por 100 (está decimal), multiplica de volta
    if (typeof valorMDR === 'number') {
      // Se for menor que 1 (ex: 0.98), mantém porque é 0,98%
      // Se for entre 0.01 e 0.99, está correto (é decimal percentual)
      // Não faz nada - mantém como está
    }
    
    if (typeof valorTIS === 'number') {
      // Mantém como está
    }
    
    if (typeof valorRebate === 'number') {
      // Mantém como está
    }
  }
  
  sheet.getRange(2, COLUNAS.MDR + 1, ultimaLinha - 1, 3).setNumberFormat('0.00"%"');
  SpreadsheetApp.flush();
  
  return { success: true, message: `✅ Formato corrigido para 2 casas decimais!`, correcoes: correcoes };
}

function formatarPercentualParaExibicao(valor) {
  if (valor === null || valor === undefined || valor === '') {
    return '';
  }
  
  console.log("🔧 GS->Front: formatarPercentualParaExibicao recebeu:", valor, "tipo:", typeof valor);
  
  try {
    if (typeof valor === 'string' && valor.includes('%')) {
      return valor;
    }
    
    if (typeof valor === 'number') {
      // 🔥 CORREÇÃO: NÃO multiplica por 100 - usa o valor direto
      // Se é 0.98, mostra 0,98%
      const formatado = valor.toFixed(2).replace('.', ',') + '%';
      console.log("✅ GS->Front: Número formatado:", valor, "→", formatado);
      return formatado;
    }
    
    return String(valor || '');
    
  } catch (error) {
    console.error("❌ GS->Front: Erro em formatarPercentualParaExibicao:", error);
    return String(valor || '');
  }
}

function corrigirParaDuasCasasDecimais() {
  const waitlabel = 'Sim_Facilita';
  const sheet = getSheetByName(waitlabel);
  const ultimaLinha = sheet.getLastRow();
  
  console.log("🔧 CORRIGINDO PARA 2 CASAS DECIMAIS...");
  
  if (ultimaLinha < 2) return { success: false, message: "Nenhum dado" };
  
  let correcoes = 0;
  
  for (let linha = 2; linha <= ultimaLinha; linha++) {
    // Obtém o valor que APARECE na célula
    const mdrDisplay = sheet.getRange(linha, COLUNAS.MDR + 1).getDisplayValue();
    const tisDisplay = sheet.getRange(linha, COLUNAS.TIS + 1).getDisplayValue();
    const rebateDisplay = sheet.getRange(linha, COLUNAS.REBATE + 1).getDisplayValue();
    
    console.log(`Linha ${linha} ANTES: MDR=${mdrDisplay}, TIS=${tisDisplay}, Rebate=${rebateDisplay}`);
    
    // Função para converter mantendo 2 casas decimais
    const converterParaDuasCasas = (displayValue) => {
      if (!displayValue || displayValue.trim() === '' || displayValue === '0,00%') {
        return '';
      }
      
      // Remove o % e troca vírgula por ponto
      const limpo = displayValue.replace('%', '').replace(',', '.');
      const numero = parseFloat(limpo);
      
      if (isNaN(numero)) {
        return '';
      }
      
      // Arredonda para 2 casas decimais
      return parseFloat(numero.toFixed(2));
    };
    
    // Converte cada valor
    const mdrCorrigido = converterParaDuasCasas(mdrDisplay);
    const tisCorrigido = converterParaDuasCasas(tisDisplay);
    const rebateCorrigido = converterParaDuasCasas(rebateDisplay);
    
    // Atualiza as células se houve alteração
    if (mdrCorrigido !== '') {
      sheet.getRange(linha, COLUNAS.MDR + 1).setValue(mdrCorrigido);
      correcoes++;
      console.log(`  MDR: ${mdrDisplay} → ${mdrCorrigido}`);
    }
    
    if (tisCorrigido !== '') {
      sheet.getRange(linha, COLUNAS.TIS + 1).setValue(tisCorrigido);
      correcoes++;
      console.log(`  TIS: ${tisDisplay} → ${tisCorrigido}`);
    }
    
    if (rebateCorrigido !== '') {
      sheet.getRange(linha, COLUNAS.REBATE + 1).setValue(rebateCorrigido);
      correcoes++;
      console.log(`  Rebate: ${rebateDisplay} → ${rebateCorrigido}`);
    }
  }
  
  // 🔥 FORMATO CRÍTICO: Use este formato EXATO
  // #,##0.00 - mostra número com 2 casas decimais
  // "%" - adiciona o símbolo de percentual
  sheet.getRange(2, COLUNAS.MDR + 1, ultimaLinha - 1, 3).setNumberFormat('#,##0.00"%"');
  
  SpreadsheetApp.flush();
  
  return { 
    success: true, 
    message: `✅ ${correcoes} valores corrigidos para 2 casas decimais!`,
    correcoes: correcoes
  };
}

function verificarValoresAtuais() {
  const waitlabel = 'Sim_Facilita';
  const sheet = getSheetByName(waitlabel);
  const ultimaLinha = Math.min(10, sheet.getLastRow()); // Verifica só as primeiras 10 linhas
  
  console.log("🔍 VERIFICANDO VALORES ATUAIS (primeiras 10 linhas):");
  console.log("=================================================");
  
  for (let linha = 2; linha <= ultimaLinha; linha++) {
    const mdrValor = sheet.getRange(linha, COLUNAS.MDR + 1).getValue();
    const mdrDisplay = sheet.getRange(linha, COLUNAS.MDR + 1).getDisplayValue();
    const mdrFormula = sheet.getRange(linha, COLUNAS.MDR + 1).getFormula();
    
    const tisValor = sheet.getRange(linha, COLUNAS.TIS + 1).getValue();
    const tisDisplay = sheet.getRange(linha, COLUNAS.TIS + 1).getDisplayValue();
    
    const rebateValor = sheet.getRange(linha, COLUNAS.REBATE + 1).getValue();
    const rebateDisplay = sheet.getRange(linha, COLUNAS.REBATE + 1).getDisplayValue();
    
    console.log(`Linha ${linha}:`);
    console.log(`  MDR - Valor: ${mdrValor}, Display: ${mdrDisplay}, Fórmula: ${mdrFormula || 'Nenhuma'}`);
    console.log(`  TIS - Valor: ${tisValor}, Display: ${tisDisplay}`);
    console.log(`  Rebate - Valor: ${rebateValor}, Display: ${rebateDisplay}`);
    console.log("---");
  }
  
  // Verifica também o formato atual
  const formatoMDR = sheet.getRange(2, COLUNAS.MDR + 1).getNumberFormat();
  console.log(`📋 Formato atual da coluna MDR: ${formatoMDR}`);
  
  return { success: true };
}
function formatarPercentualParaSalvar(percentual) {
  if (percentual === null || percentual === undefined || percentual === '') {
    return 0; // 🔥 ALTERADO: Retorna 0 em vez de string vazia
  }
  
  console.log("💾 Front->GS: formatarPercentualParaSalvar recebeu:", percentual, "tipo:", typeof percentual);
  
  try {
    if (typeof percentual === 'string') {
      const limpo = percentual.trim();
      if (limpo === '') return 0; // 🔥 Adicionado
      
      const numeroStr = limpo.replace('%', '').replace(',', '.');
      const numero = parseFloat(numeroStr);
      
      if (!isNaN(numero)) {
        // 🔥 CORREÇÃO: NÃO divide por 100 - salva o valor direto
        return parseFloat(numero.toFixed(4));
      } else {
        return 0; // 🔥 Se não for número válido, retorna 0
      }
    }
    
    if (typeof percentual === 'number') {
      // 🔥 CORREÇÃO: NÃO divide por 100
      return parseFloat(percentual.toFixed(4));
    }
    
    return 0; // 🔥 ALTERADO: Retorna 0 por padrão
    
  } catch (error) {
    console.error("❌ Front->GS: Erro em formatarPercentualParaSalvar:", error);
    return 0; // 🔥 ALTERADO: Retorna 0 em caso de erro
  }
}

// 🔥 FUNÇÃO PARA GARANTIR VALORES PADRÃO NOS PERCENTUAIS
function garantirValoresPadraoPercentuais() {
  const waitlabelAtual = getWaitlabelAtual();
  const sheet = getSheetByName(waitlabelAtual);
  const ultimaLinha = sheet.getLastRow();
  
  if (ultimaLinha < 2) return { success: true, message: "Nenhum registro" };
  
  let correcoes = 0;
  
  const dados = sheet.getDataRange().getValues();
  
  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    
    // Verificar coluna MDR (coluna 14 - índice 13)
    if (linha[COLUNAS.MDR] === '' || linha[COLUNAS.MDR] === null || linha[COLUNAS.MDR] === undefined) {
      sheet.getRange(i + 1, COLUNAS.MDR + 1).setValue(0);
      correcoes++;
    }
    
    // Verificar coluna TIS (coluna 15 - índice 14)
    if (linha[COLUNAS.TIS] === '' || linha[COLUNAS.TIS] === null || linha[COLUNAS.TIS] === undefined) {
      sheet.getRange(i + 1, COLUNAS.TIS + 1).setValue(0);
      correcoes++;
    }
    
    // Verificar coluna Rebate (coluna 16 - índice 15)
    if (linha[COLUNAS.REBATE] === '' || linha[COLUNAS.REBATE] === null || linha[COLUNAS.REBATE] === undefined) {
      sheet.getRange(i + 1, COLUNAS.REBATE + 1).setValue(0);
      correcoes++;
    }
  }
  
  // Aplicar formato correto
  sheet.getRange(2, COLUNAS.MDR + 1, ultimaLinha - 1, 3).setNumberFormat('0.00"%"');
  SpreadsheetApp.flush();
  
  return { 
    success: true, 
    message: `✅ ${correcoes} células definidas como 0%!`,
    correcoes: correcoes
  };
}

function resetarValoresPercentuais() {
  const waitlabel = 'Sim_Facilita';
  const sheet = getSheetByName(waitlabel);
  const ultimaLinha = sheet.getLastRow();
  
  console.log("🔄 RESETANDO VALORES PERCENTUAIS...");
  
  if (ultimaLinha < 2) return { success: false, message: "Nenhum dado" };
  
  const dados = sheet.getRange(2, 1, ultimaLinha - 1, 18).getValues();
  let correcoes = 0;
  
  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    if (!linha[0]) continue;
    
    const linhaReal = i + 2;
    
    const mdrExibido = sheet.getRange(linhaReal, COLUNAS.MDR + 1).getDisplayValue();
    const tisExibido = sheet.getRange(linhaReal, COLUNAS.TIS + 1).getDisplayValue();
    const rebateExibido = sheet.getRange(linhaReal, COLUNAS.REBATE + 1).getDisplayValue();
    
    if (mdrExibido && mdrExibido !== '') {
      const mdrLimpo = mdrExibido.replace('%', '').replace(',', '.');
      const mdrNumero = parseFloat(mdrLimpo);
      if (!isNaN(mdrNumero)) {
        // 🔥 CORREÇÃO: Salva o valor direto, sem dividir por 100
        sheet.getRange(linhaReal, COLUNAS.MDR + 1).setValue(mdrNumero);
        correcoes++;
      }
    }
    
    if (tisExibido && tisExibido !== '') {
      const tisLimpo = tisExibido.replace('%', '').replace(',', '.');
      const tisNumero = parseFloat(tisLimpo);
      if (!isNaN(tisNumero)) {
        sheet.getRange(linhaReal, COLUNAS.TIS + 1).setValue(tisNumero);
        correcoes++;
      }
    }
    
    if (rebateExibido && rebateExibido !== '') {
      const rebateLimpo = rebateExibido.replace('%', '').replace(',', '.');
      const rebateNumero = parseFloat(rebateLimpo);
      if (!isNaN(rebateNumero)) {
        sheet.getRange(linhaReal, COLUNAS.REBATE + 1).setValue(rebateNumero);
        correcoes++;
      }
    }
  }
  
  sheet.getRange(2, COLUNAS.MDR + 1, ultimaLinha - 1, 3).setNumberFormat('0.00"%"');
  SpreadsheetApp.flush();
  
  return { success: true, message: `✅ ${correcoes} valores resetados!`, correcoes: correcoes };
}

function getSheetByName(nome) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const sheet = ss.getSheetByName(nome);
    
    if (!sheet) {
      throw new Error('Planilha "' + nome + '" não encontrada!');
    }
    
    return sheet;
  } catch (error) {
    console.error("❌ Erro em getSheetByName:", error);
    throw error;
  }
}

function formatarDataBrasilSimples() {
  const agora = new Date();
  
  const dia = String(agora.getDate()).padStart(2, '0');
  const mes = String(agora.getMonth() + 1).padStart(2, '0');
  const ano = agora.getFullYear();
  const horas = String(agora.getHours()).padStart(2, '0');
  const minutos = String(agora.getMinutes()).padStart(2, '0');
  const segundos = String(agora.getSeconds()).padStart(2, '0');
  
  const dataFormatada = `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
  
  return dataFormatada;
}

function normalizarTexto(texto) {
  if (!texto || typeof texto !== 'string') return texto;
  return texto
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim();
}

function converterMoedaParaNumero(valorMoeda) {
  if (!valorMoeda) return 0;
  try {
    if (typeof valorMoeda === 'number') return valorMoeda;
    if (typeof valorMoeda === 'string') {
      const valorLimpo = valorMoeda
        .replace('R$', '')
        .replace(/\./g, '')
        .replace(',', '.')
        .trim();
      const numero = parseFloat(valorLimpo);
      return isNaN(numero) ? 0 : numero;
    }
    return parseFloat(valorMoeda) || 0;
  } catch (error) {
    console.error("❌ Erro ao converter moeda:", valorMoeda, error);
    return 0;
  }
}

function processarAdesao(valorAdesao) {
  if (!valorAdesao && valorAdesao !== 0) return 'Isento';
  const valorStr = valorAdesao.toString().trim();
  if (valorStr === 'Isento' || valorStr === '0' || valorStr === '0.00' || valorStr === 'R$ 0,00') {
    return 'Isento';
  }
  const numero = parseFloat(valorStr);
  if (!isNaN(numero)) {
    return numero;
  }
  return valorStr;
}

function processarAdesaoParaSalvar(valorAdesao) {
  if (!valorAdesao && valorAdesao !== 0) return 0;
  if (typeof valorAdesao === 'number') return valorAdesao;
  
  const valorStr = valorAdesao.toString().trim();
  if (valorStr === 'Isento' || valorStr === '0' || valorStr === '0.00' || valorStr === 'R$ 0,00') {
    return 0;
  }
  
  try {
    const valorLimpo = valorStr
      .replace('R$', '')
      .replace(/\./g, '')
      .replace(',', '.')
      .trim();
    
    const numero = parseFloat(valorLimpo);
    return isNaN(numero) ? 0 : numero;
    
  } catch (error) {
    console.error("Erro ao processar adesão:", error);
    return 0;
  }
}

function formatarCNPJNoSheets(cnpj) {
  if (!cnpj) return '';
  if (cnpj.toString().includes('.') || cnpj.toString().includes('/') || cnpj.toString().includes('-')) {
    return cnpj.toString();
  }
  const cnpjStr = cnpj.toString().replace(/\D/g, '');
  if (cnpjStr.length === 14) {
    return cnpjStr.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
  }
  return cnpj;
}

function validarEtapa(etapa, situacao) {
  const situacaoNormalizada = normalizarTexto(situacao || '');
  const precisaValidarEtapa = situacaoNormalizada === 'EM ANDAMENTO' || situacaoNormalizada === 'NOVO REGISTRO';
  
  if (!precisaValidarEtapa) {
    return { valida: true, etapa: etapa ? normalizarTexto(etapa) : '' };
  }
  
  if (!etapa || etapa.trim() === '') {
    return { 
      valida: false, 
      mensagem: `❌ Para situações "${situacao}" o campo Etapa é obrigatório!` 
    };
  }
  
  const etapasValidas = [
    "PENDENTE FORNECEDOR(ES)",
    "PENDENTE SIM", 
    "PENDENTE RETORNO EXTERNO"
  ];
  
  const etapaNormalizada = normalizarTexto(etapa);
  
  const etapasBloqueadas = ["DESISTIU", "REJEITADO", "CADASTRADO", "NOVO REGISTRO", "EM ANDAMENTO", "DESCREDENCIADO"];
  
  if (etapasBloqueadas.includes(etapaNormalizada)) {
    const mensagemErro = `❌ ETAPA NÃO PERMITIDA!\n\nA etapa "${etapa}" é uma SITUAÇÃO, não uma etapa do processo.\n\n📋 ETAPAS VÁLIDAS (do processo):\n• PENDENTE FORNECEDOR(ES)\n• PENDENTE SIM\n• PENDENTE RETORNO EXTERNO\n\n💡 Use o campo "Situação" para: ${etapa}`;
    
    return { valida: false, mensagem: mensagemErro };
  }
  
  if (!etapasValidas.includes(etapaNormalizada)) {
    const mensagemErro = `❌ ETAPA INVÁLIDA!\n\nA etapa "${etapa}" não é válida.\n\n📋 ETAPAS VÁLIDAS:\n• PENDENTE FORNECEDOR(ES)\n• PENDENTE SIM\n• PENDENTE RETORNO EXTERNO\n\nSelecione uma das etapas acima para continuar.`;
    
    return { valida: false, mensagem: mensagemErro };
  }
  
  return { valida: true, etapa: etapaNormalizada };
}

function processarCadastroComWaitlabel(dados, waitlabel) {
  try {
    console.log("🎯 PROCESSAR CADASTRO - INICIANDO");
    
    const sheet = getSheetByName(waitlabel);
    
    if (dados.acao === 'atualizar' && dados.id) {
      console.log("✏️ MODO ATUALIZAÇÃO");
      
      const linhaAtualizar = parseInt(dados.id);
      if (linhaAtualizar < 2 || linhaAtualizar > sheet.getLastRow()) {
        return { success: false, message: "Registro não encontrado" };
      }

      const dadosAtuais = sheet.getRange(linhaAtualizar, 1, 1, 18).getValues()[0];
      
      let houveAlteracaoRelevante = false;
      
      if (dados.etapa && dados.etapa !== dadosAtuais[COLUNAS.ETAPA]) {
        houveAlteracaoRelevante = true;
      }
      
      if (dados.situacao && dados.situacao !== dadosAtuais[COLUNAS.SITUACAO]) {
        houveAlteracaoRelevante = true;
      }
      
      if (dados.observacoes && dados.observacoes !== dadosAtuais[COLUNAS.OBSERVACAO]) {
        houveAlteracaoRelevante = true;
      }
      
      const novosDados = Array(18).fill('');
      
      for (let i = 0; i < 18; i++) {
        novosDados[i] = dadosAtuais[i] || '';
      }
      
      if (dados.razao_social) novosDados[COLUNAS.RAZAO_SOCIAL] = normalizarTexto(dados.razao_social);
      if (dados.nome_fantasia) novosDados[COLUNAS.NOME_FANTASIA] = normalizarTexto(dados.nome_fantasia);
      if (dados.cnpj) novosDados[COLUNAS.CNPJ] = dados.cnpj.toString();
      
      if (dados.fornecedores && dados.fornecedores.length > 0) {
        const primeiroFornecedor = dados.fornecedores[0];
        novosDados[COLUNAS.FORNECEDOR] = normalizarTexto(primeiroFornecedor.nome || primeiroFornecedor);
        
        // 🔥 CORREÇÃO: Se vazio ou indefinido, salvar como 0
        novosDados[COLUNAS.MDR] = (primeiroFornecedor.mdr !== undefined && primeiroFornecedor.mdr !== null && primeiroFornecedor.mdr !== '') 
            ? parseFloat(primeiroFornecedor.mdr.toString().replace('%', '').replace(',', '.')) 
            : 0;
            
        novosDados[COLUNAS.TIS] = (primeiroFornecedor.tis !== undefined && primeiroFornecedor.tis !== null && primeiroFornecedor.tis !== '') 
            ? parseFloat(primeiroFornecedor.tis.toString().replace('%', '').replace(',', '.')) 
            : 0;
            
        novosDados[COLUNAS.REBATE] = (primeiroFornecedor.rebate !== undefined && primeiroFornecedor.rebate !== null && primeiroFornecedor.rebate !== '') 
            ? parseFloat(primeiroFornecedor.rebate.toString().replace('%', '').replace(',', '.')) 
            : 0;
      }
      
      if (dados.etapa) novosDados[COLUNAS.ETAPA] = normalizarTexto(dados.etapa);
      if (dados.observacoes !== undefined) novosDados[COLUNAS.OBSERVACAO] = normalizarTexto(dados.observacoes);
      if (dados.contrato_enviado) novosDados[COLUNAS.CONTRATO_ENVIADO] = normalizarTexto(dados.contrato_enviado);
      if (dados.contrato_assinado) novosDados[COLUNAS.CONTRATO_ASSINADO] = normalizarTexto(dados.contrato_assinado);
      if (dados.ativacao) novosDados[COLUNAS.ATIVACAO] = dados.ativacao;
      if (dados.link) novosDados[COLUNAS.LINK] = dados.link;
      if (dados.mensalidade !== undefined) novosDados[COLUNAS.MENSALIDADE] = converterMoedaParaNumero(dados.mensalidade);
      if (dados.mensalidade_sim !== undefined) novosDados[COLUNAS.MENSALIDADE_SIM] = converterMoedaParaNumero(dados.mensalidade_sim);
      if (dados.adesao !== undefined) novosDados[COLUNAS.ADESAO] = processarAdesaoParaSalvar(dados.adesao);
      if (dados.situacao) novosDados[COLUNAS.SITUACAO] = normalizarTexto(dados.situacao);
      
      if (houveAlteracaoRelevante) {
        const novaData = formatarDataBrasilSimples();
        novosDados[COLUNAS.ULTIMA_ETAPA] = novaData;
      }
      
      sheet.getRange(linhaAtualizar, 1, 1, novosDados.length).setValues([novosDados]);
      
      sheet.getRange(linhaAtualizar, COLUNAS.MENSALIDADE + 1).setNumberFormat('"R$"#,##0.00');
      sheet.getRange(linhaAtualizar, COLUNAS.MENSALIDADE_SIM + 1).setNumberFormat('"R$"#,##0.00');
      sheet.getRange(linhaAtualizar, COLUNAS.ADESAO + 1).setNumberFormat('"R$"#,##0.00');
      sheet.getRange(linhaAtualizar, COLUNAS.MDR + 1).setNumberFormat('0.00"%"');
      sheet.getRange(linhaAtualizar, COLUNAS.TIS + 1).setNumberFormat('0.00"%"');
      sheet.getRange(linhaAtualizar, COLUNAS.REBATE + 1).setNumberFormat('0.00"%"');
      
      SpreadsheetApp.flush();
      
      return {
        success: true,
        message: '✅ Cadastro atualizado com sucesso!' + (houveAlteracaoRelevante ? ' (Data da última etapa atualizada)' : '')
      };
      
    } else {
      console.log("🆕 MODO NOVO CADASTRO");
      
      if (!dados.fornecedores || dados.fornecedores.length === 0) {
        return { success: false, message: '❌ Nenhum fornecedor selecionado!' };
      }
      
      const ultimaLinha = sheet.getLastRow();
      let linhaInserir = Math.max(2, ultimaLinha + 1);
      let registrosCriados = 0;
      
      const validacaoEtapa = validarEtapa(dados.etapa, dados.situacao);
      if (!validacaoEtapa.valida) {
        return { success: false, message: validacaoEtapa.mensagem };
      }
      const etapaValidada = validacaoEtapa.etapa;
      
      const cadastrosExistentes = buscarTodosCadastrosPorCNPJComWaitlabel(dados.cnpj, waitlabel);
      const fornecedoresDuplicados = [];
      
      for (let fornecedor of dados.fornecedores) {
        const nomeFornecedor = fornecedor.nome || fornecedor;
        const jaExiste = cadastrosExistentes.some(cad => cad.fornecedor === nomeFornecedor);
        if (jaExiste) {
          fornecedoresDuplicados.push(nomeFornecedor);
        }
      }
      
      if (fornecedoresDuplicados.length > 0) {
        return { 
          success: false, 
          message: `❌ Este CNPJ já possui cadastro no ${waitlabel} para: ${fornecedoresDuplicados.join(', ')}` 
        };
      }
      
      for (let i = 0; i < dados.fornecedores.length; i++) {
        const fornecedor = dados.fornecedores[i];
        const nomeFornecedor = fornecedor.nome || fornecedor;
        
        const dataAtual = formatarDataBrasilSimples();
        
        const linhaDados = [
          normalizarTexto(dados.razao_social) || '',
          normalizarTexto(dados.nome_fantasia) || '',
          dados.cnpj ? dados.cnpj.toString() : '',
          normalizarTexto(nomeFornecedor),
          dataAtual,
          etapaValidada,
          normalizarTexto(dados.observacoes) || '',
          normalizarTexto(dados.contrato_enviado) || '',
          normalizarTexto(dados.contrato_assinado) || '',
          dados.ativacao || '',
          dados.link || '',
          converterMoedaParaNumero(dados.mensalidade) || 0,
          converterMoedaParaNumero(dados.mensalidade_sim) || 0,
          // 🔥 CORREÇÃO: Se vazio ou indefinido, salvar como 0
          (fornecedor.mdr !== undefined && fornecedor.mdr !== null && fornecedor.mdr !== '') 
              ? parseFloat(fornecedor.mdr.toString().replace('%', '').replace(',', '.')) 
              : 0,
          (fornecedor.tis !== undefined && fornecedor.tis !== null && fornecedor.tis !== '') 
              ? parseFloat(fornecedor.tis.toString().replace('%', '').replace(',', '.')) 
              : 0,
          (fornecedor.rebate !== undefined && fornecedor.rebate !== null && fornecedor.rebate !== '') 
              ? parseFloat(fornecedor.rebate.toString().replace('%', '').replace(',', '.')) 
              : 0,
          processarAdesaoParaSalvar(dados.adesao),
          normalizarTexto(dados.situacao) || 'NOVO REGISTRO'
        ];
        
        try {
          sheet.getRange(linhaInserir, 1, 1, linhaDados.length).setValues([linhaDados]);
          
          sheet.getRange(linhaInserir, COLUNAS.MENSALIDADE + 1).setNumberFormat('"R$"#,##0.00');
          sheet.getRange(linhaInserir, COLUNAS.MENSALIDADE_SIM + 1).setNumberFormat('"R$"#,##0.00');
          sheet.getRange(linhaInserir, COLUNAS.ADESAO + 1).setNumberFormat('"R$"#,##0.00');
          sheet.getRange(linhaInserir, COLUNAS.MDR + 1).setNumberFormat('0.00"%"');
          sheet.getRange(linhaInserir, COLUNAS.TIS + 1).setNumberFormat('0.00"%"');
          sheet.getRange(linhaInserir, COLUNAS.REBATE + 1).setNumberFormat('0.00"%"');
          
          linhaInserir++;
          registrosCriados++;
        } catch (erro) {
          console.error(`❌ Erro ao salvar fornecedor ${nomeFornecedor}:`, erro);
        }
      }
      
      SpreadsheetApp.flush();
      
      return {
        success: true,
        message: `✅ "${dados.razao_social}" cadastrado com sucesso no ${waitlabel} para ${registrosCriados} fornecedor(es)!`,
        registrosCriados: registrosCriados
      };
    }
    
  } catch (error) {
    console.error("❌ Erro em processarCadastroComWaitlabel:", error);
    return {
      success: false,
      message: '❌ Erro: ' + error.toString()
    };
  }
}

function aplicarAlteracoesATodos(cnpj, dados, camposSelecionados) {
  try {
    console.log("🎯 APLICAR A TODOS - INICIANDO");
    
    const waitlabelAtual = getWaitlabelAtual();
    const sheet = getSheetByName(waitlabelAtual);
    
    const dadosCompletos = sheet.getDataRange().getValues();
    const registrosParaAtualizar = [];
    
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    
    for (let i = 1; i < dadosCompletos.length; i++) {
      const linha = dadosCompletos[i];
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      const cnpjLinha = linha[COLUNAS.CNPJ]?.toString().replace(/\D/g, '') || '';
      
      if (cnpjLinha === cnpjBuscado) {
        registrosParaAtualizar.push({
          linhaNumero: i + 1,
          dadosOriginais: linha
        });
      }
    }
    
    console.log(`🔍 Encontrados ${registrosParaAtualizar.length} registros`);
    
    const mudouEtapa = camposSelecionados.includes('etapa') || camposSelecionados.includes('inputEtapaSearch');
    const mudouSituacao = camposSelecionados.includes('situacao');
    const mudouObservacao = camposSelecionados.includes('observacoes');
    
    const precisaAtualizarData = mudouEtapa || mudouSituacao || mudouObservacao;
    
    let registrosAtualizados = 0;
    
    for (const registro of registrosParaAtualizar) {
      const novosDados = [...registro.dadosOriginais];
      
      for (const campo of camposSelecionados) {
        const valor = obterValorParaAplicarTodos(campo, dados);
        
        switch(campo) {
          case 'razao_social': novosDados[COLUNAS.RAZAO_SOCIAL] = valor; break;
          case 'nome_fantasia': novosDados[COLUNAS.NOME_FANTASIA] = valor; break;
          case 'cnpj_cadastro': novosDados[COLUNAS.CNPJ] = valor; break;
          case 'etapa':
          case 'inputEtapaSearch': novosDados[COLUNAS.ETAPA] = valor; break;
          case 'situacao': novosDados[COLUNAS.SITUACAO] = valor; break;
          case 'observacoes': novosDados[COLUNAS.OBSERVACAO] = valor; break;
          case 'contrato_enviado': novosDados[COLUNAS.CONTRATO_ENVIADO] = valor; break;
          case 'contrato_assinado': novosDados[COLUNAS.CONTRATO_ASSINADO] = valor; break;
          case 'ativacao': novosDados[COLUNAS.ATIVACAO] = valor; break;
          case 'link': novosDados[COLUNAS.LINK] = valor; break;
          case 'mensalidade': novosDados[COLUNAS.MENSALIDADE] = valor; break;
          case 'mensalidade_sim': novosDados[COLUNAS.MENSALIDADE_SIM] = valor; break;
          case 'adesao': novosDados[COLUNAS.ADESAO] = valor; break;
          case 'mdr': novosDados[COLUNAS.MDR] = valor; break;
          case 'tis': novosDados[COLUNAS.TIS] = valor; break;
          case 'rebate': novosDados[COLUNAS.REBATE] = valor; break;
        }
      }
      
      if (precisaAtualizarData) {
        const novaData = formatarDataBrasilSimples();
        novosDados[COLUNAS.ULTIMA_ETAPA] = novaData;
      }
      
      sheet.getRange(registro.linhaNumero, 1, 1, novosDados.length).setValues([novosDados]);
      
      sheet.getRange(registro.linhaNumero, COLUNAS.MENSALIDADE + 1).setNumberFormat('"R$"#,##0.00');
      sheet.getRange(registro.linhaNumero, COLUNAS.MENSALIDADE_SIM + 1).setNumberFormat('"R$"#,##0.00');
      sheet.getRange(registro.linhaNumero, COLUNAS.ADESAO + 1).setNumberFormat('"R$"#,##0.00');
      sheet.getRange(registro.linhaNumero, COLUNAS.MDR + 1).setNumberFormat('0.00"%"');
      sheet.getRange(registro.linhaNumero, COLUNAS.TIS + 1).setNumberFormat('0.00"%"');
      sheet.getRange(registro.linhaNumero, COLUNAS.REBATE + 1).setNumberFormat('0.00"%"');
      
      registrosAtualizados++;
    }
    
    SpreadsheetApp.flush();
    
    return {
      success: true,
      registrosAtualizados: registrosAtualizados,
      message: `✅ ${registrosAtualizados} registro(s) atualizado(s) com sucesso!` +
               (precisaAtualizarData ? ' (Data da última etapa atualizada)' : '')
    };
    
  } catch (error) {
    console.error("❌ ERRO em aplicarAlteracoesATodos:", error);
    return { success: false, message: "❌ Erro: " + error.toString() };
  }
}

function obterValorParaAplicarTodos(campo, dados) {
  switch(campo) {
    case 'razao_social': return normalizarTexto(dados.razao_social) || '';
    case 'nome_fantasia': return normalizarTexto(dados.nome_fantasia) || '';
    case 'cnpj_cadastro': return dados.cnpj ? dados.cnpj.toString() : '';
    case 'etapa':
    case 'inputEtapaSearch': return normalizarTexto(dados.etapa) || '';
    case 'observacoes': return normalizarTexto(dados.observacoes) || '';
    case 'contrato_enviado': return normalizarTexto(dados.contrato_enviado) || '';
    case 'contrato_assinado': return normalizarTexto(dados.contrato_assinado) || '';
    case 'ativacao': return dados.ativacao || '';
    case 'link': return dados.link || '';
    case 'mensalidade': return converterMoedaParaNumero(dados.mensalidade) || 0;
    case 'mensalidade_sim': return converterMoedaParaNumero(dados.mensalidade_sim) || 0;
    case 'adesao': return processarAdesaoParaSalvar(dados.adesao);
    // 🔥 CORREÇÃO: Se vazio ou indefinido, retornar 0
    case 'mdr': 
        return (dados.mdr !== undefined && dados.mdr !== null && dados.mdr !== '') 
            ? parseFloat(dados.mdr.toString().replace('%', '').replace(',', '.')) 
            : 0;
    case 'tis': 
        return (dados.tis !== undefined && dados.tis !== null && dados.tis !== '') 
            ? parseFloat(dados.tis.toString().replace('%', '').replace(',', '.')) 
            : 0;
    case 'rebate': 
        return (dados.rebate !== undefined && dados.rebate !== null && dados.rebate !== '') 
            ? parseFloat(dados.rebate.toString().replace('%', '').replace(',', '.')) 
            : 0;
    case 'situacao': 
      let situacao = normalizarTexto(dados.situacao) || 'NOVO REGISTRO';
      if (situacao === 'NOVO REGISTRO') situacao = 'Novo Registro';
      return situacao;
    default: return '';
  }
}

// 🔥 FUNÇÕES DE BUSCA CORRIGIDAS
function buscarTodosCadastrosComWaitlabel(waitlabel) {
  try {
    const sheet = getSheetByName(waitlabel);
    const ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 2) return [];
    
    const dados = sheet.getRange(2, 1, ultimaLinha - 1, 18).getValues();
    const cadastros = [];
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      if (!linha[0] || linha[0].toString().trim() === '') continue;

      let ultimaEtapaFormatada = '';
      if (linha[COLUNAS.ULTIMA_ETAPA] instanceof Date) {
        const dataUTC = linha[COLUNAS.ULTIMA_ETAPA];
        const dataBrasilia = new Date(dataUTC.getTime() - (5 * 60 * 60 * 1000));
        
        const dia = String(dataBrasilia.getDate()).padStart(2, '0');
        const mes = String(dataBrasilia.getMonth() + 1).padStart(2, '0');
        const ano = dataBrasilia.getFullYear();
        const horas = String(dataBrasilia.getHours()).padStart(2, '0');
        const minutos = String(dataBrasilia.getMinutes()).padStart(2, '0');
        const segundos = String(dataBrasilia.getSeconds()).padStart(2, '0');
        ultimaEtapaFormatada = `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
      } else {
        ultimaEtapaFormatada = linha[COLUNAS.ULTIMA_ETAPA]?.toString().trim() || '';
      }
      
      const cadastro = {
        id: i + 2,
        razao_social: linha[COLUNAS.RAZAO_SOCIAL]?.toString().trim() || '',
        nome_fantasia: linha[COLUNAS.NOME_FANTASIA]?.toString().trim() || '',
        cnpj: formatarCNPJNoSheets(linha[COLUNAS.CNPJ]?.toString().trim() || ''),
        fornecedor: linha[COLUNAS.FORNECEDOR]?.toString().trim() || '',
        ultima_etapa: ultimaEtapaFormatada,
        etapa: linha[COLUNAS.ETAPA]?.toString().trim() || '',
        observacoes: linha[COLUNAS.OBSERVACAO]?.toString().trim() || '',
        contrato_enviado: linha[COLUNAS.CONTRATO_ENVIADO]?.toString().trim() || '',
        contrato_assinado: linha[COLUNAS.CONTRATO_ASSINADO]?.toString().trim() || '',
        ativacao: linha[COLUNAS.ATIVACAO]?.toString().trim() || '',
        link: linha[COLUNAS.LINK]?.toString().trim() || '',
        mensalidade: parseFloat(linha[COLUNAS.MENSALIDADE]) || 0,
        mensalidade_sim: parseFloat(linha[COLUNAS.MENSALIDADE_SIM]) || 0,
        mdr: formatarPercentualParaExibicao(linha[COLUNAS.MDR]),
        tis: formatarPercentualParaExibicao(linha[COLUNAS.TIS]),
        rebate: formatarPercentualParaExibicao(linha[COLUNAS.REBATE]),
        adesao: processarAdesao(linha[COLUNAS.ADESAO]),
        situacao: (linha[COLUNAS.SITUACAO]?.toString().trim() || 'Novo registro'),
        waitlabel: waitlabel
      };
      
      cadastros.push(cadastro);
    }
    
    return cadastros;
    
  } catch (error) {
    console.error("❌ Erro em buscarTodosCadastrosComWaitlabel:", error);
    return [];
  }
}

function buscarTodosCadastrosPorCNPJComWaitlabel(cnpj, waitlabel) {
  try {
    const sheet = getSheetByName(waitlabel);
    const ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 2) return [];
    
    const dados = sheet.getRange(2, 1, ultimaLinha - 1, 18).getValues();
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    const cadastrosEncontrados = [];
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      const cnpjCadastro = linha[COLUNAS.CNPJ]?.toString().replace(/\D/g, '') || '';
      
      if (cnpjCadastro === cnpjBuscado) {
        let ultimaEtapaFormatada = '';
        if (linha[COLUNAS.ULTIMA_ETAPA] instanceof Date) {
          const dataUTC = linha[COLUNAS.ULTIMA_ETAPA];
          const dataBrasilia = new Date(dataUTC.getTime() - (5 * 60 * 60 * 1000));
          
          const dia = String(dataBrasilia.getDate()).padStart(2, '0');
          const mes = String(dataBrasilia.getMonth() + 1).padStart(2, '0');
          const ano = dataBrasilia.getFullYear();
          const horas = String(dataBrasilia.getHours()).padStart(2, '0');
          const minutos = String(dataBrasilia.getMinutes()).padStart(2, '0');
          const segundos = String(dataBrasilia.getSeconds()).padStart(2, '0');
          ultimaEtapaFormatada = `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
        } else {
          ultimaEtapaFormatada = linha[COLUNAS.ULTIMA_ETAPA]?.toString().trim() || '';
        }
        
        const cadastro = {
          id: i + 2,
          razao_social: linha[COLUNAS.RAZAO_SOCIAL]?.toString().trim() || '',
          nome_fantasia: linha[COLUNAS.NOME_FANTASIA]?.toString().trim() || '',
          cnpj: formatarCNPJNoSheets(linha[COLUNAS.CNPJ]?.toString().trim() || ''),
          fornecedor: linha[COLUNAS.FORNECEDOR]?.toString().trim() || '',
          ultima_etapa: ultimaEtapaFormatada,
          etapa: linha[COLUNAS.ETAPA]?.toString().trim() || '',
          observacoes: linha[COLUNAS.OBSERVACAO]?.toString().trim() || '',
          contrato_enviado: linha[COLUNAS.CONTRATO_ENVIADO]?.toString().trim() || '',
          contrato_assinado: linha[COLUNAS.CONTRATO_ASSINADO]?.toString().trim() || '',
          ativacao: linha[COLUNAS.ATIVACAO]?.toString().trim() || '',
          link: linha[COLUNAS.LINK]?.toString().trim() || '',
          mensalidade: parseFloat(linha[COLUNAS.MENSALIDADE]) || 0,
          mensalidade_sim: parseFloat(linha[COLUNAS.MENSALIDADE_SIM]) || 0,
          mdr: formatarPercentualParaExibicao(linha[COLUNAS.MDR]),
          tis: formatarPercentualParaExibicao(linha[COLUNAS.TIS]),
          rebate: formatarPercentualParaExibicao(linha[COLUNAS.REBATE]),
          adesao: processarAdesao(linha[COLUNAS.ADESAO]),
          situacao: (linha[COLUNAS.SITUACAO]?.toString().trim() || 'Novo registro'),
          waitlabel: waitlabel
        };
        
        cadastrosEncontrados.push(cadastro);
      }
    }
    
    return cadastrosEncontrados;
    
  } catch (error) {
    console.error("❌ Erro em buscarTodosCadastrosPorCNPJComWaitlabel:", error);
    return [];
  }
}

function buscarCadastroPorIDComWaitlabel(id, waitlabel) {
  try {
    const sheet = getSheetByName(waitlabel);
    const ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < id) return { encontrado: false, mensagem: "Registro não encontrado" };
    
    const linha = sheet.getRange(id, 1, 1, 18).getValues()[0];
    
    if (!linha[0] || linha[0].toString().trim() === '') {
      return { encontrado: false, mensagem: "Registro vazio ou não encontrado" };
    }

    let ultimaEtapaFormatada = '';
    if (linha[COLUNAS.ULTIMA_ETAPA] instanceof Date) {
      const dataUTC = linha[COLUNAS.ULTIMA_ETAPA];
      const dataBrasilia = new Date(dataUTC.getTime() - (5 * 60 * 60 * 1000));
      
      const dia = String(dataBrasilia.getDate()).padStart(2, '0');
      const mes = String(dataBrasilia.getMonth() + 1).padStart(2, '0');
      const ano = dataBrasilia.getFullYear();
      const horas = String(dataBrasilia.getHours()).padStart(2, '0');
      const minutos = String(dataBrasilia.getMinutes()).padStart(2, '0');
      const segundos = String(dataBrasilia.getSeconds()).padStart(2, '0');
      ultimaEtapaFormatada = `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
    } else {
      ultimaEtapaFormatada = linha[COLUNAS.ULTIMA_ETAPA]?.toString().trim() || '';
    }

    const fornecedorParaFormulario = {
      nome: linha[COLUNAS.FORNECEDOR]?.toString().trim() || '',
      mdr: formatarPercentualParaExibicao(linha[COLUNAS.MDR]),
      tis: formatarPercentualParaExibicao(linha[COLUNAS.TIS]),
      rebate: formatarPercentualParaExibicao(linha[COLUNAS.REBATE])
    };
    
    const resultado = {
      encontrado: true,
      id: id,
      razao_social: linha[COLUNAS.RAZAO_SOCIAL]?.toString().trim() || '',
      nome_fantasia: linha[COLUNAS.NOME_FANTASIA]?.toString().trim() || '',
      cnpj: formatarCNPJNoSheets(linha[COLUNAS.CNPJ]?.toString().trim() || ''),
      fornecedor: linha[COLUNAS.FORNECEDOR]?.toString().trim() || '',
      fornecedores: [fornecedorParaFormulario],
      ultima_etapa: ultimaEtapaFormatada,
      etapa: linha[COLUNAS.ETAPA]?.toString().trim() || '',
      observacoes: linha[COLUNAS.OBSERVACAO]?.toString().trim() || '',
      contrato_enviado: linha[COLUNAS.CONTRATO_ENVIADO]?.toString().trim() || '',
      contrato_assinado: linha[COLUNAS.CONTRATO_ASSINADO]?.toString().trim() || '',
      ativacao: linha[COLUNAS.ATIVACAO]?.toString().trim() || '',
      link: linha[COLUNAS.LINK]?.toString().trim() || '',
      mensalidade: parseFloat(linha[COLUNAS.MENSALIDADE]) || 0,
      mensalidade_sim: parseFloat(linha[COLUNAS.MENSALIDADE_SIM]) || 0,
      mdr: fornecedorParaFormulario.mdr,
      tis: fornecedorParaFormulario.tis,
      rebate: fornecedorParaFormulario.rebate,
      adesao: processarAdesao(linha[COLUNAS.ADESAO]),
      situacao: (linha[COLUNAS.SITUACAO]?.toString().trim() || 'Novo registro'),
      waitlabel: waitlabel
    };

    return resultado;
    
  } catch (error) {
    console.error("❌ Erro em buscarCadastroPorIDComWaitlabel:", error);
    return { encontrado: false, mensagem: "Erro: " + error.message };
  }
}

// 🔥 FUNÇÕES DE WAITLABELS
function getWaitlabelAtual() {
  const cache = CacheService.getScriptCache();
  const waitlabelAtual = cache.get('waitlabel_atual');
  return waitlabelAtual || WAITLABELS_CONFIG.WAITLABEL_PADRAO;
}

function setWaitlabelAtual(waitlabel) {
  if (WAITLABELS_CONFIG.WAITLABELS.includes(waitlabel)) {
    const cache = CacheService.getScriptCache();
    cache.put('waitlabel_atual', waitlabel, 21600);
    return { success: true, message: `Waitlabel alterado para: ${waitlabel}` };
  }
  return { success: false, message: 'Waitlabel inválido' };
}

function getCoresWaitlabels() {
  return WAITLABELS_CONFIG.CORES;
}

function getWaitlabels() {
  return WAITLABELS_CONFIG.WAITLABELS;
}

// 🔥 FUNÇÃO PRINCIPAL DO WEB APP
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistema - Gestão de Cadastros')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 🔥 ADICIONE ESTA FUNÇÃO NO FINAL DO ARQUIVO GS
function corrigirTodosPercentuais() {
  const waitlabel = 'Sim_Facilita';
  const sheet = getSheetByName(waitlabel);
  const ultimaLinha = sheet.getLastRow();
  
  console.log("🔄 CORRIGINDO TODOS OS PERCENTUAIS...");
  
  if (ultimaLinha < 2) return { success: true, message: "Nenhum dado para corrigir" };
  
  let correcoes = 0;
  
  for (let linha = 2; linha <= ultimaLinha; linha++) {
    const mdrRange = sheet.getRange(linha, COLUNAS.MDR + 1);
    const tisRange = sheet.getRange(linha, COLUNAS.TIS + 1);
    const rebateRange = sheet.getRange(linha, COLUNAS.REBATE + 1);
    
    const mdrValor = mdrRange.getValue();
    const tisValor = tisRange.getValue();
    const rebateValor = rebateRange.getValue();
    
    // Se o valor for muito grande (ex: 1112), divide por 100
    if (typeof mdrValor === 'number' && mdrValor > 10) {
      mdrRange.setValue(mdrValor / 100);
      correcoes++;
      console.log(`Linha ${linha}: MDR ${mdrValor} → ${mdrValor/100}`);
    }
    
    if (typeof tisValor === 'number' && tisValor > 10) {
      tisRange.setValue(tisValor / 100);
      correcoes++;
      console.log(`Linha ${linha}: TIS ${tisValor} → ${tisValor/100}`);
    }
    
    if (typeof rebateValor === 'number' && rebateValor > 10) {
      rebateRange.setValue(rebateValor / 100);
      correcoes++;
      console.log(`Linha ${linha}: Rebate ${rebateValor} → ${rebateValor/100}`);
    }
  }
  
  // Aplica formato correto para TODAS as células
  sheet.getRange(2, COLUNAS.MDR + 1, ultimaLinha - 1, 3).setNumberFormat('0.00"%"');
  
  SpreadsheetApp.flush();
  
  return { 
    success: true, 
    message: `✅ ${correcoes} valores de percentual corrigidos! Formato: 0.00%`,
    correcoes: correcoes
  };
}

// 🔥 FUNÇÃO DE TESTE
function testar() {
  return { 
    success: true, 
    message: "✅ Sistema funcionando!",
    timestamp: new Date().toISOString()
  };
}
