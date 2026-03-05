// CONFIGURAÇÕES
const CONFIG = {
  ID_PLANILHA: "1V4iGN14UpIQcwf3qKU0_Wbiy2exdW2WUmrYTniy0upA",
  TIMEZONE: "America/Sao_Paulo"
};

// 🔥 ESTRUTURA DAS COLUNAS - ATUALIZADA COM NOVAS COLUNAS
const COLUNAS_PADRAO = {
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
  DATA_CRIACAO: 17,
  SITUACAO: 18
};

const COLUNAS_SIM_FACILITA = {
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
  PLANO: 11,
  MENSALIDADE: 12,
  VENC: 13,
  METODO_PGTO: 14,
  MDR: 15,
  TIS: 16,
  REBATE: 17,
  ADESAO: 18,
  PGTO_ADESAO: 19,
  DATA_CRIACAO: 20,
  TREINADO: 21,
  SITUACAO: 22
};

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

// 🔥 FUNÇÃO AUXILIAR PARA OBTER ESTRUTURA CORRETA
function getColunasConfig(waitlabel) {
  return waitlabel === 'Sim_Facilita' ? COLUNAS_SIM_FACILITA : COLUNAS_PADRAO;
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

// 🔥 CORREÇÃO: Função para converter porcentagem para número decimal
function converterPercentualParaDecimal(valorPercentual) {
  if (!valorPercentual && valorPercentual !== 0) return 0;
  
  console.log("🔢 converterPercentualParaDecimal recebeu:", valorPercentual, "tipo:", typeof valorPercentual);
  
  try {
    if (typeof valorPercentual === 'number') {
      console.log("📊 É número:", valorPercentual);
      
      if (valorPercentual >= 1) {
        console.log("📊 Número >= 1, dividindo por 100:", valorPercentual, "→", valorPercentual / 100);
        return valorPercentual / 100;
      } else {
        console.log("📊 Número < 1, já é decimal:", valorPercentual);
        return valorPercentual;
      }
    }
    
    if (typeof valorPercentual === 'string') {
      console.log("📊 É string:", valorPercentual);
      
      const valorSemPercentual = valorPercentual.replace('%', '').trim();
      const valorComPonto = valorSemPercentual.replace(',', '.');
      
      const partes = valorComPonto.split('.');
      let valorFinalStr = '';
      
      if (partes.length > 1) {
        if (partes[1].length > 0) {
          valorFinalStr = partes[0].replace(/\./g, '') + '.' + partes[1];
        } else {
          valorFinalStr = partes[0].replace(/\./g, '');
        }
      } else {
        valorFinalStr = valorComPonto.replace(/\./g, '');
      }
      
      const numero = parseFloat(valorFinalStr);
      
      if (isNaN(numero)) {
        console.log("⚠️ Não é número válido:", valorPercentual);
        return 0;
      }
      
      console.log("📊 Número convertido da string:", numero);
      
      if (numero >= 1) {
        console.log("📊 Número >= 1, dividindo por 100:", numero, "→", numero / 100);
        return numero / 100;
      } else {
        console.log("📊 Número < 1, já é decimal:", numero);
        return numero;
      }
    }
    
    console.log("⚠️ Tipo não reconhecido, retornando 0");
    return 0;
    
  } catch (error) {
    console.error("❌ Erro ao converter percentual:", error);
    return 0;
  }
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

function formatarCNPJParaSalvar(cnpj) {
  if (!cnpj) return '';
  
  console.log("🔢 formatarCNPJParaSalvar recebeu:", cnpj, "tipo:", typeof cnpj);
  
  try {
    if (typeof cnpj === 'string' && (cnpj.includes('.') || cnpj.includes('/') || cnpj.includes('-'))) {
      console.log("✅ CNPJ já formatado:", cnpj);
      return cnpj;
    }
    
    const cnpjStr = cnpj.toString();
    const cnpjLimpo = cnpjStr.replace(/\D/g, '');
    console.log("🔢 CNPJ limpo:", cnpjLimpo);
    
    if (cnpjLimpo.length === 14) {
      const cnpjFormatado = cnpjLimpo.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
      console.log("✅ CNPJ formatado:", cnpjFormatado);
      return cnpjFormatado;
    }
    
    console.log("⚠️ CNPJ com formato incompleto:", cnpjLimpo);
    return cnpjLimpo;
    
  } catch (error) {
    console.error("❌ Erro ao formatar CNPJ:", error);
    return cnpj ? cnpj.toString() : '';
  }
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

// 🔥 FUNÇÃO CORRIGIDA: formatarPercentualParaExibicao
function formatarPercentualParaExibicao(valor) {
  if (valor === null || valor === undefined || valor === '') {
    return '';
  }
  
  try {
    console.log("📊 formatarPercentualParaExibicao recebeu:", valor, "tipo:", typeof valor);
    
    if (typeof valor === 'string' && valor.includes('%')) {
      console.log("✅ Já tem %, retornando:", valor);
      return valor;
    }
    
    let numero;
    
    if (typeof valor === 'string') {
      const valorLimpo = valor.replace('%', '').replace(',', '.');
      numero = parseFloat(valorLimpo);
      
      if (isNaN(numero)) {
        console.log("⚠️ String não é número válido:", valor);
        return valor;
      }
    } else if (typeof valor === 'number') {
      numero = valor;
    } else {
      console.log("⚠️ Tipo não suportado:", typeof valor);
      return String(valor);
    }
    
    let valorParaExibir;
    
    if (numero < 1) {
      valorParaExibir = numero * 100;
      console.log("📊 Decimal <1 convertido para exibição:", numero, "→", valorParaExibir);
    } else if (numero >= 1 && numero <= 100) {
      valorParaExibir = numero;
      console.log("📊 Já está em percentual, mantendo:", valorParaExibir);
    } else {
      valorParaExibir = numero;
      console.log("📊 Número grande, mantendo como está:", valorParaExibir);
    }
    
    const formatado = valorParaExibir.toFixed(2).replace('.', ',') + '%';
    console.log("✅ Formatado para exibição:", formatado);
    
    return formatado;
    
  } catch (error) {
    console.error("❌ Erro em formatarPercentualParaExibicao:", error);
    return String(valor || '');
  }
}

function processarTreinado(valorTreinado) {
  console.log("🎯 processarTreinado recebeu:", valorTreinado, "tipo:", typeof valorTreinado);
  
  // 🔥 CORREÇÃO: Se for undefined, null ou string vazia, retorna 'NAO'
  if (valorTreinado === undefined || valorTreinado === null) {
    console.log("⚠️ Treinado undefined/null, retornando 'NAO'");
    return 'NAO';
  }
  
  // Se for string vazia
  if (typeof valorTreinado === 'string' && valorTreinado.trim() === '') {
    console.log("⚠️ Treinado string vazia, retornando 'NAO'");
    return 'NAO';
  }
  
  // Se for string
  if (typeof valorTreinado === 'string') {
    const valorUpper = valorTreinado.trim().toUpperCase();
    
    // 🔥 ACEITAR MÚLTIPLAS FORMAS DE "SIM"
    if (valorUpper === 'SIM' || valorUpper === 'S' || 
        valorUpper === 'YES' || valorUpper === 'Y' || 
        valorUpper === 'TRUE' || valorUpper === 'V' ||
        valorUpper === 'VERDADEIRO' || valorUpper === '1') {
      console.log("✅ String reconhecida como SIM");
      return 'SIM';
    }
    
    // 🔥 ACEITAR MÚLTIPLAS FORMAS DE "NAO"
    if (valorUpper === 'NÃO' || valorUpper === 'NAO' || 
        valorUpper === 'N' || valorUpper === 'NO' ||
        valorUpper === 'FALSE' || valorUpper === 'F' ||
        valorUpper === 'FALSO' || valorUpper === '0') {
      console.log("✅ String reconhecida como NAO");
      return 'NAO';
    }
    
    // Se não reconhecer, assume como NAO
    console.log("⚠️ String não reconhecida, retornando NAO");
    return 'NAO';
  }
  
  // Se for booleano
  if (typeof valorTreinado === 'boolean') {
    console.log("✅ Boolean:", valorTreinado);
    return valorTreinado ? 'SIM' : 'NAO';
  }
  
  // Se for número
  if (typeof valorTreinado === 'number') {
    console.log("✅ Number:", valorTreinado);
    return valorTreinado > 0 ? 'SIM' : 'NAO';
  }
  
  // Para qualquer outro tipo
  try {
    const valorStr = String(valorTreinado).trim().toUpperCase();
    if (valorStr === 'SIM' || valorStr === 'S' || valorStr === 'YES' || 
        valorStr === 'Y' || valorStr === 'TRUE' || valorStr === '1') {
      return 'SIM';
    }
  } catch (e) {
    console.error("❌ Erro ao processar treinado:", e);
  }
  
  // Padrão seguro
  console.log("⚠️ Tipo não reconhecido, retornando NAO");
  return 'NAO';
}

// 🔥 ATUALIZE esta função para:
function obterTreinadoPorFornecedor(dados, nomeFornecedor) {
  const fornecedorLower = nomeFornecedor.toLowerCase();
  
  console.log(`\n🔍 [obterTreinadoPorFornecedor] Buscando treinado para: ${nomeFornecedor} (${fornecedorLower})`);
  console.log("📋 Dados recebidos:", Object.keys(dados).filter(k => k.includes('treinado')).map(k => `${k}: "${dados[k]}"`));
  
  // 🔥 PRIMEIRO: Verificar se existe um campo específico (treinado_agil, treinado_bc, etc)
  const campoEspecifico = `treinado_${fornecedorLower}`;
  if (dados[campoEspecifico] !== undefined && 
      dados[campoEspecifico] !== null && 
      dados[campoEspecifico] !== '') {
    
    const resultado = processarTreinado(dados[campoEspecifico]);
    console.log(`✅ Usando campo específico "${campoEspecifico}": ${dados[campoEspecifico]} -> ${resultado}`);
    return resultado;
  }
  
  // 🔥 SEGUNDO: Verificar se há um campo "treinado" genérico (que é o que o formulário envia)
  if (dados.treinado !== undefined && dados.treinado !== null) {
    const resultado = processarTreinado(dados.treinado);
    console.log(`⚠️ Campo específico não encontrado, usando 'treinado' genérico: ${dados.treinado} -> ${resultado}`);
    return resultado;
  }
  
  // 🔥 TERCEIRO: Verificar campos com nomes similares
  const fornecedorSemEspacos = fornecedorLower.replace(/\s+/g, '');
  const possiveisCampos = Object.keys(dados).filter(key => 
    key.startsWith('treinado') && key.toLowerCase().includes(fornecedorSemEspacos)
  );
  
  if (possiveisCampos.length > 0) {
    const resultado = processarTreinado(dados[possiveisCampos[0]]);
    console.log(`✅ Usando campo similar "${possiveisCampos[0]}": ${dados[possiveisCampos[0]]} -> ${resultado}`);
    return resultado;
  }
  
  // 🔥 QUARTO: Default
  console.log(`❌ Nenhum treinado encontrado para ${nomeFornecedor}, usando 'NAO'`);
  return 'NAO';
}

// 🔥🔥🔥 FUNÇÃO COMPLETA CORRIGIDA PARA PROCESSAR CADASTRO COM NOVAS COLUNAS
function processarCadastroComWaitlabel(dados, waitlabel) {
  try {
    console.log("🎯 PROCESSAR CADASTRO - INICIANDO");
    console.log("Waitlabel:", waitlabel);
    console.log("Dados recebidos:", JSON.stringify(dados));
    
    // 🔥🔥🔥 DEBUG EXTENDIDO PARA TREINAMENTO
    console.log("🎓 DEBUG COMPLETO DOS CAMPOS DE TREINAMENTO:");
    const camposTreinado = Object.keys(dados).filter(k => k.includes('treinado'));
    if (camposTreinado.length > 0) {
      camposTreinado.forEach(k => {
        console.log(`  ${k}: "${dados[k]}" (tipo: ${typeof dados[k]})`);
      });
    } else {
      console.log("  ❌ NENHUM CAMPO DE TREINAMENTO ENCONTRADO!");
    }
    
    const sheet = getSheetByName(waitlabel);
    const COLUNAS = getColunasConfig(waitlabel);
    
    let fornecedoresArray = [];
    
    console.log("🔍 Analisando fornecedor:", dados.fornecedor);
    
    if (typeof dados.fornecedor === 'string') {
      fornecedoresArray = dados.fornecedor.split(',').map(f => f.trim()).filter(f => f);
      console.log("✅ Fornecedores processados de string:", fornecedoresArray);
    } else if (Array.isArray(dados.fornecedor)) {
      fornecedoresArray = dados.fornecedor;
      console.log("✅ Fornecedores processados de array:", fornecedoresArray);
    } else {
      console.log("⚠️ Tipo de fornecedor não reconhecido:", typeof dados.fornecedor);
      return { 
        success: false, 
        message: '❌ ERRO: Formato de fornecedor inválido. Selecione pelo menos um fornecedor.' 
      };
    }
    
    if (fornecedoresArray.length === 0) {
      return { 
        success: false, 
        message: '❌ NENHUM FORNECEDOR SELECIONADO! Selecione pelo menos um fornecedor.' 
      };
    }
    
    if (!dados.cnpj || dados.cnpj.toString().replace(/\D/g, '').length < 14) {
      return {
        success: false,
        message: '❌ CNPJ INVÁLIDO! Informe um CNPJ válido com 14 dígitos.'
      };
    }
    
    const cnpjFormatado = formatarCNPJParaSalvar(dados.cnpj);
    console.log("✅ CNPJ formatado para salvar:", cnpjFormatado);
    
    const situacaoNormalizada = normalizarTexto(dados.situacao || 'NOVO REGISTRO');
    let etapaValidada = normalizarTexto(dados.etapa || '');
    
    const situacoesComEtapaObrigatoria = ['EM ANDAMENTO', 'NOVO REGISTRO'];
    const situacaoEmAndamentoOuNovo = situacoesComEtapaObrigatoria.includes(situacaoNormalizada);
    
    if (!situacaoEmAndamentoOuNovo) {
      etapaValidada = situacaoNormalizada;
    } else {
      const validacaoEtapa = validarEtapa(dados.etapa, dados.situacao);
      if (!validacaoEtapa.valida) {
        return { success: false, message: validacaoEtapa.mensagem };
      }
      etapaValidada = validacaoEtapa.etapa;
    }
    
    // 🔥 NOVA COLUNA: DATA_CRIACAO - sempre salvar a data/hora atual
    const dataCriacao = formatarDataBrasilSimples();
    
    if (dados.acao === 'atualizar' && dados.id) {
      console.log("✏️ MODO ATUALIZAÇÃO - ID:", dados.id);
      
      const linhaAtualizar = parseInt(dados.id);
      if (linhaAtualizar < 2 || linhaAtualizar > sheet.getLastRow()) {
        return { success: false, message: "Registro não encontrado" };
      }

      const totalColunas = waitlabel === 'Sim_Facilita' ? 23 : 19;
      const dadosAtuais = sheet.getRange(linhaAtualizar, 1, 1, totalColunas).getValues()[0];
      
      let houveAlteracaoRelevante = false;
      
      if (dados.situacao && dados.situacao !== dadosAtuais[COLUNAS.SITUACAO]) {
        houveAlteracaoRelevante = true;
      }
      
      if (dados.etapa && dados.etapa !== dadosAtuais[COLUNAS.ETAPA]) {
        houveAlteracaoRelevante = true;
      }
      
      if (dados.observacoes && dados.observacoes !== dadosAtuais[COLUNAS.OBSERVACAO]) {
        houveAlteracaoRelevante = true;
      }
      
      const novosDados = Array(totalColunas).fill('');
      
      for (let i = 0; i < totalColunas; i++) {
        novosDados[i] = dadosAtuais[i] || '';
      }
      
      if (dados.razao_social) novosDados[COLUNAS.RAZAO_SOCIAL] = normalizarTexto(dados.razao_social);
      if (dados.nome_fantasia) novosDados[COLUNAS.NOME_FANTASIA] = normalizarTexto(dados.nome_fantasia);
      
      if (dados.cnpj) novosDados[COLUNAS.CNPJ] = cnpjFormatado;
      
      if (fornecedoresArray.length > 0) {
        const primeiroFornecedor = fornecedoresArray[0];
        novosDados[COLUNAS.FORNECEDOR] = normalizarTexto(primeiroFornecedor);
        
        const fornecedorLower = primeiroFornecedor.toLowerCase();
        
        const mdrVal = dados[`mdr_${fornecedorLower}`] !== undefined ? dados[`mdr_${fornecedorLower}`] : dados.mdr;
        const tisVal = dados[`tis_${fornecedorLower}`] !== undefined ? dados[`tis_${fornecedorLower}`] : dados.tis;
        const rebateVal = dados[`rebate_${fornecedorLower}`] !== undefined ? dados[`rebate_${fornecedorLower}`] : dados.rebate;
        
        console.log(`📊 Processando percentuais para ${fornecedorLower}:`);
        console.log("  MDR:", mdrVal, "TIS:", tisVal, "Rebate:", rebateVal);
        
        const mdrDecimal = converterPercentualParaDecimal(mdrVal);
        const tisDecimal = converterPercentualParaDecimal(tisVal);
        const rebateDecimal = converterPercentualParaDecimal(rebateVal);
        
        console.log(`✅ Percentuais convertidos para ${fornecedorLower}:`);
        console.log("  MDR decimal:", mdrDecimal);
        console.log("  TIS decimal:", tisDecimal);
        console.log("  Rebate decimal:", rebateDecimal);
        
        if (waitlabel === 'Sim_Facilita') {
          novosDados[COLUNAS.MDR] = mdrDecimal;
          novosDados[COLUNAS.TIS] = tisDecimal;
          novosDados[COLUNAS.REBATE] = rebateDecimal;
        } else {
          novosDados[COLUNAS_PADRAO.MDR] = mdrDecimal;
          novosDados[COLUNAS_PADRAO.TIS] = tisDecimal;
          novosDados[COLUNAS_PADRAO.REBATE] = rebateDecimal;
        }
      }
      
      novosDados[COLUNAS.ETAPA] = etapaValidada;
      
      if (dados.observacoes !== undefined) novosDados[COLUNAS.OBSERVACAO] = normalizarTexto(dados.observacoes);
      if (dados.contrato_enviado) novosDados[COLUNAS.CONTRATO_ENVIADO] = normalizarTexto(dados.contrato_enviado);
      if (dados.contrato_assinado) novosDados[COLUNAS.CONTRATO_ASSINADO] = normalizarTexto(dados.contrato_assinado);
      if (dados.ativacao) novosDados[COLUNAS.ATIVACAO] = dados.ativacao;
      if (dados.link) novosDados[COLUNAS.LINK] = dados.link;
      if (dados.situacao) novosDados[COLUNAS.SITUACAO] = normalizarTexto(dados.situacao);
      
      // 🔥 MANTER DATA_CRIACAO (não atualizar na edição)
      if (dadosAtuais[COLUNAS.DATA_CRIACAO]) {
        novosDados[COLUNAS.DATA_CRIACAO] = dadosAtuais[COLUNAS.DATA_CRIACAO];
      }
      
      // 🔥🔥🔥 CORREÇÃO CRÍTICA: ATUALIZAR COLUNA TREINADO APENAS PARA SIM_FACILITA
      if (waitlabel === 'Sim_Facilita') {
        console.log("🔥🔥🔥 [ATUALIZAÇÃO] Processando campo TREINADO para Sim_Facilita");
        console.log("   - Fornecedor atual dos dados:", dadosAtuais[COLUNAS.FORNECEDOR]);
        console.log("   - Fornecedor do formulário:", fornecedoresArray[0]);
        
        const fornecedorAtual = dadosAtuais[COLUNAS.FORNECEDOR]?.toString().trim() || '';
        let treinadoProcessado = 'NAO';
        
        if (fornecedorAtual) {
          console.log(`   🎯 Fornecedor encontrado na planilha: "${fornecedorAtual}"`);
          
          // 🔥 BUSCAR TODOS OS CAMPOS POSSÍVEIS DE TREINADO
          const fornecedorLower = fornecedorAtual.toLowerCase();
          const campoTreinadoEspecifico = `treinado_${fornecedorLower}`;
          
          // 🔥 PRIORIDADE 1: Campo específico do fornecedor
          if (dados[campoTreinadoEspecifico] !== undefined && dados[campoTreinadoEspecifico] !== null) {
            treinadoProcessado = processarTreinado(dados[campoTreinadoEspecifico]);
            console.log(`   ✅ Usando campo específico "${campoTreinadoEspecifico}": ${dados[campoTreinadoEspecifico]} -> ${treinadoProcessado}`);
          }
          // 🔥 PRIORIDADE 2: Campo 'treinado' genérico
          else if (dados.treinado !== undefined && dados.treinado !== null) {
            treinadoProcessado = processarTreinado(dados.treinado);
            console.log(`   ⚠️ Usando campo 'treinado' genérico: ${dados.treinado} -> ${treinadoProcessado}`);
          }
          // 🔥 PRIORIDADE 3: Valor antigo
          else if (dadosAtuais[COLUNAS.TREINADO] !== undefined && dadosAtuais[COLUNAS.TREINADO] !== null) {
            treinadoProcessado = processarTreinado(dadosAtuais[COLUNAS.TREINADO]);
            console.log(`   🔄 Mantendo valor antigo: ${dadosAtuais[COLUNAS.TREINADO]} -> ${treinadoProcessado}`);
          }
        } else {
          // Se não encontrar fornecedor na planilha, usar o primeiro fornecedor do formulário
          const primeiroFornecedorForm = fornecedoresArray[0].toLowerCase();
          const campoTreinadoEspecifico = `treinado_${primeiroFornecedorForm}`;
          
          if (dados[campoTreinadoEspecifico] !== undefined && dados[campoTreinadoEspecifico] !== null) {
            treinadoProcessado = processarTreinado(dados[campoTreinadoEspecifico]);
            console.log(`   🔄 Usando campo específico do formulário "${campoTreinadoEspecifico}": ${dados[campoTreinadoEspecifico]} -> ${treinadoProcessado}`);
          } else if (dados.treinado !== undefined && dados.treinado !== null) {
            treinadoProcessado = processarTreinado(dados.treinado);
            console.log(`   🔄 Usando 'treinado' genérico do formulário: ${dados.treinado} -> ${treinadoProcessado}`);
          }
        }
        
        novosDados[COLUNAS.TREINADO] = treinadoProcessado;
        console.log("🔥🔥🔥 Treinado salvo como:", treinadoProcessado);
      }
      
      if (waitlabel === 'Sim_Facilita') {
        if (dados.plano !== undefined) novosDados[COLUNAS.PLANO] = normalizarTexto(dados.plano);
        if (dados.mensalidade !== undefined) novosDados[COLUNAS.MENSALIDADE] = converterMoedaParaNumero(dados.mensalidade);
        if (dados.vencimento !== undefined) novosDados[COLUNAS.VENC] = dados.vencimento;
        if (dados.metodo_pgto !== undefined) novosDados[COLUNAS.METODO_PGTO] = dados.metodo_pgto;
        if (dados.adesao !== undefined) novosDados[COLUNAS.ADESAO] = processarAdesaoParaSalvar(dados.adesao);
        if (dados.pgto_adesao !== undefined) novosDados[COLUNAS.PGTO_ADESAO] = dados.pgto_adesao;
      } else {
        if (dados.mensalidade !== undefined) novosDados[COLUNAS_PADRAO.MENSALIDADE] = converterMoedaParaNumero(dados.mensalidade);
        if (dados.mensalidade_sim !== undefined) novosDados[COLUNAS_PADRAO.MENSALIDADE_SIM] = converterMoedaParaNumero(dados.mensalidade_sim);
        if (dados.adesao !== undefined) novosDados[COLUNAS_PADRAO.ADESAO] = processarAdesaoParaSalvar(dados.adesao);
      }
      
      if (houveAlteracaoRelevante) {
        const novaData = formatarDataBrasilSimples();
        novosDados[COLUNAS.ULTIMA_ETAPA] = novaData;
      }
      
      if (!houveAlteracaoRelevante) {
        novosDados[COLUNAS.ULTIMA_ETAPA] = formatarDataBrasilSimples();
      }
      
      sheet.getRange(linhaAtualizar, 1, 1, novosDados.length).setValues([novosDados]);
      
      // 🔥 CORREÇÃO: Formatar células corretamente
      if (waitlabel === 'Sim_Facilita') {
        sheet.getRange(linhaAtualizar, COLUNAS.MENSALIDADE + 1).setNumberFormat('"R$"#,##0.00');
        sheet.getRange(linhaAtualizar, COLUNAS.ADESAO + 1).setNumberFormat('"R$"#,##0.00');
        if (dados.pgto_adesao) {
          sheet.getRange(linhaAtualizar, COLUNAS.PGTO_ADESAO + 1).setNumberFormat('dd/mm/yyyy');
        }
        // Formatar percentuais
        sheet.getRange(linhaAtualizar, COLUNAS.MDR + 1).setNumberFormat('0.00%');
        sheet.getRange(linhaAtualizar, COLUNAS.TIS + 1).setNumberFormat('0.00%');
        sheet.getRange(linhaAtualizar, COLUNAS.REBATE + 1).setNumberFormat('0.00%');
        // 🔥 Formatar DATA_CRIACAO
        if (dadosAtuais[COLUNAS.DATA_CRIACAO]) {
          sheet.getRange(linhaAtualizar, COLUNAS.DATA_CRIACAO + 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
        }
        // 🔥 Formatar TREINADO como texto
        sheet.getRange(linhaAtualizar, COLUNAS.TREINADO + 1).setNumberFormat('@');
      } else {
        sheet.getRange(linhaAtualizar, COLUNAS_PADRAO.MENSALIDADE + 1).setNumberFormat('"R$"#,##0.00');
        sheet.getRange(linhaAtualizar, COLUNAS_PADRAO.MENSALIDADE_SIM + 1).setNumberFormat('"R$"#,##0.00');
        sheet.getRange(linhaAtualizar, COLUNAS_PADRAO.ADESAO + 1).setNumberFormat('"R$"#,##0.00');
        // Formatar percentuais
        sheet.getRange(linhaAtualizar, COLUNAS_PADRAO.MDR + 1).setNumberFormat('0.00%');
        sheet.getRange(linhaAtualizar, COLUNAS_PADRAO.TIS + 1).setNumberFormat('0.00%');
        sheet.getRange(linhaAtualizar, COLUNAS_PADRAO.REBATE + 1).setNumberFormat('0.00%');
        // 🔥 Formatar DATA_CRIACAO
        if (dadosAtuais[COLUNAS_PADRAO.DATA_CRIACAO]) {
          sheet.getRange(linhaAtualizar, COLUNAS_PADRAO.DATA_CRIACAO + 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
        }
      }
      
      SpreadsheetApp.flush();
      
      return {
        success: true,
        message: '✅ Cadastro atualizado com sucesso!'
      };
      
    } else {
      console.log("🆕 MODO NOVO CADASTRO");
      
      const cadastrosExistentes = buscarTodosCadastrosPorCNPJComWaitlabel(dados.cnpj, waitlabel);
      const fornecedoresDuplicados = [];
      
      for (let fornecedor of fornecedoresArray) {
        const jaExiste = cadastrosExistentes.some(cad => cad.fornecedor === normalizarTexto(fornecedor));
        if (jaExiste) {
          fornecedoresDuplicados.push(fornecedor);
        }
      }
      
      if (fornecedoresDuplicados.length > 0) {
        return { 
          success: false, 
          message: `❌ Este CNPJ já possui cadastro no ${waitlabel} para: ${fornecedoresDuplicados.join(', ')}` 
        };
      }
      
      let linhaInserir = Math.max(2, sheet.getLastRow() + 1);
      let registrosCriados = 0;
      
      for (let i = 0; i < fornecedoresArray.length; i++) {
        const nomeFornecedor = fornecedoresArray[i];
        const fornecedorLower = nomeFornecedor.toLowerCase();
        
        const mdrVal = dados[`mdr_${fornecedorLower}`] !== undefined ? dados[`mdr_${fornecedorLower}`] : dados.mdr;
        const tisVal = dados[`tis_${fornecedorLower}`] !== undefined ? dados[`tis_${fornecedorLower}`] : dados.tis;
        const rebateVal = dados[`rebate_${fornecedorLower}`] !== undefined ? dados[`rebate_${fornecedorLower}`] : dados.rebate;
        
        console.log(`📊 Processando percentuais para ${fornecedorLower}:`);
        console.log("  MDR:", mdrVal, "TIS:", tisVal, "Rebate:", rebateVal);
        
        const mdrDecimal = converterPercentualParaDecimal(mdrVal);
        const tisDecimal = converterPercentualParaDecimal(tisVal);
        const rebateDecimal = converterPercentualParaDecimal(rebateVal);
        
        console.log(`✅ Percentuais convertidos para ${fornecedorLower}:`);
        console.log("  MDR decimal:", mdrDecimal);
        console.log("  TIS decimal:", tisDecimal);
        console.log("  Rebate decimal:", rebateDecimal);
        
        const dataAtual = formatarDataBrasilSimples();
        
        let linhaDados = [];
        
        if (waitlabel === 'Sim_Facilita') {
          // 🔥 CORREÇÃO CRÍTICA: Usar obterTreinadoPorFornecedor
          const treinadoProcessado = obterTreinadoPorFornecedor(dados, nomeFornecedor);
          console.log(`🔥🔥🔥 [NOVO CADASTRO] Treinado para ${nomeFornecedor}:`, {
            'dados.treinado': dados.treinado,
            [`dados.treinado_${fornecedorLower}`]: dados[`treinado_${fornecedorLower}`],
            'processado': treinadoProcessado
          });
          
          linhaDados = [
            normalizarTexto(dados.razao_social) || '',
            normalizarTexto(dados.nome_fantasia) || '',
            cnpjFormatado,
            normalizarTexto(nomeFornecedor),
            dataAtual, // ULTIMA_ETAPA
            etapaValidada,
            normalizarTexto(dados.observacoes) || '',
            normalizarTexto(dados.contrato_enviado) || '',
            normalizarTexto(dados.contrato_assinado) || '',
            dados.ativacao || '',
            dados.link || '',
            normalizarTexto(dados.plano) || '',
            converterMoedaParaNumero(dados.mensalidade) || 0,
            dados.vencimento || '',
            dados.metodo_pgto || '',
            mdrDecimal,
            tisDecimal,
            rebateDecimal,
            processarAdesaoParaSalvar(dados.adesao),
            dados.pgto_adesao || '',
            dataCriacao, // DATA_CRIACAO
            treinadoProcessado, // 🔥🔥🔥 CORREÇÃO AQUI!
            situacaoNormalizada
          ];
        } else {
          linhaDados = [
            normalizarTexto(dados.razao_social) || '',
            normalizarTexto(dados.nome_fantasia) || '',
            cnpjFormatado,
            normalizarTexto(nomeFornecedor),
            dataAtual, // ULTIMA_ETAPA
            etapaValidada,
            normalizarTexto(dados.observacoes) || '',
            normalizarTexto(dados.contrato_enviado) || '',
            normalizarTexto(dados.contrato_assinado) || '',
            dados.ativacao || '',
            dados.link || '',
            converterMoedaParaNumero(dados.mensalidade) || 0,
            converterMoedaParaNumero(dados.mensalidade_sim) || 0,
            mdrDecimal,
            tisDecimal,
            rebateDecimal,
            processarAdesaoParaSalvar(dados.adesao),
            dataCriacao, // DATA_CRIACAO
            situacaoNormalizada
          ];
        }
        
        try {
          console.log(`Inserindo linha ${linhaInserir} para fornecedor ${nomeFornecedor}`);
          console.log(`Dados do treinado na linha: ${linhaDados[COLUNAS.TREINADO] || 'N/A'}`);
          
          sheet.getRange(linhaInserir, 1, 1, linhaDados.length).setValues([linhaDados]);
          
          // 🔥 CORREÇÃO: Formatar células corretamente
          if (waitlabel === 'Sim_Facilita') {
            sheet.getRange(linhaInserir, COLUNAS.MENSALIDADE + 1).setNumberFormat('"R$"#,##0.00');
            sheet.getRange(linhaInserir, COLUNAS.ADESAO + 1).setNumberFormat('"R$"#,##0.00');
            if (dados.pgto_adesao) {
              sheet.getRange(linhaInserir, COLUNAS.PGTO_ADESAO + 1).setNumberFormat('dd/mm/yyyy');
            }
            // Formatar percentuais
            sheet.getRange(linhaInserir, COLUNAS.MDR + 1).setNumberFormat('0.00%');
            sheet.getRange(linhaInserir, COLUNAS.TIS + 1).setNumberFormat('0.00%');
            sheet.getRange(linhaInserir, COLUNAS.REBATE + 1).setNumberFormat('0.00%');
            // 🔥 Formatar DATA_CRIACAO
            sheet.getRange(linhaInserir, COLUNAS.DATA_CRIACAO + 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
            // 🔥 Formatar TREINADO como texto
            sheet.getRange(linhaInserir, COLUNAS.TREINADO + 1).setNumberFormat('@');
          } else {
            sheet.getRange(linhaInserir, COLUNAS_PADRAO.MENSALIDADE + 1).setNumberFormat('"R$"#,##0.00');
            sheet.getRange(linhaInserir, COLUNAS_PADRAO.MENSALIDADE_SIM + 1).setNumberFormat('"R$"#,##0.00');
            sheet.getRange(linhaInserir, COLUNAS_PADRAO.ADESAO + 1).setNumberFormat('"R$"#,##0.00');
            // Formatar percentuais
            sheet.getRange(linhaInserir, COLUNAS_PADRAO.MDR + 1).setNumberFormat('0.00%');
            sheet.getRange(linhaInserir, COLUNAS_PADRAO.TIS + 1).setNumberFormat('0.00%');
            sheet.getRange(linhaInserir, COLUNAS_PADRAO.REBATE + 1).setNumberFormat('0.00%');
            // 🔥 Formatar DATA_CRIACAO
            sheet.getRange(linhaInserir, COLUNAS_PADRAO.DATA_CRIACAO + 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
          }
          
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

// 🔥🔥🔥 FUNÇÃO ATUALIZADA PARA BUSCAR CADASTROS
function buscarTodosCadastrosComWaitlabel(waitlabel) {
  try {
    console.log(`🔍 Buscando cadastros para ${waitlabel}...`);
    
    const sheet = getSheetByName(waitlabel);
    if (!sheet) {
      console.log(`❌ Sheet ${waitlabel} não encontrado`);
      return [];
    }
    
    const ultimaLinha = sheet.getLastRow();
    console.log(`📊 Última linha na planilha ${waitlabel}: ${ultimaLinha}`);
    
    if (ultimaLinha < 2) {
      console.log(`ℹ️ Nenhum dado na planilha ${waitlabel}`);
      return [];
    }
    
    const COLUNAS = getColunasConfig(waitlabel);
    const totalColunas = waitlabel === 'Sim_Facilita' ? 23 : 19;
    
    const dados = sheet.getRange(2, 1, ultimaLinha - 1, totalColunas).getValues();
    console.log(`📋 ${dados.length} linhas de dados obtidas`);
    
    const cadastros = [];
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      
      if (!linha[0] || linha[0].toString().trim() === '') {
        continue;
      }

      let ultimaEtapaFormatada = '';
      if (linha[COLUNAS.ULTIMA_ETAPA] instanceof Date) {
        ultimaEtapaFormatada = formatarDataParaExibicao(linha[COLUNAS.ULTIMA_ETAPA]);
      } else {
        ultimaEtapaFormatada = linha[COLUNAS.ULTIMA_ETAPA]?.toString().trim() || '';
      }
      
      // 🔥 Formatar DATA_CRIACAO
      let dataCriacaoFormatada = '';
      if (linha[COLUNAS.DATA_CRIACAO] instanceof Date) {
        dataCriacaoFormatada = formatarDataParaExibicao(linha[COLUNAS.DATA_CRIACAO]);
      } else {
        dataCriacaoFormatada = linha[COLUNAS.DATA_CRIACAO]?.toString().trim() || '';
      }
      
      const cnpjDisplay = linha[COLUNAS.CNPJ] ? formatarCNPJParaExibicao(linha[COLUNAS.CNPJ]) : '';
      
      const cadastro = {
        id: i + 2,
        razao_social: linha[COLUNAS.RAZAO_SOCIAL]?.toString().trim() || '',
        nome_fantasia: linha[COLUNAS.NOME_FANTASIA]?.toString().trim() || '',
        cnpj: cnpjDisplay,
        fornecedor: linha[COLUNAS.FORNECEDOR]?.toString().trim() || '',
        ultima_etapa: ultimaEtapaFormatada,
        etapa: linha[COLUNAS.ETAPA]?.toString().trim() || '',
        observacoes: linha[COLUNAS.OBSERVACAO]?.toString().trim() || '',
        contrato_enviado: linha[COLUNAS.CONTRATO_ENVIADO]?.toString().trim() || '',
        contrato_assinado: linha[COLUNAS.CONTRATO_ASSINADO]?.toString().trim() || '',
        ativacao: linha[COLUNAS.ATIVACAO]?.toString().trim() || '',
        link: linha[COLUNAS.LINK]?.toString().trim() || '',
        situacao: (linha[COLUNAS.SITUACAO]?.toString().trim() || 'NOVO REGISTRO'),
        data_criacao: dataCriacaoFormatada,
        waitlabel: waitlabel
      };
      
      if (waitlabel === 'Sim_Facilita') {
        cadastro.plano = linha[COLUNAS.PLANO]?.toString().trim() || '';
        cadastro.mensalidade = parseFloat(linha[COLUNAS.MENSALIDADE]) || 0;
        cadastro.vencimento = linha[COLUNAS.VENC]?.toString().trim() || '';
        cadastro.metodo_pgto = linha[COLUNAS.METODO_PGTO]?.toString().trim() || '';
        cadastro.mdr = formatarPercentualParaExibicao(linha[COLUNAS.MDR]);
        cadastro.tis = formatarPercentualParaExibicao(linha[COLUNAS.TIS]);
        cadastro.rebate = formatarPercentualParaExibicao(linha[COLUNAS.REBATE]);
        cadastro.adesao = processarAdesao(linha[COLUNAS.ADESAO]);
        cadastro.pgto_adesao = linha[COLUNAS.PGTO_ADESAO]?.toString().trim() || '';
        cadastro.treinado = processarTreinado(linha[COLUNAS.TREINADO]); // 🔥 CORREÇÃO: usar função processarTreinado
      } else {
        cadastro.mensalidade = parseFloat(linha[COLUNAS_PADRAO.MENSALIDADE]) || 0;
        cadastro.mensalidade_sim = parseFloat(linha[COLUNAS_PADRAO.MENSALIDADE_SIM]) || 0;
        cadastro.mdr = formatarPercentualParaExibicao(linha[COLUNAS_PADRAO.MDR]);
        cadastro.tis = formatarPercentualParaExibicao(linha[COLUNAS_PADRAO.TIS]);
        cadastro.rebate = formatarPercentualParaExibicao(linha[COLUNAS_PADRAO.REBATE]);
        cadastro.adesao = processarAdesao(linha[COLUNAS_PADRAO.ADESAO]);
      }
      
      cadastros.push(cadastro);
    }
    
    console.log(`✅ ${cadastros.length} cadastros encontrados para ${waitlabel}`);
    return cadastros;
    
  } catch (error) {
    console.error("❌ Erro em buscarTodosCadastrosComWaitlabel:", error);
    return [];
  }
}

// 🔥 FUNÇÃO AUXILIAR PARA FORMATAR DATA
function formatarDataParaExibicao(data) {
  if (!data) return '';
  
  try {
    if (data instanceof Date) {
      const dia = String(data.getDate()).padStart(2, '0');
      const mes = String(data.getMonth() + 1).padStart(2, '0');
      const ano = data.getFullYear();
      const horas = String(data.getHours()).padStart(2, '0');
      const minutos = String(data.getMinutes()).padStart(2, '0');
      const segundos = String(data.getSeconds()).padStart(2, '0');
      return `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
    } else if (typeof data === 'string') {
      return data;
    }
    
    return String(data);
  } catch (error) {
    console.error("❌ Erro ao formatar data:", error);
    return String(data || '');
  }
}

function formatarCNPJParaExibicao(cnpj) {
  if (!cnpj) return '';
  
  try {
    const cnpjStr = cnpj.toString();
    
    if (cnpjStr.includes('.') || cnpjStr.includes('/') || cnpjStr.includes('-')) {
      return cnpjStr;
    }
    
    const cnpjLimpo = cnpjStr.replace(/\D/g, '');
    
    if (cnpjLimpo.length === 14) {
      return cnpjLimpo.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
    }
    
    return cnpjLimpo;
    
  } catch (error) {
    console.error("❌ Erro ao formatar CNPJ para exibição:", error);
    return cnpj ? cnpj.toString() : '';
  }
}

// 🔥🔥🔥 FUNÇÃO COMPLETA CORRIGIDA PARA APLICAR A TODOS
function aplicarAlteracoesATodos(cnpj, dados, campos, waitlabel) {
  try {
    console.log("🎯 APLICAR A TODOS - INICIANDO");
    console.log("CNPJ:", cnpj);
    console.log("Campos:", campos);
    console.log("Waitlabel recebido:", waitlabel); // 🔥 AGORA VEM DIRETO DO FRONTEND
    console.log("Dados recebidos:", JSON.stringify(dados));
    
    // 🔥🔥🔥 AGORA O WAITLABEL VEM DO FRONTEND, NÃO PRECISA BUSCAR
    // const waitlabel = dados.waitlabel || WAITLABELS_CONFIG.WAITLABEL_PADRAO; // ❌ REMOVA ESTA LINHA SE EXISTIR
    
    console.log("Waitlabel final:", waitlabel);
    
    // 🔥 VALIDAÇÃO EXTRA: Se não veio waitlabel, usa Result como padrão
    if (!waitlabel || waitlabel === '') {
      console.warn("⚠️ Waitlabel não recebido, usando 'Result' como padrão");
      waitlabel = 'Result';
    }
    
    const sheet = getSheetByName(waitlabel);
    if (!sheet) {
      return {
        success: false,
        message: `❌ Planilha "${waitlabel}" não encontrada! Verifique o nome da aba.`
      };
    }
    
    const COLUNAS = getColunasConfig(waitlabel);
    
    const cadastrosDoCNPJ = buscarTodosCadastrosPorCNPJComWaitlabel(cnpj, waitlabel);
    
    if (cadastrosDoCNPJ.length === 0) {
      return {
        success: false,
        message: `❌ Nenhum registro encontrado com o CNPJ ${cnpj} no ${waitlabel}`
      };
    }
    
    console.log(`🔍 Encontrados ${cadastrosDoCNPJ.length} registros para o CNPJ`);
    
    let registrosAtualizados = 0;
    
    const situacaoUpper = dados.situacao ? normalizarTexto(dados.situacao) : '';
    const ehSituacaoFinal = situacaoUpper === 'CADASTRADO' || 
                           situacaoUpper === 'REJEITADO' || 
                           situacaoUpper === 'DESCREDENCIADO' || 
                           situacaoUpper === 'DESISTIU';
    
    console.log(`🔥 Situação: ${situacaoUpper}, É situação final? ${ehSituacaoFinal}`);
    
    if (ehSituacaoFinal) {
      console.log(`🔥🔥🔥 SITUAÇÃO FINAL DETECTADA! Forçando etapa = situação: ${situacaoUpper}`);
      
      delete dados.etapa;
      delete dados.inputEtapaSearch;
      
      dados.etapa = situacaoUpper;
      dados.inputEtapaSearch = situacaoUpper;
      
      if (!campos.includes('etapa') && !campos.includes('inputEtapaSearch')) {
        campos.push('etapa');
        campos.push('inputEtapaSearch');
      }
      
      console.log(`🔥 Dados após correção: etapa=${dados.etapa}, inputEtapaSearch=${dados.inputEtapaSearch}`);
    }
    
    for (const cadastro of cadastrosDoCNPJ) {
      const linha = cadastro.id;
      const totalColunas = waitlabel === 'Sim_Facilita' ? 23 : 19;
      const dadosAtuais = sheet.getRange(linha, 1, 1, totalColunas).getValues()[0];
      
      const novosDados = Array(totalColunas).fill('');
      
      for (let i = 0; i < totalColunas; i++) {
        novosDados[i] = dadosAtuais[i] || '';
      }
      
      let houveAlteracao = false;
      
      for (const campo of campos) {
        console.log(`\n🔥 Processando campo: ${campo}, valor: ${dados[campo]}`);
        
        let colunaIndex = -1;
        let valorParaSalvar = dados[campo];
        
        if (campo === 'razao_social') {
          colunaIndex = COLUNAS.RAZAO_SOCIAL;
        } else if (campo === 'nome_fantasia') {
          colunaIndex = COLUNAS.NOME_FANTASIA;
        } else if (campo === 'contrato_enviado') {
          colunaIndex = COLUNAS.CONTRATO_ENVIADO;
        } else if (campo === 'contrato_assinado') {
          colunaIndex = COLUNAS.CONTRATO_ASSINADO;
        } else if (campo === 'ativacao') {
          colunaIndex = COLUNAS.ATIVACAO;
        } else if (campo === 'link') {
          colunaIndex = COLUNAS.LINK;
        } else if (campo === 'mensalidade') {
          colunaIndex = waitlabel === 'Sim_Facilita' ? COLUNAS.MENSALIDADE : COLUNAS_PADRAO.MENSALIDADE;
        } else if (campo === 'mensalidade_sim' && waitlabel !== 'Sim_Facilita') {
          colunaIndex = COLUNAS_PADRAO.MENSALIDADE_SIM;
        } else if (campo === 'situacao') {
          colunaIndex = COLUNAS.SITUACAO;
        } else if (campo === 'etapa' || campo === 'inputEtapaSearch') {
          colunaIndex = COLUNAS.ETAPA;
          console.log(`🔥🔥🔥 APLICANDO ETAPA: colunaIndex=${colunaIndex}, valor=${dados[campo]}`);
        } else if (campo === 'observacoes') {
          colunaIndex = COLUNAS.OBSERVACAO;
        } else if (campo === 'adesao') {
          colunaIndex = waitlabel === 'Sim_Facilita' ? COLUNAS.ADESAO : COLUNAS_PADRAO.ADESAO;
        } else if (campo === 'data_criacao') {
          colunaIndex = COLUNAS.DATA_CRIACAO;
          console.log(`🔥🔥🔥 APLICANDO DATA_CRIACAO: colunaIndex=${colunaIndex}, valor=${dados[campo]}`);
        } else if (campo === 'treinado' && waitlabel === 'Sim_Facilita') {
          colunaIndex = COLUNAS.TREINADO;
          console.log(`🔥🔥🔥 [APLICAR A TODOS] Campo TREINADO para ${waitlabel}`);
          
          const fornecedorAtual = dadosAtuais[COLUNAS.FORNECEDOR]?.toString().trim() || '';
          console.log(`   Fornecedor deste registro: ${fornecedorAtual}`);
          
          if (fornecedorAtual) {
            const fornecedorLower = fornecedorAtual.toLowerCase();
            const campoTreinadoEspecifico = `treinado_${fornecedorLower}`;
            
            if (dados[campoTreinadoEspecifico] !== undefined && dados[campoTreinadoEspecifico] !== null) {
              valorParaSalvar = processarTreinado(dados[campoTreinadoEspecifico]);
              console.log(`   ✅ Usando campo específico "${campoTreinadoEspecifico}": ${dados[campoTreinadoEspecifico]} -> ${valorParaSalvar}`);
            }
            else if (dados.treinado !== undefined && dados.treinado !== null) {
              valorParaSalvar = processarTreinado(dados.treinado);
              console.log(`   ⚠️ Usando campo 'treinado' genérico: ${dados.treinado} -> ${valorParaSalvar}`);
            }
            else {
              valorParaSalvar = processarTreinado(dadosAtuais[COLUNAS.TREINADO]);
              console.log(`   🔄 Mantendo valor atual: ${dadosAtuais[COLUNAS.TREINADO]} -> ${valorParaSalvar}`);
            }
          } else {
            valorParaSalvar = processarTreinado(dados[campo] || dados.treinado || 'NAO');
            console.log(`   Fornecedor não encontrado, usando valor: ${valorParaSalvar}`);
          }
        } else if (waitlabel === 'Sim_Facilita') {
          if (campo === 'plano') {
            colunaIndex = COLUNAS.PLANO;
            console.log(`🔥🔥🔥 APLICANDO PLANO: colunaIndex=${colunaIndex}, valor=${dados[campo]}`);
          } else if (campo === 'vencimento') {
            colunaIndex = COLUNAS.VENC;
          } else if (campo === 'metodo_pgto') {
            colunaIndex = COLUNAS.METODO_PGTO;
          } else if (campo === 'pgto_adesao') {
            colunaIndex = COLUNAS.PGTO_ADESAO;
          }
        }
        
        if (colunaIndex !== -1 && dados[campo] !== undefined) {
          
          if (campo === 'treinado' && waitlabel === 'Sim_Facilita') {
            // Já processado acima
          } else if (campo === 'mensalidade' || campo === 'mensalidade_sim' || campo === 'adesao') {
            valorParaSalvar = converterMoedaParaNumero(valorParaSalvar);
          } else if (campo === 'situacao' || campo === 'razao_social' || campo === 'nome_fantasia' || 
                     campo === 'contrato_enviado' || campo === 'contrato_assinado' || campo === 'observacoes' ||
                     campo === 'plano' || campo === 'etapa' || campo === 'inputEtapaSearch') {
            valorParaSalvar = normalizarTexto(valorParaSalvar);
            console.log(`🔥🔥🔥 Salvando ${campo}: "${valorParaSalvar}" na coluna ${colunaIndex + 1}`);
          }
          
          const valorAtual = dadosAtuais[colunaIndex];
          const valorNovo = valorParaSalvar;
          
          console.log(`🔍 Comparando campo ${campo}:`);
          console.log(`  Valor atual: "${valorAtual}"`);
          console.log(`  Valor novo: "${valorNovo}"`);
          console.log(`  São iguais? ${String(valorAtual).trim() === String(valorNovo).trim()}`);
          
          if (String(valorAtual).trim() !== String(valorNovo).trim()) {
            novosDados[colunaIndex] = valorParaSalvar;
            houveAlteracao = true;
            console.log(`✅ Campo "${campo}" alterado: "${valorAtual}" -> "${valorParaSalvar}" na coluna ${colunaIndex + 1}`);
          } else {
            console.log(`⚠️ Campo "${campo}" não alterado (mesmo valor)`);
          }
        } else if (dados[campo] !== undefined) {
          console.log(`⚠️ Campo "${campo}" não mapeado ou waitlabel incorreto`);
        }
      }
      
      if (houveAlteracao) {
        novosDados[COLUNAS.ULTIMA_ETAPA] = formatarDataBrasilSimples();
        console.log(`📅 Última etapa atualizada: ${novosDados[COLUNAS.ULTIMA_ETAPA]}`);
      }
      
      if (houveAlteracao) {
        console.log(`💾 Salvando linha ${linha}:`, novosDados);
        sheet.getRange(linha, 1, 1, novosDados.length).setValues([novosDados]);
        
        if (waitlabel === 'Sim_Facilita') {
          if (novosDados[COLUNAS.PLANO]) {
            sheet.getRange(linha, COLUNAS.PLANO + 1).setNumberFormat('@');
            console.log(`✅ Plano formatado como texto: ${novosDados[COLUNAS.PLANO]}`);
          }
          
          sheet.getRange(linha, COLUNAS.MDR + 1).setNumberFormat('0.00%');
          sheet.getRange(linha, COLUNAS.TIS + 1).setNumberFormat('0.00%');
          sheet.getRange(linha, COLUNAS.REBATE + 1).setNumberFormat('0.00%');
          
          sheet.getRange(linha, COLUNAS.MENSALIDADE + 1).setNumberFormat('"R$"#,##0.00');
          sheet.getRange(linha, COLUNAS.ADESAO + 1).setNumberFormat('"R$"#,##0.00');
          
          if (novosDados[COLUNAS.PGTO_ADESAO]) {
            sheet.getRange(linha, COLUNAS.PGTO_ADESAO + 1).setNumberFormat('dd/mm/yyyy');
          }
          
          if (novosDados[COLUNAS.DATA_CRIACAO]) {
            sheet.getRange(linha, COLUNAS.DATA_CRIACAO + 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
          }
          
          if (novosDados[COLUNAS.TREINADO]) {
            sheet.getRange(linha, COLUNAS.TREINADO + 1).setNumberFormat('@');
            console.log(`✅ Treinado formatado como texto: ${novosDados[COLUNAS.TREINADO]}`);
          }
        } else {
          sheet.getRange(linha, COLUNAS_PADRAO.MDR + 1).setNumberFormat('0.00%');
          sheet.getRange(linha, COLUNAS_PADRAO.TIS + 1).setNumberFormat('0.00%');
          sheet.getRange(linha, COLUNAS_PADRAO.REBATE + 1).setNumberFormat('0.00%');
          
          sheet.getRange(linha, COLUNAS_PADRAO.MENSALIDADE + 1).setNumberFormat('"R$"#,##0.00');
          sheet.getRange(linha, COLUNAS_PADRAO.MENSALIDADE_SIM + 1).setNumberFormat('"R$"#,##0.00');
          sheet.getRange(linha, COLUNAS_PADRAO.ADESAO + 1).setNumberFormat('"R$"#,##0.00');
          
          if (novosDados[COLUNAS_PADRAO.DATA_CRIACAO]) {
            sheet.getRange(linha, COLUNAS_PADRAO.DATA_CRIACAO + 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
          }
        }
        
        registrosAtualizados++;
        
        Utilities.sleep(50);
      } else {
        console.log(`⚠️ Nenhuma alteração na linha ${linha}`);
      }
    }
    
    SpreadsheetApp.flush();
    
    const mensagem = registrosAtualizados > 0 
      ? `✅ ${registrosAtualizados} registro(s) atualizado(s) com sucesso!`
      : `⚠️ Nenhuma alteração foi aplicada (valores já estavam corretos).`;
    
    return {
      success: true,
      registrosAtualizados: registrosAtualizados,
      message: mensagem
    };
    
  } catch (error) {
    console.error("❌ Erro em aplicarAlteracoesATodos:", error);
    console.error("Stack trace:", error.stack);
    
    return {
      success: false,
      message: '❌ Erro: ' + error.message
    };
  }
}

// 🔥🔥🔥 FUNÇÃO ATUALIZADA PARA BUSCAR CADASTROS POR CNPJ
function buscarTodosCadastrosPorCNPJComWaitlabel(cnpj, waitlabel) {
  try {
    const sheet = getSheetByName(waitlabel);
    const ultimaLinha = sheet.getLastRow();
    if (ultimaLinha < 2) return [];
    
    const COLUNAS = getColunasConfig(waitlabel);
    const totalColunas = waitlabel === 'Sim_Facilita' ? 23 : 19;
    
    const dados = sheet.getRange(2, 1, ultimaLinha - 1, totalColunas).getValues();
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    const cadastrosEncontrados = [];
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      const cnpjCadastro = linha[COLUNAS.CNPJ]?.toString().replace(/\D/g, '') || '';
      
      if (cnpjCadastro === cnpjBuscado) {
        let ultimaEtapaFormatada = '';
        if (linha[COLUNAS.ULTIMA_ETAPA] instanceof Date) {
          ultimaEtapaFormatada = formatarDataParaExibicao(linha[COLUNAS.ULTIMA_ETAPA]);
        } else {
          ultimaEtapaFormatada = linha[COLUNAS.ULTIMA_ETAPA]?.toString().trim() || '';
        }
        
        // 🔥 Formatar DATA_CRIACAO
        let dataCriacaoFormatada = '';
        if (linha[COLUNAS.DATA_CRIACAO] instanceof Date) {
          dataCriacaoFormatada = formatarDataParaExibicao(linha[COLUNAS.DATA_CRIACAO]);
        } else {
          dataCriacaoFormatada = linha[COLUNAS.DATA_CRIACAO]?.toString().trim() || '';
        }
        
        const cadastro = {
          id: i + 2,
          razao_social: linha[COLUNAS.RAZAO_SOCIAL]?.toString().trim() || '',
          nome_fantasia: linha[COLUNAS.NOME_FANTASIA]?.toString().trim() || '',
          cnpj: formatarCNPJParaExibicao(linha[COLUNAS.CNPJ]?.toString().trim() || ''),
          fornecedor: linha[COLUNAS.FORNECEDOR]?.toString().trim() || '',
          ultima_etapa: ultimaEtapaFormatada,
          etapa: linha[COLUNAS.ETAPA]?.toString().trim() || '',
          observacoes: linha[COLUNAS.OBSERVACAO]?.toString().trim() || '',
          contrato_enviado: linha[COLUNAS.CONTRATO_ENVIADO]?.toString().trim() || '',
          contrato_assinado: linha[COLUNAS.CONTRATO_ASSINADO]?.toString().trim() || '',
          ativacao: linha[COLUNAS.ATIVACAO]?.toString().trim() || '',
          link: linha[COLUNAS.LINK]?.toString().trim() || '',
          situacao: (linha[COLUNAS.SITUACAO]?.toString().trim() || 'NOVO REGISTRO'),
          data_criacao: dataCriacaoFormatada,
          waitlabel: waitlabel
        };
        
        if (waitlabel === 'Sim_Facilita') {
          cadastro.plano = linha[COLUNAS.PLANO]?.toString().trim() || '';
          cadastro.mensalidade = parseFloat(linha[COLUNAS.MENSALIDADE]) || 0;
          cadastro.vencimento = linha[COLUNAS.VENC]?.toString().trim() || '';
          cadastro.metodo_pgto = linha[COLUNAS.METODO_PGTO]?.toString().trim() || '';
          cadastro.mdr = formatarPercentualParaExibicao(linha[COLUNAS.MDR]);
          cadastro.tis = formatarPercentualParaExibicao(linha[COLUNAS.TIS]);
          cadastro.rebate = formatarPercentualParaExibicao(linha[COLUNAS.REBATE]);
          cadastro.adesao = processarAdesao(linha[COLUNAS.ADESAO]);
          cadastro.pgto_adesao = linha[COLUNAS.PGTO_ADESAO]?.toString().trim() || '';
          cadastro.treinado = processarTreinado(linha[COLUNAS.TREINADO]); // 🔥 CORREÇÃO: usar função processarTreinado
        } else {
          cadastro.mensalidade = parseFloat(linha[COLUNAS_PADRAO.MENSALIDADE]) || 0;
          cadastro.mensalidade_sim = parseFloat(linha[COLUNAS_PADRAO.MENSALIDADE_SIM]) || 0;
          cadastro.mdr = formatarPercentualParaExibicao(linha[COLUNAS_PADRAO.MDR]);
          cadastro.tis = formatarPercentualParaExibicao(linha[COLUNAS_PADRAO.TIS]);
          cadastro.rebate = formatarPercentualParaExibicao(linha[COLUNAS_PADRAO.REBATE]);
          cadastro.adesao = processarAdesao(linha[COLUNAS_PADRAO.ADESAO]);
        }
        
        cadastrosEncontrados.push(cadastro);
      }
    }
    
    return cadastrosEncontrados;
    
  } catch (error) {
    console.error("❌ Erro em buscarTodosCadastrosPorCNPJComWaitlabel:", error);
    return [];
  }
}

function getWaitlabelAtual() {
  const cache = CacheService.getScriptCache();
  const waitlabelAtual = cache.get('waitlabel_atual');
  return waitlabelAtual || WAITLABELS_CONFIG.WAITLABEL_PADRAO;
}

// 🔥🔥🔥 FUNÇÃO ATUALIZADA PARA BUSCAR CADASTRO POR ID E WAITLABEL
function buscarCadastroPorIDComWaitlabel(id, waitlabel) {
  try {
    console.log("🔍 Buscando cadastro ID:", id, "Waitlabel:", waitlabel);
    
    const sheet = getSheetByName(waitlabel);
    if (!sheet) {
      throw new Error(`Sheet ${waitlabel} não encontrado`);
    }
    
    const linha = parseInt(id);
    if (isNaN(linha) || linha < 2 || linha > sheet.getLastRow()) {
      throw new Error("ID inválido ou registro não encontrado");
    }
    
    const COLUNAS = getColunasConfig(waitlabel);
    const totalColunas = waitlabel === 'Sim_Facilita' ? 23 : 19;
    
    const linhaDados = sheet.getRange(linha, 1, 1, totalColunas).getValues()[0];
    
    let ultimaEtapaFormatada = '';
    if (linhaDados[COLUNAS.ULTIMA_ETAPA] instanceof Date) {
      ultimaEtapaFormatada = formatarDataParaExibicao(linhaDados[COLUNAS.ULTIMA_ETAPA]);
    } else {
      ultimaEtapaFormatada = linhaDados[COLUNAS.ULTIMA_ETAPA]?.toString().trim() || '';
    }
    
    // 🔥 Formatar DATA_CRIACAO
    let dataCriacaoFormatada = '';
    if (linhaDados[COLUNAS.DATA_CRIACAO] instanceof Date) {
      dataCriacaoFormatada = formatarDataParaExibicao(linhaDados[COLUNAS.DATA_CRIACAO]);
    } else {
      dataCriacaoFormatada = linhaDados[COLUNAS.DATA_CRIACAO]?.toString().trim() || '';
    }
    
    const cnpjDisplay = linhaDados[COLUNAS.CNPJ] ? formatarCNPJParaExibicao(linhaDados[COLUNAS.CNPJ]) : '';
    
    const cadastro = {
      id: linha,
      razao_social: linhaDados[COLUNAS.RAZAO_SOCIAL]?.toString().trim() || '',
      nome_fantasia: linhaDados[COLUNAS.NOME_FANTASIA]?.toString().trim() || '',
      cnpj: cnpjDisplay,
      fornecedor: linhaDados[COLUNAS.FORNECEDOR]?.toString().trim() || '',
      ultima_etapa: ultimaEtapaFormatada,
      etapa: linhaDados[COLUNAS.ETAPA]?.toString().trim() || '',
      observacoes: linhaDados[COLUNAS.OBSERVACAO]?.toString().trim() || '',
      contrato_enviado: linhaDados[COLUNAS.CONTRATO_ENVIADO]?.toString().trim() || '',
      contrato_assinado: linhaDados[COLUNAS.CONTRATO_ASSINADO]?.toString().trim() || '',
      ativacao: linhaDados[COLUNAS.ATIVACAO]?.toString().trim() || '',
      link: linhaDados[COLUNAS.LINK]?.toString().trim() || '',
      situacao: (linhaDados[COLUNAS.SITUACAO]?.toString().trim() || 'NOVO REGISTRO'),
      data_criacao: dataCriacaoFormatada,
      waitlabel: waitlabel
    };
    
    if (waitlabel === 'Sim_Facilita') {
      cadastro.plano = linhaDados[COLUNAS.PLANO]?.toString().trim() || '';
      cadastro.mensalidade = parseFloat(linhaDados[COLUNAS.MENSALIDADE]) || 0;
      cadastro.vencimento = linhaDados[COLUNAS.VENC]?.toString().trim() || '';
      cadastro.metodo_pgto = linhaDados[COLUNAS.METODO_PGTO]?.toString().trim() || '';
      cadastro.mdr = formatarPercentualParaExibicao(linhaDados[COLUNAS.MDR]);
      cadastro.tis = formatarPercentualParaExibicao(linhaDados[COLUNAS.TIS]);
      cadastro.rebate = formatarPercentualParaExibicao(linhaDados[COLUNAS.REBATE]);
      cadastro.adesao = processarAdesao(linhaDados[COLUNAS.ADESAO]);
      cadastro.pgto_adesao = linhaDados[COLUNAS.PGTO_ADESAO]?.toString().trim() || '';
      cadastro.treinado = processarTreinado(linhaDados[COLUNAS.TREINADO]); // 🔥 CORREÇÃO: usar função processarTreinado
    } else {
      cadastro.mensalidade = parseFloat(linhaDados[COLUNAS_PADRAO.MENSALIDADE]) || 0;
      cadastro.mensalidade_sim = parseFloat(linhaDados[COLUNAS_PADRAO.MENSALIDADE_SIM]) || 0;
      cadastro.mdr = formatarPercentualParaExibicao(linhaDados[COLUNAS_PADRAO.MDR]);
      cadastro.tis = formatarPercentualParaExibicao(linhaDados[COLUNAS_PADRAO.TIS]);
      cadastro.rebate = formatarPercentualParaExibicao(linhaDados[COLUNAS_PADRAO.REBATE]);
      cadastro.adesao = processarAdesao(linhaDados[COLUNAS_PADRAO.ADESAO]);
    }
    
    console.log("✅ Cadastro encontrado:", cadastro);
    return cadastro;
    
  } catch (error) {
    console.error("❌ Erro em buscarCadastroPorIDComWaitlabel:", error);
    throw error;
  }
}

// 🔥 FUNÇÃO SIMPLES PARA BUSCAR CADASTRO POR ID (USA O WAITLABEL ATUAL)
function buscarCadastroPorID(id) {
  try {
    const waitlabel = getWaitlabelAtual();
    console.log("🔍 Buscando cadastro ID:", id, "Waitlabel:", waitlabel);
    
    return buscarCadastroPorIDComWaitlabel(id, waitlabel);
    
  } catch (error) {
    console.error("❌ Erro em buscarCadastroPorID:", error);
    throw error;
  }
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

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistema - Gestão de Cadastros')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function testar() {
  return { 
    success: true, 
    message: "✅ Sistema funcionando!",
    timestamp: new Date().toISOString()
  };
}

function testarBuscaCadastros(waitlabel) {
  try {
    console.log(`🔍 Testando busca para ${waitlabel}...`);
    const resultado = buscarTodosCadastrosComWaitlabel(waitlabel);
    console.log(`✅ Resultado: ${resultado.length} cadastros`);
    
    return {
      success: true,
      quantidade: resultado.length,
      cadastros: resultado.slice(0, 3)
    };
  } catch (error) {
    console.error("❌ Erro no teste:", error);
    return {
      success: false,
      error: error.message
    };
  }
}

// 🔥 FUNÇÃO PARA TESTAR A CONVERSÃO DE PERCENTUAIS
function testarConversaoPercentual() {
  const testes = [
    { input: '2,67%', esperado: 0.0267 },
    { input: '7,77%', esperado: 0.0777 },
    { input: '3%', esperado: 0.03 },
    { input: '2.67', esperado: 0.0267 },
    { input: '7.77', esperado: 0.0777 },
    { input: 3, esperado: 0.03 },
    { input: 0.0267, esperado: 0.0267 }
  ];
  
  const resultados = [];
  
  testes.forEach((teste, index) => {
    const resultado = converterPercentualParaDecimal(teste.input);
    const formatado = formatarPercentualParaExibicao(resultado);
    
    resultados.push({
      teste: index + 1,
      input: teste.input,
      tipo: typeof teste.input,
      resultadoDecimal: resultado,
      resultadoFormatado: formatado,
      esperado: teste.esperado,
      ok: Math.abs(resultado - teste.esperado) < 0.0001
    });
  });
  
  console.log("🧪 RESULTADOS DOS TESTES:");
  resultados.forEach(r => {
    console.log(`${r.ok ? '✅' : '❌'} Teste ${r.teste}: ${r.input} -> ${r.resultadoDecimal} (${r.resultadoFormatado})`);
  });
  
  return resultados;
}

// 🔥 NOVA FUNÇÃO PARA TESTAR O CAMPO TREINADO
function testarCampoTreinado() {
  const testes = [
    { input: 'SIM', esperado: 'SIM' },
    { input: 'Sim', esperado: 'SIM' },
    { input: 'sim', esperado: 'SIM' },
    { input: 'S', esperado: 'SIM' },
    { input: 'YES', esperado: 'SIM' },
    { input: 'Y', esperado: 'SIM' },
    { input: 'TRUE', esperado: 'SIM' },
    { input: '1', esperado: 'SIM' },
    { input: 'NAO', esperado: 'NAO' },
    { input: 'Não', esperado: 'NAO' },
    { input: 'N', esperado: 'NAO' },
    { input: 'NO', esperado: 'NAO' },
    { input: 'FALSE', esperado: 'NAO' },
    { input: '0', esperado: 'NAO' },
    { input: '', esperado: 'NAO' },
    { input: undefined, esperado: 'NAO' },
    { input: null, esperado: 'NAO' },
    { input: true, esperado: 'SIM' },
    { input: false, esperado: 'NAO' },
    { input: 1, esperado: 'SIM' },
    { input: 0, esperado: 'NAO' }
  ];
  
  console.log("🧪 TESTANDO CAMPO TREINADO:");
  
  const resultados = [];
  
  testes.forEach((teste, index) => {
    const resultado = processarTreinado(teste.input);
    const ok = resultado === teste.esperado;
    
    console.log(`${ok ? '✅' : '❌'} Teste ${index + 1}: "${teste.input}" -> "${resultado}" (esperado: "${teste.esperado}")`);
    
    resultados.push({
      teste: index + 1,
      input: teste.input,
      resultado: resultado,
      esperado: teste.esperado,
      ok: ok
    });
  });
  
  return {
    success: true,
    resultados: resultados,
    total: testes.length,
    acertos: resultados.filter(r => r.ok).length
  };
}

function debugRecebimentoDados(dados) {
  console.log("🔍 DEBUG RECEBIMENTO DADOS:");
  console.log("Dados recebidos do HTML:", JSON.stringify(dados, null, 2));
  console.log("Tipo dos dados:", typeof dados);
  console.log("Fornecedor:", dados.fornecedor, "tipo:", typeof dados.fornecedor);
  console.log("CNPJ:", dados.cnpj);
  console.log("Treinado (campo geral):", dados.treinado);
  console.log("Treinado BC:", dados.treinado_bc);
  console.log("Treinado Parcelex:", dados.treinado_parcelex);
  
  return {
    success: true,
    message: "Dados recebidos com sucesso",
    dadosRecebidos: dados
  };
}

function testarCadastroSimples() {
  console.log("🧪 TESTE SIMPLES - " + new Date());
  
  const dados = {
    razao_social: "TESTE " + new Date().getTime(),
    nome_fantasia: "TESTE",
    cnpj: "11.222.333/0001-44",
    fornecedor: ["BC"],
    treinado: "SIM",
    mensalidade: "R$ 100,00",
    contrato_enviado: "SIM",
    contrato_assinado: "SIM",
    situacao: "CADASTRADO",
    etapa: "CADASTRADO",
    acao: "cadastrar"
  };
  
  console.log("📤 Dados para teste:", dados);
  
  const resultado = processarCadastroComWaitlabel(dados, "Sim_Facilita");
  console.log("📥 Resultado:", resultado);
  
  return resultado;
}

function atualizarCadastroCompleto(dados, waitlabel) {
  try {
    console.log('🔄 ATUALIZAR CADASTRO COMPLETO CHAMADO');
    console.log('📋 Waitlabel:', waitlabel);
    console.log('📊 Dados recebidos:', dados);
    
    // 🔥 DEBUG ESPECIAL PARA TREINAMENTO
    console.log('🎓 CAMPOS DE TREINAMENTO NO DADOS:');
    for (var key in dados) {
      if (key.includes('treinado')) {
        console.log('  ' + key + ': "' + dados[key] + '"');
      }
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(waitlabel);
    if (!sheet) {
      return { success: false, message: 'Waitlabel não encontrado: ' + waitlabel };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // Encontrar índice da coluna ID
    var idIndex = headers.indexOf('id');
    if (idIndex === -1) {
      idIndex = headers.indexOf('ID');
    }
    
    if (idIndex === -1) {
      console.error('❌ Coluna ID não encontrada');
      return { success: false, message: 'Coluna ID não encontrada na planilha' };
    }
    
    console.log('🔍 Procurando ID:', dados.id);
    
    // Procurar a linha com o ID
    var linhaEncontrada = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][idIndex] == dados.id) {
        linhaEncontrada = i;
        console.log('✅ ID encontrado na linha:', i + 1);
        break;
      }
    }
    
    if (linhaEncontrada === -1) {
      console.error('❌ ID não encontrado na planilha:', dados.id);
      return { success: false, message: 'ID não encontrado: ' + dados.id };
    }
    
    // 🔥 ATUALIZAR CADA CAMPO
    var camposAtualizados = [];
    
    for (var key in dados) {
      if (key === 'id' || key === 'acao' || key === 'ultima_etapa') continue;
      
      var colIndex = headers.indexOf(key);
      if (colIndex === -1) {
        console.log('⚠️ Coluna não encontrada:', key);
        continue;
      }
      
      var valorAntigo = data[linhaEncontrada][colIndex];
      var valorNovo = dados[key];
      
      // Se o valor for diferente, atualizar
      if (String(valorAntigo) !== String(valorNovo)) {
        sheet.getRange(linhaEncontrada + 1, colIndex + 1).setValue(valorNovo);
        camposAtualizados.push(key);
        
        // 🔥 LOG ESPECIAL PARA TREINAMENTO
        if (key.includes('treinado')) {
          console.log('🎓 TREINAMENTO ATUALIZADO:');
          console.log('  Coluna:', colIndex + 1);
          console.log('  Linha:', linhaEncontrada + 1);
          console.log('  De:', valorAntigo);
          console.log('  Para:', valorNovo);
        }
      }
    }
    
    // Atualizar data da última atualização
    var colunaUltimaEtapa = headers.indexOf('ultima_etapa');
    if (colunaUltimaEtapa !== -1) {
      var dataAtual = new Date();
      var dataFormatada = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
      sheet.getRange(linhaEncontrada + 1, colunaUltimaEtapa + 1).setValue(dataFormatada);
      console.log('⏰ Última atualização:', dataFormatada);
    }
    
    console.log('✅ Atualização concluída. Campos atualizados:', camposAtualizados.length, camposAtualizados);
    
    return { 
      success: true, 
      message: 'Cadastro atualizado com sucesso!',
      camposAtualizados: camposAtualizados,
      totalCampos: camposAtualizados.length
    };
    
  } catch (error) {
    console.error('❌ ERRO em atualizarCadastroCompleto:', error);
    return { 
      success: false, 
      message: 'Erro ao atualizar: ' + error.toString(),
      erro: error.toString()
    };
  }
}

function testarAtualizacaoSimples() {
  console.log("🧪 TESTANDO ATUALIZAÇÃO SIMPLES");
  
  const dados = {
    id: "255", // Use um ID que existe
    acao: "atualizar",
    waitlabel: "Sim_Facilita",
    razao_social: "TESTE ATUALIZAÇÃO",
    fornecedor: ["AGIL"],
    cnpj: "33.333.333/3444-44",
    situacao: "CADASTRADO",
    etapa: "CADASTRADO",
    treinado_agil: "SIM",
    mdr_agil: "0.037",
    tis_agil: "0.03"
  };
  
  console.log("📤 Dados de teste:", dados);
  
  try {
    const resultado = processarCadastroComWaitlabel(dados, "Sim_Facilita");
    console.log("📥 Resultado:", resultado);
    return resultado;
  } catch (error) {
    console.error("❌ Erro no teste:", error);
    return { success: false, message: error.toString() };
  }
}
