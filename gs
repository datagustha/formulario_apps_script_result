// CONFIGURAÇÕES
const CONFIG = {
  ID_PLANILHA: "1V4iGN14UpIQcwf3qKU0_Wbiy2exdW2WUmrYTniy0upA",
  ABA_PRINCIPAL: "Result"
};

// FUNÇÃO PRINCIPAL
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistema RESULT - Gestão de Cadastros')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// INCLUIR ARQUIVOS HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 🔥🔥🔥 FUNÇÃO PRINCIPAL CORRIGIDA - PROCESSAR CADASTRO (PARA AMBOS CADASTRO E ATUALIZAÇÃO)
function processarCadastro(dados) {
  try {
    console.log("🎯 PROCESSAR CADASTRO - Dados recebidos:", dados);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    let aba = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);

    if (!aba) {
      console.log("📝 Criando nova aba...");
      aba = ss.insertSheet(CONFIG.ABA_PRINCIPAL);
      // Cabeçalho com 17 colunas na ORDEM CORRETA
      const cabecalho = [
        'Razão Social', 'Nome Fantasia', 'CNPJ', 'Tipo', 'Fornecedor', 
        'Evento', 'Data Status', 'Observação', 'Contrato Enviado', 'Contrato Assinado',
        'Ativação', 'Link', 'Mensalidade', 'Tarifa', '% Tarifa', 'Adesão', 'Situação'
      ];
      aba.getRange('A1:Q1').setValues([cabecalho]);
      aba.getRange(1, 1, 1, cabecalho.length)
        .setBackground("#7E3E9A")
        .setFontColor("white")
        .setFontWeight("bold");
      aba.setFrozenRows(1);
    }

    if (dados.acao === 'cadastrar') {
      return cadastrarNovo(aba, dados);
    } else if (dados.acao === 'atualizar') {
      return atualizarCadastro(aba, dados);
    } else {
      return { success: false, message: "Ação não reconhecida" };
    }

  } catch (error) {
    console.error("❌ Erro em processarCadastro:", error);
    return { success: false, message: "Erro: " + error.message };
  }
}

// 🔥🔥🔥 FUNÇÃO CADASTRAR NOVO - MULTIPLOS FORNECEDORES FUNCIONANDO
function cadastrarNovo(aba, dados) {
  try {
    console.log("🆕 CADASTRAR NOVO - INICIANDO");
    console.log("📋 Fornecedores recebidos:", dados.fornecedores);
    
    // Verificar se já existe algum cadastro com este CNPJ
    const cadastroExistente = buscarCadastroPorCNPJ(dados.cnpj);
    if (cadastroExistente.encontrado) {
      return { success: false, message: "❌ Este CNPJ já está cadastrado!" };
    }

    const ultimaLinha = aba.getLastRow();
    let linhaInserir = Math.max(2, ultimaLinha + 1);
    const resultados = [];
    let registrosCriados = 0;

    // ✅ CORREÇÃO: Apenas ajustar "Novo registro" para "Novo Registro"
    let situacaoParaSalvar = dados.situacao || 'Novo Registro';
    if (situacaoParaSalvar === 'Novo registro') {
      situacaoParaSalvar = 'Novo Registro';
    }

    console.log(`🎯 Situação: "${dados.situacao}" → "${situacaoParaSalvar}"`);

    for (let i = 0; i < dados.fornecedores.length; i++) {
      const fornecedorObj = dados.fornecedores[i];
      
      // Processar fornecedor
      let nomeFornecedor = '';
      let tarifaFornecedor = '';
      let percentualTarifaFornecedor = '0%';
      
      if (typeof fornecedorObj === 'object' && fornecedorObj !== null) {
        nomeFornecedor = fornecedorObj.nome || '';
        tarifaFornecedor = fornecedorObj.tarifa || '';
        percentualTarifaFornecedor = fornecedorObj.percentual_tarifa || '0%';
      }

      // Validar se o nome do fornecedor está preenchido
      if (!nomeFornecedor || nomeFornecedor.trim() === '') {
        resultados.push(`❌ Fornecedor sem nome - pulado`);
        continue;
      }

      // Converter valores monetários
      let mensalidadeNumero = parseFloat(dados.mensalidade) || 0;
      let adesaoNumero = processarAdesaoParaSalvar(dados.adesao);

      // Array com 17 colunas na ORDEM CORRETA
      const linhaDados = [
        dados.razao_social || '',
        dados.nome_fantasia || '',
        dados.cnpj ? dados.cnpj.toString() : '',
        dados.tipo || '',
        nomeFornecedor,
        dados.evento || '',
        dados.data_status || '',
        dados.observacoes || '',
        dados.contrato_enviado || '',
        dados.contrato_assinado || '',
        dados.ativacao || '',
        dados.link || '',
        mensalidadeNumero,
        tarifaFornecedor,
        percentualTarifaFornecedor,
        adesaoNumero,
        situacaoParaSalvar
      ];

      console.log(`📝 Inserindo fornecedor ${i + 1}: ${nomeFornecedor}`);
      
      try {
        const range = aba.getRange(linhaInserir, 1, 1, linhaDados.length);
        range.setValues([linhaDados]);
        
        // Formatar colunas monetárias
        aba.getRange(linhaInserir, 13).setNumberFormat('"R$"#,##0.00');
        aba.getRange(linhaInserir, 16).setNumberFormat('"R$"#,##0.00');
        
        SpreadsheetApp.flush();
        
        linhaInserir++;
        registrosCriados++;
        resultados.push(`✅ ${nomeFornecedor} - ${tarifaFornecedor} ${percentualTarifaFornecedor}`);
        
      } catch (erroInsercao) {
        console.error(`❌ Erro:`, erroInsercao.message);
        resultados.push(`❌ ${nomeFornecedor} - ERRO: ${erroInsercao.message}`);
      }
    }

    // Mensagem final
    const sucessos = resultados.filter(r => r.includes('✅')).length;
    const erros = resultados.filter(r => r.includes('❌')).length;
    
    let mensagem = '';
    if (erros === 0) {
      mensagem = `✅ "${dados.razao_social}" cadastrado com sucesso para ${sucessos} fornecedor(es)!`;
    } else if (sucessos === 0) {
      mensagem = `❌ Erro ao cadastrar "${dados.razao_social}" para todos os fornecedores!`;
    } else {
      mensagem = `⚠️ "${dados.razao_social}" cadastrado parcialmente: ${sucessos} sucesso(s), ${erros} erro(s)`;
    }

    return { 
      success: erros === 0,
      message: mensagem,
      registrosCriados: registrosCriados,
      detalhes: resultados
    };

  } catch (error) {
    console.error("❌ Erro:", error);
    return { 
      success: false, 
      message: "Erro ao cadastrar: " + error.message 
    };
  }
}

// 🔥🔥🔥 FUNÇÃO ATUALIZAR CADASTRO - CORRIGIDA E FUNCIONANDO
function atualizarCadastro(aba, dados) {
  try {
    console.log("✏️ ATUALIZAR CADASTRO - INICIANDO");
    console.log("📋 Dados recebidos:", dados);
    
    const linhaAtualizar = parseInt(dados.id);

    if (linhaAtualizar < 2 || linhaAtualizar > aba.getLastRow()) {
      return { success: false, message: "Registro não encontrado" };
    }

    // 🔥 CORREÇÃO: Processar fornecedores corretamente
    let fornecedorParaAtualizar = '';
    let tarifaParaAtualizar = dados.tarifa || '';
    let percentualParaAtualizar = dados.percentual_tarifa || '0%';

    if (Array.isArray(dados.fornecedores) && dados.fornecedores.length > 0) {
      const primeiroFornecedor = dados.fornecedores[0];
      fornecedorParaAtualizar = primeiroFornecedor.nome || primeiroFornecedor;
      tarifaParaAtualizar = primeiroFornecedor.tarifa || tarifaParaAtualizar;
      percentualParaAtualizar = primeiroFornecedor.percentual_tarifa || percentualParaAtualizar;
    } else if (typeof dados.fornecedores === 'string') {
      fornecedorParaAtualizar = dados.fornecedores;
    } else {
      fornecedorParaAtualizar = dados.fornecedor || '';
    }

    // Converter valores monetários para número
    let mensalidadeNumero = converterMoedaParaNumero(dados.mensalidade);
    let adesaoNumero = processarAdesaoParaSalvar(dados.adesao);

    // Garantir que a situação seja válida
    const situacaoValida = (dados.situacao && dados.situacao.trim() !== '') ? dados.situacao : 'Novo registro';

    // Array com 17 colunas na ORDEM CORRETA
    const novosDados = [
      dados.razao_social || '',
      dados.nome_fantasia || '',
      dados.cnpj ? formatarCNPJNoSheets(dados.cnpj) : '',
      dados.tipo || '',
      fornecedorParaAtualizar,
      dados.evento || '',
      dados.data_status || '',
      dados.observacoes || '',
      dados.contrato_enviado || '',
      dados.contrato_assinado || '',
      dados.ativacao || '',
      dados.link || '',
      mensalidadeNumero,
      tarifaParaAtualizar,
      percentualParaAtualizar,
      adesaoNumero,
      situacaoValida
    ];

    console.log("📝 Atualizando linha:", linhaAtualizar);
    console.log("📊 Novos dados:", novosDados);
    
    aba.getRange(linhaAtualizar, 1, 1, novosDados.length).setValues([novosDados]);
    
    // Formatar colunas monetárias
    aba.getRange(linhaAtualizar, 13).setNumberFormat('"R$"#,##0.00');
    aba.getRange(linhaAtualizar, 16).setNumberFormat('"R$"#,##0.00');

    SpreadsheetApp.flush();

    return { 
      success: true, 
      message: `✅ "${dados.razao_social}" atualizado com sucesso!` 
    };

  } catch (error) {
    console.error("❌ Erro em atualizarCadastro:", error);
    return { success: false, message: "Erro ao atualizar: " + error.message };
  }
}

// 🔥 MANTER FUNÇÃO salvarCadastro PARA COMPATIBILIDADE
function salvarCadastro(dados) {
  return processarCadastro(dados);
}

// 🔥 MANTER FUNÇÃO processarAtualizacao PARA COMPATIBILIDADE
function processarAtualizacao(dados) {
  return processarCadastro(dados);
}

// 🔥🔥🔥 FUNÇÕES DE BUSCA (MANTIDAS)
function buscarTodosCadastros() {
  try {
    console.log("🔍 Iniciando busca de todos os cadastros...");
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);
    if (!aba) {
      console.log("❌ Aba não encontrada:", CONFIG.ABA_PRINCIPAL);
      return [];
    }
    
    const ultimaLinha = aba.getLastRow();
    console.log("📊 Última linha:", ultimaLinha);
    
    if (ultimaLinha < 2) {
      console.log("ℹ️ Nenhum dado além do cabeçalho");
      return [];
    }
    
    // Buscar dados na ORDEM CORRETA (17 colunas)
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    console.log("📈 Dados brutos encontrados:", dados.length);
    
    const cadastros = [];
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      
      // Pular linhas vazias
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      // Formatar data corretamente
      let dataStatusFormatada = '';
      if (linha[6] && linha[6] instanceof Date) {
        dataStatusFormatada = Utilities.formatDate(linha[6], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (linha[6]) {
        dataStatusFormatada = linha[6].toString();
      }
      
      let ativacaoFormatada = '';
      if (linha[10] && linha[10] instanceof Date) {
        ativacaoFormatada = Utilities.formatDate(linha[10], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (linha[10]) {
        ativacaoFormatada = linha[10].toString();
      }
      
      const cadastro = {
        id: i + 2,
        razao_social: linha[0]?.toString().trim() || '',
        nome_fantasia: linha[1]?.toString().trim() || '',
        cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''),
        tipo: linha[3]?.toString().trim() || '',
        fornecedor: linha[4]?.toString().trim() || '',
        evento: linha[5]?.toString().trim() || '',
        data_status: dataStatusFormatada,
        observacoes: linha[7]?.toString().trim() || '',
        contrato_enviado: linha[8]?.toString().trim() || '',
        contrato_assinado: linha[9]?.toString().trim() || '',
        ativacao: ativacaoFormatada,
        link: linha[11]?.toString().trim() || '',
        mensalidade: parseFloat(linha[12]) || 0,
        tarifa: linha[13]?.toString().trim() || '',
        percentual_tarifa: linha[14]?.toString().trim() || '',
        adesao: processarAdesao(linha[15]),
        situacao: (linha[16]?.toString().trim() || 'Novo registro')
      };
      
      cadastros.push(cadastro);
    }
    
    console.log("✅ Cadastros processados:", cadastros.length);
    return cadastros;
    
  } catch (error) {
    console.error("❌ Erro em buscarTodosCadastros:", error);
    return [];
  }
}

function buscarCadastroPorCNPJ(cnpj) {
  try {
    console.log("🔍 Buscando CNPJ:", cnpj);
    
    if (!cnpj || cnpj.toString().replace(/\D/g, '').length < 11) {
      return { encontrado: false, mensagem: "CNPJ inválido" };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);
    if (!aba) return { encontrado: false, mensagem: "Planilha não encontrada" };
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) return { encontrado: false, mensagem: "Nenhum dado encontrado" };
    
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    
    console.log("🔎 Procurando CNPJ limpo:", cnpjBuscado);
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      const cnpjCadastro = linha[2]?.toString().replace(/\D/g, '') || '';
      
      // Pular linhas vazias
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      if (cnpjCadastro === cnpjBuscado) {
        console.log("✅ Cadastro encontrado na linha:", i + 2);
        
        // Formatar datas corretamente
        let dataStatusFormatada = '';
        if (linha[6] && linha[6] instanceof Date) {
          dataStatusFormatada = Utilities.formatDate(linha[6], Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (linha[6]) {
          dataStatusFormatada = linha[6].toString();
        }
        
        let ativacaoFormatada = '';
        if (linha[10] && linha[10] instanceof Date) {
          ativacaoFormatada = Utilities.formatDate(linha[10], Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (linha[10]) {
          ativacaoFormatada = linha[10].toString();
        }
        
        return {
          encontrado: true,
          id: i + 2,
          razao_social: linha[0]?.toString().trim() || '',
          nome_fantasia: linha[1]?.toString().trim() || '',
          cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''),
          tipo: linha[3]?.toString().trim() || '',
          fornecedor: linha[4]?.toString().trim() || '',
          evento: linha[5]?.toString().trim() || '',
          data_status: dataStatusFormatada,
          observacoes: linha[7]?.toString().trim() || '',
          contrato_enviado: linha[8]?.toString().trim() || '',
          contrato_assinado: linha[9]?.toString().trim() || '',
          ativacao: ativacaoFormatada,
          link: linha[11]?.toString().trim() || '',
          mensalidade: parseFloat(linha[12]) || 0,
          tarifa: linha[13]?.toString().trim() || '',
          percentual_tarifa: linha[14]?.toString().trim() || '',
          adesao: processarAdesao(linha[15]),
          situacao: (linha[16]?.toString().trim() || 'Novo registro')
        };
      }
    }
    
    console.log("❌ Cadastro não encontrado para CNPJ:", cnpjBuscado);
    return { encontrado: false, mensagem: "Cadastro não encontrado" };
    
  } catch (error) {
    console.error("❌ Erro em buscarCadastroPorCNPJ:", error);
    return { encontrado: false, mensagem: "Erro: " + error.message };
  }
}

function buscarCadastroPorID(id) {
  try {
    console.log("🔍 Buscando cadastro por ID:", id);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);
    if (!aba) return { encontrado: false, mensagem: "Planilha não encontrada" };
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < id) return { encontrado: false, mensagem: "Registro não encontrado" };
    
    const linha = aba.getRange(id, 1, 1, 17).getValues()[0];
    
    // Verificar se a linha não está vazia
    if (!linha[0] || linha[0].toString().trim() === '') {
      return { encontrado: false, mensagem: "Registro vazio ou não encontrado" };
    }
    
    // Formatar datas corretamente
    let dataStatusFormatada = '';
    if (linha[6] && linha[6] instanceof Date) {
      dataStatusFormatada = Utilities.formatDate(linha[6], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (linha[6]) {
      dataStatusFormatada = linha[6].toString();
    }
    
    let ativacaoFormatada = '';
    if (linha[10] && linha[10] instanceof Date) {
      ativacaoFormatada = Utilities.formatDate(linha[10], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (linha[10]) {
      ativacaoFormatada = linha[10].toString();
    }
    
    return {
      encontrado: true,
      id: id,
      razao_social: linha[0]?.toString().trim() || '',
      nome_fantasia: linha[1]?.toString().trim() || '',
      cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''),
      tipo: linha[3]?.toString().trim() || '',
      fornecedor: linha[4]?.toString().trim() || '',
      evento: linha[5]?.toString().trim() || '',
      data_status: dataStatusFormatada,
      observacoes: linha[7]?.toString().trim() || '',
      contrato_enviado: linha[8]?.toString().trim() || '',
      contrato_assinado: linha[9]?.toString().trim() || '',
      ativacao: ativacaoFormatada,
      link: linha[11]?.toString().trim() || '',
      mensalidade: parseFloat(linha[12]) || 0,
      tarifa: linha[13]?.toString().trim() || '',
      percentual_tarifa: linha[14]?.toString().trim() || '',
      adesao: processarAdesao(linha[15]),
      situacao: (linha[16]?.toString().trim() || 'Novo registro')
    };
    
  } catch (error) {
    console.error("❌ Erro em buscarCadastroPorID:", error);
    return { encontrado: false, mensagem: "Erro: " + error.message };
  }
}

// 🔥🔥🔥 FUNÇÕES AUXILIARES
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
  if (!valorAdesao) return 0;
  const valorStr = valorAdesao.toString().trim();
  if (valorStr === 'Isento') {
    return 0;
  }
  return converterMoedaParaNumero(valorStr);
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

// 🔥 FUNÇÃO DEBUG
function debugFormulario(dados) {
  console.log("🎯 DEBUG FORMULÁRIO - DADOS RECEBIDOS:");
  console.log("Razão Social:", dados.razao_social);
  console.log("CNPJ:", dados.cnpj);
  console.log("Tipo:", dados.tipo);
  console.log("Quantidade de fornecedores:", dados.fornecedores ? dados.fornecedores.length : 0);
  console.log("Fornecedores detalhados:", dados.fornecedores);
  console.log("Ação:", dados.acao);
  console.log("DADOS COMPLETOS:", JSON.stringify(dados, null, 2));
  
  return {
    success: true,
    message: "✅ Debug recebido - verifique os logs",
    quantidadeFornecedores: dados.fornecedores ? dados.fornecedores.length : 0,
    estruturaFornecedores: dados.fornecedores ? dados.fornecedores.map(f => ({
      tipo: typeof f,
      nome: f.nome || f,
      tarifa: f.tarifa || 'N/A',
      percentual: f.percentual_tarifa || 'N/A'
    })) : []
  };
}

function testar() {
  return { 
    success: true, 
    message: "✅ Sistema funcionando!",
    timestamp: new Date().toISOString(),
    totalCadastros: buscarTodosCadastros().length
  };
}
