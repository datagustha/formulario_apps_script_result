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

// BUSCAR TODOS OS CADASTROS - CORRIGIDA COM ORDEM CERTA
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
    
    // 🔥 CORREÇÃO: Buscar dados na ORDEM CORRETA (17 colunas)
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    console.log("📈 Dados brutos encontrados:", dados.length);
    
    const cadastros = [];
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      
      // Pular linhas vazias
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      // Formatar data corretamente
      let dataStatusFormatada = '';
      if (linha[5] && linha[5] instanceof Date) {
        dataStatusFormatada = Utilities.formatDate(linha[5], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (linha[5]) {
        dataStatusFormatada = linha[5].toString();
      }
      
      let ativacaoFormatada = '';
      if (linha[9] && linha[9] instanceof Date) {
        ativacaoFormatada = Utilities.formatDate(linha[9], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (linha[9]) {
        ativacaoFormatada = linha[9].toString();
      }
      
      const cadastro = {
        id: i + 2,
        // 🔥 ORDEM CORRETA DAS COLUNAS:
        razao_social: linha[0]?.toString().trim() || '',
        nome_fantasia: linha[1]?.toString().trim() || '',
        cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''),
        tipo: linha[3]?.toString().trim() || '',
        fornecedor: linha[4]?.toString().trim() || '',
        evento: linha[5]?.toString().trim() || '', // ✅ COLUNA 6: EVENTO
        data_status: dataStatusFormatada, // ✅ COLUNA 7: DATA STATUS
        observacoes: linha[7]?.toString().trim() || '', // ✅ COLUNA 8: STATUS (OBSERVAÇÕES)
        contrato_enviado: linha[8]?.toString().trim() || '', // ✅ COLUNA 9
        contrato_assinado: linha[9]?.toString().trim() || '', // ✅ COLUNA 10
        ativacao: ativacaoFormatada, // ✅ COLUNA 11: ATIVAÇÃO
        link: linha[11]?.toString().trim() || '', // ✅ COLUNA 12: LINK
        mensalidade: parseFloat(linha[12]) || 0, // ✅ COLUNA 13: MENSALIDADE
        tarifa: linha[13]?.toString().trim() || '', // ✅ COLUNA 14: TARIFA
        percentual_tarifa: linha[14]?.toString().trim() || '', // ✅ COLUNA 15: % TARIFA
        adesao: parseFloat(linha[15]) || 0, // ✅ COLUNA 16: ADESÃO (AGORA IGUAL MENSALIDADE)
        situacao: linha[16]?.toString().trim() || 'Novo registro' // ✅ COLUNA 17: SITUAÇÃO
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

// BUSCAR CADASTRO POR CNPJ - CORRIGIDA
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
    
    // 🔥 CORREÇÃO: Buscar 17 colunas na ORDEM CORRETA
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
        if (linha[5] && linha[5] instanceof Date) {
          dataStatusFormatada = Utilities.formatDate(linha[5], Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (linha[5]) {
          dataStatusFormatada = linha[5].toString();
        }
        
        let ativacaoFormatada = '';
        if (linha[9] && linha[9] instanceof Date) {
          ativacaoFormatada = Utilities.formatDate(linha[9], Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (linha[9]) {
          ativacaoFormatada = linha[9].toString();
        }
        
        return {
          encontrado: true,
          id: i + 2,
          // 🔥 ORDEM CORRETA:
          razao_social: linha[0]?.toString().trim() || '',
          nome_fantasia: linha[1]?.toString().trim() || '',
          cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''),
          tipo: linha[3]?.toString().trim() || '',
          fornecedor: linha[4]?.toString().trim() || '',
          evento: linha[5]?.toString().trim() || '', // ✅ EVENTO
          data_status: dataStatusFormatada, // ✅ DATA STATUS
          observacoes: linha[7]?.toString().trim() || '', // ✅ STATUS (OBSERVAÇÕES)
          contrato_enviado: linha[8]?.toString().trim() || '',
          contrato_assinado: linha[9]?.toString().trim() || '',
          ativacao: ativacaoFormatada, // ✅ ATIVAÇÃO
          link: linha[11]?.toString().trim() || '', // ✅ LINK
          adesao: parseFloat(linha[15]) || 0, // ✅ MENSALIDADE
          tarifa: linha[13]?.toString().trim() || '', // ✅ TARIFA
          percentual_tarifa: linha[14]?.toString().trim() || '', // ✅ % TARIFA
          adesao: processarAdesao(linha[15]), // ✅ ADESÃO
          situacao: linha[16]?.toString().trim() || 'Novo registro' // ✅ SITUAÇÃO
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

// BUSCAR CADASTRO POR ID - CORRIGIDA
function buscarCadastroPorID(id) {
  try {
    console.log("🔍 Buscando cadastro por ID:", id);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);
    if (!aba) return { encontrado: false, mensagem: "Planilha não encontrada" };
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < id) return { encontrado: false, mensagem: "Registro não encontrado" };
    
    // 🔥 CORREÇÃO: Buscar 17 colunas na ORDEM CORRETA
    const linha = aba.getRange(id, 1, 1, 17).getValues()[0];
    
    // Verificar se a linha não está vazia
    if (!linha[0] || linha[0].toString().trim() === '') {
      return { encontrado: false, mensagem: "Registro vazio ou não encontrado" };
    }
    
    // Formatar datas corretamente
    let dataStatusFormatada = '';
    if (linha[5] && linha[5] instanceof Date) {
      dataStatusFormatada = Utilities.formatDate(linha[5], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (linha[5]) {
      dataStatusFormatada = linha[5].toString();
    }
    
    let ativacaoFormatada = '';
    if (linha[9] && linha[9] instanceof Date) {
      ativacaoFormatada = Utilities.formatDate(linha[9], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (linha[9]) {
      ativacaoFormatada = linha[9].toString();
    }
    
    return {
      encontrado: true,
      id: id,
      // 🔥 ORDEM CORRETA:
      razao_social: linha[0]?.toString().trim() || '',
      nome_fantasia: linha[1]?.toString().trim() || '',
      cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''),
      tipo: linha[3]?.toString().trim() || '',
      fornecedor: linha[4]?.toString().trim() || '',
      evento: linha[5]?.toString().trim() || '', // ✅ EVENTO
      data_status: dataStatusFormatada, // ✅ DATA STATUS
      observacoes: linha[7]?.toString().trim() || '', // ✅ STATUS (OBSERVAÇÕES)
      contrato_enviado: linha[8]?.toString().trim() || '',
      contrato_assinado: linha[9]?.toString().trim() || '',
      ativacao: ativacaoFormatada, // ✅ ATIVAÇÃO
      link: linha[11]?.toString().trim() || '', // ✅ LINK
      adesao: parseFloat(linha[15]) || 0, // ✅ CORRETO: ADESÃO IGUAL MENSALIDADE // ✅ MENSALIDADE
      tarifa: linha[13]?.toString().trim() || '', // ✅ TARIFA
      percentual_tarifa: linha[14]?.toString().trim() || '', // ✅ % TARIFA
      adesao: processarAdesao(linha[15]), // ✅ ADESÃO
      situacao: linha[16]?.toString().trim() || 'Novo registro' // ✅ SITUAÇÃO
    };
    
  } catch (error) {
    console.error("❌ Erro em buscarCadastroPorID:", error);
    return { encontrado: false, mensagem: "Erro: " + error.message };
  }
}

// SALVAR CADASTRO - CORRIGIDA COM MÚLTIPLOS FORNECEDORES
function salvarCadastro(dados) {
  try {
    console.log("💾 Salvando cadastro:", dados);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    let aba = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);

    if (!aba) {
      console.log("📝 Criando nova aba...");
      aba = ss.insertSheet(CONFIG.ABA_PRINCIPAL);
      // 🔥 CORREÇÃO: Cabeçalho com 17 colunas na ORDEM CORRETA
      const cabecalho = [
        'Razão Social', 'Nome Fantasia', 'CNPJ', 'Tipo', 'Fornecedor', 
        'Evento', 'Data Status', 'Status', 'Contrato Enviado', 'Contrato Assinado',
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
    console.error("❌ Erro em salvarCadastro:", error);
    return { success: false, message: "Erro: " + error.message };
  }
}

function cadastrarNovo(aba, dados) {
  try {
    console.log("🆕 Cadastrando novo:", dados.razao_social);
    console.log("📋 Fornecedores selecionados:", dados.fornecedores);
    
    // Verificar se já existe algum cadastro com este CNPJ
    const cadastroExistente = buscarCadastroPorCNPJ(dados.cnpj);
    if (cadastroExistente.encontrado) {
      return { success: false, message: "❌ Este CNPJ já está cadastrado!" };
    }

    const ultimaLinha = aba.getLastRow();
    let linhaInserir = Math.max(2, ultimaLinha + 1);

    // 🔥 CORREÇÃO: Criar um registro para CADA fornecedor selecionado
    const resultados = [];
    
    for (let i = 0; i < dados.fornecedores.length; i++) {
      const fornecedor = dados.fornecedores[i];
      
      // Converter valores monetários para número
      let mensalidadeNumero = converterMoedaParaNumero(dados.mensalidade);
      let adesaoNumero = processarAdesaoParaSalvar(dados.adesao);

      // 🔥 CORREÇÃO: Array com 17 colunas na ORDEM CORRETA
      const linhaDados = [
        dados.razao_social || '',
        dados.nome_fantasia || '',
        dados.cnpj ? formatarCNPJNoSheets(dados.cnpj) : '',
        dados.tipo || '',
        fornecedor, // ✅ FORNECEDOR INDIVIDUAL
        dados.evento || '', // ✅ EVENTO
        dados.data_status || '', // ✅ DATA STATUS
        dados.observacoes || '', // ✅ STATUS (OBSERVAÇÕES)
        dados.contrato_enviado || '',
        dados.contrato_assinado || '',
        dados.ativacao || '', // ✅ ATIVAÇÃO
        dados.link || '', // ✅ LINK
        mensalidadeNumero, // ✅ MENSALIDADE
        dados.tarifa || '', // ✅ TARIFA
        dados.percentual_tarifa || '', // ✅ % TARIFA
        adesaoNumero, // ✅ ADESÃO
        dados.situacao || 'Novo registro' // ✅ SITUAÇÃO (padrão: Novo registro)
      ];

      console.log(`📝 Inserindo registro ${i + 1}/${dados.fornecedores.length} para fornecedor: ${fornecedor}`);
      console.log("📊 Dados da linha:", linhaDados);
      
      aba.getRange(linhaInserir, 1, 1, linhaDados.length).setValues([linhaDados]);
      
      // Formatar colunas monetárias
      aba.getRange(linhaInserir, 13).setNumberFormat('"R$"#,##0.00'); // Mensalidade (coluna 13)
      aba.getRange(linhaInserir, 16).setNumberFormat('"R$"#,##0.00'); // Adesão (coluna 16)
      
      linhaInserir++;
      resultados.push(`✅ ${fornecedor}`);
    }

    const mensagem = resultados.length === 1 
      ? `✅ "${dados.razao_social}" cadastrado com sucesso para ${dados.fornecedores[0]}!`
      : `✅ "${dados.razao_social}" cadastrado com sucesso para ${dados.fornecedores.length} fornecedores!`;

    return { 
      success: true, 
      message: mensagem 
    };

  } catch (error) {
    console.error("❌ Erro em cadastrarNovo:", error);
    return { success: false, message: "Erro ao cadastrar: " + error.message };
  }
}

function atualizarCadastro(aba, dados) {
  try {
    console.log("✏️ Atualizando cadastro ID:", dados.id);
    
    const linhaAtualizar = parseInt(dados.id);

    if (linhaAtualizar < 2 || linhaAtualizar > aba.getLastRow()) {
      return { success: false, message: "Registro não encontrado" };
    }

    // Converter valores monetários para número
    let mensalidadeNumero = converterMoedaParaNumero(dados.mensalidade);
    let adesaoNumero = processarAdesaoParaSalvar(dados.adesao);

    // 🔥 CORREÇÃO: Pegar o PRIMEIRO fornecedor do array (na edição só temos um)
    const fornecedorParaAtualizar = dados.fornecedores && dados.fornecedores.length > 0 
      ? dados.fornecedores[0] 
      : '';

    // 🔥 CORREÇÃO: Array com 17 colunas na ORDEM CORRETA
    const novosDados = [
      dados.razao_social || '',
      dados.nome_fantasia || '',
      dados.cnpj ? formatarCNPJNoSheets(dados.cnpj) : '',
      dados.tipo || '',
      fornecedorParaAtualizar,
      dados.evento || '', // ✅ EVENTO
      dados.data_status || '', // ✅ DATA STATUS
      dados.observacoes || '', // ✅ STATUS (OBSERVAÇÕES)
      dados.contrato_enviado || '',
      dados.contrato_assinado || '',
      dados.ativacao || '', // ✅ ATIVAÇÃO
      dados.link || '', // ✅ LINK
      mensalidadeNumero, // ✅ MENSALIDADE
      dados.tarifa || '', // ✅ TARIFA
      dados.percentual_tarifa || '', // ✅ % TARIFA
      adesaoNumero, // ✅ ADESÃO
      dados.situacao || 'Novo registro' // ✅ SITUAÇÃO
    ];

    console.log("📝 Atualizando linha:", linhaAtualizar);
    console.log("📊 Novos dados:", novosDados);
    
    aba.getRange(linhaAtualizar, 1, 1, novosDados.length).setValues([novosDados]);
    
    // Formatar colunas monetárias
    aba.getRange(linhaAtualizar, 13).setNumberFormat('"R$"#,##0.00'); // Mensalidade
    aba.getRange(linhaAtualizar, 16).setNumberFormat('"R$"#,##0.00'); // Adesão

    return { 
      success: true, 
      message: `✅ "${dados.razao_social}" atualizado com sucesso!` 
    };

  } catch (error) {
    console.error("❌ Erro em atualizarCadastro:", error);
    return { success: false, message: "Erro ao atualizar: " + error.message };
  }
}

// 🔥 FUNÇÕES PARA PROCESSAR ADESÃO
function processarAdesao(valorAdesao) {
  if (!valorAdesao) return 'Isento';
  const valorStr = valorAdesao.toString().trim();
  if (valorStr === 'Isento' || valorStr === '0' || valorStr === '0.00' || valorStr === 'R$ 0,00') {
    return 'Isento';
  }
  const numero = parseFloat(valorStr);
  if (!isNaN(numero)) {
    return formatarMoedaParaExibicao(numero);
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

function formatarMoedaParaExibicao(valor) {
  if (!valor && valor !== 0) return 'R$ 0,00';
  const numero = typeof valor === 'number' ? valor : parseFloat(valor);
  if (isNaN(numero)) return 'R$ 0,00';
  return 'R$ ' + numero.toLocaleString('pt-BR', { minimumFractionDigits: 2 });
}

// FUNÇÕES AUXILIARES
function formatarCNPJ(cnpj) {
  if (!cnpj) return '';
  const cnpjLimpo = cnpj.toString().replace(/\D/g, '');
  if (cnpjLimpo.length === 14) {
    return cnpjLimpo.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, '$1.$2.$3/$4-$5');
  }
  return cnpjLimpo;
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

// FUNÇÕES EXISTENTES
function buscarCadastrosPorSituacao(situacao) {
  try {
    console.log("🔍 Filtrando por situação:", situacao);
    const todosCadastros = buscarTodosCadastros();
    if (situacao === 'all') return todosCadastros;
    const cadastrosFiltrados = todosCadastros.filter(cadastro => 
      cadastro.situacao === situacao
    );
    console.log("✅ Cadastros filtrados:", cadastrosFiltrados.length);
    return cadastrosFiltrados;
  } catch (error) {
    console.error("❌ Erro em buscarCadastrosPorSituacao:", error);
    return [];
  }
}

function testar() {
  return { 
    success: true, 
    message: "✅ Sistema funcionando!",
    timestamp: new Date().toISOString(),
    totalCadastros: buscarTodosCadastros().length
  };
}
