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
      // 🔥 CORREÇÃO: Cabeçalho com 17 colunas SEM "Data Status"
      const cabecalho = [
        'Razão Social', 'Nome Fantasia', 'CNPJ', 'Tipo', 'Fornecedor', 
        'Ultimo evento', 'Evento', 'Observação', 'Contrato Enviado', 'Contrato Assinado',
        'Ativação', 'Link', 'Mensalidade', 'Tarifa', '% Tarifa', 'Adesão', 'Situação'
      ];
      aba.getRange('A1:Q1').setValues([cabecalho]); // ✅ A:Q (17 colunas)
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

// 🔥 FUNÇÃO DEBUG DETALHADA PARA DATAS
function debugDatas(dados) {
  console.log("🎯 DEBUG DETALHADO - DATAS E TARIFAS");
  console.log("📦 Dados completos:", JSON.stringify(dados, null, 2));
  console.log("📅 Data ativação recebida:", dados.ativacao, "Tipo:", typeof dados.ativacao);
  console.log("💰 Tarifa recebida:", dados.tarifa, "Tipo:", typeof dados.tarifa);
  console.log("📊 Fornecedores:", dados.fornecedores);
  
  if (dados.fornecedores && Array.isArray(dados.fornecedores)) {
    dados.fornecedores.forEach((fornecedor, index) => {
      console.log(`🔍 Fornecedor ${index + 1}:`, fornecedor);
      console.log(`   Nome: ${fornecedor.nome || fornecedor}`);
      console.log(`   Tarifa: ${fornecedor.tarifa || 'N/A'}`);
      console.log(`   % Tarifa: ${fornecedor.percentual_tarifa || 'N/A'}`);
    });
  }
  
  return { success: true, message: "Debug realizado - verifique logs" };
}

// 🔥🔥🔥 FUNÇÃO CADASTRAR NOVO - COM DEBUG MAXIMO
function cadastrarNovo(aba, dados) {
  try {
    console.log("🆕 CADASTRAR NOVO - INICIANDO COM DEBUG");
    console.log("📋 Dados recebidos:", dados);
    
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
    let situacaoParaSalvar = normalizarTexto(dados.situacao) || 'NOVO REGISTRO';
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

      console.log(`🔍 Processando fornecedor ${i + 1}:`);
      console.log(`   Nome: ${nomeFornecedor}`);
      console.log(`   Tarifa: ${tarifaFornecedor}`);
      console.log(`   % Tarifa: ${percentualTarifaFornecedor}`);

      // Validar se o nome do fornecedor está preenchido
      if (!nomeFornecedor || nomeFornecedor.trim() === '') {
        resultados.push(`❌ Fornecedor sem nome - pulado`);
        continue;
      }

      // Converter valores monetários
      let mensalidadeNumero = parseFloat(dados.mensalidade) || 0;
      let adesaoNumero = processarAdesaoParaSalvar(dados.adesao);

      // 🔥🔥🔥 CORREÇÃO: Datas FRESCAS para CADA fornecedor
      const dataAtual = new Date();
      const dataAtivacao = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy");
      const dataUltimoEvento = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

      console.log(`📅 Datas geradas para fornecedor ${i + 1}:`);
      console.log(`   Data Ativação: ${dataAtivacao}`);
      console.log(`   Data Último Evento: ${dataUltimoEvento}`);

      // Array com 17 colunas na ORDEM CORRETA
      const linhaDados = [
        normalizarTexto(dados.razao_social) || '',
        normalizarTexto(dados.nome_fantasia) || '',
        dados.cnpj ? dados.cnpj.toString() : '',
        normalizarTexto(dados.tipo) || '',
        normalizarTexto(nomeFornecedor),
        // Data ÚLTIMO EVENTO
        dataUltimoEvento,
        normalizarTexto(dados.evento) || '',
        normalizarTexto(dados.observacoes) || '',
        normalizarTexto(dados.contrato_enviado) || '',
        normalizarTexto(dados.contrato_assinado) || '',
        // Data ATIVAÇÃO
        dataAtivacao,
        dados.link || '',
        mensalidadeNumero,
        tarifaFornecedor || '', // NÃO aplicar normalizarTexto
        percentualTarifaFornecedor,
        adesaoNumero,
        normalizarTexto(situacaoParaSalvar)
      ];

      console.log(`📝 Linha de dados ${i + 1}:`, linhaDados);
      
      try {
        const range = aba.getRange(linhaInserir, 1, 1, linhaDados.length);
        console.log(`💾 Salvando na linha: ${linhaInserir}`);
        range.setValues([linhaDados]);
        
        // 🔥 FORMATAR COLUNAS IMEDIATAMENTE
        aba.getRange(linhaInserir, 13).setNumberFormat('"R$"#,##0.00'); // Mensalidade (M)
        aba.getRange(linhaInserir, 16).setNumberFormat('"R$"#,##0.00'); // Adesão (P)
        aba.getRange(linhaInserir, 15).setNumberFormat('0%'); // % Tarifa (O)
        aba.getRange(linhaInserir, 14).setNumberFormat('@'); // Tarifa como texto (N)
        aba.getRange(linhaInserir, 11).setNumberFormat('dd/MM/yyyy'); // 🔥 FORMATAR DATA ATIVAÇÃO (K)
        
        SpreadsheetApp.flush();
        
        // 🔥 VERIFICAR O QUE FOI SALVO
        const dadosSalvos = aba.getRange(linhaInserir, 1, 1, 17).getValues()[0];
        console.log(`✅ Dados salvos na linha ${linhaInserir}:`, dadosSalvos);
        console.log(`📅 Data ativação salva: ${dadosSalvos[10]}`);
        console.log(`💰 Tarifa salva: ${dadosSalvos[13]}`);
        console.log(`📊 % Tarifa salva: ${dadosSalvos[14]}`);
        
        linhaInserir++;
        registrosCriados++;
        resultados.push(`✅ ${nomeFornecedor} - ${tarifaFornecedor} ${percentualTarifaFornecedor}`);
        
      } catch (erroInsercao) {
        console.error(`❌ Erro ao salvar:`, erroInsercao);
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
    console.error("❌ Erro geral:", error);
    return { 
      success: false, 
      message: "Erro ao cadastrar: " + error.message 
    };
  }
}

// 🔥🔥🔥 FUNÇÃO ATUALIZAR CADASTRO - CORRIGIDA (DATA ATIVAÇÃO NÃO MUDA)
function atualizarCadastro(aba, dados) {
  try {
    console.log("✏️ ATUALIZAR CADASTRO - INICIANDO");
    console.log("📋 Dados recebidos:", dados);
    
    const linhaAtualizar = parseInt(dados.id);

    if (linhaAtualizar < 2 || linhaAtualizar > aba.getLastRow()) {
      return { success: false, message: "Registro não encontrado" };
    }

    // 🔥🔥🔥 CORREÇÃO 1: BUSCAR A DATA DE ATIVAÇÃO ORIGINAL
    const dadosAtuais = aba.getRange(linhaAtualizar, 1, 1, 17).getValues()[0];
    const dataAtivacaoOriginal = dadosAtuais[10]; // Coluna K - Ativação
    
    console.log("📅 Data ativação original:", dataAtivacaoOriginal);
    console.log("📅 Tipo da data original:", typeof dataAtivacaoOriginal);

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

    // 🔥🔥🔥 CORREÇÃO 2: MANTER A DATA DE ATIVAÇÃO ORIGINAL
    let dataAtivacaoParaSalvar = dataAtivacaoOriginal;
    
    // Se for um objeto Date, formatar corretamente
    if (dataAtivacaoOriginal instanceof Date) {
      dataAtivacaoParaSalvar = Utilities.formatDate(dataAtivacaoOriginal, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    // Se já for string, manter como está
    else if (typeof dataAtivacaoOriginal === 'string') {
      dataAtivacaoParaSalvar = dataAtivacaoOriginal;
    }
    // Se estiver vazia, usar a data atual (apenas para novos registros)
    else if (!dataAtivacaoOriginal || dataAtivacaoOriginal === '') {
      dataAtivacaoParaSalvar = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    }

    console.log("📅 Data ativação que será salva:", dataAtivacaoParaSalvar);

    // Array com 17 colunas na ORDEM CORRETA
    const novosDados = [
      normalizarTexto(dados.razao_social) || '',
      normalizarTexto(dados.nome_fantasia) || '',
      dados.cnpj ? dados.cnpj.toString() : '',
      normalizarTexto(dados.tipo) || '',
      normalizarTexto(fornecedorParaAtualizar),
      // ✅ Data ÚLTIMO EVENTO atualizada (com segundos)
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
      normalizarTexto(dados.evento) || '',
      normalizarTexto(dados.observacoes) || '',
      normalizarTexto(dados.contrato_enviado) || '',
      normalizarTexto(dados.contrato_assinado) || '',
      // 🔥🔥🔥 DATA ATIVAÇÃO ORIGINAL (NÃO MUDA)
      dataAtivacaoParaSalvar,
      dados.link || '',
      mensalidadeNumero,
      tarifaParaAtualizar || '', // 🔥 NÃO aplicar normalizarTexto
      percentualParaAtualizar,
      adesaoNumero,
      normalizarTexto(situacaoValida)
    ];

    console.log("📝 Atualizando linha:", linhaAtualizar);
    console.log("📊 Novos dados:", novosDados);
    
    aba.getRange(linhaAtualizar, 1, 1, novosDados.length).setValues([novosDados]);
    
    // 🔥🔥🔥 CORREÇÃO: ADICIONAR FORMATAÇÃO DA TARIFA
    aba.getRange(linhaAtualizar, 13).setNumberFormat('"R$"#,##0.00'); // Mensalidade (coluna M)
    aba.getRange(linhaAtualizar, 16).setNumberFormat('"R$"#,##0.00'); // Adesão (coluna P)
    aba.getRange(linhaAtualizar, 15).setNumberFormat('0%'); // % Tarifa (coluna O)
    aba.getRange(linhaAtualizar, 14).setNumberFormat('@'); // 🔥 Tarifa como texto (coluna N)

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

// 🔥 CORREÇÃO DEFINITIVA - FUNÇÃO BUSCAR CADASTRO POR ID (17 COLUNAS)
function buscarCadastroPorID(id) {
  try {
    console.log("🔍 Buscando cadastro por ID:", id);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);
    if (!aba) return { encontrado: false, mensagem: "Planilha não encontrada" };
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < id) return { encontrado: false, mensagem: "Registro não encontrado" };
    
    const linha = aba.getRange(id, 1, 1, 17).getValues()[0]; // ✅ 17 colunas
    
    // Verificar se a linha não está vazia
    if (!linha[0] || linha[0].toString().trim() === '') {
      return { encontrado: false, mensagem: "Registro vazio ou não encontrado" };
    }
    
    console.log("📊 Linha bruta encontrada:", linha);
    
    // 🔥 CORREÇÃO: ÍNDICES CORRETOS PARA 17 COLUNAS
    let ultimoEventoFormatado = '';
    if (linha[5] && linha[5] instanceof Date) { // ✅ CORRETO: linha[5] é Último evento
      ultimoEventoFormatado = Utilities.formatDate(linha[5], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    } else if (linha[5]) {
      ultimoEventoFormatado = linha[5].toString();
    }
    
    let ativacaoFormatada = '';
    if (linha[10] && linha[10] instanceof Date) { // ✅ CORRETO: linha[10] é Ativação
      ativacaoFormatada = Utilities.formatDate(linha[10], Session.getScriptTimeZone(), "yyyy-MM-dd"); // 🔥 FORMATO PARA INPUT DATE
    } else if (linha[10]) {
      // Se já é string, converter de dd/MM/yyyy para yyyy-MM-dd se necessário
      if (linha[10].includes('/')) {
        const partes = linha[10].split('/');
        ativacaoFormatada = `${partes[2]}-${partes[1]}-${partes[0]}`;
      } else {
        ativacaoFormatada = linha[10].toString();
      }
    }

          // 🔥 CORREÇÃO: Processar tarifa e percentual corretamente
      let tarifa = linha[13]?.toString().trim() || '';

      // 🔥🔥🔥 CORREÇÃO CRÍTICA: Converter número para porcentagem
      let percentualTarifa = '0%';
      if (linha[14] !== null && linha[14] !== undefined && linha[14] !== '') {
        const valor = parseFloat(linha[14]);
        if (!isNaN(valor)) {
          // Converter 0.07 para 7%
          percentualTarifa = Math.round(valor * 100) + '%';
        } else {
          percentualTarifa = linha[14]?.toString().trim() || '0%';
        }
      }
    
    console.log(`💰 Tarifa encontrada: "${tarifa}"`);
    console.log(`📊 % Tarifa encontrada: "${percentualTarifa}"`);
    console.log(`📅 Ativação encontrada: "${linha[10]}" → Formatada: "${ativacaoFormatada}"`);
    
    // 🔥🔥🔥 CORREÇÃO CRÍTICA: Estrutura de fornecedores para o formulário
    const fornecedorParaFormulario = {
      nome: linha[4]?.toString().trim() || '', // E - Fornecedor
      tarifa: tarifa,                          // N - Tarifa
      percentual_tarifa: percentualTarifa      // O - % Tarifa
    };
    
    console.log("👥 Fornecedor para formulário:", fornecedorParaFormulario);
    console.log("🎯 DEBUG DA SITUAÇÃO:");
    console.log("Coluna 16 (Situação):", linha[16], "Tipo:", typeof linha[16]);
    console.log("Situação como string:", linha[16]?.toString());
    console.log("Situação trimmed:", linha[16]?.toString().trim());
    
    // 🔥 CORREÇÃO: RETORNO COM ÍNDICES CORRETOS PARA 17 COLUNAS
    const resultado = {
      encontrado: true,
      id: id,
      razao_social: linha[0]?.toString().trim() || '',     // A - Razão Social
      nome_fantasia: linha[1]?.toString().trim() || '',    // B - Nome Fantasia
      cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''), // C - CNPJ
      tipo: linha[3]?.toString().trim() || '',             // D - Tipo
      fornecedor: linha[4]?.toString().trim() || '',       // E - Fornecedor (para compatibilidade)
      fornecedores: [fornecedorParaFormulario],            // 🔥 ESTRUTURA QUE O FORMULÁRIO ESPERA
      ultimo_evento: ultimoEventoFormatado,                // F - Último evento
      evento: linha[6]?.toString().trim() || '',           // G - Evento
      observacoes: linha[7]?.toString().trim() || '',      // H - Observação
      contrato_enviado: linha[8]?.toString().trim() || '', // I - Contrato Enviado
      contrato_assinado: linha[9]?.toString().trim() || '', // J - Contrato Assinado
      ativacao: ativacaoFormatada,                         // K - Ativação ⭐
      link: linha[11]?.toString().trim() || '',            // L - Link
      mensalidade: parseFloat(linha[12]) || 0,             // M - Mensalidade
      tarifa: tarifa,                                      // N - Tarifa ⭐ (para compatibilidade)
      percentual_tarifa: percentualTarifa,                 // O - % Tarifa ⭐ (para compatibilidade)
      adesao: processarAdesao(linha[15]),                  // P - Adesão
      situacao: (linha[16]?.toString().trim() || 'Novo registro') // Q - Situação
    };
    
    console.log("✅ Resultado final para formulário:", resultado);
    return resultado;
    
  } catch (error) {
    console.error("❌ Erro em buscarCadastroPorID:", error);
    return { encontrado: false, mensagem: "Erro: " + error.message };
  }
}

// 🔥 ATUALIZE TAMBÉM A FUNÇÃO buscarTodosCadastros() para 17 colunas:
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
      
      // Formatar último evento
      let ultimoEventoFormatado = '';
      if (linha[5] && linha[5] instanceof Date) { // ✅ Último evento
        ultimoEventoFormatado = Utilities.formatDate(linha[5], Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else if (linha[5]) {
        ultimoEventoFormatado = linha[5].toString();
      }
      
      let ativacaoFormatada = '';
      if (linha[10] && linha[10] instanceof Date) { // ✅ Ativação
        ativacaoFormatada = Utilities.formatDate(linha[10], Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else if (linha[10]) {
        ativacaoFormatada = linha[10].toString();
      }
      
      // 🔥 CORREÇÃO: ESTRUTURA COM 17 COLUNAS
      const cadastro = {
        id: i + 2,
        razao_social: linha[0]?.toString().trim() || '',     // A - Razão Social
        nome_fantasia: linha[1]?.toString().trim() || '',    // B - Nome Fantasia
        cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''), // C - CNPJ
        tipo: linha[3]?.toString().trim() || '',             // D - Tipo
        fornecedor: linha[4]?.toString().trim() || '',       // E - Fornecedor
        ultimo_evento: ultimoEventoFormatado,                // F - Último evento
        evento: linha[6]?.toString().trim() || '',           // G - Evento
        observacoes: linha[7]?.toString().trim() || '',      // H - Observação
        contrato_enviado: linha[8]?.toString().trim() || '', // I - Contrato Enviado
        contrato_assinado: linha[9]?.toString().trim() || '', // J - Contrato Assinado
        ativacao: ativacaoFormatada,                         // K - Ativação ⭐
        link: linha[11]?.toString().trim() || '',            // L - Link
        mensalidade: parseFloat(linha[12]) || 0,             // M - Mensalidade
        tarifa: linha[13]?.toString().trim() || '',          // N - Tarifa
        percentual_tarifa: linha[14]?.toString().trim() || '', // O - % Tarifa
        adesao: processarAdesao(linha[15]),                  // P - Adesão
        situacao: (linha[16]?.toString().trim() || 'Novo registro') // Q - Situação
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

// 🔥🔥🔥 FUNÇÃO BUSCAR CADASTRO POR CNPJ - VERSÃO CORRIGIDA (APENAS UMA)
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
    
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues(); // ✅ 17 colunas
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    
    console.log("🔎 Procurando CNPJ limpo:", cnpjBuscado);
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      const cnpjCadastro = linha[2]?.toString().replace(/\D/g, '') || '';
      
      // Pular linhas vazias
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      if (cnpjCadastro === cnpjBuscado) {
        console.log("✅ Cadastro encontrado na linha:", i + 2);

        // 🔥🔥🔥 ADICIONA O DEBUG AQUI
  console.log("🔍 DEBUG DETALHADO DA LINHA ENCONTRADA:");
  console.log("Linha completa:", linha);
  console.log("Coluna 13 (Tarifa):", linha[13], "Tipo:", typeof linha[13]);
  console.log("Coluna 14 (% Tarifa):", linha[14], "Tipo:", typeof linha[14]);
  console.log("Coluna 14 como string:", linha[14]?.toString());
  console.log("Coluna 14 trimmed:", linha[14]?.toString().trim());
        
        // Formatar último evento
        let ultimoEventoFormatado = '';
        if (linha[5] && linha[5] instanceof Date) { // ✅ Último evento
          ultimoEventoFormatado = Utilities.formatDate(linha[5], Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else if (linha[5]) {
          ultimoEventoFormatado = linha[5].toString();
        }
        
        // 🔥 CORREÇÃO: Data ativação para formato do input date
        let ativacaoFormatada = '';
        if (linha[10] && linha[10] instanceof Date) { // ✅ Ativação
          ativacaoFormatada = Utilities.formatDate(linha[10], Session.getScriptTimeZone(), "yyyy-MM-dd"); // 🔥 FORMATO PARA INPUT DATE
        } else if (linha[10]) {
          // Se já é string, converter de dd/MM/yyyy para yyyy-MM-dd se necessário
          if (linha[10].includes('/')) {
            const partes = linha[10].split('/');
            ativacaoFormatada = `${partes[2]}-${partes[1]}-${partes[0]}`;
          } else {
            ativacaoFormatada = linha[10].toString();
          }
        }

        // 🔥 CORREÇÃO: Processar tarifa e percentual corretamente
        // 🔥 CORREÇÃO: Processar tarifa e percentual corretamente
          let tarifa = linha[13]?.toString().trim() || '';

          // 🔥🔥🔥 CORREÇÃO CRÍTICA: Converter número para porcentagem
          let percentualTarifa = '0%';
          if (linha[14] !== null && linha[14] !== undefined && linha[14] !== '') {
            const valor = parseFloat(linha[14]);
            if (!isNaN(valor)) {
              // Converter 0.07 para 7%
              percentualTarifa = Math.round(valor * 100) + '%';
            } else {
              percentualTarifa = linha[14]?.toString().trim() || '0%';
            }
          }
        
        console.log(`💰 Tarifa encontrada: "${tarifa}"`);
        console.log(`📊 % Tarifa encontrada: "${percentualTarifa}"`);
        console.log(`📅 Ativação encontrada: "${linha[10]}" → Formatada: "${ativacaoFormatada}"`);
        
        // 🔥🔥🔥 CORREÇÃO CRÍTICA: Estrutura de fornecedores para o formulário
        const fornecedorParaFormulario = {
          nome: linha[4]?.toString().trim() || '', // E - Fornecedor
          tarifa: tarifa,                          // N - Tarifa
          percentual_tarifa: percentualTarifa      // O - % Tarifa
        };
        
        console.log("👥 Fornecedor para formulário:", fornecedorParaFormulario);

        console.log("🎯 DEBUG DA SITUAÇÃO:");
        console.log("Coluna 16 (Situação):", linha[16], "Tipo:", typeof linha[16]);
        console.log("Situação como string:", linha[16]?.toString());
        console.log("Situação trimmed:", linha[16]?.toString().trim());
        
        // 🔥 CORREÇÃO: ESTRUTURA COM 17 COLUNAS
        return {
          encontrado: true,
          id: i + 2,
          razao_social: linha[0]?.toString().trim() || '',     // A
          nome_fantasia: linha[1]?.toString().trim() || '',    // B
          cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''), // C
          tipo: linha[3]?.toString().trim() || '',             // D
          fornecedor: linha[4]?.toString().trim() || '',       // E
          fornecedores: [fornecedorParaFormulario],            // 🔥 ESTRUTURA QUE O FORMULÁRIO ESPERA
          ultimo_evento: ultimoEventoFormatado,                // F
          evento: linha[6]?.toString().trim() || '',           // G
          observacoes: linha[7]?.toString().trim() || '',      // H
          contrato_enviado: linha[8]?.toString().trim() || '', // I
          contrato_assinado: linha[9]?.toString().trim() || '', // J
          ativacao: ativacaoFormatada,                         // K ⭐
          link: linha[11]?.toString().trim() || '',            // L
          mensalidade: parseFloat(linha[12]) || 0,             // M
          tarifa: tarifa,                                      // N ⭐ (para compatibilidade)
          percentual_tarifa: percentualTarifa,                 // O ⭐ (para compatibilidade)
          adesao: processarAdesao(linha[15]),                  // P
          situacao: (linha[16]?.toString().trim() || 'Novo registro') // Q
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

// 🔥 FUNÇÃO: Converter TUDO para maiúsculas e remover acentos
function normalizarTexto(texto) {
  if (!texto || typeof texto !== 'string') return texto;
  
  // ✅ TUDO vai para maiúsculas, sem exceções
  return texto
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '') // Remove acentos
    .toUpperCase() // Converte para maiúsculas
    .trim();
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
