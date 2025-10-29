// CONFIGURAÇÕES
const CONFIG = {
  ID_PLANILHA: "1V4iGN14UpIQcwf3qKU0_Wbiy2exdW2WUmrYTniy0upA",
  ABA_PRINCIPAL: "Result"
};

// 🔥🔥🔥 CONFIGURAÇÕES DOS WAITLABELS
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

// 🔥🔥🔥 FUNÇÕES DE GERENCIAMENTO DE WAITLABELS
function getWaitlabelAtual() {
  const cache = CacheService.getScriptCache();
  const waitlabelAtual = cache.get('waitlabel_atual');
  return waitlabelAtual || WAITLABELS_CONFIG.WAITLABEL_PADRAO;
}

function setWaitlabelAtual(waitlabel) {
  if (WAITLABELS_CONFIG.WAITLABELS.includes(waitlabel)) {
    const cache = CacheService.getScriptCache();
    cache.put('waitlabel_atual', waitlabel, 21600); // 6 horas
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

// 🔥🔥🔥 FUNÇÃO PRINCIPAL
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistema RESULT - Gestão de Cadastros')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 🔥🔥🔥 FUNÇÕES PRINCIPAIS COM WAITLABEL
function processarCadastroComWaitlabel(dados, waitlabel) {
  try {
    console.log("🎯 PROCESSAR CADASTRO COM WAITLABEL - Dados recebidos:", dados, "Waitlabel:", waitlabel);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    let aba = ss.getSheetByName(waitlabel);

    if (!aba) {
      console.log("📝 Criando nova aba para waitlabel:", waitlabel);
      aba = ss.insertSheet(waitlabel);
      const cabecalho = [
        'Razão Social', 'Nome Fantasia', 'CNPJ', 'Tipo', 'Fornecedor', 
        'Ultimo evento', 'Evento', 'Observação', 'Contrato Enviado', 'Contrato Assinado',
        'Ativação', 'Link', 'Mensalidade', 'Tarifa', '% Tarifa', 'Adesão', 'Situação'
      ];
      aba.getRange('A1:Q1').setValues([cabecalho]);
      aba.getRange(1, 1, 1, cabecalho.length)
        .setBackground(WAITLABELS_CONFIG.CORES[waitlabel] || "#7E3E9A")
        .setFontColor("white")
        .setFontWeight("bold");
      aba.setFrozenRows(1);
    }

    if (dados.acao === 'cadastrar') {
      return cadastrarNovoComWaitlabel(aba, dados, waitlabel);
    } else if (dados.acao === 'atualizar') {
      return atualizarCadastroComWaitlabel(aba, dados, waitlabel);
    } else {
      return { success: false, message: "Ação não reconhecida" };
    }

  } catch (error) {
    console.error("❌ Erro em processarCadastroComWaitlabel:", error);
    return { success: false, message: "Erro: " + error.message };
  }
}

function cadastrarNovoComWaitlabel(aba, dados, waitlabel) {
  try {
    console.log("🆕 CADASTRAR NOVO COM WAITLABEL - INICIANDO");
    console.log("📋 Dados recebidos:", dados);
    console.log("🏷️ Waitlabel:", waitlabel);
    
    // ✅ NOVA VERIFICAÇÃO: Verificar se já existe MESMO CNPJ + MESMO FORNECEDOR
    const fornecedoresParaCadastrar = dados.fornecedores || [];
    const fornecedoresDuplicados = [];
    
    // Buscar todos os cadastros existentes deste CNPJ NO WAITLABEL ATUAL
    const cadastrosExistentes = buscarTodosCadastrosPorCNPJComWaitlabel(dados.cnpj, waitlabel);
    
    for (let fornecedor of fornecedoresParaCadastrar) {
      const nomeFornecedor = fornecedor.nome || fornecedor;
      
      // Verificar se já existe este CNPJ + este fornecedor
      const jaExiste = cadastrosExistentes.some(cad => 
        cad.fornecedor === nomeFornecedor
      );
      
      if (jaExiste) {
        fornecedoresDuplicados.push(nomeFornecedor);
      }
    }
    
    // Se há fornecedores duplicados, avisar
    if (fornecedoresDuplicados.length > 0) {
      return { 
        success: false, 
        message: `❌ Este CNPJ já possui cadastro no ${waitlabel} para o(s) fornecedor(es): ${fornecedoresDuplicados.join(', ')}` 
      };
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

      // 🔥🔥🔥 CORREÇÃO: Datas - USAR DATA DO USUÁRIO SE INFORMADA, SENÃO VAZIO
      const dataAtual = new Date();
      const dataUltimoEvento = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

      // ✅ CORREÇÃO: Usar data informada pelo usuário OU ficar vazio (CORRIGIDO FUSO HORÁRIO)
      let dataAtivacaoParaSalvar = '';
      if (dados.ativacao && dados.ativacao.trim() !== '') {
        // Se usuário informou data, formatar corretamente (CORREÇÃO FUSO HORÁRIO)
        try {
          // 🔥 CORREÇÃO: Adicionar 1 dia para compensar o fuso horário
          const dataUsuario = new Date(dados.ativacao);
          dataUsuario.setDate(dataUsuario.getDate() + 1); // 🔥 ADICIONA 1 DIA
          dataAtivacaoParaSalvar = Utilities.formatDate(dataUsuario, Session.getScriptTimeZone(), "dd/MM/yyyy");
          console.log("📅 Data ativação informada pelo usuário (CORRIGIDA):", dataAtivacaoParaSalvar);
        } catch (e) {
          console.error("❌ Erro ao processar data do usuário:", e);
          dataAtivacaoParaSalvar = ''; // Manter vazio se houver erro
        }
      } else {
        console.log("📅 Nenhuma data de ativação informada - campo ficará vazio");
      }

      console.log(`📅 Datas geradas para fornecedor ${i + 1}:`);
      console.log(`   Data Ativação: ${dataAtivacaoParaSalvar}`);
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
        // 🔥 DATA ATIVAÇÃO - usar a data informada pelo usuário (pode ser vazia)
        dataAtivacaoParaSalvar,
        dados.link || '',
        mensalidadeNumero,
        tarifaFornecedor || '',
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
      mensagem = `✅ "${dados.razao_social}" cadastrado com sucesso no ${waitlabel} para ${sucessos} fornecedor(es)!`;
    } else if (sucessos === 0) {
      mensagem = `❌ Erro ao cadastrar "${dados.razao_social}" no ${waitlabel} para todos os fornecedores!`;
    } else {
      mensagem = `⚠️ "${dados.razao_social}" cadastrado parcialmente no ${waitlabel}: ${sucessos} sucesso(s), ${erros} erro(s)`;
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

function atualizarCadastroComWaitlabel(aba, dados, waitlabel) {
  try {
    console.log("✏️ ATUALIZAR CADASTRO COM WAITLABEL - INICIANDO");
    console.log("📋 Dados recebidos:", dados);
    console.log("🏷️ Waitlabel:", waitlabel);
    
    // 🔥🔥🔥 ADICIONE ESTES DEBUGS PARA A ADESÃO
    console.log("💰💰💰 DEBUG ADESÃO - VALOR RECEBIDO DO HTML:", dados.adesao);
    console.log("💰💰💰 DEBUG ADESÃO - TIPO:", typeof dados.adesao);
    
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
      message: `✅ "${dados.razao_social}" atualizado com sucesso no ${waitlabel}!` 
    };

  } catch (error) {
    console.error("❌ Erro em atualizarCadastroComWaitlabel:", error);
    return { success: false, message: "Erro ao atualizar: " + error.message };
  }
}

// 🔥🔥🔥 FUNÇÕES DE BUSCA COM WAITLABEL
function buscarTodosCadastrosComWaitlabel(waitlabel) {
  try {
    console.log("🔍 Iniciando busca de todos os cadastros no waitlabel:", waitlabel);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabel);
    if (!aba) {
      console.log("❌ Aba não encontrada:", waitlabel);
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
        situacao: (linha[16]?.toString().trim() || 'Novo registro'), // Q - Situação
        waitlabel: waitlabel // 🔥 ADICIONAR WAITLABEL
      };
      
      cadastros.push(cadastro);
    }
    
    console.log("✅ Cadastros processados no", waitlabel + ":", cadastros.length);
    return cadastros;
    
  } catch (error) {
    console.error("❌ Erro em buscarTodosCadastrosComWaitlabel:", error);
    return [];
  }
}

function buscarTodosCadastrosPorCNPJComWaitlabel(cnpj, waitlabel) {
  try {
    console.log("🔍 Buscando TODOS os cadastros do CNPJ:", cnpj, "no waitlabel:", waitlabel);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabel);
    if (!aba) return [];
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) return [];
    
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    
    const cadastrosEncontrados = [];
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      const cnpjCadastro = linha[2]?.toString().replace(/\D/g, '') || '';
      
      // Pular linhas vazias
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      if (cnpjCadastro === cnpjBuscado) {
        cadastrosEncontrados.push({
          id: i + 2,
          fornecedor: linha[4]?.toString().trim() || '',
          situacao: linha[16]?.toString().trim() || '',
          waitlabel: waitlabel
        });
      }
    }
    
    console.log(`✅ Encontrados ${cadastrosEncontrados.length} cadastros para o CNPJ no ${waitlabel}`);
    return cadastrosEncontrados;
    
  } catch (error) {
    console.error("❌ Erro em buscarTodosCadastrosPorCNPJComWaitlabel:", error);
    return [];
  }
}

// 🔥🔥🔥 FUNÇÃO AUXILIAR PARA PROCESSAR LINHAS (CRÍTICA - FALTANTE)
function processarLinhaParaRetorno(linha, id) {
  // Formatar último evento
  let ultimoEventoFormatado = '';
  if (linha[5] && linha[5] instanceof Date) {
    ultimoEventoFormatado = Utilities.formatDate(linha[5], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  } else if (linha[5]) {
    ultimoEventoFormatado = linha[5].toString();
  }
  
  // Formatar data ativação
  let ativacaoFormatada = '';
  if (linha[10] && linha[10] instanceof Date) {
    ativacaoFormatada = Utilities.formatDate(linha[10], Session.getScriptTimeZone(), "yyyy-MM-dd");
  } else if (linha[10]) {
    if (linha[10].includes('/')) {
      const partes = linha[10].split('/');
      ativacaoFormatada = `${partes[2]}-${partes[1]}-${partes[0]}`;
    } else {
      ativacaoFormatada = linha[10].toString();
    }
  }

  // Processar tarifa e percentual
  let tarifa = linha[13]?.toString().trim() || '';
  let percentualTarifa = '0%';
  if (linha[14] !== null && linha[14] !== undefined && linha[14] !== '') {
    const valor = parseFloat(linha[14]);
    if (!isNaN(valor)) {
      percentualTarifa = Math.round(valor * 100) + '%';
    } else {
      percentualTarifa = linha[14]?.toString().trim() || '0%';
    }
  }
  
  // Estrutura de fornecedor para formulário
  const fornecedorParaFormulario = {
    nome: linha[4]?.toString().trim() || '',
    tarifa: tarifa,
    percentual_tarifa: percentualTarifa
  };
  
  return {
    encontrado: true,
    id: id,
    razao_social: linha[0]?.toString().trim() || '',
    nome_fantasia: linha[1]?.toString().trim() || '',
    cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''),
    tipo: linha[3]?.toString().trim() || '',
    fornecedor: linha[4]?.toString().trim() || '',
    fornecedores: [fornecedorParaFormulario],
    ultimo_evento: ultimoEventoFormatado,
    evento: linha[6]?.toString().trim() || '',
    observacoes: linha[7]?.toString().trim() || '',
    contrato_enviado: linha[8]?.toString().trim() || '',
    contrato_assinado: linha[9]?.toString().trim() || '',
    ativacao: ativacaoFormatada,
    link: linha[11]?.toString().trim() || '',
    mensalidade: parseFloat(linha[12]) || 0,
    tarifa: tarifa,
    percentual_tarifa: percentualTarifa,
    adesao: processarAdesao(linha[15]),
    situacao: (linha[16]?.toString().trim() || 'Novo registro')
  };
}

function buscarCadastroPorIDComWaitlabel(id, waitlabel) {
  try {
    console.log("🔍 Buscando cadastro por ID:", id, "no waitlabel:", waitlabel);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabel);
    if (!aba) return { encontrado: false, mensagem: "Waitlabel não encontrado" };
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < id) return { encontrado: false, mensagem: "Registro não encontrado" };
    
    const linha = aba.getRange(id, 1, 1, 17).getValues()[0];
    
    if (!linha[0] || linha[0].toString().trim() === '') {
      return { encontrado: false, mensagem: "Registro vazio ou não encontrado" };
    }
    
    const resultado = processarLinhaParaRetorno(linha, id);
    resultado.waitlabel = waitlabel;
    
    return resultado;
    
  } catch (error) {
    console.error("❌ Erro em buscarCadastroPorIDComWaitlabel:", error);
    return { encontrado: false, mensagem: "Erro: " + error.message };
  }
}

// 🔥🔥🔥 FUNÇÕES PARA "APLICAR A TODOS"
function aplicarAlteracoesATodos(cnpj, dadosParaAplicar, camposSelecionados) {
  try {
    console.log("🎯 APLICAR A TODOS - INICIANDO");
    console.log("📋 CNPJ alvo:", cnpj);
    console.log("📦 Dados para aplicar:", dadosParaAplicar);
    console.log("🔧 Campos selecionados:", camposSelecionados);
    
    const waitlabelAtual = getWaitlabelAtual();
    console.log("🏷️ Waitlabel atual:", waitlabelAtual);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabelAtual);
    
    if (!aba) {
      return { success: false, message: "Waitlabel não encontrado: " + waitlabelAtual };
    }
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) {
      return { success: false, message: "Nenhum cadastro encontrado" };
    }
    
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    
    let registrosAtualizados = 0;
    const resultados = [];
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      const linhaNumero = i + 2;
      
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      const cnpjCadastro = linha[2]?.toString().replace(/\D/g, '') || '';
      
      if (cnpjCadastro === cnpjBuscado) {
        console.log(`🔍 Encontrado registro na linha ${linhaNumero} para aplicar alterações`);
        
        const novosDados = [...linha];
        
        camposSelecionados.forEach(campo => {
          const indiceColuna = obterIndiceColuna(campo);
          if (indiceColuna !== -1) {
            const novoValor = obterValorParaCampo(campo, dadosParaAplicar, linha);
            novosDados[indiceColuna] = novoValor;
            console.log(`   ✅ Campo "${campo}" [coluna ${indiceColuna + 1}]: "${novoValor}"`);
          }
        });
        
        novosDados[5] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        
        try {
          aba.getRange(linhaNumero, 1, 1, novosDados.length).setValues([novosDados]);
          aplicarFormatacao(aba, linhaNumero, camposSelecionados);
          
          registrosAtualizados++;
          resultados.push(`✅ Linha ${linhaNumero} - ${linha[4]}`);
          
        } catch (erroSalvamento) {
          console.error(`❌ Erro ao salvar linha ${linhaNumero}:`, erroSalvamento);
          resultados.push(`❌ Linha ${linhaNumero} - ERRO: ${erroSalvamento.message}`);
        }
      }
    }
    
    SpreadsheetApp.flush();
    
    console.log(`✅ CONCLUSÃO: ${registrosAtualizados} registro(s) atualizado(s)`);
    
    return {
      success: true,
      registrosAtualizados: registrosAtualizados,
      message: `✅ Alterações aplicadas para ${registrosAtualizados} registro(s) do CNPJ ${cnpj}`,
      detalhes: resultados
    };
    
  } catch (error) {
    console.error("❌ Erro em aplicarAlteracoesATodos:", error);
    return { 
      success: false, 
      message: "Erro ao aplicar alterações: " + error.message 
    };
  }
}

function obterIndiceColuna(campo) {
  const mapeamentoCampos = {
    'razao_social': 0,
    'nome_fantasia': 1,  
    'cnpj': 2,
    'tipo': 3,
    'fornecedores': 4,
    'evento': 6,
    'observacoes': 7,
    'contrato_enviado': 8,
    'contrato_assinado': 9,
    'ativacao': 10,
    'link': 11,
    'mensalidade': 12,
    'adesao': 15,
    'situacao': 16
  };
  
  return mapeamentoCampos[campo] !== undefined ? mapeamentoCampos[campo] : -1;
}

function obterValorParaCampo(campo, dadosParaAplicar, linhaAtual) {
  switch(campo) {
    case 'razao_social':
      return normalizarTexto(dadosParaAplicar.razao_social) || '';
    case 'nome_fantasia':
      return normalizarTexto(dadosParaAplicar.nome_fantasia) || '';
    case 'cnpj':
      return dadosParaAplicar.cnpj ? dadosParaAplicar.cnpj.toString() : '';
    case 'tipo':
      return normalizarTexto(dadosParaAplicar.tipo) || '';
    case 'evento':
      return normalizarTexto(dadosParaAplicar.evento) || '';
    case 'observacoes':
      return normalizarTexto(dadosParaAplicar.observacoes) || '';
    case 'contrato_enviado':
      return normalizarTexto(dadosParaAplicar.contrato_enviado) || '';
    case 'contrato_assinado':
      return normalizarTexto(dadosParaAplicar.contrato_assinado) || '';
    case 'ativacao':
      if (dadosParaAplicar.ativacao && dadosParaAplicar.ativacao.trim() !== '') {
        try {
          const dataUsuario = new Date(dadosParaAplicar.ativacao);
          dataUsuario.setDate(dataUsuario.getDate() + 1);
          return Utilities.formatDate(dataUsuario, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } catch (e) {
          console.error("❌ Erro ao processar data:", e);
          return '';
        }
      }
      return '';
    case 'link':
      return dadosParaAplicar.link || '';
    case 'mensalidade':
      return converterMoedaParaNumero(dadosParaAplicar.mensalidade) || 0;
    case 'adesao':
      return processarAdesaoParaSalvar(dadosParaAplicar.adesao);
    case 'situacao':
      let situacao = normalizarTexto(dadosParaAplicar.situacao) || 'NOVO REGISTRO';
      if (situacao === 'NOVO REGISTRO') situacao = 'Novo Registro';
      return situacao;
    case 'fornecedores':
      return linhaAtual[4];
    default:
      return linhaAtual[obterIndiceColuna(campo)];
  }
}

function aplicarFormatacao(aba, linhaNumero, camposSelecionados) {
  try {
    aba.getRange(linhaNumero, 13).setNumberFormat('"R$"#,##0.00');
    aba.getRange(linhaNumero, 16).setNumberFormat('"R$"#,##0.00');
    aba.getRange(linhaNumero, 15).setNumberFormat('0%');
    aba.getRange(linhaNumero, 11).setNumberFormat('dd/MM/yyyy');
    
    if (camposSelecionados.includes('mensalidade')) {
      aba.getRange(linhaNumero, 13).setNumberFormat('"R$"#,##0.00');
    }
    
    if (camposSelecionados.includes('adesao')) {
      aba.getRange(linhaNumero, 16).setNumberFormat('"R$"#,##0.00');
    }
    
  } catch (error) {
    console.error("❌ Erro na formatação:", error);
  }
}

function excluirTodosFornecedoresCNPJ(cnpj) {
  try {
    console.log("🗑️ EXCLUIR TODOS - INICIANDO para CNPJ:", cnpj);
    
    const waitlabelAtual = getWaitlabelAtual();
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabelAtual);
    
    if (!aba) {
      return { success: false, message: "Waitlabel não encontrado" };
    }
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) {
      return { success: false, message: "Nenhum cadastro encontrado" };
    }
    
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    
    const linhasParaExcluir = [];
    
    for (let i = dados.length - 1; i >= 0; i--) {
      const linha = dados[i];
      const cnpjCadastro = linha[2]?.toString().replace(/\D/g, '') || '';
      
      if (cnpjCadastro === cnpjBuscado) {
        linhasParaExcluir.push(i + 2);
      }
    }
    
    console.log(`🔍 Encontradas ${linhasParaExcluir.length} linhas para excluir`);
    
    linhasParaExcluir.forEach(linha => {
      try {
        aba.deleteRow(linha);
        console.log(`✅ Linha ${linha} excluída`);
      } catch (erroExclusao) {
        console.error(`❌ Erro ao excluir linha ${linha}:`, erroExclusao);
      }
    });
    
    return {
      success: true,
      message: `✅ ${linhasParaExcluir.length} registro(s) excluído(s) do CNPJ ${cnpj}`,
      registrosExcluidos: linhasParaExcluir.length
    };
    
  } catch (error) {
    console.error("❌ Erro em excluirTodosFornecedoresCNPJ:", error);
    return { 
      success: false, 
      message: "Erro ao excluir registros: " + error.message 
    };
  }
}

function contarRegistrosPorCNPJ(cnpj) {
  try {
    const waitlabelAtual = getWaitlabelAtual();
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabelAtual);
    
    if (!aba) return 0;
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) return 0;
    
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    
    let contador = 0;
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      const cnpjCadastro = linha[2]?.toString().replace(/\D/g, '') || '';
      
      if (cnpjCadastro === cnpjBuscado) {
        contador++;
      }
    }
    
    console.log(`🔍 CNPJ ${cnpj} tem ${contador} registro(s)`);
    return contador;
    
  } catch (error) {
    console.error("❌ Erro em contarRegistrosPorCNPJ:", error);
    return 0;
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
  console.log("💰💰💰 processarAdesaoParaSalvar - VALOR ENTRADA:", valorAdesao, "Tipo:", typeof valorAdesao);
  
  if (!valorAdesao && valorAdesao !== 0) {
    console.log("💰💰💰 Retornando 0 (valor vazio)");
    return 0;
  }
  
  // Se já é número, retorna direto (SEM multiplicar)
  if (typeof valorAdesao === 'number') {
    console.log("💰💰💰 Já é número, retornando:", valorAdesao);
    return valorAdesao;
  }
  
  const valorStr = valorAdesao.toString().trim();
  console.log("💰💰💰 Valor como string:", valorStr);
  
  if (valorStr === 'Isento' || valorStr === '0' || valorStr === '0.00' || valorStr === 'R$ 0,00') {
    console.log("💰💰💰 Retornando 0 (isento)");
    return 0;
  }
  
  // 🔥🔥🔥 CORREÇÃO: Converter sem multiplicações
  try {
    const valorLimpo = valorStr
      .replace('R$', '')
      .replace(/\./g, '')
      .replace(',', '.')
      .trim();
    
    console.log("💰💰💰 Valor limpo:", valorLimpo);
    
    const numero = parseFloat(valorLimpo);
    
    if (isNaN(numero)) {
      console.log("💰💰💰 Não é número válido, retornando 0");
      return 0;
    }
    
    console.log("💰💰💰 Número final para salvar:", numero);
    return numero;
    
  } catch (error) {
    console.error("💰💰💰 Erro ao processar adesão:", error);
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

function normalizarTexto(texto) {
  if (!texto || typeof texto !== 'string') return texto;
  return texto
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
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

// 🔥🔥🔥 FUNÇÕES ORIGINAIS (PARA COMPATIBILIDADE)
function processarCadastro(dados) {
  try {
    console.log("🎯 PROCESSAR CADASTRO - Dados recebidos:", dados);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    let aba = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);

    if (!aba) {
      console.log("📝 Criando nova aba...");
      aba = ss.insertSheet(CONFIG.ABA_PRINCIPAL);
      const cabecalho = [
        'Razão Social', 'Nome Fantasia', 'CNPJ', 'Tipo', 'Fornecedor', 
        'Ultimo evento', 'Evento', 'Observação', 'Contrato Enviado', 'Contrato Assinado',
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

function cadastrarNovo(aba, dados) {
  try {
    console.log("🆕 CADASTRAR NOVO - INICIANDO COM DEBUG");
    console.log("📋 Dados recebidos:", dados);
    
    // ✅ NOVA VERIFICAÇÃO: Verificar se já existe MESMO CNPJ + MESMO FORNECEDOR
    const fornecedoresParaCadastrar = dados.fornecedores || [];
    const fornecedoresDuplicados = [];
    
    // Buscar todos os cadastros existentes deste CNPJ
    const cadastrosExistentes = buscarTodosCadastrosPorCNPJ(dados.cnpj);
    
    for (let fornecedor of fornecedoresParaCadastrar) {
      const nomeFornecedor = fornecedor.nome || fornecedor;
      
      // Verificar se já existe este CNPJ + este fornecedor
      const jaExiste = cadastrosExistentes.some(cad => 
        cad.fornecedor === nomeFornecedor
      );
      
      if (jaExiste) {
        fornecedoresDuplicados.push(nomeFornecedor);
      }
    }
    
    // Se há fornecedores duplicados, avisar
    if (fornecedoresDuplicados.length > 0) {
      return { 
        success: false, 
        message: `❌ Este CNPJ já possui cadastro para o(s) fornecedor(es): ${fornecedoresDuplicados.join(', ')}` 
      };
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

      // 🔥🔥🔥 CORREÇÃO: Datas - USAR DATA DO USUÁRIO SE INFORMADA, SENÃO VAZIO
      const dataAtual = new Date();
      const dataUltimoEvento = Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

      // ✅ CORREÇÃO: Usar data informada pelo usuário OU ficar vazio (CORRIGIDO FUSO HORÁRIO)
      let dataAtivacaoParaSalvar = '';
      if (dados.ativacao && dados.ativacao.trim() !== '') {
        // Se usuário informou data, formatar corretamente (CORREÇÃO FUSO HORÁRIO)
        try {
          // 🔥 CORREÇÃO: Adicionar 1 dia para compensar o fuso horário
          const dataUsuario = new Date(dados.ativacao);
          dataUsuario.setDate(dataUsuario.getDate() + 1); // 🔥 ADICIONA 1 DIA
          dataAtivacaoParaSalvar = Utilities.formatDate(dataUsuario, Session.getScriptTimeZone(), "dd/MM/yyyy");
          console.log("📅 Data ativação informada pelo usuário (CORRIGIDA):", dataAtivacaoParaSalvar);
        } catch (e) {
          console.error("❌ Erro ao processar data do usuário:", e);
          dataAtivacaoParaSalvar = ''; // Manter vazio se houver erro
        }
      } else {
        console.log("📅 Nenhuma data de ativação informada - campo ficará vazio");
      }

      console.log(`📅 Datas geradas para fornecedor ${i + 1}:`);
      console.log(`   Data Ativação: ${dataAtivacaoParaSalvar}`);
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
        // 🔥 DATA ATIVAÇÃO - usar a data informada pelo usuário (pode ser vazia)
        dataAtivacaoParaSalvar,
        dados.link || '',
        mensalidadeNumero,
        tarifaFornecedor || '',
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

function buscarTodosCadastrosPorCNPJ(cnpj) {
  try {
    console.log("🔍 Buscando TODOS os cadastros do CNPJ:", cnpj);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(CONFIG.ABA_PRINCIPAL);
    if (!aba) return [];
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) return [];
    
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    
    const cadastrosEncontrados = [];
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      const cnpjCadastro = linha[2]?.toString().replace(/\D/g, '') || '';
      
      // Pular linhas vazias
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      if (cnpjCadastro === cnpjBuscado) {
        cadastrosEncontrados.push({
          id: i + 2,
          fornecedor: linha[4]?.toString().trim() || '',
          situacao: linha[16]?.toString().trim() || ''
        });
      }
    }
    
    console.log(`✅ Encontrados ${cadastrosEncontrados.length} cadastros para o CNPJ`);
    return cadastrosEncontrados;
    
  } catch (error) {
    console.error("❌ Erro em buscarTodosCadastrosPorCNPJ:", error);
    return [];
  }
}

function salvarCadastro(dados) {
  return processarCadastro(dados);
}

function processarAtualizacao(dados) {
  return processarCadastro(dados);
}

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

function testar() {
  return { 
    success: true, 
    message: "✅ Sistema funcionando!",
    timestamp: new Date().toISOString(),
    totalCadastros: buscarTodosCadastros().length
  };
}
