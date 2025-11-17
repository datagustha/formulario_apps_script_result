// CONFIGURAÇÕES
const CONFIG = {
  ID_PLANILHA: "1V4iGN14UpIQcwf3qKU0_Wbiy2exdW2WUmrYTniy0upA",
  ABA_PRINCIPAL: "Result",
  TIMEZONE: "America/Sao_Paulo" // 🔥 CORREÇÃO: Fuso horário padronizado
};

// 🔥🔥🔥 FUNÇÃO CORRIGIDA PARA HORÁRIO BRASIL
// 🔥🔥🔥 FUNÇÃO CORRIGIDA PARA HORÁRIO BRASIL - VERSÃO DEFINITIVA
function formatarDataBrasil(data) {
  if (!data) return '';
  
  try {
    console.log("🔥 formatarDataBrasil - Entrada:", data, "Tipo:", typeof data);
    
    // Se já é string no formato brasileiro, retornar COMO ESTÁ
    if (typeof data === 'string' && data.includes('/') && data.includes(':')) {
      console.log("✅ Já está no formato brasileiro - retornando como está:", data);
      return data;
    }
    
    // Se é objeto Date, formatar CORRETAMENTE com fuso do Brasil
    if (data instanceof Date) {
      const dataBrasil = Utilities.formatDate(data, CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm:ss");
      console.log("✅ Date convertido:", data.toString(), "→", dataBrasil);
      return dataBrasil;
    }
    
    // Para outros casos, tentar converter mantendo o horário ORIGINAL
    try {
      const dataObj = new Date(data);
      if (!isNaN(dataObj.getTime())) {
        const dataBrasil = Utilities.formatDate(dataObj, CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm:ss");
        console.log("✅ Outro tipo convertido:", data, "→", dataBrasil);
        return dataBrasil;
      }
    } catch (e) {
      console.log("⚠️ Não conseguiu converter, retornando original:", data);
      return data.toString();
    }
    
    // Fallback
    console.log("⚠️ Fallback - retornando como string:", data);
    return data.toString();
    
  } catch (error) {
    console.error("❌ Erro em formatarDataBrasil:", error);
    return data ? data.toString() : '';
  }
}

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
      // NOVA ESTRUTURA SEM "Tipo" - 16 colunas
      const cabecalho = [
        'Razão Social', 'Nome Fantasia', 'CNPJ', 'Fornecedor', 
        'Ultimo evento', 'Evento', 'Observação', 'Contrato Enviado', 'Contrato Assinado',
        'Ativação', 'Link', 'Mensalidade', 'Mensalidade SIM', 'Tarifa', '% Tarifa', 'Adesão', 'Situação'
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
      const dataUltimoEvento = formatarDataBrasil(dataAtual);

      // ✅✅✅ CORREÇÃO: Usar data informada pelo usuário COM +1 DIA
      let dataAtivacaoParaSalvar = '';
      if (dados.ativacao && dados.ativacao.trim() !== '') {
        try {
          const dataUsuario = new Date(dados.ativacao);
          // 🔥🔥🔥 CORREÇÃO: ADICIONAR +1 DIA PARA COMPENSAR FUSO HORÁRIO
          dataUsuario.setDate(dataUsuario.getDate() + 1);
          dataAtivacaoParaSalvar = Utilities.formatDate(dataUsuario, CONFIG.TIMEZONE, "dd/MM/yyyy");
          console.log("📅 Data ativação informada pelo usuário (CORRIGIDA +1):", dataAtivacaoParaSalvar);
        } catch (e) {
          console.error("❌ Erro ao processar data do usuário:", e);
          dataAtivacaoParaSalvar = '';
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
        normalizarTexto(nomeFornecedor),
        dataUltimoEvento,
        normalizarTexto(dados.evento) || '',
        normalizarTexto(dados.observacoes) || '',
        normalizarTexto(dados.contrato_enviado) || '',
        normalizarTexto(dados.contrato_assinado) || '',
        dataAtivacaoParaSalvar,
        dados.link || '',
        mensalidadeNumero,                    
        converterMoedaParaNumero(dados.mensalidade_sim) || 0,
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
        aba.getRange(linhaInserir, 12).setNumberFormat('"R$"#,##0.00'); // L - Mensalidade
        aba.getRange(linhaInserir, 13).setNumberFormat('"R$"#,##0.00'); // M - Mensalidade SIM
        aba.getRange(linhaInserir, 15).setNumberFormat('0.00%');        // O - % Tarifa
        aba.getRange(linhaInserir, 16).setNumberFormat('"R$"#,##0.00'); // P - Adesão
        aba.getRange(linhaInserir, 14).setNumberFormat('@');            // N - Tarifa (texto)
        aba.getRange(linhaInserir, 10).setNumberFormat('dd/MM/yyyy');   // J - Ativação
        
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
    
    const linhaAtualizar = parseInt(dados.id);

    if (linhaAtualizar < 2 || linhaAtualizar > aba.getLastRow()) {
      return { success: false, message: "Registro não encontrado" };
    }

    // 🔥 BUSCAR OS DADOS ATUAIS
    const dadosAtuais = aba.getRange(linhaAtualizar, 1, 1, 17).getValues()[0];
    const dataAtivacaoOriginal = dadosAtuais[9]; // Coluna J - Ativação
    
    console.log("📅 Data ativação original:", dataAtivacaoOriginal);

    // Processar fornecedor
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

    // Converter valores monetários
    let mensalidadeNumero = converterMoedaParaNumero(dados.mensalidade);
    let adesaoNumero = processarAdesaoParaSalvar(dados.adesao);

    // Garantir situação válida
    const situacaoValida = (dados.situacao && dados.situacao.trim() !== '') ? dados.situacao : 'Novo registro';

    // 🔥🔥🔥 CORREÇÃO DEFINITIVA: MANTER DATA ATIVAÇÃO ORIGINAL OU USAR NOVA COM +1
    let dataAtivacaoParaSalvar = dataAtivacaoOriginal;
    
    if (dados.ativacao && dados.ativacao.trim() !== '') {
      try {
        const dataUsuario = new Date(dados.ativacao);
        // ✅✅✅ CORREÇÃO: ADICIONAR +1 DIA PARA COMPENSAR FUSO HORÁRIO
        dataUsuario.setDate(dataUsuario.getDate() + 1);
        dataAtivacaoParaSalvar = Utilities.formatDate(dataUsuario, CONFIG.TIMEZONE, "dd/MM/yyyy");
        console.log("📅 NOVA data ativação (COM +1 DIA):", dataAtivacaoParaSalvar);
      } catch (e) {
        console.error("❌ Erro ao processar data:", e);
        dataAtivacaoParaSalvar = dataAtivacaoOriginal;
      }
    } else {
      console.log("📅 Mantendo data ativação original:", dataAtivacaoOriginal);
      if (dataAtivacaoOriginal instanceof Date) {
        dataAtivacaoParaSalvar = Utilities.formatDate(dataAtivacaoOriginal, CONFIG.TIMEZONE, "dd/MM/yyyy");
      }
    }

    // 🔥🔥🔥 CORREÇÃO CRÍTICA: ATUALIZAR AMBAS AS COLUNAS E e F
    const dataAtual = new Date();
    const dataHoraAtual = formatarDataBrasil(dataAtual);
    
    console.log("🕐 Data/hora atual para Último evento:", dataHoraAtual);
    console.log("📝 Evento digitado pelo usuário:", dados.evento);

    // Array com 17 colunas na ORDEM CORRETA
    const novosDados = [
      normalizarTexto(dados.razao_social) || '',
      normalizarTexto(dados.nome_fantasia) || '',
      dados.cnpj ? dados.cnpj.toString() : '',
      normalizarTexto(fornecedorParaAtualizar),
      dataHoraAtual,                                        
      normalizarTexto(dados.evento) || '',                  
      normalizarTexto(dados.observacoes) || '',
      normalizarTexto(dados.contrato_enviado) || '',
      normalizarTexto(dados.contrato_assinado) || '',
      dataAtivacaoParaSalvar, // ✅ DATA COM +1 DIA
      dados.link || '',
      mensalidadeNumero,                                    
      converterMoedaParaNumero(dados.mensalidade_sim) || 0, 
      tarifaParaAtualizar || '',                            
      percentualParaAtualizar,                              
      adesaoNumero,                                         
      normalizarTexto(situacaoValida)                       
    ];

    console.log("📝 Atualizando linha:", linhaAtualizar);
    console.log("🎯 COLUNA E (Último evento):", novosDados[4]);
    console.log("🎯 COLUNA F (Evento):", novosDados[5]);
    console.log("🎯 COLUNA J (Ativação - COM +1 DIA):", novosDados[9]);
    
    // Salvar os dados
    aba.getRange(linhaAtualizar, 1, 1, novosDados.length).setValues([novosDados]);
    
    // Aplicar formatação
    aba.getRange(linhaAtualizar, 12).setNumberFormat('"R$"#,##0.00'); // L - Mensalidade
    aba.getRange(linhaAtualizar, 13).setNumberFormat('"R$"#,##0.00'); // M - Mensalidade SIM
    aba.getRange(linhaAtualizar, 15).setNumberFormat('0.00%');        // O - % Tarifa
    aba.getRange(linhaAtualizar, 16).setNumberFormat('"R$"#,##0.00'); // P - Adesão
    aba.getRange(linhaAtualizar, 14).setNumberFormat('@');            // N - Tarifa (texto)
    aba.getRange(linhaAtualizar, 10).setNumberFormat('dd/MM/yyyy');   // J - Ativação

    SpreadsheetApp.flush();

    console.log("✅ Atualização concluída - Data de ativação salva COM +1 DIA");

    return { 
      success: true, 
      message: `✅ "${dados.razao_social}" atualizado com sucesso no ${waitlabel}!` 
    };

  } catch (error) {
    console.error("❌ Erro em atualizarCadastroComWaitlabel:", error);
    return { success: false, message: "Erro ao atualizar: " + error.message };
  }
}

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

      // 🔥🔥🔥 DEBUG DAS COLUNAS E e F
      if (i < 3) { // Debug apenas dos primeiros 3 registros
        console.log(`🔍 DEBUG Registro ${i + 2}:`);
        console.log(`   Coluna E [4] - Último evento:`, linha[4], "Tipo:", typeof linha[4]);
        console.log(`   Coluna F [5] - Evento:`, linha[5], "Tipo:", typeof linha[5]);
      }
      
      // 🔥🔥🔥 CORREÇÃO: Último evento deve ser da COLUNA E (índice 4) - DATA
      let ultimoEventoFormatado = '';
      if (linha[4] && linha[4] instanceof Date) { // ✅ COLUNA E - DATA
        ultimoEventoFormatado = formatarDataBrasil(linha[4]);
      } else if (linha[4]) {
        ultimoEventoFormatado = linha[4].toString();
      }
      
      // 🔥🔥🔥 CORREÇÃO: Evento deve ser da COLUNA F (índice 5) - TEXTO
      let evento = linha[5]?.toString().trim() || ''; // ✅ COLUNA F - TEXTO
      
      let ativacaoFormatada = '';
      if (linha[9] && linha[9] instanceof Date) { // ✅ Ativação
        ativacaoFormatada = Utilities.formatDate(linha[9], CONFIG.TIMEZONE, "dd/MM/yyyy");
      } else if (linha[9]) {
        ativacaoFormatada = linha[9].toString();
      }
      
      // 🔥 CORREÇÃO: ESTRUTURA COM 17 COLUNAS - CORRIGIDO
      const cadastro = {
        id: i + 2,
        razao_social: linha[0]?.toString().trim() || '',     // A - Razão Social (0)
        nome_fantasia: linha[1]?.toString().trim() || '',    // B - Nome Fantasia (1)
        cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''), // C - CNPJ (2)
        fornecedor: linha[3]?.toString().trim() || '',       // D - Fornecedor (3)
        ultimo_evento: ultimoEventoFormatado,                // ✅ E - DATA
        evento: evento,                                      // ✅ F - TEXTO
        observacoes: linha[6]?.toString().trim() || '',      // G - Observação (6)
        contrato_enviado: linha[7]?.toString().trim() || '', // H - Contrato Enviado (7)
        contrato_assinado: linha[8]?.toString().trim() || '', // I - Contrato Assinado (8)
        ativacao: ativacaoFormatada,                         // J - Ativação (9)
        link: linha[10]?.toString().trim() || '',            // K - Link (10)
        mensalidade: parseFloat(linha[11]) || 0,             // L - Mensalidade (11)
        mensalidade_sim: parseFloat(linha[12]) || 0,         // M - Mensalidade SIM (12)
        tarifa: linha[13]?.toString().trim() || '',          // N - Tarifa (13)
        percentual_tarifa: linha[14]?.toString().trim() || '', // O - % Tarifa (14)
        adesao: processarAdesao(linha[15]),                  // P - Adesão (15)
        situacao: (linha[16]?.toString().trim() || 'Novo registro'), // Q - Situação (16)
        waitlabel: waitlabel
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
    console.log("🔍 BUSCAR TODOS CADASTROS POR CNPJ COM WAITLABEL - INICIANDO");
    console.log("📋 CNPJ:", cnpj, "Waitlabel:", waitlabel);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabel);
    
    if (!aba) {
      console.log("❌ Waitlabel não encontrado:", waitlabel);
      return [];
    }
    
    const ultimaLinha = aba.getLastRow();
    console.log("📊 Última linha:", ultimaLinha);
    
    if (ultimaLinha < 2) {
      console.log("ℹ️ Nenhum dato além do cabeçalho");
      return [];
    }
    
    // Buscar dados na ORDEM CORRETA (17 colunas)
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    
    console.log("🔎 Procurando CNPJ limpo:", cnpjBuscado);
    console.log("📈 Total de registros para filtrar:", dados.length);
    
    const cadastrosEncontrados = [];
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      
      // Pular linhas vazias
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      const cnpjCadastro = linha[2]?.toString().replace(/\D/g, '') || '';
      
      if (cnpjCadastro === cnpjBuscado) {
        console.log("✅ Cadastro encontrado na linha:", i + 2);
        
        // 🔥🔥🔥 CORREÇÃO: Formatar último evento - COLUNA E (índice 4) - DATA
        let ultimoEventoFormatado = '';
        if (linha[4] && linha[4] instanceof Date) { // ✅ COLUNA E - DATA
          ultimoEventoFormatado = formatarDataBrasil(linha[4]);
        } else if (linha[4]) {
          ultimoEventoFormatado = linha[4].toString();
        }
        
        let ativacaoFormatada = '';
        if (linha[9] && linha[9] instanceof Date) { // ✅ COLUNA J - Ativação
          ativacaoFormatada = Utilities.formatDate(linha[9], CONFIG.TIMEZONE, "dd/MM/yyyy");
        } else if (linha[9]) {
          ativacaoFormatada = linha[9].toString();
        }
        
        // 🔥 CORREÇÃO: ESTRUTURA COM 17 COLUNAS
        const cadastro = {
          id: i + 2,
          razao_social: linha[0]?.toString().trim() || '',     // A - Razão Social (0)
          nome_fantasia: linha[1]?.toString().trim() || '',    // B - Nome Fantasia (1)
          cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''), // C - CNPJ (2)
          fornecedor: linha[3]?.toString().trim() || '',       // D - Fornecedor (3)
          ultimo_evento: ultimoEventoFormatado,                // ✅ E - DATA
          evento: linha[5]?.toString().trim() || '',           // ✅ F - TEXTO
          observacoes: linha[6]?.toString().trim() || '',      // G - Observação (6)
          contrato_enviado: linha[7]?.toString().trim() || '', // H - Contrato Enviado (7)
          contrato_assinado: linha[8]?.toString().trim() || '', // I - Contrato Assinado (8)
          ativacao: ativacaoFormatada,                         // J - Ativação (9)
          link: linha[10]?.toString().trim() || '',            // K - Link (10)
          mensalidade: parseFloat(linha[11]) || 0,             // L - Mensalidade (11)
          mensalidade_sim: parseFloat(linha[12]) || 0,         // M - Mensalidade SIM (12)
          tarifa: linha[13]?.toString().trim() || '',          // N - Tarifa (13)
          percentual_tarifa: linha[14]?.toString().trim() || '', // O - % Tarifa (14)
          adesao: processarAdesao(linha[15]),                  // P - Adesão (15)
          situacao: (linha[16]?.toString().trim() || 'Novo registro'), // Q - Situação (16)
          waitlabel: waitlabel
        };
        
        cadastrosEncontrados.push(cadastro);
      }
    }
    
    console.log(`✅ Encontrados ${cadastrosEncontrados.length} cadastro(s) para o CNPJ ${cnpj}`);
    return cadastrosEncontrados;
    
  } catch (error) {
    console.error("❌ Erro em buscarTodosCadastrosPorCNPJComWaitlabel:", error);
    return [];
  }
}

function processarLinhaParaRetorno(linha, id) {
  console.log("=== 🔍 DEBUG processarLinhaParaRetorno - INÍCIO ===");
  
  // 🔥🔥🔥 CORREÇÃO DEFINITIVA: COLUNAS E e F CORRETAS
  // Coluna E (índice 4) = DATA DO ÚLTIMO EVENTO (14/11/2025 15:46:23)
  // Coluna F (índice 5) = TEXTO DO EVENTO ("NOVO CADASTRO")
  
  let ultimoEventoFormatado = '';
  if (linha[4] && linha[4] instanceof Date) { // ✅ COLUNA E - DATA
    ultimoEventoFormatado = formatarDataBrasil(linha[4]);
  } else if (linha[4]) {
    ultimoEventoFormatado = linha[4].toString(); // JÁ ESTÁ NO FORMATO CERTO
  }
  
  let evento = linha[5]?.toString().trim() || ''; // ✅ COLUNA F - TEXTO
  
  console.log("🎯🎯🎯 DEBUG CRÍTICO DAS COLUNAS E e F:");
  console.log("Coluna E [4] - Último evento BRUTO:", linha[4], "Tipo:", typeof linha[4]);
  console.log("Coluna F [5] - Evento BRUTO:", linha[5], "Tipo:", typeof linha[5]);
  console.log("Último evento formatado:", ultimoEventoFormatado);
  console.log("Evento texto:", evento);
  
  // Formatar data ativação
  let ativacaoFormatada = '';
  if (linha[9] && linha[9] instanceof Date) { // ✅ COLUNA J - Ativação
    ativacaoFormatada = Utilities.formatDate(linha[9], CONFIG.TIMEZONE, "yyyy-MM-dd");
  } else if (linha[9]) {
    if (linha[9].includes('/')) {
      const partes = linha[9].split('/');
      ativacaoFormatada = `${partes[2]}-${partes[1]}-${partes[0]}`;
    } else {
      ativacaoFormatada = linha[9].toString();
    }
  }

  // 🔥🔥🔥 CORREÇÃO: Referências corretas das colunas financeiras
  console.log("🔍 Dados brutos das colunas financeiras:");
  console.log("Coluna 13 (N - Tarifa):", linha[13], "Tipo:", typeof linha[13]);
  console.log("Coluna 14 (O - % Tarifa):", linha[14], "Tipo:", typeof linha[14]);
  console.log("Coluna 15 (P - Adesão):", linha[15], "Tipo:", typeof linha[15]);

  let tarifa = linha[13]?.toString().trim() || ''; // Coluna N - Tarifa
  
  let percentualTarifa = '0%';
  if (linha[14] !== null && linha[14] !== undefined && linha[14] !== '') { // Coluna O - % Tarifa
    const valor = parseFloat(linha[14]);
    if (!isNaN(valor)) {
      percentualTarifa = (valor * 100).toFixed(2) + '%';
    } else {
      percentualTarifa = linha[14]?.toString().trim() || '0%';
    }
  }

  let adesaoProcessada = processarAdesao(linha[15]); // Coluna P - Adesão
  
  console.log("💰 Valores processados:");
  console.log("   Tarifa:", tarifa);
  console.log("   % Tarifa:", percentualTarifa);
  console.log("   Adesão:", adesaoProcessada);
  
  // Estrutura de fornecedor para formulário
  const fornecedorParaFormulario = {
    nome: linha[3]?.toString().trim() || '',
    tarifa: tarifa,
    percentual_tarifa: percentualTarifa
  };
  
  // 🔥🔥🔥 CORREÇÃO: ESTRUTURA COM REFERÊNCIAS CORRETAS
  const resultado = {
    encontrado: true,
    id: id,
    razao_social: linha[0]?.toString().trim() || '',     // A - Razão Social
    nome_fantasia: linha[1]?.toString().trim() || '',    // B - Nome Fantasia
    cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''), // C - CNPJ
    fornecedor: linha[3]?.toString().trim() || '',       // D - Fornecedor
    fornecedores: [fornecedorParaFormulario],
    ultimo_evento: ultimoEventoFormatado,                // ✅ E - DATA (14/11/2025 15:46:23)
    evento: evento,                                      // ✅ F - TEXTO ("NOVO CADASTRO")
    observacoes: linha[6]?.toString().trim() || '',      // G - Observação
    contrato_enviado: linha[7]?.toString().trim() || '', // H - Contrato Enviado
    contrato_assinado: linha[8]?.toString().trim() || '', // I - Contrato Assinado
    ativacao: ativacaoFormatada,                         // J - Ativação
    link: linha[10]?.toString().trim() || '',            // K - Link
    mensalidade: parseFloat(linha[11]) || 0,             // L - Mensalidade
    mensalidade_sim: parseFloat(linha[12]) || 0,         // M - Mensalidade SIM
    tarifa: tarifa,                                      // N - Tarifa
    percentual_tarifa: percentualTarifa,                 // O - % Tarifa
    adesao: adesaoProcessada,                            // P - Adesão
    situacao: (linha[16]?.toString().trim() || 'Novo registro') // Q - Situação
  };

  console.log("=== ✅ DEBUG processarLinhaParaRetorno - FIM ===");
  console.log("🎯 RESULTADO FINAL:");
  console.log("   Último evento (DATA):", resultado.ultimo_evento);
  console.log("   Evento (TEXTO):", resultado.evento);
  return resultado;
}

function debugOrdemColunasReal() {
  try {
    const waitlabelAtual = getWaitlabelAtual();
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabelAtual);
    
    if (!aba) {
      console.log("❌ Aba não encontrada:", waitlabelAtual);
      return { error: "Aba não encontrada" };
    }
    
    const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    const primeiraLinha = aba.getRange(2, 1, 1, aba.getLastColumn()).getValues()[0];
    
    console.log("=== 🔍 ORDEM REAL DAS COLUNAS ===");
    cabecalhos.forEach((cabecalho, index) => {
      const letraColuna = String.fromCharCode(65 + index);
      console.log(`Coluna ${letraColuna} [${index}]: "${cabecalho}" = ${primeiraLinha[index]}`);
    });
    
    // Foco especial nas colunas E e F
    console.log("=== 🎯 FOCO COLUNAS E e F ===");
    console.log("Coluna E [4]:", cabecalhos[4], "=", primeiraLinha[4]);
    console.log("Coluna F [5]:", cabecalhos[5], "=", primeiraLinha[5]);
    
    return {
      cabecalhos: cabecalhos,
      dados: primeiraLinha
    };
    
  } catch (error) {
    console.error("❌ Erro:", error);
    return { error: error.message };
  }
}

function debugOrdemColunasSimFacilita() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName('Sim_Facilita'); // 🔥 Mudei para Sim_Facilita
    
    if (!aba) {
      console.log("❌ Aba Sim_Facilita não encontrada!");
      return { error: "Aba Sim_Facilita não encontrada" };
    }
    
    const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    const primeiraLinha = aba.getRange(2, 1, 1, aba.getLastColumn()).getValues()[0];
    
    console.log("=== 🔍 ORDEM REAL DAS COLUNAS - SIM_FACILITA ===");
    cabecalhos.forEach((cabecalho, index) => {
      console.log(`Coluna ${index}: "${cabecalho}" = ${primeiraLinha[index]}`);
    });
    
    return {
      cabecalhos: cabecalhos,
      dados: primeiraLinha
    };
    
  } catch (error) {
    console.error("❌ Erro:", error);
    return { error: error.message };
  }
}

function buscarCadastroPorIDComWaitlabel(id, waitlabel) {
  try {
    console.log("🔍🔍🔍 DEBUG COMPLETO - Buscando cadastro por ID:", id, "no waitlabel:", waitlabel);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabel);
    if (!aba) return { encontrado: false, mensagem: "Waitlabel não encontrado" };
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < id) return { encontrado: false, mensagem: "Registro não encontrado" };
    
    const linha = aba.getRange(id, 1, 1, 17).getValues()[0];
    
    if (!linha[0] || linha[0].toString().trim() === '') {
      return { encontrado: false, mensagem: "Registro vazio ou não encontrado" };
    }

    // 🔥🔥🔥 DEBUG SUPER DETALHADO - VERIFICAR O QUE ESTÁ SENDO PROCESSADO
    console.log("=== 🎯 DEBUG DAS COLUNAS NA FONTE ===");
    console.log("📊 Linha completa:", linha);
    
    // 🔥🔥🔥 DEBUG CRÍTICO - COLUNAS E e F
    console.log("🎯🎯🎯 DEBUG CRÍTICO - COLUNAS E e F:");
    console.log("🔍 Coluna E [4] - Último evento BRUTO:", linha[4], "Tipo:", typeof linha[4]);
    console.log("🔍 Coluna F [5] - Evento BRUTO:", linha[5], "Tipo:", typeof linha[5]);
    console.log("🔍 Coluna E como string:", linha[4]?.toString());
    console.log("🔍 Coluna F como string:", linha[5]?.toString());
    
    // Debug das colunas financeiras
    console.log("💰 COLUNAS FINANCEIRAS BRUTAS:");
    console.log("🔍 Coluna 13 (N - Tarifa) BRUTO:", linha[13], "Tipo:", typeof linha[13]);
    console.log("🔍 Coluna 14 (O - % Tarifa) BRUTO:", linha[14], "Tipo:", typeof linha[14]);
    console.log("🔍 Coluna 15 (P - Adesão) BRUTO:", linha[15], "Tipo:", typeof linha[15]);
    
    console.log("🔍 Coluna 8 (Contrato Enviado) BRUTO:", linha[8], "Tipo:", typeof linha[8]);
    console.log("🔍 Coluna 9 (Contrato Assinado) BRUTO:", linha[9], "Tipo:", typeof linha[9]);
    console.log("🔍 Coluna 10 (Ativação) BRUTO:", linha[10], "Tipo:", typeof linha[10]);
    
    // 🔥🔥🔥 TESTE DIRETO - PROCESSAR NA MÃO
    const contratoEnviadoTeste = linha[8]?.toString().trim() || '';
    const contratoAssinadoTeste = linha[9]?.toString().trim() || '';
    console.log("🧪 TESTE DIRETO - Contrato Enviado:", contratoEnviadoTeste);
    console.log("🧪 TESTE DIRETO - Contrato Assinado:", contratoAssinadoTeste);
    
    const resultado = processarLinhaParaRetorno(linha, id);
    resultado.waitlabel = waitlabel;
    
    console.log("=== ✅ RESULTADO FINAL DA FUNÇÃO processarLinhaParaRetorno ===");
    console.log("🎯 DADOS TEMPORAIS:");
    console.log("   Último evento:", resultado.ultimo_evento);
    console.log("   Evento:", resultado.evento);
    console.log("   Ativação:", resultado.ativacao);
    console.log("Contrato Enviado no resultado:", resultado.contrato_enviado);
    console.log("Contrato Assinado no resultado:", resultado.contrato_assinado);
    console.log("💰 DADOS FINANCEIROS NO RESULTADO:");
    console.log("   Tarifa:", resultado.tarifa);
    console.log("   % Tarifa:", resultado.percentual_tarifa);
    console.log("   Adesão:", resultado.adesao);
    console.log("   Mensalidade:", resultado.mensalidade);
    console.log("   Mensalidade SIM:", resultado.mensalidade_sim);
    console.log("   Ativação:", resultado.ativacao);
    
    return resultado;
    
  } catch (error) {
    console.error("❌ Erro em buscarCadastroPorIDComWaitlabel:", error);
    return { encontrado: false, mensagem: "Erro: " + error.message };
  }
}

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
        
        // 🔥🔥🔥 CORREÇÃO 1: Aplicar campos selecionados
        camposSelecionados.forEach(campo => {
          const indiceColuna = obterIndiceColuna(campo);
          if (indiceColuna !== -1) {
            const novoValor = obterValorParaCampo(campo, dadosParaAplicar, linha);
            novosDados[indiceColuna] = novoValor;
            console.log(`   ✅ Campo "${campo}" [coluna ${indiceColuna + 1}]: "${novoValor}"`);
          }
        });

        // 🔥🔥🔥 CORREÇÃO CRÍTICA: SALVAR CORRETAMENTE NAS COLUNAS E e F
        if (camposSelecionados.includes('evento')) {
          // COLUNA E = DATA atual (Último evento)
          novosDados[4] = formatarDataBrasil(new Date());
          // COLUNA F = TEXTO do evento (já foi salvo acima pelo forEach)
          console.log("🎯 COLUNA E (Data - Último evento):", novosDados[4]);
          console.log("🎯 COLUNA F (Evento - texto):", novosDados[5]);
        } else {
          // Se não está aplicando evento, atualizar apenas a data do último evento
          novosDados[4] = formatarDataBrasil(new Date());
        }
        
        try {
          aba.getRange(linhaNumero, 1, 1, novosDados.length).setValues([novosDados]);
          aplicarFormatacao(aba, linhaNumero, camposSelecionados);
          
          registrosAtualizados++;
          resultados.push(`✅ Linha ${linhaNumero} - ${linha[3]}`);
          
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

// 🔥🔥🔥 CORREÇÃO 2: Função auxiliar atualizada
function obterIndiceColuna(campo) {
  const mapeamentoCampos = {
    'razao_social': 0,      // A
    'nome_fantasia': 1,     // B  
    'cnpj': 2,              // C
    'fornecedores': 3,      // D - Fornecedor
    'evento': 5,            // F - Evento (TEXTO) - CORRETO!
    'observacoes': 6,       // G - Observação
    'contrato_enviado': 7,  // H - Contrato Enviado
    'contrato_assinado': 8, // I - Contrato Assinado
    'ativacao': 9,          // J - Ativação
    'link': 10,             // K - Link
    'mensalidade': 11,      // L - Mensalidade
    'mensalidade_sim': 12,  // M - Mensalidade SIM
    'adesao': 15,           // P - Adesão
    'situacao': 16          // Q - Situação
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
    case 'evento':
      return normalizarTexto(dadosParaAplicar.evento) || ''; // ✅ COLUNA F - EVENTO TEXTO
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
          // ✅✅✅ CORREÇÃO: ADICIONAR +1 DIA PARA COMPENSAR FUSO HORÁRIO
          dataUsuario.setDate(dataUsuario.getDate() + 1);
          return Utilities.formatDate(dataUsuario, CONFIG.TIMEZONE, "dd/MM/yyyy");
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
    case 'mensalidade_sim':
      return converterMoedaParaNumero(dadosParaAplicar.mensalidade_sim) || 0;
    case 'adesao':
      return processarAdesaoParaSalvar(dadosParaAplicar.adesao);
    case 'situacao':
      let situacao = normalizarTexto(dadosParaAplicar.situacao) || 'NOVO REGISTRO';
      if (situacao === 'NOVO REGISTRO') situacao = 'Novo Registro';
      return situacao;
    case 'fornecedores':
      return linhaAtual[3];
    default:
      return linhaAtual[obterIndiceColuna(campo)];
  }
}

function aplicarFormatacao(aba, linhaNumero, camposSelecionados) {
  try {
    aba.getRange(linhaNumero, 12).setNumberFormat('"R$"#,##0.00'); // Mensalidade (L) - índice 11
    aba.getRange(linhaNumero, 13).setNumberFormat('"R$"#,##0.00'); // Mensalidade SIM (M) - índice 12
    aba.getRange(linhaNumero, 16).setNumberFormat('"R$"#,##0.00'); // Adesão (P) - índice 15
    aba.getRange(linhaNumero, 15).setNumberFormat('0.00%');        // % Tarifa (O) - índice 14
    aba.getRange(linhaNumero, 10).setNumberFormat('dd/MM/yyyy');   // Data Ativação (J) - índice 9
    
    if (camposSelecionados.includes('mensalidade')) {
      aba.getRange(linhaNumero, 13).setNumberFormat('"R$"#,##0.00');
    }
    
    if (camposSelecionados.includes('mensalidade_sim')) { // 🔥 NOVO
      aba.getRange(linhaNumero, 14).setNumberFormat('"R$"#,##0.00');
    }
    
    if (camposSelecionados.includes('adesao')) {
      aba.getRange(linhaNumero, 17).setNumberFormat('"R$"#,##0.00'); // ATUALIZADO
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
        'Razão Social', 'Nome Fantasia', 'CNPJ', 'Fornecedor', 
        'Ultimo evento', 'Evento', 'Observação', 'Contrato Enviado', 'Contrato Assinado',
        'Ativação', 'Link', 'Mensalidade', 'Mensalidade SIM', 'Tarifa', '% Tarifa', 'Adesão', 'Situação'
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
      const dataUltimoEvento = formatarDataBrasil(dataAtual);

      // ✅ CORREÇÃO: Usar data informada pelo usuário SEM adicionar dias
      let dataAtivacaoParaSalvar = '';
      if (dados.ativacao && dados.ativacao.trim() !== '') {
        try {
          const dataUsuario = new Date(dados.ativacao);
          dataAtivacaoParaSalvar = Utilities.formatDate(dataUsuario, CONFIG.TIMEZONE, "dd/MM/yyyy");
          console.log("📅 Data ativação informada pelo usuário (CORRIGIDA):", dataAtivacaoParaSalvar);
        } catch (e) {
          console.error("❌ Erro ao processar data do usuário:", e);
          dataAtivacaoParaSalvar = '';
        }
      } else {
        console.log("📅 Nenhuma data de ativação informada - campo ficará vazio");
      }

      console.log(`📅 Datas geradas para fornecedor ${i + 1}:`);
      console.log(`   Data Ativação: ${dataAtivacaoParaSalvar}`);
      console.log(`   Data Último Evento: ${dataUltimoEvento}`);

      // Array com 17 colunas na ORDEM CORRETA
      const linhaDados = [
        normalizarTexto(dados.razao_social) || '',           // A (0)
        normalizarTexto(dados.nome_fantasia) || '',          // B (1)
        dados.cnpj ? dados.cnpj.toString() : '',             // C (2)
        normalizarTexto(nomeFornecedor),                     // D (3) - Fornecedor
        dataUltimoEvento,                                    // E (4) - Último evento
        normalizarTexto(dados.evento) || '',                 // F (5) - Evento
        normalizarTexto(dados.observacoes) || '',            // G (6) - Observação
        normalizarTexto(dados.contrato_enviado) || '',       // H (7) - Contrato Enviado
        normalizarTexto(dados.contrato_assinado) || '',      // I (8) - Contrato Assinado
        dataAtivacaoParaSalvar,                              // J (9) - Ativação
        dados.link || '',                                    // K (10) - Link
        mensalidadeNumero,                                   // L (11) - Mensalidade
        converterMoedaParaNumero(dados.mensalidade_sim) || 0,// M (12) - Mensalidade SIM
        tarifaFornecedor || '',                              // N (13) - Tarifa
        percentualTarifaFornecedor,                          // O (14) - % Tarifa
        adesaoNumero,                                        // P (15) - Adesão
        normalizarTexto(situacaoParaSalvar)                  // Q (16) - Situação
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
    const dataAtivacaoOriginal = dadosAtuais[9]; // Coluna J - Ativação
    
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

    // 🔥🔥🔥 CORREÇÃO 2: MANTER A DATA DE ATIVAÇÃO ORIGINAL OU USAR NOVA SEM +1
    let dataAtivacaoParaSalvar = dataAtivacaoOriginal;
    
    if (dados.ativacao && dados.ativacao.trim() !== '') {
      try {
        const dataUsuario = new Date(dados.ativacao);
        // ✅✅✅ CORREÇÃO: REMOVIDO O +1 - USAR DATA EXATA DO USUÁRIO
        dataAtivacaoParaSalvar = Utilities.formatDate(dataUsuario, CONFIG.TIMEZONE, "dd/MM/yyyy");
        console.log("📅 NOVA data ativação (EXATA DO USUÁRIO):", dataAtivacaoParaSalvar);
      } catch (e) {
        console.error("❌ Erro ao processar data:", e);
        dataAtivacaoParaSalvar = dataAtivacaoOriginal;
      }
    } else {
      console.log("📅 Mantendo data ativação original:", dataAtivacaoOriginal);
      if (dataAtivacaoOriginal instanceof Date) {
        dataAtivacaoParaSalvar = Utilities.formatDate(dataAtivacaoOriginal, CONFIG.TIMEZONE, "dd/MM/yyyy");
      }
    }

    console.log("📅 Data ativação que será salva (EXATA):", dataAtivacaoParaSalvar);

    // Array com 17 colunas na ORDEM CORRETA
    const novosDados = [
      normalizarTexto(dados.razao_social) || '',           // A (0)
      normalizarTexto(dados.nome_fantasia) || '',          // B (1)
      dados.cnpj ? dados.cnpj.toString() : '',             // C (2)
      normalizarTexto(fornecedorParaAtualizar),            // D (3)
      Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm:ss"), // E (4)
      normalizarTexto(dados.evento) || '',                 // F (5)
      normalizarTexto(dados.observacoes) || '',            // G (6)
      normalizarTexto(dados.contrato_enviado) || '',       // H (7)
      normalizarTexto(dados.contrato_assinado) || '',      // I (8)
      dataAtivacaoParaSalvar,                              // J (9) - ✅ DATA EXATA
      dados.link || '',                                    // K (10)
      mensalidadeNumero,                                   // L (11)
      converterMoedaParaNumero(dados.mensalidade_sim) || 0,// M (12)
      tarifaParaAtualizar || '',                           // N (13)
      percentualParaAtualizar,                             // O (14)
      adesaoNumero,                                        // P (15)
      normalizarTexto(situacaoValida)                      // Q (16)
    ];

    console.log("📝 Atualizando linha:", linhaAtualizar);
    console.log("📊 Novos dados:", novosDados);
    console.log("🎯 Data ativação salva (EXATA):", novosDados[9]);
    
    aba.getRange(linhaAtualizar, 1, 1, novosDados.length).setValues([novosDados]);
    
    // 🔥🔥🔥 CORREÇÃO: ADICIONAR FORMATAÇÃO DA TARIFA
    aba.getRange(linhaAtualizar, 12).setNumberFormat('"R$"#,##0.00'); // Mensalidade (L)
    aba.getRange(linhaAtualizar, 13).setNumberFormat('"R$"#,##0.00'); // Mensalidade SIM (M)
    aba.getRange(linhaAtualizar, 16).setNumberFormat('"R$"#,##0.00'); // Adesão (P)
    aba.getRange(linhaAtualizar, 15).setNumberFormat('0.00%');        // % Tarifa (O)
    aba.getRange(linhaAtualizar, 14).setNumberFormat('@');            // Tarifa como texto (N)
    aba.getRange(linhaAtualizar, 10).setNumberFormat('dd/MM/yyyy');   // Data Ativação (J)

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
        ultimoEventoFormatado = Utilities.formatDate(linha[5], CONFIG.TIMEZONE, "dd/MM/yyyy");
      } else if (linha[5]) {
        ultimoEventoFormatado = linha[5].toString();
      }
      
      let ativacaoFormatada = '';
      if (linha[10] && linha[10] instanceof Date) { // ✅ Ativação
        ativacaoFormatada = Utilities.formatDate(linha[10], CONFIG.TIMEZONE, "dd/MM/yyyy");
      } else if (linha[10]) {
        ativacaoFormatada = linha[10].toString();
      }
      
      // 🔥 CORREÇÃO: ESTRUTURA COM 17 COLUNAS
      const cadastro = {
        id: i + 2,
        razao_social: linha[0]?.toString().trim() || '',     // A - Razão Social (0)
        nome_fantasia: linha[1]?.toString().trim() || '',    // B - Nome Fantasia (1)
        cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''), // C - CNPJ (2)
        fornecedor: linha[3]?.toString().trim() || '',       // D - Fornecedor (3)
        ultimo_evento: ultimoEventoFormatado,                // E - Último evento (4)
        evento: linha[5]?.toString().trim() || '',           // F - Evento (5)
        observacoes: linha[6]?.toString().trim() || '',      // G - Observação (6)
        contrato_enviado: linha[7]?.toString().trim() || '', // H - Contrato Enviado (7)
        contrato_assinado: linha[8]?.toString().trim() || '', // I - Contrato Assinado (8)
        ativacao: ativacaoFormatada,                         // J - Ativação (9)
        link: linha[10]?.toString().trim() || '',            // K - Link (10)
        mensalidade: parseFloat(linha[11]) || 0,             // L - Mensalidade (11)
        mensalidade_sim: parseFloat(linha[12]) || 0,         // 🔥 M - Mensalidade SIM (12) - VOCÊ ESQUECEU ESTA!
        tarifa: linha[13]?.toString().trim() || '',          // N - Tarifa (13)
        percentual_tarifa: linha[14]?.toString().trim() || '', // O - % Tarifa (14)
        adesao: processarAdesao(linha[15]),                  // P - Adesão (15)
        situacao: (linha[16]?.toString().trim() || 'Novo registro') // Q - Situação (16)
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
          ultimoEventoFormatado = Utilities.formatDate(linha[5], CONFIG.TIMEZONE, "dd/MM/yyyy");
        } else if (linha[5]) {
          ultimoEventoFormatado = linha[5].toString();
        }
        
        // 🔥 CORREÇÃO: Data ativação para formato do input date
        let ativacaoFormatada = '';
        if (linha[10] && linha[10] instanceof Date) { // ✅ Ativação
          ativacaoFormatada = Utilities.formatDate(linha[10], CONFIG.TIMEZONE, "yyyy-MM-dd"); // 🔥 FORMATO PARA INPUT DATE
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
        // ✅ CORREÇÃO: MANTER O VALOR EXATO SEM ARREDONDAMENTO
        let percentualTarifa = '0%';
        if (linha[15] !== null && linha[15] !== undefined && linha[15] !== '') {
          const valor = parseFloat(linha[15]);
          if (!isNaN(valor)) {
            // 🔥 CORREÇÃO: Usar toFixed(2) para manter casas decimais
            percentualTarifa = (valor * 100).toFixed(2) + '%'; // 0.035 → 3.50%
          } else {
            // Se já está como string com %, manter como está
            percentualTarifa = linha[15]?.toString().trim() || '0%';
          }
        }
        
        console.log(`💰 Tarifa encontrada: "${tarifa}"`);
        console.log(`📊 % Tarifa encontrada: "${percentualTarifa}"`);
        console.log(`📅 Ativação encontrada: "${linha[10]}" → Formatada: "${ativacaoFormatada}"`);
        
        // 🔥🔥🔥 CORREÇÃO CRÍTICA: Estrutura de fornecedores para o formulário
        const fornecedorParaFormulario = {
          nome: linha[3]?.toString().trim() || '', // ✅ CORRIGIDO: índice 3 (Fornecedor)
          tarifa: tarifa,
          percentual_tarifa: percentualTarifa
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
          razao_social: linha[0]?.toString().trim() || '',     // A - Razão Social (0)
          nome_fantasia: linha[1]?.toString().trim() || '',    // B - Nome Fantasia (1)
          cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''), // C - CNPJ (2)
          fornecedor: linha[3]?.toString().trim() || '',       // D - Fornecedor (3)
          fornecedores: [fornecedorParaFormulario],            // 🔥 ESTRUTURA QUE O FORMULÁRIO ESPERA
          ultimo_evento: ultimoEventoFormatado,                // E - Último evento (4)
          evento: linha[5]?.toString().trim() || '',           // F - Evento (5)
          observacoes: linha[6]?.toString().trim() || '',      // G - Observação (6)
          contrato_enviado: linha[7]?.toString().trim() || '', // H - Contrato Enviado (7)
          contrato_assinado: linha[8]?.toString().trim() || '', // I - Contrato Assinado (8)
          ativacao: ativacaoFormatada,                         // J - Ativação (9)
          link: linha[10]?.toString().trim() || '',            // K - Link (10)
          mensalidade: parseFloat(linha[11]) || 0,             // L - Mensalidade (11)
          mensalidade_sim: parseFloat(linha[12]) || 0,         // M - Mensalidade SIM (12)
          tarifa: tarifa,                                      // N - Tarifa (13)
          percentual_tarifa: percentualTarifa,                 // O - % Tarifa (14)
          adesao: processarAdesao(linha[15]),                  // P - Adesão (15)
          situacao: (linha[16]?.toString().trim() || 'Novo registro') // Q - Situação (16)
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
      ultimoEventoFormatado = Utilities.formatDate(linha[5], CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm:ss");
    } else if (linha[5]) {
      ultimoEventoFormatado = linha[5].toString();
    }
    
    let ativacaoFormatada = '';
    if (linha[10] && linha[10] instanceof Date) { // ✅ CORRETO: linha[10] é Ativação
      ativacaoFormatada = Utilities.formatDate(linha[10], CONFIG.TIMEZONE, "yyyy-MM-dd"); // 🔥 FORMATO PARA INPUT DATE
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
    // ✅ CORREÇÃO: MANTER O VALOR EXATO SEM ARREDONDAMENTO
    let percentualTarifa = '0%';
    if (linha[15] !== null && linha[15] !== undefined && linha[15] !== '') {
      const valor = parseFloat(linha[15]);
      if (!isNaN(valor)) {
        // 🔥 CORREÇÃO: Usar toFixed(2) para manter casas decimais
        percentualTarifa = (valor * 100).toFixed(2) + '%'; // 0.035 → 3.50%
      } else {
        // Se já está como string com %, manter como está
        percentualTarifa = linha[15]?.toString().trim() || '0%';
      }
    }
  
    console.log(`💰 Tarifa encontrada: "${tarifa}"`);
    console.log(`📊 % Tarifa encontrada: "${percentualTarifa}"`);
    console.log(`📅 Ativação encontrada: "${linha[10]}" → Formatada: "${ativacaoFormatada}"`);
    
    // 🔥🔥🔥 CORREÇÃO CRÍTICA: Estrutura de fornecedores para o formulário
    const fornecedorParaFormulario = {
      nome: linha[3]?.toString().trim() || '', // ✅ índice 3 (Fornecedor)
      tarifa: tarifa,
      percentual_tarifa: percentualTarifa
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
      razao_social: linha[0]?.toString().trim() || '',     // A - Razão Social (0)
      nome_fantasia: linha[1]?.toString().trim() || '',    // B - Nome Fantasia (1)
      cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''), // C - CNPJ (2)
      fornecedor: linha[3]?.toString().trim() || '',       // D - Fornecedor (3)
      fornecedores: [fornecedorParaFormulario],
      ultimo_evento: ultimoEventoFormatado,
      evento: linha[5]?.toString().trim() || '',           // F - Evento (5)
      observacoes: linha[6]?.toString().trim() || '',      // G - Observação (6)
      contrato_enviado: linha[7]?.toString().trim() || '', // H - Contrato Enviado (7)
      contrato_assinado: linha[8]?.toString().trim() || '', // I - Contrato Assinado (8)
      ativacao: ativacaoFormatada,
      link: linha[10]?.toString().trim() || '',            // K - Link (10)
      mensalidade: parseFloat(linha[11]) || 0,             // L - Mensalidade (11)
      mensalidade_sim: parseFloat(linha[12]) || 0,         // M - Mensalidade SIM (12)
      tarifa: tarifa,                                      // N - Tarifa (13)
      percentual_tarifa: percentualTarifa,                 // O - % Tarifa (14)
      adesao: processarAdesao(linha[15]),                  // P - Adesão (15)
      situacao: (linha[16]?.toString().trim() || 'Novo registro') // Q - Situação (16)
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
          fornecedor: linha[3]?.toString().trim() || '',
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

// 🔥🔥🔥 FUNÇÃO TEMPORÁRIA PARA DEBUG DO TWO SISTERS
function debugTwoSisters() {
  try {
    console.log("=== 🎯 DEBUG ESPECÍFICO TWO SISTERS ===");
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName('Result');
    
    if (!aba) {
      console.log("❌ Aba Result não encontrada");
      return;
    }
    
    // Buscar especificamente a linha 2 (que é o TWO SISTERS)
    const linha = aba.getRange(2, 1, 1, 17).getValues()[0];
    
    console.log("📊 LINHA COMPLETA DO TWO SISTERS:");
    for (let i = 0; i < linha.length; i++) {
      const letraColuna = String.fromCharCode(65 + i);
      console.log(`Coluna ${letraColuna} [${i}]:`, linha[i], "Tipo:", typeof linha[i]);
    }
    
    console.log("=== 🔍 DETALHES CONTRATO ASSINADO ===");
    console.log("Coluna J [9] - Contrato Assinado:", linha[9]);
    console.log("Como string:", linha[9]?.toString());
    console.log("Trimmed:", linha[9]?.toString().trim());
    console.log("Uppercase:", linha[9]?.toString().trim().toUpperCase());
    console.log("É exatamente 'SIM':", linha[9]?.toString().trim().toUpperCase() === 'SIM');
    
    // Testar a função processarLinhaParaRetorno
    console.log("=== 🧪 TESTE processarLinhaParaRetorno ===");
    const resultado = processarLinhaParaRetorno(linha, 2);
    console.log("Contrato Assinado no resultado:", resultado.contrato_assinado);
    
    return {
      linhaCompleta: linha,
      contratoAssinadoBruto: linha[9],
      contratoAssinadoProcessado: resultado.contrato_assinado
    };
    
  } catch (error) {
    console.error("❌ Erro no debug:", error);
    return { error: error.message };
  }
}

function testarContratoAssinado() {
  try {
    console.log("=== 🧪 TESTE CONTRATO ASSINADO ===");
    
    const ss = SpreadsheetApp.openById("1V4iGN14UpIQcwf3qKU0_Wbiy2exdW2WUmrYTniy0upA");
    const aba = ss.getSheetByName('Result');
    
    if (!aba) {
      console.log("❌ Aba não encontrada");
      return;
    }
    
    // Buscar linha 2 (TWO SISTERS)
    const linha = aba.getRange(2, 1, 1, 17).getValues()[0];
    
    console.log("📊 LINHA COMPLETA:");
    for (let i = 0; i < linha.length; i++) {
      const letraColuna = String.fromCharCode(65 + i);
      console.log(`Coluna ${letraColuna} [${i}]:`, linha[i], "Tipo:", typeof linha[i]);
    }
    
    console.log("=== 🔍 DETALHES CONTRATOS ===");
    console.log("Coluna I [8] - Contrato Enviado:", linha[8]);
    console.log("Coluna J [9] - Contrato Assinado:", linha[9]);
    
    // Testar processamento
    const contratoEnviado = linha[8]?.toString().trim() || '';
    const contratoAssinado = linha[9]?.toString().trim() || '';
    
    console.log("✅ Contrato Enviado processado:", contratoEnviado);
    console.log("✅ Contrato Assinado processado:", contratoAssinado);
    
    return {
      contrato_enviado: contratoEnviado,
      contrato_assinado: contratoAssinado
    };
    
  } catch (error) {
    console.error("❌ Erro no teste:", error);
    return { error: error.message };
  }
}

function testarBuscaComWaitlabel() {
  try {
    console.log("=== 🧪 TESTE BUSCA COM WAITLABEL ===");
    
    // Testar a busca pelo ID 2 no waitlabel Result
    const resultado = buscarCadastroPorIDComWaitlabel(2, 'Result');
    
    console.log("=== 📋 RESULTADO FINAL ===");
    console.log("Encontrado:", resultado.encontrado);
    console.log("Contrato Enviado:", resultado.contrato_enviado);
    console.log("Contrato Assinado:", resultado.contrato_assinado);
    console.log("Tipo Contrato Enviado:", typeof resultado.contrato_enviado);
    console.log("Tipo Contrato Assinado:", typeof resultado.contrato_assinado);
    
    return resultado;
    
  } catch (error) {
    console.error("❌ Erro no teste:", error);
    return { error: error.message };
  }
}

function testarPercentualCorrigido() {
  const resultado = buscarCadastroPorIDComWaitlabel(988, 'Result');
  console.log("🎯 RESULTADO DO TESTE:");
  console.log("Percentual tarifa:", resultado.percentual_tarifa);
  console.log("Deve ser 3.50% (não 4%)");
  return resultado;
}

// 🔥🔥🔥 FUNÇÃO DE TESTE DO FUSO HORÁRIO
function testarFusoHorario() {
  console.log("=== 🧪 TESTE FUSO HORÁRIO ===");
  
  const dataTeste = new Date();
  const resultado = {
    dataOriginal: dataTeste.toString(),
    comFormatarDataBrasil: formatarDataBrasil(dataTeste),
    comUtilities: Utilities.formatDate(dataTeste, CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm:ss"),
    timezoneConfig: CONFIG.TIMEZONE
  };
  
  console.log("📊 Resultado do teste:", resultado);
  
  return resultado;
}

function testar() {
  return { 
    success: true, 
    message: "✅ Sistema funcionando!",
    timestamp: new Date().toISOString(),
    totalCadastros: buscarTodosCadastros().length
  };
}
