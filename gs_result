// CONFIGURAÇÕES
const CONFIG = {
  ID_PLANILHA: "1V4iGN14UpIQcwf3qKU0_Wbiy2exdW2WUmrYTniy0upA",
  ABA_PRINCIPAL: "Result",
  TIMEZONE: "America/Sao_Paulo"
};

// 🔥 ESTRUTURA DAS COLUNAS - ATUALIZADA COM SEUS NOMES
const ESTRUTURA_COLUNAS = {
  RAZAO_SOCIAL: 'Razão Social',
  NOME_FANTASIA: 'Nome Fantasia', 
  CNPJ: 'CNPJ',
  FORNECEDOR: 'Fornecedor',
  ULTIMA_ETAPA: 'Ultima etapa',
  ETAPA: 'Etapa',
  OBSERVACAO: 'Observação',
  CONTRATO_ENVIADO: 'Contrato Enviado',
  CONTRATO_ASSINADO: 'Contrato Assinado',
  ATIVACAO: 'Ativação',
  LINK: 'Link',
  MENSALIDADE: 'Mensalidade',
  MENSALIDADE_SIM: 'Mensalidade SIM',
  TARIFA: 'Tarifa',
  PERCENTUAL_TARIFA: '% Tarifa',
  ADESAO: 'Adesão',
  SITUACAO: 'Situação'
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

// 🔥🔥🔥 FUNÇÃO CORRIGIDA PARA HORÁRIO BRASIL - COM FUSO CORRETO
function formatarDataBrasil(data) {
  if (!data) return '';
  
  try {
    // Se já é string no formato brasileiro, retornar COMO ESTÁ
    if (typeof data === 'string' && data.includes('/') && data.includes(':')) {
      return data;
    }
    
    // Se é objeto Date, formatar CORRETAMENTE com fuso do Brasil
    if (data instanceof Date) {
      // 🔥 CORREÇÃO: Usar o fuso horário de Brasília corretamente
      const dataBrasil = Utilities.formatDate(data, CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm:ss");
      console.log("✅ Date convertido:", data.toString(), "→", dataBrasil);
      return dataBrasil;
    }
    
    // Para outros casos, tentar converter
    try {
      const dataObj = new Date(data);
      if (!isNaN(dataObj.getTime())) {
        const dataBrasil = Utilities.formatDate(dataObj, CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm:ss");
        return dataBrasil;
      }
    } catch (e) {
      return data.toString();
    }
    
    return data.toString();
    
  } catch (error) {
    console.error("❌ Erro em formatarDataBrasil:", error);
    return data ? data.toString() : '';
  }
}

// 🔥🔥🔥 FUNÇÃO SIMPLES - SEM COMPENSAR FUSO HORÁRIO
function formatarDataBrasilCorrigida(data) {
  if (!data) return '';
  
  try {
    // Se já está no formato correto, retornar como está
    if (typeof data === 'string' && data.includes('/') && data.includes(':')) {
      return data;
    }
    
    let dataObj;
    
    // Se já é Date, usar direto
    if (data instanceof Date) {
      dataObj = data;
    } else {
      // Tentar converter para Date
      dataObj = new Date(data);
      if (isNaN(dataObj.getTime())) {
        return data.toString();
      }
    }
    
    // 🔥🔥🔥 MÉTODO SIMPLES: Usar Utilities.formatDate sem compensações
    // Isso deve pegar automaticamente o fuso horário de Brasília
    const dataFormatada = Utilities.formatDate(dataObj, CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm:ss");
    
    console.log("🔥 DATA SIMPLES:", dataFormatada);
    
    return dataFormatada;
    
  } catch (error) {
    console.error("❌ Erro em formatarDataBrasilCorrigida:", error);
    return data ? data.toString() : '';
  }
}

// 🔥🔥🔥 VERSÃO MAIS SIMPLES - APENAS USANDO HORÁRIO LOCAL DO USUÁRIO
function formatarDataBrasilSimples() {
  const agora = new Date();
  
  // Usar métodos locais do JavaScript que pegam o fuso do usuário
  const dia = String(agora.getDate()).padStart(2, '0');
  const mes = String(agora.getMonth() + 1).padStart(2, '0');
  const ano = agora.getFullYear();
  const horas = String(agora.getHours()).padStart(2, '0');
  const minutos = String(agora.getMinutes()).padStart(2, '0');
  const segundos = String(agora.getSeconds()).padStart(2, '0');
  
  const dataFormatada = `${dia}/${mes}/${ano} ${horas}:${minutos}:${segundos}`;
  
  console.log("🔥 DATA SIMPLES (Local):", dataFormatada);
  
  return dataFormatada;
}



// 🔥 FUNÇÕES DE GERENCIAMENTO DE WAITLABELS
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

// 🔥 FUNÇÃO PRINCIPAL
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistema - Gestão de Cadastros')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 🔥 FUNÇÕES PRINCIPAIS COM WAITLABEL
function processarCadastroComWaitlabel(dados, waitlabel) {
  try {
    console.log("🎯 PROCESSAR CADASTRO COM WAITLABEL - Dados:", dados, "Waitlabel:", waitlabel);
    
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    let aba = ss.getSheetByName(waitlabel);

    if (!aba) {
      console.log("📝 Criando nova aba para waitlabel:", waitlabel);
      aba = ss.insertSheet(waitlabel);
      const cabecalho = [
        'Razão Social', 'Nome Fantasia', 'CNPJ', 'Fornecedor', 
        'Ultima etapa', 'Etapa',
        'Observação', 'Contrato Enviado', 'Contrato Assinado',
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
    
    // 🔥🔥🔥 VALIDAÇÃO DA ETAPA - USANDO FUNÇÃO AUXILIAR COM SITUAÇÃO
    console.log("🎯 Validando etapa selecionada no cadastro...");
    const validacaoEtapa = validarEtapa(dados.etapa, dados.situacao);
    if (!validacaoEtapa.valida) {
      return { success: false, message: validacaoEtapa.mensagem };
    }
    const etapaValidada = validacaoEtapa.etapa;
    
    // Verificar duplicatas
    const fornecedoresParaCadastrar = dados.fornecedores || [];
    const fornecedoresDuplicados = [];
    const cadastrosExistentes = buscarTodosCadastrosPorCNPJComWaitlabel(dados.cnpj, waitlabel);
    
    for (let fornecedor of fornecedoresParaCadastrar) {
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

    const ultimaLinha = aba.getLastRow();
    let linhaInserir = Math.max(2, ultimaLinha + 1);
    const resultados = [];
    let registrosCriados = 0;

    let situacaoParaSalvar = normalizarTexto(dados.situacao) || 'NOVO REGISTRO';
    if (situacaoParaSalvar === 'Novo registro') {
      situacaoParaSalvar = 'Novo Registro';
    }

    for (let i = 0; i < dados.fornecedores.length; i++) {
      const fornecedorObj = dados.fornecedores[i];
      
      let nomeFornecedor = '';
      let tarifaFornecedor = '';
      let percentualTarifaFornecedor = '0%';
      
      if (typeof fornecedorObj === 'object' && fornecedorObj !== null) {
        nomeFornecedor = fornecedorObj.nome || '';
        tarifaFornecedor = fornecedorObj.tarifa || '';
        percentualTarifaFornecedor = fornecedorObj.percentual_tarifa || '0%';
      }

      if (!nomeFornecedor || nomeFornecedor.trim() === '') {
        resultados.push(`❌ Fornecedor sem nome - pulado`);
        continue;
      }

      let mensalidadeNumero = parseFloat(dados.mensalidade) || 0;
      let adesaoNumero = processarAdesaoParaSalvar(dados.adesao);

      // 🔥🔥🔥 CORREÇÃO: USAR A MESMA FUNÇÃO SIMPLES
      const dataUltimaEtapa = formatarDataBrasilSimples();

      let dataAtivacaoParaSalvar = '';
      if (dados.ativacao && dados.ativacao.trim() !== '') {
        try {
          const dataUsuario = new Date(dados.ativacao);
          dataUsuario.setDate(dataUsuario.getDate() + 1);
          dataAtivacaoParaSalvar = Utilities.formatDate(dataUsuario, CONFIG.TIMEZONE, "dd/MM/yyyy");
        } catch (e) {
          console.error("❌ Erro ao processar data do usuário:", e);
          dataAtivacaoParaSalvar = '';
        }
      }

      const linhaDados = [
        normalizarTexto(dados.razao_social) || '',
        normalizarTexto(dados.nome_fantasia) || '',
        dados.cnpj ? dados.cnpj.toString() : '',
        normalizarTexto(nomeFornecedor),
        dataUltimaEtapa, // 🔥 AGORA COM HORÁRIO CORRETO
        etapaValidada, // 🔥 USANDO A ETAPA JÁ VALIDADA
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

      try {
        const range = aba.getRange(linhaInserir, 1, 1, linhaDados.length);
        range.setValues([linhaDados]);
        
        // Formatar colunas
        aba.getRange(linhaInserir, 12).setNumberFormat('"R$"#,##0.00');
        aba.getRange(linhaInserir, 13).setNumberFormat('"R$"#,##0.00');
        aba.getRange(linhaInserir, 15).setNumberFormat('0.00%');
        aba.getRange(linhaInserir, 16).setNumberFormat('"R$"#,##0.00');
        aba.getRange(linhaInserir, 14).setNumberFormat('@');
        aba.getRange(linhaInserir, 10).setNumberFormat('dd/MM/yyyy');
        
        SpreadsheetApp.flush();
        
        linhaInserir++;
        registrosCriados++;
        resultados.push(`✅ ${nomeFornecedor} - ${tarifaFornecedor} ${percentualTarifaFornecedor}`);
        
      } catch (erroInsercao) {
        console.error(`❌ Erro ao salvar:`, erroInsercao);
        resultados.push(`❌ ${nomeFornecedor} - ERRO: ${erroInsercao.message}`);
      }
    }

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
    
    const linhaAtualizar = parseInt(dados.id);

    if (linhaAtualizar < 2 || linhaAtualizar > aba.getLastRow()) {
      return { success: false, message: "Registro não encontrado" };
    }

    // 🔥🔥🔥 VALIDAÇÃO DA ETAPA - USANDO FUNÇÃO AUXILIAR COM SITUAÇÃO
    console.log("🎯 Validando etapa selecionada...");
    const validacaoEtapa = validarEtapa(dados.etapa, dados.situacao);
    if (!validacaoEtapa.valida) {
      return { success: false, message: validacaoEtapa.mensagem };
    }
    const etapaNova = validacaoEtapa.etapa;

    const dadosAtuais = aba.getRange(linhaAtualizar, 1, 1, 17).getValues()[0];
    const dataAtivacaoOriginal = dadosAtuais[9];
    const etapaAtual = dadosAtuais[5]?.toString().trim() || '';
    const situacaoAtual = dadosAtuais[16]?.toString().trim() || '';
    const dataUltimaEtapaAtual = dadosAtuais[4];
    
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

    let mensalidadeNumero = converterMoedaParaNumero(dados.mensalidade);
    let adesaoNumero = processarAdesaoParaSalvar(dados.adesao);

    const situacaoValida = (dados.situacao && dados.situacao.trim() !== '') ? dados.situacao : 'Novo registro';

    let dataAtivacaoParaSalvar = dataAtivacaoOriginal;
    if (dados.ativacao && dados.ativacao.trim() !== '') {
      try {
        const dataUsuario = new Date(dados.ativacao);
        dataUsuario.setDate(dataUsuario.getDate() + 1);
        dataAtivacaoParaSalvar = Utilities.formatDate(dataUsuario, CONFIG.TIMEZONE, "dd/MM/yyyy");
      } catch (e) {
        console.error("❌ Erro ao processar data:", e);
        dataAtivacaoParaSalvar = dataAtivacaoOriginal;
      }
    } else {
      if (dataAtivacaoOriginal instanceof Date) {
        dataAtivacaoParaSalvar = Utilities.formatDate(dataAtivacaoOriginal, CONFIG.TIMEZONE, "dd/MM/yyyy");
      }
    }

    const situacaoNova = normalizarTexto(situacaoValida);
    
    const mudouEtapa = etapaAtual !== etapaNova;
    const mudouSituacao = situacaoAtual !== situacaoNova;
    
    let dataUltimaEtapaParaSalvar = dataUltimaEtapaAtual;
    
    // 🔥🔥🔥 CORREÇÃO: USAR A MESMA FUNÇÃO DO "APLICAR A TODOS"
    if (mudouEtapa || mudouSituacao) {
      const dataAtual = new Date();
      
      // 🔥 USAR A MESMA FUNÇÃO SIMPLES QUE O "APLICAR A TODOS" USA
      dataUltimaEtapaParaSalvar = formatarDataBrasilSimples();
      
      console.log("🔄 ETAPA OU SITUAÇÃO MUDOU - ATUALIZANDO DATA DA ÚLTIMA ETAPA");
      console.log("📅 NOVA DATA (HORÁRIO BRASIL):", dataUltimaEtapaParaSalvar);
    }

    const novosDados = [
      normalizarTexto(dados.razao_social) || '',
      normalizarTexto(dados.nome_fantasia) || '',
      dados.cnpj ? dados.cnpj.toString() : '',
      normalizarTexto(fornecedorParaAtualizar),
      dataUltimaEtapaParaSalvar, // 🔥 AGORA COM HORÁRIO CORRETO
      etapaNova, // 🔥 USANDO A ETAPA JÁ VALIDADA
      normalizarTexto(dados.observacoes) || '',
      normalizarTexto(dados.contrato_enviado) || '',
      normalizarTexto(dados.contrato_assinado) || '',
      dataAtivacaoParaSalvar,
      dados.link || '',
      mensalidadeNumero,                                    
      converterMoedaParaNumero(dados.mensalidade_sim) || 0, 
      tarifaParaAtualizar || '',                            
      percentualParaAtualizar,                              
      adesaoNumero,                                         
      normalizarTexto(situacaoValida)                       
    ];

    aba.getRange(linhaAtualizar, 1, 1, novosDados.length).setValues([novosDados]);
    
    // Aplicar formatação
    aba.getRange(linhaAtualizar, 12).setNumberFormat('"R$"#,##0.00');
    aba.getRange(linhaAtualizar, 13).setNumberFormat('"R$"#,##0.00');
    aba.getRange(linhaAtualizar, 15).setNumberFormat('0.00%');
    aba.getRange(linhaAtualizar, 16).setNumberFormat('"R$"#,##0.00');
    aba.getRange(linhaAtualizar, 14).setNumberFormat('@');
    aba.getRange(linhaAtualizar, 10).setNumberFormat('dd/MM/yyyy');

    SpreadsheetApp.flush();

    return { 
      success: true, 
      message: `✅ "${dados.razao_social}" atualizado com sucesso no ${waitlabel}!` + 
               (mudouEtapa || mudouSituacao ? ' (Data da última etapa atualizada)' : '')
    };

  } catch (error) {
    console.error("❌ Erro em atualizarCadastroComWaitlabel:", error);
    return { success: false, message: "Erro ao atualizar: " + error.message };
  }
}

function aplicarAlteracoesATodos(cnpj, dados, camposSelecionados) {
  console.log("🎯 APLICAR A TODOS - VERSÃO COM VALIDAÇÃO INTELIGENTE");
  
  try {
    // 🔥🔥🔥 VALIDAÇÃO DA ETAPA - CONSIDERANDO A SITUAÇÃO
    console.log("🎯 Validando etapa no Aplicar a Todos...");
    const etapaParaValidar = dados.etapa;
    const situacaoParaValidar = dados.situacao || '';
    
    // Se está tentando aplicar uma nova etapa, validar ela
    if (camposSelecionados.includes('etapa') || camposSelecionados.includes('inputEtapaSearch')) {
      console.log("🔍 Validando NOVA etapa...");
      const validacaoEtapa = validarEtapa(etapaParaValidar, situacaoParaValidar);
      if (!validacaoEtapa.valida) {
        return { success: false, message: validacaoEtapa.mensagem };
      }
      console.log("✅ Nova etapa válida");
    } else {
      // 🔥🔥🔥 VALIDAÇÃO INTELIGENTE: Só validar etapa se a situação for EM ANDAMENTO
      const situacaoNormalizada = normalizarTexto(situacaoParaValidar);
      const ehEmAndamento = situacaoNormalizada === 'EM ANDAMENTO';
      
      if (ehEmAndamento) {
        console.log("🔍 Situação é EM ANDAMENTO - validando etapa atual do formulário...");
        const validacaoEtapa = validarEtapa(etapaParaValidar, situacaoParaValidar);
        if (!validacaoEtapa.valida) {
          return { 
            success: false, 
            message: `❌ OPERAÇÃO BLOQUEADA!\n\nPara situações "EM ANDAMENTO" a etapa é obrigatória.\n\nA etapa atual no formulário ("${etapaParaValidar}") não é válida.\n\nCorrija a etapa para uma das opções válidas:\n\n• PENDENTE FORNECEDOR(ES)\n• PENDENTE SIM\n• PENDENTE WL\n• PENDENTE CLÍNICA/LOJA` 
          };
        }
        console.log("✅ Etapa atual do formulário é válida para EM ANDAMENTO");
      } else {
        console.log("✅ Situação não é EM ANDAMENTO - etapa não precisa ser validada");
      }
    }

    const waitlabelAtual = getWaitlabelAtual();
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabelAtual);
    
    if (!aba) {
      return { success: false, message: "❌ Planilha não encontrada" };
    }
    
    const dadosCompletos = aba.getDataRange().getValues();
    const cabecalhos = dadosCompletos[0];
    
    const cnpjIndex = cabecalhos.indexOf("CNPJ");
    const ultimaEtapaIndex = cabecalhos.indexOf("Ultima etapa");
    const etapaIndex = cabecalhos.indexOf("Etapa");
    const situacaoIndex = cabecalhos.indexOf("Situação");
    
    console.log("🎯 Índices: UltimaEtapa=" + ultimaEtapaIndex);
    
    // BUSCAR REGISTROS
    const cnpjBuscado = cnpj.toString().replace(/\D/g, '');
    const registrosParaAtualizar = [];
    
    for (let i = 1; i < dadosCompletos.length; i++) {
      const linha = dadosCompletos[i];
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      const cnpjLinha = linha[cnpjIndex]?.toString().replace(/\D/g, '') || '';
      
      if (cnpjLinha === cnpjBuscado) {
        registrosParaAtualizar.push({
          linhaNumero: i + 1,
          dadosOriginais: linha
        });
      }
    }
    
    console.log(`🔍 Encontrados ${registrosParaAtualizar.length} registros`);
    
    // 🔥🔥🔥 VALIDAÇÃO ADICIONAL: Verificar se algum registro atual tem etapa inválida
    console.log("🔍 Verificando etapas existentes nos registros...");
    for (const registro of registrosParaAtualizar) {
      const etapaExistente = registro.dadosOriginais[etapaIndex]?.toString().trim() || '';
      if (etapaExistente) {
        const etapasValidas = ["PENDENTE FORNECEDOR(ES)", "PENDENTE SIM", "PENDENTE WL", "PENDENTE CLÍNICA/LOJA"];
        const etapaNormalizada = normalizarTexto(etapaExistente);
        
        if (!etapasValidas.includes(etapaNormalizada)) {
          console.log(`⚠️ Registro linha ${registro.linhaNumero} tem etapa inválida: "${etapaExistente}"`);
        }
      }
    }
    
    const mudouEtapa = camposSelecionados.includes('etapa') || camposSelecionados.includes('inputEtapaSearch');
    const mudouSituacao = camposSelecionados.includes('situacao');
    
    console.log("🔄 Mudanças: Etapa=" + mudouEtapa + ", Situacao=" + mudouSituacao);
    
    let registrosAtualizados = 0;
    let atualizouDataUltimaEtapa = false;
    
    for (const registro of registrosParaAtualizar) {
      console.log(`🔄 Atualizando linha ${registro.linhaNumero}...`);
      
      const novosDados = [...registro.dadosOriginais];
      
      // APLICAR ALTERAÇÕES
      for (const campo of camposSelecionados) {
        const valor = obterValorParaAplicarTodos(campo, dados);
        
        switch(campo) {
          case 'razao_social': novosDados[0] = valor; break;
          case 'nome_fantasia': novosDados[1] = valor; break;
          case 'cnpj_cadastro': novosDados[2] = valor; break;
          case 'etapa':
          case 'inputEtapaSearch': 
            novosDados[etapaIndex] = valor; 
            break;
          case 'situacao':
            novosDados[situacaoIndex] = valor;
            break;
          case 'observacoes': novosDados[6] = valor; break;
          case 'contrato_enviado': novosDados[7] = valor; break;
          case 'contrato_assinado': novosDados[8] = valor; break;
          case 'ativacao': novosDados[9] = valor; break;
          case 'link': novosDados[10] = valor; break;
          case 'mensalidade': novosDados[11] = valor; break;
          case 'mensalidade_sim': novosDados[12] = valor; break;
          case 'adesao': novosDados[15] = valor; break;
        }
      }
      
      // 🔥🔥🔥 ATUALIZAR DATA DA ÚLTIMA ETAPA SE MUDOU ETAPA/SITUAÇÃO
      if (ultimaEtapaIndex !== -1 && (mudouEtapa || mudouSituacao)) {
        const dataAtual = new Date();
        
        // 🔥 USAR A FORMATAÇÃO CORRIGIDA PARA HORÁRIO BRASIL
        const dataFormatada = formatarDataBrasilSimples();
        novosDados[ultimaEtapaIndex] = dataFormatada; // 🔥 AGORA COM HORÁRIO CORRETO
        
        atualizouDataUltimaEtapa = true;
        console.log(`   📅📅📅 DATA ATUALIZADA (HORÁRIO BRASIL): ${dataFormatada}`);
      }
      
      // SALVAR
      aba.getRange(registro.linhaNumero, 1, 1, novosDados.length).setValues([novosDados]);
      registrosAtualizados++;
    }
    
    SpreadsheetApp.flush();
    
    return {
      success: true,
      registrosAtualizados: registrosAtualizados,
      message: `✅ ${registrosAtualizados} registro(s) atualizado(s) com sucesso!` +
               (atualizouDataUltimaEtapa ? ' (Data da última etapa atualizada)' : '')
    };
    
  } catch (error) {
    console.error("❌ ERRO:", error);
    return { success: false, message: "❌ Erro: " + error.toString() };
  }
}

function obterValorParaAplicarTodos(campo, dados) {
  switch(campo) {
    case 'razao_social':
      return normalizarTexto(dados.razao_social) || '';
    case 'nome_fantasia':
      return normalizarTexto(dados.nome_fantasia) || '';
    case 'cnpj_cadastro':
      return dados.cnpj ? dados.cnpj.toString() : '';
    case 'etapa':
    case 'inputEtapaSearch':
      // 🔥 GARANTIR QUE A ETAPA SEJA NORMALIZADA CORRETAMENTE
      let etapa = normalizarTexto(dados.etapa) || '';
      // Se estiver vazia, não aplicar
      if (!etapa) return '';
      return etapa;
    case 'observacoes':
      return normalizarTexto(dados.observacoes) || '';
    case 'contrato_enviado':
      return normalizarTexto(dados.contrato_enviado) || '';
    case 'contrato_assinado':
      return normalizarTexto(dados.contrato_assinado) || '';
    case 'ativacao':
      return dados.ativacao || '';
    case 'link':
      return dados.link || '';
    case 'mensalidade':
      return converterMoedaParaNumero(dados.mensalidade) || 0;
    case 'mensalidade_sim':
      return converterMoedaParaNumero(dados.mensalidade_sim) || 0;
    case 'adesao':
      return processarAdesaoParaSalvar(dados.adesao);
    case 'situacao':
      let situacao = normalizarTexto(dados.situacao) || 'NOVO REGISTRO';
      if (situacao === 'NOVO REGISTRO') situacao = 'Novo Registro';
      return situacao;
    default:
      return '';
  }
}

// 🔥 FUNÇÃO AUXILIAR PARA VALIDAR ETAPAS - VERSÃO COM SITUAÇÃO
function validarEtapa(etapa, situacao) {
  // 🔥 SE NÃO FOR "EM ANDAMENTO", ETAPA NÃO É OBRIGATÓRIA
  const situacaoNormalizada = normalizarTexto(situacao || '');
  const naoEhEmAndamento = situacaoNormalizada !== 'EM ANDAMENTO';
  
  if (naoEhEmAndamento) {
    console.log("✅ Situação não é EM ANDAMENTO - etapa não é obrigatória");
    return { valida: true, etapa: etapa ? normalizarTexto(etapa) : '' };
  }
  
  // 🔥 SE É "EM ANDAMENTO", ENTÃO ETAPA É OBRIGATÓRIA
  if (!etapa || etapa.trim() === '') {
    return { 
      valida: false, 
      mensagem: '❌ Para situações "EM ANDAMENTO" o campo Etapa é obrigatório!' 
    };
  }
  
  const etapasValidas = [
    "PENDENTE FORNECEDOR(ES)",
    "PENDENTE SIM", 
    "PENDENTE WL",
    "PENDENTE CLÍNICA/LOJA"
  ];
  
  const etapaNormalizada = normalizarTexto(etapa);
  
  // 🔥 BLOQUEAR EXPLICITAMENTE "DESISTIU" E OUTRAS ETAPAS INVÁLIDAS
  const etapasBloqueadas = ["DESISTIU", "REJEITADO", "CADASTRADO", "NOVO REGISTRO", "EM ANDAMENTO", "DESCREDENCIADO"];
  
  if (etapasBloqueadas.includes(etapaNormalizada)) {
    const mensagemErro = `❌ ETAPA NÃO PERMITIDA!\n\nA etapa "${etapa}" é uma SITUAÇÃO, não uma etapa do processo.\n\n📋 ETAPAS VÁLIDAS (do processo):\n• PENDENTE FORNECEDOR(ES)\n• PENDENTE SIM\n• PENDENTE WL\n• PENDENTE CLÍNICA/LOJA\n\n💡 Use o campo "Situação" para: ${etapa}`;
    
    return { valida: false, mensagem: mensagemErro };
  }
  
  if (!etapasValidas.includes(etapaNormalizada)) {
    const mensagemErro = `❌ ETAPA INVÁLIDA!\n\nA etapa "${etapa}" não é válida.\n\n📋 ETAPAS VÁLIDAS:\n• PENDENTE FORNECEDOR(ES)\n• PENDENTE SIM\n• PENDENTE WL\n• PENDENTE CLÍNICA/LOJA\n\nSelecione uma das etapas acima para continuar.`;
    
    return { valida: false, mensagem: mensagemErro };
  }
  
  return { valida: true, etapa: etapaNormalizada };
}

// 🔥 FUNÇÕES DE BUSCA
function buscarTodosCadastrosComWaitlabel(waitlabel) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabel);
    if (!aba) return [];
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) return [];
    
    const dados = aba.getRange(2, 1, ultimaLinha - 1, 17).getValues();
    const cadastros = [];
    
    for (let i = 0; i < dados.length; i++) {
      const linha = dados[i];
      if (!linha[0] || linha[0].toString().trim() === '') continue;

      let ultimaEtapaFormatada = '';
      if (linha[4] && linha[4] instanceof Date) {
        ultimaEtapaFormatada = formatarDataBrasil(linha[4]);
      } else if (linha[4]) {
        ultimaEtapaFormatada = linha[4].toString();
      }
      
      let ativacaoFormatada = '';
      if (linha[9] && linha[9] instanceof Date) {
        ativacaoFormatada = Utilities.formatDate(linha[9], CONFIG.TIMEZONE, "dd/MM/yyyy");
      } else if (linha[9]) {
        ativacaoFormatada = linha[9].toString();
      }
      
      const cadastro = {
        id: i + 2,
        razao_social: linha[0]?.toString().trim() || '',
        nome_fantasia: linha[1]?.toString().trim() || '',
        cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''),
        fornecedor: linha[3]?.toString().trim() || '',
        ultima_etapa: ultimaEtapaFormatada,
        etapa: linha[5]?.toString().trim() || '',
        observacoes: linha[6]?.toString().trim() || '',
        contrato_enviado: linha[7]?.toString().trim() || '',
        contrato_assinado: linha[8]?.toString().trim() || '',
        ativacao: ativacaoFormatada,
        link: linha[10]?.toString().trim() || '',
        mensalidade: parseFloat(linha[11]) || 0,
        mensalidade_sim: parseFloat(linha[12]) || 0,
        tarifa: linha[13]?.toString().trim() || '',
        percentual_tarifa: linha[14]?.toString().trim() || '',
        adesao: processarAdesao(linha[15]),
        situacao: (linha[16]?.toString().trim() || 'Novo registro'),
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
      if (!linha[0] || linha[0].toString().trim() === '') continue;
      
      const cnpjCadastro = linha[2]?.toString().replace(/\D/g, '') || '';
      
      if (cnpjCadastro === cnpjBuscado) {
        let ultimaEtapaFormatada = '';
        if (linha[4] && linha[4] instanceof Date) {
          ultimaEtapaFormatada = formatarDataBrasil(linha[4]);
        } else if (linha[4]) {
          ultimaEtapaFormatada = linha[4].toString();
        }
        
        let ativacaoFormatada = '';
        if (linha[9] && linha[9] instanceof Date) {
          ativacaoFormatada = Utilities.formatDate(linha[9], CONFIG.TIMEZONE, "dd/MM/yyyy");
        } else if (linha[9]) {
          ativacaoFormatada = linha[9].toString();
        }
        
        const cadastro = {
          id: i + 2,
          razao_social: linha[0]?.toString().trim() || '',
          nome_fantasia: linha[1]?.toString().trim() || '',
          cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''),
          fornecedor: linha[3]?.toString().trim() || '',
          ultima_etapa: ultimaEtapaFormatada,
          etapa: linha[5]?.toString().trim() || '',
          observacoes: linha[6]?.toString().trim() || '',
          contrato_enviado: linha[7]?.toString().trim() || '',
          contrato_assinado: linha[8]?.toString().trim() || '',
          ativacao: ativacaoFormatada,
          link: linha[10]?.toString().trim() || '',
          mensalidade: parseFloat(linha[11]) || 0,
          mensalidade_sim: parseFloat(linha[12]) || 0,
          tarifa: linha[13]?.toString().trim() || '',
          percentual_tarifa: linha[14]?.toString().trim() || '',
          adesao: processarAdesao(linha[15]),
          situacao: (linha[16]?.toString().trim() || 'Novo registro'),
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
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabel);
    if (!aba) return { encontrado: false, mensagem: "Waitlabel não encontrado" };
    
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < id) return { encontrado: false, mensagem: "Registro não encontrado" };
    
    const linha = aba.getRange(id, 1, 1, 17).getValues()[0];
    
    if (!linha[0] || linha[0].toString().trim() === '') {
      return { encontrado: false, mensagem: "Registro vazio ou não encontrado" };
    }

    let ultimaEtapaFormatada = '';
    if (linha[4] && linha[4] instanceof Date) {
      ultimaEtapaFormatada = formatarDataBrasil(linha[4]);
    } else if (linha[4]) {
      ultimaEtapaFormatada = linha[4].toString();
    }
    
    let ativacaoFormatada = '';
    if (linha[9] && linha[9] instanceof Date) {
      ativacaoFormatada = Utilities.formatDate(linha[9], CONFIG.TIMEZONE, "yyyy-MM-dd");
    } else if (linha[9]) {
      if (linha[9].includes('/')) {
        const partes = linha[9].split('/');
        ativacaoFormatada = `${partes[2]}-${partes[1]}-${partes[0]}`;
      } else {
        ativacaoFormatada = linha[9].toString();
      }
    }

    let tarifa = linha[13]?.toString().trim() || '';
    let percentualTarifa = '0%';
    if (linha[14] !== null && linha[14] !== undefined && linha[14] !== '') {
      const valor = parseFloat(linha[14]);
      if (!isNaN(valor)) {
        percentualTarifa = (valor * 100).toFixed(2) + '%';
      } else {
        percentualTarifa = linha[14]?.toString().trim() || '0%';
      }
    }

    const fornecedorParaFormulario = {
      nome: linha[3]?.toString().trim() || '',
      tarifa: tarifa,
      percentual_tarifa: percentualTarifa
    };
    
    const resultado = {
      encontrado: true,
      id: id,
      razao_social: linha[0]?.toString().trim() || '',
      nome_fantasia: linha[1]?.toString().trim() || '',
      cnpj: formatarCNPJNoSheets(linha[2]?.toString().trim() || ''),
      fornecedor: linha[3]?.toString().trim() || '',
      fornecedores: [fornecedorParaFormulario],
      ultima_etapa: ultimaEtapaFormatada,
      etapa: linha[5]?.toString().trim() || '',
      observacoes: linha[6]?.toString().trim() || '',
      contrato_enviado: linha[7]?.toString().trim() || '',
      contrato_assinado: linha[8]?.toString().trim() || '',
      ativacao: ativacaoFormatada,
      link: linha[10]?.toString().trim() || '',
      mensalidade: parseFloat(linha[11]) || 0,
      mensalidade_sim: parseFloat(linha[12]) || 0,
      tarifa: tarifa,
      percentual_tarifa: percentualTarifa,
      adesao: processarAdesao(linha[15]),
      situacao: (linha[16]?.toString().trim() || 'Novo registro'),
      waitlabel: waitlabel
    };

    return resultado;
    
  } catch (error) {
    console.error("❌ Erro em buscarCadastroPorIDComWaitlabel:", error);
    return { encontrado: false, mensagem: "Erro: " + error.message };
  }
}

// 🔥 FUNÇÕES AUXILIARES
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

// 🔥 FUNÇÕES DE EXCLUSÃO
function excluirTodosFornecedoresCNPJ(cnpj) {
  try {
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
    
    linhasParaExcluir.forEach(linha => {
      try {
        aba.deleteRow(linha);
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
    
    return contador;
    
  } catch (error) {
    console.error("❌ Erro em contarRegistrosPorCNPJ:", error);
    return 0;
  }
}

// 🔥 FUNÇÃO DE TESTE
function testar() {
  return { 
    success: true, 
    message: "✅ Sistema funcionando!",
    timestamp: new Date().toISOString()
  };
}

// 🔥 FUNÇÃO PARA VERIFICAR COLUNAS
function verificarColunas() {
  try {
    const waitlabelAtual = getWaitlabelAtual();
    const ss = SpreadsheetApp.openById(CONFIG.ID_PLANILHA);
    const aba = ss.getSheetByName(waitlabelAtual);
    
    if (!aba) return { error: "Aba não encontrada" };
    
    const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
    
    console.log("🔍 COLUNAS ENCONTRADAS:");
    cabecalhos.forEach((cabecalho, index) => {
      const letraColuna = String.fromCharCode(65 + index);
      console.log(`Coluna ${letraColuna} [${index}]: "${cabecalho}"`);
    });
    
    const ultimaEtapaIndex = cabecalhos.indexOf("Ultima etapa");
    console.log("🎯 Índice da coluna 'Ultima etapa':", ultimaEtapaIndex);
    
    return {
      cabecalhos: cabecalhos,
      ultimaEtapaIndex: ultimaEtapaIndex,
      encontrada: ultimaEtapaIndex !== -1
    };
    
  } catch (error) {
    console.error("❌ Erro:", error);
    return { error: error.message };
  }
}
