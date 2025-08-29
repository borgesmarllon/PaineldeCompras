// ===============================================
    // FUNÇÕES PARA PEDIDOS DE COMPRA
    // ===============================================

    /**
     * Retorna o próximo número sequencial para um novo pedido.
     * Cria a planilha 'Pedidos' se não existir.
     * @returns {string} O próximo número de pedido formatado como '0001'.
     */
    function getProximoNumeroPedido(empresaCodigo) {
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(30000)){
       throw new Error("Não foi póssivel obter o acesso para gerar um novo número de pedido. Tente novamente em alguns instantes."); 
      }
      try{
      const spreadsheet = SpreadsheetApp.getActive();
      let sheet = spreadsheet.getSheetByName('Pedidos');

      if (!sheet) {
        sheet = spreadsheet.insertSheet('Pedidos');
        const headers = [
          'Número do Pedido', 'ID da Empresa', 'Data', 'Fornecedor', 'CNPJ Fornecedor',
          'Endereço Fornecedor', 'Condição Pagamento Fornecedor', 'Forma Pagamento Fornecedor',
          'Placa Veiculo', 'Nome Veiculo', 'Observacoes', 'Total Geral', 'Status', 'Itens'
        ];
        sheet.appendRow(headers);
        sheet.getRange('A:B').setNumberFormat('@');
        return '000001';
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return '000001';
      }

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const colNumero = headers.findIndex(h => h.toUpperCase() === 'NÚMERO DO PEDIDO');
      const colEmpresa = headers.findIndex(h => ['ID DA EMPRESA', 'ID EMPRESA', 'EMPRESA'].includes(h.toUpperCase()));

      if (colEmpresa === -1 || colNumero === -1) {
        throw new Error('Cabeçalhos "ID da Empresa" ou "Número do Pedido" não encontrados na planilha "Pedidos".');
      }

      const idEmpresaColumn = sheet.getRange(2, colEmpresa + 1, lastRow - 1).getValues();
      const numeroPedidoColumn = sheet.getRange(2, colNumero + 1, lastRow - 1).getValues();

      const empresaCodigoTratado = String(empresaCodigo).trim();
      let maxNumero = 0;
      
      for (let i = 0; i < idEmpresaColumn.length; i++) {
      const idNaLinha = String(idEmpresaColumn[i][0]).trim();
      
      // A sua lógica de comparação com parseInt é excelente
      if (parseInt(idNaLinha, 10) === parseInt(empresaCodigoTratado, 10)) {
        const numeroAtual = parseInt(numeroPedidoColumn[i][0], 10);
        if (!isNaN(numeroAtual) && numeroAtual > maxNumero) {
          maxNumero = numeroAtual;
        }
      }
    }
    
    const proximoNumero = maxNumero > 0 ? maxNumero + 1 : 1;
    
    return proximoNumero.toString().padStart(6, '0');

  } finally {
    // Libera o "semáforo" para a próxima execução, aconteça o que acontecer.
    lock.releaseLock();
  }
}

/**
 * Retorna o objeto de configuração principal do sistema.
 * Colocada neste arquivo para garantir 100% de visibilidade.
 * @returns {Object} O objeto de configuração.
 */
function getConfig() {
  return {
    sheets: {
      pedidos: 'Pedidos',
      fornecedores: 'Fornecedores'
    },
    status: {
      aguardandoAprovacao: 'AGUARDANDO APROVACAO'
    }
  };
}

function normalizarDataPedido(data) {
    if (!data) return new Date();
    if (data instanceof Date) return data;
    // Tenta converter string para data
    try {
        return new Date(data.replace(' ', 'T'));
    } catch (e) {
        return new Date();
    }
}


/**
 * Salva um objeto de pedido em uma linha na planilha "Pedidos".
 * @param {Object} pedido O objeto contendo os detalhes do pedido.
 * @param {string} usuarioLogado O email do usuário que está realizando a ação.
 * @returns {{status: string, message: string}} Um objeto com o resultado da operação.
 */
function salvarPedido(pedido, usuarioLogado) {
  console.log('📋 === INÍCIO salvarPedido ===');
  console.log('📋 Objeto pedido recebido:', JSON.stringify(pedido, null, 2));

  try {
    const config = getConfig();
    const sheet = SpreadsheetApp.getActive().getSheetByName(config.sheets.pedidos);
    if (!sheet) {
      throw new Error('Planilha "Pedidos" não encontrada. Verifique o nome na função getConfig().');
    }

  const empresaId = pedido.empresaId || pedido.empresa;
  const numeroPedido = pedido.numeroPedido || pedido.numero;
  console.log(`Número do pedido gerado para a empresa ${empresaId}: ${numeroPedido}`);
  const idFornecedorParaBusca = String(pedido.fornecedorId || pedido.fornecedor);
  const dataObj = normalizarDataPedido(pedido.data);
  //const dataFinalParaFormatar = dataObj || new Date();

  const dataFormatada = Utilities.formatDate(
    dataObj,
    "America/Sao_Paulo",    // O fuso horário de referência
    "dd/MM/yyyy HH:mm:ss"   // O formato de texto que você quer na planilha
  );

  const mapaEmpresas = _criarMapaDeEmpresas();
  const mapaFornecedores = criarMapaDeFornecedores();
  let dadosFornecedor = {}; 
  const chaveDeBusca = String(pedido.fornecedorId || pedido.fornecedor).toUpperCase().trim();
  if (mapaFornecedores && mapaFornecedores.hasOwnProperty(chaveDeBusca)) {
      dadosFornecedor = mapaFornecedores[chaveDeBusca];
      Logger.log(`✅ Fornecedor encontrado: [Chave: ${chaveDeBusca}, Nome: ${dadosFornecedor.nome}]`);
  } else {
      Logger.log(`⚠️ Aviso: Fornecedor com a chave "${chaveDeBusca}" não encontrado.`);
  }
  const cidadeFornecedor = (dadosFornecedor.cidade || '').toUpperCase();
  let statusFinal = config.status.aguardandoAprovacao;

  if (cidadeFornecedor.includes("VITORIA DA CONQUISTA")) {
    statusFinal = "Aprovado"; // Sobreescreve para fornecedor local
    Logger.log('APROVAÇÃO AUTOMATICA: Fornecedor de Vitória da Conquista.');
  } else {
    Logger.log(`APROVAÇÃO MANUAL: Fornecedor de '${cidadeFornecedor}'. Status definido como '${statusFinal}'.`);
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const dataToSave = {
    'Número do Pedido': `'${numeroPedido}`,
    'Empresa': `'${empresaId}`,
    'Data': dataFormatada,
    'Fornecedor': dadosFornecedor.nome || pedido.fornecedor,
    'CNPJ Fornecedor': dadosFornecedor.cnpj || '',
    'Endereço Fornecedor': dadosFornecedor.endereco || '',
    'Estado Fornecedor': dadosFornecedor.estado || '',
    'Cidade Fornecedor': dadosFornecedor.cidade || '',
    'Condição Pagamento Fornecedor': dadosFornecedor.condicao || '',
    'Forma Pagamento Fornecedor': dadosFornecedor.forma || '',
    'Placa Veiculo': pedido.placaVeiculo,
    'Nome Veiculo': pedido.nomeVeiculo,
    'Observacoes': pedido.observacoes,
    'Total Geral': parseFloat(pedido.totalGeral) || 0,
    'ICMS ST Total': parseFloat(pedido.valorIcms) || 0,
    'Status': statusFinal,
    'Itens': JSON.stringify(pedido.itens || []),
    'Data Criacao': new Date().toISOString(), 
    'Produto Fornecedor': pedido.produtoFornecedor,
    'Usuario Criador': usuarioLogado,
    'Aliquota imposto': parseFloat(pedido.aliquotaImposto) || 0
  };

  const rowData = headers.map(header => dataToSave.hasOwnProperty(header) ? dataToSave[header] : '');
  sheet.appendRow(rowData);
  
  // ===== LOG DE DEPURAÇÃO ADICIONADO =====
    Logger.log("--- PREPARANDO PARA CHAMAR desmembrarJsonDeItens ---");

  // Após salvar o pedido principal, chama a função para desmembrar os itens.
    desmembrarJsonDeItens(
        numeroPedido, 
        empresaId, 
        dataToSave['Itens'],
        dataToSave['Estado Fornecedor'],
        dataToSave['Aliquota imposto'],
        dataToSave['ICMS ST Total']
    );
    Logger.log("--- RETORNOU DE desmembrarJsonDeItens ---");
        // Verificamos se o status final exige aprovação
          const listaAdmins = _getAdminUsers().data || [];  // já retorna lista direta

          Logger.log(`adminsResponse: ${JSON.stringify(listaAdmins)}`);

          if (listaAdmins.length > 0) {
            Logger.log(`Encontrados ${listaAdmins.length} administradores para notificação.`);

            const numeroPedidoLimpo = String(dataToSave['Número do Pedido']).replace(/'/g, '');
            const empresaIdLimpo = String(dataToSave['Empresa']).replace(/'/g, '');
            
            // Esta linha agora funciona, pois 'mapaEmpresas' foi definido no início.
            const nomeDaEmpresa = mapaEmpresas[empresaIdLimpo]?.empresa || `Empresa ID ${empresaIdLimpo}`;
            const urlBaseDoApp = ScriptApp.getService().getUrl();
            let mensagem;
            
            if (statusFinal === config.status.aguardandoAprovacao) {
            const linkDoPedido = `${urlBaseDoApp}?page=pedido&id=${numeroPedidoLimpo}&empresa=${empresaIdLimpo}`;
                mensagem = `🔔 <b>Novo Pedido para Aprovação!</b>\n\n` +
                           `<b>Empresa:</b> ${nomeDaEmpresa}\n` +
                           `<b>Nº Pedido:</b> ${numeroPedidoLimpo}\n` +
                           `<b>Criado por:</b> ${usuarioLogado}\n` +
                           `<b>Fornecedor:</b> ${dataToSave['Fornecedor']}\n` +
                           `<b>Valor Total:</b> ${dataToSave['Total Geral'].toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}`                            
            } else { // Caso de aprovação automática
                mensagem = `✅ <b>Novo Pedido Aprovado (Automático)</b>\n\n` +
                            `<b>Empresa:</b> ${nomeDaEmpresa}\n` +
                            `<b>Nº Pedido:</b> ${numeroPedidoLimpo}\n` +
                            `<b>Fornecedor:</b> ${dataToSave['Fornecedor']}\n` +
                            `<b>Valor Total:</b> ${dataToSave['Total Geral'].toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}\n\n` +
                            `O pedido foi criado e aprovado automaticamente.`;
            }
            
            let botoesDeAcao = null;

            if (statusFinal === config.status.aguardandoAprovacao) {
                botoesDeAcao = [
                    [
                        { text: "✅ Aprovar", callback_data: `a:${numeroPedidoLimpo}:${empresaIdLimpo}` },
                        { text: "❌ Rejeitar", callback_data: `r:${numeroPedidoLimpo}:${empresaIdLimpo}` }
                    ],
                    [
                        { text: "📄 Ver Detalhes", callback_data: `d:${numeroPedidoLimpo}:${empresaIdLimpo}` }
                    ]
                ];
            } else if (statusFinal === 'Aprovado') {
                botoesDeAcao = [
                    [
                        { text: "📄 Ver Detalhes", callback_data: `d:${numeroPedidoLimpo}:${empresaIdLimpo}` }
                    ]
                ];
            }
            // Envio centralizado
            listaAdmins.forEach((admin, index) => {
              if (index > 0) Utilities.sleep(1000);
              Logger.log(`Enviando notificação para chatId: ${admin.chatId}`);
              const respostaEnvio = enviarNotificacaoTelegram(admin.chatId, mensagem, botoesDeAcao);
              Logger.log(`Resposta da API Telegram para chatId ${admin.chatId}: ${respostaEnvio}`);
            });
          }

          return { status: 'ok', message: `Pedido ${numeroPedido} salvo com sucesso com status '${statusFinal}'!` };

        } catch (e) {
          Logger.log(`ERRO em salvarPedido: ${e.message}\nStack: ${e.stack}`);
          return { status: 'error', message: `Ocorreu um erro no servidor: ${e.message}` };
        }
}

function criarMapaDeFornecedores() {
    try {
        const config = getConfig();
        const fornecedoresSheet = SpreadsheetApp.getActive().getSheetByName(config.sheets.fornecedores);
        if (!fornecedoresSheet) {
            Logger.log("ERRO CRÍTICO: A aba de fornecedores não foi encontrada.");
            return {};
        }

        const data = fornecedoresSheet.getDataRange().getValues();
        const headers = data.shift();

        // Encontra o índice das colunas dinamicamente
        const nomeIndex = headers.indexOf("RAZAO SOCIAL");
        
        // <<< MUDANÇA 1 de 4 >>> Encontrar a coluna do ID/CÓDIGO.
        // !!! ATENÇÃO: Verifique se o nome da sua coluna de ID é "CÓDIGO". Se for "ID", troque abaixo.
        const idIndex = headers.indexOf("ID"); 
        
        const cidadeIndex = headers.indexOf("CIDADE");
        const cnpjIndex = headers.indexOf("CNPJ");
        const enderecoIndex = headers.indexOf("ENDERECO");
        const grupoIndex = headers.indexOf("GRUPO");
        const condicaoIndex = headers.indexOf("CONDICAO DE PAGAMENTO");
        const formaIndex = headers.indexOf("FORMA DE PAGAMENTO");
        const estadoIndex = headers.indexOf("ESTADO");
        const statusIndex = headers.indexOf("STATUS");
        
        // <<< MUDANÇA 2 de 4 >>> Adicionar verificação para a coluna de ID.
        if (nomeIndex === -1 || idIndex === -1) {
            Logger.log("ERRO CRÍTICO: Não foi possível encontrar as colunas 'RAZAO SOCIAL' e/ou 'CÓDIGO' na aba 'Fornecedores'.");
            return {};
        }

        const mapa = {};

        data.forEach(row => {
            const status = (statusIndex !== -1) ? String(row[statusIndex] || '').toUpperCase().trim() : 'ATIVO';
            
            if (status === 'ATIVO') {
                // <<< MUDANÇA 3 de 4 >>> Usar o ID como chave do mapa, não mais o nome.
                const idFornecedor = String(row[idIndex] || '').trim();

                // Só adiciona ao mapa se a linha tiver um ID válido.
                if (idFornecedor) {
                    mapa[idFornecedor] = {
                        id: idFornecedor, // Adiciona o ID também dentro do objeto para referência
                        nome: row[nomeIndex],
                        cidade: row[cidadeIndex] || '',
                        cnpj: row[cnpjIndex] || '',
                        endereco: row[enderecoIndex] || '',
                        grupo: row[grupoIndex] || '',
                        condicao: row[condicaoIndex] || '',
                        forma: row[formaIndex] || '',
                        estado: row[estadoIndex] || '',
                    };
                }
            }
        });
        
        // <<< MUDANÇA 4 de 4 >>> Atualizar a mensagem de log para refletir a mudança.
        Logger.log(`Mapa de fornecedores (por ID) criado com sucesso. Total de entradas: ${Object.keys(mapa).length}`);
        
        // Opcional: Descomente a linha abaixo para depurar e ver o novo mapa.
        // Logger.log(JSON.stringify(mapa, null, 2)); 
        
        return mapa;

    } catch (e) {
        Logger.log(`ERRO FATAL em criarMapaDeFornecedores: ${e.message}`);
        return {};
    }
}

/**
 * Busca todos os dados de um pedido específico para exibição na tela de impressão.
 * @param {string} numeroPedido - O número do pedido a ser buscado.
 * @returns {Object|null} Objeto com todos os dados do pedido, ou null se não encontrado.
 */
function getDadosPedidoParaImpressao(numeroPedido) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log(`Planilha 'Pedidos' vazia ou não encontrada ao buscar pedido ${numeroPedido}.`);
    return null;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  Logger.log(`[getDadosPedidoParaImpressao] Buscando pedido: "${numeroPedido}"`);

  const pedidoRow = dados.find(row => {
    const sheetNumeroPedido = String(row[0]).trim(); // Pega o valor da primeira coluna e remove espaços
    Logger.log(`[getDadosPedidoParaImpressao] Comparando "${sheetNumeroPedido}" (na planilha) com "${String(numeroPedido).trim()}" (recebido).`);
    return sheetNumeroPedido === String(numeroPedido).trim();
  });
  
  if (!pedidoRow) {
    Logger.log(`Pedido "${numeroPedido}" não encontrado na planilha após a busca.`);
    return null;
  }

  const pedidoData = {};
  headers.forEach((header, index) => {
    const camelCaseHeader = toCamelCase(header); // Usa a função toCamelCase para padronizar
    let value = pedidoRow[index];

    // Converte objetos Date para string formatada (YYYY-MM-DD)
    if (value instanceof Date) {
      value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
    pedidoData[camelCaseHeader] = value;
  });

  try {
    pedidoData.itens = JSON.parse(pedidoData.itens || '[]');
  } catch (e) {
    Logger.log(`Erro ao parsear itens JSON para pedido ${numeroPedido}: ${e.message}`);
    pedidoData.itens = [];
  }

  Logger.log(`[getDadosPedidoParaImpressao] Pedido encontrado e processado: ${JSON.stringify(pedidoData)}`);
  return pedidoData;
}


/**
 * Busca pedidos na planilha com base em múltiplos critérios.
 * @param {object} params Objeto com os parâmetros de busca.
 * @param {string} [params.mainSearch] Termo para buscar em "Número do Pedido" ou "Fornecedor".
 * @param {string} [params.dateStart] Data inicial no formato YYYY-MM-DD.
 * @param {string} [params.dateEnd] Data final no formato YYYY-MM-DD.
 * @param {string} [params.plateSearch] Placa do veículo a ser buscada.
 * @param {string} [params.usuarioCriador] O nome do usuário criador para filtrar (apenas admin).
 * @returns {object} Um objeto com o status da operação e os dados dos pedidos encontrados.
 */
function buscarPedidosv2(params) {
  try {
    Logger.log("Iniciando busca com parâmetros: " + JSON.stringify(params));
    
    // --- LÓGICA DE EMPRESA REFORÇADA ---
    if (!params || !params.empresaId) {
      Logger.log("Busca interrompida: ID da empresa é obrigatório.");
      return { status: 'success', data: [] };
    }
    
    const sheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName('Pedidos');
    if (!sheet) { throw new Error("Aba 'Pedidos' não encontrada."); }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'success', data: [] };

    const lastCol = sheet.getLastColumn();
    const all = sheet.getRange(1, 1, lastRow, lastCol).getValues(); // inclui header
    const rawHeaders = all[0] || [];

    // Normaliza header (remove acentos, trim, uppercase)
    const normalizeHeader = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().trim();
    const headers = rawHeaders.map(h => normalizeHeader(h));

    // Busca índice robusto por palavras-chave
    const findIndexByKeywords = (keywords) => {
      for (let i = 0; i < headers.length; i++) {
        for (let k = 0; k < keywords.length; k++) {
          if (headers[i].includes(keywords[k])) return i;
        }
      }
      return -1;
    };
    const mapaFornecedores = criarMapaDeFornecedores();
    //const data = sheet.getDataRange().getValues();
    //const headers = data.shift();

    // Lógica robusta para encontrar as colunas, independente de pequenas variações no nome
    const col = {
      numeroDoPedido: headers.findIndex(h => h.toUpperCase().includes('NUMERO DO PEDIDO')),
      empresa: headers.findIndex(h => h.toUpperCase() === 'EMPRESA'),
      data: headers.findIndex(h => h.toUpperCase() === 'DATA'),
      fornecedor: headers.findIndex(h => h.toUpperCase() === 'FORNECEDOR'),
      placaVeiculo: headers.findIndex(h => h.toUpperCase().includes('PLACA')),
      veiculo: headers.findIndex(h => h.toUpperCase().includes('NOME VEICULO')),
      observacoes: headers.findIndex(h => h.toUpperCase().includes('OBSERVACOES')),
      totalGeral: headers.findIndex(h => h.toUpperCase().includes('TOTAL GERAL') || h.toUpperCase() === 'VALOR'),
      status: headers.findIndex(h => h.toUpperCase() === 'STATUS'),
      itens: headers.findIndex(h => h.toUpperCase() .includes('ITENS')),
      estado: headers.findIndex(h => h.toUpperCase().includes('ESTADO FORNECEDOR')),
      dataCriacao: headers.findIndex(h => h.toUpperCase() === 'DATA CRIACAO'),
      aliquota: headers.findIndex(h => h.toUpperCase().includes('ALIQUOTA IMPOSTO')),
      icmsSt: headers.findIndex(h => h.toUpperCase().includes('ICMS ST TOTAL')),
      usuarioCriador: headers.findIndex(h => h.toUpperCase().includes('USUARIO CRIADOR'))
    };

    // Validar colunas essenciais (falhar cedo)
    if (col.numeroDoPedido === -1 || col.empresa === -1) {
      Logger.log("Colunas essenciais não encontradas: numeroDoPedido ou empresa.");
      return { status: 'error', message: 'Colunas essenciais não encontradas na planilha.' };
    }

    // Pré-processa parâmetros
    const empresaFiltro = String(params.empresaId).trim();
    const mainSearch = params.mainSearch ? String(params.mainSearch).toLowerCase().trim() : null;
    const plateFilter = params.plateSearch ? normalizePlate(params.plateSearch) : null;
    const usuarioCriadorFiltro = params.usuarioCriador ? String(params.usuarioCriador).toLowerCase().trim() : null;
    const bypassStatusFilter = params.bypassStatusFilter || (String(params.perfil || '').toLowerCase() === 'admin');

    const dateStartTime = params.dateStart ? (new Date(params.dateStart + 'T00:00:00')).getTime() : null;
    const dateEndTime = params.dateEnd ? (new Date(params.dateEnd + 'T23:59:59')).getTime() : null;

    const statusExcluidos = new Set(['RASCUNHO', 'AGUARDANDO APROVACAO']);

    // Itera linhas com for (mais performático)
    const results = [];
    for (let r = 1; r < all.length; r++) {
      const row = all[r];

      // 1) Empresa (filtro rápido)
      const empresaPlanilha = String(row[col.empresa] || '').trim();
      if (empresaPlanilha !== empresaFiltro) continue;

      // 2) Status (rápido reject se aplicável)
      if (!bypassStatusFilter && col.status !== -1) {
        const statusVal = String(row[col.status] || '').trim().toUpperCase();
        if (!statusVal || statusExcluidos.has(statusVal)) continue;
      }

      // 3) Main search (nº pedido ou fornecedor)
      if (mainSearch) {
        const numLower = String(row[col.numeroDoPedido] || '').toLowerCase();
        const fornLower = String(row[col.fornecedor] || '').toLowerCase();
        if (!(numLower.includes(mainSearch) || fornLower.includes(mainSearch))) continue;
      }

      // 4) Data range
      if (dateStartTime !== null && col.data !== -1) {
        const cell = row[col.data];
        let time = NaN;
        if (cell instanceof Date) time = cell.getTime();
        else {
          const parsed = Date.parse(String(cell || ''));
          time = isNaN(parsed) ? NaN : parsed;
        }
        if (isNaN(time) || time < dateStartTime || time > dateEndTime) continue;
      }

      // 5) Placa
      if (plateFilter && col.placaVeiculo !== -1) {
        const placaPlan = normalizePlate(row[col.placaVeiculo] || '');
        if (!placaPlan.includes(plateFilter)) continue;
      }

      // 6) Usuário criador
      if (usuarioCriadorFiltro && col.usuarioCriador !== -1) {
        const criadorPlan = String(row[col.usuarioCriador] || '').toLowerCase().trim();
        if (criadorPlan !== usuarioCriadorFiltro) continue;
      }

      // --- passou todos os filtros: montar objeto resultante ---
      const dataCell = row[col.data];
      const dataStr = (dataCell instanceof Date)
        ? Utilities.formatDate(dataCell, "GMT-03:00", "yyyy-MM-dd'T'HH:mm:ss'Z'")
        : String(dataCell || '');
        // Pega o nome do fornecedor da linha atual da planilha "Pedidos"
        const nomeFornecedor = row[col.fornecedor];
        
        // Usa o mapa para encontrar o objeto completo do fornecedor
        const infoFornecedor = mapaFornecedores[String(nomeFornecedor).toUpperCase().trim()] || {};

        const pedido ={
        numeroDoPedido: row[col.numeroDoPedido],
        empresaId: row[col.empresa],
        data: dataStr,
        fornecedor: row[col.fornecedor],
        fornecedorId: infoFornecedor.id || null,
        totalGeral: row[col.totalGeral],
        status: row[col.status],
        placa: row[col.placaVeiculo],
        veiculo: row[col.veiculo],
        observacoes: row[col.observacoes],
        //itens: row[colunas.itens]
        estado: row[col.estado],
        dataCriacao: row[col.dataCriacao] instanceof Date ? Utilities.formatDate(row[col.dataCriacao], "GMT-03:00", "yyyy-MM-dd'T'HH:mm:ss'Z'") : row[col.dataCriacao],
        aliquota: row[col.aliquota],
        icmsSt: row[col.icmsSt],
        usuarioCriador: row[col.usuarioCriador]
      };

      const itensJSON = row[col.itens];
      if (col.itens !== -1 && itensJSON && String(itensJSON).trim() !== '') {
        try {
          pedido.itens = JSON.parse(itensJSON);
        } catch (e) {
          // MELHORIA: Usar Logger.log para erros do servidor
          Logger.log(`Erro ao parsear JSON de itens do pedido ${pedido.numeroDoPedido}: ` + e);
          pedido.itens = [];
          pedido.erroItens = "Formato inválido";
        }
      } else { 
        pedido.itens = [];
      }
      results.push(pedido);
    }
     return { status: 'success', data: results };

  } catch (e) {
    Logger.log("Erro na função buscarPedidos: " + e + "\nStack: " + e.stack);
    return { status: 'error', message: e.toString() };
  }


 function normalizePlate(s) {
    return String(s || '').replace(/[^A-Za-z0-9]/g, '').toUpperCase().trim();
  }
}

function _TESTE_inspecionarSaidaDaBuscaV2() {
  try {
    Logger.log("--- 🔬 INICIANDO TESTE DO MOTOR DE BUSCA buscarPedidosv2 ---");

    // --- PASSO MAIS IMPORTANTE: ADAPTE OS DADOS ABAIXO ---
    // Coloque aqui parâmetros de um pedido que você sabe que existe,
    // para garantir que a busca encontre um resultado.
    const paramsDeTeste = {
      mainSearch: "000108",   // << SUBSTITUA por um NÚMERO DE PEDIDO ou NOME DE FORNECEDOR real
      empresaId: "002"         // << SUBSTITUA pelo ID DE EMPRESA correto para o pedido acima
      // Você pode adicionar outros filtros aqui se quiser testar, como:
      // dateStart: "2025-08-01",
      // dateEnd: "2025-08-31"
    };

    Logger.log(`Executando buscarPedidosv2 com os parâmetros: ${JSON.stringify(paramsDeTeste)}`);

    // --- Executa a função real que queremos testar ---
    const resultado = buscarPedidosv2(paramsDeTeste);

    // --- Imprime o resultado completo no log de forma legível ---
    Logger.log("--- ⬇️ RESPOSTA COMPLETA RETORNADA PELA FUNÇÃO ⬇️ ---");
    Logger.log(JSON.stringify(resultado, null, 2));

    // --- Análise do resultado ---
    if (resultado && resultado.status === 'success' && resultado.data.length > 0) {
      Logger.log("--- ✅ ANÁLISE DO TESTE: SUCESSO! ---");
      Logger.log(`A busca encontrou ${resultado.data.length} pedido(s).`);
      Logger.log("==> POR FAVOR, VERIFIQUE NO LOG ACIMA se a propriedade 'fornecedorId' está presente e com o ID numérico correto dentro do objeto do pedido.");
    } else if (resultado && resultado.status === 'success') {
      Logger.log("--- ⚠️ ANÁLISE DO TESTE: AVISO ---");
      Logger.log("A função executou sem erros, mas não encontrou nenhum pedido com os critérios informados.");
    } else {
      Logger.log(`--- 🔥 ANÁLISE DO TESTE: FALHA ---`);
      Logger.log("A função retornou um status de erro. Mensagem: " + (resultado ? resultado.message : "N/A"));
    }

  } catch (e) {
    Logger.log(`🔥🔥 ERRO CRÍTICO na função de teste: ${e.message}`);
  }
}

    /**
     * Busca um único pedido pelo seu número e pelo ID da empresa para edição.
     * @param {string} numeroDoPedido - O número do pedido a ser encontrado.
     * @param {string} idEmpresa - O ID da empresa à qual o pedido pertence.
     * @returns {object|null} O objeto do pedido encontrado ou null se não encontrar.
     */
    function getPedidoParaEditar(numeroDoPedido, idEmpresa) {
    Logger.log(`[getPedidoParaEditar] Iniciando busca. Pedido: "${numeroDoPedido}", Empresa: "${idEmpresa}"`);
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
        if (!sheet) {
            Logger.log("[getPedidoParaEditar] ERRO: Planilha 'Pedidos' não encontrada.");
            throw new Error("Planilha 'Pedidos' não encontrada.");
        }

        const data = sheet.getDataRange().getValues();
        Logger.log(`[getPedidoParaEditar] Planilha "Pedidos" lida. Total de ${data.length - 1} registros de dados.`);
        const originalHeaders = data[0];

        // Encontra os índices das colunas usando os cabeçalhos originais
        const indexNumero = originalHeaders.findIndex(h => String(h).toUpperCase().trim() === 'NÚMERO DO PEDIDO');
        const indexEmpresa = originalHeaders.findIndex(h => ['ID DA EMPRESA', 'ID EMPRESA', 'EMPRESA'].includes(String(h).toUpperCase().trim()));
        
        if (indexNumero === -1 || indexEmpresa === -1) {
            Logger.log(`[getPedidoParaEditar] ERRO: Colunas não encontradas. Índice 'Número do Pedido': ${indexNumero}, Índice 'Empresa': ${indexEmpresa}`);
            throw new Error("Colunas 'Número do Pedido' ou 'ID da Empresa' não encontradas.");
        }

        // Procura pela linha que corresponde ao número do pedido E ao ID da empresa
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const numPedidoNaLinha = String(row[indexNumero]).trim();
            const idEmpresaNaLinha = String(row[indexEmpresa]).trim();

            if (numPedidoNaLinha == numeroDoPedido && idEmpresaNaLinha == idEmpresa) {
                Logger.log(`[getPedidoParaEditar] Pedido encontrado na linha ${i + 1}. Montando o objeto de retorno...`);
                // Encontrou o pedido, agora monta o objeto
                const pedido = {};
                originalHeaders.forEach((header, index) => {
                    const headerTrimmed = String(header).trim();
                    let value = row[index];
                    Logger.log(`  -> Processando coluna "${headerTrimmed}". Valor bruto: "${value}" (Tipo: ${typeof value})`);

                    // Garante que todos os campos de data sejam convertidos para uma string padronizada
                    if (value instanceof Date) {
                        const headerUpper = headerTrimmed.toUpperCase();
                        
                        // Se for 'Data Criacao' ou 'Data Ultima Edicao', formata COM a hora.
                        if (headerUpper === 'DATA CRIACAO' || headerUpper === 'DATA ULTIMA EDICAO') {
                            value = formatarDataParaISO(value); // Retorna 'YYYY-MM-DD HH:MM:SS'
                            Logger.log(`     - Data com hora. Valor formatado: "${value}"`);
                        } 
                        // Para a coluna 'Data' principal, formata SEM a hora.
                        else {
                            value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                            Logger.log(`     - Data sem hora. Valor formatado: "${value}"`);
                        }
                    }
                    
                    pedido[toCamelCase(headerTrimmed)] = value;
                });
                
                // Faz o parse dos itens
                if (typeof pedido.itens === 'string' && pedido.itens) {
                    try {
                        pedido.itens = JSON.parse(pedido.itens);
                    } catch(e) {
                        Logger.log(`[getPedidoParaEditar] ERRO ao fazer parse dos itens para o pedido ${numeroDoPedido}: ${e.message}`);
                        pedido.itens = [];
                    }
                }
                
                Logger.log(`[getPedidoParaEditar] Objeto final do pedido montado e pronto para ser retornado.`);
                return pedido; // Retorna o objeto do pedido encontrado
            }
        }

        Logger.log(`[getPedidoParaEditar] Finalizou o loop. Pedido "${numeroDoPedido}" não foi encontrado para a empresa "${idEmpresa}".`);
        return null; // Retorna null se não encontrar o pedido

    } catch (e) {
        // O log de erro agora inclui o stack trace para mais detalhes
        Logger.log(`[getPedidoParaEditar] ERRO FATAL: ${e.message}. Stack: ${e.stack}`);
        return null;
    }
}

// ===============================================
    // FUNÇÕES PARA VEICULOS, PLACAS E FORNECEDORES
    // ===============================================
    /**
     * Adiciona um novo nome de veículo à planilha "Veiculos".
     * @param {string} nomeVeiculo - O nome do novo veículo a ser adicionado.
     * @returns {object} Um objeto com o status da operação.
     */
    /**
     * Retorna uma lista de todos os nomes de veículos cadastrados.
     */
    function getVeiculosList() {
      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
        if (!sheet) {
          Logger.log("Planilha 'Config' não encontrada.");
          return []; 
        }

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        // Lê apenas a primeira coluna (A)
        const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
        
        // Mapeia para um array de strings e remove espaços em branco
        const veiculos = data.map(row => String(row[0]).trim()).filter(nome => nome !== "");

        // Ordena alfabeticamente
        veiculos.sort((a, b) => a.localeCompare(b));
        
        return veiculos;
      } catch (e) {
        Logger.log("Erro em getVeiculosList: " + e.message);
        return [];
      }
    }

    /**
     * Adiciona um novo nome de veículo à planilha "Veiculos".
     * @param {string} nomeVeiculo - O nome do novo veículo a ser adicionado.
     * @returns {object} Um objeto com o status da operação.
     */
    function adicionarNovoVeiculo(nomeVeiculo) {
      if (!nomeVeiculo || typeof nomeVeiculo !== 'string' || nomeVeiculo.trim() === '') {
        return { status: 'error', message: 'O nome do veículo não pode estar vazio.' };
      }

      const nomeLimpo = nomeVeiculo.trim().toUpperCase(); // Padroniza para maiúsculas

      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
        
        // Pega todos os valores da coluna C para verificar se o veículo já existe
        const rangeVeiculos = sheet.getRange("C2:C" + sheet.getLastRow());
        const veiculosExistentes = rangeVeiculos.getValues().map(row => String(row[0]).trim().toUpperCase());

        // --- LÓGICA DE VALIDAÇÃO INTELIGENTE ---
        const semelhançaMinima = 2; // Aceita até 2 letras diferentes. Você pode ajustar este valor.

        for (const existente of veiculosExistentes) {
          const distancia = levenshteinDistance(nomeLimpo, existente);

          if (distancia === 0) {
            return { status: 'exists', message: 'Este veículo já está cadastrado.' };
          }
          
          if (distancia <= semelhançaMinima) {
            return { status: 'similar', message: `Erro: O nome '${nomeVeiculo}' é muito parecido com '${existente}', que já está cadastrado.` };
          }
        }
        // --- FIM DA VALIDAÇÃO ---   
        
        // Encontra a próxima linha vazia na coluna C e adiciona o novo veículo lá
        const proximaLinhaVazia = rangeVeiculos.getValues().filter(String).length + 2;
        sheet.getRange(proximaLinhaVazia, 3).setValue(nomeLimpo);
        
        return { status: 'ok', message: 'Veículo adicionado com sucesso!', novoVeiculo: nomeLimpo };
      } catch (e) {
        Logger.log("Erro em adicionarNovoVeiculo: " + e.message);
        return { status: 'error', message: 'Ocorreu um erro ao salvar o novo veículo.' };
      }
    }

    /**
     * Calcula a Distância de Levenshtein entre duas strings.
     * Retorna o número de edições necessárias para transformar uma string na outra.
     */
    function levenshteinDistance(a, b) {
      if (a.length === 0) return b.length;
      if (b.length === 0) return a.length;

      const matrix = Array(b.length + 1).fill(null).map(() => Array(a.length + 1).fill(null));

      for (let i = 0; i <= a.length; i++) {
        matrix[0][i] = i;
      }
      for (let j = 0; j <= b.length; j++) {
        matrix[j][0] = j;
      }

      for (let j = 1; j <= b.length; j++) {
        for (let i = 1; i <= a.length; i++) {
          const cost = a[i - 1] === b[j - 1] ? 0 : 1;
          matrix[j][i] = Math.min(
            matrix[j][i - 1] + 1,      // Deletion
            matrix[j - 1][i] + 1,      // Insertion
            matrix[j - 1][i - 1] + cost // Substitution
          );
        }
      }

      return matrix[b.length][a.length];
    }

    // ===============================================
    // FUNÇÕES PARA RASCUNHO
    // ===============================================
    /**
     * Salva um rascunho na planilha
     * @param {Object} dadosRascunho - Dados do rascunho a ser salvo
     * @returns {Object} - Resposta com status e ID do rascunho
     */
    function salvarRascunho(dadosRascunho) {
      try {
        console.log('📝 Salvando rascunho:', dadosRascunho);
        const idDaPlanilha = '1M0GTX9WmnggiNnDynU0kC457yoy0iRHcRJ39d_B109o'; // Coloque o ID aqui
        const colunas = mapearCabecalhoPedidos(idDaPlanilha);
        // Se o mapa não for criado, pare a execução
        if (!colunas) {
            return { status: 'error', message: 'Não foi possível ler a estrutura da planilha.' };
        }
        // Obter a planilha
        const planilha = SpreadsheetApp.openById('1M0GTX9WmnggiNnDynU0kC457yoy0iRHcRJ39d_B109o');
        const aba = planilha.getSheetByName('Pedidos') || planilha.insertSheet('Pedidos');
        const agora = new Date();

        //const indiceDaAliquota = colunas.aliquotaImposto; // Vai retornar 19
        //const indiceDataCriacao = colunas.dataCriacao; // Vai retornar 15
        // Validações básicas
        if (!dadosRascunho.fornecedor || !dadosRascunho.fornecedor.trim()) {
          return {
            status: 'error',
            message: 'Fornecedor é obrigatório para salvar o rascunho.'
          };
        }
        
        if (!dadosRascunho.itens || !Array.isArray(dadosRascunho.itens) || dadosRascunho.itens.length === 0) {
          return {
            status: 'error',
            message: 'Pelo menos um item é obrigatório para salvar o rascunho.'
          };
        }
        
        // Validar se pelo menos um item tem descrição
        const itemValido = dadosRascunho.itens.some(item => item.descricao && item.descricao.trim());
        if (!itemValido) {
          return {
            status: 'error',
            message: 'Pelo menos um item deve ter uma descrição.'
          };
        }
                
        // Gerar ID único para o rascunho
        
        const ano = agora.getFullYear();
        const mes = String(agora.getMonth() + 1).padStart(2, '0');
        const dia = String(agora.getDate()).padStart(2, '0');
        const timestamp = agora.getTime();
        const rascunhoId = `RASC-${ano}${mes}${dia}-${timestamp}`;
        
        // Obter dados do fornecedor da aba Fornecedores (se existir)
        const fornecedoresSheet = planilha.getSheetByName('Fornecedores');
        let fornecedorCnpj = '';
        let fornecedorEndereco = '';
        let condicaoPagamentoFornecedor = '';
        let formaPagamentoFornecedor = '';
        let estadoFornecedor = '';
        let cidadeFornecedor = '';

        if (fornecedoresSheet) {
          const fornecedoresData = fornecedoresSheet.getRange(2, 1, fornecedoresSheet.getLastRow() - 1, fornecedoresSheet.getLastColumn()).getValues();
          const foundFornecedor = fornecedoresData.find(row => String(row[0]) === dadosRascunho.fornecedor); 
          if (foundFornecedor) {
            razaoSocialFornecedor = String(foundFornecedor[1] || '');
            fornecedorCnpj = String(foundFornecedor[3] || '');
            fornecedorEndereco = String(foundFornecedor[4] || '');
            condicaoPagamentoFornecedor = String(foundFornecedor[5] || '');
            formaPagamentoFornecedor = String(foundFornecedor[6] || '');
            estadoFornecedor = String(foundFornecedor[10] || ''); // Coluna 11 (índice 10) = Estado
            cidadeFornecedor = String(foundFornecedor[11] || '');
          }
        }

        // Preparar dados para salvar (mesma estrutura do salvarPedido)
        const dadosParaSalvar = {
          'Número do Pedido': "'" + rascunhoId, // Usando ID do rascunho como número
          'Empresa': "'" + (dadosRascunho.empresa || Session.getActiveUser().getEmail()),
          'Data': dadosRascunho.data ? formatarDataParaISO(dadosRascunho.data) : formatarDataParaISO(agora),
          'Fornecedor': razaoSocialFornecedor || dadosRascunho.fornecedor.trim(),
          'CNPJ Fornecedor': fornecedorCnpj,
          'Endereço Fornecedor': fornecedorEndereco,
          'Estado Fornecedor': estadoFornecedor,
          'Condição Pagamento Fornecedor': condicaoPagamentoFornecedor,
          'Forma Pagamento Fornecedor': formaPagamentoFornecedor,
          'Placa Veiculo': dadosRascunho.placaVeiculo || '',
          'Nome Veiculo': dadosRascunho.nomeVeiculo || '',
          'Observacoes': dadosRascunho.observacoes || '',
          'Total Geral': dadosRascunho.totalGeral || 0,
          'Status': 'RASCUNHO', // Diferença principal: status RASCUNHO em vez de "Em Aberto"
          'Itens': JSON.stringify(dadosRascunho.itens),
          'ICMS ST Total': dadosRascunho.valorIcms || 0,
          'Data Ultima Edicao': formatarDataParaISO(agora), // Sempre usar data/hora atual padronizada
          'Aliquota Imposto': dadosRascunho.aliquotaImposto || 0,
          'Usuario Criador': dadosRascunho.usuarioCriador
          
        };
        
        
        // Verificar se é uma atualização de rascunho existente
          if (dadosRascunho.rascunhoId) {
              const linhaExistente = encontrarLinhaRascunho(aba, dadosRascunho.rascunhoId, colunas);

              if (linhaExistente > 0) {
                  // --- MODO ATUALIZAÇÃO "À PROVA DE BALAS" ---

                  // 1. Preserva a Data de Criação lendo o valor antigo (seu código, que está correto)
                  const indiceDataCriacao = colunas.dataCriacao;
                  if (indiceDataCriacao !== -1) {
                      const dataCriacaoAntiga = aba.getRange(linhaExistente, indiceDataCriacao + 1).getValue();
                      if (dataCriacaoAntiga) {
                          dadosLimpos.dataCriacao = dataCriacaoAntiga; 
                      }
                  }

                  // Garante que o ID do Pedido está nos dados a serem salvos
                  dadosLimpos.numeroDoPedido = "'" + dadosRascunho.rascunhoId;

                  // 2. Este loop SUBSTITUI a 'salvarDadosNaPlanilha'
                  // Ele atualiza cada célula individualmente, de forma segura.
                  console.log(`Atualizando rascunho na linha ${linhaExistente}.`);
                  for (const chave in dadosLimpos) {
                      const indiceColuna = colunas[chave]; // Pega o índice (ex: aliquotaImposto -> 19)
                      if (indiceColuna !== -1 && indiceColuna !== undefined) {
                          // Escreve o valor na célula exata (linha 12, coluna 20, por exemplo)
                          aba.getRange(linhaExistente, indiceColuna + 1).setValue(dadosLimpos[chave]);
                      }
                  }

                  // 3. Força a sincronização DEPOIS de dar os comandos .setValue()
                  SpreadsheetApp.flush();

                  Logger.log(`✅ Alterações salvas e sincronizadas na planilha para a linha ${linhaExistente}.`);
                  return { status: 'success', message: 'Rascunho atualizado com sucesso!', rascunhoId: dadosRascunho.rascunhoId };
              } else {
                  // Se não encontrou a linha, retorna um erro claro
                  return { status: 'error', message: `Rascunho ${dadosRascunho.rascunhoId} não encontrado para atualizar.` };
              }
          }
          
        /// ---------- É UM NOVO RASCUNHO ----------
        //const timestamp = agora.getTime();
        //const rascunhoId = `RASC-${agora.getFullYear()}${(agora.getMonth() + 1).toString().padStart(2, '0')}${agora.getDate().toString().padStart(2, '0')}-${timestamp}`;
        dadosParaSalvar['Número do Pedido'] = "'" + rascunhoId;
        
        // Adicionamos a 'Data Criacao' e o 'Usuario Criador' apenas para rascunhos novos
        dadosParaSalvar['Data Criacao'] = formatarDataParaISO(agora);
        dadosParaSalvar['Usuario Criador'] = dadosRascunho.usuarioCriador || Session.getActiveUser().getEmail();

        salvarDadosNaPlanilha(aba, dadosParaSalvar);
        return { status: 'success', message: 'Rascunho salvo!', rascunhoId: rascunhoId };
        
    } catch (error) {
        console.error('❌ Erro ao salvar rascunho:', error);
        return { status: 'error', message: 'Erro interno ao salvar rascunho: ' + error.message };
    }
}

function encontrarLinhaRascunho(aba, rascunhoId, colunas) {
  try {
    const indiceId = colunas.numeroDoPedido;
    if (indiceId === -1) {
      Logger.log("Erro em encontrarLinhaRascunho: a coluna 'Número do Pedido' não foi encontrada.");
      return 0;
    }

    const valoresId = aba.getRange(2, indiceId + 1, aba.getLastRow() - 1, 1).getValues();
    Logger.log(`Procurando pelo rascunhoId: "${rascunhoId}"`);

    for (let i = 0; i < valoresId.length; i++) {
      // ✅ A CORREÇÃO FINAL ESTÁ AQUI: Adicionamos .trim() para limpar a célula
      const valorCelula = String(valoresId[i][0]).trim();

      // A lógica de comparação robusta que já tínhamos
      if (valorCelula === rascunhoId || valorCelula === "'" + rascunhoId) {
        const numeroLinha = i + 2;
        Logger.log(`✅ Rascunho encontrado! ID "${rascunhoId}" corresponde ao valor da célula "${valorCelula}" na linha ${numeroLinha}.`);
        return numeroLinha;
      }
    }

    Logger.log(`❌ Rascunho com ID "${rascunhoId}" não foi encontrado na planilha.`);
    return 0;

  } catch (e) {
    Logger.log('Erro crítico na função encontrarLinhaRascunho: ' + e.stack);
    return 0;
  }
}

    /**
     * Busca todos os rascunhos de uma empresa
     * @param {string} empresaId - ID da empresa
     * @returns {Object} - Lista de rascunhos
     */
    function buscarRascunhosv2(empresaId) {
      console.log('🔍 [BACKEND] === INÍCIO buscarRascunhos ===');
      console.log('🔍 [BACKEND] Parâmetro empresaId:', empresaId);
      console.log('🔍 [BACKEND] Tipo do empresaId:', typeof empresaId);
      
      try {
        // ID da planilha definido localmente
        var planilhaId = '1xVLFSqL5SVT6cmZ_9foOkKxJIHqHeCGHLcyryBJ44g0';
        
        // Validação básica
        if (!empresaId) {
          console.error('❌ [BACKEND] empresaId é obrigatório');
          var erro = {
            status: 'error',
            message: 'ID da empresa é obrigatório',
            rascunhos: []
          };
          console.log('📤 [BACKEND] Retornando erro de validação:', erro);
          return erro;
        }
        
        console.log('✅ [BACKEND] Validação OK, tentando acessar planilha...');
        console.log('🔍 [BACKEND] planilhaId:', planilhaId);
        
        var planilha = SpreadsheetApp.openById(planilhaId);
        console.log('✅ [BACKEND] Planilha acessada com sucesso');
        
        var aba = planilha.getSheetByName('Pedidos');
        
        if (!aba) {
          console.log('📋 [BACKEND] Aba Pedidos não encontrada');
          var sucesso = {
            status: 'success',
            rascunhos: [],
            message: 'Aba Pedidos não encontrada'
          };
          console.log('📤 [BACKEND] Retornando lista vazia:', sucesso);
          return sucesso;
        }
        
        console.log('✅ [BACKEND] Aba Pedidos encontrada');
        
        var dados = aba.getDataRange().getValues();
        console.log('📊 [BACKEND] Dados obtidos - Total de linhas:', dados.length);
        
        if (dados.length < 2) {
          console.log('📋 [BACKEND] Planilha vazia ou só cabeçalho');
          var vazio = {
            status: 'success',
            rascunhos: [],
            message: 'Planilha vazia'
          };
          console.log('📤 [BACKEND] Retornando planilha vazia:', vazio);
          return vazio;
        }
        
        var cabecalhos = dados[0];
        var rascunhos = [];
        
        console.log('📊 [BACKEND] Cabeçalhos:', cabecalhos);
        
        // Encontrar índices das colunas (usando os nomes reais da planilha)
        var indices = {
          numeroPedido: cabecalhos.indexOf('Número do Pedido'),
          empresa: cabecalhos.indexOf('Empresa'),
          status: cabecalhos.indexOf('Status'),
          data: cabecalhos.indexOf('Data'),
          fornecedor: cabecalhos.indexOf('Fornecedor'),
          nomeVeiculo: cabecalhos.indexOf('Nome Veiculo'),
          placaVeiculo: cabecalhos.indexOf('Placa Veiculo'),
          observacoes: cabecalhos.indexOf('Observacoes'),
          itens: cabecalhos.indexOf('Itens'),
          totalGeral: cabecalhos.indexOf('Total Geral'),
          produtoFornecedor: cabecalhos.indexOf('Produto Fornecedor'),
          icmsSTTotal: cabecalhos.indexOf('ICMS ST Total'),
          //dataUltimaCriacao: cabecalhos.indexOf('Data Ultima Edicao'),
          //usuarioCriador: cabecalhos.indexOf('Usuario Criador')
        };
        
        console.log('📊 [BACKEND] Índices encontrados:', indices);
        
        // Verificar colunas críticas
        if (indices.status === -1) {
          console.error('❌ [BACKEND] Coluna Status não encontrada');
          var erro = {
            status: 'error',
            message: 'Coluna Status não encontrada na planilha',
            rascunhos: []
          };
          console.log('📤 [BACKEND] Retornando erro de estrutura:', erro);
          return erro;
        }
        
        if (indices.empresa === -1) {
          console.error('❌ [BACKEND] Coluna Empresa não encontrada');
          var erro = {
            status: 'error',
            message: 'Coluna Empresa não encontrada na planilha',
            rascunhos: []
          };
          console.log('📤 [BACKEND] Retornando erro de estrutura:', erro);
          return erro;
        }
        
        console.log('✅ [BACKEND] Estrutura da planilha validada');
        
        // Processar dados
        var rascunhosEncontrados = 0;
        var empresaIdStr = String(empresaId).trim();
        
        console.log('🔍 [BACKEND] Processando linhas para empresa:', empresaIdStr);
        
        for (var i = 1; i < dados.length; i++) {
          var linha = dados[i];
          var statusLinha = linha[indices.status];
          var empresaLinha = linha[indices.empresa];
          
          // Debug das primeiras 3 linhas
          if (i <= 3) {
            console.log('📊 [BACKEND] Linha ' + i + ': Status="' + statusLinha + '", Empresa="' + empresaLinha + '"');
          }
          
          // Verificar se é rascunho da empresa
          if (statusLinha === 'RASCUNHO' && empresaLinha) {
            // Remover apóstrofo do campo empresa para comparação
            var empresaNaPlanilha = String(empresaLinha).replace(/'/g, '').trim();
            
            if (i <= 3) {
              console.log('🔍 [BACKEND] Comparando linha ' + i + ': "' + empresaNaPlanilha + '" === "' + empresaIdStr + '"');
            }
            
            if (empresaNaPlanilha === empresaIdStr) {
              rascunhosEncontrados++;
              console.log('✅ [BACKEND] Rascunho ' + rascunhosEncontrados + ' encontrado na linha ' + (i + 1));
              
              var itensArray = [];
              try {
                if (linha[indices.itens]) {
                  itensArray = JSON.parse(linha[indices.itens]);
                }
              } catch (e) {
                console.warn('⚠️ [BACKEND] Erro ao parsear itens:', linha[indices.numeroPedido]);
                itensArray = [];
              }
              
              var rascunho = {
                id: linha[indices.numeroPedido] ? String(linha[indices.numeroPedido]).replace(/'/g, '') : '',
                data: linha[indices.data] ? String(linha[indices.data]) : '',
                fornecedor: linha[indices.fornecedor] || '',
                nomeVeiculo: linha[indices.nomeVeiculo] || '',
                placaVeiculo: linha[indices.placaVeiculo] || '',
                observacoes: linha[indices.observacoes] || '',
                itens: itensArray,
                totalGeral: Number(linha[indices.totalGeral]) || 0,
                produtoFornecedor: linha[indices.produtoFornecedor] || '',
                icmsStTotal: linha[indices.icmsSTTotal] || '',
                //dataUltimaCriacao: linha[indices.dataUltimaCriacao] || '',
                //usuarioCriador: linha[indices.usuarioCriador] || ''
              };
              
              rascunhos.push(rascunho);
            }
          }
        }
        
        console.log(`✅ [BACKEND] Processamento concluído - ${rascunhos.length} rascunhos encontrados`);
        
        // Ordenar por data (mais recente primeiro)
        try {
          rascunhos.sort((a, b) => new Date(b.data) - new Date(a.data));
          console.log('✅ [BACKEND] Rascunhos ordenados por data');
        } catch (sortError) {
          console.warn('⚠️ [BACKEND] Erro ao ordenar:', sortError);
        }
        
        const resultado = {
          status: 'success',
          rascunhos: rascunhos,
          message: `${rascunhos.length} rascunho(s) encontrado(s)`
        };
        
        console.log('📤 [BACKEND] Retornando resultado final:', resultado);
        return resultado;
        
      } catch (error) {
        console.error('❌ [BACKEND] Erro na função buscarRascunhos:', error);
        console.error('❌ [BACKEND] Stack trace:', error.stack);
        
        const erro = {
          status: 'error',
          message: 'Erro interno: ' + error.message,
          rascunhos: []
        };
        
        console.log('📤 [BACKEND] Retornando erro:', erro);
        return erro;
      } finally {
        console.log('🔍 [BACKEND] === FIM buscarRascunhos ===');
      }
    }

    /**
     * Busca um rascunho específico por ID
     * @param {string} rascunhoId - ID do rascunho
     * @returns {Object} - Dados do rascunho
     */
    function buscarRascunhoPorId(rascunhoId) {
      try {
        console.log('🔍 [BUSCAR ID] Buscando rascunho por ID:', rascunhoId);
        
        // ID da planilha definido localmente
        var planilhaId = '1xVLFSqL5SVT6cmZ_9foOkKxJIHqHeCGHLcyryBJ44g0';
        var planilha = SpreadsheetApp.openById(planilhaId);
        var aba = planilha.getSheetByName('Pedidos');
        
        if (!aba) {
          return {
            status: 'error',
            message: 'Planilha de pedidos não encontrada.'
          };
        }
        
        var dados = aba.getDataRange().getValues();
        var cabecalhos = dados[0];
        
        // Buscar possíveis variações do nome da coluna de data última edição
        var possiveisNomes = ['Data Ultima Edicao', 'Data Última Edição', 'Ultima Edicao', 'Última Edição', 'Data da Ultima Edicao'];
        var indiceDataUltimaEdicao = -1;
        
        for (var nomeColuna of possiveisNomes) {
          var indice = cabecalhos.indexOf(nomeColuna);
          if (indice !== -1) {
            indiceDataUltimaEdicao = indice;
            console.log('🔍 [BUSCAR ID] ✅ Coluna encontrada:', nomeColuna, 'no índice:', indice);
            break;
          }
        }
        
        // Encontrar índices das colunas (usando os nomes reais da planilha)
        var indices = {
          numeroPedido: cabecalhos.indexOf('Número do Pedido'),
          status: cabecalhos.indexOf('Status'),
          data: cabecalhos.indexOf('Data'),
          empresa: cabecalhos.indexOf('Empresa'),
          fornecedor: cabecalhos.indexOf('Fornecedor'),
          nomeVeiculo: cabecalhos.indexOf('Nome Veiculo'),
          placaVeiculo: cabecalhos.indexOf('Placa Veiculo'),
          observacoes: cabecalhos.indexOf('Observacoes'),
          itens: cabecalhos.indexOf('Itens'),
          totalGeral: cabecalhos.indexOf('Total Geral'),
          dataUltimaEdicao: indiceDataUltimaEdicao,
          produtoFornecedor: cabecalhos.indexOf('produtoFornecedor'),
          icmsStTotal: cabecalhos.indexOf('ICMS ST Total')
        };
        
        console.log('🔍 [BUSCAR ID] Processando ' + (dados.length - 1) + ' linhas...');
        
        // Procurar o rascunho
        for (var i = 1; i < dados.length; i++) {
          var linha = dados[i];
          
          var numeroRascunho = linha[indices.numeroPedido] ? String(linha[indices.numeroPedido]).replace(/'/g, '') : '';
          if (numeroRascunho === rascunhoId && linha[indices.status] === 'RASCUNHO') {
            var itensArray = [];
            try {
              if (linha[indices.itens]) {
                itensArray = JSON.parse(linha[indices.itens]);
              }
            } catch (e) {
              console.warn('⚠️ [BUSCAR ID] Erro ao parsear itens do rascunho:', rascunhoId);
              itensArray = [];
            }
            
            var rascunho = {
              id: numeroRascunho,
              data: linha[indices.data] ? String(linha[indices.data]) : '',
              empresa: linha[indices.empresa] ? String(linha[indices.empresa]).replace(/'/g, '') : '',
              fornecedor: linha[indices.fornecedor] || '',
              nomeVeiculo: linha[indices.nomeVeiculo] || '',
              placaVeiculo: linha[indices.placaVeiculo] || '',
              observacoes: linha[indices.observacoes] || '',
              itens: itensArray,
              totalGeral: Number(linha[indices.totalGeral]) || 0,
              dataUltimaEdicao: (indices.dataUltimaEdicao !== -1 && linha[indices.dataUltimaEdicao]) ? String(linha[indices.dataUltimaEdicao]) : '',
              produtoFornecedor: linha[indices.produtoFornecedor] || '',
              icmsStTotal: linha[indices.icmsStTotal] || ''
            };
            
            console.log('✅ [BUSCAR ID] Rascunho encontrado:', rascunhoId);
            return {
              status: 'success',
              rascunho: rascunho
            };
          }
        }
        
        console.log('❌ [BUSCAR ID] Rascunho não encontrado:', rascunhoId);
        return {
          status: 'error',
          message: 'Rascunho não encontrado.'
        };
        
      } catch (error) {
        console.error('❌ [BUSCAR ID] Erro ao buscar rascunho por ID:', error);
        return {
          status: 'error',
          message: 'Erro ao buscar rascunho: ' + error.message
        };
      }
    }

    /**
     * Finaliza um rascunho como pedido oficial
     * @param {string} rascunhoId - ID do rascunho
     * @returns {Object} - Resultado da operação
     */
    function finalizarRascunho(rascunhoId) {
      try {
        console.log('✅ Finalizando rascunho:', rascunhoId);
        
        // Buscar dados do rascunho
        const resultadoBusca = buscarRascunhoPorId(rascunhoId);
        if (resultadoBusca.status !== 'success') {
          return resultadoBusca;
        }
        
        const dadosRascunho = resultadoBusca.rascunho;
        console.log('📋 Dados do rascunho encontrado:', dadosRascunho);
        
        // Validar dados para finalização
        const validacao = validarDadosParaPedido(dadosRascunho);
        if (!validacao.valido) {
          return {
            status: 'error',
            message: validacao.mensagem
          };
        }
        
        // Obter empresa do rascunho ou usar empresa do usuário logado
        let empresaCodigo = dadosRascunho.empresa;
        
        // Se não houver empresa no rascunho, tentar obter do usuário logado
        if (!empresaCodigo) {
          const usuarioLogado = obterUsuarioLogado();
          if (usuarioLogado && usuarioLogado.idEmpresa) {
            empresaCodigo = usuarioLogado.idEmpresa;
          } else {
            return {
              status: 'error',
              message: 'Não foi possível determinar a empresa para gerar o número do pedido.'
            };
          }
        }
        
        console.log('🏢 Empresa para geração do pedido:', empresaCodigo);
        
        // Gerar número do pedido sequencial por empresa
        const numeroPedido = getProximoNumeroPedido(empresaCodigo);
        console.log('📝 Número do pedido gerado:', numeroPedido);
        
        // Preparar dados do pedido
        const dadosPedido = {
          numero: numeroPedido,
          data: formatarDataParaISO(dadosRascunho.data || new Date()),
          fornecedor: dadosRascunho.fornecedor,
          nomeVeiculo: dadosRascunho.nomeVeiculo || '',
          placaVeiculo: dadosRascunho.placaVeiculo || '',
          observacoes: dadosRascunho.observacoes || '',
          itens: dadosRascunho.itens || [],
          totalGeral: dadosRascunho.totalGeral || 0,
          produtoFornecedor: dadosRascunho.produtoFornecedor || '',
          icmsStTtotal: dadosRascunho.icmsStTotal || '',
          empresa: empresaCodigo
        };
        
        console.log('📦 Dados do pedido preparados:', dadosPedido);
        
        // Salvar como pedido usando função existente
        const resultadoSalvamento = salvarPedido(dadosPedido);
        
        if (resultadoSalvamento.status === 'ok') {
          // Excluir o rascunho
          const resultadoExclusao = excluirRascunho(rascunhoId);
          
          console.log('✅ Rascunho finalizado como pedido:', numeroPedido);
          return {
            status: 'success',
            message: 'Rascunho finalizado com sucesso!',
            numeroPedido: numeroPedido
          };
        } else {
          return {
            status: 'error',
            message: 'Erro ao finalizar rascunho: ' + resultadoSalvamento.message
          };
        }
        
      } catch (error) {
        console.error('❌ Erro ao finalizar rascunho:', error);
        return {
          status: 'error',
          message: 'Erro interno ao finalizar rascunho: ' + error.message
        };
      }
    }

    /**
     * Exclui um rascunho
     * @param {string} rascunhoId - ID do rascunho
     * @returns {Object} - Resultado da operação
     */
    function excluirRascunho(rascunhoId) {
      try {
        console.log('🗑️ Excluindo rascunho:', rascunhoId);
        
        const planilha = SpreadsheetApp.openById(PLANILHA_ID);
        const aba = planilha.getSheetByName('Pedidos');
        
        if (!aba) {
          return {
            status: 'error',
            message: 'Planilha de pedidos não encontrada.'
          };
        }
        
        const linhaRascunho = encontrarLinhaRascunho(aba, rascunhoId);
        
        if (linhaRascunho > 0) {
          aba.deleteRow(linhaRascunho);
          
          console.log('✅ Rascunho excluído:', rascunhoId);
          return {
            status: 'success',
            message: 'Rascunho excluído com sucesso!'
          };
        } else {
          return {
            status: 'error',
            message: 'Rascunho não encontrado.'
          };
        }
        
      } catch (error) {
        console.error('❌ Erro ao excluir rascunho:', error);
        return {
          status: 'error',
          message: 'Erro ao excluir rascunho: ' + error.message
        };
      }
    }

/**
 * ATUALIZA um pedido existente na planilha 'Pedidos'.
 * Esta função deve ser adicionada a um dos seus arquivos .gs (ex: Pedidos.gs).
 *
 * @param {object} pedidoObject - O objeto do pedido com os dados atualizados. DEVE conter a propriedade 'numeroDoPedido'.
 * @returns {object} Um objeto de status com uma mensagem de sucesso ou erro.
 */
function editarPedido(pedidoObject) {
    Logger.log(`[editarPedido] 1. Iniciando atualização para o pedido: ${pedidoObject.numeroDoPedido}`);
    try {
        const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
        if (!sheet) {
            throw new Error("Planilha 'Pedidos' não encontrada.");
        }

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const indexNumeroPedido = headers.findIndex(h => String(h).toUpperCase().trim() === 'NÚMERO DO PEDIDO');

        if (indexNumeroPedido === -1) {
            throw new Error("Coluna 'Número do Pedido' não encontrada na planilha.");
        }

        const data = sheet.getDataRange().getValues();
        let rowIndexToUpdate = -1;
        let originalRowData = null;

        // Procura a linha que corresponde ao número do pedido
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][indexNumeroPedido]).trim() === String(pedidoObject.numeroDoPedido).trim()) {
                rowIndexToUpdate = i + 1; // +1 porque getRange é 1-indexed
                originalRowData = data[i]; // Armazena os dados originais da linha
                break;
            }
        }

        if (rowIndexToUpdate === -1) {
            return { status: 'error', message: 'Pedido para atualização não encontrado.' };
        }
        Logger.log(`[editarPedido] 2. Pedido encontrado na linha ${rowIndexToUpdate}.`);

        // Busca dados do fornecedor para garantir que estão atualizados
        const fornecedoresSheet = SpreadsheetApp.getActive().getSheetByName('Fornecedores');
        let fornecedorCnpj = '', fornecedorEndereco = '', condicaoPagamentoFornecedor = '', formaPagamentoFornecedor = '', estadoFornecedor = '';
        if (fornecedoresSheet) {
            const fornecedoresData = fornecedoresSheet.getRange(2, 1, fornecedoresSheet.getLastRow() - 1, fornecedoresSheet.getLastColumn()).getValues();
            const foundFornecedor = fornecedoresData.find(row => String(row[1]) === pedidoObject.fornecedor);
            if (foundFornecedor) {
                fornecedorCnpj = String(foundFornecedor[3] || '');
                fornecedorEndereco = String(foundFornecedor[4] || '');
                condicaoPagamentoFornecedor = String(foundFornecedor[5] || '');
                formaPagamentoFornecedor = String(foundFornecedor[6] || '');
                estadoFornecedor = String(foundFornecedor[10] || '');
                cidadeFornecedor = String(foundFornecedor[11] || '');
            }
        }

        // Mapeia os cabeçalhos para seus índices para facilitar a busca de dados originais
        const headerMap = {};
        headers.forEach((header, i) => headerMap[String(header).trim()] = i);

        // ================================================================
        // CORREÇÃO APLICADA AQUI
        // ================================================================
        const dataToSave = {
            // IDs são formatados como texto para preservar zeros à esquerda
            'Número do Pedido': "'" + pedidoObject.numeroDoPedido,
            'Empresa': "'" + pedidoObject.empresaId,
            
            // Para outros campos, usa o novo valor se ele existir, senão mantém o valor original
            'Data': pedidoObject.data || originalRowData[headerMap['Data']],
            'Fornecedor': pedidoObject.fornecedor,
            'CNPJ Fornecedor': fornecedorCnpj,
            'Endereço Fornecedor': fornecedorEndereco,
            'Estado Fornecedor': estadoFornecedor,
            'Condição Pagamento Fornecedor': condicaoPagamentoFornecedor,
            'Forma Pagamento Fornecedor': formaPagamentoFornecedor,
            'Placa Veiculo': pedidoObject.placaVeiculo,
            'Nome Veiculo': pedidoObject.nomeVeiculo,
            'Observacoes': pedidoObject.observacoes || originalRowData[headerMap['Observacoes']],
            'Total Geral': pedidoObject.totalGeral,
            'Status': 'AGUARDANDO APROVACAO' || 'EM ABERTO' || 'APROVADO',
            'Itens': JSON.stringify(pedidoObject.itens),
            'Data Criacao': pedidoObject.dataCriacao || originalRowData[headerMap['Data Criacao']],
            'Data Ultima Edicao': formatarDataParaISO(new Date()), // Sempre atualiza a data de edição
            'Usuario_Criador': pedidoObject.usuarioCriador || originalRowData[headerMap['Usuario_Criador']],
            'Produto Fornecedor': pedidoObject.produtoFornecedor,
            'Icms St Total': pedidoObject.icmsStTotal
        };

        // Cria a linha de dados na ordem exata dos cabeçalhos da planilha
        const rowData = headers.map(header => dataToSave[String(header).trim()] !== undefined ? dataToSave[String(header).trim()] : '');
        Logger.log(`[editarPedido] 3. Dados prontos para serem escritos na linha ${rowIndexToUpdate}.`);

        // Atualiza a linha inteira na planilha
        sheet.getRange(rowIndexToUpdate, 1, 1, rowData.length).setValues([rowData]);
        Logger.log(`[editarPedido] 4. Pedido ${pedidoObject.numeroDoPedido} atualizado com sucesso.`);

        return { status: 'success', message: 'Pedido atualizado com sucesso!' };

    } catch (e) {
        Logger.log(`[editarPedido] ERRO FATAL: ${e.message}. Stack: ${e.stack}`);
        return { status: 'error', message: `Erro no servidor ao atualizar o pedido: ${e.message}` };
    }
}

// ================================================================
// FUNÇÕES PARA PEDIDOS APROVADOS
// ================================================================
/**
 * Busca na planilha todos os pedidos criados por um usuário específico que tenham o status "Aprovado".
 * @param {string} usuarioLogado O nome de usuário (login) do criador do pedido.
 * @returns {object} Um objeto com o status da operação e os dados dos pedidos encontrados.
 */
function getMeusPedidosAprovados(usuarioLogado) {
  try {
    if (!usuarioLogado) {
      throw new Error("O nome do usuário não foi fornecido.");
    }

    const sheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName(NOME_DA_ABA_DE_PEDIDOS);
    if (!sheet) { throw new Error("Aba de pedidos não encontrada."); }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Pega os cabeçalhos

        Logger.log("DEBUG: Cabeçalhos encontrados na planilha: " + JSON.stringify(headers));

    // Encontra os índices das colunas necessárias
    const colunas = {
      usuarioCriador: headers.indexOf("Usuario Criador"),
      status: headers.indexOf("Status"),
      numeroDoPedido: headers.indexOf("Número do Pedido"),
      data: headers.indexOf("Data"),
      fornecedor: headers.indexOf("Fornecedor"),
      empresa: headers.indexOf("Empresa"),
      totalGeral: headers.indexOf("Total Geral")
    };
    Logger.log("DEBUG: Índices das colunas encontrados: " + JSON.stringify(colunas));

    // Validação para garantir que todas as colunas foram encontradas
    for (const key in colunas) {
        if (colunas[key] === -1) {
            throw new Error(`A coluna "${key}" não foi encontrada na planilha de Pedidos.`);
        }
    }

    // ✅ NOVO: Lógica para o filtro de data
    const hoje = new Date();
    const tresDiasAtras = new Date();
    tresDiasAtras.setDate(hoje.getDate() - 3);
    // Zera a hora para garantir que a comparação inclua o dia inteiro
    tresDiasAtras.setHours(0, 0, 0, 0);

    const pedidosAprovados = data
      .filter(row => {
        const dataDoPedido = new Date(row[colunas.data]);
        // ✅ NOVO: Adiciona a verificação da data ao filtro
        return row[colunas.usuarioCriador] === usuarioLogado && 
               row[colunas.status] === "Aprovado" &&
               dataDoPedido >= tresDiasAtras;
      })
      .map(row => {
        const dataDoPedido = row[colunas.data];
        return {
          'Número_do_Pedido': row[colunas.numeroDoPedido],
          'Data': dataDoPedido instanceof Date ? Utilities.formatDate(dataDoPedido, "GMT-03:00", "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'") : dataDoPedido,
          'Fornecedor': row[colunas.fornecedor],
          'Empresa': row[colunas.empresa],
          'Total Geral': row[colunas.totalGeral]
        };
      
    });
    
    Logger.log(`Encontrados ${pedidosAprovados.length} pedidos aprovados para o usuário ${usuarioLogado}.`);
    return { status: 'success', data: pedidosAprovados };

  } catch(e) {
    Logger.log(`Erro em getMeusPedidosAprovados: ${e.message}`);
    return { status: 'error', message: e.message };
  }
}     

/**
 * Altera o status de um pedido para "Cancelado" na planilha 'Pedidos'.
 * @param {string} numeroPedido - O número do pedido a ser cancelado.
 * @param {string} empresaId - O ID da empresa do pedido.
 * @returns {object} Um objeto com o status da operação.
 */
function cancelarPedidoBackend(numeroPedido, empresaId, usuarioCancelou, dataCancelamento) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
    if (!sheet) {
      throw new Error("Planilha 'Pedidos' não encontrada.");
    }

    const mapaEmpresas = _criarMapaDeEmpresas();

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values.shift();

    // Encontra os índices das colunas necessárias
    const colNumero = headers.indexOf("Número do Pedido");
    const colEmpresa = headers.indexOf("Empresa");
    const colStatus = headers.indexOf("Status");
    const colUsuarioCanc = headers.indexOf("Usuario Cancelamento");
    const colDataCanc = headers.indexOf("Data Cancelamento");

    if (colNumero === -1 || colEmpresa === -1 || colStatus === -1) {
      throw new Error("Não foi possível encontrar as colunas 'Número do Pedido', 'Empresa' ou 'Status'.");
    }

    // Procura pela linha correspondente ao pedido
    let rowIndex = -1;
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][colNumero]).trim() === String(numeroPedido).trim() && String(values[i][colEmpresa]).trim() === String(empresaId).trim()) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      return { status: 'error', message: 'Pedido não encontrado para cancelamento.' };
    }

    // Atualiza o status na planilha. A linha no array 'values' é 'rowIndex',
    // mas na planilha é 'rowIndex + 2' (porque o array começa em 0 e removemos o cabeçalho).
    const rowToUpdate = rowIndex + 2;
    sheet.getRange(rowToUpdate, colStatus + 1).setValue("Cancelado");

    if (colUsuarioCanc !== -1) {
      sheet.getRange(rowToUpdate, colUsuarioCanc + 1).setValue(usuarioCancelou);
    }
    if (colDataCanc !== -1) {
      // Formata a data para um formato mais legível na planilha
      const dataFormatada = Utilities.formatDate(new Date(dataCancelamento), "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");
      sheet.getRange(rowToUpdate, colDataCanc + 1).setValue(dataFormatada);
    }

    // --- BLOCO DE NOTIFICAÇÃO AJUSTADO PARA ADMINS ---
    Logger.log(`Iniciando notificação de cancelamento para os administradores...`);

    const adminsParaNotificar = _getAdminUsers().data || [];

    if (adminsParaNotificar.length > 0) {
      const nomeDaEmpresa = mapaEmpresas[empresaId]?.empresa || `Empresa ID ${empresaId}`;
      const mensagem = `🚫 <b>Alerta: Pedido Cancelado</b>\n\n` +
                         `O pedido <b>Nº ${numeroPedido}</b> da empresa ${nomeDaEmpresa}, foi cancelado no portal.\n\n` +
                         `<b>Cancelado por:</b> ${usuarioCancelou}`;
    
      adminsParaNotificar.forEach(admin => {
        enviarNotificacaoTelegram(admin.chatId, mensagem);
      });
      Logger.log(`Notificação de cnacelamente enviado para ${adminsParaNotificar.length} admin(s).`);
  }  
    Logger.log(`Pedido #${numeroPedido} da empresa #${empresaId} foi cancelado com sucesso.`);
    return { status: 'ok', message: `Pedido #${numeroPedido} foi cancelado com sucesso.` };

  } catch (e) {
    Logger.log(`ERRO em cancelarPedidoBackend: ${e.message}`);
    return { status: 'error', message: `Ocorreu um erro no servidor: ${e.message}` };
  }
}

function mapearCabecalhoPedidos(idDaPlanilha) {
  try {
    const planilha = SpreadsheetApp.openById(idDaPlanilha);
    const aba = planilha.getSheetByName('Pedidos');

    if (!aba) {
      Logger.log('Erro: A aba "Pedidos" não foi encontrada na planilha.');
      return null;
    }

    // Pega todos os valores da primeira linha (cabeçalho)
    const cabecalhos = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];

    // Define um "esquema" com os nomes que queremos usar no código e suas possíveis variações na planilha
    const schema = {
      numeroDoPedido: ['NUMERO DO PEDIDO', 'NÚMERO DO PEDIDO'],
      empresa: ['EMPRESA'],
      data: ['DATA'],
      fornecedor: ['FORNECEDOR'],
      status: ['STATUS'],
      itens: ['ITENS'],
      totalGeral: ['TOTAL GERAL'],
      dataCriacao: ['DATA CRIACAO', 'DATA CRIAÇÃO'],
      dataUltimaEdicao: ['DATA ULTIMA EDICAO', 'DATA ÚLTIMA EDIÇÃO'],
      aliquotaImposto: ['ALIQUOTA IMPOSTO', 'ALÍQUOTA IMPOSTO'],
      usuarioCriador: ['USUARIO CRIADOR', 'USUÁRIO CRIADOR'],
      // Adicione outros campos que você precisar aqui...
    };

    const mapaDeColunas = {};

    // Normaliza os cabeçalhos da planilha para comparação (converte para maiúsculas e remove espaços)
    const cabecalhosNormalizados = cabecalhos.map(h => String(h).toUpperCase().trim());

    // Itera sobre o nosso esquema para encontrar o índice de cada campo
    for (const chave in schema) {
      mapaDeColunas[chave] = -1; // Valor padrão caso não encontre
      const nomesPossiveis = schema[chave];
      
      for (const nome of nomesPossiveis) {
        const index = cabecalhosNormalizados.indexOf(nome);
        if (index !== -1) {
          mapaDeColunas[chave] = index;
          break; // Encontrou, pode parar de procurar por este campo
        }
      }
    }

    Logger.log('Mapeamento de cabeçalho criado com sucesso: ' + JSON.stringify(mapaDeColunas));
    return mapaDeColunas;

  } catch (e) {
    Logger.log('Erro crítico ao tentar mapear o cabeçalho: ' + e.toString());
    return null;
  }
}
