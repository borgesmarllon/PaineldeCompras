/**
 * Envia uma mensagem de notificação para um usuário específico no Telegram.
 * @param {string} chatId O ID do chat do Telegram do destinatário.
 * @param {string} mensagem A mensagem a ser enviada. Suporta tags HTML simples como <b>, <i>, <a>.
 */
function enviarNotificacaoTelegram(chatId, mensagem, botoes) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
    if (!token) {
      throw new Error("Token do bot do Telegram não encontrado nas Propriedades de Script.");
    }

    const url = `https://api.telegram.org/bot${token}/sendMessage`;
    
    const payload = {
      'chat_id': chatId,
      'text': mensagem,
      'parse_mode': 'HTML' // Permite usar <b>para negrito</b>, <i>para itálico</i>, etc.
    };

    if (botoes && botoes.length > 0) {
      payload.reply_markup = JSON.stringify({ inline_keyboard: botoes });
    }

    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const content = response.getContentText();
    
    // Logar a resposta completa do Telegram para depuração
    Logger.log(`Notificação enviada para o Chat ID ${chatId}. Resposta: ${response.getResponseCode()}`);
    Logger.log(`Resposta do Telegram: ${response.getContentText()}`);

    const jsonResponse = JSON.parse(content);

    // Se a resposta for um sucesso (código 200), retornamos true
    if (code === 200 && jsonResponse.ok) {
      return true;
    } else {
      Logger.log(`Erro ao enviar a notificação para o Chat ID ${chatId}: ${jsonResponse.description || 'Sem descrição do erro'}`);
      return false;
    }

  } catch (e) {
    // Caso algum erro ocorra durante o processo, logamos a exceção
    Logger.log(`ERRO ao enviar notificação para o Telegram (Chat ID: ${chatId}): ${e.message}`);
    return false;
  }
}

/**
 * Função auxiliar para buscar o Chat ID de um usuário na planilha 'Usuários' usando o nome de usuário.
 * @param {string} username O nome de usuário a ser buscado.
 * @return {string|null} O Chat ID do Telegram ou null se não for encontrado.
 */
function _getChatIdPorUsername(username) {
  try {
    const userSheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName("Usuarios");
    const userData = userSheet.getDataRange().getValues();
    const headers = userData.shift();

    const colUsername = headers.indexOf("USUARIO"); 
    const colChatId = headers.indexOf("TELEGRAM CHAT ID");

    if (colUsername === -1 || colChatId === -1) {
        Logger.log("AVISO: Coluna 'Nome de Usuário' ou 'Telegram Chat ID' não encontrada na aba 'Usuários'.");
        return null;
    }

    // Procura pelo username (sem diferenciar maiúsculas/minúsculas)
    for (const row of userData) {
      if (String(row[colUsername]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
        return row[colChatId]; // Retorna o ID encontrado
      }
    }
    return null; // Usuário não encontrado

  } catch (e) {
    Logger.log("Erro ao buscar Chat ID por username: " + e.message);
    return null;
  }
}

function _getAdminUsers() {
  try {
    const userSheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName("Usuarios");
    const userData = userSheet.getDataRange().getValues();
    const headers = userData.shift();

    const colUsername = headers.indexOf("USUARIO");
    const colChatId = headers.indexOf("TELEGRAM CHAT ID");
    const colPerfil = headers.indexOf("PERFIL");

    if ([colUsername, colChatId, colPerfil].includes(-1)) {
      throw new Error("Colunas essenciais ('Usuario', 'Telegram Chat ID', 'Perfil') não encontradas na aba 'Usuários'.");
    }

    const admins = [];
    userData.forEach(row => {
      const perfil = String(row[colPerfil]).trim().toLowerCase();
      const chatId = row[colChatId];

      // Se o perfil é 'admin' e existe um Chat ID, adiciona à lista
      if (perfil === 'admin' && chatId) {
        admins.push({
          nome: row[colUsername],
          chatId: chatId
        });
      }
    });

    Logger.log(`Encontrados ${admins.length} administradores para notificação.`);
    return { status: 'success', data: admins };

  } catch (e) {
    Logger.log("Erro ao buscar usuários admin: " + e.message);
    return { status: 'error', data: [], message: e.message }; // Retorna lista vazia em caso de erro
  }
}

function _isAdmin(telegramId) {
    try {
    if (!telegramId) return false;

    const resultadoAdmins = _getAdminUsers(); // Retorna { status, data }
    const listaDeAdmins = resultadoAdmins.data || []; // Extrai o array de dados

    return listaDeAdmins.some(admin => admin.chatId == telegramId);
    
  } catch (e) {
    Logger.log(`Erro em _isAdmin: ${e.stack}`);
    return false;
  }
}
/**
 * Busca na planilha 'Usuários' e retorna uma lista de todos os usuários
 * com perfil 'Admin' que possuem um Telegram Chat ID.
 * @return {Array} Um array de objetos, onde cada objeto é um admin com seu chatId. Ex: [{nome: 'joao.silva', chatId: '123456789'}]
 */
function _api_getAdminUsers() {
  try {
    if (!telegramId) return false;

    // 1. Pega o OBJETO de resposta da função
    const resultadoAdmins = _getAdminUsers();

    // 2. Extrai o ARRAY de dados de dentro do objeto
    const listaDeAdmins = resultadoAdmins.data || [];

    // 3. Procura o ID no ARRAY e retorna true ou false
    return listaDeAdmins.some(admin => admin.chatId == telegramId);
    
  } catch (e) {
    Logger.log(`Erro em _isAdmin: ${e.stack}`);
    return false; // Retorna 'false' em caso de erro para segurança
  }
}

/**
 * Ponto de entrada para TODAS as interações do Telegram (Webhook).
 * Processa tanto mensagens de texto quanto cliques em botões.

function processarRejeicao(chatId, texto, adminInfo, userCache) {
  const numeroPedidoParaRejeitar = userCache.get(`rejeitando_${chatId}`);
  if (numeroPedidoParaRejeitar) {
    userCache.remove(`rejeitando_${chatId}`);
    const resultado = atualizarStatusPedido(numeroPedidoParaRejeitar, "Rejeitado", texto, adminInfo);
    let mensagemConfirmacao;
    if (resultado && resultado.status === 'success') {
      mensagemConfirmacao = `👍 Pedido <b>${numeroPedidoParaRejeitar}</b> foi rejeitado com sucesso. O criador será notificado.`;
    } else {
      mensagemConfirmacao = `⚠️ Falha ao registrar a rejeição para o pedido ${numeroPedidoParaRejeitar}.`;
    }
    enviarNotificacaoTelegram(chatId, mensagemConfirmacao);
  }
}

function _api_buscarPedido(params) {
  const { numeroPedido, empresaId } = params;

  if (!numeroPedido || !empresaId) {
    return { status: "bad_request", data: "Número do pedido ou ID da empresa não fornecidos." };
  }

  const pedido = getPedidoCompletoPorId(numeroPedido, empresaId); // Você já tem essa lógica

   Logger.log("Objeto Pedido Recebido: " + JSON.stringify(pedido)); 

  if (pedido) {
    // A função _formatarPedidoParaTelegram já retorna o texto que precisamos!
    const mensagemResposta = _formatarPedidoParaTelegramv2(pedido);
    return { status: "success", data: mensagemResposta };
  } else {
    const mensagemErro = `❌ Pedido <b>${numeroPedido}</b> não encontrado na empresa <b>${empresaId}</b>.`;
    return { status: "not_found", data: mensagemErro };
  }
}*/

function doPost(e) {
  // A requisição do Python virá com um 'action' e 'params'.
  const request = JSON.parse(e.postData.contents);
  const action = request.action; 
  const params = request.params || {};
  
  // Usamos um switch para rotear a ação para a função correta.
  switch (action) {
    case 'ping':
      return apiResponse({ status: "success", data: "pong" });
    
    case 'buscar_por_fornecedor':
      return apiResponse(_api_buscarPorFornecedor(params));
    
    case 'buscar_pedido':
      return apiResponse(_api_buscarPedido(params));
    
    case 'obter_detalhes_pedido':
      return apiResponse(_api_obterDetalhes(params));
      
    case 'aprovar_pedido':
      return apiResponse(_api_processarAprovacao(params));
      
    case 'rejeitar_pedido':
      return apiResponse(_api_processarRejeicao(params));

    case 'obter_admins':
      return apiResponse(_getAdminUsers());

    case 'criarMapaDeFornecedoresv2':
      return apiResponse({ status: "success", data: criarMapaDeFornecedoresv2() });
    
    case 'gerar_pdf_pedido':
      return apiResponse(gerarPdfPedido(params.numeroPedido, params.empresaId));
    
    case 'buscar_por_placa':
       return apiResponse(buscarPorPlaca(params));

    case 'relatorio_pdf':
        return apiResponse(generatePdfReport(params));

    case 'relatorio_xls':
        return apiResponse(generateXlsReport(params));
    
    case 'criar_pedido':
        return apiResponse(_api_criarPedido(params.pedido, params.userInfo));
    
    case 'obter_empresas':
        return apiResponse({ status: "success", data: _criarMapaDeEmpresas()});
    
    case 'salvar_rascunho':
      return apiResponse(_api_salvarRascunho(params));

    case 'carregar_rascunho':
      return apiResponse(_api_carregarRascunho(params));

    case 'calcular_imposto_simples':
      return apiResponse(calcularStModal(params));

    case 'getDashboardData':
      return apiResponse(getDashboardData(params));

    case 'listar_veiculos':
      return apiResponse(_api_listarVeiculos());

    case 'obter_opcoes_pagamento':
      return apiResponse(_api_obterOpcoesPagamento());
    
    case 'consultar_cnpj_e_opcoes':
      return apiResponse(_api_consultarCnpjEopcoes(params));
    
    case 'finalizar_cadastro_fornecedor':
      return apiResponse(_api_finalizarCadastroFornecedor(params));

    default:
      return apiResponse({ status: "error", data: `Ação desconhecida: ${action}` });
  }
}

/**
 * Função auxiliar para padronizar e retornar todas as respostas para o Python.
 * @param {Object} payload O objeto de resultado da função da API.
 * @return {ContentService} Uma resposta em formato JSON.
 */
function apiResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// =================================================================
// FUNÇÕES DA API INTERNA
// Estas são as funções que o `doPost` chama. Elas contêm a lógica principal.
// Elas NÃO falam com o Telegram, apenas retornam um objeto.
// =================================================================

/**
 * [API] Busca um pedido completo e retorna o texto formatado.
 * @param {Object} params Objeto com `numeroPedido` e `empresaId`.
 * @return {Object} Payload com status e dados.
 */
function _api_buscarPedido(params) {
  try {
    const { mainSearch: numeroPedido, empresaId, userInfo } = params; 
    if (!numeroPedido || !empresaId) {
      throw new Error("Parâmetros insuficientes.");
    }

    loggersheet(`API: buscando pedido ${numeroPedido} na empresa ${empresaId}`);
    const pedido = getPedidoCompletoPorId(numeroPedido, empresaId);

    if (!pedido) {
      // Se o pedido não existe, retorna 'não encontrado' para todos.
      return { status: "not_found", data: `❌ Pedido <b>${numeroPedido}</b> não encontrado na empresa <b>${empresaId}</b>.` };
    }
    
    // --- FILTRO DE SEGURANÇA APLICADO AQUI ---
    const isAdmin = _isAdmin(userInfo ? userInfo.id : null);
    const statusDoPedido = (pedido.status || '').toUpperCase().trim();

    if (statusDoPedido === 'AGUARDANDO APROVACAO' && !isAdmin) {
      Logger.log(`Acesso negado ao pedido ${numeroPedido} para o usuário não-admin ID ${userInfo.id}.`);
      return { status: "not_found", data: `❌ Pedido <b>${numeroPedido}</b> não encontrado.` };
    }
    // --- FIM DO FILTRO ---

    // 2. Se o usuário tem permissão, formata e retorna a mensagem
    const mensagemFormatada = _formatarPedidoParaTelegramv2(pedido);
    const botoes = [[{ text: "🔎 Ver detalhes", callback_data: `detalhes:${numeroPedido}:${empresaId}`}, { text: "📄 PDF", callback_data: `pdf:${numeroPedido}:${empresaId}`}]];
    return { status: "success", data: mensagemFormatada, botoes: botoes };
    
  } catch (err) {
    loggersheet(`Erro em _api_buscarPedido: ${err.message}`);
    return { status: "error", data: "⚠️ Ocorreu um erro ao buscar o pedido." };
  }
}

/**
 * [API] Obtém os detalhes de um pedido para exibição com botões de ação.
 * @param {Object} params Objeto com `numeroPedido` e `empresaId`.
 * @return {Object} Payload com status e o texto formatado para detalhes.
 */
function _api_obterDetalhes(params) {
  try {
    const { numeroPedido, empresaId } = params;
    if (!numeroPedido || !empresaId) throw new Error("Parâmetros insuficientes.");

    const pedido = getPedidoCompletoPorId(numeroPedido, empresaId);
    if (!pedido) {
      return { status: "not_found", data: `😕 Pedido ${numeroPedido} não encontrado.` };
    }
    
    // Formata o corpo principal da mensagem de detalhes
    const textoDetalhes = _formatarDetalhesParaTelegram(pedido);
    return { status: "success", data: textoDetalhes };
  } catch(err) {
    Logger.log(`Erro em _api_obterDetalhes: ${err.message}`);
    return { status: "error", data: "⚠️ Ocorreu um erro ao buscar os detalhes do pedido." };
  }
}

/**
 * Processa a lógica de um clique de botão (callback_query).
*/
function processCallbackQuery(callbackQuery) {
  // Pega o serviço de bloqueio do script.
  const lock = LockService.getScriptLock();
    
  // Tenta obter o bloqueio por 10 segundos.
  // Se outro processo já estiver rodando, ele não conseguirá o bloqueio.
  const gotLock = lock.tryLock(10000); 

  if (!gotLock) {
    // Não conseguiu o bloqueio. Significa que outro processo já está em execução.
    // Loga o evento e para a execução imediatamente para evitar o travamento.
    Logger.log("Não foi possível obter o bloqueio. Outra instância provavelmente está em execução. Ignorando este clique.");
    return; // Para a função aqui.
  }

  try{
  const chatId = callbackQuery.message.chat.id;
  const messageId = callbackQuery.message.message_id;
   
    // 1. Primeiro verifica e responde à callback query
    _answerCallbackQuery(callbackQuery.id);
      Logger.log(`Callback expirada ou inválida: ${callbackQuery.id}`);

    const callbackData = callbackQuery.data;
    Logger.log(`Processando callback: ${callbackData}`);
    
    // 2. Validação básica
    if (!callbackData || typeof callbackData !== 'string') {
      throw new Error("Dados de callback inválidos");
    }

    const [action, numeroPedido, empresaId] = callbackData.split(':');
    
    // 3. Validação dos parâmetros
    if (!action || !numeroPedido || !empresaId) {
      throw new Error(`Formato de callback inválido: ${callbackData}`);
    }

    // 4. Processamento das ações
    switch (action) {
      case 'aprovar':
        _processarAprovacao(chatId, messageId, numeroPedido, empresaId, callbackQuery.from);
        break;
        
      case 'rejeitar':
        _processarRejeicaoInicial(chatId, messageId, empresaId, numeroPedido);
        break;
        
      case 'detalhes':
        _processarDetalhesPedido(chatId, messageId, numeroPedido, empresaId);
        break;
        
      default:
        throw new Error(`Ação desconhecida: ${action}`);
    }
    
  } catch (err) {
    Logger.log(`ERRO em processCallbackQuery: ${err.toString()}\n${err.stack}`);
    const chatId = callbackQuery.message.chat.id;
    const messageId = callbackQuery.message.message_id;
    _handleCallbackError(chatId, messageId, err);

  } finally {
    // ESSA PARTE É CRUCIAL!
    // Não importa se o script funcionou ou deu erro,
    // o bloqueio (crachá) DEVE ser liberado para que a próxima execução possa funcionar.
    lock.releaseLock();
    Logger.log("Bloqueio liberado.");
  }
}

/**
 * [API] Processa a aprovação de um pedido. Usa LockService para segurança.
 * @param {Object} params Objeto com `numeroPedido` e `adminInfo`.
 * @return {Object} Payload com a mensagem de confirmação ou erro.
 */
function _api_processarAprovacao(params) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { status: "locked", data: "Sistema ocupado processando outra solicitação. Tente novamente em alguns segundos." };
  }
  
  try {
    const { numeroPedido, adminInfo } = params;
    const resultado = atualizarStatusPedido(numeroPedido, "Aprovado", "", adminInfo);
    
    if (resultado?.status === 'success') {
      return { status: "success", data: `✅ Pedido <b>${numeroPedido}</b> APROVADO por ${adminInfo.first_name}.` };
    } else {
      return { status: "error", data: `⚠️ Falha ao aprovar o pedido ${numeroPedido}.` };
    }
  } catch(err) {
    loggersheet(`Erro em _api_processarAprovacao: ${err.message}`);
    return { status: "error", data: "⚠️ Ocorreu um erro crítico ao aprovar o pedido." };
  } finally {
    lock.releaseLock();
  }
}

/**
 * [API] Processa a rejeição de um pedido com motivo. Usa LockService.
 * @param {Object} params Objeto com `numeroPedido`, `motivoRejeicao`, `adminInfo`.
 * @return {Object} Payload com a mensagem de confirmação ou erro.
 */
function _api_processarRejeicao(params) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { status: "locked", data: "Sistema ocupado. Tente novamente em alguns segundos." };
  }
  
  try {
    const { numeroPedido, motivoRejeicao, adminInfo } = params;
    const resultado = atualizarStatusPedido(numeroPedido, "Rejeitado", motivoRejeicao, adminInfo);

    if (resultado?.status === 'success') {
      return { status: "success", data: `👍 Pedido <b>${numeroPedido}</b> foi REJEITADO com sucesso.`,criador_chat_id: resultado.criador_chat_id };
    } else {
      return { status: "error", data: `⚠️ Falha ao registrar a rejeição para o pedido ${numeroPedido}.` };
    }
  } catch(err) {
    loggersheet(`Erro em _api_processarRejeicao: ${err.message}`);
    return { status: "error", data: "⚠️ Ocorreu um erro crítico ao rejeitar o pedido." };
  } finally {
    lock.releaseLock();
  }
}

/**
 * [API] Busca TODOS os pedidos de um determinado fornecedor.
 * @param {Object} params Objeto com `nomeFornecedor`.
 * @return {Object} Payload com status e uma LISTA de pedidos encontrados.
 */
function _api_buscarPorFornecedor(params) {
  try {
    const { nomeFornecedor } = params;
    if (!nomeFornecedor) throw new Error("Nome do fornecedor não fornecido.");

    loggersheet(`API: buscando pedidos para o fornecedor: ${nomeFornecedor}`);

    const pedidosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
    const todosOsPedidos = pedidosSheet.getDataRange().getValues();
    const headers = todosOsPedidos.shift(); // Remove e guarda o cabeçalho

    const colFornecedor = headers.indexOf("Fornecedor");
    const colNumeroPedido = headers.indexOf("Número do Pedido");
    const colStatus = headers.indexOf("Status");
    const colTotal = headers.indexOf("Total Geral");
    const colEmpresa = headers.indexOf("Empresa");

    const pedidosEncontrados = [];

    // Itera por todos os pedidos para encontrar correspondências
    todosOsPedidos.forEach(linha => {
      const fornecedorNaPlanilha = linha[colFornecedor].toString().toLowerCase();
      
      // Usamos .includes() para uma busca mais flexível (ex: "Carlos" encontra "Carlos Augusto")
      if (fornecedorNaPlanilha.includes(nomeFornecedor.toLowerCase())) {
        pedidosEncontrados.push({
          numero: String(linha[colNumeroPedido]).trim(),
          status: linha[colStatus],
          total: linha[colTotal],
          empresaId: String(linha[colEmpresa]).trim() 
        });
      }
    });

    if (pedidosEncontrados.length > 0) {
      // Retorna uma lista de objetos, não mais um texto formatado
      return { status: "success", data: pedidosEncontrados };
    } else {
      return { status: "not_found", data: `Nenhum pedido encontrado para o fornecedor contendo "${nomeFornecedor}".` };
    }
  } catch (err) {
    loggersheet(`Erro em _api_buscarPorFornecedor: ${err.message}`);
    return { status: "error", data: "⚠️ Ocorreu um erro ao buscar os pedidos." };
  }
}

/**
 * [API] Busca os pedidos por placa;
 * @param {Object} params Objeto com placaVeiculo
 * @return {Object} Payload com status e uma LISTA de pedidos encontrados.
 */

function buscarPorPlaca(params) {
  try {
    if (!params || !params.placaVeiculo) {
      return { status: 'error', message: 'Placa obrigatória.' };
    }
    const placaBusca = String(params.placaVeiculo).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
    const sheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName('Pedidos');
    if (!sheet) {
      throw new Error("Aba 'Pedidos' não encontrada.");
    }
    const dataRows = sheet.getDataRange().getValues();
    const headers = dataRows.shift();
    const colunas = {
      placa: headers.findIndex(h => h.toUpperCase().includes('PLACA')),
      numero: headers.findIndex(h => 
        h.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").includes('NUMERO DO PEDIDO')
      ),
      empresaId: headers.findIndex(h => h.toUpperCase() === 'EMPRESA'),
      fornecedor: headers.findIndex(h => h.toUpperCase() === 'FORNECEDOR'),
      status: headers.findIndex(h => h.toUpperCase() === 'STATUS')
    };
    if (colunas.placa === -1) throw new Error("Coluna 'Placa' não encontrada na planilha.");   
    if (colunas.numero === -1) throw new Error("Coluna contendo 'Número do Pedido' não foi encontrada. Verifique o cabeçalho.");

    const pedidos = [];

    for (const row of dataRows) {
      const placaPlanilha = String(row[colunas.placa]).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
      
      if (placaPlanilha === placaBusca) {
 
        pedidos.push({
          numero: row[colunas.numero],
          empresaId: row[colunas.empresaId],
          fornecedor: row[colunas.fornecedor],
          status: row[colunas.status],
          placa: row[colunas.placa]
        });
      }
    }
    
    return { status: 'success', data: pedidos };
    
  } catch (e) {
    Logger.log(`Erro em buscarPorPlaca_otimizada: ${e.toString()}`);
    return { status: 'error', message: e.toString() };
  }
}

/**
 * [UTILITÁRIO - MODIFICADO] Formata os detalhes para o corpo da mensagem.
 * O Python cuidará de adicionar os botões.
 */
function _formatarDetalhesParaTelegram(pedido) {

  const mapaEmpresas = _criarMapaDeEmpresas();
  const idDaEmpresa = pedido.empresaId || pedido.empresa_id;
  const nomeDaEmpresa = mapaEmpresas[idDaEmpresa]?.empresa || `Empresa ID ${idDaEmpresa}`;  
  let itensTexto = pedido.itens?.map((item, index) => 
    `\n    ${index + 1}. ${item.descricao} ` +
    `(${item.quantidade} x ${_formatCurrency(item.precoUnitario)}) = ` +
    `<b>${_formatCurrency(item.totalItem)}</b>`
  ).join('') || "\n    Nenhum item detalhado encontrado.";

  return `📄 <b>Detalhes do Pedido Nº ${pedido.numero_do_pedido}</b>\n\n` +
    `<b>Empresa:</b> ${nomeDaEmpresa}\n` +
    `<b>Fornecedor:</b> ${pedido.fornecedor}\n` +
    `<b>Criado por:</b> ${pedido.usuarioCriadorInfo?.nome || pedido.usuario_criador}\n` +
    `<b>Valor Total:</b> ${_formatCurrency(pedido.total_geral)}\n` +
    `<b>Impostos:</b> ${_formatCurrency(pedido.icms_st_total)}\n\n` +
    `<b>Itens:</b>${itensTexto}\n\n` +
    `<i>O que deseja fazer?</i>`;
}

/**
 * [UTILITÁRIO - MANTIDO] Formata a visualização completa de um pedido..
 */
function _formatarPedidoParaTelegramv2(pedido) {
  if (!pedido) {
    return '❌ Não foi possível encontrar as informações do pedido.';
  }

  const itensTexto = pedido.itens?.map((item, index) =>
    `  • ${item.descricao} (${item.quantidade} x ${_formatCurrency(item.precoUnitario)}) = <b>${_formatCurrency(item.totalItem)}</b>`
  ).join('\n') || "Nenhum item detalhado encontrado.";

  return `
📄 <b>Detalhes do Pedido Nº ${pedido.numero_do_pedido || 'N/A'}</b>
<b>Empresa:</b> ${pedido.empresaInfo?.empresa || pedido.empresaId || 'N/A'}
<b>Status:</b> ${pedido.status || 'N/A'}
<b>Fornecedor:</b> ${pedido.fornecedorInfo?.nome || 'N/A'}
<b>Criado por:</b> ${pedido.usuarioCriadorInfo?.nome || 'N/A'}
<b>Valor Total:</b> ${_formatCurrency(pedido.total_geral || 0)}
<b>Impostos:</b> ${_formatCurrency(pedido.icms_st_total || 0)}

<b>Itens:</b>
${itensTexto}
`;
}

/**
 * [UTILITÁRIO - MANTIDO] Funções de apoio
 */
function _formatCurrency(value) {
  // Implementação da sua função
  if (typeof value !== 'number') value = parseFloat(value) || 0;
  return value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}
function _processarDetalhesPedido(chatId, messageId, numeroPedido, empresaId) {
  const pedido = getPedidoCompletoPorId(numeroPedido, empresaId);
  
  if (!pedido) {
    _editMessageText(chatId, messageId, `😕 Pedido ${numeroPedido} não encontrado.`);
    return;
  }

  // Formatação dos itens
  let itensTexto = pedido.itens?.map((item, index) => 
    `\n    ${index + 1}. ${item.descricao} ` +
    `(${item.quantidade} x ${_formatCurrency(item.precoUnitario)}) = ` +
    `<b>${_formatCurrency(item.totalItem)}</b>`
  ).join('') || "\n    Nenhum item detalhado encontrado.";

  // Montagem da mensagem
  const mensagem = `📄 <b>Detalhes do Pedido Nº ${numeroPedido}</b>\n\n` +
    `<b>Empresa:</b> ${pedido.empresaInfo?.razao_social || `ID ${pedido.empresa}`}\n` +
    `<b>Fornecedor:</b> ${pedido.fornecedor}\n` +
    `<b>Criado por:</b> ${pedido.usuarioCriadorInfo?.nome || pedido.usuario_criador}\n` +
    `<b>Valor Total:</b> ${_formatCurrency(pedido.total_geral)}\n` +
    `<b>Impostos:</b> ${_formatCurrency(pedido.icms_st_total)}\n\n` +
    `<b>Itens:</b>${itensTexto}\n\n` +
    `<i>O que deseja fazer?</i>`;

  const botoes = [
    [
      { text: "✅ Aprovar", callback_data: `aprovar:${numeroPedido}:${empresaId}` },
      { text: "❌ Rejeitar", callback_data: `rejeitar:${numeroPedido}:${empresaId}` }
    ]
  ];

  _editMessageText(chatId, messageId, mensagem, botoes);
}

function _handleCallbackError(chatId, messageId, error) {
  Logger.log(`Erro no chat ${chatId}: ${error.message}`);
  
  // Só envia mensagem de erro se a callback ainda for recente
  if (_isRecentInteraction()) {
    _editMessageText(chatId, messageId, 
      "⚠️ Ocorreu um erro ao processar sua solicitação. Por favor, tente novamente.");
  }
}


// Funções utilitárias
function _isRecentInteraction(timestamp = Date.now(), maxAgeSeconds = 10) {
  return (Date.now() - timestamp) < (maxAgeSeconds * 1000);
}

function _formatCurrency(value) {
  return _sanitizeCurrency(value).toLocaleString('pt-BR', { 
    style: 'currency', 
    currency: 'BRL' 
  });
}

/**
 * Responde a uma callback query. Isso faz o ícone de 'carregando' no botão desaparecer.
 * @param {string} callbackQueryId O ID da query de callback.
 
function _answerCallbackQuery(callbackQueryId) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
    const url = `https://api.telegram.org/bot${token}/answerCallbackQuery`;

    const payload = {
      'callback_query_id': String(callbackQueryId)
    };

    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true // Evita que uma falha aqui pare o script
    };

    UrlFetchApp.fetch(url, options);
    // Não é necessário retornar nada.

  } catch (e) {
    // Apenas loga o erro, não queremos que isso pare o fluxo principal.
    Logger.log(`Falha ao tentar responder à callback query ${callbackQueryId}: ${e.message}`);
  }
}
*/

/**
 * Edita uma mensagem existente no Telegram, geralmente para remover os botões e mostrar um status.
 * @param {string} chatId O ID do chat onde a mensagem está.
 * @param {string} messageId O ID da mensagem a ser editada.
 * @param {string} novoTexto O novo texto da mensagem.
 */
function _editMessageText(chatId, messageId, novoTexto, botoes) {
 try {
    const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
    const url = `https://api.telegram.org/bot${token}/editMessageText`;
    
    const payload = {
      chat_id: String(chatId),
      message_id: messageId,
      text: novoTexto,
      parse_mode: 'HTML',
      // A "mágica" está aqui:
      // Se 'botoes' for fornecido, usa-o. Senão, usa um array vazio [] para remover o teclado.
      reply_markup: JSON.stringify({
        'inline_keyboard': botoes || [] 
      })
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    Logger.log("Resposta da API 'editMessageText': " + responseText);

  } catch (e) {
    Logger.log(`ERRO CRÍTICO em _editMessageText: ${e.message}`);
  }
}

/**
 * [API] Ponto de entrada para criar um novo pedido de forma segura e completa.
 * Esta função orquestra a geração do número e o salvamento.
 * @param {Object} params - Dados do pedido vindos do bot (sem o número).
 * @param {Object} usuarioInfo - Informações do usuário do Telegram que está criando.
 * @return {Object} Objeto de resposta para o bot.
 */
function _api_criarPedido(params, usuarioInfo) {
  try {
    Logger.log("API: Iniciando processo de criação de pedido...");
    
    // Validação dos dados recebidos do bot
    if (!params.empresaId || !params.fornecedor || !params.itens) {
      throw new Error("Dados insuficientes do bot para criar o pedido (empresa, fornecedor ou itens faltando).");
    }

    let totalGeralCalculado = 0;
    let icmsStTotalCalculado = 0;

    if (params.itens && Array.isArray(params.itens)) {
      params.itens.forEach(item => {
        totalGeralCalculado += Number(item.totalItem) || 0;
        icmsStTotalCalculado += Number(item.icmsSt) || 0;
      });
    }

    // 1.1 GERA O NÚMERO DE FORMA SEGURA NO SERVIDOR
    const novoNumeroPedido = getProximoNumeroPedido(params.empresaId);

    // 2. MONTA O OBJETO COMPLETO
    // Combina os dados do bot com os dados gerados no servidor
    const pedidoCompleto = {
      numero: novoNumeroPedido,
      empresaId: params.empresaId,
      fornecedor: params.fornecedor,
      itens: params.itens,
      data: new Date(), // O servidor define a data para garantir consistência
      totalGeral: totalGeralCalculado,
      valorIcms: icmsStTotalCalculado, 
      placaVeiculo: params.placaVeiculo || '', // Preparando para quando o bot enviar
      nomeVeiculo: params.nomeVeiculo || '',
      observacoes: params.observacoes || ''
    };

    // 3. PEGA O NOME DO USUÁRIO
    const nomeUsuarioCriador = _getUsuarioCriadorPorTelegramId(usuarioInfo);
    
    // 4. CHAMA A FUNÇÃO PARA SALVAR NA PLANILHA
    const resultadoSalvar = salvarPedidoBot(pedidoCompleto, nomeUsuarioCriador);
    
    // 5. RETORNA UMA RESPOSTA FINAL PARA O BOT
    if (resultadoSalvar.status === 'ok') {
        return { 
          status: "success", 
          // A mensagem de sucesso agora inclui o número gerado!
          data: `✅ Pedido <b>Nº ${novoNumeroPedido}</b> criado com sucesso!` 
        };
    } else {
        // Se salvarPedidoBot falhou, repassa o erro
        throw new Error(resultadoSalvar.message);
    }

  } catch (e) {
    Logger.log(`ERRO em _api_criarPedido: ${e.stack}`);
    return { status: 'error', message: `Erro no servidor: ${e.message}` };
  }
}

/**
 * Salva um objeto de pedido em uma linha na planilha "Pedidos".
 * @param {Object} pedido O objeto contendo os detalhes do pedido.
 * @param {string} usuarioLogado O login do usuário que está realizando a ação.
 * Salva um pedido montado no bot do telegram
 * @returns {{status: string, message: string}} Um objeto com o resultado da operação.
 */
function salvarPedidoBot(pedido, usuarioLogado) {
  console.log('📋 === INÍCIO salvarPedidoBot ===');
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
  const mapaFornecedores = criarMapaDeFornecedoresv2();
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
          const listaAdmins = _getAdminUsers();  // já retorna lista direta

          Logger.log(`adminsResponse: ${JSON.stringify(listaAdmins)}`);

          if (listaAdmins.length > 0) {
            Logger.log(`Encontrados ${listaAdmins.length} administradores para notificação.`);

            const numeroPedidoLimpo = String(dataToSave['Número do Pedido']).replace(/'/g, '');
            const empresaIdLimpo = String(dataToSave['Empresa']).replace(/'/g, '');
            
            // Esta linha agora funciona, pois 'mapaEmpresas' foi definido no início.
            const nomeDaEmpresa = mapaEmpresas[empresaIdLimpo]?.razao_social || `Empresa ID ${empresaIdLimpo}`;
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
                           `O pedido Nº ${numeroPedidoLimpo} da empresa <b>${nomeDaEmpresa}</b> foi criado e aprovado automaticamente.`;
            }
            
            const botoesDeAcao = [
              [
                { text: "✅ Aprovar", callback_data: `aprovar:${numeroPedidoLimpo}:${empresaIdLimpo}` },
                { text: "❌ Rejeitar", callback_data: `rejeitar:${numeroPedidoLimpo}:${empresaIdLimpo}` }
              ]
            ];
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

/**
 * Busca o NOME de um usuário na planilha "USUARIOS" usando o ID do Telegram como chave.
 * @param {object} userInfo O objeto {id, first_name} vindo do bot.
 * @return {string} O nome encontrado na planilha ou o primeiro nome do Telegram como fallback.
 */
function _getUsuarioCriadorPorTelegramId(userInfo) {
  try {
    if (!userInfo || !userInfo.id) {
      Logger.log("Aviso: userInfo ou ID do Telegram não fornecido. Usando 'Bot' como padrão.");
      return 'Usuário do Bot';
    }

    const telegramId = userInfo.id;
    const fallbackName = userInfo.first_name || 'Usuário Desconhecido';

    const userSheet = SpreadsheetApp.getActive().getSheetByName("USUARIOS");
    if (!userSheet) {
      Logger.log("Aviso: Aba 'USUARIOS' não encontrada. Usando o nome do Telegram como fallback.");
      return fallbackName; // Se a planilha não existe, retorna o nome do Telegram
    }

    const headers = userSheet.getRange(1, 1, 1, userSheet.getLastColumn()).getValues()[0];
    const idColIndex = headers.indexOf("TELEGRAM CHAT ID");
    const nomeColIndex = headers.indexOf("NOME");

    if (idColIndex === -1 || nomeColIndex === -1) {
      Logger.log("Aviso: Colunas 'TELEGRAM CHAT ID' ou 'NOME' não encontradas. Usando o nome do Telegram como fallback.");
      return fallbackName;
    }

    // Otimização: Lê apenas as duas colunas necessárias
    const idsColumn = userSheet.getRange(2, idColIndex + 1, userSheet.getLastRow()).getValues();
    const nomesColumn = userSheet.getRange(2, nomeColIndex + 1, userSheet.getLastRow()).getValues();

    // Procura pelo ID
    for (let i = 0; i < idsColumn.length; i++) {
      // Usa '==' para comparar string com número de forma flexível
      if (idsColumn[i][0] == telegramId) {
        const nomeEncontrado = nomesColumn[i][0];
        Logger.log(`ID ${telegramId} encontrado. Usuário: ${nomeEncontrado}`);
        return nomeEncontrado; // Retorna o nome da planilha
      }
    }

    // Se o loop terminar e não encontrar, significa que o usuário não está cadastrado
    Logger.log(`Aviso: Usuário com ID ${telegramId} não encontrado na planilha. Usando o nome do Telegram como fallback.`);
    return fallbackName;

  } catch (e) {
    Logger.log(`ERRO em _getUsuarioCriadorPorTelegramId: ${e.stack}`);
    // Em caso de erro, retorna o nome do Telegram para não quebrar o processo.
    return userInfo.first_name || 'Erro na Busca';
  }
}

function setWebhook() {
  const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  const webAppUrl = "https://script.google.com/macros/s/AKfycbym1M47L17jyjuHJEBOQpoOplT1AnwO3aAKW1j9-imYu1TrY-da-xFlvIYskyo3tmTbkw/exec"; // Pegue em "Implantar" > "Gerenciar Implantações"

  const url = `https://api.telegram.org/bot${token}/setWebhook?url=${webAppUrl}`;
  const response = UrlFetchApp.fetch(url);

  Logger.log(response.getContentText());
}

/**
 * [API] Recebe uma LISTA de itens e retorna a mesma lista com os impostos calculados.
 * Usa a técnica de planilha temporária para ser seguro para múltiplos usuários.
 */
function _api_calcularImpostosLote(params) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaMestre = spreadsheet.getSheetByName('CalculoICMS');
  let tempSheet = null;

  if (!planilhaMestre) {
    return { status: 'error', message: "Planilha mestre 'CalculoICMS' não encontrada." };
  }

  try {
    const itens = params.itens;
    const estado = params.estado; // Estado do fornecedor para o cálculo

    if (!itens || !Array.isArray(itens)) {
      throw new Error("A lista de itens é inválida.");
    }
    
    // 1. CRIA A CÓPIA TEMPORÁRIA (UMA VEZ SÓ)
    const tempSheetName = `TempCalc_${new Date().getTime()}`;
    tempSheet = planilhaMestre.copyTo(spreadsheet).setName(tempSheetName);
    spreadsheet.setActiveSheet(tempSheet);
    
    // 2. ESCREVE OS DADOS DE TODOS OS ITENS NA CÓPIA
    // Prepara os dados para escrita em lote (muito mais rápido)
    const dadosParaEscrever = itens.map(item => [
      estado, // Coluna C (Estado)
      2,      // Coluna D (Regime)
      '',     // Coluna E
      Number(item.totalItem) || 0 // Coluna F (Valor)
    ]);

    // Escreve todos os dados de uma vez, começando da célula C2
    tempSheet.getRange(2, 3, dadosParaEscrever.length, 4).setValues(dadosParaEscrever);

    // 3. FORÇA O RECÁLCULO
    SpreadsheetApp.flush();
    Utilities.sleep(1000); // Pausa para garantir que os cálculos complexos terminem

    // 4. LÊ OS RESULTADOS DE TODOS OS ITENS
    const resultadosRange = tempSheet.getRange(2, 9, itens.length, 6); // Colunas I (Alíquota) até N (Valor Calculado)
    const resultadosCalculados = resultadosRange.getValues();
    
    // 5. ATUALIZA A LISTA ORIGINAL COM OS RESULTADOS
    const itensAtualizados = itens.map((item, index) => {
      const resultadoLinha = resultadosCalculados[index];
      item.aliquotaUsada = resultadoLinha[0]; // Coluna I
      item.icmsSt = resultadoLinha[5];       // Coluna N
      return item;
    });

    return { status: 'success', data: itensAtualizados };

  } catch (e) {
    Logger.log(`Erro durante o cálculo em lote: ${e.stack}`);
    return { status: 'error', message: `Erro no cálculo: ${e.message}` };
  } finally {
    // 6. DELETA A CÓPIA
    if (tempSheet) {
      spreadsheet.deleteSheet(tempSheet);
    }
  }
}

/**
 * [API] Salva o estado da conversa de um usuário na planilha "Rascunhos".
 * Se já existir um rascunho, ele é sobrescrito.
 */
function _api_salvarRascunho(params) {
  try {
    const { usuarioId, dadosRascunho } = params;
    if (!usuarioId || !dadosRascunho) throw new Error("ID do usuário e dados do rascunho são obrigatórios.");

    const rascunhosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rascunhos");
    const idsColumn = rascunhosSheet.getRange("A:A").getValues();
    let userRow = -1;

    for (let i = 0; i < idsColumn.length; i++) {
      if (idsColumn[i][0] == usuarioId) {
        userRow = i + 1;
        break;
      }
    }
    
    const dadosComoJson = JSON.stringify(dadosRascunho);

    if (userRow > 0) {
      // Usuário já tem um rascunho, vamos atualizar
      rascunhosSheet.getRange(userRow, 2).setValue(dadosComoJson);
    } else {
      // Novo rascunho, adiciona nova linha
      rascunhosSheet.appendRow([`'${usuarioId}`, dadosComoJson]);
    }

    return { status: "success", data: "Rascunho salvo com sucesso!" };
  } catch (e) {
    return { status: "error", message: `Erro ao salvar rascunho: ${e.message}` };
  }
}

/**
 * [API] Carrega o rascunho de um usuário e o apaga da planilha.
 */
function _api_carregarRascunho(params) {
  try {
    const { usuarioId } = params;
    if (!usuarioId) throw new Error("ID do usuário é obrigatório.");

    const rascunhosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rascunhos");
    const idsColumn = rascunhosSheet.getRange("A:A").getValues();
    let userRow = -1;

    for (let i = 0; i < idsColumn.length; i++) {
      if (idsColumn[i][0] == usuarioId) {
        userRow = i + 1;
        break;
      }
    }

    if (userRow > 0) {
      const dadosJson = rascunhosSheet.getRange(userRow, 2).getValue();
      const dadosRascunho = JSON.parse(dadosJson);
      
      // Apaga a linha para que o rascunho não seja carregado duas vezes
      rascunhosSheet.deleteRow(userRow);
      
      return { status: "success", data: dadosRascunho };
    } else {
      return { status: "not_found", data: "Nenhum rascunho encontrado." };
    }
  } catch (e) {
    return { status: "error", message: `Erro ao carregar rascunho: ${e.message}` };
  }
}

// Função de API que "publica" sua lista para o bot
function _api_listarVeiculos() {
  try {
    // Sua função perfeita é chamada aqui!
    const listaDeVeiculos = getVeiculosList(); 
    
    // Retorna os dados no formato padrão que o bot espera
    return { status: "success", data: listaDeVeiculos };

  } catch (e) {
    Logger.log(`ERRO em _api_listarVeiculos: ${e.stack}`);
    return { status: "error", message: e.message };
  }
}

/**
 * API para cadastrar um novo fornecedor usando a BrasilAPI para autocompletar os dados.
 * Recebe apenas o CNPJ.
 */
function _api_cadastrarFornecedorViaCnpj(params) {
  try {
    const { cnpj } = params;
    if (!cnpj) {
      throw new Error("O CNPJ é obrigatório.");
    }

    // 1. Consulta o CNPJ na BrasilAPI para obter os dados básicos.
    // Reutilizando sua função já existente para isso.
    const resultadoCnpj = consultarCnpj_V2(cnpj);
    if (resultadoCnpj.status !== 'ok') {
      throw new Error(resultadoCnpj.message);
    }
    const dadosApi = resultadoCnpj.data;

    // 2. Busca os valores padrão da planilha 'Config'.
    // Pega o primeiro item da lista como padrão.
    const condicaoPadrao = getCondicoesPagamento()[0] || 'A PRAZO';
    const formaPadrao = getFormasPagamento()[0] || 'BOLETO';

    // 3. Monta um objeto 'fornecedorObject' completo,
    // idêntico ao que o app web montaria para um NOVO fornecedor.
    const fornecedorParaSalvar = {
      // 'codigo' é deixado em branco para acionar a lógica de CRIAÇÃO
      // dentro de adicionarOuAtualizarFornecedorv2.
      razaoSocial: dadosApi.razaoSocial,
      nomeFantasia: dadosApi.nomeFantasia,
      cnpj: cnpj, // Usamos o CNPJ original para manter a formatação
      endereco: dadosApi.endereco,
      cidade: dadosApi.cidade,
      estado: dadosApi.uf,
      condicaoPagamento: condicaoPadrao,
      formaPagamento: formaPadrao,
      grupo: '', // Grupo será definido pela lógica interna da função de salvar
      regimeTributario: dadosApi.regimeTributario,
    };

    // 4. Chama a sua função de salvamento principal já existente.
    // Toda a lógica de código sequencial, status, grupo e validações será executada aqui.
    const resultadoFinal = adicionarOuAtualizarFornecedorv2(fornecedorParaSalvar);

    // Retorna a resposta da função de salvamento para o bot.
    return resultadoFinal;

  } catch (e) {
    Logger.log(`ERRO em _api_cadastrarFornecedorViaCnpj: ${e.stack}`);
    return { status: "error", message: e.message };
  }
}

/**
 * API para retornar as listas de Condições e Formas de Pagamento da planilha 'Config'.
 */
function _api_obterOpcoesPagamento() {
  try {
    // Reutilizamos as funções que você já tem!
    const condicoes = getCondicoesPagamento();
    const formas = getFormasPagamento();

    return { 
      status: "success", 
      data: {
        condicoes: condicoes,
        formas: formas
      } 
    };
  } catch (e) {
    Logger.log(`ERRO em _api_obterOpcoesPagamento: ${e.stack}`);
    return { status: "error", message: e.message };
  }
}

/**
 * API que recebe os dados de um fornecedor (já consultados na BrasilAPI e confirmados pelo usuário)
 * e os salva na planilha, acionando a lógica principal de criação.
 */
function _api_finalizarCadastroFornecedor(params) {
  try {
    const fornecedorData = params.fornecedorData; // Espera receber o objeto com os dados
    if (!fornecedorData || !fornecedorData.cnpj) {
      throw new Error("Dados do fornecedor para finalização estão incompletos.");
    }
    
    // Busca os valores padrão da planilha 'Config'
    const condicaoPadrao = getCondicoesPagamento()[0] || 'A PRAZO';
    const formaPadrao = getFormasPagamento()[0] || 'BOLETO';
    
    // Monta o objeto final no formato que a sua função principal espera
    const fornecedorParaSalvar = {
      razaoSocial: fornecedorData.razaoSocial,
      nomeFantasia: fornecedorData.nomeFantasia,
      cnpj: fornecedorData.cnpj,
      endereco: fornecedorData.endereco,
      cidade: fornecedorData.cidade,
      estado: fornecedorData.uf,
      condicaoPagamento: fornecedorData.condicaoPagamento || 'A PRAZO', // Usa o padrão
      formaPagamento: fornecedorData.formaPagamento || 'BOLETO',       // Usa o padrão
      grupo: '',
      regimeTributario: fornecedorData.regimeTributario,
    };

    // Chama sua função de salvamento principal e robusta que já existe!
    return adicionarOuAtualizarFornecedorv2(fornecedorParaSalvar);

  } catch (e) {
    Logger.log(`ERRO em _api_finalizarCadastroFornecedor: ${e.stack}`);
    return { status: "error", message: e.message };
  }
}

/**
 * Consulta o CNPJ na BrasilAPI E busca as opções de pagamento na planilha 'Config'.
 * Retorna todos os dados necessários para o bot iniciar a conversa de cadastro.
 */
function _api_consultarCnpjEopcoes(params) {
  try {
    const cnpj = params.cnpj
    Logger.log(`_api_consultarCnpjEopcoes recebeu: ${cnpj} (length: ${cnpj.length})`);
    // Reutiliza a função de consulta de CNPJ
    const resultadoCnpj = consultarCnpj_V2(cnpj);
    if (resultadoCnpj.status !== 'ok') {
      throw new Error(resultadoCnpj.message);
    }
    
    // Reutiliza suas funções para buscar as listas de opções
    const condicoes = getCondicoesPagamento();
    const formas = getFormasPagamento();
    
    return { 
      status: "success", 
      data: {
        dadosFornecedor: resultadoCnpj.data, // Dados da BrasilAPI
        opcoesCondicoes: condicoes,          // Lista de Condições
        opcoesFormas: formas                 // Lista de Formas
      } 
    };
  } catch (e) {
    Logger.log(`ERRO em _api_consultarCnpjEopcoes: ${e.stack}`);
    return { status: "error", message: e.message };
  }
}

/**
 * Função de teste dedicada para a consulta de um CNPJ na BrasilAPI.
 * Ela chama a função 'consultarCnpj_V2' para verificar a comunicação externa.
 */
function testar_consultaCnpj() {
  Logger.log("--- INICIANDO TESTE: consultarCnpj_V2 ---");

  // 1. Defina um CNPJ VÁLIDO para o teste.
  //    Pode ser com ou sem formatação.
  const cnpjDeTeste = "12.275.282/0001-19"; // <-- TROQUE SE QUISER TESTAR OUTRO CNPJ

  try {
    // 2. Chama a sua função de consulta de CNPJ
    const resultado = _api_consultarCnpjEopcoes(cnpjDeTeste);
    
    // 3. Imprime o resultado completo no log para análise
    Logger.log("--- RESULTADO DO TESTE ---");
    Logger.log(JSON.stringify(resultado, null, 2)); // Formata o JSON para ser fácil de ler
    Logger.log("--- FIM DO TESTE ---");
    
  } catch (e) {
    Logger.log("!!! O TESTE FALHOU COM UM ERRO CRÍTICO !!!");
    Logger.log("Erro: " + e.message);
    Logger.log("Stack: " + e.stack);
  }
}

function buscarPedidoTelegram(params) {
   try {
    Logger.log("Iniciando busca otimizada com parâmetros: " + JSON.stringify(params));

    if (!params || !params.empresaId) {
      Logger.log("Busca interrompida: ID da empresa é obrigatório.");
      return { status: 'success', data: [] }; // Retorna sucesso com dados vazios
    }

    const sheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName('Pedidos');
    if (!sheet) { throw new Error("Aba 'Pedidos' não encontrada."); }

    const dataRows = sheet.getDataRange().getValues();
    const headers = dataRows.shift(); // Remove e armazena o cabeçalho

    // Mapeamento de colunas (lógica mantida, pois já é eficiente)
    const colunas = {
      numeroDoPedido: headers.findIndex(h => h.toUpperCase().includes('NÚMERO DO PEDIDO')),
      empresa: headers.findIndex(h => h.toUpperCase() === 'EMPRESA'),
      data: headers.findIndex(h => h.toUpperCase() === 'DATA'),
      fornecedor: headers.findIndex(h => h.toUpperCase() === 'FORNECEDOR'),
      placaVeiculo: headers.findIndex(h => h.toUpperCase().includes('PLACA')),
      veiculo: headers.findIndex(h => h.toUpperCase().includes('VEICULO')),
      observacoes: headers.findIndex(h => h.toUpperCase().includes('OBSERVACOES')),
      totalGeral: headers.findIndex(h => h.toUpperCase().includes('TOTAL GERAL') || h.toUpperCase() === 'VALOR'),
      status: headers.findIndex(h => h.toUpperCase() === 'STATUS'),
      itens: headers.findIndex(h => h.toUpperCase().includes('ITENS')),
      estado: headers.findIndex(h => h.toUpperCase().includes('ESTADO FORNECEDOR')),
      dataCriacao: headers.findIndex(h => h.toUpperCase() === 'DATA CRIACAO'),
      aliquota: headers.findIndex(h => h.toUpperCase().includes('ALIQUOTA IMPOSTO')),
      icmsSt: headers.findIndex(h => h.toUpperCase().includes('ICMS ST TOTAL')),
      usuarioCriador: headers.findIndex(h => h.toUpperCase().includes('USUARIO CRIADOR'))
    };

    // --- OTIMIZAÇÃO: Pré-cálculo dos filtros ---
    const empresaFiltro = String(params.empresaId).trim();
    const statusExcluidos = ['RASCUNHO', 'AGUARDANDO APROVACAO'];
    const termoPrincipal = params.mainSearch ? String(params.mainSearch).trim().toLowerCase() : null;
    const placaFiltro = params.plateSearch ? String(params.plateSearch).trim().toLowerCase() : null;
    const criadorFiltro = params.usuarioCriador ? String(params.usuarioCriador).trim().toLowerCase() : null;
    const dataInicio = params.dateStart ? new Date(params.dateStart + 'T00:00:00') : null;
    const dataFim = params.dateEnd ? new Date(params.dateEnd + 'T23:59:59') : null;

    // --- LÓGICA DE PRÉ-BUSCA (pedido oculto) ---
    // Integrada ao loop principal para evitar uma segunda varredura dos dados
    if (termoPrincipal && params.perfil !== 'admin') {
      const pedidoOculto = dataRows.find(row => 
        String(row[colunas.numeroDoPedido]).toLowerCase().trim().includes(termoPrincipal) &&
        String(row[colunas.empresa]).trim() === empresaFiltro
      );
      if (pedidoOculto) {
        const statusDoPedido = (pedidoOculto[colunas.status] || '').trim().toUpperCase();
        if (statusExcluidos.includes(statusDoPedido)) {
          const numeroDoPedidoEncontrado = pedidoOculto[colunas.numeroDoPedido];
          const mensagem = `O pedido #${numeroDoPedidoEncontrado} foi encontrado, mas está com o status "${pedidoOculto[colunas.status]}" e não pode ser exibido.`;
          Logger.log(`[backend] Pedido oculto encontrado: ${mensagem}`);
          return { status: 'found_but_hidden', message: mensagem };
        }
      }
    }

    const pedidosEncontrados = [];
    // --- OTIMIZAÇÃO: Loop único para filtrar e mapear ---
    for (const row of dataRows) {
      
      // Filtro 1: Empresa (o mais importante, executado primeiro)
      if (String(row[colunas.empresa]).trim() !== empresaFiltro) {
        continue; // Pula para a próxima linha
      }
      
      // Filtro 2: Status
      if (!params.bypassStatusFilter) {
        const statusDoPedido = (row[colunas.status] || '').trim().toUpperCase();
        if (statusDoPedido === '' || statusExcluidos.includes(statusDoPedido)) {
          continue;
        }
      }

      // Filtro 3: Termo Principal (Nº Pedido ou Fornecedor)
      if (termoPrincipal) {
        const numPedido = String(row[colunas.numeroDoPedido]).toLowerCase();
        const fornecedor = String(row[colunas.fornecedor] ?? '').toLowerCase();
        if (!numPedido.includes(termoPrincipal) && !fornecedor.includes(termoPrincipal)) {
          continue;
        }
      }

      // Filtro 4: Data
      if (dataInicio && dataFim) {
        const dataPedido = new Date(row[colunas.data]);
        if (dataPedido < dataInicio || dataPedido > dataFim) {
          continue;
        }
      }

      // Filtro 5: Placa
      if (placaFiltro && colunas.placaVeiculo !== -1) {
        if (String(row[colunas.placaVeiculo]).toLowerCase().trim() !== placaFiltro) {
          continue;
        }
      }

      // Filtro 6: Usuário Criador
      if (criadorFiltro && colunas.usuarioCriador !== -1) {
        if (String(row[colunas.usuarioCriador]).toLowerCase().trim() !== criadorFiltro) {
          continue;
        }
      }

      // --- Se passou por todos os filtros, mapeia o objeto ---
      const dataDoPedido = row[colunas.data];
      const dataCriacao = row[colunas.dataCriacao];
      const pedido = {
        numeroDoPedido: row[colunas.numeroDoPedido],
        empresaId: row[colunas.empresa],
        data: dataDoPedido instanceof Date ? Utilities.formatDate(dataDoPedido, "GMT-03:00", "yyyy-MM-dd'T'HH:mm:ss'Z'") : dataDoPedido,
        fornecedor: row[colunas.fornecedor],
        totalGeral: row[colunas.totalGeral],
        status: row[colunas.status],
        placa: row[colunas.placaVeiculo],
        veiculo: row[colunas.veiculo],
        observacoes: row[colunas.observacoes],
        estado: row[colunas.estado],
        dataCriacao: dataCriacao instanceof Date ? Utilities.formatDate(dataCriacao, "GMT-03:00", "yyyy-MM-dd'T'HH:mm:ss'Z'") : dataCriacao,
        aliquota: row[colunas.aliquota],
        icmsSt: row[colunas.icmsSt],
        usuarioCriador: row[colunas.usuarioCriador],
        itens: [] // Inicializa com array vazio
      };

      // Processamento dos itens (lógica mantida)
      const itensJSON = row[colunas.itens];
      if (colunas.itens !== -1 && itensJSON && String(itensJSON).trim() !== '') {
        try {
          pedido.itens = JSON.parse(itensJSON);
        } catch (e) {
          Logger.log(`Erro ao parsear JSON de itens do pedido ${pedido.numeroDoPedido}: ` + e);
          pedido.erroItens = "Formato inválido";
        }
      }
      
      pedidosEncontrados.push(pedido);
    }
    
    Logger.log(`Busca finalizada. Encontrados ${pedidosEncontrados.length} pedidos.`);
    return { status: 'success', data: pedidosEncontrados };

  } catch (e) {
    Logger.log("Erro na função buscarPedidos_otimizada: " + e + "\nStack: " + e.stack);
    return { status: 'error', message: e.toString() };
  }
}

