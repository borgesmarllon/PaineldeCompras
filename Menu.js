// =================================================================
// CONFIGURAÇÕES GLOBAIS - AJUSTE ESTAS VARIÁVEIS
// =================================================================
const ID_DA_PLANILHA = "1M0GTX9WmnggiNnDynU0kC457yoy0iRHcRJ39d_B109o"; // <-- IMPORTANTE: Troque por seu ID real
const NOME_DA_ABA_DE_PEDIDOS = "Pedidos"; // <-- Ajuste se o nome da sua aba for diferente
const NOME_DA_ABA_DE_AVISOS = "Avisos";   // <-- Ajuste se o nome da sua aba for diferente


/**
 * Busca na planilha os pedidos que estão com status "Pendente de Aprovação".
 * @returns {Array<Object>} Um array de objetos, onde cada objeto é um pedido pendente.
 */
function getPendingApprovals() {
    try {
      const sheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName(NOME_DA_ABA_DE_PEDIDOS);
      if (!sheet) { throw new Error(`Aba "${NOME_DA_ABA_DE_PEDIDOS}" não encontrada.`); }

      const data = sheet.getDataRange().getValues();
      const headers = data.shift();

      const colunas = {
        numeroDoPedido: headers.indexOf("Número do Pedido"),
        fornecedor: headers.indexOf("Fornecedor"),
        totalGeral: headers.indexOf("Total Geral"),
        status: headers.indexOf("Status"),
        empresaId: headers.indexOf("Empresa") // Só precisamos do ID
    };

    const pendingApprovals = [];
    data.forEach(row => {
      if (String(row[colunas.status]).trim().toUpperCase() === "AGUARDANDO APROVACAO") {
          pendingApprovals.push({
            numeroDoPedido: row[colunas.numeroDoPedido],
            fornecedor: row[colunas.fornecedor],
            totalGeral: row[colunas.totalGeral],
            empresaId: row[colunas.empresaId] // Retorna só o ID
          });
      }
    });
    return pendingApprovals;
  } catch (e) { 
    console.error(`Erro em getPendingApprovals: ${e.message}`);
    return []; // Retorna um array vazio em caso de erro para não quebrar o frontend
  }
}

/**
 * Busca na planilha os avisos que estão marcados como ativos para o Mural.
 * @returns {Array<Object>} Um array de objetos, onde cada objeto é um aviso.
 */
function getNotices() {
  try {
    const sheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName(NOME_DA_ABA_DE_AVISOS);
    if (!sheet) {
      throw new Error(`A aba "${NOME_DA_ABA_DE_AVISOS}" não foi encontrada.`);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    // Encontra o índice de cada coluna dinamicamente
    const colunas = {
      data: headers.indexOf("Data"),
      mensagem: headers.indexOf("Mensagem"),
      ativo: headers.indexOf("Ativo")
    };

    // Valida se as colunas foram encontradas
    for (let key in colunas) {
      if (colunas[key] === -1) {
        throw new Error(`A coluna "${headers[key] || key}" não foi encontrada na aba de avisos.`);
      }
    }

    const activeNotices = [];
    data.forEach(row => {
      // Verifica se a linha não está vazia e se o aviso está ativo (marcado como "SIM" ou TRUE)
      if (row[colunas.ativo] === "SIM" || row[colunas.ativo] === true) {
        // Formata a data para o padrão DD/MM/YYYY, caso seja um objeto Date
        const dataAviso = row[colunas.data] instanceof Date 
          ? Utilities.formatDate(row[colunas.data], "GMT-03:00", "dd/MM/yyyy") 
          : row[colunas.data];

        activeNotices.push({
          data: dataAviso,
          mensagem: row[colunas.mensagem]
        });
      }
    });
    
    console.log(`Encontrados ${activeNotices.length} avisos ativos.`);
    // Inverte o array para mostrar os avisos mais recentes (de baixo) primeiro
    return activeNotices.reverse(); 

  } catch (e) {
    console.error(`Erro em getNotices: ${e.message}`);
    return [];
  }
}

/**
 * Busca os 3 avisos mais recentes com status "Ativo".
 */
function getAvisosAtivos() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Avisos");
    if (!sheet || sheet.getLastRow() < 2) return [];

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0); // Zera a hora para comparar apenas a data

    const allAvisos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues(); // Lê 5 colunas
    return allAvisos
      .filter(row => {
          const status = String(row[0]).trim().toUpperCase() === 'ATIVO';
          const dataVencimento = row[4]; // Coluna 5 (índice 4) é a data de vencimento
          const naoExpirado = !dataVencimento || new Date(dataVencimento) >= hoje;
          return status && naoExpirado;
      })
      .map(row => ({
        status: row[0], data: row[1] instanceof Date ? row[1].toISOString() : new Date().toISOString(),
        titulo: row[2], mensagem: row[3]
      }))
      .sort((a, b) => new Date(b.data) - new Date(a.data))
      .slice(0, 3);
  } catch (e) {
    Logger.log(`ERRO em getAvisosAtivos: ${e.message}`);
    return [];
  }
}

/**
 * Busca TODOS os avisos para a tela de gerenciamento.
 */
function getTodosOsAvisos() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Avisos");
    if (!sheet || sheet.getLastRow() < 2) return [];

    const allAvisos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    return allAvisos
      .map((row, index) => ({
        row: index + 2, // Guarda o número da linha real para edições futuras
        status: String(row[0]).trim().toUpperCase(),
        data: row[1] instanceof Date ? row[1].toISOString() : new Date().toISOString(),
        titulo: row[2],
        mensagem: row[3],
        vencimento: row[4] instanceof Date ? row[4].toISOString() : null
      }))
      .sort((a, b) => new Date(b.data) - new Date(a.data));
  } catch (e) {
    Logger.log(`ERRO em getTodosOsAvisos: ${e.message}`);
    return [];
  }
}

/**
 * Busca os dados de um único aviso pela sua linha na planilha.
 */
function getAvisoByRow(rowNumber) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Avisos");
    if (!sheet) throw new Error("Planilha 'Avisos' não encontrada.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Mapeia os dados para um objeto usando os cabeçalhos como chaves
    const avisoData = {};
    headers.forEach((header, index) => {
      avisoData[header] = rowData[index];
    });

    // --- LÓGICA DE DEPURAÇÃO E CORREÇÃO DA DATA ---
    let dataVencimentoISO = null;
    const vencimentoRaw = avisoData["Data de Vencimento"];
    Logger.log(`[Debug getAvisoByRow] Raw 'Data de Vencimento' da planilha: ${vencimentoRaw} (Tipo: ${typeof vencimentoRaw})`);

    if (vencimentoRaw && String(vencimentoRaw).trim() !== '') {
        // Tenta converter para um objeto Date, não importa o formato original
        const tempDate = new Date(vencimentoRaw);
        // Verifica se a conversão resultou em uma data válida
        if (!isNaN(tempDate.getTime())) {
            dataVencimentoISO = tempDate.toISOString();
            Logger.log(`[Debug getAvisoByRow] Data de vencimento convertida com sucesso para: ${dataVencimentoISO}`);
        } else {
            Logger.log(`[Debug getAvisoByRow] AVISO: Não foi possível converter '${vencimentoRaw}' para uma data válida.`);
        }
    }
    // --- FIM DA LÓGICA DE DEPURAÇÃO ---

    // Monta o objeto de retorno de forma segura, usando os nomes exatos das colunas
    const aviso = {
      row: rowNumber,
      status: String(avisoData["Status"] || '').trim().toUpperCase(),
      data: avisoData["Data"] instanceof Date ? avisoData["Data"].toISOString() : null,
      titulo: avisoData["Título"],
      mensagem: avisoData["Mensagem"],
      vencimento: dataVencimentoISO // Usa a data processada e validada
    };

    return { status: 'ok', data: aviso };
  } catch (e) {
    Logger.log(`ERRO em getAvisoByRow: ${e.message}`);
    return { status: 'error', message: e.message };
  }
}

/**
 * Adiciona um novo aviso à planilha.
 */
function adicionarNovoAviso(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Avisos");
    // Colunas: Status, Data, Título, Mensagem, Data de Vencimento
    const dataVencimento = data.vencimento ? new Date(data.vencimento) : null;
    sheet.appendRow(['Ativo', new Date(), data.titulo, data.mensagem, dataVencimento]);
    return { status: 'ok' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * Edita o título e a mensagem de um aviso existente.
 */
function editarAviso(data) {
    try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Avisos");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Encontra a posição de cada coluna
    const colData = headers.indexOf("Data") + 1;
    const colTitulo = headers.indexOf("Título") + 1;
    const colMensagem = headers.indexOf("Mensagem") + 1;
    const colVencimento = headers.indexOf("Data de Vencimento") + 1;

    // Valida se as colunas foram encontradas
    if ([colData, colTitulo, colMensagem, colVencimento].includes(0)) {
        throw new Error("Uma ou mais colunas necessárias (Data, Título, Mensagem, Data de Vencimento) não foram encontradas.");
    }

    // ===== CORREÇÃO APLICADA AQUI =====
    // Trata a data de vencimento de forma mais robusta para evitar problemas de fuso horário.
    // Se 'data.vencimento' for uma string como "2025-07-30", isso a converte para uma data local.
    const dataVencimento = data.vencimento ? new Date(data.vencimento + 'T00:00:00') : null;
    Logger.log(`Backend: Tentando salvar a data de vencimento como: ${dataVencimento}`);
    
    // Atualiza os valores nas colunas corretas
    sheet.getRange(data.row, colData).setValue(new Date());
    sheet.getRange(data.row, colTitulo).setValue(data.titulo);
    sheet.getRange(data.row, colMensagem).setValue(data.mensagem);
    sheet.getRange(data.row, colVencimento).setValue(dataVencimento);
    
    Logger.log(`Aviso na linha ${data.row} atualizado com sucesso.`);
    return { status: 'ok', message: 'Aviso atualizado com sucesso!' };
  } catch (e) {
    Logger.log(`ERRO em editarAviso: ${e.message}`);
    return { status: 'error', message: e.message };
  }
}

/**
 * Altera o status de um aviso (Ativo/Inativo).
 */
function alterarStatusAviso(rowNumber, novoStatus) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Avisos");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Encontra a posição das colunas de Status e Data
    const colStatus = headers.indexOf("Status") + 1;
    const colData = headers.indexOf("Data") + 1;

    // Valida se as colunas foram encontradas
    if (colStatus === 0) {
        throw new Error("A coluna 'Status' não foi encontrada.");
    }
    if (colData === 0) {
        throw new Error("A coluna 'Data' não foi encontrada.");
    }

    // Atualiza o status e a data da modificação
    sheet.getRange(rowNumber, colStatus).setValue(novoStatus);
    sheet.getRange(rowNumber, colData).setValue(new Date());
    
    return { status: 'ok' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * Encontra um pedido pelo seu número e atualiza seu status.
 * @param {string} numeroPedido O número do pedido a ser atualizado.
 * @param {string} novoStatus O novo status a ser definido (ex: "Aprovado", "Rejeitado").
 * @returns {object} Um objeto indicando o sucesso ou falha da operação.
 */
function atualizarStatusPedido(numeroPedido, novoStatus, motivo, infoAprovador) {
  try {
    // Validação básica de entrada
    if (!numeroPedido || !novoStatus) {
      throw new Error("Número do pedido e novo status são obrigatórios.");
    }

    const sheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName(NOME_DA_ABA_DE_PEDIDOS);
    if (!sheet) { throw new Error(`Aba "${NOME_DA_ABA_DE_PEDIDOS}" não encontrada.`); }

    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // Pega apenas os cabeçalhos

    const colunaNumeroPedido = headers.indexOf("Número do Pedido");
    const colunaStatus = headers.indexOf("Status");
    const colunaMotivo = headers.indexOf("Motivo Rejeição");
    const colunaNotificacao = headers.indexOf("NotificacaoAprovadoVisto");
    const colCriadorUsername = headers.indexOf("Usuario Criador");

    if (colCriadorUsername === -1) {
      throw new Error("Coluna 'Usuário Criador' não foi encontrada na planilha de Pedidos.");
    }

    if (colunaNumeroPedido === -1 || colunaStatus === -1) {
      throw new Error("Não foi possível encontrar as colunas 'Número do Pedido' ou 'Status' na planilha.");
    }
    if (novoStatus === 'Rejeitado' && colunaMotivo === -1) {
        throw new Error("A coluna 'Motivo Rejeição' é necessária para rejeitar um pedido, mas não foi encontrada.");
    }
    if (novoStatus === 'Aprovado' && colunaNotificacao === -1) {
      throw new Error("A coluna 'NotificacaoAprovadoVisto' é necessária para notificar o pedido.");
    }

    // Encontra a linha correspondente ao pedido (começando da linha 2 da planilha, que é o índice 1 nos dados)
    for (let i = 1; i < data.length; i++) {
      const cellPedido = String(data[i][colunaNumeroPedido]).replace(/^'/, '').trim();
      if (cellPedido === String(numeroPedido).trim()) {
        const rowIndex = i + 1; // O número da linha na planilha é o índice + 1
         const linhaDoPedido = data[i];
        
        // Atualiza a célula do status na linha encontrada
        sheet.getRange(rowIndex, colunaStatus + 1).setValue(novoStatus);

        // Se for uma rejeição, salva o motivo. Se for aprovação, limpa o motivo.
        if (colunaMotivo !== -1) {
            sheet.getRange(rowIndex, colunaMotivo + 1).setValue(motivo);
        }
        
        // Se for aprovação, inicia o contador de notificação
        if (novoStatus === 'Aprovado') {
          Logger.log(`[Depuração] Pedido #${numeroPedido} APROVADO. A tentar definir 'NotificacaoAprovadoVisto' como 0 na linha ${rowIndex}, coluna ${colunaNotificacao + 1}.`);
          sheet.getRange(rowIndex, colunaNotificacao + 1).setValue(0);
        }

        // 1. Pega o NOME DE USUÁRIO do criador do pedido
                const usernameDoCriador = linhaDoPedido[colCriadorUsername];
                
                // 2. Usa o NOME DE USUÁRIO para encontrar o Chat ID do Telegram
                const chatIdDoCriador = _getChatIdPorUsername(usernameDoCriador);
                
                let mensagem = "";
                if (novoStatus === 'Aprovado') {
                    mensagem = `✅ <b>Pedido Aprovado!</b>\n\nSeu pedido <b>Nº ${numeroPedido}</b> foi aprovado.`;
                } else if (novoStatus === 'Rejeitado') {
                    mensagem = `❌ <b>Pedido Rejeitado.</b>\n\nSeu pedido <b>Nº ${numeroPedido}</b> foi rejeitado.\n<b>Motivo:</b> <i>${motivo || 'N/A'}</i>`;
                }

                // 3. Envia a notificação para o Chat ID encontrado
                if (mensagem && chatIdDoCriador) {
                    enviarNotificacaoTelegram(chatIdDoCriador, mensagem);
                }


        Logger.log(`Pedido #${numeroPedido} atualizado para o status "${novoStatus}" na linha ${rowIndex}.`);
        return { status: 'success', message: `Pedido ${novoStatus.toLowerCase()} com sucesso!`, criador_chat_id: chatIdDoCriador };
      }
    }

    // Se o loop terminar e não encontrar o pedido
    throw new Error(`Pedido #${numeroPedido} não encontrado.`);

  } catch (e) {
    console.error(`Erro em atualizarStatusPedido: ${e.message}`);
    return { status: 'error', message: e.message };
  }
}

/**
 * ✅ NOVA FUNÇÃO: Busca os pedidos aprovados recentemente para o utilizador
 * e incrementa o contador de visualização.
 * @param {string} usuarioLogado O nome do utilizador logado.
 * @returns {Array<Object>} Um array de objetos de pedidos aprovados.
 */
function getMinhasAprovacoesRecentes(usuarioLogado) {
  if (!usuarioLogado) return [];
  
  try {
    const sheetPedidos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("pedidos");
    if (!sheetPedidos) return [];

    const data = sheetPedidos.getDataRange().getValues();
    const headers = data.shift();
    const idxNumPed = headers.indexOf('Número do Pedido');
    const idxStatus = headers.indexOf('Status');
    const idxCriador = headers.indexOf('Usuario Criador');
    const idxNotificacao = headers.indexOf('NotificacaoAprovadoVisto');

    if (idxNotificacao === -1) {
        console.log("Aviso: Coluna 'NotificacaoAprovadoVisto' não encontrada. A funcionalidade de avisos de aprovação está desativada.");
        return [];
    }

    const notificacoes = [];
    const usuarioLogadoKey = usuarioLogado.trim().toLowerCase();

    data.forEach((row, index) => {
        const contador = row[idxNotificacao];
        const status = row[idxStatus];
        const criador = String(row[idxCriador]).trim().toLowerCase();

        if (status === 'Aprovado' && criador === usuarioLogadoKey) {
            Logger.log(`[Depuração] A verificar Pedido #${row[idxNumPed]} para ${criador}. Valor do contador: "${contador}" (Tipo: ${typeof contador})`);
            
            if (typeof contador === 'number' && contador < 2) {
                Logger.log(`--> CONDIÇÃO CUMPRIDA. A adicionar à lista de notificações.`);
                notificacoes.push({
                    numeroPedido: row[idxNumPed]
                });
                
                const novoValor = contador + 1;
                Logger.log(`--> A incrementar contador para ${novoValor} na linha ${index + 2}.`);
                sheetPedidos.getRange(index + 2, idxNotificacao + 1).setValue(novoValor);
            }
        }
    });
    
    return notificacoes;

  } catch (e) {
    Logger.log(`Erro em getMinhasAprovacoesRecentes: ${e.message}`);
    return [];
  }
}

/**
 * Busca e processa os dados para o dashboard, usando um filtro mais robusto.
 */
function getDashboardAdminData() {
  try {
    const ss = SpreadsheetApp.openById(ID_DA_PLANILHA);
    const empresasSheet = ss.getSheetByName('Empresas');
    const pedidosSheet = ss.getSheetByName(NOME_DA_ABA_DE_PEDIDOS);

    if (!empresasSheet || !pedidosSheet) {
      throw new Error("Planilhas 'Empresas' ou 'Pedidos' não encontradas.");
    }
    
    // Obter empresas
    const empresasData = empresasSheet.getRange(2, 1, empresasSheet.getLastRow() - 1, 2).getValues();
    const todasAsEmpresas = empresasData.map(([id, nome]) => ({ id: String(id).trim(), nome: nome, ultimosPedidosAprovados: [], ultimosPedidosCancelados: [] }));

    // Obter TODOS os pedidos e normalizar cabeçalhos
    const pedidosData = pedidosSheet.getDataRange().getValues();
    const pedidosHeaders = pedidosData.shift().map(h => String(h || '').toUpperCase().trim());
    Logger.log("Cabeçalhos encontrados: " + JSON.stringify(pedidosHeaders));
    const colunas = {
      empresaId: pedidosHeaders.indexOf("EMPRESA"),
      status: pedidosHeaders.indexOf("STATUS"),
      dataCriacao: pedidosHeaders.indexOf("DATA CRIACAO"),
      data: pedidosHeaders.indexOf("DATA"),
      numeroDoPedido: pedidosHeaders.indexOf("NÚMERO DO PEDIDO"),
      fornecedor: pedidosHeaders.indexOf("FORNECEDOR"),
      totalGeral: pedidosHeaders.indexOf("TOTAL GERAL"),
      estadoFornecedor: pedidosHeaders.indexOf("ESTADO FORNECEDOR"),
      icmsStTotal: pedidosHeaders.indexOf("ICMS ST TOTAL"),
      usuarioCancelamento: pedidosHeaders.indexOf("USUARIO CANCELAMENTO")
    };
    Logger.log("Índice da coluna ICMS ST Total: " + colunas.icmsStTotal);
    if (colunas.status === -1) throw new Error("A coluna 'STATUS' não foi encontrada.");
    if (colunas.empresaId === -1) throw new Error("A coluna 'EMPRESA' não foi encontrada.");
    if (colunas.estadoFornecedor === -1) throw new Error("A coluna 'ESTADO FORNECEDOR' não foi encontrada.");

    // Agrupar pedidos APROVADOS
    const pedidosAprovadosPorEmpresa = {};
    const pedidosCanceladosPorEmpresa = {};
    pedidosData.forEach(row => {
      // Lógica de filtro idêntica à da função de teste que funcionou
      const status = String(row[colunas.status] || '').trim().toUpperCase();
      const empresaId = String(row[colunas.empresaId] || '').trim(); 
      Logger.log(`Pedido #${row[colunas.numeroDoPedido]} - ICMS ST: ${row[colunas.icmsStTotal]}`);
      if (!empresaId) return;

      const dataObj = row[colunas.dataCriacao] || row[colunas.data];
      const dataFormatada = (dataObj instanceof Date) ? dataObj.toISOString() : dataObj;
      if (status === "APROVADO" && empresaId) {
        if (!pedidosAprovadosPorEmpresa[empresaId]) {
          pedidosAprovadosPorEmpresa[empresaId] = [];
        }
        pedidosAprovadosPorEmpresa[empresaId].push({
          numeroDoPedido: row[colunas.numeroDoPedido],
          fornecedor: row[colunas.fornecedor],
          totalGeral: row[colunas.totalGeral],
          data: dataFormatada,
          icmsStTotal: row[colunas.icmsStTotal],         
          estadoFornecedor: row[colunas.estadoFornecedor]
        })
      }
      else if (status === "CANCELADO") {
        if (!pedidosCanceladosPorEmpresa[empresaId]) {
          pedidosCanceladosPorEmpresa[empresaId] = [];
        }
        pedidosCanceladosPorEmpresa[empresaId].push({
          numeroDoPedido: row[colunas.numeroDoPedido],
          usuarioCancelamento: row[colunas.usuarioCancelamento] || '',
          data: dataFormatada
        })
      }
    });
    Logger.log(`📊 Empresas com pedidos aprovados: ${Object.keys(pedidosAprovadosPorEmpresa).length}`);
    
    
    // Ordenar, fatiar e juntar os dados
    todasAsEmpresas.forEach(empresa => {
      const aprovados = pedidosAprovadosPorEmpresa[empresa.id];
      if (aprovados && aprovados.length > 0) {
        aprovados.sort((a, b) => new Date(b.data) - new Date(a.data));
        empresa.ultimosPedidosAprovados = aprovados.slice(0, 2);
      }
      // --- NOVO: Processa os cancelados ---
      const cancelados = pedidosCanceladosPorEmpresa[empresa.id];
      if (cancelados && cancelados.length > 0) {
        cancelados.sort((a, b) => new Date(b.data) - new Date(a.data));
        empresa.ultimosPedidosCancelados = cancelados.slice(0, 1);
      }
    });

    return { status: 'ok', data: todasAsEmpresas };

  } catch (e) {
    Logger.log(`ERRO em getDashboardAdminData: ${e.message}`);
    return { status: 'error', message: e.message, data: [] };
  }
}

function _TESTE_verificarDadosDoDashboardAdmin() {
  try {
    Logger.log("--- 🚀 INICIANDO TESTE DO BACK-END DO DASHBOARD ADMIN ---");

    // 1. Chama a sua função principal de busca de dados
    const resultado = getDashboardAdminData();

    // 2. Verifica se a chamada foi bem-sucedida
    if (resultado && resultado.status === 'ok') {
      Logger.log("✅ A função foi executada com sucesso!");
      
      const todasAsEmpresas = resultado.data;
      Logger.log(`Encontradas informações para ${todasAsEmpresas.length} empresa(s).`);
      Logger.log("--- INSPECIONANDO DADOS DE CADA EMPRESA ---");

      // 3. Itera sobre cada empresa e mostra os dados encontrados
      todasAsEmpresas.forEach(empresa => {
        Logger.log(`\n🏢 Empresa: ${empresa.nome} (ID: ${empresa.id})`);

        // Verifica e loga os pedidos APROVADOS
        if (empresa.ultimosPedidosAprovados.length > 0) {
          Logger.log(`   -> Encontrados ${empresa.ultimosPedidosAprovados.length} pedidos APROVADOS.`);
          Logger.log(JSON.stringify(empresa.ultimosPedidosAprovados, null, 2));
        } else {
          Logger.log("   -> Nenhum pedido APROVADO recente encontrado.");
        }

        // Verifica e loga os pedidos CANCELADOS
        if (empresa.ultimosPedidosCancelados.length > 0) {
          Logger.log(`   -> Encontrados ${empresa.ultimosPedidosCancelados.length} pedidos CANCELADOS.`);
          Logger.log(JSON.stringify(empresa.ultimosPedidosCancelados, null, 2));
        } else {
          Logger.log("   -> Nenhum pedido CANCELADO recente encontrado.");
        }
      });

    } else {
      // Se a função retornou um erro, mostra a mensagem
      throw new Error(resultado ? resultado.message : "A função não retornou um resultado válido.");
    }

    Logger.log("\n--- ✅ TESTE DO BACK-END CONCLUÍDO ---");

  } catch (e) {
    Logger.log(`🔥🔥 FALHA NO TESTE DO BACK-END: ${e.message}`);
  }
}

// =========================================================================
// NOVA FUNÇÃO - EXCLUSIVA PARA OS CARDS DO MENU PRINCIPAL
// =========================================================================
function buscarUltimosPedidosDoUsuario(params) {
    // 1. Pega o nome de usuário DIRETAMENTE dos parâmetros enviados pelo cliente.
    const nomeDoUsuario = params.usuarioCriador;
    const idDaEmpresa = params.empresaId;

    // 2. Validações básicas
    if (!nomeDoUsuario || !idDaEmpresa) {
        return { status: 'error', message: 'Informações insuficientes (usuário ou empresa) para a busca.' };
    }

    try {
        Logger.log(`[buscarUltimosPedidosDoUsuario] Buscando últimos pedidos para: ${nomeDoUsuario}`);

        // 3. Sua lógica para acessar a planilha continua a mesma
        const sheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName('Pedidos');
        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const colunas = mapearColunas(headers);
         Logger.log("Colunas mapeadas: " + JSON.stringify(colunas, null, 2));

        // 4. O filtro agora confia no nome de usuário enviado
        const pedidosDoUsuario = data
            .filter(row => {
                const criadorPlanilha = String(row[colunas.usuarioCriador] || '').trim().toLowerCase();
                const empresaPlanilha = String(row[colunas.empresa]).trim();
                
                // O filtro duplo: tem que ser deste usuário E desta empresa
                return criadorPlanilha === nomeDoUsuario.trim().toLowerCase() && empresaPlanilha === idDaEmpresa;
            })
            .sort((a, b) => { // Ordena para garantir que pegamos os mais recentes
                let valorDataA = a[colunas.dataCriacao] || a[colunas.data];
                let valorDataB = b[colunas.dataCriacao] || b[colunas.data];
                
                const dataA = valorDataA instanceof Date ? valorDataA : new Date(String(valorDataA).replace(' ', 'T'));
                const dataB = valorDataB instanceof Date ? valorDataB : new Date(String(valorDataB).replace(' ', 'T'));
                
                return dataB - dataA;
            });
        
        // 5. Pega os 2 últimos
        const ultimosDoisPedidos = pedidosDoUsuario.slice(0, 2);

        // 6. Mapeia para o formato de objeto que seu cliente espera
        const resultadoFinal = ultimosDoisPedidos.map(row => {
        const dataDoPedido = row[colunas.data];
        const dataCriacaoDoPedido = row[colunas.dataCriacao];

        const pedido ={
        numeroDoPedido: row[colunas.numeroDoPedido],
        empresaId: row[colunas.empresa],
        data: dataDoPedido instanceof Date ? Utilities.formatDate(dataDoPedido, "GMT-03:00", "yyyy-MM-dd'T'HH:mm:ss'Z'") : dataDoPedido,
        fornecedor: row[colunas.fornecedor],
        totalGeral: row[colunas.totalGeral],
        status: row[colunas.status],
        placa: row[colunas.placaVeiculo],
        veiculo: row[colunas.veiculo],
        observacoes: row[colunas.observacoes],
        //itens: row[colunas.itens]
        estadoFornecedor: row[colunas.estadoFornecedor], 
        icmsStTotal: row[colunas.icmsStTotal],      
        dataCriacao: dataCriacaoDoPedido instanceof Date ? Utilities.formatDate(dataCriacaoDoPedido, "GMT-03:00", "yyyy-MM-dd HH:mm:ss") : dataCriacaoDoPedido,
        aliquota: row[colunas.aliquota],
        usuarioCriador: row[colunas.usuarioCriador]
      };
           // Lógica para adicionar os itens, igual à sua outra função
            const itensJSON = row[colunas.itens];
            if (colunas.itens !== -1 && itensJSON && String(itensJSON).trim() !== '') {
                try {
                    pedido.itens = JSON.parse(itensJSON);
                } catch (e) {
                    Logger.log(`Erro ao parsear JSON de itens do pedido ${pedido.numeroDoPedido}: ` + e);
                    pedido.erroItens = "Formato inválido";
                }
            }
            
            // AGORA RETORNAMOS O OBJETO CORRETO E COMPLETO
            return pedido; 
        });
        
        return { status: 'success', data: resultadoFinal };

    } catch (e) {
        Logger.log("Erro em buscarUltimosPedidosDoUsuario: " + e.stack);
        return { status: 'error', message: e.toString() };
    }
}

/**
 * BUSCA OS PEDIDOS REJEITADOS CRIADOS PELO USUÁRIO ATUAL.
 * @returns {Array<Object>} Um array de objetos, cada um com {numeroPedido, motivoRejeicao}.
 */
function getMeusPedidosRejeitados(usuarioLogado) {
  try {
    // Validação para garantir que o usuário foi passado como parâmetro.
    if (!usuarioLogado) {
        Logger.log("Aviso em getMeusPedidosRejeitados: Nenhum usuário logado fornecido.");
        return [];
    }

    const usuarioLogadoKey = usuarioLogado.trim().toLowerCase();
    Logger.log(`Buscando pedidos rejeitados para o utilizador: "${usuarioLogadoKey}"`);

    const sheetPedidos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("pedidos");
    if (!sheetPedidos) throw new Error("Aba 'pedidos' não encontrada.");

    const data = sheetPedidos.getDataRange().getValues();
    const headers = data.shift();
    const idxNumPed = headers.indexOf('Número do Pedido');
    const idxStatus = headers.indexOf('Status');
    const idxCriador = headers.indexOf('Usuario Criador');
    const idxMotivo = headers.indexOf('Motivo Rejeição'); // Garanta que esta coluna exista

    const pedidosRejeitados = [];

    data.forEach((row, index) => {
        const status = row[idxStatus];
        const criador = row[idxCriador];
        
        // Compara os valores padronizados
        if (status === 'Rejeitado' && String(criador).trim().toLowerCase() === usuarioLogadoKey) {
            Logger.log(`Encontrado pedido rejeitado na linha ${index + 2} para o utilizador ${criador}`);
            pedidosRejeitados.push({
                numeroPedido: row[idxNumPed],
                motivoRejeicao: row[idxMotivo] || "Nenhum motivo fornecido."
            });
        }
    });
    
    Logger.log(`Finalizado. Total de ${pedidosRejeitados.length} pedidos rejeitados encontrados para "${usuarioLogadoKey}".`);
    return pedidosRejeitados;

  } catch (e) {
    Logger.log(`Erro em getMeusPedidosRejeitados: ${e.message}`);
    return [];
  }
}


/**
 * BUSCA TODOS OS DADOS DE UM PEDIDO ESPECÍFICO PARA CORREÇÃO.
 * @param {string} numeroPedido O número do pedido a ser buscado.
 * @returns {Object} Um objeto completo com os dados do cabeçalho e a lista de itens do pedido.
 */
function getDadosCompletosDoPedido(numeroPedido) {
   try {
    Logger.log(`Iniciando busca de dados completos para o Pedido #${numeroPedido}...`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPedidos = ss.getSheetByName("pedidos");
    const sheetFornecedores = ss.getSheetByName("fornecedores");

    // 1. Busca os dados do cabeçalho do pedido
    const pedidosData = sheetPedidos.getDataRange().getValues();
    const headersPed = pedidosData.shift();
    const rowPedido = pedidosData.find(row => String(row[headersPed.indexOf('Número do Pedido')]) === String(numeroPedido));
    if (!rowPedido) throw new Error(`Pedido ${numeroPedido} não encontrado.`);
    
    const dadosPedidoRaw = getObjectFromRow(rowPedido, headersPed);

    // 2. Busca o ID do fornecedor correspondente
    const nomeFornecedorDoPedido = dadosPedidoRaw['Fornecedor'];
    let idFornecedor = null;
    if (nomeFornecedorDoPedido) {
        const fornecedoresData = sheetFornecedores.getDataRange().getValues();
        const headersForn = fornecedoresData.shift();
        const headersFornLowerCase = headersForn.map(h => String(h).toLowerCase());
        const idxIdForn = headersFornLowerCase.indexOf('id');
        const idxNomeForn = headersFornLowerCase.indexOf('razao social');

        if (idxIdForn > -1 && idxNomeForn > -1) {
            const nomeFornecedorKey = String(nomeFornecedorDoPedido).trim().toLowerCase();
            const fornecedorEncontrado = fornecedoresData.find(row => String(row[idxNomeForn]).trim().toLowerCase() === nomeFornecedorKey);
            if (fornecedorEncontrado) {
              idFornecedor = fornecedorEncontrado[idxIdForn];
            }
        }
    }

    // 3. ✅ CORREÇÃO: Lê e converte a string JSON de 'Itens' para um array de objetos.
    let itensDoPedido = [];
    const itensJsonString = dadosPedidoRaw['Itens']; // A coluna com 'I' maiúsculo
    if (itensJsonString && typeof itensJsonString === 'string') {
        try {
            itensDoPedido = JSON.parse(itensJsonString);
            Logger.log(`${itensDoPedido.length} itens encontrados e convertidos a partir do JSON do pedido.`);
        } catch (e) {
            Logger.log(`Erro ao converter o JSON de itens para o pedido #${numeroPedido}: ${e.message}`);
        }
    } else {
        Logger.log(`Aviso: A coluna 'Itens' para o pedido #${numeroPedido} está vazia ou não é uma string.`);
    }

    // 4. Cria um objeto de retorno limpo e padronizado
    const dadosLimpos = {
        'Número do Pedido': dadosPedidoRaw['Número do Pedido'],
        'Data': dadosPedidoRaw['Data'],
        'Observacoes': dadosPedidoRaw['Observacoes'],
        'Nome Veiculo': dadosPedidoRaw['Nome Veiculo'],
        'Placa Veiculo': dadosPedidoRaw['Placa Veiculo'],
        'ID_FORNECEDOR': idFornecedor,
        'itens': itensDoPedido // Agora contém o array de itens convertido
    };
    
    Logger.log("Retornando dados limpos para o frontend: %s", JSON.stringify(dadosLimpos, null, 2));
    return dadosLimpos;

  } catch (e) {
    Logger.log(`Erro em getDadosCompletosDoPedido: ${e.message}`);
    return null;
  }
}

/**
 * ATUALIZA UM PEDIDO EXISTENTE APÓS CORREÇÃO E O REENVIA PARA APROVAÇÃO.
 * @param {Object} dadosPedido O objeto completo do pedido com os dados atualizados.
 * @returns {Object} Um objeto de status.
 */
function reenviarPedidoCorrigido(dadosPedido) {
  Logger.log("--- INICIANDO reenviarPedidoCorrigido ---");
  Logger.log("Dados recebidos do frontend: %s", JSON.stringify(dadosPedido, null, 2));

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPedidos = ss.getSheetByName("pedidos");
    const sheetItens = ss.getSheetByName("itens pedido");
    const sheetFornecedores = ss.getSheetByName("fornecedores");

    // 1. Encontra a linha para atualizar
    const pedidosData = sheetPedidos.getDataRange().getValues();
    const headersPed = pedidosData.shift();
    const idxNumPed = headersPed.indexOf('Número do Pedido');
    
    const rowIndexInData = pedidosData.findIndex(row => row[idxNumPed] == dadosPedido.numeroPedido);
    if (rowIndexInData === -1) {
      throw new Error("Não foi possível encontrar o pedido original para atualizar.");
    }
    
    const sheetRowIndex = rowIndexInData + 2;
    const oldRowData = sheetPedidos.getRange(sheetRowIndex, 1, 1, headersPed.length).getValues()[0];
    Logger.log(`Pedido #${dadosPedido.numeroPedido} encontrado na linha ${sheetRowIndex}. Dados antigos: ${oldRowData}`);

    // 2. Procura o nome do fornecedor a partir do ID recebido
    const fornecedoresData = sheetFornecedores.getDataRange().getValues();
    const headersForn = fornecedoresData.shift();
    const idxIdForn = headersForn.indexOf('ID');
    const idxRazaoSocial = headersForn.indexOf('RAZAO SOCIAL');
    
    let nomeFornecedor = '';
    const fornecedorEncontrado = fornecedoresData.find(row => String(row[idxIdForn]) === String(dadosPedido.fornecedorId));
    if (fornecedorEncontrado) {
      nomeFornecedor = fornecedorEncontrado[idxRazaoSocial];
      Logger.log(`Fornecedor encontrado. ID: ${dadosPedido.fornecedorId} -> Nome: ${nomeFornecedor}`);
    } else {
      Logger.log(`AVISO: Fornecedor com ID ${dadosPedido.fornecedorId} não encontrado. A manter o nome antigo.`);
      nomeFornecedor = oldRowData[headersPed.indexOf('Fornecedor')]; // Fallback
    }

    // 3. Cria a nova linha fundindo os dados antigos com os novos
    const novaLinhaPedido = headersPed.map((header, index) => {
        switch(header) {
            case 'Status': return 'Aguardando Aprovacao';
            case 'Motivo Rejeição': return '';
            case 'Data Ultima Edicao': return new Date();
            case 'Fornecedor': return nomeFornecedor;
            // Pega os valores atualizados do frontend
            case 'Observacoes': return dadosPedido.observacoes;
            case 'Placa Veiculo': return dadosPedido.placaVeiculo;
            case 'Nome Veiculo': return dadosPedido.nomeVeiculo;
            case 'Total Geral': return dadosPedido.totalGeral;
            case 'ICMS ST Total': return dadosPedido.valorIcms;
            // Para todas as outras colunas (Data, Empresa, etc.), mantém o valor antigo
            case 'Empresa':
                const empresaId = oldRowData[index];
                // Adiciona um apóstrofo para forçar o Google Sheets a tratar como texto
                return "'" + empresaId;
            case 'Data':
            case 'Data Criacao':
            case 'Usuario Criador':
                return oldRowData[index]; // Mantém o valor original da planilha
            
            // Para todas as outras colunas, mantém o valor antigo por segurança
            default: return oldRowData[index];
        }
    });
    Logger.log("Nova linha de pedido preparada para ser guardada: %s", JSON.stringify(novaLinhaPedido));

    // 4. Escreve a linha atualizada de volta
    sheetPedidos.getRange(sheetRowIndex, 1, 1, novaLinhaPedido.length).setValues([novaLinhaPedido]);
    Logger.log("Linha do pedido na folha de cálculo 'pedidos' foi atualizada.");
    
    // 5. Atualiza os itens (remove os antigos e adiciona os novos)
    Logger.log("A iniciar a atualização dos itens...");
    const itensData = sheetItens.getDataRange().getValues();
    const idxNumPedItem = itensData[0].indexOf('NUMERO PEDIDO');
    const numeroPedidoAlvo = parseInt(dadosPedido.numeroPedido, 10);
    let itensRemovidos = 0;

    for (let i = itensData.length - 1; i > 0; i--) {
        if (parseInt(itensData[i][idxNumPedItem], 10) === numeroPedidoAlvo) {
            sheetItens.deleteRow(i + 2);
            itensRemovidos++;
        }
    }
    Logger.log(`${itensRemovidos} itens antigos foram removidos.`);

    dadosPedido.itens.forEach(item => {
        const novaLinhaItem = headersItens.map(header => item[header] || '');
        sheetItens.appendRow(novaLinhaItem);
    });
    Logger.log(`${dadosPedido.itens.length} novos itens foram adicionados.`);

    return { status: 'success', message: `Pedido #${dadosPedido.numeroPedido} foi corrigido e reenviado.` };
  } catch (e) {
    Logger.log(`ERRO FATAL em reenviarPedidoCorrigido: ${e.message}\nStack: ${e.stack}`);
    return { status: 'error', message: e.message };
  }
}
