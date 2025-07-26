// =================================================================
// CONFIGURAÇÕES GLOBAIS - AJUSTE ESTAS VARIÁVEIS
// =================================================================
const ID_DA_PLANILHA = "1xVLFSqL5SVT6cmZ_9foOkKxJIHqHeCGHLcyryBJ44g0"; // <-- IMPORTANTE: Troque por seu ID real
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
      if (row[colunas.status] === "Aguardando Aprovacao") {
          pendingApprovals.push({
            numeroDoPedido: row[colunas.numeroDoPedido],
            fornecedor: row[colunas.fornecedor],
            totalGeral: row[colunas.totalGeral],
            empresaId: row[colunas.empresaId] // Retorna só o ID
          });
      }
    });
    return pendingApprovals;
  } catch (e) { // <-- ESTE É O BLOCO QUE FALTAVA
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
 * Encontra um pedido pelo seu número e atualiza seu status.
 * @param {string} numeroPedido O número do pedido a ser atualizado.
 * @param {string} novoStatus O novo status a ser definido (ex: "Aprovado", "Rejeitado").
 * @returns {object} Um objeto indicando o sucesso ou falha da operação.
 */
function atualizarStatusPedido(numeroPedido, novoStatus) {
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

    if (colunaNumeroPedido === -1 || colunaStatus === -1) {
      throw new Error("Não foi possível encontrar as colunas 'Número do Pedido' ou 'Status' na planilha.");
    }

    // Encontra a linha correspondente ao pedido (começando da linha 2 da planilha, que é o índice 1 nos dados)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colunaNumeroPedido]) === String(numeroPedido)) {
        const rowIndex = i + 1; // O número da linha na planilha é o índice + 1
        
        // Atualiza a célula do status na linha encontrada
        sheet.getRange(rowIndex, colunaStatus + 1).setValue(novoStatus);
        
        console.log(`Pedido #${numeroPedido} atualizado para o status "${novoStatus}" na linha ${rowIndex}.`);
        return { status: 'success', message: `Pedido ${novoStatus.toLowerCase()} com sucesso!` };
      }
    }

    // Se o loop terminar e não encontrar o pedido
    throw new Error(`Pedido #${numeroPedido} não encontrado.`);

  } catch (e) {
    console.error(`Erro em atualizarStatusPedido: ${e.message}`);
    return { status: 'error', message: e.message };
  }
}
