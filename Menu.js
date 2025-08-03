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

function testarAppCache() {
    console.log("--- Testando o AppCache ---");
    if (AppCache.userCompanies.length > 0) {
        console.log("✅ SUCESSO! O cache de empresas foi carregado.");
        console.log("Empresas no cache:", AppCache.userCompanies);
    } else {
        console.error("❌ FALHA! O cache de empresas está vazio. Verifique os logs da função 'loadUserCompanies' e do backend.");
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
function atualizarStatusPedido(numeroPedido, novoStatus, motivo) {
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

    if (colunaNumeroPedido === -1 || colunaStatus === -1) {
      throw new Error("Não foi possível encontrar as colunas 'Número do Pedido' ou 'Status' na planilha.");
    }
    if (novoStatus === 'Rejeitado' && colunaMotivo === -1) {
        throw new Error("A coluna 'Motivo Rejeição' é necessária para rejeitar um pedido, mas não foi encontrada.");
    }

    // Encontra a linha correspondente ao pedido (começando da linha 2 da planilha, que é o índice 1 nos dados)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colunaNumeroPedido]) === String(numeroPedido)) {
        const rowIndex = i + 1; // O número da linha na planilha é o índice + 1
        
        // Atualiza a célula do status na linha encontrada
        sheet.getRange(rowIndex, colunaStatus + 1).setValue(novoStatus);

        // Se for uma rejeição, salva o motivo. Se for aprovação, limpa o motivo.
        if (colunaMotivo !== -1) {
            sheet.getRange(rowIndex, colunaMotivo + 1).setValue(motivo);
        }
        
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

/**
 * Busca e processa todos os dados necessários para o dashboard do administrador em uma única chamada.
 * Retorna uma lista de empresas, e para cada uma, os seus 2 últimos pedidos APROVADOS.
 */
/**
 * VERSÃO FINAL E DEFINITIVA
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
    const todasAsEmpresas = empresasData.map(([id, nome]) => ({ id: String(id).trim(), nome: nome, ultimosPedidosAprovados: [] }));

    // Obter TODOS os pedidos e normalizar cabeçalhos
    const pedidosData = pedidosSheet.getDataRange().getValues();
    const pedidosHeaders = pedidosData.shift().map(h => String(h || '').toUpperCase().trim());
    
    const colunas = {
      empresaId: pedidosHeaders.indexOf("EMPRESA"),
      status: pedidosHeaders.indexOf("STATUS"),
      dataCriacao: pedidosHeaders.indexOf("DATA CRIACAO"),
      data: pedidosHeaders.indexOf("DATA"),
      numeroDoPedido: pedidosHeaders.indexOf("NÚMERO DO PEDIDO"),
      fornecedor: pedidosHeaders.indexOf("FORNECEDOR"),
      totalGeral: pedidosHeaders.indexOf("TOTAL GERAL"),
      estadoFornecedor: pedidosHeaders.indexOf("ESTADO FORNECEDOR")
    };
    
    if (colunas.status === -1) throw new Error("A coluna 'STATUS' não foi encontrada.");
    if (colunas.empresaId === -1) throw new Error("A coluna 'EMPRESA' não foi encontrada.");
    if (colunas.estadoFornecedor === -1) throw new Error("A coluna 'ESTADO FORNECEDOR' não foi encontrada.");

    // Agrupar pedidos APROVADOS
    const pedidosAprovadosPorEmpresa = {};
    pedidosData.forEach(row => {
      // Lógica de filtro idêntica à da função de teste que funcionou
      const status = String(row[colunas.status] || '').trim().toUpperCase();
      const empresaId = String(row[colunas.empresaId] || '').trim(); 

      if (status === "APROVADO" && empresaId) {
        if (!pedidosAprovadosPorEmpresa[empresaId]) {
          pedidosAprovadosPorEmpresa[empresaId] = [];
        }
        
        const dataObj = row[colunas.dataCriacao] || row[colunas.data];
        const dataFormatada = (dataObj instanceof Date) ? dataObj.toISOString() : dataObj;

        pedidosAprovadosPorEmpresa[empresaId].push({
          numeroDoPedido: row[colunas.numeroDoPedido],
          fornecedor: row[colunas.fornecedor],
          totalGeral: row[colunas.totalGeral],
          data: dataFormatada,
          estadoFornecedor: row[colunas.estadoFornecedor]
        });
      }
    });
    
    Logger.log("Pedidos aprovados agrupados (após filtro final): " + Object.keys(pedidosAprovadosPorEmpresa).length + " empresas com pedidos aprovados.");
    
    // Ordenar, fatiar e juntar os dados
    todasAsEmpresas.forEach(empresa => {
      const pedidosDaEmpresa = pedidosAprovadosPorEmpresa[empresa.id];
      if (pedidosDaEmpresa && pedidosDaEmpresa.length > 0) {
        pedidosDaEmpresa.sort((a, b) => new Date(b.data) - new Date(a.data));
        empresa.ultimosPedidosAprovados = pedidosDaEmpresa.slice(0, 2);
      }
    });
    
    return { status: 'ok', data: todasAsEmpresas };

  } catch (e) {
    Logger.log(`ERRO em getDashboardAdminData: ${e.message}`);
    return { status: 'error', message: e.message, data: [] };
  }
}

function testarFiltroDePedidos() {
  try {
    const sheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName(NOME_DA_ABA_DE_PEDIDOS);
    if (!sheet) {
      Logger.log("ERRO: Planilha de Pedidos não encontrada.");
      return;
    }

    Logger.log("--- INICIANDO TESTE DE FILTRO FINAL ---");

    const data = sheet.getDataRange().getValues();
    const headers = data.shift().map(h => String(h).toUpperCase().trim());

    const statusIndex = headers.indexOf("STATUS");
    const empresaIndex = headers.indexOf("EMPRESA");

    Logger.log(`Índice da coluna STATUS: ${statusIndex}`);
    Logger.log(`Índice da coluna EMPRESA: ${empresaIndex}`);

    if (statusIndex === -1 || empresaIndex === -1) {
      Logger.log("ERRO: Não foi possível encontrar as colunas 'STATUS' ou 'EMPRESA'. Verifique os nomes no cabeçalho.");
      return;
    }

    let contadorDeAprovados = 0;

    // Vamos verificar cada linha
    data.forEach((row, index) => {
      const statusRaw = row[statusIndex];
      const empresaRaw = row[empresaIndex];

      const statusProcessed = String(statusRaw || '').trim().toUpperCase();
      const empresaProcessed = String(empresaRaw || '').trim();

      const aCondicaoEhVerdadeira = (statusProcessed === "APROVADO" && empresaProcessed.length > 0);

      // Log detalhado para as 10 primeiras linhas
      if (index < 10) {
        Logger.log(`Linha ${index + 2}: | Status Lido: "${statusRaw}" | Empresa Lido: "${empresaRaw}" | A Condição é: ${aCondicaoEhVerdadeira}`);
      }

      if (aCondicaoEhVerdadeira) {
        contadorDeAprovados++;
      }
    });

    Logger.log(`--- RESULTADO FINAL DO TESTE ---`);
    Logger.log(`Total de pedidos que passaram no filtro: ${contadorDeAprovados}`);
    Logger.log(`---------------------------------`);

  } catch(e) {
    Logger.log("ERRO durante o teste de filtro: " + e.message);
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

function testarBuscaDeRejeitados() {
    // IMPORTANTE: Coloque aqui o seu nome de utilizador exatamente como está no localStorage.
    const meuUsuarioDeTeste = "admin"; // ou "seu.nome", etc.

    Logger.log(`--- INICIANDO TESTE DE BUSCA DE PEDIDOS REJEITADOS PARA: "${meuUsuarioDeTeste}" ---`);
    const resultado = getMeusPedidosRejeitados(meuUsuarioDeTeste);

    if (resultado.length > 0) {
        Logger.log(`✅ SUCESSO! Foram encontrados ${resultado.length} pedidos:`);
        Logger.log(JSON.stringify(resultado, null, 2));
    } else {
        Logger.log("⚠️ AVISO: A busca não retornou nenhum pedido rejeitado para este utilizador.");
        Logger.log("Verifique se o nome de utilizador está correto e se existem pedidos com o status 'Rejeitado' para ele na planilha.");
    }
}

/**
 * BUSCA TODOS OS DADOS DE UM PEDIDO ESPECÍFICO PARA CORREÇÃO.
 * @param {string} numeroPedido O número do pedido a ser buscado.
 * @returns {Object} Um objeto completo com os dados do cabeçalho e a lista de itens do pedido.
 */
function getDadosCompletosDoPedido(numeroPedido) {
   try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPedidos = ss.getSheetByName("pedidos");
    const sheetItens = ss.getSheetByName("itens pedido");
    const sheetFornecedores = ss.getSheetByName("fornecedores");

    // 1. Busca os dados do cabeçalho do pedido
    const pedidosData = sheetPedidos.getDataRange().getValues();
    const headersPed = pedidosData.shift();
    const rowPedido = pedidosData.find(row => String(row[headersPed.indexOf('Número do Pedido')]) === String(numeroPedido));
    if (!rowPedido) throw new Error(`Pedido ${numeroPedido} não encontrado.`);
    
    const dadosPedido = getObjectFromRow(rowPedido, headersPed);

    // 2. Busca o ID do fornecedor correspondente de forma robusta
    const nomeFornecedorDoPedido = dadosPedido['Fornecedor'];
    if (!nomeFornecedorDoPedido) {
        Logger.log(`Aviso: O pedido ${numeroPedido} não tem um fornecedor associado.`);
    } else {
        const fornecedoresData = sheetFornecedores.getDataRange().getValues();
        const headersForn = fornecedoresData.shift();
        const headersFornLowerCase = headersForn.map(h => String(h).toLowerCase()); // Cabeçalhos em minúsculas para busca

        const idxIdForn = headersFornLowerCase.indexOf('id');
        
        // Tenta encontrar a coluna de nome do fornecedor com nomes comuns, ignorando maiúsculas/minúsculas
        let idxNomeForn = headersFornLowerCase.indexOf('fornecedor');
        if (idxNomeForn === -1) idxNomeForn = headersFornLowerCase.indexOf('razao social');
        if (idxNomeForn === -1) idxNomeForn = headersFornLowerCase.indexOf('nome_fornecedor');

        if (idxIdForn > -1 && idxNomeForn > -1) {
            const nomeFornecedorKey = String(nomeFornecedorDoPedido).trim().toLowerCase();
            const fornecedorEncontrado = fornecedoresData.find(row => String(row[idxNomeForn]).trim().toLowerCase() === nomeFornecedorKey);
            
            if (fornecedorEncontrado) {
              dadosPedido['ID_FORNECEDOR'] = fornecedorEncontrado[idxIdForn];
              Logger.log(`Fornecedor encontrado para "${nomeFornecedorDoPedido}". ID adicionado: ${dadosPedido['ID_FORNECEDOR']}`);
            } else {
              Logger.log(`AVISO: Fornecedor "${nomeFornecedorDoPedido}" do pedido ${numeroPedido} não foi encontrado na folha de cálculo 'fornecedores'.`);
            }
        } else {
            Logger.log(`AVISO: Não foi possível encontrar as colunas de ID ou Nome do Fornecedor na folha de cálculo 'fornecedores'.`);
        }
    }

    // 3. Busca todos os itens associados a esse pedido
    const itensData = sheetItens.getDataRange().getValues();
    const headersItens = itensData.shift();
    const idxNumPedItem = headersItens.indexOf('NUMERO PEDIDO');

    const itensDoPedido = itensData.filter(row => String(row[idxNumPedItem]) === String(numeroPedido))
      .map(row => getObjectFromRow(row, headersItens));

    dadosPedido.itens = itensDoPedido;
    
    return dadosPedido;
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
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPedidos = ss.getSheetByName("pedidos");
    const sheetItens = ss.getSheetByName("itens pedido");

    const pedidosData = sheetPedidos.getDataRange().getValues();
    const headersPed = pedidosData[0];
    const idxNumPed = headersPed.indexOf('Número do Pedido');
    const linhaParaAtualizar = pedidosData.findIndex(row => row[idxNumPed] == dadosPedido.numeroPedido);

    const rowIndexInArray = pedidosData.findIndex(row => String(row[idxNumPed]) === String(dadosPedido.numeroPedido));
    
    if (rowIndexInArray === -1) {
      throw new Error("Não foi possível encontrar o pedido original para atualizar.");
    }
    const rowIndexInSheet = rowIndexInArray + 1; // A linha real na folha de cálculo

    // Pega a linha inteira de dados antigos para preservar campos não editáveis
    let linhaAntiga = pedidosData[rowIndexInArray];

    // Atualiza os campos específicos
    linhaAntiga[headersPed.indexOf('Status')] = 'Aguardando Aprovacao';
    linhaAntiga[headersPed.indexOf('Motivo Rejeição')] = ''; // Limpa o motivo
    linhaAntiga[headersPed.indexOf('Data Ultima Edicao')] = new Date();
    linhaAntiga[headersPed.indexOf('Fornecedor')] = dadosPedido.fornecedor;
    linhaAntiga[headersPed.indexOf('Observacoes')] = dadosPedido.observacoes;
    linhaAntiga[headersPed.indexOf('Nome Veiculo')] = dadosPedido.nomeVeiculo;
    linhaAntiga[headersPed.indexOf('Placa Veiculo')] = dadosPedido.placaVeiculo;
    linhaAntiga[headersPed.indexOf('Total Geral')] = dadosPedido.totalGeral;
    linhaAntiga[headersPed.indexOf('ICMS ST Total')] = dadosPedido.valorIcms;
    
    // Reescreve a linha inteira com os dados atualizados
    sheetPedidos.getRange(rowIndexInSheet, 1, 1, linhaAntiga.length).setValues([linhaAntiga]);
    // Garante que o formato do número do pedido seja mantido como texto
    sheetPedidos.getRange(rowIndexInSheet, idxNumPed + 1).setNumberFormat('@');


    // Remove os itens antigos e adiciona os novos
    const itensData = sheetItens.getDataRange().getValues();
    const idxNumPedItem = itensData[0].indexOf('NUMERO PEDIDO');
    for (let i = itensData.length - 1; i > 0; i--) {
        if (String(itensData[i][idxNumPedItem]) === String(dadosPedido.numeroPedido)) {
            sheetItens.deleteRow(i + 1);
        }
    }
    
    const headersItens = sheetItens.getRange(1, 1, 1, sheetItens.getLastColumn()).getValues()[0];
    dadosPedido.itens.forEach(item => {
        const novaLinhaItem = headersItens.map(header => item[header] || '');
        sheetItens.appendRow(novaLinhaItem);
    });

    return { status: 'success', message: `Pedido #${dadosPedido.numeroPedido} foi corrigido e reenviado.` };
  } catch (e) {
    Logger.log(`Erro em reenviarPedidoCorrigido: ${e.message}\n${e.stack}`);
    return { status: 'error', message: e.message };
  }
}

function testarBuscaDeDadosCompletos() {
    // IMPORTANTE: Coloque aqui o número de um pedido que foi rejeitado.
    const numeroPedidoParaTestar = "001387"; 

    Logger.log(`--- INICIANDO TESTE DE BUSCA DE DADOS PARA O PEDIDO: "${numeroPedidoParaTestar}" ---`);
    const resultado = getDadosCompletosDoPedido(numeroPedidoParaTestar);

    if (resultado) {
        Logger.log(`✅ SUCESSO! Dados encontrados para o pedido:`);
        Logger.log(JSON.stringify(resultado, null, 2));
        
        // Verificações importantes
        if (resultado.ID_FORNECEDOR) {
            Logger.log(`-> Verificação de ID do Fornecedor: SUCESSO (ID: ${resultado.ID_FORNECEDOR})`);
        } else {
            Logger.log(`-> Verificação de ID do Fornecedor: FALHA (ID não foi encontrado/adicionado)`);
        }
        Logger.log(`-> Verificação de Itens: SUCESSO (${resultado.itens.length} itens encontrados)`);

    } else {
        Logger.log(`❌ FALHA: A busca não retornou nenhum dado para este pedido.`);
        Logger.log("Verifique se o número do pedido está correto e se ele existe na folha de cálculo 'pedidos'.");
    }
}

