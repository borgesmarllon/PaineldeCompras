// ===============================================
    // FUNÇÕES PARA PEDIDOS DE COMPRA
    // ===============================================

    /**
     * Retorna o próximo número sequencial para um novo pedido.
     * Cria a planilha 'Pedidos' se não existir.
     * @returns {string} O próximo número de pedido formatado como '0001'.
     */
    function getProximoNumeroPedido(empresaCodigo) {
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

      const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      const empresaCodigoTratado = String(empresaCodigo).trim();
      
      const numeros = data
        .filter(row => {
          const idNaLinha = String(row[colEmpresa]).trim();
          
          // --- AQUI ESTÁ A CORREÇÃO FINAL ---
          // Converte ambos os IDs para números antes de comparar.
          // parseInt("1") vira 1. parseInt("001") também vira 1. A comparação funciona.
          return parseInt(idNaLinha, 10) === parseInt(empresaCodigoTratado, 10);
        })
        .map(row => parseInt(row[colNumero], 10))
        .filter(n => !isNaN(n));

      const proximoNumero = numeros.length > 0 ? Math.max(...numeros) + 1 : 1;
      
      return proximoNumero.toString().padStart(6, '0');
    }


        /**
     * Salva um novo pedido de compra na planilha 'Pedidos'.
     * @param {Object} pedido - Objeto contendo os detalhes do pedido (numero, data, fornecedor, itens, totalGeral, placaVeiculo, nomeVeiculo, observacoes).
     * @returns {Object} Um objeto com status e mensagem.
     */
    function salvarPedido(pedido, usuarioLogado) {
      console.log('📋 === INÍCIO salvarPedido ===');
      console.log('📋 Objeto pedido recebido:', JSON.stringify(pedido, null, 2));
      
      const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
      if (!sheet) {
        return { status: 'error', message: 'Planilha "Pedidos" não encontrada. Contate o administrador.' };
      }

      // Normalizar o número do pedido - aceitar tanto 'numeroPedido' quanto 'numero'
      const numeroPedido = pedido.numeroPedido || pedido.numero;
      console.log('📋 Número do pedido normalizado:', numeroPedido);
      
      if (!numeroPedido) {
        console.error('❌ Número do pedido não encontrado em:', pedido);
        return { status: 'error', message: 'Número do pedido é obrigatório.' };
      }

      // Normalizar empresa (aceitar empresaId ou empresa)
      const empresaId = pedido.empresaId || pedido.empresa;
      console.log('📋 ID da empresa normalizado:', empresaId);
      
      console.log('📋 Total geral recebido:', pedido.totalGeral);
      console.log('📋 Itens recebidos:', pedido.itens ? pedido.itens.length : 0, 'itens');

      const itensJSON = JSON.stringify(pedido.itens);

      const fornecedoresSheet = SpreadsheetApp.getActive().getSheetByName('Fornecedores');
      let fornecedorCnpj = '';
      let fornecedorEndereco = '';
      let condicaoPagamentoFornecedor = '';
      let formaPagamentoFornecedor = '';
      let estadoFornecedor = '';
      let cidadeFornecedor = '';

      if (fornecedoresSheet) {
        const fornecedoresData = fornecedoresSheet.getRange(2, 1, fornecedoresSheet.getLastRow() - 1, fornecedoresSheet.getLastColumn()).getValues();
        const foundFornecedor = fornecedoresData.find(row => String(row[0]) === pedido.fornecedorId); 
        if (foundFornecedor) {
          fornecedorCnpj = String(foundFornecedor[3] || '');
          fornecedorEndereco = String(foundFornecedor[4] || '');
          condicaoPagamentoFornecedor = String(foundFornecedor[5] || '');
          formaPagamentoFornecedor = String(foundFornecedor[6] || '');
          estadoFornecedor = String(foundFornecedor[10] || ''); // Coluna 11 (índice 10) = Estado
          cidadeFornecedor = String(foundFornecedor[11] || '');
        }
      }

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const colunaUsuarioCriador = headers.indexOf("Usuario Criador");
      if (colunaUsuarioCriador === -1) {
        throw new Error("A coluna 'Usuario Criador' não foi encontrada na planilha de Pedidos.");
    }
      const rowData = new Array(headers.length).fill('');

      const dataToSave = {
        'Número do Pedido': "'" + numeroPedido, // Adiciona o apóstrofo
        'Empresa': "'" + empresaId,     // Usar o mesmo nome que está em getProximoNumeroPedido
        'Data': pedido.data,
        'Fornecedor': pedido.fornecedor,
        'CNPJ Fornecedor': fornecedorCnpj,
        'Endereço Fornecedor': fornecedorEndereco,
        'Estado Fornecedor': estadoFornecedor,
        'Condição Pagamento Fornecedor': condicaoPagamentoFornecedor,
        'Forma Pagamento Fornecedor': formaPagamentoFornecedor,
        'Placa Veiculo': pedido.placaVeiculo,
        'Nome Veiculo': pedido.nomeVeiculo,
        'Observacoes': pedido.observacoes,
        'Total Geral': parseFloat(pedido.totalGeral) || 0, // Garantir que é um número
        'ICMS ST Total': parseFloat(pedido.valorIcms) || 0,
        'Status': 'AGUARDANDO APROVACAO',
        'Itens': itensJSON,
        'Data Criacao': formatarDataParaISO(new Date()), // Timestamp de criação padronizado
        'Produto Fornecedor': pedido.produtoFornecedor,
        'Usuario Criador': pedido.usuarioLogado
      };



      headers.forEach((header, index) => {
        if (dataToSave.hasOwnProperty(header)) {
          rowData[index] = dataToSave[header];
          console.log(`📋 Mapeando coluna "${header}":`, dataToSave[header]);
        }
      });

      console.log('📋 Dados finais para salvar:', rowData);
      sheet.appendRow(rowData);

      console.log('✅ Pedido salvo com sucesso:', numeroPedido);
      console.log('📋 === FIM salvarPedido ===');
      return { status: 'ok', message: `Pedido ${numeroPedido} salvo com sucesso!` };
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

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    // Lógica robusta para encontrar as colunas, independente de pequenas variações no nome
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
      itens: headers.findIndex(h => h.toUpperCase() .includes('ITENS')),
      estado: headers.findIndex(h => h.toUpperCase().includes('ESTADO FORNECEDOR')),
      dataCriacao: headers.findIndex(h => h.toUpperCase() === 'DATA CRIACAO'),
      aliquota: headers.findIndex(h => h.toUpperCase().includes('ALIQUOTA IMPOSTO')),
      icmsSt: headers.findIndex(h => h.toUpperCase().includes('ICMS ST TOTAL')),
      usuarioCriador: headers.findIndex(h => h.toUpperCase().includes('USUARIO CRIADOR'))
    };

    // Validação para garantir que colunas essenciais foram encontradas
    for (const key in colunas) {
        if (colunas[key] === -1) {
            Logger.log(`AVISO: A coluna "${key}" não foi encontrada. O filtro ou o dado retornado para este campo será ignorado.`);
        }
    }

    const pedidosEncontrados = data.filter(row => {
      // O filtro de empresa é o primeiro e obrigatório
      const statusDoPedido = row[colunas.status];
      if (!statusDoPedido || String(statusDoPedido).trim() === '') {
        return false; //
      }

      const empresaPlanilha = String(row[colunas.empresa]).trim();
      const empresaFiltro = String(params.empresaId).trim();
      if (empresaPlanilha !== empresaFiltro) {
         return false; // Se não for da empresa correta, já descarta a linha
      }
      
      let match = true;

      // --- Filtro por Empresa (SEMPRE APLICADO) ---
      if (params.empresaId && match) {
          const empresaPlanilha = String(row[colunas.empresa]).trim();
          const empresaFiltro = String(params.empresaId).trim();
          if (empresaPlanilha !== empresaFiltro) {
              return false;
         }
      } else if (params.empresaId) { // Se não houver empresa selecionada, não retorna nada
          return false;
      }
      
      // Filtro Principal (Nº Pedido ou Fornecedor)
      if (params.mainSearch && match) {
        const termo = params.mainSearch.toLowerCase().trim();
        const numPedido = String(row[colunas.numeroDoPedido]).toLowerCase();
        const fornecedor = String(row[colunas.fornecedor] ?? '').toLowerCase();
        if (!numPedido.includes(termo) && !fornecedor.includes(termo)) {
          return false;
        }
      }

      // Filtro por Data
      if (params.dateStart && params.dateEnd && match) {
        const dataPedido = new Date(row[colunas.data]);
        const dataInicio = new Date(params.dateStart + 'T00:00:00');
        const dataFim = new Date(params.dateEnd + 'T23:59:59');
        if (dataPedido < dataInicio || dataPedido > dataFim) {
          return false;
        }
      }
      
      // Filtro por Placa
      if (params.plateSearch && match && colunas.placaVeiculo !== -1) {
          const placaPlanilha = String(row[colunas.placaVeiculo]).toLowerCase().trim();
          const placaFiltro = params.plateSearch.toLowerCase().trim();
          Logger.log(`Comparando Placa: Planilha='${placaPlanilha}', Filtro='${placaFiltro}'`);
          if (placaPlanilha !== placaFiltro) {
            return false;
          }
      }

      // Filtro por Usuário Criador
      if (params.usuarioCriador && match && colunas.usuarioCriador !== -1) {
          const criadorPlanilha = String(row[colunas.usuarioCriador]).toLowerCase().trim();
          const criadorFiltro = params.usuarioCriador.toLowerCase().trim();
          Logger.log(`Comparando Criador: Planilha='${criadorPlanilha}', Filtro='${criadorFiltro}'`);
          if (criadorPlanilha !== criadorFiltro) {
            return false;
          }
      }

      return true;
    }).map(row => {
      // Mapeia a linha para um objeto, garantindo que a data seja serializável
      const dataDoPedido = row[colunas.data];

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
        estado: row[colunas.estado],
        dataCriacao: row[colunas.dataCriacao] instanceof Date ? Utilities.formatDate(row[colunas.dataCriacao], "GMT-03:00", "yyyy-MM-dd'T'HH:mm:ss'Z'") : row[colunas.dataCriacao],
        aliquota: row[colunas.aliquota],
        icmsSt: row[colunas.icmsSt],
        usuarioCriador: row[colunas.usuarioCriador]
      };

      const itensJSON = row[colunas.itens];
      if (colunas.itens !== -1 && itensJSON && String(itensJSON).trim() !== '') {
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
      return pedido;
    });
     return { status: 'success', data: pedidosEncontrados };

  } catch (e) {
    Logger.log("Erro na função buscarPedidos: " + e + "\nStack: " + e.stack);
    return { status: 'error', message: e.toString() };
  }
}

function testarBuscarPedidos() {
  const paramsTeste = {
    empresaId: "001", mainSearch: "001", dateStart: "", dateEnd: "", plateSearch: "", usuarioCriador: ""
  };
  Logger.log(`--- INICIANDO TESTE para a função buscarPedidos ---`);
  const resultado = buscarPedidosv2(paramsTeste);
  Logger.log("--- RESULTADO DO TESTE ---");
  Logger.log(JSON.stringify(resultado, null, 2));
}

 function listarPedidosPorEmpresa(empresa) {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
      if (!sheet) {
        return [];
      }

      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const pedidos = [];

      const indexEmpresa = headers.indexOf('Empresa');
      if (indexEmpresa === -1) {
        throw new Error('Coluna "Empresa" não encontrada na planilha Pedidos.');
      }

      for (let i = 1; i < data.length; i++) {
        const linha = data[i];
        const empresaDaLinha = String(linha[indexEmpresa]).trim();

        if (empresaDaLinha === String(empresa).trim()) {
          const pedido = {};
          headers.forEach((header, idx) => {
            pedido[header] = linha[idx];
          });
          pedidos.push(pedido);
        }
      }

      return pedidos;
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
     * Função de teste para verificar comunicação backend
     */
    function testarComunicacao() {
      console.log('✅ [TESTE] Função testarComunicacao chamada com sucesso');
      return {
        status: 'success',
        message: 'Comunicação funcionando',
        timestamp: new Date().toISOString()
      };
    }

    /**
     * Função de teste ainda mais simples
     */
    function testeSimples() {
      return 'OK';
    }

    /**
     * ===============================================
     * BACKEND - SISTEMA DE RASCUNHOS
     * Google Apps Script Functions
     * ===============================================
     */

    /**
     * Salva um rascunho na planilha
     * @param {Object} dadosRascunho - Dados do rascunho a ser salvo
     * @returns {Object} - Resposta com status e ID do rascunho
     */
    function salvarRascunho(dadosRascunho) {
      try {
        console.log('📝 Salvando rascunho:', dadosRascunho);
        
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
        
        // Obter a planilha
        const planilha = SpreadsheetApp.openById(PLANILHA_ID);
        const aba = planilha.getSheetByName('Pedidos') || planilha.insertSheet('Pedidos');
        
        // Gerar ID único para o rascunho
        const agora = new Date();
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
          const foundFornecedor = fornecedoresData.find(row => String(row[1]) === dadosRascunho.fornecedor); 
          if (foundFornecedor) {
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
          'Fornecedor': dadosRascunho.fornecedor.trim(),
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
          'Data Ultima Edicao': formatarDataParaISO(agora), // Sempre usar data/hora atual padronizada
          'Produto Fornecedor': dadosRascunho.produtoFornecedor
        };
        
        // Verificar se é uma atualização de rascunho existente
        if (dadosRascunho.rascunhoId) {
          const linhaExistente = encontrarLinhaRascunho(aba, dadosRascunho.rascunhoId);
          if (linhaExistente > 0) {
            // Atualizar rascunho existente usando a mesma estrutura
            dadosParaSalvar['Número do Pedido'] = "'" + dadosRascunho.rascunhoId;
            salvarDadosNaPlanilha(aba, dadosParaSalvar, linhaExistente);
            
            console.log('✅ Rascunho atualizado com sucesso:', dadosRascunho.rascunhoId);
            return {
              status: 'success',
              message: 'Rascunho atualizado com sucesso!',
              rascunhoId: dadosRascunho.rascunhoId
            };
          }
        }
        
        // Salvar novo rascunho usando a mesma estrutura da função salvarPedido
        salvarDadosNaPlanilha(aba, dadosParaSalvar);
        
        console.log('✅ Rascunho salvo com sucesso:', rascunhoId);
        return {
          status: 'success',
          message: 'Rascunho salvo com sucesso!',
          rascunhoId: rascunhoId
        };
        
      } catch (error) {
        console.error('❌ Erro ao salvar rascunho:', error);
        return {
          status: 'error',
          message: 'Erro interno ao salvar rascunho: ' + error.message
        };
      }
    }

    /**
     * Busca todos os rascunhos de uma empresa
     * @param {string} empresaId - ID da empresa
     * @returns {Object} - Lista de rascunhos
     */
    function buscarRascunhos(empresaId) {
      console.log('🔍 [BACKEND] === INÍCIO buscarRascunhos ===');
      console.log('🔍 [BACKEND] Parâmetro empresaId:', empresaId);
      console.log('🔍 [BACKEND] Tipo do empresaId:', typeof empresaId);
      
      try {
        // ID da planilha definido localmente
        var planilhaId = '1J7CE_BZ8eUsXhjkmgxAIIWjMTOr2FfSfIMONqE4UpHA';
        
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
          produtoFornecedor: cabecalhos.indexOf('Produto Fornecedor')
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
                produtoFornecedor: linha[indices.produtoFornecedor] || ''
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
        var planilhaId = '1J7CE_BZ8eUsXhjkmgxAIIWjMTOr2FfSfIMONqE4UpHA';
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
          produtoFornecedor: cabecalhos.indexOf('produtoFornecedor')
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
              produtoFornecedor: linha[indices.produtoFornecedor] || ''
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

    const pedidosAprovados = [];
    data.forEach(row => {
      // Filtra pelo usuário logado E pelo status "Aprovado"
      if (row[colunas.usuarioCriador] === usuarioLogado && row[colunas.status] === "Aprovado") {
        
        // Monta um objeto limpo com todos os dados da linha do pedido
        const dataDoPedido = row[colunas.data];
        
        // --- CORREÇÃO DE SERIALIZAÇÃO APLICADA AQUI ---
        // Monta um objeto apenas com os dados necessários, garantindo que a data seja texto.
        pedidosAprovados.push({
          'Número_do_Pedido': row[colunas.numeroDoPedido],
          'Data': dataDoPedido instanceof Date ? Utilities.formatDate(dataDoPedido, "GMT-03:00", "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'") : dataDoPedido,
          'Fornecedor': row[colunas.fornecedor],
          'Empresa': row[colunas.empresa],
          'Total Geral': row[colunas.totalGeral]
        });
      }
    });
    
    console.log(`Encontrados ${pedidosAprovados.length} pedidos aprovados para o usuário ${usuarioLogado}.`);
    return { status: 'success', data: pedidosAprovados };

  } catch(e) {
    console.error(`Erro em getMeusPedidosAprovados: ${e.message}`);
    return { status: 'error', message: e.message };
  }
}     

function testarGetMeusPedidosAprovados() {
  // --- CONFIGURE AQUI ---
  // Coloque o nome de um usuário que você sabe que tem pedidos aprovados na planilha.
  const usuarioTeste = "admin"; 
  const empresa = "002";
  
  Logger.log(`--- INICIANDO TESTE para a função getMeusPedidosAprovados ---`);
  Logger.log(`Procurando pedidos aprovados para o usuário: "${usuarioTeste}"`);
  
  // Chama a função principal com os dados de teste
  const resultado = getMeusPedidosAprovados(usuarioTeste, empresa);
  
  // Mostra o resultado completo no log
  Logger.log("--- RESULTADO DO TESTE ---");
  Logger.log(JSON.stringify(resultado, null, 2)); // Usamos JSON.stringify para ver o objeto de forma clara
  Logger.log("--------------------------");
  
  if (resultado.status === 'success') {
    Logger.log(`✅ Teste bem-sucedido! Encontrados ${resultado.data.length} pedidos.`);
  } else {
    Logger.log(`❌ Teste falhou! Mensagem de erro: ${resultado.message}`);
  }
}
