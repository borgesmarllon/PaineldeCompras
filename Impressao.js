//const ID_DA_PLANILHA = "1xVLFSqL5SVT6cmZ_9foOkKxJIHqHeCGHLcyryBJ44g0"

// ===============================================
    // FUNÇÕES PARA IMPRESSAO
    // ===============================================    

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

    function getDadosPedidoParaImpressaoAdmin(numeroPedido, empresaId) {
      console.log(`Buscando pedido ${numeroPedido} da empresa ${empresaId} para admin`);
      
      try {
        // Sua lógica para buscar o pedido específico da empresa
        // Exemplo (adapte para sua estrutura):
        
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
        var data = sheet.getDataRange().getValues();
        
        // Encontrar o pedido específico
        for (var i = 1; i < data.length; i++) {
          if (data[i][0] == numeroPedido && data[i][1] == empresaId) { // Ajuste os índices conforme sua planilha
            return {
              numeroDoPedido: data[i][0],
              empresaId: data[i][1],
              data: data[i][2],
              fornecedor: data[i][3],
              totalGeral: data[i][4],
              // ... outros campos
              // Incluir dados da empresa também
              enderecoEmpresa: "endereço da empresa",
              cnpjEmpresa: "cnpj da empresa",
              // etc.
            };
          }
        }
        
        return null; // Pedido não encontrado
        
      } catch (error) {
        console.error("Erro ao buscar pedido para admin:", error);
        throw error;
      }
    }

/**
 * SUBSTITUIÇÃO UNIFICADA E SEGURA para as funções de impressão.
 * Busca um pedido pelo seu número E ID da empresa, e enriquece o objeto
 * com os dados cadastrais completos da empresa e do fornecedor.
 * @param {string} numeroPedido O número do pedido a ser buscado.
 * @param {string} empresaId O ID da empresa do pedido.
 * @returns {Object|null} Um objeto completo com todos os dados para impressão ou null se não encontrado.
 */
function getPedidoCompletoPorId(numeroPedido, empresaId) {
  try {
    // 1. BUSCA OS DADOS DO PEDIDO
    const pedidoData = buscarPedidoPorId(numeroPedido, empresaId);
    if (!pedidoData) {
      throw new Error(`Pedido ${numeroPedido} da empresa ${empresaId} não encontrado.`);
    }

    // 2. BUSCA OS DADOS DA EMPRESA E ANEXA AO PEDIDO
    const dadosEmpresa = getDadosEmpresaPorId(empresaId);
    pedidoData.empresaInfo = dadosEmpresa || {};

    // 3. BUSCA OS DADOS DO FORNECEDOR E ANEXA AO PEDIDO
    const dadosFornecedor = getDadosFornecedorPorNome(pedidoData.fornecedor);
    pedidoData.fornecedorInfo = dadosFornecedor || {};

    Logger.log(`Pedido ${numeroPedido} encontrado e enriquecido com sucesso.`);
    return pedidoData;

  } catch (e) {
    Logger.log(`ERRO em getPedidoCompletoPorId: ${e.message}`);
    return null;
  }
}

function buscarPedidoPorId(numeroPedido, empresaId) {
  // Esta função agora abre sua própria conexão
  const planilha = SpreadsheetApp.openById(ID_DA_PLANILHA);
  const sheet = planilha.getSheetByName('Pedidos'); // Pega a aba
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const colunas = {};
  headers.forEach((header, index) => { colunas[toCamelCase(header)] = index; });

  const pedidoRow = data.find(row => 
    String(row[colunas.numeroDoPedido]).trim() === String(numeroPedido).trim() &&
    String(row[colunas.empresa]).trim() === String(empresaId).trim()
  );

  if (!pedidoRow) return null;

  const pedidoData = {};
  for (const key in colunas) {
    if (colunas.hasOwnProperty(key)) {
      let value = pedidoRow[colunas[key]];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
      }
      pedidoData[key] = value;
    }
  }
  
  pedidoData.itens = JSON.parse(pedidoData.itens || '[]');
  return pedidoData;
}


/**
 * Função auxiliar para converter "Nome Cabeçalho" em "nomeCabecalho".
 * @param {string} text - O texto a ser convertido.
 * @returns {string} O texto em camelCase.
 */
function toCamelCase(text) {
  if (!text) return '';
  // Remove acentos, caracteres especiais, e depois converte para camelCase
  const a = 'àáâäæãåāăąçćčđďèéêëēėęěğǵḧîïíīįìłḿñńǹňôöòóœøōõőṕŕřßśšşșťțûüùúūǘůűųẃẍÿýžźż·/_,:;'
  const b = 'aaaaaaaaaacccddeeeeeeeegghiiiiiilmnnnnoooooooooprrssssssttuuuuuuuuuwxyyzzz------'
  const p = new RegExp(a.split('').join('|'), 'g')

  return text.toString().toLowerCase()
    .replace(/\s+/g, '-') // substitui espaços por -
    .replace(p, c => b.charAt(a.indexOf(c))) // substitui caracteres especiais
    .replace(/&/g, '-e-') // substitui & por 'e'
    .replace(/[^\w\-]+/g, '') // remove caracteres inválidos
    .replace(/\-\-+/g, '-') // substitui múltiplos - por um único -
    .replace(/^-+/, '') // remove - do início
    .replace(/-+$/, '') // remove - do final
    .replace(/-(\w)/g, (match, R) => R.toUpperCase()); // Converte para camelCase
}

/**
 * Busca os dados cadastrais de uma empresa específica pelo seu ID.
 * @param {string} empresaId O ID da empresa a ser buscada (ex: "001").
 * @returns {object|null} Um objeto com os dados da empresa ou null se não for encontrada.
 */
function getDadosEmpresaPorId(empresaId) {
  // CORREÇÃO DO ERRO DE DIGITAÇÃO E LÓGICA DE COMPARAÇÃO
  try {
    const planilha = SpreadsheetApp.openById(ID_DA_PLANILHA);
    const sheet = planilha.getSheetByName('Empresas'); // Corrigido de getSheetByNem
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const colunas = {};
    headers.forEach((h, i) => colunas[toCamelCase(h)] = i);

    if (colunas.id === undefined) return null;

    // --- LÓGICA DE COMPARAÇÃO ROBUSTA ---
    // Converte ambos os lados para número antes de comparar.
    // Assim, 1 (da planilha) será igual a "001" (do front-end).
    const empresaRow = data.find(row => 
        parseInt(row[colunas.id], 10) === parseInt(empresaId, 10)
    );
    // --- FIM DA MELHORIA ---

    if (empresaRow) {
      const empresaData = {};
      for (const key in colunas) {
          if (colunas.hasOwnProperty(key)) {
              empresaData[key] = empresaRow[colunas[key]];
          }
      }
      // Garante que o ID retornado seja sempre uma string com zeros à esquerda
      empresaData.id = String(empresaData.id).padStart(3, '0');
      return empresaData;
    }
    return null;
  } catch(e) {
    Logger.log(`ERRO em getDadosEmpresaPorId: ${e.message}`);
    return null;
  }
}

function getDadosFornecedorPorNome(nomeFornecedor) {
    const planilha = SpreadsheetApp.openById(ID_DA_PLANILHA); // Abre o arquivo
    const sheet = planilha.getSheetByName('Fornecedores'); // Pega a aba
    if (!sheet || !nomeFornecedor) return null;

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const colunas = {};
    headers.forEach((h, i) => colunas[toCamelCase(h)] = i);

    // Garante que a coluna com o nome (ex: razaoSocial) existe
    if (colunas.razaoSocial === undefined) return null; 

    const nomeFornecedorNormalizado = String(nomeFornecedor).trim().toUpperCase();
    const fornecedorRow = data.find(row => 
        String(row[colunas.razaoSocial]).trim().toUpperCase() === nomeFornecedorNormalizado
    );

    if (fornecedorRow) {
        const fornecedorData = {};
        for (const key in colunas) {
            if (colunas.hasOwnProperty(key)) {
                fornecedorData[key] = fornecedorRow[colunas[key]];
            }
        }
        return fornecedorData;
    }
    return null;
}

/**
 * ===================================================================
 * FUNÇÃO DE TESTE DEDICADA PARA O CADASTRO DE EMPRESAS
 * ===================================================================
 * Testa a busca de uma empresa específica na planilha 'Empresas' pelo seu ID.
 * Em caso de falha, lista os IDs disponíveis para facilitar a depuração.
 */
function testarBuscaEmpresa() {
  // Para usar, mude o valor da variável 'idParaTestar' abaixo para o ID que você quer verificar.
  const idParaTestar = "001";

  Logger.log("--- INICIANDO TESTE DE BUSCA DE EMPRESA ---");
  Logger.log(`Procurando por ID exato: "${idParaTestar}"`);

  try {
    // 1. Tenta a busca exata usando a mesma função auxiliar da impressão.
    const resultadoExato = getDadosEmpresaPorId(idParaTestar);

    // 2. Analisa o resultado.
    if (resultadoExato) {
      Logger.log("✅ SUCESSO: Correspondência exata encontrada para o ID!");
      Logger.log("Dados encontrados:");
      Logger.log(resultadoExato); // Loga o objeto completo da empresa encontrada.
    } else {
      Logger.log(`❌ FALHA: Nenhuma correspondência 100% exata foi encontrada para o ID "${idParaTestar}".`);
      Logger.log("--- Verificando IDs disponíveis na planilha 'Empresas' ---");

      // 3. Se a busca falhar, lista os IDs que existem na planilha.
      const planilha = SpreadsheetApp.openById(ID_DA_PLANILHA);
      const sheet = planilha.getSheetByName('Empresas');
      if (!sheet) {
        Logger.log("ERRO: Aba 'Empresas' não encontrada.");
        return;
      }
      
      const data = sheet.getDataRange().getValues();
      const headers = data.shift();
      const idIndex = headers.findIndex(h => h.toUpperCase() === 'ID');

      if (idIndex === -1) {
        Logger.log("ERRO: Coluna 'ID' não encontrada na aba 'Empresas'.");
        return;
      }

      // Extrai todos os IDs da coluna, remove vazios e exibe.
      const idsDisponiveis = data.map(row => row[idIndex]).filter(id => id); 
      
      if (idsDisponiveis.length > 0) {
        Logger.log(`💡 SUGESTÃO: Os seguintes IDs foram encontrados na sua planilha: [${idsDisponiveis.join(', ')}]`);
        Logger.log(`Compare o ID que você está buscando ('${idParaTestar}') com a lista acima. Há alguma diferença (espaços, zeros à esquerda, etc.)?`);
      } else {
        Logger.log("Nenhum ID foi encontrado na coluna 'ID' da planilha 'Empresas'.");
      }
    }
  } catch(e) {
    Logger.log(`ERRO CRÍTICO DURANTE O TESTE: ${e.message}`);
  }
  Logger.log("--- TESTE CONCLUÍDO ---");
}

/**
 * Calcula a distância de Levenshtein entre duas strings.
 * É uma medida da diferença entre duas sequências de caracteres.
 * @param {string} a A primeira string.
 * @param {string} b A segunda string.
 * @returns {number} A distância (número de edições).
 */
function levenshteinDistance(a, b) {
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  const matrix = [];

  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }

  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(matrix[i - 1][j - 1] + 1, Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1));
      }
    }
  }
  return matrix[b.length][a.length];
}
