/**
 * Busca produtos no catálogo (DB_PRODUTOS) com base em um termo de busca.
 * Procura tanto na descrição quanto na referência do produto.
 * @param {string} searchTerm O texto a ser buscado.
 * @returns {Array<Object>} Um array de objetos de produto, limitado aos 15 primeiros resultados.
 */
function buscarProdutos(searchTerm) {
   try {
    const cache = CacheService.getScriptCache();
    const CACHE_KEY = 'lista_produtos';
    
    // 1. Tenta buscar a lista de produtos do cache
    let produtosJson = cache.get(CACHE_KEY);
    let todosOsProdutos;

    if (produtosJson) {
      // Se encontrou no cache, usa os dados de lá (muito mais rápido)
      Logger.log("Cache HIT: Produtos encontrados na memória temporária.");
      todosOsProdutos = JSON.parse(produtosJson);
    } else {
      // Se não, lê da planilha (operação mais lenta)
      Logger.log("Cache MISS: Lendo produtos da planilha pela primeira vez.");
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_PRODUTOS');
      if (!sheet) throw new Error("Aba 'DB_PRODUTOS' não foi encontrada.");
      
      const data = sheet.getDataRange().getValues();
      const headers = data.shift();
      
      todosOsProdutos = data.map(row => {
          let obj = {};
          headers.forEach((header, i) => {
              let value = row[i];
              if (value instanceof Date) value = value.toISOString();
              obj[header] = value;
          });
          return obj;
      });

      // 2. Salva a lista lida no cache para as próximas buscas (válido por 1 hora)
      cache.put(CACHE_KEY, JSON.stringify(todosOsProdutos), 3600); 
    }

    // 3. Filtra a lista (do cache ou da planilha) com o termo da busca
    const searchTermLower = searchTerm.toLowerCase();
    const results = todosOsProdutos.filter(p => {
        const descricao = String(p.DESCRICAO_PRODUTO || '').toLowerCase();
        const ref = String(p.REF_PRODUTO || '').toLowerCase();
        return descricao.includes(searchTermLower) || ref.includes(searchTermLower);
    }).slice(0, 15); // Pega apenas os 15 primeiros resultados

    Logger.log(`Busca por "${searchTerm}" encontrou ${results.length} resultados.`);
    return results;

  } catch (e) {
    Logger.log(`Erro em buscarProdutos (com cache): ${e.stack}`);
    throw new Error(`Erro ao buscar produtos: ${e.message}`);
  }
}


/**
 * Verifica se um produto já existe pelo nome ou referência. Se não existir, cadastra-o automaticamente.
 * Retorna o objeto do produto (existente ou recém-criado) com seu ID_PRODUTO padronizado.
 * @param {Object} dadosProduto - Um objeto com {descricao, ref, unidade}.
 * @returns {Object} Um objeto representando a linha do produto na planilha DB_PRODUTOS.
 */
function verificarOuCadastrarProduto(dadosProduto) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_PRODUTOS');
    if (!sheet) {
      throw new Error("A aba 'DB_PRODUTOS' não foi encontrada.");
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxDesc = headers.indexOf('DESCRICAO_PRODUTO');
    const idxRef = headers.indexOf('REF_PRODUTO');

    const descBusca = dadosProduto.descricao.trim().toLowerCase();
    const refBusca = dadosProduto.ref.trim().toLowerCase();

    // Procura por um produto existente
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const descPlanilha = String(row[idxDesc]).trim().toLowerCase();
      const refPlanilha = String(row[idxRef]).trim().toLowerCase();

      if ((refBusca && refPlanilha === refBusca) || (descPlanilha === descBusca)) {
        console.log(`Produto encontrado.`);
        return getObjectFromRow(row, headers);
      }
    }

    // Se chegou aqui, o produto é novo.
    console.log(`Produto não encontrado. Cadastrando e limpando o cache.`);
    
    const novoId = getProximoIdProduto(data);
    const dataCadastro = new Date();
    
    // ✅ CORREÇÃO APLICADA AQUI: A declaração duplicada foi removida.
    // Usamos o método robusto que itera sobre os cabeçalhos.
    const novaLinha = [];
    headers.forEach(header => {
        switch(header) {
            case 'ID_PRODUTO':
                novaLinha.push(novoId);
                break;
            case 'DESCRICAO_PRODUTO':
                novaLinha.push(dadosProduto.descricao.trim().toUpperCase()); // Padroniza para maiúsculas
                break;
            case 'REF_PRODUTO':
                novaLinha.push(dadosProduto.ref.trim().toUpperCase()); // Padroniza para maiúsculas
                break;
            case 'UNIDADE_MEDIDA':
                novaLinha.push(dadosProduto.unidade);
                break;
            case 'DATA_CADASTRO':
                novaLinha.push(dataCadastro);
                break;
            // Adicione outras colunas do seu catálogo aqui se necessário
            case 'CATEGORIA_PRODUTO':
                novaLinha.push(''); // Exemplo: Deixa a categoria em branco para ser preenchida depois
                break;
            default:
                novaLinha.push(''); // Garante que a linha terá o mesmo número de colunas que o cabeçalho
        }
    });

    sheet.appendRow(novaLinha);
    
    // Limpa o cache para forçar a releitura na próxima busca.
    const cache = CacheService.getScriptCache();
    cache.remove('lista_produtos');
    
    // Retorna o objeto do produto recém-criado
    return getObjectFromRow(novaLinha, headers);

  } catch (e) {
    console.error(`Erro em verificarOuCadastrarProduto: ${e.message}`);
    throw new Error(`Erro ao processar produto: ${e.message}`);
  }
}

/**
 * Função auxiliar para gerar o próximo ID_PRODUTO sequencial.
 * Ex: Se o último for P0123, o próximo será P0124.
 */
function getProximoIdProduto(data) {
  const idxId = data[0].indexOf('ID_PRODUTO');
  if (data.length <= 1) { // Apenas o cabeçalho
    return 'PD_001';
  }
  
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const idAtualStr = data[i][idxId];
    if (idAtualStr && typeof idAtualStr === 'string' && idAtualStr.startsWith('PD')) {
      const idNum = parseInt(idAtualStr.substring(3), 10);
      if (idNum > maxId) {
        maxId = idNum;
      }
    }
  }
  
  const novoNumero = maxId + 1;
  const novoId = 'PD_' + String(novoNumero).padStart(3, '0'); // Garante PD_001, PD_002, etc.
  return novoId;
}
//_ Adicione esta função auxiliar genérica se ainda não a tiver
function getObjectFromRow(row, headers) {
    const obj = {};
  headers.forEach((header, i) => {
    if (header) { // Ignora colunas que não têm um cabeçalho
      let value = row[i];
      
      // ✅ CORREÇÃO APLICADA AQUI
      // Se o valor for um objeto de Data, converte para uma string no formato ISO.
      if (value instanceof Date) {
        value = value.toISOString();
      }
      
      obj[header] = value;
    }
  });
  return obj;
}

function testarPerformanceDaBusca() {
  Logger.log("--- INICIANDO TESTE DE PERFORMANCE DA BUSCA DE PRODUTOS ---");

  // 1. Limpa o cache para garantir que a primeira busca seja "fria" (lendo da planilha)
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('lista_produtos');
    Logger.log("✅ Cache 'lista_produtos' limpo para iniciar o teste.");
  } catch (e) {
    Logger.log("⚠️ Não foi possível limpar o cache: " + e.message);
  }

  // --- PRIMEIRA BUSCA (CACHE MISS) ---
  Logger.log("\n--- Executando a PRIMEIRA busca (deve ser a mais LENTA) ---");
  console.time("Tempo da Primeira Busca (Cache Miss)"); // Inicia o cronômetro

  const resultados1 = buscarProdutos("RV MANUAL VW DELIVERY"); // Use um termo de busca comum nos seus produtos

  console.timeEnd("Tempo da Primeira Busca (Cache Miss)"); // Para o cronômetro e exibe o tempo
  Logger.log(`Encontrados ${resultados1.length} resultados.`);

  // --- SEGUNDA BUSCA (CACHE HIT) ---
  Logger.log("\n--- Executando a SEGUNDA busca (deve ser INSTANTÂNEA) ---");
  console.time("Tempo da Segunda Busca (Cache Hit)"); // Inicia um novo cronômetro

  const resultados2 = buscarProdutos("CAPA DE PORCA 27"); // Use outro termo de busca comum

  console.timeEnd("Tempo da Segunda Busca (Cache Hit)"); // Para o cronômetro
  Logger.log(`Encontrados ${resultados2.length} resultados.`);
  
  Logger.log("\n--- TESTE DE PERFORMANCE CONCLUÍDO ---");
}
