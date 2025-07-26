// ===============================================
// FUNÇÕES AUXILIARES GERAIS
// (Estas devem estar no TOPO do seu arquivo .gs)
// ===============================================

/**
 * Converte string para camelCase.
 * Ex: "NOME DA COLUNA" -> "nomeDaColuna"
 */
function toCamelCase(str) {
  return String(str).toLowerCase().replace(/[^a-zA-Z0-9]+(.)?/g, (match, chr) => chr ? chr.toUpperCase() : '');
}

/**
 * Retorna a classe CSS para o status do pedido.
 
function getStatusClass(status) {
  switch (String(status).toLowerCase().trim()) {
    case 'concluído': return 'bg-green-100 text-green-800';
    case 'pendente': return 'bg-yellow-100 text-yellow-800';
    case 'cancelado': return 'bg-red-100 text-red-800';
    default: return 'bg-gray-100 text-gray-800';
  }
}*/

/**
 * Função auxiliar para formatar a data para ISO string (YYYY-MM-DD).
 * Útil para serialização ou logs.
 */
function formatarDataParaISO(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return null;
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}-${month}-${day}`;
}


// ===============================================
// FUNÇÕES DE OBTENÇÃO DE DADOS BRUTOS / COMUNS
// (Funções que leem diretamente as planilhas ou interagem com APIs externas)
// ===============================================

/**
 * Função auxiliar para obter os dados de pedidos da planilha e formatá-los,
 * aplicando o filtro de empresa.
 * @param {Object} filters - Objeto de filtros contendo { empresa: { id: '...', nome: '...' } }.
 * @returns {Array<Object>} Uma lista de objetos de pedido formatados.
 */
function _getPedidosDatav2(filters) {
     Logger.log(`[backend] _getPedidosData iniciado com filtros: ${JSON.stringify(filters)}`);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
    if (!sheet) {
        Logger.log('[backend] Planilha "Pedidos" não encontrada.');
        return [];
    }
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    if (values.length < 2) {
        Logger.log('[backend] Planilha "Pedidos" vazia ou contém apenas cabeçalho.');
        return [];
    }
    const headers = values[0];
    const pedidosComEmpresaFiltrada = [];
    const headerMap = {};
    headers.forEach((header, index) => {
        headerMap[toCamelCase(header)] = index;
    });

    const indexEmpresaColuna = headers.findIndex(h => ["EMPRESA", "IDEMPRESA", "IDDAEMPRESA"].includes(h.toUpperCase()));
    if (indexEmpresaColuna === -1) {
        Logger.log('[backend] ERRO CRÍTICO: Coluna da empresa não encontrada na planilha "Pedidos".');
        return [];
    }

    //const chaveEmpresaColuna = toCamelCase(headers[indexEmpresaColuna]);
    const idEmpresaParaFiltrar = filters.empresa && filters.empresa.id != null ? String(filters.empresa.id).trim() : null;
    const idEmpresaParaFiltrarNum = parseInt(idEmpresaParaFiltrar, 10);
    Logger.log(`[backend] Filtrando para o ID da empresa (numérico): ${idEmpresaParaFiltrarNum}`);

    //const pedidoEmpresaValorNaColuna = pedido[chaveEmpresaColuna];
    Logger.log(`[backend] ID da Empresa para filtrar (do localStorage via frontend): "${idEmpresaParaFiltrar}"`);

    for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const idEmpresaNaLinha = row[indexEmpresaColuna];
        const idEmpresaNaLinhaNum = parseInt(idEmpresaNaLinha, 10);

        // Compara os valores e pula para a próxima linha se a empresa for diferente
        if (idEmpresaParaFiltrar && idEmpresaNaLinhaNum !== idEmpresaParaFiltrarNum) {
            continue; 
        }
        const pedido = {};
        for (const key in headerMap) {
            pedido[key] = row[headerMap[key]];
        }if (i < 5) { // Log apenas para as primeiras 5 linhas para não encher o log
  Logger.log(`[backend] _getPedidosDatav2: Pedido ${i} - Objeto mapeado: ${JSON.stringify(pedido)}`);
  Logger.log(`[backend] _getPedidosDatav2: Pedido ${i} - Valor de p.estado: "${pedido.estadoFornecedor}"`); // Verifique se 'estado' é o nome correto
  Logger.log(`[backend] _getPedidosDatav2: Pedido ${i} - Headers mapeados: ${JSON.stringify(Object.keys(headerMap))}`); // Veja todos os campos mapeados
}
        pedido.totalGeral = parseFloat(pedido.totalGeral || 0);

        // Lógica para tratar pedido.itens
        if (typeof pedido.itens === 'string' && pedido.itens.trim() !== '') {
            try {
                const parsedItems = JSON.parse(pedido.itens);
                pedido.itens = Array.isArray(parsedItems) ? parsedItems : [];
            } catch (e) {
                Logger.log(`[backend] Erro ao parsear itens JSON para pedido ${pedido.numeroDoPedido || 'N/A'}: ${e.message}. Valor bruto: "${pedido.itens}"`);
                pedido.itens = [];
            }
        } else if (!Array.isArray(pedido.itens)) {
            pedido.itens = [];
        }

        // Lógica para tratar pedido.data
        if (typeof pedido.data === 'string' && pedido.data.trim() !== '') {
            const parts = pedido.data.match(/(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})/);
            if (parts) {
                const year = parseInt(parts[3], 10);
                const month = parseInt(parts[2], 10) - 1;
                const day = parseInt(parts[1], 10);
                const hour = parseInt(parts[4], 10);
                const minute = parseInt(parts[5], 10);
                const second = parseInt(parts[6], 10);
                const parsedDate = new Date(year, month, day, hour, minute, second);
                pedido.data = (parsedDate instanceof Date && !isNaN(parsedDate.getTime())) ? parsedDate : null;
            } else {
                const defaultParsedDate = new Date(pedido.data);
                pedido.data = (defaultParsedDate instanceof Date && !isNaN(defaultParsedDate.getTime())) ? defaultParsedDate : null;
            }
        } else if (!(pedido.data instanceof Date)) {
            pedido.data = null;
        }
        //const idEmpresaParaFiltrarNum = parseInt(idEmpresaParaFiltrar, 10);
        //const pedidoEmpresaValorNaColunaNum = parseInt(pedidoEmpresaValorNaColuna, 10);
        // A comparação numérica (ex: 1 === 1) funciona de forma confiável
        //if (idEmpresaParaFiltrar && pedidoEmpresaValorNaColunaNum !== idEmpresaParaFiltrarNum) {
        //    continue; // Pula para a próxima linha se a empresa for diferente

        //}    
        pedidosComEmpresaFiltrada.push(pedido);

    }
    Logger.log(`[backend] _getPedidosDatav2 finalizado. Total de pedidos filtrados por empresa: ${pedidosComEmpresaFiltrada.length}`);

    return pedidosComEmpresaFiltrada;

} 

/**
 * Obtém o próximo código sequencial para um novo fornecedor.
 */
function getProximoCodigoFornecedor() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Fornecedores');
  if (!sheet || sheet.getLastRow() < 2) return '0001';

  const codigos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const numeros = codigos.flat()
    .map(c => parseInt(c))
    .filter(n => !isNaN(n));

  const proximo = numeros.length ? Math.max(...numeros) + 1 : 1;
  return proximo.toString().padStart(4, '0');
}

/**
 * Obtém a lista de condições de pagamento da planilha 'Config'.
 */
function getCondicoesPagamento() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
  if (!sheet) {
    Logger.log('ERRO: Planilha "Config" não encontrada! Verifique o nome da aba.');
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const dados = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return dados.flat().filter(e => e !== '');
}

/**
 * Obtém a lista de formas de pagamento da planilha 'Config'.
 */
function getFormasPagamento() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
  if (!sheet) {
    Logger.log('ERRO: Planilha "Config" não encontrada! Verifique o nome da aba.');
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const dados = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  return dados.flat().filter(e => e !== '');
}

/**
 * Obtém a lista de estados (UF e Nome) da planilha 'Config'.
 */
function getEstados() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
  if (!sheet) {
    Logger.log('ERRO: Planilha "Config" não encontrada! Verifique o nome da aba.');
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) { 
    return [];
  }
  
  const dados = sheet.getRange(3, 4, lastRow - 2, 2).getValues();

  return dados
    .filter(([uf, nome]) => uf && String(uf).trim() !== '' && nome && String(nome).trim() !== '') 
    .map(([uf, nome]) => ({
      value: String(uf).trim(),
      text: String(nome).trim()
    }));
}

/**
 * Obtém a lista de fornecedores de uma planilha (apenas os nomes dos ativos).
 * Usado para popular o filtro de fornecedores no dashboard.
 * @returns {Array<string>} Uma lista de nomes de fornecedores.
 */
function getFornecedoresFromSheet() {
  Logger.log('[backend] getFornecedoresFromSheet: Iniciando busca de fornecedores para filtro.');
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('Fornecedores');
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log('[backend] Planilha "Fornecedores" vazia ou não encontrada.');
      return [];
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    
    const indexStatus = headers.findIndex(h => h.toUpperCase() === 'STATUS');
    const indexRazaoSocial = headers.findIndex(h => h.toUpperCase() === 'RAZAO SOCIAL'); 

    if (indexRazaoSocial === -1) {
        Logger.log('[backend] ERRO: Coluna "RAZÃO SOCIAL" não encontrada em Fornecedores. Verifique o cabeçalho.');
        return [];
    }

    const fornecedoresAtivos = values.slice(1) 
      .filter(row => indexStatus === -1 || String(row[indexStatus]).trim().toUpperCase() === 'ATIVO') 
      .map(row => String(row[indexRazaoSocial]).trim()) 
      .filter(name => name); 
    
    Logger.log(`[backend] getFornecedoresFromSheet: Encontrados ${fornecedoresAtivos.length} fornecedores ativos.`);
    return fornecedoresAtivos;

  } catch (e) {
    Logger.log(`[backend] getFornecedoresFromSheet ERRO FATAL: ${e.message}. Stack: ${e.stack}`);
    return [];
  }
}

/**
 * Retorna uma lista completa de TODOS os fornecedores para a tela de gerenciamento.
 */
function getFornecedoresParaGerenciar() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
    if (!sheet || sheet.getLastRow() < 2) {
      return []; 
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const fornecedores = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fornecedor = {};
      headers.forEach((header, index) => {
        fornecedor[toCamelCase(header)] = row[index];
      });
      fornecedores.push(fornecedor);
    }
    return fornecedores;
  } catch (e) {
    Logger.log("Erro em getFornecedoresParaGerenciar: " + e.message);
    return [];
  }
}

/**
 * Obtém a lista de fornecedores com mais detalhes (razao, cnpj, etc.) para outros usos (ex: tela de pedido).
 */
function getFornecedoresList() {
    Logger.log('[backend] getFornecedoresList: Iniciando busca de fornecedores com detalhes.');
    const sheet = SpreadsheetApp.getActive().getSheetByName('Fornecedores');

    if (!sheet || sheet.getLastRow() < 2) return [];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

    const indexStatus = headers.findIndex(h => h.toUpperCase() === 'STATUS');

    const fornecedores = data
        .filter(row => indexStatus === -1 || String(row[indexStatus]).trim().toUpperCase() === 'ATIVO')
        .map(row => {
            const fornecedor = {};
            headers.forEach((header, idx) => {
                fornecedor[toCamelCase(header)] = row[idx];
            });
            fornecedor.codigo = String(fornecedor.codigo || '');
            fornecedor.razao = String(fornecedor.razaoSocial || ''); 
            fornecedor.fantasia = String(fornecedor.nomeFantasia || ''); 
            fornecedor.cnpj = String(fornecedor.cnpj || '');
            fornecedor.endereco = String(fornecedor.endereco || '');
            fornecedor.condicao = String(fornecedor.condicaoDePagamento || ''); 
            fornecedor.forma = String(fornecedor.formaDePagamento || ''); 
            fornecedor.grupo = String(fornecedor.grupo || '');
            fornecedor.estado = String(fornecedor.estado || '');
            fornecedor.cidade = String(fornecedor.cidade || '');
            return fornecedor;
        });
    Logger.log(`[backend] getFornecedoresList: Encontrados ${fornecedores.length} fornecedores ativos com detalhes.`);
    return fornecedores;
}

// Funções de cadastro/edição/exclusão de fornecedores (mantidas como estão no seu código original)
//function salvarFornecedor(fornecedor) { /* ... */ }
//function adicionarOuAtualizarFornecedor(fornecedorObject) { /* ... */ }
//function excluirFornecedor(codigoFornecedor) { /* ... */ }
//function alternarStatusFornecedor(codigoFornecedor) { /* ... */ }
//function consultarCnpj(cnpj) { /* ... */ }


// ===============================================
// FUNÇÕES DE CÁLCULO / ANÁLISE PARA O DASHBOARD
// (Estas devem vir ANTES de generatePurchaseSuggestions e getDashboardData)
// ===============================================

/**
 * Retorna um resumo financeiro dos pedidos.
 * @param {Array<Object>} pedidos - A lista de pedidos JÁ FILTRADA.
 * @returns {Object} Resumo financeiro.
 */
function calculateFinancialSummary(pedidos) {
  let totalGeralPedidos = 0;
  let numeroTotalPedidos = 0;

  pedidos.forEach(pedido => {
    totalGeralPedidos += pedido.totalGeral || 0;
    numeroTotalPedidos++;
  });

  const summary = {
    totalPedidos: numeroTotalPedidos,
    valorTotalPedidos: totalGeralPedidos
  };
  Logger.log(`[backend] calculateFinancialSummary: Resumo financeiro: ${JSON.stringify(summary)}`);
  return summary;
}

/**
 * Retorna os fornecedores mais comprados, do maior para o menor volume.
 * @param {Array<Object>} pedidos - A lista de pedidos JÁ FILTRADA.
 * @returns {Array<Object>} Lista de fornecedores com total comprado.
 */
function calculateTopSuppliers(pedidos) {
  const supplierVolumes = {};
  pedidos.forEach(pedido => {
    const fornecedor = pedido.fornecedor || 'Desconhecido';
    const total = pedido.totalGeral || 0;
    supplierVolumes[fornecedor] = (supplierVolumes[fornecedor] || 0) + total;
  });

  const sortedSuppliers = Object.keys(supplierVolumes).map(fornecedor => ({
    fornecedor: fornecedor,
    totalComprado: supplierVolumes[fornecedor]
  })).sort((a, b) => b.totalComprado - a.totalComprado);

  Logger.log(`[backend] calculateTopSuppliers: Fornecedores por volume (top 5): ${JSON.stringify(sortedSuppliers.slice(0,5))}`);
  return sortedSuppliers;
}

/**
 * Retorna os produtos mais pedidos, do maior para o menor volume (quantidade).
 * @param {Array<Object>} pedidos - A lista de pedidos JÁ FILTRADA.
 * @returns {Array<Object>} Lista de produtos com quantidade total pedida.
 */
function calculateTopProducts(pedidos) {
  const productQuantities = {};
  pedidos.forEach(pedido => {
    (pedido.itens || []).forEach(item => { 
      const produto = item.descricao || 'Produto Desconhecido';
      const quantidade = parseFloat(item.quantidade) || 0;
      productQuantities[produto] = (productQuantities[produto] || 0) + quantidade;
    });
  });

  const sortedProducts = Object.keys(productQuantities).map(produto => ({
    produto: produto,
    totalQuantidade: productQuantities[produto]
  })).sort((a, b) => b.totalQuantidade - a.totalQuantidade);

  Logger.log(`[backend] calculateTopProducts: Produtos por volume (top 5): ${JSON.stringify(sortedProducts.slice(0,5))}`);
  return sortedProducts;
}

/**
 * Calcula os dados mensais para o gráfico.
 * Adapte esta função para como suas datas estão armazenadas e qual período você quer considerar.
 */
function calculateMonthlyAnalysis(pedidos) {
    const monthlyData = {}; 

    pedidos.forEach(pedido => {
        if (!pedido.data || !(pedido.data instanceof Date) || isNaN(pedido.data.getTime())) {
            Logger.log(`[backend] calculateMonthlyAnalysis: Pedido com data inválida ou nula, ignorado para análise mensal. Pedido ID: ${pedido.numeroDoPedido}`);
            return; 
        }

        const yearMonth = `${pedido.data.getFullYear()}-${(pedido.data.getMonth() + 1).toString().padStart(2, '0')}`;
        
        if (!monthlyData[yearMonth]) {
            monthlyData[yearMonth] = { total: 0, count: 0 };
        }
        monthlyData[yearMonth].total += pedido.totalGeral || 0;
        monthlyData[yearMonth].count++;
    });

    const sortedMonths = Object.keys(monthlyData).sort();

    const labels = sortedMonths.map(ym => {
        const [year, month] = ym.split('-');
        const monthNames = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
        return `${monthNames[parseInt(month) - 1]}/${year.slice(2)}`;
    });
    
    const pedidosData = sortedMonths.map(ym => monthlyData[ym].count);
    const gastosData = sortedMonths.map(ym => monthlyData[ym].total);

    Logger.log(`[backend] calculateMonthlyAnalysis: Dados mensais gerados: Labels: ${labels.length}, Pedidos: ${pedidosData.length}, Gastos: ${gastosData.length}`);
    return { labels, pedidosData, gastosData };
}

/**
 * Formata as sugestões do Gemini (que vêm em Markdown) para um formato que o frontend espera.
 * Isso pode envolver parsear o Markdown ou simplesmente dividir em linhas.
 * (COLE ESTA FUNÇÃO AQUI!)
 */
function formatGeminiSuggestionsForFrontend(markdownText) {
    if (!markdownText) return [];

    const lines = markdownText.split('\n').filter(line => line.trim() !== '');

    const formattedSuggestions = [];
    lines.forEach(line => {
        let icon = 'fa-lightbulb';
        let color = 'text-purple-500';
        let text = line;

        text = text.replace(/^-+\s*/, '').trim();
        text = text.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>'); 

        if (text.toLowerCase().includes('estoque')) {
            icon = 'fa-boxes-stacked';
            color = 'text-blue-500';
        } else if (text.toLowerCase().includes('fornecedores')) {
            icon = 'fa-handshake';
            color = 'text-orange-500';
        } else if (text.toLowerCase().includes('produtos')) {
            icon = 'fa-cube';
            color = 'text-green-500';
        } else if (text.toLowerCase().includes('custo') || text.toLowerCase().includes('reduzir gastos')) {
            icon = 'fa-dollar-sign';
            color = 'text-red-500';
        } else if (text.toLowerCase().includes('plano de ação') || text.toLowerCase().includes('próximo passo')) {
            icon = 'fa-clipboard-list';
            color = 'text-gray-600';
        }

        formattedSuggestions.push({
            icon: icon,
            color: color,
            text: text
        });
    });

    return formattedSuggestions;
}
// ===============================================
// FUNÇÃO DE GERAÇÃO DE SUGESTÕES (GEMINI API)
// (Esta deve vir ANTES de getDashboardData)
// ===============================================

/**
 * Gera sugestões de redução de compras usando o Gemini API.
 * @param {Object} pFinancialSummary - Resumo financeiro (objeto).
 * @param {Array<Object>} pTopProducts - Lista de top produtos (array de objetos).
 * @param {Array<Object>} pTopSuppliers - Lista de top fornecedores (array de objetos).
 * @returns {string} Sugestões de redução de compras (texto bruto do Gemini).
 */
function generatePurchaseSuggestions(pFinancialSummary, pTopProducts, pTopSuppliers) { 
  Logger.log('[backend] generatePurchaseSuggestions: Iniciando geração de sugestões...');
  
  pFinancialSummary = pFinancialSummary || {};
  pTopProducts = pTopProducts || [];
  pTopSuppliers = pTopSuppliers || [];

  try {
    const prompt = `
      **PERSONA:** Você é um consultor de supply chain especializado na otimização de compras de mercadorias.

      **CONTEXTO:** Os dados a seguir são de compras emergenciais ("apaga-incêndio") realizadas em fornecedores locais de alto custo. A empresa deseja criar um plano de ação para minimizar a necessidade dessas compras e reduzir a dependência desses fornecedores.

      **TAREFA:** Com base nos dados, forneça de 3 a 5 recomendações estratégicas e acionáveis. Organize suas sugestões nas seguintes categorias: GESTÃO DE ESTOQUE, ESTRATÉGIA DE FORNECEDORES e ANÁLISE DE PRODUTOS. Para cada sugestão, seja direto e indique um próximo passo prático.

      **DADOS DE ENTRADA:**
      - Resumo Financeiro das Compras Locais: ${JSON.stringify(pFinancialSummary)}
      - Top 5 Produtos de Compra Urgente: ${JSON.stringify(pTopProducts.slice(0, 5))}
      - Top 5 Fornecedores Locais Mais Utilizados: ${JSON.stringify(pTopSuppliers.slice(0, 5))}

      **PLANO DE AÇÃO ESTRATÉGICO:**
      `;

    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    
    if (!apiKey) {
        Logger.log('[backend] generatePurchaseSuggestions: Erro: GEMINI_API_KEY não configurada nas Propriedades do Script.');
        return "Erro de configuração: Chave da API Gemini não encontrada. Por favor, configure a 'GEMINI_API_KEY' nas Propriedades do Script.";
    }

    const payload = {
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: {
          temperature: 0.7,
          maxOutputTokens: 500
      }
    };

    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

    let response;
    try {
      response = UrlFetchApp.fetch(apiUrl, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
    } catch (fetchError) {
      Logger.log(`[backend] generatePurchaseSuggestions: Erro ao executar UrlFetchApp.fetch: ${fetchError.message}. Stack: ${fetchError.stack}`);
      return `Erro de rede ou comunicação com a API Gemini: ${fetchError.message}`;
    }

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    Logger.log(`[backend] generatePurchaseSuggestions: Código de Resposta da API: ${responseCode}`); 
    Logger.log(`[backend] generatePurchaseSuggestions: Texto de Resposta da API (bruto): ${responseText}`); 

    if (!responseText || responseText.trim() === '') {
      let errorMessage = `Erro: Resposta vazia ou nula da API Gemini. Código: ${responseCode}.`;
      Logger.log(`[backend] generatePurchaseSuggestions: Erro detalhado (resposta vazia): ${errorMessage}`);
      return `Erro ao gerar sugestões: ${errorMessage}`;
    }

    let result;
    try {
      result = JSON.parse(responseText);
    } catch (parseError) {
      Logger.log(`[backend] generatePurchaseSuggestions: SyntaxError: Erro ao parsear JSON da resposta da API Gemini: ${parseError.message}. Conteúdo bruto: ${responseText.substring(0, 500)}... Stack: ${parseError.stack}`);
      return `Erro ao gerar sugestões: Resposta inválida da API Gemini. Tente novamente ou verifique a chave/permissões.`;
    }
    
    Logger.log(`[backend] generatePurchaseSuggestions: Resposta da API Gemini (parseada): ${JSON.stringify(result)}`);

    if (result.candidates && result.candidates.length > 0 &&
        result.candidates[0].content && result.candidates[0].content.parts &&
        result.candidates[0].content.parts.length > 0) {
      const suggestions = result.candidates[0].content.parts[0].text;
      Logger.log(`[backend] generatePurchaseSuggestions: Sugestões geradas: ${suggestions}`);
      return suggestions;
    } else {
      Logger.log(`[backend] generatePurchaseSuggestions: Resposta da API Gemini não contém sugestões válidas (candidates ausentes): ${JSON.stringify(result)}`);
      if (result.error && result.error.message) {
        return `Erro da API Gemini: ${result.error.message}`;
      }
      return "Não foi possível gerar sugestões no momento. A API retornou uma estrutura inesperada ou incompleta.";
    }

  } catch (e) {
    Logger.log(`[backend] generatePurchaseSuggestions: Erro fatal (fora do bloco fetch/parse): ${e.message}. Stack: ${e.stack}`);
    return `Ocorreu um erro inesperado ao gerar sugestões: ${e.message}`;
  }
}

// ===============================================
// FUNÇÃO PRINCIPAL PARA O DASHBOARD
// (Esta deve ser a ÚLTIMA função principal no seu arquivo antes de doGet)
// ===============================================

/**
 * Função principal para o Dashboard, que coleta todos os dados e gera sugestões,
 * formatando-os para o frontend.
 * @param {Object} filters - Um objeto contendo os filtros (startDate, endDate, supplier, state, empresa).
 * @returns {Object} Objeto contendo todos os dados do dashboard e sugestões, formatado para o frontend.
 */
function getDashboardData(filters) {
  Logger.log('--- [backend] getDashboardData: INICIANDO EXECUÇÃO (vFINAL CORRIGIDA) ---');
  Logger.log(`[backend] getDashboardData: Filtros recebidos: ${JSON.stringify(filters)}`);

  try {
    // 1. Obter TODOS os pedidos já filtrados PELA EMPRESA LOGADA dentro de _getPedidosDatav2
    let finalFilteredPedidos = _getPedidosDatav2(filters); 
    Logger.log(`[backend] getDashboardData: Total de pedidos após filtro por empresa: ${finalFilteredPedidos.length}`);

    // 2. Aplicar FILTROS ADICIONAIS (data, fornecedor, estado) sobre os pedidos já filtrados pela empresa
    if (filters.startDate) {
        const start = new Date(filters.startDate + 'T00:00:00');
        finalFilteredPedidos = finalFilteredPedidos.filter(p => p.data && p.data instanceof Date && p.data >= start);
        Logger.log(`[backend] getDashboardData: Após filtro de data de início (${filters.startDate}): ${finalFilteredPedidos.length} pedidos.`);
    }
    if (filters.endDate) {
        const end = new Date(filters.endDate + 'T23:59:59');
        finalFilteredPedidos = finalFilteredPedidos.filter(p => p.data && p.data instanceof Date && p.data <= end);
        Logger.log(`[backend] getDashboardData: Após filtro de data de fim (${filters.endDate}): ${finalFilteredPedidos.length} pedidos.`);
    }
    if (filters.supplier && filters.supplier !== "") {
        const selectedSupplier = filters.supplier.toLowerCase();
        finalFilteredPedidos = finalFilteredPedidos.filter(p => (p.fornecedor || '').toLowerCase() === selectedSupplier);
        Logger.log(`[backend] getDashboardData: Após filtro de fornecedor (${filters.supplier}): ${finalFilteredPedidos.length} pedidos.`);
    }
    if (filters.state && filters.state !== "") {
        const selectedState = filters.state.toLowerCase();
        finalFilteredPedidos = finalFilteredPedidos.filter(p => (p.estadoFornecedor || '').toLowerCase() === selectedState);
        Logger.log(`[backend] getDashboardData: Após filtro de estado (${filters.state}): ${finalFilteredPedidos.length} pedidos.`);
    }
    Logger.log(`[backend] getDashboardData: Total final de pedidos após TODOS os filtros: ${finalFilteredPedidos.length}`);

    
    //    Garanta que o status na sua planilha seja 'Aprovado'.
    const pedidosParaCalculo = finalFilteredPedidos.filter(p => 
        p.status && String(p.status).trim().toUpperCase() === 'APROVADO'
    );
    Logger.log(`[backend] getDashboardData: Destes, ${pedidosParaCalculo.length} estão 'Aprovados' e serão usados para os cálculos.`);
  
    // 3. Calcular resumos e tops com base nos 'finalFilteredPedidos'
    const financialSummary = calculateFinancialSummary(pedidosParaCalculo);
    const topSuppliers = calculateTopSuppliers(pedidosParaCalculo);
    const topProducts = calculateTopProducts(pedidosParaCalculo);
    const monthlyAnalysisData = calculateMonthlyAnalysis(pedidosParaCalculo);
    
    // Ordena por data decrescente para pegar os mais recentes para a tabela
    const recentOrdersSorted = [...finalFilteredPedidos].sort((a, b) => {
        if (!a.data || !b.data || !(a.data instanceof Date) || !(b.data instanceof Date)) return 0;
        return b.data.getTime() - a.data.getTime();
    });

    const recentOrders = recentOrdersSorted.slice(0, 10).map(p => ({
        id: p.nUmeroDoPedido || p.id || p.numeroPedido || p.númedoDoPedido || 'N/A', 
        supplier: p.fornecedor || 'Desconhecido',
        value: 'R$ ' + (p.totalGeral || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 }),
        status: p.status || 'Desconhecido',
        statusClass: getStatusClass(p.status)
    }));

    // 4. Gerar sugestões da IA
    const iaSuggestionsRaw = generatePurchaseSuggestions(
        financialSummary || {}, // Passando o OBJETO financialSummary
        topProducts || [],     // Passando o ARRAY topProducts
        topSuppliers || []     // Passando o ARRAY topSuppliers
    );
    const iaSuggestionsFormatted = formatGeminiSuggestionsForFrontend(iaSuggestionsRaw);

    Logger.log('[backend] getDashboardData: SUCESSO! Retornando dados formatados para o frontend.');
    
    // Retorna o objeto com todos os dados formatados para o frontend
    return {
      totalGasto: 'R$ ' + (financialSummary.valorTotalPedidos || 0).toFixed(2).replace('.', ','),
      totalPedidos: financialSummary.totalPedidos || 0,
      ticketMedio: financialSummary.totalPedidos > 0 ? 'R$ ' + (financialSummary.valorTotalPedidos / financialSummary.totalPedidos).toFixed(2).replace('.', ',') : 'R$ 0,00',
      
      monthlyLabels: monthlyAnalysisData.labels,
      monthlyPedidosData: monthlyAnalysisData.pedidosData,
      monthlyGastosData: monthlyAnalysisData.gastosData,
      
      topSuppliersLabels: topSuppliers.slice(0,5).map(s => s.fornecedor),
      topSuppliersValues: topSuppliers.slice(0,5).map(s => s.totalComprado),

      recentOrders: recentOrders,
      iaSuggestions: iaSuggestionsFormatted
    };

  } catch (e) {
    Logger.log(`[backend] getDashboardData: ERRO FATAL: ${e.message}. Stack: ${e.stack}`);
    return { 
        status: 'error', 
        message: `Erro ao carregar dados do Dashboard: ${e.message}`,
        totalGasto: 'Erro', totalPedidos: 'Erro', ticketMedio: 'Erro',
        monthlyLabels: [], monthlyPedidosData: [], monthlyGastosData: [],
        topSuppliersLabels: [], topSuppliersValues: [],
        recentOrders: [], iaSuggestions: [{icon: 'fa-exclamation-triangle', color: 'text-red-500', text: `Erro no backend: ${e.message}`}]
    };
  }
}

function getStatusClass(status) {
  if (!status) return 'bg-gray-200 text-gray-800';
  switch (status.toUpperCase()) {
    case 'EM ABERTO':
      return 'bg-blue-100 text-blue-800';
    case 'APROVADO':
      return 'bg-green-100 text-green-800';
    case 'CANCELADO':
      return 'bg-red-100 text-red-800';
    default:
      return 'bg-gray-200 text-gray-800';
  }
}
// ===============================================
// FUNÇÕES DE SERVIÇO HTML (Para doGet e inclusão)
// ===============================================

/**
 * Função para servir o HTML para a interface do usuário.
 * Altere 'Dashboard' para o nome do seu arquivo HTML principal.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Dashboard') // Nome do seu arquivo HTML
      .evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Inclui arquivos CSS e JS no HTML (se você usar arquivos separados).
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename)
      .evaluate()
      .getContent();
}
