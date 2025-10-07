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
 * Formata um número para o padrão numérico brasileiro.
 * Ex: 12345.67 -> "12.345,67"
 * @param {number} value O número a ser formatado.
 * @return {string} O número formatado como string.
 */
function _formatNumberBRL(value) {
  const number = parseFloat(value) || 0;
  return new Intl.NumberFormat('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(number);
}

/**
 * Converte uma string de moeda BRL para um número float puro.
 * Ex: "R$ 1.234,56" -> 1234.56
 * @param {string|number} currencyValue O valor a ser convertido.
 * @return {number} O valor como um número float.
 */
function _sanitizeCurrency(currencyValue) {
  if (currencyValue === null || currencyValue === '' || typeof currencyValue === 'undefined') {
    return 0;
  }
  if (typeof currencyValue === 'number') {
    return currencyValue; // Já é um número, não faz nada
  }

  const stringValue = String(currencyValue)
    .replace("R$", "")      // Remove o símbolo de real
    .trim()                 // Remove espaços extras
    .replace(/\./g, '')     // Remove o separador de milhar (.)
    .replace(",", ".");     // Troca a vírgula decimal por ponto

  const numberValue = parseFloat(stringValue);

  return isNaN(numberValue) ? 0 : numberValue; // Se ainda falhar, retorna 0
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
      - Top 5 Produtos de Compra Urgente: ${JSON.stringify(pTopProducts.slice(0, 10))}
      - Top 5 Fornecedores Locais Mais Utilizados: ${JSON.stringify(pTopSuppliers.slice(0, 10))}

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

/**
 * VERSÃO APRIMORADA: Gera uma análise estratégica completa da operação de compras nacional.
 * Esta função é autossuficiente: ela busca os dados necessários, monta um prompt rico e chama a IA.
 * @returns {string} O texto da análise estratégica gerado pela IA.
 */
function gerarAnaliseEstrategica(filters) {
  Logger.log('[backend] gerarAnaliseEstrategica (versão completa): Iniciando...');

  try {
    // --- ETAPA 1: COLETAR DADOS CONTEXTUAIS ---
        // Chama a nossa nova função central para obter os dados já filtrados e aprovados.
        const pedidosParaAnalise = _getFilteredPedidos(filters);

        if (pedidosParaAnalise.length === 0) {
            Logger.log("[backend] Nenhum pedido aprovado encontrado com os filtros atuais. Retornando análise vazia.");
            return JSON.stringify({ analiseEstrategica: [] });
        }

    Logger.log("Coletando dados para a análise...");
    // Calcula os dados agregados necessários para o prompt a partir dos dados filtrados.    
    const produtosRanqueados = getProductsRankedByValue(pedidosParaAnalise);
    const analisePorEstado = getAnalysisByState(pedidosParaAnalise);
    
    // Busca o histórico do produto mais importante
    const produtoPrincipal = produtosRanqueados.length > 0 ? produtosRanqueados[0] : null;
    let historicoProdutoPrincipal = {};
    if (produtoPrincipal) {
    historicoProdutoPrincipal = getPurchaseHistoryForItem(produtoPrincipal.descricao);  
    }
    
    // --- ETAPA 2: MONTAR O PROMPT COMPLETO ---
      const prompt = `
        **PERSONA:** Você é um consultor sênior de Supply Chain e Estratégia Tributária, especializado em compras no mercado brasileiro, com profundo conhecimento da legislação fiscal estadual, ICMS ST, e impactos de compras interestaduais e estaduais para empresas no regime de Lucro Presumido.

        **CONTEXTO:** Você está analisando os dados de compra de uma empresa com sede na Bahia (BA), operando sob regime de Lucro Presumido. As compras da empresa incluem tanto fornecedores **dentro da Bahia** quanto de **outros estados**.

        Considere que aproximadamente **90% dos produtos comprados são sujeitos ao regime de ICMS Substituição Tributária (ICMS ST)**, com exceção das **ferramentas manuais** (como chave de fenda, alicate, etc), que são tratadas no sistema pela descrição e **não estão sujeitas à ST**. Essa distinção deve ser considerada ao avaliar:

        - O impacto da **margem de valor agregado (MVA)** e da **base de cálculo presumida** nas compras interestaduais
        - A viabilidade de **compras internas na Bahia** para evitar o pagamento do ICMS ST
        - O redirecionamento de fornecedores para produtos **não sujeitos à ST**, visando aproveitar alíquotas interestaduais mais vantajosas

        O objetivo é identificar oportunidades de:

        - Redução de custos totais (produto + impostos)
        - Otimização fiscal (principalmente ICMS ST e diferencial de alíquota)
        - Consolidação e redirecionamento geográfico da base de fornecedores
        - Mitigação da volatilidade de preços de itens estratégicos

        Avalie também se há **possibilidades de substituição de fornecedores por alternativas mais vantajosas geograficamente**, sem comprometer o abastecimento (considerando lead time, confiabilidade logística e histórico de entrega).

        Ao identificar produtos não sujeitos à ST (ex: ferramentas manuais), priorize a análise de **diferença de alíquota interestadual** e **potenciais créditos de ICMS** que possam ser aproveitados.

        **TAREFA:** Forneça exatamente 6 recomendações (2 por categoria). Organize suas sugestões nas seguintes categorias:
        - ANÁLISE TRIBUTÁRIA E GEOGRÁFICA
        - ESTRATÉGIA DE FORNECEDORES
        - GESTÃO DE ESTOQUE E PRODUTOS

         **FORMATO DE SAÍDA OBRIGATÓRIO E ESTRITO:**
        - Use '@@@' para separar CADA recomendação completa.
        - Dentro de cada recomendação, use '|||' para separar a Categoria, a Recomendação, a Justificativa e o Próximo Passo, EXATAMENTE NESTA ORDEM.
        - NÃO use JSON. NÃO use Markdown. NÃO use títulos ou listas. Apenas texto puro com os separadores '@@@' e '|||'.
        - Para cada uma das categorias listadas, é OBRIGATÓRIO fornecer pelo menos duas recomendação completa.
        - NÃO REPITA as categorias na sua resposta.
        - Se por algum motivo você não tiver dados suficientes para analisar uma categoria, preencha a recomendação com o texto "Não há dados suficientes para uma análise aprofundada nesta categoria, tentarei na próxima vez trazer o retorno desejado!".

        Para cada recomendação:
        - Justifique com base nos dados
        - Indique claramente um **próximo passo prático**      
        - Formate todos os valores monetários no padrão brasileiro (exemplo: R$ 1.234,56)
        - Quando aplicável, indique: “*A compra do produto X seria mais vantajosa se realizada a partir do estado Y, considerando custo final com impostos.*”

        **EXEMPLO DE RESPOSTA:**
        ANÁLISE TRIBUTÁRIA E GEOGRÁFICA|||Avaliar a compra do item X do estado de MG.|||O ICMS em MG é menor para este produto.|||Contatar 3 fornecedores em MG para cotação.@@@ESTRATÉGIA DE FORNECEDORES|||Consolidar compras dos itens A e B em um único fornecedor.|||A compra em volume gera um desconto de 5%.|||Iniciar negociação com o Fornecedor Y para um contrato anual.

        **DADOS PARA ANÁLISE:**
        Estados: ${JSON.stringify(analisePorEstado, null, 2)}
        Top Products: ${JSON.stringify(produtosRanqueados.slice(0, 10), null, 2)}

        **PLANO DE AÇÃO ESTRATÉGICO:**
        `;

    // --- ETAPA 3: CHAMAR A API GEMINI ---
    const respostaDaIA = callGeminiAPI(prompt); // Usa a função que já criamos
    Logger.log("[backend] Resposta em texto recebida da IA: " + respostaDaIA);

    // --- ETAPA 4: PROCESSAR O TEXTO E CONSTRUIR O JSON
    const analiseFinal = { analiseEstrategica: [] };
        const mapaCategorias = {}; // Usado para agrupar recomendações por categoria

        const recomendacoesIndividuais = respostaDaIA.split('@@@');

        recomendacoesIndividuais.forEach(recTexto => {
            if (recTexto.trim() === '') return; // Pula blocos vazios

            const partes = recTexto.split('|||');
            if (partes.length === 4) {
                const [categoria, recomendacao, justificativa, proximoPasso] = partes.map(p => p.trim());

                const novaRecomendacao = {
                    recomendacao: recomendacao,
                    justificativa: justificativa,
                    proximoPasso: proximoPasso
                };

                // Se a categoria ainda não foi vista, crie-a no nosso objeto final
                if (!mapaCategorias[categoria]) {
                    mapaCategorias[categoria] = {
                        categoria: categoria,
                        recomendacoes: []
                    };
                    analiseFinal.analiseEstrategica.push(mapaCategorias[categoria]);
                }
                
                // Adicione a nova recomendação à categoria correspondente
                mapaCategorias[categoria].recomendacoes.push(novaRecomendacao);
            }
        });

        Logger.log("[backend] Objeto JSON construído com sucesso no código.");
        return JSON.stringify(analiseFinal);

    } catch (e) {
        Logger.log(`[backend] Erro fatal em gerarAnaliseEstrategica: ${e.message}. Stack: ${e.stack}`);
        return `Ocorreu um erro no servidor ao processar a análise: ${e.message}`;
    }
}

/**
 * Agrega os dados de compra por estado do fornecedor.
 * Retorna um resumo de valor total comprado, ICMS ST pago e contagem de fornecedores/pedidos.
 * Depende que a aba 'pedidos' tenha uma coluna com o ID do Fornecedor para ligar as informações.
 */
function getAnalysisByState(pedidosAprovados) {
   try {
        Logger.log(`[getAnalysisByState] Iniciando análise por estado com ${pedidosAprovados.length} pedidos aprovados.`);

        const analise = {};

        // NOVO: Itera sobre os pedidos já filtrados. Não precisa mais ler planilhas.
        pedidosAprovados.forEach(pedido => {
            const estado = pedido.estadoFornecedor; // Pega o estado diretamente do objeto pedido

            if (estado) {
                if (!analise[estado]) {
                    analise[estado] = {
                        uf: estado,
                        valorTotalComprado: 0,
                        valorTotalIcmsSt: 0,
                        pedidos: new Set(),
                        fornecedores: new Set()
                    };
                }
                
                // Agrega os valores totais de todos os itens daquele pedido
                let totalPedido = 0;
                let icmsStPedido = 0;
                if(pedido.itens && Array.isArray(pedido.itens)){
                   pedido.itens.forEach(item => {
                       totalPedido += _sanitizeCurrency(item.totalItem);
                       icmsStPedido += _sanitizeCurrency(item.icmsStTotal); // Supondo que cada item tenha seu ICMS
                   });
                }

                analise[estado].valorTotalComprado += totalPedido;
                analise[estado].valorTotalIcmsSt += icmsStPedido;
                analise[estado].pedidos.add(pedido.nUmeroDoPedido); // Use a chave correta para o número do pedido
                analise[estado].fornecedores.add(pedido.fornecedor);
            }
        });

        const resultadoFinal = Object.values(analise).map(res => ({
            uf: res.uf,
            valorTotalComprado: res.valorTotalComprado,
            valorTotalIcmsSt: res.valorTotalIcmsSt,
            numeroDePedidos: res.pedidos.size,
            numeroDeFornecedores: res.fornecedores.size
        })).sort((a, b) => b.valorTotalComprado - a.valorTotalComprado);

        Logger.log("[getAnalysisByState] Análise por estado concluída. Estados encontrados: " + resultadoFinal.length);
        return resultadoFinal;

    } catch (e) {
        Logger.log(`ERRO em getAnalysisByState: ${e.message}`);
        return [];
    }
}

/**
 * Busca o histórico de todas as compras de um produto específico.
 * @param {string} idProduto O ID do produto a ser pesquisado (ex: "PD_018").
 * @returns {Array<Object>} Um array com o histórico de compras do item, ordenado por data.
 */
function getPurchaseHistoryForItem(idProduto) {
  try {
    Logger.log(`Buscando histórico para o produto ID: ${idProduto}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetItens = ss.getSheetByName("itens pedido");
    const sheetPedidos = ss.getSheetByName("pedidos");

    if (!sheetItens || !sheetPedidos) {
      throw new Error("As abas 'itens pedido' ou 'pedidos' não foram encontradas.");
    }
    
    // 1. Mapeia Número do Pedido para o Nome do Fornecedor e a Data
    const pedidoInfo = {};
    const pedidosData = sheetPedidos.getDataRange().getValues();
    const headersPed = pedidosData.shift();
    const idxNumPed = headersPed.indexOf('Número do Pedido');
    const idxFornNome = headersPed.indexOf('Fornecedor');
    const idxData = headersPed.indexOf('Data'); // Assumindo que a data principal está na aba 'pedidos'
    
    pedidosData.forEach(row => {
      pedidoInfo[row[idxNumPed]] = {
        fornecedor: row[idxFornNome],
        data: row[idxData]
      };
    });

    // 2. Filtra os itens pelo ID do produto e monta o histórico
    const historico = [];
    const itensData = sheetItens.getDataRange().getValues();
    const headersItens = itensData.shift();
    const idxIdProdItem = headersItens.indexOf('ID_PRODUTO'); // Assumindo que você tem um ID
    const idxDescricaoItem = headersItens.indexOf('DESCRICAO'); // Usado como fallback
    const idxNumPedItem = headersItens.indexOf('NUMERO PEDIDO');
    const idxQtd = headersItens.indexOf('QUANTIDADE');
    const idxPreco = headersItens.indexOf('PRECO UNITARIO');
    
    itensData.forEach(row => {
      // Procura pelo ID do produto. Se não tiver, usa a descrição como fallback
      const produtoIdentifier = row[idxIdProdItem] || row[idxDescricaoItem];
      
      if (produtoIdentifier === idProduto) {
        const numPedido = row[idxNumPedItem];
        const info = pedidoInfo[numPedido];
        
        if (info) {
          historico.push({
            data: info.data,
            fornecedor: info.fornecedor,
            quantidade: row[idxQtd],
            precoUnitario: row[idxPreco]
          });
        }
      }
    });
    
    // 3. Ordena o histórico pela data, do mais recente para o mais antigo
    historico.sort((a, b) => new Date(b.data) - new Date(a.data));
    
    Logger.log(`Histórico encontrado para ${idProduto}: ${historico.length} registros.`);
    return historico.slice(0, 10); // Retorna os 10 registros mais recentes para não sobrecarregar o prompt

  } catch (e) {
    Logger.log(`ERRO em getPurchaseHistoryForItem: ${e.message}`);
    return [];
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

      let totalIcmsSt = 0;
    pedidosParaCalculo.forEach(p => {
        totalIcmsSt += _sanitizeCurrency(p.icmsStTotal); 
    });
    totalIcmsSt = parseFloat(totalIcmsSt.toFixed(2));
    Logger.log(`[backend] getDashboardData: Total de ICMS ST calculado: ${totalIcmsSt}`);

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
        id: p.nUmeroDoPedido || p.id || p.numeroPedido || p.numeroDoPedido || 'N/A', 
        supplier: p.fornecedor || 'Desconhecido',
        value: 'R$ ' + (p.totalGeral || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 }),
        icmsStTotal: p.icmsStTotal || 0, // Adicionado para a tabela
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
      totalGasto: 'R$ ' + _formatNumberBRL(financialSummary.valorTotalPedidos),
      totalPedidos: financialSummary.totalPedidos || 0,
      ticketMedio: 'R$ ' + _formatNumberBRL(financialSummary.valorTotalPedidos / financialSummary.totalPedidos),
      totalIcmsSt: (totalIcmsSt || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }),

      monthlyLabels: monthlyAnalysisData.labels,
      monthlyPedidosData: monthlyAnalysisData.pedidosData,
      monthlyGastosData: monthlyAnalysisData.gastosData,
      
      topSuppliersLabels: topSuppliers.slice(0,10).map(s => s.fornecedor),
      topSuppliersValues: topSuppliers.slice(0,10).map(s => s.totalComprado),

      // ===== DADOS ATUALIZADOS PARA O GRÁFICO DE PRODUTOS =====
      topProductsLabels: topProducts.slice(0, 10).map(p => p.produto),
      topProductsValues: topProducts.slice(0, 10).map(p => p.totalQuantidade), // Usa a propriedade correta 'totalQuantidade'

      recentOrders: recentOrders,
      iaSuggestions: iaSuggestionsFormatted
    };

  } catch (e) {
    Logger.log(`[backend] getDashboardData: ERRO FATAL: ${e.message}. Stack: ${e.stack}`);
    return { 
        status: 'error', 
        message: `Erro ao carregar dados do Dashboard: ${e.message}`,
        totalIcmsSt: 'Erro', // Adicionado para o caso de erro
        totalGasto: 'Erro', totalPedidos: 'Erro', ticketMedio: 'Erro',
        monthlyLabels: [], monthlyPedidosData: [], monthlyGastosData: [],
        topSuppliersLabels: [], topSuppliersValues: [],
        recentOrders: [], iaSuggestions: [{icon: 'fa-exclamation-triangle', color: 'text-red-500', text: `Erro no backend: ${e.message}`}]
    };
  }
}

function getStatusClass(status) {
        const statusNormalizado = (status || '').toUpperCase().replace(' ', '');
        switch (statusNormalizado) {
            case 'APROVADO':
                return 'bg-green-100 text-green-800';
            case 'CANCELADO':
                return 'bg-red-100 text-red-800';
            case 'RASCUNHO':
                return 'bg-yellow-100 text-yellow-800';
            case 'AGUARDANDOAPROVACAO':
                return 'bg-blue-100 text-blue-800';
            case 'EMABERTO':
                return 'bg-purple-100 text-purple-800';
            default:
                return 'bg-gray-100 text-gray-800';
        }
    }

    /**
 * Analisa a planilha 'itens pedido', agrupa os itens por descrição de produto,
 * soma o valor total comprado de cada um e retorna uma lista ranqueada.
 * @returns {Array<Object>} Um array de objetos, cada um com {descricao, valorTotal},
 * ordenado do maior valorTotal para o menor.
 */
function getProductsRankedByValue(pedidosAprovados) {
  try {
    Logger.log(`[getProductsRankedByValue] Iniciando ranking com ${pedidosAprovados.length} pedidos aprovados.`);

    const produtosAgregados = {};

    // NOVO: Itera sobre os pedidos já filtrados que foram recebidos como argumento
    pedidosAprovados.forEach(pedido => {
      // Supondo que seu objeto 'pedido' tem uma propriedade 'itens' que é um array
      if (pedido.itens && Array.isArray(pedido.itens)) {
        pedido.itens.forEach(item => {
          const descricao = item.descricao;
          const valor = _sanitizeCurrency(item.totalItem); // USA A FUNÇÃO DE HIGIENIZAÇÃO

          if (descricao && valor > 0) {
            if (produtosAgregados[descricao]) {
              produtosAgregados[descricao] += valor;
            } else {
              produtosAgregados[descricao] = valor;
            }
          }
        });
      }
    });

    const resultadoArray = Object.keys(produtosAgregados).map(key => {
      return {
        descricao: key,
        valorTotal: produtosAgregados[key]
      };
    });

    resultadoArray.sort((a, b) => b.valorTotal - a.valorTotal);
    
    Logger.log(`[getProductsRankedByValue] Agregação concluída. ${resultadoArray.length} produtos únicos encontrados.`);
    
    return resultadoArray;

  } catch (e) {
    Logger.log(`ERRO em getProductsRankedByValue: ${e.message}`);
    return [];
  }
}
/**
 * Orquestra a criação da análise de Curva ABC.
 * 1. Busca os dados ranqueados dos produtos.
 * 2. Monta o prompt para a IA.
 * 3. (Simula) a chamada à IA e retorna a resposta JSON estruturada.
 * @returns {string} Uma string contendo o JSON da análise da IA.
 */
function gerarAnaliseABC_comIA(filters) {
  Logger.log("Iniciando geração da análise ABC com chamada real à IA...");
  
  try {
    const pedidosParaAnalise = _getFilteredPedidos(filters);
    const produtosRanqueados = getProductsRankedByValue(pedidosParaAnalise);

    if (!produtosRanqueados || produtosRanqueados.length === 0) {
      return JSON.stringify({ insights: {ponto1: "Nenhum dado de produto para analisar."}, chartData: {} });
    }

    const promptCurvaABC = `
      **PERSONA:** Você é um analista de dados especialista em gestão de inventário e supply chain.

      **TAREFA:** A partir da lista dos top 15 produtos ranqueados por valor total de compra, gere uma análise completa de Curva ABC. Sua saída deve ser **exclusivamente um objeto JSON bem-formado**, sem nenhum texto ou explicação adicional antes ou depois. O objeto JSON deve conter duas chaves principais: "insights" e "chartData".

      1. **Na chave "insights"**: Forneça um objeto com três chaves ("ponto1", "ponto2", "ponto3"), cada uma contendo uma frase concreta e clara sobre a análise da Curva ABC considerando somente os top 15 produtos. Utilize o critério padrão ABC, onde:
        - Classe A: itens que juntos representam aproximadamente 70% do valor acumulado.
        - Classe B: próximos 20%.
        - Classe C: últimos 10%.

        As frases devem incluir, por exemplo:
        - Quantidade de itens em cada classe.
        - Percentual do valor total representado pela Classe A.
        - Relevância da classe C em termos de número de itens vs. valor.

        Os valores monetários devem estar formatados no padrão brasileiro PT-BR (exemplo: "1.234,56").

      2. **Na chave "chartData"**: Gere os dados para um gráfico de Pareto no formato do Chart.js, considerando os mesmos 15 produtos:
        - "labels": array com as descrições dos top 15 produtos.
        - "datasets": array com dois objetos:
            - Primeiro objeto (barras): chave "data" contendo valores monetários puros (números) de compra de cada produto.
            - Segundo objeto (linha): chave "data" contendo o percentual acumulado para cada produto (valores entre 0 e 100).

      **DADOS DE ENTRADA:**
      - Lista de Produtos Ranqueados: ${JSON.stringify(produtosRanqueados)}

      **OBSERVAÇÕES IMPORTANTES:**
      - Certifique-se de que os produtos estão ordenados em ordem decrescente de valor total de compra antes de gerar a análise.
      - Caso a lista contenha menos de 15 produtos, aplique a análise com os itens disponíveis.
      - Os valores monetários devem estar formatados no padrão brasileiro PT-BR (exemplo: "1.234,56").
      - Os percentuais acumulados devem ser arredondados para uma casa decimal.

      **SAÍDA (APENAS O OBJETO JSON):**
  `;
    
    Logger.log("Chamando a IA para gerar a análise...");
    const respostaDaIA = callGeminiAPI(promptCurvaABC);
    
    const jsonLimpo = respostaDaIA.replace(/```json/g, "").replace(/```/g, "").trim();
    
    Logger.log("Análise da IA recebida. Retornando para o frontend.");
    return jsonLimpo;

  } catch (e) {
    Logger.log(`Erro ao chamar a IA: ${e.message}`);
    return JSON.stringify({ 
      insights: { ponto1: `Erro ao gerar análise: ${e.message}` },
      chartData: {} 
    });
  }
}

/**
 * FUNÇÃO CENTRAL E REUTILIZÁVEL
 * Busca todos os pedidos e aplica um conjunto de filtros sobre eles.
 * Esta é a única função que acessa os dados brutos e os filtra.
 * @param {object} filters Objeto contendo os filtros da tela (startDate, endDate, etc.).
 * @return {Array} Um array de objetos, onde cada objeto é um pedido que passou por todos os filtros.
 */
function _getFilteredPedidos(filters) {
    Logger.log(`[backend] _getFilteredPedidos: Iniciando busca com filtros: ${JSON.stringify(filters)}`);

    const filtrosParaV2 = {
        // Copia os outros filtros que possam existir (startDate, endDate, etc.)
        startDate: filters.startDate,
        endDate: filters.endDate,
        supplier: filters.supplier,
        state: filters.state
    };
    
    // A tradução principal: converte { empresaId: "001" } para { empresa: { id: "001" } }
    if (filters && filters.empresaId) {
        filtrosParaV2.empresa = { id: filters.empresaId };
    }
    
    Logger.log(`[ADAPTADOR _getFilteredPedidos] Filtro traduzido para: ${JSON.stringify(filtrosParaV2)}`);

    // 1. Obter TODOS os pedidos já filtrados PELA EMPRESA LOGADA
    let pedidos = _getPedidosDatav2(filtrosParaV2); // Supondo que esta função já lida com o filtro de empresa.

    // 2. Aplica os filtros recebidos do front-end
    if (filters && filters.startDate) {
        const start = new Date(filters.startDate + 'T00:00:00');
        pedidos = pedidos.filter(p => p.data && p.data instanceof Date && p.data >= start);
    }
    if (filters && filters.endDate) {
        const end = new Date(filters.endDate + 'T23:59:59');
        pedidos = pedidos.filter(p => p.data && p.data instanceof Date && p.data <= end);
    }
    if (filters && filters.supplier) {
        pedidos = pedidos.filter(p => (p.fornecedor || '').toLowerCase() === filters.supplier.toLowerCase());
    }
    if (filters && filters.state) {
        pedidos = pedidos.filter(p => (p.estadoFornecedor || '').toLowerCase() === filters.state.toLowerCase());
    }

    // 3. Aplica o filtro de status "APROVADO" que é padrão para o dashboard
    const pedidosAprovados = pedidos.filter(p => 
        p.status && String(p.status).trim().toUpperCase() === 'APROVADO'
    );
    
    Logger.log(`[backend] _getFilteredPedidos: Retornando ${pedidosAprovados.length} pedidos APROVADOS e filtrados.`);
    return pedidosAprovados;
}

/**
 * Testa o fluxo da Análise ABC, verificando se o filtro de empresa é aplicado.
 */
function _TESTE_filtragemDaAnaliseABC() {
  Logger.log("--- 🔬 INICIANDO TESTE DE FILTRAGEM DA ANÁLISE ABC ---");

  // Simula os filtros que o front-end enviaria, incluindo um ID de empresa específico
  const mockFilters = {
    empresaId: "003" // << Altere para um ID de empresa que você queira testar
  };

  Logger.log(`Testando com o filtro: ${JSON.stringify(mockFilters)}`);

  // Chama a função principal da Análise ABC
  gerarAnaliseABC_comIA(mockFilters);

  Logger.log("--- 🔬 TESTE DE FILTRAGEM CONCLUÍDO ---");
  Logger.log("--> Verifique os logs abaixo para ver o resultado do filtro de empresa.");
}

/**
 * Função principal que envia um prompt para a API do Gemini.
 * @param {string} prompt O texto do prompt a ser enviado.
 * @returns {string} A resposta de texto da IA.
 */
function callGeminiAPI(prompt) {
  const apiKey = _getGeminiApiKey();
  // Usando um modelo poderoso e atualizado
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey;

  const payload = {
    "contents": [{
      "parts": [{ "text": prompt }]
    }],
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    Logger.log("Enviando prompt para a API do Gemini...");
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      const textResponse = jsonResponse.candidates[0].content.parts[0].text;
      Logger.log("Resposta da IA recebida com sucesso.");
      return textResponse;
    } else {
      Logger.log(`Erro da API Gemini: Status ${responseCode}, Resposta: ${responseBody}`);
      throw new Error(`Erro na chamada da API: ${responseBody}`);
    }
  } catch (e) {
    Logger.log(e);
    throw new Error(`Falha na comunicação com a API do Gemini: ${e.message}`);
  }
}

/**
 * Função para testar a geração da análise ABC pela IA.
 */
function testarGeracaoDeAnaliseABC() {
  const respostaJson = gerarAnaliseABC_comIA();
  
  Logger.log("--- RESPOSTA (SIMULADA) RECEBIDA PELA IA ---");
  
  // Tenta "parsear" a resposta para garantir que é um JSON válido
  try {
    const dados = JSON.parse(respostaJson);
    Logger.log("JSON é válido!");
    Logger.log("Insights: %s", dados.insights.ponto1);
    Logger.log("Dados do Gráfico: %s", JSON.stringify(dados.chartData, null, 2));
  } catch (e) {
    Logger.log("ERRO: A resposta não é um JSON válido.");
    Logger.log(respostaJson);
  }
}

function testarAnaliseEstrategicaCompleta() {
  const analise = gerarAnaliseEstrategica();
  Logger.log("--- ANÁLISE COMPLETA GERADA PELA IA ---");
  Logger.log(analise);
}

/**
 * Função para testar a busca de histórico de um produto.
 */
function testarHistoricoDeProduto() {
  // ✅ TROQUE AQUI pelo ID ou Descrição de um produto real do seu sistema
  const produtoParaTestar = "RV EL VW CONSTELLATION LE"; 
  
  const historico = getPurchaseHistoryForItem(produtoParaTestar);
  
  Logger.log(`--- Histórico de Compras para: ${produtoParaTestar} ---`);
  Logger.log(JSON.stringify(historico, null, 2));
}
