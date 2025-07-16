// ===============================================
    // FUNÇÕES PARA DASHBOARD (NOVAS)
    // ===============================================

    /**
     * Função auxiliar para obter dados de pedidos de forma padronizada.
     * @returns {Array<Object>} Lista de objetos de pedidos com cabeçalhos em camelCase.
     */
    function _getPedidosData(empresa) {
      Logger.log(`[getPedidosData v5] Função iniciada. Filtro: ${JSON.stringify(empresa)}`);
      try {
        const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
        if (!sheet || sheet.getLastRow() < 2) {
          Logger.log('[getPedidosData v5] Planilha "Pedidos" vazia ou não encontrada.');
          return [];
        }

        const values = sheet.getDataRange().getValues();
        const headers = values[0];
        const indexEmpresa = headers.findIndex(h => ["EMPRESA", "IDEMPRESA", "IDDAEMPRESA"].includes(h.toUpperCase()));

        if (indexEmpresa === -1) {
          Logger.log('[getPedidosData v5] ERRO CRÍTICO: Coluna da empresa não encontrada.');
          return [];
        }
        
        const chaveEmpresa = toCamelCase(headers[indexEmpresa]);

        const pedidos = values.slice(1).map(row => {
          const pedido = {};
          headers.forEach((header, index) => {
            pedido[toCamelCase(header)] = row[index];
          });

          
          // Converte a string de itens em um array de objetos
          if (typeof pedido.itens === 'string' && pedido.itens.trim().startsWith('[')) {
            try {
              pedido.itens = JSON.parse(pedido.itens);
            } catch (e) {
              Logger.log(`Erro ao converter itens para o pedido ${pedido.numeroDoPedido || 'N/A'}: ${e.message}`);
              pedido.itens = []; // Garante que seja um array vazio em caso de erro
            }
          } else if (!Array.isArray(pedido.itens)) {
            // Se não for uma string JSON ou um array, força a ser um array vazio
            pedido.itens = [];
          }
          

          pedido.totalGeral = parseFloat(pedido.totalGeral || 0);
          return pedido;
        });

        if (empresa && empresa.id != null) {
          const idEmpresa = String(empresa.id).trim();
          const filtrados = pedidos.filter(p => {
            const pedidoEmpresaValor = String(p[chaveEmpresa] || '').trim();
            return pedidoEmpresaValor == idEmpresa;
          });
          Logger.log(`[getPedidosData v5] Filtro aplicado. Encontrados ${filtrados.length} de ${pedidos.length} pedidos.`);
          return filtrados;
        }
        
        return pedidos;

      } catch (e) {
        Logger.log(`[getPedidosData v5] ERRO FATAL: ${e.message}`);
        return [];
      }
    }

    function sanitizePedidos(pedidos) {
      return pedidos.map(pedido => {
        const clone = { ...pedido };
        // Transforma datas em string padronizada
        if (clone.data instanceof Date) {
          clone.data = formatarDataParaISO(clone.data);
        }
        // Se tiver outros campos Date, serialize-os aqui!
        // Remova funções e campos não serializáveis, se houver
        return clone;
      });
    }

    /**
     * Retorna os fornecedores mais comprados, do maior para o menor volume.
     * @param {Array<Object>} pedidos - A lista de pedidos JÁ FILTRADA.
     * @returns {Array<Object>} Lista de fornecedores com total comprado.
     */
    function _getTopSuppliersByVolume(empresa) {
      const pedidos = _getPedidosData(empresa);
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

      Logger.log(`[Dashboard - _getTopSuppliersByVolume] Fornecedores por volume: ${JSON.stringify(sortedSuppliers)}`);
      return sortedSuppliers;
    }

    /**
     * Retorna os produtos mais pedidos, do maior para o menor volume (quantidade).
     * @param {Array<Object>} pedidos - A lista de pedidos JÁ FILTRADA.
     * @returns {Array<Object>} Lista de produtos com quantidade total pedida.
     */
    function _getTopProductsByVolume(empresa) {
      const pedidos = _getPedidosData(empresa);
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

      Logger.log(`[Dashboard - _getTopProductsByVolume] Produtos por volume: ${JSON.stringify(sortedProducts)}`);
      return sortedProducts;
    }

    /**
     * Retorna um resumo financeiro dos pedidos.
     * @param {Array<Object>} pedidos - A lista de pedidos JÁ FILTRADA.
     * @returns {Object} Resumo financeiro.
     */
    function _getFinancialSummary(empresa) {
      const pedidos = _getPedidosData(empresa);
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

      Logger.log(`[Dashboard - _getFinancialSummary] Resumo financeiro: ${JSON.stringify(summary)}`);
      return summary;
    }

    /**
     * Gera sugestões de redução de compras usando o Gemini API.
     * @param {string} financialSummaryStr - Resumo financeiro em string JSON.
     * @param {string} topProductsStr - Lista de top produtos em string JSON.
     * @param {string} topSuppliersStr - Lista de top fornecedores em string JSON.
     * @returns {string} Sugestões de redução de compras.
     */
    function generatePurchaseSuggestions(financialSummaryStr, topProductsStr, topSuppliersStr) {
      Logger.log('[generatePurchaseSuggestions] Iniciando geração de sugestões...');
      Logger.log('financialSummaryStr: ' + financialSummaryStr);
      Logger.log('topProductsStr: ' + topProductsStr);
      Logger.log('topSuppliersStr: ' + topSuppliersStr);
      // Corrige: se algum parâmetro não vier, substitui por um valor válido
      if (!financialSummaryStr || financialSummaryStr === "undefined") financialSummaryStr = "{}";
      if (!topProductsStr || topProductsStr === "undefined") topProductsStr = "[]";
      if (!topSuppliersStr || topSuppliersStr === "undefined") topSuppliersStr = "[]";

      try {
        const financialSummary = JSON.parse(financialSummaryStr);
        const topProducts = JSON.parse(topProductsStr);
        const topSuppliers = JSON.parse(topSuppliersStr);

        const prompt = `
          **PERSONA:** Você é um consultor de supply chain especializado na otimização de compras de mercadorias.

          **CONTEXTO:** Os dados a seguir são de compras emergenciais ("apaga-incêndio") realizadas em fornecedores locais de alto custo. A empresa deseja criar um plano de ação para minimizar a necessidade dessas compras e reduzir a dependência desses fornecedores.

          **TAREFA:** Com base nos dados, forneça de 3 a 5 recomendações estratégicas e acionáveis. Organize suas sugestões nas seguintes categorias: GESTÃO DE ESTOQUE, ESTRATÉGIA DE FORNECEDORES e ANÁLISE DE PRODUTOS. Para cada sugestão, seja direto e indique um próximo passo prático.

          **DADOS DE ENTRADA:**
          - Resumo Financeiro das Compras Locais: ${JSON.stringify(financialSummary)}
          - Top 5 Produtos de Compra Urgente: ${JSON.stringify(topProducts.slice(0, 5))}
          - Top 5 Fornecedores Locais Mais Utilizados: ${JSON.stringify(topSuppliers.slice(0, 5))}

          **PLANO DE AÇÃO ESTRATÉGICO:**
          `;

        // Obtém a chave da API das propriedades do script
        const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
        
        if (!apiKey) {
            Logger.log('[generatePurchaseSuggestions] Erro: GEMINI_API_KEY não configurada nas Propriedades do Script.');
            return "Erro de configuração: Chave da API Gemini não encontrada. Por favor, configure a 'GEMINI_API_KEY' nas Propriedades do Script.";
        }

        const payload = {
          contents: [{ role: "user", parts: [{ text: prompt }] }],
          generationConfig: {
              temperature: 0.7, // Um pouco mais criativo, mas ainda focado
              maxOutputTokens: 500
          }
        };

        const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

        Logger.log(`[generatePurchaseSuggestions] Enviando requisição para Gemini API: ${apiUrl}`);
        Logger.log(`[generatePurchaseSuggestions] Payload (antes do fetch): ${JSON.stringify(payload)}`); 

        let response;
        try {
          response = UrlFetchApp.fetch(apiUrl, {
            method: 'POST',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true // Para capturar erros na resposta
          });
        } catch (fetchError) {
          Logger.log(`[generatePurchaseSuggestions] Erro ao executar UrlFetchApp.fetch: ${fetchError.message}. Stack: ${fetchError.stack}`);
          return `Erro de rede ou comunicação com a API Gemini: ${fetchError.message}`;
        }

        // Adicionado logs para inspecionar a resposta antes do JSON.parse
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        Logger.log(`[generatePurchaseSuggestions] Código de Resposta da API: ${responseCode}`); 
        Logger.log(`[generatePurchaseSuggestions] Texto de Resposta da API (bruto): ${responseText}`); 

        // ** Correção principal aqui: Verifica se responseText é válido antes de tentar JSON.parse **
        if (!responseText || responseText.trim() === '') {
          let errorMessage = `Erro: Resposta vazia ou nula da API Gemini. Código: ${responseCode}.`;
          Logger.log(`[generatePurchaseSuggestions] Erro detalhado (resposta vazia): ${errorMessage}`);
          return `Erro ao gerar sugestões: ${errorMessage}`;
        }

        let result;
        try {
          result = JSON.parse(responseText); // Agora processa responseText após a verificação de nulo/vazio
        } catch (parseError) {
          Logger.log(`[generatePurchaseSuggestions] SyntaxError: Erro ao parsear JSON da resposta da API Gemini: ${parseError.message}. Conteúdo bruto: ${responseText.substring(0, 500)}... Stack: ${parseError.stack}`);
          return `Erro ao gerar sugestões: Resposta inválida da API Gemini. Tente novamente ou verifique a chave/permissões.`;
        }
        
        Logger.log(`[generatePurchaseSuggestions] Resposta da API Gemini (parseada): ${JSON.stringify(result)}`);

        if (result.candidates && result.candidates.length > 0 &&
            result.candidates[0].content && result.candidates[0].content.parts &&
            result.candidates[0].content.parts.length > 0) {
          const suggestions = result.candidates[0].content.parts[0].text;
          Logger.log(`[generatePurchaseSuggestions] Sugestões geradas: ${suggestions}`);
          return suggestions;
        } else {
          Logger.log(`[generatePurchaseSuggestions] Resposta da API Gemini não contém sugestões válidas (candidates ausentes): ${JSON.stringify(result)}`);
          if (result.error && result.error.message) {
            return `Erro da API Gemini: ${result.error.message}`;
          }
          return "Não foi possível gerar sugestões no momento. A API retornou uma estrutura inesperada ou incompleta.";
        }

      } catch (e) {
        Logger.log(`[generatePurchaseSuggestions] Erro fatal (fora do bloco fetch/parse): ${e.message}. Stack: ${e.stack}`);
        return `Ocorreu um erro inesperado ao gerar sugestões: ${e.message}`;
      }
    }

    /**
     * Função principal para o Dashboard, que coleta todos os dados e gera sugestões.
     * @returns {Object} Objeto contendo todos os dados do dashboard e sugestões.
     */
    function getDashboardData(empresa) {
      // Este novo log confirma que a versão correta do código está executando.
      Logger.log('--- INICIANDO EXECUÇÃO COM O CÓDIGO CORRIGIDO (v3) ---'); 
      Logger.log(`Filtro recebido para a empresa: ${JSON.stringify(empresa)}`);
      
      try {
        // As chamadas abaixo agora passam o filtro 'empresa'
        const financialSummary = _getFinancialSummary(empresa);
        const topProducts = _getTopProductsByVolume(empresa);
        const topSuppliers = _getTopSuppliersByVolume(empresa);
        
        const suggestions = generatePurchaseSuggestions(
          JSON.stringify(financialSummary || {}),
          JSON.stringify(topProducts || []),
          JSON.stringify(topSuppliers || [])
        );

        Logger.log('--- SUCESSO! Retornando dados JÁ FILTRADOS. ---');
        
        return {
          status: 'success',
          topSuppliers: topSuppliers,
          topProducts: topProducts,
          financialSummary: financialSummary,
          suggestions: suggestions
        };
      } catch (e) {
        Logger.log(`ERRO FATAL: ${e.message}`);
        return { status: 'error', message: `Erro ao carregar dados do Dashboard: ${e.message}` };
      }
    }

    function serializePedidos(pedidos) {
      return pedidos.map(p => {
        const novoPedido = { ...p };
        // Transforma datas em string padronizada
        if (novoPedido.data instanceof Date) {
          novoPedido.data = formatarDataParaISO(novoPedido.data);
        }
        // Se tiver outros campos do tipo Date, faça igual aqui
        return novoPedido;
      });
    }