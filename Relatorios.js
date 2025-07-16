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
     * Busca pedidos na planilha 'Pedidos' com base em um termo de busca.
     * O termo de busca pode ser o número do pedido ou o nome do fornecedor.
     * @param {string} termoBusca - O termo a ser buscado.
     * @param {string} empresaCodigo - O código da empresa (ex: "E001").
     * @returns {Object} Um objeto com status e uma lista de pedidos que correspondem ao termo de busca.
     */
    /**
     * VERSÃO DE DIAGNÓSTICO
     * Busca pedidos com base em um termo e no código da empresa, com logs detalhados.
     */
    function buscarPedidos(termoBusca, empresaCodigo) {
      console.log('🔍 === INÍCIO buscarPedidos ===');
      console.log('🔍 Termo de busca:', termoBusca);
      console.log('🔍 Código da empresa:', empresaCodigo);
      
      try {
        const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
        if (!sheet) {
          console.error('❌ Planilha "Pedidos" não encontrada');
          return { status: "error", data: [], message: "Aba 'Pedidos' não encontrada." };
        }

        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        console.log('📋 Headers da planilha:', headers);
        
        const resultados = [];
        
        const idEmpresaFiltro = String(empresaCodigo).trim();
        const termoNormalizado = (termoBusca || "").toString().toLowerCase().trim();
        console.log('🔍 ID empresa filtro:', idEmpresaFiltro);
        console.log('🔍 Termo normalizado:', termoNormalizado);

        const colEmpresa = headers.findIndex(h => ["EMPRESA", "IDEMPRESA", "IDDAEMPRESA", "ID DA EMPRESA", "Empresa"].includes(String(h).toUpperCase()));
        const colNumeroPedido = headers.findIndex(h => String(h).toUpperCase() === 'NÚMERO DO PEDIDO');
        const colFornecedor = headers.findIndex(h => String(h).toUpperCase() === 'FORNECEDOR');
        
        console.log('📍 Índices das colunas:');
        console.log('  - Empresa:', colEmpresa, headers[colEmpresa]);
        console.log('  - Número do Pedido:', colNumeroPedido, headers[colNumeroPedido]);
        console.log('  - Fornecedor:', colFornecedor, headers[colFornecedor]);

        if (colEmpresa === -1) {
          console.error('❌ Coluna da empresa não encontrada');
          return { status: "error", data: [], message: "Coluna da empresa não encontrada." };
        }

        console.log(`🔍 Total de linhas para processar: ${data.length - 1}`);
        
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const idEmpresaNaLinha = String(row[colEmpresa]).trim();

          console.log(`📋 Linha ${i}: ID empresa na linha: "${idEmpresaNaLinha}", Filtro: "${idEmpresaFiltro}"`);

          if (parseInt(idEmpresaNaLinha, 10) !== parseInt(idEmpresaFiltro, 10)) {
            console.log(`⏭️ Linha ${i}: Empresa não corresponde - pulando`);
            continue;
          }

          console.log(`✅ Linha ${i}: Empresa corresponde - verificando termo de busca`);

          const numeroPedidoNormalizado = String(row[colNumeroPedido] || '').toLowerCase().trim();
          const fornecedorNormalizado = String(row[colFornecedor] || '').toLowerCase().trim();

          console.log(`🔍 Linha ${i}: Número: "${numeroPedidoNormalizado}", Fornecedor: "${fornecedorNormalizado}"`);

          const shouldAddRow = (termoNormalizado === "") || 
                              numeroPedidoNormalizado.includes(termoNormalizado) || 
                              fornecedorNormalizado.includes(termoNormalizado);

          console.log(`🔍 Linha ${i}: Deve adicionar? ${shouldAddRow}`);
                              
          if (shouldAddRow) {
            const pedidoData = {};
            headers.forEach((header, index) => {
              let value = row[index];

              // --- AQUI ESTÁ A CORREÇÃO FINAL ---
              // Se o valor for um objeto de Data, converte para texto no formato AAAA-MM-DD
              if (value instanceof Date) {
                value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
              }
              // --- FIM DA CORREÇÃO ---

              pedidoData[toCamelCase(header)] = value;
            });

            // O resto da sua lógica de parse e etc.
            pedidoData.totalGeral = parseFloat(pedidoData.totalGeral || 0);
            if (typeof pedidoData.itens === 'string') {
              try {
                pedidoData.itens = JSON.parse(pedidoData.itens);
              } catch (e) {
                pedidoData.itens = [];
              }
            } else if (!Array.isArray(pedidoData.itens)) {
              pedidoData.itens = [];
            }

            console.log(`➕ Linha ${i}: Adicionando pedido aos resultados`);
            resultados.push(pedidoData);
          }
        }
        
        console.log(`📊 Total de resultados encontrados: ${resultados.length}`);
        console.log('🔍 === FIM buscarPedidos ===');
        
        // Agora o objeto 'resultados' não contém mais objetos de Data e pode ser retornado
        return { status: "success", data: resultados };

      } catch (e) {
        Logger.log(`Erro fatal em buscarPedidos: ${e.message}`);
        return { status: "error", data: [], message: `Erro no servidor: ${e.message}` };
      }
    }

    // ===============================================
    // NOVAS FUNÇÕES PARA RELATÓRIOS
    // ===============================================

    /**
     * Retorna uma lista de todos os produtos únicos de todos os pedidos.
     * @returns {Array<string>} Uma lista de nomes de produtos únicos.
     */

    function _getPedidosDatas(reportParams) {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
      if (!sheet || sheet.getLastRow() < 2) {
        Logger.log('[_getPedidosData] Planilha "Pedidos" vazia ou não encontrada.');
        return [];
      }
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      const headers = values[0];
      const pedidos = [];

      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const pedido = {};
        headers.forEach((header, index) => {
          const camelCaseHeader = toCamelCase(header);
          let value = row[index];
          // Converte string de data para Date
          if (camelCaseHeader === 'data' && typeof value === 'string') {
            try {
              value = new Date(value + 'T12:00:00');
            } catch (e) {
              Logger.log(`Erro ao parsear data ${value} no pedido ${row[0]}: ${e.message}`);
              value = null;
            }
          } else if (value instanceof Date) {
            // ok
          } else if (camelCaseHeader === 'data' && !(value instanceof Date)) {
            value = null;
          }
          pedido[camelCaseHeader] = value;
        });

        // Parse itens JSON caso necessário
        if (typeof pedido.itens === 'string') {
          try {
            pedido.itens = JSON.parse(pedido.itens);
          } catch (e) {
            Logger.log(`[Pedidos] Erro ao parsear itens JSON para pedido ${pedido.numeroDoPedido}: ${e.message}`);
            pedido.itens = [];
          }
        } else if (!Array.isArray(pedido.itens)) {
          pedido.itens = [];
        }

        pedido.totalGeral = parseFloat(pedido.totalGeral || 0);

        pedidos.push(pedido);
      }

      // --- FILTRO MULTI-EMPRESA ---
      if (reportParams && reportParams.empresaCnpj) {
        const idEmpresaFiltro = String(reportParams.companyId).trim();
        Logger.log(`[_getPedidosDatas] Filtrando pedidos pelo ID da empresa: ${idEmpresaFiltro}`);

        const pedidosFiltrados = pedidos.filter(p => {
          // Assumindo que a coluna na sua planilha 'Pedidos' se chama 'Empresa' ou 'ID Empresa'.
          // O toCamelCase vai transformar isso em 'empresa' ou 'idEmpresa'.
          const idDoPedidoNaLinha = String(p.empresa || p.idEmpresa || '').trim();
          
          // Usamos '==' para o caso da planilha ter o número 1 e o filtro ser "001"
          return idDoPedidoNaLinha == idEmpresaFiltro;
        });

        Logger.log(`[_getPedidosDatas] Encontrados ${pedidosFiltrados.length} pedidos para a empresa.`);
        return pedidosFiltrados;
      }

      Logger.log('[_getPedidosDatas] Nenhum filtro de empresa aplicado. Retornando todos os pedidos.');
      return pedidos;
    }

    function getUniqueProducts(reportParams) {
      Logger.log('[getUniqueProducts - SERVER] Iniciando busca por produtos únicos.');
      const pedidos = _getPedidosDatas(reportParams); // Reutiliza a função que busca e formata os dados dos pedidos
      const uniqueProducts = new Set();

      pedidos.forEach(pedido => {
        (pedido.itens || []).forEach(item => {
          if (item.descricao && item.descricao.trim() !== '') {
            uniqueProducts.add(item.descricao.trim());
          }
        });
      });

      const sortedProducts = Array.from(uniqueProducts).sort((a, b) => a.localeCompare(b));
      Logger.log(`[getUniqueProducts - SERVER] Produtos únicos encontrados: ${JSON.stringify(sortedProducts)}`);
      return sortedProducts;
    }

    /**
     * Função para filtrar pedidos com base em parâmetros.
     * @param {Array<Object>} allPedidos - Todos os pedidos brutos.
     * @param {Object} reportParams - Parâmetros de filtro (startDate, endDate, supplier).
     * @returns {Array<Object>} Pedidos filtrados.
     */
    function _filterPedidos(allPedidos, reportParams) {
        let filtered = allPedidos;

        // Adiciona uma verificação explícita para reportParams aqui
        if (!reportParams || typeof reportParams !== 'object') {
            Logger.log('[_filterPedidos] Aviso: reportParams é undefined ou não é um objeto. Nenhum filtro de data/fornecedor será aplicado.');
            return filtered; // Retorna todos os pedidos se não houver parâmetros válidos
        }

        // Filtro por data
        if (reportParams.startDate && reportParams.endDate) {
            const start = new Date(reportParams.startDate + 'T00:00:00');
            const end = new Date(reportParams.endDate + 'T23:59:59');
            filtered = filtered.filter(pedido => {
                // Se a data do pedido não for um Date object válido, exclua-o ou trate
                if (!(pedido.data instanceof Date)) {
                    Logger.log(`Data inválida no pedido ${pedido.numeroDoPedido}: ${pedido.data}. Excluindo do filtro.`);
                    return false;
                }
                return pedido.data >= start && pedido.data <= end;
            });
            Logger.log(`[_filterPedidos] Filtrado por data: ${filtered.length} pedidos restantes.`);
        }

        // Filtro por fornecedor
        if (reportParams.supplier) {
            const selectedSupplier = reportParams.supplier.toLowerCase();
            filtered = filtered.filter(pedido => (pedido.fornecedor || '').toLowerCase() === selectedSupplier);
            Logger.log(`[_filterPedidos] Filtrado por fornecedor '${reportParams.supplier}': ${filtered.length} pedidos restantes.`);
        }

        return filtered;
    }

    /**
     * Agrupa e sumariza dados de pedidos para geração de relatório.
     * @param {Array<Object>} pedidos - Pedidos já filtrados.
     * @param {Object} reportParams - Parâmetros do relatório (reportType).
     * @returns {Object} Dados agrupados e/ou sumarizados.
     */
    function _groupAndSummarizePedidos(pedidos, reportParams) {
        const data = {};

        if (reportParams.reportType === 'detailed') {
            // Relatório Detalhado: Agrupa por data, depois por fornecedor
            pedidos.sort((a, b) => a.data.getTime() - b.data.getTime()); // Ordena por data

            pedidos.forEach(pedido => {
                const dateStr = Utilities.formatDate(pedido.data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                if (!data[dateStr]) {
                    data[dateStr] = {
                        date: pedido.data,
                        fornecedores: {},
                        totalDate: 0
                    };
                }
                const fornecedor = pedido.fornecedor || 'Desconhecido';
                if (!data[dateStr].fornecedores[fornecedor]) {
                    data[dateStr].fornecedores[fornecedor] = {
                        pedidos: [],
                        totalFornecedor: 0
                    };
                }
                data[dateStr].fornecedores[fornecedor].pedidos.push(pedido);
                data[dateStr].fornecedores[fornecedor].totalFornecedor += pedido.totalGeral || 0;
                data[dateStr].totalDate += pedido.totalGeral || 0;
            });
        } else if (reportParams.reportType === 'financial') {
            // Relatório Financeiro: Sumariza totais
            data.totalGeralPedidos = 0;
            data.numeroTotalPedidos = 0;
            data.totalPorFornecedor = {};
            
            pedidos.forEach(pedido => {
                data.totalGeralPedidos += pedido.totalGeral || 0;
                data.numeroTotalPedidos++;
                const fornecedor = pedido.fornecedor || 'Desconhecido';
                data.totalPorFornecedor[fornecedor] = (data.totalPorFornecedor[fornecedor] || 0) + (pedido.totalGeral || 0);
            });
            // Converter para array para facilitar o uso no HTML, se necessário
            data.listaTotalPorFornecedor = Object.keys(data.totalPorFornecedor).map(f => ({
                fornecedor: f,
                total: data.totalPorFornecedor[f]
            })).sort((a, b) => b.total - a.total); // Ordena do maior para o menor
        }
        Logger.log(`[_groupAndSummarizePedidos] Dados agrupados: ${JSON.stringify(data, null, 2)}`);
        return data;
    }


    /**
     * Gera o conteúdo HTML para o relatório.
     * @param {Object} reportData - Dados do relatório (agrupados/sumarizados por _groupAndSummarizePedidos).
     * @param {Object} reportParams - Parâmetros do relatório para cabeçalho (reportType, startDate, endDate, supplier).
     * @returns {string} Conteúdo HTML do relatório.
     */
    function _generatePdfHtmlContent(reportData, reportParams) {
        const companyName = reportParams.companyName || "EMPRESA NÃO INFORMADA";
        const companyAddress = reportParams.companyAddress || "";
        const companyCnpj = reportParams.empresaCnpj || "";
        const reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

        let reportTitle = "Relatório de Compras";
        if (reportParams.reportType === 'detailed') {
            reportTitle += " Detalhado";
        } else if (reportParams.reportType === 'financial') {
            reportTitle += " Financeiro";
        }

        let filtersApplied = [];
        if (reportParams.startDate && reportParams.endDate) {
            filtersApplied.push(`Período: ${Utilities.formatDate(new Date(reportParams.startDate + 'T00:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy')} a ${Utilities.formatDate(new Date(reportParams.endDate + 'T00:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy')}`);
        }
        if (reportParams.supplier) {
            filtersApplied.push(`Fornecedor: ${reportParams.supplier}`);
        }
        const filtersHtml = filtersApplied.length > 0 ? `<p class="filters">Filtros Aplicados: ${filtersApplied.join(' | ')}</p>` : '';

        let bodyContent = '';

        if (reportParams.reportType === 'detailed') {
            if (Object.keys(reportData).length === 0) {
                bodyContent += '<p class="no-data">Nenhum pedido encontrado para os filtros selecionados.</p>';
            } else {
                for (const dateStr in reportData) {
                    const dateGroup = reportData[dateStr];
                    bodyContent += `<div class="date-group">`;
                    bodyContent += `<h3>Data: ${Utilities.formatDate(dateGroup.date, Session.getScriptTimeZone(), 'dd/MM/yyyy')} (Total: R$ ${dateGroup.totalDate.toFixed(2).replace('.', ',')})</h3>`;
                    
                    for (const supplierName in dateGroup.fornecedores) {
                        const supplierGroup = dateGroup.fornecedores[supplierName];
                        bodyContent += `<div class="supplier-group">`;
                        bodyContent += `<h4>Fornecedor: ${supplierName} (Total: R$ ${supplierGroup.totalFornecedor.toFixed(2).replace('.', ',')})</h4>`;
                        bodyContent += `<table>`;
                        bodyContent += `<thead><tr>
                            <th>Número Pedido</th>
                            <th>CNPJ Fornecedor</th>
                            <th>Razao Social</th>
                            <th>Item</th>
                            <th>Unidade</th>
                            <th>Qtd.</th>
                            <th>Preço Unit.</th>
                            <th>Subtotal Item</th>
                        </tr></thead>`;
                        bodyContent += `<tbody>`;
                        supplierGroup.pedidos.forEach(pedido => {
                            pedido.itens.forEach(item => {
                                bodyContent += `<tr>
                                    <td>${pedido.numeroDoPedido}</td>
                                    <td>${pedido.cnpjFornecedor || ''}</td>
                                    <td>${supplierName}</td>
                                    <td>${item.descricao || ''}</td>
                                    <td>${item.unidade || ''}</td>
                                    <td>${item.quantidade}</td>
                                    <td>R$ ${parseFloat(item.precoUnitario || 0).toFixed(2).replace('.', ',')}</td>
                                    <td>R$ ${parseFloat(item.totalItem || 0).toFixed(2).replace('.', ',')}</td>
                                </tr>`;
                            });
                            // Add a row for total of this specific order
                            bodyContent += `<tr class="pedido-total-row">
                                <td colspan="7" style="text-align:right; font-weight:bold;">Total Pedido ${pedido.numeroDoPedido}:</td>
                                <td style="font-weight:bold;">R$ ${parseFloat(pedido.totalGeral || 0).toFixed(2).replace('.', ',')}</td>
                            </tr>`;
                        });
                        bodyContent += `</tbody>`;
                        bodyContent += `</table>`;
                        bodyContent += `</div>`; // Close supplier-group
                    }
                    bodyContent += `</div>`; // Close date-group
                }
            }
        } else if (reportParams.reportType === 'financial') {
            if (reportData.numeroTotalPedidos === 0) {
                bodyContent += '<p class="no-data">Nenhum dado financeiro encontrado para os filtros selecionados.</p>';
            } else {
                bodyContent += `<p class="financial-summary-total">Total Geral de Pedidos: ${reportData.numeroTotalPedidos}</p>`;
                bodyContent += `<p class="financial-summary-total">Valor Total das Compras: <span class="total-value-display">R$ ${reportData.totalGeralPedidos.toFixed(2).replace('.', ',')}</span></p>`;
                
                bodyContent += `<h4>Totais por Fornecedor:</h4>`;
                bodyContent += `<table>`;
                bodyContent += `<thead><tr><th>Fornecedor</th><th>Valor Total Comprado</th></tr></thead>`;
                bodyContent += `<tbody>`;
                reportData.listaTotalPorFornecedor.forEach(item => {
                    bodyContent += `<tr>
                        <td>${item.fornecedor}</td>
                        <td>R$ ${item.total.toFixed(2).replace('.', ',')}</td>
                    </tr>`;
                });
                bodyContent += `</tbody>`;
                bodyContent += `</table>`;
            }
        }

        // Constrói o HTML completo do relatório
        const html = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>${reportTitle}</title>
            <style>
                body { font-family: 'Arial', sans-serif; margin: 20px; font-size: 10pt; }
                .header, .footer { text-align: center; margin-bottom: 20px; }
                .header h1 { margin: 0; color: #0056b3; font-size: 16pt; }
                .header p { margin: 2px 0; font-size: 8pt; }
                .report-title { text-align: center; font-size: 14pt; margin-bottom: 15px; color: #007BFF; }
                .filters { font-size: 9pt; text-align: center; margin-bottom: 15px; color: #555; }
                table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
                th, td { border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 9pt; }
                th { background-color: #f2f2f2; font-weight: bold; text-align: center; }
                .date-group, .supplier-group { margin-bottom: 25px; border: 1px solid #eee; padding: 15px; border-radius: 5px; }
                .date-group h3 { margin-top: 0; color: #0056b3; font-size: 12pt; border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-bottom: 15px; }
                .supplier-group h4 { margin-top: 0; color: #007BFF; font-size: 11pt; margin-bottom: 10px; }
                .no-data { text-align: center; font-style: italic; color: #888; margin-top: 30px; }
                .financial-summary-total { font-size: 12pt; font-weight: bold; text-align: center; margin-bottom: 10px; }
                .total-value-display { color: #28a745; }
                .pedido-total-row { background-color: #f0f8ff; font-style: italic; }
                .pedido-total-row td { border-top: 2px solid #007bff; }
                .page-break { page-break-before: always; } /* Para quebrar página antes de cada grupo de data/fornecedor se necessário */
            </style>
        </head>
        <body>
            <div class="header">
                <h1>${companyName}</h1>
                <p>${companyAddress}</p>
                <p>CNPJ: ${companyCnpj}</p>
                <p>Relatório Gerado em: ${reportDate}</p>
            </div>
            <h2 class="report-title">${reportTitle}</h2>
            ${filtersHtml}
            ${bodyContent}
            <div class="footer">
                <p>&copy; ${new Date().getFullYear()} ${companyName}. Todos os direitos reservados.</p>
            </div>
        </body>
        </html>
        `;
        return html;
    }

    /**
     * Gera um relatório em PDF com base nos parâmetros fornecidos.
     * @param {Object} reportParams - Objeto com: startDate, endDate, supplier, reportType.
     * @returns {string} URL temporária do arquivo PDF gerado no Google Drive.
     */
    function generatePdfReport(reportParams) {
      Logger.log(`[generatePdfReport] Iniciando. Parâmetros recebidos (antes de qualquer processamento): ${JSON.stringify(reportParams)}`);

      // Adiciona uma verificação explícita para reportParams
      if (reportParams === undefined || reportParams === null) {
          Logger.log('[generatePdfReport] ERRO CRÍTICO: reportParams é undefined ou null. Abortando.');
          throw new Error('Parâmetros de relatório ausentes ou inválidos. Tente novamente.');
      }

      try {
        const allPedidos = _getPedidosDatas(reportParams);
        const filteredPedidos = _filterPedidos(allPedidos, reportParams);
        const reportData = _groupAndSummarizePedidos(filteredPedidos, reportParams);
        const htmlContent = _generatePdfHtmlContent(reportData, reportParams);
        Logger.log(`[generatePdfReport] Conteúdo HTML gerado (primeiros 500 caracteres): ${htmlContent.substring(0, Math.min(htmlContent.length, 500))}`);


        const htmlBlob = Utilities.newBlob(htmlContent, MimeType.HTML, 'RelatorioTemporario.html');
        Logger.log(`[generatePdfReport] htmlBlob criado. Tipo: ${htmlBlob.getContentType()}, Tamanho: ${htmlBlob.getBytes().length} bytes.`);

        const folderName = "RelatoriosComprasTemporarios";
        let folder;
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) { 
          folder = folders.next();
          Logger.log(`[generatePdfReport] Pasta existente encontrada: '${folder.getName()}' (ID: ${folder.getId()})`);
        } else {
          folder = DriveApp.createFolder(folderName);
          Logger.log(`[generatePdfReport] Pasta '${folderName}' criada no Google Drive com ID: ${folder.getId()}.`);
        }

        // Verifica se a pasta foi obtida/criada com sucesso antes de prosseguir
        if (!folder) {
            Logger.log('[generatePdfReport] ERRO: A pasta de destino é nula ou indefinida.');
            throw new Error('Não foi possível obter ou criar a pasta de destino no Google Drive.');
        }

        const htmlFile = folder.createFile(htmlBlob);
        Logger.log(`[generatePdfReport] Arquivo HTML temporário criado: ${htmlFile.getUrl()} (ID: ${htmlFile.getId()}).`);

        // Verifica se o arquivo HTML foi criado com sucesso antes de tentar converter para PDF
        if (!htmlFile) {
            Logger.log('[generatePdfReport] ERRO: O arquivo HTML temporário é nulo ou indefinido após a criação.');
            throw new Error('Falha ao criar o arquivo HTML temporário.');
        }

        // --- PONTO CRÍTICO: Conversão para PDF ---
        Logger.log('[generatePdfReport] Tentando converter o htmlFile para PDF...');
        const pdfBlob = htmlFile.getAs(MimeType.PDF);
        Logger.log(`[generatePdfReport] Conversão para PDF concluída. pdfBlob criado: ${pdfBlob ? 'Sim' : 'Não'}. Tamanho: ${pdfBlob ? pdfBlob.getBytes().length : 'N/A'} bytes.`);

        if (!pdfBlob) {
            Logger.log('[generatePdfReport] ERRO LÓGICO: pdfBlob é nulo ou indefinido após getAs(MimeType.PDF).');
            throw new Error('Falha ao converter o arquivo HTML para PDF: o blob resultante é nulo.');
        }
        
        let titleForFileName = "RelatorioCompras";
        if (reportParams.reportType === 'detailed') {
            titleForFileName += "_Detalhado";
        } else if (reportParams.reportType === 'financial') {
            titleForFileName += "_Financeiro";
        }
        const pdfFileName = `${titleForFileName}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.pdf`;
        
        const pdfFile = folder.createFile(pdfBlob).setName(pdfFileName);
        Logger.log(`[generatePdfReport] Arquivo PDF final criado: ${pdfFile.getUrl()} (ID: ${pdfFile.getId()}).`);

        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileUrl = pdfFile.getDownloadUrl(); 
        Logger.log(`[generatePdfReport] URL de download do PDF: ${fileUrl}`);
        
        // Opcional: Remover o arquivo HTML temporário imediatamente após a conversão
        htmlFile.setTrashed(true); 
        Logger.log(`[generatePdfReport] Arquivo HTML temporário '${htmlFile.getName()}' movido para a lixeira.`);

        return fileUrl;

      } catch (e) {
        Logger.log(`[generatePdfReport] ERRO DURANTE A GERAÇÃO DO PDF: ${e.message}. Stack: ${e.stack}`);
        // Adicione esta linha para relançar um erro mais limpo para o cliente
        throw new Error(`Erro ao gerar relatório PDF: ${e.message}`);
      }
    }

    /**
     * Gera um relatório em XLS (planilha) com base nos parâmetros fornecidos.
     * @param {Object} reportParams - Objeto com parâmetros do relatório (startDate, endDate, supplier, reportType).
     * @returns {string} URL temporária do arquivo XLS gerado no Google Drive.
     */
    function generateXlsReport(reportParams) {
      Logger.log(`[generateXlsReport] Iniciando. Parâmetros recebidos: ${JSON.stringify(reportParams)}`);

      if (reportParams === undefined || reportParams === null) {
          Logger.log('[generateXlsReport] ERRO CRÍTICO: reportParams é undefined ou null. Abortando.');
          throw new Error('Parâmetros de relatório ausentes ou inválidos para XLS. Tente novamente.');
      }

      try {
        const allPedidos = _getPedidosDatas(reportParams);
        const filteredPedidos = _filterPedidos(allPedidos, reportParams);
        const reportData = _groupAndSummarizePedidos(filteredPedidos, reportParams);

        const ssTitle = `RelatorioCompras_${reportParams.reportType}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
        const spreadsheet = SpreadsheetApp.create(ssTitle);
        const spreadsheetId = spreadsheet.getId();
        const sheet = spreadsheet.getActiveSheet();
        sheet.setName("Relatorio");

        const reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
        // Adiciona cabeçalho do relatório na planilha
        const companyName = reportParams.companyName || "EMPRESA NÃO INFORMADA";
        const companyAddress = reportParams.companyAddress || "";
        const companyCnpj = reportParams.empresaCnpj || "";
        
        sheet.getRange('A1').setValue(companyName);
        sheet.getRange('A2').setValue(companyAddress);
        sheet.getRange('A3').setValue(`CNPJ: ${companyCnpj}`);
        sheet.getRange('A4').setValue(`Relatório Gerado em: ${reportDate}`);
        sheet.getRange('A6').setValue(`Relatório de Compras ${reportParams.reportType === 'detailed' ? 'Detalhado' : 'Financeiro'}`);

        let filtersApplied = [];
        if (reportParams.startDate && reportParams.endDate) {
            filtersApplied.push(`Período: ${Utilities.formatDate(new Date(reportParams.startDate + 'T00:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy')} a ${Utilities.formatDate(new Date(reportParams.endDate + 'T00:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy')}`);
        }
        if (reportParams.supplier) {
            filtersApplied.push(`Fornecedor: ${reportParams.supplier}`);
        }
        if (filtersApplied.length > 0) {
            sheet.getRange('A7').setValue(`Filtros Aplicados: ${filtersApplied.join(' | ')}`);
        }

        let startRow = 9; // Começa a inserir dados a partir da linha 9

        if (reportParams.reportType === 'detailed') {
            const headers = [
                'Número do Pedido', 'Data do Pedido', 'Fornecedor', 'CNPJ Fornecedor', 'Razão Social',
                'Endereço Fornecedor', 'Condição Pagamento', 'Forma Pagamento', 
                'Placa Veículo', 'Nome Veículo', 'Observações',
                'Descrição do Item', 'Unidade', 'Quantidade', 'Preço Unitário', 'Subtotal Item', 'Total Geral do Pedido'
            ];
            sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
            startRow++;

            let rowData = [];
            // Ordenar por data e depois por fornecedor
            const sortedDates = Object.keys(reportData).sort();
            sortedDates.forEach(dateStr => {
                const dateGroup = reportData[dateStr];
                const sortedSupplierNames = Object.keys(dateGroup.fornecedores).sort();
                sortedSupplierNames.forEach(supplierName => {
                    const supplierGroup = dateGroup.fornecedores[supplierName];
                    supplierGroup.pedidos.forEach(pedido => {
                        const pedidoBaseRow = [
                            pedido.numeroDoPedido,
                            pedido.data ? Utilities.formatDate(pedido.data, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '', // Formata a data
                            pedido.fornecedor,
                            pedido.cnpjFornecedor,
                            supplierName,
                            pedido.enderecoFornecedor,
                            pedido.condicaoPagamentoFornecedor,
                            pedido.formaPagamentoFornecedor,
                            pedido.placaVeiculo,
                            pedido.nomeVeiculo,
                            pedido.observacoes,
                            '', '', '', '', '', '' // Colunas vazias para itens
                        ];

                        if (pedido.itens && pedido.itens.length > 0) {
                            pedido.itens.forEach(item => {
                                const itemRow = [
                                    '', '', '', '', '', '', '', '', '', '', // Colunas vazias para dados do pedido
                                    item.descricao || '',
                                    item.unidade || '',
                                    item.quantidade || 0,
                                    parseFloat(item.precoUnitario || 0),
                                    parseFloat(item.totalItem || 0),
                                    '' // Coluna vazia para total geral do pedido
                                ];
                                rowData.push(pedidoBaseRow.slice(0, 11).concat(itemRow.slice(11))); // Combina dados do pedido com item
                            });
                            // Adiciona a linha do total geral do pedido na última linha de itens
                            const lastItemRow = rowData[rowData.length - 1];
                            if (lastItemRow) {
                                lastItemRow[headers.length - 1] = parseFloat(pedido.totalGeral || 0); // Define o total geral na última coluna do pedido
                            }
                        } else {
                            // Se não houver itens, ainda adiciona a linha do pedido com o total geral
                            const emptyItemRow = ['', '', '', '', '', ''];
                            rowData.push(pedidoBaseRow.slice(0, 11).concat(emptyItemRow));
                            rowData[rowData.length - 1][headers.length - 1] = parseFloat(pedido.totalGeral || 0);
                        }
                    });
                });
            });

            if (rowData.length > 0) {
                sheet.getRange(startRow, 1, rowData.length, headers.length).setValues(rowData);
                // Formatar colunas numéricas
                sheet.getRange(startRow, 13, rowData.length, 1).setNumberFormat('0'); // Quantidade
                sheet.getRange(startRow, 14, rowData.length, 1).setNumberFormat('R$#,##0.00'); // Preço Unitário
                sheet.getRange(startRow, 15, rowData.length, 1).setNumberFormat('R$#,##0.00'); // Subtotal Item
                sheet.getRange(startRow, 16, rowData.length, 1).setNumberFormat('R$#,##0.00'); // Total Geral do Pedido
                Logger.log(`[generateXlsReport] Dados detalhados inseridos na planilha.`);
            } else {
                sheet.getRange(startRow, 1).setValue('Nenhum pedido encontrado para os filtros selecionados.').setFontStyle('italic');
            }
                  // Ajusta as larguras (pode precisar ajustar os índices se quiser)
                sheet.autoResizeColumns(1, headers.length);
                sheet.setColumnWidth(1, 120); // Número do Pedido
                sheet.setColumnWidth(2, 90);  // Data
                sheet.setColumnWidth(3, 180); // Fornecedor
                sheet.setColumnWidth(11, 160); // Descrição do Item

                // Congela cabeçalho da tabela e primeiras colunas
                sheet.setFrozenRows(startRow - 1);
                sheet.setFrozenColumns(2);
        
        // Ajusta larguras
            sheet.autoResizeColumns(1, headers.length);
            sheet.setColumnWidth(1, 120); // Número do Pedido
            sheet.setColumnWidth(2, 90);  // Data
            sheet.setColumnWidth(3, 180); // Fornecedor
            sheet.setColumnWidth(11, 160); // Descrição do Item

            // Congela cabeçalho da tabela e primeiras colunas
            sheet.setFrozenRows(startRow - 1);
            sheet.setFrozenColumns(2);

        } else if (reportParams.reportType === 'financial') {
            sheet.getRange(startRow, 1).setValue(`Total Geral de Pedidos: ${reportData.numeroTotalPedidos}`).setFontWeight('bold');
            sheet.getRange(startRow + 1, 1).setValue(`Valor Total das Compras: R$ ${reportData.totalGeralPedidos.toFixed(2).replace('.', ',')}`).setFontWeight('bold');
            sheet.getRange(startRow + 1, 2).setNumberFormat('R$#,##0.00');

            startRow += 3;

            const supplierHeaders = ['Fornecedor', 'Valor Total Comprado'];
            sheet.getRange(startRow, 1, 1, supplierHeaders.length).setValues([supplierHeaders]);
            
            // Formatação do cabeçalho de fornecedores
            const supHeaderRange = sheet.getRange(startRow, 1, 1, supplierHeaders.length);
            supHeaderRange.setFontWeight('bold')
                          .setFontSize(11)
                          .setBackground('#0056b3')
                          .setFontColor('white')
                          .setHorizontalAlignment('center')
                          .setVerticalAlignment('middle')
                          .setBorder(true, true, true, true, true, true);

            startRow++;

            const supplierData = reportData.listaTotalPorFornecedor.map(item => [
                item.fornecedor,
                item.total
            ]);

            if (supplierData.length > 0) {
                sheet.getRange(startRow, 1, supplierData.length, supplierHeaders.length).setValues(supplierData);
                sheet.getRange(startRow, 2, supplierData.length, 1).setNumberFormat('R$#,##0.00');
                // Formatação dos dados
                const supDataRange = sheet.getRange(startRow, 1, supplierData.length, supplierHeaders.length);
                supDataRange.setFontSize(10).setVerticalAlignment('middle');
                supDataRange.setBorder(true, true, true, true, false, false);
            } else {
                sheet.getRange(startRow, 1).setValue('Nenhum dado financeiro por fornecedor encontrado.').setFontStyle('italic');
            }
            sheet.autoResizeColumns(1, supplierHeaders.length);
            sheet.setFrozenRows(startRow - 1);
        }

        // Cria uma pasta temporária (ou usa uma existente)
        const folderName = "RelatoriosComprasTemporarios";
        let folder;
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) { 
          folder = folders.next();
          Logger.log(`[generateXlsReport] Pasta existente encontrada: '${folder.getName()}' (ID: ${folder.getId()})`);
        } else {
          folder = DriveApp.createFolder(folderName);
          Logger.log(`[generateXlsReport] Pasta '${folderName}' criada no Google Drive com ID: ${folder.getId()}.`);
        }

        if (!folder) {
            Logger.log('[generateXlsReport] ERRO: A pasta de destino é nula ou indefinida.');
            throw new Error('Não foi possível obter ou criar a pasta de destino no Google Drive.');
        }

        Logger.log(`[generateXlsReport] ID da planilha temporária criada: ${spreadsheetId}`);

        const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;
        Logger.log(`[generateXlsReport] URL de exportação para XLSX: ${exportUrl}`);

        let xlsBlob;
        try {
            // Obtenha o token de acesso para autenticação
            const token = ScriptApp.getOAuthToken();
            Logger.log('[generateXlsReport] Token OAuth obtido com sucesso.');

            // Faça a requisição para exportar o arquivo como XLSX
            const response = UrlFetchApp.fetch(exportUrl, {
                headers: {
                    Authorization: `Bearer ${token}`
                },
                muteHttpExceptions: true // Para capturar erros na resposta
            });

            if (response.getResponseCode() === 200) {
                xlsBlob = response.getBlob();
                Logger.log(`[generateXlsReport] Blob XLS (exportado via URL) criado. Tipo: ${xlsBlob.getContentType()}, Tamanho: ${xlsBlob.getBytes().length} bytes.`);
            } else {
                const errorText = response.getContentText();
                Logger.log(`[generateXlsReport] Erro na resposta da exportação: Código ${response.getResponseCode()}, Texto: ${errorText}`);
                throw new Error(`Falha ao exportar a planilha para XLSX via URL. Código: ${response.getResponseCode()}, Mensagem: ${errorText}`);
            }

        } catch (exportError) {
            Logger.log(`[generateXlsReport] Erro durante a exportação via URLFetchApp: ${exportError.message}. Stack: ${exportError.stack}`);
            throw new Error(`Erro ao gerar relatório XLS (exportação): ${exportError.message}`);
        } finally {
            // Sempre mova a planilha temporária para a lixeira, independentemente do sucesso da exportação
            if (spreadsheetId) {
                try {
                    DriveApp.getFileById(spreadsheetId).setTrashed(true);
                    Logger.log(`[generateXlsReport] Planilha temporária '${spreadsheet.getName()}' (ID: ${spreadsheetId}) movida para a lixeira.`);
                } catch (cleanupError) {
                    Logger.log(`[generateXlsReport] Erro ao mover planilha temporária para a lixeira (ID: ${spreadsheetId}): ${cleanupError.message}`);
                }
            }
        }

        if (!xlsBlob) {
            Logger.log('[generateXlsReport] ERRO LÓGICO: xlsBlob é nulo após tentativa de exportação via URL.');
            throw new Error('Falha ao gerar relatório XLS: o blob resultante da exportação é nulo.');
        }

        const xlsFileName = `${ssTitle}.xlsx`;
        const xlsFile = folder.createFile(xlsBlob).setName(xlsFileName);
        Logger.log(`[generateXlsReport] Arquivo XLS final criado: ${xlsFile.getUrl()} (ID: ${xlsFile.getId()}).`);

        xlsFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileUrl = xlsFile.getDownloadUrl(); 
        Logger.log(`[generateXlsReport] URL de download do XLS: ${fileUrl}`);

        return fileUrl;

      } catch (e) {
        Logger.log(`[generateXlsReport] ERRO DURANTE A GERAÇÃO DO XLS: ${e.message}. Stack: ${e.stack}`);
        throw new Error(`Erro ao gerar relatório XLS: ${e.message}`);
      }
    }
