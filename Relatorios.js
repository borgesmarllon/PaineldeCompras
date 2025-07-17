   // =================================================================
// FUNÇÕES PRINCIPAIS DE GERAÇÃO DE RELATÓRIO
// =================================================================

function generatePdfReport(reportParams) {
    try {
        // PASSO 1: O único objetivo é ver se a função é chamada.
        Logger.log("--- DIAGNÓSTICO PROFUNDO ---");
        Logger.log("✅ PASSO 1: A função 'generatePdfReport' foi chamada com sucesso.");
        
        // PASSO 2: Verificar o tipo e o conteúdo do parâmetro recebido.
        Logger.log(`✅ PASSO 2: Tipo do parâmetro 'reportParams' recebido: ${typeof reportParams}`);
        
        // PASSO 3: Tentar registrar o conteúdo do parâmetro.
        // Se reportParams for um objeto inválido, o JSON.stringify pode falhar,
        // o que nos daria uma pista importante.
        Logger.log(`✅ PASSO 3: Conteúdo de 'reportParams': ${JSON.stringify(reportParams, null, 2)}`);
        
        // Se chegou até aqui, a comunicação e os parâmetros estão funcionando.
        return { status: 'success', url: '#', message: 'Teste de comunicação e parâmetros bem-sucedido! Verifique os logs do backend.' };

    } catch (e) {
        Logger.log(`❌ ERRO NO TESTE DE DIAGNÓSTICO PROFUNDO: ${e.message}`);
        return { status: 'error', message: `Falha no teste de diagnóstico: ${e.message}` };
    }
}

function generateXlsReport(reportParams) {
    try {
        Logger.log(`[generateXlsReport] 1. Iniciando. Parâmetros: ${JSON.stringify(reportParams)}`);

        if (!reportParams || !reportParams.companyId) {
            throw new Error("Parâmetro 'companyId' é obrigatório.");
        }

        const allPedidos = _getPedidosData(reportParams);
        const filteredPedidos = _filterPedidos(allPedidos, reportParams);
        const reportData = _groupAndSummarizePedidos(filteredPedidos, reportParams);

        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
        const ss = SpreadsheetApp.create(`Relatorio_${reportParams.reportType}_${timestamp}`);
        const sheet = ss.getActiveSheet();
        _populateSheetWithReportData(sheet, reportData, reportParams);
        Logger.log(`[generateXlsReport] 5. Planilha populada.`);
        
        SpreadsheetApp.flush();

        const folder = _getOrCreateFolder("RelatoriosComprasTemporarios");
        const xlsxBlob = _exportSpreadsheetAsXlsx(ss.getId());
        const xlsxFile = folder.createFile(xlsxBlob).setName(ss.getName() + ".xlsx");

        xlsxFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileUrl = xlsxFile.getDownloadUrl();
        Logger.log(`[generateXlsReport] 6. Sucesso! URL do XLSX: ${fileUrl}`);

        DriveApp.getFileById(ss.getId()).setTrashed(true);

        return { status: 'success', url: fileUrl };

    } catch (e) {
        Logger.log(`[generateXlsReport] ERRO FATAL: ${e.message}. Stack: ${e.stack}`);
        return { status: 'error', message: `Erro ao gerar relatório XLSX: ${e.message}` };
    }
}

// =================================================================
// FUNÇÕES AUXILIARES COMPLETAS
// =================================================================

function _getPedidosData(reportParams) {
    Logger.log(`[_getPedidosData] 2. Buscando dados da planilha...`);
    const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
    if (!sheet || sheet.getLastRow() < 2) return [];

    const range = sheet.getDataRange();
    const values = range.getValues();
    const headers = values[0];
    const pedidos = [];

    const colEmpresa = headers.findIndex(h => ['EMPRESA', 'ID EMPRESA', 'ID DA EMPRESA'].includes(String(h).toUpperCase().trim()));
    if (colEmpresa === -1) throw new Error("Coluna da Empresa não encontrada na planilha 'Pedidos'.");

    const idEmpresaFiltro = String(reportParams.companyId).trim();

    for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (String(row[colEmpresa]).trim() !== idEmpresaFiltro) {
            continue;
        }

        const pedido = {};
        headers.forEach((header, index) => {
            pedido[toCamelCase(header)] = row[index];
        });

        if (pedido.data && !(pedido.data instanceof Date)) {
            try {
                let dateString = String(pedido.data);
                if (!dateString.includes('T') && !dateString.includes('Z')) {
                   dateString = dateString.replace(' ', 'T');
                }
                pedido.data = new Date(dateString);
            } catch (e) { pedido.data = null; }
        }
        
        if (typeof pedido.itens === 'string' && pedido.itens) {
            try {
                pedido.itens = JSON.parse(pedido.itens);
            } catch (e) { pedido.itens = []; }
        } else if (!Array.isArray(pedido.itens)) {
            pedido.itens = [];
        }

        pedido.totalGeral = parseFloat(pedido.totalGeral || 0);
        pedidos.push(pedido);
    }
    Logger.log(`[_getPedidosData] Encontrados ${pedidos.length} pedidos para a empresa ${idEmpresaFiltro}.`);
    return pedidos;
}

function _filterPedidos(allPedidos, reportParams) {
    let filtered = allPedidos;
    Logger.log(`[_filterPedidos] 3. Filtrando pedidos. Inicial: ${filtered.length}`);

    if (reportParams.startDate && reportParams.endDate) {
        const start = new Date(reportParams.startDate + 'T00:00:00');
        const end = new Date(reportParams.endDate + 'T23:59:59');
        filtered = filtered.filter(p => p.data && p.data >= start && p.data <= end);
    }

    if (reportParams.supplier) {
        const selectedSupplier = reportParams.supplier.toLowerCase();
        filtered = filtered.filter(p => (p.fornecedor || '').toLowerCase() === selectedSupplier);
    }
    Logger.log(`[_filterPedidos] Final: ${filtered.length} pedidos.`);
    return filtered;
}

function _groupAndSummarizePedidos(pedidos, reportParams) {
    Logger.log(`[_groupAndSummarizePedidos] 4. Agrupando ${pedidos.length} pedidos para relatório do tipo "${reportParams.reportType}".`);
    const data = {};
    if (reportParams.reportType === 'detailed') {
        pedidos.sort((a, b) => a.data.getTime() - b.data.getTime());
        pedidos.forEach(pedido => {
            const dateStr = Utilities.formatDate(pedido.data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            if (!data[dateStr]) data[dateStr] = { date: pedido.data, fornecedores: {}, totalDate: 0 };
            const fornecedor = pedido.fornecedor || 'Desconhecido';
            if (!data[dateStr].fornecedores[fornecedor]) data[dateStr].fornecedores[fornecedor] = { pedidos: [], totalFornecedor: 0 };
            data[dateStr].fornecedores[fornecedor].pedidos.push(pedido);
            data[dateStr].fornecedores[fornecedor].totalFornecedor += pedido.totalGeral;
            data[dateStr].totalDate += pedido.totalGeral;
        });
    } else if (reportParams.reportType === 'financial') {
        data.totalGeralPedidos = 0;
        data.numeroTotalPedidos = 0;
        data.totalPorFornecedor = {};
        pedidos.forEach(pedido => {
            data.totalGeralPedidos += pedido.totalGeral;
            data.numeroTotalPedidos++;
            const fornecedor = pedido.fornecedor || 'Desconhecido';
            data.totalPorFornecedor[fornecedor] = (data.totalPorFornecedor[fornecedor] || 0) + pedido.totalGeral;
        });
        data.listaTotalPorFornecedor = Object.keys(data.totalPorFornecedor).map(f => ({
            fornecedor: f,
            total: data.totalPorFornecedor[f]
        })).sort((a, b) => b.total - a.total);
    }
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
        let reportTitle = "Relatório de Compras " + (reportParams.reportType === 'detailed' ? "Detalhado" : "Financeiro");
        
        const reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
        let filtersHtml = `<p>Relatório Gerado em: ${reportDate}</p>`;
        if (reportParams.startDate) filtersHtml += `<p>Período: ${reportParams.startDate} a ${reportParams.endDate}</p>`;
        if (reportParams.supplier) filtersHtml += `<p>Fornecedor: ${reportParams.supplier}</p>`;

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
    return `<!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;font-size:10pt}table{width:100%;border-collapse:collapse}th,td{border:1px solid #ddd;padding:4px}th{background-color:#f2f2f2}</style></head><body><h1>${reportTitle}</h1><div>${filtersHtml}</div>${bodyContent}</body></html>`;

    }

    /**
 * Popula uma planilha do Google Sheets com os dados do relatório para exportação em XLSX.
 */
function _populateSheetWithReportData(sheet, reportData, reportParams) {
    let row = 1;
    sheet.getRange(row, 1).setValue(reportParams.companyName || "Relatório de Compras").setFontWeight('bold').setFontSize(14);
    row++;

    const reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    sheet.getRange(row, 1).setValue(`Relatório Gerado em: ${reportDate}`);
    row += 2;

    if (reportParams.startDate && reportParams.endDate) {
        sheet.getRange(row, 1).setValue(`Período: ${reportParams.startDate} a ${reportParams.endDate}`);
        row++;
    }
    if (reportParams.supplier) {
        sheet.getRange(row, 1).setValue(`Fornecedor: ${reportParams.supplier}`);
        row++;
    }
    row++;

    if (reportParams.reportType === 'detailed') {
        const headers = ['Data', 'Fornecedor', 'Nº Pedido', 'Item', 'Qtd.', 'Preço Unit.', 'Subtotal'];
        sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f2f2f2');
        row++;
        Object.values(reportData).forEach(dateGroup => {
            Object.values(dateGroup.fornecedores).forEach(supplierGroup => {
                supplierGroup.pedidos.forEach(pedido => {
                    pedido.itens.forEach(item => {
                        const rowData = [
                            Utilities.formatDate(pedido.data, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
                            pedido.fornecedor,
                            pedido.numeroDoPedido,
                            item.descricao,
                            item.quantidade,
                            item.precoUnitario,
                            item.totalItem
                        ];
                        sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
                        row++;
                    });
                });
            });
        });
        sheet.getRange(row, 6, sheet.getLastRow() - row + 1, 2).setNumberFormat('R$ #,##0.00');
        sheet.autoResizeColumns(1, headers.length);

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
 * Exporta uma Planilha Google como um blob XLSX.
 */
function _exportSpreadsheetAsXlsx(spreadsheetId) {
    Logger.log(`[_exportSpreadsheetAsXlsx] Exportando planilha ID: ${spreadsheetId}`);
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
    Logger.log(`[_exportSpreadsheetAsXlsx] Exportação concluída.`);
    return response.getBlob();
}

/**
 * Obtém ou cria uma pasta no Google Drive.
 */
function _getOrCreateFolder(folderName) {
    Logger.log(`[_getOrCreateFolder] Verificando/Criando pasta: "${folderName}"`);
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
        const folder = folders.next();
        Logger.log(`[_getOrCreateFolder] Pasta encontrada.`);
        return folder;
    }
    const newFolder = DriveApp.createFolder(folderName);
    Logger.log(`[_getOrCreateFolder] Pasta criada.`);
    return newFolder;
}

function testBackendConnection() {
    Logger.log("✅ [testBackendConnection] A função de teste foi chamada com sucesso!");
    return { 
        status: "ok", 
        message: "Backend respondeu!", 
        timestamp: new Date().toLocaleString('pt-BR') 
    };
}
    // ========================

    /**
 * Serve o conteúdo HTML da tela de Relatórios para o frontend.
 * @returns {string} Conteúdo HTML.
 */
function getRelatoriosHtmlContent() { // Esta função será chamada pelo frontend
  Logger.log('[backend] getRelatoriosHtmlContent: Servindo o HTML da tela de Relatórios.');
  return HtmlService.createTemplateFromFile('Relatorios') // Certifique-se que 'Relatorios' é o nome do seu arquivo HTML
      .evaluate() // IMPORTANTE: .evaluate() AQUI para injetar google.script.run no HTML
      .getContent(); // .getContent() para retornar o HTML como string
}

// Sua função doGet() principal deve servir a tela inicial do seu SPA (ex: Dashboard ou Menu)
function doGet() {
  Logger.log('[backend] doGet: Servindo a tela inicial.');
  return HtmlService.createTemplateFromFile('Login') // Ou 'Menu', ou 'Login', a tela que inicia seu app
      .evaluate() // ESSENCIAL aqui para que google.script.run funcione em TODO o seu SPA
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}
