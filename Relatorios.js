   // =================================================================
// FUNÇÕES PRINCIPAIS DE GERAÇÃO DE RELATÓRIO
// =================================================================

function generatePdfReport(reportParams) {
     try {
        Logger.log(`[generatePdfReport] 1. Iniciando. Parâmetros: ${JSON.stringify(reportParams)}`);
        
        if (!reportParams || !reportParams.companyId) {
            throw new Error("Parâmetro 'companyId' é obrigatório.");
        }

        const allPedidos = _getPedidosData(reportParams);
        const filteredPedidos = _filterPedidos(allPedidos, reportParams);
        const reportData = _groupAndSummarizePedidos(filteredPedidos, reportParams);
        const htmlContent = _generatePdfHtmlContent(reportData, reportParams);
        Logger.log(`[generatePdfReport] 5. Conteúdo HTML gerado.`);

        const htmlBlob = Utilities.newBlob(htmlContent, MimeType.HTML, 'Relatorio.html');
        const pdfBlob = htmlBlob.getAs(MimeType.PDF);
        Logger.log(`[generatePdfReport] 6. Blob PDF criado.`);

        const folder = _getOrCreateFolder("RelatoriosComprasTemporarios");
        const fileName = `Relatorio_${reportParams.reportType}_${new Date().getTime()}.pdf`;
        const pdfFile = folder.createFile(pdfBlob).setName(fileName);
        
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileUrl = pdfFile.getDownloadUrl();
        Logger.log(`[generatePdfReport] 7. Sucesso! URL do PDF: ${fileUrl}`);
        
        return { status: 'success', url: fileUrl };

    } catch (e) {
        Logger.log(`[generatePdfReport] ERRO FATAL: ${e.message}. Stack: ${e.stack}`);
        // Retorna um objeto de erro em vez de lançar uma exceção
        return { status: 'error', message: `Erro ao gerar relatório PDF: ${e.message}` };
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

    const indexNumeroPedido = headers.findIndex(h => ['NÚMERO DO PEDIDO', 'NUMERO DO PEDIDO', 'NUMERO PEDIDO'].includes(String(h).toUpperCase().trim()));
    const idEmpresaFiltro = String(reportParams.companyId).trim();

    for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (parseInt(row[colEmpresa], 10) !== parseInt(idEmpresaFiltro, 10)) {
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
    const companyAddress = reportParams.companyAddress || "";
    const companyCnpj = reportParams.empresaCnpj || "";
    const reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

    let reportTitle = "Relatório de Compras " + (reportParams.reportType === 'detailed' ? "Detalhado" : "Financeiro");
    let filtersHtml = '';
    if (reportParams.startDate && reportParams.endDate) {
        filtersHtml += `<strong>Período:</strong> ${Utilities.formatDate(new Date(reportParams.startDate + 'T00:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy')} a ${Utilities.formatDate(new Date(reportParams.endDate + 'T00:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy')}<br>`;
    }
    if (reportParams.supplier) {
        filtersHtml += `<strong>Fornecedor:</strong> ${reportParams.supplier}`;
    }

    let bodyContent = '';

    if (reportParams.reportType === 'detailed') {
        if (Object.keys(reportData).length === 0) {
            bodyContent = '<p class="no-data">Nenhum pedido encontrado para os filtros selecionados.</p>';
        } else {
            const sortedDates = Object.keys(reportData).sort();
            sortedDates.forEach(dateStr => {
                const dateGroup = reportData[dateStr];
                bodyContent += `<div class="date-group"><h3>Data: ${Utilities.formatDate(dateGroup.date, Session.getScriptTimeZone(), 'dd/MM/yyyy')} &nbsp;&nbsp;<span class="total-day">Total do Dia: R$ ${dateGroup.totalDate.toFixed(2).replace('.', ',')}</span></h3>`;
                const sortedSuppliers = Object.keys(dateGroup.fornecedores).sort();
                sortedSuppliers.forEach(supplierName => {
                    const supplierGroup = dateGroup.fornecedores[supplierName];
                    bodyContent += `<div class="supplier-group"><h4>Fornecedor: ${supplierName} &nbsp;&nbsp;<span class="total-supplier">Total: R$ ${supplierGroup.totalFornecedor.toFixed(2).replace('.', ',')}</span></h4><table><thead><tr><th>Nº Pedido</th><th>Item</th><th>Qtd.</th><th>Preço Unit.</th><th>Subtotal</th></tr></thead><tbody>`;
                    supplierGroup.pedidos.forEach(pedido => {
                        pedido.itens.forEach(item => {
                            bodyContent += `<tr>
                                <td>${pedido.numeroDoPedido || ''}</td>
                                <td>${item.descricao || ''}</td>
                                <td class="text-right">${item.quantidade || 0}</td>
                                <td class="text-right">R$ ${parseFloat(item.precoUnitario || 0).toFixed(2).replace('.', ',')}</td>
                                <td class="text-right">R$ ${parseFloat(item.totalItem || 0).toFixed(2).replace('.', ',')}</td>
                            </tr>`;
                        });
                    });
                    bodyContent += `</tbody></table></div>`;
                });
                bodyContent += `</div>`;
            });
        }
    } else if (reportParams.reportType === 'financial') {
        if (reportData.numeroTotalPedidos === 0) {
            bodyContent = '<p class="no-data">Nenhum dado financeiro encontrado para os filtros selecionados.</p>';
        } else {
            bodyContent += `<div class="summary-box">
                <p><strong>Total de Pedidos:</strong> ${reportData.numeroTotalPedidos}</p>
                <p><strong>Valor Total das Compras:</strong> <span class="total-value">R$ ${reportData.totalGeralPedidos.toFixed(2).replace('.', ',')}</span></p>
            </div>`;
            bodyContent += `<h4>Totais por Fornecedor:</h4><table><thead><tr><th>Fornecedor</th><th class="text-right">Valor Total Comprado</th></tr></thead><tbody>`;
            reportData.listaTotalPorFornecedor.forEach(item => {
                bodyContent += `<tr><td>${item.fornecedor}</td><td class="text-right">R$ ${item.total.toFixed(2).replace('.', ',')}</td></tr>`;
            });
            bodyContent += `</tbody></table>`;
        }
    }

    const html = `
    <!DOCTYPE html><html><head><title>${reportTitle}</title>
    <style>
        @page { size: A4; margin: 1cm; }
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; font-size: 10pt; color: #333; }
        .header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #004a99; padding-bottom: 10px; margin-bottom: 20px; }
        .company-info h1 { margin: 0; color: #004a99; font-size: 18pt; }
        .company-info p { margin: 2px 0; font-size: 9pt; color: #555; }
        .report-info { text-align: right; font-size: 9pt; color: #555; }
        .report-title { text-align: center; font-size: 16pt; margin-bottom: 5px; color: #333; font-weight: bold; }
        .filters { font-size: 8pt; text-align: center; margin-bottom: 20px; color: #666; background-color: #f9f9f9; padding: 8px; border-radius: 4px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { border: 1px solid #e0e0e0; padding: 8px; text-align: left; }
        th { background-color: #f2f7fc; font-weight: bold; color: #004a99; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .text-right { text-align: right; }
        .date-group { margin-bottom: 25px; page-break-inside: avoid; }
        .date-group h3 { margin-top: 0; color: #004a99; font-size: 12pt; border-bottom: 1px solid #ccc; padding-bottom: 5px; }
        .supplier-group { margin-left: 15px; margin-top: 15px; }
        .supplier-group h4 { margin-top: 0; color: #0056b3; font-size: 11pt; }
        .total-day, .total-supplier, .total-value { font-weight: bold; }
        .no-data { text-align: center; font-style: italic; color: #888; margin-top: 30px; }
        .summary-box { background-color: #f2f7fc; border: 1px solid #c9deff; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .footer { position: fixed; bottom: -20px; left: 0; right: 0; text-align: center; font-size: 8pt; color: #999; }
    </style></head><body>
    <div class="header">
        <div class="company-info">
            <h1>${companyName}</h1>
            <p>${companyAddress}</p>
            <p>CNPJ: ${companyCnpj}</p>
        </div>
        <div class="report-info">
            <strong>${reportTitle}</strong><br>
            Gerado em: ${reportDate}
        </div>
    </div>
    <div class="filters">${filtersHtml}</div>
    <div class="content">${bodyContent}</div>
    <div class="footer"><p>Página <span class="pageNumber"></span> de <span class="totalPages"></span></p></div>
    </body></html>`;
    return html;
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
        
        const sortedDates = Object.keys(reportData).sort();
        sortedDates.forEach(dateStr => {
            const dateGroup = reportData[dateStr];
            const sortedSuppliers = Object.keys(dateGroup.fornecedores).sort();
            sortedSuppliers.forEach(supplierName => {
                const supplierGroup = dateGroup.fornecedores[supplierName];
                supplierGroup.pedidos.forEach(pedido => {
                    pedido.itens.forEach(item => {
                        const rowData = [
                            Utilities.formatDate(pedido.data, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
                            supplierName,
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
        sheet.getRange(10, 5, Math.max(1, row - 10), 3).setNumberFormat('R$ #,##0.00');
        sheet.autoResizeColumns(1, headers.length);

    } else if (reportParams.reportType === 'financial') {
        sheet.getRange(row, 1).setValue('Total de Pedidos:').setFontWeight('bold');
        sheet.getRange(row, 2).setValue(reportData.numeroTotalPedidos);
        row++;
        sheet.getRange(row, 1).setValue('Valor Total das Compras:').setFontWeight('bold');
        sheet.getRange(row, 2).setValue(reportData.totalGeralPedidos).setNumberFormat('R$ #,##0.00');
        row += 2;

        const headers = ['Fornecedor', 'Valor Total Comprado'];
        sheet.getRange(row, 1, 1, 2).setValues([headers]).setFontWeight('bold').setBackground('#f2f2f2');
        row++;
        
        reportData.listaTotalPorFornecedor.forEach(item => {
            sheet.getRange(row, 1, 1, 2).setValues([[item.fornecedor, item.total]]);
            sheet.getRange(row, 2).setNumberFormat('R$ #,##0.00');
            row++;
        });
        sheet.autoResizeColumns(1, 2);
    }
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
