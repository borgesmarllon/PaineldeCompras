   // =================================================================
// FUNÇÕES PRINCIPAIS DE GERAÇÃO DE RELATÓRIO
// =================================================================

function generatePdfReport(reportParams) {
     try {
        Logger.log(`[generatePdfReport] 1. Iniciando. Parâmetros: ${JSON.stringify(reportParams)}`);
        
        const allPedidos = _getPedidosData(reportParams);
        const filteredPedidos = _filterPedidos(allPedidos, reportParams);
        const reportData = _groupAndSummarizePedidos(filteredPedidos, reportParams);

        // ==========================================================
        // ✅ 1. RAIO-X DOS DADOS (ANTES DE GERAR O HTML)
        // ==========================================================
        Logger.log("--- DEBUG: DADOS ANTES DE GERAR O HTML ---");
        // Usamos JSON.stringify para ver a estrutura completa do objeto
        Logger.log(JSON.stringify(reportData, null, 2));
        Logger.log("-----------------------------------------");
        // ==========================================================

        const htmlContent = _generatePdfHtmlContent(reportData, reportParams);
        Logger.log(`[generatePdfReport] 5. Conteúdo HTML gerado.`);

        // ==========================================================
        // ✅ 2. RAIO-X DO HTML (O RESULTADO FINAL)
        // ==========================================================
        Logger.log("--- DEBUG: CONTEÚDO HTML GERADO ---");
        Logger.log(htmlContent);
        Logger.log("-----------------------------------");
        // ==========================================================

        const pdfBlob = Utilities.newBlob(htmlContent, MimeType.HTML).getAs(MimeType.PDF).setName(`Relatorio.pdf`);
        const folder = _getOrCreateFolder("RelatoriosComprasTemporarios");
        const pdfFile = folder.createFile(pdfBlob);
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        Utilities.sleep(1000); 
        const fileUrl = `https://drive.google.com/uc?export=download&id=${pdfFile.getId()}`;
        
        Logger.log(`[generatePdfReport] 7. Sucesso! URL do PDF: ${fileUrl}`);
        return { status: 'success', url: fileUrl };

    } catch (e) {
        Logger.log(`[generatePdfReport] ERRO FATAL: ${e.message}. Stack: ${e.stack}`);
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

    if (reportParams.status) {
        const selectedStatus = reportParams.status.toLowerCase();
        filtered = filtered.filter(p => (p.status || '').toLowerCase() === selectedStatus);
    }
    Logger.log(`[_filterPedidos] Final: ${filtered.length} pedidos.`);
    return filtered;
}

function _groupAndSummarizePedidos(pedidos, reportParams) {
    Logger.log(`[_groupAndSummarizePedidos] 4. Agrupando ${pedidos.length} pedidos para relatório do tipo "${reportParams.reportType}".`);
    const data = {};
    if (reportParams.reportType === 'detailed') {
      const grupos = {};
      let valorTotalPeriodo = 0; 
      let valorTotalIcmsPeriodo = 0;

        pedidos.sort((a, b) => a.data.getTime() - b.data.getTime());

        pedidos.forEach(pedido => {
          valorTotalPeriodo += pedido.totalGeral;
          valorTotalIcmsPeriodo += pedido.icmsStTotal || 0;

            const dateStr = Utilities.formatDate(pedido.data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            if (!grupos[dateStr]) {
              grupos[dateStr] = { date: pedido.data, fornecedores: {}, totalDate: 0 };
            }
            const fornecedor = pedido.fornecedor || 'Desconhecido';
            if (!grupos[dateStr].fornecedores[fornecedor]) {
            grupos[dateStr].fornecedores[fornecedor] = { pedidos: [], totalFornecedor: 0 };
            }
            grupos[dateStr].fornecedores[fornecedor].pedidos.push(pedido);
            grupos[dateStr].fornecedores[fornecedor].totalFornecedor += pedido.totalGeral;
            grupos[dateStr].totalDate += pedido.totalGeral;
        });
        return {
            valorTotalPeriodo: valorTotalPeriodo,
            valorTotalIcmsPeriodo: valorTotalIcmsPeriodo,
            grupos: grupos
        };

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
        
        return data;
        }

        return{};
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
        if (!reportData || !reportData.grupos || Object.keys(reportData.grupos).length === 0) {
            bodyContent = '<p class="no-data">Nenhum pedido encontrado para os filtros selecionados.</p>';
        } else {
            // ✅ 1. ADICIONA A CAIXA DE RESUMO COM O VALOR TOTAL DO PERÍODO
            if (reportData.valorTotalPeriodo) {
                bodyContent += `
                <div class="summary-box">                
                  <p><strong>Valor Total do Período:</strong> <span class="total-value">R$ ${_formatCurrency(reportData.valorTotalPeriodo)}</span></p>
                  <p><strong>Valor Total ICMS ST do Período:</strong> <span class="total-value">R$ ${_formatCurrency(reportData.valorTotalIcmsPeriodo)}</span></p>
                </div>`;
            }
            const sortedDates = Object.keys(reportData.grupos).sort();
            sortedDates.forEach(dateStr => {
                const dateGroup = reportData.grupos[dateStr];
                bodyContent += `<div class="date-group">
                <h3>Data: ${Utilities.formatDate(dateGroup.date, Session.getScriptTimeZone(), 'dd/MM/yyyy')}</h3> &nbsp;&nbsp;<div class="total-day">Total do Dia: R$ ${_formatCurrency(dateGroup.totalDate)}</div>`;
                const sortedSuppliers = Object.keys(dateGroup.fornecedores).sort();
                sortedSuppliers.forEach(supplierName => {
                    const supplierGroup = dateGroup.fornecedores[supplierName];
                    const totalIcmsFornecedor = supplierGroup.pedidos.reduce((sum, pedido) => sum + (pedido.icmsStTotal || 0), 0);                    
                    bodyContent += `<div class="supplier-group">
                    <h4>Fornecedor: ${supplierName} &nbsp;&nbsp;
                    <span class="total-supplier">Total: R$ ${_formatCurrency(supplierGroup.totalFornecedor)}</span>
                    <span class="total-supplier">ICMS ST: R$ ${_formatCurrency(totalIcmsFornecedor)}</span>                    
                    </h4>
                    <table>
                    <thead>
                    <tr>
                    <th>Nº Pedido</th>
                    <th>Item</th>
                    <th class="text-right">Qtd.</th>
                    <th class="text-right">Preço Unit.</th>
                    <th class="text-right">Subtotal</th>    
                    <th class="text-right">Vlr. Icms ST</th>                
                    </tr>
                    </thead>
                    <tbody>`;
                    supplierGroup.pedidos.forEach(pedido => {
                      const icmsPedidoFormatado = `R$ ${parseFloat(pedido.icmsStTotal || 0).toFixed(2).replace('.', ',')}`;
                        pedido.itens.forEach(item => {
                            bodyContent += `<tr>
                                <td>${pedido.numeroDoPedido || ''}</td>
                                <td>${item.descricao || ''}</td>
                                <td class="text-right">${item.quantidade || 0}</td>
                                <td class="text-right">R$ ${_formatCurrency(item.precoUnitario)}</td>
                                <td class="text-right">R$ ${_formatCurrency(item.totalItem)}</td>
                                <td></td>
                            </tr>`;
                        });
                        
                        const subtotalCalculado = pedido.itens.reduce((sum, item) => sum + (parseFloat(item.totalItem) || 0), 0);
                        const subtotalPedidoFormatado = `R$ ${subtotalCalculado.toFixed(2).replace('.', ',')}`;                        
                        bodyContent += `<tr class="order-total-row">
                            <td colspan="4"><strong>Total do Pedido ${pedido.numeroDoPedido}</strong></td>
                            <td class="text-right"><strong>${subtotalPedidoFormatado}</strong></td>
                            <td class="text-right"><strong>${icmsPedidoFormatado}</strong></td>
                        </tr>`;
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
            bodyContent += `<h4>Totais por Fornecedor:</h4>
            <table><thead><tr><th>Fornecedor</th><th class="text-right">Valor Total Comprado</th></tr></thead><tbody>`;
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
        .date-group { 
          margin-bottom: 25px; 
          padding-top: 15px;
          border-top: 2px solid #004a99;
          page-break-inside: avoid; 
          }
        .date-group:first-child{
          margin-top: 0;
          padding-top: 0;
          border-top: none;
        }
        .date-group h3 { 
          margin-top: 0; 
          color: #004a99; 
          font-size: 12pt; 
          border-bottom: 1px 
          solid #ccc; 
          padding-bottom: 5px; 
          }
        .supplier-group { margin-left: 15px; margin-top: 15px; }
        .supplier-group h4 { margin-top: 0; color: #0056b3; font-size: 11pt; }
        .total-day, .total-supplier, .total-value { font-weight: bold; }
        .no-data { text-align: center; font-style: italic; color: #888; margin-top: 30px; }
        .summary-box { background-color: #f2f7fc; border: 1px solid #c9deff; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .footer { position: fixed; bottom: -20px; left: 0; right: 0; text-align: center; font-size: 8pt; color: #999; }
        .order-total-row{
          background-color: #e9eef5 !important; /* Usa !important para sobrescrever o nth-child(even) */
          font-weigth: bold;
          color: #333;
        }
        .order-total-row td{
          border-top: 2px solid #ccc;
        }
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
        const headers = ['Data', 'Fornecedor', 'Nº Pedido', 'Item', 'Qtd.', 'Preço Unit.', 'Subtotal', 'Valor Icms'];
        sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f2f2f2');
        row++;
        
        const grupos = reportData.grupos || {};
        const sortedDates = Object.keys(grupos).sort();
        sortedDates.forEach(dateStr => {
            const dateGroup = grupos[dateStr];
            const fornecedores = dateGroup.fornecedores || {};
            const sortedSuppliers = Object.keys(dateGroup.fornecedores).sort();
            sortedSuppliers.forEach(supplierName => {
                const supplierGroup = fornecedores[supplierName];
                if (!supplierGroup || !supplierGroup.pedidos) return;
                    supplierGroup.pedidos.forEach(pedido => {
                    pedido.itens.forEach(item => {
                        const rowData = [
                            Utilities.formatDate(pedido.data, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
                            supplierName,
                            pedido.numeroDoPedido,
                            item.descricao,
                            item.quantidade,
                            item.precoUnitario,
                            item.totalItem,
                            pedido.valorIcms
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
function _formatCurrency(value) {
  const number = Number(value) || 0;
  return number.toLocaleString('pt-BR', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}
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

function testRelatorioPdf() {
  // Parâmetros de teste
  const params = {
    companyId: "001",           // Altere conforme necessário
    startDate: "2025-08-01",    // Altere conforme necessário
    endDate: "2025-08-12",      // Altere conforme necessário
    reportType: "detailed"      // Ou "financial"
  };
  const resultado = generatePdfReport(params);
  Logger.log("Retorno da função generatePdfReport:");
  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}
