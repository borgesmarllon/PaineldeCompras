/**
 * =================================================================================
 * NOVA FUNÇÃO PARA DESMEMBRAR O JSON DE ITENS
 * =================================================================================
 */


function getConfig() {
    // Objeto de configuração integrado com os nomes das suas planilhas
    const CONFIG = {
        PLANILHA_ID: '1xVLFSqL5SVT6cmZ_9foOkKxJIHqHeCGHLcyryBJ44g0',
        sheets: {
            pedidos: 'Pedidos',
            usuarios: 'Usuarios',
            fornecedores: 'Fornecedores',
            config: 'Config',
            empresas: 'Empresas',
            usuario_logado: 'usuario_logado',
            calculoIcms: 'Calculo ICMS',
            aviso: 'Aviso',
            itens: 'Itens Pedido' // Corrigido para o nome correto da planilha
        },
        GEMINI_API_KEY_PROPERTY: 'GEMINI_API_KEY',
        DRIVE_FOLDER_REPORTS: "RelatoriosComprasTemporarios",
        status: { // Adicionado para compatibilidade com o resto do código
            aguardandoAprovacao: 'Aguardando Aprovacao'
        }
    };
    return CONFIG;
}

/**
 * Pega uma string JSON de itens, a converte em objetos e salva cada item
 * em uma nova linha na planilha 'Itens'.
 * @param {string} numeroPedido - O número do pedido para vincular os itens.
 * @param {string} empresaId - O ID da empresa para vincular os itens.
 * @param {string} itensJsonString - A string JSON contendo o array de itens.
 */
function desmembrarJsonDeItens(numeroPedido, empresaId, itensJsonString, estadoFornecedor, aliquotaImposto, icmsStTotal) {
    try {
        Logger.log(`Iniciando desmembramento de itens para o pedido ${numeroPedido}`);
        const config = getConfig();
        const sheetItens = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheets.itens);
        if (!sheetItens) {
            throw new Error(`Planilha de itens com o nome "${config.sheets.itens}" não foi encontrada.`);
        }

        const itens = JSON.parse(itensJsonString);

        if (!Array.isArray(itens) || itens.length === 0) {
            Logger.log(`Nenhum item para desmembrar para o pedido ${numeroPedido}.`);
            return;
        }

        const rowsToAdd = itens.map(item => {
            // A ordem aqui deve corresponder exatamente à ordem das colunas na sua planilha 'Itens'
            return [
                `'${numeroPedido}`,
                `'${empresaId}`,
                item.descricao || '',
                item.produtoFornecedor || '',
                item.quantidade || 0,
                item.unidade || '',
                item.precoUnitario || 0,
                item.totalItem || 0,
                item.icmsSt || 0,
                estadoFornecedor || '',
                aliquotaImposto || 0,
                icmsStTotal || 0
            ];
        });

        // Adiciona todas as novas linhas de uma só vez para melhor performance
        sheetItens.getRange(sheetItens.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);

        Logger.log(`✅ ${rowsToAdd.length} itens do pedido ${numeroPedido} foram salvos com sucesso na planilha 'Itens'.`);

    } catch (e) {
        Logger.log(`❌ ERRO em desmembrarJsonDeItens para o pedido ${numeroPedido}: ${e.message}`);
        // É importante que este erro não pare o fluxo principal, então apenas registramos o log.
    }
}

/**
 * Testa a função de desmembrar itens para pedidos específicos.
 * Você deve rodar esta função diretamente do editor do Apps Script.
 */
function testarDesmembramentoDeItens() {
    Logger.log("--- INICIANDO TESTE DE DESMEMBRAMENTO DE ITENS ---");

    // <<< IMPORTANTE: Substitua estes valores pelos IDs dos pedidos que você quer testar >>>
    const pedidosParaTestar = [
        { numero: '001346', empresa: '001' },
    
    ];

    pedidosParaTestar.forEach(infoPedido => {
        Logger.log(`\n--- Processando Pedido: ${infoPedido.numero} | Empresa: ${infoPedido.empresa} ---`);
        
        // Simula a busca do objeto completo do pedido, como se ele viesse do seu frontend/banco de dados
        // Em um cenário real, você poderia usar uma função como 'getPedidoCompletoPorId'
        const pedidoCompleto = _simularBuscaPedido(infoPedido.numero, infoPedido.empresa);

        if (pedidoCompleto) {
            Logger.log(`Pedido encontrado. JSON de itens: ${pedidoCompleto.itens}`);
            
            // Chama a função de desmembramento com os dados do pedido encontrado
            desmembrarJsonDeItens(
                pedidoCompleto.numero_do_pedido,
                pedidoCompleto.empresa,
                pedidoCompleto.itens, // Passa a string JSON
                pedidoCompleto.estado_fornecedor,
                pedidoCompleto.aliquota_imposto,
                pedidoCompleto.icms_st_total
            );
        } else {
            Logger.log(`AVISO: Pedido ${infoPedido.numero} não encontrado na simulação. Pulando.`);
        }
    });

    Logger.log("\n--- TESTE DE DESMEMBRAMENTO CONCLUÍDO ---");
}

function _simularBuscaPedido(numeroPedido, empresaId) {
    const sheet = SpreadsheetApp.getActive().getSheetByName(getConfig().sheets.pedidos);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    const colNumero = headers.indexOf("Número do Pedido");
    const colEmpresa = headers.indexOf("Empresa");
    
    const pedidoRow = data.find(row => 
        String(row[colNumero]).replace("'", "") === String(numeroPedido) && 
        String(row[colEmpresa]).replace("'", "") === String(empresaId)
    );

    if (!pedidoRow) return null;

    // Monta um objeto simples com os campos necessários para o desmembramento
    const pedidoCompleto = {};
    headers.forEach((header, index) => {
        const key = header.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '_').toLowerCase();
        pedidoCompleto[key] = pedidoRow[index];
    });
    
    return pedidoCompleto;
}
