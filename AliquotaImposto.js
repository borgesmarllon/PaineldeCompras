/**
 * Grava os dados de um pedido na aba de cálculo, força o recálculo das fórmulas
 * e lê os resultados para retorná-los ao frontend.
 * @param {object} dadosPedido - Um objeto contendo os dados a serem gravados.
 * @returns {object} Um objeto com os resultados calculados pela planilha.
 */
function calcularImpostoPedidoCompleto(itensDoPedido, dadosGerais) {

  Logger.log("--- FUNÇÃO 'calcularImpostoPedidoCompleto' INICIADA ---");
  Logger.log("Parâmetro 'itensDoPedido' recebido: %s", JSON.stringify(itensDoPedido));
  Logger.log("Parâmetro 'dadosGerais' recebido: %s", JSON.stringify(dadosGerais));
   try {
    const planilhaAtual = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = planilhaAtual.getSheetByName('CalculoICMS');
    if (!sheet) throw new Error("Aba 'CalculoICMS' não encontrada.");

    // ====================================================================
    // ===== BUSCA SEGURA PELA LINHA TOTAL (A CORREÇÃO DEFINITIVA) =====
    // ====================================================================
    Logger.log("Iniciando busca pela linha TOTAL...");
    
    // Define o intervalo de busca para a coluna B, começando da linha 2.
    // Isso FORÇA o script a IGNORAR a linha 1 (cabeçalhos).
    // A célula mesclada B21:D21 com "TOTAL" será encontrada por esta busca.
    const rangeBusca = sheet.getRange("B2:B" + sheet.getLastRow()); 
    
    const textFinder = rangeBusca.createTextFinder("TOTAL")
    .matchEntireCell(true) // Procura pela célula com o texto exato
    .matchCase(false);    // Ignora maiúsculas/minúsculas
    
    const totalCell = textFinder.findNext();
    
    if (!totalCell) {
      // Se ainda não encontrar, o erro é claro.
      throw new Error("CRÍTICO: A célula com o texto 'TOTAL' não foi encontrada na coluna B (a partir da linha 2). Verifique se a célula B21 contém exatamente a palavra 'TOTAL'.");
    }
    
    const linhaTotal = totalCell.getRow();
    Logger.log(`SUCESSO: Linha TOTAL encontrada na linha correta: ${linhaTotal}`);
    // ====================================================================

    // O resto da sua função, que já está correto, continua a partir daqui...
    if (linhaTotal > 3) {
      sheet.getRange(`A2:C${linhaTotal - 1}`).clearContent();
      sheet.getRange(`D2:D${linhaTotal - 1}`).clearContent();
      sheet.getRange(`F2:F${linhaTotal - 1}`).clearContent();
    }

    const numeroDeItens = itensDoPedido.length;
    const linhasDisponiveis = linhaTotal - 2;
    if (numeroDeItens > linhasDisponiveis) {
      const linhasParaAdicionar = numeroDeItens - linhasDisponiveis;
      sheet.insertRowsBefore(linhaTotal, linhasParaAdicionar);
      const linhaModelo = linhaTotal - 1;
      const rangeModelo = sheet.getRange(linhaModelo, 1, 1, sheet.getLastColumn());
      const rangeDestino = sheet.getRange(linhaTotal, 1, linhasParaAdicionar, sheet.getLastColumn());
      rangeModelo.copyTo(rangeDestino);
    }

    itensDoPedido.forEach((item, index) => {
      const linhaAtual = 2 + index;
      sheet.getRange(linhaAtual, 1).setValue(dadosGerais.numeroNFE);
      sheet.getRange(linhaAtual, 3).setValue(dadosGerais.ufFornecedor);
      sheet.getRange(linhaAtual, 4).setValue(dadosGerais.regimeTributario);
      sheet.getRange(linhaAtual, 6).setValue(item.subtotal);
    });

    SpreadsheetApp.flush();
    
    // Reencontra a linha total para garantir que pegamos a posição certa após inserir linhas
    const novaLinhaTotal = sheet.getRange("B2:B" + sheet.getLastRow()).createTextFinder("TOTAL").findNext().getRow();
    const resultadosFinais = sheet.getRange(novaLinhaTotal, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const icmsStTotal = resultadosFinais[13];
    const diferencaIcmsSn = resultadosFinais[14];
    
    return {
      status: 'ok',
      resultados: { icmsStTotal: icmsStTotal, diferencaIcmsSn: diferencaIcmsSn }
    };

  } catch (e) {
    Logger.log("Erro em calcularImpostoPedidoCompleto: " + e.message);
    return { status: 'error', message: e.message };
  }
}

/**
 * Limpa a área de dados da planilha de cálculo, preservando os cabeçalhos,
 * a linha TOTAL e as fórmulas.
 */
/**
 * Limpa a área de dados da planilha de cálculo, preservando os cabeçalhos,
 * a linha TOTAL, as fórmulas e a formatação.
 */
function limparPlanilhaDeCalculo() {
  try {
    const planilhaAtual = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = planilhaAtual.getSheetByName('CalculoICMS');
    if (!sheet) throw new Error("Aba 'CalculoICMS' não encontrada.");

    const textFinder = sheet.createTextFinder("TOTAL").matchEntireCell(true).matchCase(false);
    const totalCell = textFinder.findNext();
    if (!totalCell) throw new Error("Linha 'TOTAL' não encontrada.");
    
    const linhaTotal = totalCell.getRow();

    if (linhaTotal > 2) {
      // Limpa apenas as colunas de ENTRADA na área de cálculo
      sheet.getRange(`A2:C${linhaTotal - 1}`).clearContent();
      sheet.getRange(`D2:D${linhaTotal - 1}`).clearContent();
    }
    
    return { status: 'ok', message: 'Planilha de cálculo limpa.' };
  } catch(e) {
    Logger.log("Erro ao limpar planilha de cálculo: " + e.message);
    return { status: 'error', message: e.message };
  }
}

/**
 * Função criada apenas para testar a 'calcularEGravarImposto' diretamente no editor,
 * sem precisar usar o App da Web.
 */
function testarCalculoDeImposto() {
  Logger.log("--- INICIANDO TESTE da função 'calcularImpostoPedidoCompleto' ---");

  // 1. Simula a LISTA DE ITENS que o frontend enviaria.
  //    O backend usará a propriedade 'totalItem' para o valor base de cada linha.
  const itensDeTeste = [
    { descricao: 'PRODUTO DE TESTE A', totalItem: 1550.75 },
    { descricao: 'PRODUTO DE TESTE B', totalItem: 3116.19 },
    { descricao: 'PRODUTO DE TESTE C', totalItem: 850.00 }
  ];
  
  // 2. Simula os DADOS GERAIS do pedido que o frontend enviaria.
  const dadosGeraisDeTeste = {
    numeroNFE: 'TESTE-CALCULO-789',
    ufFornecedor: 'SP', // Mude aqui para testar outros estados
    regimeTributario: 2   // Mude aqui: 1 para Simples Nacional, 2 para Outro
  };

  Logger.log("Enviando Itens: %s", JSON.stringify(itensDeTeste, null, 2));
  Logger.log("Enviando Dados Gerais: %s", JSON.stringify(dadosGeraisDeTeste, null, 2));

  // 3. Chama a sua função real com os dados de teste
  const resultado = calcularImpostoPedidoCompleto(itensDeTeste, dadosGeraisDeTeste);

  // 4. Mostra o resultado final no log
  Logger.log("--- RESULTADO RECEBIDO ---");
  Logger.log(JSON.stringify(resultado, null, 2));
}

/**
 * Busca a tabela de alíquotas por estado na aba 'Config'.
 * @returns {Object} Um objeto no formato { "BA": 0.18, "SP": 0.12, ... }
 */
function getAliquotasConfig() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
    if (!sheet) return {};

    const data = sheet.getRange(2, 5, numRows, 2).getValues();
    const aliquotasMap = {};
    
    data.forEach(row => {
      const estado = String(row[4]).trim().toUpperCase(); // Coluna A: Estado
      let aliquotaValor = row[5]; // Coluna B: Alíquota

      // Lógica robusta para converter a alíquota para número
      if (typeof aliquotaValor === 'string') {
        aliquotaValor = parseFloat(aliquotaValor.replace(',', '.'));
      }
      
      if (estado && typeof aliquotaValor === 'number' && !isNaN(aliquotaValor)) {
        aliquotasMap[estado] = aliquotaValor;
      }
    });
    
    return aliquotasMap;

  } catch (e) {
    Logger.log("Erro em getAliquotasConfig: " + e.message);
    return {};
  }
}
