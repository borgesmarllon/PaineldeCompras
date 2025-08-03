/**
 * Grava os dados de um pedido na aba de cálculo, força o recálculo das fórmulas
 * e lê os resultados TOTAIS e INDIVIDUAIS para retorná-los ao frontend.
 * @param {object} dadosPedido - Um objeto contendo os dados a serem gravados.
 * @returns {object} Um objeto com os resultados calculados pela planilha.
 */
function calcularImpostoPedidoCompleto(itensDoPedido, dadosGerais) {

  Logger.log("--- FUNÇÃO 'calcularImpostoPedidoCompleto' INICIADA ---");
  Logger.log("Parâmetro 'itensDoPedido' recebido: %s", JSON.stringify(itensDoPedido));
  Logger.log("Parâmetro 'dadosGerais' recebido: %s", JSON.stringify(dadosGerais));
   try {
    const planilhaCalculo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CalculoICMS');
    if (!planilhaCalculo) {
      throw new Error("A aba 'CalculoICMS' não foi encontrada.");
    }

    Logger.log("Preparando a planilha de cálculo...");
    
    // Define o intervalo de busca para a coluna B, começando da linha 2.
    // Isso FORÇA o script a IGNORAR a linha 1 (cabeçalhos).
    // A célula mesclada B21:D21 com "TOTAL" será encontrada por esta busca.
    const rangeBusca = planilhaCalculo.getRange("B2:B" + planilhaCalculo.getMaxRows());
    const textFinder = rangeBusca.createTextFinder("TOTAL").matchEntireCell(true).matchCase(false);
    const totalCell = textFinder.findNext();
    if (!totalCell) {
      throw new Error("Célula 'TOTAL' não foi encontrada na coluna B da aba 'CalculoICMS'.");
    }
    const linhaTotal = totalCell.getRow();
    Logger.log(`Linha TOTAL encontrada na linha: ${linhaTotal}`);
    // ====================================================================
    // Limpa apenas as colunas de entrada de dados da área de cálculo.
    if (linhaTotal > 2) {
      planilhaCalculo.getRange(`A2:A${linhaTotal - 1}`).clearContent(); // Coluna A (NFE)
      planilhaCalculo.getRange(`C2:D${linhaTotal - 1}`).clearContent(); // Colunas C e D (UF e Regime)
      planilhaCalculo.getRange(`F2:F${linhaTotal - 1}`).clearContent(); // Coluna F (Valor do Item)
      Logger.log(`Colunas de entrada A, C, D, F limpas, preservando B e E.`);
    }
// Se o número de itens for maior que o espaço disponível, insere novas linhas.
    const espacoDisponivel = linhaTotal - 2;
    if (itensDoPedido.length > espacoDisponivel) {
      const linhasParaAdicionar = itensDoPedido.length - espacoDisponivel;
      planilhaCalculo.insertRowsBefore(linhaTotal, linhasParaAdicionar);
      Logger.log(`${linhasParaAdicionar} linha(s) adicionada(s) para comportar os itens.`);
    }

    // --- 2. ESCRITA DOS DADOS DE ENTRADA ---
    if (itensDoPedido.length > 0) {
      Logger.log(`Escrevendo ${itensDoPedido.length} itens na planilha em colunas separadas...`);
      const numItens = itensDoPedido.length;

      // Prepara os dados para cada coluna/bloco que será escrito
      const dadosColunaA = itensDoPedido.map(item => [dadosGerais.numeroNFE || 'CALCULO_TEMP']);
      const dadosColunasCD = itensDoPedido.map(item => [dadosGerais.ufFornecedor || '', dadosGerais.regimeTributario || '']);
      const dadosColunaF = itensDoPedido.map(item => [item.subtotal || item.totalItem || 0]);

      // Escreve na planilha em chamadas separadas, pulando as colunas B e E
      planilhaCalculo.getRange(2, 1, numItens, 1).setValues(dadosColunaA);  // Escreve na Coluna A
      planilhaCalculo.getRange(2, 3, numItens, 2).setValues(dadosColunasCD); // Escreve nas Colunas C e D
      planilhaCalculo.getRange(2, 6, numItens, 1).setValues(dadosColunaF);  // Escreve na Coluna F
      
      Logger.log("Dados dos itens escritos com sucesso, preservando as colunas B e E.");
    }

    // Força a atualização de todas as fórmulas na planilha. É um passo crucial.
    SpreadsheetApp.flush();
    Utilities.sleep(1000); // Adiciona uma pequena pausa para garantir que os cálculos mais complexos terminem.

    // --- 3. LEITURA DOS RESULTADOS CALCULADOS ---
    Logger.log("Lendo resultados da planilha...");
    
    // Reencontra a linha total, caso novas linhas tenham sido inseridas.
    // Usamos o mesmo textFinder de antes para garantir que ele continue a busca de onde parou.
    Logger.log("Reiniciando busca pela linha TOTAL para leitura final...");
    const rangeBuscaFinal = planilhaCalculo.getRange("B2:B" + planilhaCalculo.getMaxRows());
    const finalFinder = rangeBuscaFinal.createTextFinder("TOTAL").matchEntireCell(true).matchCase(false);
    const finalTotalCell = finalFinder.findNext();

    if (!finalTotalCell) {
       throw new Error("CRÍTICO: Não foi possível reencontrar a célula 'TOTAL' após a escrita dos dados.");
    }
    const linhaTotalFinal = finalTotalCell.getRow();
    Logger.log(`Leitura final será feita na linha TOTAL: ${linhaTotalFinal}`);

    const resultadosTotais = planilhaCalculo.getRange(linhaTotalFinal, 1, 1, planilhaCalculo.getLastColumn()).getValues()[0];

    // **IMPORTANTE**: Verifique se os índices (13 e 14) correspondem às colunas corretas (N e O) na sua linha de TOTAL.
    const icmsStTotalCalculado = resultadosTotais[13] || 0;     // Coluna N
    const diferencaIcmsSnCalculada = resultadosTotais[14] || 0; // Coluna O
    
    Logger.log(`Total ICMS ST lido da linha ${linhaTotalFinal}: ${icmsStTotalCalculado}`);
    
    const icmsPorItemCalculado = [];
    if (itensDoPedido.length > 0) {
      // Lê o valor da coluna N (índice da coluna é 14) para cada linha de item
      const rangeResultadosItens = planilhaCalculo.getRange(2, 14, itensDoPedido.length, 1).getValues();
      rangeResultadosItens.forEach(linha => {
        icmsPorItemCalculado.push(linha[0] || 0);
      });
      Logger.log(`ICMS por item lido (da Coluna N): ${JSON.stringify(icmsPorItemCalculado)}`);
    }
    
    // --- 4. RETORNO DOS DADOS PARA O FRONTEND ---
    return {
      status: 'ok',
      resultados: {
        icmsStTotal: icmsStTotalCalculado,
        diferencaIcmsSn: diferencaIcmsSnCalculada,
        icmsPorItem: icmsPorItemCalculado // Retorna o array com o imposto de cada item
      }
    };

  } catch (e) {
    Logger.log(e.stack); // Log completo do erro para depuração
    return { status: 'error', message: `Erro no backend: ${e.message}` };
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
function testarCalculoDeImpostoCompleto() {
  Logger.log("--- INICIANDO TESTE da função 'calcularImpostoPedidoCompleto' ---");

  // 1. Simula a LISTA DE ITENS que o frontend enviaria.
  //    A propriedade importante aqui é 'subtotal'.
  const mockItens = [
    { descricao: 'PRODUTO DE TESTE A', subtotal: 1550.75 },
    { descricao: 'PRODUTO DE TESTE B', subtotal: 3116.19 },
    { descricao: 'PRODUTO DE TESTE C', subtotal: 850.00 }
  ];
  
  // 2. Simula os DADOS GERAIS do pedido que o frontend enviaria.
  const mockDadosGerais = {
    //numeroNFE: 'TESTE-CALCULO-789',
    ufFornecedor: 'SP', // Mude aqui para testar outros estados (ex: 'MG', 'BA')
    regimeTributario: 2   // Mude aqui: 1 para Simples Nacional, 2 para Outro
  };

  Logger.log("Enviando Itens: %s", JSON.stringify(mockItens, null, 2));
  Logger.log("Enviando Dados Gerais: %s", JSON.stringify(mockDadosGerais, null, 2));

  // 3. Chama a sua função real com os dados de teste
  const resultado = calcularImpostoPedidoCompleto(mockItens, mockDadosGerais);

  // 4. Mostra o resultado final no log de forma organizada
  Logger.log("--- RESULTADO RECEBIDO DO BACKEND ---");
  Logger.log(JSON.stringify(resultado, null, 2));
  
  if(resultado.status === 'ok') {
    Logger.log("✅ TESTE BEM-SUCEDIDO!");
    Logger.log("Total ICMS ST Calculado: %s", resultado.resultados.icmsStTotal);
    Logger.log("ICMS por Item: %s", resultado.resultados.icmsPorItem.join(', '));
  } else {
    Logger.log("❌ TESTE FALHOU! Mensagem: %s", resultado.message);
  }
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
