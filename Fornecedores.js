
    // ===============================================
    // FUNÇÕES PARA FORNECEDORES
    // ===============================================

    function getProximoCodigoFornecedor() {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Fornecedores');
      if (sheet.getLastRow() < 2) return '0001';

      const codigos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      const numeros = codigos.flat()
        .map(c => parseInt(c))
        .filter(n => !isNaN(n));

      const proximo = numeros.length ? Math.max(...numeros) + 1 : 1;
      return proximo.toString().padStart(4, '0');
    }

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

    /** function salvarFornecedor(fornecedor) {
      try {
        // 1. Limpa o CNPJ, deixando apenas os números
        const cnpjLimpo = String(fornecedor.cnpj).replace(/\D/g, '');
        if (cnpjLimpo.length !== 14) {
          throw new Error("O CNPJ deve conter 14 dígitos.");
        }
      const sheet = SpreadsheetApp.getActive().getSheetByName('Fornecedores');
      const data = sheet.getDataRange().getValues(); // Todas as linhas e colunas


      // Descubra em qual coluna está o CNPJ (assumindo headers na primeira linha)
      const header = data[0];
      const cnpjColIndex = header.findIndex(h => h.toString().toUpperCase().includes('CNPJ'));
      if (cnpjColIndex === -1) {
        return { status: 'error', message: 'Coluna CNPJ não encontrada.' };
      }

      // Verifica se o CNPJ já existe na planilha
        const cnpjJaExiste = data.slice(1).some(row => {
          // Limpa o CNPJ de cada linha da planilha para comparar apenas os números
          const cnpjDaLinha = String(row[cnpjColIndex] || '').replace(/\D/g, '');
          return cnpjDaLinha === cnpjLimpo;
        });

        if (cnpjJaExiste) {
          return { status: 'error', message: 'Já existe um fornecedor cadastrado com este CNPJ!' };
        }

      const codigoFornecedorTexto = "'" + fornecedor.codigo;

      sheet.appendRow([
        codigoFornecedorTexto, // Usa o código que veio do formulário
        fornecedor.razao,
        fornecedor.fantasia,
        fornecedor.cnpj,
        fornecedor.endereco,
        fornecedor.condicao,
        fornecedor.forma,
        "",
        "",
        'INATIVO',
        fornecedor.estado,
        fornecedor.cidade
      ]);
      return { status: 'INATIVO', message: "Fornecedor salvo com sucesso, solicite a ativação ao administrador!" };
    } catch (e) {
        Logger.log(`Erro em salvarFornecedor: ${e.message}`);
        return { status: 'error', message: `Erro no servidor: ${e.message}` };
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
        // Começa em i = 1 para pular o cabeçalho
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const fornecedor = {};
          headers.forEach((header, index) => {
            fornecedor[toCamelCase(header)] = row[index];
          });
          fornecedores.push(fornecedor);
        }
        Logger.log("DADOS FINAIS ENVIADOS PARA O FRONTEND: %s", JSON.stringify(fornecedores, null, 2));
        return fornecedores;

      } catch (e) {
        Logger.log("Erro em getFornecedoresParaGerenciar: " + e.message);
        return [];
      }
    }

    /**
     * Cria um novo fornecedor ou atualiza um existente de forma segura.
     * VERSÃO CORRIGIDA que não apaga o código ao editar.
     */
    /**
 * Adiciona um novo fornecedor ou atualiza um existente com base na presença do 'codigo'.
 * Esta função é a única porta de entrada para salvar dados do fornecedor.
 * @param {object} fornecedorObject O objeto de dados enviado pelo frontend.
 * @returns {object} Um objeto com o status e a mensagem da operação.
 */
function adicionarOuAtualizarFornecedorv2(fornecedorObject) {
    Logger.log("Iniciando 'adicionarOuAtualizarFornecedor' com dados: " + JSON.stringify(fornecedorObject));
    
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
        if (!sheet) {
            throw new Error("A planilha 'Fornecedores' não foi encontrada.");
        }
        
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .map(h => String(h).toUpperCase().trim());

        // Função auxiliar para encontrar o índice de uma coluna (mais robusta)
        const findHeaderIndex = (name) => {
            const index = headers.indexOf(name);
            if (index === -1) throw new Error(`Coluna essencial '${name}' não encontrada. Verifique o cabeçalho da planilha.`);
            return index;
        };

        // Mapeamento dos índices das colunas pelos nomes exatos nos cabeçalhos
        const indexCodigo = findHeaderIndex('ID'); // Ajuste se o nome for 'ID' ou outro
        const indexRazao = findHeaderIndex('RAZAO SOCIAL');
        const indexFantasia = findHeaderIndex('NOME FANTASIA');
        const indexCnpj = findHeaderIndex('CNPJ');
        const indexEndereco = findHeaderIndex('ENDERECO');
        const indexCidade = findHeaderIndex('CIDADE'); // Adicionado
        const indexEstado = findHeaderIndex('ESTADO'); // Adicionado
        const indexCondicao = findHeaderIndex('CONDICAO DE PAGAMENTO');
        const indexForma = findHeaderIndex('FORMA DE PAGAMENTO');
        const indexGrupo = findHeaderIndex('GRUPO');
        const indexStatus = findHeaderIndex('STATUS');
        const indexRegime = findHeaderIndex('REGIME TRIBUTARIO');

        const cnpjLimpo = String(fornecedorObject.cnpj).replace(/\D/g, '');
        if (cnpjLimpo.length !== 14) {
            throw new Error("O CNPJ deve conter 14 dígitos.");
        }

        const allData = sheet.getDataRange().getValues();
        const cnpjsDaPlanilha = allData.slice(1).map(row => String(row[indexCnpj] || '').replace(/\D/g, ''));

        // --- LÓGICA DE ATUALIZAÇÃO ---
        if (fornecedorObject.codigo) {
            Logger.log(`Modo ATUALIZAÇÃO para o código: ${fornecedorObject.codigo}`);
            
            // Verifica se o CNPJ já existe EM OUTRO fornecedor
            const codigosDaPlanilha = allData.slice(1).map(row => String(row[indexCodigo]));
            const codigoNumericoParaBuscar = parseInt(fornecedorObject.codigo, 10);
            const indexDoCodigoAtual = codigosDaPlanilha
                .map(c => parseInt(c, 10))
                .indexOf(codigoNumericoParaBuscar);

            cnpjsDaPlanilha.forEach((cnpj, index) => {
                if (cnpj === cnpjLimpo && index !== indexDoCodigoAtual) {
                    throw new Error("Este CNPJ já está cadastrado para outro fornecedor.");
                }
            });

            const rowIndexToUpdate = indexDoCodigoAtual + 2; // +1 porque slice(1) e +1 porque planilhas começam em 1
            if (rowIndexToUpdate > 1) {
                sheet.getRange(rowIndexToUpdate, indexRazao + 1).setValue(fornecedorObject.razaoSocial);
                sheet.getRange(rowIndexToUpdate, indexFantasia + 1).setValue(fornecedorObject.nomeFantasia);
                sheet.getRange(rowIndexToUpdate, indexCnpj + 1).setValue(fornecedorObject.cnpj);
                sheet.getRange(rowIndexToUpdate, indexEndereco + 1).setValue(fornecedorObject.endereco);
                sheet.getRange(rowIndexToUpdate, indexCidade + 1).setValue(fornecedorObject.cidade);
                sheet.getRange(rowIndexToUpdate, indexEstado + 1).setValue(fornecedorObject.estado);
                sheet.getRange(rowIndexToUpdate, indexCondicao + 1).setValue(fornecedorObject.condicaoPagamento);
                sheet.getRange(rowIndexToUpdate, indexForma + 1).setValue(fornecedorObject.formaPagamento);
                sheet.getRange(rowIndexToUpdate, indexGrupo + 1).setValue(fornecedorObject.grupo);
                sheet.getRange(rowIndexToUpdate, indexRegime + 1).setValue(fornecedorObject.regimeTributario);
                
                return { status: 'ok', message: 'Fornecedor atualizado com sucesso!' };
            } else {
                throw new Error('Fornecedor para atualização não encontrado com o código fornecido.');
            }

        // --- LÓGICA DE CRIAÇÃO ---
        } else {
            Logger.log("Modo CRIAÇÃO de novo fornecedor.");
            // Verifica se o CNPJ já existe
            if (cnpjsDaPlanilha.includes(cnpjLimpo)) {
                throw new Error("Já existe um fornecedor cadastrado com este CNPJ!");
            }
            
            const codigosNumericos = allData.slice(1)
                                            .map(row => parseInt(row[indexCodigo]))
                                            .filter(n => !isNaN(n));
            const novoCodigo = codigosNumericos.length ? Math.max(...codigosNumericos) + 1 : 1;
            
            const newRowData = [];
            newRowData[indexCodigo] = "'" + novoCodigo; // Formata como texto para evitar problemas
            newRowData[indexRazao] = fornecedorObject.razaoSocial;
            newRowData[indexFantasia] = fornecedorObject.nomeFantasia;
            newRowData[indexCnpj] = fornecedorObject.cnpj;
            newRowData[indexEndereco] = fornecedorObject.endereco;
            newRowData[indexCidade] = fornecedorObject.cidade;
            newRowData[indexEstado] = fornecedorObject.estado;
            newRowData[indexCondicao] = fornecedorObject.condicaoPagamento;
            newRowData[indexForma] = fornecedorObject.formaPagamento;
            newRowData[indexGrupo] = fornecedorObject.grupo || ''; // Garante que não seja undefined
            //newRowData[indexStatus] = 'INATIVO'; // Ou 'INATIVO' se preferir um fluxo de aprovação
            const estado = (fornecedorObject.estado || '').toUpperCase();
            const cidade = (fornecedorObject.cidade || '').toUpperCase();

            if (estado === 'BA' && cidade === 'VITORIA DA CONQUISTA') {
                newRowData[indexStatus] = 'INATIVO';
                newRowData[indexGrupo] = fornecedorObject.grupo || '';
            } else {
                newRowData[indexStatus] = 'ATIVO';
                newRowData[indexGrupo] = 'EXTERNO';
            }
            newRowData[indexRegime] = fornecedorObject.regimeTributario;

            sheet.appendRow(newRowData);
            return { status: 'ok', message: 'Fornecedor adicionado com sucesso!' };
        }
    } catch (e) {
        Logger.log("ERRO em suaFuncaoDeBackend: " + e.message);
        return { status: 'error', message: e.message }; // Retorna a mensagem de erro específica para o usuário
    }
}

    /**
     * Exclui um fornecedor da planilha.
     */
    function excluirFornecedor(codigoFornecedor) {
      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
        const codigos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
        const rowIndexToDelete = codigos.findIndex(codigo => String(codigo) == String(codigoFornecedor)) + 2;

        if (rowIndexToDelete > 1) {
          sheet.deleteRow(rowIndexToDelete);
          return { status: 'ok', message: 'Fornecedor excluído com sucesso!' };
        } else {
          return { status: 'error', message: 'Fornecedor não encontrado para exclusão.' };
        }
      } catch (e) {
        Logger.log("Erro em excluirFornecedor: " + e.message);
        return { status: 'error', message: 'Erro ao excluir o fornecedor.' };
      }
    }

    /**
 * Altera o status de um fornecedor entre 'ATIVO' e 'INATIVO'.
 * Esta versão é robusta e contém tratamento de erro para nunca retornar 'null'.
 * @param {string} codigoFornecedor - O código do fornecedor a ser alterado.
 * @returns {object} Um objeto com o status da operação e o novo status do fornecedor.
 */
function alternarStatusFornecedorv2(codigoFornecedor) {
  Logger.log("Iniciando alternância de status para código: " + codigoFornecedor);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
    if (!sheet) {
      Logger.log("Planilha 'Fornecedores' não encontrada.");
      return { status: 'error', message: 'Planilha não encontrada.' };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log("Cabeçalhos: " + JSON.stringify(headers));

    const indexCodigo = headers.indexOf("ID");
    const indexStatus = headers.indexOf("STATUS");
    Logger.log("Index ID: " + indexCodigo + " | Index STATUS: " + indexStatus);

    if (indexCodigo === -1 || indexStatus === -1) {
      Logger.log("Colunas 'ID' ou 'STATUS' não encontradas!");
      return { status: 'error', message: 'Colunas não encontradas.' };
    }

    const codigos = sheet.getRange(2, indexCodigo + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    Logger.log("Códigos lidos: " + JSON.stringify(codigos));

    // Normaliza ambos para string para evitar problemas de tipo
    const rowIndexToUpdate = codigos.findIndex(codigo => String(codigo) === String(codigoFornecedor)) + 2;
    Logger.log("Linha do código para alterar: " + rowIndexToUpdate);

    if (rowIndexToUpdate > 1) {
      const statusCell = sheet.getRange(rowIndexToUpdate, indexStatus + 1);
      const statusAtual = statusCell.getValue().toString().trim().toUpperCase();
      Logger.log("Status atual: " + statusAtual);

      const novoStatus = (statusAtual === "ATIVO") ? "INATIVO" : "ATIVO";
      statusCell.setValue(novoStatus);
      Logger.log("Novo status definido: " + novoStatus);

      return { status: 'ok', message: `Status alterado para '${novoStatus}'.`, novoStatus: novoStatus };
    } else {
      Logger.log("Código não encontrado na coluna ID.");
      return { status: 'error', message: 'Fornecedor não encontrado.' };
    }
  } catch (e) {
    Logger.log("Erro capturado: " + e.message);
    return { status: 'error', message: 'Erro inesperado: ' + e.message };
  }
}

    /**
     * Consulta um CNPJ em uma API externa e retorna os dados da empresa.
     * @param {string} cnpj - O CNPJ a ser consultado.
     * @returns {object} Um objeto com o status da operação e os dados da empresa.
     */
    function consultarCnpj_V2(cnpj) {
      try {
        // 1. Limpa o CNPJ, deixando apenas os números
        const cnpjLimpo = String(cnpj).replace(/\D/g, '');
        if (cnpjLimpo.length !== 14) {
          throw new Error("O CNPJ deve conter 14 dígitos.");
        }
        
        // 2. Monta a URL da API para a consulta
        const apiUrl = `https://brasilapi.com.br/api/cnpj/v1/${cnpjLimpo}`;
        
        // 3. Faz a chamada para a API externa
        const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        
        if (responseCode === 200) {
        
      const dadosApi = JSON.parse(responseText);
      let regimeCodigo = 2; // Padrão para 'Outro' (Lucro Presumido/Real)
      
      // Se for optante pelo MEI ou pelo Simples, o código é 1
      if (dadosApi.opcao_pelo_mei || dadosApi.opcao_pelo_simples) {
        regimeCodigo = 1;
      }
      const partesEndereco = [
          dadosApi.logradouro, dadosApi.numero, dadosApi.bairro, dadosApi.complemento
      ].filter(Boolean);
      const enderecoFormatado = partesEndereco.join(', ');

      return {
        status: 'ok',
        data: {
          razaoSocial: dadosApi.razao_social,
          nomeFantasia: dadosApi.nome_fantasia,
          endereco: enderecoFormatado,
          uf: dadosApi.uf,
          cidade: dadosApi.municipio,
          regimeTributario: regimeCodigo
        }
      };
    } else {
      const erroApi = JSON.parse(responseText);
      return { status: 'error', message: erroApi.message || 'CNPJ não encontrado na API.' };
    }
  } catch (e) {
    Logger.log("ERRO em consultarCnpj: " + e.message);
    return { status: 'error', message: 'Erro interno do servidor: ' + e.message };
  }
}

/**
 * Função criada apenas para testar a 'consultarCnpj' diretamente no editor do Apps Script.
 
function testarConsultaCnpj() {
  // Vamos usar um CNPJ válido e conhecido para o teste.
  const cnpjDeTeste = "06.990.590/0001-23"; // CNPJ da Google Brasil
  
  console.log(`Iniciando teste com o CNPJ: ${cnpjDeTeste}`);
  
  // Chama a sua função real que queremos testar
  const resultado = consultarCnpj(cnpjDeTeste);
  
  // Imprime o resultado formatado no log para fácil visualização
  console.log("Resultado da consulta:");
  console.log(JSON.stringify(resultado, null, 2));
}
*/
function getEstadosv2() {
  Logger.log("--- Iniciando getEstados ---");
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
    if (!sheet) {
      Logger.log('ERRO: Planilha "Config" não encontrada!');
      return [];
    }
    Logger.log("✅ Planilha 'Config' encontrada com sucesso.");

    const lastRow = sheet.getLastRow();
    if (lastRow < 3) {
      Logger.log("AVISO: Não há dados de estados a partir da linha 3.");
      return [];
    }
    
    // Busca os dados das colunas D (4), E (5) e F (6)
    const dados = sheet.getRange(3, 4, lastRow - 2, 3).getValues(); 
    Logger.log("1. Dados brutos lidos da planilha (UF, Nome, Alíquota):");
    Logger.log(JSON.stringify(dados, null, 2));
    
    const estadosMapeados = dados
      .filter(([uf, nome, aliquota]) => uf && nome) // Remove linhas onde UF ou Nome são vazios
      .map(([uf, nome, aliquota]) => ({
        value: String(uf).trim(),
        text: String(nome).trim(),
        aliquota: parseFloat(String(aliquota || '0').replace(',', '.')) || 0 // Converte para número de forma segura
      }));
    
    Logger.log("2. Dados finais mapeados e formatados:");
    Logger.log(JSON.stringify(estadosMapeados, null, 2));
    Logger.log("--- Finalizando getEstados ---");

    return estadosMapeados;

  } catch (e) {
    Logger.log('!!! ERRO em getEstados: ' + e.message);
    return [];
  }
}

function testarGetEstados() {
  Logger.log("--- INICIANDO TESTE da função 'getEstados' ---");
  
  const resultado = getEstadosv2();
  
  Logger.log("--- RESULTADO FINAL RETORNADO PELA FUNÇÃO ---");
  Logger.log(JSON.stringify(resultado, null, 2));
}

    /**
     * Retorna uma lista de fornecedores (razão social) para preencher o dropdown de pedidos.
     * @returns {Array<Object>} Uma lista de objetos { codigo: string, razao: string, cnpj: string, endereco: string, condicao: string, forma: string }.
     */
    function getFornecedoresListv2() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Fornecedores');
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    // Lê os cabeçalhos da linha 1
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).toUpperCase().trim());
    
    // Encontra a posição (índice) de cada coluna pelo nome
    const indexStatus = headers.indexOf('STATUS');
    const indexCodigo = headers.indexOf('ID'); // Assumindo que a coluna se chama ID
    const indexRazao = headers.indexOf('RAZAO SOCIAL');
    const indexFantasia = headers.indexOf('NOME FANTASIA');
    const indexCnpj = headers.indexOf('CNPJ');
    const indexGrupo = headers.indexOf('GRUPO');
    const indexEstado = headers.indexOf('ESTADO');
    const indexCidade = headers.indexOf('CIDADE');
    const indexRegime = headers.indexOf('REGIME TRIBUTARIO');
    const indexCondicao = headers.indexOf('CONDICAO DE PAGAMENTO');
    const indexForma = headers.indexOf('FORMA DE PAGAMENTO');

    // Validação para garantir que a coluna essencial 'Status' existe
    if (indexStatus === -1) {
      throw new Error("Coluna 'STATUS' não encontrada na planilha 'Fornecedores'.");
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

    const fornecedores = data
      // Filtra para incluir apenas os que têm status "ATIVO"
      .filter(row => String(row[indexStatus]).trim().toUpperCase() === 'ATIVO')
      // Mapeia para um objeto, pegando os dados da coluna correta pelo índice
      .map(row => ({
        id: row[indexCodigo],
        razaoSocial: row[indexRazao], // Usando 'razaoSocial' para manter o padrão
        nomeFantasia: row[indexFantasia], // Usando 'nomeFantasia' para manter o padrão
        cnpj: row[indexCnpj],
        grupo: String(row[indexGrupo] || '').trim().toUpperCase(),
        estado: String(row[indexEstado] || ''),
        cidade: String(row[indexCidade] || ''),
        regime: row[indexRegime] || '2',
        condicao: row[indexCondicao],
        forma: row[indexForma]
      }));
    Logger.log("DADOS FINAIS SENDO ENVIADOS PARA O FRONTEND (amostra de até 5):");
    // Mostra uma amostra dos 5 primeiros para não poluir o log
    Logger.log(JSON.stringify(fornecedores.slice(0, 5), null, 2)); 
    return fornecedores;
    
  } catch (e) {
    Logger.log("ERRO em getFornecedoresListv2: " + e.message);
    return []; // Retorna vazio em caso de erro
  }
}

    /**
 * Retorna uma lista otimizada de fornecedores para a tela de Gerenciamento,
 * contendo apenas os campos necessários (código, nomes, cnpj, grupo, status).
 * @returns {Array<Object>} Uma lista de objetos de fornecedor.
 */
function getFornecedoresParaGerenciamento() {
  Logger.log("--- Iniciando getFornecedoresParaGerenciamento ---");
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Fornecedores');
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log("AVISO: Aba 'Fornecedores' não encontrada ou vazia. Retornando array vazio.");

      return [];
    }
    Logger.log("Aba 'Fornecedores' encontrada com sucesso.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase().trim());
    Logger.log("Cabeçalhos encontrados e padronizados: " + JSON.stringify(headers));
    const indexCodigo = headers.indexOf('ID');
    const indexRazao = headers.indexOf('RAZAO SOCIAL');
    const indexFantasia = headers.indexOf('NOME FANTASIA');
    const indexCnpj = headers.indexOf('CNPJ');
    const indexGrupo = headers.indexOf('GRUPO');
    const indexStatus = headers.indexOf('STATUS');
    const indexRegime = headers.indexOf('REGIME TRIBUTARIO');

    if (indexCodigo === -1 || indexStatus === -1) {
      throw new Error("Coluna 'CÓDIGO' ou 'STATUS' não encontrada para a tela de gerenciamento.");
      Logger.log(`ERRO: ${erroMsg}`);
    }

    const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    Logger.log(`${allData.length} linhas de dados encontradas para processar.`);

    // Nota: Removi o filtro de 'ATIVO' daqui para que a tela de gerenciamento
    // possa mostrar TODOS os fornecedores e permitir ativar/desativar.
    const fornecedores = allData.map(row => {
      return {
        codigo: row[indexCodigo],
        razaoSocial: row[indexRazao],
        nomeFantasia: row[indexFantasia],
        cnpj: row[indexCnpj],
        grupo: row[indexGrupo],
        status: row[indexStatus],
        regimeTributario: row[indexRegime] || 'Outro'
      };
    });

    Logger.log(`Índice da coluna 'ID': ${indexCodigo}`);
    Logger.log(`Índice da coluna 'RAZAO SOCIAL': ${indexRazao}`);
    Logger.log(`Índice da coluna 'STATUS': ${indexStatus}`);

    Logger.log(`Processamento concluído. Retornando ${fornecedores.length} fornecedores.`);
    return fornecedores;

  } catch (e) {
    Logger.log("ERRO em getFornecedoresParaGerenciamento: " + e.message);
    return []; 
  }
}

/**
 * Busca um fornecedor específico pelo seu código e retorna todos os seus dados.
 * @param {string} codigo - O código/ID do fornecedor a ser buscado.
 * @returns {object | null} Um objeto com os dados do fornecedor ou null se não for encontrado.
 */
function obterFornecedorPorCodigo(codigo) {
  try {
    // Validação inicial
    if (!codigo) {
      throw new Error("O código do fornecedor não foi fornecido para a busca.");
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
    if (!sheet) {
      throw new Error("Planilha 'Fornecedores' não encontrada.");
    }

    // Pega todos os cabeçalhos e dados da planilha
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    
    // Encontra o índice da coluna de código (usando 'ID' como confirmado)
    const indexCodigo = headers.map(h => String(h).toUpperCase().trim()).indexOf('ID');
    if (indexCodigo === -1) {
      throw new Error("A coluna de cabeçalho 'ID' não foi encontrada na planilha.");
    }

    // Procura pela linha que corresponde ao código fornecido
    const rowEncontrada = data.find(row => String(row[indexCodigo]) == String(codigo));

    // Se encontrou a linha, monta o objeto de resposta
    if (rowEncontrada) {
      const fornecedor = {};
      
      // Mapeia dinamicamente os cabeçalhos para as propriedades do objeto
      // Isso torna a função flexível a novas colunas no futuro
      headers.forEach((header, index) => {
        // Converte nomes como 'RAZAO SOCIAL' para 'razaoSocial' (camelCase)
        const key = String(header).trim().toLowerCase()
                                .replace(/ \(\w+\)/g, '') // remove (ex)
                                .replace(/[^a-zA-Z0-9]+(.)/g, (m, chr) => chr.toUpperCase());
        
        fornecedor[key] = rowEncontrada[index];
      });
      
      Logger.log("Fornecedor encontrado para edição: " + JSON.stringify(fornecedor));
      return fornecedor; // Retorna o objeto completo com os dados do fornecedor

    } else {
      // Se não encontrou o fornecedor, retorna null
      Logger.log("Nenhum fornecedor encontrado com o código: " + codigo);
      return null;
    }

  } catch (e) {
    Logger.log("Erro em obterFornecedorPorCodigo: " + e.message);
    return null; // Retorna null também em caso de qualquer erro
  }
}

function testeForcado_AlternarStatus(codigo) {
  Logger.log("--- INICIANDO TESTE FORÇADO para o código: " + codigo + " ---");
  
  let resultado;
  try {
    // Chama a sua função real que queremos inspecionar
    resultado = alternarStatusFornecedor(codigo);
    
    // Loga o que quer que a função tenha retornado
    Logger.log("A função 'alternarStatusFornecedor' retornou um valor.");
    Logger.log("Tipo do resultado: " + typeof resultado);
    Logger.log("Conteúdo do resultado: " + JSON.stringify(resultado, null, 2));

  } catch (e) {
    Logger.log("UM ERRO OCORREU AO TENTAR CHAMAR 'alternarStatusFornecedor'. Erro: " + e.message);
    resultado = { status: 'error', message: "Falha catastrófica ao chamar a função: " + e.message };
  }

  // Verificação final para a causa do 'null' no frontend
  if (resultado === undefined) {
      Logger.log("ALERTA: O resultado da função foi 'undefined'. Esta é a causa do 'null' no frontend.");
      // Se for undefined, nós forçamos um objeto de erro para o frontend não quebrar.
      return { status: 'error', message: "Resultado foi 'undefined' no backend." };
  }
  
  Logger.log("--- TESTE FORÇADO CONCLUÍDO ---");
  return resultado; // Retorna o resultado para o frontend
}

function runTesteAlternarStatusFornecedor() {
  var resultado = alternarStatusFornecedorv2(3); // Troque para o código desejado
  Logger.log("Resultado do teste: " + JSON.stringify(resultado, null, 2));
  return resultado; // <-- Isso só tem efeito se for chamado por outra função
}

/**
 * Lê os cabeçalhos e os dados da planilha "Fornecedores" exatamente como estão,
 * sem normalização ou conversão.
 * @returns {object} Um objeto com os campos 'headers' e 'data'.
 */
function lerCabecalhosEDadosFornecedores() {
  Logger.log("Iniciando leitura da planilha 'Fornecedores'...");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
  if (!sheet) {
    Logger.log("ERRO: Planilha 'Fornecedores' não encontrada.");
    return { headers: [], data: [] };
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  Logger.log("lastRow: " + lastRow + ", lastCol: " + lastCol);
  if (lastRow < 1 || lastCol < 1) {
    Logger.log("Planilha sem dados.");
    return { headers: [], data: [] };
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log("Cabeçalhos encontrados: " + JSON.stringify(headers));
  var data = [];
  if (lastRow > 1) {
    data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    Logger.log("Total de linhas de dados: " + data.length);
  } else {
    Logger.log("Não há dados além dos cabeçalhos.");
  }
  return {
    headers: headers,
    data: data
  };
}
function lerCabecalhoEDadosOriginaisFornecedores() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
  if (!sheet) {
    Logger.log("Planilha 'Fornecedores' não encontrada.");
    return { headers: [], rows: [] };
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    Logger.log("Planilha sem dados.");
    return { headers: [], rows: [] };
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]; // Cabeçalhos como estão na planilha
  var rows = [];
  if (lastRow > 1) {
    rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues(); // Dados das linhas como estão
  }
  return {
    headers: headers,
    rows: rows
  };
}

function lerCabecalhoEDadosComLogFornecedores() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
  if (!sheet) {
    Logger.log("Planilha 'Fornecedores' não encontrada.");
    return { headers: [], rows: [] };
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  Logger.log("lastRow: " + lastRow + ", lastCol: " + lastCol);

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log("Cabeçalhos: " + JSON.stringify(headers));

  var rows = [];
  if (lastRow > 1) {
    rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    Logger.log("Total de linhas lidas: " + rows.length);
    for (var i = 0; i < rows.length; i++) {
      Logger.log("Linha " + (i+2) + ": " + JSON.stringify(rows[i]));
    }
  } else {
    Logger.log("Não há dados além do cabeçalho.");
  }
  return {
    headers: headers,
    rows: rows
  };
}

function testarGetFornecedoresParaGerenciamento() {
  Logger.log("--- INICIANDO TESTE da função 'getFornecedoresParaGerenciamento' ---");
  
  const resultado = getFornecedoresParaGerenciamento();
  
  if (resultado && resultado.length > 0) {
    Logger.log(`✅ SUCESSO! A função retornou ${resultado.length} fornecedores.`);
    Logger.log("Amostra dos primeiros 2 fornecedores: " + JSON.stringify(resultado.slice(0, 2), null, 2));
  } else {
    Logger.log("⚠️ ATENÇÃO: A função retornou um array vazio.");
    Logger.log("Verifique os logs acima para entender o motivo. As causas mais comuns são:");
    Logger.log("1. A aba 'Fornecedores' está vazia ou não foi encontrada.");
    Logger.log("2. Os nomes dos cabeçalhos 'ID' ou 'STATUS' na planilha não correspondem exatamente ao esperado.");
  }
}
