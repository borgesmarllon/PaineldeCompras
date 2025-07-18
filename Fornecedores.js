
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
function adicionarOuAtualizarFornecedor(fornecedorObject) {
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
        const indexCodigo = findHeaderIndex('CÓDIGO'); // Ajuste se o nome for 'ID' ou outro
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
            const indexDoCodigoAtual = codigosDaPlanilha.indexOf(String(fornecedorObject.codigo));

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
            newRowData[indexStatus] = 'ATIVO'; // Ou 'INATIVO' se preferir um fluxo de aprovação
            
            sheet.appendRow(newRowData);
            return { status: 'ok', message: 'Fornecedor adicionado com sucesso!' };
        }
    } catch (e) {
        Logger.log("ERRO em 'adicionarOuAtualizarFornecedor': " + e.message);
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
function alternarStatusFornecedor(codigoFornecedor) {
  // Bloco try...catch para garantir que um objeto sempre seja retornado.
  try {
    if (!codigoFornecedor) {
      throw new Error('Código do fornecedor não foi fornecido.');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
    if (!sheet) {
      throw new Error('Planilha "Fornecedores" não encontrada.');
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase().trim());
    
    // Usando 'ID' como você confirmou
    const indexCodigo = headers.indexOf('ID');
    const indexStatus = headers.indexOf('STATUS');

    if (indexCodigo === -1 || indexStatus === -1) {
      throw new Error('Coluna "ID" ou "STATUS" não encontrada.');
    }

    const codigos = sheet.getRange(2, indexCodigo + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIndexToUpdate = codigos.findIndex(codigo => String(codigo) == String(codigoFornecedor)) + 2;

    if (rowIndexToUpdate > 1) {
      const statusCell = sheet.getRange(rowIndexToUpdate, indexStatus + 1);
      const statusAtual = statusCell.getValue().toString().trim().toUpperCase();
      
      const novoStatus = (statusAtual === 'ATIVO') ? 'INATIVO' : 'ATIVO';
      statusCell.setValue(novoStatus);
      
      // Retorno de sucesso
      return { 
        status: 'ok', 
        message: `Fornecedor definido como '${novoStatus}'.`,
        novoStatus: novoStatus 
      };
    } else {
      // Retorno de erro controlado
      return { status: 'error', message: 'Fornecedor não encontrado.' };
    }
  } catch (e) {
    Logger.log("ERRO em alternarStatusFornecedor: " + e.message);
    // Retorno de erro de exceção
    return { status: 'error', message: 'Erro no servidor: ' + e.message };
  }
}

    /**
     * Consulta um CNPJ em uma API externa e retorna os dados da empresa.
     * @param {string} cnpj - O CNPJ a ser consultado.
     * @returns {object} Um objeto com o status da operação e os dados da empresa.
     */
    function consultarCnpj(cnpj) {
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
          // 4. Se a resposta for bem-sucedida, analisa os dados
          const dadosApi = JSON.parse(responseText);
          
          const partesEndereco = [
                dadosApi.logradouro,
                dadosApi.numero,
                dadosApi.bairro,
                dadosApi.complemento
            ].filter(Boolean); // O .filter(Boolean) remove itens nulos, vazios ou undefined da lista

            // Junta as partes existentes com uma formatação limpa
            const enderecoFormatado = partesEndereco.join(', ');

          // 5. Retorna um objeto limpo e padronizado
          return {
            status: 'ok',
            data: {
              razaoSocial: dadosApi.razao_social,
              nomeFantasia: dadosApi.nome_fantasia,
              endereco: enderecoFormatado,
              uf: dadosApi.uf,
              cidade: dadosApi.municipio
              // Adicione outros campos que desejar
            }
          };
        } else {
          // Se a API retornar um erro (ex: CNPJ não encontrado)
          const erroApi = JSON.parse(responseText);
          return { status: 'error', message: erroApi.message || 'CNPJ não encontrado ou inválido.' };
        }
      } catch (e) {
        Logger.log("Erro em consultarCnpj: " + e.message);
        return { status: 'error', message: 'Erro ao consultar o CNPJ. Verifique o console do servidor.' };
      }
    }


    function getEstados() {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
      if (!sheet) {
        Logger.log('ERRO: Planilha "Config" não encontrada! Verifique o nome da aba.');
        return [];
      }
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return [];
      }
      
      // Busca os dados das colunas D (UF) e E (Nome do Estado)
      const dados = sheet.getRange(3, 4, lastRow - 1, 2).getValues();
      
      // Filtra e formata os dados para o padrão esperado pelo frontend
      return dados
        .filter(([uf, nome]) => uf && nome) // Remove linhas vazias
        .map(([uf, nome]) => ({
          value: String(uf).trim(),
          text: String(nome).trim()
        }));
    }

    /**
     * Retorna uma lista de fornecedores (razão social) para preencher o dropdown de pedidos.
     * @returns {Array<Object>} Uma lista de objetos { codigo: string, razao: string, cnpj: string, endereco: string, condicao: string, forma: string }.
     */
    function getFornecedoresList() {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Fornecedores');

      if (!sheet || sheet.getLastRow() < 2) return [];
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

      // Encontra o índice da coluna "Status"
      const indexStatus = headers.findIndex(h => h.toUpperCase() === 'STATUS');

      const fornecedores = data
        // FILTRA para incluir apenas os que têm status "Ativo"
        .filter(row => indexStatus === -1 || String(row[indexStatus]).trim().toUpperCase() === 'ATIVO')
        .map(row => {

      const [codigo, razao, fantasia, cnpj, endereco, condicao, forma, idEmpresa, grupo, status, estado, cidade] = row;
      return {
        codigo: String(codigo),
        razao: String(razao),
        fantasia: String(fantasia),
        cnpj: String(cnpj),
        endereco: String(endereco),
        condicao: String(condicao),
        forma: String(forma),
        grupo: String(grupo || '').trim().toUpperCase(),
        estado: String(estado || ''),
        cidade: String(cidade || '')
        };
      });

      return fornecedores;
    } 

    /**
 * Retorna uma lista otimizada de fornecedores para a tela de Gerenciamento,
 * contendo apenas os campos necessários (código, nomes, cnpj, grupo, status).
 * @returns {Array<Object>} Uma lista de objetos de fornecedor.
 */
function getFornecedoresParaGerenciamento() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Fornecedores');
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toUpperCase().trim());
    
    const indexCodigo = headers.indexOf('ID');
    const indexRazao = headers.indexOf('RAZAO SOCIAL');
    const indexFantasia = headers.indexOf('NOME FANTASIA');
    const indexCnpj = headers.indexOf('CNPJ');
    const indexGrupo = headers.indexOf('GRUPO');
    const indexStatus = headers.indexOf('STATUS');

    if (indexCodigo === -1 || indexStatus === -1) {
      throw new Error("Coluna 'CÓDIGO' ou 'STATUS' não encontrada para a tela de gerenciamento.");
    }

    const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

    // Nota: Removi o filtro de 'ATIVO' daqui para que a tela de gerenciamento
    // possa mostrar TODOS os fornecedores e permitir ativar/desativar.
    const fornecedores = allData.map(row => {
      return {
        codigo: row[indexCodigo],
        razaoSocial: row[indexRazao],
        nomeFantasia: row[indexFantasia],
        cnpj: row[indexCnpj],
        grupo: row[indexGrupo],
        status: row[indexStatus]
      };
    });

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
