
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

    function salvarFornecedor(fornecedor) {
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
        fornecedor.estado
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
    function adicionarOuAtualizarFornecedor(fornecedorObject) {
      Logger.log("--- [DIAGNÓSTICO SALVAR] ---");
      Logger.log("1. Objeto recebido do frontend: " + JSON.stringify(fornecedorObject));
      
      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
        if (!sheet) {
          throw new Error("A planilha 'Fornecedores' não foi encontrada.");
        }
        
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                            .map(h => String(h).toUpperCase().trim());
        Logger.log("2. Cabeçalhos encontrados na planilha: [" + headers.join(', ') + "]");

        // Função auxiliar para encontrar o índice de uma coluna ou retornar um erro claro
        const findHeaderIndex = (possibleNames) => {
          for (const name of possibleNames) {
            const index = headers.indexOf(name);
            if (index !== -1) {
              return index;
            }
            return -1;
          }
          throw new Error(`Nenhuma das colunas esperadas (${possibleNames.join(', ')}) foi encontrada. Verifique os cabeçalhos da planilha "Fornecedores".`);
        };

        // Adapte os nomes dos cabeçalhos abaixo para corresponderem à sua planilha
        const indexCodigo = headers.indexOf('ID');
        const indexRazao = headers.indexOf('RAZAO SOCIAL');
        const indexFantasia = headers.indexOf('NOME FANTASIA');
        const indexCnpj = headers.indexOf('CNPJ');
        const indexEndereco = headers.indexOf('ENDERECO');
        const indexCondicao = headers.indexOf('CONDICAO DE PAGAMENTO');
        const indexForma = headers.indexOf('FORMA DE PAGAMENTO');
        const indexGrupo = headers.indexOf('GRUPO');
        const indexStatus = headers.indexOf('STATUS');
        Logger.log(`3. Índices encontrados -> Código: ${indexCodigo}, Razão: ${indexRazao}, Condição: ${indexCondicao}, Forma: ${indexForma}`);

        // Validação para garantir que as colunas essenciais foram encontradas
        if ([indexCodigo, indexRazao, indexFantasia, indexCnpj, indexEndereco, indexCondicao, indexForma, indexGrupo, indexStatus].includes(-1)) {
            throw new Error("Uma ou mais colunas essenciais não foram encontradas. Verifique os nomes no log acima.");
        }
        // Verifica se é uma atualização (se um código foi enviado)
        if (fornecedorObject.codigo) {
          Logger.log("4. Modo ATUALIZAÇÃO detectado.");
          // --- LÓGICA DE ATUALIZAÇÃO SEGURA ---
          const codigos = sheet.getRange(2, indexCodigo + 1, sheet.getLastRow() - 1, 1).getValues().flat();
          const rowIndexToUpdate = codigos.findIndex(codigo => String(codigo) == String(fornecedorObject.codigo)) + 2;
          Logger.log(`5. Procurando pelo código "${fornecedorObject.codigo}". Linha encontrada: ${rowIndexToUpdate > 1 ? rowIndexToUpdate : 'NENHUMA'}`);

          if (rowIndexToUpdate > 1) {
            // Atualiza apenas as células necessárias, preservando o resto da linha
            sheet.getRange(rowIndexToUpdate, indexRazao + 1).setValue(fornecedorObject.razaoSocial);
            sheet.getRange(rowIndexToUpdate, indexFantasia + 1).setValue(fornecedorObject.nomeFantasia);
            sheet.getRange(rowIndexToUpdate, indexCnpj + 1).setValue(fornecedorObject.cnpj);
            sheet.getRange(rowIndexToUpdate, indexEndereco + 1).setValue(fornecedorObject.endereco);
            sheet.getRange(rowIndexToUpdate, indexCondicao + 1).setValue(fornecedorObject.condicaoPagamento);
            sheet.getRange(rowIndexToUpdate, indexForma + 1).setValue(fornecedorObject.formaPagamento);
            sheet.getRange(rowIndexToUpdate, indexGrupo + 1).setValue(fornecedorObject.grupo);
            
            return { status: 'ok', message: 'Fornecedor atualizado com sucesso!' };
          } else {
            return { status: 'error', message: 'Fornecedor para atualização não encontrado.' };
          }

        } else {
          // --- LÓGICA DE CRIAÇÃO ---
          const ids = sheet.getRange(2, indexCodigo + 1, sheet.getLastRow() - 1, 1).getValues().flat().map(id => parseInt(id)).filter(n => !isNaN(n));
          const novoId = ids.length ? Math.max(...ids) + 1 : 1;
          
          const newRowData = [];
          newRowData[indexCodigo] = "'" + novoId;
          newRowData[indexRazao] = fornecedorObject.razaoSocial;
          newRowData[indexFantasia] = fornecedorObject.nomeFantasia;
          newRowData[indexCnpj] = fornecedorObject.cnpj;
          newRowData[indexEndereco] = fornecedorObject.endereco;
          newRowData[indexCondicao] = fornecedorObject.condicaoPagamento;
          newRowData[indexForma] = fornecedorObject.formaPagamento;
          newRowData[indexGrupo] = fornecedorObject.grupo;
          if (indexStatus > -1) newRowData[indexStatus] = 'Ativo';

          sheet.appendRow(newRowData);
          return { status: 'ok', message: 'Fornecedor adicionado com sucesso!' };
        }
      } catch (e) {
        Logger.log("Erro em adicionarOuAtualizarFornecedor: " + e.message);
        return { status: 'error', message: 'Erro ao salvar o fornecedor.' };
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
     * Altera o status de um fornecedor para 'Inativo' na planilha.
     * @param {string} codigoFornecedor - O código do fornecedor a ser inativado.
     * @returns {object} Um objeto com o status da operação.
     */
    function alternarStatusFornecedor(codigoFornecedor) {
      Logger.log("alternarStatusFornecedor - codigoFornecedor recebido: " + codigoFornecedor)
      if (!codigoFornecedor) {
        return { status: 'error', message: 'Código do fornecedor não fornecido.' };
      }

      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
        if (!sheet) {
          throw new Error('Planilha "Fornecedores" não encontrada.');
        }

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        const indexCodigo = headers.findIndex(h => ['CODIGO', 'ID'].includes(h.toUpperCase()));
        const indexStatus = headers.findIndex(h => h.toUpperCase() === 'STATUS');

        if (indexCodigo === -1 || indexStatus === -1) {
          throw new Error('Coluna "Codigo" ou "Status" não encontrada na planilha "Fornecedores".');
        }

        const codigos = sheet.getRange(2, indexCodigo + 1, sheet.getLastRow() - 1, 1).getValues().flat();
        const rowIndexToUpdate = codigos.findIndex(codigo => parseInt(codigo) === parseInt(codigoFornecedor)) + 2;

        if (rowIndexToUpdate > 1) {
          const statusCell = sheet.getRange(rowIndexToUpdate, indexStatus + 1);
          const statusAtual = statusCell.getValue().toString().trim().toUpperCase();
          
          // Lógica do "interruptor"
          const novoStatus = (statusAtual === 'ATIVO') ? 'Inativo' : 'Ativo';
          
          statusCell.setValue(novoStatus);
          
          return { status: 'ok', message: `Fornecedor definido como '${novoStatus}' com sucesso!` };
        } else {
          return { status: 'error', message: 'Fornecedor não encontrado para alterar o status.' };
        }
      } catch (e) {
        Logger.log("Erro em alternarStatusFornecedor: " + e.message);
        return { status: 'error', message: 'Erro ao alterar o status do fornecedor.' };
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
          
          // 5. Retorna um objeto limpo e padronizado
          return {
            status: 'ok',
            data: {
              razaoSocial: dadosApi.razao_social,
              nomeFantasia: dadosApi.nome_fantasia,
              endereco: `${dadosApi.logradouro}, ${dadosApi.numero}. ${dadosApi.bairro}`,
              uf: dadosApi.uf
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