// ===============================================
    // FUNÇÕES PARA PEDIDOS DE COMPRA
    // ===============================================

    /**
     * Retorna o próximo número sequencial para um novo pedido.
     * Cria a planilha 'Pedidos' se não existir.
     * @returns {string} O próximo número de pedido formatado como '0001'.
     */
    function getProximoNumeroPedido(empresaCodigo) {
      const spreadsheet = SpreadsheetApp.getActive();
      let sheet = spreadsheet.getSheetByName('Pedidos');

      if (!sheet) {
        sheet = spreadsheet.insertSheet('Pedidos');
        const headers = [
          'Número do Pedido', 'ID da Empresa', 'Data', 'Fornecedor', 'CNPJ Fornecedor',
          'Endereço Fornecedor', 'Condição Pagamento Fornecedor', 'Forma Pagamento Fornecedor',
          'Placa Veiculo', 'Nome Veiculo', 'Observacoes', 'Total Geral', 'Status', 'Itens'
        ];
        sheet.appendRow(headers);
        sheet.getRange('A:B').setNumberFormat('@');
        return '000001';
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return '000001';
      }

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const colNumero = headers.findIndex(h => h.toUpperCase() === 'NÚMERO DO PEDIDO');
      const colEmpresa = headers.findIndex(h => ['ID DA EMPRESA', 'ID EMPRESA', 'EMPRESA'].includes(h.toUpperCase()));

      if (colEmpresa === -1 || colNumero === -1) {
        throw new Error('Cabeçalhos "ID da Empresa" ou "Número do Pedido" não encontrados na planilha "Pedidos".');
      }

      const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      const empresaCodigoTratado = String(empresaCodigo).trim();
      
      const numeros = data
        .filter(row => {
          const idNaLinha = String(row[colEmpresa]).trim();
          
          // --- AQUI ESTÁ A CORREÇÃO FINAL ---
          // Converte ambos os IDs para números antes de comparar.
          // parseInt("1") vira 1. parseInt("001") também vira 1. A comparação funciona.
          return parseInt(idNaLinha, 10) === parseInt(empresaCodigoTratado, 10);
        })
        .map(row => parseInt(row[colNumero], 10))
        .filter(n => !isNaN(n));

      const proximoNumero = numeros.length > 0 ? Math.max(...numeros) + 1 : 1;
      
      return proximoNumero.toString().padStart(6, '0');
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

      const [codigo, razao, fantasia, cnpj, endereco, condicao, forma, idEmpresa, grupo, status, estado] = row;
      return {
        codigo: String(codigo),
        razao: String(razao),
        fantasia: String(fantasia),
        cnpj: String(cnpj),
        endereco: String(endereco),
        condicao: String(condicao),
        forma: String(forma),
        grupo: String(grupo || '').trim().toUpperCase(),
        estado: String(estado || '')
        };
      });

      return fornecedores;
    } 

    /**
     * Salva um novo pedido de compra na planilha 'Pedidos'.
     * @param {Object} pedido - Objeto contendo os detalhes do pedido (numero, data, fornecedor, itens, totalGeral, placaVeiculo, nomeVeiculo, observacoes).
     * @returns {Object} Um objeto com status e mensagem.
     */
    function salvarPedido(pedido) {
      console.log('📋 === INÍCIO salvarPedido ===');
      console.log('📋 Objeto pedido recebido:', JSON.stringify(pedido, null, 2));
      

      
      const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
      if (!sheet) {
        return { status: 'error', message: 'Planilha "Pedidos" não encontrada. Contate o administrador.' };
      }

      // Normalizar o número do pedido - aceitar tanto 'numeroPedido' quanto 'numero'
      const numeroPedido = pedido.numeroPedido || pedido.numero;
      console.log('📋 Número do pedido normalizado:', numeroPedido);
      
      if (!numeroPedido) {
        console.error('❌ Número do pedido não encontrado em:', pedido);
        return { status: 'error', message: 'Número do pedido é obrigatório.' };
      }

      // Normalizar empresa (aceitar empresaId ou empresa)
      const empresaId = pedido.empresaId || pedido.empresa;
      console.log('📋 ID da empresa normalizado:', empresaId);
      
      console.log('📋 Total geral recebido:', pedido.totalGeral);
      console.log('📋 Itens recebidos:', pedido.itens ? pedido.itens.length : 0, 'itens');

      const itensJSON = JSON.stringify(pedido.itens);

      const fornecedoresSheet = SpreadsheetApp.getActive().getSheetByName('Fornecedores');
      let fornecedorCnpj = '';
      let fornecedorEndereco = '';
      let condicaoPagamentoFornecedor = '';
      let formaPagamentoFornecedor = '';
      let estadoFornecedor = '';

      if (fornecedoresSheet) {
        const fornecedoresData = fornecedoresSheet.getRange(2, 1, fornecedoresSheet.getLastRow() - 1, fornecedoresSheet.getLastColumn()).getValues();
        const foundFornecedor = fornecedoresData.find(row => String(row[1]) === pedido.fornecedor); 
        if (foundFornecedor) {
          fornecedorCnpj = String(foundFornecedor[3] || '');
          fornecedorEndereco = String(foundFornecedor[4] || '');
          condicaoPagamentoFornecedor = String(foundFornecedor[5] || '');
          formaPagamentoFornecedor = String(foundFornecedor[6] || '');
          estadoFornecedor = String(foundFornecedor[10] || ''); // Coluna 11 (índice 10) = Estado
        }
      }

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const rowData = new Array(headers.length).fill('');

      const dataToSave = {
        'Número do Pedido': "'" + numeroPedido, // Adiciona o apóstrofo
        'Empresa': "'" + empresaId,     // Usar o mesmo nome que está em getProximoNumeroPedido
        'Data': formatarDataParaISO(pedido.data),
        'Fornecedor': pedido.fornecedor,
        'CNPJ Fornecedor': fornecedorCnpj,
        'Endereço Fornecedor': fornecedorEndereco,
        'Estado Fornecedor': estadoFornecedor,
        'Condição Pagamento Fornecedor': condicaoPagamentoFornecedor,
        'Forma Pagamento Fornecedor': formaPagamentoFornecedor,
        'Placa Veiculo': pedido.placaVeiculo,
        'Nome Veiculo': pedido.nomeVeiculo,
        'Observacoes': pedido.observacoes,
        'Total Geral': parseFloat(pedido.totalGeral) || 0, // Garantir que é um número
        'Status': 'Em Aberto',
        'Itens': itensJSON,
        'Data Hora Criacao': formatarDataParaISO(new Date()) // Timestamp de criação padronizado
      };



      headers.forEach((header, index) => {
        if (dataToSave.hasOwnProperty(header)) {
          rowData[index] = dataToSave[header];
          console.log(`📋 Mapeando coluna "${header}":`, dataToSave[header]);
        }
      });

      console.log('📋 Dados finais para salvar:', rowData);
      sheet.appendRow(rowData);

      console.log('✅ Pedido salvo com sucesso:', numeroPedido);
      console.log('📋 === FIM salvarPedido ===');
      return { status: 'ok', message: `Pedido ${numeroPedido} salvo com sucesso!` };
    }
 function listarPedidosPorEmpresa(empresa) {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
      if (!sheet) {
        return [];
      }

      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const pedidos = [];

      const indexEmpresa = headers.indexOf('Empresa');
      if (indexEmpresa === -1) {
        throw new Error('Coluna "Empresa" não encontrada na planilha Pedidos.');
      }

      for (let i = 1; i < data.length; i++) {
        const linha = data[i];
        const empresaDaLinha = String(linha[indexEmpresa]).trim();

        if (empresaDaLinha === String(empresa).trim()) {
          const pedido = {};
          headers.forEach((header, idx) => {
            pedido[header] = linha[idx];
          });
          pedidos.push(pedido);
        }
      }

      return pedidos;
    }

    /**
     * Busca um único pedido pelo seu número e pelo ID da empresa para edição.
     * @param {string} numeroDoPedido - O número do pedido a ser encontrado.
     * @param {string} idEmpresa - O ID da empresa à qual o pedido pertence.
     * @returns {object|null} O objeto do pedido encontrado ou null se não encontrar.
     */
    function getPedidoParaEditar(numeroDoPedido, idEmpresa) {
      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
        if (!sheet) throw new Error("Planilha 'Pedidos' não encontrada.");

        const data = sheet.getDataRange().getValues();
        const headers = data[0].map(h => toCamelCase(h));

        // Encontra os índices das colunas
        const indexNumero = headers.indexOf('numeroDoPedido');
        const indexEmpresa = headers.indexOf('idDaEmpresa') > -1 ? headers.indexOf('idDaEmpresa') : headers.indexOf('empresa');

        if (indexNumero === -1 || indexEmpresa === -1) {
          throw new Error("Colunas 'Número do Pedido' ou 'ID da Empresa' não encontradas.");
        }

        // Procura pela linha que corresponde ao número do pedido E ao ID da empresa
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const numPedidoNaLinha = String(row[indexNumero]).trim();
          const idEmpresaNaLinha = String(row[indexEmpresa]).trim();

          if (numPedidoNaLinha == numeroDoPedido && idEmpresaNaLinha == idEmpresa) {
            // Encontrou o pedido, agora monta o objeto
            const pedido = {};
            headers.forEach((header, index) => {
              let value = row[index];
              // Converte a data para um formato de texto padronizado
              if (value instanceof Date) {
                value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
              }
              pedido[header] = value;
            });
            
            // Faz o parse dos itens
            if (typeof pedido.itens === 'string') {
              pedido.itens = JSON.parse(pedido.itens);
            }
            
            return pedido; // Retorna o objeto do pedido encontrado
          }
        }

        return null; // Retorna null se não encontrar o pedido

      } catch (e) {
        Logger.log("Erro em getPedidoParaEditar: " + e.message);
        return null;
      }
    }
// ===============================================
    // FUNÇÕES PARA VEICULOS, PLACAS E FORNECEDORES
    // ===============================================
    /**
     * Adiciona um novo nome de veículo à planilha "Veiculos".
     * @param {string} nomeVeiculo - O nome do novo veículo a ser adicionado.
     * @returns {object} Um objeto com o status da operação.
     */
    /**
     * Retorna uma lista de todos os nomes de veículos cadastrados.
     */
    function getVeiculosList() {
      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
        if (!sheet) {
          Logger.log("Planilha 'Config' não encontrada.");
          return []; 
        }

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return [];

        // Lê apenas a primeira coluna (A)
        const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
        
        // Mapeia para um array de strings e remove espaços em branco
        const veiculos = data.map(row => String(row[0]).trim()).filter(nome => nome !== "");

        // Ordena alfabeticamente
        veiculos.sort((a, b) => a.localeCompare(b));
        
        return veiculos;
      } catch (e) {
        Logger.log("Erro em getVeiculosList: " + e.message);
        return [];
      }
    }

    /**
     * Adiciona um novo nome de veículo à planilha "Veiculos".
     * @param {string} nomeVeiculo - O nome do novo veículo a ser adicionado.
     * @returns {object} Um objeto com o status da operação.
     */
    function adicionarNovoVeiculo(nomeVeiculo) {
      if (!nomeVeiculo || typeof nomeVeiculo !== 'string' || nomeVeiculo.trim() === '') {
        return { status: 'error', message: 'O nome do veículo não pode estar vazio.' };
      }

      const nomeLimpo = nomeVeiculo.trim().toUpperCase(); // Padroniza para maiúsculas

      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
        
        // Pega todos os valores da coluna C para verificar se o veículo já existe
        const rangeVeiculos = sheet.getRange("C2:C" + sheet.getLastRow());
        const veiculosExistentes = rangeVeiculos.getValues().map(row => String(row[0]).trim().toUpperCase());

        // --- LÓGICA DE VALIDAÇÃO INTELIGENTE ---
        const semelhançaMinima = 2; // Aceita até 2 letras diferentes. Você pode ajustar este valor.

        for (const existente of veiculosExistentes) {
          const distancia = levenshteinDistance(nomeLimpo, existente);

          if (distancia === 0) {
            return { status: 'exists', message: 'Este veículo já está cadastrado.' };
          }
          
          if (distancia <= semelhançaMinima) {
            return { status: 'similar', message: `Erro: O nome '${nomeVeiculo}' é muito parecido com '${existente}', que já está cadastrado.` };
          }
        }
        // --- FIM DA VALIDAÇÃO ---   
        
        // Encontra a próxima linha vazia na coluna C e adiciona o novo veículo lá
        const proximaLinhaVazia = rangeVeiculos.getValues().filter(String).length + 2;
        sheet.getRange(proximaLinhaVazia, 3).setValue(nomeLimpo);
        
        return { status: 'ok', message: 'Veículo adicionado com sucesso!', novoVeiculo: nomeLimpo };
      } catch (e) {
        Logger.log("Erro em adicionarNovoVeiculo: " + e.message);
        return { status: 'error', message: 'Ocorreu um erro ao salvar o novo veículo.' };
      }
    }

    /**
     * Calcula a Distância de Levenshtein entre duas strings.
     * Retorna o número de edições necessárias para transformar uma string na outra.
     */
    function levenshteinDistance(a, b) {
      if (a.length === 0) return b.length;
      if (b.length === 0) return a.length;

      const matrix = Array(b.length + 1).fill(null).map(() => Array(a.length + 1).fill(null));

      for (let i = 0; i <= a.length; i++) {
        matrix[0][i] = i;
      }
      for (let j = 0; j <= b.length; j++) {
        matrix[j][0] = j;
      }

      for (let j = 1; j <= b.length; j++) {
        for (let i = 1; i <= a.length; i++) {
          const cost = a[i - 1] === b[j - 1] ? 0 : 1;
          matrix[j][i] = Math.min(
            matrix[j][i - 1] + 1,      // Deletion
            matrix[j - 1][i] + 1,      // Insertion
            matrix[j - 1][i - 1] + cost // Substitution
          );
        }
      }

      return matrix[b.length][a.length];
    }

    // ===============================================
    // FUNÇÕES PARA RASCUNHO
    // ===============================================

    /**
     * Função de teste para verificar comunicação backend
     */
    function testarComunicacao() {
      console.log('✅ [TESTE] Função testarComunicacao chamada com sucesso');
      return {
        status: 'success',
        message: 'Comunicação funcionando',
        timestamp: new Date().toISOString()
      };
    }

    /**
     * Função de teste ainda mais simples
     */
    function testeSimples() {
      return 'OK';
    }

    /**
     * ===============================================
     * BACKEND - SISTEMA DE RASCUNHOS
     * Google Apps Script Functions
     * ===============================================
     */

    /**
     * Salva um rascunho na planilha
     * @param {Object} dadosRascunho - Dados do rascunho a ser salvo
     * @returns {Object} - Resposta com status e ID do rascunho
     */
    function salvarRascunho(dadosRascunho) {
      try {
        console.log('📝 Salvando rascunho:', dadosRascunho);
        
        // Validações básicas
        if (!dadosRascunho.fornecedor || !dadosRascunho.fornecedor.trim()) {
          return {
            status: 'error',
            message: 'Fornecedor é obrigatório para salvar o rascunho.'
          };
        }
        
        if (!dadosRascunho.itens || !Array.isArray(dadosRascunho.itens) || dadosRascunho.itens.length === 0) {
          return {
            status: 'error',
            message: 'Pelo menos um item é obrigatório para salvar o rascunho.'
          };
        }
        
        // Validar se pelo menos um item tem descrição
        const itemValido = dadosRascunho.itens.some(item => item.descricao && item.descricao.trim());
        if (!itemValido) {
          return {
            status: 'error',
            message: 'Pelo menos um item deve ter uma descrição.'
          };
        }
        
        // Obter a planilha
        const planilha = SpreadsheetApp.openById(PLANILHA_ID);
        const aba = planilha.getSheetByName('Pedidos') || planilha.insertSheet('Pedidos');
        
        // Gerar ID único para o rascunho
        const agora = new Date();
        const ano = agora.getFullYear();
        const mes = String(agora.getMonth() + 1).padStart(2, '0');
        const dia = String(agora.getDate()).padStart(2, '0');
        const timestamp = agora.getTime();
        const rascunhoId = `RASC-${ano}${mes}${dia}-${timestamp}`;
        
        // Obter dados do fornecedor da aba Fornecedores (se existir)
        const fornecedoresSheet = planilha.getSheetByName('Fornecedores');
        let fornecedorCnpj = '';
        let fornecedorEndereco = '';
        let condicaoPagamentoFornecedor = '';
        let formaPagamentoFornecedor = '';
        let estadoFornecedor = '';

        if (fornecedoresSheet) {
          const fornecedoresData = fornecedoresSheet.getRange(2, 1, fornecedoresSheet.getLastRow() - 1, fornecedoresSheet.getLastColumn()).getValues();
          const foundFornecedor = fornecedoresData.find(row => String(row[1]) === dadosRascunho.fornecedor); 
          if (foundFornecedor) {
            fornecedorCnpj = String(foundFornecedor[3] || '');
            fornecedorEndereco = String(foundFornecedor[4] || '');
            condicaoPagamentoFornecedor = String(foundFornecedor[5] || '');
            formaPagamentoFornecedor = String(foundFornecedor[6] || '');
            estadoFornecedor = String(foundFornecedor[10] || ''); // Coluna 11 (índice 10) = Estado
          }
        }

        // Preparar dados para salvar (mesma estrutura do salvarPedido)
        const dadosParaSalvar = {
          'Número do Pedido': "'" + rascunhoId, // Usando ID do rascunho como número
          'Empresa': "'" + (dadosRascunho.empresa || Session.getActiveUser().getEmail()),
          'Data': dadosRascunho.data ? formatarDataParaISO(dadosRascunho.data) : formatarDataParaISO(agora),
          'Fornecedor': dadosRascunho.fornecedor.trim(),
          'CNPJ Fornecedor': fornecedorCnpj,
          'Endereço Fornecedor': fornecedorEndereco,
          'Estado Fornecedor': estadoFornecedor,
          'Condição Pagamento Fornecedor': condicaoPagamentoFornecedor,
          'Forma Pagamento Fornecedor': formaPagamentoFornecedor,
          'Placa Veiculo': dadosRascunho.placaVeiculo || '',
          'Nome Veiculo': dadosRascunho.nomeVeiculo || '',
          'Observacoes': dadosRascunho.observacoes || '',
          'Total Geral': dadosRascunho.totalGeral || 0,
          'Status': 'RASCUNHO', // Diferença principal: status RASCUNHO em vez de "Em Aberto"
          'Itens': JSON.stringify(dadosRascunho.itens),
          'Data Ultima Edicao': formatarDataParaISO(agora) // Sempre usar data/hora atual padronizada
        };
        
        // Verificar se é uma atualização de rascunho existente
        if (dadosRascunho.rascunhoId) {
          const linhaExistente = encontrarLinhaRascunho(aba, dadosRascunho.rascunhoId);
          if (linhaExistente > 0) {
            // Atualizar rascunho existente usando a mesma estrutura
            dadosParaSalvar['Número do Pedido'] = "'" + dadosRascunho.rascunhoId;
            salvarDadosNaPlanilha(aba, dadosParaSalvar, linhaExistente);
            
            console.log('✅ Rascunho atualizado com sucesso:', dadosRascunho.rascunhoId);
            return {
              status: 'success',
              message: 'Rascunho atualizado com sucesso!',
              rascunhoId: dadosRascunho.rascunhoId
            };
          }
        }
        
        // Salvar novo rascunho usando a mesma estrutura da função salvarPedido
        salvarDadosNaPlanilha(aba, dadosParaSalvar);
        
        console.log('✅ Rascunho salvo com sucesso:', rascunhoId);
        return {
          status: 'success',
          message: 'Rascunho salvo com sucesso!',
          rascunhoId: rascunhoId
        };
        
      } catch (error) {
        console.error('❌ Erro ao salvar rascunho:', error);
        return {
          status: 'error',
          message: 'Erro interno ao salvar rascunho: ' + error.message
        };
      }
    }

    /**
     * Busca todos os rascunhos de uma empresa
     * @param {string} empresaId - ID da empresa
     * @returns {Object} - Lista de rascunhos
     */
    function buscarRascunhos(empresaId) {
      console.log('🔍 [BACKEND] === INÍCIO buscarRascunhos ===');
      console.log('🔍 [BACKEND] Parâmetro empresaId:', empresaId);
      console.log('🔍 [BACKEND] Tipo do empresaId:', typeof empresaId);
      
      try {
        // ID da planilha definido localmente
        var planilhaId = '1J7CE_BZ8eUsXhjkmgxAIIWjMTOr2FfSfIMONqE4UpHA';
        
        // Validação básica
        if (!empresaId) {
          console.error('❌ [BACKEND] empresaId é obrigatório');
          var erro = {
            status: 'error',
            message: 'ID da empresa é obrigatório',
            rascunhos: []
          };
          console.log('📤 [BACKEND] Retornando erro de validação:', erro);
          return erro;
        }
        
        console.log('✅ [BACKEND] Validação OK, tentando acessar planilha...');
        console.log('🔍 [BACKEND] planilhaId:', planilhaId);
        
        var planilha = SpreadsheetApp.openById(planilhaId);
        console.log('✅ [BACKEND] Planilha acessada com sucesso');
        
        var aba = planilha.getSheetByName('Pedidos');
        
        if (!aba) {
          console.log('📋 [BACKEND] Aba Pedidos não encontrada');
          var sucesso = {
            status: 'success',
            rascunhos: [],
            message: 'Aba Pedidos não encontrada'
          };
          console.log('📤 [BACKEND] Retornando lista vazia:', sucesso);
          return sucesso;
        }
        
        console.log('✅ [BACKEND] Aba Pedidos encontrada');
        
        var dados = aba.getDataRange().getValues();
        console.log('📊 [BACKEND] Dados obtidos - Total de linhas:', dados.length);
        
        if (dados.length < 2) {
          console.log('📋 [BACKEND] Planilha vazia ou só cabeçalho');
          var vazio = {
            status: 'success',
            rascunhos: [],
            message: 'Planilha vazia'
          };
          console.log('📤 [BACKEND] Retornando planilha vazia:', vazio);
          return vazio;
        }
        
        var cabecalhos = dados[0];
        var rascunhos = [];
        
        console.log('📊 [BACKEND] Cabeçalhos:', cabecalhos);
        
        // Encontrar índices das colunas (usando os nomes reais da planilha)
        var indices = {
          numeroPedido: cabecalhos.indexOf('Número do Pedido'),
          empresa: cabecalhos.indexOf('Empresa'),
          status: cabecalhos.indexOf('Status'),
          data: cabecalhos.indexOf('Data'),
          fornecedor: cabecalhos.indexOf('Fornecedor'),
          nomeVeiculo: cabecalhos.indexOf('Nome Veiculo'),
          placaVeiculo: cabecalhos.indexOf('Placa Veiculo'),
          observacoes: cabecalhos.indexOf('Observacoes'),
          itens: cabecalhos.indexOf('Itens'),
          totalGeral: cabecalhos.indexOf('Total Geral')
        };
        
        console.log('📊 [BACKEND] Índices encontrados:', indices);
        
        // Verificar colunas críticas
        if (indices.status === -1) {
          console.error('❌ [BACKEND] Coluna Status não encontrada');
          var erro = {
            status: 'error',
            message: 'Coluna Status não encontrada na planilha',
            rascunhos: []
          };
          console.log('📤 [BACKEND] Retornando erro de estrutura:', erro);
          return erro;
        }
        
        if (indices.empresa === -1) {
          console.error('❌ [BACKEND] Coluna Empresa não encontrada');
          var erro = {
            status: 'error',
            message: 'Coluna Empresa não encontrada na planilha',
            rascunhos: []
          };
          console.log('📤 [BACKEND] Retornando erro de estrutura:', erro);
          return erro;
        }
        
        console.log('✅ [BACKEND] Estrutura da planilha validada');
        
        // Processar dados
        var rascunhosEncontrados = 0;
        var empresaIdStr = String(empresaId).trim();
        
        console.log('🔍 [BACKEND] Processando linhas para empresa:', empresaIdStr);
        
        for (var i = 1; i < dados.length; i++) {
          var linha = dados[i];
          var statusLinha = linha[indices.status];
          var empresaLinha = linha[indices.empresa];
          
          // Debug das primeiras 3 linhas
          if (i <= 3) {
            console.log('📊 [BACKEND] Linha ' + i + ': Status="' + statusLinha + '", Empresa="' + empresaLinha + '"');
          }
          
          // Verificar se é rascunho da empresa
          if (statusLinha === 'RASCUNHO' && empresaLinha) {
            // Remover apóstrofo do campo empresa para comparação
            var empresaNaPlanilha = String(empresaLinha).replace(/'/g, '').trim();
            
            if (i <= 3) {
              console.log('🔍 [BACKEND] Comparando linha ' + i + ': "' + empresaNaPlanilha + '" === "' + empresaIdStr + '"');
            }
            
            if (empresaNaPlanilha === empresaIdStr) {
              rascunhosEncontrados++;
              console.log('✅ [BACKEND] Rascunho ' + rascunhosEncontrados + ' encontrado na linha ' + (i + 1));
              
              var itensArray = [];
              try {
                if (linha[indices.itens]) {
                  itensArray = JSON.parse(linha[indices.itens]);
                }
              } catch (e) {
                console.warn('⚠️ [BACKEND] Erro ao parsear itens:', linha[indices.numeroPedido]);
                itensArray = [];
              }
              
              var rascunho = {
                id: linha[indices.numeroPedido] ? String(linha[indices.numeroPedido]).replace(/'/g, '') : '',
                data: linha[indices.data] ? String(linha[indices.data]) : '',
                fornecedor: linha[indices.fornecedor] || '',
                nomeVeiculo: linha[indices.nomeVeiculo] || '',
                placaVeiculo: linha[indices.placaVeiculo] || '',
                observacoes: linha[indices.observacoes] || '',
                itens: itensArray,
                totalGeral: Number(linha[indices.totalGeral]) || 0
              };
              
              rascunhos.push(rascunho);
            }
          }
        }
        
        console.log(`✅ [BACKEND] Processamento concluído - ${rascunhos.length} rascunhos encontrados`);
        
        // Ordenar por data (mais recente primeiro)
        try {
          rascunhos.sort((a, b) => new Date(b.data) - new Date(a.data));
          console.log('✅ [BACKEND] Rascunhos ordenados por data');
        } catch (sortError) {
          console.warn('⚠️ [BACKEND] Erro ao ordenar:', sortError);
        }
        
        const resultado = {
          status: 'success',
          rascunhos: rascunhos,
          message: `${rascunhos.length} rascunho(s) encontrado(s)`
        };
        
        console.log('📤 [BACKEND] Retornando resultado final:', resultado);
        return resultado;
        
      } catch (error) {
        console.error('❌ [BACKEND] Erro na função buscarRascunhos:', error);
        console.error('❌ [BACKEND] Stack trace:', error.stack);
        
        const erro = {
          status: 'error',
          message: 'Erro interno: ' + error.message,
          rascunhos: []
        };
        
        console.log('📤 [BACKEND] Retornando erro:', erro);
        return erro;
      } finally {
        console.log('🔍 [BACKEND] === FIM buscarRascunhos ===');
      }
    }

    /**
     * Busca um rascunho específico por ID
     * @param {string} rascunhoId - ID do rascunho
     * @returns {Object} - Dados do rascunho
     */
    function buscarRascunhoPorId(rascunhoId) {
      try {
        console.log('🔍 [BUSCAR ID] Buscando rascunho por ID:', rascunhoId);
        
        // ID da planilha definido localmente
        var planilhaId = '1J7CE_BZ8eUsXhjkmgxAIIWjMTOr2FfSfIMONqE4UpHA';
        var planilha = SpreadsheetApp.openById(planilhaId);
        var aba = planilha.getSheetByName('Pedidos');
        
        if (!aba) {
          return {
            status: 'error',
            message: 'Planilha de pedidos não encontrada.'
          };
        }
        
        var dados = aba.getDataRange().getValues();
        var cabecalhos = dados[0];
        
        // Buscar possíveis variações do nome da coluna de data última edição
        var possiveisNomes = ['Data Ultima Edicao', 'Data Última Edição', 'Ultima Edicao', 'Última Edição', 'Data da Ultima Edicao'];
        var indiceDataUltimaEdicao = -1;
        
        for (var nomeColuna of possiveisNomes) {
          var indice = cabecalhos.indexOf(nomeColuna);
          if (indice !== -1) {
            indiceDataUltimaEdicao = indice;
            console.log('🔍 [BUSCAR ID] ✅ Coluna encontrada:', nomeColuna, 'no índice:', indice);
            break;
          }
        }
        
        // Encontrar índices das colunas (usando os nomes reais da planilha)
        var indices = {
          numeroPedido: cabecalhos.indexOf('Número do Pedido'),
          status: cabecalhos.indexOf('Status'),
          data: cabecalhos.indexOf('Data'),
          empresa: cabecalhos.indexOf('Empresa'),
          fornecedor: cabecalhos.indexOf('Fornecedor'),
          nomeVeiculo: cabecalhos.indexOf('Nome Veiculo'),
          placaVeiculo: cabecalhos.indexOf('Placa Veiculo'),
          observacoes: cabecalhos.indexOf('Observacoes'),
          itens: cabecalhos.indexOf('Itens'),
          totalGeral: cabecalhos.indexOf('Total Geral'),
          dataUltimaEdicao: indiceDataUltimaEdicao
        };
        
        console.log('🔍 [BUSCAR ID] Processando ' + (dados.length - 1) + ' linhas...');
        
        // Procurar o rascunho
        for (var i = 1; i < dados.length; i++) {
          var linha = dados[i];
          
          var numeroRascunho = linha[indices.numeroPedido] ? String(linha[indices.numeroPedido]).replace(/'/g, '') : '';
          if (numeroRascunho === rascunhoId && linha[indices.status] === 'RASCUNHO') {
            var itensArray = [];
            try {
              if (linha[indices.itens]) {
                itensArray = JSON.parse(linha[indices.itens]);
              }
            } catch (e) {
              console.warn('⚠️ [BUSCAR ID] Erro ao parsear itens do rascunho:', rascunhoId);
              itensArray = [];
            }
            
            var rascunho = {
              id: numeroRascunho,
              data: linha[indices.data] ? String(linha[indices.data]) : '',
              empresa: linha[indices.empresa] ? String(linha[indices.empresa]).replace(/'/g, '') : '',
              fornecedor: linha[indices.fornecedor] || '',
              nomeVeiculo: linha[indices.nomeVeiculo] || '',
              placaVeiculo: linha[indices.placaVeiculo] || '',
              observacoes: linha[indices.observacoes] || '',
              itens: itensArray,
              totalGeral: Number(linha[indices.totalGeral]) || 0,
              dataUltimaEdicao: (indices.dataUltimaEdicao !== -1 && linha[indices.dataUltimaEdicao]) ? String(linha[indices.dataUltimaEdicao]) : ''
            };
            
            console.log('✅ [BUSCAR ID] Rascunho encontrado:', rascunhoId);
            return {
              status: 'success',
              rascunho: rascunho
            };
          }
        }
        
        console.log('❌ [BUSCAR ID] Rascunho não encontrado:', rascunhoId);
        return {
          status: 'error',
          message: 'Rascunho não encontrado.'
        };
        
      } catch (error) {
        console.error('❌ [BUSCAR ID] Erro ao buscar rascunho por ID:', error);
        return {
          status: 'error',
          message: 'Erro ao buscar rascunho: ' + error.message
        };
      }
    }

    /**
     * Finaliza um rascunho como pedido oficial
     * @param {string} rascunhoId - ID do rascunho
     * @returns {Object} - Resultado da operação
     */
    function finalizarRascunho(rascunhoId) {
      try {
        console.log('✅ Finalizando rascunho:', rascunhoId);
        
        // Buscar dados do rascunho
        const resultadoBusca = buscarRascunhoPorId(rascunhoId);
        if (resultadoBusca.status !== 'success') {
          return resultadoBusca;
        }
        
        const dadosRascunho = resultadoBusca.rascunho;
        console.log('📋 Dados do rascunho encontrado:', dadosRascunho);
        
        // Validar dados para finalização
        const validacao = validarDadosParaPedido(dadosRascunho);
        if (!validacao.valido) {
          return {
            status: 'error',
            message: validacao.mensagem
          };
        }
        
        // Obter empresa do rascunho ou usar empresa do usuário logado
        let empresaCodigo = dadosRascunho.empresa;
        
        // Se não houver empresa no rascunho, tentar obter do usuário logado
        if (!empresaCodigo) {
          const usuarioLogado = obterUsuarioLogado();
          if (usuarioLogado && usuarioLogado.idEmpresa) {
            empresaCodigo = usuarioLogado.idEmpresa;
          } else {
            return {
              status: 'error',
              message: 'Não foi possível determinar a empresa para gerar o número do pedido.'
            };
          }
        }
        
        console.log('🏢 Empresa para geração do pedido:', empresaCodigo);
        
        // Gerar número do pedido sequencial por empresa
        const numeroPedido = getProximoNumeroPedido(empresaCodigo);
        console.log('📝 Número do pedido gerado:', numeroPedido);
        
        // Preparar dados do pedido
        const dadosPedido = {
          numero: numeroPedido,
          data: formatarDataParaISO(dadosRascunho.data || new Date()),
          fornecedor: dadosRascunho.fornecedor,
          nomeVeiculo: dadosRascunho.nomeVeiculo || '',
          placaVeiculo: dadosRascunho.placaVeiculo || '',
          observacoes: dadosRascunho.observacoes || '',
          itens: dadosRascunho.itens || [],
          totalGeral: dadosRascunho.totalGeral || 0,
          empresa: empresaCodigo
        };
        
        console.log('📦 Dados do pedido preparados:', dadosPedido);
        
        // Salvar como pedido usando função existente
        const resultadoSalvamento = salvarPedido(dadosPedido);
        
        if (resultadoSalvamento.status === 'ok') {
          // Excluir o rascunho
          const resultadoExclusao = excluirRascunho(rascunhoId);
          
          console.log('✅ Rascunho finalizado como pedido:', numeroPedido);
          return {
            status: 'success',
            message: 'Rascunho finalizado com sucesso!',
            numeroPedido: numeroPedido
          };
        } else {
          return {
            status: 'error',
            message: 'Erro ao finalizar rascunho: ' + resultadoSalvamento.message
          };
        }
        
      } catch (error) {
        console.error('❌ Erro ao finalizar rascunho:', error);
        return {
          status: 'error',
          message: 'Erro interno ao finalizar rascunho: ' + error.message
        };
      }
    }

    /**
     * Exclui um rascunho
     * @param {string} rascunhoId - ID do rascunho
     * @returns {Object} - Resultado da operação
     */
    function excluirRascunho(rascunhoId) {
      try {
        console.log('🗑️ Excluindo rascunho:', rascunhoId);
        
        const planilha = SpreadsheetApp.openById(PLANILHA_ID);
        const aba = planilha.getSheetByName('Pedidos');
        
        if (!aba) {
          return {
            status: 'error',
            message: 'Planilha de pedidos não encontrada.'
          };
        }
        
        const linhaRascunho = encontrarLinhaRascunho(aba, rascunhoId);
        
        if (linhaRascunho > 0) {
          aba.deleteRow(linhaRascunho);
          
          console.log('✅ Rascunho excluído:', rascunhoId);
          return {
            status: 'success',
            message: 'Rascunho excluído com sucesso!'
          };
        } else {
          return {
            status: 'error',
            message: 'Rascunho não encontrado.'
          };
        }
        
      } catch (error) {
        console.error('❌ Erro ao excluir rascunho:', error);
        return {
          status: 'error',
          message: 'Erro ao excluir rascunho: ' + error.message
        };
      }
    }