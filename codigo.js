    // ===============================================
    // CONFIGURACAO GLOBAL DA PLANILHA
    // ===============================================
    var PLANILHA_ID = '1J7CE_BZ8eUsXhjkmgxAIIWjMTOr2FfSfIMONqE4UpHA';

    /**
     * FUNCAO DE TESTE PARA VERIFICAR COLUNAS DA PLANILHA PEDIDOS
     */
    function verificarColunasPlanilaaPedidos() {
      try {
        const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
        if (!sheet) {
          return {
            status: 'error',
            message: 'Planilha "Pedidos" não encontrada'
          };
        }

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        const colunasRelacionadas = {
          'Fornecedor': headers.indexOf('Fornecedor'),
          'CNPJ Fornecedor': headers.indexOf('CNPJ Fornecedor'),
          'Endereço Fornecedor': headers.indexOf('Endereço Fornecedor'),
          'Estado Fornecedor': headers.indexOf('Estado Fornecedor'),
          'Condição Pagamento Fornecedor': headers.indexOf('Condição Pagamento Fornecedor'),
          'Forma Pagamento Fornecedor': headers.indexOf('Forma Pagamento Fornecedor')
        };

        return {
          status: 'success',
          headers: headers,
          colunasRelacionadas: colunasRelacionadas,
          estadoFornecedorExiste: colunasRelacionadas['Estado Fornecedor'] !== -1,
          totalColunas: headers.length
        };
        
      } catch (error) {
        return {
          status: 'error',
          message: 'Erro ao verificar colunas: ' + error.message
        };
      }
    }

    /**
     * FUNCAO DE TESTE PARA COMUNICACAO BACKEND
     */
    function testarComunicacao() {
      return {
        status: 'success',
        message: 'Comunicacao funcionando',
        timestamp: new Date().toISOString()
      };
    }

    /**
     * FUNCAO ALTERNATIVA PARA TESTE DE COMUNICACAO
     */
    function testeComunicacao() {
      return 'FUNCIONANDO';
    }

    /**
     * FUNCOES DE TESTE SEM PARAMETROS PARA EXECUCAO DIRETA NO APPS SCRIPT
     */
    function testeBuscarRascunhos001() {
      return buscarRascunhosCorrigida('001');
    }

    function testeBuscarRascunhosSimples001() {
      return buscarRascunhosSimples('001');
    }

    /**
     * FUNÇÃO DE TESTE PARA DIAGNÓSTICO - EXECUTE ESTA!
     */
    function testeCompleto() {
      return diagnosticoCompleto();
    }

    /**
     * FUNÇÃO DE TESTE ESPECÍFICA PARA ESTADO DO FORNECEDOR
     * Substitua "NOME_DO_FORNECEDOR" pelo nome real de um fornecedor da sua planilha
     */
    function testeEstado() {
      // ALTERE AQUI: coloque o nome exato de um fornecedor da sua planilha
      return testarEstadoFornecedor("ACLENILTON IVES DA SILVA");
    }

    /**
     * FUNÇÃO DE TESTE PARA VERIFICAR A CORREÇÃO DO ÍNDICE
     */
    function testeCorrecaoIndice() {
      return diagnosticoCompleto();
    }

    /**
     * DIAGNOSTICO COMPLETO DO SISTEMA
     * EXECUTE ESTA FUNCAO NO GOOGLE APPS SCRIPT, NAO NO NAVEGADOR!
     */
    function diagnosticoSistema() {
      try {
        var resultado = {
          timestamp: new Date().toISOString(),
          googleAppsScript: 'OK',
          planilhaAcess: 'TENTANDO...',
          planilhaId: PLANILHA_ID
        };
        
        // Testar acesso a planilha
        try {
          var planilha = SpreadsheetApp.openById(PLANILHA_ID);
          resultado.planilhaAcess = 'OK';
          resultado.planilhaNome = planilha.getName();
          
          // Testar acesso a aba Pedidos
          var aba = planilha.getSheetByName('Pedidos');
          if (aba) {
            resultado.abaPedidos = 'OK';
            resultado.totalLinhas = aba.getLastRow();
          } else {
            resultado.abaPedidos = 'NAO ENCONTRADA';
          }
          
        } catch (e) {
          resultado.planilhaAcess = 'ERRO: ' + e.message;
        }
        
        return resultado;
      } catch (error) {
        return {
          erro: error.message,
          timestamp: new Date().toISOString()
        };
      }
    }

    /**
     * TESTE ESPECÍFICO PARA BUSCAR RASCUNHOS COM DEBUG
     */
    function testarBuscarRascunhos(empresaId) {
      try {
        console.log('🔧 [DEBUG] Testando buscarRascunhos com empresaId:', empresaId);
        
        // Usar ID da planilha diretamente
        var planilhaId = '1J7CE_BZ8eUsXhjkmgxAIIWjMTOr2FfSfIMONqE4UpHA';
        console.log('🔧 [DEBUG] Usando planilhaId:', planilhaId);
        
        // Testar acesso à planilha
        var planilha = SpreadsheetApp.openById(planilhaId);
        console.log('🔧 [DEBUG] Planilha acessada:', planilha.getName());
        
        // Testar acesso à aba
        var aba = planilha.getSheetByName('Pedidos');
        if (!aba) {
          return {
            status: 'error',
            message: 'Aba Pedidos não encontrada',
            debug: 'Aba inexistente'
          };
        }
        
        console.log('🔧 [DEBUG] Aba encontrada');
        
        // Chamar a função real
        var resultado = buscarRascunhos(empresaId);
        console.log('🔧 [DEBUG] Resultado da buscarRascunhos:', resultado);
        
        return {
          status: 'success',
          message: 'Teste concluído',
          resultadoBusca: resultado,
          debug: 'Função executada com sucesso'
        };
        
      } catch (error) {
        console.error('🔧 [DEBUG] Erro no teste:', error);
        return {
          status: 'error',
          message: error.message,
          debug: error.stack
        };
      }
    }

    /**
     * BUSCAR RASCUNHOS VERSAO SIMPLIFICADA PARA TESTE
     */
    function buscarRascunhosSimples(empresaId) {
      try {
        console.log('🔍 [SIMPLES] Iniciando busca para empresa:', empresaId);
        
        // ID da planilha diretamente
        var planilhaId = '1J7CE_BZ8eUsXhjkmgxAIIWjMTOr2FfSfIMONqE4UpHA';
        var planilha = SpreadsheetApp.openById(planilhaId);
        var aba = planilha.getSheetByName('Pedidos');
        
        if (!aba) {
          return {
            status: 'success',
            rascunhos: [],
            message: 'Aba Pedidos não encontrada'
          };
        }
        
        var dados = aba.getDataRange().getValues();
        console.log('🔍 [SIMPLES] Total de linhas:', dados.length);
        
        if (dados.length < 2) {
          return {
            status: 'success',
            rascunhos: [],
            message: 'Planilha vazia'
          };
        }
        
        var rascunhos = [];
        var cabecalhos = dados[0];
        
        // Encontrar colunas importantes
        var colunaStatus = cabecalhos.indexOf('Status');
        var colunaEmpresa = cabecalhos.indexOf('Empresa');
        var colunaNumero = cabecalhos.indexOf('Número do Pedido');
        var colunaFornecedor = cabecalhos.indexOf('Fornecedor');
        
        console.log('🔍 [SIMPLES] Colunas - Status:', colunaStatus, 'Empresa:', colunaEmpresa);
        
        // Processar dados
        for (var i = 1; i < dados.length; i++) {
          var linha = dados[i];
          
          if (linha[colunaStatus] === 'RASCUNHO' && linha[colunaEmpresa]) {
            var empresaNaLinha = String(linha[colunaEmpresa]).replace(/'/g, '').trim();
            
            if (empresaNaLinha === String(empresaId).trim()) {
              rascunhos.push({
                id: linha[colunaNumero] ? String(linha[colunaNumero]).replace(/'/g, '') : '',
                fornecedor: linha[colunaFornecedor] || '',
                empresa: empresaNaLinha
              });
            }
          }
        }
        
        console.log('🔍 [SIMPLES] Rascunhos encontrados:', rascunhos.length);
        
        return {
          status: 'success',
          rascunhos: rascunhos,
          message: rascunhos.length + ' rascunho(s) encontrado(s)'
        };
        
      } catch (error) {
        console.error('🔍 [SIMPLES] Erro:', error);
        return {
          status: 'error',
          message: error.message,
          rascunhos: []
        };
      }
    }

    /**
     * VERSAO CORRIGIDA DA FUNCAO BUSCAR RASCUNHOS
     * Esta versão garante compatibilidade total com Google Apps Script
     */
    function buscarRascunhosCorrigida(empresaId) {
      console.log('🔍 [CORRIGIDA] === INÍCIO buscarRascunhosCorrigida ===');
      console.log('🔍 [CORRIGIDA] Parâmetro empresaId RECEBIDO:', empresaId);
      console.log('🔍 [CORRIGIDA] Tipo do empresaId:', typeof empresaId);
      console.log('🔍 [CORRIGIDA] empresaId é null?', empresaId === null);
      console.log('🔍 [CORRIGIDA] empresaId é undefined?', empresaId === undefined);
      console.log('🔍 [CORRIGIDA] empresaId convertido para string:', String(empresaId));
      
      try {
        // ID da planilha definido localmente para evitar problemas de escopo
        var planilhaId = '1J7CE_BZ8eUsXhjkmgxAIIWjMTOr2FfSfIMONqE4UpHA';
        
        // Validação básica
        if (!empresaId) {
          console.error('❌ [CORRIGIDA] empresaId é obrigatório');
          console.error('❌ [CORRIGIDA] Valor recebido:', empresaId);
          return {
            status: 'error',
            message: 'ID da empresa é obrigatório. Valor recebido: ' + empresaId,
            rascunhos: []
          };
        }
        
        console.log('✅ [CORRIGIDA] Validação OK, acessando planilha...');
        var planilha = SpreadsheetApp.openById(planilhaId);
        var aba = planilha.getSheetByName('Pedidos');
        
        if (!aba) {
          console.log('📋 [CORRIGIDA] Aba Pedidos não encontrada');
          return {
            status: 'success',
            rascunhos: [],
            message: 'Aba Pedidos não encontrada'
          };
        }
        
        var dados = aba.getDataRange().getValues();
        console.log('📊 [CORRIGIDA] Total de linhas:', dados.length);
        
        if (dados.length < 2) {
          return {
            status: 'success',
            rascunhos: [],
            message: 'Planilha vazia'
          };
        }
        
        var cabecalhos = dados[0];
        var rascunhos = [];
        var empresaIdStr = String(empresaId).trim();
        
        console.log('🔍 [CORRIGIDA] empresaIdStr após conversão:', empresaIdStr);
        
        // Buscar possíveis variações do nome da coluna de data última edição
        var possiveisNomes = ['Data Ultima Edicao', 'Data Última Edição', 'Ultima Edicao', 'Última Edição', 'Data da Ultima Edicao'];
        var indiceDataUltimaEdicao = -1;
        
        for (var nomeColuna of possiveisNomes) {
          var indice = cabecalhos.indexOf(nomeColuna);
          if (indice !== -1) {
            indiceDataUltimaEdicao = indice;
            console.log('📊 [CORRIGIDA] ✅ Coluna encontrada:', nomeColuna, 'no índice:', indice);
            break;
          }
        }
        
        if (indiceDataUltimaEdicao === -1) {
          console.log('📊 [CORRIGIDA] ⚠️ Nenhuma coluna de data última edição encontrada. Cabeçalhos disponíveis:', cabecalhos);
        }
        
        // Índices das colunas
        var indices = {
          numeroPedido: cabecalhos.indexOf('Número do Pedido'),
          empresa: cabecalhos.indexOf('Empresa'),
          status: cabecalhos.indexOf('Status'),
          data: cabecalhos.indexOf('Data'),
          fornecedor: cabecalhos.indexOf('Fornecedor'),
          estadoFornecedor: cabecalhos.indexOf('Estado Fornecedor'),
          nomeVeiculo: cabecalhos.indexOf('Nome Veiculo'),
          placaVeiculo: cabecalhos.indexOf('Placa Veiculo'),
          observacoes: cabecalhos.indexOf('Observacoes'),
          itens: cabecalhos.indexOf('Itens'),
          totalGeral: cabecalhos.indexOf('Total Geral'),
          dataUltimaEdicao: indiceDataUltimaEdicao
        };
        
        console.log('📊 [CORRIGIDA] Índices das colunas:', indices);
        console.log('📊 [CORRIGIDA] ⚠️ Coluna Data Ultima Edicao encontrada?', indices.dataUltimaEdicao !== -1 ? 'SIM' : 'NÃO');
        console.log('📊 [CORRIGIDA] Processando ' + (dados.length - 1) + ' linhas...');
        
        // Processar todas as linhas
        for (var i = 1; i < dados.length; i++) {
          var linha = dados[i];
          var statusLinha = linha[indices.status];
          var empresaLinha = linha[indices.empresa];
          
          // Log detalhado das primeiras 5 linhas
          if (i <= 5) {
            console.log('📊 [CORRIGIDA] Linha ' + i + ': Status="' + statusLinha + '", Empresa="' + empresaLinha + '"');
          }
          
          if (statusLinha === 'RASCUNHO' && empresaLinha) {
            var empresaNaPlanilha = String(empresaLinha).replace(/'/g, '').trim();
            
            console.log('🔍 [CORRIGIDA] Linha ' + i + ' é RASCUNHO - Comparando "' + empresaNaPlanilha + '" com "' + empresaIdStr + '"');
            
            if (empresaNaPlanilha === empresaIdStr) {
              console.log('✅ [CORRIGIDA] MATCH! Rascunho encontrado na linha ' + (i + 1));
              
              var itensArray = [];
              try {
                if (linha[indices.itens]) {
                  itensArray = JSON.parse(linha[indices.itens]);
                }
              } catch (e) {
                console.warn('⚠️ [CORRIGIDA] Erro ao parsear itens na linha ' + (i + 1));
                itensArray = [];
              }
                  rascunhos.push({
              id: linha[indices.numeroPedido] ? String(linha[indices.numeroPedido]).replace(/'/g, '') : '',
              data: linha[indices.data] ? String(linha[indices.data]) : '',
              fornecedor: linha[indices.fornecedor] || '',
              estadoFornecedor: linha[indices.estadoFornecedor] || '',
              nomeVeiculo: linha[indices.nomeVeiculo] || '',
              placaVeiculo: linha[indices.placaVeiculo] || '',
              observacoes: linha[indices.observacoes] || '',
              itens: itensArray,
              totalGeral: Number(linha[indices.totalGeral]) || 0,
              dataUltimaEdicao: (indices.dataUltimaEdicao !== -1 && linha[indices.dataUltimaEdicao]) ? String(linha[indices.dataUltimaEdicao]) : ''
            });
            }
          }
        }
        
        console.log('✅ [CORRIGIDA] Processamento concluído - ' + rascunhos.length + ' rascunhos encontrados');
        
        var resultado = {
          status: 'success',
          rascunhos: rascunhos,
          message: rascunhos.length + ' rascunho(s) encontrado(s)'
        };
            console.log('📤 [CORRIGIDA] Retornando resultado:', resultado);
      console.log('📤 [CORRIGIDA] Tentando serializar resultado...');
      try {
        var resultadoSerializado = JSON.stringify(resultado);
        console.log('✅ [CORRIGIDA] Serialização bem-sucedida');
      } catch (serializationError) {
        console.error('❌ [CORRIGIDA] Erro na serialização:', serializationError);
      }
      return resultado;
        
      } catch (error) {
        console.error('❌ [CORRIGIDA] Erro:', error);
        console.error('❌ [CORRIGIDA] Stack:', error.stack);
        return {
          status: 'error',
          message: 'Erro interno: ' + error.message,
          rascunhos: []
        };
      } finally {
        console.log('🔍 [CORRIGIDA] === FIM buscarRascunhosCorrigida ===');
      }
    }

    /**
     * Converte uma string de cabeçalho (ex: "Número do Pedido", "CNPJ Fornecedor")
     * para o formato camelCase compatível com JavaScript (ex: "numeroDoPedido", "cnpjFornecedor").
     * Remove acentos e caracteres não alfanuméricos.
     * @param {string} str O cabeçalho da coluna da planilha.
     * @returns {string} O cabeçalho formatado em camelCase.
     */
    function toCamelCase(str) {
      if (!str) return '';

      // 1. Converte para minúsculas e remove acentos
      let s = String(str)
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '');

      // 2. Substitui qualquer coisa que não seja letra ou número por um espaço
      s = s.replace(/[^a-z0-9]+/g, ' ');

      // 3. Converte para camelCase
      return s.replace(/ (\w)/g, (match, p1) => p1.toUpperCase());
    }


    function doGet() {
      // Ajuste o nome do arquivo HTML principal aqui se ele mudou
      return HtmlService.createHtmlOutputFromFile('Login')
        .setTitle('Portal de Compras')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    function incluir(pagina) {
      return HtmlService.createTemplateFromFile(pagina).evaluate().getContent();
    }

    function validarLogin(usuario, senha, empresaSelecionada) {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return { status: 'erro', message: 'Nenhum usuário cadastrado.' };
      }
      const dados = sheet.getRange(2, 1, lastRow - 1, 7).getValues(); // Lê até a coluna 7 (Empresas)

      // Obtém lista de empresas para exibir nome (busca na tabela "Empresas")
      const empresasSheet = SpreadsheetApp.getActive().getSheetByName('Empresas');
      let empresasLista = {};
      if (empresasSheet && empresasSheet.getLastRow() >= 2) {
        const empresasData = empresasSheet.getRange(2, 1, empresasSheet.getLastRow() - 1, 2).getValues();
        empresasData.forEach(([codigo, nome]) => {
          empresasLista[String(codigo).trim()] = String(nome).trim();
        });
      }

      for (let i = 0; i < dados.length; i++) {
        const [id, nome, user, passHashDaPlanilha, perfil, status, empresasStr] = dados[i];


        // 1. Gera o hash da senha que o usuário digitou no formulário.
        const senhaDigitadaHash = gerarHash(String(senha));

        // 2. Compara o nome de usuário E o HASH da senha.
        if (String(usuario) === String(user) && senhaDigitadaHash === passHashDaPlanilha) {

            if (status === 'Ativo') {
            // Processa lista de empresas permitidas
            //const empresasCodigos = String(empresasStr || '')
            //.split(',')
            //.map(e => e.trim().padStart(3, '0'))
            //.filter(e => e !== "");
            // 1. Pega a lista de códigos de empresa que o usuário tem permissão. (SEU CÓDIGO)
            const empresasPermitidas = String(empresasStr || '').split(',').map(e => e.trim());

            // 2. Verifica se a empresa que o usuário SELECIONOU está na lista de permissões dele.
            if (!empresasPermitidas.includes(String(empresaSelecionada))) {
              return { status: 'erro', message: 'Você não tem permissão para acessar esta empresa.' };
            }

            // 3. Se a permissão estiver OK, busca os dados COMPLETOS da empresa selecionada.
            //    Isso usa a função auxiliar que já discutimos.
            const empresaObjetoCompleto = _getEmpresaDataById(empresaSelecionada);

            // 4. Se não encontrar os dados da empresa (ex: ID não existe na planilha 'Empresas'), retorna um erro.
            if (!empresaObjetoCompleto) {
              return { status: 'erro', message: `Os dados para a empresa ID ${empresaSelecionada} não foram encontrados.` };
            }
            return {
              status: 'ok',
              idUsuario: id,           // <-- Incluído aqui
              nomeUsuario: nome,
              nome: nome,       // <-- Incluído aqui
              perfil: perfil,
              empresa: empresaObjetoCompleto
            };
          } else {
            return { status: 'inativo', message: 'Usuário aguardando aprovação do administrador ou inativo' };
          }
        }
      }
      return { status: 'erro', message: 'Usuário ou senha inválidos!' };
    }

    // Crie esta função auxiliar no seu backend
    function _getEmpresaDataById(id) {
        // Esta função deve ler sua planilha de 'Empresas'
        // e retornar o objeto da empresa (com id, nome, cnpj, endereco)
        // que corresponde ao ID fornecido.
      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Empresas");
        if (!sheet) return null;

        const values = sheet.getDataRange().getValues();
        const headers = values[0];
        const idColumnIndex = headers.findIndex(h => String(h).toUpperCase() === 'ID');

        if (idColumnIndex === -1) return null;

        for (let i = 1; i < values.length; i++) {
          const idNaPlanilha = String(values[i][idColumnIndex]).trim();
          const idProcurado = String(id).trim();

          // Compara os IDs como números para ignorar a diferença de zeros
          if (parseInt(idNaPlanilha, 10) === parseInt(idProcurado, 10)) {
              
              const empresa = {};
              headers.forEach((header, index) => {
                  const chave = toCamelCase(header);
                  let valor = values[i][index];

                  
                  // Se esta é a coluna do ID, formata o valor para ser uma string
                  // com 3 dígitos, preenchendo com zeros à esquerda se necessário.
                  if (index === idColumnIndex) {
                      valor = String(valor).padStart(3, '0');
                  }
                  // --- FIM DA CORREÇÃO ---

                  empresa[chave] = valor;
              });
              return empresa; // Retorna o objeto da empresa com o ID já formatado
          }
        }
        return null;
      } catch(e) {
        Logger.log("Erro em _getEmpresaDataById: " + e.message);
        return null;
      }
    }

    // ===============================================
    // FUNÇÕES PARA USUARIOS
    // ===============================================

    /**
     * Cria um novo usuário com empresas permitidas.
     * @param {string} nome Nome do usuário.
     * @param {string} usuario Login do usuário.
     * @param {string} senha Senha do usuário.
     * @param {string} empresasCódigos Códigos das empresas (ex: "1,2,3").
     * @param {string} [perfil] Perfil do usuário ("usuario" ou "admin"). Opcional.
     * @returns {Object} Objeto status/mensagem.
     */
    function criarUsuario(nome, usuario, senha, empresasCodigos, perfil) {

      // Validações iniciais
      if (!nome || !senha) {
        return { status: 'error', message: 'Nome e senha são obrigatórios.' };
      }
      
      // Verifica se a senha tem pelo menos 6 caracteres
      if (!senha || senha.length < 6) {
        return { status: 'error', message: 'A senha deve ter pelo menos 6 caracteres.' };
      }

      try {
        const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
        
        // Pega todos os usuários existentes para a verificação de duplicidade
        const dadosUsuariosExistentes = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues().flat();
        
        // GERA O NOME DE USUÁRIO ÚNICO AQUI
        const novoUsuario = _gerarUsernameUnico(nome, dadosUsuariosExistentes);
        //const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
        //const dadosUsuariosExistentes = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues().flat();
        if (dadosUsuariosExistentes.includes(usuario)) {
          return { status: 'error', message: 'Nome de usuário já existe. Escolha outro.' };
        }

      const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(id => parseInt(id)).filter(n => !isNaN(n));
      const novoId = ids.length ? Math.max(...ids) + 1 : 1;

      // GERA O HASH DA SENHA ANTES DE SALVAR
      const senhaHash = gerarHash(senha);

      sheet.appendRow([
        novoId,
        nome,
        usuario,
        senhaHash,
        perfil || 'usuario', // Perfil padrão ou informado
        'Inativo',           // Status padrão
        empresasCodigos      // Novidade: códigos das empresas permitidas, ex: "1,2"
      ]);
      return { status: 'ok', message: `Solicitação para o usuário '${novoUsuario}' enviada. Aguarde ativação pelo Administrador.` };

      } catch (e) {
        Logger.log("Erro em criarUsuario: " + e.message);
        return { status: 'error', message: 'Ocorreu um erro ao criar a solicitação.' };
      }
    }

    /**
     * Gera um nome de usuário único no formato "primeiro.ultimo".
     * Se o nome de usuário já existir, adiciona um número ao final (ex: joao.silva2).
     * @param {string} nomeCompleto O nome completo do usuário.
     * @param {Array<string>} usuariosExistentes Uma lista de todos os nomes de usuário já cadastrados.
     * @returns {string} Um nome de usuário único.
     */
    function _gerarUsernameUnico(nomeCompleto, usuariosExistentes) {
      if (!nomeCompleto) return '';

      const nomes = nomeCompleto.trim().toLowerCase().split(' ');
      const primeiroNome = nomes[0];
      const ultimoNome = nomes.length > 1 ? nomes[nomes.length - 1] : '';

      let usernameBase = ultimoNome ? `${primeiroNome}.${ultimoNome}` : primeiroNome;
      
      // Normaliza o nome de usuário para remover acentos e caracteres especiais
      usernameBase = usernameBase.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

      // Verifica se o nome de usuário já existe e adiciona um sufixo numérico se necessário
      let finalUsername = usernameBase;
      let counter = 2;
      while (usuariosExistentes.map(u => u.toLowerCase()).includes(finalUsername)) {
        finalUsername = `${usernameBase}${counter}`;
        counter++;
      }
      
      return finalUsername;
    }

    /**
     * Lista todos os usuários da planilha 'Usuarios'.
     * @returns {Array<Object>} Uma lista de objetos de usuário.
     */
    function listarUsuarios() {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
      if (!sheet || sheet.getLastRow() < 2) return [];

      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();

      return data.map(([id, nome, usuario, senha, perfil, status, empresasStr]) => ({
        id: String(id),
        nome: String(nome),
        usuario: String(usuario),
        perfil: String(perfil),
        status: String(status),
        empresas: empresasStr ? String(empresasStr).split(',').map(e => e.trim()) : []
      }));
    }

    function carregarEmpresasPorUsuario(userId, empresasSelecionadas) {
      google.script.run.withSuccessHandler(function(empresas) {
        const container = document.getElementById(`empresas-${userId}`);
        if (!container) return;

        container.innerHTML = empresas.map(emp => {
          const codigo = String(emp.codigo).trim();
          const checked = selecionadas.includes(codigo) ? 'checked' : '';
          return `<label style="display:block;">
                    <input type="checkbox" value="${emp.codigo}" ${checked}>
                    ${emp.nome}
                  </label>`;
        }).join('');
      }).listarEmpresas();
    }

    function salvarPermissoesUsuario(userId, listaDeIdsPermitidos, idEmpresaPadrao) {
      try {
        const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
        if (!sheet) throw new Error("Planilha 'Usuarios' não encontrada.");

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const indexId = headers.findIndex(h => h.toUpperCase() === 'ID');
        const indexEmpresas = headers.findIndex(h => h.toUpperCase() === 'ID EMPRESA');
        const indexEmpresaPadrao = headers.findIndex(h => h.toUpperCase() === 'ID EMPRESA PADRÃO');

        if (indexId === -1 || indexEmpresas === -1 || indexEmpresaPadrao === -1) {
          throw new Error("Colunas 'ID', 'ID EMPRESA' ou 'ID EMPRESA PADRÃO' não encontradas.");
        }
        
        const idsUsuarios = sheet.getRange(2, indexId + 1, sheet.getLastRow() - 1, 1).getValues().flat();
        const rowIndexToUpdate = idsUsuarios.findIndex(id => String(id) == String(userId)) + 2;

        if (rowIndexToUpdate < 2) {
          return { status: 'error', message: `Usuário com ID ${userId} não encontrado.` };
        }

        const empresasString = Array.isArray(listaDeIdsPermitidos) ? listaDeIdsPermitidos.join(',') : '';
        
        sheet.getRange(rowIndexToUpdate, indexEmpresas + 1).setValue(empresasString);
        sheet.getRange(rowIndexToUpdate, indexEmpresaPadrao + 1).setValue("'" + (idEmpresaPadrao || ''));

        return { status: 'ok', message: 'Permissões atualizadas com sucesso!' };

      } catch (e) {
        Logger.log(`Erro em salvarPermissoesUsuario: ${e.message}`);
        return { status: 'error', message: `Erro no servidor: ${e.message}` };
      }
    }

    function buscarEmpresasDoUsuario() {
      const usuario = obterUsuarioLogado(); // Sua função de controle de sessão

      const sheet = SpreadsheetApp.getActive().getSheetByName('Empresas');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const colId = headers.indexOf('ID');
      const colNome = headers.indexOf('NOME');

      const empresasPermitidas = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // Aqui você pode adicionar o filtro conforme o usuário logado
        empresasPermitidas.push({
          idEmpresa: row[colId],
          nomeEmpresa: row[colNome]
        });
      }

      return empresasPermitidas;
    }

    function registrarEmpresaSelecionadaNoLogin(empresaId) {
      const sheet = SpreadsheetApp.getActive().getSheetByName('USUARIO_LOGADO');
      sheet.getRange(2, 1).setValue(empresaId); // Exemplo simples, adapte conforme seu controle
    }

    function obterUsuarioLogado() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('usuario_logado');
      if (!sheet) throw new Error("Planilha 'usuario_logado' não encontrada.");

      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return null; // Sem dados

      const headers = data[0];
      const idxIdUsuario = headers.indexOf("ID_USUARIO");
      const idxNomeUsuario = headers.indexOf("NOME_USUARIO");
      const idxIdEmpresa = headers.indexOf("ID_EMPRESA");
      const idxNomeEmpresa = headers.indexOf("NOME_EMPRESA");
      const idxDataLogin = headers.indexOf("DATA_LOGIN");

      if ([idxIdUsuario, idxNomeUsuario, idxIdEmpresa, idxNomeEmpresa, idxDataLogin].includes(-1)) {
        throw new Error("Colunas necessárias não encontradas na planilha 'usuario_logado'.");
      }

      // Encontra a linha com a DATA_LOGIN mais recente
      let ultimoRegistro = null;
      let dataMaisRecente = new Date(0); // data mínima para iniciar

      for (let i = 1; i < data.length; i++) {
        const linha = data[i];
        const dataLogin = new Date(linha[idxDataLogin]);
        if (dataLogin > dataMaisRecente) {
          dataMaisRecente = dataLogin;
          ultimoRegistro = linha;
        }
      }

      if (!ultimoRegistro) return null;

      return {
        idUsuario: ultimoRegistro[idxIdUsuario],
        nomeUsuario: ultimoRegistro[idxNomeUsuario],
        idEmpresa: ultimoRegistro[idxIdEmpresa],
        nomeEmpresa: ultimoRegistro[idxNomeEmpresa],
        dataLogin: dataMaisRecente
      };
    }

    function registrarLoginUsuario(idUsuario, nomeUsuario, idEmpresa, nomeEmpresa) {
      const ss = SpreadsheetApp.getActive();
      const sheet = ss.getSheetByName('usuario_logado');
      
      if (!sheet) {
        throw new Error('Planilha usuario_logado não encontrada.');
      }
      
      const dataHoraLogin = new Date();
      // Força os IDs a serem tratados como texto pela planilha adicionando um apóstrofo.
      // Isso previne que "001" se torne 1.
      const idUsuarioTexto = "'" + idUsuario;
      const idEmpresaTexto = "'" + idEmpresa;
      const dados = sheet.getDataRange().getValues();
      // Supondo que a primeira linha seja cabeçalho
      let linhaExistente = -1;
      
      for (let i = 1; i < dados.length; i++) {
        if (dados[i][0] == idUsuario) {  // Coluna 0 = ID_USUARIO
          linhaExistente = i + 1; // Índice da linha real na planilha (1-based)
          break;
        }
      }
      
      if (linhaExistente > 0) {
        // Atualiza registro existente
        sheet.getRange(linhaExistente, 1).setValue(idUsuarioTexto);   // ID_USUARIO
        sheet.getRange(linhaExistente, 2).setValue(nomeUsuario);   // NOME_USUARIO
        sheet.getRange(linhaExistente, 3).setValue(idEmpresaTexto);     // ID_EMPRESA
        sheet.getRange(linhaExistente, 4).setValue(nomeEmpresa);   // NOME_EMPRESA
        sheet.getRange(linhaExistente, 5).setValue(dataHoraLogin); // DATA_LOGIN
      } else {
        // Insere novo registro no final
        sheet.appendRow([idUsuarioTexto, nomeUsuario, idEmpresaTexto, nomeEmpresa, dataHoraLogin]);
      }
      
      return { status: 'ok', message: 'Usuário logado registrado com sucesso.' };
    }

    function obterEmpresasDoUsuario(username) {
      try {
        if (!username) return null;

        const sheetUsuarios = SpreadsheetApp.getActive().getSheetByName('Usuarios');
        if (!sheetUsuarios || sheetUsuarios.getLastRow() < 2) return null;

        const dataUsuarios = sheetUsuarios.getDataRange().getValues();
        const headersUsuarios = dataUsuarios[0].map(h => String(h).toUpperCase());

        const indexUser = headersUsuarios.indexOf('USUARIO');
        const indexStatus = headersUsuarios.indexOf('STATUS');
        const indexEmpresasStr = headersUsuarios.indexOf('ID EMPRESA');
        const indexEmpresaPadrao = headersUsuarios.indexOf('ID EMPRESA PADRÃO');

        if (indexUser === -1 || indexStatus === -1 || indexEmpresasStr === -1) {
          throw new Error("Cabeçalhos 'USUARIO', 'STATUS' ou 'ID EMPRESA' não encontrados na planilha 'Usuarios'.");
        }

        const usuarioEncontrado = dataUsuarios.find(row => String(row[indexUser]).trim().toLowerCase() === username.trim().toLowerCase());

        if (!usuarioEncontrado || usuarioEncontrado[indexStatus].toUpperCase() !== 'ATIVO') {
          return null;
        }
        
        const idEmpresaPadrao = (indexEmpresaPadrao > -1) ? usuarioEncontrado[indexEmpresaPadrao] : null;
        const idsEmpresasString = String(usuarioEncontrado[indexEmpresasStr] || '').trim();

        if (!idsEmpresasString) {
          return { defaultEmpresaId: null, empresas: [] };
        }

        // Padroniza a lista de IDs permitidos para o usuário (ex: '1' vira '001')
        const idsPermitidosPadronizados = idsEmpresasString.split(',').map(id => id.trim().padStart(3, '0'));
      
        const sheetEmpresas = SpreadsheetApp.getActive().getSheetByName('Empresas');
        if (!sheetEmpresas) return { defaultEmpresaId: idEmpresaPadrao, empresas: [] };
        const dadosEmpresas = sheetEmpresas.getRange(2, 1, sheetEmpresas.getLastRow() - 1, 2).getValues();
        const empresasPermitidas = dadosEmpresas
          // Compara a lista padronizada com os IDs da planilha de empresas, também padronizados
          .filter(empresaRow => idsPermitidosPadronizados.includes(String(empresaRow[0]).trim().padStart(3, '0')))
          .map(empresaRow => ({
            id: String(empresaRow[0]).trim().padStart(3, '0'),
            nome: String(empresaRow[1]).trim()
          }));

        return {
          defaultEmpresaId: idEmpresaPadrao ? String(idEmpresaPadrao).trim().padStart(3, '0') : null,
          empresas: empresasPermitidas
        };

      } catch (e) {
        Logger.log('Erro em obterEmpresasDoUsuario: ' + e.message);
        return { error: e.message };
      }
    }

    /**
     * Altera o status de um usuário na planilha 'Usuarios'.
     * @param {string} userId - O ID do usuário.
     * @param {string} novoStatus - O novo status (ex: 'Ativo', 'Inativo').
     * @returns {Object} Um objeto com status e mensagem.
     */
    function alterarStatusUsuario(userId, novoStatus) {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
      if (!sheet) {
        Logger.log('[alterarStatusUsuario - SERVER] Planilha "Usuarios" não encontrada.');
        return { status: 'error', message: 'Planilha "Usuarios" não encontrada.' };
      }

      Logger.log(`[alterarStatusUsuario - SERVER] Tentando alterar status para userId: ${userId}, novoStatus: ${novoStatus}`);
      
      // Busca o usuário pelo ID na primeira coluna (coluna 1)
      const lastRow = sheet.getLastRow();
      const idsColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Obtém apenas a coluna de IDs
      let rowIndexToUpdate = -1;

      for (let i = 0; i < idsColumn.length; i++) {
        if (String(idsColumn[i][0]).trim() === String(userId).trim()) {
          rowIndexToUpdate = i + 2; // +2 porque os dados começam na linha 2 e o índice do array é 0-based
          break;
        }
      }

      if (rowIndexToUpdate !== -1) {
        // A coluna de status é a 6ª coluna (índice 5 no array getValues)
        sheet.getRange(rowIndexToUpdate, 6).setValue(novoStatus);
        Logger.log(`[alterarStatusUsuario - SERVER] Status do usuário ${userId} alterado para ${novoStatus} na linha ${rowIndexToUpdate}.`);
        return { status: 'ok', message: `Status do usuário ${userId} atualizado para ${novoStatus}.` };
      } else {
        Logger.log(`[alterarStatusUsuario - SERVER] Usuário ${userId} não encontrado para alteração de status.`);
        return { status: 'error', message: `Usuário ${userId} não encontrado.` };
      }
    }

    /**
     * Exclui um usuário da planilha 'Usuarios'.
     * @param {string} userId - O ID do usuário a ser excluído.
     * @returns {Object} Um objeto com status e mensagem.
     */
    function excluirUsuario(userId) {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
      if (!sheet) {
        Logger.log('[excluirUsuario - SERVER] Planilha "Usuarios" não encontrada.');
        return { status: 'error', message: 'Planilha "Usuarios" não encontrada.' };
      }

      Logger.log(`[excluirUsuario - SERVER] Tentando excluir userId: ${userId}`);

      const lastRow = sheet.getLastRow();
      const idsColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Obtém apenas a coluna de IDs
      let rowIndexToDelete = -1;

      for (let i = 0; i < idsColumn.length; i++) {
        if (String(idsColumn[i][0]).trim() === String(userId).trim()) {
          rowIndexToDelete = i + 2; // +2 porque os dados começam na linha 2 e o índice do array é 0-based
          break;
        }
      }

      if (rowIndexToDelete !== -1) {
        sheet.deleteRow(rowIndexToDelete);
        Logger.log(`[excluirUsuario - SERVER] Usuário ${userId} excluído da linha ${rowIndexToDelete}.`);
        return { status: 'ok', message: `Usuário ${userId} excluído com sucesso.` };
      } else {
        Logger.log(`[excluirUsuario - SERVER] Usuário ${userId} não encontrado para exclusão.`);
        return { status: 'error', message: `Usuário ${userId} não encontrado.` };
      }
    }

    /**
     * Altera o perfil de um usuário na planilha 'Usuarios'.
     * @param {string} userId - O ID do usuário.
     * @param {string} novoPerfil - O novo perfil (ex: 'admin', 'usuario').
     * @returns {Object} Um objeto com status e mensagem.
     */
    function alterarPerfilUsuario(userId, novoPerfil) {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
      if (!sheet) {
        return { status: 'error', message: 'Planilha "Usuarios" não encontrada.' };
      }
      const lastRow = sheet.getLastRow();
      const idsColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      let rowIndexToUpdate = -1;

      for (let i = 0; i < idsColumn.length; i++) {
        if (String(idsColumn[i][0]).trim() === String(userId).trim()) {
          rowIndexToUpdate = i + 2;
          break;
        }
      }
      if (rowIndexToUpdate !== -1) {
        // A coluna de perfil é a 5ª coluna
        sheet.getRange(rowIndexToUpdate, 5).setValue(novoPerfil);
        return { status: 'ok', message: `Perfil do usuário ${userId} alterado para ${novoPerfil}.` };
      } else {
        return { status: 'error', message: `Usuário ${userId} não encontrado.` };
      }
    }

    /**
     * Salva as permissões de um usuário e define a primeira empresa da lista como padrão.
     * @param {string} idUsuario - O ID do usuário a ser atualizado.
     * @param {string} empresasCodigosStr - A string com os códigos das empresas, separados por vírgula (ex: "003,001,002").
     */
    function salvarPermissoesDeEmpresaParaUsuario(idUsuario, empresasCodigosStr) {
      Logger.log(`--- [DIAGNÓSTICO SALVAR PERMISSÕES] ---`);
      Logger.log(`1. Função iniciada. ID do Usuário: "${idUsuario}", String de Empresas: "${empresasCodigosStr}"`);
      
      try {
        const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
        if (!sheet) throw new Error("Planilha 'Usuarios' não encontrada.");

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        Logger.log(`2. Cabeçalhos lidos da planilha 'Usuarios': [${headers.join(', ')}]`);

        const headersUpperCase = headers.map(h => String(h).toUpperCase());
        
        // Encontra os índices das colunas
        const indexIdUsuario = headersUpperCase.indexOf('ID');
        const indexEmpresasPermitidas = headersUpperCase.indexOf('ID EMPRESA');
        const indexEmpresaPadrao = headersUpperCase.indexOf('ID EMPRESA PADRÃO');
        
        Logger.log(`3. Índices encontrados -> ID: ${indexIdUsuario}, ID EMPRESA: ${indexEmpresasPermitidas}, ID EMPRESA PADRÃO: ${indexEmpresaPadrao}`);

        if (indexIdUsuario === -1 || indexEmpresasPermitidas === -1 || indexEmpresaPadrao === -1) {
          throw new Error("Uma ou mais colunas necessárias (ID, ID EMPRESA, ID EMPRESA PADRÃO) não foram encontradas. Verifique os nomes exatos na planilha.");
        }

        // Encontra a linha do usuário
        const ids = sheet.getRange(2, indexIdUsuario + 1, sheet.getLastRow() - 1, 1).getValues().flat();
        const rowIndexToUpdate = ids.findIndex(id => String(id) == String(idUsuario)) + 2;
        Logger.log(`4. Procurando pelo ID de usuário "${idUsuario}". Linha encontrada: ${rowIndexToUpdate > 1 ? rowIndexToUpdate : 'NENHUMA'}`);

        if (rowIndexToUpdate < 2) {
          return { status: 'error', message: `Usuário com ID ${idUsuario} não encontrado.` };
        }

        // Lógica para definir a empresa padrão
        const codigosArray = String(empresasCodigosStr || '').split(',').map(c => c.trim()).filter(String);
        const idEmpresaPadrao = (codigosArray.length > 0) ? codigosArray[0] : '';
        Logger.log(`5. Lógica da empresa padrão -> Primeiro ID da lista é: "${idEmpresaPadrao}"`);

        Logger.log(`6. TENTANDO ESCREVER NA PLANILHA...`);
        Logger.log(`   - Linha: ${rowIndexToUpdate}`);
        Logger.log(`   - Coluna de Permissões (índice ${indexEmpresasPermitidas}): Escrevendo o valor "${empresasCodigosStr}"`);
        sheet.getRange(rowIndexToUpdate, indexEmpresasPermitidas + 1).setValue(empresasCodigosStr);
        
        Logger.log(`   - Coluna Padrão (índice ${indexEmpresaPadrao}): Escrevendo o valor "'${idEmpresaPadrao}"`);
        sheet.getRange(rowIndexToUpdate, indexEmpresaPadrao + 1).setValue("'" + idEmpresaPadrao);
        
        Logger.log(`7. Escrita na planilha concluída com sucesso.`);
        return { status: 'ok', message: 'Permissões e empresa padrão atualizadas com sucesso!' };

      } catch (e) {
        Logger.log(`--- [DIAGNÓSTICO SALVAR PERMISSÕES] ERRO FATAL: ${e.message} ---`);
        return { status: 'error', message: `Erro no servidor: ${e.message}` };
      }
    }


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

    // ===============================================
    // FUNÇÕES PARA IMPRESSAO
    // ===============================================    

    /**
     * Busca todos os dados de um pedido específico para exibição na tela de impressão.
     * @param {string} numeroPedido - O número do pedido a ser buscado.
     * @returns {Object|null} Objeto com todos os dados do pedido, ou null se não encontrado.
     */
    function getDadosPedidoParaImpressao(numeroPedido) {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
      if (!sheet || sheet.getLastRow() < 2) {
        Logger.log(`Planilha 'Pedidos' vazia ou não encontrada ao buscar pedido ${numeroPedido}.`);
        return null;
      }

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

      Logger.log(`[getDadosPedidoParaImpressao] Buscando pedido: "${numeroPedido}"`);

      const pedidoRow = dados.find(row => {
        const sheetNumeroPedido = String(row[0]).trim(); // Pega o valor da primeira coluna e remove espaços
        Logger.log(`[getDadosPedidoParaImpressao] Comparando "${sheetNumeroPedido}" (na planilha) com "${String(numeroPedido).trim()}" (recebido).`);
        return sheetNumeroPedido === String(numeroPedido).trim();
      });
      
      if (!pedidoRow) {
        Logger.log(`Pedido "${numeroPedido}" não encontrado na planilha após a busca.`);
        return null;
      }

      const pedidoData = {};
      headers.forEach((header, index) => {
        const camelCaseHeader = toCamelCase(header); // Usa a função toCamelCase para padronizar
        let value = pedidoRow[index];

        // Converte objetos Date para string formatada (YYYY-MM-DD)
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        
        pedidoData[camelCaseHeader] = value;
      });

      try {
        pedidoData.itens = JSON.parse(pedidoData.itens || '[]');
      } catch (e) {
        Logger.log(`Erro ao parsear itens JSON para pedido ${numeroPedido}: ${e.message}`);
        pedidoData.itens = [];
      }

      Logger.log(`[getDadosPedidoParaImpressao] Pedido encontrado e processado: ${JSON.stringify(pedidoData)}`);
      return pedidoData;
    }

    function getDadosPedidoParaImpressaoAdmin(numeroPedido, empresaId) {
      console.log(`Buscando pedido ${numeroPedido} da empresa ${empresaId} para admin`);
      
      try {
        // Sua lógica para buscar o pedido específico da empresa
        // Exemplo (adapte para sua estrutura):
        
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
        var data = sheet.getDataRange().getValues();
        
        // Encontrar o pedido específico
        for (var i = 1; i < data.length; i++) {
          if (data[i][0] == numeroPedido && data[i][1] == empresaId) { // Ajuste os índices conforme sua planilha
            return {
              numeroDoPedido: data[i][0],
              empresaId: data[i][1],
              data: data[i][2],
              fornecedor: data[i][3],
              totalGeral: data[i][4],
              // ... outros campos
              // Incluir dados da empresa também
              enderecoEmpresa: "endereço da empresa",
              cnpjEmpresa: "cnpj da empresa",
              // etc.
            };
          }
        }
        
        return null; // Pedido não encontrado
        
      } catch (error) {
        console.error("Erro ao buscar pedido para admin:", error);
        throw error;
      }
    }

    /**
     * Busca pedidos na planilha 'Pedidos' com base em um termo de busca.
     * O termo de busca pode ser o número do pedido ou o nome do fornecedor.
     * @param {string} termoBusca - O termo a ser buscado.
     * @param {string} empresaCodigo - O código da empresa (ex: "E001").
     * @returns {Object} Um objeto com status e uma lista de pedidos que correspondem ao termo de busca.
     */
    /**
     * VERSÃO DE DIAGNÓSTICO
     * Busca pedidos com base em um termo e no código da empresa, com logs detalhados.
     */
    function buscarPedidos(termoBusca, empresaCodigo) {
      console.log('🔍 === INÍCIO buscarPedidos ===');
      console.log('🔍 Termo de busca:', termoBusca);
      console.log('🔍 Código da empresa:', empresaCodigo);
      
      try {
        const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
        if (!sheet) {
          console.error('❌ Planilha "Pedidos" não encontrada');
          return { status: "error", data: [], message: "Aba 'Pedidos' não encontrada." };
        }

        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        console.log('📋 Headers da planilha:', headers);
        
        const resultados = [];
        
        const idEmpresaFiltro = String(empresaCodigo).trim();
        const termoNormalizado = (termoBusca || "").toString().toLowerCase().trim();
        console.log('🔍 ID empresa filtro:', idEmpresaFiltro);
        console.log('🔍 Termo normalizado:', termoNormalizado);

        const colEmpresa = headers.findIndex(h => ["EMPRESA", "IDEMPRESA", "IDDAEMPRESA", "ID DA EMPRESA", "Empresa"].includes(String(h).toUpperCase()));
        const colNumeroPedido = headers.findIndex(h => String(h).toUpperCase() === 'NÚMERO DO PEDIDO');
        const colFornecedor = headers.findIndex(h => String(h).toUpperCase() === 'FORNECEDOR');
        
        console.log('📍 Índices das colunas:');
        console.log('  - Empresa:', colEmpresa, headers[colEmpresa]);
        console.log('  - Número do Pedido:', colNumeroPedido, headers[colNumeroPedido]);
        console.log('  - Fornecedor:', colFornecedor, headers[colFornecedor]);

        if (colEmpresa === -1) {
          console.error('❌ Coluna da empresa não encontrada');
          return { status: "error", data: [], message: "Coluna da empresa não encontrada." };
        }

        console.log(`🔍 Total de linhas para processar: ${data.length - 1}`);
        
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const idEmpresaNaLinha = String(row[colEmpresa]).trim();

          console.log(`📋 Linha ${i}: ID empresa na linha: "${idEmpresaNaLinha}", Filtro: "${idEmpresaFiltro}"`);

          if (parseInt(idEmpresaNaLinha, 10) !== parseInt(idEmpresaFiltro, 10)) {
            console.log(`⏭️ Linha ${i}: Empresa não corresponde - pulando`);
            continue;
          }

          console.log(`✅ Linha ${i}: Empresa corresponde - verificando termo de busca`);

          const numeroPedidoNormalizado = String(row[colNumeroPedido] || '').toLowerCase().trim();
          const fornecedorNormalizado = String(row[colFornecedor] || '').toLowerCase().trim();

          console.log(`🔍 Linha ${i}: Número: "${numeroPedidoNormalizado}", Fornecedor: "${fornecedorNormalizado}"`);

          const shouldAddRow = (termoNormalizado === "") || 
                              numeroPedidoNormalizado.includes(termoNormalizado) || 
                              fornecedorNormalizado.includes(termoNormalizado);

          console.log(`🔍 Linha ${i}: Deve adicionar? ${shouldAddRow}`);
                              
          if (shouldAddRow) {
            const pedidoData = {};
            headers.forEach((header, index) => {
              let value = row[index];

              // --- AQUI ESTÁ A CORREÇÃO FINAL ---
              // Se o valor for um objeto de Data, converte para texto no formato AAAA-MM-DD
              if (value instanceof Date) {
                value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
              }
              // --- FIM DA CORREÇÃO ---

              pedidoData[toCamelCase(header)] = value;
            });

            // O resto da sua lógica de parse e etc.
            pedidoData.totalGeral = parseFloat(pedidoData.totalGeral || 0);
            if (typeof pedidoData.itens === 'string') {
              try {
                pedidoData.itens = JSON.parse(pedidoData.itens);
              } catch (e) {
                pedidoData.itens = [];
              }
            } else if (!Array.isArray(pedidoData.itens)) {
              pedidoData.itens = [];
            }

            console.log(`➕ Linha ${i}: Adicionando pedido aos resultados`);
            resultados.push(pedidoData);
          }
        }
        
        console.log(`📊 Total de resultados encontrados: ${resultados.length}`);
        console.log('🔍 === FIM buscarPedidos ===');
        
        // Agora o objeto 'resultados' não contém mais objetos de Data e pode ser retornado
        return { status: "success", data: resultados };

      } catch (e) {
        Logger.log(`Erro fatal em buscarPedidos: ${e.message}`);
        return { status: "error", data: [], message: `Erro no servidor: ${e.message}` };
      }
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
    // FUNÇÕES PARA DASHBOARD (NOVAS)
    // ===============================================

    /**
     * Função auxiliar para obter dados de pedidos de forma padronizada.
     * @returns {Array<Object>} Lista de objetos de pedidos com cabeçalhos em camelCase.
     */
    function _getPedidosData(empresa) {
      Logger.log(`[getPedidosData v5] Função iniciada. Filtro: ${JSON.stringify(empresa)}`);
      try {
        const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
        if (!sheet || sheet.getLastRow() < 2) {
          Logger.log('[getPedidosData v5] Planilha "Pedidos" vazia ou não encontrada.');
          return [];
        }

        const values = sheet.getDataRange().getValues();
        const headers = values[0];
        const indexEmpresa = headers.findIndex(h => ["EMPRESA", "IDEMPRESA", "IDDAEMPRESA"].includes(h.toUpperCase()));

        if (indexEmpresa === -1) {
          Logger.log('[getPedidosData v5] ERRO CRÍTICO: Coluna da empresa não encontrada.');
          return [];
        }
        
        const chaveEmpresa = toCamelCase(headers[indexEmpresa]);

        const pedidos = values.slice(1).map(row => {
          const pedido = {};
          headers.forEach((header, index) => {
            pedido[toCamelCase(header)] = row[index];
          });

          
          // Converte a string de itens em um array de objetos
          if (typeof pedido.itens === 'string' && pedido.itens.trim().startsWith('[')) {
            try {
              pedido.itens = JSON.parse(pedido.itens);
            } catch (e) {
              Logger.log(`Erro ao converter itens para o pedido ${pedido.numeroDoPedido || 'N/A'}: ${e.message}`);
              pedido.itens = []; // Garante que seja um array vazio em caso de erro
            }
          } else if (!Array.isArray(pedido.itens)) {
            // Se não for uma string JSON ou um array, força a ser um array vazio
            pedido.itens = [];
          }
          

          pedido.totalGeral = parseFloat(pedido.totalGeral || 0);
          return pedido;
        });

        if (empresa && empresa.id != null) {
          const idEmpresa = String(empresa.id).trim();
          const filtrados = pedidos.filter(p => {
            const pedidoEmpresaValor = String(p[chaveEmpresa] || '').trim();
            return pedidoEmpresaValor == idEmpresa;
          });
          Logger.log(`[getPedidosData v5] Filtro aplicado. Encontrados ${filtrados.length} de ${pedidos.length} pedidos.`);
          return filtrados;
        }
        
        return pedidos;

      } catch (e) {
        Logger.log(`[getPedidosData v5] ERRO FATAL: ${e.message}`);
        return [];
      }
    }

    function sanitizePedidos(pedidos) {
      return pedidos.map(pedido => {
        const clone = { ...pedido };
        // Transforma datas em string padronizada
        if (clone.data instanceof Date) {
          clone.data = formatarDataParaISO(clone.data);
        }
        // Se tiver outros campos Date, serialize-os aqui!
        // Remova funções e campos não serializáveis, se houver
        return clone;
      });
    }

    /**
     * Retorna os fornecedores mais comprados, do maior para o menor volume.
     * @param {Array<Object>} pedidos - A lista de pedidos JÁ FILTRADA.
     * @returns {Array<Object>} Lista de fornecedores com total comprado.
     */
    function _getTopSuppliersByVolume(empresa) {
      const pedidos = _getPedidosData(empresa);
      const supplierVolumes = {};

      pedidos.forEach(pedido => {
        const fornecedor = pedido.fornecedor || 'Desconhecido';
        const total = pedido.totalGeral || 0;
        supplierVolumes[fornecedor] = (supplierVolumes[fornecedor] || 0) + total;
      });

      const sortedSuppliers = Object.keys(supplierVolumes).map(fornecedor => ({
        fornecedor: fornecedor,
        totalComprado: supplierVolumes[fornecedor]
      })).sort((a, b) => b.totalComprado - a.totalComprado);

      Logger.log(`[Dashboard - _getTopSuppliersByVolume] Fornecedores por volume: ${JSON.stringify(sortedSuppliers)}`);
      return sortedSuppliers;
    }

    /**
     * Retorna os produtos mais pedidos, do maior para o menor volume (quantidade).
     * @param {Array<Object>} pedidos - A lista de pedidos JÁ FILTRADA.
     * @returns {Array<Object>} Lista de produtos com quantidade total pedida.
     */
    function _getTopProductsByVolume(empresa) {
      const pedidos = _getPedidosData(empresa);
      const productQuantities = {};

      pedidos.forEach(pedido => {
        (pedido.itens || []).forEach(item => {
          const produto = item.descricao || 'Produto Desconhecido';
          const quantidade = parseFloat(item.quantidade) || 0;
          productQuantities[produto] = (productQuantities[produto] || 0) + quantidade;
        });
      });

      const sortedProducts = Object.keys(productQuantities).map(produto => ({
        produto: produto,
        totalQuantidade: productQuantities[produto]
      })).sort((a, b) => b.totalQuantidade - a.totalQuantidade);

      Logger.log(`[Dashboard - _getTopProductsByVolume] Produtos por volume: ${JSON.stringify(sortedProducts)}`);
      return sortedProducts;
    }

    /**
     * Retorna um resumo financeiro dos pedidos.
     * @param {Array<Object>} pedidos - A lista de pedidos JÁ FILTRADA.
     * @returns {Object} Resumo financeiro.
     */
    function _getFinancialSummary(empresa) {
      const pedidos = _getPedidosData(empresa);
      let totalGeralPedidos = 0;
      let numeroTotalPedidos = 0;

      pedidos.forEach(pedido => {
        totalGeralPedidos += pedido.totalGeral || 0;
        numeroTotalPedidos++;
      });

      const summary = {
        totalPedidos: numeroTotalPedidos,
        valorTotalPedidos: totalGeralPedidos
      };

      Logger.log(`[Dashboard - _getFinancialSummary] Resumo financeiro: ${JSON.stringify(summary)}`);
      return summary;
    }

    /**
     * Gera sugestões de redução de compras usando o Gemini API.
     * @param {string} financialSummaryStr - Resumo financeiro em string JSON.
     * @param {string} topProductsStr - Lista de top produtos em string JSON.
     * @param {string} topSuppliersStr - Lista de top fornecedores em string JSON.
     * @returns {string} Sugestões de redução de compras.
     */
    function generatePurchaseSuggestions(financialSummaryStr, topProductsStr, topSuppliersStr) {
      Logger.log('[generatePurchaseSuggestions] Iniciando geração de sugestões...');
      Logger.log('financialSummaryStr: ' + financialSummaryStr);
      Logger.log('topProductsStr: ' + topProductsStr);
      Logger.log('topSuppliersStr: ' + topSuppliersStr);
      // Corrige: se algum parâmetro não vier, substitui por um valor válido
      if (!financialSummaryStr || financialSummaryStr === "undefined") financialSummaryStr = "{}";
      if (!topProductsStr || topProductsStr === "undefined") topProductsStr = "[]";
      if (!topSuppliersStr || topSuppliersStr === "undefined") topSuppliersStr = "[]";

      try {
        const financialSummary = JSON.parse(financialSummaryStr);
        const topProducts = JSON.parse(topProductsStr);
        const topSuppliers = JSON.parse(topSuppliersStr);

        const prompt = `
          **PERSONA:** Você é um consultor de supply chain especializado na otimização de compras de mercadorias.

          **CONTEXTO:** Os dados a seguir são de compras emergenciais ("apaga-incêndio") realizadas em fornecedores locais de alto custo. A empresa deseja criar um plano de ação para minimizar a necessidade dessas compras e reduzir a dependência desses fornecedores.

          **TAREFA:** Com base nos dados, forneça de 3 a 5 recomendações estratégicas e acionáveis. Organize suas sugestões nas seguintes categorias: GESTÃO DE ESTOQUE, ESTRATÉGIA DE FORNECEDORES e ANÁLISE DE PRODUTOS. Para cada sugestão, seja direto e indique um próximo passo prático.

          **DADOS DE ENTRADA:**
          - Resumo Financeiro das Compras Locais: ${JSON.stringify(financialSummary)}
          - Top 5 Produtos de Compra Urgente: ${JSON.stringify(topProducts.slice(0, 5))}
          - Top 5 Fornecedores Locais Mais Utilizados: ${JSON.stringify(topSuppliers.slice(0, 5))}

          **PLANO DE AÇÃO ESTRATÉGICO:**
          `;

        // Obtém a chave da API das propriedades do script
        const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
        
        if (!apiKey) {
            Logger.log('[generatePurchaseSuggestions] Erro: GEMINI_API_KEY não configurada nas Propriedades do Script.');
            return "Erro de configuração: Chave da API Gemini não encontrada. Por favor, configure a 'GEMINI_API_KEY' nas Propriedades do Script.";
        }

        const payload = {
          contents: [{ role: "user", parts: [{ text: prompt }] }],
          generationConfig: {
              temperature: 0.7, // Um pouco mais criativo, mas ainda focado
              maxOutputTokens: 500
          }
        };

        const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

        Logger.log(`[generatePurchaseSuggestions] Enviando requisição para Gemini API: ${apiUrl}`);
        Logger.log(`[generatePurchaseSuggestions] Payload (antes do fetch): ${JSON.stringify(payload)}`); 

        let response;
        try {
          response = UrlFetchApp.fetch(apiUrl, {
            method: 'POST',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true // Para capturar erros na resposta
          });
        } catch (fetchError) {
          Logger.log(`[generatePurchaseSuggestions] Erro ao executar UrlFetchApp.fetch: ${fetchError.message}. Stack: ${fetchError.stack}`);
          return `Erro de rede ou comunicação com a API Gemini: ${fetchError.message}`;
        }

        // Adicionado logs para inspecionar a resposta antes do JSON.parse
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        Logger.log(`[generatePurchaseSuggestions] Código de Resposta da API: ${responseCode}`); 
        Logger.log(`[generatePurchaseSuggestions] Texto de Resposta da API (bruto): ${responseText}`); 

        // ** Correção principal aqui: Verifica se responseText é válido antes de tentar JSON.parse **
        if (!responseText || responseText.trim() === '') {
          let errorMessage = `Erro: Resposta vazia ou nula da API Gemini. Código: ${responseCode}.`;
          Logger.log(`[generatePurchaseSuggestions] Erro detalhado (resposta vazia): ${errorMessage}`);
          return `Erro ao gerar sugestões: ${errorMessage}`;
        }

        let result;
        try {
          result = JSON.parse(responseText); // Agora processa responseText após a verificação de nulo/vazio
        } catch (parseError) {
          Logger.log(`[generatePurchaseSuggestions] SyntaxError: Erro ao parsear JSON da resposta da API Gemini: ${parseError.message}. Conteúdo bruto: ${responseText.substring(0, 500)}... Stack: ${parseError.stack}`);
          return `Erro ao gerar sugestões: Resposta inválida da API Gemini. Tente novamente ou verifique a chave/permissões.`;
        }
        
        Logger.log(`[generatePurchaseSuggestions] Resposta da API Gemini (parseada): ${JSON.stringify(result)}`);

        if (result.candidates && result.candidates.length > 0 &&
            result.candidates[0].content && result.candidates[0].content.parts &&
            result.candidates[0].content.parts.length > 0) {
          const suggestions = result.candidates[0].content.parts[0].text;
          Logger.log(`[generatePurchaseSuggestions] Sugestões geradas: ${suggestions}`);
          return suggestions;
        } else {
          Logger.log(`[generatePurchaseSuggestions] Resposta da API Gemini não contém sugestões válidas (candidates ausentes): ${JSON.stringify(result)}`);
          if (result.error && result.error.message) {
            return `Erro da API Gemini: ${result.error.message}`;
          }
          return "Não foi possível gerar sugestões no momento. A API retornou uma estrutura inesperada ou incompleta.";
        }

      } catch (e) {
        Logger.log(`[generatePurchaseSuggestions] Erro fatal (fora do bloco fetch/parse): ${e.message}. Stack: ${e.stack}`);
        return `Ocorreu um erro inesperado ao gerar sugestões: ${e.message}`;
      }
    }

    /**
     * Função principal para o Dashboard, que coleta todos os dados e gera sugestões.
     * @returns {Object} Objeto contendo todos os dados do dashboard e sugestões.
     */
    function getDashboardData(empresa) {
      // Este novo log confirma que a versão correta do código está executando.
      Logger.log('--- INICIANDO EXECUÇÃO COM O CÓDIGO CORRIGIDO (v3) ---'); 
      Logger.log(`Filtro recebido para a empresa: ${JSON.stringify(empresa)}`);
      
      try {
        // As chamadas abaixo agora passam o filtro 'empresa'
        const financialSummary = _getFinancialSummary(empresa);
        const topProducts = _getTopProductsByVolume(empresa);
        const topSuppliers = _getTopSuppliersByVolume(empresa);
        
        const suggestions = generatePurchaseSuggestions(
          JSON.stringify(financialSummary || {}),
          JSON.stringify(topProducts || []),
          JSON.stringify(topSuppliers || [])
        );

        Logger.log('--- SUCESSO! Retornando dados JÁ FILTRADOS. ---');
        
        return {
          status: 'success',
          topSuppliers: topSuppliers,
          topProducts: topProducts,
          financialSummary: financialSummary,
          suggestions: suggestions
        };
      } catch (e) {
        Logger.log(`ERRO FATAL: ${e.message}`);
        return { status: 'error', message: `Erro ao carregar dados do Dashboard: ${e.message}` };
      }
    }

    function serializePedidos(pedidos) {
      return pedidos.map(p => {
        const novoPedido = { ...p };
        // Transforma datas em string padronizada
        if (novoPedido.data instanceof Date) {
          novoPedido.data = formatarDataParaISO(novoPedido.data);
        }
        // Se tiver outros campos do tipo Date, faça igual aqui
        return novoPedido;
      });
    }

    // ===============================================
    // NOVAS FUNÇÕES PARA RELATÓRIOS
    // ===============================================

    /**
     * Retorna uma lista de todos os produtos únicos de todos os pedidos.
     * @returns {Array<string>} Uma lista de nomes de produtos únicos.
     */

    function _getPedidosDatas(reportParams) {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
      if (!sheet || sheet.getLastRow() < 2) {
        Logger.log('[_getPedidosData] Planilha "Pedidos" vazia ou não encontrada.');
        return [];
      }
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      const headers = values[0];
      const pedidos = [];

      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const pedido = {};
        headers.forEach((header, index) => {
          const camelCaseHeader = toCamelCase(header);
          let value = row[index];
          // Converte string de data para Date
          if (camelCaseHeader === 'data' && typeof value === 'string') {
            try {
              value = new Date(value + 'T12:00:00');
            } catch (e) {
              Logger.log(`Erro ao parsear data ${value} no pedido ${row[0]}: ${e.message}`);
              value = null;
            }
          } else if (value instanceof Date) {
            // ok
          } else if (camelCaseHeader === 'data' && !(value instanceof Date)) {
            value = null;
          }
          pedido[camelCaseHeader] = value;
        });

        // Parse itens JSON caso necessário
        if (typeof pedido.itens === 'string') {
          try {
            pedido.itens = JSON.parse(pedido.itens);
          } catch (e) {
            Logger.log(`[Pedidos] Erro ao parsear itens JSON para pedido ${pedido.numeroDoPedido}: ${e.message}`);
            pedido.itens = [];
          }
        } else if (!Array.isArray(pedido.itens)) {
          pedido.itens = [];
        }

        pedido.totalGeral = parseFloat(pedido.totalGeral || 0);

        pedidos.push(pedido);
      }

      // --- FILTRO MULTI-EMPRESA ---
      if (reportParams && reportParams.empresaCnpj) {
        const idEmpresaFiltro = String(reportParams.companyId).trim();
        Logger.log(`[_getPedidosDatas] Filtrando pedidos pelo ID da empresa: ${idEmpresaFiltro}`);

        const pedidosFiltrados = pedidos.filter(p => {
          // Assumindo que a coluna na sua planilha 'Pedidos' se chama 'Empresa' ou 'ID Empresa'.
          // O toCamelCase vai transformar isso em 'empresa' ou 'idEmpresa'.
          const idDoPedidoNaLinha = String(p.empresa || p.idEmpresa || '').trim();
          
          // Usamos '==' para o caso da planilha ter o número 1 e o filtro ser "001"
          return idDoPedidoNaLinha == idEmpresaFiltro;
        });

        Logger.log(`[_getPedidosDatas] Encontrados ${pedidosFiltrados.length} pedidos para a empresa.`);
        return pedidosFiltrados;
      }

      Logger.log('[_getPedidosDatas] Nenhum filtro de empresa aplicado. Retornando todos os pedidos.');
      return pedidos;
    }

    function getUniqueProducts(reportParams) {
      Logger.log('[getUniqueProducts - SERVER] Iniciando busca por produtos únicos.');
      const pedidos = _getPedidosDatas(reportParams); // Reutiliza a função que busca e formata os dados dos pedidos
      const uniqueProducts = new Set();

      pedidos.forEach(pedido => {
        (pedido.itens || []).forEach(item => {
          if (item.descricao && item.descricao.trim() !== '') {
            uniqueProducts.add(item.descricao.trim());
          }
        });
      });

      const sortedProducts = Array.from(uniqueProducts).sort((a, b) => a.localeCompare(b));
      Logger.log(`[getUniqueProducts - SERVER] Produtos únicos encontrados: ${JSON.stringify(sortedProducts)}`);
      return sortedProducts;
    }

    /**
     * Função para filtrar pedidos com base em parâmetros.
     * @param {Array<Object>} allPedidos - Todos os pedidos brutos.
     * @param {Object} reportParams - Parâmetros de filtro (startDate, endDate, supplier).
     * @returns {Array<Object>} Pedidos filtrados.
     */
    function _filterPedidos(allPedidos, reportParams) {
        let filtered = allPedidos;

        // Adiciona uma verificação explícita para reportParams aqui
        if (!reportParams || typeof reportParams !== 'object') {
            Logger.log('[_filterPedidos] Aviso: reportParams é undefined ou não é um objeto. Nenhum filtro de data/fornecedor será aplicado.');
            return filtered; // Retorna todos os pedidos se não houver parâmetros válidos
        }

        // Filtro por data
        if (reportParams.startDate && reportParams.endDate) {
            const start = new Date(reportParams.startDate + 'T00:00:00');
            const end = new Date(reportParams.endDate + 'T23:59:59');
            filtered = filtered.filter(pedido => {
                // Se a data do pedido não for um Date object válido, exclua-o ou trate
                if (!(pedido.data instanceof Date)) {
                    Logger.log(`Data inválida no pedido ${pedido.numeroDoPedido}: ${pedido.data}. Excluindo do filtro.`);
                    return false;
                }
                return pedido.data >= start && pedido.data <= end;
            });
            Logger.log(`[_filterPedidos] Filtrado por data: ${filtered.length} pedidos restantes.`);
        }

        // Filtro por fornecedor
        if (reportParams.supplier) {
            const selectedSupplier = reportParams.supplier.toLowerCase();
            filtered = filtered.filter(pedido => (pedido.fornecedor || '').toLowerCase() === selectedSupplier);
            Logger.log(`[_filterPedidos] Filtrado por fornecedor '${reportParams.supplier}': ${filtered.length} pedidos restantes.`);
        }

        return filtered;
    }

    /**
     * Agrupa e sumariza dados de pedidos para geração de relatório.
     * @param {Array<Object>} pedidos - Pedidos já filtrados.
     * @param {Object} reportParams - Parâmetros do relatório (reportType).
     * @returns {Object} Dados agrupados e/ou sumarizados.
     */
    function _groupAndSummarizePedidos(pedidos, reportParams) {
        const data = {};

        if (reportParams.reportType === 'detailed') {
            // Relatório Detalhado: Agrupa por data, depois por fornecedor
            pedidos.sort((a, b) => a.data.getTime() - b.data.getTime()); // Ordena por data

            pedidos.forEach(pedido => {
                const dateStr = Utilities.formatDate(pedido.data, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                if (!data[dateStr]) {
                    data[dateStr] = {
                        date: pedido.data,
                        fornecedores: {},
                        totalDate: 0
                    };
                }
                const fornecedor = pedido.fornecedor || 'Desconhecido';
                if (!data[dateStr].fornecedores[fornecedor]) {
                    data[dateStr].fornecedores[fornecedor] = {
                        pedidos: [],
                        totalFornecedor: 0
                    };
                }
                data[dateStr].fornecedores[fornecedor].pedidos.push(pedido);
                data[dateStr].fornecedores[fornecedor].totalFornecedor += pedido.totalGeral || 0;
                data[dateStr].totalDate += pedido.totalGeral || 0;
            });
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

        let reportTitle = "Relatório de Compras";
        if (reportParams.reportType === 'detailed') {
            reportTitle += " Detalhado";
        } else if (reportParams.reportType === 'financial') {
            reportTitle += " Financeiro";
        }

        let filtersApplied = [];
        if (reportParams.startDate && reportParams.endDate) {
            filtersApplied.push(`Período: ${Utilities.formatDate(new Date(reportParams.startDate + 'T00:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy')} a ${Utilities.formatDate(new Date(reportParams.endDate + 'T00:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy')}`);
        }
        if (reportParams.supplier) {
            filtersApplied.push(`Fornecedor: ${reportParams.supplier}`);
        }
        const filtersHtml = filtersApplied.length > 0 ? `<p class="filters">Filtros Aplicados: ${filtersApplied.join(' | ')}</p>` : '';

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
        return html;
    }

    /**
     * Gera um relatório em PDF com base nos parâmetros fornecidos.
     * @param {Object} reportParams - Objeto com: startDate, endDate, supplier, reportType.
     * @returns {string} URL temporária do arquivo PDF gerado no Google Drive.
     */
    function generatePdfReport(reportParams) {
      Logger.log(`[generatePdfReport] Iniciando. Parâmetros recebidos (antes de qualquer processamento): ${JSON.stringify(reportParams)}`);

      // Adiciona uma verificação explícita para reportParams
      if (reportParams === undefined || reportParams === null) {
          Logger.log('[generatePdfReport] ERRO CRÍTICO: reportParams é undefined ou null. Abortando.');
          throw new Error('Parâmetros de relatório ausentes ou inválidos. Tente novamente.');
      }

      try {
        const allPedidos = _getPedidosDatas(reportParams);
        const filteredPedidos = _filterPedidos(allPedidos, reportParams);
        const reportData = _groupAndSummarizePedidos(filteredPedidos, reportParams);
        const htmlContent = _generatePdfHtmlContent(reportData, reportParams);
        Logger.log(`[generatePdfReport] Conteúdo HTML gerado (primeiros 500 caracteres): ${htmlContent.substring(0, Math.min(htmlContent.length, 500))}`);


        const htmlBlob = Utilities.newBlob(htmlContent, MimeType.HTML, 'RelatorioTemporario.html');
        Logger.log(`[generatePdfReport] htmlBlob criado. Tipo: ${htmlBlob.getContentType()}, Tamanho: ${htmlBlob.getBytes().length} bytes.`);

        const folderName = "RelatoriosComprasTemporarios";
        let folder;
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) { 
          folder = folders.next();
          Logger.log(`[generatePdfReport] Pasta existente encontrada: '${folder.getName()}' (ID: ${folder.getId()})`);
        } else {
          folder = DriveApp.createFolder(folderName);
          Logger.log(`[generatePdfReport] Pasta '${folderName}' criada no Google Drive com ID: ${folder.getId()}.`);
        }

        // Verifica se a pasta foi obtida/criada com sucesso antes de prosseguir
        if (!folder) {
            Logger.log('[generatePdfReport] ERRO: A pasta de destino é nula ou indefinida.');
            throw new Error('Não foi possível obter ou criar a pasta de destino no Google Drive.');
        }

        const htmlFile = folder.createFile(htmlBlob);
        Logger.log(`[generatePdfReport] Arquivo HTML temporário criado: ${htmlFile.getUrl()} (ID: ${htmlFile.getId()}).`);

        // Verifica se o arquivo HTML foi criado com sucesso antes de tentar converter para PDF
        if (!htmlFile) {
            Logger.log('[generatePdfReport] ERRO: O arquivo HTML temporário é nulo ou indefinido após a criação.');
            throw new Error('Falha ao criar o arquivo HTML temporário.');
        }

        // --- PONTO CRÍTICO: Conversão para PDF ---
        Logger.log('[generatePdfReport] Tentando converter o htmlFile para PDF...');
        const pdfBlob = htmlFile.getAs(MimeType.PDF);
        Logger.log(`[generatePdfReport] Conversão para PDF concluída. pdfBlob criado: ${pdfBlob ? 'Sim' : 'Não'}. Tamanho: ${pdfBlob ? pdfBlob.getBytes().length : 'N/A'} bytes.`);

        if (!pdfBlob) {
            Logger.log('[generatePdfReport] ERRO LÓGICO: pdfBlob é nulo ou indefinido após getAs(MimeType.PDF).');
            throw new Error('Falha ao converter o arquivo HTML para PDF: o blob resultante é nulo.');
        }
        
        let titleForFileName = "RelatorioCompras";
        if (reportParams.reportType === 'detailed') {
            titleForFileName += "_Detalhado";
        } else if (reportParams.reportType === 'financial') {
            titleForFileName += "_Financeiro";
        }
        const pdfFileName = `${titleForFileName}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.pdf`;
        
        const pdfFile = folder.createFile(pdfBlob).setName(pdfFileName);
        Logger.log(`[generatePdfReport] Arquivo PDF final criado: ${pdfFile.getUrl()} (ID: ${pdfFile.getId()}).`);

        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileUrl = pdfFile.getDownloadUrl(); 
        Logger.log(`[generatePdfReport] URL de download do PDF: ${fileUrl}`);
        
        // Opcional: Remover o arquivo HTML temporário imediatamente após a conversão
        htmlFile.setTrashed(true); 
        Logger.log(`[generatePdfReport] Arquivo HTML temporário '${htmlFile.getName()}' movido para a lixeira.`);

        return fileUrl;

      } catch (e) {
        Logger.log(`[generatePdfReport] ERRO DURANTE A GERAÇÃO DO PDF: ${e.message}. Stack: ${e.stack}`);
        // Adicione esta linha para relançar um erro mais limpo para o cliente
        throw new Error(`Erro ao gerar relatório PDF: ${e.message}`);
      }
    }

    /**
     * Gera um relatório em XLS (planilha) com base nos parâmetros fornecidos.
     * @param {Object} reportParams - Objeto com parâmetros do relatório (startDate, endDate, supplier, reportType).
     * @returns {string} URL temporária do arquivo XLS gerado no Google Drive.
     */
    function generateXlsReport(reportParams) {
      Logger.log(`[generateXlsReport] Iniciando. Parâmetros recebidos: ${JSON.stringify(reportParams)}`);

      if (reportParams === undefined || reportParams === null) {
          Logger.log('[generateXlsReport] ERRO CRÍTICO: reportParams é undefined ou null. Abortando.');
          throw new Error('Parâmetros de relatório ausentes ou inválidos para XLS. Tente novamente.');
      }

      try {
        const allPedidos = _getPedidosDatas(reportParams);
        const filteredPedidos = _filterPedidos(allPedidos, reportParams);
        const reportData = _groupAndSummarizePedidos(filteredPedidos, reportParams);

        const ssTitle = `RelatorioCompras_${reportParams.reportType}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}`;
        const spreadsheet = SpreadsheetApp.create(ssTitle);
        const spreadsheetId = spreadsheet.getId();
        const sheet = spreadsheet.getActiveSheet();
        sheet.setName("Relatorio");

        const reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
        // Adiciona cabeçalho do relatório na planilha
        const companyName = reportParams.companyName || "EMPRESA NÃO INFORMADA";
        const companyAddress = reportParams.companyAddress || "";
        const companyCnpj = reportParams.empresaCnpj || "";
        
        sheet.getRange('A1').setValue(companyName);
        sheet.getRange('A2').setValue(companyAddress);
        sheet.getRange('A3').setValue(`CNPJ: ${companyCnpj}`);
        sheet.getRange('A4').setValue(`Relatório Gerado em: ${reportDate}`);
        sheet.getRange('A6').setValue(`Relatório de Compras ${reportParams.reportType === 'detailed' ? 'Detalhado' : 'Financeiro'}`);

        let filtersApplied = [];
        if (reportParams.startDate && reportParams.endDate) {
            filtersApplied.push(`Período: ${Utilities.formatDate(new Date(reportParams.startDate + 'T00:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy')} a ${Utilities.formatDate(new Date(reportParams.endDate + 'T00:00:00'), Session.getScriptTimeZone(), 'dd/MM/yyyy')}`);
        }
        if (reportParams.supplier) {
            filtersApplied.push(`Fornecedor: ${reportParams.supplier}`);
        }
        if (filtersApplied.length > 0) {
            sheet.getRange('A7').setValue(`Filtros Aplicados: ${filtersApplied.join(' | ')}`);
        }

        let startRow = 9; // Começa a inserir dados a partir da linha 9

        if (reportParams.reportType === 'detailed') {
            const headers = [
                'Número do Pedido', 'Data do Pedido', 'Fornecedor', 'CNPJ Fornecedor', 'Razão Social',
                'Endereço Fornecedor', 'Condição Pagamento', 'Forma Pagamento', 
                'Placa Veículo', 'Nome Veículo', 'Observações',
                'Descrição do Item', 'Unidade', 'Quantidade', 'Preço Unitário', 'Subtotal Item', 'Total Geral do Pedido'
            ];
            sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
            startRow++;

            let rowData = [];
            // Ordenar por data e depois por fornecedor
            const sortedDates = Object.keys(reportData).sort();
            sortedDates.forEach(dateStr => {
                const dateGroup = reportData[dateStr];
                const sortedSupplierNames = Object.keys(dateGroup.fornecedores).sort();
                sortedSupplierNames.forEach(supplierName => {
                    const supplierGroup = dateGroup.fornecedores[supplierName];
                    supplierGroup.pedidos.forEach(pedido => {
                        const pedidoBaseRow = [
                            pedido.numeroDoPedido,
                            pedido.data ? Utilities.formatDate(pedido.data, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '', // Formata a data
                            pedido.fornecedor,
                            pedido.cnpjFornecedor,
                            supplierName,
                            pedido.enderecoFornecedor,
                            pedido.condicaoPagamentoFornecedor,
                            pedido.formaPagamentoFornecedor,
                            pedido.placaVeiculo,
                            pedido.nomeVeiculo,
                            pedido.observacoes,
                            '', '', '', '', '', '' // Colunas vazias para itens
                        ];

                        if (pedido.itens && pedido.itens.length > 0) {
                            pedido.itens.forEach(item => {
                                const itemRow = [
                                    '', '', '', '', '', '', '', '', '', '', // Colunas vazias para dados do pedido
                                    item.descricao || '',
                                    item.unidade || '',
                                    item.quantidade || 0,
                                    parseFloat(item.precoUnitario || 0),
                                    parseFloat(item.totalItem || 0),
                                    '' // Coluna vazia para total geral do pedido
                                ];
                                rowData.push(pedidoBaseRow.slice(0, 11).concat(itemRow.slice(11))); // Combina dados do pedido com item
                            });
                            // Adiciona a linha do total geral do pedido na última linha de itens
                            const lastItemRow = rowData[rowData.length - 1];
                            if (lastItemRow) {
                                lastItemRow[headers.length - 1] = parseFloat(pedido.totalGeral || 0); // Define o total geral na última coluna do pedido
                            }
                        } else {
                            // Se não houver itens, ainda adiciona a linha do pedido com o total geral
                            const emptyItemRow = ['', '', '', '', '', ''];
                            rowData.push(pedidoBaseRow.slice(0, 11).concat(emptyItemRow));
                            rowData[rowData.length - 1][headers.length - 1] = parseFloat(pedido.totalGeral || 0);
                        }
                    });
                });
            });

            if (rowData.length > 0) {
                sheet.getRange(startRow, 1, rowData.length, headers.length).setValues(rowData);
                // Formatar colunas numéricas
                sheet.getRange(startRow, 13, rowData.length, 1).setNumberFormat('0'); // Quantidade
                sheet.getRange(startRow, 14, rowData.length, 1).setNumberFormat('R$#,##0.00'); // Preço Unitário
                sheet.getRange(startRow, 15, rowData.length, 1).setNumberFormat('R$#,##0.00'); // Subtotal Item
                sheet.getRange(startRow, 16, rowData.length, 1).setNumberFormat('R$#,##0.00'); // Total Geral do Pedido
                Logger.log(`[generateXlsReport] Dados detalhados inseridos na planilha.`);
            } else {
                sheet.getRange(startRow, 1).setValue('Nenhum pedido encontrado para os filtros selecionados.').setFontStyle('italic');
            }
                  // Ajusta as larguras (pode precisar ajustar os índices se quiser)
                sheet.autoResizeColumns(1, headers.length);
                sheet.setColumnWidth(1, 120); // Número do Pedido
                sheet.setColumnWidth(2, 90);  // Data
                sheet.setColumnWidth(3, 180); // Fornecedor
                sheet.setColumnWidth(11, 160); // Descrição do Item

                // Congela cabeçalho da tabela e primeiras colunas
                sheet.setFrozenRows(startRow - 1);
                sheet.setFrozenColumns(2);
        
        // Ajusta larguras
            sheet.autoResizeColumns(1, headers.length);
            sheet.setColumnWidth(1, 120); // Número do Pedido
            sheet.setColumnWidth(2, 90);  // Data
            sheet.setColumnWidth(3, 180); // Fornecedor
            sheet.setColumnWidth(11, 160); // Descrição do Item

            // Congela cabeçalho da tabela e primeiras colunas
            sheet.setFrozenRows(startRow - 1);
            sheet.setFrozenColumns(2);

        } else if (reportParams.reportType === 'financial') {
            sheet.getRange(startRow, 1).setValue(`Total Geral de Pedidos: ${reportData.numeroTotalPedidos}`).setFontWeight('bold');
            sheet.getRange(startRow + 1, 1).setValue(`Valor Total das Compras: R$ ${reportData.totalGeralPedidos.toFixed(2).replace('.', ',')}`).setFontWeight('bold');
            sheet.getRange(startRow + 1, 2).setNumberFormat('R$#,##0.00');

            startRow += 3;

            const supplierHeaders = ['Fornecedor', 'Valor Total Comprado'];
            sheet.getRange(startRow, 1, 1, supplierHeaders.length).setValues([supplierHeaders]);
            
            // Formatação do cabeçalho de fornecedores
            const supHeaderRange = sheet.getRange(startRow, 1, 1, supplierHeaders.length);
            supHeaderRange.setFontWeight('bold')
                          .setFontSize(11)
                          .setBackground('#0056b3')
                          .setFontColor('white')
                          .setHorizontalAlignment('center')
                          .setVerticalAlignment('middle')
                          .setBorder(true, true, true, true, true, true);

            startRow++;

            const supplierData = reportData.listaTotalPorFornecedor.map(item => [
                item.fornecedor,
                item.total
            ]);

            if (supplierData.length > 0) {
                sheet.getRange(startRow, 1, supplierData.length, supplierHeaders.length).setValues(supplierData);
                sheet.getRange(startRow, 2, supplierData.length, 1).setNumberFormat('R$#,##0.00');
                // Formatação dos dados
                const supDataRange = sheet.getRange(startRow, 1, supplierData.length, supplierHeaders.length);
                supDataRange.setFontSize(10).setVerticalAlignment('middle');
                supDataRange.setBorder(true, true, true, true, false, false);
            } else {
                sheet.getRange(startRow, 1).setValue('Nenhum dado financeiro por fornecedor encontrado.').setFontStyle('italic');
            }
            sheet.autoResizeColumns(1, supplierHeaders.length);
            sheet.setFrozenRows(startRow - 1);
        }

        // Cria uma pasta temporária (ou usa uma existente)
        const folderName = "RelatoriosComprasTemporarios";
        let folder;
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) { 
          folder = folders.next();
          Logger.log(`[generateXlsReport] Pasta existente encontrada: '${folder.getName()}' (ID: ${folder.getId()})`);
        } else {
          folder = DriveApp.createFolder(folderName);
          Logger.log(`[generateXlsReport] Pasta '${folderName}' criada no Google Drive com ID: ${folder.getId()}.`);
        }

        if (!folder) {
            Logger.log('[generateXlsReport] ERRO: A pasta de destino é nula ou indefinida.');
            throw new Error('Não foi possível obter ou criar a pasta de destino no Google Drive.');
        }

        Logger.log(`[generateXlsReport] ID da planilha temporária criada: ${spreadsheetId}`);

        const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;
        Logger.log(`[generateXlsReport] URL de exportação para XLSX: ${exportUrl}`);

        let xlsBlob;
        try {
            // Obtenha o token de acesso para autenticação
            const token = ScriptApp.getOAuthToken();
            Logger.log('[generateXlsReport] Token OAuth obtido com sucesso.');

            // Faça a requisição para exportar o arquivo como XLSX
            const response = UrlFetchApp.fetch(exportUrl, {
                headers: {
                    Authorization: `Bearer ${token}`
                },
                muteHttpExceptions: true // Para capturar erros na resposta
            });

            if (response.getResponseCode() === 200) {
                xlsBlob = response.getBlob();
                Logger.log(`[generateXlsReport] Blob XLS (exportado via URL) criado. Tipo: ${xlsBlob.getContentType()}, Tamanho: ${xlsBlob.getBytes().length} bytes.`);
            } else {
                const errorText = response.getContentText();
                Logger.log(`[generateXlsReport] Erro na resposta da exportação: Código ${response.getResponseCode()}, Texto: ${errorText}`);
                throw new Error(`Falha ao exportar a planilha para XLSX via URL. Código: ${response.getResponseCode()}, Mensagem: ${errorText}`);
            }

        } catch (exportError) {
            Logger.log(`[generateXlsReport] Erro durante a exportação via URLFetchApp: ${exportError.message}. Stack: ${exportError.stack}`);
            throw new Error(`Erro ao gerar relatório XLS (exportação): ${exportError.message}`);
        } finally {
            // Sempre mova a planilha temporária para a lixeira, independentemente do sucesso da exportação
            if (spreadsheetId) {
                try {
                    DriveApp.getFileById(spreadsheetId).setTrashed(true);
                    Logger.log(`[generateXlsReport] Planilha temporária '${spreadsheet.getName()}' (ID: ${spreadsheetId}) movida para a lixeira.`);
                } catch (cleanupError) {
                    Logger.log(`[generateXlsReport] Erro ao mover planilha temporária para a lixeira (ID: ${spreadsheetId}): ${cleanupError.message}`);
                }
            }
        }

        if (!xlsBlob) {
            Logger.log('[generateXlsReport] ERRO LÓGICO: xlsBlob é nulo após tentativa de exportação via URL.');
            throw new Error('Falha ao gerar relatório XLS: o blob resultante da exportação é nulo.');
        }

        const xlsFileName = `${ssTitle}.xlsx`;
        const xlsFile = folder.createFile(xlsBlob).setName(xlsFileName);
        Logger.log(`[generateXlsReport] Arquivo XLS final criado: ${xlsFile.getUrl()} (ID: ${xlsFile.getId()}).`);

        xlsFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileUrl = xlsFile.getDownloadUrl(); 
        Logger.log(`[generateXlsReport] URL de download do XLS: ${fileUrl}`);

        return fileUrl;

      } catch (e) {
        Logger.log(`[generateXlsReport] ERRO DURANTE A GERAÇÃO DO XLS: ${e.message}. Stack: ${e.stack}`);
        throw new Error(`Erro ao gerar relatório XLS: ${e.message}`);
      }
    }

    // ===============================================
    // FUNÇÕES PARA TROCA DE SENHA 
    // ===============================================
    function listarEmpresas() {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Empresas');
      if (!sheet || sheet.getLastRow() < 2) return [];
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
      return data.map(([codigo, nome]) => ({ codigo: String(codigo), nome: String(nome) }));
    }

    function getEmpresasDoUsuario(usuarioId) {
      const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
      const data = sheet.getDataRange().getValues();
      const header = data[0];
      const idxUsuario = header.indexOf('USUARIO');
      const idxIdEmpresa = header.indexOf('ID EMPRESA');
      const idxNomeEmpresa = header.indexOf('EMPRESA'); // Se você tiver uma coluna com nome da empresa

      const empresas = [];

      for(let i = 1; i < data.length; i++) {
        if(data[i][idxUsuario] === usuarioId) {
          empresas.push({
            idEmpresa: data[i][idxIdEmpresa],
            nomeEmpresa: data[i][idxNomeEmpresa] || `Empresa ${data[i][idxIdEmpresa]}`
          });
        }
      }
      return empresas;
    }

    // Exemplo para salvar empresa selecionada no usuário_logado
    function registrarEmpresaSelecionadaNoLogin(empresaId) {
      // Você pode salvar em uma tabela USUARIO_LOGADO ou onde preferir
      // Exemplo simples:
      const sheet = SpreadsheetApp.getActive().getSheetByName('UsuarioLogado');
      sheet.clearContents(); // limpa registros anteriores
      sheet.appendRow([Session.getActiveUser().getEmail(), empresaId, new Date()]);
    }

    /**
     * Altera a senha do usuário logado.
     * @param {string} login - login do usuário (NÃO é e-mail)
     * @param {string} senhaAtual - senha atual digitada
     * @param {string} novaSenha - nova senha digitada
     * @returns {Object} {status: 'success'|'error', message: string}
     */
    function alterarSenhaUsuario(login, senhaAtual, novaSenha) {
      try {
        if (!login) return {status: 'error', message: 'Login não informado.'};
        if (!senhaAtual || !novaSenha) return {status: 'error', message: 'Preencha todos os campos.'};
        
        var sh = SpreadsheetApp.getActive().getSheetByName('Usuarios');
        var dados = sh.getDataRange().getValues();
        var idxLogin = dados[0].indexOf('USUARIO');
        var idxSenha = dados[0].indexOf('SENHA');
        if (idxLogin < 0 || idxSenha < 0)
          return {status: 'error', message: 'Planilha de usuários mal configurada.'};

        var rowIdx = -1;
        for (var i=1; i<dados.length; i++) {
          if (String(dados[i][idxLogin]).toLowerCase().trim() === login.toLowerCase().trim()) {
            rowIdx = i;
            break;
          }
        }
        if (rowIdx < 0) return {status: 'error', message: 'Usuário não encontrado.'};

        // Verificar a senha atual (Atenção: simples, para produção use hash)
        var senhaAtualArmazenada = dados[rowIdx][idxSenha];
        var hashInformado = gerarHash(senhaAtual);
        if (senhaAtualArmazenada !== senhaAtual) {
          return {status: 'error', message: 'Senha atual incorreta.'};
        }

        // Atualiza senha com novo hash
        var novoHash = gerarHash(novaSenha);
        sh.getRange(rowIdx+1, idxSenha+1).setValue(novoHash);

        return {status: 'success'};
      } catch (e) {
        return {status: 'error', message: 'Falha ao trocar senha: ' + e.message};
      }
    }

    /**
     * Função de teste para gerar o hash de uma senha específica e exibi-lo no log.
     */
    function testarHashDeSenha() {
      const senhaParaTestar = '1234';
      const hashResultante = gerarHash(senhaParaTestar);

      Logger.log(`O hash SHA-256 para a senha "${senhaParaTestar}" é:`);
      Logger.log(hashResultante);
    }

    function gerarHash(senha) {
      var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, senha);
      return Utilities.base64Encode(digest);
    }

    function adminRedefinirSenha(usuario, novaSenha) {
      var sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
      var dados = sheet.getDataRange().getValues();
      var idxUsuario = dados[0].indexOf('USUARIO');
      var idxSenha = dados[0].indexOf('SENHA');
      var linha = -1;

      for (var i = 1; i < dados.length; i++) {
        if (String(dados[i][idxUsuario]).toLowerCase().trim() === usuario.toLowerCase().trim()) {
          linha = i;
          break;
        }
      }
      if (linha < 0) return { status: 'error', message: 'Usuário não encontrado.' };

      var hashNovaSenha = gerarHash(novaSenha);
      sheet.getRange(linha + 1, idxSenha + 1).setValue(hashNovaSenha);

      return { status: 'ok', message: 'Senha redefinida com sucesso.' };
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
    // FUNÇÕES PARA GERENCIAMENTO DE USUÁRIOS COM AUDITORIA
    // ===============================================

    /**
     * Lista todos os usuários com dados formatados para o frontend
     * @returns {Array<Object>} Lista de usuários com empresas processadas
     */
    function listarUsuariosCompleto() {
      try {
        const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
        if (!sheet || sheet.getLastRow() < 2) return [];

        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
        const sheetEmpresas = SpreadsheetApp.getActive().getSheetByName('Empresas');
        
        // Mapear empresas por ID para conversão
        let empresasMap = {};
        if (sheetEmpresas && sheetEmpresas.getLastRow() >= 2) {
          const empresasData = sheetEmpresas.getRange(2, 1, sheetEmpresas.getLastRow() - 1, 2).getValues();
          empresasData.forEach(([id, nome]) => {
            empresasMap[String(id).trim().padStart(3, '0')] = String(nome).trim();
          });
        }

        return data.map(([id, nome, usuario, senha, perfil, status, empresasStr, empresaPadrao]) => {
          const empresasIds = empresasStr ? String(empresasStr).split(',').map(e => e.trim()) : [];
          const empresasNomes = empresasIds.map(id => empresasMap[String(id).padStart(3, '0')] || `Empresa ${id}`);
          const empresaPadraoNome = empresaPadrao ? empresasMap[String(empresaPadrao).padStart(3, '0')] || '' : '';

          return {
            id: String(id),
            nome: String(nome),
            usuario: String(usuario),
            perfil: String(perfil),
            status: String(status),
            empresas: empresasNomes.join(','),
            empresaPadrao: empresaPadraoNome,
            empresasIds: empresasIds,
            empresaPadraoId: empresaPadrao ? String(empresaPadrao) : ''
          };
        });
      } catch (error) {
        Logger.log('Erro em listarUsuariosCompleto: ' + error.message);
        return [];
      }
    }

    /**
     * Obter dados de auditoria de um usuário
     * @param {string} userId - ID do usuário
     * @returns {Object} Dados de auditoria
     */
    function obterDadosAuditoria(userId) {
      try {
        const result = {
          lastLogin: null,
          lastOrder: null,
          lastPrint: null
        };

        // 1. Último login
        const loginSheet = SpreadsheetApp.getActive().getSheetByName('usuario_logado');
        if (loginSheet && loginSheet.getLastRow() > 1) {
          const loginData = loginSheet.getDataRange().getValues();
          const loginEntries = loginData.slice(1).filter(row => String(row[0]).replace("'", "") === String(userId));
          if (loginEntries.length > 0) {
            // Pegar o mais recente
            const lastLoginEntry = loginEntries.sort((a, b) => new Date(b[4]) - new Date(a[4]))[0];
            result.lastLogin = {
              date: lastLoginEntry[4].toISOString(),
              ip: '192.168.1.100' // IP simulado - você pode implementar captura real
            };
          }
        }

        // 2. Último pedido criado
        const pedidosSheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
        if (pedidosSheet && pedidosSheet.getLastRow() > 1) {
          const pedidosData = pedidosSheet.getDataRange().getValues();
          const headers = pedidosData[0];
          const numeroIdx = headers.findIndex(h => String(h).toUpperCase().includes('NÚMERO') || String(h).toUpperCase().includes('NUMERO'));
          const dataIdx = headers.findIndex(h => String(h).toUpperCase().includes('DATA'));
          
          if (numeroIdx > -1 && dataIdx > -1) {
            const userPedidos = pedidosData.slice(1).filter(row => {
              // Aqui você pode ajustar a lógica para identificar pedidos do usuário
              // Por exemplo, se há uma coluna "USUARIO_CRIADOR" ou similar
              return true; // Temporário - implementar lógica específica
            });
            
            if (userPedidos.length > 0) {
              const lastPedido = userPedidos.sort((a, b) => new Date(b[dataIdx]) - new Date(a[dataIdx]))[0];
              result.lastOrder = {
                id: String(lastPedido[numeroIdx]),
                date: lastPedido[dataIdx].toISOString()
              };
            }
          }
        }

        // 3. Última impressão - dados simulados
        // Você pode implementar um log de impressões se necessário
        if (Math.random() > 0.5) { // Simular alguns usuários com impressões
          result.lastPrint = {
            id: 'PC-2025-0789',
            date: new Date(Date.now() - Math.random() * 7 * 24 * 60 * 60 * 1000).toISOString()
          };
        }

        return result;
      } catch (error) {
        Logger.log('Erro em obterDadosAuditoria: ' + error.message);
        return { lastLogin: null, lastOrder: null, lastPrint: null };
      }
    }

    /**
     * Alternar status de usuário (Ativo/Inativo)
     * @param {string} userId - ID do usuário
     * @returns {Object} Resultado da operação
     */
    function alternarStatusUsuario(userId) {
      try {
        const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
        if (!sheet) return { status: 'error', message: 'Planilha Usuarios não encontrada' };

        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const idIdx = headers.findIndex(h => String(h).toUpperCase() === 'ID');
        const statusIdx = headers.findIndex(h => String(h).toUpperCase() === 'STATUS');
        const nomeIdx = headers.findIndex(h => String(h).toUpperCase() === 'NOME');

        if (idIdx === -1 || statusIdx === -1) {
          return { status: 'error', message: 'Colunas ID ou STATUS não encontradas' };
        }

        for (let i = 1; i < data.length; i++) {
          if (String(data[i][idIdx]) === String(userId)) {
            const statusAtual = String(data[i][statusIdx]);
            const novoStatus = statusAtual === 'Ativo' ? 'Inativo' : 'Ativo';
            const nomeUsuario = String(data[i][nomeIdx]);
            
            sheet.getRange(i + 1, statusIdx + 1).setValue(novoStatus);
            
            return { 
              status: 'ok', 
              message: `Usuário ${nomeUsuario} foi ${novoStatus === 'Ativo' ? 'ativado' : 'desativado'}`,
              novoStatus: novoStatus,
              nomeUsuario: nomeUsuario
            };
          }
        }

        return { status: 'error', message: 'Usuário não encontrado' };
      } catch (error) {
        Logger.log('Erro em alternarStatusUsuario: ' + error.message);
        return { status: 'error', message: 'Erro interno: ' + error.message };
      }
    }

    /**
     * Salvar permissões de empresa para usuário (versão completa)
     * @param {string} userId - ID do usuário
     * @param {Array} empresasIds - Array de IDs das empresas
     * @param {string} empresaPadraoId - ID da empresa padrão
     * @returns {Object} Resultado da operação
     */
    function salvarPermissoesEmpresaUsuario(userId, empresasIds, empresaPadraoId) {
      try {
        Logger.log(`[salvarPermissoesEmpresaUsuario] Iniciando para usuário ${userId}`);
        Logger.log(`[salvarPermissoesEmpresaUsuario] Empresas: ${JSON.stringify(empresasIds)}`);
        Logger.log(`[salvarPermissoesEmpresaUsuario] Empresa padrão: ${empresaPadraoId}`);
        
        const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
        if (!sheet) return { status: 'error', message: 'Planilha Usuarios não encontrada' };

        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const idIdx = headers.findIndex(h => String(h).toUpperCase() === 'ID');
        const empresasIdx = headers.findIndex(h => String(h).toUpperCase().includes('EMPRESA') && !String(h).toUpperCase().includes('PADRÃO'));
        const empresaPadraoIdx = headers.findIndex(h => String(h).toUpperCase().includes('EMPRESA') && String(h).toUpperCase().includes('PADRÃO'));
        const nomeIdx = headers.findIndex(h => String(h).toUpperCase() === 'NOME');

        Logger.log(`[salvarPermissoesEmpresaUsuario] Índices - ID: ${idIdx}, Empresas: ${empresasIdx}, Padrão: ${empresaPadraoIdx}`);

        if (idIdx === -1 || empresasIdx === -1) {
          return { status: 'error', message: 'Colunas necessárias não encontradas' };
        }

        for (let i = 1; i < data.length; i++) {
          if (String(data[i][idIdx]) === String(userId)) {
            const empresasStr = Array.isArray(empresasIds) ? empresasIds.join(',') : '';
            const nomeUsuario = String(data[i][nomeIdx]);
            
            // Salva lista de empresas
            sheet.getRange(i + 1, empresasIdx + 1).setValue(empresasStr);
            
            // Salva empresa padrão se a coluna existir e a empresa padrão estiver na lista
            if (empresaPadraoIdx > -1) {
              const empresaPadraoValida = empresaPadraoId && empresasIds.includes(empresaPadraoId) ? empresaPadraoId : '';
              sheet.getRange(i + 1, empresaPadraoIdx + 1).setValue(empresaPadraoValida);
            }
            
            Logger.log(`[salvarPermissoesEmpresaUsuario] Permissões salvas para usuário ${nomeUsuario}`);
            
            return { 
              status: 'ok', 
              message: `Permissões de empresas atualizadas para ${nomeUsuario}`,
              empresas: empresasStr,
              empresaPadrao: empresaPadraoId
            };
          }
        }

        return { status: 'error', message: 'Usuário não encontrado' };
      } catch (error) {
        Logger.log('Erro em salvarPermissoesEmpresaUsuario: ' + error.message);
        return { status: 'error', message: 'Erro interno: ' + error.message };
      }
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

    /**
     * ===============================================
     * FUNÇÕES DE DIAGNÓSTICO
     * ===============================================
     */

    /**
     * FUNÇÃO DE DIAGNÓSTICO COMPLETO PARA VERIFICAR ESTRUTURA DAS PLANILHAS
     * Execute esta função no Google Apps Script para verificar se tudo está correto
     */
    function diagnosticoCompleto() {
      console.log('🔍 === DIAGNÓSTICO COMPLETO DO SISTEMA ===');
      
      try {
        const planilha = SpreadsheetApp.openById(PLANILHA_ID);
        console.log('✅ Planilha acessada:', planilha.getName());
        
        // 1. Verificar aba Fornecedores
        console.log('\n📋 === ABA FORNECEDORES ===');
        const abaFornecedores = planilha.getSheetByName('Fornecedores');
        if (abaFornecedores) {
          const headersFornecedores = abaFornecedores.getRange(1, 1, 1, abaFornecedores.getLastColumn()).getValues()[0];
          console.log('📊 Colunas encontradas na aba Fornecedores:', headersFornecedores.length);
          headersFornecedores.forEach((header, index) => {
            console.log(`   Coluna ${index + 1}: "${header}"`);
          });
          
          // Verificar se há fornecedores com estado preenchido
          const dadosFornecedores = abaFornecedores.getRange(2, 1, Math.min(3, abaFornecedores.getLastRow() - 1), abaFornecedores.getLastColumn()).getValues();
          console.log('\n🔍 Primeiros fornecedores (amostra):');
          dadosFornecedores.forEach((fornecedor, index) => {
            console.log(`   Fornecedor ${index + 1}:`);
            console.log(`     - Razão Social: "${fornecedor[1]}"`);
            console.log(`     - Estado (Coluna 11): "${fornecedor[10]}"`);
          });
        } else {
          console.log('❌ Aba Fornecedores não encontrada!');
        }
        
        // 2. Verificar aba Pedidos
        console.log('\n📋 === ABA PEDIDOS ===');
        const abaPedidos = planilha.getSheetByName('Pedidos');
        if (abaPedidos) {
          const headersPedidos = abaPedidos.getRange(1, 1, 1, abaPedidos.getLastColumn()).getValues()[0];
          console.log('📊 Colunas encontradas na aba Pedidos:', headersPedidos.length);
          headersPedidos.forEach((header, index) => {
            console.log(`   Coluna ${index + 1}: "${header}"`);
          });
          
          // Verificar se existe coluna "Estado Fornecedor"
          const indiceEstadoFornecedor = headersPedidos.indexOf('Estado Fornecedor');
          if (indiceEstadoFornecedor !== -1) {
            console.log(`✅ Coluna "Estado Fornecedor" encontrada na posição ${indiceEstadoFornecedor + 1}`);
          } else {
            console.log('⚠️ Coluna "Estado Fornecedor" NÃO encontrada na aba Pedidos');
            console.log('   Você precisa adicionar esta coluna manualmente na planilha');
          }
        } else {
          console.log('❌ Aba Pedidos não encontrada!');
        }
        
        // 3. Testar busca de fornecedor
        console.log('\n🔍 === TESTE DE BUSCA DE FORNECEDOR ===');
        if (abaFornecedores) {
          const dadosTesteFornecedor = abaFornecedores.getRange(2, 1, 1, abaFornecedores.getLastColumn()).getValues()[0];
          if (dadosTesteFornecedor && dadosTesteFornecedor[1]) {
            const nomeFornecedor = String(dadosTesteFornecedor[1]);
            console.log(`🔍 Testando busca do fornecedor: "${nomeFornecedor}"`);
            
            const fornecedoresData = abaFornecedores.getRange(2, 1, abaFornecedores.getLastRow() - 1, abaFornecedores.getLastColumn()).getValues();
            const foundFornecedor = fornecedoresData.find(row => String(row[1]) === nomeFornecedor);
            
            if (foundFornecedor) {
              console.log('✅ Fornecedor encontrado!');
              console.log(`   - CNPJ: "${foundFornecedor[3]}"`);
              console.log(`   - Endereço: "${foundFornecedor[4]}"`);
              console.log(`   - Condição Pagamento: "${foundFornecedor[5]}"`);
              console.log(`   - Forma Pagamento: "${foundFornecedor[6]}"`);
              console.log(`   - Estado: "${foundFornecedor[10]}" <- IMPORTANTE!`);
            } else {
              console.log('❌ Fornecedor não encontrado na busca');
            }
          }
        }
        
        console.log('\n✅ === DIAGNÓSTICO CONCLUÍDO ===');
        return {
          status: 'success',
          message: 'Diagnóstico executado com sucesso! Verifique o console para detalhes.'
        };
        
      } catch (error) {
        console.error('❌ Erro no diagnóstico:', error);
        return {
          status: 'error',
          message: 'Erro no diagnóstico: ' + error.message
        };
      }
    }

    /**
     * FUNÇÃO ESPECÍFICA PARA TESTAR A CAPTURA DO ESTADO DO FORNECEDOR
     * Execute esta função com o nome de um fornecedor específico
     */
    function testarEstadoFornecedor(nomeFornecedor) {
      console.log(`🔍 === TESTE ESPECÍFICO DO ESTADO DO FORNECEDOR ===`);
      console.log(`📋 Fornecedor: "${nomeFornecedor}"`);
      
      try {
        const planilha = SpreadsheetApp.openById(PLANILHA_ID);
        const fornecedoresSheet = planilha.getSheetByName('Fornecedores');
        
        if (!fornecedoresSheet) {
          console.log('❌ Aba Fornecedores não encontrada!');
          return { status: 'error', message: 'Aba Fornecedores não encontrada' };
        }
        
        const fornecedoresData = fornecedoresSheet.getRange(2, 1, fornecedoresSheet.getLastRow() - 1, fornecedoresSheet.getLastColumn()).getValues();
        const foundFornecedor = fornecedoresData.find(row => String(row[1]) === nomeFornecedor);
        
        if (foundFornecedor) {
          console.log('✅ Fornecedor encontrado!');
          console.log('📊 Dados capturados:');
          console.log(`   - Código: "${foundFornecedor[0]}"`);
          console.log(`   - Razão Social: "${foundFornecedor[1]}"`);
          console.log(`   - Nome Fantasia: "${foundFornecedor[2]}"`);
          console.log(`   - CNPJ: "${foundFornecedor[3]}"`);
          console.log(`   - Endereço: "${foundFornecedor[4]}"`);
          console.log(`   - Condição Pagamento: "${foundFornecedor[5]}"`);
          console.log(`   - Forma Pagamento: "${foundFornecedor[6]}"`);
          console.log(`   - Grupo: "${foundFornecedor[7]}"`);
          console.log(`   - Estado: "${foundFornecedor[10]}" <- ESTE É O ESTADO!`);
          
          // Simular o que seria salvo no pedido
          const estadoFornecedor = String(foundFornecedor[10] || '');
          console.log(`\n💾 Estado que seria salvo no pedido: "${estadoFornecedor}"`);
          
          if (estadoFornecedor.trim()) {
            console.log('✅ Estado preenchido e será salvo corretamente!');
          } else {
            console.log('⚠️ Estado vazio - verifique se o fornecedor tem estado preenchido na planilha');
          }
          
          return {
            status: 'success',
            estadoEncontrado: estadoFornecedor,
            dadosCompletos: {
              cnpj: String(foundFornecedor[3] || ''),
              endereco: String(foundFornecedor[4] || ''),
              condicao: String(foundFornecedor[5] || ''),
              forma: String(foundFornecedor[6] || ''),
              estado: estadoFornecedor
            }
          };
        } else {
          console.log('❌ Fornecedor não encontrado!');
          console.log('📋 Fornecedores disponíveis:');
          fornecedoresData.slice(0, 5).forEach((row, index) => {
            console.log(`   ${index + 1}. "${row[1]}"`);
          });
          
          return {
            status: 'error',
            message: 'Fornecedor não encontrado'
          };
        }
        
      } catch (error) {
        console.error('❌ Erro ao testar estado do fornecedor:', error);
        return {
          status: 'error',
          message: 'Erro: ' + error.message
        };
      }
    }

    /**
     * ===============================================
     * FUNÇÕES AUXILIARES
     * ===============================================
     */

    /**
     * Formata uma data para o padrão ISO com data e hora
     * @param {string|Date} data - Data a ser formatada
     * @returns {string} - Data no formato ISO (YYYY-MM-DD HH:mm:ss)
     */
    function formatarDataParaISO(data) {
      try {
        let dataObj;
        
        if (!data) {
          dataObj = new Date();
        } else if (typeof data === 'string') {
          // Se for apenas data (YYYY-MM-DD), adicionar hora atual para evitar problemas de fuso horário
          if (data.match(/^\d{4}-\d{2}-\d{2}$/)) {
            // Para data simples, usar a hora atual em vez de meio-dia
            const agora = new Date();
            const horaAtual = agora.getHours().toString().padStart(2, '0');
            const minutoAtual = agora.getMinutes().toString().padStart(2, '0');
            const segundoAtual = agora.getSeconds().toString().padStart(2, '0');
            
            dataObj = new Date(data + `T${horaAtual}:${minutoAtual}:${segundoAtual}`);
            console.log('📅 Data simples convertida com hora atual:', data, '→', dataObj.toISOString());
          } else {
            // Para data com hora, usar como está
            dataObj = new Date(data);
          }
          
          // Se a conversão resultou em data inválida, usar data atual
          if (isNaN(dataObj.getTime())) {
            console.warn('Data inválida recebida:', data, 'Usando data atual');
            dataObj = new Date();
          }
        } else if (data instanceof Date) {
          dataObj = data;
        } else {
          console.warn('Tipo de data não reconhecido:', typeof data, data, 'Usando data atual');
          dataObj = new Date();
        }
        
        // Formatar para YYYY-MM-DD HH:mm:ss usando fuso horário do Brasil
        return Utilities.formatDate(dataObj, 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm:ss');
        
      } catch (error) {
        console.error('Erro ao formatar data:', error, 'Data recebida:', data);
        // Em caso de erro, usar data atual
        return Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm:ss');
      }
    }

    /**
     * Salva dados na planilha usando a mesma estrutura da função salvarPedido
     * @param {Sheet} aba - Aba da planilha
     * @param {Object} dados - Dados para salvar
     * @param {number} linha - Linha para atualizar (opcional, para novos registros)
     */
    function salvarDadosNaPlanilha(aba, dados, linha = null) {
      const headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
      const rowData = new Array(headers.length).fill('');

      // Preencher dados conforme os cabeçalhos da planilha
      headers.forEach((header, index) => {
        if (dados.hasOwnProperty(header)) {
          rowData[index] = dados[header];
        }
      });

      if (linha) {
        // Atualizar linha existente
        aba.getRange(linha, 1, 1, rowData.length).setValues([rowData]);
      } else {
        // Adicionar nova linha
        aba.getRange(aba.getLastRow() + 1, 1, 1, rowData.length).setValues([rowData]);
      }
    }

    /**
     * Encontra a linha de um rascunho na planilha
     * @param {Sheet} aba - Aba da planilha
     * @param {string} rascunhoId - ID do rascunho
     * @returns {number} - Número da linha (0 se não encontrado)
     */
    function encontrarLinhaRascunho(aba, rascunhoId) {
      const dados = aba.getDataRange().getValues();
      const cabecalhos = dados[0];
      const indiceNumeroPedido = cabecalhos.indexOf('Número do Pedido');
      const indiceStatus = cabecalhos.indexOf('Status');
      
      for (let i = 1; i < dados.length; i++) {
        const numeroRascunho = dados[i][indiceNumeroPedido] ? dados[i][indiceNumeroPedido].replace("'", "") : '';
        if (numeroRascunho === rascunhoId && dados[i][indiceStatus] === 'RASCUNHO') {
          return i + 1; // +1 porque getRange é 1-indexed
        }
      }
      
      return 0;
    }

    /**
     * Valida dados para conversão de rascunho em pedido
     * @param {Object} dados - Dados do rascunho
     * @returns {Object} - Resultado da validação
     */
    function validarDadosParaPedido(dados) {
      if (!dados.fornecedor || !dados.fornecedor.trim()) {
        return {
          valido: false,
          mensagem: 'Fornecedor é obrigatório.'
        };
      }
      
      if (!dados.itens || dados.itens.length === 0) {
        return {
          valido: false,
          mensagem: 'Pelo menos um item é obrigatório.'
        };
      }
      
      // Validar se todos os itens têm descrição
      const itensValidos = dados.itens.every(item => item.descricao && item.descricao.trim());
      if (!itensValidos) {
        return {
          valido: false,
          mensagem: 'Todos os itens devem ter uma descrição.'
        };
      }
      
      return {
        valido: true,
        mensagem: 'Dados válidos para finalização.'
      };
    }

    /**
     * Gera um número único para o pedido
     * @returns {string} - Número do pedido
     */
    function gerarNumeroPedido() {
      const agora = new Date();
      const ano = agora.getFullYear();
      const mes = String(agora.getMonth() + 1).padStart(2, '0');
      const dia = String(agora.getDate()).padStart(2, '0');
      const timestamp = agora.getTime();
      
      return `PED-${ano}${mes}${dia}-${timestamp}`;
    }

    // ===============================================
    // CONSTANTES E CONFIGURAÇÕES
    // ===============================================
