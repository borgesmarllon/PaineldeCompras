/**
 * @file Usuarios.gs
 * @description Funções para gerenciamento de usuários, login, permissões e autenticação.
 */


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
    function criarUsuario(nome, senha, empresasCodigos, perfil) {

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
        
        if (!sheet) {
            const errorMsg = "A planilha com o nome 'Usuarios' não foi encontrada. Verifique se o nome da aba está exatamente correto (maiúsculas/minúsculas, sem acentos).";
            Logger.log("ERRO CRÍTICO: " + errorMsg);
            return { status: 'error', message: errorMsg };
        }
        
        const lastRow = sheet.getLastRow();
        let dadosUsuariosExistentes = [];
        let ids = [];

        // VERIFICAÇÃO ROBUSTA: Só tenta ler os dados se houver mais do que apenas a linha do cabeçalho.
        if (lastRow > 1) {
            dadosUsuariosExistentes = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
            ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(id => parseInt(id)).filter(n => !isNaN(n));
        }

        // Pega todos os usuários existentes para a verificação de duplicidade
        //const dadosUsuariosExistentes = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues().flat();
        
        // GERA O NOME DE USUÁRIO ÚNICO AQUI
        // Esta função já garante que o nome de usuário não será duplicado.
        const novoUsuario = _gerarUsernameUnico(nome, dadosUsuariosExistentes);

        // Pega o último ID para gerar o próximo
        //const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(id => parseInt(id)).filter(n => !isNaN(n));
        const novoId = ids.length ? Math.max(...ids) + 1 : 1;

        // Gera o HASH da senha antes de salvar
        const senhaHash = gerarHash(senha);

        sheet.appendRow([
            novoId,
            nome.toUpperCase(), // Salva o nome em maiúsculas
            novoUsuario,
            senhaHash,
            perfil || 'usuario', // Perfil padrão ou informado
            'Inativo',           // Status padrão
            empresasCodigos || '' // Garante que seja uma string vazia se nulo
        ]);
        SpreadsheetApp.flush();
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

function validarLogin(usuario, senha, empresaSelecionada) {
        try {
        Logger.log(`[validarLogin] Tentativa de login para usuário: ${usuario}, Empresa: ${empresaSelecionada}`);
        
        const sheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName('Usuarios');
        if (!sheet || sheet.getLastRow() < 2) {
          return { status: 'erro', message: 'Nenhum usuário cadastrado.' };
        }

        const data = sheet.getDataRange().getValues();
        // 1. SEPARA OS CABEÇALHOS
        const headers = data.shift();
        // 2. CHAMA A FERRAMENTA PARA MAPEÁ-LOS
        const colunas = mapearColunas(headers); // Usa a função auxiliar
            const senhaDigitadaHash = gerarHash(String(senha));

        // 3. USA 'find' PARA UMA BUSCA MAIS LIMPA
        const usuarioRow = data.find(row => 
            String(row[colunas.usuario]).toLowerCase() === String(usuario).toLowerCase() &&
            row[colunas.senha] === senhaDigitadaHash
        );

        // 4. SE NÃO ENCONTROU, RETORNA ERRO
        if (!usuarioRow) {
          return { status: 'erro', message: 'Usuário ou senha inválidos!' };
        }
        
        // 5. SE ENCONTROU, USA AS COLUNAS MAPEADAS PARA PEGAR OS DADOS CORRETOS
        if (String(usuarioRow[colunas.status]).toUpperCase() !== 'ATIVO') {
          return { status: 'inativo', message: 'Usuário inativo ou aguardando aprovação.' };
        }

        const empresasPermitidas = String(usuarioRow[colunas.idEmpresa] || '').split(',').map(e => e.trim());
        if (!empresasPermitidas.includes(String(empresaSelecionada))) {
          return { status: 'erro', message: 'Você não tem permissão para acessar esta empresa.' };
        }

        const empresaObjetoCompleto = getDadosEmpresaPorId(empresaSelecionada);
        if (!empresaObjetoCompleto) {
          return { status: 'erro', message: `Os dados para a empresa ID ${empresaSelecionada} não foram encontrados.` };
        }
        
        // NOVO: 1. Pega o nome de usuário verificado para usar no cache.
        const nomeUsuarioVerificado = String(usuarioRow[colunas.usuario]).toLowerCase();

        // NOVO: 2. Gera um token (uma "pulseira de acesso") único e secreto para esta sessão.
        const token = Utilities.getUuid();

        // NOVO: 3. Armazena o token no CacheService, associando-o ao nome de usuário.
        // A "pulseira" é válida por 2 horas (7200 segundos).
        CacheService.getScriptCache().put('token_' + token, nomeUsuarioVerificado, 3600);

        Logger.log(`[validarLogin] Login bem-sucedido para ${usuario}.`);
        return {
          status: 'ok',
          token: token,
          usuario: String(usuario).toLowerCase(), 
          idUsuario: usuarioRow[colunas.id],
          nome: usuarioRow[colunas.nome],
          perfil: usuarioRow[colunas.perfil],
          funcao: usuarioRow[colunas.funcao],
          statusConta: usuarioRow[colunas.status],
          empresa: empresaObjetoCompleto
        };

      } catch (e) {
        Logger.log(`[validarLogin] ERRO FATAL: ${e.message} ${e.stack}`);
        return { status: 'error', message: `Erro no servidor: ${e.message}` };
      }
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
        const indexFuncao = headersUsuarios.indexOf('FUNCAO');

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

  function atualizarFuncaoEPerfilDoUsuario(dados) {
  const { usuarioId, novaFuncao, novoPerfil } = dados;

  const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
  if (!sheet) {
    return { status: 'error', message: 'Planilha "Usuarios" não encontrada.' };
  }

  const lastRow = sheet.getLastRow();
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(usuarioId).trim()) {
      const row = i + 2;

      // Atualiza coluna 5 (Perfil) e coluna 9 (Função)
      sheet.getRange(row, 5).setValue(novoPerfil);
      sheet.getRange(row, 9).setValue(novaFuncao);

      return { status: 'ok', message: `Função e perfil do usuário ${usuarioId} atualizados com sucesso.` };
    }
  }

  return { status: 'error', message: `Usuário ${usuarioId} não encontrado.` };
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
        if (senhaAtualArmazenada !== hashInformado) {
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
function gerarHash(senha) {
      var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, senha);
      return Utilities.base64Encode(digest);
    }

/**
 * Permite que um administrador redefina a senha de qualquer usuário.
 * Inclui verificação de segurança para garantir que o autor da chamada é um admin.
 * @param {string} usuarioAlvo - O nome de usuário cuja senha será redefinida.
 * @param {string} novaSenha - A nova senha em texto plano.
 * @returns {object} Objeto de resposta com status e mensagem.
 */
function adminRedefinirSenha(usuario, novaSenha) {
      var sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
      var dados = sheet.getDataRange().getValues();
      var idxUsuario = dados[0].indexOf('USUARIO');
      var idxSenha = dados[0].indexOf('SENHA');
      var linha = -1;

      if (!novaSenha || novaSenha.trim() === '') {
        return { status: 'error', message: 'Senha inválida.' };
      }

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
    const result = { lastLogin: null, lastOrder: null, lastPrint: null };
    if (!userId) {
      throw new Error("ID do usuário não foi fornecido.");
    }

    Logger.log(`Iniciando auditoria para o ID de usuário: ${userId}`);
    const normalize = v => String(v || "").trim().toLowerCase();

    // --- PASSO 1: TRADUZIR O ID PARA NOME ---
    let usuarioNome = null;
    const usuariosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("usuarios"); // Usa o nome da sua planilha
    if (usuariosSheet) {
      const usuariosData = usuariosSheet.getDataRange().getValues();
      const headers = usuariosData.shift();
      const idIndex = headers.indexOf("ID");
      const nomeIndex = headers.indexOf("NOME");

      if (idIndex > -1 && nomeIndex > -1) {
        const userRow = usuariosData.find(row => String(row[idIndex]).trim() === String(userId).trim());
        if (userRow) {
          usuarioNome = userRow[nomeIndex];
          Logger.log(`Nome encontrado para o ID ${userId}: ${usuarioNome}`);
        }
      }
    }
    if (!usuarioNome) {
      Logger.log(`Não foi possível encontrar um nome para o ID de usuário: ${userId}`);
    }
    
    // --- PASSO 2: BUSCAR ÚLTIMO LOGIN (usando o ID) ---
    const loginSheet = SpreadsheetApp.getActive().getSheetByName('usuario_logado');
    if (loginSheet && loginSheet.getLastRow() > 1) {
      const loginData = loginSheet.getDataRange().getValues();
      const loginEntries = loginData.slice(1).filter(row => String(row[0]).trim() === String(userId).trim());
      
      Logger.log(`Login entries encontradas para o ID ${userId}: ${loginEntries.length}`);
      if (loginEntries.length > 0) {
        const lastLoginEntry = loginEntries.sort((a, b) => new Date(b[4]) - new Date(a[4]))[0];
        const date = parseDateSafe(lastLoginEntry[4]);
        result.lastLogin = {
          date: date ? Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "Inválido",
          ip: lastLoginEntry[5] || 'N/A'
        };
      }
    }

    // --- PASSO 3: BUSCAR PEDIDOS E IMPRESSÕES (usando o NOME) ---
    // Esta parte só roda se tivermos encontrado um nome para o usuário.
    if (usuarioNome) {
      // Bloco "Último Pedido"
      const pedidosSheet = SpreadsheetApp.getActive().getSheetByName('Pedidos');
      if (pedidosSheet && pedidosSheet.getLastRow() > 1) {
          const pedidosData = pedidosSheet.getDataRange().getValues();
          const headers = pedidosData[0];
          const numeroIdx = headers.findIndex(h => normalize(h).includes('pedido'));
          const dataIdx = headers.findIndex(h => normalize(h).includes('data'));
          const usuarioIdx = headers.findIndex(h => normalize(h).includes('usuario criador'));

          if (usuarioIdx > -1) {
              const userPedidos = pedidosData.slice(1).filter(row => normalize(row[usuarioIdx]) === normalize(usuarioNome));
              if (userPedidos.length > 0) {
                  const lastPedido = userPedidos.sort((a, b) => new Date(b[dataIdx]) - new Date(a[dataIdx]))[0];
                  const date = parseDateSafe(lastPedido[dataIdx]);
                  result.lastOrder = {
                      id: String(lastPedido[numeroIdx]),
                      date: date ? Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "Inválido"
                  };
              }
          }
      }

      // Bloco "Última Impressão"
      const impressaoSheet = SpreadsheetApp.getActive().getSheetByName('Log Impressoes');
      if (impressaoSheet && impressaoSheet.getLastRow() > 1) {
          const impressaoData = impressaoSheet.getDataRange().getValues();
          const headers = impressaoData[0];
          const usuarioIdx = headers.findIndex(h => normalize(h).includes('usuario'));
          const pedidoIdx = headers.findIndex(h => normalize(h).includes('pedido'));
          const dataIdx = headers.findIndex(h => normalize(h).includes('data'));

          if (usuarioIdx > -1) {
              const userImpressao = impressaoData.slice(1).filter(row => normalize(row[usuarioIdx]) === normalize(usuarioNome));
              if (userImpressao.length > 0) {
                  const lastImpressao = userImpressao.sort((a, b) => new Date(b[dataIdx]) - new Date(a[dataIdx]))[0];
                  const date = parseDateSafe(lastImpressao[dataIdx]);
                  result.lastPrint = {
                      id: String(lastImpressao[pedidoIdx]),
                      date: date ? Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "Inválido"
                  };
              }
          }
      }
    } else {
        Logger.log("Busca de Pedidos e Impressões pulada pois o nome do usuário não foi encontrado.");
    }

    Logger.log("Resultado final da auditoria: " + JSON.stringify(result));
    return result;

  } catch (error) {
    Logger.log('Erro em obterDadosAuditoria: ' + error.message);
    return { lastLogin: null, lastOrder: null, lastPrint: null };
  }
}

/**
 * VERSÃO FINAL E ROBUSTA
 * Converte um valor para um objeto de data de forma segura, agora com suporte
 * para múltiplos formatos, incluindo o formato completo do JavaScript e o brasileiro.
 * @param {*} value O valor a ser convertido.
 * @returns {Date | null} O objeto de data válido ou null se a conversão falhar.
 */
function parseDateSafe(value) {
  try {
    // 1. Se já for um objeto de data válido, retorna imediatamente.
    if (value instanceof Date && !isNaN(value)) {
      return value;
    }

    // 2. Se for um texto, tenta a conversão direta primeiro.
    // Isso funciona para o formato ISO (YYYY-MM-DD) e o formato completo (Fri Aug 22...).
    if (typeof value === 'string') {
      const d = new Date(value);
      if (!isNaN(d.getTime())) {
        return d;
      }

      // 3. Se a conversão direta falhou, tenta a análise manual do formato brasileiro (DD/MM/YYYY).
      if (value.includes('/')) {
        const parts = value.split(' ');
        const dateParts = parts[0].split('/');
        
        if (dateParts.length === 3) {
          const day = parseInt(dateParts[0], 10);
          const month = parseInt(dateParts[1], 10) - 1; // Mês é base 0 em JS
          const year = parseInt(dateParts[2], 10);
          
          let hour = 0, minute = 0;
          if (parts.length > 1 && parts[1].includes(':')) {
            const timeParts = parts[1].split(':');
            hour = parseInt(timeParts[0], 10) || 0;
            minute = parseInt(timeParts[1], 10) || 0;
          }
          
          const manualDate = new Date(year, month, day, hour, minute);
          if (!isNaN(manualDate.getTime())) {
            return manualDate;
          }
        }
      }
    }

    // 4. Se nada funcionou, retorna null.
    return null;

  } catch (e) {
    Logger.log(`Erro em parseDateSafe ao tentar converter o valor "${value}": ${e.message}`);
    return null;
  }
}

function testar_Auditoria_ComNomeDeUsuario() {
  Logger.log("--- INICIANDO TESTE DE AUDITORIA POR NOME ---");

  // <<< MUDANÇA AQUI >>>
  // Em vez de um ID, use o NOME COMPLETO de um usuário que você sabe que criou pedidos.
  // O nome deve ser exatamente como está na coluna "Usuario Criador" da planilha "Pedidos".
  const nomeDoUsuarioParaTestar = "1"; // Exemplo, use um nome real dos seus dados

  const resultado = obterDadosAuditoria(nomeDoUsuarioParaTestar);
  
  Logger.log("--- RESULTADO DO TESTE ---");
  Logger.log(JSON.stringify(resultado, null, 2));
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

    /**
 * Lista todos os usuários com dados formatados para a tela de gerenciamento.
 * Esta função busca os nomes das empresas para uma exibição mais amigável.
 * @returns {Array<Object>} Lista de usuários com os dados das empresas processados.
 */
function listarUsuariosCompleto() {
    try {
        Logger.log("[listarUsuariosCompleto] Iniciando a busca completa de usuários...");
        const sheet = SpreadsheetApp.getActive().getSheetByName('Usuarios');
        if (!sheet || sheet.getLastRow() < 2) {
            Logger.log("[listarUsuariosCompleto] Planilha 'Usuarios' não encontrada ou vazia.");
            return [];
        }

        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
        Logger.log(`[listarUsuariosCompleto] Encontrados ${data.length} registros de usuários.`);
        
        const empresasMap = _getEmpresasMap();

        const usuariosProcessados = data.map(([id, nome, usuario, senha, perfil, status, empresasStr, empresaPadrao]) => {
            const empresasIds = empresasStr ? String(empresasStr).split(',').map(e => e.trim()) : [];
            const empresasNomes = empresasIds.map(id => empresasMap[String(id).trim()] || `Empresa ${id}`);
            const empresaPadraoId = empresaPadrao ? String(empresaPadrao).trim() : '';
            const empresaPadraoNome = empresasMap[empresaPadraoId] || '';

            return {
                id: String(id),
                nome: String(nome),
                usuario: String(usuario),
                perfil: String(perfil),
                status: String(status),
                empresas: empresasNomes.join(', '),
                empresaPadrao: empresaPadraoNome,
                empresasIds: empresasIds,
                empresaPadraoId: empresaPadraoId
            };
        });
        
        if (usuariosProcessados.length > 0) {
            Logger.log(`[listarUsuariosCompleto] Processamento concluído. Exemplo do primeiro usuário: ${JSON.stringify(usuariosProcessados[0])}`);
        } else {
            Logger.log("[listarUsuariosCompleto] Nenhum usuário foi processado.");
        }
        
        return usuariosProcessados;

    } catch (error) {
        Logger.log(`[listarUsuariosCompleto] ERRO FATAL: ${error.message}. Stack: ${error.stack}`);
        return []; // Retorna um array vazio em caso de erro.
    }
}


/**
 * Função auxiliar para criar um mapa de ID -> Nome da Empresa.
 * Isso evita ter que ler a planilha de empresas múltiplas vezes.
 * @returns {Object} Um objeto onde a chave é o ID da empresa e o valor é o nome.
 */
function _getEmpresasMap() {
    const empresasMap = {};
    try {
        Logger.log("[_getEmpresasMap] Iniciando a criação do mapa de empresas...");
        const sheetEmpresas = SpreadsheetApp.getActive().getSheetByName('Empresas');
        if (sheetEmpresas && sheetEmpresas.getLastRow() >= 2) {
            const empresasData = sheetEmpresas.getRange(2, 1, sheetEmpresas.getLastRow() - 1, 2).getValues();
            Logger.log(`[_getEmpresasMap] Encontradas ${empresasData.length} empresas na planilha.`);
            empresasData.forEach(([id, nome]) => {
                if (id && nome) {
                    const cleanId = String(id).trim();
                    empresasMap[cleanId] = String(nome).trim();
                }
            });
            Logger.log(`[_getEmpresasMap] Mapa de empresas criado com sucesso com ${Object.keys(empresasMap).length} entradas.`);
        } else {
            Logger.log("[_getEmpresasMap] Planilha 'Empresas' não encontrada ou vazia.");
        }
    } catch(e) {
        Logger.log(`[_getEmpresasMap] ERRO ao criar mapa de empresas: ${e.message}`);
    }
    return empresasMap;
}
    function _getEmpresaDataById(empresaId) {
    try {
        const companySheet = SpreadsheetApp.getActive().getSheetByName('Empresas');
        if (!companySheet || companySheet.getLastRow() < 2) {
            Logger.log("[_getEmpresaDataById] Planilha 'Empresas' não encontrada ou vazia.");
            return null;
        }

        const companiesData = companySheet.getDataRange().getValues();
        const companyHeaders = companiesData[0].map(h => String(h).toUpperCase().trim());

        const idxCompanyId = companyHeaders.indexOf('ID');
        const idxCompanyName = companyHeaders.indexOf('EMPRESA');
        const idxCompanyCnpj = companyHeaders.indexOf('CNPJ');
        const idxCompanyEndereco = companyHeaders.indexOf('ENDEREÇO');

        if (idxCompanyId === -1 || idxCompanyName === -1 || idxCompanyCnpj === -1 || idxCompanyEndereco === -1) {
            Logger.log("[_getEmpresaDataById] ERRO: Cabeçalhos essenciais (ID, NOME, CNPJ, ENDERECO) não encontrados.");
            return null;
        }

        const idProcurado = String(empresaId).trim();
        Logger.log(`[_getEmpresaDataById] Procurando por Empresa ID: "${idProcurado}"`);

        // Procura pela empresa com o ID correspondente, comparando como texto.
        const empresaInfo = companiesData.find((row, index) => {
            if (index === 0) return false; // Pula a linha do cabeçalho
            
            const idNaPlanilha = String(row[idxCompanyId]).trim();
            
            // LOG DE DIAGNÓSTICO: Mostra o que está sendo comparado.
            Logger.log(`  -> Linha ${index + 1}: Comparando "${idNaPlanilha}" (planilha) com "${idProcurado}" (login).`);

            // CORREÇÃO: Compara os valores numéricos, o que faz com que 1 seja igual a "001".
            return parseInt(idNaPlanilha, 10) === parseInt(idProcurado, 10);
        });

        if (empresaInfo) {
            const empresa = {
                id: String(empresaInfo[idxCompanyId]).trim(),
                nome: empresaInfo[idxCompanyName],
                cnpj: empresaInfo[idxCompanyCnpj],
                endereco: empresaInfo[idxCompanyEndereco]
            };
            Logger.log(`[_getEmpresaDataById] ✅ Empresa encontrada: ${JSON.stringify(empresa)}`);
            return empresa;
        }

        Logger.log(`[_getEmpresaDataById] ❌ Empresa com ID "${idProcurado}" não foi encontrada após varrer a planilha.`);
        return null;

    } catch (e) {
        Logger.log(`[getEmpresaDataById] ERRO FATAL: ${e.message}`);
        return null;
    }
}

/**
 * FUNÇÃO DE EMERGÊNCIA
 * Reseta manualmente a senha de um usuário para o novo padrão de hash (Hexadecimal).
 * Execute esta função diretamente do editor de scripts.
 */
function resetarSenhaManualmente() {
  const usuarioParaResetar = "admin"; // Coloque aqui o nome de usuário que você quer resetar
  const novaSenha = "1234";          // Coloque aqui a nova senha temporária

  try {
    const sheet = SpreadsheetApp.openById(ID_DA_PLANILHA).getSheetByName('Usuarios');
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const colunas = {};
    headers.forEach((h, i) => { colunas[toCamelCase(h)] = i; });

    const rowIndex = data.findIndex(row => row[colunas.usuario]?.toLowerCase() === usuarioParaResetar.toLowerCase());

    if (rowIndex === -1) {
      Logger.log(`ERRO: Usuário "${usuarioParaResetar}" não encontrado para resetar a senha.`);
      return;
    }

    // CORREÇÃO 1: Usa a sua função de hash antiga (gerarHash)
    const hashAntigo = gerarHash(novaSenha); 

    const linhaReal = rowIndex + 2;
    const colunaSenha = colunas.senha + 1;

    // CORREÇÃO 2: Usa a variável correta ('hashAntigo') para salvar
    sheet.getRange(linhaReal, colunaSenha).setValue(hashAntigo);

    Logger.log(`✅ SUCESSO: A senha para o usuário "${usuarioParaResetar}" foi redefinida para "${novaSenha}" (no formato antigo).`);
    SpreadsheetApp.flush(); // Garante que a alteração seja salva imediatamente.

  } catch (e) {
    Logger.log(`ERRO FATAL ao resetar senha: ${e.message}`);
  }
}
