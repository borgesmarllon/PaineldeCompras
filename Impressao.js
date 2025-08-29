//const ID_DA_PLANILHA = "1xVLFSqL5SVT6cmZ_9foOkKxJIHqHeCGHLcyryBJ44g0"

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
 * SUBSTITUIÇÃO UNIFICADA E SEGURA para as funções de impressão.
 * Busca um pedido pelo seu número E ID da empresa, e enriquece o objeto
 * com os dados cadastrais completos da empresa e do fornecedor.
 * @param {string} numeroPedido O número do pedido a ser buscado.
 * @param {string} empresaId O ID da empresa do pedido.
 * @returns {Object|null} Um objeto completo com todos os dados para impressão ou null se não encontrado.
 
function getPedidoCompletoPorId(numeroPedido, empresaId) {
  try {
    // 1. BUSCA OS DADOS DO PEDIDO
    const pedidoData = buscarPedidoPorId(numeroPedido, empresaId);
    if (!pedidoData) {
      throw new Error(`Pedido ${numeroPedido} da empresa ${empresaId} não encontrado.`);
    }

    // 2. BUSCA OS DADOS DA EMPRESA E ANEXA AO PEDIDO
    const dadosEmpresa = getDadosEmpresaPorId(empresaId);
    pedidoData.empresaInfo = dadosEmpresa || {};

    // 3. BUSCA OS DADOS DO FORNECEDOR E ANEXA AO PEDIDO
    const dadosFornecedor = getDadosFornecedorPorNome(pedidoData.fornecedor);
    pedidoData.fornecedorInfo = dadosFornecedor || {};

    Logger.log(`Pedido ${numeroPedido} encontrado e enriquecido com sucesso.`);
    return pedidoData;

  } catch (e) {
    Logger.log(`ERRO em getPedidoCompletoPorId: ${e.message}`);
    return null;
  }
}*/

function buscarPedidoPorId(numeroPedido, empresaId) {
  // Esta função agora abre sua própria conexão
  const planilha = SpreadsheetApp.openById(ID_DA_PLANILHA);
  const sheet = planilha.getSheetByName('Pedidos'); // Pega a aba
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const colunas = {};
  headers.forEach((header, index) => { colunas[toCamelCase(header)] = index; });

  const pedidoRow = data.find(row => 
    String(row[colunas.numeroDoPedido]).trim() === String(numeroPedido).trim() &&
    String(row[colunas.empresa]).trim() === String(empresaId).trim()
  );

  if (!pedidoRow) return null;

  const pedidoData = {};
  for (const key in colunas) {
    if (colunas.hasOwnProperty(key)) {
      let value = pedidoRow[colunas[key]];
      if (value instanceof Date) {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
      }
      pedidoData[key] = value;
    }
  }
  
  pedidoData.itens = JSON.parse(pedidoData.itens || '[]');
  return pedidoData;
}


/**
 * Função auxiliar para converter "Nome Cabeçalho" em "nomeCabecalho".
 * @param {string} text - O texto a ser convertido.
 * @returns {string} O texto em camelCase.
 */
function toCamelCase(text) {
  if (!text) return '';
  // Remove acentos, caracteres especiais, e depois converte para camelCase
  const a = 'àáâäæãåāăąçćčđďèéêëēėęěğǵḧîïíīįìłḿñńǹňôöòóœøōõőṕŕřßśšşșťțûüùúūǘůűųẃẍÿýžźż·/_,:;'
  const b = 'aaaaaaaaaacccddeeeeeeeegghiiiiiilmnnnnoooooooooprrssssssttuuuuuuuuuwxyyzzz------'
  const p = new RegExp(a.split('').join('|'), 'g')

  return text.toString().toLowerCase()
    .replace(/\s+/g, '-') // substitui espaços por -
    .replace(p, c => b.charAt(a.indexOf(c))) // substitui caracteres especiais
    .replace(/&/g, '-e-') // substitui & por 'e'
    .replace(/[^\w\-]+/g, '') // remove caracteres inválidos
    .replace(/\-\-+/g, '-') // substitui múltiplos - por um único -
    .replace(/^-+/, '') // remove - do início
    .replace(/-+$/, '') // remove - do final
    .replace(/-(\w)/g, (match, R) => R.toUpperCase()); // Converte para camelCase
}

/**
 * Busca os dados cadastrais de uma empresa específica pelo seu ID.
 * @param {string} empresaId O ID da empresa a ser buscada (ex: "001").
 * @returns {object|null} Um objeto com os dados da empresa ou null se não for encontrada.
 */
function getDadosEmpresaPorId(empresaId) {
  // CORREÇÃO DO ERRO DE DIGITAÇÃO E LÓGICA DE COMPARAÇÃO
  try {
    const planilha = SpreadsheetApp.openById(ID_DA_PLANILHA);
    const sheet = planilha.getSheetByName('Empresas'); // Corrigido de getSheetByNem
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const colunas = {};
    headers.forEach((h, i) => colunas[toCamelCase(h)] = i);

    if (colunas.id === undefined) return null;

    // --- LÓGICA DE COMPARAÇÃO ROBUSTA ---
    // Converte ambos os lados para número antes de comparar.
    // Assim, 1 (da planilha) será igual a "001" (do front-end).
    const empresaRow = data.find(row => 
        parseInt(row[colunas.id], 10) === parseInt(empresaId, 10)
    );
    // --- FIM DA MELHORIA ---

    if (empresaRow) {
      const empresaData = {};
      for (const key in colunas) {
          if (colunas.hasOwnProperty(key)) {
              empresaData[key] = empresaRow[colunas[key]];
          }
      }
      // Garante que o ID retornado seja sempre uma string com zeros à esquerda
      empresaData.id = String(empresaData.id).padStart(3, '0');
      return empresaData;
    }
    return null;
  } catch(e) {
    Logger.log(`ERRO em getDadosEmpresaPorId: ${e.message}`);
    return null;
  }
}

function getDadosFornecedorPorNome(nomeFornecedor) {
    const planilha = SpreadsheetApp.openById(ID_DA_PLANILHA); // Abre o arquivo
    const sheet = planilha.getSheetByName('Fornecedores'); // Pega a aba
    if (!sheet || !nomeFornecedor) return null;

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const colunas = {};
    headers.forEach((h, i) => colunas[toCamelCase(h)] = i);

    // Garante que a coluna com o nome (ex: razaoSocial) existe
    if (colunas.razaoSocial === undefined) return null; 

    const nomeFornecedorNormalizado = String(nomeFornecedor).trim().toUpperCase();
    const fornecedorRow = data.find(row => 
        String(row[colunas.razaoSocial]).trim().toUpperCase() === nomeFornecedorNormalizado
    );

    if (fornecedorRow) {
        const fornecedorData = {};
        for (const key in colunas) {
            if (colunas.hasOwnProperty(key)) {
                fornecedorData[key] = fornecedorRow[colunas[key]];
            }
        }
        return fornecedorData;
    }
    return null;
}


/**
 * Calcula a distância de Levenshtein entre duas strings.
 * É uma medida da diferença entre duas sequências de caracteres.
 * @param {string} a A primeira string.
 * @param {string} b A segunda string.
 * @returns {number} A distância (número de edições).
 */
function levenshteinDistance(a, b) {
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  const matrix = [];

  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }

  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(matrix[i - 1][j - 1] + 1, Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1));
      }
    }
  }
  return matrix[b.length][a.length];
}

// ID DA PASTA DO DRIVE QUE SALVA O ARQUIVO PDF
const ID_PASTA_PDFS = '1t7mQk5pY1g-Gxl4kFT_1sy-R0RYty6Do'
// ====================================================
/**
 * FUNÇÃO PRINCIPAL DA IMPRESSÃO
 * Orquestra a busca de dados, criação do HTML e geração do PDF.
 * @param {string} numeroPedido - O número do pedido a ser impresso.
 * @param {string} empresaId - O ID da empresa do pedido.
 * @returns {object} Um objeto com o status e a URL do PDF gerado.
 */
function gerarPdfPedido(numeroPedido, empresaId, usuarioLogado, nomeUsuario) {
  try {
    Logger.log(`Iniciando geração de PDF para pedido ${numeroPedido}, empresa ${empresaId}`);
     Logger.log("Parâmetros recebidos:");
  Logger.log("Número do Pedido: " + numeroPedido);
  Logger.log("Empresa: " + empresaId);
  Logger.log("Usuário logado: " + usuarioLogado);
  Logger.log("Nome do usuário: " + nomeUsuario);
    // 1. Obtem os dados completos do pedido para saber seu status
    const pedidoCompleto = getPedidoCompletoPorId(numeroPedido, empresaId);
    if (!pedidoCompleto) {
      throw new Error("pedido não encontrado no servidor.");
    }

    const nomeArquivo = `Pedido_${numeroPedido}_${empresaId}.pdf`;
    const pastaDestino = DriveApp.getFolderById(ID_PASTA_PDFS);
    const arquivosExistentes = pastaDestino.getFilesByName(nomeArquivo);
    const statusPedido = (pedidoCompleto.status || '').toUpperCase();

    let pdfFile, fileId;

    // 2. Verifica se o arquivo já existe para evitar recriação
    if (statusPedido === 'Cancelado' || statusPedido === 'RASCUNHO') {
      if (arquivosExistentes.hasNext()) {
        const arquivoAntigo = arquivosExistentes.next();
        Logger.log(`Pedido com status '${statusPedido}'. Removendo PDF antigo para gerar um novo com marca d'água.`);
        arquivoAntigo.setTrashed(true); // Envia para lixeira, mais seguro que exclusão permanente.
      }
    } else {
    if (arquivosExistentes.hasNext()) {
       pdfFile = arquivosExistentes.next();
      fileId = pdfFile.getId();
      Logger.log(`PDF já existe. Retornando URLs: download=${fileId}, visualização=${pdfFile.getUrl()}`);
      return { 
      status: 'ok', 
      pdfUrl: `https://drive.google.com/uc?export=download&id=${fileId}`, // Para Telegram
      pdfViewUrl: pdfFile.getUrl() // Para web
    };
  }
} 

    // Se não existir, prossegue com a criação
    Logger.log("Nenhum PDF existente encontrado. Gerando um novo arquivo.");
    
    // 3. Construir o HTML para o PDF
    const htmlParaPdf = construirHtmlParaPdf(pedidoCompleto);

    // 4. Criar o blob do PDF a partir do HTML
    const pdfBlob = Utilities.newBlob(htmlParaPdf, 'text/html', `Pedido_${numeroPedido}.html`)
                             .getAs('application/pdf');
    pdfBlob.setName(nomeArquivo);

    // 4. Salvar o PDF na pasta do Google Drive
    pdfFile = pastaDestino.createFile(pdfBlob);

    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileId = pdfFile.getId();

    registrarImpressao(usuarioLogado, nomeUsuario, numeroPedido, 'Impressora1');
    // 5. Retornar a URL do arquivo para o front-end
        Logger.log(`PDF do pedido ${numeroPedido} gerado com sucesso: ${pdfFile.getUrl()}`);
        return {status: 'ok',
          pdfUrl: `https://drive.google.com/uc?export=download&id=${fileId}`, // Para Telegram
          pdfViewUrl: pdfFile.getUrl() // Para web
        };

      } catch (error) {
        Logger.log(`ERRO ao gerar PDF para o pedido ${numeroPedido}: ${error.message}\nStack: ${error.stack}`);
        return { status: 'error', message: error.message };
      }
      
}

/**
 * FUNÇÃO DE DADOS
 * Busca todos os dados de um pedido, sua empresa e fornecedor.
 * (Esta é uma implementação robusta, verifique se os nomes das abas e colunas batem com os seus)
 */
function getPedidoCompletoPorId(numeroPedido, empresaId) {
  try {
    Logger.log(`Buscando pedido. Parâmetros recebidos: numeroPedido='${numeroPedido}', empresaId='${empresaId}'`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pedidosSheet = ss.getSheetByName("Pedidos");
    if (!pedidosSheet) throw new Error("Planilha 'Pedidos' não encontrada.");

    // 1. Criar os mapas de busca para performance
    const mapaEmpresas = _criarMapaDeEmpresas();
    const mapaFornecedores = criarMapaDeFornecedoresv2();    
    const mapaDeUsuarios = _criarMapaDeUsuarios();

    // 2. Encontrar a linha do pedido
    const pedidosData = pedidosSheet.getDataRange().getValues();
    const pedidoHeaders = pedidosData.shift();

    Logger.log(`Cabeçalhos da planilha 'Pedidos': [${pedidoHeaders.join(", ")}]`);
    const colunas = {
        numero: pedidoHeaders.indexOf("Número do Pedido"),
        empresa: pedidoHeaders.indexOf("Empresa"),
        fornecedor: pedidoHeaders.indexOf("Fornecedor"),
        usuario: pedidoHeaders.indexOf("Usuario Criador"),
        status: pedidoHeaders.indexOf("Status")
    };
    Logger.log(`Índices das colunas encontrados: Numero=${colunas.numero}, Empresa=${colunas.empresa}, Usuario=${colunas.usuario}`);

    if (colunas.numero === -1 || colunas.empresa === -1) {
        throw new Error("Não foi possível encontrar as colunas 'Número do Pedido' e/ou 'Empresa' na planilha. Verifique os nomes dos cabeçalhos.");
    }

    let logCount = 0;
    const pedidoRow = pedidosData.find(row => {
        const numeroNaPlanilha = String(row[colunas.numero]).trim();
        const empresaNaPlanilha = String(row[colunas.empresa]).trim();
        const numeroBuscado = String(numeroPedido).trim();
        const empresaBuscada = String(empresaId).trim();

        // Adiciona um log detalhado para as primeiras 5 comparações para ajudar a depurar
        if (logCount < 5) {
            Logger.log(`Comparando: (Planilha) Num='${numeroNaPlanilha}', Emp='${empresaNaPlanilha}' vs (Buscado) Num='${numeroBuscado}', Emp='${empresaBuscada}'`);
            logCount++;
        }
        
        return numeroNaPlanilha === numeroBuscado && empresaNaPlanilha === empresaBuscada;
    });

    if (!pedidoRow) {
      Logger.log(`Pedido ${numeroPedido} da empresa ${empresaId} não encontrado.`);
      return null;
    }

    // 3. Montar o objeto base do pedido
    const pedidoCompleto = {};
    pedidoHeaders.forEach((header, index) => {
        const key = header.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '_').toLowerCase();
        pedidoCompleto[key] = pedidoRow[index];
    });
    Logger.log(`[DEPURAÇÃO] Chaves criadas para o objeto pedidoCompleto: ${Object.keys(pedidoCompleto).join(', ')}`);
    pedidoCompleto.empresaId = pedidoCompleto.empresa;
    // 4. Anexar informações da empresa e fornecedor usando os mapas
    pedidoCompleto.empresaInfo = mapaEmpresas[pedidoCompleto.empresa] || {};
    
    const nomeFornecedorNormalizado = (pedidoCompleto.fornecedor || '').trim().toUpperCase();
    pedidoCompleto.fornecedorInfo = mapaFornecedores[nomeFornecedorNormalizado] || {};

    const idUsuarioCriador = pedidoCompleto.usuario_criador;
    pedidoCompleto.usuarioCriadorInfo = mapaDeUsuarios[idUsuarioCriador] || {};

    // 5. Garantir que os itens sejam um objeto
    try {
      pedidoCompleto.itens = JSON.parse(pedidoCompleto.itens);
    } catch(e) {
      pedidoCompleto.itens = [];
    }
   
    Logger.log("Pedido completo encontrado e montado com sucesso.");
    Logger.log(`[DEPURAÇÃO FINAL] Objeto 'pedidoCompleto' final: ${JSON.stringify(pedidoCompleto, null, 2)}`);
    return pedidoCompleto;

  } catch (e) {
    Logger.log(`ERRO em getPedidoCompletoPorId: ${e.stack}`);
    return null;
  }
}

function testarGetPedidoCompleto() {
  Logger.log("--- INICIANDO TESTE: getPedidoCompletoPorId ---");

  // 1. Defina os parâmetros de um pedido que você sabe que existe.
  const numeroPedidoTeste = "001422"; // <-- COLOQUE UM NÚMERO DE PEDIDO REAL AQUI
  const empresaIdTeste = "001";      // <-- COLOQUE O ID DA EMPRESA CORRETO AQUI

  try {
    // ==========================================================
    // <<< A CORREÇÃO ESTÁ AQUI >>>
    // Garantimos que estamos passando as variáveis corretas para a função.
    // ==========================================================
    const pedidoCompleto = getPedidoCompletoPorId(numeroPedidoTeste, empresaIdTeste);
    
    if (pedidoCompleto) {
      Logger.log("✅ Pedido encontrado! Verificando as propriedades do objeto...");
      Logger.log(`[DEPURAÇÃO] Chaves criadas: ${Object.keys(pedidoCompleto).join(', ')}`);
    } else {
      Logger.log("❌ Pedido não encontrado. Verifique se o número e o ID da empresa estão corretos no teste E na planilha.");
    }

    Logger.log("--- FIM DO TESTE ---");
    
  } catch (e) {
    Logger.log("!!! O TESTE FALHOU COM UM ERRO CRÍTICO !!!");
    Logger.log("Erro: " + e.message);
  }
}
/**
 * FUNÇÃO AUXILIAR
 * Cria um mapa de busca rápida para todas as empresas.
 * @returns {object} Um objeto onde a chave é o ID da empresa.
 */
function _criarMapaDeEmpresas() {
    const empresasSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Empresas");
    if (!empresasSheet) return {};
    
    const data = empresasSheet.getDataRange().getValues();
    const headers = data.shift();
    const mapa = {};
    
    data.forEach(row => {
        const empresaObj = {};
        headers.forEach((header, index) => {
            const key = header.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '_').toLowerCase();
            empresaObj[key] = row[index];
        });
        const id = String(empresaObj.id || empresaObj.codigo).trim();
        if (id) mapa[id] = empresaObj;
    });
    return mapa;
}


function _criarMapaDeUsuarios() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    if (!sheet) return {};
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const mapa = {};
    
    data.forEach(row => {
        const userObj = {};
        headers.forEach((header, index) => {
            const key = header.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, '_').toLowerCase();
            userObj[key] = row[index];
        });
        const login = String(userObj.usuario || userObj.login).trim().toLowerCase();
        if (login) mapa[login] = userObj;
    });
    return mapa;
}

/**
 * Cria um mapa de fornecedores usando a RAZÃO SOCIAL (NOME) como chave.
 * Esta versão é compatível com o resto do sistema que busca pelo nome.
 */
function criarMapaDeFornecedoresv2() {
    try {
        const config = getConfig();
        const fornecedoresSheet = SpreadsheetApp.getActive().getSheetByName(config.sheets.fornecedores);
        if (!fornecedoresSheet) {
            Logger.log("ERRO CRÍTICO: A aba de fornecedores não foi encontrada.");
            return {}; // Retorna um objeto vazio para segurança.
        }

        const data = fornecedoresSheet.getDataRange().getValues();
        const headers = data.shift(); // Pega a primeira linha (cabeçalhos) para referência

        // Encontra o índice das colunas dinamicamente (muito mais seguro que números fixos)
        // **Atenção:** Adapte os nomes "Nome" e "Cidade" se os cabeçalhos na sua planilha forem diferentes.
        const nomeIndex = headers.indexOf("RAZAO SOCIAL"); 
        const cidadeIndex = headers.indexOf("CIDADE");
        const cnpjIndex = headers.indexOf("CNPJ");
        const enderecoIndex = headers.indexOf("ENDERECO"); // Exemplo
        const grupoIndex = headers.indexOf("GRUPO");
        const condicaoIndex = headers.indexOf("CONDICAO DE PAGAMENTO");
        const formaIndex = headers.indexOf("FORMA DE PAGAMENTO");
        const estadoIndex = headers.indexOf("ESTADO");

        const statusIndex = headers.indexOf("STATUS");
        if (nomeIndex === -1) {
            Logger.log("ERRO CRÍTICO: Não foi possível encontrar a coluna 'Nome' na aba 'Fornecedores'.");
            return {};
        }

        const mapa = {}; // <-- Usamos um objeto simples {}, não new Map()

        data.forEach(row => {
          const status = (statusIndex !== -1) ? String(row[statusIndex] || '').toUpperCase().trim() : 'ATIVO';
            
            // --- PASSO 3: Crie o mapa APENAS SE o status for 'ATIVO' ---
            if (status === 'ATIVO') {
            // A CHAVE do mapa será o nome do fornecedor, normalizado para evitar erros.
            const nomeFornecedor = String(row[nomeIndex] || '').trim().toUpperCase();

            // Só adiciona ao mapa se a linha tiver um nome válido.
            if (nomeFornecedor) {
                // O VALOR do mapa é um objeto com todos os dados do fornecedor.
                mapa[nomeFornecedor] = {
                    nome: row[nomeIndex], // Guardamos o nome original
                    cidade: row[cidadeIndex] || '',
                    cnpj: row[cnpjIndex] || '',
                    endereco: row[enderecoIndex] || '',
                    grupo: row[grupoIndex] || '',
                    condicao: row[condicaoIndex] || '',
                    forma: row[formaIndex] || '',
                    estado: row[estadoIndex] || '',
                    // Adicione aqui os outros campos que você precisar, usando os índices.
                };
            }
            }
        });
        
        Logger.log(`Mapa de fornecedores (por NOME) criado com sucesso. Total de entradas: ${Object.keys(mapa).length}`);
        Logger.log("Dados completos do mapa:");
        Logger.log(JSON.stringify(mapa, null, 2)); // Mostra o objeto em formato legível
        return mapa;

    } catch (e) {
        Logger.log(`ERRO FATAL em criarMapaDeFornecedoresv2: ${e.message}`);
        return {}; // Retorna um objeto vazio em caso de erro.
    }
}

/**
 * FUNÇÃO DE LAYOUT
 * Constrói a string HTML para o documento de impressão, imitando o seu layout.
 * @param {object} pedidoCompleto - O objeto completo do pedido.
 * @returns {string} Uma string contendo todo o HTML do documento.
 */
function construirHtmlParaPdf(pedidoCompleto) {
          // --- 1. Preparação dos Dados ---
          const empresa = pedidoCompleto.empresaInfo || {};
          const fornecedor = pedidoCompleto.fornecedorInfo || {};
          const itens = pedidoCompleto.itens || [];
          const statusPedido = (pedidoCompleto.status || '').toUpperCase();

          // --- LÓGICA DA MARCA D'ÁGUA ---
            let marcaDaguaHtml = '';
            if (statusPedido === 'CANCELADO' || statusPedido === 'RASCUNHO') {
                const textoMarcaDagua = statusPedido === 'CANCELADO' ? 'CANCELADO' : 'RASCUNHO';
                const corMarcaDagua = statusPedido === 'CANCELADO' ? 'rgba(220, 38, 38, 0.15)' : 'rgba(249, 115, 22, 0.15)'; // Vermelho para cancelado, Laranja para rascunho
                marcaDaguaHtml = `<div class="marca-dagua" style="color: ${corMarcaDagua};">${textoMarcaDagua}</div>`;
            }

          const dados = {
            empresa: {
              nome: empresa.empresa || empresa.razao_social || empresa.nome || 'Nome da Empresa não fornecido',
              endereco: empresa.endereco || '',
              cidadeuf: empresa.cidadeuf || '',
              cnpj: empresa.cnpj || ''
            },
            fornecedor: {
               nome: fornecedor.razao_social || fornecedor.razaoSocial || fornecedor.nome || pedidoCompleto.fornecedor || '',
              cnpj: fornecedor.cnpj || pedidoCompleto.cnpj_fornecedor || '',
              endereco: fornecedor.endereco || pedidoCompleto.endereco_fornecedor || '',
              formaPagamento: fornecedor.forma_de_pagamento || fornecedor.formaDePagamento || pedidoCompleto.forma_pagamento_fornecedor || '',
              condicaoPagamento: fornecedor.condicao_de_pagamento || fornecedor.condicaoDePagamento || pedidoCompleto.condicao_pagamento_fornecedor || ''
            },
            pedido: {
              numero: pedidoCompleto.numero_do_pedido || '',
              data: new Date(pedidoCompleto.data || Date.now()),
              observacoes: pedidoCompleto.observacoes || 'Sem observações.',
              placaVeiculo: pedidoCompleto.placa_veiculo || '',
              nomeVeiculo: pedidoCompleto.nome_veiculo || ''
            },
            financeiro: {
              subtotal: parseFloat(pedidoCompleto.total_geral || 0),
              imposto: parseFloat(pedidoCompleto.icms_st_total || 0)
            },
            usuario: {
             nome: (pedidoCompleto.usuarioCriadorInfo || {}).nome || pedidoCompleto.usuario_criador || 'Usuário não informado',
              funcao: (pedidoCompleto.usuarioCriadorInfo || {}).funcao || 'Função não informada'
            }
          };
          // --- 2. VERIFICAR A CONDIÇÃO DA CIDADE ---
          const cidadeFornecedor = (fornecedor.cidade || '').trim().toLowerCase();
          const isFornecedorLocal = cidadeFornecedor.includes('vitoria da conquista');
          Logger.log(`Fornecedor: '${fornecedor.nome}', Cidade: '${cidadeFornecedor}', É Local? ${isFornecedorLocal}`);
          
          dados.financeiro.totalFinal = dados.financeiro.subtotal + (isFornecedorLocal ? 0 : dados.financeiro.imposto);

          const formatarMoeda = (valor) => (valor || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
          const dataFormatada = isNaN(dados.pedido.data.getTime()) ? 'Data Inválida' : dados.pedido.data.toLocaleDateString('pt-BR');

          const itensHtml = itens.map((item, index) => {
            const produtoForn = item.produtoFornecedor || item.produto_fornecedor || '';
            const precoUnitario = parseFloat(item.precoUnitario) || 0;
            const totalItem = parseFloat(item.totalItem) || 0;
            const quantidade = Number(item.quantidade) || 0;
            return `
               <tr>
                <td class="text-left">${item.descricao || ''}${produtoForn ? `<br><span class="produto-fornecedor">${produtoForn}</span>` : ''}</td>
                <td class="text-center">${quantidade.toLocaleString('pt-BR')}</td>
                <td class="text-center">${item.unidade || ''}</td>
                <td class="text-right">${formatarMoeda(precoUnitario)}</td>
                <td class="text-right">${formatarMoeda(totalItem)}</td>
              </tr>`;
          }).join('');
         
          // --- 3. CRIAR O "COMPONENTE" DE UMA VIA COMPLETA ---
          const umaViaCompletaHtml = `
          <div class="via-impressao">
              ${marcaDaguaHtml}
              <table class="header-table">
                  <tr>
                      <td class="header-info">
                          <strong>${dados.empresa.nome}</strong><br>
                          ${dados.empresa.endereco}<br>
                          ${dados.empresa.cidadeuf}<br>
                          CNPJ: ${dados.empresa.cnpj}
                      </td>
                      <td class="header-pedido">
                          <strong class="titulo-pedido">PEDIDO DE COMPRA</strong><br>
                          <strong class="titulo-pedido">Nº ${dados.pedido.numero}</strong><br>
                          <span class="subtitulo-pedido">DATA DE EMISSÃO</span><br>
                          ${dataFormatada}
                      </td>
                  </tr>
              </table>
              <hr class="separador-forte">
              <table class="info-section">
                  <tr><td colspan="2" class="titulo-secao"><strong>INFORMAÇÕES DO FORNECEDOR</strong></td></tr>
                  <tr>
                      <td style="width: 60%;"><strong>Razão Social:</strong> ${dados.fornecedor.nome}<br><strong>Endereço:</strong> ${dados.fornecedor.endereco}<br><strong>CNPJ:</strong> ${dados.fornecedor.cnpj}</td>
                      <td style="width: 40%;"><strong>CONDIÇÕES DE PAGAMENTO</strong><br><strong>Forma:</strong> ${dados.fornecedor.formaPagamento}<br><strong>Condição:</strong> ${dados.fornecedor.condicaoPagamento}</td>
                  </tr>
              </table>
              <hr class="separador-suave">
              <table class="info-section">
                  <tr><td colspan="2" class="titulo-secao"><strong>INFORMAÇÕES DO VEÍCULO</strong></td></tr>
                  <tr>
                      <td style="width: 60%;"><strong>Modelo:</strong> ${dados.pedido.nomeVeiculo}</td>
                      <td style="width: 40%;"><strong>Placa:</strong> ${dados.pedido.placaVeiculo}</td>
                  </tr>
              </table>
              <hr class="separador-suave">
              <table class="items-table">
                  <thead><tr>
                      <th style="width: 50%; text-align: left;">Item</th><th style="width: 10%;">Qtd.</th><th style="width: 10%;">Unid.</th>
                      <th style="width: 15%; text-align: right;">Vl. Unitário</th><th style="width: 15%; text-align: right;">Subtotal</th>
                  </tr></thead>
                  <tbody>${itensHtml}</tbody>
              </table>
              <p class="total-items-label"><strong>Total dos Itens: ${formatarMoeda(dados.financeiro.subtotal)}</strong></p>
              <table class="footer-section">
                  <tr>
                      <td style="width: 60%; vertical-align: top;"><strong>OBSERVAÇÕES</strong><br><p class="observacoes">${dados.pedido.observacoes}</p><p class="aviso">Atenção: Qualquer alteração só pode ser realizada mediante autorização prévia, sob pena de não pagamento.</p></td>
                      <td style="width: 40%; text-align: right; vertical-align: bottom;" class="bloco-totais">
                          <p>Soma dos Itens &nbsp; <strong>${formatarMoeda(dados.financeiro.subtotal)}</strong></p>
                          ${!isFornecedorLocal ? ` <p>Impostos (ICMS ST) &nbsp; <strong>${formatarMoeda(dados.financeiro.imposto)}</strong></p>` : ''}
                          <p class="total-geral">TOTAL GERAL &nbsp; <strong>${formatarMoeda(dados.financeiro.totalFinal)}</strong></p>
                      </td>
                  </tr>
              </table>
              <div class="assinatura">
                  <p class="linha-assinatura">${dados.usuario.nome}<br><span class="funcao-assinatura">${dados.usuario.funcao}</span></p>
              </div>
          </div>
      `;

          // --- 4. MONTAR O CORPO FINAL COM O LAYOUT DE TABELA PARA DUPLICAÇÃO ---
          let corpoHtmlFinal = '';
          if (isFornecedorLocal) {
              corpoHtmlFinal = `
                  <table class="container-tabela-duas-vias">
                      <tr>
                          <td class="coluna-via">${umaViaCompletaHtml}</td>
                          <td class="coluna-via">${umaViaCompletaHtml}</td>
                      </tr>
                  </table>
              `;
          } else {
              corpoHtmlFinal = umaViaCompletaHtml;
          }

          // --- 5. MONTAR E RETORNAR O DOCUMENTO COMPLETO ---
          return `<!DOCTYPE html>
          <html lang="pt-BR">
          <head>
              <meta charset="UTF-8">
              <title>Pedido de Compra - ${dados.pedido.numero}</title>
              <style>
                  body{font-family:Arial,sans-serif;font-size:9pt;color:#333;margin:0}
                  table{width:100%;border-collapse:collapse}
                  strong{font-weight:bold}
                  hr{border:none;margin:10px 0}
                  hr.separador-forte{border-top:2px solid #000}
                  hr.separador-suave{border-top:1px solid #000}
                  .text-left{text-align:left}
                  .text-center{text-align:center}
                  .text-right{text-align:right}
                  .via-impressao{position:relative;width:100%;width: 100%;margin:auto;padding:20px;background:white;border:1px solid #555;box-sizing:border-box}
                  .header-table td{vertical-align:top;padding:0}
                  .header-info{width:65%}
                  .header-pedido{width:35%;text-align:right}
                  .titulo-pedido{font-size:12pt}
                  .subtitulo-pedido{font-size:8pt;color:#555}
                  .info-section{margin-top:15px}
                  .info-section td{padding:2px 0;vertical-align:top}
                  .info-section .titulo-secao{font-size:10pt;padding-bottom:5px}
                  .items-table{margin-top:15px}
                  .items-table th,.items-table td{border:1px solid #333;padding:5px 6px}
                  .items-table th{background-color:#e0e0e0;font-weight:bold}
                  .items-table tbody tr:nth-child(even){background-color:#f9f9f9}
                  .produto-fornecedor{font-size:8px;color:#c00}
                  .total-items-label{text-align:right;margin-top:5px;font-size:10pt}
                  .footer-section{margin-top:15px}
                  .observacoes{font-size:9pt}
                  .aviso{color:#c00;font-weight:bold}
                  .bloco-totais p{margin:2px 0}
                  .total-geral{font-size:11pt;border-top:1px solid #000;padding-top:5px}
                  .assinatura{text-align:center;margin-top:70px}
                  .linha-assinatura{border-top:1px solid #000;display:inline-block;padding:5px 60px 0 60px;margin:0}
                  .funcao-assinatura{font-size:8pt;color:#555}
                  .marca-dagua{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%) rotate(-45deg);font-size:100px;font-weight:bold;color:rgba(0,0,0,0.1);z-index:1;pointer-events:none;text-align:center;width:100%}

                  @media print {
                      @page { size: A4 landscape; margin: 5mm; }
                      body { width: 100%; margin: 0; }
                      
                      /* Padrão para via única: ocupa a página inteira deitada */
                      .via-impressao { width: 100%; box-sizing: border-box; border: none; padding: 10mm;}
                      
                      /* Container da via dupla (a tabela) */
                      .container-tabela-duas-vias { width: 100%; border-spacing: 5mm 0; border-collapse: separate; }
                      
                      /* Colunas da via dupla */
                      .coluna-via { width: 50%; vertical-align: top; }
                      
                      /* Borda para as vias quando estão duplicadas */
                      .coluna-via .via-impressao { border: 1px dashed #999; padding: 10px; }

                      .rodape-pedido {
                      page-break-inside: avoid !important; /* Tenta não quebrar este bloco */
                      }
                      .assinatura {
                      margin-top: 40px !important; /* Reduz a margem apenas na impressão */
                      }
                  }
              </style>
          </head>
          <body>
              ${corpoHtmlFinal}
          </body>
          </html>`;
}

function registrarImpressao(usuarioLogado, nomeUsuario, pedido, impressora) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Log Impressoes');
  sheet.appendRow([usuarioLogado, nomeUsuario, pedido, new Date(), impressora || '', '']);
}

function getDadosParaImpressao(numeroPedido, empresaId) {
  try {
    const pedidoCompleto = getPedidoCompletoPorId(numeroPedido, empresaId);
    
    if (!pedidoCompleto) {
      return null;
    }

    // "Sanitiza" o objeto: converte o campo de data para uma string no formato ISO.
    // O JavaScript no frontend saberá como ler isso.
    if (pedidoCompleto.data && typeof pedidoCompleto.data.toISOString === 'function') {
      pedidoCompleto.data = pedidoCompleto.data.toISOString();
    }
    // Faça o mesmo para outros campos de data, se houver.

    return pedidoCompleto; // Retorna o objeto "seguro" para transporte.

  } catch (e) {
    Logger.log(`Erro em getDadosParaImpressao: ${e.stack}`);
    return null;
  }
}

/**
 * Função de teste para verificar a 'getPedidoCompletoPorId'.
 * Ela busca um pedido específico e exibe o objeto completo nos logs.
 */
function testarGeracaoDePdf() {
  const numeroPedidoTeste = '001404'; 
  const empresaIdTeste = '001';

  Logger.log(`--- INICIANDO TESTE DE GERAÇÃO DE PDF para Pedido Nº ${numeroPedidoTeste} ---`);

  // Chamada para obter o objeto completo
  const pedidoCompleto = buscarPedidoPorId(numeroPedidoTeste, empresaIdTeste);
  
  if (!pedidoCompleto) {
    Logger.log("❌ ERRO: pedidoCompleto não retornado.");
    return;
  }

  // Log completo do objeto pedidoCompleto
  Logger.log("📦 Objeto pedidoCompleto:");
  Logger.log(JSON.stringify(pedidoCompleto, null, 2));

  // Verifica subestruturas
  Logger.log("🏢 Empresa:");
  Logger.log(JSON.stringify(pedidoCompleto.empresaInfo || pedidoCompleto.empresa || {}, null, 2));
  
  Logger.log("🏭 Fornecedor:");
  Logger.log(JSON.stringify(pedidoCompleto.fornecedorInfo || pedidoCompleto.fornecedor || {}, null, 2));
  
  Logger.log("📋 Itens:");
  Logger.log(JSON.stringify(pedidoCompleto.itens || [], null, 2));
  
  Logger.log("🚚 Dados do veículo:");
  Logger.log("Veículo:", pedidoCompleto.nomeVeiculo);
  Logger.log("Placa:", pedidoCompleto.placaVeiculo);
  
  Logger.log("👤 Criador do pedido:");
  Logger.log("Usuário: " + (pedidoCompleto.usuarioCriador || pedidoCompleto.usuario || ''));
Logger.log("Função: " + (pedidoCompleto.funcaoCriador || pedidoCompleto.cargo || ''));


  // Agora chama a geração do PDF
  const resultado = gerarPdfPedido(numeroPedidoTeste, empresaIdTeste);
  
  Logger.log("📄 Resultado da geração do PDF:");
  Logger.log(resultado);
  
  if (resultado && resultado.status === 'ok') {
    Logger.log("✅ SUCESSO! PDF gerado.");
    Logger.log("Abra este link no seu navegador para ver o arquivo: " + resultado.pdfUrl);
  } else {
    Logger.log(`❌ FALHA! Não foi possível gerar o PDF. Mensagem: ${resultado ? resultado.message : 'Nenhuma resposta'}`);
  }

  Logger.log("--- TESTE CONCLUÍDO ---");
}

function testeMapas() {
  try {
    Logger.log("========== Testando _criarMapaDeEmpresas ==========");
    const mapaEmpresas = _criarMapaDeEmpresas();
    Logger.log("Empresas encontradas: " + Object.keys(mapaEmpresas).length);
    Logger.log(JSON.stringify(mapaEmpresas, null, 2));

    Logger.log("========== Testando _criarMapaDeUsuarios ==========");
    const mapaUsuarios = _criarMapaDeUsuarios();
    Logger.log("Usuários encontrados: " + Object.keys(mapaUsuarios).length);
    Logger.log(JSON.stringify(mapaUsuarios, null, 2));

    Logger.log("========== Testando criarMapaDeFornecedores ==========");
    const mapaFornecedores = criarMapaDeFornecedoresv2();
    Logger.log("Fornecedores encontrados: " + (mapaFornecedores?.size || 0));

    if (mapaFornecedores instanceof Map) {
      mapaFornecedores.forEach((fornecedor, id) => {
        Logger.log("ID: " + id + " => " + JSON.stringify(fornecedor));
      });
    } else {
      Logger.log("Mapa de fornecedores está vazio ou nulo.");
    }

  } catch (erro) {
    Logger.log("Erro ao testar mapas: " + erro.message);
  }
}
