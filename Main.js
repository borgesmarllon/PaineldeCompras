/**
 * @file Main.gs
 * @description Funções principais para servir a aplicação web.
 */

function doGet() {
  // O nome do arquivo HTML principal (página de login) deve estar aqui.
  return HtmlService.createHtmlOutputFromFile('Login')
    .setTitle('Portal de Compras')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function incluir(pagina) {
  return HtmlService.createTemplateFromFile(pagina).evaluate().getContent();
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
