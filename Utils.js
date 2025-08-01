/**
 * @file Utils.gs
 * @description Funções de utilidade usadas em múltiplos módulos do projeto.
 */

/**
 * Converte uma string de cabeçalho para o formato camelCase.
 * Ex: "Número do Pedido" -> "numeroDoPedido"
 * @param {string} str O cabeçalho da coluna.
 * @returns {string} O cabeçalho formatado em camelCase.
 */
function toCamelCase(str) {
  if (!str) return '';

  return String(str)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/ (\w)/g, (match, p1) => p1.toUpperCase());
}


/**
 * Gera um hash SHA-256 para uma string.
 * @param {string} senha A string a ser hasheada.
 * @returns {string} O hash em formato base64.
 */
function gerarHash(senha) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, senha);
  return Utilities.base64Encode(digest);
}


/**
 * Formata uma data para o padrão ISO (YYYY-MM-DD HH:mm:ss) no fuso de São Paulo.
 * @param {string|Date} data - Data a ser formatada.
 * @returns {string} Data no formato ISO.
 */
function formatarDataParaISO(data) {
  try {
    let dataObj;

    if (!data) {
      dataObj = new Date();
    } else if (typeof data === 'string') {
      // Tenta cobrir formatos como YYYY-MM-DD
      if (data.match(/^\d{4}-\d{2}-\d{2}$/)) {
        const agora = new Date();
        dataObj = new Date(`${data}T${agora.toTimeString().split(' ')[0]}`);
      } else {
        dataObj = new Date(data);
      }
      if (isNaN(dataObj.getTime())) {
        dataObj = new Date();
      }
    } else if (data instanceof Date) {
      dataObj = data;
    } else {
      dataObj = new Date();
    }
    return Utilities.formatDate(dataObj, 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm:ss');
  } catch (error) {
    Logger.log(`Erro ao formatar data: ${error}. Data recebida: ${data}`);
    return Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd HH:mm:ss');
  }
}

/**
 * Calcula a Distância de Levenshtein entre duas strings.
 * @param {string} a Primeira string.
 * @param {string} b Segunda string.
 * @returns {number} O número de edições para transformar a em b.
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
