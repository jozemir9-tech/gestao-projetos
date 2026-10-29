/**
 * ============================================================================
 *  API REST DA PLANILHA — App Gestão de Projetos (Google Apps Script)
 * ============================================================================
 *  Expõe as 6 abas do Google Sheets como uma API HTTP (JSON), protegida por
 *  API Key, para o app externo (Streamlit) — ou qualquer outro aplicativo —
 *  consultar os dados sem precisar de conta de serviço.
 *
 *  ABAS ATENDIDAS (devem existir na planilha, com cabeçalho na 1ª linha):
 *    DB_Cronograma_MSProject   (chave: ID_Tarefa)
 *    DB_Histograma_MO          (chave: ID_Linha)
 *    DB_Histograma_Equip       (chave: ID_Linha)
 *    DB_Recursos_Humanos       (chave: Matricula)
 *    DB_Contratos              (chave: ID_Contrato)
 *    Timeline_Visao            (chave: ID_Marco)
 *
 *  ENDPOINTS (GET):
 *    ?action=list&aba=DB_Contratos                -> todas as linhas da aba (JSON)
 *    ?action=get&aba=DB_Contratos&chave=CTR-01    -> 1 linha pela chave
 *    ?action=info                                 -> estrutura (abas, colunas, totais)
 *    ?action=ping                                 -> teste de vida
 *    Filtros opcionais: &busca=texto&limite=500&coluna=Status&valor=Em Execução
 *
 *  ESCRITA (POST, corpo JSON):
 *    ?action=create&aba=DB_Contratos   (corpo: objeto da linha)
 *    ?action=update&aba=DB_Contratos   (corpo: objeto com a chave)
 *    ?action=delete&aba=DB_Contratos&chave=CTR-01
 *
 *  AUTENTICAÇÃO: envie a API Key em 1 destes lugares:
 *    - Header  "x-api-key: SUA_CHAVE"
 *    - Query   "?api_key=SUA_CHAVE"
 *    - Corpo   {"api_key": "SUA_CHAVE", ...}  (POST)
 *
 *  PARA GERAR A CHAVE: no editor, selecione a função gerarEInstalarChave e
 *  clique em "Executar" (1 vez). Copie a chave do "Registro de execução".
 * ============================================================================
 */

// =========================== CONFIGURAÇÕES ================================

const CONFIG = {
  /** Nome de cada aba → coluna usada como identificador único (chave). */
  ABAS: {
    DB_Cronograma_MSProject: 'ID_Tarefa',
    DB_Histograma_MO: 'ID_Linha',
    DB_Histograma_Equip: 'ID_Linha',
    DB_Recursos_Humanos: 'Matricula',
    DB_Contratos: 'ID_Contrato',
    Timeline_Visao: 'ID_Marco',
  },
  /** Límite padrão de linhas por listagem (evita estourar a cota do Apps Script). */
  LIMITE_PADRAO: 5000,
};

// ============================ PONTO DE ENTRADA =============================

function doGet(e) {
  try {
    autorizar_(e);
    const p = (e && e.parameter) || {};
    const acao = String(p.action || 'list').toLowerCase();

    switch (acao) {
      case 'list':
        return json_(listar_(p));
      case 'get':
        return json_({ data: buscarPorChave_(p) });
      case 'info':
        return json_(informacoes_());
      case 'ping':
        return json_({ message: 'API funcionando!' });
      default:
        throw new ErroApi_(400, 'Ação inválida no GET: "' + acao + '".');
    }
  } catch (erro) {
    return responderErro_(erro);
  }
}

function doPost(e) {
  try {
    autorizar_(e);
    const p = (e && e.parameter) || {};
    const corpo = lerCorpoJson_(e);
    const acao = String(corpo.action || p.action || '').toLowerCase();
    if (!acao) throw new ErroApi_(400, 'Informe a ação: create, update ou delete.');
    const aba = corpo.aba || p.aba;

    switch (acao) {
      case 'create':
        return json_({ data: criar_(aba, corpo) }, 201);
      case 'update':
        return json_({ data: atualizar_(aba, corpo) });
      case 'delete':
        return json_({ data: excluir_(aba, corpo.chave || p.chave) });
      default:
        throw new ErroApi_(400, 'Ação inválida no POST: "' + acao + '".');
    }
  } catch (erro) {
    return responderErro_(erro);
  }
}

// ================================ LEITURA ==================================

/** Lista as linhas de uma aba, com filtros, busca e limite. */
function listar_(p) {
  const aba = exigirAba_(p.aba);
  const { cabecalhos, registros } = lerAba_(aba);

  let resultado = registros.map((r) => {
    const c = Object.assign({}, r);
    delete c._linha;
    return c;
  });

  // Filtro exato por coluna: &coluna=Status&valor=Em Execução
  if (p.coluna && p.valor !== undefined) {
    resultado = resultado.filter(
      (r) => String(r[p.coluna] == null ? '' : r[p.coluna]) === String(p.valor)
    );
  }

  // Busca textual em todas as colunas: &busca=texto
  if (p.busca) {
    const termo = String(p.busca).toLowerCase();
    resultado = resultado.filter((r) =>
      cabecalhos.some((h) => String(r[h] == null ? '' : r[h]).toLowerCase().indexOf(termo) !== -1)
    );
  }

  const total = resultado.length;
  const limite = Math.max(parseInt(p.limite, 10) || CONFIG.LIMITE_PADRAO, 1);

  return {
    data: resultado.slice(0, limite),
    meta: { aba: aba.getName(), columns: cabecalhos, total: total, limit: limite },
  };
}

/** Busca 1 linha pela coluna-chave da aba. */
function buscarPorChave_(p) {
  const aba = exigirAba_(p.aba);
  if (!p.chave) throw new ErroApi_(400, 'O parâmetro "chave" é obrigatório.');
  const chaveCol = CONFIG.ABAS[aba.getName()];
  const { registros } = lerAba_(aba);
  const alvo = registros.find((r) => String(r[chaveCol]) === String(p.chave));
  if (!alvo) throw new ErroApi_(404, 'Registro "' + p.chave + '" não encontrado em ' + aba.getName() + '.');
  delete alvo._linha;
  return alvo;
}

/** Estrutura da planilha: abas, colunas e total de linhas (para o app montar os dados). */
function informacoes_() {
  const out = { sheets: {} };
  Object.keys(CONFIG.ABAS).forEach((nome) => {
    try {
      const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
      if (!aba) {
        out.sheets[nome] = { found: false };
        return;
      }
      const { cabecalhos, registros } = lerAba_(aba);
      out.sheets[nome] = { found: true, key: CONFIG.ABAS[nome], columns: cabecalhos, rows: registros.length };
    } catch (err) {
      out.sheets[nome] = { found: false, error: String(err) };
    }
  });
  return out;
}

// ================================ ESCRITA ==================================

function criar_(nomeAba, dados) {
  return comTrava_(() => {
    const aba = exigirAba_(nomeAba);
    const chaveCol = CONFIG.ABAS[aba.getName()];
    let cabecalhos = cabecalhosDe_(aba);

    // Aba vazia: cria cabeçalho com a chave + campos do payload
    if (cabecalhos.length === 0) {
      cabecalhos = [chaveCol].concat(
        Object.keys(dados).filter((k) => k !== 'action' && k !== 'aba' && k !== 'api_key' && k !== chaveCol)
      );
      aba.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    }

    // Garante a coluna-chave
    if (cabecalhos.indexOf(chaveCol) === -1) {
      aba.insertColumnBefore(1);
      aba.getRange(1, 1).setValue(chaveCol);
      cabecalhos = [chaveCol].concat(cabecalhos);
    }

    // Campos novos viram colunas novas
    Object.keys(dados).forEach((k) => {
      if (k === 'action' || k === 'aba' || k === 'api_key') return;
      if (cabecalhos.indexOf(k) === -1) {
        aba.getRange(1, cabecalhos.length + 1).setValue(k);
        cabecalhos.push(k);
      }
    });

    const registro = Object.assign({}, dados);
    delete registro.action;
    delete registro.aba;
    delete registro.api_key;
    registro[chaveCol] = registro[chaveCol] || proximaChave_(aba, chaveCol);

    const linha = cabecalhos.map((h) => (registro[h] !== undefined ? registro[h] : ''));
    aba.appendRow(linha);
    delete registro._linha;
    return registro;
  });
}

function atualizar_(nomeAba, dados) {
  return comTrava_(() => {
    const aba = exigirAba_(nomeAba);
    const chaveCol = CONFIG.ABAS[aba.getName()];
    if (!dados[chaveCol]) throw new ErroApi_(400, 'Informe o campo-chave "' + chaveCol + '".');

    const { registros } = lerAba_(aba);
    const alvo = registros.find((r) => String(r[chaveCol]) === String(dados[chaveCol]));
    if (!alvo) throw new ErroApi_(404, 'Registro "' + dados[chaveCol] + '" não encontrado em ' + aba.getName() + '.');

    delete dados.action;
    delete dados.aba;
    delete dados.api_key;

    const atualizado = Object.assign({}, alvo, dados);
    const cabecalhos = cabecalhosComNovos_(aba, Object.keys(dados), chaveCol);
    const linha = cabecalhos.map((h) => (atualizado[h] !== undefined ? atualizado[h] : ''));
    aba.getRange(alvo._linha, 1, 1, linha.length).setValues([linha]);
    delete atualizado._linha;
    return atualizado;
  });
}

function excluir_(nomeAba, chave) {
  return comTrava_(() => {
    const aba = exigirAba_(nomeAba);
    const chaveCol = CONFIG.ABAS[aba.getName()];
    if (!chave) throw new ErroApi_(400, 'Informe o campo-chave "' + chaveCol + '".');
    const { registros } = lerAba_(aba);
    const alvo = registros.find((r) => String(r[chaveCol]) === String(chave));
    if (!alvo) throw new ErroApi_(404, 'Registro "' + chave + '" não encontrado em ' + aba.getName() + '.');
    aba.deleteRow(alvo._linha);
    return { chave: chave, deleted: true };
  });
}

// =============================== AUXILIARES ================================

/** Lê uma aba inteira → { cabecalhos, registros }. Cada linha vira objeto. */
function lerAba_(aba) {
  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();
  if (ultimaLinha < 1 || ultimaColuna < 1) return { cabecalhos: [], registros: [] };

  const valores = aba.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();
  const cabecalhos = valores[0].map((c) => String(c).trim());

  const registros = [];
  for (let i = 1; i < valores.length; i++) {
    const linha = valores[i];
    if (linha.every((c) => c === '' || c === null)) continue;
    const obj = { _linha: i + 1 };
    for (let j = 0; j < cabecalhos.length; j++) {
      if (!cabecalhos[j]) continue;
      obj[cabecalhos[j]] = formatarCelula_(linha[j]);
    }
    registros.push(obj);
  }
  return { cabecalhos, registros };
}

/** Datas do Sheets → texto ISO local (o app Streamlit faz o parse). */
function formatarCelula_(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return valor;
}

function cabecalhosDe_(aba) {
  if (aba.getLastRow() < 1 || aba.getLastColumn() < 1) return [];
  return aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0].map((c) => String(c).trim());
}

function cabecalhosComNovos_(aba, novosCampos, chaveCol) {
  let cabecalhos = cabecalhosDe_(aba);
  novosCampos.forEach((k) => {
    if (cabecalhos.indexOf(k) === -1) {
      aba.getRange(1, cabecalhos.length + 1).setValue(k);
      cabecalhos.push(k);
    }
  });
  return cabecalhos;
}

/** Próxima chave: numérica sequencial se a coluna for numérica; senão, UUID. */
function proximaChave_(aba, chaveCol) {
  const { registros } = lerAba_(aba);
  const ids = registros.map((r) => r[chaveCol]).filter((v) => v !== '' && v != null);
  const numericos = ids.filter((v) => !isNaN(Number(v)));
  if (ids.length === numericos.length) {
    const max = numericos.reduce((m, v) => Math.max(m, Number(v)), 0);
    return max + 1;
  }
  return Utilities.getUuid().slice(0, 8);
}

function exigirAba_(nome) {
  if (!nome) throw new ErroApi_(400, 'Informe o parâmetro "aba". Ex.: aba=DB_Contratos.');
  if (!CONFIG.ABAS[nome]) {
    throw new ErroApi_(400, 'Aba "' + nome + '" não é suportada. Use: ' + Object.keys(CONFIG.ABAS).join(', ') + '.');
  }
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!aba) throw new ErroApi_(404, 'Aba "' + nome + '" não existe nesta planilha.');
  return aba;
}

/** Trava de edição para evitar conflitos em escritas simultâneas. */
function comTrava_(funcao) {
  const trava = LockService.getScriptLock();
  trava.waitLock(20000);
  try {
    return funcao();
  } finally {
    trava.releaseLock();
  }
}

// ============================ AUTENTICAÇÃO =================================

function autorizar_(e) {
  const chaveEsperada = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!chaveEsperada) {
    throw new ErroApi_(
      400,
      'API Key não configurada. No editor, execute a função "gerarEInstalarChave" uma vez.'
    );
  }

  let informada = obterHeader_(e, 'x-api-key');
  if (!informada && e && e.parameter) informada = e.parameter.api_key;
  if (!informada && e && e.postData && e.postData.contents) {
    try {
      const corpo = JSON.parse(e.postData.contents);
      informada = corpo.api_key || null;
    } catch (err) { /* corpo não é JSON */ }
  }

  if (informada !== chaveEsperada) throw new ErroApi_(401, 'API Key inválida ou ausente.');
}

function obterHeader_(e, nome) {
  if (!e) return null;
  const alvo = String(nome).toLowerCase();
  const fontes = [e.headers, e.commonHeaders];
  for (let f = 0; f < fontes.length; f++) {
    const fonte = fontes[f];
    if (!fonte) continue;
    for (const k of Object.keys(fonte)) {
      if (String(k).toLowerCase() === alvo) return fonte[k];
    }
  }
  return null;
}

function lerCorpoJson_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  try {
    const corpo = JSON.parse(e.postData.contents);
    if (corpo === null || typeof corpo !== 'object' || Array.isArray(corpo)) {
      throw new ErroApi_(400, 'O corpo da requisição deve ser um objeto JSON.');
    }
    return corpo;
  } catch (err) {
    if (err instanceof ErroApi_) throw err;
    throw new ErroApi_(400, 'JSON inválido no corpo da requisição.');
  }
}

// ============================== RESPOSTAS ==================================

function json_(objeto, codigo) {
  return ContentService.createTextOutput(JSON.stringify(Object.assign({ success: true }, objeto)))
    .setMimeType(ContentService.MimeType.JSON)
    .setResponseCode(codigo || 200);
}

function ErroApi_(codigo, mensagem) {
  this.code = codigo;
  this.message = mensagem;
  this.name = 'ErroApi';
}
ErroApi_.prototype = Object.create(Error.prototype);

function responderErro_(erro) {
  const codigo = (erro && erro.code) || 500;
  const mensagem = (erro && erro.message) || 'Erro interno: ' + String(erro);
  if (codigo === 500) console.error(erro && erro.stack ? erro.stack : erro);
  return ContentService.createTextOutput(
    JSON.stringify({ success: false, error: mensagem, code: codigo })
  )
    .setMimeType(ContentService.MimeType.JSON)
    .setResponseCode(codigo);
}

// ============ FUNÇÃO PARA EXECUTAR UMA VEZ NO EDITOR =======================
// 1. Selecione "gerarEInstalarChave" no menu de funções do editor.
// 2. Clique em "Executar" e autorize (primeira vez pede permissão da planilha).
// 3. Copie a chave mostrada no "Registro de execução".

function gerarEInstalarChave() {
  const chave = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('API_KEY', chave);
  Logger.log('API_KEY criada com sucesso: ' + chave);
}
