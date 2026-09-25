/**********************************************************************
 * FLUXO VENDA DYNAMICS — EXTRATOR DE TABELAS DE PREÇOS
 * Arquivo: Tabelas          Versão: V1.2
 *
 * O QUE FAZ
 *   Lê os arquivos de tabela de preços que o Orçamento sobe na pasta
 *   "01 Tabelas Atualizadas", extrai os itens com Código Dyn válido e
 *   grava na aba TabelaPrecos da planilha-banco. Cada carga vira uma
 *   VERSÃO datada: nada é apagado, as linhas anteriores daquele cliente
 *   passam a vigente = NAO.
 *
 * V1.1 — correção: cabeçalho de preço no PLURAL ("Valores 2025") não era
 *   reconhecido, e por isso 100% dos itens da HONDA caíam como "SEM PREÇO
 *   VÁLIDO" mesmo com Código Dyn certo. Padrão de preço ampliado para
 *   aceitar singular e plural, mais exceção explícita pro código 331.
 *
 * V1.2 — correção: o cabeçalho só era procurado nas primeiras 15 linhas
 *   da aba. Em abas com o bloco de preço real no meio da planilha (ex.:
 *   RENNER, catálogo geral em cima + "ORÇAMENTOS À PARTE" lá embaixo), a
 *   aba inteira caía como ABA IGNORADA. Agora a aba inteira é varrida e
 *   CADA linha com coluna de Código Dyn vira o início de um bloco de
 *   itens — cobre qualquer quantidade de blocos, em qualquer posição.
 *
 * ANTES DE RODAR
 *   1) Serviços > Serviços do Google avançados > ativar "Drive API" (v2).
 *   2) Rodar configurarTabelas() uma vez — cria as pastas e as abas.
 *   3) Rodar processarTabelas().
 *
 * FUNÇÕES PARA USAR NO EDITOR
 *   configurarTabelas()      cria pastas e abas, grava os IDs no Config
 *   processarTabelas()       processa tudo que está na pasta de entrada
 *   testarTabela(nome)       lê UM arquivo e imprime o resultado, sem gravar
 *   situacaoTabelas()        resumo do que já está na base
 *   buscarPreco(codigo)      consulta usada pelo fluxo de pedidos
 **********************************************************************/

/* ====================== PARÂMETROS ====================== */

var PASTA_RAIZ      = 'FLUXO VENDA DYNAMICS';
var P_ENTRADA       = '01 Tabelas Atualizadas';
var P_PROCESSADAS   = '02 Tabelas Processadas';
var P_FALHAS        = '05 Falhas';
var ABA_PRECOS      = 'TabelaPrecos';
var ABA_ERROS       = 'TabelaErros';
var ABA_CONFIG      = 'Config';
var ABA_CLIENTES    = 'Clientes';
var FUSO            = 'America/Sao_Paulo';
var ORCAMENTO_MS    = 4.5 * 60 * 1000;   // teto de tempo por execução

/* Exceções por código de cliente Dyn, enquanto as tabelas não seguem o
   modelo padrão. Quando o cliente migrar para o modelo, basta apagar a
   linha correspondente daqui. */
var EXCECOES = {
  '331': { precoRe: /valores?\s*20\d\d/, nota: 'HONDA: preço fica em "Valores <ano>", nunca em "Valores Antigos"' },
  '399': { precoCol: 8,  nota: 'SHELL: preço na coluna H' },
  '405': { precoRe: /preco\s*dyn/,  altRe: /cola\s*extra/, altRotulo: 'COM COLA EXTRA' },
  '416': { precoRe: /#\s*4\s*mm/,   altRe: /#\s*3\s*mm/,   altRotulo: 'ACM 3MM' }
};

/* Padrões de cabeçalho aceitos (texto já normalizado: minúsculo, sem acento) */
var PADROES = {
  cod_dyn    : [/\bcod(igo)?s?\b.*\bdyn\b/],
  cod_pailon : [/\bcod(igo)?s?\b.*pailon/],
  desc       : [/^descricao/, /^item$/, /^produto$/],
  detalhe    : [/descritivo/, /informacoes e caracteristicas/, /consideracoes/, /anotacoes/, /informacoes tecnicas/],
  unidade    : [/^un\.?$/, /^unidade$/, /^um$/],
  data       : [/^data/],
  preco      : [/valor unitario/, /preco dyn/, /valores?\s*20\d\d/, /^valores?$/, /^preco/],
  preco_alt  : [/valor variante/],
  desc_alt   : [/descricao da variante/],
  grupo      : [/^grupo$/],
  situacao   : [/^situacao$/]
};

var RE_COD_DYN    = /^\d{3}\.\d{2}\.\d{3}$/;
var RE_COD_PAILON = /^\d{3}-\d{6,}$/;
var UNIDADES = ['UN','UND','UNID','UNIDADE','M2','M²','M3','M³','ML','M','RL','CNJ','PC','PÇ','KG','CX','JG','PAR','L'];

var COLS_PRECOS = ['id_linha','versao_carga','vigente','cliente','cli_dyn','cli_pailon','codigo',
  'cod_pailon','descricao','detalhe','unidade','preco','rotulo_preco','preco_alt','rotulo_alt',
  'data_revisao','status','arquivo','aba','linha_origem'];
var COLS_ERROS = ['versao_carga','arquivo','aba','linha','motivo','conteudo'];

/* ====================== CONFIGURAÇÃO ====================== */

function configurarTabelas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raiz = pastaPorNome_(null, PASTA_RAIZ);
  var ent  = pastaPorNome_(raiz, P_ENTRADA);
  var pro  = pastaPorNome_(raiz, P_PROCESSADAS);
  var fal  = pastaPorNome_(raiz, P_FALHAS);

  criarAba_(ss, ABA_PRECOS, COLS_PRECOS);
  criarAba_(ss, ABA_ERROS,  COLS_ERROS);
  criarAba_(ss, ABA_CLIENTES, ['cod_dyn','cod_pailon','cliente','situacao']);

  var cfg = ss.getSheetByName(ABA_CONFIG);
  if (!cfg) { cfg = ss.insertSheet(ABA_CONFIG); cfg.appendRow(['parametro','valor']); }
  gravarCfg_(cfg, 'pasta_raiz_id', raiz.getId());
  gravarCfg_(cfg, 'pasta_tabelas_entrada_id', ent.getId());
  gravarCfg_(cfg, 'pasta_tabelas_processadas_id', pro.getId());
  gravarCfg_(cfg, 'pasta_falhas_id', fal.getId());
  gravarCfg_(cfg, 'tolerancia_pct', '3');
  gravarCfg_(cfg, 'tabela_antiga_dias', '180');

  Logger.log('Pastas e abas prontas.');
  Logger.log('Raiz: ' + raiz.getUrl());
  Logger.log('Entrada: ' + ent.getUrl());
}

function pastaPorNome_(pai, nome) {
  var it = pai ? pai.getFoldersByName(nome) : DriveApp.getFoldersByName(nome);
  if (it.hasNext()) return it.next();
  return pai ? pai.createFolder(nome) : DriveApp.createFolder(nome);
}

function criarAba_(ss, nome, cols) {
  var sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function gravarCfg_(cfg, chave, valor) {
  var v = cfg.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === chave) { cfg.getRange(i + 1, 2).setValue(valor); return; }
  }
  cfg.appendRow([chave, valor]);
}

function lerCfg_(chave, padrao) {
  var cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CONFIG);
  if (!cfg) return padrao;
  var v = cfg.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === chave) return String(v[i][1]).trim();
  }
  return padrao;
}

/* ====================== PROCESSAMENTO ====================== */

function processarTabelas() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { Logger.log('Outra execução em andamento.'); return; }
  var t0 = Date.now();
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var ent = DriveApp.getFolderById(lerCfg_('pasta_tabelas_entrada_id', ''));
    var pro = DriveApp.getFolderById(lerCfg_('pasta_tabelas_processadas_id', ''));
    var fal = DriveApp.getFolderById(lerCfg_('pasta_falhas_id', ''));
    var shP = ss.getSheetByName(ABA_PRECOS);
    var shE = ss.getSheetByName(ABA_ERROS);
    var versao = Utilities.formatDate(new Date(), FUSO, 'yyyy-MM-dd HH:mm');

    var arquivos = [], it = ent.getFiles();
    while (it.hasNext()) arquivos.push(it.next());
    if (!arquivos.length) { Logger.log('Pasta de entrada vazia.'); return; }

    var totItens = 0, totErros = 0, processados = 0;

    for (var k = 0; k < arquivos.length; k++) {
      if (Date.now() - t0 > ORCAMENTO_MS) { Logger.log('Tempo esgotado — rode de novo para continuar.'); break; }
      var arq = arquivos[k], nome = arq.getName();
      try {
        var res = extrairArquivo_(arq);
        if (!res.itens.length && !res.erros.length) throw new Error('Nenhuma linha reconhecida no arquivo.');

        // desativa a versão anterior dos clientes presentes nesta carga
        var clientes = {};
        for (var i = 0; i < res.itens.length; i++) clientes[res.itens[i].cli_dyn] = true;
        desativarVersaoAnterior_(shP, clientes);

        gravarItens_(shP, res.itens, versao);
        gravarErros_(shE, res.erros, versao);

        arq.moveTo(pro);
        totItens += res.itens.length; totErros += res.erros.length; processados++;
        Logger.log(nome + ' -> ' + res.itens.length + ' itens, ' + res.erros.length + ' rejeitados');
      } catch (e) {
        gravarErros_(shE, [{ arquivo: nome, aba: '', linha: '', motivo: 'FALHA NA LEITURA', conteudo: String(e) }], versao);
        arq.moveTo(fal);
        Logger.log('FALHA em ' + nome + ': ' + e);
      }
    }
    Logger.log('--- ' + processados + ' arquivo(s), ' + totItens + ' itens, ' + totErros + ' rejeitados ---');
  } finally {
    lock.releaseLock();
  }
}

/* Lê um arquivo e devolve {itens, erros}. Não grava nada. */
function extrairArquivo_(arq) {
  var nome = arq.getName();
  var ss = abrirComoPlanilha_(arq);
  var itens = [], erros = [];
  try {
    var abas = ss.ss.getSheets();
    for (var s = 0; s < abas.length; s++) {
      var r = extrairAba_(abas[s], nome);
      itens = itens.concat(r.itens);
      erros = erros.concat(r.erros);
    }
  } finally {
    if (ss.temporaria) DriveApp.getFileById(ss.ss.getId()).setTrashed(true);
  }
  return { itens: itens, erros: erros };
}

/* Converte xlsx em planilha Google temporária; se já for planilha, usa direto. */
function abrirComoPlanilha_(arq) {
  if (arq.getMimeType() === MimeType.GOOGLE_SHEETS) {
    return { ss: SpreadsheetApp.openById(arq.getId()), temporaria: false };
  }
  var copia = Drive.Files.copy(
    { title: '__tmp__' + arq.getName(), mimeType: MimeType.GOOGLE_SHEETS },
    arq.getId()
  );
  return { ss: SpreadsheetApp.openById(copia.id), temporaria: true };
}

/* Uma aba pode ter mais de um bloco de itens — um catálogo geral em cima
   e um "ORÇAMENTOS X À PARTE" mais abaixo, por exemplo. Acha TODAS as
   linhas de cabeçalho (qualquer linha com uma coluna de Código Dyn) e
   processa cada bloco separadamente, do seu cabeçalho até o próximo
   cabeçalho encontrado (ou até o fim da aba). */
function extrairAba_(aba, arquivo) {
  var itens = [], erros = [];
  var nomeAba = aba.getName();
  var ult = aba.getLastRow(), ultC = aba.getLastColumn();
  if (ult < 2 || ultC < 3) return { itens: itens, erros: erros };
  var v = aba.getRange(1, 1, ult, ultC).getValues();

  var hdrs = acharCabecalhos_(v);
  if (!hdrs.length) {
    erros.push({ arquivo: arquivo, aba: nomeAba, linha: '', motivo: 'ABA IGNORADA',
                 conteudo: 'Sem coluna de Código Dyn no cabeçalho' });
    return { itens: itens, erros: erros };
  }

  for (var b = 0; b < hdrs.length; b++) {
    var hi  = hdrs[b];
    var fim = (b + 1 < hdrs.length) ? hdrs[b + 1] : v.length;   // fim exclusivo do bloco
    var r = extrairBloco_(v, hi, fim, ultC, arquivo, nomeAba);
    itens = itens.concat(r.itens);
    erros = erros.concat(r.erros);
  }
  return { itens: itens, erros: erros };
}

function extrairBloco_(v, hi, fim, ultC, arquivo, nomeAba) {
  var itens = [], erros = [];
  var mapa = mapearColunas_(v[hi]);
  if (mapa.cod_dyn === undefined) return { itens: itens, erros: erros };

  var cj = mapa.cod_dyn;
  var cli = clientePredominante_(v, hi, fim, cj);
  var ex  = EXCECOES[cli] || {};
  var hdr = v[hi].map(nrm_);

  // coluna de preço
  var pj = null, rot = '';
  if (ex.precoCol !== undefined && ex.precoCol - 1 < ultC) pj = ex.precoCol - 1;
  else if (ex.precoRe) {
    for (var i = 0; i < hdr.length; i++) {
      if (hdr[i] && ex.precoRe.test(hdr[i])) { pj = i; break; }
    }
  }
  if (pj === null && mapa._preco.length) pj = escolherColunaPreco_(hdr, mapa._preco);
  if (pj !== null) rot = hdr[pj] || ('coluna ' + letra_(pj + 1));

  // coluna de preço variante
  var aj = null, arot = '';
  if (mapa.preco_alt !== undefined) { aj = mapa.preco_alt; arot = ''; }
  else if (ex.altRe) {
    // varre o cabeçalho inteiro: a coluna de variante nem sempre tem rótulo
    // reconhecido como preço (ex.: GWM "Valor com cola extra")
    for (var i2 = 0; i2 < hdr.length; i2++) {
      if (hdr[i2] && i2 !== pj && ex.altRe.test(hdr[i2])) { aj = i2; arot = ex.altRotulo || ''; break; }
    }
  }

  var dj  = escolherDescricao_(v, hi, fim, mapa._desc);
  var tj  = mapa.detalhe;
  var uj  = colunaUnidade_(v, hi, fim, mapa);
  var dtj = mapa.data;
  var paj = mapa.cod_pailon;
  var alj = mapa.desc_alt;
  var sij = mapa.situacao;

  for (var r = hi + 1; r < fim; r++) {
    var linha = v[r];
    var cod = String(linha[cj] === null || linha[cj] === undefined ? '' : linha[cj]).trim();
    if (!RE_COD_DYN.test(cod)) {
      if (preenchidas_(linha) >= 3) {
        erros.push({ arquivo: arquivo, aba: nomeAba, linha: r + 1, motivo: 'SEM CÓDIGO DYN',
                     conteudo: dj !== null ? corta_(linha[dj], 80) : '' });
      }
      continue;
    }
    var preco = numero_(pj !== null ? linha[pj] : null);
    if (preco === null) {
      erros.push({ arquivo: arquivo, aba: nomeAba, linha: r + 1, motivo: 'SEM PREÇO VÁLIDO',
                   conteudo: cod + ' · ' + (dj !== null ? corta_(linha[dj], 70) : '') });
      continue;
    }
    var sit = sij !== undefined ? String(linha[sij] || '').trim().toUpperCase() : '';
    if (sit === 'INATIVO') {
      erros.push({ arquivo: arquivo, aba: nomeAba, linha: r + 1, motivo: 'ITEM INATIVO', conteudo: cod });
      continue;
    }
    var cp = paj !== undefined ? String(linha[paj] || '').trim() : '';
    if (!RE_COD_PAILON.test(cp)) cp = '';
    var dataRev = dtj !== undefined ? dataBr_(linha[dtj]) : '';

    itens.push({
      cliente    : '',
      cli_dyn    : cod.substring(0, 3),
      cli_pailon : cp ? cp.substring(0, 3) : '',
      codigo     : cod,
      cod_pailon : cp,
      descricao  : dj !== null ? limpa_(linha[dj]) : '',
      detalhe    : tj !== undefined ? corta_(limpa_(linha[tj]), 250) : '',
      unidade    : uj !== null ? String(linha[uj] || '').trim().toUpperCase() : '',
      preco      : preco,
      rotulo     : rot,
      preco_alt  : aj !== null ? numero_(linha[aj]) : null,
      rotulo_alt : alj !== undefined ? limpa_(linha[alj]) : arot,
      data_rev   : dataRev || 'SEM DATA DE REVISÃO',
      arquivo    : arquivo,
      aba        : nomeAba,
      linha      : r + 1
    });
  }
  return { itens: itens, erros: erros };
}

/* ====================== GRAVAÇÃO ====================== */

function desativarVersaoAnterior_(sh, clientes) {
  if (sh.getLastRow() < 2) return;
  var n = sh.getLastRow() - 1;
  var cliCol = sh.getRange(2, 5, n, 1).getValues();     // cli_dyn
  var vigCol = sh.getRange(2, 3, n, 1).getValues();     // vigente
  var mudou = false;
  for (var i = 0; i < n; i++) {
    if (clientes[String(cliCol[i][0]).trim()] && String(vigCol[i][0]).trim() === 'SIM') {
      vigCol[i][0] = 'NAO'; mudou = true;
    }
  }
  if (mudou) sh.getRange(2, 3, n, 1).setValues(vigCol);
}

function gravarItens_(sh, itens, versao) {
  if (!itens.length) return;
  var nomes = mapaClientes_();
  // detecta código duplicado DENTRO da carga
  var conta = {};
  for (var i = 0; i < itens.length; i++) conta[itens[i].codigo] = (conta[itens[i].codigo] || 0) + 1;

  var base = sh.getLastRow();
  var linhas = itens.map(function (it, idx) {
    var status = 'OK';
    if (conta[it.codigo] > 1) status = 'CÓDIGO DUPLICADO NA BASE';
    else if (it.data_rev === 'SEM DATA DE REVISÃO') status = 'SEM DATA DE REVISÃO';
    return [
      'TP' + (base + idx), versao, 'SIM',
      nomes[it.cli_dyn] || '', it.cli_dyn, it.cli_pailon,
      it.codigo, it.cod_pailon, it.descricao, it.detalhe, it.unidade,
      it.preco, it.rotulo, it.preco_alt === null ? '' : it.preco_alt, it.rotulo_alt,
      it.data_rev, status, it.arquivo, it.aba, it.linha
    ];
  });
  sh.getRange(sh.getLastRow() + 1, 1, linhas.length, COLS_PRECOS.length).setValues(linhas);
}

function gravarErros_(sh, erros, versao) {
  if (!erros.length) return;
  var linhas = erros.map(function (e) {
    return [versao, e.arquivo, e.aba, e.linha, e.motivo, e.conteudo];
  });
  sh.getRange(sh.getLastRow() + 1, 1, linhas.length, COLS_ERROS.length).setValues(linhas);
}

function mapaClientes_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_CLIENTES);
  var m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < v.length; i++) {
    var c = String(v[i][0]).trim();
    if (c) m[c.length === 2 ? '0' + c : c] = String(v[i][2]).trim();
  }
  return m;
}

/* ====================== CONSULTA USADA PELO PEDIDO ====================== */

/**
 * Devolve a referência de preço de um Código Dyn.
 * status possíveis:
 *   OK                          preço encontrado e utilizável
 *   SEM PREÇO EM TABELA         código não existe em nenhuma tabela vigente
 *   CÓDIGO DUPLICADO NA BASE    mais de uma linha vigente com o mesmo código
 *   SEM DATA DE REVISÃO         preço existe, mas sem data de revisão
 *   TABELA DESATUALIZADA        data de revisão mais velha que o parâmetro
 */
function buscarPreco(codigo) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PRECOS);
  if (!sh || sh.getLastRow() < 2) return { status: 'SEM PREÇO EM TABELA', preco: null };
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_PRECOS.length).getValues();
  var achados = [];
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][2]).trim() === 'SIM' && String(v[i][6]).trim() === String(codigo).trim()) achados.push(v[i]);
  }
  if (!achados.length) return { status: 'SEM PREÇO EM TABELA', preco: null };
  if (achados.length > 1) {
    return { status: 'CÓDIGO DUPLICADO NA BASE', preco: null,
             precos: achados.map(function (a) { return a[11]; }) };
  }
  var a = achados[0];
  var status = 'OK';
  var dataRev = String(a[15]).trim();
  if (dataRev === 'SEM DATA DE REVISÃO') status = 'SEM DATA DE REVISÃO';
  else {
    var lim = Number(lerCfg_('tabela_antiga_dias', '180'));
    var d = dataDe_(dataRev);
    if (d && (new Date() - d) / 86400000 > lim) status = 'TABELA DESATUALIZADA';
  }
  return {
    status     : status,
    cliente    : a[3],
    codigo     : a[6],
    descricao  : a[8],
    unidade    : a[10],
    preco      : Number(a[11]),
    preco_alt  : a[13] === '' ? null : Number(a[13]),
    rotulo_alt : a[14],
    data_rev   : dataRev,
    versao     : a[1]
  };
}

/**
 * Compara o preço praticado com a referência. Se não bater com o preço base
 * mas bater com a variante dentro da tolerância, resolve como variante.
 */
function compararPreco(codigo, precoPraticado) {
  var ref = buscarPreco(codigo);
  if (ref.preco === null) return { status: ref.status, variacao: null, referencia: null };
  var tol = Number(lerCfg_('tolerancia_pct', '3'));
  var vBase = (precoPraticado - ref.preco) / ref.preco * 100;
  var usado = { base: ref.preco, rotulo: 'TABELA', variacao: vBase };
  if (ref.preco_alt && Math.abs(vBase) > tol) {
    var vAlt = (precoPraticado - ref.preco_alt) / ref.preco_alt * 100;
    if (Math.abs(vAlt) < Math.abs(vBase)) {
      usado = { base: ref.preco_alt, rotulo: ref.rotulo_alt || 'VARIANTE', variacao: vAlt };
    }
  }
  return {
    status     : ref.status,
    cliente    : ref.cliente,
    descricao  : ref.descricao,
    referencia : usado.base,
    rotulo     : usado.rotulo,
    variacao   : Math.round(usado.variacao * 100) / 100,
    dentro     : Math.abs(usado.variacao) <= tol,
    data_rev   : ref.data_rev
  };
}

/* ====================== DIAGNÓSTICO ====================== */

function testarTabela(nome) {
  var ent = DriveApp.getFolderById(lerCfg_('pasta_tabelas_entrada_id', ''));
  var arq = null;
  if (nome) {
    var it = ent.getFilesByName(nome);
    if (it.hasNext()) arq = it.next();
  } else {
    var it2 = ent.getFiles(), mais = null;
    while (it2.hasNext()) { var f = it2.next(); if (!mais || f.getDateCreated() > mais.getDateCreated()) mais = f; }
    arq = mais;
  }
  if (!arq) { Logger.log('Arquivo não encontrado na pasta de entrada.'); return; }

  var r = extrairArquivo_(arq);
  Logger.log('ARQUIVO: ' + arq.getName());
  Logger.log('Itens lidos: ' + r.itens.length + ' | rejeitados: ' + r.erros.length);
  var porCli = {};
  r.itens.forEach(function (i) { porCli[i.cli_dyn] = (porCli[i.cli_dyn] || 0) + 1; });
  Logger.log('Clientes: ' + JSON.stringify(porCli));
  Logger.log('--- primeiros 5 itens ---');
  r.itens.slice(0, 5).forEach(function (i) {
    Logger.log([i.codigo, i.descricao, i.unidade, i.preco,
                i.preco_alt ? '(' + i.rotulo_alt + ' ' + i.preco_alt + ')' : '',
                i.data_rev].join(' | '));
  });
  var motivos = {};
  r.erros.forEach(function (e) { motivos[e.motivo] = (motivos[e.motivo] || 0) + 1; });
  Logger.log('--- motivos de rejeição --- ' + JSON.stringify(motivos));
  r.erros.slice(0, 8).forEach(function (e) { Logger.log('  linha ' + e.linha + ' · ' + e.motivo + ' · ' + e.conteudo); });
}

function situacaoTabelas() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PRECOS);
  if (!sh || sh.getLastRow() < 2) { Logger.log('Base vazia.'); return; }
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_PRECOS.length).getValues();
  var res = {};
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][2]).trim() !== 'SIM') continue;
    var c = v[i][4];
    if (!res[c]) res[c] = { cliente: v[i][3], itens: 0, ok: 0, semData: 0, dup: 0, versao: v[i][1] };
    res[c].itens++;
    var st = String(v[i][16]);
    if (st === 'OK') res[c].ok++;
    else if (st.indexOf('DUPLICADO') >= 0) res[c].dup++;
    else if (st.indexOf('SEM DATA') >= 0) res[c].semData++;
  }
  Logger.log('cliente | itens | ok | sem data | duplicado | versão');
  Object.keys(res).sort().forEach(function (c) {
    var r = res[c];
    Logger.log([c + ' ' + r.cliente, r.itens, r.ok, r.semData, r.dup, r.versao].join(' | '));
  });
}

/* ====================== UTILITÁRIOS ====================== */

function nrm_(s) {
  if (s === null || s === undefined) return '';
  return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
         .toLowerCase().replace(/\s+/g, ' ').trim();
}

/* Acha TODAS as linhas de cabeçalho da aba inteira (sem limite de linha) —
   qualquer linha que tenha uma célula reconhecida como Código Dyn é o
   início de um bloco de itens. Uma aba pode ter o cabeçalho longe do topo
   (catálogo geral em cima, tabela de preço real bem mais abaixo) ou mais
   de um bloco (ex.: "ORÇAMENTOS X À PARTE" no meio da mesma aba). */
function acharCabecalhos_(v) {
  var hdrs = [];
  for (var i = 0; i < v.length; i++) {
    if (temColunaCodDyn_(v[i])) hdrs.push(i);
  }
  return hdrs;
}

function temColunaCodDyn_(linha) {
  for (var j = 0; j < linha.length; j++) {
    var t = nrm_(linha[j]);
    if (!t) continue;
    for (var p = 0; p < PADROES.cod_dyn.length; p++) {
      if (PADROES.cod_dyn[p].test(t)) return true;
    }
  }
  return false;
}

function mapearColunas_(linha) {
  var mapa = { _preco: [], _desc: [] };
  for (var j = 0; j < linha.length; j++) {
    var t = nrm_(linha[j]);
    if (!t) continue;
    for (var k in PADROES) {
      for (var p = 0; p < PADROES[k].length; p++) {
        if (PADROES[k][p].test(t)) {
          if (k === 'preco') mapa._preco.push(j);
          else if (k === 'desc') mapa._desc.push(j);
          else if (mapa[k] === undefined) mapa[k] = j;
          p = PADROES[k].length;
        }
      }
    }
  }
  if (mapa._desc.length) mapa.desc = mapa._desc[0];
  return mapa;
}

/* Havendo mais de uma coluna candidata a preço (ex.: "Valores Antigos" e
   "Valores 2025" no mesmo arquivo), fica com a que tiver o ano mais recente
   no rótulo. Sem ano em nenhuma candidata, fica com a última (mais à
   direita — convenção mais comum de "preço vigente" nas tabelas do
   Orçamento). */
function escolherColunaPreco_(hdr, cands) {
  var melhor = cands[cands.length - 1], anoMax = -1;
  for (var i = 0; i < cands.length; i++) {
    var m = /20\d\d/.exec(hdr[cands[i]] || '');
    if (m) {
      var ano = Number(m[0]);
      if (ano > anoMax) { anoMax = ano; melhor = cands[i]; }
    }
  }
  return melhor;
}

/* Havendo mais de uma coluna candidata a descrição (ex.: OMODA tem duas
   colunas ITEM), fica com a de texto mais longo no corpo da tabela. */
function escolherDescricao_(v, hi, fim, cands) {
  if (!cands || !cands.length) return null;
  var melhor = cands[0], sc = -1;
  for (var c = 0; c < cands.length; c++) {
    var soma = 0, n = 0;
    for (var r = hi + 1; r < Math.min(fim, hi + 40); r++) {
      var t = v[r][cands[c]];
      if (t !== null && t !== undefined && String(t).trim()) { soma += String(t).length; n++; }
    }
    var m = n ? soma / n : 0;
    if (m > sc) { sc = m; melhor = cands[c]; }
  }
  return melhor;
}

/* Quando a coluna de unidade não tem rótulo (caso MG) ou tem rótulo errado
   (caso GWM, rotulada "Qtd"), acha pela vocabulário do conteúdo. */
function colunaUnidade_(v, hi, fim, mapa) {
  if (mapa.unidade !== undefined) return mapa.unidade;
  var melhor = null, sc = 0;
  var nCols = v[hi].length;
  for (var j = 0; j < nCols; j++) {
    var bons = 0, tot = 0;
    for (var r = hi + 1; r < Math.min(fim, hi + 60); r++) {
      var t = String(v[r][j] === null || v[r][j] === undefined ? '' : v[r][j]).trim().toUpperCase();
      if (!t) continue;
      tot++;
      if (UNIDADES.indexOf(t) >= 0) bons++;
    }
    if (tot >= 3) {
      var razao = bons / tot;
      if (razao > sc && razao > 0.5) { sc = razao; melhor = j; }
    }
  }
  return melhor;
}

function clientePredominante_(v, hi, fim, cj) {
  var conta = {};
  for (var r = hi + 1; r < fim; r++) {
    var c = String(v[r][cj] === null || v[r][cj] === undefined ? '' : v[r][cj]).trim();
    if (RE_COD_DYN.test(c)) { var p = c.substring(0, 3); conta[p] = (conta[p] || 0) + 1; }
  }
  var melhor = null, n = 0;
  for (var k in conta) if (conta[k] > n) { n = conta[k]; melhor = k; }
  return melhor;
}

function numero_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v > 0 ? Math.round(v * 10000) / 10000 : null;
  var s = String(v).trim().replace(/R\$/g, '').replace(/\s/g, '');
  if (!s || s === '*' || s === '-' || s === '—') return null;
  if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  else s = s.replace(/,/g, '');
  var f = parseFloat(s);
  return isNaN(f) || f <= 0 ? null : Math.round(f * 10000) / 10000;
}

function dataBr_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, FUSO, 'dd/MM/yyyy');
  }
  var s = String(v).trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  var d = new Date(s);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, FUSO, 'dd/MM/yyyy');
}

function dataDe_(s) {
  var m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(s).trim());
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

function limpa_(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/\s+/g, ' ').trim();
}

function corta_(v, n) {
  var s = limpa_(v);
  return s.length > n ? s.substring(0, n) : s;
}

function preenchidas_(linha) {
  var n = 0;
  for (var i = 0; i < linha.length; i++) {
    if (linha[i] !== null && linha[i] !== undefined && String(linha[i]).trim() !== '') n++;
  }
  return n;
}

function letra_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
  return s;
}
