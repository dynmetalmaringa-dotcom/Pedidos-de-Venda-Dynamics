/**********************************************************************
 * FLUXO VENDA DYNAMICS — EXTRATOR DE PEDIDOS DE VENDA
 * Arquivo: Pedidos          Versão: V1.4
 *
 * O QUE FAZ
 *   Lê os PDFs de pedido de venda do ForWood que o Orçamento sobe na
 *   pasta "03 Pedidos de Venda Novos", converte o valor de cada item
 *   para o valor cheio (x2,5), cruza com a tabela de preços do cliente
 *   pelo Código Dyn e grava tudo com a comparação já feita.
 *
 *   Revisão nova do mesmo pedido NÃO apaga a anterior: entra como
 *   registro próprio, a antiga vira vigente = NAO e o app grava na aba
 *   Revisoes exatamente o que mudou, item a item.
 *
 * DEPENDE DO ARQUIVO "Tabelas"
 *   Usa buscarPreco() e compararPreco() de lá. Os dois arquivos precisam
 *   estar no MESMO projeto do Apps Script.
 *
 * ANTES DE RODAR
 *   1) Drive API v2 ativada (mesma da etapa 3 do guia).
 *   2) configurarPedidos() uma vez.
 *
 * V1.2 — correção crítica: a conversão do PDF para Documento Google junta
 *   TODOS os itens de uma página numa linha só (e cola o rodapé no fim dela).
 *   A leitura por linha só enxergava o primeiro item e pegava os valores do
 *   último — 3 dos 4 pedidos de teste entraram com 1 item e valor errado.
 *   Agora o texto é lido como fluxo único de palavras e cada Código Dyn abre
 *   um item, com os 5 PRIMEIROS números seguintes (e não os últimos da linha).
 *   Também: a conferência de quantidade virou auxiliar ("SEM REFERENCIA"
 *   quando o rodapé não traz o número) e não reprova mais o pedido sozinha.
 *
 * V1.3 — correção: o Sheets converte código em número ao gravar ("00" vira
 *   0, "039" vira 39, "009" vira 9). Isso quebrava a detecção de pedido já
 *   processado (comparava "0" com "00" e nunca batia), o que duplicaria o
 *   pedido e criaria uma falsa revisão a cada reprocessamento. As colunas de
 *   código passam a ser gravadas como TEXTO e toda comparação normaliza o
 *   valor, para que as linhas antigas continuem batendo.
 *
 * V1.4 — correção: o desvio do pedido comparava o TOTAL do pedido contra
 *   a tabela só dos itens comparáveis. Todo item sem preço de tabela entrava
 *   como "vendido acima" (ex.: MG MOTOR +11,7%, que era só um item sem
 *   tabela). Agora os dois lados usam apenas os itens com tabela.
 *
 * FUNÇÕES PARA USAR NO EDITOR
 *   configurarPedidos()      cria pastas e abas
 *   processarPedidos()       processa tudo que está na pasta de entrada
 *   testarPedido(nome)       lê UM PDF e imprime o resultado, sem gravar
 *   dumpTexto(nome)          imprime o texto convertido cru (diagnóstico)
 *   situacaoPedidos()        resumo da carteira
 *   limparTestePedidos(ocs)  V1.1 — apaga da base pedidos/itens/revisões
 *                            de uma lista de OCs (uso: cargas de teste)
 **********************************************************************/

/* ====================== PARÂMETROS ====================== */

var PV_RAIZ        = 'FLUXO VENDA DYNAMICS';
var PV_ENTRADA     = '03 Pedidos de Venda Novos';
var PV_PROCESSADOS = '04 Pedidos de Venda Processados';
var PV_FALHAS      = '05 Falhas';

var ABA_PEDIDOS  = 'Pedidos';
var ABA_ITENS    = 'ItensVenda';
var ABA_REVISOES = 'Revisoes';
var ABA_PVERROS  = 'PedidoErros';

var FATOR_CHEIO  = 2.5;     // valor do PDF é 40% do valor de venda
var PV_FUSO      = 'America/Sao_Paulo';
var PV_ORCAMENTO = 4.5 * 60 * 1000;

var COLS_PEDIDOS = ['id_pedido','oc','vigente','cli_pailon','cliente','ano','sequencial','revisao',
  'numero_erp','data_pedido','data_entrega','cond_pagto','observacao','revenda','cidade','uf',
  'qtd_itens','total_40','total_cheio','total_tabela','desvio_valor','desvio_pct',
  'confere_valor','confere_qtde','status','oc_divergente','arquivo','data_carga'];

var COLS_ITENS = ['id_item','id_pedido','oc','revisao','seq','codigo','descricao','unidade','qtde',
  'preco_40','preco_cheio','total_cheio','preco_tabela','rotulo_ref','desvio_pct','desvio_valor',
  'status_preco','data_rev_tabela'];

var COLS_REVISOES = ['data_hora','oc','rev_de','rev_para','classe','tipo','codigo','descricao','de','para'];

var COLS_PVERROS = ['data_hora','arquivo','motivo','detalhe'];

var RE_CODIGO = /\b(\d{3}\.\d{2}\.\d{3})\b/;
var RE_NUM    = /^-?\d{1,3}(?:\.\d{3})*(?:,\d+)?$|^-?\d+(?:,\d+)?$/;

/* ====================== CONFIGURAÇÃO ====================== */

function configurarPedidos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raiz = pvPasta_(null, PV_RAIZ);
  var ent  = pvPasta_(raiz, PV_ENTRADA);
  var pro  = pvPasta_(raiz, PV_PROCESSADOS);
  var fal  = pvPasta_(raiz, PV_FALHAS);

  pvAba_(ss, ABA_PEDIDOS,  COLS_PEDIDOS);
  pvAba_(ss, ABA_ITENS,    COLS_ITENS);
  pvAba_(ss, ABA_REVISOES, COLS_REVISOES);
  pvAba_(ss, ABA_PVERROS,  COLS_PVERROS);

  var cfg = ss.getSheetByName('Config');
  if (!cfg) { cfg = ss.insertSheet('Config'); cfg.appendRow(['parametro','valor']); }
  pvCfgGrava_(cfg, 'pasta_pedidos_entrada_id', ent.getId());
  pvCfgGrava_(cfg, 'pasta_pedidos_processados_id', pro.getId());
  pvCfgGrava_(cfg, 'pasta_falhas_id', fal.getId());
  pvCfgGrava_(cfg, 'fator_cheio', String(FATOR_CHEIO));

  Logger.log('Pastas e abas de pedidos prontas.');
  Logger.log('Entrada: ' + ent.getUrl());
}

function pvPasta_(pai, nome) {
  var it = pai ? pai.getFoldersByName(nome) : DriveApp.getFoldersByName(nome);
  if (it.hasNext()) return it.next();
  return pai ? pai.createFolder(nome) : DriveApp.createFolder(nome);
}

/* Colunas que guardam código e NÃO podem virar número: o Sheets transforma
   "00" em 0, "039" em 39 e "009" em 9, o que quebra a comparação de revisão
   e a busca do cliente. Formatadas como texto na configuração. */
var COLS_TEXTO = {
  'Pedidos'   : [2, 4, 6, 7, 8],   // oc, cli_pailon, ano, sequencial, revisao
  'ItensVenda': [3, 4, 5, 6],      // oc, revisao, seq, codigo
  'Revisoes'  : [2, 3, 4, 7]       // oc, rev_de, rev_para, codigo
};

function pvAba_(ss, nome, cols) {
  var sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var tx = COLS_TEXTO[nome];
  if (tx) {
    for (var i = 0; i < tx.length; i++) {
      if (tx[i] <= sh.getMaxColumns()) {
        sh.getRange(1, tx[i], sh.getMaxRows(), 1).setNumberFormat('@');
      }
    }
  }
  return sh;
}

/* Normaliza um código que pode ter voltado do Sheets como número:
   0 -> "00", 1 -> "01", 39 -> "039". Usado em toda comparação de revisão e
   de código de cliente, para que linhas antigas (gravadas antes da coluna
   virar texto) continuem batendo. */
function pvPad_(v, n) {
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (s === '') return '';
  if (!/^\d+$/.test(s)) return s;
  while (s.length < n) s = '0' + s;
  return s;
}

function pvRev_(v) { return pvPad_(v, 2); }

function pvCfgGrava_(cfg, chave, valor) {
  var v = cfg.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === chave) { cfg.getRange(i + 1, 2).setValue(valor); return; }
  }
  cfg.appendRow([chave, valor]);
}

function pvCfg_(chave, padrao) {
  var cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!cfg) return padrao;
  var v = cfg.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === chave) return String(v[i][1]).trim();
  }
  return padrao;
}

/* ====================== LEITURA DO PDF ====================== */

/** Converte o PDF em Documento Google temporário e devolve o texto. */
function pvTexto_(arq) {
  var doc = Drive.Files.copy(
    { title: '__tmp__' + arq.getName(), mimeType: MimeType.GOOGLE_DOCS },
    arq.getId()
  );
  try {
    return DocumentApp.openById(doc.id).getBody().getText();
  } finally {
    try { DriveApp.getFileById(doc.id).setTrashed(true); } catch (e) {}
  }
}

function pvLimpa_(l) {
  return String(l).replace(/;/g, ' ').replace(/\s+/g, ' ').trim();
}

function pvNum_(t) {
  if (t === null || t === undefined) return null;
  var s = String(t).trim();
  if (!RE_NUM.test(s)) return null;
  var f = parseFloat(s.replace(/\./g, '').replace(',', '.'));
  return isNaN(f) ? null : f;
}

/** Cabeçalho do pedido. Os rótulos são casados de forma tolerante a acento. */
function pvCabecalho_(txt) {
  var h = {}, m;
  m = /Pedido Nr\.\s*:\s*;?\s*(\d+)/.exec(txt);                     h.numero = m ? m[1] : '';
  m = /Data Ped\.\s*:\s*;?\s*(\d{2}\/\d{2}\/\d{2})/.exec(txt);      h.data_pedido = m ? m[1] : '';
  m = /Previs[aã]o\s+Entrega\s*:\s*;?\s*(\d{2}\/\d{2}\/\d{2})/.exec(txt); h.data_entrega = m ? m[1] : '';
  m = /OC\s*:\s*;?\s*(\d{1,3}\.\d{2}\.\d{3}-\d{2})/.exec(txt);      var ocImp = m ? m[1] : '';
  m = /Nr\. Pedido Rep\.\s*:\s*;?\s*(\d{10})/.exec(txt);            var rep = m ? m[1] : '';

  // O campo OC perde o zero à esquerda (39.26.063 em vez de 039.26.063).
  // O Nr. Pedido Rep. traz o mesmo número sem pontuação e com o zero.
  if (rep) {
    h.oc = rep.substr(0,3) + '.' + rep.substr(3,2) + '.' + rep.substr(5,3) + '-' + rep.substr(8,2);
  } else {
    h.oc = ocImp;
  }
  h.oc_impresso = ocImp;
  h.oc_divergente = (ocImp && rep && ocImp.replace(/^(\d\d)\./, '0$1.') !== h.oc) ? 'SIM' : 'NAO';

  if (h.oc) {
    var p = h.oc.split('.');
    h.cli_pailon = p[0];
    h.ano        = p[1];
    h.sequencial = p[2].split('-')[0];
    h.revisao    = h.oc.slice(-2);
  } else { h.cli_pailon = ''; h.ano = ''; h.sequencial = ''; h.revisao = ''; }

  m = /Observa[cç][aã]o\s*:\s*;?\s*([^;\n]+)/.exec(txt);            h.observacao = m ? m[1].trim() : '';
  var partes = h.observacao ? h.observacao.split(' - ') : [];
  h.cliente_obs = partes[0] ? partes[0].trim() : '';
  h.revenda     = partes[1] ? partes[1].trim() : '';
  h.cidade      = partes.length >= 3 ? partes[partes.length - 2].trim() : '';
  h.uf          = partes.length >= 2 ? partes[partes.length - 1].trim() : '';
  if (/^[A-Z]{2}$/.test(h.cidade)) { h.uf = h.cidade; h.cidade = ''; }

  m = /Cond\. Pagto\.\s*:\s*;?\s*([^;\n]+)/.exec(txt);              h.cond_pagto = m ? m[1].trim() : '';
  m = /Total\s+das\s+Previs[oõ]es\s*:\s*;?\s*([\d.,]+)/.exec(txt);  h.total_pedido = m ? pvNum_(m[1]) : null;
  m = /^([\d.]+,\d+)\s*\n\s*Peso Bruto/m.exec(txt);                 h.qtde_total = m ? pvNum_(m[1]) : null;
  return h;
}

/**
 * Itens — V1.2.
 *
 * A conversão do PDF para Documento Google NÃO respeita uma linha por item:
 * dependendo do layout ela junta TODOS os itens de uma página numa linha só,
 * e ainda cola o rodapé ("Qtde. de Produtos...") no fim dessa mesma linha.
 * Por isso a leitura não pode ser por linha — o texto inteiro é tratado como
 * um fluxo único de palavras e cada CÓDIGO DYN encontrado abre um item.
 *
 * Para cada código:
 *   quantidade  = último número ANTES do código
 *   seq         = número inteiro imediatamente antes da quantidade, quando
 *                 existir (layout com coluna Seq.); senão, a ordem do item
 *   os 5 números  = PRIMEIRA sequência de 5 números seguidos DEPOIS da
 *                 descrição: unitário, frete, valor total, %IPI e valor IPI
 *   unidade     = palavra não-numérica logo antes desses 5 números
 *   descrição   = o que está entre o código e a unidade
 *
 * Pegar os 5 PRIMEIROS números (e não os 5 últimos da linha) é o que permite
 * ler vários itens grudados na mesma linha sem embaralhar valor de um item
 * com o de outro.
 */
var RE_COD_TOKEN = /^\d{3}\.\d{2}\.\d{3}$/;
var RE_RODAPE    = /(Qtde\.|Vlr Bruto|Total de Volumes|Peso |Cubagem|TOTAL|VLR TOTAL|I\.C\.M\.S|Base C|Al[ií]quota|DYNAMICS METALURGICA|CNPJ|P[áa]gina|Comercial|L[óo]gica|P E D I D O|Previs)/i;
/* Palavras que marcam início de cabeçalho/rodapé de página — usadas para
   cortar a continuação da descrição quando o item é o último da página. */
var RE_CORTE     = /^(DYNAMICS|METALURGICA|LTDA|CNPJ:?|I\.E\.:?|Data:?|Hora:?|P[áa]gina:?|Comercial|L[óo]gica|Qtde\.|Vlr|Vlr\.|Total|TOTAL|Peso|Cubagem|I\.C\.M\.S\.|Al[ií]quota|Base|VLR|Representante|Cond\.|Observa[cç][aã]o:?|Transportador:?|Redespacho:?|[»«])$/i;

function pvItens_(linhas) {
  // fluxo único de palavras, na ordem de leitura do documento
  var toks = [];
  for (var i = 0; i < linhas.length; i++) {
    var l = pvLimpa_(linhas[i]);
    if (!l) continue;
    var t = l.split(' ');
    for (var j = 0; j < t.length; j++) if (t[j] !== '') toks.push(t[j]);
  }

  // posição de todos os códigos do documento
  var cods = [];
  for (var p = 0; p < toks.length; p++) if (RE_COD_TOKEN.test(toks[p])) cods.push(p);

  var out = [], fimAnterior = 0;
  for (var c = 0; c < cods.length; c++) {
    var ip = cods[c];

    // --- quantidade e seq: olhando só o trecho entre o item anterior e este
    var ini = Math.max(fimAnterior, 0);
    var qi = -1;
    for (var a = ip - 1; a >= ini; a--) { if (RE_NUM.test(toks[a])) { qi = a; break; } }
    if (qi < 0) continue;
    var qtde = pvNum_(toks[qi]);
    if (qtde === null) continue;
    var seq = String(out.length + 1);
    if (qi - 1 >= ini && /^\d{1,4}$/.test(toks[qi - 1])) seq = toks[qi - 1];

    // --- primeira sequência de 5 números seguidos depois do código
    var n0 = -1;
    for (var d = ip + 1; d + 4 < toks.length; d++) {
      if (RE_NUM.test(toks[d]) && RE_NUM.test(toks[d+1]) && RE_NUM.test(toks[d+2]) &&
          RE_NUM.test(toks[d+3]) && RE_NUM.test(toks[d+4])) { n0 = d; break; }
      // se esbarrar no próximo código antes de achar os 5 números, desiste
      if (RE_COD_TOKEN.test(toks[d])) break;
    }
    if (n0 < 0) continue;

    var unit  = pvNum_(toks[n0]);
    var frete = pvNum_(toks[n0 + 1]);
    var vtot  = pvNum_(toks[n0 + 2]);
    var ipipc = pvNum_(toks[n0 + 3]);
    var ipivl = pvNum_(toks[n0 + 4]);
    if (unit === null || vtot === null) continue;

    // --- unidade e descrição
    var um = (n0 - 1 > ip && !RE_NUM.test(toks[n0 - 1])) ? toks[n0 - 1] : '';
    var desc = toks.slice(ip + 1, um ? n0 - 1 : n0).join(' ').trim();

    // --- continuação da descrição truncada (layout com coluna Seq., corta em
    //     25 caracteres e joga o resto depois dos números)
    if (desc.length >= 24) {
      var cont = [];
      for (var e = n0 + 5; e < toks.length; e++) {
        // para no próximo item, em qualquer número, ou ao esbarrar no
        // cabeçalho/rodapé da página seguinte
        if (RE_NUM.test(toks[e]) || RE_COD_TOKEN.test(toks[e]) || RE_CORTE.test(toks[e])) break;
        cont.push(toks[e]);
        if (cont.length >= 4) break;
      }
      var txtCont = cont.join(' ').trim();
      if (txtCont && txtCont.length <= 40 && !RE_RODAPE.test(txtCont)) {
        desc += (txtCont.charAt(0) === '-' ? ' ' : '') + txtCont;
      }
    }

    out.push({
      seq: seq, codigo: toks[ip], descricao: desc, unidade: um, qtde: qtde,
      preco_40: unit, total_40: vtot,
      preco_cheio: Math.round(unit * FATOR_CHEIO * 10000) / 10000,
      total_cheio: Math.round(vtot * FATOR_CHEIO * 100) / 100,
      frete: frete, ipi_pct: ipipc, ipi_vlr: ipivl
    });
    fimAnterior = n0 + 5;
  }
  return out;
}

/**
 * Quantidade total de produtos impressa no rodapé. Na conversão do Google a
 * posição desse número varia: às vezes na linha logo abaixo de "Qtde. de
 * Produtos", às vezes algumas linhas acima. Procura, numa janela de 5 linhas
 * ao redor do rótulo, uma linha que contenha SÓ um número — e fica com a mais
 * próxima do rótulo.
 */
function pvQtdeTotal_(linhas) {
  var alvo = -1;
  for (var i = 0; i < linhas.length; i++) {
    if (/Qtde\.\s*de\s*Produtos/i.test(linhas[i])) { alvo = i; break; }
  }
  if (alvo < 0) return null;
  var melhor = null, dist = 99;
  for (var j = Math.max(0, alvo - 5); j <= Math.min(linhas.length - 1, alvo + 5); j++) {
    var l = pvLimpa_(linhas[j]);
    if (!/^[\d.]+,\d+$/.test(l)) continue;
    var v = pvNum_(l);
    if (v === null || v <= 0) continue;
    var d2 = Math.abs(j - alvo);
    if (d2 < dist) { dist = d2; melhor = v; }
  }
  return melhor;
}

/** Lê um PDF inteiro e devolve {cab, itens}, com a conferência aritmética feita. */
function pvLer_(arq) {
  var txt = pvTexto_(arq);
  var linhas = txt.split('\n');
  var h = pvCabecalho_(txt);
  var it = pvItens_(linhas);
  h.qtde_total = pvQtdeTotal_(linhas);

  var soma = 0, somaQ = 0;
  for (var i = 0; i < it.length; i++) { soma += it[i].total_40; somaQ += it[i].qtde; }
  soma  = Math.round(soma * 100) / 100;
  somaQ = Math.round(somaQ * 10000) / 10000;

  h.qtd_itens      = it.length;
  h.total_40       = soma;
  h.total_cheio    = Math.round(soma * FATOR_CHEIO * 100) / 100;
  h.confere_valor  = (h.total_pedido !== null && Math.abs(soma - h.total_pedido) < 0.02) ? 'SIM' : 'NAO';

  // A conferência de quantidade é auxiliar: quando o rodapé não traz o número,
  // ela fica "SEM REFERENCIA" e NÃO reprova o pedido — quem manda é o valor,
  // que já denuncia qualquer item perdido na leitura.
  if (h.qtde_total === null) h.confere_qtde = 'SEM REFERENCIA';
  else h.confere_qtde = (Math.abs(somaQ - h.qtde_total) < 0.02) ? 'SIM' : 'NAO';

  h.status = (h.confere_valor === 'SIM' && h.confere_qtde !== 'NAO' && it.length > 0) ? 'OK-AUTO' : 'REVISAR';
  return { cab: h, itens: it };
}

/* ====================== PROCESSAMENTO ====================== */

function processarPedidos() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { Logger.log('Outra execução em andamento.'); return; }
  var t0 = Date.now();
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var ent = DriveApp.getFolderById(pvCfg_('pasta_pedidos_entrada_id', ''));
    var pro = DriveApp.getFolderById(pvCfg_('pasta_pedidos_processados_id', ''));
    var fal = DriveApp.getFolderById(pvCfg_('pasta_falhas_id', ''));
    var shP = ss.getSheetByName(ABA_PEDIDOS);
    var shI = ss.getSheetByName(ABA_ITENS);
    var shR = ss.getSheetByName(ABA_REVISOES);
    var agora = Utilities.formatDate(new Date(), PV_FUSO, 'dd/MM/yyyy HH:mm');

    var arquivos = [], it = ent.getFiles();
    while (it.hasNext()) arquivos.push(it.next());
    if (!arquivos.length) { Logger.log('Pasta de entrada vazia.'); return; }

    var n = 0, nDup = 0, nRev = 0;
    for (var k = 0; k < arquivos.length; k++) {
      if (Date.now() - t0 > PV_ORCAMENTO) { Logger.log('Tempo esgotado — rode de novo.'); break; }
      var arq = arquivos[k], nome = arq.getName();
      try {
        var r = pvLer_(arq);
        if (!r.cab.oc) throw new Error('Não foi possível ler o número do pedido (campo OC).');
        if (!r.itens.length) throw new Error('Nenhum item reconhecido.');

        var ant = pvPedidoVigente_(shP, r.cab.oc);
        if (ant && ant.revisao === pvRev_(r.cab.revisao)) {
          pvErro_(ss, agora, nome, 'PEDIDO JA PROCESSADO',
                  r.cab.oc + ' revisão ' + r.cab.revisao + ' já está na base');
          arq.moveTo(pro); nDup++;
          Logger.log(nome + ' -> já processado, ignorado');
          continue;
        }

        pvGravar_(ss, shP, shI, r, nome, agora);

        if (ant) {
          var difs = pvDiff_(shI, ant, r);
          pvGravarRevisoes_(shR, agora, r.cab, ant.revisao, difs);
          nRev++;
          Logger.log(nome + ' -> rev ' + ant.revisao + ' para ' + r.cab.revisao + ', ' + difs.length + ' mudanças');
        } else {
          Logger.log(nome + ' -> ' + r.cab.oc + ' rev ' + r.cab.revisao + ', ' +
                     r.itens.length + ' itens, R$ ' + r.cab.total_cheio + ' [' + r.cab.status + ']');
        }
        arq.moveTo(pro); n++;
      } catch (e) {
        pvErro_(ss, agora, nome, 'FALHA NA LEITURA', String(e));
        arq.moveTo(fal);
        Logger.log('FALHA em ' + nome + ': ' + e);
      }
    }
    Logger.log('--- ' + n + ' pedido(s), ' + nRev + ' revisão(ões), ' + nDup + ' duplicado(s) ---');
  } finally {
    lock.releaseLock();
  }
}

/** Devolve o pedido vigente daquele OC, ou null. */
function pvPedidoVigente_(shP, oc) {
  if (shP.getLastRow() < 2) return null;
  var v = shP.getRange(2, 1, shP.getLastRow() - 1, COLS_PEDIDOS.length).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    if (String(v[i][1]).trim() === oc && String(v[i][2]).trim() === 'SIM') {
      return { linha: i + 2, id_pedido: v[i][0], oc: v[i][1], revisao: pvRev_(v[i][7]),
               data_entrega: v[i][10], total_cheio: v[i][18] };
    }
  }
  return null;
}

function pvGravar_(ss, shP, shI, r, nome, agora) {
  var h = r.cab;

  // a revisão anterior deixa de ser vigente, mas continua na base
  if (shP.getLastRow() >= 2) {
    var n = shP.getLastRow() - 1;
    var ocs = shP.getRange(2, 2, n, 1).getValues();
    var vig = shP.getRange(2, 3, n, 1).getValues();
    var mudou = false;
    for (var i = 0; i < n; i++) {
      if (String(ocs[i][0]).trim() === h.oc && String(vig[i][0]).trim() === 'SIM') {
        vig[i][0] = 'NAO'; mudou = true;
      }
    }
    if (mudou) shP.getRange(2, 3, n, 1).setValues(vig);
  }

  var idPedido = 'PV' + h.oc.replace(/[.\-]/g, '');
  var nomeCli = pvNomeCliente_(h.cli_pailon) || h.cliente_obs;

  // itens com a comparação de preço já resolvida
  var linhasIt = [], totTab = 0, totComp = 0, comRef = 0;
  for (var j = 0; j < r.itens.length; j++) {
    var i2 = r.itens[j];
    var cmp = { status: 'SEM PREÇO EM TABELA', referencia: '', rotulo: '', variacao: '', data_rev: '' };
    try { cmp = compararPreco(i2.codigo, i2.preco_cheio); } catch (e) {}
    var desvioV = '';
    if (cmp.referencia) {
      totTab  += cmp.referencia * i2.qtde;
      totComp += i2.total_cheio;
      comRef++;
      desvioV = Math.round((i2.preco_cheio - cmp.referencia) * i2.qtde * 100) / 100;
    }
    linhasIt.push([idPedido + '-' + (j + 1), idPedido, h.oc, h.revisao, i2.seq, i2.codigo,
      i2.descricao, i2.unidade, i2.qtde, i2.preco_40, i2.preco_cheio, i2.total_cheio,
      cmp.referencia || '', cmp.rotulo || '', cmp.variacao === null ? '' : cmp.variacao,
      desvioV, cmp.status, cmp.data_rev || '']);
  }
  totTab  = Math.round(totTab * 100) / 100;
  totComp = Math.round(totComp * 100) / 100;
  // V1.4: o desvio compara SÓ os itens que têm tabela, dos dois lados.
  // Antes usava o total do pedido inteiro, e item sem tabela virava "acima".
  var desvio = comRef ? Math.round((totComp - totTab) * 100) / 100 : '';
  var desvioPct = (comRef && totTab) ? Math.round((totComp - totTab) / totTab * 10000) / 100 : '';

  shP.appendRow([idPedido, h.oc, 'SIM', h.cli_pailon, nomeCli, h.ano, h.sequencial, h.revisao,
    h.numero, h.data_pedido, h.data_entrega, h.cond_pagto, h.observacao, h.revenda, h.cidade, h.uf,
    h.qtd_itens, h.total_40, h.total_cheio, comRef ? totTab : '', desvio, desvioPct,
    h.confere_valor, h.confere_qtde, h.status, h.oc_divergente, nome, agora]);

  if (linhasIt.length) {
    shI.getRange(shI.getLastRow() + 1, 1, linhasIt.length, COLS_ITENS.length).setValues(linhasIt);
  }
}

function pvNomeCliente_(codPailon) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
  if (!sh || sh.getLastRow() < 2) return '';
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < v.length; i++) {
    // pvPad_ nos dois lados: o código do cliente na aba Clientes pode ter sido
    // gravado como número (039 vira 39) e não bateria com o código do pedido
    if (pvPad_(v[i][1], 3) === pvPad_(codPailon, 3)) return String(v[i][2]).trim();
  }
  return '';
}

/* ====================== DIFERENÇA ENTRE REVISÕES ====================== */

/**
 * Compara a revisão nova com a anterior e classifica cada mudança:
 *   COMERCIAL    — não afeta a fábrica (condição de pagamento, observação)
 *   PROGRAMACAO  — mexe na data de entrega: avisa o PCP
 *   PRODUTIVA    — item, quantidade ou descrição: trava e avisa
 */
function pvDiff_(shI, ant, r) {
  var difs = [];
  var antigos = {};
  if (shI.getLastRow() >= 2) {
    var v = shI.getRange(2, 1, shI.getLastRow() - 1, COLS_ITENS.length).getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][2]).trim() === String(ant.oc).trim() && pvRev_(v[i][3]) === ant.revisao) {
        antigos[String(v[i][5]).trim()] = { qtde: Number(v[i][8]), preco: Number(v[i][10]), desc: v[i][6] };
      }
    }
  }
  var novos = {};
  for (var j = 0; j < r.itens.length; j++) novos[r.itens[j].codigo] = r.itens[j];

  for (var c in novos) {
    if (!antigos[c]) {
      difs.push(['PRODUTIVA', 'ITEM INCLUIDO', c, novos[c].descricao, '', novos[c].qtde]);
    } else {
      if (Math.abs(antigos[c].qtde - novos[c].qtde) > 0.0001) {
        difs.push(['PRODUTIVA', 'QUANTIDADE ALTERADA', c, novos[c].descricao, antigos[c].qtde, novos[c].qtde]);
      }
      if (Math.abs(antigos[c].preco - novos[c].preco_cheio) > 0.01) {
        difs.push(['COMERCIAL', 'PRECO ALTERADO', c, novos[c].descricao, antigos[c].preco, novos[c].preco_cheio]);
      }
    }
  }
  for (var c2 in antigos) {
    if (!novos[c2]) difs.push(['PRODUTIVA', 'ITEM EXCLUIDO', c2, antigos[c2].desc, antigos[c2].qtde, '']);
  }
  if (String(ant.data_entrega).trim() !== String(r.cab.data_entrega).trim()) {
    difs.push(['PROGRAMACAO', 'DATA DE ENTREGA ALTERADA', '', '', ant.data_entrega, r.cab.data_entrega]);
  }
  var difTot = Math.abs(Number(ant.total_cheio) - r.cab.total_cheio);
  if (difTot > 0.01) {
    difs.push(['COMERCIAL', 'TOTAL DO PEDIDO ALTERADO', '', '', ant.total_cheio, r.cab.total_cheio]);
  }
  return difs;
}

function pvGravarRevisoes_(shR, agora, cab, revDe, difs) {
  if (!difs.length) {
    shR.appendRow([agora, cab.oc, revDe, cab.revisao, 'COMERCIAL', 'SEM MUDANCA DETECTADA', '', '', '', '']);
    return;
  }
  var linhas = difs.map(function (d) {
    return [agora, cab.oc, revDe, cab.revisao, d[0], d[1], d[2], d[3], d[4], d[5]];
  });
  shR.getRange(shR.getLastRow() + 1, 1, linhas.length, COLS_REVISOES.length).setValues(linhas);
}

function pvErro_(ss, agora, arquivo, motivo, detalhe) {
  var sh = ss.getSheetByName(ABA_PVERROS);
  if (sh) sh.appendRow([agora, arquivo, motivo, String(detalhe).substring(0, 400)]);
}

/* ====================== DIAGNÓSTICO ====================== */

function pvArquivo_(nome) {
  var ent = DriveApp.getFolderById(pvCfg_('pasta_pedidos_entrada_id', ''));
  if (nome) {
    var it = ent.getFilesByName(nome);
    return it.hasNext() ? it.next() : null;
  }
  var it2 = ent.getFiles(), mais = null;
  while (it2.hasNext()) { var f = it2.next(); if (!mais || f.getDateCreated() > mais.getDateCreated()) mais = f; }
  return mais;
}

function testarPedido(nome) {
  var arq = pvArquivo_(nome);
  if (!arq) { Logger.log('Arquivo não encontrado na pasta de entrada.'); return; }
  var r = pvLer_(arq);
  var h = r.cab;
  Logger.log('ARQUIVO: ' + arq.getName());
  Logger.log('OC ' + h.oc + ' (impresso: ' + h.oc_impresso + ') · revisão ' + h.revisao +
             ' · divergência OC: ' + h.oc_divergente);
  Logger.log('Cliente Pailon ' + h.cli_pailon + ' · ' + h.cliente_obs + ' · ' + h.revenda +
             ' · ' + h.cidade + '/' + h.uf);
  Logger.log('Pedido ERP ' + h.numero + ' · emissão ' + h.data_pedido +
             ' · ENTREGA ' + h.data_entrega + ' · ' + h.cond_pagto);
  Logger.log('Itens ' + h.qtd_itens + ' · soma ' + h.total_40 + ' vs total ' + h.total_pedido +
             ' · confere valor: ' + h.confere_valor + ' · confere qtde: ' + h.confere_qtde);
  Logger.log('VALOR CHEIO (x' + FATOR_CHEIO + '): R$ ' + h.total_cheio + '  >>> ' + h.status);
  Logger.log('--- itens ---');
  for (var i = 0; i < r.itens.length; i++) {
    var x = r.itens[i], ref = '';
    try {
      var c = compararPreco(x.codigo, x.preco_cheio);
      ref = c.referencia ? ('tabela ' + c.referencia + ' · ' + c.variacao + '% · ' + c.status)
                         : c.status;
    } catch (e) { ref = 'arquivo Tabelas não carregado'; }
    Logger.log([x.codigo, x.descricao, x.unidade, x.qtde, x.preco_40 + ' -> ' + x.preco_cheio, ref].join(' | '));
  }
}

/** Texto cru da conversão — use quando o parser não reconhecer algo. */
function dumpTexto(nome) {
  var arq = pvArquivo_(nome);
  if (!arq) { Logger.log('Arquivo não encontrado.'); return; }
  var linhas = pvTexto_(arq).split('\n');
  Logger.log('ARQUIVO: ' + arq.getName() + ' · ' + linhas.length + ' linhas');
  for (var i = 0; i < linhas.length; i++) Logger.log(i + ': ' + linhas[i]);
}

function situacaoPedidos() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PEDIDOS);
  if (!sh || sh.getLastRow() < 2) { Logger.log('Nenhum pedido na base.'); return; }
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_PEDIDOS.length).getValues();
  var tot = 0, n = 0, rev = 0, revisar = 0, porCli = {};
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][2]).trim() !== 'SIM') continue;
    n++; tot += Number(v[i][18]) || 0;
    if (pvRev_(v[i][7]) !== '00') rev++;
    if (String(v[i][24]).trim() !== 'OK-AUTO') revisar++;
    var c = (v[i][4] || v[i][3]);
    porCli[c] = (porCli[c] || 0) + (Number(v[i][18]) || 0);
  }
  Logger.log('Pedidos vigentes: ' + n + ' · valor cheio R$ ' + Math.round(tot * 100) / 100);
  Logger.log('Com revisão posterior à 00: ' + rev + ' · precisando de revisão manual: ' + revisar);
  Logger.log('--- por cliente ---');
  Object.keys(porCli).sort().forEach(function (c) {
    Logger.log(c + ' | R$ ' + Math.round(porCli[c] * 100) / 100);
  });
}

/* ====================== LIMPEZA DE TESTE ====================== */

/**
 * Apaga da base (abas Pedidos, ItensVenda e Revisoes) tudo que pertence
 * aos OCs informados — útil para tirar uma carga de teste antes de ir
 * para produção. NÃO mexe nos arquivos da pasta "04 Pedidos de Venda
 * Processados": se quiser reprocessar o mesmo PDF depois, mova-o de volta
 * manualmente para "03 Pedidos de Venda Novos".
 *
 * Uso no editor, indo em Executar com esta função selecionada:
 *   limparTestePedidos(['153.26.011-00','135.26.009-00','039.26.063-00'])
 */
function limparTestePedidos(ocs) {
  if (!ocs || !ocs.length) { Logger.log('Informe uma lista de OCs. Ex.: ["153.26.011-00"]'); return; }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var set = {};
  ocs.forEach(function (o) { set[String(o).trim()] = true; });

  var nPed = pvExcluirLinhas_(ss.getSheetByName(ABA_PEDIDOS), 2, set);    // coluna 2 = oc
  var nIt  = pvExcluirLinhas_(ss.getSheetByName(ABA_ITENS), 3, set);     // coluna 3 = oc
  var nRev = pvExcluirLinhas_(ss.getSheetByName(ABA_REVISOES), 2, set);  // coluna 2 = oc

  Logger.log('Removidos -> Pedidos: ' + nPed + ' | ItensVenda: ' + nIt + ' | Revisoes: ' + nRev);
  Logger.log('Os PDFs continuam em "04 Pedidos de Venda Processados" — mova de volta pra ' +
             '"03 Pedidos de Venda Novos" se quiser reprocessar.');
}

/* Apaga, de baixo para cima, toda linha cuja coluna "col" (1-based) bate
   com um dos valores do conjunto — de baixo para cima pra não bagunçar
   os índices durante a exclusão. */
function pvExcluirLinhas_(sh, col, set) {
  if (!sh || sh.getLastRow() < 2) return 0;
  var n = sh.getLastRow() - 1;
  var vals = sh.getRange(2, col, n, 1).getValues();
  var n2 = 0;
  for (var i = n - 1; i >= 0; i--) {
    if (set[String(vals[i][0]).trim()]) { sh.deleteRow(i + 2); n2++; }
  }
  return n2;
}
