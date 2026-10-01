/**********************************************************************
 * FLUXO VENDA DYNAMICS — APLICATIVO WEB
 * Arquivo: App              Versão: V3.5
 *
 * Login por usuário e senha, termo de confidencialidade/LGPD,
 * registro de tudo (acessos, downloads, uploads) na aba LogAcoes.
 * Lê a base que Tabelas e Pedidos alimentam. Não recalcula preço.
 *
 * ABA ACESSOS (editar direto na planilha, sem rodar nada)
 *   usuario | senha | nome | telas | ativo | aceite_termo | ultimo_acesso
 *   telas: DIA, MES, CLIENTES, AUDITORIA, UPLOAD  (separadas por vírgula)
 *          ou TODAS.   ativo: SIM ou NAO.
 *   O usuário SUPORTE não aparece no registro de atividades.
 *   email: e-mail que recebe os avisos.   avisos: PEDIDOS, TABELAS (ou os dois).
 *     PEDIDOS = aviso de novo pedido de venda (Diretoria)
 *     TABELAS = itens vendidos sem tabela / fora da tabela e tabelas vencidas (Orçamento)
 *
 * FUNÇÕES PARA O EDITOR
 *   instalarV2()   cria as abas Acessos e LogAcoes (uma vez só)
 *   situacaoApp()  confere a base e o tempo de carga
 *   ativarAvisosV2_3()  V2.3 — cria as colunas email/avisos e liga os e-mails (uma vez)
 *   testarAvisos()      manda um e-mail de teste para cada destinatário
 **********************************************************************/

var APP_TITULO   = 'Fluxo Venda Dynamics';
var ABA_LOGOS    = 'Logos';
var ABA_ACESSOS  = 'Acessos';
var ABA_LOG      = 'LogAcoes';
var P_LOGOS      = '06 Logos de Clientes';
var P_UP_PEDIDOS = '03 Pedidos de Venda Novos';
var P_UP_TABELAS = '01 Tabelas Atualizadas';
var LIMITE_LOGO  = 45000;
var SESSAO_SEG   = 21600;            // 6 horas
var TERMO_VERSAO = 'V1.1';
var COLS_ACESSOS = ['usuario','senha','nome','telas','ativo','aceite_termo','ultimo_acesso','email','avisos'];
var LINK_APP = 'https://dynmetalmaringa-dotcom.github.io/Pedidos-de-Venda-Dynamics/';
var DIAS_TABELA_VELHA = 180;
var COLS_LOG     = ['data_hora','usuario','perfil','acao','detalhe'];

/* ====================== PUBLICAÇÃO ====================== */

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle(APP_TITULO)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function inc(nome) { return HtmlService.createHtmlOutputFromFile(nome).getContent(); }

/* ====================== INSTALAÇÃO ====================== */

/* Nomes das telas aceitos na coluna "telas" da aba Acessos */
/* V3.0 — telas: VENDA (dia, mes, cli) · RETRABALHO (rdia, rmes, rcli) · ORÇAMENTO (aud)
   · ARQUIVOS (up, reg) · FATURAMENTO (flan, fdash).
   Na coluna "telas": DIA, MES, CLIENTES, RETRABALHO, AUDITORIA, UPLOAD, REGISTROS,
   FATURAMENTO — ou TODAS. Quem tem DIA/MES/CLIENTES vê também o RETRABALHO;
   UPLOAD inclui REGISTROS. */
var TELAS_OK = { DIA: ['dia'], MES: ['mes'], 'MÊS': ['mes'], CLIENTES: ['cli'], VENDA: ['dia', 'mes', 'cli'],
  RETRABALHO: ['rdia', 'rmes', 'rcli'], RETRABALHOS: ['rdia', 'rmes', 'rcli'], AUDITORIA: ['aud'],
  UPLOAD: ['up', 'reg'], REGISTROS: ['reg'], FATURAMENTO: ['flan', 'fdash'], PCP: ['pcar'], CARTEIRA: ['pcar'] };
var TODAS_TELAS = ['dia', 'mes', 'cli', 'rdia', 'rmes', 'rcli', 'aud', 'pcar', 'up', 'reg', 'flan', 'fdash'];
var ABA_FAT = 'Faturamento';
var COLS_FAT = ['data_hora','usuario','nf','data_nf','id_pedido','oc','os','revisao','cliente','codigo','descricao',
  'qtde_pedido','qtde_faturada','unit_pedido_40','unit_faturado_40','total_faturado_40','alterado'];

function instalarV2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA_ACESSOS);
  if (sh && String(sh.getRange(1, 2).getValues()[0][0]).trim() !== 'senha') {
    sh.setName('Acessos_antigo_' + Utilities.formatDate(new Date(), FUSO, 'ddMMyy_HHmm'));
    sh = null;
  }
  var novos = 0;
  if (!sh) {
    sh = ss.insertSheet(ABA_ACESSOS);
    sh.getRange(1, 1, 1, COLS_ACESSOS.length).setValues([COLS_ACESSOS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), COLS_ACESSOS.length).setNumberFormat('@');
    var dir = 'DIA, MES, CLIENTES, AUDITORIA, UPLOAD', orc = 'AUDITORIA, UPLOAD';
    var USU = [['MICHELE','Michele Puma',dir],['PAULA','Paula',dir],['LUCIANO','Luciano',dir],['CESAR','Cesar',dir],
               ['FERNANDA','Fernanda',orc],['JULIANA','Juliana',orc],['SUPORTE','Suporte','TODAS']];
    var linhas = USU.map(function (u) {
      return [u[0], 'Dyn' + Utilities.getUuid().replace(/-/g, '').substring(0, 5).toUpperCase(), u[1], u[2], 'SIM', '', ''];
    });
    sh.getRange(2, 1, linhas.length, COLS_ACESSOS.length).setValues(linhas);
    novos = linhas.length;
  }
  var lg = ss.getSheetByName(ABA_LOG);
  if (!lg) {
    lg = ss.insertSheet(ABA_LOG);
    lg.getRange(1, 1, 1, COLS_LOG.length).setValues([COLS_LOG]).setFontWeight('bold');
    lg.setFrozenRows(1);
  }
  if (!ss.getSheetByName(ABA_LOGOS)) {
    var shL = ss.insertSheet(ABA_LOGOS);
    shL.getRange(1, 1, 1, 5).setValues([['cod_dyn','cliente','mime','base64','atualizado']]).setFontWeight('bold');
    shL.setFrozenRows(1);
  }
  Logger.log(novos + ' usuário(s) criado(s) com senha provisória — veja a aba Acessos.');
  Logger.log('Senha e telas se alteram direto na aba Acessos. Não precisa rodar nada.');
}

function _acharUsuario(usuario) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACESSOS);
  if (!sh || sh.getLastRow() < 2) return null;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_ACESSOS.length).getValues();
  var u = String(usuario || '').trim().toUpperCase();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim().toUpperCase() === u) return { sh: sh, linha: i + 2, r: v[i] };
  }
  return null;
}

/** "DIA, MES, AUDITORIA" -> ['dia','mes','aud']; "TODAS" -> todas */
function _telas(txt) {
  var t = String(txt || '').toUpperCase();
  if (/TODAS|TUDO/.test(t)) return TODAS_TELAS.slice();
  var out = [];
  t.split(/[,;\/\s]+/).forEach(function (x) {
    (TELAS_OK[x.trim()] || []).forEach(function (k) { if (out.indexOf(k) < 0) out.push(k); });
  });
  if (/\bDIA\b|\bMES\b|MÊS|CLIENTES|VENDA/.test(t)) ['rdia', 'rmes', 'rcli'].forEach(function (k) { if (out.indexOf(k) < 0) out.push(k); });
  return TODAS_TELAS.filter(function (k) { return out.indexOf(k) >= 0; });
}

/* ====================== LOGIN E SESSÃO ====================== */

function _cpfOk(c) {
  c = String(c || '').replace(/\D/g, '');
  if (c.length !== 11 || /^(\d)\1{10}$/.test(c)) return false;
  for (var t = 9; t < 11; t++) {
    var s = 0;
    for (var i = 0; i < t; i++) s += Number(c.charAt(i)) * (t + 1 - i);
    if (((s * 10) % 11) % 10 !== Number(c.charAt(t))) return false;
  }
  return true;
}

function login(usuario, senha, aceite) {
  var u = String(usuario || '').trim().toUpperCase();
  var a = _acharUsuario(u);
  if (!a || String(a.r[1]).trim() === '' || String(a.r[1]).trim() !== String(senha || '').trim()) {
    _log({ u: u || '?', p: '' }, 'LOGIN RECUSADO', 'usuário ou senha incorretos');
    return { ok: false, msg: 'Usuário ou senha incorretos.' };
  }
  if (String(a.r[4]).trim().toUpperCase() === 'NAO' || String(a.r[4]).trim().toUpperCase() === 'NÃO') {
    return { ok: false, msg: 'Usuário bloqueado. Fale com a Diretoria.' };
  }
  var telas = _telas(a.r[3]);
  if (!telas.length) return { ok: false, msg: 'Nenhuma tela liberada para este usuário. Fale com a Diretoria.' };
  var aceitou = String(a.r[5]).indexOf(TERMO_VERSAO) >= 0;
  if (!aceitou) {
    var nm = aceite && String(aceite.nome || '').trim().toUpperCase().replace(/\s+/g, ' ');
    var cpf = aceite && String(aceite.cpf || '').replace(/\D/g, '');
    if (!nm || nm.split(' ').length < 2) return { ok: false, termo: true, msg: 'Leia o termo e assine com nome completo e CPF.' };
    if (!_cpfOk(cpf)) return { ok: false, termo: true, msg: 'CPF inválido.' };
    aceite = { nome: nm, cpf: cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4') };
  }
  var agora = _agora();
  var s = { u: u, p: u === 'SUPORTE' ? 'oculto' : 'usuario', n: String(a.r[2] || u).trim(), t: telas };
  if (!aceitou) {
    a.sh.getRange(a.linha, 6).setValue('Termo ' + TERMO_VERSAO + ' assinado em ' + agora + ' por ' + aceite.nome + ' · CPF ' + aceite.cpf);
    _log(s, 'ASSINATURA DO TERMO', 'Confidencialidade e LGPD ' + TERMO_VERSAO + ' · ' + aceite.nome +
         ' · CPF ***.' + aceite.cpf.substring(4, 11) + '-**');
  }
  a.sh.getRange(a.linha, 7).setValue(agora);
  var tk = Utilities.getUuid();
  CacheService.getScriptCache().put('fvd_' + tk, JSON.stringify(s), SESSAO_SEG);
  _log(s, 'LOGIN', telas.join(', '));
  return { ok: true, token: tk, telas: telas, nome: s.n, usuario: u };
}

function sair(tk) {
  var s = _sessao(tk, true);
  if (s) { _log(s, 'SAIR', ''); CacheService.getScriptCache().remove('fvd_' + tk); }
  return true;
}

function _sessao(tk, silencioso) {
  var j = tk ? CacheService.getScriptCache().get('fvd_' + tk) : null;
  if (!j) { if (silencioso) return null; throw new Error('SESSAO_EXPIRADA'); }
  CacheService.getScriptCache().put('fvd_' + tk, j, SESSAO_SEG);
  return JSON.parse(j);
}

function _agora() { return Utilities.formatDate(new Date(), FUSO, 'dd/MM/yyyy HH:mm:ss'); }

function _log(s, acao, detalhe) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ABA_LOG);
    if (!sh) {
      sh = ss.insertSheet(ABA_LOG);
      sh.getRange(1, 1, 1, COLS_LOG.length).setValues([COLS_LOG]);
    }
    sh.appendRow([_agora(), s.u, s.p, acao, String(detalhe || '').substring(0, 400)]);
  } catch (e) {}
}

/** A tela chama quando alguém baixa PDF/Excel ou imprime. */
function registrar(tk, acao, detalhe) {
  var s = _sessao(tk);
  _log(s, String(acao).substring(0, 60), detalhe);
  return true;
}

/* ====================== DADOS PARA AS TELAS ====================== */

function _idx(cols) { var m = {}; cols.forEach(function (c, i) { m[c] = i; }); return m; }
function _r2(n) { return Math.round((Number(n) || 0) * 100) / 100; }

function carregar(tk) {
  var s = _sessao(tk);
  var t0 = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = { eu: { u: s.u, n: s.n, t: s.t }, hoje: Utilities.formatDate(new Date(), FUSO, 'dd/MM/yy'),
              logos: _logos(), logoNomes: _logoNomes(), P: [], tabelas: [], tabItens: {}, logs: [] };

  var shP = ss.getSheetByName(ABA_PEDIDOS), shI = ss.getSheetByName(ABA_ITENS);
  if (shP && shP.getLastRow() > 1) {
    var P = _idx(COLS_PEDIDOS), I = _idx(COLS_ITENS);
    var vp = shP.getRange(2, 1, shP.getLastRow() - 1, COLS_PEDIDOS.length).getValues();
    var vi = shI && shI.getLastRow() > 1
      ? shI.getRange(2, 1, shI.getLastRow() - 1, COLS_ITENS.length).getValues() : [];
    var porPed = {};
    vi.forEach(function (r) {
      var k = String(r[I.id_pedido]).trim() + '|' + pvRev_(r[I.revisao]);
      (porPed[k] = porPed[k] || []).push({
        c: String(r[I.codigo]).trim(), d: String(r[I.descricao]).trim(), un: String(r[I.unidade]).trim(),
        q: Number(r[I.qtde]) || 0, u: Number(r[I.preco_cheio]) || 0, t: Number(r[I.total_cheio]) || 0,
        tb: Number(r[I.preco_tabela]) || 0, st: String(r[I.status_preco] || '').trim(),
        dt: _dataBr(r[I.data_rev_tabela])
      });
    });
    var DYN = _mapaDyn(), REV = _revisoes(), FAT = _faturadoPorItem();
    vp.forEach(function (r) {
      if (String(r[P.vigente]).trim() !== 'SIM') return;
      var id = String(r[P.id_pedido]).trim(), rev = pvRev_(r[P.revisao]);
      var itens = porPed[id + '|' + rev] || [];
      var v = Number(r[P.total_cheio]) || 0;
      var v40 = Number(r[P.total_40]) || v / FATOR_CHEIO;
      var nome = String(r[P.cliente] || '').trim() || String(r[P.cli_pailon]).trim();
      var p = {
        id: id, oc: String(r[P.oc]).trim(), rev: rev, erp: String(r[P.numero_erp]).trim(),
        cli: nome, cp: pvPad_(r[P.cli_pailon], 3), rv: String(r[P.revenda] || '').trim(),
        loc: [String(r[P.cidade] || '').trim(), String(r[P.uf] || '').trim()].filter(String).join(' - '),
        dp: _dataBr(r[P.data_pedido]), ent: _dataBr(r[P.data_entrega]),
        v: _r2(v), v40: _r2(v40), itens: itens,
        dc: (function (x) { return Object.prototype.toString.call(x) === '[object Date]' ? Utilities.formatDate(x, FUSO, 'dd/MM/yyyy HH:mm') : String(x || '').trim(); })(r[P.data_carga]),
        rt: String(r[P.tipo] || '').trim() === 'RETRABALHO', org: String(r[P.os_origem] || '').trim()
      };
      /* faturamento: por OS (vale para qualquer revisão) + código */
      var os = pvOs_(p.oc), fv = 0, comp1 = true, algum = false;
      itens.forEach(function (x) {
        var f = FAT[os + '|' + x.c]; x.fq = f ? _r2(f.q) : 0; x.fv = f ? _r2(f.v) : 0;
        fv += x.fv; if (x.fq > 0) algum = true; if (x.fq + 0.0001 < x.q) comp1 = false;
      });
      p.fv40 = _r2(fv); p.fs = !algum ? '' : (comp1 && itens.length ? 'COMPLETO' : 'PARCIAL');
      p.sv = v < 1;
      p.cd = DYN.cod[p.cp] || _dynPorNome(DYN, nome) ||
             (!p.sv && itens.length ? String(itens[0].c).substring(0, 3) : '') || p.cp;
      var comp = 0, tab = 0, n = 0;
      itens.forEach(function (x) { if (x.tb > 0) { comp += x.t; tab += x.tb * x.q; n++; } });
      p.comp = _r2(comp); p.tab = _r2(tab); p.ct = n;
      var ra = REV[p.oc + '|' + rev];
      if (ra) p.alt = ra;
      out.P.push(p);
    });
  }

  if (s.t.indexOf('aud') >= 0 || s.t.indexOf('up') >= 0) {
    var a = _auditoria();
    out.tabelas = a.tabelas; out.tabItens = a.itens;
    out.sitTab = _situacaoTabelas();
  }
  if (s.t.indexOf('up') >= 0) out.logs = _ultimosLogs(s.p === 'oculto');
  if (s.t.indexOf('reg') >= 0 && !out.logs.length) out.logs = _ultimosLogs(s.p === 'oculto');
  if (s.t.indexOf('flan') >= 0) out.fat = _lancamentos();
  if (!s.t.some(function (k) { return /^(dia|mes|cli|rdia|rmes|rcli|aud|flan|pcar)$/.test(k); })) out.P = [];
  out.ms = Date.now() - t0;
  return out;
}

function _diasDesde(v) {
  if (v === '' || v === null || v === undefined) return null;
  var d = Object.prototype.toString.call(v) === '[object Date]' ? v : null;
  if (!d) {
    var m = /^(\d{2})\/(\d{2})\/(\d{2,4})/.exec(String(v).trim());
    if (!m) return null;
    var a = +m[3]; if (a < 100) a += 2000;
    d = new Date(a, +m[2] - 1, +m[1]);
  }
  var h = new Date(); h.setHours(0, 0, 0, 0);
  return Math.max(0, Math.round((h - d) / 86400000));
}

/** Tabelas por cliente: itens vigentes, revisão, dias sem revisão, duplicados. */
/* ====================== ATUALIZAR TABELAS DE VENDA (V2.6) ====================== */

/** Refaz a comparação de preço de todos os itens dos pedidos vigentes com as
    tabelas vigentes de hoje. Sem lock (quem chama segura o lock). */
function recompararPedidos_(quem) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shI = ss.getSheetByName(ABA_ITENS), shP = ss.getSheetByName(ABA_PEDIDOS), shT = ss.getSheetByName(ABA_PRECOS);
  var res = { itens: 0, mud: 0, ped: 0 };
  if (!shI || shI.getLastRow() < 2 || !shP || shP.getLastRow() < 2) return res;
  TP_CACHE = shT && shT.getLastRow() > 1 ? shT.getRange(2, 1, shT.getLastRow() - 1, COLS_PRECOS.length).getValues() : [];
  CFG_CACHE = null;
  try {
    var P = _idx(COLS_PEDIDOS), I = _idx(COLS_ITENS);
    var vp = shP.getRange(2, 1, shP.getLastRow() - 1, COLS_PEDIDOS.length).getValues();
    var vig = {};
    vp.forEach(function (r) { if (String(r[P.vigente]).trim() === 'SIM') vig[String(r[P.id_pedido]).trim() + '|' + pvRev_(r[P.revisao])] = { comp: 0, tab: 0, n: 0 }; });
    var n = shI.getLastRow() - 1;
    var vi = shI.getRange(2, 1, n, COLS_ITENS.length).getValues();
    var c0 = I.preco_tabela, bloco = vi.map(function (r) { return r.slice(c0, c0 + 6); });
    var memo = {};
    vi.forEach(function (r, i) {
      var k = String(r[I.id_pedido]).trim() + '|' + pvRev_(r[I.revisao]);
      var pd = vig[k]; if (!pd) return;
      var cod = String(r[I.codigo]).trim(), pc = Number(r[I.preco_cheio]) || 0, q = Number(r[I.qtde]) || 0;
      var mk = cod + '|' + pc, cmp = memo[mk];
      if (!cmp) { cmp = { status: 'SEM PREÇO EM TABELA' }; try { cmp = compararPreco(cod, pc); } catch (e) {} memo[mk] = cmp; }
      var ref = cmp.referencia || '';
      var nova = [ref, cmp.rotulo || '', cmp.variacao === null || cmp.variacao === undefined ? '' : cmp.variacao,
                  ref ? Math.round((pc - ref) * q * 100) / 100 : '', cmp.status || '', cmp.data_rev || ''];
      if (Number(bloco[i][0] || 0) !== Number(ref || 0)) res.mud++;
      bloco[i] = nova; res.itens++;
      if (ref) { pd.comp += Number(r[I.total_cheio]) || 0; pd.tab += ref * q; pd.n++; }
    });
    shI.getRange(2, c0 + 1, n, 6).setValues(bloco);
    var cols = vp.map(function (r) {
      var pd = vig[String(r[P.id_pedido]).trim() + '|' + pvRev_(r[P.revisao])];
      if (!pd) return [r[P.total_tabela], r[P.desvio_valor], r[P.desvio_pct]];
      res.ped++;
      if (!pd.n) return ['', '', ''];
      var t = Math.round(pd.tab * 100) / 100, d = Math.round((pd.comp - pd.tab) * 100) / 100;
      return [t, d, t ? Math.round(d / t * 10000) / 100 : ''];
    });
    shP.getRange(2, P.total_tabela + 1, vp.length, 3).setValues(cols);
  } finally { TP_CACHE = null; }
  var info = { dh: _agora(), quem: quem, itens: res.itens, mud: res.mud, ped: res.ped };
  PropertiesService.getScriptProperties().setProperty('rc_ult', JSON.stringify(info));
  _log({ u: String(quem).split(' ')[0], p: 'sistema' }, 'ATUALIZAÇÃO DE TABELAS', res.ped + ' pedidos · ' + res.itens + ' itens · ' + res.mud + ' com preço de tabela alterado');
  return res;
}

/** Botão "Atualizar tabelas de venda" da Auditoria: lê as tabelas que estiverem na
    fila e refaz a comparação de todos os pedidos. */
function atualizarTabelasVenda(tk) {
  var s = _sessao(tk);
  var fila = _pendentes('pasta_tabelas_entrada_id');
  if (fila > 0) { try { processarTabelas(); } catch (e) {} }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { ok: false, msg: 'Outro processamento em andamento. Tente em 1 minuto.' };
  try {
    var r = recompararPedidos_(s.n);
    return { ok: true, msg: (fila ? fila + ' tabela(s) lida(s) · ' : '') + r.ped + ' pedidos recalculados · ' + r.mud + ' item(ns) com preço de tabela alterado.' };
  } finally { lock.releaseLock(); }
}

function _pendentes(idCfg) {
  try { var it = DriveApp.getFolderById(lerCfg_(idCfg, '')).getFiles(), n = 0; while (it.hasNext()) { it.next(); n++; } return n; }
  catch (e) { return 0; }
}

/** Situação para o topo da Auditoria. */
function _situacaoTabelas() {
  var o = { up: null, carga: '', rc: null, fila: _pendentes('pasta_tabelas_entrada_id') };
  try { o.rc = JSON.parse(PropertiesService.getScriptProperties().getProperty('rc_ult') || 'null'); } catch (e) {}
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_LOG);
  if (sh && sh.getLastRow() > 1) {
    var n = Math.min(3000, sh.getLastRow() - 1);
    var v = sh.getRange(sh.getLastRow() - n + 1, 1, n, COLS_LOG.length).getValues();
    for (var i = v.length - 1; i >= 0; i--) {
      if (String(v[i][3]) === 'UPLOAD TABELA') {
        var dh = v[i][0]; if (Object.prototype.toString.call(dh) === '[object Date]') dh = Utilities.formatDate(dh, FUSO, 'dd/MM/yyyy HH:mm:ss');
        o.up = { dh: String(dh), u: String(v[i][1]), d: String(v[i][4]).split(' · ')[0] }; break;
      }
    }
  }
  return o;
}

function _dhNum(s) { var m = /(\d{2})\/(\d{2})\/(\d{4})\s*(\d{2})?:?(\d{2})?/.exec(String(s)); return m ? Number(m[3] + m[2] + m[1] + (m[4] || '00') + (m[5] || '00')) : 0; }

/* ====================== FATURAMENTO (V3.0) ====================== */

function _abaFat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(ABA_FAT);
  if (!sh) {
    sh = ss.insertSheet(ABA_FAT);
    sh.getRange(1, 1, 1, COLS_FAT.length).setValues([COLS_FAT]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 3, sh.getMaxRows(), 1).setNumberFormat('@');
    sh.getRange(1, 6, sh.getMaxRows(), 5).setNumberFormat('@');
  }
  return sh;
}
function _faturadoPorItem() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FAT), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, COLS_FAT.length).getValues().forEach(function (r) {
    var k = String(r[6]).trim() + '|' + String(r[9]).trim();
    m[k] = m[k] || { q: 0, v: 0 }; m[k].q += Number(r[12]) || 0; m[k].v += Number(r[15]) || 0;
  });
  return m;
}
function _lancamentos() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FAT);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = Math.min(4000, sh.getLastRow() - 1);
  return sh.getRange(sh.getLastRow() - n + 1, 1, n, COLS_FAT.length).getValues().map(function (r) {
    return { dh: _txtDh(r[0]), u: String(r[1]), nf: String(r[2]), dnf: _dataBr(r[3]), id: String(r[4]), oc: String(r[5]),
             cli: String(r[8]), c: String(r[9]), d: String(r[10]), qp: Number(r[11]) || 0, qf: Number(r[12]) || 0,
             up: Number(r[13]) || 0, uf: Number(r[14]) || 0, tf: Number(r[15]) || 0, alt: String(r[16]) === 'SIM' };
  }).reverse();
}
function _txtDh(v) { return Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, FUSO, 'dd/MM/yyyy HH:mm') : String(v || ''); }

/** Lançamento do faturista. L = {id, oc, cli, nf, dnf (dd/MM/aaaa), itens:[{c,d,qp,qf,up,uf}]} */
function faturar(tk, L) {
  var s = _sessao(tk);
  if (s.t.indexOf('flan') < 0) return { ok: false, msg: 'Seu usuário não tem a tela Faturamento.' };
  var nf = String(L.nf || '').replace(/\D/g, '');
  if (!nf) return { ok: false, msg: 'Informe o número da nota fiscal.' };
  if (!/^\d{2}\/\d{2}\/\d{4}$/.test(String(L.dnf || ''))) return { ok: false, msg: 'Informe a data da nota.' };
  var its = (L.itens || []).filter(function (x) { return Number(x.qf) > 0; });
  if (!its.length) return { ok: false, msg: 'Selecione ao menos um item com quantidade.' };
  var lock = LockService.getScriptLock(); lock.tryLock(20000);
  try {
    var sh = _abaFat(), agora = _agora(), os = pvOs_(L.oc), alt = 0, tot = 0;
    var linhas = its.map(function (x) {
      var qf = Number(x.qf), uf = Number(x.uf), up = Number(x.up) || 0, qp = Number(x.qp) || 0;
      var a = Math.abs(uf - up) > 0.005 || Math.abs(qf - qp) > 0.0001;
      if (a) alt++; tot += qf * uf;
      return [agora, s.n, nf, L.dnf, L.id, L.oc, os, pvRev_(String(L.oc).slice(-2)), L.cli, x.c, x.d, qp, qf, up, uf,
              Math.round(qf * uf * 100) / 100, a ? 'SIM' : 'NAO'];
    });
    sh.getRange(sh.getLastRow() + 1, 1, linhas.length, COLS_FAT.length).setValues(linhas);
    _log(s, 'FATURAMENTO', 'NF ' + nf + ' · ' + L.oc + ' ' + L.cli + ' · ' + its.length + ' item(ns) · R$ ' + (Math.round(tot * 100) / 100) + (alt ? ' · ' + alt + ' alterado(s)' : ''));
    return { ok: true, msg: 'NF ' + nf + ' registrada: ' + its.length + ' item(ns)' + (alt ? ' · ' + alt + ' com valor/quantidade diferente do pedido' : '') + '.' };
  } finally { lock.releaseLock(); }
}

/* ====================== E-MAILS DIÁRIOS (V3.0) ====================== */

function ativarV3() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACESSOS);
  if (sh && sh.getLastRow() > 1) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    sh.getRange(2, 9, v.length, 1).setValues(v.map(function (r) {
      var a = String(r[8] || '').toUpperCase();
      if (/PEDIDOS/.test(a)) { if (!/FATURAMENTO/.test(a)) a += ', FATURAMENTO'; if (!/REVISOES/.test(a)) a += ', REVISOES'; }
      if (/TABELAS/.test(a) && !/REVISOES/.test(a)) a += ', REVISOES';
      return [a.replace(/^,\s*/, '')];
    }));
  }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (/^(emailFaturamentoDia|emailRevisoesDia)$/.test(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('emailFaturamentoDia').timeBased().everyDays(1).atHour(18).create();
  ScriptApp.newTrigger('emailRevisoesDia').timeBased().everyDays(1).atHour(18).create();
  _abaFat();
  Logger.log('V3.0 pronta: aba Faturamento criada; e-mails diários de faturamento e de revisões às 18h.');
  Logger.log('Coluna avisos: FATURAMENTO e REVISOES foram incluídos para quem recebia PEDIDOS/TABELAS.');
}

function emailFaturamentoDia() {
  var hoje = Utilities.formatDate(new Date(), FUSO, 'dd/MM/yyyy');
  var L = _lancamentos().filter(function (x) { return x.dh.indexOf(hoje) === 0; });
  if (!L.length) return;
  var ped = {}, tot = 0, nfs = {};
  L.forEach(function (x) {
    var p = ped[x.oc] = ped[x.oc] || { oc: x.oc, cli: x.cli, it: [], v: 0 };
    p.it.push(x); p.v += x.tf; tot += x.tf; nfs[x.nf] = 1;
  });
  var corpo = '<p style="margin:0 0 10px"><b>' + Object.keys(ped).length + '</b> pedido(s) · <b>' + Object.keys(nfs).length + '</b> nota(s) · valor Dynamics <b>' + _moeda(tot) +
    '</b> · valor cheio <b>' + _moeda(tot * FATOR_CHEIO) + '</b></p>';
  Object.keys(ped).forEach(function (k) {
    var p = ped[k];
    corpo += '<h3 style="margin:16px 0 6px;font-size:15px">' + p.oc + ' · ' + p.cli + ' · ' + _moeda(p.v) + '</h3>' +
      _tab(['NF', 'Código', 'Descrição', 'Qtde', 'Unit. pedido', 'Unit. faturado', 'Total', 'Lançado por'], p.it.map(function (x) {
        return [x.nf, x.c, x.d, x.qf + (x.qf !== x.qp ? ' de ' + x.qp : ''), _moeda(x.up), (x.alt ? '<b style="color:#D0121C">' : '') + _moeda(x.uf) + (x.alt ? '</b>' : ''), _moeda(x.tf), x.u];
      }));
  });
  _email(_destinos('FATURAMENTO'), 'Faturamento do dia ' + hoje + ' · ' + _moeda(tot), 'Faturamento do dia ' + hoje, corpo);
}

function emailRevisoesDia() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_REVISOES);
  if (!sh || sh.getLastRow() < 2) return;
  var hoje = Utilities.formatDate(new Date(), FUSO, 'dd/MM/yyyy'), g = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, COLS_REVISOES.length).getValues().forEach(function (r) {
    var dh = _txtDh(r[0]); if (dh.indexOf(hoje) !== 0) return;
    var k = String(r[1]).trim(); (g[k] = g[k] || { oc: k, de: pvRev_(r[2]), para: pvRev_(r[3]), l: [] }).l.push(r);
  });
  var ks = Object.keys(g); if (!ks.length) return;
  var ped = {};
  var shP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PEDIDOS), P = _idx(COLS_PEDIDOS);
  shP.getRange(2, 1, shP.getLastRow() - 1, COLS_PEDIDOS.length).getValues().forEach(function (r) {
    if (String(r[P.vigente]).trim() === 'SIM') ped[String(r[P.oc]).trim()] = r;
  });
  var corpo = '<p style="margin:0 0 10px"><b>' + ks.length + '</b> pedido(s) revisado(s) hoje.</p>';
  ks.forEach(function (k) {
    var x = g[k], r = ped[k], sem = x.l.some(function (l) { return /SEM TROCA/.test(l[5]); });
    corpo += '<h3 style="margin:16px 0 4px;font-size:15px">' + k + (r ? ' · ' + (r[P.cliente] || '') + ' · ' + (r[P.revenda] || '') + ' · ' + _moeda(r[P.total_cheio]) : '') + '</h3>' +
      '<p style="margin:0 0 6px;font-size:13px">' + (sem ? '<b style="color:#D0121C">ALTERADO SEM TROCA DE REVISÃO</b>' : 'Revisão ' + x.de + ' → ' + x.para) +
      (r ? ' · entrega ' + _dataBr(r[P.data_entrega]) : '') + '</p>' +
      _tab(['Tipo', 'Código', 'Descrição', 'De', 'Para'], x.l.filter(function (l) { return !/^ALERTA$/.test(l[4]); }).map(function (l) {
        return [String(l[5]).toLowerCase(), l[6], l[7], _txtRev(l[8]), '<b>' + _txtRev(l[9]) + '</b>'];
      }));
  });
  _email(_destinos('REVISOES'), 'Pedidos revisados hoje (' + hoje + ') · ' + ks.length, 'Pedidos revisados hoje', corpo);
}

/** V2.8 — a tela não aceita Date vindo do servidor (trava o carregamento):
    data vira texto dd/MM/aa, número continua número. */
function _txtRev(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    // V3.4 — número que a planilha exibiu como data (célula com formato de data): volta a ser quantidade
    if (v.getFullYear() < 1950) return Math.round((v.getTime() - new Date(1899, 11, 30).getTime()) / 864e5 * 10000) / 10000;
    return Utilities.formatDate(v, FUSO, 'dd/MM/yy');
  }
  if (typeof v === 'string' && /^-?\d+(,\d+)?$/.test(v.trim())) return Number(v.trim().replace(',', '.'));
  return typeof v === 'number' ? v : String(v);
}

/** V2.5 — última revisão de cada OS: o que mudou (aba Revisoes). */
function _revisoes() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_REVISOES), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, COLS_REVISOES.length).getValues().forEach(function (r) {
    var oc = String(r[1]).trim(), para = pvRev_(r[3]);
    if (!oc) return;
    var k = oc + '|' + para, dh = r[0];
    if (Object.prototype.toString.call(dh) === '[object Date]') dh = Utilities.formatDate(dh, FUSO, 'dd/MM/yyyy HH:mm');
    dh = String(dh);
    if (m[k] && m[k].dh !== dh && _dhNum(dh) > _dhNum(m[k].dh)) delete m[k];     // fica o lote mais recente
    if (m[k] && m[k].dh !== dh) return;
    var x = m[k] = m[k] || { de: pvRev_(r[2]), para: para, dh: dh, dias: _diasDesde(dh.substring(0, 10)), ch: [] };
    x.ch.push({ t: String(r[5] || ''), c: String(r[6] || '').trim(), d: String(r[7] || ''),
                de: _txtRev(r[8]), pa: _txtRev(r[9]) });
  });
  return m;
}

/** V2.5 — devolve a tabela vigente do cliente (arquivo da pasta de processadas). */
function baixarTabela(tk, cod) {
  var s = _sessao(tk);
  var t = _auditoria().tabelas.filter(function (x) { return x.cod === cod; })[0];
  if (!t || !t.arq) return { ok: false, msg: 'Tabela não encontrada.' };
  var f = null;
  try {
    var it = DriveApp.getFolderById(lerCfg_('pasta_tabelas_processadas_id', '')).getFilesByName(t.arq);
    while (it.hasNext()) { var g = it.next(); if (!f || g.getDateCreated() > f.getDateCreated()) f = g; }
  } catch (e) {}
  if (!f) return { ok: false, msg: 'Arquivo ' + t.arq + ' não está na pasta 02 Tabelas Processadas.' };
  var b = f.getBlob();
  _log(s, 'DOWNLOAD TABELA', t.cod + ' ' + t.cli + ' · ' + t.arq);
  return { ok: true, nome: t.arq, mime: b.getContentType(), b64: Utilities.base64Encode(b.getBytes()) };
}

function _auditoria() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PRECOS);
  var porCli = {}, itens = {};
  if (sh && sh.getLastRow() > 1) {
    var T = _idx(COLS_PRECOS);
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_PRECOS.length).getValues();
    var cont = {};
    v.forEach(function (r) {
      if (String(r[T.vigente]).trim() !== 'SIM') return;
      var c = pvPad_(r[T.cli_dyn], 3), cod = String(r[T.codigo]).trim();
      cont[c + '|' + cod] = (cont[c + '|' + cod] || 0) + 1;
    });
    v.forEach(function (r) {
      if (String(r[T.vigente]).trim() !== 'SIM') return;
      var c = pvPad_(r[T.cli_dyn], 3);
      var t = porCli[c] = porCli[c] || { cod: c, cli: String(r[T.cliente] || '').trim() || c, it: 0, ok: 0,
                                         sd: 0, dup: 0, arq: String(r[T.arquivo] || '').trim(),
                                         carga: String(r[T.versao_carga] || '').trim(), dMax: null, dMin: null };
      var cod = String(r[T.codigo]).trim();
      var dias = _diasDesde(r[T.data_revisao]);
      var dup = cont[c + '|' + cod] > 1;
      t.it++;
      if (dias === null) t.sd++; else {
        if (t.dMin === null || dias < t.dMin) t.dMin = dias;
        if (t.dMax === null || dias > t.dMax) t.dMax = dias;
      }
      if (dup) t.dup++;
      if (!dup && dias !== null) t.ok++;
      (itens[c] = itens[c] || []).push({
        c: cod, d: String(r[T.descricao] || '').trim(), un: String(r[T.unidade] || '').trim(),
        p: Number(r[T.preco]) || 0, dr: _dataBr(r[T.data_revisao]), dias: dias, dup: dup,
        st: String(r[T.status] || '').trim(), aba: String(r[T.aba] || '').trim(), ln: r[T.linha_origem]
      });
    });
  }
  var tabelas = Object.keys(porCli).map(function (k) { return porCli[k]; })
    .sort(function (a, b) { return a.cli < b.cli ? -1 : 1; });
  return { tabelas: tabelas, itens: itens };
}

function _ultimosLogs(verOculto) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = Math.min(600, sh.getLastRow() - 1);
  var v = sh.getRange(sh.getLastRow() - n + 1, 1, n, COLS_LOG.length).getValues();
  return v.reverse().filter(function (r) {
    return verOculto || String(r[2]).trim() !== 'oculto';
  }).map(function (r) {
    var dh = r[0];
    if (Object.prototype.toString.call(dh) === '[object Date]') dh = Utilities.formatDate(dh, FUSO, 'dd/MM/yyyy HH:mm:ss');
    return { dh: String(dh), u: String(r[1]), p: String(r[2]), a: String(r[3]), d: String(r[4]) };
  }).slice(0, 400);
}

/* datas: o Sheets transforma "14/09/26" em Date. Sempre devolve dd/MM/aa. */
function _dataBr(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, FUSO, 'dd/MM/yy');
  var s = String(v).trim();
  var m = /^(\d{2})\/(\d{2})\/(\d{2,4})$/.exec(s);
  if (m) return m[1] + '/' + m[2] + '/' + m[3].slice(-2);
  var d = new Date(s);
  return isNaN(d.getTime()) ? s : Utilities.formatDate(d, FUSO, 'dd/MM/yy');
}

function _mapaDyn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var m = { cod: {}, nome: {} };
  var sh = ss.getSheetByName(ABA_CLIENTES);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
      var dyn = pvPad_(r[0], 3), pai = pvPad_(r[1], 3), nm = String(r[2] || '').trim().toUpperCase();
      if (!/^\d{3}$/.test(dyn)) return;
      if (pai) m.cod[pai] = dyn;
      m.cod[dyn] = dyn;
      if (nm) m.nome[nm] = dyn;
    });
  }
  var shL = ss.getSheetByName(ABA_LOGOS);
  if (shL && shL.getLastRow() > 1) {
    shL.getRange(2, 1, shL.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var dyn = pvPad_(r[0], 3), nm = String(r[1] || '').trim().toUpperCase();
      if (/^\d{3}$/.test(dyn) && nm && !m.nome[nm]) m.nome[nm] = dyn;
    });
  }
  return m;
}
function _dynPorNome(mapa, nome) {
  var n = String(nome || '').trim().toUpperCase();
  if (!n) return '';
  if (mapa.nome[n]) return mapa.nome[n];
  var ks = Object.keys(mapa.nome).sort(function (a, b) { return b.length - a.length; });
  for (var i = 0; i < ks.length; i++) {
    if (n.indexOf(ks[i]) === 0 || ks[i].indexOf(n) === 0) return mapa.nome[ks[i]];
  }
  return '';
}

function _logos() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_LOGOS), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
    var cod = pvPad_(r[0], 3);
    if (cod && r[3]) m[cod] = 'data:' + (r[2] || 'image/png') + ';base64,' + r[3];
  });
  return m;
}
function _logoNomes() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_LOGOS), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    var cod = pvPad_(r[0], 3); if (cod) m[cod] = String(r[1] || '').trim();
  });
  return m;
}

/* ====================== ENVIO DE ARQUIVOS ====================== */

function pastaApp_(pai, nome) {
  var it = pai ? pai.getFoldersByName(nome) : DriveApp.getFoldersByName(nome);
  if (it.hasNext()) return it.next();
  return pai ? pai.createFolder(nome) : DriveApp.createFolder(nome);
}

/** tipo: 'pedido' (PDF) | 'tabela' (Excel) | 'logo' (PNG/JPG) */
function enviarArquivo(tk, tipo, nome, mime, dados, extra) {
  var s = _sessao(tk);
  if (s.t.indexOf('up') < 0) return { ok: false, msg: 'Seu usuário não tem a tela Upload.' };
  try {
    var bytes = Utilities.base64Decode(dados);
    var blob  = Utilities.newBlob(bytes, mime, nome);
    var raiz  = pastaApp_(null, PASTA_RAIZ);
    var kb = Math.round(bytes.length / 1024) + ' KB';
    if (tipo === 'pedido') {
      if (!/\.pdf$/i.test(nome)) return { ok: false, msg: nome + ': não é PDF.' };
      pastaApp_(raiz, P_UP_PEDIDOS).createFile(blob);
      _log(s, 'UPLOAD PEDIDO', nome + ' · ' + kb);
      return { ok: true, msg: nome + ' na fila de pedidos.' };
    }
    if (tipo === 'tabela') {
      if (!/\.(xlsx|xlsm|xls)$/i.test(nome)) return { ok: false, msg: nome + ': não é Excel.' };
      var ext = (/\.(xlsx|xlsm|xls)$/i.exec(nome) || ['', '.xlsx'])[0];
      var base = nome.replace(/\.(xlsx|xlsm|xls)$/i, '').replace(/\s*\(\d+\)$/, '').replace(/_\d{4}-\d{2}-\d{2}_\d{4}$/, '').trim();
      nome = base + '_' + Utilities.formatDate(new Date(), FUSO, 'yyyy-MM-dd_HHmm') + ext;
      blob.setName(nome);
      pastaApp_(raiz, P_UP_TABELAS).createFile(blob);
      _log(s, 'UPLOAD TABELA', nome + ' · ' + kb + (extra && extra.cli ? ' · cliente ' + extra.cli : ''));
      return { ok: true, msg: nome + ' na fila de tabelas.' };
    }
    if (tipo === 'logo') {
      var cod = pvPad_((extra && extra.cod) || '', 3);
      var cli = String((extra && extra.cli) || '').trim().toUpperCase();
      if (!/^\d{3}$/.test(cod)) return { ok: false, msg: 'Informe o código Dyn de três dígitos.' };
      if (!cli) return { ok: false, msg: 'Informe o nome do cliente.' };
      if (bytes.length > LIMITE_LOGO) return { ok: false, msg: 'Logo com ' + kb + '. Limite ' + Math.round(LIMITE_LOGO / 1024) + ' KB.' };
      blob.setName(cod + '_' + cli.replace(/\s+/g, '_') + (/jpe?g/i.test(mime) ? '.jpg' : '.png'));
      pastaApp_(raiz, P_LOGOS).createFile(blob);
      _gravarLogo(cod, cli, mime, Utilities.base64Encode(bytes));
      _log(s, 'UPLOAD LOGO', cod + ' ' + cli);
      return { ok: true, msg: 'Logo de ' + cli + ' salva.' };
    }
    return { ok: false, msg: 'Tipo de arquivo não reconhecido.' };
  } catch (e) {
    _log(s, 'UPLOAD FALHOU', nome + ' · ' + e);
    return { ok: false, msg: 'Falha ao enviar ' + nome + ': ' + e };
  }
}

function _gravarLogo(cod, cli, mime, b64) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_LOGOS);
  if (!sh) return;
  var agora = _agora();
  if (sh.getLastRow() > 1) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < v.length; i++) {
      if (pvPad_(v[i][0], 3) === cod) { sh.getRange(i + 2, 1, 1, 5).setValues([[cod, cli, mime, b64, agora]]); return; }
    }
  }
  sh.appendRow([cod, cli, mime, b64, agora]);
}

/* ====================== AVISOS POR E-MAIL (V2.3) ====================== */

function ativarAvisosV2_3() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA_ACESSOS);
  if (!sh) { Logger.log('Aba Acessos não existe. Rode instalarV2 primeiro.'); return; }
  sh.getRange(1, 8, 1, 2).setValues([['email', 'avisos']]).setFontWeight('bold');
  var n = sh.getLastRow() - 1;
  if (n > 0) {
    var v = sh.getRange(2, 1, n, 9).getValues();
    var av = v.map(function (r) {
      if (String(r[8]).trim()) return [r[8]];
      var t = _telas(r[3]), u = String(r[0]).toUpperCase();
      if (u === 'SUPORTE') return [''];
      var a = [];
      if (t.indexOf('dia') >= 0 || t.indexOf('mes') >= 0) a.push('PEDIDOS');
      if (t.indexOf('aud') >= 0 && t.indexOf('dia') < 0) a.push('TABELAS');
      return [a.join(', ')];
    });
    sh.getRange(2, 9, n, 1).setValues(av);
  }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'verificarAvisos' || f === 'avisoTabelasSemanal') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('verificarAvisos').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('avisoTabelasSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  var shP = ss.getSheetByName(ABA_PEDIDOS);
  PropertiesService.getScriptProperties().setProperty('av_linha', String(shP ? shP.getLastRow() : 1));
  Logger.log('Avisos ligados. Preencha a coluna email da aba Acessos. Coluna avisos: PEDIDOS e/ou TABELAS.');
}

function _destinos(tipo) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_ACESSOS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().filter(function (r) {
    var at = String(r[4]).trim().toUpperCase();
    return at !== 'NAO' && at !== 'NÃO' && /@/.test(String(r[7])) &&
           (String(r[8]).toUpperCase().indexOf(tipo) >= 0 || /TODOS/i.test(String(r[8])));
  }).map(function (r) { return String(r[7]).trim(); });
}

function _moeda(n) {
  var s = (Math.round((Number(n) || 0) * 100) / 100).toFixed(2).split('.');
  return 'R$ ' + s[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + s[1];
}
function _email(para, assunto, titulo, corpo) {
  if (!para.length) return 0;
  var html = '<div style="font-family:Arial,sans-serif;max-width:760px;color:#15181D">' +
    '<div style="background:#15181D;padding:14px 18px;border-bottom:3px solid #D0121C">' +
    '<span style="color:#D0121C;font-weight:800;letter-spacing:.08em">DYNAMICS</span>' +
    '<span style="color:#E6E9ED;font-size:12px;margin-left:10px">PEDIDOS DE VENDA</span></div>' +
    '<div style="padding:16px 18px"><h2 style="margin:0 0 12px;font-size:19px">' + titulo + '</h2>' + corpo +
    '<p style="margin:18px 0"><a href="' + LINK_APP + '" style="background:#D0121C;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none;font-weight:700">Abrir o app</a></p>' +
    '<p style="font-size:11px;color:#8C96A4">Aviso automático do Fluxo Venda Dynamics · uso interno e confidencial · Dynamics Metalurgica Ltda.</p></div></div>';
  MailApp.sendEmail({ to: para.join(','), subject: assunto, htmlBody: html, name: 'Pedidos de Venda Dynamics' });
  return para.length;
}
function _tab(cab, linhas) {
  var th = cab.map(function (c) { return '<th style="text-align:left;font-size:11px;color:#566170;padding:6px 8px;border-bottom:2px solid #E1E5EA">' + c + '</th>'; }).join('');
  var tr = linhas.map(function (l) {
    return '<tr>' + l.map(function (c) { return '<td style="padding:6px 8px;border-bottom:1px solid #EEF1F4;font-size:13px">' + c + '</td>'; }).join('') + '</tr>';
  }).join('');
  return '<table style="border-collapse:collapse;width:100%">' + '<tr>' + th + '</tr>' + tr + '</table>';
}

/** Gatilho a cada 10 min: avisa pedidos novos (Diretoria) e itens fora da tabela (Orçamento). */
function verificarAvisos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shP = ss.getSheetByName(ABA_PEDIDOS), shI = ss.getSheetByName(ABA_ITENS);
  if (!shP) return;
  var pr = PropertiesService.getScriptProperties();
  var ult = Number(pr.getProperty('av_linha') || 0), fim = shP.getLastRow();
  if (!ult) { pr.setProperty('av_linha', String(fim)); return; }
  if (fim <= ult) return;
  var P = _idx(COLS_PEDIDOS), I = _idx(COLS_ITENS);
  var novos = shP.getRange(ult + 1, 1, fim - ult, COLS_PEDIDOS.length).getValues()
    .filter(function (r) { return String(r[P.vigente]).trim() === 'SIM'; });
  pr.setProperty('av_linha', String(fim));
  if (!novos.length) return;

  var ids = {};
  novos.forEach(function (r) { ids[String(r[P.id_pedido]).trim() + '|' + pvRev_(r[P.revisao])] = r; });
  var pend = [];
  if (shI && shI.getLastRow() > 1) {
    shI.getRange(2, 1, shI.getLastRow() - 1, COLS_ITENS.length).getValues().forEach(function (x) {
      var r = ids[String(x[I.id_pedido]).trim() + '|' + pvRev_(x[I.revisao])];
      if (!r) return;
      var tb = Number(x[I.preco_tabela]) || 0, u = Number(x[I.preco_cheio]) || 0;
      var st = String(x[I.status_preco] || '').trim();
      if (tb > 0 && Math.abs(u - tb) <= 0.005 && st === 'OK') return;
      pend.push([String(r[P.cliente] || r[P.cli_pailon]), String(x[I.codigo]), String(x[I.descricao]),
                 String(r[P.oc] || 'ERP ' + r[P.numero_erp]), _moeda(u), tb > 0 ? _moeda(tb) : '—',
                 tb > 0 ? ((u > tb ? '+' : '') + ((u - tb) / tb * 100).toFixed(1).replace('.', ',') + '%') : '—',
                 tb > 0 ? (st === 'OK' ? 'preço diferente' : st.toLowerCase()) : st.toLowerCase() || 'sem preço em tabela']);
    });
  }

  var tot = 0, linhas = novos.map(function (r) {
    tot += Number(r[P.total_cheio]) || 0;
    var d = r[P.desvio_pct];
    return ['<b>' + (String(r[P.oc]).trim() || 'ERP ' + r[P.numero_erp]) + '</b>', String(r[P.cliente] || r[P.cli_pailon]),
            String(r[P.revenda] || ''), _dataBr(r[P.data_pedido]), _dataBr(r[P.data_entrega]),
            _moeda(r[P.total_cheio]), d === '' ? 'sem tabela' : (d > 0 ? '+' : '') + String(d).replace('.', ',') + '%'];
  });
  var n1 = _email(_destinos('PEDIDOS'),
    novos.length + ' novo(s) pedido(s) de venda · ' + _moeda(tot),
    novos.length + ' novo(s) pedido(s) de venda',
    '<p style="margin:0 0 10px">Valor cheio total: <b>' + _moeda(tot) + '</b></p>' +
    _tab(['OS', 'Cliente', 'Revenda', 'Pedido', 'Entrega', 'Valor cheio', 'Contra tabela'], linhas));

  var n2 = 0;
  if (pend.length) {
    pend.sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
    n2 = _email(_destinos('TABELAS'),
      'Tabela para atualizar: ' + pend.length + ' item(ns) em pedidos novos',
      'Itens vendidos fora da tabela ou sem preço',
      '<p style="margin:0 0 10px">Confira na tabela do cliente, corrija e envie a tabela atualizada pelo app (Upload).</p>' +
      _tab(['Cliente', 'Código', 'Descrição', 'OS', 'Vendido', 'Tabela', 'Dif.', 'Situação'], pend.slice(0, 200)) +
      (pend.length > 200 ? '<p>… e mais ' + (pend.length - 200) + ' item(ns). Veja todos na Auditoria.</p>' : ''));
  }
  _log({ u: 'SISTEMA', p: 'sistema' }, 'AVISOS ENVIADOS', novos.length + ' pedido(s) → ' + n1 + ' e-mail(s) · ' + pend.length + ' item(ns) → ' + n2 + ' e-mail(s)');
}

/** Gatilho semanal (segunda 7h): tabelas vencidas, sem data ou com código duplicado. */
function avisoTabelasSemanal() {
  var a = _auditoria();
  var l = a.tabelas.filter(function (t) { return t.sd || t.dup || (t.dMax !== null && t.dMax > DIAS_TABELA_VELHA); })
    .sort(function (x, y) { return (y.dMax || 0) - (x.dMax || 0); })
    .map(function (t) { return ['<b>' + t.cli + '</b> (' + t.cod + ')', t.it, t.dMax === null ? '—' : t.dMax + ' dias', t.sd, t.dup, t.arq]; });
  if (!l.length) return;
  _email(_destinos('TABELAS'), 'Tabelas para revisar: ' + l.length + ' cliente(s)', 'Tabelas de preço para revisar',
    '<p style="margin:0 0 10px">Clientes com tabela há mais de ' + DIAS_TABELA_VELHA + ' dias sem revisão, itens sem data ou código duplicado.</p>' +
    _tab(['Cliente', 'Itens', 'Revisão mais antiga', 'Sem data', 'Duplicados', 'Arquivo'], l));
}

function testarAvisos() {
  ['PEDIDOS', 'TABELAS'].forEach(function (t) {
    var d = _destinos(t);
    Logger.log(t + ': ' + (d.join(', ') || 'ninguém — preencha email e avisos na aba Acessos'));
    _email(d, 'Teste de aviso · ' + t, 'Teste de aviso (' + t + ')', '<p>Se você recebeu este e-mail, os avisos de ' + t + ' estão funcionando.</p>');
  });
}

/* ====================== DIAGNÓSTICO ====================== */

function situacaoApp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  function conta(n) { var s = ss.getSheetByName(n); return s ? Math.max(0, s.getLastRow() - 1) : -1; }
  [ABA_PEDIDOS, ABA_ITENS, ABA_PRECOS, ABA_LOGOS, ABA_ACESSOS, ABA_LOG].forEach(function (n) {
    Logger.log(n + ': ' + conta(n) + ' linha(s)');
  });
  var tk = Utilities.getUuid();
  CacheService.getScriptCache().put('fvd_' + tk, JSON.stringify({ u: 'EDITOR', p: 'oculto', n: 'Editor', t: TODAS_TELAS }), 60);
  var t = Date.now(), d = carregar(tk);
  Logger.log(d.P.length + ' pedidos vigentes · ' + d.tabelas.length + ' tabelas · carga ' + (Date.now() - t) + ' ms');
}
