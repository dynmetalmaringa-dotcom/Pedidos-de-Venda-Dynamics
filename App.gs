/**********************************************************************
 * FLUXO VENDA DYNAMICS — APLICATIVO WEB
 * Arquivo: App              Versão: V4.4
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
  'qtde_pedido','qtde_faturada','unit_pedido_40','unit_faturado_40','total_faturado_40','alterado','data_expedicao'];

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
    var DYN = _mapaDyn(), REV = _revisoes(), FAT = _faturadoPorItem(), STT = _statusPorOs(), HDT = _alteracoesData();
    vp.forEach(function (r) {
      if (String(r[P.vigente]).trim() !== 'SIM') return;
      var id = String(r[P.id_pedido]).trim(), rev = pvRev_(r[P.revisao]);
      var itens = porPed[id + '|' + rev] || [];
      var v = Number(r[P.total_cheio]) || 0;
      var v40 = Number(r[P.total_40]) || v / FATOR_CHEIO;
      var nome = String(r[P.cliente] || '').trim() || String(r[P.cli_pailon]).trim();
      var p = {
        id: id, oc: String(r[P.oc]).trim(), rev: rev, erp: String(r[P.numero_erp]).trim(),
        cli: nome, cp: String(r[P.cli_pailon] || '').trim() ? pvPad_(r[P.cli_pailon], 3) : '', rv: String(r[P.revenda] || '').trim(),
        loc: [String(r[P.cidade] || '').trim(), String(r[P.uf] || '').trim()].filter(String).join(' - '),
        dp: _dataBr(r[P.data_pedido]), ent: _dataBr(r[P.data_entrega]),
        v: _r2(v), v40: _r2(v40), itens: itens, fx: (v40 > 0 && Math.abs(v - v40) < 0.01) ? 1 : FATOR_CHEIO,
        dc: (function (x) { return Object.prototype.toString.call(x) === '[object Date]' ? Utilities.formatDate(x, FUSO, 'dd/MM/yyyy HH:mm') : String(x || '').trim(); })(r[P.data_carga]),
        rt: String(r[P.tipo] || '').trim() === 'RETRABALHO', org: String(r[P.os_origem] || '').trim()
      };
      /* faturamento: por OS (vale para qualquer revisão) + código */
      var os = pvOs_(p.oc), fv = 0, comp1 = true, algum = false;
      itens.forEach(function (x) {
        var f = FAT[os + '|' + x.c]; x.fq = f ? _r2(f.q) : 0; x.fv = f ? _r2(f.v) : 0; x.fd = f ? f.d.join(', ') : '';
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
      var stt = STT[os]; if (stt && stt.s !== 'LIBERADO') p.st = stt;      // V3.7 — cancelado / paralisado
      if (HDT[id]) p.da = HDT[id];                                        // V4.3 — data de entrega alterada pelo relatório do ForWood
      out.P.push(p);
    });
  }

  if (s.t.indexOf('aud') >= 0 || s.t.indexOf('up') >= 0) {
    var a = _auditoria();
    out.tabelas = a.tabelas; out.tabItens = a.itens;
    out.sitTab = _situacaoTabelas();
  }
  if (s.t.indexOf('up') >= 0) out.logs = _ultimosLogs(s.p === 'oculto');
  if (s.t.indexOf('up') >= 0) out.hd = _ultimasDatas();                  // V4.3
  if (s.t.indexOf('reg') >= 0 && !out.logs.length) out.logs = _ultimosLogs(s.p === 'oculto');
  if (s.t.indexOf('flan') >= 0) out.fat = _lancamentos();
  out.motivos = _motivos();
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
  if (String(sh.getRange(1, COLS_FAT.length).getValue()).trim() !== 'data_expedicao') sh.getRange(1, COLS_FAT.length).setValue('data_expedicao').setFontWeight('bold');   // V4.1
  return sh;
}
function _faturadoPorItem() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_FAT), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, COLS_FAT.length).getValues().forEach(function (r) {
    var k = String(r[6]).trim() + '|' + String(r[9]).trim();
    m[k] = m[k] || { q: 0, v: 0, d: [] }; m[k].q += Number(r[12]) || 0; m[k].v += Number(r[15]) || 0;
    var dn = _dataBr(r[3]); if (dn && m[k].d.indexOf(dn) < 0) m[k].d.push(dn);          // V4.2 — datas de faturamento
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
             up: Number(r[13]) || 0, uf: Number(r[14]) || 0, tf: Number(r[15]) || 0, alt: String(r[16]) === 'SIM', dex: _dataBr(r[17]) };
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
              Math.round(qf * uf * 100) / 100, a ? 'SIM' : 'NAO', _diaUtilAnterior(L.dnf)];
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
  ativarV4();   // V4.0 — gatilhos de e-mail ficam em ativarV4
  _abaFat();
  Logger.log('V3.0 pronta: aba Faturamento criada; e-mails diários de faturamento e de revisões às 18h.');
  Logger.log('Coluna avisos: FATURAMENTO e REVISOES foram incluídos para quem recebia PEDIDOS/TABELAS.');
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
  var n1 = 0;   // V4.0 — pedidos novos só no fechamento do dia (emailPedidosDia, 18h)

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


/* ===================== V3.7 — CANCELADO / PARALISADO ===================== */
var ABA_STATUS = 'StatusPedido', ABA_MOTIVOS = 'Motivos';
var COLS_STATUS = ['data_hora', 'usuario', 'os', 'id_pedido', 'status', 'motivo', 'observacao', 'solicitante'];
var MOTIVOS_PADRAO = [
  ['CANCELADO', 'Cliente desistiu'], ['CANCELADO', 'Pedido duplicado'], ['CANCELADO', 'Substituído por outro pedido'],
  ['CANCELADO', 'Obra / projeto cancelado pelo cliente'], ['CANCELADO', 'Erro de lançamento'], ['CANCELADO', 'Outro'],
  ['PARALISADO', 'Aguardando cliente'], ['PARALISADO', 'Aguardando projeto'], ['PARALISADO', 'Aguardando medição'],
  ['PARALISADO', 'Aguardando aprovação comercial'], ['PARALISADO', 'Pendência financeira'], ['PARALISADO', 'Obra paralisada'], ['PARALISADO', 'Outro']];
function _abaStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(ABA_STATUS);
  if (!sh) { sh = ss.insertSheet(ABA_STATUS); sh.setFrozenRows(1); }
  if (String(sh.getRange(1, COLS_STATUS.length).getValue()).trim() !== COLS_STATUS[COLS_STATUS.length - 1]) sh.getRange(1, 1, 1, COLS_STATUS.length).setValues([COLS_STATUS]).setFontWeight('bold');
  return sh;
}
function _motivos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(ABA_MOTIVOS);
  if (!sh) {
    sh = ss.insertSheet(ABA_MOTIVOS);
    sh.getRange(1, 1, 1, 2).setValues([['tipo', 'motivo']]).setFontWeight('bold');
    sh.getRange(2, 1, MOTIVOS_PADRAO.length, 2).setValues(MOTIVOS_PADRAO); sh.setFrozenRows(1);
  }
  var m = { CANCELADO: [], PARALISADO: [] };
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    var t = String(r[0]).trim().toUpperCase(), x = String(r[1]).trim();
    if (m[t] && x && m[t].indexOf(x) < 0) m[t].push(x);
  });
  return m;
}
function _statusPorOs() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_STATUS), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, COLS_STATUS.length).getValues().forEach(function (r) {
    var os = String(r[2]).trim(); if (!os) return;
    m[os] = { s: String(r[4]).trim(), m: String(r[5]).trim(), o: String(r[6]).trim(), u: String(r[1]).trim(), dh: _txtDh(r[0]), sol: String(r[7] || '').trim() };
  });
  return m;
}
/** Cancela, paralisa ou libera (reativa) um pedido. Fica registrado quem e quando. */
function mudarStatus(tk, L) {
  var s = _sessao(tk);
  if (!s.t.some(function (k) { return /^(dia|rdia|pcar)$/.test(k); })) return { ok: false, msg: 'Seu usuário não pode alterar a situação do pedido.' };
  var st = String(L.st || '').toUpperCase();
  if (['CANCELADO', 'PARALISADO', 'LIBERADO'].indexOf(st) < 0) return { ok: false, msg: 'Situação inválida.' };
  if (st !== 'LIBERADO' && !String(L.motivo || '').trim()) return { ok: false, msg: 'Escolha o motivo.' };
  if (!String(L.sol || '').trim()) return { ok: false, msg: 'Informe quem solicitou.' };
  var os = pvOs_(L.oc), agora = _agora();
  var sol = String(L.sol).trim().substring(0, 80);
  _abaStatus().appendRow([agora, s.n, os, L.id, st, String(L.motivo || ''), String(L.obs || '').substring(0, 300), sol]);
  _log(s, 'PEDIDO ' + st, os + ' · ' + (L.cli || '') + (L.motivo ? ' · ' + L.motivo : '') + ' · solicitado por ' + sol + (L.obs ? ' · ' + L.obs : ''));
  return { ok: true, msg: os + (st === 'LIBERADO' ? ' liberado novamente.' : ' marcado como ' + st.toLowerCase() + '.'),
           st: st === 'LIBERADO' ? null : { s: st, m: String(L.motivo || ''), o: String(L.obs || ''), u: s.n, dh: agora.substring(0, 16), sol: sol } };
}

/* ===================== V3.7 — UPLOAD DO RELATÓRIO DE FATURAMENTO (ERP) ===================== */
/** Lê a planilha do ERP (Codigo, Produto Codigo, Qtde Produto, Valor Produto Unitario,
 *  Data Atendimento...) e devolve as linhas. O app confere com os pedidos antes de gravar. */
function lerFaturamento(tk, b64, nome) {
  var s = _sessao(tk);
  if (s.t.indexOf('flan') < 0) return { ok: false, msg: 'Seu usuário não tem a tela Faturamento.' };
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), MimeType.MICROSOFT_EXCEL, nome || 'faturamento.xlsx');
  // V3.9 — funciona com o serviço Drive na versão 2 (insert) ou 3 (create)
  var f = Drive.Files.insert ? Drive.Files.insert({ title: '__fat__' + (nome || ''), mimeType: MimeType.GOOGLE_SHEETS }, blob, { convert: true })
                             : Drive.Files.create({ name: '__fat__' + (nome || ''), mimeType: MimeType.GOOGLE_SHEETS }, blob);
  try {
    var v = SpreadsheetApp.openById(f.id).getSheets()[0].getDataRange().getValues();
    var cab = v[0].map(function (x) { return String(x).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim(); });
    function col(re) { for (var i = 0; i < cab.length; i++) if (re.test(cab[i])) return i; return -1; }
    var C = { erp: col(/^codigo$/), obs: col(/observacao/), c: col(/produto codigo/), d: col(/produto descricao/), q: col(/qtde/),
              u: col(/unitario/), t: col(/produto total/), dat: col(/atendimento/), sit: col(/situacao/), nf: col(/(^| )(nf|nota)( |$)|nota fiscal/) };
    if (C.erp < 0 || C.c < 0 || C.q < 0 || C.u < 0) return { ok: false, msg: 'Planilha fora do padrão: preciso das colunas Codigo, Produto Codigo, Qtde Produto e Valor Produto Unitario.' };
    var L = [];
    v.slice(1).forEach(function (r) {
      var erp = String(r[C.erp]).trim(); if (!erp) return;
      var dt = C.dat >= 0 ? r[C.dat] : '';
      if (Object.prototype.toString.call(dt) === '[object Date]') dt = Utilities.formatDate(dt, FUSO, 'dd/MM/yyyy');
      L.push({ erp: erp, obs: C.obs >= 0 ? String(r[C.obs]).trim() : '', c: String(r[C.c]).trim(), d: C.d >= 0 ? String(r[C.d]).trim() : '',
               q: Number(String(r[C.q]).replace(',', '.')) || 0, u: Number(String(r[C.u]).replace(',', '.')) || 0,
               dat: String(dt || '').trim(), sit: C.sit >= 0 ? String(r[C.sit]).trim() : '', nf: C.nf >= 0 ? String(r[C.nf]).trim() : '' });
    });
    _log(s, 'FATURAMENTO LIDO', (nome || '') + ' · ' + L.length + ' linha(s)');
    return { ok: true, linhas: L, temNf: C.nf >= 0 };
  } finally { try { DriveApp.getFileById(f.id).setTrashed(true); } catch (e) {} }
}
/** Grava vários pedidos de uma vez (vindo do relatório do ERP). */
function faturarLote(tk, lotes) {
  var s = _sessao(tk);
  if (s.t.indexOf('flan') < 0) return { ok: false, msg: 'Seu usuário não tem a tela Faturamento.' };
  var lock = LockService.getScriptLock(); lock.tryLock(30000);
  try {
    var sh = _abaFat(), agora = _agora(), linhas = [], tot = 0;
    (lotes || []).forEach(function (L) {
      var os = pvOs_(L.oc), nf = String(L.nf || 'S/NF').trim();
      (L.itens || []).forEach(function (x) {
        var qf = Number(x.qf), uf = Number(x.uf), up = Number(x.up) || 0, qp = Number(x.qp) || 0;
        if (!(qf > 0)) return;
        var a = Math.abs(uf - up) > 0.005 || Math.abs(qf - qp) > 0.0001; tot += qf * uf;
        linhas.push([agora, s.n + ' (relatório ERP)', nf, L.dnf, L.id, L.oc, os, pvRev_(String(L.oc).slice(-2)), L.cli, x.c, x.d, qp, qf, up, uf, Math.round(qf * uf * 100) / 100, a ? 'SIM' : 'NAO', L.dex || _diaUtilAnterior(L.dnf)]);
      });
    });
    if (!linhas.length) return { ok: false, msg: 'Nenhum item para gravar.' };
    sh.getRange(sh.getLastRow() + 1, 1, linhas.length, COLS_FAT.length).setValues(linhas);
    _log(s, 'FATURAMENTO LOTE', lotes.length + ' pedido(s) · ' + linhas.length + ' item(ns) · R$ ' + (Math.round(tot * 100) / 100));
    return { ok: true, msg: lotes.length + ' pedido(s) e ' + linhas.length + ' item(ns) registrados.' };
  } finally { lock.releaseLock(); }
}

/* V3.7 — acentua os textos gravados sem acento na base (e-mails) */
function _acent(t) {
  var A = [[/incluido/g,'incluído'],[/excluido/g,'excluído'],[/revisao/g,'revisão'],[/mudanca/g,'mudança'],[/programacao/g,'programação'],[/condicao/g,'condição'],[/descricao/g,'descrição'],[/preco/g,'preço'],[/observacao/g,'observação'],[/\bja\b/g,'já'],[/\bnao\b/g,'não']];
  t = String(t || ''); A.forEach(function (a) { t = t.replace(a[0], a[1]); }); return t;
}


/* ===================== V4.0 — E-MAILS DE FECHAMENTO (layout para celular) =====================
   emailPedidosDia   → 18h · pedidos novos do dia            · coluna avisos: PEDIDOS
   emailRevisoesDia  → 18h · pedidos revisados do dia        · coluna avisos: REVISOES
   enviarEmailFaturamento → botão na tela Faturamento (sem gatilho) · coluna avisos: FATURAMENTO
   Rode ativarV4() uma vez. */
function ativarV4() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (/^(emailFaturamentoDia|emailRevisoesDia|emailPedidosDia)$/.test(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('emailPedidosDia').timeBased().everyDays(1).atHour(18).nearMinute(0).inTimezone(FUSO).create();
  ScriptApp.newTrigger('emailRevisoesDia').timeBased().everyDays(1).atHour(18).nearMinute(5).inTimezone(FUSO).create();
  Logger.log('V4.0: e-mails de pedidos novos e revisados às 18h. Faturamento sem gatilho (botão no app).');
  Logger.log('Destinatários de PEDIDOS: ' + (_destinos('PEDIDOS').join(', ') || 'ninguém'));
  Logger.log('Destinatários de REVISOES: ' + (_destinos('REVISOES').join(', ') || 'ninguém'));
  Logger.log('Destinatários de FATURAMENTO: ' + (_destinos('FATURAMENTO').join(', ') || 'ninguém'));
}

var M_F = "font-family:-apple-system,'Segoe UI',Roboto,Arial,sans-serif";
function _mChip(t, bg, fg) { return '<span style="display:inline-block;background:' + bg + ';color:' + fg + ';font-size:11px;font-weight:700;padding:3px 8px;border-radius:4px;white-space:nowrap">' + t + '</span>'; }
function _mSec(t) { return '<tr><td style="padding:18px 20px 8px;font-size:11px;color:#566070;letter-spacing:1.2px;text-transform:uppercase;font-weight:800;border-top:1px solid #EEF1F5">' + t + '</td></tr>'; }
function _mKpis(l) {
  var rows = '';
  for (var i = 0; i < l.length; i += 2) {
    rows += '<tr>' + l.slice(i, i + 2).map(function (k, j) {
      return '<td width="50%" style="padding:12px 14px;border-bottom:1px solid #2A313B;' + (j === 0 ? 'border-right:1px solid #2A313B;' : '') + 'vertical-align:top"><div style="font-size:10px;color:#9AA4B2;letter-spacing:1px;text-transform:uppercase;font-weight:700">' + k[0] +
        '</div><div style="font-size:20px;font-weight:800;color:' + (k[3] || '#fff') + ';margin-top:3px;white-space:nowrap">' + k[1] + '</div><div style="font-size:11.5px;color:#9AA4B2;margin-top:1px">' + k[2] + '</div></td>';
    }).join('') + (l.slice(i, i + 2).length === 1 ? '<td></td>' : '') + '</tr>';
  }
  return '<tr><td style="padding:14px 20px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#15181D;border-radius:10px">' + rows + '</table></td></tr>';
}
function _mEnviar(tipo, assunto, sub, titulo, lead, corpo) {
  var para = _destinos(tipo);
  if (!para.length) { Logger.log('Ninguém com ' + tipo + ' na coluna avisos.'); return 0; }
  var html = '<div style="margin:0;background:#EEF1F5;' + M_F + '"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#EEF1F5"><tr><td align="center" style="padding:16px 8px">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#ffffff;border-radius:12px;overflow:hidden;' + M_F + ';color:#15181D">' +
    '<tr><td style="background:#15181D;padding:14px 20px;border-bottom:3px solid #D0121C"><span style="color:#E5242E;font-weight:800;letter-spacing:2px;font-size:14px">DYNAMICS</span><br><span style="color:#9AA4B2;font-size:10px;letter-spacing:1.5px">FECHAMENTO DO DIA · PEDIDOS DE VENDA</span></td></tr>' +
    '<tr><td style="padding:20px 20px 4px"><div style="font-size:11px;color:#8C96A4;letter-spacing:1px;text-transform:uppercase;font-weight:700">' + sub + '</div><div style="font-size:24px;line-height:1.2;font-weight:800;margin-top:6px">' + titulo + '</div>' +
    (lead ? '<div style="font-size:14px;line-height:1.5;color:#3A4350;margin-top:8px">' + lead + '</div>' : '') + '</td></tr>' + corpo +
    '<tr><td style="padding:20px"><a href="' + LINK_APP + '" style="display:block;text-align:center;background:#D0121C;color:#fff;padding:13px;border-radius:8px;text-decoration:none;font-weight:700;font-size:15px">Abrir no app</a></td></tr>' +
    '<tr><td style="padding:14px 20px;background:#F4F6F9;font-size:11px;line-height:1.5;color:#8C96A4">Envio automático · uso interno e confidencial · Dynamics Metalurgica Ltda.<br>Valor cheio = venda ao cliente · Valor Dynamics = valor dos PDFs (40%; pedido direto = 100%).</td></tr></table></td></tr></table></div>';
  MailApp.sendEmail({ to: para.join(','), subject: assunto, htmlBody: html, name: 'Pedidos de Venda Dynamics' });
  _log({ u: 'SISTEMA', p: 'sistema' }, 'E-MAIL ' + tipo, assunto + ' → ' + para.length + ' destinatário(s)');
  return para.length;
}
function _mDiaSemana(d) { return ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'][d.getDay()]; }
function _mHoje() { var d = new Date(); return { d: d, br: Utilities.formatDate(d, FUSO, 'dd/MM/yyyy'), sub: _mDiaSemana(d) + ' · ' + Utilities.formatDate(d, FUSO, 'dd/MM/yyyy') }; }
function _mDataCarga(v) { return Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, FUSO, 'dd/MM/yyyy') : String(v || '').substring(0, 10); }
function _mDias(ent) { var m = /^(\d{2})\/(\d{2})\/(\d{2,4})$/.exec(_dataBr(ent) || ''); if (!m) return null; var a = +m[3]; if (a < 100) a += 2000; var h = new Date(); h.setHours(0, 0, 0, 0); return Math.round((new Date(a, +m[2] - 1, +m[1]) - h) / 864e5); }
function _mOsHoje(hoje) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_REVISOES), o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (_txtDh(r[0]).indexOf(hoje) === 0) o[pvOs_(r[1])] = 1; });
  return o;
}

/** 18h — pedidos que entraram hoje e não são revisão de pedido que já existia. */
function emailPedidosDia() {
  var H = _mHoje(), sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PEDIDOS);
  if (!sh || sh.getLastRow() < 2) return;
  var P = _idx(COLS_PEDIDOS), rev = _mOsHoje(H.br);
  var L = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_PEDIDOS.length).getValues().filter(function (r) {
    return _mDataCarga(r[P.data_carga]) === H.br && String(r[P.vigente]).trim() === 'SIM' && !rev[pvOs_(r[P.oc])] && String(r[P.somente_pendentes] || '') !== 'SIM';
  });
  if (!L.length) { Logger.log('Nenhum pedido novo hoje.'); return; }
  L.sort(function (a, b) { return (Number(b[P.total_cheio]) || 0) - (Number(a[P.total_cheio]) || 0); });
  var tot = 0, t40 = 0, nit = 0, urg = 0, cl = {}, semTab = 0, abaixo = 0;
  L.forEach(function (r) {
    var v = Number(r[P.total_cheio]) || 0; tot += v; t40 += Number(r[P.total_40]) || 0; nit += Number(r[P.qtd_itens]) || 0;
    var d = _mDias(r[P.data_entrega]); if (d !== null && d <= 15) urg++;
    var c = String(r[P.cliente] || r[P.cli_pailon]); cl[c] = cl[c] || [0, 0]; cl[c][0]++; cl[c][1] += v;
    if (r[P.desvio_valor] === '' ) semTab++; else if (Number(r[P.desvio_valor]) < -0.005) abaixo++;
  });
  var ks = Object.keys(cl).sort(function (a, b) { return cl[b][1] - cl[a][1]; }), top = ks[0];
  var corpo = _mKpis([['Valor cheio', _moeda0(tot), 'venda ao cliente'], ['Valor Dynamics', _moeda0(t40), 'valor dos PDFs'],
                      ['Pedidos · itens', L.length + ' · ' + nit, ks.length + ' cliente(s)'], ['Entrega ≤ 15 dias', String(urg), 'prazo curto', '#FFC35C']]);
  corpo += _mSec('Por cliente');
  ks.forEach(function (k) { var w = Math.max(2, Math.round(cl[k][1] / (tot || 1) * 100));
    corpo += '<tr><td style="padding:6px 20px"><table width="100%" cellpadding="0" cellspacing="0"><tr><td style="font-size:14px;font-weight:700">' + k + ' <span style="color:#8C96A4;font-weight:400;font-size:12px">· ' + cl[k][0] + ' ped.</span></td><td align="right" style="font-size:14px;font-weight:700;white-space:nowrap">' + _moeda0(cl[k][1]) + '</td></tr>' +
      '<tr><td colspan="2" style="padding-top:5px"><table width="100%" cellpadding="0" cellspacing="0"><tr><td width="' + w + '%" style="background:#15181D;height:6px"></td><td style="background:#E9ECF0;height:6px"></td></tr></table></td></tr></table></td></tr>'; });
  corpo += _mSec('Pedidos do dia');
  L.forEach(function (r) {
    var d = _mDias(r[P.data_entrega]), ent = _dataBr(r[P.data_entrega]), dv = r[P.desvio_valor], rt = String(r[P.tipo]) === 'RETRABALHO';
    var tb = dv === '' ? _mChip('sem tabela', '#EEF1F5', '#566070') : Number(dv) < -0.005 ? _mChip('abaixo da tabela', '#FDE8E9', '#D0121C') : Number(dv) > 0.005 ? _mChip('acima da tabela', '#E3F5EC', '#0E7A4E') : _mChip('igual à tabela', '#E6EEFB', '#1F5FBF');
    var pz = d === null ? '<span style="font-size:12.5px;color:#566070">sem data de entrega</span>' : d <= 15 ? _mChip('entrega ' + ent + ' · ' + (d < 0 ? Math.abs(d) + ' dias vencida' : d + ' dias'), '#FFF3DC', '#B86E00') : '<span style="font-size:12.5px;color:#566070">entrega ' + ent + ' · ' + d + ' dias</span>';
    var n = Number(r[P.qtd_itens]) || 0;
    corpo += '<tr><td style="padding:0 20px"><table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:1px solid #EEF1F5"><tr><td style="padding:12px 0 2px;font-size:12px;color:#8C96A4;font-family:Consolas,Menlo,monospace">' + (rt ? 'RETRABALHO ' + String(r[P.oc]).replace(/^RT/, 'ID ') : r[P.oc]) + '</td><td align="right" style="padding:12px 0 2px">' + tb + '</td></tr>' +
      '<tr><td style="font-size:15px;font-weight:800">' + (r[P.cliente] || r[P.cli_pailon]) + '</td><td align="right" style="font-size:16px;font-weight:800;white-space:nowrap">' + _moeda(r[P.total_cheio]) + '</td></tr>' +
      '<tr><td colspan="2" style="font-size:12.5px;color:#566070;padding:2px 0 4px">' + [r[P.revenda], r[P.cidade], r[P.uf]].filter(String).join(' · ') + ' · ' + n + (n === 1 ? ' item' : ' itens') + '</td></tr>' +
      '<tr><td colspan="2" style="padding:0 0 12px">' + pz + '</td></tr></table></td></tr>';
  });
  if (semTab || abaixo) corpo += _mSec('Atenção comercial') + '<tr><td style="padding:0 20px 4px"><table width="100%" cellpadding="0" cellspacing="0"><tr><td style="background:#FFF8E6;border-left:4px solid #B86E00;padding:12px 14px;font-size:13.5px;line-height:1.6">' +
    (semTab ? '• <b>' + semTab + ' pedido(s) sem preço em tabela</b> — conferir a tabela do cliente<br>' : '') + (abaixo ? '• <b>' + abaixo + ' pedido(s) abaixo da tabela</b>' : '• Nenhum pedido abaixo da tabela') + '</td></tr></table></td></tr>';
  var lead = '<b>' + top + '</b> respondeu por ' + Math.round(cl[top][1] / (tot || 1) * 100) + '% do valor do dia.' + (urg ? ' ' + urg + ' pedido(s) com entrega em até 15 dias.' : '');
  _mEnviar('PEDIDOS', 'Fechamento ' + H.br.substring(0, 5) + ' · ' + L.length + ' pedido(s) novo(s) · ' + _moeda0(tot), H.sub, L.length + ' pedido(s) novo(s)<br><span style="color:#D0121C">' + _moeda0(tot) + '</span>', lead, corpo);
}

/** 18h — pedidos revisados hoje, agrupados por tipo de mudança. */
function emailRevisoesDia() {
  var H = _mHoje(), ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(ABA_REVISOES);
  if (!sh || sh.getLastRow() < 2) return;
  var g = {}, ord = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, COLS_REVISOES.length).getValues().forEach(function (r) {
    if (_txtDh(r[0]).indexOf(H.br) !== 0) return;
    var k = String(r[1]).trim(); if (!g[k]) { g[k] = { oc: k, de: pvRev_(r[2]), para: pvRev_(r[3]), l: [] }; ord.push(k); }
    g[k].l.push(r);
  });
  if (!ord.length) { Logger.log('Nenhuma revisão hoje.'); return; }
  var shP = ss.getSheetByName(ABA_PEDIDOS), P = _idx(COLS_PEDIDOS), vig = {}, ant = {};
  shP.getRange(2, 1, shP.getLastRow() - 1, COLS_PEDIDOS.length).getValues().forEach(function (r) {
    var oc = String(r[P.oc]).trim(), os = pvOs_(oc);
    if (String(r[P.vigente]).trim() === 'SIM') vig[oc] = r; else (ant[os] = ant[os] || []).push(r);
  });
  var nSem = 0, nAnt = 0, nExc = 0, imp = 0, cards = [];
  ord.forEach(function (k) {
    var x = g[k], r = vig[k], sem = x.l.some(function (l) { return /SEM TROCA/.test(l[5]); }), sa = x.l.some(function (l) { return /SEM ANTERIOR/.test(l[5]); });
    var a = (ant[pvOs_(k)] || []).filter(function (q) { return pvRev_(q[P.revisao]) === x.de; }).pop();
    var dv = r && a ? (Number(r[P.total_cheio]) || 0) - (Number(a[P.total_cheio]) || 0) : 0;
    imp += dv; if (sem) nSem++; if (sa) nAnt++;
    var ch = x.l.filter(function (l) { return !/^ALERTA$/.test(l[4]) && !/SEM MUDANCA/.test(l[5]); });
    var exc = ch.filter(function (l) { return /EXCLU/.test(l[5]); }).length; nExc += exc;
    var grav = sem ? 3 : exc ? 2 : sa ? 1 : 0;
    var cor = grav >= 2 ? '#D0121C' : sa ? '#B86E00' : '#1F5FBF';
    var bd = (sem ? _mChip('SEM TROCA DE REVISÃO', '#D0121C', '#fff') : sa ? _mChip('REV ' + x.para + ' sem anterior na base', '#FFF3DC', '#B86E00') : _mChip('REV ' + x.de + ' → ' + x.para, '#15181D', '#fff')) +
      (r && a && Math.abs(dv) > 0.005 ? ' ' + _mChip((dv > 0 ? '+' : '−') + _moeda0(Math.abs(dv)).replace('R$ ', 'R$ ') + (a[P.total_cheio] ? ' · ' + (dv > 0 ? '+' : '−') + Math.abs(dv / a[P.total_cheio] * 100).toFixed(1).replace('.', ',') + '%' : ''), dv < 0 ? '#FDE8E9' : '#E3F5EC', dv < 0 ? '#D0121C' : '#0E7A4E') : '');
    var alerta = sem ? 'O pedido mudou e o número da revisão não foi trocado no ForWood. O app já atualizou; alinhar com o comercial.' : exc ? '⚠ ' + exc + ' item(ns) excluído(s) na revisão — o PCP deve confirmar se já estavam em produção.' : '';
    var gr = {}, ordG = [['EXCLU', 'Excluídos', '#FDE8E9', '#D0121C'], ['INCLU', 'Incluídos', '#E3F5EC', '#0E7A4E'], ['QUANT', 'Quantidade alterada', '#FFF3DC', '#B86E00'], ['PRE', 'Preço alterado', '#FFF3DC', '#B86E00'], ['DESCRI', 'Descrição alterada', '#FFF3DC', '#B86E00'], ['ENTREGA', 'Data de entrega', '#E6EEFB', '#1F5FBF'], ['', 'Comercial / outros', '#EEF1F5', '#566070']];
    ch.forEach(function (l) { var t = String(l[5]).toUpperCase(), o = ordG.filter(function (q) { return !q[0] || t.indexOf(q[0]) >= 0; })[0]; (gr[o[1]] = gr[o[1]] || { o: o, l: [] }).l.push(l); });
    var muds = ordG.filter(function (o) { return gr[o[1]]; }).map(function (o) {
      return '<tr><td style="padding:10px 14px 4px;border-top:1px solid #EEF1F5">' + _mChip(o[1].toUpperCase() + ' · ' + gr[o[1]].l.length, o[2], o[3]) + '</td></tr>' + gr[o[1]].l.map(function (l) {
        var de = _txtRev(l[8]), pa = _txtRev(l[9]), ex = /EXCLU/.test(l[5]), inc = /INCLU/.test(l[5]);
        return '<tr><td style="padding:4px 14px 8px">' + (l[6] ? '<div style="font-size:13px"><b style="font-family:Consolas,Menlo,monospace">' + l[6] + '</b> · ' + String(l[7] || '') + '</div>' : (!/ENTREGA/.test(l[5]) ? '<div style="font-size:13px;color:#3A4350">' + _acent(String(l[5]).toLowerCase()) + '</div>' : '')) +
          '<div style="font-size:13.5px;margin-top:3px">' + (ex ? '<span style="color:#8C96A4;text-decoration:line-through">' + de + '</span> &nbsp;<b style="color:#D0121C">excluído</b>' : inc ? '<b>' + pa + '</b> &nbsp;<span style="color:#0E7A4E">novo</span>' : '<span style="color:#8C96A4">' + de + '</span> &nbsp;→&nbsp; <b>' + pa + '</b>') + '</div></td></tr>';
      }).join('');
    }).join('');
    var sub = r ? [r[P.revenda], r[P.cidade], r[P.uf]].filter(String).join(' · ') + ' · entrega ' + _dataBr(r[P.data_entrega]) : '';
    cards.push({ g: grav, h: '<tr><td style="padding:6px 20px 10px"><table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #E2E6EC;border-left:5px solid ' + cor + ';border-radius:8px"><tr><td style="padding:12px 14px"><div style="font-size:12px;color:#8C96A4;font-family:Consolas,Menlo,monospace">' + k + '</div><div style="font-size:16px;font-weight:800;margin-top:2px">' + (r ? r[P.cliente] : '') + '</div><div style="font-size:12.5px;color:#566070;margin-top:2px">' + sub + '</div><div style="margin-top:8px">' + bd + '</div></td></tr>' +
      (alerta ? '<tr><td style="padding:0 14px 12px"><div style="background:' + (sem || exc ? '#FDE8E9' : '#FFF3DC') + ';color:#8A0B12;font-size:13px;line-height:1.5;padding:10px 12px;border-radius:6px;font-weight:600">' + alerta + '</div></td></tr>' : '') +
      (sa && !ch.length ? '<tr><td style="padding:0 14px 12px;font-size:13px;color:#566070">Mudanças não apuradas: a revisão anterior não está no app.</td></tr>' : muds) + '</table></td></tr>' });
  });
  cards.sort(function (a, b) { return b.g - a.g; });
  var corpo = _mKpis([['Revisados', String(ord.length), 'no dia'], ['Impacto no valor', (imp < 0 ? '−' : imp > 0 ? '+' : '') + _moeda0(Math.abs(imp)), 'valor cheio', imp < 0 ? '#FF7A7F' : '#fff'],
                      ['Sem troca de revisão', String(nSem), 'alinhar com o comercial', nSem ? '#FF7A7F' : '#fff'], ['Itens excluídos', String(nExc), 'conferir na produção', nExc ? '#FF7A7F' : '#fff']]) +
    _mSec('Revisões · por gravidade') + cards.map(function (c) { return c.h; }).join('');
  var lead = [nSem ? '<b>' + nSem + '</b> pedido(s) mudaram <b>sem troca de revisão</b>' : '', nExc ? '<b>' + nExc + '</b> item(ns) excluído(s)' : '', nAnt ? '<b>' + nAnt + '</b> revisão(ões) sem a anterior na base' : ''].filter(String).join(' · ');
  _mEnviar('REVISOES', 'Fechamento ' + H.br.substring(0, 5) + ' · ' + ord.length + ' pedido(s) revisado(s)' + (nSem ? ' · ' + nSem + ' sem troca de revisão' : ''), H.sub, ord.length + ' pedido(s) revisado(s)', lead, corpo);
}

/** Botão na tela Faturamento: envia o faturamento lançado hoje (sem gatilho). */
function enviarEmailFaturamento(tk) {
  var s = _sessao(tk);
  if (s.t.indexOf('flan') < 0) return { ok: false, msg: 'Seu usuário não tem a tela Faturamento.' };
  var H = _mHoje(), L = _lancamentos().filter(function (x) { return x.dh.indexOf(H.br) === 0; });
  if (!L.length) return { ok: false, msg: 'Nenhum faturamento lançado hoje.' };
  if (!_destinos('FATURAMENTO').length) return { ok: false, msg: 'Ninguém marcado com FATURAMENTO na coluna avisos da aba Acessos.' };
  // V4.0 — o faturamento lançado hoje é a data da nota (Data Atendimento) dos pedidos
  //         expedidos no dia útil anterior. O e-mail mostra as duas datas.
  var dts = {}; L.forEach(function (x) { dts[x.dnf] = 1; });
  var dl = Object.keys(dts).sort(function (a, b) { return _dNum(a) - _dNum(b); });
  var dxs = {}; L.forEach(function (x) { dxs[x.dex || _dataBr(_diaUtilAnterior(x.dnf))] = 1; });
  var rotF = dl.map(function (d) { return d.substring(0, 5); }).join(', '), rotE = Object.keys(dxs).sort(function (a, b) { return _dNum(a) - _dNum(b); }).map(function (d) { return d.substring(0, 5); }).join(', ');
  var ped = {}, ord = [], tot = 0, nfs = {}, alt = 0;
  L.forEach(function (x) { if (!ped[x.oc]) { ped[x.oc] = { oc: x.oc, cli: x.cli, it: [], v: 0, nf: {} }; ord.push(x.oc); } var p = ped[x.oc]; p.it.push(x); p.v += x.tf; p.nf[x.nf] = 1; tot += x.tf; nfs[x.nf] = 1; if (x.alt) alt++; });
  ord.sort(function (a, b) { return ped[b].v - ped[a].v; });
  var corpo = _mKpis([['Faturado (Dynamics)', _moeda0(tot), 'valor das notas'], ['Valor cheio', _moeda0(tot * FATOR_CHEIO), 'referência 40%'],
                      ['Pedidos · itens', ord.length + ' · ' + L.length, Object.keys(nfs).length + ' nota(s)'], ['Itens alterados', String(alt), 'qtde ou valor diferente', alt ? '#FFC35C' : '#fff']]) + _mSec('Pedidos faturados');
  ord.forEach(function (k) { var p = ped[k];
    corpo += '<tr><td style="padding:6px 20px 10px"><table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #E2E6EC;border-left:5px solid #0E7A4E;border-radius:8px"><tr><td style="padding:12px 14px"><table width="100%" cellpadding="0" cellspacing="0"><tr><td style="font-size:12px;color:#8C96A4;font-family:Consolas,Menlo,monospace">' + p.oc + '</td><td align="right">' + _mChip('NF ' + Object.keys(p.nf).join(', '), '#E3F5EC', '#0E7A4E') + '</td></tr>' +
      '<tr><td style="font-size:16px;font-weight:800">' + p.cli + '</td><td align="right" style="font-size:16px;font-weight:800;white-space:nowrap">' + _moeda(p.v) + '</td></tr></table></td></tr>' +
      p.it.map(function (x) { return '<tr><td style="padding:6px 14px;border-top:1px solid #EEF1F5;font-size:13px"><b style="font-family:Consolas,Menlo,monospace">' + x.c + '</b> · ' + x.d + '<div style="margin-top:3px;color:#3A4350">' + x.qf + (x.qf !== x.qp ? ' <span style="color:#B86E00">de ' + x.qp + '</span>' : '') + ' × ' + (x.alt && Math.abs(x.uf - x.up) > 0.005 ? '<b style="color:#D0121C">' + _moeda(x.uf) + '</b> <span style="color:#8C96A4">(pedido ' + _moeda(x.up) + ')</span>' : _moeda(x.uf)) + ' = <b>' + _moeda(x.tf) + '</b></div></td></tr>'; }).join('') +
      '</table></td></tr>'; });
  var n = _mEnviar('FATURAMENTO', 'Faturamento ' + rotF + ' (expedição ' + rotE + ') · ' + ord.length + ' pedido(s) · ' + _moeda0(tot),
    'Notas de ' + rotF + ' · pedidos expedidos em ' + rotE, 'Faturamento de ' + rotF + '<br><span style="color:#0E7A4E">' + _moeda0(tot) + '</span>',
    'Pedidos <b>expedidos em ' + rotE + '</b> e faturados em <b>' + rotF + '</b>. Conferido e enviado por <b>' + s.n + '</b>.', corpo);
  _log(s, 'FATURAMENTO E-MAIL', ord.length + ' pedido(s) · ' + _moeda(tot) + ' → ' + n + ' destinatário(s)');
  return { ok: true, msg: 'E-mail do faturamento enviado para ' + n + ' pessoa(s).' };
}
function _moeda0(n) { return 'R$ ' + Math.round(Number(n) || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }

function _dNum(d) { var m = /^(\d{2})\/(\d{2})\/(\d{2,4})$/.exec(String(d || '')); if (!m) return 0; var a = +m[3]; if (a < 100) a += 2000; return a * 10000 + (+m[2]) * 100 + (+m[1]); }
/** dia útil anterior (sexta, se a nota é de segunda) */
function _diaUtilAnterior(d) {
  var m = /^(\d{2})\/(\d{2})\/(\d{2,4})$/.exec(String(d || '')); if (!m) return '';
  var a = +m[3]; if (a < 100) a += 2000; var x = new Date(a, +m[2] - 1, +m[1]);
  do { x.setDate(x.getDate() - 1); } while (x.getDay() === 0 || x.getDay() === 6);
  return Utilities.formatDate(x, FUSO, 'dd/MM/yyyy');
}

/* ===================== V4.3 — MOTOR DE DATAS DE ENTREGA E CONFERÊNCIA DE REVISÕES =====================
 * A data de entrega é gerida por terceiros no ForWood e não chega ao app. O Upload recebe o relatório de
 * pedidos de venda do ForWood (Excel), compara com a base e:
 *   1) mostra as datas que mudaram e, depois da conferência, atualiza SOMENTE a data de entrega;
 *   2) registra cada mudança na aba HistoricoDatas (e marca o pedido no app);
 *   3) devolve as diferenças de revisão/itens/valores para o Orçamento subir os PDFs mais atuais.
 * Colunas esperadas: Codigo, Produto Codigo, QTD, VALOR UNIT, PREVISAO, Numero Ordem Compra, Status. */
var ABA_HD = 'HistoricoDatas';
var COLS_HD = ['data_hora', 'usuario', 'id_pedido', 'oc', 'os', 'numero_erp', 'cliente', 'data_anterior', 'data_nova', 'dif_dias', 'arquivo'];
var DATAS_SEM_DATA = ['30/12', '31/12'];      // datas-curinga do ForWood: significam "sem data definida"

function _abaHD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(ABA_HD);
  if (!sh) {
    sh = ss.insertSheet(ABA_HD);
    sh.getRange(1, 1, 1, COLS_HD.length).setValues([COLS_HD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
/** Data (Date, dd/mm/aa, dd/mm/aaaa ou ISO) -> {s:'dd/MM/yyyy', n:aaaammdd, md:'dd/MM', ms:'aaaa-mm'} */
function _pdt(v) {
  if (v === null || v === undefined || v === '') return null;
  var d;
  if (Object.prototype.toString.call(v) === '[object Date]') { if (isNaN(v.getTime())) return null; d = Utilities.formatDate(v, FUSO, 'dd/MM/yyyy'); }
  else d = String(v).trim();
  var m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(d);
  if (!m) {
    var t = new Date(d); if (isNaN(t.getTime())) return null;
    m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(Utilities.formatDate(t, FUSO, 'dd/MM/yyyy')); if (!m) return null;
  }
  var dd = ('0' + m[1]).slice(-2), mm = ('0' + m[2]).slice(-2), aa = +m[3]; if (aa < 100) aa += 2000;
  return { s: dd + '/' + mm + '/' + aa, n: aa * 10000 + (+mm) * 100 + (+dd), md: dd + '/' + mm, ms: aa + '-' + mm };
}
function _dnDias(a, b) {
  function t(n) { return Date.UTC(Math.floor(n / 10000), Math.floor(n / 100) % 100 - 1, n % 100); }
  return Math.round((t(b.n) - t(a.n)) / 864e5);
}
function _nErp(x) { return String(x === null || x === undefined ? '' : x).trim().replace(/\.0+$/, ''); }
function _nn(x) {
  if (typeof x === 'number') return x;
  var t = String(x === null || x === undefined ? '' : x).trim(); if (!t) return 0;
  if (t.indexOf(',') >= 0) t = t.replace(/\./g, '').replace(',', '.');
  var f = parseFloat(t); return isNaN(f) ? 0 : f;
}
/** OC no mesmo padrão nos dois lados: 35.26.001-02 = 035.26.001-02 · RT11189 = 11189 */
function _nOc(t) {
  t = String(t || '').trim().replace(/^RT/i, '');
  var m = /^(\d+)\.(\d+)\.(\d+)-(\d+)$/.exec(t);
  return m ? ('00' + m[1]).slice(-3) + '.' + m[2] + '.' + ('00' + m[3]).slice(-3) + '-' + ('0' + m[4]).slice(-2) : t;
}
function _nOcPartes(t) { var m = /^(.*)-(\d+)$/.exec(_nOc(t)); return m ? { os: m[1], rev: +m[2] } : { os: _nOc(t), rev: null }; }
function _nh(x) { return String(x).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim(); }

/** Compara as linhas do relatório do ForWood (matriz com cabeçalho) com a base. Não grava nada. */
function _prevAnalisar_(v) {
  var cab = (v[0] || []).map(_nh);
  function col(re) { for (var i = 0; i < cab.length; i++) if (re.test(cab[i])) return i; return -1; }
  var C = { erp: col(/^codigo$/), obs: col(/observacao/), c: col(/produto codigo/), d: col(/produto descricao/), q: col(/^qtde?$|^qtd /), u: col(/^valor unit/),
            t: col(/^valor total/), pv: col(/previsao/), sit: col(/^situacao$/), oc: col(/ordem (de )?compra/), st: col(/^status$/) };
  if (C.erp < 0 || C.pv < 0 || C.oc < 0 || C.st < 0 || C.c < 0 || C.q < 0 || C.u < 0)
    return { ok: false, msg: 'Planilha fora do padrão. Preciso do relatório de pedidos de venda do ForWood, com as colunas Codigo, Produto Codigo, QTD, VALOR UNIT, PREVISAO, Numero Ordem Compra e Status.' };
  var E = {};
  for (var i = 1; i < v.length; i++) {
    var r = v[i], erp = _nErp(r[C.erp]); if (!erp) continue;
    var e = E[erp] = E[erp] || { erp: erp, oc: '', obs: '', st: '', sit: '', tot: 0, prev: [], it: {} };
    if (!e.oc) e.oc = String(r[C.oc]).trim();
    if (!e.obs && C.obs >= 0) e.obs = String(r[C.obs]).trim();
    if (!e.st) e.st = String(r[C.st]).trim();
    if (!e.sit && C.sit >= 0) e.sit = String(r[C.sit]).trim();
    var pd = _pdt(r[C.pv]); if (pd && !e.prev.some(function (x) { return x.n === pd.n; })) e.prev.push(pd);
    var q = _nn(r[C.q]), u = _nn(r[C.u]), t = C.t >= 0 ? _nn(r[C.t]) : q * u, c = String(r[C.c]).trim();
    var x = e.it[c] = e.it[c] || { c: c, d: C.d >= 0 ? String(r[C.d]).trim() : '', q: 0, t: 0 };
    x.q += q; x.t += t; e.tot += t;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet(), shP = ss.getSheetByName(ABA_PEDIDOS), shI = ss.getSheetByName(ABA_ITENS);
  var P = _idx(COLS_PEDIDOS), I = _idx(COLS_ITENS), app = {};
  if (shP && shP.getLastRow() > 1) shP.getRange(2, 1, shP.getLastRow() - 1, COLS_PEDIDOS.length).getValues().forEach(function (r) {
    if (String(r[P.vigente]).trim() !== 'SIM') return;
    var erp = _nErp(r[P.numero_erp]); if (!erp) return;
    app[erp] = { id: String(r[P.id_pedido]).trim(), oc: String(r[P.oc]).trim(), rev: pvRev_(r[P.revisao]), erp: erp,
      cli: String(r[P.cliente] || '').trim() || String(r[P.cli_pailon] || '').trim(), ent: _pdt(r[P.data_entrega]),
      v40: _r2(r[P.total_40]), v: _r2(r[P.total_cheio]), rt: String(r[P.tipo] || '').trim() === 'RETRABALHO' };
  });
  var itApp = {};
  if (shI && shI.getLastRow() > 1) shI.getRange(2, 1, shI.getLastRow() - 1, COLS_ITENS.length).getValues().forEach(function (r) {
    var k = String(r[I.id_pedido]).trim() + '|' + pvRev_(r[I.revisao]), c = String(r[I.codigo]).trim();
    var m = itApp[k] = itApp[k] || {}, x = m[c] = m[c] || { c: c, d: String(r[I.descricao] || '').trim(), q: 0, t: 0, u: Number(r[I.preco_40]) || 0 };
    var q = Number(r[I.qtde]) || 0; x.q += q; x.t += q * (Number(r[I.preco_40]) || 0);
  });
  var FAT = _faturadoPorItem(), STT = _statusPorOs();
  var R = { ok: true, datas: [], semdata: [], revs: [], itens: [], faltam: [], naoAprov: [], baixas: [], baixaBloq: [], parcSem: [], res: { app: Object.keys(app).length, conf: 0, noRel: 0, igual: 0, atend: 0, datas: 0, semdata: 0, nova: 0, antiga: 0, os: 0, valor: 0, faltam: 0, cancel: 0, baixa: 0, baixaIt: 0, jaBaixado: 0, bloq: 0, parcSem: 0 } };
  Object.keys(app).forEach(function (erp) {
    var a = app[erp], e = E[erp];
    if (!e) { R.res.noRel++; return; }
    R.res.conf++;
    var stc = STT[pvOs_(a.oc)]; if (stc && stc.s === 'CANCELADO') { R.res.cancel++; return; }   // V4.4 — cancelado no app continua cancelado: sem data, sem baixa, sem revisão
    if (/nao aprovado/.test(_nh(e.sit))) R.naoAprov.push({ oc: a.oc, erp: erp, cli: a.cli, st: e.st });
    if (/atendido total/.test(_nh(e.st))) { R.res.atend++; _prevBaixa_(R, a, e, itApp, FAT); return; }   // V4.4 — baixa do que o ERP já atendeu
    if (/parcial/.test(_nh(e.st)) && !a.rt) {                                                          // conferência: parcial no ERP sem nenhum lançamento no app
      var os0 = pvOs_(a.oc), ia0 = itApp[a.id + '|' + a.rev] || {}, alg = Object.keys(ia0).some(function (c) { var f = FAT[os0 + '|' + c]; return f && f.q > 0; });
      if (!alg) { R.parcSem.push({ oc: a.oc, erp: erp, cli: a.cli, ent: a.ent ? a.ent.s : '', vE: _r2(e.tot) }); R.res.parcSem++; }
    }
    var pk = e.prev.slice().sort(function (x, y) { return x.n - y.n; }), nova = pk.length ? pk[pk.length - 1] : null;
    // ---- data de entrega
    if (nova) {
      var curinga = DATAS_SEM_DATA.indexOf(nova.md) >= 0;
      if (curinga && (!a.ent || DATAS_SEM_DATA.indexOf(a.ent.md) < 0)) {
        R.semdata.push({ id: a.id, oc: a.oc, erp: erp, cli: a.cli, ent: a.ent ? a.ent.s : '', fw: nova.s, st: e.st }); R.res.semdata++;
      } else if (!a.ent || a.ent.n !== nova.n) {
        R.datas.push({ id: a.id, oc: a.oc, erp: erp, cli: a.cli, ant: a.ent ? a.ent.s : '', nov: nova.s, dd: a.ent ? _dnDias(a.ent, nova) : null,
          mm: !!(a.ent && a.ent.ms !== nova.ms), st: e.st, v: a.v, multi: pk.length > 1 }); R.res.datas++;
      } else R.res.igual++;
    }
    // ---- revisão / OS / valor
    if (a.rt) return;
    var pa = _nOcPartes(a.oc), pe = _nOcPartes(e.oc), tipo = '', acao = '';
    var parcial = /parcial/.test(_nh(e.st)), dv = _r2(a.v40 - e.tot);
    if (_nOc(a.oc) !== _nOc(e.oc)) {
      if (pa.os === pe.os && pe.rev !== null && pa.rev !== null) {
        if (pe.rev > pa.rev) { tipo = 'NOVA'; acao = 'Subir o PDF da revisão ' + ('0' + pe.rev).slice(-2) + ' (o app está na ' + ('0' + pa.rev).slice(-2) + ')'; }
        else { tipo = 'ANTIGA'; acao = 'O ForWood está na revisão ' + ('0' + pe.rev).slice(-2) + ' e o app na ' + ('0' + pa.rev).slice(-2) + ': conferir no ForWood'; }
      } else { tipo = 'OS'; acao = 'No ForWood o pedido ' + erp + ' é a OS ' + e.oc + '; no app está como ' + a.oc + ': conferir o PDF carregado'; }
    } else if (!parcial && Math.abs(dv) > Math.max(2, e.tot * 0.0005)) { tipo = 'VALOR'; acao = 'Mesma revisão, valor diferente: conferir preços e itens e subir o PDF atualizado se necessário'; }
    if (!tipo) return;
    var ni = 0;
    if (!parcial) {
      var ia = itApp[a.id + '|' + a.rev] || {}, codes = {}; Object.keys(ia).forEach(function (c) { codes[c] = 1; }); Object.keys(e.it).forEach(function (c) { codes[c] = 1; });
      Object.keys(codes).sort().forEach(function (c) {
        var xa = ia[c], xe = e.it[c], ue = xe && xe.q ? xe.t / xe.q : 0, ua = xa && xa.q ? xa.t / xa.q : 0, d = '';
        if (xa && !xe) d = 'NÃO EXISTE NO FORWOOD'; else if (!xa && xe) d = 'INCLUÍDO NO FORWOOD';
        else if (Math.abs(xa.q - xe.q) > 0.0001) d = 'QUANTIDADE'; else if (Math.abs(ua - ue) > 0.011) d = 'PREÇO';
        if (d) { ni++; R.itens.push({ erp: erp, oc: a.oc, cli: a.cli, c: c, d: (xe || xa).d, t: d, qa: xa ? xa.q : null, qe: xe ? xe.q : null, ua: xa ? _r2(ua) : null, ue: xe ? _r2(ue) : null }); }
      });
    } else acao += ' (parcial: o ForWood não informa o saldo; conferir os itens ao subir)';
    R.revs.push({ id: a.id, oc: a.oc, revA: a.rev, ocE: e.oc, erp: erp, cli: a.cli, st: e.st, ent: nova ? nova.s : '', v40: a.v40, vE: _r2(e.tot), dv: parcial ? null : dv, tipo: tipo, acao: acao, ni: ni });
    R.res[tipo === 'NOVA' ? 'nova' : tipo === 'ANTIGA' ? 'antiga' : tipo === 'OS' ? 'os' : 'valor']++;
  });
  Object.keys(E).forEach(function (erp) {
    var e = E[erp]; if (app[erp] || /atendido total/.test(_nh(e.st))) return;
    var pk = e.prev.slice().sort(function (x, y) { return x.n - y.n; });
    R.faltam.push({ erp: erp, oc: e.oc, cli: e.obs, ent: pk.length ? pk[pk.length - 1].s : '', st: e.st, vE: _r2(e.tot) }); R.res.faltam++;
  });
  var po = { NOVA: 0, OS: 1, VALOR: 2, ANTIGA: 3 };
  R.revs.sort(function (x, y) { return po[x.tipo] - po[y.tipo] || String(x.oc).localeCompare(String(y.oc)); });
  R.datas.sort(function (x, y) { return _pdt(x.nov).n - _pdt(y.nov).n; });
  R.faltam.sort(function (x, y) { return (_pdt(x.ent) || { n: 0 }).n - (_pdt(y.ent) || { n: 0 }).n; });
  return R;
}
/** V4.4 — pedido "Atendido Total" no ERP: calcula o saldo que ainda não foi baixado no app. Só baixa se OS e valor conferem. */
function _prevBaixa_(R, a, e, itApp, FAT) {
  var ia = itApp[a.id + '|' + a.rev] || {}, os = pvOs_(a.oc), dv = _r2(a.v40 - e.tot), motivo = '';
  if (_nOc(a.oc) !== _nOc(e.oc) && !a.rt) motivo = 'OS diferente no ForWood (' + e.oc + ')';
  else if (!a.rt && Math.abs(dv) > Math.max(2, e.tot * 0.0005)) motivo = 'valor diferente (app ' + _r2(a.v40) + ' × ForWood ' + _r2(e.tot) + ')';
  var its = [], tot = 0;
  Object.keys(ia).forEach(function (c) {
    var x = ia[c], f = FAT[os + '|' + c], fq = f ? f.q : 0, sd = _r2(x.q - fq);
    if (sd > 0.0001) { its.push({ c: c, d: x.d, qp: x.q, qf: sd, up: _r2(x.u), uf: _r2(x.u) }); tot += sd * x.u; }
  });
  if (!its.length) { R.res.jaBaixado++; return; }
  if (motivo) { R.baixaBloq.push({ oc: a.oc, erp: a.erp, cli: a.cli, motivo: motivo, vE: _r2(e.tot) }); R.res.bloq++; return; }
  R.baixas.push({ id: a.id, oc: a.oc, rev: a.rev, erp: a.erp, cli: a.cli, itens: its, total: _r2(tot) }); R.res.baixa++; R.res.baixaIt += its.length;
}
/** V4.4 — grava as baixas na aba Faturamento (NF "S/NF", data de hoje, origem relatório ForWood). Refaz o saldo dentro do bloqueio: não duplica. */
function _gravarBaixas_(s, baixas, arq) {
  if (!baixas || !baixas.length) return { n: 0, it: 0, tot: 0 };
  var sh = _abaFat(), agora = _agora(), hoje = Utilities.formatDate(new Date(), FUSO, 'dd/MM/yyyy'), dex = _diaUtilAnterior(hoje), FAT = _faturadoPorItem(), linhas = [], peds = 0, tot = 0;
  baixas.forEach(function (B) {
    var os = pvOs_(B.oc), n0 = linhas.length;
    B.itens.forEach(function (x) {
      var f = FAT[os + '|' + x.c], sd = _r2(x.qp - (f ? f.q : 0));
      if (!(sd > 0.0001)) return;
      tot += sd * x.uf;
      linhas.push([agora, s.n + ' (baixa relatório ForWood)', 'S/NF', hoje, B.id, B.oc, os, pvRev_(B.rev), B.cli, x.c, x.d, x.qp, sd, x.up, x.uf, Math.round(sd * x.uf * 100) / 100, 'NAO', dex]);
    });
    if (linhas.length > n0) peds++;
  });
  if (!linhas.length) return { n: 0, it: 0, tot: 0 };
  sh.getRange(sh.getLastRow() + 1, 1, linhas.length, COLS_FAT.length).setValues(linhas);
  _log(s, 'BAIXA PELO RELATÓRIO FORWOOD', peds + ' pedido(s) · ' + linhas.length + ' item(ns) · R$ ' + (Math.round(tot * 100) / 100) + ' · ' + (arq || ''));
  return { n: peds, it: linhas.length, tot: Math.round(tot * 100) / 100 };
}
/** V4.4 — passo único: lê o Excel do ForWood, ATUALIZA as datas e DÁ BAIXA no que o ERP já atendeu. Devolve a conferência (revisões, parcial sem lançamento, etc.). */
function processarPrevisoes(tk, b64, nome) {
  var s = _sessao(tk);
  if (s.t.indexOf('up') < 0) return { ok: false, msg: 'Seu usuário não tem a tela Upload.' };
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), MimeType.MICROSOFT_EXCEL, nome || 'pedidos.xlsx');
  var f = Drive.Files.insert ? Drive.Files.insert({ title: '__prev__' + (nome || ''), mimeType: MimeType.GOOGLE_SHEETS }, blob, { convert: true })
                             : Drive.Files.create({ name: '__prev__' + (nome || ''), mimeType: MimeType.GOOGLE_SHEETS }, blob);
  var v;
  try { v = SpreadsheetApp.openById(f.id).getSheets()[0].getDataRange().getValues(); }
  finally { try { DriveApp.getFileById(f.id).setTrashed(true); } catch (e) {} }
  var lock = LockService.getScriptLock(); lock.tryLock(30000);
  try {
    var R = _prevAnalisar_(v);
    if (!R.ok) return R;
    var g = _gravarDatas_(s, R.datas.map(function (d) { return { id: d.id, nov: d.nov }; }), nome);
    R.datasFeitas = g.n; R.baixaFeita = _gravarBaixas_(s, R.baixas, nome);
    _log(s, 'RELATÓRIO FORWOOD PROCESSADO', (nome || '') + ' · ' + (v.length - 1) + ' linha(s) · ' + g.n + ' data(s) atualizada(s) · ' + R.baixaFeita.n + ' pedido(s) baixado(s) · ' + (R.res.nova + R.res.os + R.res.valor) + ' revisão(ões)/valor(es) a conferir');
    return R;
  } finally { lock.releaseLock(); }
}
/** Passo 1 — lê o Excel do ForWood e devolve a conferência (nada é gravado). */
function lerPrevisoes(tk, b64, nome) {
  var s = _sessao(tk);
  if (s.t.indexOf('up') < 0) return { ok: false, msg: 'Seu usuário não tem a tela Upload.' };
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), MimeType.MICROSOFT_EXCEL, nome || 'pedidos.xlsx');
  var f = Drive.Files.insert ? Drive.Files.insert({ title: '__prev__' + (nome || ''), mimeType: MimeType.GOOGLE_SHEETS }, blob, { convert: true })
                             : Drive.Files.create({ name: '__prev__' + (nome || ''), mimeType: MimeType.GOOGLE_SHEETS }, blob);
  try {
    var v = SpreadsheetApp.openById(f.id).getSheets()[0].getDataRange().getValues();
    var R = _prevAnalisar_(v);
    if (R.ok) _log(s, 'DATAS FORWOOD LIDAS', (nome || '') + ' · ' + (v.length - 1) + ' linha(s) · ' + R.res.datas + ' data(s) diferente(s) · ' + (R.res.nova + R.res.os + R.res.valor) + ' revisão(ões)/valor(es) a conferir');
    return R;
  } finally { try { DriveApp.getFileById(f.id).setTrashed(true); } catch (e) {} }
}
/** Passo 2 — grava SOMENTE a data de entrega dos pedidos marcados e registra o histórico. lista = [{id, nov:'dd/mm/aaaa'}] */
function gravarPrevisoes(tk, lista, arq) {
  var s = _sessao(tk);
  if (s.t.indexOf('up') < 0) return { ok: false, msg: 'Seu usuário não tem a tela Upload.' };
  var lock = LockService.getScriptLock(); lock.tryLock(30000);
  try { var g = _gravarDatas_(s, lista, arq); return g.n ? { ok: true, n: g.n, msg: g.n + ' data(s) de entrega atualizada(s).' } : { ok: true, n: 0, msg: g.msg || 'Nenhuma data precisou ser alterada.' }; }
  finally { lock.releaseLock(); }
}
function _gravarDatas_(s, lista, arq) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_PEDIDOS), P = _idx(COLS_PEDIDOS);
  if (!sh || sh.getLastRow() < 2) return { n: 0, msg: 'Base de pedidos vazia.' };
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, COLS_PEDIDOS.length).getValues(), pos = {};
  v.forEach(function (r, i) { if (String(r[P.vigente]).trim() === 'SIM') pos[String(r[P.id_pedido]).trim()] = i; });
  var hd = [], agora = _agora();
  (lista || []).forEach(function (L) {
    var i = pos[String(L.id).trim()]; if (i === undefined) return;
    var nov = _pdt(L.nov); if (!nov) return;
    var r = v[i], ant = _pdt(r[P.data_entrega]);
    if (ant && ant.n === nov.n) return;
    var cel = sh.getRange(i + 2, P.data_entrega + 1);
    cel.setValue(Utilities.parseDate(nov.s, FUSO, 'dd/MM/yyyy')); cel.setNumberFormat('dd/MM/yy');
    hd.push([agora, s.n, String(r[P.id_pedido]), String(r[P.oc]), pvOs_(r[P.oc]), String(r[P.numero_erp]),
      String(r[P.cliente] || '').trim() || String(r[P.cli_pailon] || '').trim(), ant ? ant.s : '', nov.s, ant ? _dnDias(ant, nov) : '', arq || '']);
  });
  if (!hd.length) return { n: 0 };
  var h = _abaHD(); h.getRange(h.getLastRow() + 1, 1, hd.length, COLS_HD.length).setValues(hd);
  _log(s, 'DATAS DE ENTREGA ATUALIZADAS', hd.length + ' pedido(s) · ' + (arq || ''));
  return { n: hd.length };
}
/** Última alteração de data por pedido (para marcar no app). */
function _alteracoesData() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_HD), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, COLS_HD.length).getValues().forEach(function (r) {
    var id = String(r[2]).trim(); if (!id) return;
    var o = m[id] = m[id] || { k: 0 };
    o.a = _dataBr(r[7]); o.n = _dataBr(r[8]); o.dh = _txtDh(r[0]).substring(0, 10); o.u = String(r[1]); o.k++;
  });
  return m;
}
function _ultimasDatas() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA_HD);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = Math.min(300, sh.getLastRow() - 1);
  return sh.getRange(sh.getLastRow() - n + 1, 1, n, COLS_HD.length).getValues().map(function (r) {
    return { dh: _txtDh(r[0]), u: String(r[1]), oc: String(r[3]), erp: String(r[5]), cli: String(r[6]), a: _dataBr(r[7]), n: _dataBr(r[8]), dd: r[9] === '' ? '' : Number(r[9]) };
  }).reverse();
}
