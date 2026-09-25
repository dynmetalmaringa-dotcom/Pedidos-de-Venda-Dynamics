/**********************************************************************
 * FLUXO VENDA DYNAMICS — APLICATIVO WEB
 * Arquivo: App              Versão: V2.0
 *
 * Login por usuário e senha, termo de confidencialidade/LGPD,
 * registro de tudo (acessos, downloads, uploads) na aba LogAcoes.
 * Lê a base que Tabelas e Pedidos alimentam. Não recalcula preço.
 *
 * FUNÇÕES PARA O EDITOR
 *   instalarV2()                 cria abas Acessos/LogAcoes e os usuários
 *   redefinirSenha('PAULA','Nova@123')   troca a senha de alguém
 *   bloquearUsuario('PAULA')     tira o acesso (ativo = NAO)
 *   situacaoApp()                confere a base e o tempo de carga
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
var TERMO_VERSAO = 'V1.0';
var COLS_ACESSOS = ['usuario','perfil','nome','senha_hash','ativo','aceite_termo','ultimo_acesso'];
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

function instalarV2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var velha = ss.getSheetByName(ABA_ACESSOS);
  if (velha && String(velha.getRange(1, 1).getValues()[0][0]).trim() !== 'usuario') {
    velha.setName('Acessos_V1_antigo');
    velha = null;
  }
  var sh = velha || ss.insertSheet(ABA_ACESSOS);
  if (sh.getLastRow() < 1 || String(sh.getRange(1, 1).getValues()[0][0]).trim() !== 'usuario') {
    sh.getRange(1, 1, 1, COLS_ACESSOS.length).setValues([COLS_ACESSOS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var USU = [
    ['MICHELE','dir','Michele Puma','f606b5e21e5363be17df490442f512ee673ac6548df8acd1568cb2b749443b0d'],
    ['PAULA','dir','Paula','74806745aa91b5d502411c3416e541a7122ea230f53696c376ab1bd379a81798'],
    ['LUCIANO','dir','Luciano','cb30ed08a079e63cd291f8a90985eed04920ab7a3d0282e3f23d220163ebecdf'],
    ['CESAR','dir','Cesar','6fc05843858102e43a71b51b4cab7fc3986c770532036021dd124308bc9e6fa7'],
    ['FERNANDA','orc','Fernanda','fee7d54368e6c0280ff0ebfe67d1b4c97eb83510e61bcf88565a417a397477be'],
    ['JULIANA','orc','Juliana','8c32648eeff04b920d29aad66ac98f1b3c26ba25eb7e9bb5dc48b5a00d098fac'],
    ['SUPORTE','adm','Suporte','47e8d4f632497f0e0cb37e094e0b0f261cd256107eb26e82bbcec2487e4ea918']
  ];
  var ja = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .forEach(function (r) { ja[String(r[0]).trim().toUpperCase()] = true; });
  var novos = USU.filter(function (u) { return !ja[u[0]]; })
    .map(function (u) { return [u[0], u[1], u[2], u[3], 'SIM', '', '']; });
  if (novos.length) sh.getRange(sh.getLastRow() + 1, 1, novos.length, COLS_ACESSOS.length).setValues(novos);

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
  Logger.log(novos.length + ' usuário(s) criado(s). Abas Acessos e LogAcoes prontas.');
  Logger.log('Perfis: dir = Diretoria + Orçamento | orc = Orçamento | adm = suporte (oculto).');
}

function _hash(usuario, senha) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
          'FVD|' + String(usuario).trim().toUpperCase() + '|' + String(senha), Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}

function redefinirSenha(usuario, nova) {
  if (!usuario || !nova || String(nova).length < 6) { Logger.log('Informe usuário e senha com 6+ caracteres.'); return; }
  var a = _acharUsuario(usuario);
  if (!a) { Logger.log('Usuário não encontrado: ' + usuario); return; }
  a.sh.getRange(a.linha, 4).setValue(_hash(usuario, nova));
  a.sh.getRange(a.linha, 5).setValue('SIM');
  _log({ u: 'EDITOR', p: 'adm' }, 'SENHA REDEFINIDA', String(usuario).toUpperCase());
  Logger.log('Senha de ' + String(usuario).toUpperCase() + ' redefinida.');
}
function bloquearUsuario(usuario) {
  var a = _acharUsuario(usuario);
  if (!a) { Logger.log('Usuário não encontrado: ' + usuario); return; }
  a.sh.getRange(a.linha, 5).setValue('NAO');
  Logger.log(String(usuario).toUpperCase() + ' bloqueado.');
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

/* ====================== LOGIN E SESSÃO ====================== */

function login(usuario, senha, aceite) {
  var a = _acharUsuario(usuario);
  var u = String(usuario || '').trim().toUpperCase();
  if (!a || String(a.r[3]) !== _hash(u, senha)) {
    _log({ u: u || '?', p: '' }, 'LOGIN RECUSADO', 'usuário ou senha incorretos');
    return { ok: false, msg: 'Usuário ou senha incorretos.' };
  }
  if (String(a.r[4]).trim().toUpperCase() === 'NAO') return { ok: false, msg: 'Usuário bloqueado. Fale com a Diretoria.' };
  var aceitou = String(a.r[5]).indexOf(TERMO_VERSAO) >= 0;
  if (!aceitou && !aceite) return { ok: false, termo: true, msg: 'Leia e aceite o termo para entrar.' };
  var agora = _agora();
  var perfil = String(a.r[1]).trim().toLowerCase();
  if (perfil === 'ambos' || perfil === 'todos') perfil = 'dir';
  var s = { u: u, p: perfil, n: String(a.r[2] || u).trim() };
  if (!aceitou) {
    a.sh.getRange(a.linha, 6).setValue('Termo ' + TERMO_VERSAO + ' aceito em ' + agora);
    _log(s, 'ACEITE DO TERMO', 'Confidencialidade e LGPD ' + TERMO_VERSAO);
  }
  a.sh.getRange(a.linha, 7).setValue(agora);
  var tk = Utilities.getUuid();
  CacheService.getScriptCache().put('fvd_' + tk, JSON.stringify(s), SESSAO_SEG);
  _log(s, 'LOGIN', '');
  return { ok: true, token: tk, perfil: perfil, nome: s.n, usuario: u };
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
  var out = { eu: { u: s.u, p: s.p, n: s.n }, hoje: Utilities.formatDate(new Date(), FUSO, 'dd/MM/yy'),
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
    var DYN = _mapaDyn();
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
        v: _r2(v), v40: _r2(v40), itens: itens
      };
      p.sv = v < 1;
      p.cd = DYN.cod[p.cp] || _dynPorNome(DYN, nome) ||
             (!p.sv && itens.length ? String(itens[0].c).substring(0, 3) : '') || p.cp;
      var comp = 0, tab = 0, n = 0;
      itens.forEach(function (x) { if (x.tb > 0) { comp += x.t; tab += x.tb * x.q; n++; } });
      p.comp = _r2(comp); p.tab = _r2(tab); p.ct = n;
      out.P.push(p);
    });
  }

  if (s.p === 'orc' || s.p === 'dir' || s.p === 'adm') {
    var a = _auditoria();
    out.tabelas = a.tabelas; out.tabItens = a.itens;
    out.logs = _ultimosLogs(s.p === 'adm');
  }
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
    return verOculto || String(r[2]).trim() !== 'adm';
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

/* ====================== DIAGNÓSTICO ====================== */

function situacaoApp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  function conta(n) { var s = ss.getSheetByName(n); return s ? Math.max(0, s.getLastRow() - 1) : -1; }
  [ABA_PEDIDOS, ABA_ITENS, ABA_PRECOS, ABA_LOGOS, ABA_ACESSOS, ABA_LOG].forEach(function (n) {
    Logger.log(n + ': ' + conta(n) + ' linha(s)');
  });
  var tk = Utilities.getUuid();
  CacheService.getScriptCache().put('fvd_' + tk, JSON.stringify({ u: 'EDITOR', p: 'adm', n: 'Editor' }), 60);
  var t = Date.now(), d = carregar(tk);
  Logger.log(d.P.length + ' pedidos vigentes · ' + d.tabelas.length + ' tabelas · carga ' + (Date.now() - t) + ' ms');
}
