// Preencha apenas se o script NÃO estiver vinculado à planilha (Extensões > Apps Script).
const ID_PLANILHA = '';

const CONFIG = {
  estagiarios: {
    titulo: 'Estagiários',
    singular: 'estagiário',
    aba: 'Estágiários',
    linhaGrupo: 1,
    linhaSubgrupo: 2,
    linhaCabecalho: 3,
    grupos: [
      ['A', 'H', 'Dados pessoais'],
      ['I', 'O', 'Análise documental'],
      ['P', 'W', 'Contratação'],
      ['X', 'AE', 'Responsáveis'],
      ['AF', 'BC', 'Acompanhamento mensal'],
      ['BD', 'BQ', 'Cursos de capacitação']
    ],
    // Rótulos exibidos no lugar do cabeçalho da planilha, por coluna.
    rotulos: {
      Z: 'Supervisor de estágio', AA: 'CPF do supervisor', AB: 'E-mail do supervisor',
      AC: 'Professor orientador', AD: 'CPF do orientador', AE: 'E-mail do orientador'
    },
    // Campos extras que editam o link (hiperlink) da célula da coluna indicada.
    links: { A: 'Link da pasta' },
    tabela: ['Nome', 'Curso', 'Instituição de Ensino', 'UG', 'Mentor', 'Status', 'Status da Análise'],
    filtros: ['Status', 'Status da Análise', 'Chamada', 'UG', 'Instituição de Ensino', 'Gestor Responsável', 'Mentor'],
    sugestoes: ['Curso', 'Área', 'Cidade', 'Gestor Responsável', 'Mentor', 'Reitor', 'Sup. de Estágio', 'Prof. Orientador'],
    textoLongo: ['Observações'],
    obrigatorios: ['Nome', 'CPF'],
    unico: 'CPF',
    dicaBusca: 'Pesquisar por nome, CPF, curso, instituição, cidade…'
  },
  mentores: {
    titulo: 'Mentores',
    singular: 'mentor',
    aba: 'Mentores',
    linhaCabecalho: 2,
    grupos: [['A', 'E', 'Dados do mentor']],
    tabela: ['Nome', 'Área', 'CPF', 'Email', 'Whatsapp'],
    filtros: ['Área'],
    sugestoes: ['Área'],
    textoLongo: [],
    obrigatorios: ['Nome'],
    unico: 'CPF',
    dicaBusca: 'Pesquisar por nome, área, CPF, e-mail…',
    vinculo: { entidade: 'estagiarios', rotulo: 'Mentor' }
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('EstágioFlow')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

function carregarDados() {
  const entidades = {};
  Object.keys(CONFIG).forEach(function (k) {
    const ctx = contexto_(k), cfg = ctx.cfg;
    entidades[k] = {
      titulo: cfg.titulo,
      singular: cfg.singular,
      dicaBusca: cfg.dicaBusca,
      campos: ctx.campos,
      tabela: indices_(ctx, cfg.tabela),
      filtros: indices_(ctx, cfg.filtros),
      sugestoes: indices_(ctx, cfg.sugestoes || []),
      obrigatorios: indices_(ctx, cfg.obrigatorios || []),
      vinculo: cfg.vinculo || null,
      registros: lerRegistros_(ctx)
    };
  });
  return { usuario: Session.getActiveUser().getEmail(), entidades: entidades };
}

function salvarRegistro(chave, linha, versao, alteracoes) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ctx = contexto_(chave);
    linha = Number(linha);
    if (!(linha >= ctx.inicio && linha <= ctx.sh.getLastRow())) throw new Error('Linha inválida.');
    const rg = ctx.sh.getRange(linha, 1, 1, ctx.n);
    const atual = lerLinhas_(ctx, linha, 1)[0];
    if (versao_(atual) !== versao) {
      throw new Error('CONFLITO: Este registro foi alterado por outra pessoa ou mudou de posição na planilha. Clique em “Atualizar dados” e refaça a alteração.');
    }
    const indices = Object.keys(alteracoes || {}).map(Number).filter(function (i) { return i >= 0 && i < ctx.campos.length; });
    if (!indices.length) return registro_(ctx, linha);

    const novo = atual.slice();
    indices.forEach(function (i) { novo[i] = String(alteracoes[i] == null ? '' : alteracoes[i]).trim(); });
    validar_(ctx, novo, linha, indices);

    const formulas = rg.getFormulas()[0];
    indices.forEach(function (i) {
      const c = ctx.campos[i];
      if (i < ctx.n && c.link === undefined && formulas[i]) throw new Error('O campo "' + c.rotulo + '" é calculado por fórmula e não pode ser editado.');
    });
    const comLink = {};
    indices.forEach(function (i) {
      const c = ctx.campos[i];
      if (c.tipo === 'link') comLink[c.origem] = true;
      else if (c.link !== undefined) comLink[i] = true;
      else ctx.sh.getRange(linha, i + 1).setValue(converter_(c, novo[i], ctx.tz));
    });
    Object.keys(comLink).forEach(function (o) { escreverLink_(ctx, linha, Number(o), novo); });
    SpreadsheetApp.flush();
    return registro_(ctx, linha);
  } finally {
    lock.releaseLock();
  }
}

function criarRegistro(chave, valores) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ctx = contexto_(chave), sh = ctx.sh;
    const v = ctx.campos.map(function (c, i) { return String(valores && valores[i] != null ? valores[i] : '').trim(); });
    validar_(ctx, v, null, ctx.campos.map(function (c) { return c.i; }));

    const linha = Math.max(sh.getLastRow(), ctx.cfg.linhaCabecalho) + 1;
    if (linha > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);
    const destino = sh.getRange(linha, 1, 1, ctx.n);
    if (linha > ctx.inicio) {
      const modelo = sh.getRange(ctx.inicio, 1, 1, ctx.n);
      modelo.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      modelo.copyTo(destino, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
    destino.setValues([v.slice(0, ctx.n).map(function (s, i) { return converter_(ctx.campos[i], s, ctx.tz); })]);
    ctx.campos.forEach(function (c) { if (c.tipo === 'link' && v[c.i]) escreverLink_(ctx, linha, c.origem, v); });
    SpreadsheetApp.flush();
    return registro_(ctx, linha);
  } finally {
    lock.releaseLock();
  }
}

// Execute no editor para conferir se a leitura da planilha está correta.
function diagnostico() {
  const d = carregarDados();
  Object.keys(d.entidades).forEach(function (k) {
    const e = d.entidades[k];
    console.log(e.titulo + ': ' + e.registros.length + ' registros, ' + e.campos.length + ' campos');
    e.campos.forEach(function (c) {
      console.log('  [' + c.grupo + '] ' + (c.sub ? c.sub + ' / ' : '') + c.rotulo + ' (' + c.tipo + (c.opcoes ? ': ' + c.opcoes.join(', ') : '') + ')');
    });
  });
}

function planilha_() {
  const ss = ID_PLANILHA ? SpreadsheetApp.openById(ID_PLANILHA) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Planilha não encontrada. Preencha ID_PLANILHA em Code.gs.');
  return ss;
}

function contexto_(chave) {
  const cfg = CONFIG[chave];
  if (!cfg) throw new Error('Entidade inválida: ' + chave);
  const ss = planilha_();
  const sh = ss.getSheetByName(cfg.aba);
  if (!sh) throw new Error('Aba "' + cfg.aba + '" não encontrada na planilha.');
  const tz = ss.getSpreadsheetTimeZone();
  const cab = sh.getRange(cfg.linhaCabecalho, 1, 1, sh.getLastColumn()).getValues()[0];
  let n = cab.length;
  while (n > 0 && String(cab[n - 1]).trim() === '') n--;
  if (!n) throw new Error('Cabeçalho vazio na aba "' + cfg.aba + '".');
  return { cfg: cfg, sh: sh, tz: tz, n: n, inicio: cfg.linhaCabecalho + 1, campos: lerCampos_(sh, cfg, cab.slice(0, n), tz) };
}

function lerCampos_(sh, cfg, cab, tz) {
  const n = cab.length;
  const linhaGrupo = cfg.linhaGrupo ? sh.getRange(cfg.linhaGrupo, 1, 1, n).getDisplayValues()[0] : [];
  const linhaSub = cfg.linhaSubgrupo ? sh.getRange(cfg.linhaSubgrupo, 1, 1, n).getDisplayValues()[0] : [];
  const modelo = sh.getRange(cfg.linhaCabecalho + 1, 1, 1, n);
  const validacoes = modelo.getDataValidations()[0];
  const formatos = modelo.getNumberFormats()[0];
  const grupos = cfg.grupos.map(function (g) {
    const ini = coluna_(g[0]);
    return { ini: ini, fim: coluna_(g[1]), nome: String(linhaGrupo[ini - 1] || '').trim() || g[2] };
  });

  const fixos = {};
  Object.keys(cfg.rotulos || {}).forEach(function (l) { fixos[coluna_(l)] = cfg.rotulos[l]; });

  let sub = '';
  const campos = cab.map(function (v, i) {
    const col = i + 1;
    const g = grupos.filter(function (x) { return col >= x.ini && col <= x.fim; })[0];
    if (!g || g.ini === col) sub = '';
    if (String(linhaSub[i] || '').trim()) sub = String(linhaSub[i]).trim();
    const rotulo = fixos[col] || (v instanceof Date ? Utilities.formatDate(v, tz, 'MM/yyyy') : String(v).trim());
    const opcoes = opcoes_(validacoes[i]);
    let tipo = 'texto';
    if (opcoes) tipo = 'lista';
    else if (cfg.textoLongo.indexOf(rotulo) >= 0) tipo = 'textoLongo';
    else if (/d/i.test(formatos[i]) && /y/i.test(formatos[i])) tipo = 'data';
    else if (/e-?mail/i.test(rotulo)) tipo = 'email';
    return { i: i, rotulo: rotulo, sub: g ? sub : '', grupo: g ? g.nome : 'Outros', tipo: tipo, opcoes: opcoes };
  });

  const total = {}, cont = {};
  const chaveDe = function (c) { return c.grupo + '|' + c.sub + '|' + c.rotulo; };
  campos.forEach(function (c) { total[chaveDe(c)] = (total[chaveDe(c)] || 0) + 1; });
  campos.forEach(function (c) {
    const k = chaveDe(c);
    if (total[k] > 1) { cont[k] = (cont[k] || 0) + 1; c.rotulo += ' (' + cont[k] + ')'; }
    c.pos = c.i;
  });
  Object.keys(cfg.links || {}).forEach(function (l) {
    const base = campos[coluna_(l) - 1];
    if (!base) return;
    base.link = campos.length;
    campos.push({ i: campos.length, rotulo: cfg.links[l], sub: base.sub, grupo: base.grupo, tipo: 'link', opcoes: null, origem: base.i, pos: base.i + 0.5 });
  });
  return campos;
}

function opcoes_(dv) {
  if (!dv) return null;
  const tipo = dv.getCriteriaType(), args = dv.getCriteriaValues();
  if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return args[0].map(String);
  if (tipo === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    return [].concat.apply([], args[0].getDisplayValues()).filter(String);
  }
  return null;
}

function lerRegistros_(ctx) {
  const total = ctx.sh.getLastRow() - ctx.inicio + 1;
  if (total < 1) return [];
  const saida = [];
  lerLinhas_(ctx, ctx.inicio, total).forEach(function (v, k) {
    if (v.some(function (x) { return x !== ''; })) saida.push({ linha: ctx.inicio + k, versao: versao_(v), v: v });
  });
  return saida;
}

function registro_(ctx, linha) {
  const v = lerLinhas_(ctx, linha, 1)[0];
  return { linha: linha, versao: versao_(v), v: v };
}

function lerLinhas_(ctx, inicio, total) {
  const linhas = ctx.sh.getRange(inicio, 1, total, ctx.n).getValues().map(function (r) {
    return r.map(function (x, i) { return serializar_(x, ctx.campos[i], ctx.tz); });
  });
  ctx.campos.forEach(function (c) {
    if (c.tipo !== 'link') return;
    const rg = ctx.sh.getRange(inicio, c.origem + 1, total, 1);
    const ricos = rg.getRichTextValues(), formulas = rg.getFormulas();
    linhas.forEach(function (v, k) { v[c.i] = urlDe_(ricos[k][0], formulas[k][0]); });
  });
  return linhas;
}

function urlDe_(rico, formula) {
  const m = /^=HYPERLINK\(\s*"([^"]+)"/i.exec(formula || '');
  if (m) return m[1];
  if (!rico) return '';
  if (rico.getLinkUrl()) return rico.getLinkUrl();
  const trecho = rico.getRuns().filter(function (r) { return r.getLinkUrl(); })[0];
  return trecho ? trecho.getLinkUrl() : '';
}

function escreverLink_(ctx, linha, origem, v) {
  const c = ctx.campos[origem], texto = v[origem], url = v[c.link];
  const cel = ctx.sh.getRange(linha, origem + 1);
  if (!texto) { cel.setValue(''); return; }
  const rico = SpreadsheetApp.newRichTextValue().setText(texto);
  if (url) rico.setLinkUrl(url);
  cel.setRichTextValue(rico.build());
}

function validar_(ctx, v, linha, indices) {
  const cfg = ctx.cfg, campos = ctx.campos, erros = [];
  indices_(ctx, cfg.obrigatorios || []).forEach(function (i) {
    if ((linha === null || indices.indexOf(i) >= 0) && !v[i]) erros.push('Preencha o campo "' + campos[i].rotulo + '".');
  });
  indices.forEach(function (i) {
    const c = campos[i], s = v[i];
    if (!s) return;
    if (c.opcoes && c.opcoes.indexOf(s) < 0) erros.push('O valor "' + s + '" não é permitido em "' + c.rotulo + '".');
    if (c.tipo === 'data' && !/^\d{4}-\d{2}-\d{2}$/.test(s)) erros.push('Data inválida em "' + c.rotulo + '".');
    if (c.tipo === 'email' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) erros.push('E-mail inválido em "' + c.rotulo + '".');
    if (c.tipo === 'link' && !/^https?:\/\/\S+$/i.test(s)) erros.push('O link em "' + c.rotulo + '" deve começar com https://.');
    if (c.tipo === 'link' && !v[c.origem]) erros.push('Preencha "' + campos[c.origem].rotulo + '" para salvar o link.');
  });

  const u = cfg.unico ? indices_(ctx, [cfg.unico])[0] : undefined;
  if (u !== undefined && indices.indexOf(u) >= 0 && digitos_(v[u])) {
    const total = ctx.sh.getLastRow() - ctx.inicio + 1;
    if (total > 0) {
      const alvo = digitos_(v[u]);
      const coluna = ctx.sh.getRange(ctx.inicio, u + 1, total, 1).getValues();
      for (let k = 0; k < coluna.length; k++) {
        if (ctx.inicio + k !== linha && digitos_(coluna[k][0]) === alvo) {
          erros.push('Já existe um registro com este ' + campos[u].rotulo + ' (linha ' + (ctx.inicio + k) + ' da planilha).');
          break;
        }
      }
    }
  }
  if (erros.length) throw new Error(erros.join(' '));
}

function serializar_(v, campo, tz) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, tz, campo.tipo === 'data' ? 'yyyy-MM-dd' : 'dd/MM/yyyy');
  return String(v);
}

function converter_(campo, s, tz) {
  if (s === '') return '';
  if (campo.tipo === 'data' && /^\d{4}-\d{2}-\d{2}$/.test(s)) return Utilities.parseDate(s, tz, 'yyyy-MM-dd');
  if (!campo.opcoes && /^[=+\-@]/.test(s)) return "'" + s;
  return s;
}

function versao_(v) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(v), Utilities.Charset.UTF_8));
}

function indices_(ctx, rotulos) {
  return rotulos.map(function (r) {
    for (let i = 0; i < ctx.campos.length; i++) if (ctx.campos[i].rotulo === r) return i;
    return -1;
  }).filter(function (i) { return i >= 0; });
}

function digitos_(s) {
  const d = String(s == null ? '' : s).replace(/\D/g, '');
  return d ? ('00000000000' + d).slice(-Math.max(11, d.length)) : '';
}

function coluna_(letras) {
  return String(letras).toUpperCase().split('').reduce(function (n, ch) { return n * 26 + ch.charCodeAt(0) - 64; }, 0);
}
