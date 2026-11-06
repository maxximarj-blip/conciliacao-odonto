/* ===== Motor de conciliação — Odonto Premier ===== */
const Core = (() => {
  const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
  const nameKey = s => norm(s).replace(/[^a-z0-9 ]/g, ' ').replace(/\b(dr|dra|drª|doutor|doutora)\b/g, ' ').replace(/\s+/g, ' ').trim();

  function guiaKey(v) {
    if (v === null || v === undefined) return '';
    let s = typeof v === 'number' ? String(Math.round(v)) : String(v);
    s = s.trim().toUpperCase().replace(/\s+/g, '').replace(/\.0+$/, '');
    if (/^\d+$/.test(s)) s = s.replace(/^0+(?=\d)/, '');
    return s;
  }

  function toNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    if (v === null || v === undefined || v === '') return 0;
    let s = String(v).replace(/[R$\s]/g, '');
    if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }

  function toDate(v) {
    if (v instanceof Date && !isNaN(v)) return v;
    if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000) + 43200000);
    if (typeof v === 'string') {
      const m = v.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
      if (m) return new Date(+(m[3].length === 2 ? '20' + m[3] : m[3]), +m[2] - 1, +m[1], 12);
      const d = new Date(v); if (!isNaN(d)) return d;
    }
    return null;
  }
  const GENERIC = /cl[ií]nico geral|sem dentista|n[aã]o informado|avalia[cç][aã]o|endodontia \(|^-?$/i;
  function limpaTexto(v, proc = false) {
    let t = String(v ?? '').replace(/\s+/g, ' ').trim();
    t = t.replace(/^[\s:;,.\-–—•·|/\\]+/, '').replace(/[\s:;,\-–—|]+$/, '');
    t = t.replace(/\s+([,.;:)])/g, '$1').replace(/([(])\s+/g, '$1').replace(/([,;])(?=[^\s\d])/g, '$1 ');
    if (proc) {
      const letras = t.replace(/[^A-Za-zÀ-ÿ]/g, '');
      const up = letras.replace(/[^A-ZÀ-Þ]/g, '').length;
      if (letras.length > 3 && up / letras.length > 0.6) t = t.toLowerCase();
      t = t.charAt(0).toUpperCase() + t.slice(1);
    }
    return t;
  }
  const dayKey = d => d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}` : '';

  /* ---- detecção de colunas da planilha do sistema ---- */
  const FIELD_RULES = {
    guia: h => /guia/.test(h) && !/status|valor|situa|pend/.test(h),
    dentista: h => /dentista|profissional|executor/.test(h),
    valor: h => /valor|executado|vlr|total/.test(h) && !/pend/.test(h),
    data: h => /^data|dt /.test(h) || /data do atend/.test(h),
    situacao: h => /situa/.test(h),
    status: h => /status/.test(h),
    procedimento: h => /procedimento/.test(h) && !/cod/.test(h),
    paciente: h => /paciente/.test(h) && !/conv|cpf/.test(h),
    atendimento: h => /atendimento/.test(h) && /id/.test(h),
    convenio: h => /conv[eê]nio|plano|operadora/.test(h),
    obs: h => /observ|motivo/.test(h),
  };

  function findHeaderRow(rows) {
    let best = -1, bestScore = 0;
    for (let i = 0; i < Math.min(rows.length, 25); i++) {
      const r = rows[i] || [];
      let score = 0;
      r.forEach(c => { const h = norm(c); if (!h) return; for (const k in FIELD_RULES) if (FIELD_RULES[k](h)) { score++; break; } });
      if (score > bestScore) { bestScore = score; best = i; }
    }
    return { index: best, score: bestScore };
  }

  function mapColumns(header) {
    const map = {};
    const hs = header.map(norm);
    for (const k in FIELD_RULES) {
      const idx = hs.findIndex((h, i) => h && FIELD_RULES[k](h) && !Object.values(map).includes(i));
      if (idx >= 0) map[k] = idx;
    }
    // preferências específicas
    const pick = (k, re) => { const i = hs.findIndex(h => re.test(h)); if (i >= 0) map[k] = i; };
    pick('guia', /^(numero|n[ºo°]?|num)[ _.]*(da )?guia$/);
    pick('valor', /^valor guia|^executado$/);
    pick('procedimento', /^nome[ _]procedimento$|^procedimento$/);
    return map;
  }

  /* Avalia cada aba e escolhe a melhor para a produção do sistema */
  function analyzeSystemWorkbook(sheets) { // sheets: [{name, rows}]
    return sheets.map(s => {
      const h = findHeaderRow(s.rows);
      if (h.index < 0) return { name: s.name, ok: false, score: 0 };
      const header = s.rows[h.index].map(c => String(c ?? '').trim());
      const map = mapColumns(header);
      const dataRows = s.rows.slice(h.index + 1).filter(r => map.guia !== undefined && guiaKey(r[map.guia]));
      let score = (map.guia !== undefined ? 10 : 0) + (map.valor !== undefined ? 5 : 0) + (map.dentista !== undefined ? 3 : 0) + (map.data !== undefined ? 2 : 0) + (map.situacao !== undefined ? 2 : 0) + (map.procedimento !== undefined ? 1 : 0);
      if (/detalh/.test(norm(s.name))) score += 3;
      if (!dataRows.length) score = 0;
      return { name: s.name, ok: score >= 15, score, headerIndex: h.index, header, map, count: dataRows.length };
    }).sort((a, b) => b.score - a.score);
  }

  function buildSystemRows(rows, headerIndex, map) {
    const out = [];
    rows.slice(headerIndex + 1).forEach((r, i) => {
      const g = guiaKey(r[map.guia]);
      if (!g) return;
      const get = k => map[k] !== undefined ? r[map[k]] : null;
      const d = toDate(get('data'));
      out.push({
        id: out.length,
        linha: headerIndex + i + 2,
        guia: g,
        guiaRaw: String(r[map.guia]).replace(/\.0$/, ''),
        data: d, dia: dayKey(d),
        dentistaSistema: String(get('dentista') ?? '').trim() || 'Sem Dentista',
        valor: toNum(get('valor')),
        situacao: String(get('situacao') ?? get('status') ?? '').trim() || '—',
        status: String(get('status') ?? '').trim(),
        procedimento: limpaTexto(get('procedimento'), true),
        paciente: limpaTexto(get('paciente')),
        obs: limpaTexto(get('obs')),
        convenio: String(get('convenio') ?? '').trim(),
        atendimento: get('atendimento') != null ? guiaKey(get('atendimento')) : '',
      });
    });
    return out;
  }

  /* Planilha do dentista: descobre a coluna de guia pelo maior cruzamento com o sistema;
     também acha ID do atendimento, valor, paciente e data quando existirem */
  function analyzeDentistWorkbook(sheets, systemGuias, systemAtend = new Set()) {
    let best = null;
    const colHits = (rows, c, set, from) => { const vals = new Set(); let hits = 0; rows.forEach((r, i) => { if (i <= from) return; const g = guiaKey(r?.[c]); if (g && g.length >= 5 && !vals.has(g)) { vals.add(g); if (set.has(g)) hits++; } }); return hits; };
    sheets.forEach(s => {
      const h = findHeaderRow(s.rows);
      const ncols = Math.max(0, ...s.rows.slice(0, 400).map(r => (r || []).length));
      for (let c = 0; c < ncols; c++) {
        const hits = colHits(s.rows, c, systemGuias, h.index);
        const hdr = h.index >= 0 ? norm(s.rows[h.index]?.[c]) : '';
        const score = hits + (/guia/.test(hdr) && !/status|valor/.test(hdr) ? 0.5 : 0);
        if (!best || score > best.score) best = { sheet: s.name, col: c, header: h.index >= 0 ? String(s.rows[h.index]?.[c] ?? '').trim() : '', score, headerIndex: h.index, rows: s.rows, ncols };
      }
    });
    if (!best) return { guias: [], hits: 0, itens: [], atendExtra: [] };
    const hdrs = best.headerIndex >= 0 ? (best.rows[best.headerIndex] || []).map(norm) : [];
    const findCol = re => hdrs.findIndex((x, i) => i !== best.col && re.test(x));
    // coluna de atendimento: maior cruzamento com IDs de atendimento do sistema
    let atCol = -1, atHits = 0;
    if (systemAtend.size) for (let c = 0; c < best.ncols; c++) { if (c === best.col) continue; const hh = colHits(best.rows, c, systemAtend, best.headerIndex); if (hh > atHits) { atHits = hh; atCol = c; } }
    if (atHits < 3) atCol = -1;
    const valCol = findCol(/valor|vlr|total/), pacCol = findCol(/paciente|nome/), datCol = findCol(/data/), procCol = findCol(/procedimento/);
    const itens = [], guias = [], seen = new Set();
    best.rows.forEach((r, i) => {
      if (i <= best.headerIndex || !r) return;
      const raw = r[best.col];
      let g = guiaKey(raw); if (!(g && g.length >= 5 && /\d{4,}/.test(g))) g = '';
      const at = atCol >= 0 ? guiaKey(r[atCol]) : '';
      if (!g && !at) return;
      itens.push({ linha: i + 1, guia: g, guiaRaw: g ? String(raw).replace(/\.0$/, '').trim() : '', atend: at, valor: valCol >= 0 ? toNum(r[valCol]) : null, paciente: pacCol >= 0 ? limpaTexto(r[pacCol]) : '', procedimento: procCol >= 0 ? limpaTexto(r[procCol], true) : '', data: datCol >= 0 ? toDate(r[datCol]) : null });
      if (g && !seen.has(g)) { seen.add(g); guias.push(g); }
    });
    // atendimentos usados como 2º critério: só para linhas cuja guia não existe no sistema
    const atendExtra = [...new Set(itens.filter(it => (!it.guia || !systemGuias.has(it.guia)) && it.atend && systemAtend.has(it.atend)).map(it => it.atend))];
    return { sheet: best.sheet, col: best.col, header: best.header, atCol, atHeader: atCol >= 0 && best.headerIndex >= 0 ? String(best.rows[best.headerIndex]?.[atCol] ?? '').trim() : '', temValor: valCol >= 0, guias, itens, atendExtra, hits: guias.filter(g => systemGuias.has(g)).length };
  }

  /* Encontra dentista cadastrado por nome (sistema ou nome do arquivo) */
  function matchDentist(text, dentistas) {
    const t = nameKey(text);
    if (!t) return null;
    // exato por nome ou alias
    for (const d of dentistas) {
      const keys = [d.nome, ...(d.aliases || '').split(';')].map(nameKey).filter(Boolean);
      if (keys.includes(t)) return d;
    }
    return null;
  }
  function guessDentistFromText(text, dentistas) { return guessDentistFromFilename(String(text || '') + '.x', dentistas); }
  function guessDentistFromFilename(fname, dentistas) {
    const t = ' ' + nameKey(fname.replace(/\.[^.]+$/, '').replace(/[_\-.]/g, ' ')) + ' ';
    let best = null, bestScore = 0;
    dentistas.forEach(d => {
      const keys = [d.nome, ...(d.aliases || '').split(';')].map(nameKey).filter(Boolean);
      keys.forEach(k => {
        const toks = k.split(' ').filter(w => w.length >= 3);
        if (!toks.length) return;
        const hit = toks.filter(w => t.includes(' ' + w + ' ')).length;
        const sc = hit / toks.length + (hit === toks.length ? 1 : 0) + hit * 0.01;
        if (hit && sc > bestScore) { bestScore = sc; best = d; }
      });
    });
    return best;
  }

  /* ===== Conciliação =====
     Regras:
     1) Guia informada por UM dentista → ele é o executor (o responsável do sistema não decide).
     2) Guia informada por MAIS de um → conflito (usuário escolhe).
     3) Linha informada sem guia válida, só com ID do atendimento → atribui só se o valor bater (tolerância R$ 1).
        Se não bater, fica pendente na divergência para o usuário decidir.
     4) Sem nenhuma planilha → usa o dentista do sistema, se estiver cadastrado.
     5) Decisões do usuário: resoluções de divergência, lançamentos não localizados e troca linha a linha. */
  function reconcile({ systemRows, dentistFiles, dentistas, especialidades, situacoesBase, overrides = {}, diariasOverride = {}, resolucoes = {}, lancamentos = {}, splits = {}, regrasProc = [], aprovadas = {}, tolAtend = 1, tolDiv = 1 }) {
    const byId = Object.fromEntries(dentistas.map(d => [d.id, d]));
    const files = dentistFiles.filter(f => f.dentistaId);
    const add = (m, k, v) => { if (!m.has(k)) m.set(k, new Set()); m.get(k).add(v); };
    const sysGuias = new Set(systemRows.map(r => r.guia));
    const sysAt = new Set(systemRows.map(r => r.atendimento).filter(Boolean));
    const guiaAt = new Map(); systemRows.forEach(r => { if (r.atendimento && !guiaAt.has(r.guia)) guiaAt.set(r.guia, r.atendimento); });

    const guiaToDent = new Map();
    files.forEach(f => f.guias.forEach(g => { if (sysGuias.has(g)) add(guiaToDent, g, f.dentistaId); }));

    // 2º critério: atendimento (somente linhas cuja guia não existe no sistema)
    const atInf = new Map(); // atend -> [{dentistaId, valor, temValor}]
    files.forEach(f => {
      const m = new Map();
      (f.itens || []).forEach(it => { if ((!it.guia || !sysGuias.has(it.guia)) && it.atend && sysAt.has(it.atend)) m.set(it.atend, (m.get(it.atend) || 0) + (it.valor || 0)); });
      m.forEach((v, at) => { if (!atInf.has(at)) atInf.set(at, []); atInf.get(at).push({ dentistaId: f.dentistaId, valor: v, temValor: f.temValor }); });
    });

    const sysCache = {};
    const sysDent = name => (name in sysCache) ? sysCache[name] : (sysCache[name] = matchDentist(name, dentistas));
    const baseSet = new Set(situacoesBase);

    // guias com status/situação Pendente ficam FORA do movimento (só aparecem na lista de pendentes)
    const ehPendente = r => /pend/i.test(r.situacao || '') || /pend/i.test(r.status || '');
    // ---- etapa 1: atribuição automática
    const linhas = systemRows.map(r => {
      const cands = [...(guiaToDent.get(r.guia) || [])];
      let exec = null, origem = 'nao_identificado', alerta = '', atPendente = null;
      if (cands.length === 1) { exec = cands[0]; origem = 'dentista'; }
      else if (cands.length > 1) { origem = 'conflito'; alerta = 'Guia informada por: ' + cands.map(id => byId[id]?.nome).join(', '); }
      return { ...r, executorId: exec, origem, alerta, candidatos: cands, atPendente, pend: ehPendente(r) && !aprovadas[r.id], aprovado: ehPendente(r) && !!aprovadas[r.id], pendOriginal: ehPendente(r) };
    });
    // regras automáticas por procedimento (ex.: Consulta odontológica → CLINICA) têm prioridade sobre os cruzamentos
    const regras = regrasProc.filter(rg => rg.termo && byId[rg.dentistaId]).map(rg => ({ ...rg, k: norm(rg.termo) }));
    if (regras.length) linhas.forEach(l => {
      const p = norm(l.procedimento); const rg = regras.find(x => p.includes(x.k));
      if (rg) { l.executorId = rg.dentistaId; l.origem = 'regra'; l.alerta = ''; l.regraTxt = `Regra: "${rg.termo}" → ${byId[rg.dentistaId].nome}`; l.candidatos = []; l.regra = true; }
    });
    // atendimento: candidatos = linhas do atendimento sem informante pela guia
    atInf.forEach((infs, at) => {
      const ls = linhas.filter(l => l.atendimento === at && !l.candidatos.length && !l.regra && !l.pend);
      if (!ls.length) return;
      const soma = ls.reduce((a, l) => a + l.valor, 0);
      if (infs.length === 1) {
        const i = infs[0];
        if (!i.temValor || Math.abs(i.valor - soma) <= tolAtend) ls.forEach(l => { l.executorId = i.dentistaId; l.origem = 'atendimento'; l.alerta = 'Cruzado pelo ID do atendimento (valor confere)'; });
        else ls.forEach(l => { l.atPendente = i.dentistaId; l.alerta = `Atendimento informado por ${byId[i.dentistaId]?.nome} com valor diferente — decida na divergência`; });
      } else ls.forEach(l => { l.origem = 'conflito'; l.candidatos = infs.map(i => i.dentistaId); l.alerta = 'Atendimento informado por: ' + infs.map(i => byId[i.dentistaId]?.nome).join(', '); });
    });
    // sem planilha: dentista do sistema, se cadastrado
    linhas.forEach(l => {
      if (l.executorId || l.origem === 'conflito') return;
      const sd = sysDent(l.dentistaSistema);
      if (sd) { l.executorId = sd.id; l.origem = 'sistema'; }
    });

    // ---- etapa 2: divergências (sobre a atribuição automática)
    const divergencias = [];
    files.forEach(f => {
      if (!f.temValor) return;
      const inf = new Map();
      (f.itens || []).forEach(it => {
        let k = null;
        if (it.guia && sysGuias.has(it.guia)) k = guiaAt.has(it.guia) ? 'a:' + guiaAt.get(it.guia) : 'g:' + it.guia;
        else if (it.atend && sysAt.has(it.atend)) k = 'a:' + it.atend;
        if (!k) return;
        const o = inf.get(k) || { valor: 0, paciente: it.paciente, guias: new Set(), procs: [], data: it.data };
        o.valor += it.valor || 0; if (it.guiaRaw) o.guias.add(it.guiaRaw); if (it.procedimento) o.procs.push(it.procedimento); inf.set(k, o);
      });
      inf.forEach((o, k) => {
        const tp = k[0], key = k.slice(2);
        const doKey = linhas.filter(l => (tp === 'g' ? l.guia === key : l.atendimento === key));
        if (doKey.length && doKey.every(l => l.pend)) return; // tudo pendente: fora do movimento
        const afetadas = doKey.filter(l => !l.pend && !l.regra && (l.executorId === f.dentistaId || (!l.executorId && l.origem !== 'conflito')));
        const sis = afetadas.reduce((a, l) => a + l.valor, 0);
        if (Math.abs(o.valor - sis) <= Math.max(0.05, tolDiv)) {
          afetadas.forEach(l => { if (!l.executorId) { l.executorId = f.dentistaId; l.origem = 'atendimento'; l.atPendente = null; l.alerta = 'Mesmo atendimento de guia informada; valor total confere'; } });
          return;
        }
        const id = f.name + '|' + k;
        divergencias.push({ id, arquivo: f.name, informanteId: f.dentistaId, atendimento: tp === 'a' ? key : '', guia: [...o.guias].join(', ') || '—', paciente: o.paciente, data: o.data,
          procInformado: o.procs.join('; '), procSistema: afetadas.map(l => l.procedimento).join('; '), linhaIds: afetadas.map(l => l.id), pendentes: afetadas.filter(l => !l.executorId).length,
          informado: o.valor, sistema: sis, dif: o.valor - sis, resolucao: resolucoes[id] || null });
      });
    });
    divergencias.sort((a, b) => Math.abs(b.dif) - Math.abs(a.dif));

    // ---- etapa 3: decisões do usuário
    const extras = []; // ajustes de valor, lançamentos atribuídos e divisões entre dentistas
    const inBase = l => l.aprovado || baseSet.size === 0 || baseSet.has(l.situacao);
    const nomeParts = partes => partes.map(p => `${byId[p.dentistaId]?.nome || '?'} ${Math.round(p.pct * 100) / 100}%`).join(' · ');
    const partesOk = partes => Array.isArray(partes) && partes.filter(p => p.dentistaId && p.pct > 0).length >= 1;
    const dividirLinha = (l, partes, motivo) => {
      l.executorId = null; l.origem = 'split'; l.atPendente = null; l.split = partes; l.alerta = `${motivo}: ${nomeParts(partes)}`;
      partes.forEach(p => { if (p.dentistaId && p.pct > 0) extras.push({ tipo: 'split', dentistaId: p.dentistaId, valor: l.valor * p.pct / 100, comissionavel: inBase(l), data: l.data, guia: l.guia, ref: l.guiaRaw, paciente: l.paciente, descricao: `${l.procedimento} (${p.pct}%)` }); });
    };
    divergencias.forEach(d => {
      const r = d.resolucao; if (!r || !r.decisao || r.decisao === 'pendente') return;
      if (r.decisao === 'split') {
        if (!partesOk(r.partes)) return;
        if (r.base === 'informado') {
          d.linhaIds.forEach(i => { const l = linhas[i]; l.executorId = null; l.origem = 'split'; l.atPendente = null; l.split = r.partes; l.alerta = 'Dividida (valor informado): ' + nomeParts(r.partes); l.splitInformado = true; });
          r.partes.forEach(p => { if (p.dentistaId && p.pct > 0) extras.push({ tipo: 'split', dentistaId: p.dentistaId, valor: d.informado * p.pct / 100, comissionavel: true, data: d.data || null, guia: null, ref: d.atendimento || d.guia, paciente: d.paciente, descricao: `Divisão do valor informado (${p.pct}%)` }); });
        } else d.linhaIds.forEach(i => dividirLinha(linhas[i], r.partes, 'Dividida'));
        return;
      }
      const alvo = r.dentistaId || null;
      d.linhaIds.forEach(i => { const l = linhas[i]; l.executorId = alvo; l.origem = alvo ? 'manual' : 'nao_identificado'; l.atPendente = null; l.alerta = 'Definido na divergência de valor'; });
      if (alvo && r.decisao === 'informado' && Math.abs(d.dif) > 0.004) extras.push({ tipo: 'ajuste', dentistaId: alvo, valor: d.dif, data: d.data || null, ref: d.atendimento || d.guia, paciente: d.paciente, descricao: 'Ajuste p/ valor informado' });
    });
    linhas.forEach(l => {
      if (overrides[l.id] !== undefined && l.origem !== 'split') { l.executorId = overrides[l.id] || null; l.origem = l.executorId ? 'manual' : 'nao_identificado'; l.atPendente = null; }
    });
    // divisão de guia definida na linha (vale para todas as linhas da guia)
    linhas.forEach(l => { const sp = splits[l.guia]; if (sp && partesOk(sp.partes) && !l.pend && l.origem !== 'split' && !(l.regra && overrides[l.id] === undefined)) dividirLinha(l, sp.partes, 'Guia dividida'); });
    linhas.forEach(l => { l.comissionavel = !l.pend && inBase(l) && !!l.executorId; });
    const mov = linhas.filter(l => !l.pend);
    const pendentes = linhas.filter(l => l.pendOriginal);

    // ---- lançamentos informados e não localizados no sistema
    const naoEncontradas = [];
    files.forEach(f => (f.itens || []).forEach(it => {
      if ((it.guia && sysGuias.has(it.guia)) || (it.atend && sysAt.has(it.atend))) return;
      const id = f.name + '|' + it.linha;
      const dec = lancamentos[id] || null;
      const x = { id, guia: it.guiaRaw || it.guia || '', atend: it.atend || '', paciente: it.paciente || '', procedimento: it.procedimento || '', valor: it.valor, data: it.data || null, informanteId: f.dentistaId, arquivo: f.name, decisao: dec };
      naoEncontradas.push(x);
      if (dec?.dentistaId) extras.push({ tipo: 'lancamento', dentistaId: dec.dentistaId, valor: dec.valor ?? it.valor ?? 0, data: it.data || null, ref: x.guia || x.atend, paciente: x.paciente, descricao: it.procedimento || 'Lançamento não localizado' });
    }));

    // ---- resumo por dentista
    const espById = Object.fromEntries(especialidades.map(e => [e.id, e]));
    const resumo = {};
    const get = d => resumo[d.id] || (resumo[d.id] = { dentista: d, guias: new Set(), procs: 0, producao: 0, base: 0, ajustes: 0, divididas: 0, dias: new Set(), pendente: 0 });
    mov.forEach(l => {
      const d = byId[l.executorId]; if (!d) return;
      const s = get(d); s.guias.add(l.guia); s.procs++; s.producao += l.valor;
      if (l.comissionavel) { s.base += l.valor; if (l.dia) s.dias.add(l.dia); } else s.pendente += l.valor;
    });
    extras.forEach(e => {
      const d = byId[e.dentistaId]; if (!d) return;
      const s = get(d); s.producao += e.valor;
      if (e.tipo === 'split') { s.divididas += e.valor; s.procs++; if (e.guia) s.guias.add(e.guia); }
      else s.ajustes += e.valor;
      if (e.comissionavel === false) s.pendente += e.valor;
      else { s.base += e.valor; if (e.data && e.tipo !== 'ajuste') s.dias.add(dayKey(e.data)); }
    });
    const resumoArr = Object.values(resumo).map(s => {
      const d = s.dentista, esp = espById[d.especialidadeId];
      const forma = d.forma || esp?.forma || 'percentual';
      const valorRef = Number(d.valor ?? esp?.valor ?? 0);
      const diariasCalc = s.dias.size, diarias = diariasOverride[d.id] ?? diariasCalc;
      let comissao = 0;
      if (forma === 'percentual') comissao = s.base * valorRef / 100;
      else if (forma === 'diaria') comissao = diarias * valorRef;
      else if (forma === 'fixo') comissao = valorRef;
      comissao = Math.round(comissao * 100) / 100;
      return { dentistaId: d.id, nome: d.nome, especialidade: esp?.nome || '—', forma, valorRef, guias: s.guias.size, procs: s.procs, producao: s.producao, base: s.base, ajustes: s.ajustes, divididas: s.divididas, pendente: s.pendente, diarias, diariasCalc, comissao };
    }).sort((a, b) => b.producao - a.producao);

    const divPend = divergencias.filter(d => !d.resolucao || d.resolucao.decisao === 'pendente').length;
    const tot = {
      linhas: mov.length, guias: new Set(mov.map(l => l.guia)).size,
      producao: mov.reduce((a, l) => a + l.valor, 0),
      base: resumoArr.reduce((a, r) => a + r.base, 0),
      comissao: resumoArr.reduce((a, r) => a + r.comissao, 0),
      ajustes: extras.filter(e => e.tipo !== 'split').reduce((a, e) => a + e.valor, 0),
      identificadas: mov.filter(l => l.executorId || l.origem === 'split').length,
      divididas: mov.filter(l => l.origem === 'split').length,
      viaDentista: mov.filter(l => l.origem === 'dentista' || l.origem === 'atendimento').length,
      viaAtend: mov.filter(l => l.origem === 'atendimento').length,
      viaSistema: mov.filter(l => l.origem === 'sistema').length,
      viaRegra: mov.filter(l => l.origem === 'regra').length,
      manual: mov.filter(l => l.origem === 'manual').length,
      naoIdent: mov.filter(l => !l.executorId && l.origem !== 'conflito' && l.origem !== 'split').length,
      conflitos: mov.filter(l => l.origem === 'conflito' && !l.executorId).length,
      naoIdentValor: mov.filter(l => !l.executorId && l.origem !== 'split').reduce((a, l) => a + l.valor, 0),
      naoEncontradas: naoEncontradas.length,
      neAbertos: naoEncontradas.filter(x => !x.decisao?.dentistaId && !x.decisao?.descartado).length,
      divergencias: divergencias.length, divPend,
    };
    tot.pendencias = tot.divPend + tot.neAbertos + tot.conflitos;
    const abertas = pendentes.filter(l => l.pend), aprov = pendentes.filter(l => l.aprovado);
    tot.pendentesQtd = abertas.length; tot.pendentesValor = abertas.reduce((a, l) => a + l.valor, 0);
    tot.aprovadasQtd = aprov.length; tot.aprovadasValor = aprov.reduce((a, l) => a + l.valor, 0);
    return { linhas: mov, resumo: resumoArr, naoEncontradas, divergencias, extras, pendentes, tot };
  }

  return { norm, nameKey, guiaKey, toNum, toDate, dayKey, findHeaderRow, mapColumns, analyzeSystemWorkbook, buildSystemRows, analyzeDentistWorkbook, limpaTexto, matchDentist, guessDentistFromFilename, guessDentistFromText, reconcile };
})();
if (typeof module !== 'undefined') module.exports = Core;
