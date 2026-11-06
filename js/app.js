/* =====================================================================
   Conciliação de Produção Odontológica — app (Supabase + GitHub Pages)
   ===================================================================== */
(() => {
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const brl = v => (+v || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const n0 = v => (+v || 0).toLocaleString('pt-BR');
const dtf = v => v ? new Date(v).toLocaleString('pt-BR', { day: '2-digit', month: '2-digit', year: '2-digit', hour: '2-digit', minute: '2-digit' }) : '—';
const GENERIC = /cl[ií]nico geral|sem dentista|n[aã]o informado|avalia[cç][aã]o|^$/i;
const FORMAS = { percentual: 'Percentual', diaria: 'Diária fixa', fixo: 'Fixo mensal' };
const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
const compLabel = c => { if (!c) return ''; const [y, m] = String(c).split('-'); return `${MESES[+m - 1]}/${y}`; };

const CFG = window.APP_CONFIG || {};
const S = {
  user: null, perfil: null,
  cad: { dentistas: [], especialidades: [], regras: [] },
  conc: null,          // linha da tabela conciliacoes (ou {competencia, novo:true})
  ro: false,           // competência fechada → somente leitura
  sys: null, denFiles: [],
  situacoes: [], sitBase: new Set(),
  overrides: {}, overridesAuto: {}, aprovadas: {}, pagamentos: [], diarias: {}, resolucoes: {}, lancamentos: {}, splits: {}, pendView: { dent: '', q: '', page: 0, f: 'abertas' },
  result: null, view: { orig: 'todas', dent: '', q: '', page: 0 }, divView: { f: 'pend', page: 0 },
  editId: null, audFiltro: '',
};

/* ===================== Supabase ===================== */
if (!CFG.supabaseUrl || !CFG.supabaseAnonKey || /SEU-PROJETO|COLE_AQUI/.test(CFG.supabaseUrl + CFG.supabaseAnonKey)) {
  document.body.innerHTML = `<div style="max-width:560px;margin:80px auto;font-family:Manrope,sans-serif;padding:24px;background:#fff;border-radius:14px;border:1px solid #dfe8e6">
    <h2 style="margin-top:0">Falta configurar o Supabase</h2><p>Edite o arquivo <b>js/config.js</b> com a URL e a chave <i>anon public</i> do seu projeto (Supabase → Project Settings → API) e publique de novo.</p></div>`;
  return;
}
const sb = window.supabase.createClient(CFG.supabaseUrl, CFG.supabaseAnonKey, { auth: { persistSession: true, autoRefreshToken: true } });
const BUCKET = 'planilhas';
async function q(promise, msg) {
  const { data, error } = await promise;
  if (error) { console.error(error); throw new Error((msg ? msg + ': ' : '') + (error.message || 'erro')); }
  return data;
}
function setSave(t, err) { const el = $('#saveState'); el.textContent = t; el.style.color = err ? '#ffb4ab' : ''; }
const fail = e => { toast(e.message || String(e)); setSave('erro ao salvar', true); };

/* ===================== UI utils ===================== */
function toast(t) { const el = $('#toast'); el.textContent = t; el.classList.add('on'); clearTimeout(el._t); el._t = setTimeout(() => el.classList.remove('on'), 3200); }
function dialog(title, text, { input = false, placeholder = '', okText = 'Confirmar', required = false } = {}) {
  return new Promise(res => {
    $('#mTitle').textContent = title; $('#mText').textContent = text; $('#mYes').textContent = okText;
    const inp = $('#mInput'); inp.classList.toggle('hide', !input); inp.value = ''; inp.placeholder = placeholder;
    $('#modal').classList.add('on'); if (input) setTimeout(() => inp.focus(), 50);
    const done = v => { $('#modal').classList.remove('on'); $('#mYes').onclick = $('#mNo').onclick = null; res(v); };
    $('#mYes').onclick = () => { if (input && required && !inp.value.trim()) { inp.focus(); return; } done(input ? inp.value.trim() : true); };
    $('#mNo').onclick = () => done(input ? null : false);
  });
}
const TABS = ['conc', 'res', 'pend', 'dent', 'esp', 'hist'];
function showTab(t) {
  document.querySelectorAll('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === t));
  TABS.forEach(x => $('#tab-' + x).classList.toggle('hide', x !== t));
  if (t === 'hist') loadHistorico();
  window.scrollTo(0, 0);
}
document.querySelectorAll('#tabs button').forEach(b => b.onclick = () => showTab(b.dataset.tab));
function setupDrop(dropEl, inputEl, onFiles) {
  inputEl.onchange = () => { const f = [...inputEl.files]; inputEl.value = ''; if (!S.ro) onFiles(f); };
  ['dragenter', 'dragover'].forEach(ev => dropEl.addEventListener(ev, e => { e.preventDefault(); dropEl.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => dropEl.addEventListener(ev, e => { e.preventDefault(); dropEl.classList.remove('over'); }));
  dropEl.addEventListener('drop', e => { if (!S.ro) onFiles([...e.dataTransfer.files]); });
}
async function parseBook(buf) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  return wb.SheetNames.map(n => ({ name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }));
}
const safeName = n => n.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w.\-]+/g, '_').slice(-120);
const espById = id => S.cad.especialidades.find(e => e.id === id);
const dentById = id => S.cad.dentistas.find(d => d.id === id);
const ativos = () => S.cad.dentistas.filter(d => d.ativo).sort((a, b) => a.nome.localeCompare(b.nome));
const dentOpts = (sel, blank = '— selecione —') => `<option value="">${blank}</option>` + ativos().map(d => `<option value="${d.id}" ${d.id === sel ? 'selected' : ''}>${esc(d.nome)}</option>`).join('')
  + (sel && dentById(sel) && !dentById(sel).ativo ? `<option value="${sel}" selected>${esc(dentById(sel).nome)} (inativo)</option>` : '');

/* ===================== Login ===================== */
$('#loginClinica').textContent = CFG.clinica ? `${CFG.clinica} · comissões` : 'Odontológica · comissões';
$('#loginForm').onsubmit = async e => {
  e.preventDefault();
  $('#lgBtn').disabled = true; $('#lgMsg').textContent = 'Entrando…';
  const { error } = await sb.auth.signInWithPassword({ email: $('#lgEmail').value.trim(), password: $('#lgSenha').value });
  $('#lgBtn').disabled = false;
  if (error) { $('#lgMsg').textContent = /invalid/i.test(error.message) ? 'E-mail ou senha inválidos.' : error.message; return; }
  $('#lgMsg').textContent = '';
};
$('#btnSair').onclick = async () => { await sb.auth.signOut(); location.reload(); };
/* ----- Solicitar acesso / Esqueci minha senha ----- */
let reqTipo = 'acesso';
function abrirReq(tipo) {
  reqTipo = tipo;
  $('#loginForm').classList.add('hide'); $('#reqForm').classList.remove('hide');
  $('#reqTitle').textContent = tipo === 'acesso' ? 'Solicitar acesso' : 'Esqueci minha senha';
  $('#reqHint').textContent = tipo === 'acesso' ? 'Preencha seus dados. O administrador vai analisar e enviar o acesso para o seu e-mail.' : 'Informe o e-mail cadastrado. O administrador vai redefinir sua senha e te avisar por e-mail.';
  $('#reqNomeBox').classList.toggle('hide', tipo !== 'acesso'); $('#reqNome').required = tipo === 'acesso';
  $('#reqMsgBox').querySelector('span').textContent = tipo === 'acesso' ? 'Função / observação' : 'Observação (opcional)';
  $('#reqEmail').value = $('#lgEmail').value || ''; $('#reqOut').textContent = ''; $('#reqBtn').disabled = false;
  setTimeout(() => (tipo === 'acesso' ? $('#reqNome') : $('#reqEmail')).focus(), 50);
}
$('#lnkAcesso').onclick = e => { e.preventDefault(); abrirReq('acesso'); };
$('#lnkSenha').onclick = e => { e.preventDefault(); abrirReq('senha'); };
$('#lnkVoltar').onclick = e => { e.preventDefault(); $('#reqForm').classList.add('hide'); $('#loginForm').classList.remove('hide'); };
$('#reqForm').onsubmit = async e => {
  e.preventDefault();
  const out = $('#reqOut');
  const dados = { tipo: reqTipo, nome: reqTipo === 'acesso' ? $('#reqNome').value.trim() : null, email: $('#reqEmail').value.trim().toLowerCase(), mensagem: $('#reqMsg').value.trim() || null };
  if (reqTipo === 'acesso' && !dados.nome) { out.style.color = 'var(--bad)'; out.textContent = 'Informe seu nome.'; return; }
  $('#reqBtn').disabled = true; out.style.color = 'var(--ink2)'; out.textContent = 'Enviando…';
  const { error } = await sb.from('solicitacoes').insert(dados);
  if (error) { $('#reqBtn').disabled = false; out.style.color = 'var(--bad)'; out.textContent = /Muitas/.test(error.message) ? error.message : /check constraint|email/i.test(error.message) ? 'Confira o e-mail digitado.' : 'Não foi possível enviar agora. Tente novamente.'; return; }
  out.style.color = 'var(--ok)'; out.textContent = '✓ Solicitação enviada! Você receberá o retorno por e-mail.';
  $('#reqNome').value = ''; $('#reqMsg').value = '';
};

/* ----- Painel de solicitações (admin) ----- */
let solFiltro = 'pendente';
async function contarSolicitacoes() {
  if (S.perfil?.papel !== 'admin') return;
  const { count } = await sb.from('solicitacoes').select('id', { count: 'exact', head: true }).eq('status', 'pendente');
  $('#btnSol').classList.remove('hide'); $('#solCnt').textContent = count ? `(${count})` : '';
}
$('#btnSol').onclick = () => { $('#solModal').classList.add('on'); carregarSolicitacoes(); };
$('#solFechar').onclick = () => $('#solModal').classList.remove('on');
document.querySelectorAll('#segSol button').forEach(b => b.onclick = () => { document.querySelectorAll('#segSol button').forEach(x => x.classList.toggle('on', x === b)); solFiltro = b.dataset.s; carregarSolicitacoes(); });
const senhaTemp = () => { const c = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789'; let s = ''; const a = crypto.getRandomValues(new Uint32Array(10)); a.forEach(n => s += c[n % c.length]); return s; };
async function carregarSolicitacoes() {
  const el = $('#solList'); el.innerHTML = '<p class="muted">Carregando…</p>';
  try {
    const rows = await q(sb.from('solicitacoes').select('*').eq('status', solFiltro).order('criado_em', { ascending: false }).limit(100), 'Solicitações');
    if (!rows.length) { el.innerHTML = `<p class="empty">${solFiltro === 'pendente' ? '✓ Nenhuma solicitação pendente.' : 'Nada aqui.'}</p>`; return; }
    el.innerHTML = rows.map(r => {
      const pw = senhaTemp();
      const sqlSenha = `update auth.users set encrypted_password = extensions.crypt('${pw}', extensions.gen_salt('bf')), updated_at = now()\nwhere email = '${r.email}';`;
      return `<div class="sol"><div class="top"><span class="pill ${r.tipo === 'acesso' ? 'novo' : 'rasc'}">${r.tipo === 'acesso' ? 'Solicitar acesso' : 'Esqueci minha senha'}</span>
        <b>${esc(r.nome || r.email)}</b><span class="muted">${esc(r.email)}</span><span class="sp"></span><span class="muted" style="font-size:12px">${dtf(r.criado_em)}</span></div>
        ${r.mensagem ? `<div style="margin-top:6px;font-size:13px">${esc(r.mensagem)}</div>` : ''}
        ${r.status === 'pendente' ? (r.tipo === 'acesso'
          ? `<div class="muted" style="font-size:12.5px;margin-top:8px">Para liberar: Supabase → <b>Authentication → Users → Add user</b> → e-mail <b>${esc(r.email)}</b>, senha <b class="mono">${pw}</b>, marque <i>Auto Confirm User</i>. Depois envie a senha à pessoa e marque como atendida.</div>`
          : `<div class="muted" style="font-size:12.5px;margin-top:8px">Rode no Supabase → <b>SQL Editor</b> (senha temporária já gerada) e envie a nova senha à pessoa:</div><pre>${esc(sqlSenha)}</pre>`)
          + `<div class="row" style="margin-top:8px;gap:6px"><button class="btn sm" data-copy="${esc(r.tipo === 'acesso' ? `E-mail: ${r.email}\nSenha: ${pw}` : sqlSenha)}">Copiar ${r.tipo === 'acesso' ? 'dados' : 'SQL'}</button><button class="btn sm" data-msg="${esc(`Olá${r.nome ? ', ' + r.nome.split(' ')[0] : ''}! ${r.tipo === 'acesso' ? 'Seu acesso' : 'Sua senha'} ao sistema de Conciliação foi ${r.tipo === 'acesso' ? 'liberado' : 'redefinida'}.\n\nEndereço: ${location.href.split('#')[0]}\nE-mail: ${r.email}\nSenha: ${pw}\n\nRecomendamos guardar esta senha em local seguro.`)}">Copiar mensagem p/ enviar</button><span class="sp"></span>
            <button class="btn sm ghost danger" data-st="recusada" data-id="${r.id}">Recusar</button><button class="btn sm pri" data-st="atendida" data-id="${r.id}">Marcar como atendida</button></div>`
        : `<div class="muted" style="font-size:12px;margin-top:6px">${r.status === 'atendida' ? '✓ Atendida' : '✕ Recusada'} em ${dtf(r.atualizado_em)}${r.observacao ? ' · ' + esc(r.observacao) : ''}</div>`}</div>`;
    }).join('');
    el.querySelectorAll('[data-copy],[data-msg]').forEach(b => b.onclick = async () => { try { await navigator.clipboard.writeText(b.dataset.copy ?? b.dataset.msg); toast('Copiado'); } catch { toast('Não consegui copiar; selecione o texto manualmente.'); } });
    el.querySelectorAll('[data-st]').forEach(b => b.onclick = async () => {
      const obs = await dialog(b.dataset.st === 'atendida' ? 'Marcar como atendida?' : 'Recusar solicitação?', 'Fica registrado na auditoria.', { input: true, placeholder: 'Observação (opcional)', okText: 'Confirmar' });
      if (obs === null) return;
      try { await q(sb.from('solicitacoes').update({ status: b.dataset.st, observacao: obs || null }).eq('id', b.dataset.id), 'Atualizar'); toast('Solicitação atualizada'); carregarSolicitacoes(); contarSolicitacoes(); } catch (e) { fail(e); }
    });
  } catch (e) { el.innerHTML = `<p class="muted">${esc(e.message)}</p>`; }
}

let started = false;
sb.auth.onAuthStateChange((ev, session) => {
  if (session?.user && !started) { started = true; S.user = session.user; boot(); }
  if (!session) { $('#login').classList.remove('hide'); }
});
(async () => { const { data } = await sb.auth.getSession(); if (!data.session) $('#login').classList.remove('hide'); })();

async function boot() {
  $('#login').classList.add('hide');
  try {
    S.perfil = await q(sb.from('perfis').select('*').eq('id', S.user.id).maybeSingle());
    $('#userChip').innerHTML = `<b>${esc(S.perfil?.nome || S.user.email)}</b> · ${S.perfil?.papel === 'admin' ? 'admin' : 'operador'}`;
    await loadCadastro();
    contarSolicitacoes();
    const d = new Date(); d.setMonth(d.getMonth() - 1);
    $('#compMes').value = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
    await abrirCompetencia($('#compMes').value);
  } catch (e) { toast(e.message); }
}

/* ===================== Cadastro (Supabase) ===================== */
async function loadCadastro() {
  const [esp, den] = await Promise.all([
    q(sb.from('especialidades').select('*').order('criado_em'), 'Especialidades'),
    q(sb.from('dentistas').select('*').order('nome'), 'Dentistas'),
  ]);
  try { S.cad.regras = (await q(sb.from('regras_procedimento').select('*').order('criado_em'), 'Regras')).map(r => ({ id: r.id, termo: r.termo, dentistaId: r.dentista_id, ativo: r.ativo })); }
  catch (e) { S.cad.regras = []; console.warn('Tabela regras_procedimento ausente — rode o schema.sql', e); }
  S.cad.especialidades = esp.map(e => ({ id: e.id, nome: e.nome, forma: e.forma, valor: Number(e.valor), ativo: e.ativo }));
  S.cad.dentistas = den.map(d => ({ id: d.id, nome: d.nome, aliases: d.aliases || '', especialidadeId: d.especialidade_id, forma: d.forma || null, valor: d.valor == null ? null : Number(d.valor), ativo: d.ativo }));
  renderDentistas(); renderEsp(); renderDen(); renderRegras(); checkReady();
}
async function cadChanged() {
  await loadCadastro();
  if (S.result) run(false);
}

/* ===================== Competência ===================== */
$('#compMes').onchange = () => { if ($('#compMes').value) abrirCompetencia($('#compMes').value); };
function resetEstado() {
  S.sys = null; S.denFiles = []; S.situacoes = []; S.sitBase = new Set();
  S.overrides = {}; S.overridesAuto = {}; S.aprovadas = {}; S.pagamentos = []; S.diarias = {}; S.resolucoes = {}; S.lancamentos = {}; S.splits = {}; S.result = null;
  S.view = { orig: 'todas', dent: '', q: '', page: 0 }; S.divView = { f: 'pend', page: 0 };
  $('#tolDiv').value = 1;
}
async function abrirCompetencia(mes) {
  const comp = mes + '-01';
  try { await flushNow(); } catch {}
  document.body.classList.add('loading'); setSave('carregando…');
  try {
    resetEstado();
    const c = await q(sb.from('conciliacoes').select('*').eq('competencia', comp).maybeSingle(), 'Conciliação');
    S.conc = c || { competencia: comp, novo: true };
    if (c) {
      const regras = c.regras || {}, dec = c.decisoes || {};
      S.overrides = dec.overrides || {}; S.diarias = dec.diarias || {}; S.resolucoes = dec.resolucoes || {}; S.lancamentos = dec.lancamentos || {}; S.splits = dec.splits || {}; S.overridesAuto = dec.overridesAuto || {}; S.aprovadas = dec.aprovadas || {}; S.pagamentos = dec.pagamentos || [];
      if (regras.tolDiv != null) $('#tolDiv').value = regras.tolDiv;
      const arqs = await q(sb.from('conciliacao_arquivos').select('*').eq('conciliacao_id', c.id).eq('ativo', true).order('enviado_em'), 'Arquivos');
      const sysA = arqs.filter(a => a.tipo === 'sistema').pop();
      if (sysA) {
        const sheets = await baixarPlanilha(sysA.storage_path);
        montarSistema(sysA.nome, sheets, sysA.id, regras);
      }
      for (const a of arqs.filter(x => x.tipo === 'dentista')) {
        const sheets = await baixarPlanilha(a.storage_path);
        const df = { id: a.id, arquivoId: a.id, name: a.nome, sheets, dentistaId: a.dentista_id || '' };
        analyzeDen(df, false); S.denFiles.push(df);
      }
    }
    S.ro = S.conc.status === 'fechada';
    if (S.sys) run(false); else renderResult();
    renderAllConc(); renderCompBar();
    setSave(c ? 'tudo salvo' : '');
  } catch (e) { toast(e.message); setSave('erro', true); }
  document.body.classList.remove('loading');
}
async function baixarPlanilha(path) {
  const blob = await q(sb.storage.from(BUCKET).download(path), 'Baixar planilha');
  return parseBook(await blob.arrayBuffer());
}
async function garantirConciliacao() {
  if (S.conc && !S.conc.novo) return S.conc;
  const row = await q(sb.from('conciliacoes').insert({ competencia: S.conc.competencia, regras: regrasAtuaisFull(), decisoes: decisoesAtuais() }).select().single(), 'Criar competência');
  S.conc = row; renderCompBar(); return row;
}
async function enviarArquivo(file, tipo, dentistaId) {
  const c = await garantirConciliacao();
  const path = `${c.id}/${Date.now()}_${tipo}_${safeName(file.name)}`;
  await q(sb.storage.from(BUCKET).upload(path, file, { upsert: false, contentType: file.type || 'application/octet-stream' }), 'Enviar arquivo');
  return q(sb.from('conciliacao_arquivos').insert({ conciliacao_id: c.id, tipo, nome: file.name, dentista_id: dentistaId || null, storage_path: path, tamanho: file.size }).select().single(), 'Registrar arquivo');
}
const regrasAtuais = () => ({ sitBase: [...S.sitBase], tolDiv: Math.max(0, +$('#tolDiv').value || 0) });
const decisoesAtuais = () => ({ overrides: S.overrides, diarias: S.diarias, resolucoes: S.resolucoes, lancamentos: S.lancamentos, splits: S.splits, overridesAuto: S.overridesAuto, aprovadas: S.aprovadas, pagamentos: S.pagamentos });
const totaisAtuais = () => {
  if (!S.result) return {};
  const T = S.result.tot;
  return { producao: +T.producao.toFixed(2), base: +T.base.toFixed(2), comissao: +T.comissao.toFixed(2), ajustes: +T.ajustes.toFixed(2), linhas: T.linhas, guias: T.guias, pendencias: T.pendencias, dentistas: S.result.resumo.length, arquivosDentistas: S.denFiles.filter(f => f.dentistaId).length };
};
const resumoAtual = () => (S.result?.resumo || []).map(r => ({ ...r, producao: +r.producao.toFixed(2), base: +r.base.toFixed(2), ajustes: +r.ajustes.toFixed(2), comissao: +r.comissao.toFixed(2) }));

let saveTimer = null, saving = false, again = false;
function salvar() {
  if (S.ro || !S.conc || S.conc.novo) return;
  setSave('salvando…');
  clearTimeout(saveTimer); saveTimer = setTimeout(flush, 900);
}
async function flush() {
  if (saving) { again = true; return; }
  saving = true;
  try {
    const row = await q(sb.from('conciliacoes').update({ regras: regrasAtuaisFull(), decisoes: decisoesAtuais(), totais: totaisAtuais(), arquivo_sistema: S.sys?.fileName || null }).eq('id', S.conc.id).select().single(), 'Salvar');
    S.conc = row; setSave('✓ tudo salvo');
  } catch (e) { fail(e); }
  saving = false;
  if (again) { again = false; flush(); }
}
async function flushNow() { clearTimeout(saveTimer); if (S.conc && !S.conc.novo && !S.ro) await flush(); }

function renderCompBar() {
  const c = S.conc; if (!c) return;
  const adm = S.perfil?.papel === 'admin';
  const st = c.novo ? '<span class="pill novo">Nova — suba a planilha do sistema</span>' : c.status === 'fechada' ? `<span class="pill fech">✓ Fechada</span>` : '<span class="pill rasc">● Rascunho</span>';
  $('#compStatus').innerHTML = `${st}${c.novo ? '' : `<span class="muted" style="font-size:12px">versão ${c.versao} · atualizado ${dtf(c.atualizado_em)}</span>`}`;
  $('#btnVersao').classList.toggle('hide', !!c.novo || S.ro);
  $('#btnFechar').classList.toggle('hide', !!c.novo || S.ro);
  $('#btnReabrir').classList.toggle('hide', !S.ro || !adm);
  $('#btnFechar').disabled = !S.result;
  const b = $('#roBanner');
  b.classList.toggle('hide', !S.ro);
  if (S.ro) b.innerHTML = `🔒 <b>${compLabel(c.competencia)}</b> está fechada desde ${dtf(c.fechada_em)}. Tudo fica somente leitura. ${adm ? 'Para corrigir, use <b>Reabrir</b> (o motivo fica registrado).' : 'Para corrigir, peça a um administrador para reabrir.'}`;
  document.body.classList.toggle('ro', S.ro);
}
$('#btnVersao').onclick = async () => {
  if (!S.result) { toast('Faça a conciliação primeiro.'); return; }
  const obs = await dialog('Salvar versão', `Grava uma foto de ${compLabel(S.conc.competencia)} no histórico (decisões, totais e comissões).`, { input: true, placeholder: 'Observação (opcional)', okText: 'Salvar versão' });
  if (obs === null) return;
  try { await flushNow(); const v = await q(sb.rpc('salvar_versao', { p_id: S.conc.id, p_totais: totaisAtuais(), p_resumo: resumoAtual(), p_observacao: obs || null }), 'Salvar versão'); S.conc.versao = v; renderCompBar(); toast(`Versão ${v} gravada`); }
  catch (e) { fail(e); }
};
$('#btnFechar').onclick = async () => {
  if (!S.result) { toast('Faça a conciliação primeiro.'); return; }
  const T = S.result.tot;
  const aviso = T.pendencias ? `\n\n⚠ Ainda há ${T.pendencias} pendência(s) sem decisão (${T.divPend} divergências, ${T.neAbertos} não localizados, ${T.conflitos} conflitos). Elas ficarão fora da comissão.` : '';
  const obs = await dialog(`Fechar ${compLabel(S.conc.competencia)}?`, `Comissão total: ${brl(T.comissao)} para ${S.result.resumo.length} dentista(s).\nDepois de fechada, a competência fica travada e as comissões vão para o histórico.${aviso}`, { input: true, placeholder: 'Observação do fechamento (opcional)', okText: 'Fechar competência' });
  if (obs === null) return;
  try {
    await flushNow();
    const v = await q(sb.rpc('fechar_conciliacao', { p_id: S.conc.id, p_totais: totaisAtuais(), p_resumo: resumoAtual(), p_detalhe: detalheCompacto(), p_observacao: obs || null }), 'Fechar');
    clearTimeout(saveTimer); S.ro = true;
    toast(`Competência fechada (versão ${v})`);
    await abrirCompetencia($('#compMes').value);
  } catch (e) { fail(e); }
};
$('#btnReabrir').onclick = async () => {
  const motivo = await dialog(`Reabrir ${compLabel(S.conc.competencia)}?`, 'O fechamento atual continua guardado no histórico. Informe o motivo da reabertura.', { input: true, placeholder: 'Motivo (obrigatório)', okText: 'Reabrir', required: true });
  if (!motivo) return;
  try { await q(sb.rpc('reabrir_conciliacao', { p_id: S.conc.id, p_motivo: motivo }), 'Reabrir'); toast('Competência reaberta'); await abrirCompetencia($('#compMes').value); }
  catch (e) { fail(e); }
};
function detalheCompacto() {
  return S.result.linhas.map(l => ({ g: l.guiaRaw, a: l.atendimento, d: l.dia, p: l.paciente, pr: l.procedimento, s: l.situacao, ds: l.dentistaSistema, e: l.executorId, en: dentById(l.executorId)?.nome || null, o: l.origem, c: l.comissionavel, v: l.valor, ob: l.obs || undefined, sp: l.split || undefined, pg: linhaPgto(l.id) || undefined }));
}

/* ===================== 1. sistema ===================== */
setupDrop($('#dropSys'), $('#fileSys'), async files => {
  const f = files[0]; if (!f) return;
  try {
    const sheets = await parseBook(await f.arrayBuffer());
    const analysis = Core.analyzeSystemWorkbook(sheets);
    if (!analysis[0] || analysis[0].score < 10) { toast('Não encontrei uma coluna de guia nessa planilha.'); return; }
    const temDecisoes = Object.keys(S.overrides).length + Object.keys(S.resolucoes).length + Object.keys(S.lancamentos).length;
    if (S.sys && !await dialog('Substituir a planilha do sistema?', `A atual (${S.sys.fileName}) fica guardada no histórico.${temDecisoes ? `\nAs ${temDecisoes} decisão(ões) desta competência serão zeradas, pois dependem das linhas da planilha.` : ''}`, { okText: 'Substituir' })) return;
    setSave('enviando arquivo…');
    const antigo = S.sys?.arquivoId;
    const arq = await enviarArquivo(f, 'sistema');
    if (antigo) await q(sb.from('conciliacao_arquivos').update({ ativo: false }).eq('id', antigo), 'Arquivo anterior');
    S.overrides = {}; S.overridesAuto = {}; S.aprovadas = {}; S.pagamentos = []; S.diarias = {}; S.resolucoes = {}; S.lancamentos = {}; S.splits = {}; S.result = null;
    montarSistema(f.name, sheets, arq.id, null);
    S.denFiles.forEach(df => analyzeDen(df, false));
    renderAllConc(); run(false); salvar();
    toast('Produção do sistema carregada e guardada');
  } catch (e) { fail(e); }
});
function montarSistema(fileName, sheets, arquivoId, regras) {
  const analysis = Core.analyzeSystemWorkbook(sheets);
  const pref = regras?.aba && analysis.find(a => a.name === regras.aba) || analysis[0];
  S.sys = { fileName, sheets, analysis, arquivoId, sel: pref.name, map: { ...(regras?.map || pref.map) }, headerIndex: regras?.headerIndex ?? pref.headerIndex };
  buildSys(regras?.sitBase);
}
function buildSys(sitSalvas) {
  const sh = S.sys.sheets.find(s => s.name === S.sys.sel);
  S.sys.rows = Core.buildSystemRows(sh.rows, S.sys.headerIndex, S.sys.map);
  S.sys.guiaSet = new Set(S.sys.rows.map(r => r.guia));
  S.sys.atSet = new Set(S.sys.rows.map(r => r.atendimento).filter(Boolean));
  const ehPend = r => /pend/i.test(r.situacao || '') || /pend/i.test(r.status || '');
  S.sys.pendRows = S.sys.rows.filter(ehPend);
  S.situacoes = [...new Set(S.sys.rows.filter(r => !ehPend(r)).map(r => r.situacao))];
  S.sitBase = Array.isArray(sitSalvas) ? new Set(sitSalvas.filter(x => S.situacoes.includes(x))) : new Set(S.situacoes.filter(s => !/pend|glos|negad|cancel|recus/i.test(s)));
  renderSit();
}
const regrasAtuaisFull = () => ({ ...regrasAtuais(), aba: S.sys?.sel, map: S.sys?.map, headerIndex: S.sys?.headerIndex });
function renderSys() {
  const el = $('#sysInfo');
  if (!S.sys) { el.innerHTML = ''; return; }
  const rows = S.sys.rows.filter(r => !S.sys.pendRows.includes(r));
  const ds = S.sys.rows.map(r => r.data).filter(Boolean).sort((a, b) => a - b);
  const per = ds.length ? `${ds[0].toLocaleDateString('pt-BR')} a ${ds[ds.length - 1].toLocaleDateString('pt-BR')}` : '—';
  const sh = S.sys.sheets.find(s => s.name === S.sys.sel);
  const header = sh.rows[S.sys.headerIndex] || [];
  const F = { guia: 'Nº da guia *', dentista: 'Dentista', valor: 'Valor *', data: 'Data', situacao: 'Situação', procedimento: 'Procedimento', paciente: 'Paciente', convenio: 'Convênio', atendimento: 'ID atendimento', obs: 'Motivo / observação' };
  const colOpts = k => `<option value="">— não usar —</option>` + header.map((h, i) => `<option value="${i}" ${S.sys.map[k] === i ? 'selected' : ''}>${esc(h || 'Coluna ' + (i + 1))}</option>`).join('');
  el.innerHTML = `
    <div class="row" style="margin-top:14px"><span class="chip ok" title="Guardado no histórico">✓ ${esc(S.sys.fileName)}</span><span class="sp"></span>
      <label class="fld" style="width:230px"><span>Aba usada</span><select class="inp" id="sysSheet">${S.sys.analysis.filter(a => a.score > 0).map(a => `<option ${a.name === S.sys.sel ? 'selected' : ''}>${esc(a.name)}</option>`).join('')}</select></label></div>
    <div class="meta"><div><span>Linhas no movimento</span><b>${n0(rows.length)}</b></div><div><span>Guias</span><b>${n0(new Set(rows.map(r => r.guia)).size)}</b></div><div><span>Produção</span><b>${brl(rows.reduce((a, r) => a + r.valor, 0))}</b></div><div><span>Período</span><b style="font-size:12px">${per}</b></div></div>
    ${S.sys.pendRows.length ? `<div class="notice warn" style="margin-top:10px">${n0(S.sys.pendRows.length)} linha(s) com status <b>Pendente</b> (${brl(S.sys.pendRows.reduce((a, r) => a + r.valor, 0))}) ficam <b>fora do movimento</b> até o pagamento ser aprovado na aba <b>Guias pendentes</b>.</div>` : ''}
    <div class="chips" style="margin-top:10px">${Object.keys(F).map(k => `<span class="chip ${S.sys.map[k] !== undefined ? 'ok' : 'miss'}">${S.sys.map[k] !== undefined ? '✓' : '✕'} ${F[k].replace(' *', '')}</span>`).join('')}</div>
    <details class="map"><summary>Ajustar colunas manualmente</summary><div class="mapgrid">${Object.keys(F).map(k => `<label class="fld"><span>${F[k]}</span><select class="inp" data-map="${k}">${colOpts(k)}</select></label>`).join('')}</div></details>`;
  const afterChange = () => { S.denFiles.forEach(df => analyzeDen(df, false)); renderAllConc(); if (S.result) run(false); salvarRegras(); };
  $('#sysSheet').onchange = e => { const a = S.sys.analysis.find(x => x.name === e.target.value); S.sys.sel = a.name; S.sys.map = { ...a.map }; S.sys.headerIndex = a.headerIndex; buildSys(); afterChange(); };
  el.querySelectorAll('[data-map]').forEach(s => s.onchange = () => { const k = s.dataset.map; if (s.value === '') delete S.sys.map[k]; else S.sys.map[k] = +s.value; buildSys([...S.sitBase]); afterChange(); });
}
function salvarRegras() { salvar(); }
function renderSit() {
  const box = $('#sitBox');
  if (!S.situacoes.length) { box.innerHTML = '<span class="muted">Carregue a planilha do sistema.</span>'; return; }
  box.innerHTML = S.situacoes.map(s => {
    const n = S.sys.rows.filter(r => r.situacao === s && !S.sys.pendRows.includes(r)); const v = n.reduce((a, r) => a + r.valor, 0);
    return `<label class="chk"><input type="checkbox" data-sit="${esc(s)}" ${S.sitBase.has(s) ? 'checked' : ''}>${esc(s)} <span class="muted mono" style="font-size:11px">${n.length} · ${brl(v)}</span></label>`;
  }).join('');
  if (S.sys.pendRows.length) box.insertAdjacentHTML('beforeend', `<span class="chip" title="Status Pendente nunca entra no movimento">🔒 Pendentes: fora do movimento · ${n0(S.sys.pendRows.length)} · ${brl(S.sys.pendRows.reduce((a, r) => a + r.valor, 0))}</span>`);
  box.querySelectorAll('[data-sit]').forEach(c => c.onchange = () => { c.checked ? S.sitBase.add(c.dataset.sit) : S.sitBase.delete(c.dataset.sit); if (S.result) run(false); salvar(); });
  applyRO();
}

/* ===================== 2. planilhas dos dentistas ===================== */
setupDrop($('#dropDen'), $('#fileDen'), async files => {
  let ok = 0;
  setSave('enviando arquivos…');
  for (const f of files) {
    try {
      const sheets = await parseBook(await f.arrayBuffer());
      const df = { name: f.name, sheets, dentistaId: Core.guessDentistFromFilename(f.name, ativos())?.id || '' };
      analyzeDen(df, true);
      const arq = await enviarArquivo(f, 'dentista', df.dentistaId);
      df.id = df.arquivoId = arq.id;
      S.denFiles.push(df); ok++;
    } catch (e) { fail(new Error(`${f.name}: ${e.message}`)); }
  }
  if (ok) toast(`${ok} planilha(s) de dentista carregada(s) e guardada(s)`);
  renderDen(); checkReady(); if (S.result) run(false); salvar();
});
function analyzeDen(df, guessDent) {
  const a = Core.analyzeDentistWorkbook(df.sheets, S.sys?.guiaSet || new Set(), S.sys?.atSet || new Set());
  Object.assign(df, { guias: a.guias || [], hits: a.hits || 0, col: a.col, header: a.header, sheet: a.sheet, itens: a.itens || [], atendExtra: a.atendExtra || [], atHeader: a.atHeader, temValor: a.temValor });
  for (const s of df.sheets) {
    const h = Core.findHeaderRow(s.rows); if (h.index < 0) continue;
    const ci = (s.rows[h.index] || []).findIndex(c => /dentista|profissional/i.test(String(c ?? '')));
    if (ci < 0) continue;
    const names = [...new Set(s.rows.slice(h.index + 1).map(r => String(r?.[ci] ?? '').trim()).filter(Boolean))];
    if (names.length === 1) {
      df.nomeInformado = names[0];
      if (guessDent && !df.dentistaId) { const d = Core.matchDentist(names[0], ativos()) || Core.guessDentistFromText(names[0], ativos()); if (d) df.dentistaId = d.id; }
      break;
    }
  }
}
function renderDen() {
  const el = $('#denList');
  if (!S.denFiles.length) { el.innerHTML = ''; return; }
  el.innerHTML = `<table class="t" style="margin-top:12px"><colgroup><col style="width:34%"><col style="width:20%"><col style="width:36%"><col style="width:10%"></colgroup>
    <thead><tr><th>Arquivo</th><th class="num">Guias · cruzadas</th><th>Dentista</th><th></th></tr></thead><tbody>
    ${S.denFiles.map(f => {
      const pct = f.guias.length ? Math.round(f.hits / f.guias.length * 100) : 0;
      return `<tr><td title="${esc(f.name)}"><b>${esc(f.name)}</b><div class="muted" style="font-size:11px;overflow:hidden;text-overflow:ellipsis">${f.sheet ? esc(f.sheet) + ' · ' + esc(f.header || 'col ' + (f.col + 1)) + (f.atHeader ? ' + ' + esc(f.atHeader) : '') : 'coluna de guia não encontrada'}${f.nomeInformado ? ' · ' + esc(f.nomeInformado) : ''}</div></td>
      <td class="num">${n0(f.guias.length)} · <span style="color:${!S.sys ? 'var(--muted)' : pct >= 70 ? 'var(--ok)' : pct >= 30 ? 'var(--warn)' : 'var(--bad)'}">${S.sys ? n0(f.hits) : '–'}</span>${f.atendExtra?.length ? `<div class="muted" style="font-size:11px">+${f.atendExtra.length} por atend.</div>` : ''}</td>
      <td><select class="inp" data-fd="${f.id}" style="${f.dentistaId ? '' : 'border-color:var(--bad)'}">${dentOpts(f.dentistaId)}</select></td>
      <td style="text-align:right"><button class="btn ghost sm danger" data-rm="${f.id}" title="Remover (fica guardado no histórico)">✕</button></td></tr>`;
    }).join('')}</tbody></table>
    ${ativos().length ? '' : '<div class="notice warn">Cadastre os dentistas na aba <b>Dentistas</b> para vincular cada planilha.</div>'}`;
  el.querySelectorAll('[data-fd]').forEach(s => s.onchange = async () => {
    const f = S.denFiles.find(x => x.id === s.dataset.fd); f.dentistaId = s.value;
    renderDen(); checkReady(); if (S.result) run(false);
    try { await q(sb.from('conciliacao_arquivos').update({ dentista_id: s.value || null }).eq('id', f.arquivoId), 'Vincular dentista'); salvar(); } catch (e) { fail(e); }
  });
  el.querySelectorAll('[data-rm]').forEach(b => b.onclick = async () => {
    const f = S.denFiles.find(x => x.id === b.dataset.rm);
    if (!await dialog('Remover planilha?', `${f.name} sai desta conciliação, mas continua guardada no histórico.`, { okText: 'Remover' })) return;
    try { await q(sb.from('conciliacao_arquivos').update({ ativo: false }).eq('id', f.arquivoId), 'Remover'); } catch (e) { fail(e); return; }
    S.denFiles = S.denFiles.filter(x => x !== f); renderDen(); checkReady(); if (S.result) run(false); salvar();
  });
  applyRO();
}
function checkReady() {
  const btn = $('#btnRun'), msg = $('#readyMsg');
  if (!S.sys) { btn.disabled = true; msg.textContent = 'Falta a planilha do sistema.'; return; }
  const semDent = S.denFiles.filter(f => !f.dentistaId).length;
  btn.disabled = false;
  msg.innerHTML = semDent ? `<span style="color:var(--warn)">⚠ ${semDent} planilha(s) sem dentista vinculado — serão ignoradas.</span>` : `${n0(S.sys.rows.length)} linhas do sistema · ${S.denFiles.length} planilha(s) de dentista`;
}
$('#btnRun').onclick = () => { run(true); salvar(); };
$('#tolDiv').onchange = () => { if (S.result) run(false); salvar(); };

/* ===================== conciliação ===================== */
function run(go) {
  if (!S.sys) return;
  S.result = Core.reconcile({
    systemRows: S.sys.rows, dentistFiles: S.denFiles.filter(f => f.dentistaId), dentistas: S.cad.dentistas, especialidades: S.cad.especialidades,
    situacoesBase: [...S.sitBase], overrides: S.overrides, diariasOverride: S.diarias, resolucoes: S.resolucoes, lancamentos: S.lancamentos, splits: S.splits, regrasProc: S.cad.regras.filter(r => r.ativo && dentById(r.dentistaId)?.ativo), aprovadas: S.aprovadas, tolDiv: Math.max(0, +$('#tolDiv').value || 0),
  });
  renderResult(); renderCompBar();
  if (go) { S.view.page = 0; showTab('res'); }
}
function renderResult() {
  const R = S.result;
  $('#resEmpty').classList.toggle('hide', !!R); $('#resBody').classList.toggle('hide', !R);
  if (!R) return;
  const T = R.tot;
  $('#resSub').textContent = `${compLabel(S.conc?.competencia)}${S.ro ? ' · 🔒 fechada' : ''} · ${S.sys.fileName} · ${S.denFiles.filter(f => f.dentistaId).length} planilha(s) de dentista`;
  const pId = T.linhas ? T.identificadas / T.linhas * 100 : 0;
  const w = v => (T.linhas ? v / T.linhas * 100 : 0) + '%';
  $('#kpis').innerHTML = `
    <div class="kpi"><span>Produção total</span><b>${brl(T.producao)}</b><small>${n0(T.guias)} guias · ${n0(T.linhas)} linhas</small></div>
    <div class="kpi"><span>Base comissionável</span><b>${brl(T.base)}</b><small>${[...S.sitBase].join(', ') || 'todas as situações'}${T.aprovadasQtd ? ` + ${n0(T.aprovadasQtd)} pendência(s) aprovada(s)` : ''}${T.ajustes ? ' · ajustes ' + brl(T.ajustes) : ''}</small></div>
    <div class="kpi hl"><span>Comissão total</span><b>${brl(T.comissao)}</b><small>${R.resumo.length} dentistas</small></div>
    <div class="kpi"><span>Linhas identificadas</span><b>${pId.toFixed(1).replace('.', ',')}%</b>
      <div class="bar2"><i style="width:${w(T.viaDentista)};background:var(--ok)"></i><i style="width:${w(T.viaSistema)};background:var(--info)"></i><i style="width:${w(T.manual)};background:#6b46c1"></i><i style="width:${w(T.conflitos)};background:var(--warn)"></i></div></div>
    <div class="kpi ${T.naoIdent + T.conflitos ? 'bad' : ''}"><span>Sem executor</span><b>${brl(T.naoIdentValor)}</b><small>${n0(T.naoIdent)} linhas · ${n0(T.conflitos)} conflitos</small></div>
    <div class="kpi ${T.pendencias ? 'bad' : ''}" style="cursor:pointer" id="kpiPend"><span>Pendências p/ decidir</span><b>${n0(T.pendencias)}</b><small>${n0(T.divPend)} diverg. · ${n0(T.neAbertos)} não localiz. · ${n0(T.conflitos)} confl.</small></div>`;
  renderResumo(); renderSeg(); renderDet(); renderNE(); renderDiv(); renderPend(); renderAutoResp();
  $('#kpiPend').onclick = () => (T.divPend ? $('#cardDiv') : T.neAbertos ? $('#cardNE') : $('#tblDet')).scrollIntoView({ behavior: 'smooth', block: 'start' });
  applyRO();
}
function renderResumo() {
  const rows = S.result.resumo;
  const tot = k => rows.reduce((a, r) => a + r[k], 0);
  $('#tblResumo').innerHTML = `<colgroup><col style="width:17%"><col style="width:9%"><col style="width:11%"><col style="width:6%"><col style="width:5%"><col style="width:11%"><col style="width:9%"><col style="width:11%"><col style="width:6%"><col style="width:9%"><col style="width:6%"></colgroup>
  <thead><tr><th>Dentista</th><th>Especialidade</th><th>Remuneração</th><th class="num">Guias</th><th class="num">Proc.</th><th class="num">Produção</th><th class="num" title="Ajustes de divergência e lançamentos não localizados atribuídos">Ajustes</th><th class="num">Base</th><th class="num">Diárias</th><th class="num">Comissão</th><th class="num">Pgto</th></tr></thead>
  <tbody>${rows.length ? rows.map(r => `<tr style="cursor:pointer" data-dent="${r.dentistaId}" title="Conferir informado × apurado de ${esc(r.nome)}">
    <td><b style="color:var(--brand);text-decoration:underline;text-underline-offset:3px">${esc(r.nome)}</b></td><td>${esc(r.especialidade)}</td>
    <td>${r.forma === 'percentual' ? r.valorRef + '% da base' : r.forma === 'diaria' ? brl(r.valorRef) + '/dia' : brl(r.valorRef) + ' fixo'}</td>
    <td class="num">${n0(r.guias)}</td><td class="num">${n0(r.procs)}</td><td class="num">${brl(r.producao)}</td><td class="num" style="color:${r.ajustes ? 'var(--ink)' : 'var(--muted)'}">${r.ajustes ? brl(r.ajustes) : '—'}</td><td class="num">${brl(r.base)}</td>
    <td class="num">${r.forma === 'diaria' ? `<input class="inp num" type="number" min="0" step="1" value="${r.diarias}" data-dia="${r.dentistaId}" title="Calculado: ${r.diariasCalc}" style="width:60px;text-align:right">` : `<span class="muted">${r.diariasCalc}</span>`}</td>
    <td class="num"><b>${brl(r.comissao)}</b></td><td class="num">${pgtoTxt(r.dentistaId)}</td></tr>`).join('') : `<tr><td colspan="11" class="empty">Nenhum dentista identificado. Confira o cadastro e os vínculos das planilhas.</td></tr>`}</tbody>
  <tfoot><tr><td>Total</td><td></td><td></td><td class="num">${n0(tot('guias'))}</td><td class="num">${n0(tot('procs'))}</td><td class="num">${brl(tot('producao'))}</td><td class="num">${brl(tot('ajustes'))}</td><td class="num">${brl(tot('base'))}</td><td></td><td class="num">${brl(tot('comissao'))}</td><td></td></tr></tfoot>`;
  $('#tblResumo').querySelectorAll('tr[data-dent]').forEach(tr => tr.onclick = e => {
    if (e.target.tagName === 'INPUT') return;
    abrirConferencia(tr.dataset.dent);
  });
  renderPgInfo();
  $('#tblResumo').querySelectorAll('[data-dia]').forEach(i => i.onchange = () => { S.diarias[i.dataset.dia] = Math.max(0, +i.value || 0); run(false); salvar(); });
}
/* ===================== conferência do dentista ===================== */
const ST_LBL = { ok: '✓ Conciliado', dif: '≠ Diferença', so_sistema: 'Só no sistema', outro: 'Apurado p/ outro / sem executor', pendente: 'Pendente no sistema', nao_localizado: 'Não localizado no sistema' };
const ST_CURTO = { ok: '✓ Conciliado', dif: '≠ Diferença', so_sistema: 'Só no sistema', outro: 'Outro executor', pendente: 'Pendente', nao_localizado: 'Não localizado' };
let dmState = { id: null, f: 'todas' };
function abrirConferencia(id) { dmState = { id, f: 'todas' }; renderConferencia(); $('#dentModal').classList.add('on'); }
$('#dmFechar').onclick = () => $('#dentModal').classList.remove('on');
$('#dmDet').onclick = () => { $('#dentModal').classList.remove('on'); S.view.dent = dmState.id; S.view.orig = 'todas'; S.view.page = 0; showTab('res'); renderSeg(); renderDet(); setTimeout(() => $('#tblDet').scrollIntoView({ behavior: 'smooth', block: 'start' }), 50); };
function renderConferencia() {
  const id = dmState.id, d = dentById(id);
  const C = Core.compararDentista({ dentistaId: id, dentistFiles: S.denFiles.filter(f => f.dentistaId), result: S.result, systemRows: S.sys.rows, tol: Math.max(0, +$('#tolDiv').value || 0) });
  const conc = C.temPlanilha && C.difs + C.outro + C.naoLocalizado + C.soSistema + C.pendente === 0;
  $('#dmTitle').textContent = d?.nome || '—';
  $('#dmStatus').innerHTML = !C.temPlanilha ? '<span class="pill novo">Sem planilha do dentista</span>' : conc ? '<span class="pill fech">✓ Conciliado</span>' : '<span class="pill rasc">≠ Não conciliado</span>';
  $('#dmHint').textContent = C.temPlanilha
    ? 'Comparação por atendimento entre o que o dentista informou na planilha de produção e o que foi apurado no sistema para ele (linhas atribuídas, divisões e ajustes). Guias pendentes em aberto não entram no apurado.'
    : 'Este dentista não enviou planilha de produção nesta competência: tudo que aparece foi apurado pelo sistema (regras, Resp. no sistema ou atribuição manual).';
  const k = (t, v, sub, cls = '') => `<div class="kpi ${cls}"><span>${t}</span><b>${v}</b><small>${sub}</small></div>`;
  $('#dmKpis').innerHTML = k('Informado pelo dentista', brl(C.informado), `${n0(C.rows.filter(r => r.temInf).length)} atendimento(s)`) + k('Apurado no sistema', brl(C.apurado), `${n0(C.rows.filter(r => r.temSis).length)} atendimento(s)`)
    + k('Diferença', brl(C.dif), Math.abs(C.dif) < 0.05 ? 'valores batem' : C.dif > 0 ? 'informou mais do que foi apurado' : 'apurado maior que o informado', Math.abs(C.dif) >= 0.05 ? 'bad' : '')
    + k('Conciliados', `${n0(C.ok)} de ${n0(C.rows.length)}`, `${n0(C.difs)} com diferença · ${n0(C.soSistema)} só no sistema · ${n0(C.outro + C.pendente + C.naoLocalizado)} outros`);
  const tabs = [['todas', 'Todos', C.rows.length], ['ok', ST_LBL.ok, C.ok], ['dif', ST_LBL.dif, C.difs], ['so_sistema', ST_LBL.so_sistema, C.soSistema], ['outro', 'Apurado p/ outro', C.outro], ['pendente', 'Pendente', C.pendente], ['nao_localizado', 'Não localizado', C.naoLocalizado]].filter(t => t[0] === 'todas' || t[2]);
  $('#dmSeg').innerHTML = tabs.map(([kk, t, c]) => `<button class="${dmState.f === kk ? 'on' : ''}" data-dmf="${kk}">${t}<em>${c}</em></button>`).join('');
  $('#dmSeg').querySelectorAll('button').forEach(b => b.onclick = () => { dmState.f = b.dataset.dmf; renderConferencia(); });
  const rows = C.rows.filter(r => dmState.f === 'todas' || r.status === dmState.f);
  $('#dmTbl').innerHTML = `<colgroup><col style="width:17%"><col style="width:13%"><col style="width:32%"><col style="width:9%"><col style="width:9%"><col style="width:8%"><col style="width:12%"></colgroup>
    <thead><tr><th>Atendimento / paciente</th><th>Guia(s)</th><th>Procedimentos (informado · apurado)</th><th class="num">Informado</th><th class="num">Apurado</th><th class="num">Dif.</th><th>Situação</th></tr></thead>
    <tbody>${rows.length ? rows.map(r => `<tr>
      <td class="two"><span class="l"><span class="mono" style="font-size:12px">${r.tipoRef === 'a' ? esc(r.ref) : '—'}</span> · ${esc(r.paciente || '—')}</span><small>${r.data ? new Date(r.data).toLocaleDateString('pt-BR') : ''}</small></td>
      <td class="two"><span class="l mono" style="font-size:12px" title="informada(s): ${esc(r.guiasInf.join(', '))}">${esc(r.guiasInf.join(', ') || '—')}</span><small class="mono" title="no sistema: ${esc(r.guiasSis.join(', '))}">${r.guiasSis.length ? 'sist.: ' + esc(r.guiasSis.join(', ')) : ''}</small></td>
      <td class="two"><span class="l" title="${esc(r.procInf.join('; '))}">${esc(r.procInf.join('; ') || '—')}</span><small title="${esc(r.procSis.join('; '))}">${esc(r.procSis.join('; ') || '—')}</small></td>
      <td class="num">${r.temInf ? brl(r.informado) : '—'}</td><td class="num">${r.temSis ? brl(r.apurado) : '—'}</td>
      <td class="num" style="color:${Math.abs(r.dif) < 0.05 ? 'var(--muted)' : r.dif > 0 ? 'var(--bad)' : 'var(--info)'}">${brl(r.dif)}</td>
      <td><span class="tag st-${r.status}" title="${ST_LBL[r.status]}">${ST_CURTO[r.status]}</span></td></tr>`).join('') : '<tr><td colspan="7" class="empty">Nada neste filtro.</td></tr>'}</tbody>
    <tfoot><tr><td colspan="3">${n0(rows.length)} atendimento(s)</td><td class="num">${brl(rows.reduce((a, r) => a + r.informado, 0))}</td><td class="num">${brl(rows.reduce((a, r) => a + r.apurado, 0))}</td><td class="num">${brl(rows.reduce((a, r) => a + r.dif, 0))}</td><td></td></tr></tfoot>`;
}

/* ===================== enviado para pagamento ===================== */
function linhasDoDentista(id) { return S.result.linhas.filter(l => l.executorId === id || (l.origem === 'split' && (l.split || []).some(p => p.dentistaId === id))); }
function pgtoDe(id) { return S.pagamentos.filter(p => p.dentistas.some(x => x.id === id)).sort((a, b) => b.em.localeCompare(a.em))[0] || null; }
function linhaPgto(lid) { const p = S.pagamentos.filter(pg => pg.linhas.includes(lid)).sort((a, b) => b.em.localeCompare(a.em))[0]; return p ? p.data : null; }
const dataBR = d => d ? d.split('-').reverse().join('/') : '';
function pgtoTxt(id) {
  const p = pgtoDe(id); if (!p) return '<span class="muted">—</span>';
  const ls = linhasDoDentista(id); const pend = ls.filter(l => !linhaPgto(l.id)).length;
  return `<span title="Enviado em ${dataBR(p.data)} por ${esc(p.por)}${pend ? ` · ${pend} linha(s) nova(s) ainda não enviadas` : ''}" style="color:${pend ? 'var(--warn)' : 'var(--ok)'};font-weight:700;font-size:12px">${dataBR(p.data).slice(0, 5)}${pend ? '*' : ''}</span>`;
}
function renderPgInfo() {
  const n = S.pagamentos.length;
  $('#pgInfo').innerHTML = n ? `último envio: <b>${dataBR(S.pagamentos[n - 1].data)}</b> (${n0(S.pagamentos[n - 1].dentistas.length)} dentista(s)) · <a href="#" id="pgUndo" style="color:var(--bad)">desfazer</a>` : '';
  $('#btnPgto').disabled = S.ro || !S.result?.resumo.length;
  const u = $('#pgUndo'); if (u) u.onclick = async e => {
    e.preventDefault(); if (S.ro) return;
    const p = S.pagamentos[S.pagamentos.length - 1];
    if (!await dialog('Desfazer o último envio para pagamento?', `Envio de ${dataBR(p.data)} para ${p.dentistas.map(x => x.nome).join(', ')}. As transações deixam de ter essa data de pagamento.`, { okText: 'Desfazer envio' })) return;
    S.pagamentos.pop(); run(false); salvar(); toast('Envio para pagamento desfeito');
  };
}
$('#btnPgto').onclick = () => {
  if (S.ro || !S.result) return;
  const R = S.result.resumo;
  $('#pgData').value = new Date().toISOString().slice(0, 10);
  $('#pgLista').innerHTML = `<table class="t"><colgroup><col style="width:8%"><col style="width:44%"><col style="width:18%"><col style="width:16%"><col style="width:14%"></colgroup>
    <thead><tr><th></th><th>Dentista</th><th class="num">Comissão</th><th class="num">Linhas</th><th class="num">Já enviado</th></tr></thead><tbody>
    ${R.map(r => { const p = pgtoDe(r.dentistaId); return `<tr><td><input type="checkbox" data-pgd="${r.dentistaId}" ${p ? '' : 'checked'} style="accent-color:var(--brand)"></td><td><b>${esc(r.nome)}</b></td><td class="num">${brl(r.comissao)}</td><td class="num">${n0(linhasDoDentista(r.dentistaId).length)}</td><td class="num">${p ? dataBR(p.data) : '—'}</td></tr>`; }).join('')}</tbody></table>`;
  const upd = () => { const ids = [...document.querySelectorAll('[data-pgd]:checked')].map(x => x.dataset.pgd); $('#pgTotal').textContent = `${ids.length} dentista(s) · comissão ${brl(R.filter(r => ids.includes(r.dentistaId)).reduce((a, r) => a + r.comissao, 0))}`; };
  document.querySelectorAll('[data-pgd]').forEach(x => x.onchange = upd);
  $('#pgTodos').onclick = () => { document.querySelectorAll('[data-pgd]').forEach(x => x.checked = true); upd(); };
  $('#pgNenhum').onclick = () => { document.querySelectorAll('[data-pgd]').forEach(x => x.checked = false); upd(); };
  $('#pgMsg').textContent = ''; upd(); $('#pgModal').classList.add('on');
};
$('#pgCancel').onclick = () => $('#pgModal').classList.remove('on');
$('#pgOk').onclick = () => {
  const data = $('#pgData').value; const ids = [...document.querySelectorAll('[data-pgd]:checked')].map(x => x.dataset.pgd);
  if (!data) { $('#pgMsg').textContent = 'Informe a data do pagamento.'; return; }
  if (!ids.length) { $('#pgMsg').textContent = 'Selecione pelo menos um dentista.'; return; }
  const R = S.result.resumo;
  const linhas = [...new Set(ids.flatMap(id => linhasDoDentista(id).map(l => l.id)))];
  S.pagamentos.push({ id: Date.now().toString(36), data, em: new Date().toISOString(), por: S.user?.email || '',
    dentistas: ids.map(id => { const r = R.find(x => x.dentistaId === id); return { id, nome: r?.nome || dentById(id)?.nome, comissao: +(r?.comissao || 0).toFixed(2), producao: +(r?.producao || 0).toFixed(2) }; }), linhas });
  $('#pgModal').classList.remove('on'); run(false); salvar();
  toast(`${n0(linhas.length)} transação(ões) de ${ids.length} dentista(s) marcadas com pagamento em ${dataBR(data)}`);
};

const ORIG = [['todas', 'Todas'], ['dentista', 'Via guia'], ['atendimento', 'Via atendimento'], ['sistema', 'Via sistema'], ['manual', 'Manual'], ['nao_identificado', 'Sem executor'], ['conflito', 'Conflito'], ['split', 'Dividida'], ['regra', 'Por regra'], ['aprovada', 'Pendência aprovada'], ['alerta', 'Com alerta']];
const ORIG_LBL = { dentista: 'Dentista (guia)', atendimento: 'Dentista (atendimento)', sistema: 'Sistema', manual: 'Manual', nao_identificado: 'Sem executor', conflito: 'Conflito', split: 'Dividida entre dentistas', regra: 'Regra por procedimento' };
function filtered() {
  const { orig, dent, q: qs } = S.view; const qq = Core.norm(qs);
  return S.result.linhas.filter(l =>
    (orig === 'todas' || (orig === 'alerta' ? !!l.alerta : orig === 'aprovada' ? l.aprovado : l.origem === orig)) &&
    (!dent || (dent === '_none' ? !l.executorId : l.executorId === dent)) &&
    (!qq || Core.norm([l.guiaRaw, l.atendimento, l.paciente, l.procedimento, l.dentistaSistema, l.convenio].join(' ')).includes(qq)));
}
function renderSeg() {
  const L = S.result.linhas;
  const c = k => k === 'todas' ? L.length : k === 'alerta' ? L.filter(l => l.alerta).length : k === 'aprovada' ? L.filter(l => l.aprovado).length : L.filter(l => l.origem === k).length;
  $('#segOrig').innerHTML = ORIG.map(([k, t]) => `<button class="${S.view.orig === k ? 'on' : ''}" data-o="${k}">${t}<em>${c(k)}</em></button>`).join('');
  $('#segOrig').querySelectorAll('button').forEach(b => b.onclick = () => { S.view.orig = b.dataset.o; S.view.page = 0; renderSeg(); renderDet(); });
  $('#fDent').innerHTML = `<option value="">Todos os executores</option><option value="_none">— sem executor —</option>` + ativos().map(d => `<option value="${d.id}">${esc(d.nome)}</option>`).join('');
  $('#fDent').value = S.view.dent;
}
$('#fDent').onchange = e => { S.view.dent = e.target.value; S.view.page = 0; renderDet(); };
$('#fBusca').oninput = e => { S.view.q = e.target.value; S.view.page = 0; renderDet(); };
/* ===== Associar pelo Resp. no sistema (resolve conflitos) ===== */
function analisarResp() {
  const out = { aplicar: [], genericos: [], semCadastro: [], fora: [] };
  S.result.linhas.filter(l => l.origem === 'conflito' && !l.executorId).forEach(l => {
    if (GENERIC.test(l.dentistaSistema || '')) { out.genericos.push(l); return; }
    const d = Core.matchDentist(l.dentistaSistema, ativos());
    if (!d) { out.semCadastro.push(l); return; }
    if ((l.candidatos || []).includes(d.id)) out.aplicar.push([l, d]); else out.fora.push(l);
  });
  return out;
}
$('#btnAutoResp').onclick = async () => {
  if (S.ro || !S.result) return;
  const a = analisarResp();
  const total = a.aplicar.length + a.genericos.length + a.semCadastro.length + a.fora.length;
  if (!total) { toast('Nenhum conflito para resolver.'); return; }
  const g = n => n ? `\n• ${n}` : '';
  const txt = `Dos ${n0(total)} conflito(s):` +
    `\n✓ ${n0(a.aplicar.length)} será(ão) associado(s) ao dentista do "Resp. no sistema" (${brl(a.aplicar.reduce((x, [l]) => x + l.valor, 0))})` +
    (a.genericos.length ? `\n⚠ ${n0(a.genericos.length)} fica(m) pendente(s): Resp. genérico (Clínico Geral, Sem Dentista…)` : '') +
    (a.semCadastro.length ? `\n⚠ ${n0(a.semCadastro.length)} fica(m) pendente(s): Resp. no sistema não está cadastrado` : '') +
    (a.fora.length ? `\n⚠ ${n0(a.fora.length)} fica(m) pendente(s): Resp. no sistema não está entre os dentistas que informaram a guia` : '') +
    `\n\nConsultas odontológicas continuam com a regra (CLINICA). Dá para desfazer depois.`;
  if (!a.aplicar.length) { await dialog('Nada a associar automaticamente', txt, { okText: 'Entendi' }); return; }
  if (!await dialog('Associar pelo Resp. no sistema?', txt, { okText: `Associar ${n0(a.aplicar.length)} linha(s)` })) return;
  a.aplicar.forEach(([l, d]) => { S.overrides[l.id] = d.id; S.overridesAuto[l.id] = true; });
  run(false); salvar();
  toast(`${n0(a.aplicar.length)} linha(s) associadas · ${n0(total - a.aplicar.length)} continuam em conflito para avaliar`);
  S.view.orig = 'conflito'; S.view.page = 0; renderSeg(); renderDet();
};
$('#btnAutoRespUndo').onclick = async () => {
  const ids = Object.keys(S.overridesAuto);
  if (!ids.length) { toast('Não há associações automáticas para desfazer.'); return; }
  if (!await dialog('Desfazer associação automática?', `${n0(ids.length)} linha(s) voltam a ficar como estavam antes do botão (em conflito). Trocas manuais feitas depois não são afetadas.`, { okText: 'Desfazer' })) return;
  ids.forEach(id => { delete S.overrides[id]; delete S.overridesAuto[id]; });
  run(false); salvar(); toast('Associação automática desfeita');
};
function renderAutoResp() {
  const n = S.result ? S.result.linhas.filter(l => l.origem === 'conflito' && !l.executorId).length : 0;
  const u = Object.keys(S.overridesAuto).length;
  $('#btnAutoResp').innerHTML = `⚡ Associar pelo Resp. no sistema${n ? ` <span class="cnt" style="background:var(--warnSoft);color:var(--warn)">${n0(n)}</span>` : ''}`;
  $('#btnAutoResp').disabled = S.ro || !n;
  $('#btnAutoRespUndo').classList.toggle('hide', !u || S.ro);
  $('#btnAutoRespUndo').textContent = `Desfazer (${n0(u)})`;
}
const PAGE = 50;
function renderDet() {
  const list = filtered();
  const pages = Math.max(1, Math.ceil(list.length / PAGE));
  if (S.view.page >= pages) S.view.page = pages - 1;
  const slice = list.slice(S.view.page * PAGE, (S.view.page + 1) * PAGE);
  $('#tblDet').innerHTML = `<colgroup><col style="width:7%"><col style="width:10%"><col style="width:22%"><col style="width:14%"><col style="width:8%"><col style="width:13%"><col style="width:16%"><col style="width:10%"></colgroup>
  <thead><tr><th>Data</th><th>Guia</th><th>Procedimento</th><th>Paciente</th><th>Situação</th><th>Resp. no sistema</th><th>Executor</th><th class="num">Valor</th></tr></thead>
  <tbody>${slice.length ? slice.map(l => `<tr>
    <td class="mono" style="font-size:12px">${l.data ? l.data.toLocaleDateString('pt-BR').slice(0, 5) : '—'}</td>
    <td class="mono" style="font-size:12px" title="Guia ${esc(l.guiaRaw)} · atend. ${esc(l.atendimento)}">${esc(l.guiaRaw)}</td>
    <td title="${esc(l.procedimento)}">${esc(l.procedimento || '—')}</td>
    <td title="${esc(l.paciente)}${l.convenio ? ' · ' + esc(l.convenio) : ''}">${esc(l.paciente || '—')}</td>
    <td><span class="${S.sitBase.has(l.situacao) ? '' : 'muted'}">${esc(l.situacao)}</span></td>
    <td title="${esc(l.dentistaSistema)}" class="${GENERIC.test(l.dentistaSistema) ? 'muted' : ''}">${esc(l.dentistaSistema)}</td>
    <td title="${esc(S.overridesAuto[l.id] && l.origem === 'manual' ? 'Associado pelo botão Resp. no sistema' : (l.alerta || l.regraTxt || ORIG_LBL[l.origem]))}"><div class="row" style="gap:4px;flex-wrap:nowrap"><span class="dot" style="background:${{ dentista: 'var(--ok)', atendimento: '#3fae7f', sistema: 'var(--info)', manual: '#6b46c1', nao_identificado: 'var(--bad)', conflito: 'var(--warn)', split: '#2f6fb0', regra: '#0e6b63' }[l.origem]}"></span>
      <select class="inp" data-ov="${l.id}" style="flex:1;min-width:0">${l.origem === 'split' ? `<option value="_cur" selected>÷ ${esc(partesTxt(l.split))}</option>` : ''}${dentOpts(l.executorId, l.origem === 'conflito' ? '⚠ escolher…' : '— sem executor —')}<option value="_split">÷ Dividir guia entre dentistas…</option></select>${l.alerta ? '<span title="' + esc(l.alerta) + '" style="color:var(--warn);cursor:help">⚠</span>' : ''}</div></td>
    <td class="num">${brl(l.valor)}${linhaPgto(l.id) ? `<div style="font-size:10.5px;color:var(--ok);font-family:Manrope" title="Enviado para pagamento em ${dataBR(linhaPgto(l.id))}">pg ${dataBR(linhaPgto(l.id)).slice(0, 5)}</div>` : ''}</td></tr>`).join('') : `<tr><td colspan="8" class="empty">Nenhuma linha neste filtro.</td></tr>`}</tbody>`;
  $('#tblDet').querySelectorAll('[data-ov]').forEach(s => s.onchange = async () => {
    const l = S.result.linhas.find(x => x.id === +s.dataset.ov);
    const daDivergencia = l.origem === 'split' && !S.splits[l.guia];
    if (s.value === '_cur') return;
    if (s.value === '_split') {
      if (daDivergencia) { toast('Esta linha foi dividida na divergência de valor. Ajuste por lá.'); renderDet(); return; }
      const gl = S.result.linhas.filter(x => x.guia === l.guia);
      const cur = S.splits[l.guia]?.partes || [{ dentistaId: l.executorId || '', pct: 50 }, { dentistaId: '', pct: 50 }];
      const r = await abrirSplit({ titulo: `Dividir a guia ${l.guiaRaw}`, hint: `${gl.length} linha(s) · ${l.paciente || ''}. A divisão vale para todas as linhas desta guia; cada uma mantém sua situação (faturada ou pendente).`, total: gl.reduce((a, x) => a + x.valor, 0), partes: cur, podeRemover: !!S.splits[l.guia] });
      if (r?.remover) delete S.splits[l.guia]; else if (r) S.splits[l.guia] = { partes: r.partes };
      run(false); salvar(); return;
    }
    if (daDivergencia) { toast('Esta linha foi dividida na divergência de valor. Ajuste por lá.'); renderDet(); return; }
    if (S.splits[l.guia]) { delete S.splits[l.guia]; toast(`Divisão da guia ${l.guiaRaw} removida`); }
    S.overrides[s.dataset.ov] = s.value; delete S.overridesAuto[s.dataset.ov]; run(false); salvar();
  });
  const sum = list.reduce((a, l) => a + l.valor, 0);
  $('#pager').innerHTML = `<span>${n0(list.length)} linhas · ${brl(sum)}</span><span class="sp"></span>
    <button class="btn sm" ${S.view.page ? '' : 'disabled'} data-p="-1">‹</button><span>pág. ${S.view.page + 1} de ${pages}</span><button class="btn sm" ${S.view.page < pages - 1 ? '' : 'disabled'} data-p="1">›</button>`;
  $('#pager').querySelectorAll('[data-p]').forEach(b => b.onclick = () => { S.view.page += +b.dataset.p; renderDet(); });
  applyRO();
}
function renderNE() {
  const ne = S.result.naoEncontradas;
  $('#cardNE').classList.toggle('hide', !ne.length);
  const val = x => x.decisao?.valor ?? x.valor ?? 0;
  $('#tblNE').innerHTML = `<colgroup><col style="width:7%"><col style="width:15%"><col style="width:20%"><col style="width:18%"><col style="width:8%"><col style="width:20%"><col style="width:12%"></colgroup>
  <thead><tr><th>Data</th><th>Guia / atend.</th><th>Paciente</th><th>Procedimento</th><th class="num">Informado</th><th>Atribuir a</th><th class="num">Valor a considerar</th></tr></thead>
  <tbody>${ne.map(x => { const d = x.decisao || {}; const sel = d.descartado ? '_x' : (d.dentistaId || ''); return `<tr>
    <td class="mono" style="font-size:12px">${x.data ? x.data.toLocaleDateString('pt-BR').slice(0, 5) : '—'}</td>
    <td class="two"><span class="l mono" style="font-size:12px">${esc(x.guia || '—')}</span><small>atend. ${esc(x.atend || '—')}</small></td>
    <td class="two"><span class="l" title="${esc(x.paciente)}">${esc(x.paciente || '—')}</span><small>informado por ${esc(dentById(x.informanteId)?.nome || '—')}</small></td>
    <td title="${esc(x.procedimento)}">${esc(x.procedimento || '—')}</td>
    <td class="num">${x.valor == null ? '—' : brl(x.valor)}</td>
    <td><select class="inp" data-ne="${esc(x.id)}" style="${sel ? '' : 'border-color:var(--warn)'}"><option value="">⚠ pendente</option><option value="_x" ${sel === '_x' ? 'selected' : ''}>Não considerar</option>${ativos().map(dd => `<option value="${dd.id}" ${dd.id === sel ? 'selected' : ''}>${esc(dd.nome)}${dd.id === x.informanteId ? ' (informou)' : ''}</option>`).join('')}</select></td>
    <td class="num"><input class="inp num" type="number" step="0.01" min="0" data-nev="${esc(x.id)}" value="${(+val(x)).toFixed(2)}" ${d.dentistaId ? '' : 'disabled'} style="text-align:right"></td></tr>`; }).join('')}</tbody>
  <tfoot><tr><td colspan="4">${n0(ne.length)} lançamento(s) · ${n0(S.result.tot.neAbertos)} pendente(s)</td><td class="num">${brl(ne.reduce((a, x) => a + (x.valor || 0), 0))}</td><td>Considerado</td><td class="num">${brl(ne.filter(x => x.decisao?.dentistaId).reduce((a, x) => a + (+val(x)), 0))}</td></tr></tfoot>`;
  $('#tblNE').querySelectorAll('[data-ne]').forEach(sl => sl.onchange = () => {
    const id = sl.dataset.ne, x = ne.find(y => y.id === id), cur = S.lancamentos[id] || {};
    if (!sl.value) delete S.lancamentos[id];
    else if (sl.value === '_x') S.lancamentos[id] = { descartado: true };
    else S.lancamentos[id] = { dentistaId: sl.value, valor: cur.valor ?? x.valor ?? 0 };
    run(false); salvar();
  });
  $('#tblNE').querySelectorAll('[data-nev]').forEach(i => i.onchange = () => { const l = S.lancamentos[i.dataset.nev]; if (!l?.dentistaId) return; l.valor = Math.max(0, +i.value || 0); run(false); salvar(); });
}
$('#bulkNE').onclick = () => {
  if (S.ro) return;
  let n = 0; S.result.naoEncontradas.forEach(x => { if (!S.lancamentos[x.id] && x.informanteId) { S.lancamentos[x.id] = { dentistaId: x.informanteId, valor: x.valor ?? 0 }; n++; } });
  if (!n) { toast('Nenhum lançamento pendente.'); return; }
  run(false); salvar(); toast(`${n} lançamento(s) atribuído(s) a quem informou`);
};
const DEC = { pendente: '⚠ Pendente', sistema: 'Manter sistema', informado: 'Usar valor informado', split: '÷ Dividir entre dentistas' };
function renderDiv() {
  const all = S.result.divergencias;
  $('#cardDiv').classList.toggle('hide', !all.length);
  if (!all.length) return;
  const isPend = d => !d.resolucao || d.resolucao.decisao === 'pendente';
  const V = S.divView;
  const nP = all.filter(isPend).length;
  $('#segDiv').innerHTML = [['pend', 'Pendentes', nP], ['ok', 'Resolvidas', all.length - nP], ['all', 'Todas', all.length]].map(([k, t, c]) => `<button class="${V.f === k ? 'on' : ''}" data-df="${k}">${t}<em>${c}</em></button>`).join('');
  $('#segDiv').querySelectorAll('button').forEach(b => b.onclick = () => { V.f = b.dataset.df; V.page = 0; renderDiv(); });
  const dv = all.filter(d => V.f === 'all' || (V.f === 'pend' ? isPend(d) : !isPend(d)));
  const inf = all.reduce((a, d) => a + d.informado, 0), sis = all.reduce((a, d) => a + d.sistema, 0);
  $('#divResumo').innerHTML = `<span class="chip">${n0(all.length)} atendimentos</span><span class="chip">Informado ${brl(inf)}</span><span class="chip">Sistema ${brl(sis)}</span><span class="chip">Diferença ${brl(inf - sis)}</span>${nP ? `<span class="chip miss">${n0(nP)} pendente(s)</span>` : '<span class="chip ok">✓ todas resolvidas</span>'}`;
  const PG = 25, pages = Math.max(1, Math.ceil(dv.length / PG)); if (V.page >= pages) V.page = pages - 1;
  const sl = dv.slice(V.page * PG, (V.page + 1) * PG);
  $('#tblDiv').innerHTML = `<colgroup><col style="width:19%"><col style="width:27%"><col style="width:8%"><col style="width:8%"><col style="width:8%"><col style="width:14%"><col style="width:16%"></colgroup>
  <thead><tr><th>Atendimento · guia / paciente</th><th>Procedimentos (informado · sistema)</th><th class="num">Informado</th><th class="num">Sistema</th><th class="num">Dif.</th><th>Decisão</th><th>Atribuir a</th></tr></thead>
  <tbody>${sl.length ? sl.map(d => { const r = d.resolucao || {}; const dec = r.decisao || 'pendente'; const alvo = r.dentistaId ?? d.informanteId; return `<tr>
    <td class="two"><span class="l"><span class="mono" style="font-size:12px">${esc(d.atendimento || d.guia)}</span> · ${esc(d.paciente || '—')}</span><small title="Guia(s) informada(s): ${esc(d.guia)}">${d.atendimento ? 'guia ' + esc(d.guia) + ' · ' : ''}${esc(dentById(d.informanteId)?.nome || '—')}${d.data ? ' · ' + d.data.toLocaleDateString('pt-BR').slice(0, 5) : ''}</small></td>
    <td class="two"><span class="l" title="${esc(d.procInformado)}">${esc(d.procInformado || '—')}</span><small title="${esc(d.procSistema)}">${d.procSistema ? esc(d.procSistema) : 'nenhuma linha do sistema para este dentista'}${d.pendentes ? ` · ${d.pendentes} linha(s) sem executor` : ''}</small></td>
    <td class="num">${brl(d.informado)}</td><td class="num">${brl(d.sistema)}</td>
    <td class="num" style="color:${d.dif > 0 ? 'var(--bad)' : 'var(--info)'}">${brl(d.dif)}</td>
    <td><select class="inp" data-dd="${esc(d.id)}" style="${dec === 'pendente' ? 'border-color:var(--warn)' : ''}">${Object.entries(DEC).map(([k, t]) => `<option value="${k}" ${k === dec ? 'selected' : ''}>${t}</option>`).join('')}</select></td>
    <td>${dec === 'split' ? `<div class="row" style="gap:4px;flex-wrap:nowrap"><span title="${esc(partesTxt(r.partes))} · base ${r.base === 'informado' ? 'informado' : 'sistema'}" style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;font-size:12.5px">${esc(partesTxt(r.partes))}</span><button class="btn sm" data-dse="${esc(d.id)}">Editar</button></div>` : `<select class="inp" data-dt="${esc(d.id)}" ${dec === 'pendente' ? 'disabled' : ''}>${dentOpts(alvo, '— nenhum (não comissiona) —')}</select>`}</td></tr>`; }).join('') : `<tr><td colspan="7" class="empty">${V.f === 'pend' ? '✓ Nenhuma divergência pendente.' : 'Nada aqui.'}</td></tr>`}</tbody>`;
  const editarSplitDiv = async d => {
    const cur = S.resolucoes[d.id]?.decisao === 'split' ? S.resolucoes[d.id] : null;
    const r = await abrirSplit({ titulo: `Dividir ${d.atendimento ? 'o atendimento ' + d.atendimento : 'a guia ' + d.guia}`, hint: `${d.paciente || ''} · informado por ${dentById(d.informanteId)?.nome || '—'}. Escolha o valor a dividir e o percentual de cada dentista.`,
      total: d.sistema, totalInf: d.informado, base: cur?.base || 'sistema', partes: cur?.partes || [{ dentistaId: d.informanteId, pct: 50 }, { dentistaId: '', pct: 50 }], podeRemover: !!cur });
    if (r?.remover) delete S.resolucoes[d.id];
    else if (r) S.resolucoes[d.id] = { decisao: 'split', base: r.base, partes: r.partes };
    run(false); salvar();
  };
  $('#tblDiv').querySelectorAll('[data-dd]').forEach(x => x.onchange = () => {
    const d = all.find(y => y.id === x.dataset.dd); const cur = S.resolucoes[d.id] || {};
    if (x.value === 'split') { editarSplitDiv(d); return; }
    if (x.value === 'pendente') delete S.resolucoes[d.id]; else S.resolucoes[d.id] = { decisao: x.value, dentistaId: cur.dentistaId ?? d.informanteId };
    run(false); salvar();
  });
  $('#tblDiv').querySelectorAll('[data-dse]').forEach(b => b.onclick = () => editarSplitDiv(all.find(y => y.id === b.dataset.dse)));
  $('#tblDiv').querySelectorAll('[data-dt]').forEach(x => x.onchange = () => { const r = S.resolucoes[x.dataset.dt]; if (!r) return; r.dentistaId = x.value; run(false); salvar(); });
  $('#pagerDiv').innerHTML = pages > 1 ? `<span>${n0(dv.length)} atendimentos</span><span class="sp"></span><button class="btn sm" ${V.page ? '' : 'disabled'} data-dp="-1">‹</button><span>pág. ${V.page + 1} de ${pages}</span><button class="btn sm" ${V.page < pages - 1 ? '' : 'disabled'} data-dp="1">›</button>` : '';
  $('#pagerDiv').querySelectorAll('[data-dp]').forEach(b => b.onclick = () => { V.page += +b.dataset.dp; renderDiv(); });
}
function bulkDiv(decisao) {
  if (S.ro) return;
  let n = 0; S.result.divergencias.forEach(d => { if (!S.resolucoes[d.id]) { S.resolucoes[d.id] = { decisao, dentistaId: d.informanteId }; n++; } });
  if (!n) { toast('Nenhuma divergência pendente.'); return; }
  run(false); salvar(); toast(`${n} divergência(s) resolvida(s): ${DEC[decisao]}`);
}
$('#bulkSis').onclick = () => bulkDiv('sistema');
$('#bulkInf').onclick = () => bulkDiv('informado');

/* ===================== divisão entre dentistas ===================== */
const partesTxt = partes => (partes || []).filter(p => p.dentistaId).map(p => `${dentById(p.dentistaId)?.nome || '?'} ${String(Math.round(p.pct * 100) / 100).replace('.', ',')}%`).join(' · ');
function abrirSplit({ titulo, hint, total, totalInf = null, base = 'sistema', partes, podeRemover }) {
  return new Promise(res => {
    let P = partes.map(p => ({ dentistaId: p.dentistaId || '', pct: +p.pct || 0 }));
    let B = base;
    const totalAtual = () => (B === 'informado' && totalInf != null ? totalInf : total);
    $('#spTitle').textContent = titulo; $('#spHint').textContent = hint || ''; $('#spMsg').textContent = '';
    $('#spBase').classList.toggle('hide', totalInf == null);
    $('#spRemover').classList.toggle('hide', !podeRemover);
    const draw = () => {
      if (totalInf != null) {
        $('#spBaseSeg').innerHTML = [['sistema', `Sistema ${brl(total)}`], ['informado', `Informado ${brl(totalInf)}`]].map(([k, t]) => `<button class="${B === k ? 'on' : ''}" data-b="${k}">${t}</button>`).join('');
        $('#spBaseSeg').querySelectorAll('button').forEach(b => b.onclick = () => { B = b.dataset.b; draw(); });
      }
      $('#spTbl').innerHTML = `<colgroup><col style="width:52%"><col style="width:18%"><col style="width:22%"><col style="width:8%"></colgroup>
        <thead><tr><th>Dentista</th><th class="num">%</th><th class="num">Valor</th><th></th></tr></thead>
        <tbody>${P.map((p, i) => `<tr><td><select class="inp" data-spd="${i}">${dentOpts(p.dentistaId)}</select></td>
          <td><input class="inp num" type="number" min="0" max="100" step="0.01" data-spp="${i}" value="${p.pct}" style="text-align:right"></td>
          <td class="num">${brl(totalAtual() * p.pct / 100)}</td>
          <td style="text-align:right">${P.length > 2 ? `<button class="btn ghost sm danger" data-spx="${i}">✕</button>` : ''}</td></tr>`).join('')}</tbody>`;
      const soma = P.reduce((a, p) => a + (+p.pct || 0), 0);
      $('#spSoma').innerHTML = `Total: <b style="color:${Math.abs(soma - 100) < 0.01 ? 'var(--ok)' : 'var(--bad)'}">${String(Math.round(soma * 100) / 100).replace('.', ',')}%</b> · ${brl(totalAtual() * soma / 100)} de ${brl(totalAtual())}`;
      $('#spTbl').querySelectorAll('[data-spd]').forEach(x => x.onchange = () => { P[+x.dataset.spd].dentistaId = x.value; });
      $('#spTbl').querySelectorAll('[data-spp]').forEach(x => x.onchange = () => { P[+x.dataset.spp].pct = Math.max(0, Math.min(100, +x.value || 0)); draw(); });
      $('#spTbl').querySelectorAll('[data-spx]').forEach(x => x.onclick = () => { P.splice(+x.dataset.spx, 1); draw(); });
    };
    $('#spAdd').onclick = () => { P.push({ dentistaId: '', pct: 0 }); draw(); };
    $('#spIgual').onclick = () => { const n = P.length; P.forEach((p, i) => p.pct = i < n - 1 ? Math.floor(10000 / n) / 100 : Math.round((100 - Math.floor(10000 / n) / 100 * (n - 1)) * 100) / 100); draw(); };
    const fim = v => { $('#splitModal').classList.remove('on'); res(v); };
    $('#spCancel').onclick = () => fim(null);
    $('#spRemover').onclick = () => fim({ remover: true });
    $('#spOk').onclick = () => {
      const ps = P.filter(p => p.dentistaId && p.pct > 0);
      const soma = ps.reduce((a, p) => a + p.pct, 0);
      const ids = ps.map(p => p.dentistaId);
      if (ps.length < 2) { $('#spMsg').textContent = 'Escolha pelo menos dois dentistas com percentual.'; return; }
      if (new Set(ids).size !== ids.length) { $('#spMsg').textContent = 'O mesmo dentista aparece duas vezes.'; return; }
      if (Math.abs(soma - 100) > 0.01) { $('#spMsg').textContent = `Os percentuais somam ${String(Math.round(soma * 100) / 100).replace('.', ',')}%; precisam fechar 100%.`; return; }
      fim({ partes: ps, base: B });
    };
    draw(); $('#splitModal').classList.add('on');
  });
}

/* ===================== guias pendentes ===================== */
$('#pFDent').onchange = e => { S.pendView.dent = e.target.value; S.pendView.page = 0; renderPend(); };
$('#pFBusca').oninput = e => { S.pendView.q = e.target.value; S.pendView.page = 0; renderPend(); };
const execTxt = l => l.origem === 'split' ? '÷ ' + partesTxt(l.split) : (dentById(l.executorId)?.nome || '— sem executor —');
function renderPend() {
  const R = S.result;
  $('#pendEmpty').classList.toggle('hide', !!R); $('#pendBody').classList.toggle('hide', !R);
  const cnt = $('#cntPend');
  if (!R) { cnt.style.display = 'none'; return; }
  const TODAS = R.pendentes, P = TODAS.filter(l => l.pend), A = TODAS.filter(l => l.aprovado);
  cnt.textContent = P.length; cnt.style.display = P.length ? '' : 'none';
  $('#pendChips').innerHTML = `<span class="chip miss">🔒 Em aberto: ${n0(P.length)} · <b>${brl(R.tot.pendentesValor)}</b> · fora do movimento</span>
    <span class="chip ok">✓ Pagamento aprovado: ${n0(A.length)} · <b>${brl(R.tot.aprovadasValor)}</b> · entram na produção e na comissão</span>`;
  // por executor / por motivo (somente em aberto)
  const byE = {};
  P.forEach(l => {
    if (l.origem === 'split') (l.split || []).forEach(p => { const k = p.dentistaId; const o = byE[k] || (byE[k] = { nome: dentById(k)?.nome || '?', n: 0, v: 0 }); o.n++; o.v += l.valor * p.pct / 100; });
    else { const k = l.executorId || '_'; const o = byE[k] || (byE[k] = { nome: dentById(l.executorId)?.nome || '— sem executor —', n: 0, v: 0 }); o.n++; o.v += l.valor; }
  });
  const eRows = Object.entries(byE).sort((a, b) => b[1].v - a[1].v);
  $('#tblPendDent').innerHTML = `<colgroup><col style="width:52%"><col style="width:18%"><col style="width:30%"></colgroup><thead><tr><th>Executor (em aberto)</th><th class="num">Linhas</th><th class="num">Valor</th></tr></thead>
    <tbody>${eRows.map(([k, o]) => `<tr style="cursor:pointer" data-pe="${k}"><td>${k === '_' ? '<span class="muted">— sem executor —</span>' : '<b>' + esc(o.nome) + '</b>'}</td><td class="num">${n0(o.n)}</td><td class="num">${brl(o.v)}</td></tr>`).join('') || '<tr><td colspan="3" class="empty">Nenhuma guia pendente em aberto 🎉</td></tr>'}</tbody>`;
  $('#tblPendDent').querySelectorAll('[data-pe]').forEach(tr => tr.onclick = () => { S.pendView.dent = tr.dataset.pe; S.pendView.page = 0; renderPend(); });
  const byM = {};
  P.forEach(l => { const txt = (l.obs || '').replace(/[.\s]+$/, ''); const k = Core.norm(txt) || '_'; const o = byM[k] || (byM[k] = { nome: txt || '(sem motivo informado)', n: 0, v: 0 }); o.n++; o.v += l.valor; });
  const mRows = Object.entries(byM).map(([k, o]) => [o.nome, o]).sort((a, b) => b[1].n - a[1].n);
  $('#tblPendMot').innerHTML = `<colgroup><col style="width:66%"><col style="width:12%"><col style="width:22%"></colgroup><thead><tr><th>Motivo (em aberto)</th><th class="num">Linhas</th><th class="num">Valor</th></tr></thead>
    <tbody>${mRows.slice(0, 12).map(([k, o]) => `<tr><td title="${esc(k)}">${esc(k)}</td><td class="num">${n0(o.n)}</td><td class="num">${brl(o.v)}</td></tr>`).join('') || '<tr><td colspan="3" class="empty">—</td></tr>'}${mRows.length > 12 ? `<tr><td colspan="3" class="muted" style="font-size:12px">+ ${mRows.length - 12} outros motivos (veja a lista abaixo)</td></tr>` : ''}</tbody>`;
  // filtros
  const V = S.pendView;
  $('#segPend').innerHTML = [['abertas', 'Em aberto', P.length], ['aprovadas', 'Aprovadas', A.length], ['todas', 'Todas', TODAS.length]].map(([k, t, c]) => `<button class="${V.f === k ? 'on' : ''}" data-pf="${k}">${t}<em>${c}</em></button>`).join('');
  $('#segPend').querySelectorAll('button').forEach(b => b.onclick = () => { V.f = b.dataset.pf; V.page = 0; renderPend(); });
  $('#pFDent').innerHTML = `<option value="">Todos os executores</option><option value="_">— sem executor —</option>` + ativos().map(d => `<option value="${d.id}">${esc(d.nome)}</option>`).join('');
  $('#pFDent').value = V.dent;
  const qq = Core.norm(V.q);
  const base = V.f === 'abertas' ? P : V.f === 'aprovadas' ? A : TODAS;
  const list = base.filter(l => (!V.dent || (V.dent === '_' ? !l.executorId && l.origem !== 'split' : l.executorId === V.dent || (l.split || []).some(p => p.dentistaId === V.dent)))
    && (!qq || Core.norm([l.guiaRaw, l.atendimento, l.paciente, l.procedimento, l.obs, l.dentistaSistema].join(' ')).includes(qq)));
  const aprovaveis = list.filter(l => l.pend && (l.executorId || l.origem === 'split'));
  $('#pBulk').textContent = `✓ Aprovar pgto das filtradas com executor (${n0(aprovaveis.length)})`;
  $('#pBulk').disabled = S.ro || !aprovaveis.length;
  const PG = 50, pages = Math.max(1, Math.ceil(list.length / PG)); if (V.page >= pages) V.page = pages - 1;
  const sl = list.slice(V.page * PG, (V.page + 1) * PG);
  $('#tblPend').innerHTML = `<colgroup><col style="width:6%"><col style="width:10%"><col style="width:13%"><col style="width:14%"><col style="width:17%"><col style="width:17%"><col style="width:8%"><col style="width:15%"></colgroup>
    <thead><tr><th>Data</th><th>Guia</th><th>Paciente</th><th>Procedimento</th><th>Executor</th><th>Motivo da pendência</th><th class="num">Valor</th><th style="text-align:right">Ação</th></tr></thead>
    <tbody>${sl.length ? sl.map(l => { const ap = S.aprovadas[l.id]; return `<tr style="${l.aprovado ? 'background:var(--okSoft)' : ''}"><td class="mono" style="font-size:12px">${l.data ? l.data.toLocaleDateString('pt-BR').slice(0, 5) : '—'}</td>
      <td class="mono" style="font-size:12px" title="atend. ${esc(l.atendimento)}">${esc(l.guiaRaw)}</td><td title="${esc(l.paciente)}">${esc(l.paciente || '—')}</td>
      <td title="${esc(l.procedimento)}">${esc(l.procedimento || '—')}</td>
      <td title="${esc(l.alerta || l.regraTxt || ORIG_LBL[l.origem] || '')}"><select class="inp" data-pov="${l.id}" ${S.ro ? 'disabled' : ''} style="${l.executorId || l.origem === 'split' ? '' : 'border-color:var(--warn)'}">${l.origem === 'split' ? `<option value="_cur" selected>÷ ${esc(partesTxt(l.split))}</option>` : ''}${dentOpts(l.executorId, '— sem executor —')}</select></td>
      <td class="two"><span class="l" title="${esc(l.obs)}">${esc(l.obs || '—')}</span><small>${esc(l.status || l.situacao)}</small></td><td class="num">${brl(l.valor)}</td>
      <td style="text-align:right">${l.aprovado ? `<div class="row" style="gap:6px;justify-content:flex-end;flex-wrap:nowrap"><div style="text-align:right;line-height:1.2" title="${esc(ap?.por || '')}"><b style="color:var(--ok);font-size:12px">✓ Aprovado</b><div class="muted" style="font-size:11px">${ap?.em ? dtf(ap.em) : ''}</div></div><button class="btn ghost sm" data-desap="${l.id}" ${S.ro ? 'disabled' : ''} title="Desfazer aprovação">↺</button></div>` : `<button class="btn sm pri" data-apr="${l.id}" ${S.ro ? 'disabled' : ''}>Aprovar pgto</button>`}</td></tr>`; }).join('') : `<tr><td colspan="8" class="empty">${V.f === 'abertas' ? 'Nenhuma guia pendente em aberto.' : 'Nenhuma guia neste filtro.'}</td></tr>`}</tbody>`;
  const T = $('#tblPend');
  T.querySelectorAll('[data-pov]').forEach(x => x.onchange = () => { if (x.value === '_cur') return; S.overrides[x.dataset.pov] = x.value; delete S.overridesAuto[x.dataset.pov]; run(false); salvar(); });
  T.querySelectorAll('[data-apr]').forEach(b => b.onclick = () => {
    const l = TODAS.find(x => x.id === +b.dataset.apr);
    if (!l.executorId && l.origem !== 'split') { toast('Escolha o dentista executor antes de aprovar.'); T.querySelector(`[data-pov="${l.id}"]`)?.focus(); return; }
    S.aprovadas[l.id] = { em: new Date().toISOString(), por: S.user?.email || '' };
    run(false); salvar(); toast(`Guia ${l.guiaRaw} aprovada: ${brl(l.valor)} entrou na produção de ${execTxt(l)}`);
  });
  T.querySelectorAll('[data-desap]').forEach(b => b.onclick = () => { delete S.aprovadas[b.dataset.desap]; run(false); salvar(); toast('Aprovação desfeita: a guia voltou para pendente'); });
  $('#pagerPend').innerHTML = `<span>${n0(list.length)} linhas · ${brl(list.reduce((a, l) => a + l.valor, 0))}</span><span class="sp"></span>
    <button class="btn sm" ${V.page ? '' : 'disabled'} data-pp="-1">‹</button><span>pág. ${V.page + 1} de ${pages}</span><button class="btn sm" ${V.page < pages - 1 ? '' : 'disabled'} data-pp="1">›</button>`;
  $('#pagerPend').querySelectorAll('[data-pp]').forEach(b => b.onclick = () => { V.page += +b.dataset.pp; renderPend(); });
  $('#pBulk').onclick = async () => {
    if (!aprovaveis.length) return;
    const semExec = list.filter(l => l.pend && !l.executorId && l.origem !== 'split').length;
    if (!await dialog('Aprovar pagamento das guias filtradas?', `${n0(aprovaveis.length)} linha(s) · ${brl(aprovaveis.reduce((a, l) => a + l.valor, 0))} entram na produção e na comissão dos executores.${semExec ? `\n${n0(semExec)} linha(s) sem executor ficam de fora.` : ''}`, { okText: 'Aprovar' })) return;
    const em = new Date().toISOString();
    aprovaveis.forEach(l => { S.aprovadas[l.id] = { em, por: S.user?.email || '' }; });
    run(false); salvar(); toast(`${n0(aprovaveis.length)} guia(s) aprovada(s)`);
  };
}

/* somente leitura quando fechada */
function applyRO() {
  if (!S.ro) { ['#tolDiv', '#bulkSis', '#bulkInf', '#bulkNE', '#fileSys', '#fileDen'].forEach(x => { const el = $(x); if (el) el.disabled = false; }); return; }
  document.querySelectorAll('#tab-conc input, #tab-conc select, #tab-res select, #tab-res input:not(#fBusca), #btnRun, #btnPgto, #bulkSis, #bulkInf, #bulkNE, [data-rm], [data-dse]').forEach(el => { if (el.id !== 'compMes') el.disabled = true; });
}
function renderAllConc() { applyRO(); renderSys(); renderSit(); renderDen(); checkReady(); renderResult(); renderSugestoes(); renderRegras(); applyRO(); }

/* ===================== exportar ===================== */
$('#btnExport').onclick = () => {
  if (!S.result) { toast('Faça a conciliação primeiro.'); return; }
  exportarExcel(S.result, `Conciliacao_${compLabel(S.conc.competencia).replace('/', '_')}${S.ro ? '_fechada' : ''}.xlsx`);
};
function exportarExcel(R, nomeArquivo) {
  const nm = id => dentById(id)?.nome || '';
  const wb = XLSX.utils.book_new();
  const res = R.resumo.map(r => ({ Dentista: r.nome, Especialidade: r.especialidade, Remuneração: FORMAS[r.forma], 'Valor/%': r.valorRef, Guias: r.guias, Procedimentos: r.procs, Produção: +(+r.producao).toFixed(2), Ajustes: +(+r.ajustes || 0).toFixed(2), 'Base comissionável': +(+r.base).toFixed(2), Diárias: r.forma === 'diaria' ? r.diarias : '', Comissão: +(+r.comissao).toFixed(2), 'Enviado p/ pgto': r.dentistaId && S.pagamentos ? dataBR(pgtoDe(r.dentistaId)?.data || '') : '' }));
  res.push({ Dentista: 'TOTAL', Produção: +R.resumo.reduce((a, r) => a + +r.producao, 0).toFixed(2), Ajustes: +R.resumo.reduce((a, r) => a + (+r.ajustes || 0), 0).toFixed(2), 'Base comissionável': +R.resumo.reduce((a, r) => a + +r.base, 0).toFixed(2), Comissão: +R.resumo.reduce((a, r) => a + +r.comissao, 0).toFixed(2) });
  const ws1 = XLSX.utils.json_to_sheet(res); ws1['!cols'] = [28, 14, 14, 9, 8, 12, 13, 12, 16, 8, 13].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws1, 'Comissões');
  if (R.linhas) {
    const det = R.linhas.map(l => ({ Data: l.data ? l.data.toLocaleDateString('pt-BR') : (l.d || ''), Guia: l.guiaRaw ?? l.g, Atendimento: l.atendimento ?? l.a, Procedimento: l.procedimento ?? l.pr, Paciente: l.paciente ?? l.p, Situação: l.situacao ?? l.s, 'Resp. no sistema': l.dentistaSistema ?? l.ds, Executor: l.en ?? (l.origem === 'split' ? 'Dividida: ' + partesTxt(l.split) : nm(l.executorId)), Origem: ORIG_LBL[l.origem ?? l.o], Comissionável: (l.comissionavel ?? l.c) ? 'Sim' : 'Não', Valor: l.valor ?? l.v, 'Data pgto': l.id != null && S.pagamentos ? dataBR(linhaPgto(l.id) || '') : '', Observação: l.alerta || '' }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(det), 'Detalhamento');
    if (R.pendentes?.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(R.pendentes.map(l => ({ Data: l.data ? l.data.toLocaleDateString('pt-BR') : '', Guia: l.guiaRaw, Atendimento: l.atendimento, Paciente: l.paciente, Procedimento: l.procedimento, Executor: execTxt(l), 'Motivo da pendência': l.obs || '', Status: l.status || l.situacao, Valor: l.valor, 'Pgto aprovado': l.aprovado ? 'Sim' : 'Não', 'Aprovado em': S.aprovadas[l.id]?.em ? new Date(S.aprovadas[l.id].em).toLocaleString('pt-BR') : '', 'Aprovado por': S.aprovadas[l.id]?.por || '' }))), 'Guias pendentes');
  }
  if (R.divergencias?.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(R.divergencias.map(d => ({ Atendimento: d.atendimento, 'Guia(s) informada(s)': d.guia, Paciente: d.paciente, 'Informado por': nm(d.informanteId), 'Procedimento informado': d.procInformado, 'Procedimento no sistema': d.procSistema, 'Valor informado': +d.informado.toFixed(2), 'Valor no sistema': +d.sistema.toFixed(2), Diferença: +d.dif.toFixed(2), Decisão: DEC[d.resolucao?.decisao || 'pendente'].replace('⚠ ', ''), 'Atribuído a': d.resolucao ? nm(d.resolucao.dentistaId) : '' }))), 'Divergências');
  if (R.naoEncontradas?.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(R.naoEncontradas.map(x => ({ Data: x.data ? x.data.toLocaleDateString('pt-BR') : '', Guia: x.guia, Atendimento: x.atend, Paciente: x.paciente, Procedimento: x.procedimento, 'Valor informado': x.valor ?? '', 'Informado por': nm(x.informanteId), Decisão: x.decisao?.dentistaId ? 'Atribuído' : x.decisao?.descartado ? 'Não considerar' : 'Pendente', 'Atribuído a': nm(x.decisao?.dentistaId), 'Valor considerado': x.decisao?.dentistaId ? x.decisao.valor : '', Arquivo: x.arquivo }))), 'Não localizados');
  if (R.extras?.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(R.extras.map(e => ({ Tipo: e.tipo === 'ajuste' ? 'Ajuste de divergência' : e.tipo === 'split' ? 'Divisão entre dentistas' : 'Lançamento não localizado', Comissionável: e.comissionavel === false ? 'Não' : 'Sim', Dentista: nm(e.dentistaId), Referência: e.ref, Paciente: e.paciente, Descrição: e.descricao, Valor: +e.valor.toFixed(2) }))), 'Ajustes');
  XLSX.writeFile(wb, nomeArquivo);
}

/* ===================== cadastro de dentistas ===================== */
function espOpts(sel) { return S.cad.especialidades.filter(e => e.ativo || e.id === sel).map(e => `<option value="${e.id}" ${e.id === sel ? 'selected' : ''}>${esc(e.nome)}</option>`).join(''); }
function updValLabel() {
  const f = $('#dForma').value || espById($('#dEsp').value)?.forma || 'percentual';
  $('#dValLbl').textContent = f === 'percentual' ? 'Percentual (%)' : f === 'diaria' ? 'Valor da diária (R$)' : 'Valor mensal (R$)';
  const e = espById($('#dEsp').value);
  $('#dValor').placeholder = !$('#dForma').value && e ? `padrão: ${e.forma === 'percentual' ? e.valor + '%' : brl(e.valor)}` : '';
}
$('#dEsp').onchange = updValLabel; $('#dForma').onchange = updValLabel;
function resetForm() {
  S.editId = null; $('#dNome').value = ''; $('#dAlias').value = ''; $('#dForma').value = ''; $('#dValor').value = '';
  $('#dSave').textContent = 'Adicionar'; $('#dCancel').classList.add('hide'); updValLabel();
}
$('#dCancel').onclick = resetForm;
$('#dSave').onclick = async () => {
  const nome = $('#dNome').value.trim();
  if (!nome) { toast('Informe o nome do dentista.'); $('#dNome').focus(); return; }
  if (S.cad.dentistas.find(d => d.ativo && d.id !== S.editId && Core.nameKey(d.nome) === Core.nameKey(nome))) { toast('Já existe um dentista ativo com esse nome.'); return; }
  const v = $('#dValor').value;
  const row = { nome, aliases: $('#dAlias').value.trim() || null, especialidade_id: $('#dEsp').value || null, forma: $('#dForma').value || null, valor: v === '' ? null : +v };
  if (row.forma && row.valor === null) { toast('Informe o valor/percentual da forma escolhida.'); return; }
  $('#dSave').disabled = true;
  try {
    if (S.editId) await q(sb.from('dentistas').update(row).eq('id', S.editId), 'Salvar dentista');
    else await q(sb.from('dentistas').insert(row), 'Cadastrar dentista');
    toast(S.editId ? 'Dentista atualizado' : 'Dentista cadastrado'); resetForm(); await cadChanged();
  } catch (e) { fail(e); }
  $('#dSave').disabled = false;
};
function regra(d) {
  const e = espById(d.especialidadeId);
  const f = d.forma || e?.forma || 'percentual'; const v = d.valor ?? e?.valor ?? 0;
  return { f, v, txt: f === 'percentual' ? `${v}%` : f === 'diaria' ? `${brl(v)} / dia` : `${brl(v)} / mês`, padrao: !d.forma };
}
$('#showInativos').onchange = () => renderDentistas();
function renderDentistas() {
  $('#cntDent').textContent = ativos().length;
  const cur = $('#dEsp').value; $('#dEsp').innerHTML = espOpts(cur || S.cad.especialidades.find(e => e.ativo)?.id); updValLabel();
  const list = S.cad.dentistas.filter(d => d.ativo || $('#showInativos').checked).sort((a, b) => (b.ativo - a.ativo) || a.nome.localeCompare(b.nome));
  $('#tblDent').innerHTML = `<colgroup><col style="width:24%"><col style="width:26%"><col style="width:15%"><col style="width:12%"><col style="width:10%"><col style="width:13%"></colgroup>
  <thead><tr><th>Dentista</th><th>Nomes no sistema</th><th>Especialidade</th><th>Remuneração</th><th class="num">Valor / %</th><th></th></tr></thead>
  <tbody>${list.length ? list.map(d => { const r = regra(d); return `<tr style="${d.ativo ? '' : 'opacity:.55'}">
    <td><b>${esc(d.nome)}</b>${d.ativo ? '' : ' <span class="tag nao_identificado">inativo</span>'}</td><td class="muted" title="${esc(d.aliases)}">${esc(d.aliases || '—')}</td>
    <td>${esc(espById(d.especialidadeId)?.nome || '—')}</td><td>${FORMAS[r.f]}${r.padrao ? ' <span class="muted" style="font-size:11px">(padrão)</span>' : ''}</td>
    <td class="num">${r.txt}</td>
    <td style="text-align:right">${d.ativo ? `<button class="btn ghost sm" data-ed="${d.id}">Editar</button><button class="btn ghost sm danger" data-ina="${d.id}" title="Inativar">Inativar</button>` : `<button class="btn ghost sm" data-rea="${d.id}">Reativar</button>`}</td></tr>`; }).join('')
    : `<tr><td colspan="6" class="empty"><b>Nenhum dentista cadastrado</b>Use o formulário acima ou cadastre a partir dos nomes encontrados na produção.</td></tr>`}</tbody>`;
  $('#tblDent').querySelectorAll('[data-ed]').forEach(b => b.onclick = () => {
    const d = dentById(b.dataset.ed); S.editId = d.id;
    $('#dNome').value = d.nome; $('#dAlias').value = d.aliases || ''; $('#dEsp').value = d.especialidadeId || ''; $('#dForma').value = d.forma || ''; $('#dValor').value = d.valor ?? '';
    $('#dSave').textContent = 'Salvar alterações'; $('#dCancel').classList.remove('hide'); updValLabel(); window.scrollTo({ top: 0, behavior: 'smooth' }); $('#dNome').focus();
  });
  $('#tblDent').querySelectorAll('[data-ina]').forEach(b => b.onclick = async () => {
    const d = dentById(b.dataset.ina);
    if (!await dialog('Inativar dentista?', `${d.nome} sai das listas, mas continua no histórico das competências.`, { okText: 'Inativar' })) return;
    try { await q(sb.from('dentistas').update({ ativo: false }).eq('id', d.id), 'Inativar'); if (S.editId === d.id) resetForm(); await cadChanged(); } catch (e) { fail(e); }
  });
  $('#tblDent').querySelectorAll('[data-rea]').forEach(b => b.onclick = async () => { try { await q(sb.from('dentistas').update({ ativo: true }).eq('id', b.dataset.rea), 'Reativar'); await cadChanged(); } catch (e) { fail(e); } });
  renderSugestoes();
}
function renderSugestoes() {
  const el = $('#sugCard');
  if (!S.sys) { el.innerHTML = `<h3 style="margin:0 0 4px">Cadastro rápido</h3><p class="hint" style="margin:0">Carregue a planilha do sistema na aba Conciliação para ver os nomes de dentistas encontrados e cadastrá-los com um clique.</p>`; return; }
  const counts = {};
  S.sys.rows.forEach(r => { counts[r.dentistaSistema] = (counts[r.dentistaSistema] || 0) + r.valor; });
  const names = Object.keys(counts).filter(n => !Core.matchDentist(n, ativos()));
  const reais = names.filter(n => !GENERIC.test(n)).sort((a, b) => counts[b] - counts[a]);
  const gen = names.filter(n => GENERIC.test(n));
  el.innerHTML = `<h3 style="margin:0 0 4px">Nomes encontrados na produção sem cadastro</h3>
    <p class="hint">Clique para preencher o formulário. ${gen.length ? `Os genéricos (${gen.map(esc).join(', ')}) não precisam de cadastro: são resolvidos pelas planilhas dos dentistas.` : ''}</p>
    <div class="chips">${reais.length ? reais.map(n => `<button class="chip" style="cursor:pointer;border-color:var(--line)" data-sug="${esc(n)}">+ ${esc(n)} <span class="mono muted" style="font-size:11px">${brl(counts[n])}</span></button>`).join('') : '<span class="chip ok">✓ Todos os dentistas da produção estão cadastrados</span>'}</div>`;
  el.querySelectorAll('[data-sug]').forEach(b => b.onclick = () => { resetForm(); $('#dNome').value = b.dataset.sug; showTab('dent'); $('#dEsp').focus(); });
}

/* ===================== regras por procedimento ===================== */
function renderRegras() {
  $('#rgDent').innerHTML = dentOpts('', '— escolha o dentista —');
  const sug = S.sys ? [...new Set(S.sys.rows.map(r => r.procedimento).filter(Boolean))].sort() : [];
  $('#rgSug').innerHTML = sug.slice(0, 300).map(p => `<option value="${esc(p)}">`).join('');
  const regs = S.cad.regras.filter(r => r.ativo);
  const conta = r => S.sys ? S.sys.rows.filter(x => Core.norm(x.procedimento).includes(Core.norm(r.termo))).length : null;
  $('#tblRegras').innerHTML = `<colgroup><col style="width:44%"><col style="width:30%"><col style="width:14%"><col style="width:12%"></colgroup>
    <thead><tr><th>Procedimento contém</th><th>Executor</th><th class="num">Linhas na competência</th><th></th></tr></thead>
    <tbody>${regs.length ? regs.map(r => { const d = dentById(r.dentistaId); const n = conta(r); return `<tr>
      <td><b>${esc(r.termo)}</b></td><td>${esc(d?.nome || '—')}${d && !d.ativo ? ' <span class="tag nao_identificado">inativo — regra parada</span>' : ''}</td>
      <td class="num">${n == null ? '<span class="muted">—</span>' : n0(n)}</td>
      <td style="text-align:right"><button class="btn ghost sm danger" data-rgx="${r.id}">Inativar</button></td></tr>`; }).join('')
      : '<tr><td colspan="4" class="empty">Nenhuma regra. Ex.: "Consulta Odontológica" → CLINICA.</td></tr>'}</tbody>`;
  $('#tblRegras').querySelectorAll('[data-rgx]').forEach(b => b.onclick = async () => {
    const r = S.cad.regras.find(x => x.id === b.dataset.rgx);
    if (!await dialog('Inativar regra?', `"${r.termo}" deixa de ser aplicada a partir de agora. Competências já fechadas não mudam.`, { okText: 'Inativar' })) return;
    try { await q(sb.from('regras_procedimento').update({ ativo: false }).eq('id', r.id), 'Inativar regra'); await cadChanged(); toast('Regra inativada'); } catch (e) { fail(e); }
  });
}
$('#rgAdd').onclick = async () => {
  const termo = $('#rgTermo').value.trim(), dentista_id = $('#rgDent').value;
  if (termo.length < 2) { toast('Informe o texto do procedimento.'); $('#rgTermo').focus(); return; }
  if (!dentista_id) { toast('Escolha o dentista executor.'); return; }
  if (S.cad.regras.some(r => r.ativo && Core.norm(r.termo) === Core.norm(termo))) { toast('Já existe uma regra ativa para esse procedimento.'); return; }
  try { await q(sb.from('regras_procedimento').insert({ termo, dentista_id }), 'Criar regra'); $('#rgTermo').value = ''; await cadChanged(); toast('Regra criada e aplicada'); } catch (e) { fail(e); }
};

/* ===================== especialidades ===================== */
function renderEsp() {
  const used = id => S.cad.dentistas.filter(d => d.ativo && d.especialidadeId === id).length;
  const list = S.cad.especialidades.filter(e => e.ativo);
  $('#tblEsp').innerHTML = `<colgroup><col style="width:34%"><col style="width:26%"><col style="width:22%"><col style="width:10%"><col style="width:8%"></colgroup>
  <thead><tr><th>Especialidade</th><th>Forma de remuneração</th><th class="num">Valor / %</th><th class="num">Dentistas</th><th></th></tr></thead>
  <tbody>${list.map(e => `<tr>
    <td><input class="inp" data-en="${e.id}" value="${esc(e.nome)}"></td>
    <td><select class="inp" data-ef="${e.id}">${Object.entries(FORMAS).map(([k, v]) => `<option value="${k}" ${e.forma === k ? 'selected' : ''}>${v}</option>`).join('')}</select></td>
    <td><div class="row" style="gap:6px;flex-wrap:nowrap;justify-content:flex-end"><span class="muted">${e.forma === 'percentual' ? '%' : 'R$'}</span><input class="inp num" type="number" step="0.01" min="0" data-ev="${e.id}" value="${e.valor}" style="width:120px;text-align:right"></div></td>
    <td class="num">${used(e.id)}</td>
    <td style="text-align:right"><button class="btn ghost sm danger" data-edel="${e.id}" ${used(e.id) ? 'disabled title="Em uso por dentistas ativos"' : 'title="Inativar"'}>✕</button></td></tr>`).join('')}</tbody>`;
  const T = $('#tblEsp');
  const upd = async (id, patch) => { try { await q(sb.from('especialidades').update(patch).eq('id', id), 'Especialidade'); await cadChanged(); toast('Especialidade atualizada'); } catch (e) { fail(e); } };
  T.querySelectorAll('[data-en]').forEach(i => i.onchange = () => upd(i.dataset.en, { nome: i.value.trim() || 'Sem nome' }));
  T.querySelectorAll('[data-ef]').forEach(i => i.onchange = () => upd(i.dataset.ef, { forma: i.value }));
  T.querySelectorAll('[data-ev]').forEach(i => i.onchange = () => upd(i.dataset.ev, { valor: Math.max(0, +i.value || 0) }));
  T.querySelectorAll('[data-edel]').forEach(b => b.onclick = async () => {
    const e = espById(b.dataset.edel);
    if (!await dialog('Inativar especialidade?', e.nome, { okText: 'Inativar' })) return;
    upd(e.id, { ativo: false });
  });
}
$('#eAdd').onclick = async () => { try { await q(sb.from('especialidades').insert({ nome: 'Nova especialidade', forma: 'percentual', valor: 0 }), 'Especialidade'); await cadChanged(); } catch (e) { fail(e); } };

/* ===================== histórico ===================== */
async function loadHistorico() {
  try {
    const [concs, vig] = await Promise.all([
      q(sb.from('conciliacoes').select('id,competencia,status,versao,totais,arquivo_sistema,atualizado_em,fechada_em').order('competencia', { ascending: false }), 'Histórico'),
      q(sb.from('comissoes_vigentes').select('competencia,dentista_id,dentista_nome,comissao,producao').order('competencia', { ascending: false }).limit(2000), 'Comissões'),
    ]);
    $('#tblConc').innerHTML = `<colgroup><col style="width:14%"><col style="width:11%"><col style="width:7%"><col style="width:13%"><col style="width:13%"><col style="width:9%"><col style="width:15%"><col style="width:18%"></colgroup>
      <thead><tr><th>Competência</th><th>Status</th><th class="num">Versão</th><th class="num">Produção</th><th class="num">Comissão</th><th class="num">Pendências</th><th>Atualizado</th><th></th></tr></thead>
      <tbody>${concs.length ? concs.map(c => `<tr><td><b>${compLabel(c.competencia)}</b></td>
        <td>${c.status === 'fechada' ? '<span class="pill fech">✓ Fechada</span>' : '<span class="pill rasc">● Rascunho</span>'}</td>
        <td class="num">${c.versao}</td><td class="num">${brl(c.totais?.producao)}</td><td class="num"><b>${brl(c.totais?.comissao)}</b></td><td class="num">${c.totais?.pendencias ?? '—'}</td>
        <td class="muted" style="font-size:12px">${dtf(c.atualizado_em)}</td>
        <td style="text-align:right"><button class="btn sm ghost" data-vers="${c.id}" data-comp="${c.competencia}">Versões</button><button class="btn sm" data-abrir="${c.competencia}">Abrir</button></td></tr>`).join('')
      : '<tr><td colspan="8" class="empty">Nenhuma competência ainda.</td></tr>'}</tbody>`;
    $('#tblConc').querySelectorAll('[data-abrir]').forEach(b => b.onclick = async () => { $('#compMes').value = b.dataset.abrir.slice(0, 7); showTab('conc'); await abrirCompetencia($('#compMes').value); });
    $('#tblConc').querySelectorAll('[data-vers]').forEach(b => b.onclick = () => loadVersoes(b.dataset.vers, b.dataset.comp));
    // pivot de comissões
    const comps = [...new Set(vig.map(v => v.competencia))].sort().reverse().slice(0, 6).reverse();
    const dents = {};
    vig.filter(v => comps.includes(v.competencia)).forEach(v => { const k = v.dentista_id || v.dentista_nome; (dents[k] = dents[k] || { nome: v.dentista_nome, m: {} }).m[v.competencia] = +v.comissao; });
    const rows = Object.values(dents).sort((a, b) => a.nome.localeCompare(b.nome));
    const w = comps.length ? Math.floor(60 / (comps.length + 1)) : 60;
    $('#tblPivot').innerHTML = comps.length ? `<colgroup><col style="width:40%">${comps.map(() => `<col style="width:${w}%">`).join('')}<col style="width:${w}%"></colgroup>
      <thead><tr><th>Dentista</th>${comps.map(c => `<th class="num">${compLabel(c).replace(/^(\w{3})\w*/, '$1')}</th>`).join('')}<th class="num">Total</th></tr></thead>
      <tbody>${rows.map(r => `<tr><td><b>${esc(r.nome)}</b></td>${comps.map(c => `<td class="num">${r.m[c] != null ? brl(r.m[c]) : '<span class="muted">—</span>'}</td>`).join('')}<td class="num"><b>${brl(Object.values(r.m).reduce((a, b) => a + b, 0))}</b></td></tr>`).join('')}</tbody>
      <tfoot><tr><td>Total</td>${comps.map(c => `<td class="num">${brl(rows.reduce((a, r) => a + (r.m[c] || 0), 0))}</td>`).join('')}<td class="num">${brl(rows.reduce((a, r) => a + Object.values(r.m).reduce((x, y) => x + y, 0), 0))}</td></tr></tfoot>`
      : '<tbody><tr><td class="empty">Nenhuma competência fechada ainda.</td></tr></tbody>';
    loadAuditoria();
  } catch (e) { toast(e.message); }
}
async function loadVersoes(id, comp) {
  try {
    const vs = await q(sb.from('conciliacao_versoes').select('id,versao,tipo,totais,observacao,salvo_por_email,salvo_em').eq('conciliacao_id', id).order('versao', { ascending: false }), 'Versões');
    $('#cardVers').classList.remove('hide'); $('#versTitle').textContent = `Versões de ${compLabel(comp)}`;
    const TP = { salvamento: 'Salvamento', fechamento: '🔒 Fechamento', reabertura: '↺ Reabertura' };
    $('#tblVers').innerHTML = `<colgroup><col style="width:7%"><col style="width:14%"><col style="width:15%"><col style="width:18%"><col style="width:12%"><col style="width:22%"><col style="width:12%"></colgroup>
      <thead><tr><th class="num">Versão</th><th>Tipo</th><th>Quando</th><th>Quem</th><th class="num">Comissão</th><th>Observação</th><th></th></tr></thead>
      <tbody>${vs.length ? vs.map(v => `<tr><td class="num">${v.versao}</td><td>${TP[v.tipo]}</td><td class="muted" style="font-size:12px">${dtf(v.salvo_em)}</td><td title="${esc(v.salvo_por_email)}">${esc(v.salvo_por_email || '—')}</td>
        <td class="num">${brl(v.totais?.comissao)}</td><td title="${esc(v.observacao)}">${esc(v.observacao || '—')}</td>
        <td style="text-align:right">${v.tipo !== 'reabertura' ? `<button class="btn sm" data-xv="${v.id}">⬇ Excel</button>` : ''}</td></tr>`).join('') : '<tr><td colspan="7" class="empty">Nenhuma versão gravada. Use "Salvar versão" ou "Fechar competência".</td></tr>'}</tbody>`;
    $('#tblVers').querySelectorAll('[data-xv]').forEach(b => b.onclick = async () => {
      try {
        const v = await q(sb.from('conciliacao_versoes').select('*').eq('id', b.dataset.xv).single(), 'Versão');
        exportarExcel({ resumo: v.resumo || [], linhas: v.detalhe || null }, `Conciliacao_${compLabel(comp).replace('/', '_')}_v${v.versao}.xlsx`);
      } catch (e) { toast(e.message); }
    });
    $('#cardVers').scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (e) { toast(e.message); }
}
document.querySelectorAll('#segAud button').forEach(b => b.onclick = () => { document.querySelectorAll('#segAud button').forEach(x => x.classList.toggle('on', x === b)); S.audFiltro = b.dataset.a; loadAuditoria(); });
const TAB_LBL = { regras_procedimento: 'Regra', solicitacoes: 'Solicitação', conciliacoes: 'Conciliação', conciliacao_arquivos: 'Arquivo', dentistas: 'Dentista', especialidades: 'Especialidade' };
const IGN = new Set(['atualizado_em', 'atualizado_por', 'criado_em', 'criado_por', 'totais', 'id']);
function descAud(a) {
  const d = a.depois || {}, b = a.antes || {};
  const ref = a.tabela === 'conciliacoes' ? compLabel(d.competencia || b.competencia) : a.tabela === 'conciliacao_arquivos' ? (d.nome || b.nome) : (d.nome || b.nome);
  if (a.acao === 'insert') return `${ref}: criado`;
  const ch = Object.keys(d).filter(k => !IGN.has(k) && JSON.stringify(d[k]) !== JSON.stringify(b[k]));
  const fmt = k => {
    if (k === 'decisoes') { const c = x => Object.keys(x?.resolucoes || {}).length + Object.keys(x?.lancamentos || {}).length + Object.keys(x?.overrides || {}).length; return `decisões (${c(b.decisoes)} → ${c(d.decisoes)})`; }
    if (k === 'status') return `status ${b.status} → ${d.status}`;
    if (k === 'ativo') return d.ativo ? 'reativado' : 'inativado/removido';
    if (k === 'dentista_id') return `dentista → ${dentById(d.dentista_id)?.nome || '—'}`;
    if (['valor', 'forma', 'nome', 'aliases', 'versao'].includes(k)) return `${k}: ${b[k] ?? '—'} → ${d[k] ?? '—'}`;
    return k;
  };
  return `${ref}: ${ch.map(fmt).join(', ') || 'totais recalculados'}`;
}
async function loadAuditoria() {
  try {
    let qy = sb.from('auditoria').select('*').order('em', { ascending: false }).limit(80);
    if (S.audFiltro) qy = qy.eq('tabela', S.audFiltro);
    const rows = await q(qy, 'Auditoria');
    $('#tblAud').innerHTML = `<colgroup><col style="width:14%"><col style="width:13%"><col style="width:20%"><col style="width:53%"></colgroup>
      <thead><tr><th>Quando</th><th>Onde</th><th>Quem</th><th>O que mudou</th></tr></thead>
      <tbody>${rows.length ? rows.map(a => `<tr><td class="muted" style="font-size:12px">${dtf(a.em)}</td><td>${TAB_LBL[a.tabela] || a.tabela}</td><td title="${esc(a.usuario_email)}">${esc(a.usuario_email || 'sistema')}</td><td title="${esc(descAud(a))}">${esc(descAud(a))}</td></tr>`).join('') : '<tr><td colspan="4" class="empty">Sem registros.</td></tr>'}</tbody>`;
  } catch (e) { toast(e.message); }
}
})();
