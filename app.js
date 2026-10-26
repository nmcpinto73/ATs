/* Autorização de Trabalho — lógica replicada do Autorizacoes_de_Trabalho.xlsm
   Fórmulas: Risco = Severidade × Exposição × Controlo; Global: >=33 Alto, >=21 Médio, senão Baixo
   Numeração: <prefixo>-<ano>-<nº 4 dígitos>  (Config!C3/C4/C5)
   Macros VBA replicadas: EmitirAutorizacao, FecharAutorizacao, OrdenarHistorico,
   VerificarAlertas, NovaAutorizacao, AddSel, Worksheet_Change (alerta espaço confinado). */

'use strict';

const SEV = [1,2,3,4,5];            // lstSever: Listas!K3:K7
const EXPO = [1,2,3,4];             // lstExpo: Listas!L3:L6
const CTRL = [1,2,3];               // lstControl: Listas!M3:M5
const SIMNAO = ['Sim','Não'];       // lstSN

const DEFAULT_LISTS = {
  solicitantes: ['António Monteiro','António Pereira','Nuno Pinto','Frederico Magro','Paulo Casanova','Paulo Oliveira','Ricardo Carrilho','Ricardo Pereira','João Carrilho','Vitor Monteiro'],
  empresas: ['A.M.S.','Cimontubo','Meivecore','Nova Citacor','Soandaimes','Termiso'],
  locais: ['Fábrica de Resinas','Fábrica de Formaldeído','Armazenagem de Resinas','Armazenagem de Formaldeído','Geral','Armazém de Ureia','Armazém de Embalados','Silo Melamina','ETAI','ETAR','Central Térmica','Turbina','RIA','TK de Metanol','TK de Metanol A.P.S.'],
  ferramentas: ['Ferramentas manuais','Máquina pressão 200 Bar','Máquina pressão 500 Bar','Máquina pressão 1000 Bar','Martelo pneumático','Rebarbadora','Máquina de soldar','Berbequim','Chave de impacto','Escada','Andaime','Empilhador','Plataforma elevatória','Maçarico','Ferramenta antifaísca','Multímetro','Chave dinamométrica','Ferramentas para pintura'],
  quimicos: ['Resina','Resina residuos','Ureia','Melamina','Soda cáustica','Ácido acético','Formaldeído','Metanol','Solventes orgânicos','Ácido sulfúrico','Tintas e vernizes','Gases comprimidos','Óleos e lubrificantes','Amónia','Cloro','Combustíveis'],
  trabalhadores: ['Cristian Dobrinescu','Cosmin Simion','Carlos Salgueiro','Paulo Gonçalves','Pedro Matondo','Érica','Rafael','Francisco Barrinho','Paulo Parreira','José Aguamel','António Dias','Ivo Moreira','Lucas','Daniel Peres','Paulo Pinela','Rafael Klementsov','Ricardo Monteiro','António Monteiro','António Pereira','Nuno Pinto','Frederico Magro','Paulo Casanova'],
  riscos: ['Afogamento','Acesso inadequado','Ambiente térmico','Arranque do equipamento ou exposição a energia eletrica','Arranque do equipamento ou exposição a energia gravitica','Arranque do equipamento ou exposição a energia hidrálica','Arranque do equipamento ou exposição a energia Mecânica','Arranque do equipamento ou exposição a energia Pneumatica','Arranque do equipamento ou exposição a energia radioativa','Atmosferas perigosas','Atropelamento','Contato cutâneo com agente químico (Irritação cutânea, dermatites)','Contato térmico (Queimadura)','Corte / Entalamento / Esmagamento','Electrização / Electrocussão','Exposição a agente quÍmico (vapores / poeiras/gases)','Exposição a Formaldeído','Exposição a Monóxido de Carbono','Exposição a radiações','Exposição a vibrações','Exposição ao Fenol','Exposição ao metanol','Exposição ao ruído','Exposição Soda Caustica','Gases ou pressões não inflamáveis acima de 2 bar?','Iluminação inadequada','Incêndio / Explosão','Intoxicação por ingestão','Lesões musculo-esqueléticas','Produtos inflamáveis/ areas ATEX','Projeção de particulas/ fragmentos','Qualquer energia que pode pôr o equipamento a funcionar (por exemplo, gravidade)?','Queda a diferentes níveis','Queda ao mesmo nivel','Queda de objetos','Queda de objetos em manipulação','Queda em altura','Sistema elétrico acima de 50 V?','Soterramento / Esmagamento','Temperaturas acima de 50 ºC?','Vapor ou líquido criogénico?','Risco Ambiental','Risco Biológico','Ficha de procedimentos de Segurança'],
};

const CONFIG_KEY = 'at_config', LISTS_KEY = 'at_lists', HIST_KEY = 'at_hist', DRAFT_KEY = 'at_draft', NUM_KEY = 'at_nextnum';

/* ---------- Núcleo de cálculo (réplica das fórmulas do Excel) ---------- */
// F39: =IF(OR(C39="",D39="",E39=""),"",C39*D39*E39)
function nivelRisco(sev, expo, ctrl){
  if (sev==null || expo==null || ctrl==null || sev==='' || expo==='' || ctrl==='') return null;
  return Number(sev) * Number(expo) * Number(ctrl);
}
function corRisco(v){ if (v==null) return ''; return v>=33 ? 'r' : v>=21 ? 'y' : 'g'; }
// D46: =IF(COUNT(F39:F44)=0,"",IF(MAX(F39:F44)>=33,"Alto",IF(MAX(F39:F44)>=21,"Médio","Baixo")))
function riscoGlobal(valores){
  const nums = valores.filter(v => v !== null && v !== '');
  if (!nums.length) return null;
  const mx = Math.max(...nums);
  return mx >= 33 ? 'Alto' : mx >= 21 ? 'Médio' : 'Baixo';
}
// C6 via macro: prefixo-ano-0000
function formatNum(prefixo, ano, n){ return `${String(prefixo).trim()}-${ano}-${String(n).padStart(4,'0')}`; }

/* ---------- Estado ---------- */
let state = {
  cfg: { prefixo:'AT', ano:'2026', nextnum:1, empresa:'Nuno Pinto & Excel' },
  lists: JSON.parse(JSON.stringify(DEFAULT_LISTS)),
  hist: [],
};
let draft = newDraft();
function newDraft(){ return {
  solicitante:'', empresa:'', local:'', equip:'', desc:'',
  atex:'Não', altura:'Não', quente:'Não', loto:'Não', confinados:'Não', escavacoes:'Não', pressao:'Não',
  ferr:[], quim:[], trab:[],
  riscos: Array.from({length:6},()=>({p:'',s:'',e:'',c:'',m:''})),
  lotoRows: Array.from({length:3},()=>({ex:'',id:'',co:'',ob:''})),
  numero:null, estado:'-', emissao:null,
};}

function save(){
  localStorage.setItem(CONFIG_KEY, JSON.stringify(state.cfg));
  localStorage.setItem(LISTS_KEY, JSON.stringify(state.lists));
  localStorage.setItem(HIST_KEY, JSON.stringify(state.hist));
  localStorage.setItem(NUM_KEY, JSON.stringify(state.cfg.nextnum));
  localStorage.setItem(DRAFT_KEY, JSON.stringify(draft));
}
function load(){
  try{ const c = localStorage.getItem(CONFIG_KEY); if(c) state.cfg = {...state.cfg, ...JSON.parse(c)}; }catch(e){}
  try{ const l = localStorage.getItem(LISTS_KEY); if(l) state.lists = {...JSON.parse(JSON.stringify(DEFAULT_LISTS)), ...JSON.parse(l)}; }catch(e){}
  try{ const h = localStorage.getItem(HIST_KEY); if(h) state.hist = JSON.parse(h); }catch(e){}
  try{ const d = localStorage.getItem(DRAFT_KEY); if(d) draft = {...newDraft(), ...JSON.parse(d)}; }catch(e){}
}

/* ---------- Utilitários ---------- */
const $ = s => document.querySelector(s);
const esc = s => String(s??'').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
function fmtDate(ts){ if(!ts) return ''; const d=new Date(ts); return d.toLocaleDateString('pt-PT')+' '+d.toLocaleTimeString('pt-PT',{hour:'2-digit',minute:'2-digit'}); }
function toast(msg){ const t=$('#toast'); t.textContent=msg; t.classList.add('show'); clearTimeout(t._h); t._h=setTimeout(()=>t.classList.remove('show'),2600); }
function fillSelect(sel, items, placeholder){
  sel.innerHTML = `<option value="">${placeholder||'— selecionar —'}</option>` + items.map(i=>`<option${i===sel.value?' selected':''}>${esc(i)}</option>`).join('');
}

/* ---------- Modal (substitui MsgBox do VBA) ---------- */
function showModal(title, msg, buttons){
  $('#modal-title').textContent = title;
  $('#modal-msg').textContent = msg;
  const bx = $('#modal-btns'); bx.innerHTML='';
  buttons.forEach(b=>{
    const el=document.createElement('button');
    el.className='btn '+(b.style||'primary'); el.textContent=b.label;
    el.onclick=()=>{ $('#modal-root').hidden=true; if(b.fn) b.fn(); };
    bx.appendChild(el);
  });
  $('#modal-root').hidden=false;
}

/* ---------- Formulário ---------- */
function renderForm(){
  $('#hdr-company').textContent = state.cfg.empresa;
  fillSelect($('#f-solicitante'), state.lists.solicitantes);
  fillSelect($('#f-empresa'), state.lists.empresas);
  fillSelect($('#f-local'), state.lists.locais);
  fillSelect($('#sel-ferr'), state.lists.ferramentas);
  fillSelect($('#sel-quim'), state.lists.quimicos);
  fillSelect($('#sel-trab'), state.lists.trabalhadores);

  $('#f-solicitante').value = draft.solicitante||'';
  $('#f-empresa').value = draft.empresa||'';
  $('#f-local').value = draft.local||'';
  $('#f-equip').value = draft.equip||'';
  $('#f-desc').value = draft.desc||'';
  document.querySelectorAll('.seg').forEach(seg=>{
    const f = seg.dataset.field;
    seg.querySelectorAll('button').forEach(b=>b.classList.toggle('on', b.dataset.val === draft[f]));
  });
  renderChips();
  renderRisks();
  renderLoto();
  renderSigns();
  updateStatusStrip();
  updateGlobalBadge();
}
function renderChips(){
  const defs = [['ferr','#chips-ferr'],['quim','#chips-quim'],['trab','#chips-trab']];
  defs.forEach(([k,sel])=>{
    $(sel).innerHTML = (draft[k]||[]).map((v,i)=>`<span class="chip">${esc(v)}<button type="button" data-k="${k}" data-i="${i}">✕</button></span>`).join('') || '<span class="note">Nenhum item adicionado.</span>';
    $(sel).querySelectorAll('button').forEach(b=>b.onclick=()=>{ draft[b.dataset.k].splice(Number(b.dataset.i),1); renderChips(); save(); });
  });
}
function renderRisks(){
  const box = $('#risk-rows'); box.innerHTML='';
  draft.riscos.forEach((r,i)=>{
    const v = nivelRisco(r.s, r.e, r.c);
    const div=document.createElement('div'); div.className='risk-row';
    div.innerHTML = `
      <div class="field" style="margin-bottom:0"><label>Perigo / Risco ${i+1}</label>
      <select data-i="${i}" data-f="p">${'<option value="">— selecionar —</option>' + state.lists.riscos.map(x=>`<option${x===r.p?' selected':''}>${esc(x)}</option>`).join('')}</select></div>
      <div class="rgrid">
        <div><label>Severidade (1–5)</label><select data-i="${i}" data-f="s">${opts(r.s, SEV)}</select></div>
        <div><label>Exposição (1–4)</label><select data-i="${i}" data-f="e">${opts(r.e, EXPO)}</select></div>
        <div><label>Controlo (1–3)</label><select data-i="${i}" data-f="c">${opts(r.c, CTRL)}</select></div>
        <div><label>Nível de Risco</label><div class="rv ${corRisco(v)}">${v==null?'—':v}</div></div>
      </div>
      <div class="field" style="margin:8px 0 0"><label>Medidas de controlo</label><input type="text" data-i="${i}" data-f="m" value="${esc(r.m)}" placeholder="Medidas preventivas..."></div>`;
    box.appendChild(div);
  });
  box.querySelectorAll('select,input').forEach(el=>{
    el.onchange = ()=>{ draft.riscos[Number(el.dataset.i)][el.dataset.f] = el.value; save(); renderRisks(); updateGlobalBadge(); };
  });
}
function opts(cur, arr){ return '<option value="">—</option>'+arr.map(v=>`<option value="${v}"${String(cur)===String(v)?' selected':''}>${v}</option>`).join(''); }

function renderLoto(){
  const box = $('#loto-rows'); box.innerHTML='';
  draft.lotoRows.forEach((r,i)=>{
    const div=document.createElement('div'); div.className='loto-row';
    div.innerHTML = `
      <div class="row2">
        <div class="field"><label>Executante</label>${inputLoto(i,'ex',r.ex,'Nome')}</div>
        <div class="field"><label>ID do Cadeado</label>${inputLoto(i,'id',r.id,'Ex.: CAD-014')}</div>
      </div>
      <div class="row2">
        <div class="field"><label>Consignado?</label><select data-i="${i}" data-f="co">${optsSimNao(r.co)}</select></div>
        <div class="field"><label>Observações</label>${inputLoto(i,'ob',r.ob,'')}</div>
      </div>`;
    box.appendChild(div);
  });
  box.querySelectorAll('input,select').forEach(el=>{
    el.onchange = ()=>{ draft.lotoRows[Number(el.dataset.i)][el.dataset.f] = el.value; save(); };
  });
}
function inputLoto(i,f,val,ph){ return `<input type="text" data-i="${i}" data-f="${f}" value="${esc(val)}" placeholder="${ph}">`; }
function optsSimNao(cur){ return '<option value="">—</option><option value="Sim"'+(cur==='Sim'?' selected':'')+'>Sim</option><option value="Não"'+(cur==='Não'?' selected':'')+'>Não</option>'; }

function renderSigns(){
  const g = riscoGlobal(draft.riscos.map(r=>nivelRisco(r.s,r.e,r.c)))==='Alto';
  $('#sign-row').innerHTML = `
    <div class="sign">_______________________<br><b>O Solicitante</b></div>
    ${g?'<div class="sign sst">_______________________<br><b>Responsável SST</b><br><small>Nível de Risco Elevado requer SST</small></div>':''}
    <div class="sign">_______________________<br><b>Empresa Executante</b></div>`;
}
function updateStatusStrip(){
  $('#st-num').textContent = draft.numero || '(automático)';
  $('#st-estado').textContent = draft.estado || '-';
  $('#st-data').textContent = draft.emissao ? 'Emitida: '+fmtDate(draft.emissao) : '';
}
function updateGlobalBadge(){
  const g = riscoGlobal(draft.riscos.map(r=>nivelRisco(r.s,r.e,r.c)));
  const b = $('#global-badge');
  b.textContent = g || '—';
  b.className = 'badge ' + (g ? 'bg-'+g.toLowerCase() : 'bg-none');
}

/* --- Botões + Adicionar (AddSel do VBA: adiciona sem repetidos) --- */
function bindAdd(btnId, selId, key){
  $(btnId).onclick = ()=>{
    const v = $(selId).value;
    if(!v){ toast('Escolha primeiro um item na lista.'); return; }
    if(!(draft[key]||[]).includes(v)) draft[key].push(v);
    $(selId).value=''; save(); renderChips();
  };
}

/* --- Segmentos Sim/Não --- */
document.querySelectorAll('.seg').forEach(seg=>{
  seg.querySelectorAll('button').forEach(b=>b.onclick=()=>{
    const f = seg.dataset.field;
    draft[f] = b.dataset.val;
    save(); renderForm();
    // Worksheet_Change do VBA: alerta imediato de espaço confinado
    if(f==='confinados' && b.dataset.val==='Sim'){
      showModal('⚠️ Espaço Confinado',
        'ALERTA DE SEGURANÇA: ESPAÇO CONFINADO.\n\nRequer autorização de entrada específica, monitorização atmosférica contínua, vigia permanente no exterior e meios de resgate disponíveis ANTES do início dos trabalhos.',
        [{label:'Compreendi'}]);
    }
  });
});

/* --- Emitir (EmitirAutorizacao do VBA) --- */
$('#btn-emitir').onclick = ()=>{
  if(!draft.solicitante || !draft.empresa || !draft.local || !draft.desc){
    showModal('Campos em falta','Preencha Solicitante, Empresa, Local e Descrição antes de emitir.',[{label:'OK'}]); return;
  }
  const g = riscoGlobal(draft.riscos.map(r=>nivelRisco(r.s,r.e,r.c)));
  // Regra LOTO da macro: se LOTO=Sim e risco ALTO → exigir executante + cadeado
  if(g==='Alto' && draft.loto==='Sim'){
    const ok = draft.lotoRows.some(r=>r.ex.trim() && r.id.trim());
    if(!ok){
      showModal('Detalhe LOTO',"LOTO assinalado como 'Sim': indique pelo menos o executante e o ID do cadeado na secção DETALHE LOTO.",[{label:'OK'}]);
      $('#loto-card').scrollIntoView({behavior:'smooth'});
      return;
    }
  }
  // Registar no histórico (linha do Histórico)
  const num = formatNum(state.cfg.prefixo, state.cfg.ano, state.cfg.nextnum);
  const reg = {
    num, estado:'Aberta', emissao: new Date().toISOString(),
    solicitante:draft.solicitante, empresa:draft.empresa, local:draft.local, equip:draft.equip, desc:draft.desc,
    atex:draft.atex, altura:draft.altura, loto:draft.loto, quente:draft.quente, confinados:draft.confinados,
    escavacoes:draft.escavacoes, pressao:draft.pressao,
    ferr:draft.ferr.join('; '), quim:draft.quim.join('; '), trab:draft.trab.join('; '),
    risco:g, fecho:null,
    lotoExec:draft.lotoRows.filter(r=>r.ex).map(r=>r.ex).join('; '),
    lotoIds:draft.lotoRows.filter(r=>r.id).map(r=>r.id).join('; '),
    lotoConsign:draft.lotoRows.filter(r=>r.co).map(r=>r.co).join('; '),
    lotoObs:draft.lotoRows.filter(r=>r.ob).map(r=>r.ob).join('; '),
    riscosDet: draft.riscos.filter(r=>r.p && r.s && r.e && r.c).map(r=>({p:r.p,s:r.s,e:r.e,c:r.c,m:r.m})),
  };
  state.hist.unshift(reg);                       // OrdenarHistorico: mais recente primeiro
  state.cfg.nextnum += 1;
  draft.numero = num; draft.estado='Aberta'; draft.emissao = reg.emissao;
  save(); renderForm();

  const msgs = [`Autorização ${num} emitida e guardada no histórico (estado: Aberta).`];
  if(g==='Alto') msgs.push('ATENÇÃO: Risco GLOBAL ALTO. Requer medidas adicionais e aprovação SST.');
  if(draft.confinados==='Sim') msgs.push('Lembrete: esta autorização envolve ESPAÇO CONFINADO.\nGaranta autorização de entrada, monitorização atmosférica, vigia e plano de resgate.');
  showModal(g==='Alto'?'⚠️ Risco elevado':'✔ Emitida', msgs.join('\n\n'), [{label:'OK'}]);
};

/* --- Nova (NovaAutorizacao do VBA) --- */
$('#btn-nova').onclick = ()=>{
  showModal('Nova autorização','Limpar o formulário para emitir uma nova autorização? (O histórico não é afetado.)',[
    {label:'Cancelar', style:'ghost'},
    {label:'Limpar', fn:()=>{ draft = newDraft(); save(); renderForm(); toast('Formulário limpo. Pode preencher uma nova autorização.'); }},
  ]);
};

/* ---------- Histórico ---------- */
function renderHist(){
  const q = ($('#hist-search').value||'').toLowerCase();
  const box = $('#hist-list');
  const items = state.hist.filter(h=>!q || [h.num,h.local,h.empresa,h.solicitante,h.risco,h.desc].join(' ').toLowerCase().includes(q));
  box.innerHTML = items.length ? '' : '<div class="note">Sem registos. Emita a primeira autorização no separador «Nova».</div>';
  items.forEach(h=>{
    const div=document.createElement('div'); div.className='hist-item';
    div.innerHTML = `
      <div class="hist-top"><span class="num">${esc(h.num)}</span>
        <span><span class="badge ${h.estado.toLowerCase()}">${esc(h.estado)}</span>
        ${h.risco?`<span class="badge ${h.risco.toLowerCase()==='médio'?'medio':h.risco.toLowerCase()}">${esc(h.risco)}</span>`:''}</span></div>
      <div class="hist-meta">
        <span>📅 ${fmtDate(h.emissao)}${h.fecho?' → fechada '+fmtDate(h.fecho):''}</span>
        <span>👤 ${esc(h.solicitante)} · 🏢 ${esc(h.empresa)}</span>
        <span>📍 ${esc(h.local)}${h.equip?' · '+esc(h.equip):''}</span>
        <span>📝 ${esc(String(h.desc).slice(0,90))}${String(h.desc).length>90?'…':''}</span>
      </div>
      <div class="hist-actions">
        <button class="btn ghost small" data-act="ver">👁 Ver / Imprimir</button>
        ${h.estado==='Aberta'?'<button class="btn success small" data-act="fechar">🔒 Fechar</button>':''}
      </div>`;
    div.querySelector('[data-act="ver"]').onclick = ()=>showDetail(h);
    const fb = div.querySelector('[data-act="fechar"]');
    if(fb) fb.onclick = ()=>fecharAut(h);
    box.appendChild(div);
  });
}
$('#hist-search').oninput = renderHist;

/* --- Fechar (FecharAutorizacao + frmFecharAut do VBA) --- */
function fecharAut(h){
  if(h.estado!=='Aberta'){ toast('Esta autorização já não está aberta.'); return; }
  showModal('Confirmar fecho', `Pretende fechar a autorização ${h.num}?`, [
    {label:'Não', style:'ghost'},
    {label:'Sim, fechar', fn:()=>{
      h.estado='Fechada'; h.fecho=new Date().toISOString();
      if(draft.numero===h.num){ draft.estado='Fechada'; }
      save(); renderHist(); renderForm(); toast(`Autorização ${h.num} fechada.`);
    }},
  ]);
}

/* --- Detalhe / impressão (ImprimirAutorizacao do VBA) --- */
function showDetail(h){
  const rows = (h.riscosDet||[]).map(r=>{
    const v = nivelRisco(r.s,r.e,r.c);
    return `<tr><td>${esc(r.p)}</td><td>${r.s}</td><td>${r.e}</td><td>${r.c}</td><td style="font-weight:800">${v}</td><td>${esc(r.m)}</td></tr>`;
  }).join('') || '<tr><td colspan="6">Sem riscos avaliados.</td></tr>';
  const loto = (h.lotoExec||h.lotoIds) ? `
    <table class="dtable"><tr><td class="k">Executantes LOTO</td><td>${esc(h.lotoExec)}</td></tr>
    <tr><td class="k">IDs dos Cadeados</td><td>${esc(h.lotoIds)}</td></tr>
    <tr><td class="k">Consignado (LOTO)</td><td>${esc(h.lotoConsign)}</td></tr>
    <tr><td class="k">Observações</td><td>${esc(h.lotoObs)}</td></tr></table>` : '';
  $('#detail-sheet').innerHTML = `
    <h2>AUTORIZAÇÃO DE TRABALHO</h2>
    <div class="sub">${esc(state.cfg.empresa)} — Documento de controlo de segurança</div>
    <table class="dtable">
      <tr><td class="s">Nº</td><td><b>${esc(h.num)}</b></td><td class="s">Estado</td><td>${esc(h.estado)}</td></tr>
      <tr><td class="s">Data de emissão</td><td>${fmtDate(h.emissao)}</td><td class="s">Data de fecho</td><td>${h.fecho?fmtDate(h.fecho):'—'}</td></tr>
      <tr><td class="s">Solicitante</td><td>${esc(h.solicitante)}</td><td class="s">Empresa executante</td><td>${esc(h.empresa)}</td></tr>
      <tr><td class="s">Local do trabalho</td><td>${esc(h.local)}</td><td class="s">Equipamento</td><td>${esc(h.equip)||'—'}</td></tr>
      <tr><td class="k">Descrição dos trabalhos</td><td colspan="3">${esc(h.desc)}</td></tr>
    </table>
    <table class="dtable">
      <tr><td class="s">Zona ATEX</td><td>${esc(h.atex)}</td><td class="s">Trab. em altura</td><td>${esc(h.altura)}</td></tr>
      <tr><td class="s">Trab. a quente</td><td>${esc(h.quente)}</td><td class="s">Esp. confinados</td><td>${esc(h.confinados)}</td></tr>
      <tr><td class="s">Escavações</td><td>${esc(h.escavacoes)}</td><td class="s">Máq. de pressão</td><td>${esc(h.pressao)}</td></tr>
      <tr><td class="k">LOTO (Bloq./Etiq.)</td><td colspan="3">${esc(h.loto)}</td></tr>
      <tr><td class="k">Ferramentas</td><td colspan="3">${esc(h.ferr)||'—'}</td></tr>
      <tr><td class="k">Produtos químicos</td><td colspan="3">${esc(h.quim)||'—'}</td></tr>
      <tr><td class="k">Trabalhadores</td><td colspan="3">${esc(h.trab)||'—'}</td></tr>
    </table>
    <b style="font-size:13px">⚠️ Avaliação de riscos — Nível global: <span class="badge ${h.risco?h.risco.toLowerCase():''}">${esc(h.risco||'—')}</span></b>
    <table class="risks-d"><thead><tr><th>Perigo / Risco</th><th>Sev.</th><th>Exp.</th><th>Contr.</th><th>Risco</th><th>Medidas</th></tr></thead><tbody>${rows}</tbody></table>
    ${loto}
    <div class="sign-row" style="margin-top:14px">
      <div class="sign">_______________________<br><b>O Solicitante</b></div>
      ${h.risco==='Alto'?'<div class="sign sst">_______________________<br><b>Responsável SST</b><br><small>Nível de Risco Elevado requer SST</small></div>':''}
      <div class="sign">_______________________<br><b>Empresa Executante</b></div>
    </div>
    <div class="detail-actions">
      <button class="btn primary small" onclick="window.print()">🖨 Imprimir / PDF</button>
      <button class="btn ghost small" id="detail-close">Voltar</button>
    </div>`;
  $('#detail-root').hidden=false;
  $('#detail-close').onclick = ()=>{ $('#detail-root').hidden=true; };
}

/* --- Exportar CSV --- */
$('#btn-csv').onclick = ()=>{
  const cols = ['num','estado','emissao','solicitante','empresa','local','desc','atex','altura','loto','quente','confinados','escavacoes','pressao','ferr','quim','trab','risco','fecho','lotoExec','lotoIds','lotoConsign','lotoObs','equip'];
  const head = ['Nº Autorização','Estado','Data Emissão','Solicitante','Empresa Executante','Local','Descrição','ATEX','Trab. em Altura','LOTO','Trab. a Quente','Esp. Confinados','Escavações','Máq. de Pressão','Ferramentas','Produtos Químicos','Trabalhadores','Nível de Risco','Data Fecho','Executantes LOTO','IDs Cadeados','Consignado','Observações LOTO','Equipamento'];
  const lines = [head.join(';')].concat(state.hist.map(h=>cols.map(c=>String(h[c]??'').replaceAll(';',',')).join(';')));
  const blob = new Blob(['\ufeff'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'});
  const a = document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='historico_autorizacoes.csv'; a.click();
};

/* --- VerificarAlertas (macro Workbook_Open do VBA) --- */
function verificarAlertas(){
  const ab = state.hist.filter(h=>h.estado==='Aberta');
  const alto = ab.filter(h=>h.risco==='Alto').length;
  const conf = ab.filter(h=>h.confinados==='Sim').length;
  const loto = ab.filter(h=>h.loto==='Sim').length;
  const banner = $('#alerts-banner'), inner = $('#alerts-inner'), txt = $('#alerts-text');
  if(ab.length){
    const crit = alto>0;
    inner.classList.toggle('crit', crit);
    txt.textContent = `🔔 ${ab.length} autorização(ões) ABERTA(S) · ${alto} com risco ALTO · ${conf} em espaços confinados · ${loto} com LOTO`;
    banner.hidden=false;
  } else banner.hidden=true;
}
$('#btn-alertas').onclick = verificarAlertas;
$('#alerts-close').onclick = ()=>{ $('#alerts-banner').hidden=true; };

/* ---------- Painel (fórmulas COUNTIF do Dashboard) ---------- */
function renderDash(){
  const T = state.hist.length;
  const ab = state.hist.filter(h=>h.estado==='Aberta').length;
  const fe = T-ab;
  const alto = state.hist.filter(h=>h.risco==='Alto').length;
  const medio = state.hist.filter(h=>h.risco==='Médio').length;
  const baixo = state.hist.filter(h=>h.risco==='Baixo').length;
  const pct = n => T ? Math.round(n/T*100)+'%' : '0%';
  const byType = flag => {
    const tot = state.hist.filter(h=>h[flag]==='Sim').length;
    const aberto = state.hist.filter(h=>h[flag]==='Sim'&&h.estado==='Aberta').length;
    return {tot, aberto, pct:pct(tot)};
  };
  const kpis = [['TOTAL',T,''],['ABERTAS',ab,'abertas'],['FECHADAS',fe,'fechadas'],['RISCO ALTO',alto,'alto'],['RISCO MÉDIO',medio,'medio'],['RISCO BAIXO',baixo,'baixo']];
  $('#dash-kpis').innerHTML = kpis.map(([l,v,c])=>`<div class="kpi ${c}"><div class="v">${v}</div><div class="l">${l}</div></div>`).join('');
  const esp = [['💥 ATEX','atex'],['🚪 CONFINADOS','confinados'],['🔥 QUENTE','quente'],['🔒 LOTO','loto'],['🪜 ALTURA','altura'],['🌋 PRESSÃO','pressao']];
  $('#dash-especiais').innerHTML = esp.map(([l,f])=>{ const d=byType(f); return `<div class="esp ${d.aberto?'hot':''}"><div class="v">${d.tot}</div><div class="l">${l}<br><small>${d.aberto} aberto(s)</small></div></div>`; }).join('');
  const linhas = [
    ['Risco global: ALTO', alto, state.hist.filter(h=>h.risco==='Alto'&&h.estado==='Aberta').length],
    ['Risco global: Médio', medio, state.hist.filter(h=>h.risco==='Médio'&&h.estado==='Aberta').length],
    ['Risco global: Baixo', baixo, state.hist.filter(h=>h.risco==='Baixo'&&h.estado==='Aberta').length],
    ['Zona ATEX', byType('atex').tot, byType('atex').aberto],
    ['Trabalhos em altura', byType('altura').tot, byType('altura').aberto],
    ['Trabalhos a quente', byType('quente').tot, byType('quente').aberto],
    ['Bloqueio/Etiquetagem (LOTO)', byType('loto').tot, byType('loto').aberto],
    ['Espaços confinados', byType('confinados').tot, byType('confinados').aberto],
    ['Escavações', byType('escavacoes').tot, byType('escavacoes').aberto],
    ['Máquinas de pressão', byType('pressao').tot, byType('pressao').aberto],
  ];
  $('#dash-tabela tbody').innerHTML = linhas.map(([l,t,a])=>`
    <tr><td>${l}</td><td class="num">${t}</td><td class="num" style="${a>0?'color:var(--red)':''}">${a}</td><td class="num ${t===alto&&alto?'ralto':t===medio&&medio?'rmedio':t===baixo&&baixo?'rbaixo':''}">${pct(t)}</td></tr>`).join('');
}

/* ---------- Listas / Config ---------- */
function renderListas(){
  $('#cfg-prefixo').value = state.cfg.prefixo;
  $('#cfg-ano').value = state.cfg.ano;
  $('#cfg-num').value = state.cfg.nextnum;
  $('#cfg-empresa').value = state.cfg.empresa;
  ['cfg-prefixo','cfg-ano','cfg-num','cfg-empresa'].forEach(id=>{
    $('#'+id).onchange = ()=>{
      state.cfg.prefixo=$('#cfg-prefixo').value.trim()||'AT';
      state.cfg.ano=$('#cfg-ano').value.trim()||'2026';
      state.cfg.nextnum=parseInt($('#cfg-num').value,10)||1;
      state.cfg.empresa=$('#cfg-empresa').value;
      save(); renderForm();
    };
  });
  const defs = [['Solicitantes','solicitantes'],['Empresas','empresas'],['Locais','locais'],['Ferramentas','ferramentas'],['Produtos químicos','quimicos'],['Trabalhadores','trabalhadores'],['Perigos & Riscos','riscos']];
  $('#listas-edit').innerHTML = defs.map(([label,key])=>`
    <div class="lstblock"><h3>${label} (${state.lists[key].length})</h3>
      <div class="chips">${state.lists[key].map((v,i)=>`<span class="chip">${esc(v)}<button type="button" data-lk="${key}" data-i="${i}">✕</button></span>`).join('')||'<span class="note">Vazia.</span>'}</div>
      <div class="lstadd"><input type="text" placeholder="Adicionar novo item..." data-lkinput="${key}"><button class="btn ghost small" data-lkadd="${key}">+ Adicionar</button></div>
    </div>`).join('');
  document.querySelectorAll('#listas-edit [data-lk]').forEach(b=>b.onclick=()=>{
    state.lists[b.dataset.lk].splice(Number(b.dataset.i),1); save(); renderListas(); renderForm();
  });
  document.querySelectorAll('#listas-edit [data-lkadd]').forEach(b=>b.onclick=()=>{
    const inp = document.querySelector(`[data-lkinput="${b.dataset.lkadd}"]`);
    const v = inp.value.trim();
    if(v && !state.lists[b.dataset.lkadd].includes(v)){ state.lists[b.dataset.lkadd].push(v); save(); renderListas(); renderForm(); toast('Item adicionado à lista.'); }
  });
}

/* --- Backup JSON --- */
$('#btn-backup').onclick = ()=>{
  const blob = new Blob([JSON.stringify({cfg:state.cfg,lists:state.lists,hist:state.hist},null,2)],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='backup_autorizacoes.json'; a.click();
};
$('#file-import').onchange = e=>{
  const f = e.target.files[0]; if(!f) return;
  const r = new FileReader();
  r.onload = ()=>{ try{
      const d = JSON.parse(r.result);
      if(d.cfg) state.cfg={...state.cfg,...d.cfg};
      if(d.lists) state.lists={...state.lists,...d.lists};
      if(d.hist) state.hist=d.hist;
      save(); load(); renderAll(); toast('Dados importados com sucesso.');
    }catch(err){ toast('Ficheiro inválido.'); } };
  r.readAsText(f);
};

/* ---------- Navegação ---------- */
function show(view){
  ['form','hist','dash','listas'].forEach(v=>{
    document.getElementById('view-'+v).hidden = v!==view;
  });
  document.querySelectorAll('.navbtn').forEach(b=>b.classList.toggle('on', b.dataset.view===view));
  if(view==='hist') renderHist();
  if(view==='dash') renderDash();
  if(view==='listas') renderListas();
  window.scrollTo({top:0});
}
document.querySelectorAll('.navbtn').forEach(b=>b.onclick=()=>show(b.dataset.view));

function renderAll(){ renderForm(); verificarAlertas(); }

/* ---------- Init ---------- */
load(); renderAll();

/* ---------- PWA: service worker + instalação ---------- */
if('serviceWorker' in navigator && location.protocol.startsWith('http')){
  navigator.serviceWorker.register('sw.js').catch(()=>{});
}
let deferredPrompt=null;
window.addEventListener('beforeinstallprompt', e=>{ e.preventDefault(); deferredPrompt=e; $('#btn-install').hidden=false; });
$('#btn-install').onclick = async ()=>{ if(deferredPrompt){ deferredPrompt.prompt(); deferredPrompt=null; $('#btn-install').hidden=true; } };
