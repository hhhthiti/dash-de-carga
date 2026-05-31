/* ═══════════════════════════════════════════════════════
   SUPABASE REST
═══════════════════════════════════════════════════════ */
const SB_URL_DEFAULT = 'https://pwjatxqtkvwcmzmjjvbi.supabase.co/rest/v1';
const SB_KEY_DEFAULT = 'sb_publishable_bUPTDkrOzc0_I3xwNw15aA_lk76gg4w';
const SB_URL = localStorage.getItem('sb_url') || SB_URL_DEFAULT;
const SB_KEY = localStorage.getItem('sb_key') || SB_KEY_DEFAULT;
const HDR = {'Content-Type':'application/json','apikey':SB_KEY,'Authorization':'Bearer '+SB_KEY};
const SB_RETRY_MS = [300, 900, 1800];
const sbEq = (v)=>`eq.${encodeURIComponent(String(v??''))}`;
const sbIn = (vals)=>`in.(${(vals||[]).map(v=>`"${encodeURIComponent(String(v??''))}"`).join(',')})`;

async function sbFetch(url, options={}){
  let lastErr;
  for(let i=0;i<=SB_RETRY_MS.length;i++){
    try{
      const r=await fetch(url,options);
      if(r.ok) return r;
      const msg=await r.text();
      const isRetryable=r.status>=500||r.status===429;
      r._bodyText=msg;
      if(!isRetryable || i===SB_RETRY_MS.length) return r;
      await new Promise(res=>setTimeout(res,SB_RETRY_MS[i]));
    }catch(e){
      lastErr=e;
      if(i===SB_RETRY_MS.length) break;
      await new Promise(res=>setTimeout(res,SB_RETRY_MS[i]));
    }
  }
  throw lastErr || new Error('Falha de rede ao acessar Supabase.');
}

async function sbErrorText(r){
  return r._bodyText!==undefined ? r._bodyText : await r.text();
}

function sbMissingColumnFromError(msg){
  const text=String(msg||'');
  const parsed=(()=>{try{return JSON.parse(text);}catch(e){return null;}})();
  const fullMsg=parsed&&parsed.message?parsed.message:text;
  const m=fullMsg.match(/Could not find the '([^']+)' column/i);
  return m?m[1]:'';
}

function rowsWithoutColumn(rows, col){
  return rows.map(row=>{
    const copy={...row};
    delete copy[col];
    return copy;
  });
}

function queryWithoutSelectedColumn(qs, col){
  const parts=String(qs||'').split('&');
  let changed=false;
  const next=parts.map(part=>{
    if(!part.startsWith('select=')) return part;
    const selected=part.slice(7).split(',').filter(c=>c!==col);
    changed=true;
    return 'select='+selected.join(',');
  });
  return changed?next.join('&'):qs;
}

async function sbGet(table, qs=''){
  if(!SB_KEY || SB_KEY.includes('COLE_SUA_ANON_KEY_AQUI')){
    throw new Error('Chave Supabase não configurada. Defina localStorage sb_key com a ANON KEY JWT (Settings → API → anon public).');
  }
  if(!SB_URL || !SB_URL.includes('/rest/v1')){
    throw new Error('URL Supabase inválida. Defina localStorage sb_url com https://SEU-PROJETO.supabase.co/rest/v1');
  }
  let query=qs;
  const ignoredCols=new Set();
  while(true){
    const r=await sbFetch(`${SB_URL}/${table}?${query}`,{headers:HDR});
    if(r.ok) return r.json();
    const msg=await sbErrorText(r);
    const missing=sbMissingColumnFromError(msg);
    if(missing&&!ignoredCols.has(missing)&&String(query||'').includes('select=')){
      const next=queryWithoutSelectedColumn(query,missing);
      if(next!==query){
        ignoredCols.add(missing);
        query=next;
        continue;
      }
    }
    throw new Error(`GET ${table}: `+msg);
  }
}
async function sbUpsert(table, rows, onConflict=''){
  for(let i=0;i<rows.length;i+=30){
    let lote=rows.slice(i,i+30);
    const ignoredCols=new Set();
    while(true){
      const conflict=onConflict?`?on_conflict=${encodeURIComponent(onConflict)}`:'';
      const r=await sbFetch(`${SB_URL}/${table}${conflict}`,{
        method:'POST',
        headers:{...HDR,'Prefer':'resolution=merge-duplicates,return=minimal'},
        body:JSON.stringify(lote)
      });
      if(r.ok) break;
      const msg=await sbErrorText(r);
      const missing=sbMissingColumnFromError(msg);
      if(missing&&!ignoredCols.has(missing)&&lote.some(row=>Object.prototype.hasOwnProperty.call(row,missing))){
        ignoredCols.add(missing);
        lote=rowsWithoutColumn(lote,missing);
        continue;
      }
      throw new Error(`UPSERT ${table} lote ${i/30+1}: `+msg);
    }
  }
}
async function sbPatch(table, data, filters){
  const q=Object.entries(filters).map(([k,v])=>`${k}=${sbEq(v)}`).join('&');
  let payload={...data};
  const ignoredCols=new Set();
  while(Object.keys(payload).length){
    const r=await sbFetch(`${SB_URL}/${table}?${q}`,{method:'PATCH',headers:HDR,body:JSON.stringify(payload)});
    if(r.ok) return;
    const msg=await sbErrorText(r);
    const missing=sbMissingColumnFromError(msg);
    if(missing&&!ignoredCols.has(missing)&&Object.prototype.hasOwnProperty.call(payload,missing)){
      ignoredCols.add(missing);
      delete payload[missing];
      continue;
    }
    throw new Error(`PATCH ${table}: `+msg);
  }
}
async function sbDelete(table, filters){
  const q=Object.entries(filters).map(([k,v])=>`${k}=${sbEq(v)}`).join('&');
  const r=await sbFetch(`${SB_URL}/${table}?${q}`,{method:'DELETE',headers:HDR});
  if(!r.ok) throw new Error(`DELETE ${table}: `+await sbErrorText(r));
}
async function sbInsert(table, rows){
  let payload=rows;
  const ignoredCols=new Set();
  while(true){
    const r=await sbFetch(`${SB_URL}/${table}`,{method:'POST',headers:HDR,body:JSON.stringify(payload)});
    if(r.ok) return;
    const msg=await sbErrorText(r);
    const missing=sbMissingColumnFromError(msg);
    if(missing&&!ignoredCols.has(missing)&&payload.some(row=>Object.prototype.hasOwnProperty.call(row,missing))){
      ignoredCols.add(missing);
      payload=rowsWithoutColumn(payload,missing);
      continue;
    }
    throw new Error(`INSERT ${table}: `+msg);
  }
}

/* ═══════════════════════════════════════════════════════
   ESTADO
═══════════════════════════════════════════════════════ */
const STATUS_OPTIONS=['AG CHEGADA','CARREGANDO','EXPEDIDO','NO SHOW','VEICULO RECUSADO','SEPARANDO','EM FATURAMENTO','PATIO','DT EXCLUIDA','FOI EMBORA'];
const STATUS_FINAIS=['EXPEDIDO','NO SHOW','VEICULO RECUSADO'];
function isFinalStatus(status){
  return STATUS_FINAIS.includes(String(status||'').trim().toUpperCase());
}
const SC={'EXPEDIDO':'#22c55e','CARREGANDO':'#3b82f6','AG CHEGADA':'#f59e0b','NO SHOW':'#ef4444','VEICULO RECUSADO':'#dc2626','SEPARANDO':'#8b5cf6','EM FATURAMENTO':'#06b6d4','PATIO':'#64748b','DT EXCLUIDA':'#374151','FOI EMBORA':'#6b7280'};
// REGRA 9: DTs novas — badge NOVO até a próxima atualização da agenda
let novoDTs = new Set(); // DTs novas do último upload

function buildTipoOperacaoBadge(tipoOperacao,{small=false,emptyDash=true}={}){
  const op=String(tipoOperacao||'').toUpperCase();
  const style=small?' style="padding:3px 10px;font-size:11px;"':'';
  if(op.includes('TRANSFER')) return `<span class="tag-transf"${style}>🔄 TRANSFERÊNCIA</span>`;
  if(op.includes('VENDA ARUJA')) return `<span class="tag-venda"${style}>🏭 VENDA ARUJA</span>`;
  if(op.includes('VENDA MOGI')) return `<span class="tag-mogi"${style}>🏙️ VENDA MOGI</span>`;
  if(op.includes('VENDA')) return `<span class="tag-venda"${style}>🛒 VENDA</span>`;
  return emptyDash?'<span class="tag-sem">—</span>':'';
}

function buildPaletizacaoBadge(raw){
  const tipo=normalizePaletizacaoLabel(raw);
  if(tipo==='TORDESILHAS') return '<span class="tag-palet tag-tordesilhas">TORDESILHAS</span>';
  if(tipo==='PALETIZADA') return '<span class="tag-palet tag-paletizada">PALETIZADA</span>';
  return '<span class="tag-palet tag-estivada">ESTIVADA</span>';
}

let agendRows=[], agendaDiagRows=[], dtsMescladas=[], exportMap={}, tipoOpMap={}, descDocMap={}, centroMap={}, infoAgendaMap={}, tableData=[];
let remessaMap={}, pesoLiquidoMap={}, horaChegadaCSVMap={}, sapNumMap={}, paletizacaoMap={};
let preFatMode=false; // Ctrl+Ç ativa checkboxes de Pré-Fat nas DTs
let panelDT=null;
let currentTableTab='todas';
let reagendDT=null;
let dtSearchTerm='';
let tableSortMode=localStorage.getItem('table_sort_mode')||'inicio';
let activeRefsOverride=[];
let autoSyncPausedUntil=0;
let fileWorkflowInProgress=false;


function pauseAutoSync(ms=120000){
  autoSyncPausedUntil=Math.max(autoSyncPausedUntil,Date.now()+ms);
}

function releaseAutoSyncAfterSave(ms=45000){
  fileWorkflowInProgress=false;
  pauseAutoSync(ms);
}

function isVisible(id){
  const el=document.getElementById(id);
  return !!el && el.style.display!=='none';
}

function shouldSkipAutoSync(){
  const importModal=document.getElementById('import-overlay');
  return fileWorkflowInProgress ||
    Date.now()<autoSyncPausedUntil ||
    !isVisible('passo4') ||
    (importModal && importModal.style.display!=='none');
}

/* ═══════════════════════════════════════════════════════
   INIT
═══════════════════════════════════════════════════════ */
// Ctrl+Ç — ativa/desativa modo Pré-Fat
document.addEventListener('keydown', async e=>{
  if((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='q'){
    e.preventDefault();
    openAgendaSearchModal();
    return;
  }
  if(e.ctrlKey && (e.key==='ç'||e.key==='Ç'||e.keyCode===231||e.keyCode===199)){
    e.preventDefault();
    preFatMode=!preFatMode;
    document.body.classList.toggle('prefat-mode', preFatMode);
    showOk(preFatMode?'✅ Modo PRÉ-FAT ativado — marque as DTs pré-faturadas':'❌ Modo PRÉ-FAT desativado');
    renderRows();
  }
  // Ctrl+S — salva a grade (grade_carregamento + fim_carregamento) de todas as DTs no banco
  if((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='s'){
    e.preventDefault();
    if(!tableData.length) return;
    spin(true);
    showInf('Salvando grade no banco…');
    try{
      for(const row of tableData){
        if(!row.dt||!row.data_ref) continue;
        const diaAtual=diaRefAtualPorDataRef(row.data_ref)||row.dia_ref;
        row.dia_ref=diaAtual;
        await sbPatch('reporte_carga',{
          grade_carregamento:String(row.grade_carregamento||''),
          fim_carregamento:String(row.fim_carregamento||''),
          dia_ref:String(diaAtual||''),
          updated_at:new Date().toISOString(),
        },{dt:row.dt,data_ref:row.data_ref});
      }
      hideInf();
      showOk('✅ Grade salva no banco para '+tableData.length+' DT(s)!');
    }catch(err){
      hideInf();
      showErr('Erro ao salvar grade: '+err.message);
    }
    spin(false);
  }
});

window.addEventListener('DOMContentLoaded', async ()=>{
  const syncTxt = document.getElementById('sync-txt');
  const syncDot = document.getElementById('sync-dot');
  syncTxt.textContent = 'CONECTANDO…';
  if(syncDot){ syncDot.style.background='#f59e0b'; }
  try {
    await tryLoadExisting();
    syncTxt.textContent = 'ONLINE';
    if(syncDot){ syncDot.style.background='#22c55e'; }
    // Se não carregou nada do banco, mostra o passo 1 para upload da agenda
    if(!tableData.length) setStep(1);
  } catch(e) {
    syncTxt.textContent = 'ERRO';
    if(syncDot){ syncDot.style.background='#ef4444'; syncDot.style.animation='none'; }
    showErr('Falha ao conectar: ' + e.message);
    setStep(1);
  }
  // Ticker para atualizar emojis de relógio a cada minuto
  updateStatusReminder();
  setInterval(()=>{
    updateStatusReminder();
    if(!tableData.length) return;
    if(currentTableTab==='reporte') renderReporte();
    else if(currentTableTab==='todas'||currentTableTab==='finalizadas') renderRows();
  }, 60000);
  // Auto-sync para ambiente com múltiplos usuários
  setInterval(async ()=>{
    if(shouldSkipAutoSync()) return;
    try{
      await tryLoadExisting();
      syncTxt.textContent='ONLINE';
      if(syncDot){ syncDot.style.background='#22c55e'; syncDot.style.animation='pulse 1.5s infinite'; }
    }catch(e){
      syncTxt.textContent='OFFLINE';
      if(syncDot){ syncDot.style.background='#f59e0b'; }
    }
  }, 20000);
});

/* ═══════════════════════════════════════════════════════
   UI
═══════════════════════════════════════════════════════ */
function showErr(m){const e=document.getElementById('aerr');e.textContent='⚠️ '+m;e.style.display='block';}
function showOk(m) {const e=document.getElementById('aok'); e.textContent='✅ '+m;e.style.display='block';setTimeout(()=>e.style.display='none',4500);}
function showInf(m){const e=document.getElementById('ainf');e.textContent='⏳ '+m;e.style.display='block';}
function hideInf(){document.getElementById('ainf').style.display='none';}
function spin(on){document.getElementById('spinner').style.display=on?'block':'none';}

function escHtml(v){
  return String(v??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
}

function openAgendaSearchModal(){
  const ov=document.getElementById('agenda-search-overlay');
  if(!ov) return;
  ov.style.display='flex';
  const input=document.getElementById('agenda-search-input');
  if(input){
    input.focus();
    input.select();
  }
  renderAgendaSearch();
}

function closeAgendaSearchModal(){
  const ov=document.getElementById('agenda-search-overlay');
  if(ov) ov.style.display='none';
}

function agendaDiagnosticRowsForTerm(term,{dtOnly=false}={}){
  const q=normalizeDT(term).toLowerCase();
  if(!q) return [];
  return (agendaDiagRows||[]).filter(r=>{
    const dt=normalizeDT(r.DT).toLowerCase();
    if(dt===q || dt.includes(q)) return true;
    if(dtOnly) return false;
    return String(r.TRANSPORTADORA||'').toLowerCase().includes(q) ||
      String(r.DOCA||'').toLowerCase().includes(q);
  });
}

function agendaSearchTerms(raw){
  const text=String(raw||'').trim();
  if(!text) return [];
  const numericTerms=[...text.matchAll(/\d{5,}/g)]
    .map(m=>normalizeDT(m[0]))
    .filter(Boolean);
  const uniqueNumeric=[];
  numericTerms.forEach(dt=>{if(!uniqueNumeric.includes(dt)) uniqueNumeric.push(dt);});
  if(uniqueNumeric.length>1) return uniqueNumeric.map(value=>({value,dtOnly:true}));
  return [{value:text,dtOnly:false}];
}

function agendaDashRowsForTerm(term,{dtOnly=false}={}){
  const dtTerm=normalizeDT(term);
  if(!dtTerm) return [];
  return tableData.filter(r=>{
    const dt=String(r.dt||'');
    if(dt===dtTerm || dt.includes(dtTerm)) return true;
    if(dtOnly) return false;
    return String(r.transportadora||'').toLowerCase().includes(String(term).toLowerCase());
  });
}

function buildAgendaSearchMissingCard(term){
  return `<div style="color:#fca5a5;background:#450a0a;border:1px solid #ef444455;border-radius:8px;padding:12px;margin-bottom:10px;">
    Não achei a DT <b>${escHtml(term)}</b> na última agenda lida nem na tabela atual. Se a agenda foi carregada antes de abrir esta tela, importe o arquivo de agenda novamente e pesquise de novo.
  </div>`;
}

function buildAgendaSearchCard(dt,diagRows,dashRows){
  const rows=diagRows.filter(r=>normalizeDT(r.DT)===normalizeDT(dt));
  const inDash=dashRows.filter(r=>normalizeDT(r.dt)===normalizeDT(dt));
  const anyIncluded=rows.some(r=>r.included);
  const refLabels=[...new Set(rows.map(r=>agendaRefDate(r)).filter(Boolean).map(dKey))];
  const statusColor=inDash.length?'#22c55e':anyIncluded?'#f59e0b':'#ef4444';
  const statusText=inDash.length
    ? 'Entrou no dashboard atual'
    : anyIncluded
      ? 'Entrou na agenda filtrada, mas nao esta na tabela atual'
      : 'Apareceu no arquivo, mas foi descartada pelos filtros';
  const rowHtml=rows.length?rows.map(r=>{
    const ref=agendaRefDate(r);
    return `<div style="display:grid;grid-template-columns:90px 1fr;gap:8px;padding:8px 0;border-top:1px solid #334155;">
      <div style="color:#64748b;font-size:10px;">Linha ${escHtml(r.linha)}</div>
      <div style="font-size:12px;line-height:1.6;">
        <b style="color:${r.included?'#86efac':'#fca5a5'};">${r.included?'Incluida':'Fora'}</b>
        <span style="color:#94a3b8;"> · ${escHtml(r.motivo||'')}</span><br>
        <span style="color:#64748b;">Inicio:</span> ${escHtml(fmtDT(r.AGENDA,true)||'-')}
        <span style="color:#64748b;margin-left:10px;">Fim:</span> ${escHtml(fmtDT(r.FIM_AGENDA,true)||'-')}
        <span style="color:#64748b;margin-left:10px;">Dia usado:</span> ${escHtml(ref?dKey(ref):'-')}<br>
        <span style="color:#64748b;">Local:</span> ${escHtml(r.LOCAL||'-')}
        <span style="color:#64748b;margin-left:10px;">Doca:</span> ${escHtml(r.DOCA||'-')}<br>
        <span style="color:#64748b;">Transp.:</span> ${escHtml(r.TRANSPORTADORA||'-')}
        <span style="color:#64748b;margin-left:10px;">Peso:</span> ${escHtml(r.PESO||'-')}
      </div>
    </div>`;
  }).join(''):'<div style="padding:8px 0;color:#64748b;">Nao apareceu na ultima agenda lida, mas existe na tabela atual.</div>';
  const dashHtml=inDash.length?inDash.map(r=>`<div style="margin-top:8px;background:#052e16;border:1px solid #22c55e55;border-radius:8px;padding:8px 10px;font-size:12px;">
    Tabela atual: <b>${escHtml(r.dia_ref||'-')}</b> · data_ref ${escHtml(r.data_ref||'-')} · status ${escHtml(r.status||'-')} · fim ${escHtml(r.fim_carregamento||'-')}
  </div>`).join(''):'';
  return `<div style="background:#0f172a;border:1px solid #334155;border-left:3px solid ${statusColor};border-radius:8px;padding:12px 14px;margin-bottom:10px;">
    <div style="display:flex;align-items:center;gap:10px;justify-content:space-between;flex-wrap:wrap;">
      <div style="font-size:15px;font-weight:900;color:#93c5fd;">DT ${escHtml(dt)}</div>
      <div style="font-size:11px;color:${statusColor};font-weight:800;">${statusText}</div>
    </div>
    <div style="font-size:11px;color:#64748b;margin-top:3px;">Referencia pelo fim da agenda: ${escHtml(refLabels.join(', ')||'-')}</div>
    ${rowHtml}
    ${dashHtml}
  </div>`;
}

function renderAgendaSearch(){
  const input=document.getElementById('agenda-search-input');
  const body=document.getElementById('agenda-search-body');
  const meta=document.getElementById('agenda-search-meta');
  if(!body) return;
  const term=(input&&input.value||'').trim();
  if(meta){
    meta.textContent=(agendaDiagRows&&agendaDiagRows.length)
      ? agendaDiagRows.length+' linhas da ultima agenda lida em memoria'
      : 'Nenhuma agenda carregada nesta sessao';
  }
  if(!term){
    body.innerHTML='<div style="color:#64748b;padding:18px;text-align:center;">Digite uma ou várias DTs, transportadora ou doca para investigar.</div>';
    return;
  }

  const searches=agendaSearchTerms(term);
  const cards=[];
  const renderedDTs=new Set();
  searches.forEach(search=>{
    const diagRows=agendaDiagnosticRowsForTerm(search.value,{dtOnly:search.dtOnly});
    const dashRows=agendaDashRowsForTerm(search.value,{dtOnly:search.dtOnly});
    const foundDTs=[...new Set([...diagRows.map(r=>normalizeDT(r.DT)),...dashRows.map(r=>normalizeDT(r.dt))].filter(Boolean))];
    if(!foundDTs.length){
      cards.push(buildAgendaSearchMissingCard(search.value));
      return;
    }
    foundDTs.forEach(dt=>{
      const key=normalizeDT(dt);
      if(renderedDTs.has(key)) return;
      renderedDTs.add(key);
      cards.push(buildAgendaSearchCard(dt,diagRows,dashRows));
    });
  });

  body.innerHTML=cards.join('') || '<div style="color:#fca5a5;background:#450a0a;border:1px solid #ef444455;border-radius:8px;padding:12px;">Nenhum resultado encontrado.</div>';
}

function setStep(n){
  // Passos: 1,2,3,4 (removemos o passo 5, agora é 4)
  [1,2,3,4].forEach(i=>{
    document.getElementById('passo'+i).style.display=i===n?'block':'none';
    const d=document.getElementById('dot'+i);
    if(!d)return;
    d.className='sdot'+(i<n?' done':i===n?' active':'');
    d.textContent=i<n?'✓':i;
  });
  document.getElementById('aerr').style.display='none';
  // Pausa o auto-sync enquanto o usuário estiver no assistente de upload.
  if(n!==4) pauseAutoSync();
  // Quando entrar no passo 3, mostra dica de materiais já importados
  if(n===1) activeRefsOverride=[];
  if(n===3){
    const hint=document.getElementById('p3-mat-hint');
    const nDTs=Object.keys(exportMap).length;
    if(hint){
      if(nDTs>0){
        hint.textContent='✨ '+nDTs+' DTs já com materiais importados do arquivo de agendamento! Você pode pular este passo.';
        hint.style.display='inline';
      }else{
        hint.style.display='none';
      }
    }
  }
}

function switchTab(tab){
  ['file','paste','xlsx'].forEach(t=>{
    document.getElementById('tab-'+t+'-body').style.display=t===tab?'block':'none';
    const btn=document.getElementById('tab-'+t);
    btn.style.background=t===tab?'#3b82f6':'#1e293b';
    btn.style.color=t===tab?'#fff':'#94a3b8';
  });
}

let matFiltro='todos';

function switchTableTab(tab){
  if(tab==='materiais') tab='todas';
  currentTableTab=tab;
  const allTabs=['todas','finalizadas','reporte','planejamento','mudancas','semgrade','dtfake','docanull','nf'];
  allTabs.forEach(t=>{
    const el=document.getElementById('tab-'+t);
    if(el) el.classList.toggle('active',t===tab);
  });
  // Hide all content areas first
  ['mat-board','twrap','materiais-wrap','reporte-wrap','planejamento-wrap','mudancas-wrap','semgrade-wrap','dtfake-wrap','docanull-wrap','nf-wrap'].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.style.display='none';
  });
  if(tab==='reporte'){
    document.getElementById('reporte-wrap').style.display='block';
    renderReporte();
  } else if(tab==='planejamento'){
    document.getElementById('planejamento-wrap').style.display='block';
    renderReporte();
    renderPlanejamento();
  } else if(tab==='mudancas'){
    document.getElementById('mudancas-wrap').style.display='block';
    renderMudancasGrade();
  } else if(tab==='semgrade'){
    document.getElementById('semgrade-wrap').style.display='block';
    renderSemGrade();
  } else if(tab==='dtfake'){
    document.getElementById('dtfake-wrap').style.display='block';
  } else if(tab==='docanull'){
    document.getElementById('docanull-wrap').style.display='block';
    renderDocaNullAudit();
  } else if(tab==='nf'){
    document.getElementById('nf-wrap').style.display='block';
    nfRenderLista();
  } else {
    document.getElementById('twrap').style.display='block';
    renderRows();
  }
}

function filtrarMateriais(f){
  matFiltro=f;
  ['todos','venda','transf'].forEach(id=>{
    const btn=document.getElementById('mfilt-'+id);
    if(!btn) return;
    btn.style.background=f===id?'#1e3a5f':'transparent';
    btn.style.fontWeight=f===id?'800':'400';
  });
  renderMateriaisAba();
}

async function renderMateriaisAba(){
  const board=document.getElementById('materiais-board');
  const resumo=document.getElementById('mfilt-resumo');
  if(!board) return;
  board.innerHTML='<div style="color:#64748b;padding:20px;">Carregando materiais…</div>';
  try{
    // Ordem das DTs conforme grade
    const ordem=[...tableData]
      .sort(compareAgendaRows)
      .map(r=>String(r.dt));
    const uniq=[]; ordem.forEach(dt=>{if(!uniq.includes(dt)) uniq.push(dt);});
    // Adiciona DTs de remessaMap não presentes na tabela
    Object.keys(remessaMap).forEach(dt=>{if(!uniq.includes(dt)) uniq.push(dt);});

    // Aplica filtro de tipo
    const dtsFiltradas=uniq.filter(dt=>{
      const row=tableData.find(r=>r.dt===dt)||{};
      const op=(row.tipo_operacao||tipoOpMap[dt]||'').toUpperCase();
      if(matFiltro==='venda')  return op.includes('ARUJA')||(op.includes('VENDA')&&!op.includes('MOGI')&&!op.includes('TRANSFER'));
      if(matFiltro==='mogi')   return op.includes('MOGI');
      if(matFiltro==='transf') return op.includes('TRANSFER');
      return true;
    });

    let totalDTs=0, totalRemessas=0, totalItens=0;
    let html='';

    dtsFiltradas.forEach(dt=>{
      const itens=remessaMap[dt];
      if(!itens||!itens.length) return;
      totalDTs++;
      const row=tableData.find(r=>r.dt===dt)||{};
      const op=(row.tipo_operacao||tipoOpMap[dt]||'').toUpperCase();
      const isTransf=op.includes('TRANSFER');
      const opBadge=isTransf
        ?'<span class="tag-transf" style="padding:2px 8px;font-size:10px;">🔄 TRANSFERÊNCIA</span>'
        :op.includes('MOGI')?'<span class="tag-mogi" style="padding:2px 8px;font-size:10px;">🏙️ VENDA MOGI</span>':'<span class="tag-venda" style="padding:2px 8px;font-size:10px;">🏭 VENDA ARUJA</span>';
      const diaTag=row.dia_ref==='HOJE'
        ?'<span class="tag-hoje" style="padding:1px 6px;font-size:9px;">HOJE</span>'
        :(row.dia_ref==='AMANHÃ'?'<span class="tag-amanha" style="padding:1px 6px;font-size:9px;">AMANHÃ</span>':'');
      const transp=row.transportadora||'—';
      const grade=row.grade_carregamento||'—';
      const status=row.status||'';
      const sc=SC[status]||'#334155';
      const descDoc=row.descricao_documento||descDocMap[dt]||'';

      // Agrupa por remessa para exibição
      const byRemessa={};
      itens.forEach(i=>{
        if(!byRemessa[i.remessa]) byRemessa[i.remessa]=[];
        byRemessa[i.remessa].push({material:i.material, qtde:i.qtde});
      });
      const remessas=Object.keys(byRemessa);
      totalRemessas+=remessas.length;
      totalItens+=itens.length;

      // Soma total em paletes (a quantidade original vem em fardos)
      const paletesTotalDT = itens.reduce((sum, i) => sum + formatPaletes(i.material,i.qtde).paletes, 0);
      const sobrasTotalDT = itens.reduce((sum, i) => sum + formatPaletes(i.material,i.qtde).sobra, 0);

      html+=`<div class="mat-group" style="margin-bottom:14px;">`;
      html+=`<div class="mat-group-h" style="display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center;padding:10px 14px;">`;
      html+=`<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">`;
      html+=`<span style="color:#93c5fd;font-size:14px;font-weight:900;">🚛 ${dt}</span>`;
      html+=`${diaTag} ${opBadge}`;
      if(descDoc) html+=`<span style="color:#94a3b8;font-size:10px;font-style:italic;">${descDoc}</span>`;
      html+=`<span style="color:#64748b;font-size:10px;">| ${transp}</span>`;
      html+=`<span style="color:#a5b4fc;font-size:10px;">| ⏱ ${grade}</span>`;
      html+=`<span style="background:${sc}22;border:1px solid ${sc}55;color:${sc};border-radius:4px;padding:1px 7px;font-size:9px;font-weight:700;">${status}</span>`;
      html+=`</div>`;
      html+=`<div style="text-align:right;">`;
      html+=`<div style="color:#64748b;font-size:10px;">${remessas.length} remessa(s) · ${itens.length} item(ns)</div>`;
      html+=`<div style="font-size:12px;font-weight:700;color:#a7f3d0;">Total: ${paletesTotalDT} palete(s) necessário(s) + ${formatNumeroBR(sobrasTotalDT)} fardo(s) sobra</div>`;
      html+=`</div>`;
      html+=`</div>`;

      // Tabela agrupada por remessa
      html+=`<div style="padding:0 14px 10px;">`;
      remessas.forEach(remessa=>{
        const linhas=byRemessa[remessa];
        html+=`<div style="margin-bottom:6px;">`;
        html+=`<div style="display:grid;grid-template-columns:1fr 120px 160px;gap:4px;padding:4px 0;border-bottom:1px solid #334155;margin-bottom:2px;">`;
        html+=`<span style="font-size:9px;letter-spacing:1.5px;color:#3b82f6;font-weight:700;">REMESSA ${remessa}</span>`;
        html+=`<span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;">MATERIAL</span>`;
        html+=`<span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;text-align:right;">PALETES</span>`;
        html+=`</div>`;
        linhas.forEach(({material,qtde},idx)=>{
          const pal=formatPaletes(material,qtde);
          const detalhe=pal.porPalete?`<div style="font-size:9px;color:#64748b;font-weight:500;">${parseQuantidadeFardos(qtde).toLocaleString('pt-BR',{maximumFractionDigits:2})} fardos · ${pal.porPalete}/palete</div>`:'';
          html+=`<div style="display:grid;grid-template-columns:1fr 120px 160px;gap:4px;padding:2px 0;border-bottom:1px solid #1e293b;font-size:12px;">`;
          html+=`<span style="color:#475569;font-size:10px;">${idx===0?'':'—'}</span>`;
          html+=`<span style="color:#cbd5e1;">${material}</span>`;
          html+=`<span style="text-align:right;color:#a7f3d0;font-weight:700;">${pal.label}${detalhe}</span>`;
          html+=`</div>`;
        });
        html+=`</div>`;
      });
      html+=`</div></div>`;
    });

    if(resumo) resumo.textContent=`${totalDTs} transportes · ${totalRemessas} remessas · ${totalItens} itens`;
    board.innerHTML=html||'<div style="color:#334155;padding:20px;">Nenhum material encontrado'+
      (matFiltro!=='todos'?' para o filtro selecionado.':'. Carregue o CSV do Relatório de Expedição no Passo 3.')+'</div>';
  }catch(e){
    board.innerHTML='<div style="color:#ef4444;padding:20px;">Erro ao carregar: '+e.message+'</div>';
  }
}

function onSearchDT(v){
  dtSearchTerm=(v||'').replace(/\D/g,'');
  renderRows();
}

function setTableSortMode(v){
  tableSortMode=v==='fim'?'fim':'inicio';
  try{localStorage.setItem('table_sort_mode',tableSortMode);}catch(e){}
  renderRows();
}

/* ═══════════════════════════════════════════════════════
   DATAS
═══════════════════════════════════════════════════════ */
function today()   {const d=new Date();d.setHours(0,0,0,0);return d;}
function yesterday(){const d=today();d.setDate(d.getDate()-1);return d;}
function tomorrow(){const d=today();d.setDate(d.getDate()+1);return d;}
function sameDay(a,b){return a&&b&&a.getDate()===b.getDate()&&a.getMonth()===b.getMonth()&&a.getFullYear()===b.getFullYear();}
function dKey(d){return d.toLocaleDateString('pt-BR');}
function fmtDT(d,t){
  if(!d)return '';
  return t?d.toLocaleString('pt-BR',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'}):d.toLocaleDateString('pt-BR');
}
function parseBR(s){
  if(!s||!s.trim())return null;
  const m=s.match(/(\d{2})\/(\d{2})\/(\d{4})(?:[,T\s]+(\d{2}):(\d{2}))?/);
  return m?new Date(+m[3],+m[2]-1,+m[1],+(m[4]||0),+(m[5]||0)):null;
}

function agendaRefDate(dt){
  // REGRA 2/3: data de referência é sempre o FIM de carregamento, nunca o início.
  return (dt&&dt.FIM_AGENDA) || null;
}

function rowRefDate(row){
  return parseBR(String((row&&row.fim_carregamento)||'')) ||
    parseBR(String((row&&row.fim_agenda)||'')) ||
    parseBR(String((row&&row.data_ref)||''));
}

function rowLoadingEndDate(row){
  return parseBR(String((row&&row.fim_carregamento)||'')) ||
    parseBR(String((row&&row.fim_agenda)||''));
}

function rowReportRefDate(row){
  return rowLoadingEndDate(row) ||
    parseBR(String((row&&row.data_ref)||''));
}

function rowStatusRefDate(row){
  // Janela ativa da importação de status continua baseada no FIM de carregamento.
  return rowLoadingEndDate(row);
}

function rowImportAgendaDate(row){
  // Para identificar DTs duplicadas na importação, a data/hora da planilha deve
  // ser comparada com a agenda (início), não com o dia atual nem com o fim.
  return parseBR(String((row&&row.grade_carregamento)||'')) ||
    parseBR(String((row&&row.agenda)||''));
}

function rowReportRefKey(row){
  const ref=rowReportRefDate(row);
  return ref?dKey(ref):String((row&&row.data_ref)||'').trim();
}

function parseAgendaDateTime(row){
  const grade=parseBR(row.grade_carregamento||'');
  if(grade) return grade;
  const agenda=parseBR(row.agenda||'');
  if(agenda) return agenda;
  return new Date(9999,0,1);
}
function diaRefAtualPorDataRef(dataRef){
  const d=parseBR(String(dataRef||''));
  if(!d) return '';
  if(sameDay(d,today())) return 'HOJE';
  if(sameDay(d,tomorrow())) return 'AMANHÃ';
  return '';
}
function diaRefAtualPorRow(row){
  const fim=rowRefDate(row);
  if(fim){
    if(sameDay(fim,today())) return 'HOJE';
    if(sameDay(fim,tomorrow())) return 'AMANHÃ';
  }
  return diaRefAtualPorDataRef(row&&row.data_ref);
}
function normalizeDiaRefRow(row){
  const diaAtual=diaRefAtualPorRow(row);
  if(!diaAtual) return row;
  return {...row,_stored_dia_ref:row.dia_ref,dia_ref:diaAtual};
}
function normalizeDiaRefRows(rows){
  return dedupeCargaRowsByDTRef((rows||[]).map(normalizeDiaRefRow));
}

const REPORTE_CARGA_UPSERT_COLUMNS=[
  'dt','transportadora','grade_carregamento','fim_carregamento','hora_chegada','n_portaria',
  'status','descricao_documento','centro','toneladas','peso_liquido','agenda','local_cd',
  'dia_ref','data_ref','tipo_operacao','reagendada','doca_null','paletizacao','updated_at'
];

function cleanCargaRowForInsert(row){
  const refDate=rowReportRefDate(row);
  const dataRef=refDate?dKey(refDate):String(row.data_ref||'');
  const clean={};
  REPORTE_CARGA_UPSERT_COLUMNS.forEach(col=>{ clean[col]=row&&row[col]!==undefined&&row[col]!==null?row[col]:''; });
  clean.dt=String(normalizeDT(clean.dt));
  clean.data_ref=String(dataRef);
  clean.dia_ref=sameDay(refDate,today())?'HOJE':(sameDay(refDate,tomorrow())?'AMANHÃ':'');
  clean.reagendada=!!row.reagendada;
  clean.doca_null=!!row.doca_null;
  clean.updated_at=new Date().toISOString();
  return clean;
}

function isRowInActiveReportWindow(row){
  const ref=rowReportRefDate(row);
  return !!ref && (sameDay(ref,today())||sameDay(ref,tomorrow()));
}

async function rewriteActiveCargaSnapshot(){
  if(!tableData||!tableData.length) return 0;
  const rows=dedupeCargaRowsByDTRef(tableData.filter(isRowInActiveReportWindow))
    .map(cleanCargaRowForInsert)
    .filter(r=>r.dt&&r.data_ref);

  if(rows.length) await sbUpsert('reporte_carga',rows,'dt,data_ref');
  tableData=normalizeDiaRefRows(rows).sort(compareAgendaRows);
  return rows.length;
}
async function persistNormalizedDiaRefs(rows){
  for(const row of rows||[]){
    if(!row||!row.dt||!row.data_ref||!row._stored_dia_ref) continue;
    if(row._stored_dia_ref===row.dia_ref) continue;
    try{
      await sbPatch('reporte_carga',{dia_ref:row.dia_ref,updated_at:new Date().toISOString()},{dt:String(row.dt),data_ref:String(row.data_ref)});
      row._stored_dia_ref=row.dia_ref;
    }catch(e){
      console.warn('Não foi possível corrigir dia_ref da DT '+row.dt,e);
    }
  }
}
function compareAgendaRows(a,b){
  const iniA=parseBR(a.grade_carregamento||'')||new Date(9999,0,1);
  const iniB=parseBR(b.grade_carregamento||'')||new Date(9999,0,1);
  const diffIni=iniA-iniB;
  if(diffIni) return diffIni;
  return String(a.dt||'').localeCompare(String(b.dt||''),'pt-BR',{numeric:true});
}

function compareFimAgendaRows(a,b){
  const fimA=parseBR(a.fim_carregamento||'')||parseBR(a.grade_carregamento||'')||new Date(9999,0,1);
  const fimB=parseBR(b.fim_carregamento||'')||parseBR(b.grade_carregamento||'')||new Date(9999,0,1);
  const diffFim=fimA-fimB;
  if(diffFim) return diffFim;
  return compareAgendaRows(a,b);
}

function compareTableRows(a,b){
  return tableSortMode==='fim'?compareFimAgendaRows(a,b):compareAgendaRows(a,b);
}

const FARDOS_POR_PALETE={
  '20104414':36,'20104415':36,'20104416':36,'20104419':33,'20104418':28,
  '20106704':36,'20106705':36,'20104405':48,'20104407':32,'20104410':36,
  '20104412':28,'20104413':36,'20104409':36,'20104408':36,'20081464':36,
  '20081466':36,'20081481':28,'20081961':36,'20089387':24,'20095269':36,
  '20081984':28,'20081465':36,'20081467':36,'20081469':28,'20111061':36,
  '20104309':45,'20104421':63,'20104422':56,'20109594':56,'20104430':15,
  '20104429':27,'20104426':27,'20104313':36,'20104425':36,'20104312':36,
  '20110078':36,'20091834':36,'20091835':42,'30227945':15,'30179006':18,
  '30228198':15,'20109727':24,'20091836':36,'20104310':225
};

function parseQuantidadeFardos(qtde){
  const n=parseFloat(String(qtde||'0').replace(/\./g,'').replace(',','.'));
  return Number.isFinite(n)?n:0;
}

function formatNumeroBR(n){
  return Number(n||0).toLocaleString('pt-BR',{maximumFractionDigits:2});
}

function formatPaletes(material,qtde){
  const sku=String(material||'').trim().replace(/\D/g,'');
  const fardos=parseQuantidadeFardos(qtde);
  const porPalete=FARDOS_POR_PALETE[sku];
  if(!porPalete){
    return {
      paletes:0,
      sobra:fardos,
      porPalete:null,
      label:fardos?`Sem fator: ${formatNumeroBR(fardos)} fardo(s)`:'—'
    };
  }
  const paletes=Math.floor(fardos/porPalete);
  const sobra=Math.round((fardos%porPalete)*100)/100;
  return {
    paletes,
    sobra,
    porPalete,
    label:`${paletes} palete(s) necessário(s) + ${formatNumeroBR(sobra)} fardo(s) sobra`
  };
}
function normalizeDT(raw){
  const v=String(raw||'').trim();
  if(!v) return '';
  if(/^\d+$/.test(v)) return v.replace(/^0+(?=\d)/,'');
  return v;
}

function mergeAgendaDuplicateRows(base,incoming){
  const merged={...base};
  Object.keys(incoming||{}).forEach(k=>{
    const current=merged[k];
    const next=incoming[k];
    if((current===undefined||current===null||current==='') && next!==undefined && next!==null && next!==''){
      merged[k]=next;
    }
  });
  return merged;
}

function dedupeAgendaRowsByDTRef(rows){
  const byKey=new Map();
  (rows||[]).forEach(row=>{
    const dt=normalizeDT(row&&row.DT);
    const refDate=agendaRefDate(row);
    const key=dt+'__'+(refDate?dKey(refDate):'');
    if(!dt||!refDate){
      byKey.set(key+'__'+byKey.size,row);
      return;
    }
    byKey.set(key,byKey.has(key)?mergeAgendaDuplicateRows(byKey.get(key),row):row);
  });
  return [...byKey.values()];
}

function dedupeCargaRowsByDTRef(rows){
  const byKey=new Map();
  (rows||[]).forEach(row=>{
    const dt=normalizeDT(row&&row.dt);
    const refKey=rowReportRefKey(row);
    const key=dt+'__'+refKey;
    if(!dt||!refKey){
      byKey.set(key+'__'+byKey.size,row);
      return;
    }
    byKey.set(key,byKey.has(key)?mergeAgendaDuplicateRows(byKey.get(key),row):row);
  });
  return [...byKey.values()];
}

function normalizeHora(raw){
  const s=String(raw||'').trim();
  if(!s) return '';
  const m=s.match(/(\d{1,2})[:hH](\d{2})/);
  if(m){
    const hh=Math.max(0,Math.min(23,parseInt(m[1],10)));
    const mm=Math.max(0,Math.min(59,parseInt(m[2],10)));
    return String(hh).padStart(2,'0')+':'+String(mm).padStart(2,'0');
  }
  if(/^\d+(?:[,.]\d+)?$/.test(s)){
    const n=Number(s.replace(',','.'));
    if(n>0 && n<1){
      const total=Math.round(n*24*60);
      return String(Math.floor(total/60)%24).padStart(2,'0')+':'+String(total%60).padStart(2,'0');
    }
    if(s.length===3||s.length===4){
      const pad=s.padStart(4,'0');
      return pad.slice(0,2)+':'+pad.slice(2);
    }
  }
  return '';
}

function normalizeSap(raw){
  return String(raw||'').trim().replace(/\.0+$/,'').replace(/\D/g,'');
}

function parseImportDate(raw){
  const s=String(raw||'').trim();
  if(!s) return null;
  const br=parseBR(s);
  if(br) return br;
  const iso=s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if(iso) return new Date(+iso[1],+iso[2]-1,+iso[3]);
  if(/^\d+(?:[,.]\d+)?$/.test(s)){
    const n=Number(s.replace(',','.'));
    if(n>20000&&n<80000){
      const excelEpoch=new Date(1899,11,30);
      excelEpoch.setDate(excelEpoch.getDate()+Math.floor(n));
      return excelEpoch;
    }
  }
  return null;
}

function combineImportDateHora(dataRaw,horaRaw){
  const d=parseImportDate(dataRaw);
  if(!d) return null;
  const hora=normalizeHora(horaRaw);
  if(hora){
    const [hh,mm]=hora.split(':').map(n=>parseInt(n,10));
    d.setHours(hh||0,mm||0,0,0);
  }else{
    d.setHours(0,0,0,0);
  }
  return d;
}

function sameMinute(a,b){
  return a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate()&&a.getHours()===b.getHours()&&a.getMinutes()===b.getMinutes();
}

function findDashRowForImportedStatus(csvRow,dashRows){
  const dt=String(csvRow&&csvRow.dt||'');
  const candidates=(dashRows||[]).filter(r=>String(r.dt)===dt);
  if(!candidates.length) return null;
  const importAgenda=csvRow&&csvRow.fimAgendamento;
  if(importAgenda){
    const exactAgenda=candidates.find(r=>sameMinute(rowImportAgendaDate(r),importAgenda));
    if(exactAgenda) return exactAgenda;
    const sameAgendaDay=candidates.find(r=>sameDay(rowImportAgendaDate(r),importAgenda));
    if(sameAgendaDay) return sameAgendaDay;

    // Fallback para arquivos antigos em que DATA/HORA tenha sido preenchida com o fim.
    const exactEnd=candidates.find(r=>sameMinute(rowStatusRefDate(r),importAgenda));
    if(exactEnd) return exactEnd;
    const sameEndDay=candidates.find(r=>sameDay(rowStatusRefDate(r),importAgenda));
    if(sameEndDay) return sameEndDay;
  }
  return candidates[0];
}

/* ═══════════════════════════════════════════════════════
   DECODE BUFFER
═══════════════════════════════════════════════════════ */
function decodeBuf(buf){
  const b=new Uint8Array(buf);
  if(b[0]===0xFF&&b[1]===0xFE)return new TextDecoder('utf-16le').decode(buf);
  if(b[0]===0xFE&&b[1]===0xFF)return new TextDecoder('utf-16be').decode(buf);
  let nulls=0;for(let i=1;i<Math.min(200,b.length);i+=2)if(b[i]===0)nulls++;
  if(nulls>15)return new TextDecoder('utf-16le').decode(buf);
  try{return new TextDecoder('utf-8').decode(buf);}catch(e){}
  return new TextDecoder('latin1').decode(buf);
}


function normHeader(v){
  return String(v||'')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[^A-Z0-9]+/gi,' ')
    .trim().toUpperCase();
}

function pickHeader(keys,...names){
  const normalizedKeys=(keys||[]).map(k=>({raw:k,norm:normHeader(k)})).filter(k=>k.norm);
  const wanted=(names||[]).map(normHeader).filter(Boolean);
  for(const w of wanted){
    const exact=normalizedKeys.find(k=>k.norm===w);
    if(exact) return exact.raw;
  }
  for(const w of wanted){
    const starts=normalizedKeys.find(k=>k.norm.startsWith(w+' ')||k.norm.endsWith(' '+w));
    if(starts) return starts.raw;
  }
  for(const w of wanted){
    const contains=normalizedKeys.find(k=>k.norm.includes(w)||w.includes(k.norm));
    if(contains) return contains.raw;
  }
  return '';
}


function pickAgendaStartHeader(keys,...names){
  const withoutEnd=(keys||[]).filter(k=>{
    const n=normHeader(k);
    return !/^(FIM|TERMINO|TÉRMINO)\b/.test(n) && !/\b(FIM|TERMINO|TÉRMINO)$/.test(n);
  });
  return pickHeader(withoutEnd,...names);
}

function getAgendaHeaderNames(){
  return {
    dt:['DT','Nº TRANSPORTE','N TRANSPORTE','NUMERO TRANSPORTE','NÚMERO TRANSPORTE','TRANSPORTE','NR TRANSPORTE','NR. TRANSPORTE'],
    agenda:['AGENDA TRANSPORTADOR','INICIO AGENDA TRANSPORTADOR','INÍCIO AGENDA TRANSPORTADOR','INICIO AGENDA','INÍCIO AGENDA','DATA/HORA AGENDA TRANSPORTADOR','DATA AGENDA TRANSPORTADOR','AGENDA'],
    local:['LOCAL','LOCAL CARREGAMENTO','LOCAL DE CARREGAMENTO','CENTRO','CENTRO CD','CD','PLANTA'],
    transportadora:['NOME TRANSPORTADORA','TRANSPORTADORA','NOME TRANSP','TRANSP.','TRANSP'],
    fim:['FIM AGENDA TRANSPORTADOR','FIM DA AGENDA TRANSPORTADOR','DATA/HORA FIM AGENDA TRANSPORTADOR','FIM AGENDA','AGENDA FIM','FIM'],
    peso:['PESO','PESO LIQUIDO','PESO LÍQUIDO'],
    tipo:['TIPO VEICULO','TIPO VEÍCULO','TIPO DE VEICULO','TIPO DE VEÍCULO','TIPO'],
    doca:['DOCA','DOCA CARREGAMENTO','DOCA DE CARREGAMENTO'],
  };
}

const AGENDA_FIM_TRANSPORTADOR_COL_INDEX=22; // coluna 23 da agenda completa

function getAgendaCompleteFimValue(cols){
  return Array.isArray(cols)&&cols.length>AGENDA_FIM_TRANSPORTADOR_COL_INDEX ? cols[AGENDA_FIM_TRANSPORTADOR_COL_INDEX] : '';
}

function getAgendaFimValue(cols,mapFim,getByHeader){
  const col23=getAgendaCompleteFimValue(cols);
  // Na agenda completa, a coluna 23 é a referência oficial: FIM AGENDA TRANSPORTADOR.
  if(parseAnyDate(col23)) return col23;
  return mapFim?getByHeader(cols,mapFim):col23;
}

function getAgendaHeaderMap(keys){
  const h=getAgendaHeaderNames();
  return {
    DT:pickHeader(keys,...h.dt),
    LOCAL:pickHeader(keys,...h.local),
    TRANSPORTADORA:pickHeader(keys,...h.transportadora),
    AGENDA:pickAgendaStartHeader(keys,...h.agenda),
    FIM_AGENDA:pickHeader(keys,...h.fim),
    PESO:pickHeader(keys,...h.peso),
    TIPO:pickHeader(keys,...h.tipo),
    DOCA:pickHeader(keys,...h.doca),
  };
}

function agendaHeaderScore(keys){
  const map=getAgendaHeaderMap(keys);
  let score=0;
  if(map.DT) score+=4;
  if(map.AGENDA) score+=4;
  if(map.LOCAL) score+=2;
  if(map.DOCA) score+=2;
  if(map.FIM_AGENDA) score+=1;
  if(map.TRANSPORTADORA) score+=1;
  return score;
}

function parseAnyDate(v){
  if(v instanceof Date && !Number.isNaN(v.getTime())) return v;
  if(typeof v==='number' && Number.isFinite(v) && typeof XLSX!=='undefined'){
    const d=XLSX.SSF.parse_date_code(v);
    if(d) return new Date(d.y,d.m-1,d.d,d.H||0,d.M||0,Math.floor(d.S||0));
  }
  const raw=String(v||'').trim();
  if(!raw) return null;
  if(/^\d{4,6}(?:[,.]\d+)?$/.test(raw) && typeof XLSX!=='undefined'){
    const d=XLSX.SSF.parse_date_code(Number(raw.replace(',','.')));
    if(d) return new Date(d.y,d.m-1,d.d,d.H||0,d.M||0,Math.floor(d.S||0));
  }
  const br=parseBR(raw);
  if(br) return br;
  const iso=raw.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[T\s]+(\d{1,2}):(\d{2}))?/);
  if(iso) return new Date(+iso[1],+iso[2]-1,+iso[3],+(iso[4]||0),+(iso[5]||0));
  const flex=raw.match(/(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})(?:[,T\s]+(\d{1,2}):(\d{2}))?/);
  if(flex){
    const year=+flex[3]<100?2000+(+flex[3]):+flex[3];
    return new Date(year,+flex[2]-1,+flex[1],+(flex[4]||0),+(flex[5]||0));
  }
  return null;
}

function normalizeDocaName(v){
  return String(v||'')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .replace(/[\s\-]+/g,'_')
    .replace(/_+/g,'_');
}

function isEmptyDoca(v){
  const s=String(v??'').trim();
  return !s || /^(null|undefined|n\/?a|nao_informado)$/i.test(s);
}

function isLocalFinal1110Or1111(v){
  const s=String(v||'').trim().replace(/\.0+$/,'');
  const nums=s.match(/\d+/g)||[];
  const last=nums.length?nums[nums.length-1].replace(/^0+(?=\d)/,''):'';
  return /(?:1110|1111)$/.test(last);
}

function agendaInvalidSummary(totalRows){
  const counts=new Map();
  (agendaDiagRows||[]).forEach(r=>counts.set(r.motivo||'Sem motivo', (counts.get(r.motivo||'Sem motivo')||0)+1));
  const parts=[...counts.entries()]
    .sort((a,b)=>b[1]-a[1])
    .slice(0,4)
    .map(([motivo,total])=>`${total} ${motivo}`);
  const sample=(agendaDiagRows||[]).slice(0,3)
    .map(r=>`linha ${r.linha||'?'}: LOCAL="${r.LOCAL||'-'}", DOCA="${r.DOCA||'-'}", FIM="${fmtDT(r.FIM_AGENDA,true)||'-'}"`)
    .join(' | ');
  return `Nenhuma linha válida encontrada. Lidas ${totalRows||0} linhas. Motivos: ${parts.join('; ')||'sem diagnostico'}.${sample?' Amostras: '+sample:''}`;
}

function isDocaFabMogSemIfnt(docaNorm){
  return /(?:^|_)doca_\d+_fab_mog(?:_|$)/.test(docaNorm) && !/(?:^|_)ifnt(?:_|$)/.test(docaNorm);
}

function getAgendaRecordsFromWorkbook(buf){
  try{
    if(typeof XLSX==='undefined') return [];
    const wb=XLSX.read(new Uint8Array(buf),{type:'array',cellDates:true});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const table=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true,blankrows:false});
    if(!table.length) return [];
    let headerIdx=-1, header=[], bestScore=0;
    table.slice(0,50).forEach((row,i)=>{
      const keys=(row||[]).map(v=>String(v||'').trim());
      const score=agendaHeaderScore(keys);
      if(score>bestScore){bestScore=score;headerIdx=i;header=keys;}
    });
    const map=getAgendaHeaderMap(header);
    if(headerIdx<0||!map.DT) return [];
    const idx=k=>k?header.indexOf(k):-1;
    const get=(cols,k)=>{const i=idx(k);return i>=0?cols[i]:'';};
    return table.slice(headerIdx+1).map((cols,i)=>({
      linha:headerIdx+i+2,
      DT:get(cols,map.DT),
      LOCAL:get(cols,map.LOCAL),
      TRANSPORTADORA:get(cols,map.TRANSPORTADORA),
      AGENDA:get(cols,map.AGENDA),
      FIM_AGENDA:getAgendaFimValue(cols,map.FIM_AGENDA,get),
      PESO:get(cols,map.PESO),
      TIPO:get(cols,map.TIPO),
      DOCA:get(cols,map.DOCA),
    })).filter(r=>String(r.DT||'').trim());
  }catch(e){return [];}
}

function splitAgendaLine(line,sep){
  const cols=[];
  let cur='', inQ=false;
  for(let i=0;i<line.length;i++){
    const c=line[i];
    if(c==='"'){
      if(inQ&&line[i+1]==='"'){cur+='"';i++;}
      else inQ=!inQ;
    }else if(c===sep&&!inQ){
      cols.push(cur);cur='';
    }else cur+=c;
  }
  cols.push(cur);
  return cols.map(s=>String(s||'').trim().replace(/^"|"$/g,''));
}

function getAgendaRecordsFromText(buf){
  const lines=decodeBuf(buf).split(/\r?\n/).filter(l=>l.trim());
  if(!lines.length) return [];
  const sep=lines.find(l=>l.includes('\t'))?'\t':(lines.find(l=>l.includes(';'))?';':',');
  let headerIdx=-1, h=[], bestScore=0;
  lines.slice(0,50).forEach((line,i)=>{
    const keys=splitAgendaLine(line,sep);
    const score=agendaHeaderScore(keys);
    if(score>bestScore){bestScore=score;headerIdx=i;h=keys;}
  });
  const map=getAgendaHeaderMap(h);
  if(headerIdx<0||!map.DT) throw new Error('Coluna DT não encontrada. A data oficial deve estar na coluna 23 (FIM AGENDA TRANSPORTADOR).\nColunas: '+(h.length?h.join(' | '):'nenhuma linha de cabeçalho reconhecida'));
  const idx=k=>k?h.indexOf(k):-1;
  const get=(cols,k)=>{const i=idx(k);return i>=0?cols[i]||'':'';};
  return lines.slice(headerIdx+1).map((line,i)=>{
    const c=splitAgendaLine(line,sep);
    return {linha:headerIdx+i+2,DT:get(c,map.DT),LOCAL:get(c,map.LOCAL),TRANSPORTADORA:get(c,map.TRANSPORTADORA),AGENDA:get(c,map.AGENDA),FIM_AGENDA:getAgendaFimValue(c,map.FIM_AGENDA,get),PESO:get(c,map.PESO),TIPO:get(c,map.TIPO),DOCA:get(c,map.DOCA)};
  }).filter(r=>String(r.DT||'').trim());
}

function getAgendaRecords(buf){
  const workbookRows=getAgendaRecordsFromWorkbook(buf);
  return workbookRows.length?workbookRows:getAgendaRecordsFromText(buf);
}

function agendaRowsForDia(dia){
  return agendRows.filter(r=>agendaRefDate(r)&&sameDay(agendaRefDate(r),dia));
}

/* ═══════════════════════════════════════════════════════
   PASSO 1 — AGENDAMENTO
═══════════════════════════════════════════════════════ */
function processAgend(file){
  if(!file)return;
  fileWorkflowInProgress=true;
  pauseAutoSync(180000);
  showInf('Lendo agendamento…');
  const isInlineUpload = document.getElementById('passo4').style.display !== 'none';
  if(document.getElementById('dz1l')) document.getElementById('dz1l').innerHTML='<span class="dz-ok">✓ '+file.name+'</span>';
  file.arrayBuffer().then(async buf=>{
    try{
      // Tenta importar materiais automaticamente se o arquivo tiver as colunas certas (XLSX)
      exportMap={};remessaMap={};pesoLiquidoMap={};
      const matCnt=tryImportMaterialsFromAgend(buf);
      if(matCnt>0) showOk('✨ '+matCnt+' linhas de materiais importadas automaticamente do arquivo de agendamento!');
      const records=getAgendaRecords(buf);
      if(!records.length) throw new Error('Nenhuma linha encontrada no arquivo de agendamento.');
      agendRows=[];
      agendaDiagRows=[];
      for(const rec of records){
        const dtNorm=normalizeDT(rec.DT);
        if(!dtNorm) continue;
        const loc=String(rec.LOCAL||'').trim();
        const doca=String(rec.DOCA||'').trim();
        const docaNorm=normalizeDocaName(doca);
        const agendaRaw=parseAnyDate(rec.AGENDA);
        const fimAgendaRaw=parseAnyDate(rec.FIM_AGENDA);
        const isMogiOuAruja = isLocalFinal1110Or1111(loc);
        const diag={
          DT:dtNorm,LOCAL:loc,DOCA:doca,
          TRANSPORTADORA:String(rec.TRANSPORTADORA||'').trim(),
          AGENDA:agendaRaw,
          FIM_AGENDA:fimAgendaRaw,
          TIPO:String(rec.TIPO||'').trim(),
          PESO:String(rec.PESO||'').trim(),
          included:false,
          motivo:'',
          linha:rec.linha||'',
        };
        if(!isMogiOuAruja){
          diag.motivo='Fora do escopo: LOCAL nao termina com 1110/1111.';
          agendaDiagRows.push(diag);
          continue;
        }
        // DTs com DOCA vazia/null entram na grade para auditoria, mas ficam fora do reporte
        // enquanto não forem liberadas manualmente pelo operador.
        const docaNula = isEmptyDoca(doca);
        if(docaNula){
          diag.motivo='DOCA vazia/null — importada para auditoria e fora do reporte até liberação manual.';
        }
        if(isDocaFabMogSemIfnt(docaNorm)){
          diag.motivo='Descartada pela regra de DOCA FAB_MOG sem sufixo IFNT.';
          agendaDiagRows.push(diag);
          continue;
        }
        if(!fimAgendaRaw){
          diag.motivo='Sem FIM AGENDA TRANSPORTADOR valido na coluna 23.';
          agendaDiagRows.push(diag);
          continue;
        }
        diag.included=true;
        diag.motivo='Entrou na agenda filtrada.';
        agendaDiagRows.push(diag);
        agendRows.push({
          DT:dtNorm,LOCAL:loc,DOCA:doca,
          TRANSPORTADORA:diag.TRANSPORTADORA,
          AGENDA:agendaRaw,
          FIM_AGENDA:fimAgendaRaw,
          TIPO:diag.TIPO,
          PESO:diag.PESO,
          DOCA_NULL:docaNula,
        });
      }
      agendRows=dedupeAgendaRowsByDTRef(agendRows);
      if(!agendRows.length)throw new Error(agendaInvalidSummary(records.length));
      hideInf();
      if(isInlineUpload){
        // Upload feito de dentro da tabela: pula passo 2/3 e vai direto para o banco
        const T=today(),AM=tomorrow();
        const dtsH=agendaRowsForDia(T);
        const dtsA=agendaRowsForDia(AM);
        dtsMescladas=dedupeAgendaRowsByDTRef([...dtsH,...dtsA]);
        showOk('Agenda atualizada ('+agendRows.length+' linhas). Sincronizando banco…');
        await buildTable();
        releaseAutoSyncAfterSave();
      } else {
        renderStep2();
      }
    }catch(e){hideInf();showErr(e.message);releaseAutoSyncAfterSave(60000);}
  }).catch(e=>{hideInf();showErr('Erro ao abrir o arquivo: '+e.message);releaseAutoSyncAfterSave(60000);});
}

function renderStep2(){
  const T=today(),AM=tomorrow();
  const dtsH=agendaRowsForDia(T);
  const dtsA=agendaRowsForDia(AM);
  dtsMescladas=dedupeAgendaRowsByDTRef([...dtsH,...dtsA]);
  document.getElementById('p2info').innerHTML=
    '✅ <b style="color:#22c55e">'+agendRows.length+' linhas</b> encontradas · '+
    '<b style="color:#60a5fa">'+dtsH.length+'</b> hoje · <b style="color:#a78bfa">'+dtsA.length+'</b> amanhã'+
    (Object.keys(exportMap).length>0?
      ' · <b style="color:#4ade80">✨ '+Object.keys(exportMap).length+' DTs com materiais já importados</b>':
      '');
  const mk=(arr,id,cls)=>{
    const w=document.getElementById(id);w.innerHTML='';
    if(!arr.length){w.innerHTML='<span style="font-size:11px;color:#334155">Nenhuma</span>';return;}
    arr.forEach(r=>{const s=document.createElement('span');s.className='chip'+cls;s.textContent=r.DT;w.appendChild(s);});
  };
  mk(dtsH,'chipsHoje','');
  mk(dtsA,'chipsAmanha',' old');
  setStep(2);
}

/* Copia DTs em coluna para colar no SAP */
function copyDTs(dia){
  const arr = dia==='hoje'
    ? agendaRowsForDia(today())
    : agendaRowsForDia(tomorrow());
  const texto = arr.map(r=>r.DT).join('\n');
  navigator.clipboard.writeText(texto).then(()=>{
    const btn=document.getElementById('copy-'+dia+'-btn');
    btn.textContent='✅ Copiado!';
    btn.classList.add('copied');
    setTimeout(()=>{btn.textContent='📋 Copiar DTs em coluna';btn.classList.remove('copied');},2000);
  }).catch(()=>showErr('Não foi possível copiar. Use Ctrl+C manualmente.'));
}

/* ═══════════════════════════════════════════════════════
   AUTO-IMPORT DE MATERIAIS DO ARQUIVO DO PASSO 1 (XLSX)
   Detecta colunas: Nº transporte | Material | Qtde Remessa
═══════════════════════════════════════════════════════ */
function tryImportMaterialsFromAgend(buf){
  try{
    const wb=XLSX.read(new Uint8Array(buf),{type:'array'});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
    if(!rows.length) return 0;
    const keys=Object.keys(rows[0]);
    const fc=(...ts)=>{for(const t of ts){const k=keys.find(k=>k.toUpperCase().includes(t.toUpperCase()));if(k)return k;}return null;};
    const kDT=fc('Nº TRANSPORTE','N TRANSPORTE','TRANSPORTE','NR. TRANSPORTE');
    const kMat=fc('MATERIAL');
    const kQtd=fc('QTDE REMESSA','QTD REMESSA','QTDE');
    if(!kDT||!kMat) return 0; // não tem colunas de materiais, ignora silenciosamente
    let cnt=0;
    for(const row of rows){
      const rawDT=row[kDT];
      const dt=normalizeDT(rawDT);
      const rawMat=row[kMat];
      const mat=(rawMat===''||rawMat===null||rawMat===undefined)?'':String(Math.round(Number(rawMat)));
      const qtdRaw=kQtd?String(row[kQtd]||''):'';
      const qtd=qtdRaw.replace(/\./g,'').replace(',','.');
      const qtdFmt=qtd?String(Math.round(Number(qtd)||0)):'';
      if(!dt||!mat||dt==='NaN'||mat==='NaN') continue;
      cnt++;
      if(!exportMap[dt])exportMap[dt]=[];
      const existing=exportMap[dt].find(m=>m.material===mat);
      if(existing){
        const q1=Number(existing.quantidade)||0;
        const q2=Number(qtdFmt)||0;
        existing.quantidade=String(q1+q2);
      }else{
        exportMap[dt].push({material:mat,quantidade:qtdFmt});
      }
    }
    return cnt;
  }catch(e){return 0;}
}


/* ═══════════════════════════════════════════════════════
   PASSO 3 — RELATÓRIO DE EXPEDIÇÃO CSV
   Lê CSV do relatorio_expedicao:
   Nº transporte, Descrição de Documento, Peso líquido,
   Nr. Remessa/Recebimento, Material, Qtde Remessa
═══════════════════════════════════════════════════════ */

function parseCSVRelatorio(text, sep) {
  const rows = [];
  const lines = text.split(/\r?\n/);
  for (const line of lines) {
    if (!line.trim()) continue;
    const cols = [];
    let cur = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') {
        if (inQ && line[i+1] === '"') { cur += '"'; i++; }
        else inQ = !inQ;
      } else if (c === sep && !inQ) {
        cols.push(cur); cur = '';
      } else { cur += c; }
    }
    cols.push(cur);
    rows.push(cols);
  }
  return rows;
}

function processRelatorioCSV(file) {
  if (!file) return;
  fileWorkflowInProgress=true;
  pauseAutoSync(180000);
  showInf('Lendo CSV do Relatório de Expedição…');
  const dz2Label=document.getElementById('dz2l');
  if(dz2Label) dz2Label.innerHTML = '<span class="dz-ok">✓ ' + file.name + '</span>';
  const reader = new FileReader();
  reader.onload = async function(e) {
    try {
      const text = e.target.result;
      const rows = parseCSVRelatorio(text, ';');
      if (rows.length < 2) throw new Error('Arquivo vazio ou sem dados.');
      const headers = rows[0].map(h => h.trim().replace(/^"|"$/g,''));
      const ci = n => headers.findIndex(h => h.toUpperCase().includes(n.toUpperCase()));
      const iTransp   = ci('Nº transporte') !== -1 ? ci('Nº transporte') : ci('TRANSPORTE');
      const iDesc     = ci('Descrição de Documento') !== -1 ? ci('Descrição de Documento') : ci('DESCRI');
      const iPeso     = ci('Peso líquido') !== -1 ? ci('Peso líquido') : ci('PESO');
      const iRemessa  = ci('Nr. Remessa') !== -1 ? ci('Nr. Remessa') : ci('REMESSA');
      const iMat      = ci('Material');
      const iQtde     = ci('Qtde Remessa') !== -1 ? ci('Qtde Remessa') : ci('QTDE');
      const iHora     = ci('Hora') !== -1 ? ci('Hora') : ci('HORA');
      const iSap      = ci('SAP') !== -1 ? ci('SAP') : ci('Número SAP') !== -1 ? ci('Número SAP') : ci('NR SAP') !== -1 ? ci('NR SAP') : ci('Nº SAP') !== -1 ? ci('Nº SAP') : -1;
      const iCentro   = ci('Centro') !== -1 ? ci('Centro') : ci('CENTRO') !== -1 ? ci('CENTRO') : ci('Ctr') !== -1 ? ci('Ctr') : -1;
      const iInfoAgenda = ci('Inf. Agenda Entrega') !== -1 ? ci('Inf. Agenda Entrega') : ci('AGENDA ENTREGA');

      if (iTransp === -1) throw new Error('Coluna Nº transporte não encontrada.\nColunas: ' + headers.join(' | '));

      exportMap = {}; tipoOpMap = {}; descDocMap = {}; centroMap = {}; infoAgendaMap = {};
      remessaMap = {}; pesoLiquidoMap = {}; horaChegadaCSVMap = {}; sapNumMap = {}; paletizacaoMap = {};
      let linhasOk = 0;

      const records = rows.slice(1).filter(r => r.some(c => c.trim() !== ''));
      for (const cols of records) {
        const strip = v => (v||'').trim().replace(/^"|"$/g,'');
        const dt = normalizeDT(strip(cols[iTransp]).replace(/\.0+$/, '').replace(/\D/g, ''));
        if (!dt || dt.length < 5) continue;

        const descDoc = iDesc !== -1 ? strip(cols[iDesc]) : '';
        const pesoRaw = iPeso !== -1 ? strip(cols[iPeso]) : '';
        const remessa = iRemessa !== -1 ? strip(cols[iRemessa]).replace(/\.0+$/, '') : '';
        const matRaw  = iMat !== -1 ? strip(cols[iMat]).replace(/\.0+$/, '').replace(/\D/g,'') : '';
        const qtdeRaw = iQtde !== -1 ? strip(cols[iQtde]) : '';
        const qtde    = qtdeRaw.replace(/\./g,'').replace(',','.');
        const horaCsv = iHora !== -1 ? normalizeHora(strip(cols[iHora])) : '';
        const sapCsv  = iSap !== -1 ? normalizeSap(strip(cols[iSap])) : '';
        const centroRaw = iCentro !== -1 ? String(cols[iCentro]||'').trim().replace(/\.0+$/,'') : '';
        const infoAgenda = iInfoAgenda !== -1 ? String(cols[iInfoAgenda]||'') : '';
        if (infoAgenda && !infoAgendaMap[dt]) infoAgendaMap[dt] = infoAgenda;
        paletizacaoMap[dt] = mergePaletizacaoLabel(
          paletizacaoMap[dt],
          infoAgenda
        );

        if (horaCsv) horaChegadaCSVMap[dt] = horaCsv;
        if (sapCsv) sapNumMap[dt] = sapCsv;
        if (centroRaw && !centroMap[dt]) centroMap[dt] = centroRaw;

        // descricao_documento e tipo operacao — centro 1111=ARUJA, 1110=MOGI
        if (!descDocMap[dt]) descDocMap[dt] = descDoc;
        if (!tipoOpMap[dt]) {
          const raw = String(descDoc || '').toUpperCase();
          const localRaw = String((cols[ci('Local')]||'')).trim().toUpperCase();
          const isMogi  = centroRaw === '1110';
          const isAruja = centroRaw === '1111' || centroRaw === '';
          const isTransf = raw.includes('TRANSFER') || raw.includes('FILIAL') || raw.includes('ABAST') || raw.includes('TNF');
          const isPrefat = raw.includes('PREFAT') || raw.includes('PRÉ-FAT') || raw.includes('PRE FAT');
          if (isTransf) {
            tipoOpMap[dt] = 'TRANSFERÊNCIA';
          } else if (isPrefat) {
            tipoOpMap[dt] = 'PREFATURA';
          } else {
            const mogiByLocal = localRaw.includes('MOGI');
            const arujaByLocal = localRaw.includes('ARUJA');
            tipoOpMap[dt] = (isMogi||mogiByLocal) ? 'VENDA MOGI' : ((isAruja||arujaByLocal) ? 'VENDA ARUJA' : 'VENDA NORMAL');
          }
        }

        // peso liquido — soma por DT
        if (pesoRaw) {
          const p = parseCargaNumber(pesoRaw);
          pesoLiquidoMap[normalizeDT(dt)] = (pesoLiquidoMap[normalizeDT(dt)] || 0) + p;
        }

        // remessas por transporte (para aba Materiais / Quantidades)
        if (remessa && matRaw) {
          if (!remessaMap[dt]) remessaMap[dt] = [];
          remessaMap[dt].push({ remessa, material: matRaw, qtde: qtde || qtdeRaw });
        }

        // exportMap — material+qtde por DT (para painel de separação)
        if (matRaw) {
          if (!exportMap[dt]) exportMap[dt] = [];
          const existing = exportMap[dt].find(m => m.material === matRaw);
          const q2 = parseFloat(qtde) || 0;
          if (existing) {
            existing.quantidade = String((Number(existing.quantidade) || 0) + q2);
          } else {
            exportMap[dt].push({ material: matRaw, quantidade: String(q2 || qtdeRaw) });
          }
        }
        linhasOk++;
      }

      const totalDTs = Object.keys(exportMap).length;
      const status = document.getElementById('upload-status-p3');
      if (status) status.innerHTML = `<b style="color:#4ade80">✓ ${file.name}</b> — <b style="color:#60a5fa">${totalDTs} transportes</b> · ${linhasOk} linhas processadas`;
      const isInlineUpload = document.getElementById('passo4').style.display !== 'none';
      if(isInlineUpload){
        const importedCount=await persistImportedMaterialsForCurrentTable();
        await persistRelatorioFieldsForCurrentTable();
        await persistImportedMaterialsForCurrentTable();
        hideInf();
        showOk(`Relatório OK — ${totalDTs} transportes · ${linhasOk} linhas · materiais salvos para ${importedCount} DT(s)`);
        await reloadTable();
        setStep(4);
        releaseAutoSyncAfterSave();
      }else{
        hideInf();
        showOk(`Relatório OK — ${totalDTs} transportes · ${linhasOk} linhas`);
        await buildTable();
        releaseAutoSyncAfterSave();
      }
    } catch(err) {
      hideInf();
      showErr('Erro ao ler CSV: ' + err.message);
      console.error(err);
      releaseAutoSyncAfterSave(60000);
    }
  };
  reader.onerror = () => { hideInf(); showErr('Erro ao abrir o arquivo.'); releaseAutoSyncAfterSave(60000); };
  reader.readAsText(file, 'ISO-8859-1');
}


async function persistImportedMaterials(cargaRows=null){
  const dtsWithAgenda=new Map();
  if(Array.isArray(cargaRows)){
    cargaRows.forEach(r=>{
      const dt=normalizeDT(r&&r.dt);
      const ref=String((r&&r.data_ref)||'');
      if(dt&&ref) dtsWithAgenda.set(dt,ref);
    });
  }else{
    dtsMescladas.forEach(r=>{
      const dt=normalizeDT(r&&r.DT);
      const ref=dKey(agendaRefDate(r)||today());
      if(dt&&ref) dtsWithAgenda.set(dt,ref);
    });
  }
  const inserts=[];
  for(const [dt,mats] of Object.entries(exportMap)){
    const dataRef=dtsWithAgenda.get(normalizeDT(dt));
    if(!dataRef||!Array.isArray(mats)) continue;
    mats.forEach((m,i)=>{
      if(!m.material) return;
          inserts.push({dt:String(dt),data_ref:String(dataRef),material:String(m.material),quantidade:String(m.quantidade||''),observacao:'',ordem:i,paletizacao:paletizacaoMap[String(dt)]||'ESTIVADA'});
    });
  }
  if(!inserts.length) return 0;
  const pares=[...new Set(inserts.map(m=>`${m.dt}__${m.data_ref}`))];
  for(const p of pares){
    const [dt,data_ref]=p.split('__');
    await sbDelete('reporte_materiais',{dt,data_ref});
  }
  await sbInsert('reporte_materiais',inserts);
  return pares.length;
}

function hasPesoMateriais(dt){
  return Object.prototype.hasOwnProperty.call(pesoLiquidoMap,normalizeDT(dt));
}

function pesoMateriaisFinal(dt){
  return hasPesoMateriais(dt)?String(toTonInt(pesoLiquidoMap[normalizeDT(dt)])):'';
}

function getCurrentTableRefsByDT(){
  const refsByDT=new Map();
  tableData.forEach(r=>{
    const dt=normalizeDT(r.dt);
    const ref=String(r.data_ref||'');
    if(!dt||!ref) return;
    if(!refsByDT.has(dt)) refsByDT.set(dt,[]);
    if(!refsByDT.get(dt).includes(ref)) refsByDT.get(dt).push(ref);
  });
  return refsByDT;
}

async function persistImportedMaterialsForCurrentTable(){
  const refsByDT=getCurrentTableRefsByDT();
  const inserts=[];
  for(const [dt,mats] of Object.entries(exportMap)){
    const refs=refsByDT.get(normalizeDT(dt))||[];
    if(!refs.length||!Array.isArray(mats)) continue;
    refs.forEach(dataRef=>{
      mats.forEach((m,i)=>{
        if(!m.material) return;
        inserts.push({dt:String(dt),data_ref:String(dataRef),material:String(m.material),quantidade:String(m.quantidade||''),observacao:'',ordem:i,paletizacao:paletizacaoMap[String(dt)]||'ESTIVADA'});
      });
    });
  }
  const pares=[...new Set(inserts.map(m=>`${m.dt}__${m.data_ref}`))];
  for(const p of pares){
    const [dt,data_ref]=p.split('__');
    await sbDelete('reporte_materiais',{dt,data_ref});
  }
  if(inserts.length) await sbInsert('reporte_materiais',inserts);
  return pares.length;
}

async function persistRelatorioFieldsForCurrentTable(){
  const refsByDT=getCurrentTableRefsByDT();
  for(const [dtRaw,refs] of refsByDT.entries()){
    const dt=normalizeDT(dtRaw);
    const patch={updated_at:new Date().toISOString()};
    if(descDocMap[dt]) patch.descricao_documento=String(descDocMap[dt]);
    if(tipoOpMap[dt]) patch.tipo_operacao=String(tipoOpMap[dt]);
    if(centroMap[dt]) patch.centro=String(centroMap[dt]);
    if(paletizacaoMap[dt]) patch.paletizacao=String(paletizacaoMap[dt]);
    // Peso é exclusivamente do relatório/materiais importado; sem material, limpa o peso.
    const peso=pesoMateriaisFinal(dt);
    patch.peso_liquido=peso;
    patch.toneladas=peso;
    if(horaChegadaCSVMap[dt]) patch.hora_chegada=String(horaChegadaCSVMap[dt]);
    if(sapNumMap[dt]) patch.n_portaria=String(sapNumMap[dt]);
    if(Object.keys(patch).length===1) continue;
    for(const dataRef of refs){
      await sbPatch('reporte_carga',patch,{dt,data_ref:dataRef});
      const row=tableData.find(r=>normalizeDT(r.dt)===dt && String(r.data_ref||'')===String(dataRef));
      if(row) Object.assign(row,patch);
    }
  }
}

async function saveUploadSnapshot(rows){
  try{
    const refs=[...new Set(rows.map(r=>String(r.data_ref||'')).filter(Boolean))];
    const uploadTs=new Date().toISOString();
    const logs=refs.map(ref=>({
      dt:'__AGENDA__',
      data_ref:ref,
      event_type:'UPLOAD_AGENDA',
      status_after:'UPLOAD_AGENDA',
      transportadora:'AGENDA_ATIVA',
      hora_evento:uploadTs,
      source:'DASHBOARD',
      payload:{refs,total_dts:rows.length},
    }));
    if(logs.length){
      for(const log of logs) await tryInsertLog(log);
    }
  }catch(e){}
}

/* ═══════════════════════════════════════════════════════
   BUILD TABLE — Upsert preservando status/hora_chegada
═══════════════════════════════════════════════════════ */
async function registrarSaidasGrade(rows,{origem='UPLOAD_DIFF'}={}){
  const saidas=(rows||[]).filter(r=>r&&r.dt&&r.data_ref&&!isFinalStatus(r.status)&&String(r.status||'')!=='DT EXCLUIDA');
  if(!saidas.length) return 0;
  const now=new Date().toISOString();
  for(const row of saidas){
    const dt=normalizeDT(row.dt);
    const dataRef=String(row.data_ref||'');
    await sbPatch('reporte_carga',{status:'DT EXCLUIDA',updated_at:now,source:origem},{dt,data_ref:dataRef});
    await tryInsertLog({
      dt,
      data_ref:dataRef,
      event_type:'DT_EXCLUIDA_GRADE',
      status:'DT EXCLUIDA',
      transportadora:row.transportadora||'Saiu da grade ativa',
      created_at:now,
      source:origem,
      payload:{status_before:row.status||'',grade_carregamento:row.grade_carregamento||'',fim_carregamento:row.fim_carregamento||''},
    });
  }
  return saidas.length;
}

function agendaScheduleKey(row){
  const fim=parseBR(String((row&&row.fim_carregamento)||(row&&row.fim_agenda)||''));
  return fim?String(fim.getTime()):String((row&&row.fim_carregamento)||(row&&row.fim_agenda)||'').trim();
}

function hasAgendaScheduleChanged(oldRow,newRow){
  if(!oldRow||!newRow) return false;
  const oldRef=String(oldRow.data_ref||'');
  const newRef=String(newRow.data_ref||'');
  if(oldRef&&newRef&&oldRef!==newRef) return true;
  return agendaScheduleKey(oldRow)!==agendaScheduleKey(newRow);
}

function buildExistingByDT(rows){
  const byDT=new Map();
  (rows||[]).forEach(row=>{
    const dt=normalizeDT(row&&row.dt);
    if(!dt) return;
    if(!byDT.has(dt)) byDT.set(dt,[]);
    byDT.get(dt).push(row);
  });
  return byDT;
}

function bestExistingForIncoming(incoming,byDT){
  const dt=normalizeDT(incoming&&incoming.dt);
  const rows=(byDT&&byDT.get(dt))||[];
  if(!rows.length) return null;
  const sameRef=rows.find(r=>String(r.data_ref||'')===String(incoming.data_ref||''));
  if(sameRef) return sameRef;
  return rows.find(r=>!isFinalStatus(r.status)&&String(r.status||'')!=='DT EXCLUIDA') || rows[0];
}

async function registrarReagendamentosAgenda(incomingRows,existingRows){
  const byDT=buildExistingByDT(existingRows);
  const now=new Date().toISOString();
  let total=0;
  for(const incoming of incomingRows||[]){
    const dt=normalizeDT(incoming&&incoming.dt);
    const oldRow=bestExistingForIncoming(incoming,byDT);
    if(!dt||!oldRow||!hasAgendaScheduleChanged(oldRow,incoming)) continue;
    const oldRef=String(oldRow.data_ref||'');
    const newRef=String(incoming.data_ref||'');
    await tryInsertLog({
      dt,
      data_ref:newRef||oldRef,
      event_type:'REAGENDADA',
      status_before:oldRow.status||'',
      status_after:'REAGENDADA',
      transportadora:incoming.transportadora||oldRow.transportadora||'DT reagendada na agenda',
      grade_carregamento:incoming.grade_carregamento||'',
      fim_carregamento:incoming.fim_carregamento||'',
      created_at:now,
      source:'UPLOAD_DIFF',
      payload:{
        data_ref_before:oldRef,
        data_ref_after:newRef,
        grade_carregamento:incoming.grade_carregamento||incoming.agenda||'',
        fim_carregamento:incoming.fim_carregamento||'',
        grade_carregamento_before:oldRow.grade_carregamento||oldRow.agenda||'',
        fim_carregamento_before:oldRow.fim_carregamento||'',
        grade_carregamento_after:incoming.grade_carregamento||incoming.agenda||'',
        fim_carregamento_after:incoming.fim_carregamento||'',
        status_before:oldRow.status||'',
        observacao:'DT já existia no banco com outra data/horário; tratada como reagendamento, não como DT excluída.'
      },
    });
    if(oldRef&&newRef&&oldRef!==newRef){
      await sbDelete('reporte_carga',{dt:String(oldRow.dt||dt),data_ref:oldRef});
    }
    total++;
  }
  return total;
}

async function markMissingRowsAsExcluded(uploadedRefs, incomingRows, existingRows){
  // REGRA 5: NUNCA remover automaticamente DTs pendentes (GAP).
  // Apenas registrar em log que a DT saiu da nova agenda.
  // O usuário deve excluir manualmente quando necessário.
  const incomingKeys=new Set((incomingRows||[]).map(r=>`${normalizeDT(r.dt)}__${String(r.data_ref||'')}`));
  const incomingDTs=new Set((incomingRows||[]).map(r=>normalizeDT(r.dt)).filter(Boolean));
  const refs=new Set((uploadedRefs||[]).map(String));
  const now=new Date().toISOString();
  for(const row of existingRows||[]){
    const dataRef=String(row.data_ref||'');
    const dt=normalizeDT(row.dt);
    if(!dt || !refs.has(dataRef)) continue;
    if(incomingKeys.has(`${dt}__${dataRef}`)) continue;
    if(incomingDTs.has(dt)) continue; // mesma DT veio em outro horário/data: é reagendamento, não exclusão
    const st=String(row.status||'').trim();
    if(isFinalStatus(st) || st==='DT EXCLUIDA') continue;
    // GAP: DT pendente que sumiu da nova agenda → apenas logar, NÃO excluir automaticamente
    await tryInsertLog({dt,data_ref:dataRef,event_type:'DT_SAIU_GRADE',status:st,transportadora:row.transportadora||'Saiu da nova agenda',created_at:now,source:'UPLOAD_DIFF',payload:{status_before:st,grade_carregamento:row.grade_carregamento||'',fim_carregamento:row.fim_carregamento||'',observacao:'GAP: DT pendente não encontrada na nova agenda — mantida no sistema conforme Regra 5'}});
  }
}

async function buildTable(){
  fileWorkflowInProgress=true;
  pauseAutoSync(180000);
  showInf('Sincronizando com o banco…');
  const T=today(),AM=tomorrow();
  const kT=dKey(T), kAM=dKey(AM);
  try{
    // Busca existentes por data de FIM e também por DT para limpar registros antigos gravados pelo início.
    const refsUpload=[...new Set(dtsMescladas.map(r=>dKey(agendaRefDate(r)||T)))];
    const dtsUpload=[...new Set(dtsMescladas.map(r=>normalizeDT(r.DT)).filter(Boolean))];
    const selectCarga='dt,data_ref,status,hora_chegada,n_portaria,tipo_operacao,descricao_documento,centro,reagendada,peso_liquido,toneladas,doca_null,transportadora,grade_carregamento,fim_carregamento,agenda';
    const refsExisting=[...new Set([...refsUpload,kT,kAM].filter(Boolean))];
    const existingByRef=await sbGet('reporte_carga',
      `data_ref=${sbIn(refsExisting)}&select=${selectCarga}`
    );
    let existingByDt=[];
    if(dtsUpload.length){
      existingByDt=await sbGet('reporte_carga',
        `dt=${sbIn(dtsUpload)}&select=${selectCarga}`
      );
    }
    const existingMap=new Map();
    [...(existingByRef||[]),...(existingByDt||[])].forEach(r=>{
      existingMap.set(`${normalizeDT(r.dt)}__${String(r.data_ref||'')}`,r);
    });
    const existing=[...existingMap.values()];
    const exMap={};
    const exByDT=buildExistingByDT(existing||[]);
    (existing||[]).forEach(r=>{exMap[normalizeDT(r.dt)+'_'+r.data_ref]=r;});
    const logsStatusMap={};
    const finalStatusSet=new Set(STATUS_FINAIS);
    const finalizedDTs=new Set();
    (existing||[]).forEach(r=>{
      const dtEx=normalizeDT(r.dt);
      const st=String(r.status||'').trim();
      if(dtEx&&isFinalStatus(st)) finalizedDTs.add(dtEx);
    });
    if(dtsUpload.length){
      let logRows=[];
      try{
        logRows=await sbGet('dt_logs',
          `dt=${sbIn(dtsUpload)}&select=dt,event_type,status_after,hora_evento&order=hora_evento.desc&limit=1000`
        );
      }catch(e){
        if(!isMissingLogsError(e)) throw e;
        try{
          logRows=await sbGet('reporte_logs',
            `dt=${sbIn(dtsUpload)}&select=dt,status,created_at&order=created_at.desc&limit=1000`
          );
        }catch(e2){
          if(!isMissingLogsError(e2)) throw e2;
        }
      }
      (logRows||[]).forEach(l=>{
        const dtLog=normalizeDT(l.dt);
        const st=String(l.status_after||l.status||'').trim();
        if(!dtLog||!st) return;
        if(st==='UPLOAD_AGENDA'||st==='REAGENDADA') return;
        if(isFinalStatus(st)){
          finalizedDTs.add(dtLog);
          return;
        }
        if(logsStatusMap[dtLog]) return;
        logsStatusMap[dtLog]=st;
      });
    }


    const rows=dedupeCargaRowsByDTRef(dtsMescladas.filter(dt=>!finalizedDTs.has(normalizeDT(dt.DT))).map(dt=>{
      const refDate=agendaRefDate(dt)||T;
      const ref=dKey(refDate);
      const dtNorm=normalizeDT(dt.DT);
      const sameRefEx=exMap[dtNorm+'_'+ref]||{};
      const fallbackEx=(exByDT.get(dtNorm)||[]).find(r=>!isFinalStatus(r.status)&&String(r.status||'')!=='DT EXCLUIDA')||{};
      const ex=Object.keys(sameRefEx).length?sameRefEx:fallbackEx;
      const diaRef=sameDay(refDate,T)?'HOJE':'AMANHÃ';
      const tipoOp=tipoOpMap[dt.DT]||(ex.tipo_operacao&&ex.tipo_operacao!==''?ex.tipo_operacao:'');
      const pesoFinal = pesoMateriaisFinal(dt.DT);
      const docaNullFinal = dt.DOCA_NULL ? ex.doca_null !== false : false;
      // Preserva status/hora_chegada e liberações manuais; peso não tem fallback: só materiais.
      const statusAtual = ex.status && ex.status !== '' ? ex.status : (logsStatusMap[normalizeDT(dt.DT)]||'AG CHEGADA');
      const horaAtual   = ex.hora_chegada || '';
      return {
        dt:String(dt.DT),
        transportadora:String(dt.TRANSPORTADORA||''),
        grade_carregamento:String(dt.AGENDA?fmtDT(dt.AGENDA,true):''),
        fim_carregamento:String(dt.FIM_AGENDA?fmtDT(dt.FIM_AGENDA,true):''),
        hora_chegada:horaAtual || horaChegadaCSVMap[dt.DT] || '',
        n_portaria:sapNumMap[dt.DT] || ex.n_portaria || '',
        status:statusAtual,
        descricao_documento:String(descDocMap[dt.DT]||ex.descricao_documento||''),
        centro:String(centroMap[dt.DT]||ex.centro||''),
        toneladas:pesoFinal,
        peso_liquido:pesoFinal,
        agenda:String(dt.AGENDA?fmtDT(dt.AGENDA,true):''),
        local_cd:String(dt.LOCAL||''),
        dia_ref:diaRef,
        data_ref:String(ref),
        tipo_operacao:String(tipoOp),
        reagendada: ex.reagendada || false,
        doca_null: docaNullFinal,
      };
    }));

    // REGRA 9: detectar DTs novas (não existiam no banco antes desse upload)
    novoDTs = new Set(); // reset a cada upload
    const existingDTSet = new Set((existing||[]).map(r=>normalizeDT(r.dt)).filter(Boolean));
    rows.forEach(r=>{ if(!existingDTSet.has(normalizeDT(r.dt))) novoDTs.add(normalizeDT(r.dt)); });

    const uploadedRefs=[...new Set(refsUpload.filter(Boolean))];
    const activeRowRefs=[...new Set(rows.map(r=>r.data_ref).filter(Boolean))];
    activeRefsOverride=(activeRowRefs.length?activeRowRefs:uploadedRefs).slice(0,2);
    // REGRA UNICIDADE DT: apaga do banco qualquer linha com mesmo DT mas data_ref diferente
    // (evita duplicatas ao re-subir a grade com horários alterados).
    const staleRows=[];
    rows.forEach(row=>{
      const dtNorm=normalizeDT(row.dt);
      const expectedRef=String(row.data_ref||'');
      (exByDT.get(dtNorm)||[]).forEach(ex=>{
        if(String(ex.data_ref||'')!==expectedRef) staleRows.push({dt:String(ex.dt||dtNorm),data_ref:String(ex.data_ref||'')});
      });
    });
    for(const s of staleRows){
      try{await sbDelete('reporte_carga',s);}catch(e){console.warn('Falha ao limpar data_ref obsoleto:',s,e);}
    }
    // Como não existe constraint única (dt,data_ref), fazemos replace por data_ref
    await markMissingRowsAsExcluded(uploadedRefs,rows,existing);
    if(rows.length) await sbUpsert('reporte_carga',rows,'dt,data_ref');
    await registrarReagendamentosAgenda(rows,existing);
    await saveUploadSnapshot(rows);
    const importedCount=await persistImportedMaterials(rows);
    hideInf();
    if(importedCount>0) showOk('Materiais sincronizados automaticamente para '+importedCount+' DT(s).');
    await reloadTable();
    setStep(4);
    releaseAutoSyncAfterSave();
  }catch(e){
    hideInf();
    showErr('Erro banco (tabela ainda aparece): '+e.message);
    tableData=dtsMescladas.map(dt=>({
      dt:dt.DT,transportadora:dt.TRANSPORTADORA,
      grade_carregamento:dt.AGENDA?fmtDT(dt.AGENDA,true):'',
      fim_carregamento:dt.FIM_AGENDA?fmtDT(dt.FIM_AGENDA,true):'',
      hora_chegada:horaChegadaCSVMap[dt.DT]||'',n_portaria:sapNumMap[dt.DT]||'',status:'AG CHEGADA',descricao_documento:descDocMap[dt.DT]||'',centro:centroMap[dt.DT]||'',toneladas:pesoMateriaisFinal(dt.DT),
      peso_liquido: pesoMateriaisFinal(dt.DT),
      agenda:dt.AGENDA?fmtDT(dt.AGENDA,true):'',
      dia_ref:sameDay(agendaRefDate(dt)||T,T)?'HOJE':'AMANHÃ',
      data_ref:dKey(agendaRefDate(dt)||T),
      tipo_operacao:tipoOpMap[dt.DT]||'',
      reagendada:false,
    }));
    renderRows();
    setStep(4);
    releaseAutoSyncAfterSave();
  }
}


async function getActiveRefs(){
  if(activeRefsOverride.length) return activeRefsOverride;
  const ON=yesterday(),T=today(),AM=tomorrow();
  const base=[dKey(T),dKey(AM)];
  try{
    // Hoje/amanhã têm prioridade quando já existe agenda atual no banco.
    const todayRows=await sbGet('reporte_carga',`data_ref=eq.${base[0]}&select=dt&limit=1`);
    if(todayRows&&todayRows.length) return base;

    // Depois de cada upload salvamos marcadores __AGENDA__; eles evitam reabrir uma grade antiga (ex.: 05/05).
    let markers=[];
    try{
      markers=await sbGet('dt_logs','dt=eq.__AGENDA__&event_type=eq.UPLOAD_AGENDA&select=data_ref,hora_evento&order=hora_evento.desc&limit=20');
    }catch(e){
      if(!isMissingLogsError(e)) throw e;
      markers=await sbGet('reporte_logs','dt=eq.__AGENDA__&status=eq.UPLOAD_AGENDA&select=data_ref,created_at&order=created_at.desc&limit=20');
    }
    const latestTs=markers&&markers.length?String(markers[0].hora_evento||markers[0].created_at||''):null;
    if(latestTs){
      const active=[];
      (markers||[]).forEach(r=>{
        if(String(r.hora_evento||r.created_at||'')!==latestTs) return;
        const k=String(r.data_ref||'');
        if(k&&!active.includes(k)) active.push(k);
      });
      if(active.length) return active.slice(0,2);
    }

    // Fallback legado: usa os data_ref mais recentes em logs de upload existentes.
    let lastUpload=[];
    try{
      lastUpload=await sbGet('dt_logs',`event_type=eq.UPLOAD_AGENDA&select=data_ref,hora_evento&order=hora_evento.desc&limit=300`);
    }catch(e){
      if(!isMissingLogsError(e)) throw e;
      lastUpload=await sbGet('reporte_logs',`status=eq.UPLOAD_AGENDA&select=data_ref,created_at&order=created_at.desc&limit=300`);
    }
    const byUpload=[];
    (lastUpload||[]).forEach(r=>{
      const k=String(r.data_ref||'');
      if(k && !byUpload.includes(k)) byUpload.push(k);
    });
    if(byUpload.length) return byUpload.slice(0,2);

    const rec=await sbGet('reporte_carga','select=data_ref&limit=1000');
    const uniq=[...new Set((rec||[]).map(r=>String(r.data_ref||'')).filter(Boolean))];
    const toDate=(k)=>{const m=String(k).match(/(\d{2})\/(\d{2})\/(\d{4})/);return m?new Date(+m[3],+m[2]-1,+m[1]):new Date(0);};
    uniq.sort((a,b)=>toDate(b)-toDate(a));
    return uniq.slice(0,2);
  }catch(e){
    if(isMissingReporteLogsError(e)) return base;
    return base;
  }
}

async function reloadTable(){
  try{
    const refs=await getActiveRefs();
    if(!refs.length){tableData=[]; if(currentTableTab==='reporte') renderReporte(); else renderRows(); return;}
    const data=await sbGet('reporte_carga',
      `data_ref=${sbIn(refs)}&order=dia_ref.asc,agenda.asc`
    );
    tableData=normalizeDiaRefRows(data||[]).sort((a,b)=>{
      return compareAgendaRows(a,b);
    });
    await enrichRowsWithPaletizacao(tableData);
    await persistNormalizedDiaRefs(tableData);
    renderRows();
  }catch(e){showErr('Erro ao carregar tabela: '+e.message);}
}

async function enrichRowsWithPaletizacao(rows){
  if(!rows||!rows.length) return rows;
  const refs=[...new Set(rows.map(r=>String(r.data_ref||'')).filter(Boolean))];
  if(!refs.length) return rows;

  const byDTRef={};
  const byDT={};
  try{
    const mats=await sbGet('reporte_materiais',
      `data_ref=${sbIn(refs)}&select=dt,data_ref,paletizacao`
    );
    (mats||[]).forEach(m=>{
      const dt=normalizeDT(m.dt);
      const dataRef=String(m.data_ref||'');
      if(!dt||!dataRef) return;
      const key=dt+'__'+dataRef;
      byDTRef[key]=mergePaletizacaoLabel(byDTRef[key],m.paletizacao);
      byDT[dt]=mergePaletizacaoLabel(byDT[dt],m.paletizacao);
    });

    const allDTs=[...new Set(rows.map(r=>normalizeDT(r.dt)).filter(Boolean))];
    if(allDTs.length){
      const extra=await sbGet('reporte_materiais',
        `dt=${sbIn(allDTs)}&select=dt,paletizacao`
      );
      (extra||[]).forEach(m=>{
        const dt=normalizeDT(m.dt);
        if(!dt) return;
        byDT[dt]=mergePaletizacaoLabel(byDT[dt],m.paletizacao);
      });
    }
  }catch(e){
    console.warn('Nao foi possivel carregar paletizacao dos materiais.',e);
  }

  rows.forEach(r=>{
    const dt=normalizeDT(r.dt);
    const key=dt+'__'+String(r.data_ref||'');
    r.paletizacao=mergePaletizacaoLabel(
      mergePaletizacaoLabel(r.paletizacao,byDTRef[key]),
      byDT[dt]
    );
  });
  return rows;
}

/* ═══════════════════════════════════════════════════════
   RELÓGIO — verifica se deve mostrar emoji
   Regra: se hoje e status não é final e grade definida,
   mostra 🕐 se falta ≤1h para o início da grade
═══════════════════════════════════════════════════════ */
function clockEmoji(row){
  if(row.dia_ref!=='HOJE') return '';
  if(row.n_portaria) return '';
  if(isFinalStatus(row.status)) return '';
  const agora=new Date();
  // 🚛 Carreta deve sair: fim_carregamento chegando ou passado (para DTs pendentes)
  if(row.fim_carregamento){
    const fim=parseBR(row.fim_carregamento);
    if(fim){
      const diffFimMin=(fim-agora)/60000;
      const fimStr=fim.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'});
      if(diffFimMin<=0 && diffFimMin>-120){
        return '<span class="dt-clock" title="🚛 Fim do carregamento atingido ('+fimStr+') — carreta deve sair!">🚛</span>';
      }
      if(diffFimMin>0 && diffFimMin<=30){
        return '<span class="dt-clock" title="🚛 Falta '+Math.round(diffFimMin)+'min para o fim do carregamento ('+fimStr+') — prepare saída!">🚛</span>';
      }
    }
  }
  if(row.status==='CARREGANDO') return '<span class="dt-clock" title="DT em carregamento sem SAP/Portaria. Preencher SAP agora.">🚨</span>';
  if(!row.grade_carregamento) return '';
  const grade=parseBR(row.grade_carregamento);
  if(!grade) return '';
  const chegadaLimite=new Date(grade.getTime()-3600000);
  const diffGradeMin=(grade-agora)/60000;
  if(diffGradeMin>=0 && diffGradeMin<=60){
    return '<span class="dt-clock" title="Falta 1h ou menos para iniciar a grade; carreta deveria chegar até '+chegadaLimite.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})+'">🕐</span>';
  }
  if(diffGradeMin<0 && diffGradeMin>-120) return '<span style="font-size:11px;opacity:.6" title="Grade iniciada sem portaria preenchida">⏰</span>';
  return '';
}

function shouldSapNow(row){
  return !row.n_portaria && row.status==='CARREGANDO';
}


async function renderMaterialsBoard(){
  const board=document.getElementById('mat-board');
  const twrap=document.getElementById('twrap');
  if(!board||!twrap) return;
  twrap.style.display='none';
  board.style.display='block';
  board.innerHTML='<div style="color:#64748b;">Carregando materiais…</div>';
  try{
    const refs=[...new Set(tableData.map(r=>r.data_ref).filter(Boolean))];
    if(!refs.length){board.innerHTML='<div style="color:#334155;">Sem DTs carregadas.</div>';return;}
    const q='data_ref='+sbIn(refs)+'&order=dt.asc,ordem.asc';
    const mats=await sbGet('reporte_materiais',q);
    if(!mats||!mats.length){board.innerHTML='<div style="color:#334155;">Nenhum material salvo para as DTs atuais.</div>';return;}
    const byDT={};
    mats.forEach(m=>{
      const key=String(m.dt||'');
      if(!key) return;
      if(!byDT[key]) byDT[key]={};
      const mat=String(m.material||'').trim();
      if(!mat) return;
      const qtd=Number(String(m.quantidade||'0').replace(',','.'))||0;
      byDT[key][mat]=(byDT[key][mat]||0)+qtd;
    });
    const ordem=[...tableData].sort(compareAgendaRows).map(r=>String(r.dt));
    const uniq=[]; ordem.forEach(dt=>{if(!uniq.includes(dt)) uniq.push(dt);});
    let html='';
    uniq.forEach(dt=>{
      const itens=byDT[dt]; if(!itens) return;
      const pairs=Object.entries(itens);
      const totalPaletes=pairs.reduce((s,[mat,q])=>s+formatPaletes(mat,q).paletes,0);
      const totalSobra=pairs.reduce((s,[mat,q])=>s+formatPaletes(mat,q).sobra,0);
      const totalLabel=totalPaletes+' palete(s)'+(totalSobra?' + '+Math.round(totalSobra)+' fardo(s) sobra':'');
      html+='<div class="mat-group"><div class="mat-group-h"><span>DT '+dt+'</span><span>'+pairs.length+' materiais · '+totalLabel+'</span></div><div class="mat-group-b">';
      pairs.forEach(([mat,q])=>{
        const pal=formatPaletes(mat,q);
        const detalhe=pal.porPalete?'<div style="font-size:9px;color:#64748b;font-weight:500;">'+Math.round(q)+' fardos · '+pal.porPalete+'/palete</div>':'';
        html+='<div class="mat-item"><span>'+mat+detalhe+'</span><span style="text-align:right;color:#a7f3d0;font-weight:700;">'+pal.label+'</span></div>';
      });
      html+='</div></div>';
    });
    board.innerHTML=html||'<div style="color:#334155;">Nenhum material encontrado para as DTs listadas.</div>';
  }catch(e){
    board.innerHTML='<div style="color:#ef4444;">Erro ao carregar materiais: '+e.message+'</div>';
  }
}

/* ═══════════════════════════════════════════════════════
   RENDERIZAR
═══════════════════════════════════════════════════════ */
function renderRows(){
  const board=document.getElementById('mat-board');
  const twrap=document.getElementById('twrap');
  const isTableView=currentTableTab==='todas'||currentTableTab==='finalizadas';
  if(board && isTableView) board.style.display='none';
  if(twrap) twrap.style.display=isTableView?'block':'none';
  // Resumo de status
  const sum={};
  tableData.forEach(r=>{sum[r.status]=(sum[r.status]||0)+1;});
  const sw=document.getElementById('tstatus');sw.innerHTML='';
  Object.entries(sum).forEach(([s,n])=>{
    const c=SC[s]||'#334155';
    const sp=document.createElement('span');sp.className='sbadge';
    sp.style.cssText=`background:${c}22;border:1px solid ${c}55;color:${c}`;
    sp.textContent=s+': '+n;sw.appendChild(sp);
  });
  const cH=tableData.filter(r=>r.dia_ref==='HOJE').length;
  const cA=tableData.filter(r=>r.dia_ref==='AMANHÃ').length;
  document.getElementById('tinfo').innerHTML=
    `<span style="color:#60a5fa">${cH} hoje</span> · <span style="color:#a78bfa">${cA} amanhã</span>`;

  // Filtro de aba
  const sortSelect=document.getElementById('table-sort-mode');
  if(sortSelect && sortSelect.value!==tableSortMode) sortSelect.value=tableSortMode;

  let rows=[...tableData].sort(compareTableRows);
  if(currentTableTab==='finalizadas'){
    rows=rows.filter(r=>isFinalStatus(r.status)&&!r.doca_null);
  }else{
    // Exclui doca_null da tabela principal: essas DTs ficam apenas na aba ⚠️ DOCA S/ INFO
    rows=rows.filter(r=>!isFinalStatus(r.status)&&!r.doca_null);
  }
  if(dtSearchTerm){
    rows=rows.filter(r=>String(r.dt||'').includes(dtSearchTerm));
  }

  const tbody=document.getElementById('tbody');tbody.innerHTML='';

  // Atualiza badge da aba finalizadas
  const nFin=tableData.filter(r=>isFinalStatus(r.status)).length;
  document.getElementById('tab-finalizadas').textContent=`🏁 EXPEDIDAS / NO SHOW / RECUSADAS (${nFin})`;
  // Atualiza badge da aba DOCA S/ INFO
  const nDocaNull=tableData.filter(r=>r.doca_null).length;
  const tabDocaNull=document.getElementById('tab-docanull');
  if(tabDocaNull) tabDocaNull.textContent=`⚠️ DOCA S/ INFO${nDocaNull>0?' ('+nDocaNull+')':''}`;

  rows.forEach(row=>{
    const isH=row.dia_ref==='HOJE';
    const c=SC[row.status]||'#334155';
    const tr=document.createElement('tr');
    const isFinal=isFinalStatus(row.status);
    const isReagend=row.reagendada;
    if(isReagend) tr.className='row-reagendada-flag';
    else tr.className=isH?'row-hoje':'row-amanha';
    const opTag=buildTipoOperacaoBadge(row.tipo_operacao,{small:false,emptyDash:true});
    const paletTag=buildPaletizacaoBadge(row.paletizacao);
    const diaTag=isH?'<span class="tag-hoje">HOJE</span>':'<span class="tag-amanha">AMANHÃ</span>';
    const clock=clockEmoji(row);
    // Botão reagendamento — aparece nas finalizadas
    const reagBtn=isFinal
      ? `<button class="bg" style="font-size:10px;padding:3px 9px;border-color:#f59e0b55;color:#f59e0b;" onclick="openReagend('${row.dt}','${row.data_ref}')">🔄 Reagendar</button>`
      : '';
    // REGRA 9 + 10: badge NOVO para DTs recém-adicionadas; borda amarela para DOCA nula
    const isNovoDT = novoDTs.has(normalizeDT(row.dt));
    const novoBadge = isNovoDT ? ' <span style="background:#f59e0b22;border:1px solid #f59e0b88;color:#f59e0b;font-size:9px;font-weight:900;border-radius:4px;padding:1px 5px;letter-spacing:1px;vertical-align:middle;">NOVO ✨</span>' : '';
    const docaNullStyle = row.doca_null ? 'outline:2px solid #f59e0b;outline-offset:1px;' : '';
    const docaReporteToggle = (preFatMode && row.doca_null)
      ? `<label style="font-size:9px;color:#f59e0b;display:flex;align-items:center;gap:3px;margin-top:2px;cursor:pointer;"><input type="checkbox" onchange="toggleDocaNullReporte('${row.dt}','${row.data_ref}',this.checked)" style="accent-color:#f59e0b;"/>REPORTE</label>`
      : '';
    tr.innerHTML=
      `<td>${diaTag}</td>`+
      `<td><span class="td-dt" style="${docaNullStyle}" onclick="openPanel('${row.dt}','${row.data_ref}')">${row.dt}</span>${novoBadge}${row.doca_null?' <span title="DOCA não informada — fora do reporte até liberação manual" style="color:#f59e0b;font-size:11px;cursor:default;">⚠️</span>':''}${clock?` <span style="margin-left:4px;vertical-align:middle;">${clock}</span>`:''}${preFatMode?`<label style="font-size:9px;color:#a78bfa;display:flex;align-items:center;gap:3px;margin-top:2px;cursor:pointer;"><input type="checkbox" ${row.tipo_operacao==='PRÉ-FAT'?'checked':''} onchange="togglePreFat('${row.dt}','${row.data_ref}',this.checked)" style="accent-color:#a78bfa;"/>PRÉ-FAT</label>`:''}${docaReporteToggle}</td>` +
      `<td class="td-transp">${row.transportadora||'—'}</td>`+
      `<td class="td-sm">${paletTag}</td>`+
      `<td class="td-time">${row.grade_carregamento||'—'}</td>`+
      `<td class="td-time">${row.fim_carregamento||'—'}</td>`+
      `<td><input type="time" value="${row.hora_chegada||''}" onchange="saveField('${row.dt}','${row.data_ref}','hora_chegada',this.value)"/></td>`+
      `<td style="text-align:center;" data-dt="${row.dt}" data-ref="${row.data_ref}" data-sap="${row.n_portaria||''}" onclick="promptPortaria(this.dataset.dt,this.dataset.ref,this.dataset.sap)" style="cursor:pointer;text-align:center;">${
  row.n_portaria
    ? '<span title="SAP: '+row.n_portaria+'" style="color:#4ade80;font-size:15px;">✅</span><div style="font-size:9px;color:#4ade80;letter-spacing:.5px;">'+row.n_portaria+'</div>'
    : (shouldSapNow(row)?'<span title="Status CARREGANDO sem SAP" style="color:#f97316;font-size:15px;">⚠️</span>':(clockEmoji(row)||'<span style="color:#334155;font-size:11px;">—</span>'))+'<div style="font-size:9px;color:'+(shouldSapNow(row)?'#f97316':'#475569')+';margin-top:1px;">SAP</div>'
}</td>`+
      `<td><div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;"><select class="ss" style="background:${c}22;border:1.5px solid ${c}99;color:${c}" onchange="saveStatus('${row.dt}','${row.data_ref}',this.value)">`+
        STATUS_OPTIONS.map(s=>`<option value="${s}"${s===row.status?' selected':''}>${s}</option>`).join('')+
      `</select>${opTag?opTag:''}</div></td>`+
      `<td class="td-sm">${row.descricao_documento||'—'}</td>`+
      `<td class="td-sm">${fmtTon(row.peso_liquido||row.toneladas||0)}</td>`;
    tbody.appendChild(tr);
  });

  if(!rows.length){
    const tr=document.createElement('tr');
    tr.innerHTML=`<td colspan="11" style="text-align:center;padding:30px;color:#334155;font-size:13px;">
      ${currentTableTab==='finalizadas'?'Nenhuma DT finalizada ainda.':'Todas as DTs do dia estão finalizadas.'}
    </td>`;
    tbody.appendChild(tr);
  }
}

/* ═══════════════════════════════════════════════════════
   N° PORTARIA (SAP)
═══════════════════════════════════════════════════════ */
async function promptPortaria(dt,dataRef,current){
  const val=prompt('Informe o N° SAP/Portaria para DT '+dt+':', current||'');
  if(val===null) return; // cancelado
  const sap=val.trim().replace(/[^\d]/g,'');
  await saveField(dt,dataRef,'n_portaria',sap);
  const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef);
  if(row){ row.n_portaria=sap; renderRows(); }
}

/* ═══════════════════════════════════════════════════════
   SALVAR CAMPO GENÉRICO
═══════════════════════════════════════════════════════ */
async function togglePreFat(dt,dataRef,checked){
  const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef);
  const novoTipo=checked?'PRÉ-FAT':(tipoOpMap[dt]||'VENDA ARUJA');
  await saveField(dt,dataRef,'tipo_operacao',novoTipo);
  if(row){row.tipo_operacao=novoTipo; renderRows(); renderReporte();}
}

async function toggleDocaNullReporte(dt,dataRef,checked){
  const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef);
  const docaNull=checked?false:true;
  await saveField(dt,dataRef,'doca_null',docaNull);
  if(row){row.doca_null=docaNull; renderRows(); renderReporte(); renderDocaNullAudit();}
}

async function saveField(dt,dataRef,field,value){
  spin(true);
  try{
    const patchValue=typeof value==='boolean'?value:String(value);
    await sbPatch('reporte_carga',{[field]:patchValue,updated_at:new Date().toISOString()},{dt,data_ref:dataRef});
    const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef);
    if(row){row[field]=patchValue;}
    await reloadTable();
    renderRows();
    renderReporte();
  }catch(e){showErr('Erro ao salvar: '+e.message);}
  spin(false);
}

async function tryInsertLog(row){
  const eventType=String(row.event_type||row.status||'STATUS_CHANGED');
  const payload={
    ...(row.payload&&typeof row.payload==='object'?row.payload:{}),
    legacy_status:row.status||row.status_after||null,
  };
  const dtLog={
    dt:String(row.dt||''),
    data_ref:String(row.data_ref||''),
    event_type:eventType,
    status_after:row.status_after||row.status||null,
    transportadora:row.transportadora||null,
    hora_evento:row.hora_evento||row.created_at||new Date().toISOString(),
    source:row.source||'DASHBOARD',
    payload,
  };
  try{
    await sbInsert('dt_logs',[dtLog]);
  }catch(e){
    const msg=String(e&&e.message||'');
    if(isMissingDtLogsError(e)){
      try{
        await sbInsert('reporte_logs',[{
          dt:dtLog.dt,
          data_ref:dtLog.data_ref,
          status:dtLog.status_after||eventType,
          transportadora:dtLog.transportadora||'',
          created_at:dtLog.hora_evento,
        }]);
      }catch(e2){
        if(isMissingReporteLogsError(e2)) return;
        console.warn('Falha ao gravar log legado:',String(e2&&e2.message||''));
      }
      return;
    }
    console.warn('Falha ao gravar log:',msg);
  }
}

/* SALVAR STATUS + gravar no log */
async function saveStatus(dt,dataRef,value){
  spin(true);
  try{
    await sbPatch('reporte_carga',{status:String(value),updated_at:new Date().toISOString()},{dt,data_ref:dataRef});
    const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef);
    if(row){row.status=value;}
    await reloadTable();
    renderRows();
    renderReporte();
    // Grava log só para status relevantes (evita poluição)
    const STATUS_LOG=[];
    if(STATUS_LOG.includes(value)){
      await tryInsertLog({
        dt:String(dt),
        data_ref:String(dataRef),
        status:String(value),
        transportadora:row?String(row.transportadora||''):'',
        created_at:new Date().toISOString(),
      });
    }
  }catch(e){showErr('Erro ao salvar status: '+e.message);}
  spin(false);
}

/* ═══════════════════════════════════════════════════════
   LOGS
═══════════════════════════════════════════════════════ */
function openLogs(){
  document.getElementById('log-overlay').classList.add('open');
  loadLogs();
}
function closeLogs(){
  document.getElementById('log-overlay').classList.remove('open');
}
function isMissingDtLogsError(err){
  const msg=String(err&&err.message||err||'').toLowerCase();
  return msg.includes("public.dt_logs")
    || msg.includes('relation "dt_logs" does not exist')
    || msg.includes('404')
    || msg.includes('pgrst');
}
function isMissingReporteLogsError(err){
  const msg=String(err&&err.message||err||'').toLowerCase();
  return msg.includes("public.reporte_logs")
    || msg.includes('relation "reporte_logs" does not exist')
    || msg.includes('404')
    || msg.includes('pgrst');
}
function isMissingLogsError(err){
  return isMissingDtLogsError(err)||isMissingReporteLogsError(err);
}

async function loadLogs(searchDT=""){
  const list=document.getElementById('log-list');
  list.innerHTML='<div style="color:#64748b;padding:12px;">Carregando…</div>';
  try{
    let logs=[];
    let legacy=false;
    try{
      logs=await sbGet('dt_logs','select=hora_evento,dt,event_type,status_before,status_after,transportadora,grade_carregamento,fim_carregamento,source,payload&order=hora_evento.desc&limit=200');
    }catch(e){
      if(!isMissingDtLogsError(e)) throw e;
      legacy=true;
      logs=await sbGet('reporte_logs','order=created_at.desc&limit=200');
    }
    let rows=logs||[];
    const filtro=String(searchDT||'').replace(/\D/g,'');
    if(filtro) rows=rows.filter(l=>String(l.dt||'').includes(filtro));
    if(!rows.length){list.innerHTML='<div style="color:#334155;padding:12px;">Nenhum log para o filtro informado.</div>';return;}
    list.innerHTML='';
    rows.forEach(l=>{
      const status=String(l.status_after||l.status||'').trim();
      const eventType=String(l.event_type||status||'-');
      const sc=SC[status]||'#64748b';
      const when=l.hora_evento||l.created_at;
      const grade=[l.grade_carregamento,l.fim_carregamento].filter(Boolean).join(' - ');
      const div=document.createElement('div');
      div.className='log-entry';
      div.innerHTML=
        `<span style="color:#64748b">${when?new Date(when).toLocaleString('pt-BR'):'-'}</span>`+
        `<span style="color:#93c5fd;font-weight:700">${l.dt||'-'}</span>`+
        `<span style="color:#c4b5fd;font-weight:700">${eventType}</span>`+
        `<span style="background:${sc}22;border:1px solid ${sc}55;color:${sc};border-radius:4px;padding:2px 8px;font-weight:700;">${status||'-'}</span>`+
        `<span style="color:#94a3b8">${l.transportadora||'—'}</span>`+
        `<span style="color:#94a3b8">${grade||'—'}</span>`+
        `<span style="color:#64748b">${legacy?'legado':(l.source||'SYSTEM')}</span>`;
      list.appendChild(div);
    });
  }catch(e){
    const msg=String(e&&e.message||'');
    if(isMissingLogsError(e)){
      list.innerHTML='<div style="color:#94a3b8;padding:12px;">ℹ️ Log desativado: tabela <b>reporte_logs</b> não existe neste projeto Supabase. O painel de carga continua funcionando normalmente.</div>';
      return;
    }
    list.innerHTML='<div style="color:#ef4444;padding:12px;">Erro ao carregar logs: '+e.message+'<br/><small style="color:#64748b">Verifique se a tabela reporte_logs existe no Supabase.</small></div>';
  }
}


async function renderMudancasGrade(){
  const body=document.getElementById('mudancas-list');
  const resumo=document.getElementById('mudancas-resumo');
  const diaSel=document.getElementById('mudancas-dia');
  if(!body) return;
  const ref=diaSel&&diaSel.value?diaSel.value:dKey(today());
  if(diaSel && !diaSel.value) diaSel.value=ref;
  body.innerHTML='<div style="color:#64748b;padding:12px;">Carregando mudanças…</div>';
  if(resumo) resumo.textContent='';
  try{
    let logs=[],legacy=false;
    try{
      logs=await sbGet('dt_logs',`data_ref=${sbEq(ref)}&select=hora_evento,dt,event_type,status_before,status_after,transportadora,grade_carregamento,fim_carregamento,source,payload&order=hora_evento.desc&limit=500`);
    }catch(e){
      if(!isMissingDtLogsError(e)) throw e;
      legacy=true;
      logs=await sbGet('reporte_logs',`data_ref=${sbEq(ref)}&order=created_at.desc&limit=500`);
    }
    const rows=(logs||[]).filter(l=>{
      const ev=String(l.event_type||l.status||l.status_after||'').toUpperCase();
      const st=String(l.status_after||l.status||'').toUpperCase();
      return ev.includes('REAGEND') || st.includes('REAGEND') || ev.includes('EXCLUID') || st.includes('EXCLUID') || ev.includes('REMOVED_FROM_CURRENT');
    });
    const reag=rows.filter(l=>String(l.event_type||l.status||l.status_after||'').toUpperCase().includes('REAGEND') || String(l.status_after||l.status||'').toUpperCase().includes('REAGEND')).length;
    const excl=rows.length-reag;
    if(resumo) resumo.innerHTML=`<span style="color:#f59e0b">${reag} reagendada(s)</span> · <span style="color:#ef4444">${excl} excluída(s)/fora da grade</span> · ${ref}`;
    if(!rows.length){
      body.innerHTML='<div style="color:#334155;padding:18px;text-align:center;">Nenhuma DT reagendada ou excluída da grade neste dia.</div>';
      return;
    }
    body.innerHTML='';
    rows.forEach(l=>{
      const ev=String(l.event_type||l.status||l.status_after||'-');
      const st=String(l.status_after||l.status||'-');
      const isReag=ev.toUpperCase().includes('REAGEND')||st.toUpperCase().includes('REAGEND');
      const payload=l.payload&&typeof l.payload==='object'?l.payload:{};
      const grade=[l.grade_carregamento||payload.grade_carregamento,l.fim_carregamento||payload.fim_carregamento].filter(Boolean).join(' → ');
      const div=document.createElement('div');
      div.className='mud-item';
      div.innerHTML=`
        <div style="font-size:11px;color:#64748b;">${new Date(l.hora_evento||l.created_at||Date.now()).toLocaleString('pt-BR')}</div>
        <div style="font-weight:900;color:#93c5fd;">${escHtml(l.dt||'-')}</div>
        <div><span class="mud-badge" style="border-color:${isReag?'#f59e0b55':'#ef444455'};color:${isReag?'#fbbf24':'#fca5a5'};background:${isReag?'#f59e0b18':'#ef444418'};">${isReag?'REAGENDADA':'EXCLUÍDA DA GRADE'}</span></div>
        <div style="color:#94a3b8;">${escHtml(l.transportadora||'—')}</div>
        <div style="color:#64748b;">${escHtml(grade || (legacy?'legado':'—'))}</div>`;
      body.appendChild(div);
    });
  }catch(e){
    body.innerHTML='<div style="color:#ef4444;padding:12px;">Erro ao carregar mudanças: '+escHtml(e.message)+'</div>';
  }
}

function mudancasSetDia(v){
  const sel=document.getElementById('mudancas-dia');
  if(sel) sel.value=v;
  renderMudancasGrade();
}

/* ═══════════════════════════════════════════════════════
   REAGENDAMENTO
═══════════════════════════════════════════════════════ */
function openReagend(dt,dataRef){
  reagendDT={dt,dataRef};
  const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef)||{};
  document.getElementById('reagend-dt-num').textContent=dt;
  document.getElementById('reagend-obs').value='';
  // Pré-preenche com grade atual
  if(row.grade_carregamento){
    const g=parseBR(row.grade_carregamento);
    if(g) document.getElementById('reagend-grade').value=g.toISOString().slice(0,16);
  }
  if(row.fim_carregamento){
    const f=parseBR(row.fim_carregamento);
    if(f) document.getElementById('reagend-fim').value=f.toISOString().slice(0,16);
  }
  document.getElementById('reagend-overlay').classList.add('open');
}
function closeReagend(){
  document.getElementById('reagend-overlay').classList.remove('open');
  reagendDT=null;
}
async function confirmarReagendamento(){
  if(!reagendDT)return;
  const {dt,dataRef}=reagendDT;
  const gradeVal=document.getElementById('reagend-grade').value;
  const fimVal=document.getElementById('reagend-fim').value;
  const obs=document.getElementById('reagend-obs').value.trim();
  if(!gradeVal){showErr('Informe a nova grade.');return;}
  // Converte datetime-local para formato do sistema
  const fmtLocal=v=>{
    if(!v)return '';
    const d=new Date(v);
    return d.toLocaleString('pt-BR',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
  };
  spin(true);
  try{
    await sbPatch('reporte_carga',{
      grade_carregamento:fmtLocal(gradeVal),
      fim_carregamento:fimVal?fmtLocal(fimVal):'',
      status:'AG CHEGADA',
      reagendada:true,
      updated_at:new Date().toISOString(),
    },{dt,data_ref:dataRef});
    // Log do reagendamento
    await tryInsertLog({
      dt:String(dt),
      data_ref:String(dataRef),
      status:'REAGENDADA',
      transportadora:obs||'Reagendamento via sistema',
      created_at:new Date().toISOString(),
    });
    const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef);
    if(row){
      row.grade_carregamento=fmtLocal(gradeVal);
      row.fim_carregamento=fimVal?fmtLocal(fimVal):'';
      row.status='AG CHEGADA';
      row.reagendada=true;
    }
    showOk('DT '+dt+' reagendada! Aparece em destaque laranja.');
    closeReagend();
    renderRows();
  }catch(e){showErr('Erro ao reagendar: '+e.message);}
  spin(false);
}

/* ═══════════════════════════════════════════════════════
   PAINEL — MAPA DE SEPARAÇÃO
═══════════════════════════════════════════════════════ */
async function openPanel(dt,dataRef){
  panelDT={dt,dataRef};
  const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef)||{};
  document.getElementById('ptitle').textContent='📦  DT '+dt;
  const opBadge=buildTipoOperacaoBadge(row.tipo_operacao,{small:true,emptyDash:false});
  document.getElementById('pdt-info').innerHTML=
    `<b style="color:#93c5fd;font-size:15px">${dt}</b>  ${opBadge}<br/>`+
    `<span style="color:#64748b">Transportadora:</span> <b style="color:#e2e8f0">${row.transportadora||'—'}</b><br/>`+
    `<span style="color:#64748b">Grade:</span> <b style="color:#a5b4fc">${row.grade_carregamento||'—'}</b>  →  <b style="color:#a5b4fc">${row.fim_carregamento||'—'}</b>`;

  let listaMats=[];
  try{
    const saved=await sbGet('reporte_materiais',`dt=eq.${encodeURIComponent(dt)}&data_ref=eq.${encodeURIComponent(dataRef)}&order=ordem.asc`);
    listaMats=saved||[];
  }catch(e){}

  if(!listaMats.length){
    const dtKey=String(dt);
    const mapa=exportMap[dtKey]||[];
    listaMats=mapa.map((m,i)=>({material:String(m.material),quantidade:String(m.quantidade),ordem:i}));
  }

  document.getElementById('pobs').value='';
  renderMatList(listaMats);
  document.getElementById('panel-overlay').style.display='block';
  document.getElementById('panel').classList.add('open');
}

function closePanel(){
  document.getElementById('panel-overlay').style.display='none';
  document.getElementById('panel').classList.remove('open');
  panelDT=null;
}

function renderMatList(mats){
  const list=document.getElementById('mat-list');list.innerHTML='';
  (mats||[]).forEach(m=>{
    const row=document.createElement('div');row.className='mat-row';
    row.innerHTML=
      `<input class="mat-input" placeholder="Material" value="${m.material||''}" data-f="material" oninput="updateMatRowPreview(this.closest('.mat-row'));updatePanelTotal()"/>`+
      `<div class="mat-qty-cell"><input class="mat-input" placeholder="Fardos" value="${m.quantidade||''}" data-f="quantidade" oninput="updateMatRowPreview(this.closest('.mat-row'));updatePanelTotal()"/><div class="mat-palete-preview"></div></div>`+
      `<button class="mat-del" onclick="this.parentElement.remove();updatePanelTotal();">🗑</button>`;
    list.appendChild(row);
    updateMatRowPreview(row);
  });
  updatePanelTotal();
}

function addMatRow(){
  const row=document.createElement('div');row.className='mat-row';
  row.innerHTML=
    `<input class="mat-input" placeholder="Material" data-f="material" oninput="updateMatRowPreview(this.closest('.mat-row'));updatePanelTotal()"/>`+
    `<div class="mat-qty-cell"><input class="mat-input" placeholder="Fardos" data-f="quantidade" oninput="updateMatRowPreview(this.closest('.mat-row'));updatePanelTotal()"/><div class="mat-palete-preview"></div></div>`+
    `<button class="mat-del" onclick="this.parentElement.remove();updatePanelTotal();">🗑</button>`;
  document.getElementById('mat-list').appendChild(row);
  row.querySelector('input').focus();
  updateMatRowPreview(row);
  updatePanelTotal();
}

function updateMatRowPreview(row){
  if(!row) return;
  const material=row.querySelector('[data-f=material]')?.value||'';
  const quantidade=row.querySelector('[data-f=quantidade]')?.value||'';
  const preview=row.querySelector('.mat-palete-preview');
  if(!preview) return;
  const pal=formatPaletes(material,quantidade);
  const fardos=parseQuantidadeFardos(quantidade);
  preview.textContent=pal.porPalete
    ? `${pal.label} · ${formatNumeroBR(fardos)} fardo(s) no pedido · ${pal.porPalete}/palete`
    : pal.label;
}

function updatePanelTotal(){
  const rows=[...document.querySelectorAll('#mat-list .mat-row')];
  const total=rows.reduce((s,row)=>{
    const material=row.querySelector('[data-f=material]')?.value||'';
    const quantidade=row.querySelector('[data-f=quantidade]')?.value||'';
    return s+formatPaletes(material,quantidade).paletes;
  },0);
  const sobras=rows.reduce((s,row)=>{
    const material=row.querySelector('[data-f=material]')?.value||'';
    const quantidade=row.querySelector('[data-f=quantidade]')?.value||'';
    return s+formatPaletes(material,quantidade).sobra;
  },0);
  const el=document.getElementById('mat-total');
  if(el) el.textContent=`Total: ${total} palete(s) necessário(s) + ${formatNumeroBR(sobras)} fardo(s) sobra`;
}

function printPanelMapa(){
  if(!panelDT) return;
  // Build a clean print-friendly table from current mat-list
  const rows=[...document.querySelectorAll('#mat-list .mat-row')];
  const mats=rows.map(r=>({
    mat:r.querySelector('[data-f=material]').value.trim(),
    qtd:r.querySelector('[data-f=quantidade]').value.trim()
  })).filter(m=>m.mat);
  const row=tableData.find(r=>r.dt===panelDT.dt&&r.data_ref===panelDT.dataRef)||{};
  const total=mats.reduce((s,m)=>s+formatPaletes(m.mat,m.qtd).paletes,0);
  const sobras=mats.reduce((s,m)=>s+formatPaletes(m.mat,m.qtd).sobra,0);
  const printArea=document.getElementById('print-panel-area');
  // Temporarily swap content with print table
  const origInner=printArea.innerHTML;
  printArea.innerHTML=`
    <div class="print-header">
      <div>
        <h2>📦 Mapa de Separação — DT ${panelDT.dt}</h2>
        <div style="font-size:11px;color:#475569;margin-top:4px;">
          ${row.transportadora||''} &nbsp;|&nbsp; Grade: ${row.grade_carregamento||'—'} → ${row.fim_carregamento||'—'}
          &nbsp;|&nbsp; ${row.tipo_operacao||''}
        </div>
      </div>
      <div class="print-meta">
        Impresso em ${new Date().toLocaleString('pt-BR')}<br/>
        Total: <b>${total} palete(s) necessário(s) + ${formatNumeroBR(sobras)} fardo(s) sobra</b>
      </div>
    </div>
    <table class="print-mat-table">
      <thead><tr><th>#</th><th>MATERIAL</th><th>QUANTIDADE</th></tr></thead>
      <tbody>${mats.map((m,i)=>`<tr><td>${i+1}</td><td>${m.mat}</td><td style="text-align:right;">${formatPaletes(m.mat,m.qtd).label}</td></tr>`).join('')}</tbody>
      <tfoot><tr><td colspan="2" style="text-align:right;font-weight:800;padding:6px 10px;">TOTAL</td><td style="text-align:right;font-weight:800;padding:6px 10px;">${total} palete(s) necessário(s) + ${formatNumeroBR(sobras)} fardo(s) sobra</td></tr></tfoot>
    </table>
    ${document.getElementById('pobs').value ? `<div style="margin-top:14px;font-size:11px;color:#475569;">Obs: ${document.getElementById('pobs').value}</div>` : ''}
  `;
  document.body.classList.add('print-panel-mode');
  window.print();
  setTimeout(()=>{
    document.body.classList.remove('print-panel-mode');
    printArea.innerHTML=origInner;
    // Re-render mat list after restore
    openPanel(panelDT.dt, panelDT.dataRef);
  },600);
}

async function saveMateriais(){
  if(!panelDT)return;
  const {dt,dataRef}=panelDT;
  const row=tableData.find(r=>String(r.dt)===String(dt)&&String(r.data_ref)===String(dataRef));
  const paletizacao=normalizePaletizacaoLabel(row&&row.paletizacao);
  const rows=[...document.getElementById('mat-list').querySelectorAll('.mat-row')];
  const mats=rows.map((r,i)=>({
    dt,data_ref:dataRef,
    material:r.querySelector('[data-f=material]').value.trim(),
    quantidade:r.querySelector('[data-f=quantidade]').value.trim(),
    observacao:document.getElementById('pobs').value.trim(),
    ordem:i,
    paletizacao
  })).filter(m=>m.material);
  spin(true);
  try{
    await sbDelete('reporte_materiais',{dt,data_ref:dataRef});
    if(mats.length)await sbInsert('reporte_materiais',mats);
    showOk('Materiais da DT '+dt+' salvos!');
    closePanel();
  }catch(e){showErr('Erro ao salvar materiais: '+e.message);}
  spin(false);
}

/* ═══════════════════════════════════════════════════════
   CARREGAR AO ABRIR
═══════════════════════════════════════════════════════ */
async function tryLoadExisting(){
  try{
    const refs=await getActiveRefs();
    if(!refs.length){tableData=[];renderRows();return;}
    const data=await sbGet('reporte_carga',
      `data_ref=${sbIn(refs)}&order=dia_ref.asc,agenda.asc`
    );
    if(data&&data.length){
      tableData=normalizeDiaRefRows(data).sort((a,b)=>{
        const ga=parseBR(a.grade_carregamento)||new Date(9999,0);
        const gb=parseBR(b.grade_carregamento)||new Date(9999,0);
        return ga-gb;
      });
      await enrichRowsWithPaletizacao(tableData);
      await persistNormalizedDiaRefs(tableData);
      // Popula tipoOpMap a partir do banco para o painel funcionar sem ZLES002
      tableData.forEach(r=>{ if(r.dt&&r.tipo_operacao) tipoOpMap[r.dt]=r.tipo_operacao; });
      if(currentTableTab==='reporte') renderReporte(); else renderRows();
      setStep(4);
    }
  }catch(e){ throw e; /* erro de conexão — propaga para o badge */ }
}

/* ═══════════════════════════════════════════════════════
   DRAG & DROP
═══════════════════════════════════════════════════════ */
['dz1','dz2'].forEach(id=>{
  const el=document.getElementById(id);if(!el)return;
  el.addEventListener('dragover',e=>{e.preventDefault();el.classList.add('drag');});
  el.addEventListener('dragleave',()=>el.classList.remove('drag'));
  el.addEventListener('drop',e=>{
    e.preventDefault();el.classList.remove('drag');
    const f=e.dataTransfer.files[0];if(!f)return;
    if(id==='dz1')processAgend(f);
    else processRelatorioCSV(f);
  });
});

/* ═══════════════════════════════════════════════════════
   EXPORT CSV
═══════════════════════════════════════════════════════ */
function csvGradeValue(v){
  return String(v||'').replace(/,\s*/,' ');
}

function exportCSV(){
  const cols=['DIA','DT','TRANSPORTADORA','GRADE','FIM','HORA CHEGADA','N° PORTARIA','STATUS','DESC. DOCUMENTO','PESO LÍQUIDO','TIPO OPERAÇÃO','MATERIAL','PALETES','SOBRA (FARDOS)','QTD TOTAL (FARDOS)'];
  const rows=[];
  const exportRows=dedupeCargaRowsByDTRef((tableData||[])
    .filter(isRowInActiveReportWindow)
    .filter(rpRowContaNoReporte)
  ).sort(compareAgendaRows);
  exportRows.forEach(r=>{
    const dtKey=String(r.dt||'').trim();
    const mats=exportMap[dtKey]||[];
    const base=[r.dia_ref,r.dt,r.transportadora,csvGradeValue(r.grade_carregamento),csvGradeValue(r.fim_carregamento),r.hora_chegada,r.n_portaria||'',r.status,r.descricao_documento||'',fmtTon(r.peso_liquido||r.toneladas||0),r.tipo_operacao||''];
    if(!mats.length){
      rows.push([...base,'','','','']);
    } else {
      mats.forEach((m,i)=>{
        const pal=formatPaletes(m.material,m.quantidade);
        const qtdFardos=parseQuantidadeFardos(m.quantidade);
        if(i===0){
          rows.push([...base,m.material,pal.paletes,pal.sobra,Math.round(qtdFardos)]);
        } else {
          rows.push(['','','','','','','','','','','',m.material,pal.paletes,pal.sobra,Math.round(qtdFardos)]);
        }
      });
    }
  });
  const csv=[cols,...rows].map(r=>r.map(v=>'"'+String(v||'').replace(/"/g,'""')+'"').join(';')).join('\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));
  a.download='reporte_'+new Date().toLocaleDateString('pt-BR').replace(/\//g,'-')+'.csv';
  a.click();
}

/* ═══════════════════════════════════════════════════════
   REPORTE DE STATUS
═══════════════════════════════════════════════════════ */
const RP_TIPOS = ['GRADE','TRANSFERÊNCIA','VENDA ARUJA','VENDA MOGI','PRÉ-FATURA'];
const RP_SUZANO_TIPOS = ['VENDA ARUJA','VENDA MOGI','TRANSFERENCIA','PRE-FATURA'];
const RP_COLORS = {
  'VENDA ARUJA':'#22c55e','VENDA MOGI':'#06b6d4','PRÉ-FATURA':'#a78bfa','TRANSFERÊNCIA':'#fb923c','GRADE':'#60a5fa'
};
let rpTurnoFiltro = 'todos';
let rpMetas = JSON.parse(localStorage.getItem('rp_metas')||'{}');
let rpHoraCorte = localStorage.getItem('rp_hora_corte') || '';
let rpDataRef = localStorage.getItem('rp_data_ref') || '';
let rpPlanejadoSuzano = JSON.parse(localStorage.getItem('rp_planejado_suzano')||'{}');
let rpLastReport = null;
let rpSnapshotCache = {};
let appConfig = JSON.parse(localStorage.getItem('reporte_app_config')||'{}');

const STATUS_REALIZADO = ['EM FATURAMENTO','EXPEDIDO'];
const STATUS_FORA_REPORTE = ['NO SHOW','VEICULO RECUSADO'];

function rpRefDate(row){
  // Usado apenas para classificação de turno e corte de horário.
  // NÃO é mais usado como chave de dia (rpDateKey) — evita que DTs com
  // fim_carregamento no dia seguinte desapareçam do reporte quando a
  // agenda anterior é reenviada e essas DTs são excluídas.
  return parseBR(row.fim_carregamento||'') || parseBR(row.fim_agenda||'');
}

function rpDateKey(row){
  return rowReportRefKey(row);
}

function compareDateRefs(a,b){
  const da=parseBR(a), db=parseBR(b);
  return (da?da.getTime():0)-(db?db.getTime():0);
}

function isTransferencia(row){
  const d=String(row.descricao_documento||'').toUpperCase();
  return d.includes('TRANSFER') || d.includes('TNF') || d.includes('FILIAL') || d.includes('INTERCOMP') || d.includes('REMESSA');
}

function isVendaNormal(row){
  const d=String(row.descricao_documento||'').toUpperCase();
  if(isTransferencia(row)) return false;
  return !d || d.includes('VENDA NORMAL') || d.includes('VENDA');
}

function isVendaAruja(row){
  const centro=String(row.centro||'').trim();
  const op=String(row.tipo_operacao||'').toUpperCase();
  if(!isVendaNormal(row) || op.includes('TRANSFER')) return false;
  return centro==='1111' || (!centro && op.includes('ARUJA'));
}

function isVendaMogi(row){
  const centro=String(row.centro||'').trim();
  const op=String(row.tipo_operacao||'').toUpperCase();
  if(!isVendaNormal(row) || op.includes('TRANSFER')) return false;
  return centro==='1110' || (!centro && op.includes('MOGI'));
}

function rpModalidade(row){
  const op=(row.tipo_operacao||'').trim().toUpperCase();
  if(op==='PRÉ-FAT'||op==='PRÉ-FATURA'||op.includes('PRÉ')||op.includes('PRE')) return 'PRÉ-FATURA';
  if(isTransferencia(row)) return 'TRANSFERÊNCIA';
  if(isVendaMogi(row)) return 'VENDA MOGI';
  if(isVendaAruja(row)) return 'VENDA ARUJA';
  if(op.includes('TRANSFER')) return 'TRANSFERÊNCIA';
  if(op.includes('MOGI')) return 'VENDA MOGI';
  if(op.includes('ARUJA')) return 'VENDA ARUJA';
  return '';
}

function rpGetTurno(row){
  // Classifica pelo horário do FIM da agenda
  const grade = rpRefDate(row);
  if(!grade) return null;
  const h = grade.getHours();
  if(h>=7 && h<=14) return 'T1';
  if(h>=15 && h<=22) return 'T2';
  return 'T3'; // 23-06
}

function rpNormalizeTipo(row){
  return rpModalidade(row) || 'VENDA ARUJA';
}

function rpSetTurno(t){
  rpTurnoFiltro=t;
  document.querySelectorAll('.rp-turno').forEach(btn=>{
    btn.classList.remove('active-turno');
  });
  const id=t==='todos'?'rp-t0':t==='T1'?'rp-t1':t==='T2'?'rp-t2':'rp-t3';
  const el=document.getElementById(id);
  if(el) el.classList.add('active-turno');
  renderReporte();
}

function rpSaveMeta(tipo,val){
  rpMetas[tipo]=parseInt(val)||0;
  try{localStorage.setItem('rp_metas',JSON.stringify(rpMetas));}catch(e){}
  renderReporte();
}

function rpSetDataRef(v){
  rpDataRef=String(v||'').trim();
  try{localStorage.setItem('rp_data_ref',rpDataRef);}catch(e){}
  renderReporte();
  renderPlanejamento();
}

function rpSetHoraCorte(v){
  rpHoraCorte=String(v||'').trim();
  try{localStorage.setItem('rp_hora_corte',rpHoraCorte);}catch(e){}
  renderReporte();
}

function rpParsePlanejadoInput(raw){
  const n=parseCargaNumber(raw);
  if(!n) return 0;
  return n<1000?Math.round(n*1000):Math.floor(n);
}

function rpGetPlanejadoSuzano(){
  const item=rpPlanejadoSuzano[rpDataRef];
  if(item&&typeof item==='object') return Object.values(item).reduce((s,v)=>s+Math.max(0,Number(v)||0),0);
  return Math.max(0,Number(item||0));
}

function rpGetPlanejadoSuzanoDetalhado(){
  const item=rpPlanejadoSuzano[rpDataRef];
  const det={};
  RP_SUZANO_TIPOS.forEach(tipo=>det[tipo]=0);
  if(item&&typeof item==='object'){
    RP_SUZANO_TIPOS.forEach(tipo=>{det[tipo]=Math.max(0,Number(item[tipo]||0));});
  }else if(Number(item||0)>0){
    det['VENDA ARUJA']=Number(item)||0;
  }
  return det;
}

function rpSetPlanejadoSuzano(v){
  const val=rpParsePlanejadoInput(v);
  if(!rpDataRef) rpDataRef=dKey(today());
  if(val) rpPlanejadoSuzano[rpDataRef]={'VENDA ARUJA':val,'VENDA MOGI':0,'TRANSFERENCIA':0,'PRE-FATURA':0};
  else delete rpPlanejadoSuzano[rpDataRef];
  try{localStorage.setItem('rp_planejado_suzano',JSON.stringify(rpPlanejadoSuzano));}catch(e){}
  renderReporte();
}

function rpSetPlanejadoSuzanoTipo(tipo,v){
  if(!rpDataRef) rpDataRef=dKey(today());
  const det=rpGetPlanejadoSuzanoDetalhado();
  const val=rpParsePlanejadoInput(v);
  if(val) det[tipo]=val;
  else det[tipo]=0;
  rpPlanejadoSuzano[rpDataRef]=det;
  if(!rpGetPlanejadoSuzano()) delete rpPlanejadoSuzano[rpDataRef];
  try{localStorage.setItem('rp_planejado_suzano',JSON.stringify(rpPlanejadoSuzano));}catch(e){}
  renderReporte();
  renderPlanejamento();
}

function getEmailList(raw){
  return String(raw||'').split(/[;,]/).map(s=>s.trim()).filter(Boolean);
}

function openConfigModal(){
  const ov=document.getElementById('config-overlay');
  if(!ov) return;
  appConfig=JSON.parse(localStorage.getItem('reporte_app_config')||'{}');
  const set=(id,val)=>{const el=document.getElementById(id); if(el) el.value=val||'';};
  set('cfg-email-provider',appConfig.emailProvider||'');
  set('cfg-email-from',appConfig.emailFrom||'');
  set('cfg-email-endpoint',appConfig.emailEndpoint||'');
  set('cfg-email-to',(appConfig.emailTo||[]).join('; '));
  set('cfg-email-cc',(appConfig.emailCc||[]).join('; '));
  set('cfg-report-inicial',appConfig.reportInicial||'00:00');
  set('cfg-report-final',appConfig.reportFinal||'23:59');
  ov.style.display='flex';
}

function closeConfigModal(){
  const ov=document.getElementById('config-overlay');
  if(ov) ov.style.display='none';
}

function saveConfigModal(){
  const val=id=>document.getElementById(id)?.value || '';
  const provider=val('cfg-email-provider');
  appConfig={
    emailProvider:provider,
    emailFrom:val('cfg-email-from').trim(),
    emailEndpoint:val('cfg-email-endpoint').trim() || (provider==='brevo'?'http://localhost:8787/send-report':''),
    emailTo:getEmailList(val('cfg-email-to')),
    emailCc:getEmailList(val('cfg-email-cc')),
    reportInicial:val('cfg-report-inicial') || '00:00',
    reportFinal:val('cfg-report-final') || '23:59',
    updatedAt:new Date().toISOString(),
  };
  try{localStorage.setItem('reporte_app_config',JSON.stringify(appConfig));}catch(e){}
  closeConfigModal();
  showOk('Configurações do reporte salvas.');
}

function syncConfigFromOpenModal(){
  const ov=document.getElementById('config-overlay');
  if(!ov || ov.style.display==='none') return;
  const val=id=>document.getElementById(id)?.value || '';
  const provider=val('cfg-email-provider');
  appConfig={
    emailProvider:provider,
    emailFrom:val('cfg-email-from').trim(),
    emailEndpoint:val('cfg-email-endpoint').trim() || (provider==='brevo'?'http://localhost:8787/send-report':''),
    emailTo:getEmailList(val('cfg-email-to')),
    emailCc:getEmailList(val('cfg-email-cc')),
    reportInicial:val('cfg-report-inicial') || '00:00',
    reportFinal:val('cfg-report-final') || '23:59',
    updatedAt:new Date().toISOString(),
  };
  try{localStorage.setItem('reporte_app_config',JSON.stringify(appConfig));}catch(e){}
}

function rpBuildSnapshotPayload(source='DASHBOARD',evento='Snapshot manual'){
  renderReporte();
  const report=rpLastReport||{};
  const status=report.statusCounts||{};
  const dataRef=report.dataRef||rpDataRef||dKey(today());
  const dataOperacao=parseBR(dataRef);
  return {
    data_ref:dataRef,
    data_operacao:dataOperacao?dataOperacao.toISOString().slice(0,10):null,
    snapshot_at:new Date().toISOString(),
    source,
    evento,
    planejado_suzano_kg:report.planejadoSuzano||rpGetPlanejadoSuzano()||null,
    nossa_grade_kg:report.nossaGrade||0,
    realizado_kg:report.realizadoGrade||0,
    pendente_kg:Math.max(0,(report.nossaGrade||0)-(report.realizadoGrade||0)),
    dts_total:tableData.filter(r=>rpDateKey(r)===dataRef).length,
    dts_abertas:tableData.filter(r=>rpDateKey(r)===dataRef&&!isFinalStatus(r.status)).length,
    dts_expedidas:Number(status.EXPEDIDO||0),
    dts_no_show:Number(status['NO SHOW']||0),
    dts_recusadas:Number(status['VEICULO RECUSADO']||0),
    detalhes:{
      corte:report.corteLabel||'',
      turno:rpTurnoFiltro,
      email_configurado:!!(appConfig.emailProvider&&appConfig.emailFrom&&appConfig.emailTo&&appConfig.emailTo.length),
      planejado_suzano_detalhado:rpGetPlanejadoSuzanoDetalhado(),
      tipos:report.tipos||[],
      paletizacao:report.paletizacao||{},
      turnos:report.turnos||{},
    }
  };
}

async function rpCreateGradeSnapshot(source='DASHBOARD',evento='Snapshot manual'){
  const payload=rpBuildSnapshotPayload(source,evento);
  await sbInsert('grade_snapshots',[payload]);
  await rpLoadSnapshots(payload.data_ref);
  renderTimelineChart();
  return payload;
}

async function rpSalvarSnapshotManual(){
  spin(true);
  try{
    await rpCreateGradeSnapshot('MANUAL','Snapshot manual do reporte');
    showOk('Snapshot da grade salvo no banco.');
  }catch(e){showErr('Erro ao salvar snapshot: '+e.message);}
  spin(false);
}

async function rpSalvarPlanejadoSuzano(){
  spin(true);
  try{
    const kg=rpGetPlanejadoSuzano();
    if(!kg){showErr('Informe o planejado Suzano antes de salvar.');spin(false);return;}
    const dataRef=rpDataRef||dKey(today());
    const dataOperacao=parseBR(dataRef);
    await sbInsert('planejamento_suzano_snapshots',[{
      data_ref:dataRef,
      data_operacao:dataOperacao?dataOperacao.toISOString().slice(0,10):null,
      planejado_suzano_kg:kg,
      dts_total:tableData.filter(r=>rpDateKey(r)===dataRef).length,
      source:'DASHBOARD',
      observacao:'Planejado informado no planejamento',
      detalhes:rpGetPlanejadoSuzanoDetalhado()
    }]);
    await rpCreateGradeSnapshot('SUZANO','Planejado Suzano atualizado');
    showOk('Planejado Suzano salvo e snapshot criado.');
  }catch(e){showErr('Erro ao salvar planejado Suzano: '+e.message);}
  spin(false);
}

async function rpSalvarReporteManual(){
  spin(true);
  try{
    const snap=await rpCreateGradeSnapshot('REPORTE','Reporte manual salvo');
    const report=rpLastReport||{};
    const dataRef=report.dataRef||rpDataRef||dKey(today());
    const dataOperacao=parseBR(dataRef);
    const status=report.statusCounts||{};
    await sbInsert('reportes_diarios',[{
      data_ref:dataRef,
      data_operacao:dataOperacao?dataOperacao.toISOString().slice(0,10):null,
      tipo:'MANUAL',
      hora_reporte:new Date().toISOString(),
      planejado_suzano_kg:report.planejadoSuzano||null,
      grade_atual_kg:report.nossaGrade||0,
      realizado_kg:report.realizadoGrade||0,
      pendente_kg:Math.max(0,(report.nossaGrade||0)-(report.realizadoGrade||0)),
      variacao_suzano_kg:(report.nossaGrade||0)-(report.planejadoSuzano||0),
      variacao_grade_kg:null,
      dts_expedidas:Number(status.EXPEDIDO||0),
      dts_no_show:Number(status['NO SHOW']||0),
      dts_recusadas:Number(status['VEICULO RECUSADO']||0),
      resumo:{...report,email:appConfig},
    }]);
    showOk('Reporte manual salvo no banco.');
  }catch(e){showErr('Erro ao salvar reporte: '+e.message);}
  spin(false);
}

function rpGetCorteDate(){
  const selectedDate=parseBR(rpDataRef||'') || today();
  const corte=new Date(selectedDate);
  if(!rpHoraCorte || !/^\d{2}:\d{2}$/.test(rpHoraCorte)){
    corte.setHours(23,59,59,999);
    return corte;
  }
  const [hh,mm]=rpHoraCorte.split(':').map(n=>parseInt(n,10));
  corte.setHours(hh||0,mm||0,59,999);
  return corte;
}

function rpRowDentroDoCorte(row,corteDate){
  const fim=rpRefDate(row);
  if(!fim) return false;
  return fim.getTime()<=corteDate.getTime();
}

function rpRowContaNoReporte(row){
  if(row&&row.doca_null) return false;
  return !STATUS_FORA_REPORTE.includes(String((row&&row.status)||'').trim().toUpperCase());
}

async function rpLoadSnapshots(dataRef){
  const ref=String(dataRef||rpDataRef||'').trim();
  if(!ref) return [];
  try{
    const rows=await sbGet('grade_snapshots',
      `data_ref=${sbEq(ref)}&select=snapshot_at,source,evento,planejado_suzano_kg,nossa_grade_kg,realizado_kg,pendente_kg,dts_total,dts_expedidas,dts_no_show,dts_recusadas&order=snapshot_at.asc&limit=120`
    );
    rpSnapshotCache[ref]=rows||[];
    return rpSnapshotCache[ref];
  }catch(e){
    rpSnapshotCache[ref]=rpSnapshotCache[ref]||[];
    return rpSnapshotCache[ref];
  }
}

function renderTimelineChart(){
  const el=document.getElementById('rp-timeline-chart');
  const info=document.getElementById('rp-snapshot-info');
  if(!el) return;
  const rows=(rpSnapshotCache[rpDataRef]||[]).filter(Boolean);
  if(info) info.textContent=rows.length?`${rows.length} snapshots salvos`:'Sem snapshots salvos ainda';
  if(!rows.length){
    el.innerHTML='<div style="height:100%;display:flex;align-items:center;justify-content:center;color:#64748b;font-size:12px;">Salve um snapshot para começar a linha histórica do dia.</div>';
    return;
  }
  const W=900,H=250,padL=58,padR=22,padT=22,padB=42;
  const vals=rows.flatMap(r=>[r.planejado_suzano_kg,r.nossa_grade_kg,r.realizado_kg].map(Number).filter(Number.isFinite));
  const max=Math.max(1000,...vals);
  const min=0;
  const xFor=i=>padL+(rows.length===1?(W-padL-padR)/2:i*(W-padL-padR)/(rows.length-1));
  const yFor=v=>padT+(max-(Number(v)||0))*(H-padT-padB)/(max-min);
  const line=(key,color,label)=> {
    const pts=rows.map((r,i)=>`${xFor(i)},${yFor(r[key])}`).join(' ');
    const dots=rows.map((r,i)=>`<circle cx="${xFor(i)}" cy="${yFor(r[key])}" r="4" fill="${color}"><title>${label}: ${fmtCargaCompact(r[key]||0)} - ${new Date(r.snapshot_at).toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})}</title></circle>`).join('');
    return `<polyline points="${pts}" fill="none" stroke="${color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>${dots}`;
  };
  const labels=rows.map((r,i)=>{
    const d=new Date(r.snapshot_at);
    const x=xFor(i);
    return `<text x="${x}" y="${H-17}" fill="#64748b" font-size="10" text-anchor="middle">${d.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})}</text>`;
  }).join('');
  const eventDots=rows.filter(r=>r.evento).map((r,i)=>{
    const idx=rows.indexOf(r), x=xFor(idx), y=padT+8;
    return `<circle cx="${x}" cy="${y}" r="3" fill="#f59e0b"><title>${escHtml(r.evento||'Evento')}</title></circle>`;
  }).join('');
  el.innerHTML=`<svg viewBox="0 0 ${W} ${H}" width="100%" height="100%" role="img" aria-label="Evolução do planejamento">
    <rect x="0" y="0" width="${W}" height="${H}" fill="#0f172a"/>
    <line x1="${padL}" y1="${padT}" x2="${padL}" y2="${H-padB}" stroke="#334155"/>
    <line x1="${padL}" y1="${H-padB}" x2="${W-padR}" y2="${H-padB}" stroke="#334155"/>
    <text x="12" y="${padT+4}" fill="#64748b" font-size="10">${fmtCargaCompact(max)}</text>
    <text x="12" y="${H-padB+4}" fill="#64748b" font-size="10">0 kg</text>
    ${line('planejado_suzano_kg','#38bdf8','Suzano')}
    ${line('nossa_grade_kg','#60a5fa','Nossa grade')}
    ${line('realizado_kg','#22c55e','Realizado')}
    ${eventDots}
    ${labels}
    <g transform="translate(${padL},14)" font-size="11" font-weight="700">
      <circle cx="0" cy="0" r="4" fill="#38bdf8"/><text x="8" y="4" fill="#94a3b8">Suzano</text>
      <circle cx="88" cy="0" r="4" fill="#60a5fa"/><text x="96" y="4" fill="#94a3b8">Nossa grade</text>
      <circle cx="208" cy="0" r="4" fill="#22c55e"/><text x="216" y="4" fill="#94a3b8">Realizado</text>
    </g>
  </svg>`;
}

function rpSuzanoLabel(tipo){
  return {
    'VENDA ARUJA':'Venda Aruja',
    'VENDA MOGI':'Venda Mogi',
    'TRANSFERENCIA':'Transferencia',
    'PRE-FATURA':'Pre-fatura'
  }[tipo]||tipo;
}

function renderPlanejamento(){
  const wrap=document.getElementById('planejamento-wrap');
  if(!wrap) return;
  const refs=[...new Set(tableData.map(r=>rpDateKey(r)).filter(Boolean))]
    .filter(ref=>{const d=parseBR(ref);return d&&(sameDay(d,today())||sameDay(d,tomorrow()));})
    .sort(compareDateRefs);
  if(!rpDataRef || (refs.length && !refs.includes(rpDataRef))) rpDataRef=refs[0]||dKey(today());
  const sel=document.getElementById('pl-data-ref');
  if(sel){
    sel.innerHTML=refs.map(ref=>`<option value="${ref}">${ref}</option>`).join('');
    sel.value=rpDataRef;
  }
  const det=rpGetPlanejadoSuzanoDetalhado();
  const inputs=document.getElementById('pl-suzano-inputs');
  if(inputs){
    inputs.innerHTML=RP_SUZANO_TIPOS.map(tipo=>{
      const color=RP_COLORS[tipo]||'#38bdf8';
      return `<div class="rp-tipo-card" style="border-color:${color}44;text-align:left;">
        <label class="rp-tipo-title" style="color:${color};display:block;">${rpSuzanoLabel(tipo)}</label>
        <input class="rp-meta-input" inputmode="decimal" value="${det[tipo]?fmtCargaCompact(det[tipo]).replace(' kg','').replace(' t',''):''}" placeholder="ex: 45,5 t" oninput="rpSetPlanejadoSuzanoTipo('${tipo}',this.value)" onchange="rpSetPlanejadoSuzanoTipo('${tipo}',this.value)" style="text-align:left;"/>
        <div class="rp-num-label" style="margin-top:7px;">Pedido Suzano para esta operação</div>
      </div>`;
    }).join('');
  }
  const info=document.getElementById('pl-info');
  if(info) info.textContent='Total Suzano: '+fmtCargaCompact(rpGetPlanejadoSuzano());
  renderTimelineChart();
}

function rpBuildEmailPayload(){
  renderReporte();
  const report=rpLastReport||{};
  const det=rpGetPlanejadoSuzanoDetalhado();
  const linhas=[
    `Reporte Operacional - ${report.dataRef||rpDataRef||dKey(today())}`,
    '',
    `Planejado Suzano: ${fmtCargaCompact(report.planejadoSuzano||0)}`,
    `- Venda Aruja: ${fmtCargaCompact(det['VENDA ARUJA']||0)}`,
    `- Venda Mogi: ${fmtCargaCompact(det['VENDA MOGI']||0)}`,
    `- Transferencia: ${fmtCargaCompact(det.TRANSFERENCIA||0)}`,
    `- Pre-fatura: ${fmtCargaCompact(det['PRE-FATURA']||0)}`,
    '',
    `Nossa grade: ${fmtCargaCompact(report.nossaGrade||0)}`,
    `Realizado: ${fmtCargaCompact(report.realizadoGrade||0)}`,
    `Pendente: ${fmtCargaCompact(Math.max(0,(report.nossaGrade||0)-(report.realizadoGrade||0)))}`,
    '',
    `DTs expedidas: ${(report.statusCounts&&report.statusCounts.EXPEDIDO)||0}`,
    `No-shows: ${(report.statusCounts&&report.statusCounts['NO SHOW'])||0}`,
    `Recusas: ${(report.statusCounts&&report.statusCounts['VEICULO RECUSADO'])||0}`,
  ];
  return {
    subject:`Reporte Operacional - ${report.dataRef||rpDataRef||dKey(today())}`,
    body:linhas.join('\n'),
    report,
    planejadoSuzanoDetalhado:det,
    config:appConfig
  };
}

async function rpEnviarEmailAgora(){
  syncConfigFromOpenModal();
  appConfig=JSON.parse(localStorage.getItem('reporte_app_config')||'{}');
  const payload=rpBuildEmailPayload();
  if(!appConfig.emailTo||!appConfig.emailTo.length){
    showErr('Configure pelo menos um destinatário antes de enviar.');
    openConfigModal();
    return;
  }
  if(appConfig.emailEndpoint){
    spin(true);
    try{
      const r=await fetch(appConfig.emailEndpoint,{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify(payload)
      });
      if(!r.ok) throw new Error(await r.text());
      showOk('Solicitação de envio enviada para a automação.');
      closeConfigModal();
    }catch(e){showErr('Erro ao chamar automação de e-mail: '+e.message);}
    spin(false);
    return;
  }
  const to=(appConfig.emailTo||[]).join(',');
  const cc=(appConfig.emailCc&&appConfig.emailCc.length)?'&cc='+encodeURIComponent(appConfig.emailCc.join(',')):'';
  window.location.href=`mailto:${to}?subject=${encodeURIComponent(payload.subject)}${cc}&body=${encodeURIComponent(payload.body)}`;
  showOk('E-mail preparado. Para envio automático direto, configure a URL da automação.');
}


function rpParseToneladas(row){
  const raw = row.peso_liquido ?? row.toneladas ?? row.peso ?? '';
  let s=String(raw).trim();
  if(!s) return 0;
  s=s.replace(/\s/g,'');
  if(s.includes(',') && s.includes('.')){
    if(s.lastIndexOf(',')>s.lastIndexOf('.')) s=s.replace(/\./g,'').replace(',','.');
    else s=s.replace(/,/g,'');
  }else if(s.includes(',')) s=s.replace(',', '.');
  const n=parseFloat(s.replace(/[^\d.-]/g,''));
  if(!Number.isFinite(n)) return 0;
  return Math.floor(n);
}

function toTonInt(raw){
  const s=String(raw??'0').trim();
  if(!s) return 0;
  if(s.includes(',')){
    const n=parseInt(s.split(',')[0].replace(/\./g,'').replace(/[^\d-]/g,''),10);
    return Number.isFinite(n)?n:0;
  }
  if(/^\d+\.\d+$/.test(s)){
    const n=Math.round(parseFloat(s));
    return Number.isFinite(n)?n:0;
  }
  const n=parseInt(s.replace(/\./g,'').replace(/[^\d-]/g,''),10);
  return Number.isFinite(n)?n:0;
}

function parseCargaNumber(raw){
  let s=String(raw??'').trim();
  if(!s) return 0;
  s=s.replace(/\s/g,'');
  if(s.includes(',') && s.includes('.')){
    if(s.lastIndexOf(',')>s.lastIndexOf('.')) s=s.replace(/\./g,'').replace(',','.');
    else s=s.replace(/,/g,'');
  }else if(s.includes(',')){
    s=s.replace(',', '.');
  }else if(/^\d{1,3}(\.\d{3})+$/.test(s)){
    s=s.replace(/\./g,'');
  }
  const n=parseFloat(s.replace(/[^\d.-]/g,''));
  return Number.isFinite(n)?n:0;
}

function fmtTon(raw){
  return toTonInt(raw).toLocaleString('pt-BR');
}

function fmtCargaCompact(raw){
  const n=Number(raw)||0;
  const abs=Math.abs(n);
  if(abs>=1000){
    const t=Math.trunc((n/1000)*10)/10;
    const txt=t.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1});
    return txt+' t';
  }
  return Math.round(n).toLocaleString('pt-BR')+' kg';
}

function fmtCargaValue(raw){
  return fmtCargaCompact(raw).replace(' t','').replace(' kg','');
}

function fmtCargaUnit(raw){
  return Math.abs(Number(raw)||0)>=1000?'t':'kg';
}

function updateStatusReminder(){
  const el=document.getElementById('status-reminder');
  if(!el) return;
  const now=new Date();
  const min=now.getMinutes();
  if(min<50){
    el.style.display='none';
    el.textContent='';
    return;
  }
  const next=new Date(now);
  next.setHours(now.getHours()+1,0,0,0);
  const faltam=Math.max(1,Math.ceil((next-now)/60000));
  el.textContent='Enviar status em '+faltam+' min ('+next.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})+')';
  el.style.display='inline-flex';
}

function rpWithinWindow(dt,start,end){
  return !!dt && dt>=start && dt<=end;
}

function rpTurnoFromDate(dt){
  if(!dt) return null;
  const h=dt.getHours();
  if(h>=7 && h<=14) return 'T1';
  if(h>=15 && h<=22) return 'T2';
  return 'T3'; // 23-06
}

function normalizePaletizacaoLabel(raw){
  const base=String(raw||'')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g,'')
    .toUpperCase();

  // Qualquer variação contendo TORDE ou o rótulo já normalizado
  if(base.includes('TORDE') || base.includes('TORDESILHAS')) return 'TORDESILHAS';

  // Qualquer variação contendo PLT, PALET ou o rótulo já normalizado
  if(base.includes('PLT') || base.includes('PALET') || base.includes('PALETIZADA')) return 'PALETIZADA';

  if(base.includes('ESTIVADA')) return 'ESTIVADA';

  // Caso não encontre nenhum padrão
  return 'ESTIVADA';
}

function mergePaletizacaoLabel(currentLabel,newRaw){
  const current=normalizePaletizacaoLabel(currentLabel);
  const incoming=normalizePaletizacaoLabel(newRaw);
  const priority={ESTIVADA:0,PALETIZADA:1,TORDESILHAS:2};
  return (priority[incoming]||0)>(priority[current]||0)?incoming:current;
}

function renderReporte(){
  const horaInput=document.getElementById('rp-hora-corte');
  if(horaInput && horaInput.value!==rpHoraCorte) horaInput.value=rpHoraCorte;

  const refs=[...new Set(tableData.map(r=>rpDateKey(r)).filter(Boolean))]
    .filter(ref=>{const d=parseBR(ref);return d&&(sameDay(d,today())||sameDay(d,tomorrow()));})
    .sort(compareDateRefs);
  if(!rpDataRef || (refs.length && !refs.includes(rpDataRef))) rpDataRef=refs[0]||dKey(today());
  if(!rpSnapshotCache[rpDataRef]){
    rpSnapshotCache[rpDataRef]=[];
    rpLoadSnapshots(rpDataRef).then(renderTimelineChart).catch(()=>renderTimelineChart());
  }

  const diaSelect=document.getElementById('rp-data-ref');
  if(diaSelect){
    diaSelect.innerHTML=refs.map(ref=>`<option value="${ref}">${ref}</option>`).join('');
    diaSelect.value=rpDataRef;
  }
  const suzanoInput=document.getElementById('rp-planejado-suzano');
  if(suzanoInput && document.activeElement!==suzanoInput){
    const suzanoVal=rpGetPlanejadoSuzano();
    suzanoInput.value=suzanoVal?fmtCargaCompact(suzanoVal).replace(' kg','').replace(' t',''):'';
  }

  const dayRows=tableData.filter(r=>rpDateKey(r)===rpDataRef);
  const rows=(rpTurnoFiltro==='todos'?[...dayRows]
    :dayRows.filter(r=>rpGetTurno(r)===rpTurnoFiltro)).sort(compareFimAgendaRows);
  const reportRows=rows.filter(rpRowContaNoReporte);
  const excludedRows=rows.length-reportRows.length;

  const infoEl=document.getElementById('rp-turno-info');

  // Inputs de meta
  const metasEl=document.getElementById('rp-metas-inputs');
  if(metasEl){
    metasEl.innerHTML='';
  }

  const corte=rpGetCorteDate();
  const corteLabel=corte.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'});
  if(infoEl) infoEl.textContent=reportRows.length+' DTs no filtro atual · corte até '+corteLabel+
    (excludedRows?` · ${excludedRows} NO SHOW/RECUSADA/DOCA S/ INFO fora do reporte`:'');

  const planejadoTon={};
  const realizadoTon={};
  RP_TIPOS.forEach(t=>{planejadoTon[t]=0;realizadoTon[t]=0;});

  reportRows.forEach(r=>{
    const modalidade=rpModalidade(r);
    const ton=rpParseToneladas(r);
    if(!rpRowDentroDoCorte(r,corte)) return;
    const realizado=STATUS_REALIZADO.includes(r.status);

    planejadoTon['GRADE']+=ton;
    if(realizado) realizadoTon['GRADE']+=ton;

    if(modalidade && planejadoTon[modalidade]!==undefined){
      planejadoTon[modalidade]+=ton;
      if(realizado) realizadoTon[modalidade]+=ton;
    }
  });

  const planejadoSuzano=rpGetPlanejadoSuzano();
  const nossaGrade=planejadoTon['GRADE']||0;
  const realizadoGrade=realizadoTon['GRADE']||0;
  const deltaSuzanoGrade=nossaGrade-planejadoSuzano;
  const deltaRealGrade=realizadoGrade-nossaGrade;
  const pctGrade=nossaGrade>0?Math.round(realizadoGrade/nossaGrade*100):0;
  const compEl=document.getElementById('rp-comparativo-suzano');
  if(compEl){
    const deltaLabel=(v)=>`${v>=0?'+':''}${fmtCargaCompact(v)}`;
    compEl.innerHTML=[
      {title:'PLANEJADO SUZANO',value:planejadoSuzano,color:'#38bdf8',sub:planejadoSuzano?'Informado manualmente':'Informe no campo SUZANO'},
      {title:'NOSSA GRADE',value:nossaGrade,color:'#60a5fa',sub:planejadoSuzano?`vs Suzano: ${deltaLabel(deltaSuzanoGrade)}`:'Grade atual'},
      {title:'REALIZADO',value:realizadoGrade,color:'#22c55e',sub:`${pctGrade}% da nossa grade`},
      {title:'PENDENTE',value:Math.max(0,nossaGrade-realizadoGrade),color:'#f59e0b',sub:`vs grade: ${deltaLabel(deltaRealGrade)}`},
    ].map(c=>`<div class="rp-tipo-card" style="border-color:${c.color}44;">
      <div class="rp-tipo-title" style="color:${c.color}">${c.title}</div>
      <div class="rp-turno-num" style="color:${c.color};font-size:24px;">${fmtCargaCompact(c.value)}</div>
      <div class="rp-num-label" style="margin-top:5px;">${c.sub}</div>
    </div>`).join('');
  }
  renderTimelineChart();

  rpLastReport={
    dataRef:rpDataRef,
    corteLabel,
    planejadoSuzano,
    planejadoSuzanoDetalhado:rpGetPlanejadoSuzanoDetalhado(),
    nossaGrade,
    realizadoGrade,
    variacaoSuzanoGrade:deltaSuzanoGrade,
    tipos:RP_TIPOS.map(tipo=>({
      tipo,
      planejado:planejadoTon[tipo]||0,
      realizado:realizadoTon[tipo]||0,
      pendente:Math.max(0,(planejadoTon[tipo]||0)-(realizadoTon[tipo]||0)),
    })),
    statusCounts:{},
    paletizacao:null,
    turnos:null,
  };

// Gráfico de barras
  const chart=document.getElementById('rp-chart');
  if(chart){
    const maxVal=Math.max(1,...RP_TIPOS.map(t=>Math.max(planejadoTon[t]||0,realizadoTon[t]||0)));
    chart.innerHTML=RP_TIPOS.map(tipo=>{
      const plan=planejadoTon[tipo]||0;
      const real=realizadoTon[tipo]||0;
      const pend=Math.max(0,plan-real);
      const hPlan=plan?Math.round((plan/maxVal)*130):0;
      const hReal=real?Math.round((real/maxVal)*130):0;
      const hPend=pend?Math.round((pend/maxVal)*130):0;
      const c=RP_COLORS[tipo]||'#3b82f6';
      return `<div class="rp-bar-group">
        <div class="rp-bar-wrap">
          <div class="rp-bar" style="height:${hPlan}px;background:#3b82f6;width:16px;" title="Planejado: ${plan}"></div>
          <div class="rp-bar" style="height:${hReal}px;background:#22c55e;width:16px;" title="Realizado: ${real}"></div>
          <div class="rp-bar" style="height:${hPend}px;background:#f59e0b;width:16px;" title="Pendente: ${pend}"></div>
        </div>
        <div class="rp-bar-label" style="color:${c}">${tipo}</div>
      </div>`;
    }).join('');
  }

  // Números consolidados
  const numEl=document.getElementById('rp-numeros');
  if(numEl){
    numEl.innerHTML=RP_TIPOS.map(tipo=>{
      const realTon=(realizadoTon[tipo]||0);
      const plan = planejadoTon[tipo]||0;
      const realStr=fmtCargaValue(realTon);
      const pend=Math.max(0,plan-realTon);
      const pct=plan>0?Math.min(100,Math.round(realTon/plan*100)):0;
      const c=RP_COLORS[tipo]||'#3b82f6';
      const barW=plan>0?Math.round(realTon/plan*100):0;
      return `<div class="rp-tipo-card" style="border-color:${c}44;">
        <div class="rp-tipo-title" style="color:${c}">${tipo}</div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin-bottom:8px;">
          <div><div class="rp-num-plan">${fmtCargaValue(plan)}</div><div class="rp-num-label">Planej. ${fmtCargaUnit(plan)}</div></div>
          <div><div class="rp-num-real">${realStr}</div><div class="rp-num-label">Real. ${fmtCargaUnit(realTon)}</div></div>
          <div><div class="rp-num-pend">${fmtCargaValue(pend)}</div><div class="rp-num-label">Pend. ${fmtCargaUnit(pend)}</div></div>
        </div>
        <div style="background:#0f172a;border-radius:4px;height:4px;margin-bottom:6px;overflow:hidden;">
          <div style="height:100%;width:${barW}%;background:${c};border-radius:4px;transition:.4s;"></div>
        </div>
        <div style="display:flex;justify-content:space-between;font-size:10px;color:#64748b;">
          <span>${fmtCargaCompact(realTon)} realizadas</span><span style="color:${pct>=100?'#22c55e':c}">${pct}%</span>
        </div>
      </div>`;
    }).join('');
  }

  // Divisao por paletizacao
  const paletResumoEl=document.getElementById('rp-paletizacao-resumo');
  const paletizacaoResumo={
    PALETIZADA:{label:'PALETIZADA',count:0,ton:0,color:'#10b981'},
    TORDESILHAS:{label:'TORDESILHAS',count:0,ton:0,color:'#e879f9'},
    ESTIVADA:{label:'ESTIVADA',count:0,ton:0,color:'#818cf8'},
  };
  reportRows.forEach(r=>{
    const fim=rpRefDate(r);
    if(!fim) return;
    if(!rpRowDentroDoCorte(r,corte)) return;
    const key=normalizePaletizacaoLabel(r.paletizacao);
    if(!paletizacaoResumo[key]) return;
    paletizacaoResumo[key].count++;
    paletizacaoResumo[key].ton+=rpParseToneladas(r);
  });
  if(rpLastReport) rpLastReport.paletizacao=paletizacaoResumo;
  if(paletResumoEl){
    paletResumoEl.innerHTML=['PALETIZADA','TORDESILHAS','ESTIVADA'].map(key=>{
      const r=paletizacaoResumo[key];
      return `<div class="rp-turno-card" style="border-color:${r.color}44;">
        <div class="rp-turno-label" style="color:${r.color}">${r.label}</div>
        <div class="rp-turno-num" style="color:${r.color}">${r.count}</div>
        <div style="font-size:10px;color:#64748b;">DTs</div>
        <div style="font-size:12px;color:#e2e8f0;margin-top:6px;font-weight:800;">${fmtCargaCompact(r.ton)}</div>
        <div style="font-size:10px;color:#64748b;">planejados</div>
      </div>`;
    }).join('');
  }

  // Contagem por Status atual
  const statusCountEl=document.getElementById('rp-status-counts');
  if(statusCountEl){
    const statusList=['AG CHEGADA','PATIO','CARREGANDO','EM FATURAMENTO','SEPARANDO','EXPEDIDO','NO SHOW','VEICULO RECUSADO','FOI EMBORA'];
    const statusColors={'AG CHEGADA':'#f59e0b','PATIO':'#64748b','CARREGANDO':'#3b82f6','EM FATURAMENTO':'#06b6d4','SEPARANDO':'#8b5cf6','EXPEDIDO':'#22c55e','NO SHOW':'#ef4444','VEICULO RECUSADO':'#dc2626','FOI EMBORA':'#6b7280'};
    const sCounts={};
    reportRows.forEach(r=>{ sCounts[r.status]=(sCounts[r.status]||0)+1; });
    if(rpLastReport) rpLastReport.statusCounts={...sCounts};
    statusCountEl.innerHTML=statusList
      .filter(s=>sCounts[s]>0)
      .map(s=>{
        const c=statusColors[s]||'#334155';
        return `<div class="rp-turno-card" style="border-color:${c}44;">
          <div class="rp-turno-label" style="color:${c};font-size:9px;">${s}</div>
          <div class="rp-turno-num" style="color:${c}">${sCounts[s]}</div>
          <div style="font-size:10px;color:#64748b;">DTs</div>
        </div>`;
      }).join('');
  }

  // Separações por turno
  const turnosEl=document.getElementById('rp-turnos-resumo');
  if(turnosEl){
    const byTurno={T1:0,T2:0,T3:0};
    const paletByTurno={
      T1:{TORDESILHAS:0,PALETIZADA:0,ESTIVADA:0},
      T2:{TORDESILHAS:0,PALETIZADA:0,ESTIVADA:0},
      T3:{TORDESILHAS:0,PALETIZADA:0,ESTIVADA:0},
    };
    let t3DiaProd=0;
    let t3TurnoProd=0;
    const hoje0=parseBR(rpDataRef||'')||today(); hoje0.setHours(0,0,0,0);
    const amanha0=new Date(hoje0); amanha0.setDate(amanha0.getDate()+1);
    const t3DiaIni=new Date(hoje0); t3DiaIni.setHours(23,0,0,0);
    const t3DiaFim=new Date(hoje0); t3DiaFim.setHours(23,59,59,999);
    const t3TurnoIni=new Date(hoje0); t3TurnoIni.setHours(23,0,0,0);
    const t3TurnoFim=new Date(amanha0); t3TurnoFim.setHours(6,59,59,999);

    reportRows.forEach(r=>{
      const fim=rpRefDate(r);
      if(!fim) return;

      // Turno noturno cruza a meia-noite:
      // - fechamento do dia: 23:00-23:59 do dia selecionado
      // - continuidade do turno: 00:00-06:59 do dia seguinte
      const dentroJanelaDia = sameDay(fim,hoje0);
      const dentroJanelaNoiteSeguinte = fim>=t3TurnoIni && fim<=t3TurnoFim;
      if(!dentroJanelaDia && !dentroJanelaNoiteSeguinte) return;

      const turno=rpTurnoFromDate(fim);
      if(!turno) return;
      const ton=rpParseToneladas(r);
      byTurno[turno]+=ton;

      const paletKey=normalizePaletizacaoLabel(r.paletizacao);
      paletByTurno[turno][paletKey]++;

      if(turno==='T3'){
        if(fim>=t3DiaIni && fim<=t3DiaFim) t3DiaProd+=ton;         // produtividade do dia
        if(fim>=t3TurnoIni && fim<=t3TurnoFim) t3TurnoProd+=ton;   // produtividade do turno (23:00-06:59)
      }
    });
    const totalTurnos=byTurno.T1+byTurno.T2+byTurno.T3;
    const labels={T1:'🌅 T1 · 07-14h',T2:'☀️ T2 · 15-22h',T3:'🌙 T3 · 23-06h'};
    const cores={T1:'#60a5fa',T2:'#f59e0b',T3:'#a78bfa'};
    const cards=['T1','T2','T3'].map(t=>`
      <div class="rp-turno-card" style="border-color:${cores[t]}44;">
        <div class="rp-turno-label" style="color:${cores[t]}">${labels[t]}</div>
        <div class="rp-turno-num" style="color:${cores[t]}">${fmtCargaValue(byTurno[t])}</div>
        <div style="font-size:10px;color:#64748b;">${fmtCargaUnit(byTurno[t])} planejados</div>
        <div style="font-size:10px;color:#94a3b8;margin-top:6px;line-height:1.5;">
          Tordesilhas: <b style="color:#c4b5fd;">${paletByTurno[t].TORDESILHAS}</b> ·
          Paletizada: <b style="color:#c4b5fd;">${paletByTurno[t].PALETIZADA}</b> ·
          Estivada: <b style="color:#c4b5fd;">${paletByTurno[t].ESTIVADA}</b>
        </div>
        ${t==='T3'
          ? `<div style="font-size:10px;color:#94a3b8;margin-top:6px;line-height:1.6;">
              Dia (23-00): <b style="color:#c4b5fd;">${fmtCargaCompact(t3DiaProd)}</b><br/>
              Turno (23-06): <b style="color:#c4b5fd;">${fmtCargaCompact(t3TurnoProd)}</b>
            </div>`
          : ''
        }
      </div>
    `).join('');
    const totalCard=`
      <div class="rp-turno-card" style="border-color:#22c55e44;">
        <div class="rp-turno-label" style="color:#22c55e">📦 TOTAL DO DIA</div>
        <div class="rp-turno-num" style="color:#22c55e">${fmtCargaValue(totalTurnos)}</div>
        <div style="font-size:10px;color:#64748b;">${fmtCargaUnit(totalTurnos)} planejados</div>
      </div>
    `;
    if(rpLastReport){
      rpLastReport.turnos={...byTurno,total:totalTurnos,t3DiaProd,t3TurnoProd,paletByTurno};
    }
    turnosEl.innerHTML=cards+totalCard;
  }
}


/* ═══════════════════════════════════════════════════════
   REGRA 12: EXPORTAR REPORTE (XLSX / CSV)
═══════════════════════════════════════════════════════ */
function exportReporteCSV(){
  // Reutiliza a mesma lógica do exportCSV() já existente — apenas garante
  // que é chamado do contexto do reporte para consistência.
  exportCSV();
}

function exportReporteXLSX(){
  if(typeof XLSX==='undefined'){showErr('Biblioteca XLSX não carregada. Recarregue a página.');return;}
  const cols=['DIA','DT','TRANSPORTADORA','GRADE','FIM','HORA CHEGADA','N° PORTARIA','STATUS','DESC. DOCUMENTO','PESO LÍQUIDO','TIPO OPERAÇÃO','MATERIAL','PALETES','SOBRA (FARDOS)','QTD TOTAL (FARDOS)'];
  const wsData=[cols];
  const exportRows=dedupeCargaRowsByDTRef((tableData||[])
    .filter(isRowInActiveReportWindow)
    .filter(rpRowContaNoReporte)
  ).sort(compareAgendaRows);
  exportRows.forEach(r=>{
    const dtKey=String(r.dt||'').trim();
    const mats=exportMap[dtKey]||[];
    const base=[r.dia_ref,r.dt,r.transportadora,csvGradeValue(r.grade_carregamento),csvGradeValue(r.fim_carregamento),r.hora_chegada,r.n_portaria||'',r.status,r.descricao_documento||'',fmtTon(r.peso_liquido||r.toneladas||0),r.tipo_operacao||''];
    if(!mats.length){
      wsData.push([...base,'','','','']);
    } else {
      mats.forEach((m,i)=>{
        const pal=formatPaletes(m.material,m.quantidade);
        const qtdFardos=parseQuantidadeFardos(m.quantidade);
        if(i===0){
          wsData.push([...base,m.material,pal.paletes,pal.sobra,Math.round(qtdFardos)]);
        } else {
          wsData.push(['','','','','','','','','','','',m.material,pal.paletes,pal.sobra,Math.round(qtdFardos)]);
        }
      });
    }
  });
  const ws=XLSX.utils.aoa_to_sheet(wsData);
  // Destaque de cabeçalho
  const range=XLSX.utils.decode_range(ws['!ref']);
  for(let C=range.s.c;C<=range.e.c;C++){
    const cellAddr=XLSX.utils.encode_cell({r:0,c:C});
    if(!ws[cellAddr]) continue;
    ws[cellAddr].s={font:{bold:true},fill:{fgColor:{rgb:'1E293B'}}};
  }
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Reporte');
  const filename='reporte_'+new Date().toLocaleDateString('pt-BR').replace(/\//g,'-')+'.xlsx';
  XLSX.writeFile(wb,filename);
}

function exportReporteJPG(){
  renderReporte();
  if(!rpLastReport){showErr('Reporte ainda sem dados para exportar.');return;}
  const W=1300,H=900;
  const canvas=document.createElement('canvas');
  canvas.width=W; canvas.height=H;
  const ctx=canvas.getContext('2d');
  const rr=(x,y,w,h,r,fill,stroke)=>{
    ctx.beginPath();
    ctx.moveTo(x+r,y);
    ctx.arcTo(x+w,y,x+w,y+h,r);
    ctx.arcTo(x+w,y+h,x,y+h,r);
    ctx.arcTo(x,y+h,x,y,r);
    ctx.arcTo(x,y,x+w,y,r);
    ctx.closePath();
    ctx.fillStyle=fill; ctx.fill();
    if(stroke){ctx.strokeStyle=stroke;ctx.lineWidth=1;ctx.stroke();}
  };
  const text=(s,x,y,size,color,weight='600',align='left')=>{
    ctx.font=weight+' '+size+'px Segoe UI, Arial';
    ctx.fillStyle=color;
    ctx.textAlign=align;
    ctx.fillText(String(s),x,y);
  };

  ctx.fillStyle='#0b1220';
  ctx.fillRect(0,0,W,H);
  text('Reporte de Status',40,54,28,'#e2e8f0','800');
  text('Dia '+rpLastReport.dataRef+'  |  Corte ate '+rpLastReport.corteLabel,40,84,15,'#94a3b8','600');
  if(rpLastReport.planejadoSuzano){
    text('Suzano '+fmtCargaCompact(rpLastReport.planejadoSuzano)+'  |  Nossa grade '+fmtCargaCompact(rpLastReport.nossaGrade)+'  |  Realizado '+fmtCargaCompact(rpLastReport.realizadoGrade),40,106,13,'#38bdf8','700');
  }
  text(new Date().toLocaleString('pt-BR'),W-40,54,14,'#64748b','600','right');

  const colors=RP_COLORS;
  let x=40,y=118;
  rpLastReport.tipos.forEach((r,i)=>{
    const c=colors[r.tipo]||'#60a5fa';
    const cardX=x+(i%2)*390;
    const cardY=y+Math.floor(i/2)*128;
    rr(cardX,cardY,360,104,10,'#111c2e',c+'66');
    text(r.tipo,cardX+18,cardY+28,13,c,'800');
    text(fmtCargaCompact(r.planejado),cardX+18,cardY+66,22,'#60a5fa','800');
    text('Planejado',cardX+18,cardY+88,11,'#64748b','700');
    text(fmtCargaCompact(r.realizado),cardX+154,cardY+66,22,'#22c55e','800');
    text('Realizado',cardX+154,cardY+88,11,'#64748b','700');
    text(fmtCargaCompact(r.pendente),cardX+282,cardY+66,22,'#f59e0b','800','center');
    text('Pendente',cardX+282,cardY+88,11,'#64748b','700','center');
  });

  const rightX=850;
  rr(rightX,118,410,250,10,'#111c2e','#334155');
  text('Status Atual das DTs',rightX+20,150,16,'#e2e8f0','800');
  const statusColors={'AG CHEGADA':'#f59e0b','PATIO':'#64748b','CARREGANDO':'#3b82f6','EM FATURAMENTO':'#06b6d4','SEPARANDO':'#8b5cf6','EXPEDIDO':'#22c55e','NO SHOW':'#ef4444','VEICULO RECUSADO':'#dc2626','FOI EMBORA':'#6b7280'};
  let sy=184;
  Object.entries(rpLastReport.statusCounts||{}).forEach(([st,n])=>{
    const c=statusColors[st]||'#94a3b8';
    text(st,rightX+22,sy,13,c,'800');
    text(n,rightX+360,sy,18,'#e2e8f0','800','right');
    sy+=28;
  });

  rr(rightX,398,410,250,10,'#111c2e','#334155');
  text('Separacoes por Turno',rightX+20,430,16,'#e2e8f0','800');
  const t=rpLastReport.turnos||{T1:0,T2:0,T3:0,total:0,t3DiaProd:0,t3TurnoProd:0,paletByTurno:{}};
  [
    ['T1 07-14h',t.T1,'#60a5fa','T1'],
    ['T2 15-22h',t.T2,'#f59e0b','T2'],
    ['T3 23-06h',t.T3,'#a78bfa','T3'],
    ['Total do dia',t.total,'#22c55e',''],
  ].forEach((row,idx)=>{
    const yy=468+idx*38;
    text(row[0],rightX+22,yy,13,row[2],'800');
    text(fmtCargaCompact(row[1]),rightX+360,yy,18,row[2],'800','right');
    if(row[3]){
      const p=(t.paletByTurno&&t.paletByTurno[row[3]])||{TORDESILHAS:0,PALETIZADA:0,ESTIVADA:0};
      text('Tord: '+p.TORDESILHAS+'  |  Pal: '+p.PALETIZADA+'  |  Est: '+p.ESTIVADA,rightX+22,yy+16,10,'#94a3b8','700');
    }
  });
  text('T3 dia 23-00: '+fmtCargaCompact(t.t3DiaProd),rightX+22,628,12,'#94a3b8','700');
  text('T3 turno 23-06: '+fmtCargaCompact(t.t3TurnoProd),rightX+210,628,12,'#94a3b8','700');

  rr(40,540,760,170,10,'#111c2e','#334155');
  text('Planejado x Realizado',64,574,16,'#e2e8f0','800');
  const max=Math.max(1,...rpLastReport.tipos.flatMap(r=>[r.planejado,r.realizado,r.pendente]));
  rpLastReport.tipos.forEach((r,i)=>{
    const bx=80+i*138;
    const base=680;
    const vals=[
      [r.planejado,'#3b82f6'],
      [r.realizado,'#22c55e'],
      [r.pendente,'#f59e0b'],
    ];
    vals.forEach((v,j)=>{
      const h=Math.round((v[0]/max)*78);
      ctx.fillStyle=v[1];
      ctx.fillRect(bx+j*18,base-h,12,h);
    });
    text(r.tipo.replace('TRANSFERENCIA','TRANSF.'),bx+18,704,10,colors[r.tipo]||'#94a3b8','800','center');
  });

  rr(40,730,760,120,10,'#111c2e','#334155');
  text('Divisao por Paletizacao',64,762,16,'#e2e8f0','800');
  const paletJpg=rpLastReport.paletizacao||{};
  [
    ['PALETIZADA','#10b981'],
    ['TORDESILHAS','#e879f9'],
    ['ESTIVADA','#818cf8'],
  ].forEach((row,idx)=>{
    const p=paletJpg[row[0]]||{count:0,ton:0};
    const px=74+idx*240;
    text(row[0],px,794,13,row[1],'800');
    text(String(p.count||0),px,825,28,row[1],'800');
    text('DTs',px+42,825,12,'#64748b','700');
    text(fmtCargaCompact(p.ton||0),px+112,825,20,'#e2e8f0','800');
  });

  downloadCanvasJPG(
    canvas,
    'reporte_status_'+String(rpLastReport.dataRef||'').replace(/\//g,'-')+'_'+String(rpLastReport.corteLabel||'').replace(':','h')+'.jpg',
    0.92
  );
  exportGradeTodasDTsJPG();
}

function downloadCanvasJPG(canvas,filename,quality=0.92,delay=0){
  canvas.toBlob(blob=>{
    if(!blob){showErr('Nao foi possivel gerar a imagem.');return;}
    setTimeout(()=>{
      const a=document.createElement('a');
      a.href=URL.createObjectURL(blob);
      a.download=filename;
      a.click();
      setTimeout(()=>URL.revokeObjectURL(a.href),1500);
    },delay);
  },'image/jpeg',quality);
}

function exportGradeTodasDTsJPG(){
  const rows=[...tableData]
    .filter(r=>!isFinalStatus(r.status))
    .sort(compareTableRows);
  const visibleRows=rows.slice(0,54);
  const W=1800;
  const rowH=34;
  const H=Math.max(820,150+visibleRows.length*rowH+70);
  const canvas=document.createElement('canvas');
  canvas.width=W; canvas.height=H;
  const ctx=canvas.getContext('2d');
  const text=(s,x,y,size,color,weight='600',align='left',maxWidth=null)=>{
    ctx.font=weight+' '+size+'px Segoe UI, Arial';
    ctx.fillStyle=color;
    ctx.textAlign=align;
    const value=String(s||'');
    if(maxWidth) ctx.fillText(value,x,y,maxWidth);
    else ctx.fillText(value,x,y);
  };
  const rect=(x,y,w,h,fill,stroke=null)=>{
    ctx.fillStyle=fill; ctx.fillRect(x,y,w,h);
    if(stroke){ctx.strokeStyle=stroke;ctx.lineWidth=1;ctx.strokeRect(x,y,w,h);}
  };
  const ellipsis=(value,maxChars)=>{
    const s=String(value||'');
    return s.length>maxChars?s.slice(0,Math.max(0,maxChars-1))+'…':s;
  };
  const paletColor=(raw)=>{
    const p=normalizePaletizacaoLabel(raw);
    if(p==='PALETIZADA') return '#10b981';
    if(p==='TORDESILHAS') return '#e879f9';
    return '#818cf8';
  };
  const statusColors={'AG CHEGADA':'#f59e0b','PATIO':'#64748b','CARREGANDO':'#3b82f6','EM FATURAMENTO':'#06b6d4','SEPARANDO':'#8b5cf6','EXPEDIDO':'#22c55e','NO SHOW':'#ef4444','VEICULO RECUSADO':'#dc2626','FOI EMBORA':'#6b7280'};
  const cols=[
    ['DIA',40,90],
    ['DT',132,115],
    ['TRANSPORTADORA',252,250],
    ['PALETIZACAO',510,145],
    ['GRADE',665,180],
    ['FIM',852,180],
    ['STATUS',1040,190],
    ['OPERACAO',1238,170],
    ['DESC. DOCUMENTO',1418,190],
    ['PESO',1628,110],
  ];

  rect(0,0,W,H,'#0b1220');
  text('Grade - Todas as DTs',40,54,30,'#e2e8f0','800');
  text(new Date().toLocaleString('pt-BR'),W-40,54,15,'#64748b','600','right');
  text(rows.length+' DTs abertas na grade atual',40,84,15,'#94a3b8','700');
  if(rows.length>visibleRows.length) text('Mostrando primeiras '+visibleRows.length+' de '+rows.length+' DTs',W-40,84,13,'#f59e0b','700','right');

  rect(32,110,W-64,34,'#162032','#334155');
  cols.forEach(([label,x])=>text(label,x,132,12,'#94a3b8','800'));

  visibleRows.forEach((r,i)=>{
    const y=150+i*rowH;
    rect(32,y-22,W-64,rowH,i%2?'#0f172a':'#111c2e','#1e293b');
    const palet=normalizePaletizacaoLabel(r.paletizacao);
    const status=String(r.status||'');
    const op=String(r.tipo_operacao||'');
    text(r.dia_ref||'',cols[0][1],y,12,r.dia_ref==='HOJE'?'#60a5fa':'#a78bfa','800');
    text(r.dt||'',cols[1][1],y,14,'#93c5fd','800');
    text(ellipsis(r.transportadora||'',30),cols[2][1],y,12,'#cbd5e1','700');
    text(palet,cols[3][1],y,12,paletColor(palet),'800');
    text(String(r.grade_carregamento||'—').replace(',',''),cols[4][1],y,12,'#e2e8f0','700');
    text(String(r.fim_carregamento||'—').replace(',',''),cols[5][1],y,12,'#e2e8f0','700');
    text(status,cols[6][1],y,12,statusColors[status]||'#94a3b8','800');
    text(ellipsis(op||'—',18),cols[7][1],y,12,'#c4b5fd','800');
    text(ellipsis(r.descricao_documento||'—',24),cols[8][1],y,12,'#94a3b8','700');
    text(fmtTon(r.peso_liquido||r.toneladas||0),cols[9][1]+86,y,13,'#e2e8f0','800','right');
  });

  const footerY=H-28;
  text('Gerado pelo dashboard de carga',40,footerY,12,'#475569','700');
  downloadCanvasJPG(
    canvas,
    'grade_todas_dts_'+new Date().toLocaleDateString('pt-BR').replace(/\//g,'-')+'.jpg',
    0.92,
    350
  );
}

/* ═══════════════════════════════════════════════════════
   DT SEM GRADE
═══════════════════════════════════════════════════════ */
let sgRegistros = JSON.parse(localStorage.getItem('sg_registros')||'[]');

async function registrarSemGrade(){
  const dt=document.getElementById('sg-dt').value.trim();
  const transp=document.getElementById('sg-transp').value.trim();
  const motivo=document.getElementById('sg-motivo').value.trim();
  const obs=document.getElementById('sg-obs').value.trim();
  const gradeIniRaw=document.getElementById('sg-grade').value.trim();
  const gradeFimRaw=document.getElementById('sg-fim').value.trim();
  if(!dt){showErr('Informe o número da DT.');return;}
  const dtNorm=normalizeDT(dt);
  const fmtLocal=(v)=>{
    if(!v) return '';
    const d=new Date(v);
    return Number.isNaN(d.getTime())?'':d.toLocaleString('pt-BR',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
  };
  const gradeIni=fmtLocal(gradeIniRaw);
  const gradeFim=fmtLocal(gradeFimRaw);
  const reg={dt:dtNorm,transp,motivo,obs,gradeIni,gradeFim,ts:new Date().toLocaleString('pt-BR')};
  sgRegistros.unshift(reg);
  try{localStorage.setItem('sg_registros',JSON.stringify(sgRegistros));}catch(e){}
  const dataRef=dKey(today());
  const row={dt:String(dtNorm),transportadora:transp,grade_carregamento:gradeIni,fim_carregamento:gradeFim,hora_chegada:'',n_portaria:'',status:'AG CHEGADA',descricao_documento:motivo||'ADICIONADA MANUAL',toneladas:'',peso_liquido:'',agenda:gradeIni||'',local_cd:'',dia_ref:'HOJE',data_ref:dataRef,tipo_operacao:'',reagendada:false};
  try{
    const exists=await sbGet('reporte_carga',`dt=eq.${encodeURIComponent(dtNorm)}&data_ref=eq.${encodeURIComponent(dataRef)}&select=dt&limit=1`);
    if(!exists||!exists.length) await sbInsert('reporte_carga',[row]);
  }catch(e){/* tabela pode não existir ainda */}
  const idx=tableData.findIndex(r=>String(r.dt)===String(dtNorm)&&r.data_ref===dataRef);
  if(idx===-1) tableData.push(row); else tableData[idx]={...tableData[idx],...row};
  document.getElementById('sg-dt').value='';
  document.getElementById('sg-transp').value='';
  document.getElementById('sg-motivo').value='';
  document.getElementById('sg-obs').value='';
  document.getElementById('sg-grade').value='';
  document.getElementById('sg-fim').value='';
  showOk('DT '+dtNorm+' adicionada e incluída na grade principal.');
  renderSemGrade();
  renderRows();
}

/* ═══════════════════════════════════════════════════════
   REGRA 10: AUDITORIA DE DTs SEM DOCA
═══════════════════════════════════════════════════════ */
function renderDocaNullAudit(){
  const list=document.getElementById('docanull-list');
  const empty=document.getElementById('docanull-empty');
  if(!list) return;
  const rows=(tableData||[]).filter(r=>r.doca_null);
  if(!rows.length){
    list.innerHTML='';
    if(empty) empty.style.display='block';
    return;
  }
  if(empty) empty.style.display='none';
  list.innerHTML=rows.map(r=>`
    <div style="background:#1e293b;border:1.5px solid #f59e0b55;border-radius:8px;padding:12px 16px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;">
      <span style="font-size:14px;font-weight:900;color:#f59e0b;min-width:90px;">⚠️ ${escHtml(r.dt)}</span>
      <span style="font-size:12px;color:#94a3b8;flex:1;">${escHtml(r.transportadora||'—')}</span>
      <span style="font-size:11px;color:#64748b;">Grade: ${escHtml(r.grade_carregamento||'—')}</span>
      <span style="font-size:11px;color:#64748b;">Fim: ${escHtml(r.fim_carregamento||'—')}</span>
      <span style="font-size:11px;color:#64748b;">Status: ${escHtml(r.status||'—')}</span>
      <button class="bp" onclick="toggleDocaNullReporte('${r.dt}','${r.data_ref}',true)" style="font-size:10px;padding:3px 10px;background:#b45309;border-color:#f59e0b;color:#fff7ed;">✅ Liberar no reporte</button>
      <button class="bg" onclick="docaNullExcluir('${r.dt}','${r.data_ref}')" style="font-size:10px;padding:3px 10px;border-color:#ef444455;color:#ef4444;">🗑️ Excluir</button>
    </div>
  `).join('');
}

async function docaNullExcluir(dt, dataRef){
  if(!confirm('Excluir DT '+dt+' (DOCA não informada)?\nEsta ação não pode ser desfeita.')) return;
  try{
    await sbPatch('reporte_carga',{status:'DT EXCLUIDA',updated_at:new Date().toISOString(),source:'AUDIT_DOCA_NULL'},{dt,data_ref:dataRef});
    await tryInsertLog({dt,data_ref:dataRef,event_type:'DT_EXCLUIDA_MANUAL',status:'DT EXCLUIDA',transportadora:'',created_at:new Date().toISOString(),source:'AUDIT_DOCA_NULL',payload:{motivo:'Excluída manualmente via auditoria DOCA NULL'}});
    showOk('DT '+dt+' excluída.');
    await reloadTable();
    renderDocaNullAudit();
  }catch(e){showErr('Erro ao excluir: '+e.message);}
}

function renderSemGrade(){
  const lista=document.getElementById('sg-lista');
  if(!lista) return;
  if(!sgRegistros.length){
    lista.innerHTML='<div style="color:#334155;font-size:12px;padding:10px;">Nenhuma DT sem grade registrada ainda.</div>';
    return;
  }
  lista.innerHTML=sgRegistros.map((r,i)=>`
    <div class="sg-item">
      <div><div style="font-weight:800;color:#f59e0b;">${r.dt}</div><div style="font-size:10px;color:#64748b;">${r.ts}</div></div>
      <div style="font-size:11px;color:#cbd5e1;">${r.transp||'—'}</div>
      <div style="font-size:11px;color:#94a3b8;">${r.motivo||'—'}${r.obs?' · '+r.obs:''}${r.gradeIni?` · 🕒 ${r.gradeIni}${r.gradeFim?` → ${r.gradeFim}`:''}`:''}</div>
      <button class="bg" onclick="sgRemover(${i})" style="font-size:10px;padding:3px 8px;">✕</button>
    </div>
  `).join('');
}

function sgRemover(i){
  sgRegistros.splice(i,1);
  try{localStorage.setItem('sg_registros',JSON.stringify(sgRegistros));}catch(e){}
  renderSemGrade();
}

/* ═══════════════════════════════════════════════════════
   DT FAKE — SUBSTITUIÇÃO
═══════════════════════════════════════════════════════ */
let fakeSubstData=null;

async function substituirDTFake(){
  const dtFrom=document.getElementById('fake-dt-from').value.trim();
  const dtTo=document.getElementById('fake-dt-to').value.trim();
  const statusEl=document.getElementById('fake-status');
  const preview=document.getElementById('fake-preview');
  const previewBody=document.getElementById('fake-preview-body');
  if(!dtFrom||!dtTo){showErr('Informe os dois números de DT.');return;}
  if(dtFrom===dtTo){showErr('DT fake e DT original devem ser diferentes.');return;}
  statusEl.textContent='🔍 Buscando…';
  preview.style.display='none';
  try{
    // Busca DT fake no banco
    const refs=await getActiveRefs();
    const fromRows=await sbGet('reporte_carga',`dt=eq.${encodeURIComponent(dtFrom)}&select=*&limit=5`);
    const fromMats=await sbGet('reporte_materiais',`dt=eq.${encodeURIComponent(dtFrom)}&select=*`);
    if(!fromRows||!fromRows.length){showErr('DT fake '+dtFrom+' não encontrada no banco.');statusEl.textContent='';return;}
    const row=fromRows[0];
    fakeSubstData={dtFrom,dtTo,row,mats:fromMats||[]};
    previewBody.innerHTML=`
      <div style="margin-bottom:10px;">
        <b style="color:#f59e0b;">DT ${dtFrom}</b> será renomeada para <b style="color:#22c55e;">DT ${dtTo}</b>
      </div>
      <div style="font-size:11px;color:#94a3b8;margin-bottom:8px;">
        Transportadora: ${row.transportadora||'—'} &nbsp;|&nbsp; Grade: ${row.grade_carregamento||'—'}<br/>
        Materiais atuais (serão substituídos): ${fakeSubstData.mats.length} item(s)
      </div>
      <div style="font-size:11px;color:#f59e0b;">⚠️ Os materiais da DT fake serão removidos. Após a substituição, adicione os materiais corretos no painel da DT ${dtTo}.</div>
    `;
    preview.style.display='block';
    statusEl.textContent='';
  }catch(e){showErr('Erro: '+e.message);statusEl.textContent='';}
}

async function confirmarSubstituicao(){
  if(!fakeSubstData){return;}
  const {dtFrom,dtTo,row,mats}=fakeSubstData;
  spin(true);
  try{
    // Cria nova linha com dtTo copiando dados da fake
    const novaRow={...row,dt:dtTo,updated_at:new Date().toISOString()};
    delete novaRow.id;
    await sbInsert('reporte_carga',[novaRow]);
    // Migra materiais
    if(mats.length){
      const novosMats=mats.map(m=>({...m,dt:dtTo}));
      novosMats.forEach(m=>{delete m.id;});
      await sbDelete('reporte_materiais',{dt:dtTo,data_ref:row.data_ref});
      await sbInsert('reporte_materiais',novosMats);
    }
    // Remove a DT fake
    await sbDelete('reporte_carga',{dt:dtFrom,data_ref:row.data_ref});
    await sbDelete('reporte_materiais',{dt:dtFrom,data_ref:row.data_ref});
    // Log
    await tryInsertLog({dt:dtTo,data_ref:row.data_ref,status:'SUBSTITUIÇÃO DE DT FAKE',transportadora:'DT fake: '+dtFrom,created_at:new Date().toISOString()});
    showOk('DT '+dtFrom+' substituída por '+dtTo+' com sucesso!');
    document.getElementById('fake-preview').style.display='none';
    document.getElementById('fake-dt-from').value='';
    document.getElementById('fake-dt-to').value='';
    document.getElementById('fake-status').textContent='';
    fakeSubstData=null;
    await reloadTable();
  }catch(e){showErr('Erro na substituição: '+e.message);}
  spin(false);
}

function cancelarSubstituicao(){
  fakeSubstData=null;
  document.getElementById('fake-preview').style.display='none';
  document.getElementById('fake-status').textContent='';
}

/* ═══════════════════════════════════════════════════════
   JUNTAR NFs — PDF
═══════════════════════════════════════════════════════ */
let nfArquivos=[];
let nfBlob=null;

function nfAddFiles(files){
  nfArquivos=[...nfArquivos,...Array.from(files)];
  nfBlob=null;
  nfRenderLista();
}

function nfRenderLista(){
  const lista=document.getElementById('nf-lista');
  if(!lista)return;
  if(!nfArquivos.length){lista.innerHTML='<div style="color:#334155;font-size:12px;padding:6px;">Nenhum arquivo selecionado.</div>';return;}
  lista.innerHTML=nfArquivos.map((f,i)=>`
    <div class="nf-item">
      <div>
        <b style="color:#e2e8f0;">${i+1}. ${f.name}</b>
        <span style="color:#64748b;font-size:11px;"> — ${(f.size/1024).toFixed(1)} KB</span>
      </div>
      <div style="display:flex;gap:6px;">
        ${i>0?`<button class="bg" onclick="nfMover(${i},-1)" style="padding:3px 8px;font-size:11px;">↑</button>`:'<span style="width:34px;"></span>'}
        ${i<nfArquivos.length-1?`<button class="bg" onclick="nfMover(${i},1)" style="padding:3px 8px;font-size:11px;">↓</button>`:'<span style="width:34px;"></span>'}
        <button class="bg" onclick="nfRemover(${i})" style="padding:3px 8px;font-size:11px;color:#ef4444;">✕</button>
      </div>
    </div>
  `).join('');
}

function nfMover(i,d){
  const j=i+d;
  if(j<0||j>=nfArquivos.length)return;
  [nfArquivos[i],nfArquivos[j]]=[nfArquivos[j],nfArquivos[i]];
  nfBlob=null;nfRenderLista();
}

function nfRemover(i){
  nfArquivos.splice(i,1);nfBlob=null;nfRenderLista();
}

function nfLimpar(){
  nfArquivos=[];nfBlob=null;
  nfRenderLista();
  const msg=document.getElementById('nf-msg');
  if(msg){msg.style.display='none';}
}

function nfMsg(txt,ok){
  const el=document.getElementById('nf-msg');
  if(!el)return;
  el.style.display='block';
  el.style.color=ok?'#22c55e':'#ef4444';
  el.textContent=txt;
}

async function nfGerar(){
  if(!nfArquivos.length){nfMsg('Adicione pelo menos 1 PDF.',false);return;}
  if(typeof PDFLib==='undefined'){
    // Load pdf-lib dynamically
    await new Promise((res,rej)=>{
      const s=document.createElement('script');
      s.src='https://unpkg.com/pdf-lib/dist/pdf-lib.min.js';
      s.onload=res;s.onerror=rej;
      document.head.appendChild(s);
    });
  }
  nfMsg('⏳ Juntando PDFs…',true);
  const dtNum=(document.getElementById('nf-dt-num').value.trim()||'pdf-unico');
  try{
    const merged=await PDFLib.PDFDocument.create();
    for(const file of nfArquivos){
      const bytes=await file.arrayBuffer();
      const src=await PDFLib.PDFDocument.load(bytes,{ignoreEncryption:true});
      const pages=await merged.copyPages(src,src.getPageIndices());
      pages.forEach(p=>merged.addPage(p));
    }
    const mergedBytes=await merged.save();
    nfBlob=new Blob([mergedBytes],{type:'application/pdf'});
    const a=document.createElement('a');
    a.href=URL.createObjectURL(nfBlob);
    a.download=`DT-${dtNum}.pdf`;
    document.body.appendChild(a);a.click();a.remove();
    nfMsg(`✅ PDF gerado: DT-${dtNum}.pdf (${nfArquivos.length} arquivo(s))`,true);
  }catch(e){nfMsg('Erro ao juntar: '+e.message,false);}
}

async function nfImprimir(){
  if(!nfBlob) await nfGerar();
  if(!nfBlob)return;
  const url=URL.createObjectURL(nfBlob);
  const w=window.open(url,'_blank');
  if(!w){nfMsg('Bloqueio de pop-up detectado. Baixe o PDF e imprima.',false);return;}
  setTimeout(()=>{w.focus();w.print();},800);
}

/* ═══════════════════════════════════════════════════════
   IMPORTAR PLANILHA — Ctrl+B
═══════════════════════════════════════════════════════ */
let importCsvData = null; // linhas parseadas da planilha do Ctrl+B

document.addEventListener('keydown', e => {
  if ((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='b') {
    e.preventDefault();
    openImportModal();
  }
  if ((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='p') {
    e.preventDefault();
    openPortariaPdfModal();
  }
});

function openPortariaPdfModal(){
  const ov=document.getElementById('portaria-overlay');
  if(!ov) return;
  document.getElementById('portaria-result').style.display='none';
  document.getElementById('portaria-err').style.display='none';
  document.getElementById('portaria-file').value='';
  ov.style.display='flex';
}
function closePortariaPdfModal(){ document.getElementById('portaria-overlay').style.display='none'; }

async function runPortariaPdfImport(file){
  if(!file){ showErr('Selecione um PDF da portaria.'); return; }
  spin(true); showInf('Lendo PDF da portaria…');
  try{
    if(typeof pdfjsLib==='undefined'){
      await new Promise((res,rej)=>{
        const s=document.createElement('script');
        s.src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
        s.onload=res; s.onerror=rej; document.head.appendChild(s);
      });
      pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }
    const buf=await file.arrayBuffer();
    const pdf=await pdfjsLib.getDocument({data:buf}).promise;
    let txt='';
    for(let i=1;i<=pdf.numPages;i++){
      const p=await pdf.getPage(i);
      const c=await p.getTextContent();
      txt+=c.items.map(x=>x.str).join(' ')+'\n';
    }
    const dts=[...new Set((txt.match(/\b\d{8}\b/g)||[]).map(n=>String(n).trim()))];
    if(!dts.length) throw new Error('Nenhuma DT (8 dígitos) encontrada no PDF.');
    let atualizadas=0;
    for(const dt of dts){
      const rows=tableData.filter(r=>String(r.dt)===dt);
      for(const row of rows){
        if(row.status!=='AG CHEGADA') continue;
        const patch={status:'PATIO',updated_at:new Date().toISOString()};
        if(!row.hora_chegada) patch.hora_chegada=new Date().toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'});
        await sbPatch('reporte_carga',patch,{dt:row.dt,data_ref:row.data_ref});
        await tryInsertLog({dt:row.dt,data_ref:row.data_ref,status:'PATIO',transportadora:row.transportadora||'',created_at:new Date().toISOString()});
        row.status='PATIO';
        if(!row.hora_chegada) row.hora_chegada=patch.hora_chegada;
        atualizadas++;
      }
    }
    renderRows();
    hideInf();
    const rs=document.getElementById('portaria-result');
    rs.style.display='block';
    rs.textContent=`✅ PDF processado. DTs lidas: ${dts.length}. Status alterado para PATIO: ${atualizadas}.`;
    closePortariaPdfModal();
  }catch(e){
    hideInf();
    const er=document.getElementById('portaria-err');
    er.style.display='block'; er.textContent='Erro ao importar PDF: '+e.message;
  }
  spin(false);
}

function openImportModal(){
  pauseAutoSync(180000);
  importCsvData=null;
  document.getElementById('import-btn').disabled=true;
  document.getElementById('import-drop-label').textContent='Clique ou arraste a planilha aqui';
  document.getElementById('import-preview').style.display='none';
  document.getElementById('import-result').style.display='none';
  document.getElementById('import-err').style.display='none';
  document.getElementById('import-file-input').value='';
  document.getElementById('import-overlay').style.display='flex';
}

function closeImportModal(){
  document.getElementById('import-overlay').style.display='none';
  releaseAutoSyncAfterSave();
}

// Mapa de status da planilha → status do dashboard
const CSV_STATUS_MAP = {
  'CARREGANDO':    'CARREGANDO',
  'EXPEDIDO':      'EXPEDIDO',
  'PATIO':         'PATIO',
  'EM PATIO':      'PATIO',
  'NO SHOW':       'NO SHOW',
  'SEPARANDO':     'SEPARANDO',
  'EM FATURAMENTO':'EM FATURAMENTO',
  'AG CHEGADA':    'AG CHEGADA',
  'VEICULO RECUSADO':'VEICULO RECUSADO',
  'MOTORISTA FOI EMBORA':'FOI EMBORA',
  'FOI EMBORA':    'FOI EMBORA',
  'DT EXCLUIDA':   'DT EXCLUIDA',
  'BAIXA NOCUPAÇÃO':'AG CHEGADA',
};

// Motivos da nova coluna FATURAMENTO. Quando o STATUS vem vazio,
// qualquer um deles indica que a DT está em faturamento.
const CSV_FATURAMENTO_MAP = {
  'CUSTO DE FRETE': 'EM FATURAMENTO',
  'PROBLEMA DE DT': 'EM FATURAMENTO',
  'PROBELMA DE DT': 'EM FATURAMENTO',
  'PROBLEMA JSL':   'EM FATURAMENTO',
  'PROBELMA JSL':   'EM FATURAMENTO',
};

function normalizeImportText(raw){
  return String(raw||'')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ')
    .toUpperCase();
}

function normalizeStatus(raw){
  const s=normalizeImportText(raw);
  return CSV_STATUS_MAP[s] || null;
}

function normalizeFaturamentoStatus(raw){
  const s=normalizeImportText(raw);
  return CSV_FATURAMENTO_MAP[s] || null;
}

function detectImportSeparator(text){
  const firstLine=(text.split(/\r?\n/).find(l=>l.trim())||'');
  const candidates=['\t',';',','];
  return candidates
    .map(sep=>({sep,count:(firstLine.match(new RegExp(sep==='\t'?'\\t':sep,'g'))||[]).length}))
    .sort((a,b)=>b.count-a.count)[0].sep;
}

function parseImportFileRows(file, data){
  const ext=(file.name.split('.').pop()||'').toLowerCase();

  if(['xlsx','xls'].includes(ext)){
    if(typeof XLSX==='undefined') throw new Error('Biblioteca XLSX não carregada. Recarregue a página e tente novamente.');
    const wb=XLSX.read(data,{type:'array'});
    const sheetName=wb.SheetNames[0];
    if(!sheetName) throw new Error('Planilha sem abas.');
    return XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:'',raw:false});
  }

  const text=decodeBuf(data);
  const lines=text.split(/\r?\n/).filter(l=>l.trim());
  if(!lines.length) throw new Error('Arquivo vazio');
  return parseCSVRelatorio(text,detectImportSeparator(text));
}

function buildImportRows(parsed){
  const rows=parsed.filter(r=>Array.isArray(r)&&r.some(c=>String(c||'').trim()));
  if(!rows.length) throw new Error('Arquivo vazio');

  const headers=rows[0].map(h=>normalizeImportText(String(h||'').replace(/^"|"$/g,'')));

  // Find column indices. Este import aceita a planilha nova:
  // DT, HORA, DATA, FATURAMENTO, SAP, TRANSPORTADORA, STATUS, Mapa, Grade, TIPO, PESO.
  const iDT=headers.findIndex(h=>h==='DT'||h.includes('TRANSPORTE'));
  const iHora=headers.findIndex(h=>h==='HORA'||h.includes('HORA CHEGADA'));
  const iData=headers.findIndex(h=>h==='DATA'||h.includes('DATA'));
  const iSap=headers.findIndex(h=>h==='SAP'||h.includes('PORTARIA')||h.includes('N SAP')||h.includes('NR SAP')||h.includes('NO SAP'));
  const iFaturamento=headers.findIndex(h=>h==='FATURAMENTO'||h.includes('FATURAMENTO'));
  const statusIndexes=headers.map((h,i)=>({h,i})).filter(x=>x.h==='STATUS'||x.h.startsWith('STATUS.')).map(x=>x.i);
  const iStatusFinal=statusIndexes.length?statusIndexes[statusIndexes.length-1]:-1;

  if(iDT===-1) throw new Error('Coluna DT não encontrada na planilha. Verifique o formato.');

  const data=[];
  const byDT=new Map();
  for(let i=1;i<rows.length;i++){
    const cols=rows[i].map(c=>String(c||'').trim().replace(/^"|"$/g,''));
    const dt=normalizeDT(String(cols[iDT]||'').replace(/\.0+$/,'').replace(/\D/g,''));
    if(!dt) continue;

    const rawStatus=iStatusFinal!==-1?(cols[iStatusFinal]||''):'';
    const faturamento=iFaturamento!==-1?(cols[iFaturamento]||''):'';
    const mappedStatus=normalizeStatus(rawStatus) || normalizeFaturamentoStatus(faturamento);
    const hora=iHora!==-1?normalizeHora(cols[iHora]):'';
    const dataAgenda=iData!==-1?(cols[iData]||''):'';
    const fimAgendamento=combineImportDateHora(dataAgenda,hora);
    const sap=iSap!==-1?normalizeSap(cols[iSap]):'';
    const importKey=dt+'__'+(fimAgendamento?fimAgendamento.getTime():(dataAgenda||''));
    if(byDT.has(importKey)){
      const current=byDT.get(importKey);
      if(!current.rawStatus&&rawStatus) current.rawStatus=rawStatus;
      if(!current.faturamento&&faturamento) current.faturamento=faturamento;
      if(!current.mappedStatus&&mappedStatus) current.mappedStatus=mappedStatus;
      if(!current.hora&&hora) current.hora=hora;
      if(!current.dataAgenda&&dataAgenda) current.dataAgenda=dataAgenda;
      if(!current.fimAgendamento&&fimAgendamento) current.fimAgendamento=fimAgendamento;
      if(!current.sap&&sap) current.sap=sap;
    }else{
      const item={dt, rawStatus, faturamento, mappedStatus, hora, dataAgenda, fimAgendamento, sap};
      byDT.set(importKey,item);
      data.push(item);
    }
  }

  if(!data.length) throw new Error('Nenhuma DT encontrada na planilha.');
  return data;
}

function renderImportPreview(){
  const previewHtml=importCsvData.slice(0,8).map(r=>{
    const mapped=r.mappedStatus || normalizeStatus(r.rawStatus) || normalizeFaturamentoStatus(r.faturamento);
    const statusLabel=mapped||r.rawStatus||'(sem status)';
    return `<div style="display:grid;grid-template-columns:82px 58px 80px 105px 1fr;gap:6px;padding:4px 0;border-bottom:1px solid #1e293b;">
      <span style="color:#93c5fd;font-weight:700;">${r.dt}</span>
      <span style="color:#e2e8f0;">${r.hora||'—'}</span>
      <span style="color:#e2e8f0;">${r.sap||'—'}</span>
      <span style="color:#c4b5fd;">${r.faturamento||'—'}</span>
      <span style="color:${mapped?'#22c55e':'#f59e0b'};">${statusLabel}</span>
    </div>`;
  }).join('')+(importCsvData.length>8?`<div style="color:#64748b;font-size:10px;padding-top:4px;">…mais ${importCsvData.length-8} DTs</div>`:'');

  document.getElementById('import-preview-body').innerHTML=
    `<div style="display:grid;grid-template-columns:82px 58px 80px 105px 1fr;gap:6px;padding:4px 0;border-bottom:1px solid #334155;margin-bottom:4px;">
      <span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;">DT</span>
      <span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;">HORA</span>
      <span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;">SAP</span>
      <span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;">FATURAMENTO</span>
      <span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;">STATUS</span>
    </div>`+previewHtml;
  document.getElementById('import-preview').style.display='block';
}

function handleImportFile(file){
  if(!file) return;
  fileWorkflowInProgress=true;
  pauseAutoSync(180000);
  document.getElementById('import-err').style.display='none';
  document.getElementById('import-result').style.display='none';
  document.getElementById('import-drop-label').textContent='⏳ Lendo arquivo…';

  const reader=new FileReader();
  reader.onload=e=>{
    try{
      importCsvData=buildImportRows(parseImportFileRows(file,e.target.result));
      renderImportPreview();
      document.getElementById('import-drop-label').textContent=`✅ ${file.name} — ${importCsvData.length} DTs encontradas`;
      document.getElementById('import-btn').disabled=false;
      pauseAutoSync(180000);
    }catch(err){
      document.getElementById('import-err').textContent='❌ '+err.message;
      document.getElementById('import-err').style.display='block';
      document.getElementById('import-drop-label').textContent='Clique ou arraste a planilha aqui';
      releaseAutoSyncAfterSave(60000);
    }
  };
  reader.readAsArrayBuffer(file);
}

async function runImport(){
  if(!importCsvData||!importCsvData.length) return;
  fileWorkflowInProgress=true;
  pauseAutoSync(180000);

  document.getElementById('import-btn').disabled=true;
  document.getElementById('import-btn').textContent='⏳ Processando…';
  document.getElementById('import-result').style.display='none';
  document.getElementById('import-err').style.display='none';

  try{
    const now=new Date();

    // Filtra DTs pelo FIM de carregamento, nunca pelo início nem por fallback da data_ref.
    const dashHojeAmanha=tableData.filter(r=>{
      const ref=rowStatusRefDate(r);
      return ref && (sameDay(ref,today())||sameDay(ref,tomorrow()));
    });

    // Build a Set of DTs in the imported sheet
    const csvDtSet=new Set(importCsvData.map(r=>String(r.dt)));

    let updated=0, fieldUpdated=0, removed=0, skipped=0;

    // 1) Update DTs present in imported sheet (so as que estao no dash de hoje/amanha pelo fim)
    for(const csvRow of importCsvData){
      const dt=String(csvRow.dt);
      const mappedStatus=csvRow.mappedStatus || normalizeStatus(csvRow.rawStatus) || normalizeFaturamentoStatus(csvRow.faturamento);

      const dashRow=findDashRowForImportedStatus(csvRow,dashHojeAmanha);
      if(!dashRow){skipped++;continue;}

      const patch={};
      if(csvRow.hora && dashRow.hora_chegada!==csvRow.hora) patch.hora_chegada=csvRow.hora;
      if(csvRow.sap && dashRow.n_portaria!==csvRow.sap) patch.n_portaria=csvRow.sap;
      if(Object.keys(patch).length){
        await sbPatch('reporte_carga',{...patch,updated_at:new Date().toISOString()},{dt,data_ref:dashRow.data_ref});
        Object.assign(dashRow,patch);
        fieldUpdated++;
      }

      if(!mappedStatus){
        if(!Object.keys(patch).length) skipped++;
        continue;
      }
      if(dashRow.status===mappedStatus){continue;} // ja correto

      try{
        await saveStatus(dt,dashRow.data_ref,mappedStatus);
        dashRow.status=mappedStatus;
        updated++;
      }catch(e){skipped++;}
    }

    // 2) DT ausente na planilha saiu da grade ativa: registra como excluida da grade.
    const removidas=[];
    tableData=tableData.filter(r=>{
      const ref=rowStatusRefDate(r);
      const inActiveWindow=ref && (sameDay(ref,today())||sameDay(ref,tomorrow()));
      if(!inActiveWindow) return true;
      if(csvDtSet.has(String(r.dt))) return true;
      removidas.push(r);
      return false;
    });
    removed=await registrarSaidasGrade(removidas,{origem:'IMPORT_STATUS'});

    await rewriteActiveCargaSnapshot();

    // Atualiza tela
    renderRows();
    renderReporte();

    const msg=`✅ Concluído! Status atualizados: ${updated} | Hora/SAP preenchidos: ${fieldUpdated} | Removidas da grade: ${removed} | Ignoradas: ${skipped}`;
    document.getElementById('import-result').textContent=msg;
    document.getElementById('import-result').style.display='block';

  }catch(err){
    document.getElementById('import-err').textContent='❌ Erro: '+err.message;
    document.getElementById('import-err').style.display='block';
  }

  document.getElementById('import-btn').disabled=false;
  document.getElementById('import-btn').textContent='⚡ Atualizar Dashboard';
  releaseAutoSyncAfterSave();
}
