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
async function sbUpsert(table, rows){
  for(let i=0;i<rows.length;i+=30){
    let lote=rows.slice(i,i+30);
    const ignoredCols=new Set();
    while(true){
      const r=await sbFetch(`${SB_URL}/${table}`,{
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
const SC={'EXPEDIDO':'#22c55e','CARREGANDO':'#3b82f6','AG CHEGADA':'#f59e0b','NO SHOW':'#ef4444','VEICULO RECUSADO':'#dc2626','SEPARANDO':'#8b5cf6','EM FATURAMENTO':'#06b6d4','PATIO':'#64748b','DT EXCLUIDA':'#374151','FOI EMBORA':'#6b7280'};

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

let agendRows=[], dtsMescladas=[], exportMap={}, tipoOpMap={}, descDocMap={}, centroMap={}, infoAgendaMap={}, tableData=[];
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
  const allTabs=['todas','finalizadas','reporte','semgrade','dtfake','nf'];
  allTabs.forEach(t=>{
    const el=document.getElementById('tab-'+t);
    if(el) el.classList.toggle('active',t===tab);
  });
  // Hide all content areas first
  ['mat-board','twrap','materiais-wrap','reporte-wrap','semgrade-wrap','dtfake-wrap','nf-wrap'].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.style.display='none';
  });
  if(tab==='reporte'){
    document.getElementById('reporte-wrap').style.display='block';
    renderReporte();
  } else if(tab==='semgrade'){
    document.getElementById('semgrade-wrap').style.display='block';
    renderSemGrade();
  } else if(tab==='dtfake'){
    document.getElementById('dtfake-wrap').style.display='block';
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
function normalizeDiaRefRow(row){
  const diaAtual=diaRefAtualPorDataRef(row&&row.data_ref);
  if(!diaAtual) return row;
  return {...row,_stored_dia_ref:row.dia_ref,dia_ref:diaAtual};
}
function normalizeDiaRefRows(rows){
  return (rows||[]).map(normalizeDiaRefRow);
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
      const lines=decodeBuf(buf).split(/\r?\n/).filter(l=>l.trim());
      const h=lines[0].split('\t').map(s=>s.trim());
      const ci=n=>h.findIndex(x=>String(x||'').trim().toUpperCase()===String(n||'').toUpperCase());
      const iDT=ci('DT'),iLoc=ci('LOCAL'),iTransp=ci('NOME TRANSPORTADORA');
      const iAg=ci('AGENDA TRANSPORTADOR'),iFim=ci('FIM AGENDA TRANSPORTADOR');
      const iPeso=ci('PESO'),iTipo=ci('TIPO VEICULO'),iDoca=ci('DOCA');
      if(iDT===-1||iAg===-1)throw new Error('Colunas DT ou AGENDA TRANSPORTADOR não encontradas.\nColunas: '+h.join(' | '));
      agendRows=[];
      for(let i=1;i<lines.length;i++){
        const c=lines[i].split('\t');
        if(!c[iDT]?.trim())continue;
        const loc=(c[iLoc]||'').trim();
        const doca=iDoca!==-1?(c[iDoca]||'').trim():'';
        const docaNorm=doca
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g,'')
          .replace(/[\s\-]+/g,'_')
          .replace(/_+/g,'_');
        if(!loc.endsWith('1110')&&!loc.endsWith('1111'))continue;
        // Exclui docas fab_mog sem _ifnt (DOCA_1_FAB_MOG, DOCA_2_FAB_MOG etc.),
        // mas mantém variantes com sufixo _ifnt (DOCA_9_FAB_MOG_IFNT, DOCA_12_FAB_MOG_IFNT_EXTRA etc.).
        const isFabMog = docaNorm.includes('fab_mog');
        const isIfnt = docaNorm.includes('_ifnt');
        if(isFabMog && !isIfnt) continue;
        // Usa AGENDA TRANSPORTADOR como data de referência
        const agendaRaw=parseBR((c[iAg]||'').trim());
        if(!agendaRaw) continue; // sem agenda, ignora
        agendRows.push({
          DT:normalizeDT(c[iDT]),LOCAL:loc,DOCA:doca,
          TRANSPORTADORA:(c[iTransp]||'').trim(),
          AGENDA:agendaRaw,
          FIM_AGENDA:parseBR((c[iFim]||'').trim()),
          TIPO:iTipo!==-1?(c[iTipo]||'').trim():'',
          PESO:iPeso!==-1?(c[iPeso]||'').trim():'',
        });
      }
      if(!agendRows.length)throw new Error('Nenhuma linha com LOCAL 1110/1111 e DOCA _IFNT ou ARUJA encontrada.');
      hideInf();
      if(isInlineUpload){
        // Upload feito de dentro da tabela: pula passo 2/3 e vai direto para o banco
        const T=today(),AM=tomorrow();
        const dtsH=agendRows.filter(r=>r.AGENDA&&sameDay(r.AGENDA,T));
        const dtsA=agendRows.filter(r=>r.AGENDA&&sameDay(r.AGENDA,AM));
        dtsMescladas=[...dtsH,...dtsA];
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
  const dtsH=agendRows.filter(r=>r.AGENDA&&sameDay(r.AGENDA,T));
  const dtsA=agendRows.filter(r=>r.AGENDA&&sameDay(r.AGENDA,AM));
  dtsMescladas=[...dtsH,...dtsA];
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
    ? agendRows.filter(r=>r.AGENDA&&sameDay(r.AGENDA,today()))
    : agendRows.filter(r=>r.AGENDA&&sameDay(r.AGENDA,tomorrow()));
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
        const dt = strip(cols[iTransp]).replace(/\.0+$/, '').replace(/\D/g, '');
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
          const p = parseFloat(pesoRaw.replace(/\./g,'').replace(',','.')) || 0;
          pesoLiquidoMap[dt] = (pesoLiquidoMap[dt] || 0) + p;
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


async function persistImportedMaterials(){
  const dtsWithAgenda=new Map(dtsMescladas.map(r=>[String(r.DT), dKey(r.AGENDA||today())]));
  const inserts=[];
  for(const [dt,mats] of Object.entries(exportMap)){
    const dataRef=dtsWithAgenda.get(String(dt));
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
    if(pesoLiquidoMap[dt]) patch.peso_liquido=String(pesoLiquidoMap[dt].toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}));
    if(horaChegadaCSVMap[dt]) patch.hora_chegada=String(horaChegadaCSVMap[dt]);
    if(sapNumMap[dt]) patch.n_portaria=String(sapNumMap[dt]);
    if(Object.keys(patch).length===1) continue;
    for(const dataRef of refs){
      await sbPatch('reporte_carga',patch,{dt,data_ref:dataRef});
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
      status:'UPLOAD_AGENDA',
      transportadora:'AGENDA_ATIVA',
      created_at:uploadTs,
    }));
    if(logs.length){
      for(const log of logs) await tryInsertLog(log);
    }
  }catch(e){}
}

/* ═══════════════════════════════════════════════════════
   BUILD TABLE — Upsert preservando status/hora_chegada
═══════════════════════════════════════════════════════ */
async function buildTable(){
  fileWorkflowInProgress=true;
  pauseAutoSync(180000);
  showInf('Sincronizando com o banco…');
  const T=today(),AM=tomorrow();
  const kT=dKey(T), kAM=dKey(AM);
  try{
    // Busca o que já existe para PRESERVAR status e hora_chegada nas datas do upload
    const refsUpload=[...new Set(dtsMescladas.map(r=>dKey(r.AGENDA||T)))];
    const existing=await sbGet('reporte_carga',
      `data_ref=${sbIn(refsUpload)}&select=dt,data_ref,status,hora_chegada,n_portaria,tipo_operacao,descricao_documento,centro,reagendada`
    );
    const exMap={};
    (existing||[]).forEach(r=>{exMap[r.dt+'_'+r.data_ref]=r;});
    const dtsUpload=[...new Set(dtsMescladas.map(r=>normalizeDT(r.DT)).filter(Boolean))];
    const logsStatusMap={};
    if(dtsUpload.length){
      let logRows=[];
      try{
        logRows=await sbGet('reporte_logs',
          `dt=in.("${dtsUpload.join('\",\"')}")&order=created_at.desc&limit=1000`
        );
      }catch(e){
        if(!isMissingReporteLogsError(e)) throw e;
      }
      (logRows||[]).forEach(l=>{
        const dtLog=normalizeDT(l.dt);
        const st=String(l.status||'').trim();
        if(!dtLog||!st||logsStatusMap[dtLog]) return;
        if(st==='UPLOAD_AGENDA'||st==='REAGENDADA') return;
        logsStatusMap[dtLog]=st;
      });
    }


    const rows=dtsMescladas.map(dt=>{
      const ref=dKey(dt.AGENDA||T);
      const ex=exMap[dt.DT+'_'+ref]||{};
      const diaRef=sameDay(dt.AGENDA,T)?'HOJE':'AMANHÃ';
      const tipoOp=tipoOpMap[dt.DT]||(ex.tipo_operacao&&ex.tipo_operacao!==''?ex.tipo_operacao:'');
      // Preserva status e hora_chegada existentes!
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
        toneladas:String(toTonInt(dt.PESO)),
        peso_liquido: String(toTonInt(pesoLiquidoMap[dt.DT] ?? dt.PESO)),
        agenda:String(dt.AGENDA?fmtDT(dt.AGENDA,true):''),
        local_cd:String(dt.LOCAL||''),
        dia_ref:diaRef,
        data_ref:String(ref),
        tipo_operacao:String(tipoOp),
        reagendada: ex.reagendada || false,
      };
    });

    const uploadedRefs=[...new Set(rows.map(r=>r.data_ref).filter(Boolean))];
    activeRefsOverride=uploadedRefs.slice(0,2);
    // Como não existe constraint única (dt,data_ref), fazemos replace por data_ref
    for(const ref of uploadedRefs){
      await sbDelete('reporte_carga',{data_ref:ref});
    }
    await sbInsert('reporte_carga',rows);
    await saveUploadSnapshot(rows);
    const importedCount=await persistImportedMaterials();
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
      hora_chegada:horaChegadaCSVMap[dt.DT]||'',n_portaria:sapNumMap[dt.DT]||'',status:'AG CHEGADA',descricao_documento:descDocMap[dt.DT]||'',centro:centroMap[dt.DT]||'',toneladas:String(toTonInt(dt.PESO)),
      peso_liquido: String(toTonInt(pesoLiquidoMap[dt.DT] ?? dt.PESO)),
      agenda:dt.AGENDA?fmtDT(dt.AGENDA,true):'',
      dia_ref:sameDay(dt.AGENDA,T)?'HOJE':'AMANHÃ',
      data_ref:dKey(dt.AGENDA||T),
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
  const T=today(),AM=tomorrow();
  const base=[dKey(T),dKey(AM)];
  try{
    // Hoje/amanhã têm prioridade quando já existe agenda atual no banco.
    const todayRows=await sbGet('reporte_carga',`data_ref=eq.${base[0]}&select=dt&limit=1`);
    if(todayRows&&todayRows.length) return base;

    // Depois de cada upload salvamos marcadores __AGENDA__; eles evitam reabrir uma grade antiga (ex.: 05/05).
    const markers=await sbGet('reporte_logs','dt=eq.__AGENDA__&status=eq.UPLOAD_AGENDA&select=data_ref,created_at&order=created_at.desc&limit=20');
    const latestTs=markers&&markers.length?markers[0].created_at:null;
    if(latestTs){
      const active=[];
      (markers||[]).forEach(r=>{
        if(r.created_at!==latestTs) return;
        const k=String(r.data_ref||'');
        if(k&&!active.includes(k)) active.push(k);
      });
      if(active.length) return active.slice(0,2);
    }

    // Fallback legado: usa os data_ref mais recentes em logs de upload existentes.
    const lastUpload=await sbGet('reporte_logs',`status=eq.UPLOAD_AGENDA&select=data_ref,created_at&order=created_at.desc&limit=300`);
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
  if(STATUS_FINAIS.includes(row.status)) return '';
  if(row.status==='CARREGANDO') return '<span class="dt-clock" title="DT em carregamento sem SAP/Portaria. Preencher SAP agora.">🚨</span>';
  if(!row.grade_carregamento) return '';
  const grade=parseBR(row.grade_carregamento);
  if(!grade) return '';
  const agora=new Date();
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
    rows=rows.filter(r=>STATUS_FINAIS.includes(r.status));
  }else{
    rows=rows.filter(r=>!STATUS_FINAIS.includes(r.status));
  }
  if(dtSearchTerm){
    rows=rows.filter(r=>String(r.dt||'').includes(dtSearchTerm));
  }

  const tbody=document.getElementById('tbody');tbody.innerHTML='';

  // Atualiza badge da aba finalizadas
  const nFin=tableData.filter(r=>STATUS_FINAIS.includes(r.status)).length;
  document.getElementById('tab-finalizadas').textContent=`🏁 EXPEDIDAS / NO SHOW / RECUSADAS (${nFin})`;

  rows.forEach(row=>{
    const isH=row.dia_ref==='HOJE';
    const c=SC[row.status]||'#334155';
    const tr=document.createElement('tr');
    const isFinal=STATUS_FINAIS.includes(row.status);
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
    tr.innerHTML=
      `<td>${diaTag}</td>`+
      `<td><span class="td-dt" onclick="openPanel('${row.dt}','${row.data_ref}')">${row.dt}</span>${clock?` <span style="margin-left:4px;vertical-align:middle;">${clock}</span>`:''}${preFatMode?`<label style="font-size:9px;color:#a78bfa;display:flex;align-items:center;gap:3px;margin-top:2px;cursor:pointer;"><input type="checkbox" ${row.tipo_operacao==='PRÉ-FAT'?'checked':''} onchange="togglePreFat('${row.dt}','${row.data_ref}',this.checked)" style="accent-color:#a78bfa;"/>PRÉ-FAT</label>`:''}</td>` +
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

async function saveField(dt,dataRef,field,value){
  spin(true);
  try{
    await sbPatch('reporte_carga',{[field]:String(value),updated_at:new Date().toISOString()},{dt,data_ref:dataRef});
    const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef);
    if(row){row[field]=value;}
    await reloadTable();
    renderRows();
    renderReporte();
  }catch(e){showErr('Erro ao salvar: '+e.message);}
  spin(false);
}

async function tryInsertLog(row){
  try{
    await sbInsert('reporte_logs',[row]);
  }catch(e){
    const msg=String(e&&e.message||'');
    if(isMissingReporteLogsError(e)) return;
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
    const STATUS_LOG=['EXPEDIDO','NO SHOW','EM FATURAMENTO','VEICULO RECUSADO','FOI EMBORA'];
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
function isMissingReporteLogsError(err){
  const msg=String(err&&err.message||err||'').toLowerCase();
  return msg.includes("public.reporte_logs")
    || msg.includes('relation "reporte_logs" does not exist')
    || msg.includes('404')
    || msg.includes('pgrst');
}

async function loadLogs(searchDT=""){
  const list=document.getElementById('log-list');
  list.innerHTML='<div style="color:#64748b;padding:12px;">Carregando…</div>';
  try{
    const logs=await sbGet('reporte_logs','order=created_at.desc&limit=200');
    let rows=logs||[];
    const filtro=String(searchDT||'').replace(/\D/g,'');
    if(filtro) rows=rows.filter(l=>String(l.dt||'').includes(filtro));
    if(!rows.length){list.innerHTML='<div style="color:#334155;padding:12px;">Nenhum log para o filtro informado.</div>';return;}
    list.innerHTML='';
    rows.forEach(l=>{
      const sc=SC[l.status]||'#64748b';
      const div=document.createElement('div');
      div.className='log-entry';
      div.innerHTML=
        `<span style="color:#64748b">${l.created_at?new Date(l.created_at).toLocaleString('pt-BR'):'-'}</span>`+
        `<span style="color:#93c5fd;font-weight:700">${l.dt||'-'}</span>`+
        `<span style="background:${sc}22;border:1px solid ${sc}55;color:${sc};border-radius:4px;padding:2px 8px;font-weight:700;">${l.status||'-'}</span>`+
        `<span style="color:#94a3b8">${l.transportadora||'—'}</span>`;
      list.appendChild(div);
    });
  }catch(e){
    const msg=String(e&&e.message||'');
    if(isMissingReporteLogsError(e)){
      list.innerHTML='<div style="color:#94a3b8;padding:12px;">ℹ️ Log desativado: tabela <b>reporte_logs</b> não existe neste projeto Supabase. O painel de carga continua funcionando normalmente.</div>';
      return;
    }
    list.innerHTML='<div style="color:#ef4444;padding:12px;">Erro ao carregar logs: '+e.message+'<br/><small style="color:#64748b">Verifique se a tabela reporte_logs existe no Supabase.</small></div>';
  }
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
  tableData.forEach(r=>{
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
const RP_COLORS = {
  'VENDA ARUJA':'#22c55e','VENDA MOGI':'#06b6d4','PRÉ-FATURA':'#a78bfa','TRANSFERÊNCIA':'#fb923c','GRADE':'#60a5fa'
};
let rpTurnoFiltro = 'todos';
let rpMetas = JSON.parse(localStorage.getItem('rp_metas')||'{}');
let rpHoraCorte = localStorage.getItem('rp_hora_corte') || '';
let rpDataRef = localStorage.getItem('rp_data_ref') || '';
let rpLastReport = null;

const STATUS_REALIZADO = ['EM FATURAMENTO','EXPEDIDO','NO SHOW'];

function rpRefDate(row){
  // Reporte deve considerar o fim da agenda (não o início).
  // Evita contabilizar toneladas antes da janela realmente encerrar.
  return parseBR(row.fim_carregamento||'') || parseBR(row.fim_agenda||'');
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
}

function rpSetHoraCorte(v){
  rpHoraCorte=String(v||'').trim();
  try{localStorage.setItem('rp_hora_corte',rpHoraCorte);}catch(e){}
  renderReporte();
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

  const refs=[...new Set(tableData.map(r=>String(r.data_ref||'')).filter(Boolean))].sort((a,b)=>{
    const da=parseBR(a), db=parseBR(b);
    return (da?da.getTime():0)-(db?db.getTime():0);
  });
  if(!rpDataRef || (refs.length && !refs.includes(rpDataRef))) rpDataRef=refs[0]||dKey(today());

  const diaSelect=document.getElementById('rp-data-ref');
  if(diaSelect){
    diaSelect.innerHTML=refs.map(ref=>`<option value="${ref}">${ref}</option>`).join('');
    diaSelect.value=rpDataRef;
  }

  const dayRows=tableData.filter(r=>String(r.data_ref||'')===rpDataRef);
  const rows=(rpTurnoFiltro==='todos'?[...dayRows]
    :dayRows.filter(r=>rpGetTurno(r)===rpTurnoFiltro)).sort(compareFimAgendaRows);

  const infoEl=document.getElementById('rp-turno-info');

  // Inputs de meta
  const metasEl=document.getElementById('rp-metas-inputs');
  if(metasEl){
    metasEl.innerHTML='';
  }

  const corte=rpGetCorteDate();
  const corteLabel=corte.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'});
  if(infoEl) infoEl.textContent=rows.length+' DTs no filtro atual · corte até '+corteLabel;

  const planejadoTon={};
  const realizadoTon={};
  RP_TIPOS.forEach(t=>{planejadoTon[t]=0;realizadoTon[t]=0;});

  rows.forEach(r=>{
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

  rpLastReport={
    dataRef:rpDataRef,
    corteLabel,
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
  rows.forEach(r=>{
    const grade=parseBR(r.grade_carregamento||'');
    if(!grade || !sameDay(grade,parseBR(rpDataRef||'')||today())) return;
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
    rows.forEach(r=>{ sCounts[r.status]=(sCounts[r.status]||0)+1; });
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

    rows.forEach(r=>{
      const grade=parseBR(r.grade_carregamento||'');
      if(!grade || !sameDay(grade,hoje0)) return;
      const turno=rpTurnoFromDate(grade);
      if(!turno) return;
      const ton=rpParseToneladas(r);
      byTurno[turno]+=ton;

      const paletKey=normalizePaletizacaoLabel(r.paletizacao);
      paletByTurno[turno][paletKey]++;

      if(turno==='T3' && grade){
        if(grade>=t3DiaIni && grade<=t3DiaFim) t3DiaProd+=ton;         // produtividade do dia
        if(grade>=t3TurnoIni && grade<=t3TurnoFim) t3TurnoProd+=ton;   // produtividade do turno
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
    .filter(r=>!STATUS_FINAIS.includes(r.status))
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
  const iSap=headers.findIndex(h=>h==='SAP'||h.includes('PORTARIA')||h.includes('N SAP')||h.includes('NR SAP')||h.includes('NO SAP'));
  const iFaturamento=headers.findIndex(h=>h==='FATURAMENTO'||h.includes('FATURAMENTO'));
  const statusIndexes=headers.map((h,i)=>({h,i})).filter(x=>x.h==='STATUS'||x.h.startsWith('STATUS.')).map(x=>x.i);
  const iStatusFinal=statusIndexes.length?statusIndexes[statusIndexes.length-1]:-1;

  if(iDT===-1) throw new Error('Coluna DT não encontrada na planilha. Verifique o formato.');

  const data=[];
  const seen=new Set();
  for(let i=1;i<rows.length;i++){
    const cols=rows[i].map(c=>String(c||'').trim().replace(/^"|"$/g,''));
    const dt=normalizeDT(String(cols[iDT]||'').replace(/\.0+$/,'').replace(/\D/g,''));
    if(!dt||seen.has(dt)) continue;
    seen.add(dt);

    const rawStatus=iStatusFinal!==-1?(cols[iStatusFinal]||''):'';
    const faturamento=iFaturamento!==-1?(cols[iFaturamento]||''):'';
    const mappedStatus=normalizeStatus(rawStatus) || normalizeFaturamentoStatus(faturamento);
    const hora=iHora!==-1?normalizeHora(cols[iHora]):'';
    const sap=iSap!==-1?normalizeSap(cols[iSap]):'';
    data.push({dt, rawStatus, faturamento, mappedStatus, hora, sap});
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
    const nowMinutes=now.getHours()*60+now.getMinutes();

    // Filtrar só DTs de HOJE e AMANHÃ do dashboard
    const dashHojeAmanha=tableData.filter(r=>r.dia_ref==='HOJE'||r.dia_ref==='AMANHÃ');

    // Build a Set of DTs in the imported sheet
    const csvDtSet=new Set(importCsvData.map(r=>String(r.dt)));

    let updated=0, fieldUpdated=0, expedited=0, skipped=0;

    // 1) Update DTs present in imported sheet (só as que estão no dash de hoje/amanhã)
    for(const csvRow of importCsvData){
      const dt=String(csvRow.dt);
      const mappedStatus=csvRow.mappedStatus || normalizeStatus(csvRow.rawStatus) || normalizeFaturamentoStatus(csvRow.faturamento);

      const dashRow=dashHojeAmanha.find(r=>String(r.dt)===dt);
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
      if(dashRow.status===mappedStatus){continue;} // já correto

      try{
        await saveStatus(dt,dashRow.data_ref,mappedStatus);
        dashRow.status=mappedStatus;
        updated++;
      }catch(e){skipped++;}
    }

    // 2) Marcar como EXPEDIDO: DTs de HOJE no dash, grade ≤ agora, ausentes da planilha
    const finalSet=new Set(STATUS_FINAIS);
    for(const dashRow of dashHojeAmanha){
      if(dashRow.dia_ref!=='HOJE') continue;            // só HOJE
      const dt=String(dashRow.dt);
      if(csvDtSet.has(dt)) continue;                    // está na planilha, pular
      if(finalSet.has(dashRow.status)) continue;         // já finalizada

      // Checa se a grade já passou
      const gradeStr=dashRow.grade_carregamento||'';
      const gradeMatch=gradeStr.match(/(\d{1,2}):(\d{2})/);
      if(gradeMatch){
        const gradeMinutes=parseInt(gradeMatch[1])*60+parseInt(gradeMatch[2]);
        if(gradeMinutes>nowMinutes) continue; // grade ainda no futuro
      } else {
        if(!dashRow.hora_chegada) continue;
        const [hh,mm]=(dashRow.hora_chegada||'').split(':').map(Number);
        if(isNaN(hh)||hh*60+mm>nowMinutes) continue;
      }

      try{
        await saveStatus(dt,dashRow.data_ref,'EXPEDIDO');
        dashRow.status='EXPEDIDO';
        expedited++;
      }catch(e){}
    }

    // Atualiza tela
    renderRows();
    renderReporte();

    const msg=`✅ Concluído! Status atualizados: ${updated} | Hora/SAP preenchidos: ${fieldUpdated} | Expedidas automáticas: ${expedited} | Ignoradas: ${skipped}`;
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
