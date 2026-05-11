/* ═══════════════════════════════════════════════════════
   SUPABASE REST
═══════════════════════════════════════════════════════ */
const SB_URL = 'https://pwjatxqtkvwcmzmjjvbi.supabase.co/rest/v1';
// ⚠️ TROQUE PELA anon key do Supabase: Dashboard → Settings → API → anon public (começa com eyJ...)
const SB_KEY = 'sb_publishable_bUPTDkrOzc0_I3xwNw15aA_lk76gg4w';
const HDR = {'Content-Type':'application/json','apikey':SB_KEY,'Authorization':'Bearer '+SB_KEY};

async function sbGet(table, qs=''){
  if(!SB_KEY || SB_KEY.includes('COLE_SUA_ANON_KEY_AQUI')){
    throw new Error('Chave Supabase não configurada em cco_script.js (SB_KEY).');
  }
  const r=await fetch(`${SB_URL}/${table}?${qs}`,{headers:HDR});
  if(!r.ok) throw new Error(`GET ${table}: `+await r.text());
  return r.json();
}
async function sbUpsert(table, rows){
  for(let i=0;i<rows.length;i+=30){
    const r=await fetch(`${SB_URL}/${table}`,{
      method:'POST',
      headers:{...HDR,'Prefer':'resolution=merge-duplicates,return=minimal'},
      body:JSON.stringify(rows.slice(i,i+30))
    });
    if(!r.ok) throw new Error(`UPSERT ${table} lote ${i/30+1}: `+await r.text());
  }
}
async function sbPatch(table, data, filters){
  const q=Object.entries(filters).map(([k,v])=>`${k}=eq.${v.includes('/')?'"'+v+'"':encodeURIComponent(v)}`).join('&');
  const r=await fetch(`${SB_URL}/${table}?${q}`,{method:'PATCH',headers:HDR,body:JSON.stringify(data)});
  if(!r.ok) throw new Error(`PATCH ${table}: `+await r.text());
}
async function sbDelete(table, filters){
  const q=Object.entries(filters).map(([k,v])=>{
    const vs=String(v); return `${k}=eq.${vs.includes('/')?'"'+vs+'"':encodeURIComponent(vs)}`;
  }).join('&');
  const r=await fetch(`${SB_URL}/${table}?${q}`,{method:'DELETE',headers:HDR});
  if(!r.ok) throw new Error(`DELETE ${table}: `+await r.text());
}
async function sbInsert(table, rows){
  const r=await fetch(`${SB_URL}/${table}`,{method:'POST',headers:HDR,body:JSON.stringify(rows)});
  if(!r.ok) throw new Error(`INSERT ${table}: `+await r.text());
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

let agendRows=[], dtsMescladas=[], exportMap={}, tipoOpMap={}, descDocMap={}, tableData=[];
let remessaMap={}, pesoLiquidoMap={}, horaChegadaCSVMap={}, sapNumMap={};
let preFatMode=false; // Ctrl+Ç ativa checkboxes de Pré-Fat nas DTs
let panelDT=null;
let currentTableTab='todas';
let reagendDT=null;
let dtSearchTerm='';
let activeRefsOverride=[];

/* ═══════════════════════════════════════════════════════
   INIT
═══════════════════════════════════════════════════════ */
// Ctrl+Ç — ativa/desativa modo Pré-Fat
document.addEventListener('keydown', e=>{
  if(e.ctrlKey && (e.key==='ç'||e.key==='Ç'||e.keyCode===231||e.keyCode===199)){
    e.preventDefault();
    preFatMode=!preFatMode;
    document.body.classList.toggle('prefat-mode', preFatMode);
    showOk(preFatMode?'✅ Modo PRÉ-FAT ativado — marque as DTs pré-faturadas':'❌ Modo PRÉ-FAT desativado');
    renderRows();
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
  } catch(e) {
    syncTxt.textContent = 'ERRO';
    if(syncDot){ syncDot.style.background='#ef4444'; syncDot.style.animation='none'; }
    showErr('Falha ao conectar: ' + e.message);
  }
  // Ticker para atualizar emojis de relógio a cada minuto
  setInterval(()=>{ if(tableData.length) renderRows(); }, 60000);
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
      .sort((a,b)=>parseAgendaDateTime(a)-parseAgendaDateTime(b))
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

      // Soma total de qtde de todos os itens da DT
      const qtdeTotalDT = itens.reduce((sum, i) => {
        const n = parseFloat(String(i.qtde||'0').replace(/\./g,'').replace(',','.')) || 0;
        return sum + n;
      }, 0);

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
      html+=`<div style="font-size:12px;font-weight:700;color:#a7f3d0;">Total: ${Math.round(qtdeTotalDT)} un.</div>`;
      html+=`</div>`;
      html+=`</div>`;

      // Tabela agrupada por remessa
      html+=`<div style="padding:0 14px 10px;">`;
      remessas.forEach(remessa=>{
        const linhas=byRemessa[remessa];
        html+=`<div style="margin-bottom:6px;">`;
        html+=`<div style="display:grid;grid-template-columns:1fr 120px 100px;gap:4px;padding:4px 0;border-bottom:1px solid #334155;margin-bottom:2px;">`;
        html+=`<span style="font-size:9px;letter-spacing:1.5px;color:#3b82f6;font-weight:700;">REMESSA ${remessa}</span>`;
        html+=`<span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;">MATERIAL</span>`;
        html+=`<span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;text-align:right;">QTDE</span>`;
        html+=`</div>`;
        linhas.forEach(({material,qtde},idx)=>{
          html+=`<div style="display:grid;grid-template-columns:1fr 120px 100px;gap:4px;padding:2px 0;border-bottom:1px solid #1e293b;font-size:12px;">`;
          html+=`<span style="color:#475569;font-size:10px;">${idx===0?'':'—'}</span>`;
          html+=`<span style="color:#cbd5e1;">${material}</span>`;
          html+=`<span style="text-align:right;color:#a7f3d0;font-weight:700;">${qtde||'—'}</span>`;
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
  const m=s.match(/(\d{2})\/(\d{2})\/(\d{4})(?:[T\s](\d{2}):(\d{2}))?/);
  return m?new Date(+m[3],+m[2]-1,+m[1],+(m[4]||0),+(m[5]||0)):null;
}

function parseAgendaDateTime(row){
  const grade=parseBR(row.grade_carregamento||'');
  if(grade) return grade;
  const agenda=parseBR(row.agenda||'');
  if(agenda) return agenda;
  return new Date(9999,0,1);
}
function normalizeDT(raw){
  const v=String(raw||'').trim();
  if(!v) return '';
  if(/^\d+$/.test(v)) return v.replace(/^0+(?=\d)/,'');
  return v;
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
  showInf('Lendo agendamento…');
  document.getElementById('dz1l').innerHTML='<span class="dz-ok">✓ '+file.name+'</span>';
  file.arrayBuffer().then(buf=>{
    try{
      // Tenta importar materiais automaticamente se o arquivo tiver as colunas certas (XLSX)
      exportMap={};remessaMap={};pesoLiquidoMap={};
      const matCnt=tryImportMaterialsFromAgend(buf);
      if(matCnt>0) showOk('✨ '+matCnt+' linhas de materiais importadas automaticamente do arquivo de agendamento!');
      const lines=decodeBuf(buf).split(/\r?\n/).filter(l=>l.trim());
      const h=lines[0].split('\t').map(s=>s.trim());
      const ci=n=>h.indexOf(n);
      const iDT=ci('DT'),iLoc=ci('LOCAL'),iTransp=ci('NOME TRANSPORTADORA');
      const iAg=ci('AGENDA TRANSPORTADOR'),iFim=ci('FIM AGENDA TRANSPORTADOR');
      const iPeso=ci('PESO'),iTipo=ci('TIPO VEICULO');
      if(iDT===-1||iAg===-1)throw new Error('Colunas DT ou AGENDA TRANSPORTADOR não encontradas.\nColunas: '+h.join(' | '));
      agendRows=[];
      for(let i=1;i<lines.length;i++){
        const c=lines[i].split('\t');
        if(!c[iDT]?.trim())continue;
        const loc=(c[iLoc]||'').trim();
        if(!loc.endsWith('1110')&&!loc.endsWith('1111'))continue;
        agendRows.push({
          DT:normalizeDT(c[iDT]),LOCAL:loc,
          TRANSPORTADORA:(c[iTransp]||'').trim(),
          AGENDA:parseBR((c[iAg]||'').trim()),
          FIM_AGENDA:parseBR((c[iFim]||'').trim()),
          TIPO:iTipo!==-1?(c[iTipo]||'').trim():'',
          PESO:iPeso!==-1?(c[iPeso]||'').trim():'',
        });
      }
      if(!agendRows.length)throw new Error('Nenhuma linha com LOCAL 1110/1111 encontrada.');
      hideInf();
      renderStep2();
    }catch(e){hideInf();showErr(e.message);}
  });
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
  showInf('Lendo CSV do Relatório de Expedição…');
  document.getElementById('dz2l').innerHTML = '<span class="dz-ok">✓ ' + file.name + '</span>';
  const reader = new FileReader();
  reader.onload = function(e) {
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

      if (iTransp === -1) throw new Error('Coluna Nº transporte não encontrada.\nColunas: ' + headers.join(' | '));

      exportMap = {}; tipoOpMap = {}; descDocMap = {};
      remessaMap = {}; pesoLiquidoMap = {}; horaChegadaCSVMap = {}; sapNumMap = {};
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

        // descricao_documento e tipo operacao — centro 1111=ARUJA, 1110=MOGI
        if (!descDocMap[dt]) descDocMap[dt] = descDoc;
        if (!tipoOpMap[dt]) {
          const raw = descDoc.toUpperCase();
          const centroRaw = iCentro !== -1 ? String(cols[iCentro]||'').trim().replace(/\.0+$/,'') : '';
          const isMogi  = centroRaw === '1110';
          const isAruja = centroRaw === '1111' || centroRaw === '';
          const isTransf = raw.includes('TRANSFER') || raw.includes('FILIAL') || raw.includes('ABAST') || raw.includes('TNF');
          if (isTransf) {
            tipoOpMap[dt] = 'TRANSFERÊNCIA';
          } else {
            tipoOpMap[dt] = isMogi ? 'VENDA MOGI' : 'VENDA ARUJA';
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
      hideInf();
      showOk(`Relatório OK — ${totalDTs} transportes · ${linhasOk} linhas`);
      buildTable();
    } catch(err) {
      hideInf();
      showErr('Erro ao ler CSV: ' + err.message);
      console.error(err);
    }
  };
  reader.onerror = () => { hideInf(); showErr('Erro ao abrir o arquivo.'); };
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
      inserts.push({dt:String(dt),data_ref:String(dataRef),material:String(m.material),quantidade:String(m.quantidade||''),observacao:'',ordem:i});
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

async function saveUploadSnapshot(rows){
  try{
    const logs=rows.map(r=>({
      dt:String(r.dt||''),
      data_ref:String(r.data_ref||''),
      status:'UPLOAD_AGENDA',
      transportadora:String(r.transportadora||''),
      created_at:new Date().toISOString(),
    })).filter(r=>r.dt&&r.data_ref);
    if(logs.length) await sbInsert('reporte_logs',logs);
  }catch(e){}
}

/* ═══════════════════════════════════════════════════════
   BUILD TABLE — Upsert preservando status/hora_chegada
═══════════════════════════════════════════════════════ */
async function buildTable(){
  showInf('Sincronizando com o banco…');
  const T=today(),AM=tomorrow();
  const kT=dKey(T), kAM=dKey(AM);
  try{
    // Busca o que já existe para PRESERVAR status e hora_chegada nas datas do upload
    const refsUpload=[...new Set(dtsMescladas.map(r=>dKey(r.AGENDA||T)))];
    const existing=await sbGet('reporte_carga',
      `data_ref=in.("${refsUpload.join('\",\"')}")&select=dt,data_ref,status,hora_chegada,n_portaria,tipo_operacao,descricao_documento,reagendada`
    );
    const exMap={};
    (existing||[]).forEach(r=>{exMap[r.dt+'_'+r.data_ref]=r;});
    const dtsUpload=[...new Set(dtsMescladas.map(r=>normalizeDT(r.DT)).filter(Boolean))];
    const logsStatusMap={};
    if(dtsUpload.length){
      const logRows=await sbGet('reporte_logs',
        `dt=in.("${dtsUpload.join('\",\"')}")&order=created_at.desc&limit=1000`
      );
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
        toneladas:String(dt.PESO||''),
        peso_liquido: pesoLiquidoMap[dt.DT] ? String(pesoLiquidoMap[dt.DT].toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})) : String(dt.PESO||''),
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
  }catch(e){
    hideInf();
    showErr('Erro banco (tabela ainda aparece): '+e.message);
    tableData=dtsMescladas.map(dt=>({
      dt:dt.DT,transportadora:dt.TRANSPORTADORA,
      grade_carregamento:dt.AGENDA?fmtDT(dt.AGENDA,true):'',
      fim_carregamento:dt.FIM_AGENDA?fmtDT(dt.FIM_AGENDA,true):'',
      hora_chegada:horaChegadaCSVMap[dt.DT]||'',n_portaria:sapNumMap[dt.DT]||'',status:'AG CHEGADA',descricao_documento:descDocMap[dt.DT]||'',toneladas:dt.PESO,
      peso_liquido: pesoLiquidoMap[dt.DT] ? String(pesoLiquidoMap[dt.DT].toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})) : String(dt.PESO||''),
      agenda:dt.AGENDA?fmtDT(dt.AGENDA,true):'',
      dia_ref:sameDay(dt.AGENDA,T)?'HOJE':'AMANHÃ',
      data_ref:dKey(dt.AGENDA||T),
      tipo_operacao:tipoOpMap[dt.DT]||'',
      reagendada:false,
    }));
    renderRows();
    setStep(4);
  }
}


async function getActiveRefs(){
  if(activeRefsOverride.length) return activeRefsOverride;
  const T=today(),AM=tomorrow();
  const base=[dKey(T),dKey(AM)];
  try{
    // Sempre prioriza hoje/amanhã se existir no banco
    const todayRows=await sbGet('reporte_carga',`data_ref=eq.${base[0]}&select=dt&limit=1`);
    if(todayRows&&todayRows.length) return base;

    // Regra principal: usar o último upload de agenda registrado em log
    const lastUpload=await sbGet('reporte_logs',`status=eq.UPLOAD_AGENDA&select=data_ref,created_at&order=created_at.desc&limit=300`);
    const byUpload=[];
    (lastUpload||[]).forEach(r=>{
      const k=String(r.data_ref||'');
      if(k && !byUpload.includes(k)) byUpload.push(k);
    });
    if(byUpload.length) return byUpload.slice(0,2);

    // Fallback robusto: ordena por data_ref real (dd/mm/aaaa), não por texto
    const rec=await sbGet('reporte_carga','select=data_ref,updated_at&order=updated_at.desc&limit=1000');
    const uniq=[...new Set((rec||[]).map(r=>String(r.data_ref||'')).filter(Boolean))];
    const toDate=(k)=>{const m=String(k).match(/(\d{2})\/(\d{2})\/(\d{4})/);return m?new Date(+m[3],+m[2]-1,+m[1]):new Date(0);};
    uniq.sort((a,b)=>toDate(b)-toDate(a));
    return uniq.slice(0,2);
  }catch(e){return base;}
}

async function reloadTable(){
  try{
    const refs=await getActiveRefs();
    if(!refs.length){tableData=[];renderRows();return;}
    const data=await sbGet('reporte_carga',
      `data_ref=in.("${refs.join('\",\"')}")&order=dia_ref.asc,agenda.asc`
    );
    tableData=(data||[]).sort((a,b)=>{
      return parseAgendaDateTime(a)-parseAgendaDateTime(b);
    });
    renderRows();
  }catch(e){showErr('Erro ao carregar tabela: '+e.message);}
}

/* ═══════════════════════════════════════════════════════
   RELÓGIO — verifica se deve mostrar emoji
   Regra: se hoje e status não é final e grade definida,
   mostra 🕐 se falta ≤1h para o início da grade
═══════════════════════════════════════════════════════ */
function clockEmoji(row){
  if(row.dia_ref!=='HOJE') return '';
  if(STATUS_FINAIS.includes(row.status)) return '';
  if(!row.grade_carregamento) return '';
  const grade=parseBR(row.grade_carregamento);
  if(!grade) return '';
  const agora=new Date();
  const diffMin=(grade - agora)/60000; // positivo = grade ainda no futuro
  // Carreta deve chegar 1h antes = diffMin entre 0 e 60 => alerta
  if(diffMin>=0 && diffMin<=60) return '<span class="dt-clock" title="Carreta deve chegar até '+new Date(grade-3600000).toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'})+'">🕐</span>';
  if(diffMin<0 && diffMin>-120) return '<span style="font-size:11px;opacity:.6" title="Grade iniciada">⏰</span>';
  return '';
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
    const q='data_ref=in.("'+refs.join('","')+'")&order=dt.asc,ordem.asc';
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
    const ordem=[...tableData].sort((a,b)=>parseAgendaDateTime(a)-parseAgendaDateTime(b)).map(r=>String(r.dt));
    const uniq=[]; ordem.forEach(dt=>{if(!uniq.includes(dt)) uniq.push(dt);});
    let html='';
    uniq.forEach(dt=>{
      const itens=byDT[dt]; if(!itens) return;
      const pairs=Object.entries(itens);
      const total=pairs.reduce((s,[,v])=>s+v,0);
      html+='<div class="mat-group"><div class="mat-group-h"><span>DT '+dt+'</span><span>'+pairs.length+' materiais · Qtd total '+Math.round(total)+'</span></div><div class="mat-group-b">';
      pairs.forEach(([mat,q])=>{html+='<div class="mat-item"><span>'+mat+'</span><span style="text-align:right;color:#a7f3d0;font-weight:700;">'+Math.round(q)+'</span></div>';});
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
  if(board) board.style.display='none';
  if(twrap) twrap.style.display='block';
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
  let rows=[...tableData].sort((a,b)=>parseAgendaDateTime(a)-parseAgendaDateTime(b));
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
    const diaTag=isH?'<span class="tag-hoje">HOJE</span>':'<span class="tag-amanha">AMANHÃ</span>';
    const clock=clockEmoji(row);
    // Botão reagendamento — aparece nas finalizadas
    const reagBtn=isFinal
      ? `<button class="bg" style="font-size:10px;padding:3px 9px;border-color:#f59e0b55;color:#f59e0b;" onclick="openReagend('${row.dt}','${row.data_ref}')">🔄 Reagendar</button>`
      : '';
    tr.innerHTML=
      `<td>${diaTag}</td>`+
      `<td><span class="td-dt" onclick="openPanel('${row.dt}','${row.data_ref}')">${row.dt}</span>${preFatMode?`<label style="font-size:9px;color:#a78bfa;display:flex;align-items:center;gap:3px;margin-top:2px;cursor:pointer;"><input type="checkbox" ${row.tipo_operacao==='PRÉ-FAT'?'checked':''} onchange="togglePreFat('${row.dt}','${row.data_ref}',this.checked)" style="accent-color:#a78bfa;"/>PRÉ-FAT</label>`:''}` +
      `<td class="td-transp">${row.transportadora||'—'}</td>`+
      `<td class="td-time">${row.grade_carregamento||'—'}</td>`+
      `<td class="td-time">${row.fim_carregamento||'—'}</td>`+
      `<td><input type="time" value="${row.hora_chegada||''}" onchange="saveField('${row.dt}','${row.data_ref}','hora_chegada',this.value)"/></td>`+
      `<td style="text-align:center;" data-dt="${row.dt}" data-ref="${row.data_ref}" data-sap="${row.n_portaria||''}" onclick="promptPortaria(this.dataset.dt,this.dataset.ref,this.dataset.sap)" style="cursor:pointer;text-align:center;">${
  row.n_portaria
    ? '<span title="SAP: '+row.n_portaria+'" style="color:#4ade80;font-size:15px;">✅</span><div style="font-size:9px;color:#4ade80;letter-spacing:.5px;">'+row.n_portaria+'</div>'
    : (clockEmoji(row)||'<span style="color:#334155;font-size:11px;">—</span>')+'<div style="font-size:9px;color:#475569;margin-top:1px;">SAP</div>'
}</td>`+
      `<td><div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;"><select class="ss" style="background:${c}22;border:1.5px solid ${c}99;color:${c}" onchange="saveStatus('${row.dt}','${row.data_ref}',this.value)">`+
        STATUS_OPTIONS.map(s=>`<option value="${s}"${s===row.status?' selected':''}>${s}</option>`).join('')+
      `</select></div></td>`+
      `<td class="td-sm"><div style="display:flex;align-items:center;gap:5px;flex-wrap:wrap;">${opTag?opTag+' ':''}<span>${row.descricao_documento||'—'}</span></div></td>`+
      `<td class="td-sm">${row.peso_liquido||row.toneladas||'—'}</td>`;
    tbody.appendChild(tr);
  });

  if(!rows.length){
    const tr=document.createElement('tr');
    tr.innerHTML=`<td colspan="9" style="text-align:center;padding:30px;color:#334155;font-size:13px;">
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
    if(row){row[field]=value;renderRows();}
  }catch(e){showErr('Erro ao salvar: '+e.message);}
  spin(false);
}

/* SALVAR STATUS + gravar no log */
async function saveStatus(dt,dataRef,value){
  spin(true);
  try{
    await sbPatch('reporte_carga',{status:String(value),updated_at:new Date().toISOString()},{dt,data_ref:dataRef});
    const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef);
    // Grava log só para status relevantes (evita poluição)
    const STATUS_LOG=['EXPEDIDO','NO SHOW','EM FATURAMENTO','VEICULO RECUSADO','FOI EMBORA'];
    if(STATUS_LOG.includes(value)){
      await sbInsert('reporte_logs',[{
        dt:String(dt),
        data_ref:String(dataRef),
        status:String(value),
        transportadora:row?String(row.transportadora||''):'',
        created_at:new Date().toISOString(),
      }]);
    }
    if(row){row.status=value;renderRows();}
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
    await sbInsert('reporte_logs',[{
      dt:String(dt),
      data_ref:String(dataRef),
      status:'REAGENDADA',
      transportadora:obs||'Reagendamento via sistema',
      created_at:new Date().toISOString(),
    }]);
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
      `<input class="mat-input" placeholder="Material" value="${m.material||''}" data-f="material"/>`+
      `<input class="mat-input" placeholder="Qtd"      value="${m.quantidade||''}" data-f="quantidade" oninput="updatePanelTotal()"/>`+
      `<button class="mat-del" onclick="this.parentElement.remove();updatePanelTotal();">🗑</button>`;
    list.appendChild(row);
  });
  updatePanelTotal();
}

function addMatRow(){
  const row=document.createElement('div');row.className='mat-row';
  row.innerHTML=
    `<input class="mat-input" placeholder="Material" data-f="material"/>`+
    `<input class="mat-input" placeholder="Qtd"      data-f="quantidade" oninput="updatePanelTotal()"/>`+
    `<button class="mat-del" onclick="this.parentElement.remove();updatePanelTotal();">🗑</button>`;
  document.getElementById('mat-list').appendChild(row);
  row.querySelector('input').focus();
  updatePanelTotal();
}

function updatePanelTotal(){
  const qs=[...document.querySelectorAll('#mat-list [data-f=quantidade]')];
  const total=qs.reduce((s,i)=>s+(parseFloat(String(i.value||'0').replace(/\./g,'').replace(',','.'))||0),0);
  const el=document.getElementById('mat-total');
  if(el) el.textContent='Total: '+Math.round(total);
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
  const total=mats.reduce((s,m)=>s+(parseFloat(String(m.qtd||'0').replace(/\./g,'').replace(',','.'))||0),0);
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
        Total de volumes: <b>${Math.round(total)}</b>
      </div>
    </div>
    <table class="print-mat-table">
      <thead><tr><th>#</th><th>MATERIAL</th><th>QUANTIDADE</th></tr></thead>
      <tbody>${mats.map((m,i)=>`<tr><td>${i+1}</td><td>${m.mat}</td><td style="text-align:right;">${m.qtd}</td></tr>`).join('')}</tbody>
      <tfoot><tr><td colspan="2" style="text-align:right;font-weight:800;padding:6px 10px;">TOTAL</td><td style="text-align:right;font-weight:800;padding:6px 10px;">${Math.round(total)}</td></tr></tfoot>
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
  const rows=[...document.getElementById('mat-list').querySelectorAll('.mat-row')];
  const mats=rows.map((r,i)=>({
    dt,data_ref:dataRef,
    material:r.querySelector('[data-f=material]').value.trim(),
    quantidade:r.querySelector('[data-f=quantidade]').value.trim(),
    observacao:document.getElementById('pobs').value.trim(),
    ordem:i
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
      `data_ref=in.("${refs.join('\",\"')}")&order=dia_ref.asc,agenda.asc`
    );
    if(data&&data.length){
      tableData=data.sort((a,b)=>{
        const ga=parseBR(a.grade_carregamento)||new Date(9999,0);
        const gb=parseBR(b.grade_carregamento)||new Date(9999,0);
        return ga-gb;
      });
      // Popula tipoOpMap a partir do banco para o painel funcionar sem ZLES002
      tableData.forEach(r=>{ if(r.dt&&r.tipo_operacao) tipoOpMap[r.dt]=r.tipo_operacao; });
      renderRows();
      setStep(4);
      showOk('Tabela carregada do banco com a agenda mais recente salva.');
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
function exportCSV(){
  const cols=['DIA','DT','TRANSPORTADORA','GRADE','FIM','HORA CHEGADA','N° PORTARIA','STATUS','DESC. DOCUMENTO','PESO LÍQUIDO','TIPO OPERAÇÃO'];
  const rows=tableData.map(r=>[r.dia_ref,r.dt,r.transportadora,r.grade_carregamento,r.fim_carregamento,r.hora_chegada,r.n_portaria||'',r.status,r.descricao_documento||'',r.peso_liquido||r.toneladas||'',r.tipo_operacao||'']);
  const csv=[cols,...rows].map(r=>r.map(v=>'"'+String(v||'').replace(/"/g,'""')+'"').join(';')).join('\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));
  a.download='reporte_'+new Date().toLocaleDateString('pt-BR').replace(/\//g,'-')+'.csv';
  a.click();
}

/* ═══════════════════════════════════════════════════════
   REPORTE DE STATUS
═══════════════════════════════════════════════════════ */
const RP_TIPOS = ['VENDA ARUJA','VENDA MOGI','PRÉ-FATURA','TRANSFERÊNCIA','GRADE'];
const RP_COLORS = {
  'VENDA ARUJA':'#22c55e','VENDA MOGI':'#06b6d4','PRÉ-FATURA':'#a78bfa','TRANSFERÊNCIA':'#fb923c','GRADE':'#60a5fa'
};
let rpTurnoFiltro = 'todos';
let rpMetas = JSON.parse(localStorage.getItem('rp_metas')||'{}');

function rpGetTurno(row){
  // Classifica pelo horário do FIM da agenda
  const grade = parseBR(row.grade_carregamento||'') || parseBR(row.agenda||'');
  if(!grade) return null;
  const h = grade.getHours();
  if(h>=7 && h<=14) return 'T1';
  if(h>=15 && h<=22) return 'T2';
  return 'T3'; // 23-06
}

function rpNormalizeTipo(row){
  const op=(row.tipo_operacao||'').trim().toUpperCase();
  // Tipos exatos que vêm do tipoOpMap (ZLES002)
  if(op==='VENDA ARUJA')         return 'VENDA ARUJA';
  if(op==='VENDA MOGI')          return 'VENDA MOGI';
  if(op==='TRANSFERÊNCIA ARUJA'||op==='TRANSFERÊNCIA MOGI'||op==='TRANSFERÊNCIA') return 'TRANSFERÊNCIA';
  if(op==='PRÉ-FAT'||op==='PRÉ-FATURA'||op.includes('PRÉ')||op.includes('PRE')) return 'PRÉ-FATURA';
  if(op==='GRADE'||op.includes('GRADE')) return 'GRADE';
  // Fallback por texto
  if(op.includes('TRANSFER')) return 'TRANSFERÊNCIA';
  if(op.includes('MOGI'))     return 'VENDA MOGI';
  return 'VENDA ARUJA';
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
  return Number.isFinite(n)?n:0;
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

function renderReporte(){
  const rows=rpTurnoFiltro==='todos'?[...tableData]
    :tableData.filter(r=>rpGetTurno(r)===rpTurnoFiltro);

  const infoEl=document.getElementById('rp-turno-info');
  if(infoEl) infoEl.textContent=rows.length+' DTs no filtro atual';

  // Inputs de meta
  const metasEl=document.getElementById('rp-metas-inputs');
  if(metasEl){
    metasEl.innerHTML=RP_TIPOS.map(tipo=>`
      <div>
        <label style="font-size:9px;letter-spacing:1.5px;color:${RP_COLORS[tipo]||'#64748b'};font-weight:700;display:block;margin-bottom:4px;">${tipo}</label>
        <input class="rp-meta-input" type="number" min="0" value="${rpMetas[tipo]||0}"
          oninput="rpSaveMeta('${tipo}',this.value)" title="Meta do dia para ${tipo}"/>
      </div>
    `).join('');
  }

    const agora=new Date();
  const inicioDia=new Date(agora); inicioDia.setHours(0,0,0,0);

  const realizadoTon={};
  RP_TIPOS.forEach(t=>{realizadoTon[t]=0;});

  const GRADE_STATUS_REALIZADO=['CARREGANDO','EM FATURAMENTO','EXPEDIDO'];

  rows.forEach(r=>{
    const tipo=rpNormalizeTipo(r);
    const ton=rpParseToneladas(r);

    // GRADE: planejado automático por fim_carregamento no período
    const fim=parseBR(r.fim_carregamento||'');
    if(rpWithinWindow(fim,inicioDia,agora)){
      rpMetas['GRADE']=(rpMetas['GRADE']||0); // preserva campo mas planejado visual é dinâmico
      if(GRADE_STATUS_REALIZADO.includes(r.status)) realizadoTon['GRADE']+=ton;
    }

    if(tipo==='VENDA ARUJA'){
      const centro=String(r.centro||'').trim();
      const desc=String(r.descricao_documento||'').toUpperCase();
      const ref=parseBR(r.grade_carregamento||'')||parseBR(r.agenda||'')||parseBR(r.fim_carregamento||'');
      if(centro==='1111' && desc.includes('VENDA') && rpWithinWindow(ref,inicioDia,agora)) realizadoTon['VENDA ARUJA']+=ton;
      return;
    }
    if(tipo==='VENDA MOGI'){
      const centro=String(r.centro||'').trim();
      const desc=String(r.descricao_documento||'').toUpperCase();
      const ref=parseBR(r.grade_carregamento||'')||parseBR(r.agenda||'')||parseBR(r.fim_carregamento||'');
      if(centro==='1110' && desc.includes('VENDA') && rpWithinWindow(ref,inicioDia,agora)) realizadoTon['VENDA MOGI']+=ton;
      return;
    }
    if(tipo==='TRANSFERÊNCIA'){
      const desc=String(r.descricao_documento||'').toUpperCase();
      const ref=parseBR(r.grade_carregamento||'')||parseBR(r.agenda||'')||parseBR(r.fim_carregamento||'');
      if((desc.includes('TRANSFER')||desc.includes('TNF')) && rpWithinWindow(ref,inicioDia,agora)) realizadoTon['TRANSFERÊNCIA']+=ton;
      return;
    }
  });

  // Planejado da grade é automático por fim_carregamento no período
  const planejadoGrade = rows
    .filter(r=>rpWithinWindow(parseBR(r.fim_carregamento||''),inicioDia,agora))
    .reduce((acc,r)=>acc+rpParseToneladas(r),0);
// Gráfico de barras
  const chart=document.getElementById('rp-chart');
  if(chart){
    const maxVal=Math.max(1,...RP_TIPOS.map(t=>{ const p=t==='GRADE'?planejadoGrade:(rpMetas[t]||0); return Math.max(p,realizadoTon[t]||0);}));
    chart.innerHTML=RP_TIPOS.map(tipo=>{
      const plan=tipo==='GRADE'?planejadoGrade:(rpMetas[tipo]||0);
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
      const plan = tipo==='GRADE' ? planejadoGrade : (rpMetas[tipo]||0);
      const tonStr=realTon.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1});
      const pend=Math.max(0,plan-realTon);
      const pct=plan>0?Math.min(100,Math.round(realTon/plan*100)):0;
      const c=RP_COLORS[tipo]||'#3b82f6';
      const barW=plan>0?Math.round(realTon/plan*100):0;
      return `<div class="rp-tipo-card" style="border-color:${c}44;">
        <div class="rp-tipo-title" style="color:${c}">${tipo}</div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin-bottom:8px;">
          <div><div class="rp-num-plan">${plan}</div><div class="rp-num-label">Planej.</div></div>
          <div><div class="rp-num-real">${tonStr}</div><div class="rp-num-label">⚖️ Tons</div></div>
          <div><div class="rp-num-pend">${pend}</div><div class="rp-num-label">Pendente</div></div>
        </div>
        <div style="background:#0f172a;border-radius:4px;height:4px;margin-bottom:6px;overflow:hidden;">
          <div style="height:100%;width:${barW}%;background:${c};border-radius:4px;transition:.4s;"></div>
        </div>
        <div style="display:flex;justify-content:space-between;font-size:10px;color:#64748b;">
          <span>${realTon.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})} t realizadas</span><span style="color:${pct>=100?'#22c55e':c}">${pct}%</span>
        </div>
      </div>`;
    }).join('');
  }

  // Contagem por Status atual
  const statusCountEl=document.getElementById('rp-status-counts');
  if(statusCountEl){
    const statusList=['AG CHEGADA','PATIO','CARREGANDO','EM FATURAMENTO','SEPARANDO','EXPEDIDO','NO SHOW','VEICULO RECUSADO','FOI EMBORA'];
    const statusColors={'AG CHEGADA':'#f59e0b','PATIO':'#64748b','CARREGANDO':'#3b82f6','EM FATURAMENTO':'#06b6d4','SEPARANDO':'#8b5cf6','EXPEDIDO':'#22c55e','NO SHOW':'#ef4444','VEICULO RECUSADO':'#dc2626','FOI EMBORA':'#6b7280'};
    const sCounts={};
    tableData.forEach(r=>{ sCounts[r.status]=(sCounts[r.status]||0)+1; });
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
    let t3DiaProd=0;
    let t3TurnoProd=0;
    const hoje0=new Date(); hoje0.setHours(0,0,0,0);
    const amanha0=new Date(hoje0); amanha0.setDate(amanha0.getDate()+1);
    const t3DiaIni=new Date(hoje0); t3DiaIni.setHours(23,0,0,0);
    const t3DiaFim=new Date(hoje0); t3DiaFim.setHours(23,59,59,999);
    const t3TurnoIni=new Date(hoje0); t3TurnoIni.setHours(23,0,0,0);
    const t3TurnoFim=new Date(amanha0); t3TurnoFim.setHours(6,59,59,999);

    rows.forEach(r=>{
      const grade=parseBR(r.grade_carregamento||'');
      const turno=rpTurnoFromDate(grade);
      if(!turno) return;
      const ton=rpParseToneladas(r);
      byTurno[turno]+=ton;

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
        <div class="rp-turno-num" style="color:${cores[t]}">${byTurno[t].toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})}</div>
        <div style="font-size:10px;color:#64748b;">Toneladas planejadas</div>
        ${t==='T3'
          ? `<div style="font-size:10px;color:#94a3b8;margin-top:6px;line-height:1.6;">
              Dia (23-00): <b style="color:#c4b5fd;">${t3DiaProd.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})} t</b><br/>
              Turno (23-06): <b style="color:#c4b5fd;">${t3TurnoProd.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})} t</b>
            </div>`
          : ''
        }
      </div>
    `).join('');
    const totalCard=`
      <div class="rp-turno-card" style="border-color:#22c55e44;">
        <div class="rp-turno-label" style="color:#22c55e">📦 TOTAL DO DIA</div>
        <div class="rp-turno-num" style="color:#22c55e">${totalTurnos.toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})}</div>
        <div style="font-size:10px;color:#64748b;">Toneladas planejadas</div>
      </div>
    `;
    turnosEl.innerHTML=cards+totalCard;
  }
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
  if(!dt){showErr('Informe o número da DT.');return;}
  const reg={dt,transp,motivo,obs,ts:new Date().toLocaleString('pt-BR')};
  sgRegistros.unshift(reg);
  try{localStorage.setItem('sg_registros',JSON.stringify(sgRegistros));}catch(e){}
  // Salvar no Supabase também
  try{
    await sbInsert('reporte_semgrade',[{
      dt:String(dt),transportadora:transp,motivo,observacao:obs,
      created_at:new Date().toISOString()
    }]);
  }catch(e){/* tabela pode não existir ainda */}
  document.getElementById('sg-dt').value='';
  document.getElementById('sg-transp').value='';
  document.getElementById('sg-motivo').value='';
  document.getElementById('sg-obs').value='';
  showOk('DT '+dt+' registrada como sem grade.');
  renderSemGrade();
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
      <div style="font-size:11px;color:#94a3b8;">${r.motivo||'—'}${r.obs?' · '+r.obs:''}</div>
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
    await sbInsert('reporte_logs',[{dt:dtTo,data_ref:row.data_ref,status:'SUBSTITUIÇÃO DE DT FAKE',transportadora:'DT fake: '+dtFrom,created_at:new Date().toISOString()}]);
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
let importCsvData = null; // parsed rows from CSV

document.addEventListener('keydown', e => {
  if ((e.ctrlKey||e.metaKey) && e.key.toLowerCase()==='b') {
    e.preventDefault();
    openImportModal();
  }
});

function openImportModal(){
  importCsvData=null;
  document.getElementById('import-btn').disabled=true;
  document.getElementById('import-drop-label').textContent='Clique ou arraste o CSV aqui';
  document.getElementById('import-preview').style.display='none';
  document.getElementById('import-result').style.display='none';
  document.getElementById('import-err').style.display='none';
  document.getElementById('import-file-input').value='';
  document.getElementById('import-overlay').style.display='flex';
}

function closeImportModal(){
  document.getElementById('import-overlay').style.display='none';
}

// Mapa de status do CSV → status do dashboard
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

function normalizeStatus(raw){
  const s=(raw||'').trim().toUpperCase();
  return CSV_STATUS_MAP[s] || null;
}

function handleImportFile(file){
  if(!file) return;
  document.getElementById('import-err').style.display='none';
  document.getElementById('import-result').style.display='none';
  document.getElementById('import-drop-label').textContent='⏳ Lendo arquivo…';

  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const text=e.target.result;
      const lines=text.split(/\r?\n/).filter(l=>l.trim());
      if(!lines.length) throw new Error('Arquivo vazio');

      // Detect separator (semicolon or comma)
      const sep=lines[0].includes(';')?';':',';
      const headers=lines[0].split(sep).map(h=>h.trim().replace(/^"|"$/g,'').toUpperCase());

      // Find column indices
      const iDT=headers.findIndex(h=>h==='DT');
      const iStatus=headers.findIndex(h=>h==='STATUS'||h==='STATUS.1');
      // prefer last STATUS col (STATUS.1 in the sample)
      const iStatusFinal=headers.lastIndexOf('STATUS.1')!==-1
        ?headers.lastIndexOf('STATUS.1')
        :(iStatus!==-1?iStatus:-1);

      if(iDT===-1) throw new Error('Coluna DT não encontrada no CSV. Verifique o formato.');

      importCsvData=[];
      const seen=new Set();
      for(let i=1;i<lines.length;i++){
        const cols=lines[i].split(sep).map(c=>c.trim().replace(/^"|"$/g,''));
        const dt=String(cols[iDT]||'').trim();
        if(!dt||seen.has(dt)) continue;
        seen.add(dt);
        const rawStatus=iStatusFinal!==-1?(cols[iStatusFinal]||''):'';
        importCsvData.push({dt, rawStatus});
      }

      if(!importCsvData.length) throw new Error('Nenhuma DT encontrada no CSV.');

      // Preview
      const previewHtml=importCsvData.slice(0,8).map(r=>{
        const mapped=normalizeStatus(r.rawStatus);
        return `<div style="display:grid;grid-template-columns:120px 1fr 1fr;gap:6px;padding:4px 0;border-bottom:1px solid #1e293b;">
          <span style="color:#93c5fd;font-weight:700;">${r.dt}</span>
          <span style="color:#64748b;">${r.rawStatus||'—'}</span>
          <span style="color:${mapped?'#22c55e':'#f59e0b'};">${mapped||'(não mapeado)'}</span>
        </div>`;
      }).join('')+(importCsvData.length>8?`<div style="color:#64748b;font-size:10px;padding-top:4px;">…mais ${importCsvData.length-8} DTs</div>`:'');

      document.getElementById('import-preview-body').innerHTML=
        `<div style="display:grid;grid-template-columns:120px 1fr 1fr;gap:6px;padding:4px 0;border-bottom:1px solid #334155;margin-bottom:4px;">
          <span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;">DT</span>
          <span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;">STATUS CSV</span>
          <span style="font-size:9px;letter-spacing:1.5px;color:#475569;font-weight:700;">→ DASH</span>
        </div>`+previewHtml;
      document.getElementById('import-preview').style.display='block';
      document.getElementById('import-drop-label').textContent=`✅ ${file.name} — ${importCsvData.length} DTs encontradas`;
      document.getElementById('import-btn').disabled=false;
    }catch(err){
      document.getElementById('import-err').textContent='❌ '+err.message;
      document.getElementById('import-err').style.display='block';
      document.getElementById('import-drop-label').textContent='Clique ou arraste o CSV aqui';
    }
  };
  reader.readAsText(file,'latin1');
}

async function runImport(){
  if(!importCsvData||!importCsvData.length) return;

  document.getElementById('import-btn').disabled=true;
  document.getElementById('import-btn').textContent='⏳ Processando…';
  document.getElementById('import-result').style.display='none';
  document.getElementById('import-err').style.display='none';

  try{
    const now=new Date();
    const nowMinutes=now.getHours()*60+now.getMinutes();

    // Filtrar só DTs de HOJE e AMANHÃ do dashboard
    const dashHojeAmanha=tableData.filter(r=>r.dia_ref==='HOJE'||r.dia_ref==='AMANHÃ');

    // Build a Set of DTs in the CSV
    const csvDtSet=new Set(importCsvData.map(r=>String(r.dt)));

    let updated=0, expedited=0, skipped=0;

    // 1) Update DTs present in CSV (só as que estão no dash de hoje/amanhã)
    for(const csvRow of importCsvData){
      const dt=String(csvRow.dt);
      const mappedStatus=normalizeStatus(csvRow.rawStatus);
      if(!mappedStatus){skipped++;continue;}

      const dashRow=dashHojeAmanha.find(r=>String(r.dt)===dt);
      if(!dashRow){skipped++;continue;}
      if(dashRow.status===mappedStatus){continue;} // já correto

      try{
        await saveStatus(dt,dashRow.data_ref,mappedStatus);
        dashRow.status=mappedStatus;
        updated++;
      }catch(e){skipped++;}
    }

    // 2) Marcar como EXPEDIDO: DTs de HOJE no dash, grade ≤ agora, ausentes do CSV
    const finalSet=new Set(STATUS_FINAIS);
    for(const dashRow of dashHojeAmanha){
      if(dashRow.dia_ref!=='HOJE') continue;            // só HOJE
      const dt=String(dashRow.dt);
      if(csvDtSet.has(dt)) continue;                    // está no CSV, pular
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

    // Atualiza tabela
    renderRows();

    const msg=`✅ Concluído! Atualizadas: ${updated} | Expedidas automáticas: ${expedited} | Ignoradas: ${skipped}`;
    document.getElementById('import-result').textContent=msg;
    document.getElementById('import-result').style.display='block';

  }catch(err){
    document.getElementById('import-err').textContent='❌ Erro: '+err.message;
    document.getElementById('import-err').style.display='block';
  }

  document.getElementById('import-btn').disabled=false;
  document.getElementById('import-btn').textContent='⚡ Atualizar Dashboard';
}
