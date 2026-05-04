/* ═══════════════════════════════════════════════════════
   SUPABASE REST
═══════════════════════════════════════════════════════ */
const SB_URL = 'https://pwjatxqtkvwcmzmjjvbi.supabase.co/rest/v1';
const SB_KEY = 'sb_publishable_bUPTDkrOzc0_I3xwNw15aA_lk76gg4w';
const HDR = {'Content-Type':'application/json','apikey':SB_KEY,'Authorization':'Bearer '+SB_KEY};

async function sbGet(table, qs=''){
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

let agendRows=[], dtsMescladas=[], exportMap={}, tipoOpMap={}, tableData=[];
let panelDT=null;
let currentTableTab='todas';
let reagendDT=null;

/* ═══════════════════════════════════════════════════════
   INIT
═══════════════════════════════════════════════════════ */
window.addEventListener('DOMContentLoaded', async ()=>{
  document.getElementById('sync-txt').textContent='ONLINE';
  await tryLoadExisting();
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
  [1,2,3,4].forEach(i=>{
    document.getElementById('passo'+i).style.display=i===n?'block':'none';
    const d=document.getElementById('dot'+i);
    if(!d)return;
    d.className='sdot'+(i<n?' done':i===n?' active':'');
    d.textContent=i<n?'✓':i;
  });
  document.getElementById('aerr').style.display='none';
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

function switchTableTab(tab){
  currentTableTab=tab;
  ['todas','finalizadas','agenda'].forEach(t=>{
    const el=document.getElementById('tab-'+t);
    if(el) el.classList.toggle('active',t===tab);
  });
  const board=document.getElementById('mat-board');
  const twrap=document.getElementById('twrap');
  const agendaWrap=document.getElementById('agenda-wrap');
  if(board) board.style.display='none';
  if(tab==='agenda'){
    if(twrap) twrap.style.display='none';
    if(agendaWrap) agendaWrap.style.display='block';
    renderAgendaTab();
  } else {
    if(twrap) twrap.style.display='block';
    if(agendaWrap) agendaWrap.style.display='none';
    renderRows();
  }
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
/* Detecta se o arquivo é o Relatório de Expedição do SAP (CSV com colunas pt-BR) */
function isRelatorioExpedicao(h){
  return h.some(c=>c.includes('Nº transporte')||c.includes('Nr. transporte')||c.includes('Nome Transportadora'))
      && h.some(c=>c.includes('Data Agendamento')||c.includes('Data Carregar'));
}

/* Converte data DD.MM.YYYY + hora HH:MM:SS → Date ou null */
function parseDotDate(dateStr, timeStr){
  if(!dateStr||!dateStr.trim()) return null;
  const m=dateStr.trim().match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if(!m) return null;
  const [,d,mo,y]=m;
  const t=timeStr?timeStr.trim().match(/(\d{2}):(\d{2})/):null;
  return new Date(+y,+mo-1,+d,t?+t[1]:0,t?+t[2]:0);
}

/* Processa Relatório de Expedição — agrupa por Nº transporte e usa Data/Hora Agendamento */
function processRelatorioExpedicao(text){
  const sep=text.indexOf(';')>text.indexOf('\t')?'\t':';';
  const lines=text.split(/\r?\n/).filter(l=>l.trim());
  const strip=s=>(s||'').trim().replace(/^"|"$/g,'');
  const h=lines[0].split(sep).map(strip);
  const ci=(...names)=>{for(const n of names){const i=h.findIndex(c=>c.includes(n));if(i!==-1)return i;} return -1;};
  const iDT   =ci('Nº transporte','Nr. transporte');
  const iTransp=ci('Nome Transportadora');
  const iDataAg=ci('Data Agendamento');
  const iHoraAg=ci('Hora Agendamento');
  const iDataCa=ci('Data Carregar');
  const iHoraCa=ci('Hora Carregar');
  const iLocal =ci('Local','LOCAL');
  const iPeso  =ci('Peso líquido','Peso');
  if(iDT===-1) throw new Error('Coluna "Nº transporte" não encontrada no arquivo.\nColunas detectadas: '+h.join(' | '));
  const byDT={};
  for(let i=1;i<lines.length;i++){
    const c=lines[i].split(sep).map(strip);
    const dtRaw=(c[iDT]||'').replace(/\.0+$/,'').trim();
    if(!dtRaw) continue;
    const dt=dtRaw.replace(/\D/g,'');
    if(!dt||dt.length<5) continue;
    if(!byDT[dt]){
      // Prioriza Data Agendamento; fallback: Data Carregar
      const agDate=iDataAg!==-1?parseDotDate(c[iDataAg],iHoraAg!==-1?c[iHoraAg]:''):null;
      const caDate=iDataCa!==-1?parseDotDate(c[iDataCa],iHoraCa!==-1?c[iHoraCa]:''):null;
      const agenda=agDate||caDate;
      byDT[dt]={
        DT:dt,
        LOCAL:iLocal!==-1?(c[iLocal]||'').trim():'1110',
        TRANSPORTADORA:iTransp!==-1?(c[iTransp]||'').trim():'',
        AGENDA:agenda,
        FIM_AGENDA:null,
        TIPO:'',
        PESO:iPeso!==-1?(c[iPeso]||'').trim():'',
      };
    }
  }
  agendRows=Object.values(byDT);
  if(!agendRows.length) throw new Error('Nenhuma DT encontrada no Relatório de Expedição.');
  return agendRows.length;
}

function processAgend(file){
  if(!file)return;
  showInf('Lendo agendamento…');
  document.getElementById('dz1l').innerHTML='<span class="dz-ok">✓ '+file.name+'</span>';
  file.arrayBuffer().then(buf=>{
    try{
      exportMap={};
      const matCnt=tryImportMaterialsFromAgend(buf);
      if(matCnt>0) showOk('✨ '+matCnt+' linhas de materiais importadas automaticamente do arquivo de agendamento!');
      const text=decodeBuf(buf);
      const lines=text.split(/\r?\n/).filter(l=>l.trim());
      // Detecta o separador e as colunas do cabeçalho
      const firstLine=lines[0];
      const sep=firstLine.indexOf(';')>(firstLine.match(/\t/g)||[]).length?';':'\t';
      const h=firstLine.split(sep).map(s=>s.trim().replace(/^"|"$/g,''));

      if(isRelatorioExpedicao(h)){
        // ── Formato: Relatório de Expedição SAP ──
        const cnt=processRelatorioExpedicao(text);
        showOk('📋 Relatório de Expedição lido! '+cnt+' DTs encontradas.');
        hideInf();
        renderStep2();
        return;
      }

      // ── Formato original: arquivo de agendamento TSV ──
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
          DT:(c[iDT]||'').trim(),LOCAL:loc,
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
  const dtsSG=agendRows.filter(r=>!r.AGENDA);
  dtsMescladas=[...dtsH,...dtsA,...dtsSG];
  const semGradeInfo=dtsSG.length>0?' · <b style="color:#f59e0b">⚠️ '+dtsSG.length+' sem grade</b>':'';  document.getElementById('p2info').innerHTML=
    '✅ <b style="color:#22c55e">'+agendRows.length+' linhas</b> encontradas · '+
    '<b style="color:#60a5fa">'+dtsH.length+'</b> hoje · <b style="color:#a78bfa">'+dtsA.length+'</b> amanhã'+
    semGradeInfo+
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
    if(!kDT||!kMat) return 0;
    let cnt=0;
    for(const row of rows){
      const rawDT=row[kDT];
      const dt=(rawDT===''||rawDT===null||rawDT===undefined)?'':String(Math.round(Number(rawDT)));
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
   PASSO 3 — ZLES002
═══════════════════════════════════════════════════════ */
function processZles(file){
  if(!file)return;
  showInf('Lendo ZLES002…');
  document.getElementById('dz2l').innerHTML='<span class="dz-ok">✓ '+file.name+'</span>';
  file.arrayBuffer().then(buf=>{
    const text=decodeBuf(buf);
    parseZlesText(text);
  });
}

function processZlesPaste(){
  const text=document.getElementById('zles-paste').value.trim();
  if(!text){showErr('Cole o conteúdo do SAP no campo de texto primeiro.');return;}
  showInf('Processando texto colado…');
  setTimeout(()=>parseZlesText(text),50);
}

function parseZlesText(text){
  const dbg=document.getElementById('dbgZles');
  dbg.textContent='';dbg.classList.add('show');
  try{
    const lines=text.split(/\r?\n/).filter(l=>l.trim());
    if(!lines.length)throw new Error('Conteúdo vazio.');
    const L0=lines[0];
    const sep=(L0.match(/;/g)||[]).length>=(L0.match(/\t/g)||[]).length?';':'\t';
    dbg.textContent+=`Separador: "${sep===';'?'ponto-e-vírgula (;)':'tab (\\t)'}"\n`;
    dbg.textContent+=`Total de linhas: ${lines.length}\n`;
    const strip=s=>(s||'').trim().replace(/^"|"$/g,'');
    const h=lines[0].split(sep).map(strip);
    dbg.textContent+=`Colunas (${h.length}): ${h.slice(0,10).join(' | ')}…\n`;
    const findIdx=(...termos)=>{
      for(const t of termos){
        const i=h.findIndex(v=>v.toUpperCase().includes(t.toUpperCase()));
        if(i!==-1)return i;
      }
      return -1;
    };
    let iDT=6;
    if(lines.length>1){
      const primeiraLinha=lines[1].split(sep).map(strip);
      const valG=(primeiraLinha[6]||'').replace(/\D/g,'');
      if(valG.length<5){
        const alt=h.findIndex(v=>{
          const u=v.toUpperCase();
          return (u.includes('TRANSPORTE')||u.includes('NR TRANSP')||u.includes('Nº TRANSP'))
                 && !u.includes('NOME') && !u.includes('COD') && !u.includes('CÓD');
        });
        if(alt!==-1) iDT=alt;
      }
    }
    const iTipo=findIdx('TIPO','OPERAÇÃO','OPERACAO')!==-1?findIdx('TIPO','OPERAÇÃO','OPERACAO'):9;
    const iDesc=findIdx('DESCRI')!==-1?findIdx('DESCRI'):9;
    const iMat =findIdx('MATERIAL')!==-1?findIdx('MATERIAL'):11;
    const iQtd =findIdx('QTDE REMESSA','QTD REMESSA')!==-1?findIdx('QTDE REMESSA','QTD REMESSA'):12;
    dbg.textContent+=`Usando: DT=col${iDT}(${h[iDT]||'?'}) | Tipo=col${iTipo}(${h[iTipo]||'?'}) | Mat=col${iMat}(${h[iMat]||'?'}) | Qtd=col${iQtd}(${h[iQtd]||'?'})\n`;
    exportMap={};tipoOpMap={};
    let linhasOk=0,linhasIgnoradas=0;
    for(let i=1;i<lines.length;i++){
      const c=lines[i].split(sep).map(strip);
      if(c.every(v=>!v)){linhasIgnoradas++;continue;}
      const dtRaw=(c[iDT]||'').replace(/\.0+$/,'').trim();
      const dt=dtRaw.replace(/\D/g,'');
      if(!dt||dt.length<5){linhasIgnoradas++;continue;}
      const matRaw=(c[iMat]||'').replace(/\.0+$/,'').trim();
      const mat=matRaw.replace(/\D/g,'');
      const qtdRaw=(c[iQtd]||'').replace(/\./g,'').replace(',','.').trim();
      const qtd=qtdRaw?String(Math.round(Number(qtdRaw)||0)):'';
      if(!tipoOpMap[dt]){
        const t1=(iTipo!==-1?c[iTipo]||'':'').toUpperCase();
        const t2=(iDesc!==-1&&iDesc!==iTipo?c[iDesc]||'':'').toUpperCase();
        const tipoRaw=t1+' '+t2;
        tipoOpMap[dt]=(tipoRaw.includes('TRANSFER')||tipoRaw.includes('FILIAL')||tipoRaw.includes('INTERCOMP')||tipoRaw.includes('REMESSA P/')||tipoRaw.includes('TNF'))?'TRANSFERÊNCIA':'VENDA NORMAL';
      }
      if(mat){
        if(!exportMap[dt])exportMap[dt]=[];
        const existing=exportMap[dt].find(m=>m.material===mat);
        if(existing){
          const q1=Number(existing.quantidade)||0;
          const q2=Number(qtd)||0;
          existing.quantidade=String(q1+q2);
        }else{
          exportMap[dt].push({material:mat,quantidade:qtd});
        }
      }
      linhasOk++;
    }
    const totalDTs=Object.keys(exportMap).length;
    const tipos=[...new Set(Object.values(tipoOpMap))];
    dbg.textContent+=`\n✅ Linhas processadas: ${linhasOk} | Ignoradas: ${linhasIgnoradas}\n`;
    dbg.textContent+=`DTs com materiais: ${totalDTs}\n`;
    Object.keys(exportMap).slice(0,3).forEach(dt=>{
      const mats=exportMap[dt].slice(0,4).map(m=>`${m.material}(${m.quantidade})`).join(', ');
      dbg.textContent+=`  DT ${dt}: ${mats}${exportMap[dt].length>4?'…':''}\n`;
    });
    hideInf();
    showOk(`ZLES002 OK — ${totalDTs} DTs · ${linhasOk} linhas · tipos: ${tipos.join(', ')}`);
    buildTable();
  }catch(e){
    hideInf();
    dbg.textContent+='\n❌ ERRO: '+e.message;
    showErr('Erro ZLES002: '+e.message);
  }
}

/* ═══════════════════════════════════════════════════════
   PLANILHA EXTRA XLSX
═══════════════════════════════════════════════════════ */
function processExtra(file){
  if(!file)return;
  showInf('Lendo planilha extra…');
  document.getElementById('dz3l').innerHTML='<span class="dz-ok">✓ '+file.name+'</span>';
  const dbg=document.getElementById('dbgExtra');
  dbg.textContent='';dbg.classList.add('show');
  file.arrayBuffer().then(buf=>{
    try{
      const wb=XLSX.read(new Uint8Array(buf),{type:'array'});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:''});
      if(!rows.length)throw new Error('Planilha vazia.');
      const keys=Object.keys(rows[0]);
      dbg.textContent='Colunas: '+keys.join(' | ')+'\n';
      const fc=(...ts)=>{for(const t of ts){const k=keys.find(k=>k.toUpperCase().includes(t.toUpperCase()));if(k)return k;}return null;};
      const kDT=fc('Nº TRANSPORTE','N TRANSPORTE','TRANSPORTE');
      const kMat=fc('MATERIAL');
      const kQtd=fc('QTDE REMESSA','QTD REMESSA','QTDE');
      dbg.textContent+=`Col DT: ${kDT||'NÃO ACHADA'} | Mat: ${kMat||'NÃO ACHADA'} | Qtd: ${kQtd||'NÃO ACHADA'}\n`;
      if(!kDT||!kMat)throw new Error(
        '⚠️ Este campo é para uma planilha auxiliar com colunas: "Nº transporte", "Material" e "Qtde Remessa".\n\n'+
        'Para a ZLES002 do SAP, use as abas "📂 Subir arquivo CSV" ou "📋 Colar do SAP" acima.\n\n'+
        'Colunas encontradas neste arquivo: '+keys.join(', ')
      );
      let cnt=0;
      for(const row of rows){
        const rawDT=row[kDT];
        const dt=(rawDT===''||rawDT===null||rawDT===undefined)?'':String(Math.round(Number(rawDT)));
        const rawMat=row[kMat];
        const mat=(rawMat===''||rawMat===null||rawMat===undefined)?'':String(Math.round(Number(rawMat)));
        const qtd=kQtd?String(row[kQtd]||''):'';
        if(!dt||!mat||dt==='NaN'||mat==='NaN')continue;
        cnt++;
        if(!exportMap[dt])exportMap[dt]=[];
        if(!exportMap[dt].some(m=>m.material===mat))exportMap[dt].push({material:mat,quantidade:qtd});
      }
      dbg.textContent+=`Linhas mescladas: ${cnt} | Total DTs com materiais: ${Object.keys(exportMap).length}\n`;
      hideInf();
      showOk('Planilha extra OK: '+cnt+' linhas mescladas em '+Object.keys(exportMap).length+' DTs.');
      buildTable();
    }catch(e){hideInf();dbg.textContent+='\n❌ ERRO: '+e.message;showErr('Erro planilha extra: '+e.message);}
  });
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

/* ═══════════════════════════════════════════════════════
   BUILD TABLE
═══════════════════════════════════════════════════════ */
async function buildTable(){
  showInf('Sincronizando com o banco…');
  const T=today(),AM=tomorrow();
  const kT=dKey(T), kAM=dKey(AM);
  try{
    const existing=await sbGet('reporte_carga',
      `data_ref=in.("${kT}","${kAM}")&select=dt,data_ref,status,hora_chegada,tipo_operacao,reagendada`
    );
    const exMap={};
    (existing||[]).forEach(r=>{exMap[r.dt+'_'+r.data_ref]=r;});

    const STATUS_MANTER=['EXPEDIDO','NO SHOW','VEICULO RECUSADO','FOI EMBORA','PATIO','CARREGANDO','EM FATURAMENTO','SEPARANDO'];
    const newDTKeys=new Set(dtsMescladas.map(r=>r.DT+'_'+dKey(r.AGENDA||T)));
    const paraRemover=(existing||[]).filter(r=>{
      const key=r.dt+'_'+r.data_ref;
      return !newDTKeys.has(key) && !STATUS_MANTER.includes(r.status||'');
    });
    for(const r of paraRemover){
      try{ await sbDelete('reporte_carga',{dt:r.dt,data_ref:r.data_ref}); }catch(e2){}
    }

    const rows=dtsMescladas.map(dt=>{
      const ref=dKey(dt.AGENDA||T);
      const ex=exMap[dt.DT+'_'+ref]||{};
      const diaRef=sameDay(dt.AGENDA,T)?'HOJE':'AMANHÃ';
      const tipoOp=tipoOpMap[dt.DT]||ex.tipo_operacao||'';
      const statusAtual = ex.status && ex.status !== '' ? ex.status : 'AG CHEGADA';
      const horaAtual   = ex.hora_chegada || '';
      return {
        dt:String(dt.DT),
        transportadora:String(dt.TRANSPORTADORA||''),
        grade_carregamento:String(dt.AGENDA?fmtDT(dt.AGENDA,true):''),
        fim_carregamento:String(dt.FIM_AGENDA?fmtDT(dt.FIM_AGENDA,true):''),
        hora_chegada:horaAtual,
        status:statusAtual,
        tipo:String(dt.TIPO||''),
        toneladas:String(dt.PESO||''),
        agenda:String(dt.AGENDA?fmtDT(dt.AGENDA,true):''),
        local_cd:String(dt.LOCAL||''),
        dia_ref:diaRef,
        data_ref:String(ref),
        tipo_operacao:String(tipoOp),
        reagendada: ex.reagendada || false,
      };
    });

    await sbUpsert('reporte_carga',rows);
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
      hora_chegada:'',status:'AG CHEGADA',tipo:dt.TIPO,toneladas:dt.PESO,
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
  const T=today(),AM=tomorrow();
  const base=[dKey(T),dKey(AM)];
  try{
    const todayRows=await sbGet('reporte_carga',`data_ref=in.("${base[0]}","${base[1]}")&select=dt&limit=1`);
    if(todayRows&&todayRows.length) return base;
    const rec=await sbGet('reporte_carga','select=data_ref,updated_at&order=updated_at.desc&limit=500');
    const uniq=[];
    (rec||[]).forEach(r=>{const k=String(r.data_ref||'');if(k&&!uniq.includes(k)) uniq.push(k);});
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
   RELÓGIO
═══════════════════════════════════════════════════════ */
function clockEmoji(row){
  if(row.dia_ref!=='HOJE') return '';
  if(STATUS_FINAIS.includes(row.status)) return '';
  if(!row.grade_carregamento) return '';
  const grade=parseBR(row.grade_carregamento);
  if(!grade) return '';
  const agora=new Date();
  const diffMin=(grade - agora)/60000;
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

  let rows=[...tableData].sort((a,b)=>parseAgendaDateTime(a)-parseAgendaDateTime(b));
  if(currentTableTab==='finalizadas'){
    rows=rows.filter(r=>STATUS_FINAIS.includes(r.status));
  }else{
    rows=rows.filter(r=>!STATUS_FINAIS.includes(r.status));
  }

  const tbody=document.getElementById('tbody');tbody.innerHTML='';
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
    const op=(row.tipo_operacao||'').toUpperCase();
    const opTag=op.includes('TRANSFER')
      ?'<span class="tag-transf">🔄 TRANSFERÊNCIA</span>'
      :op.includes('VENDA')
      ?'<span class="tag-venda">🛒 VENDA NORMAL</span>'
      :'<span class="tag-sem">—</span>';
    const diaTag=isH?'<span class="tag-hoje">HOJE</span>':'<span class="tag-amanha">AMANHÃ</span>';
    const clock=clockEmoji(row);
    const reagBtn=isFinal
      ? `<button class="bg" style="font-size:10px;padding:3px 9px;border-color:#f59e0b55;color:#f59e0b;" onclick="openReagend('${row.dt}','${row.data_ref}')">🔄 Reagendar</button>`
      : '';
    tr.innerHTML=
      `<td>${diaTag}</td>`+
      `<td><span class="td-dt" onclick="openPanel('${row.dt}','${row.data_ref}')">${clock}${row.dt}</span></td>`+
      `<td class="td-transp">${row.transportadora||'—'}</td>`+
      `<td class="td-time">${row.grade_carregamento||'—'}</td>`+
      `<td class="td-time">${row.fim_carregamento||'—'}</td>`+
      `<td><input type="time" value="${row.hora_chegada||''}" onchange="saveField('${row.dt}','${row.data_ref}','hora_chegada',this.value)"/></td>`+
      `<td><div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;"><select class="ss" style="background:${c}22;border:1.5px solid ${c}99;color:${c}" onchange="saveStatus('${row.dt}','${row.data_ref}',this.value)">`+
        STATUS_OPTIONS.map(s=>`<option value="${s}"${s===row.status?' selected':''}>${s}</option>`).join('')+
      `</select>${opTag}</div></td>`+
      `<td class="td-sm">${row.tipo||'—'}</td>`+
      `<td class="td-sm">${row.toneladas||'—'}</td>`+
      `<td>${opTag}</td>`+
      `<td class="td-ag">${row.agenda||'—'}</td>`+
      `<td>${reagBtn}</td>`;
    tbody.appendChild(tr);
  });

  if(!rows.length){
    const tr=document.createElement('tr');
    tr.innerHTML=`<td colspan="12" style="text-align:center;padding:30px;color:#334155;font-size:13px;">
      ${currentTableTab==='finalizadas'?'Nenhuma DT finalizada ainda.':'Todas as DTs do dia estão finalizadas.'}
    </td>`;
    tbody.appendChild(tr);
  }
}

/* ═══════════════════════════════════════════════════════
   SALVAR
═══════════════════════════════════════════════════════ */
async function saveField(dt,dataRef,field,value){
  spin(true);
  try{
    await sbPatch('reporte_carga',{[field]:String(value),updated_at:new Date().toISOString()},{dt,data_ref:dataRef});
    const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef);
    if(row){row[field]=value;renderRows();}
  }catch(e){showErr('Erro ao salvar: '+e.message);}
  spin(false);
}

async function saveStatus(dt,dataRef,value){
  spin(true);
  try{
    await sbPatch('reporte_carga',{status:String(value),updated_at:new Date().toISOString()},{dt,data_ref:dataRef});
    const row=tableData.find(r=>r.dt===dt&&r.data_ref===dataRef);
    await sbInsert('reporte_logs',[{
      dt:String(dt),
      data_ref:String(dataRef),
      status:String(value),
      transportadora:row?String(row.transportadora||''):'',
      created_at:new Date().toISOString(),
    }]);
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
async function loadLogs(){
  const list=document.getElementById('log-list');
  list.innerHTML='<div style="color:#64748b;padding:12px;">Carregando…</div>';
  try{
    const logs=await sbGet('reporte_logs','order=created_at.desc&limit=200');
    if(!logs||!logs.length){list.innerHTML='<div style="color:#334155;padding:12px;">Nenhum log registrado ainda.</div>';return;}
    list.innerHTML='';
    logs.forEach(l=>{
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
  const op=(row.tipo_operacao||'').toUpperCase();
  const opBadge=op.includes('TRANSFER')
    ?'<span class="tag-transf" style="padding:3px 10px;font-size:11px;">🔄 TRANSFERÊNCIA</span>'
    :op.includes('VENDA')
    ?'<span class="tag-venda"  style="padding:3px 10px;font-size:11px;">🛒 VENDA NORMAL</span>'
    :'';
  document.getElementById('pdt-info').innerHTML=
    `<b style="color:#93c5fd;font-size:15px">${dt}</b>  ${opBadge}<br/>`+
    `<span style="color:#64748b">Transportadora:</span> <b style="color:#e2e8f0">${row.transportadora||'—'}</b><br/>`+
    `<span style="color:#64748b">Grade:</span> <b style="color:#a5b4fc">${row.grade_carregamento||'—'}</b>  →  <b style="color:#a5b4fc">${row.fim_carregamento||'—'}</b>`;

  let listaMats=[];
  try{
    const saved=await sbGet('reporte_materiais',`dt=eq.${encodeURIComponent(dt)}&data_ref=eq.${encodeURIComponent('"'+dataRef+'"')}&order=ordem.asc`);
    listaMats=saved||[];
  }catch(e){ console.error('Erro ao buscar materiais:',e); }

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
      `<input class="mat-input" placeholder="Qtd"      value="${m.quantidade||''}" data-f="quantidade"/>`+
      `<button class="mat-del" onclick="this.parentElement.remove()">🗑</button>`;
    list.appendChild(row);
  });
}

function addMatRow(){
  const row=document.createElement('div');row.className='mat-row';
  row.innerHTML=
    `<input class="mat-input" placeholder="Material" data-f="material"/>`+
    `<input class="mat-input" placeholder="Qtd"      data-f="quantidade"/>`+
    `<button class="mat-del" onclick="this.parentElement.remove()">🗑</button>`;
  document.getElementById('mat-list').appendChild(row);
  row.querySelector('input').focus();
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
      tableData.forEach(r=>{ if(r.dt&&r.tipo_operacao) tipoOpMap[r.dt]=r.tipo_operacao; });
      renderRows();
      setStep(4);
      showOk('Tabela carregada do banco. Suba o agendamento para atualizar.');
    }
  }catch(e){/*sem dados, fica no passo 1*/}
}

/* ═══════════════════════════════════════════════════════
   DRAG & DROP
═══════════════════════════════════════════════════════ */
['dz1','dz2','dz3'].forEach(id=>{
  const el=document.getElementById(id);if(!el)return;
  el.addEventListener('dragover',e=>{e.preventDefault();el.classList.add('drag');});
  el.addEventListener('dragleave',()=>el.classList.remove('drag'));
  el.addEventListener('drop',e=>{
    e.preventDefault();el.classList.remove('drag');
    const f=e.dataTransfer.files[0];if(!f)return;
    if(id==='dz1')processAgend(f);
    else if(id==='dz2')processZles(f);
    else processExtra(f);
  });
});

/* ═══════════════════════════════════════════════════════
   EXPORT CSV
═══════════════════════════════════════════════════════ */
function exportCSV(){
  const cols=['DIA','DT','TRANSPORTADORA','GRADE','FIM','HORA CHEGADA','STATUS','TIPO VEÍ.','TONELADAS','OPERAÇÃO','AGENDA'];
  const rows=tableData.map(r=>[r.dia_ref,r.dt,r.transportadora,r.grade_carregamento,r.fim_carregamento,r.hora_chegada,r.status,r.tipo,r.toneladas,r.tipo_operacao||'',r.agenda]);
  const csv=[cols,...rows].map(r=>r.map(v=>'"'+String(v||'').replace(/"/g,'""')+'"').join(';')).join('\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'}));
  a.download='reporte_'+new Date().toLocaleDateString('pt-BR').replace(/\//g,'-')+'.csv';
  a.click();
}

/* ═══════════════════════════════════════════════════════
   ABA AGENDA — renderiza tabela da planilha importada
═══════════════════════════════════════════════════════ */
function renderAgendaTab(){
  const tbody=document.getElementById('agenda-tbody');
  if(!tbody) return;
  tbody.innerHTML='';
  if(!agendRows||!agendRows.length){
    tbody.innerHTML=`<tr><td colspan="8" style="text-align:center;padding:30px;color:#334155;">Nenhum arquivo de agendamento carregado ainda. Suba o arquivo no Passo 1.</td></tr>`;
    document.getElementById('agenda-count').textContent='0 linhas';
    return;
  }
  document.getElementById('agenda-count').textContent=agendRows.length+' linhas';
  agendRows.forEach(r=>{
    const tr=document.createElement('tr');
    tr.innerHTML=
      `<td class="td-sm">${r.DT||'—'}</td>`+
      `<td class="td-transp">${r.TRANSPORTADORA||'—'}</td>`+
      `<td class="td-sm">${r.LOCAL||'—'}</td>`+
      `<td class="td-time">${r.AGENDA?fmtDT(r.AGENDA,true):'—'}</td>`+
      `<td class="td-time">${r.FIM_AGENDA?fmtDT(r.FIM_AGENDA,true):'—'}</td>`+
      `<td class="td-sm">${r.TIPO||'—'}</td>`+
      `<td class="td-sm">${r.PESO||'—'}</td>`+
      `<td><span class="${!r.AGENDA?'tag-sem-grade':r.AGENDA&&sameDay(r.AGENDA,today())?'tag-hoje':'tag-amanha'}">${!r.AGENDA?'SEM GRADE':r.AGENDA&&sameDay(r.AGENDA,today())?'HOJE':'AMANHÃ'}</span></td>`;
    tbody.appendChild(tr);
  });
}

function exportAgendaXLSX(){
  if(!agendRows||!agendRows.length){showErr('Carregue o arquivo de agendamento primeiro.');return;}
  const cols=['DT','TRANSPORTADORA','LOCAL','GRADE','FIM GRADE','TIPO VEÍCULO','PESO','DIA'];
  const rows=agendRows.map(r=>[
    r.DT||'',
    r.TRANSPORTADORA||'',
    r.LOCAL||'',
    r.AGENDA?fmtDT(r.AGENDA,true):'',
    r.FIM_AGENDA?fmtDT(r.FIM_AGENDA,true):'',
    r.TIPO||'',
    r.PESO||'',
    r.AGENDA&&sameDay(r.AGENDA,today())?'HOJE':'AMANHÃ',
  ]);
  const ws=XLSX.utils.aoa_to_sheet([cols,...rows]);
  // Largura das colunas
  ws['!cols']=[{wch:14},{wch:30},{wch:12},{wch:18},{wch:18},{wch:16},{wch:10},{wch:8}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Agendamento');
  XLSX.writeFile(wb,'agendamento_'+new Date().toLocaleDateString('pt-BR').replace(/\//g,'-')+'.xlsx');
}
