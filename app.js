// ─── Field mappings ────────────────────────────────────────────────────────
const FIELDS = {
  personen:    { Kuerzel:'Title', Vorname:'field_1', Nachname:'field_2' },
  projekte:    { Projekt_Kuerzel:'Title', Beschreibung:'field_1' },
  lagerfolgen: { Lagerfolge_ID:'Title', Name:'field_1', Norm:'field_2', Anwendung:'field_3' },
  experimente: {
    Experiment_ID:'Title', Projekt_Kuerzel:'field_2',
    Datum:'field_3', Person_Kuerzel:'field_4', Projekttitel:'field_5',
    Beschreibung:'field_6', Beobachtungen:'field_7', Kommentar:'field_8',
    Presse:'Presse', Pressdruck:'Pressdruck', Presstemperatur:'Presstemperatur', Presszeit:'Presszeit',
  },
  material: {
    Experiment_ID:'Title', Laenge_mm:'field_1', Breite_mm:'field_2',
    Lagerfolge_ID:'field_3', Kraft_N:'field_4', Holzbruch_pct:'field_5', Kommentar:'field_6',
    Maschine:'Maschine',
  },
  komponenten: {
    Experiment_ID:'Title', Quelle_Typ:'field_1', Chemikalie_Name:'field_2',
    Experiment_Ref:'field_3', Komponente_Name:'field_4', Hersteller:'field_5',
    Menge:'field_7', Einheit:'field_8', Rolle:'field_9',
  },
  chemikalien: {
    Chemikalie_Name:'Title', IUPAC:'field_1', Formel:'field_2',
    Hersteller:'field_3', Menge:'field_4', Einheit:'field_5',
    Ort:'field_6', Herkunft:'field_7', Kommentar:'field_8', Fuellstand:'F_x00fc_llstand',
  },
  feststoffgehalt: {
    Experiment_ID:'Title', Probe:'field_1', Leergewicht_g:'field_2',
    Einwaage_g:'field_3', Endgewicht_g:'field_4', Kommentar:'field_5',
  },
  lagerfolgen_schritte: {
    Lagerfolge_ID:'Title', Schritt_Nr:'field_1', Behandlung:'field_2', Medium:'field_3',
    Temperatur_C:'field_4', RH_pct:'field_5', Dauer_h:'field_6', Kommentar:'field_7',
  },
  maschinen: {
    Name:'Title', Hersteller:'field_1', Typ:'field_2', Maschinen_ID:'field_3',
    Kuerzel:'field_4', Kommentar:'field_5',
  },
};
function mapFrom(items,fm){return items.map(item=>{const o={_spId:item.Id};for(const[k,v]of Object.entries(fm))o[k]=item[v]??null;return o;});}
function mapTo(obj,fm){const o={};for(const[k,v]of Object.entries(fm)){if(k in obj)o[v]=obj[k];}return o;}

// ─── MSAL ─────────────────────────────────────────────────────────────────
const msalConfig={auth:{clientId:'9c5f89c5-d994-4b6e-9215-5f5cd4c09753',authority:'https://login.microsoftonline.com/08a7f19b-4557-486a-8f49-71dcef345176',redirectUri:window.location.origin+window.location.pathname.replace(/index\.html$/,'')},cache:{cacheLocation:'sessionStorage'}};
const SP='https://oucvbj.sharepoint.com/sites/FirmaLeibnizUniversittHannoverITE-LIGARO';
const SP_HOST='https://oucvbj.sharepoint.com';
const SP_SCOPE='https://oucvbj.sharepoint.com/AllSites.Write';
const LIST={experimente:'Experimente',material:'Materialpruefung',lagerfolgen:'Lagerfolgen',personen:'Personen',projekte:'Projektkuerzel',komponenten:'Komponenten',chemikalien:'Chemikalienbestand',feststoffgehalt:'Feststoffgehalt',lagerfolgen_schritte:'Lagerfolgen_Schritte',maschinen:'Maschinen'};

let msalApp;
async function getMsal(){if(!msalApp){msalApp=new msal.PublicClientApplication(msalConfig);await msalApp.initialize();}return msalApp;}
async function login(){document.getElementById('login-error').textContent='';try{const app=await getMsal();const r=await app.loginPopup({scopes:[SP_SCOPE]});onLoggedIn(r.account);}catch(e){document.getElementById('login-error').textContent='Anmeldung fehlgeschlagen: '+e.message;}}
async function getToken(){const app=await getMsal(),accs=app.getAllAccounts();if(!accs.length)throw new Error('Nicht angemeldet');return(await app.acquireTokenSilent({scopes:[SP_SCOPE],account:accs[0]}).catch(()=>app.acquireTokenPopup({scopes:[SP_SCOPE]}))).accessToken;}
function onLoggedIn(account){document.getElementById('login-screen').style.display='none';document.getElementById('app').style.display='block';document.getElementById('user-info').textContent=account.name||account.username;loadAll();}
(async()=>{const app=await getMsal();const a=app.getAllAccounts();if(a.length)onLoggedIn(a[0]);})();

// ─── State ────────────────────────────────────────────────────────────────
let allExp=[],allMat=[],allChem=[],allKomps=[],allSC=[],allLafoSchritte=[],allMaschinen=[],personen=[],projekte=[],lagerfolgen=[];
let entityTypes={},digestVal=null,digestExp=0;
let selectedProj=new Set(),selectedPers=new Set(),selectedLafo=new Set(),selectedOrt=new Set();
let selectedResProj=new Set(),selectedResPers=new Set(),selectedResLafo=new Set();
let mpaRanges={};
let activeTextCol='Beschreibung';
let editingExp=null,editingMat=null,editingChem=null,editingSC=null,editingLafo=null,editingMach=null,editingProj=null;
let lafoSchrittIdx=0,deletedLafoSchritte=[];
let kompIdx=0,deletedKomps=[];
// Sort state: col + dir (1=asc, -1=desc)
let sortState={
  exp:{col:'Datum',dir:-1},
  mat:{col:'_datum',dir:-1},
  res:{col:'Datum',dir:-1},
};

// ─── API ──────────────────────────────────────────────────────────────────
function setStatus(msg){document.getElementById('status-msg').textContent=msg;}
function enc(s){return encodeURIComponent(s);}
function esc(v){if(v==null)return'';return String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function fmtDate(v){if(!v||v.startsWith('0001'))return'';return v.substring(0,10);}
function fmtDec(n,d){if(n==null||isNaN(n))return'';return Number(n).toFixed(d).replace('.',',');}
function calcMpa(l,b,f){const a=parseFloat(l)*parseFloat(b),fo=parseFloat(f);if(!a||!fo)return null;return(fo/a).toFixed(2);}

async function getDigest(token){if(digestVal&&Date.now()<digestExp)return digestVal;const r=await fetch(`${SP}/_api/contextinfo`,{method:'POST',headers:{Accept:'application/json;odata=verbose',Authorization:'Bearer '+token}});if(!r.ok)throw new Error('contextinfo '+r.status);const d=await r.json();digestVal=d.d.GetContextWebInformation.FormDigestValue;digestExp=Date.now()+25*60*1000;return digestVal;}
async function getEntityType(listName,token){if(entityTypes[listName])return entityTypes[listName];const r=await fetch(`${SP}/_api/web/lists/getbytitle('${enc(listName)}')?$select=ListItemEntityTypeFullName`,{headers:{Accept:'application/json;odata=verbose',Authorization:'Bearer '+token}});if(!r.ok)throw new Error('type '+listName+': '+r.status);const d=await r.json();entityTypes[listName]=d.d.ListItemEntityTypeFullName;return entityTypes[listName];}
async function spGet(listName,fieldMap,filter=''){const token=await getToken();let url=`${SP}/_api/web/lists/getbytitle('${enc(listName)}')/items?$top=5000`;if(filter)url+=`&$filter=${encodeURIComponent(filter)}`;const r=await fetch(url,{headers:{Accept:'application/json;odata=verbose',Authorization:'Bearer '+token}});if(!r.ok)throw new Error(`GET ${listName}: ${r.status}`);const d=await r.json();return fieldMap?mapFrom(d.d.results,fieldMap):d.d.results;}
async function spPost(listName,item){const token=await getToken();const[type,digest]=await Promise.all([getEntityType(listName,token),getDigest(token)]);const r=await fetch(`${SP}/_api/web/lists/getbytitle('${enc(listName)}')/items`,{method:'POST',headers:{'Accept':'application/json;odata=verbose','Content-Type':'application/json;odata=verbose','X-RequestDigest':digest,'Authorization':'Bearer '+token},body:JSON.stringify({__metadata:{type},...item})});if(!r.ok){const t=await r.text();let m=`HTTP ${r.status}`;try{m=JSON.parse(t).error?.message?.value||m;}catch{}throw new Error(m);}return r.json();}
async function spPatch(listName,spId,item){const token=await getToken();const[type,digest]=await Promise.all([getEntityType(listName,token),getDigest(token)]);const r=await fetch(`${SP}/_api/web/lists/getbytitle('${enc(listName)}')/items(${spId})`,{method:'POST',headers:{'Accept':'application/json;odata=verbose','Content-Type':'application/json;odata=verbose','X-RequestDigest':digest,'Authorization':'Bearer '+token,'X-HTTP-Method':'MERGE','IF-MATCH':'*'},body:JSON.stringify({__metadata:{type},...item})});if(!r.ok){const t=await r.text();let m=`HTTP ${r.status}`;try{m=JSON.parse(t).error?.message?.value||m;}catch{}throw new Error(m);}}
async function spDelete(listName,spId){const token=await getToken();const digest=await getDigest(token);const r=await fetch(`${SP}/_api/web/lists/getbytitle('${enc(listName)}')/items(${spId})`,{method:'POST',headers:{'Accept':'application/json;odata=verbose','X-RequestDigest':digest,'Authorization':'Bearer '+token,'X-HTTP-Method':'DELETE','IF-MATCH':'*'}});if(!r.ok&&r.status!==204){const t=await r.text();let m=`HTTP ${r.status}`;try{m=JSON.parse(t).error?.message?.value||m;}catch{}throw new Error(m);}}

async function spGetAttachments(listName,itemId){const token=await getToken();const r=await fetch(`${SP}/_api/web/lists/getbytitle('${enc(listName)}')/items(${itemId})/AttachmentFiles`,{headers:{Accept:'application/json;odata=verbose',Authorization:'Bearer '+token}});if(!r.ok)return[];const d=await r.json();return d.d.results;}
async function spAttach(listName,itemId,fileName,fileData){const token=await getToken();const digest=await getDigest(token);const r=await fetch(`${SP}/_api/web/lists/getbytitle('${enc(listName)}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(fileName)}')`,{method:'POST',headers:{Accept:'application/json;odata=verbose','X-RequestDigest':digest,Authorization:'Bearer '+token},body:fileData});if(!r.ok){const t=await r.text();let m=`HTTP ${r.status}`;try{m=JSON.parse(t).error?.message?.value||m;}catch{}throw new Error(m);}return r.json();}
async function spDeleteAttachment(listName,itemId,fileName){const token=await getToken();const digest=await getDigest(token);await fetch(`${SP}/_api/web/lists/getbytitle('${enc(listName)}')/items(${itemId})/AttachmentFiles('${encodeURIComponent(fileName)}')`,{method:'POST',headers:{Accept:'application/json;odata=verbose','X-RequestDigest':digest,Authorization:'Bearer '+token,'X-HTTP-Method':'DELETE','IF-MATCH':'*'}});}

// ─── Experiment status ────────────────────────────────────────────────────
const EXP_STATUS_COLORS=['#2CD2A7','#FF9933','#424242'];
const EXP_STATUS_LABELS=['geplant','heute','fertig'];
function getExpStatus(id){return parseInt(localStorage.getItem('expStatus-'+id)||'2');}
function cycleExpStatus(id,event){event.stopPropagation();const s=(getExpStatus(id)+1)%3;localStorage.setItem('expStatus-'+id,s);const btn=document.getElementById('sb-'+id);if(btn){btn.style.background=EXP_STATUS_COLORS[s];btn.title=EXP_STATUS_LABELS[s];}}

// ─── MPa color ────────────────────────────────────────────────────────────
function computeMpaRanges(){mpaRanges={};allMat.forEach(m=>{const v=parseFloat(calcMpa(m.Laenge_mm,m.Breite_mm,m.Kraft_N));if(isNaN(v))return;const id=m.Lagerfolge_ID||'';if(!mpaRanges[id])mpaRanges[id]={min:v,max:v};mpaRanges[id].min=Math.min(mpaRanges[id].min,v);mpaRanges[id].max=Math.max(mpaRanges[id].max,v);});}
function mpaStyle(mpa,lafoId){const r=mpaRanges[lafoId];if(!r||r.max===r.min)return'';const pct=(parseFloat(mpa)-r.min)/(r.max-r.min);if((lafoId||'').toUpperCase().includes('WET')){return`background:hsl(280,${Math.round(pct*65)}%,${Math.round(88-pct*8)}%);`;}else{return`background:hsl(120,${Math.round(pct*65)}%,${Math.round(88-pct*6)}%);`;}}

// ─── Sort helpers ─────────────────────────────────────────────────────────
function setSort(pane,col){
  const s=sortState[pane];
  if(s.col===col)s.dir*=-1; else{s.col=col;s.dir=col==='Datum'||col==='MWMpa'||col==='MWSC'?-1:1;}
  updateSortHeaders(pane);
  if(pane==='exp')filterExp();else if(pane==='mat')filterMat();else filterRes();
}
function updateSortHeaders(pane){
  const s=sortState[pane];
  document.querySelectorAll(`[id^="si-${pane}-"]`).forEach(el=>{
    const col=el.id.replace(`si-${pane}-`,'');
    const th=el.closest('th');
    if(col===s.col){el.textContent=s.dir===1?'▲':'▼';th.classList.add('sort-active');}
    else{el.textContent='↕';th.classList.remove('sort-active');}
  });
}
function applySort(rows,pane,numericCols=[]){
  const s=sortState[pane];
  return [...rows].sort((a,b)=>{
    let va=a[s.col]??'',vb=b[s.col]??'';
    if(numericCols.includes(s.col)){va=parseFloat(va)||0;vb=parseFloat(vb)||0;}
    if(va<vb)return -s.dir;if(va>vb)return s.dir;return 0;
  });
}

// ─── Helpers ──────────────────────────────────────────────────────────────
function fillSelect(id,data,valKey,labelFn){const sel=document.getElementById(id);const first=sel.options[0];sel.innerHTML='';sel.appendChild(first);data.forEach(d=>{const o=document.createElement('option');o.value=d[valKey];o.textContent=labelFn(d);sel.appendChild(o);});}
function showTab(id,btn){document.querySelectorAll('.pane').forEach(p=>p.classList.remove('active'));document.querySelectorAll('nav button').forEach(b=>b.classList.remove('active'));document.getElementById('pane-'+id).classList.add('active');btn.classList.add('active');}

// ─── Multi-select ─────────────────────────────────────────────────────────
function buildMs(panelId,data,valKey,labelFn,selectedSet,onChangeFn){const panel=document.getElementById(panelId);const clearBtn=panel.querySelector('.ms-clear');panel.innerHTML='';panel.appendChild(clearBtn);data.forEach(d=>{const val=String(d[valKey]);const lbl=document.createElement('label');lbl.className='ms-item';lbl.innerHTML=`<input type="checkbox" value="${esc(val)}" onchange="${onChangeFn}(this)"> ${esc(labelFn(d))}`;panel.appendChild(lbl);});}
function toggleMs(type){const panel=document.getElementById('ms-'+type+'-panel');const isOpen=panel.classList.contains('open');document.querySelectorAll('.ms-panel').forEach(p=>p.classList.remove('open'));if(!isOpen)panel.classList.add('open');}
document.addEventListener('click',e=>{if(!e.target.closest('.ms-wrap'))document.querySelectorAll('.ms-panel').forEach(p=>p.classList.remove('open'));});
function updateMsBtn(type,set){const btn=document.getElementById('ms-'+type+'-btn');const count=document.getElementById('ms-'+type+'-count');if(set.size===0){btn.classList.remove('active');count.innerHTML='';}else{btn.classList.add('active');count.innerHTML=`<span class="ms-tag">${set.size}</span> `;}}
function clearMs(type){
  const map={'proj':selectedProj,'pers':selectedPers,'lafo':selectedLafo,'ort':selectedOrt,'res-proj':selectedResProj,'res-pers':selectedResPers,'res-lafo':selectedResLafo};
  map[type].clear();
  document.querySelectorAll(`#ms-${type}-panel input[type=checkbox]`).forEach(cb=>cb.checked=false);
  updateMsBtn(type,map[type]);
  if(type==='proj'||type==='pers')filterExp();else if(type==='lafo')filterMat();else if(type==='ort')filterChem();else filterRes();
}
function onProjChange(cb){cb.checked?selectedProj.add(cb.value):selectedProj.delete(cb.value);updateMsBtn('proj',selectedProj);filterExp();}
function onPersChange(cb){cb.checked?selectedPers.add(cb.value):selectedPers.delete(cb.value);updateMsBtn('pers',selectedPers);filterExp();}
function onLafoChange(cb){cb.checked?selectedLafo.add(cb.value):selectedLafo.delete(cb.value);updateMsBtn('lafo',selectedLafo);filterMat();}
function onOrtChange(cb){cb.checked?selectedOrt.add(cb.value):selectedOrt.delete(cb.value);updateMsBtn('ort',selectedOrt);filterChem();}
function onResProjChange(cb){cb.checked?selectedResProj.add(cb.value):selectedResProj.delete(cb.value);updateMsBtn('res-proj',selectedResProj);filterRes();}
function onResPersChange(cb){cb.checked?selectedResPers.add(cb.value):selectedResPers.delete(cb.value);updateMsBtn('res-pers',selectedResPers);filterRes();}
function onResLafoChange(cb){cb.checked?selectedResLafo.add(cb.value):selectedResLafo.delete(cb.value);updateMsBtn('res-lafo',selectedResLafo);filterRes();}

function checkDupExpId(input){
  const val=input.value;
  const isDup=val&&!editingExp&&allExp.some(e=>e.Experiment_ID===val);
  input.style.borderColor=isDup?'#f59e0b':'';
  input.style.background=isDup?'#fffbeb':'';
  const hint=document.getElementById('f-exp-id-hint');
  if(hint){if(isDup)hint.textContent='⚠ Diese ID existiert bereits – bitte ändern!';else if(hint.textContent.startsWith('⚠'))hint.textContent='';}
}
function setTextCol(col,btn){activeTextCol=col;document.querySelectorAll('.col-toggle button').forEach(b=>b.classList.remove('active'));btn.classList.add('active');document.getElementById('text-col-header').textContent=col;filterExp();}

// ─── Load all ─────────────────────────────────────────────────────────────
async function loadAll(){
  setStatus('Lade…');
  try{
    [personen,projekte,lagerfolgen,allExp,allMat]=await Promise.all([
      spGet(LIST.personen,FIELDS.personen),spGet(LIST.projekte,FIELDS.projekte),
      spGet(LIST.lagerfolgen,FIELDS.lagerfolgen),spGet(LIST.experimente,FIELDS.experimente),
      spGet(LIST.material,FIELDS.material),
    ]);
    allChem        = await spGet(LIST.chemikalien,        FIELDS.chemikalien).catch(()=>[]);
    allKomps       = await spGet(LIST.komponenten,        FIELDS.komponenten).catch(()=>[]);
    allSC          = await spGet(LIST.feststoffgehalt,    FIELDS.feststoffgehalt).catch(()=>[]);
    allLafoSchritte= await spGet(LIST.lagerfolgen_schritte, FIELDS.lagerfolgen_schritte).catch(()=>[]);
    allMaschinen   = await spGet(LIST.maschinen,          FIELDS.maschinen).catch(()=>[]);

    allExp.sort((a,b)=>(b.Datum||'').localeCompare(a.Datum||''));
    computeMpaRanges();

    buildMs('ms-proj-panel', projekte,    'Projekt_Kuerzel', p=>p.Projekt_Kuerzel+(p.Beschreibung?` – ${p.Beschreibung}`:''), selectedProj, 'onProjChange');
    buildMs('ms-pers-panel', personen,    'Kuerzel',         p=>`${p.Kuerzel} – ${p.Vorname} ${p.Nachname}`, selectedPers, 'onPersChange');
    buildMs('ms-lafo-panel', lagerfolgen, 'Lagerfolge_ID',   l=>`${l.Lagerfolge_ID} – ${l.Name}`, selectedLafo, 'onLafoChange');
    fillSelect('f-mat-lafo', lagerfolgen, 'Lagerfolge_ID',   l=>`${l.Lagerfolge_ID} – ${l.Name}`);
    buildMs('ms-res-proj-panel', projekte,    'Projekt_Kuerzel', p=>p.Projekt_Kuerzel+(p.Beschreibung?` – ${p.Beschreibung}`:''), selectedResProj, 'onResProjChange');
    buildMs('ms-res-pers-panel', personen,    'Kuerzel',         p=>`${p.Kuerzel} – ${p.Vorname} ${p.Nachname}`, selectedResPers, 'onResPersChange');
    buildMs('ms-res-lafo-panel', lagerfolgen, 'Lagerfolge_ID',   l=>`${l.Lagerfolge_ID} – ${l.Name}`, selectedResLafo, 'onResLafoChange');

    const orte=[...new Set(allChem.map(c=>c.Ort).filter(Boolean))].sort().map(o=>({Ort:o}));
    buildMs('ms-ort-panel', orte, 'Ort', o=>o.Ort, selectedOrt, 'onOrtChange');

    updateSortHeaders('exp');updateSortHeaders('mat');updateSortHeaders('res');
    cachedErgebnisse=computeErgebnisse();
    filterExp();filterMat();filterRes();filterChem();filterLafo();filterMasch();filterProj();setStatus('');
  }catch(e){
    setStatus('Fehler: '+e.message);console.error(e);
    document.getElementById('exp-tbody').innerHTML=`<tr><td colspan="8" class="state">${esc(e.message)}</td></tr>`;
    document.getElementById('mat-tbody').innerHTML=`<tr><td colspan="10" class="state">${esc(e.message)}</td></tr>`;
    document.getElementById('res-tbody').innerHTML=`<tr><td colspan="12" class="state">${esc(e.message)}</td></tr>`;
    document.getElementById('lafo-tbody').innerHTML=`<tr><td colspan="5" class="state">${esc(e.message)}</td></tr>`;
    document.getElementById('mach-tbody').innerHTML=`<tr><td colspan="6" class="state">${esc(e.message)}</td></tr>`;
  }
}

// ─── Experiments list ──────────────────────────────────────────────────────
function filterExp(){
  const q=document.getElementById('exp-search').value.toLowerCase();
  const df=document.getElementById('exp-date-from').value;
  const dt=document.getElementById('exp-date-to').value;
  let rows=allExp.filter(e=>
    (!selectedProj.size||selectedProj.has(e.Projekt_Kuerzel))&&
    (!selectedPers.size||selectedPers.has(e.Person_Kuerzel))&&
    (!df||fmtDate(e.Datum)>=df)&&(!dt||fmtDate(e.Datum)<=dt)&&
    (!q||`${e.Experiment_ID} ${e.Projekttitel} ${e.Beschreibung} ${e.Beobachtungen}`.toLowerCase().includes(q))
  );
  rows=applySort(rows,'exp');
  document.getElementById('exp-count').textContent=rows.length+' Einträge';
  document.getElementById('exp-tbody').innerHTML=rows.length
    ?rows.map(e=>{
        const eid=esc(e.Experiment_ID);
        const text=esc(e[activeTextCol]||'');
        const presseK=e.Presse||'';
        const pd=e.Pressdruck!=null?e.Pressdruck:'';
        const pt=e.Presstemperatur!=null?e.Presstemperatur:'';
        const pz=e.Presszeit!=null?e.Presszeit:'';
        const hasPress=presseK||pd!==''||pt!==''||pz!=='';
        const paramShort=hasPress?[presseK,pd,pt,pz].filter(v=>v!=='').join(' '):'';
        const mach=allMaschinen.find(m=>m.Kuerzel===presseK);
        const paramTip=hasPress?[mach?`Presse: ${mach.Name}`:presseK?`Presse: ${presseK}`:'',pd!==''?`${pd} N/mm²`:'',pt!==''?`${pt} °C`:'',pz!==''?`${pz} min`:''].filter(Boolean).join(', '):'';
        const st=getExpStatus(e.Experiment_ID);
        return `<tr class="exp-row" id="erow-${eid}" onclick="toggleExpRow('${eid}')">
          <td><span class="badge" onclick="openDetail(event,'${eid}')">${eid}</span></td>
          <td class="exp-titel-td" id="etit-${eid}" title="${esc(e.Projekttitel)}">${esc(e.Projekttitel)}</td>
          <td style="white-space:nowrap">${fmtDate(e.Datum)}</td>
          <td>${esc(e.Person_Kuerzel)}</td>
          <td class="text-cell collapsed" id="tc-${eid}">${text}</td>
          <td style="white-space:nowrap;font-size:11px;color:#666;font-family:monospace" title="${esc(paramTip)}">${esc(paramShort)}</td>
          <td style="width:28px;text-align:center" onclick="event.stopPropagation()"><button id="sb-${eid}" class="status-btn" style="background:${EXP_STATUS_COLORS[st]}" title="${EXP_STATUS_LABELS[st]}" onclick="cycleExpStatus('${eid}',event)"></button></td>
          <td class="col-actions" onclick="event.stopPropagation()" style="width:88px">
            <button class="btn-icon" title="Bearbeiten" onclick="editExp('${eid}')">✏️</button>
            <button class="btn-icon" title="Duplizieren" onclick="dupExp('${eid}')">📋</button>
            <button class="btn-icon" title="Feststoffgehalt" onclick="openSCPanel('${eid}')">⚗️</button>
          </td>
        </tr>`;
      }).join('')
    :'<tr><td colspan="8" class="state">Keine Einträge gefunden.</td></tr>';
}

function toggleExpRow(expId){const row=document.getElementById('erow-'+expId);const tc=document.getElementById('tc-'+expId);if(!row)return;const exp=!row.classList.contains('expanded');row.classList.toggle('expanded',exp);if(tc){tc.classList.toggle('collapsed',!exp);tc.classList.toggle('expanded-text',exp);}}
function toggleAllWrap(on){document.getElementById('pane-exp').classList.toggle('wrap-mode',on);}

// ─── Material list ─────────────────────────────────────────────────────────
function filterMat(){
  const q=document.getElementById('mat-search').value.toLowerCase();
  const df=document.getElementById('mat-date-from').value;
  const dt=document.getElementById('mat-date-to').value;
  let rows=allMat.filter(m=>{
    const exp=allExp.find(e=>e.Experiment_ID===m.Experiment_ID);
    const datum=fmtDate(exp?.Datum||'');
    const titel=(exp?.Projekttitel||'').toLowerCase();
    return (!selectedLafo.size||selectedLafo.has(m.Lagerfolge_ID))&&
      (!df||datum>=df)&&(!dt||datum<=dt)&&
      (!q||((m.Experiment_ID||'').toLowerCase().includes(q)||titel.includes(q)));
  });
  // Add computed fields for sorting
  rows=rows.map(m=>{const exp=allExp.find(e=>e.Experiment_ID===m.Experiment_ID);return{...m,_mpa:parseFloat(calcMpa(m.Laenge_mm,m.Breite_mm,m.Kraft_N))||0,_datum:exp?.Datum||''};});
  rows=applySort(rows,'mat',['_mpa']);
  document.getElementById('mat-count').textContent=rows.length+' Einträge';
  document.getElementById('mat-tbody').innerHTML=rows.length
    ?rows.map(m=>{
        const mpa=calcMpa(m.Laenge_mm,m.Breite_mm,m.Kraft_N);
        const style=mpa?mpaStyle(mpa,m.Lagerfolge_ID):'';
        const hb=m.Holzbruch_pct!=null?Math.round(m.Holzbruch_pct*100)+'%':'';
        const lxb=(m.Laenge_mm!=null&&m.Breite_mm!=null)?`${m.Laenge_mm}×${m.Breite_mm}`:'';
        const eid=esc(m.Experiment_ID);
        const exp=allExp.find(e=>e.Experiment_ID===m.Experiment_ID);
        const titel=esc(exp?.Projekttitel||'');
        return `<tr>
          <td style="white-space:nowrap;width:110px"><span class="badge" onclick="openDetail(event,'${eid}')">${eid}</span></td>
          <td style="max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px;color:#555" title="${titel}">${titel}</td>
          <td style="white-space:nowrap;font-size:12px;color:#666">${fmtDate(m._datum)}</td>
          <td style="text-align:right"><span class="mpa-chip" style="${style}">${mpa??''}</span></td>
          <td style="text-align:right;font-size:12px">${hb}</td>
          <td style="white-space:nowrap;font-size:12px">${esc(m.Lagerfolge_ID)}</td>
          <td style="text-align:right;font-size:12px;white-space:nowrap">${lxb}</td>
          <td style="text-align:right;font-size:12px">${m.Kraft_N??''}</td>
          <td class="truncate" style="font-size:12px">${esc(m.Kommentar)}</td>
          <td class="col-actions"><button class="btn-icon" title="Bearbeiten" onclick="editMat(${m._spId})">✏️</button></td>
        </tr>`;
      }).join('')
    :'<tr><td colspan="10" class="state">Keine Einträge gefunden.</td></tr>';
}

// ─── Ergebnisse ───────────────────────────────────────────────────────────
function calcSC(leer,ein,end){if(ein==null||ein===''||leer==null||end==null||end==='')return null;const v=(parseFloat(end)-parseFloat(leer))/parseFloat(ein)*100;return isNaN(v)?null:v;}

function computeErgebnisse(){
  const groups={};
  allMat.forEach(m=>{
    const mpa=parseFloat(calcMpa(m.Laenge_mm,m.Breite_mm,m.Kraft_N));
    if(isNaN(mpa))return;
    const key=`${m.Experiment_ID}||${m.Lagerfolge_ID||''}`;
    if(!groups[key])groups[key]={Experiment_ID:m.Experiment_ID,Lagerfolge_ID:m.Lagerfolge_ID||'',mpaVals:[],hbVals:[]};
    groups[key].mpaVals.push(mpa);
    if(m.Holzbruch_pct!=null)groups[key].hbVals.push(m.Holzbruch_pct*100);
  });
  // SC pro Experiment (unabhängig von Lagerfolge)
  const scByExp={};
  allSC.forEach(s=>{
    const v=calcSC(s.Leergewicht_g,s.Einwaage_g,s.Endgewicht_g);
    if(v==null)return;
    if(!scByExp[s.Experiment_ID])scByExp[s.Experiment_ID]=[];
    scByExp[s.Experiment_ID].push(v);
  });
  function scStats(expId){
    const vals=scByExp[expId]||[];
    if(!vals.length)return{nSC:0,MWSC:null,StdAbwSC:null};
    const n=vals.length,mean=vals.reduce((a,b)=>a+b,0)/n;
    const std=n>1?Math.sqrt(vals.reduce((a,b)=>a+(b-mean)**2,0)/(n-1)):0;
    return{nSC:n,MWSC:mean,StdAbwSC:std};
  }
  return Object.values(groups).map(g=>{
    const n=g.mpaVals.length;
    const mean=g.mpaVals.reduce((a,b)=>a+b,0)/n;
    const variance=n>1?g.mpaVals.reduce((a,b)=>a+(b-mean)**2,0)/(n-1):0;
    const exp=allExp.find(e=>e.Experiment_ID===g.Experiment_ID);
    const sc=scStats(g.Experiment_ID);
    return{
      Experiment_ID:g.Experiment_ID,Lagerfolge_ID:g.Lagerfolge_ID,
      Titel:exp?.Projekttitel||'',Datum:exp?.Datum||'',
      Person:exp?.Person_Kuerzel||'',Projekt:exp?.Projekt_Kuerzel||'',
      n,MWMpa:mean,StdAbw:Math.sqrt(variance),mpaVals:g.mpaVals,
      MWHolzbruch:g.hbVals.length?(g.hbVals.reduce((a,b)=>a+b,0)/g.hbVals.length):null,
      ...sc,
    };
  });
}

let cachedErgebnisse=[],currentResRows=[];
function filterRes(){
  const q=document.getElementById('res-search').value.toLowerCase();
  const df=document.getElementById('res-date-from').value;
  const dt=document.getElementById('res-date-to').value;
  let rows=cachedErgebnisse.filter(r=>
    (!selectedResProj.size||selectedResProj.has(r.Projekt))&&
    (!selectedResPers.size||selectedResPers.has(r.Person))&&
    (!selectedResLafo.size||selectedResLafo.has(r.Lagerfolge_ID))&&
    (!df||fmtDate(r.Datum)>=df)&&(!dt||fmtDate(r.Datum)<=dt)&&
    (!q||`${r.Experiment_ID} ${r.Titel}`.toLowerCase().includes(q))
  );
  rows=applySort(rows,'res',['MWMpa','n','MWSC']);
  currentResRows=rows;
  document.getElementById('res-count').textContent=rows.length+' Gruppen';
  document.getElementById('res-tbody').innerHTML=rows.length
    ?rows.map(r=>{
        const style=mpaStyle(r.MWMpa.toFixed(2),r.Lagerfolge_ID);
        const eid=esc(r.Experiment_ID);
        return `<tr>
          <td><span class="badge" onclick="openDetail(event,'${eid}')">${eid}</span></td>
          <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(r.Titel)}</td>
          <td style="white-space:nowrap">${fmtDate(r.Datum)}</td>
          <td>${esc(r.Person)}</td>
          <td style="font-size:12px">${esc(r.Lagerfolge_ID)}</td>
          <td style="text-align:right" class="stat-n">${r.n}</td>
          <td style="text-align:right"><span class="mpa-chip" style="${style}">${fmtDec(r.MWMpa,3)}</span></td>
          <td style="text-align:right;font-size:12px;color:#666">${fmtDec(r.StdAbw,3)}</td>
          <td style="text-align:right;font-size:12px">${r.MWHolzbruch!=null?fmtDec(r.MWHolzbruch,1)+'%':''}</td>
          <td style="text-align:right;font-size:12px;border-left:2px solid #e2e8f0">${r.nSC||''}</td>
          <td style="text-align:right;font-size:12px">${r.MWSC!=null?fmtDec(r.MWSC,2)+'%':''}</td>
          <td style="text-align:right;font-size:12px;color:#666">${r.StdAbwSC!=null&&r.nSC>1?fmtDec(r.StdAbwSC,2)+'%':''}</td>
        </tr>`;
      }).join('')
    :'<tr><td colspan="12" class="state">Keine Daten.</td></tr>';
}

// ─── Plot ─────────────────────────────────────────────────────────────────
const errDotsPlugin={
  id:'errDots',
  afterDatasetsDraw(chart){
    const ctx=chart.ctx;
    chart.data.datasets.forEach((ds,di)=>{
      const meta=chart.getDatasetMeta(di);if(meta.hidden)return;
      if(ds._errBars){
        const yScale=chart.scales[ds.yAxisID||'y'];
        ds._errBars.forEach((sd,i)=>{
          if(sd==null||isNaN(sd)||sd===0)return;
          const bar=meta.data[i];if(!bar)return;
          const mean=ds.data[i];if(mean==null)return;
          const x=bar.x,yTop=yScale.getPixelForValue(mean+sd),yBot=yScale.getPixelForValue(mean-sd),hw=4;
          ctx.save();ctx.strokeStyle='#222';ctx.lineWidth=1.5;
          ctx.beginPath();ctx.moveTo(x,yTop);ctx.lineTo(x,yBot);ctx.moveTo(x-hw,yTop);ctx.lineTo(x+hw,yTop);ctx.moveTo(x-hw,yBot);ctx.lineTo(x+hw,yBot);ctx.stroke();ctx.restore();
        });
      }
      if(ds._dots){
        const yScale=chart.scales[ds.yAxisID||'y'];
        ds._dots.forEach((vals,i)=>{
          const bar=meta.data[i];if(!bar||!vals)return;
          vals.forEach(v=>{ctx.save();ctx.fillStyle='#000';ctx.beginPath();ctx.arc(bar.x,yScale.getPixelForValue(v),3,0,Math.PI*2);ctx.fill();ctx.restore();});
        });
      }
    });
  }
};

function buildPlotConfig(rows){
  const labels=rows.map(r=>r.Experiment_ID+(r.Lagerfolge_ID?' | '+r.Lagerfolge_ID:''));
  const mpaColors=rows.map(r=>(r.Lagerfolge_ID||'').toUpperCase().includes('WET')?'rgba(120,60,170,0.72)':'rgba(30,58,95,0.75)');
  return{
    type:'bar',
    data:{
      labels,
      datasets:[
        {label:'MPa (MW)',data:rows.map(r=>r.MWMpa),backgroundColor:mpaColors,yAxisID:'y',
          _errBars:rows.map(r=>r.n>1?r.StdAbw:null),_dots:rows.map(r=>r.mpaVals||[]),order:2},
        {label:'Holzbruch %',data:rows.map(r=>r.MWHolzbruch),backgroundColor:'rgba(255,153,51,0.72)',yAxisID:'y2',order:3},
        {label:'SC %',data:rows.map(r=>r.MWSC),backgroundColor:'rgba(44,210,167,0.72)',yAxisID:'y2',order:4},
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{position:'top',labels:{font:{size:12}}},
        tooltip:{callbacks:{label(ctx){const v=ctx.parsed.y;if(v==null)return null;return`${ctx.dataset.label}: ${fmtDec(v,3)}`;}}}
      },
      scales:{
        x:{ticks:{font:{size:10},maxRotation:45}},
        y:{type:'linear',position:'left',title:{display:true,text:'MPa',font:{size:12}}},
        y2:{type:'linear',position:'right',title:{display:true,text:'%',font:{size:12}},grid:{drawOnChartArea:false}},
      }
    },
    plugins:[errDotsPlugin],
  };
}

let _plotCharts=[];
function plotRes(){
  const rows=currentResRows;
  if(!rows.length){alert('Keine Daten für den Plot vorhanden.');return;}
  const MAX_PER=10,MAX_CHARTS=5;
  if(rows.length>MAX_PER*MAX_CHARTS){alert(`Zu viele Einträge (${rows.length}) für den Plot. Maximal ${MAX_PER*MAX_CHARTS}. Bitte Filter einschränken.`);return;}
  _plotCharts.forEach(c=>c.destroy());_plotCharts=[];
  const chunks=[];for(let i=0;i<rows.length;i+=MAX_PER)chunks.push(rows.slice(i,i+MAX_PER));
  const body=document.getElementById('plot-body');body.innerHTML='';
  chunks.forEach((chunk,ci)=>{
    const wrap=document.createElement('div');wrap.className='plot-chart-wrap';
    if(chunks.length>1){const h=document.createElement('div');h.className='plot-chart-title';h.textContent=`Grafik ${ci+1} / ${chunks.length}`;wrap.appendChild(h);}
    const cw=document.createElement('div');cw.style.cssText='position:relative;height:420px';
    const canvas=document.createElement('canvas');canvas.id='plot-cv-'+ci;cw.appendChild(canvas);wrap.appendChild(cw);
    const foot=document.createElement('div');foot.className='plot-chart-foot';
    const dlBtn=document.createElement('button');dlBtn.className='btn btn-secondary btn-sm';dlBtn.textContent='↓ PNG';dlBtn.onclick=()=>downloadPlotChart(ci);
    const lbl=document.createElement('span');lbl.style.cssText='font-size:13px;margin-left:10px';lbl.textContent='Als Anhang zu:';
    const sel=document.createElement('select');sel.id='plot-sel-'+ci;sel.style.cssText='margin:0 6px;min-width:140px';
    [...new Set(chunk.map(r=>r.Experiment_ID))].forEach(id=>{const o=document.createElement('option');o.value=id;o.textContent=id;sel.appendChild(o);});
    const saveBtn=document.createElement('button');saveBtn.className='btn btn-primary btn-sm';saveBtn.textContent='📎 Speichern';saveBtn.onclick=()=>savePlotChart(ci);
    const st=document.createElement('span');st.id='plot-st-'+ci;st.style.cssText='font-size:12px;margin-left:8px';
    foot.append(dlBtn,lbl,sel,saveBtn,st);wrap.appendChild(foot);body.appendChild(wrap);
    _plotCharts.push(new Chart(canvas,buildPlotConfig(chunk)));
  });
  document.getElementById('plot-overlay').classList.add('open');
}
function closePlot(){document.getElementById('plot-overlay').classList.remove('open');}
function downloadPlotChart(ci){
  const canvas=document.getElementById('plot-cv-'+ci);if(!canvas)return;
  const a=document.createElement('a');a.download=`ligaro_plot_${ci+1}_${new Date().toISOString().slice(0,10)}.png`;a.href=canvas.toDataURL('image/png');a.click();
}
async function savePlotChart(ci){
  const canvas=document.getElementById('plot-cv-'+ci),sel=document.getElementById('plot-sel-'+ci),st=document.getElementById('plot-st-'+ci);
  if(!canvas||!sel)return;
  const expId=sel.value,exp=allExp.find(e=>e.Experiment_ID===expId);
  if(!exp){st.textContent='Experiment nicht gefunden.';return;}
  st.textContent='Speichert…';st.style.color='';
  try{
    const blob=await new Promise(res=>canvas.toBlob(res,'image/png'));
    const buf=await blob.arrayBuffer();
    const fname=`plot_${expId}_${new Date().toISOString().slice(0,19).replace(/:/g,'-')}.png`;
    await spAttach(LIST.experimente,exp._spId,fname,buf);
    st.textContent='✓ Gespeichert!';st.style.color='#1a6b3c';setTimeout(()=>{st.textContent='';st.style.color='';},3000);
  }catch(e){st.textContent='Fehler: '+e.message;st.style.color='#c53030';}
}

function exportExcel(){
  const q=document.getElementById('res-search').value.toLowerCase();
  const df=document.getElementById('res-date-from').value;
  const dt=document.getElementById('res-date-to').value;
  let rows=cachedErgebnisse.filter(r=>
    (!selectedResProj.size||selectedResProj.has(r.Projekt))&&
    (!selectedResPers.size||selectedResPers.has(r.Person))&&
    (!selectedResLafo.size||selectedResLafo.has(r.Lagerfolge_ID))&&
    (!df||fmtDate(r.Datum)>=df)&&(!dt||fmtDate(r.Datum)<=dt)&&
    (!q||`${r.Experiment_ID} ${r.Titel}`.toLowerCase().includes(q))
  );
  rows=applySort(rows,'res',['MWMpa','n','MWSC']);
  const data=[
    ['Experiment-ID','Titel','Datum','Person','Protokoll','n','MW MPa','StdAbw (MPa)','MW Holzbruch (%)','n SC','MW SC%','StdAbw SC%'],
    ...rows.map(r=>[r.Experiment_ID,r.Titel,fmtDate(r.Datum),r.Person,r.Lagerfolge_ID,r.n,
      parseFloat(r.MWMpa.toFixed(3)),parseFloat(r.StdAbw.toFixed(3)),
      r.MWHolzbruch!=null?parseFloat(r.MWHolzbruch.toFixed(1)):'',
      r.nSC||0,r.MWSC!=null?parseFloat(r.MWSC.toFixed(2)):'',r.StdAbwSC!=null&&r.nSC>1?parseFloat(r.StdAbwSC.toFixed(2)):'']),
  ];
  const ws=XLSX.utils.aoa_to_sheet(data);
  ws['!cols']=[{wch:16},{wch:30},{wch:12},{wch:8},{wch:14},{wch:4},{wch:10},{wch:12},{wch:16},{wch:5},{wch:10},{wch:12}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Ergebnisse');
  XLSX.writeFile(wb,'LIGARO_Ergebnisse.xlsx');
}

// ─── Chemicals list ────────────────────────────────────────────────────────
const FUELLSTAND_CFG={
  voll:        {color:'#22c55e',bg:'#dcfce7',label:'Voll'},
  mittel:      {color:'#ca8a04',bg:'#fef9c3',label:'Mittel'},
  fast_leer:   {color:'#ea580c',bg:'#ffedd5',label:'Fast leer'},
  nachbestellen:{color:'#dc2626',bg:'#fee2e2',label:'Nachbestellen'},
};
function fuellstandChip(val){
  if(!val)return'';
  const cfg=FUELLSTAND_CFG[val];
  if(!cfg)return esc(val);
  return `<span style="display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700;background:${cfg.bg};color:${cfg.color}">${cfg.label}</span>`;
}
function filterChem(){
  const q=document.getElementById('chem-search').value.toLowerCase();
  const rows=allChem.filter(c=>
    (!selectedOrt.size||selectedOrt.has(c.Ort))&&
    (!q||`${c.Chemikalie_Name} ${c.Hersteller} ${c.Ort} ${c.IUPAC}`.toLowerCase().includes(q))
  ).sort((a,b)=>(a.Chemikalie_Name||'').localeCompare(b.Chemikalie_Name||''));
  document.getElementById('chem-count').textContent=rows.length+' Einträge';
  document.getElementById('chem-tbody').innerHTML=rows.length
    ?rows.map(c=>{
        const cid=(c.Chemikalie_Name||'').replace(/[^a-z0-9]/gi,'_')+c._spId;
        const spId=c._spId;
        const usedIn=[...new Set(allKomps.filter(k=>k.Chemikalie_Name===c.Chemikalie_Name).map(k=>k.Experiment_ID).filter(Boolean))];
        const hasExpand=usedIn.length>0;
        const expandRow=hasExpand?`<tr class="chem-expand-row" id="chem-exp-${cid}"><td colspan="11"><div class="chem-expand-inner"><span class="extra-label">Verwendet in ${usedIn.length} Experiment(en)</span><div class="chem-exp-list">${usedIn.map(id=>`<span class="badge" onclick="openDetail(event,'${esc(id)}')">${esc(id)}</span>`).join('')}</div></div></td></tr>`:'';
        const chevron=hasExpand?`<span style="font-size:10px;opacity:.5" id="chev-${cid}">▶</span>`:'';
        const flChip=fuellstandChip(c.Fuellstand);
        return `<tr class="chem-row" onclick="${hasExpand?`toggleChemRow('${cid}')`:''}" id="crow-${cid}">
          <td style="width:20px;padding:9px 5px 9px 13px">${chevron}</td>
          <td><strong title="${esc(c.IUPAC||'')}">${esc(c.Chemikalie_Name)}</strong></td>
          <td style="font-family:monospace;font-size:12px">${esc(c.Formel)}</td>
          <td>${esc(c.Hersteller)}</td>
          <td style="text-align:right">${esc(c.Menge)}</td>
          <td>${esc(c.Einheit)}</td>
          <td style="text-align:center">${flChip}</td>
          <td>${esc(c.Ort)}</td>
          <td>${esc(c.Herkunft)}</td>
          <td class="truncate">${esc(c.Kommentar)}</td>
          <td class="col-actions" onclick="event.stopPropagation()" style="width:60px">
            <button class="btn-icon" title="SDS / Anhänge" onclick="openSDS(${spId},'${esc(c.Chemikalie_Name)}')">📄</button>
            <button class="btn-icon" title="Bearbeiten" onclick="editChem(${spId})">✏️</button>
          </td>
        </tr>${expandRow}`;
      }).join('')
    :allChem.length===0
      ?'<tr><td colspan="11" class="state">Chemikalien-Liste noch nicht in SharePoint vorhanden.</td></tr>'
      :'<tr><td colspan="11" class="state">Keine Einträge gefunden.</td></tr>';
}
function toggleChemRow(cid){const row=document.getElementById('crow-'+cid);const expRow=document.getElementById('chem-exp-'+cid);const chev=document.getElementById('chev-'+cid);if(!expRow)return;const open=!expRow.classList.contains('show');expRow.classList.toggle('show',open);row.classList.toggle('expanded',open);if(chev)chev.textContent=open?'▼':'▶';}

// ─── Lagerfolgen list ──────────────────────────────────────────────────────
function filterLafo(){
  const q=(document.getElementById('lafo-search')?.value||'').toLowerCase();
  const rows=lagerfolgen.filter(l=>!q||`${l.Lagerfolge_ID} ${l.Name} ${l.Norm}`.toLowerCase().includes(q));
  const el=document.getElementById('lafo-tbody');if(!el)return;
  document.getElementById('lafo-count').textContent=rows.length+' Einträge';
  el.innerHTML=rows.length
    ?rows.map(l=>{
        const lid=(l.Lagerfolge_ID||'').replace(/[^a-z0-9]/gi,'_');
        const steps=[...allLafoSchritte.filter(s=>s.Lagerfolge_ID===l.Lagerfolge_ID)].sort((a,b)=>(parseInt(a.Schritt_Nr)||0)-(parseInt(b.Schritt_Nr)||0));
        const hasSteps=steps.length>0;
        const chevron=hasSteps?`<span style="font-size:10px;opacity:.5" id="lchev-${lid}">▶</span>`:'';
        const stepsRow=hasSteps?`<tr class="chem-expand-row" id="lafo-exp-${lid}"><td colspan="5"><div class="chem-expand-inner"><table class="modal-table" style="font-size:12px"><thead><tr><th>Nr.</th><th>Behandlung</th><th>Medium</th><th style="text-align:right">Temp. (°C)</th><th style="text-align:right">rH (%)</th><th style="text-align:right">Dauer (h)</th><th>Kommentar</th></tr></thead><tbody>${steps.map(s=>`<tr><td>${esc(s.Schritt_Nr)}</td><td>${esc(s.Behandlung)}</td><td>${esc(s.Medium)}</td><td style="text-align:right">${esc(s.Temperatur_C)}</td><td style="text-align:right">${esc(s.RH_pct)}</td><td style="text-align:right">${esc(s.Dauer_h)}</td><td style="font-size:11px;color:#666">${esc(s.Kommentar)}</td></tr>`).join('')}</tbody></table></div></td></tr>`:'';
        return`<tr class="chem-row" onclick="${hasSteps?`toggleLafoRow('${lid}')`:''}" id="lrow-${lid}"><td style="width:20px;padding:9px 5px 9px 13px">${chevron}</td><td><strong>${esc(l.Lagerfolge_ID)}</strong></td><td>${esc(l.Name)}</td><td style="font-size:12px;color:#555">${esc(l.Norm)}</td><td class="col-actions" onclick="event.stopPropagation()"><button class="btn-icon" title="Bearbeiten" onclick="editLafo('${esc(l.Lagerfolge_ID)}')">✏️</button></td></tr>${stepsRow}`;
      }).join('')
    :'<tr><td colspan="5" class="state">Keine Lagerfolgen gefunden.</td></tr>';
}
function toggleLafoRow(lid){const expRow=document.getElementById('lafo-exp-'+lid);const chev=document.getElementById('lchev-'+lid);if(!expRow)return;const open=!expRow.classList.contains('show');expRow.classList.toggle('show',open);if(chev)chev.textContent=open?'▼':'▶';}

// ─── Maschinen list ────────────────────────────────────────────────────────
function filterMasch(){
  const q=(document.getElementById('mach-search')?.value||'').toLowerCase();
  const rows=allMaschinen.filter(m=>!q||`${m.Kuerzel} ${m.Name} ${m.Hersteller} ${m.Typ}`.toLowerCase().includes(q));
  const el=document.getElementById('mach-tbody');if(!el)return;
  document.getElementById('mach-count').textContent=rows.length+' Maschinen';
  el.innerHTML=rows.length
    ?rows.map(m=>`<tr><td style="font-weight:700;white-space:nowrap;width:48px">${esc(m.Kuerzel)}</td><td>${esc(m.Name)}</td><td>${esc(m.Hersteller)}</td><td style="font-size:12px;color:#555">${esc(m.Typ)}</td><td style="font-size:12px;color:#888">${esc(m.Kommentar)}</td><td class="col-actions"><button class="btn-icon" title="Bearbeiten" onclick="editMach(${m._spId})">✏️</button></td></tr>`).join('')
    :allMaschinen.length===0
      ?'<tr><td colspan="6" class="state">Maschinen-Liste noch nicht in SharePoint vorhanden.</td></tr>'
      :'<tr><td colspan="6" class="state">Keine Einträge gefunden.</td></tr>';
}

// ─── Projekte list ────────────────────────────────────────────────────────
function filterProj(){
  const q=(document.getElementById('proj-search')?.value||'').toLowerCase();
  const rows=[...projekte].sort((a,b)=>a.Projekt_Kuerzel.localeCompare(b.Projekt_Kuerzel)).filter(p=>!q||`${p.Projekt_Kuerzel} ${p.Beschreibung}`.toLowerCase().includes(q));
  const el=document.getElementById('proj-tbody');if(!el)return;
  document.getElementById('proj-count').textContent=rows.length+' Projekte';
  const expCounts={};allExp.forEach(e=>{const p=e.Projekt_Kuerzel;if(p)expCounts[p]=(expCounts[p]||0)+1;});
  el.innerHTML=rows.length
    ?rows.map(p=>`<tr><td style="font-weight:700;white-space:nowrap;width:80px">${esc(p.Projekt_Kuerzel)}</td><td>${esc(p.Beschreibung)}</td><td style="text-align:right;color:#888;font-size:12px">${expCounts[p.Projekt_Kuerzel]||0}</td><td class="col-actions"><button class="btn-icon" title="Bearbeiten" onclick="editProj('${esc(p.Projekt_Kuerzel)}')">✏️</button></td></tr>`).join('')
    :'<tr><td colspan="4" class="state">Keine Projekte gefunden.</td></tr>';
}
function editProj(kuerzel){
  const item=projekte.find(p=>p.Projekt_Kuerzel===kuerzel);if(!item)return;
  editingProj=item;
  document.getElementById('overlay').classList.add('open');
  document.getElementById('panel-proj').classList.add('open');
  activePanel='proj';
  document.getElementById('f-proj-kuerzel').value=item.Projekt_Kuerzel;
  document.getElementById('f-proj-beschreibung').value=item.Beschreibung||'';
  document.getElementById('proj-alert').innerHTML='';
}
async function saveProj(){
  if(!editingProj)return;
  const beschreibung=document.getElementById('f-proj-beschreibung').value.trim();
  const alertEl=document.getElementById('proj-alert');
  try{
    await spPatch(LIST.projekte,editingProj._spId,mapTo({Projekt_Kuerzel:editingProj.Projekt_Kuerzel,Beschreibung:beschreibung},FIELDS.projekte));
    editingProj.Beschreibung=beschreibung;
    buildMs('ms-proj-panel',projekte,'Projekt_Kuerzel',p=>p.Projekt_Kuerzel+(p.Beschreibung?` – ${p.Beschreibung}`:''),selectedProj,'onProjChange');
    buildMs('ms-res-proj-panel',projekte,'Projekt_Kuerzel',p=>p.Projekt_Kuerzel+(p.Beschreibung?` – ${p.Beschreibung}`:''),selectedResProj,'onResProjChange');
    filterProj();
    alertEl.innerHTML='<div class="alert alert-ok">Gespeichert!</div>';setTimeout(closePanel,900);
  }catch(e){alertEl.innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
}

// ─── Lagerfolge form ───────────────────────────────────────────────────────
function initLafoForm(item){
  document.getElementById('lafo-alert').innerHTML='';
  document.getElementById('panel-lafo-title').textContent=item?'Lagerfolge bearbeiten':'Neue Lagerfolge';
  document.getElementById('btn-del-lafo').style.display=item?'':'none';
  document.getElementById('f-lafo-id').disabled=!!item;
  deletedLafoSchritte=[];
  document.getElementById('lafo-schritt-tbody').innerHTML='';
  lafoSchrittIdx=0;
  if(item){
    document.getElementById('f-lafo-id').value=item.Lagerfolge_ID||'';
    document.getElementById('f-lafo-name').value=item.Name||'';
    document.getElementById('f-lafo-norm').value=item.Norm||'';
    document.getElementById('f-lafo-anwendung').value=item.Anwendung||'';
    const steps=[...allLafoSchritte.filter(s=>s.Lagerfolge_ID===item.Lagerfolge_ID)].sort((a,b)=>(parseInt(a.Schritt_Nr)||0)-(parseInt(b.Schritt_Nr)||0));
    steps.forEach(s=>addLafoSchritt(s));
  } else {
    ['f-lafo-id','f-lafo-name','f-lafo-norm','f-lafo-anwendung'].forEach(id=>document.getElementById(id).value='');
    addLafoSchritt();
  }
}
function editLafo(lafoId){
  const item=lagerfolgen.find(l=>l.Lagerfolge_ID===lafoId);if(!item)return;
  editingLafo=item;
  document.getElementById('overlay').classList.add('open');
  document.getElementById('panel-lafo').classList.add('open');
  activePanel='lafo';initLafoForm(item);
}
function addLafoSchritt(item){
  const i=lafoSchrittIdx++;
  const tr=document.createElement('tr');
  tr.id='lafo-sr-'+i;
  if(item?._spId)tr.dataset.spid=String(item._spId);
  const spIdArg=item?._spId?','+item._spId:'';
  tr.innerHTML=`<td><input type="number" id="lsn-${i}" value="${item?.Schritt_Nr||i+1}" min="1" step="1" style="width:44px"></td><td><input type="text" id="lsb-${i}" value="${esc(item?.Behandlung||'')}" style="width:110px"></td><td><input type="text" id="lsm-${i}" value="${esc(item?.Medium||'')}" style="width:70px"></td><td><input type="number" id="lst-${i}" value="${item?.Temperatur_C??''}" step="0.1" style="width:54px"></td><td><input type="number" id="lsr-${i}" value="${item?.RH_pct??''}" step="1" style="width:54px"></td><td><input type="number" id="lsd-${i}" value="${item?.Dauer_h??''}" step="0.5" style="width:54px"></td><td><input type="text" id="lsk-${i}" value="${esc(item?.Kommentar||'')}" style="width:130px"></td><td><button class="del-btn" onclick="removeLafoSchritt(${i}${spIdArg})">×</button></td>`;
  document.getElementById('lafo-schritt-tbody').appendChild(tr);
}
function removeLafoSchritt(i,spId){document.getElementById('lafo-sr-'+i)?.remove();if(spId)deletedLafoSchritte.push(spId);}
async function saveLafo(){
  const alertEl=document.getElementById('lafo-alert'),btn=document.getElementById('btn-save-lafo');
  const id=document.getElementById('f-lafo-id').value.trim();
  const name=document.getElementById('f-lafo-name').value.trim();
  if(!id||!name){alertEl.innerHTML='<div class="alert alert-err">ID und Name sind Pflichtfelder.</div>';return;}
  if(!editingLafo&&lagerfolgen.some(l=>l.Lagerfolge_ID===id)){alertEl.innerHTML=`<div class="alert alert-err">ID „${id}" existiert bereits.</div>`;return;}
  btn.disabled=true;btn.textContent='Speichert…';alertEl.innerHTML='';
  try{
    const int={Lagerfolge_ID:id,Name:name,Norm:document.getElementById('f-lafo-norm').value.trim(),Anwendung:document.getElementById('f-lafo-anwendung').value.trim()};
    const sp=mapTo(int,FIELDS.lagerfolgen);
    if(editingLafo){await spPatch(LIST.lagerfolgen,editingLafo._spId,sp);Object.assign(editingLafo,int);}
    else{const saved=await spPost(LIST.lagerfolgen,sp);lagerfolgen.push({...int,_spId:saved.d.Id});lagerfolgen.sort((a,b)=>a.Lagerfolge_ID.localeCompare(b.Lagerfolge_ID));}
    // Delete removed steps
    await Promise.all(deletedLafoSchritte.map(sid=>spDelete(LIST.lagerfolgen_schritte,sid)));
    allLafoSchritte=allLafoSchritte.filter(s=>!deletedLafoSchritte.includes(s._spId));
    deletedLafoSchritte=[];
    // Save / update steps
    const schrittRows=[];
    document.querySelectorAll('#lafo-schritt-tbody tr').forEach(tr=>{
      const i=tr.id.replace('lafo-sr-','');
      const spId=tr.dataset.spid?parseInt(tr.dataset.spid):null;
      const s={Lagerfolge_ID:id,Schritt_Nr:document.getElementById('lsn-'+i)?.value||'',Behandlung:document.getElementById('lsb-'+i)?.value.trim()||'',Medium:document.getElementById('lsm-'+i)?.value.trim()||'',Temperatur_C:document.getElementById('lst-'+i)?.value||null,RH_pct:document.getElementById('lsr-'+i)?.value||null,Dauer_h:document.getElementById('lsd-'+i)?.value||null,Kommentar:document.getElementById('lsk-'+i)?.value.trim()||''};
      schrittRows.push({int:s,sp:mapTo(s,FIELDS.lagerfolgen_schritte),spId});
    });
    for(const r of schrittRows){
      if(r.spId){await spPatch(LIST.lagerfolgen_schritte,r.spId,r.sp);const ex=allLafoSchritte.find(s=>s._spId===r.spId);if(ex)Object.assign(ex,r.int);}
      else{const saved=await spPost(LIST.lagerfolgen_schritte,r.sp);allLafoSchritte.push({...r.int,_spId:saved.d.Id});}
    }
    buildMs('ms-lafo-panel',lagerfolgen,'Lagerfolge_ID',l=>`${l.Lagerfolge_ID} – ${l.Name}`,selectedLafo,'onLafoChange');
    fillSelect('f-mat-lafo',lagerfolgen,'Lagerfolge_ID',l=>`${l.Lagerfolge_ID} – ${l.Name}`);
    buildMs('ms-res-lafo-panel',lagerfolgen,'Lagerfolge_ID',l=>`${l.Lagerfolge_ID} – ${l.Name}`,selectedResLafo,'onResLafoChange');
    filterLafo();
    alertEl.innerHTML='<div class="alert alert-ok">Gespeichert!</div>';setTimeout(closePanel,900);
  }catch(e){alertEl.innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
  finally{btn.disabled=false;btn.textContent='Speichern';}
}
async function deleteLafo(){
  if(!editingLafo||!confirm(`Lagerfolge „${editingLafo.Lagerfolge_ID}" wirklich löschen?`))return;
  try{
    const steps=allLafoSchritte.filter(s=>s.Lagerfolge_ID===editingLafo.Lagerfolge_ID);
    await Promise.all(steps.map(s=>spDelete(LIST.lagerfolgen_schritte,s._spId)));
    allLafoSchritte=allLafoSchritte.filter(s=>s.Lagerfolge_ID!==editingLafo.Lagerfolge_ID);
    await spDelete(LIST.lagerfolgen,editingLafo._spId);
    lagerfolgen=lagerfolgen.filter(l=>l._spId!==editingLafo._spId);
    buildMs('ms-lafo-panel',lagerfolgen,'Lagerfolge_ID',l=>`${l.Lagerfolge_ID} – ${l.Name}`,selectedLafo,'onLafoChange');
    fillSelect('f-mat-lafo',lagerfolgen,'Lagerfolge_ID',l=>`${l.Lagerfolge_ID} – ${l.Name}`);
    buildMs('ms-res-lafo-panel',lagerfolgen,'Lagerfolge_ID',l=>`${l.Lagerfolge_ID} – ${l.Name}`,selectedResLafo,'onResLafoChange');
    filterLafo();closePanel();
  }catch(e){document.getElementById('lafo-alert').innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
}

// ─── Maschine form ─────────────────────────────────────────────────────────
function initMachForm(item){
  document.getElementById('mach-alert').innerHTML='';
  document.getElementById('panel-mach-title').textContent=item?'Maschine bearbeiten':'Neue Maschine';
  document.getElementById('btn-del-mach').style.display=item?'':'none';
  if(item){
    document.getElementById('f-mach-name').value=item.Name||'';
    document.getElementById('f-mach-hersteller').value=item.Hersteller||'';
    document.getElementById('f-mach-typ').value=item.Typ||'';
    document.getElementById('f-mach-kuerzel').value=item.Kuerzel||'';
    document.getElementById('f-mach-komm').value=item.Kommentar||'';
  } else {
    ['f-mach-name','f-mach-hersteller','f-mach-kuerzel','f-mach-komm'].forEach(id=>document.getElementById(id).value='');
    document.getElementById('f-mach-typ').value='';
  }
}
function editMach(spId){const item=allMaschinen.find(m=>m._spId===spId);if(!item)return;editingMach=item;document.getElementById('overlay').classList.add('open');document.getElementById('panel-mach').classList.add('open');activePanel='mach';initMachForm(item);}
async function saveMach(){
  const alertEl=document.getElementById('mach-alert'),btn=document.getElementById('btn-save-mach');
  const name=document.getElementById('f-mach-name').value.trim();
  const kuerzel=document.getElementById('f-mach-kuerzel').value.trim();
  if(!name||!kuerzel){alertEl.innerHTML='<div class="alert alert-err">Name und Kürzel sind Pflichtfelder.</div>';return;}
  btn.disabled=true;btn.textContent='Speichert…';alertEl.innerHTML='';
  const int={Name:name,Hersteller:document.getElementById('f-mach-hersteller').value.trim(),Typ:document.getElementById('f-mach-typ').value,Maschinen_ID:'',Kuerzel:kuerzel,Kommentar:document.getElementById('f-mach-komm').value.trim()};
  try{
    const sp=mapTo(int,FIELDS.maschinen);
    if(editingMach){await spPatch(LIST.maschinen,editingMach._spId,sp);Object.assign(editingMach,int);}
    else{const saved=await spPost(LIST.maschinen,sp);allMaschinen.push({...int,_spId:saved.d.Id});allMaschinen.sort((a,b)=>(a.Kuerzel||'').localeCompare(b.Kuerzel||''));}
    filterMasch();
    alertEl.innerHTML='<div class="alert alert-ok">Gespeichert!</div>';setTimeout(closePanel,900);
  }catch(e){alertEl.innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
  finally{btn.disabled=false;btn.textContent='Speichern';}
}
async function deleteMach(){
  if(!editingMach||!confirm(`Maschine „${editingMach.Name}" wirklich löschen?`))return;
  try{await spDelete(LIST.maschinen,editingMach._spId);allMaschinen=allMaschinen.filter(m=>m._spId!==editingMach._spId);filterMasch();closePanel();}
  catch(e){document.getElementById('mach-alert').innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
}

// ─── Detail modal ──────────────────────────────────────────────────────────
async function openDetail(event,expId){
  event.stopPropagation();
  const expForTitle=allExp.find(e=>e.Experiment_ID===expId);
  document.getElementById('detail-title').textContent=expId+(expForTitle?.Projekttitel?' – '+expForTitle.Projekttitel:'');
  document.getElementById('detail-body').innerHTML='<div class="state">Lade…</div>';
  document.getElementById('detail-overlay').classList.add('open');
  try{
    let komps=[],kompsNote='';
    try{komps=await spGet(LIST.komponenten,FIELDS.komponenten,`Title eq '${expId}'`);}
    catch(e){kompsNote=e.message.includes('404')?'<p class="modal-warn">Komponenten-Liste noch nicht in SharePoint vorhanden.</p>':`<p class="modal-warn">Fehler: ${esc(e.message)}</p>`;}
    const mats=allMat.filter(m=>m.Experiment_ID===expId);
    const exp=allExp.find(e=>e.Experiment_ID===expId);
    const attachments=exp?await spGetAttachments(LIST.experimente,exp._spId).catch(()=>[]):[];
    let html='';
    html+='<div class="modal-section"><h3>Zusammensetzung</h3>';
    if(kompsNote)html+=kompsNote;
    html+='<table class="modal-table"><thead><tr><th>Komponente</th><th>Hersteller</th><th style="text-align:right">Menge</th><th>Einheit</th><th>Rolle</th><th></th></tr></thead><tbody id="komp-detail-tbody">';
    if(komps.length)komps.forEach(k=>{html+=dKompRow(k,expId);});
    else html+=`<tr id="komp-empty-row"><td colspan="6" class="modal-empty">Keine Komponenten eingetragen.</td></tr>`;
    html+='</tbody></table>';
    html+=`<button class="btn btn-secondary btn-sm" style="margin-top:8px" onclick="dKompAdd('${esc(expId)}')">+ Komponente</button>`;
    html+='</div>';
    html+='<div class="modal-section"><h3>Materialprüfung</h3>';
    if(mats.length){html+='<table class="modal-table"><thead><tr><th>Protokoll</th><th style="text-align:right">MPa</th><th style="text-align:right">Holzbruch %</th><th style="text-align:right">L × B</th><th style="text-align:right">Kraft (N)</th><th>Kommentar</th></tr></thead><tbody>';mats.forEach(m=>{const mpa=calcMpa(m.Laenge_mm,m.Breite_mm,m.Kraft_N);const style=mpa?mpaStyle(mpa,m.Lagerfolge_ID):'';const hb=m.Holzbruch_pct!=null?Math.round(m.Holzbruch_pct*100)+'%':'';const lxb=(m.Laenge_mm!=null&&m.Breite_mm!=null)?`${m.Laenge_mm} × ${m.Breite_mm}`:'';html+=`<tr><td>${esc(m.Lagerfolge_ID)}</td><td style="text-align:right"><span class="mpa-chip" style="${style}">${mpa!=null?String(mpa).replace('.',','):''}</span></td><td style="text-align:right">${hb}</td><td style="text-align:right">${lxb}</td><td style="text-align:right">${m.Kraft_N??''}</td><td>${esc(m.Kommentar)}</td></tr>`;});html+='</tbody></table>';}
    else html+='<p class="modal-empty">Keine Messungen vorhanden.</p>';
    html+='</div>';
    const scs=allSC.filter(s=>s.Experiment_ID===expId);
    html+='<div class="modal-section"><h3>Feststoffgehalt</h3>';
    if(scs.length){html+='<table class="modal-table"><thead><tr><th>Probe</th><th style="text-align:right">Leergewicht (g)</th><th style="text-align:right">Einwaage (g)</th><th style="text-align:right">Endgewicht (g)</th><th style="text-align:right">SC%</th><th>Kommentar</th><th></th></tr></thead><tbody>';scs.forEach(s=>{const sc=calcSC(s.Leergewicht_g,s.Einwaage_g,s.Endgewicht_g);html+=`<tr><td>${esc(s.Probe)}</td><td style="text-align:right">${s.Leergewicht_g??''}</td><td style="text-align:right">${s.Einwaage_g??''}</td><td style="text-align:right">${s.Endgewicht_g??''}</td><td style="text-align:right;font-weight:600">${sc!=null?fmtDec(sc,2)+'%':''}</td><td style="font-size:12px;color:#666">${esc(s.Kommentar)}</td><td><button class="btn-icon" title="Bearbeiten" onclick="editSC(${s._spId});closeDetailDirect()">✏️</button></td></tr>`;});html+='</tbody></table>';}
    else html+='<p class="modal-empty">Keine SC-Messungen vorhanden. <button class="btn btn-secondary btn-sm" onclick="openSCPanel(\''+esc(expId)+'\');closeDetailDirect()">SC erfassen</button></p>';
    html+='</div>';
    if(exp&&(exp.Beschreibung||exp.Beobachtungen||exp.Kommentar)){html+='<div class="modal-section"><h3>Notizen</h3>';if(exp.Beschreibung)html+=`<div class="extra-label">Beschreibung</div><p style="font-size:13px;margin-bottom:9px;white-space:pre-wrap">${esc(exp.Beschreibung)}</p>`;if(exp.Beobachtungen)html+=`<div class="extra-label">Beobachtungen</div><p style="font-size:13px;margin-bottom:9px;white-space:pre-wrap">${esc(exp.Beobachtungen)}</p>`;if(exp.Kommentar)html+=`<div class="extra-label">Kommentar</div><p style="font-size:13px;white-space:pre-wrap">${esc(exp.Kommentar)}</p>`;html+='</div>';}
    if(exp){
      html+='<div class="modal-section" id="attach-section"><h3>Anhänge</h3>';
      html+=renderAttachmentList(attachments,LIST.experimente,exp._spId);
      html+=`<label class="btn btn-secondary btn-sm" style="cursor:pointer;margin-top:8px;display:inline-block">📎 Datei hochladen<input type="file" multiple style="display:none" onchange="uploadAttachments(this,'${LIST.experimente}',${exp._spId})"></label>`;
      html+='</div>';
    }
    document.getElementById('detail-body').innerHTML=html;
  }catch(e){document.getElementById('detail-body').innerHTML=`<p class="modal-empty">Fehler: ${esc(e.message)}</p>`;}
}
function closeDetail(e){if(e.target===document.getElementById('detail-overlay'))closeDetailDirect();}
function closeDetailDirect(){document.getElementById('detail-overlay').classList.remove('open');}

// ─── Detail modal: Komponenten bearbeiten ──────────────────────────────────
function dKompRow(k,expId){
  const name=esc(k.Chemikalie_Name||k.Experiment_Ref||k.Komponente_Name||'–');
  return `<tr id="komp-row-${k._spId}" data-spid="${k._spId}" data-expid="${esc(expId)}" data-name="${name}" data-her="${esc(k.Hersteller||'')}" data-menge="${k.Menge!=null?k.Menge:''}" data-einheit="${esc(k.Einheit||'')}" data-rolle="${esc(k.Rolle||'')}">
    <td>${name}</td><td>${esc(k.Hersteller)}</td><td style="text-align:right">${k.Menge!=null?k.Menge:''}</td><td>${esc(k.Einheit)}</td><td>${esc(k.Rolle)}</td>
    <td style="white-space:nowrap"><button class="btn-icon" title="Bearbeiten" onclick="dKompEdit(${k._spId})">✏️</button><button class="btn-icon" title="Löschen" onclick="dKompDel(${k._spId})">🗑️</button></td></tr>`;
}
function dKompEditInputs(id,name,her,menge,einheit,rolle,expId){
  return `<td><input type="text" id="kd-n-${id}" value="${name}" style="width:120px" list="komp-chem-list"></td>
    <td><input type="text" id="kd-h-${id}" value="${her}" style="width:85px"></td>
    <td><input type="number" id="kd-m-${id}" value="${menge}" style="width:55px"></td>
    <td><input type="text" id="kd-e-${id}" value="${einheit}" style="width:55px"></td>
    <td><input type="text" id="kd-r-${id}" value="${rolle}" style="width:75px"></td>
    <td style="white-space:nowrap">
      <button class="btn btn-primary btn-sm" onclick="dKompSave(${id},'${expId}')">💾</button>
      <button class="btn btn-secondary btn-sm" onclick="dKompCancel(${id})">✕</button></td>`;
}
function dKompEdit(spId){
  const row=document.getElementById('komp-row-'+spId);if(!row)return;
  row.innerHTML=dKompEditInputs(spId,row.dataset.name,row.dataset.her,row.dataset.menge,row.dataset.einheit,row.dataset.rolle,row.dataset.expid);
}
function dKompCancel(spId){
  const row=document.getElementById('komp-row-'+spId);if(!row)return;
  const d=row.dataset;
  row.innerHTML=`<td>${d.name}</td><td>${d.her}</td><td style="text-align:right">${d.menge}</td><td>${d.einheit}</td><td>${d.rolle}</td><td style="white-space:nowrap"><button class="btn-icon" onclick="dKompEdit(${spId})">✏️</button><button class="btn-icon" onclick="dKompDel(${spId})">🗑️</button></td>`;
}
async function dKompSave(spId,expId){
  const nameVal=(document.getElementById('kd-n-'+spId)?.value||'').trim();if(!nameVal)return;
  const chem=allChem.find(c=>c.Chemikalie_Name===nameVal);
  const kint={Experiment_ID:expId,Quelle_Typ:chem?'Chemikalie':'Sonstiges',Chemikalie_Name:chem?nameVal:null,Komponente_Name:chem?null:nameVal,Experiment_Ref:null,Hersteller:(document.getElementById('kd-h-'+spId)?.value||'').trim(),Menge:document.getElementById('kd-m-'+spId)?.value!==''?parseFloat(document.getElementById('kd-m-'+spId)?.value):null,Einheit:(document.getElementById('kd-e-'+spId)?.value||'').trim(),Rolle:(document.getElementById('kd-r-'+spId)?.value||'').trim()};
  try{
    await spPatch(LIST.komponenten,spId,mapTo(kint,FIELDS.komponenten));
    const cached=allKomps.find(k=>k._spId===spId);if(cached)Object.assign(cached,kint);
    const row=document.getElementById('komp-row-'+spId);if(!row)return;
    const name=esc(kint.Chemikalie_Name||kint.Komponente_Name||'');
    row.dataset.name=name;row.dataset.her=esc(kint.Hersteller);row.dataset.menge=kint.Menge!=null?kint.Menge:'';row.dataset.einheit=esc(kint.Einheit);row.dataset.rolle=esc(kint.Rolle);
    row.innerHTML=`<td>${name}</td><td>${esc(kint.Hersteller)}</td><td style="text-align:right">${kint.Menge!=null?kint.Menge:''}</td><td>${esc(kint.Einheit)}</td><td>${esc(kint.Rolle)}</td><td style="white-space:nowrap"><button class="btn-icon" onclick="dKompEdit(${spId})">✏️</button><button class="btn-icon" onclick="dKompDel(${spId})">🗑️</button></td>`;
  }catch(e){alert('Fehler: '+e.message);}
}
async function dKompDel(spId){
  if(!confirm('Komponente löschen?'))return;
  try{
    await spDelete(LIST.komponenten,spId);
    allKomps=allKomps.filter(k=>k._spId!==spId);
    document.getElementById('komp-row-'+spId)?.remove();
    const tbody=document.getElementById('komp-detail-tbody');
    if(tbody&&!tbody.querySelector('tr[id^="komp-row-"]'))tbody.innerHTML='<tr id="komp-empty-row"><td colspan="6" class="modal-empty">Keine Komponenten eingetragen.</td></tr>';
  }catch(e){alert('Fehler: '+e.message);}
}
function dKompAdd(expId){
  document.getElementById('komp-empty-row')?.remove();
  const tbody=document.getElementById('komp-detail-tbody');if(!tbody)return;
  const tid='new'+Date.now();
  const tr=document.createElement('tr');tr.id='komp-row-'+tid;tr.dataset.expid=expId;
  tr.innerHTML=`<td><input type="text" id="kd-n-${tid}" placeholder="Name/Chemikalie" style="width:120px" list="komp-chem-list"></td>
    <td><input type="text" id="kd-h-${tid}" placeholder="Hersteller" style="width:85px"></td>
    <td><input type="number" id="kd-m-${tid}" placeholder="Menge" style="width:55px"></td>
    <td><input type="text" id="kd-e-${tid}" placeholder="Einheit" style="width:55px"></td>
    <td><input type="text" id="kd-r-${tid}" placeholder="Rolle" style="width:75px"></td>
    <td style="white-space:nowrap">
      <button class="btn btn-primary btn-sm" onclick="dKompCreate('${tid}','${esc(expId)}')">💾</button>
      <button class="btn btn-secondary btn-sm" onclick="document.getElementById('komp-row-${tid}')?.remove()">✕</button></td>`;
  tbody.appendChild(tr);
  document.getElementById('kd-n-'+tid)?.focus();
}
async function dKompCreate(tid,expId){
  const nameVal=(document.getElementById('kd-n-'+tid)?.value||'').trim();if(!nameVal)return;
  const chem=allChem.find(c=>c.Chemikalie_Name===nameVal);
  const kint={Experiment_ID:expId,Quelle_Typ:chem?'Chemikalie':'Sonstiges',Chemikalie_Name:chem?nameVal:null,Komponente_Name:chem?null:nameVal,Experiment_Ref:null,Hersteller:(document.getElementById('kd-h-'+tid)?.value||'').trim(),Menge:document.getElementById('kd-m-'+tid)?.value!==''?parseFloat(document.getElementById('kd-m-'+tid)?.value):null,Einheit:(document.getElementById('kd-e-'+tid)?.value||'').trim(),Rolle:(document.getElementById('kd-r-'+tid)?.value||'').trim()};
  try{
    const saved=await spPost(LIST.komponenten,mapTo(kint,FIELDS.komponenten));
    const newSpId=saved.d.Id;
    allKomps.push({...kint,_spId:newSpId});
    const row=document.getElementById('komp-row-'+tid);if(!row)return;
    row.id='komp-row-'+newSpId;row.dataset.spid=newSpId;
    const name=esc(kint.Chemikalie_Name||kint.Komponente_Name||'');
    row.dataset.name=name;row.dataset.her=esc(kint.Hersteller);row.dataset.menge=kint.Menge!=null?kint.Menge:'';row.dataset.einheit=esc(kint.Einheit);row.dataset.rolle=esc(kint.Rolle);
    row.innerHTML=`<td>${name}</td><td>${esc(kint.Hersteller)}</td><td style="text-align:right">${kint.Menge!=null?kint.Menge:''}</td><td>${esc(kint.Einheit)}</td><td>${esc(kint.Rolle)}</td><td style="white-space:nowrap"><button class="btn-icon" onclick="dKompEdit(${newSpId})">✏️</button><button class="btn-icon" onclick="dKompDel(${newSpId})">🗑️</button></td>`;
  }catch(e){alert('Fehler: '+e.message);}
}

// ─── Attachments ───────────────────────────────────────────────────────────
function renderAttachmentList(attachments,listName,itemId){
  if(!attachments.length)return'<p class="modal-empty" id="attach-empty">Noch keine Anhänge.</p>';
  return'<div id="attach-list">'+attachments.map(a=>{
    const name=esc(a.FileName);
    const url=esc(a.ServerRelativeUrl);
    return`<div style="display:flex;align-items:center;gap:8px;padding:4px 0;border-bottom:1px solid #f0f2f5">
      <span style="font-size:16px">${fileIcon(a.FileName)}</span>
      <a href="${SP_HOST}${url}" target="_blank" style="font-size:13px;flex:1;color:#1e3a5f;text-decoration:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${name}</a>
      <button class="btn-icon" title="Löschen" onclick="deleteAttachment('${listName}',${itemId},'${name}',this)">🗑️</button>
    </div>`;
  }).join('')+'</div>';
}
function fileIcon(name){const ext=(name||'').split('.').pop().toLowerCase();if(['jpg','jpeg','png','gif','webp'].includes(ext))return'🖼️';if(ext==='pdf')return'📄';return'📎';}
async function uploadAttachments(input,listName,itemId){
  const files=[...input.files];if(!files.length)return;
  const section=document.getElementById('attach-section');
  const status=document.createElement('p');status.style.fontSize='12px';status.style.color='#888';status.textContent=`Lade ${files.length} Datei(en)…`;section.appendChild(status);
  input.disabled=true;
  try{
    for(const f of files){
      const buf=await f.arrayBuffer();
      await spAttach(listName,itemId,f.name,buf);
    }
    const attachments=await spGetAttachments(listName,itemId);
    const existing=document.getElementById('attach-list')||document.getElementById('attach-empty');
    if(existing){const newList=document.createElement('div');newList.innerHTML=renderAttachmentList(attachments,listName,itemId);existing.replaceWith(...newList.childNodes);}
    status.textContent=`✓ ${files.length} Datei(en) hochgeladen.`;status.style.color='#1a6b3c';
  }catch(e){status.textContent='Fehler: '+e.message;status.style.color='#c53030';}
  finally{input.value='';input.disabled=false;}
}
async function deleteAttachment(listName,itemId,fileName,btn){
  if(!confirm(`„${fileName}" wirklich löschen?`))return;
  btn.disabled=true;
  try{
    await spDeleteAttachment(listName,itemId,fileName);
    const attachments=await spGetAttachments(listName,itemId);
    const existing=document.getElementById('attach-list')||document.getElementById('attach-empty');
    if(existing){const newList=document.createElement('div');newList.innerHTML=renderAttachmentList(attachments,listName,itemId);existing.replaceWith(...newList.childNodes);}
  }catch(e){alert('Fehler: '+e.message);btn.disabled=false;}
}

// ─── SDS (Sicherheitsdatenblatt) ───────────────────────────────────────────
async function openSDS(spId,chemName){
  const attachments=await spGetAttachments(LIST.chemikalien,spId).catch(()=>[]);
  const sds=attachments.filter(a=>a.FileName.toLowerCase().includes('sds')||a.FileName.toLowerCase().includes('sicherheit')||a.FileName.toLowerCase().endsWith('.pdf'));
  if(sds.length===1&&!attachments.find(a=>a!==sds[0])){
    window.open(SP+sds[0].ServerRelativeUrl,'_blank');return;
  }
  // show quick picker
  const modal=document.getElementById('sds-modal');
  document.getElementById('sds-modal-name').textContent=chemName;
  document.getElementById('sds-list').innerHTML=attachments.length
    ?attachments.map(a=>`<div style="display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid #f0f2f5">
        <span>${fileIcon(a.FileName)}</span>
        <a href="${SP_HOST}${esc(a.ServerRelativeUrl)}" target="_blank" style="flex:1;font-size:13px;color:#1e3a5f">${esc(a.FileName)}</a>
        <button class="btn-icon" onclick="deleteAttachment('${LIST.chemikalien}',${spId},'${esc(a.FileName)}',this);this.closest('[id=sds-list]').querySelector('a')?.click()">🗑️</button>
      </div>`).join('')
    :'<p class="modal-empty">Noch keine Anhänge.</p>';
  document.getElementById('sds-expid').value='';
  document.getElementById('sds-spid').value=spId;
  modal.classList.add('open');
}
function closeSDS(){document.getElementById('sds-modal').classList.remove('open');}
async function uploadSDS(){
  const input=document.getElementById('sds-file-input');const files=[...input.files];if(!files.length)return;
  const spId=parseInt(document.getElementById('sds-spid').value);
  const btn=document.getElementById('btn-upload-sds');btn.disabled=true;btn.textContent='Lädt…';
  try{
    for(const f of files){await spAttach(LIST.chemikalien,spId,f.name,await f.arrayBuffer());}
    const attachments=await spGetAttachments(LIST.chemikalien,spId);
    document.getElementById('sds-list').innerHTML=attachments.map(a=>`<div style="display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid #f0f2f5">
      <span>${fileIcon(a.FileName)}</span>
      <a href="${SP_HOST}${esc(a.ServerRelativeUrl)}" target="_blank" style="flex:1;font-size:13px;color:#1e3a5f">${esc(a.FileName)}</a>
      <button class="btn-icon" onclick="deleteAttachment('${LIST.chemikalien}',${spId},'${esc(a.FileName)}',this)">🗑️</button>
    </div>`).join('');
    input.value='';
  }catch(e){alert('Fehler: '+e.message);}
  finally{btn.disabled=false;btn.textContent='Hochladen';}
}

// ─── Komponenten (in Experiment-Formular) ─────────────────────────────────
function addKompRow(item){
  const i=kompIdx++;
  const tr=document.createElement('tr');
  tr.id='komp-row-'+i;
  if(item?._spId)tr.dataset.spid=String(item._spId);
  const spIdArg=item?._spId?','+item._spId:'';
  const nameVal=esc(item?.Chemikalie_Name||item?.Komponente_Name||'');
  tr.innerHTML=`<td><input type="text" id="kn-${i}" value="${nameVal}" list="komp-chem-list" autocomplete="off" oninput="onKompNameInput(${i},this.value)" style="width:140px"></td><td><input type="text" id="kh-${i}" value="${esc(item?.Hersteller||'')}" style="width:100px"></td><td><input type="number" id="km-${i}" value="${item?.Menge??''}" min="0" step="any" style="width:60px"></td><td><input type="text" id="ke-${i}" value="${esc(item?.Einheit||'')}" style="width:50px"></td><td><input type="text" id="kr-${i}" value="${esc(item?.Rolle||'')}" style="width:80px"></td><td><button class="del-btn" onclick="removeKompRow(${i}${spIdArg})">×</button></td>`;
  document.getElementById('komp-tbody').appendChild(tr);
}
function removeKompRow(i,spId){document.getElementById('komp-row-'+i)?.remove();if(spId)deletedKomps.push(spId);}
function onKompNameInput(i,val){const chem=allChem.find(c=>c.Chemikalie_Name===val);if(chem){const hEl=document.getElementById('kh-'+i);if(hEl&&!hEl.value)hEl.value=chem.Hersteller||'';}}

// ─── Panels ───────────────────────────────────────────────────────────────
let activePanel=null;
function openPanel(type){document.getElementById('overlay').classList.add('open');document.getElementById('panel-'+type).classList.add('open');activePanel=type;if(type==='exp'){editingExp=null;initExpForm();}else if(type==='mat'){editingMat=null;initMatForm();}else if(type==='chem'){editingChem=null;initChemForm();}else if(type==='sc'){editingSC=null;initSCForm(null);}else if(type==='lafo'){editingLafo=null;initLafoForm();}else if(type==='mach'){editingMach=null;initMachForm();}}
function closePanel(){document.getElementById('overlay').classList.remove('open');if(activePanel){document.getElementById('panel-'+activePanel).classList.remove('open');activePanel=null;}editingExp=null;editingMat=null;editingChem=null;editingSC=null;editingLafo=null;editingMach=null;editingProj=null;}

// ─── Experiment form ───────────────────────────────────────────────────────
function initExpForm(item){
  document.getElementById('exp-alert').innerHTML='';
  document.getElementById('panel-exp-title').textContent=item?'Experiment bearbeiten':'Neues Experiment';
  document.getElementById('btn-del-exp').style.display=item?'':'none';
  document.getElementById('f-exp-id').disabled=!!item;
  fillSelectProj();
  fillSelectPers();
  fillSelectPresse();
  document.getElementById('new-proj-fields').style.display='none';
  document.getElementById('new-pers-fields').style.display='none';
  // Refresh chemical datalist for komponenten
  const kompDl=document.getElementById('komp-chem-list');
  if(kompDl){kompDl.innerHTML='';allChem.forEach(c=>{const o=document.createElement('option');o.value=c.Chemikalie_Name;kompDl.appendChild(o);});}
  // Reset komponenten rows
  deletedKomps=[];kompIdx=0;
  const kompTbody=document.getElementById('komp-tbody');if(kompTbody)kompTbody.innerHTML='';
  if(item){
    document.getElementById('f-exp-id').value=item.Experiment_ID||'';
    document.getElementById('f-exp-proj').value=item.Projekt_Kuerzel||'';
    document.getElementById('f-exp-datum').value=fmtDate(item.Datum)||'';
    document.getElementById('f-exp-person').value=item.Person_Kuerzel||'';
    document.getElementById('f-exp-titel').value=item.Projekttitel||'';
    document.getElementById('f-exp-beschreibung').value=item.Beschreibung||'';
    document.getElementById('f-exp-beob').value=item.Beobachtungen||'';
    document.getElementById('f-exp-komm').value=item.Kommentar||'';
    document.getElementById('f-exp-presse').value=item.Presse||'';
    document.getElementById('f-exp-pressdruck').value=item.Pressdruck??'';
    document.getElementById('f-exp-presstemperatur').value=item.Presstemperatur??'';
    document.getElementById('f-exp-presszeit').value=item.Presszeit??'';
    document.getElementById('f-exp-id-hint').textContent='';
    // Load existing Komponenten
    allKomps.filter(k=>k.Experiment_ID===item.Experiment_ID).forEach(k=>addKompRow(k));
  } else {
    document.getElementById('f-exp-datum').value=new Date().toISOString().slice(0,10);
    document.getElementById('f-exp-proj').value='';
    document.getElementById('f-exp-person').value='';
    ['f-exp-id','f-exp-titel','f-exp-beschreibung','f-exp-beob','f-exp-komm'].forEach(id=>document.getElementById(id).value='');
    document.getElementById('f-exp-presse').value='P1';
    document.getElementById('f-exp-pressdruck').value='0.7';
    document.getElementById('f-exp-presstemperatur').value='140';
    document.getElementById('f-exp-presszeit').value='10';
    document.getElementById('f-exp-id-hint').textContent='';
  }
}
function editExp(expId){const item=allExp.find(e=>e.Experiment_ID===expId);if(!item)return;editingExp=item;document.getElementById('overlay').classList.add('open');document.getElementById('panel-exp').classList.add('open');activePanel='exp';initExpForm(item);}
function dupExp(expId){
  const item=allExp.find(e=>e.Experiment_ID===expId);if(!item)return;
  const prefix=item.Experiment_ID.replace(/-.*/, '');
  const nums=allExp.filter(e=>e.Experiment_ID?.startsWith(prefix+'-')).map(e=>parseInt(e.Experiment_ID.split('-')[1])).filter(n=>Number.isFinite(n));
  const newId=`${prefix}-${String(nums.length?Math.max(...nums)+1:1).padStart(3,'0')}`;
  editingExp=null;
  document.getElementById('overlay').classList.add('open');document.getElementById('panel-exp').classList.add('open');activePanel='exp';
  initExpForm(null);
  document.getElementById('f-exp-id').value=newId;
  document.getElementById('f-exp-id-hint').textContent=`Dupliziert von ${item.Experiment_ID}`;
  document.getElementById('f-exp-proj').value=item.Projekt_Kuerzel||'';
  document.getElementById('f-exp-person').value=item.Person_Kuerzel||'';
  document.getElementById('f-exp-titel').value=item.Projekttitel||'';
  document.getElementById('f-exp-beschreibung').value=item.Beschreibung||'';
  document.getElementById('f-exp-beob').value=item.Beobachtungen||'';
  document.getElementById('f-exp-komm').value=item.Kommentar||'';
  document.getElementById('f-exp-presse').value=item.Presse||'';
  document.getElementById('f-exp-pressdruck').value=item.Pressdruck??'';
  document.getElementById('f-exp-presstemperatur').value=item.Presstemperatur??'';
  document.getElementById('f-exp-presszeit').value=item.Presszeit??'';
}
function fillSelectPers(){
  const sel=document.getElementById('f-exp-person');
  const cur=sel.value;
  sel.innerHTML='<option value="">– wählen –</option>';
  personen.forEach(p=>{const o=document.createElement('option');o.value=p.Kuerzel;o.textContent=`${p.Kuerzel} – ${p.Vorname} ${p.Nachname}`;sel.appendChild(o);});
  const newOpt=document.createElement('option');newOpt.value='__new__';newOpt.textContent='+ Neue Person anlegen…';newOpt.style.fontWeight='600';newOpt.style.color='#1e3a5f';sel.appendChild(newOpt);
  if(cur)sel.value=cur;
}
function onPersonSelect(){
  const val=document.getElementById('f-exp-person').value;
  document.getElementById('new-pers-fields').style.display=val==='__new__'?'block':'none';
}
function fillSelectProj(){
  const sel=document.getElementById('f-exp-proj');
  const cur=sel.value;
  sel.innerHTML='<option value="">– wählen –</option>';
  projekte.forEach(p=>{const o=document.createElement('option');o.value=p.Projekt_Kuerzel;o.textContent=`${p.Projekt_Kuerzel} – ${p.Beschreibung||''}`;sel.appendChild(o);});
  const newOpt=document.createElement('option');newOpt.value='__new__';newOpt.textContent='+ Neues Projekt anlegen…';newOpt.style.fontWeight='600';newOpt.style.color='#1e3a5f';sel.appendChild(newOpt);
  if(cur)sel.value=cur;
}
function fillSelectPresse(){
  const sel=document.getElementById('f-exp-presse');
  const cur=sel.value;
  sel.innerHTML='<option value="">– keine –</option>';
  allMaschinen.filter(m=>m.Typ==='Heizpresse').forEach(m=>{const o=document.createElement('option');o.value=m.Kuerzel;o.textContent=`${m.Kuerzel} – ${m.Name}`;sel.appendChild(o);});
  if(cur)sel.value=cur;
}
function onProjSelect(){
  const val=document.getElementById('f-exp-proj').value;
  const fields=document.getElementById('new-proj-fields');
  fields.style.display=val==='__new__'?'block':'none';
  if(val!=='__new__')suggestExpId();
  else{document.getElementById('f-exp-id').value='';document.getElementById('f-exp-id-hint').textContent='';}
}
function suggestExpId(){
  if(editingExp)return;
  const sel=document.getElementById('f-exp-proj');
  const proj=sel.value==='__new__'?document.getElementById('f-new-proj-kuerzel').value.trim():sel.value;
  if(!proj)return;
  const nums=allExp.filter(e=>e.Experiment_ID?.startsWith(proj+'-')).map(e=>parseInt(e.Experiment_ID.split('-')[1])).filter(n=>Number.isFinite(n));
  const sug=`${proj}-${String(nums.length?Math.max(...nums)+1:1).padStart(3,'0')}`;
  document.getElementById('f-exp-id').value=sug;document.getElementById('f-exp-id-hint').textContent=`Nächste freie ID: ${sug}`;
}
async function saveExperiment(){
  const alertEl=document.getElementById('exp-alert'),btn=document.getElementById('btn-save-exp');
  const id=document.getElementById('f-exp-id').value.trim().toUpperCase();
  const isNewProj=document.getElementById('f-exp-proj').value==='__new__';
  const isNewPers=document.getElementById('f-exp-person').value==='__new__';
  let proj=isNewProj?document.getElementById('f-new-proj-kuerzel').value.trim().toUpperCase():document.getElementById('f-exp-proj').value;
  let person=isNewPers?document.getElementById('f-new-pers-kuerzel').value.trim().toUpperCase():document.getElementById('f-exp-person').value;
  const datum=document.getElementById('f-exp-datum').value;
  if(!id||!proj||!datum||!person){alertEl.innerHTML='<div class="alert alert-err">Bitte alle Pflichtfelder (*) ausfüllen.</div>';return;}
  if(isNewProj&&projekte.some(p=>p.Projekt_Kuerzel===proj)){alertEl.innerHTML=`<div class="alert alert-err">Projektkürzel „${proj}" existiert bereits.</div>`;return;}
  if(isNewPers&&personen.some(p=>p.Kuerzel===person)){alertEl.innerHTML=`<div class="alert alert-err">Kürzel „${person}" existiert bereits.</div>`;return;}
  if(!editingExp&&allExp.some(e=>e.Experiment_ID===id)){alertEl.innerHTML=`<div class="alert alert-err">ID „${id}" existiert bereits.</div>`;return;}
  btn.disabled=true;btn.textContent='Speichert…';alertEl.innerHTML='';
  try{
    // 1. Neues Projekt anlegen falls nötig
    if(isNewProj){
      const beschreibung=document.getElementById('f-new-proj-beschreibung').value.trim();
      await spPost(LIST.projekte,mapTo({Projekt_Kuerzel:proj,Beschreibung:beschreibung},FIELDS.projekte));
      const newP={Projekt_Kuerzel:proj,Beschreibung:beschreibung};
      projekte.push(newP);projekte.sort((a,b)=>a.Projekt_Kuerzel.localeCompare(b.Projekt_Kuerzel));
      buildMs('ms-proj-panel',projekte,'Projekt_Kuerzel',p=>p.Projekt_Kuerzel+(p.Beschreibung?` – ${p.Beschreibung}`:''),selectedProj,'onProjChange');
      buildMs('ms-res-proj-panel',projekte,'Projekt_Kuerzel',p=>p.Projekt_Kuerzel+(p.Beschreibung?` – ${p.Beschreibung}`:''),selectedResProj,'onResProjChange');
    }
    // 1b. Neue Person anlegen falls nötig
    if(isNewPers){
      const vorname=document.getElementById('f-new-pers-vorname').value.trim();
      const nachname=document.getElementById('f-new-pers-nachname').value.trim();
      await spPost(LIST.personen,mapTo({Kuerzel:person,Vorname:vorname,Nachname:nachname},FIELDS.personen));
      personen.push({Kuerzel:person,Vorname:vorname,Nachname:nachname});
      personen.sort((a,b)=>a.Kuerzel.localeCompare(b.Kuerzel));
    }
    // 2. Experiment speichern
    const pdVal=document.getElementById('f-exp-pressdruck').value;
    const ptVal=document.getElementById('f-exp-presstemperatur').value;
    const pzVal=document.getElementById('f-exp-presszeit').value;
    const int={Experiment_ID:id,Projekt_Kuerzel:proj,Datum:datum+'T00:00:00Z',Person_Kuerzel:person,Projekttitel:document.getElementById('f-exp-titel').value.trim(),Beschreibung:document.getElementById('f-exp-beschreibung').value.trim(),Beobachtungen:document.getElementById('f-exp-beob').value.trim(),Kommentar:document.getElementById('f-exp-komm').value.trim(),Presse:document.getElementById('f-exp-presse').value||null,Pressdruck:pdVal!==''?parseFloat(pdVal):null,Presstemperatur:ptVal!==''?parseFloat(ptVal):null,Presszeit:pzVal!==''?parseFloat(pzVal):null};
    const sp=mapTo(int,FIELDS.experimente);
    if(editingExp){await spPatch(LIST.experimente,editingExp._spId,sp);Object.assign(editingExp,int);}
    else{const saved=await spPost(LIST.experimente,sp);const newId=int.Experiment_ID;localStorage.setItem('expStatus-'+newId,'0');allExp.unshift({...int,_spId:saved.d.Id});}
    // Handle Komponenten
    await Promise.all(deletedKomps.map(sid=>spDelete(LIST.komponenten,sid)));
    allKomps=allKomps.filter(k=>!deletedKomps.includes(k._spId));
    deletedKomps=[];
    const kompRows=[];
    document.querySelectorAll('#komp-tbody tr').forEach(tr=>{
      const i=tr.id.replace('komp-row-','');
      const spId=tr.dataset.spid?parseInt(tr.dataset.spid):null;
      const nameVal=(document.getElementById('kn-'+i)?.value||'').trim();
      if(!nameVal)return;
      const chem=allChem.find(c=>c.Chemikalie_Name===nameVal);
      const kint={Experiment_ID:id,Quelle_Typ:chem?'Chemikalie':'Sonstiges',Chemikalie_Name:chem?nameVal:null,Komponente_Name:chem?null:nameVal,Experiment_Ref:null,Hersteller:(document.getElementById('kh-'+i)?.value||'').trim(),Menge:document.getElementById('km-'+i)?.value!==''?parseFloat(document.getElementById('km-'+i)?.value):null,Einheit:(document.getElementById('ke-'+i)?.value||'').trim(),Rolle:(document.getElementById('kr-'+i)?.value||'').trim()};
      kompRows.push({int:kint,sp:mapTo(kint,FIELDS.komponenten),spId});
    });
    for(const r of kompRows){
      if(r.spId){await spPatch(LIST.komponenten,r.spId,r.sp);const ex=allKomps.find(k=>k._spId===r.spId);if(ex)Object.assign(ex,r.int);}
      else{const saved=await spPost(LIST.komponenten,r.sp);allKomps.push({...r.int,_spId:saved.d.Id});}
    }
    cachedErgebnisse=computeErgebnisse();filterExp();filterRes();
    alertEl.innerHTML='<div class="alert alert-ok">Gespeichert!</div>';setTimeout(closePanel,900);
  }catch(e){alertEl.innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
  finally{btn.disabled=false;btn.textContent='Speichern';}
}
async function deleteExp(){
  if(!editingExp||!confirm(`Experiment ${editingExp.Experiment_ID} wirklich löschen?`))return;
  try{await spDelete(LIST.experimente,editingExp._spId);allExp=allExp.filter(e=>e._spId!==editingExp._spId);cachedErgebnisse=computeErgebnisse();filterExp();filterRes();closePanel();}
  catch(e){document.getElementById('exp-alert').innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
}

// ─── Material form ─────────────────────────────────────────────────────────
let specIdx=0;
function initMatForm(item){
  document.getElementById('mat-alert').innerHTML='';
  document.getElementById('panel-mat-title').textContent=item?'Messung bearbeiten':'Neue Messung(en)';
  document.getElementById('btn-del-mat').style.display=item?'':'none';
  document.getElementById('btn-save-mat').textContent=item?'Speichern':'Alle speichern';
  document.getElementById('mat-multi-fields').style.display=item?'none':'';
  document.getElementById('mat-single-fields').style.display=item?'':'none';
  fillSelect('f-mat-lafo',lagerfolgen,'Lagerfolge_ID',l=>`${l.Lagerfolge_ID} – ${l.Name}`);
  const mSel=document.getElementById('f-mat-maschine');mSel.innerHTML='<option value="">– keine –</option>';
  allMaschinen.filter(m=>m.Typ==='Zugprüfmaschine').forEach(m=>{const o=document.createElement('option');o.value=m.Kuerzel;o.textContent=`${m.Kuerzel} – ${m.Name}`;mSel.appendChild(o);});
  const dl=document.getElementById('mat-expid-list');dl.innerHTML='';
  allExp.forEach(e=>{const o=document.createElement('option');o.value=e.Experiment_ID;dl.appendChild(o);});
  if(item){
    document.getElementById('f-mat-expid').value=item.Experiment_ID||'';
    document.getElementById('f-mat-lafo').value=item.Lagerfolge_ID||'';
    document.getElementById('f-mat-maschine').value=item.Maschine||'';
    document.getElementById('f-mat-l').value=item.Laenge_mm??'';
    document.getElementById('f-mat-b').value=item.Breite_mm??'';
    document.getElementById('f-mat-f').value=item.Kraft_N??'';
    document.getElementById('f-mat-h').value=item.Holzbruch_pct!=null?item.Holzbruch_pct*100:'';
    document.getElementById('f-mat-k').value=item.Kommentar||'';
  } else {document.getElementById('f-mat-expid').value='';document.getElementById('f-mat-maschine').value='Z2';document.getElementById('f-mat-lafo').value='LAFO-DRY-01';document.getElementById('spec-tbody').innerHTML='';specIdx=0;addSpecRow();addSpecRow();addSpecRow();}
}
function editMat(spId){const item=allMat.find(m=>m._spId===spId);if(!item)return;editingMat=item;document.getElementById('overlay').classList.add('open');document.getElementById('panel-mat').classList.add('open');activePanel='mat';initMatForm(item);}
function addSpecRow(){const i=specIdx++;const tr=document.createElement('tr');tr.id='spec-row-'+i;tr.innerHTML=`<td><input type="number" id="sl-${i}" value="10" min="0" step="0.1" oninput="recalcMpa(${i})"></td><td><input type="number" id="sb-${i}" value="20" min="0" step="0.1" oninput="recalcMpa(${i})"></td><td><input type="number" id="sf-${i}" min="0" step="1" oninput="recalcMpa(${i})"></td><td class="mpa-cell" id="sm-${i}">–</td><td><input type="number" id="sh-${i}" min="0" max="100" step="1" placeholder="0–100"></td><td><input type="text" id="sk-${i}"></td><td><button class="del-btn" onclick="removeSpecRow(${i})">×</button></td>`;document.getElementById('spec-tbody').appendChild(tr);}
function removeSpecRow(i){document.getElementById('spec-row-'+i)?.remove();}
function recalcMpa(i){const l=document.getElementById('sl-'+i)?.value,b=document.getElementById('sb-'+i)?.value,f=document.getElementById('sf-'+i)?.value,m=document.getElementById('sm-'+i);if(m){const v=calcMpa(l,b,f);m.textContent=v?v.replace('.',',')+' MPa':'–';}}
async function saveMaterial(){
  const alertEl=document.getElementById('mat-alert'),btn=document.getElementById('btn-save-mat');
  const expId=document.getElementById('f-mat-expid').value.trim().toUpperCase(),lafo=document.getElementById('f-mat-lafo').value;
  if(!expId||!lafo){alertEl.innerHTML='<div class="alert alert-err">Experiment-ID und Protokoll sind Pflichtfelder.</div>';return;}
  if(!allExp.find(e=>e.Experiment_ID===expId)){alertEl.innerHTML=`<div class="alert alert-err">Experiment-ID „${esc(expId)}" nicht gefunden.</div>`;return;}
  btn.disabled=true;alertEl.innerHTML='';
  try{
    if(editingMat){
      const fVal=document.getElementById('f-mat-f').value,hVal=document.getElementById('f-mat-h').value;
      if(!fVal||hVal===''){alertEl.innerHTML='<div class="alert alert-err">Kraft und Holzbruch sind Pflichtfelder.</div>';btn.disabled=false;return;}
      const int={Experiment_ID:expId,Lagerfolge_ID:lafo,Maschine:document.getElementById('f-mat-maschine').value||null,Laenge_mm:parseFloat(document.getElementById('f-mat-l').value)||null,Breite_mm:parseFloat(document.getElementById('f-mat-b').value)||null,Kraft_N:parseFloat(fVal),Holzbruch_pct:parseFloat(hVal)/100,Kommentar:document.getElementById('f-mat-k').value||''};
      await spPatch(LIST.material,editingMat._spId,mapTo(int,FIELDS.material));Object.assign(editingMat,int);
    } else {
      const rows=[];
      document.querySelectorAll('#spec-tbody tr').forEach(tr=>{const i=tr.id.replace('spec-row-','');const l=document.getElementById('sl-'+i)?.value,b=document.getElementById('sb-'+i)?.value,f=document.getElementById('sf-'+i)?.value;const h=document.getElementById('sh-'+i)?.value,k=document.getElementById('sk-'+i)?.value;if(l&&b&&f&&h!==''){const int={Experiment_ID:expId,Lagerfolge_ID:lafo,Laenge_mm:parseFloat(l),Breite_mm:parseFloat(b),Kraft_N:parseFloat(f),Holzbruch_pct:parseFloat(h)/100,Kommentar:k||''};rows.push({int,sp:mapTo(int,FIELDS.material)});}});
      if(!rows.length){alertEl.innerHTML='<div class="alert alert-err">Keine gültigen Probekörper.</div>';btn.disabled=false;return;}
      btn.textContent=`Speichert ${rows.length}…`;
      const saved=await Promise.all(rows.map(r=>spPost(LIST.material,r.sp)));
      saved.forEach((s,i)=>allMat.unshift({...rows[i].int,_spId:s.d.Id}));
    }
    computeMpaRanges();cachedErgebnisse=computeErgebnisse();filterMat();filterRes();
    alertEl.innerHTML='<div class="alert alert-ok">Gespeichert!</div>';setTimeout(closePanel,900);
  }catch(e){alertEl.innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
  finally{btn.disabled=false;btn.textContent=editingMat?'Speichern':'Alle speichern';}
}
async function deleteMat(){
  if(!editingMat||!confirm('Messung wirklich löschen?'))return;
  try{await spDelete(LIST.material,editingMat._spId);allMat=allMat.filter(m=>m._spId!==editingMat._spId);computeMpaRanges();cachedErgebnisse=computeErgebnisse();filterMat();filterRes();closePanel();}
  catch(e){document.getElementById('mat-alert').innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
}

// ─── SC form ───────────────────────────────────────────────────────────────
let scIdx=0;
function initSCForm(item){
  document.getElementById('sc-alert').innerHTML='';
  document.getElementById('panel-sc-title').textContent=item?'SC-Messung bearbeiten':'Feststoffgehalt erfassen';
  document.getElementById('btn-del-sc').style.display=item?'':'none';
  document.getElementById('btn-save-sc').textContent=item?'Speichern':'Alle speichern';
  document.getElementById('sc-multi-fields').style.display=item?'none':'';
  document.getElementById('sc-single-fields').style.display=item?'':'none';
  if(item){
    document.getElementById('f-sc-expid').value=item.Experiment_ID||'';
    document.getElementById('f-sc-probe').value=item.Probe||'';
    document.getElementById('f-sc-leer').value=item.Leergewicht_g??'';
    document.getElementById('f-sc-ein').value=item.Einwaage_g??'';
    document.getElementById('f-sc-end').value=item.Endgewicht_g??'';
    document.getElementById('f-sc-komm').value=item.Kommentar||'';
  } else {
    document.getElementById('f-sc-expid').value='';
    document.getElementById('sc-tbody').innerHTML='';
    scIdx=0;addSCRow();
  }
}
function editSC(spId){const item=allSC.find(s=>s._spId===spId);if(!item)return;editingSC=item;document.getElementById('overlay').classList.add('open');document.getElementById('panel-sc').classList.add('open');activePanel='sc';initSCForm(item);}
function openSCPanel(expId){document.getElementById('overlay').classList.add('open');document.getElementById('panel-sc').classList.add('open');activePanel='sc';editingSC=null;initSCForm(null);if(expId)document.getElementById('f-sc-expid').value=expId;}
function addSCRow(){const i=scIdx++;const tr=document.createElement('tr');tr.id='sc-row-'+i;tr.innerHTML=`<td><input type="text" id="sp-${i}" style="width:80px"></td><td><input type="number" id="sl-sc-${i}" min="0" step="0.0001" oninput="recalcSC(${i})" style="width:90px"></td><td><input type="number" id="se-${i}" min="0" step="0.0001" oninput="recalcSC(${i})" style="width:90px"></td><td><input type="number" id="sg-${i}" min="0" step="0.0001" oninput="recalcSC(${i})" style="width:90px"></td><td class="mpa-cell" id="ss-${i}" style="white-space:nowrap">–</td><td><input type="text" id="sk-sc-${i}" style="width:100px"></td><td><button class="del-btn" onclick="removeSCRow(${i})">×</button></td>`;document.getElementById('sc-tbody').appendChild(tr);}
function removeSCRow(i){document.getElementById('sc-row-'+i)?.remove();}
function recalcSC(i){const leer=document.getElementById('sl-sc-'+i)?.value,ein=document.getElementById('se-'+i)?.value,end=document.getElementById('sg-'+i)?.value,sc=document.getElementById('ss-'+i);if(sc){const v=calcSC(leer,ein,end);sc.textContent=v!=null?fmtDec(v,2)+'%':'–';}}
async function saveSC(){
  const alertEl=document.getElementById('sc-alert'),btn=document.getElementById('btn-save-sc');
  const expId=document.getElementById('f-sc-expid').value.trim().toUpperCase();
  if(!expId){alertEl.innerHTML='<div class="alert alert-err">Experiment-ID ist Pflichtfeld.</div>';return;}
  btn.disabled=true;alertEl.innerHTML='';
  try{
    if(editingSC){
      const leer=document.getElementById('f-sc-leer').value,ein=document.getElementById('f-sc-ein').value,end=document.getElementById('f-sc-end').value;
      if(!leer||!ein){alertEl.innerHTML='<div class="alert alert-err">Leergewicht und Einwaage sind Pflichtfelder.</div>';btn.disabled=false;return;}
      const int={Experiment_ID:expId,Probe:document.getElementById('f-sc-probe').value.trim(),Leergewicht_g:parseFloat(leer),Einwaage_g:parseFloat(ein),Endgewicht_g:end!==''?parseFloat(end):null,Kommentar:document.getElementById('f-sc-komm').value.trim()};
      await spPatch(LIST.feststoffgehalt,editingSC._spId,mapTo(int,FIELDS.feststoffgehalt));Object.assign(editingSC,int);
    } else {
      const rows=[];
      document.querySelectorAll('#sc-tbody tr').forEach(tr=>{const i=tr.id.replace('sc-row-','');const leer=document.getElementById('sl-sc-'+i)?.value,ein=document.getElementById('se-'+i)?.value,end=document.getElementById('sg-'+i)?.value,probe=document.getElementById('sp-'+i)?.value,k=document.getElementById('sk-sc-'+i)?.value;if(leer&&ein){const int={Experiment_ID:expId,Probe:probe||'',Leergewicht_g:parseFloat(leer),Einwaage_g:parseFloat(ein),Endgewicht_g:end!==''?parseFloat(end):null,Kommentar:k||''};rows.push({int,sp:mapTo(int,FIELDS.feststoffgehalt)});}});
      if(!rows.length){alertEl.innerHTML='<div class="alert alert-err">Keine gültigen Proben (Leergewicht + Einwaage erforderlich).</div>';btn.disabled=false;return;}
      btn.textContent=`Speichert ${rows.length}…`;
      const saved=await Promise.all(rows.map(r=>spPost(LIST.feststoffgehalt,r.sp)));
      saved.forEach((s,i)=>allSC.unshift({...rows[i].int,_spId:s.d.Id}));
    }
    cachedErgebnisse=computeErgebnisse();filterRes();
    alertEl.innerHTML='<div class="alert alert-ok">Gespeichert!</div>';setTimeout(closePanel,900);
  }catch(e){alertEl.innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
  finally{btn.disabled=false;btn.textContent=editingSC?'Speichern':'Alle speichern';}
}
async function deleteSC(){
  if(!editingSC||!confirm('SC-Messung wirklich löschen?'))return;
  try{await spDelete(LIST.feststoffgehalt,editingSC._spId);allSC=allSC.filter(s=>s._spId!==editingSC._spId);cachedErgebnisse=computeErgebnisse();filterRes();closePanel();}
  catch(e){document.getElementById('sc-alert').innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
}

// ─── Chemikalien form ──────────────────────────────────────────────────────
function initChemForm(item){
  document.getElementById('chem-alert').innerHTML='';
  document.getElementById('panel-chem-title').textContent=item?'Chemikalie bearbeiten':'Neue Chemikalie';
  document.getElementById('btn-del-chem').style.display=item?'':'none';
  document.getElementById('pubchem-status').textContent='';document.getElementById('pubchem-status').className='pubchem-status';
  const dl=document.getElementById('hersteller-list');dl.innerHTML='';
  [...new Set(allChem.map(c=>c.Hersteller).filter(Boolean))].sort().forEach(h=>{const o=document.createElement('option');o.value=h;dl.appendChild(o);});
  const ortDl=document.getElementById('ort-list');ortDl.innerHTML='';
  [...new Set(allChem.map(c=>c.Ort).filter(Boolean))].sort().forEach(v=>{const o=document.createElement('option');o.value=v;ortDl.appendChild(o);});
  if(item){
    document.getElementById('f-chem-name').value=item.Chemikalie_Name||'';
    document.getElementById('f-chem-formel').value=item.Formel||'';
    document.getElementById('f-chem-iupac').value=item.IUPAC||'';
    document.getElementById('f-chem-hersteller').value=item.Hersteller||'';
    document.getElementById('f-chem-menge').value=item.Menge??'';
    document.getElementById('f-chem-einheit').value=item.Einheit||'';
    document.getElementById('f-chem-ort').value=item.Ort||'';
    document.getElementById('f-chem-herkunft').value=item.Herkunft||'';
    document.getElementById('f-chem-fuellstand').value=item.Fuellstand||'';
    document.getElementById('f-chem-komm').value=item.Kommentar||'';
  } else {
    ['f-chem-name','f-chem-formel','f-chem-iupac','f-chem-hersteller','f-chem-einheit','f-chem-ort','f-chem-komm'].forEach(id=>document.getElementById(id).value='');
    document.getElementById('f-chem-menge').value='';
    document.getElementById('f-chem-herkunft').value='Gekauft';
    document.getElementById('f-chem-fuellstand').value='';
  }
}
function editChem(spId){const item=allChem.find(c=>c._spId===spId);if(!item)return;editingChem=item;document.getElementById('overlay').classList.add('open');document.getElementById('panel-chem').classList.add('open');activePanel='chem';initChemForm(item);}

let pubchemTimer=null;
function onChemNameInput(val){const status=document.getElementById('pubchem-status');clearTimeout(pubchemTimer);if(val.trim().length<3){status.textContent='';return;}status.textContent='Suche in PubChem…';status.className='pubchem-status';pubchemTimer=setTimeout(()=>fetchPubchem(val.trim()),600);}
async function fetchPubchem(name){const status=document.getElementById('pubchem-status');try{const r=await fetch(`https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/${encodeURIComponent(name)}/property/MolecularFormula,IUPACName/JSON`);if(!r.ok){status.textContent='Nicht in PubChem gefunden.';status.className='pubchem-status err';return;}const d=await r.json();const p=d?.PropertyTable?.Properties?.[0];if(!p){status.textContent='Nicht in PubChem gefunden.';status.className='pubchem-status err';return;}const fe=document.getElementById('f-chem-formel'),ie=document.getElementById('f-chem-iupac');if(!fe.value&&p.MolecularFormula)fe.value=p.MolecularFormula;if(!ie.value&&p.IUPACName)ie.value=p.IUPACName;status.textContent='✓ PubChem: Formel und IUPAC vorgeschlagen.';status.className='pubchem-status found';}catch(e){status.textContent='PubChem nicht erreichbar.';status.className='pubchem-status err';}}

async function saveChem(){
  const alertEl=document.getElementById('chem-alert'),btn=document.getElementById('btn-save-chem');
  const name=document.getElementById('f-chem-name').value.trim();
  if(!name){alertEl.innerHTML='<div class="alert alert-err">Name ist ein Pflichtfeld.</div>';return;}
  btn.disabled=true;btn.textContent='Speichert…';alertEl.innerHTML='';
  const int={Chemikalie_Name:name,IUPAC:document.getElementById('f-chem-iupac').value.trim(),Formel:document.getElementById('f-chem-formel').value.trim(),Hersteller:document.getElementById('f-chem-hersteller').value.trim(),Menge:document.getElementById('f-chem-menge').value!==''?parseFloat(document.getElementById('f-chem-menge').value):null,Einheit:document.getElementById('f-chem-einheit').value.trim(),Ort:document.getElementById('f-chem-ort').value.trim(),Herkunft:document.getElementById('f-chem-herkunft').value.trim(),Fuellstand:document.getElementById('f-chem-fuellstand').value||null,Kommentar:document.getElementById('f-chem-komm').value.trim()};
  try{
    const sp=mapTo(int,FIELDS.chemikalien);
    if(editingChem){await spPatch(LIST.chemikalien,editingChem._spId,sp);Object.assign(editingChem,int);}
    else{const saved=await spPost(LIST.chemikalien,sp);allChem.push({...int,_spId:saved.d.Id});allChem.sort((a,b)=>(a.Chemikalie_Name||'').localeCompare(b.Chemikalie_Name||''));}
    filterChem();alertEl.innerHTML='<div class="alert alert-ok">Gespeichert!</div>';setTimeout(closePanel,900);
  }catch(e){alertEl.innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
  finally{btn.disabled=false;btn.textContent='Speichern';}
}
async function deleteChem(){
  if(!editingChem||!confirm(`„${editingChem.Chemikalie_Name}" wirklich löschen?`))return;
  try{await spDelete(LIST.chemikalien,editingChem._spId);allChem=allChem.filter(c=>c._spId!==editingChem._spId);filterChem();closePanel();}
  catch(e){document.getElementById('chem-alert').innerHTML=`<div class="alert alert-err">Fehler: ${esc(e.message)}</div>`;}
}
