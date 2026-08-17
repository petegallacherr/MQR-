'use strict';

const state={herd:null,mastitis:null,dct:null,mastaplex:null,result:null};
const ids=['farmName','vetName','seasonYear','herdTests','peakCows','firstCalvers','sccTs','sccLa','dairyCompany','supplyNumber','ptptCode','heifersSealed','bmsccPrevious','bmsccCurrent','calvingStart','expectedDryOff','prescriptionStatus','herdFile','mastitisFile','dctFile','mastaplexFile','herdStatus','mastitisStatus','dctStatus','mastaplexStatus','mappingPanel','generateBtn','viewReportBtn','generateStatus','demoBtn','errorBox','report','reportTitle','reportMeta','consultFacts','dataQuality','kpis','executiveSummary','monthlyTable','caseTiming','caseSummary','drugSummary','sccSummary','sccTransitions','quarterSummary','mastaplexSummary','previousDctSummary','dctSummary','dctOrder','cowTable','cowCountText','cowSearch','csvBtn','printBtn'];
const els=Object.fromEntries(ids.map(id=>[id,document.getElementById(id)]));
els.seasonYear.value=new Date().getFullYear()-1;
els.sccTs.value=150;
els.sccLa.value=250;

const clean=s=>String(s??'').trim();
const norm=s=>clean(s).toLowerCase().replace(/[\s_\-\/\\().]+/g,'').replace(/[^a-z0-9]/g,'');
const safe=s=>String(s??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
const num=v=>{if(typeof v==='number'&&Number.isFinite(v))return v;const x=parseFloat(String(v??'').replace(/,/g,''));return Number.isFinite(x)?x:null};
const fmt=n=>n==null?'—':Number(n).toLocaleString();
const pct=(a,b,d=1)=>b!=null&&b!==0&&a!=null?`${(a/b*100).toFixed(d)}%`:'N/A';
const median=arr=>{const a=arr.filter(Boolean).map(d=>d.getTime()).sort((x,y)=>x-y);if(!a.length)return null;const m=Math.floor(a.length/2);return new Date(a.length%2?a[m]:(a[m-1]+a[m])/2)};

function parseDate(v){
  if(!v&&v!==0)return null;
  if(v instanceof Date&&!isNaN(v))return v;
  if(typeof v==='number'&&v>20000&&v<80000){const d=new Date(Date.UTC(1899,11,30));d.setUTCDate(d.getUTCDate()+v);return new Date(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())}
  const s=clean(v);if(!s)return null;
  const nz=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);if(nz){let y=+nz[3];if(y<100)y+=2000;const d=new Date(y,+nz[2]-1,+nz[1]);return isNaN(d)?null:d}
  const d=new Date(s);return isNaN(d)?null:d;
}
function isoDate(d){return d?d.toLocaleDateString('en-NZ',{day:'2-digit',month:'short',year:'numeric'}):''}
function dateKey(d){return d?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`:''}
function daysBetween(a,b){if(!a||!b)return null;return Math.round((a.getTime()-b.getTime())/86400000)}

function parseCsvMatrix(text){
  const rows=[];let row=[],cell='',quoted=false;
  for(let i=0;i<text.length;i++){const ch=text[i],next=text[i+1];if(ch==='"'&&quoted&&next==='"'){cell+='"';i++}else if(ch==='"')quoted=!quoted;else if(ch===','&&!quoted){row.push(cell);cell=''}else if((ch==='\n'||ch==='\r')&&!quoted){if(ch==='\r'&&next==='\n')i++;row.push(cell);cell='';if(row.some(x=>clean(x)))rows.push(row);row=[]}else cell+=ch}
  if(cell||row.length){row.push(cell);if(row.some(x=>clean(x)))rows.push(row)}return rows;
}
function headerScore(row,kind){
  const s=row.map(norm).join('|');let score=0;const has=(...x)=>x.some(k=>s.includes(norm(k)));
  if(has('tag','animal','cow','id'))score+=3;
  if(kind==='herd'){if(has('scc','herdtest','somatic'))score+=3;if(has('calving','pregnancy','bcs'))score+=2}
  if(kind==='mastitis'||kind==='dct'){if(has('treatment','product','remedy','medicine'))score+=3;if(has('date','condition','quarter','teat'))score+=2}
  if(kind==='mastaplex'){if(has('species','organism','growth','result','bacteria'))score+=4;if(has('quarter','sample','date'))score+=1}
  score+=Math.min(2,row.filter(x=>clean(x)).length/5);return score;
}
function matrixToObjects(matrix,kind){
  if(!matrix.length)return[];let best=0,bestScore=-1;
  for(let i=0;i<Math.min(10,matrix.length);i++){const sc=headerScore(matrix[i],kind);if(sc>bestScore){bestScore=sc;best=i}}
  const headers=matrix[best].map((h,i)=>clean(h)||`Column ${i+1}`);const seen={};
  headers.forEach((h,i)=>{if(seen[h])headers[i]=`${h} ${i+1}`;seen[headers[i]]=1});
  return matrix.slice(best+1).filter(r=>r.some(x=>clean(x))).map(r=>Object.fromEntries(headers.map((h,i)=>[h,r[i]??''])));
}
let xlsxLoadPromise=null;
function ensureXlsx(){
  if(typeof XLSX!=='undefined')return Promise.resolve();
  if(xlsxLoadPromise)return xlsxLoadPromise;
  xlsxLoadPromise=new Promise((resolve,reject)=>{
    const script=document.createElement('script');
    script.src='https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
    script.async=true;
    const timer=setTimeout(()=>reject(new Error('Excel reader could not be loaded. For reliable offline use, export the MINDA file as CSV.')),8000);
    script.onload=()=>{clearTimeout(timer);typeof XLSX!=='undefined'?resolve():reject(new Error('Excel reader loaded but was unavailable. Please use CSV.'))};
    script.onerror=()=>{clearTimeout(timer);reject(new Error('Excel reader could not be loaded. Please use the original MINDA CSV export.'))};
    document.head.appendChild(script);
  }).catch(e=>{xlsxLoadPromise=null;throw e});
  return xlsxLoadPromise;
}
async function readTabular(file,kind){
  const ext=file.name.split('.').pop().toLowerCase();let matrix;
  if(ext==='csv')matrix=parseCsvMatrix(await file.text());
  else{await ensureXlsx();const data=await file.arrayBuffer();const wb=XLSX.read(data,{type:'array',cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];matrix=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true})}
  return matrixToObjects(matrix,kind);
}
function headersOf(rows){return rows?.length?Object.keys(rows[0]):[]}
function findHeader(headers,aliases,contains=[]){const byNorm=new Map(headers.map(h=>[norm(h),h]));for(const a of aliases)if(byNorm.has(norm(a)))return byNorm.get(norm(a));for(const h of headers)if(contains.some(x=>norm(h).includes(norm(x))))return h;return null}

function parseSccHeaderDate(header){
  const s=clean(header);const m=s.match(/(?:Herd\s*Test\s*Results[-\s]*)?([A-Za-z]{3,9})\s+(20\d{2})\s*SCC/i);if(!m)return null;
  const months={jan:0,january:0,feb:1,february:1,mar:2,march:2,apr:3,april:3,may:4,jun:5,june:5,jul:6,july:6,aug:7,august:7,sep:8,sept:8,september:8,oct:9,october:9,nov:10,november:10,dec:11,december:11};
  const month=months[m[1].toLowerCase()];if(month==null)return null;return new Date(+m[2],month,1);
}
function sccDateLabel(d){return d?d.toLocaleDateString('en-NZ',{month:'short',year:'numeric'}):''}
function detectHerd(rows){
  const h=headersOf(rows);const map={
    tag:findHeader(h,['Tag','Animal ID','Animal-Animal ID','Cow Number','Cow No','Animal','PTPT','ID'],['animalanimalid','tag','cow','animalid']),
    yearBorn:findHeader(h,['Animal-Year Born','Year Born','Birth Year'],['yearborn']),
    calving:findHeader(h,['Animal-Calving Date','Calving date','Calving Date (this season)','Calved'],['calvingdate']),
    bcs:findHeader(h,['Animal-BCS','BCS','Body Condition Score'],['bcs','bodycondition']),
    preg:findHeader(h,['Mating-Pregnancy Diagnosis','Pregnancy Diagnosis','Preg Diagnosis','Pregnancy Status'],['pregnancy','pregdiag']),
    expectedCalving:findHeader(h,['Pre-Calving-Expected Calving Date','Expected Calving date','Expected Calving Date','Due Date'],['expectedcalving','duedate'])
  };
  const dated=h.map(header=>({header,date:parseSccHeaderDate(header)})).filter(x=>x.date).sort((a,b)=>b.date-a.date);
  if(dated.length){
    const seasonYear=+els.seasonYear.value||new Date().getFullYear()-1,start=new Date(seasonYear,5,1),end=new Date(seasonYear+1,5,1);
    const current=dated.filter(x=>x.date>=start&&x.date<end);const previous=dated.filter(x=>x.date<start)[0]||null;
    const selected=[...current,...(previous?[previous]:[])];map.scc=selected.map(x=>x.header);map.sccDates=selected.map(x=>sccDateLabel(x.date));map.sccAll=dated.map(x=>x.header);map.sccAllDates=dated.map(x=>sccDateLabel(x.date));map.currentHerdTests=current.length;map.previousHerdTest=previous?sccDateLabel(previous.date):'';
  }else{
    const explicit=['Most recent H/T','2nd','3rd','4th','5th'].map(x=>findHeader(h,[x])).filter(Boolean);const sccLike=h.filter(x=>/scc|herd.?test|h\/t/i.test(x)&&!Object.values(map).includes(x));map.scc=[...new Set([...explicit,...sccLike])].slice(0,5);map.currentHerdTests=Math.min(4,Math.max(1,map.scc.length-(map.scc.length>1?1:0)));
  }
  if(!map.tag&&h.length)map.tag=h[0];return map;
}
function detectTreatment(rows){
  const h=headersOf(rows);const map={
    tag:findHeader(h,['Animal No','Tag','Animal ID','Cow Number','Cow No','Animal','ID'],['animalno','tag','cow','animalid']),
    historic:findHeader(h,['IsHistoricAnimal','Historic Animal'],['historic']),
    category:findHeader(h,['Condition','Category','Case Type','Treatment Type','Diagnosis','Reason'],['condition','category','diagnosis']),
    date:findHeader(h,['Created Date','Treatment Date','Date','Event Date','Start Date'],['createddate','treatmentdate','eventdate','date']),
    quarter:findHeader(h,['Areas Affected','Quarter','Affected Quarter','Teat'],['areasaffected','quarter','teat']),
    drug:findHeader(h,['Treatment Used','Product','Drug','Treatment','Remedy','Medicine'],['treatmentused','product','drug','remedy','medicine','treatment']),
    duration:findHeader(h,['No of Doses','No. of Doses','Number of Doses','Days','Duration','Treatments','Number of treatments','No. Treatments'],['noofdoses','numberofdoses','duration','days','treatments']),
    interval:findHeader(h,['Interval'],['interval']),amount:findHeader(h,['Amount'],['amount']),lastTreatment:findHeader(h,['Last Treatment Date'],['lasttreatmentdate']),milkWithholding:findHeader(h,['Milk Withholding Period'],['milkwithholding']),meatWithholding:findHeader(h,['Meat Withholding Period'],['meatwithholding']),returnToVat:findHeader(h,['Return to Vat'],['returntovat']),completed:findHeader(h,['Completed Date'],['completeddate'])
  };
  if(!map.tag&&h.length)map.tag=h[0];return map;
}
function detectMastaplex(rows){const h=headersOf(rows);const map={tag:findHeader(h,['Tag','Animal ID','Cow Number','Cow No','Animal','ID'],['tag','cow','animalid']),species:findHeader(h,['Species','Organism','Growth','Result','Bacteria'],['species','organism','growth','bacteria']),quarter:findHeader(h,['Quarter','Affected Quarter'],['quarter']),date:findHeader(h,['Date','Sample Date','Test Date'],['sampledate','testdate'])};if(!map.tag&&h.length>=3)map.tag=h[2];if(!map.species&&h.length>=5)map.species=h[4];return map}
function confidence(kind,map){const req=kind==='herd'?['tag','scc']:kind==='mastaplex'?['tag','species']:['tag','date','drug'];return req.filter(k=>Array.isArray(map[k])?map[k].length:map[k]).length/req.length}
function setStatus(kind,rows,map,name){
  const el=els[`${kind}Status`],c=confidence(kind,map);el.className=`file-status ${c>=.99?'good':c>=.5?'warn':'bad'}`;
  if(kind==='herd'&&map.sccAllDates?.length){const selected=(map.sccDates||[]).join(', ');el.textContent=`${rows.length.toLocaleString()} rows · ${map.sccAllDates.length} dated herd tests detected · using ${selected} · ${name}`}
  else el.textContent=`${rows.length.toLocaleString()} rows · ${c>=.99?'columns recognised':c>=.5?'check detected columns':'column match poor'} · ${name}`;
  renderMappings();
}
function renderMappings(){const labels={herd:'Herd',mastitis:'Mastitis',dct:'DCT / ITS',mastaplex:'Mastaplex'};const items=[];for(const kind of Object.keys(labels)){const d=state[kind];if(!d)continue;const parts=Object.entries(d.map).filter(([k])=>!['sccAll'].includes(k)).map(([k,v])=>`${k}: ${Array.isArray(v)?v.join(', '):(v||'not found')}`);items.push(`<div class="mapping-row"><strong>${labels[kind]}:</strong><span>${safe(parts.join(' · '))}</span></div>`)}els.mappingPanel.innerHTML=items.join('');els.mappingPanel.classList.toggle('hidden',!items.length)}
async function handleFile(kind,file){if(!file)return;try{const rows=await readTabular(file,kind);if(!rows.length)throw new Error('No data rows found.');const map=kind==='herd'?detectHerd(rows):kind==='mastaplex'?detectMastaplex(rows):detectTreatment(rows);state[kind]={rows,map,name:file.name};if(kind==='herd'&&map.currentHerdTests>=1&&map.currentHerdTests<=10)els.herdTests.value=String(map.currentHerdTests);setStatus(kind,rows,map,file.name)}catch(e){const el=els[`${kind}Status`];el.className='file-status bad';el.textContent=e.message}}
els.herdFile.addEventListener('change',e=>handleFile('herd',e.target.files[0]));els.mastitisFile.addEventListener('change',e=>handleFile('mastitis',e.target.files[0]));els.dctFile.addEventListener('change',e=>handleFile('dct',e.target.files[0]));els.mastaplexFile.addEventListener('change',e=>handleFile('mastaplex',e.target.files[0]));

els.seasonYear.addEventListener('change',()=>{if(!state.herd)return;state.herd.map=detectHerd(state.herd.rows);if(state.herd.map.currentHerdTests>=1&&state.herd.map.currentHerdTests<=10)els.herdTests.value=String(state.herd.map.currentHerdTests);setStatus('herd',state.herd.rows,state.herd.map,state.herd.name)});

const prescriptionFieldIds=['dairyCompany','supplyNumber','ptptCode','heifersSealed','bmsccPrevious','bmsccCurrent','calvingStart','expectedDryOff'];
function updatePrescriptionFields(){
  let complete=0;
  for(const id of prescriptionFieldIds){
    const input=els[id],field=input?.closest('[data-prescription-field]');
    if(!input||!field)continue;
    const filled=clean(input.value)!=='';
    field.classList.toggle('prescription-complete',filled);
    field.classList.toggle('prescription-missing',!filled);
    if(filled)complete++;
  }
  if(els.prescriptionStatus){
    const missing=prescriptionFieldIds.length-complete;
    els.prescriptionStatus.innerHTML=missing
      ? `<strong>${missing} prescription detail${missing===1?'':'s'} still needed.</strong> You can still generate the consultation report.`
      : '<strong>Prescription details complete.</strong>';
    els.prescriptionStatus.classList.toggle('complete',missing===0);
  }
}
for(const id of prescriptionFieldIds){els[id]?.addEventListener('input',updatePrescriptionFields);els[id]?.addEventListener('change',updatePrescriptionFields)}
updatePrescriptionFields();

const DCT_DRUGS=['quadrant dc','duramast dc 500','duramast dc 600','cepravin dry cow','cefa-safe','dryclox dc','ceprotect dc','juraclox','orbenin dry','noroclox dc 600'];
const SEALANTS=['dryzen','duraseal','dryseal','teatseal','sureseal'];
const NSAIDS=['metacam 20mg/ml for injection','metacam 20mg/ml','key injection','norflunix','metacam'];
function isNSAID(drug){const d=clean(drug).toLowerCase();return NSAIDS.some(x=>d.includes(x))}
function isDctDrug(drug){const d=clean(drug).toLowerCase();return DCT_DRUGS.some(x=>d.includes(x))}
function isSealant(drug){const d=clean(drug).toLowerCase();return SEALANTS.some(x=>d.includes(x))}
function classifyMastitis(cat){const c=clean(cat).toLowerCase();if(c.includes('subclinical'))return'Mastitis - Subclinical';if(c.includes('clinical')||c.includes('mastitis'))return'Mastitis - Clinical';return''}
function parseQuarters(v){const raw=clean(v).toUpperCase();if(!raw)return[];return [...new Set(raw.split(/[^A-Z]+/).filter(q=>['LF','RF','LR','RR'].includes(q)))]}
function treatmentRecords(data,mode){
  const out=[],m=data?.map||{};
  (data?.rows||[]).forEach((r,i)=>{const id=clean(r[m.tag]);if(!id)return;const drug=clean(r[m.drug]),cat=clean(r[m.category]),quarter=clean(r[m.quarter]).toUpperCase();out.push({id,date:parseDate(r[m.date]),quarter,quarters:parseQuarters(quarter),drug,duration:num(r[m.duration]),historic:m.historic?clean(r[m.historic]):'',category:mode==='mastitis'?classifyMastitis(cat):isDctDrug(drug)?'DCT':isSealant(drug)?'Sealant':'',raw:r,index:i})});
  return out;
}
function indexByCow(records){const m=new Map();for(const r of records){if(!m.has(r.id))m.set(r.id,[]);m.get(r.id).push(r)}return m}
function indexMastaplex(data){const byCow=new Map(),m=data?.map||{};for(const r of data?.rows||[]){const id=clean(r[m.tag]),species=clean(r[m.species]);if(!id||!species)continue;if(!byCow.has(id))byCow.set(id,[]);byCow.get(id).push(species)}return byCow}

function dctAdvice(scc,preg,treatments,herdTests,ts,la){
  if(/empty/i.test(preg))return'Empty';
  if(scc[0]==null||scc[0]===0)return'Absent from H/T (Culled/Carryover)';
  if(treatments.some(x=>x.category==='Mastitis - Clinical'))return'Mastitis - Clinical';
  if(treatments.some(x=>x.category==='Mastitis - Subclinical'))return'Mastitis - Subclinical';
  // Assess every available current-season herd test. One high result anywhere
  // in the season is enough to move the cow into the appropriate DCT band.
  const relevant=scc.slice(0,Math.max(1,herdTests)).filter(v=>v!=null&&v!==0);
  const highest=relevant.length?Math.max(...relevant):null;
  if(highest==null)return'Absent from H/T (Culled/Carryover)';
  if(highest<ts)return'Teat Sealant only';
  if(highest>=la)return'LA DCT/Sealant Combo';
  return'SA DCT/Sealant Combo';
}
function dryPeriodStatus(scc,herdTests,threshold){
  if(herdTests<1)return{status:'Not available',pre:null,post:null,note:'No current herd test detected'};
  const firstCurrentIndex=Math.max(0,herdTests-1),preIndex=herdTests,post=scc[firstCurrentIndex],pre=scc[preIndex];
  if(pre==null||pre===0)return{status:'Not available',pre,post,note:'Pre-dry SCC unavailable'};
  if(post==null||post===0)return{status:'Missed H/T',pre,post,note:'First post-calving herd test unavailable'};
  if(pre<threshold&&post>=threshold)return{status:'New Infection',pre,post,note:''};
  if(pre>=threshold&&post<threshold)return{status:'Cured',pre,post,note:''};
  if(pre>=threshold&&post>=threshold)return{status:'Retained Infection',pre,post,note:''};
  return{status:'Low SCC',pre,post,note:''};
}
function dryOffTiming(cow,seasonYear){
  if(/empty/i.test(cow.preg))return{advice:'Cull/Carryover',targetDate:null};
  if(!cow.expectedCalving)return{advice:'N/A — expected calving date missing',targetDate:null};
  if(cow.bcs==null)return{advice:'N/A — BCS missing',targetDate:null};
  const target=cow.yearBorn===seasonYear-2?5.5:5;const gap=target-cow.bcs-.3;const days=gap/.5*30+50;const d=new Date(cow.expectedCalving);d.setDate(d.getDate()-Math.round(days));const adjustedMonth=d.getDate()>15?d.getMonth()+1:d.getMonth();let advice='';
  if(adjustedMonth===3)advice='1st April Dry Off';else if(adjustedMonth===4)advice='1st May Dry Off';else if(adjustedMonth>=5)advice='Dry Off End of Season';else advice='Dry Off NOW';
  return{advice,targetDate:d,targetBCS:target,days};
}

function buildCaseAnalysis(records,herdById,seasonYear,peak,firstCalvers){
  const mastitisRows=records.filter(r=>r.category);const antibioticRows=mastitisRows.filter(r=>!isNSAID(r.drug));const nsaidRows=mastitisRows.filter(r=>isNSAID(r.drug));
  const groups=new Map();
  for(const r of mastitisRows){const key=`${r.id}|${dateKey(r.date)||'nodate-'+r.index}`;if(!groups.has(key))groups.set(key,{id:r.id,date:r.date,rows:[],quarters:new Set(),antibiotics:new Set(),nsaid:false,maxDoses:0});const g=groups.get(key);g.rows.push(r);for(const qt of r.quarters)g.quarters.add(qt);if(isNSAID(r.drug))g.nsaid=true;else{if(r.drug)g.antibiotics.add(r.drug);g.maxDoses=Math.max(g.maxDoses,r.duration||0)}}
  const cases=[...groups.values()].filter(g=>g.antibiotics.size>0).sort((a,b)=>(a.date||0)-(b.date||0));
  const byCow=new Map();for(const c of cases){if(!byCow.has(c.id))byCow.set(c.id,[]);byCow.get(c.id).push(c)}
  const repeatCowIds=[...byCow.entries()].filter(([,v])=>v.length>1).map(([id])=>id);
  let repeatQuarterCowIds=[];for(const [id,arr] of byCow){if(arr.length<2)continue;let repeat=false;const seen=new Set();for(const c of arr){for(const q of c.quarters){if(seen.has(q))repeat=true;seen.add(q)}}if(repeat)repeatQuarterCowIds.push(id)}
  const combo=cases.filter(c=>c.nsaid).length,extended=cases.filter(c=>c.maxDoses>3).length,txChanges=cases.filter(c=>c.antibiotics.size>1).length;
  const productCounts={};for(const r of antibioticRows){if(r.drug)productCounts[r.drug]=(productCounts[r.drug]||0)+1}
  const quarters={LF:0,RF:0,LR:0,RR:0,Multi:0,'No quarter':0};for(const c of cases){const qs=[...c.quarters];if(qs.length>1)quarters.Multi++;else if(qs.length===1)quarters[qs[0]]++;else quarters['No quarter']++}
  const timing={d0_7:0,d8_30:0,over30:0,pre:0,unavailable:0,matched:0};
  for(const c of cases){const cow=herdById.get(c.id);if(!cow||!cow.calving||!c.date){timing.unavailable++;continue}const days=daysBetween(c.date,cow.calving);if(days==null){timing.unavailable++;continue}if(days < -30){timing.unavailable++;continue}timing.matched++;if(days<0)timing.pre++;else if(days<=7)timing.d0_7++;else if(days<=30)timing.d8_30++;else timing.over30++}
  const months=[5,6,7,8,9,10,11,0,1,2,3,4],names=['June','July','August','September','October','November','December','January','February','March','April','May'],triggers=[.01,.025,.04,.025,.01,.01,.01,.01,.01,.01,.01,.01];let cumulative=0;
  const monthly=months.map((m,i)=>{const cs=cases.filter(c=>c.date&&c.date.getMonth()===m);const nr=nsaidRows.filter(r=>r.date&&r.date.getMonth()===m);const heifer=cs.filter(c=>herdById.get(c.id)?.yearBorn===seasonYear-2).length;const mature=cs.filter(c=>herdById.has(c.id)&&herdById.get(c.id)?.yearBorn!==seasonYear-2).length;const rate=peak?cs.length/peak:null;if(rate!=null)cumulative+=rate;return{name:names[i],cases:cs.length,nsaid:nr.length,heifer,mature,rate,trigger:triggers[i],seasonPct:rate==null?null:cumulative,overTrigger:rate!=null&&rate>triggers[i]}});
  const unmatchedCowCases=cases.filter(c=>!herdById.has(c.id)).length;
  return{mastitisRows,antibioticRows,nsaidRows,cases,uniqueCows:byCow.size,repeatCows:repeatCowIds.length,repeatCowIds,repeatQuarterCows:repeatQuarterCowIds.length,repeatQuarterCowIds,combo,extended,txChanges,productCounts,quarters,timing,monthly,unmatchedCowCases};
}

function buildResult(){
  const errors=[];if(!state.herd)errors.push('Load the herd / SCC file.');if(!state.mastitis)errors.push('Load the current-season mastitis Treatment Register.');const ts=num(els.sccTs.value),la=num(els.sccLa.value);if(ts==null)errors.push('Enter the SCC low/high threshold.');if(la==null)errors.push('Enter the SCC cut-off for LA DCT.');if(ts!=null&&la!=null&&ts>=la)errors.push('The SCC low/high threshold should be lower than the LA DCT cut-off.');if(errors.length)throw new Error(errors.join('\n'));
  const seasonYear=+els.seasonYear.value,herdTests=+els.herdTests.value,hm=state.herd.map,mastitisRecords=treatmentRecords(state.mastitis,'mastitis'),mastitisByCow=indexByCow(mastitisRecords),dctRecords=treatmentRecords(state.dct,'dct'),mastaIndex=indexMastaplex(state.mastaplex);const cows=[];
  for(const r of state.herd.rows){const id=clean(r[hm.tag]);if(!id)continue;const scc=(hm.scc||[]).map(h=>num(r[h]));const cow={id,yearBorn:Math.trunc(num(r[hm.yearBorn])??0),calving:parseDate(r[hm.calving]),bcs:num(r[hm.bcs]),preg:clean(r[hm.preg]),expectedCalving:parseDate(r[hm.expectedCalving]),scc,treatments:mastitisByCow.get(id)||[],mastaplex:mastaIndex.get(id)||[]};cow.dct=dctAdvice(scc,cow.preg,cow.treatments,herdTests,ts,la);cow.dryPeriod=dryPeriodStatus(scc,herdTests,ts);Object.assign(cow,dryOffTiming(cow,seasonYear));const missing=[];if(!cow.preg)missing.push('pregnancy');if(cow.bcs==null)missing.push('BCS');if(!cow.expectedCalving)missing.push('expected calving');if(cow.scc.slice(0,herdTests).some(v=>v==null))missing.push('partial SCC');cow.dataNote=missing.join(', ');cows.push(cow)}
  const herdById=new Map(cows.map(c=>[c.id,c])),peak=num(els.peakCows.value)||cows.length,firstCalvers=num(els.firstCalvers.value)||cows.filter(c=>c.yearBorn===seasonYear-2).length,recent=cows.map(c=>c.scc[0]).filter(x=>x!=null&&x>0),sccBuckets={le100:0,b101_200:0,b201_499:0,ge500:0,over1000:0};for(const x of recent){if(x<=100)sccBuckets.le100++;else if(x<=200)sccBuckets.b101_200++;else if(x<500)sccBuckets.b201_499++;else sccBuckets.ge500++;if(x>1000)sccBuckets.over1000++}
  const dctCounts={};for(const c of cows)dctCounts[c.dct]=(dctCounts[c.dct]||0)+1;const clinical=dctCounts['Mastitis - Clinical']||0,subclinical=dctCounts['Mastitis - Subclinical']||0,sealantOnly=dctCounts['Teat Sealant only']||0,sa=dctCounts['SA DCT/Sealant Combo']||0,laBase=dctCounts['LA DCT/Sealant Combo']||0,laOrder=laBase+clinical+subclinical,sealantCows=sealantOnly+sa+laOrder,sealantTubes=sealantCows*4;const pregnant=cows.filter(c=>/pregnant/i.test(c.preg)).length;const pregnantSa=cows.filter(c=>/pregnant/i.test(c.preg)&&c.dct==='SA DCT/Sealant Combo').length;const pregnantLa=cows.filter(c=>/pregnant/i.test(c.preg)&&['LA DCT/Sealant Combo','Mastitis - Clinical','Mastitis - Subclinical'].includes(c.dct)).length;const pregnantSealant=cows.filter(c=>/pregnant/i.test(c.preg)&&c.dct==='Teat Sealant only').length;const emptyHeifers=cows.filter(c=>/empty/i.test(c.preg)&&c.yearBorn===seasonYear-2).length;
  const caseAnalysis=buildCaseAnalysis(mastitisRecords,herdById,seasonYear,peak,firstCalvers);
  const transitionCounts={'Cured':0,'Retained Infection':0,'New Infection':0,'Low SCC':0,'Missed H/T':0,'Not available':0};for(const c of cows)transitionCounts[c.dryPeriod.status]=(transitionCounts[c.dryPeriod.status]||0)+1;const transitionAvailable=cows.length-(transitionCounts['Not available']||0)-(transitionCounts['Missed H/T']||0);
  const prevProducts={},prevCowDct=new Set(),prevCowSeal=new Set();for(const r of dctRecords){if(r.drug)prevProducts[r.drug]=(prevProducts[r.drug]||0)+1;if(r.category==='DCT')prevCowDct.add(r.id);if(r.category==='Sealant')prevCowSeal.add(r.id)}const species={};for(const arr of mastaIndex.values())for(const s of arr)species[s]=(species[s]||0)+1;
  const expectedDates=cows.map(c=>c.expectedCalving).filter(Boolean),calvingStart=expectedDates.length?new Date(Math.min(...expectedDates.map(d=>d.getTime()))):null,midCalving=median(expectedDates);
  const coverage={preg:cows.filter(c=>c.preg).length,bcs:cows.filter(c=>c.bcs!=null).length,expected:cows.filter(c=>c.expectedCalving).length,latestScc:recent.length,transition:transitionAvailable,caseCalving:caseAnalysis.timing.matched,caseTotal:caseAnalysis.cases.length,dctLoaded:!!state.dct,mastaplexLoaded:!!state.mastaplex};
  return{cows,peak,firstCalvers,recent,sccBuckets,dctCounts,clinical,subclinical,sealantOnly,sa,laBase,laOrder,sealantCows,sealantTubes,pregnant,pregnantSa,pregnantLa,pregnantSealant,emptyHeifers,caseAnalysis,transitionCounts,transitionAvailable,dctRecords,prevProducts,prevCowDct,prevCowSeal,species,calvingStart,midCalving,coverage,threshold:ts,laThreshold:la};
}

function metricsHtml(items){return`<div class="metric-list">${items.map(([a,b,c])=>`<div class="metric-row"><span>${safe(a)}${c?`<small>${safe(c)}</small>`:''}</span><span>${safe(b)}</span></div>`).join('')}</div>`}
function barsHtml(items,total){return items.map(([name,n])=>`<div class="bar-row"><div class="bar-label"><span>${safe(name)}</span><strong>${fmt(n)}${total?` · ${pct(n,total)}`:''}</strong></div><div class="bar-track"><div class="bar-fill" style="width:${total?Math.min(100,n/total*100):0}%"></div></div></div>`).join('')}
function statusBadge(text,tone='neutral'){return`<span class="status-badge ${tone}">${safe(text)}</span>`}
function coverageBox(label,have,total,note=''){const available=total?have/total:null,tone=available==null?'neutral':available>=.95?'good':available>=.7?'warn':'bad';const value=total?`${fmt(have)} / ${fmt(total)} · ${pct(have,total)}`:safe(have);return`<div class="coverage-item ${tone}"><b>${safe(label)}</b><strong>${value}</strong>${note?`<span>${safe(note)}</span>`:''}</div>`}
function renderMonthly(rows){
  els.monthlyTable.querySelector('tbody').innerHTML=rows.map(x=>`<tr class="${x.overTrigger?'trigger-row':''}"><td>${x.name}</td><td>${fmt(x.cases)}</td><td>${fmt(x.nsaid)}</td><td>${fmt(x.heifer)}</td><td>${fmt(x.mature)}</td><td>${pct(x.rate,1)}</td><td>${pct(x.trigger,1)}</td><td>${pct(x.seasonPct,1)}</td></tr>`).join('')+`<tr class="total-row"><td>Total</td><td>${fmt(rows.reduce((a,x)=>a+x.cases,0))}</td><td>${fmt(rows.reduce((a,x)=>a+x.nsaid,0))}</td><td>${fmt(rows.reduce((a,x)=>a+x.heifer,0))}</td><td>${fmt(rows.reduce((a,x)=>a+x.mature,0))}</td><td>${pct(rows.reduce((a,x)=>a+x.cases,0),state.result?.peak)}</td><td>14.0%</td><td>${pct(rows.length?rows[rows.length-1].seasonPct:null,1)}</td></tr>`;
}
function renderResult(r){
  state.result=r;els.report.classList.remove('hidden');els.errorBox.classList.add('hidden');const farm=clean(els.farmName.value)||'Unnamed farm',vet=clean(els.vetName.value)||'Veterinarian not entered';els.reportTitle.textContent=farm;els.reportMeta.textContent=`${vet} · Season ${els.seasonYear.value}/${String(+els.seasonYear.value+1).slice(-2)} · Generated ${new Date().toLocaleDateString('en-NZ')}`;
  const herdDates=state.herd?.map?.sccDates||[],currentDates=herdDates.slice(0,+els.herdTests.value),previousDate=state.herd?.map?.previousHerdTest||herdDates[+els.herdTests.value]||'';const enteredStart=parseDate(els.calvingStart.value),expectedDry=parseDate(els.expectedDryOff.value);
  const facts=[['Veterinarian',vet],['Dairy company',clean(els.dairyCompany.value)||'N/A'],['Supply number',clean(els.supplyNumber.value)||'N/A'],['LIC / MyHerd',clean(els.ptptCode.value)||'N/A'],['Heifers teat sealed?',clean(els.heifersSealed.value)||'N/A'],['Previous-season BMSCC',clean(els.bmsccPrevious.value)||'N/A'],['Current-season BMSCC',clean(els.bmsccCurrent.value)||'N/A'],['Current herd tests',currentDates.length?currentDates.join(', '):els.herdTests.value],['Pre-dry comparison',previousDate||'N/A'],['Peak cows',fmt(r.peak)],['First calvers',fmt(r.firstCalvers)],['SCC cut-offs',`${r.threshold} / ${r.laThreshold}`],['Plan start of calving',isoDate(enteredStart||r.calvingStart)||'N/A'],['Expected mid-calving',isoDate(r.midCalving)||'N/A'],['Expected dry-off',isoDate(expectedDry)||'N/A']];els.consultFacts.innerHTML=facts.map(([a,b])=>`<div class="fact"><b>${safe(a)}</b><span>${safe(b)}</span></div>`).join('');

  els.dataQuality.innerHTML=[
    coverageBox('Pregnancy status',r.coverage.preg,r.cows.length,'Used for DCT/carryover classification'),
    coverageBox('BCS',r.coverage.bcs,r.cows.length,'Needed for individual dry-off timing'),
    coverageBox('Expected calving date',r.coverage.expected,r.cows.length,'Needed for dry-off timing'),
    coverageBox('Latest SCC',r.coverage.latestScc,r.cows.length,'Cows present at the latest selected H/T'),
    coverageBox('Dry-period SCC comparison',r.coverage.transition,r.cows.length,'Requires pre-dry + first post-calving SCC'),
    coverageBox('Case-to-calving match',r.coverage.caseCalving,r.coverage.caseTotal,'Unmatched cases remain in season totals'),
    coverageBox('Previous DCT / ITS',r.coverage.dctLoaded?'Loaded':'Not supplied',null,'Optional'),
    coverageBox('Mastaplex',r.coverage.mastaplexLoaded?'Loaded':'Not supplied',null,'Optional')
  ].join('');

  const seasonCases=r.caseAnalysis.cases.length,over200=r.sccBuckets.b201_499+r.sccBuckets.ge500,monthsOver=r.caseAnalysis.monthly.filter(x=>x.overTrigger).length,early30=r.caseAnalysis.timing.d0_7+r.caseAnalysis.timing.d8_30;
  els.kpis.innerHTML=[['Clinical case events',fmt(seasonCases)],['Season case %',pct(seasonCases,r.peak)],['Latest SCC >200',pct(over200,r.recent.length)],['Repeat cows',fmt(r.caseAnalysis.repeatCows)],['Retained post-DCT',fmt(r.transitionCounts['Retained Infection'])],['Sealant tubes',fmt(r.sealantTubes)]].map(([label,value])=>`<div class="kpi"><div class="value">${safe(value)}</div><div class="label">${safe(label)}</div></div>`).join('');

  const summary=[];
  summary.push(`There were ${fmt(seasonCases)} clinical antibiotic case events across ${fmt(r.caseAnalysis.uniqueCows)} cows (${pct(seasonCases,r.peak)} of peak cow numbers by case count).`);
  summary.push(`${fmt(monthsOver)} month${monthsOver===1?'':'s'} exceeded the monthly SmartSAMM-style trigger reference used in the spreadsheet.`);
  if(r.caseAnalysis.timing.matched)summary.push(`${fmt(early30)} of ${fmt(r.caseAnalysis.timing.matched)} cases with a usable calving-date match occurred within 30 days after calving; ${fmt(r.caseAnalysis.timing.unavailable)} case${r.caseAnalysis.timing.unavailable===1?' was':'s were'} unavailable for timing analysis.`);else summary.push('Case timing from calving is unavailable because no treatment cases could be matched to a usable season calving date.');
  if(r.transitionAvailable)summary.push(`Dry-period SCC outcome was calculable for ${fmt(r.transitionAvailable)} cows: ${fmt(r.transitionCounts['Cured'])} cured, ${fmt(r.transitionCounts['Retained Infection'])} retained infections, ${fmt(r.transitionCounts['New Infection'])} new infections and ${fmt(r.transitionCounts['Low SCC'])} remained low.`);else summary.push('Dry-period SCC cure/new/retained status could not be calculated from the supplied herd-test history.');
  if(r.coverage.preg < r.cows.length*0.95)summary.push(`Pregnancy status is available for ${fmt(r.coverage.preg)} of ${fmt(r.cows.length)} cows. The DCT order should be treated as provisional because empty cows cannot be reliably excluded where pregnancy status is missing.`);
  summary.push(`Current DCT estimate: ${fmt(r.sa)} SA combination cows, ${fmt(r.laOrder)} LA combination cows (including clinical/subclinical cases), ${fmt(r.sealantOnly)} teat-sealant-only cows and ${fmt(r.sealantTubes)} sealant tubes.`);
  els.executiveSummary.innerHTML=summary.map(x=>`<div class="summary-point"><span>•</span><p>${safe(x)}</p></div>`).join('');

  renderMonthly(r.caseAnalysis.monthly);
  const t=r.caseAnalysis.timing;els.caseTiming.innerHTML=metricsHtml([['0–7 days',fmt(t.d0_7),`${pct(t.d0_7,t.matched)} of matched cases`],['8–30 days',fmt(t.d8_30),`${pct(t.d8_30,t.matched)} of matched cases`],['Over 30 days',fmt(t.over30),`${pct(t.over30,t.matched)} of matched cases`],['Pre-calving',fmt(t.pre),`${pct(t.pre,t.matched)} of matched cases`],['Timing unavailable',fmt(t.unavailable),'Missing/unusable calving-date match'],['Usable timing coverage',`${fmt(t.matched)} / ${fmt(r.caseAnalysis.cases.length)}`,pct(t.matched,r.caseAnalysis.cases.length)]]);
  els.caseSummary.innerHTML=metricsHtml([['Cows treated with antibiotic',fmt(r.caseAnalysis.uniqueCows)],['Repeat cows',`${fmt(r.caseAnalysis.repeatCows)} · ${pct(r.caseAnalysis.repeatCows,r.peak)}`],['Repeat cow + same quarter',fmt(r.caseAnalysis.repeatQuarterCows)],['Ab + NSAID combination cases',`${fmt(r.caseAnalysis.combo)} · ${pct(r.caseAnalysis.combo,r.caseAnalysis.cases.length)}`],['Extended antibiotic cases (>3 doses)',fmt(r.caseAnalysis.extended)],['Treatment change within same case',fmt(r.caseAnalysis.txChanges)],['Treatment cases not matched to herd file',fmt(r.caseAnalysis.unmatchedCowCases)]]);
  const products=Object.entries(r.caseAnalysis.productCounts).sort((a,b)=>b[1]-a[1]);els.drugSummary.innerHTML=products.length?barsHtml(products,r.caseAnalysis.antibioticRows.length)+metricsHtml([['Antibiotic treatment rows',fmt(r.caseAnalysis.antibioticRows.length)],['NSAID treatment rows',fmt(r.caseAnalysis.nsaidRows.length)],['NSAID-paired case events',fmt(r.caseAnalysis.combo)]]):'<p class="na-copy">N/A — no antibiotic treatment products detected.</p>';
  const q=Object.entries(r.caseAnalysis.quarters).filter(([,n])=>n>0);els.quarterSummary.innerHTML=q.length?barsHtml(q,r.caseAnalysis.cases.length)+`<p class="context-note">Percentages use clinical case events as the denominator. A multi-quarter case is counted once under Multi.</p>`:'<p class="na-copy">N/A — no quarter information detected.</p>';
  els.sccSummary.innerHTML=r.recent.length?barsHtml([['≤100',r.sccBuckets.le100],['101–200',r.sccBuckets.b101_200],['201–499',r.sccBuckets.b201_499],['≥500',r.sccBuckets.ge500]],r.recent.length)+metricsHtml([['Over 1,000',fmt(r.sccBuckets.over1000)],['Cows with latest SCC',fmt(r.recent.length)],['Latest SCC >200',`${fmt(over200)} · ${pct(over200,r.recent.length)}`]]):'<p class="na-copy">N/A — no usable latest herd-test SCC values.</p>';
  const transitionItems=[['Cured',r.transitionCounts['Cured']],['Retained infection',r.transitionCounts['Retained Infection']],['New infection',r.transitionCounts['New Infection']],['Remained low',r.transitionCounts['Low SCC']]];els.sccTransitions.innerHTML=r.transitionAvailable?barsHtml(transitionItems,r.transitionAvailable)+metricsHtml([['Usable comparisons',`${fmt(r.transitionAvailable)} / ${fmt(r.cows.length)}`],['Missed first post-calving H/T',fmt(r.transitionCounts['Missed H/T'])],['Pre-dry comparison unavailable',fmt(r.transitionCounts['Not available'])],['Threshold used',`${fmt(r.threshold)} ×1000`]]):'<p class="na-copy">N/A — pre-dry and first post-calving SCC values were not both available.</p>';
  const pp=Object.entries(r.prevProducts).sort((a,b)=>b[1]-a[1]);els.previousDctSummary.innerHTML=state.dct?(pp.length?barsHtml(pp,Math.max(1,r.dctRecords.length))+metricsHtml([['Cows with DCT',fmt(r.prevCowDct.size)],['Cows with sealant',fmt(r.prevCowSeal.size)]]):'<p class="na-copy">Previous DCT / ITS file loaded, but no recognised DCT or sealant products were detected.</p>'):'<p class="na-copy">N/A — previous DCT / ITS data not supplied.</p>';
  const sp=Object.entries(r.species).sort((a,b)=>b[1]-a[1]);els.mastaplexSummary.innerHTML=state.mastaplex?(sp.length?barsHtml(sp,sp.reduce((a,x)=>a+x[1],0)):'<p class="na-copy">Mastaplex file loaded, but no cultured-growth results were detected.</p>'):'<p class="na-copy">N/A — Mastaplex data not supplied.</p>';

  const order=[['Teat Sealant only',r.sealantOnly],['SA DCT/Sealant Combo',r.sa],['LA DCT/Sealant Combo',r.laBase],['Mastitis - Clinical',r.clinical],['Mastitis - Subclinical',r.subclinical],['Empty',r.dctCounts.Empty||0],['Absent from H/T',r.dctCounts['Absent from H/T (Culled/Carryover)']||0]];els.dctSummary.innerHTML=barsHtml(order,r.cows.length);
  els.dctOrder.innerHTML=`<div class="order-box"><h4>Order estimate</h4>${metricsHtml([['SA DCT combo cows',fmt(r.sa)],['LA DCT combo cows',fmt(r.laOrder),'Includes clinical/subclinical cows'],['Sealant-only cows',fmt(r.sealantOnly)],['Total sealant cows',fmt(r.sealantCows)],['Total sealant tubes',fmt(r.sealantTubes)],['Pregnant cows',fmt(r.pregnant)],['Pregnant SA combo',fmt(r.pregnantSa)],['Pregnant LA combo',fmt(r.pregnantLa)],['Pregnant sealant only',fmt(r.pregnantSealant)],['Empty heifers / carryover candidates',fmt(r.emptyHeifers)],['Pregnancy-data coverage',`${fmt(r.coverage.preg)} / ${fmt(r.cows.length)}`,r.coverage.preg<r.cows.length?'Order provisional where pregnancy is missing':'Complete']])}</div>`;
  renderCowTable(r.cows);els.report.scrollIntoView({behavior:'smooth',block:'start'});
}
function renderCowTable(cows){
  const q=clean(els.cowSearch.value).toLowerCase(),filtered=cows.filter(c=>!q||[c.id,c.dct,c.advice,c.dryPeriod.status,c.preg,c.dataNote,...c.mastaplex].join(' ').toLowerCase().includes(q));els.cowCountText.textContent=`Showing ${filtered.length.toLocaleString()} of ${cows.length.toLocaleString()} cows`;
  els.cowTable.querySelector('tbody').innerHTML=filtered.map(c=>{const tone=c.dryPeriod.status==='Cured'||c.dryPeriod.status==='Low SCC'?'good':c.dryPeriod.status==='Retained Infection'||c.dryPeriod.status==='New Infection'?'warn':'neutral';return`<tr><td><span class="tag">${safe(c.id)}</span></td><td>${c.scc[0]??''}</td><td>${c.dryPeriod.pre??''}</td><td>${safe(c.preg||'N/A')}</td><td>${safe(c.dct)}</td><td>${safe(c.advice)}</td><td>${statusBadge(c.dryPeriod.status,tone)}</td><td>${safe(c.dataNote||'—')}</td><td>${safe(c.mastaplex.join(', '))}</td></tr>`}).join('');
}
els.cowSearch.addEventListener('input',()=>state.result&&renderCowTable(state.result.cows));
els.generateBtn.addEventListener('click',()=>{
  els.generateBtn.disabled=true;els.generateBtn.textContent='Generating…';els.generateStatus.textContent='';els.viewReportBtn.classList.add('hidden');
  try{renderResult(buildResult());els.generateBtn.textContent='Report generated ✓';els.generateStatus.textContent='The report is below.';els.viewReportBtn.classList.remove('hidden')}
  catch(e){els.errorBox.textContent=`Report could not be generated: ${e.message}`;els.errorBox.classList.remove('hidden');els.report.classList.add('hidden');els.generateBtn.textContent='Generate report';els.generateStatus.textContent='Check the message below.';els.errorBox.scrollIntoView({behavior:'smooth',block:'center'})}
  finally{els.generateBtn.disabled=false}
});
els.viewReportBtn.addEventListener('click',()=>els.report.scrollIntoView({behavior:'smooth',block:'start'}));
els.printBtn.addEventListener('click',()=>window.print());
els.csvBtn.addEventListener('click',()=>{if(!state.result)return;const rows=[['Tag','Latest SCC','Pre-dry SCC','Pregnancy Diagnosis','Expected Calving Date','BCS','DCT Advice','Dry Off Advice','Dry-period SCC Status','Data Note','Mastaplex']];for(const c of state.result.cows)rows.push([c.id,c.scc[0]??'',c.dryPeriod.pre??'',c.preg,isoDate(c.expectedCalving),c.bcs??'',c.dct,c.advice,c.dryPeriod.status,c.dataNote,c.mastaplex.join('; ')]);const csv=rows.map(r=>r.map(x=>`"${String(x??'').replace(/"/g,'""')}"`).join(',')).join('\r\n');const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download=`${(clean(els.farmName.value)||'farm').replace(/[^a-z0-9]+/gi,'-')}-dct-recommendations.csv`;a.click();URL.revokeObjectURL(a.href)});

function demoData(){
  els.farmName.value='Demo Dairy';els.vetName.value='Dr Example';els.seasonYear.value=2025;els.peakCows.value=20;els.firstCalvers.value=5;els.sccTs.value=150;els.sccLa.value=250;els.herdTests.value=4;els.dairyCompany.value='Demo Dairy Co';els.supplyNumber.value='12345';els.ptptCode.value='ABCD';els.heifersSealed.value='Y';els.bmsccPrevious.value=145;els.bmsccCurrent.value=128;els.calvingStart.value='2025-07-01';els.expectedDryOff.value='2026-05-20';updatePrescriptionFields();
  const herd=Array.from({length:20},(_,i)=>({'Tag':String(101+i),'Animal-Year Born':i<5?2023:2021,'Calving date':new Date(2025,6,(i%20)+1),'BCS':[4.2,4.5,4.7,5,5.2][i%5],'Most recent H/T':[82,116,188,275,520,73,146,205,340,110,620,95,155,230,410,90,180,260,780,130][i],'2nd':[95,120,170,260,480,80,160,190,315,135,590,120,140,250,390,100,160,240,650,145][i],'3rd':[110,130,160,240,450,88,180,175,290,150,550,140,130,270,360,120,145,220,580,160][i],'4th':[120,140,175,220,430,95,190,165,280,160,510,145,120,260,330,125,140,210,560,170][i],'5th':[130,155,140,200,400,110,150,155,250,170,470,160,110,240,300,140,130,190,520,180][i],'Pregnancy Diagnosis':i===18?'Empty':'Pregnant','Expected Calving date':new Date(2026,6,(i%20)+1)}));
  const mastitis=[{'Cow Number':'103','Condition':'Mastitis - Clinical','Treatment Date':new Date(2025,7,12),'Quarter':'LF','Product':'Intracillin 1000 Milking Cow','Days':3},{'Cow Number':'103','Condition':'Mastitis - Clinical','Treatment Date':new Date(2025,7,12),'Quarter':'','Product':'Metacam 20mg/Ml For Injection','Days':1},{'Cow Number':'103','Condition':'Mastitis - Clinical','Treatment Date':new Date(2025,9,4),'Quarter':'LF','Product':'Intracillin 1000 Milking Cow','Days':3},{'Cow Number':'107','Condition':'Mastitis - Clinical','Treatment Date':new Date(2025,8,7),'Quarter':'RR','Product':'Intracillin 1000 Milking Cow','Days':5},{'Cow Number':'111','Condition':'Mastitis - Subclinical','Treatment Date':new Date(2025,10,10),'Quarter':'RF','Product':'Albiotic','Days':3},{'Cow Number':'111','Condition':'Mastitis - Subclinical','Treatment Date':new Date(2025,10,10),'Quarter':'','Product':'Metacam 20mg/Ml For Injection','Days':1}];
  const dct=[{'Cow Number':'101','Treatment Date':new Date(2025,4,20),'Product':'Dryzen'},{'Cow Number':'101','Treatment Date':new Date(2025,4,20),'Product':'Duramast DC 500'},{'Cow Number':'102','Treatment Date':new Date(2025,4,20),'Product':'Dryzen'},{'Cow Number':'102','Treatment Date':new Date(2025,4,20),'Product':'Quadrant DC'}];const mp=[{'Cow Number':'103','Species':'Staph. aureus','Quarter':'LF'},{'Cow Number':'107','Species':'Strep. uberis','Quarter':'RR'},{'Cow Number':'111','Species':'CNS','Quarter':'RF'}];
  state.herd={rows:herd,map:detectHerd(herd),name:'Demo herd data'};state.mastitis={rows:mastitis,map:detectTreatment(mastitis),name:'Demo mastitis treatments'};state.dct={rows:dct,map:detectTreatment(dct),name:'Demo DCT / ITS'};state.mastaplex={rows:mp,map:detectMastaplex(mp),name:'Demo Mastaplex'};for(const k of ['herd','mastitis','dct','mastaplex'])setStatus(k,state[k].rows,state[k].map,state[k].name);try{renderResult(buildResult());els.generateBtn.textContent='Report generated ✓';els.generateStatus.textContent='Demo report loaded below.';els.viewReportBtn.classList.remove('hidden')}catch(e){els.errorBox.textContent=`Demo report could not be generated: ${e.message}`;els.errorBox.classList.remove('hidden');els.report.classList.add('hidden')}
}
els.demoBtn.addEventListener('click',demoData);
if('serviceWorker'in navigator&&location.protocol.startsWith('http'))navigator.serviceWorker.register('./sw.js').catch(()=>{});
