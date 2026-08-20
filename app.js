'use strict';

const state={herd:null,mastitis:null,dct:null,mastaplex:null,result:null};
const ids=['farmName','vetName','seasonYear','herdTests','herdTestsGuard','peakCows','firstCalvers','sccTs','sccLa','dairyCompany','supplyNumber','ptptCode','heifersSealed','bmsccPrevious','bmsccCurrent','calvingStart','expectedDryOff','prescriptionStatus','herdFile','mastitisFile','dctFile','mastaplexFile','herdStatus','mastitisStatus','dctStatus','mastaplexStatus','mappingPanel','generateBtn','viewReportBtn','generateStatus','demoBtn','errorBox','report','reportTitle','reportMeta','consultFacts','dataQuality','kpis','executiveSummary','mastitisChart','mastitisChartTooltip','monthlyTable','caseTiming','caseSummary','drugSummary','sccSummary','sccTransitions','quarterSummary','mastaplexSummary','previousDctSummary','dctSummary','dctOrder','cowTable','cowCountText','cowSearch','cowCsvBtn','csvBtn','printBtn','newFarmBtn'];
const els=Object.fromEntries(ids.map(id=>[id,document.getElementById(id)]));
els.seasonYear.value=new Date().getFullYear()-1;
function defaultCalvingStartFromSeason(){
  const y=Number(els.seasonYear.value);
  if(Number.isFinite(y)&&y>=2000&&y<=2100){
    els.calvingStart.value=`${y+1}-08-01`;
  }
}
defaultCalvingStartFromSeason();
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
  if(v instanceof Date&&!isNaN(v))return new Date(v.getFullYear(),v.getMonth(),v.getDate());
  const fromExcelSerial=n=>{if(!Number.isFinite(n)||n<=20000||n>=80000)return null;const d=new Date(Date.UTC(1899,11,30));d.setUTCDate(d.getUTCDate()+n);return new Date(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate())};
  if(typeof v==='number'){const excel=fromExcelSerial(v);if(excel)return excel;}
  const s=clean(v);if(!s)return null;
  if(/^\d{5}(?:\.\d+)?$/.test(s)){const excel=fromExcelSerial(Number(s));if(excel)return excel;}
  let m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:\s|T|$)/);
  if(m){let y=+m[3];if(y<100)y+=2000;const d=new Date(y,+m[2]-1,+m[1]);if(d.getFullYear()===y&&d.getMonth()===+m[2]-1&&d.getDate()===+m[1])return d;}
  m=s.match(/^(20\d{2})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:\s|T|$)/);
  if(m){const d=new Date(+m[1],+m[2]-1,+m[3]);if(d.getFullYear()===+m[1]&&d.getMonth()===+m[2]-1&&d.getDate()===+m[3])return d;}
  const months={jan:0,january:0,feb:1,february:1,mar:2,march:2,apr:3,april:3,may:4,jun:5,june:5,jul:6,july:6,aug:7,august:7,sep:8,sept:8,september:8,oct:9,october:9,nov:10,november:10,dec:11,december:11};
  m=s.match(/^(\d{1,2})\s+([A-Za-z]{3,9})[\s,.-]+(\d{2,4})/);
  if(m){let y=+m[3];if(y<100)y+=2000;const mo=months[m[2].toLowerCase()];if(mo!=null){const d=new Date(y,mo,+m[1]);if(d.getFullYear()===y&&d.getMonth()===mo&&d.getDate()===+m[1])return d;}}
  m=s.match(/^([A-Za-z]{3,9})\s+(\d{1,2})[\s,.-]+(\d{2,4})/);
  if(m){let y=+m[3];if(y<100)y+=2000;const mo=months[m[1].toLowerCase()];if(mo!=null){const d=new Date(y,mo,+m[2]);if(d.getFullYear()===y&&d.getMonth()===mo&&d.getDate()===+m[2])return d;}}
  const d=new Date(s);return isNaN(d)?null:new Date(d.getFullYear(),d.getMonth(),d.getDate());
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
  const s=clean(header);
  // Only inspect SCC / herd-test headings, but accept the common MINDA date styles:
  // Apr 2025 SCC, 24 Apr 2025 SCC, 24/04/2025 SCC and 2025-04-24 SCC.
  if(!/(scc|somatic|herd\s*test)/i.test(s))return null;
  const months={jan:0,january:0,feb:1,february:1,mar:2,march:2,apr:3,april:3,may:4,jun:5,june:5,jul:6,july:6,aug:7,august:7,sep:8,sept:8,september:8,oct:9,october:9,nov:10,november:10,dec:11,december:11};
  let m=s.match(/\b(20\d{2})[\/.-](\d{1,2})[\/.-](\d{1,2})\b/);
  if(m){const d=new Date(+m[1],+m[2]-1,+m[3]);return isNaN(d)?null:d;}
  m=s.match(/\b(\d{1,2})[\/.-](\d{1,2})[\/.-](20\d{2})\b/);
  if(m){const d=new Date(+m[3],+m[2]-1,+m[1]);return isNaN(d)?null:d;}
  m=s.match(/\b(\d{1,2})?\s*([A-Za-z]{3,9})[\s,.-]+(20\d{2})\b/i);
  if(m){const month=months[m[2].toLowerCase()];if(month!=null){const d=new Date(+m[3],month,m[1]?+m[1]:1);return isNaN(d)?null:d;}}
  m=s.match(/\b([A-Za-z]{3,9})[\s,.-]+(20\d{2})\b/i);
  if(m){const month=months[m[1].toLowerCase()];if(month!=null)return new Date(+m[2],month,1);}
  return null;
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
    const selected=[...current,...(previous?[previous]:[])];map.scc=selected.map(x=>x.header);map.sccDates=selected.map(x=>sccDateLabel(x.date));map.sccAll=dated.map(x=>x.header);map.sccAllDates=dated.map(x=>sccDateLabel(x.date));map.currentHerdTests=current.length;map.previousHerdTest=previous?sccDateLabel(previous.date):'';map.herdTestMode='dated';
  }else{
    const explicit=['Most recent H/T','2nd','3rd','4th','5th','6th','7th','8th','9th','10th'].map(x=>findHeader(h,[x])).filter(Boolean);
    const sccLike=h.filter(x=>/scc|herd.?test|h\/t/i.test(x)&&!Object.values(map).includes(x));
    map.scc=[...new Set([...explicit,...sccLike])].slice(0,11);
    // With generic MINDA H/T labels, reserve the oldest available column as the
    // previous/pre-dry comparison when there is more than one SCC column.
    map.currentHerdTests=Math.min(10,Math.max(1,map.scc.length-(map.scc.length>1?1:0)));
    map.previousHerdTestHeader=map.scc[map.currentHerdTests]||'';map.herdTestMode='generic';
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
function detectMastaplex(rows){
  const h=headersOf(rows);
  const map={
    tag:findHeader(h,['Tag','Animal ID','Cow Number','Cow No','Animal','ID','Cow'],['tag','cow','animalid']),
    species:findHeader(h,['Species','Organism','Growth','Result','Bacteria','Pathogen','Identification'],['species','organism','growth','bacteria','pathogen','identification']),
    quarter:findHeader(h,['Quarter','Affected Quarter'],['quarter']),
    date:findHeader(h,['Sample Date','Test Date','Collection Date','Culture Date','Result Date','Date','Date/Time','Date Time','Test Date/Time','Sample Date/Time','Created Date','Created At','Date Created','Started Date','Start Date','Test Started','Started At','Completed Date','Completed At'],['sampledate','testdate','collectiondate','culturedate','resultdate','datetime','createddate','createdat','datecreated','startdate','teststarted','startedat','completeddate','completedat'])
  };
  if(!map.date){
    const candidates=h.filter(x=>/(date|time|created|started|start|collected|collection|completed)/i.test(x));
    let best=null,bestScore=0;
    for(const header of candidates){
      const vals=rows.slice(0,200).map(r=>r[header]).filter(v=>clean(v));
      if(!vals.length)continue;
      const parsed=vals.filter(v=>parseDate(v)).length;
      const score=parsed/vals.length;
      if(score>bestScore){bestScore=score;best=header;}
    }
    if(best&&bestScore>=0.5)map.date=best;
  }
  if(!map.tag&&h.length>=3)map.tag=h[2];
  if(!map.species&&h.length>=5)map.species=h[4];
  return map;
}
function confidence(kind,map){const req=kind==='herd'?['tag','scc']:kind==='mastaplex'?['tag','species','date']:['tag','date','drug'];return req.filter(k=>Array.isArray(map[k])?map[k].length:map[k]).length/req.length}
function setStatus(kind,rows,map,name){
  const el=els[`${kind}Status`],c=confidence(kind,map);el.className=`file-status ${c>=.99?'good':c>=.5?'warn':'bad'}`;
  if(kind==='herd'&&map.sccAllDates?.length){const selected=(map.sccDates||[]).join(', ');el.textContent=`${rows.length.toLocaleString()} rows · ${map.sccAllDates.length} dated herd tests detected · using ${selected} · ${name}`}
  else if(kind==='mastaplex')el.textContent=`${rows.length.toLocaleString()} rows · ${c>=.99?'columns recognised':c>=.5?'check detected columns':'column match poor'} · date column: ${map.date||'NOT FOUND'} · ${name}`;
  else el.textContent=`${rows.length.toLocaleString()} rows · ${c>=.99?'columns recognised':c>=.5?'check detected columns':'column match poor'} · ${name}`;
  renderMappings();
}
function renderMappings(){const labels={herd:'Herd',mastitis:'Mastitis',dct:'DCT / ITS',mastaplex:'Mastaplex'};const items=[];for(const kind of Object.keys(labels)){const d=state[kind];if(!d)continue;const parts=Object.entries(d.map).filter(([k])=>!['sccAll'].includes(k)).map(([k,v])=>`${k}: ${Array.isArray(v)?v.join(', '):(v||'not found')}`);items.push(`<div class="mapping-row"><strong>${labels[kind]}:</strong><span>${safe(parts.join(' · '))}</span></div>`)}els.mappingPanel.innerHTML=items.join('');els.mappingPanel.classList.toggle('hidden',!items.length)}
function syncHerdTestSafety(){
  const guard=els.herdTestsGuard,select=els.herdTests;
  if(!guard||!select)return;
  if(!state.herd){select.disabled=false;guard.className='herd-test-guard neutral';guard.textContent='Load the herd / SCC file to verify this automatically.';return;}
  const map=state.herd.map||{},count=map.currentHerdTests||0,total=(map.scc||[]).length;
  if(map.herdTestMode==='dated'){
    select.disabled=true;
    if(count>=1&&count<=10){select.value=String(count);guard.className='herd-test-guard good';guard.textContent=`Auto-detected and locked: ${count} current-season herd test${count===1?'':'s'} from dated MINDA SCC headings.`;}
    else if(count>10){guard.className='herd-test-guard bad';guard.textContent=`${count} current-season herd tests were detected, but this build supports up to 10. Review the source file before generating recommendations.`;}else{guard.className='herd-test-guard bad';guard.textContent='No dated herd tests fall inside the selected 1 June–31 May season. Check the Season start year before generating the report.';}
    return;
  }
  select.disabled=false;
  if(count>=1&&count<=10&&!select.dataset.userSet)select.value=String(count);
  const chosen=+select.value||0,pre=(map.scc||[])[chosen]||'';
  guard.className='herd-test-guard warn';
  guard.textContent=total?`Manual confirmation required: ${total} SCC/H/T columns detected. With ${chosen} current-season test${chosen===1?'':'s'}, ${pre?`“${pre}” will be used as the pre-dry comparison.`:'no separate pre-dry SCC column is available.'}`:'Manual confirmation required: dated SCC headings were not detected.';
}
async function handleFile(kind,file){if(!file)return;try{const rows=await readTabular(file,kind);if(!rows.length)throw new Error('No data rows found.');const map=kind==='herd'?detectHerd(rows):kind==='mastaplex'?detectMastaplex(rows):detectTreatment(rows);state[kind]={rows,map,name:file.name};if(kind==='herd'){delete els.herdTests.dataset.userSet;syncHerdTestSafety()}setStatus(kind,rows,map,file.name)}catch(e){const el=els[`${kind}Status`];el.className='file-status bad';el.textContent=e.message}}
els.herdFile.addEventListener('change',e=>handleFile('herd',e.target.files[0]));els.mastitisFile.addEventListener('change',e=>handleFile('mastitis',e.target.files[0]));els.dctFile.addEventListener('change',e=>handleFile('dct',e.target.files[0]));els.mastaplexFile.addEventListener('change',e=>handleFile('mastaplex',e.target.files[0]));
els.herdTests.addEventListener('change',()=>{els.herdTests.dataset.userSet='1';syncHerdTestSafety()});

els.seasonYear.addEventListener('change',()=>{defaultCalvingStartFromSeason();updatePrescriptionFields();if(!state.herd)return;state.herd.map=detectHerd(state.herd.rows);delete els.herdTests.dataset.userSet;syncHerdTestSafety();setStatus('herd',state.herd.rows,state.herd.map,state.herd.name)});

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
// LIC / MyHerd PTPT codes are conventionally entered in capitals. Normalise the
// stored value as the user types so reports and exports also receive uppercase text.
els.ptptCode?.addEventListener('input',e=>{
  const input=e.target,start=input.selectionStart,end=input.selectionEnd;
  const upper=input.value.toUpperCase();
  if(input.value!==upper){input.value=upper;try{input.setSelectionRange(start,end)}catch{}}
});
for(const id of prescriptionFieldIds){els[id]?.addEventListener('input',updatePrescriptionFields);els[id]?.addEventListener('change',updatePrescriptionFields)}
updatePrescriptionFields();

const DCT_DRUGS=['quadrant dc','duramast dc 500','duramast dc 600','cepravin dry cow','cefa-safe','dryclox dc','ceprotect dc','juraclox','orbenin dry','noroclox dc 600'];
const SEALANTS=['dryzen','duraseal','dryseal','teatseal','sureseal'];
const NSAIDS=['metacam 20mg/ml for injection','metacam 20mg/ml','key injection','norflunix','metacam'];
function isNSAID(drug){const d=clean(drug).toLowerCase();return NSAIDS.some(x=>d.includes(x))}
function isDctDrug(drug){const d=clean(drug).toLowerCase();return DCT_DRUGS.some(x=>d.includes(x))}
function isSealant(drug){const d=clean(drug).toLowerCase();return SEALANTS.some(x=>d.includes(x))}

// AMR reference catalogue. Categories follow the NZVA antimicrobial traffic-light
// guidance checked August 2026. Product aliases are stored locally so the report
// remains fully offline. Combination products use the highest recognised category.
const AMR_CATALOGUE_VERSION='NZVA traffic-light guidance · checked Aug 2026';
const AMR_CATALOGUE=[
  {aliases:['intracillin 1000 milking cow','intracillin 1000','intracillin mc','masticillin'],active:'Procaine penicillin G',category:'green'},
  {aliases:['mamyzin','penthaone'],active:'Penethamate hydriodide',category:'green'},
  {aliases:['engemycin','bivatop 200','bivatop','oxytetracycline'],active:'Oxytetracycline',category:'green'},
  {aliases:['penclox'],active:'Procaine penicillin G + cloxacillin',category:'yellow'},
  {aliases:['clavulox lc','synulox lc','combiclav lc'],active:'Amoxicillin + clavulanic acid',category:'yellow'},
  {aliases:['albiotic'],active:'Lincomycin + neomycin',category:'yellow'},
  {aliases:['mastiplan'],active:'Cephapirin',category:'yellow'},
  {aliases:['rilexine'],active:'Cephalexin (1st-generation cephalosporin)',category:'yellow'},
  {aliases:['cepravin','quadrant dc','ceprotect'],active:'Cefalonium (1st-generation cephalosporin)',category:'yellow'},
  {aliases:['cefa-safe','cefa safe'],active:'Cephapirin (1st-generation cephalosporin)',category:'yellow'},
  {aliases:['dryclox dc','duramast dc 500'],active:'Ampicillin + cloxacillin',category:'yellow'},
  {aliases:['duramast dc 600','juraclox','orbenin dry','orbenin enduro','orbenin la','noroclox dc 600','cloxamp dc'],active:'Cloxacillin',category:'yellow'},
  {aliases:['amphoprim'],active:'Trimethoprim + sulphonamide',category:'yellow'},
  {aliases:['mastalone'],active:'Oleandomycin + oxytetracycline + neomycin',category:'red'},
  {aliases:['tylan','tylosin'],active:'Tylosin (macrolide)',category:'red'},
  {aliases:['cobactan','cefquinome'],active:'Cefquinome (4th-generation cephalosporin)',category:'red'},
  {aliases:['excenel','excede','ceftiofur'],active:'Ceftiofur (3rd-generation cephalosporin)',category:'red'},
  {aliases:['baytril','enrofloxacin'],active:'Enrofloxacin (fluoroquinolone)',category:'red'},
  {aliases:['gentamicin','gentamycin'],active:'Gentamicin',category:'red'},
  {aliases:['teatseal','sureseal','u-seal','useal','dryseal','duraseal','dryzen'],active:'Internal teat sealant · non-antibiotic',category:'non-antibiotic'}
];
function amrInfo(drug){
  const d=clean(drug).toLowerCase().replace(/[®™]/g,'').replace(/\s+/g,' ').trim();
  if(!d)return{category:'unclassified',active:'Active ingredient not identified'};
  const exact=AMR_CATALOGUE.find(item=>item.aliases.some(alias=>d.includes(alias)));
  if(exact)return exact;
  // Generic active-ingredient fallbacks help with brands that include the ingredient
  // in their exported product name, while still refusing to guess unknown brands.
  if(/procaine penicillin|penicillin g procaine|penethamate|tetracycline|oxytetracycline/.test(d))return{category:'green',active:'Active ingredient recognised from product name'};
  if(/cloxacillin|ampicillin|cephapirin|cefapirin|cefalonium|cephalonium|cephalexin|cefalexin|lincomycin/.test(d))return{category:'yellow',active:'Active ingredient recognised from product name'};
  if((/amoxicillin|amoxycillin/.test(d)&&/clavulan/.test(d))||(/trimethoprim/.test(d)&&/sulfa|sulpho/.test(d)))return{category:'yellow',active:'Active ingredient combination recognised from product name'};
  if(/ceftiofur|cefquinome|enrofloxacin|marbofloxacin|danofloxacin|tylosin|erythromycin|spiramycin|oleandomycin|gentamicin|gentamycin/.test(d))return{category:'red',active:'Active ingredient recognised from product name'};
  if(isSealant(drug))return{category:'non-antibiotic',active:'Internal teat sealant · non-antibiotic'};
  return{category:'unclassified',active:'AMR category not classified'};
}
function amrBadgeHtml(info){const labels={green:'GREEN',yellow:'YELLOW',red:'RED','non-antibiotic':'NON-ANTIBIOTIC',unclassified:'UNCLASSIFIED'};return`<span class="amr-badge amr-${info.category}">${labels[info.category]||'UNCLASSIFIED'}</span>`}
function amrBarsHtml(items,total){return items.map(([name,n])=>{const info=amrInfo(name),width=total?Math.min(100,n/total*100):0;return`<div class="bar-row amr-row"><div class="bar-label"><span class="amr-product"><span class="amr-product-title">${safe(name)} ${amrBadgeHtml(info)}</span><small>${safe(info.active)}</small></span><span class="bar-value"><strong>${fmt(n)}</strong>${total?`<small>${pct(n,total)}</small>`:''}</span></div><div class="bar-track"><div class="bar-fill amr-fill amr-${info.category}" style="width:${width}%"></div></div></div>`}).join('')}
function classifyMastitis(cat){const c=clean(cat).toLowerCase();if(c.includes('subclinical'))return'Mastitis - Subclinical';if(c.includes('clinical')||c.includes('mastitis'))return'Mastitis - Clinical';return''}
function isSevereClinicalMastitis(cat){const c=clean(cat).toLowerCase().replace(/[-_/]+/g,' ').replace(/\s+/g,' ').trim();return c.includes('mastitis')&&(c.includes('black')||c.includes('necrotic'))}
function parseQuarters(v){const raw=clean(v).toUpperCase();if(!raw)return[];return [...new Set(raw.split(/[^A-Z]+/).filter(q=>['LF','RF','LR','RR'].includes(q)))]}
function dairySeasonWindow(seasonYear){const start=new Date(seasonYear,5,1),end=new Date(seasonYear+1,5,1);return{start,end,label:`1 Jun ${seasonYear}–31 May ${seasonYear+1}`}}
function inDairySeason(date,seasonYear){if(!date)return false;const {start,end}=dairySeasonWindow(seasonYear);return date>=start&&date<end}
function treatmentRecords(data,mode){
  const out=[],m=data?.map||{};
  (data?.rows||[]).forEach((r,i)=>{const id=clean(r[m.tag]);if(!id)return;const drug=clean(r[m.drug]),cat=clean(r[m.category]),quarter=clean(r[m.quarter]).toUpperCase();out.push({id,date:parseDate(r[m.date]),quarter,quarters:parseQuarters(quarter),drug,duration:num(r[m.duration]),historic:m.historic?clean(r[m.historic]):'',rawCategory:cat,forceClinicalCase:mode==='mastitis'&&isSevereClinicalMastitis(cat),category:mode==='mastitis'?classifyMastitis(cat):isDctDrug(drug)?'DCT':isSealant(drug)?'Sealant':'',raw:r,index:i})});
  return out;
}
function indexByCow(records){const m=new Map();for(const r of records){if(!m.has(r.id))m.set(r.id,[]);m.get(r.id).push(r)}return m}
function indexMastaplex(data,seasonYear){
  const byCow=new Map(),m=data?.map||{};
  let recognised=0,inSeason=0,outOfSeason=0,undated=0,firstParsed=null,lastParsed=null;
  for(const r of data?.rows||[]){
    const id=clean(r[m.tag]),species=clean(r[m.species]);
    if(!id||!species)continue;
    recognised++;
    const date=m.date?parseDate(r[m.date]):null;
    if(!date){undated++;continue;}
    if(!firstParsed||date<firstParsed)firstParsed=date;
    if(!lastParsed||date>lastParsed)lastParsed=date;
    if(!inDairySeason(date,seasonYear)){outOfSeason++;continue;}
    inSeason++;
    if(!byCow.has(id))byCow.set(id,[]);
    byCow.get(id).push(species);
  }
  return{byCow,recognised,inSeason,outOfSeason,undated,dateHeader:m.date||'',firstParsed,lastParsed};
}
function mastaplexGroup(result){
  const s=clean(result).toLowerCase().replace(/[._-]+/g,' ').replace(/\s+/g,' ').trim();
  if(!s)return'Other / unclassified';
  if(/no (bacterial )?growth|no growth|negative culture|nothing isolated/.test(s))return'No bacterial growth';
  if(/mixed|contamin/.test(s))return'Mixed / contaminated';
  if(/gram\s*negative|gram\s*-ve|e\.?\s*coli|escherichia|klebsiella|serratia|enterobacter|pseudomonas|proteus|citrobacter/.test(s))return'Gram negative';
  if(/gram\s*positive|gram\s*\+ve|staph|staphyl|coagulase|\bcns\b|\bnas\b|strep|strept|enterococcus|aerococcus|corynebacter|bacillus|trueperella|arcanobacter/.test(s))return'Gram positive';
  return'Other / unclassified';
}

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
  if(/empty/i.test(cow.preg))return{advice:'Cull/Carryover',targetDate:null,timingBasis:'empty'};
  if(!cow.expectedCalving)return{advice:'N/A — expected calving date missing',targetDate:null,timingBasis:'unavailable'};

  // When BCS is available, retain the established BCS-adjusted timing calculation.
  // When BCS is unavailable, fall back to a standard 50-day dry period from the
  // cow's expected calving date so a useful dry-off recommendation can still be made.
  const target=cow.yearBorn===seasonYear-2?5.5:5;
  const hasBcs=cow.bcs!=null;
  const days=hasBcs ? (target-cow.bcs-.3)/.5*30+50 : 50;
  const d=new Date(cow.expectedCalving);
  d.setDate(d.getDate()-Math.round(days));
  const adjustedMonth=d.getDate()>15?d.getMonth()+1:d.getMonth();
  let advice='';
  if(adjustedMonth===3)advice='1st April Dry Off';else if(adjustedMonth===4)advice='1st May Dry Off';else if(adjustedMonth>=5)advice='Dry Off End of Season';else advice='Dry Off NOW';
  return{advice,targetDate:d,targetBCS:hasBcs?target:null,days,timingBasis:hasBcs?'BCS + expected calving':'Expected calving date only'};
}

function buildCaseAnalysis(records,herdById,seasonYear,peak,firstCalvers){
  const mastitisRows=records.filter(r=>r.category);const antibioticRows=mastitisRows.filter(r=>!isNSAID(r.drug));const nsaidRows=mastitisRows.filter(r=>isNSAID(r.drug));
  const groups=new Map();
  for(const r of mastitisRows){const key=`${r.id}|${dateKey(r.date)||'nodate-'+r.index}`;if(!groups.has(key))groups.set(key,{id:r.id,date:r.date,rows:[],quarters:new Set(),antibiotics:new Set(),nsaid:false,maxDoses:0,forceClinicalCase:false});const g=groups.get(key);g.rows.push(r);if(r.forceClinicalCase)g.forceClinicalCase=true;for(const qt of r.quarters)g.quarters.add(qt);if(isNSAID(r.drug))g.nsaid=true;else{if(r.drug)g.antibiotics.add(r.drug);g.maxDoses=Math.max(g.maxDoses,r.duration||0)}}
  // Standard mastitis events continue to require a recorded non-NSAID treatment row.
  // Black mastitis / enzootic necrotic mastitis are severe clinical diagnoses and
  // must count as cases even when treatment is supportive, euthanasia, or blank.
  const cases=[...groups.values()].filter(g=>g.antibiotics.size>0||g.forceClinicalCase).sort((a,b)=>(a.date||0)-(b.date||0));
  const byCow=new Map();for(const c of cases){if(!byCow.has(c.id))byCow.set(c.id,[]);byCow.get(c.id).push(c)}
  const repeatCowIds=[...byCow.entries()].filter(([,v])=>v.length>1).map(([id])=>id);
  let repeatQuarterCowIds=[];for(const [id,arr] of byCow){if(arr.length<2)continue;let repeat=false;const seen=new Set();for(const c of arr){for(const q of c.quarters){if(seen.has(q))repeat=true;seen.add(q)}}if(repeat)repeatQuarterCowIds.push(id)}
  const combo=cases.filter(c=>c.nsaid).length,extended=cases.filter(c=>c.maxDoses>3).length,txChanges=cases.filter(c=>c.antibiotics.size>1).length;
  const productCounts={};for(const r of antibioticRows){if(r.drug)productCounts[r.drug]=(productCounts[r.drug]||0)+1}
  const quarters={LF:0,RF:0,LR:0,RR:0,Multi:0,'No quarter':0};for(const c of cases){const qs=[...c.quarters];if(qs.length>1)quarters.Multi++;else if(qs.length===1)quarters[qs[0]]++;else quarters['No quarter']++}
  const timing={d0_7:0,d8_30:0,over30:0,pre:0,unavailable:0,matched:0};
  for(const c of cases){const cow=herdById.get(c.id);if(!cow||!cow.calving||!c.date){timing.unavailable++;continue}const days=daysBetween(c.date,cow.calving);if(days==null){timing.unavailable++;continue}if(days < -30){timing.unavailable++;continue}timing.matched++;if(days<0)timing.pre++;else if(days<=7)timing.d0_7++;else if(days<=30)timing.d8_30++;else timing.over30++}
  const months=[5,6,7,8,9,10,11,0,1,2,3,4],names=['June','July','August','September','October','November','December','January','February','March','April','May'],triggers=[.01,.025,.04,.025,.01,.01,.01,.01,.01,.01,.01,.01];let cumulative=0;
  const matureCows=peak!=null&&firstCalvers!=null?Math.max(0,peak-firstCalvers):null;
  const monthly=months.map((m,i)=>{const cs=cases.filter(c=>c.date&&c.date.getMonth()===m);const nr=nsaidRows.filter(r=>r.date&&r.date.getMonth()===m);const heifer=cs.filter(c=>herdById.get(c.id)?.yearBorn===seasonYear-2).length;const mature=cs.filter(c=>herdById.has(c.id)&&herdById.get(c.id)?.yearBorn!==seasonYear-2).length;const rate=peak?cs.length/peak:null;const heiferRate=firstCalvers?heifer/firstCalvers:null;const matureRate=matureCows?mature/matureCows:null;if(rate!=null)cumulative+=rate;return{name:names[i],cases:cs.length,nsaid:nr.length,heifer,mature,rate,heiferRate,matureRate,trigger:triggers[i],seasonPct:rate==null?null:cumulative,overTrigger:rate!=null&&rate>triggers[i]}});
  const unmatchedCowCases=cases.filter(c=>!herdById.has(c.id)).length;
  return{mastitisRows,antibioticRows,nsaidRows,cases,uniqueCows:byCow.size,repeatCows:repeatCowIds.length,repeatCowIds,repeatQuarterCows:repeatQuarterCowIds.length,repeatQuarterCowIds,combo,extended,txChanges,productCounts,quarters,timing,monthly,unmatchedCowCases};
}

function buildResult(){
  const errors=[];if(!state.herd)errors.push('Load the herd / SCC file.');if(!state.mastitis)errors.push('Load the current-season mastitis Treatment Register.');const ts=num(els.sccTs.value),la=num(els.sccLa.value);if(ts==null)errors.push('Enter the SCC low/high threshold.');if(la==null)errors.push('Enter the SCC cut-off for LA DCT.');if(ts!=null&&la!=null&&ts>=la)errors.push('The SCC low/high threshold should be lower than the LA DCT cut-off.');
  const hm=state.herd?.map||{},detectedHerdTests=hm.herdTestMode==='dated'?hm.currentHerdTests:null,manualHerdTests=+els.herdTests.value;
  if(hm.herdTestMode==='dated'&&(!detectedHerdTests||detectedHerdTests<1))errors.push('No dated herd tests fall inside the selected 1 June–31 May season. Check the Season start year.');
  if(hm.herdTestMode==='dated'&&detectedHerdTests>10)errors.push(`The herd file contains ${detectedHerdTests} current-season herd tests; this build supports up to 10. Please review the file before generating recommendations.`);
  if(hm.herdTestMode!=='dated'&&((hm.scc||[]).length===0))errors.push('No SCC / herd-test columns were detected in the herd file.');
  if(hm.herdTestMode!=='dated'&&manualHerdTests>(hm.scc||[]).length)errors.push(`Current-season herd tests is set to ${manualHerdTests}, but only ${(hm.scc||[]).length} SCC/H/T columns were detected. Correct the herd-test count before generating recommendations.`);
  if(errors.length)throw new Error(errors.join('\n'));
  const seasonYear=+els.seasonYear.value,herdTests=detectedHerdTests||manualHerdTests,seasonWindow=dairySeasonWindow(seasonYear);
  const allMastitisRecords=treatmentRecords(state.mastitis,'mastitis').filter(r=>r.category);
  const mastitisUndated=allMastitisRecords.filter(r=>!r.date).length;
  const mastitisOutOfSeason=allMastitisRecords.filter(r=>r.date&&!inDairySeason(r.date,seasonYear)).length;
  const mastitisRecords=allMastitisRecords.filter(r=>inDairySeason(r.date,seasonYear));
  const mastitisByCow=indexByCow(mastitisRecords),dctRecords=treatmentRecords(state.dct,'dct'),mastaSeason=indexMastaplex(state.mastaplex,seasonYear),mastaIndex=mastaSeason.byCow;const cows=[];
  for(const r of state.herd.rows){const id=clean(r[hm.tag]);if(!id)continue;const scc=(hm.scc||[]).map(h=>num(r[h]));const cow={id,yearBorn:Math.trunc(num(r[hm.yearBorn])??0),calving:parseDate(r[hm.calving]),bcs:num(r[hm.bcs]),preg:clean(r[hm.preg]),expectedCalving:parseDate(r[hm.expectedCalving]),scc,treatments:mastitisByCow.get(id)||[],mastaplex:mastaIndex.get(id)||[]};cow.dct=dctAdvice(scc,cow.preg,cow.treatments,herdTests,ts,la);cow.dryPeriod=dryPeriodStatus(scc,herdTests,ts);Object.assign(cow,dryOffTiming(cow,seasonYear));const missing=[];if(!cow.preg)missing.push('pregnancy');if(cow.bcs==null)missing.push('BCS');if(!/empty/i.test(cow.preg)&&!cow.expectedCalving)missing.push('expected calving');if(cow.scc.slice(0,herdTests).some(v=>v==null||v===0))missing.push('partial SCC');cow.dataNote=missing.join(', ');cows.push(cow)}
  const herdById=new Map(cows.map(c=>[c.id,c])),peak=num(els.peakCows.value)||cows.length,firstCalvers=num(els.firstCalvers.value)||cows.filter(c=>c.yearBorn===seasonYear-2).length,recent=cows.map(c=>c.scc[0]).filter(x=>x!=null&&x>0),sccBuckets={b0_149:0,b150_500:0,over500:0,over1000:0};for(const x of recent){if(x<150)sccBuckets.b0_149++;else if(x<=500)sccBuckets.b150_500++;else sccBuckets.over500++;if(x>1000)sccBuckets.over1000++}
  const dctCounts={};for(const c of cows)dctCounts[c.dct]=(dctCounts[c.dct]||0)+1;const clinical=dctCounts['Mastitis - Clinical']||0,subclinical=dctCounts['Mastitis - Subclinical']||0,sealantOnly=dctCounts['Teat Sealant only']||0,sa=dctCounts['SA DCT/Sealant Combo']||0,laBase=dctCounts['LA DCT/Sealant Combo']||0,laOrder=laBase+clinical+subclinical,sealantCows=sealantOnly+sa+laOrder,sealantTubes=sealantCows*4,saDctTubes=sa*4,laDctTubes=laOrder*4;const pregnant=cows.filter(c=>/pregnant/i.test(c.preg)).length;const pregnantSa=cows.filter(c=>/pregnant/i.test(c.preg)&&c.dct==='SA DCT/Sealant Combo').length;const pregnantLa=cows.filter(c=>/pregnant/i.test(c.preg)&&['LA DCT/Sealant Combo','Mastitis - Clinical','Mastitis - Subclinical'].includes(c.dct)).length;const pregnantSealant=cows.filter(c=>/pregnant/i.test(c.preg)&&c.dct==='Teat Sealant only').length;const emptyHeifers=cows.filter(c=>/empty/i.test(c.preg)&&c.yearBorn===seasonYear-2).length;
  const caseAnalysis=buildCaseAnalysis(mastitisRecords,herdById,seasonYear,peak,firstCalvers);
  const transitionCounts={'Cured':0,'Retained Infection':0,'New Infection':0,'Low SCC':0,'Missed H/T':0,'Not available':0};for(const c of cows)transitionCounts[c.dryPeriod.status]=(transitionCounts[c.dryPeriod.status]||0)+1;const transitionAvailable=cows.length-(transitionCounts['Not available']||0)-(transitionCounts['Missed H/T']||0);
  const prevProducts={},prevCowDct=new Set(),prevCowSeal=new Set(),prevTreatmentByCow=new Map();
  for(const r of dctRecords){
    if(r.drug)prevProducts[r.drug]=(prevProducts[r.drug]||0)+1;
    if(r.category==='DCT')prevCowDct.add(r.id);
    if(r.category==='Sealant')prevCowSeal.add(r.id);
    if(!prevTreatmentByCow.has(r.id))prevTreatmentByCow.set(r.id,{dct:false,sealant:false});
    const p=prevTreatmentByCow.get(r.id);if(r.category==='DCT')p.dct=true;if(r.category==='Sealant')p.sealant=true;
  }
  const prevComboCurrent=new Set(),prevSealantOnlyCurrent=new Set();
  for(const [id,p] of prevTreatmentByCow){
    if(!herdById.has(id))continue;
    if(p.dct&&p.sealant)prevComboCurrent.add(id);
    else if(p.sealant&&!p.dct)prevSealantOnlyCurrent.add(id);
  }
  const currentSeasonMastitisCows=new Set(caseAnalysis.cases.map(c=>c.id));
  const prevComboMastitis=[...prevComboCurrent].filter(id=>currentSeasonMastitisCows.has(id)).length;
  const prevSealantOnlyMastitis=[...prevSealantOnlyCurrent].filter(id=>currentSeasonMastitisCows.has(id)).length;
  const species={},mastaplexGroups={'Gram positive':{},'Gram negative':{},'No bacterial growth':{},'Mixed / contaminated':{},'Other / unclassified':{}};for(const arr of mastaIndex.values())for(const s of arr){species[s]=(species[s]||0)+1;const g=mastaplexGroup(s);mastaplexGroups[g][s]=(mastaplexGroups[g][s]||0)+1;}
  const expectedDates=cows.map(c=>c.expectedCalving).filter(Boolean),calvingStart=expectedDates.length?new Date(Math.min(...expectedDates.map(d=>d.getTime()))):null,midCalving=median(expectedDates);
  const coverage={preg:cows.filter(c=>c.preg).length,bcs:cows.filter(c=>c.bcs!=null).length,expected:cows.filter(c=>c.expectedCalving).length,latestScc:recent.length,transition:transitionAvailable,caseCalving:caseAnalysis.timing.matched,caseTotal:caseAnalysis.cases.length,mastitisRecognised:allMastitisRecords.length,mastitisInSeason:mastitisRecords.length,mastitisOutOfSeason,mastitisUndated,dctLoaded:!!state.dct,mastaplexLoaded:!!state.mastaplex,mastaplexRecognised:mastaSeason.recognised,mastaplexInSeason:mastaSeason.inSeason,mastaplexOutOfSeason:mastaSeason.outOfSeason,mastaplexUndated:mastaSeason.undated,mastaplexDateHeader:mastaSeason.dateHeader,mastaplexFirstParsed:mastaSeason.firstParsed,mastaplexLastParsed:mastaSeason.lastParsed};
  return{cows,peak,firstCalvers,recent,sccBuckets,dctCounts,clinical,subclinical,sealantOnly,sa,laBase,laOrder,sealantCows,sealantTubes,saDctTubes,laDctTubes,pregnant,pregnantSa,pregnantLa,pregnantSealant,emptyHeifers,caseAnalysis,transitionCounts,transitionAvailable,dctRecords,prevProducts,prevCowDct,prevCowSeal,prevComboCurrent,prevSealantOnlyCurrent,prevComboMastitis,prevSealantOnlyMastitis,species,mastaplexGroups,calvingStart,midCalving,coverage,seasonWindow,herdTests,threshold:ts,laThreshold:la};
}

function metricsHtml(items){return`<div class="metric-list">${items.map(([a,b,c])=>`<div class="metric-row"><span class="metric-name">${safe(a)}</span><span class="metric-value"><strong>${safe(b)}</strong>${c?`<small>${safe(c)}</small>`:''}</span></div>`).join('')}</div>`}
function barsHtml(items,total){return items.map(([name,n])=>`<div class="bar-row"><div class="bar-label"><span class="bar-name">${safe(name)}</span><span class="bar-value"><strong>${fmt(n)}</strong>${total?`<small>${pct(n,total)}</small>`:''}</span></div><div class="bar-track"><div class="bar-fill" style="width:${total?Math.min(100,n/total*100):0}%"></div></div></div>`).join('')}
function statusBadge(text,tone='neutral'){return`<span class="status-badge ${tone}">${safe(text)}</span>`}
function coverageBox(label,have,total,note=''){const available=total?have/total:null,tone=available==null?'neutral':available>=.95?'good':available>=.7?'warn':'bad';const value=total?`${fmt(have)} / ${fmt(total)} · ${pct(have,total)}`:safe(have);return`<div class="coverage-item ${tone}"><b>${safe(label)}</b><strong>${value}</strong>${note?`<span>${safe(note)}</span>`:''}</div>`}
function ratioPct(v,d=1){return v==null?'N/A':`${(v*100).toFixed(d)}%`}
function mastitisAxisLabel(v){const n=v*100;return `${n.toFixed(Math.abs(n-Math.round(n))<1e-8?0:1)}%`}
function renderMastitisChart(rows){
  if(!els.mastitisChart)return;
  if(!rows?.length){els.mastitisChart.innerHTML='<p class="na-copy">N/A — no monthly mastitis data.</p>';return}
  const W=1180,H=500,L=72,R=80,T=30,B=72,plotW=W-L-R,plotH=H-T-B;
  const allMonthly=rows.flatMap(x=>[x.heiferRate,x.matureRate,x.rate,x.trigger]).filter(v=>v!=null&&Number.isFinite(v));
  const seasonValues=rows.map(x=>x.seasonPct).filter(v=>v!=null&&Number.isFinite(v));
  const monthlyMax=Math.max(.04,...allMonthly),leftMax=Math.max(.05,Math.ceil(monthlyMax*1.08/.01)*.01);
  const seasonMax=Math.max(.14,...seasonValues),rightMax=Math.max(.14,Math.ceil(seasonMax*1.05/.02)*.02);
  const leftStep=leftMax<=.08?.01:leftMax<=.16?.02:.05,rightStep=rightMax<=.20?.02:.05;
  const yLeft=v=>T+plotH-(Math.max(0,v)/leftMax)*plotH;
  const yRight=v=>T+plotH-(Math.max(0,v)/rightMax)*plotH;
  const groupW=plotW/rows.length,barW=Math.min(17,groupW*.19),gap=4;
  const esc=safe;
  let svg=`<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" aria-labelledby="mastitisChartTitle mastitisChartDesc"><title id="mastitisChartTitle">Clinical mastitis monthly and seasonal trend</title><desc id="mastitisChartDesc">Grouped monthly bars show heifer, mature-age cow and whole-herd clinical mastitis rates. Whole-herd monthly bars turn red when they exceed the SMARTSAMM monthly trigger. The red line is the trigger and the orange line is cumulative season percentage.</desc>`;
  for(let v=0;v<=leftMax+1e-9;v+=leftStep){const y=yLeft(v);svg+=`<line class="chart-grid" x1="${L}" y1="${y}" x2="${W-R}" y2="${y}"/><text class="chart-axis-label left" x="${L-10}" y="${y+4}">${mastitisAxisLabel(v)}</text>`}
  for(let v=0;v<=rightMax+1e-9;v+=rightStep){const y=yRight(v);svg+=`<text class="chart-axis-label right" x="${W-R+10}" y="${y+4}">${mastitisAxisLabel(v)}</text>`}
  svg+=`<line class="chart-axis" x1="${L}" y1="${T}" x2="${L}" y2="${T+plotH}"/><line class="chart-axis" x1="${W-R}" y1="${T}" x2="${W-R}" y2="${T+plotH}"/><line class="chart-axis" x1="${L}" y1="${T+plotH}" x2="${W-R}" y2="${T+plotH}"/><text class="chart-axis-title" x="20" y="${T+plotH/2}" transform="rotate(-90 20 ${T+plotH/2})">Cases per month (%)</text><text class="chart-axis-title" x="${W-18}" y="${T+plotH/2}" transform="rotate(90 ${W-18} ${T+plotH/2})">Accumulated season %</text>`;
  const triggerPts=[],seasonPts=[];
  rows.forEach((x,i)=>{
    const cx=L+groupW*(i+.5),baseY=T+plotH,barStart=cx-(barW*1.5+gap);
    if(x.overTrigger)svg+=`<rect class="trigger-breach-zone" x="${L+groupW*i+2}" y="${T}" width="${Math.max(0,groupW-4)}" height="${plotH}" rx="4"/>`;
    const vals=[['heifer',x.heiferRate],['mature',x.matureRate],['monthly',x.rate]];
    vals.forEach(([cls,v],j)=>{if(v==null)return;const y=yLeft(v),h=Math.max(0,baseY-y),breach=cls==='monthly'&&x.overTrigger?' over-trigger':'';svg+=`<rect class="mastitis-bar ${cls}${breach}" x="${barStart+j*(barW+gap)}" y="${y}" width="${barW}" height="${h}" rx="1"/>`;if(cls==='monthly'&&x.overTrigger){const bx=barStart+j*(barW+gap)+barW/2,by=Math.max(T+10,y-11);svg+=`<circle class="breach-dot" cx="${bx}" cy="${by}" r="8"/><text class="breach-mark" x="${bx}" y="${by+3.5}" text-anchor="middle">!</text>`}});
    triggerPts.push(`${cx},${yLeft(x.trigger)}`);if(x.seasonPct!=null)seasonPts.push(`${cx},${yRight(x.seasonPct)}`);
    svg+=`<text class="chart-month ${x.overTrigger?'over-trigger':''}" x="${cx}" y="${baseY+25}" text-anchor="middle">${esc(x.name.slice(0,3))}</text><rect class="mastitis-hit" data-index="${i}" tabindex="0" x="${L+groupW*i}" y="${T}" width="${groupW}" height="${plotH+38}" fill="transparent"><title>${esc(x.name)} — Heifer ${ratioPct(x.heiferRate)}, MA cows ${ratioPct(x.matureRate)}, Monthly ${ratioPct(x.rate)}, Trigger ${ratioPct(x.trigger)}, Season ${ratioPct(x.seasonPct)}${x.overTrigger?', ABOVE TRIGGER':''}</title></rect>`;
  });
  svg+=`<polyline class="mastitis-line trigger-line" points="${triggerPts.join(' ')}"/><polyline class="mastitis-line season-line" points="${seasonPts.join(' ')}"/>`;
  seasonPts.forEach(pt=>{const [x,y]=pt.split(',');svg+=`<circle class="season-point" cx="${x}" cy="${y}" r="3.5"/>`});
  svg+='</svg>';els.mastitisChart.innerHTML=svg;
  const showMonth=i=>{const x=rows[i];if(!x||!els.mastitisChartTooltip)return;els.mastitisChartTooltip.innerHTML=`<strong>${esc(x.name)}</strong><span>Cases: ${fmt(x.cases)}</span><span>Heifer: ${ratioPct(x.heiferRate)}</span><span>MA cows: ${ratioPct(x.matureRate)}</span><span>Whole herd: ${ratioPct(x.rate)}</span><span>Trigger: ${ratioPct(x.trigger)}</span><span>Season to date: ${ratioPct(x.seasonPct)}</span>`};
  els.mastitisChart.querySelectorAll('.mastitis-hit').forEach(hit=>{const i=Number(hit.dataset.index);hit.addEventListener('mouseenter',()=>showMonth(i));hit.addEventListener('focus',()=>showMonth(i));hit.addEventListener('click',()=>showMonth(i))});
}
function renderMonthly(rows){
  els.monthlyTable.querySelector('tbody').innerHTML=rows.map(x=>{const monthly=x.overTrigger?`<span class="trigger-breach-value"><strong>${pct(x.rate,1)}</strong><small>Above trigger</small></span>`:pct(x.rate,1);return`<tr class="${x.overTrigger?'trigger-row':''}"><td>${x.name}${x.overTrigger?' <span class="table-alert-dot" title="Above SMARTSAMM trigger">!</span>':''}</td><td>${fmt(x.cases)}</td><td>${fmt(x.nsaid)}</td><td>${fmt(x.heifer)}</td><td>${fmt(x.mature)}</td><td>${monthly}</td><td>${pct(x.trigger,1)}</td><td>${pct(x.seasonPct,1)}</td></tr>`}).join('')+`<tr class="total-row"><td>Total</td><td>${fmt(rows.reduce((a,x)=>a+x.cases,0))}</td><td>${fmt(rows.reduce((a,x)=>a+x.nsaid,0))}</td><td>${fmt(rows.reduce((a,x)=>a+x.heifer,0))}</td><td>${fmt(rows.reduce((a,x)=>a+x.mature,0))}</td><td>${pct(rows.reduce((a,x)=>a+x.cases,0),state.result?.peak)}</td><td>14.0%</td><td>${pct(rows.length?rows[rows.length-1].seasonPct:null,1)}</td></tr>`;
  renderMastitisChart(rows);
}
function renderResult(r){
  state.result=r;els.report.classList.remove('hidden');els.errorBox.classList.add('hidden');const farm=clean(els.farmName.value)||'Unnamed farm',vet=clean(els.vetName.value)||'Veterinarian not entered';els.reportTitle.textContent=farm;els.reportMeta.textContent=`${vet} · Season ${els.seasonYear.value}/${String(+els.seasonYear.value+1).slice(-2)} · Generated ${new Date().toLocaleDateString('en-NZ')}`;
  const herdMap=state.herd?.map||{},herdDates=herdMap.sccDates||[],herdTests=r.herdTests||+els.herdTests.value;
  const currentDates=herdDates.slice(0,herdTests);
  const previousSource=herdMap.previousHerdTest||herdDates[herdTests]||herdMap.previousHerdTestHeader||herdMap.scc?.[herdTests]||'';
  const firstCurrentSource=(currentDates.length?currentDates[currentDates.length-1]:herdMap.scc?.[Math.max(0,herdTests-1)])||'';
  const comparisonSource=previousSource&&firstCurrentSource?`${previousSource} → ${firstCurrentSource}`:(previousSource||'N/A — previous/pre-dry H/T not detected');
  const currentTestDisplay=currentDates.length?currentDates.join(', '):(herdMap.scc?.slice(0,herdTests).join(', ')||String(herdTests));
  const enteredStart=parseDate(els.calvingStart.value),expectedDry=parseDate(els.expectedDryOff.value);
  const facts=[['Veterinarian',vet],['Analysis season',r.seasonWindow.label],['Dairy company',clean(els.dairyCompany.value)||'N/A'],['Supply number',clean(els.supplyNumber.value)||'N/A'],['LIC / MyHerd',clean(els.ptptCode.value).toUpperCase()||'N/A'],['Heifers teat sealed?',clean(els.heifersSealed.value)||'N/A'],['Previous-season BMSCC',clean(els.bmsccPrevious.value)||'N/A'],['Current-season BMSCC',clean(els.bmsccCurrent.value)||'N/A'],['Current herd tests',currentTestDisplay],['Dry-period SCC comparison source',comparisonSource],['Peak cows',fmt(r.peak)],['First calvers',fmt(r.firstCalvers)],['SCC cut-offs',`${r.threshold} / ${r.laThreshold}`],['Plan start of calving',isoDate(enteredStart||r.calvingStart)||'N/A'],['Expected mid-calving',isoDate(r.midCalving)||'N/A'],['Expected dry-off',isoDate(expectedDry)||'N/A']];els.consultFacts.innerHTML=facts.map(([a,b])=>`<div class="fact"><b>${safe(a)}</b><span>${safe(b)}</span></div>`).join('');

  els.dataQuality.innerHTML=[
    coverageBox('Pregnancy status',r.coverage.preg,r.cows.length,'Used for DCT/carryover classification'),
    coverageBox('BCS',r.coverage.bcs,r.cows.length,'Used for BCS-adjusted dry-off timing'),
    coverageBox('Expected calving date',r.coverage.expected,r.cows.length,'Required for individual dry-off timing'),
    coverageBox('Latest SCC',r.coverage.latestScc,r.cows.length,'Cows present at the latest selected H/T'),
    coverageBox('Dry-period SCC comparison',r.coverage.transition,r.cows.length,'Requires pre-dry + first post-calving SCC'),
    coverageBox('Case-to-calving match',r.coverage.caseCalving,r.coverage.caseTotal,'Unmatched cases remain in season totals'),
    coverageBox('Mastitis rows in selected season',r.coverage.mastitisInSeason,r.coverage.mastitisRecognised,`${fmt(r.coverage.mastitisOutOfSeason)} outside ${r.seasonWindow.label}; ${fmt(r.coverage.mastitisUndated)} undated excluded`),
    coverageBox('Previous DCT / ITS',r.coverage.dctLoaded?'Loaded':'Not supplied',null,'Optional'),
    r.coverage.mastaplexLoaded?coverageBox('Mastaplex results in selected season',r.coverage.mastaplexInSeason,r.coverage.mastaplexRecognised,`${fmt(r.coverage.mastaplexOutOfSeason)} outside ${r.seasonWindow.label}; ${fmt(r.coverage.mastaplexUndated)} undated excluded · date column: ${r.coverage.mastaplexDateHeader||'not detected'}${r.coverage.mastaplexFirstParsed?` · parsed range ${isoDate(r.coverage.mastaplexFirstParsed)}–${isoDate(r.coverage.mastaplexLastParsed)}`:''}`):coverageBox('Mastaplex','Not supplied',null,'Optional')
  ].join('');

  const seasonCases=r.caseAnalysis.cases.length,over500=r.sccBuckets.over500,monthsOverRows=r.caseAnalysis.monthly.filter(x=>x.overTrigger),monthsOver=monthsOverRows.length,monthsOverNames=monthsOverRows.map(x=>x.name),early30=r.caseAnalysis.timing.d0_7+r.caseAnalysis.timing.d8_30;
  els.kpis.innerHTML=[['Clinical case events',fmt(seasonCases)],['Season case %',pct(seasonCases,r.peak)],['Latest SCC >500',pct(over500,r.recent.length)],['Repeat cows',fmt(r.caseAnalysis.repeatCows)],['Retained post-DCT',fmt(r.transitionCounts['Retained Infection'])]].map(([label,value])=>`<div class="kpi"><div class="value">${safe(value)}</div><div class="label">${safe(label)}</div></div>`).join('');

  const summary=[];
  summary.push(`There were ${fmt(seasonCases)} clinical mastitis case events across ${fmt(r.caseAnalysis.uniqueCows)} cows (${pct(seasonCases,r.peak)} of peak cow numbers by case count).`);
  if(r.coverage.mastitisOutOfSeason||r.coverage.mastitisUndated)summary.push(`Treatment analysis was restricted to ${r.seasonWindow.label}. ${fmt(r.coverage.mastitisOutOfSeason)} recognised mastitis treatment row${r.coverage.mastitisOutOfSeason===1?' was':'s were'} outside the selected season and ${fmt(r.coverage.mastitisUndated)} row${r.coverage.mastitisUndated===1?' was':'s were'} undated; these were excluded from case calculations and cow DCT treatment-history overrides.`);
  if(monthsOver){const monthList=monthsOverNames.length===1?monthsOverNames[0]:monthsOverNames.length===2?`${monthsOverNames[0]} and ${monthsOverNames[1]}`:`${monthsOverNames.slice(0,-1).join(', ')} and ${monthsOverNames[monthsOverNames.length-1]}`;summary.push(`${fmt(monthsOver)} month${monthsOver===1?'':'s'} exceeded the monthly SMARTSAMM trigger: ${monthList}.`);}else summary.push('No months exceeded the monthly SMARTSAMM trigger.');
  if(r.caseAnalysis.timing.matched)summary.push(`${fmt(early30)} of ${fmt(r.caseAnalysis.timing.matched)} cases with a usable calving-date match occurred within 30 days after calving; ${fmt(r.caseAnalysis.timing.unavailable)} case${r.caseAnalysis.timing.unavailable===1?' was':'s were'} unavailable for timing analysis.`);else summary.push('Case timing from calving is unavailable because no treatment cases could be matched to a usable season calving date.');
  if(r.transitionAvailable)summary.push(`Dry-period SCC outcome was calculable for ${fmt(r.transitionAvailable)} cows: ${fmt(r.transitionCounts['Cured'])} cured, ${fmt(r.transitionCounts['Retained Infection'])} retained infections, ${fmt(r.transitionCounts['New Infection'])} new infections and ${fmt(r.transitionCounts['Low SCC'])} remained low.`);else summary.push('Dry-period SCC cure/new/retained status could not be calculated from the supplied herd-test history.');
  if(r.coverage.preg < r.cows.length*0.95)summary.push(`Pregnancy status is available for ${fmt(r.coverage.preg)} of ${fmt(r.cows.length)} cows. The DCT order should be treated as provisional because empty cows cannot be reliably excluded where pregnancy status is missing.`);
  summary.push(`Current DCT tube estimate: ${fmt(r.saDctTubes)} SA DCT tubes, ${fmt(r.laDctTubes)} LA DCT tubes (including clinical/subclinical cows), and ${fmt(r.sealantTubes)} teat sealant tubes across all cows receiving sealant.`);
  els.executiveSummary.innerHTML=summary.map(x=>`<div class="summary-point"><span>•</span><p>${safe(x)}</p></div>`).join('');

  renderMonthly(r.caseAnalysis.monthly);
  const t=r.caseAnalysis.timing;els.caseTiming.innerHTML=metricsHtml([['0–7 days',fmt(t.d0_7),`${pct(t.d0_7,t.matched)} of matched cases`],['8–30 days',fmt(t.d8_30),`${pct(t.d8_30,t.matched)} of matched cases`],['Over 30 days',fmt(t.over30),`${pct(t.over30,t.matched)} of matched cases`],['Pre-calving',fmt(t.pre),`${pct(t.pre,t.matched)} of matched cases`],['Timing unavailable',fmt(t.unavailable),'Missing/unusable calving-date match'],['Usable timing coverage',`${fmt(t.matched)} / ${fmt(r.caseAnalysis.cases.length)}`,pct(t.matched,r.caseAnalysis.cases.length)]]);
  els.caseSummary.innerHTML=metricsHtml([['Cows with mastitis case events',fmt(r.caseAnalysis.uniqueCows)],['Repeat cows',fmt(r.caseAnalysis.repeatCows),`${pct(r.caseAnalysis.repeatCows,r.peak)} of peak cows`],['Repeat cow + same quarter',fmt(r.caseAnalysis.repeatQuarterCows)],['Ab + NSAID combination cases',fmt(r.caseAnalysis.combo),`${pct(r.caseAnalysis.combo,r.caseAnalysis.cases.length)} of clinical case events`],['Extended antibiotic cases (>3 doses)',fmt(r.caseAnalysis.extended)],['Treatment change within same case',fmt(r.caseAnalysis.txChanges)],['Mastitis cases not matched to herd file',fmt(r.caseAnalysis.unmatchedCowCases)]]);
  const products=Object.entries(r.caseAnalysis.productCounts).sort((a,b)=>b[1]-a[1]);els.drugSummary.innerHTML=products.length?amrBarsHtml(products,r.caseAnalysis.antibioticRows.length)+`<div class="amr-legend"><span>${amrBadgeHtml({category:'green'})} first-line</span><span>${amrBadgeHtml({category:'yellow'})} restricted</span><span>${amrBadgeHtml({category:'red'})} critically important</span></div><p class="amr-note">NZVA traffic-light category by recognised active ingredient. Combination products use the highest recognised ingredient category. Unknown products remain Unclassified. ${safe(AMR_CATALOGUE_VERSION)}.</p>`+metricsHtml([['Antibiotic treatment rows',fmt(r.caseAnalysis.antibioticRows.length)],['NSAID treatment rows',fmt(r.caseAnalysis.nsaidRows.length)],['NSAID-paired case events',fmt(r.caseAnalysis.combo)]]):'<p class="na-copy">N/A — no antibiotic treatment products detected.</p>';
  const q=Object.entries(r.caseAnalysis.quarters).filter(([,n])=>n>0);els.quarterSummary.innerHTML=q.length?barsHtml(q,r.caseAnalysis.cases.length)+`<p class="context-note">Percentages use clinical case events as the denominator. A multi-quarter case is counted once under Multi.</p>`:'<p class="na-copy">N/A — no quarter information detected.</p>';
  els.sccSummary.innerHTML=r.recent.length?barsHtml([['0–150',r.sccBuckets.b0_149],['150–500',r.sccBuckets.b150_500],['Over 500',r.sccBuckets.over500],['Over 1,000',r.sccBuckets.over1000]],r.recent.length)+metricsHtml([['Cows with latest SCC',fmt(r.recent.length)]])+`<p class="context-note">For threshold handling, an SCC of exactly 150 is counted in the 150–500 group. Over 1,000 is a subset of Over 500. SCC values of 0 are treated as unavailable rather than a true SCC result.</p>`:'<p class="na-copy">N/A — no usable latest herd-test SCC values.</p>';
  const transitionItems=[['Cured',r.transitionCounts['Cured']],['Retained infection',r.transitionCounts['Retained Infection']],['New infection',r.transitionCounts['New Infection']],['Remained low',r.transitionCounts['Low SCC']]];els.sccTransitions.innerHTML=r.transitionAvailable?barsHtml(transitionItems,r.transitionAvailable)+metricsHtml([['Usable comparisons',`${fmt(r.transitionAvailable)} / ${fmt(r.cows.length)}`],['Missed first post-calving H/T',fmt(r.transitionCounts['Missed H/T'])],['Pre-dry comparison unavailable',fmt(r.transitionCounts['Not available'])],['Threshold used',`${fmt(r.threshold)} ×1000`]]):'<p class="na-copy">N/A — pre-dry and first post-calving SCC values were not both available.</p>';
  const pp=Object.entries(r.prevProducts).sort((a,b)=>b[1]-a[1]);
  if(state.dct){
    if(pp.length){
      const comboDen=r.prevComboCurrent.size,sealDen=r.prevSealantOnlyCurrent.size;
      const comboOutcome=comboDen?`${fmt(r.prevComboMastitis)} / ${fmt(comboDen)}`:'N/A';
      const sealOutcome=sealDen?`${fmt(r.prevSealantOnlyMastitis)} / ${fmt(sealDen)}`:'N/A';
      const comboOutcomeNote=comboDen?`${pct(r.prevComboMastitis,comboDen)} · ≥1 current-season case`:'No matched previous combo cows';
      const sealOutcomeNote=sealDen?`${pct(r.prevSealantOnlyMastitis,sealDen)} · ≥1 current-season case`:'No matched previous sealant-only cows';
      const selectiveNote=sealDen
        ? 'Rates use cows matched to the current herd file. Mastitis means at least one current-season mastitis case event; this is an outcome comparison, not proof that the previous dry-off treatment caused or prevented a case.'
        : 'No previous sealant-only cows were detected among cows matched to the current herd. A selective-treatment comparison is therefore not available for this season.';
      els.previousDctSummary.innerHTML=`<div class="previous-dct-layout"><div><h4>Previous dry-off products</h4>${amrBarsHtml(pp,Math.max(1,r.dctRecords.length))}<p class="amr-note">Antibiotic products use the NZVA AMR traffic-light category; combination products use the highest recognised ingredient category, and teat sealants are marked Non-antibiotic. Unknown products remain Unclassified.</p>${metricsHtml([['Cows with DCT',fmt(r.prevCowDct.size)],['Cows with sealant',fmt(r.prevCowSeal.size)]])}</div><div class="order-box selective-dct-box"><h4>Current-season mastitis by previous treatment</h4>${metricsHtml([['Previous DCT + sealant combo cows',fmt(comboDen),'Matched to current herd'],['DCT combo → mastitis this season',comboOutcome,comboOutcomeNote],['Previous sealant-only cows',fmt(sealDen),'Matched to current herd'],['Sealant only → mastitis this season',sealOutcome,sealOutcomeNote]])}<p class="context-note">${safe(selectiveNote)}</p></div></div>`;
    }else els.previousDctSummary.innerHTML='<p class="na-copy">Previous DCT / ITS file loaded, but no recognised DCT or sealant products were detected.</p>';
  }else els.previousDctSummary.innerHTML='<p class="na-copy">N/A — previous DCT / ITS data not supplied.</p>';
  const sp=Object.entries(r.species).sort((a,b)=>b[1]-a[1]);
  if(state.mastaplex){
    if(sp.length){
      const gp=Object.entries(r.mastaplexGroups['Gram positive']||{}).sort((a,b)=>b[1]-a[1]);
      const gn=Object.entries(r.mastaplexGroups['Gram negative']||{}).sort((a,b)=>b[1]-a[1]);
      const ng=Object.entries(r.mastaplexGroups['No bacterial growth']||{}).sort((a,b)=>b[1]-a[1]);
      const mix=Object.entries(r.mastaplexGroups['Mixed / contaminated']||{}).sort((a,b)=>b[1]-a[1]);
      const other=Object.entries(r.mastaplexGroups['Other / unclassified']||{}).sort((a,b)=>b[1]-a[1]);
      const gpN=gp.reduce((a,x)=>a+x[1],0),gnN=gn.reduce((a,x)=>a+x[1],0),bacterialN=gpN+gnN,total=sp.reduce((a,x)=>a+x[1],0);
      const groupMetrics=metricsHtml([['Gram-positive growths',fmt(gpN),bacterialN?pct(gpN,bacterialN)+' of classified bacterial growths':'No classified bacterial growths'],['Gram-negative growths',fmt(gnN),bacterialN?pct(gnN,bacterialN)+' of classified bacterial growths':'No classified bacterial growths']]);
      const groupBlock=(title,items,cls)=>`<section class="mastaplex-group ${cls}"><h4>${safe(title)}</h4>${items.length?barsHtml(items,Math.max(1,total)):'<p class="na-copy">None detected.</p>'}</section>`;
      const otherItems=[...ng,...mix,...other].sort((a,b)=>b[1]-a[1]);
      els.mastaplexSummary.innerHTML=`<div class="mastaplex-overview">${groupMetrics}</div><div class="mastaplex-group-grid">${groupBlock('Gram-positive organisms',gp,'gram-positive')}${groupBlock('Gram-negative organisms',gn,'gram-negative')}</div>${otherItems.length?`<div class="mastaplex-other">${groupBlock('Other Mastaplex results',otherItems,'other-results')}</div>`:''}<p class="context-note">Only Mastaplex results dated within ${safe(r.seasonWindow.label)} are included. Gram percentages compare classified bacterial growths only. No-growth, mixed/contaminated and unclassified results are shown separately. ${fmt(r.coverage.mastaplexOutOfSeason)} result${r.coverage.mastaplexOutOfSeason===1?' was':'s were'} outside the selected season and ${fmt(r.coverage.mastaplexUndated)} result${r.coverage.mastaplexUndated===1?' was':'s were'} undated; these were excluded.</p>`;
    }else els.mastaplexSummary.innerHTML=`<p class="na-copy">Mastaplex file loaded, but no cultured-growth results fell within ${safe(r.seasonWindow.label)}.</p><p class="context-note">Date column used: <strong>${safe(r.coverage.mastaplexDateHeader||'not detected')}</strong>. ${r.coverage.mastaplexFirstParsed?`Parsed date range: ${safe(isoDate(r.coverage.mastaplexFirstParsed))}–${safe(isoDate(r.coverage.mastaplexLastParsed))}. `:''}${fmt(r.coverage.mastaplexOutOfSeason)} result${r.coverage.mastaplexOutOfSeason===1?' was':'s were'} outside the selected season and ${fmt(r.coverage.mastaplexUndated)} result${r.coverage.mastaplexUndated===1?' was':'s were'} undated; these were excluded.</p>`;
  }else els.mastaplexSummary.innerHTML='<p class="na-copy">N/A — Mastaplex data not supplied.</p>';

  const order=[['Teat Sealant only',r.sealantOnly],['SA DCT/Sealant Combo',r.sa],['LA DCT/Sealant Combo',r.laBase],['Mastitis - Clinical',r.clinical],['Mastitis - Subclinical',r.subclinical],['Empty',r.dctCounts.Empty||0],['Absent from H/T',r.dctCounts['Absent from H/T (Culled/Carryover)']||0]];els.dctSummary.innerHTML=barsHtml(order,r.cows.length)+`<p class="amr-note">SA / LA recommendation bands are not AMR-coloured because the prescribed DCT product has not yet been selected. Recorded previous DCT products above are colour-coded where recognised.</p>`;
  els.dctOrder.innerHTML=`<div class="order-box"><h4>Order estimate</h4>${metricsHtml([['SA DCT combo cows',fmt(r.sa)],['LA DCT combo cows',fmt(r.laOrder),'Includes clinical/subclinical cows'],['Sealant-only cows',fmt(r.sealantOnly)],['Total sealant cows',fmt(r.sealantCows)],['Total sealant tubes',fmt(r.sealantTubes)],['Pregnant cows',fmt(r.pregnant)],['Pregnant SA combo',fmt(r.pregnantSa)],['Pregnant LA combo',fmt(r.pregnantLa)],['Pregnant sealant only',fmt(r.pregnantSealant)],['Empty heifers / carryover candidates',fmt(r.emptyHeifers)],['Pregnancy-data coverage',`${fmt(r.coverage.preg)} / ${fmt(r.cows.length)}`,r.coverage.preg<r.cows.length?'Order provisional where pregnancy is missing':'Complete']])}</div>`;
  renderCowTable(r.cows);els.report.scrollIntoView({behavior:'smooth',block:'start'});
}
function renderCowTable(cows){
  const q=clean(els.cowSearch.value).toLowerCase(),filtered=cows.filter(c=>!q||[c.id,c.dct,c.advice,c.dryPeriod.status,c.dryPeriod.pre,c.dryPeriod.post,c.preg,c.dataNote,...c.mastaplex].join(' ').toLowerCase().includes(q));els.cowCountText.textContent=`Showing ${filtered.length.toLocaleString()} of ${cows.length.toLocaleString()} cows`;
  const threshold=state.result?.threshold;
  els.cowTable.querySelector('tbody').innerHTML=filtered.map(c=>{
    const tone=c.dryPeriod.status==='Cured'||c.dryPeriod.status==='Low SCC'?'good':c.dryPeriod.status==='Retained Infection'||c.dryPeriod.status==='New Infection'?'warn':'neutral';
    const hasPair=c.dryPeriod.pre!=null&&c.dryPeriod.pre!==0&&c.dryPeriod.post!=null&&c.dryPeriod.post!==0;
    const comparison=hasPair?`<div class="scc-status-explain"><strong>${safe(c.dryPeriod.pre)} → ${safe(c.dryPeriod.post)}</strong><span>cutoff ${safe(threshold)}</span></div>`:'';
    const clinicalMastitis=c.dct==='Mastitis - Clinical',subclinicalMastitis=c.dct==='Mastitis - Subclinical';
    const mastitisFlag=clinicalMastitis?'<span class="mastitis-history-flag clinical">CLINICAL MASTITIS THIS SEASON</span>':subclinicalMastitis?'<span class="mastitis-history-flag subclinical">SUBCLINICAL MASTITIS THIS SEASON</span>':'';
    const dctDisplay=(clinicalMastitis||subclinicalMastitis)?`${mastitisFlag}<div class="dct-reason"><strong>LA DCT + sealant</strong><span>Reason: mastitis history</span></div>`:safe(c.dct);
    return`<tr><td><span class="tag">${safe(c.id)}</span></td><td>${c.scc[0]??''}</td><td>${c.dryPeriod.pre??''}</td><td>${c.dryPeriod.post??''}</td><td>${safe(c.preg||'N/A')}</td><td>${dctDisplay}</td><td>${safe(c.advice)}</td><td>${statusBadge(c.dryPeriod.status,tone)}${comparison}</td><td>${safe(c.dataNote||'—')}</td><td>${safe(c.mastaplex.join(', '))}</td></tr>`
  }).join('');
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
function resetForNewFarm(){
  const confirmed=window.confirm('Start a new report? Current farm details, uploaded files and generated results will be cleared.');
  if(!confirmed)return;

  state.herd=null;state.mastitis=null;state.dct=null;state.mastaplex=null;state.result=null;

  // Reset consult inputs to a clean new-farm state while retaining the agreed defaults.
  const clearIds=['farmName','vetName','peakCows','firstCalvers','dairyCompany','supplyNumber','ptptCode','bmsccPrevious','bmsccCurrent','expectedDryOff'];
  for(const id of clearIds)if(els[id])els[id].value='';
  if(els.heifersSealed)els.heifersSealed.value='';
  if(els.seasonYear)els.seasonYear.value=String(new Date().getFullYear()-1);
  if(els.herdTests){els.herdTests.disabled=false;els.herdTests.value='4';delete els.herdTests.dataset.userSet;}
  if(els.sccTs)els.sccTs.value='150';
  if(els.sccLa)els.sccLa.value='250';
  defaultCalvingStartFromSeason();

  // Clear the browser file inputs as well as the in-memory data so the same
  // filename can be selected again for the next client if necessary.
  for(const id of ['herdFile','mastitisFile','dctFile','mastaplexFile'])if(els[id])els[id].value='';
  for(const kind of ['herd','mastitis','dct','mastaplex']){
    const el=els[`${kind}Status`];
    if(el){el.className='file-status neutral';el.textContent='No file loaded';}
  }

  if(els.mappingPanel){els.mappingPanel.innerHTML='';els.mappingPanel.classList.add('hidden');}
  if(els.cowSearch)els.cowSearch.value='';
  if(els.errorBox){els.errorBox.textContent='';els.errorBox.classList.add('hidden');}
  if(els.report)els.report.classList.add('hidden');
  if(els.viewReportBtn)els.viewReportBtn.classList.add('hidden');
  if(els.generateBtn){els.generateBtn.disabled=false;els.generateBtn.textContent='Generate report';}
  if(els.generateStatus)els.generateStatus.textContent='';

  syncHerdTestSafety();
  updatePrescriptionFields();
  window.scrollTo({top:0,behavior:'smooth'});
}
els.newFarmBtn?.addEventListener('click',resetForNewFarm);
function exportCowRecommendations(){
  if(!state.result)return;
  const rows=[['Cow Tag','Latest SCC','Pre-dry SCC','First comparison SCC','SCC low/high cutoff','Pregnancy Diagnosis','Expected Calving Date','BCS','Individual DCT Recommendation','Dry-off Advice','Dry-period SCC Status','Dry-period SCC comparison','Data Note','Mastaplex Result']];
  for(const c of state.result.cows)rows.push([c.id,c.scc[0]??'',c.dryPeriod.pre??'',c.dryPeriod.post??'',state.result.threshold??'',c.preg,isoDate(c.expectedCalving),c.bcs??'',c.dct,c.advice,c.dryPeriod.status,(c.dryPeriod.pre&&c.dryPeriod.post)?`${c.dryPeriod.pre} -> ${c.dryPeriod.post}`:'',c.dataNote,c.mastaplex.join('; ')]);
  const csv='\uFEFF'+rows.map(r=>r.map(x=>`"${String(x??'').replace(/"/g,'""')}"`).join(',')).join('\r\n');
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));
  a.download=`${(clean(els.farmName.value)||'farm').replace(/[^a-z0-9]+/gi,'-')}-individual-cow-recommendations.csv`;
  document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(a.href);
}
// The report-header button owns the export action.
// The lower button calls this control directly from its inline click so it also works
// during the first load after a PWA update when an older cached app.js may still be active.
els.csvBtn.addEventListener('click',exportCowRecommendations);

function demoData(){
  els.farmName.value='Demo Dairy';els.vetName.value='Dr Example';els.seasonYear.value=2025;els.peakCows.value=20;els.firstCalvers.value=5;els.sccTs.value=150;els.sccLa.value=250;els.herdTests.value=4;els.dairyCompany.value='Fonterra Co-operative Group';els.supplyNumber.value='12345';els.ptptCode.value='ABCD';els.heifersSealed.value='Y';els.bmsccPrevious.value=145;els.bmsccCurrent.value=128;els.calvingStart.value='2026-08-01';els.expectedDryOff.value='2026-05-20';updatePrescriptionFields();
  const herd=Array.from({length:20},(_,i)=>({'Tag':String(101+i),'Animal-Year Born':i<5?2023:2021,'Calving date':new Date(2025,6,(i%20)+1),'BCS':[4.2,4.5,4.7,5,5.2][i%5],'Most recent H/T':[82,116,188,275,520,73,146,205,340,110,620,95,155,230,410,90,180,260,780,130][i],'2nd':[95,120,170,260,480,80,160,190,315,135,590,120,140,250,390,100,160,240,650,145][i],'3rd':[110,130,160,240,450,88,180,175,290,150,550,140,130,270,360,120,145,220,580,160][i],'4th':[120,140,175,220,430,95,190,165,280,160,510,145,120,260,330,125,140,210,560,170][i],'5th':[130,155,140,200,400,110,150,155,250,170,470,160,110,240,300,140,130,190,520,180][i],'Pregnancy Diagnosis':i===18?'Empty':'Pregnant','Expected Calving date':new Date(2026,6,(i%20)+1)}));
  const mastitis=[{'Cow Number':'103','Condition':'Mastitis - Clinical','Treatment Date':new Date(2025,7,12),'Quarter':'LF','Product':'Intracillin 1000 Milking Cow','Days':3},{'Cow Number':'103','Condition':'Mastitis - Clinical','Treatment Date':new Date(2025,7,12),'Quarter':'','Product':'Metacam 20mg/Ml For Injection','Days':1},{'Cow Number':'103','Condition':'Mastitis - Clinical','Treatment Date':new Date(2025,9,4),'Quarter':'LF','Product':'Intracillin 1000 Milking Cow','Days':3},{'Cow Number':'107','Condition':'Mastitis - Clinical','Treatment Date':new Date(2025,8,7),'Quarter':'RR','Product':'Intracillin 1000 Milking Cow','Days':5},{'Cow Number':'111','Condition':'Mastitis - Subclinical','Treatment Date':new Date(2025,10,10),'Quarter':'RF','Product':'Albiotic','Days':3},{'Cow Number':'111','Condition':'Mastitis - Subclinical','Treatment Date':new Date(2025,10,10),'Quarter':'','Product':'Metacam 20mg/Ml For Injection','Days':1}];
  const dct=[{'Cow Number':'101','Treatment Date':new Date(2025,4,20),'Product':'Dryzen'},{'Cow Number':'101','Treatment Date':new Date(2025,4,20),'Product':'Duramast DC 500'},{'Cow Number':'102','Treatment Date':new Date(2025,4,20),'Product':'Dryzen'},{'Cow Number':'102','Treatment Date':new Date(2025,4,20),'Product':'Quadrant DC'},{'Cow Number':'103','Treatment Date':new Date(2025,4,20),'Product':'Dryzen'},{'Cow Number':'107','Treatment Date':new Date(2025,4,20),'Product':'Dryzen'},{'Cow Number':'107','Treatment Date':new Date(2025,4,20),'Product':'Quadrant DC'}];const mp=[{'Cow Number':'103','Species':'Staph. aureus','Quarter':'LF','Sample Date':new Date(2025,7,13)},{'Cow Number':'107','Species':'Strep. uberis','Quarter':'RR','Sample Date':new Date(2025,8,8)},{'Cow Number':'111','Species':'CNS','Quarter':'RF','Sample Date':new Date(2025,10,11)},{'Cow Number':'102','Species':'E. coli','Quarter':'LR','Sample Date':new Date(2024,9,1)}];
  state.herd={rows:herd,map:detectHerd(herd),name:'Demo herd data'};state.mastitis={rows:mastitis,map:detectTreatment(mastitis),name:'Demo mastitis treatments'};state.dct={rows:dct,map:detectTreatment(dct),name:'Demo DCT / ITS'};state.mastaplex={rows:mp,map:detectMastaplex(mp),name:'Demo Mastaplex'};for(const k of ['herd','mastitis','dct','mastaplex'])setStatus(k,state[k].rows,state[k].map,state[k].name);try{renderResult(buildResult());els.generateBtn.textContent='Report generated ✓';els.generateStatus.textContent='Demo report loaded below.';els.viewReportBtn.classList.remove('hidden')}catch(e){els.errorBox.textContent=`Demo report could not be generated: ${e.message}`;els.errorBox.classList.remove('hidden');els.report.classList.add('hidden')}
}
els.demoBtn.addEventListener('click',demoData);
if('serviceWorker'in navigator&&location.protocol.startsWith('http'))navigator.serviceWorker.register('./sw.js').catch(()=>{});
