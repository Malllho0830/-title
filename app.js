const $=id=>document.getElementById(id);
const fmt=x=>Number(x).toLocaleString('ko-KR',{maximumFractionDigits:3});
let worker, nextId=0, requests=new Map(), sectors=[], rowChoices=[], cells=[], basket=[], result=null, linkage=[], currentLevel='basic';
const YEARS=['2024','2023','2022'];
let currentYear=YEARS[0], demand=null, comparing=false;
const LEVELS={
 basic:{label:'기본부문',files:y=>['io-'+y+'-basic.xlsx'].concat(y==='2023'?['io-2023.xlsx']:[])},
 small:{label:'소분류',files:y=>['io-'+y+'-small.xlsx']},
 mid:{label:'중분류',files:y=>['io-'+y+'-mid.xlsx']},
 large:{label:'대분류',files:y=>['io-'+y+'-large.xlsx']}
};
function text(tag,value){const e=document.createElement(tag);e.textContent=value;return e;}
function call(type,payload){return new Promise((resolve,reject)=>{const id=++nextId;requests.set(id,{resolve,reject});worker.postMessage({id,type,payload});});}
function startWorker(){worker=new Worker('./worker.js');worker.onmessage=({data:d})=>{if(d.type==='status'){$('status').textContent=d.text;return;}const p=requests.get(d.id);if(!p)return;requests.delete(d.id);d.ok?p.resolve(d.result):p.reject(Error(d.error));};worker.onerror=e=>{for(const p of requests.values())p.reject(Error(e.message||'Worker 실행 오류'));requests.clear();};}
function options(el,list){el.replaceChildren(...list.map(x=>{const o=document.createElement('option');o.value=x.code;o.textContent=x.code+' '+x.name;return o;}));}
function invalidate(){result=null;$('results').hidden=true;}
async function load(buffer,level,year){
 invalidate();basket=[];renderBasket();$('compare-results').hidden=true;$('compare-status').textContent='시뮬레이션에서 시나리오를 추가한 뒤 비교하세요.';
 for(const id of ['sheet','row','sector','add','run','row-search','col-search','sector-search','amount','link-sort','link-search','dem-sector','dem-sector-search','compare-run'])$(id).disabled=true;
 if(worker)worker.terminate();for(const p of requests.values())p.reject(Error('데이터 다시 로드'));requests.clear();startWorker();
 try{const info=await call('init',buffer);sectors=info.sectors;linkage=info.linkage||[];currentLevel=level;currentYear=year;options($('sector'),sectors);setLevelTexts();$('link-search').value='';renderLinkage();demand=info.demand||null;$('dem-sector-search').value='';options($('dem-sector'),sectors);renderDemand();$('sector-search').value='';$('row-search').value='';$('col-search').value='';$('sheet').replaceChildren(...info.sheets.map(s=>{const o=text('option',s);o.value=s;return o;}));
 $('validation').textContent=JSON.stringify(info.residuals,null,2)+'\n최대 절대 잔차: 허용값 1e-8 미만';
 for(const id of ['sheet','row','sector','add','run','row-search','col-search','sector-search','amount','link-sort','link-search','dem-sector','dem-sector-search','compare-run'])$(id).disabled=false;
 await explore();$('status').textContent='준비 완료 · '+year+'년 '+LEVELS[level].label+' · '+sectors.length+'개 부문 · '+info.sheets.length+'개 시트 · 계수 검증 통과';
 }catch(e){$('status').textContent='준비 실패: '+e.message+' — 새로고침 후 재시도하세요.';}
}
let exploreSeq=0;
async function explore(){const seq=++exploreSeq;try{const d=await call('explore',{sheet:$('sheet').value,row:$('row').value});if(seq!==exploreSeq)return;rowChoices=d.rows;filterRows(d.selected);cells=d.cells;$('unit').textContent=d.unit;$('explore-caption').textContent=$('sheet').value+' / 행: '+d.selected+' '+d.name;renderCells();}catch(e){$('status').textContent=e.message;}}
function filterRows(selected){const q=$('row-search').value.trim();options($('row'),rowChoices.filter(r=>(r.code+' '+r.name).includes(q)));if([...$('row').options].some(o=>o.value===selected))$('row').value=selected;}
function renderCells(){const q=$('col-search').value.trim();$('cells').replaceChildren(...cells.filter(c=>(c.code+' '+c.name).includes(q)).map(c=>{const tr=document.createElement('tr');for(const v of [c.code,c.name,c.value===null?'빈 셀':typeof c.value==='number'?fmt(c.value):c.value])tr.append(text('td',v));return tr;}));}
$('sheet').onchange=()=>{$('row-search').value='';$('row').replaceChildren();explore();};$('row').onchange=explore;
$('row-search').oninput=()=>{filterRows();if($('row').value)explore();else{$('cells').replaceChildren();$('explore-caption').textContent='일치하는 행이 없습니다.';}};
$('col-search').oninput=renderCells;
$('sector-search').oninput=()=>{const q=$('sector-search').value.trim();options($('sector'),sectors.filter(s=>(s.code+' '+s.name).includes(q)));};
function renderBasket(){$('basket').replaceChildren(...basket.map((s,i)=>{const li=text('li',s.code+' '+s.name+' / '+fmt(s.amount)+'백만원 ');const b=text('button','삭제');b.type='button';b.onclick=()=>{basket.splice(i,1);invalidate();renderBasket();};li.append(b);return li;}));$('sim-status').textContent=basket.length?basket.length+'개 입력 준비. 계산 버튼을 누르세요.':'시나리오를 추가하세요.';}
$('scenario').onsubmit=e=>{e.preventDefault();const amount=Number($('amount').value),s=sectors.find(s=>s.code===$('sector').value);if(!s||!Number.isFinite(amount)||amount<=0){$('sim-status').textContent='산업과 양의 증가액을 확인하세요.';return;}basket.push({...s,amount});invalidate();renderBasket();};
$('clear').onclick=()=>{basket=[];invalidate();renderBasket();};
$('run').onclick=async()=>{if(!basket.length){$('sim-status').textContent='시나리오를 먼저 추가하세요.';return;}const snapshot=basket.map(s=>({...s}));$('run').disabled=true;$('sim-status').textContent='계산 중…';try{const r=await call('simulate',snapshot);if(JSON.stringify(snapshot)!==JSON.stringify(basket)){$('sim-status').textContent='입력이 변경되었습니다. 다시 계산하세요.';return;}result=r;renderResult();$('sim-status').textContent='계산 완료. 아래 결과는 현재 시나리오에 대응합니다.';}catch(e){invalidate();$('sim-status').textContent='계산 실패: '+e.message;}finally{$('run').disabled=false;}};
function renderResult(){$('results').hidden=false;$('cards').replaceChildren(...[['수요 증가액',result.demand],['생산유발액',result.production],['수입유발액',result.imports],['부가가치유발액',result.valueAdded]].map(([title,v])=>{const d=text('div',title);d.className='card';d.append(text('strong',fmt(v)),text('span','백만원'));return d;}));const max=Math.max(...result.rows.map(r=>r.production),1);$('bars').replaceChildren(...result.rows.slice(0,15).map(r=>{const d=text('div',r.name+' · '+fmt(r.production)+'백만원');d.className='bar';const t=document.createElement('div');t.className='track';const f=document.createElement('div');f.className='fill';f.style.width=Math.max(0,r.production/max*100)+'%';t.append(f);d.append(t);return d;}));$('result-rows').replaceChildren(...result.rows.map(r=>{const tr=document.createElement('tr');[r.code,r.name,fmt(r.production),fmt(r.imports),fmt(r.valueAdded)].forEach(v=>tr.append(text('td',v)));return tr;}));}
function csvCell(v){let s=String(v);if(/^[=+@-]/.test(s)&&typeof v!=='number')s="'"+s;return '"'+s.replaceAll('"','""')+'"';}
$('download').onclick=()=>{if(!result)return;const rows=[['산업코드','산업명','생산유발액_백만원','수입유발액_백만원','부가가치유발액_백만원'],...result.rows.map(r=>[r.code,r.name,r.production,r.imports,r.valueAdded])];const url=URL.createObjectURL(new Blob(['\ufeff'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download='io-'+currentYear+'-'+currentLevel+'-results.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};

function setLevelTexts(){
 const L=LEVELS[currentLevel].label,n=sectors.length;
 $('hdr-sub').textContent=currentYear+' · 생산자가격 · '+L;
 $('hdr-desc').textContent='12개 원자료 시트 탐색, '+n+'개 부문 수요충격·연쇄효과 분석, 최종수요 항목별 분해, 연도 간 비교';
 $('download').textContent='전체 '+n+'개 산업 CSV 다운로드';
 $('result-caption').textContent=currentYear+'년 '+L+' 산업별 유발액 — 백만원, 생산유발액 내림차순';
 $('link-caption').textContent=currentYear+'년 '+L+' 산업 연쇄효과 ('+n+'개 부문) — 무단위 지수, 전 산업 평균 = 1';
}
function renderLinkage(){
 const q=$('link-search').value.trim(),k=$('link-sort').value;
 const rows=linkage.filter(r=>(r.code+' '+r.name).includes(q)).sort((a,b)=>b[k]-a[k]);
 $('link-rows').replaceChildren(...rows.map(r=>{const tr=document.createElement('tr');[r.code,r.name,r.backward.toFixed(3),r.forward.toFixed(3),r.type].forEach(v=>tr.append(text('td',v)));return tr;}));
 const cnt={};linkage.forEach(r=>cnt[r.type]=(cnt[r.type]||0)+1);
 $('link-summary').textContent=linkage.length?Object.entries(cnt).map(([t,v])=>t+' '+v+'개').join(' · '):'자료를 불러온 후 표시됩니다.';
}
$('link-sort').onchange=renderLinkage;$('link-search').oninput=renderLinkage;


// ---- 4. 최종수요 항목별 유발 분해 ----
const GROUPS=['소비','투자','수출'];
function demandRows(v){
 const it=demand.items,ti=it.findIndex(x=>x.code==='9190');
 const sum=(k,g)=>it.reduce((a,x,i)=>x.group===g?a+v[k][i]:a,0);
 const rows=[];
 for(const g of GROUPS){
  rows.push({label:g,bold:true,p:sum('production',g),m:sum('imports',g),v:sum('valueAdded',g)});
  it.forEach((x,i)=>{if(x.group===g&&x.name!==g)rows.push({label:'　'+x.name,p:v.production[i],m:v.imports[i],v:v.valueAdded[i]});});
 }
 rows.push({label:'최종수요계',bold:true,p:v.production[ti],m:v.imports[ti],v:v.valueAdded[ti]});
 return rows;
}
function fillDemandTable(tbody,rows){
 const tot=rows[rows.length-1].p;
 tbody.replaceChildren(...rows.map(r=>{
  const tr=document.createElement('tr'),td=document.createElement('td');
  if(r.bold)td.append(text('strong',r.label));else td.textContent=r.label;
  tr.append(td);
  [fmt(r.p),tot?(r.p/tot*100).toFixed(1):'-',fmt(r.m),fmt(r.v)].forEach(x=>tr.append(text('td',x)));
  return tr;}));
}
function renderDemand(){
 const L=LEVELS[currentLevel].label,y=currentYear;
 if(!demand){
  $('dem-rows').replaceChildren();$('dem-bars').replaceChildren();$('dem-sector-rows').replaceChildren();
  $('dem-caption').textContent='이 파일에서는 최종수요 항목별 유발액 시트를 읽지 못했습니다.';return;
 }
 const rows=demandRows(demand.total),tot=rows[rows.length-1].p;
 $('dem-caption').textContent=y+'년 '+L+' 전 산업 합계 — 백만원, 구성비는 생산유발액 기준(%)';
 fillDemandTable($('dem-rows'),rows);
 $('dem-bars').replaceChildren(...rows.filter(r=>r.bold&&GROUPS.includes(r.label)).map(r=>{
  const share=tot?r.p/tot*100:0,d=text('div',r.label+' · '+share.toFixed(1)+'% ('+fmt(r.p)+'백만원)');
  d.className='bar';const t=document.createElement('div');t.className='track';
  const f=document.createElement('div');f.className='fill';f.style.width=Math.max(0,share)+'%';
  t.append(f);d.append(t);return d;}));
 renderDemandSector();
}
function renderDemandSector(){
 const code=$('dem-sector').value,v=demand&&demand.sectors[code],s=sectors.find(x=>x.code===code);
 if(!v||!s){$('dem-sector-rows').replaceChildren();$('dem-sector-caption').textContent='산업을 선택하세요.';return;}
 $('dem-sector-caption').textContent=currentYear+'년 '+LEVELS[currentLevel].label+' · '+s.code+' '+s.name+' — 이 산업의 생산을 유발한 최종수요 (백만원)';
 fillDemandTable($('dem-sector-rows'),demandRows(v));
}
$('dem-sector').onchange=renderDemandSector;
$('dem-sector-search').oninput=()=>{const q=$('dem-sector-search').value.trim();options($('dem-sector'),sectors.filter(x=>(x.code+' '+x.name).includes(q)));renderDemandSector();};

// ---- 5. 연도 간 비교 ----
async function fetchBuffer(level,year){
 const files=LEVELS[level].files(year);
 for(const f of files){const r=await fetch('./data/'+f);if(r.ok)return r.arrayBuffer();}
 throw Error('data/'+files[0]+' 파일을 찾을 수 없습니다.');
}
function runTemp(buffer,shocks,year){
 return new Promise((resolve,reject)=>{
  const w=new Worker('./worker.js'),pend=new Map();let id=0;
  const ask=(type,payload)=>new Promise((res,rej)=>{const i=++id;pend.set(i,{res,rej});w.postMessage({id:i,type,payload});});
  w.onmessage=({data:d})=>{if(d.type==='status')return;const p=pend.get(d.id);if(!p)return;pend.delete(d.id);d.ok?p.res(d.result):p.rej(Error(d.error));};
  w.onerror=e=>{for(const p of pend.values())p.rej(Error(e.message||'Worker 실행 오류'));pend.clear();};
  (async()=>{
   try{
    const info=await ask('init',buffer);
    const miss=shocks.find(s=>!info.sectors.some(x=>x.code===s.code));
    if(miss)throw Error(year+'년 표에 '+miss.code+' '+miss.name+' 부문이 없습니다.');
    resolve(await ask('simulate',shocks));
   }catch(e){reject(e);}finally{w.terminate();}
  })();
 });
}
function renderCompare(out,level){
 const L=LEVELS[level].label,last=out[out.length-1],first=out[0];
 $('cmp-sum-caption').textContent=L+' · 같은 증가액을 연도별 산업연관표로 계산한 결과 (백만원, 계수는 무단위)';
 $('cmp-sum-rows').replaceChildren(...out.map(o=>{const r=o.r,tr=document.createElement('tr');
  [o.year+'년',fmt(r.demand),fmt(r.production),fmt(r.imports),fmt(r.valueAdded),(r.production/r.demand).toFixed(4),(r.imports/r.demand).toFixed(4),(r.valueAdded/r.demand).toFixed(4)].forEach(v=>tr.append(text('td',v)));return tr;}));
 const max=Math.max(...out.map(o=>o.r.production),1);
 $('cmp-bars').replaceChildren(...out.map(o=>{
  const d=text('div',o.year+'년 · 생산유발액 '+fmt(o.r.production)+'백만원 (생산유발계수 '+(o.r.production/o.r.demand).toFixed(4)+')');
  d.className='bar';const t=document.createElement('div');t.className='track';
  const f=document.createElement('div');f.className='fill';f.style.width=Math.max(0,o.r.production/max*100)+'%';
  t.append(f);d.append(t);return d;}));
 $('cmp-ind-caption').textContent=L+' · '+last.year+'년 기준 생산유발액 상위 15개 산업의 연도별 비교 (백만원)';
 const hr=document.createElement('tr');
 ['코드','산업',...out.map(o=>o.year+'년'),first.year+'→'+last.year+' 변화율'].forEach(h=>hr.append(text('th',h)));
 $('cmp-ind-head').replaceChildren(hr);
 const maps=out.map(o=>new Map(o.r.rows.map(x=>[x.code,x.production])));
 $('cmp-ind-rows').replaceChildren(...last.r.rows.slice(0,15).map(row=>{
  const tr=document.createElement('tr'),vals=maps.map(m=>m.get(row.code));
  const a=vals[0],b=vals[vals.length-1];
  [row.code,row.name,...vals.map(v=>v===undefined?'-':fmt(v)),a>0&&b!==undefined?((b-a)/a*100).toFixed(1)+'%':'-'].forEach(v=>tr.append(text('td',v)));
  return tr;}));
 $('compare-results').hidden=false;
}
$('compare-run').onclick=async()=>{
 if(comparing)return;
 if(!basket.length){$('compare-status').textContent='위의 시뮬레이션에서 시나리오를 먼저 추가하세요.';return;}
 const level=currentLevel,shocks=basket.map(s=>({code:s.code,name:s.name,amount:s.amount})),years=YEARS.slice().reverse(),btn=$('compare-run'),out=[];
 comparing=true;btn.disabled=true;$('compare-results').hidden=true;
 try{
  for(let i=0;i<years.length;i++){
   $('compare-status').textContent=years[i]+'년 '+LEVELS[level].label+' 계산 중 ('+(i+1)+'/'+years.length+') — 연도마다 파일을 불러오므로 시간이 걸립니다.';
   const r=await runTemp(await fetchBuffer(level,years[i]),shocks,years[i]);
   out.push({year:years[i],r});
  }
  renderCompare(out,level);
  $('compare-status').textContent='비교 완료 · '+LEVELS[level].label+' · 시나리오 '+shocks.length+'개 산업, 합계 '+fmt(out[0].r.demand)+'백만원';
 }catch(e){$('compare-status').textContent='비교 실패: '+e.message;}
 finally{comparing=false;btn.disabled=$('sheet').disabled;}
};

async function loadWorkbook(level,year){
 const files=LEVELS[level].files(year),label=year+'년 '+LEVELS[level].label;
 try {
  $('status').textContent=label+' Excel 내려받는 중…';
  let response=null;
  for(const f of files){
   const r=await fetch('./data/'+f);
   if(r.ok){response=r;break;}
  }
  if(!response)throw Error('data/'+files[0]+' 파일을 찾을 수 없습니다. data 폴더에 올렸는지, 파일 이름이 같은지 확인하세요.');
  await load(await response.arrayBuffer(),level,year);
 }catch(error){
  $('status').textContent='준비 실패: '+error.message+' — 연결 상태를 확인하고 새로고침하세요.';
 }
}
$('year').replaceChildren(...YEARS.map(y=>{const o=text('option',y+'년');o.value=y;return o;}));
const syncButton=()=>{const b=$('load-btn');b.disabled=false;b.textContent=$('year').value+'년 '+LEVELS[$('level').value].label+' 불러오기';};
$('level').onchange=syncButton;$('year').onchange=syncButton;syncButton();
$('load-btn').onclick=async()=>{
 const button=$('load-btn'),level=$('level').value,year=$('year').value,name=year+'년 '+LEVELS[level].label;
 button.disabled=true;
 button.textContent='불러오는 중…';
 try {
  await loadWorkbook(level,year);
  const ready=!$('sheet').disabled&&currentLevel===level&&currentYear===year;
  button.textContent=ready?name+' 불러오기 완료':name+' 다시 불러오기';
  button.disabled=ready;
 }finally{
  if($('sheet').disabled){button.disabled=false;button.textContent=name+' 다시 불러오기';}
 }
};
