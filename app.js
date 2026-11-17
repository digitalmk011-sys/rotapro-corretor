'use strict';
let originals=[], processed=[], headers=[], sourceWorkbook=null, sourceSheetName=null, sourceFilename="", sourceRowIndexes=[], originalCount=0;
const $=id=>document.getElementById(id);
const clean=s=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[.,;#]/g,' ').replace(/\s+/g,' ').trim();
const pretty=s=>String(s||'').replace(/\b\w/g,c=>c.toUpperCase());
const STREET_TYPES={r:'rua',rua:'rua',av:'avenida',aven:'avenida',avenida:'avenida',al:'alameda',alam:'alameda',alameda:'alameda',trav:'travessa',tv:'travessa',travessa:'travessa',pr:'praca',pca:'praca',praca:'praca',rod:'rodovia',rodovia:'rodovia',estr:'estrada',estrada:'estrada'};
const WORD_EXPANSIONS={pion:'pioneiro',pioneir:'pioneiro',prof:'professor',profa:'professora',dr:'doutor',dra:'doutora',cel:'coronel',gen:'general',dep:'deputado',pres:'presidente',sto:'santo',sta:'santa',pe:'padre',jd:'jardim'};
const WORD_ACCENTS={jose:'José',joao:'João',antonio:'Antônio',sao:'São',vitor:'Vítor',conceicao:'Conceição',goncalves:'Gonçalves',marcia:'Márcia',luiz:'Luiz',luis:'Luís',sebastiao:'Sebastião',tres:'Três',praca:'Praça'};
function normalizeStreet(s){
 let words=clean(s).split(/\s+/).filter(Boolean);
 if(!words.length)return '';
 if(STREET_TYPES[words[0]])words[0]=STREET_TYPES[words[0]];
 words=words.map(w=>WORD_EXPANSIONS[w]||w);
 return words.join(' ');
}
function displayStreet(s){return s.split(' ').map((w,i)=>i>0&&['da','de','do','das','dos','e'].includes(w)?w:(WORD_ACCENTS[w]||w.charAt(0).toUpperCase()+w.slice(1))).join(' ')}
function parseAddress(raw){
 const text=String(raw??'').trim();
 // Remove primeiro os complementos, inclusive quando aparecem antes do numero do predio.
 // Numeros precedidos por apartamento, apto, bloco, unidade etc. nunca viram numero da rua.
 const normalized=text.replace(/[;|]/g,',').replace(/\s+/g,' ').trim();
 const tokens=normalized.split(/(?=\b(?:apto?\.?|apartamento|apart\.?|ap\.?|bloco|bl\.?|torre|unidade|und\.?|sala|conjunto|cj\.?|fundos|lote|quadra)\b)/i);
 const unitPattern=/(?:^|[\s,])(?:apto?\.?|apartamento|apart\.?|ap\.?|bloco|bl\.?|torre|unidade|und\.?|sala|conjunto|cj\.?|fundos|lote|quadra)\s*(?:n[ºo°.]?\s*)?([a-z]?\d{1,6}[a-z]?)/gi;
 const units=[];let match;
 while((match=unitPattern.exec(normalized))!==null)units.push({start:match.index,end:unitPattern.lastIndex,number:match[1]});
 // Um complemento introduzido por palavra-chave pode estar no meio do endereco.
 let base=normalized.replace(/(?:^|[\s,])(?:apto?\.?|apartamento|apart\.?|ap\.?|bloco|bl\.?|torre|unidade|und\.?|sala|conjunto|cj\.?|fundos|lote|quadra)\s*(?:n[ºo°.]?\s*)?[a-z]?\d{1,6}[a-z]?(?=$|[\s,])/gi,' ');
 base=base.replace(/\s*,\s*,/g,',').trim().replace(/,$/,'');
 // Complementos sem numero, no fim, nao sao parte do logradouro.
 base=base.replace(/\s*,?\s*(?:apartamento|apto|ap|bloco|torre|fundos)\s*$/i,'');
 const candidates=[...base.matchAll(/(?:^|[\s,])(?:n[ºo°.]?\s*)?(\d{1,6}[a-z]?)(?=$|[\s,])/gi)];
 let street='',num='',ambiguousNumber=false;
 if(candidates.length){
  // A primeira referencia numerica restante e o numero predial; se houver varias, pedir revisao.
  const m=candidates[0],prefix=base.slice(0,m.index).replace(/[\s,]+$/,'');
  street=normalizeStreet(prefix);num=m[1].toUpperCase();
  ambiguousNumber=candidates.length>1;
 }else street=normalizeStreet(base.replace(/\s*,.*$/,''));
 return {street,num,address:displayStreet(street)+(num?', '+num:''),ambiguousNumber,unitNumbers:units.map(u=>u.number)};
}
function wordsMatch(shortName,longName){
 const a=shortName.split(' '),b=longName.split(' ');
 if(a.length!==b.length)return false;
 return a.every((w,i)=>w===b[i]||(w.length===1&&b[i].startsWith(w))||(w.length>=3&&b[i].startsWith(w)&&b[i].length>w.length));
}
function chooseCanonical(processed){
 // Uma forma longa só é adotada se for única e corroborada pela própria planilha.
 const buckets=new Map();
 for(const x of processed){if(!x.p.num||!x.p.street)continue;const k=x.city+'|'+x.p.num;if(!buckets.has(k))buckets.set(k,[]);buckets.get(k).push(x)}
 for(const bucket of buckets.values()){
  const names=[...new Set(bucket.map(x=>x.p.street))];
  for(const x of bucket){
   const candidates=names.filter(n=>wordsMatch(x.p.street,n)&&n.length>=x.p.street.length);
   const max=Math.max(0,...candidates.map(n=>n.length));const best=candidates.filter(n=>n.length===max);
   if(best.length===1&&best[0]!==x.p.street){x.p.street=best[0];x.p.address=displayStreet(best[0])+', '+x.p.num;x.expanded=true}
   else if(best.length>1){x.ambiguous=true}
  }
 }
}
function field(row,names){const keys=Object.keys(row);const k=keys.find(k=>names.some(n=>clean(k)===clean(n)));return k?row[k]:''}
function process(data){
 originals=data;headers=[...new Set(data.flatMap(r=>Object.keys(r)))];
 processed=data.map((r,i)=>{
 const raw=field(r,['Destination Address','Endereço','Endereco','Address','Logradouro']);
 const city=clean(field(r,['City','Cidade','Município','Municipio']));
 const cep=String(field(r,['Zipcode','Zip Code','CEP','Postal code'])||'').replace(/\D/g,'');
 const lat=Number(String(field(r,['Latitude','Lat'])||'').replace(',','.'));
 const lon=Number(String(field(r,['Longitude','Long','Lng'])||'').replace(',','.'));
 const coord=Number.isFinite(lat)&&Number.isFinite(lon)&&lat!==0&&lon!==0&&Math.abs(lat)<=90&&Math.abs(lon)<=180;
 const p=parseAddress(raw);const key=city+'|'+p.street+'|'+p.num;
 return {r,raw,city,cep,lat,lon,coord,p,key,i};
 });
 chooseCanonical(processed);
 for(const x of processed)x.key=x.city+'|'+x.p.street+'|'+x.p.num;
 const groups=new Map();for(const x of processed){if(!groups.has(x.key))groups.set(x.key,[]);groups.get(x.key).push(x)}
 let n=0;for(const [,group] of groups){n++;const cp=new Set(group.filter(x=>x.cep.length===8).map(x=>x.cep));const c=group.filter(x=>x.coord);let divergent=false;
 for(let i=0;i<c.length;i++)for(let j=i+1;j<c.length;j++)if(haversine(c[i],c[j])>150)divergent=true;
 for(const x of group){x.stop=n;x.review=!x.p.num||!x.city||x.p.ambiguousNumber||cp.size>1||divergent||x.ambiguous;x.reason=!x.p.num?'Número não identificado':!x.city?'Cidade ausente':x.p.ambiguousNumber?'Mais de um número predial possível — revisar':cp.size>1?'CEPs diferentes no mesmo endereço':divergent?'Coordenadas distantes para o mesmo endereço':x.ambiguous?'Nome de rua ambíguo — revisar':x.expanded?'Abreviação expandida com referência na planilha':'Padronizado por rua, número e cidade';}}
 render();$('export').disabled=!processed.length; $('exportNotice').hidden=true;
}
function haversine(a,b){const r=Math.PI/180,dlat=(a.lat-b.lat)*r,dlon=(a.lon-b.lon)*r;const q=Math.sin(dlat/2)**2+Math.cos(a.lat*r)*Math.cos(b.lat*r)*Math.sin(dlon/2)**2;return 12742000*Math.asin(Math.min(1,Math.sqrt(q)))}
function escapeHtml(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function render(){const unique=new Set(processed.map(x=>x.key)).size;$('count').textContent=`${processed.length} encomendas · ${unique} endereços físicos · ${processed.filter(x=>x.review).length} para revisar · ${processed.filter(x=>x.p.num&&x.p.address!==x.raw).length} padronizados`;
 $('rows').innerHTML=processed.map(x=>{const q=encodeURIComponent(x.p.address+(x.city?', '+x.city:'')+', Brasil');const google='https://www.google.com/maps/search/?api=1&query='+q;const waze='https://waze.com/ul?q='+q+'&navigate=no';return `<tr><td>${escapeHtml(x.raw)}</td><td>${escapeHtml(x.p.address)}<br><small>${escapeHtml(x.city)}</small></td><td>${x.stop}</td><td class="${x.review?'warn':'ok'}">${x.review?'⚠ Revisar':'✓ Agrupado'}<br><small>${escapeHtml(x.reason)}</small></td><td><a href="${google}" target="_blank" rel="noopener">Google</a> · <a href="${waze}" target="_blank" rel="noopener">Waze</a></td></tr>`}).join('')}
// Importa pela matriz da planilha para preservar a posição de cada linha, inclusive vazias.
$('file').addEventListener('change',async e=>{
 const f=e.target.files[0];if(!f)return;
 try{
  if(typeof XLSX==='undefined')throw Error('Biblioteca XLSX indisponível. Verifique a internet.');
  const wb=XLSX.read(await f.arrayBuffer(),{type:'array',cellStyles:true,cellDates:true});
  const name=wb.SheetNames[0], sheet=wb.Sheets[name];
  const matrix=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:true,blankrows:true});
  const headerIdx=matrix.findIndex(row=>row.some(v=>['destination address','endereco','endereço','address','logradouro'].includes(clean(v))));
  if(headerIdx<0)throw Error('Coluna Destination Address / Endereço não encontrada.');
  const h=matrix[headerIdx].map((v,i)=>String(v||'Coluna '+(i+1)));
  const addrIdx=h.findIndex(v=>['destination address','endereco','endereço','address','logradouro'].includes(clean(v)));
  const data=[],indexes=[];
  for(let i=headerIdx+1;i<matrix.length;i++){
   const vals=matrix[i]||[];
   if(!vals.some(v=>v!==''&&v!==null&&v!==undefined))continue;
   const row={};h.forEach((key,j)=>row[key]=vals[j]??'');
   // Inclui toda linha da planilha, mesmo que o endereço esteja vazio.
   data.push(row);indexes.push(i);
  }
  if(!data.length)throw Error('A planilha não contém entregas.');
  sourceWorkbook=wb;sourceSheetName=name;sourceFilename=f.name;
  sourceRowIndexes=indexes;originalCount=data.length;
  process(data);
  $('status').textContent='Planilha importada: '+data.length+' linhas. Exportação manterá todas as linhas, colunas e abas.';
 }catch(err){$('status').textContent='Erro ao importar: '+err.message;sourceWorkbook=null;$('export').disabled=true}
});
$('demo').onclick=()=>{sourceWorkbook=null;sourceSheetName=null;sourceFilename='exemplo.xlsx';sourceRowIndexes=[];originalCount=2;process([{'Destination Address':'R Pion Alfredo J da Costa 472, Ap 405',City:'Maringá',Sequence:1},{'Destination Address':'Rua Pioneiro Alfredo Jose da Costa, 472, Apto 908',City:'Maringá',Sequence:2}]);$('status').textContent='EXEMPLO com 2 entregas. Para exportar sua rota completa, importe o arquivo XLSX.'};
$('export').onclick=()=>{
 if(typeof XLSX==='undefined'){alert('Biblioteca XLSX indisponível.');return}
 if(!processed.length)return;
 try{
  let wb, exportedCount=processed.length;
  if(sourceWorkbook&&sourceSheetName){
   wb=sourceWorkbook;
   const ws=wb.Sheets[sourceSheetName];
   const range=XLSX.utils.decode_range(ws['!ref']);
   let headerRow=-1,addressCol=-1;
   for(let row=range.s.r;row<=Math.min(range.e.r,range.s.r+30);row++){
    for(let col=range.s.c;col<=range.e.c;col++){
     const cell=ws[XLSX.utils.encode_cell({r:row,c:col})];
     if(cell&&['destination address','endereco','endereço','address','logradouro'].includes(clean(cell.v))){headerRow=row;addressCol=col;break}
    }
    if(addressCol>=0)break;
   }
   if(addressCol<0)throw Error('Não encontrei a coluna de endereço.');
   if(sourceRowIndexes.length!==processed.length)throw Error('Contagem de linhas inconsistente.');
   processed.forEach((item,i)=>{
    if(!item.p.num||!item.p.street||item.ambiguous||item.p.ambiguousNumber)return; // Endereço incerto fica intacto.
    const row=sourceRowIndexes[i];
    const ref=XLSX.utils.encode_cell({r:row,c:addressCol});
    const old=ws[ref]||{};
    ws[ref]={...old,t:'s',v:item.p.address};delete ws[ref].w;
   });
  }else{
   wb=XLSX.utils.book_new();
   const rows=processed.map(x=>({...x.r,'Destination Address':x.p.num&&!x.ambiguous&&!x.p.ambiguousNumber?x.p.address:x.raw}));
   XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows,{header:[...headers]}),'Entregas corrigidas');
  }
  const filename=sourceFilename.replace(/\.xlsx?$/i,'')+'_CORRIGIDA.xlsx';
  const output=XLSX.write(wb,{bookType:'xlsx',type:'array',cellStyles:true});
  // Baixa o arquivo completo, não uma lista filtrada de correções.
  const blob=new Blob([output],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url=URL.createObjectURL(blob);const a=document.createElement('a');
  a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),60000);
  $('exportNotice').hidden=false;
  $('exportNotice').textContent='✓ PLANILHA CORRIGIDA — '+exportedCount+' de '+originalCount+' linhas exportadas em XLSX: '+filename;
  $('status').textContent='Exportação da planilha completa realizada. Nenhuma entrega foi removida.';
 }catch(err){alert('Falha ao exportar a planilha completa: '+err.message)}
};
