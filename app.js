'use strict';
let originals=[], processed=[], headers=[];
const $=id=>document.getElementById(id);
const clean=s=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[.,;#]/g,' ').replace(/\s+/g,' ').trim();
const pretty=s=>String(s||'').replace(/\b\w/g,c=>c.toUpperCase());
function normalizeStreet(s){
 let x=clean(s).replace(/^(?:r|r ua|rua)\s+/,'rua ').replace(/^(?:av|aven|avenida)\s+/,'avenida ')
 .replace(/\bpion\b/g,'pioneiro').replace(/\bpio\b/g,'pioneiro').replace(/\bprof\b/g,'professor').replace(/\bdr\b/g,'doutor');
 // Expandir iniciais somente em padrões conhecidos e conferíveis; nunca em todas as ruas.
 x=x.replace(/\balfredo j da costa\b/g,'alfredo jose da costa');
 return x;
}
function parseAddress(raw){
 const text=String(raw??'').trim();
 let base=text.split(/\s*,\s*(?=(?:ap(?:to|artamento)?|bloco|torre|casa|fundos|sala|conj|quadra|lote|unidade)\b)/i)[0];
 base=base.replace(/\s+(?:ap(?:to|artamento)?|bloco|torre|casa|fundos|sala|conj|unidade)\s*[\w-]+.*$/i,'');
 let m=base.match(/^(.*?)[,\s]+(\d+[a-zA-Z]?)(?:\s*[-,].*)?$/);
 if(!m) return {street:normalizeStreet(base),num:'',address:pretty(normalizeStreet(base))};
 const street=normalizeStreet(m[1]);return {street,num:m[2].toUpperCase(),address:pretty(street)+', '+m[2].toUpperCase()};
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
 const groups=new Map();for(const x of processed){if(!groups.has(x.key))groups.set(x.key,[]);groups.get(x.key).push(x)}
 let n=0;for(const [,group] of groups){n++;const cp=new Set(group.filter(x=>x.cep.length===8).map(x=>x.cep));const c=group.filter(x=>x.coord);let divergent=false;
 for(let i=0;i<c.length;i++)for(let j=i+1;j<c.length;j++)if(haversine(c[i],c[j])>150)divergent=true;
 for(const x of group){x.stop=n;x.review=!x.p.num||!x.city||cp.size>1||divergent;x.reason=!x.p.num?'Número não identificado':!x.city?'Cidade ausente':cp.size>1?'CEPs diferentes no mesmo endereço':divergent?'Coordenadas distantes para o mesmo endereço':'Agrupado por rua, número e cidade';}}
 render();$('export').disabled=!processed.length;
}
function haversine(a,b){const r=Math.PI/180,dlat=(a.lat-b.lat)*r,dlon=(a.lon-b.lon)*r;const q=Math.sin(dlat/2)**2+Math.cos(a.lat*r)*Math.cos(b.lat*r)*Math.sin(dlon/2)**2;return 12742000*Math.asin(Math.min(1,Math.sqrt(q)))}
function escapeHtml(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function render(){const unique=new Set(processed.map(x=>x.key)).size;$('count').textContent=`${processed.length} encomendas · ${unique} endereços físicos · ${processed.filter(x=>x.review).length} para revisar`;
 $('rows').innerHTML=processed.map(x=>{const q=encodeURIComponent(x.p.address+(x.city?', '+x.city:'')+', Brasil');const google='https://www.google.com/maps/search/?api=1&query='+q;const waze='https://waze.com/ul?q='+q+'&navigate=no';return `<tr><td>${escapeHtml(x.raw)}</td><td>${escapeHtml(x.p.address)}<br><small>${escapeHtml(x.city)}</small></td><td>${x.stop}</td><td class="${x.review?'warn':'ok'}">${x.review?'⚠ Revisar':'✓ Agrupado'}<br><small>${escapeHtml(x.reason)}</small></td><td><a href="${google}" target="_blank" rel="noopener">Google</a> · <a href="${waze}" target="_blank" rel="noopener">Waze</a></td></tr>`}).join('')}
$('file').addEventListener('change',async e=>{const f=e.target.files[0];if(!f)return;try{if(typeof XLSX==='undefined')throw Error('Biblioteca XLSX indisponível. Verifique a conexão com a internet.');const wb=XLSX.read(await f.arrayBuffer(),{type:'array'});const sheet=wb.Sheets[wb.SheetNames[0]];process(XLSX.utils.sheet_to_json(sheet,{defval:''}));$('status').textContent='Planilha carregada: '+f.name}catch(err){$('status').textContent='Erro: '+err.message}});
$('demo').onclick=()=>process([{'Destination Address':'R Pion Alfredo J da Costa 472, Ap 405',City:'Maringá',Sequence:1},{'Destination Address':'Rua Pioneiro Alfredo Jose da Costa, 472, Apto 908',City:'Maringá',Sequence:2}]);
$('export').onclick=()=>{if(typeof XLSX==='undefined'){alert('Biblioteca XLSX indisponível.');return}const rows=processed.map(x=>({...x.r,'Endereço padronizado':x.p.address,'Chave agrupamento':x.key,'Parada corrigida':x.stop,'Status correção':x.review?'REVISAR':'AGRUPADO','Motivo':x.reason}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows,{header:[...headers,'Endereço padronizado','Chave agrupamento','Parada corrigida','Status correção','Motivo']}),'Entregas corrigidas');XLSX.writeFile(wb,'RotaPro_Enderecos_Corrigidos.xlsx')};
