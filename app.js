const $ = id => document.getElementById(id);
let state = { sales: [], stock: [], order: [], collections: new Map(), selected: new Set() };

const THEME = {
  navy: '1F406A', navy2: '173554', pale: 'EAF2F8', green: 'D9EAD3',
  yellow: 'FFF2CC', orange: 'FCE5CD', red: 'F4CCCC', white: 'FFFFFF', gray: 'F3F4F6', text: '1F2937'
};
const money = n => Number(n || 0).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
const number = n => Number(n || 0).toLocaleString('pt-BR',{maximumFractionDigits:2});
const norm = s => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toUpperCase();
const numBR = v => {
  if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
  let s = String(v ?? '').trim().replace(/R\$/gi,'').replace(/\s/g,'');
  if (!s) return 0;
  if (s.includes(',') && s.includes('.')) s = s.replace(/\./g,'').replace(',','.');
  else if (s.includes(',')) s = s.replace(',','.');
  const n = Number(s); return Number.isFinite(n) ? n : 0;
};
const excelDate = v => {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number' && XLSX.SSF?.parse_date_code) {
    const d = XLSX.SSF.parse_date_code(v); if (d) return new Date(d.y,d.m-1,d.d);
  }
  const s=String(v||'').trim(); if(!s) return null;
  const br=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if(br) return new Date(+br[3],+br[2]-1,+br[1]);
  const iso=s.match(/^(\d{4})-(\d{2})-(\d{2})/); if(iso) return new Date(+iso[1],+iso[2]-1,+iso[3]);
  const d=new Date(s); return isNaN(d)?null:d;
};
function referenceFromCode(code){ const m=String(code??'').trim().match(/^(\d{6})/); return m?m[1]:''; }
function collectionFromReference(ref){
  ref=String(ref||''); if(!/^\d{3}/.test(ref)) return {key:'NA',name:'Coleção não identificada'};
  const family=ref[0], year=ref.slice(1,3); const names={'1':'Inverno','3':'Verão','4':'Alto Verão'};
  return names[family]?{key:ref.slice(0,3),name:`${names[family]} ${year}`}:{key:'NA',name:'Coleção não identificada'};
}
function parseProduct(text){
  const raw=String(text||''); const name=raw.split(/\s+Cor:/i)[0].trim();
  const cor=(raw.match(/Cor:([^;]+)/i)||[])[1]?.trim()||''; const tamanho=(raw.match(/Tamanho:([^;]+)/i)||[])[1]?.trim()||'';
  const first=norm(name).split(/\s+/)[0]||'OUTROS';
  const catMap={BLAZER:'BLAZER',VESTIDO:'VESTIDO',CALCA:'CALÇA',BLUSA:'BLUSA',SAIA:'SAIA',CONJUNTO:'CONJUNTO',SHORT:'SHORT',SHORTS:'SHORT',JAQUETA:'CASACO/JAQUETA',CASACO:'CASACO/JAQUETA',MACACAO:'MACACÃO',BODY:'BODY',CAMISA:'CAMISA',CAMISETA:'CAMISA',CINTO:'CINTO',COLETE:'COLETE',TOP:'TOP',ACESSORIO:'ACESSÓRIO',BERMUDA:'BERMUDA'};
  return {name,cor,tamanho,categoria:catMap[first]||first};
}
function normalizeColor(c){ return norm(c).replace(/\s+-.*$/,'').replace(/\s+/g,' ').trim() || 'N/I'; }
function normalizeSize(s){ return norm(s).replace('ÚNICO','UNICO') || 'N/I'; }
function toRows(ws){ return XLSX.utils.sheet_to_json(ws,{defval:'',raw:true}); }
async function readArrayBuffer(file){ return await file.arrayBuffer(); }
async function readTextSmart(file){
  const buf=await file.arrayBuffer(); for(const enc of ['utf-8','windows-1252']){ try{return new TextDecoder(enc,{fatal:true}).decode(buf)}catch(e){} }
  return new TextDecoder('windows-1252').decode(buf);
}
function findHeader(row, variants){
  const entries=Object.keys(row); const targets=variants.map(norm); return entries.find(k=>targets.includes(norm(k)));
}
async function parseSales(file){
  const wb=XLSX.read(await readArrayBuffer(file),{type:'array',cellDates:true}); const rows=toRows(wb.Sheets[wb.SheetNames[0]]);
  if(!rows.length) throw new Error('A planilha de venda está vazia.');
  const sample=rows[0], codeKey=findHeader(sample,['Código','Codigo']), prodKey=findHeader(sample,['Produto']), qtyKey=findHeader(sample,['Quantidade','Qtd','Qtde']), valueKey=findHeader(sample,['Valor total NF','Valor Total','Faturamento','Valor']), dateKey=findHeader(sample,['Data','Data da venda','Data emissão','Data de emissão','Emissão']);
  if(!codeKey||!prodKey||!qtyKey||!valueKey) throw new Error('VENDA: não encontrei as colunas obrigatórias Código, Produto, Quantidade e Valor total NF.');
  return rows.map(r=>{ const code=r[codeKey]??''; const ref=referenceFromCode(code); const p=parseProduct(r[prodKey]); const q=numBR(r[qtyKey]); const value=numBR(r[valueKey]); return {codigo:String(code),referencia:ref,colecao:collectionFromReference(ref),produto:p.name,categoria:p.categoria,cor:normalizeColor(p.cor),tamanho:normalizeSize(p.tamanho),qtd:q,faturamento:value,data:dateKey?excelDate(r[dateKey]):null}; }).filter(x=>x.referencia);
}
function parseDelimited(text, delim=';'){
  return text.split(/\r?\n/).filter(x=>x.trim()).map(line=>{ const out=[]; let cur='',q=false; for(let i=0;i<line.length;i++){ const c=line[i]; if(c==='"'){ if(q&&line[i+1]==='"'){cur+='"';i++;} else q=!q; } else if(c===delim&&!q){out.push(cur);cur='';} else cur+=c; } out.push(cur); return out; });
}
async function parseStock(file){
  const rows=parseDelimited(await readTextSmart(file)); if(rows.length<2) throw new Error('ESTOQUE: arquivo vazio ou inválido.');
  const h=rows[0].map(x=>norm(x)); const idx=n=>h.indexOf(norm(n)); const ci=idx('Código'),pi=idx('Produto'),qi=idx('Quantidade');
  if(ci<0||pi<0||qi<0) throw new Error('ESTOQUE: não encontrei Código, Produto e Quantidade.');
  return rows.slice(1).map(r=>{const code=r[ci]; const ref=referenceFromCode(code); const p=parseProduct(r[pi]); return {codigo:String(code||''),referencia:ref,colecao:collectionFromReference(ref),produto:p.name,categoria:p.categoria,cor:normalizeColor(p.cor),tamanho:normalizeSize(p.tamanho),qtd:numBR(r[qi])};}).filter(x=>x.referencia);
}
async function parseOrder(file){
  const rows=parseDelimited(await readTextSmart(file)); const start=rows.findIndex(r=>norm(r[0])==='REFERENCIA');
  if(start<0) throw new Error('PEDIDO: não encontrei a seção “Referência” do pedido original.');
  const h=rows[start].map(x=>norm(x)); const sizeCols=h.map((x,i)=>({x,i})).filter(z=>['PP','P','M','G','GG','UNICO','36','38','40','42','44','46','48'].includes(z.x)); const totalI=h.indexOf('TOTAL'), priceI=h.indexOf('PRECO'), corI=h.indexOf('COR');
  const out=[];
  for(const r of rows.slice(start+1)){
    const m=String(r[0]||'').match(/^(\d{6})\s*-\s*(.*)$/); if(!m) continue;
    const ref=m[1], produto=m[2].trim(), cor=normalizeColor(r[corI]||''), price=priceI>=0?numBR(r[priceI]):0, cat=parseProduct(produto).categoria;
    for(const sc of sizeCols){ const q=numBR(r[sc.i]); if(q!==0) out.push({referencia:ref,colecao:collectionFromReference(ref),produto,categoria:cat,cor,tamanho:normalizeSize(sc.x),qtd:q,preco:price}); }
    if(!sizeCols.length && totalI>=0){ const q=numBR(r[totalI]); if(q) out.push({referencia:ref,colecao:collectionFromReference(ref),produto,categoria:cat,cor,tamanho:'N/I',qtd:q,preco:price}); }
  }
  if(!out.length) throw new Error('PEDIDO: nenhuma quantidade foi reconhecida.');
  return out;
}
function aggregate(rows,keyFn,fields){
  const m=new Map(); for(const r of rows){ const k=keyFn(r); if(!m.has(k))m.set(k,{...r}); else {const a=m.get(k); for(const f of fields)a[f]=(a[f]||0)+(r[f]||0); if(!a.produto&&r.produto)a.produto=r.produto;} } return m;
}
function collectCollections(){
  const m=new Map(); for(const r of [...state.sales,...state.stock,...state.order]){const c=r.colecao;if(!m.has(c.key))m.set(c.key,{...c,refs:new Set()});m.get(c.key).refs.add(r.referencia);} state.collections=m;
}
function validationRow(label,value,ok=true){return `<div class="validation-item"><span>${ok?'✅':'⚠️'} ${label}</span><span class="meta">${value}</span></div>`}
function renderCollections(){
  const box=$('collectionsList'); box.innerHTML=''; state.selected.clear();
  [...state.collections.values()].sort((a,b)=>a.key.localeCompare(b.key)).forEach(c=>{const id=`col-${c.key}`; const lab=document.createElement('label'); lab.className='collection-item'; lab.innerHTML=`<input type="checkbox" id="${id}" value="${c.key}" checked><span>${c.name}</span><small>${c.refs.size} refs.</small>`; box.appendChild(lab); state.selected.add(c.key); lab.querySelector('input').addEventListener('change',e=>e.target.checked?state.selected.add(c.key):state.selected.delete(c.key));});
}
function selectedFilter(r){return state.selected.has(r.colecao.key)}
function buildAnalysis(){
  const sales=state.sales.filter(selectedFilter), stock=state.stock.filter(selectedFilter), order=state.order.filter(selectedFilter);
  const sRef=aggregate(sales,r=>r.referencia,['qtd','faturamento']), stRef=aggregate(stock,r=>r.referencia,['qtd']), oRef=aggregate(order,r=>r.referencia,['qtd']);
  const refs=new Set([...sRef.keys(),...stRef.keys(),...oRef.keys()]); const master=[];
  for(const ref of refs){
    const s=sRef.get(ref)||{}, st=stRef.get(ref)||{}, o=oRef.get(ref)||{};
    const vendidoBruto=s.qtd||0, devolvido=Math.max(0,-vendidoBruto), vendido=Math.max(0,vendidoBruto), estoque=Math.max(0,st.qtd||0), pedido=Math.max(0,o.qtd||0), disponivel=vendido+estoque, extra=Math.max(0,disponivel-pedido), faturamento=s.faturamento||0, preco=o.preco||0;
    master.push({Referencia:ref,Colecao:collectionFromReference(ref).name,Produto:s.produto||st.produto||o.produto||'',Categoria:s.categoria||st.categoria||o.categoria||'',Preco_Tabela:preco,Pedido_Original:pedido,Vendido_Liq:vendido,Devolvido_Pecas:devolvido,Estoque_Atual:estoque,Disponivel_Real:disponivel,Entradas_Extras:extra,Giro_vs_Pedido:pedido?vendido/pedido:0,Giro_vs_Disponivel:disponivel?vendido/disponivel:0,Faturamento_Liq:faturamento,Ticket_Peca:vendido?faturamento/vendido:0,Capital_Estoque:preco?estoque*preco:0});
  }
  master.sort((a,b)=>b.Faturamento_Liq-a.Faturamento_Liq); return {sales,stock,order,master};
}
function categoryTable(master){
  const m=new Map(); for(const r of master){const k=r.Categoria||'OUTROS'; if(!m.has(k))m.set(k,{Categoria:k,Refs:new Set(),Pedido:0,Vendido:0,Estoque:0,Disponivel:0,Faturamento:0}); const a=m.get(k);a.Refs.add(r.Referencia);a.Pedido+=r.Pedido_Original;a.Vendido+=r.Vendido_Liq;a.Estoque+=r.Estoque_Atual;a.Disponivel+=r.Disponivel_Real;a.Faturamento+=r.Faturamento_Liq;}
  const arr=[...m.values()], tP=arr.reduce((s,x)=>s+x.Pedido,0), tV=arr.reduce((s,x)=>s+x.Vendido,0), tF=arr.reduce((s,x)=>s+x.Faturamento,0);
  return arr.map(x=>({Categoria:x.Categoria,Referencias:x.Refs.size,Pedido:x.Pedido,'% do Pedido':tP?x.Pedido/tP:0,'Vendido Líq':x.Vendido,'% das Peças':tV?x.Vendido/tV:0,'Estoque Atual':x.Estoque,'Disponível':x.Disponivel,'Giro % Pedido':x.Pedido?x.Vendido/x.Pedido:0,'Giro % Disponível':x.Disponivel?x.Vendido/x.Disponivel:0,'Faturamento Líq':x.Faturamento,'% do Faturamento':tF?x.Faturamento/tF:0,'Ticket Médio':x.Vendido?x.Faturamento/x.Vendido:0,'Índice Performance':(tP&&x.Pedido)?((tF?x.Faturamento/tF:0)/(x.Pedido/tP)):0})).sort((a,b)=>b['Faturamento Líq']-a['Faturamento Líq']);
}
function abcTable(master){
  const rows=master.filter(x=>x.Faturamento_Liq>0).sort((a,b)=>b.Faturamento_Liq-a.Faturamento_Liq), total=rows.reduce((s,x)=>s+x.Faturamento_Liq,0); let acc=0;
  return rows.map(r=>{acc+=r.Faturamento_Liq; const p=total?acc/total:0; return {Classe:p<=.8?'A':p<=.95?'B':'C',Referência:r.Referencia,Produto:r.Produto,Categoria:r.Categoria,'Vendido Líq':r.Vendido_Liq,'Faturamento Líq':r.Faturamento_Liq,'% Acumulado':p,'Estoque Atual':r.Estoque_Atual,'Giro % Disponível':r.Giro_vs_Disponivel,Ação:r.Giro_vs_Disponivel>=.7&&r.Estoque_Atual<=5?'REPOSIÇÃO / PROTEGER':r.Giro_vs_Disponivel<.15&&r.Estoque_Atual>=10?'AÇÃO COMERCIAL':'ACOMPANHAR'};});
}
function gradeTable(sales,stock,order){
  const map=new Map(), add=(rows,col)=>rows.forEach(r=>{const k=normalizeSize(r.tamanho); if(!map.has(k))map.set(k,{Tamanho:k,Pedido:0,Vendido:0,Estoque:0,Faturamento:0}); const a=map.get(k);a[col]+=r.qtd||0;if(col==='Vendido')a.Faturamento+=r.faturamento||0;});
  add(order,'Pedido');add(sales,'Vendido');add(stock,'Estoque'); const arr=[...map.values()],tp=arr.reduce((s,x)=>s+x.Pedido,0),tv=arr.reduce((s,x)=>s+x.Vendido,0);
  const orderRank=['PP','P','M','G','GG','UNICO','36','38','40','42','44','46','48','N/I'];
  return arr.map(x=>({Tamanho:x.Tamanho,Pedido:x.Pedido,'% do Pedido':tp?x.Pedido/tp:0,'Vendido Líq':x.Vendido,'% das Vendas':tv?x.Vendido/tv:0,'Desvio Grade':(tv?x.Vendido/tv:0)-(tp?x.Pedido/tp:0),'Estoque Atual':x.Estoque,'Disponível':x.Vendido+x.Estoque,'Giro % Pedido':x.Pedido?x.Vendido/x.Pedido:0,'Giro % Disponível':(x.Vendido+x.Estoque)?x.Vendido/(x.Vendido+x.Estoque):0,'Faturamento Líq':x.Faturamento})).sort((a,b)=>(orderRank.indexOf(a.Tamanho)<0?99:orderRank.indexOf(a.Tamanho))-(orderRank.indexOf(b.Tamanho)<0?99:orderRank.indexOf(b.Tamanho)));
}
function colorTable(sales,stock,order){
  const map=new Map(), add=(rows,col)=>rows.forEach(r=>{const k=normalizeColor(r.cor);if(!map.has(k))map.set(k,{Cor:k,Pedido:0,Vendido:0,Estoque:0,Faturamento:0});const a=map.get(k);a[col]+=r.qtd||0;if(col==='Vendido')a.Faturamento+=r.faturamento||0;});add(order,'Pedido');add(sales,'Vendido');add(stock,'Estoque');
  const arr=[...map.values()],tf=arr.reduce((s,x)=>s+x.Faturamento,0); return arr.map(x=>({Cor:x.Cor,Pedido:x.Pedido,'Vendido Líq':x.Vendido,'Estoque Atual':x.Estoque,'Disponível':x.Vendido+x.Estoque,'Giro % Disponível':(x.Vendido+x.Estoque)?x.Vendido/(x.Vendido+x.Estoque):0,'Faturamento Líq':x.Faturamento,'% do Faturamento':tf?x.Faturamento/tf:0})).sort((a,b)=>b['Faturamento Líq']-a['Faturamento Líq']);
}
function monthlyTable(sales){
  if(!sales.some(x=>x.data)) return [];
  const m=new Map(); for(const r of sales.filter(x=>x.data)){const k=`${r.data.getFullYear()}-${String(r.data.getMonth()+1).padStart(2,'0')}`; if(!m.has(k))m.set(k,{Mes:k,Refs:new Set(),Vendido:0,Faturamento:0});const a=m.get(k);a.Refs.add(r.referencia);a.Vendido+=r.qtd||0;a.Faturamento+=r.faturamento||0;}
  return [...m.values()].sort((a,b)=>a.Mes.localeCompare(b.Mes)).map(x=>({Mês:x.Mes,'Refs Ativas':x.Refs.size,'Vendido Líq':x.Vendido,'Faturamento Líq':x.Faturamento,'Ticket Médio':x.Vendido?x.Faturamento/x.Vendido:0}));
}
function daysBetween(a,b){const da=new Date(a+'T00:00:00'),db=new Date(b+'T00:00:00');return Math.max(1,Math.floor((db-da)/86400000)+1);}
function renderPreview(a){
  const totalPedido=a.master.reduce((s,x)=>s+x.Pedido_Original,0),vend=a.master.reduce((s,x)=>s+x.Vendido_Liq,0),est=a.master.reduce((s,x)=>s+x.Estoque_Atual,0),fat=a.master.reduce((s,x)=>s+x.Faturamento_Liq,0),disp=vend+est;
  $('summaryCard').classList.remove('hidden'); $('kpis').innerHTML=`<div class="kpi"><b>${number(totalPedido)}</b><span>Pedido original</span></div><div class="kpi"><b>${number(vend)}</b><span>Peças vendidas</span></div><div class="kpi"><b>${number(est)}</b><span>Estoque atual</span></div><div class="kpi"><b>${money(fat)}</b><span>Faturamento</span></div><div class="kpi"><b>${totalPedido?((vend/totalPedido)*100).toFixed(1).replace('.',',')+'%':'-'}</b><span>Giro vs pedido</span></div><div class="kpi"><b>${disp?((vend/disp)*100).toFixed(1).replace('.',',')+'%':'-'}</b><span>Giro vs disponível</span></div>`;
}
function styleCell(cell, style){cell.s={...(cell.s||{}),...style};}
function applyHeader(ws, range){
  const r=XLSX.utils.decode_range(range); for(let R=r.s.r;R<=r.e.r;R++)for(let C=r.s.c;C<=r.e.c;C++){const addr=XLSX.utils.encode_cell({r:R,c:C});if(ws[addr])styleCell(ws[addr],{fill:{fgColor:{rgb:THEME.navy}},font:{bold:true,color:{rgb:THEME.white}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{bottom:{style:'thin',color:{rgb:'AAB7C4'}}}});}
}
function applyNumberFormats(ws, headers, startRow, endRow){
  const pctHeaders=['Giro % Pedido','Giro % Disponível','Giro vs Pedido','Giro vs Disponível','% do Pedido','% das Peças','% das Vendas','% do Faturamento','% Acumulado','Índice Performance','Desvio Grade'];
  const moneyHeaders=['Preço Tabela','Faturamento Líq','Faturamento','Ticket Médio','Ticket por Peça','Capital Parado','Capital Estoque'];
  headers.forEach((h,i)=>{for(let r=startRow;r<=endRow;r++){const addr=XLSX.utils.encode_cell({r:r-1,c:i}); if(!ws[addr])continue; if(pctHeaders.includes(h))ws[addr].z=h==='Índice Performance'?'0.00':'0.0%'; if(moneyHeaders.includes(h))ws[addr].z='R$ #,##0.00';}});
}
function addSheet(wb,name,title,subtitle,rows,opts={}){
  const data=rows.length?rows:[{Mensagem:'Sem dados para os filtros selecionados'}], headers=Object.keys(data[0]);
  const aoa=[[title], [subtitle||''], [], headers, ...data.map(r=>headers.map(h=>r[h]))]; const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:Math.max(0,headers.length-1)}},{s:{r:1,c:0},e:{r:1,c:Math.max(0,headers.length-1)}}];
  if(ws.A1)styleCell(ws.A1,{font:{bold:true,sz:16,color:{rgb:THEME.navy}},alignment:{vertical:'center'}});
  if(ws.A2)styleCell(ws.A2,{font:{italic:true,sz:10,color:{rgb:'6B7280'}}});
  applyHeader(ws,`A4:${XLSX.utils.encode_col(headers.length-1)}4`); ws['!autofilter']={ref:`A4:${XLSX.utils.encode_col(headers.length-1)}${data.length+4}`}; ws['!freeze']={xSplit:0,ySplit:4};
  ws['!rows']=[{hpt:24},{hpt:18},{hpt:8},{hpt:32}];
  ws['!cols']=headers.map(h=>({wch: Math.min(38,Math.max(11,h.length+2, h.includes('Produto')?34:0))}));
  applyNumberFormats(ws,headers,5,data.length+4);
  if(opts.colorScaleHeader){const ci=headers.indexOf(opts.colorScaleHeader);if(ci>=0){for(let i=0;i<data.length;i++){const v=Number(data[i][opts.colorScaleHeader]||0),addr=XLSX.utils.encode_cell({r:i+4,c:ci});if(ws[addr]){const fill=v>=.6?THEME.green:v>=.3?THEME.yellow:v>0?THEME.orange:THEME.red;styleCell(ws[addr],{fill:{fgColor:{rgb:fill}},alignment:{horizontal:'center'}});}}}}
  XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31));
}
function generateWorkbook(){
  const a=buildAnalysis(); renderPreview(a); const wb=XLSX.utils.book_new();
  const start=$('dateStart').value,end=$('dateEnd').value,stockDate=$('stockDate').value,days=daysBetween(start,end),months=days/30.4375;
  const totalPedido=a.master.reduce((s,x)=>s+x.Pedido_Original,0),vend=a.master.reduce((s,x)=>s+x.Vendido_Liq,0),est=a.master.reduce((s,x)=>s+x.Estoque_Atual,0),disp=vend+est,fat=a.master.reduce((s,x)=>s+x.Faturamento_Liq,0),extras=a.master.reduce((s,x)=>s+x.Entradas_Extras,0),refsVend=a.master.filter(x=>x.Vendido_Liq>0).length,refsDisp=a.master.filter(x=>x.Disponivel_Real>0).length,semVenda=a.master.filter(x=>x.Vendido_Liq===0&&x.Estoque_Atual>0).length,ritmoMensal=months?vend/months:0,cobertura=ritmoMensal?est/ritmoMensal:0;
  const cols=[...state.selected].map(k=>state.collections.get(k)?.name).filter(Boolean).join(', '), sub=`Período: ${start} a ${end} • Estoque: ${stockDate||'não informado'} • Coleções: ${cols}`;
  const resumo=[
    {INDICADOR:'Peças do pedido original',VALOR:totalPedido,LEITURA:'Compra inicial do e-commerce'},
    {INDICADOR:'Entradas extras estimadas',VALOR:extras,LEITURA:'Reforço de estoque além do pedido original'},
    {INDICADOR:'Total disponibilizado',VALOR:disp,LEITURA:'Vendido no período + estoque atual'},
    {INDICADOR:'Peças vendidas (líquido)',VALOR:vend,LEITURA:'Quantidade reconhecida na base de venda'},
    {INDICADOR:'Estoque atual',VALOR:est,LEITURA:'Saldo na data informada'},
    {INDICADOR:'CENÁRIO 1 · Giro vs pedido original',VALOR:totalPedido?vend/totalPedido:0,LEITURA:'Do que foi comprado, quanto girou'},
    {INDICADOR:'CENÁRIO 2 · Giro vs disponível real',VALOR:disp?vend/disp:0,LEITURA:'Considerando entradas extras'},
    {INDICADOR:'Faturamento líquido',VALOR:fat,LEITURA:'Receita informada no relatório de vendas'},
    {INDICADOR:'Ticket médio por peça',VALOR:vend?fat/vend:0,LEITURA:'Receita média por peça'},
    {INDICADOR:'Referências vendidas',VALOR:refsVend,LEITURA:`de ${refsDisp} referências com disponibilidade`},
    {INDICADOR:'Referências sem nenhuma venda',VALOR:semVenda,LEITURA:'Com estoque atual e venda zero no período'},
    {INDICADOR:'Ritmo médio de venda',VALOR:ritmoMensal,LEITURA:'Peças por 30,4 dias no período informado'},
    {INDICADOR:'Cobertura do estoque atual',VALOR:cobertura,LEITURA:'Meses de estoque no ritmo médio do período'}
  ];
  addSheet(wb,'1. Resumo Executivo',`CHARRY | Desempenho ${cols||'Coleções selecionadas'}`,sub,resumo);
  const wsR=wb.Sheets['1. Resumo Executivo']; for(let r=5;r<=resumo.length+4;r++){const ind=resumo[r-5].INDICADOR,addr=`B${r}`;if(!wsR[addr])continue;if(ind.includes('Giro'))wsR[addr].z='0.0%';else if(ind.includes('Faturamento')||ind.includes('Ticket'))wsR[addr].z='R$ #,##0.00';else if(ind.includes('Cobertura'))wsR[addr].z='0.0';}

  const c1=a.master.filter(x=>x.Pedido_Original>0).map(x=>({Referência:x.Referencia,Produto:x.Produto,Categoria:x.Categoria,'Preço Tabela':x.Preco_Tabela,Pedido:x.Pedido_Original,'Vendido Líq':x.Vendido_Liq,'Saldo vs Pedido':x.Pedido_Original-x.Vendido_Liq,'Giro % Pedido':x.Giro_vs_Pedido,'Estoque Atual':x.Estoque_Atual,'Faturamento Líq':x.Faturamento_Liq}));
  addSheet(wb,'2. Cenário 1 - vs Pedido','CENÁRIO 1 · Giro sobre o pedido original','Quanto de cada referência comprada no pedido original já foi vendida.',c1,{colorScaleHeader:'Giro % Pedido'});
  const c2=a.master.map(x=>({Referência:x.Referencia,Produto:x.Produto,Categoria:x.Categoria,Origem:x.Pedido_Original>0?(x.Entradas_Extras>0?'Pedido + Estoque':'Pedido'):'Só estoque (entrada extra)',Pedido:x.Pedido_Original,'Entrada Extra':x.Entradas_Extras,Disponível:x.Disponivel_Real,'Vendido Líq':x.Vendido_Liq,'Estoque Atual':x.Estoque_Atual,'Giro % Disponível':x.Giro_vs_Disponivel,'Faturamento Líq':x.Faturamento_Liq}));
  addSheet(wb,'3. Cenário 2 - vs Estoque','CENÁRIO 2 · Giro sobre o disponível real','Disponível = vendido no período + estoque atual. Entradas extras = disponível - pedido, quando positivo.',c2,{colorScaleHeader:'Giro % Disponível'});
  addSheet(wb,'4. Cenário 3 - Categoria','CENÁRIO 3 · Giro e representatividade por categoria','Índice de performance = % do faturamento ÷ % do pedido. Acima de 1,00 = categoria rende mais do que pesou na compra.',categoryTable(a.master),{colorScaleHeader:'Giro % Disponível'});
  addSheet(wb,'5. Curva ABC','CURVA ABC · Concentração do faturamento','Classe A = até 80% do faturamento • B = até 95% • C = cauda final.',abcTable(a.master),{colorScaleHeader:'Giro % Disponível'});
  addSheet(wb,'6. Grade e Cor','GRADE · Giro por tamanho','Desvio grade = participação nas vendas menos participação no pedido. Positivo = vendeu proporcionalmente mais.',gradeTable(a.sales,a.stock,a.order),{colorScaleHeader:'Giro % Disponível'});
  addSheet(wb,'6B. Cor','CORES · Giro e faturamento','Ranking de cores por faturamento no período informado.',colorTable(a.sales,a.stock,a.order),{colorScaleHeader:'Giro % Disponível'});
  const monthly=monthlyTable(a.sales);
  if(monthly.length) addSheet(wb,'7. Evolução Mensal','EVOLUÇÃO MENSAL DA COLEÇÃO','Gerada automaticamente porque a base de venda possui uma coluna de data reconhecida.',monthly);
  else addSheet(wb,'7. Evolução Mensal','EVOLUÇÃO MENSAL DA COLEÇÃO','A base atual de venda é consolidada e não traz data por linha. O período informado é usado para ritmo e cobertura, mas não permite dividir os meses.',[{Status:'Evolução mensal indisponível nesta carga',Período:`${start} a ${end}`,Orientação:'Se futuramente o relatório de venda trouxer Data/Data da venda, esta aba será preenchida automaticamente.'}]);
  const sem=a.master.filter(x=>x.Vendido_Liq===0&&x.Estoque_Atual>0).map(x=>({Referência:x.Referencia,Produto:x.Produto,Categoria:x.Categoria,Origem:x.Pedido_Original>0?(x.Entradas_Extras>0?'Pedido + Estoque':'Pedido'):'Só estoque (entrada extra)',Pedido:x.Pedido_Original,'Estoque Parado':x.Estoque_Atual,'Preço Tabela':x.Preco_Tabela,'Capital Parado':x.Preco_Tabela?x.Estoque_Atual*x.Preco_Tabela:0})).sort((a,b)=>b['Capital Parado']-a['Capital Parado']);
  addSheet(wb,'8. Sem Venda','REFERÊNCIAS SEM NENHUMA VENDA','Produtos sem uma única venda no período e com estoque atual. Prioridade para ação comercial.',sem);
  addSheet(wb,'9. Base Mestre','BASE MESTRE · Auditoria','Base consolidada que alimenta todas as análises.',a.master);
  const fname=`Analise_Giro_${[...state.selected].join('-')||'colecoes'}_${end||'periodo'}.xlsx`; XLSX.writeFile(wb,fname,{bookType:'xlsx'});
}

$('analyzeBtn').addEventListener('click',async()=>{
  const files=[$('salesFile').files[0],$('stockFile').files[0],$('orderFile').files[0]]; if(files.some(x=>!x)){alert('Selecione os 3 arquivos antes de continuar.');return;}
  $('analyzeBtn').disabled=true;$('analyzeBtn').textContent='Lendo arquivos...';
  try{
    state={sales:await parseSales(files[0]),stock:await parseStock(files[1]),order:await parseOrder(files[2]),collections:new Map(),selected:new Set()}; collectCollections();
    const noDate=!state.sales.some(x=>x.data), unknown=state.collections.has('NA');
    $('validationCard').classList.remove('hidden');$('collectionsCard').classList.remove('hidden');$('statusBadge').textContent='Arquivos válidos';$('statusBadge').className='badge ok';
    $('validationList').innerHTML=validationRow('Venda',`${state.sales.length} linhas reconhecidas`)+validationRow('Estoque',`${state.stock.length} SKUs reconhecidos`)+validationRow('Pedido original',`${state.order.length} combinações referência/tamanho reconhecidas`)+validationRow('Coleções',`${state.collections.size} grupos identificados`,!unknown)+validationRow('Data na venda',noDate?'Não existe — evolução mensal ficará informativa':'Reconhecida — evolução mensal será gerada',!noDate);
    renderCollections();
  }catch(e){$('validationCard').classList.remove('hidden');$('statusBadge').textContent='Revisar arquivo';$('statusBadge').className='badge warn';$('validationList').innerHTML=validationRow('Erro',e.message,false);console.error(e);}
  finally{$('analyzeBtn').disabled=false;$('analyzeBtn').textContent='Ler e validar arquivos';}
});
$('toggleAllBtn').addEventListener('click',()=>{const boxes=[...document.querySelectorAll('#collectionsList input[type=checkbox]')];const all=boxes.every(x=>x.checked);boxes.forEach(x=>{x.checked=!all;x.dispatchEvent(new Event('change'));});$('toggleAllBtn').textContent=all?'Selecionar todas':'Desmarcar todas';});
$('generateBtn').addEventListener('click',()=>{if(!state.selected.size){alert('Selecione ao menos uma coleção.');return;}if(!$('dateStart').value||!$('dateEnd').value){alert('Informe o período inicial e final das vendas.');return;}if($('dateEnd').value<$('dateStart').value){alert('A data final não pode ser anterior à data inicial.');return;}generateWorkbook();});
