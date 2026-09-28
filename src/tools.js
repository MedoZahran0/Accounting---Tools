import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';

// PDF.js worker: Vite must be given the worker URL explicitly, otherwise
// PDF rendering fails in the browser with GlobalWorkerOptions.workerSrc.
pdfjsLib.GlobalWorkerOptions.workerSrc=new URL('pdfjs-dist/build/pdf.worker.mjs',import.meta.url).toString();
import {PDFDocument,degrees,StandardFonts,rgb} from 'pdf-lib';
import {createWorker} from 'tesseract.js';

const AR='٠١٢٣٤٥٦٧٨٩';
export const toEnglish=v=>String(v??'').replace(/[٠-٩]/g,d=>String(AR.indexOf(d)));
export const normalize=v=>toEnglish(v).replace(/[\u200e\u200f]/g,'').replace(/\s+/g,' ').trim();
export const safeName=(n='result')=>n.replace(/\.[^.]+$/,'').replace(/[^\w\u0600-\u06FF-]+/g,'_')||'result';
export const download=(blob,name)=>{const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),1500)};
export const bytesDownload=(b,n,t='application/octet-stream')=>download(new Blob([b],{type:t}),n,t);
const pdfBytes=async f=>{
  if(f instanceof Uint8Array)return f;
  if(f instanceof ArrayBuffer)return new Uint8Array(f);
  if(ArrayBuffer.isView(f))return new Uint8Array(f.buffer,f.byteOffset,f.byteLength);
  if(!f||typeof f.arrayBuffer!=='function')throw Error('ملف PDF غير صالح أو غير قابل للقراءة');
  return new Uint8Array(await f.arrayBuffer());
};
export const loadPdf=async f=>{
  const data=await pdfBytes(f);
  const head=new TextDecoder('latin1').decode(data.slice(0,1024));
  if(!head.includes('%PDF-'))throw Error('الملف المحدد ليس PDF صالحاً (لم يتم العثور على ترويسة PDF). اختر ملف PDF حقيقياً ثم جرّب مرة أخرى.');
  try{
    // Preview must remain reliable even when the browser cannot start PDF.js's worker.
    // The worker is faster, but the main-thread fallback is intentionally supported.
    try{
      return await pdfjsLib.getDocument({
        data,
        useWorkerFetch:false,
        isEvalSupported:false,
        disableAutoFetch:false
      }).promise;
    }catch(workerError){
      return await pdfjsLib.getDocument({
        data,
        disableWorker:true,
        useWorkerFetch:false,
        isEvalSupported:false,
        disableAutoFetch:false
      }).promise;
    }
  }catch(e){
    const m=String(e?.message||e);
    if(/No PDF header found|Invalid PDF|InvalidPDFException/i.test(m))throw Error('تعذر قراءة هذا الملف كـPDF. قد يكون الملف تالفاً أو ليس PDF حقيقياً.');
    throw e;
  }
};
export const renderPage=async(page,scale=2)=>{const v=page.getViewport({scale}),c=document.createElement('canvas');c.width=Math.max(1,Math.ceil(v.width));c.height=Math.max(1,Math.ceil(v.height));const ctx=c.getContext('2d',{willReadFrequently:true,alpha:false});ctx.fillStyle='#fff';ctx.fillRect(0,0,c.width,c.height);await page.render({canvasContext:ctx,viewport:v,background:'#fff'}).promise;return c};

const rowGroup=(items,tolerance=0.55)=>{
  const rows=[];
  for(const i of items.filter(x=>String(x.str??x.text??'').trim())){
    const x=Number(i.transform?.[4]??i.bbox?.x0??0), y=Number(i.transform?.[5]??((i.bbox?.y0??0)+(i.bbox?.y1??0))/2);
    const fallbackH=((i.bbox?.y1??0)-(i.bbox?.y0??0));
    const rawH=Number(i.transform?.[3]??i.transform?.[0]??fallbackH);
    const h=Math.max(6,Math.abs(Number.isFinite(rawH)?rawH:10));
    const text=String(i.str??i.text??'').trim();
    const fallbackW=((i.bbox?.x1??0)-(i.bbox?.x0??0));
    const rawW=Number(i.width??fallbackW);
    const w=Number.isFinite(rawW)?rawW:0;
    let r=rows.find(z=>Math.abs(z.y-y)<=Math.max(3,h*tolerance));
    if(!r){r={y,h,items:[]};rows.push(r)}
    r.items.push({x,w,h,text});
  }
  return rows.sort((a,b)=>b.y-a.y).map(r=>({...r,items:r.items.sort((a,b)=>a.x-b.x)}));
};
export const groupItemsDetailed=items=>rowGroup(items);
export const groupItems=items=>rowGroup(items).map(r=>r.items.map(i=>i.text));
export const wordsToRowsDetailed=words=>rowGroup(words.map(w=>({text:w.text,bbox:w.bbox})),0.75).sort((a,b)=>a.y-b.y);
export const wordsToRows=rows=>rows.map(r=>r.items.map(i=>i.text));

const dateValue=s=>{const x=normalize(s);return /^(?:\d{1,2}[\/-]\d{1,2}(?:[\/-]\d{2,4})?|\d{4}[\/-]\d{1,2}[\/-]\d{1,2})$/.test(x)?x:''};
const numberValue=s=>{let x=normalize(s).replace(/[,٬،\s]/g,'').replace(/[^\d.()\-]/g,'');if(!x)return null;if(/^\(.*\)$/.test(x))x='-'+x.slice(1,-1);const n=Number(x);return Number.isFinite(n)?n:null};
const headerKind=s=>{const x=normalize(s).toLowerCase();if(/date|التاريخ|تاريخ/.test(x))return'date';if(/description|details|narration|البيان|الوصف|التفاصيل|الحركة/.test(x))return'description';if(/debit|withdraw|مدين|سحب|خصم/.test(x))return'debit';if(/credit|deposit|دائن|إيداع|إضافة/.test(x))return'credit';if(/balance|الرصيد|رصيد/.test(x))return'balance';return''};

export const detectStatementColumns=(detailedRows)=>{
  const rows=detailedRows.slice(0,45), maxX=Math.max(1,...rows.flatMap(r=>r.items.map(i=>i.x+i.w)));
  let best=null;
  for(const r of rows){
    const hs=r.items.map(i=>({x:i.x+i.w/2,k:headerKind(i.text)})).filter(x=>x.k);
    if(hs.length>=2 && (!best||hs.length>best.length))best=hs;
  }
  const headerCols=best?.map(h=>({kind:h.k,x:h.x}))||[];
  const numericByX=[];
  for(const r of detailedRows.slice(0,80))for(const i of r.items)if(numberValue(i.text)!==null)numericByX.push(i.x+i.w/2);
  const clusters=[];
  for(const x of numericByX.sort((a,b)=>a-b)){let c=clusters.find(c=>Math.abs(c.x-x)<Math.max(12,maxX*.018));if(!c)c={x,n:0};c.x=(c.x*c.n+x)/(c.n+1);c.n++;if(!clusters.includes(c))clusters.push(c)}
  clusters.sort((a,b)=>a.x-b.x);
  const cols=headerCols.length?headerCols:clusters.slice(0,6).map((c,i)=>({x:c.x,kind:['date','description','debit','credit','balance'][i]||'amount'}));
  return {columns:cols,confidence:cols.length>=3?'جيدة':cols.length===2?'متوسطة':'مبدئية'};
};


function canonicalBankRows(detailed,statement){
  const cols=statement.columns||[];
  const byKind={};
  for(const c of cols)if(!byKind[c.kind])byKind[c.kind]=c;
  const ordered=['date','description','debit','credit','balance'];
  const out=[['التاريخ','الوصف','المدين','الدائن','الرصيد']];
  for(const r of detailed){
    const cells={date:[],description:[],debit:[],credit:[],balance:[]};
    for(const i of r.items){
      const x=i.x+i.w/2;
      let bestKind='',bestDist=Infinity;
      for(const kind of ordered){
        const col=byKind[kind]; if(!col) continue;
        const d=Math.abs(x-col.x);
        if(d<bestDist){bestDist=d;bestKind=kind}
      }
      if(bestKind)cells[bestKind].push(i.text);
    }
    const date=cells.date.join(' ').trim();
    const description=cells.description.join(' ').trim();
    const debit=cells.debit.map(numberValue).find(v=>v!==null);
    const credit=cells.credit.map(numberValue).find(v=>v!==null);
    const balance=cells.balance.map(numberValue).find(v=>v!==null);
    const all=[date,description,debit,credit,balance];
    const hasDate=!!dateValue(date);
    const hasAmount=[debit,credit,balance].some(v=>v!==undefined&&v!==null);
    const headerText=normalize(r.items.map(i=>i.text).join(' ')).toLowerCase();
    const isHeader=/date|description|details|debit|credit|balance|التاريخ|الوصف|البيان|مدين|دائن|الرصيد/.test(headerText);
    if(!isHeader&&(hasDate||hasAmount)&&description)out.push([
      hasDate?date:'',
      description,
      credit??'',
      debit??'',
      balance??''
    ]);
  }
  return out.length>1?out:[['التاريخ','الوصف','المدين','الدائن','الرصيد']];
}
function assignStatementRows(detailed,statement){
  if(!statement.columns.length)return detailed.map(r=>r.items.map(i=>i.text));
  const cols=[...statement.columns].sort((a,b)=>a.x-b.x), out=[];
  for(const r of detailed){
    const cells=cols.map(c=>[]);
    for(const i of r.items){
      const x=i.x+i.w/2;let best=0,dist=Infinity;
      cols.forEach((c,j)=>{const d=Math.abs(x-c.x);if(d<dist){dist=d;best=j}});
      cells[best].push(i.text);
    }
    const row=cells.map(c=>c.join(' ').trim());
    if(row.some(Boolean))out.push(row);
  }
  return out;
}

export async function pdfRows(file,onProgress){
  const pdf=await loadPdf(file), digitalRows=[], ocrRows=[];let digitalPages=0;
  for(let p=1;p<=pdf.numPages;p++){
    const page=await pdf.getPage(p), rows=groupItemsDetailed((await page.getTextContent()).items);
    if(rows.length){digitalPages++;digitalRows.push(...rows.map(r=>({...r,page:p})))} 
    onProgress?.(Math.round(p/pdf.numPages*25),`قراءة الصفحة ${p} من ${pdf.numPages}`);
  }
  let mode='text',detailed=digitalRows;
  if(digitalPages<Math.max(1,Math.ceil(pdf.numPages*.65))){
    mode='ocr';detailed=[];
    const worker=await createWorker('ara+eng',1,{logger:m=>onProgress?.(25+Math.round((m.progress||0)*60),m.status||'OCR')});
    for(let p=1;p<=pdf.numPages;p++){
      const c=await renderPage(await pdf.getPage(p),2.1),res=await worker.recognize(c);
      detailed.push(...wordsToRowsDetailed(res.data.words||[]).map(r=>({...r,page:p})));
      onProgress?.(25+Math.round(p/pdf.numPages*60),`OCR الصفحة ${p} من ${pdf.numPages}`);
    }
    await worker.terminate();
  }
  const statement=detectStatementColumns(detailed);
  const rawRows=assignStatementRows(detailed,statement);
  const statementKinds=new Set(statement.columns.map(c=>c.kind));
  const looksLikeBank=statementKinds.has('date')&&statementKinds.has('description')&&statementKinds.has('balance')&&(statementKinds.has('debit')||statementKinds.has('credit'));
  const rows=looksLikeBank?canonicalBankRows(detailed,statement):rawRows;
  onProgress?.(100,'اكتملت القراءة');
  return {rows,mode,detailed,statement,isBankStatement:looksLikeBank};
}

export const rowsWorkbook=(rows,sheet='Data')=>{
  const data=rows.filter(r=>r?.some(v=>String(v??'').trim())).map(r=>r.map(v=>v??''));
  const max=Math.max(1,...data.map(r=>r.length)),a=data.map(r=>{const x=r.slice(0,max);while(x.length<max)x.push('');return x});
  const ws=XLSX.utils.aoa_to_sheet(a.length?a:[['لا توجد بيانات']]);
  ws['!cols']=Array.from({length:max},(_,c)=>({wch:Math.min(45,Math.max(12,...a.slice(0,150).map(r=>String(r[c]??'').length+2)))}));
  ws['!sheetViews']=[{rightToLeft:true}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheet);return wb;
};
export const exportRows=(rows,name)=>XLSX.writeFile(rowsWorkbook(rows),name);
export const excelRows=async f=>{const wb=XLSX.read(await f.arrayBuffer(),{type:'array',cellDates:true});return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:''})};
const buildSequentialOrder=docs=>{const out=[];for(let fi=0;fi<docs.length;fi++){const count=docs[fi]?.getPageCount?.()||0;for(let page=1;page<=count;page++)out.push({fileIndex:fi,page})}return out};
const normalizeMergeOrder=(docs,order)=>{
 const all=buildSequentialOrder(docs);
 if(!Array.isArray(order)||order.length!==all.length)return all;
 const seen=new Set(),safe=[];
 for(const item of order){
  const fi=Number(item?.fileIndex),page=Number(item?.page),count=docs[fi]?.getPageCount?.()||0;
  if(!Number.isInteger(fi)||fi<0||fi>=docs.length||!Number.isInteger(page)||page<1||page>count)return all;
  const key=fi+':'+page;if(seen.has(key))return all;seen.add(key);safe.push({fileIndex:fi,page});
 }
 return safe.length===all.length?safe:all;
};
const mergePreservingPdf=async(files,order,onProgress)=>{
 const out=await PDFDocument.create();
 const sources=await Promise.all(files.map(async f=>{
  const bytes=await pdfBytes(f);
  return PDFDocument.load(bytes,{ignoreEncryption:true,updateMetadata:false});
 }));
 const safeOrder=normalizeMergeOrder(sources,order);
 if(!safeOrder.length)throw Error('لم يتم العثور على صفحات صالحة للدمج');
 for(let i=0;i<safeOrder.length;i++){
  const item=safeOrder[i],src=sources[item.fileIndex];
  if(!src)throw Error('ملف المصدر غير موجود للصفحة '+(i+1));
  const copied=await out.copyPages(src,[item.page-1]);
  if(copied.length!==1)throw Error('تعذر نسخ الصفحة '+item.page+' من الملف '+(item.fileIndex+1));
  out.addPage(copied[0]);
  onProgress?.(Math.round((i+1)/safeOrder.length*100),'جاري دمج الصفحة '+(i+1)+' من '+safeOrder.length);
 }
 if(out.getPageCount()!==safeOrder.length)throw Error('فشل إنشاء صفحات ملف الدمج');
 const bytes=await out.save({useObjectStreams:false,addDefaultPage:false});
 const check=await PDFDocument.load(bytes,{ignoreEncryption:true,updateMetadata:false});
 if(check.getPageCount()!==safeOrder.length)throw Error('تم إنشاء ملف دمج غير صالح');
 return bytes;
};
export const mergePdfs=async(files,onProgress)=>{
 const docs=await Promise.all(files.map(async f=>{
  const bytes=await pdfBytes(f);
  return PDFDocument.load(bytes,{ignoreEncryption:true,updateMetadata:false});
 }));
 return mergePreservingPdf(files,buildSequentialOrder(docs),onProgress);
};
export const mergeOrderedPdfs=async(files,order,onProgress)=>{
 const docs=await Promise.all(files.map(async f=>{
  const bytes=await pdfBytes(f);
  return PDFDocument.load(bytes,{ignoreEncryption:true,updateMetadata:false});
 }));
 return mergePreservingPdf(files,normalizeMergeOrder(docs,order),onProgress);
};
export const pagesPdf=async(f,numbers)=>{const s=await PDFDocument.load(await f.arrayBuffer(),{ignoreEncryption:true}),o=await PDFDocument.create(),idx=numbers.map(Number).map(n=>n-1).filter(n=>n>=0&&n<s.getPageCount());if(!idx.length)throw Error('لم يتم تحديد صفحات صحيحة');(await o.copyPages(s,idx)).forEach(p=>o.addPage(p));return o.save()};
export const rotatePdf=async(f,a)=>{const p=await PDFDocument.load(await f.arrayBuffer());p.getPages().forEach(x=>x.setRotation(degrees((x.getRotation().angle+a+360)%360)));return p.save()};
export const annotatePdf=async(f,kind,text,opts={})=>{
 const pdf=await loadPdf(f),out=await PDFDocument.create();
 const size=Math.max(8,Math.min(180,Number(opts.size)||24));
 const opacity=Math.max(0,Math.min(1,Number(opts.opacity)??.22));
 const angle=Number(opts.angle)||35;
 const label=String(text||'ACCOUNTING TOOLS').trim()||'ACCOUNTING TOOLS';
 for(let i=1;i<=pdf.numPages;i++){
  const srcPage=await pdf.getPage(i),vp=srcPage.getViewport({scale:3.2});
  const canvas=document.createElement('canvas');canvas.width=Math.ceil(vp.width);canvas.height=Math.ceil(vp.height);
  const ctx=canvas.getContext('2d',{willReadFrequently:true});
  await srcPage.render({canvasContext:ctx,viewport:vp}).promise;
  ctx.save();
  if(kind==='number'){
   const px=Math.max(24,Math.round(size*3.2));
   ctx.fillStyle='rgba(90,90,90,.9)';ctx.font=px+'px Arial';ctx.textAlign='right';ctx.textBaseline='bottom';
   ctx.fillText(String(i),vp.width-58,vp.height-34);
  }else{
   const px=Math.max(24,Math.round(size*3.2));
   ctx.translate(vp.width/2,vp.height/2);ctx.rotate(angle*Math.PI/180);
   ctx.globalAlpha=opacity;ctx.fillStyle='rgb(90,90,90)';ctx.font='bold '+px+'px Arial';ctx.textAlign='center';ctx.textBaseline='middle';
   ctx.fillText(label,0,0);
  }
  ctx.restore();
  const png=await new Promise((resolve,reject)=>canvas.toBlob(async b=>b?resolve(await b.arrayBuffer()):reject(Error('تعذر تجهيز الصفحة '+i)),'image/png'));
  const img=await out.embedPng(png),page=out.addPage([vp.width/3.2,vp.height/3.2]);
  page.drawImage(img,{x:0,y:0,width:page.getWidth(),height:page.getHeight()});
  canvas.width=canvas.height=1;
  opts.onProgress?.(Math.round(i/pdf.numPages*100),'جاري تجهيز الصفحة '+i+' من '+pdf.numPages);
 }
 return out.save({useObjectStreams:false});
};
export const rasterCompress=async(f,quality,onProgress)=>{const p=await loadPdf(f),o=await PDFDocument.create();for(let i=1;i<=p.numPages;i++){const page=await p.getPage(i),vp=page.getViewport({scale:1}),scale=quality<.55?1:quality<.8?1.25:1.6,c=await renderPage(page,scale),b=await new Promise((resolve,reject)=>c.toBlob(x=>x?x.arrayBuffer().then(resolve,reject):reject(Error('تعذر ضغط الصفحة '+i)),'image/jpeg',quality)),img=await o.embedJpg(b),pg=o.addPage([vp.width,vp.height]);pg.drawImage(img,{x:0,y:0,width:vp.width,height:vp.height});c.width=c.height=1;onProgress?.(Math.round(i/p.numPages*100))}return o.save()};
export const imagesPdf=async fs=>{const o=await PDFDocument.create();for(const f of fs){const b=await f.arrayBuffer(),img=f.type.includes('png')?await o.embedPng(b):await o.embedJpg(b),p=o.addPage([img.width,img.height]);p.drawImage(img,{x:0,y:0,width:img.width,height:img.height})}return o.save()};

const AI_ENDPOINT=import.meta.env.VITE_AI_ENDPOINT||'/api/ai';

const repairArabicText=v=>{let s=String(v??'').normalize('NFKC').replace(/[\u200E\u200F\u202A-\u202E\u2066-\u2069]/g,'').trim();if(!s)return s;const tokens=s.split(/\s+/).filter(Boolean);const isolated=tokens.filter(t=>/^[\u0600-\u06FF]$/.test(t)).length;const ar=(s.match(/[\u0600-\u06FF]/g)||[]).length;if(ar>=4&&isolated>=Math.max(3,Math.ceil(ar*.55))){const joined=tokens.join('');return Array.from(joined).reverse().join('')}return s};
const normalizeAiRows=rows=>Array.isArray(rows)?rows.map(r=>Array.isArray(r)?r.map(v=>String(v??'').trim()):[]):[];

export async function aiRefineRows(rows,{endpoint=AI_ENDPOINT,signal,task='document_ocr',file=null}={}){
  const cleaned=normalizeAiRows(rows).map(r=>r.map(v=>v.replace(/[\u200B\u200C\u200D\uFEFF]/g,'').replace(/[ \\t]+/g,' ').trim()));
  const local=cleaned.map(r=>r.map(v=>repairArabicText(v.replace(/[|¦]/g,'I').replace(/[٠-٩]/g,d=>toEnglish(d)))));
  if(!endpoint)return{rows:local,source:'local',confidence:'محلي'};
  try{
    const body={task,rows:local,language:'ar+en'};
    if(file){
      const bytes=new Uint8Array(await file.arrayBuffer());
      let binary='';const chunk=0x8000;
      for(let p=0;p<bytes.length;p+=chunk)binary+=String.fromCharCode(...bytes.subarray(p,p+chunk));
      body.pdfBase64=btoa(binary);body.mimeType=file.type||'application/pdf';
    }
    const res=await fetch(endpoint,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body),signal});
    if(!res.ok)throw Error('AI endpoint '+res.status);
    const data=await res.json(),refined=normalizeAiRows(data?.rows);
    if(!Array.isArray(data?.rows)||!refined.length)throw Error('Invalid AI response');
    if(local.length&&refined.length!==local.length)throw Error('AI changed row count');
    const safe=local.length?refined.map((r,i)=>r.length===local[i].length?r.map((v,j)=>{const ai=repairArabicText(v),src=repairArabicText(local[i][j]);const aiAr=(ai.match(/[\u0600-\u06FF]/g)||[]).length,srcAr=(src.match(/[\u0600-\u06FF]/g)||[]).length;const aiWords=(ai.match(/\s+/g)||[]).length,srcWords=(src.match(/\s+/g)||[]).length;if(aiAr>=4&&srcAr>=4&&srcWords>0&&aiWords===0&&src.length>ai.length*.8)return src;return ai;}):local[i]):refined;
    return{rows:safe,source:'ai',confidence:data?.confidence||'AI'};
  }catch(e){return{rows:local,source:'local-fallback',confidence:'محلي بعد تعذر AI'};}
}
