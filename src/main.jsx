import React,{useMemo,useRef,useState} from 'react';
import {createRoot} from 'react-dom/client';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import {PDFDocument} from 'pdf-lib';
import * as XLSX from 'xlsx';
import {createWorker} from 'tesseract.js';
import {Document as WordDocument,Paragraph,TextRun,Packer} from 'docx';
import {FileSpreadsheet,FileText,Image as ImageIcon,Search,Sun,Moon,Upload,ShieldCheck,WandSparkles,Merge,Scissors,Download,Trash2,X,Play,FileDown,Table2,LoaderCircle} from 'lucide-react';
import './styles.css';

pdfjsLib.GlobalWorkerOptions.workerSrc=pdfWorker;

const OWNER='Accountant Abdel-hamid Zahran';
const PHONE='01095247005';

const tools=[
 {id:'pdf-excel',cat:'PDF',title:'PDF إلى Excel',desc:'حوّل PDF النصي أو الممسوح ضوئياً إلى Excel قابل للتعديل مع ترتيب الصفوف والأعمدة.',icon:FileSpreadsheet,featured:true,accept:'.pdf'},
 {id:'merge-pdf',cat:'PDF',title:'دمج PDF',desc:'ادمج عدة ملفات PDF في ملف واحد بالترتيب الذي تختاره.',icon:Merge,accept:'.pdf'},
 {id:'split-pdf',cat:'PDF',title:'تقسيم PDF',desc:'استخرج صفحات محددة أو نطاقاً من الصفحات إلى ملف PDF جديد.',icon:Scissors,accept:'.pdf'},
 {id:'pdf-images',cat:'PDF',title:'PDF إلى صور',desc:'حوّل كل صفحة إلى PNG عالية الدقة قابلة للتحميل.',icon:ImageIcon,accept:'.pdf'},
 {id:'pdf-word',cat:'PDF',title:'PDF إلى Word',desc:'استخرج النص من PDF وأنشئ ملف Word قابل للتحرير.',icon:FileText,accept:'.pdf'},
 {id:'excel-clean',cat:'Excel',title:'تنظيف Excel',desc:'احذف الصفوف الفارغة والتكرارات ونظّف المسافات.',icon:FileSpreadsheet,accept:'.xlsx,.xls,.csv'},
 {id:'excel-csv',cat:'Excel',title:'Excel ↔ CSV',desc:'حوّل ملفات Excel إلى CSV أو CSV إلى Excel.',icon:FileSpreadsheet,accept:'.xlsx,.xls,.csv'},
 {id:'image-pdf',cat:'صور',title:'صور إلى PDF',desc:'اجمع JPG وPNG في ملف PDF واحد بالترتيب.',icon:ImageIcon,accept:'image/*'},
 {id:'ocr',cat:'OCR',title:'OCR استخراج النص',desc:'استخرج النص العربي والإنجليزي من الصور والمستندات الممسوحة.',icon:WandSparkles,accept:'image/*,.pdf'}
];

function downloadBlob(blob,name){
 const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download=name; a.click();
 setTimeout(()=>URL.revokeObjectURL(url),1000);
}
function downloadBytes(bytes,name,type='application/octet-stream'){downloadBlob(new Blob([bytes],{type}),name)}
function safeName(name){return name.replace(/\.[^.]+$/,'').replace(/[^\w\u0600-\u06FF-]+/g,'_')}

async function loadPdf(file){
 return pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;
}
function groupItems(items){
 const rows=[];
 for(const item of items.filter(x=>x.str?.trim())){
   const x=item.transform?.[4]??0, y=item.transform?.[5]??0;
   let row=rows.find(r=>Math.abs(r.y-y)<4);
   if(!row){row={y,items:[]};rows.push(row)}
   row.items.push({x,text:item.str.trim()});
 }
 return rows.sort((a,b)=>b.y-a.y).map(r=>r.items.sort((a,b)=>a.x-b.x).map(i=>i.text));
}
function wordsToRows(words){
 const clean=words.filter(w=>w.text?.trim()&&Number(w.conf??100)>15);
 const rows=[];
 for(const w of clean){
   const y=(w.bbox?.y0??0)+(w.bbox?.y1??0)/2, x=w.bbox?.x0??0;
   let row=rows.find(r=>Math.abs(r.y-y)<Math.max(8,(w.bbox?.y1-w.bbox?.y0||12)*.65));
   if(!row){row={y,items:[]};rows.push(row)}
   row.items.push({x,text:w.text.trim()});
 }
 return rows.sort((a,b)=>a.y-b.y).map(r=>r.items.sort((a,b)=>a.x-b.x).map(i=>i.text));
}
function rowsToWorkbook(rows){
 const normalized=rows.filter(r=>r.some(v=>String(v).trim()));
 const max=Math.min(30,Math.max(1,...normalized.map(r=>r.length)));
 const data=normalized.map(r=>{const a=r.slice(0,max); while(a.length<max)a.push(''); return a});
 const ws=XLSX.utils.aoa_to_sheet(data.length?data:[['لا توجد بيانات']]);
 ws['!cols']=Array.from({length:max},(_,c)=>({wch:Math.min(42,Math.max(12,...data.map(r=>String(r[c]??'').length+2)))}));
 const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Extracted');
 return wb;
}
async function renderPage(page,scale=2){
 const viewport=page.getViewport({scale}); const canvas=document.createElement('canvas');
 canvas.width=Math.ceil(viewport.width); canvas.height=Math.ceil(viewport.height);
 await page.render({canvasContext:canvas.getContext('2d'),viewport}).promise;
 return canvas;
}
async function pdfRows(file,onProgress){
 const pdf=await loadPdf(file); let all=[]; let digital=0;
 for(let p=1;p<=pdf.numPages;p++){
   const page=await pdf.getPage(p); const text=await page.getTextContent();
   const rows=groupItems(text.items);
   if(rows.length){digital++; all.push(...rows)}
   onProgress?.(Math.round(p/pdf.numPages*35),`قراءة الصفحة ${p} من ${pdf.numPages}`);
 }
 if(digital>=Math.max(1,Math.ceil(pdf.numPages*.6))) return {rows:all,mode:'text'};
 const worker=await createWorker('ara+eng',1,{logger:m=>onProgress?.(35+Math.round((m.progress||0)*55),m.status||'OCR')});
 all=[];
 for(let p=1;p<=pdf.numPages;p++){
   const page=await pdf.getPage(p); const canvas=await renderPage(page,2);
   const result=await worker.recognize(canvas,{}, {text:true,tsv:true});
   all.push(...wordsToRows(result.data.words||[]));
   onProgress?.(35+Math.round(p/pdf.numPages*55),`OCR الصفحة ${p} من ${pdf.numPages}`);
 }
 await worker.terminate();
 return {rows:all,mode:'ocr'};
}

async function mergePdfs(files){
 const out=await PDFDocument.create();
 for(const file of files){
   const src=await PDFDocument.load(await file.arrayBuffer());
   const pages=await out.copyPages(src,src.getPageIndices()); pages.forEach(p=>out.addPage(p));
 }
 return out.save();
}
async function splitPdf(file,start,end){
 const src=await PDFDocument.load(await file.arrayBuffer()), out=await PDFDocument.create();
 const from=Math.max(1,start), to=Math.min(src.getPageCount(),end);
 const pages=await out.copyPages(src,Array.from({length:Math.max(0,to-from+1)},(_,i)=>from-1+i)); pages.forEach(p=>out.addPage(p));
 return out.save();
}
async function imagesToPdf(files){
 const out=await PDFDocument.create();
 for(const file of files){
   const bytes=await file.arrayBuffer(); const isPng=file.type.includes('png');
   const img=isPng?await out.embedPng(bytes):await out.embedJpg(bytes);
   const page=out.addPage([img.width,img.height]); page.drawImage(img,{x:0,y:0,width:img.width,height:img.height});
 }
 return out.save();
}
async function excelRows(file){
 const data=await file.arrayBuffer(); const wb=XLSX.read(data,{type:'array'}); const ws=wb.Sheets[wb.SheetNames[0]];
 return XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
}

function ToolModal({tool,onClose}){
 const [files,setFiles]=useState([]),[busy,setBusy]=useState(false),[progress,setProgress]=useState(0),[message,setMessage]=useState(''),[resultRows,setResultRows]=useState(null),[split,setSplit]=useState({start:1,end:1});
 const inputRef=useRef(null);
 const pick=e=>setFiles(Array.from(e.target.files||[]));
 const run=async()=>{
   if(!files.length){setMessage('اختر ملفاً أولاً');return}
   setBusy(true);setMessage('');setResultRows(null);
   try{
    if(tool.id==='pdf-excel'){
      const r=await pdfRows(files[0],(p,m)=>{setProgress(p);setMessage(m)});
      setResultRows(r.rows);setMessage(r.mode==='ocr'?'تم استخدام OCR لأن الملف ممسوح ضوئياً. راجع الجدول قبل التصدير.':'تم استخراج النص والجداول من PDF.');
    } else if(tool.id==='merge-pdf'){
      const b=await mergePdfs(files);downloadBytes(b,'merged.pdf','application/pdf');setMessage('تم دمج الملفات وتنزيل الملف.');
    } else if(tool.id==='split-pdf'){
      const b=await splitPdf(files[0],Number(split.start),Number(split.end));downloadBytes(b,`${safeName(files[0].name)}_pages_${split.start}-${split.end}.pdf`,'application/pdf');setMessage('تم تقسيم الملف وتنزيل النتيجة.');
    } else if(tool.id==='pdf-images'){
      const pdf=await loadPdf(files[0]);
      for(let p=1;p<=pdf.numPages;p++){const canvas=await renderPage(await pdf.getPage(p),2);const blob=await new Promise(r=>canvas.toBlob(r,'image/png'));downloadBlob(blob,`${safeName(files[0].name)}_page_${p}.png`);setProgress(Math.round(p/pdf.numPages*100))}
      setMessage('تم تحويل الصفحات إلى PNG.');
    } else if(tool.id==='pdf-word'){
      const pdf=await loadPdf(files[0]); const paras=[];
      for(let p=1;p<=pdf.numPages;p++){const page=await pdf.getPage(p);const text=await page.getTextContent();const rows=groupItems(text.items);rows.forEach(r=>paras.push(new Paragraph({children:[new TextRun(r.join(' '))]})));setProgress(Math.round(p/pdf.numPages*100))}
      if(!paras.length) throw new Error('هذا PDF يبدو ممسوحاً ضوئياً. استخدم OCR أو PDF إلى Excel.');
      const doc=new WordDocument({sections:[{children:paras}]});const blob=await Packer.toBlob(doc);downloadBlob(blob,`${safeName(files[0].name)}.docx`);setMessage('تم إنشاء ملف Word.');
    } else if(tool.id==='ocr'){
      let text='';
      if(files[0].type==='application/pdf'){const r=await pdfRows(files[0],(p,m)=>{setProgress(p);setMessage(m)});text=r.rows.map(r=>r.join(' ')).join('\n')}
      else {const worker=await createWorker('ara+eng',1,{logger:m=>{setProgress(Math.round((m.progress||0)*100));setMessage(m.status||'OCR')}});const r=await worker.recognize(files[0]);text=r.data.text;await worker.terminate()}
      const blob=new Blob([text],{type:'text/plain;charset=utf-8'});downloadBlob(blob,`${safeName(files[0].name)}_OCR.txt`);setMessage('تم استخراج النص وتنزيله.');
    } else if(tool.id==='image-pdf'){
      const b=await imagesToPdf(files);downloadBytes(b,'images.pdf','application/pdf');setMessage('تم إنشاء PDF من الصور.');
    } else if(tool.id==='excel-clean'){
      const rows=await excelRows(files[0]);const seen=new Set();const cleaned=rows.filter(r=>r.some(v=>String(v).trim())).filter(r=>{const k=JSON.stringify(r);if(seen.has(k))return false;seen.add(k);return true});
      setResultRows(cleaned);setMessage(`تم تنظيف البيانات: ${rows.length-cleaned.length} صف تم حذفه.`);
    } else if(tool.id==='excel-csv'){
      if(/\.csv$/i.test(files[0].name)){const text=await files[0].text();const wb=XLSX.read(text,{type:'string'});XLSX.writeFile(wb,`${safeName(files[0].name)}.xlsx`);setMessage('تم تحويل CSV إلى Excel.')}
      else {const wb=XLSX.read(await files[0].arrayBuffer(),{type:'array'});const ws=wb.Sheets[wb.SheetNames[0]];const csv=XLSX.utils.sheet_to_csv(ws);downloadBlob(new Blob([csv],{type:'text/csv;charset=utf-8'}),`${safeName(files[0].name)}.csv`);setMessage('تم تحويل Excel إلى CSV.')}
    }
   }catch(e){setMessage('حدث خطأ: '+(e?.message||'تعذر تنفيذ الأداة'))}
   finally{setBusy(false)}
 };
 const exportRows=()=>{if(!resultRows?.length)return;const wb=rowsToWorkbook(resultRows);XLSX.writeFile(wb,`${safeName(files[0]?.name||'result')}_processed.xlsx`)};
 return <div className="overlay" onMouseDown={e=>e.target===e.currentTarget&&onClose()}>
  <div className="modal">
   <div className="modalhead"><div><span className="modalicon"><tool.icon size={22}/></span><div><h3>{tool.title}</h3><p>{tool.desc}</p></div></div><button className="close" onClick={onClose}><X/></button></div>
   <label className="filebox"><input ref={inputRef} type="file" multiple={['merge-pdf','image-pdf'].includes(tool.id)} accept={tool.accept} onChange={pick}/><Upload size={28}/><strong>{files.length?files.map(f=>f.name).join(' • '):'اسحب الملفات هنا أو اضغط للاختيار'}</strong><span>{tool.accept}</span></label>
   {tool.id==='split-pdf'&&<div className="range"><label>من صفحة <input type="number" min="1" value={split.start} onChange={e=>setSplit({...split,start:e.target.value})}/></label><label>إلى صفحة <input type="number" min="1" value={split.end} onChange={e=>setSplit({...split,end:e.target.value})}/></label></div>}
   {busy&&<div className="progress"><div><span>{message||'جاري المعالجة...'}</span><b>{progress}%</b></div><i><em style={{width:`${progress}%`}}/></i></div>}
   {message&&!busy&&<div className="notice">{message}</div>}
   {resultRows&&<div className="preview"><div className="previewhead"><strong><Table2 size={18}/> معاينة النتيجة</strong><button onClick={exportRows}><Download size={17}/> تصدير Excel</button></div><div className="tablewrap"><table><tbody>{resultRows.slice(0,80).map((r,i)=><tr key={i}>{r.map((v,j)=><td key={j}>{String(v)}</td>)}</tr>)}</tbody></table></div><small>المعاينة تعرض أول 80 صفاً فقط، والتصدير يشمل كل الصفوف.</small></div>}
   <div className="modalactions"><button className="secondary" onClick={onClose}>إغلاق</button><button className="primary" disabled={busy} onClick={run}>{busy?<LoaderCircle className="spin"/>:<Play size={17}/>} تنفيذ الأداة</button></div>
  </div>
 </div>
}

function App(){
 const [dark,setDark]=useState(false),[query,setQuery]=useState(''),[cat,setCat]=useState('الكل'),[selected,setSelected]=useState(null),[drag,setDrag]=useState(false);
 const cats=['الكل','PDF','Excel','صور','OCR'];
 const filtered=useMemo(()=>tools.filter(t=>(cat==='الكل'||t.cat===cat)&&((t.title+' '+t.desc).includes(query))),[cat,query]);
 const handleDrop=e=>{e.preventDefault();setDrag(false);const file=e.dataTransfer.files?.[0];if(!file)return;const match=tools.find(t=>(file.type==='application/pdf'&&t.id==='pdf-excel')||(file.type.includes('spreadsheet')&&t.id==='excel-clean')||(file.type.startsWith('image/')&&t.id==='ocr'));setSelected(match||tools[0])};
 return <div className={dark?'app dark':'app'}>
  <header><div className="brand"><div className="logo">أ</div><div><h1>أدوات المحاسب</h1><p>أدوات عملية تنجز الشغل</p></div></div><button className="theme" onClick={()=>setDark(!dark)}>{dark?<Sun size={19}/>:<Moon size={19}/>}</button></header>
  <main>
   <section className="hero"><span className="eyebrow"><ShieldCheck size={16}/> معالجة محلية وبدون قاعدة بيانات</span><h2>كل أدوات شغلك<br/><span>في مكان واحد.</span></h2><p>PDF وExcel والصور وOCR — ارفع الملف، نفّذ العملية، ونزّل النتيجة مباشرة.</p><div className="search"><Search size={20}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="ابحث عن أداة..."/></div></section>
   <div className="cats">{cats.map(c=><button className={cat===c?'active':''} onClick={()=>setCat(c)} key={c}>{c}</button>)}</div>
   <section className="grid">{filtered.map(t=>{const I=t.icon;return <button className={'tool '+(t.featured?'featured':'')} key={t.id} onClick={()=>setSelected(t)}><div className="tooltop"><span className="icon"><I size={22}/></span><span className="go">←</span></div><h3>{t.title}</h3><p>{t.desc}</p>{t.featured&&<span className="badge">الأداة الأساسية</span>}</button>})}</section>
   <section className={'drop '+(drag?'drag':'')} onDragOver={e=>{e.preventDefault();setDrag(true)}} onDragLeave={()=>setDrag(false)} onDrop={handleDrop}><Upload size={24}/><div><strong>اسحب ملفك هنا</strong><span>وسنقترح لك الأداة المناسبة حسب نوع الملف</span></div><button onClick={()=>document.getElementById('globalPick').click()}>اختيار ملف</button><input id="globalPick" hidden type="file" onChange={e=>{const f=e.target.files?.[0];if(f){const m=tools.find(t=>(f.type==='application/pdf'&&t.id==='pdf-excel')||(f.type.startsWith('image/')&&t.id==='ocr'));setSelected(m||tools[0])}}}/></section>
   <section className="privacy"><ShieldCheck size={20}/><div><strong>خصوصيتك أولاً</strong><span>المعالجة تتم داخل متصفحك، ولا توجد قاعدة بيانات أو حسابات أو حفظ دائم لملفاتك.</span></div></section>
  </main>
  <footer><div>أدوات المحاسب <span>•</span> أدوات عملية للمحاسب والعمل المكتبي</div><div className="watermark">{OWNER} <b>•</b> {PHONE}</div></footer>
  {selected&&<ToolModal tool={selected} onClose={()=>setSelected(null)}/>}
 </div>
}
createRoot(document.getElementById('root')).render(<App/>);
