import React,{useState} from 'react';
import {createRoot} from 'react-dom/client';
import {FileSpreadsheet,FileText,Image as ImageIcon,Search,Sun,Moon,ArrowLeft,Upload,ShieldCheck,WandSparkles} from 'lucide-react';
import './styles.css';

const tools=[
 {id:'pdf-excel',cat:'PDF',title:'PDF إلى Excel',desc:'حوّل ملفات PDF العادية والممسوحة ضوئياً إلى جداول Excel قابلة للتعديل.',icon:FileSpreadsheet,featured:true},
 {id:'merge-pdf',cat:'PDF',title:'دمج PDF',desc:'ادمج عدة ملفات PDF في ملف واحد مع ترتيب الصفحات.',icon:FileText},
 {id:'split-pdf',cat:'PDF',title:'تقسيم PDF',desc:'استخرج صفحات أو نطاقات محددة من ملف PDF.',icon:FileText},
 {id:'pdf-images',cat:'PDF',title:'PDF إلى صور',desc:'حوّل صفحات PDF إلى JPG أو PNG.',icon:ImageIcon},
 {id:'excel-clean',cat:'Excel',title:'تنظيف Excel',desc:'تنظيف الصفوف الفارغة والتكرارات وتوحيد البيانات.',icon:FileSpreadsheet},
 {id:'excel-csv',cat:'Excel',title:'Excel ↔ CSV',desc:'تحويل سريع بين ملفات Excel وCSV.',icon:FileSpreadsheet},
 {id:'image-pdf',cat:'صور',title:'صور إلى PDF',desc:'اجمع الصور في ملف PDF واحد.',icon:ImageIcon},
 {id:'ocr',cat:'OCR',title:'استخراج النص OCR',desc:'استخرج النص من الصور والملفات الممسوحة ضوئياً.',icon:WandSparkles}
];

function App(){
 const [dark,setDark]=useState(false),[query,setQuery]=useState(''),[cat,setCat]=useState('الكل');
 const cats=['الكل','PDF','Excel','صور','OCR'];
 const filtered=tools.filter(t=>(cat==='الكل'||t.cat===cat)&&((t.title+' '+t.desc).includes(query)));
 return <div className={dark?'app dark':'app'}>
  <header><div className="brand"><div className="logo">أ</div><div><h1>أدوات المحاسب</h1><p>أدوات عملية تنجز الشغل</p></div></div><button className="theme" onClick={()=>setDark(!dark)}>{dark?<Sun size={19}/>:<Moon size={19}/>}</button></header>
  <main>
   <section className="hero"><span className="eyebrow"><ShieldCheck size={16}/> ملفاتك على جهازك</span><h2>كل الأدوات التي تحتاجها<br/><span>في مكان واحد.</span></h2><p>حوّل، نظّف، استخرج ونظّم ملفاتك بسهولة. بدون حسابات وبدون قاعدة بيانات.</p>
   <div className="search"><Search size={20}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="ابحث عن أداة..."/></div></section>
   <div className="cats">{cats.map(c=><button className={cat===c?'active':''} onClick={()=>setCat(c)} key={c}>{c}</button>)}</div>
   <section className="grid">{filtered.map(t=>{const I=t.icon;return <button className={'tool '+(t.featured?'featured':'')} key={t.id} onClick={()=>alert('الأداة جاهزة للربط بالمعالجة: '+t.title)}><div className="tooltop"><span className="icon"><I size={22}/></span><ArrowLeft size={18}/></div><h3>{t.title}</h3><p>{t.desc}</p>{t.featured&&<span className="badge">أداة أساسية</span>}</button>})}</section>
   <section className="drop"><Upload size={22}/><div><strong>اسحب ملفك هنا</strong><span>وسنختار لك الأدوات المناسبة حسب نوعه</span></div><button>اختيار ملف</button></section>
  </main>
  <footer>أدوات المحاسب <span>•</span> معالجة محلية بدون تخزين دائم للملفات</footer>
 </div>
}
createRoot(document.getElementById('root')).render(<App/>);