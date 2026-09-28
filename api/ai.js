export default async function handler(req,res){
  if(req.method!=='POST') return res.status(405).json({error:'Method not allowed'});
  const key=process.env.GEMINI_API_KEY;
  if(!key) return res.status(503).json({error:'AI backend is not configured'});
  try{
    const {task='document_ocr',rows=[]}=req.body||{};
    if(!Array.isArray(rows)||rows.length>1200) return res.status(400).json({error:'Invalid rows'});
    const rules={
      pdf_to_word_ocr_correction:'Correct OCR errors using surrounding context. Preserve every number, date, amount, name, code and punctuation unless strongly supported. Never invent missing text. Keep exactly the same row and cell counts.',
      pdf_table_reconstruction:'Reconstruct PDF table cells from OCR/text. Preserve financial numbers and dates exactly when readable. Fix obvious OCR mistakes only when context proves the correction. Keep the same row/cell counts.',
      bank_statement_reconstruction:'Act as a meticulous accounting-document OCR reviewer. Correct obvious OCR mistakes in dates, descriptions and numeric amounts, but never invent transactions or alter a number without strong contextual evidence. Preserve opening/closing balances and transaction values. Keep the same row/cell counts.'
    };
    const instructions=rules[task]||rules.pdf_to_word_ocr_correction;
    const prompt='You are a precision accounting-document reconstruction engine. Return ONLY JSON matching the requested structure. '+instructions+'\\nInput rows:\\n'+JSON.stringify(rows);
    const payload={
      contents:[{role:'user',parts:[{text:prompt}]}],
      generationConfig:{
        responseMimeType:'application/json',
        responseSchema:{
          type:'OBJECT',
          properties:{
            rows:{
              type:'ARRAY',
              items:{type:'ARRAY',items:{type:'STRING'}}
            }
          },
          required:['rows']
        }
      }
    };
    const model=process.env.GEMINI_MODEL||'gemini-2.5-flash-lite';
    const url='https://generativelanguage.googleapis.com/v1beta/models/'+encodeURIComponent(model)+':generateContent';
    const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json','x-goog-api-key':key},body:JSON.stringify(payload)});
    if(!r.ok){const t=await r.text();return res.status(502).json({error:'AI provider error',detail:t.slice(0,500)})}
    const data=await r.json();
    const raw=data?.candidates?.[0]?.content?.parts?.map(x=>x.text||'').join('')||'';
    if(!raw) throw Error('AI returned an empty response');
    const parsed=JSON.parse(raw);
    if(!Array.isArray(parsed.rows)||parsed.rows.length!==rows.length) throw Error('AI returned an invalid structure');
    const safe=parsed.rows.map((r,i)=>Array.isArray(r)&&r.length===rows[i].length?r:rows[i]);
    return res.status(200).json({rows:safe,confidence:'AI precision review'});
  }catch(e){return res.status(500).json({error:'AI processing failed',detail:String(e.message||e)})}
}
