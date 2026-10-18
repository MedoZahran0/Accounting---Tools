export default async function handler(req,res){
  if(req.method!=='POST') return res.status(405).json({error:'Method not allowed'});
  const key=process.env.GEMINI_API_KEY;
  if(!key) return res.status(503).json({error:'AI backend is not configured'});
  try{
    const {task='document_ocr',rows=[],pdfBase64='',mimeType='application/pdf'}=req.body||{};
    if(!Array.isArray(rows)||rows.length>1200) return res.status(400).json({error:'Invalid rows'});
    const rules={
      image_ocr:'Correct OCR from the supplied image. Preserve Arabic word boundaries, spaces, punctuation, numbers and line breaks when readable. Keep Arabic in normal Unicode logical reading order. Do not concatenate separate words, reverse characters, transliterate or invent text.',
      pdf_to_word_ocr_correction:'Correct OCR errors using surrounding context. Preserve every number, date, amount, name, code and punctuation unless strongly supported. Never invent missing text. Keep exactly the same row and cell counts.',
      pdf_table_reconstruction:'Read the supplied PDF visually and reconstruct its tables accurately. Arabic text must remain correct Arabic Unicode in logical reading order. Preserve every date, amount, account number and balance exactly when readable. Do not reverse Arabic characters, do not transliterate Arabic, do not invent values. Preserve spaces between Arabic words exactly; never concatenate separate Arabic words into one string. Return rows in normal table reading order.',
      bank_statement_reconstruction:'Extract ONLY the transaction table from the supplied bank statement PDF. Return one header row with EXACTLY these five cells in this exact order: date, description, debit, credit, balance. Then return one row per real transaction with exactly five cells: transaction date, full description, BANK-SIDE debit amount, BANK-SIDE credit amount, and the bank balance shown after that transaction if present. Keep debit and credit in the bank statement perspective; do NOT reverse them here because the application reverses them later. Ignore page headers, account numbers, branch numbers, reference numbers, page numbers, statement period, opening/closing summary boxes, totals, and footers. Preserve all Arabic words with spaces in logical reading order. Dates must be DD/MM/YYYY when readable. Amounts must be plain numeric strings with decimal point and no thousands separators; blank means the cell is truly blank. Do not treat dates, account/reference numbers, or balances as transaction amounts. Do not invent or interpolate missing values. If the PDF has multiple pages, combine all transaction rows in chronological order and do not repeat the header.'
    };
    const instructions=rules[task]||rules.pdf_to_word_ocr_correction;
    const parts=[];
    if(pdfBase64){
      const clean=String(pdfBase64).replace(/^data:[^;]+;base64,/,'');
      if(clean.length>70*1024*1024) return res.status(413).json({error:'PDF is too large for inline AI processing'});
      parts.push({inlineData:{mimeType:mimeType||'application/pdf',data:clean}});
    }
    const prompt='You are a precision accounting-document reconstruction engine. '+instructions+'\n'+(pdfBase64?'Use the PDF itself as the primary source. The following OCR rows are only a hint and may be wrong:\n'+JSON.stringify(rows):'Input rows:\n'+JSON.stringify(rows))+'\nReturn ONLY JSON with {"rows":[["cell",...],...]}. '+(task==='bank_statement_reconstruction'?'For this task, the first row MUST be exactly ["date","description","debit","credit","balance"], every later row must contain exactly five cells, and the number of rows MUST match the transactions actually visible in the PDF, not the OCR row count.':'Keep a consistent column count and preserve the original row count and cell count.');
    parts.push({text:prompt});
    const payload={contents:[{role:'user',parts}],generationConfig:{responseMimeType:'application/json',responseSchema:{type:'OBJECT',properties:{rows:{type:'ARRAY',items:{type:'ARRAY',items:{type:'STRING'}}}},required:['rows']}}};
    const model=process.env.GEMINI_MODEL||'gemini-2.5-flash-lite';
    const url='https://generativelanguage.googleapis.com/v1beta/models/'+encodeURIComponent(model)+':generateContent';
    const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json','x-goog-api-key':key},body:JSON.stringify(payload)});
    if(!r.ok){const t=await r.text();return res.status(502).json({error:'AI provider error',detail:t.slice(0,500)})}
    const data=await r.json(),raw=data?.candidates?.[0]?.content?.parts?.map(x=>x.text||'').join('')||'';
    if(!raw) throw Error('AI returned an empty response');
    const parsed=JSON.parse(raw);
    if(!Array.isArray(parsed.rows)) throw Error('AI returned an invalid structure');
    const safe=task==='bank_statement_reconstruction'?parsed.rows.filter(r=>Array.isArray(r)&&r.length===5).slice(0,2000):(rows.length?parsed.rows.map((r,i)=>Array.isArray(r)&&r.length===rows[i].length?r:rows[i]).slice(0,rows.length):parsed.rows);
    return res.status(200).json({rows:safe,confidence:'AI visual PDF review'});
  }catch(e){return res.status(500).json({error:'AI processing failed',detail:String(e.message||e)})}
}