// Reports are composed from computed evidence. No network or language-model API is used.
const n=v=>{try{return BigInt(v).toLocaleString('fa-IR');}catch{return String(v??'');}};
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export const nextAction = rule => ({
 balance:'ریز گردش و مانده ابتدا و انتهای دوره را دریافت و ثبت‌های اصلاحی را تطبیق دهید.',
 subtotal:'معین‌های سرفصل و محدوده جمع در گزارش مبدأ را کنترل کنید.',
 grand:'کامل بودن سرفصل‌ها، فیلترهای گزارش و محدوده جمع کل را کنترل کنید.',
 nature:'ماهیت مصوب سرفصل و آخرین ثبت‌های بدهکار و بستانکار را با مسئول حساب تطبیق دهید.',
 count:'تعریف تعداد در خروجی سامانه و ریز ثبت‌های خودکار و اصلاحی را از بانک دریافت کنید.',
 turnover:'اصلاحات، برگشت اسناد و مبنای گردش تجمعی را با ریز سند بررسی کنید.',
 quality:'نسخه کامل و معتبر خروجی را دریافت کنید؛ فرمول‌های بدون نتیجه را در اکسل محاسبه و ذخیره کنید.',
 duplicate:'شناسه‌های تکراری را با دفتر معین تطبیق دهید؛ پیش از جمع‌زدن، تکراری بودن رکورد را تأیید کنید.',
 suspense:'ریز اقلام باز، تاریخ ایجاد، آخرین پیگیری و مسئول تعیین تکلیف را دریافت کنید.',
 cross:'دوره و واحد دو حساب را تأیید و اقلام تشکیل‌دهنده اختلاف خالص را به‌صورت ردیفی تطبیق دهید.',
 commitment:'نگاشت حساب‌های مقابل را تأیید و فهرست تعهدات و قراردادهای مرتبط را تطبیق دهید.',
 watch:'ریز تسهیلات، تاریخ سررسید، وصول، طبقه‌بندی و وضعیت وثایق را برای پیگیری دریافت کنید.'
}[rule]??'ریز مستندات و توضیح مسئول حساب را دریافت کنید.');

export function reportBlocks(report,settings={}){
 const f=report.findings,s=report.summary,primary=f.filter(x=>x.rule!=='watch'),watch=f.filter(x=>x.rule==='watch');
 const meta=report.sheets.filter(s=>s.supported).map(s=>s.metadata).join(' | ');
 const blocks=[
  {type:'kicker',text:'BSI / BANK INSPECTION / 02'},
  {type:'title',text:'گزارش بررسی اسناد بانکی'},
  {type:'subtitle',text:report.name},
  {type:'paragraph',text:meta||'مشخصات دوره در منبع موجود نیست.'},
  {type:'paragraph',text:`تهیه‌کننده: ${settings.inspector?.trim()||'ثبت نشده'}   |   تاریخ تهیه: ${new Date(report.generatedAt).toLocaleDateString('fa-IR')}   |   نسخه قواعد: ${report.version??'2.0'}`},
  {type:'notice',text:'محرمانه - برای بررسی داخلی. این گزارش ماشینی، پیش‌نویس مستند برای بازرس است و تأیید یا اثبات تخلف محسوب نمی‌شود.'},
  {type:'heading',text:'۱. خلاصه مدیریتی'},
  {type:'paragraph',text:`در فایل حاضر ${n(s.records)} حساب معین و ${n(s.subtotals)} جمع سرفصل بررسی شد. ${n(report.checks.filter(c=>c.status==='pass').length)} کنترل جمع منطبق، ${n(report.checks.filter(c=>c.status==='fail').length)} کنترل نامنطبق و ${n(report.checks.filter(c=>c.status==='unavailable').length)} کنترل غیرقابل انجام بود. ${n(s.confirmed??0)} یافته مغایرت محاسباتی ثبت شد. این نتیجه فقط به کنترل‌های اجراشده مربوط است.`},
  {type:'paragraph',text:`${n(primary.length)} هشدار یا یافته بررسی در ${n(s.affected)} ردیف و ${n(watch.length)} حساب مطالبات غیرجاری برای پیگیری فهرست شده است. تعداد یافته با تعداد حساب یکسان نیست؛ یک حساب می‌تواند چند هشدار داشته باشد. اولویت، پیشنهاد ترتیب بررسی است و برآورد زیان یا شدت تخلف نیست.`},
  {type:'paragraph',text:`حد بررسی انتخابی: ${n(report.threshold)} واحد فایل. واحد پول از عنوان حساب‌ها به‌طور یکسان قابل اثبات نیست؛ مقادیر بدون تبدیل و تجمیع ارزهای متفاوت گزارش شده‌اند. جمع مطلق مبالغ هشدارها به‌عنوان زیان یا جمع مغایرت ارائه نشده است.`},
  {type:'heading',text:'۲. نکات قابل اقدام'},
 ];
 const counts=new Map();for(const x of primary)counts.set(x.rule,(counts.get(x.rule)??0)+1);
 for(const rule of ['balance','subtotal','grand','quality','duplicate','nature','cross','suspense','count','turnover','commitment'])if(counts.has(rule))blocks.push({type:'action',text:`${n(counts.get(rule))} مورد - ${nextAction(rule)}`});
 if(!primary.length)blocks.push({type:'paragraph',text:'در کنترل‌های اجراشده هشدار بررسی ثبت نشد؛ تکمیل بررسی به دسترسی به ریز اسناد وابسته است.'});
 if(report.crossChecks?.length){blocks.push({type:'heading',text:'۳. تطبیق حساب‌های مقابل'});for(const c of report.crossChecks)blocks.push({type:c.status==='review'?'notice':'paragraph',text:`${c.title}: ${c.status==='pass'?'منطبق':c.status==='review'?'نیازمند تأیید بازرس':'بررسی نشد'}. شیت ${c.sheet}، ردیف‌های ${c.rows.map(n).join('، ')}${c.net!==undefined?`، خالص ${n(c.net)} واحد فایل`:''}. ${c.basis}`});}
 blocks.push({type:'heading',text:'۴. شواهد یافته‌ها و اقدام پیشنهادی'});
 for(const [i,x]of primary.entries()){
  blocks.push({type:'finding',text:`${n(i+1)}. ${x.title}`,severity:x.severity});
  blocks.push({type:'source',text:`${x.account} | شیت ${x.sheet} | ردیف ${n(x.row)} | کد ${x.code||'ناموجود'} | سلول‌ها ${x.cells?.join('، ')||'ساختار گزارش'}`});
  blocks.push({type:'paragraph',text:`نوع نتیجه: ${x.classification}. ${x.detail}`});
  blocks.push({type:'action',text:`اقدام پیشنهادی: ${nextAction(x.rule)}`});
 }
 blocks.push({type:'heading',text:'۵. پیوست مطالبات غیرجاری'});
 blocks.push({type:'paragraph',text:'این فهرست از عنوان حساب و سرفصل مادر استخراج شده است. مانده غیرجاری، به‌خودی‌خود تخلف یا مبلغ زیان نیست و برای احتساب ذخیره یا نرخ مطالبات به اطلاعات تکمیلی نیاز است.'});
 for(const x of watch)blocks.push({type:'appendix',text:`ردیف ${n(x.row)}، شیت ${x.sheet}، کد ${x.code}: ${x.account}؛ مانده ${n(x.amount)} واحد فایل.`});
 if(!watch.length)blocks.push({type:'paragraph',text:'حسابی با عنوان قابل شناسایی و مانده غیرصفر در این دسته ثبت نشد.'});
 blocks.push({type:'heading',text:'۶. دامنه و محدودیت بررسی'});
 blocks.push({type:'paragraph',text:'کنترل‌ها شامل مانده و گردش، جمع سرفصل و کل، ماهیت، تعداد و مبلغ، کیفیت داده و تطبیق‌های تعریف‌شده است. فرمول‌های اکسل بازحساب نشده‌اند. نبود ریز تراکنش، اطلاعات مشتری، دوره‌های مقایسه‌ای، مانده ابتدای دوره و قدمت اقلام، مانع اظهار نظر درباره همه تخلف‌ها و ریسک‌ها است.'});
 for(const sh of report.sheets)blocks.push({type:'paragraph',text:`شیت ${sh.name}: ${sh.supported?`${n(sh.records)} معین بررسی شد.`:`بررسی نشده؛ ${sh.reason}`}`});
 if(settings.note?.trim())blocks.push({type:'heading',text:'یادداشت بازرس (افزوده کاربر)'},{type:'paragraph',text:settings.note.trim()});
 blocks.push({type:'source',text:`شناسه منبع SHA-256: ${report.sha256||'موجود نیست'}`},{type:'source',text:'منبع همه ارقام: فایل اکسل انتخاب‌شده. شماره ردیف‌ها، شماره واقعی ردیف در اکسل هستند. نسخه نهایی نیازمند بازبینی و امضای بازرس است.'});
 return blocks;
}
export function reportHTML(report,settings={}){
 return reportBlocks(report,settings).map(b=>`<${['title','heading'].includes(b.type)?'h2':'p'} class="report-${b.type}${b.severity?' '+b.severity:''}">${esc(b.text)}</${['title','heading'].includes(b.type)?'h2':'p'}>`).join('');
}

// Offline PDF export uses the browser's Persian text shaping. Raster pages preserve glyphs
// consistently across PDF viewers. Use the separate print action for selectable text.
export async function createReportPDF(report,settings={},onProgress=()=>{}){
 const W=1240,H=1754,margin=88,maxW=W-2*margin;
 const pages=[];let canvas,ctx,y,page=0;
 async function finish(){if(!canvas)return;ctx.fillStyle='#72829c';ctx.font='19px Tahoma';ctx.direction='rtl';ctx.textAlign='right';ctx.fillText('BSI  /  گزارش داخلی - نسخه قواعد ۲.۰',W-margin,H-48);ctx.textAlign='left';ctx.fillText(n(++page),margin,H-48);pages.push(new Uint8Array(await (await new Promise(r=>canvas.toBlob(r,'image/jpeg',.94))).arrayBuffer()));onProgress(page);await new Promise(r=>setTimeout(r,0));}
 function start(){canvas=document.createElement('canvas');canvas.width=W;canvas.height=H;ctx=canvas.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,W,H);ctx.fillStyle='#1947cf';ctx.fillRect(0,0,W,17);ctx.font='bold 21px Arial';ctx.textAlign='left';ctx.fillText('BSI',margin,60);ctx.fillStyle='#8190a6';ctx.font='17px Tahoma';ctx.textAlign='right';ctx.direction='rtl';ctx.fillText('گزارش بررسی اسناد بانکی',W-margin,60);ctx.strokeStyle='#e4e9f2';ctx.beginPath();ctx.moveTo(margin,82);ctx.lineTo(W-margin,82);ctx.stroke();y=122;}
 function wrap(text){const out=[];for(const line of String(text).split('\n')){let current='';for(const word of line.split(/\s+/)){const next=current?current+' '+word:word;if(ctx.measureText(next).width<=maxW-20){current=next;continue;}if(current){out.push(current);current='';}if(ctx.measureText(word).width<=maxW-20)current=word;else{for(const char of word){if(ctx.measureText(current+char).width>maxW-20){out.push(current);current='';}current+=char;}}}if(current)out.push(current);}return out;}
 start();
 for(const b of reportBlocks(report,settings)){
  const large=b.type==='title',heading=b.type==='heading',finding=b.type==='finding';
  const size=large?44:heading?29:finding?25:b.type==='kicker'?17:b.type==='source'?19:22;
  const lineHeight=Math.ceil(size*1.85),gap=heading?20:large?18:10;
  const font=`${large||heading||finding?'bold ':''}${size}px Tahoma`;
  ctx.font=font;const lines=wrap(b.text);const keep=(heading||finding)?Math.min(lines.length*lineHeight+100,300):Math.min(lines.length*lineHeight+gap,260);
  if(y+keep>H-100){await finish();start();}
  for(const line of lines){
   if(y+lineHeight>H-100){await finish();start();}
   ctx.font=font;ctx.direction='rtl';ctx.textAlign='right';
   if(['notice','action','finding'].includes(b.type)){ctx.fillStyle=b.type==='notice'?'#fff5dd':b.type==='action'?'#f2f6ff':'#edf2ff';ctx.fillRect(margin,y-8,maxW,lineHeight);ctx.fillStyle=b.type==='notice'?'#e1a53a':'#4e75dd';ctx.fillRect(W-margin-4,y-8,4,lineHeight);}
   ctx.fillStyle=large||heading?'#173882':b.type==='source'?'#6a7991':'#263954';ctx.fillText(line,W-margin-14,y+size);y+=lineHeight;
  }y+=gap;
 }
 await finish();return jpegPDF(pages,W,H);
}
function jpegPDF(images,width,height){
 const enc=new TextEncoder(),chunks=[],offsets=[0];let offset=0;
 const write=v=>{const bytes=typeof v==='string'?enc.encode(v):v;chunks.push(bytes);offset+=bytes.length;};
 const object=(id,body)=>{offsets[id]=offset;write(`${id} 0 obj\n`);write(body);write('\nendobj\n');};
 write('%PDF-1.4\n%BSI\n');object(1,'<< /Type /Catalog /Pages 2 0 R >>');object(2,`<< /Type /Pages /Count ${images.length} /Kids [${images.map((_,i)=>`${5+i*3} 0 R`).join(' ')}] >>`);
 images.forEach((image,i)=>{const imageId=3+i*3,contentId=imageId+1,pageId=imageId+2;offsets[imageId]=offset;write(`${imageId} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${image.length} >>\nstream\n`);write(image);write('\nendstream\nendobj\n');const commands='q\n595.28 0 0 841.89 0 0 cm\n/Scan Do\nQ';object(contentId,`<< /Length ${enc.encode(commands).length} >>\nstream\n${commands}\nendstream`);object(pageId,`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.28 841.89] /Resources << /XObject << /Scan ${imageId} 0 R >> >> /Contents ${contentId} 0 R >>`);});
 const xref=offset;write(`xref\n0 ${offsets.length}\n0000000000 65535 f \n`);for(let i=1;i<offsets.length;i++)write(`${String(offsets[i]).padStart(10,'0')} 00000 n \n`);write(`trailer\n<< /Size ${offsets.length} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`);return new Blob(chunks,{type:'application/pdf'});
}
