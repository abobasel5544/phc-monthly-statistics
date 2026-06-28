'use client';
import {useMemo,useState,useEffect} from 'react';
import {modules} from '../lib/templates';
import {doc, getDoc, setDoc, addDoc, collection, getDocs, query, orderBy, limit, serverTimestamp, deleteDoc} from 'firebase/firestore';
import {db} from '../lib/firebase';

type Role='admin'|'executive'|'center';
const centers=[
 {id:'c01',name:'مركز صحي الواحة'},{id:'c02',name:'مركز صحي الرغامة'},{id:'c03',name:'مركز صحي قويزة'},{id:'c04',name:'مركز صحي المطار القديم'},{id:'c05',name:'مركز صحي الروابي'},{id:'c06',name:'مركز صحي السليمانية'},{id:'c07',name:'مركز صحي شرق الخط'}
];
const users=[
 {username:'admin',password:'admin123',name:'مدير النظام',role:'admin' as Role,centerId:'',isActive:true},
 {username:'executive',password:'123456',name:'إدارة شؤون المراكز الصحية',role:'executive' as Role,centerId:'',isActive:true},
 ...centers.map((c,i)=>({username:`center${i+1}`,password:'123456',name:c.name,role:'center' as Role,centerId:c.id,isActive:true})),
 ...centers.map((c)=>({username:c.id,password:'123456',name:c.name,role:'center' as Role,centerId:c.id,isActive:true}))
];
const hijriYears=Array.from({length:16},(_,i)=>String(1445+i));
function fileSafeName(v:string){return String(v||'').replace(/[\\/:*?"<>|]/g,'-').replace(/\s+/g,'_')}
function colName(n:number){let s='';while(n){let r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26)}return s}
function cellToParts(ref:string){const m=ref.match(/([A-Z]+)(\d+)/)!;let c=0;for(const ch of m[1]) c=c*26+ch.charCodeAt(0)-64;return {c,r:Number(m[2])}}
function normalizeNumberText(v:string){return String(v||'').replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d).toString()).replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d).toString())}
function cleanYearInput(v:string){return normalizeNumberText(v).replace(/[^0-9]/g,'').slice(0,4)}
function setWorkbookPeriodHeaders(wb:any, ws:any, m:any, selectedMonth:string, selectedYear:string){
 const y=Number(cleanYearInput(selectedYear));
 const mo=Number(normalizeNumberText(selectedMonth));
 if(y){
  // تحديث سنة النموذج في جميع مواضع رأس القالب حتى لا تبقى ثابتة على 1447.
  wb.worksheets.forEach((sheet:any)=>{sheet.eachRow((row:any)=>{row.eachCell({includeEmpty:false},(cell:any)=>{
   if(typeof cell.value==='number' && cell.value>=1400 && cell.value<=1500) cell.value=y;
   if(typeof cell.value==='string' && /^(14|15)\d{2}$/.test(cell.value.trim())) cell.value=String(y);
  })})});
 }
 // تثبيت رقم الشهر المحدد في رأس البلوك المصدر.
 const target=m.monthColumns?.find((x:any)=>String(x.month)===String(mo));
 if(target && mo){
  ws.getCell(`${colName(target.col)}${target.row}`).value=mo;
  if(y){
   // أغلب القوالب يكون ترتيب رأس الفترة: شهر، رقم الشهر، عام، السنة، هـ.
   // نكتب السنة في الخانة المجاورة المناسبة دون المساس بالمعادلات أو صفوف البيانات.
   ws.getCell(`${colName(target.col+2)}${target.row}`).value=y;
  }
 }
}
function key(centerId:string,year:string,month:string,moduleId:string,cell:string){return `stat:${centerId}:${year}:${month}:${moduleId}:${cell}`}
function firestoreFormDocId(centerId:string,year:string,month:string,moduleId:string){return `${centerId}_${year}_${month}_${moduleId}`.replace(/[^a-zA-Z0-9_-]/g,'_')}
function cleanLabel(label:string, cell:string){
 const raw=(label||'').trim();
 const looksLikeCell=/^[A-Z]+\d+$/.test(raw);
 const hasOldNumbers=/\d+\s*\/\s*\d+|^[\d\s:.-]+$/.test(raw);
 const cleaned=raw.replace(/\s*\([A-Z]+\d+\)\s*/g,'').trim();
 if(!cleaned || looksLikeCell || hasOldNumbers) return 'حقل إدخال إحصائي';
 return cleaned.replace(/^[0-9]+\s*\/\s*/g,'').replace(/\s*\/\s*[0-9]+$/g,'').trim();
}
const defaultFiveRowStarts:any={
 c07:5,   // الربيع والتوفيق / شرق الخط حسب القالب القديم
 c00:10,  // الجامعة - موجود في بعض قوالب Excel القديمة
 c06:15,  // السليمانية
 c04:20,  // المطار القديم
 c05:25,  // الروابي
 c03:30,  // قويزة
 c02:35,  // الرغامة
 c01:40   // الواحة
};
const configs:any={
 population:{starts:defaultFiveRowStarts,rowOffsets:[0,1,2,3],inputCols:[4,5,6,7,8,9,11],autoTotalCols:[10],totalRowOffset:4,rowCats:['سعودي / ذكر','سعودي / أنثى','غير سعودي / ذكر','غير سعودي / أنثى']},
 workforce:{starts:defaultFiveRowStarts,rowOffsets:[0,1,2,3],inputCols:[4,5,6,7,8,9,10,11,12,13,14,15,16],autoTotalCols:[17],totalRowOffset:4,rowCats:['سعودي / ذكر','سعودي / أنثى','غير سعودي / ذكر','غير سعودي / أنثى']},
 visitors:{starts:{c07:6,c00:13,c06:20,c04:27,c05:34,c03:41,c02:48,c01:55},rowOffsets:[0,1,3,4],inputCols:[4,5,6,7,8,10,11,12],autoTotalCols:[9,13],rowCats:['سعودي / ذكر','سعودي / أنثى','غير سعودي / ذكر','غير سعودي / أنثى']},
 chronic:{starts:defaultFiveRowStarts,rowOffsets:[0,1,2,3],inputCols:[4,5,6,7,8,9,10,11,12,13,14,15],rowCats:['سعودي / ذكر','سعودي / أنثى','غير سعودي / ذكر','غير سعودي / أنثى']},
 referrals:{starts:defaultFiveRowStarts,rowOffsets:[0,1,2,3],inputCols:[4,5,6,8],autoTotalCols:[7],totalRowOffset:4,rowCats:['سعودي / ذكر','سعودي / أنثى','غير سعودي / ذكر','غير سعودي / أنثى']},
 childcare:{starts:{c07:5,c00:11,c06:17,c04:23,c05:29,c03:35,c02:41,c01:47},rowOffsets:[0,1,2,3],inputCols:[4,5,6,7,8,9,10,11,12,13,14],totalRowOffset:4,rowCats:['سعودي / ذكر','سعودي / أنثى','غير سعودي / ذكر','غير سعودي / أنثى']},
 lab:{starts:{c07:6,c00:13,c06:20,c04:27,c05:34,c03:41,c02:48,c01:55},rowOffsets:[0,1,3,4],inputCols:[6,7,8,9,10,11,12,13,14,15,16],autoTotalCols:[17],rowCats:['سعودي / ذكر','سعودي / أنثى','غير سعودي / ذكر','غير سعودي / أنثى']},
 minor_surgery:{starts:defaultFiveRowStarts,rowOffsets:[0,1,2,3],inputCols:[4,5,6,7],rowCats:['سعودي / ذكر','سعودي / أنثى','غير سعودي / ذكر','غير سعودي / أنثى']},
 positive:{starts:{c07:6,c00:13,c06:20,c04:27,c05:34,c03:41,c02:48,c01:55},rowOffsets:[0,1,3,4],inputCols:[4,5,6,7,8,9],rowCats:['سعودي / ذكر','سعودي / أنثى','غير سعودي / ذكر','غير سعودي / أنثى']},
 xray:{starts:{c07:6,c00:13,c06:20,c04:27,c05:34,c03:41,c02:48,c01:55},rowOffsets:[0,1,3,4],inputCols:[4,5],rowCats:['سعودي / ذكر','سعودي / أنثى','غير سعودي / ذكر','غير سعودي / أنثى']},
 maternity:{starts:{c07:5,c00:9,c06:13,c04:17,c05:21,c03:25,c02:29,c01:33},rowOffsets:[0,1],inputCols:[3,4,5,6,8],rowCats:['سعودي','غير سعودي']},
 pregnancy_end:{starts:{c07:5,c00:16,c06:27,c04:38,c05:49,c03:60,c02:71,c01:82},rowOffsets:[0,1,6,7],inputCols:[3,4,5,6,7,8,9,10,11,12,13,14],rowCats:['سعوديات','غير سعوديات','سعوديات - مكان/نتيجة الولادة','غير سعوديات - مكان/نتيجة الولادة']},
 dental:{starts:{c07:2,c00:28,c06:54,c04:80,c05:106,c03:132,c02:158,c01:184},rowOffsets:[],inputCols:[],autoTotalCols:[11],rowCats:[]},
 common:{starts:{c07:6,c00:6,c06:6,c04:6,c05:6,c03:6,c02:6,c01:6},rowOffsets:Array.from({length:60},(_,i)=>i),inputCols:[2,3,4,5,7,8,9,10,11,12],autoTotalCols:[],rowCats:[]}
};

const populationAgeColumns:any={
 4:'أقل من سنة',
 5:'من 4:1',
 6:'من14:5',
 7:'من44:15',
 8:'من60:45',
 9:'أكثر من 60'
};
const populationRows:any=[
 {offset:0,nationality:'سعودي',gender:'ذكر'},
 {offset:1,nationality:'سعودي',gender:'أنثى'},
 {offset:2,nationality:'غير سعودي',gender:'ذكر'},
 {offset:3,nationality:'غير سعودي',gender:'أنثى'}
];

const workforceColumns:any={
 4:'أطباء — عام',
 5:'أطباء — ط.أسرة',
 6:'طبيب أسنان',
 7:'تمريض',
 8:'قابلة',
 9:'فني صيدلة',
 10:'فني مختبر',
 11:'فني أشعة',
 12:'مراقب صحي',
 13:'فني إحصاء',
 14:'إداري',
 15:'مستخدم',
 16:'أخرى'
};
const visitorsColumns:any={
 4:'عدد مراجعي العيادات العامة',
 5:'مراجعي الأمراض المزمنة',
 6:'عدد مراجعي عيادة الأسنان',
 7:'عدد زيارات الحوامل',
 8:'الطفل السليم والتطعيمات',
 10:'الضماد والحقن والطوارئ',
 11:'المختبر',
 12:'الأشعة'
};

const childcareColumns:any={
 4:'عدد الأطفال المسجلين أقل من عام',
 5:'التحصينات الأساسية — أكملوا التحصينات الأساسية خلال الشهر',
 6:'التحصينات الأساسية — بلغ عمرهم 15 شهراً ولم يكملوا التحصينات الأساسية',
 7:'عدد الأطفال المسجلين أقل من 5 أعوام',
 8:'عدد من رضعوا طبيعياً',
 9:'عدد المعرضين للخطر أقل من 5 سنوات',
 10:'زيارات عيادة الطفل السليم خلال الشهر',
 11:'وفيات الأطفال — من الولادة إلى أقل من 7 أيام',
 12:'وفيات الأطفال — من 7 أيام إلى أقل من شهر',
 13:'وفيات الأطفال — من شهر إلى أقل من عام',
 14:'وفيات الأطفال — من عام إلى أقل من 5 سنوات'
};
const childcareRows:any=[
 {offset:0,nationality:'سعودي',gender:'ذكر'},
 {offset:1,nationality:'سعودي',gender:'أنثى'},
 {offset:2,nationality:'غير سعودي',gender:'ذكر'},
 {offset:3,nationality:'غير سعودي',gender:'أنثى'}
];
function getChildcareInputFields(selectedCenterId:string){
 const cfg=configs.childcare;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 23;
 const list:any[]=[];
 for(const rowInfo of childcareRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(childcareColumns).map(Number)){
   const cell=`${colName(col)}${row}`;
   list.push({cell,row,col,label:`${childcareColumns[col]} — ${rowInfo.nationality} — ${rowInfo.gender}`});
  }
 }
 // هذا البند موجود أسفل صف المجموع في القالب الأصلي، وليس إجمالياً، لذلك يظهر كإدخال مستقل.
 list.push({cell:`L${startRow+5}`,row:startRow+5,col:12,label:'عدد الأطفال الذين يقل وزنهم بالنسبة لعمرهم تحت المنطقة C في منحنى النمو خلال الشهر — ذكر'});
 list.push({cell:`N${startRow+5}`,row:startRow+5,col:14,label:'عدد الأطفال الذين يقل وزنهم بالنسبة لعمرهم تحت المنطقة C في منحنى النمو خلال الشهر — أنثى'});
 return list;
}


const labColumns:any={
 6:'مصلي — فحص الإيدز HIV',
 7:'مصلي — فصائل الدم B.GROUP',
 8:'مصلي — التهاب كبدي HEP',
 9:'مصلي — فحص الحمل في الدم',
 10:'مصلي — أخرى',
 11:'كيماوي — نسبة السكر في الدم',
 12:'كيماوي — أخرى',
 13:'بول — روتين',
 14:'بول — حمل',
 15:'براز',
 16:'بصاق'
};
const labRows:any=[
 {offset:0,nationality:'سعودي',gender:'ذكر'},
 {offset:1,nationality:'سعودي',gender:'أنثى'},
 {offset:3,nationality:'غير سعودي',gender:'ذكر'},
 {offset:4,nationality:'غير سعودي',gender:'أنثى'}
];
function getLabInputFields(selectedCenterId:string){
 const cfg=configs.lab;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 27;
 const list:any[]=[];
 for(const rowInfo of labRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(labColumns).map(Number)){
   list.push({cell:`${colName(col)}${row}`,row,col,label:`${labColumns[col]} — ${rowInfo.nationality} — ${rowInfo.gender}`});
  }
 }
 return list;
}


const positiveColumns:any={
 4:'بلهارسيا بولية',
 5:'بلهارسيا معوية',
 6:'دوسنتاريا',
 7:'ملاريا',
 8:'حمل',
 9:'أخرى'
};
const positiveRows:any=[
 {offset:0,nationality:'سعودي',gender:'ذكر'},
 {offset:1,nationality:'سعودي',gender:'أنثى'},
 {offset:3,nationality:'غير سعودي',gender:'ذكر'},
 {offset:4,nationality:'غير سعودي',gender:'أنثى'}
];
function getPositiveInputFields(selectedCenterId:string){
 const cfg=configs.positive;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 27;
 const list:any[]=[];
 for(const rowInfo of positiveRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(positiveColumns).map(Number)){
   list.push({cell:`${colName(col)}${row}`,row,col,label:`${positiveColumns[col]} — ${rowInfo.nationality} — ${rowInfo.gender}`});
  }
 }
 return list;
}


const xrayColumns:any={
 4:'عدد المراجعين',
 5:'عدد الأفلام'
};
const xrayRows:any=[
 {offset:0,nationality:'سعودي',gender:'ذكر'},
 {offset:1,nationality:'سعودي',gender:'أنثى'},
 {offset:3,nationality:'غير سعودي',gender:'ذكر'},
 {offset:4,nationality:'غير سعودي',gender:'أنثى'}
];
function getXrayInputFields(selectedCenterId:string){
 const cfg=configs.xray;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 27;
 const list:any[]=[];
 for(const rowInfo of xrayRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(xrayColumns).map(Number)){
   list.push({cell:`${colName(col)}${row}`,row,col,label:`${xrayColumns[col]} — ${rowInfo.nationality} — ${rowInfo.gender}`});
  }
 }
 return list;
}

const minorSurgeryColumns:any={
 4:'جراحات صغرى',
 5:'غيارات',
 6:'حقن',
 7:'أكسجين'
};
const minorSurgeryRows:any=[
 {offset:0,nationality:'سعودي',gender:'ذكر'},
 {offset:1,nationality:'سعودي',gender:'أنثى'},
 {offset:2,nationality:'غير سعودي',gender:'ذكر'},
 {offset:3,nationality:'غير سعودي',gender:'أنثى'}
];
function getMinorSurgeryInputFields(selectedCenterId:string){
 const cfg=configs.minor_surgery;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 20;
 const list:any[]=[];
 for(const rowInfo of minorSurgeryRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(minorSurgeryColumns).map(Number)){
   list.push({cell:`${colName(col)}${row}`,row,col,label:`${minorSurgeryColumns[col]} — ${rowInfo.nationality} — ${rowInfo.gender}`});
  }
 }
 return list;
}


const referralsColumns:any={
 4:'إسعافات أولية',
 5:'إحالات طارئة',
 6:'إحالات عادية',
 8:'الإحالات الراجعة'
};
const referralsRows:any=[
 {offset:0,nationality:'سعودي',gender:'ذكر'},
 {offset:1,nationality:'سعودي',gender:'أنثى'},
 {offset:2,nationality:'غير سعودي',gender:'ذكر'},
 {offset:3,nationality:'غير سعودي',gender:'أنثى'}
];


const pregnancyEndFirstPartColumns:any={
 3:'عدد النساء اللاتي انتهى حملهن',
 4:'عدد الزيارات أثناء الحمل — أقل من 4 زيارات',
 6:'عدد الزيارات أثناء الحمل — أكثر من 4 زيارات',
 8:'عدد زيارات النفاس — خلال 10 أيام الأولى',
 9:'عدد زيارات النفاس — نهاية الفترة',
 11:'استكمال تحصينات الكزاز — أكملت',
 13:'استكمال تحصينات الكزاز — لم تستكمل'
};
const pregnancyEndDeliveryColumns:any={
 4:'مكان الولادة — بالمنزل — بإشراف',
 5:'مكان الولادة — بالمنزل — بدون إشراف',
 6:'مكان الولادة — بالمركز الصحي',
 7:'مكان الولادة — بالمستشفى',
 8:'نتيجة الحمل — مولود حي — طبيعي',
 10:'نتيجة الحمل — مولود حي — مبتسر',
 11:'نتيجة الحمل — مولود حي — أقل من 2.5 كجم',
 13:'نتيجة الحمل — إجهاض',
 14:'نتيجة الحمل — مولود ميت'
};
const pregnancyEndMainRows:any=[
 {offset:0,label:'سعوديات'},
 {offset:1,label:'غير سعوديات'}
];
const pregnancyEndDeliveryRows:any=[
 {offset:6,label:'سعوديات'},
 {offset:7,label:'غير سعوديات'}
];
function getPregnancyEndInputFields(selectedCenterId:string){
 const cfg=configs.pregnancy_end;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 38;
 const list:any[]=[];
 for(const rowInfo of pregnancyEndMainRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(pregnancyEndFirstPartColumns).map(Number)){
   list.push({cell:`${colName(col)}${row}`,row,col,label:`${pregnancyEndFirstPartColumns[col]} — ${rowInfo.label}`});
  }
 }
 for(const rowInfo of pregnancyEndDeliveryRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(pregnancyEndDeliveryColumns).map(Number)){
   list.push({cell:`${colName(col)}${row}`,row,col,label:`${pregnancyEndDeliveryColumns[col]} — ${rowInfo.label}`});
  }
 }
 return list;
}

const maternityColumns:any={
 3:'إجمالي الحوامل المسجلات',
 4:'الحوامل الجدد',
 5:'المعرضات للخطر',
 6:'زيارات الحوامل خلال الشهر'
};
const maternityRows:any=[
 {offset:0,nationality:'سعودي'},
 {offset:1,nationality:'غير سعودي'}
];
function getMaternityInputFields(selectedCenterId:string){
 const cfg=configs.maternity;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 17;
 const list:any[]=[];
 for(const rowInfo of maternityRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(maternityColumns).map(Number)){
   const cell=`${colName(col)}${row}`;
   list.push({cell,row,col,label:`${maternityColumns[col]} — ${rowInfo.nationality}`});
  }
 }
 // يظهر في قالب Excel مرة واحدة لكل مركز، ولا يتم تكراره لغير السعودي.
 list.push({cell:`H${startRow}`,row:startRow,col:8,label:'عدد النساء في سن الإنجاب'});
 return list;
}

function getReferralsInputFields(selectedCenterId:string){
 const cfg=configs.referrals;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 5;
 const list:any[]=[];
 for(const rowInfo of referralsRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(referralsColumns).map(Number)){
   const cell=`${colName(col)}${row}`;
   list.push({cell,row,col,label:`${referralsColumns[col]} — ${rowInfo.nationality} — ${rowInfo.gender}`});
  }
 }
 return list;
}


const dentalServiceRows:any[]=[
 {offset:2,label:'الأشعة — إطباقية'},
 {offset:3,label:'الأشعة — ذروية'},
 {offset:4,label:'الأشعة — خارج الفم'},
 {offset:5,label:'الإجراءات العلاجية / الخلع — لبني'},
 {offset:6,label:'الإجراءات العلاجية / الخلع — دائم'},
 {offset:7,label:'الإجراءات العلاجية / الخلع — جراحي'},
 {offset:8,label:'الإجراءات العلاجية / الجراحة — جراحة'},
 {offset:9,label:'الإجراءات العلاجية / الجراحة — متابعة'},
 {offset:10,label:'الإجراءات العلاجية / الحشو — مؤقت'},
 {offset:11,label:'الإجراءات العلاجية / الحشو — أملجم'},
 {offset:12,label:'الإجراءات العلاجية / الحشو — كومبوزيت'},
 {offset:13,label:'الإجراءات العلاجية / الحشو — لا ينومر زجاج'},
 {offset:14,label:'الإجراءات العلاجية / علاج العصب — لبني'},
 {offset:15,label:'الإجراءات العلاجية / علاج العصب — تهيئة'},
 {offset:16,label:'الإجراءات العلاجية / علاج العصب — حشو'},
 {offset:17,label:'الإجراءات العلاجية / علاج اللثة — كحت جير'},
 {offset:18,label:'الإجراءات العلاجية / علاج اللثة — جراحة لثة'},
 {offset:19,label:'الإجراءات العلاجية — أمراض الفم'},
 {offset:20,label:'الإجراءات الوقائية — توعية صحية'},
 {offset:21,label:'الإجراءات الوقائية — رعاية حوامل'},
 {offset:22,label:'الإجراءات الوقائية — حشو الحفر'},
 {offset:23,label:'الإجراءات الوقائية — فلوريد موضعي'}
];
const dentalPatientRows:any[]=[
 {offset:1,label:'عدد المراجعين — بالغين — ذكور'},
 {offset:3,label:'عدد المراجعين — بالغين — إناث'},
 {offset:5,label:'عدد المراجعين — أطفال — ذكور'},
 {offset:7,label:'عدد المراجعين — أطفال — إناث'},
 {offset:12,label:'عدد المراجعين — سعودي'},
 {offset:15,label:'عدد المراجعين — غير سعودي'},
 {offset:19,label:'عدد المراجعين — جديد'},
 {offset:22,label:'عدد المراجعين — متابعة'}
];
const dentalMedicineCols:any={
 2:'أدوية منصرفة — مضادات حيوية',
 3:'أدوية منصرفة — مسكنات',
 4:'أدوية منصرفة — غسول ومضمضة للفم',
 5:'أدوية منصرفة — علاج لتقرحات الفم',
 6:'أدوية منصرفة — أدوية حساسية',
 7:'أدوية منصرفة — أخرى'
};
function getDentalInputFields(selectedCenterId:string){
 const cfg=configs.dental;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 80;
 const list:any[]=[];
 for(const item of dentalServiceRows){
  const row=startRow+item.offset;
  list.push({cell:`F${row}`,row,col:6,label:`${item.label} — عدد المراجعين`});
 }
 for(const item of dentalPatientRows){
  const row=startRow+item.offset;
  list.push({cell:`K${row}`,row,col:11,label:item.label});
 }
 const medicineRow=startRow+25;
 for(const col of Object.keys(dentalMedicineCols).map(Number)){
  list.push({cell:`${colName(col)}${medicineRow}`,row:medicineRow,col,label:dentalMedicineCols[col]});
 }
 return list;
}

const chronicColumns:any={
 4:'مرضى السكري — المسجلين',
 5:'مرضى السكري — الجدد',
 6:'مرضى السكري — المحالين',
 7:'مرضى السكري — عدد الزيارات خلال الشهر',
 8:'مرضى ضغط الدم — المسجلين',
 9:'مرضى ضغط الدم — الجدد',
 10:'مرضى ضغط الدم — المحالين',
 11:'مرضى ضغط الدم — عدد الزيارات خلال الشهر',
 12:'مرضى الربو — المسجلين',
 13:'مرضى الربو — الجدد',
 14:'مرضى الربو — المحالين',
 15:'مرضى الربو — عدد الزيارات خلال الشهر'
};
const chronicRows:any=[
 {offset:0,nationality:'سعودي',gender:'ذكر'},
 {offset:1,nationality:'سعودي',gender:'أنثى'},
 {offset:2,nationality:'غير سعودي',gender:'ذكر'},
 {offset:3,nationality:'غير سعودي',gender:'أنثى'}
];
function getChronicInputFields(selectedCenterId:string){
 const cfg=configs.chronic;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 5;
 const list:any[]=[];
 for(const rowInfo of chronicRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(chronicColumns).map(Number)){
   const cell=`${colName(col)}${row}`;
   list.push({cell,row,col,label:`${chronicColumns[col]} — ${rowInfo.nationality} — ${rowInfo.gender}`});
  }
 }
 return list;
}

const visitorsRows:any=[
 {offset:0,nationality:'سعودي',gender:'ذكر'},
 {offset:1,nationality:'سعودي',gender:'أنثى'},
 {offset:3,nationality:'غير سعودي',gender:'ذكر'},
 {offset:4,nationality:'غير سعودي',gender:'أنثى'}
];
function getVisitorsInputFields(selectedCenterId:string){
 const cfg=configs.visitors;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 6;
 const list:any[]=[];
 for(const rowInfo of visitorsRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(visitorsColumns).map(Number)){
   const cell=`${colName(col)}${row}`;
   list.push({cell,row,col,label:`${visitorsColumns[col]} — ${rowInfo.nationality} — ${rowInfo.gender}`});
  }
 }
 return list;
}

const workforceRows:any=[
 {offset:0,nationality:'سعودي',gender:'ذكر'},
 {offset:1,nationality:'سعودي',gender:'أنثى'},
 {offset:2,nationality:'غير سعودي',gender:'ذكر'},
 {offset:3,nationality:'غير سعودي',gender:'أنثى'}
];
function getWorkforceInputFields(selectedCenterId:string){
 const cfg=configs.workforce;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 5;
 const list:any[]=[];
 for(const rowInfo of workforceRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(workforceColumns).map(Number)){
   const cell=`${colName(col)}${row}`;
   list.push({cell,row,col,label:`${workforceColumns[col]} — ${rowInfo.nationality} — ${rowInfo.gender}`});
  }
 }
 return list;
}

function getPopulationInputFields(selectedCenterId:string){
 const cfg=configs.population;
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 5;
 const list:any[]=[];
 for(const rowInfo of populationRows){
  const row=startRow+rowInfo.offset;
  for(const col of Object.keys(populationAgeColumns).map(Number)){
   const cell=`${colName(col)}${row}`;
   list.push({cell,row,col,label:`${populationAgeColumns[col]} — ${rowInfo.nationality} — ${rowInfo.gender}`});
  }
 }
 // عدد الأسر في القالب الأصلي يظهر كخلايا مدمجة لكل جنسية، لذلك يظهر للمستخدم مرة واحدة لكل جنسية فقط.
 list.push({cell:`K${startRow}`,row:startRow,col:11,label:'عدد الأسر — سعودي'});
 list.push({cell:`K${startRow+2}`,row:startRow+2,col:11,label:'عدد الأسر — غير سعودي'});
 return list;
}

function baseBlockWidth(m:any){return Number(m.blockWidth||999)}
function headerForColumn(m:any,col:number){
 const hit=m.fields.find((x:any)=>Number(x.row)===4 && Number(x.col)===col) || m.fields.find((x:any)=>Number(x.row)===3 && Number(x.col)===col) || m.fields.find((x:any)=>Number(x.row)===2 && Number(x.col)===col);
 const label=hit?cleanLabel(hit.label,hit.cell):'';
 if(!label || label==='حقل إدخال إحصائي') return `حقل إدخال ${colName(col)}`;
 return label.replace('بيان بالقوى العاملة بالمركز خلال الشهر / ','').replace('بيان عدد السكان والأسر المشمولين بالخدمة من واقع الملفات الصحية / ','').replace('المرضى والأفلام / ','').replace('العمر بالسنين / ','');
}
function isAutoTotalHeader(text:string){
 const t=(text||'').replace(/أ/g,'ا').replace(/إ/g,'ا').trim();
 return /(^|\s)(اجمالي|المجموع|مجموع|اجمالي\s+المراجعين|مجموع\s+مراجعي|مجموع\s+الخدمات)(\s|$)/.test(t) || t.includes('حقل اجمالي');
}
function isManualInputColumn(m:any,cfg:any,col:number){
 const header=headerForColumn(m,col);
 if((cfg.autoTotalCols||[]).includes(col)) return false;
 if(isAutoTotalHeader(header)) return false;
 return true;
}
function getInputFields(m:any,selectedCenterId:string){
 if(m.id==='population') return getPopulationInputFields(selectedCenterId);
 if(m.id==='workforce') return getWorkforceInputFields(selectedCenterId);
 if(m.id==='visitors') return getVisitorsInputFields(selectedCenterId);
 if(m.id==='chronic') return getChronicInputFields(selectedCenterId);
 if(m.id==='referrals') return getReferralsInputFields(selectedCenterId);
 if(m.id==='childcare') return getChildcareInputFields(selectedCenterId);
 if(m.id==='dental') return getDentalInputFields(selectedCenterId);
 if(m.id==='maternity') return getMaternityInputFields(selectedCenterId);
 if(m.id==='pregnancy_end') return getPregnancyEndInputFields(selectedCenterId);
 if(m.id==='lab') return getLabInputFields(selectedCenterId);
 if(m.id==='positive') return getPositiveInputFields(selectedCenterId);
 if(m.id==='xray') return getXrayInputFields(selectedCenterId);
 if(m.id==='minor_surgery') return getMinorSurgeryInputFields(selectedCenterId);
 const cfg=configs[m.id] || {starts:defaultFiveRowStarts,rowOffsets:[0,1,2,3],inputCols:Array.from({length:baseBlockWidth(m)},(_,i)=>i+1),rowCats:['سعودي / ذكر','سعودي / أنثى','غير سعودي / ذكر','غير سعودي / أنثى']};
 const startRow=cfg.starts[selectedCenterId] || cfg.starts.c04 || 5;
 const list:any[]=[];
 for(let i=0;i<cfg.rowOffsets.length;i++){
  const row=startRow+cfg.rowOffsets[i];
  for(const col of cfg.inputCols){
   if(col<1 || col>baseBlockWidth(m)) continue;
   if(!isManualInputColumn(m,cfg,col)) continue;
   const cell=`${colName(col)}${row}`;
   const header=headerForColumn(m,col);
   const cat=cfg.rowCats?.[i] ? ` — ${cfg.rowCats[i]}` : '';
   list.push({cell,row,col,label:`${header}${cat}`});
  }
 }
 return list;
}

function applyAutomaticTotals(ws:any,m:any,cid:string,delta:number){
 const cfg=configs[m.id];
 if(!cfg) return;

 if(m.id==='pregnancy_end'){
  const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 38;
  const totalMain=startRow+2;
  for(const col of [3,4,6,8,9,11,13]){
   ws.getCell(`${colName(col+delta)}${totalMain}`).value={formula:`SUM(${colName(col+delta)}${startRow}:${colName(col+delta)}${startRow+1})`};
  }
  const totalDelivery=startRow+8;
  for(const col of [4,5,6,7,8,10,11,13,14]){
   ws.getCell(`${colName(col+delta)}${totalDelivery}`).value={formula:`SUM(${colName(col+delta)}${startRow+6}:${colName(col+delta)}${startRow+7})`};
  }
  return;
 }
 if(m.id==='lab'){
  const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 27;
  // صفوف المجموع لا تظهر للمستخدم: مجموع السعوديين، مجموع غير السعوديين، والمجموع الكلي.
  for(const col of Object.keys(labColumns).map(Number)){
   ws.getCell(`${colName(col+delta)}${startRow+2}`).value={formula:`SUM(${colName(col+delta)}${startRow}:${colName(col+delta)}${startRow+1})`};
   ws.getCell(`${colName(col+delta)}${startRow+5}`).value={formula:`SUM(${colName(col+delta)}${startRow+3}:${colName(col+delta)}${startRow+4})`};
   ws.getCell(`${colName(col+delta)}${startRow+6}`).value={formula:`SUM(${colName(col+delta)}${startRow+2},${colName(col+delta)}${startRow+5})`};
  }
  for(const row of [startRow,startRow+1,startRow+2,startRow+3,startRow+4,startRow+5,startRow+6]){
   ws.getCell(`${colName(17+delta)}${row}`).value={formula:`SUM(${colName(6+delta)}${row}:${colName(16+delta)}${row})`};
  }
  return;
 }
 if(m.id==='positive'){
  const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 27;
  // صفوف المجموع لا تظهر للمستخدم: مجموع السعوديين، مجموع غير السعوديين، والمجموع الكلي.
  for(const col of Object.keys(positiveColumns).map(Number)){
   ws.getCell(`${colName(col+delta)}${startRow+2}`).value={formula:`SUM(${colName(col+delta)}${startRow}:${colName(col+delta)}${startRow+1})`};
   ws.getCell(`${colName(col+delta)}${startRow+5}`).value={formula:`SUM(${colName(col+delta)}${startRow+3}:${colName(col+delta)}${startRow+4})`};
   ws.getCell(`${colName(col+delta)}${startRow+6}`).value={formula:`SUM(${colName(col+delta)}${startRow+2},${colName(col+delta)}${startRow+5})`};
  }
  for(const row of [startRow,startRow+1,startRow+2,startRow+3,startRow+4,startRow+5,startRow+6]){
   ws.getCell(`${colName(10+delta)}${row}`).value={formula:`SUM(${colName(4+delta)}${row}:${colName(9+delta)}${row})`};
  }
  return;
 }
 if(m.id==='xray'){
  const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 27;
  // صفوف المجموع لا تظهر للمستخدم: مجموع السعوديين، مجموع غير السعوديين، والمجموع الكلي.
  for(const col of Object.keys(xrayColumns).map(Number)){
   ws.getCell(`${colName(col+delta)}${startRow+2}`).value={formula:`SUM(${colName(col+delta)}${startRow}:${colName(col+delta)}${startRow+1})`};
   ws.getCell(`${colName(col+delta)}${startRow+5}`).value={formula:`SUM(${colName(col+delta)}${startRow+3}:${colName(col+delta)}${startRow+4})`};
   ws.getCell(`${colName(col+delta)}${startRow+6}`).value={formula:`SUM(${colName(col+delta)}${startRow+2},${colName(col+delta)}${startRow+5})`};
  }
  return;
 }

 if(m.id==='minor_surgery'){
  const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 20;
  const totalRow=startRow+4;
  // صف المجموع لا يظهر للمستخدم، ويُحسب تلقائياً من الصفوف الأساسية الأربعة.
  for(const col of Object.keys(minorSurgeryColumns).map(Number)){
   ws.getCell(`${colName(col+delta)}${totalRow}`).value={formula:`SUM(${colName(col+delta)}${startRow}:${colName(col+delta)}${startRow+3})`};
  }
  return;
 }

 if(m.id==='maternity'){
  const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 17;
  const totalRow=startRow+2;
  // صف المجموع لا يظهر للمستخدم، ويحسب تلقائياً من سعودي + غير سعودي.
  for(const col of [3,4,5,6]){
   ws.getCell(`${colName(col+delta)}${totalRow}`).value={formula:`SUM(${colName(col+delta)}${startRow}:${colName(col+delta)}${startRow+1})`};
  }
  // عدد النساء في سن الإنجاب يظهر كقيمة مركزية واحدة، والمتزوجات = 38% كما في القالب الأصلي.
  ws.getCell(`${colName(8+delta)}${totalRow}`).value={formula:`${colName(8+delta)}${startRow}`};
  ws.getCell(`${colName(10+delta)}${startRow}`).value={formula:`${colName(8+delta)}${startRow}*0.38`};
  ws.getCell(`${colName(10+delta)}${totalRow}`).value={formula:`${colName(10+delta)}${startRow}`};
  return;
 }
 if(m.id==='dental'){
  const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 80;
  // مجاميع عدد المراجعين في نموذج الأسنان لا تظهر للمستخدم، وتُعاد كمعادلات داخل القالب.
  ws.getCell(`${colName(11+delta)}${startRow+9}`).value={formula:`SUM(${colName(11+delta)}${startRow+1}:${colName(11+delta)}${startRow+8})`};
  ws.getCell(`${colName(11+delta)}${startRow+18}`).value={formula:`SUM(${colName(11+delta)}${startRow+12}:${colName(11+delta)}${startRow+17})`};
  ws.getCell(`${colName(11+delta)}${startRow+25}`).value={formula:`SUM(${colName(11+delta)}${startRow+19}:${colName(11+delta)}${startRow+24})`};
  return;
 }
 if(m.id==='visitors'){
  const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 6;
  const dataRows=[startRow,startRow+1,startRow+3,startRow+4];
  // مجموع مراجعي العيادات لكل صف = 1+2+3+4+5 حسب النموذج الأصلي.
  for(const row of dataRows){
   ws.getCell(`${colName(9+delta)}${row}`).value={formula:`SUM(${colName(4+delta)}${row}:${colName(8+delta)}${row})`};
   // مجموع الخدمات والمساندة لكل صف = 7+8+9 حسب النموذج الأصلي.
   ws.getCell(`${colName(13+delta)}${row}`).value={formula:`SUM(${colName(10+delta)}${row}:${colName(12+delta)}${row})`};
  }
  // مجموع السعوديين.
  const saudiTotal=startRow+2;
  for(const col of [4,5,6,7,8,9,10,11,12,13]){
   ws.getCell(`${colName(col+delta)}${saudiTotal}`).value={formula:`SUM(${colName(col+delta)}${startRow}:${colName(col+delta)}${startRow+1})`};
  }
  // مجموع غير السعوديين.
  const nonSaudiTotal=startRow+5;
  for(const col of [4,5,6,7,8,9,10,11,12,13]){
   ws.getCell(`${colName(col+delta)}${nonSaudiTotal}`).value={formula:`SUM(${colName(col+delta)}${startRow+3}:${colName(col+delta)}${startRow+4})`};
  }
  // المجموع الكلي للمركز.
  const grandTotal=startRow+6;
  for(const col of [4,5,6,7,8,9,10,11,12,13]){
   ws.getCell(`${colName(col+delta)}${grandTotal}`).value={formula:`SUM(${colName(col+delta)}${saudiTotal},${colName(col+delta)}${nonSaudiTotal})`};
  }
  // إجمالي المراجعين للمركز خلال الشهر في أعلى نفس بلوك الشهر.
  ws.getCell(`${colName(13+delta)}2`).value={formula:`${colName(9+delta)}${grandTotal}`};
  return;
 }
 if(m.id==='childcare'){
  const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 23;
  const totalRow=startRow+4;
  // صف المجموع لا يظهر للمستخدم، ويُحسب تلقائياً من الصفوف الأساسية الأربعة.
  for(const col of [4,5,6,7,8,9,10,11,12,13,14]){
   ws.getCell(`${colName(col+delta)}${totalRow}`).value={formula:`SUM(${colName(col+delta)}${startRow}:${colName(col+delta)}${startRow+3})`};
  }
  return;
 }
 if(m.id==='population'){
  const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 5;
  const rows=[startRow,startRow+1,startRow+2,startRow+3];
  for(const row of rows){
   ws.getCell(`${colName(10+delta)}${row}`).value={formula:`SUM(${colName(4+delta)}${row}:${colName(9+delta)}${row})`};
  }
  const totalRow=startRow+4;
  for(const col of [4,5,6,7,8,9,10]){
   ws.getCell(`${colName(col+delta)}${totalRow}`).value={formula:`SUM(${colName(col+delta)}${startRow}:${colName(col+delta)}${startRow+3})`};
  }
  ws.getCell(`${colName(11+delta)}${totalRow}`).value={formula:`SUM(${colName(11+delta)}${startRow},${colName(11+delta)}${startRow+2})`};
  return;
 }
 const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 5;
 const manualCols=(cfg.inputCols||[]).filter((col:number)=>isManualInputColumn(m,cfg,col));
 const autoTotalCols:number[]=(cfg.autoTotalCols||[]);
 const rows=(cfg.rowOffsets||[]).map((off:number)=>startRow+off);
 // حساب أعمدة الإجمالي لكل صف من صفوف الإدخال، دون أن تظهر للمستخدم.
 for(const row of rows){
  for(const totalCol of autoTotalCols){
   const left=Math.min(...manualCols.filter((c:number)=>c<totalCol));
   const right=Math.max(...manualCols.filter((c:number)=>c<totalCol));
   if(Number.isFinite(left) && Number.isFinite(right) && right>=left){
    const target=ws.getCell(`${colName(totalCol+delta)}${row}`);
    target.value={formula:`SUM(${colName(left+delta)}${row}:${colName(right+delta)}${row})`};
   }
  }
 }
 // صف المجموع النهائي للمركز: يجمع الصفوف الأساسية السابقة فقط.
 if(typeof cfg.totalRowOffset==='number'){
  const totalRow=startRow+cfg.totalRowOffset;
  const cols=[...manualCols,...autoTotalCols];
  for(const col of cols){
   const refs=rows.map((r:number)=>`${colName(col+delta)}${r}`).join(',');
   if(refs){
    ws.getCell(`${colName(col+delta)}${totalRow}`).value={formula:`SUM(${refs})`};
   }
  }
 }
}


async function ensureDefaultUsers(){
 try{
  const flagKey='firebaseDefaultUsersSeeded:v30';
  if(typeof window!=='undefined' && localStorage.getItem(flagKey)==='yes') return;
  await Promise.all(users.map((x:any)=>setDoc(doc(db,'users',x.username),{
   username:x.username,
   password:x.password,
   name:x.name,
   role:x.role,
   centerId:x.centerId||'',
   centerName:x.centerId?(centers.find(c=>c.id===x.centerId)?.name||''):'جميع المراكز',
   isActive:x.isActive!==false,
   updatedAt:serverTimestamp()
  },{merge:true})));
  await Promise.all(centers.map((c:any)=>setDoc(doc(db,'centers',c.id),{
   id:c.id,
   name:c.name,
   isActive:true,
   updatedAt:serverTimestamp()
  },{merge:true})));
  if(typeof window!=='undefined') localStorage.setItem(flagKey,'yes');
 }catch(error){
  console.warn('تعذر تهيئة مستخدمي Firebase، سيتم استخدام المستخدمين المحليين مؤقتاً', error);
 }
}

function sanitizeUser(raw:any){
 if(!raw) return null;
 return {
  username:String(raw.username||''),
  password:String(raw.password||''),
  name:String(raw.name||raw.username||''),
  role:(raw.role||'center') as Role,
  centerId:String(raw.centerId||''),
  isActive:raw.isActive!==false
 };
}
function sanitizeCenter(raw:any, id?:string){
 if(!raw && !id) return null;
 const centerId=String(raw?.id||id||'').trim();
 if(!centerId) return null;
 return {id:centerId,name:String(raw?.name||raw?.centerName||centerId).trim(),sector:String(raw?.sector||'شرق جدة').trim(),isActive:raw?.isActive!==false};
}

export default function Page(){
 const [user,setUser]=useState<any>(null); const [u,setU]=useState(''); const [p,setP]=useState('');
 const [tab,setTab]=useState('dashboard'); const [moduleId,setModuleId]=useState<string>(modules[0].id);
 const [centerId,setCenterId]=useState('c04'); const [year,setYear]=useState('1447'); const [month,setMonth]=useState('1');
 const [search,setSearch]=useState(''); const [tick,setTick]=useState(0); const [toast,setToast]=useState(''); const [firebaseStatus,setFirebaseStatus]=useState<'idle'|'loading'|'saved'|'error'>('idle');
 const [auditRows,setAuditRows]=useState<any[]>([]); const [monthLocked,setMonthLocked]=useState(false);
 const emptyUserForm:any={username:'',password:'123456',name:'',role:'center' as Role,centerId:'c01',isActive:true};
 const [userRows,setUserRows]=useState<any[]>([]); const [userForm,setUserForm]=useState<any>(emptyUserForm); const [editingUsername,setEditingUsername]=useState('');
 const emptyCenterForm:any={id:'',name:'',sector:'شرق جدة',isActive:true};
 const [centerRows,setCenterRows]=useState<any[]>(centers.map(c=>({...c,sector:'شرق جدة',isActive:true}))); const [centerForm,setCenterForm]=useState<any>(emptyCenterForm); const [editingCenterId,setEditingCenterId]=useState('');
 useEffect(()=>{const s=localStorage.getItem('currentUser'); if(s) setUser(JSON.parse(s)); ensureDefaultUsers(); loadCenterRows();},[]);
 const centerName=(id:string)=>centerRows.find((c:any)=>c.id===id)?.name || centers.find(c=>c.id===id)?.name || id;
 const activeCenters=useMemo(()=>centerRows.filter((c:any)=>c.isActive!==false),[centerRows]);
 const allowedCenters=useMemo(()=> user?.role==='center'?activeCenters.filter((c:any)=>c.id===user.centerId):activeCenters,[user,activeCenters]);
 useEffect(()=>{ if(user?.role==='center') setCenterId(user.centerId)},[user]);
 const mod=modules.find(m=>m.id===moduleId)!;
 const showToast=(message:string)=>{setToast(message); window.setTimeout(()=>setToast(''),2800)};
 const metaKey=(cid:string,yr:string,mo:string,mId:string)=>`formMeta:${cid}:${yr}:${mo}:${mId}`;
 const readMeta=(cid=centerId,yr=year,mo=month,mId=moduleId)=>{
  if(typeof window==='undefined') return null;
  const raw=localStorage.getItem(metaKey(cid,yr,mo,mId));
  if(!raw) return null;
  try{return JSON.parse(raw)}catch{return null}
 };
 const formatDateTime=(value?:string)=>value?new Date(value).toLocaleString('ar-SA',{hour:'2-digit',minute:'2-digit',day:'2-digit',month:'long',year:'numeric'}):'لا يوجد حفظ معتمد';
 const roleTitle=(role:Role)=> role==='admin'?'مدير النظام':role==='executive'?'إدارة شؤون المراكز الصحية':'مسؤول إحصاء مركز';
 const canManage=()=>user?.role==='admin'||user?.role==='executive';
 const addAudit=async(action:string,details:any={})=>{
  try{await addDoc(collection(db,'audit_logs'),{action,details,username:user?.username||u||'',userName:user?.name||'',role:user?.role||'',centerId,year,month,moduleId,createdAt:serverTimestamp(),createdAtText:new Date().toISOString()})}catch(e){console.warn('audit log failed',e)}
 };
 const loadAuditRows=async()=>{
  try{const qs=await getDocs(query(collection(db,'audit_logs'),orderBy('createdAtText','desc'),limit(60))); setAuditRows(qs.docs.map(d=>({id:d.id,...d.data()})))}catch(e){console.warn('audit load failed',e)}
 };
 const loadUserRows=async()=>{
  if(user?.role!=='admin') return;
  try{
   const qs=await getDocs(collection(db,'users'));
   const rows=qs.docs.map(d=>sanitizeUser({...d.data(), username:d.id})).filter(Boolean) as any[];
   rows.sort((a:any,b:any)=>String(a.role).localeCompare(String(b.role)) || String(a.username).localeCompare(String(b.username)));
   setUserRows(rows.length?rows:users);
  }catch(e){console.warn('users load failed',e); setUserRows(users)}
 };
 const resetUserForm=()=>{setEditingUsername(''); setUserForm({...emptyUserForm})};
 const editUser=(row:any)=>{
  setEditingUsername(row.username);
  setUserForm({username:row.username||'',password:row.password||'',name:row.name||'',role:(row.role||'center') as Role,centerId:row.centerId||'c01',isActive:row.isActive!==false});
 };
 const saveUser=async()=>{
  if(user?.role!=='admin') return;
  const username=String(userForm.username||'').trim();
  const password=String(userForm.password||'').trim();
  if(!username || !password){alert('يرجى إدخال اسم المستخدم وكلمة المرور'); return;}
  const role=(userForm.role||'center') as Role;
  const centerId=role==='center'?String(userForm.centerId||''):'';
  if(role==='center' && !centerId){alert('يرجى اختيار المركز للمستخدم'); return;}
  const payload:any={
   username,
   password,
   name:String(userForm.name||username).trim(),
   role,
   centerId,
   centerName:centerId?(centerName(centerId)):'جميع المراكز',
   isActive:userForm.isActive!==false,
   updatedBy:user?.username||'',
   updatedAt:serverTimestamp(),
   updatedAtText:new Date().toISOString()
  };
  await setDoc(doc(db,'users',username),payload,{merge:true});
  await addAudit(editingUsername?'تعديل مستخدم':'إضافة مستخدم',{username,role,centerId});
  showToast(editingUsername?'تم تعديل المستخدم':'تم إضافة المستخدم');
  resetUserForm();
  await loadUserRows();
 };
 const removeUser=async(row:any)=>{
  if(user?.role!=='admin') return;
  if(row.username==='admin'){alert('لا يمكن حذف حساب مدير النظام الرئيسي'); return;}
  if(!confirm(`هل تريد حذف المستخدم ${row.username}؟`)) return;
  await deleteDoc(doc(db,'users',row.username));
  await addAudit('حذف مستخدم',{username:row.username,role:row.role,centerId:row.centerId||''});
  showToast('تم حذف المستخدم');
  if(editingUsername===row.username) resetUserForm();
  await loadUserRows();
 };
 const toggleUserActive=async(row:any)=>{
  if(user?.role!=='admin') return;
  if(row.username==='admin'){alert('لا يمكن إيقاف حساب مدير النظام الرئيسي'); return;}
  const next=row.isActive===false;
  await setDoc(doc(db,'users',row.username),{isActive:next,updatedBy:user?.username||'',updatedAt:serverTimestamp(),updatedAtText:new Date().toISOString()},{merge:true});
  await addAudit(next?'تفعيل مستخدم':'إيقاف مستخدم',{username:row.username});
  showToast(next?'تم تفعيل المستخدم':'تم إيقاف المستخدم');
  await loadUserRows();
 };

 const loadCenterRows=async()=>{
  try{
   const qs=await getDocs(collection(db,'centers'));
   const rows=qs.docs.map(d=>sanitizeCenter({...d.data(),id:d.id},d.id)).filter(Boolean) as any[];
   rows.sort((a:any,b:any)=>String(a.id).localeCompare(String(b.id)));
   setCenterRows(rows.length?rows:centers.map(c=>({...c,sector:'شرق جدة',isActive:true})));
  }catch(e){console.warn('centers load failed',e); setCenterRows(centers.map(c=>({...c,sector:'شرق جدة',isActive:true})))}
 };
 const resetCenterForm=()=>{setEditingCenterId(''); setCenterForm({...emptyCenterForm})};
 const editCenter=(row:any)=>{setEditingCenterId(row.id); setCenterForm({id:row.id||'',name:row.name||'',sector:row.sector||'شرق جدة',isActive:row.isActive!==false})};
 const saveCenter=async()=>{
  if(user?.role!=='admin') return;
  const id=String(centerForm.id||'').trim().toLowerCase();
  const name=String(centerForm.name||'').trim();
  if(!id || !name){alert('يرجى إدخال رمز المركز واسم المركز'); return;}
  const payload:any={id,name,sector:String(centerForm.sector||'شرق جدة').trim(),isActive:centerForm.isActive!==false,updatedBy:user?.username||'',updatedAt:serverTimestamp(),updatedAtText:new Date().toISOString()};
  await setDoc(doc(db,'centers',id),payload,{merge:true});
  await addAudit(editingCenterId?'تعديل مركز':'إضافة مركز',{centerId:id,centerName:name});
  showToast(editingCenterId?'تم تعديل المركز':'تم إضافة المركز');
  resetCenterForm();
  await loadCenterRows();
 };
 const toggleCenterActive=async(row:any)=>{
  if(user?.role!=='admin') return;
  const next=row.isActive===false;
  await setDoc(doc(db,'centers',row.id),{isActive:next,updatedBy:user?.username||'',updatedAt:serverTimestamp(),updatedAtText:new Date().toISOString()},{merge:true});
  await addAudit(next?'تفعيل مركز':'إيقاف مركز',{centerId:row.id,centerName:row.name});
  showToast(next?'تم تفعيل المركز':'تم إيقاف المركز');
  await loadCenterRows();
 };
 const removeCenter=async(row:any)=>{
  if(user?.role!=='admin') return;
  if(centers.some(c=>c.id===row.id)){alert('لا يمكن حذف مركز أساسي مرتبط بقوالب Excel. يمكن إيقافه فقط.'); return;}
  if((userRows.length?userRows:users).some((u:any)=>u.centerId===row.id)){alert('لا يمكن حذف مركز مرتبط بمستخدمين. انقل المستخدمين أو احذفهم أولاً.'); return;}
  if(!confirm(`هل تريد حذف المركز ${row.name}؟`)) return;
  await deleteDoc(doc(db,'centers',row.id));
  await addAudit('حذف مركز',{centerId:row.id,centerName:row.name});
  showToast('تم حذف المركز');
  await loadCenterRows();
 };

 const checkMonthLock=async()=>{
  try{const snap=await getDoc(doc(db,'month_locks',`${year}_${month}`)); setMonthLocked(snap.exists() && snap.data()?.locked===true)}catch(e){setMonthLocked(false)}
 };
 const toggleMonthLock=async()=>{
  if(!canManage()) return;
  const next=!monthLocked;
  await setDoc(doc(db,'month_locks',`${year}_${month}`),{year,month,locked:next,updatedBy:user?.username||'',updatedByName:user?.name||'',updatedAt:serverTimestamp(),updatedAtText:new Date().toISOString()},{merge:true});
  setMonthLocked(next);
  await addAudit(next?'اعتماد وقفل الشهر':'إلغاء قفل الشهر',{year,month});
  showToast(next?'تم اعتماد وقفل الشهر':'تم إلغاء قفل الشهر');
 };
 const loadPeriodFromFirebase=async()=>{
  if(!user) return;
  try{
   const scope=user.role==='center'?centers.filter(c=>c.id===user.centerId):allowedCenters;
   await Promise.all(scope.flatMap((c:any)=>modules.map(async(m:any)=>{
    const snap=await getDoc(doc(db,'monthly_statistics',firestoreFormDocId(c.id,year,month,m.id)));
    if(snap.exists()){
     const data:any=snap.data(); const values=data?.values||{};
     Object.entries(values).forEach(([cell,value])=>localStorage.setItem(key(c.id,year,month,m.id,cell),String(value??'')));
     if(data?.meta) localStorage.setItem(metaKey(c.id,year,month,m.id),JSON.stringify(data.meta));
    }
   })));
   setTick(x=>x+1);
  }catch(e){console.warn('period firebase load failed',e)}
 };
 const login=async()=>{
  const username=u.trim();
  const password=p.trim();
  if(!username || !password){alert('يرجى إدخال اسم المستخدم وكلمة المرور'); return;}
  try{
   setFirebaseStatus('loading');
   let found:any=null;
   const snap=await getDoc(doc(db,'users',username));
   if(snap.exists()) found=sanitizeUser(snap.data());
   if(!found){
    found=users.find((x:any)=>x.username===username) || null;
    if(found){
     await setDoc(doc(db,'users',found.username),{
      ...found,
      centerName:found.centerId?(centerName(found.centerId)):'جميع المراكز',
      isActive:found.isActive!==false,
      updatedAt:serverTimestamp()
     },{merge:true});
    }
   }
   if(found && found.isActive!==false && found.password===password){
    localStorage.setItem('currentUser',JSON.stringify(found));
    setUser(found);
    setFirebaseStatus('idle');
    await addDoc(collection(db,'audit_logs'),{action:'تسجيل دخول',details:{username:found.username,role:found.role,centerId:found.centerId||''},username:found.username,userName:found.name,role:found.role,centerId:found.centerId||'',year,month,moduleId:'',createdAt:serverTimestamp(),createdAtText:new Date().toISOString()}).catch(()=>{});
    return;
   }
   setFirebaseStatus('idle');
   alert('بيانات الدخول غير صحيحة أو المستخدم غير مفعل');
  }catch(error){
   console.error('Firebase login error',error);
   const fallback=users.find((x:any)=>x.username===username&&x.password===password);
   if(fallback){localStorage.setItem('currentUser',JSON.stringify(fallback)); setUser(fallback); setFirebaseStatus('idle'); return;}
   setFirebaseStatus('error');
   alert('تعذر الاتصال بقاعدة البيانات وبيانات الدخول غير صحيحة');
  }
 };
 const saveVal=(cell:string,val:string)=>{if(typeof window==='undefined')return; const clean=normalizeNumberText(val).replace(/[^0-9.]/g,'').replace(/(\..*)\./g,'$1'); localStorage.setItem(key(centerId,year,month,moduleId,cell),clean); setTick(x=>x+1)};
 const getVal=(mId:string,cell:string,cid=centerId,mo=month,yr=year)=> typeof window==='undefined'?'':(localStorage.getItem(key(cid,yr,mo,mId,cell))||'');
 const getNum=(mId:string,cell:string,cid=centerId,mo=month,yr=year)=>Number(getVal(mId,cell,cid,mo,yr))||0;
 const getVisitorsSummary=(cid=centerId,mo=month,yr=year)=>{
  const cfg=configs.visitors;
  const startRow=cfg.starts?.[cid] || cfg.starts?.c04 || 6;
  const dataRows=[startRow,startRow+1,startRow+3,startRow+4];
  const clinicTotal=dataRows.reduce((sum,row)=>sum+[4,5,6,7,8].reduce((s,col)=>s+getNum('visitors',`${colName(col)}${row}`,cid,mo,yr),0),0);
  const supportTotal=dataRows.reduce((sum,row)=>sum+[10,11,12].reduce((s,col)=>s+getNum('visitors',`${colName(col)}${row}`,cid,mo,yr),0),0);
  return {monthlyTotal:clinicTotal, clinicTotal, supportTotal};
 };
 const moduleFields=(m:any,cid:string)=>getInputFields(m,cid);
 const moduleFilled=(m:any,cid:string)=>moduleFields(m,cid).filter((f:any)=>getVal(m.id,f.cell,cid)).length;
 const moduleTotal=(m:any,cid:string)=>moduleFields(m,cid).length || 1;
 const formCompletion=(m:any,cid:string)=>{const total=moduleTotal(m,cid); const done=moduleFilled(m,cid); const pct=Math.round((done/Math.max(total,1))*100); return {done,total,pct,status:pct===100?'مكتمل':pct>0?'قيد الإدخال':'لم يبدأ'}};
 const centerCompletion=(cid:string)=>{const total=modules.reduce((s,m)=>s+moduleTotal(m,cid),0); const done=modules.reduce((s,m)=>s+moduleFilled(m,cid),0); const completedForms=modules.filter(m=>formCompletion(m,cid).pct===100).length; const startedForms=modules.filter(m=>formCompletion(m,cid).pct>0).length; return {done,total,completedForms,startedForms,pct:Math.round((done/Math.max(total,1))*100)}};
 const getRecentUpdates=()=>{if(typeof window==='undefined')return []; const rows:any[]=[]; const prefix='stat:'; for(let i=0;i<localStorage.length;i++){const k=localStorage.key(i)||''; if(!k.startsWith(prefix))continue; const parts=k.split(':'); if(parts.length<6)continue; const [,cid,yr,mo,mId,cell]=parts; const value=localStorage.getItem(k)||''; if(!value)continue; const c=centers.find(x=>x.id===cid); const m=modules.find(x=>x.id===mId); if(c&&m&&yr===year&&mo===month) rows.push({center:c.name,module:m.title,value,cell}); } return rows.slice(-6).reverse()};
 const getExecutiveIndicators=()=>{
  const sumModule=(mId:string)=>allowedCenters.reduce((sum,c)=>sum+getInputFields(modules.find(m=>m.id===mId),c.id).reduce((s:any,f:any)=>s+(Number(getVal(mId,f.cell,c.id))||0),0),0);
  return [
   {title:'إجمالي السكان المدخل',value:sumModule('population'),note:'من نموذج السكان والأسر'},
   {title:'إجمالي القوى العاملة',value:sumModule('workforce'),note:'من نموذج القوى العاملة'},
   {title:'إجمالي المراجعين',value:allowedCenters.reduce((sum,c)=>sum+getVisitorsSummary(c.id).monthlyTotal,0),note:'من نموذج المراجعين'},
   {title:'إجمالي الإحالات',value:sumModule('referrals'),note:'من نموذج الإحالات'}
  ];
 };

 const getExecutiveAlerts=()=>{
  const alerts:any[]=[];
  const notStarted=allowedCenters.filter(c=>centerCompletion(c.id).pct===0);
  if(notStarted.length) alerts.push({level:'متأخر',title:'مراكز لم تبدأ الإدخال',text:notStarted.map(c=>c.name.replace('مركز صحي ', '')).join('، ')});
  const incomplete=allowedCenters.filter(c=>{const p=centerCompletion(c.id).pct; return p>0 && p<100}).sort((a,b)=>centerCompletion(a.id).pct-centerCompletion(b.id).pct).slice(0,3);
  incomplete.forEach(c=>alerts.push({level:'متابعة',title:`${c.name.replace('مركز صحي ', '')} قيد الاستكمال`,text:`نسبة الإنجاز الحالية ${centerCompletion(c.id).pct}%، يلزم استكمال النماذج الناقصة قبل التصدير النهائي.`}));
  const moduleGaps=modules.map(m=>({m,done:allowedCenters.filter(c=>formCompletion(m,c.id).pct===100).length})).filter(x=>x.done<allowedCenters.length).slice(0,4);
  moduleGaps.forEach(x=>alerts.push({level:'نموذج ناقص',title:x.m.title,text:`مكتمل في ${x.done} من ${allowedCenters.length} مراكز.`}));
  return alerts.slice(0,6);
 };
 const printExecutiveReport=()=>{showToast('سيتم فتح نافذة الطباعة للتقرير التنفيذي'); window.setTimeout(()=>window.print(),250)};

 const currentFields=moduleFields(mod,centerId);
 const visibleFields=currentFields.filter((f:any)=>cleanLabel(f.label,f.cell).includes(search));
 const currentFilled=currentFields.filter((f:any)=>getVal(moduleId,f.cell)).length;
 const currentPct=Math.round((currentFilled/Math.max(currentFields.length,1))*100);
 const globalStats=allowedCenters.reduce((acc:any,c)=>{const p=centerCompletion(c.id); acc.done+=p.done; acc.total+=p.total; if(p.pct===100) acc.completeCenters++; return acc},{done:0,total:0,completeCenters:0});
 const globalPct=Math.round((globalStats.done/Math.max(globalStats.total,1))*100);
 const incompleteCenters=allowedCenters.length-globalStats.completeCenters;
 const executiveRows=allowedCenters.map(c=>({center:c,stats:centerCompletion(c.id),latest:modules.map(m=>readMeta(c.id,year,month,m.id)).filter(Boolean).sort((a:any,b:any)=>String(b.savedAt).localeCompare(String(a.savedAt)))[0]})).sort((a:any,b:any)=>b.stats.pct-a.stats.pct);
 const fullyCompletedModules=modules.filter(m=>allowedCenters.every(c=>formCompletion(m,c.id).pct===100)).length;
 const startedModules=modules.filter(m=>allowedCenters.some(c=>formCompletion(m,c.id).pct>0)).length;
 const printReport=()=>{showToast('تم تجهيز التقرير للطباعة'); window.setTimeout(()=>window.print(),250)};
 const lastUpdate= new Date().toLocaleString('ar-SA',{hour:'2-digit',minute:'2-digit',day:'2-digit',month:'long',year:'numeric'});
 const clearCurrentForm=()=>{if(!confirm('سيتم مسح بيانات هذا النموذج للمركز والشهر المحددين فقط. هل تريد المتابعة؟'))return; currentFields.forEach((f:any)=>localStorage.removeItem(key(centerId,year,month,moduleId,f.cell))); setTick(x=>x+1); showToast('تم مسح بيانات النموذج الحالي')};
 const copyPreviousMonth=()=>{const mNum=Number(month); const prevMonth=String(mNum===1?12:mNum-1); const prevYear=String(mNum===1?Number(year)-1:year); let count=0; currentFields.forEach((f:any)=>{const v=getVal(moduleId,f.cell,centerId,prevMonth,prevYear); if(v){localStorage.setItem(key(centerId,year,month,moduleId,f.cell),v); count++}}); setTick(x=>x+1); showToast(count?`تم نسخ ${count} قيمة من الشهر السابق`:'لا توجد بيانات في الشهر السابق')};
 const saveCurrentForm=async()=>{
  if(monthLocked && user?.role==='center'){showToast('الشهر معتمد ومغلق، لا يمكن تعديل البيانات إلا من الإدارة'); return;}
  const values:any={};
  currentFields.forEach((f:any)=>{const v=getVal(moduleId,f.cell); if(v!=='') values[f.cell]=v});
  const count=Object.keys(values).length;
  const now=new Date().toISOString();
  const meta={savedAt:now,updatedAt:now,savedBy:user?.name||'',role:roleTitle(user?.role),centerName:centerName(centerId),formTitle:mod.title,completion:currentPct,fieldsCount:count,totalFields:currentFields.length};
  localStorage.setItem(metaKey(centerId,year,month,moduleId),JSON.stringify(meta));
  setTick(x=>x+1);
  try{
   setFirebaseStatus('loading');
   await setDoc(doc(db,'monthly_statistics',firestoreFormDocId(centerId,year,month,moduleId)),{
    centerId,
    centerName:centerName(centerId),
    year,
    month,
    moduleId,
    moduleTitle:mod.title,
    values,
    meta,
    savedBy:user?.username||'',
    savedByName:user?.name||'',
    savedByRole:user?.role||'',
    savedByCenterId:user?.centerId||'',
    updatedAt:serverTimestamp()
   },{merge:true});
   setFirebaseStatus('saved');
   await addAudit('حفظ بيانات نموذج',{centerId,centerName:centerName(centerId),moduleId,moduleTitle:mod.title,fieldsCount:count,completion:currentPct});
   await loadAuditRows();
   showToast(count?`تم حفظ ${count} حقل في قاعدة البيانات بنجاح`:'تم حفظ النموذج بدون بيانات مدخلة');
  }catch(error){
   console.error('Firebase save error',error);
   setFirebaseStatus('error');
   showToast('تم الحفظ محلياً، وتعذر الحفظ في Firebase');
  }
 };
 const currentMeta=readMeta();
 const currentLastSaveText=formatDateTime(currentMeta?.savedAt);
 const currentSavedBy=currentMeta?.savedBy||user?.name||'غير محدد';
 const currentStatusLabel=currentPct===100?'مكتمل':currentPct>0?'قيد الاستكمال':'لم يبدأ';
 const visitorsSummary=mod.id==='visitors'?getVisitorsSummary(centerId):null;
 useEffect(()=>{
  if(!user || typeof window==='undefined') return;
  let cancelled=false;
  const loadRemote=async()=>{
   try{
    setFirebaseStatus('loading');
    const snap=await getDoc(doc(db,'monthly_statistics',firestoreFormDocId(centerId,year,month,moduleId)));
    if(cancelled) return;
    if(snap.exists()){
     const data:any=snap.data();
     const values=data?.values||{};
     currentFields.forEach((f:any)=>localStorage.removeItem(key(centerId,year,month,moduleId,f.cell)));
     Object.entries(values).forEach(([cell,value])=>localStorage.setItem(key(centerId,year,month,moduleId,cell),String(value??'')));
     if(data?.meta) localStorage.setItem(metaKey(centerId,year,month,moduleId),JSON.stringify(data.meta));
     setTick(x=>x+1);
    }
    setFirebaseStatus('idle');
   }catch(error){
    console.error('Firebase load error',error);
    if(!cancelled) setFirebaseStatus('error');
   }
  };
  loadRemote();
  return ()=>{cancelled=true};
 },[user?.username,centerId,year,month,moduleId]);
 useEffect(()=>{ if(user){checkMonthLock(); loadPeriodFromFirebase(); if(user.role==='admin'||user.role==='executive') loadAuditRows();} },[user?.username,year,month]);
 const filled=globalStats.done;
 if(!user) return <div className="loginWrap"><div className="login"><section className="hero"><div className="heroTop"><div className="emblem emblemLarge">PHC</div><span className="officialTag">منصة حكومية داخلية</span></div><h2>منصة الإحصائية الشهرية</h2><p>تحويل نماذج Excel المعتمدة إلى منصة إلكترونية رسمية لإدخال بيانات المراكز الصحية، مع تصدير مطابق للقوالب الأصلية.</p><div className="heroStats"><div><strong>14</strong><span>نموذج شهري</span></div><div><strong>7</strong><span>مراكز صحية</span></div><div><strong>Excel</strong><span>تصدير مطابق</span></div></div></section><section className="loginBox"><div className="loginHeader"><span className="miniLine"></span><h2>تسجيل الدخول</h2><p className="muted">تجمع جدة الصحي الأول — مستشفى شرق جدة — إدارة شؤون المراكز الصحية</p></div><div className="field"><label>اسم المستخدم</label><input value={u} onChange={e=>setU(e.target.value)}/></div><div className="field"><label>كلمة المرور</label><input type="password" value={p} onChange={e=>setP(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')login()}}/></div><button className="btn btnPrimary" style={{width:'100%'}} onClick={login} disabled={firebaseStatus==='loading'}>{firebaseStatus==='loading'?'جاري التحقق...':'دخول المنصة'}</button></section></div></div>;
 async function exportModule(single=true){
  const ExcelJS=(await import('exceljs')).default; const saveAs=(await import('file-saver')).saveAs; const JSZip=(await import('jszip')).default;
  const exportOne=async(m:any,cid:string)=>{const wb=new ExcelJS.Workbook(); const res=await fetch(`/templates/${m.templateFile}`); await wb.xlsx.load(await res.arrayBuffer()); const ws=wb.getWorksheet(m.sheetName) || wb.worksheets[0]; const target=m.monthColumns.find((x:any)=>String(x.month)===String(month)); const delta=(target?target.col:m.firstMonthCol)-m.firstMonthCol;
   setWorkbookPeriodHeaders(wb,ws,m,month,year);
   wb.worksheets.forEach((sheet:any)=>{sheet.eachRow((row:any,rowNumber:number)=>{ if(rowNumber>=5){ row.eachCell({includeEmpty:false},(cell:any)=>{ if(typeof cell.value==='number'){cell.value=null} }); } });});
   getInputFields(m,cid).forEach((f:any)=>{const v=getVal(m.id,f.cell,cid); const {c,r}=cellToParts(f.cell); const cell=ws.getCell(`${colName(c+delta)}${r}`); if(v!==''){const num=Number(v); cell.value=isNaN(num)?v:num;}});
   applyAutomaticTotals(ws,m,cid,delta);
   const b=await wb.xlsx.writeBuffer(); return new Blob([b],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});};
  if(single){const blob=await exportOne(mod,centerId); saveAs(blob,`${fileSafeName(mod.title)}-${fileSafeName(centerName(centerId)||'مركز')}-${cleanYearInput(year)}-شهر-${month}.xlsx`); showToast('تم تصدير النموذج بنجاح')}
  else{const zip=new JSZip(); for(const c of allowedCenters){for(const m of modules){zip.file(`${fileSafeName(c.name)}/${fileSafeName(m.title)}-${cleanYearInput(year)}-شهر-${month}.xlsx`, await exportOne(m,c.id));}} saveAs(await zip.generateAsync({type:'blob'}),`تصدير-الإحصائية-الشهرية-${month}-${year}.zip`); showToast('تم تجهيز ملف التصدير المجمع')}
 }
 return <>
  {toast&&<div className="toast">{toast}</div>}
  <header className="topbar"><div className="brand"><div className="emblem">PHC</div><div><h1>منصة الإحصائية الشهرية</h1><p>تجمع جدة الصحي الأول · مستشفى شرق جدة · إدارة شؤون المراكز الصحية</p></div></div><div className="topActions"><span className="badge badgeLight">{roleTitle(user.role)}</span><span className="badge badgeLight">{user.name}</span><button className="btn btnGhost" onClick={()=>{localStorage.removeItem('currentUser');setUser(null)}}>خروج</button></div></header>
  <main className="layout"><aside><div className="sideTitle"><b>القائمة الرئيسية</b><small>النطاق: {user.role==='center'?centerName(user.centerId):'جميع المراكز'}</small></div><button className={'navBtn '+(tab==='dashboard'?'active':'')} onClick={()=>setTab('dashboard')}>🏠 الرئيسية</button><button className={'navBtn '+(tab==='followup'?'active':'')} onClick={()=>setTab('followup')}>📌 متابعة المراكز</button><button className={'navBtn '+(tab==='report'?'active':'')} onClick={()=>setTab('report')}>📄 التقرير التنفيذي</button><button className={'navBtn '+(tab==='entry'?'active':'')} onClick={()=>setTab('entry')}>📊 الإحصائية الشهرية</button><button className={'navBtn '+(tab==='exports'?'active':'')} onClick={()=>setTab('exports')}>📁 التصدير</button>{(user.role==='admin'||user.role==='executive')&&<button className={'navBtn '+(tab==='audit'?'active':'')} onClick={()=>{setTab('audit');loadAuditRows()}}>🧾 سجل العمليات</button>}{user.role==='admin'&&<button className={'navBtn '+(tab==='users'?'active':'')} onClick={()=>{setTab('users');loadUserRows();loadCenterRows()}}>⚙️ المستخدمون والصلاحيات</button>}{user.role==='admin'&&<button className={'navBtn '+(tab==='centers'?'active':'')} onClick={()=>{setTab('centers');loadCenterRows();loadUserRows()}}>🏥 إدارة المراكز الصحية</button>}<hr/><p className="muted sideSection">النماذج المعتمدة</p>{modules.map(m=><button key={m.id} className={'navBtn moduleBtn '+(moduleId===m.id&&tab==='entry'?'active':'')} onClick={()=>{setModuleId(m.id);setTab('entry')}}><span>{m.title}</span><small>{moduleFilled(m,centerId)}/{moduleTotal(m,centerId)}</small></button>)}</aside>
   <section className="content">
    {tab==='dashboard'&&<div className="pageStack"><div className="welcomeCard executiveHero"><div><span className="officialTag dark">مركز القيادة التنفيذي</span><h2>مركز القيادة التنفيذي للإحصائية الشهرية</h2><p>لوحة تنفيذية موحدة لمتابعة حالة المراكز، المؤشرات الصحية، التنبيهات، والتقارير الشهرية.</p><div className="heroButtons"><button className="btn btnPrimary" onClick={()=>setTab('report')}>عرض التقرير التنفيذي</button><button className="btn btnGhost" onClick={()=>setTab('exports')}>التصدير الشهري</button></div></div><div className="statusPill"><strong>{globalPct}%</strong><span>نسبة الإنجاز العامة</span></div></div>
     <div className="kpis"><div className="kpi"><small>المراكز الصحية</small><strong>{allowedCenters.length}</strong><span>ضمن نطاق المستخدم</span></div><div className="kpi"><small>النماذج المعتمدة</small><strong>{modules.length}</strong><span>نماذج Excel جاهزة</span></div><div className="kpi"><small>المراكز المكتملة</small><strong>{globalStats.completeCenters}</strong><span>من أصل {allowedCenters.length}</span></div><div className="kpi"><small>آخر تحديث</small><strong className="smallStrong">{lastUpdate}</strong><span>توقيت الجهاز الحالي</span></div></div>
     <div className="executiveIndicators">{getExecutiveIndicators().map((x:any)=><div className="indicatorCard" key={x.title}><small>{x.title}</small><strong>{x.value.toLocaleString('en-US')}</strong><span>{x.note}</span></div>)}</div>
     <div className="centerCommandGrid">{allowedCenters.map(c=>{const p=centerCompletion(c.id);return <button key={c.id} className="centerCommandCard" onClick={()=>{setCenterId(c.id);setTab('entry')}}><div><b>{c.name.replace('مركز صحي ', '')}</b><span>{p.completedForms} نماذج مكتملة من {modules.length}</span></div><strong>{p.pct}%</strong><div className="progress full"><i style={{width:`${p.pct}%`}}></i></div></button>})}</div>
     <div className="dashboardGrid"><div className="card"><div className="sectionHead"><div><h2>حالة إنجاز المراكز</h2><p className="muted">متابعة إجمالية لكل مركز حسب عدد الحقول المكتملة في النماذج.</p></div><span className="badge">{globalStats.done} / {globalStats.total}</span></div><table className="table"><thead><tr><th>المركز</th><th>نسبة الإنجاز</th><th>النماذج المكتملة</th><th>المدخل</th><th>الحالة</th></tr></thead><tbody>{allowedCenters.map(c=>{const p=centerCompletion(c.id);return <tr key={c.id}><td>{c.name}</td><td><div className="progress"><i style={{width:`${p.pct}%`}}></i></div><b>{p.pct}%</b></td><td>{p.completedForms} / {modules.length}</td><td>{p.done} / {p.total}</td><td><span className={'badge '+(p.pct===100?'badgeOk':p.pct>0?'badgeWarn':'badgeLate')}>{p.pct===100?'مكتمل':p.pct>0?'قيد الإدخال':'لم يبدأ'}</span></td></tr>})}</tbody></table></div>
      <div className="card"><div className="sectionHead"><div><h2>آخر الإدخالات</h2><p className="muted">آخر القيم المحفوظة للشهر المحدد.</p></div></div><div className="updatesList">{getRecentUpdates().length?getRecentUpdates().map((r:any,i:number)=><div className="updateItem" key={i}><b>{r.center}</b><span>{r.module}</span><small>قيمة مدخلة: {r.value}</small></div>):<div className="emptyState">لا توجد إدخالات محفوظة بعد لهذا الشهر.</div>}</div></div></div>
     <div className="card smartAlerts"><div className="sectionHead"><div><h2>التنبيهات الذكية</h2><p className="muted">تنبيهات تنفيذية مبنية على حالة الإدخال الحالية لكل مركز ونموذج.</p></div><span className="badge badgeWarn">{getExecutiveAlerts().length} تنبيهات</span></div><div className="alertsGrid">{getExecutiveAlerts().length?getExecutiveAlerts().map((a:any,i:number)=><div className="alertItem" key={i}><small>{a.level}</small><b>{a.title}</b><span>{a.text}</span></div>):<div className="emptyState">لا توجد تنبيهات حالياً. جميع المدخلات مستقرة حسب البيانات المحفوظة.</div>}</div></div>
     <div className="card"><div className="sectionHead"><div><h2>مصفوفة حالة النماذج</h2><p className="muted">عرض سريع يوضح هل النموذج مكتمل أو قيد الإدخال أو لم يبدأ لكل مركز.</p></div><button className="btn btnGhost" onClick={()=>setTab('exports')}>الانتقال للتصدير</button></div><div className="matrixWrap"><table className="matrix"><thead><tr><th>المركز</th>{modules.map(m=><th key={m.id}>{m.title}</th>)}</tr></thead><tbody>{allowedCenters.map(c=><tr key={c.id}><td>{c.name}</td>{modules.map(m=>{const f=formCompletion(m,c.id);return <td key={m.id}><button className={'dotStatus '+(f.pct===100?'done':f.pct>0?'progressing':'empty')} title={`${m.title}: ${f.pct}%`} onClick={()=>{setCenterId(c.id);setModuleId(m.id);setTab('entry')}}>{f.pct===100?'✓':f.pct>0?'◐':'—'}</button></td>})}</tr>)}</tbody></table></div></div>
     <div className="card"><div className="sectionHead"><div><h2>جاهزية النماذج للمركز المحدد</h2><p className="muted">اضغط على أي نموذج للانتقال مباشرة إلى صفحة الإدخال.</p></div><span className="badge">{centerName(centerId)}</span></div><div className="moduleCards moduleCardsGrid">{modules.map(m=>{const f=formCompletion(m,centerId);return <button key={m.id} className="miniModule" onClick={()=>{setModuleId(m.id);setTab('entry')}}><span>{m.title}</span><b>{f.done}/{f.total}</b><em>{f.pct}%</em></button>})}</div></div>
    </div>}
    {tab==='followup'&&<div className="pageStack"><div className="welcomeCard executiveHero"><div><span className="officialTag dark">متابعة شهرية</span><h2>متابعة إدخال المراكز الصحية</h2><p>استعراض حالة إنجاز كل مركز صحي للنماذج الشهرية، مع آخر حفظ معتمد ونسبة الإكمال.</p></div><div className="statusPill compact"><strong>{globalPct}%</strong><span>الإنجاز العام</span></div></div><div className="card"><div className="sectionHead"><div><h2>ملخص حالة المراكز</h2><p className="muted">تُحدّث الحالة بناءً على البيانات المحفوظة لكل مركز وشهر وسنة.</p></div><div className="toolbar smallToolbar"><span className={'badge '+(monthLocked?'badgeOk':'badgeWarn')}>{monthLocked?'الشهر معتمد ومغلق':'الشهر مفتوح للإدخال'}</span>{canManage()&&<button className="btn btnGhost" onClick={toggleMonthLock}>{monthLocked?'إلغاء قفل الشهر':'اعتماد وقفل الشهر'}</button>}<span className="badge">{month} / {year}</span></div></div><table className="table"><thead><tr><th>المركز</th><th>نسبة الإنجاز</th><th>النماذج المكتملة</th><th>آخر حفظ معتمد</th><th>الحالة</th></tr></thead><tbody>{allowedCenters.map(c=>{const p=centerCompletion(c.id); const metas=modules.map(m=>readMeta(c.id,year,month,m.id)).filter(Boolean); const latest=metas.sort((a:any,b:any)=>String(b.savedAt).localeCompare(String(a.savedAt)))[0]; return <tr key={c.id}><td>{c.name}</td><td><div className="progress"><i style={{width:`${p.pct}%`}}></i></div><b>{p.pct}%</b></td><td>{p.completedForms} / {modules.length}</td><td>{latest?formatDateTime(latest.savedAt):'لا يوجد حفظ معتمد'}</td><td><span className={'badge '+(p.pct===100?'badgeOk':p.pct>0?'badgeWarn':'badgeLate')}>{p.pct===100?'مكتمل':p.pct>0?'قيد الاستكمال':'لم يبدأ'}</span></td></tr>})}</tbody></table></div><div className="card"><div className="sectionHead"><div><h2>تفصيل النماذج حسب المركز</h2><p className="muted">اختر أي حالة للانتقال إلى نموذج الإدخال مباشرة.</p></div></div><div className="matrixWrap"><table className="matrix"><thead><tr><th>المركز</th>{modules.map(m=><th key={m.id}>{m.title}</th>)}</tr></thead><tbody>{allowedCenters.map(c=><tr key={c.id}><td>{c.name}</td>{modules.map(m=>{const f=formCompletion(m,c.id);return <td key={m.id}><button className={'dotStatus '+(f.pct===100?'done':f.pct>0?'progressing':'empty')} title={`${m.title}: ${f.pct}%`} onClick={()=>{setCenterId(c.id);setModuleId(m.id);setTab('entry')}}>{f.pct===100?'✓':f.pct>0?'◐':'—'}</button></td>})}</tr>)}</tbody></table></div></div></div>}
    {tab==='entry'&&<div className="pageStack"><div className="formHeader"><div><span className="officialTag dark">نموذج إدخال معتمد</span><h2>{mod.title}</h2><p className="muted">إدخال الإحصائية الشهرية حسب النموذج المعتمد، مع احتساب المجاميع تلقائياً عند التصدير.</p></div><div className="statusPill compact"><strong>{currentPct}%</strong><span>{currentFilled} / {currentFields.length}</span></div></div><div className="card"><div className="toolbar"><div className="field"><label>المركز</label><select value={centerId} onChange={e=>setCenterId(e.target.value)} disabled={user.role==='center'}>{allowedCenters.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div><div className="field"><label>الشهر الهجري</label><select value={month} onChange={e=>setMonth(e.target.value)}>{Array.from({length:12},(_,i)=><option key={i+1}>{i+1}</option>)}</select></div><div className="field"><label>السنة</label><select value={year} onChange={e=>setYear(e.target.value)}>{hijriYears.map(y=><option key={y} value={y}>{y}</option>)}</select></div><div className="field grow"><label>بحث داخل الحقول</label><input className="search" value={search} onChange={e=>setSearch(e.target.value)} placeholder="اكتب اسم المؤشر"/></div>{monthLocked&&<span className="badge badgeOk">الشهر معتمد ومغلق</span>}<button className="btn btnPrimary" onClick={saveCurrentForm} disabled={firebaseStatus==='loading'||(monthLocked&&user.role==='center')}>{firebaseStatus==='loading'?'جاري الحفظ...':'حفظ البيانات'}</button><button className="btn btnSoft" onClick={copyPreviousMonth}>نسخ الشهر السابق</button><button className="btn btnGhost" onClick={()=>exportModule(true)}>تصدير Excel</button><button className="btn btnDanger" onClick={clearCurrentForm}>مسح النموذج</button></div>{mod.id==='visitors'&&visitorsSummary&&<div className="calculatedPanel"><div><small>إجمالي المراجعين للمركز خلال الشهر</small><b>{visitorsSummary.monthlyTotal.toLocaleString('en-US')}</b></div><div><small>مجموع مراجعي العيادات</small><b>{visitorsSummary.clinicTotal.toLocaleString('en-US')}</b></div><div><small>مجموع الخدمات والمساندة</small><b>{visitorsSummary.supportTotal.toLocaleString('en-US')}</b></div></div>}<div className="approvalPanel"><div><small>حالة النموذج</small><b>{currentStatusLabel}</b></div><div><small>آخر حفظ</small><b>{currentLastSaveText}</b></div><div><small>تم بواسطة</small><b>{currentSavedBy}</b></div><div><small>نسبة الإكمال</small><b>{currentPct}%</b></div></div><div className="grid">{visibleFields.map((f:any)=><div className="entry" key={f.cell}><label>{cleanLabel(f.label,f.cell)}</label><input inputMode="decimal" pattern="[0-9]*" value={getVal(moduleId,f.cell)} onChange={e=>saveVal(f.cell,e.target.value)} placeholder="0" disabled={monthLocked&&user.role==='center'}/></div>)}</div></div></div>}
    {tab==='report'&&<div className="pageStack reportPage"><div className="welcomeCard executiveHero"><div><span className="officialTag dark">تقرير الإدارة التنفيذي</span><h2>التقرير التنفيذي للإحصائية الشهرية</h2><p>ملخص رسمي لحالة إدخال الإحصائية الشهرية للمراكز الصحية حسب الشهر والسنة المحددين.</p><div className="heroButtons"><button className="btn btnPrimary" onClick={printReport}>طباعة التقرير</button><button className="btn btnGhost" onClick={()=>setTab('followup')}>متابعة المراكز</button></div></div><div className="statusPill"><strong>{globalPct}%</strong><span>نسبة الإنجاز العامة</span></div></div>
     <div className="executiveIndicators"><div className="indicatorCard"><small>الفترة</small><strong>{month} / {year}</strong><span>الشهر والسنة الهجرية</span></div><div className="indicatorCard"><small>المراكز المكتملة</small><strong>{globalStats.completeCenters}</strong><span>من أصل {allowedCenters.length} مراكز</span></div><div className="indicatorCard"><small>المراكز غير المكتملة</small><strong>{incompleteCenters}</strong><span>تحتاج متابعة واستكمال</span></div><div className="indicatorCard"><small>النماذج المكتملة بكل المراكز</small><strong>{fullyCompletedModules}</strong><span>من أصل {modules.length} نموذجاً</span></div></div>
     <div className="dashboardGrid"><div className="card"><div className="sectionHead"><div><h2>ملخص المراكز الصحية</h2><p className="muted">ترتيب المراكز حسب نسبة الإنجاز للفترة المحددة.</p></div><span className={'badge '+(monthLocked?'badgeOk':'badgeWarn')}>{monthLocked?'الفترة معتمدة ومغلقة':'الفترة مفتوحة للإدخال'}</span></div><table className="table"><thead><tr><th>المركز</th><th>نسبة الإنجاز</th><th>النماذج المكتملة</th><th>آخر تحديث</th><th>الحالة</th></tr></thead><tbody>{executiveRows.map((r:any)=><tr key={r.center.id}><td>{r.center.name}</td><td><div className="progress"><i style={{width:`${r.stats.pct}%`}}></i></div><b>{r.stats.pct}%</b></td><td>{r.stats.completedForms} / {modules.length}</td><td>{r.latest?formatDateTime(r.latest.savedAt):'لا يوجد حفظ'}</td><td><span className={'badge '+(r.stats.pct===100?'badgeOk':r.stats.pct>0?'badgeWarn':'badgeLate')}>{r.stats.pct===100?'مكتمل':r.stats.pct>0?'قيد الاستكمال':'لم يبدأ'}</span></td></tr>)}</tbody></table></div>
      <div className="card"><div className="sectionHead"><div><h2>خلاصة تنفيذية</h2><p className="muted">قراءة مختصرة لحالة المنصة.</p></div></div><div className="reportSummary"><div><b>{globalPct}%</b><span>نسبة الإنجاز العامة</span></div><div><b>{startedModules}</b><span>نماذج بدأ إدخالها</span></div><div><b>{globalStats.done}</b><span>حقول مدخلة</span></div><div><b>{globalStats.total-globalStats.done}</b><span>حقول متبقية</span></div></div></div></div>
     <div className="card"><div className="sectionHead"><div><h2>مصفوفة الاعتماد التنفيذية</h2><p className="muted">حالة كل نموذج لكل مركز صحي.</p></div></div><div className="matrixWrap"><table className="matrix"><thead><tr><th>المركز</th>{modules.map(m=><th key={m.id}>{m.title}</th>)}</tr></thead><tbody>{allowedCenters.map(c=><tr key={c.id}><td>{c.name}</td>{modules.map(m=>{const f=formCompletion(m,c.id);return <td key={m.id}><span className={'dotStatus '+(f.pct===100?'done':f.pct>0?'progressing':'empty')}>{f.pct===100?'✓':f.pct>0?'◐':'—'}</span></td>})}</tr>)}</tbody></table></div></div>
    </div>}
    {tab==='exports'&&<div className="card"><h2>التصدير الشهري</h2><p className="muted">تصدير نموذج واحد من صفحة الإدخال أو تصدير جميع النماذج لجميع المراكز في ملف ZIP، مع استخدام القوالب الأصلية بعد تنظيف البيانات القديمة.</p><div className="toolbar"><div className="field"><label>الشهر</label><select value={month} onChange={e=>setMonth(e.target.value)}>{Array.from({length:12},(_,i)=><option key={i+1}>{i+1}</option>)}</select></div><div className="field"><label>السنة</label><select value={year} onChange={e=>setYear(e.target.value)}>{hijriYears.map(y=><option key={y} value={y}>{y}</option>)}</select></div><button className="btn btnPrimary" onClick={()=>exportModule(false)}>تصدير جميع النماذج ZIP</button></div></div>}
    {tab==='audit'&&<div className="pageStack"><div className="welcomeCard executiveHero"><div><span className="officialTag dark">سجل تدقيق</span><h2>سجل العمليات</h2><p>متابعة عمليات الدخول والحفظ واعتماد الشهور داخل المنصة.</p></div><div className="statusPill compact"><strong>{auditRows.length}</strong><span>عملية مسجلة</span></div></div><div className="card"><div className="sectionHead"><div><h2>آخر العمليات</h2><p className="muted">يعرض آخر 60 عملية محفوظة في Firebase.</p></div><button className="btn btnGhost" onClick={loadAuditRows}>تحديث السجل</button></div><table className="table"><thead><tr><th>الوقت</th><th>المستخدم</th><th>الإجراء</th><th>النطاق</th><th>التفاصيل</th></tr></thead><tbody>{auditRows.length?auditRows.map((r:any)=><tr key={r.id}><td>{formatDateTime(r.createdAtText)}</td><td>{r.userName||r.username||'-'}</td><td>{r.action}</td><td>{r.centerId||'عام'} / {r.year||'-'} / {r.month||'-'}</td><td>{r.details?.moduleTitle||r.details?.centerName||r.details?.username||'-'}</td></tr>):<tr><td colSpan={5}>لا توجد عمليات مسجلة حالياً.</td></tr>}</tbody></table></div></div>}
    {tab==='centers'&&<div className="pageStack"><div className="welcomeCard executiveHero"><div><span className="officialTag dark">إدارة النظام</span><h2>إدارة المراكز الصحية</h2><p>إضافة وتعديل وتفعيل وإيقاف المراكز الصحية من داخل المنصة، مع ربطها بالمستخدمين والصلاحيات.</p></div><div className="statusPill compact"><strong>{centerRows.filter((c:any)=>c.isActive!==false).length}</strong><span>مركز نشط</span></div></div><div className="kpis"><div className="kpi"><small>إجمالي المراكز</small><strong>{centerRows.length}</strong><span>مسجلة في النظام</span></div><div className="kpi"><small>المراكز النشطة</small><strong>{centerRows.filter((c:any)=>c.isActive!==false).length}</strong><span>متاحة للإدخال</span></div><div className="kpi"><small>المراكز الموقوفة</small><strong>{centerRows.filter((c:any)=>c.isActive===false).length}</strong><span>غير متاحة حالياً</span></div><div className="kpi"><small>المستخدمون المرتبطون</small><strong>{(userRows.length?userRows:users).filter((u:any)=>u.role==='center').length}</strong><span>مسؤولو إحصاء</span></div></div><div className="card"><div className="sectionHead"><div><h2>{editingCenterId?'تعديل مركز':'إضافة مركز جديد'}</h2><p className="muted">المراكز الأساسية مرتبطة بقوالب Excel، ويمكن تعديل الاسم أو الحالة. إضافة مركز جديد متاحة للإدارة عند الحاجة التشغيلية.</p></div><button className="btn btnGhost" onClick={resetCenterForm}>تفريغ النموذج</button></div><div className="toolbar"><div className="field"><label>رمز المركز</label><input value={centerForm.id} disabled={!!editingCenterId} onChange={e=>setCenterForm((x:any)=>({...x,id:e.target.value.trim().toLowerCase()}))} placeholder="مثال: c08"/></div><div className="field"><label>اسم المركز</label><input value={centerForm.name} onChange={e=>setCenterForm((x:any)=>({...x,name:e.target.value}))} placeholder="مثال: مركز صحي ..."/></div><div className="field"><label>القطاع</label><input value={centerForm.sector} onChange={e=>setCenterForm((x:any)=>({...x,sector:e.target.value}))} placeholder="شرق جدة"/></div><div className="field"><label>الحالة</label><select value={centerForm.isActive?'active':'inactive'} onChange={e=>setCenterForm((x:any)=>({...x,isActive:e.target.value==='active'}))}><option value="active">مفعل</option><option value="inactive">موقوف</option></select></div><button className="btn btnPrimary" onClick={saveCenter}>{editingCenterId?'حفظ التعديل':'إضافة المركز'}</button></div></div><div className="card"><div className="sectionHead"><div><h2>قائمة المراكز الصحية</h2><p className="muted">يمكن إيقاف المركز مؤقتاً دون حذف بياناته أو بيانات المستخدمين المرتبطين به.</p></div><button className="btn btnGhost" onClick={()=>{loadCenterRows();loadUserRows()}}>تحديث القائمة</button></div><table className="table"><thead><tr><th>الرمز</th><th>اسم المركز</th><th>القطاع</th><th>الحالة</th><th>المستخدمون</th><th>الإجراء</th></tr></thead><tbody>{centerRows.map((c:any)=>{const linked=(userRows.length?userRows:users).filter((u:any)=>u.centerId===c.id);return <tr key={c.id}><td>{c.id}</td><td>{c.name}</td><td>{c.sector||'شرق جدة'}</td><td><span className={'badge '+(c.isActive!==false?'badgeOk':'badgeLate')}>{c.isActive!==false?'مفعل':'موقوف'}</span></td><td>{linked.length?linked.map((u:any)=>u.username).join('، '):'لا يوجد'}</td><td><button className="btn btnGhost" onClick={()=>editCenter(c)}>تعديل</button> <button className="btn btnSoft" onClick={()=>toggleCenterActive(c)}>{c.isActive!==false?'إيقاف':'تفعيل'}</button> <button className="btn btnDanger" onClick={()=>removeCenter(c)}>حذف</button></td></tr>})}</tbody></table></div></div>}
    {tab==='users'&&<div className="pageStack"><div className="welcomeCard executiveHero"><div><span className="officialTag dark">إدارة النظام</span><h2>المستخدمون والصلاحيات</h2><p>إضافة وتعديل وتعطيل مستخدمي المنصة من داخل النظام، دون الحاجة للدخول إلى Firebase.</p></div><div className="statusPill compact"><strong>{userRows.length||users.length}</strong><span>مستخدم</span></div></div><div className="card"><div className="sectionHead"><div><h2>{editingUsername?'تعديل مستخدم':'إضافة مستخدم جديد'}</h2><p className="muted">هذه الصفحة متاحة لمدير النظام فقط.</p></div><button className="btn btnGhost" onClick={resetUserForm}>تفريغ النموذج</button></div><div className="toolbar"><div className="field"><label>اسم المستخدم</label><input value={userForm.username} disabled={!!editingUsername} onChange={e=>setUserForm((x:any)=>({...x,username:e.target.value.trim()}))} placeholder="مثال: c08"/></div><div className="field"><label>كلمة المرور</label><input value={userForm.password} onChange={e=>setUserForm((x:any)=>({...x,password:e.target.value}))} placeholder="كلمة المرور"/></div><div className="field"><label>اسم العرض</label><input value={userForm.name} onChange={e=>setUserForm((x:any)=>({...x,name:e.target.value}))} placeholder="اسم المستخدم أو المركز"/></div><div className="field"><label>الدور</label><select value={userForm.role} onChange={e=>setUserForm((x:any)=>({...x,role:e.target.value as Role,centerId:e.target.value==='center'?x.centerId:''}))}><option value="admin">مدير النظام</option><option value="executive">إدارة شؤون المراكز الصحية</option><option value="center">مسؤول إحصاء مركز</option></select></div>{userForm.role==='center'&&<div className="field"><label>المركز</label><select value={userForm.centerId} onChange={e=>setUserForm((x:any)=>({...x,centerId:e.target.value}))}>{activeCenters.map((c:any)=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>}<div className="field"><label>الحالة</label><select value={userForm.isActive?'active':'inactive'} onChange={e=>setUserForm((x:any)=>({...x,isActive:e.target.value==='active'}))}><option value="active">مفعل</option><option value="inactive">موقوف</option></select></div><button className="btn btnPrimary" onClick={saveUser}>{editingUsername?'حفظ التعديل':'إضافة المستخدم'}</button></div></div><div className="card"><div className="sectionHead"><div><h2>قائمة المستخدمين</h2><p className="muted">يمكن تعديل كلمة المرور، الدور، النطاق، أو إيقاف المستخدم دون حذف بياناته السابقة.</p></div><button className="btn btnGhost" onClick={loadUserRows}>تحديث القائمة</button></div><table className="table"><thead><tr><th>المستخدم</th><th>الاسم</th><th>الدور</th><th>النطاق</th><th>الحالة</th><th>كلمة المرور</th><th>الإجراء</th></tr></thead><tbody>{(userRows.length?userRows:users).map((x:any)=><tr key={x.username}><td>{x.username}</td><td>{x.name}</td><td>{roleTitle(x.role)}</td><td>{x.centerId?centerName(x.centerId):'جميع المراكز'}</td><td><span className={'badge '+(x.isActive!==false?'badgeOk':'badgeLate')}>{x.isActive!==false?'مفعل':'موقوف'}</span></td><td>{x.password}</td><td><button className="btn btnGhost" onClick={()=>editUser(x)}>تعديل</button> <button className="btn btnSoft" onClick={()=>toggleUserActive(x)}>{x.isActive!==false?'إيقاف':'تفعيل'}</button> <button className="btn btnDanger" onClick={()=>removeUser(x)}>حذف</button></td></tr>)}</tbody></table></div></div>}
   </section></main></>
}