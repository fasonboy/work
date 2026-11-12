// Rebuild the four market research screenshots with native charts and a comparison table.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/market-research',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation,applyPresentationChartFont}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1677,H=938,FONT='Noto Sans CJK SC',NAVY='#071b45',BLUE='#074ee7',GRAY='#63748d';
if(process.env.PPT_OPERATION_MARKED!=='1') execFileSync(process.env.RUNTIME_NODE,[path.join(skill,'container_tools/mark_artifact_operation_started.mjs'),'--operation-kind','create','--expected-output-count','1','--output-format','pptx'],{stdio:'inherit'});
const deck=Presentation.create({slideSize:{width:W,height:H}});let slide;
const grad=(a,b,angleDeg=40)=>({type:'gradient',gradientKind:'linear',angleDeg,stops:[{offset:0,color:a},{offset:100000,color:b}]});
function box(name,x,y,w,h,fill='none',r=0,geometry='rect',stroke='none',lineWidth=1) {
  return slide.shapes.add({name,geometry,position:{left:x,top:y,width:w,height:h},fill,line:{fill:stroke,width:stroke==='none'?0:lineWidth},...(r&&['rect','textbox','roundRect'].includes(geometry)?{borderRadius:r}:{})});
}
function txt(name,value,x,y,w,h,size=24,color=NAVY,bold=false,extra={}) {
  const o=box(name,x,y,w,h);o.text=value;o.text.style={typeface:FONT,fontSize:size,color,bold,autoFit:'none',insets:0,wrap:'none',verticalAlignment:'middle',...extra};return o;
}
const M=(x,y)=>({moveTo:{x,y}}),L=(x,y)=>({lineTo:{x,y}}),B=(x1,y1,x2,y2,x,y)=>({cubicBezTo:{x1,y1,x2,y2,x,y}}),Z=()=>({close:{}});
function curve(name,x,y,w,h,commands,fill,stroke='none',lineWidth=1) {
  return slide.shapes.add({geometry:'custom',name,position:{left:x,top:y,width:w,height:h},fill,line:{fill:stroke,width:stroke==='none'?0:lineWidth},customPaths:[{width:w,height:h,commands}]});
}
function poly(name,pts,fill) {
  const x=Math.min(...pts.map(p=>p[0])),y=Math.min(...pts.map(p=>p[1])),w=Math.max(...pts.map(p=>p[0]))-x,h=Math.max(...pts.map(p=>p[1]))-y;
  return curve(name,x,y,w,h,[M(pts[0][0]-x,pts[0][1]-y),...pts.slice(1).map(p=>L(p[0]-x,p[1]-y)),Z()],fill);
}
const paths={
 trend:'<path d="M6 32L24 17l13 10L63 4M48 4h15v15M7 64V47h11v17zM29 64V37h11v27zM51 64V26h11v38z"/>',
 cart:'<path d="M5 5h10l9 39h32l9-28H19M24 53h31"/><circle cx="28" cy="60" r="4"/><circle cx="53" cy="60" r="4"/>',
 box:'<path d="M6 18L33 4l28 14v32L33 65 6 50zM6 18l27 15 28-15M33 33v32M20 11l28 15"/>',
 fish:'<path d="M4 34c18-22 37-22 52-7l14-8-3 15 3 15-14-8C41 56 22 56 4 34z"/><circle cx="20" cy="31" r="2"/>',
 rocket:'<path d="M22 45C29 13 43 3 64 3c0 21-10 35-42 42zM23 29l-13 4-7 16 17-3M38 43l-3 17-16 7 5-15M14 52L4 64"/><circle cx="47" cy="20" r="6"/>',
 diamond:'<path d="M3 21L18 5h33l15 16-31 44zM3 21h63M18 5l17 60L51 5M18 5l17 16L51 5"/>',
 shield:'<path d="M34 4l26 10v20c0 15-13 26-26 32C21 60 8 49 8 34V14zM22 34l10 10 18-22"/>',
 bulb:'<path d="M17 45C-4 9 58-3 57 29c0 8-4 14-10 19v9H23v-9zM24 62h22M28 67h14M35 3V0M4 27H0M65 27h5M11 9L6 4M59 9l5-5"/>',
 molecules:'<path d="M34 7L13 20v26l21 13 22-13V20zM34 7v-5M13 20l-8-4M56 20l9-4M13 46l-8 5M56 46l9 5M34 59v7"/><circle cx="34" cy="5" r="4"/><circle cx="9" cy="18" r="4"/><circle cx="59" cy="18" r="4"/><circle cx="9" cy="48" r="4"/><circle cx="59" cy="48" r="4"/>',
 leaf:'<path d="M6 51C9 13 36 20 61 4c0 39-18 57-46 51M4 65l40-40"/>',
 layers:'<path d="M5 21L34 5l29 16-29 16zM6 34l28 16 28-16M6 47l28 16 28-16"/>',
 microscope:'<path d="M22 8L35 0l22 33-13 8zM42 42C24 53 12 40 16 28M35 49v13M8 66h56M45 43c17 0 22-15 9-27M5 52h27"/>',
 target:'<circle cx="31" cy="36" r="27"/><circle cx="31" cy="36" r="17"/><circle cx="31" cy="36" r="7"/><path d="M31 36L60 7M53 1l-2 13 13-2"/>',
 user:'<circle cx="34" cy="17" r="13"/><path d="M5 64c0-36 58-36 58 0"/>',
 people:'<circle cx="34" cy="15" r="11"/><path d="M13 62V49c0-22 42-22 42 0v13zM12 8c-13 0-13 21 0 21M56 8c13 0 13 21 0 21M8 38c-9 0-11 9-11 20M60 38c9 0 11 9 11 20"/>',
 coins:'<ellipse cx="45" cy="13" rx="17" ry="7"/><path d="M28 13v28c0 11 34 11 34 0V13M28 24c0 10 34 10 34 0M28 35c0 10 34 10 34 0"/><ellipse cx="21" cy="39" rx="17" ry="7"/><path d="M4 39v17c0 11 34 11 34 0V39M4 48c0 11 34 11 34 0"/>',
 stethoscope:'<path d="M5 7v22c0 22 31 22 31 0V7M5 8h7M29 8h7M21 47v10c0 22 40 22 40-3V39"/><circle cx="61" cy="33" r="8"/>',
 video:'<rect x="5" y="8" width="59" height="47" rx="4"/><path d="M29 21l16 10-16 12z"/>',
 head:'<path d="M24 65V51C-7 28 19-2 45 6c22 7 24 26 13 40v19zM21 17h2M37 16h2M49 28h2M21 33h2M34 28h2M39 42h2"/>',
 lotus:'<path d="M34 62C7 60 4 39 5 29c17 0 28 16 29 33zM34 62c27-2 30-23 29-33-17 0-28 16-29 33zM34 61C12 39 23 17 34 4c14 13 24 34 0 57zM4 47c-10 16 22 28 30 15 10 13 42 0 31-15"/>',
 house:'<path d="M4 31L34 3l30 28H54v34H39V47H27v18H13V31z"/>',
 ship:'<path d="M4 40h60L53 59H16zM14 40V19h38v21M25 19V7h15v12M34 7V0M19 30h7M33 30h7M46 30h5M4 66c9 5 14-4 22 0s14 4 22 0 13 4 20 0"/>',
 factory:'<path d="M6 64V32l16 10V28l16 12V20h20v44zM11 31V7h10v29M42 20V7h13v13M16 52h5M30 52h5M45 52h5"/>',
 building:'<path d="M5 64V17l46-9v56M51 29h13v35M13 26h7M28 23h7M13 39h7M28 36h7M13 52h7M28 49h7M2 65h66"/>',
 lab:'<path d="M40 4l22 22-8 8L32 12zM36 16L6 49c-9 15 7 24 18 12l29-32M12 47h25M50 5l12 12"/>',
};
function icon(kind,x,y,size,color=BLUE){assert(paths[kind],kind);slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="84" height="84" viewBox="-7 -5 84 84"><g fill="none" stroke="${color}" stroke-width="3.8" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} icon`,position:{left:x,top:y,width:size,height:size}});}
function newSlide(note,bg='#ffffff'){slide=deck.slides.add();slide.background.fill=bg;slide.speakerNotes.text=note;}
function panel(name,x,y,w,h,fill='#ffffff',r=24){box(`${name}-shadow`,x,y+5,w,h,'#d9e4f2/55',r);return box(name,x,y,w,h,fill,r,'rect','#edf3fa',1);}
function header(title,subtitle,first=false){
 txt('market-research-label',first?'MARKET RESEARCH  /':'MARKET RESEARCH  /  南极磷虾油项目',67,20,first?225:800,37,21,first?GRAY:BLUE,true);
 if(first)txt('project-label','南极磷虾油项目',290,20,350,37,21,'#0d4b88',true);
 txt('page-title',title,67,62,1562,68,45,NAVY,true);txt('page-subtitle',subtitle,67,132,1547,39,21.5,first?GRAY:'#3c536f');
 box('header-rule',67,184,1549,5,'#edf0f4');box('header-accent',67,184,95,5,'#105bcb');
}
function sourceNote(t){return '来源：用户本轮按发送顺序提供的4张市场调研截图。'+t+' 仅按截图转录与重建；市场数据、研究结论、公司能力和来源归属未独立核验，不补造文献链接或原始材料。图标及装饰为近似重绘，无整页截图背景。';}
function check(name,x,y,size=26){box(name,x,y,size,size,'#085ce0',0,'ellipse');curve(`${name}-tick`,x+size*.24,y+size*.27,size*.55,size*.45,[M(0,size*.2),L(size*.17,size*.37),L(size*.52,0)],'none','#ffffff',2.5);}

// Slide 1: market metrics and a native chart. Preserve the derived-2025 note.
newSlide(sourceNote('第1张：2026上半年市场数据。2025 H1约3.5亿元为按5.5亿元与56.9%同比倒算的一位小数近似值，非单独公布数据；柱图使用截图已显示的3.5/5.5，不以四舍五入的柱值重算同比。'),grad('#f6fbff','#e2f4fd',80));
box('polar-sea',0,886,W,52,grad('#93cde8','#d7f2fc',90));
poly('polar-left-ice',[[0,798],[44,825],[25,836],[101,866],[193,881],[0,888]],'#ddf2ff');poly('polar-right-ice',[[1555,810],[1605,767],[1677,790],[1677,899],[1512,890]],'#b9e1f7');
for(let i=0;i<7;i++)box(`sea-ripple-${i}`,32+(i%3)*44,894+i*6,Math.min(1580-i*47,W-40-(i%3)*44),1.3,i%2?'#effcff':'#b4e0f4');
header('中国磷虾油市场进入高增长阶段，品类渗透仍有空间','2026年上半年主流电商平台监测：磷虾油增速领先，鱼油仍保持更大的销售规模',true);
const kpis=[{x:72,w:377,label:'磷虾油线上销售额 | 2026 H1',value:'5.5亿元',detail:'同比增长 56.9% ↑',kind:'trend'},{x:480,w:354,label:'磷虾油销售增速',value:'+56.9%',detail:'约为保健食品线上大盘增速的6倍',kind:'cart'},{x:859,w:369,label:'升级型磷虾油商品 | 2026 H1',value:'2.4亿元',detail:'同比增长 102.5% ↑',kind:'box'},{x:1253,w:369,label:'鱼油及 Omega-3 | 2026 H1',value:'29.1亿元',detail:'同比增长 41.1% ↑',kind:'fish'}];
for(const [i,k] of kpis.entries()){
 box(`kpi-${i+1}-accent`,k.x-5,225,7,159,grad('#1259f2','#73cdf7'),4);panel(`kpi-${i+1}`,k.x,202,k.w,187,grad('#ffffff','#eaf5ff'),19);
 box(`kpi-${i+1}-badge`,k.x+25,221,46,46,grad('#78b8ff','#1e4eea'),0,'ellipse');icon(k.kind,k.x+34,230,29,'#ffffff');
 txt(`kpi-${i+1}-label`,k.label,k.x+93,226,k.w-106,33,18.5,'#3c5474',true);txt(`kpi-${i+1}-value`,k.value,k.x+93,267,k.w-107,63,46,i===1?'#0545fb':NAVY,true);
 if(i!==1)box(`kpi-${i+1}-detail-bg`,k.x+92,335,165,31,'#dfedff',16);
 txt(`kpi-${i+1}-detail`,k.detail,k.x+99,333,k.w-113,34,i===1?16.2:17,'#2d5489',true);
}
panel('sales-panel',66,411,792,450,grad('#ffffff','#f6fbff'),23);
txt('sales-heading','磷虾油线上销售额持续扩大',93,426,697,46,29,NAVY,true);txt('sales-unit','单位：亿元 | 主流电商平台监测',93,471,697,31,18,GRAY);
const salesChart=slide.charts.add('bar',{position:{left:96,top:513,width:708,height:288},categories:['2025 H1*','2026 H1'],series:[{name:'磷虾油线上销售额（亿元）',values:[3.5,5.5],points:[{idx:0,fill:grad('#becdfa','#a6baf1',90)},{idx:1,fill:grad('#2a8de4','#0c53b4',90)}]}],barOptions:{direction:'column',grouping:'clustered',gapWidth:130},hasLegend:false,titlePlacement:'none',xAxis:{visible:true,position:'bottom',line:{fill:'#bac8da',width:1.3},textStyle:{typeface:FONT,fontSize:19,fill:'#415979'},majorGridlines:null,minorGridlines:null},yAxis:{visible:true,min:0,max:6,majorUnit:2,numberFormatCode:'0',line:{fill:'none',width:0},textStyle:{typeface:FONT,fontSize:16,fill:'#63748d'},majorGridlines:{fill:'#e7edf5',width:1},minorGridlines:null},dataLabels:{showValue:true,position:'outEnd',numberFormatCode:'0.0',textStyle:{typeface:FONT,fontSize:25,fill:NAVY,bold:true}},chartFill:'none',chartLine:{fill:'none',width:0},plotAreaFill:'none',plotAreaLine:{fill:'none',width:0}});applyPresentationChartFont(salesChart,{fontFamily:FONT});
box('sales-yoy-badge',387,503,131,35,'#d9e7ff',18);txt('sales-yoy-label','+56.9% YoY',400,507,111,28,18,BLUE,true);
const arrowA=box('sales-growth-from',370,557,1,1),arrowB=box('sales-growth-to',524,542,1,1);slide.shapes.connect(arrowA,arrowB,{kind:'straight',line:{fill:grad('#d9eaff','#4085df'),width:6},head:{type:'none'},tail:{type:'triangle',width:'sm',length:'sm'}}).bringToFront();
txt('sales-derived-note','*2025 H1约3.5亿元：按2026 H1销售额及同比增速倒算，非单独公布值。',93,811,731,35,14,'#506782');
panel('trend-panel',880,411,742,450,grad('#ffffff','#f2faff'),24);box('trend-heading-badge',912,432,48,48,grad('#47acfa','#0c77dd'),0,'ellipse');icon('rocket',923,442,29,'#ffffff');txt('trend-heading','市场发展趋势',981,431,576,51,31,NAVY,true);
const trendRows=[{y:503,n:'01',kind:'trend',title:'品类增长快于大盘',detail:'磷虾油增速56.9%，高于鱼油41.1%及保健食品大盘9.5%。'},{y:600,n:'02',kind:'diamond',title:'产品向高规格升级',detail:'突出磷脂、虾青素等明确成分卖点的升级型商品增长更快。'},{y:697,n:'03',kind:'shield',title:'竞争转向品质与信任',detail:'原料来源、成分含量、批次检测、资质和生产信息透明度成为关键。'}];
for(const [i,r] of trendRows.entries()){
 box(`trend-${i+1}-number-bg`,915,r.y+17,42,42,grad('#67bafa','#1688db'),0,'ellipse');txt(`trend-${i+1}-number`,r.n,922,r.y+22,30,31,22,'#ffffff',true,{alignment:'center'});
 box(`trend-${i+1}-badge`,987,r.y,61,61,'#e7f5ff',0,'ellipse');icon(r.kind,1000,r.y+12,37,'#4a93d0');
 txt(`trend-${i+1}-heading`,r.title,1072,r.y-3,502,36,21,NAVY,true);txt(`trend-${i+1}-detail`,r.detail,1072,r.y+34,503,53,17.2,GRAY,false,{wrap:'square',lineSpacing:1.15});
 if(i<2)box(`trend-${i+1}-rule`,984,r.y+79,589,1,'#e4edf5');
}
box('market-judgment-bg',911,790,681,52,'#def2ff',26);icon('bulb',924,802,29,'#2b8add');txt('market-judgment','市场判断：小品类高增长，增长动力来自渗透率提升与高规格产品升级。',965,800,608,34,17.6,'#0952a3',true);

// Slide 2: actual comparison table and the contrasting study results.
newSlide(sourceNote('第2张：营养对比与人体吸收研究。2026 RCT（72名健康成年人、12周、Omega-3 1.1g/日）和2015 RCT（66名健康成年人、4周、Omega-3 1.3g/日）及NIH综述措辞只作截图转录，未核验原始文献。保留“观察血液指标变化，不等于临床疾病结局”及不能宣传为药物替代品的结论。'));
header('磷虾油更可能成为鱼油的高端补充与结构性替代','两者都提供 EPA / DHA，核心差异在脂质形态、伴随营养成分及实际利用效果');
panel('comparison-panel',66,221,820,559,'#ffffff',13);
const compareValues=[
 ['比较维度','南极磷虾油','鱼油'],
 ['Omega-3\n形态','以磷脂结合型为主\n（更易被细胞利用）','以甘油三酯型为主\n（也有乙酯型、再酯化型）'],
 ['伴随成分','含天然磷脂、虾青素\n及胆碱等多种活性营养','主要提供 EPA / DHA\n成分相对单一'],
 ['有效含量','单位油脂中的 EPA / DHA\n通常更高','EPA / DHA 含量相对较低\n需更高剂量才能达到相似摄入'],
 ['市场定位','差异化、高端化\n强调复合营养与原料稀缺性','成熟大众品类\n价格带宽、剂量选择多'],
];
const compareTable=slide.tables.add({rows:5,columns:3,left:66,top:221,width:820,height:538,columnWidths:[235,315,270],values:compareValues});
for(let r=0;r<5;r++){
 compareTable.rows[r].height=[65,118,118,119,118][r];
 for(let c=0;c<3;c++){
  const cell=compareTable.getCell(r,c);cell.fill=r===0?grad('#0866dc','#0752be'):(c===1?'#eaf3ff':r%2?'#f8fafc':'#ffffff');
  cell.text.style={typeface:FONT,fontSize:r===0?27:c===0?24:c===1?20:18.5,color:r===0?'#ffffff':NAVY,bold:r===0||c<2,autoFit:'none',wrap:'square',alignment:r===0?'center':'left',verticalAlignment:'middle',lineSpacing:1.1};
  compareTable.cells.block({row:r,column:c,rowCount:1,columnCount:1}).assign({margins:{left:r===0?20:c===0?103:c===1?61:21,right:9,top:8,bottom:8}});
 }
}
compareTable.borders.assign({fill:'#e9eef5',color:'#e9eef5',width:1,style:'solid'});
for(const [i,kind,y] of [[1,'molecules',325],[2,'leaf',443],[3,'layers',561],[4,'trend',679]]){icon(kind,94,y,47,'#075cbd');check(`compare-check-${i}`,324,y+3,25);}
panel('study-panel',914,206,703,574,'#ffffff',13);box('study-icon-badge',938,224,55,55,'#0962dc',12);icon('microscope',949,234,34,'#ffffff');txt('study-heading','人体吸收研究：显示两者存在显著差异',1018,229,574,50,29,NAVY,true);
panel('rct-2026-panel',938,302,657,144,'#edf4ff',12);poly('rct-2026-tag-bg',[[963,303],[1101,303],[1090,322],[1101,341],[963,341]],'#2f79fd');txt('rct-2026-tag','2026 RCT',976,306,119,33,21,'#ffffff',true);
txt('rct-2026-method','72名健康成年人 | 12周 | Omega-3 1.1g/日',1129,305,449,35,19.2,'#596b80',true);
txt('rct-2026-result','磷虾油组血浆 EPA / DHA 增幅约为鱼油组的',969,350,478,55,23.5,NAVY,true);txt('rct-2026-ratio','1.5 倍',1454,344,129,65,43,BLUE,true,{alignment:'center'});
txt('rct-2026-limit','观察的是血液指标变化，不等于临床疾病结局。',969,401,600,34,18.5,'#53657c');
panel('rct-2015-panel',938,476,657,105,'#ffffff',11);poly('rct-2015-tag-bg',[[963,477],[1101,477],[1090,497],[1101,517],[963,517]],'#f2f4f7');txt('rct-2015-tag','2015 RCT',976,480,121,33,21,NAVY,true);
txt('rct-2015-method','66名健康成年人 | 4周 | Omega-3 1.3g/日',1129,480,449,35,19.2,'#596b80',true);txt('rct-2015-result','血浆及红细胞 EPA+DHA 水平未见显著组间差异',970,525,608,46,25.5,NAVY,true);
txt('nih-summary','NIH综述：部分研究提示磷脂型 Omega-3 生物利用度较高，另一些研究未发现明显差异；仍需结合剂量和具体配方判断。',938,600,657,65,19.5,'#40546e',false,{wrap:'square',lineSpacing:1.18});
box('study-rule',938,674,655,2,'#edf0f3');panel('substitution-logic-panel',938,695,657,66,'#ffffff',12);box('substitution-label-bg',939,698,128,52,'#e9f2ff',10);txt('substitution-label','替代逻辑',953,704,108,41,25,BLUE,true);
txt('substitution-logic','鱼油保留大众与高浓度市场；磷虾油争取注重磷脂、虾青素、品质溯源的中高端消费者。',1099,699,479,58,18.5,NAVY,true,{wrap:'square',lineSpacing:1.15});
panel('comparison-conclusion-bg',66,810,1551,75,grad('#eff7ff','#e9f2ff'),17);poly('comparison-conclusion-label-bg',[[82,810],[279,810],[246,885],[82,885],[66,870],[66,827]],grad('#0b66d1','#0747a6'));icon('target',89,827,45,'#ffffff');txt('comparison-conclusion-label','结论',148,824,104,51,35,'#ffffff',true);
txt('comparison-conclusion','磷虾油不是鱼油的全面替代，而是 Omega-3 市场中的高端升级选项；不能宣传为药物替代品。',303,822,1284,54,29,NAVY,true);

// Native percentage charts use only observed values. Gray rails are drawing shapes.
function percentageBars(name,categories,values,x,y,w,h,cap,color=BLUE,size=25){
 const fills=values.map((_,i)=>({idx:i,fill:Array.isArray(color)?color[i]:color}));
 const gap=values.length===4?170:245,rowHeight=(h-8)/values.length,barHeight=rowHeight/(1+gap/100);
 for(let i=0;i<values.length;i++)box(`${name}-display-rail-${i+1}`,x+9,y+4+rowHeight*(i+.5)-barHeight/2,w-18,barHeight,'#e9eef4');
 const chart=slide.charts.add('bar',{position:{left:x,top:y,width:w,height:h},categories,series:[
  {name:`${name}（实际比例）`,values,points:fills,dataLabels:{showValue:true,position:'outEnd',numberFormatCode:'0%',textStyle:{typeface:FONT,fontSize:size,fill:name==='购买信任因素'?NAVY:BLUE,bold:true}}},
 ],barOptions:{direction:'bar',grouping:'clustered',gapWidth:gap},hasLegend:false,titlePlacement:'none',xAxis:{visible:false,line:{fill:'none',width:0},tickLabelPosition:'none',majorGridlines:null,minorGridlines:null},yAxis:{visible:false,min:0,max:cap,line:{fill:'none',width:0},tickLabelPosition:'none',majorGridlines:null,minorGridlines:null},chartFill:'none',chartLine:{fill:'none',width:0},plotAreaFill:'none',plotAreaLine:{fill:'none',width:0}});applyPresentationChartFont(chart,{fontFamily:FONT});return chart;
}
// Slide 3: keep fish-oil buyer proxies distinct from krill survey denominators.
newSlide(sourceNote('第3张：顶部57%、31%、31%、42%及信任因素38%、29%、11%、3%均为鱼油消费者/相邻品类购买力代理，不能当成磷虾油消费者直接测量。右侧中国消费者调查2024、N=1,000，29%为全部受访者听说过，48%仅在已认知人群中，76%仅在磷虾油使用者中。两张条形图灰色底轨分别按50%和100%作视觉标尺；底轨为形状，并非额外调查数据，原生图表及内嵌工作簿仅存实际比例。'));
header('消费人群不是单一银发市场，核心在中年健康管理者与家庭决策者','以鱼油消费者调研作为相邻品类购买力代理，并结合磷虾油认知调查识别市场教育空间');
const proxyMetrics=[{x:56,w:365,title:'女性消费者占比',value:'57%',detail:'鱼油线上购买者',kind:'user'},{x:439,w:381,title:'36–45岁年龄段',value:'31%',detail:'鱼油线上购买者',kind:'people'},{x:841,w:384,title:'年收入8–12万元',value:'31%',detail:'鱼油线上购买者',kind:'coins'},{x:1244,w:381,title:'单次消费100–300元',value:'42%',detail:'501元以上约13%',kind:'cart'}];
for(const [i,m] of proxyMetrics.entries()){
 box(`proxy-${i+1}-accent`,m.x,207,7,143,'#1d65ff',4);panel(`proxy-${i+1}`,m.x+7,206,m.w-7,145,grad('#ffffff','#edf4ff'),17);box(`proxy-${i+1}-badge`,m.x+23,239,77,77,'#e8f2ff',0,'ellipse');icon(m.kind,m.x+40,257,44,'#2169fd');box(`proxy-${i+1}-divider`,m.x+115,239,1,77,'#dce9fb');
 txt(`proxy-${i+1}-title`,m.title,m.x+133,216,m.w-146,34,21,NAVY,true);txt(`proxy-${i+1}-value`,m.value,m.x+133,248,m.w-146,65,58,'#071460',true);txt(`proxy-${i+1}-detail`,m.detail,m.x+133,309,m.w-146,34,20,'#354f7c');
}
panel('trust-panel',55,373,792,354,'#ffffff',25);box('trust-heading-badge',79,391,64,64,grad('#4193ff','#1557df'),12);icon('cart',90,403,42,'#ffffff');txt('trust-heading','购买信任因素',160,391,637,45,31,NAVY,true);txt('trust-subtitle','专业人士与健康垂直内容对购买决策影响较大',160,434,637,31,18,'#3e5576');
const trustLabels=['认证医生 / 营养师','健康垂直大V','真实用户分享','自媒体评测大V'],trustValues=[.38,.29,.11,.03];
for(const [i,kind] of ['stethoscope','shield','people','video'].entries()){icon(kind,88,490+i*58,34,'#236afe');txt(`trust-label-${i+1}`,trustLabels[i],139,487+i*58,171,40,20,NAVY,true);}
percentageBars('购买信任因素',trustLabels,trustValues,319,480,480,232,.5,['#0b32c7','#0a3fd0','#82aaff','#82aaff'],23);
panel('awareness-panel',867,373,758,354,'#ffffff',25);box('awareness-heading-badge',891,391,64,64,grad('#4193ff','#1557df'),12);icon('head',901,401,44,'#ffffff');txt('awareness-heading','磷虾油认知与接受度',977,391,610,45,31,NAVY,true);txt('awareness-source','中国消费者调查 | 2024 | 样本 N=1,000',977,434,610,31,19,'#3e5576');
const awareness=[{title:'听说过磷虾油',base:'全部受访者',value:.29},{title:'已认知者正在使用磷虾油',base:'仅在已认知人群中',value:.48},{title:'认为磷虾油是优质 Omega-3 来源',base:'仅在磷虾油使用者中',value:.76}];
for(const [i,a] of awareness.entries()){txt(`awareness-${i+1}-label`,a.title,902,479+i*79,324,40,20.5,NAVY,true);txt(`awareness-${i+1}-base`,a.base,902,514+i*79,324,31,19,'#3d5579');}
percentageBars('磷虾油认知与接受度',awareness.map(a=>a.title),awareness.map(a=>a.value),1234,480,360,228,1,BLUE,28);
const segments=[{x:55,w:527,title:'35–55岁 | 健康管理者',detail:'核心客群：36–45岁是鱼油购买者中占比最高的年龄段',kind:'user'},{x:600,w:509,title:'25–40岁 | 品质营养人群',detail:'增量客群：关注配方、便捷体验与科学内容',kind:'lotus'},{x:1128,w:497,title:'55+ | 家庭健康购买者',detail:'家庭场景：重视安全、说明清晰与服务保障',kind:'house'}];
for(const [i,g] of segments.entries()){panel(`audience-${i+1}`,g.x,746,g.w,119,grad('#ffffff','#edf6ff'),20);icon(g.kind,g.x+33,775,62,'#2671ff');box(`audience-${i+1}-divider`,g.x+115,774,1.2,64,'#d2e4fb');txt(`audience-${i+1}-heading`,g.title,g.x+133,766,g.w-146,43,27,'#071c68',true);txt(`audience-${i+1}-detail`,g.detail,g.x+133,811,g.w-147,43,18,'#3d557a',false,{wrap:'square',lineSpacing:1.15});}

// Slide 4: production landscape. Unavailable photographs are not fabricated.
newSlide(sourceNote('第4张：生产格局。公司名称、产能、许可、合作关系和进入门店数量仅为截图转录，未验证企业原始公告、SKU标签或备案资料。当前云端没有原照片，四块照片区域改为明确标注“非原图／原图待补”的通用矢量示意，不能作为企业实际设施或船舶照片。'));
header('国内生产端逐步形成原料、精深加工与终端品牌协同格局','品牌方、原料供应商与实际生产企业可能不是同一主体，具体关系须落实到SKU标签与备案资料');
panel('upstream-banner',58,211,1566,68,'#edf5ff',16);poly('upstream-label-bg',[[75,211],[298,211],[279,279],[75,279],[58,266],[58,228]],grad('#287bd1','#1456a0'));icon('ship',90,228,41,'#ffffff');txt('upstream-label','上游原料',145,220,143,51,29,'#ffffff',true);box('upstream-divider',320,228,1.2,34,'#bad4f0');
txt('upstream-body','挪威 Aker BioMarine | 全球南极磷虾原料供应商之一，与VIK等品牌在原料及研发合作；原料端优势集中在捕捞、加工与供应链管理。',346,218,1243,54,21,NAVY,false,{wrap:'square',lineSpacing:1.1});
const companies=[
 {x:48,y:310,w:782,h:213,title:'辽渔南极磷虾科技发展有限公司',sub:'辽渔集团全资公司 | 陆上精深加工与销售',body:'产品覆盖磷虾油、凝胶糖果、磷虾肽、磷虾粉。2026年公开首发60%磷脂磷虾油及“极之极”磷虾油番茄红素软胶囊；原料来自“福兴海”轮捕捞并船载加工。',kind:'ship',caption:'船舶示意（非原图）'},
 {x:848,y:310,w:783,h:213,title:'逢时（青岛）海洋科技股份有限公司',sub:'VIK品牌 | 品牌运营与自有生产体系',body:'布局透明工厂，与挪威Aker BioMarine开展原料及产品合作；企业公开披露旗下业务进入2,100余家线下门店，覆盖会员商超及药房。',kind:'factory',caption:'车间示意（非原图）'},
 {x:48,y:542,w:782,h:201,title:'山东鲁华海洋生物科技有限公司',sub:'山东健康集团旗下 | 专业磷虾油生产商',body:'官方资料披露年产能约500吨，产品覆盖南极磷虾油、凝胶糖果、膏剂、微囊粉、片剂和滴剂等形态。',kind:'building',caption:'厂房示意（非原图）'},
 {x:848,y:542,w:783,h:201,title:'青岛南极维康生物科技有限公司',sub:'研发、原料提取与终端生产',body:'公开企业资料披露拥有SC生产许可、10万级GMP车间及年产1,000吨磷虾油原料生产线，具备凝胶糖果等终端产品生产能力。',kind:'lab',caption:'实验示意（非原图）'},
];
for(const [i,c] of companies.entries()){
 panel(`company-${i+1}`,c.x,c.y,c.w,c.h,'#ffffff',25);box(`company-${i+1}-image-placeholder`,c.x+17,c.y+16,198,c.h-32,grad('#eff8ff','#d6eaf7'),17);icon(c.kind,c.x+65,c.y+37,103,'#2b83ba');
 txt(`company-${i+1}-image-caption`,c.caption,c.x+26,c.y+c.h-55,181,24,15,'#416681',true,{alignment:'center'});txt(`company-${i+1}-image-pending`,'原图待补',c.x+26,c.y+c.h-33,181,22,14,'#6e8799',false,{alignment:'center'});
 box(`company-${i+1}-divider`,c.x+237,c.y+22,1.5,c.h-43,'#367dff');txt(`company-${i+1}-heading`,c.title,c.x+261,c.y+10,c.w-281,48,i===1?28:29,NAVY,true);
 txt(`company-${i+1}-subtitle`,c.sub,c.x+261,c.y+59,c.w-280,40,21,BLUE,true);txt(`company-${i+1}-body`,c.body,c.x+261,c.y+111,c.w-280,c.h-119,18.5,NAVY,false,{wrap:'square',lineSpacing:1.1});
}
panel('project-impact-bg',48,778,1582,111,'#edf5ff',23);poly('project-impact-label-bg',[[73,778],[370,778],[337,889],[73,889],[48,868],[48,799]],grad('#2b79cc','#1755a0'));icon('trend',78,813,47,'#ffffff');txt('project-impact-label','对本项目的意义',134,803,218,62,29,'#ffffff',true);box('project-impact-divider',380,810,1.3,44,'#b4ceeb');
const impacts=[{x:404,w:383,text:'具备捕捞、船载加工与陆上精深加工的产业基础'},{x:852,w:346,text:'在研发、生产与保健食品开发上具备协同优势'},{x:1248,w:352,text:'项目拟议的56%磷脂 / 60粒SKU具备产业落地条件'}];
for(const [i,p] of impacts.entries()){box(`impact-${i+1}-badge`,p.x,812,40,40,'#85b8f7',0,'ellipse');curve(`impact-${i+1}-tick`,p.x+10,822,22,19,[M(0,8),L(7,15),L(21,0)],'none','#ffffff',4);txt(`impact-${i+1}-text`,p.text,p.x+51,801,p.w-51,72,21,NAVY,true,{wrap:'square',lineSpacing:1.15});}

const raw=path.join(build,'raw.pptx'),styled=path.join(build,'native-styled.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
// Explicit native table borders and chart category ordering/label inheritance.
const nativeStyles=String.raw`
import sys,json
from zipfile import ZipFile
from lxml import etree as E
A='http://schemas.openxmlformats.org/drawingml/2006/main'
C='http://schemas.openxmlformats.org/drawingml/2006/chart'
tables=0;locks=0;bars=0
with ZipFile(sys.argv[1]) as src,ZipFile(sys.argv[2],'w') as dst:
 for info in src.infolist():
  data=src.read(info.filename)
  if info.filename.startswith('ppt/slides/slide') and info.filename.endswith('.xml'):
   root=E.fromstring(data)
   for e in root.iter():
    if e.get('noGrp')=='1':e.attrib.pop('noGrp');locks+=1
   for table in root.findall('.//{%s}tbl'%A):
    for row in table.findall('{%s}tr'%A):
     for cell in row.findall('{%s}tc'%A):
      props=cell.find('{%s}tcPr'%A)
      if props is None:props=E.SubElement(cell,'{%s}tcPr'%A)
      for key in ['lnL','lnR','lnT','lnB']:
       for e in props.findall('{%s}%s'%(A,key)):props.remove(e)
       ln=E.Element('{%s}%s'%(A,key),w='9525' if key=='lnB' else '0')
       if key=='lnB':E.SubElement(E.SubElement(ln,'{%s}solidFill'%A),'{%s}srgbClr'%A,val='E7EDF5')
       else:E.SubElement(ln,'{%s}noFill'%A)
       props.insert(0,ln)
    tables+=1
   data=E.tostring(root,encoding='UTF-8',xml_declaration=True)
  elif '/charts/' in info.filename and info.filename.endswith('.xml'):
   root=E.fromstring(data)
   direction=root.find('.//{%s}barDir'%C)
   horizontal=direction is not None and direction.get('val')=='bar'
   if horizontal:
    bars+=1
    for axis in root.findall('.//{%s}catAx'%C):
     scaling=axis.find('{%s}scaling'%C)
     if scaling is None:scaling=E.SubElement(axis,'{%s}scaling'%C)
     orient=scaling.find('{%s}orientation'%C)
     if orient is None:orient=E.SubElement(scaling,'{%s}orientation'%C)
     orient.set('val','maxMin')
   for label in root.findall('.//{%s}dLbl'%C):
    custom=label.find('{%s}tx'%C) is not None
    for key in ['showLegendKey','showVal','showCatName','showSerName','showPercent','showBubbleSize','showLeaderLines']:
     for e in label.findall('{%s}%s'%(C,key)):label.remove(e)
     E.SubElement(label,'{%s}%s'%(C,key),val='0' if custom or key!='showVal' else '1')
   for text in root.findall('.//{%s}txPr'%C):
    body=text.find('{%s}bodyPr'%A)
    if body is not None:body.set('wrap','none')
   data=E.tostring(root,encoding='UTF-8',xml_declaration=True)
  dst.writestr(info,data)
assert tables==1 and bars==2,(tables,bars)
print(json.dumps({'native_tables':tables,'horizontal_charts':bars,'removed_group_locks':locks}))
`;
await fs.writeFile(path.join(build,'native-styling.json'),execFileSync(process.env.RUNTIME_PYTHON,['-c',nativeStyles,raw,styled],{encoding:'utf8'}));
await fs.writeFile(path.join(build,'normalization.json'),execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',styled,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-tables','1','--min-native-connectors','1','--min-target-arrows','1'],{encoding:'utf8'}));
const finalPath=path.join(output,'market-research-4slides-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit','--require-native-table-slide','2'],explicitTotalSlideCount:4,requiredNativeTableOwnerSlides:[2],requiredNativeChartOwnerSlides:[1,3],materializeLiteralChartWorkbooks:true,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));for(const [i,s] of imported.slides.items.entries()){const png=await imported.export({slide:s,format:'png',scale:1});await fs.writeFile(path.join(build,`preview-${i+1}.png`),Buffer.from(await png.arrayBuffer()));}
console.log(JSON.stringify({finalPath,slides:4,charts:3,nativeTables:1,photoLimitation:'Slide 4 has explicit non-original schematic placeholders'},null,2));
