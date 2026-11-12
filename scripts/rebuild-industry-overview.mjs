// Rebuild the industry overview screenshot with editable range data and driver cards.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/industry-overview',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation,applyPresentationChartFont}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1672,H=941,FONT='Noto Sans CJK SC',NAVY='#133f70',RED='#e34442',BLUE='#347fc1';
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
slide=deck.slides.add();slide.background.fill=grad('#ffffff','#eef5ff',70);
// Decorative polar scene is made from native facets and water lines.
box('polar-water',1065,151,607,50,grad('#d6ebf8/55','#ffffff/0',90));
poly('mountains-back',[[1080,156],[1229,89],[1270,107],[1352,55],[1408,71],[1498,9],[1553,33],[1672,62],[1672,157]],'#eaf5fd');
poly('mountains-front',[[1172,157],[1298,97],[1358,118],[1457,45],[1503,10],[1542,54],[1594,45],[1672,74],[1672,158]],'#d9eaf7');
for(const [i,points] of [[1,[[1503,10],[1457,113],[1483,91],[1495,138],[1532,61]]],[2,[[1503,10],[1542,54],[1594,45],[1540,104]]],[3,[[1457,45],[1408,139],[1435,121],[1451,148],[1492,94]]],[4,[[1352,55],[1303,150],[1341,115],[1373,136],[1408,71]]],[5,[[1594,45],[1581,146],[1615,105],[1651,145],[1672,74]]]])poly(`snow-facet-${i}`,points,'#ffffff/86');
for(const [i,points] of [[1,[[1504,29],[1487,72],[1511,60]]],[2,[[1538,64],[1528,113],[1550,98]]],[3,[[1448,91],[1435,139],[1464,120]]],[4,[[1599,81],[1585,145],[1619,117]]],[5,[[1656,94],[1630,144],[1662,131]]]])poly(`blue-facet-${i}`,points,'#a3c7e4/70');
for(let i=0;i<8;i++) {
  const x=1124+(i%3)*25;
  box(`water-line-${i}`,x,157+i*5,Math.min(548-(i%4)*21,W-x),1.2,i%2?'#c7deef/48':'#eef7fc/75');
}
const paths={
  globe:'<circle cx="34" cy="34" r="28"/><path d="M9 19l12 2 4 10-6 8 7 12-2 10M29 7l-5 9 11 6 4-8 12 5 7 12-13 1-5 15 7 8M13 9l7 5M57 51l-7-4"/>',
  trend:'<path d="M6 32L24 17l13 10L63 4M48 4h15v15M7 64V47h11v17zM29 64V37h11v27zM51 64V26h11v38z"/>',
  bars:'<path d="M8 62V42h11v20zM29 62V27h11v35zM50 62V10h11v52z" fill="currentColor"/>',
  linechart:'<path d="M5 5v57h59M12 46l14-17 12 9 23-27M48 11h13v13M17 59V49M30 59V43M43 59V39M56 59V29"/>',
  pie:'<path d="M30 4a27 27 0 1 0 27 31H30zM37 1v27h27A27 27 0 0 0 37 1z"/>',
  gear:'<path d="M29 4h11l3 8 8 3 8-3 8 9-5 7 1 9 7 5-4 11-9 1-7 6-1 9H35l-3-9-8-4-8 3-8-10 5-8-1-9-7-5 4-10 9-1 7-6z"/><circle cx="37" cy="36" r="11"/>',
  heart:'<path d="M34 63C18 48 4 35 4 20 4 1 25 0 34 15 43 0 64 1 64 20c0 15-14 28-30 43zM6 33h15l6-12 10 24 6-12h18"/>',
  drop:'<path d="M34 3C21 21 9 35 9 47a25 25 0 0 0 50 0C59 35 47 21 34 3zM20 48c1 10 9 15 17 14"/>',
  sprout:'<path d="M34 66V20M34 45C11 46 5 31 6 21c19 0 31 12 28 24zM34 46c-3-22 14-26 27-27 0 20-12 31-27 27zM34 35C20 24 24 9 34 0c13 10 16 26 0 35z"/>',
  shield:'<path d="M34 4l26 10v20c0 15-13 26-26 32C21 60 8 49 8 34V14zM22 34l10 10 18-22"/>',
};
function pict(kind,x,y,size,color=NAVY){
  const body=paths[kind].replaceAll('currentColor',color);
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="84" height="84" viewBox="-6 -5 84 84"><g fill="none" stroke="${color}" stroke-width="4.2" stroke-linecap="round" stroke-linejoin="round">${body}</g></svg>`,alt:`${kind} icon`,position:{left:x,top:y,width:size,height:size}});
}
box('title-accent',52,31,14,66,grad('#ed5255','#c72032'),3);
txt('page-title','一、行业概况：市场扩容与产品认知升级',91,22,1281,65,43,NAVY,true);
txt('page-subtitle','全球需求稳步增长，健康消费升级与原料品质成为行业发展的重要驱动力',92,91,1200,40,23,'#3e5573');
box('header-divider',92,140,1518,1.3,'#a0b7cd');

box('market-summary-shadow',40,173,788,282,'#dae3ee/60',28);
box('market-summary',40,168,787,281,grad('#ffffff','#f8fbff'),28);
curve('summary-heading-bg',40,168,495,55,[M(27,0),L(495,0),L(474,55),L(0,55),L(0,26),B(0,10,10,0,27,0),Z()],grad('#174778','#416d99'));
poly('summary-heading-accent',[[527,168],[538,168],[516,223],[505,223]],'#9bc7e7');
pict('linechart',66,177,43,'#ffffff');txt('summary-heading','全球市场规模与增长预期',130,174,386,45,30,'#ffffff',true);
box('summary-rule-1',300,258,1.2,167,'#becade');box('summary-rule-2',591,258,1.2,167,'#becade');
const metrics=[{x:61,w:218,kind:'globe',label:'2025年市场规模',val:'5.4–9.3',sub:'亿美元',font:46},{x:313,w:266,kind:'trend',label:'预计年复合增长率',val:'8.4%–11.3%',sub:'（2026–2034年）',font:43},{x:605,w:207,kind:'bars',label:'2034年市场预测',val:'11–24',sub:'亿美元',font:46}];
for(const [i,m] of metrics.entries()){
  pict(m.kind,m.x+(m.w-51)/2,247,51,'#377fbe');
  txt(`metric-${i+1}-label`,m.label,m.x,300,m.w,36,20,'#445873',true,{alignment:'center'});
  txt(`metric-${i+1}-value`,m.val,m.x,335,m.w,61,m.font,RED,true,{alignment:'center'});
  txt(`metric-${i+1}-unit-or-period`,m.sub,m.x,396,m.w,39,24,'#445873',false,{alignment:'center'});
}

box('market-chart-shadow',40,481,788,387,'#d9e5f2/60',28);box('market-chart-panel',40,476,787,386,grad('#ffffff','#f7fbff'),28);
curve('market-chart-heading-bg',40,476,570,57,[M(27,0),L(570,0),L(549,57),L(0,57),L(0,26),B(0,10,10,0,27,0),Z()],grad('#174778','#476f99'));
poly('market-chart-heading-accent',[[601,476],[611,476],[589,533],[579,533]],'#99c3e5');
pict('pie',67,485,43,'#ffffff');txt('market-chart-heading','全球磷虾油市场规模区间（亿美元）',130,482,459,45,29,'#ffffff',true);
box('market-chart-accent',65,543,119,4,'#5c9ed4');
// Two stacked components preserve both range limits. Total bar height is the upper bound.
const chart=slide.charts.add('bar',{
  position:{left:91,top:552,width:685,height:261},categories:['2025年','2034年预测'],titlePlacement:'none',hasLegend:false,
  series:[
    {name:'区间下限（亿美元）',values:[5.4,11],points:[{idx:0,fill:'#78add9'},{idx:1,fill:'#e85a56'}],dataLabels:{showValue:false,textStyle:{typeface:FONT}}},
    {name:'上限减下限（亿美元）',values:[3.9,13],points:[{idx:0,fill:'#78add9'},{idx:1,fill:'#e85a56'}],dataLabels:{showValue:false,position:'inEnd',textStyle:{typeface:FONT,fontSize:30,fill:'#ffffff',bold:true}},dataLabelOverrides:[{idx:0,text:'5.4–9.3',position:'inEnd',textStyle:{typeface:FONT,fontSize:29,fill:'#ffffff',bold:true}},{idx:1,text:'11–24',position:'inEnd',textStyle:{typeface:FONT,fontSize:30,fill:'#ffffff',bold:true}}]},
  ],
  barOptions:{direction:'column',grouping:'stacked',gapWidth:125},
  xAxis:{visible:true,position:'bottom',line:{fill:'#9aadc5',width:1.5,style:'solid'},textStyle:{typeface:FONT,fontSize:25,fill:'#465976',bold:true},majorGridlines:null,minorGridlines:null},
  yAxis:{visible:false,min:0,max:28,line:{fill:'none',width:0},majorGridlines:null,minorGridlines:null,tickLabelPosition:'none'},
  chartFill:'none',chartLine:{fill:'none',width:0},plotAreaFill:'none',plotAreaLine:{fill:'none',width:0},
});applyPresentationChartFont(chart,{fontFamily:FONT});
const ca=box('market-growth-arrow-start',331,700,1,1),cb=box('market-growth-arrow-end',506,617,1,1);
slide.shapes.connect(ca,cb,{kind:'curved',line:{fill:grad('#fbdde6','#f87898'),width:13},head:{type:'none'},tail:{type:'triangle',width:'sm',length:'sm'}}).bringToFront();
box('market-chart-note-accent',65,826,4,18,NAVY);txt('market-chart-note','区间反映不同机构统计口径差异，非单一机构连续预测值。',80,817,711,35,16,'#4c607b');

curve('drivers-heading-bg',849,168,500,55,[M(24,0),L(500,0),L(481,55),L(0,55),L(0,25),B(0,10,10,0,24,0),Z()],grad('#174778','#456c94'));
poly('drivers-heading-red-accent',[[1342,168],[1352,168],[1332,223],[1322,223]],'#f16a69');pict('gear',877,176,42,'#ffffff');txt('drivers-heading','行业增长与产品认知驱动',938,174,392,45,30,'#ffffff',true);
const drivers=[
 {x:853,y:234,w:382,h:232,n:'01',title:'健康消费升级',text:'营养补充与日常健康管理需求\n持续提升。',kind:'heart',color:RED,fill:'#fff1f2'},
 {x:1261,y:234,w:379,h:232,n:'02',title:'磷脂型Omega-3',text:'磷脂结合形态是磷虾油区别于\n常见鱼油的认知点。',kind:'drop',color:BLUE,fill:'#eaf4ff'},
 {x:853,y:491,w:382,h:241,n:'03',title:'天然虾青素',text:'原料天然含有虾青素，形成产\n品差异化表达。',kind:'sprout',color:BLUE,fill:'#eaf4ff'},
 {x:1261,y:491,w:379,h:241,n:'04',title:'品质与可追溯',text:'原料来源、污染物检测与可持\n续捕捞信息影响信任。',kind:'shield',color:RED,fill:'#fff0f1'},
];
for(const [i,d] of drivers.entries()){
  box(`driver-${i+1}-shadow`,d.x,d.y+5,d.w,d.h,'#d7e0ec/70',27);
  box(`driver-${i+1}-card`,d.x,d.y,d.w,d.h,grad('#ffffff',d.fill,45),27);
  curve(`driver-${i+1}-number-bg`,d.x+24,d.y+21,70,50,[M(8,0),L(50,0),B(62,0,64,9,64,18),L(70,26),L(64,32),B(64,45,58,50,48,50),L(8,50),B(0,50,0,46,0,41),L(0,8),B(0,0,3,0,8,0),Z()],grad(d.color,d.color));
  txt(`driver-${i+1}-number`,d.n,d.x+35,d.y+27,48,37,26,'#ffffff',true,{alignment:'center'});
  box(`driver-${i+1}-icon-bg`,d.x+177,d.y+13,67,67,i===0||i===3?'#fff5f5':'#eaf5ff',0,'ellipse');pict(d.kind,d.x+178,d.y+17,64,d.color);
  txt(`driver-${i+1}-title`,d.title,d.x+26,d.y+87,d.w-45,46,32,NAVY,true);
  box(`driver-${i+1}-rule`,d.x+28,d.y+137,i===0?194:i===3?44:38,2.2,d.color);
  txt(`driver-${i+1}-text`,d.text,d.x+26,d.y+146,d.w-46,72,24,'#334f72',false,{wrap:'square',lineSpacing:1.22});
}
box('reporting-note-panel',852,764,788,99,grad('#e6f1ff','#e0ecfb'),23);
box('reporting-note-info-badge',890,788,41,41,grad('#4c76a7','#244b7d'),0,'ellipse');txt('reporting-note-info','i',902,790,20,36,31,'#ffffff',true,{alignment:'center'});
txt('reporting-note-heading','汇报口径提示',946,784,166,48,27,NAVY,true);box('reporting-note-divider',1121,790,1.4,39,'#7c9cbd');
txt('reporting-note-text','市场规模数据因原料、终端成品及统计区域不同而存在差异。产品吸收率与健康功效表述，应以具体产品检测资料及合规范围为依据，不宜直接表述为确定的临床优势。',1143,782,476,71,15.5,'#465d7b',false,{wrap:'square',lineSpacing:1.18});
curve('footer-light-blue',0,880,W,61,[M(0,0),B(484,43,985,100,W,0),L(W,61),L(0,61),Z()],grad('#60a4dc','#bcdcff'));
curve('footer-ice-blue',0,898,W,43,[M(0,0),B(500,31,929,59,1370,17),B(1485,7,1595,20,W,36),L(W,43),L(0,43),Z()],'#8fc1ea/65');
slide.speakerNotes.text='来源：用户当前聊天“行业概况：市场扩容与产品认知升级”截图，1672×941。2025年市场规模5.4–9.3亿美元、预计年复合增长率8.4%–11.3%（2026–2034年）、2034年市场预测11–24亿美元按截图保留，不独立核验或选取单值。左下为原生堆积柱形图：下限系列5.4、11；跨度系列为9.3−5.4=3.9、24−11=13，柱顶表示上限，标签保留完整区间。内嵌新建数据快照，不声称原始工作簿或公式。图形高度根据上限绘制，不把它解释为同一机构连续预测。四项认知驱动按图转录，不增加营养功效断言。右下原截图小字部分模糊，提示按可辨含义重述，非逐字转录，使用者可编辑复核。雪山与图标为近似矢量重绘，不是原始照片；没有整页图片背景。';
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-connectors','1','--min-target-arrows','1'],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,'industry-overview-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,requiredNativeChartOwnerSlides:[1],materializeLiteralChartWorkbooks:true,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),marketRanges:[[5.4,9.3],[11,24]],cagrRange:'8.4%–11.3%',smallPrint:'Restated from legible meaning'},null,2));
