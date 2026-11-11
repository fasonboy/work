// Rebuild the krill-market page with editable device and phone illustrations.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/krill-market',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation,applyPresentationChartFont}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1672,H=941,FONT='Noto Sans CJK SC',NAVY='#062357',BLUE='#007ed4';
if(process.env.PPT_OPERATION_MARKED!=='1') execFileSync(process.env.RUNTIME_NODE,[path.join(skill,'container_tools/mark_artifact_operation_started.mjs'),'--operation-kind','create','--expected-output-count','1','--output-format','pptx'],{stdio:'inherit'});
const deck=Presentation.create({slideSize:{width:W,height:H}}),slide=deck.slides.add();
slide.background.fill='#ffffff';
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
function poly(name,points,fill) {
  const xs=points.map(p=>p[0]),ys=points.map(p=>p[1]),x=Math.min(...xs),y=Math.min(...ys),w=Math.max(...xs)-x,h=Math.max(...ys)-y;
  return curve(name,x,y,w,h,[M(points[0][0]-x,points[0][1]-y),...points.slice(1).map(p=>L(p[0]-x,p[1]-y)),Z()],fill);
}
function icon(kind,x,y,size,color='#0962aa') {
  const paths={
    coins:'<ellipse cx="40" cy="12" rx="17" ry="7"/><path d="M23 12v31c0 10 34 10 34 0V12M23 23c0 10 34 10 34 0M23 34c0 10 34 10 34 0"/><ellipse cx="21" cy="37" rx="17" ry="7"/><path d="M4 37v19c0 10 34 10 34 0V37M4 47c0 10 34 10 34 0"/>',
    trend:'<path d="M4 35l17-15 13 10L59 4M46 4h13v13"/><path d="M6 62V49h10v13zM26 62V39h10v23zM47 62V25h10v37z" fill="#38abe9"/>',
    pie:'<path d="M29 4a27 27 0 1 0 27 31H29z"/><path d="M36 1v27h27A27 27 0 0 0 36 1z" fill="#6ac6f5"/>',
    people:'<circle cx="32" cy="17" r="11"/><path d="M13 60V48c0-22 38-22 38 0v12M13 14C-1 10-3 32 11 34M53 14c14-4 16 18 2 20M5 42c-7 1-7 9-7 18M59 42c7 1 7 9 7 18"/>',
    drop:'<path d="M32 2C20 19 7 32 7 44a25 25 0 0 0 50 0C57 32 44 19 32 2z"/>',
    document:'<path d="M43 4H9a5 5 0 0 0-5 5v48a5 5 0 0 0 5 5h23M43 4v13h13zM14 19h19M14 29h12M14 39h12M14 49h12"/><path d="M45 27l17 7v13c0 10-9 17-17 21-8-4-17-11-17-21V34zM36 46l6 7 13-15"/>',
    trophy:'<path d="M14 5h36l-3 28c-2 13-9 19-15 19s-13-6-15-19zM32 52v9M21 63h23M14 9H3c-3 22 2 30 16 31M50 9h11c3 22-2 30-16 31"/>',
    target:'<circle cx="28" cy="36" r="27"/><circle cx="28" cy="36" r="17"/><circle cx="28" cy="36" r="6"/><path d="M28 36L58 6M47 5l3 11 12 2-9 10-12-2-3-12z"/>',
  };
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="76" height="80" viewBox="-7 -5 80 82"><g fill="none" stroke="${color}" stroke-width="3.8" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} icon`,position:{left:x,top:y,width:size,height:size}});
}
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
txt('page-title','南极磷虾油市场现状与发展机遇',65,27,1220,95,70,NAVY,true);
box('subtitle-accent',69,149,68,10,grad('#7ecbff','#009be8',0),6);
txt('page-subtitle','全球市场稳步增长，健康消费催生发展空间',156,126,1080,57,36,'#354d6b');

const panels=[{x:21,w:588,num:'01',label:'全球市场规模',titleX:159},{x:626,w:493,num:'02',label:'需求增长驱动',titleX:755},{x:1137,w:517,num:'03',label:'市场竞争与商业机会',titleX:1265}];
for(const [i,p] of panels.entries()) {
  box(`panel-${i+1}`,p.x,222,p.w,525,grad('#ffffff','#fcfeff'),21,'rect','#bdd7f8');
  box(`panel-${i+1}-header-round`,p.x+1,223,p.w-2,77,grad('#ffffff','#e0f0ff'),20);
  box(`panel-${i+1}-header-lower`,p.x+1,267,p.w-2,33,grad('#f1f9ff','#e0f0ff',90));
  curve(`panel-${i+1}-number-bg`,p.x+13,236,112,53,[M(17,0),L(112,0),L(94,53),L(0,53),Z()],grad('#00a9f5','#064ca1',50));
  txt(`panel-${i+1}-number`,p.num,p.x+42,236,67,51,40,'#ffffff',true);
  txt(`panel-${i+1}-heading`,p.label,p.titleX,237,p.x+p.w-p.titleX-10,55,i===2?33.5:35,NAVY,true);
}
icon('coins',47,349,55);txt('market-2024-caption','2024年市场规模',121,334,224,40,26,'#111111');
txt('market-about','约',121,377,33,43,28,'#141414');txt('market-2024-value','11',157,364,68,62,46,NAVY,true);txt('market-2024-unit','亿美元',225,377,108,45,30,'#111111');
box('market-divider-1',50,441,288,1.3,'#c2d7f1');
icon('trend',47,474,55);txt('market-2030-caption','2030年预计达到',121,458,225,42,26,'#111111');
txt('market-2030-value','21',121,489,77,66,51,NAVY,true);txt('market-2030-unit','亿美元',198,503,134,48,30,'#111111');
box('market-divider-2',50,567,288,1.3,'#c2d7f1');
icon('pie',45,615,56);txt('growth-period','2025—2030年',121,580,235,42,27,'#111111');
txt('growth-label','年复合增长率 (CAGR)',121,616,240,39,24,'#111111');txt('growth-value','11.1%',121,645,229,73,57,NAVY,true);

const chart=slide.charts.add('bar',{
  position:{left:352,top:338,width:246,height:392},titlePlacement:'none',categories:['2024年','2030年'],
  series:[{name:'市场规模（亿美元）',values:[11,21],fill:grad('#048fdc','#084b9e',90),points:[{idx:0,fill:grad('#ace3ff','#54bafb',90)},{idx:1,fill:grad('#008fd8','#034b9f',90)}]}],
  barOptions:{direction:'column',grouping:'clustered',gapWidth:26},hasLegend:false,
  xAxis:{visible:true,position:'bottom',line:{fill:'#1a407b',width:1.4,style:'solid'},textStyle:{fontSize:24,fill:'#111111'},majorGridlines:null,minorGridlines:null},
  yAxis:{visible:false,min:0,max:25,majorGridlines:null,minorGridlines:null,line:{fill:'none',width:0},tickLabelPosition:'none'},
  dataLabels:{showValue:true,position:'outEnd',numberFormatCode:'0"亿美元"',textStyle:{fontSize:25,fill:NAVY,bold:true}},
  chartFill:'none',chartLine:{fill:'none',width:0},plotAreaFill:'none',plotAreaLine:{fill:'none',width:0},
});
applyPresentationChartFont(chart,{fontFamily:FONT});

const drivers=[
 {y:332,kind:'people',text:'健康管理意识提升，\n带动营养补充需求增长',ty:336,h:90},
 {y:468,kind:'drop',text:'南极磷虾油含磷脂及\nOmega-3（EPA、DHA），\n具备成分差异化优势',ty:447,h:123,size:25.5},
 {y:606,kind:'document',text:'海洋营养食品发展\n受到政策和市场关注',ty:610,h:90},
];
for(const [i,d] of drivers.entries()) {
  box(`driver-${i+1}-icon-bg`,657,d.y,96,96,grad('#f3faff','#e4f2ff'),0,'ellipse','#e7f3ff');
  icon(d.kind,673,d.y+16,64,'#085ca5');
  txt(`driver-${i+1}-text`,d.text,778,d.ty,322,d.h,d.size??27.3,'#111111',false,{wrap:'square',lineSpacing:1.14});
}
box('drivers-divider-1',655,445,437,1.3,'#bad2ef');box('drivers-divider-2',655,584,437,1.3,'#bad2ef');
const opportunities=[
 {y:340,kind:'trophy',title:'竞争现状',detail:'品牌较多、\n品质参差'},
 {y:473,kind:'people',title:'消费关注',detail:'原料来源、\n成分含量、产品安全'},
 {y:613,kind:'target',title:'发展路径',detail:'品质透明、\n合理定价、渠道拓展'},
];
for(const [i,o] of opportunities.entries()) {
  icon(o.kind,1164,o.y+10,62,'#0759a2');txt(`opportunity-${i+1}-title`,o.title,1237,o.y+8,139,60,31,NAVY,true);
  box(`opportunity-${i+1}-divider`,1382,o.y+7,1.3,60,'#7597c7');
  txt(`opportunity-${i+1}-detail`,o.detail,1405,o.y-6,233,88,25.5,'#111111',false,{wrap:'square',lineSpacing:1.2});
}
box('opportunity-rule-1',1164,445,469,1.3,'#bad2ef');box('opportunity-rule-2',1164,581,469,1.3,'#bad2ef');

box('judgment-banner',23,796,1629,92,grad('#ffffff','#e6f3ff'),17);
curve('judgment-label-bg',23,797,269,89,[M(29,0),L(269,0),L(242,89),L(0,89),Z()],grad('#087dc8','#04377d',60));
txt('judgment-heading','市场判断',69,809,205,65,43,'#ffffff',true);
txt('judgment-text','行业具备长期成长空间，新品牌应以品质、价格竞争力和消费者信任建立市场优势。',311,812,1126,61,29,NAVY,true);
// Softgel decorations are individual editable ellipses and highlights.
for(const [i,x,y,w,h,rotation] of [[1,1476,834,102,35,-16],[2,1541,811,88,34,15]]) {
  const shadow=box(`capsule-${i}-shadow`,x+2,y+31,w,16,'#d6a4a8/22',0,'ellipse');shadow.position={left:x+2,top:y+31,width:w,height:16,rotation};
  const c=box(`capsule-${i}`,x,y,w,h,grad('#ffb2a1','#df170d',90),0,'ellipse','#cc3026',1.7);c.position={left:x,top:y,width:w,height:h,rotation};
  const shine=box(`capsule-${i}-shine`,x+10,y+4,w-22,8,'#fff6ed/72',0,'ellipse');shine.position={left:x+10,top:y+4,width:w-22,height:8,rotation};
}
slide.speakerNotes.text='来源：用户当前聊天“南极磷虾油市场现状与发展机遇”图片。市场规模按图保留为2024年约11亿美元、2030年预计21亿美元；2025—2030年CAGR 11.1%亦按原图保留，没有据端点重新推算。市场预测、营养、政策与行业结论未独立核验，也未新增外部数据来源。柱状图使用原生PowerPoint图表，数据为11和21，单位亿美元，内嵌工作簿。图表数据标签使用自身数值，排版与原图略有差异。雪山、水面、图标与胶囊均为近似重绘的矢量装饰，不是原始照片。';
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp')],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,'krill-market-editable.pptx');
// Source is a screenshot with two literal values, not an original workbook.
// Materialize a new data snapshot and validate its workbook links; do not claim source formulas.
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,requiredNativeChartOwnerSlides:[1],materializeLiteralChartWorkbooks:true,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),nativeChart:{categories:['2024年','2030年'],values:[11,21],unit:'亿美元'}},null,2));
