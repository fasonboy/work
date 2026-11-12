// Rebuild the global/China krill oil market page with an editable column chart.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/krill-market-china',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation,applyPresentationChartFont}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1672,H=941,FONT='Noto Sans CJK SC',NAVY='#062452',GREEN='#24af88';
if(process.env.PPT_OPERATION_MARKED!=='1') execFileSync(process.env.RUNTIME_NODE,[path.join(skill,'container_tools/mark_artifact_operation_started.mjs'),'--operation-kind','create','--expected-output-count','1','--output-format','pptx'],{stdio:'inherit'});
const deck=Presentation.create({slideSize:{width:W,height:H}}),slide=deck.slides.add();slide.background.fill='#ffffff';
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
function icon(kind,x,y,size,color=NAVY) {
  const paths={
    coins:'<ellipse cx="44" cy="14" rx="16" ry="7"/><path d="M28 14v29c0 10 32 10 32 0V14M28 24c0 10 32 10 32 0M28 34c0 10 32 10 32 0"/><ellipse cx="22" cy="39" rx="16" ry="7"/><path d="M6 39v17c0 10 32 10 32 0V39M6 48c0 10 32 10 32 0"/>',
    trend:'<path d="M4 36l17-15 13 9L59 5M46 5h13v13M6 62V50h10v12zM27 62V40h10v22zM48 62V25h10v37z"/>',
    pie:'<path d="M29 4a27 27 0 1 0 27 31H29zM36 1v27h27A27 27 0 0 0 36 1z"/>',
    heart:'<path d="M34 63C18 49 3 35 3 20 3 2 24 1 34 15 44 1 65 2 65 20c0 15-15 29-31 43zM8 32h13l7-12 10 25 7-13h14"/>',
    brain:'<path d="M34 9C19-3 10 6 11 17-3 19-3 37 9 41c-7 12 5 23 17 19 4 12 16 9 17 0 16 8 26-2 23-15 13-9 8-24-3-28 1-16-17-19-24-9zM34 10v48M18 21c-8 5-2 16 5 13M18 48c12 5 17-8 10-13M50 20c12 1 11 13 3 17M45 49c8 7 16-1 13-8"/>',
    runner:'<circle cx="46" cy="9" r="7"/><path d="M39 23l-11-8-13 11-6 13 6 4 13-17 7 5-9 14 12 10-8 17 7 3 12-21-11-12 8-14 8 8 12-8-4-6-9 6zM24 49L8 61l4 6 19-12"/>',
    leaf:'<path d="M7 46C14 18 35 27 61 12c0 38-23 54-47 39M4 61l39-30"/>',
    cart:'<rect x="4" y="5" width="60" height="44" rx="4"/><path d="M27 50v9M42 50v9M18 63h34M17 15h8l4 18h22l4-13H27"/><circle cx="31" cy="39" r="2"/><circle cx="49" cy="39" r="2"/>',
    target:'<circle cx="34" cy="34" r="24"/><circle cx="34" cy="34" r="15"/><path d="M34 1v17M34 51v17M1 34h17M51 34h17"/>',
    layers:'<path d="M5 21L34 4l29 17-29 18zM6 35l28 17 28-17M6 49l28 17 28-17"/>',
  };
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="82" height="84" viewBox="-5 -6 82 88"><g fill="none" stroke="${color}" stroke-width="3.7" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} market icon`,position:{left:x,top:y,width:size,height:size}});
}
// Editable polar scene and krill illustration replace photographic decoration only.
curve('polar-scene-bg',879,0,793,180,[M(38,0),B(-28,90,12,161,159,180),L(793,180),L(793,0),Z()],grad('#effff9','#c3e4fc',0));
poly('mountains-back',[[935,129],[1051,83],[1094,95],[1155,65],[1214,86],[1300,30],[1346,58],[1436,93],[1473,72],[1538,93],[1595,69],[1672,91],[1672,139]],'#b9d8ee');
poly('mountains-front',[[1013,136],[1142,96],[1182,112],[1300,30],[1343,91],[1390,69],[1451,128],[1558,82],[1595,114],[1672,91],[1672,143]],'#88bce1');
for(const [i,pts] of [[1,[[1300,30],[1230,126],[1279,94],[1296,133],[1323,72]]],[2,[[1300,30],[1343,91],[1390,69],[1337,128]]],[3,[[1155,65],[1103,133],[1135,110],[1168,135],[1214,86]]],[4,[[1558,82],[1539,137],[1594,115],[1627,141],[1672,91]]]])poly(`snow-facet-${i}`,pts,'#f5fbff');
box('polar-water',1000,137,672,43,grad('#80bedf','#d3edfb',90));
for(let i=0;i<6;i++)box(`water-line-${i}`,1040+(i%3)*22,142+i*6,Math.min(615-(i%4)*35,W-(1040+(i%3)*22)),1.3,i%2?'#b9e3f5':'#f1fbff');
for(let i=0;i<7;i++) {
  const x=846+i*15,y=111-i*5;
  const seg=box(`krill-segment-${i}`,x,y,24,14,'#ffe6d4',0,'ellipse','#f08b58',1);seg.position={left:x,top:y,width:24,height:14,rotation:-23};
  curve(`krill-leg-${i}`,x+6,y+9,22,29,[M(0,0),B(21,6,2,19,12,29)],'none','#adb9ce',1.4);
}
const shell=box('krill-carapace',911,72,60,21,grad('#fff0e3','#f56c2c'),0,'ellipse','#f08a56',1);shell.position={left:911,top:72,width:60,height:21,rotation:-17};
box('krill-head',959,63,21,15,'#ffab6b',0,'ellipse');box('krill-eye',969,64,5,5,'#13233b',0,'ellipse');
for(const [i,ey] of [[1,0],[2,16],[3,40]])curve(`krill-antenna-${i}`,975,35,55,46,[M(0,34),B(16,23,37,ey,55,ey)],'none','#efb38a',1);
poly('title-blue-slash',[[66,34],[80,34],[65,101],[51,101]],grad('#1e82aa','#034c75',90));
txt('page-title','南极磷虾油市场现状',96,10,826,112,72,NAVY,true,{italic:true});
txt('page-subtitle','全球市场增长与中国消费趋势',52,120,976,56,35,'#363636');

const panels=[{x:17,w:572,n:'01',title:'全球市场规模',tx:176},{x:607,w:422,n:'02',title:'需求增长动力',tx:741},{x:1048,w:608,n:'03',title:'国内市场与消费趋势',tx:1181}];
for(const [i,p] of panels.entries()) {
  box(`panel-${i+1}`,p.x,201,p.w,527,'#ffffff',30,'rect','#b6d1ee',1.5);
  box(`panel-${i+1}-header`,p.x+1,202,p.w-2,69,grad('#174f8b','#073763',90),29);
  box(`panel-${i+1}-header-lower`,p.x+1,237,p.w-2,34,grad('#174f8b','#073763',90));
  curve(`panel-${i+1}-number-bg`,p.x+9,208,123,62,[M(20,0),L(123,0),L(102,62),L(0,62),Z()],grad('#36bd93','#18a67f',50));
  txt(`panel-${i+1}-number`,p.n,p.x+48,212,66,56,40,'#ffffff',true);
  txt(`panel-${i+1}-heading`,p.title,p.tx,210,p.x+p.w-p.tx-15,61,i===2?35.5:37,'#ffffff',true);
}
txt('global-market-scope','全球磷虾油补充剂市场',170,279,386,43,23.5,'#111111');
icon('coins',53,352,56);txt('global-2024-label','2024年',135,336,193,44,26,'#111111');
txt('global-2024-value','11',135,365,67,64,51,NAVY,true);txt('global-2024-unit','亿美元',198,383,126,44,31,NAVY,true);
box('global-rule-1',49,443,271,1.3,'#d2e1f1');
icon('trend',53,482,56);txt('global-2030-label','2030年预计',135,459,193,44,26,'#111111');
txt('global-2030-value','21',135,490,67,65,51,NAVY,true);txt('global-2030-unit','亿美元',198,508,126,44,31,NAVY,true);
box('global-rule-2',49,568,271,1.3,'#d2e1f1');
icon('pie',51,612,57);txt('global-growth-period','2025—2030年',135,583,208,42,27,'#111111');
txt('global-cagr-label','CAGR',135,621,208,40,25,'#111111');txt('global-cagr-value','11.1%',135,648,217,66,53,NAVY,true);
const chart=slide.charts.add('bar',{
  position:{left:351,top:310,width:212,height:385},titlePlacement:'none',categories:['2024年','2030年'],
  series:[{name:'全球市场规模（亿美元）',values:[11,21],fill:grad('#039bd2','#0569af',90),points:[{idx:0,fill:grad('#b0e8fa','#69c6e9',90)},{idx:1,fill:grad('#00a2d6','#0067ae',90)}]}],
  barOptions:{direction:'column',grouping:'clustered',gapWidth:26},hasLegend:false,
  xAxis:{visible:true,position:'bottom',line:{fill:'#6a94b0',width:1.2,style:'solid'},textStyle:{fontSize:23,fill:'#111111'},majorGridlines:null,minorGridlines:null},
  yAxis:{visible:false,min:0,max:25,line:{fill:'none',width:0},majorGridlines:null,minorGridlines:null,tickLabelPosition:'none'},
  dataLabels:{showValue:true,position:'outEnd',numberFormatCode:'0"亿美元"',textStyle:{fontSize:25,fill:NAVY,bold:true}},
  chartFill:'none',chartLine:{fill:'none',width:0},plotAreaFill:'none',plotAreaLine:{fill:'none',width:0},
});applyPresentationChartFont(chart,{fontFamily:FONT});
const ga=box('chart-growth-start',402,476,1,1),gb=box('chart-growth-end',450,405,1,1);
slide.shapes.connect(ga,gb,{kind:'curved',line:{fill:GREEN,width:6},head:{type:'none'},tail:{type:'triangle',width:'med',length:'sm'}}).bringToFront();

const drivers=[{y:293,kind:'heart',text:'心血管健康'},{y:400,kind:'brain',text:'认知健康'},{y:507,kind:'runner',text:'运动营养'},{y:614,kind:'leaf',text:'天然与高品质营养'}];
for(const [i,d] of drivers.entries()) {
  box(`driver-${i+1}-pill`,683,d.y+5,326,79,grad('#ffffff','#edf6fe'),18);
  box(`driver-${i+1}-badge`,638,d.y,93,93,grad('#eefbf6','#e0f5ee'),0,'ellipse','#cff2e7');
  icon(d.kind,653,d.y+15,63,'#20a888');
  txt(`driver-${i+1}-text`,d.text,752,d.y+15,249,62,i===3?30.5:33,NAVY,true);
}

icon('cart',1081,337,63);txt('china-period','2024年1—9月',1163,299,218,41,26,'#111111');
txt('china-sales-label','国内线上销售额近',1163,335,218,42,25,'#111111');
box('china-sales-highlight',1169,391,151,41,grad('#f3fbff','#e6f2ff'),9);
txt('china-sales-value','10',1163,364,82,75,59,NAVY,true);txt('china-sales-unit','亿元',1242,390,84,49,33,NAVY,true);
box('china-stat-divider',1387,307,1.3,121,'#c7d7ed');
icon('trend',1427,312,58);txt('china-yoy-label','同比增长',1496,310,146,51,30,'#151515',true);
box('china-growth-highlight',1420,369,217,65,grad('#f6fffa','#e0f5ed'),16);
txt('china-yoy-value','24.3%',1461,360,181,80,55,GREEN,true);
box('china-rule-1',1074,463,557,1.3,'#d2e1f1');
icon('target',1091,502,58);box('china-quality-divider',1178,501,1.3,56,'#7b9bc5');
txt('china-quality-text','关注原料来源、EPA/DHA含量、\n磷脂含量及产品安全',1208,489,428,86,27.5,'#111111',false,{wrap:'square',lineSpacing:1.16});
box('china-rule-2',1074,584,475,1.3,'#d2e1f1');
icon('layers',1093,615,57);box('china-trend-divider',1178,613,1.3,58,'#7b9bc5');
txt('china-trend-text','品类认知提升，品质与行业\n规范化受到关注',1208,602,428,88,27.5,'#111111',false,{wrap:'square',lineSpacing:1.16});

box('outlook-banner',34,773,1622,116,grad('#ffffff','#e0efff'),32);
curve('global-growth-label-bg',35,791,335,79,[M(25,0),L(335,0),L(311,79),L(0,79),Z()],grad('#0b5797','#123d75',0));
txt('global-growth-label','全球市场增长',94,800,258,60,34,'#ffffff',true,{italic:true});
const fa=box('outlook-arrow-from',382,831,1,1),fb=box('outlook-arrow-to',545,831,1,1);
slide.shapes.connect(fa,fb,{kind:'straight',line:{fill:grad('#dcefff','#0087d7',0),width:12},head:{type:'none'},tail:{type:'triangle',width:'med',length:'sm'}}).bringToFront();
curve('china-development-label-bg',562,791,337,79,[M(25,0),L(337,0),L(312,79),L(0,79),Z()],grad('#32bb92','#10a57f',0));
txt('china-development-label','中国市场发展',621,800,263,60,34,'#ffffff',true,{italic:true});
box('outlook-divider',960,792,1.5,78,'#294974');
txt('outlook-conclusion','全球市场持续扩容，\n中国市场进入品类认知与品质升级阶段。',1017,781,617,99,33,NAVY,true,{wrap:'square',lineSpacing:1.06});
slide.speakerNotes.text='来源：用户当前聊天“南极磷虾油市场现状—全球市场增长与中国消费趋势”截图。全球市场2024年11亿美元、2030年预计21亿美元、2025—2030年CAGR 11.1%以及国内2024年1—9月线上销售额近10亿元、同比增长24.3%按图保留，单位未互换，未独立核验来源或重算CAGR。全球柱状图使用原生PowerPoint图表，内嵌依据11和21新建的Excel数据快照，不声称保留原始工作簿/公式。需求驱动、营养与消费趋势文案按图转录，不构成独立验证。雪山、水面、磷虾与图标近似矢量重绘，不是原始照片。';
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-connectors','2','--min-target-arrows','2'],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,'krill-market-china-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,requiredNativeChartOwnerSlides:[1],materializeLiteralChartWorkbooks:true,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),globalChartValues:[11,21],globalUnit:'亿美元',chinaSales:'近10亿元',chinaGrowth:'24.3%'},null,2));
