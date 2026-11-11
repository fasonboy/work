// Rebuild the fish-oil/krill-oil comparison with a native table and percentage chart.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/fish-krill-opportunity',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation,applyPresentationChartFont}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1672,H=941,FONT='Noto Sans CJK SC',NAVY='#061f54';
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
function icon(kind,x,y,size,color='#0868c6') {
  const paths={
    fish:'<path d="M5 34c18-22 38-22 53-7l15-8-4 15 4 15-15-8C43 56 23 56 5 34zM31 13c1-9 11-11 15-3M33 52c0 9 9 10 13 6"/><circle cx="21" cy="31" r="2"/>',
    heart:'<path d="M37 65C21 50 4 36 4 20 4 1 27-2 37 14 47-2 70 1 70 20c0 16-17 30-33 45zM14 32h13l7-12 10 24 7-12h12"/>',
    trend:'<path d="M5 35C28 31 42 20 60 3M45 3h15v15"/><path d="M7 62V48h11v14zM29 62V38h11v24zM51 62V26h11v36z" fill="#1492e8"/>',
    drop:'<path d="M34 4C22 20 9 35 9 47a25 25 0 0 0 50 0C59 35 46 20 34 4z"/>',
    molecules:'<path d="M35 2l14 8v16l-14 8-14-8V10zM17 31l14 8v16l-14 8-14-8V39zM53 31l14 8v16l-14 8-14-8V39zM35 34v8M31 45h8"/>',
    shield:'<path d="M34 3l26 11v19c0 15-14 26-26 32C22 59 8 48 8 33V14zM20 33l10 10 20-22"/>',
    thumb:'<path d="M26 29l8-21c10-4 14 2 10 17h20c7 0 7 6 6 10l-6 24H26z" fill="#ffffff" stroke="#ffffff"/><path d="M10 29h10v30H10z" fill="#ffffff" stroke="#ffffff"/>',
  };
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="84" height="80" viewBox="-4 -6 84 84"><g fill="none" stroke="${color}" stroke-width="4.1" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} icon`,position:{left:x,top:y,width:size,height:size}});
}
// Editable logo and polar decoration, not original image assets.
const nodes=[[72,39,'#ffcd00'],[94,39,'#ffbf00'],[113,60,'#ff9f00'],[123,52,'#ffae00'],[138,41,'#ffab00'],[151,63,'#ff9400'],[130,74,'#ff6b00'],[97,71,'#8fbd00'],[81,57,'#a6c600']];
for(const [i,j] of [[0,1],[1,2],[2,3],[3,4],[4,5],[5,6],[6,2],[2,7],[7,8]]) {
  const [x1,y1,c]=nodes[i],[x2,y2]=nodes[j];slide.shapes.add({name:`logo-edge-${i}-${j}`,geometry:'line',position:{left:Math.min(x1,x2),top:Math.min(y1,y2),width:Math.abs(x2-x1)||.1,height:Math.abs(y2-y1)||.1,verticalFlip:(x2-x1)*(y2-y1)<0},fill:'none',line:{fill:c,width:4}});
}
for(const [i,[x,y,c]] of nodes.entries())box(`logo-node-${i}`,x-5.5,y-5.5,11,11,c,0,'ellipse');
const logo=txt('logo-wordmark','',43,82,149,37,24,'#333333');logo.text=[[{run:'JKN ',style:{typeface:FONT,fontSize:'24px',color:'#333333'}},{run:'Biotech',style:{typeface:FONT,fontSize:'24px',color:'#333333',bold:true}}]];
box('logo-divider',197,32,1.5,78,'#a1a8b5');
poly('polar-mountain-back',[[1259,117],[1373,81],[1400,93],[1471,53],[1519,64],[1581,22],[1622,55],[1672,74],[1672,120]],'#dff2ff');
poly('polar-mountain-front',[[1320,118],[1438,88],[1489,102],[1581,22],[1615,80],[1640,69],[1672,99],[1672,121]],'#b9dff8');
for(const [i,pts] of [[1,[[1581,22],[1522,111],[1568,79],[1575,113],[1601,60]]],[2,[[1581,22],[1615,80],[1640,69],[1609,111]]],[3,[[1471,53],[1414,115],[1450,92],[1478,116],[1519,64]]],[4,[[1640,69],[1623,118],[1651,102],[1672,119],[1672,99]]]])poly(`snow-facet-${i}`,pts,'#ffffff');
box('polar-water',1259,120,413,17,grad('#c2e6fc','#ffffff',90));
txt('page-title','从鱼油的发展看南极磷虾油的市场机会',223,0,1121,85,60,NAVY,true,{italic:true});
txt('page-subtitle','鱼油被广泛认知后市场快速增长，磷虾油具备更优的营养优势，市场前景巨大',227,79,1092,49,29,'#101d32');

const panels=[{x:27,w:801,n:'01',title:'鱼油：认知提升带动市场快速增长',tx:168},{x:849,w:799,n:'02',title:'鱼油与磷虾油：磷虾油更具优势',tx:997}];
for(const [i,p] of panels.entries()) {
  box(`panel-${i+1}`,p.x,158,p.w,523,'#ffffff',30,'rect','#c2dcff',1.5);
  box(`panel-${i+1}-header`,p.x+1,159,p.w-2,64,grad('#f8fcff','#e2f0ff'),29);
  box(`panel-${i+1}-header-lower`,p.x+1,191,p.w-2,32,grad('#f8fcff','#e2f0ff'));
  curve(`panel-${i+1}-number-bg`,p.x+1,159,122,62,[M(23,0),L(122,0),L(103,62),L(0,62),B(2,34,9,4,23,0),Z()],grad('#2b9bff','#024caf',60));
  txt(`panel-${i+1}-number`,p.n,p.x+32,158,83,60,46,'#ffffff',true,{italic:true});
  txt(`panel-${i+1}-title`,p.title,p.tx,163,p.x+p.w-p.tx-10,54,36,NAVY,true);
}
const stages=[{x:67,w:191,title:'2000年前',desc:'小众补充剂',kind:'fish'},{x:324,w:213,title:'2000–2010年',desc:'健康价值被广泛认知',kind:'heart'},{x:611,w:185,title:'2010年后',desc:'进入大众消费',kind:'trend'}];
for(const [i,s] of stages.entries()) {
  box(`stage-${i+1}-surface`,s.x,251,s.w,167,grad('#f1f8ff','#e7f3ff'),28);
  icon(s.kind,s.x+(s.w-82)/2,264,82);
  txt(`stage-${i+1}-title`,s.title,s.x+5,330,s.w-10,48,i===1?28:29,NAVY,true,{alignment:'center'});
  txt(`stage-${i+1}-description`,s.desc,s.x-2,373,s.w+4,38,i===1?22:25,'#101010',false,{alignment:'center'});
}
for(const [i,x] of [[1,277],[2,559]]) {
  const a=box(`stage-arrow-${i}-from`,x,339,1,1),b=box(`stage-arrow-${i}-to`,x+28,339,1,1);
  slide.shapes.connect(a,b,{kind:'straight',line:{fill:grad('#cceaff','#007fff',0),width:8},head:{type:'none'},tail:{type:'triangle',width:'med',length:'sm'}}).bringToFront();
}
box('left-panel-divider',49,459,759,1.3,'#d6e7fb');
txt('fish-use-label','美国成年人\n鱼油使用率',64,521,164,97,30,'#101010',true,{wrap:'square',lineSpacing:1.17});
const chart=slide.charts.add('bar',{
  position:{left:229,top:474,width:335,height:194},titlePlacement:'none',categories:['1999–2000年','2011–2012年'],
  series:[{name:'美国成年人鱼油使用率',values:[.013,.12],valuesFormatCode:'0.0%',fill:grad('#229dff','#0754bd',90),points:[{idx:0,fill:grad('#badfff','#82b8f6',90)},{idx:1,fill:grad('#269fff','#0055c5',90)}]}],
  barOptions:{direction:'column',grouping:'clustered',gapWidth:45},hasLegend:false,
  xAxis:{visible:true,position:'bottom',line:{fill:'#828ba3',width:1.4,style:'solid'},textStyle:{fontSize:22,fill:'#111111'},majorGridlines:null,minorGridlines:null},
  yAxis:{visible:false,min:0,max:.14,line:{fill:'none',width:0},majorGridlines:null,minorGridlines:null,tickLabelPosition:'none'},
  dataLabels:{showValue:true,position:'outEnd',numberFormatCode:'0.0%',textStyle:{fontSize:29,fill:'#0a3495',bold:true}},
  chartFill:'none',chartLine:{fill:'none',width:0},plotAreaFill:'none',plotAreaLine:{fill:'none',width:0},
});
applyPresentationChartFont(chart,{fontFamily:FONT});
const growthA=box('growth-arrow-start',355,566,1,1),growthB=box('growth-arrow-end',413,516,1,1);
slide.shapes.connect(growthA,growthB,{kind:'curved',line:{fill:'#ff8200',width:7},head:{type:'none'},tail:{type:'triangle',width:'med',length:'sm'}}).bringToFront();
box('ninefold-card',580,501,226,134,grad('#f1f8ff','#d1e5ff'),26);
txt('ninefold-heading','使用率提升',621,512,170,48,29,NAVY,true);
txt('ninefold-about','约',621,570,47,50,36,NAVY,true);txt('ninefold-value','9',670,546,61,80,62,'#0759d7',true,{italic:true});txt('ninefold-unit','倍',730,560,64,65,48,'#0759d7',true,{italic:true});
curve('ninefold-underline',626,616,134,13,[M(0,10),B(54,-1,91,-1,134,6),B(86,7,54,8,17,13),Z()],'#8bbff9');

const table=slide.tables.add({rows:4,columns:3,left:867,top:250,width:763,height:368,columnWidths:[212,222,329],values:[
 ['对比维度','鱼油','南极磷虾油'],
 ['营养形式','甘油三酯型\n（传统形式）','磷脂型Omega-3\n（更易吸收利用）'],
 ['营养成分','主要提供\nEPA、DHA','EPA、DHA + 磷脂 +\n天然虾青素（抗氧化）'],
 ['综合优势','应用成熟','吸收好、抗氧化、\n营养更全面'],
]});
const rowHeights=[62,102,104,100];
for(let r=0;r<4;r++) {
  table.rows[r].height=rowHeights[r];
  for(let c=0;c<3;c++) {
    const cell=table.getCell(r,c);
    cell.fill=r===0?(c===1?grad('#d7eaff','#bcdbfc'):grad('#168ff3','#0049ae',60)):(c===0?grad('#f8fcff','#edf4fc'):c===2?grad('#fffdf8','#fff2dc'):'#ffffff');
    cell.text.style={typeface:FONT,fontSize:r===0?(c===1?32:30):c===2?25.5:27,color:r===0?(c===1?NAVY:'#ffffff'):c===1?'#111111':NAVY,bold:r===0||c!==1,autoFit:'none',wrap:'square',alignment:r>0&&c===0?'left':'center',verticalAlignment:'middle'};
    table.cells.block({row:r,column:c,rowCount:1,columnCount:1}).assign({margins:{left:r>0&&c===0?80:10,right:r>0&&c===2?43:8,top:8,bottom:8}});
  }
}
table.borders.assign({fill:'#c5d9f3',width:1.1,style:'solid'});
table.getCell(1,2).text.get('（更易吸收利用）').bold=false;
table.getCell(1,2).text.get('（更易吸收利用）').color='#101010';
for(const [kind,y] of [['drop',335],['molecules',438],['shield',540]])icon(kind,884,y,54,'#0966c9');
for(const [i,y] of [[1,343],[2,446],[3,550]]) {
  box(`table-thumb-${i}-bg`,1580,y,35,35,grad('#3baaff','#066bd4'),0,'ellipse');icon('thumb',1582,y+2,30,'#ffffff');
}

box('outlook-banner',29,723,1618,155,grad('#ffffff','#deeeff'),33);
curve('outlook-label-bg',29,749,315,103,[M(34,0),L(315,0),L(280,103),L(0,103),Z()],grad('#0869df','#0345a0',60));
txt('outlook-label','市场前景',85,763,230,77,47,'#ffffff',true,{italic:true});
const separator=box('outlook-separator',371,761,1.5,74,'#344a6b');separator.position={left:371,top:761,width:1.5,height:74,rotation:4};
txt('outlook-line-1','鱼油的快速普及验证了Omega-3产品的巨大市场潜力，',414,744,996,54,32.5,NAVY,true);
txt('outlook-line-2','南极磷虾油凭借更优的营养形式和差异化优势，',414,794,715,55,32.5,NAVY,true);
txt('outlook-emphasis','市场前景广阔。',1118,786,278,67,39,'#03449e',true,{italic:true});
curve('outlook-blue-underline',1119,845,247,18,[M(0,14),B(115,-3,183,-5,247,3),B(175,4,102,9,16,18),Z()],'#4399f6');
poly('ice-bed',[[1280,873],[1394,833],[1435,846],[1483,817],[1540,840],[1579,815],[1647,807],[1647,873]],'#cbe9fd');
for(const [i,pts] of [[1,[[1394,833],[1386,873],[1435,846]]],[2,[[1483,817],[1446,873],[1502,858]]],[3,[[1579,815],[1537,873],[1647,851]]]])poly(`ice-facet-${i}`,pts,'#f3fbff');
for(const [i,x,y,w,h,rotation] of [[1,1417,796,130,48,-14],[2,1507,760,113,47,18]]) {
  const shadow=box(`capsule-${i}-shadow`,x+2,y+h-2,w,19,'#cf9da0/25',0,'ellipse');shadow.position={left:x+2,top:y+h-2,width:w,height:19,rotation};
  const cap=box(`capsule-${i}`,x,y,w,h,grad('#ffb6a2','#da170c',90),0,'ellipse','#c6251c',2);cap.position={left:x,top:y,width:w,height:h,rotation};
  const shine=box(`capsule-${i}-shine`,x+12,y+5,w-26,11,'#fff6ef/75',0,'ellipse');shine.position={left:x+12,top:y+5,width:w-26,height:11,rotation};
}
slide.speakerNotes.text='来源：用户当前聊天“从鱼油的发展看南极磷虾油的市场机会”图片。时间线、营养对比、美国成年人鱼油使用率1.3%和12.0%、约9倍及市场判断均按图保留，未独立核验来源或营养功效。原生柱状图内的数据分别为0.013和0.12，以0.0%显示；内嵌新建数据快照，不声称原始工作簿或公式血缘。对比内容使用真正的4行3列PowerPoint表格，可按表格操作编辑。文字、Logo、图标、雪山、冰面和胶囊为重建对象，装饰外观与原照片有差异。';
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-connectors','3','--min-target-arrows','3','--min-native-tables','1'],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,'fish-krill-opportunity-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit','--require-native-table-slide','1'],explicitTotalSlideCount:1,requiredNativeTableOwnerSlides:[1],requiredNativeChartOwnerSlides:[1],materializeLiteralChartWorkbooks:true,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),nativeTable:{rows:4,columns:3},chartValues:[.013,.12]},null,2));
