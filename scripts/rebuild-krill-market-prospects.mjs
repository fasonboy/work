// Rebuild the krill-oil market outlook with editable comparison cards and a native table.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/krill-market-prospects',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
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
poly('title-blue-slash',[[60,30],[74,30],[56,100],[42,100]],grad('#007ce1','#0059a7',90));
txt('page-title','南极磷虾油市场前景',88,9,767,105,70,NAVY,true,{italic:true});
txt('page-subtitle','鱼油消费教育基础与磷虾油差异化增长空间',52,113,807,52,34,'#303030');

const panels=[{x:16,w:691,n:'01',title:'鱼油：消费认知提升推动市场渗透',tx:155},{x:726,w:933,n:'02',title:'鱼油与磷虾油：磷虾油的差异化优势',tx:868}];
for(const [i,p] of panels.entries()) {
  box(`panel-${i+1}`,p.x,193,p.w,533,'#ffffff',26,'rect','#91c9fc',1.3);
  box(`panel-${i+1}-header`,p.x+1,194,p.w-2,66,grad('#004583','#0077ae',0),25);
  box(`panel-${i+1}-header-lower`,p.x+1,225,p.w-2,35,grad('#004583','#0077ae',0));
  curve(`panel-${i+1}-number-bg`,p.x+1,194,124,65,[M(26,0),L(124,0),L(97,65),L(0,65),B(2,29,12,2,26,0),Z()],grad('#00abf3','#0065cd',80));
  txt(`panel-${i+1}-number`,p.n,p.x+35,196,80,57,39,'#ffffff',true,{italic:true});
  txt(`panel-${i+1}-title`,p.title,p.tx,198,p.x+p.w-p.tx-17,55,i===0?33:35,'#ffffff',true);
}
for(const [i,x,w] of [[1,35,193],[2,267,196],[3,501,192]])box(`usage-card-${i}`,x,305,w,284,grad('#f1f9ff','#e1f1fe',65),28);
icon('fish',96,326,75);icon('heart',330,326,74);
for(const [i,x,w,period,pct] of [[1,35,193,'1999—2000年','1.3%'],[2,267,196,'2011—2012年','12%']]) {
  txt(`usage-${i}-period`,period,x+5,394,w-10,44,27,NAVY,true,{alignment:'center'});
  txt(`usage-${i}-population`,'美国成年人\n鱼油补充剂使用率',x+5,441,w-10,66,22,'#111111',false,{alignment:'center',wrap:'square',lineSpacing:1.25});
  txt(`usage-${i}-value`,pct,x+9,508,w-18,69,56,NAVY,true,{alignment:'center'});
}
for(const [i,x] of [[1,241],[2,475]]) {
  const a=box(`usage-arrow-${i}-start`,x,448,1,1),b=box(`usage-arrow-${i}-end`,x+19,448,1,1);
  slide.shapes.connect(a,b,{kind:'straight',line:{fill:grad('#c3e8ff','#55b6ff',0),width:5},head:{type:'none'},tail:{type:'triangle',width:'sm',length:'sm'}}).bringToFront();
}
txt('increase-heading','使用率提升',522,374,160,51,31,NAVY,true);
txt('increase-about','约',516,452,45,58,39,NAVY,true);
const nine=txt('increase-value','9',568,415,67,93,81,'#0067e9',true,{italic:true});nine.text.fill=grad('#0649ea','#00a1ff',90);
txt('increase-unit','倍',629,444,60,68,49,'#0067e9',true,{italic:true});
curve('increase-underline',523,515,153,16,[M(0,11),B(63,-1,109,-1,153,7),B(104,7,57,12,17,16),Z()],'#68b4ff');
box('left-conclusion-rule',48,620,630,1.4,'#d3e8fc');icon('trend',69,641,69);
box('left-conclusion-divider',165,646,1.4,50,'#1e4387');
txt('left-conclusion','鱼油普及带动Omega-3消费认知提升',192,645,490,55,28.5,NAVY,true);

const values=[
  ['对比维度','鱼油','南极磷虾油'],
  ['Omega-3\n形式','以甘油三酯型、\n乙酯型等形式存在','部分以磷脂\n形式存在'],
  ['营养成分','EPA、DHA','EPA、DHA、\n磷脂及天然虾青素'],
  ['产品特点','成熟的Omega-3\n补充来源','具有差异化的\n营养结构'],
];
const table=slide.tables.add({rows:4,columns:3,left:744,top:295,width:550,height:372,columnWidths:[166,196,188],values});
for(let r=0;r<4;r++) {
  table.rows[r].height=[56,102,106,108][r];
  for(let c=0;c<3;c++) {
    const cell=table.getCell(r,c);
    cell.fill=r===0?(c===1?grad('#d8efff','#b3ddf8'):grad('#00a3f5','#0070ce',90)):(c===0?grad('#fbfdff','#eef7ff'):c===2?grad('#fffcf5','#fff1d9'):'#ffffff');
    cell.text.style={typeface:FONT,fontSize:r===0?26:r===1&&c===0?21:r===2&&c===2?21.5:c===0?23:c===1?22.5:23,color:r===0?(c===1?NAVY:'#ffffff'):c===1?'#111111':NAVY,bold:r===0||c!==1,autoFit:'none',wrap:'square',alignment:r>0&&c===0?'left':'center',verticalAlignment:'middle'};
    table.cells.block({row:r,column:c,rowCount:1,columnCount:1}).assign({margins:{left:r>0&&c===0?60:7,right:6,top:7,bottom:7}});
  }
}
table.borders.assign({fill:'#152b48',width:1.1,style:'solid'});
table.cells.block({row:1,column:0,rowCount:1,columnCount:1}).assign({margins:{left:57,right:3,top:7,bottom:7}});
icon('drop',756,376,43);icon('molecules',756,482,43);icon('shield',756,589,43);

// Editable capsule illustration and independently editable ingredient labels.
box('ingredient-scene',1308,361,339,351,grad('#ffffff','#e5f6ff',90),30);
box('capsule-orbit',1320,391,319,181,grad('#f4fcff','#d7efff'),0,'ellipse','#ace0ff',6);
box('capsule-orbit-inner',1329,400,301,162,'none',0,'ellipse','#ffffff',2);
box('ingredient-platform',1310,647,337,51,grad('#f5fdff','#b4e7fc',90),0,'ellipse');
box('ingredient-platform-top',1333,651,289,33,grad('#ddf6ff','#ffffff'),0,'ellipse','#e7fbff',3);
function capsule(name,x,y,w,h,rotation) {
  const sh=box(`${name}-shadow`,x+4,y+h-4,w,18,'#c46c77/20',0,'ellipse');sh.position={left:x+4,top:y+h-4,width:w,height:18,rotation};
  const cap=box(name,x,y,w,h,grad('#ffb89c','#e01000',90),0,'ellipse','#d41100',2.6);cap.position={left:x,top:y,width:w,height:h,rotation};
  const dark=box(`${name}-body`,x+7,y+9,w-14,h-13,grad('#ff5736','#b90000',90),0,'ellipse');dark.position={left:x+7,top:y+9,width:w-14,height:h-13,rotation};
  const shine=box(`${name}-shine`,x+15,y+5,w-40,13,'#fff4e5/83',0,'ellipse');shine.position={left:x+15,top:y+5,width:w-40,height:13,rotation};
}
capsule('large-capsule-back',1463,425,143,63,23);capsule('large-capsule-front',1352,461,163,65,-15);
for(const [i,label,x,y,w,px,py] of [[1,'EPA',1334,324,108,1409,400],[2,'DHA',1519,324,109,1553,400],[3,'磷脂',1331,593,116,1402,554],[4,'天然虾青素',1487,593,155,1552,554]]) {
  const labelShape=box(`ingredient-label-${i}-bg`,x,y,w,46,grad('#ffffff','#def1ff',90),25,'rect','#b5dcff',2.5);
  txt(`ingredient-label-${i}`,label,x+6,y+1,w-12,43,label==='天然虾青素'?27:29,NAVY,true,{italic:i<3,alignment:'center'});
  const anchor=box(`ingredient-anchor-${i}`,px,py,1,1);
  slide.shapes.connect(labelShape,anchor,{kind:'straight',line:{fill:'#078cff',width:2.2},head:{type:'none'},tail:{type:'none'}}).bringToFront();
  box(`ingredient-dot-ring-${i}`,px-8,py-8,16,16,'#ffffff',0,'ellipse');box(`ingredient-dot-${i}`,px-4,py-4,8,8,'#078cff',0,'ellipse');
}

box('outlook-banner',22,762,1629,142,grad('#ffffff','#e0f1ff'),32);
curve('outlook-label-bg',22,784,314,101,[M(33,0),L(314,0),L(279,101),L(0,101),Z()],grad('#007ae2','#004590',35));
txt('outlook-label','市场前景',78,797,234,75,46,'#ffffff',true,{italic:true});
const separator=box('outlook-separator',363,797,1.5,70,'#264879');separator.position={left:363,top:797,width:1.5,height:70,rotation:4};
txt('outlook-line-1','鱼油的普及推动了Omega-3消费认知，南极磷虾油凭借磷脂型Omega-3',403,779,1020,51,31.5,NAVY,true);
txt('outlook-line-2','及天然虾青素形成差异化定位，具备',403,828,516,52,31.5,NAVY,true);
txt('outlook-emphasis','进一步增长潜力。',913,824,313,59,35,'#0066d5',true,{italic:true});
curve('outlook-emphasis-underline',916,879,273,15,[M(0,10),B(108,-4,197,-4,273,3),B(180,5,81,10,19,15),Z()],'#60aaff');
capsule('footer-capsule-back',1521,798,104,45,20);capsule('footer-capsule-front',1449,827,112,44,-16);
slide.speakerNotes.text='来源：用户当前聊天“南极磷虾油市场前景—鱼油消费教育基础与磷虾油差异化增长空间”截图。美国成年人鱼油补充剂使用率1999—2000年1.3%、2011—2012年12%、约9倍、鱼油及磷虾油营养形式与成分比较、市场前景结论按截图保留。未独立核验数据来源、营养或市场主张，未用精确比值替换截图约9倍。对比为4行3列原生PowerPoint表格；数据卡片为可编辑文字，未新增截图中不存在的图表。原图文件未在云端提供，雪山、水面、磷虾、胶囊和图标为近似重绘而非原始照片。';
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-connectors','6','--min-target-arrows','2','--min-native-tables','1'],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,'krill-market-prospects-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit','--require-native-table-slide','1'],explicitTotalSlideCount:1,requiredNativeTableOwnerSlides:[1],fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),nativeTable:{rows:4,columns:3}},null,2));
