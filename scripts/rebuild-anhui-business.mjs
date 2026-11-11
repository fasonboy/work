// Rebuild the screenshot's editable layout and map. Supply --reference for the actual photos.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/anhui-business',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__anhui_business__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1672,H=941,FONT='Noto Sans CJK SC',NAVY='#001849';
let source;
if(opt('--reference')) {
  const meta=await sharp(opt('--reference')).metadata();
  assert(Math.abs(meta.width/meta.height-W/H)<.01,'Reference must match this screenshot aspect ratio');
  source=await sharp(opt('--reference')).resize(W,H).png().toBuffer();
}
if(process.env.PPT_OPERATION_MARKED!=='1') execFileSync(process.env.RUNTIME_NODE,[path.join(skill,'container_tools/mark_artifact_operation_started.mjs'),'--operation-kind','create','--expected-output-count','1','--output-format','pptx'],{stdio:'inherit'});
const deck=Presentation.create({slideSize:{width:W,height:H}}),slide=deck.slides.add();
slide.background.fill='#ffffff';
const grad=(a,b,angleDeg=40)=>({type:'gradient',gradientKind:'linear',angleDeg,stops:[{offset:0,color:a},{offset:100000,color:b}]});
function box(name,x,y,w,h,fill='none',radius=0,line='none',geometry='rect') {
  return slide.shapes.add({name,geometry,position:{left:x,top:y,width:w,height:h},fill,line:{fill:line,width:line==='none'?0:1},...(radius?{borderRadius:radius}:{})});
}
function text(name,value,x,y,w,h,size=24,color=NAVY,bold=false,extra={}) {
  const o=box(name,x,y,w,h);o.text=value;o.text.style={typeface:FONT,fontSize:size,color,bold,autoFit:'none',insets:0,wrap:'none',verticalAlignment:'middle',...extra};return o;
}
function svgIcon(kind,x,y,size,color='#ffffff') {
  const paths={store:'<path d="M10 27v29h44V27M13 9h38l7 17c0 10-12 12-14 4-2 8-13 8-16 0-2 8-14 8-17 0-4 6-11 2-10-4z"/><path d="M18 48l8-7 8 6 7-3 7 8"/>',sofa:'<path d="M13 28v-8c0-17 38-17 38 0v8M14 43h36v12H14zM14 43v-7c-12-1-11-18 0-16 7 1 7 11 7 11h22s0-10 7-11c11-2 12 15 0 16v7M17 55v6m30-6v6"/>',wine:'<path d="M18 6h28l3 19c2 15-36 15-34 0zM17 20h30M32 37v22M22 59h20"/>',people:'<circle cx="27" cy="18" r="10"/><path d="M10 54v-9c0-17 34-17 34 0v9M45 9c14 0 14 22 0 22M51 39c9 0 11 8 11 15"/>',crown:'<path d="M9 20l12 13L32 8l11 25 12-13-6 32H15zM16 59h33"/><circle cx="8" cy="18" r="3"/><circle cx="32" cy="7" r="3"/><circle cx="56" cy="18" r="3"/>',heart:'<path d="M32 57C20 47 4 36 4 22 4 5 23 3 32 17 41 3 60 5 60 22c0 14-16 25-28 35z"/>',check:'<path d="M14 33l11 11 26-27"/>',pin:'<path d="M32 60S9 35 9 22a23 23 0 0 1 46 0c0 13-23 38-23 38z"/><circle cx="32" cy="22" r="9"/>'};
  return slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 66"><g fill="none" stroke="${color}" stroke-width="3.6" stroke-linejoin="round" stroke-linecap="round">${paths[kind]}</g></svg>`,alt:`${kind} vector icon`,position:{left:x,top:y,width:size,height:size}});
}
function curve(name,x,y,w,h,commands,fill) {
  return slide.shapes.add({geometry:'custom',name,position:{left:x,top:y,width:w,height:h},fill,line:{fill:'none',width:0},customPaths:[{width:w,height:h,commands}]});
}
const M=(x,y)=>({moveTo:{x,y}}),L=(x,y)=>({lineTo:{x,y}}),B=(x1,y1,x2,y2,x,y)=>({cubicBezTo:{x1,y1,x2,y2,x,y}}),Z=()=>({close:{}});
slide.background.fill=grad('#ffffff','#f8fbff',65);
curve('title-blue-slash',41,39,34,76,[M(15,0),L(34,0),L(19,76),L(0,76),Z()],grad('#008cff','#0046aa'));
text('editable-title','安徽业务展示',84,19,454,104,68,NAVY,true,{italic:true});
box('header-divider',555,38,1.5,73,'#6987d2');
text('editable-subheading','打造高品质的雪茄与威士忌品鉴空间',601,22,1010,59,40,'#072778',true);
text('editable-subtitle','融合品鉴 · 交流 · 休闲于一体的专业服务场所',602,80,1020,41,27,'#313b52');
curve('title-gold-swoosh',83,119,193,25,[M(0,12),B(62,-2,135,-4,193,14),B(109,5,53,11,3,25),Z()],grad('#ffe6a0','#f4a014',0));
curve('header-gold-swoosh',1124,80,322,52,[M(0,48),B(101,49,139,-33,322,15),B(165,-11,110,66,0,48),Z()],grad('#ffd978','#d99025',0));

const photos=[{x:31,y:175,w:692,h:411,label:'门店外观照片'},{x:737,y:175,w:532,h:411,label:'门店内部照片'}];
for(const [i,p] of photos.entries()) {
  if(source) {
    const crop=await sharp(source).extract({left:p.x,top:p.y,width:p.w,height:p.h}).png().toBuffer();
    slide.images.add({blob:crop,contentType:'image/png',alt:p.label,position:{left:p.x,top:p.y,width:p.w,height:p.h},geometry:'roundRect',borderRadius:23});
  } else {
    box(`photo-${i+1}-slot`,p.x,p.y,p.w,p.h,grad('#eff4fc','#dae6f7'),23);
    text(`photo-${i+1}-missing`,`${p.label}\n待补截图原文件`,p.x+30,p.y+151,p.w-60,106,28,'#758bac',false,{alignment:'center',wrap:'square'});
  }
}

// The screenshot's small map is a schematic reconstruction, not geocoded data.
const mx=1288,my=175;
box('map-background',mx,my,352,411,'#f5f6f5',23);
for(let row=0;row<5;row++)for(let col=0;col<4;col++) {
  const xx=mx+10+col*84,yy=my+9+row*77;
  box(`map-building-${row}-${col}`,xx,yy,74,65,row>1?'#e7f3fb':'#f0f1f1',3);
}
for(const [i,x,y,w,h] of [[1,1298,238,333,10],[2,1298,299,333,10],[3,1298,383,333,12],[4,1298,508,333,11],[5,1334,181,10,395],[6,1465,181,11,395],[7,1576,181,9,395]])box(`map-road-${i}`,x,y,w,h,'#ffffff');
box('map-river-bank',1298,398,333,12,'#bdf0b5');
curve('map-river',1298,408,333,29,[M(0,0),L(333,0),L(333,13),B(260,23,228,13,162,14),B(101,13,78,40,26,17),L(0,14),Z()],'#a4def7');
box('map-park',1346,443,72,65,'#cef5c2',3);
box('map-garden',1349,356,111,25,'#c8f2bd',2);
const mapText=[['road-main','汉兴大道',1439,375,102,23,13,'#596061'],['dijing','帝景天成',1500,343,83,24,13,'#1c4d84'],['plaza','蒂朗公园',1356,474,80,24,13,'#338747'],['center','合肥市中心',1456,479,107,25,13,'#df761f'],['road-east','东\n风\n北\n路',1465,516,19,61,11,'#596061']];
for(const [name,value,x,y,w,h,size,color] of mapText)text(`map-${name}`,value,x,y,w,h,size,color,true,{wrap:'square',lineSpacing:.95});
for(const [i,x,y,color] of [[1,1382,333,'#48b4f0'],[2,1525,341,'#48b4f0'],[3,1381,469,'#2db440'],[4,1490,474,'#ffab16'],[5,1539,467,'#48b4f0'],[6,1394,542,'#48b4f0']])box(`map-poi-${i}`,x-5,y-5,11,11,color,0,'none','ellipse');
box('map-address-callout',1404,217,216,70,grad('#126ae0','#002a77'),14);
svgIcon('pin',1418,232,33);
text('map-address-text','安徽省合肥市\n东风北路119号',1462,224,153,56,18.5,'#ffffff',true,{wrap:'square'});
curve('map-callout-pointer',1474,286,23,14,[M(0,0),L(23,0),L(11,14),Z()],'#01347f');
curve('map-main-pin',1471,304,29,38,[M(14,38),B(7,28,0,20,0,13),B(0,-4,29,-4,29,13),B(29,21,21,30,14,38),Z()],'#087bff');
box('map-pin-center',1479,311,12,12,'#ffffff',0,'none','ellipse');

const cards=[
 {x:31,w:536,iconX:74,divider:203,tx:231,label:'门店外观展示',detail:'现代雅致的门店形象',tag:'高端格调 · 辨识度强',kind:'store',tagW:264},
 {x:591,w:504,iconX:623,divider:753,tx:784,label:'门店内部环境',detail:'舒适私密的品鉴空间',tag:'专业氛围 · 尊享体验',kind:'sofa',tagW:263},
 {x:1119,w:521,iconX:1150,divider:1279,tx:1307,label:'门店地理位置',detail:'安徽省合肥市 · 东风北路119号',tag:'交通便利 · 商圈成熟',kind:'pin',tagW:275},
];
for(const [i,c] of cards.entries()) {
  box(`card-${i+1}`,c.x,615,c.w,179,grad('#fbfdff','#eaf2ff'),26,'#bad6ff');
  curve(`card-${i+1}-wave`,c.x+c.w-185,742,184,49,[M(0,49),B(75,37,114,1,184,0),L(184,49),Z()],'#bedcff');
  curve(`card-${i+1}-wave-light`,c.x+c.w-134,764,133,27,[M(0,27),B(73,0,90,0,133,0),L(133,27),Z()],'#d7e9ff');
  box(`card-${i+1}-badge`,c.iconX,641,99,95,grad('#0079f7','#00327f',50),24);
  svgIcon(c.kind,c.iconX+21,658,58);
  box(`card-${i+1}-divider`,c.divider,648,1.2,111,'#7295d0');
  text(`card-${i+1}-heading`,c.label,c.tx,625,c.x+c.w-c.tx-15,54,36,'#002374',true);
  text(`card-${i+1}-detail`,c.detail,c.tx,682,c.x+c.w-c.tx-12,38,i===2?21.5:25,'#2f3547');
  box(`card-${i+1}-tag-bg`,c.tx,728,c.tagW,45,grad('#fff2d5','#ffecb7',0),23);
  box(`card-${i+1}-tag-check-bg`,c.tx+18,739,23,23,'#df9300',0,'none','ellipse');
  svgIcon('check',c.tx+19,739,21);
  text(`card-${i+1}-tag-text`,c.tag,c.tx+52,731,c.tagW-58,38,21,'#28251d',true);
}
box('highlights-label',97,827,161,43,grad('#0759ca','#002267'),19);
text('highlights-title','核心亮点',120,829,125,38,25,'#ffffff',true);
const a=box('highlights-arrow-from',279,848,1,1),b=box('highlights-arrow-to',307,848,1,1);
slide.shapes.connect(a,b,{kind:'straight',line:{fill:'#f5c86f',width:7},head:{type:'none'},tail:{type:'triangle',width:'med',length:'sm'}}).bringToFront();
const highlights=[{x:383,kind:'wine',label:'精选雪茄与威士忌',w:243},{x:709,kind:'people',label:'专业品鉴与交流',w:225},{x:1025,kind:'crown',label:'私密舒适的环境',w:223},{x:1343,kind:'heart',label:'高品质会员服务',w:230}];
for(const [i,h] of highlights.entries()) {
  if(i)box(`highlight-${i}-divider`,h.x-47,832,1.3,33,'#ccd9ed');
  svgIcon(h.kind,h.x,829,38,'#103c83');
  text(`highlight-${i+1}-text`,h.label,h.x+49,830,h.w,38,24,'#0d337d',true);
}
// Each wave is an independent native path, kept within the slide canvas.
curve('footer-gold-wave',0,867,1672,74,[M(0,60),B(465,110,816,3,1335,43),B(1435,52,1544,-5,1672,1),L(1672,74),L(0,74),Z()],grad('#ffe69f','#d68d20',0));
curve('footer-light-blue-wave',0,883,1672,58,[M(0,1),B(430,113,914,-20,1330,34),B(1437,51,1561,-9,1672,1),L(1672,58),L(0,58),Z()],grad('#71b1ff','#236bd1',0));
curve('footer-deep-blue-wave',0,884,1672,57,[M(0,25),B(423,109,966,-22,1335,33),B(1452,40,1557,-11,1672,1),L(1672,57),L(0,57),Z()],grad('#0048aa','#00357e',0));
slide.speakerNotes.text=`来源：用户当前聊天“安徽业务展示”图片。保持1672×941画布。所有标题、说明、金色标签、卡片、底部亮点和波浪为独立可编辑对象；地图为按可见截图近似重绘，非导航底图，不能据此核验地址。${source?'两张门店照片由用户截图局部裁切嵌入，照片内部建筑、家具及招牌不变成可编辑形状。':'原始截图未以可读取文件提供，两张门店照片缺失；本文件为待补照片草稿，不能声称完整复原。'}小地图无法清晰读取的地名未猜写。`;
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-connectors','1','--min-target-arrows','1'],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,source?'anhui-business-editable.pptx':'anhui-business-editable-draft.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
await fs.writeFile(path.join(build,'status.json'),JSON.stringify({finalPath,missingPhotos:source?0:2,map:'approximate editable redraw',normalization:JSON.parse(normal)},null,2));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),missingPhotos:source?0:2},null,2));
