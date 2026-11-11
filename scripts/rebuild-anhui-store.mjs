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
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/anhui-store',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__anhui__.cjs'));
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
  const paths={store:'<path d="M10 27v29h44V27M13 9h38l7 17c0 10-12 12-14 4-2 8-13 8-16 0-2 8-14 8-17 0-4 6-11 2-10-4z"/><path d="M18 48l8-7 8 6 7-3 7 8"/>',sofa:'<path d="M13 28v-8c0-17 38-17 38 0v8M14 43h36v12H14zM14 43v-7c-12-1-11-18 0-16 7 1 7 11 7 11h22s0-10 7-11c11-2 12 15 0 16v7M17 55v6m30-6v6"/>',pin:'<path d="M32 60S9 35 9 22a23 23 0 0 1 46 0c0 13-23 38-23 38z"/><circle cx="32" cy="22" r="9"/>'};
  return slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 66"><g fill="none" stroke="${color}" stroke-width="3.6" stroke-linejoin="round" stroke-linecap="round">${paths[kind]}</g></svg>`,alt:`${kind} vector icon`,position:{left:x,top:y,width:size,height:size}});
}
box('title-accent-shadow',61,52,25,94,'#eef4ff');
box('title-accent',68,52,15,94,'#0444f0');
text('editable-title','安徽线下门店展示',110,45,1060,90,58,NAVY,true);
box('title-divider',110,138,1506,2,'#bdd1ff');
const photos=[{x:35,y:212,w:600,h:451,label:'门店外观照片'},{x:648,y:212,w:507,h:451,label:'门店内部照片'}];
for(const [i,p] of photos.entries()) {
  if(source) {
    const crop=await sharp(source).extract({left:p.x,top:p.y,width:p.w,height:p.h}).png().toBuffer();
    slide.images.add({blob:crop,contentType:'image/png',alt:p.label,position:{left:p.x,top:p.y,width:p.w,height:p.h},geometry:'roundRect',borderRadius:24});
  } else {
    box(`photo-${i+1}-slot`,p.x,p.y,p.w,p.h,grad('#edf3fc','#dae7f9'),24);
    text(`photo-${i+1}-source-label`,`${p.label}\n待补截图原文件`,p.x+35,p.y+170,p.w-70,95,27,'#6b84a9',false,{alignment:'center',wrap:'square'});
  }
}

// Editable map redrawn from the screenshot, not an externally fetched navigation map.
const mx=1172,my=212;
box('editable-map-frame',mx,my,472,451,'#f6f7f6',24);
const mb=(name,x,y,w,h,fill,r=0,geometry='rect')=>box(`map-${name}`,mx+x,my+y,w,h,fill,r,'none',geometry);
const mt=(name,value,x,y,w,h,size=14,color='#124f86',bold=true,extra={})=>text(`map-${name}`,value,mx+x,my+y,w,h,size,color,bold,extra);
// Building parcels and a few dominant blocks follow the reference's visible arrangement.
for(const [i,x,y,w,h] of [[1,28,10,50,60],[2,86,12,78,60],[3,173,15,95,55],[4,285,15,91,57],[5,388,16,60,55],[6,35,87,80,51],[7,126,87,115,51],[8,258,86,170,52],[9,38,159,93,67],[10,153,155,86,68],[11,277,155,99,67],[12,388,155,62,67],[13,32,286,61,64],[14,221,294,96,69],[15,326,295,117,65],[16,25,387,120,49],[17,158,387,91,49],[18,277,386,124,50]]) mb(`parcel-${i}`,x,y,w,h,i>8?'#e6f2fa':'#f4f4f3',3);
for(const [name,x,y,w,h] of [['road-top',20,71,432,8],['road-center',18,140,434,11],['road-main',17,236,438,16],['road-south',17,373,438,12],['road-west',23,8,9,430],['road-middle',244,8,13,430],['road-east',377,8,10,430]]) mb(name,x,y,w,h,'#ffffff');
mb('road-main-centerline',17,242,438,2,'#d9dedf');
mb('park-north',105,210,144,22,'#c5f4c2',2);
mb('park-plaza',95,298,95,75,'#c9f4c3',3);
mb('park-river-top',30,252,425,9,'#b6ecba');
slide.shapes.add({geometry:'custom',name:'editable-map-river',position:{left:mx+29,top:my+263,width:426,height:30},fill:'#a2dffc',line:{fill:'none',width:0},customPaths:[{width:426,height:30,commands:[{moveTo:{x:0,y:0}},{lineTo:{x:426,y:0}},{lineTo:{x:426,y:23}},{cubicBezTo:{x1:329,y1:8,x2:240,y2:23,x:176,y:16}},{cubicBezTo:{x1:115,y1:12,x2:115,y2:38,x:65,y:22}},{lineTo:{x:0,y:15}},{close:{}}]}]});
const mapLabels=[
 ['factory','固镇县创新\n服装厂',13,13,100,42,14,'#124f86'],
 ['valves','蚌埠帆必得\n阀门有限公司',138,31,128,43,14,'#124f86'],
 ['clothing','登泰服饰\n有限公司',290,22,100,44,14,'#124f86'],
 ['ecommerce','固镇电子商\n务产业园',14,74,120,44,14,'#124f86'],
 ['tuanjie','团结路',80,130,80,26,14,'#4c4f50'],
 ['court','法院',10,171,55,30,15,'#ed1b2e'],
 ['investments','投资大厦',114,186,94,29,14,'#124f86'],
 ['kindergarten','固镇县示范\n幼儿园',313,134,101,43,14,'#be5d1c'],
 ['dijing','帝景天成',292,192,92,30,14,'#124f86'],
 ['xincheng','新城观邸',401,192,69,30,13,'#124f86'],
 ['hanxing','汉兴大道',213,224,108,29,15,'#4c4f50'],
 ['plaza','市民广场',113,335,94,30,15,'#2a883e'],
 ['mall','百大购物中心',229,349,130,29,14,'#c96d18'],
 ['duojin','多金中心',347,322,88,27,14,'#124f86'],
 ['hospital','固镇县\n医院住院部',407,322,65,45,13,'#124f86'],
 ['chaoyang-left','朝阳路',111,370,85,28,14,'#55595b'],
 ['chaoyang-right','朝阳路',357,370,90,28,14,'#55595b'],
 ['kaixuan','凯旋幸福里',12,390,109,27,14,'#124f86'],
 ['kaixiang','凯祥·时代公馆',107,418,147,25,14,'#124f86'],
 ['dongfeng','东\n风\n北\n路',244,381,20,63,12,'#55595b'],
 ['zhongyuan','中远翡翠居',356,421,112,24,14,'#124f86'],
];
for(const [name,value,x,y,w,h,size,color] of mapLabels)mt(name,value,x,y,w,h,size,color,true,{wrap:'square',alignment:name==='dongfeng'?'center':'left',lineSpacing:.96});
// Recreate visible POI dots and the main green location pin with editable objects.
for(const [i,x,y] of [[1,144,181],[2,327,196],[3,444,195],[4,142,331],[5,357,331],[6,50,386],[7,161,413],[8,350,435]]) {
  mb(`poi-${i}`,x-7,y-8,15,15,i===4?'#2caa36':'#5a9fe2',0,'ellipse');
}
slide.shapes.add({geometry:'custom',name:'editable-map-location-pin',position:{left:mx+249,top:my+146,width:38,height:52},fill:'#00ba55',line:{fill:'none',width:0},customPaths:[{width:38,height:52,commands:[{moveTo:{x:19,y:52}},{cubicBezTo:{x1:14,y1:37,x2:0,y2:28,x:0,y:19}},{cubicBezTo:{x1:0,y1:-5,x2:38,y2:-5,x:38,y:19}},{cubicBezTo:{x1:38,y1:30,x2:24,y2:38,x:19,y:52}},{close:{}}]}]});
box('map-pin-center',mx+259,my+157,17,17,'#ffffff',0,'none','ellipse');

const captions=[
 {x:35,w:576,iconX:115,divider:235,textX:268,label:'门店外观展示',detail:'实体经营场所',icon:'store'},
 {x:633,w:497,iconX:690,divider:808,textX:837,label:'门店内部环境',detail:'舒适专业的服务空间',icon:'sofa'},
 {x:1153,w:491,iconX:1192,divider:1298,textX:1327,label:'门店地理位置',detail:'安徽省合肥市 · 东风北路119号',icon:'pin'},
];
for(const [i,c] of captions.entries()) {
  box(`caption-card-${i+1}`,c.x,690,c.w,141,grad('#f6f9ff','#edf3fd'),27);
  box(`caption-badge-${i+1}`,c.iconX,715,86,86,grad('#0878eb','#003b99',50),0,'none','ellipse');
  svgIcon(c.icon,c.iconX+16,731,54);
  box(`caption-divider-${i+1}`,c.divider,724,1.5,53,'#1049bc');
  text(`caption-title-${i+1}`,c.label,c.textX,704,c.x+c.w-c.textX-12,56,36,NAVY,true);
  text(`caption-detail-${i+1}`,c.detail,c.textX,763,c.x+c.w-c.textX-10,39,i===2?21:24,'#343434');
}
slide.speakerNotes.text=`来源：用户当前聊天中的“安徽线下门店展示”截图。保留1672×941画布和原始可见文字。标题、文字、卡片和地图均为可编辑对象。地图按截图示意重绘，标签位置和细节有近似，不能用于导航或校验地址；原图标题/地址与地图地名可能不一致，未擅自改写。${source?'门店外观和室内两张照片从用户提供的原始截图裁切并嵌入，照片内部不是可编辑对象。':'本次云端没有可读取的原始截图，两张照片未恢复；当前输出仅为明确标记的待补照片草稿，不能声称完整复原。'}`;
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp')],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,source?'anhui-store-editable.pptx':'anhui-store-editable-draft.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
await fs.writeFile(path.join(build,'status.json'),JSON.stringify({finalPath,missingPhotos:source?0:2,map:'editable approximation redrawn from visible screenshot; not a navigation map',normalization:JSON.parse(normal)},null,2));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),missingPhotos:source?0:2},null,2));
