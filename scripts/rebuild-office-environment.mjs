// Rebuild the office-photo layout; original photos must be supplied for a complete slide.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/office-environment',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1672,H=941,FONT='Noto Sans CJK SC',NAVY='#06265e';
const reference=opt('--reference'),officeImage=opt('--office-image'),receptionImage=opt('--reception-image');
let source;
if(reference) {
  const m=await sharp(reference).metadata();
  assert(Math.abs(m.width/m.height-W/H)<.01,'Screenshot aspect ratio differs from the office layout');
  source=await sharp(reference).resize(W,H).png().toBuffer();
}
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
function curve(name,x,y,w,h,commands,fill) {
  return slide.shapes.add({geometry:'custom',name,position:{left:x,top:y,width:w,height:h},fill,line:{fill:'none',width:0},customPaths:[{width:w,height:h,commands}]});
}
function icon(kind,x,y,size,color=NAVY) {
  const paths={
    monitor:'<rect x="4" y="7" width="60" height="39" rx="3"/><path d="M25 47v9M43 47v9M14 60h40"/>',
    building:'<path d="M22 65V7h26v58M4 65V27h18M48 65V27h16v38M0 65h68M29 17h11M29 28h11M29 39h11M29 50h11M10 36h6M10 46h6M54 36h5M54 46h5"/>',
    people:'<circle cx="34" cy="18" r="12"/><path d="M13 65V51c0-23 42-23 42 0v14zM13 13C-2 10-4 34 11 36M55 13c15-3 17 21 2 23M5 44c-7 2-7 10-7 21M63 44c7 2 7 10 7 21"/>',
    handshake:'<path d="M4 15l13-8 15 8 13-8 20 10 4 24-20 23M17 7L3 21l5 20 23 24M18 18l17 17c7 8 14-1 9-6l-6-7M32 15L20 29c-6 8 3 13 8 8l10-10M8 41l23 24M16 38l24 24M25 40l22 17M39 40l17 10"/>',
    diamond:'<path d="M17 5h35l17 20-35 40L0 25zM18 28l13 14 23-26"/>',
  };
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="80" height="80" viewBox="-7 -4 84 78"><g fill="none" stroke="${color}" stroke-width="3.8" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} office icon`,position:{left:x,top:y,width:size,height:size}});
}
curve('upper-right-blue-band',1335,0,337,199,[M(0,199),B(120,74,218,24,297,0),L(337,0),L(337,21),B(180,66,90,122,0,199),Z()],grad('#c3e1ff','#77b8fa',0));
curve('upper-right-wash',1335,20,337,179,[M(0,179),B(143,53,230,27,337,0),L(337,179),Z()],grad('#ffffff','#e7f3ff',0));
curve('footer-wash',0,844,1672,97,[M(0,24),B(445,83,1017,98,1672,0),L(1672,97),L(0,97),Z()],grad('#e2f1ff','#eef7ff',0));
curve('footer-left-blue',0,860,416,81,[M(0,0),B(123,15,259,50,416,81),L(0,81),Z()],grad('#77b6f9','#badfff',0));
const nodes=[[126,49,'#ffcd00'],[152,49,'#ffbf00'],[175,74,'#ff9f00'],[185,64,'#ffae00'],[203,52,'#ffab00'],[219,78,'#ff9400'],[195,91,'#ff6b00'],[154,86,'#8fbd00'],[135,69,'#a6c600']];
for(const [i,j] of [[0,1],[1,2],[2,3],[3,4],[4,5],[5,6],[6,2],[2,7],[7,8]]) {
  const [x1,y1,c]=nodes[i],[x2,y2]=nodes[j];slide.shapes.add({name:`logo-edge-${i}-${j}`,geometry:'line',position:{left:Math.min(x1,x2),top:Math.min(y1,y2),width:Math.abs(x2-x1)||.1,height:Math.abs(y2-y1)||.1,verticalFlip:(x2-x1)*(y2-y1)<0},fill:'none',line:{fill:c,width:4.5}});
}
for(const [i,[x,y,c]] of nodes.entries())box(`logo-node-${i}`,x-6,y-6,12,12,c,0,'ellipse');
const logo=txt('logo-wordmark','',88,102,184,43,28,'#333333');logo.text=[[{run:'JKN ',style:{typeface:FONT,fontSize:'28px',color:'#333333'}},{run:'Biotech',style:{typeface:FONT,fontSize:'28px',color:'#333333',bold:true}}]];
box('header-divider',296,44,1.5,88,'#9bacc6');
txt('company-title','台湾加康诺生物科技有限公司',341,16,1269,113,74,NAVY,true,{italic:true});
box('subtitle-line-left',516,159,146,1.6,'#b8c7df');box('subtitle-line-right',1010,159,146,1.6,'#b8c7df');
txt('photo-section-subtitle','办公环境实拍展示',692,133,292,55,35,'#838383',false,{alignment:'center'});
const photos=[{id:'office',x:49,w:739,label:'办公区',input:officeImage},{id:'reception',x:810,w:812,label:'会议与接待区',input:receptionImage}];
let missing=0;
for(const p of photos) {
  box(`${p.id}-photo-frame`,p.x,224,p.w,532,'#ffffff',31,'rect','#e2e7ee');
  let imageBytes;
  if(p.input) imageBytes=await fs.readFile(p.input);
  else if(source) imageBytes=await sharp(source).extract({left:p.x,top:224,width:p.w,height:532}).png().toBuffer();
  if(imageBytes) {
    const png=await sharp(imageBytes).png().toBuffer();
    slide.images.add({blob:png,contentType:'image/png',alt:`${p.label}实拍照片`,fit:'cover',position:{left:p.x,top:224,width:p.w,height:532},geometry:'roundRect',borderRadius:31});
  } else {
    missing++;
    box(`${p.id}-photo-slot`,p.x+1,225,p.w-2,530,grad('#f7faff','#eaf1fb'),30);
    txt(`${p.id}-photo-missing`,`${p.label}实拍照片待补\n请提供原始图片文件`,p.x+40,430,p.w-80,102,29,'#7b8fac',false,{alignment:'center',wrap:'square',lineSpacing:1.14});
  }
}
box('office-tag-bg',66,687,273,58,grad('#075ba8','#194c8c/0',0),20);
icon('monitor',85,695,44,'#ffffff');txt('office-tag','办公区',141,689,175,51,33,'#ffffff',true);
box('reception-tag-bg',830,687,335,58,grad('#034a92','#2b5d8c/0',0),20);
icon('people',850,696,46,'#ffffff');txt('reception-tag','会议与接待区',909,689,247,51,32,'#ffffff',true);
const strengths=[{x:160,kind:'building',label:'现代化办公环境',tx:239},{x:542,kind:'people',label:'专业服务团队',tx:624},{x:914,kind:'handshake',label:'高效协作沟通',tx:995},{x:1288,kind:'diamond',label:'舒适接待空间',tx:1360}];
for(const [i,s] of strengths.entries()) {
  icon(s.kind,s.x,802,63,NAVY);txt(`strength-${i+1}`,s.label,s.tx,807,258,52,29,NAVY,true);
}
for(const [i,x] of [[1,472],[2,845],[3,1218]])box(`strength-divider-${i}`,x,805,1.2,50,'#c8d7eb');
slide.speakerNotes.text=`来源：用户当前聊天“台湾加康诺生物科技有限公司—办公环境实拍展示”图片。标题、标签、底部说明和版式按可见内容重建；Logo、图标和蓝色曲线是近似矢量重绘。${missing?`当前缺少${missing}张可读取的实拍原图；这是明确标为待补照片的版式草稿，不使用生成或虚构的办公室照片替代。`:'两张照片由用户提供的文件或截图局部嵌入，照片内部家具、建筑和人物不因此成为可编辑对象。'}如从完整截图裁切，照片内原来烘焙的标签也会随照片保留；独立原照片可避免标签重复。`;
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp')],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,missing?'office-environment-editable-draft.pptx':'office-environment-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
await fs.writeFile(path.join(build,'status.json'),JSON.stringify({finalPath,missingPhotos:missing,referenceProvided:!!source,officeImageProvided:!!officeImage,receptionImageProvided:!!receptionImage},null,2));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),missingPhotos:missing},null,2));
