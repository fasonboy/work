// Rebuild the N9 slide with native editable packaging, not a substituted photograph.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/n9-product',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1676,H=938,FONT='Noto Sans CJK SC',INK='#07152f',ORANGE='#f58a00';
if(process.env.PPT_OPERATION_MARKED!=='1') execFileSync(process.env.RUNTIME_NODE,[path.join(skill,'container_tools/mark_artifact_operation_started.mjs'),'--operation-kind','create','--expected-output-count','1','--output-format','pptx'],{stdio:'inherit'});
const deck=Presentation.create({slideSize:{width:W,height:H}}),slide=deck.slides.add();
slide.background.fill='#ffffff';
const grad=(a,b,angleDeg=40)=>({type:'gradient',gradientKind:'linear',angleDeg,stops:[{offset:0,color:a},{offset:100000,color:b}]});
function box(name,x,y,w,h,fill='none',r=0,geometry='rect',stroke='none') {
  return slide.shapes.add({name,geometry,position:{left:x,top:y,width:w,height:h},fill,line:{fill:stroke,width:stroke==='none'?0:1},...(r?{borderRadius:r}:{})});
}
function txt(name,value,x,y,w,h,size=24,color=INK,bold=false,extra={}) {
  const o=box(name,x,y,w,h);o.text=value;o.text.style={typeface:FONT,fontSize:size,color,bold,autoFit:'none',insets:0,wrap:'none',verticalAlignment:'middle',...extra};return o;
}
const M=(x,y)=>({moveTo:{x,y}}),L=(x,y)=>({lineTo:{x,y}}),B=(x1,y1,x2,y2,x,y)=>({cubicBezTo:{x1,y1,x2,y2,x,y}}),Z=()=>({close:{}});
function curve(name,x,y,w,h,commands,fill,stroke='none',lineWidth=1) {
  return slide.shapes.add({geometry:'custom',name,position:{left:x,top:y,width:w,height:h},fill,line:{fill:stroke,width:stroke==='none'?0:lineWidth},customPaths:[{width:w,height:h,commands}]});
}
function poly(name,points,fill,stroke='none') {
  const xs=points.map(p=>p[0]),ys=points.map(p=>p[1]);
  const x=Math.min(...xs),y=Math.min(...ys),w=Math.max(...xs)-x,h=Math.max(...ys)-y;
  return curve(name,x,y,w,h,[M(points[0][0]-x,points[0][1]-y),...points.slice(1).map(p=>L(p[0]-x,p[1]-y)),Z()],fill,stroke);
}
function icon(kind,x,y,size,color=ORANGE) {
  const paths={
    shield:'<path d="M32 4l25 10v18c0 15-13 25-25 30C20 57 7 47 7 32V14z"/><path d="M20 32l9 9 17-19"/>',
    cycle:'<path d="M9 29C7 7 43 0 55 20M55 7v14H41M55 36c2 22-34 29-46 9M9 58V44h14"/>',
    glass:'<path d="M14 5h36l3 18c2 24-44 24-42 0zM32 42v19M21 61h23"/>',
    leaf:'<path d="M6 49C12 13 37 21 58 7c1 36-15 56-44 47M5 60l38-33"/>',
    check:'<circle cx="29" cy="34" r="23" opacity=".45"/><path d="M13 33l14 14 28-31"/>',
    parcel:'<path d="M4 16L30 3l29 14v32L31 63 4 49zM4 16l27 14 28-13M31 30v33M16 10l29 14v13"/>',
    people:'<circle cx="32" cy="17" r="10"/><path d="M14 59v-8c0-21 36-21 36 0v8M12 11c-14 0-14 21 0 21M52 11c14 0 14 21 0 21M8 40c-9 0-11 8-11 19M56 40c9 0 11 8 11 19"/>',
  };
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="70" height="70" viewBox="-5 -3 74 74"><g fill="none" stroke="${color}" stroke-width="4.3" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} icon`,position:{left:x,top:y,width:size,height:size}});
}
box('title-accent',66,36,14,92,grad('#0963ed','#0e32a3',90));
txt('page-title','台湾主营业务：N9胶囊',96,36,1220,73,50,'#101319',true);
box('title-divider',96,122,1518,1.4,'#bdcce5');
const nodes=[[1487,23,'#ffcd00'],[1510,23,'#ffbf00'],[1530,45,'#ff9f00'],[1540,36,'#ffae00'],[1557,25,'#ffab00'],[1572,49,'#ff9400'],[1550,61,'#ff6b00'],[1514,57,'#8fbd00'],[1496,43,'#a6c600']];
for(const [i,j] of [[0,1],[1,2],[2,3],[3,4],[4,5],[5,6],[6,2],[2,7],[7,8]]) {
  const [x1,y1,c]=nodes[i],[x2,y2]=nodes[j];
  slide.shapes.add({name:`logo-edge-${i}-${j}`,geometry:'line',position:{left:Math.min(x1,x2),top:Math.min(y1,y2),width:Math.abs(x2-x1)||.1,height:Math.abs(y2-y1)||.1,verticalFlip:(x2-x1)*(y2-y1)<0},fill:'none',line:{fill:c,width:4}});
}
for(const [i,[x,y,c]] of nodes.entries())box(`logo-node-${i}`,x-6,y-6,12,12,c,0,'ellipse');
const logo=txt('logo-wordmark','',1452,70,192,41,27,'#333333');
logo.text=[[{run:'JKN ',style:{typeface:FONT,fontSize:'27px',color:'#333333'}},{run:'Biotech',style:{typeface:FONT,fontSize:'27px',bold:true,color:'#333333'}}]];

curve('gold-background-arc',900,178,776,477,[M(0,414),B(144,112,342,-28,511,2),B(690,16,745,143,776,229),L(776,300),B(581,363,405,402,0,414),Z()],grad('#fff9f1','#ffe1b3',300));
curve('gold-background-ribbon',1450,423,226,170,[M(0,170),B(86,39,153,27,226,0),L(226,61),B(146,73,71,111,0,170),Z()],grad('#fff9ef','#ffdcab'));
const n9=txt('hero-n9','N9',65,170,160,132,101,'#ff8a00',true,{italic:true});
n9.text.fill=grad('#ffad00','#fa5b00',90);
txt('hero-product-name','牛饮无罪代谢胶囊',215,191,696,105,72,INK,true,{italic:true});
txt('hero-subtitle','饮前防护 | 饮中减负 | 饮后修复 | 轻松应对社交饮酒',65,301,876,54,30,'#3b3b3b');
box('hero-accent-rule',67,368,89,5,ORANGE,3);
const benefits=[['shield','守护肝脏',85],['cycle','促进代谢',310],['glass','减轻不适',535],['leaf','温和安全',760]];
for(const [i,[kind,label,x]] of benefits.entries()) {
  box(`benefit-${i+1}-surface`,x,405,154,144,grad('#fff5e8','#ffffff',90),0,'ellipse');
  icon(kind,x+47,421,60);
  txt(`benefit-${i+1}-text`,label,x+6,485,142,48,28,INK,true,{alignment:'center'});
  if(i<3)box(`benefit-${i+1}-divider`,x+189,443,1.4,57,'#c8cedb');
}
box('advantages-panel',64,584,874,193,grad('#fafbff','#f4f6fc'),21,'rect','#e0e6f2');
box('advantages-label-bg',84,616,116,133,grad('#316fb3','#101d3d',90),21);
txt('advantages-label','核心\n优势',97,632,91,100,33,'#ffffff',true,{alignment:'center',wrap:'square'});
const advantages=[
 {y:605,label:'多重专利成分',detail:'二重专利牛樟芝 × 韩国枳椇子 DHM 配合',size:20.8},
 {y:662,label:'高纯度ADH活性',detail:'日本大阪 98% 高纯度 GSH 搭配 Cys 协同代谢',size:19.1},
 {y:716,label:'专业研发支持',detail:'营养师 Renee 推荐，安全合规添加',size:21},
];
for(const [i,a] of advantages.entries()) {
  icon('check',230,a.y+8,33,'#ff9500');
  txt(`advantage-${i+1}-title`,a.label,271,a.y,187,46,25.5,INK,true);
  box(`advantage-${i+1}-divider`,460,a.y+13,1.3,21,'#a5aab7');
  txt(`advantage-${i+1}-description`,a.detail,472,a.y+1,453,45,a.size,'#30353f');
}
icon('parcel',73,813,57,'#f58b00');
txt('spec-heading','规格',146,806,112,44,26,INK,true);
txt('spec-value','10包/盒',146,846,112,35,21,'#1e222b');
box('spec-divider',264,817,1.3,54,'#bdc6d8');
icon('shield',307,812,58,'#f58b00');
txt('usage-heading','食用方式',373,804,244,44,26,INK,true);
txt('usage-value','饮前服用，24小时强化代谢',373,845,245,35,19.5,'#1e222b');
box('usage-divider',629,817,1.3,54,'#bdc6d8');
icon('people',664,812,61,'#f58b00');
txt('scenario-heading','适用场景',739,804,302,44,26,INK,true);
txt('scenario-value','聚餐应酬、商务社交、朋友聚会等',739,845,325,35,19.5,'#1e222b');

// Packaging is an editable vector reconstruction, not a photograph or exact print master.
curve('package-floor',996,730,616,174,[M(0,0),L(616,28),L(616,174),L(0,124),Z()],grad('#c8cad5/55','#ffffff/0',90));
function productBox(id,front,top,right,scale) {
  poly(`${id}-top`,top,grad('#53535a','#25252a',80),'#16161a');
  poly(`${id}-side`,right,grad('#49494e','#242428',0),'#131317');
  poly(`${id}-front`,front,grad('#202023','#09090c',85),'#101014');
  const [x,y]=front[0];
  function pt(name,value,dx,dy,w,h,size,color='#ffffff',bold=false,extra={}) {
    const o=txt(`${id}-${name}`,value,x+dx*scale,y+dy*scale,w*scale,h*scale,size*scale,color,bold,extra);
    o.position={left:x+dx*scale,top:y+dy*scale,width:w*scale,height:h*scale,rotation:6};
    return o;
  }
  const mark=pt('editable-n9','N9',17,3,224,112,108,'#ff6b12',true,{italic:true});
  mark.text.fill=grad('#fc8b25','#dd4209',90);
  pt('product-name','N9牛饮无罪代谢胶囊',19,119,209,23,13.1,'#ffffff',true);
  pt('english','Guilt-Free\nSupport\nCapsule',19,140,192,72,20,'#ffffff',true,{wrap:'square',lineSpacing:.95});
  pt('net-weight','NET:10包',205,199,71,25,11.8,'#ffffff');
  pt('japanese','夜\n飲\n暢\n快\nサ\nポ\nー\nト',245,19,20,119,12.3,'#ff8a00',false,{wrap:'square',lineSpacing:.91});
  const [tx,ty]=top[1];
  const t=txt(`${id}-top-slogan`,'SUPPORT YOUR NIGHT. RESET YOUR DAY.',tx+6,ty+9,240*scale,22*scale,8.2*scale,'#d2d2d9');
  t.position={left:tx+6,top:ty+9,width:240*scale,height:22*scale,rotation:5};
  // A few neutral printing lines suggest a side panel without inventing ingredients.
  const [rx,ry]=right[1];
  for(let i=0;i<9;i++) {
    const line=box(`${id}-side-print-${i}`,rx-29*scale,ry+45*scale+i*9*scale,22*scale,1.4,'#8c8c95/65');
    line.position={left:rx-29*scale,top:ry+45*scale+i*9*scale,width:22*scale,height:1.4,rotation:-40};
  }
}
productBox('lower-left',[[1018,535],[1265,565],[1265,772],[1018,737]],[[1018,535],[1088,482],[1321,506],[1265,565]],[[1265,565],[1321,506],[1321,715],[1265,772]],.91);
productBox('lower-right',[[1265,587],[1528,614],[1528,814],[1265,780]],[[1265,587],[1326,535],[1585,557],[1528,614]],[[1528,614],[1585,557],[1585,754],[1528,814]],.95);
productBox('upper',[[1133,329],[1406,353],[1406,571],[1133,542]],[[1133,329],[1206,282],[1461,300],[1406,353]],[[1406,353],[1461,300],[1461,518],[1406,571]],1);

slide.speakerNotes.text='来源：用户当前聊天的N9胶囊产品页截图。所有标题、卖点、核心优势、食用方式及规格按截图保留，仅用于版式重建，未独立验证功效、安全性、专利、推荐或饮酒相关宣传。右侧三盒包装使用可编辑原生形状与文字近似重绘，不是原产品照片，也不是可用于印刷的官方包装稿；N9字形纹理、透视、反光及不可辨认小字有省略或简化，未猜写配方。Logo、图标同为重绘。没有将整页截图作为背景。';
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp')],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,'n9-product-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),packaging:'native editable approximation'},null,2));
