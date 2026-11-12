// Rebuild the online promotion expense screenshot as editable slide objects.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/online-promotion-budget',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1672,H=941,FONT='Noto Sans CJK SC',NAVY='#07325e',GREEN='#168b57',BLUE='#155999';
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
  coins:'<ellipse cx="45" cy="13" rx="17" ry="7"/><path d="M28 13v28c0 11 34 11 34 0V13M28 24c0 10 34 10 34 0M28 35c0 10 34 10 34 0"/><ellipse cx="21" cy="39" rx="17" ry="7"/><path d="M4 39v17c0 11 34 11 34 0V39M4 48c0 11 34 11 34 0"/>',
  screen:'<rect x="6" y="8" width="56" height="39" rx="2"/><path d="M8 53h52M27 53l-2 9M43 53l2 9M20 62h30M31 18l5 17 4-6 5-2zM41 29l6 7"/>',
  sprout:'<path d="M33 64V29M33 41C9 43 3 25 5 11c23 0 31 14 28 30zM34 32C34 8 54 3 66 5c-1 22-17 32-32 27zM16 23l17 18M34 32l20-15"/>',
  camera:'<rect x="3" y="12" width="44" height="40" rx="3"/><path d="M47 26l17-10v33L47 39z"/><path d="M20 24l13 9-13 9z" fill="#ffffff" stroke="#ffffff"/>',
  people:'<circle cx="34" cy="15" r="11"/><path d="M13 62V49c0-22 42-22 42 0v13zM12 8c-13 0-13 21 0 21M56 8c13 0 13 21 0 21M8 38c-9 0-11 9-11 20M60 38c9 0 11 9 11 20"/>',
  target:'<circle cx="31" cy="36" r="27"/><circle cx="31" cy="36" r="17"/><circle cx="31" cy="36" r="7"/><path d="M31 36L60 7M53 1l-2 13 13-2"/>',
};
function pict(kind,x,y,size,color=NAVY){
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="84" height="84" viewBox="-7 -5 84 84"><g fill="none" stroke="${color}" stroke-width="3.9" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} icon`,position:{left:x,top:y,width:size,height:size}});
}
slide=deck.slides.add();slide.background.fill=grad('#fcfdff','#f0f5ff',80);
box('header-rule',88,81,1497,4,grad('#cbe6ff','#007bda',0));
box('title-pill',473,39,721,87,grad('#2eb3da','#003b6a',15),46);
txt('page-title','线上推广费用规划',510,38,647,88,60,'#ffffff',true,{alignment:'center'});
txt('page-subtitle','通过线上流量获取、内容种草和私域运营，持续拓展目标消费人群。',341,135,1125,49,31.5,'#505050');

// Full circular frame is decorative; the source does not depict a 35% progress arc.
box('budget-ring',96,232,457,465,grad('#36c1ee','#0074ae',52),0,'ellipse');
box('budget-ring-inner',125,261,400,407,grad('#ffffff','#f5f9ff',50),0,'ellipse');
curve('budget-wave-back',183.6,600,282.8,68,[M(0,8.4),B(70,-8,100,58,151,35),B(207,18,243,-5,282.8,8.4),B(245.3,46.6,194.44,68,141.4,68),B(88.36,68,37.5,46.6,0,8.4),Z()],'#bddffb');
curve('budget-wave-front',225,615,200,53,[M(0,25.74),B(75,50,130,0,200,25.74),B(169.59,43.6,135.1,53,100,53),B(64.89,53,30.41,43.6,0,25.74),Z()],'#6bbdf2/70');
pict('coins',289,297,90,NAVY);
txt('budget-label','线上预算',238,375,239,65,47,NAVY,true);
box('budget-value-pill',149,450,361,74,grad('#2bb5dc','#0053a1',15),39);
txt('budget-value','24.5–28万元',183,449,314,75,52,'#ffffff',true);
txt('budget-share','占总推广预算约35%',191,536,326,53,31,'#575757');
const a=box('budget-to-allocation-from',590,464,1,1),b=box('budget-to-allocation-to',616,464,1,1);
slide.shapes.connect(a,b,{kind:'straight',line:{fill:'#a6d5fc',width:7},head:{type:'none'},tail:{type:'triangle',width:'sm',length:'sm'}}).bringToFront();

const allocations=[
 {y:226,title:'线上平台广告',pct:'45%',desc:'搜索推广、信息流广告、精准流量获取',kind:'screen',color:'#056aad',fill:'#d9eaff'},
 {y:355,title:'内容种草合作',pct:'25%',desc:'达人合作、产品科普、内容传播',kind:'sprout',color:'#00a663',fill:'#dff2ed'},
 {y:483,title:'短视频及直播',pct:'20%',desc:'视频制作、直播推广、内容引流',kind:'camera',color:'#028bca',fill:'#d9eaff'},
 {y:611,title:'私域会员运营',pct:'10%',desc:'社群活动、会员触达、老客复购',kind:'people',color:'#048cca',fill:'#d9eaff'},
];
for(const [i,r] of allocations.entries()){
  box(`allocation-${i+1}-surface`,715,r.y,872,110,grad('#ffffff',r.fill,0),19);
  box(`allocation-${i+1}-badge`,660,r.y+10,120,89,grad(i===1?'#4bd180':'#27b5de',i===1?'#009756':'#00549b',60),18);
  pict(r.kind,687,r.y+29,68,'#ffffff');
  box(`allocation-${i+1}-divider`,818,r.y+13,1.3,79,'#1b72b9');
  txt(`allocation-${i+1}-heading`,r.title,854,r.y-2,268,57,41,i===1?'#009c5a':'#07508e',true);
  box(`allocation-${i+1}-pct-pill`,1125,r.y+11,126,43,grad(i===1?'#31ba73':'#20a9d4',r.color,65),24);
  txt(`allocation-${i+1}-pct`,r.pct,1141,r.y+7,94,47,36,'#ffffff',true,{alignment:'center'});
  txt(`allocation-${i+1}-description`,r.desc,854,r.y+53,704,48,29.5,'#555555');
}

box('conclusion-panel',65,750,1541,78,grad('#ffffff','#dceaff'),41);
box('conclusion-label-bg',65,750,250,78,grad('#073e67','#06345b'),40);
box('conclusion-label-square',277,750,38,78,'#07395f');
pict('target',171,761,59,'#ffffff');
box('conclusion-divider',336,765,1.3,46,'#3f82b7');
txt('conclusion-text','以投产比和获客成本为核心，动态优化线上投放节奏。',366,759,1206,57,35,NAVY,true);
curve('footer-green',0,858,W,83,[M(0,52),B(530,-14,780,111,W,0),L(W,83),L(0,83),Z()],grad('#72ba91','#09b176'));
curve('footer-navy',0,869,W,72,[M(0,12),B(290,-49,535,114,1019,47),B(1330,40,1466,15,W,27),L(W,72),L(0,72),Z()],grad('#164878','#003462'));
slide.speakerNotes.text='来源：用户当前聊天“线上推广费用规划”截图，1672×941。标题、副标题、线上预算24.5–28万元、占总推广预算约35%、线上平台广告45%、内容种草合作25%、短视频及直播20%、私域会员运营10%以及各项说明和底部结论按图转录，未独立核验预算口径或更新数字。四项占比总和100%，是本页线上预算内部的分配；35%是总推广预算占比，两个口径分别保留。本页是独立单页，没有修改前一份5页运营方案中的约60%线上预算数字，原材料存在口径差异，不擅自改动。左侧完整蓝圈是装饰，不是35%进度图。所有文字为原生文本框，关系箭头为原生连接符；图标为近似矢量重绘并附PNG回退，波浪为独立原生形状，无整页截图背景。';
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-connectors','1','--min-target-arrows','1'],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,'online-promotion-budget-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),budget:'24.5–28万元',overallShare:'约35%',allocations:[45,25,20,10]},null,2));
