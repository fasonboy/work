// Rebuild the user-supplied short-video dashboard with editable PowerPoint objects.
// Usage: source scripts/ppt-runtime.sh; "$RUNTIME_NODE" scripts/rebuild-short-video.mjs [--reference /path/source.png]
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const option = key => { const i = args.indexOf(key); return i < 0 ? undefined : args[i + 1]; };
const reference = option('--reference');
const work = path.resolve(option('--work-dir') || path.join(project, '.local/short-video-rebuild', `run-${Date.now()}`));
const build = path.join(work, 'build');
const output = path.join(work, 'output');
await fs.mkdir(build, { recursive: true });
await fs.mkdir(output, { recursive: true });
const runtime = process.env.RUNTIME_NODE_MODULES;
const skill = process.env.PRESENTATIONS_SKILL_DIR;
assert(runtime && skill && process.env.RUNTIME_PYTHON && process.env.RUNTIME_NODE, 'First source scripts/ppt-runtime.sh');
const req = createRequire(path.join(runtime, '__dashboard__.cjs'));
const { Presentation, PresentationFile, FileBlob } = await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp = req('sharp');
const { finalizePresentation } = await import(pathToFileURL(path.join(skill, 'container_tools/artifact_tool_utils.mjs')).href);

// Keep the 1694 × 928 reference canvas rather than silently cropping to 16:9.
const W = 1694, H = 928;
let source;
if (reference) {
  const meta = await sharp(reference).metadata();
  assert(Math.abs(meta.width / meta.height - W / H) < 0.01, 'Reference aspect ratio differs from this dashboard');
  source = await sharp(reference).resize(W, H).png().toBuffer();
}
if (process.env.PPT_OPERATION_MARKED !== '1') {
  execFileSync(process.env.RUNTIME_NODE, [path.join(skill, 'container_tools/mark_artifact_operation_started.mjs'), '--operation-kind', 'create', '--expected-output-count', '1', '--output-format', 'pptx'], { stdio: 'inherit' });
}
const deck = Presentation.create({ slideSize: { width: W, height: H } });
const slide = deck.slides.add();
const C = { ink: '#07152f', muted: '#476397', blue: '#0754d4', line: '#d9e8ff', red: '#ed0043' };
const FONT = 'Noto Sans CJK SC';
const inventory = [];
const grad = (a,b,angleDeg=30) => ({ type: 'gradient', gradientKind: 'linear', angleDeg, stops: [{ offset: 0, color: a }, { offset: 100000, color: b }] });
function shape(name, x,y,w,h,fill='none',stroke='none',radius=0,geometry='rect') {
  return slide.shapes.add({ name, geometry, position: { left:x, top:y, width:w, height:h }, fill, line:{ fill:stroke, width:stroke==='none'?0:1 }, ...(radius ? { borderRadius:radius } : {}) });
}
function txt(name, content, x,y,w,h,size=22,color=C.ink,bold=false,extra={}) {
  const s=shape(name,x,y,w,h);
  s.text=content;
  s.text.style={ typeface:FONT, fontSize:size, color, bold, autoFit:'none', insets:0, wrap:'none', verticalAlignment:'middle', ...extra };
  inventory.push({name,text:content,x,y,width:w,height:h});
  return s;
}
const rule=(name,x,y,w,color=C.line)=>shape(name,x,y,w,1.3,color);
function icon(kind,x,y,size,color='#ffffff') {
  const paths={
    coins:`<ellipse cx="39" cy="14" rx="12" ry="6" fill="${color}"/><path d="M27 20c4 6 19 6 24 0v7c-5 6-20 6-24 0zm0 13c4 6 19 6 24 0v7c-5 6-20 6-24 0z" fill="${color}"/><ellipse cx="23" cy="32" rx="13" ry="6" fill="${color}"/><path d="M10 38c5 6 21 6 26 0v7c-5 6-21 6-26 0zm0 12c5 6 21 6 26 0v5c-5 6-21 6-26 0z" fill="${color}"/>`,
    cart:`<g fill="none" stroke="${color}" stroke-width="4.3" stroke-linecap="round" stroke-linejoin="round"><path d="M8 14h9l7 29h28l6-23H21"/><path d="M25 49h25"/></g><circle cx="27" cy="55" r="4" fill="${color}"/><circle cx="49" cy="55" r="4" fill="${color}"/>`,
    cursor:`<path d="M20 20l9 35 9-11 11 12 8-8-13-11 12-7z" fill="${color}"/><g stroke="${color}" stroke-width="3.4" stroke-linecap="round"><path d="M20 7v6M7 20h6M9 9l5 5M30 10l-3 5M10 30l5-3"/></g>`,
    eye:`<path d="M6 32Q32 0 58 32Q32 64 6 32z" fill="${color}"/><circle cx="32" cy="32" r="13" fill="#3b86f9"/><circle cx="32" cy="32" r="6" fill="${color}"/>`,
    bars:`<rect x="9" y="35" width="10" height="22" rx="2" fill="${color}"/><rect x="27" y="23" width="10" height="34" rx="2" fill="${color}"/><rect x="45" y="10" width="10" height="47" rx="2" fill="${color}"/>`,
    play:`<path d="M23 14L51 32 23 50z" fill="${color}"/>`,
    heart:`<path d="M32 55C16 44 5 35 5 22 5 6 23 4 32 17 41 4 59 6 59 22c0 13-11 22-27 33z" fill="${color}"/>`,
  };
  return slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64">${paths[kind]}</svg>`,alt:`${kind} icon`,position:{left:x,top:y,width:size,height:size}});
}
function medallion(kind,x,y,size,a='#62a5ff',b='#1773f5') {
  shape(`${kind}-badge-${x}-${y}`,x,y,size,size,grad(a,b,55),'none',0,'ellipse');
  icon(kind,x+size*.2,y+size*.2,size*.6);
}

slide.background.fill=grad('#ffffff','#f9fcff',75);
// Pale background accents stay separate and editable.
shape('header-ribbon-1',805,1,484,70,grad('#f4f9ff','#e8f3ff'),'none',36);
shape('header-ribbon-2',1240,0,343,74,grad('#eaf4ff','#f6faff'),'none',36);
const watermark=shape('video-watermark',1419,19,110,82,grad('#c4e1ff','#edf6ff'),'#d6eaff',20);
watermark.position={left:1419,top:19,width:110,height:82,rotation:-16};
icon('play',1448,38,44,'#ffffff');
txt('title','超级短视频总结',51,9,680,78,63,C.ink,true,{italic:true});
shape('title-underline',49,88,497,11,grad('#0879ff','#ffffff',0),'none',6);

const metrics=[
  {x:41,w:325,label:'推广花费',value:'0.98万',detail:'￥9,786.53',kind:'coins',bg:'#eff7ff',line:'#d5e7ff',a:'#67a7ff',b:'#247cf5',color:C.blue},
  {x:382,w:305,label:'成交笔数',value:'51',detail:'单笔成交成本 ￥191.89',kind:'cart',bg:'#fff7ef',line:'#ffead8',a:'#ff9828',b:'#ff6207',color:C.ink},
  {x:701,w:344,label:'点击转化率',value:'0.29%',detail:'全店整体 2.44%',kind:'cursor',bg:'#fcecf2',line:'#ffdfeb',a:'#ff6897',b:'#f52061',color:C.red},
];
for(const m of metrics) {
  shape(`${m.label}-card`,m.x,123,m.w,159,grad('#ffffff',m.bg),m.line,16);
  shape(`${m.label}-edge`,m.x+1,141,5,33,m.b,'none',3);
  if(m.label!=='点击转化率') for(let k=0;k<3;k++)shape(`${m.label}-mini-bar-${k}`,m.x+m.w-67+k*18,247-k*16,12,18+k*16,m.label==='推广花费'?'#d6eaff':'#ffe8d4','none',3);
  medallion(m.kind,m.x+15,159,76,m.a,m.b);
  txt(`${m.label}-label`,m.label,m.x+104,142,m.w-113,36,24,C.ink,true);
  txt(`${m.label}-value`,m.value,m.x+103,176,m.w-110,65,51,m.color,true);
  txt(`${m.label}-detail`,m.detail,m.x+105,240,m.w-110,29,m.label==='成交笔数'?17.6:19,C.muted,false);
}

shape('analysis-panel',41,300,1023,497,'#ffffff/80',C.line,19);
txt('analysis-heading','流量转化分析',75,313,460,59,36,C.ink,true);
txt('analysis-tagline','低点击成本 · 低成交转化',760,326,275,36,21,C.muted,true,{alignment:'right'});
rule('analysis-divider-top',75,374,960);

const flow=[
 {x:75,w:259,bg:'#eaf3ff',kind:'eye',label:'广告展现',value:'48.50万',detail:'485,036次',color:C.blue},
 {x:436,w:259,bg:'#eff7ff',kind:'cursor',label:'付费点击',value:'17,494',detail:'点击成本 ￥0.56',color:C.blue},
 {x:778,w:260,bg:'#fff0f4',kind:'cart',label:'成交',value:'51笔',detail:'成交成本 ￥191.89',color:C.red},
];
for(const f of flow) {
  shape(`${f.label}-surface`,f.x,400,f.w,146,grad(f.bg,'#ffffff',30),f.bg,16);
  medallion(f.kind,f.x+17,420,70,f.kind==='cart'?'#ff687d':'#65a3ff',f.kind==='cart'?'#f11e45':'#297bf5');
  txt(`${f.label}-label`,f.label,f.x+102,414,f.w-105,34,22,C.ink,true);
  txt(`${f.label}-value`,f.value,f.x+101,451,f.w-102,53,f.label==='广告展现'?41:43,f.color,true);
  txt(`${f.label}-detail`,f.detail,f.x+102,505,f.w-104,31,18.6,C.muted,false);
}
for(const [i,x,label,value,color] of [[1,356,'点击率','3.61%',C.blue],[2,713,'转化率','0.29%',C.red]]) {
  txt(`flow-${i}-rate-label`,label,x-4,415,60,25,16.5,color,true,{alignment:'center'});
  txt(`flow-${i}-rate-value`,value,x-4,435,60,24,16.5,color,true,{alignment:'center'});
  const a=shape(`arrow-${i}-source`,x+8,486,1,1), b=shape(`arrow-${i}-target`,x+37,486,1,1);
  slide.shapes.connect(a,b,{kind:'straight',fromSide:'right',toSide:'left',line:{fill:grad('#c3dcff','#4284ed',0),width:14},head:{type:'none'},tail:{type:'triangle',width:'med',length:'sm'}});
}
rule('analysis-divider-bottom',75,568,960);
txt('contribution-heading','对全店推广的贡献占比',75,578,720,46,29,C.ink,true);
const bars=[
 {y:635,label:'点击量占比',value:'29.5%',fill:grad('#2c7be9','#0757cf'),length:410,kind:'cursor',color:C.blue},
 {y:686,label:'花费占比',value:'7.1%',fill:grad('#8ca9d9','#7298d8'),length:119,kind:'coins',color:'#7e9bcf'},
 {y:738,label:'成交金额占比',value:'4.0%',fill:grad('#ff617f','#ed1754'),length:75,kind:'cart',color:C.red},
];
for(const bar of bars) {
  shape(`${bar.label}-icon-bg`,75,bar.y,44,42,bar.kind==='cart'?'#fff0f4':'#edf5ff','none',10);
  icon(bar.kind,81,bar.y+4,32,bar.kind==='cart'?C.red:C.blue);
  txt(`${bar.label}-label`,bar.label,140,bar.y+1,171,39,22,C.muted,true);
  shape(`${bar.label}-track`,303,bar.y+8,607,24,'#e9eff7','none',12);
  shape(`${bar.label}-bar`,303,bar.y+8,bar.length,24,bar.fill,'none',12);
  txt(`${bar.label}-value`,bar.value,932,bar.y-3,102,40,30,bar.color,true,{alignment:'right'});
}

shape('video-panel',1079,108,587,691,'#ffffff',C.line,20);
shape('video-panel-header',1079,108,587,75,grad('#187dff','#65acff',5),'none',20);
shape('video-panel-header-lower',1079,144,587,40,grad('#2586ff','#65acff',5));
shape('video-panel-interior',1080,167,585,630,'#ffffff','none',18);
shape('video-header-play-bg',1100,118,39,39,'#ffffff','none',0,'ellipse');
icon('play',1106,123,28,'#267bf4');
txt('video-heading','视频效果展示',1156,113,294,46,30,'#ffffff',true);
txt('video-heading-caption','精选优质视频素材',1468,123,176,29,18,'#edf6ff',false,{alignment:'right'});
const cards=[
 {x:1098,y:183,ph:250,title:'出发前必看！\n申根签证避坑指南',views:'12.8万',likes:'1.2万'},
 {x:1286,y:183,ph:250,title:'去欧洲办申根签证',views:'8.6万',likes:'8,264'},
 {x:1476,y:183,ph:250,title:'不知道申根签证保险\n怎么买？',views:'6.3万',likes:'6,321'},
 {x:1098,y:502,ph:227,title:'旅行保险怎么选？',views:'5.2万',likes:'4,851'},
 {x:1286,y:502,ph:227,title:'申根签证保险怎么买？',views:'4.8万',likes:'4,126'},
 {x:1476,y:502,ph:227,title:'正在办申根签证？\n保险可别买错！',views:'3.9万',likes:'3,746'},
];
for(const [i,c] of cards.entries()) {
  shape(`video-${i+1}-card`,c.x,c.y,173,c.ph+54,'#eef4fb',C.line,11);
  if(source) {
    const crop=await sharp(source).extract({left:c.x,top:c.y,width:173,height:c.ph-31}).png().toBuffer();
    slide.images.add({blob:crop,contentType:'image/png',alt:`原图视频缩略图 ${i+1}`,position:{left:c.x,top:c.y,width:173,height:c.ph-31},geometry:'roundRect',borderRadius:9});
  } else {
    shape(`video-${i+1}-unavailable-image`,c.x,c.y,173,c.ph-31,grad('#e0edfb','#c7dcf5',70),'none',10);
    txt(`video-${i+1}-source-status`,'视频画面\n待补原图',c.x+14,c.y+37,145,61,19,'#6d88ae',false,{alignment:'center',wrap:'square'});
    shape(`video-${i+1}-play-bg`,c.x+62,c.y+c.ph*.48,50,50,'#526b8d/55','#ffffff/90',0,'ellipse');
    icon('play',c.x+68,c.y+c.ph*.48+5,39);
    shape(`video-${i+1}-menu-bg`,c.x+133,c.y+10,32,22,'#506989/55','none',11);
    txt(`video-${i+1}-menu`,'···',c.x+135,c.y+5,28,28,20,'#ffffff',true,{alignment:'center'});
  }
  shape(`video-${i+1}-stats-bg`,c.x,c.y+c.ph-31,173,31,'#16253a');
  shape(`video-${i+1}-view-play`,c.x+9,c.y+c.ph-23,15,15,'#ffffff','none',0,'ellipse');
  icon('play',c.x+10,c.y+c.ph-22,13,'#26364a');
  txt(`video-${i+1}-views`,c.views,c.x+29,c.y+c.ph-28,65,25,14,'#ffffff');
  icon('heart',c.x+108,c.y+c.ph-23,16);
  txt(`video-${i+1}-likes`,c.likes,c.x+127,c.y+c.ph-28,44,25,13.5,'#ffffff');
  txt(`video-${i+1}-title`,c.title,c.x+11,c.y+c.ph+5,157,43,15.6,C.ink,true,{wrap:'square',verticalAlignment:'middle',lineSpacing:1.05});
}

shape('summary-banner',33,814,1632,88,grad('#eaf4ff','#f6faff',0),C.line,22);
medallion('bars',64,826,64,'#6cafff','#0062e5');
shape('summary-pill',143,827,199,55,grad('#0864eb','#5db5ff',0),'none',28);
txt('summary-period','9月投放总结',160,832,174,43,28,'#ffffff',true);
shape('summary-divider',367,842,1.5,31,'#76a0ff');
const summary=txt('summary-text','',392,829,1138,59,21,C.ink,true);
summary.text=[[
 {run:'超级短视频获取点击的',style:{bold:true,color:C.ink,typeface:FONT,fontSize:'21px'}},
 {run:'能力较强、点击成本较低',style:{bold:true,color:C.blue,typeface:FONT,fontSize:'21px'}},
 {run:'，可以有效的低成本获客，增加品牌用户人群，方便后续转化。',style:{bold:true,color:C.ink,typeface:FONT,fontSize:'21px'}},
]];
const sa=shape('summary-trend-start',1518,887,1,1),sb=shape('summary-trend-end',1627,834,1,1);
slide.shapes.connect(sa,sb,{kind:'straight',line:{fill:'#c6e0ff',width:13},head:{type:'none'},tail:{type:'triangle',width:'med',length:'sm'}}).bringToFront();
slide.speakerNotes.text=`来源：用户在当前聊天提供的“超级短视频总结”参考图。保持原图1694×928比例；按可见内容转录原始指标，不擅自更改数值。字体使用可用的 Noto Sans CJK SC 重建。${source?'六个视频画面来自原图局部裁切；仅这些缩略图保留为栅格。':'原始图片未挂载到云端，六个视频画面缺失，当前为明确标记的草稿，不能声称完整复原。'}视频观看量、点赞量和卡片标题独立可编辑。三条贡献条保留参考图的可见长度，它们是可编辑形状，不构造未经原图定义的额外坐标轴。`;

const raw=path.join(build,'raw.pptx');
const normalized=path.join(build,'portable.pptx');
await (await PresentationFile.exportPptx(deck)).save(raw);
const normalization=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',normalized,'--sharp-module',req.resolve('sharp'),'--min-native-connectors','3','--min-target-arrows','3'],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normalization);
const normalizedStats=JSON.parse(normalization);
assert.equal(normalizedStats.svgCount,normalizedStats.pngFallbackCount);
const finalPath=path.join(output,source?'超级短视频总结_可编辑.pptx':'超级短视频总结_可编辑草稿_待补缩略图.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:normalized,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const finalDeck=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const preview=await finalDeck.export({slide:finalDeck.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'slide-1.png'),Buffer.from(await preview.arrayBuffer()));
await fs.writeFile(path.join(build,'text-inventory.json'),JSON.stringify(inventory,null,2));
await fs.writeFile(path.join(build,'reconstruction-status.json'),JSON.stringify({finalPath,referenceProvided:!!source,missingVideoThumbnails:source?0:6,canvas:{width:W,height:H},editableTextBoxes:inventory.length,normalization:normalizedStats},null,2));
console.log(JSON.stringify({finalPath,preview:path.join(build,'slide-1.png'),missingVideoThumbnails:source?0:6},null,2));
