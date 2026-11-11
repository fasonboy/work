// Rebuild the private-resources page with editable device and phone illustrations.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/private-resources',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1676,H=939,FONT='Noto Sans CJK SC',INK='#06152d',ORANGE='#ff8a00';
if(process.env.PPT_OPERATION_MARKED!=='1') execFileSync(process.env.RUNTIME_NODE,[path.join(skill,'container_tools/mark_artifact_operation_started.mjs'),'--operation-kind','create','--expected-output-count','1','--output-format','pptx'],{stdio:'inherit'});
const deck=Presentation.create({slideSize:{width:W,height:H}}),slide=deck.slides.add();
slide.background.fill='#ffffff';
const grad=(a,b,angleDeg=40)=>({type:'gradient',gradientKind:'linear',angleDeg,stops:[{offset:0,color:a},{offset:100000,color:b}]});
function box(name,x,y,w,h,fill='none',r=0,geometry='rect',stroke='none',lineWidth=1) {
  return slide.shapes.add({name,geometry,position:{left:x,top:y,width:w,height:h},fill,line:{fill:stroke,width:stroke==='none'?0:lineWidth},...(r&&['rect','textbox','roundRect'].includes(geometry)?{borderRadius:r}:{})});
}
function txt(name,value,x,y,w,h,size=24,color=INK,bold=false,extra={}) {
  const o=box(name,x,y,w,h);o.text=value;o.text.style={typeface:FONT,fontSize:size,color,bold,autoFit:'none',insets:0,wrap:'none',verticalAlignment:'middle',...extra};return o;
}
const M=(x,y)=>({moveTo:{x,y}}),L=(x,y)=>({lineTo:{x,y}}),B=(x1,y1,x2,y2,x,y)=>({cubicBezTo:{x1,y1,x2,y2,x,y}}),Z=()=>({close:{}});
function curve(name,x,y,w,h,commands,fill,stroke='none',lineWidth=1) {
  return slide.shapes.add({geometry:'custom',name,position:{left:x,top:y,width:w,height:h},fill,line:{fill:stroke,width:stroke==='none'?0:lineWidth},customPaths:[{width:w,height:h,commands}]});
}
function icon(kind,x,y,size,color=ORANGE) {
  const paths={
    store:'<path d="M10 28v30h44V28M9 9h46l6 17c0 10-12 11-15 3-2 8-14 8-16 0-2 8-14 8-17 0-3 8-14 6-12-3zM28 43v15M41 43v15"/>',
    people:'<circle cx="26" cy="15" r="10"/><path d="M7 56V44c0-19 38-19 38 0v12zM46 9c15 0 15 20 0 20M51 37c10 0 12 8 12 19"/>',
    heart:'<path d="M32 58C18 47 5 34 5 21 5 5 23 3 32 17 41 3 59 5 59 21c0 13-13 26-27 37zM6 33h13l6-11 9 22 8-12h16"/>',
    thermometer:'<path d="M24 39V12a8 8 0 0 1 16 0v27a15 15 0 1 1-16 0zM32 19v29"/><circle cx="32" cy="50" r="6"/>',
    id:'<rect x="2" y="12" width="60" height="43" rx="5"/><circle cx="20" cy="27" r="6"/><path d="M10 45c0-13 20-13 20 0M39 25h14M39 34h14M39 43h10"/>',
    chat:'<path d="M8 43L3 60l17-7c35 9 56-16 38-37C39-6-5 6 3 34z"/><circle cx="21" cy="29" r="2"/><circle cx="33" cy="29" r="2"/><circle cx="45" cy="29" r="2"/>',
    dualchat:'<path d="M25 2C-2 2-6 32 10 42l-4 12 15-7c12 2 26-1 31-12C65 20 48 2 25 2z"/><path d="M46 28C21 28 19 56 43 59l15 4-3-10c14-13 6-25-9-25z"/><circle cx="18" cy="21" r="2"/><circle cx="34" cy="21" r="2"/><circle cx="40" cy="43" r="2"/><circle cx="52" cy="43" r="2"/>',
    headset:'<path d="M6 33C6-4 58-4 58 33M6 31h9v24H6c-7 0-7-24 0-24zM58 31h-9v24h9c7 0 7-24 0-24zM55 54c0 12-15 10-23 10"/><rect x="24" y="59" width="13" height="5" rx="2"/>',
  };
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="72" height="72" viewBox="-5 -5 76 76"><g fill="none" stroke="${color}" stroke-width="4.5" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} icon`,position:{left:x,top:y,width:size,height:size}});
}
curve('background-cream',1190,61,486,825,[M(0,275),B(190,32,361,-3,486,3),L(486,770),B(268,862,91,741,0,658),Z()],grad('#fffaf4','#fff0dd',0));
curve('footer-left-orange',0,886,324,53,[M(0,0),B(95,21,205,45,324,53),L(0,53),Z()],grad('#ffc27c','#ffe4c5',0));
curve('footer-right-orange',947,831,729,108,[M(0,108),B(231,80,506,48,729,0),L(729,108),Z()],grad('#fff9ee','#ff9600',0));
box('title-accent',61,43,14,78,grad('#ffaa10','#ff8200',90),7);
txt('page-title','线下私域资源展示',105,17,1295,92,63,INK,true);
txt('page-subtitle','整合线下设备与门店流量，搭建“健康服务 + 私域运营 + 在线支持”的闭环体系，提升用户粘性与转化',105,108,1465,49,26,'#343434');
box('title-rule',105,170,89,5,ORANGE,3);

box('offline-panel-header',53,231,644,112,grad('#ffbd00','#f67500',90),31);
box('offline-panel-body',53,307,644,580,'#ffffff',31,'rect','#fff0e8');
icon('store',100,239,61,'#ffffff');
txt('offline-heading','线下设备与服务资源',198,240,463,62,38,'#ffffff',true);
box('private-panel-header',753,231,400,112,grad('#ffbd00','#f67500',90),31);
box('private-panel-body',753,307,400,525,'#ffffff',31,'rect','#fff0e8');
icon('people',793,240,61,'#ffffff');
txt('private-heading','私域运营与支持',871,241,267,60,35.5,'#ffffff',true);

const resources=[{y:337,h:174,title:'血压仪',detail:'血压监测 · 健康评估 · 数据记录',kind:'heart'},{y:527,h:163,title:'医用恒温箱',detail:'药品/试剂存储 · 恒温保鲜 · 安全合规',kind:'thermometer'},{y:706,h:153,title:'中老年体检卡',detail:'健康体检 · 早筛评估 · 健康管理',kind:'id'}];
for(const [i,r] of resources.entries()) {
  box(`resource-${i+1}-surface`,78,r.y,609,r.h,grad('#fffaf4','#fff2e8'),27);
  box(`resource-${i+1}-illustration-bg`,78,r.y,252,r.h,grad('#e7d7b7','#fff3de',0),25);
  icon(r.kind,353,i===0?373:i===1?554:731,64);
  txt(`resource-${i+1}-title`,r.title,424,i===0?372:i===1?559:732,250,58,i===2?35:37,INK,true);
  txt(`resource-${i+1}-description`,r.detail,i===1?348:356,i===0?440:i===1?622:797,i===1?333:322,40,i===1?20.1:21.5,'#202020');
}
// Original equipment photos are unavailable: use clearly documented editable illustrations.
for(const [i,x,y,rot] of [[1,117,350,-30],[2,139,344,30],[3,108,375,-50]]) {
  const leaf=box(`monitor-scene-leaf-${i}`,x,y,16,39,'#75864b',0,'ellipse');leaf.position={left:x,top:y,width:16,height:39,rotation:rot};
}
box('monitor-scene-table',87,480,235,22,'#d9c6aa',8);
box('monitor-cuff',102,421,67,40,grad('#9da3a3','#c8cecc'),0,'ellipse');
for(let i=0;i<4;i++)box(`monitor-cuff-ring-${i}`,105+i*6,427,39,29,'none',0,'ellipse','#7a8281',1.4);
curve('monitor-tube',140,448,63,28,[M(0,0),B(13,27,48,28,63,2)],'none','#8d9593',4);
curve('monitor-body',161,373,136,129,[M(22,0),L(136,18),L(114,129),L(0,109),Z()],grad('#ffffff','#dce0df'), '#ced4d2');
curve('monitor-screen',188,393,82,66,[M(11,0),L(82,11),L(71,66),L(0,55),Z()],'#d0dcd3','#a8b8af');
for(const [i,label,y] of [[1,'SYS',401],[2,'DIA',417],[3,'PUL',433]]) {
  txt(`monitor-screen-label-${i}`,label,195,y,31,17,8,'#63736a');
  txt(`monitor-screen-value-${i}`,'—',233,y-3,26,22,16,'#37473f');
}
box('monitor-button',210,469,25,16,'#d9e3ed',8,'ellipse','#c8d3d8');
for(let i=0;i<5;i++)box(`fridge-room-curtain-${i}`,281+i*8,536,4,146,'#e8b87d/55');
box('fridge-side',154,542,108,135,'#c5cccd',4);
box('fridge-cabinet',144,543,109,131,grad('#f7f8f5','#d9e0de'),4,'rect','#c5cecd');
box('fridge-controller',156,550,84,13,'#f4f4f1',2);
txt('fridge-controller-label','TEMP',162,550,30,12,6,'#7e8c8d');
box('fridge-controller-screen',218,552,16,7,'#c2d7dd',1);
box('fridge-glass',157,569,83,89,grad('#a5c3cf','#e5f1ed'),2,'rect','#91acb6');
for(let row=0;row<3;row++) {
  box(`fridge-shelf-${row}`,161,591+row*26,74,3,'#dce9e6');
  for(let col=0;col<3;col++) {
    box(`fridge-bottle-${row}-${col}`,165+col*23,577+row*26,12,14,'#f5fcfc',2);
    box(`fridge-bottle-label-${row}-${col}`,166+col*23,581+row*26,10,5,col===1?'#b4dce6':'#c2e8de');
  }
}
box('fridge-handle',243,600,4,27,'#9baaaa',2);
box('fridge-foot-left',154,674,13,6,'#7e8989',1);box('fridge-foot-right',232,674,13,6,'#7e8989',1);

for(let i=0;i<5;i++)box(`consultation-wall-line-${i}`,276+i*8,714,2,126,'#e5c89b/60');
box('consultation-tv',129,733,80,45,'#818b86',2);box('consultation-tv-display',133,737,72,37,'#d7eaf0');
box('consultation-cross-v',146,747,5,17,'#d59783');box('consultation-cross-h',140,753,17,5,'#d59783');
box('consultation-tv-line-1',164,747,31,3,'#94b9c7');box('consultation-tv-line-2',164,754,27,2,'#a1c0cb');
box('consultation-chair-left',127,801,43,36,'#aaa99d',10);box('consultation-chair-right',267,801,39,36,'#a7a798',10);
box('consultation-table-top',163,801,107,20,'#fffdf1',0,'ellipse');
box('consultation-table-leg-left',183,817,4,29,'#b6aa94');box('consultation-table-leg-right',247,817,4,29,'#b6aa94');
box('consultation-pot',278,768,17,24,'#dbd1b9',3);
for(const [i,x,y,rot] of [[1,275,746,-30],[2,288,738,30],[3,287,751,50]]) {
  const leaf=box(`consultation-plant-${i}`,x,y,12,31,'#72834b',0,'ellipse');leaf.position={left:x,top:y,width:12,height:31,rotation:rot};
}

box('private-chat-card',770,342,367,217,grad('#fffaf4','#fff2e8'),29);
icon('dualchat',814,380,80);txt('private-chat-heading','私信沟通',919,391,202,59,35,INK,true);
txt('private-chat-details','精准触达 · 一对一沟通\n需求解答 · 转化跟进',817,460,293,82,25,'#202020',false,{alignment:'center',wrap:'square',lineSpacing:1.24});
box('support-card',770,580,367,216,grad('#fffaf4','#fff2e8'),29);
icon('headset',816,605,77);txt('support-heading','在线客服',919,619,202,59,35,INK,true);
txt('support-details','实时响应 · 专业咨询\n售后支持 · 会员维护',817,687,293,80,25,'#202020',false,{alignment:'center',wrap:'square',lineSpacing:1.24});

for(const [gap,x] of [[1,709],[2,1174]]) for(let i=0;i<2;i++) {
  const a=box(`flow-${gap}-${i}-from`,x+i*17,547,1,1),b=box(`flow-${gap}-${i}-to`,x+i*17+10,547,1,1);
  slide.shapes.connect(a,b,{kind:'straight',fromSide:'right',toSide:'left',line:{fill:ORANGE,width:5},head:{type:'none'},tail:{type:'triangle',width:'med',length:'sm'}}).bringToFront();
}

box('community-label-bg',1227,185,398,60,grad('#ffad00','#ff8100',90),30);
icon('chat',1274,191,48,'#ffffff');txt('community-heading','私域社群',1341,189,159,51,35,'#ffffff',true);
txt('community-example','（示例）',1495,196,117,43,25,'#ffffff',true);
box('phone-outer-shadow',1227,259,397,646,'#ccc8c3',65);box('phone-outer',1230,261,391,641,'#161616',63);
box('phone-rim',1236,267,379,629,'#e4e4e2',59);box('phone-screen',1243,274,365,614,'#ffffff',53);
box('phone-header-bg',1248,305,355,51,'#fafafa',0);
box('phone-camera-pill',1384,278,84,25,'#040404',13);box('phone-camera',1454,286,7,7,'#101b29',0,'ellipse');
curve('phone-back',1259,319,13,21,[M(11,0),L(1,10),L(11,20)],'none','#161616',2.3);
txt('phone-title','群聊',1374,309,100,40,24,'#111111',true,{alignment:'center'});
box('phone-search-circle',1530,320,15,15,'none',0,'ellipse','#161616',2);curve('phone-search-tail',1542,331,9,10,[M(0,0),L(9,10)],'none','#161616',2);
box('phone-add-circle',1571,317,23,23,'none',0,'ellipse','#161616',1.7);box('phone-add-h',1576,328,13,1.7,'#161616');box('phone-add-v',1582,322,1.7,13,'#161616');
box('phone-list-divider',1249,356,353,1,'#eeeeee');
const groups=[
 {title:'【群公告领券】…',preview:'[26条] 1点点大连福利官…',time:'10:39',color:'#d7cfb6'},
 {title:'大连东港商务区美…',preview:'你已添加了大连东港商务区美…',time:'10:14',color:'#00b66d'},
 {title:'青囊客服群',preview:'[22条] 心理-情感咨询助理…',time:'08:56',color:'#c9b19c'},
 {title:'宁珊旅行-慧珊',preview:'皇家双十一终极促销上…',time:'昨天 22:11',color:'#8e9b8b'},
 {title:'美团按摩洗…',preview:'[2条] 小周周：[小程序] 足…',time:'昨天 19:21',color:'#bbd6bd'},
 {title:'鲁华南极磷虾油…',preview:'双11狂欢购 | 鲁华南…',time:'昨天 18:54',color:'#d2c5aa'},
 {title:'麦知名品商标专…',preview:'[图片]',time:'昨天 18:46',color:'#afbdc6'},
 {title:'A2区19号楼…',preview:'[2条] 公园上城…',time:'昨天 18:14',color:'#cbb6a0'},
];
for(const [i,g] of groups.entries()) {
  const y=367+i*64;
  box(`chat-${i+1}-avatar`,1273,y,45,45,g.color,5);
  if(i===1) {
    box('chat-green-avatar-bubble',1282,y+11,27,19,'#ffffff',7);box('chat-green-avatar-dot-1',1288,y+18,3,3,g.color,0,'ellipse');box('chat-green-avatar-dot-2',1299,y+18,3,3,g.color,0,'ellipse');
  } else for(let k=0;k<4;k++)box(`chat-${i+1}-avatar-tile-${k}`,1277+(k%2)*20,y+4+Math.floor(k/2)*20,17,17,k%2?'#d8ded4':i%2?'#87908a':'#b0a68f',2);
  if(i!==1)box(`chat-${i+1}-unread`,1314,y-3,8,8,'#ff3939',0,'ellipse');
  txt(`chat-${i+1}-group-name`,g.title,1335,y-4,208,30,19,'#111111',true);
  txt(`chat-${i+1}-preview`,g.preview,1335,y+27,205,24,14,'#777777');
  txt(`chat-${i+1}-time`,g.time,1543,y,55,24,11,'#8c8c8c',false,{alignment:'right'});
  box(`chat-${i+1}-muted-1`,1570,y+30,8,8,'#d4d4d4',0,'ellipse');box(`chat-${i+1}-muted-2`,1581,y+30,8,8,'#d4d4d4',0,'ellipse');
  if(i<7)box(`chat-${i+1}-divider`,1335,y+58,266,1,'#eeeeee');
}
box('phone-home-indicator',1375,879,108,4,'#242424',2);
slide.speakerNotes.text='来源：用户当前聊天“线下私域资源展示”截图。标题、资源名称、服务说明和流程关系按可见内容重建。原图文件未提供，因此血压计、恒温箱、体检场景及手机列表均为可编辑矢量示意，不是原始照片或真实截图。血压计屏幕用横线代替无法确认的测量数值，不表示真实读数。手机头像使用简化色块；群名称和预览仅转录可辨认部分并用省略号表示截断，不补全未知消息或人物身份。右上保留“示例”标识。去除了标题截图中的校对红色波浪线，不作为设计元素复现。';
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-connectors','4','--min-target-arrows','4'],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,'private-resources-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),illustrations:'native schematic redraws, not original photos or chat screenshots'},null,2));
