// Rebuild the four event and market-channel screenshots as editable slides.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/market-channels-events',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation,makeNativeBulletParagraphs}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1672,H=941,FONT='Noto Sans CJK SC',NAVY='#133f70',RED='#e34442',BLUE='#347fc1';
if(process.env.PPT_OPERATION_MARKED!=='1') execFileSync(process.env.RUNTIME_NODE,[path.join(skill,'container_tools/mark_artifact_operation_started.mjs'),'--operation-kind','create','--expected-output-count','1','--output-format','pptx'],{stdio:'inherit'});
const deck=Presentation.create({slideSize:{width:W,height:H}});let slide,scale=1;
const grad=(a,b,angleDeg=40)=>({type:'gradient',gradientKind:'linear',angleDeg,stops:[{offset:0,color:a},{offset:100000,color:b}]});
function box(name,x,y,w,h,fill='none',r=0,geometry='rect',stroke='none',lineWidth=1) {
  return slide.shapes.add({name,geometry,position:{left:x*scale,top:y*scale,width:w*scale,height:h*scale},fill,line:{fill:stroke,width:stroke==='none'?0:lineWidth*scale},...(r&&['rect','textbox','roundRect'].includes(geometry)?{borderRadius:r*scale}:{})});
}
function txt(name,value,x,y,w,h,size=24,color=NAVY,bold=false,extra={}) {
  const o=box(name,x,y,w,h);o.text=value;o.text.style={typeface:FONT,fontSize:size*scale,color,bold,autoFit:'none',insets:0,wrap:'none',verticalAlignment:'middle',...extra};return o;
}
const M=(x,y)=>({moveTo:{x,y}}),L=(x,y)=>({lineTo:{x,y}}),B=(x1,y1,x2,y2,x,y)=>({cubicBezTo:{x1,y1,x2,y2,x,y}}),Z=()=>({close:{}});
function curve(name,x,y,w,h,commands,fill,stroke='none',lineWidth=1) {
  return slide.shapes.add({geometry:'custom',name,position:{left:x*scale,top:y*scale,width:w*scale,height:h*scale},fill,line:{fill:stroke,width:stroke==='none'?0:lineWidth*scale},customPaths:[{width:w,height:h,commands}]});
}
function poly(name,pts,fill) {
  const x=Math.min(...pts.map(p=>p[0])),y=Math.min(...pts.map(p=>p[1])),w=Math.max(...pts.map(p=>p[0]))-x,h=Math.max(...pts.map(p=>p[1]))-y;
  return curve(name,x,y,w,h,[M(pts[0][0]-x,pts[0][1]-y),...pts.slice(1).map(p=>L(p[0]-x,p[1]-y)),Z()],fill);
}
const bulletSpecs={};
function bullets(name,items,x,y,w,h,size=24,bold=false,spacing=10){
  if(name.startsWith('competition-'))items=items.map(t=>t.replaceAll('\n',''));
  const o=box(name,x,y,w,h);bulletSpecs[name]=items;
  o.text=makeNativeBulletParagraphs(items.map(t=>t.replaceAll('\n','')),{marginLeftPoints:22*scale,hangingPoints:16*scale,spaceAfterPoints:spacing*scale,bulletCharacter:'•'});
  o.text.style={typeface:FONT,fontSize:size*scale,color:NAVY,bold,autoFit:'none',wrap:'square',insets:0,verticalAlignment:'top',lineSpacing:1.18};return o;
}
const paths={
 people:'<circle cx="34" cy="15" r="11"/><path d="M13 62V49c0-22 42-22 42 0v13zM12 8c-13 0-13 21 0 21M56 8c13 0 13 21 0 21M8 38c-9 0-11 9-11 20M60 38c9 0 11 9 11 20"/>',
 pair:'<circle cx="17" cy="16" r="10"/><circle cx="49" cy="16" r="10"/><path d="M-3 61c0-28 40-28 40 0M29 61c0-28 40-28 40 0"/>',
 user:'<circle cx="34" cy="17" r="13"/><path d="M5 64c0-36 58-36 58 0"/>',
 chat:'<rect x="7" y="8" width="53" height="37" rx="6"/><path d="M24 45L14 58l23-13M18 21h30M18 32h22"/>',
 bottle:'<rect x="13" y="17" width="41" height="47" rx="7"/><path d="M24 16V4h20v12M22 40h24"/>',
 gift:'<path d="M5 24h58v12H5zM10 36v28h48V36M34 23v41M34 22C6 25 10-9 26 5l8 17c28 3 24-31 8-17z"/>',
 handshake:'<path d="M5 39l16-18 13 3 13-3 16 18-12 14-18-17-13 15zM26 29l-7 9 7 7 9-8M42 44l-8 10M50 51l-9 9"/>',
 network:'<circle cx="12" cy="15" r="9"/><circle cx="55" cy="15" r="9"/><circle cx="34" cy="53" r="9"/><path d="M21 15h25M17 23l13 22M50 23L38 45"/>',
 cart:'<path d="M5 5h10l9 39h32l9-28H19M24 53h31"/><circle cx="28" cy="60" r="4"/><circle cx="53" cy="60" r="4"/>',
 video:'<rect x="4" y="8" width="60" height="46" rx="5"/><path d="M28 21l17 10-17 11zM14 63h40"/>',
 search:'<circle cx="28" cy="28" r="23"/><path d="M45 45l22 22"/>',
 store:'<path d="M10 28v33h45V28M8 8h49l6 17c0 11-13 12-16 3-3 9-15 9-17 0-3 9-15 9-17 0-3 9-15 7-12-3zM28 44h15v17"/>',
 medical:'<rect x="6" y="15" width="56" height="46" rx="3"/><path d="M25 15V5h19v10M34 28v22M23 39h22"/>',
 shield:'<path d="M34 4l26 10v20c0 15-13 26-26 32C21 60 8 49 8 34V14zM22 34l10 10 18-22"/>',
 layers:'<path d="M5 21L34 5l29 16-29 16zM6 34l28 16 28-16M6 47l28 16 28-16"/>',
 target:'<circle cx="31" cy="36" r="27"/><circle cx="31" cy="36" r="17"/><circle cx="31" cy="36" r="7"/><path d="M31 36L60 7M53 1l-2 13 13-2"/>',
 trophy:'<path d="M15 5h36l-3 28c-2 14-9 20-15 20s-13-6-15-20zM33 53v10M20 65h26M15 10H4C1 30 7 40 21 41M51 10h11c3 20-3 30-17 31"/>',
 trend:'<path d="M6 32L24 17l13 10L63 4M48 4h15v15M7 64V47h11v17zM29 64V37h11v27zM51 64V26h11v38z"/>',
 factory:'<path d="M5 63V29l16 10V25l17 11V17h14v46zM12 49h5M27 49h5M42 49h5M10 29V8h10v27M44 17V7h10v10"/>',
 tag:'<path d="M5 38L35 7h27v28L31 65z"/><circle cx="51" cy="19" r="4"/>',
 laptop:'<rect x="12" y="8" width="44" height="36" rx="2"/><path d="M12 44L3 57h62l-9-13M25 52h18"/>',
};
function pict(kind,x,y,size,color=NAVY){
  assert(paths[kind],kind);
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="84" height="84" viewBox="-7 -5 84 84"><g fill="none" stroke="${color}" stroke-width="4.2" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} icon`,position:{left:x*scale,top:y*scale,width:size*scale,height:size*scale}});
}
function portrait(kind,x,y,w,h,color){
  const hair=kind==='woman'?'<path d="M22 40C10 7 64-2 60 38l6 26-18-2M22 38l-7 25 16-1"/>':'<path d="M21 32C11 1 64 0 59 33L50 20 35 23 25 19z" fill="COLOR"/>';
  const clothes=kind==='office'?'<path d="M16 91l6-30 16 15 6-15 17 30M37 68l4 9-4 13"/>':'<path d="M9 91V78c0-25 58-25 58 0v13M28 61l12 12 11-12M39 73v17"/>';
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="84" height="100" viewBox="0 0 80 100"><g fill="#ffffff" stroke="${color}" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round">${hair.replaceAll('COLOR',color)}<path d="M22 31c-3 41 39 41 37 0M30 54v10M49 54v10"/><path d="M28 36h3M48 36h3M36 46h8"/>${clothes}</g></svg>`,alt:`${kind} audience illustration`,position:{left:x*scale,top:y*scale,width:w*scale,height:h*scale}});
}
function newSlide(note,s=1){scale=s;slide=deck.slides.add();slide.background.fill='#fbfdff';slide.speakerNotes.text=note;}
function arrow(name,x1,y1,x2,y2,color='#86bddd',width=5){const a=box(`${name}-from`,x1,y1,1,1),b=box(`${name}-to`,x2,y2,1,1);slide.shapes.connect(a,b,{kind:'straight',line:{fill:color,width:width*scale},head:{type:'none'},tail:{type:'triangle',width:'sm',length:'sm'}}).bringToFront();}
function polar(){
  box('polar-water',1065,131,607,47,grad('#d6ebf8/70','#ffffff/0',90));
  poly('polar-back',[[1043,130],[1192,87],[1250,103],[1360,53],[1408,71],[1501,9],[1553,44],[1672,74],[1672,137]],'#d8ecf9');
  poly('polar-front',[[1172,137],[1298,97],[1365,119],[1460,44],[1501,9],[1542,54],[1594,45],[1672,94],[1672,142]],'#a9d3ee');
  for(const [i,p] of [[1,[[1501,9],[1457,113],[1483,83],[1495,127],[1532,61]]],[2,[[1501,9],[1542,54],[1594,45],[1540,115]]],[3,[[1360,53],[1303,131],[1341,104],[1373,128],[1408,71]]],[4,[[1594,45],[1581,136],[1615,105],[1651,136],[1672,94]]]])poly(`polar-snow-${i}`,p,'#ffffff/90');
}
function header(title,subtitle){polar();box('title-accent',52,30,14,65,grad('#f24e51','#cc2036'),3);txt('page-title',title,91,22,1396,64,43,NAVY,true);txt('page-subtitle',subtitle,91,91,1376,41,22,NAVY,true);box('header-rule',91,140,1518,1.2,'#b4cbe0');}
function footer(){curve('footer-light',0,868,W,73,[M(0,0),B(450,119,1030,73,1370,15),B(1490,0,1590,0,W,9),L(W,73),L(0,73),Z()],'#a9ddfa');curve('footer-blue',0,884,W,57,[M(0,3),B(530,75,1020,66,1407,13),B(1523,0,1590,0,W,3),L(W,57),L(0,57),Z()],grad('#1497ed','#7bc8fa'));}
function ribbon(name,title,kind,x,y,w,color=BLUE){
  curve(`${name}-ribbon`,x,y,w,60,[M(22,0),L(w,0),L(w-40,60),L(0,60),L(0,22),B(0,8,8,0,22,0),Z()],grad(color,'#074875'));
  pict(kind,x+25,y+11,44,'#ffffff');txt(`${name}-heading`,title,x+105,y+7,w-128,48,31,'#ffffff',true);
}
const sourceNote=t=>'来源：用户本轮按发送顺序提供的4张截图。'+t+' 仅转录并重建截图内容，未独立验证市场、品牌、产品或渠道主张。人物、图标和雪山为近似矢量重绘，不是原始照片；没有整页图片背景。';

// Slide 1 uses its original 1600×900 coordinate grid, scaled uniformly to deck width.
newSlide(sourceNote('第1张：线下健康主题活动与产品会。线下推广预算42–48万元按图保留。原图1600×900，按宽度等比适配1672×941，保留全部内容。'),1672/1600);
box('event-top-rule',90,104,1421,4,'#69a7cf');box('event-title-pill',390,55,820,101,grad('#35a5d5','#053e69'),51);
txt('page-title','线下健康主题活动与产品会',486,75,661,68,47,'#ffffff',true,{alignment:'center'});
txt('page-subtitle','以产品会聚集目标客户，通过讲解、体验与礼品互动促进现场转化。',380,185,982,45,28,'#515d6c');
for(const [i,x,w] of [[1,71,429],[2,536,529],[3,1101,429]]){box(`event-panel-${i}-shadow`,x,287,w,400,'#d9e2ec/45',29);box(`event-panel-${i}`,x,281,w,399,'#ffffff',29,'rect','#c7deef',1.3);}
for(const [i,x,label,color] of [[1,98,'目标人群','#2187c5'],[2,563,'产品会形式','#28a779'],[3,1128,'组织模式','#2187c5']]){box(`event-chip-${i}`,x,305,208,59,color,31);txt(`event-chip-${i}-label`,label,x+25,310,176,47,32,'#ffffff',true);}
for(const [i,y,title,desc] of [[1,410,'中老年人群','关注日常营养与健康管理'],[2,546,'亚健康关注人群','关注营养补充与生活方式']]){
  box(`event-audience-${i}`,105,y,360,92,'#f2f7fd',20,'rect','#d5e6f3',.7);pict('pair',108,y+10,100,'#2789c4');txt(`event-audience-${i}-title`,title,220,y+13,235,43,29,NAVY,true);txt(`event-audience-${i}-detail`,desc,220,y+55,238,30,19,'#627082');
}
const eventSteps=[{x:592,title:'健康分享',desc:'主题讲解',kind:'chat'},{x:752,title:'产品讲解',desc:'成分与特点',kind:'bottle'},{x:912,title:'体验礼遇',desc:'体验服务/礼品',kind:'gift'}];
for(const [i,d] of eventSteps.entries()){box(`event-step-${i+1}-circle`,d.x,391,96,96,'#edf7f2',0,'ellipse','#a8dcca',1.2);pict(d.kind,d.x+17,405,65,'#29ac7c');txt(`event-step-${i+1}-title`,d.title,d.x-9,511,116,46,28,NAVY,true,{alignment:'center'});txt(`event-step-${i+1}-detail`,d.desc,d.x-14,552,129,34,19,'#627082',false,{alignment:'center'});}
arrow('event-step-1-to-2',699,439,737,439);arrow('event-step-2-to-3',859,439,897,439);
box('event-conversion-strip',590,610,421,42,'#f0f8f4',22);txt('event-conversion-text','讲解 · 体验 · 转化',686,613,262,36,22,'#219568',true,{alignment:'center'});
for(const [i,y,title,desc,kind,color] of [[1,391,'自营活动','自建团队','user','#208ccc'],[2,477,'合作活动','外部资源','handshake','#2cac83'],[3,563,'混合模式','协同执行','network','#208ccc']]){
  box(`event-mode-${i}`,1135,y,360,66,'#f2f7fd',20,'rect','#d8e7f3',.7);pict(kind,1154,y+11,52,color);txt(`event-mode-${i}-title`,title,1225,y+6,221,35,28,NAVY,true);txt(`event-mode-${i}-detail`,desc,1225,y+40,221,25,19,'#627082');
}
box('event-budget-panel',100,730,1401,101,'#e8f3fd',52);box('event-budget-label-bg',100,730,371,101,'#0c396a',52);txt('event-budget-label','线下推广预算',155,752,277,54,32,'#ffffff',true);
txt('event-budget-value','42–48万元',533,747,366,68,50,NAVY,true);txt('event-budget-note','先试点验证，再复制放大',940,769,430,44,26,'#396386',true);
curve('event-footer-green',0,838,1600,62,[M(0,32),B(530,39,1024,1,1600,0),L(1600,62),L(0,62),Z()],'#2ba57d');curve('event-footer-navy',0,867,1600,33.4,[M(0,18),B(588,31,1170,-1,1600,0),L(1600,33.4),L(0,33.4),Z()],NAVY);

// Slide 2: online platforms, native bullet lists, conversion flow and consumer cards.
newSlide(sourceNote('第2张：线上市场。副标题个别字较模糊，按可辨内容整理为“线上渠道承担新品教育、需求承接与规模化销售，是品牌测试产品的重要阵地”，非保证逐字一致。其余项目、年龄段和说明按图转录。'));
header('二、线上市场：内容种草与货架转化协同','线上渠道承担新品教育、需求承接与规模化销售，是品牌测试产品的重要阵地');
const online=[
 {x:57,w:500,title:'综合电商平台',kind:'cart',color:'#0c88ca',fill:'#fff0f2',items:['搜索需求明确，承接主动购买流量','依靠品牌背书、评价与产品信息促成决策','大促节点适合家庭装与周期性囤货']},
 {x:594,w:496,title:'内容种草平台',kind:'video',color:'#ed4145',fill:'#e8f6ff',items:['通过短视频、图文与达人内容建立认知','围绕营养知识、成分特点和生活场景传播','小红书、抖音等适合新品种草与人群拓展']},
 {x:1126,w:504,title:'私域与社群',kind:'people',color:'#168c71',fill:'#e6f6ff',items:['以会员服务、社群内容持续触达用户','通过咨询与产品教育提升购买信任','适合周期管理、老客维护与复购转化']},
];
for(const [i,d] of online.entries()){
  box(`online-card-${i+1}-shadow`,d.x,179,d.w,242,'#dbe6f1/70',27);box(`online-card-${i+1}`,d.x,173,d.w,241,grad('#ffffff',d.fill,45),27);
  ribbon(`online-card-${i+1}`,d.title,d.kind,d.x,173,d.w-111,d.color);
  bullets(`online-card-${i+1}-bullet-body`,d.items,d.x+23,254,d.w-46,153,22.5,false,12);
}
box('online-flow-panel',57,459,1574,101,grad('#0b5a8a','#003b70'),22);
const flow=[{x:126,kind:'video',title:'内容触达',sub:'建立认知'},{x:531,kind:'search',title:'搜索比较',sub:'评估品牌与产品'},{x:938,kind:'cart',title:'下单转化',sub:'完成购买'},{x:1336,kind:'user',title:'会员复购',sub:'持续运营'}];
for(const [i,f] of flow.entries()){pict(f.kind,f.x,477,61,'#ffffff');txt(`online-flow-${i+1}-heading`,f.title,f.x+80,469,195,51,32,'#ffffff',true);txt(`online-flow-${i+1}-detail`,f.sub,f.x+98,513,190,36,22,'#ffffff');}
for(const [i,x] of [[1,426],[2,834],[3,1233]])arrow(`online-flow-arrow-${i}`,x,509,x+13,509,'#f0b8c5',4.5);
box('online-audience-accent',61,599,11,40,RED,2);txt('online-audience-heading','线上核心消费人群',98,591,830,56,35,NAVY,true);
const audiences=[{x:57,w:500,kind:'adult',color:'#1c87c3',title:'45岁以上 | 健康关注人群',desc:'关注心脑血管健康、营养补充与产品安全；重视口碑和长期使用体验。',fill:'#e9f5ff'},{x:594,w:496,kind:'office',color:'#2674aa',title:'25–40岁 | 精致白领与熬夜人群',desc:'关注日常营养、精力管理与成分信息；偏好清晰、可信的科普内容。',fill:'#edf7f6'},{x:1126,w:504,kind:'woman',color:'#e7464e',title:'30–45岁 | 家庭健康与送礼人群',desc:'为父母及家庭成员选购；重视品牌、品质、包装与购买便利性。',fill:'#fff0f3'}];
for(const [i,d] of audiences.entries()){box(`online-audience-${i+1}-shadow`,d.x,663,d.w,178,'#ccddeb/65',27);box(`online-audience-${i+1}`,d.x,658,d.w,177,grad('#ffffff',d.fill,50),27);portrait(d.kind,d.x+27,681,83,100,d.color);txt(`online-audience-${i+1}-title`,d.title,d.x+133,679,d.w-146,46,i===0?27:25.2,NAVY,true);txt(`online-audience-${i+1}-detail`,d.desc,d.x+133,736,d.w-151,85,20,NAVY,false,{wrap:'square',lineSpacing:1.2});}
footer();

// Slide 3: offline channels and operating thresholds.
newSlide(sourceNote('第3张：线下市场。四类渠道、三大门槛及渠道建议按截图保留；没有增加资质或准入承诺。'));
header('三、线下市场：专业信任与体验服务是关键','线下渠道在中老年客群触达、专业咨询与品牌信任建立方面具备独特价值');
const offline=[
 {x:58,y:171,w:506,h:241,title:'01 连锁药店与OTC渠道',kind:'store',color:'#1089d6',fill:'#fff0f4',items:['覆盖中老年健康消费人群','依托专业服务与店员推荐建立信任','需满足相应产品资质与宣传合规要求']},
 {x:598,y:171,w:490,h:241,title:'02 高端商超与会员店',kind:'cart',color:'#e83d40',fill:'#e6f7ff',items:['中高收入家庭客群集中','重视品牌、品质、配料与性价比','适合家庭装、组合装及礼赠场景']},
 {x:58,y:448,w:506,h:242,title:'03 健康管理与体检机构',kind:'medical',color:'#3a9b79',fill:'#e6f7ff',items:['连接体检后健康咨询场景','可围绕营养知识开展用户教育','更强调服务能力与证据支持']},
 {x:598,y:448,w:490,h:242,title:'04 母婴与精品零售',kind:'people',color:'#c99d49',fill:'#f6efe3',items:['覆盖孕产、女性及家庭营养需求','适合细分人群和品质型产品','需明确适用人群与合规表达边界']},
];
for(const [i,d] of offline.entries()){
 box(`offline-card-${i+1}-shadow`,d.x,d.y+6,d.w,d.h,'#d7e5ef/60',22);box(`offline-card-${i+1}`,d.x,d.y,d.w,d.h,grad('#ffffff',d.fill,45),22);
 ribbon(`offline-card-${i+1}`,d.title,d.kind,d.x,d.y,d.w-53,d.color);
 // Keep the full channel name on one editable heading line.
 const heading=slide.shapes.items.find(s=>s.name===`offline-card-${i+1}-heading`);
 if(heading)heading.text.style={typeface:FONT,fontSize:(i===0||i===2)?28:29,color:'#ffffff',bold:true,autoFit:'none',wrap:'none',insets:0,verticalAlignment:'middle'};
 bullets(`offline-card-${i+1}-bullet-body`,d.items,d.x+29,d.y+82,d.w-53,153,24,true,12);
}
box('offline-threshold-panel',1122,172,506,521,grad('#113f74','#043875'),49);
txt('offline-threshold-heading','线下渠道三大门槛',1164,192,399,62,37,'#ffffff',true);
poly('offline-threshold-red-1',[[1479,198],[1491,198],[1466,251],[1454,251]],'#e14e50');poly('offline-threshold-red-2',[[1492,198],[1500,198],[1474,251],[1467,251]],'#e14e50/60');
const thresholds=[{y:286,n:'01',title:'资质与合规',text:'普通食品与保健食品的资质、\n标签及宣传边界不同。',kind:'shield',color:'#e84440'},{y:415,n:'02',title:'价格与渠道',text:'需管理线上线下价差，\n避免渠道冲突与窜货。',kind:'layers',color:'#38a4ef'},{y:545,n:'03',title:'服务与动销',text:'依赖人员培训、产品知识和\n持续动销支持。',kind:'people',color:'#38a4ef'}];
for(const [i,d] of thresholds.entries()){
 box(`offline-threshold-${i+1}-number-bg`,1167,d.y+4,57,56,d.color,10);txt(`offline-threshold-${i+1}-number`,d.n,1177,d.y+11,40,40,28,'#ffffff',true,{alignment:'center'});
 pict(d.kind,1245,d.y+11,47,'#ffffff');box(`offline-threshold-${i+1}-divider`,1312,d.y+4,1.3,56,'#46a4da');
 txt(`offline-threshold-${i+1}-heading`,d.title,1337,d.y-10,264,49,29,'#ffffff',true);
 txt(`offline-threshold-${i+1}-detail`,d.text.replaceAll('\n',''),1337,d.y+34,279,65,20.5,'#ffffff',false,{wrap:'square',lineSpacing:1.22});
 if(i<2)box(`offline-threshold-${i+1}-rule`,1167,d.y+106,416,1,'#3281b1');
}
box('offline-advice-panel',58,738,1570,89,grad('#ffffff','#eaf5fd'),22,'rect','#e2edf5');
curve('offline-advice-label-bg',58,738,263,89,[M(22,0),L(263,0),L(231,89),L(22,89),B(7,89,0,83,0,67),L(0,22),B(0,8,8,0,22,0),Z()],grad('#176bc0','#074572'));
pict('target',91,755,54,'#ffffff');txt('offline-advice-label','渠道建议',158,750,146,61,33,'#ffffff',true);box('offline-advice-divider',337,765,1.4,36,BLUE);
txt('offline-advice','优先选择具备目标客群、专业服务能力和稳定动销机制的合作渠道；通过专供规格、组合装或服务权益实现渠道区隔。',359,748,1236,66,23.5,NAVY,true,{wrap:'square',lineSpacing:1.15});
footer();

// Slide 4: named competitors, product trends and channel strategies.
newSlide(sourceNote('第4张：竞争格局、产品趋势与全渠道策略。品牌和企业名单仅按截图转录，不构成品牌背书或独立调查。'));
header('四、竞争格局、产品趋势与全渠道策略','上游原料、品牌能力与渠道运营共同构成竞争壁垒，未来竞争将进一步转向品质与综合运营能力');
const panels=[{x:60,w:498,title:'竞争格局',kind:'trophy',color:'#2473af'},{x:583,w:511,title:'产品发展趋势',kind:'trend',color:'#2473af'},{x:1118,w:510,title:'全渠道运营策略',kind:'layers',color:'#e44243'}];
for(const [i,p] of panels.entries()){box(`strategy-panel-${i+1}-shadow`,p.x,166,p.w,597,'#d3e0ef/60',29);box(`strategy-panel-${i+1}`,p.x,160,p.w,597,'#ffffff',29);ribbon(`strategy-panel-${i+1}`,p.title,p.kind,p.x,160,p.w-45,p.color);}
box('competition-upstream-icon-bg',87,249,73,73,'#edf7ff',0,'ellipse');pict('factory',101,262,48,BLUE);txt('competition-upstream-heading','上游原料与技术',179,247,350,48,28,NAVY,true);
bullets('competition-upstream-bullet-body',['国际原料企业：Aker BioMarine等，\n具备原料与技术积累。','国内企业：逢时、辽渔等持续布局\n原料、加工与品牌。'],176,301,367,133,20.5,false,9);
box('competition-divider',82,443,448,1.2,'#dfebf6');
box('competition-brands-icon-bg',87,466,73,73,'#edf7ff',0,'ellipse');pict('tag',101,480,48,BLUE);txt('competition-brands-heading','中下游品牌',179,467,350,48,28,NAVY,true);
bullets('competition-brands-bullet-body',['国际品牌：Swisse、Schiff、\nBlackmores、GNC等。','专业与本土品牌：VIK、汤臣倍健、\n修正、同仁堂等。'],176,520,367,129,20.5,false,9);
box('competition-focus-bg',78,658,463,80,'#edf6ff',18);
txt('competition-focus','竞争重点：原料品质、品牌信任、\n产品差异化与渠道覆盖。',102,668,422,60,21.5,NAVY,true,{wrap:'square',lineSpacing:1.15});
const trends=[{y:246,h:100,n:'01',title:'高纯度化',detail:'磷脂含量与原料标准成为\n产品比较重点。',red:true},{y:357,h:106,n:'02',title:'复配功能化',detail:'与辅酶Q10、叶黄素等成分组合，\n拓展细分需求。',red:false},{y:485,h:105,n:'03',title:'剂型多元化',detail:'软胶囊、粉剂及便携型产品\n覆盖更多人群。',red:true},{y:610,h:105,n:'04',title:'宠物营养应用',detail:'向宠物皮毛、关节等营养场景延伸。',red:false}];
for(const [i,t] of trends.entries()){
 box(`product-trend-${i+1}-bg`,598,t.y,479,t.h,t.red?'#fff3f5':'#edf7ff',18);box(`product-trend-${i+1}-number-bg`,617,t.y+14,50,50,t.red?'#e13b3e':'#338dc6',10);txt(`product-trend-${i+1}-number`,t.n,624,t.y+22,37,34,24,'#ffffff',true,{alignment:'center'});
 box(`product-trend-${i+1}-divider`,692,t.y+14,2,t.h-30,t.red?RED:BLUE);txt(`product-trend-${i+1}-heading`,t.title,715,t.y+1,343,44,28,NAVY,true);
 txt(`product-trend-${i+1}-detail`,t.detail,715,t.y+43,347,t.h-45,21,NAVY,false,{wrap:'square',lineSpacing:1.15});
}
const channelStrategies=[{y:247,title:'线上：做认知与转化',detail:'通过内容科普、场景种草和\n搜索承接，建立明星单品。',kind:'laptop',red:false},{y:409,title:'线下：做信任与服务',detail:'布局药店、健康管理机构及\n优质零售渠道，强化专业体验。',kind:'store',red:true},{y:574,title:'私域：做留存与复购',detail:'以会员服务、知识内容和周期提醒\n提升用户长期价值。',kind:'people',red:false}];
for(const [i,c] of channelStrategies.entries()){
 box(`channel-strategy-${i+1}-bg`,1134,c.y,478,134,c.red?'#fff4f6':'#eef7ff',21);box(`channel-strategy-${i+1}-badge`,1150,c.y+24,75,75,c.red?'#ffe9ec':'#e2f1ff',0,'ellipse');pict(c.kind,1164,c.y+38,49,c.red?RED:BLUE);
 txt(`channel-strategy-${i+1}-heading`,c.title,1254,c.y+13,340,44,27,NAVY,true);txt(`channel-strategy-${i+1}-detail`,c.detail,1254,c.y+62,347,65,20.5,NAVY,false,{wrap:'square',lineSpacing:1.18});
}
box('strategy-conclusion-bg',60,784,1568,74,grad('#ffffff','#f2f9ff'),23,'rect','#e0ebf6');
curve('strategy-conclusion-label-bg',60,784,325,74,[M(24,0),L(325,0),L(269,74),L(24,74),B(6,74,0,67,0,51),L(0,24),B(0,9,8,0,24,0),Z()],grad('#1565ac','#064471'));
pict('target',97,795,55,'#ffffff');txt('strategy-conclusion-label','核心方向',165,791,155,57,34,'#ffffff',true);
txt('strategy-conclusion','以高品质产品为基础，形成“线上获客、线下背书、私域复购”的全渠道增长闭环。',412,790,1180,61,29,NAVY,true,{wrap:'square',lineSpacing:1.1});
footer();

const raw=path.join(build,'raw.pptx'),formatted=path.join(build,'native-bullets.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
// Restore standard DrawingML markers and source line breaks where the runtime
// exports buNone; keep paragraphs native and remove actual grouping locks.
const restoreBullets=String.raw`
import sys,json,copy
from zipfile import ZipFile
from lxml import etree as E
A='http://schemas.openxmlformats.org/drawingml/2006/main'
P='http://schemas.openxmlformats.org/presentationml/2006/main'
spec=json.loads(sys.argv[3]);count=0;removed=0
with ZipFile(sys.argv[1]) as src,ZipFile(sys.argv[2],'w') as dst:
 for info in src.infolist():
  data=src.read(info.filename)
  if info.filename.startswith('ppt/slides/slide') and info.filename.endswith('.xml'):
   root=E.fromstring(data)
   for e in root.iter():
    if e.get('noGrp')=='1':e.attrib.pop('noGrp');removed+=1
   for shape in root.findall('.//{%s}sp'%P):
    ident=shape.find('{%s}nvSpPr/{%s}cNvPr'%(P,P))
    key=ident.get('name','') if ident is not None else ''
    if key not in spec:continue
    paragraphs=shape.findall('.//{%s}p'%A)
    assert len(paragraphs)==len(spec[key]),(key,len(paragraphs),len(spec[key]))
    for paragraph,value in zip(paragraphs,spec[key]):
     p=paragraph.find('{%s}pPr'%A)
     assert p is not None
     for tag in ['buClr','buClrTx','buSzPct','buSzPts','buSzTx','buNone','buAutoNum','buBlip','buChar']:
      for e in p.findall('{%s}%s'%(A,tag)):p.remove(e)
     color=E.Element('{%s}buClr'%A);E.SubElement(color,'{%s}srgbClr'%A,val='E33E40')
     size=E.Element('{%s}buSzPct'%A,val='140000')
     marker=E.Element('{%s}buChar'%A,char='•')
     i=next((i for i,e in enumerate(p) if E.QName(e).localname in ['buFont','buFontTx','tabLst','defRPr','extLst']),len(p))
     p.insert(i,color);p.insert(i+1,size)
     i=next((i for i,e in enumerate(p) if E.QName(e).localname in ['tabLst','defRPr','extLst']),len(p));p.insert(i,marker)
     runs=paragraph.findall('{%s}r'%A);style=copy.deepcopy(runs[0].find('{%s}rPr'%A)) if runs else None
     for run in runs:paragraph.remove(run)
     for i,line in enumerate(value.split('\n')):
      if i:
       br=E.SubElement(paragraph,'{%s}br'%A)
       if style is not None:br.append(copy.deepcopy(style))
      r=E.SubElement(paragraph,'{%s}r'%A)
      if style is not None:r.append(copy.deepcopy(style))
      E.SubElement(r,'{%s}t'%A).text=line
     count+=1
   data=E.tostring(root,encoding='UTF-8',xml_declaration=True)
  dst.writestr(info,data)
assert count==25,count
print(json.dumps({'native_bullet_items':count,'removed_group_locks':removed}))
`;
const nativeFormat=execFileSync(process.env.RUNTIME_PYTHON,['-c',restoreBullets,raw,formatted,JSON.stringify(bulletSpecs)],{encoding:'utf8'});await fs.writeFile(path.join(build,'native-formatting.json'),nativeFormat);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',formatted,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-connectors','5','--min-target-arrows','5'],{encoding:'utf8'});await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,'market-channels-events-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-bullet-geometry','--validate-heading-fit'],explicitTotalSlideCount:4,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
for(const [i,s] of imported.slides.items.entries()){const png=await imported.export({slide:s,format:'png',scale:1});await fs.writeFile(path.join(build,`preview-${i+1}.png`),Buffer.from(await png.arrayBuffer()));}
console.log(JSON.stringify({finalPath,slides:4,slideOrder:'User attachment order',offlineBudget:'42–48万元',nativeBulletItems:25},null,2));
