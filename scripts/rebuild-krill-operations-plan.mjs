// Rebuild the five operating-plan screenshots as editable slides.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/krill-operations-plan',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation,applyPresentationChartFont}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
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
const paths={
  cart:'<path d="M5 5h10l9 39h32l9-28H19M24 53h31"/><circle cx="28" cy="60" r="4"/><circle cx="53" cy="60" r="4"/>',
  store:'<path d="M10 28v33h45V28M8 8h49l6 17c0 11-13 12-16 3-3 9-15 9-17 0-3 9-15 9-17 0-3 9-15 7-12-3zM28 44h15v17"/>',
  people:'<circle cx="34" cy="15" r="11"/><path d="M13 62V49c0-22 42-22 42 0v13zM12 8c-13 0-13 21 0 21M56 8c13 0 13 21 0 21M8 38c-9 0-11 9-11 20M60 38c9 0 11 9 11 20"/>',
  calendar:'<rect x="7" y="11" width="54" height="51" rx="4"/><path d="M19 4v14M49 4v14M7 26h54"/><path d="M18 37h4M33 37h4M48 37h4M18 48h4M33 48h4M48 48h4"/>',
  trend:'<path d="M5 34l17-15 13 9L61 3M47 3h14v15M6 62V47h11v15zM28 62V37h11v25zM50 62V25h11v37z"/>',
  target:'<circle cx="31" cy="36" r="27"/><circle cx="31" cy="36" r="17"/><circle cx="31" cy="36" r="7"/><path d="M31 36L60 7M53 1l-2 13 13-2"/>',
  coins:'<ellipse cx="45" cy="13" rx="17" ry="7"/><path d="M28 13v28c0 11 34 11 34 0V13M28 24c0 10 34 10 34 0M28 35c0 10 34 10 34 0"/><ellipse cx="21" cy="39" rx="17" ry="7"/><path d="M4 39v17c0 11 34 11 34 0V39M4 48c0 11 34 11 34 0"/>',
  box:'<path d="M6 18L33 4l28 14v32L33 65 6 50zM6 18l27 15 28-15M33 33v32M20 11l28 15"/>',
  megaphone:'<path d="M7 28l40-16v42L7 43zM19 46l6 18h10l-7-23M47 9l7-3v54l-7-6M62 17l6-5M62 32h8M62 47l6 5"/>',
  layers:'<path d="M5 21L34 5l29 16-29 16zM6 34l28 16 28-16M6 47l28 16 28-16"/>',
  network:'<rect x="27" y="2" width="14" height="13" rx="2"/><path d="M34 15v17M9 32h50M9 32v15M34 32v15M59 32v15"/><rect x="2" y="48" width="14" height="13" rx="2"/><rect x="27" y="48" width="14" height="13" rx="2"/><rect x="52" y="48" width="14" height="13" rx="2"/>',
  heart:'<path d="M34 62C18 48 4 35 4 20 4 1 25 0 34 15 43 0 64 1 64 20c0 15-14 28-30 42zM6 32h14l6-11 10 23 7-12h18"/>',
  video:'<rect x="4" y="7" width="60" height="45" rx="4"/><path d="M28 20l17 10-17 11zM16 62h37"/>',
  chat:'<path d="M25 4C-2 4-3 34 12 43L7 55l16-7c20 3 35-11 32-26C53 11 40 4 25 4zM44 29C23 29 22 55 44 58l16 7-3-14c11-15 3-22-13-22z"/><circle cx="18" cy="23" r="2"/><circle cx="35" cy="23" r="2"/><circle cx="40" cy="43" r="2"/><circle cx="52" cy="43" r="2"/>',
  sprout:'<path d="M33 64V29M33 41C9 43 3 25 5 11c23 0 31 14 28 30zM34 32C34 8 54 3 66 5c-1 22-17 32-32 27zM16 23l17 18M34 32l20-15"/>',
  repeat:'<path d="M9 25C13 3 48 0 60 20M60 7v14H46M60 42C55 64 20 68 8 48M8 61V47h14"/>',
};
function pict(kind,x,y,size,color=NAVY) {
  assert(paths[kind],kind);
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="84" height="84" viewBox="-7 -5 84 84"><g fill="none" stroke="${color}" stroke-width="4.2" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} icon`,position:{left:x,top:y,width:size,height:size}});
}
const TITLE='南极磷虾油线上线下一体化运营方案';
function newSlide(note) {slide=deck.slides.add();slide.background.fill='#ffffff';slide.speakerNotes.text=note;}
function mountains(x,y,w,h) {
  box('polar-scene',x,y,w,h,grad('#f5fbff','#c3e6f5',90));
  const pts=a=>a.map(([px,py])=>[x+px*w,y+py*h]);
  poly('mountain-back',pts([[0,.8],[.12,.43],[.21,.61],[.34,.19],[.46,.58],[.64,.05],[.78,.5],[.87,.32],[1,.61],[1,1],[0,1]]),'#b6d7eb');
  poly('mountain-front',pts([[.03,.84],[.28,.51],[.37,.72],[.64,.05],[.75,.59],[.84,.47],[1,.81],[1,1],[0,1]]),'#75afd1');
  poly('snow-1',pts([[.64,.05],[.43,.82],[.59,.53],[.63,.87],[.68,.37],[.75,.59],[.84,.47],[.73,.89]]),'#f9fdff');
  poly('snow-2',pts([[.34,.19],[.18,.81],[.29,.62],[.36,.87],[.46,.58]]),'#effaff');
  box('polar-water',x,y+h*.86,w,h*.14,grad('#84c0dd','#e7f7ff',90));
  for(let i=0;i<4;i++)box(`polar-ripple-${i}`,x+w*.06,y+h*(.88+i*.025),w*(.88-i*.06),1,'#ffffff/70');
}
function footer(orange=false) {
  if(orange)curve('footer-orange',0,821,W,120,[M(0,90),B(400,110,1090,110,W,0),L(W,120),L(0,120),Z()],'#f99100');
  curve('footer-green',0,858,W,83,[M(0,52),B(530,-14,780,111,W,0),L(W,83),L(0,83),Z()],grad('#72ba91','#15916a'));
  curve('footer-navy',0,869,W,72,[M(0,12),B(290,-49,535,114,1019,47),B(1330,40,1466,15,W,27),L(W,72),L(0,72),Z()],grad('#164878','#032b5a'));
}
function darkHeader(label='',wide=false) {
  box('top-header',0,0,W,wide?90:74,grad('#07385f','#002c61'));
  poly('header-green-accent',[[50,0],[74,0],[42,74],[18,74]],grad('#41b77d','#137e64'));
  txt('project-header',TITLE,80,6,wide?630:900,wide?75:60,wide?38:40,'#ffffff',true);
  if(label){box('header-divider',wide?729:1052,18,1.3,44,'#96b0c8');txt('header-label',label,wide?760:1074,13,wide?450:423,54,wide?28:29,'#ffffff');}
  mountains(wide?1260:1455,0,wide?412:217,wide?100:74);
}
function lineArrow(name,x1,y1,x2,y2,color='#83adcf',width=7) {
  const a=box(`${name}-from`,x1,y1,1,1),b=box(`${name}-to`,x2,y2,1,1);
  slide.shapes.connect(a,b,{kind:'straight',line:{fill:color,width},head:{type:'none'},tail:{type:'triangle',width:'sm',length:'sm'}}).bringToFront();
}
function note(text) {return '来源：用户本轮按发送顺序提供的5张运营方案截图。'+text+' 所有数据、规划、营养表述按图转录，未独立验证商业可行性或产品属性。原始图片文件未在云端提供；雪山、瓶装产品、场景及平台标识为近似矢量重绘，并非原始照片或官方Logo文件。';}
function capsule(name,x,y,w,h,rotation=0) {
  const c=box(name,x,y,w,h,grad('#ffb381','#d71500',90),0,'ellipse','#ae1304',2);c.position={left:x,top:y,width:w,height:h,rotation};
  const shine=box(`${name}-shine`,x+8,y+3,w-20,h*.22,'#fff1db/85',0,'ellipse');shine.position={left:x+8,top:y+3,width:w-20,height:h*.22,rotation};
}

// Slide 1: overall plan and existing resources.
newSlide(note('第1张：项目整体规划；首期资金140–150万元，采购70万、推广70–80万。'));
mountains(1093,181,579,496);
box('project-top-accent',58,34,10,40,GREEN);
txt('project-header',TITLE,80,25,616,56,38,NAVY,true);
box('header-divider',721,37,1.3,36,'#9baec2');
box('section-chip',753,29,264,49,'#eef7f1',8);txt('section-label','项目整体规划',769,27,241,51,38,'#207d51',true);
box('header-rule',1018,64,590,1,'#7a95a5');
poly('header-mark',[[1569,51],[1620,51],[1612,72],[1556,72]],grad(NAVY,GREEN));
curve('title-mint-underline',58,163,574,20,[M(0,0),L(574,0),L(567,20),L(0,20),Z()],'#e4f3eb');
txt('page-title','线上线下双轮驱动，打造自主品牌销售体系',58,98,1540,93,62,NAVY,true);
txt('page-subtitle','依托现有产品供应链、经销商及私域资源，构建线上电商与线下实体渠道协同发展的经营模式。',62,198,1478,49,28.5,'#242424');
box('capital-panel-shadow',65,275,1058,400,'#dbe6ef/45',25);box('capital-panel',64,268,1054,397,'#ffffff',25);
box('capital-hero',79,282,1025,181,grad('#5a8ca4','#103b68',15),30);
curve('capital-green-wave',850,388,254,75,[M(0,75),B(87,23,160,-6,254,0),L(254,40),B(167,18,102,63,0,75),Z()],grad('#4bba7d','#21815f'));
txt('capital-total','140–150万元',331,280,692,111,94,'#ffffff',true,{italic:true});
box('capital-line-left',327,417,103,1.3,'#ffffff');box('capital-line-right',761,417,108,1.3,'#ffffff');
txt('capital-label','项目首期资金规划',448,390,347,56,37,'#ffffff',true);
for(const [i,x,w,color,val,label,kind] of [[1,94,479,GREEN,'70万','产品采购','cart'],[2,594,494,BLUE,'70–80万','市场推广','megaphone']]) {
  box(`capital-allocation-${i}`,x,479,w,162,i===1?grad('#f0f8f3','#e3f2eb'):grad('#eff6ff','#e7f0fb'),22);
  box(`capital-badge-${i}`,x+31,503,113,112,i===1?'#dff0e8':'#dbe9f9',0,'ellipse');pict(kind,x+51,523,74,color);
  box(`capital-divider-${i}`,x+169,505,1.3,108,'#c7ddda');
  txt(`capital-value-${i}`,val,x+194,486,w-205,83,i===1?65:62,color,true);
  txt(`capital-purpose-${i}`,label,x+194,565,w-205,57,37,'#202020',true);
}
// Product bottle is a native schematic reconstruction.
box('product-ground',1180,610,430,65,'#e7f5fb',0,'ellipse');
box('bottle-neck',1320,292,171,74,grad('#113c73','#061e4f',0),24);
box('bottle-body',1286,329,232,298,grad('#244f79','#001c4b',0),46,'rect','#355b82',2);
box('bottle-label',1294,369,216,224,grad('#264e77','#021c43',0),12);
box('bottle-cap',1307,248,188,66,grad('#d1d7dc','#fafcff',0),13,'rect','#a9b7c1',2);
box('bottle-cap-bottom',1307,298,188,17,grad('#6e8798','#e4edf3'),5);
for(let i=0;i<12;i++)box(`bottle-cap-ridge-${i}`,1316+i*15,254,1.5,42,'#9eafbd/50');
txt('bottle-product-name','南极磷虾油',1315,391,187,48,32,'#ffffff',true,{alignment:'center'});
txt('bottle-english','ANTARCTIC\nKRILL OIL',1327,443,163,56,19,'#ffffff',false,{alignment:'center',wrap:'square'});
mountains(1295,511,215,72);
curve('bottle-green-swoosh',1295,533,215,53,[M(0,14),B(91,60,148,-8,215,0),L(215,15),B(145,26,91,63,0,14),Z()],'#2c997a');
capsule('bottle-capsule-1',1208,596,78,31,-8);capsule('bottle-capsule-2',1253,609,86,34,-12);capsule('bottle-capsule-3',1321,590,52,26,14);
txt('bottle-schematic-caption','产品示意重绘',1338,650,180,27,15,'#62828e',false,{alignment:'right'});
curve('resource-label-bg',58,707,215,135,[M(22,0),L(215,0),L(174,135),L(22,135),B(5,135,0,127,0,113),L(0,22),B(0,6,8,0,22,0),Z()],grad('#123f6e','#07345f'));
txt('resource-heading','现有资源\n基础',89,725,150,104,33,'#ffffff',true,{wrap:'square'});
const resources=[{x:282,w:413,title:'产品：',detail:'56%磷脂南极磷虾油',kind:'box'},{x:718,w:426,title:'私域：',detail:'200多个社群，10万+\n潜在触达人群',kind:'people'},{x:1167,w:447,title:'渠道：',detail:'北京、安徽已有经销商\n及实体店资源',kind:'store'}];
for(const [i,r] of resources.entries()){
  box(`resource-${i+1}`,r.x,707,r.w,135,grad('#f5faf7','#edf5f3'),22);
  box(`resource-${i+1}-badge`,r.x+17,730,86,86,'#e5f2eb',0,'ellipse');pict(r.kind,r.x+34,747,53,'#3a9f74');
  box(`resource-${i+1}-divider`,r.x+116,726,1.2,100,'#d0e4dc');
  txt(`resource-${i+1}-heading`,r.title,r.x+133,716,r.w-145,47,32,'#227c50',true);
  txt(`resource-${i+1}-detail`,r.detail,r.x+133,764,r.w-144,65,i===0?26:27,'#202020',false,{wrap:'square',lineSpacing:1.2});
}
footer();

// Slide 2: three phases and four operating measures.
newSlide(note('第2张：原图页眉“项目实施与经营目标（第5页）”按图保留；阶段1–3、4–6、7–12个月，没有新增指标目标数。'));
darkHeader('项目实施与经营目标（第5页）',true);
box('title-accent',41,124,10,49,GREEN,3);txt('page-title','分阶段推进，建立持续经营能力',72,109,1430,78,55,NAVY,true);
const phases=[{x:29,w:498,n:'01',title:'市场启动',period:'1–3个月',detail:'平台建设、渠道试点、推广测试',kind:'calendar',color:NAVY},{x:597,w:488,n:'02',title:'渠道增长',period:'4–6个月',detail:'扩大投放、拓展经销商、提升销量',kind:'trend',color:GREEN},{x:1161,w:489,n:'03',title:'规模运营',period:'7–12个月',detail:'区域扩张、客户复购、优化经营',kind:'target',color:NAVY}];
for(const [i,p] of phases.entries()) {
  box(`phase-${i+1}-frame`,p.x,246,p.w,197,'#ffffff',30,'rect','#b9cfe3');
  box(`phase-${i+1}-body`,p.x+10,282,p.w-20,134,grad('#ffffff','#edf3f6'),25);
  curve(`phase-${i+1}-header`,p.x+7,221,p.w-7,69,[M(30,0),L(p.w-34,0),L(p.w-7,34),L(p.w-34,69),L(30,69),B(-8,69,-8,0,30,0),Z()],grad(i===1?'#38b370':'#2b6984',p.color,0));
  box(`phase-${i+1}-num-badge`,p.x+44,227,55,55,'#ffffff',0,'ellipse');txt(`phase-${i+1}-number`,p.n,p.x+49,232,45,45,31,p.color,true,{alignment:'center'});
  txt(`phase-${i+1}-heading`,p.title,p.x+132,229,p.w-156,54,37,'#ffffff',true);
  pict(p.kind,p.x+25,325,68,p.color);box(`phase-${i+1}-divider`,p.x+108,320,1.2,73,'#94afc0');
  txt(`phase-${i+1}-period`,p.period,p.x+130,308,p.w-144,58,39,'#16874b',true);
  txt(`phase-${i+1}-detail`,p.detail,p.x+130,365,p.w-144,41,i===1?23:24.5,'#202020');
}
lineArrow('phase-1-to-2',545,338,576,338,'#9fb9d2',9);lineArrow('phase-2-to-3',1105,338,1137,338,'#9fb9d2',9);
box('metrics-rule-left',32,517,620,1.5,'#9ab5ca');box('metrics-rule-right',1020,517,619,1.5,'#9ab5ca');
txt('metrics-heading','四大核心经营指标',676,489,340,60,39,NAVY,true);
const metrics=[{x:26,w:388,title:'销售规模',detail:'销售额、销售盒数',kind:'trend',color:NAVY},{x:437,w:387,title:'推广效率',detail:'投放ROI、获客成本',kind:'coins',color:GREEN},{x:848,w:391,title:'渠道建设',detail:'经销商数量、终端动销',kind:'store',color:NAVY},{x:1263,w:387,title:'客户经营',detail:'复购率、私域转化率',kind:'people',color:GREEN}];
for(const [i,m] of metrics.entries()){
  box(`metric-${i+1}`,m.x,575,m.w,149,grad('#ffffff','#e8f4ee',90),24,'rect','#6faf9e',.8);pict(m.kind,m.x+30,619,68,m.color);
  box(`metric-${i+1}-divider`,m.x+122,612,1.2,74,'#8cadb4');txt(`metric-${i+1}-heading`,m.title,m.x+145,599,m.w-153,57,38,NAVY,true);
  txt(`metric-${i+1}-detail`,m.detail,m.x+145,655,m.w-153,43,i===2?23:24,'#252525');
}
curve('final-goal-bg',26,785,1624,110,[M(32,0),L(1624,0),L(1593,110),L(0,110),Z()],grad('#003466','#1aa05c',0));
pict('target',96,808,71,'#ffffff');box('final-goal-divider',202,811,1.3,60,'#b3cbd4');
txt('final-goal','最终目标：建立自主品牌、完善销售渠道、扩大市场覆盖，实现线上线下协同增长。',237,807,1350,67,36,'#ffffff',true);

// Slide 3: native doughnut with an embedded data workbook.
newSlide(note('第3张：原图“第4页：资金投入与推广计划”按图保留。推广预算70–80万元；比例为约60%、约33%、约7%。使用新建60/33/7数据快照，不声称保留原始工作簿或公式。'));
darkHeader('第4页：资金投入与推广计划');
poly('title-navy-slash',[[85,108],[103,108],[81,153],[64,153]],NAVY);poly('title-green-slash',[[109,108],[125,108],[103,153],[87,153]],'#3fa773');
txt('page-title','合理配置推广预算，提升市场投入效率',132,94,1040,77,51,NAVY,true);box('title-rule',994,140,613,1.2,'#abc1d0');
box('budget-line',96,234,1464,5,grad('#eaf1f8','#176daf',0));box('budget-total-bg',479,191,713,91,grad('#3f9cd6','#00365f',20),48);
txt('budget-total','70–80万元推广预算',569,194,594,81,61,'#ffffff',true);
const budgetChart=slide.charts.add('doughnut',{position:{left:130,top:301,width:444,height:444},categories:['线上推广','线下推广','机动储备'],series:[{name:'推广预算占比',values:[60,33,7],points:[{idx:0,fill:grad('#2183c9','#004891',90)},{idx:1,fill:grad('#51bf7f','#168b5c',90)},{idx:2,fill:grad('#70c3fb','#259ce8',90)}]}],doughnutOptions:{holeSize:59,firstSliceAngle:0},hasLegend:false,dataLabels:{showPercent:true,showValue:false,position:'inEnd',numberFormatCode:'"约"0%',textStyle:{fontSize:24,fill:'#ffffff',bold:true}},chartFill:'none',chartLine:{fill:'none',width:0},plotAreaFill:'none',plotAreaLine:{fill:'none',width:0}});
applyPresentationChartFont(budgetChart,{fontFamily:FONT});
pict('coins',313,444,79,NAVY);txt('budget-chart-title','推广预算分配',262,515,187,44,29,NAVY,true,{alignment:'center'});txt('budget-chart-total','（70–80万元）',263,560,185,38,26,'#555555',false,{alignment:'center'});
lineArrow('budget-to-channels',609,524,642,524,'#bbd8ee',11);
const allocations=[{y:305,title:'线上推广',pct:'约60%',detail:'电商广告、达人合作、内容营销、私域运营。',kind:'cart',color:BLUE,fill:'#eaf1fc'},{y:458,title:'线下推广',pct:'约33%',detail:'经销商开发、门店宣传、市场活动、渠道培训。',kind:'store',color:'#2a9e71',fill:'#eaf6f0'},{y:613,title:'机动储备',pct:'约7%',detail:'优质渠道追加及新市场测试。',kind:'layers',color:'#42a4d4',fill:'#edf3fc'}];
for(const [i,a] of allocations.entries()) {
  box(`allocation-${i+1}-surface`,746,a.y,839,131,grad('#ffffff',a.fill),17);
  box(`allocation-${i+1}-badge`,694,a.y+20,110,88,grad(i===1?'#46bf78':'#38a0e4',i===1?'#147b5a':'#004e83'),14);pict(a.kind,721,a.y+40,58,'#ffffff');
  box(`allocation-${i+1}-divider`,845,a.y+22,1.2,84,'#8eabc3');
  txt(`allocation-${i+1}-heading`,a.title,888,a.y+7,187,56,42,i===1?'#2b9569':BLUE,true);
  box(`allocation-${i+1}-pct-bg`,1076,a.y+17,133,46,a.color,24);txt(`allocation-${i+1}-pct`,a.pct,1088,a.y+16,110,46,35,'#ffffff',true,{alignment:'center'});
  txt(`allocation-${i+1}-detail`,a.detail,888,a.y+70,684,51,29,'#555555');
}
box('budget-principle-bg',64,788,1544,81,grad('#ffffff','#e7effa'),42);box('budget-principle-label-bg',64,788,346,81,NAVY,40);box('budget-principle-label-square',370,788,40,81,NAVY);
pict('target',99,803,56,'#ffffff');txt('budget-principle-label','资金使用原则：',168,800,235,58,34,'#ffffff',true);
box('budget-principle-divider',431,807,1.2,44,'#c4d6e3');
txt('budget-principle','前期测试、中期放量、后期优化；根据实际成交、获客成本及渠道利润动态调整预算。',451,801,1130,56,29,NAVY,true);
footer();

// Slide 4: offline channels and schematic replacements for unavailable photos.
newSlide(note('第4张：线下运营方案。四个原照片区域以独立可编辑的地图网络、门店陈列、机构接待、团购场景示意重建，明确标记场景示意。首批试点北京、安徽按图保留。'));
darkHeader();
curve('offline-section-bg',44,102,383,61,[M(0,0),L(348,0),L(383,30),L(348,61),L(0,61),Z()],grad('#6ac6a6','#008451',90));txt('offline-section','03  线下运营方案',68,103,337,56,41,'#ffffff',true);
poly('offline-title-slash',[[54,187],[66,187],[54,240],[43,240]],'#48b590');txt('page-title','以区域经销为基础，拓展实体销售网络',81,171,1510,83,59,NAVY,true);
const offline=[{x:25,w:397,title:'区域经销商',detail:'招募代理，建立\n区域销售网络',kind:'network',color:NAVY},{x:436,w:395,title:'实体门店',detail:'健康产品店及符合\n准入要求的药房',kind:'store',color:'#079367'},{x:845,w:394,title:'健康机构',detail:'健康管理中心及\n相关合作机构',kind:'heart',color:NAVY},{x:1253,w:396,title:'企业团购',detail:'企业福利、会员采购、\n商务团购',kind:'people',color:'#08986a'}];
for(const [i,o] of offline.entries()){
  box(`offline-card-${i+1}`,o.x,278,o.w,412,grad('#ffffff','#f2f7fa'),24,'rect','#d8e3e9');
  box(`offline-card-${i+1}-badge`,o.x+24,313,91,82,o.color,0,'ellipse');pict(o.kind,o.x+45,331,51,'#ffffff');
  box(`offline-card-${i+1}-divider`,o.x+140,311,1.3,86,'#9eb4be');
  txt(`offline-card-${i+1}-heading`,o.title,o.x+167,289,o.w-180,57,37,NAVY,true);
  txt(`offline-card-${i+1}-detail`,o.detail,o.x+167,340,o.w-176,73,i===3?22:26,'#111111',false,{wrap:'square',lineSpacing:1.25});
  box(`offline-scene-${i+1}`,o.x+1,438,o.w-2,251,i===0?grad('#28677b','#082c4c',60):i===1?grad('#d8c6a7','#faf3e5'):i===2?grad('#eee9db','#ffffff'):grad('#d0e9f0','#f4fbfd'),0);
}
// Distribution network: native connection lines link editable location nodes.
poly('distribution-map',[[51,626],[70,562],[122,543],[155,490],[210,507],[240,481],[293,527],[376,548],[391,600],[348,632],[290,646],[249,670],[186,648],[121,667]],'#25738b');
const center=box('map-network-center',207,577,17,17,'#c7faff',0,'ellipse');
for(const [i,x,y] of [[1,103,545],[2,155,508],[3,282,522],[4,343,567],[5,293,627]]){
  const n=box(`map-location-${i}`,x,y,14,14,'#bdf7ff',0,'ellipse');slide.shapes.connect(center,n,{kind:'curved',line:{fill:'#6cdeef',width:2},head:{type:'none'},tail:{type:'none'}}).bringToFront();
  box(`map-location-core-${i}`,x+4,y+4,6,6,'#ffffff',0,'ellipse');
}
// Store shelving with editable generic product cartons.
box('store-sign',466,455,271,48,'#204e5d',7);txt('store-sign-text','营养健康产品',490,460,226,38,26,'#ffffff',true,{alignment:'center'});
for(let r=0;r<3;r++){
  box(`store-shelf-${r}`,460,551+r*49,303,8,'#b58b58');
  for(let c=0;c<7;c++){
    box(`store-carton-${r}-${c}`,470+c*40,512+r*49,27,39,grad('#ffffff','#bfdfee'),3,'rect','#7ba5ba');box(`store-carton-band-${r}-${c}`,470+c*40,535+r*49,27,9,'#2c7daf');
  }
}
box('store-screen',754,500,61,164,'#17618d',7,'rect','#effbff',3);mountains(760,543,49,84);txt('store-screen-product','磷虾油',759,510,51,29,15,'#ffffff',true,{alignment:'center'});
// Health-service reception schematic.
box('clinic-wall-panel',875,465,269,111,'#edf2eb',7);pict('heart',891,483,54,'#53a889');txt('clinic-wall-name','健康管理中心',950,487,191,43,25,'#245947',true);
box('clinic-desk-side',885,578,282,76,'#c3c9c5');poly('clinic-desk-front',[[878,587],[1131,566],[1168,584],[1168,656],[878,676]],grad('#ffffff','#e0e7e2'));
box('clinic-screen',1038,552,56,28,NAVY,3);box('clinic-screen-leg',1061,578,9,12,'#5d7471');
box('clinic-pot',1180,597,28,66,'#d4d0bd',6);for(let i=0;i<4;i++){const leaf=box(`clinic-leaf-${i}`,1167+i*10,549+(i%2)*10,16,57,'#458465',0,'ellipse');leaf.position={left:1167+i*10,top:549+(i%2)*10,width:16,height:57,rotation:-35+i*22};}
// Corporate purchasing scene: generic figures and product bag.
box('corporate-window',1262,447,377,143,grad('#edfaff','#b4d9e7'),0);for(let i=0;i<3;i++)box(`corporate-window-divider-${i}`,1330+i*95,447,4,143,'#f8feff');
for(const [i,x,y] of [[1,1386,467],[2,1489,478],[3,1576,480]]){
  box(`corporate-person-${i}-head`,x,y,31,34,'#627e8c',0,'ellipse');box(`corporate-person-${i}-body`,x-14,y+35,60,96,grad('#264d63','#082d4a'),20);
}
box('corporate-table',1267,599,369,14,'#87b2c1',5);box('corporate-bag',1301,563,126,99,grad('#205685','#052d59'),4);curve('corporate-bag-handle',1329,536,69,41,[M(0,40),B(0,-6,69,-6,69,40)],'none','#c5a16b',4);txt('corporate-bag-label','南极磷虾油',1307,594,115,37,19,'#ffffff',true,{alignment:'center'});mountains(1308,633,113,23);
for(const [i,o] of offline.entries())txt(`offline-scene-caption-${i+1}`,'场景示意',o.x+o.w-103,659,89,24,15,i===0?'#d7f6ff':'#4f6c78',false,{alignment:'right'});
box('offline-strategy-bg',334,722,1294,120,grad('#ffffff','#e7f4ee'),24);
curve('offline-strategy-label-bg',45,735,329,92,[M(0,0),L(289,0),L(329,46),L(289,92),L(0,92),Z()],grad('#0c6197','#002f63'));
pict('target',82,754,51,'#ffffff');txt('offline-strategy-label','实施策略',153,748,199,63,40,'#ffffff',true);box('offline-strategy-divider',398,748,1.3,69,'#9dbbb7');
const strategy=txt('offline-strategy','以北京、安徽为首批试点区域，提供产品培训、宣传物料及市场推广支持，\n成熟后逐步拓展全国市场。',424,735,1170,93,31,'#111111',false,{wrap:'square',lineSpacing:1.25});
strategy.text.get('北京、安徽').bold=true;strategy.text.get('北京、安徽').color='#048e60';
footer();

// Slide 5: ecommerce, social content and private-domain sales.
newSlide(note('第5张：线上运营方案。天猫、京东、拼多多、抖音、小红书、视频号、微信标识为近似矢量示意；所有平台名称与经营说明保留为可编辑文字。底部流程以可编辑节点和4条原生关系箭头重建。'));
poly('online-header-navy-slash',[[53,26],[73,26],[57,65],[35,65]],NAVY);poly('online-header-green-slash',[[78,26],[93,26],[76,65],[61,65]],'#059777');
txt('project-header',TITLE,109,14,1135,65,42,NAVY,true);box('online-header-rule',35,79,1597,1.2,'#9dbacb');
curve('online-section-bg',1336,26,325,55,[M(24,0),L(325,0),L(304,55),L(0,55),Z()],NAVY);txt('online-section','2 / 线上运营方案',1391,30,251,47,30,'#ffffff',true);
poly('online-title-slash',[[65,125],[83,125],[63,180],[45,180]],'#049a7d');txt('page-title','多平台布局，建立品牌曝光与销售转化体系，',103,105,1507,89,59,NAVY,true);
const online=[{x:45,w:533,title:'天猫、京东、拼多多',kind:'cart',color:NAVY,detail:'店铺运营、搜索推广、广告投放'},{x:608,w:534,title:'抖音、小红书、视频号',kind:'video',color:'#008d75',detail:'短视频、产品科普、达人合作'},{x:1172,w:469,title:'微信私域',kind:'people',color:NAVY,detail:'社群营销、客户转化、会员复购'}];
for(const [i,o] of online.entries()){
  box(`online-card-${i+1}`,o.x,235,o.w,358,'#ffffff',23,'rect','#c5d4dd',1.2);box(`online-header-${i+1}`,o.x,235,o.w,87,grad(i===1?'#009f8d':'#005a9a',o.color),23);box(`online-header-lower-${i+1}`,o.x,277,o.w,45,o.color);
  pict(o.kind,o.x+45,251,61,'#ffffff');box(`online-divider-${i+1}`,o.x+125,252,1.3,55,'#b3d3db');txt(`online-heading-${i+1}`,o.title,o.x+145,246,o.w-158,68,i===2?41:38,'#ffffff',true);
  box(`online-description-bg-${i+1}`,o.x+18,511,o.w-36,63,'#e9f0f7',17);txt(`online-description-${i+1}`,o.detail,o.x+25,515,o.w-50,52,i===2?29:30,NAVY,true,{alignment:'center'});
}
// Approximate platform marks; names remain editable even where artwork is simplified.
for(const [i,x,label,fill] of [[1,98,'天猫','#ee0021'],[2,255,'京东','#ed1509'],[3,412,'拼多多','#f40b25'],[4,660,'抖音','#111016'],[5,819,'小红书','#ff2141'],[6,977,'视频号','#ffffff']]){
  box(`platform-${i}-tile`,x,359,117,117,fill,24,'rect',i===6?'#e8e8e8':'none');
  if(i===1){txt('tmall-word','天猫',x+9,370,100,54,38,'#ffffff',true,{alignment:'center'});poly('tmall-cat-ears',[[x+4,451],[x+8,432],[x+29,443],[x+83,443],[x+108,431],[x+113,452],[x+113,475],[x+4,475]],'#080a0c');for(const ex of [x+30,x+83]){box(`tmall-eye-${ex}`,ex-7,449,14,24,'#ffffff',0,'ellipse');box(`tmall-pupil-${ex}`,ex-2,449,4,24,'#101015',0,'ellipse');}}
  if(i===2){box('jd-dog-head',x+28,369,66,39,'#ffffff',0,'ellipse');box('jd-dog-ear',x+19,375,17,29,'#dddddd',0,'ellipse');box('jd-dog-nose',x+80,375,10,8,'#111111',0,'ellipse');curve('jd-dog-smile',x+49,389,35,10,[M(0,0),B(5,11,28,11,35,0)],'none','#111111',2);txt('jd-word',label,x+12,426,94,42,29,'#ffffff',true,{alignment:'center'});}
  if(i===3){for(let k=0;k<7;k++)box(`pdd-heart-tile-${k}`,x+29+(k%3)*21,374+Math.floor(k/3)*20,17,17,'#ffffff',3);txt('pdd-word',label,x+7,433,103,33,23,'#ffffff',true,{alignment:'center'});}
  if(i===4){
    for(const [k,dx,dy,c] of [['cyan',-3,2,'#00e4e8'],['red',3,2,'#ff1657'],['white',0,0,'#ffffff']]){
      box(`douyin-${k}-stem`,x+58+dx,373+dy,8,48,c,1);
      box(`douyin-${k}-note`,x+36+dx,409+dy,30,20,c,0,'ellipse');
      curve(`douyin-${k}-hook`,x+62+dx,373+dy,24,24,[M(0,0),B(2,13,12,20,24,20)],'none',c,8);
    }
    txt('douyin-word',label,x+11,435,95,33,25,'#ffffff',true,{alignment:'center'});
  }
  if(i===5)txt('redbook-word',label,x+8,389,102,63,31,'#ffffff',true,{alignment:'center'});
  if(i===6){curve('channels-orange-mark',x+30,375,59,48,[M(0,45),B(-15,-31,20,17,30,41),B(66,-25,75,-6,45,45),B(39,49,25,7,0,45)],'none','#f99c1c',5);txt('channels-word',label,x+8,430,102,35,27,'#111111',true,{alignment:'center'});}
}
box('wechat-tile',1343,358,125,123,grad('#20c24d','#00a52c'),24);pict('chat',1357,375,97,'#ffffff');
const flows=[{x:42,w:289,label:'内容传播',kind:'video',color:BLUE},{x:373,w:287,label:'用户种草',kind:'sprout',color:'#008e77'},{x:701,w:287,label:'电商成交',kind:'cart',color:'#ef8b00'},{x:1026,w:286,label:'私域沉淀',kind:'people',color:'#008d75'},{x:1353,w:286,label:'复购转化',kind:'repeat',color:NAVY}];
for(const [i,f] of flows.entries()){
  box(`flow-${i+1}-body`,f.x+44,684,f.w-44,90,grad(f.color,i===2?'#ffad00':f.color),8);box(`flow-${i+1}-badge`,f.x,676,106,106,'#ffffff',0,'ellipse',f.color,3.5);pict(f.kind,f.x+23,700,62,f.color);txt(`flow-${i+1}-label`,f.label,f.x+121,694,f.w-128,65,35,'#ffffff',true);
  if(i<flows.length-1)lineArrow(`flow-${i+1}-to-${i+2}`,f.x+f.w+3,729,flows[i+1].x-10,729,f.color,7);
}
footer(true);

const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-connectors','12','--min-target-arrows','7'],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,'krill-operations-plan-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:5,requiredNativeChartOwnerSlides:[3],materializeLiteralChartWorkbooks:true,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
for(const [i,s] of imported.slides.items.entries()){
  const png=await imported.export({slide:s,format:'png',scale:1});await fs.writeFile(path.join(build,`preview-${i+1}.png`),Buffer.from(await png.arrayBuffer()));
}
console.log(JSON.stringify({finalPath,slides:5,budgetChartValues:[60,33,7],slideOrder:'User attachment order',illustrations:'Approximate native redraws; original photos unavailable'},null,2));
