// Three reference screenshots: WeChat resources and two versions of company strength.
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {pathToFileURL,fileURLToPath} from 'node:url';
import {execFileSync} from 'node:child_process';
import assert from 'node:assert/strict';
const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const args=process.argv.slice(2),ix=args.indexOf('--work-dir');
const work=path.resolve(ix<0?path.join(project,'.local/private-resources-strength',`run-${Date.now()}`):args[ix+1]);
const build=path.join(work,'build'),output=path.join(work,'output');await fs.mkdir(build,{recursive:true});await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES&&process.env.PRESENTATIONS_SKILL_DIR,'Source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__private_resources__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1672,H=941,FONT='Noto Sans CJK SC',NAVY='#061942',BLUE='#0747d4',ORANGE='#ff8705',GRAY='#59677f';
if(process.env.PPT_OPERATION_MARKED!=='1')execFileSync(process.env.RUNTIME_NODE,[path.join(skill,'container_tools/mark_artifact_operation_started.mjs'),'--operation-kind','create','--expected-output-count','1','--output-format','pptx'],{stdio:'inherit'});
const deck=Presentation.create({slideSize:{width:W,height:H}});let slide;
const grad=(a,b,angleDeg=0)=>({type:'gradient',gradientKind:'linear',angleDeg,stops:[{offset:0,color:a},{offset:100000,color:b}]});
function box(name,x,y,w,h,fill='none',r=0,geometry='rect',stroke='none'){
 return slide.shapes.add({name,geometry,position:{left:x,top:y,width:w,height:h},fill,line:{fill:stroke,width:stroke==='none'?0:1},...(r&&geometry==='rect'?{borderRadius:r}:{})});
}
function text(name,value,x,y,w,h,size=24,color=NAVY,bold=false,extra={}){
 const s=box(name,x,y,w,h);s.text=value;s.text.style={typeface:FONT,fontSize:size,color,bold,autoFit:'none',insets:0,wrap:'none',verticalAlignment:'middle',lineSpacing:1.1,...extra};return s;
}
const M=(x,y)=>({moveTo:{x,y}}),L=(x,y)=>({lineTo:{x,y}}),B=(x1,y1,x2,y2,x,y)=>({cubicBezTo:{x1,y1,x2,y2,x,y}}),Z=()=>({close:{}});
function shape(name,x,y,w,h,commands,fill){return slide.shapes.add({name,geometry:'custom',position:{left:x,top:y,width:w,height:h},fill,line:{fill:'none',width:0},customPaths:[{width:w,height:h,commands}]});}
function poly(name,pts,fill){const x=Math.min(...pts.map(p=>p[0])),y=Math.min(...pts.map(p=>p[1])),w=Math.max(...pts.map(p=>p[0]))-x,h=Math.max(...pts.map(p=>p[1]))-y;return shape(name,x,y,w,h,[M(pts[0][0]-x,pts[0][1]-y),...pts.slice(1).map(p=>L(p[0]-x,p[1]-y)),Z()],fill);}
function panel(name,x,y,w,h,fill='#ffffff',r=23){box(`${name}-shadow`,x,y+5,w,h,'#b7cce4/35',r);return box(name,x,y,w,h,fill,r,'rect','#e6eef8');}
const paths={
 people:'<circle cx="34" cy="14" r="11"/><path d="M13 63V49c0-22 42-22 42 0v14zM12 8c-13 0-13 21 0 21M56 8c13 0 13 21 0 21M8 38c-9 0-11 9-11 20M60 38c9 0 11 9 11 20"/>',
 layers:'<path d="M4 19L34 4l30 15-30 15zM5 32l29 15 29-15M5 45l29 15 29-15"/>',
 calendar:'<rect x="7" y="13" width="55" height="50" rx="4"/><path d="M7 29h55M20 4v19M49 4v19M19 40h7M35 40h7M49 40h3M19 52h7M35 52h7"/>',
 diamond:'<path d="M4 22L19 5h31l16 17-31 43zM4 22h62M19 5l16 60L50 5M19 5l16 17L50 5"/>',
 building:'<path d="M5 65V18l44-9v56M49 30h15v35M14 28h7M29 25h7M14 40h7M29 37h7M14 52h7M29 49h7M2 66h66"/>',
 medical:'<rect x="7" y="20" width="55" height="43" rx="5"/><path d="M23 20v-9h22v9M35 29v25M23 41h24"/>',
 user:'<circle cx="34" cy="17" r="13"/><path d="M5 64c0-36 58-36 58 0"/>',
 contact:'<circle cx="27" cy="17" r="11"/><path d="M7 62c0-32 42-32 42 0M49 10c17-2 18 23 2 23M55 43c9 2 13 9 13 20"/>',
 compass:'<circle cx="34" cy="34" r="29"/><path d="M45 21l-6 18-18 8 7-19z"/>',
};
function icon(kind,x,y,size,color=BLUE){slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="84" height="84" viewBox="-7 -5 84 84"><g fill="none" stroke="${color}" stroke-width="3.8" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} icon redraw`,position:{left:x,top:y,width:size,height:size}});}
function wechat(name,x,y,w,h){
 box(`${name}-bg`,x,y,w,h,grad('#7ce878','#00aa35',50),w*.2);
 slide.images.add({svg:'<svg xmlns="http://www.w3.org/2000/svg" width="140" height="120" viewBox="0 0 140 120"><path d="M8 42C8 3 92 0 98 37 104 67 54 86 25 68L11 76l4-22z" fill="white"/><path d="M56 73c-4-34 67-45 73-8 4 20-13 37-26 39l4 13-20-13C72 103 57 93 56 73z" fill="#e9ffed" stroke="#38c955" stroke-width="2"/><g fill="#21bd4b"><circle cx="36" cy="32" r="5"/><circle cx="65" cy="31" r="5"/><circle cx="79" cy="66" r="4"/><circle cx="104" cy="63" r="4"/></g></svg>',alt:'WeChat logo approximation',position:{left:x+w*.06,top:y+h*.07,width:w*.88,height:h*.83}});
}
function newSlide(note){slide=deck.slides.add();slide.background.fill=grad('#ffffff','#f4f9ff',80);slide.speakerNotes.text=`来源：用户本轮按发送顺序提供的三张截图。${note} 数字及业务陈述按图转录，未独立核验财务口径或资源所有权。`;}
function waveFooter(orange=false){
 const y=H-58;
 shape('bottom-light-wave',0,y,W,58,[M(0,4),B(260,31,484,55,760,39),B(1103,12,1300,51,W,4),L(W,58),L(0,58),Z()],'#e6f1ff');
 shape('bottom-blue-wave',0,y+17,W,41,[M(0,11),B(275,18,433,55,785,28),B(1098,8,1325,57,W,3),L(W,41),L(0,41),Z()],grad('#86bdf4','#0350d4',0));
 if(orange)shape('bottom-orange-wave',0,y,W,58,[M(0,2),B(260,92,1140,55,W,0),L(W,17),B(1120,66,265,65,0,24),Z()],ORANGE);
}

// Page 1. Conversation names and unread-message indicators are text, not screenshots.
newSlide('第1张：WX私域资源展示。群聊标题按可辨部分转录，局部不清晰标题以省略号保留；40组头像为简化色块，消息预览仅保留可辨的未读条数并使用省略号，不补写姓名、电话号码或模糊消息。四列不是原始微信截图。“[条]”是消息预览，不能当作群成员数量。');
box('title-orange-rule',60,23,11,62,grad('#ffb64a','#ff7b00',90),6);
text('wx-title','WX',96,0,169,102,88,ORANGE,true,{italic:true});
text('resource-title','10万+私域资源部分展示',271,10,1172,88,68,NAVY,true,{italic:true});
const intro=text('resource-intro','深耕北京线下健康领域15年，凭借个人长期经营与积累，独立沉淀超10万私域用户资源。',98,102,1280,40,23,'#171b21',true);
intro.text.get('15年').color=ORANGE;intro.text.get('10万私域').color=ORANGE;
box('wechat-pedestal',1378,105,185,37,grad('#dfeaff','#8fb5eb',90),24);wechat('wechat-logo',1410,1,117,116);
panel('resource-metrics',59,161,1554,106,grad('#f7fbff','#e6f0ff'),52);
const metrics=[{x:122,kind:'people',title:'私域用户',value:'10万+',color:ORANGE},{x:503,kind:'layers',title:'社群数量',value:'200+'},{x:850,kind:'calendar',title:'月度活跃',value:'400–500人'},{x:1297,kind:'diamond',title:'资源沉淀',value:'15年+'}];
for(const [i,m] of metrics.entries()){
 icon(m.kind,m.x,188,63,m.color||'#1473ef');text(`metric-${i+1}-title`,m.title,m.x+91,171,Math.min(285,W-m.x-111),36,25,'#171b21',true);
 text(`metric-${i+1}-value`,m.value,m.x+86,202,Math.min(308,W-m.x-106),56,i===2?47:51,ORANGE,true,{italic:true});
 if(i<3)box(`metric-divider-${i+1}`,[439,787,1238][i],188,1.2,53,'#8baee8');
}
const chats=[
 {x:79,time:'12:22',names:['广州痕达门窗净化工程交流五群','广东医疗器械交流群','四川医疗器械设备交流群','医疗器械耗材信息交流群','二村交流群','石家庄医疗健康产业资源3群','中国净化工程/材料采购群','医用、民用、CE与FDA进出口…','山西净化工程交流群','共享会医疗器械招商群80'],counts:['286','234','338','545','279','581','213','321','61','510']},
 {x:463,time:'12:24',names:['全国医疗器械交流群六','河南洁净群','中国净化防护工程交流群','全国医院营销资源群','河南郑州医疗器械交流群','全国医疗器械资源求购群','北京冠邦洛阳交流群','二类（医疗器械）','全国医疗器械交流群','云南病房设备交流群…'],counts:['810','248','172','668','999+','643','276','818','753','573']},
 {x:847,time:'12:27',names:['北京医疗器械3群','全国医疗器械交流群','4S店中国医疗器械交流群','Y医疗设备、净化设备交流群','河南医疗器械','【陕西】医疗器械交流群','江苏…第三方一站采购平台','医疗器械3','药剂诊所药品器械采购群','湖北净化系统联系团队'],counts:['399','964','999+','395','434','706','714','999+','441','556']},
 {x:1231,time:'12:28',names:['医疗器械代加工、贴牌/…','医疗器械厂家经销商交流群','广东净化工程群','医疗器械交流买卖群','医疗器械-医用气体-净化工程','全国装饰工程供应链交流群','净化行业交流3群','广东净化群','暖通交流群','净化工程讨论群'],counts:['867','982','106','376','300','596','338','686','226','515']},
];
panel('chat-area',59,282,1554,613,'#f8fbff',28);
const avatarColors=['#c2d3cf','#304455','#b9aa86','#ece4d6','#3f747b','#82644d','#c3cfdf','#326ca4','#bcc1bd'];
for(const [c,k] of chats.entries()){
 const x=k.x,y=302,w=361;panel(`chat-${c+1}`,x,y,w,589,'#ffffff',18);box(`chat-${c+1}-header`,x+1,y+1,w-2,40,grad('#f4f4f4','#ededed',90),14);
 text(`chat-${c+1}-title`,'微信',x+145,y+5,80,30,18,'#131313',true,{alignment:'center'});box(`chat-${c+1}-plus-circle`,x+325,y+13,17,17,'none',0,'ellipse','#232323');text(`chat-${c+1}-plus`,'+',x+327,y+9,14,22,18,'#111111',false,{alignment:'center'});
 for(const [r,name] of k.names.entries()){
  const yy=y+42+r*49;
  for(let tile=0;tile<9;tile++)box(`avatar-${c+1}-${r+1}-${tile}`,x+18+(tile%3)*10.5,yy+8+Math.floor(tile/3)*10.5,9.5,9.5,avatarColors[(tile+r+c)%9],1.5);
  if(r%3!==1)box(`chat-${c+1}-row-${r+1}-dot`,x+45,yy+6,5,5,'#fb3838',0,'ellipse');
  text(`chat-${c+1}-row-${r+1}-name`,name,x+63,yy+1,232,24,name.length>17?12.2:14.2,'#121212',true);
  text(`chat-${c+1}-row-${r+1}-time`,`下午${c===0&&r>2?'12:21':k.time}`,x+294,yy+4,56,20,10.4,'#a2a2a2',false,{alignment:'right'});
  text(`chat-${c+1}-row-${r+1}-preview`,`[${k.counts[r]}条] …`,x+63,yy+24,260,20,11.7,'#a1a1a1');
  box(`chat-${c+1}-row-${r+1}-rule`,x+14,yy+48,w-28,1,'#eeeeee');
 }
 box(`chat-${c+1}-footer-bg`,x+1,y+533,w-2,53,'#fafafa',14);
 for(const [j,label] of ['微信','通讯录','发现','我'].entries()){
  const xx=x+34+j*88;
  if(j===0)wechat(`chat-${c+1}-nav-wechat`,xx+1,y+545,22,21);else icon(['','contact','compass','user'][j],xx,y+541,27,'#131313');
  text(`chat-${c+1}-nav-${j}`,label,xx-7,y+567,44,18,10.2,j===0?'#08a741':'#222222',false,{alignment:'center'});
 }
}
text('chat-rebuild-disclosure','群聊界面按可辨内容重建，头像与模糊消息预览已简化。',80,274,1150,20,11.5,'#7b89a0');waveFooter(true);

// Shared source data: revenue values are text cards, not a fabricated chart.
const businesses=[{title:'台湾加康诺生物科技\n有限公司',revenue:200},{title:'安徽省芊伊雪茄吧',revenue:300},{title:'四川省线上网络游戏\n公司',revenue:3500}];
assert.equal(businesses.reduce((s,c)=>s+c.revenue,0),4000);
function annualBanner(y,large=false){
 const x=large?862:801,blueW=large?370:389,totalW=large?750:800,h=103;
 poly('annual-blue',[[x+34,y],[x+blueW,y],[x+blueW-39,y+h],[x,y+h]],grad('#244e9e','#061f57',90));
 poly('annual-orange',[[x+blueW-6,y],[x+totalW,y],[x+totalW-30,y+h],[x+blueW-42,y+h]],grad('#ffb132','#ff8400',90));
 text('annual-heading','合计年营收',x+65,y+5,blueW-85,h-10,50,'#ffffff',true,{italic:true});
 text('annual-amount','4,000',x+blueW+(large?-5:9),y-5,large?252:263,h+8,large?86:90,'#ffffff',true,{italic:true});text('annual-unit','万元',x+totalW-(large?118:132),y+22,105,68,46,'#ffffff',true,{italic:true});
}
function section(label,y,rich,kind){
 if(rich){box(`${kind}-section-highlight`,105,y+1,211,43,'#e4f0ff',18);icon(kind,68,y+3,42);text(`${kind}-section-title`,label,126,y-2,260,47,30,NAVY,true);box(`${kind}-section-line`,331,y+24,1275,1.2,'#a8bfe7');}
 else {text(`${kind}-section-title`,label,72,y,222,41,28,NAVY,true);box(`${kind}-section-line`,253,y+23,1348,1.2,'#c8d7ed');}
}
function strengthHeader(rich){
 text('section-number','01',65,rich?21:64,104,119,rich?92:84,ORANGE,true);box('title-divider',rich?180:170,rich?38:65,1.4,rich?113:79,'#c2d5f1');
 text('strength-title','企业综合实力',208,rich?25:66,592,83,64,NAVY,true);text('strength-subtitle','多主体经营布局 · 跨区域业务协同 · 核心资源沉淀',209,rich?100:139,rich?650:585,43,rich?26:23,'#485d7b');
 if(rich)box('title-orange-accent',210,148,82,4,ORANGE,2);else {box('top-progress-bg',1347,48,254,12,'#e6eeff',6);box('top-progress-active',1347,48,127,12,ORANGE,6);}
 annualBanner(rich?48:89,rich);
}
const teamValues=['15年+','10万+','200+','400–500人'],teamLabels=['健康领域线下经验','私域资源积累','社群资源','月均社群触达'];
function teamResources(rich){
 const xlist=rich?[85,478,849,1222]:[73,460,847,1234];
 for(let i=0;i<4;i++){
  const x=xlist[i],y=rich?595:686;
  if(rich){panel(`team-${i+1}-icon-box`,x,y,87,85,'#ffffff',18);icon(['medical','layers','people','user'][i],x+15,y+18,58,'#155de1');}
  else box(`team-${i+1}-orange`,x+2,y+6,15,15,ORANGE,4);
  text(`team-${i+1}-value`,teamValues[i],x+(rich?109:32),y-8,rich?294:350,68,i===3?49:51,BLUE,true);
  text(`team-${i+1}-label`,teamLabels[i],x+(rich?109:32),y+55,rich?294:350,40,rich?25:21,GRAY);
  if(i<3)box(`team-divider-${i+1}`,rich?[443,812,1185][i]:[434,821,1208][i],y+1,1.2,rich?83:94,'#c5d7f3');
 }
}
function summary(rich){
 const y=rich?734:817,h=rich?85:75;panel('summary-band',67,y,1538,h,grad('#f1f6fc','#eef3fa'),18);
 text('summary-businesses','多元业务主体',rich?93:97,y+15,194,43,rich?26:23,NAVY,true);
 text('summary-plus-1','+',273,y+10,43,49,35,ORANGE,true);text('summary-health','健康行业经验',336,y+15,194,43,rich?26:23,NAVY,true);
 text('summary-plus-2','+',518,y+10,43,49,35,ORANGE,true);text('summary-private','私域与社群资源',577,y+15,222,43,rich?26:23,NAVY,true);
 box('summary-divider-1',rich?778:787,y+14,1.2,h-28,'#c2d2e9');
 poly('summary-annual-bg',[[rich?821:837,y],[rich?1338:1331,y],[rich?1310:1308,y+h],[rich?795:811,y+h]],grad('#fff0da','#ffe4b9',90));
 text('summary-annual-label','合计年营收',rich?845:856,y+12,208,h-20,rich?36:32,NAVY,true);
 text('summary-annual-amount','4,000',rich?1037:1047,y-6,rich?205:201,h+7,rich?68:64,ORANGE,true,{italic:true});text('summary-annual-unit','万元',rich?1250:1252,y+19,80,47,rich?34:33,ORANGE,true,{italic:true});
 box('summary-divider-2',rich?1370:1362,y+14,1.2,h-28,'#c2d2e9');text('summary-conclusion','共同构成项目拓展\n与市场运营基础',1402,y+9,207,h-13,rich?25:22,GRAY);
 text('source-note',rich?'注：私域资源及健康领域经验为核心成员长期积累；营收数据来自各主体口径展示。':'注：私域资源及健康领域经验为核心成员长期积累；营收数据按提供口径展示。',72,rich?843:902,1495,30,rich?16:14,'#687d9c');
}
// Page 2: clean version.
newSlide('第2张：企业综合实力简洁版。年营收200、300、3,500万元，算术合计4,000万元。并列经营卡片不代表财务合并报表；按提供口径展示，不补充报告期、审计或主体控制关系。');
strengthHeader(false);section('企业经营规模',231,false,'building');
for(const [i,c] of businesses.entries()){
 const x=[72,591,1110][i];panel(`company-${i+1}`,x,288,490,286,'#ffffff',26);box(`company-${i+1}-top`,x,288,490,12,BLUE,6);
 text(`company-${i+1}-number`,String(i+1).padStart(2,'0'),x+28,322,45,48,23,ORANGE,true);box(`company-${i+1}-divider`,x+83,330,1.2,38,'#c9d8ef');
 text(`company-${i+1}-title`,c.title,x+109,317,359,94,30,NAVY,true);box(`company-${i+1}-rule`,x+29,431,432,1.2,'#cddbf0');
 text(`company-${i+1}-annual-label`,'年营收',x+35,464,272,43,25,GRAY);text(`company-${i+1}-annual-value`,c.revenue.toLocaleString('en-US'),x+35,511,i===2?177:111,68,61,ORANGE,true);
 text(`company-${i+1}-annual-unit`,'万元',x+(i===2?213:150),531,104,40,28,NAVY,true);
}
section('核心团队资源',627,false,'people');teamResources(false);summary(false);

// Page 3: illustrated version; buildings are visibly marked as non-original.
newSlide('第3张：企业综合实力插图版。原始建筑照片不可读取，三块照片区域用明确标注“建筑示意（非原图）”的可编辑几何插图替代，不能视为企业办公场所实拍。企业名称、数字与本页原文保留。');
strengthHeader(true);section('企业经营规模',183,true,'building');
function buildingArt(x,y,w,h,store){
 if(store){
  poly('store-structure',[[x+10,y+48],[x+w,y],[x+w,y+h],[x+10,y+h]],'#697b8c');poly('store-front',[[x+15,y+65],[x+w-6,y+23],[x+w-6,y+h-5],[x+15,y+h-5]],'#a7bac8');
  for(let c=0;c<5;c++)poly(`store-window-${c}`,[[x+24+c*29,y+82-c*7],[x+49+c*29,y+76-c*7],[x+49+c*29,y+h-9],[x+24+c*29,y+h-9]],grad('#d8c0a0','#67503b',90));
  poly('store-canopy',[[x+9,y+48],[x+w,y],[x+w,y+26],[x+9,y+72]],'#586674');
 }else{
  poly('glass-building',[[x+8,y+57],[x+w*.68,y],[x+w,y+51],[x+w,y+h],[x+8,y+h]],'#9fb8d0');
  poly('glass-building-front',[[x+8,y+57],[x+w*.68,y],[x+w*.68,y+h],[x+8,y+h]],'#c4d9eb');
  for(let c=1;c<8;c++){const xx=x+8+c*(w*.68-8)/8;const yy=y+57-c*57/8;box(`glass-mullion-${c}`,xx,yy,1.2,y+h-yy,'#7594b0');}
  for(let r=1;r<7;r++)poly(`glass-floor-${r}`,[[x+8,y+57+r*(h-57)/7],[x+w*.68,y+r*h/7],[x+w,y+51+r*(h-51)/7],[x+w,y+53+r*(h-51)/7],[x+w*.68,y+2+r*h/7],[x+8,y+59+r*(h-57)/7]],'#8eabc2');
 }
}
for(const [i,c] of businesses.entries()){
 const x=[66,594,1119][i],w=i===1?484:489;panel(`company-${i+1}`,x,251,w,239,'#ffffff',24);
 poly(`company-${i+1}-corner`,[[x+18,251],[x+96,251],[x+77,301],[x,301],[x,270]],grad('#398ee5','#073c9e',90));text(`company-${i+1}-number`,String(i+1).padStart(2,'0'),x+25,252,58,44,32,'#ffffff',true);
 text(`company-${i+1}-title`,c.title,x+107,258,w-123,82,29,NAVY,true);box(`company-${i+1}-rule`,x+41,349,268,1.2,'#b7d0ed');
 buildingArt(x+w-184,339,180,149,i===1);box(`company-${i+1}-image-disclosure-bg`,x+w-185,466,184,21,'#ffffff/90');text(`company-${i+1}-image-disclosure`,'建筑示意（非原图）',x+w-185,466,184,21,12.5,'#315473',true,{alignment:'center'});
 text(`company-${i+1}-annual-label`,'年营收',x+40,367,202,46,30,GRAY);text(`company-${i+1}-annual-value`,c.revenue.toLocaleString('en-US'),x+40,405,i===2?185:111,75,62,ORANGE,true);
 text(`company-${i+1}-annual-unit`,'万元',x+(i===2?228:154),428,85,43,29,NAVY,true);
}
section('核心团队资源',526,true,'people');teamResources(true);summary(true);waveFooter(false);

const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx'),finalPath=path.join(output,'private-resources-company-strength-editable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
await fs.writeFile(path.join(build,'normalization.json'),execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp')],{encoding:'utf8'}));
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:3,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));for(const [i,s] of imported.slides.items.entries()){const png=await imported.export({slide:s,format:'png',scale:1});await fs.writeFile(path.join(build,`preview-${i+1}.png`),Buffer.from(await png.arrayBuffer()));}
await fs.writeFile(path.join(build,'source-data.json'),JSON.stringify({businesses,totalRevenue:4000,teamValues,teamLabels,chats},null,2));
console.log(JSON.stringify({finalPath,slides:3,limitations:'Chat avatars and unreadable previews simplified; slide 3 original building photos replaced with labeled vector illustrations.'}));
