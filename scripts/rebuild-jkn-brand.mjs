// Rebuild the JKN brand slide. Supply --reference to embed the original document crops.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/jkn-brand',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
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
  const paths={store:'<path d="M10 27v29h44V27M13 9h38l7 17c0 10-12 12-14 4-2 8-13 8-16 0-2 8-14 8-17 0-4 6-11 2-10-4z"/><path d="M18 48l8-7 8 6 7-3 7 8"/>',sofa:'<path d="M13 28v-8c0-17 38-17 38 0v8M14 43h36v12H14zM14 43v-7c-12-1-11-18 0-16 7 1 7 11 7 11h22s0-10 7-11c11-2 12 15 0 16v7M17 55v6m30-6v6"/>',building:'<path d="M9 59V16l20-7v50M29 59V6l21 5v48M3 59h54M37 17h6M37 27h6M37 37h6M16 25h6M16 35h6M16 45h6"/>',sprout:'<path d="M32 59V31M32 40C10 43 5 26 3 12c22 0 29 11 29 28zM32 31C36 12 48 8 61 9c-2 16-9 26-29 22zM32 40L13 22M32 31l17-13"/>',shield:'<path d="M32 3l25 9v17c0 17-13 28-25 34C20 57 7 46 7 29V12z" fill="#0b56be" stroke="#ffffff"/><path d="M18 31l10 10 19-23" stroke="#ffffff" stroke-width="6"/>',people:'<circle cx="32" cy="15" r="10"/><circle cx="12" cy="23" r="7"/><circle cx="52" cy="23" r="7"/><path d="M14 59V45c0-22 36-22 36 0v14zM3 52V41c0-11 9-14 14-10M61 52V41c0-11-9-14-14-10"/>',pin:'<path d="M32 60S9 35 9 22a23 23 0 0 1 46 0c0 13-23 38-23 38z"/><circle cx="32" cy="22" r="9"/>'};
  return slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 66"><g fill="none" stroke="${color}" stroke-width="3.6" stroke-linejoin="round" stroke-linecap="round">${paths[kind]}</g></svg>`,alt:`${kind} vector icon`,position:{left:x,top:y,width:size,height:size}});
}
function curve(name,x,y,w,h,commands,fill) {
  return slide.shapes.add({geometry:'custom',name,position:{left:x,top:y,width:w,height:h},fill,line:{fill:'none',width:0},customPaths:[{width:w,height:h,commands}]});
}
const M=(x,y)=>({moveTo:{x,y}}),L=(x,y)=>({lineTo:{x,y}}),B=(x1,y1,x2,y2,x,y)=>({cubicBezTo:{x1,y1,x2,y2,x,y}}),Z=()=>({close:{}});
slide.background.fill=grad('#ffffff','#f9fcff',65);
box('title-accent-shadow',71,47,23,84,'#edf4ff');
box('title-accent',77,47,13,84,grad('#1572ff','#003dd1',90));
text('company-title','台湾加康诺生物科技有限公司',114,33,1230,68,45,'#080e18',true);
text('company-subtitle','专注健康营养领域的生物科技企业',115,96,1130,40,27,'#181a1e');
box('header-divider',110,141,1490,1.5,'#c2d6ff');

// Logo recreated as independently editable nodes, lines and text.
const nodes=[[1477,32,'#ffcd00'],[1502,32,'#ffbf00'],[1520,54,'#ff9f00'],[1530,46,'#ffae00'],[1549,34,'#ffab00'],[1563,58,'#ff9400'],[1541,67,'#ff6b00'],[1505,65,'#8fbd00'],[1487,51,'#a6c600']];
for(const [i,j] of [[0,1],[1,2],[2,3],[3,4],[4,5],[5,6],[6,2],[2,7],[7,8]]) {
  const [x1,y1,c]=nodes[i],[x2,y2]=nodes[j];
  slide.shapes.add({name:`logo-edge-${i}-${j}`,geometry:'line',position:{left:Math.min(x1,x2),top:Math.min(y1,y2),width:Math.abs(x2-x1)||.1,height:Math.abs(y2-y1)||.1,verticalFlip:(x2-x1)*(y2-y1)<0},fill:'none',line:{fill:c,width:4}});
}
for(const [i,[x,y,c]] of nodes.entries())box(`logo-node-${i}`,x-6,y-6,12,12,c,0,'none','ellipse');
const logo=text('logo-wordmark','',1444,79,190,38,26,'#333333');
logo.text=[[{run:'JKN ',style:{typeface:FONT,fontSize:'26px',color:'#333333'}},{run:'Biotech',style:{typeface:FONT,fontSize:'26px',bold:true,color:'#333333'}}]];

curve('brand-ribbon',35,182,546,71,[M(32,0),L(546,0),L(520,71),L(0,71),B(-1,66,0,61,3,54),L(21,10),B(24,3,27,0,32,0),Z()],grad('#0041bd','#266bf8',0));
curve('brand-ribbon-trailing-1',557,182,45,71,[M(27,0),L(38,0),B(46,0,46,7,42,16),L(21,65),B(19,70,12,71,0,71),Z()],'#74b0ff');
curve('brand-ribbon-trailing-2',584,182,42,71,[M(25,0),L(35,0),B(43,0,43,7,39,16),L(18,65),B(16,70,9,71,0,71),Z()],'#bad7ff');
const medal='<svg xmlns="http://www.w3.org/2000/svg" width="80" height="90" viewBox="0 0 80 90"><path d="M22 45L8 85l20-5 11 9 4-35M44 45l20 39-20-6-11 11-4-35" fill="#ffffff"/><circle cx="40" cy="35" r="29" fill="#fff2b6"/><circle cx="40" cy="35" r="23" fill="#df9a20"/><circle cx="40" cy="35" r="19" fill="#1450a7"/><path d="M40 15l6 13 15 2-11 10 3 15-13-8-13 8 3-15-11-10 15-2z" fill="#ffe398"/></svg>';
slide.images.add({svg:medal,alt:'Brand medal illustration',position:{left:80,top:188,width:57,height:61}});
text('brand-heading','品牌实力',145,188,189,53,40,'#ffffff',true);
text('brand-english','B R A N D  S T R E N G T H',331,207,218,32,16,'#dceaff',false,{italic:true});

const rows=[
 {y:278,label:'台湾本土企业',detail:'台北市政府核准设立，合法合规经营',kind:'building'},
 {y:373,label:'专注营养健康',detail:'聚焦生物科技与营养保健产品研发与销售',kind:'sprout'},
 {y:473,label:'合规品质保障',detail:'产品符合台湾相关法规与质量标准',kind:'shield'},
 {y:576,label:'市场表现优异',detail:'深受消费者信赖，市场持续增长',kind:'people'},
];
for(const [i,r] of rows.entries()) {
  box(`claim-${i+1}-icon-bg`,57,r.y,119,72,grad('#f0f7ff','#d2e5ff'),17);
  svgIcon(r.kind,91,r.y+8,53,'#0750c5');
  text(`claim-${i+1}-title`,r.label,196,r.y-8,434,46,31.5,'#071b41',true);
  text(`claim-${i+1}-detail`,r.detail,196,r.y+34,438,37,i===1?21.4:22,'#171c23');
  if(i<3)box(`claim-${i+1}-divider`,196,r.y+83,433,1,'#dce8f9');
}

// Registration documents remain original image assets when a readable screenshot is supplied.
const docs=[{x:658,y:182,w:473,h:680,label:'台北市政府函'},{x:1148,y:182,w:496,h:680,label:'公司设立登记表'}];
for(const [i,d] of docs.entries()) {
  box(`document-${i+1}-frame`,d.x,d.y,d.w,d.h,'#ffffff',18,'#b7d1f9').shadow='shadow-sm';
  if(source) {
    const crop=await sharp(source).extract({left:d.x+7,top:d.y+8,width:d.w-14,height:d.h-16}).png().toBuffer();
    slide.images.add({blob:crop,contentType:'image/png',alt:d.label,position:{left:d.x+7,top:d.y+8,width:d.w-14,height:d.h-16},geometry:'roundRect',borderRadius:12});
  } else {
    box(`document-${i+1}-unavailable`,d.x+11,d.y+11,d.w-22,d.h-22,grad('#fbfdff','#f0f5fc'),12);
    text(`document-${i+1}-title`,d.label,d.x+25,d.y+255,d.w-50,59,29,'#7389aa',true,{alignment:'center'});
    text(`document-${i+1}-status`,'待补截图原文件',d.x+25,d.y+316,d.w-50,45,23,'#8b9ab2',false,{alignment:'center'});
  }
}

box('results-card',28,677,623,215,grad('#2671ff','#003182',80),24,'#7ba9fa');
text('results-title','经营成果',214,699,143,51,32,'#ffffff',true);
text('results-cutoff','（截至2026年10月）',350,700,286,48,24,'#ffffff');
box('results-label-rule',214,749,119,5,'#002975',2);
text('results-cumulative-sales','累计销售额',214,753,260,41,24,'#ffffff',true);
const num=text('results-number','1,000',209,771,248,110,86,'#f4c66b',true,{italic:true});
num.text.fill=grad('#fff9dd','#edac30',90);
num.text.outline={fill:'#ffffff',width:1.1,style:'solid'};
const wan=text('results-ten-thousand-unit','万',443,772,83,107,82,'#f4c66b',true,{italic:true});
wan.text.fill=grad('#fff5c7','#e6a334',90);
wan.text.outline={fill:'#ffffff',width:1.1,style:'solid'};
text('results-currency','台币',542,804,88,63,35,'#ffffff',true,{italic:true});
const trophy='<svg xmlns="http://www.w3.org/2000/svg" width="160" height="180" viewBox="0 0 160 180"><defs><linearGradient id="g"><stop stop-color="#cb7b10"/><stop offset=".34" stop-color="#ffeeb7"/><stop offset=".58" stop-color="#f7c95d"/><stop offset="1" stop-color="#ca7c15"/></linearGradient></defs><g fill="url(#g)" stroke="#fff2b9" stroke-width="2"><path d="M41 13h78l-7 55c-3 26-18 39-32 39S51 94 48 68z"/><path d="M41 22H18c-5 38 5 56 36 61l-4-12c-18-5-23-16-21-38h14zm78 0h23c5 38-5 56-36 61l4-12c18-5 23-16 21-38h-14z"/><path d="M71 106h18v33h20v10H51v-10h20zM45 151h70v16H45z"/><path d="M80 35l6 14 16 1-12 11 4 17-14-9-14 9 4-17-12-11 16-1z" fill="#d59425"/></g></svg>';
slide.images.add({svg:trophy,alt:'Gold trophy illustration',position:{left:66,top:707,width:115,height:127}});
// Gold laurel branches are separate editable native leaves, not a flattened results card.
for(const [branch,bx,by,sign] of [['left',61,766,-1],['right',185,766,1]]) {
  curve(`laurel-${branch}-stem`,bx-4,by,17,85,[M(sign<0?8:0,0),B(sign<0?-9:24,30,sign<0?-2:20,61,sign<0?17:-1,85)],'none').line={fill:'#f7cb61',width:2.4};
  for(let k=0;k<6;k++) {
    const yy=by+k*13,xx=bx+(k<3?-k*2:(k-4)*3)*sign;
    const leaf=box(`laurel-${branch}-leaf-${k}`,xx-5,yy,8,17,grad('#ffeb9b','#dca739'),0,'none','ellipse');
    leaf.position={left:xx-5,top:yy,width:8,height:17,rotation:sign*(24+k*5)};
  }
}
curve('results-gold-bottom',28,859,623,44,[M(0,10),B(165,67,378,-12,623,21),L(623,29),B(414,-2,182,71,0,24),Z()],grad('#c4811c','#fff1b8',0));
curve('results-white-shimmer',28,863,623,34,[M(0,7),B(182,54,384,-13,623,19),L(623,23),B(385,-6,180,57,0,11),Z()],'#fff2cf');

curve('footer-light-blue',0,877,1672,64,[M(0,31),B(390,75,781,60,1230,21),B(1439,-5,1538,-4,1672,13),L(1672,64),L(0,64),Z()],grad('#9fc8ff','#ddecff',0));
curve('footer-mid-blue',0,886,1672,55,[M(0,55),B(641,55,1254,-13,1672,8),L(1672,55),Z()],grad('#d9eaff','#83b7fd',0));
slide.speakerNotes.text=`来源：用户提供的台湾加康诺生物科技有限公司品牌实力截图。标题、公司介绍、经营成果、Logo及装饰按可见图片重建；Logo/图标/奖杯的外观为近似，非原始官方矢量文件。销售额及“截至2026年10月”按截图原样保留，未独立核验。${source?'两份登记文件从用户原始截图局部裁取，按图片对象内嵌，其小字和印章并非可编辑文字。':'截图原文件没有传入当前云端，两份登记文件原图缺失并明确标记。本文件仅为待补文件图的可编辑草稿，未模拟印章、未补写模糊文件内容，不能声称完整复原。'}`;
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp')],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,source?'jkn-brand-editable.pptx':'jkn-brand-editable-draft.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
await fs.writeFile(path.join(build,'status.json'),JSON.stringify({finalPath,missingDocumentImages:source?0:2,logo:'approximate editable redraw',normalization:JSON.parse(normal)},null,2));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),missingDocumentImages:source?0:2},null,2));
