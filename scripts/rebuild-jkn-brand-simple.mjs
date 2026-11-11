// Rebuild the latest simplified JKN company slide; --reference embeds original documents.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/jkn-brand-simple',`run-${Date.now()}`));
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
slide.background.fill='#ffffff';
box('title-accent-shadow',59,33,22,92,'#f1f6ff');
box('title-accent',64,33,14,92,grad('#1572ff','#003caf',90));
text('company-title','台湾加康诺生物科技有限公司',103,20,1270,71,48,'#080e18',true);
text('company-subtitle','专注健康营养领域的生物科技企业',104,87,1140,43,29,'#17191f');
box('header-divider',102,143,1519,1.5,'#c3d2ee');

const nodes=[[1473,43,'#ffcd00'],[1496,43,'#ffbf00'],[1516,64,'#ff9f00'],[1526,57,'#ffae00'],[1544,45,'#ffab00'],[1559,68,'#ff9400'],[1537,79,'#ff6b00'],[1501,76,'#8fbd00'],[1483,62,'#a6c600']];
for(const [i,j] of [[0,1],[1,2],[2,3],[3,4],[4,5],[5,6],[6,2],[2,7],[7,8]]) {
  const [x1,y1,c]=nodes[i],[x2,y2]=nodes[j];
  slide.shapes.add({name:`logo-edge-${i}-${j}`,geometry:'line',position:{left:Math.min(x1,x2),top:Math.min(y1,y2),width:Math.abs(x2-x1)||.1,height:Math.abs(y2-y1)||.1,verticalFlip:(x2-x1)*(y2-y1)<0},fill:'none',line:{fill:c,width:4}});
}
for(const [i,[x,y,c]] of nodes.entries())box(`logo-node-${i}`,x-6,y-6,12,12,c,0,'none','ellipse');
const logo=text('logo-wordmark','',1440,89,190,40,27,'#333333');
logo.text=[[{run:'JKN ',style:{typeface:FONT,fontSize:'27px',color:'#333333'}},{run:'Biotech',style:{typeface:FONT,fontSize:'27px',bold:true,color:'#333333'}}]];

curve('brand-ribbon',45,192,320,73,[M(28,0),L(320,0),L(292,73),L(0,73),Z()],grad('#034694','#1556e6',0));
curve('brand-ribbon-trailing',337,192,62,73,[M(28,0),L(62,0),L(33,73),L(0,73),Z()],'#b8d6ff');
text('brand-heading','品牌实力',102,194,237,65,48,'#ffffff',true);

const claims=[
 {y:303,title:'台湾本土企业',detail:'台北市政府核准设立，合法合规经营',kind:'building'},
 {y:403,title:'专注营养健康',detail:'聚焦生物科技与营养保健产品研发与销售',kind:'sprout'},
 {y:503,title:'合规品质保障',detail:'产品符合台湾相关法规与质量标准',kind:'shield'},
 {y:601,title:'市场表现稳定',detail:'深受消费者信赖，持续增长',kind:'people'},
];
for(const [i,c] of claims.entries()) {
  box(`claim-${i+1}-icon-bg`,62,c.y,106,71,grad('#f2f8ff','#e1edfc'),13);
  svgIcon(c.kind,89,c.y+9,53,'#084eb7');
  text(`claim-${i+1}-title`,c.title,195,c.y-8,438,48,32.5,'#061a3e',true);
  text(`claim-${i+1}-detail`,c.detail,196,c.y+34,438,37,i===1?21.4:22.5,'#14181f');
  if(i<3)box(`claim-${i+1}-divider`,63,c.y+85,570,1.3,'#dbe5f5');
}

const documents=[{x:661,y:198,w:473,h:678,label:'台北市政府函'},{x:1148,y:198,w:492,h:678,label:'公司设立登记表'}];
for(const [i,d] of documents.entries()) {
  box(`document-${i+1}-frame`,d.x,d.y,d.w,d.h,'#ffffff',17,'#bfcfe9');
  if(source) {
    const crop=await sharp(source).extract({left:d.x+7,top:d.y+8,width:d.w-14,height:d.h-16}).png().toBuffer();
    slide.images.add({blob:crop,contentType:'image/png',alt:d.label,position:{left:d.x+7,top:d.y+8,width:d.w-14,height:d.h-16},geometry:'roundRect',borderRadius:11});
  } else {
    box(`document-${i+1}-unavailable`,d.x+11,d.y+11,d.w-22,d.h-22,'#f7faff',11);
    text(`document-${i+1}-title`,d.label,d.x+25,d.y+255,d.w-50,59,29,'#7389aa',true,{alignment:'center'});
    text(`document-${i+1}-status`,'待补截图原文件',d.x+25,d.y+316,d.w-50,45,23,'#8b9ab2',false,{alignment:'center'});
  }
}

box('results-card',40,723,604,154,grad('#eff6ff','#e9f2ff'),16);
box('results-divider',173,744,1.3,111,'#97b5df');
for(const [i,x,y,h] of [[1,80,814,25],[2,104,804,35],[3,128,791,48]])box(`results-bar-${i}`,x,y,15,h,grad('#0b64d0','#00378a',90),1);
const start=box('results-arrow-start',82,800,1,1),end=box('results-arrow-end',142,767,1,1);
slide.shapes.connect(start,end,{kind:'curved',fromSide:'right',toSide:'left',line:{fill:'#0750ba',width:6},head:{type:'none'},tail:{type:'triangle',width:'med',length:'sm'}}).bringToFront();
text('results-title','经营成果',201,735,139,50,32,'#10151f',true);
text('results-cutoff','（截至2026年10月）',339,739,292,47,24,'#131820');
text('results-sales-label','累计销售额',201,802,145,49,27,'#10151f',true);
text('results-value','1,000万',354,781,223,82,59,'#0757d2',true,{italic:true});
text('results-currency','台币',575,807,63,41,27,'#11151c',true);
curve('results-blue-underline',365,855,178,12,[M(0,8),B(50,-2,102,-3,178,10),B(103,6,64,5,18,12),Z()],'#bad6fc');

slide.speakerNotes.text=`来源：用户最新提供的台湾加康诺品牌实力简洁版截图。保持1672×941画布；将第四条改为“市场表现稳定”“深受消费者信赖，持续增长”，并重建浅蓝经营成果卡片。销售额1,000万台币及截至2026年10月按图保留，未独立核验。Logo、图标和装饰是近似重绘。${source?'两份文件从用户提供的最新截图局部裁取后内嵌，文件内容仍是图片，不是可编辑文字。':'当前云端没有截图原文件，两份登记文件图片缺失；这是明确标记的草稿，未仿制文件扫描件、印章或补写模糊文字，不能声称完整复原。'}`;
const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-connectors','1','--min-target-arrows','1'],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,source?'jkn-brand-simple-editable.pptx':'jkn-brand-simple-editable-draft.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
await fs.writeFile(path.join(build,'status.json'),JSON.stringify({finalPath,missingDocumentImages:source?0:2,normalization:JSON.parse(normal)},null,2));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),missingDocumentImages:source?0:2},null,2));
