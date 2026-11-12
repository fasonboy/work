// Single-slide reconstruction of the revised production-landscape screenshot.
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {pathToFileURL,fileURLToPath} from 'node:url';
import {execFileSync} from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const args=process.argv.slice(2),index=args.indexOf('--work-dir');
const work=path.resolve(index<0?path.join(project,'.local/production-landscape',`run-${Date.now()}`):args[index+1]);
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true});await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES&&process.env.PRESENTATIONS_SKILL_DIR,'Source scripts/ppt-runtime.sh first');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__production_landscape__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1677,H=938,FONT='Noto Sans CJK SC',NAVY='#04143d',BLUE='#0045e5';
if(process.env.PPT_OPERATION_MARKED!=='1')execFileSync(process.env.RUNTIME_NODE,[path.join(skill,'container_tools/mark_artifact_operation_started.mjs'),'--operation-kind','create','--expected-output-count','1','--output-format','pptx'],{stdio:'inherit'});
const deck=Presentation.create({slideSize:{width:W,height:H}}),slide=deck.slides.add();
slide.background.fill='#ffffff';
slide.speakerNotes.text='来源：用户本轮提供的单张新版国内生产格局截图。按图转录企业名称、产能、合作、生产条件与项目核验说明，未独立核验企业原始材料；不补造参考资料。截图中的船名按当前图面“福六海”转录，须结合企业原始材料核验。右上冰山装饰与6个图标为近似重绘，不是企业照片。';
const grad=(a,b,angleDeg=0)=>({type:'gradient',gradientKind:'linear',angleDeg,stops:[{offset:0,color:a},{offset:100000,color:b}]});
function box(name,x,y,w,h,fill='none',r=0,stroke='none'){
 return slide.shapes.add({name,geometry:'rect',position:{left:x,top:y,width:w,height:h},fill,line:{fill:stroke,width:stroke==='none'?0:1},...(r?{borderRadius:r}:{})});
}
function text(name,value,x,y,w,h,size=22,color=NAVY,bold=false,wrap=false){
 const s=box(name,x,y,w,h);s.text=value;s.text.style={typeface:FONT,fontSize:size,color,bold,autoFit:'none',insets:0,wrap:wrap?'square':'none',verticalAlignment:'middle',lineSpacing:1.15};return s;
}
function polygon(name,pts,fill){
 const x=Math.min(...pts.map(p=>p[0])),y=Math.min(...pts.map(p=>p[1])),w=Math.max(...pts.map(p=>p[0]))-x,h=Math.max(...pts.map(p=>p[1]))-y;
 return slide.shapes.add({name,geometry:'custom',position:{left:x,top:y,width:w,height:h},fill,line:{fill:'none',width:0},customPaths:[{width:w,height:h,commands:[{moveTo:{x:pts[0][0]-x,y:pts[0][1]-y}},...pts.slice(1).map(p=>({lineTo:{x:p[0]-x,y:p[1]-y}})),{close:{}}]}]});
}
function panel(name,x,y,w,h,fill='#ffffff',r=21){box(`${name}-shadow`,x,y+4,w,h,'#d9dfea/65',r);return box(name,x,y,w,h,fill,r,'#edf0f7');}
const icons={
 factory:'<path d="M6 64V29l17 10V25l17 12V13h19v51zM11 30V6h10v29M45 13V3h10v10M16 52h6M30 52h6M46 52h6"/>',
 handshake:'<path d="M3 22l9-12 13 5 10-5 15 6 9-3 10 18-8 10-8 7-8 8-8 7-7-3-7-7-7-7-9-7zM25 15L13 29c-3 6 4 9 8 5l11-9 20 20M31 30l23 22M23 39l22 21M17 47l17 17M50 16l11 25M12 12l-6 23"/>',
 lab:'<path d="M26 5h23M31 5v21L11 59c-2 5 1 8 6 8h44c5 0 7-4 4-8L44 26V5M21 44h32M26 53h1M44 57h1M34 39h1"/>',
 gear:'<path d="M29 3h12l3 9 8 4 9-2 7 11-6 7v9l6 7-7 11-9-2-8 5-3 9H29l-3-9-8-5-9 2-7-11 6-7v-9l-6-7 7-11 9 2 8-4z"/><circle cx="35" cy="37" r="13"/>',
 target:'<circle cx="32" cy="39" r="28"/><circle cx="32" cy="39" r="18"/><circle cx="32" cy="39" r="8"/><path d="M32 39L62 8M54 1l-2 16 16-2"/>',
};
function icon(kind,x,y,size,color=BLUE){
 slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="84" height="84" viewBox="-7 -5 84 84"><g fill="none" stroke="${color}" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round">${icons[kind]}</g></svg>`,alt:`${kind} icon redraw`,position:{left:x,top:y,width:size,height:size}});
}

text('research-label','MARKET RESEARCH  /  南极磷虾油项目',63,13,1180,31,18,BLUE,true);
text('page-title','国内生产端逐步形成原料、精深加工与终端品牌协同格局',62,48,1558,65,45,NAVY,true);
text('page-subtitle','品牌方、原料供应商与实际生产企业可能不是同一主体，具体关系须落实到SKU标签与备案资料',63,116,1553,38,21,NAVY);
box('header-rule',63,161,1555,2,'#edf0f4');box('header-accent',63,162,97,6,BLUE);

panel('upstream-banner',62,192,1557,104,grad('#f0f5ff','#eaf2ff'),20);
icon('factory',116,219,53);box('upstream-icon-rule',194,219,1,49,'#b8cbe7');
text('upstream-label','上游原料',217,215,155,55,30,NAVY,true);box('upstream-label-rule',383,219,1,51,'#b8cbe7');
text('upstream-body','挪威 Aker BioMarine | 全球南极磷虾原料供应商之一，与 VIK 等品牌存在原料及研发合作；\n原料端优势集中在捕捞、加工与供应链管理。',411,205,922,76,21,NAVY,false,true);
// Reference decoration is redrawn as independently editable native shapes.
box('ice-water',1376,264,242,29,grad('#e3eff8','#4f89ad',90),8);
polygon('ice-back',[[1337,261],[1379,248],[1399,222],[1420,236],[1450,209],[1469,228],[1505,201],[1533,217],[1576,194],[1594,216],[1617,222],[1617,263]],'#dcebf5');
polygon('ice-main',[[1400,261],[1427,243],[1454,208],[1476,231],[1505,204],[1536,239],[1569,213],[1593,229],[1617,252],[1617,265]],'#9bbbd3');
polygon('ice-snow',[[1427,243],[1454,208],[1461,225],[1453,219],[1443,240],[1452,231],[1463,243],[1476,231],[1469,249],[1492,230],[1505,204],[1514,220],[1509,218],[1521,243],[1502,229],[1493,250],[1472,259],[1441,255]],'#f7fbfe');
polygon('ice-right-snow',[[1536,239],[1569,213],[1579,221],[1573,221],[1585,241],[1593,229],[1603,245],[1617,252],[1617,265],[1598,257],[1578,259],[1559,246],[1543,257]],'#f0f7fc');

const companies=[
 {x:62,y:322,w:768,h:198,kind:'factory',title:'辽渔南极磷虾科技发展有限公司',sub:'辽渔集团全资子公司 | 陆上精深加工与销售',body:'产品覆盖磷虾油、凝胶糖果、磷虾肽、磷虾粉。2026年公开首发\n60%磷脂磷虾油及“极之极”磷虾油番茄红素软胶囊；原料来自\n“福六海”轮捕捞并船载加工。'},
 {x:853,y:322,w:766,h:198,kind:'handshake',title:'逢时（青岛）海洋科技股份有限公司',sub:'VIK 品牌 | 品牌运营与自有生产体系',body:'布局透明工厂，与挪威 Aker BioMarine 开展原料及产品合作；企业公开披露旗下业务进入2,100余家线下门店，覆盖会员商超及药房。'},
 {x:62,y:542,w:768,h:184,kind:'lab',title:'山东鲁华海洋生物科技有限公司',sub:'山东健康集团旗下 | 专业磷虾油生产商',body:'官方资料披露年产能约500吨，产品覆盖南极磷虾油、凝胶糖果、膏剂、微囊粉、片剂和滴剂等形态。'},
 {x:853,y:542,w:766,h:184,kind:'gear',title:'青岛南极维康生物科技有限公司',sub:'研发、原料提取与终端生产',body:'公开企业资料披露拥有SC生产许可、10万级GMP车间及年产\n1,000吨磷虾油原料生产线，具备凝胶糖果等终端产品生产能力。'},
];
for(const [i,c] of companies.entries()){
 panel(`company-${i+1}`,c.x,c.y,c.w,c.h);if(i===0)box('company-1-left-accent',c.x,c.y+9,6,c.h-18,BLUE,3);
 box(`company-${i+1}-icon-bg`,c.x+43,c.y+20,71,68,grad('#0043d4','#7cadd9',0),15);icon(c.kind,c.x+55,c.y+32,48,'#ffffff');
 box(`company-${i+1}-divider`,c.x+135,c.y+20,1,c.h-61,'#d6e0f0');
 text(`company-${i+1}-heading`,c.title,c.x+155,c.y+9,c.w-171,43,29,NAVY,true);
 text(`company-${i+1}-subtitle`,c.sub,c.x+155,c.y+53,c.w-171,39,22,BLUE,true);
 text(`company-${i+1}-body`,c.body,c.x+155,c.y+100,c.w-176,c.h-109,20.5,NAVY,false,!c.body.includes('\n'));
}

panel('project-impact',62,761,1557,121,grad('#f1f5ff','#eff4ff'),20);
polygon('project-impact-blue',[[84,761],[393,761],[359,882],[84,882],[62,860],[62,783]],grad('#0045e5','#0538b4',30));
icon('target',88,793,51,'#ffffff');text('project-impact-label','对本项目的意义',157,790,220,59,27,'#ffffff',true);box('project-impact-rule',422,776,1,92,'#b8cce9');
text('project-impact-headline','辽渔具备捕捞、船载加工、陆上精深加工、科研平台与保健食品开发基础，产业链协同优势突出。',449,764,1138,39,25,NAVY,true);
text('project-sku-verification','项目拟议的56%磷脂 / 60粒SKU，不等同于辽渔2026年公开的60%新品，仍须单独核验规格、产品属性、备案/注册、\n委托生产关系及批次检测报告。',449,808,1138,62,21.5,NAVY);

const raw=path.join(build,'raw.pptx'),portable=path.join(build,'portable.pptx'),finalPath=path.join(output,'production-landscape-editable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
await fs.writeFile(path.join(build,'normalization.json'),execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',raw,'--output',portable,'--sharp-module',req.resolve('sharp')],{encoding:'utf8'}));
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit'],explicitTotalSlideCount:1,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,slides:1,limitation:'Icons and small iceberg decoration are approximate redraws.'}));
