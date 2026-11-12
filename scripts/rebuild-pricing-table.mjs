// Rebuild the revised pricing screenshot as one native table and three design cards.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/pricing-table',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1672,H=941,FONT='Noto Sans CJK SC',NAVY='#133f70',RED='#e34442',BLUE='#347fc1';
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
const quantities=[1,2,4,6,8,10,12],prices=[458,447,435,424,412,401,389],sourceTotals=[458,894,1740,2544,3296,4010,4668];
const totals=quantities.map((n,i)=>n*prices[i]);assert.deepEqual(totals,sourceTotals);
const money=n=>`¥${n.toLocaleString('en-US')}`;
slide=deck.slides.add();slide.background.fill=grad('#ffffff','#f0f9ff',65);
const paths={
 bottle:'<path d="M22 14V5h25v9M19 14h31v7H19zM23 21c0 8-9 8-9 16v27c0 4 3 6 6 6h27c4 0 7-2 7-6V37c0-8-10-8-10-16M15 38h38M27 48h15"/>',
 target:'<circle cx="31" cy="36" r="27"/><circle cx="31" cy="36" r="17"/><circle cx="31" cy="36" r="7"/><path d="M31 36L60 7M53 1l-2 13 13-2"/>',
 people:'<circle cx="34" cy="15" r="11"/><path d="M13 62V49c0-22 42-22 42 0v13zM12 8c-13 0-13 21 0 21M56 8c13 0 13 21 0 21M8 38c-9 0-11 9-11 20M60 38c9 0 11 9 11 20"/>',
 box:'<path d="M5 22l9-14h42l8 14v40H5zM5 22h59M34 9v53M28 31h12v18H28z"/>',
 trend:'<path d="M6 32L24 17l13 10L63 4M48 4h15v15M7 64V47h11v17zM29 64V37h11v27zM51 64V26h11v38z"/>',
};
function pict(kind,x,y,w,h=w,color=NAVY){slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="84" height="84" viewBox="-7 -4 84 84"><g fill="none" stroke="${color}" stroke-width="4.2" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} icon`,position:{left:x,top:y,width:w,height:h}});}
// Independent native snow facets approximate the decorative photograph.
box('polar-water',974,111,698,27,grad('#cdefff','#ffffff/0',90));
poly('polar-back',[[960,121],[1101,74],[1162,91],[1274,43],[1330,59],[1488,6],[1553,72],[1605,63],[1672,110],[1672,137]],'#d7effc');
poly('polar-front',[[1090,133],[1255,87],[1328,107],[1488,6],[1537,78],[1577,62],[1672,116],[1672,139]],'#99d0ef');
for(const [i,p] of [[1,[[1488,6],[1410,119],[1469,80],[1480,126],[1510,46]]],[2,[[1488,6],[1537,78],[1577,62],[1529,126]]],[3,[[1274,43],[1220,126],[1262,90],[1294,124],[1330,59]]]])poly(`polar-snow-${i}`,p,'#ffffff');
box('title-accent',43,21,14,62,RED,3);txt('page-title','五、产品经营思路：规格分层与价格梯度设计',91,13,1330,66,43,NAVY,true);
txt('page-subtitle','以阶梯价格提升客单价与长期复购，覆盖不同需求人群',93,85,1192,43,27,NAVY,true);
function panel(name,x,y,w,h,title,kind,headH=69){
 box(`${name}-shadow`,x,y+5,w,h,'#cadaea/70',29);box(name,x,y,w,h,'#ffffff',29,'rect','#dbe8f4',1.2);
 curve(`${name}-header`,x,y,w-77,headH,[M(29,0),L(w-77,0),L(w-127,headH),L(0,headH),L(0,29),B(0,11,11,0,29,0),Z()],grad('#0e78b8','#07396b'));
 poly(`${name}-accent-ice`,[[x+w-127,y],[x+w-77,y],[x+w-127-headH*.7,y+headH],[x+w-177,y+headH]],'#bde6fd/65');
 poly(`${name}-accent-red`,[[x+w-109,y],[x+w-101,y],[x+w-151,y+headH],[x+w-159,y+headH]],'#ed706e');
 box(`${name}-badge`,x+33,y+1,69,69,'#43a3d6/30',0,'ellipse');pict(kind,x+43,y+13,48,48,'#ffffff');
 txt(`${name}-heading`,title,x+120,y+7,w-182,headH-8,34,'#ffffff',true);
}
panel('pricing-panel',54,165,774,561,'规格与价格（单瓶价格阶梯）','bottle');
const values=[['规格','单瓶价','总价'],...quantities.map((n,i)=>[`${n}瓶`,money(prices[i]),money(totals[i])])];
const table=slide.tables.add({rows:8,columns:3,left:75,top:256,width:732,height:451,columnWidths:[195,270,267],values});
for(let r=0;r<8;r++){
 table.rows[r].height=r===0?54:397/7;
 for(let c=0;c<3;c++){
  const cell=table.getCell(r,c);cell.fill=r===0?'#cfe3f2':r===7?'#ffeded':r%2?'#f6fbff':'#ffffff';
  cell.text.style={typeface:FONT,fontSize:r===0?28:c===1?35:32,color:r>0&&c===1?'#bf231c':NAVY,bold:true,autoFit:'none',wrap:'none',alignment:'center',verticalAlignment:'middle'};
  table.cells.block({row:r,column:c,rowCount:1,columnCount:1}).assign({margins:{left:8,right:8,top:4,bottom:4}});
 }
}
table.borders.assign({color:'#d7e5f2',fill:'#d7e5f2',width:1.2,style:'solid'});
poly('recommended-ribbon',[[75,650],[130,650],[75,706]],'#e13b35');
const recommendation=txt('recommended-label','推荐',77,661,43,24,17,'#ffffff',true,{alignment:'center'});recommendation.position={left:77,top:661,width:43,height:24,rotation:-45};

panel('design-panel',875,171,759,548,'规格设计思路','target',73);
const design=[
 {y:271,title:'入门体验',detail:'1–2瓶：降低尝试门槛，快速建立信任。',kind:'bottle'},
 {y:414,title:'主力规格',detail:'4–8瓶：兼顾性价比与客单提升。',kind:'people'},
 {y:557,title:'长期管理',detail:'10–12瓶：适合长期服用、家庭囤货与送礼。',kind:'box'},
];
for(const [i,d] of design.entries()){
 box(`design-card-${i+1}-shadow`,894,d.y+4,725,109,'#d5e7f3/45',18);box(`design-card-${i+1}`,894,d.y,725,109,grad('#f3faff','#eaf5fd'),18);
 box(`design-card-${i+1}-badge`,912,d.y+13,86,83,'#e3f3fe',0,'ellipse');pict(d.kind,929,d.y+27,57,57,'#135e94');
 txt(`design-card-${i+1}-heading`,d.title,1017,d.y+24,136,58,32,NAVY,true);box(`design-card-${i+1}-divider`,1152,d.y+38,1.5,32,'#2f6c96');
 txt(`design-card-${i+1}-detail`,d.detail,1174,d.y+22,443,67,21,NAVY,false,{wrap:'square',lineSpacing:1.18});
}
box('conclusion-panel-shadow',54,770,1574,89,'#c8dff2/65',27);box('conclusion-panel',54,765,1574,89,grad('#ffffff','#edf8ff'),27);
curve('conclusion-label-bg',54,765,335,89,[M(27,0),L(335,0),L(275,89),L(27,89),B(9,89,0,80,0,62),L(0,27),B(0,9,9,0,27,0),Z()],grad('#1765ac','#00447a'));
pict('trend',91,785,57,57,'#ffffff');txt('conclusion-label','经营核心',168,778,182,63,38,'#ffffff',true);
txt('conclusion','阶梯定价，层级引导，提升客单价与复购率。',531,778,1070,65,40,NAVY,true);
curve('footer-light',0,869,W,72,[M(0,0),B(510,91,1064,18,1414,14),B(1524,8,1593,11,W,18),L(W,72),L(0,72),Z()],'#bde6fc');
curve('footer-blue',0,887,W,54,[M(0,0),B(610,83,1108,31,1440,1),B(1550,-5,1613,0,W,2),L(W,54),L(0,54),Z()],grad('#1298e7','#78cafd'));
slide.speakerNotes.text='来源：用户当前聊天新版“五、产品经营思路：规格分层与价格梯度设计”截图。单页，1672×941。规格1、2、4、6、8、10、12瓶；单瓶价格¥458、¥447、¥435、¥424、¥412、¥401、¥389；总价¥458、¥894、¥1,740、¥2,544、¥3,296、¥4,010、¥4,668按图保留，逐项核对数量×单价=总价。左侧为8行3列原生PowerPoint表格，12瓶行保留推荐标记。右侧为独立可编辑的入门体验、主力规格、长期管理说明。文字不构成对经营效果的独立验证。雪山、瓶子和图标为近似重绘；右上模糊白色装饰字未作为正文复制。不含整页图片背景，没有新增原图未出现的折线图，也不修改前一版文件。表格中价格与总价可分别编辑，手动改价后需同步总价。';
const raw=path.join(build,'raw.pptx'),styled=path.join(build,'table-styled.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
const styleTable=String.raw`
import sys,json
from zipfile import ZipFile
from lxml import etree as E
A='http://schemas.openxmlformats.org/drawingml/2006/main'
tables=0;locks=0
with ZipFile(sys.argv[1]) as src,ZipFile(sys.argv[2],'w') as dst:
 for info in src.infolist():
  data=src.read(info.filename)
  if info.filename.startswith('ppt/slides/slide') and info.filename.endswith('.xml'):
   root=E.fromstring(data)
   for e in root.iter():
    if e.get('noGrp')=='1':e.attrib.pop('noGrp');locks+=1
   for table in root.findall('.//{%s}tbl'%A):
    rows=table.findall('{%s}tr'%A)
    for ri,row in enumerate(rows):
     for cell in row.findall('{%s}tc'%A):
      props=cell.find('{%s}tcPr'%A)
      if props is None:props=E.SubElement(cell,'{%s}tcPr'%A)
      for key in ['lnL','lnR','lnT','lnB']:
       for e in props.findall('{%s}%s'%(A,key)):props.remove(e)
       visible=key=='lnB' and ri not in [0,6,7]
       ln=E.Element('{%s}%s'%(A,key),w='11430' if visible else '0')
       if visible:E.SubElement(E.SubElement(ln,'{%s}solidFill'%A),'{%s}srgbClr'%A,val='D4E3F0')
       else:E.SubElement(ln,'{%s}noFill'%A)
       props.insert(0,ln)
    tables+=1
   data=E.tostring(root,encoding='UTF-8',xml_declaration=True)
  dst.writestr(info,data)
assert tables==1,tables
print(json.dumps({'native_tables':tables,'removed_group_locks':locks}))
`;
await fs.writeFile(path.join(build,'table-styling.json'),execFileSync(process.env.RUNTIME_PYTHON,['-c',styleTable,raw,styled],{encoding:'utf8'}));
await fs.writeFile(path.join(build,'normalization.json'),execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',styled,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-tables','1'],{encoding:'utf8'}));
const finalPath=path.join(output,'pricing-table-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit','--require-native-table-slide','1'],explicitTotalSlideCount:1,requiredNativeTableOwnerSlides:[1],fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),table:{rows:8,columns:3},quantities,prices,totals},null,2));
