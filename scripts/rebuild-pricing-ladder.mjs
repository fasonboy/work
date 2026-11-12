// Rebuild the pricing ladder screenshot with a native line chart and two native tables.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/pricing-ladder',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation,applyPresentationChartFont}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
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
const counts=[1,2,4,6,8,10,12],unitPrices=[458,447,435,424,412,401,389],sourceTotals=[458,894,1740,2544,3296,4010,4668];
const totals=counts.map((n,i)=>n*unitPrices[i]);assert.deepEqual(totals,sourceTotals);
slide=deck.slides.add();slide.background.fill=grad('#ffffff','#eef7ff',60);
const paths={
 tag:'<path d="M5 38L35 7h27v28L31 65z"/><circle cx="51" cy="19" r="4"/>',
 layers:'<path d="M5 21L34 5l29 16-29 16zM6 34l28 16 28-16M6 47l28 16 28-16"/>',
 target:'<circle cx="31" cy="36" r="27"/><circle cx="31" cy="36" r="17"/><circle cx="31" cy="36" r="7"/><path d="M31 36L60 7M53 1l-2 13 13-2"/>',
 user:'<circle cx="30" cy="16" r="11"/><path d="M9 61V49c0-23 42-23 42 0v12M52 12l5 3M58 27h6"/>',
 cart:'<path d="M5 5h10l9 39h32l9-28H19M24 53h31"/><circle cx="28" cy="60" r="4"/><circle cx="53" cy="60" r="4"/>',
 trend:'<path d="M6 32L24 17l13 10L63 4M48 4h15v15M7 64V47h11v17zM29 64V37h11v27zM51 64V26h11v38z"/>',
 gift:'<path d="M5 24h58v12H5zM10 36v28h48V36M34 23v41M34 22C6 25 10-9 26 5l8 17c28 3 24-31 8-17z"/>',
 flask:'<path d="M24 4h20M28 4v23L8 58c-3 5 1 7 5 7h43c5 0 7-3 4-7L40 27V4M20 44h28"/>',
 coins:'<ellipse cx="46" cy="15" rx="17" ry="7"/><path d="M29 15v28c0 10 34 10 34 0V15M29 26c0 10 34 10 34 0M29 36c0 10 34 10 34 0"/><ellipse cx="21" cy="40" rx="17" ry="7"/><path d="M4 40v16c0 10 34 10 34 0V40M4 48c0 10 34 10 34 0"/>',
 cycle:'<path d="M9 27C11 3 48 0 60 20M60 7v14H46M60 40C56 65 20 68 8 48M8 61V47h14"/>',
 bottle:'<path d="M22 14V5h25v9M19 14h31v7H19zM23 21c0 8-9 8-9 16v27c0 4 3 6 6 6h27c4 0 7-2 7-6V37c0-8-10-8-10-16M15 38h38"/>',
};
function pict(kind,x,y,size,color=NAVY){slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="84" height="84" viewBox="-7 -4 84 84"><g fill="none" stroke="${color}" stroke-width="3.7" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} icon`,position:{left:x,top:y,width:size,height:size}});}
// Native polar facets approximate the decorative snow photograph.
box('polar-water',964,106,708,25,grad('#e5f6ff','#ffffff/0',90));
poly('polar-back',[[937,108],[1086,65],[1153,83],[1240,51],[1290,71],[1365,23],[1421,66],[1487,52],[1608,105],[1672,96],[1672,128]],'#d4effe');
poly('polar-front',[[1063,117],[1200,77],[1261,97],[1365,23],[1413,81],[1461,66],[1554,105],[1672,115],[1672,127]],'#8ec8ed');
for(const [i,p] of [[1,[[1365,23],[1312,105],[1349,73],[1360,111],[1391,52]]],[2,[[1365,23],[1413,81],[1461,66],[1405,111]]],[3,[[1240,51],[1189,113],[1224,88],[1258,115],[1290,71]]]])poly(`polar-snow-${i}`,p,'#ffffff');
box('title-accent',37,15,12,53,RED,3);txt('page-title','五、产品经营思路：规格分层与价格梯度设计',68,7,1267,65,43,NAVY,true);
txt('page-subtitle','通过价格梯度与规格分层，实现新品试用、客单提升与长期复购协同',69,66,1174,42,25,NAVY);
txt('report-brand','南极磷虾油市场调研',1438,10,220,36,23,NAVY,true);box('report-brand-rule',1439,47,141,4,NAVY,2);box('report-brand-accent',1585,47,52,4,RED,2);

function panel(name,x,y,w,h,title,kind){
 box(`${name}-shadow`,x,y+5,w,h,'#cadaee/65',27);box(name,x,y,w,h,'#ffffff',27,'rect','#c5dcef',1.2);
 curve(`${name}-heading-bg`,x,y,w-56,63,[M(27,0),L(w-56,0),L(w-99,63),L(0,63),L(0,27),B(0,10,10,0,27,0),Z()],grad('#0c4b85','#0871b9'));
 poly(`${name}-header-blue-accent`,[[x+w-89,y],[x+w-60,y],[x+w-102,y+63],[x+w-131,y+63]],'#b4dfff/60');
 poly(`${name}-header-red-accent`,[[x+w-74,y],[x+w-67,y],[x+w-109,y+63],[x+w-116,y+63]],'#ed6c70');
 pict(kind,x+37,y+12,46,'#ffffff');txt(`${name}-heading`,title,x+110,y+5,w-156,53,33,'#ffffff',true);
}
panel('pricing-panel',41,131,858,654,'规格价格体系','tag');
box('pricing-principle-bg',205,220,578,47,'#fff2f4',17);box('pricing-principle-dot',223,237,13,13,RED,0,'ellipse');
txt('pricing-principle','单瓶价格随购买规格递减，形成清晰的价格梯度。',248,222,523,43,23,NAVY,true);

const priceChart=slide.charts.add('line',{
 position:{left:46,top:276,width:850,height:221},titlePlacement:'none',categories:counts.map(n=>`${n}瓶`),
 series:[{name:'单瓶价格（元/瓶）',values:unitPrices,line:{fill:'#f03d43',width:3.3},marker:{symbol:'circle',size:12,fill:'#ffffff'},dataLabelOverrides:[{idx:0,showValue:true,fill:'#ee202b',line:{fill:'#ee202b',width:1},textStyle:{typeface:FONT,fontSize:25,fill:'#ffffff',bold:true},position:'top'},{idx:6,showValue:true,fill:'#ee202b',line:{fill:'#ee202b',width:1},textStyle:{typeface:FONT,fontSize:25,fill:'#ffffff',bold:true},position:'top'}]}],
 lineOptions:{grouping:'standard',smooth:false},hasLegend:false,
 xAxis:{visible:false,line:{fill:'none',width:0},tickLabelPosition:'none',majorGridlines:null,minorGridlines:null},
 yAxis:{visible:false,min:370,max:500,line:{fill:'none',width:0},tickLabelPosition:'none',majorGridlines:null,minorGridlines:null},
 dataLabels:{showValue:true,position:'top',numberFormatCode:'0"元/瓶"',textStyle:{typeface:FONT,fontSize:23,fill:NAVY,bold:true}},chartFill:'none',chartLine:{fill:'none',width:0},plotAreaFill:grad('#ffffff','#fff5f6',90),plotAreaLine:{fill:'none',width:0},
});applyPresentationChartFont(priceChart,{fontFamily:FONT});

// One real table keeps the aligned seven-tier price data editable as a table.
const table=slide.tables.add({rows:4,columns:7,left:62,top:509,width:818,height:252,columnWidths:Array(7).fill(818/7),values:[counts.map(n=>`${n}瓶`),counts.map(()=>''),unitPrices.map(p=>`${p}元/瓶`),totals.map(t=>`总价\n${t}元`)]});
for(let r=0;r<4;r++){
 table.rows[r].height=[42,82,54,74][r];
 for(let c=0;c<7;c++){
  const cell=table.getCell(r,c);cell.fill=r===0?grad('#0875bb','#084a82'):r===3?'#fff1f3':'#ffffff';
  cell.text.style={typeface:FONT,fontSize:r===0?27:r===2?22.5:26,color:r===0?'#ffffff':r===3||r===2&&(c===0||c===6)?'#df1725':NAVY,bold:true,autoFit:'none',wrap:'none',alignment:'center',verticalAlignment:'middle'};
  table.cells.block({row:r,column:c,rowCount:1,columnCount:1}).assign({margins:{left:3,right:3,top:3,bottom:3}});
  if(r===3)cell.value=[[{run:'总价',style:{typeface:FONT,fontSize:'18px',color:NAVY,bold:true}}],[{run:`${totals[c]}元`,style:{typeface:FONT,fontSize:'27px',color:'#df1725',bold:true}}]];
 }
}
table.borders.assign({color:'#ffffff',fill:'#ffffff',width:5,style:'solid'});
for(const [i,n] of counts.entries()){
 const x=62+i*(818/7);pict('bottle',x+27,563,65,'#40a4e2');txt(`bottle-quantity-${i+1}`,`×${n}`,x+37,589,48,32,21,NAVY,true,{alignment:'center'});
}

panel('specification-panel',920,131,722,337,'规格设置思路','layers');
const specRows=[
 {kind:'user',role:'引流体验规格',spec:'1瓶',detail:'降低首次购买门槛，适合试用与新品体验。'},
 {kind:'cart',role:'承接转化规格',spec:'2瓶、4瓶',detail:'适合轻度需求与初步复购，兼顾价格接受度。'},
 {kind:'trend',role:'主推提升规格',spec:'6瓶、8瓶',detail:'兼顾性价比与客单价，适合作为主推成交规格。'},
 {kind:'gift',role:'高客单利润规格',spec:'10瓶、12瓶',detail:'面向家庭囤货、长期服用及送礼需求，提升大单占比。'},
];
const specTable=slide.tables.add({rows:4,columns:4,left:936,top:193,width:690,height:264,columnWidths:[89,173,154,274],values:specRows.map(s=>['',s.role,s.spec,s.detail])});
for(let r=0;r<4;r++){
 specTable.rows[r].height=66;
 for(let c=0;c<4;c++){
  const cell=specTable.getCell(r,c);cell.fill=c===2?'#ffe5e8':r%2?'#fff3f5':'#edf7ff';
  cell.text.style={typeface:FONT,fontSize:c===1?(r===3?22.5:24):c===2?22:18,color:c===2?'#d61b27':NAVY,bold:true,autoFit:'none',wrap:c<3?'none':'square',alignment:c<3?'center':'left',verticalAlignment:'middle',lineSpacing:1.03};
  specTable.cells.block({row:r,column:c,rowCount:1,columnCount:1}).assign({margins:{left:c===3?12:4,right:c===3?8:4,top:4,bottom:4}});
 }
 pict(specRows[r].kind,956,202+r*66,48,r%2?'#e32635':'#167fbd');
}
specTable.borders.assign({color:'#ffffff',fill:'#ffffff',width:5,style:'solid'});

panel('conversion-panel',920,483,722,302,'经营逻辑与转化想法','target');
const logic=[
 {x:937,y:546,w:333,h:103,kind:'flask',color:'#1586c8',fill:'#edf8ff',title:'先试用，再放大',text:'以1瓶建立体验，降低新客决策成本。'},
 {x:1287,y:546,w:339,h:103,kind:'layers',color:'#e92835',fill:'#fff1f4',title:'中间规格承接',text:'2–8瓶构成阶梯选择，减少用户因跨度过大而流失。'},
 {x:937,y:657,w:333,h:104,kind:'coins',color:'#e92835',fill:'#fff1f4',title:'大规格做利润',text:'10–12瓶提升客单价与利润空间。'},
 {x:1287,y:657,w:339,h:104,kind:'cycle',color:'#1586c8',fill:'#edf8ff',title:'长期管理导向',text:'通过规格设计引导消费者从短期尝试走向周期性复购。'},
];
for(const [i,c] of logic.entries()){
 box(`logic-card-${i+1}`,c.x,c.y,c.w,c.h,c.fill,18);pict(c.kind,c.x+18,c.y+24,49,c.color);box(`logic-card-${i+1}-divider`,c.x+82,c.y+15,1.5,c.h-30,c.color);
 txt(`logic-card-${i+1}-heading`,c.title,c.x+102,c.y+1,c.w-114,39,25,NAVY,true);
 txt(`logic-card-${i+1}-detail`,c.text,c.x+102,c.y+38,c.w-115,c.h-41,18.5,NAVY,false,{wrap:'square',lineSpacing:1.05});
}
box('conclusion-panel',41,808,1595,74,grad('#ffffff','#f1f8ff'),24,'rect','#e0eefb');
curve('conclusion-label-bg',41,808,331,74,[M(26,0),L(331,0),L(280,74),L(26,74),B(8,74,0,65,0,49),L(0,26),B(0,9,9,0,26,0),Z()],grad('#126eae','#044b80'));
pict('target',79,822,52,'#ffffff');txt('conclusion-label','核心方向',153,816,171,57,34,'#ffffff',true);box('conclusion-label-divider',314,827,1.5,34,'#e1f4ff');
txt('conclusion','以规格分层带动价格梯度，以主推规格提升转化，以大规格放大客单与复购。',408,817,1180,56,31,NAVY,true);
curve('footer-light',0,889,W,52,[M(0,0),B(440,76,1170,15,W,10),L(W,52),L(0,52),Z()],'#bce5fc');
curve('footer-blue',0,904,W,37,[M(0,0),B(600,56,1130,22,W,23),L(W,37),L(0,37),Z()],grad('#159bea','#78cbf8'));
slide.speakerNotes.text='来源：用户当前聊天“五、产品经营思路：规格分层与价格梯度设计”截图，1672×941。规格1/2/4/6/8/10/12瓶，单瓶458/447/435/424/412/401/389元，总价458/894/1740/2544/3296/4010/4668元按图保留，逐项验证瓶数乘单价等于总价。价格走势图为原生折线图，数据标签来自图表数值，内嵌依据截图新建的Excel数据快照，不声称保留原始公式或工作簿。规格价格区重建为4行7列原生表格（含瓶图标空行），规格说明为4行4列原生表格；图标另嵌并保留可编辑数量标签。表格、图表和总价是独立对象，手动改价时需同步更新；重复运行生成脚本会从统一数据数组计算总价。右侧经营逻辑按图转录，不新增利润、转化或健康效果保证。雪山和图标近似重绘，无整页截图背景。';
const raw=path.join(build,'raw.pptx'),styled=path.join(build,'table-styled.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
// Style real table borders explicitly; keep all cells and their table operations.
const styleTables=String.raw`
import sys,json
from zipfile import ZipFile
from lxml import etree as E
A='http://schemas.openxmlformats.org/drawingml/2006/main'
C='http://schemas.openxmlformats.org/drawingml/2006/chart'
count=0;locks=0;markers=0
with ZipFile(sys.argv[1]) as src,ZipFile(sys.argv[2],'w') as dst:
 for info in src.infolist():
  data=src.read(info.filename)
  if info.filename.startswith('ppt/slides/slide') and info.filename.endswith('.xml'):
   root=E.fromstring(data)
   for e in root.iter():
    if e.get('noGrp')=='1':e.attrib.pop('noGrp');locks+=1
   for ti,table in enumerate(root.findall('.//{%s}tbl'%A)):
    for row in table.findall('{%s}tr'%A):
     for ci,cell in enumerate(row.findall('{%s}tc'%A)):
      props=cell.find('{%s}tcPr'%A)
      if props is None:props=E.SubElement(cell,'{%s}tcPr'%A)
      for key in ['lnL','lnR','lnT','lnB']:
       for e in props.findall('{%s}%s'%(A,key)):props.remove(e)
       ln=E.Element('{%s}%s'%(A,key),w=str(57150 if ti==0 and key in ['lnL','lnR'] or ti==1 and key in ['lnT','lnB'] else 19050 if ti==0 else 0))
       if ln.get('w')=='0':E.SubElement(ln,'{%s}noFill'%A)
       else:E.SubElement(E.SubElement(ln,'{%s}solidFill'%A),'{%s}srgbClr'%A,val='FFFFFF')
       props.insert(0,ln)
    count+=1
   data=E.tostring(root,encoding='UTF-8',xml_declaration=True)
  elif '/charts/' in info.filename and info.filename.endswith('.xml'):
   root=E.fromstring(data)
   for marker in root.findall('.//{%s}marker'%C):
    props=marker.find('{%s}spPr'%C)
    if props is None:props=E.SubElement(marker,'{%s}spPr'%C)
    for ln in props.findall('{%s}ln'%A):props.remove(ln)
    ln=E.SubElement(props,'{%s}ln'%A,w='23813')
    E.SubElement(E.SubElement(ln,'{%s}solidFill'%A),'{%s}srgbClr'%A,val='E8202C')
    markers+=1
   # Point overrides must explicitly hide category/series text in LibreOffice,
   # which otherwise does not inherit all collection-level false flags.
   for label in root.findall('.//{%s}dLbl'%C):
    for key in ['numFmt','dLblPos','showLegendKey','showVal','showCatName','showSerName','showPercent','showBubbleSize','showLeaderLines']:
     for e in label.findall('{%s}%s'%(C,key)):label.remove(e)
    label.insert(1,E.Element('{%s}numFmt'%C,formatCode='0"元/瓶"',sourceLinked='0'))
    E.SubElement(label,'{%s}dLblPos'%C,val='t')
    for key,value in [('showLegendKey','0'),('showVal','1'),('showCatName','0'),('showSerName','0'),('showPercent','0'),('showBubbleSize','0'),('showLeaderLines','0')]:E.SubElement(label,'{%s}%s'%(C,key),val=value)
   for labels in root.findall('.//{%s}dLbls'%C):
    for e in labels.findall('{%s}dLblPos'%C):labels.remove(e)
    i=next((i for i,e in enumerate(labels) if E.QName(e).localname.startswith('show')),len(labels))
    labels.insert(i,E.Element('{%s}dLblPos'%C,val='t'))
   for text in root.findall('.//{%s}txPr'%C):
    body=text.find('{%s}bodyPr'%A)
    if body is not None:body.set('wrap','none')
   data=E.tostring(root,encoding='UTF-8',xml_declaration=True)
  dst.writestr(info,data)
assert count==2,count
assert markers==1,markers
print(json.dumps({'native_tables':count,'native_marker_styles':markers,'removed_group_locks':locks}))
`;
await fs.writeFile(path.join(build,'table-styling.json'),execFileSync(process.env.RUNTIME_PYTHON,['-c',styleTables,raw,styled],{encoding:'utf8'}));
await fs.writeFile(path.join(build,'normalization.json'),execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',styled,'--output',portable,'--sharp-module',req.resolve('sharp'),'--min-native-tables','2'],{encoding:'utf8'}));
const finalPath=path.join(output,'pricing-ladder-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-heading-fit','--require-native-table-slide','1'],explicitTotalSlideCount:1,requiredNativeTableOwnerSlides:[1],requiredNativeChartOwnerSlides:[1],materializeLiteralChartWorkbooks:true,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),counts,unitPrices,totals,nativeTables:2,nativeCharts:1},null,2));
