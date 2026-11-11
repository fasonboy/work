// Rebuild the krill oil consumer segments as editable cards and native bullet lists.
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';

const project=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const argv=process.argv.slice(2);
const opt=k=>argv.indexOf(k)<0?undefined:argv[argv.indexOf(k)+1];
const work=path.resolve(opt('--work-dir')||path.join(project,'.local/krill-consumers',`run-${Date.now()}`));
const build=path.join(work,'build'),output=path.join(work,'output');
await fs.mkdir(build,{recursive:true}); await fs.mkdir(output,{recursive:true});
assert(process.env.RUNTIME_NODE_MODULES && process.env.PRESENTATIONS_SKILL_DIR,'First source scripts/ppt-runtime.sh');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES,'__jkn_brand__.cjs'));
const {Presentation,PresentationFile,FileBlob}=await import(pathToFileURL(req.resolve('@oai/artifact-tool')).href);
const sharp=req('sharp');
const skill=process.env.PRESENTATIONS_SKILL_DIR;
const {finalizePresentation,makeNativeBulletParagraphs}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const W=1672,H=941,FONT='Noto Sans CJK SC',NAVY='#061f50';
if(process.env.PPT_OPERATION_MARKED!=='1') execFileSync(process.env.RUNTIME_NODE,[path.join(skill,'container_tools/mark_artifact_operation_started.mjs'),'--operation-kind','create','--expected-output-count','1','--output-format','pptx'],{stdio:'inherit'});
const deck=Presentation.create({slideSize:{width:W,height:H}}),slide=deck.slides.add();slide.background.fill='#ffffff';
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
function personIcon(kind,x,y,size) {
  const paths={
    older:'<circle cx="34" cy="43" r="16"/><path d="M17 36c-2-22 34-27 35 0M25 59v9l9 9 9-9v-9M3 112V89c0-24 62-24 62 0v23M14 87v25M54 87v25"/><circle cx="27" cy="42" r="5"/><circle cx="41" cy="42" r="5"/><path d="M32 42h4M18 39h4M47 39h4M28 53h12"/><circle cx="101" cy="16" r="9"/><path d="M82 42c-5-27 39-28 38 0M83 41c2 39 34 39 35 0M81 42c13-1 17-12 18-17 4 8 11 13 22 17M93 66v8M108 66v8M74 112V91c0-26 55-26 55 0v21M84 91v21M118 91v21"/>',
    adults:'<path d="M17 37c-6-35 44-36 37 0M18 35c8 0 10-10 14-14 5 7 12 10 22 10M18 37c-3 35 39 35 36 0M26 63v9l9 8 9-8v-9M3 112V91c0-25 62-25 62 0v21M14 91v21M54 91v21"/><path d="M80 43c-9-36 46-39 41-1l3 29M82 42c12 0 17-10 18-17 3 7 11 15 22 16M82 43c-1 35 38 35 38 0M92 66v8M109 66v8M74 112V93c0-27 55-27 55 0v19M84 94v18M118 94v18"/>',
    family:'<path d="M7 58L70 7l63 51h-14v56H21V58z"/><circle cx="49" cy="75" r="12"/><circle cx="93" cy="75" r="12"/><circle cx="71" cy="94" r="11"/><path d="M32 112v-9c0-23 30-23 30-7M110 112v-9c0-23-30-23-30-7M54 115v-7c0-18 34-18 34 0v7"/>',
    shield:'<path d="M70 8l44 19v31c0 31-23 47-44 59-21-12-44-28-44-59V27zM48 61l16 17 34-36"/>',
  };
  slide.images.add({svg:`<svg xmlns="http://www.w3.org/2000/svg" width="144" height="124" viewBox="-3 0 142 124"><g fill="none" stroke="#0c5f9f" stroke-width="4.1" stroke-linecap="round" stroke-linejoin="round">${paths[kind]}</g></svg>`,alt:`${kind} consumer illustration`,position:{left:x,top:y,width:size,height:size*.86}});
}
// Krill decoration is an editable illustration rather than a source photograph.
curve('krill-tail-fan',31,99,36,32,[M(29,0),L(0,32),L(9,16),L(9,25),L(17,8),L(29,0),Z()],'#b9c9da/60');
for(let i=0;i<7;i++) {
  const x=52+i*10,y=99-i*3;
  const seg=box(`krill-body-segment-${i}`,x,y,17+i*.6,13,'#ffe4d1',0,'ellipse','#ec946b',.9);seg.position={left:x,top:y,width:17+i*.6,height:13,rotation:-23};
  curve(`krill-leg-${i}`,x+5,y+7,18,28,[M(0,0),B(19,4,3,20,10,28)],'none','#aeb9ca/80',1.2);
}
const body=box('krill-carapace',95,81,45,17,grad('#fff0e2','#f77536'),0,'ellipse','#ef8256',1);body.position={left:95,top:81,width:45,height:17,rotation:-17};
box('krill-head',130,76,17,12,'#ff9c58',0,'ellipse');box('krill-eye',139,76,4,4,'#101b2c',0,'ellipse');
curve('krill-dorsal-line',59,76,80,29,[M(0,29),B(27,18,58,3,80,0)],'none','#ef6834',2.2);
for(const [i,endY] of [[1,0],[2,13],[3,33]])curve(`krill-antenna-${i}`,142,59,45,38,[M(0,18),B(14,11,28,endY,45,endY)],'none','#efb491',.9);
txt('page-title','南极磷虾油核心消费群体',196,34,1385,104,72,NAVY,true,{italic:true});
txt('page-subtitle','健康管理需求持续扩大，消费人群由中老年向年轻化、家庭化延伸',49,141,1500,49,26.5,'#303030');

const segments=[
 {x:9,w:409,title:'中老年健康管理人群',kind:'older',pillX:56,pillW:318,age:'45岁以上 | 重点消费群体',bodyX:42,bodyW:355,after:31,items:['关注日常营养补充与健康维护','健康管理意识较强，\n具有长期消费潜力']},
 {x:434,w:396,title:'都市中青年人群',kind:'adults',pillX:470,pillW:325,age:'25—45岁 | 潜力增长群体',bodyX:480,bodyW:326,after:22,items:['工作压力较大，注重日常\n营养与健康管理','对新型营养补充产品\n接受度较高']},
 {x:846,w:395,title:'家庭健康采购人群',kind:'family',pillX:882,pillW:333,age:'30—55岁 | 家庭消费决策者',bodyX:890,bodyW:329,after:22,items:['为父母、配偶购买营养\n健康产品','家庭常备、节日送礼等需求\n形成多元消费场景']},
 {x:1257,w:406,title:'品质健康消费人群',kind:'shield',pillX:1294,pillW:336,age:'30—55岁 | 高价值潜力客群',bodyX:1306,bodyW:332,after:22,items:['已有鱼油、Omega-3等\n营养补充产品消费习惯','重视产品原料、营养成分、\n品牌及品质认证']},
];
for(const [i,s] of segments.entries()) {
  box(`segment-${i+1}-panel`,s.x,229,s.w,533,grad('#ffffff','#fcfeff'),21,'rect','#c5dfff',1.2);
  box(`segment-${i+1}-header`,s.x+1,230,s.w-2,72,grad('#fafeff','#e5f3ff'),20);
  box(`segment-${i+1}-header-lower`,s.x+1,270,s.w-2,32,grad('#fafeff','#e5f3ff'));
  curve(`segment-${i+1}-number-bg`,s.x+14,242,104,51,[M(16,0),L(104,0),L(87,51),L(0,51),Z()],grad('#00a9e6','#04539a',40));
  txt(`segment-${i+1}-number`,String(i+1).padStart(2,'0'),s.x+44,243,62,48,32,'#ffffff',true);
  txt(`segment-${i+1}-title`,s.title,s.x+124,245,s.w-136,54,29.5,NAVY,true);
  const cx=s.x+s.w/2;
  box(`segment-${i+1}-icon-bg`,cx-71,313,142,138,grad('#f5fbff','#e7f4ff'),0,'ellipse','#edf7ff');
  personIcon(s.kind,cx-69,330,138);
  box(`segment-${i+1}-pill`,s.pillX,463,s.pillW,48,grad('#e3f3ff','#d9edff'),24);
  txt(`segment-${i+1}-age`,s.age,s.pillX+5,468,s.pillW-10,39,i>1?24.4:25,NAVY,true,{alignment:'center'});
  box(`segment-${i+1}-divider`,s.x+27,532,s.w-53,1.3,grad('#bad6ff','#eff7ff',0));
  const body=box(`segment-${i+1}-bullet-body`,s.bodyX,557,s.bodyW,185);
  body.text=makeNativeBulletParagraphs(s.items.map(item=>item.replaceAll('\n','')),{marginLeftPoints:28.5,hangingPoints:28.5,spaceAfterPoints:s.after,bulletCharacter:'■'});
  body.text.style={typeface:FONT,fontSize:24,color:'#151515',autoFit:'none',wrap:'square',insets:0,verticalAlignment:'top',lineSpacing:1.18};
}
box('insight-banner',9,815,1655,91,grad('#ffffff','#e7f4ff'),20);
curve('insight-label-bg',9,816,265,88,[M(27,0),L(265,0),L(237,88),L(0,88),Z()],grad('#0a86c4','#03528c',50));
txt('insight-heading','人群洞察',53,829,211,62,43,'#ffffff',true);
txt('insight-text','南极磷虾油消费场景覆盖个人健康管理、家庭营养补充及健康礼赠，具备多层次市场拓展潜力。',300,829,1177,64,27.5,NAVY,true);
for(const [i,x,y,w,h,rotation] of [[1,1493,854,95,35,-15],[2,1557,832,84,33,16]]) {
  const shadow=box(`capsule-${i}-shadow`,x+2,y+h-1,w,14,'#cda1aa/25',0,'ellipse');shadow.position={left:x+2,top:y+h-1,width:w,height:14,rotation};
  const cap=box(`capsule-${i}`,x,y,w,h,grad('#ffb9a6','#e51a0b',90),0,'ellipse','#d52b21',1.5);cap.position={left:x,top:y,width:w,height:h,rotation};
  const shine=box(`capsule-${i}-shine`,x+9,y+3,w-21,8,'#fff7f0/80',0,'ellipse');shine.position={left:x+9,top:y+3,width:w-21,height:8,rotation};
}
slide.speakerNotes.text='来源：用户当前聊天“南极磷虾油核心消费群体”截图。四类人群、年龄范围、两条说明和人群洞察按图保留，未独立验证消费画像或营养/市场结论。每组说明使用真正的项目符号段落，换行可能与截图略有差异。人物及家庭/盾牌图标为SVG，磷虾和胶囊为原生矢量近似重绘，不是原始照片或图形文件。整页没有铺设参考截图背景。';
const raw=path.join(build,'raw.pptx'),colored=path.join(build,'native-bullets.pptx'),portable=path.join(build,'portable.pptx');
await(await PresentationFile.exportPptx(deck)).save(raw);
// Add standard DrawingML bullet color/size to generated native paragraphs.
// Preserve raw.pptx and all original assets; this changes formatting only.
const colorBullets=String.raw`
import sys,json,copy
from zipfile import ZipFile
from lxml import etree
A='http://schemas.openxmlformats.org/drawingml/2006/main'
P='http://schemas.openxmlformats.org/presentationml/2006/main'
line_breaks=json.loads(sys.argv[3])
removed_group_locks=0
restored_bullets=0
with ZipFile(sys.argv[1]) as src, ZipFile(sys.argv[2],'w') as dst:
    for info in src.infolist():
        data=src.read(info.filename)
        if info.filename.startswith('ppt/slides/slide') and info.filename.endswith('.xml'):
            root=etree.fromstring(data)
            # Namespace declarations can precede noGrp after XML serialization;
            # clear the real attribute without relying on textual attribute order.
            for element in root.iter():
                if element.get('noGrp')=='1':
                    element.attrib.pop('noGrp')
                    removed_group_locks+=1
            for shape in root.findall('.//{%s}sp'%P):
                identity=shape.find('{%s}nvSpPr/{%s}cNvPr'%(P,P))
                if identity is None or 'bullet-body' not in identity.get('name',''): continue
                segment_index=int(identity.get('name').split('-')[1])-1
                for paragraph_index,paragraph in enumerate(shape.findall('.//{%s}p'%A)):
                    p=paragraph.find('{%s}pPr'%A)
                    if p is None: continue
                    # This runtime exports the custom square as buNone. Restore a
                    # standard native DrawingML marker; never draw a fake text bullet.
                    for name in ['buClr','buClrTx','buSzPct','buSzPts','buSzTx','buNone','buAutoNum','buBlip','buChar']:
                        for e in p.findall('{%s}%s'%(A,name)): p.remove(e)
                    clr=etree.Element('{%s}buClr'%A)
                    etree.SubElement(clr,'{%s}srgbClr'%A,val='00A2E7')
                    sz=etree.Element('{%s}buSzPct'%A,val='100000')
                    marker=etree.Element('{%s}buChar'%A,char='■')
                    restored_bullets+=1
                    before=next((i for i,e in enumerate(p) if etree.QName(e).localname in ['buFont','buFontTx','tabLst','defRPr','extLst']),len(p))
                    p.insert(before,clr);p.insert(before+1,sz)
                    marker_at=next((i for i,e in enumerate(p) if etree.QName(e).localname in ['tabLst','defRPr','extLst']),len(p))
                    p.insert(marker_at,marker)
                    runs=paragraph.findall('{%s}r'%A)
                    run_style=copy.deepcopy(runs[0].find('{%s}rPr'%A)) if runs else None
                    for run in runs: paragraph.remove(run)
                    for line_index,line in enumerate(line_breaks[segment_index][paragraph_index].split('\n')):
                        if line_index:
                            br=etree.SubElement(paragraph,'{%s}br'%A)
                            if run_style is not None: br.append(copy.deepcopy(run_style))
                        run=etree.SubElement(paragraph,'{%s}r'%A)
                        if run_style is not None: run.append(copy.deepcopy(run_style))
                        etree.SubElement(run,'{%s}t'%A).text=line
            data=etree.tostring(root,encoding='UTF-8',xml_declaration=True)
        dst.writestr(info,data)
print(json.dumps({'removedGroupLocks':removed_group_locks,'nativeBulletItems':restored_bullets}))
`;
const nativeFormatting=execFileSync(process.env.RUNTIME_PYTHON,['-c',colorBullets,raw,colored,JSON.stringify(segments.map(s=>s.items))],{encoding:'utf8'});
await fs.writeFile(path.join(build,'native-formatting.json'),nativeFormatting);
const normal=execFileSync(process.env.RUNTIME_NODE,[path.join(project,'.agents/skills/image-to-editable-ppt/scripts/normalize_portable_pptx.mjs'),'--input',colored,'--output',portable,'--sharp-module',req.resolve('sharp')],{encoding:'utf8'});
await fs.writeFile(path.join(build,'normalization.json'),normal);
const finalPath=path.join(output,'krill-consumers-editable.pptx');
await finalizePresentation({workspaceDir:work,candidatePath:portable,finalPath,pythonExecutable:process.env.RUNTIME_PYTHON,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu',`${W*9525},${H*9525}`,'--validate-bullet-geometry','--validate-heading-fit'],explicitTotalSlideCount:1,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
const imported=await PresentationFile.importPptx(await FileBlob.load(finalPath));
const png=await imported.export({slide:imported.slides.items[0],format:'png',scale:1});
await fs.writeFile(path.join(build,'preview.png'),Buffer.from(await png.arrayBuffer()));
console.log(JSON.stringify({finalPath,preview:path.join(build,'preview.png'),segments:4,nativeBulletItems:8},null,2));
