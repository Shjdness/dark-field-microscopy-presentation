import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { pathToFileURL } from "node:url";

const WORKSPACE = "C:/Users/Shjdshy/Desktop/AllDeskWork/暗场显微镜PPT_DDL100818pm";
const SKILL_DIR = "C:/Users/Shjdshy/.codex/plugins/cache/openai-primary-runtime/presentations/26.915.20218/skills/presentations";
const RUNTIME_PYTHON = "C:/Users/Shjdshy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe";
const SOURCE = path.join(WORKSPACE, "C已完成/暗场显微镜_课堂汇报_样品版_v2.pptx");
const IMAGE = path.join(WORKSPACE, "B工作区/PPT构建/assets/暗场显微镜_六部件剖面示意_v1.png");
const BUILD = path.join(WORKSPACE, "B工作区/PPT构建/structure-v3");
const FINAL = path.join(WORKSPACE, "C已完成/暗场显微镜_课堂汇报_结构讲解优化版_v3.pptx");

const { importRuntimeModule } = await import(pathToFileURL(path.join(SKILL_DIR, "container_tools/runtime_helpers.mjs")).href);
const { FileBlob, PresentationFile } = await importRuntimeModule("@oai/artifact-tool");
const deck = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const slide = deck.slides.getItem(6);
const imageBytes = await fs.readFile(IMAGE);

const C = {
  navy: "#173B63", blue: "#2C73A8", cyan: "#64B6CE", pale: "#EAF4F8",
  ink: "#172536", gray: "#647486", line: "#CBDCE5", white: "#FFFFFF",
  orange: "#F2A541", green: "#3D8B7D", slate: "#7A8795",
};
const FONT_CN = "Microsoft YaHei";
const FONT_EN = "Arial";

function addText(text, x, y, w, h, size=20, color=C.ink, bold=false, align="left", family=FONT_CN) {
  const s = slide.shapes.add({ geometry:"textbox", position:{left:x,top:y,width:w,height:h}, fill:"none", line:{fill:"none",width:0} });
  s.text = text;
  s.text.style = { typeface:family, fontSize:size, color, bold, alignment:align, verticalAlignment:"middle", autoFit:"shrinkText", lineSpacing:1.04 };
  return s;
}
function rect(x,y,w,h,fill,radius=0,lineColor="none",lineWidth=0) {
  return slide.shapes.add({ geometry:radius?"roundRect":"rect", position:{left:x,top:y,width:w,height:h}, fill, line:{fill:lineColor,width:lineWidth} });
}
function circle(x,y,d,fill,lineColor="none",lineWidth=0) {
  return slide.shapes.add({ geometry:"ellipse", position:{left:x,top:y,width:d,height:d}, fill, line:{fill:lineColor,width:lineWidth} });
}
function line(x,y,w,h,color=C.line,width=2) {
  return slide.shapes.add({ geometry:"line", position:{left:x,top:y,width:w,height:h}, fill:"none", line:{fill:color,width} });
}

slide.shapes.deleteAll();
slide.background.fill = C.white;

// Keep the deck's visual system.
rect(0,0,13,720,C.navy);
addText("02 结构",54,30,400,24,12,C.blue,true,"left",FONT_EN);
addText("暗场显微镜的六个核心部件",54,60,1110,62,34,C.navy,true);
line(54,132,1170,0,C.line,1);
addText("07",1180,674,44,20,11,C.gray,false,"right",FONT_EN);

// Left: one instrument image with numbered callouts.
rect(54,154,574,474,C.white,12,C.line,1);
slide.images.add({ blob:imageBytes, contentType:"image/png", alt:"AI生成的暗场显微镜六部件剖面示意图", fit:"contain", position:{left:68,top:162,width:546,height:454} });
addText("结构示意图",72,164,120,24,12,C.gray,true);

const colors = [C.orange, C.slate, C.cyan, C.green, C.blue, C.navy];
const callouts = [
  {n:"1", cx:106, cy:560, tx:316, ty:566},
  {n:"2", cx:522, cy:490, tx:348, ty:491},
  {n:"3", cx:106, cy:413, tx:316, ty:421},
  {n:"4", cx:522, cy:342, tx:348, ty:351},
  {n:"5", cx:106, cy:256, tx:316, ty:286},
  {n:"6", cx:522, cy:174, tx:348, ty:192},
];
for (let i=0;i<callouts.length;i++) {
  const c=callouts[i], col=colors[i];
  circle(c.cx,c.cy,38,col);
  addText(c.n,c.cx+5,c.cy+4,28,28,17,C.white,true,"center",FONT_EN);
  if(c.cx<300) line(c.cx+38,c.cy+19,c.tx-(c.cx+38),0,col,2);
  else line(c.tx,c.cy+19,c.cx-c.tx,0,col,2);
  circle(c.tx-15,c.ty-15,30,"none",col,3);
}

// Right: six visually selectable rows mapped to the same numbers/colors.
const items = [
  ["光源","提供足够的照明功率"],
  ["中央遮光片","阻断低角度直射光"],
  ["暗场聚光镜","形成高角度空心光锥"],
  ["样品 + 载玻片","把照明光散射或折射进物镜"],
  ["物镜","收集落入其 NA 的散射光"],
  ["相机 / 目镜","记录亮点、轮廓和轨迹"],
];
for (let i=0;i<items.length;i++) {
  const y=154+i*75, col=colors[i];
  rect(662,y,544,65,C.white,10,C.line,1);
  rect(662,y,7,65,col,4);
  circle(682,y+14,36,col);
  addText(String(i+1),688,y+19,24,24,15,C.white,true,"center",FONT_EN);
  addText(items[i][0],735,y+8,430,26,19,C.navy,true);
  addText(items[i][1],735,y+33,430,24,14,C.ink,false);
}
rect(662,611,544,34,C.pale,8);
addText("可选扩展：偏振片、光谱仪、微流控模块",680,615,508,26,14,C.navy,true,"center");
addText("编号从照明端到探测端，讲解时按 1—6 依次向上追踪光路",662,652,544,24,12,C.gray,false,"center");

slide.speakerNotes.append("\n本页讲法：从1号光源开始，沿光轴依次讲到6号相机或目镜。左侧为AI生成的教学结构示意图，不代表某一具体商品型号；编号、部件名称和功能为可编辑PowerPoint元素。建议按1至6顺序讲解，不依赖动画或超链接，在任何播放环境下都能使用。");

await fs.mkdir(BUILD,{recursive:true});
const candidate=path.join(BUILD,"candidate.pptx");
await (await PresentationFile.exportPptx(deck)).save(candidate);

const sourceBytes=await fs.readFile(SOURCE);
const sourceSha256=crypto.createHash("sha256").update(sourceBytes).digest("hex");
const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR,"container_tools/artifact_tool_utils.mjs")).href);
const result=await finalizePresentation({
  workspaceDir:WORKSPACE,
  candidatePath:candidate,
  finalPath:FINAL,
  pythonExecutable:RUNTIME_PYTHON,
  integrityValidatorPath:path.join(SKILL_DIR,"container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath:path.join(SKILL_DIR,"container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs:["--expected-slide-size-emu","12192000,6858000","--validate-bullet-geometry","--validate-heading-fit"],
  explicitTotalSlideCount:20,
  requiredNativeTableOwnerSlides:[],
  requiredNativeChartOwnerSlides:[],
  fontPolicy:{basis:"reference",families:[FONT_CN,FONT_EN],referencePath:SOURCE,referenceSha256:sourceSha256},
  verifyArtifactToolImport:true,
  receiptPath:path.join(BUILD,"structure-v3.validation.json"),
});

const finalDeck=await PresentationFile.importPptx(await FileBlob.load(FINAL));
const renderDir=path.join(BUILD,"render"); await fs.mkdir(renderDir,{recursive:true});
for(let i=0;i<finalDeck.slides.count;i++){
  const png=await finalDeck.export({slide:finalDeck.slides.getItem(i),format:"png",scale:1});
  await fs.writeFile(path.join(renderDir,`slide-${String(i+1).padStart(2,"0")}.png`),new Uint8Array(await png.arrayBuffer()));
}
console.log(JSON.stringify({final:FINAL,slideCount:finalDeck.slides.count,renderDir,result},null,2));
