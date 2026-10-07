import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { pathToFileURL } from "node:url";

const WORKSPACE = "C:/Users/Shjdshy/Desktop/AllDeskWork/暗场显微镜PPT_DDL100818pm";
const SKILL_DIR = "C:/Users/Shjdshy/.codex/plugins/cache/openai-primary-runtime/presentations/26.915.20218/skills/presentations";
const RUNTIME_PYTHON = "C:/Users/Shjdshy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe";
const SOURCE = path.join(WORKSPACE, "C已完成/暗场显微镜_课堂汇报_结构讲解优化版_v3.pptx");
const BUILD = path.join(WORKSPACE, "B工作区/PPT构建/geometry-v4");
const FINAL = path.join(WORKSPACE, "C已完成/暗场显微镜_课堂汇报_提交版_v4.pptx");

const { importRuntimeModule } = await import(pathToFileURL(path.join(SKILL_DIR, "container_tools/runtime_helpers.mjs")).href);
const { FileBlob, PresentationFile } = await importRuntimeModule("@oai/artifact-tool");
const deck = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const slide = deck.slides.getItem(5);

const C = { navy:"#173B63", blue:"#2C73A8", cyan:"#64B6CE", orange:"#F2A541", gray:"#647486", line:"#CBDCE5", pale:"#EAF4F8", white:"#FFFFFF", red:"#C95A5A", ink:"#172536" };
const FONT_CN="Microsoft YaHei";

function addText(text,x,y,w,h,size=18,color=C.ink,bold=false,align="left") {
  const s=slide.shapes.add({geometry:"textbox",position:{left:x,top:y,width:w,height:h},fill:"none",line:{fill:"none",width:0}});
  s.text=text;
  s.text.style={typeface:FONT_CN,fontSize:size,color,bold,alignment:align,verticalAlignment:"middle",autoFit:"shrinkText",lineSpacing:1.03};
  return s;
}
function rect(x,y,w,h,fill,radius=0,lineColor="none",lineWidth=0) {
  return slide.shapes.add({geometry:radius?"roundRect":"rect",position:{left:x,top:y,width:w,height:h},fill,line:{fill:lineColor,width:lineWidth}});
}
function circle(x,y,d,fill,lineColor="none",lineWidth=0) {
  return slide.shapes.add({geometry:"ellipse",position:{left:x,top:y,width:d,height:d},fill,line:{fill:lineColor,width:lineWidth}});
}
function ray(x,y,w,h,color,width=5) {
  return slide.shapes.add({geometry:"line",position:{left:x,top:y,width:w,height:h},fill:"none",line:{fill:color,width}});
}

// Remove only the original right-side optical sketch; preserve title, equation, inequality and sources.
for (const id of [
  "sh/v2tcn650","sh/u1kbu1ov","sh/5svutgni","sh/47mt0b6x",
  "sh/j6dcr65c","sh/i54bylor","sh/twfux0ne","sh/svmt4v6t",
  "sh/cbe5g3ih","sh/dcnm98zm","sh/qpwnet0b","sh/bq547yhw"
]) deck.resolve(id).delete();

// Correct geometry: oblique illumination reaches the specimen from below;
// unscattered rays continue outside the objective pupil; scattered light enters it.
rect(784,164,186,44,C.gray,9);
addText("物镜入口瞳",796,170,162,32,18,C.white,true,"center");

rect(760,385,236,16,C.white,2,C.blue,2);
circle(866,369,24,C.orange,C.white,2);
addText("样品",822,406,112,28,16,C.navy,true,"center");

// Two symmetric high-angle illumination paths, continuous through the specimen.
ray(646,392,232,156,C.orange,5);
ray(674,228,204,164,C.orange,5);
ray(878,228,204,164,C.orange,5);
ray(878,392,232,156,C.orange,5);

// Scattered light originates at the specimen and falls inside the objective pupil.
ray(858,208,0,177,C.cyan,5);
ray(898,208,0,177,C.cyan,5);

addText("未散射光绕开物镜",590,228,190,32,15,C.orange,true,"center");
addText("未散射光绕开物镜",984,228,190,32,15,C.orange,true,"center");
addText("散射光进入物镜",784,272,188,34,18,C.cyan,true,"center");
addText("高角度照明",778,540,200,34,18,C.orange,true,"center");

rect(650,586,456,44,C.pale,9);
addText("橙色：照明光    蓝色：样品散射光",670,591,416,32,16,C.navy,true,"center");
addText("对中偏差会让照明光进入物镜，背景随之变亮",682,628,392,27,14,C.red,true,"center");

slide.speakerNotes.append("\n图形修订说明：橙色高角度照明从下方到达样品；未散射光沿原方向继续传播并绕开物镜入口瞳；蓝色散射光从样品出发并进入物镜。该图用于说明照明环最小NA大于物镜NA的几何条件。\n");

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
  fontPolicy:{basis:"reference",families:["Microsoft YaHei","Arial"],referencePath:SOURCE,referenceSha256:sourceSha256},
  verifyArtifactToolImport:true,
  receiptPath:path.join(BUILD,"geometry-v4.validation.json"),
});

const finalDeck=await PresentationFile.importPptx(await FileBlob.load(FINAL));
const renderDir=path.join(BUILD,"render"); await fs.mkdir(renderDir,{recursive:true});
for(let i=0;i<finalDeck.slides.count;i++){
  const png=await finalDeck.export({slide:finalDeck.slides.getItem(i),format:"png",scale:1});
  await fs.writeFile(path.join(renderDir,`slide-${String(i+1).padStart(2,"0")}.png`),new Uint8Array(await png.arrayBuffer()));
}
console.log(JSON.stringify({final:FINAL,slideCount:finalDeck.slides.count,renderDir,result},null,2));
