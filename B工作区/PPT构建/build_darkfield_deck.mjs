import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const WORKSPACE = "C:/Users/Shjdshy/Desktop/AllDeskWork/暗场显微镜PPT_DDL100818pm";
const SKILL_DIR = "C:/Users/Shjdshy/.codex/plugins/cache/openai-primary-runtime/presentations/26.915.20218/skills/presentations";
const RUNTIME_PYTHON = "C:/Users/Shjdshy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe";
const BUILD = path.join(WORKSPACE, "B工作区/PPT构建");
const OUT = path.join(WORKSPACE, "C已完成");
const FINAL = path.join(OUT, "暗场显微镜_课堂汇报_样品版_v2.pptx");
const W = 1280, H = 720;
const { importRuntimeModule } = await import(pathToFileURL(path.join(SKILL_DIR, "container_tools/runtime_helpers.mjs")).href);
const { Presentation, PresentationFile, FileBlob } = await importRuntimeModule("@oai/artifact-tool");
const C = {
  navy: "#173B63", blue: "#2C73A8", cyan: "#64B6CE", pale: "#EAF4F8",
  ink: "#172536", gray: "#647486", line: "#CBDCE5", white: "#FFFFFF",
  orange: "#F2A541", red: "#C95A5A", dark: "#071524", green: "#3D8B7D",
};
const FONT_CN = "Microsoft YaHei";
const FONT_EN = "Arial";
const presentation = Presentation.create({ slideSize: { width: W, height: H } });

const A = (...parts) => path.join(WORKSPACE, ...parts);
const assets = {
  stars: A("B工作区", "PPT构建", "assets", "星空背景对照.png"),
  principle: A("A资料区", "05_图片区", "01_原理图", "01_透射暗场_遮住直射光的空心光锥.png"),
  historic: A("A资料区", "05_图片区", "02_实物与核心部件", "01_历史浸没式超显微镜实物.png"),
  side: A("A资料区", "05_图片区", "03_典型结构", "01_1902超显微镜侧向照明结构.png"),
  tir: A("A资料区", "05_图片区", "03_典型结构", "02_物镜型TIR暗场光路.png"),
  diatom: A("A资料区", "05_图片区", "04_典型成像", "01_硅藻暗场图像.png"),
  virus: A("A资料区", "05_图片区", "04_典型成像", "02_流感病毒浓度与暗场亮点.png"),
  hyper: A("A资料区", "05_图片区", "05_现代扩展", "01_高光谱增强暗场系统.png"),
  scratch: A("B工作区", "PPT构建", "assets", "微划痕_暗场系统与图像.png"),
  meng1: A("B工作区", "PPT构建", "assets", "Meng_双微镜TIR暗场_Fig1.png"),
  meng2: A("B工作区", "PPT构建", "assets", "Meng_定位与粒径性能_Fig2.png"),
};
const imageBytes = {};
for (const [key, file] of Object.entries(assets)) imageBytes[key] = await fs.readFile(file);

function addText(slide, text, x, y, w, h, size=24, color=C.ink, bold=false, align="left", family=FONT_CN) {
  const s = slide.shapes.add({ geometry: "textbox", position: { left:x, top:y, width:w, height:h }, fill:"none", line:{fill:"none",width:0} });
  s.text = text;
  s.text.style = { typeface: family, fontSize:size, color, bold, alignment:align, verticalAlignment:"middle", autoFit:"shrinkText", lineSpacing:1.06 };
  return s;
}
function rect(slide, x,y,w,h, fill, radius=0, line="none", lineWidth=0) {
  return slide.shapes.add({ geometry: radius ? "roundRect" : "rect", position:{left:x,top:y,width:w,height:h}, fill, line:{fill:line,width:lineWidth} });
}
function line(slide, x,y,w,h, color=C.line, width=2, dash="solid") {
  return slide.shapes.add({ geometry:"line", position:{left:x,top:y,width:w,height:h}, fill:"none", line:{fill:color,width,style:dash} });
}
function circle(slide, x,y,d, fill, stroke="none", sw=0) {
  return slide.shapes.add({ geometry:"ellipse", position:{left:x,top:y,width:d,height:d}, fill, line:{fill:stroke,width:sw} });
}
function image(slide, key, x,y,w,h, fit="contain", alt="") {
  return slide.images.add({ blob:imageBytes[key], contentType:"image/png", alt, fit, position:{left:x,top:y,width:w,height:h} });
}
function base(slide, section, title, page) {
  slide.background.fill = C.white;
  rect(slide, 0, 0, 13, H, C.navy);
  addText(slide, section.toUpperCase(), 54, 30, 400, 24, 12, C.blue, true, "left", FONT_EN);
  addText(slide, title, 54, 60, 1110, 62, 34, C.navy, true);
  line(slide, 54, 132, 1170, 0, C.line, 1);
  addText(slide, String(page).padStart(2,"0"), 1180, 674, 44, 20, 11, C.gray, false, "right", FONT_EN);
}
function chip(slide, text, x,y,w, color=C.blue) {
  rect(slide,x,y,w,28,color,14);
  addText(slide,text,x+8,y+1,w-16,26,12,C.white,true,"center");
}
function card(slide, x,y,w,h, title, body, accent=C.blue) {
  rect(slide,x,y,w,h,C.pale,12,C.line,1);
  rect(slide,x,y,7,h,accent,4);
  addText(slide,title,x+24,y+10,w-42,32,20,C.navy,true);
  addText(slide,body,x+24,y+44,w-42,h-50,16,C.ink,false);
}
function note(slide, body, sources=[]) {
  const s = sources.length ? `\n\n来源与图片：\n${sources.map(v=>`- ${v}`).join("\n")}` : "";
  slide.speakerNotes.textFrame.setText(`${body}${s}`);
  slide.speakerNotes.setVisible(true);
}
function source(slide, text) { addText(slide,text,54,650,1080,18,10,C.gray,false); }

// 1 Cover
{
  const s=presentation.slides.add(); s.background.fill=C.white;
  rect(s,0,0,26,H,C.navy); rect(s,26,0,10,H,C.cyan);
  addText(s,"DARK-FIELD MICROSCOPY",80,82,700,28,13,C.blue,true,"left",FONT_EN);
  addText(s,"暗场显微镜",80,133,820,92,54,C.navy,true);
  addText(s,"把最亮的背景拿掉，让微弱散射成为主角",84,242,820,45,24,C.ink,false);
  line(s,84,314,500,0,C.orange,5);
  card(s,84,367,338,126,"原理","光路、NA 条件与成像边界",C.blue);
  card(s,446,367,338,126,"演化","从超显微镜到 TIR 与高光谱",C.cyan);
  card(s,808,367,338,126,"应用","病毒、纳米颗粒与表面缺陷",C.orange);
  addText(s,"汇报人：________    课程：________    日期：2026.10",84,606,820,36,17,C.gray,false);
  addText(s,"样品版 · 12–15 min",1010,638,150,24,12,C.blue,true,"right",FONT_EN);
  note(s,"开场不要先下定义。先说：这不是一台“把灯调暗”的显微镜，而是一种重新安排照明和收集方向的方法。随后用第2页观星类比建立直觉。",[
    "视觉风格参考：SciFig Conference Presentation Template — Minimal Navy, https://github.com/scifig-ai/scientific-poster-and-presentation-templates"
  ]);
}

// 2 Stars
{
  const s=presentation.slides.add(); base(s,"01 直觉引入","为什么城市里星星少了？",2);
  image(s,"stars",54,158,760,428,"cover","同一天空在城市光污染与暗夜环境中的对照；AI生成示意图");
  rect(s,54,545,760,41,C.dark); addText(s,"同一片星空：左侧背景更亮，右侧弱星更容易被察觉",72,550,724,30,17,C.white,true,"center");
  addText(s,"信号没有突然变强",862,180,300,45,24,C.navy,true);
  addText(s,"真正改变的是：",862,235,250,30,17,C.gray,false);
  card(s,850,284,332,90,"背景 ↓","弱信号与背景的差距被拉开",C.blue);
  card(s,850,391,332,90,"对比度 ↑","原本被淹没的目标显现",C.cyan);
  addText(s,"类比边界：星星自己发光；显微样品通常靠散射光被看见。",850,512,332,72,15,C.red,true);
  source(s,"图：AI 生成，仅作概念类比，不作为实验数据");
  note(s,"提问式开场：‘星星真的在城市里变少了吗？’答案是否定的，变亮的是背景。暗场的共同思想正是背景抑制。但要立刻说明：天文学中的星体是自发光对象，显微样品多是把照明光散射进物镜。更严格的天文类比其实是日冕仪。",[]);
}

// 3 Need
{
  const s=presentation.slides.add(); base(s,"01 需求","亮场的困境：目标和背景都很亮",3);
  addText(s,"相机看到的是三部分叠加",54,160,430,35,24,C.navy,true);
  const xs=[76,304,532]; const labels=["直射光\n强背景","样品散射\n弱信号","杂散与噪声\n伪信号"]; const colors=[C.orange,C.cyan,C.gray];
  xs.forEach((x,i)=>{circle(s,x,222,154,colors[i]);addText(s,labels[i],x+12,258,130,72,18,C.white,true,"center");});
  addText(s,"+",250,270,35,35,28,C.navy,true,"center"); addText(s,"+",478,270,35,35,28,C.navy,true,"center");
  rect(s,54,420,650,104,C.pale,12); addText(s,"当弱散射信号只占总光强的一小部分，细边缘、透明颗粒和浅划痕就会被背景吞没。",82,438,594,68,22,C.ink,true);
  card(s,765,168,420,112,"需求 1｜无染色","不靠吸收差异，也能看到透明或弱吸收对象",C.blue);
  card(s,765,300,420,112,"需求 2｜高通量","比逐点扫描更快地发现稀少目标",C.cyan);
  card(s,765,432,420,112,"需求 3｜可定量","把亮点升级为计数、轨迹或光谱",C.orange);
  source(s,"概念依据：Davidson & Abramowitz, Optical Microscopy (2002)");
  note(s,"这一页先不讲结构，只把问题说透：亮场并非‘看不见’，而是强背景占用了动态范围。暗场优先解决的是可见度和信噪比，而不是突破衍射极限。",[
    "Davidson & Abramowitz (2002), https://www.microscopyu.com/pdfs/Davidson_and_Abramowitz_2002.pdf"
  ]);
}

// 4 Effect
{
  const s=presentation.slides.add(); base(s,"01 效果","暗场的产品效果：黑背景上的散射轮廓",4);
  rect(s,54,158,520,430,C.dark,12); image(s,"diatom",126,174,378,400,"contain","硅藻暗场图像");
  addText(s,"硅藻壳体的边缘和纹理被散射光勾勒出来",624,178,530,64,28,C.navy,true);
  const bullets=[
    ["看见什么","边缘、颗粒、界面和粗糙处通常更亮"],
    ["为什么亮","它们把原本绕开物镜的光重新定向"],
    ["代价是什么","灰尘、划痕、盖玻片也会一起变亮"],
  ];
  bullets.forEach((b,i)=>card(s,624,270+i*104,532,84,b[0],b[1],[C.blue,C.cyan,C.orange][i]));
  chip(s,"不是“更高分辨率”",836,585,210,C.red);
  source(s,"图源：Davidson & Abramowitz (2002), Fig. 15（本地资料裁图）");
  note(s,"暗场图像最吸引人的地方是‘亮点很多’，但课堂上要把视觉冲击和科学含义分开：亮点代表有光被散射进物镜，不自动等于粒子真实尺寸或化学身份。",[
    "Davidson & Abramowitz (2002), https://www.microscopyu.com/pdfs/Davidson_and_Abramowitz_2002.pdf"
  ]);
}

// 5 Core principle
{
  const s=presentation.slides.add(); base(s,"02 原理","一句话原理：挡住直射光，只收散射光",5);
  image(s,"principle",54,158,480,460,"contain","透射暗场中央遮光与空心光锥示意图");
  addText(s,"① 造一个空心光锥",606,173,520,44,27,C.navy,true);
  addText(s,"中央遮光片挡住低角度直射光",606,216,520,35,18,C.gray,false);
  line(s,606,266,480,0,C.line,1);
  addText(s,"② 让未散射光绕开物镜",606,290,520,44,27,C.navy,true);
  addText(s,"样品不存在时，视野接近黑色",606,333,520,35,18,C.gray,false);
  line(s,606,383,480,0,C.line,1);
  addText(s,"③ 样品把光改向物镜",606,407,520,44,27,C.navy,true);
  addText(s,"散射、折射或反射后的光形成亮像",606,450,520,35,18,C.gray,false);
  rect(s,606,518,520,75,C.navy,10); addText(s,"暗场 = 照明路径与收集路径的空间分离",626,531,480,48,22,C.white,true,"center");
  source(s,"图源：Davidson & Abramowitz (2002), Fig. 14");
  note(s,"讲述节奏：造光锥—绕开物镜—样品改向。只要观众记住‘没有样品时光进不了物镜，有样品时散射光才能进来’，后续所有结构变化都能理解。",[
    "Davidson & Abramowitz (2002), https://www.microscopyu.com/pdfs/Davidson_and_Abramowitz_2002.pdf"
  ]);
}

// 6 NA geometry
{
  const s=presentation.slides.add(); base(s,"02 原理","几何条件：照明锥必须从物镜外侧掠过",6);
  addText(s,"NA = n · sin θ",74,160,430,64,38,C.navy,true,"center",FONT_EN);
  addText(s,"介质折射率 × 光线半角",74,219,430,31,16,C.gray,false,"center");
  rect(s,72,288,476,204,C.pale,12);
  addText(s,"暗场成立的关键不等式",96,308,420,30,20,C.blue,true,"center");
  addText(s,"照明环最小 NA  >  物镜 NA",96,357,420,56,29,C.navy,true,"center");
  addText(s,"比较的是照明环的最小 NA，而不是只比较聚光镜铭牌上的最大 NA。",105,427,402,47,16,C.ink,false,"center");
  // editable schematic
  rect(s,746,174,150,38,C.gray,6); addText(s,"物镜入口瞳",756,177,130,30,15,C.white,true,"center");
  circle(s,774,336,94,C.white,C.blue,3); addText(s,"样品",788,365,66,24,17,C.navy,true,"center");
  line(s,630,217,190,118,C.orange,5); line(s,844,217,190,118,C.orange,5);
  line(s,758,206,62,128,C.cyan,4); line(s,844,206,62,128,C.cyan,4);
  addText(s,"高角度照明",573,214,154,30,16,C.orange,true,"center"); addText(s,"散射光进入",930,214,156,30,16,C.cyan,true,"center");
  addText(s,"直射光绕开物镜",700,516,260,38,22,C.navy,true,"center");
  addText(s,"对中偏差 → 光锥蹭进物镜 → 背景立刻变灰",668,567,330,50,16,C.red,true,"center");
  source(s,"条件表述依据：Davidson & Abramowitz (2002); Nikon MicroscopyU Darkfield Illumination");
  note(s,"这是全场最重要的干货页。注意表述为照明环的最小数值孔径大于物镜NA。若只说‘聚光镜NA更大’并不充分，因为关键是内侧边缘也要在物镜接受角之外。",[
    "Davidson & Abramowitz (2002), https://www.microscopyu.com/pdfs/Davidson_and_Abramowitz_2002.pdf",
    "Nikon MicroscopyU, Darkfield Illumination, https://www.microscopyu.com/techniques/stereomicroscopy/darkfield-illumination"
  ]);
}

// 7 Components
{
  const s=presentation.slides.add(); base(s,"02 结构","一台暗场显微镜长什么样？",7);
  const parts=[
    ["光源","提供足够的照明功率"], ["中央遮光片","去掉低角度直射光"],
    ["暗场聚光镜","形成高角度空心光锥"], ["样品 + 载玻片","把光散射/折射进物镜"],
    ["物镜","只接收落入其 NA 的光"], ["相机 / 目镜","记录亮点、轮廓或轨迹"],
  ];
  parts.forEach((p,i)=>{
    const x=54+(i%3)*394, y=166+Math.floor(i/3)*160;
    rect(s,x,y,360,126,C.white,12,C.line,1); circle(s,x+22,y+27,56,[C.blue,C.cyan,C.orange][i%3]);
    addText(s,String(i+1),x+34,y+39,32,30,19,C.white,true,"center",FONT_EN);
    addText(s,p[0],x+94,y+20,240,34,21,C.navy,true); addText(s,p[1],x+94,y+62,238,45,15,C.ink,false);
  });
  rect(s,54,510,1148,102,C.navy,12);
  addText(s,"可升级模块",78,531,180,32,20,C.white,true);
  addText(s,"偏振片：抑制表面反光   ·   光谱仪：记录散射谱   ·   微流控：控制颗粒与浓度",260,526,900,49,18,C.white,false,"center");
  note(s,"这一页解决‘它长什么样’。普通透射显微镜加中央遮光片可以做低NA暗场；高NA成像需要专用暗场聚光镜及更严格的对中。升级模块不是暗场成立的必要条件。",[
    "Nikon MicroscopyU, Darkfield Illumination, https://www.microscopyu.com/techniques/stereomicroscopy/darkfield-illumination"
  ]);
}

// 8 Structures
{
  const s=presentation.slides.add(); base(s,"02 结构","三种结构，解决三类场景",8);
  const cards=[
    {x:54,key:"principle",title:"透射暗场",sub:"透明样品 / 微生物",body:"环形斜照明从下方入射；未散射光绕开物镜"},
    {x:444,key:"scratch",title:"反射暗场",sub:"晶圆 / 光学表面缺陷",body:"从上方斜照明；划痕与颗粒把光散射回物镜"},
    {x:834,key:"tir",title:"TIR 暗场",sub:"界面 / 单颗粒追踪",body:"全反射产生倏逝场，把照明限制在界面附近"},
  ];
  cards.forEach((c,i)=>{
    rect(s,c.x,158,352,446,C.white,12,C.line,1); image(s,c.key,c.x+18,174,316,230,"contain",c.title);
    chip(s,c.title,c.x+20,420,118,[C.blue,C.orange,C.cyan][i]);
    addText(s,c.sub,c.x+22,463,308,36,18,C.navy,true); addText(s,c.body,c.x+22,512,308,68,15,C.ink,false);
  });
  source(s,"图源：Davidson 2002；Li et al., Chinese Optics Letters 2017；Enoki et al., PLOS ONE 2012");
  note(s,"不要把三种结构讲成三台互不相关的仪器。它们共享同一原则：尽量把照明光和收集光在角度或空间上分开。透射适合透明样品，反射适合不透明表面，TIR把背景压缩到界面附近。",[
    "Davidson & Abramowitz (2002), https://www.microscopyu.com/pdfs/Davidson_and_Abramowitz_2002.pdf",
    "Li et al. (2017), doi:10.3788/COL201715.081202",
    "Enoki et al. (2012), doi:10.1371/journal.pone.0049208"
  ]);
}

// 9 Interaction
{
  const s=presentation.slides.add(); base(s,"02 原理","图像亮度从哪里来？样品把光重新分配",9);
  const stages=[
    ["入射","高角度照明",C.orange], ["相互作用","散射 / 折射 / 反射",C.cyan],
    ["收集","进入物镜 NA",C.blue], ["成像","点扩散函数 + 噪声",C.navy],
  ];
  stages.forEach((v,i)=>{
    const x=54+i*294; circle(s,x+92,193,100,v[2]); addText(s,String(i+1),x+120,221,44,44,28,C.white,true,"center",FONT_EN);
    addText(s,v[0],x+12,314,260,35,23,C.navy,true,"center"); addText(s,v[1],x+12,356,260,54,16,C.ink,false,"center");
    if(i<3){rect(s,x+252,232,74,22,C.pale,11); addText(s,"→",x+263,228,52,28,26,C.blue,true,"center");}
  });
  rect(s,80,466,1120,116,C.pale,12);
  addText(s,"Rayleigh 小颗粒近似",108,487,250,30,19,C.blue,true);
  addText(s,"Iscat ∝ d⁶",374,478,260,51,32,C.navy,true,"center",FONT_EN);
  addText(s,"粒径减半，散射信号约降为 1/64；但该标度并非对所有尺寸、形状和材料都成立。",664,480,496,70,17,C.ink,false);
  source(s,"定量边界：Olson et al., Chem. Soc. Rev. 2015; Meng et al., ACS Photonics 2021");
  note(s,"这一页解释为什么暗场既灵敏又困难：抑制背景后可以看见弱散射，但粒子变小时信号下降非常快。d^6只用于Rayleigh小颗粒近似，不能当普适定律。",[
    "Olson et al. (2015), doi:10.1039/C4CS00131A",
    "Meng et al. (2021), doi:10.1021/acsphotonics.1c01268"
  ]);
}

// 10 detect vs resolve
{
  const s=presentation.slides.add(); base(s,"02 边界","看见、定位、分辨：是三件不同的事",10);
  const rows=[
    ["看见 Detection","背景上出现可信亮点","信噪比 / 检出限",C.cyan],
    ["定位 Localization","估计亮点中心坐标","定位精度 / 轨迹",C.blue],
    ["分辨 Resolution","区分相邻结构或真实形状","点扩散函数 / Rayleigh判据",C.orange],
  ];
  rows.forEach((r,i)=>{
    const y=166+i*126; rect(s,54,y,1150,100,C.white,12,C.line,1); rect(s,54,y,234,100,r[3],12);
    addText(s,r[0],72,y+20,198,58,20,C.white,true,"center"); addText(s,r[1],322,y+18,440,32,20,C.navy,true); addText(s,r[2],322,y+57,440,25,15,C.gray,false);
    const dots=i===0?1:i===1?1:2; for(let k=0;k<dots;k++) circle(s,947+k*42,y+35,26,C.white,r[3],5);
  });
  rect(s,54,566,1150,58,C.navy,10); addText(s,"暗场主要提高“看见”的概率；后续算法可改善定位，但不会自动改变物镜的衍射分辨率。",74,574,1110,42,20,C.white,true,"center");
  note(s,"这是纠正常见误解的关键页。可以用一句话收束：看见一个4 nm颗粒的散射点，不等于用4 nm空间分辨率看见了颗粒形状。暗场提升检测对比度，定位还依赖光子数和拟合，分辨率仍受成像系统限制。",[
    "Priest et al. (2021), doi:10.1021/acs.chemrev.1c00271"
  ]);
}

// 11 Evolution
{
  const s=presentation.slides.add(); base(s,"03 演化","技术演化不是换名字，而是逐个压低背景",11);
  const items=[
    ["1902","侧向超显微","直射光太强","照明与观察方向分离"],
    ["环形聚光","透射暗场","视场与对中","把高角照明整合进显微镜"],
    ["TIR","界面暗场","体相背景","只照亮界面附近"],
    ["高光谱 / 计算","定量暗场","亮点缺少身份","用光谱、轨迹和模型赋值"],
  ];
  line(s,120,344,1030,0,C.line,8);
  items.forEach((it,i)=>{
    const x=70+i*292; circle(s,x+50,313,62,[C.navy,C.blue,C.cyan,C.orange][i]);
    addText(s,it[0],x+2,196,160,32,18,C.blue,true,"center",FONT_EN);
    addText(s,it[1],x+2,238,160,54,22,C.navy,true,"center");
    card(s,x,392,222,150,`瓶颈：${it[2]}`,it[3],[C.navy,C.blue,C.cyan,C.orange][i]);
  });
  addText(s,"主线：背景来源改变 → 光路策略改变 → 可测量指标改变",270,580,740,42,22,C.navy,true,"center");
  source(s,"历史起点：Siedentopf & Zsigmondy 1902；现代路线：Priest et al. 2021");
  note(s,"把发展史放在基本原理之后是更适合课堂的顺序。观众先理解‘为什么要隔离照明’，才看得懂每一代技术在解决哪一种背景。不要只按年份罗列仪器名称。",[
    "Siedentopf & Zsigmondy (1902), doi:10.1002/andp.19023150102",
    "Priest et al. (2021), doi:10.1021/acs.chemrev.1c00271"
  ]);
}

// 12 Virus
{
  const s=presentation.slides.add(); base(s,"04 应用","案例 1｜无标记流感病毒：亮点计数 → 浓度",12);
  image(s,"virus",54,158,696,420,"contain","流感病毒浓度与TIR暗场亮点关系");
  card(s,790,174,390,104,"对象","约 100 nm 流感病毒颗粒",C.blue);
  card(s,790,296,390,104,"方法","物镜型 TIR 暗场，压低体相背景",C.cyan);
  card(s,790,418,390,104,"论文数据","报告检出限 1.2 × 10⁴ PFU/mL",C.orange);
  addText(s,"注意：PFU/mL 表示感染性滴度；亮点数不应简单等同于病毒总颗粒数",790,544,390,58,15,C.red,true);
  source(s,"Enoki et al., PLOS ONE 7(11):e49208 (2012), CC BY；doi:10.1371/journal.pone.0049208");
  note(s,"讲这张数据图时强调‘可定量’来自标定曲线，而不是仅凭亮点漂亮。论文的1.2×10^4 PFU/mL是按背景加3倍标准差外推的检出限；PFU是感染性单位，不等于颗粒总数。",[
    "Enoki et al. (2012), doi:10.1371/journal.pone.0049208"
  ]);
}

// 13 Tracking
{
  const s=presentation.slides.add(); base(s,"04 应用","案例 2｜单颗粒追踪：暗场亮点 → 纳米轨迹",13);
  image(s,"meng1",54,166,620,334,"contain","微镜TIR暗场与iSCAT对比");
  image(s,"meng2",704,166,500,278,"contain","微镜TIR暗场定位和粒径性能");
  chip(s,"20 nm AuNPs",92,516,136,C.orange); chip(s,"6 μs exposure",246,516,144,C.blue); chip(s,"~4 nm 定位分布直径",408,516,200,C.cyan);
  addText(s,"这些数值是在特定光功率、曝光与拟合条件下得到的；“4 nm”是定位分布，不是成像分辨率。",704,478,478,80,17,C.ink,true);
  source(s,"Meng et al., ACS Photonics 8, 3111–3118 (2021), Figs. 1–2；doi:10.1021/acsphotonics.1c01268");
  note(s,"先指出传统TIR要把照明耦合进高NA物镜，背景和视场受限。作者用两块微镜把入射和收集分开。20 nm金颗粒可在6微秒曝光下追踪；4 nm指位置波动形成的定位分布直径，不等于物镜的4 nm分辨率。",[
    "Meng et al. (2021), doi:10.1021/acsphotonics.1c01268"
  ]);
}

// 14 Hyperspectral
{
  const s=presentation.slides.add(); base(s,"04 应用","案例 3｜高光谱暗场：从“亮不亮”到“散射谱是什么”",14);
  image(s,"hyper",54,162,730,420,"contain","高光谱增强暗场系统与数据立方体");
  const steps=[
    ["空间图像","定位颗粒或细胞区域"], ["光谱维度","读取每个像素的散射谱"], ["模型/标定","关联尺寸、聚集与环境变化"],
  ];
  steps.forEach((v,i)=>card(s,824,170+i*128,346,104,v[0],v[1],[C.blue,C.cyan,C.orange][i]));
  addText(s,"测到的是“系统响应过滤后的相对散射谱”，定量前要校正灯谱、暗计数与基底。",824,570,346,52,15,C.red,true);
  source(s,"Zamora-Perez et al., Materials 11, 243 (2018), CC BY；doi:10.3390/ma11020243");
  note(s,"高光谱暗场的价值是把二维图片扩展成x-y-λ数据立方体。颜色可与等离激元共振、聚集或局部折射率相关，但不是物体的‘真实颜色’，且必须做系统响应校正。",[
    "Zamora-Perez et al. (2018), doi:10.3390/ma11020243"
  ]);
}

// 15 Defects
{
  const s=presentation.slides.add(); base(s,"04 应用","案例 4｜浅划痕检测：缺陷把斜入射光散射回来",15);
  image(s,"scratch",54,154,496,468,"contain","反射暗场微划痕检测系统及图像");
  addText(s,"为什么适合缺陷分析？",600,168,560,44,28,C.navy,true);
  card(s,600,232,556,90,"背景","完好镜面主要把光反射到物镜之外",C.blue);
  card(s,600,338,556,90,"缺陷","划痕边缘与粗糙区把光散射进相机",C.cyan);
  card(s,600,444,556,90,"算法","自适应平滑 + 形态学差分分离浅划痕",C.orange);
  rect(s,600,554,556,57,C.navy,9); addText(s,"论文实验：浅划痕检出率约 82%",620,562,516,40,21,C.white,true,"center");
  source(s,"Li et al., Chinese Optics Letters 15, 081202 (2017)；doi:10.3788/COL201715.081202");
  note(s,"这张图能把反射暗场讲得非常具体：环形斜照明照向超光滑表面，完好区镜面反射不进物镜，浅划痕把光散射回来。论文针对深度低于50 nm的浅划痕，报告约82%检出率；这是特定样本和算法条件下的结果。",[
    "Li et al. (2017), doi:10.3788/COL201715.081202",
    "Open PDF: https://researching.cn/ArticlePdf/m00005/2017/15/8/COL201715081202.pdf"
  ]);
}

// 16 Comparison
{
  const s=presentation.slides.add(); base(s,"04 选择","暗场不是万能：按问题选择相关技术",16);
  const cols=[54,316,578,840,1102];
  const headers=["技术","最擅长","主要代价","适合回答","与暗场关系"];
  headers.forEach((h,i)=>{rect(s,cols[i],158,i===4?102:244,48,C.navy,0);addText(s,h,cols[i]+6,164,i===4?90:232,34,15,C.white,true,"center");});
  const rows=[
    ["暗场","弱散射可见度","缺少化学特异性","有没有颗粒/缺陷？","基线方案"],
    ["相差 / DIC","透明相位结构","定量解释复杂","细胞内部轮廓？","互补"],
    ["荧光","分子特异性","标记与漂白","它是谁？在哪里？","验证身份"],
    ["iSCAT","极弱散射灵敏度","相位背景与条纹","更小颗粒/蛋白？","干涉增强"],
  ];
  rows.forEach((r,ri)=>r.forEach((v,ci)=>{
    const widths=[244,244,244,244,102]; const y=206+ri*94; rect(s,cols[ci],y,widths[ci],94,ri%2?C.white:C.pale,0,C.line,1);
    addText(s,v,cols[ci]+8,y+8,widths[ci]-16,78,ci===4?13:15,ci===0?C.navy:C.ink,ci===0,"center");
  }));
  source(s,"技术框架参考：Priest et al., Chemical Reviews 2021");
  note(s,"这页用于回答‘为什么不用别的方法’。暗场最适合无标记、弱散射、快速筛查；需要分子身份时通常结合荧光，需要极弱信号时可用干涉散射，但系统与背景模型更复杂。",[
    "Priest et al. (2021), doi:10.1021/acs.chemrev.1c00271"
  ]);
}

// 17 Failure modes
{
  const s=presentation.slides.add(); base(s,"05 局限","最容易翻车的四个地方",17);
  const risks=[
    ["背景不黑","灰尘、指纹、盖玻片粗糙度都会散射","清洁 + 空白视野 + 背景扣除"],
    ["亮点不等于目标","非特异颗粒与气泡同样会亮","对照组 + 特异性验证"],
    ["亮度不等于厚度","强度受尺寸、材料、角度和焦面共同影响","标定曲线 + 固定几何"],
    ["颜色不等于真色","光谱受灯、物镜NA和相机响应过滤","白板/暗场校正 + 光谱响应校准"],
  ];
  risks.forEach((r,i)=>{
    const x=54+(i%2)*584, y=164+Math.floor(i/2)*212;
    rect(s,x,y,548,180,C.white,12,C.line,1); rect(s,x,y,548,48,[C.red,C.orange,C.blue,C.cyan][i],12);
    addText(s,r[0],x+18,y+6,512,34,20,C.white,true);
    addText(s,"风险｜"+r[1],x+22,y+62,500,45,16,C.ink,true);
    addText(s,"控制｜"+r[2],x+22,y+116,500,42,16,C.green,true);
  });
  source(s,"实验控制建议综合：Davidson 2002；Olson 2015；Zamora-Perez 2018");
  note(s,"如果汇报时间紧，这页至少讲前两项。暗场把所有散射体都变亮，因此清洁和对照不是附属操作，而是暗场成像质量的一部分。定量强度或颜色时还要固定照明、焦面、曝光并做校正。",[
    "Davidson & Abramowitz (2002), https://www.microscopyu.com/pdfs/Davidson_and_Abramowitz_2002.pdf",
    "Olson et al. (2015), doi:10.1039/C4CS00131A",
    "Zamora-Perez et al. (2018), doi:10.3390/ma11020243"
  ]);
}

// 18 Outlook
{
  const s=presentation.slides.add(); base(s,"05 展望","有效的展望：每条路线都对应一个瓶颈和指标",18);
  const items=[
    ["背景仍高","更窄的照明体积：TIR、光片、偏振抑制","SBR / 检出限"],
    ["缺少身份","光谱、抗体或多模态联合","分类准确率 / 特异性"],
    ["速度不足","高亮度光源 + 高速相机 + 在线算法","帧率 / 通量"],
    ["定量不稳","标准样品、物理模型与不确定度报告","重复性 / 误差"],
  ];
  items.forEach((it,i)=>{
    const y=160+i*110; chip(s,it[0],54,y+10,118,[C.red,C.orange,C.blue,C.cyan][i]);
    addText(s,it[1],200,y,650,48,20,C.navy,true); addText(s,"评价指标："+it[2],866,y+4,320,40,16,C.gray,true,"right"); line(s,200,y+66,986,0,C.line,1);
  });
  rect(s,54,606,1132,42,C.pale,10); addText(s,"判断“前沿”是否有效：能否把瓶颈转化为可验证的指标，而不是只增加系统复杂度。",74,611,1092,31,18,C.navy,true,"center");
  note(s,"避免空泛地说‘智能化、集成化、产业化’。每条展望都必须回答：现有瓶颈是什么、技术路线是什么、用什么指标验证进步。课堂上可重点讲‘从可见到可定量’。",[
    "Priest et al. (2021), doi:10.1021/acs.chemrev.1c00271",
    "Li et al. (2025), doi:10.1021/acsami.5c17222"
  ]);
}

// 19 Summary
{
  const s=presentation.slides.add(); base(s,"06 总结","带走三句话",19);
  const items=[
    ["01","暗场的本质","不是把灯调暗，而是让直射光绕开物镜。"],
    ["02","暗场的价值","提升弱散射目标的对比度、检出与追踪能力。"],
    ["03","暗场的边界","看见不等于分辨；定量必须依赖校正、标定和对照。"],
  ];
  items.forEach((it,i)=>{
    const y=162+i*136; circle(s,62,y+14,76,[C.navy,C.blue,C.orange][i]); addText(s,it[0],77,y+31,46,35,22,C.white,true,"center",FONT_EN);
    addText(s,it[1],172,y,300,46,27,C.navy,true); addText(s,it[2],172,y+51,946,52,21,C.ink,false);
  });
  rect(s,54,583,1132,62,C.navy,10); addText(s,"从星空到显微镜：真正重要的不是让目标更亮，而是让背景更安静。",74,591,1092,46,22,C.white,true,"center");
  note(s,"结尾回扣第2页。可以停顿后说：暗场最值得记住的设计思想，是先控制背景，再谈信号。随后感谢听众，并把第20页留作答疑时的文献入口。",[]);
}

// 20 References
{
  const s=presentation.slides.add(); base(s,"参考文献","核心文献与素材出处",20);
  const left=[
    "1  Davidson & Abramowitz. Optical Microscopy (2002).",
    "2  Siedentopf & Zsigmondy. Ann. Phys. 315, 1–39 (1902).",
    "3  Zsigmondy. Properties of Colloids, Nobel Lecture (1926).",
    "4  Olson et al. Chem. Soc. Rev. 44, 40–57 (2015).",
    "5  Priest et al. Chem. Rev. 121, 11937–11970 (2021).",
  ];
  const right=[
    "6  Enoki et al. PLOS ONE 7, e49208 (2012).",
    "7  Meng et al. ACS Photonics 8, 3111–3118 (2021).",
    "8  Zamora-Perez et al. Materials 11, 243 (2018).",
    "9  Li et al. Chinese Optics Letters 15, 081202 (2017).",
    "10 Li et al. ACS Appl. Mater. Interfaces 17, 64027–64047 (2025).",
  ];
  addText(s,"基础与综述",54,158,520,32,21,C.blue,true); addText(s,left.join("\n\n"),54,204,540,330,16,C.ink,false);
  addText(s,"应用与前沿",650,158,520,32,21,C.blue,true); addText(s,right.join("\n\n"),650,204,540,330,16,C.ink,false);
  rect(s,54,562,1132,62,C.pale,10); addText(s,"完整 DOI、PDF、本地图片授权说明与论文原文均已整理在 A资料区；模板来源见 B工作区/模板。",74,571,1092,44,18,C.navy,true,"center");
  source(s,"课堂展示请保留图下注明；公开发布前再次核对非开放获取图片许可");
  note(s,"参考文献页不必逐条念。答疑时可根据问题跳回对应案例。PLOS ONE与Materials为开放获取来源；其他图用于课堂演示时仍保留完整出处，若公开上传需再次核对许可。",[
    "Local bibliography: A资料区/00_导读与索引/references.bib",
    "Template: https://github.com/scifig-ai/scientific-poster-and-presentation-templates"
  ]);
}

await fs.mkdir(BUILD,{recursive:true}); await fs.mkdir(OUT,{recursive:true});
const candidate = path.join(BUILD,"darkfield_candidate.pptx");
await (await PresentationFile.exportPptx(presentation)).save(candidate);

const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR,"container_tools/artifact_tool_utils.mjs")).href);
const stagingDir = path.join(BUILD,".codex-finalizer"); await fs.mkdir(stagingDir,{recursive:true});
const result = await finalizePresentation({
  workspaceDir: WORKSPACE,
  candidatePath: candidate,
  finalPath: FINAL,
  pythonExecutable: RUNTIME_PYTHON,
  integrityValidatorPath: path.join(SKILL_DIR,"container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(SKILL_DIR,"container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs:["--expected-slide-size-emu","12192000,6858000","--validate-bullet-geometry","--validate-heading-fit"],
  explicitTotalSlideCount:20,
  requiredNativeTableOwnerSlides:[], requiredNativeChartOwnerSlides:[],
  fontPolicy:{basis:"design",families:[FONT_CN,FONT_EN]},
  verifyArtifactToolImport:true,
  receiptPath:path.join(stagingDir,"darkfield-v2.validation.json"),
});

const finalDeck = await PresentationFile.importPptx(await FileBlob.load(FINAL));
const renderDir=path.join(BUILD,"final-render"); await fs.mkdir(renderDir,{recursive:true});
const slides=Array.isArray(finalDeck.slides?.items)?finalDeck.slides.items:Array.from({length:finalDeck.slides.count},(_,i)=>finalDeck.slides.getItem(i));
for(let i=0;i<slides.length;i++){
  const png=await finalDeck.export({slide:slides[i],format:"png",scale:0.8});
  await fs.writeFile(path.join(renderDir,`slide-${String(i+1).padStart(2,"0")}.png`),new Uint8Array(await png.arrayBuffer()));
}
const inspect=await finalDeck.inspect({kind:"slide,textbox,shape,image,notes,layout",maxChars:250000});
await fs.writeFile(path.join(BUILD,"final-inspect.ndjson"),inspect.ndjson||"","utf8");
console.log(JSON.stringify({final:FINAL,slideCount:slides.length,renderDir,result},null,2));
