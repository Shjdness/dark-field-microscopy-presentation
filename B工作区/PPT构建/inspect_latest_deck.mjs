import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const workspace = "C:/Users/Shjdshy/Desktop/AllDeskWork/暗场显微镜PPT_DDL100818pm";
const source = path.join(workspace, "C已完成/暗场显微镜_课堂汇报_结构讲解优化版_v3.pptx");
const outDir = path.join(workspace, "B工作区/PPT构建/latest-inspect");
const helper = "C:/Users/Shjdshy/.codex/plugins/cache/openai-primary-runtime/presentations/26.915.20218/skills/presentations/container_tools/runtime_helpers.mjs";
const { importRuntimeModule } = await import(pathToFileURL(helper).href);
const { FileBlob, PresentationFile } = await importRuntimeModule("@oai/artifact-tool");
const deck = await PresentationFile.importPptx(await FileBlob.load(source));
await fs.mkdir(outDir, { recursive: true });
for (const i of [0, 5, 6, 7]) {
  const slide = deck.slides.getItem(i);
  const png = await deck.export({ slide, format: "png", scale: 1.2 });
  await fs.writeFile(path.join(outDir, `slide-${i + 1}.png`), new Uint8Array(await png.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(outDir, `slide-${i + 1}.layout.json`), await layout.text(), "utf8");
}
const inspect = await deck.inspect({ kind: "slide,textbox,shape,image,notes,layout", maxChars: 250000 });
await fs.writeFile(path.join(outDir, "inspect.ndjson"), inspect.ndjson || "", "utf8");
console.log(JSON.stringify({ source, slideCount: deck.slides.count, outDir }));
