import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const workspace = "C:/Users/Shjdshy/Desktop/工作桌面文件夹/暗场显微镜PPT_DDL100818pm";
const source = path.join(workspace, "B工作区/模板/SciFig_Conference_Minimal_Navy.pptx");
const outDir = path.join(workspace, "B工作区/PPT构建/template-render");
const helper = "C:/Users/Shjdshy/.codex/plugins/cache/openai-primary-runtime/presentations/26.915.20218/skills/presentations/container_tools/runtime_helpers.mjs";

const { importRuntimeModule } = await import(pathToFileURL(helper).href);
const { FileBlob, PresentationFile } = await importRuntimeModule("@oai/artifact-tool");
const deck = await PresentationFile.importPptx(await FileBlob.load(source));
await fs.mkdir(outDir, { recursive: true });
const slides = Array.isArray(deck.slides?.items)
  ? deck.slides.items
  : Array.from({ length: deck.slides.count }, (_, i) => deck.slides.getItem(i));
for (let i = 0; i < slides.length; i += 1) {
  const blob = await deck.export({ slide: slides[i], format: "png", scale: 0.7 });
  const bytes = new Uint8Array(await blob.arrayBuffer());
  await fs.writeFile(path.join(outDir, `slide-${String(i + 1).padStart(2, "0")}.png`), bytes);
}
const inspect = await deck.inspect({ kind: "slide,textbox,shape,image,table,chart,notes,layout", maxChars: 200000 });
await fs.writeFile(path.join(outDir, "inspect.ndjson"), inspect.ndjson || "", "utf8");
console.log(JSON.stringify({ slideCount: slides.length, outDir }));
