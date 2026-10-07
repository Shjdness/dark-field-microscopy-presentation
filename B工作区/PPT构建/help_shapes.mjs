import path from "node:path";
import { pathToFileURL } from "node:url";
const helper = "C:/Users/Shjdshy/.codex/plugins/cache/openai-primary-runtime/presentations/26.915.20218/skills/presentations/container_tools/runtime_helpers.mjs";
const { importRuntimeModule } = await import(pathToFileURL(helper).href);
const { Presentation } = await importRuntimeModule("@oai/artifact-tool");
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
console.log(deck.help("*", { search: "slide.images.deleteAll|image.delete|slide.shapes.deleteAll|shape.delete", include: ["index", "examples", "notes"], maxChars: 20000 }).ndjson);
