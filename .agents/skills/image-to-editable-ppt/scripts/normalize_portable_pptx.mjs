#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { execFileSync } from "node:child_process";
import { createRequire } from "node:module";

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) continue;
    args[token.slice(2)] = argv[i + 1];
    i += 1;
  }
  if (!args.input || !args.output) {
    throw new Error("Usage: normalize_portable_pptx.mjs --input INPUT.pptx --output OUTPUT.pptx [--sharp-module PATH] [--allow-autofit true] [--min-native-connectors N] [--min-target-arrows N] [--min-native-tables N]");
  }
  return args;
}

function integerArg(args, name) {
  if (args[name] == null) return 0;
  const value = Number(args[name]);
  if (!Number.isInteger(value) || value < 0) throw new Error(`--${name} must be a non-negative integer.`);
  return value;
}

function normalizeDefaultTextProperties(xml, counters) {
  for (const tag of ["defRPr", "endParaRPr"]) {
    const paired = new RegExp(`<a:${tag}([^>]*)>([\\s\\S]*?)<\\/a:${tag}>`, "g");
    xml = xml.replace(paired, (_match, attrs, body) => {
      const nextAttrs = attrs
        .replace(/\sb="[01]"/g, () => { counters.neutralizedDefaultBold += 1; return ""; })
        .replace(/\ssz="\d+"/g, () => { counters.neutralizedDefaultFontSize += 1; return ""; });
      const nextBody = body.replace(/<a:(?:latin|ea|cs)\b[^>]*\/>/g, () => {
        counters.neutralizedDefaultFontFace += 1;
        return "";
      });
      return `<a:${tag}${nextAttrs}>${nextBody}</a:${tag}>`;
    });

    const selfClosing = new RegExp(`<a:${tag}([^>]*)\\/>`, "g");
    xml = xml.replace(selfClosing, (_match, attrs) => {
      const nextAttrs = attrs
        .replace(/\sb="[01]"/g, () => { counters.neutralizedDefaultBold += 1; return ""; })
        .replace(/\ssz="\d+"/g, () => { counters.neutralizedDefaultFontSize += 1; return ""; });
      return `<a:${tag}${nextAttrs}/>`;
    });
  }
  return xml;
}

async function walk(dir) {
  const out = [];
  for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...await walk(full));
    else out.push(full);
  }
  return out;
}

function loadSharp(explicitPath) {
  const require = createRequire(import.meta.url);
  const candidates = [
    explicitPath,
    process.env.CODEX_NODE_MODULES ? path.join(process.env.CODEX_NODE_MODULES, "sharp") : null,
    "sharp",
  ].filter(Boolean);
  for (const candidate of candidates) {
    try { return require(candidate); } catch {}
  }
  throw new Error("Cannot load sharp. Pass --sharp-module using the bundled workspace dependency path.");
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const input = path.resolve(args.input);
  const output = path.resolve(args.output);
  if (input === output) throw new Error("Input and output must differ; preserve the source PPTX.");

  const sharp = loadSharp(args["sharp-module"]);
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "editable-ppt-"));
  try {
    execFileSync("unzip", ["-q", input, "-d", tempRoot]);
    const allFiles = await walk(tempRoot);
    const slideXmlFiles = allFiles.filter((file) => /\/ppt\/slides\/slide\d+\.xml$/.test(file));
    let removedGroupLocks = 0;
    const counters = {
      neutralizedDefaultBold: 0,
      neutralizedDefaultFontSize: 0,
      neutralizedDefaultFontFace: 0,
    };
    let nativeConnectorCount = 0;
    let targetArrowCount = 0;
    let nativeTableCount = 0;
    let autoFitCount = 0;

    for (const file of slideXmlFiles) {
      let xml = await fs.readFile(file, "utf8");
      xml = xml.replace(/<a:spLocks noGrp="1"[^>]*\/>/g, () => {
        removedGroupLocks += 1;
        return "";
      });
      xml = normalizeDefaultTextProperties(xml, counters);
      xml = xml.replace(/<a:rPr([^>]*)\sb="0"([^>]*)>/g, "<a:rPr$1$2>");
      const connectors = xml.match(/<p:cxnSp\b[\s\S]*?<\/p:cxnSp>/g) || [];
      nativeConnectorCount += connectors.length;
      targetArrowCount += connectors.filter((connector) => /<a:tailEnd\b[^>]*type="(?:triangle|arrow|stealth)"/.test(connector)).length;
      nativeTableCount += (xml.match(/<a:tbl\b/g) || []).length;
      autoFitCount += (xml.match(/<a:(?:normAutofit|spAutoFit)\b/g) || []).length;
      await fs.writeFile(file, xml);
    }

    const allowAutoFit = args["allow-autofit"] === "true";
    if (autoFitCount && !allowAutoFit) {
      throw new Error(`Found ${autoFitCount} automatic text-fit tag(s). Remove them so manual font-size edits remain effective, or pass --allow-autofit true when explicitly required.`);
    }
    const minNativeConnectors = integerArg(args, "min-native-connectors");
    const minTargetArrows = integerArg(args, "min-target-arrows");
    const minNativeTables = integerArg(args, "min-native-tables");
    if (nativeConnectorCount < minNativeConnectors) {
      throw new Error(`Found ${nativeConnectorCount} native connector(s), expected at least ${minNativeConnectors}. Rebuild semantic arrows with slide.shapes.connect().`);
    }
    if (targetArrowCount < minTargetArrows) {
      throw new Error(`Found ${targetArrowCount} target arrowhead(s), expected at least ${minTargetArrows}. Put the target arrowhead on the connector tail.`);
    }
    if (nativeTableCount < minNativeTables) {
      throw new Error(`Found ${nativeTableCount} native table(s), expected at least ${minNativeTables}. Rebuild actual tabular data with slide.tables.add() instead of independent rectangles and lines.`);
    }

    const relFiles = allFiles.filter((file) => file.endsWith(".rels"));
    let externalMediaLinks = 0;
    for (const file of relFiles) {
      const xml = await fs.readFile(file, "utf8");
      externalMediaLinks += (xml.match(/Type="[^"]*\/image"[^>]*TargetMode="External"/g) || []).length;
    }
    if (externalMediaLinks) {
      throw new Error(`Found ${externalMediaLinks} external image relationship(s). Embed those assets before delivery.`);
    }

    const mediaDir = path.join(tempRoot, "ppt", "media");
    let svgCount = 0;
    let pngFallbackCount = 0;
    try {
      const media = await fs.readdir(mediaDir);
      for (const svgName of media.filter((name) => name.toLowerCase().endsWith(".svg"))) {
        svgCount += 1;
        const pngName = `${svgName.slice(0, -4)}.png`;
        if (!media.includes(pngName)) continue;
        const svgPath = path.join(mediaDir, svgName);
        const pngPath = path.join(mediaDir, pngName);
        const nextPath = `${pngPath}.new`;
        await sharp(svgPath, { density: 768 })
          .resize(1024, 1024, { fit: "contain", background: { r: 255, g: 255, b: 255, alpha: 0 } })
          .png({ compressionLevel: 9 })
          .toFile(nextPath);
        await fs.rename(nextPath, pngPath);
        pngFallbackCount += 1;
      }
    } catch (error) {
      if (error?.code !== "ENOENT") throw error;
    }

    await fs.mkdir(path.dirname(output), { recursive: true });
    await fs.rm(output, { force: true });
    execFileSync("zip", ["-q", "-r", output, "."], { cwd: tempRoot });
    execFileSync("unzip", ["-t", output], { stdio: "ignore" });

    console.log(JSON.stringify({
      output,
      slides: slideXmlFiles.length,
      removedGroupLocks,
      ...counters,
      nativeConnectorCount,
      targetArrowCount,
      nativeTableCount,
      autoFitCount,
      externalMediaLinks,
      svgCount,
      pngFallbackCount,
    }, null, 2));
  } finally {
    await fs.rm(tempRoot, { recursive: true, force: true });
  }
}

main().catch((error) => {
  console.error(error?.stack || String(error));
  process.exitCode = 1;
});
