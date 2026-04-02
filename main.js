#!/usr/bin/env node

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");

function printUsage() {
  console.log("Usage: md2docx <input.md> [output.docx]");
}

function fail(message) {
  console.error(`Error: ${message}`);
  process.exit(1);
}

function checkCommandExists(cmd, displayName) {
  const result = spawnSync(cmd, ["--version"], {
    stdio: "ignore",
  });

  if (result.error && result.error.code === "ENOENT") {
    fail(
      `${displayName} ('${cmd}') not found.\n` +
        `Please install it first.\n` +
        (cmd === "mmdc"
          ? "Install with: npm install -g @mermaid-js/mermaid-cli"
          : "See: https://pandoc.org/installing.html")
    );
  }

  if (result.status !== 0) {
    fail(`${displayName} ('${cmd}') is not working correctly.`);
  }
}

function getMermaidCliCommand() {
  if (process.env.MD2DOCX_MMDC) {
    return process.env.MD2DOCX_MMDC;
  }

  const localBin = path.join(__dirname, "node_modules", ".bin", "mmdc");
  if (fs.existsSync(localBin)) {
    return localBin;
  }

  return "mmdc";
}

function parseArgs(argv) {
  const args = argv.slice(2);
  if (args.length === 0 || args.includes("-h") || args.includes("--help")) {
    printUsage();
    process.exit(args.length === 0 ? 1 : 0);
  }

  const input = args[0];
  const output =
    args[1] ||
    path.join(
      path.dirname(input),
      `${path.basename(input, path.extname(input))}.docx`
    );

  return { input, output };
}

function sanitizeMermaidForRetry(source) {
  return source.replace(/([A-Za-z0-9_]+)\[([^\]]*)\]/g, (full, id, label) => {
    const cleaned = label
      .replace(/<br\s*\/?>/gi, "<br/>")
      .replace(/"/g, '\\"');
    return `${id}["${cleaned}"]`;
  });
}

function renderDiagram(mmdcCmd, mmdPath, imgPath, puppeteerConfigFile) {
  const args = ["-i", mmdPath, "-o", imgPath, "-b", "white"];
  if (puppeteerConfigFile) {
    args.push("-p", puppeteerConfigFile);
  }
  return spawnSync(mmdcCmd, args, { encoding: "utf8" });
}

function parsePipeCells(line) {
  const t = line.trim();
  if (!t.includes("|")) return [];
  return t
    .split("|")
    .map((s) => s.trim())
    .filter(Boolean);
}

function lineToGfmPipeRow(line) {
  const t = line.trim();
  if (!t.includes("|")) return line;
  if (t.startsWith("|")) return line;
  const parts = parsePipeCells(t);
  if (parts.length < 2) return line;
  return `| ${parts.join(" | ")} |`;
}

function isGfmSeparatorRow(line) {
  const parts = parsePipeCells(line);
  if (parts.length < 2) return false;
  return parts.every((p) => /^:?-{3,}:?$/.test(p));
}

function lineLooksLikePipeTableRow(line) {
  return parsePipeCells(line).length >= 2;
}

function firstPipeCellIsNumeric(line) {
  const cells = parsePipeCells(line);
  if (!cells.length) return false;
  return /^\d+$/.test(cells[0]);
}

function tryMergeStackedHeadersAt(lines, i) {
  let j = i;
  const headers = [];
  while (
    j < lines.length &&
    lines[j].trim() !== "" &&
    !lines[j].includes("|")
  ) {
    headers.push(lines[j].trim());
    j++;
  }
  if (headers.length < 2) return null;
  if (j >= lines.length || lines[j].trim() !== "") return null;
  j++;
  if (j >= lines.length) return null;
  const normalized = lineToGfmPipeRow(lines[j]);
  if (!normalized.trim().startsWith("|")) return null;
  const cells = parsePipeCells(normalized);
  if (cells.length !== headers.length) return null;

  const newLines = [];
  newLines.push(`| ${headers.join(" | ")} |`);
  newLines.push(`| ${headers.map(() => "---").join(" | ")} |`);
  let k = j;
  while (k < lines.length) {
    const L = lines[k];
    if (L.trim() === "") break;
    const n = lineToGfmPipeRow(L);
    if (!n.trim().startsWith("|")) break;
    newLines.push(n);
    k++;
  }
  return { consumed: k - i, newLines };
}

function mergeStackedHeaders(lines) {
  const out = [];
  let i = 0;
  let inFence = false;
  while (i < lines.length) {
    const trimmed = lines[i].trim();
    if (trimmed.startsWith("```")) {
      inFence = !inFence;
      out.push(lines[i]);
      i++;
      continue;
    }
    if (inFence) {
      out.push(lines[i]);
      i++;
      continue;
    }
    const merged = tryMergeStackedHeadersAt(lines, i);
    if (merged) {
      out.push(...merged.newLines);
      i += merged.consumed;
      continue;
    }
    out.push(lines[i]);
    i++;
  }
  return out;
}

function mapNonFenceLines(lines, fn) {
  let inFence = false;
  return lines.map((line) => {
    const t = line.trim();
    if (t.startsWith("```")) {
      inFence = !inFence;
      return line;
    }
    if (inFence) return line;
    return fn(line);
  });
}

function insertMissingGfmSeparators(lines) {
  const out = [];
  let inFence = false;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();
    if (trimmed.startsWith("```")) {
      inFence = !inFence;
      out.push(line);
      continue;
    }
    if (inFence) {
      out.push(line);
      continue;
    }
    out.push(line);
    const prev = i > 0 ? lines[i - 1] : "";
    const next = lines[i + 1];
    if (!next || next.trim() === "") continue;
    if (
      !lineLooksLikePipeTableRow(line) ||
      !lineLooksLikePipeTableRow(next)
    ) {
      continue;
    }
    if (isGfmSeparatorRow(line) || isGfmSeparatorRow(next)) continue;
    if (prev.trim() !== "" && lineLooksLikePipeTableRow(prev)) continue;
    if (firstPipeCellIsNumeric(line) && firstPipeCellIsNumeric(next)) continue;
    const n = parsePipeCells(line).length;
    const m = parsePipeCells(next).length;
    if (n < 2 || m !== n) continue;
    out.push(`| ${Array(n).fill("---").join(" | ")} |`);
  }
  return out;
}

function normalizeMarkdownForGfmTables(content) {
  let lines = content.split(/\r?\n/);
  lines = mergeStackedHeaders(lines);
  lines = mapNonFenceLines(lines, lineToGfmPipeRow);
  lines = insertMissingGfmSeparators(lines);
  return lines.join("\n");
}

function processMarkdown(content, workDir, puppeteerConfigFile) {
  const lines = content.split(/\r?\n/);
  const outLines = [];

  let insideMermaid = false;
  let mermaidLines = [];
  let diagramIndex = 0;
  const mmdcCmd = getMermaidCliCommand();

  for (const line of lines) {
    if (!insideMermaid) {
      const trimmed = line.trim();
      if (/^```+ *mermaid\b/i.test(trimmed)) {
        insideMermaid = true;
        mermaidLines = [];
        diagramIndex += 1;
        continue;
      }
      outLines.push(line);
    } else {
      if (/^```+/.test(line.trim())) {
        const mmdPath = path.join(workDir, `diagram-${diagramIndex}.mmd`);
        const imgPath = path.join(workDir, `diagram-${diagramIndex}.png`);

        fs.writeFileSync(mmdPath, mermaidLines.join("\n"), "utf8");

        let res = renderDiagram(
          mmdcCmd,
          mmdPath,
          imgPath,
          puppeteerConfigFile
        );

        if (res.error && res.error.code === "ENOENT") {
          fail(
            "Mermaid CLI ('mmdc') not found while rendering.\n" +
              "Install locally in this project with:\n" +
              "  npm install @mermaid-js/mermaid-cli\n" +
              "or install globally with:\n" +
              "  npm install -g @mermaid-js/mermaid-cli"
          );
        }
        if (res.status !== 0) {
          const sanitizedMmd = sanitizeMermaidForRetry(mermaidLines.join("\n"));
          if (sanitizedMmd !== mermaidLines.join("\n")) {
            fs.writeFileSync(mmdPath, sanitizedMmd, "utf8");
            res = renderDiagram(mmdcCmd, mmdPath, imgPath, puppeteerConfigFile);
          }
        }

        if (res.status !== 0) {
          fail(
            `Mermaid CLI failed for diagram ${diagramIndex}.\n` +
              (res.stderr || "")
          );
        }

        const label = `Mermaid diagram ${diagramIndex}`;
        outLines.push(`![${label}](${imgPath})`);

        insideMermaid = false;
        mermaidLines = [];
      } else {
        mermaidLines.push(line);
      }
    }
  }

  if (insideMermaid) {
    fail("Unclosed ```mermaid code block detected.");
  }

  return outLines.join("\n");
}

function runPandoc(inputPath, outputPath, { resourcePaths = [] } = {}) {
  const args = [
    "--from=gfm",
    "--to=docx",
    inputPath,
    "-o",
    outputPath,
  ];

  if (resourcePaths.length > 0) {
    const value = resourcePaths.join(path.delimiter);
    args.push(`--resource-path=${value}`);
  }

  const res = spawnSync("pandoc", args, {
    encoding: "utf8",
  });

  if (res.error && res.error.code === "ENOENT") {
    fail(
      "Pandoc ('pandoc') not found.\n" +
        "Install from: https://pandoc.org/installing.html"
    );
  }

  if (res.status !== 0) {
    fail(`Pandoc failed.\n${res.stderr || ""}`);
  }
}

function simplifyPicSpPrForGoogleDocs(xml) {
  return xml.replace(/<pic:spPr\b[^>]*>[\s\S]*?<\/pic:spPr>/gi, (block) => {
    const close = block.lastIndexOf("</pic:spPr>");
    const openEnd = block.indexOf(">");
    if (openEnd === -1 || close === -1) return block;
    const inner = block.slice(openEnd + 1, close);
    const xfrm = inner.match(/<a:xfrm>[\s\S]*?<\/a:xfrm>/)?.[0];
    const prstGeom = inner.match(/<a:prstGeom\b[^>]*>[\s\S]*?<\/a:prstGeom>/)?.[0];
    if (!xfrm && !prstGeom) return block;
    return `<pic:spPr>${xfrm || ""}${prstGeom || ""}</pic:spPr>`;
  });
}

function listWordPartsForImagePatch(wordDir) {
  const names = new Set([
    "document.xml",
    "footnotes.xml",
    "endnotes.xml",
    "comments.xml",
  ]);
  if (!fs.existsSync(wordDir)) return [];
  for (const name of fs.readdirSync(wordDir)) {
    if (/^header\d+\.xml$/i.test(name) || /^footer\d+\.xml$/i.test(name)) {
      names.add(name);
    }
  }
  return [...names].map((n) => path.join(wordDir, n)).filter((p) => fs.existsSync(p));
}

function normalizeDocxLayout(outputDocxPath) {
  const unzip = spawnSync("unzip", ["-v"], { stdio: "ignore" });
  if (unzip.error && unzip.error.code === "ENOENT") return;
  const zip = spawnSync("zip", ["-v"], { stdio: "ignore" });
  if (zip.error && zip.error.code === "ENOENT") return;

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "md2docx-docx-"));
  try {
    const extract = spawnSync("unzip", ["-q", outputDocxPath, "-d", tempDir], {
      encoding: "utf8",
    });
    if (extract.status !== 0) return;

    const stylesPath = path.join(tempDir, "word", "styles.xml");
    const documentPath = path.join(tempDir, "word", "document.xml");
    const wordDir = path.join(tempDir, "word");
    if (!fs.existsSync(documentPath)) return;

    let stylesChanged = false;
    if (fs.existsSync(stylesPath)) {
      const styles = fs.readFileSync(stylesPath, "utf8");
      const valMatch = styles.match(/<w:lang\b[^>]*\bw:val="([^"]+)"/);
      const langVal = valMatch?.[1];
      if (langVal) {
        const updated = styles.replace(
          /(<w:lang\b[^>]*?)\bw:bidi="[^"]*"/g,
          `$1w:bidi="${langVal}"`
        );
        if (updated !== styles) {
          fs.writeFileSync(stylesPath, updated, "utf8");
          stylesChanged = true;
        }
      }
    }

    const zipEntries = [];
    let doc = fs.readFileSync(documentPath, "utf8");
    doc = doc.replace(/<w:mirrorMargins\s*\/>/g, "");
    doc = doc.replace(/<w:bookFoldPrinting\s*\/>/g, "");
    doc = doc.replace(/<w:bookFoldRevPrinting\s*\/>/g, "");

    doc = doc.replace(/(<w:sectPr\b[^>]*>)/, (m) => {
      if (doc.includes("<w:pgSz") || doc.includes("<w:pgMar")) return m;
      const pgSz = '<w:pgSz w:w="11906" w:h="16838" />';
      const pgMar =
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0" />';
      return `${m}${pgSz}${pgMar}`;
    });

    doc = simplifyPicSpPrForGoogleDocs(doc);
    fs.writeFileSync(documentPath, doc, "utf8");
    zipEntries.push("word/document.xml");

    for (const partPath of listWordPartsForImagePatch(wordDir)) {
      if (partPath === documentPath) continue;
      let part = fs.readFileSync(partPath, "utf8");
      if (!part.includes("<pic:spPr")) continue;
      const next = simplifyPicSpPrForGoogleDocs(part);
      if (next !== part) {
        fs.writeFileSync(partPath, next, "utf8");
        zipEntries.push(path.relative(tempDir, partPath).replace(/\\/g, "/"));
      }
    }

    if (stylesChanged) {
      zipEntries.push("word/styles.xml");
    }

    spawnSync("zip", ["-q", "-u", outputDocxPath, ...[...new Set(zipEntries)]], {
      cwd: tempDir,
      encoding: "utf8",
    });
  } finally {
    try {
      fs.rmSync(tempDir, { recursive: true, force: true });
    } catch (e) {
      // ignore cleanup errors
    }
  }
}

function main() {
  const { input, output } = parseArgs(process.argv);

  if (!fs.existsSync(input)) {
    fail(`Input file not found: ${input}`);
  }

  checkCommandExists("pandoc", "Pandoc");

  const outputAbs = path.resolve(output);
  const originalContent = fs.readFileSync(input, "utf8");

  const workDir = fs.mkdtempSync(
    path.join(os.tmpdir(), "md2docx-").replace(/[/\\]$/, "")
  );

  let tempMarkdownPath;
  try {
    const puppeteerConfigPath = path.join(workDir, "puppeteer-config.json");
    fs.writeFileSync(
      puppeteerConfigPath,
      JSON.stringify(
        {
          args: ["--no-sandbox", "--disable-setuid-sandbox"],
        },
        null,
        2
      ),
      "utf8"
    );

    const processed = processMarkdown(
      normalizeMarkdownForGfmTables(originalContent),
      workDir,
      puppeteerConfigPath
    );

    tempMarkdownPath = path.join(workDir, "processed.md");
    fs.writeFileSync(tempMarkdownPath, processed, "utf8");

    runPandoc(tempMarkdownPath, outputAbs, {
      resourcePaths: [path.dirname(path.resolve(input)), workDir],
    });
    normalizeDocxLayout(outputAbs);
    console.log(`Wrote: ${outputAbs}`);
  } finally {
    try {
      fs.rmSync(workDir, { recursive: true, force: true });
    } catch (e) {
      // ignore cleanup errors
    }
  }
}

if (require.main === module) {
  main();
}

