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

function runPandoc(inputPath, outputPath) {
  const res = spawnSync("pandoc", [inputPath, "-o", outputPath], {
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

function main() {
  const { input, output } = parseArgs(process.argv);

  if (!fs.existsSync(input)) {
    fail(`Input file not found: ${input}`);
  }

  checkCommandExists("pandoc", "Pandoc");

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
      originalContent,
      workDir,
      puppeteerConfigPath
    );

    tempMarkdownPath = path.join(workDir, "processed.md");
    fs.writeFileSync(tempMarkdownPath, processed, "utf8");

    runPandoc(tempMarkdownPath, output);
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

