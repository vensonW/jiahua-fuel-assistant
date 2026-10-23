// Read-only visual verification; does not author or export a workbook.
import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "file:///C:/Users/KSG/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const [sourcePath, outputPath] = process.argv.slice(2);
await fs.mkdir("outputs/qa", { recursive: true });
for (const [name, path] of [["source", sourcePath], ["result", outputPath]]) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
  const ranges = name === "source"
    ? [["加油明细记录", "A4:I9"], ["发票", "E2:J6"]]
    : [["加油明细记录", "U4:AB10"], ["发票", "AF2:AM8"]];
  for (let i = 0; i < ranges.length; i += 1) {
    const [sheetName, range] = ranges[i];
    const preview = await workbook.render({ sheetName, range, scale: 1.3, format: "png" });
    await fs.writeFile(`outputs/qa/${name}_${i + 1}.png`, new Uint8Array(await preview.arrayBuffer()));
  }
  console.log(`${name}: rendered both affected worksheets`);
}
