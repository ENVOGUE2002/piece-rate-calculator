import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "D:/New folder/Piece rate Calculator/outputs/style-creator/STYLE-style-creator.xlsx";
const outputDir = "D:/New folder/Piece rate Calculator/outputs/style-creator/renders";

const blob = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(blob);

await fs.mkdir(outputDir, { recursive: true });

for (const sheetName of ["Style Creator", "Upload Update"]) {
  const image = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  const bytes = new Uint8Array(await image.arrayBuffer());
  await fs.writeFile(`${outputDir}/${sheetName.replaceAll(" ", "_")}.png`, bytes);
  console.log(`rendered ${sheetName}`);
}
