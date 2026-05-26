import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Lenovo/Downloads/STYLE.xlsx";
const outputDir = "D:/New folder/Piece rate Calculator/outputs/style-creator";
const outputPath = `${outputDir}/STYLE-style-creator.xlsx`;

const palette = {
  navy: "#143D59",
  teal: "#1F7A8C",
  mint: "#BEE3DB",
  cream: "#F8F3E7",
  gold: "#D9A441",
  slate: "#52616B",
  white: "#FFFFFF",
  gray: "#EEF2F6",
  border: "#C7D2DC",
};

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sourceSheet = workbook.worksheets.getItemAt(0);
const sourceValues = sourceSheet.getUsedRange().values || [];

const styleSet = new Set();
for (let i = 1; i < sourceValues.length; i += 1) {
  const styleCode = sourceValues[i]?.[1];
  if (styleCode) {
    styleSet.add(String(styleCode).trim());
  }
}
const styleList = [...styleSet].sort((a, b) => a.localeCompare(b));

for (const sheetName of ["Style Creator", "Upload Update", "Lists"]) {
  try {
    workbook.worksheets.getItem(sheetName).delete();
  } catch {}
}

const listSheet = workbook.worksheets.add("Lists");
listSheet.showGridLines = false;
listSheet.getRange("A1:B1").values = [["Existing Style Numbers", "Mode Options"]];
listSheet.getRange(`A2:A${Math.max(styleList.length + 1, 2)}`).values = styleList.length
  ? styleList.map((style) => [style])
  : [[""]];
listSheet.getRange("B2:B3").values = [["Existing"], ["New"]];
listSheet.getRange("A1:B3").format = {
  fill: palette.navy,
  font: { bold: true, color: palette.white },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
listSheet.getRange(`A2:B${Math.max(styleList.length + 1, 3)}`).format = {
  fill: palette.cream,
  font: { color: palette.slate },
};
listSheet.getRange("A:B").format.columnWidthPx = 170;

const creatorSheet = workbook.worksheets.add("Style Creator");
creatorSheet.showGridLines = false;
creatorSheet.freezePanes.freezeRows(11);

creatorSheet.getRange("A1:G2").merge();
creatorSheet.getRange("A1").values = [["AUTOMATED STYLE CREATOR"]];
creatorSheet.getRange("A1:G2").format = {
  fill: palette.navy,
  font: { bold: true, color: palette.white, size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};

creatorSheet.getRange("A4:B9").values = [
  ["Style Option", "Existing"],
  ["Style Number", ""],
  ["Image Link", ""],
  ["Colour", ""],
  ["Capacity", ""],
  ["CMT Rate", ""],
];
creatorSheet.getRange("A10:B10").values = [["Total CMT", ""]];
creatorSheet.getRange("A4:A10").format = {
  fill: palette.teal,
  font: { bold: true, color: palette.white },
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
creatorSheet.getRange("B4:B10").format = {
  fill: palette.cream,
  font: { color: "#1E293B" },
};
creatorSheet.getRange("B10").formulas = [['=IF(AND(B8<>"",B9<>""),B8*B9,"")']];
creatorSheet.getRange("B8:B10").format.numberFormat = "0.00";
creatorSheet.getRange("B4").dataValidation = {
  rule: { type: "list", formula1: "Lists!$B$2:$B$3" },
};
creatorSheet.getRange("B5").dataValidation = {
  rule: { type: "list", formula1: `Lists!$A$2:$A$${Math.max(styleList.length + 1, 2)}` },
};

creatorSheet.getRange("D4:G10").merge();
creatorSheet.getRange("D4").values = [[
  "Use Existing or New, then enter image link, style number, colour, capacity, and CMT rate. Total CMT calculates automatically.",
]];
creatorSheet.getRange("D4:G10").format = {
  fill: palette.mint,
  font: { color: "#183B4A", bold: true },
  wrapText: true,
  horizontalAlignment: "left",
  verticalAlignment: "center",
};

creatorSheet.getRange("A12:G12").values = [[
  "Style Option",
  "Style Number",
  "Image Link",
  "Colour",
  "Capacity",
  "CMT Rate",
  "Total CMT",
]];
creatorSheet.getRange("A13:F32").values = Array.from({ length: 20 }, () => ["", "", "", "", "", ""]);
creatorSheet.getRange("G13:G32").formulas = Array.from({ length: 20 }, (_, index) => [
  `=IF(AND(E${13 + index}<>"",F${13 + index}<>""),E${13 + index}*F${13 + index},"")`,
]);
creatorSheet.getRange("A12:G32").format = {
  fill: palette.white,
  font: { color: "#1E293B" },
};
creatorSheet.getRange("A12:G12").format = {
  fill: palette.navy,
  font: { bold: true, color: palette.white },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
creatorSheet.getRange("A13:A32").dataValidation = {
  rule: { type: "list", formula1: "Lists!$B$2:$B$3" },
};
creatorSheet.getRange("B13:B32").dataValidation = {
  rule: { type: "list", formula1: `Lists!$A$2:$A$${Math.max(styleList.length + 1, 2)}` },
};
creatorSheet.getRange("E13:G32").format.numberFormat = "0.00";
creatorSheet.tables.add("A12:G32", true, "StyleCreatorTable");

const uploadSheet = workbook.worksheets.add("Upload Update");
uploadSheet.showGridLines = false;
uploadSheet.freezePanes.freezeRows(6);

uploadSheet.getRange("A1:E2").merge();
uploadSheet.getRange("A1").values = [["UPLOAD / UPDATE CMT RATES"]];
uploadSheet.getRange("A1:E2").format = {
  fill: palette.navy,
  font: { bold: true, color: palette.white, size: 18 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};

uploadSheet.getRange("A4:E4").values = [[
  "Style Number",
  "Operation Wise Rate",
  "Total CMT",
  "Updated By",
  "Remarks",
]];
uploadSheet.getRange("A5:D24").values = Array.from({ length: 20 }, () => ["", "", "", ""]);
uploadSheet.getRange("E5:E24").values = Array.from({ length: 20 }, () => [""]);
uploadSheet.getRange("A4:E4").format = {
  fill: palette.teal,
  font: { bold: true, color: palette.white },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
uploadSheet.getRange("A5:E24").format = {
  fill: palette.cream,
  font: { color: "#1E293B" },
};
uploadSheet.getRange("A5:A24").dataValidation = {
  rule: { type: "list", formula1: `Lists!$A$2:$A$${Math.max(styleList.length + 1, 2)}` },
};
uploadSheet.getRange("B5:C24").format.numberFormat = "0.00";
uploadSheet.tables.add("A4:E24", true, "UploadUpdateTable");

const widthMap = {
  A: 150,
  B: 150,
  C: 210,
  D: 130,
  E: 110,
  F: 110,
  G: 120,
};

for (const [col, width] of Object.entries(widthMap)) {
  creatorSheet.getRange(`${col}:${col}`).format.columnWidthPx = width;
}

const uploadWidthMap = {
  A: 150,
  B: 150,
  C: 120,
  D: 130,
  E: 220,
};
for (const [col, width] of Object.entries(uploadWidthMap)) {
  uploadSheet.getRange(`${col}:${col}`).format.columnWidthPx = width;
}

await fs.mkdir(outputDir, { recursive: true });
const exportFile = await SpreadsheetFile.exportXlsx(workbook);
await exportFile.save(outputPath);

const creatorInspect = await workbook.inspect({
  kind: "table",
  range: "Style Creator!A1:G20",
  include: "values,formulas",
  tableMaxRows: 20,
  tableMaxCols: 7,
});
const uploadInspect = await workbook.inspect({
  kind: "table",
  range: "Upload Update!A1:E12",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 5,
});
await fs.writeFile(
  "style-creator-verification.txt",
  `${creatorInspect.ndjson}\n${uploadInspect.ndjson}`,
  "utf8",
);

console.log(outputPath);
