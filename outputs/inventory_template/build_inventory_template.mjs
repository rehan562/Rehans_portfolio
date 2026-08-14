import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outputDir = "D:/My website/Rehans_portfolio/outputs/inventory_template";
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const inventory = workbook.worksheets.add("Inventory Register");
const lists = workbook.worksheets.add("Lists");

const navy = "#17365D";
const blue = "#DCE6F1";
const teal = "#0F766E";
const paleTeal = "#E6F4F1";
const paleYellow = "#FFF2CC";
const lightBorder = "#D9E2F3";

inventory.showGridLines = false;
inventory.getRange("A1:N1").merge();
inventory.getRange("A1").values = [["Inventory Register"]];
inventory.getRange("A1:N1").format = {
  fill: navy,
  font: { bold: true, color: "#FFFFFF", size: 18 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
inventory.getRange("A1:N1").format.rowHeight = 32;

inventory.getRange("A2:N2").merge();
inventory.getRange("A2").values = [["Enter one item per row. Stock Status and Stock Value update automatically."]];
inventory.getRange("A2:N2").format = {
  font: { italic: true, color: "#4A5568", size: 10 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
};
inventory.getRange("A2:N2").format.rowHeight = 22;

inventory.getRange("A4:B4").merge();
inventory.getRange("D4:E4").merge();
inventory.getRange("G4:H4").merge();
inventory.getRange("J4:K4").merge();
inventory.getRange("M4:N4").merge();
inventory.getRange("A4").values = [["Total Items"]];
inventory.getRange("D4").values = [["Units in Stock"]];
inventory.getRange("G4").values = [["Inventory Value"]];
inventory.getRange("J4").values = [["Low Stock Items"]];
inventory.getRange("M4").values = [["Out of Stock"]];
inventory.getRange("A5:B5").merge();
inventory.getRange("D5:E5").merge();
inventory.getRange("G5:H5").merge();
inventory.getRange("J5:K5").merge();
inventory.getRange("M5:N5").merge();
inventory.getRange("A5").formulas = [["=COUNTIF($B$8:$B$107,\"<>\")"]];
inventory.getRange("D5").formulas = [["=SUM($H$8:$H$107)"]];
inventory.getRange("G5").formulas = [["=SUM($M$8:$M$107)"]];
inventory.getRange("J5").formulas = [["=COUNTIF($N$8:$N$107,\"Low Stock\")"]];
inventory.getRange("M5").formulas = [["=COUNTIF($N$8:$N$107,\"Out of Stock\")"]];
for (const range of ["A4:B4", "D4:E4", "G4:H4", "J4:K4", "M4:N4"]) {
  inventory.getRange(range).format = { fill: blue, font: { bold: true, color: navy }, horizontalAlignment: "center", verticalAlignment: "center", borders: { preset: "outside", style: "thin", color: lightBorder } };
}
for (const range of ["A5:B5", "D5:E5", "G5:H5", "J5:K5", "M5:N5"]) {
  inventory.getRange(range).format = { fill: "#F8FBFF", font: { bold: true, color: navy, size: 14 }, horizontalAlignment: "center", verticalAlignment: "center", borders: { preset: "outside", style: "thin", color: lightBorder } };
}
inventory.getRange("A4:N5").format.rowHeight = 22;
inventory.getRange("G5").format.numberFormat = "$#,##0.00";

const headers = [["Item ID", "Item Name", "Category", "Description", "Supplier", "Location", "Unit", "Quantity in Stock", "Reorder Level", "Unit Cost", "Selling Price", "Last Restocked", "Stock Value", "Stock Status"]];
inventory.getRange("A7:N7").values = headers;
inventory.getRange("A7:N7").format = { fill: teal, font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
inventory.getRange("A7:N7").format.rowHeight = 34;

inventory.getRange("A8:N107").format = { borders: { preset: "inside", style: "thin", color: "#E7EDF5" }, verticalAlignment: "center" };
inventory.getRange("A8:A107").format.numberFormat = "@";
inventory.getRange("H8:I107").format.numberFormat = "#,##0";
inventory.getRange("J8:K107").format.numberFormat = "$#,##0.00";
inventory.getRange("L8:L107").format.numberFormat = "yyyy-mm-dd";
inventory.getRange("M8:M107").format.numberFormat = "$#,##0.00";
inventory.getRange("M8").formulas = [["=IF(OR(H8=\"\",J8=\"\"),\"\",H8*J8)"]];
inventory.getRange("M8:M107").fillDown();
inventory.getRange("N8").formulas = [["=IF(B8=\"\",\"\",IF(H8=0,\"Out of Stock\",IF(H8<=I8,\"Low Stock\",\"In Stock\")))"]];
inventory.getRange("N8:N107").fillDown();

inventory.getRange("C8:C107").dataValidation = { rule: { type: "list", formula1: "'Lists'!$A$2:$A$9" } };
inventory.getRange("G8:G107").dataValidation = { rule: { type: "list", formula1: "'Lists'!$B$2:$B$8" } };
inventory.getRange("H8:H107").dataValidation = { rule: { type: "whole", operator: "greaterThanOrEqual", formula1: 0 } };
inventory.getRange("I8:I107").dataValidation = { rule: { type: "whole", operator: "greaterThanOrEqual", formula1: 0 } };
inventory.getRange("J8:K107").dataValidation = { rule: { type: "decimal", operator: "greaterThanOrEqual", formula1: 0 } };
inventory.getRange("N8:N107").conditionalFormats.add("containsText", { text: "Out of Stock", format: { fill: "#FDE2E2", font: { color: "#B91C1C", bold: true } } });
inventory.getRange("N8:N107").conditionalFormats.add("containsText", { text: "Low Stock", format: { fill: paleYellow, font: { color: "#92400E", bold: true } } });
inventory.getRange("N8:N107").conditionalFormats.add("containsText", { text: "In Stock", format: { fill: paleTeal, font: { color: "#166534", bold: true } } });

const widths = [14, 24, 16, 30, 21, 16, 11, 16, 14, 13, 14, 15, 14, 16];
for (let i = 0; i < widths.length; i++) inventory.getRangeByIndexes(0, i, 108, 1).format.columnWidth = widths[i];
inventory.freezePanes.freezeRows(7);
inventory.freezePanes.freezeColumns(2);
const table = inventory.tables.add("A7:N107", true, "InventoryTable");
table.style = "TableStyleMedium2";

lists.showGridLines = false;
lists.getRange("A1:B1").values = [["Categories", "Units"]];
lists.getRange("A2:A9").values = [["Electronics"], ["Office Supplies"], ["Furniture"], ["Tools"], ["Raw Materials"], ["Packaging"], ["Cleaning Supplies"], ["Other"]];
lists.getRange("B2:B8").values = [["Each"], ["Box"], ["Pack"], ["Kg"], ["Litre"], ["Metre"], ["Set"]];
lists.getRange("A1:B1").format = { fill: navy, font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
lists.getRange("A1:B9").format.borders = { preset: "inside", style: "thin", color: "#E7EDF5" };
lists.getRange("A:B").format.columnWidth = 22;

const check = await workbook.inspect({ kind: "table", range: "Inventory Register!A1:N15", include: "values,formulas", tableMaxRows: 15, tableMaxCols: 14 });
console.log(check.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 100 }, summary: "formula error scan" });
console.log(errors.ndjson);
const preview = await workbook.render({ sheetName: "Inventory Register", range: "A1:N18", scale: 1.4, format: "png" });
await fs.writeFile(`${outputDir}/inventory_preview.png`, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/inventory_template.xlsx`);
