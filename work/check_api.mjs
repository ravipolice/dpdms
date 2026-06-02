import { Workbook } from "@oai/artifact-tool";

const workbook = Workbook.create();
console.log(await workbook.help("worksheet range formatting formulas tables row grouping data validation freeze panes conditional formatting chart"));
