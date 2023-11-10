import ExcelJS from "exceljs";
import { IResult } from "mssql";

const workbook = new ExcelJS.Workbook();

export const readFile = async () => {
  console.log('Reading "isbn.xlsx"');
  try {
    await workbook.xlsx.readFile("isbn.xlsx");
  } catch (err) {
    throw new Error(
      "Could not read isbn.xlsx. Make sure the file is placed in the root directory of this project."
    );
  }
  const column = workbook.getWorksheet(1)?.getColumn(1);
  if (!column)
    throw new Error(
      "Could not retrieve column from worksheet. Make sure ISBNs exist in column A in the first worksheet in isbn.xlsx"
    );
  const isbns: string[] = [];
  column?.eachCell((cell) => {
    const value = cell.value;
    if (typeof value === "string") isbns.push(value.trim());
    if (typeof value === "number") isbns.push(value.toString());
  });
  return isbns.filter((value) => {
    return value.match(/^\d{13}$/);
  });
};

export const writeFile = async (result: IResult<any>) => {
  console.log('Writing results to worksheet "Result" in "isbn.xlsx"');
  workbook.removeWorksheet("Result");
  const newSheet = workbook.addWorksheet("Result", {
    properties: { defaultColWidth: 20 },
  });

  const resultRows = result.recordset.map((row) => [
    row.ISBN,
    row.Tittel,
    row.Forlag,
  ]);

  newSheet.addTable({
    name: "MyTable",
    ref: "A1",
    headerRow: true,
    columns: [{ name: "ISBN" }, { name: "Tittel" }, { name: "Forlag" }],
    rows: resultRows,
  });
  await workbook.xlsx.writeFile("isbn.xlsx");
};
