const ExcelJS = require("exceljs");

const workbook = new ExcelJS.Workbook();

const readFile = async () => {
  await workbook.xlsx.readFile("isbn.xlsx");
  const column = workbook.getWorksheet("Sheet1").getColumn(1);
  const isbns = [];
  column.eachCell((cell) => {
    isbns.push(cell._value.model.value.trim());
  });
  return isbns.filter((value) => {
    return value.match(/^\d{13}$/);
  });
};

const writeFile = (values) => {
  workbook.removeWorksheet("Result");
  const newSheet = workbook.addWorksheet("Result", {
    properties: { defaultColWidth: 20 },
  });
  newSheet.addTable({
    name: "MyTable",
    ref: "A1",
    headerRow: true,
    columns: [{ name: "1" }, { name: "2" }, { name: "3" }],
    rows: values.map((value) => [value, value, value]),
  });
  workbook.xlsx.writeFile("isbn.xlsx");
};

const isbns = readFile()
  .then((values) => {
    console.log(values);
    writeFile(values);
  })
  .catch((error) => {
    console.warn(error);
  });
