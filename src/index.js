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
    columns: [{ name: "ISBN" }, { name: "Tittel" }, { name: "Forlag" }],
    rows: values.map((value) => [value, value, value]),
  });
  workbook.xlsx.writeFile("isbn.xlsx");
};

const getIsbnQuery = (isbns) => {
  return `SELECT i.Varenr AS ISBN, i.Title AS Tittel, s.Text AS Forlag
  FROM Item i
  JOIN ItemField f ON i.Item_ID = f.Item_ID AND f.FieldCode = '260'
  JOIN ItemSubField s ON f.ItemField_ID = s.ItemField_ID AND s.SubFieldCode = 'b'
  WHERE i.Varenr IN (${isbns.reduce(
    (acc, isbn, index) => `${acc}${index === 0 ? "" : ","}'${isbn}'`,
    ""
  )})
  AND i.Currentstatus = '0'`;
};

readFile()
  .then((values) => {
    const query = getIsbnQuery(values);
    console.log(query);
    writeFile(values);
  })
  .catch((error) => {
    console.warn(error);
  });
