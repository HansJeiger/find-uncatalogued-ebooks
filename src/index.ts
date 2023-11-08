import sql, { config as SqlConfig } from "mssql";
import ExcelJS from "exceljs";
import dotenv from "dotenv";
import { z } from "zod";

dotenv.config({ debug: true });

const Config = z.object({
  promus_user: z.string(),
  promus_password: z.string(),
  promus_database: z.string(),
  promus_server: z.string(),
  promus_port: z.coerce.number(),
});

const config = Config.parse(process.env);

const workbook = new ExcelJS.Workbook();

const readFile = async () => {
  await workbook.xlsx.readFile("isbn.xlsx");
  const column = workbook.getWorksheet("Sheet1")?.getColumn(1);
  if (!column) throw new Error("Could not retrieve column from worksheet.");
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

const writeFile = (values: string[]) => {
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

const getIsbnQuery = (isbns: string[]) => {
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

const connectToPromus = async (query: string) => {
  const sqlConfig: SqlConfig = {
    user: config.promus_user,
    password: config.promus_password,
    database: config.promus_database,
    server: config.promus_server,
    port: config.promus_port,
    pool: {
      max: 10,
      min: 0,
      idleTimeoutMillis: 30000,
    },
    options: {
      encrypt: true, // for azure
      trustServerCertificate: true, // change to true for local dev / self-signed certs
    },
  };

  try {
    // make sure that any items are correctly URL encoded in the connection string
    await sql.connect(
      sqlConfig //"Server=wg-sxd0e-010.i04.local,1435;Database=promus;User Id=promus_readonly;Password=zn71!xBJ!!n2;Encrypt=true;TrustServerCertificate=true"
    );
    const result = await sql.query(query);
    console.dir(result);
    process.exit(0);
  } catch (err) {
    console.error(err);
    // ... error checks
  }
};

readFile()
  .then((values) => {
    const query = getIsbnQuery(values);
    console.log(query);
    connectToPromus(query);
    writeFile(values);
  })
  .catch((error) => {
    console.warn(error);
  });
