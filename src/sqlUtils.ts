import sql, { config as SqlConfig } from "mssql";

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

export const getIsbnQuery = (isbns: string[]) => {
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

export const queryPromus = async (query: string) => {
  console.log("Sending query to Promus");
  try {
    // make sure that any items are correctly URL encoded in the connection string
    await sql.connect(
      sqlConfig //"Server=wg-sxd0e-010.i04.local,1435;Database=promus;User Id=promus_readonly;Password=zn71!xBJ!!n2;Encrypt=true;TrustServerCertificate=true"
    );
    const result = await sql.query(query);
    console.log("Retrieved result from Promus");

    return result;
  } catch (err) {
    console.error(err);
    throw new Error(
      "Cannot connect to Promus database, make sure you are connected to the WG-XWLAN network either directly or through VPN and that variables in .env are correct"
    );
  }
};
