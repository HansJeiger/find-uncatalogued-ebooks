import { getIsbnQuery, queryPromus } from "./sqlUtils";
import { readFile, writeFile } from "./excelUtils";

readFile()
  .then((values) => {
    const query = getIsbnQuery(values);
    return queryPromus(query);
  })
  .then((result) => {
    if (result === undefined)
      throw new Error("Result from Promus is undefined");
    return writeFile(result);
  })
  .then(() => process.exit(0))
  .catch((error) => {
    console.warn(error);
  });
