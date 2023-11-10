import { getIsbnQuery, queryPromus } from "./sqlUtils";
import { readFile, writeFile } from "./excelUtils";

console.log("Starting script");

readFile()
  .then((values) => {
    const query = getIsbnQuery(values);
    console.log(`Querying database for ${values.length} ISBNs:`, values);
    if (values.length === 0)
      throw new Error(
        "No ISBNs found in isbn.xlsx. Are the ISBNs contained in column A of the first worksheet?"
      );
    return queryPromus(query);
  })
  .then((result) => {
    if (result === undefined)
      throw new Error("Result from Promus is undefined");
    return writeFile(result);
  })
  .then(() => {
    console.log(
      'Results are now available in worksheet "Result" in "isbn.xlsx"'
    );
    process.exit(0);
  })
  .catch((error) => {
    console.warn(error);
  });
