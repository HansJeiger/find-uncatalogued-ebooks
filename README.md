# find-uncatalogued-ebooks
## Setup
1. Install [node.js](https://nodejs.org/en/download) if you have not already
2. Clone this project
3. In the project directory, run the file named `setup.cmd` which installs node packages and creates a `.env` file
4. In the `.env` file, fill in the required values. You can get them from someone with access to Doppler

## How to use
6. Make sure excel document is in correct format (see below)
7. Put the excel document containing the ISBNs to the root folder of the project
8. Make sure excel document is not open anywhere on your computer
9. Run `run-script.cmd`
10. Results can now be accessed in your excel file in a new worksheet named `Result`

## Correct format of excel document
1. Name of excel document must be `isbn.xlsx`
2. ISBNs are retrieved from column A in the first worksheet
3. `Result` is a reserved name for the resulting worksheet, and should not be used by any existing worksheets
