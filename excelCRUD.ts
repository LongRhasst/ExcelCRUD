import Excel from "exceljs";
import { InputDataWrite, InputDataRead } from "./interface";

/**
 * Class for processing Excel files: reading and writing data.
 * working flow:
 * - Step 1: Load the existing Excel file specified by Input.filePath.
 * - Step 2: Access the specified worksheet by Input.SheetName (default to "Sheet1" if not provided).
 * - Step 3: For writing, update cells based on Input.columnIds, Input.rowIds, and Input.qualityQuantity.
 * - Step 4: Save the modified workbook to Input.fileFinishPath or "output.xlsx" if not provided.
 */
class ExcelProcessor {
  public writeToExcel = async (Input: InputDataWrite): Promise<null> => {
    const workbook = new Excel.Workbook();

    await workbook.xlsx.readFile(Input.filePath);

    const worksheet = workbook.getWorksheet(Input.SheetName || "Sheet1");

    if (!worksheet) {
      console.error(`Sheet "${Input.SheetName}" not found in file: ${Input.filePath}`);
      return null;
    }

    console.log(`Reading from file: ${Input.filePath}`);
    console.log(`Sheet: ${Input.SheetName || "Sheet1"}`);

    // Get column and row identifiers
    const columnId = Input.columnIds;
    console.log("Column ID:", columnId);
    const rowId = Input.rowIds;
    console.log("Row ID:", rowId);

    // Write each quality value from qualityQuantity array
    if (Input.qualityQuantity && Input.qualityQuantity.length > 0) {
      for (let i = 0; i < Input.qualityQuantity.length; i++) {
        const item = Input.qualityQuantity[i];
        const quality = item.quality;

        const cellAddress = `${columnId}${parseInt(rowId) + i}`;

        worksheet.getCell(cellAddress).value = quality;
        console.log(`Written value ${quality} to cell ${cellAddress}`);
      }
    }

    // Save to the finish file path
    const outputPath = Input.fileFinishPath || "output.xlsx";
    await workbook.xlsx.writeFile(outputPath);

    console.log(`Excel file written to: ${outputPath}`);

    return null;
  };

  public readFromExcel = async (Input: InputDataRead): Promise<null> => {
    const workbook = new Excel.Workbook();

    // Load existing Excel file
    await workbook.xlsx.readFile(Input.filePath);

    // Get the specified worksheet
    const worksheet = workbook.getWorksheet(Input.SheetName || "Sheet1");

    if (!worksheet) {
      console.error(`Sheet "${Input.SheetName}" not found in file: ${Input.filePath}`);
      return null;
    }

    console.log(`Reading from file: ${Input.filePath}`);
    console.log(`Sheet: ${Input.SheetName || "Sheet1"}`);

    // Get column and row identifiers
    const columnId = Input.columnIds;
    const rowId = Input.rowIds;

    // Build cell address (e.g., "K6")
    const cellAddress = `${columnId}${rowId}`;
    console.log(`Cell address: ${cellAddress}`);

    // Read cell value
    const cell = worksheet.getCell(cellAddress);
    const cellValue = cell.value;

    console.log(`Cell value at ${cellAddress}:`, cellValue);
    console.log(`Cell type:`, cell.type);

    return null;
  };
}

export default ExcelProcessor;
