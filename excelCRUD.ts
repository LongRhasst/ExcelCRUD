import Excel from "exceljs";
import { InputDataWrite, InputDataRead, OutputDataRead } from "./interface";

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
      console.error(
        `Sheet "${Input.SheetName}" not found in file: ${Input.filePath}`,
      );
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

  public readFromExcel = async (Input: InputDataRead): Promise<null | OutputDataRead[]> => {
    const workbook = new Excel.Workbook();

    // Load existing Excel file
    await workbook.xlsx.readFile(Input.filePath);

    let output : OutputDataRead[] = [];

    // Get the specified worksheet
    const worksheet = workbook.getWorksheet(Input.SheetName || "Sheet1");

    if (!worksheet) {
      return null;
    }

    if(worksheet.getCell(`A3`).value !== "NL"){
        return null;
    }
    
    for(let paddyIndex: number = 6;;paddyIndex+=3){
        const cellAddress = `A${paddyIndex}`;
        const paddyName = worksheet.getCell(cellAddress).value;
        
        if(paddyName === null){
            break;
        }
        
        const paddyData: OutputDataRead = {
            name: paddyName?.toString() || "",
            fraction: []
        };
        
        for (let fractionIndex: number = 0; fractionIndex <= 3; fractionIndex++){
            const baseRow = paddyIndex + fractionIndex;
            const ratioCell = worksheet.getCell(`F${baseRow}`).value;
            const ratio = typeof ratioCell === 'number' ? ratioCell : parseFloat(ratioCell?.toString() || "0");
            
            const fractionData: any = {
                ratio: ratio,
                spec: {}
            };
            
            for (let col = 'G'; col <= 'M'; col = String.fromCharCode(col.charCodeAt(0) + 1)) {
                const fractionCellAddressTitle = `${col}4`;
                const fractionCellAddress = `${col}${baseRow}`;
                const titleValue = worksheet.getCell(fractionCellAddressTitle).value?.toString() || "";
                const cellValue = worksheet.getCell(fractionCellAddress).value;
                const fractionValue = typeof cellValue === 'number' ? cellValue : parseFloat(cellValue?.toString() || "0");
                
                if(titleValue){
                    fractionData.spec[titleValue] = fractionValue;
                }
            }
            
            paddyData.fraction.push(fractionData);
        }
        
        output.push(paddyData);
    }

    console.log(JSON.stringify(output, null, 2));
    
    return output;
  };
}

export default ExcelProcessor;
