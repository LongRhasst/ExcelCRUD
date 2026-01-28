import {
    InputDataWrite,
    InputDataRead
} from "./interface";
import ExcelProcessor from "./excelCRUD";

const mainFunction = async (): Promise<null> => {
    const excelProcess: ExcelProcessor = new ExcelProcessor();

    const dataWrite: InputDataWrite = {
        filePath: "MỚI.xlsx",
        fileFinishPath: "MỚI2.xlsx",
        columnIds: "B",
        rowIds: "10",
        productType: "Flour",
        SheetName: "Sheet1",
        qualityQuantity: [
            { type: "Type1", quality: 200 },
        ]
    };

    const dataRead: InputDataRead = {
        filePath: "MỚI.xlsx",
        SheetName: "Sheet1",
    };

    // await excelProcess.writeToExcel(dataWrite);
    await excelProcess.readFromExcel(dataRead);
    return null;
}

function main() {
    mainFunction().catch((error) => {
        // eslint-disable-next-line no-console
        console.error("Error in mainFunction:", error);
    });
}

main();