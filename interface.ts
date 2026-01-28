export interface InputDataWrite {
    filePath: string | "MỚI.xlsx";
    fileFinishPath: string | "MỚI1.xlsx";
    columnIds: string;
    rowIds: string;
    productType: string;
    qualityQuantity: inputProduct[];
    SheetName: string;
    overwrite?: boolean; // If true, overwrite entire file; if false/undefined, preserve existing content
}

interface inputProduct{
    type: string;
    quality: number;
}

export interface InputDataRead {
    filePath: string | "MỚI.xlsx";
    SheetName: string;
}

export interface OutputDataRead {
    name: string;
    fractions: {
        [fractionName: string] :{
            ratio: number;
            spec: {
                protein: number;
                WG: number;
                ash: number;
                loss: number;
                salt: number;
                sugar: number;
                water: number;
            };
        };
    }[];
}