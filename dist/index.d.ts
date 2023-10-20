declare enum Direction {
    Row = 0,
    Column = 1
}
type Json2ExcelProp = {
    data: [
        {
            sheetName: string;
            sheetData: unknown[];
        }
    ];
    fileName: string;
    direction: Direction;
};
declare const json2excel: (props: Json2ExcelProp) => void;
export default json2excel;
