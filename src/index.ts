import { utils, writeFileXLSX } from 'xlsx';
import {verticalExport} from "./exportVerticalExcel";
import {horizontalExport} from "./exportHorizontalExcel";

enum Direction {
    Row,
    Column
}
export type Json2ExcelProp = {
    data: [{
        sheetName: string,
        sheetData: unknown[];
    }]
    fileName: string,
    direction: Direction
}
const json2excel = (props: Json2ExcelProp): void => {
    if (props.direction === 0) {
        verticalExport(props)
    } else {
        horizontalExport(props)
    }
};
export default json2excel;
