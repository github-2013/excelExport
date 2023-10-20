import { utils, writeFileXLSX } from 'xlsx';

enum Direction {
    Row,
    Column
}
type Json2ExcelProp = {
    data: [{
        sheetName: string,
        sheetData: unknown[];
    }]
    fileName: string,
    direction: Direction
}
const json2excel = (props: Json2ExcelProp): void => {
    // 工作薄
    const workbook = utils.book_new();
    props.data.forEach((item) => {
        // 工作表
        const worksheet = utils.json_to_sheet([]);
        // 头部标题
        item.sheetData.forEach((subItem, subIndex) => {
            utils.sheet_add_aoa(worksheet, Object.entries(subItem), { origin: `A${subIndex + 1}` });
        });
        // 增加工作表
        if (item.sheetName.length >= 31) {
            utils.book_append_sheet(workbook, worksheet, `${item.sheetName.substr(0, 28)}...`);
        } else {
            utils.book_append_sheet(workbook, worksheet, item.sheetName);
        }

        // 计算列宽度
        const cellMaxWidth = item.sheetData.reduce((maxWidth, cell) => {
            const cellEntries = Object.entries(cell)
            return {
                colWidth1: Math.max(cellEntries[0][0].length, maxWidth.colWidth1),
                colWidth2: Math.max(cellEntries[0][1].length, maxWidth.colWidth2)
            };
        }, { colWidth1: 10, colWidth2: 10 });
        worksheet['!cols'] = [{ wch: cellMaxWidth.colWidth1 }, { wch: cellMaxWidth.colWidth2 }];
    });

    // 写入到文件
    writeFileXLSX(workbook, props.fileName);
};
export default json2excel;
