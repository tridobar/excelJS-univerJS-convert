// 使用：
// this.excelJsReadFile(file).then((snapshot) => {
//   this.univerAPI.createWorkbook(snapshot);
// });
async function excelJsReadFile(file) {
  function argbToRgb(argb) {
    if (!argb) return null;
    if (argb.length === 8) return { rgb: '#' + argb.substring(2, 8) };
    else if (argb.length === 6) return { rgb: '#' + argb };
    else return null;
  }
  function mapBorderStyle(style: ExcelJS.BorderStyle) {
    switch (style) {
      case 'dashDot':
        return 5;
      case 'dashDotDot':
        return 6;
      case 'dashed':
        return 4;
      case 'dotted':
        return 3;
      case 'double':
        return 7;
      case 'hair':
        return 2;
      case 'medium':
        return 8;
      case 'mediumDashDot':
        return 10;
      case 'mediumDashDotDot':
        return 11;
      case 'mediumDashed':
        return 9;
      case 'slantDashDot':
        return 12;
      case 'thick':
        return 13;
      case 'thin':
        return 1;
      default:
        return 0;
    }
  }
  const workbook = new ExcelJS.Workbook();
  const workbookId = file.name + new Date().getTime().toString();
  await workbook.xlsx.load(file);
  const sheets = {};
  const sheetOrder = [];
  workbook.eachSheet((worksheet, sheetId) => {
    const cellData = {};
    const rowData = [];
    worksheet.eachRow((row, rowNumber) => {
      const rowIndex = rowNumber - 1;
      // 遍历处理单元格数据
      row.eachCell((cell, colNumber) => {
        const colIndex = colNumber - 1;
        if (!cellData[rowIndex]) {
          cellData[rowIndex] = [];
        }
        cellData[rowIndex][colIndex] = {};
        const workCellFormula = cell.formula;
        if (workCellFormula) {
          // 公式
          cellData[rowIndex][colIndex].f = workCellFormula;
        } else {
          // 文本
          cellData[rowIndex][colIndex] = { v: cell.value };
        }
        // 单元格样式
        const cs = cell.style;
        if (cs) {
          cellData[rowIndex][colIndex].s = {
            ff: cs.font?.name, 	//字体
            fs: cs.font?.size, 	//字体大小
            it: cs.font?.italic, 	//是否斜体
            bl: cs.font?.bold, 	//是否加粗
            ul: { s: cs.font?.underline ? 1 : 0 }, 	//下划线
            st: { s: cs.font?.strike ? 1 : 0 }, 	//删除线
            ol: null,	//上划线
            // @ts-ignore
            bg:	argbToRgb(cs.fill?.fgColor?.argb), 	//背景颜色
            bd: {
              t: { s: mapBorderStyle(cs.border?.top?.style), cl: argbToRgb(cs.border?.top?.color?.argb) },
              b: { s: mapBorderStyle(cs.border?.bottom?.style), cl: argbToRgb(cs.border?.bottom?.color?.argb) },
              l: { s: mapBorderStyle(cs.border?.left?.style), cl: argbToRgb(cs.border?.left?.color?.argb) },
              r: { s: mapBorderStyle(cs.border?.right?.style), cl: argbToRgb(cs.border?.right?.color?.argb) }
            }, //边框
            cl:	argbToRgb(cs.font?.color?.argb), //字体颜色
            va:	null, //上标下标
            tr:	cs.alignment?.textRotation === 'vertical' ? { v: 1, a: 0 } : { v: 0, a: cs.alignment?.textRotation || 0 }, //文字旋转
            ht:	cs.alignment?.horizontal === 'center' ? 2 : cs.alignment?.horizontal === 'right' ? 3 : 1, //水平对齐方式
            vt:	cs.alignment?.vertical === 'top' ? 1 : cs.alignment?.vertical === 'bottom' ? 3 : 2, //垂直对齐方式
            tb:	cs.alignment?.wrapText ? 3 : 1, //截断溢出
            pd: null,	//内边距
            n: null,	//数字格式
          };
        }
      });
      rowData.push({ h: row.height, hd: row.hidden ? 1 : 0 });
    });
    sheets[sheetId] = {
      id: sheetId.toString(),
      name: worksheet.name,
      rowCount: worksheet.rowCount + 50,
      columnCount: worksheet.columnCount + 50,
      zoomRatio: 1,
      cellData: cellData,
      showGridlines: 1,
      mergeData: [],
      columnData: worksheet.columns.map(column => ({ w: column.width ? Math.round(column.width * 8) : null, hd: column.hidden ? 1 : 0 })),
      rowData,
    };
    // 合并单元格
    if (worksheet.hasMerges) {
      worksheet.model.merges.forEach(merge => {
        const mergeInfo = merge.split(':');
        const startCell = worksheet.getCell(mergeInfo[0]);
        const endCell = worksheet.getCell(mergeInfo[1]);
        sheets[sheetId].mergeData.push({
          startRow: Number(startCell.row) - 1,
          startColumn: Number(startCell.col) - 1,
          endRow: Number(endCell.row) - 1,
          endColumn: Number(endCell.col) - 1,
          rangeType: 0,
          unitId: sheetId,
          sheetId: workbookId,
        });
      });
    }
    sheetOrder.push(sheetId.toString());
  });
  return {
    id: workbookId,
    name: file.name,
    appVersion: '',
    locale: UniverCore.LocaleType.ZH_CN,
    sheetOrder,
    sheets,
  };
}
