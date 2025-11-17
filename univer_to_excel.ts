// const fWorkbook = this.univerAPI.getActiveWorkbook();
// const univerJson = fWorkbook.getSnapshot();  // fWorkbook.save() 获取的json没有style
// const workbook = new ExcelJS.Workbook();
function reverseJsonToWorkbook(workbook: ExcelJS.Workbook, univerJson) {
  function rgbToArgb(rgb) {
    if (!rgb) return null;
    return rgb.replace('#', 'FF');
  }
  function mapBorderStyle(style: number): ExcelJS.BorderStyle {
    switch (style) {
      case 1:
        return 'thin';
      case 2:
        return 'hair';
      case 3:
        return 'medium';
      case 4:
        return 'dashed';
      case 5:
        return 'dotted';
      case 6:
        return 'dashDot';
      case 7:
        return 'double';
      case 8:
        return 'mediumDashed';
      case 9:
        return 'mediumDashDot';
      case 10:
        return 'mediumDashDotDot';
      case 11:
        return 'slantDashDot';
      case 12:
        return 'thick';
      case 13:
        return 'double';
      default:
        return null;
    }
  }
  // 遍历单元格数据
  univerJson.sheetOrder.forEach(sheetId => {
    const sheetInfo = univerJson.sheets[sheetId];
    const sheetName = sheetInfo.name;
    const cellData = sheetInfo.cellData;
    const newSheet = workbook.addWorksheet(sheetName);
    for (let i = 0; i < Object.keys(cellData).reduce((acc, cur) => Math.max(acc, parseInt(cur)), 0); i++) {
      const rowKey = i.toString();
      const rowCells = cellData[rowKey];
      const newRow = newSheet.addRow([]);
      if (rowCells) {
        for (const colKey in rowCells) {
          const colIndex = parseInt(colKey);
          const cell = rowCells[colKey];
          if (!cell) continue;

          const newCell = newRow.getCell(colIndex + 1);
          // 赋值
          newCell.value = cell.v;
          // 公式
          if (cell.f) {
            newCell.value = {
              formula: cell.f,
              result: cell.v,
            };
          }
          // 样式
          if (cell.s) {
            let style = cell.s;
            if (typeof style === 'string') {
              style = univerJson.styles[style];
            }
            newCell.style = {
              font: {
                name: style.ff,
                size: style.fs,
                italic: style.it,
                bold: style.bl,
                underline: style.ul?.s === 1,
                strike: style.st?.s === 1,
                color: { argb: rgbToArgb(style.cl?.rgb) },
              },
              fill: {
                type: 'pattern',
                pattern: style.bg ? 'solid' : 'none',
                fgColor: { argb: rgbToArgb(style.bg?.rgb) },
                bgColor: { argb: rgbToArgb(style.bg?.rgb) },
              },
              border: {
                top: { style: mapBorderStyle(style.bd?.t?.s), color: { argb: rgbToArgb(style.bd?.t?.cl?.rgb) } },
                bottom: { style: mapBorderStyle(style.bd?.b?.s), color: { argb: rgbToArgb(style.bd?.b?.cl?.rgb) } },
                left: { style: mapBorderStyle(style.bd?.l?.s), color: { argb: rgbToArgb(style.bd?.l?.cl?.rgb) } },
                right: { style: mapBorderStyle(style.bd?.r?.s), color: { argb: rgbToArgb(style.bd?.r?.cl?.rgb) } },
              },
              alignment: {
                horizontal: style.ht === 2 ? 'center' : style.ht === 3 ? 'right' : 'left',
                vertical: style.vt === 1 ? 'top' : style.vt === 3 ? 'bottom' : 'middle',
                wrapText: style.tb === 3,
                indent: 0,
                shrinkToFit: false,
                textRotation: style.tr?.v === 1 ? 'vertical' : (style.tr?.a || 0),
                readingOrder: 'ltr',
              },
            };
          }
        }
      }
      // 行高
      if (sheetInfo.rowData[rowKey]) {
        newRow.height = (sheetInfo.rowData[rowKey].h || sheetInfo.defaultRowHeight) / 1.33;
        newRow.hidden = sheetInfo.rowData[rowKey].hd === 1;
      }
    }
    // 列宽
    for (let i = 0; i < Object.keys(sheetInfo.columnData).reduce((acc, cur) => Math.max(acc, parseInt(cur)), 0); i++) {
      const colKey = i.toString();
      if (sheetInfo.columnData[colKey]) {
        newSheet.columns[i].width = (sheetInfo.columnData[colKey].w ? sheetInfo.columnData[colKey].w : sheetInfo.defaultColumnWidth) / 8.43;
        newSheet.columns[i].hidden = sheetInfo.columnData[colKey].hd === 1;
      }
    }

    // 工作表隐藏
    newSheet.state = sheetInfo.hidden === 1 ? 'hidden' : 'visible';
    // 合并单元格
    sheetInfo.mergeData?.forEach(merge => {
      newSheet.mergeCells(merge.startRow + 1, merge.startColumn + 1, merge.endRow + 1, merge.endColumn + 1);
    });
  });
}
