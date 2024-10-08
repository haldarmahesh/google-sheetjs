export class SheetUtils {
  private static columnNumberToLetter(columnNumber: number) {
    let letter = "";
    while (columnNumber > 0) {
      let modulo = (columnNumber - 1) % 26;
      letter = String.fromCharCode(65 + modulo) + letter;
      columnNumber = Math.floor((columnNumber - modulo) / 26);
    }
    return letter;
  }
  private static columnIndexToLetter(columnIndex: number) {
    let letter = "";
    while (columnIndex >= 0) {
      let modulo = columnIndex % 26;
      letter = String.fromCharCode(65 + modulo) + letter;
      columnIndex = Math.floor(columnIndex / 26) - 1;
    }
    return letter;
  }
  static generateSheetRange(columns: number, rows: number, rowStartNumber = 1) {
    const startCell = `A${rowStartNumber}`;
    const endColumnLetter = this.columnNumberToLetter(columns);
    const endRowNumber = rows;
    const endCell = `${endColumnLetter}${endRowNumber}`;
    return `${startCell}:${endCell}`;
  }

  static getSheetRange(
    startColumnIndex: number,
    endColumnIndex: number,
    startRowIndex: number = 0,
    endRowIndex: number = 1
  ) {
    if (startColumnIndex > endColumnIndex) {
      throw new Error(
        "start column index should be less than end column index"
      );
    }
    if (startRowIndex > endRowIndex) {
      throw new Error("end row index should be greater than start row index");
    }
    if (startColumnIndex === endColumnIndex) {
      throw new Error("end row index should be greater than start row index");
    }
    if (startColumnIndex < 0 || endColumnIndex < 0 || startRowIndex < 0) {
      throw new Error("Invalid column or row index");
    }
    return `${this.columnIndexToLetter(startColumnIndex)}${
      startRowIndex + 1
    }:${this.columnIndexToLetter(endColumnIndex - 1)}${endRowIndex}`;
  }
}
