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
  static generateSheetRange(columns: number, rows: number, rowStartNumber = 1) {
    const startCell = `A${rowStartNumber}`;
    const endColumnLetter = this.columnNumberToLetter(columns);
    const endRowNumber = rows;
    const endCell = `${endColumnLetter}${endRowNumber}`;
    return `${startCell}:${endCell}`;
  }
}
