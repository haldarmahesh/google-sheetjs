import { SheetUtils } from "../sheet.utils";

describe('SheetUtils', () => {
  describe('columnNumberToLetter', () => {
    it('should convert column number to letter', () => {
      const result = (SheetUtils as any).columnNumberToLetter(1);
      expect(result).toBe('A');
    });

    it('should convert column number to letter for double letters', () => {
      const result = (SheetUtils as any).columnNumberToLetter(27);
      expect(result).toBe('AA');
    });

    it('should convert column number to letter for triple letters', () => {
      const result = (SheetUtils as any).columnNumberToLetter(703);
      expect(result).toBe('AAA');
    });
  });

  describe('generateSheetRange', () => {
    it('should generate sheet range for given columns and rows', () => {
      const result = SheetUtils.generateSheetRange(3, 5);
      expect(result).toBe('A1:C5');
    });

    it('should generate sheet range with custom row start number', () => {
      const result = SheetUtils.generateSheetRange(3, 5, 2);
      expect(result).toBe('A2:C5');
    });
  });

  describe('getSheetRange', () => {
    it('should be defined', () => {
      expect(SheetUtils.getSheetRange).toBeDefined();
    });
    it('should throw error when basic validation is wrong', () => {
        expect(() => SheetUtils.getSheetRange(-1, 2)).toThrow();
        expect(() => SheetUtils.getSheetRange(0, 0)).toThrow();
        expect(() => SheetUtils.getSheetRange(1, 0)).toThrow();
        expect(() => SheetUtils.getSheetRange(1, 0, 6)).toThrow();
    })
    it('should generate sheet range for given columns and rows', () => {
      expect(SheetUtils.getSheetRange(0, 1)).toBe("A1:A1");
      expect(SheetUtils.getSheetRange(0, 2)).toBe("A1:B1");
      expect(SheetUtils.getSheetRange(0, 3, 5, 6)).toBe("A6:C6");
      expect(SheetUtils.getSheetRange(0, 3, 1, 3)).toBe("A2:C3");
      expect(SheetUtils.getSheetRange(2, 3, 2, 3)).toBe("C3:C3");
      expect(SheetUtils.getSheetRange(4, 5)).toBe("E1:E1");
    });

  });
});