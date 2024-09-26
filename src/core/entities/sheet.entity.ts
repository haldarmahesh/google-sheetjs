import { sheetService } from "../../infrastructure/googlesheet-service";
import { ValidationType } from "../enums/validation-type.enum";
import { Column } from "./column.entity";
import type { Spreadsheet } from "./spreadsheet.entity";
import { SheetUtils } from "../utils/sheet.utils";
export class Sheet {
  sheetId?: number;
  spreadsheetId?: string;
  public columns: Column[] = [];
  public rows: any[][] = [];
  public headerNames: string[] = [];
  constructor(public name: string, private spreadsheet: Spreadsheet) {}
  public getSpreadsheetId(): string {
    return this.spreadsheet.spreadsheetId || "N/A";
  }
  public addColumn(headerName: string): Column {
    this.headerNames.push(headerName);
    const header = new Column(headerName);
    this.columns.push(header);
    return header;
  }

  public addColumns(columnList: Column[]): this {
    this.headerNames = columnList.map((column) => column.name);
    this.columns = columnList;
    return this;
  }

  private getHeaderCreateConfig(spreadsheetId: string): unknown {
    return {
      spreadsheetId,
      range: `${this.name}!${SheetUtils.generateSheetRange(
        this.columns.length,
        1
      )}`,
      valueInputOption: "RAW",
      requestBody: {
        values: [this.columns.map((column) => column.name)],
      },
    };
  }
  async addHeaders(spreadsheetId: string): Promise<void> {
    await sheetService(this.spreadsheet.googleAuth).spreadsheets.values.update(
      this.getHeaderCreateConfig(spreadsheetId) as any
    );
    console.log(">> Headers added");
  }
  private getConfigToProtectColumns(column: Column, index: number): unknown {
    if (column.isProtected) {
      return {
        addProtectedRange: {
          protectedRange: {
            range: {
              sheetId: this.sheetId,
              startRowIndex: 0,
              startColumnIndex: index,
              endColumnIndex: index + 1,
            },
            description: "Protecting metadata row",
            warningOnly: false,
            requestingUserCanEdit: true,
            editors: {
              users: ["mahesh.haldar@sennder.com"],
              domainUsersCanEdit: false,
            },
          },
        },
      };
    }
  }
  private getProtectedColumnFormatConfig(
    column: Column,
    index: number
  ): unknown {
    if (column.isProtected) {
      return {
        repeatCell: {
          range: {
            sheetId: this.sheetId,
            startRowIndex: 1,
            startColumnIndex: index,
            endColumnIndex: index + 1,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: {
                red: 50.0,
                green: 50.0,
                blue: 50.0,
              },
              textFormat: {
                bold: false,
              },
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat)",
        },
      };
    }
  }
  private getHeaderFormatConfig(): unknown[] {
    return [
      {
        repeatCell: {
          range: {
            sheetId: this.sheetId,
            startRowIndex: 0,
            endRowIndex: 1,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: {
                red: 150.0,
                green: 52.0,
                blue: 18.0,
              },
              textFormat: {
                bold: true,
              },
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat)",
        },
      },
    ];
  }
  private getConfigToSetDataType(column: Column, index: number): unknown {
    if (column.validation) {
      if (column.validation.type === ValidationType.ONE_OF_LIST) {
        return {
          repeatCell: {
            range: {
              sheetId: this.sheetId,
              startRowIndex: 1,
              startColumnIndex: index,
              endColumnIndex: index + 1,
            },
            cell: {
              dataValidation: {
                condition: {
                  type: "ONE_OF_LIST",
                  values: column.validation.values.map((value) => ({
                    userEnteredValue: value,
                  })),
                },
                strict: true,
                showCustomUi: true,
              },
            },
            fields: "dataValidation",
          },
        };
      }
    }
  }
  private getColumnConfig(): any[] {
    const requests: unknown[] = [];
    this.columns.forEach((column, index) => {
      requests.push(this.getProtectedColumnFormatConfig(column, index));
      requests.push(this.getConfigToProtectColumns(column, index));
      requests.push(this.getConfigToSetDataType(column, index));
    });

    return requests.filter((request) => request !== undefined);
  }
  async addFormattingAndValidations(spreadsheetId: string): Promise<void> {
    if (!this.sheetId) {
      throw new Error("Sheet ID is not available");
    }
    const requests = [
      ...this.getHeaderFormatConfig(),
      ...this.getColumnConfig(),
    ];
    await sheetService(this.spreadsheet.googleAuth).spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests,
      },
    });
  }

  async addData(data: any[][]) {
    if (!this.sheetId) {
      throw new Error("Sheet ID is not available");
    }
    try {
      await sheetService(
        this.spreadsheet.googleAuth
      ).spreadsheets.values.update({
        spreadsheetId: this.spreadsheet.spreadsheetId,
        range: `${this.name}!${SheetUtils.generateSheetRange(
          this.columns.length,
          data.length,
          2
        )}`,
        valueInputOption: "USER_ENTERED",
        resource: {
          values: data,
        },
      } as any);
    } catch (err) {
      console.log(err);
      throw err;
    }
  }
  getHeaderNames(): string[] {
    return this.columns.map((column) => column.name);
  }
  async addRow(rowData: Record<string, string>) {
    const currentColumnNames = this.getHeaderNames();
    const allColumnNameMatches = Object.keys(rowData).every((rowKey) => {
      return currentColumnNames.includes(rowKey);
    });

    if (!allColumnNameMatches) {
      console.log(
        "the columns names you passed are =>",
        Object.keys(rowData),
        "do not match the columns names in the sheet =>",
        currentColumnNames
      );
      throw new Error("Add correct column names");
    }
    sheetService(this.spreadsheet.googleAuth).spreadsheets.values.append({});
  }
  checkIfSpreadsheetIdExists(): void {
    if (!this.spreadsheet.spreadsheetId) {
      throw new Error("Spreadsheet ID is not available");
    }
  }
  async loadHeader(): Promise<void> {
    this.checkIfSpreadsheetIdExists();
    try {
      const infoObjectFromSheet = await sheetService(
        this.spreadsheet.googleAuth
      ).spreadsheets.values.get({
        spreadsheetId: this.spreadsheet.spreadsheetId,
        range: `${this.name}!A1:Z1`,
      });
      const valuesFromSheet = infoObjectFromSheet.data.values;
      if (valuesFromSheet) {
        this.headerNames = valuesFromSheet[0];
        this.headerNames.forEach((headerName) => {
          this.columns.push(new Column(headerName));
        });
      }
    } catch (err) {
      console.log(err);
      throw err;
    }
  }
  async loadData(): Promise<any[][]> {
    this.checkIfSpreadsheetIdExists();
    try {
      const infoObjectFromSheet = await sheetService(
        this.spreadsheet.googleAuth
      ).spreadsheets.values.get({
        spreadsheetId: this.spreadsheet.spreadsheetId,
        range: `${this.name}!A2:Z1000`,
      });

      const valuesFromSheet = infoObjectFromSheet.data.values;
      this.rows = valuesFromSheet || [];
      return valuesFromSheet as any[][];
    } catch (err) {
      console.log(err);
      throw err;
    }
  }

  getRow(rowNumber: number): any[] {
    return this.rows[rowNumber];
  }

  getCell(rowNumber: number, colNumber: number): any[] {
    return this.rows[rowNumber][colNumber];
  }

  getFormattedRow(rowNumber: number): any {
    const row = this.rows[rowNumber];
    const formattedRow: any = {};
    this.headerNames.forEach((header, index) => {
      formattedRow[header] = row[index];
    });
    return formattedRow;
  }
}
