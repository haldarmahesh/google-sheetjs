const { google } = require("googleapis");
import { sheetService } from "../../infrastructure/googlesheet-service";
import { ValidationType } from "../enums/validation-type.enum";
import { Column } from "./column.entity";
import type { Spreadsheet } from "./spreadsheet.entity";

export class Sheet {
  sheetId?: number;
  spreadsheetId?: string;
  public columns: Column[] = [];
  constructor(public name: string, private spreadsheet: Spreadsheet) {}
  public getSpreadsheetId(): string {
    return this.spreadsheet.spreadsheetId || "N/A";
  }
  public addColumn(name: string): Column {
    const column = new Column(name);
    this.columns.push(column);
    return column;
  }
  private getHeaderCreateConfig(spreadsheetId: string): unknown {
    return {
      spreadsheetId,
      range: `${this.name}!A1:AV`,
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
    console.log('>> Da', data)
    if(!this.sheetId) {
        throw new Error("Sheet ID is not available");
    }
    try {
      await sheetService(this.spreadsheet.googleAuth).spreadsheets.values.update({
        spreadsheetId: this.spreadsheet.spreadsheetId,
        range: `${this.name}!A2`,
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
  checkIfSpreadsheetIdExists(): void {
    if(!this.spreadsheet.spreadsheetId) {
        throw new Error("Spreadsheet ID is not available");
    }
  }
  async readData(): Promise<any[][]> {
    this.checkIfSpreadsheetIdExists()
    try {   

        const infoObjectFromSheet = await sheetService(this.spreadsheet.googleAuth).spreadsheets.values.get({
            spreadsheetId: this.spreadsheet.spreadsheetId,
            range: `${this.name}!A2:Z1000`,
        })

        const valuesFromSheet = infoObjectFromSheet.data.values;
        return valuesFromSheet as any[][];

    } catch(err) {
        console.log(err)
        throw err;
    }
  }
}
