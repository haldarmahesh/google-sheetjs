const { google } = require("googleapis");
import { googleAuth } from "../../../google-auth";
import { sheetService } from "../../../services";
import { ValidationType } from "../enums/validation-type.enum";
import type { Column } from "./column.entity";

export class Sheet {
  sheetId?: number;
  spreadsheetId?: string;
  constructor(public sheetName: string, public columns: Column[]) {}
  setSpreadsheetId(spreadsheetId: string) {
    this.spreadsheetId = spreadsheetId;
  }
  private getHeaderCreateConfig(spreadsheetId: string): unknown {
    return {
      spreadsheetId,
      range: `${this.sheetName}!A1:AV`,
      valueInputOption: "RAW",
      requestBody: {
        values: [this.columns.map((column) => column.name)],
      },
    };
  }
  async addHeaders(spreadsheetId: string): Promise<void> {
    await sheetService.spreadsheets.values.update(
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
    await sheetService.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests,
      },
    });
  }

  async addData(spreadsheetId: string, data: any[][]) {
    console.log(">> data", data);
    try {
      const updateToGsheet = [
        ["Name1", "CARRIER", "", "", "1", "ACTIVE", "2024-30-01"],
        ["Name2", "CHARTERING", "", "", "2", "ACTIVE", "2024-30-01"],
        ["Name3", "CARRIER", "", "", "3", "ACTIVE", "2024-30-01"],
      ];
      await sheetService.spreadsheets.values.update({
        spreadsheetId: spreadsheetId,
        range: `Sheet1!A2:H1000`,
        valueInputOption: "RAW",
        resource: {
          values: updateToGsheet,
        },
      } as any);
    } catch (err) {
      console.log(err);
      throw err;
    }
  }
}
