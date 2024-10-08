import { sheetService } from "../../infrastructure/googlesheet-service";
import { ValidationType } from "../enums/validation-type.enum";
import { Column } from "./column.entity";
import type { Spreadsheet } from "./spreadsheet.entity";
import { SheetUtils } from "../utils/sheet.utils";
import { ColumnGroup } from "./column-group.entity";
interface GroupHeaderConfig {
  enabled: boolean;
  groups: { name: string; startColumnIndex: number; endColumnIndex: number }[];
}
export class SheetProperties {
  constructor(public rowCount: number, public columnCount: number) {}
}
export class Sheet {
  sheetId?: number;
  spreadsheetId?: string;
  public columns: Column[] = [];
  public rows: any[][] = [];
  public headerNames: string[] = [];
  private groupHeaderConfig: GroupHeaderConfig = {
    enabled: false,
    groups: [],
  };
  constructor(
    public name: string,
    private spreadsheet: Spreadsheet,
    public properties?: SheetProperties
  ) {}
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
    this.columns = [...this.columns, ...columnList];
    return this;
  }
  public addColumnGroup(groupName: string, columnList: Column[]): this {
    this.groupHeaderConfig.enabled = true;
    this.headerNames = columnList.map((column) => column.name);
    this.columns = [...this.columns, ...columnList];
    this.groupHeaderConfig.groups.push({
      name: groupName,
      startColumnIndex: this.columns.length - columnList.length,
      endColumnIndex: this.columns.length - 1,
    });
    return this;
  }
  public addColumnGroups(list: ColumnGroup[]): this {
    list.forEach((columnGroup) => {
      this.addColumnGroup(columnGroup.name, columnGroup.columns);
    });
    return this;
  }
  // public addColumnGroups(groupName: string, columnList: Column[]): this {
  //   this.headerNames = columnList.map((column) => column.name);
  //   this.columns = columnList;
  //   return this;
  // }

  public getFormattedData(): Record<string, string>[] {
    const result: Record<string, string>[] = [];
    this.rows.forEach((_, index) => {
      result.push(this.getFormattedRow(index));
    });
    return result;
  }

  private getHeaderCreateConfigWithGroup(spreadsheetId: string): unknown {
    const headerGroupConfig: any[] = [];
    const GROUP_ROW_NUMBER = 0;
    const SUB_GROUP_ROW_NUMBER = this.groupHeaderConfig.enabled
      ? 1
      : GROUP_ROW_NUMBER;
    if (this.groupHeaderConfig.enabled) {
      this.groupHeaderConfig.groups.forEach((group) => {
        headerGroupConfig.push({
          range: `${this.name}!${SheetUtils.getSheetRange(
            group.startColumnIndex,
            group.endColumnIndex
          )}`,
          values: [[group.name]],
        });
      });
    }
    headerGroupConfig.push({
      range: `${this.name}!${SheetUtils.getSheetRange(
        0,
        this.columns.length,
        SUB_GROUP_ROW_NUMBER,
        SUB_GROUP_ROW_NUMBER + 1
      )}`,
      values: [this.columns.map((column) => column.name)],
    });
    console.log('row config', JSON.stringify(headerGroupConfig))
    return {
      spreadsheetId,
      valueInputOption: "RAW",
      requestBody: {
        data: [
          ...headerGroupConfig,
        ],
      },
    };
  }
  async addHeaders(spreadsheetId: string): Promise<void> {
    await sheetService(
      this.spreadsheet.googleAuth
    ).spreadsheets.values.batchUpdate(
      this.getHeaderCreateConfigWithGroup(spreadsheetId) as any
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

  private getGroupHeaderFormatConfig(): unknown[] {
    // TODO: (mahesh) colorize when group and columns both are used
    const COLORS = [
      {
        red: 150.0,
        green: 52.0,
        blue: 18.0,
      },
      {
        red: 50.0,
        green: 52.0,
        blue: 118.0,
      },
    ];
    let format: any[] = [];
    if (this.groupHeaderConfig.enabled) {
      this.groupHeaderConfig.groups.forEach((group, index) => {
        format = format.concat([
          {
            repeatCell: {
              range: {
                sheetId: this.sheetId,
                startColumnIndex: group.startColumnIndex,
                endColumnIndex: group.endColumnIndex + 1,
                startRowIndex: 0,
                endRowIndex: 2,
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: COLORS[index % 2],
                  textFormat: {
                    bold: true,
                  },
                },
              },
              fields: "userEnteredFormat(backgroundColor,textFormat)",
            },
          },
          {
            repeatCell: {
              range: {
                sheetId: this.sheetId,
                startColumnIndex: group.startColumnIndex,
                endColumnIndex: group.endColumnIndex + 1,
                startRowIndex: 0,
                endRowIndex: 1,
              },
              cell: {
                userEnteredFormat: {
                  horizontalAlignment: "CENTER",
                },
              },
              fields: "userEnteredFormat(horizontalAlignment)",
            },
          },
          {
            mergeCells: {
              range: {
                sheetId: this.sheetId,
                startRowIndex: 0,
                endRowIndex: 1,
                startColumnIndex: group.startColumnIndex,
                endColumnIndex: group.endColumnIndex + 1,
              },
              mergeType: "MERGE_ROWS",
            },
          },
        ]);
      });
    }

    if (!this.groupHeaderConfig.enabled) {
      format.push({
        repeatCell: {
          range: {
            sheetId: this.sheetId,
            startRowIndex: 1,
            endRowIndex: 2,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: COLORS[0],
              textFormat: {
                bold: true,
              },
            },
          },
          fields: "userEnteredFormat(backgroundColor,textFormat)",
        },
      });
    }
    return format;
  }
  private getConfigToSetDataType(column: Column, index: number): unknown {
    if (column.validation) {
      if (column.validation.type === ValidationType.ONE_OF_LIST) {
        return {
          repeatCell: {
            range: {
              sheetId: this.sheetId,
              startRowIndex: this.groupHeaderConfig.enabled ? 2 : 1,
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
  private getHeaderNotesConfig(): unknown[] {
    const format: unknown[] = [];
    this.columns.forEach((column, index) => {
      if (column.note) {
        format.push({
          repeatCell: {
            fields: "note",
            range: {
              sheetId: this.sheetId,
              startRowIndex: this.groupHeaderConfig.enabled ? 1 : 0,
              endRowIndex: 2,
              startColumnIndex: index,
              endColumnIndex: index + 1,
            },
            cell: {
              note: column.note,
            },
          },
        });
      }
    });
    return format;
  }
  async addFormattingAndValidations(spreadsheetId: string): Promise<void> {
    if (!this.sheetId) {
      throw new Error("Sheet ID is not available");
    }
    const requests = [
      ...this.getHeaderNotesConfig(),
      ...this.getGroupHeaderFormatConfig(),
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
  async appendRows(rowsData: Record<string, string>[]) {
    const rowsRequestData: any[][] = [];
    rowsData.forEach((rowData) => {
      rowsRequestData.push(this.addRow(rowData));
    });
    await sheetService(this.spreadsheet.googleAuth).spreadsheets.values.append({
      spreadsheetId: this.spreadsheet.spreadsheetId,
      range: `${this.name}!A1`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: rowsRequestData,
      },
    });
    console.log(">> Rows added");
  }
  private addRow(rowData: Record<string, string>) {
    const currentColumnNames = this.getHeaderNames();
    const rowKeys = Object.keys(rowData);
    const requestBodyRowData: any[] = [];
    const allColumnNameMatches = rowKeys.every((columnName) => {
      return currentColumnNames.includes(columnName);
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
    currentColumnNames.forEach((columnName) => {
      requestBodyRowData.push(rowData[columnName]);
    });
    return requestBodyRowData;
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
        range: `${this.name}!${SheetUtils.generateSheetRange(
          this.properties?.columnCount || 0,
          this.properties?.rowCount || 0
        )}`,
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
        range: `${this.name}!${SheetUtils.generateSheetRange(
          this.properties?.columnCount || 0,
          this.properties?.rowCount || 0,
          2
        )}`,
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
