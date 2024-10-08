import { driveService } from "../../infrastructure/drive-service";
import { sheetService } from "../../infrastructure/googlesheet-service";
import { GoogleAuth } from "./google-auth.entity";
import { Sheet, SheetProperties } from "./sheet.entity";

export class Spreadsheet {
  spreadsheetId?: string;
  public googleAuth: GoogleAuth;
  public sheets: Sheet[] = [];
  constructor(
    public title: string,
    public version: string = "1.0.0",
    {
      serviceAccountEmail,
      privateKey,
    }: { serviceAccountEmail: string; privateKey: string }
  ) {
    this.googleAuth = new GoogleAuth(serviceAccountEmail, privateKey);
    this.version = version;
  }
  public async loadProperties(spreadsheetId: string): Promise<void> {
    this.spreadsheetId = spreadsheetId;
    try {
      const spreadsheetData = await sheetService(
        this.googleAuth
      ).spreadsheets.get({
        spreadsheetId: this.spreadsheetId,
      });
      if (!spreadsheetData.data.properties?.title) {
        throw new Error("Spreadsheet title is not found");
      }
      this.title = spreadsheetData.data.properties?.title;
      // load sheets props
      if (spreadsheetData.data.sheets?.length) {
        spreadsheetData.data.sheets.forEach((sheet: any, index: number) => {
          const newSheet = new Sheet(
            sheet.properties.title,
            this,
            new SheetProperties(
              sheet.properties.gridProperties.rowCount,
              sheet.properties.gridProperties.columnCount
            )
          );
          newSheet.groupHeaderConfig.enabled = Boolean(
            spreadsheetData.data.developerMetadata?.find(
              (item) => item.metadataKey === "groupingEnabled"
            )?.metadataValue
          );
          newSheet.sheetId = sheet.properties.sheetId;
          this.sheets.push(newSheet);
          this.sheets[index].loadHeader();
        });

        await Promise.all(
          spreadsheetData.data.sheets.map(async (sheet, index: number) => {
            if (!sheet?.properties?.sheetId) {
              throw new Error("Sheet ID is not found");
            }
            const newSheet = new Sheet(
              sheet.properties?.title || "N/A",
              this,
              new SheetProperties(
                sheet.properties?.gridProperties?.rowCount || 1,
                sheet.properties?.gridProperties?.columnCount || 1
              )
            );
            newSheet.groupHeaderConfig.enabled = Boolean(
              spreadsheetData.data.developerMetadata?.find(
                (item) => item.metadataKey === "groupingEnabled"
              )?.metadataValue
            );
            newSheet.sheetId = sheet.properties.sheetId;
            this.sheets.push(newSheet);
            await this.sheets[index].loadHeader();
          })
        );
      }
      // /// remove =======
      // const res = await sheetService(this.googleAuth).spreadsheets.values.get({
      //   spreadsheetId: this.spreadsheetId,
      //   range: `${this.sheets[0].name}!F4`
      // })
      // console.log('>> res', JSON.stringify(res))
      // /// remove =======
    } catch (err: any) {
      if (err.code === 404) {
        throw new Error("Spreadsheet not found, pass a valid spreadsheet ID");
      }
      throw err;
    }
  }
  public addSheet(name: string): Sheet {
    const sheet = new Sheet(name, this);
    this.sheets.push(sheet);
    return sheet;
  }
  public isCreated(): boolean {
    return !!this.spreadsheetId;
  }
  public async createSpreasheet(): Promise<string> {
    try {
      if (this.sheets.length === 0) {
        throw new Error(
          "Sheet is missing. At least one sheet is required to create a spreadsheet, use class Sheet to add a sheet"
        );
      }
      const spreadsheet: any = await sheetService(
        this.googleAuth
      ).spreadsheets.create({
        resource: {
          properties: {
            title: this.title,
          },
          developerMetadata: [
            {
              metadataKey: "version",
              metadataValue: this.version,
              visibility: "DOCUMENT",
            },
            {
              metadataKey: "groupingEnabled",
              metadataValue:
                this.sheets[0].groupHeaderConfig.enabled.toString(), // TODO: do we need for all the sheets?
              visibility: "DOCUMENT",
            },
          ],
          sheets: this.sheets.map((sheet) => ({
            properties: {
              title: sheet.name,
              gridProperties: {
                rowCount: 1000,
                columnCount: sheet.columns.length,
              },
            },
          })),
        },
      } as any);

      console.log(`Spreadsheet ID: ${spreadsheet.data.spreadsheetId}`);

      await driveService(this.googleAuth).permissions.create({
        resource: {
          type: "domain",
          role: "writer",
          domain: "sennder.com",
        },
        fileId: spreadsheet.data.spreadsheetId,
        fields: "id",
      } as any);
      spreadsheet.data.sheets?.forEach((sheet: any, index: number) => {
        this.sheets[index].sheetId = sheet.properties?.sheetId;
      });
      return spreadsheet.data.spreadsheetId;
    } catch (err) {
      console.log(">> err1", err);
      throw err;
    }
  }

  private async addHeaders(sheetIndex: number): Promise<void> {
    if (!this.spreadsheetId) {
      throw new Error("Spreadsheet ID is not available");
    }
    this.sheets[sheetIndex].properties = new SheetProperties(
      1000,
      this.sheets[sheetIndex].columns.length
    );
    await this.sheets[sheetIndex].addHeaders(this.spreadsheetId);
    await this.sheets[sheetIndex].addFormattingAndValidations(
      this.spreadsheetId
    );
  }
  private async checkAndGetRequiredMetadata() {
    if (!this.spreadsheetId) {
      throw new Error("Spreadsheet ID is not available");
    }
    const sheetIdMissing = this.sheets.find((sheet) => !sheet.sheetId);
    if (sheetIdMissing) {
      const spreadsheetData: any = await sheetService(
        this.googleAuth
      ).spreadsheets.get({
        spreadsheetId: this.spreadsheetId,
      });
      spreadsheetData.data.sheets.forEach((sheet: any, index: number) => {
        this.sheets[index].sheetId = sheet.properties?.sheetId;
      });
    }
  }
  async create() {
    this.spreadsheetId = await this.createSpreasheet();
    // this.spreadsheetId = "1vOLEtYeGk7RQmzRidPlLRsWq9hEpngDajCEwk30JFkg";
    await this.checkAndGetRequiredMetadata();
    await this.addHeaders(0);
  }
}
