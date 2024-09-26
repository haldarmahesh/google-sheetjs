import { driveService, sheetService } from "../../../services";
import type { Sheet } from "./sheet.entity";

export class Spreadsheet {
  spreadsheetId?: string;
  title: string;
  constructor(public sheets: Sheet[]) {
    this.title = `tcm-${new Date().toISOString().slice(0, 19)}`;
  }
  private async createSpreasheet(): Promise<string> {
    try {
      const spreadsheet: any = await sheetService.spreadsheets.create({
        resource: {
          properties: {
            title: this.title,
          },
          sheets: this.sheets.map((sheet) => ({
            properties: {
              title: sheet.sheetName,
            },
          })),
        },
      } as any);
      console.log(`Spreadsheet ID: ${spreadsheet.data.spreadsheetId}`);

      await driveService.permissions.create({
        resource: {
          type: "domain",
          role: "writer",
          domain: "sennder.com",
        },
        fileId: spreadsheet.data.spreadsheetId,
        fields: "id",
      } as any);
      spreadsheet.data.sheets?.forEach((sheet: any, index: number) => {
        this.sheets[index].sheetId = sheet[index].properties?.sheetId;
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
      const spreadsheetData: any = await sheetService.spreadsheets.get({
        spreadsheetId: this.spreadsheetId,
      });
      spreadsheetData.data.sheets.forEach((sheet: any, index: number) => {
        this.sheets[index].sheetId = sheet.properties?.sheetId;
      });
    }
  }
  async create() {
    // this.spreadsheetId = await this.createSpreasheet();
    this.spreadsheetId = "1ZPizNoPvzKub96vrnb3q16L7JTic4Gcy_YjdAl41TwU";
    await this.checkAndGetRequiredMetadata();
    await this.addHeaders(0);
  }
}
