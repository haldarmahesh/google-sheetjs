
import { driveService } from "../../infrastructure/drive-service";
import { sheetService } from "../../infrastructure/googlesheet-service";
import { GoogleAuth } from "./google-auth.entity";
import { Sheet } from "./sheet.entity";

export class Spreadsheet {
  spreadsheetId?: string;
  public googleAuth: GoogleAuth;
  public sheets: Sheet[] = [];
  constructor(public title: string, {
    serviceAccountEmail,
    privateKey,
  }: {serviceAccountEmail: string, privateKey: string}) {
    this.googleAuth = new GoogleAuth(serviceAccountEmail, privateKey);
  }
  public addSheet(name: string): Sheet {
    const sheet = new Sheet(name, this);
    this.sheets.push(sheet);
    return sheet;
  }
  public isCreated(): boolean {
    return !!this.spreadsheetId; 
  }
  private async createSpreasheet(): Promise<string> {
    try {
      const spreadsheet: any = await sheetService(this.googleAuth).spreadsheets.create({
        resource: {
          properties: {
            title: this.title,
          },
          sheets: this.sheets.map((sheet) => ({
            properties: {
              title: sheet.name,
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
      const spreadsheetData: any = await sheetService(this.googleAuth).spreadsheets.get({
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
