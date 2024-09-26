import { google } from "googleapis";
import { googleAuth } from "./google-auth";

export const sheetService = google.sheets({ version: "v4", auth: googleAuth });
export const driveService = google.drive({ version: "v3", auth: googleAuth });
