import { google } from "googleapis";
import type { GoogleAuth } from "../../core/entities/google-auth.entity";

export const driveService =(googleAuth: GoogleAuth) => google.drive({ version: "v3", auth: googleAuth.authInstance});
