import { google } from "googleapis";
import { GoogleAuth } from "../../core/entities/google-auth.entity";

export const sheetService = (googleAuth: GoogleAuth) => {
    return google.sheets({ version: "v4", auth: googleAuth.authInstance})}