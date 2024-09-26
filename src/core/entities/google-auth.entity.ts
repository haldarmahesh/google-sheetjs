const { google } = require("googleapis");
export class GoogleAuth {
    // private static instance: GoogleAuth;
    constructor(readonly serviceAccountEmail: string, readonly privateKey: string) {
        // if (GoogleAuth.instance) {
        //     return GoogleAuth.instance;
        // }
        // GoogleAuth.instance = this;
    }

    get authInstance():any {
        return new google.auth.JWT(
            this.serviceAccountEmail,
            null,
            this.privateKey.replace(/\\n/g, "\n"),
            [
              "https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive",
            ]
          );
    }
}