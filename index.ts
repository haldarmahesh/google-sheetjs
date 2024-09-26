import { googleAuth } from "./google-auth";

const { google } = require("googleapis");
const auth = googleAuth;

const service = google.sheets({ version: "v4", auth });
async function create(title: string) {
  try {
    const driveService = google.drive({ version: "v3", auth });

    const spreadsheet = await service.spreadsheets.create({
      resource: {
        properties: {
          title,
        },
        sheets: [
          {
            properties: {
              title: "Contracts",
            },
          },
        ],
      },
      fields: "spreadsheetId",
    });
    console.log("?? RESPONSE", spreadsheet.data);
    console.log(`Spreadsheet ID: ${spreadsheet.data.spreadsheetId}`);

    await driveService.permissions.create({
      resource: {
        type: "domain",
        role: "writer",
        domain: "sennder.com", // Replace with your domain
      },
      fileId: spreadsheet.data.spreadsheetId,
      fields: "id",
    });

    return spreadsheet.data.spreadsheetId;
  } catch (err) {
    // TODO (developer) - Handle exception
    throw err;
  }
}

async function addHeader(spreadsheetId: string) {
  const protectedHeaders = ["contract_agreement_id", "transport_contract_id"];
  await service.spreadsheets.values.update({
    spreadsheetId,
    range: "Sheet1!A1:AV",
    valueInputOption: "RAW",
    requestBody: {
      values: [
        [
          ...[
            "shipper_contract_name",
            "contract_type",
            "entity_uuid",
            "contact_uuiid",
            "priority",
            "status",
            "validity_start",
            "validity_end",
            "base_cost",
          ],
          ...protectedHeaders,
        ],
      ],
    },
  });

  const requests = [
    {
        addProtectedRange: {
            protectedRange: {
              range: {
                sheetId: 0,
                startRowIndex: 0,
                endRowIndex: 1,
                startColumnIndex: 0,
              },
              description: "Protecting header row",
              warningOnly: false,
              requestingUserCanEdit: true,
              editors: {
                users: ["mahesh.haldar@sennder.com"],
                domainUsersCanEdit: false,
              },
            },
          },
    },
    {
      addProtectedRange: {
        protectedRange: {
          range: {
            sheetId: 0,
            startRowIndex: 0,
            startColumnIndex: 9,
            endColumnIndex: 11,
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
    },
    {
        repeatCell: {
          range: {
            sheetId: 0, // Assuming this is the first sheet (Sheet1)
            startRowIndex: 0, // Start at the second row (index 1)
            startColumnIndex: 9,
            endColumnIndex: 11,
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
      },
    {
      repeatCell: {
        range: {
          sheetId: 0, // Assuming this is the first sheet (Sheet1)
          startRowIndex: 0, // Start at the second row (index 1)
          endRowIndex: 1, // End at the third row (index 2)
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
    {
      repeatCell: {
        range: {
          sheetId: 0, // Assuming this is the first sheet (Sheet1)
          startRowIndex: 1, // Start at the second row (index 1)
          startColumnIndex: 1, // Start at the third column (index 2)
          endColumnIndex: 2, // End at the third column (index 3)
        },
        cell: {
          dataValidation: {
            condition: {
              type: "ONE_OF_LIST",
              values: [
                { userEnteredValue: "CARRIER" },
                { userEnteredValue: "CHARTERING" },
              ],
            },
            strict: true, // Enforces the validation rule
            showCustomUi: true, // Shows a dropdown UI
          },
        },
        fields: "dataValidation",
      },
    },
    {
      repeatCell: {
        range: {
          sheetId: 0, // Assuming this is the first sheet (Sheet1)
          startRowIndex: 1, // Start at the second row (index 1)
          startColumnIndex: 5, // Start at the third column (index 2)
          endColumnIndex: 6, // End at the third column (index 3)
        },
        cell: {
          dataValidation: {
            condition: {
              type: "ONE_OF_LIST",
              values: [
                { userEnteredValue: "ACTIVE" },
                { userEnteredValue: "INACTIVE" },
              ],
            },
            strict: true, // Enforces the validation rule
            showCustomUi: true, // Shows a dropdown UI
          },
        },
        fields: "dataValidation",
      },
    },
  ];
  await service.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests,
    },
  });
}
async function readSheet(googleSheetId: string) {
  try {
    const sheetInstance = await google.sheets({
      version: "v4",
      auth: googleAuth,
    });

    const infoObjectFromSheet = await sheetInstance.spreadsheets.values.get({
      auth: googleAuth,
      spreadsheetId: googleSheetId,
      range: `Sheet1!A2:N1000`,
    });
    const valuesFromSheet = infoObjectFromSheet.data.values;
    console.log(valuesFromSheet);
  } catch (err) {
    console.log("err", err);
    throw err;
  }
}
async function writeSheet(googleSheetId: string) {
  try {
    const sheetInstance = await google.sheets({
      version: "v4",
      auth: googleAuth,
    });

    const updateToGsheet = [
      ["Name1", "CARRIER", "", "", "1", "ACTIVE", "2024-30-01"],
      ["Name2", "CHARTERING", "", "", "2", "ACTIVE", "2024-30-01"],
      ["Name3", "CARRIER", "", "", "3", "ACTIVE", "2024-30-01"],
    ];

    await sheetInstance.spreadsheets.values.update({
      auth: googleAuth,
      spreadsheetId: googleSheetId,
      range: `Sheet1!A2:H1000`,
      valueInputOption: "RAW",
      resource: {
        values: updateToGsheet,
      },
    });
  } catch (err) {
    console.log("err", err);
    throw err;
  }
}
const spreadsheetId = await create(
  `tcm-${new Date().toISOString().slice(0, 19)}`
);
// const spreadsheetId = "1o79YFEU2OiMoi6Wz---7T7cjCGGgyYVuW8ivCLpHbS0";
// await addHeader(spreadsheetId);

// await readSheet(spreadsheetId);
// await writeSheet(spreadsheetId);
