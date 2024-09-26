const { google } = require("googleapis");

const clientEmail =
  "tcm-matching-bulk-upload@tcm-matching.iam.gserviceaccount.com";
const privateKey =
  "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQC9TsjHKJ243KeJ\n69GEf8myl/p4J5CBp4aah4Qmr9Dy3bMC77UFPa2cuM8pBiFdy/cIxfHselIqBzSN\ngWzfbmK9BN7dtczhrJZ3Pi2r7z4Cg1oH811QKzYBnuQt6zR57i0mzYhgX7GpVFZZ\no0gdSLis8abda2Pz/+ZsSWrwM0bvtfkFt3pomHFABr6wVsOQr3RmdRFGQSK1B7SX\nnwMQZ8xeXkPyzpE132JZmN0K/h8JRDOs24nf+45jRDTTEKHkmE73dJBO8HAye/Yh\nB612aVaMBjM99pVImkptEb1jpEUVoyO+2LapyGXGXr10lk7j+D3WaGpytvkrLeR9\nODl5OXbHAgMBAAECggEAWQD8E+rsMLiJiqZVRPsBvRaIO8q6PcMiXW/+eWPrFOyY\nF/bUgIjNoeQf/fU0ZdGaLUVHp3uhOsJVenxR0ECpap4qHEo38BiBS8Hvnikm2e6g\nuyE4C5OtWhi2xkIR04vgLaCvkEQdlvOgf5dttdr1fNZGsk6l2VfEob/o59Lr9JIk\nyFRndSftX1f+qMxUp+4zqofycIm6n5aVGZIVRRCuXyLuycgTFpHd3VCYccParQj5\njLigDxQhnwk9m57Vulz6qzloaVewNw8B+09xNr65RING5eORT7GV9qvBXmkn6ZqE\nGz639nQpCpP7rE/Tay2kAryPNSAudC3vgqifstEKGQKBgQD5fB1uq+E/wUJLWdhX\n1j+FQUg+EOfivrEXIsRjtPbvbWaGMDTm1rCLhd9Oc6UbHWYYxKlujddiDZoS0nNF\nsPcS3vXhiMod0Ixwkc6YVwUIsFiZ6Od6mJSBtJ1jldxJn4AoYir3WXs2djuZBYnl\neGIr4bH7pjrUYJs3Y76GPnWqLwKBgQDCQF3VRYEyhe4lUxIlr97IbTDOEd/3xddK\n7s5cWXFw5bUa4uREWDAPoBbHPyxYnPQ2fB6kZNiMmgjs5IQ1PxoVRfyE87hQqBKV\nHTiWhNluNoAHUP4LrStnHt+aPvPkqDpjia+xowKAW9KdtCP0pZcZSw1Ud9pq4rST\noSyLzp0O6QKBgQCUYHWcuYqogUbtS4z4iIqUtQPDLgjLeQAXs2y7pAfs09LS4d7E\nn1C2WjM6FFtQqgZrmqLuBlvfjBljMliuTRZU2dfAf7s9SigMVxtYzQBIb6DyQGtT\nJWXFUmb8sEcoXj05R1EodMZr2JuPYZTmrdctI/jXosCASMhng+HvMzyFrwKBgFE9\nfEDi9bq8mrHPgUpzuFfYms3EWggVHQqAv5uN6MzPtSOOeus+erM+P+iKujBBTD2x\nQVt9tbdwAIWauNRQFMeK4qZ0C8Tn1gW5F96Tpx/Z+UeWDvmxLfLNzbSD2Zrq5KiW\nf/1p8HTgckB0g4kg7AWvBt8p1RZYxC7t/GRoP/VpAoGBAKHBQo/yNF+gSJMB0Lfh\nSyCTQxolnszEe21ED9zD2yICuzMvyh2tEtVUMfHOC5wV5vrspCgLy5mqH9qG+u3L\nUBJZpqnKDWbh4FYbMKV44H8oW1KfELvOV4UsjlJLEDfbWEeXjYV5HLUNbhb1NWpB\nLXtWhWltkx5O3e3I1otHpoFa\n-----END PRIVATE KEY-----\n";
const googleSheetId = "your google sheet id";
const googleSheetPage = "your google sheet page name";

// authenticate the service account
export const googleAuth = new google.auth.JWT(
  clientEmail,
  null,
  privateKey.replace(/\\n/g, "\n"),
  [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
  ]
);
