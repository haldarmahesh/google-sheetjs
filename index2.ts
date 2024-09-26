import { google } from "googleapis";
import { DataValidation } from "./src/core/entities/data-validation.entity";
import { Sheet } from "./src/core/entities/sheet.entity";
import { ValidationType } from "./src/core/enums/validation-type.enum";
import { googleAuth } from "./google-auth";
import { Spreadsheet } from "./src/core/entities/spreadsheet.entity";
import { Column } from "./src/core/entities/column.entity";

const spreadsheet1 = new Spreadsheet([
  new Sheet("Contracts", [
    new Column("shipper_contract_name"),
    new Column(
      "contract_type",
      new DataValidation(ValidationType.ONE_OF_LIST, ["CARRIER", "CHARTERING"])
    ),
    new Column("entity_uuid"),
    new Column("contact_uuiid"),
    new Column("priority"),
    new Column(
      "status",
      new DataValidation(ValidationType.ONE_OF_LIST, ["ACTIVE", "INACTIVE"])
    ),
    new Column("validity_start"),
    new Column("validity_end"),
    new Column("base_cost"),
    new Column("contract_agreement_id", undefined, true),
    new Column("transport_contract_id", undefined, true),
  ]),
]);

// await spreadsheet1.create();
// await spreadsheet1.addData([
//   [1, 2],
//   [2, 3],
// ]);
await spreadsheet1.sheets[0].addData(">>", [
  [1, 2],
  [1, 2],
]);
