import type { ValidationType } from "../enums/validation-type.enum";

export class DataValidation {
  constructor(public type: ValidationType, public values: string[]) {}
}
