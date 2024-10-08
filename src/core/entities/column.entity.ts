import type { DataValidation } from "./data-validation.entity";

export class Column {
  public validation?: DataValidation;
  public isProtected: boolean;
  constructor(public name: string, public note?: string) {
    this.isProtected = false;
  }
  public setValidation(validation: DataValidation): Column {
    this.validation = validation;
    return this;
  }
  public protect(): Column {
    this.isProtected = true;
    return this;
  }
}
