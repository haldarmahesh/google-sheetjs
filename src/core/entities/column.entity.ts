import type { DataValidation } from "./data-validation.entity";

export class Column {
  constructor(
    public name: string,
    public validation?: DataValidation,
    public isProtected?: boolean
  ) {}
}
