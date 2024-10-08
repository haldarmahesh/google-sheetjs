import { Column } from "./column.entity";

export class ColumnGroup {
  constructor(public name: string, public columns: Column[]) {}
}