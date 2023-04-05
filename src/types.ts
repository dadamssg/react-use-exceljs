import { type Workbook, type Worksheet } from "exceljs"

export type Sheet = { name: string; columns: Worksheet["columns"] }
export type InterceptFn = (workbook: Workbook) => Workbook | void
export type Filename = `${string}.xlsx`
export type WorksheetName<T extends Array<Sheet>> = T[number]["name"]
export type Data<T extends Array<Sheet>> =
  | Array<any>
  | Record<WorksheetName<T>, Array<any>>
