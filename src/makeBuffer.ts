import { InterceptFn, Sheet } from "./types"

export default async function makeBuffer<T extends Array<Sheet>>({
  worksheets,
  data,
  intercept,
}: {
  worksheets: T
  data: Array<any> | Record<T[number]["name"], Array<any>>
  intercept?: InterceptFn
}) {
  const { default: {ExcelJS} } = await import("./deps")
  let workbook = new ExcelJS.Workbook()
  for (const worksheet of worksheets) {
    const sheet = workbook.addWorksheet(worksheet.name)
    sheet.columns = worksheet.columns
    const rows = (Array.isArray(data) ? data : data[worksheet.name as T[number]["name"]]) ?? []
    sheet.addRows(rows)
  }
  workbook = intercept ? intercept(workbook) ?? workbook : workbook
  return await workbook.xlsx.writeBuffer()
}
