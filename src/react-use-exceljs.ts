import { type Workbook, type Worksheet } from "exceljs"
import { saveAs } from "file-saver"
import React from "react"

type Sheet = { name: string; columns: Worksheet["columns"] }
type InterceptFn = (workbook: Workbook) => Workbook | void
type Filename = `${string}.xlsx`

export default function useExcelJS<T extends Array<Sheet>>({
  filename,
  worksheets,
  intercept,
}: {
  worksheets: T
  filename?: Filename
  intercept?: InterceptFn
}) {
  type WorksheetName = T[number]["name"]
  type Data = Array<any> | Record<WorksheetName, Array<any>>

  return {
    download: React.useCallback(
      async (data: Data) => {
        const ExcelJS = await import("exceljs")
        let workbook = new ExcelJS.Workbook()
        for (const worksheet of worksheets) {
          const sheet = workbook.addWorksheet(worksheet.name)
          sheet.columns = worksheet.columns
          const rows = (Array.isArray(data) ? data : data[worksheet.name as WorksheetName]) ?? []
          sheet.addRows(rows)
        }
        workbook = intercept ? intercept(workbook) ?? workbook : workbook
        const buffer = await workbook.xlsx.writeBuffer()
        const fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        const blob = new Blob([buffer], { type: fileType })
        saveAs(blob, filename ?? "workbook.xlsx")
      },
      [filename, worksheets, intercept]
    ),
  }
}
