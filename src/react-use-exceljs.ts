import React from "react"
import { type Data, type Filename, type InterceptFn, type Sheet } from "./types"
import makeBuffer from "./makeBuffer"

export function useExcelJS<T extends Array<Sheet>>({
  filename,
  worksheets,
  intercept,
}: {
  worksheets: T
  filename?: Filename
  intercept?: InterceptFn
}) {
  return {
    download: React.useCallback(
      (data: Data<T>) =>
        downloadExcelJS({ filename, data, worksheets, intercept }),
      [filename, worksheets, intercept]
    ),
  }
}

export async function downloadExcelJS<T extends Array<Sheet>>({
  filename,
  data,
  worksheets,
  intercept,
}: {
  worksheets: T
  data: Data<T>
  filename?: Filename
  intercept?: InterceptFn
}) {
  const buffer = await makeBuffer({ worksheets, data, intercept })
  const fileType =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  const blob = new Blob([buffer], { type: fileType })
  const {
    default: { saveAs },
  } = await import("./deps")
  saveAs(blob, filename ?? "workbook.xlsx")
}
