import { expect, test } from "vitest"
import ExcelJS from "exceljs"
import makeBuffer from "../src/makeBuffer"

const worksheetA = {
  name: "Sheet 1",
  columns: [
    {
      header: "Id",
      key: "id",
      width: 10,
    },
    {
      header: "Name",
      key: "name",
      width: 32,
    },
    {
      header: "D.O.B.",
      key: "dob",
      width: 200,
    },
  ],
}

const worksheetB = {
  name: "Cities",
  columns: [
    {
      header: "City",
      key: "city",
    },
  ],
}

test("can create one sheet", async () => {
  const buffer =  await makeBuffer({
    data: [
      { id: 1, name: "Jane Doe", dob: new Date(1984, 6, 7) },
      { id: 2, name: "John Doe", dob: new Date(1965, 1, 7) },
    ],
    worksheets: [worksheetA],
  })

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.getWorksheet(worksheetA.name)

  expect(worksheet.getCell('A1').text).toBe('Id')
  expect(worksheet.getCell('A2').text).toBe('1')
  expect(worksheet.getCell('A3').text).toBe('2')
})

test("can create two sheets", async () => {
  const buffer =  await makeBuffer({
    data: {
      [worksheetA.name]: [
        { id: 1, name: "Jane Doe", dob: new Date(1984, 6, 7) },
        { id: 2, name: "John Doe", dob: new Date(1965, 1, 7) },
      ],
      [worksheetB.name]: [
        { city: 'Dallas' },
        { city: 'New York' },
        { city: 'Miami' },
      ],
    },
    worksheets: [worksheetA, worksheetB],
  })

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheetA = workbook.getWorksheet(worksheetA.name)

  expect(sheetA.getCell('A1').text).toBe('Id')
  expect(sheetA.getCell('A2').text).toBe('1')
  expect(sheetA.getCell('A3').text).toBe('2')

  const sheetB = workbook.getWorksheet(worksheetB.name)

  expect(sheetB.getCell('A1').text).toBe('City')
  expect(sheetB.getCell('A2').text).toBe('Dallas')
  expect(sheetB.getCell('A3').text).toBe('New York')
  expect(sheetB.getCell('A4').text).toBe('Miami')
})

test("can use intercept", async () => {
  const buffer =  await makeBuffer({
    data: [
      { id: 1, name: "Jane Doe", dob: new Date(1984, 6, 7) },
      { id: 2, name: "John Doe", dob: new Date(1965, 1, 7) },
    ],
    worksheets: [worksheetA],
    intercept: (workbook) => {
      workbook.getWorksheet(worksheetA.name).getCell('A2').value = 'foobar'
    }
  })

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.getWorksheet(worksheetA.name)

  expect(worksheet.getCell('A1').text).toBe('Id')
  expect(worksheet.getCell('A2').text).toBe('foobar')
  expect(worksheet.getCell('A3').text).toBe('2')
})
