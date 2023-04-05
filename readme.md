# react-use-exceljs

A thin wrapper around the [exceljs](https://github.com/exceljs/exceljs) package that uses [file-saver](https://github.com/eligrey/FileSaver.js) for generating excel files. 

## Usage

```tsx
import { useExcelJS } from "react-use-exceljs"

const data = [
  { id: 1, name: "Jane Doe", dob: new Date(1984, 6, 7) },
  { id: 2, name: "John Doe", dob: new Date(1965, 1, 7) },
]

function App() {
  const excel = useExcelJS({
    filename: "testing.xlsx",
    worksheets: [
      {
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
      },
    ],
  })

  const onClick = () => {
    excel.download(data)
  }
  
  return (
    <div>
      <button onClick={onClick}>Download</button>
    </div>
  )
}
```

Worksheets can use different sets of data if `download()` is passed a `Record<SheetName, Array<any>>`.

```tsx
const people = [
  { id: 1, name: "Jane Doe", dob: new Date(1984, 6, 7) },
  { id: 2, name: "John Doe", dob: new Date(1965, 1, 7) },
]

const cities = [
  { city: 'Dallas' },
  { city: 'New York' },
  { city: 'Miami' },
]

function App() {
  const excel = useExcelJS({
    filename: "testing.xlsx",
    worksheets: [
      {
        name: "People",
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
      },
      {
        name: "Cities",
        columns: [
          {
            header: "City",
            key: "city",
          },
        ],
      },
    ],
  })

  const onClick = () => {
    excel.download({ People: people, Cities: cities })
  }
  
  return (
    <div>
      <button onClick={onClick}>Download</button>
    </div>
  )
}
```

You can pass an `intercept` function that will provide the generated workbook for you to modify before it is downloaded. 

```tsx
const excel = useExcelJS({
  filename: "testing.xlsx",
  intercept: (workbook) => {
    workbook.getWorksheet("Sheet 1").getColumn("id").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "55e6ed" },
    }
  },
  worksheets: [
    {
      name: "Sheet 1",
      columns: [
        {
          header: "Id",
          key: "id",
          width: 10,
        },
      ],
    },
  ],
})
```

You can also use the non-hook function version which the `useExcelJS` uses internally.

```ts
import { downloadExcelJS } from 'react-use-exceljs'

const onClick = () => {
  downloadExcelJS({
    filename: "testing.xlsx",
    data: [
      {id: 1},
      {id: 2}
    ],
    worksheets: [
      {
        name: "Sheet 1",
        columns: [
          {
            header: "Id",
            key: "id",
            width: 10,
          },
        ],
      },
    ],
  }) 
}
```


## Optimization
The `file-saver` and the rather large `exceljs` packages are lazily loaded on initiation of an excel download.
