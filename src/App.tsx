import { useExcelJS } from "./react-use-exceljs"

const data = [
  { id: 1, name: "Jane Doe", dob: new Date(1984, 6, 7) },
  { id: 2, name: "John Doe", dob: new Date(1965, 1, 7) },
]

function App() {
  const excel = useExcelJS({
    filename: "report.xlsx",
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
    void excel.download(data)
  }
  return (
    <div className="App">
      <button onClick={onClick}>Download</button>
    </div>
  )
}

export default App
