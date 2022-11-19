// @ts-nocheck
import dayjs from 'dayjs'
import { ExcelUtil } from '../export/ExcelUtils'
// 导出Excel
export const exportExcel = (
  fileName: string,
  type: string,
  tableData: Array<any>,
  tableColumns: Array<any>
) => {
  const ExcelHeader: any = []
  for (let i = 0; i < tableColumns.length!; i++) {
    if (tableColumns[i].type !== 'default') continue
    if (tableColumns[i].property === undefined) continue
    ExcelHeader.push([tableColumns[i].label])
  }
  const lstData: any = []
  const reg = /[a-zA-Z]+/
  for (const datum of tableData) {
    const rowData: Array<any> = []
    for (const dataKey in datum) {
      if (
        !reg.test(datum[dataKey]) &&
        new Date(datum[dataKey]).toLocaleString() !== '1970/1/1 08:00:00'
      ) {
        rowData.push(
          dayjs(new Date(datum[dataKey])).format('YYYY-MM-DD hh:mm:ss')
        )
      } else {
        rowData.push(datum[dataKey])
      }
    }
    lstData.push(rowData)
  }
  const lstSheet: Array<any> = []
  ExcelUtil.writeHeader(ExcelHeader, undefined, 1)
  for (let num = 1; num <= ExcelHeader.length; num++) {
    for (let i = 0, rowIndex = 1; i < lstData.length; i++, rowIndex++) {
      let colIndex = 0
      for (let j = 0; j < lstData[i].length; j++) {
        ExcelUtil.writeCellData(rowIndex, colIndex++, lstData[i][j])
      }
    }
    const sheet = ExcelUtil.getDataSheet()
    lstSheet.push({ sheet, name: `Sheet${num}` })
  }

  ExcelUtil.write(lstSheet, `${fileName}.${type}`)
}
export const importExcel = (data: any, columns: Array<any>): Array<any> => {
  const lstData: Array<any> = ExcelUtil.toMap(data, 0)
  //收集 table 的 prop 数据
  const tableProp: any = {}
  const newTableData: Array<any> = []
  for (const column of columns) {
    tableProp[column.label] = column.property
  }
  for (const lstDatum of lstData) {
    const value: any = {}
    for (const key in tableProp) {
      value[tableProp[key]] = lstDatum.get(key)
    }
    newTableData.push(value)
  }
  return newTableData
}
