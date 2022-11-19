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
  //表格头部数据
  const ExcelHeader: any = []
  // 表格对顶table props的值
  const ExcelProps: any = {}
  for (let i = 0; i < tableColumns.length!; i++) {
    if (tableColumns[i].type !== 'default') continue
    if (tableColumns[i].property === undefined) continue
    ExcelHeader.push([tableColumns[i].label])
    ExcelProps[tableColumns[i].property] = tableColumns[i].property
  }
  ExcelHeader.unshift(['序号'])
  const lstData: any = []
  const reg = /^\d+$/
  for (const datum of tableData) {
    const rowData: Array<any> = []
    for (const dataKey in ExcelProps) {
      if (reg.test(datum[dataKey]) && datum[dataKey].toString().length === 13) {
        rowData.push(dayjs(new Date(datum[dataKey])).format('YYYY-MM-DD H:m:s'))
      } else {
        rowData.push(datum[dataKey])
      }
    }
    lstData.push(rowData)
  }
  const lstSheet: Array<any> = []
  for (let num = 1; num <= ExcelHeader.length; num++) {
    ExcelUtil.writeHeader(ExcelHeader)
    for (let i = 0, rowIndex = 1; i < lstData.length; i++, rowIndex++) {
      let colIndex = 0
      ExcelUtil.writeCellData(rowIndex, colIndex++, `${i + 1}`)
      for (let j = 0; j < lstData[i].length; j++) {
        ExcelUtil.writeCellData(rowIndex, colIndex++, lstData[i][j])
      }
    }
    const sheet = ExcelUtil.getDataSheet()
    lstSheet.push({ sheet, name: `Sheet${num}` })
  }
  ExcelUtil.write(
    lstSheet,
    `${fileName}${dayjs(new Date()).format('YYYY-MM-DD hh:mm:ss')}.${type}`
  )
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
