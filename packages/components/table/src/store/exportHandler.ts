// @ts-nocheck
import dayjs from 'dayjs'
import { ExcelUtil } from '../export/ExcelUtils'
// 导出Excel
export const exportExcel = (
  fileName: string,
  type: string,
  tableData: Array<any>,
  myColumns: Array<any>,
  tableColumns: Array<any>
) => {
  //表格头部数据
  const ExcelHeader: any = []
  // 表格对顶table props的值
  const ExcelProps: any = {}
  if (myColumns.length > 1) {
    const rowsLength = myColumns.length
    const columnsLength = tableColumns.length
    const headerList: any = []
    for (let i = 0; i < rowsLength; i++) {
      const list = []
      for (let j = 0; j < columnsLength; j++) {
        list.push(false)
      }
      headerList.push(list)
    }
    for (const [i, cl] of myColumns.entries()) {
      for (const [j, element] of cl.entries()) {
        for (let k = 0; k < element.rowSpan; k++) {
          let l = 0
          while (headerList[i + k][j + l] !== false && headerList[i + k][j]) {
            l++
          }
          if (headerList[i + k][j + l] === false) {
            headerList[i + k][j + l] = element.label
          }
        }
        for (let k = 0; k < element.colSpan - 1; k++) {
          let l = 0
          while (
            headerList[i][k + j + l] !== false &&
            headerList[i][k + j + l]
          ) {
            l++
          }
          if (headerList[i][k + j + l] === false && element.colSpan > 1) {
            headerList[i][k + j + l] = element.label
          }
        }
      }
    }
    for (let i = 0; i < headerList[0].length; i++) {
      const list: any = []
      for (const element of headerList) {
        list.push(element[i])
      }
      ExcelHeader.push(list)
    }
    for (let i = 0; i < tableColumns.length!; i++) {
      if (tableColumns[i].type !== 'default') continue
      if (tableColumns[i].property === undefined) continue
      ExcelProps[tableColumns[i].property] = tableColumns[i].property
    }
  } else {
    for (let i = 0; i < tableColumns.length!; i++) {
      if (tableColumns[i].type !== 'default') continue
      if (tableColumns[i].property === undefined) continue
      ExcelHeader.push([tableColumns[i].label])
      ExcelProps[tableColumns[i].property] = tableColumns[i].property
    }
  }
  const lstData: any = []
  const reg = /^\d+$/
  for (const datum of tableData) {
    const rowData: Array<any> = []
    for (const dataKey in ExcelProps) {
      if (reg.test(datum[dataKey]) && datum[dataKey].toString().length === 13) {
        rowData.push(dayjs(new Date(datum[dataKey])).format('YYYY-MM-DD H:m:s'))
      } else if (datum[dataKey] === null || datum[dataKey] === undefined) {
        rowData.push('')
      } else {
        rowData.push(datum[dataKey])
      }
    }
    lstData.push(rowData)
  }
  const lstSheet: Array<any> = []
  ExcelUtil.writeHeader(ExcelHeader)
  for (
    let i = 0, rowIndex = myColumns.length;
    i < lstData.length;
    i++, rowIndex++
  ) {
    for (let j = 0; j < lstData[i].length; j++) {
      ExcelUtil.writeCellData(rowIndex, j, lstData[i][j])
    }
  }
  const sheet = ExcelUtil.getDataSheet()
  lstSheet.push({ sheet, name: `Sheet` })
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
