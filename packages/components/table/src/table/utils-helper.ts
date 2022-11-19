// @ts-nocheck
import { exportExcel, importExcel } from '../store/exportHandler'
import type { Store } from '../store'
function useUtils<T>(store: Store<T>) {
  const setCurrentRow = (row: T) => {
    store.commit('setCurrentRow', row)
  }
  const getSelectionRows = () => {
    return store.getSelectionRows()
  }
  const toggleRowSelection = (row: T, selected: boolean) => {
    store.toggleRowSelection(row, selected, false)
    store.updateAllSelected()
  }
  const clearSelection = () => {
    store.clearSelection()
  }
  const clearFilter = (columnKeys: string[]) => {
    store.clearFilter(columnKeys)
  }
  const toggleAllSelection = () => {
    store.commit('toggleAllSelection')
  }
  const toggleRowExpansion = (row: T, expanded?: boolean) => {
    store.toggleRowExpansionAdapter(row, expanded)
  }
  const clearSort = () => {
    store.clearSort()
  }
  const sort = (prop: string, order: string) => {
    store.commit('sort', { prop, order })
  }
  const QueryLastOnlyIdExport = (
    fileName: string,
    type: string,
    tableData: Array<T>,
    tableColumns: Array<T>
  ) => {
    exportExcel(fileName, type, tableData, tableColumns)
  }
  const importData = (data: any, columns: any) => {
    return importExcel(data, columns)
  }
  return {
    setCurrentRow,
    getSelectionRows,
    toggleRowSelection,
    clearSelection,
    clearFilter,
    toggleAllSelection,
    toggleRowExpansion,
    clearSort,
    sort,
    QueryLastOnlyIdExport,
    importData,
  }
}

export default useUtils
