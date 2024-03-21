import VxeTable from 'vxe-table'
import 'vxe-table/lib/style.css'
import VXETablePluginExportXLSX from 'vxe-table-plugin-export-xlsx'
import ExcelJS from 'exceljs'

export default defineNuxtPlugin((nuxtApp) => {
  VxeTable.use(VXETablePluginExportXLSX, {
    ExcelJS
  })
  nuxtApp.vueApp.use(VxeTable)
})
