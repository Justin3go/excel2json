import VxeTable from 'vxe-table'
import 'vxe-table/lib/style.css'

export default defineNuxtPlugin((nuxtApp) => {
  nuxtApp.vueApp.use(VxeTable)
})
