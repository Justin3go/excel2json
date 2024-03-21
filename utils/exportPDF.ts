import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'

export const htmlToPDF = async (htmlId: string, title = '报表', bgColor = '#fff') => {
  const pdfDom = document.getElementById(htmlId)
  if (!pdfDom) return
  const scrollContent = pdfDom.querySelector('table')
  if (!scrollContent) return
  scrollContent.style.padding = '0 10px !important'
  const A4Width = 595.28
  const A4Height = 841.89
  const canvas = await html2canvas(scrollContent as any, {
    scale: 4,
    useCORS: true,
    backgroundColor: bgColor,
    windowHeight: pdfDom.scrollHeight,
    windowWidth: pdfDom.scrollWidth
  })
  const pageHeight = (canvas.width / A4Width) * A4Height
  let leftHeight = canvas.height
  let position = 0
  const imgWidth = A4Width
  const imgHeight = (A4Width / canvas.width) * canvas.height

  const pageData = canvas.toDataURL('image/jpeg', 1.0)
  const PDF = new jsPDF('p', 'pt', 'a4')
  if (leftHeight < pageHeight) {
    PDF.addImage(pageData, 'JPEG', 0, 0, imgWidth, imgHeight)
  } else {
    while (leftHeight > 0) {
      PDF.addImage(pageData, 'JPEG', 0, position, imgWidth, imgHeight)
      leftHeight -= pageHeight
      position -= A4Height
      if (leftHeight > 0) PDF.addPage()
    }
  }

  PDF.save(title + '.pdf')
}
