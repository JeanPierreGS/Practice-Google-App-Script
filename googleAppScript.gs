function createMenu (){
  const menu = SpreadsheetApp.getUi().createMenu("Correspondencia");
  menu.addItem("Combinar Correspondencia", 'myFunction').addSeparator().addToUi();
  menu.addItem("Crear PDF", 'createPDF').addToUi();
}

function myFunction () {
  const sps = SpreadsheetApp.getActive()
  const sheet = sps.getSheetByName("email")
  const data  = sheet.getDataRange().getValues()
  let status = sheet.getRange(1, 8)
  status.setBackground("#1274D1")
  status.setValue("Status")
  data.forEach((item, index) => {
    if(index != 0) {
      const email = item[6]
      const name = item[2]
      const lastname = item[3]
      const cod = item[1]
      const calification = item[4]
      const observation = item[5]
      const no = item[0]
      const res = sendEmail(email, name, lastname, cod, calification, observation, no)
      if (res) {
         let statusProcces = sheet.getRange(index + 1, 8)
         statusProcces.setBackground("#3DD112")
         statusProcces.setValue("Enviado")
      } else {
        let statusProcces = sheet.getRange(index + 1, item.length)
        statusProcces.setBackground("#CD261E")
        statusProcces.setValue("Error")
      }
    }
  })
  generateDoc()
  createPDF()
}

function sendEmail (email, names, lastnames, cod, calification, observation, no) {
  const body = `Estimado (a) Estudiante ${names} ${lastnames} Código: ${cod} \n
  Nos permitimos informarle su calificación definitiva en el curso Sistemas de Información Gerencial:\n
  Calificación: ${calification}\n 
  Observaciones: ${observation}\n
  Puesto en el Grupo: ${no}\n
  Felicitaciones y éxitos en su vida académica y profesional.`
  try {
    GmailApp.sendEmail(email, "Google App Script", body); 
    return true
  } catch (err) {
    return false
  }
}

function generateDoc () {
  const docID = "1ZoBEuJdVSy9ycOdUhMx_ebm32TOyxRzQ6op1gsDpMAM"
  const sps = SpreadsheetApp.getActive()
  const sheet = sps.getSheetByName("email")
  const data  = sheet.getDataRange().getValues()
  let doc = DocumentApp.openById(docID)
  doc.getBody().clear()
  doc.getBody().appendTable(data)
}

function createPDF () {
  let document = DriveApp.getFileById("1ZoBEuJdVSy9ycOdUhMx_ebm32TOyxRzQ6op1gsDpMAM")
  let folder = DriveApp.getFolderById("1UeRnPHxFIF22qcC_HpIp6oVvc9PANzKC")
  const documentCopy = document.makeCopy(folder)
  const filePDF = documentCopy.getAs(MimeType.PDF)
  folder.createFile(filePDF).setName("1151593")
  SpreadsheetApp.getUi().alert("El PDF se ha creado correctamente")
}