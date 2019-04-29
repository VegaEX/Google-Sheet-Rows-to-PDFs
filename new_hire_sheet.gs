var TEMPLATE_ID = '--GDOC Template ID--'

var RESULTS_FOLDER_ID = 'New Hire PDFs'

var USERNAME_COLUMN_NAME = 'Username'
var EMAIL_COLUMN_NAME = 'Email'
var FIRST_NAME_COLUMN_NAME = 'First Name'
var LAST_NAME_COLUMN_NAME = 'Last Name'
var TEMP_PASSWORD_COLUMN_NAME = 'Day One Password'
var DUO_TOKEN_COLUMN_NAME = 'Duo Token'
var DATE_FORMAT = 'yyyy/MM/dd';
var PDF_FILE_NAME = 'Day One Sheet for' + USERNAME_COLUMN_NAME

var EMAIL_SUBJECT = 'New Hire onboarding information for' + USERNAME_COLUMN_NAME
var EMAIL_BODY = 'Attached is a PDF with your new hires day one information. If something looks out of place or you have additional questions please reach out to your local IT team for assistance.'

/**
 * Eventhandler
 */

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu('[ Create PDFs ]')
    .addItem('Create a PDF for each row', 'createPdfs')
    .addToUi()

}


function createPdfs() {

  var ui = SpreadsheetApp.getUi()

  if (TEMPLATE_ID === '') {
    ui.alert('TEMPLATE_ID needs to be defined in code.gs')
    return
  }

  // Set up the docs and the spreadsheet access

  var templateFile = DriveApp.getFileById(TEMPLATE_ID)
  var activeSheet = SpreadsheetApp.getActiveSheet()
  var allRows = activeSheet.getDataRange().getValues()
  var headerRow = allRows.shift()

  // Create a PDF for each row

  allRows.forEach(function(row) {

    createPdf(templateFile, headerRow, row)

    /**
     * Create a PDF
     *
     * @param {File} templateFile
     * @param {Array} headerRow
     * @param {Array} activeRow
     */

    function createPdf(templateFile, headerRow, activeRow) {

      var headerValue
      var activeCell
      var ID = null
      var recipient = null
      var copyFile
      var numberOfColumns = headerRow.length
      var copyFile = templateFile.makeCopy()
      var copyId = copyFile.getId()
      var copyDoc = DocumentApp.openById(copyId)
      var copyBody = copyDoc.getActiveSection()


      for (var columnIndex = 0; columnIndex < numberOfColumns; columnIndex++) {

        headerValue = headerRow[columnIndex]
        activeCell = activeRow[columnIndex]
        activeCell = formatCell(activeCell);

        copyBody.replaceText('<<' + headerValue + '>>', activeCell)

        if (headerValue === FILE_NAME_COLUMN_NAME) {

          ID = activeCell

        } else if (headerValue === EMAIL_COLUMN_NAME) {

          recipient = activeCell
        }
      }

      // Create the PDF file

      copyDoc.saveAndClose()
      var newFile = DriveApp.createFile(copyFile.getAs('application/pdf'))
      copyFile.setTrashed(true)

      // Rename the new PDF file

      if (PDF_FILE_NAME !== '') {

        newFile.setName(PDF_FILE_NAME)

      } else if (ID !== null){

        newFile.setName(ID)
      }

      // Put the new PDF file into the results folder

      if (RESULTS_FOLDER_ID !== '') {

        DriveApp.getFolderById(RESULTS_FOLDER_ID).addFile(newFile)
        DriveApp.removeFile(newFile)
      }

      // Email the new PDF

      if (recipient !== null) {

        MailApp.sendEmail(
          recipient,
          EMAIL_SUBJECT,
          EMAIL_BODY,
          {attachments: [newFile]})
      }

    } // createPdfs.createPdf()

  })

  ui.alert('New PDF files created')

  return
  
  /**
  * Format the cell's value
  *
  * @param {Object} value
  *
  * @return {Object} value
  */

  function formatCell(value) {

    var newValue = value;

    if (newValue instanceof Date) {

      newValue = Utilities.formatDate(
        value,
        Session.getScriptTimeZone(),
        DATE_FORMAT);

    } else if (typeof value === 'number') {

      newValue = Math.round(value * 100) / 100
    }

    return newValue;

  } // createPdf.formatCell()

} // createPdfs()
