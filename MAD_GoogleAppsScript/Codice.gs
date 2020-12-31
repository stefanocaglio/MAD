function sendMAD() {
  
  Logger.log("Start sending email(s)...");
  
  // @OnlyCurrentDoc
  var sheet = SpreadsheetApp.getActiveSheet(); // Current sheet
  var ss = SpreadsheetApp.getActive(); // Current file
  
// FOLDER DETECTION
  // Get current folder id from current file
  var directParents = DriveApp.getFileById(ss.getId()).getParents();
  
  while( directParents.hasNext() ) {
    var folder = directParents.next();
    var folderId = folder.getId();
//    Logger.log(folder.getName() + " has id " + folderId);
  }  
  // Get parent folder
  var parent = DriveApp.getFolderById(folderId); // folder Id
  
// LOOP OVER DATA ROWS
  var startRow = 2;      // Avoid header row
  var endRow = sheet.getLastRow();
  
  var mailNumber = 0;
  var mailSkip = 0;
 
  
  // Sheets data structure:
  //    1          2        3            4              5         6       7      8           9
  // Provincia	Codice	Tipologia	Denominazione	Indirizzo	Civico	Comune	CAP	  Indirizzo PEC Autonomia
  
  for (var i = startRow; i <= endRow; i++) {
  
    // Skip filtered rows, if any
    if (sheet.isRowHiddenByFilter(i)) {
//      Logger.log('Row #' + i + ' is filtered');
      mailSkip = mailSkip + 1;
      continue;
    }
    
    var provincia = sheet.getRange(i,1).getValue();
    var codice = sheet.getRange(i,2).getValue();
    var tipologia = sheet.getRange(i,3).getValue();
    var denominazione = sheet.getRange(i,4).getValue();
    var indirizzo = sheet.getRange(i,5).getValue();
    var civico = sheet.getRange(i,6).getValue();
    var comune = sheet.getRange(i,7).getValue();
    var cap = sheet.getRange(i,8).getValue();
    var mailPec = sheet.getRange(i,9).getValue();

    // SEARCH ATTACHMENT
    var filenameSearch = sheet.getRange(i,10).getValue();   // TESTING
//    var filenameSearch = codice;   // to be active
    var search = "title contains '"+ filenameSearch +"'"
    
    var files = parent.searchFiles(search);
    while (files.hasNext()) {
      var file = files.next();
    }
    
    // MAIL PREPARATION
    var subject = 'Invio Messa a Disposizione a '+ denominazione + ' di ' + comune + ' (' + provincia + ') - ' + codice;
    
    var body = 'Gentile Dirigente Scolastico,\n\nmi auguro di trovare bene Lei e tutti i Suoi collaboratori anche se in una situazione storica che, fortunatamente, viviamo per la prima volta.\n\nLe scrivo per sottoporLe la mia messa a disposizione, che trova in allegato a questa mail.\n\nNel caso La considerasse in linea con i Vostri requisiti, trova tutti i miei riferimenti sia sul modulo che in calce a questa email.\n\nCon i più cordiali saluti\n\n\n     Stefano Caglio\n\nTelefono: +39 333 3757003  / / /  +39 345 0457911\nMail: stefano.caglio@gmail.com  / / /  info@pec.stefanocaglio.com\nsito stefanocaglio.com';
    var htmlBody = HtmlService.createHtmlOutputFromFile('mail_template').getContent();
    
    GmailApp.sendEmail(mailPec, subject, body, {
      attachments: [file.getAs(MimeType.PDF)],
      name: 'Francesco Taliento',
      htmlBody: htmlBody
    });
    mailNumber = mailNumber + 1;
    Logger.log(mailPec, codice, file);
  }
  
  Logger.log(mailNumber + " email(s) sent!  -  " + mailSkip + " row(s) skipped");
  Browser.msgBox(mailNumber + " email(s) sent!  -  " + mailSkip + " row(s) skipped");
    
/**  if (mailNumber ==1) {
    Logger.log('È stata inviata ' + mailNumber + ' email!');
    Browser.msgBox('È stata inviata ' + mailNumber + ' email!');
    }
  else {
    Logger.log('Sono state inviate ' + mailNumber + ' email!');
    Browser.msgBox('Sono state inviate ' + mailNumber + ' email!');
    }
    **/
}