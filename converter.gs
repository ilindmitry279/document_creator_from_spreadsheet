function formDocument() {  
  disaForm();
  var dateNow = new Date();
  var startDate1 = new Date();
  var endDate7= new Date();
  var startDate1 = Utilities.formatDate(new Date(startDate1.setDate(dateNow.getDate()+1)), "GMT+2", "dd.MM");
  var endDate7 = Utilities.formatDate(new Date(endDate7.setDate(dateNow.getDate()+7)), "GMT+2", "dd.MM");
  var yearNow= Utilities.formatDate(dateNow, "GMT+2", "yyyy");
  var dateDoc= Utilities.formatDate(dateNow, "GMT+2", "dd.MM.yyyy")
  var raportname = "Рапорт наявність " + startDate1 + " - " + endDate7 + " " + yearNow + " року";
  // ID файлу шаблона
  var templateID = "17hDB0-skHKrST8QD83kNHKZFF3OV-lqsM4M_lzOHnYk";
  // Копіювання і перейменування файлу шаблона
  var tmpfile = DriveApp.getFileById(templateID);
  tmpfile.makeCopy(raportname);
  var newreport = DriveApp.getFilesByName(raportname).next();
  var newReportID = newreport.getId();
  var newreportURL = newreport.getUrl();
  var body = DocumentApp.openById(newReportID).getBody();
  var weekdays = weekArrayGenerator();
  var dates = dateArrayGenerator();
  shapka (body, startDate1, endDate7, yearNow, dateDoc, weekdays, dates);
  zapovnennya(body);
  clearSheetAndFormResponses();
  DocumentApp.openById(newReportID).saveAndClose();
  myMail(newreportURL,raportname, newReportID);
}

function weekArrayGenerator() {
  var weekArray = [];
  for (var i=1; i<=7; i++) {
    var dateNow = new Date();
    var nextDate = new Date();
    var nextDate = Utilities.formatDate(new Date(nextDate.setDate(dateNow.getDate()+i)), "GMT+2", "EEEE");
    var nextDow = LanguageApp.translate(nextDate,"en","uk");
    var nextdow = String(nextDow).toLowerCase();
    weekArray.push(nextdow); 
  }
  return weekArray
}

function dateArrayGenerator() {
  var dateArray = [];
  for (var i=1; i<=7; i++) {
    var dateNow = new Date();
    var nextDate = new Date();
    var nextDate = Utilities.formatDate(new Date(nextDate.setDate(dateNow.getDate()+i)), "GMT+2", "dd.MM");
    dateArray.push(nextDate); 
  }
  return dateArray
}

function formDataPostGS(i) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Відповіді форми');
  var long = ss.getLastRow();
  var ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Довідка');
  var postGS = {
    npp:i,
    username:ss1.getRange(i,1).getValue(),
    vzpib:ss1.getRange(i,2).getValue()
  }
  postGS.dayS = ss1.getRange(1,4,1,7).getValues();
  for (let l=1; l<=long; l++) {
    if (postGS.username == ss.getRange(l,2).getValue()) {
      postGS.dayS = ss.getRange(l,3,1,7).getValues();
    }
  }

  return postGS;
}

function clearSheetAndFormResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = ss.getSheetByName('Відповіді форми');
  var numRowsForDel = list.getLastRow();
  if (numRowsForDel>1) {
    list.deleteRows(2,numRowsForDel);
  }
  var urlForm = ss.getFormUrl();
  var form = FormApp.openByUrl(urlForm);
  form.deleteAllResponses();
}

function myMail(newreportURL, raportname, id) {
  var file = DriveApp.getFileById(id).getBlob();
  MailApp.sendEmail({ 
      to: "ilin.dmytry@ukr.net,tulovik76@gmail.com,bvv1805@gmail.com",
      //to: "ilin.dmytry@ukr.net", 
      subject: "На наступний тиждень.", 
      htmlBody: 'Андрій Олександрович, доповідаю про наявність особового складу на наступний тиждень. <br>'
                + '<a href=' + newreportURL + '>' + raportname + '</a>.',              
      name: "Валерій Бондар",
      attachments: [file]
      });   
}

function disaForm() {
  var formURL = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  var form = FormApp.openByUrl(formURL);
  //var formID = form.getId();
  form.setAcceptingResponses(false);
}

function enaForm() {
  var formURL = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  var form = FormApp.openByUrl(formURL);
  //var formID = form.getId();
  form.setAcceptingResponses(true);
}

function shapka (body,startDate1, endDate7, yearNow, dateDoc, weekdays, dates) {
  body.replaceText('{startDate1}', startDate1);
  body.replaceText('{endDate7}', endDate7);
  body.replaceText('{yearNow}', yearNow);
  body.replaceText('{dateDoc}', dateDoc);
  for (var j=0; j<=6; j++) {
    body.replaceText('{w'+ j + '}', weekdays[j]);
    body.replaceText('{d'+ j + '}', dates[j]);
  }
}

function zapovnennya(body) {
  let table = body.findElement(DocumentApp.ElementType.TABLE).getElement().asTable();
  let templateRow = table.getRow(2);
  for( var i=1; i<=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Довідка').getLastRow(); i++) {
    let tableRow = table.appendTableRow(templateRow.copy());
    let postGS = formDataPostGS(i);
    tableRow.getCell(0).setText(postGS.npp);
    tableRow.getCell(1).setText(postGS.vzpib);
    tableRow.getCell(2).setText(postGS.dayS[0][0]);
    tableRow.getCell(3).setText(postGS.dayS[0][1]);
    tableRow.getCell(4).setText(postGS.dayS[0][2]);
    tableRow.getCell(5).setText(postGS.dayS[0][3]);
    tableRow.getCell(6).setText(postGS.dayS[0][4]);
    tableRow.getCell(7).setText(postGS.dayS[0][5]);
    tableRow.getCell(8).setText(postGS.dayS[0][6]);      
  }
  templateRow.removeFromParent();//Видалення шаблонного рядка.
}

function createTimeDrivenTriggers() {
  ScriptApp.newTrigger(showData)
  .timeBased()
  .onWeekDay(ScriptApp.WeekDay.THURSDAY)
  .atHour(15)
  .create;
}

function showDate() {
  var todayDate = new Date ();
  var todayFormatedDate = Utilities.formatDate(todayDate.getDate(), "GMT+3")
  Logger.log(todayFormatedDate);
}
