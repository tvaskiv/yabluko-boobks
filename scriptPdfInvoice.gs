var invoiceTemplate = "13RBC_s3g6wA13q18Xm2zTKiugGDLn15dVqf66IH7XBw";
var invoiceName = "Invoice";

function onFormSubmit(e) {
  
  var languege = e.values[1]; // Choose the language you prefer for mailing
  var promoCode = e.values[2]; // Enter personal key if you have one = pesonal code for discount
  
  Logger.log(promoCode);
  // ordered books
  var elemetaryStdtsBk = e.values[3];
  var elemetaryWrktsBk = e.values[4];
  var advencedStdsBk = e.values[5];
  var advencedWrkBk = e.values[6];
  var advencedActvtBk = e.values[7];
  var intermidiateBkGetInformed = e.values[8]; // yes or not if person would like to be informed about new edition
  var firstName = e.values[9];
  var lastName = e.values[10];
  var emailAddrss = e.values[11];
  var phoneNumber = e.values[12];
  var globalLocation = e.values[13];
  var mailingAddrssSendTo = e.values[14];
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSheet = ss.getSheetByName("Form Responses 2"); 
  var newInvoiceNumer = actSheet.getLastRow()+700; // counter creates new invoice number
  
  
  // next block adds persons info to the list to be informed about intermidiate level book
  if(intermidiateBkGetInformed == "Yes"){
    var listToInform = ss.getSheetByName("IntermdListToInform");
    var listLastRow = listToInform.getLastRow()+1;
    listToInform.getRange(listLastRow, 2).setValue(emailAddrss);
    listToInform.getRange(listLastRow, 3).setValue(firstName);
    listToInform.getRange(listLastRow, 4).setValue(lastName);
    
  } else {
    Logger.log("Problem");
  }
  
  // next block mekes a copy of invoicefile and adds name and numbers of books
  var copyId = DriveApp.getFileById(invoiceTemplate)
  .makeCopy(invoiceName + " #" + newInvoiceNumer +" for "+ firstName + " " + lastName)
  .getId();
//  Logger.log(copyId);
  
  // 
  var ssInvoiceSheet = SpreadsheetApp.openById(copyId);
  var ssInvoice = ssInvoiceSheet.getSheetByName("Invoice");
  ssInvoice.getRange(12, 1).setValue(firstName + " " + lastName);
  
  ssInvoice.getRange(16, 5).setValue(elemetaryStdtsBk);
  ssInvoice.getRange(17, 5).setValue(elemetaryWrktsBk);
  ssInvoice.getRange(22, 5).setValue(advencedStdsBk);
  ssInvoice.getRange(23, 5).setValue(advencedWrkBk);
  ssInvoice.getRange(24, 5).setValue(advencedActvtBk);
  
  // set new invoice number
  ssInvoice.getRange(6, 6).setValue(newInvoiceNumer);
  
  // adds an approximate shipping cost $20
  if (globalLocation == "Україна / Ukraine"){
    ssInvoice.getRange(31, 6).setValue("");
  } else {
    ssInvoice.getRange(31, 6).setValue(20);
  }
  // applying discount
  if (promoCode != 0) {
    var output = findPromoCodeMatch(promoCode);
    ssInvoice.getRange(28, 6).setValue(output);
  } else if (promoCode == 0){
    var output = "";
    Logger.log("There is no promo code");
  } else {
    var output = "";
    Logger.log("There is no promo code");
  }
  
  var htmlEng = "<body>"+
  "Dear "+ firstName + " " + lastName + "," + "<br />" +
  "Thank you for your order. You can find your invoice in attachment." + "<br /><br />" +
  "Please check if all information is correct." + "<br /><br />" +
  "Name: " + firstName + " " + lastName + "<br />" +
  "Delivery address: " + mailingAddrssSendTo + "<br /><br />" +

  "Let us know after you make payment. It will help us to check payment and send your books faster." + "<br /><br />" +

  "Here is a little instruction how to make payment:" + "<br />" +
  "You can pay Online following this link: http://studyukrainian.org.ua/en/books/cost_and_pay" + "<br />" +
  "or this link: http://supporting.ucu.edu.ua/en/support/donate/" + "<br />" +
  "Here you can find a step by step example of the procedure how to donate for the books: http://studyukrainian.org.ua/php_uploads/data/articles/ArticleFiles_35_How_to_pay_on-line.pdf" + "<br /><br />" +
  "It is recommended to make payment through Windows system via Google Chrome browser." + "<br /><br />" +
   
  "If you have any questions about books, or you want to change order information, send us an email." + "<br /><br />" +

  "Kind regards," + "<br />" +
  "Taras Vaskiv" + "<br />" +
  "+38 067 697 18 69" + "<br />" +
  "yablukobooks@ucu.edu.ua" + "<br />" +

  "</body>";
  
  var htmlUkr = "<body>"+
  "Доброго дня "+ firstName + " " + lastName + "," + "<br />" +
  "Дякуємо за Ваше замовлення. У додатку надсилаю Вам рахунок." + "<br /><br />" +

  "Найзручнішим способом отримання книг в межах України є сервіс “Нова Пошта”. Здійснити оплату можна безпосередньо у відділенні НП при отримані замовлення." + "<br /><br />" +

  "Прошу перевірити вказану Вами інформацію. У відповіді на цього листа, підтвердіть Вашу згоду на виконання замовлення." + "<br />" +
  "Ви вказали адресу доставки: " + mailingAddrssSendTo + "<br />" +
  "Ваш номер телефону: " + phoneNumber + "<br />" +
  "Ваше ім’я та прізвище: " + firstName + " " + lastName + "<br />" +
   
  "У разі якщо у Вас виникли додаткові запитання, або Ви хочете змінити інформацію стосовно замовлення прошу написати про це." + "<br /><br />" +

  "З повагою," + "<br />" +
  "Тарас Васьків" + "<br />" +
  "+38 067 697 18 69" + "<br />" +
  "yablukobooks@ucu.edu.ua" + "<br />" +

  "</body>";
  
  if (languege == "English"){
    var html = htmlEng;
  } else if (languege == "Українська"){
    var html = htmlUkr;
  }
  
  Logger.log(ssInvoice.getRange(16, 5).getValue());
  MailApp.sendEmail(emailAddrss, "Invoice for YABLUKO books","Requires html", {
    name: "YABLUKO books",
    htmlBody: html,
    attachments: [ssInvoiceSheet.getAs(MimeType.PDF)]
  });

}



function findPromoCodeMatch(promoCode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PartnersDiscounts");
  var lastRow = ss.getLastRow();
  var lookupRangeValues = ss.getRange(2, 2, lastRow).getValues();
  for (var i = 0; i < lookupRangeValues.length-1; i++){
    
    if (lookupRangeValues[i][0] == promoCode) {
      var codeIndex = i+2;
      var partnersName = ss.getRange(codeIndex, 1).getValue();
      var discountRate = ss.getRange(codeIndex, 3).getValue();
      break;
    } else {
      
    }
  }
  return discountRate;
  
}
























