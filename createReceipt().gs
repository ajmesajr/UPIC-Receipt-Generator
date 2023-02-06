function createReceipt() // based on https://codewithcurt.com/how-to-create-invoice-generator-on-google-sheets/
{
  // DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // DEFINE MENU SHEET          
  var invoiceSheet = ss.getSheetByName("Order Summary");
  var printSheet = ss.getSheetByName("Print");
  var generatorSheet = ss.getSheetByName("Receipt Generator");

  // GET ENTRIES FROM GENERATOR
  var order_number = generatorSheet.getRange('A2').getValue();
  var customer_name = generatorSheet.getRange('B2').getValue();
  var date_of_order = generatorSheet.getRange('C2').getValue();
  var payment_method = generatorSheet.getRange('D2').getValue();
  var quantity_ordered = generatorSheet.getRange('E2').getValue();
  var xs_count = generatorSheet.getRange('F2').getValue(); 
  var s_count = generatorSheet.getRange('G2').getValue(); 
  var m_count = generatorSheet.getRange('H2').getValue();
  var l_count = generatorSheet.getRange('I2').getValue();
  var xl_count = generatorSheet.getRange('J2').getValue();
  var customer_email = generatorSheet.getRange('K2').getValue();
    
  // SET VALUES ON INVOICE
  printSheet.getRange('B12').setValue(customer_name).setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("#545454");
  printSheet.getRange('B13').setValue(customer_email).setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("#545454");
  printSheet.getRange('B15').setValue(payment_method).setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("#545454");
  printSheet.getRange('F12').setValue(order_number).setNumberFormat("0000").setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("#545454");
  printSheet.getRange('F15').setValue(date_of_order).setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("545454");

  printSheet.getRange('B19').setValue("YTS 2023 Shirt").setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("545454");
  
  var shirt_price = 350;
  var subTotal = quantity_ordered*shirt_price;

  printSheet.getRange('E19').setValue(quantity_ordered).setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("545454");
  printSheet.getRange('F19').setValue(shirt_price).setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("545454");  
  printSheet.getRange('G19').setValue(subTotal).setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("545454");  
  
  // SET TOTAL
  printSheet.getRange('G25').setValue(subTotal).setNumberFormat("# ###.00").setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("#545454");
  
  var adjustment = 0;
  if (quantity_ordered >= 10)
    adjustment = -100;  

  var totalInvoice = subTotal + adjustment;
  var note = '     XS:' + xs_count + '; S:' + s_count + '; M:' + m_count + '; L:' + l_count + '; XL:' + xl_count;

  printSheet.getRange('C19').setValue(note).setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("#545454");
  printSheet.getRange('G26').setValue(adjustment).setNumberFormat("# ###.00").setFontFamily('Helvetica Neue').setFontSize(10).setFontColor("#545454");

  // CALL INVOICE LOG
  InvoiceLog(order_number, customer_name, date_of_order, totalInvoice)

  printReceipt(); // save the receipt to PDF and email to the customer

  clearReceipt(); // clears all added entries
  return true;
}

function InvoiceLog(order_number, customer_name, date_of_order, totalInvoice) // based on https://codewithcurt.com/how-to-create-invoice-generator-on-google-sheets/
{
  
   //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE INVOICE LOG SHEET          
  var invoiceLogSheet = ss.getSheetByName("Receipt Log"); 
  
  //GET LAST ROW OF INVOICE LOG SHEET
  var nextRowInvoice = invoiceLogSheet.getLastRow() + 1;
  
  //POPULATE INVOICE LOG
  invoiceLogSheet.getRange(nextRowInvoice, 1).setValue(order_number);
  invoiceLogSheet.getRange(nextRowInvoice, 2).setValue(customer_name);
  invoiceLogSheet.getRange(nextRowInvoice, 3).setValue(date_of_order);
  invoiceLogSheet.getRange(nextRowInvoice, 4).setValue(totalInvoice).setNumberFormat("# ###.00");
}

function printReceipt() // based on https://stackoverflow.com/questions/61219628/export-a-google-sheet-range-as-pdf-using-apps-script-and-store-the-pdf-in-drive
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var printSheet = ss.getSheetByName("Print");    
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for(var i =0;i<sheets.length;i++){
    if(sheets[i].getName()!="Print"){ sheets[i].hideSheet() }
  }

  var customer_name = printSheet.getRange("B12").getValue();
  var customer_email = printSheet.getRange("B13").getValue();
  var order_number = printSheet.getRange("F12").getValue();
  
  var pdf = DriveApp.getFileById(ss.getId());

  var file_name = "Order#: " + Utilities.formatString( "%04d",order_number ) + " Customer: " + customer_name;
  var theBlob = pdf.getBlob().getAs('application/pdf').setName(file_name + ".pdf");
  var folderID = "1up5e8jGeDnzkNYAy3jqJrNjR9TPk67ds"; // Folder id to save in a folder
  var folder = DriveApp.getFolderById(folderID);
  var newFile = folder.createFile(theBlob);
  // var body = 'Dear ' + customer_name +',\n\n Please see the attached file containing the official receipt for your order of the YTS 2023: Revv Up Mechandise. We thank you for your support and we hope to see you again in our future events and merchandise.'+'\n\n Regards, \n UP Investment Club \n The Trade that Pays Off';
  // GmailApp.sendEmail(customer_email, "YTS 2023: Revv Up Merchandise Official Receipt", body, {attachments: [theBlob]});
  
  for(var i =0;i<sheets.length;i++){
    if(sheets[i].getName()!="Print"){ sheets[i].showSheet() }
  }  
}

function clearReceipt()
{
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE INVOICE SHEET          
  var invoiceSheet = ss.getSheetByName("Print");
  
  
  //SET VALUES TO NOTHING
  invoiceSheet.getRange('B12').setValue("");
  invoiceSheet.getRange('B13').setValue("");
  invoiceSheet.getRange('B15').setValue("");
  invoiceSheet.getRange('F12').setValue("");  
  invoiceSheet.getRange('F15').setValue(""); 

  invoiceSheet.getRange('B19').setValue("");  

  invoiceSheet.getRange('E19').setValue("");  
  invoiceSheet.getRange('F19').setValue("");  
  invoiceSheet.getRange('G19').setValue("");  

  invoiceSheet.getRange('G25').setValue(""); 

  invoiceSheet.getRange('C19').setValue(""); 
  invoiceSheet.getRange('G26').setValue(""); 

}
