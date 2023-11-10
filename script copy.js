

function myFunction() {
  url = 'https://www.tge.pl/energia-elektryczna-otf?dateShow='+getDateOfYesterdayOrLastFriday()+'&dateAction=prev';
  var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  
  var page_str = response.getContentText();
  
  
  var start_table_class = '<table class="footable table table-hover" data-sorting="true" id="footable_kontrakty_terminowe_0">';
  var start_table_class_pos = page_str.search(start_table_class);
  
  var page_str_from_start_table_class =  page_str.slice(start_table_class_pos, page_str.length);

  var start_table = '<table';
  var start_table_pos = page_str_from_start_table_class.search(start_table);
  
  var end_table = '</table>';
  var end_table_pos = page_str_from_start_table_class.search(end_table);

  var table_xml = page_str_from_start_table_class.slice(start_table_pos, (end_table_pos - start_table_pos) );
  
  //Logger.log('Table : ', table_xml)
  //for (var x=94; x<115; x = x+1){
    //Logger.log("Char at ["+String(x)+"] : '"+ table_xml.charAt(x)+"' - code: " + table_xml.charCodeAt(x))
    
  //}
  

  
  table_xml = table_xml.replace(/[\r\n]+/g, '');
    
  //Logger.log('Table : ', table_xml)
  //for (var x=94; x<115; x = x+1){
    //Logger.log("Char at ["+String(x)+"] : '"+ table_xml.charAt(x)+"' - code: " + table_xml.charCodeAt(x))
    
  //}
  
 

  var rows = get_rows(table_xml)
  

  data = []
  var row_count = 0;
  for (row in rows){
    
    var elements = get_elements(rows[row])
    row_count = row_count + 1
    var col_count = elements.length
    //Logger.log('Row ['+row+'] - cols: '+col_count)
    data.push(elements)
  }
  
  //Logger.log(data)

  var ss;
  ss = create_excel(data,row_count,col_count);
  
  var ss_Id = ss.getId();
  var sscopyid = ss.copy('Copy of '+ss.getName()).getId();

  ss = SpreadsheetApp.openById(ss_Id)
  //getGoogleSpreadsheetAsExcel(ss);
  //sendWorksheetAsXLSX(ss);
  //sendSpreadsheetAsXLSX(ss);
  Utilities.sleep(800)
  sendSpreadSheetAsExcel_so(ss);
  
  DriveApp.getFileById(sscopyid).setTrashed(true);
  DriveApp.getFileById(ss_Id).setTrashed(true)
  
  return 0;
}

function getDateOfYesterdayOrLastFriday() {
  var today = new Date();
  var dayOfWeek = today.getDay();

  if (dayOfWeek == 1) { // 1 is Monday
    today.setDate(today.getDate() - 3); // Go back three days for the previous Friday
  } else {
    today.setDate(today.getDate() - 1); // Go back one day for yesterday
  }
  var formattedDate = Utilities.formatDate(today, 'GMT', 'dd-MM-yyyy');
  return formattedDate;
}

function testDateFunction() {
  var date = getDateOfYesterdayOrLastFriday();
  Logger.log(date);
}

function create_excel(data, rows, cols){
  
  xlCols = {1:"A",2:"B",3:"C",4:"D",5:"E",6:"F",7:"G",8:"H",9:"I",10:"J",11:"K",12:"L",13:"M",14:"N",15:"O"}

  Logger.log('Rows: '+rows+' Cols: '+cols+' xlCol: '+xlCols[cols])


  var ws = SpreadsheetApp.create('TGE ' +  getDateOfYesterdayOrLastFriday(), rows, cols)
  var rg_str = String('A1:'+xlCols[cols]+''+String(rows))
  Logger.log('RG string: '+rg_str)
  var rg = ws.getActiveSheet().getRange(rg_str)
  rg.setValues(data)
  //var sh = ws.getActiveSheet
  return ws
}


function get_elements(row){
  
  var elements = []
  var r = row
  var res = r.match(/>(.*?)</g)
  
  for (e in res){
    var ee = res[e].slice(1, res[e].length -1)
    if ( (ee.trim().length >= 1 | ee.trim() == '-' | ee ==' ' || ee == '0') & ee != 'Suma'){
      elements.push(ee)
    }
  }
  return elements
}


function get_rows(tab){
  var table = tab
  var rows;
  //Logger.log('[get_rows] Table: ' + table)
  rows = table.match(/<tr.*?\/tr>/g)
  rows.pop()
  
  return rows

  }
  



function test() {
  var my_string = '<tr> asdasdas> sadsa < > <sad> </tr> <tr> asdasdas> sadsa < > <sad> </tr>'
  var macz_res = my_string.match(/<tr>.*?<\/tr>/g)
  var rplcd = my_string.replace('a','ADA')
  while ( rplcd.includes('a') ){
    rplcd = rplcd.replace('a','ADA')
  }
  
  
  Logger.log(rplcd)  


}

function sendWorksheetAsXLSX(ss) {
  try{
  // Define your spreadsheet and sheet
  //var spreadsheet = SpreadsheetApp.openById('YOUR_SPREADSHEET_ID');
  var spreadsheet = ss;
  //var sheet = spreadsheet.getSheetByName('Sheet1'); // Replace with your sheet name

  // Create a temporary file in XLSX format
  var temporaryFile = DriveApp.getFileById(spreadsheet.getId());
  var temporaryBlob = temporaryFile.getBlob().getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  
  // Define the recipient's email address and email subject
  var recipientEmail = 'rafalpstroko@gmail.com'; // Replace with the recipient's email
  var emailSubject = 'TGE '+getDateOfYesterdayOrLastFriday+' BASE';
  var body = 'Best regards, SEC Automation solutions';
  // Send the email
  MailApp.sendEmail({
    to: recipientEmail,
    subject: emailSubject,
    body: body,
    attachments: [temporaryBlob]
  });
    Logger.log('Email sent successfully!');
  }
  catch (f){
    Logger.log('Error: '+String(f));
  }
}


function sendSpreadsheetAsXLSX(ss) {
  try{
    var spreadsheet = ss//SpreadsheetApp.openById('YOUR_SPREADSHEET_ID');
  //var sheet = spreadsheet.getSheetByName('Sheet1'); // Replace with your sheet name

  // Export the spreadsheet to XLSX format
  var fileId = ss.getId

  // Get the URL of the newly created XLSX file
  var file = DriveApp.getFileById(fileId);
  var fileUrl = file.getUrl();

  // Define the recipient's email address and email subject
 
 
  var recipientEmail = 'rafalpstroko@gmail.com'; // Replace with the recipient's email
  var emailSubject = 'TGE '+getDateOfYesterdayOrLastFriday+' BASE';
  var body = 'Best regards, SEC Automation solutions';
  // Send the email
  MailApp.sendEmail({
    to: recipientEmail,
    subject: emailSubject,
    body: body,
    attachments: [file.getAs(MimeType.MICROSOFT_EXCEL)]
  });
  Logger.log('Email sent successfully!');
  }
  catch (f){
    Logger.log('Error: '+String(f));
  }
}
function doGet(){
  myFunction();
  return 0;
};

function sendSpreadSheetAsExcel_so(ss){
  
  var sheet = ss.getActiveSheet()
  Utilities.sleep(1000)
  //var url = "https://docs.google.com/spreadsheets/export?id=" + ss.getId() + "&exportFormat=xlsx&gid=" + sheet.getSheetId();
  //var url = "https://www.googleapis.com/drive/v3/files/"+ ss.getId() + "/export"//+"&exportFormat=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  var url = "https://docs.google.com/spreadsheets/d/"+ss.getId()+"/export?format=xlsx"
  var params = {
    method: "get",
    headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
               // mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
               },
    muteHttpExceptions: false
  };

  var blob = UrlFetchApp.fetch(url, params).getBlob();
  blob.setName(ss.getName() + ".xlsx");
  
  var recipientEmail = 'barbara.sedzimir@sec.com.pl'; // Replace with the recipient's email
  //var recipientEmail = "rafalpstroko@gmail.com"
  //var recipientEmail = 'rafal.pstrokonski.biuro@gmail.com,rafalpstroko@gmail.com'
  var emailSubject = 'TGE '+getDateOfYesterdayOrLastFriday()+' BASE';
  var body = "Dzień dobry,\n\n"
  body = body + "w załączniku przesyłamy aktualne dane TGE - kontrakty bazowe.\n"
  body = body + "Jeżeli zauważone zostaną błędy lub problemy proszę o kontakt pod adres helpdesk.ot@sec.com.pl"
  body = body + "\n\n"
  body = body + 'Best regards, SEC Automation solutions - Rafał Pstrokoński';
  // Send the email
  
  GmailApp.sendEmail(recipientEmail,
    emailSubject,
    body,
    {attachments : [blob.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')],
    bcc : "rafal.pstrokonski@sec.com.pl"}
  );
  

  Logger.log("Blob: "+String(blob.getContentType()))
  Logger.log('Email sent successfully!');
 
}

