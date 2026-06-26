function doPost(e){
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data=JSON.parse(e.postData.contents);
  var headers=["Timestamp","Full Name","Job Title","Organization","Email","Phone","Website","Organization Type","Organization Size","Services","Inspiration","Challenges","One Challenge","Event Type","Preferred Date","Location","Audience Size","Format","Budget","Source","Final Thoughts"];
  if(sheet.getLastRow()===0)sheet.appendRow(headers);
  sheet.appendRow([new Date(),data.fullName||"",data.jobTitle||"",data.organization||"",data.email||"",data.phone||"",data.website||"",data.organizationType||"",data.organizationSize||"",Array.isArray(data.services)?data.services.join(", "):(data.services||""),data.inspiration||"",Array.isArray(data.challenges)?data.challenges.join(", "):(data.challenges||""),data.oneChallenge||"",data.eventType||"",data.preferredDate||"",data.location||"",data.audienceSize||"",data.format||"",data.budget||"",data.source||"",data.finalThoughts||""]);
  MailApp.sendEmail({to:"darylcrosby78@gmail.com",subject:"New DJ Mindset Booking Inquiry",htmlBody:"A new inquiry was submitted by <strong>"+(data.fullName||"Unknown")+"</strong> from <strong>"+(data.organization||"Unknown Organization")+"</strong>.<br><br>Check your Google Sheet for details."});
  return ContentService.createTextOutput(JSON.stringify({result:"success"})).setMimeType(ContentService.MimeType.JSON);
}