//Google Apps Script

//Project OAuth Scopes, 2 Scopes Requested:
//See, edit, create, and delete all of your Google Drive files	https://www.googleapis.com/auth/drive
//View and manage your Google Docs documents	https://www.googleapis.com/auth/documents

//Based on the following tutorial which can better explain this use case, https://jeffreyeverhart.com/2018/09/17/auto-fill-google-doc-from-google-form-submission/

function autoFillGoogleDocFromForm(e) {
  //e.values is an array of form values
  var timestamp = e.values[0];
  var date = e.values[1];
  var name = e.values[2];
  var policyNumber = e.values[3];
  var veh1Year = e.values[4];
  var veh1Make = e.values[5];
  var veh1PD = e.values[6];
  var veh1Towing = e.values[7];
  var veh1Rental = e.values[8];

  var veh2Year = e.values[9];
  var veh2Make = e.values[10];
  var veh2PD = e.values[11];
  var veh2Towing = e.values[12];
  var veh2Rental = e.values[13];
  var veh3Year = e.values[14];
  var veh3Make = e.values[15];
  var veh3PD = e.values[16];
  var veh3Towing = e.values[17];
  var veh3Rental = e.values[18];
  
  var veh4Year = e.values[19];
  var veh4Make = e.values[20];
  var veh4PD = e.values[21];
  var veh4Towing = e.values[22];
  var veh4Rental = e.values[23];
  
  var medPay = e.values[24];
  var rideShare = e.values[25];
  var aDnd = e.values[26];
  var bodilyInjury = e.values[27];
  var propertyDamage = e.values[28];
  var underInsured = e.values[29];
  var underInsuredPD = e.values[30];
  
  //var templateFile = DriveApp.getFileById("1X3m5SlYUxhL78PE520W1qG7aUuLUEN0kecuR9OcCxPc"); //merged file that is not formatted properly, new plan: spit out two seperate documents
  var templateFile = DriveApp.getFileById("137B0rhUsCPcQRmSseM43eBUIawkarG2M7Klx42mrl2Y"); //coverage rejection form
  //the blank form to be filled by the script
  var templateResponseFolder = DriveApp.getFolderById("1NmqXBcKqdu_63cD02tVbEhRTNceTR1Dg");
  //the destination folder for the filled forms
  
  var copy = templateFile.makeCopy(name + 'coverage rejection form', templateResponseFolder);
  //creates a copy of the blank form templateFile and gives it a name
  
  var doc = DocumentApp.openById(copy.getId());
  //opens the document
  
  var body = doc.getBody();
  //gets the body of the document
  
  body.replaceText("{{Today's Date}}", date);
  body.replaceText("{{Insured's Name}}", name);
  body.replaceText("{{Policy Number}}", policyNumber);
  body.replaceText("{{Veh 1 Year}}", veh1Year);
  body.replaceText("{{Veh 1 Make}}", veh1Make);
  body.replaceText("{{Veh 1 Comp/Collision}}", veh1PD);
  body.replaceText("{{Veh 1 Towing}}", veh1Towing);
  body.replaceText("{{Veh 1 Rental}}", veh1Rental);
  body.replaceText("{{Medical Payments}}", medPay);
  body.replaceText("{{Rideshare}}", rideShare);
  body.replaceText("{{AD&D}}", aDnd);
  body.replaceText("{{Combined Un & Underinsured}}", underInsured);
  body.replaceText("{{Bodily Injury}}", bodilyInjury);
  body.replaceText("{{Property Damage}}", propertyDamage);
  body.replaceText("{{Underinsured PD}}", underInsuredPD);
  body.replaceText("{{Veh 2 Year}}", veh2Year);
  body.replaceText("{{Veh 2 Make}}", veh2Make);
  body.replaceText("{{Veh 2 Comp/Collision}}", veh2PD);
  body.replaceText("{{Veh 2 Towing}}", veh2Towing);
  body.replaceText("{{Veh 2 Rental}}", veh2Rental);
  body.replaceText("{{Veh 3 Year}}", veh3Year);
  body.replaceText("{{Veh 3 Make}}", veh3Make);
  body.replaceText("{{Veh 3 Comp/Collision}}", veh3PD);
  body.replaceText("{{Veh 3 Towing}}", veh3Towing);
  body.replaceText("{{Veh 3 Rental}}", veh3Rental);
  body.replaceText("{{Veh 4 Year}}", veh4Year);
  body.replaceText("{{Veh 4 Make}}", veh4Make);
  body.replaceText("{{Veh 4 Comp/Collision}}", veh4PD);
  body.replaceText("{{Veh 4 Towing}}", veh4Towing);
  body.replaceText("{{Veh 4 Rental}}", veh4Rental);
  //calls to all of our replaceText methods
  
  doc.saveAndClose();
  //saves and closes the newly created document
  
  var templateFile2 = DriveApp.getFileById("1X3m5SlYUxhL78PE520W1qG7aUuLUEN0kecuR9OcCxPc"); //policy fee form
  //the blank form to be filled by the script
  var templateResponseFolder = DriveApp.getFolderById("1NmqXBcKqdu_63cD02tVbEhRTNceTR1Dg");
  //the destination folder for the filled forms
  
  var copy2 = templateFile2.makeCopy(name + 'policy fee form', templateResponseFolder);
  //creates a copy of the blank form templateFile and gives it a name
  
  var doc2 = DocumentApp.openById(copy2.getId());
  //opens the document
  
  var body2 = doc2.getBody();
  //gets the body of the document
  
  body2.replaceText("{{Today's Date}}", date);
  body2.replaceText("{{Insured's Name}}", name);
  body2.replaceText("{{Policy Number}}", policyNumber);
  body2.replaceText("{{Veh 1 Year}}", veh1Year);
  body2.replaceText("{{Veh 1 Make}}", veh1Make);
  body2.replaceText("{{Veh 1 Comp/Collision}}", veh1PD);
  body2.replaceText("{{Veh 1 Towing}}", veh1Towing);
  body2.replaceText("{{Veh 1 Rental}}", veh1Rental);
  body2.replaceText("{{Medical Payments}}", medPay);
  body2.replaceText("{{Rideshare}}", rideShare);
  body2.replaceText("{{AD&D}}", aDnd);
  body2.replaceText("{{Combined Un & Underinsured}}", underInsured);
  body2.replaceText("{{Bodily Injury}}", bodilyInjury);
  body2.replaceText("{{Property Damage}}", propertyDamage);
  body2.replaceText("{{Underinsured PD}}", underInsuredPD);
  body2.replaceText("{{Veh 2 Year}}", veh2Year);
  body2.replaceText("{{Veh 2 Make}}", veh2Make);
  body2.replaceText("{{Veh 2 Comp/Collision}}", veh2PD);
  body2.replaceText("{{Veh 2 Towing}}", veh2Towing);
  body2.replaceText("{{Veh 2 Rental}}", veh2Rental);
  body2.replaceText("{{Veh 3 Year}}", veh3Year);
  body2.replaceText("{{Veh 3 Make}}", veh3Make);
  body2.replaceText("{{Veh 3 Comp/Collision}}", veh3PD);
  body2.replaceText("{{Veh 3 Towing}}", veh3Towing);
  body2.replaceText("{{Veh 3 Rental}}", veh3Rental);
  body2.replaceText("{{Veh 4 Year}}", veh4Year);
  body2.replaceText("{{Veh 4 Make}}", veh4Make);
  body2.replaceText("{{Veh 4 Comp/Collision}}", veh4PD);
  body2.replaceText("{{Veh 4 Towing}}", veh4Towing);
  body2.replaceText("{{Veh 4 Rental}}", veh4Rental);
  //calls to all of our replaceText methods
  
  doc2.saveAndClose();
  //saves and closes the newly created document  
}
//Todo:
//- Change getFileById to new updated document with tags added (for other template options)
//- Create variants for other lines of business and document types (Business Auto, Homeowners, General Liability, Workers Comp, Prom Note, Designation of Authorized Rep, Ride-sharing, etc)
