//Google Apps Script

//Project OAuth Scopes, 2 Scopes Requested:
//See, edit, create, and delete all of your Google Drive files	https://www.googleapis.com/auth/drive
//View and manage your Google Docs documents	https://www.googleapis.com/auth/documents

//Based on the followng tutorial which can better explain this use case, https://jeffreyeverhart.com/2018/09/17/auto-fill-google-doc-from-google-form-submission/

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

  var veh2Year = e.values[15];
  var veh2Make = e.values[16];
  var veh2PD = e.values[17];
  var veh2Towing = e.values[18];
  var veh2Rental = e.values[19];
  var veh3Year = e.values[20];
  var veh3Make = e.values[21];
  var veh3PD = e.values[22];
  var veh3Towing = e.values[23];
  var veh3Rental = e.values[24];
  
  var veh4Year = e.values[20];
  var veh4Make = e.values[21];
  var veh4PD = e.values[22];
  var veh4Towing = e.values[23];
  var veh4Rental = e.values[24];
  
  var medPay = e.values[9];
  var aDnd = e.values[10];
  var underInsured = e.values[11];
  var bodilyInjury = e.values[12];
  var propertyDamage = e.values[13];
  var underInsuredPD = e.values[14];
  
  var templateFile = DriveApp.getFileById("1X-NNKLzy13RLa4ynBkWB50uZYSV63W_zwUxMrQD9GXM");
  //the blank form to be filled by the script
  var templateResponseFolder = DriveApp.getFolderById("1vjxp9eFhXiyOk6awpdjWyp1OcA8Va5Y_");
  //the destination folder for the filled forms
  
  var copy = templateFile.makeCopy(name, templateResponseFolder);
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
  //calls to all of our replaceText methods
  
  doc.saveAndClose();
  //saves and closes the newly created document
  
}
//Todo:
//- Fix all e.values
//- Change getFileById to new updated document with tags added
//- Change getFolderById to new updated folder location
//- Logic to create merged document from two(or more) seperate template documents
//- Logic to delete tag if not needed (ie: {{Veh 2 Year}} in a situation with only 1 vehicle)
//- Create variants for other lines of business and document types (Business Auto, Homeowners, General Liability, Workers Comp, Prom Note, Designation of Authorized Rep, Ride-sharing, etc)
