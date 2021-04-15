//Google Apps Script

//Project OAuth Scopes, 2 Scopes Requested:
//See, edit, create, and delete all of your Google Drive files	https://www.googleapis.com/auth/drive
//View and manage your Google Docs documents	https://www.googleapis.com/auth/documents

//Based on the following tutorial which can better explain this use case, https://jeffreyeverhart.com/2018/09/17/auto-fill-google-doc-from-google-form-submission/

function autoFillGoogleDocFromForm(e) {
  //e.values is an array of form values
  var name = e.values[2];
  
  var templateFile = DriveApp.getFileById("137B0rhUsCPcQRmSseM43eBUIawkarG2M7Klx42mrl2Y"); //coverage rejection form
  //the blank form to be filled by the script
  var templateResponseFolder = DriveApp.getFolderById("1NmqXBcKqdu_63cD02tVbEhRTNceTR1Dg");
  //the destination folder for the filled forms
  
  var copy = templateFile.makeCopy(name + ' coverage rejection form', templateResponseFolder);
  //creates a copy of the blank form templateFile and gives it a name
  
  var doc = DocumentApp.openById(copy.getId());
  //opens the document
  
  var body = doc.getBody();
  //gets the body of the document
  
  var replacementTagArray = [
    {{Today's Date}},
    {{Insured's Name}},
    {{Policy Number}},
    {{Veh 1 Year}},
    {{Veh 1 Make}},
    {{Veh 1 Comp/Collision}},
    {{Veh 1 Towing}},
    {{Veh 1 Rental}},
    {{Medical Payments}},
    {{Rideshare}},
    {{AD&D}},
    {{Combined Un & Underinsured}},
    {{Bodily Injury}},
    {{Property Damage}},
    {{Underinsured PD}},
    {{Veh 2 Year}},
    {{Veh 2 Make}},
    {{Veh 2 Comp/Collision}},
    {{Veh 2 Towing}},
    {{Veh 2 Rental}},
    {{Veh 3 Year}},
    {{Veh 3 Make}},
    {{Veh 3 Comp/Collision}},
    {{Veh 3 Towing}},
    {{Veh 3 Rental}},
    {{Veh 4 Year}},
    {{Veh 4 Make}},
    {{Veh 4 Comp/Collision}},
    {{Veh 4 Towing}},
    {{Veh 4 Rental}}
  ]
  
  //skip i = 0 aka e.values[0] because it is unused timestamp from spreadsheet first column
  for(var i = 1;i < e.length; i++){
    body.replaceText("^replacementTagArray[i]$", e.values[i]);
  }
  //calls to all of our replaceText methods
  
  doc.saveAndClose();
  //saves and closes the newly created document
  
//   var templateFile2 = DriveApp.getFileById("1mmLvQah5uxSjKBP_zfyCqXwftISGpEKtrJ6Q-6ppGwI"); //policy fee form
//   //the blank form to be filled by the script
//   var templateResponseFolder = DriveApp.getFolderById("1NmqXBcKqdu_63cD02tVbEhRTNceTR1Dg");
//   //the destination folder for the filled forms
  
//   var copy2 = templateFile2.makeCopy(name + ' policy fee form', templateResponseFolder);
//   //creates a copy of the blank form templateFile and gives it a name
  
//   var doc2 = DocumentApp.openById(copy2.getId());
//   //opens the document
  
//   var body2 = doc2.getBody();
//   //gets the body of the document
  
//   body2.replaceText("{{Today's Date}}", date);
//   body2.replaceText("{{Insured's Name}}", name);
//   body2.replaceText("{{Policy Number}}", policyNumber);
//   body2.replaceText("{{Veh 1 Year}}", veh1Year);
//   body2.replaceText("{{Veh 1 Make}}", veh1Make);
//   body2.replaceText("{{Veh 1 Comp/Collision}}", veh1PD);
//   body2.replaceText("{{Veh 1 Towing}}", veh1Towing);
//   body2.replaceText("{{Veh 1 Rental}}", veh1Rental);
//   body2.replaceText("{{Medical Payments}}", medPay);
//   body2.replaceText("{{Rideshare}}", rideShare);
//   body2.replaceText("{{AD&D}}", aDnd);
//   body2.replaceText("{{Combined Un & Underinsured}}", underInsured);
//   body2.replaceText("{{Bodily Injury}}", bodilyInjury);
//   body2.replaceText("{{Property Damage}}", propertyDamage);
//   body2.replaceText("{{Underinsured PD}}", underInsuredPD);
//   body2.replaceText("{{Veh 2 Year}}", veh2Year);
//   body2.replaceText("{{Veh 2 Make}}", veh2Make);
//   body2.replaceText("{{Veh 2 Comp/Collision}}", veh2PD);
//   body2.replaceText("{{Veh 2 Towing}}", veh2Towing);
//   body2.replaceText("{{Veh 2 Rental}}", veh2Rental);
//   body2.replaceText("{{Veh 3 Year}}", veh3Year);
//   body2.replaceText("{{Veh 3 Make}}", veh3Make);
//   body2.replaceText("{{Veh 3 Comp/Collision}}", veh3PD);
//   body2.replaceText("{{Veh 3 Towing}}", veh3Towing);
//   body2.replaceText("{{Veh 3 Rental}}", veh3Rental);
//   body2.replaceText("{{Veh 4 Year}}", veh4Year);
//   body2.replaceText("{{Veh 4 Make}}", veh4Make);
//   body2.replaceText("{{Veh 4 Comp/Collision}}", veh4PD);
//   body2.replaceText("{{Veh 4 Towing}}", veh4Towing);
//   body2.replaceText("{{Veh 4 Rental}}", veh4Rental);
  //calls to all of our replaceText methods
  
//   doc2.saveAndClose();
  //saves and closes the newly created document  
}
//Todo:
//- Change getFileById to new updated document with tags added (for other template options)
//- Create variants for other lines of business and document types (Business Auto, Homeowners, General Liability, Workers Comp, Prom Note, Designation of Authorized Rep, Ride-sharing, etc)
