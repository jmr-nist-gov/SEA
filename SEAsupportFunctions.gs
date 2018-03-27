function addAffiliation(newAff){
  //------------------------------------------------------------------------------------------
  // Adds a new affiliation from form responses.
  //------------------------------------------------------------------------------------------
  //
  var n = [142, 148];  // List of form indices to update
  for (i=0; i<n.length; i++){
    addToFormDropdown(n[i], newAff);
  }
  //------------------------------------------------------------------------------------------
  // lastupdate: 20171007:1045                                                Jared M. Ragland
  // added support for DQ Time Point - updated indices 20180327:1212
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}

function addPatient(properties){
  //------------------------------------------------------------------------------------------
  // Fills the properties of a new Patient (by Research ID) from form responses.
  //------------------------------------------------------------------------------------------
  //
  var testing = 'FALSE';
  if (testing === 'TRUE'){
    var properties = ['testRID','testSPP','testCN','testSEX','testAGE','testMETHOD','testCONFIDENCE','testDOB','testCLASS'];
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
  var last = sheet.getLastRow();
  //
  //  Write responses for the new patient into the record to smooth future lookups
  //
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sourcePatients');
  var n = sheet.getLastRow()+1;
  sheet.getRange(n, 1).setValue(sheet.getRange(n-1,1).getValue()+1);
  sheet.getRange(n, 2).setValue(properties[0]);
  sheet.getRange(n, 3, 1, 9).setValues([properties]);
  if (sheet.getRange(n,10).getValue == ''){
    sheet.getRange(n,10).setValue('Unknown');
  }
  //
  //  Update the form to reflect the new patient choice.
  //
  addToFormDropdown(6, properties[0]);
  //------------------------------------------------------------------------------------------
  // lastupdate: 20171007:0942                                                Jared M. Ragland
  // added support for DQ Time Point - updated indices 20180327:1215
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}

function addPreparer(properties){
  //------------------------------------------------------------------------------------------
  // Fills the properties of new Personnel from form responses.
  //------------------------------------------------------------------------------------------
  //
  //  Write responses for the new preparer into the record to smooth future lookups
  //
  var testing = 'FALSE';
  if (testing === 'TRUE'){
    var properties = ['testAdd','testAff','testEmail'];
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sourcePersonnel');
  var i = sheet.getLastRow()+1;
  sheet.getRange(i,1).setValue(sheet.getRange(i-1,1).getValue()+1);
  sheet.getRange(i,2,1,3).setValues([properties]);
  sheet.getRange(i,5).setValue('Yes');
  //
  //  Update the form to reflect the new preparer choice.
  //
  var n = [34,36,69,71,147];  // List of form indices to update
  for (i=0; i<n.length; i++){
    addToFormDropdown(n[i], properties[0]);
  }
  //------------------------------------------------------------------------------------------
  // lastupdate: 20170104:1305                                                Jared M. Ragland
  // added support for DQ Time Point - updated indices 20180327:1218
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}

function addToFormDropdown(index, display){
  //------------------------------------------------------------------------------------------
  //  Adds the "display" value to the end of the dropdown options for choosing item "index".
  //  The need for this could be avoided in the future by having the item populate on load.
  //  These could benefit also from a sort, but sorting the choice objects is a pain... 
  //  for loop choice[i].getValue() -> [], etc.etc.  No method exists at this point.
  //  For now, the only workaround I could find to add a choice with navigation is to copy
  //  the .getGotoPage() object from the first item in the list, copy the properties from 
  //  the LAST item in the list (in this case '(New XXX)') and then write that back to the
  //  bottom of the list.  This only affects adding affiliations and personnel for now, but 
  //  may come in later.
  //------------------------------------------------------------------------------------------
  // 
  var form = FormApp.openById('1S5nyZ867nnMHTrgo3CW0Nm-ta14C0Xm8gPCEVOnsGc0');
  var items = form.getItems();
  var item = items[index].asListItem();
  var choices = item.getChoices();
  var navEnabled = choices[choices.length-1].getGotoPage();
  if (navEnabled === null){
    choices.push(item.createChoice(display));
  } else {
    var navInfo1 = choices[0].getGotoPage();
    var appendBack = choices[(choices.length-1)];
    var display2 = appendBack.getValue();
    var navInfo2 = appendBack.getGotoPage();
    choices.pop();
    choices.push(item.createChoice(display, navInfo1));
    choices.push(item.createChoice(display2, navInfo2));
  }
  item.setChoices(choices);
  //------------------------------------------------------------------------------------------
  // lastupdate: 20170104:1338                                                Jared M. Ragland
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}

function getLocationProperties(locationName){
  //------------------------------------------------------------------------------------------
  // Gets the County, State, and Community from the aquarium chosen in form responses.
  //------------------------------------------------------------------------------------------
  //
  //var locationName = 'Dolphin Quest Hawaii';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sourceLocations');
  var locations = sheet.getDataRange().getValues();
  //
  // Cycle through the existing locations (from the dropdown on the form) and grab
  //
  for (var i=0; i<locations.length; i++){
    if (locations[i][1] == locationName){
      return sheet.getRange(i+1,3,1,4).getValues()[0];
    }
  }
  return null;
  //------------------------------------------------------------------------------------------
  // lastupdate: 20170104:1109                                                Jared M. Ragland
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}

function getPatientProperties(researchID){
  //------------------------------------------------------------------------------------------
  // Gets the properties of a given Patient (by name) and returns them as an array.
  //------------------------------------------------------------------------------------------
  //
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sourcePatients');
  var patients = sheet.getDataRange().getValues();
  //
  // Cycle through the existing patients (from the dropdown on the form) and grab
  //
  for (var i=0; i<patients.length; i++){
    if (patients[i][1]==researchID){
      return sheet.getRange(i+1,2,1,10).getValues()[0];
    }
  }
  return null;
  //------------------------------------------------------------------------------------------
  // lastupdate: 20170104:0856                                                Jared M. Ragland
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}

function getPersonnelProperties(personnelName){
  //------------------------------------------------------------------------------------------
  // Gets the properties of a given Patient (by name) and returns them as an array.
  //------------------------------------------------------------------------------------------
  //
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sourcePersonnel');
  var personnel = sheet.getDataRange().getValues();
  //
  // Cycle through the existing patients (from the dropdown on the form) and grab
  //
  for (var i=0; i<personnel.length; i++){
    if (personnel[i][1]==personnelName){
      return sheet.getRange(i+1,2,1,4).getValues()[0];
    }
  }
  return null;
  //------------------------------------------------------------------------------------------
  // lastupdate: 20180319:1000                                                Jared M. Ragland
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}

function getPermitInfo(permitNumber){     // <- for latest, supply "last"
  //------------------------------------------------------------------------------------------
  // Returns the latest permit information for inclusion in the chain of custody footer.
  //------------------------------------------------------------------------------------------
  //
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sourcePermits');
  var permits = sheet.getDataRange().getValues();
  //
  // Cycle through the existing patients (from the dropdown on the form) and grab
  //
  if (permitNumber === 'last'){
    return sheet.getRange(permits.length, 2,1,5).getValues();
  } else {
    for (var i=1; i<permits.length+1; i++){
      if (permits[i][0]==permitNumber){
        return sheet.getRange(i+1,2,1,5).getValues();
      }
    }
  }
  return null;
  //------------------------------------------------------------------------------------------
  // lastupdate: 20171002:1345                                                Jared M. Ragland
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}

function countSamples(counts){
  //------------------------------------------------------------------------------------------
  // Returns the total sample count, rounding partials to +1.
  //------------------------------------------------------------------------------------------
  //
  return Math.ceil(
    counts.reduce(
      function (total, num){
        return total + num;
      }
    )
  );
  //------------------------------------------------------------------------------------------
  // lastupdate: 20171006:1525                                                Jared M. Ragland
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}

// Helper functions to list both responses and form items by index and name/title, and tie 
//  them together by resolving their indices against one another.
function listAll(){
  listResponseHeaders();
  listFormItems();
}

function listResponseHeaders(){
  var sheetResponses = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
  var cols = sheetResponses.getLastColumn();
  var headers = sheetResponses.getSheetValues(1, 1, 1, cols);
  var refResponses = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('refResponseHeaders');
  refResponses.clear();
  refResponses.getRange(1,1,1,3).setValues([["ArrayIndex","ColumnIndex","ResponseHeader"]]);
  var temp = '';
  for (var i=0; i<headers[0].length; i++){
    temp = headers[0][i];
    refResponses.getRange(i+2,1).setValue(temp);
    refResponses.getRange(i+2,2).setValue(i);
    refResponses.getRange(i+2,3).setValue(i+1);
  }
}
function listFormItems(){
  var form = FormApp.openById('1S5nyZ867nnMHTrgo3CW0Nm-ta14C0Xm8gPCEVOnsGc0');
  var refFormItems = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('refFormItems');
  refFormItems.clear()
  refFormItems.getRange(1, 1, 1, 3).setValues([['Index','Title','Type']]);
  var items = form.getItems();
  for (var i=0; i<items.length; i++){
    refFormItems.getRange(i+2,1).setValue(items[i].getTitle());
    refFormItems.getRange(i+2,2).setValue(items[i].getIndex());
    refFormItems.getRange(i+2,3).setValue(items[i].getType());
  }  
}
//------------------------------------------------------------------------------------------
// lastupdate: 20171006:1050                                                Jared M. Ragland
//                                                     NIST Marine ESB Data Tool Development
//------------------------------------------------------------------------------------------



// Found from http://www.codesuck.com/2012/02/transpose-javascript-array-in-one-line.html
// Transposes an array
function transpose(a){
  return a[0].map(function (_, c) { return a.map(function (r) { return r[c]; }); });
}