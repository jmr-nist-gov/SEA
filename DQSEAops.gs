function onFormSubmit(){
  //------------------------------------------------------------------------------------------
  // Activates on submission of the Dolphin Quest Sample Entry Assistant form and:
  //   - gathers the form responses from the linked DQ SEA data spreadsheet;
  //   - fills in any known values skipped during form entry as known for a given patient;
  //   - fills the report template
  //   - emails a pdf of the completed template to the form preparer and Colleen.Bryan@nist.gov
  //------------------------------------------------------------------------------------------
  //
  var testingmode = false;
  //
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
  var responses = sheet.getDataRange().getValues();
  var consolidated = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('consolidated');
  var clR = consolidated.getLastRow();
  var clC = consolidated.getLastColumn();
  var colsForPatients = [4,4,5,6,7,8,9,10,11,12];
  var colsForFacilities = [109,110,111,112];
  //
  // vals.length refers to the RANGE and not to the array, the array dim will always be -1
  //
  var last = responses.length;
  responses = responses[last-1];
  var i = 0;       // loop counter initialization
  var n = [];      // index reference initialization
  //
  //------------------------------------------------------------------------------------------
  // Replace any extraneous line feeds entered into PARAGRAPH_TEXT form items
  //
  var colsWithReturns = [34,42,49,60,108];
  var sansReturns = "";
  for (i=0; i<colsWithReturns.length; i++){
    sansReturns = sheet.getRange(last, colsWithReturns[i]).getValue().replace(/\n/g, "; ");
    sheet.getRange(last, colsWithReturns[i]).setValue(sansReturns);
    responses[colsWithReturns[i]-1] = sansReturns;
  }
  //
  //------------------------------------------------------------------------------------------
  // Check to see if this is a new Patient 
  //
  if (responses[1] == "Yes" & getPatientProperties(responses[3]) == null){
    //
    // If so, fill in the Patient properties in sourcePatients from that provided in the form
    //
    addPatient(responses[3, 4, 5, 6, 7, 8, 9, 10, 11]);
    //
    // Backfill the selected Research ID to simplify form generation
    //
    sheet.getRange(last, 3).setValue(responses[3]);
    responses[2] = responses[3];
    //
  } else {
    //
    // Otherwise grab the Patient name from the latest response, get the patient properties...
    //
    var patientProperties = getPatientProperties(responses[2]);
    //  
    // ...and fill in the skipped form responses due to being a known patient
    //   **This may need to be separated if they want a complete record of ALL age measurements**
    //    **Split age information out from sourcePatients
    //    **Create getPatientAge(){} to look for -
    //          the latest entry, OR
    //          add logic to prioritize "known" ages over "estimated" ages
    //
    // Fill Research ID, Species, Common Name, Sex, Age, Method of Aging, Age Confidence, Date of Birth, and Age Class
    //
    sheet.getRange(last, 3, 1, colsForPatients.length).setValues([patientProperties]);
  }
  //
  //------------------------------------------------------------------------------------------
  // Check to see if this is a new Affiliation
  //
  if (responses[102] == "(New Affiliation)" | responses[105] == "(New Affiliation)"){
    //
    // Change the entry of "(New Affiliation)" to that provided later in the form...
    //
    responses[102] = responses[103];
    responses[105] = responses[103];
    sheet.getRange(last, 103).setValue(responses[103]);
    sheet.getRange(last, 106).setValue(responses[103]);
    //
    // ...and add that new option as a valid entry in the affiliation dropdowns.
    //  ** Note this does not allow current users to CHANGE their affiliation.
    //  ** This will have to be done as part of DQ SEA maintenance OR
    //  ** through alteration of the form and structure.
    //
    addAffiliation(responses[103]);
  }
  // Check to see if this is a new Preparer
  if (responses[104]=="(New Personnel)"){
    //
    // If so, match the names and affiliations to those provided in the form...
    //
    responses[104] = responses[100];
    sheet.getRange(last, 105).setValue(responses[100]);    // New Name
    //
    // ...and check to see if there's a new affiliation to resolve.
    // ...and add that person to the list for preparation of future forms.
    //
    addPreparer(responses[100, 101, 102]);
  }
  //
  //------------------------------------------------------------------------------------------
  // Fill out the location info from the chosen aquarium
  //  Currently no support to automatically add facilities
  //
  var aquariumProperties = getLocationProperties(responses[12]);
  for (i=0; i<colsForFacilities.length; i++){
    sheet.getRange(last,colsForFacilities[i]).setValue(aquariumProperties[i]);
  }
  //
  //------------------------------------------------------------------------------------------
  // Fill out missing sample counts
  //
  //   - If blood samples were not collected, fill the missing collection statuses
  if (sheet.getRange(last,23).getValue()=="No"){
    sheet.getRange(last,28).setValue("No");
    sheet.getRange(last,35).setValue("No");
    sheet.getRange(last,43).setValue("No");
  }
  //      - Whole Blood
  if (sheet.getRange(last,28).getValue()=="No"){
    sheet.getRange(last,29).setValue(0);
    sheet.getRange(last,30).setValue(0);
    sheet.getRange(last,31).setValue("Not applicable");
    sheet.getRange(last,113).setValue(0);
  } else {
    sheet.getRange(last,113).setValue(countSamples(responses.slice(28,30)));
  }
  //      - Plasma
  if (sheet.getRange(last,35).getValue()=="No"){
    sheet.getRange(last,36).setValue(0);
    sheet.getRange(last,37).setValue(0);
    sheet.getRange(last,38).setValue(0);
    sheet.getRange(last,39).setValue("Not applicable");
    sheet.getRange(last,114).setValue(0);
  } else {
    sheet.getRange(last, 114).setValue(countSamples(responses.slice(35,38)));
  }
  //      - Serum
  if (sheet.getRange(last,43).getValue()=="No"){
    sheet.getRange(last,44).setValue(0);
    sheet.getRange(last,45).setValue(0);
    sheet.getRange(last,46).setValue("Not applicable");
    sheet.getRange(last,115).setValue(0);
  } else {
    sheet.getRange(last, 115).setValue(countSamples(responses.slice(43,45)));
  }
  //      - Milk
  if (sheet.getRange(last,50).getValue()=="No"){
    sheet.getRange(last,55).setValue(0);
    sheet.getRange(last,56).setValue(0);
    sheet.getRange(last,116).setValue(0);
  } else {
    sheet.getRange(last, 116).setValue(countSamples(responses.slice(54,56)));
  }
  var nTotal = sheet.getRange(last, 113, 1, 4).getValues();
  sheet.getRange(last, 117).setValue(countSamples(nTotal[0]));
  //
  //------------------------------------------------------------------------------------------
  // Get LAST permit
  //
  sheet.getRange(last, 118, 1, 5).setValues(getPermitInfo('last'));
  //------------------------------------------------------------------------------------------
  // Fill out the report template and send it if samples were recorded
  //
  // - refresh the responses list
  if (testingmode){
    Logger.log(responses);
  } else {
    // refresh responses
    responses = sheet.getDataRange().getValues();
    responses = responses[responses.length-1];
    // - add to archive
    consolidated.getRange(clR, 1, 1, clC).copyTo(consolidated.getRange(clR+1, 1, 1, clC));
    // - submit to fill and report routine
    //     - array index 21 should be 'Were samples collected?'
    //     - array index 106 should be 'Create a chain of custody document?'
    // CURRENTLY NO ERROR TRAP FOR 'Yes' TO 'Were samples collected?' AND 'No' to each tissue
    //  - maybe by countSamples(nTotal[0]) > 0 ?
    if (responses[21] === 'Yes'){
      if (responses[106] === 'Yes'){
        fillAndSendReportTemplate(responses);
      } else {
        samplesCollectedNoChain(responses);
      }
    } else {
      conditionUpdated(responses);
    }
  }
  //
  //------------------------------------------------------------------------------------------
  // lastupdate: 20171014:1024                                                Jared M. Ragland
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}