function fillAndSendReportTemplate(responses){
//function fillAndSendReportTemplate(){  //comment prior and uncomment this for testing purposes
  //------------------------------------------------------------------------------------------
  // Populates the report template directly from the responses array provided.
  //------------------------------------------------------------------------------------------
  //
  // Placeholder to test by grabbing any given line in the data source
  //
  var testingMode = false; 
  var sendmail = true;
  if (testingMode) {
    sendmail = false;
    var recordSpecific = 3;         // <- change this to the line of interest from the sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
    var AllResponses = sheet.getDataRange().getValues();
    var recordOfInterest = recordSpecific;
    var responses = AllResponses[recordOfInterest-1];
  }
  //    - set up date formatting assistance and spacing concessions and response index arrays
  var n_formatAsDateTime=[13,23,25,50,52];
  var formattedAsDateTime='';
  var n_formatAsDateOnly=[10,63,67,71,75,83,79,88,121];
  var formattedAsDateOnly='';
  var n_formatAsTimeOnly=[31,39,46,56];
  var formattedAsTimeOnly='';
  var n_preyItemSpacing=[62,63,64,70,71,72,82,83,84];
  var replacedText='';
  var affiliation = '';
  var sampleCountX=[28,35,36,43,54,55];
  var sampleCountN=[112,113,114,115];
  var forPermit = [117,118,119,120,121];
  var forEmail = [104,105,2,116,112,113,114,115,117,118,119,121];
  //
  //
  //------------------------------------------------------------------------------------------
  //  
  //  Populate the report itself...
  //
  var docTemplateID = '1FyFTDGNKMIy27so9aPk2_njzVzVlj6LBiKRlj394by4';
  var docTemplate = DriveApp.getFileById(docTemplateID);
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var date = responses[13];
  date = days[date.getDay()] + ' ' + months[date.getMonth()] + ' ' + date.getDate() + ' ' + ' ' + date.getFullYear()
  var title = responses[2] + ' on ' + date;
  var reportCopy = docTemplate.makeCopy('DQ SEA - ' + title);
  var reportID = reportCopy.getId();
  var report = DocumentApp.openById(reportID);
  var reportBody = report.getBody();
  var reportFooter = report.getFooter();
  var tab = '                ';
  //
  //  ...prep spots for all sections requiring specific numerical language...
  //    - check no reproductive condition
  if (responses[20] == 'none of the above') {
    responses[20] = tab+tab;
  }
  //    - check for blood samples
  var bloodText = '<<bloodText>>';
  var bloodTextNo = 'not collected';
  var bloodTextYes = 'collected <<entry23>> by <<entry24>> and processed <<entry25>> by <<entry26>>';
  if (responses[22]=="No"){
    reportBody.replaceText(bloodText, bloodTextNo);
  } else {
    reportBody.replaceText(bloodText, bloodTextYes);
  }
  //    - check presence and extra volume remaining for blood samples
  var extraWholeBlood = '<<extraWholeBlood>>';
  var extraWholeBloodNo = '';
  var extraWholeBloodYes = ' AND 1 aliquot with ~<<entry29>> mL';
  if (responses[29]>0){
    reportBody.replaceText(extraWholeBlood, extraWholeBloodYes);
  } else {
    reportBody.replaceText(extraWholeBlood, extraWholeBloodNo);
  }
  var extraPlasma = '<<extraPlasma>>';
  var extraPlasmaNo = '';
  var extraPlasmaYes = ' AND 1 aliquot with ~<<entry37>> mL';
  if (responses[37]>0){
    reportBody.replaceText(extraPlasma, extraPlasmaYes);
  } else {
    reportBody.replaceText(extraPlasma, extraPlasmaNo);
  }
  var extraSerum = '<<extraSerum>>';
  var extraSerumNo = '';
  var extraSerumYes = ' AND 1 aliquot with ~<<entry44>> mL';
  if (responses[44]>0){
    reportBody.replaceText(extraSerum, extraSerumYes);
  } else {
    reportBody.replaceText(extraSerum, extraSerumNo);
  }
  //    - check for milk samples
  var milkText = '<<milkText>>';
  var milkTextNo = 'not collected';
  var milkTextYes = 'collected <<entry50>> by <<entry51>> and processed <<entry52>> by <<entry53>>'
  if (responses[49]=="No"){
    reportBody.replaceText(milkText, milkTextNo);
  } else {
    reportBody.replaceText(milkText, milkTextYes);
  }
  //
  var i=0;
  //    - populate the template
  while (i < responses.length){
    replacedText = "<<entry" + i + ">>";
    if (responses[i] === ''){
        if (n_preyItemSpacing.indexOf(i)>-1) {
          reportBody.replaceText(replacedText, '');
        } else if (i === 58) {
          reportBody.replaceText('<<entry' + i + '>>', 'Not Applicable');
        } else if (i === 107) {
          reportBody.replaceText('<<entry' + i + '>>', '(None provided)');
        }
      reportBody.replaceText(replacedText, tab);
    } else {
      if (n_formatAsDateTime.indexOf(i)>-1){
        date = responses[i];
        formattedAsDateTime=date.getDate() + ' ' + months[date.getMonth()] + ' ' + date.getFullYear();
        if (date.getHours()<10){
          formattedAsDateTime+= ' 0' + date.getHours() + ':';
        } else {
          formattedAsDateTime+= ' ' + date.getHours() + ':';
        }
        if (date.getMinutes()<10){
          formattedAsDateTime+= '0' + date.getMinutes();
        } else {
          formattedAsDateTime+= date.getMinutes();
        }
        reportBody.replaceText(replacedText, formattedAsDateTime);
      } else if (n_formatAsDateOnly.indexOf(i)>-1){
        date = responses[i];
        formattedAsDateOnly=date.getDate() + ' ' + months[date.getMonth()] + ' ' + date.getFullYear();
        reportBody.replaceText(replacedText, formattedAsDateOnly);
      } else if (n_formatAsTimeOnly.indexOf(i)>-1){
        date = responses[i];
        if (date.getHours()<10){
          formattedAsTimeOnly='0'+date.getHours() + ':';
        } else {
          formattedAsTimeOnly=date.getHours() + ':';
        }
        if (date.getMinutes()<10){
          formattedAsTimeOnly+= '0' + date.getMinutes();
        } else {
          formattedAsTimeOnly+= date.getMinutes();
        }
        reportBody.replaceText(replacedText, formattedAsTimeOnly);
      } else if (sampleCountX.indexOf(i)>-1) {
        reportBody.replaceText(replacedText, responses[i].toString() + "x");
      } else if (sampleCountN.indexOf(i)>-1) {
        reportBody.replaceText(replacedText, "n=" + responses[i].toString());
      } else {
        if (i === 9){
          reportBody.replaceText(replacedText, responses[i].toString());
          if (responses[i] == "estimated" && responses[i-1] == "animal history") {
            reportBody.replaceText('<<place1>>', '');
          } else {
            reportBody.replaceText('<<place1>>', tab);
          }
        } else if (i === 30){
          reportBody.replaceText(replacedText, responses[i].toString());
          if (responses[i].toString().length < 18) {
            reportBody.replaceText('<<place2>>',tab);
          } else {
            reportBody.replaceText('<<place2>>','');
          }
        } else if (i === 45){
          reportBody.replaceText(replacedText, responses[i].toString());
          if (responses[i].toString() == "Not applicable") {
            reportBody.replaceText('<<place3>>', tab);
          } else {
            reportBody.replaceText('<<place3>>', '');
          }
        } else {
          reportBody.replaceText(replacedText, responses[i].toString());
        }
      }
    }
    i++;
  }
  //   - populate the permit information in the footer
  var permit = getPermitInfo('last')[0];
  i = 0;
  while (i < permit.length){
    if (i == (permit.length-1)){
      formattedAsDateOnly = permit[i].getDate() + ' ' + months[permit[i].getMonth()] + ' ' + permit[i].getFullYear();
      reportFooter.replaceText('<<entry' + forPermit[i] + '>>', formattedAsDateOnly);
    } else {
      reportFooter.replaceText('<<entry' + forPermit[i] + '>>', permit[i].toString());
    }
    i++;
  }
  //  
  //  ...convert to a .pdf...
  //
  var pdf = report.getAs("application/pdf");
  report.saveAndClose();
  //  
  //  ...and email to recipients.
  //
  var emailTemplateID = '1cLPHko20riKxKbPVoUUDub2y4CZtHftG7GAnptpOETg';
  var emailTemplate = DriveApp.getFileById(emailTemplateID);
  var emailCopy = emailTemplate.makeCopy('DQ Sea Email (DELETE ME)');
  var emailID = emailCopy.getId();
  var email = DocumentApp.openById(emailID);
  var emailBody = email.getBody();
  //
  //  Set up the replacements to generate both HTML and plain text format emails.
  //  First replace placeholders with response-driven data...
  //
  i = 0;
  while (i < forEmail.length){
    if (i == 4 || i == 6){
      if (responses[forEmail[i]] == 1) {
        emailBody.replaceText('<<entry' + i + '>>', responses[forEmail[i]] + ' cryovial');
      } else {
        emailBody.replaceText('<<entry' + i + '>>', responses[forEmail[i]] + ' cryovials');
      }
    } else if (i == 5 || i == 7){
      if (responses[forEmail[i]] == 1) { 
        emailBody.replaceText('<<entry' + i + '>>', responses[forEmail[i]] + ' cryovial or Teflon jar');
      } else {
        emailBody.replaceText('<<entry' + i + '>>', responses[forEmail[i]] + ' cryovials or Teflon jars');
      }
    } else if (i == 11){
      formattedAsDateOnly = responses[forEmail[i]].getDate() + ' ' + months[responses[forEmail[i]].getMonth()] + ' ' + responses[forEmail[i]].getFullYear();
      emailBody.replaceText('<<entry' + i + '>>', formattedAsDateOnly);
    } else {
      emailBody.replaceText('<<entry' + i + '>>', responses[forEmail[i]]);
    }
    i++;
  }
  //
  //  ...then make a copy of that text in HTML format to preserve formatting...
  //
  var emailBodyasHTML = emailBody.getText();
  //
  //  ...append the DTD logo...
  //
  //    - uncomment the following when the logo is that in an accessible location
  //------------------------------------------------------------------------------------------
  //var img = '"https://lh5.googleusercontent.com/1hIOiJSYj-sjjkexWoX40JtzuuDAXn7J2SyxlO6JmxnZjzmUCqF4vDpfFqvPOGMFm6Emp7cf22vHfGE=w1040-h753"';
  //img += ' style="width: 30%; height: 30%" align="right"';
  //emailBodyasHTML += '<img src=' + img + '/>';
  emailBodyasHTML +='<p style="font-size:8px" align="right">A product of NIST Environmental Specimen Bank Data Tool Development</p>';
  //------------------------------------------------------------------------------------------
  //
  //
  //  ...then strip out the HTML tags for the plain text version...
  //
  var tags = ['br','h3','hr','ul','li','/h3','/ul','/li'];
  var tagRep = ['','---------------------------------------------------------------------------','','','','',''];
  i=0;
  while (i < tags.length){
    emailBody.replaceText('<' + tags[i] + '>', tagRep[i]);
    i++;     
  }
  //
  //  ...append the DTD tag...
  //
  emailBody = emailBody.getText() +  'A product of NIST Environmental Specimen Bank Data Tool Development';
  //
  //------------------------------------------------------------------------------------------
  //  ...and send the email.
  //
  //    - always send it to those in this list:
  var recipient = "";
  if (responses[12]=="Dolphin Quest Hawaii") {
    recipient += "nlambert@dolphinquest.com,kjohnson@dolphinquest.com";
  } else if (responses[12]=="Dolphin Quest Oahu") {
    recipient += "jrocho@dolphinquest.com,nwest@dolphinquest.com";
    recipient += ',' + getPersonnelProperties(responses[104])[2];
  }
  var ccList = "jared.ragland@noaa.gov,colleen.bryan@noaa.gov";
  //    - and then add the form preparer
  //recipient += ',' + getPersonnelProperties(responses[104])[3];
  //    - optionally send it to anyone else
  //recipient += ',' + newPerson;
  //------------------------------------------------------------------------------------------
  var subject = "Dolphin Quest SEA - New Sample Entry Recorded for " + title;
  if (sendmail) {
    MailApp.sendEmail({
      to: recipient, 
      cc: ccList,
      subject: subject, 
      body: emailBody,
      htmlBody: emailBodyasHTML, 
      attachments: pdf,
      noReply: true
    });
  }
  //
  //  Close out and clean up Drive.
  //
  email.saveAndClose();
  DriveApp.getFileById(emailID).setTrashed(true);
  //------------------------------------------------------------------------------------------
  // lastupdate: 20180319:1000                                                Jared M. Ragland
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}

function samplesCollectedNoChain(responses){
  //------------------------------------------------------------------------------------------
  // Informs the assistant curator that a record was created but no chain was generated
  // by user choice.  We may want to eliminate this functinoality in the future by
  // automatically choosing whether or not to generate a chain of custody based on the
  // presence of samples.  Leaving this as an option will allow for sample records to be
  // created but the tool NOT being used to generate chains of custody.
  //------------------------------------------------------------------------------------------
  //
  var patient = responses[2];
  var recordedBy = responses[104];
  var recordedOn = responses[0];
  var subject = "Dolphin Quest SEA - Samples collected without a chain of custody";
  var recipient = "jared.ragland@noaa.gov";
  recipient += "colleen.bryan@noaa.gov";
  var emailBody = "Samples for " + patient;
  emailBody += " were recorded by " + recordedBy;
  emailBody += " on " + recordedOn;
  emailBody += " as collected, but a chain of custody was not created.  Contact a data custodian to generate a chain if necessary."
  MailApp.sendEmail({
    to: recipient, 
    subject: subject,
    body: emailBody,
    noReply: true
  });
  //
  //------------------------------------------------------------------------------------------
  // lastupdate: 20171006:1618                                                Jared M. Ragland
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}

function conditionUpdated(responses){
  //------------------------------------------------------------------------------------------
  // Informs the assistant curator that a record was created but no chain was generated
  // by user choice.  We may want to eliminate this functinoality in the future by
  // automatically choosing whether or not to generate a chain of custody based on the
  // presence of samples.  Leaving this as an option will allow for sample records to be
  // created but the tool NOT being used to generate chains of custody.
  //------------------------------------------------------------------------------------------
  //
  var patient = responses[2];
  var recordedBy = responses[104];
  var recordedOn = responses[0];
  var subject = "Dolphin Quest SEA - Patient condition was updated";
  var recipient = "jared.ragland@noaa.gov";
  recipient += "colleen.bryan@noaa.gov";
  var emailBody = "A condition update for " + patient;
  emailBody += " was recorded by " + recordedBy;
  emailBody += " on " + recordedOn;
  emailBody += " .  No samples were collected."
  MailApp.sendEmail({
    to: recipient, 
    subject: subject,
    body: emailBody,
    noReply: true
  });
  //
  //------------------------------------------------------------------------------------------
  // lastupdate: 20171014:1024                                                Jared M. Ragland
  //                                                     NIST Marine ESB Data Tool Development
  //------------------------------------------------------------------------------------------
}
