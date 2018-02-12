function exportToFW(project, buildAll, buildSpecific){
  var project = "Dolphin Quest";  // Override the project supplied, comment to active function argument
  var buildAll = true;   // Set true to regenerate ALL records, comment to activate function argument
  var buildSpecific = null;  // Grab a specific consolidated record, set null for last, comment to active function argument
  var src = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('consolidated'),
      dat = src.getDataRange().getValues();
  var lastDat = dat.length;
  dat = dat.slice(2, lastDat);
  lastDat = dat.length -1;
  var dest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FWexport'),
      xprt = dest.getDataRange(),
      destLC = xprt.getLastColumn(),
      destLR = xprt.getLastRow();
  var srcColsDirect = [],
      destColsDirect = [],
      srcColsFish = [[]],
      destColsFish = [],
      labelsFish = [],
      srcColsPermit = [[]],
      destColsPermit = [],
      srcColnSamples = [],
      destColsSamples = [],
      srcColsWB = [],
      srcColsPL = [],
      srcColsSR = [],
      srcColsMK = [],
      tissueTypes = [],
      sampled = [],
      processed = [],
      out = [[]],
      record = [],
      direct = [],
      tempSrcArr = [],
      tempDat = [],
      tempDestArr = [];
  var nSamples = 0,
      xSample = null,
      i = 0,
      ii = 0,
      n = 0,
      ndat = 0,
      si = 0,
      sii = 0,
      outi = 0,
      tissuei = 0,
      collector = null,
      processor = null,
      storeDateTime = null,
      storeTemp = null,
      notes = null,
      lot = null;
  
  if (project=="Dolphin Quest") {
    srcColsDirect = getIndices("srcColsDirect");
    destColsDirect = getIndices("destColsDirect");
    srcColsFish = getIndices("srcColsFish");
    destColsFish = getIndices("destColsFish");
    srcColsPermit = getIndices("srcColsPermit");
    destColsPermit = getIndices("destColsPermit");
    srcColsWB = [31, 32, 33, 34];
    srcColsPL = [39, 40, 41, 42];
    srcColsSR = [46, 47, 48, 49];
    srcColsMK = [57, 58, 60];
    destColsSamples = [36, [51, 49], 50, 21];
    
    // Length of srcColnSamples and tissueTypes must be equal
    srcColnSamples = [29, 36, 37, 44, 55, 56];
    tissueTypes = ["Whole Blood", "Plasma", "Plasma", "Serum", "Milk", "Milk"];
    
    // Set stopper for only building the last entry (testing and full operations)
    // Set buildAll = true to dump the entire contents of SEA
    // Catch buildSpecific specified but forgot to set buildAll to false
    if (buildSpecific != null) {buildAll = false;}
    if (buildAll) {
      ndat = dat.length;
    } else {
      ndat = 1;
    }
    for (n=0; n<ndat; n++){
      // Reset containers and counters
      direct = [];
      record = [];
      tissuei = 0;
      // Get the correct record in the dataset
      if (n<ndat) {
        tempDat = dat[n].slice(0);
      } else {
        tempDat = dat[lastDat].slice(0);
      }
      if (buildSpecific != null) {tempDat = dat[buildSpecific-1].slice(0);}
      // Those that will always go in need to be repeated for the total number of samples
      direct = buildRecordDirect(tempDat, srcColsDirect, destColsDirect);
      // Parse the permit information into a single column "Federal Permit"
      direct[destColsPermit[0]-1] = buildRecordPermit(tempDat, srcColsPermit, destColsPermit);
      // Split the reproductive status information
      direct = buildRecordReproduction(direct, tempDat, 21, [42, 34]);
      // Add the dietary information
      storeDateTime = splitDate(tempDat[60]);  // Last fed
      direct[17] = storeDateTime[0];
      direct[18] = storeDateTime[1];
      direct = buildRecordFish(direct, tempDat, srcColsFish, destColsFish);
      
      // Iterate over all samples collected to fill aliquot-specific information
      for (i=0; i<srcColnSamples.length; i++){
        xSample = 0;
        nSamples = 0;
        nSamples = tempDat[srcColnSamples[i]-1];
        if (srcColnSamples[i] == 36 | srcColnSamples[i] == 55 | srcColnSamples[i] == 56) { 
          xSample = 0;
        } else {
          xSample = tempDat[srcColnSamples[i]];
          if (xSample == '') {xSample = 0;}
        }
        if (nSamples > 0){
          for (ii=0; ii<(nSamples + Math.ceil(xSample)); ii++){
            record = direct.slice(0);
            
            // Remove the first two entries to leave it blank for MESB to fill in GUAID and DQ Time Point
            record[0] = null;
            record[1] = null;  // Time point, when included in SEA, remove this
            
            // Set the tissue type
            record[4] = tissueTypes[tissuei];
            
            // Set the current amount, adding the remainder to the LAST aliquot
            if (srcColnSamples[i] == 36 | srcColnSamples[i] == 55) {
              record[5] = 2.5;
            } else {
              record[5] = 1;
            }
            if (ii == nSamples){
              record[5] = xSample;
            }
            
            // Set sampling and processing dates and times by tissue type
            sampled = null;
            collector = null;
            processed = null;
            processor = null;
            if (record[4] == "Milk") {
              sampled = splitDate(tempDat[50]);
              collector = tempDat[51];
              processed = splitDate(tempDat[52]);
              processor = tempDat[53];
            } else {
              sampled = splitDate(tempDat[23]);
              collector = tempDat[24];
              processed = splitDate(tempDat[25]);
              processor = tempDat[26];
            }
            record[3] = sampled[0];
            record[14] = sampled[1];
            record[19] = collector;
            record[42] = processed[0];
            record[43] = processed[1];
            record[44] = processor;
            
            // Set Lot #, Storage Time/Date-DQ, Storage Temp-DQ, and Field Notes
            // Decided to do this manually for sake of specificity
            // Milk does not have a lot number
            switch(record[4]){
              case "Whole Blood":
                tempSrcArr = srcColsWB.slice(0);
                tempDestArr = destColsSamples.slice(0);
                break;
              case "Plasma":
                tempSrcArr = srcColsPL.slice(0);
                tempDestArr = destColsSamples.slice(0);
                break;
              case "Serum":
                tempSrcArr = srcColsSR.slice(0);
                tempDestArr = destColsSamples.slice(0);
                break;
              case "Milk":
                tempSrcArr = srcColsMK.slice(0);
                tempDestArr = destColsSamples.slice(1, 4);
            }
            for (si=0; si<tempSrcArr.length; si++){
              if (tempDestArr[si].length > 1) {
                storeDateTime = splitDate(tempDat[tempSrcArr[si]-1]);
                record[tempDestArr[si][1]-1] = record[3];
                record[tempDestArr[si][0]-1] = storeDateTime[1];
              } else {
                record[tempDestArr[si]-1] = tempDat[tempSrcArr[si]-1];
                if (record[4] == "Milk"){
                  record[35] = null;
                }
              }
            }
            tempSrcArr = [];
            
            // Send record to output array, refresh record, and increment output array
            out[outi] = record.slice(0);
            record = [];
            outi ++;
          }
        }
        // Move to the next tissue
        tissuei ++;
      }
    }
    
    // Build the export
    // Clear it if necessary
    if (buildAll & destLR > 1) {
      dest.getRange(2, 1, destLR-2, out[0].length).clear();
      dest.getRange(2, 1, out.length, out[0].length).setValues(out);
    } else {
      dest.getRange(destLR+1, 1, out.length, out[0].length).setValues(out);
    }
  }
  // Start build 20180205
  // Tested single 20180209:1500 - pass
  // Tested multiple entry 20180212:1215 - pass
}

function splitDate(dateToSplit){
  // Splits a date object and returns an array with readable date (e.g. 10/20/2018) and 24-hr time (e.g. 13:12:11)
  var out = [];
  out.push(dateToSplit.getMonth()+1 + '/' + dateToSplit.getDate() + '/' + dateToSplit.getFullYear());
  out.push(dateToSplit.getHours() + ":" + dateToSplit.getMinutes() + ":" + dateToSplit.getSeconds());
  return (out)
  // Built 20180208:1630
  // Tested 20180208:1645 - pass
}


function getIndices(colName){
  // Grabs the list of indices for column sources and destionation for the column named as colName
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('refResponseHeaders');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var colNames = sheet.getRange(1, 1, 1, lastCol).getValues();
  var colNum = colNames[0].indexOf(colName)+1;
  var srcCol = sheet.getRange(2, colNum, lastRow).getValues();
  var values = [];
  var ii = 0;
  var iii = 0;
  for (i=0; i<srcCol.length; i++){
    if (srcCol[i] != "") {
      var typ = typeof srcCol[i][0];
      if (typeof srcCol[i][0]  === 'string'){
        values[ii] = srcCol[i][0].split(', ');
        for (iii=0; iii<values[ii].length; iii++){
          values[ii][iii] = parseInt(values[ii][iii]);
        }
      } else {
        values[ii] = srcCol[i][0];
      }
      ii++;
    }
  }
  return (values);
  // Built 20180207:1530
  // Tested 20180207:1720 - pass
  // Refined 20180208:1400 - now returning INT for 2d array items rather than string
}

function buildRecordDirect(dat, srcColIndices, destColIndices){
  // Start by building out an array with the direct records
  // This will be repeated for each aliquot entry in a given sample set
  var out = [];
  for (var i=0; i<srcColIndices.length; i++) {
    out[destColIndices[i]-1] = dat[srcColIndices[i]-1];
    if (srcColIndices[i] == 13) {
      out[destColIndices[i]-1] = out[destColIndices[i]-1].replace("Dolphin Quest ", "DQ-");
    }
  }
  return (out);
  // Built 20180208:0830
  // Tested 20180208: 0915 - pass
}


function buildRecordReproduction(out, dat, srcColIndex, destColIndices){  
  // Split a reproductive status answer into two columns
  var status = dat[srcColIndex-1];
  if (status.indexOf("pregnant") != -1) {
    out[destColIndices[0]-1] = "Yes"
  } else {
    out[destColIndices[0]-1] = "No"
  }
  if (status.indexOf("lactating") != -1) {
    out[destColIndices[1]-1] = "Yes"
  } else {
    out[destColIndices[1]-1] = "No"
  }
  return (out);
  // Built 20180208:1345
  // Tested 20180208:1415 - pass
}

function buildRecordPermit(dat, srcColIndices, destColIndex){
  // Build the permit entry in the FW export for Dolphin Quest
  var tmp = [];
  var d = new Date();
  for (var i=0; i<srcColIndices.length-2; i++) {
    tmp[i] = dat[srcColIndices[i]-1];
  }
  for (i=i; i<srcColIndices.length; i++) {
    d = dat[srcColIndices[i]-1];
    tmp[i] = d.getMonth()+1 + '/' + d.getDate() + '/' + d.getFullYear();
  }
  return(tmp.join("-"));
  // Built 20180208:0930
  // Tested 20180208:1000 - pass
  // Modified 20180208:1040 to return only a single string rather than the entire record
  // This is less flexible as it doesn't allow for any additional manipulation of the
  // record, but is more explicitly confined to this function's purpose. - JMR
}

function buildRecordFish(out, dat, srcColIndices, destColIndices){
  // Combine the 7 possible entries for dietary items and parse out into 4 Fish Type, Fish Catch Date, and Kilocalories
  var fishes = ["capelin", "herring", "mackerel", "mullet", "squid", "sardines", "other"];
  var tempDate = null;
  var i = 0, ii = 0, oi = 0, fi = -1;
  var fish = '',
      nfish = 0;
  for (i=0; i<srcColIndices.length; i+=2) {
    fi++;
    tempDate = null;
    fish = fishes[fi];
    if (fish == "other") {
      fish = dat[srcColIndices[i][0]-1];
    } else {
      fish = fishes[fi];
    }
    if (dat[srcColIndices[i][0]-1] != "No" & dat[srcColIndices[i][0]-1] != "") {
      out[destColIndices[oi][0]-1] = fish;
      for (ii=1; ii<srcColIndices[i].length; ii++) {
        out[destColIndices[oi][ii]-1] = dat[srcColIndices[i][ii]-1];
      }
      tempDate = dat[srcColIndices[i][1]];
      // Build the fish lot number string
      if (typeof tempDate != "string") {
        out[destColIndices[oi+1]-1] = fish + "-" + 
          tempDate.getFullYear() + 
            tempDate.getMonth()+1 + 
              tempDate.getDate() + "-" + 
                dat[srcColIndices[i][2]];
      } else {
        out[destColIndices[oi+1]-1] = null;
      }
      oi+=2;
      nfish++;
    }
  }
  if (nfish<4) {
    for (oi=6; oi>=(nfish*2); oi-=2){
      i-=2;
      for (ii=0; ii<srcColIndices[i].length; ii++) {
        out[destColIndices[oi][ii]-1] = null;
      }
      out[destColIndices[oi+1]-1] = null;
    }
  }
  return (out);
  // Started 20180208:1630
  // Built 20180209:0800
  // Tested 20180209:1145 - pass
  // Modified 20180212:0900 - removes "undefined" entries in favor of null
}