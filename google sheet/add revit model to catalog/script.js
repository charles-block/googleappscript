/*How to Use


1. Set spreadsheetId with id of Catalog Sheet
2. Set sheetName with name of sheet on Catalog Sheet

3. Set spreadsheetId2 with id of Revit Catalog Sheet
4. Set sheetName2 with name of sheet on Revit Catalog Sheet

5. Set catvalue with SUB_CATEGORY


Note: 
- Copy Code To Apps Script Console Before Use
- Do Not Add or Move Columns
*/


var catvalue = "Tile - Floor";

//DSE Material Catalog
var spreadsheetId = "1LJOgByiatqIikzUWEep7kxhuTJqtfSuhoXzm9keY6HA";
var sheetName = "materials_pull__snflk__2023-05-11T16_47_44.903157Z";

//Revit Model Catalog
var spreadsheetId2 = "11ZnK70xcF72Ivi1xF6bCAgfxLc_dvgwMJV4Wrwssyzw";
var sheetName2 = "Bathroom";

function removeCharactersFromString(string) {
  var charactersToRemove = [":","(",")","<",/"/g];
  var regexPattern = new RegExp('[' + charactersToRemove.join('') + ']', 'g');
  var result = string.replace(regexPattern, '');
  return result;
}


function compareLists(list1, list2, subcategorylist) {
  var result = [];

  for (var i = 0; i < list1.length; i++) {
    var value = list1[i].toString();
    var found = false;
    var subcatvalue = removeCharactersFromString(subcategorylist[i].toString());

    for (var j = 0; j < list2.length; j++) {
      if(list2[j].includes(value)&&list2[j].includes(catvalue)&&subcatvalue === catvalue){
      //if (list2[j] === value) {
        found = true;
        result.push(j);
        break;
      }
    }

    if (!found) {
      result.push(null);
    }
  }
  return result;
}

function logData(data, chunkSize) {
  var startIndex = 0;
  var endIndex = chunkSize;
  var iteration = 1;
  
  while (startIndex < data.length) {
    var chunk = data.slice(startIndex, endIndex);
    Logger.log("Chunk " + iteration + ": " + chunk);
    
    startIndex = endIndex;
    endIndex = Math.min(endIndex + chunkSize, data.length);
    iteration++;
  }
}

function getColumnBValues() {


  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var columnBValues = sheet.getRange("B:B").getValues().flat();
  var subcategoryvalues = sheet.getRange("I:I").getValues().flat();
  
  var sheet2 = SpreadsheetApp.openById(spreadsheetId2).getSheetByName(sheetName2);
  var columnAValues2 = sheet2.getRange("A:A").getValues().flat();
  var drivelinkValue = sheet2.getRange("B:B").getValues().flat();

  var indices = compareLists(columnBValues, columnAValues2, subcategoryvalues);
  console.log(indices.length)

  for (var k = 0; k < indices.length; k++) {
    console.log(k, indices[k])
    if (indices[k] == null){continue}
    else
    {
      var drivelinkaddress = drivelinkValue[indices[k]];
      var rownum = k+1
      var drivelinkcell = sheet.getRange("AD" + rownum);
      drivelinkcell.setValue(drivelinkaddress);

      var modelstatuscell = sheet.getRange("AC" + rownum);
      modelstatuscell.setValue("Existing");

    }

  }
  //logData(indices, 100)
  //console.log(indices);
}


//AC = Exist
//AD = Revit Model Link

