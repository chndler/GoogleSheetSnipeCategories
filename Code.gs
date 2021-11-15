//SETUP
serverURL = 'SERVER_URL';
apiKey = 'API_KEY_GOES_HERE'
var API_count = 0

function onOpen(e) {
  createCommandsMenu();
}

//Create a menu option on the sheet to run the runGetModelsByCategory function
function createCommandsMenu() {
  var ui = SpreadsheetApp.getUi();
      ui.createMenu('Snipe')
      .addItem('List Assets By Category', 'runGetModelsByCategory')
      .addToUi();
}


function setupHeaders(categorySheet){
    categorySheet.appendRow(['Test Kit Name', 'Status'])
    categorySheet.setFrozenRows(1)
}

function getModelInfoByCategoryIDs(categoryID){
    console.log('Category ID: ' + categoryID)
    var url = serverURL + 'api/v1/hardware?sort=model&order=asc&category_id=' + categoryID;
    var headers = {
        "Authorization" : "Bearer " + apiKey
    };
    
    var options = {
        "method" : "GET",
        "contentType" : "application/json",
        "headers" : headers
    };
    console.log("Checking for Model ID Info (getting limit)...")
    var response = JSON.parse(UrlFetchApp.fetch(url, options));
    API_count += 1;
    var limit = response.total
    
    // now run again with limit
    var url = serverURL + 'api/v1/hardware?sort=model&order=asc&category_id=' + categoryID + "&limit=" + limit;
    console.log("Re-querying for Model ID Info (with limit)...")
    var response = JSON.parse(UrlFetchApp.fetch(url, options));
    API_count += 1;
    var rows = response.rows;
    
    var prunedRows = [];
    // since rows includes all assets for a given category, there will be duplicate models, this loop
    // removes any extras so there's only one asset listed per model 
    for (var i=0; i<response.total; i++){
        row = rows[i] 
        // console.log("Asset ID: " + row.id + ", Model Name: " + row.model.name);   
        if ((i > 0) && (row.model.name != rows[i-1].model.name)){
            prunedRows.push(row)
        }
    }
    for (var i=0; i<prunedRows.length; i++){
        row = prunedRows[i] 
        console.log("Asset ID: " + row.id + ", Model Name: " + row.model.name);   
    }
    var modelInfo = prunedRows;
    return modelInfo;
}

function getCategoryID(category){
    var url = serverURL + 'api/v1/categories?search=' + category + '&sort=name&order=asc&=asc';
    var headers = {
        "Authorization" : "Bearer " + apiKey
    };
    
    var options = {
        "method" : "GET",
        "contentType" : "application/json",
        "headers" : headers
    };

    console.log("Making query for Category ID for " + category + "...")
    var response = JSON.parse(UrlFetchApp.fetch(url, options));
    API_count += 1;
    // response _should_ be single rowed
    var limit = response.rows[0].id
    return limit
}

function getModelsByCategory(category){
    // get category ID using category name, get respective sheet, clear it, and setup headers
    var categoryID = getCategoryID(category);
    var ss = SpreadsheetApp.getActive();
    var categorySheet = ss.getSheetByName(category)
    categorySheet.clear()
    setupHeaders(categorySheet);

    // now get all hardware info in 1 function call with categoryIDs
    // getModelInfoByCategoryID() queries the API once instead of N times    
    var modelsArr = getModelInfoByCategoryIDs(categoryID)
    console.log("Appending rows to sheet " + category + "...")
    // categorySheet.setValues(modelsArr[][][])
    for (var k=0; k<modelsArr.length; k++){
        var row = modelsArr[k]
        categorySheet.appendRow([row.model.name, row.status_label.name])
        // categorySheet.appendRow([row.model.name, row.status_label.name, '=IMAGE(\"' + row.image + '\", 2)'])
    }
    console.log("Finished appending")
}

function runGetModelsByCategory(){
    var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var categories = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories').getRange("A1:A").getValues();
    categoriesLast = categories.filter(String).length;
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories')
    var range = sh.getRange('A1:A');
    var values = range.getValues();
    var categories = [];
    for (var i=0; i<categoriesLast; i++){
      var category = values[i];
      console.log("Category added: " + category)
      categories.push(category);
    }

    for (var j=0; j<categoriesLast; j++){
      var category = categories[j];
      getModelsByCategory(category)
    }
    console.log("Total API calls: " + API_count)
  }
