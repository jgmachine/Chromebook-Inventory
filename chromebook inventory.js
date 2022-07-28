var masterList = SpreadsheetApp.openById("YOUR SPREADSHEET ID HERE");
var inventorySheet = masterList.getSheetByName("Sheet1");


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Inventory')
      .addItem('Update Sheet', 'updateSheet')
      .addItem('Update Admin Console', 'updateAdminConsole')
      .addToUi();
}


function updateSheet(){
  var page; //Page object for the API request. Each page by default contains 100 results maximum
  var pageToken; //Token to determine what page to fetch from admin
  var pages = 0; //Counter to keep track of pages queried from Admin
  
  inventorySheet.getRange("A2:P").clear(); //clear admin data sheet so we can insert new data starts at A2 because headers
  
  do {
    //Try catch block to catch any errors in the request to Google Admin. If it fails we will stop requesting altogether
    try {
      Logger.log("Getting Page:" + (pages + 1));
      page = AdminDirectory.Chromeosdevices.list("my_customer", {maxResults: 100, orderBy: 'serialNumber', pageToken: pageToken}); //Get next page of Chrome devices
      pages++; //increment page counter
    }
    
    catch (e){
      Logger.log("Failed at Try Catch: " + e.message); //Print error to stack driver when there is a failure
      break; //Stop the loop
    }
    
    Logger.log("Processing Page:" + pages);
    var pageData = new Array(page.chromeosdevices.length); //Empty array for chromeosdevice data
    
    //Do stuff with each device in the page
    for(var i in page.chromeosdevices){
      var device = page.chromeosdevices[i]; //Next chromeosdevice
      var recentUsers = "None";
      var activeTimes = "None";
      
      //check if there are recent users. If so, set that value. Otherwise they show up as none
      //If last user shows up as undefined, it means that that user is not msdlt managed.
      if(device.recentUsers){
        recentUsers = device.recentUsers[0].email; //set most recent user
      }
      
      //check for activeTimes. If so, set that value. Otherwise it shows as none
      if(device.activeTimeRanges){
        activeTimes =  device.activeTimeRanges[0].date + " - " + Math.ceil(device.activeTimeRanges[0].activeTime/1000/60) + " minutes"; //format and set last active time
      }
      
      //chromeosdevice data is put into an array
      var data = [device.orgUnitPath,
                  device.notes,
                  device.annotatedLocation,
                  device.annotatedAssetId,
                  device.annotatedUser,
                  device.serialNumber,
                  device.status,
                  recentUsers,
                  activeTimes,
                  device.lastSync,
                  device.lastEnrollmentTime,
                  device.deviceId,
                  device.model,
                  device.platformVersion,
                  device.firmwareVersion,
                  device.supportEndDate
                 ];
      pageData[i] = data; //put chromeosdevice data into pageData array
    }
    
    var lastRow = inventorySheet.getLastRow() + 1; //get last empty row in the admin data sheet
    inventorySheet.getRange("A" + lastRow + ":P" + (lastRow + pageData.length - 1)).setValues(pageData); //copy over pagedata array to sheet.
    
    inventorySheet.sort(2) //sort the sheet by the 2nd colum (Notes).
    
    /*Note: Storing devices in an array and writing the entire array is much faster than writing each individual device line by line as each time "Sheet.setValues()
    * is called, communications happens with the Google sheets servers or whatever which really slows things down and isn't really good practice.
    */
    //break; //for testing purposes
    pageToken = page.nextPageToken; //get token of next page for next iteration
    Utilities.sleep(Math.random() * 1000); //wait between 0 and 1 seconds (in ms) until the next iteration. Don't want google to get mad at us for querying their servers for 23k devices all at once :)
  }
  while(pageToken); //Continue the loop while we have a valid pageToken
}


function updateAdminConsole(){
  var deviceId = inventorySheet.getRange("L2").getValue();
  Logger.log(deviceId);
  
  var device = AdminDirectory.Chromeosdevices.get('my_customer', deviceId)
  
  AdminDirectory.Chromeosdevices.update(device, 'my_customer', deviceId)
  
  var end = inventorySheet.getLastRow();
  
  if (inventorySheet.getLastRow()>1) {
    for (var i=2; i<end+1; i++) {
      var OU = inventorySheet.getRange(i,1).getDisplayValue();
      var notes = inventorySheet.getRange(i,2).getDisplayValue();
      var annotatedLocation = inventorySheet.getRange(i,3).getDisplayValue();
      var annotatedAssetId = inventorySheet.getRange(i,4).getDisplayValue();
      var annotatedUser = inventorySheet.getRange(i,5).getDisplayValue();
      var serialNumber = inventorySheet.getRange(i,6).getDisplayValue();
      
      
      var optionalArgs =  {
        "orgUnitPath": OU,
        "annotatedUser": annotatedUser,
        "annotatedAssetId": annotatedAssetId,
        "annotatedLocation": annotatedLocation,
        "notes": notes
      };
      var listChromebook = (AdminDirectory.Chromeosdevices.list("my_customer",{query: 'id:' + serialNumber}));
      var updatedChromebook =(AdminDirectory.Chromeosdevices.update(optionalArgs, "my_customer", listChromebook.chromeosdevices[0].deviceId)); // Update admin console with new values
      
    }
  }
}