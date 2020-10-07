function driveBackup(){ 
  var formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd' 'HH:mm:ss");
  var name = SpreadsheetApp.getActiveSpreadsheet().getName() + " Copy " + formattedDate;
  var cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel').getRange('J20').getValue();
  var destination = DriveApp.getFolderById(cell);
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
  var s = file.makeCopy(name, destination); 
  var s = s.getId();
  var ss = SpreadsheetApp.openById(s);
  var sheet = SpreadsheetApp.setActiveSpreadsheet(ss);
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  // Unlink and delete all forms that are linked to spreadsheet sh
  sheet.getSheets().forEach(function(sheet){
  var formUrl = sheet.getFormUrl();  // returns null if there is no linked form 
  if (formUrl) {           
    var form = FormApp.openByUrl(formUrl);
    form.removeDestination();
    DriveApp.getFileById(form.getId()).setTrashed(true);
  }
})
  addToHistory("Backup Completed");
}
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'New Student Record', functionName: 'newRecord'},
    {name: 'Backup', functionName: 'driveBackup'},
    {name: 'Reset', functionName: 'reset'},
    {name: 'Update Work', functionName: 'addDataToSheets'},
  ];
  spreadsheet.addMenu('Scripts', menuItems);
  driveBackup();
  allSandJ();
  reset(0);  
}

function allSandJ(){
   var seniors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel').getRange(26,1,15,2);
   var juniors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel').getRange(43,1,15,2);
   var sen_values = seniors.getValues();
   var jun_values = juniors.getValues(); 
   var arr_sen={};
   var arr_jun={};
   for (var i=0;i<sen_values.length;++i){
      if(sen_values[i][0]!='')
      {arr_sen[i]=[sen_values[i][0],sen_values[i][1]];}
       else
       {break;}
    }
   for (var i=0;i<jun_values.length;++i){
      if(jun_values[i][0]!='')
      {arr_jun[i]=[jun_values[i][0],jun_values[i][1]];}
     else{break;}}
    arr_sen =  JSON.stringify(arr_sen);
    arr_jun =  JSON.stringify(arr_jun)
    PropertiesService.getScriptProperties().setProperty("AllSeniors", arr_sen);  
    PropertiesService.getScriptProperties().setProperty("AllJuniors", arr_jun);
}


//Adds to individual sheets
function updateSheet(row){
  var cur_row = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(row[2]);
  cur_row.appendRow([row[4],row[5],row[6]]);
  if(row[7]!=''){
  cur_row.appendRow([row[4],row[7],row[8]]);
  }
  if(row[9]!=''){
  cur_row.appendRow([row[4],row[9],row[10]]);
  } 
  var str = "Added data to "+ row[2];
  addToHistory(str);
}
//Get data of people who didn't work and people who didn't work
function workdata(){
  var arr_jun_w=[];
  var arr_jun_nw=[];
  var arr_sen_nw=[]; 
  var arr_sen_w=[];
  var list_j= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel').getRange(6,2,15,1);
  var list_s= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel').getRange(6,1,15,1);
  for (var i = 1; i <= list_j.getNumRows(); i++) {
      j=1;
      var cell = list_j.getCell(i, j);
      if(cell.getBackground().toUpperCase() == '#E06666' )
        arr_jun_nw.push(cell.getValue());
      else
        arr_jun_w.push(cell.getValue());
  }
  for (var i = 1; i <= list_s.getNumRows(); i++) {
      j=1;
      var cell = list_s.getCell(i, j);
      if(cell.getBackground().toUpperCase() == '#E06666' )
        arr_sen_nw.push(cell.getValue());
      else
        arr_sen_w.push(cell.getValue());
  }
  arr_sen_nw=arr_sen_nw.filter(Boolean);
  arr_jun_nw=arr_jun_nw.filter(Boolean);
  return [arr_sen_nw,arr_jun_nw,arr_sen_w,arr_jun_w]
}

//Update worklist
function updateWorkList(arr_s_w,arr_j_w){//ID from new data
  var values = workdata();
  arr_sen_w = values[2];
  arr_jun_w = values[3]; //From existing data
  var arr_jun=JSON.parse(PropertiesService.getScriptProperties().getProperty("AllJuniors"));
  var arr_sen=JSON.parse(PropertiesService.getScriptProperties().getProperty("AllSeniors"));
  var count_jun = Object.keys(arr_jun).length;
  var count_sen = Object.keys(arr_sen).length;
  var names_jun=[];
  var names_sen=[];
  for(var i=0;i<arr_s_w.length;++i)
  {for(var j=0;j<count_sen;++j)
  { if(arr_s_w[i]===arr_sen[j][0])
    {names_sen.push(arr_sen[j][1])
     break;}
  }
  }
  for(var i=0;i<arr_j_w.length;++i)
  {for(var j=0;j<count_jun;++j)
  { if(arr_j_w[i]===arr_jun[j][0])
    {names_jun.push(arr_jun[j][1])
     break;}
  }  
  }   
   temp_names_j= names_jun.filter(x => arr_jun_w.indexOf(x)===-1);
   temp_names_s= names_sen.filter(x => arr_sen_w.indexOf(x)===-1);
   names_jun=temp_names_j;
   names_sen=temp_names_s;
  //Update color
  var list_j= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel').getRange(6,2,15,1);
  var list_s= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel').getRange(6,1,15,1);
  for(var k = 0; k<names_jun.length;++k){
    Logger.log(k);
    for (var i = 0; i < list_j.getNumRows(); i++) {
      j=1;
      var cell = list_j.getCell(i+1,j).getValue();
      if (cell=='')
        continue;
      if (names_jun[k].indexOf(cell) != -1)
      {list_j.getCell(i+1,j).setBackground("#93C47D")
      break;}
    }
  }
  for(var k = 0; k<names_sen.length;++k){
    for (var i = 0; i < list_s.getNumRows(); i++) {
      j=1;
      var cell = list_s.getCell(i+1,j).getValue();
      if (cell=='')
        continue;
      if (names_sen[k].indexOf(cell) != -1)
      {list_s.getCell(i+1,j).setBackground("#93C47D")
      break;}
    }
  }
  addToHistory("Updated main list of people who worked and didn't work.");
}

//Adds to history sheets
function addToHistory(text){
   var history_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("History");
   var today = new Date();
   var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
   var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds(); 
   var dateTime = date+' '+time;
   history_sheet.insertRowBefore(5);
   history_sheet.getRange(5,1,1,1).setValue(dateTime);
   history_sheet.getRange(5,2,1,1).setValue(text);
}
//Resets mainsheet
function reset(manual=1){ 
  if (manual==1)
  {var confirmation=Browser.msgBox('Warning:Reset will clear all the data on the Control Panel. Click Ok to continue', Browser.Buttons.OK_CANCEL);
   if (confirmation=="cancel"){return null;}
   }

    var today = new Date();
    const monthNames = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"
                     ];
    var month = monthNames[today.getMonth()];
    var year  = today.getFullYear();
    var date  = month+", "+year;
    var sheet =SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel');
    var currentdate = sheet.getRange(4,1).getValue();
    Logger.log(currentdate)
    Logger.log(date)
    if (!(currentdate==date))
    {sheet.getRange(4,1).setValue(date);
     if(manual==0) {addDataToAllMonths()}}
    else
      if (manual==0)return null;
    var clear_range =sheet.getRange(6,1,15,2).clearContent();
    var values=allSandJ();
    var arr_jun=JSON.parse(PropertiesService.getScriptProperties().getProperty("AllJuniors"));
    var arr_sen=JSON.parse(PropertiesService.getScriptProperties().getProperty("AllSeniors"));
    var count_jun = Object.keys(arr_jun).length-1;
    var count_sen = Object.keys(arr_sen).length-1;
    var mainlist;
    for (var i=0;i<=count_jun;++i)
    {
       mainlist= sheet.getRange(i+6,2)
       mainlist.setValue(arr_jun[i][1]);
       mainlist.setBackgroundRGB(224,102, 102);
    } 
    for (var i=0;i<=count_sen;++i)
    {
       mainlist= sheet.getRange(i+6,1)
       mainlist.setValue(arr_sen[i][1]);
       mainlist.setBackgroundRGB(224,102, 102);
    } 
       addToHistory('Reset Completed');
  return null;
}

function addDataToAllMonths(){
    const monthNames = ["January", "February", "March", "April", "May", "June",
                      "July", "August", "September", "October", "November", "December"
                     ];
    var today = new Date();
    var month = monthNames[today.getMonth()-1];
    var year  = today.getFullYear();
    var date  = month+", "+year;
    var sheet= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All Months');
    var lastrow = sheet.getLastRow();
    if (lastrow ===0) 
      lastrow=-1;
    sheet.getRange(lastrow+2,1,1,10).mergeAcross().setBackground('#FF9900').setValue(date).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(lastrow+3,1,1,2).mergeAcross().setValue('People who worked').setFontWeight('bold');
    sheet.getRange(lastrow+3,4,1,2).mergeAcross().setValue('People who didn\'t worked').setFontWeight('bold');
    sheet.getRange(lastrow+4,1,1,1).setValue('Name').setFontWeight('bold');
    sheet.getRange(lastrow+4,4,1,1).setValue('Name').setFontWeight('bold');
    var values=workdata();
    var worked= values[3].concat(values[2]);
    var notworked = values[1].concat(values[0]);  
    Logger.log(notworked.length);
    Logger.log(notworked);
    for (var i=0;i< worked.length;++i)
    {
      Logger.log(worked[i]);
    }
    for (var i=0;i< worked.length;++i)
    {
      sheet.getRange(lastrow+5+i,1).setValue(worked[i]);
    }
  for (var i =0;i< notworked.length;++i){
      sheet.getRange(lastrow+5+i,4).setValue(notworked[i]);
    }
}



//Adds to each sheet.
function addDataToSheets() {
   var confirmation=Browser.msgBox('Click Ok to run the sheet', Browser.Buttons.OK_CANCEL);
   if (confirmation=="cancel"){return null;}
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var arr_j_w=[];
  var arr_s_w=[];
  var cur_range = 0;
  for (var i=1;i<values.length;++i){
    if(values[i][0]==1){
      values[i][2]=values[i][2].toUpperCase();
      var searchvalue= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel').getRange('L18').getValue();+SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel').getRange('L19').getValue();
      var search= values[i][2].search(searchvalue);
      if(search==-1.0)
      {if (arr_s_w.indexOf(values[i][2])==-1)
        arr_s_w.push(values[i][2]);}
      else
      {if (arr_j_w.indexOf(values[i][2])==-1)
      {arr_j_w.push(values[i][2]);}}
      //Adding to individual sheet
      updateSheet(values[i]);
      //Backup Sheet
      var backup_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Backup of Form Responses');
      backup_sheet.appendRow([values[i][0],values[i][1],values[i][2],values[i][3],values[i][4],values[i][5],values[i][6],values[i][7],values[i][8],values[i][9],values[i][10]]);
      addToHistory('Added data to backup and deleted rows.');
    }
    else 
    {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1').getRange(i+1,1,1,11).setBackgroundRGB(255, 102, 102);
    }  
  }
   for(var i = values.length-1; i >= 0; i--){
    if(values[i][0]==1){
      sheet.deleteRow(i+1); 
    }}
  updateWorkList(arr_s_w,arr_j_w);
}

function getLastRowSpecial(range){
  var rowNum = 0;
  var blank = false;
  for(var row = 0; row < range.length; row++){
    if(range[row][0] === "" && !blank){
      rowNum = row;
      blank = true;
    }else if(range[row][0] !== ""){
      blank = false;
    };
  };
  return rowNum;
};
//function NewRecords
function newrecords(){
  var confirmation=Browser.msgBox('Warning:Cliking Ok will create new sheets.Click Ok to continue', Browser.Buttons.OK_CANCEL);
   if (confirmation=="cancel"){return null;}
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = source.getSheetByName('New Records');
  var lastrow = sheet.getLastRow();
  var range = sheet.getRange(5,1,lastrow-4,7);
  var values = range.getValues();
  var copysheet = source.getSheetByName('Copy-Student Record');
  copysheet.showSheet();
  var controlpanel = source.getSheetByName('Control Panel');
  var juniors = controlpanel.getRange(43,1,15).getValues();
  var seniors = controlpanel.getRange(26,1,15).getValues();
  var seniors_lastRow= 26+getLastRowSpecial(seniors);
  var juniors_lastRow= 43+getLastRowSpecial(juniors);
  Logger.log(seniors_lastRow);
  Logger.log(juniors_lastRow);
  for(var i=0;i<values.length;++i)
  {
     copysheet.copyTo(source).setName(values[i][0]);
     var searchvalue= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel').getRange('L18').getValue();+SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Control Panel').getRange('L19').getValue();
     if (values[i][0].search(searchvalue)==0)
         {
           controlpanel.getRange(juniors_lastRow,1).setValue(values[i][0]);
           controlpanel.getRange(juniors_lastRow,2).setValue(values[i][1]);           
           controlpanel.getRange(juniors_lastRow,3).setValue(values[i][3]);
           ++juniors_lastRow;
         }
     else{
           controlpanel.getRange(seniors_lastRow,1).setValue(values[i][0]);
           controlpanel.getRange(seniors_lastRow,2).setValue(values[i][1]);           
           controlpanel.getRange(seniors_lastRow,3).setValue(values[i][3]);
           ++seniors_lastRow;
     }
     var current_sheet = source.getSheetByName(values[i][0]);
     current_sheet.getRange('B2').setValue(values[i][0]);
     current_sheet.getRange('B3').setValue(values[i][1]);
     current_sheet.getRange('B4').setValue(values[i][2]);
     current_sheet.getRange('B5').setValue(values[i][3]);
     current_sheet.getRange('H3').setValue(values[i][4]);
     current_sheet.getRange('H4').setValue(values[i][5]);
     current_sheet.getRange('H5').setValue(values[i][6]);
  }
  for(var i = values.length; i > 0; i--){
      sheet.deleteRow(i+4); 
    }
  addToHistory('New Records Created');
  copysheet.hideSheet(); 
}



function newrecordsicon(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('New Records');
  sheet.activate();
}

