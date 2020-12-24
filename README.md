# excel -> csv & compare w/ gSheet!!1
## Upload an excel to gDrive. This google script converts it to csv, grabs the data &amp; compares it to what's on a gSheet ##

* Upload Excel to gDrive
* On gSheet, Tools -> Script: run `OnOpen()` :boom:

```javascript

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Send e-mail')
    .addItem('Send e-mail with my excel data compared to my gSheet !_!', 'on_button_click')
    //clicking that will trigger on_button_click
    .addToUi();
}

//Customize what happens when you click the button!

function on_button_click(){
     var ui = SpreadsheetApp.getUi();
     var result = ui.alert(
     'Please confirm',
     'Are you sure you want to send an email?',
      ui.ButtonSet.YES_NO);
      if (result == ui.Button.YES) {
        send_email(); //send_email is our main function
        ui.alert('OK, done!');
      } else {
      }
}

```

* Now click on "Send e-mail" from your gSheet toolbar, and voila you'll send something like this!

![Demo Email](../assets/demo_email.jpg?raw=true)

## Helpful tips/explanation! ##

### Where can I see an example gSheet and Excel? :fallen_leaf: ###
[gSheet!](https://docs.google.com/spreadsheets/d/1cdogBfs6bDuhEgn38_6e2VIxj-OaMBYovXQOo-vMyyY/edit?usp=sharing) 
[excel!](https://drive.google.com/file/d/1qB8GNvQiItjJZWtT8dvQe-hc6VGVaJ4w/view?usp=sharing)

### How can I debug this easily? ###
* Use `Tools -> Script Editor -> View -> Stackdriver Logging -> Apps Script Dashboard` & click on the latest run! Tip: use `Logger.log()` as opposed to `console.log()`!!

### How does the excel get turned into a csv in our drive folder? :sheep: ###
```javascript

 //We do a classic file-loop through our gDrive folder, called 'fldr', where we drop all our excel files >:)
 var fldr=DriveApp.getFolderById("your_folder_id_here!"); //read more: https://developers.google.com/apps-script/reference/drive/drive-app
 //yummy, files!
 var files=fldr.getFiles();
 
 while (files.hasNext()) { //this is what's called a file iterator!
      var file = files.next(), //notice we now singular, not plural! this is just one file at a time :)
      fn = file.getName(),
      d = "this_is_in_my_filename_and_it_has_to_be";
      if(fn.indexOf(d) >-1){ //this is in case you need to check which file you want in particular
 
        //we will be using UrlFetchApp to make requests from Google servers to use one of their apis, 
        //read more: https://developers.google.com/drive/api/v2/reference
        //we just need them to trust us, so we use this token
        var token_pls_trust_me_google = ScriptApp.getOAuthToken();
           
        //First, let's fetch our Excel's (application/vnd.ms-excel confirms this is our filetype, don't use excel 2003) byte data! (＾Ｕ＾)ノ
        var filedata = JSON.parse(UrlFetchApp.fetch(
          "https://www.googleapis.com/upload/drive/v2/files?uploadType=media&convert=true", {
            method: "POST",
            muteHttpExceptions: true,
            contentType: "application/vnd.ms-excel",
            payload: file.getBlob().getBytes(),
            headers: {
              "Authorization": "Bearer " + token_pls_trust_me_google
            }
         }
        ).getContentText());
        
        //Now that we have that byte data, let's request Google to give us back a csv file (*^▽^*)
       
        var target_file = UrlFetchApp.fetch(
            filedata.exportLinks["text/csv"], {
              method: "GET",
              headers: {
                "Authorization": "Bearer " + token_pls_trust_me_google
              }
        })
       
       
        //We can make our target_file blobby and save it in our folder as a csv, with whatever name we want! 
        fldr.createFile(target_file.getBlob()).setName(file.getName() + ".csv")
       
      }
      break;
    }
    var csvfile_name = file.getName() + ".csv"
    
    //now we just find that file in our folder again... you can write out a separate function find_this_file, or just don't do that
    //totally up to you...
    
    function find_this_file(filename) {
      var files = DriveApp.getFilesByName(filename);
      var result = [];
      while(files.hasNext())
        result.push(files.next());
      return result;
    }
    
    var file_list = find_this_file(csvfile_name)
    var csvfile = file_list[0]

```

### Want to parse dates on your gSheet? ###

```javascript
//first, for comparison reasons we get today's date as yyyy-mm-dd based off your timezone!

 const offset = new Date().getTimezoneOffset()
 const change_d = new Date()
 const d = (new Date(change_d.getTime() - (offset*60*1000))).toISOString().split('T')[0]
    
//change the date on a gSheet to mm/dd/yyyy for possible comparison with American-date-format excel data! 
//NOTE: for some reason with dates on a gSheet you have to use ''.concat(d) instead of just using d!!!
 function format_date(d){
    var s = ''.concat(d)
    var s_d = s.split(" ");
    var m = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
               
    for(var j=0; j<m.length; j++){
      if(s_d[1]==m[j]){
        s_d[1]=m.indexOf(m[j])+1;
       }
    }
    //make '1' -> '01' for January, for instance, else leave 11 for Dec as 11, not make it 011
    if(s_d[1]<10){
      s_d[1] = '0'+s_d[1];
    }
    //format as mm/dd/yyyy, same as Excel data
    var formatted_s = s_d[1]+'/'+s_d[2]+'/'+s_d[3];
    //Logger.log(formatted_s)
    d = formatted_s
    return d
 }

```
### Want to parse your CSV data? Read it in your function as a JSON! :boom: ###

```javascript
//see how to generate the excel->csv (get var 'c' below) in the send_email.gs file :frog:

//Use utilities.parseCsv, and the file you use must be a blob!
var c = Utilities.parseCsv(csvfile.getBlob().getDataAsString())

//From here, let's say we want a JSON like
let want_JSON = {
                  "'Header column' cell value in CSV, like cell A2 in excel": {
                    "b2 value":  "B2 value!", //like cell A2
                    "b3value":  "B3 value!" //like cell A3, the lack of a space is important later on
                  }
                }
//Then we can use...

let wanted_JSON = {},
    p = {} //p is like header column

for(var i=1; i<c.length; i++){
  var e_p = {} //inside p we will nest e_p
  e_p[c[row-index][column-index]] = c[i][column-index]
  p[c[i][1]] = e_p
  Object.assign(wanted_JSON, p)
}

Logger.log('this is a JSON now! \n' + JSON.stringify(wanted_JSON, 2, 2))
Logger.log('here are the keys! \n' + Object.keys(wanted_JSON)
//Now you can read your CSV file as a JSON!

```
### Want to compare with some gSheet data and send out the inconsistencies as an email? ###
```javascript
inconsistent_data_points = [] //we'll send these out as a gMail later on!

//To compare sheet data, we need to get it first!



var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
//Let's go through each Sheet that has a name with apple in it for no reason!

for (var i=0 ; i<sheets.length ; i++) if(sheets[i].getName().indexOf('apple')>-1){

//this_tab is data range, this_tab_data is its values as array-of-arrays, and lastRow so we can loop w/ j :) 
 let this_tab = sheets[i].getDataRange(),
          this_last_row = this_tab.getLastRow(),
          this_tab_data = this_tab.getValues();
 for (var j=1; j<this_last_row; j++){
   let each_sheet_B_value = this_tab_data[j][1], //this starts @ B2 is because j starts at 1! if we chose 0, we get B1
       each_sheet_A_value = this_tab_data[j][0];
       
   Logger.log('this should be B2 value of each sheet!' +each_sheet_B_value)

   //now compare to your wanted_JSON from above tip!
   for (var k=0; k<keys.length; k++){
      var this_key = keys[k], 
      JSON_header2 = wanted_JSON[this_key]["b2 value"],
      JSON_header3 = wanted_JSON.this_key.b3value
      
      Logger.log('for ' + this_key + ' its JSON value for header #2 is ' + JSON_header2)
      
      if(JSON_header2 !== each_sheet_B_value){
        //do something if it's not equal to your Excel!! I personally want to send it out as an email, 
        //so I will push it into my inconsistent_data_points array...
        Logger.log('On the excel it's saying ' + JSON_header2 + 'but on my gSheet I see ' + each_sheet_B_value + '!')
        
        inconsistent_data_points.push([this_key, JSON_header2, each_sheet_B_value])
        
      }

```
### How to send a gMail email from the data? ###

```javascript
//Let's say you've got an array of inconsistent data points between your gSheet and your Excel called inconsistent_data_points

var inconsistent_data_points = [["Apples", 5, 3], ["Oranges", 10, 2], ["Pineapples", "n/a", "n/a"]]

//Like your Excel is saying you have 5 apples but your gSheet says 3, that you have 10 Oranges 
//but your gSheet says 2, and for pineapples we just don't know..
    var perrow = 3
    var TABLEFORMAT = '"font-family:arial, sans-serif;border-collapse:collapse;"'
    var THFORMAT = 'style="padding-top:10px;padding-bottom:20px;padding-right:15px;padding-left:15px;text-align:left;font-weight:200;font-size:12px;border-bottom-width:5px;border-bottom-style:solid;border-bottom-color:#42A5F5; background-color: #4FC3F7"'
    var TRTDFORMAT = 'style="padding-top:5px;padding-bottom:5px;padding-right:5px;padding-left:5px;text-align:left;vertical-align:middle;font-weight:300;font-size:12px;"'
    //have to use inline-css for gMail API htmlBody to work properly

    var html = '<h2 style="font-size: 12px; font-weight: 200; text-align: left; margin: 10px;">the following is on the excel but different on your gSheet!</h2>'+
    '<table ' + TABLEFORMAT + '><th ' + THFORMAT + '></th>';
      for (var i = 0; i < emaildata.length; i++) {
        var each_row = '<tr ' + TRTDFORMAT + '><td ' +TRTDFORMAT + '>';
        for (var j = 0; j < emaildata[i].length; j++) {
          if(emaildata[i][j] === "n/a"){
            emaildata[i][j] === ""
          }
          else{
            each_row += emaildata[i][j]
          }
          each_row += '</td><td ' + TRTDFORMAT + '>';
        }
        html += '<td ' +TRTDFORMAT + '>'+each_row+'</td>'
        //html += "<td>" + further_beyond + "</td>";
        var next = i+1;
        if (next%perrow==0 && next!=emaildata.length){
          html +="</tr><tr>";
        }
      }
      html += "</table>"

      MailApp.sendEmail({
        to: "send_to@gmail.com",
        subject: "automated message about fruits",
        htmlBody: '<h1 style="padding-top: 50px; margin: 10px; font-size: 30px; font-weight: 300; text-align: left; margin-bottom: -15px;">Fruit Data inconsistencies!</h1>' +
        html
      })
```

