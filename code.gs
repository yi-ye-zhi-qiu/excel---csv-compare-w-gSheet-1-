//getIdFromUrl will be used in doGet(e) to return the DriveApp file ID of the wanted Excel file in the folder
function getIdFromUrl(url) { return url.match(/[-\w]{25,}/); }

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Send e-mail')
    .addItem('Send e-mail comparing excel data to this gSheet, informing of discrepancies', 'on_button_click')
    .addToUi();
}

function on_button_click(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    'Please confirm',
    'Are you sure you want to send an email?',
    ui.ButtonSet.YES_NO);
    if (result == ui.Button.YES) {
      send_email();
      ui.alert('OK, done! This does NOT change tracker data.');
    } else {
      ui.alert('ok maybe next time...')
    }
}

//this is just if you want to deploy as a webapp...
function doGet(e) {
  send_email();
  return ContentService.createTextOutput(JSON.stringify({"message":"works!"}))
                       .setMimeType(ContentService.MimeType.JSON);
}

/**Helper function**/
function find_this_file(filename) {
  /**searches gDrive and finds a file by name, will push into an array**/
  var files = DriveApp.getFilesByName(filename);
  var result = [];
  while(files.hasNext())
    result.push(files.next());
  return result;
}


function send_email() {
  /**
  If there's an Excel in your gDrive with 2020-mm-dd in its name,
  this create a csv file of that excel and drop it in the same
  gDrive folder.

  It will then compare to some tracker data w/ the csv,
  and send out a formatted gMail with any inconsistencies in the data.
  **/

  //fldr is DriveApp opening up where we drop all the excel reports
  var fldr=DriveApp.getFolderById(""); //read more: https://developers.google.com/apps-script/reference/drive/drive-app
  //yummy files!
  var files=fldr.getFiles();
  //define today's date, d, as yyyy-mm-dd

  const offset = new Date().getTimezoneOffset()
  const change_d = new Date()
  const d = (new Date(change_d.getTime() - (offset*60*1000))).toISOString().split('T')[0]
  Logger.log("Today's date, which is in the filename of the Excel we want, is: " +d)


  /**
  while there are any files in the folder, cycle through until
  we hit gold, we hit a file that matchs today's date!
  when we hit that file, make it into target_file, or a csv
  **/
  while (files.hasNext()) { //this is what's called a file iterator!
      var file = files.next(), //notice we now singular, not plural! this is just one file at a time :)
      fn = file.getName()
      if(fn.indexOf(d) >-1){ //d can very well be not 2020-mm-dd, and instead "apples report" or something like that..

        //we will be using UrlFetchApp to make requests from Google servers to use one of their apis,
        //read more: https://developers.google.com/drive/api/v2/reference
        //we just need them to trust us, so we use this token
        var token_pls_trust_me_google = ScriptApp.getOAuthToken();

        //First, let's fetch our Excel's (application/vnd.ms-excel confirms this is our filetype, doesn't work with excel 2003) byte data! (＾Ｕ＾)ノ
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
    var file_list = find_this_file(csvfile_name)
    var csvfile = file_list[0]

    //now we have our csv file, so we can parse it!

    //c = csv file contents
    var c = Utilities.parseCsv(csvfile.getBlob().getDataAsString())
    Logger.log("You're running this for " + (c.length - 1) + " rows of data.")

    var JSON_of_p = {},
        p = {};

    for(var i=1; i<c.length; i++){
      var e_p = {}

      /**Let's make a JSON like
      JSON_of_p =
      "A1": {
        "header_for_column_A": "A1",
        "header_for_column_B": "B1",
        "header_for_column_C": "C1"
      },
      "A2": {
        "header_for_column_A": "A2",
        "header_for_column_B": "B2",
        "header_for_column_C": "C1"
      }

      For the purposes of this example we will use the below:
      JSON_of_p =
      "Apple": {
        "amount sold": "100",
        "amount in warehouse": "1000000",
        "last shipment": "10/03/2020"
      },
      "Berries": {
        "amount sold": "0",
        "amount in warehouse": "10",
        "last shipment": "10/02/2020"
      }
      **/

      //This is like the header_for_column A: A1 part
      e_p[c[0][1]] = c[i][1]
      //This is like the header_for_column B: B1 part
      e_p[c[0][2]] = c[i][2]
      e_p[c[0][3]] = c[i][3]

      //This is the "A1": {} part
      p[c[i][0]] = e_p

      //This is adding "A1", "A2" into JSON_of_p, {}
      Object.assign(JSON_of_p, p)
    }

    Logger.log("This if JSON_of_p, it should include all protocol data from the excel->CSV report: \n"+JSON.stringify(JSON_of_p, 2, 2))
    var keys = Object.keys(JSON_of_p)
    Logger.log('keys of JSON_of_p, just letting you know: '+keys)

    //^ Takes about 1-5 seconds to run all that up there. Gives us a nested JSON for FPI planned/actual, LPO planned/actual for every single molecule

    /**
    Now we can compare to our gSheet

    We will loop through each Sheet of the tracker whose sheet name starts with "apple", for example

    If there's a discrepancy, we will store it in an array which we lastly send as an email.
    **/

    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

    //Loop through all tabs that start w/ apple'

    var emaildata = [['fruit', 'amount sold', 'amount remaining', 'last shipment']];

    for (var i=0 ; i<sheets.length ; i++) if(sheets[i].getName().indexOf('check my name!')>-1){

      let this_tab = sheets[i].getDataRange(),
          this_last_row = this_tab.getLastRow(),
          this_tab_data = this_tab.getValues();

      for (var j=1; j<this_last_row; j++){

        let this_fruit = this_tab_data[j][0],
            fruit_amount_sold = this_tab_data[j][1],
            fruit_amount_in_warehouse = this_tab_data[j][2],
            fruit_last_shipment = this_tab_data[j][3]

        for (var k=0; k<keys.length; k++){
          var this_key = keys[k],
          JSON_amount_sold = JSON_of_p[this_key]["amount sold"],
          JSON_amount_remaining = JSON_of_p[this_key]["amount in warehouse"],
          JSON_last_shipment = JSON_of_p[this_key]["last shipment"]

          if(this_key === this_fruit){

            //if non-empty
            if(fruit_amount_sold !== "" || fruit_amount_in_warehouse !== "" || fruit_last_shipment !== ""){

              //format our gSheet dates into mm/dd/yyyy (same as csv->Excel format)
              //note: for some reason GoogleSheets date data has to be concat into string
               function f_d(d){
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
                  Logger.log(formatted_s)
                  d = formatted_s
                  return d
               }

             fruit_last_shipment = f_d(fruit_last_shipment);

             Logger.log('this key is: '+this_key);
             Logger.log('this fruit is: '+this_fruit)
             Logger.log("yup, "+this_fruit +" from the gSheet is the same as " + this_key + " from the JSON!")
             Logger.log('Excel data: \n' + 'excel fruit sold: '+JSON_amount_sold + "\n" + 'excel fruit amount in warehouse: ' +JSON_amount_remaining + "\n" + 'excel fruit last shipment: '+JSON_last_shipment);

            function arr_indexing(a, b, c){
              //want to check if any element in a has element[0] = b[0]
              const one_equal = arr => arr.some(function(e) {return e[0] === b[0]})
              if(one_equal(a) === true){
                a.forEach(function(element) {
                  if(element[0] === b[0]){
                    element.splice(c, 1, b[c])
                  }
                })
              }
              else {
                a.push(b)
              }
              return a
            }


            var check_fruit_amount = (fruit_amount_sold === JSON_amount_sold)
            if(check_fruit_amount === false){
              Logger.log(['amount of fruit sold for ' + this_fruit + ' needs to be updated to: ' + JSON_amount_sold])
              
              let insert = [this_fruit, JSON_amount_sold, null, null]
              Logger.log(insert)
              arr_indexing(emaildata, insert, 1);
              
            }

            var check_fruit_warehouse = (fruit_amount_in_warehouse === JSON_amount_remaining)
            if(check_fruit_warehouse === false){
              Logger.log(['amount in warehouse for ' + this_fruit + ' needs to be updated to: ' + JSON_amount_remaining])
              
              let insert = [this_fruit, null, JSON_amount_remaining, null]
              Logger.log(insert)
              arr_indexing(emaildata, insert, 2)
            }

            var check_fruit_shipment_date = (fruit_last_shipment === JSON_last_shipment)
            if(check_fruit_shipment_date === false){
              Logger.log(['last shipment date for ' + this_fruit + ' needs to be updated to: ' + JSON_last_shipment])
              
              let insert = [this_fruit, null, null, JSON_last_shipment]
              Logger.log(insert)
              arr_indexing(emaildata, insert, 3);         
              
            }

           }
          }
         }
        }
       }

    Logger.log(emaildata)

    //The fun part, setting up the html table from emaildata array we will use in our gMail MailApp component

    var perrow = 4
    var TABLEFORMAT = '"font-family:arial, sans-serif;border-collapse:collapse;"'
    var THFORMAT = 'style="padding-top:10px;padding-bottom:20px;padding-right:15px;padding-left:15px;text-align:left;font-weight:200;font-size:12px;border-bottom-width:5px;border-bottom-style:solid;border-bottom-color:#42A5F5; background-color: #4FC3F7"'
    var TRTDFORMAT = 'style="padding-top:5px;padding-bottom:5px;padding-right:5px;padding-left:5px;text-align:left;vertical-align:middle;font-weight:300;font-size:12px;"'
    //have to use inline-css for gMail API htmlBody to work properly

    var html = '<table ' + TABLEFORMAT + '><th ' + THFORMAT + '></th>';
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
}
