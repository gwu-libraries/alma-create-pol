// Global variables

// Alma API URLs
var polApiUrl = "https://api-na.hosted.exlibrisgroup.com/almaws/v1/acq/po-lines"

// Current (owning) spreadsheet --> Insert your spreadsheet ID here
var spreadsheet = SpreadsheetApp.openById("");

// Keys used for configuration -- should matche the columns on the config tab of the calling (owning) spreadsheet.

var configSheet = spreadsheet.getSheetByName("config");
var rSheet = spreadsheet.getSheetByName(configSheet.getRange(2,2).getValues());
var lSheet = spreadsheet.getSheetByName("locationMapping");
var userSheet = spreadsheet.getSheetByName("authorizedUsers");
var apiKey = "apikey " + configSheet.getRange(2,1).getValues();

function createPol() {
  // Create POL in Alma from form responses that don't have completed date
  // NEED TO ADD step to loop through this function for all rows with no data in the completed column!
  // Currently this is just finding the last row, assuming it is not completed since the function is triggered on form submit.
  var lrow = rSheet.getLastRow();
  var compFirstEmpty  = lrow;

  // Mapping of form fields/columns to Alma API fields
  var title = rSheet.getRange(compFirstEmpty,3).getValue();
  var isbn = rSheet.getRange(compFirstEmpty,4).getValue();
  var location = rSheet.getRange(compFirstEmpty,5).getValue();
  var collectionName = rSheet.getRange(compFirstEmpty,6).getValue();
  var price = rSheet.getRange(compFirstEmpty,7).getValue();
  var qty = rSheet.getRange(compFirstEmpty,8).getValue();
    // make sure qty is an integer; if not, set qty to 1
    if (Number.isInteger(qty)) {}
    else {
      qty = 1;
    }
  var totalPrice = (price*qty);
  var fundCode = rSheet.getRange(compFirstEmpty,9).getValue();
  var vendorCode = rSheet.getRange(compFirstEmpty,10).getValue();
  var reportingCode = rSheet.getRange(compFirstEmpty,11).getValue();
  var vendorOrderNumber = rSheet.getRange(compFirstEmpty,12).getValue();
  var interestedUser = rSheet.getRange(compFirstEmpty,13).getValue();
    // NEED TO ADD check to ensure interested user is valid ID; or move the interested user bit as a separate action to perform after POL is created
  var holdInterested = rSheet.getRange(compFirstEmpty,14).getValue();
  var note = rSheet.getRange(compFirstEmpty,15).getValue();

  // Other cells to pull data from
  var submitterEmail = rSheet.getRange(compFirstEmpty,2).getValue();
  var completed = rSheet.getRange(compFirstEmpty,18).getValue();

  // Cells to enter responses into
  var polCell = rSheet.getRange(compFirstEmpty,16);
  var mmsCell = rSheet.getRange(compFirstEmpty,17);
  var completedCell = rSheet.getRange(compFirstEmpty,18);

  // Confirm row has not already been processed

  if (completed) {
    console.log(completed);
    console.log("Last row of spreadsheet already processed.");
  } else {

    // Confirm user is authorized to create POL
    var authorizedList = userSheet.getRange('A:A').getValues();
    var authorizedListFlat = authorizedList.map(function(row) {return row[0];});
    if (authorizedListFlat.indexOf(submitterEmail) != -1) {
      console.log("authorized");
      // find library based on location - avoids form submitter having to select library in addition to location with each submission
      // based on response here https://webapps.stackexchange.com/questions/123670/is-there-a-way-to-emulate-vlookup-in-google-script
      var data = lSheet.getRange('A2:B').getValues()
      var searchValue = location;
      var dataList = data.map(x => x[0])
      var index = dataList.indexOf(searchValue);
      if (index === -1) {
        var library = "gelman";
      } else {
          var library = data[index][1]
      }

      // Creating the REST API json string for submission
      var payloadMain = {
        "owner": {
          "value": "gelman"
        },
        "type": {
          "value": "PRINTED_BOOK_OT"
        },
        "vendor": {
          "value": vendorCode
        },
        "vendor_account": vendorCode,
        "price": {
          "sum": price
        },
        "vendor_reference_number": vendorOrderNumber,
        "resource_metadata": {
          "title": title,
          "author": collectionName,
          "isbn": isbn,
          "publisher": note
        },
        "fund_distribution": [
          {
            "fund_code": {
              "desc": "string",
              "value": fundCode
            },
            "amount": {
              "sum": totalPrice,
              "currency": {
                "value": "USD"
              }
            }
          }
        ],
        "reporting_code": reportingCode,
        "note": [
          {
            "note_text": "Created by " + submitterEmail + " via the POL creation request form."
          }
        ],
        "location": [
          {
            "quantity": qty,
            "library": {
              "value": library
            },
            "shelving_location": location,
          }
        ],
      }

      // decide whether to include interested user information in API payload
      if (interestedUser) {
        var payloadInterestedUser = {
          "interested_user": [
            {
              "primary_id": interestedUser,
              "notify_receiving_activation": true,
              "hold_item": holdInterested,
            }
          ],
        }
      }
      else {
        var payloadInterestedUser = {}
      }

      // construct API payload
      var payload = Object.assign( {}, payloadMain, payloadInterestedUser );

      // compile options to be posted to API
      var options = {
        "method": "post",
        "muteHttpExceptions" : true,
        "headers": {
          "Authorization": apiKey,
          "Content-Type": "application/json",
          "Accept": "application/json"
        },
        "payload": JSON.stringify(payload)
      }
          
      // API request
      var response = UrlFetchApp.fetch(polApiUrl, options);
      var parsedresponse = JSON.parse(response.getContentText());
      console.log(response.getContentText());  
      try {
        var pol = parsedresponse.number;
      } catch (error) {
        console.log(error);
      }
      try {
        var mms = parsedresponse.resource_metadata.mms_id.value;
      } catch (error) {
        console.log(error);
      }

      // Create date string
      var currentdate = new Date(); 
      var createdDate = currentdate.getDate() + "/"
                    + (currentdate.getMonth()+1)  + "/" 
                    + currentdate.getFullYear() + " @ "  
                    + currentdate.getHours() + ":"  
                    + currentdate.getMinutes() + ":" 
                    + currentdate.getSeconds();

      // Place returned record numbers in spreadsheet cells
      polCell.setValue(pol);
      mmsCell.setValue(mms);
      if (pol){
        completedCell.setValue(createdDate);
      } else {
        completedCell.setValue("Error " + createdDate + response.getContentText());
      }

      // Send confirmation email to submitter based on https://spreadsheet.dev/send-email-from-google-sheets
      function emailAuthorizedSubmitter() {
        if (pol) {
          // send confirmation email --> Insert bcc and replyTo email addresses, and sender name here
          var emailSubject = "Record created in Alma for " + title;
          var emailBody = "Based on your POL from submission, a record for " + title + " was created in Alma with POL number " + pol + " and MMS ID " + mms;
          var emailOptions = {
            bcc: "",
            replyTo: "",
            name: ""
          }
          MailApp.sendEmail(submitterEmail, emailSubject, emailBody, emailOptions);
        } else {
          // send error email --> Insert contact information in emailBody, bcc and replyTo email addresses, and sender name here
          var emailSubject = "Error. Record could not be created in Alma";
          var emailBody = "Based on your POL from submission, a record for " + title + " could not be created. Please contact ___ for assistance.\n\n" + response.getContentText();
          var emailOptions = {
            bcc: "",
            replyTo: "",
            name: ""
          }
          MailApp.sendEmail(submitterEmail, emailSubject, emailBody, emailOptions);
        }
      }

      emailAuthorizedSubmitter();

    }
    else {
      console.log("not authorized");
      function emailUnauthorizedSubmitter() {
        var emailSubject = "Not authorized to create POL";
        // --> Insert contact information in emailBody
        var emailBody = "You completed the POL creation request form logged into your Google account as " + submitterEmail + ", which has not been authorized to create POLs using this form. To request permission, please email ___.";
        MailApp.sendEmail(submitterEmail, emailSubject, emailBody);
      }
      emailUnauthorizedSubmitter();
    }
  }

}

