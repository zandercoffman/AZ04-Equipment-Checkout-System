//TODO
//Add Students - Done
//Get Students - Done
//Edit Students - Done
//Add Items - Done
//Get Items - Done
//Edit Items - Done
//Add Rentals - Done
//End Rentals - Done
//Edit Rental - Don't need this




//////////////////////////////////////////////////////////////////////
//The weird uncaught "" string lineral contains an unescaped line break I think was caused by the item "Zander's" laptop having a ' in it
//////////////////////////////////////////////////////////////////////

//try to load user preferences

let userPreferences = PropertiesService.getUserProperties();

let spreadsheetURL = userPreferences.getProperty("spreadsheetURL")


let datasheet
if (spreadsheetURL != null && spreadsheetURL != "") {
  datasheet = SpreadsheetApp.openByUrl(spreadsheetURL);
  let sheet = datasheet.getSheetByName("Items");
  let firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  //Going to check if it needs to migrate
  if (firstRow.indexOf("School") === -1) {
    //ruh roh need to migrate
    sheet.deleteColumns(7, 4);
    const lastCol = sheet.getLastColumn();
    sheet.insertColumnsAfter(lastCol, 1);
    sheet.getRange(1, lastCol + 1).setValue('School');
  }
};

function del() {
  userPreferences.deleteProperty("spreadsheetURL")
}


function setInventoryURL(url) {

  try {
    userPreferences.setProperty("spreadsheetURL", url);
    return "success";
  }
  catch {
    return "error";
  }



}

function resetInventoryURL() {
  userPreferences.setProperty("spreadsheetURL", "");
}

/*Currently broken (doesn't seem to persist new current version to old version instances)
function setCurrentVersion() {
  let scriptPreferences = PropertiesService.getScriptProperties();

  currentVersion = "V8"; //change this before running this function
  currentVersionURL = "https://script.google.com/a/macros/stu.tempeunion.org/s/AKfycbzxwqvux-d7U-10G-Q9t-mCNDJbi2VgCesIEhkRp9dkiTzh1SroTNaVxdXfvHxEwUTB/exec"; //change this before running this function
  scriptPreferences.setProperty("currentVersion", currentVersion);
  scriptPreferences.setProperty("currentVersionURL", currentVersionURL)
}

function getCurrentVersion() {
  let scriptPreferences = PropertiesService.getScriptProperties();

  return JSON.stringify({"currentVersion": scriptPreferences.getProperty("currentVersion"), "currentVersionURL": scriptPreferences.getProperty("currentVersionURL")});
}

*/

function getSetupPage() {
  return HtmlService.createHtmlOutputFromFile("setup").setTitle("Equipment Checkout System Setup");
}

function doGet(e) {


  if (e.pathInfo == null || e.pathInfo == "" || e.pathInfo == "/index" || e.pathInfo == "/index.html") {


    Logger.log(JSON.stringify(userPreferences))

    //return the setup page unless the spreadhsheetURL has already been set
    if (spreadsheetURL == null || spreadsheetURL == "") {
      return getSetupPage();
    }
    else {
      let htmlCode = ScriptApp.getResource('index').getDataAsString();


      /*htmlParts = htmlCode.split("</script>")
      
      let output = HtmlService.createHtmlOutput("")

      for (let i = 0; i<htmlParts.length - 2; i++) {
        let part = htmlParts[i]

        output.append(part)
        output.append("/")
        output.append("/# sourceURL=javascript.js")
        output.append("</script>")
      }
      


      let secondToLastPart = htmlParts[htmlParts.length - 2]

      output.append(secondToLastPart)
      output.append("/")
      output.append("/# sourceURL=javascript.js")
      output.append("</script>")



      output.append(htmlParts[htmlParts.length - 1])
      return output;*/





      //Did this weird workaround to insert the js at runtime on the client instead of just doing return HTMLService.createHTMLServiceFromFile('index') with the js being in a script tag to resolve the following error
      //Uncaught SyntaxError: "" string literal contains an unescaped line break userCodeAppPanel:1045:20
      //userCodeAppPanel is something that google app scripts creates that might contain the code we put in script tags, and when clicking the error line number it took me to a long one liner file that I don't think even had code we wrote
      //with this new weird workaround with js inserted at runtime, the error is gone (don't know why)

      //////////////////////////
      //////////////////////////
      //I might have figured out the weird error (since there was a ' in a new item name and stuff wasn't base64 encoded for the openItemDetails function (now it is so this workaround might no longer be needed))
      ///////////////////////////
      ///////////////////////////

      //split apart the index.html file and extract the js which needs to be removed and inserted at runtime to fix the weird error above
      htmlParts = htmlCode.split("<!-- Start of js to remove and insert at runtime-->")
      htmlPart0 = htmlParts[0];
      htmlParts2 = htmlParts[1].split("<!-- End of js to remove and insert at runtime-->")
      htmlPart1 = htmlParts2[0];
      htmlPart2 = htmlParts2[1];

      htmlParts = [htmlPart0, htmlPart1, htmlPart2]

      if (htmlParts.length != 3) {
        Logger.log("error generating index.html in Code.gs")
        return HtmlService.createHtmlOutput("error generating index.html in Code.gs")
      }

      let output = HtmlService.createHtmlOutput(htmlParts[0]) //start creating the html object to return using the first part of the html up to where the js needs to be added



      let jsCode = htmlParts[1].split('<script>')[1].split("</script>")[0] //remove the script tags from htmlParts[1] to extract only the js code


      //this adds a script tag with a self executing function whose only purpose is to add another script tag at runtime on the client with the actual js which was extracted from index.html above
      //modified from https://stackoverflow.com/questions/74050409/how-to-map-client-side-code-to-source-code
      output.append(`<script>
      (function () {
      const code = decodeURIComponent(atob('${Utilities.base64Encode(
        encodeURIComponent(jsCode)
      )}'));
      const scriptEl = document.createElement('script');
      scriptEl.textContent = code;
      //scriptEl.setAttribute("type", "module");
      document.body.appendChild(scriptEl);
      })();
      </script>`)



      //add the last bit of the html
      output.append(htmlParts[2])

      //return the html
      return output.setTitle("Equipment Checkout System");


    }

  }
  else {
    return HtmlService.createHtmlOutput("<p>" + e.pathInfo + "</p>");
  }
}


function getJSON(element) {
  switch (element) {
    case "Student":
      return JSON.stringify(getStudents())
    case "Items":
      return JSON.stringify(getItems());
    case "Rentals":
      Logger.log("Getting rental info")
      Logger.log("Rental Info:" + JSON.stringify(getRentals()))
      return JSON.stringify(getRentals())
    default:
      return ""
  }
}

/* obsolete and won't work anymore
function createRental(itemID, studentID) {
  //get the list of items in order to call createRental on the item object
  let items = getItems();

  items[itemID].createRental(studentID);


  //not sure if this needs to return anything (need to try calling this function from the frontend and test fail cases)
}
*/
// REMEMBER IT DOES NOT HAVE THE "S" in the name
//take a paramater students as input (this should be of the structure [{"id": "ID1", "classYear": "CLASS_YEAR1", "period": "PERIOD1", "name": "NAME1"}, {"id": "ID2", "classYear": "CLASS_YEAR2", "period": "PERIOD2", "name": "NAME2"}])
function addStudentsToDatabase(students) {
  Logger.log("AHHHHH: " + JSON.stringify(students));
  let currentStudents = Object.keys(getStudents());

  //check that there are no errors before editing the database
  let error = "";

  students.forEach((student) => {
    if (currentStudents.includes(student.id)) {
      error = "Error: Student ID '" + student.id + "' already exists in the database";

    }
    student.id = student.id.replace(/\D/g, "");
    student.classYear = "20" + student.id.substring(0, 2);
  })

  if (error != "") {
    return error;
  }
  else {
    //if there was no error, then write the new students to the database
    //find the sheet
    let studentSheet = datasheet.getSheetByName("Students");
    let dataRange = studentSheet.getDataRange();
    let height = dataRange.getHeight();
    Logger.log(height);


    //convert the data into the format .setValues() wants
    let dataToWrite = []
    students.forEach((student) => {
      dataToWrite.push([student.id, student.classYear, student.period, student.name]);
    });

    let rangeToWrite = studentSheet.getRange(height + 1, 1, students.length, 4);
    rangeToWrite.setValues(dataToWrite)


    return "success";
  }
}



//take a paramater items as input (this should be of the structure [{roomArea: "C999", itemDescription: "item description here", makeModel: "brand", assetTagNumber: "EX999", serialNumber: "SN999", datePurchased: "2033-2034", purchaseOrder: "", fund: "", cost: "999.99", notes: "this is a test item", status: "In", rentalID: ""}])
function addItemsToDatabase(items) {
  let currentItems = Object.keys(getItems());

  //check that there are no errors before editing the database
  let error = "";
  items.forEach((item) => {
    if (currentItems.includes(item.assetTagNumber)) {
      error = "Error: Asset ID '" + item.assetTagNumber + "' already exists in the database";

    }
  })

  if (error != "") {
    return error;
  }
  else {
    //if there was no error, then write the new students to the database
    //find the sheet
    let itemSheet = datasheet.getSheetByName("Items");
    let dataRange = itemSheet.getDataRange();

    Logger.log(dataRange.getValues());
    let height = dataRange.getHeight();
    Logger.log(height);



    //convert the data into the format .setValues() wants
    let dataToWrite = []
    items.forEach((item) => {
      dataToWrite.push([item.roomArea, item.itemDescription, item.makeModel, item.assetTagNumber, item.serialNumber, item.notes, item.status, item.rentalID, item.school]);
    });

    let rangeToWrite = itemSheet.getRange(height + 1, 2, items.length, 9);
    rangeToWrite.setValues(dataToWrite)

    return "success";
  }
}

//take a paramater items as input (this should be of the structure [{roomArea: "C999", itemDescription: "item description here", makeModel: "brand", assetTagNumber: "EX999", serialNumber: "SN999", datePurchased: "2033-2034", purchaseOrder: "", fund: "", cost: "999.99", notes: "this is a test item", status: "In", rentalID: ""}])
function editItemsInDatabase(items) {
  items = JSON.parse(items)

  let currentItems = Object.keys(getItems());

  //check that there are no errors before editing the database
  let error = "";
  items.forEach((item) => {
    if (!currentItems.includes(item.assetTagNumber)) {
      error = "Error: Asset ID '" + item.assetTagNumber + "' does not already exist in the database, call addItemsToDatabase instead";
    }
  })

  if (error != "") {
    return error;
  }
  else {
    //if there was no error, then write the new students to the database
    //find the sheet
    let itemSheet = datasheet.getSheetByName("Items");
    let dataRange = itemSheet.getDataRange();

    Logger.log(dataRange.getValues());
    let height = dataRange.getHeight();
    Logger.log(height);


    //convert the data into the format .setValues() wants and update each item
    let dataToWrite = []
    items.forEach((item) => {
      dataToWrite = [[item.roomArea, item.itemDescription, item.makeModel, item.assetTagNumber, item.serialNumber, item.notes, item.status, item.rentalID, item.school]];
      let wroteData = false;

      for (let r = 0; r < height; r++) {
        let id = dataRange.getValues()[r][4]
        Logger.log(id)

        if (id == item.assetTagNumber) {

          let rangeToWrite = itemSheet.getRange(r + 1, 2, 1, 9);
          rangeToWrite.setValues(dataToWrite)
          Logger.log("Wrote data")
          wroteData = true;
          break;
        }
      }

      if (!wroteData) {
        Logger.log("Failed to update item: " + item);
        return "Failed to update item: " + item;
      }

    });

  }
}

//take a paramater students as input (this should be of the structure [{id: "ID", classYear: "CLASSYEAR", period: "PERIOD", name: "NAME"}])
function editStudentsInDatabase(students) {
  Logger.log(students)
  students = JSON.parse(students)
  Logger.log(students)

  let currentStudents = Object.keys(getStudents());

  //check that there are no errors before editing the database
  let error = "";
  students.forEach((student) => {
    if (!currentStudents.includes(student.id)) {
      error = "Error: Student ID '" + student.id + "' does not already exist in the database, call addStudentsToDatabase instead";
    }
  })

  if (error != "") {
    return error;
  }
  else {
    //if there was no error, then write the new students to the database
    //find the sheet
    let studentSheet = datasheet.getSheetByName("Students");
    let dataRange = studentSheet.getDataRange();

    Logger.log(dataRange.getValues());
    let height = dataRange.getHeight();
    Logger.log(height);


    //convert the data into the format .setValues() wants and update each student
    let dataToWrite = []
    students.forEach((student) => {
      dataToWrite = [[student.id, student.classYear, student.period, student.name]];
      let wroteData = false;

      for (let r = 0; r < height; r++) {
        let id = dataRange.getValues()[r][0]
        Logger.log(id)

        if (id == student.id) {

          let rangeToWrite = studentSheet.getRange(r + 1, 1, 1, 4);
          rangeToWrite.setValues(dataToWrite)
          Logger.log("Wrote data")
          wroteData = true;
          break;
        }
      }

      if (!wroteData) {
        Logger.log("Failed to update student: " + student);
        return "Failed to update student: " + student;
      }
    });

  }
}

function deleteItemsFromInventory(assetTags) {
  const sheet = datasheet.getSheetByName("Items");
  var range = sheet.getDataRange();
  var values = range.getValues();

  let error = false;

  Logger.log("Attempting to delete the following items: " + JSON.stringify(assetTags))
  assetTags.forEach((assetTag) => {
    for (var i = 0; i < values.length; i++) {
      if (values[i][4] == assetTag) {
        sheet.deleteRow(i + 1);
        error = true;
      }
    }
  })

  return error;
}

//this function should now be able to be removed, deleteItemsFromInventory should work if you pass it a list with only one item
function deleteItemFromInventory(assetTag) {
  deleteItemsFromInventory([assetTag]);
}


function deleteStudentsFromDatabase(ids) {
  Logger.log(ids)
  const sheet = datasheet.getSheetByName("Students");
  var range = sheet.getDataRange();
  var values = range.getValues();

  let error = false;

  //remove the header row from the data
  values.shift();

  ids.forEach((id) => {
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] == id) {
        sheet.deleteRow(i + 2);
        error = true;
      }
    }
  })

  return error;
}


//Is this still used?
function getInventoryItems() {
  const sheet = datasheet.getSheetByName("Items");
  const data = sheet.getDataRange().getValues().slice(1);

  return data.map(row => row.slice(1));
}

function convertSchool(school) {
  let fullString = ""
  let arr = school.split(" ");
  if (arr.length <= 1) {
    return school;
  }

  for (let i = 0; i < fullString.length; i++) {
    fullString = fullString + arr[i].substring(0, 1);
  }

  return fullString;
}

function getStudents() {

  //find the sheet
  let studentSheet = datasheet.getSheetByName("Students");

  //find the range which contains data
  let dataRange = studentSheet.getDataRange();

  //get the data from the useful range
  let dataToConvert = dataRange.getValues();

  //a dictionary which will be used to store student info
  let students = {};

  //remove the first row containing headers
  dataToConvert.shift();


  //for each student in the table convert the data into an easier to use format and append it to the students dictonary
  for (student of dataToConvert) {

    students[student[0]] =
    {
      id: student[0],
      classYear: student[1],
      period: student[2],
      name: student[3]
    }
  }

  //return the student info
  return students;
}

function getItems() {

  //find the sheet
  let itemSheet = datasheet.getSheetByName("Items");

  //find the range which contains data
  let dataRange = itemSheet.getDataRange();

  //get the data from the useful range
  let dataToConvert = dataRange.getValues();


  //a dictionary which will be used to store item info
  let items = {};

  //remove the first row containing headers
  dataToConvert.shift();


  //for each item in the table convert the data into an easier to use format and append it to the items dictonary
  for (item of dataToConvert) {

    items[item[4]] =
    {
      roomArea: item[1],
      itemDescription: item[2],
      location: item[9],
      assetTagNumber: item[4],
      serialNumber: item[5],
      notes: item[6],
      status: item[7],
      rentalID: item[8],
      school: item[9]
    }
  }

  //return the item info
  return items
}

function buildRentalReceiptHTML({ studentID, itemName, assetTag, rentalID, dateOut, dateToReturn, contactEmail }) {
  return `
  <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #2c2c2c; padding: 24px; max-width: 600px; margin: auto; border: 1px solid #dcdcdc; border-radius: 8px; background-color: #ffffff;">
    <h2 style="text-align: center; color: #1a73e8; margin-bottom: 20px;">📋 Rental Receipt: ${itemName}</h2>
    
    <p style="font-size: 15px; line-height: 1.6;">
      This email confirms a rental has been successfully processed in the Equipment Checkout System.
      Please find the details of the transaction below:
    </p>

    <table style="width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px;">
      <tr>
        <td style="padding: 10px; font-weight: bold; background-color: #f5f5f5;">Rental ID:</td>
        <td style="padding: 10px;">${rentalID}</td>
      </tr>
      <tr>
  <td style="padding: 10px; font-weight: bold; background-color: #f5f5f5;">Item Name:</td>
  <td style="padding: 10px;">${itemName}</td>
</tr>
      <tr>
        <td style="padding: 10px; font-weight: bold; background-color: #f5f5f5;">Asset Tag:</td>
        <td style="padding: 10px;">${assetTag}</td>
      </tr>
      <tr>
        <td style="padding: 10px; font-weight: bold; background-color: #f5f5f5;">Student ID:</td>
        <td style="padding: 10px;">${studentID}</td>
      </tr>
      <tr>
        <td style="padding: 10px; font-weight: bold; background-color: #f5f5f5;">Date Rented Out:</td>
        <td style="padding: 10px;">${dateOut}</td>
      </tr>
      <tr>
        <td style="padding: 10px; font-weight: bold; background-color: #f5f5f5;">Expected Return Date:</td>
        <td style="padding: 10px;">${dateToReturn}</td>
      </tr>
    </table>

    <p style="margin-top: 24px; font-size: 14px; line-height: 1.6;">
      Please ensure the asset is returned by the expected return date.
    </p>

    <hr style="margin: 32px 0; border: none; border-top: 1px solid #e0e0e0;">

    <p style="font-size: 13px; color: #666;">
      This is an automated message from the Equipment Checkout System. For questions or assistance, please contact your teacher at <a href="mailto:${contactEmail}" style="color: blue; cursor: pointer;">${contactEmail}</a>.

    </p>

    <p style="text-align: center; font-size: 12px; color: #999; margin-top: 20px;">
      © ${new Date().getFullYear()} Tempe Union High School District
    </p>
  </div>
  `;
}



function createRental(studentID, assetTag, dateToReturn) {

  let items = getItems();
  let item = items[assetTag];
  let itemName = item.itemDescription || "Unnamed Item";


  // Convert the dateToReturn into a proper Date object
  let date = new Date(dateToReturn);

  // Check if the date is valid
  if (isNaN(date.getTime())) {
    return "Invalid return date provided.";
  }

  let students = getStudents();

  // Check that the student to loan the item to is in the database
  let studentExists = false;
  for (key of Object.keys(students)) {
    if (studentID == key) {
      studentExists = true;
    }
  }

  if (!studentExists) {
    return "Can't create new rental, the student ID specified is not in the database";
  }
  else {

    // Now check if the item is already checked out
    if (item["status"] != "In") {
      return "Can't create new rental, the item has a status of '" + item["status"] + "'";
    }

    // Create a rental

    // Fetch the existing rentals to figure out what the new rentalID should be
    let rentals = getRentals();
    let rentalKeys = Object.keys(rentals);
    let newRentalID;

    if (rentalKeys.length > 0) {
      newRentalID = Number.parseInt(rentalKeys[rentalKeys.length - 1]) + 1;
    }
    else {
      newRentalID = 1;
    }

    // Find the rental sheet
    let rentalSheet = datasheet.getSheetByName("Rentals");

    let rentalsRange = rentalSheet.getDataRange();
    let numberOfRentals = rentalsRange.getHeight() - 1;

    // Start looking for which row data should be written to
    let rowToWriteTo = numberOfRentals + 2;

    // Write the new rental info to the rental sheet
    dataRangeForWrite = rentalSheet.getRange(rowToWriteTo, 1, 1, 7);

    // Format the return date to match the format expected by your Google Sheets (optional)
    let formattedReturnDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    dataRangeForWrite.setValues([[newRentalID.toString(), assetTag.toString(), "Out", studentID.toString(), new Date(Date.now()).toString(), formattedReturnDate, ""]]);

    // Update the item sheet
    let itemSheet = datasheet.getSheetByName("Items");

    for (let r = 1; r <= itemSheet.getDataRange().getHeight() + 1; r++) {
      if (itemSheet.getRange(r, 5).getValue() == assetTag) {
        Logger.log("Item is located in row " + r);

        // Update the status of the item to "Out" and set the rental ID
        itemSheet.getRange(r, 8, 1, 2).setValues([["Out", newRentalID.toString()]]);
        break;
      }
    }

    let teacherEmail = Session.getActiveUser().getEmail();
    let studentEmail = `s${studentID}@stu.tempeunion.org`;


    //Sends email
    try {


      let rentalIDStr = newRentalID.toString();
      let dateOutStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

      let html = buildRentalReceiptHTML({
        studentID,
        itemName,
        assetTag,
        rentalID: newRentalID.toString(),
        dateOut: dateOutStr,
        dateToReturn: formattedReturnDate,
        contactEmail: teacherEmail
      });


      let subject = `Rental Confirmation: ${itemName}`;

      MailApp.sendEmail({
        to: teacherEmail,
        subject: subject,
        htmlBody: html
      });
    } catch (err) {
      Logger.log(JSON.stringify(err));
    } finally {
      return "success"
    }


  }
}


function getRentals() {

  //find the sheet
  let rentalSheet = datasheet.getSheetByName("Rentals");

  //find the range which contains data
  let dataRange = rentalSheet.getDataRange();

  //get the data from the useful range
  let dataToConvert = dataRange.getValues();

  //a dictionary which will be used to store rental info
  let rentals = {};

  //remove the first row containing headers
  dataToConvert.shift();


  //for each rental in the table convert the data into an easier to use format and append it to the rentals dictonary
  for (rental of dataToConvert) {
    Logger.log("rentalInfo: ")
    Logger.log(rental[3])
    rentals[rental[0]] =
    {
      rentalID: Number.parseInt(rental[0]),
      itemID: rental[1],
      status: rental[2],
      checkedOutTo: rental[3],
      dateCheckedOut: new Date(rental[4]),
      dateExpectedReturned: new Date(rental[5]),
      dateReturned: new Date(rental[6]),
      notes: rental[8],
      returnRental: function (returnedBy) {


        //update the data on the rental sheet
        for (let r = 1; r <= dataRange.getHeight() + 1; r++) {
          if (rentalSheet.getRange(r, 1).getValue() == this.rentalID) {
            Logger.log("Item is located in row " + r);

            rentalSheet.getRange(r, 3).setValue("Returned");
            rentalSheet.getRange(r, 7).setValue(new Date(Date.now()).toString());
            rentalSheet.getRange(r, 8).setValue(returnedBy);
            break;
          }
        }

        //update the data on the item sheet

        let itemSheet = datasheet.getSheetByName("Items");
        let itemDataRange = itemSheet.getDataRange();

        for (let r = 1; r <= itemDataRange.getHeight() + 1; r++) {
          if (itemSheet.getRange(r, 5).getValue() == this.itemID) {
            Logger.log("Item is located in row " + r);

            itemSheet.getRange(r, 8, 1, 2).setValues([["In", ""]]);
            break;
          }
        }
      },
      returnedBy: rental[7]
    }
  }

  //return the student info
  return rentals;
}

function buildMultiRentalReceiptHTML({ rentals, contactEmail, anyLate }) {
  const studentID = rentals[0].studentID;

  const rows = rentals.map(rental => `
    <tr>
      <td style="padding: 10px; border: 1px solid #e0e0e0;">${rental.assetTag}</td>
      <td style="padding: 10px; border: 1px solid #e0e0e0;">${rental.rentalID}</td>
      <td style="padding: 10px; border: 1px solid #e0e0e0;">${rental.dateOut}</td>
      <td style="padding: 10px; border: 1px solid #e0e0e0;">${rental.dateReturned}</td>
      <td style="padding: 10px; border: 1px solid #e0e0e0; color: ${rental.wasLate ? '#e74c3c' : '#2ecc71'};">
        ${rental.wasLate ? 'Late' : 'On Time'}
      </td>
    </tr>
  `).join("");

  return `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #2c2c2c; padding: 24px; max-width: 700px; margin: auto; border: 1px solid #dcdcdc; border-radius: 8px; background-color: #ffffff;">
      <h2 style="text-align: center; color: #1a73e8; margin-bottom: 20px;">📦 Rental Return Receipt</h2>

      <p style="font-size: 15px; line-height: 1.6;">
        This email confirms the return of the following item(s) in the Equipment Checkout System for Student ID: <strong>${studentID}</strong>.
      </p>

      <table style="width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 14px;">
        <thead>
          <tr style="background-color: #f5f5f5;">
            <th style="padding: 10px; border: 1px solid #e0e0e0;">Asset Tag</th>
            <th style="padding: 10px; border: 1px solid #e0e0e0;">Rental ID</th>
            <th style="padding: 10px; border: 1px solid #e0e0e0;">Date Out</th>
            <th style="padding: 10px; border: 1px solid #e0e0e0;">Date Returned</th>
            <th style="padding: 10px; border: 1px solid #e0e0e0;">Status</th>
          </tr>
        </thead>
        <tbody>
          ${rows}
        </tbody>
      </table>

      ${anyLate
      ? `<p style="margin-top: 24px; font-size: 14px; color: #c0392b;">
            <strong>Note:</strong> One or more items were returned past the due date.
            Please try to return items on time to avoid any issues in the future.
          </p>`
      : `<p style="margin-top: 24px; font-size: 14px; color: #27ae60;">
            All items were returned on time. Great job!
          </p>`}

      <hr style="margin: 32px 0; border: none; border-top: 1px solid #e0e0e0;">

      <p style="font-size: 13px; color: #666;">
        This is an automated message from the Equipment Checkout System. For questions or assistance, please contact your teacher at
        <a href="mailto:${contactEmail}" style="color: #1a73e8;">${contactEmail}</a>.
      </p>

      <p style="text-align: center; font-size: 12px; color: #999; margin-top: 20px;">
        © ${new Date().getFullYear()} Tempe Union High School District
      </p>
    </div>
  `;
}



//takes a list of item IDs to return the rentals for
function returnRentals(rentalsToReturn, returnedBy) {
  const rentalInfo = getRentals();
  const teacherEmail = Session.getActiveUser().getEmail();

  let returnedRentals = [];
  let anyLate = false;
  let numberOfRentalsReturned = 0;

  Logger.log("Starting rental return process...");

  rentalsToReturn.forEach((rentalID) => {
    Object.values(rentalInfo).forEach((existingRental) => {
      if (existingRental.itemID == rentalID && existingRental.status === "Out") {
        // Process the return
        existingRental.returnRental(returnedBy);
        numberOfRentalsReturned++;

        // Log to check structure
        Logger.log("Returning rental:", existingRental);

        const studentID = existingRental.checkedOutTo || "UNKNOWN";
        const studentEmail = `s${studentID}@stu.tempeunion.org`;
        const rentalIDStr = existingRental.rentalID?.toString() || "N/A";

        const dateOut = new Date(existingRental.dateCheckedOut);
        const dateOutStr = Utilities.formatDate(dateOut, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

        const dateReturned = new Date();
        const returnDateStr = Utilities.formatDate(dateReturned, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

        const dueDate = new Date(existingRental.dateExpectedReturned);
        const wasLate = dateReturned.getTime() > dueDate.getTime();

        let itemName = ""
        let assetTag = existingRental.itemID

        if (wasLate) anyLate = true;

        returnedRentals.push({
          studentID,
          studentEmail,
          itemName,
          assetTag,
          rentalID: rentalIDStr,
          dateOut: dateOutStr,
          dateReturned: returnDateStr,
          wasLate
        });
      }
    });
  });

  if (returnedRentals.length === 0) {
    return "failed, no outstanding rental with the information specified found";
  }

  const studentEmail = returnedRentals[0].studentEmail;

  const subject = anyLate
    ? `Rental Return Receipt (Some Late Returns)`
    : `Rental Return Receipt (All On Time)`;

  const messageHTML = buildMultiRentalReceiptHTML({
    rentals: returnedRentals,
    contactEmail: teacherEmail,
    anyLate
  });

  try {
    MailApp.sendEmail({
      to: teacherEmail,
      cc: teacherEmail,
      subject: subject,
      htmlBody: messageHTML
    });
  } catch (err) {
    Logger.log("Email send error:", JSON.stringify(err));
  }

  return "success";
}

//moved to frontend (no longer used here as of 2025-1-23)
function getCodeType(code) {
  let items = Object.keys(getItems());
  let students = Object.keys(getStudents());

  let isItem = false;
  let isStudent = false;

  if (items.includes(code)) {
    isItem = true;
  }
  if (students.includes(code)) {
    isStudent = true;
  }

  if (!isItem && !isStudent) {
    return "notInDatabase";
  }

  if (isItem && isStudent) {
    return "error: both a student and an item are registered with this code/id"
  }

  if (isItem) {
    return "item"
  }
  else {
    if (isStudent) {
      return "student";
    }
  }
}
function generateReport(reportType, timespanStart, timespanEnd) {



  let startDate = new Date(timespanStart)
  let endDate = new Date(timespanEnd);


  let output = { "head": [[]], "body": [] };

  switch (reportType) {



    case "itemPopularityUsageReport": {
      let rentals = getRentals();

      let itemStatistics = {}

      Object.values(rentals).forEach((rental) => {
        let rentalDate = new Date(rental["dateCheckedOut"]);
        if (rentalDate.getTime() >= startDate.getTime() && rentalDate.getTime() < endDate.getTime()) {
          let currentCount = itemStatistics[rental.itemID];
          if (currentCount == null) {
            currentCount = 0;
          }

          itemStatistics[rental.itemID] = currentCount + 1;
        }
      })




      let items = getItems();

      output["head"] = [["Asset Tag Number", "Brand", "Name", "# of Times Checked Out During This Period"]];

      //sort the item statistics (modified from https://www.geeksforgeeks.org/how-to-sort-a-dictionary-by-value-in-javascript/)
      let temp = Object.keys(itemStatistics)
        .sort((a, b) => itemStatistics[b] - itemStatistics[a])
        .reduce((temp2, key) => {
          temp2[key] = itemStatistics[key];
          return temp2;
        }, {});
      itemStatistics = temp;


      Object.keys(itemStatistics).forEach((id) => {
        Logger.log(id)
        Logger.log(items)
        output["body"].push([items[id]["assetTagNumber"], items[id]["makeModel"], items[id]["itemDescription"], itemStatistics[id]]);
      })
    }
      break;
    case "lostItemReport": {

      let items = getItems();
      let rentals = getRentals();

      let itemStatistics = [];
      let error = "";
      Object.keys(items).forEach((key) => {
        Logger.log(items[key]["status"])
        if (items[key]["status"] == "Lost") {

          Logger.log(items[key]["rentalID"])
          Logger.log(items[key])

          try {
            itemStatistics.push([key, items[key]["makeModel"], items[key]["itemDescription"], rentals[items[key]["rentalID"]]["checkedOutTo"], rentals[items[key]["rentalID"]]["dateCheckedOut"], rentals[items[key]["rentalID"]]["dateExpectedToReturn"]]);
          }
          catch (e) {
            error = "The server encountered an error while generating a lost item report";
          }

        }
      })


      if (error != "") {
        Logger.log("error");
        Logger.log(error);
        return error
      }
      output["head"] = [["Asset Tag Number", "Brand", "Name", "Checked Out To", "Date Checked Out", "Date Expected To Return"]];


      /*//sort the item statistics (modified from https://www.geeksforgeeks.org/how-to-sort-a-dictionary-by-value-in-javascript/)
      let temp = Object.keys(itemStatistics)
      .sort((a, b) => itemStatistics[b] - itemStatistics[a])
      .reduce((temp2, key) => {
        temp2[key] = itemStatistics[key];
        return temp2;
      }, {});
      itemStatistics = temp;*/

      output["body"] = itemStatistics;
    }
      break;
  }
  Logger.log("output");
  Logger.log(output);
  return JSON.stringify(output);
}
function doSomething() {
  Logger.log('I was called!');

  //let items = getItems();
  //let itemKeys = Object.keys(items);

  //items[itemKeys[0]].createRental(1);

  //let rentals = getRentals()
  //Logger.log(rentals);
  //rentals[3].returnRental();


  //Logger.log(addStudentsToDatabase([{"id": "ID3", "classYear": "2100", "period": "6", "name": "Student from the far future"}]))

  //Logger.log(editItemsInDatabase([{roomArea: "C999", itemDescription: "item description here", makeModel: "brand2", assetTagNumber: "EX999", serialNumber: "SN999", datePurchased: "2033-2034", purchaseOrder: "", fund: "", cost: "999.99", notes: "this is a test item", status: "In", rentalID: ""}]))

  //editStudentsInDatabase([{id: "99999", classYear: "99", period: "99", name: "Ninety Nine"}, {id: "55555", classYear: "505", period: "3", name: "Fifty Five"}])

  //Logger.log(addStudentsToDatabase([{id: "55555", classYear: "5555", period: "5", name: "five"}, {id: "77777", classYear: "7777", period: "7", name: "seven"}]))

  Logger.log(getPage('addStudents'));

  output = "I ran";


  return output;
}

function getSpreadsheetURL() {
  return spreadsheetURL;
}

function findMakeModelBasedOnAssetTag(assetTag) {
  let sheet = datasheet.getSheetByName("Items");
  var data = sheet.getDataRange().getValues();

  // Loop through the entire sheet to find where the assetTagNumber matches
  for (var i = 1; i < data.length; i++) {  // Start from 1 to skip headers
    for (var j = 0; j < data[i].length; j++) {  // Loop through each column in the row
      var assetTagFromSheet = data[i][j].toString().trim();  // Get the value at [i][j] and trim any whitespace

      // If we find a match with assetTagNumber
      if (assetTagFromSheet === assetTag.trim()) {
        // Return the value from the previous column (i.e., [i][j-1])
        if (j > 0) {  // Make sure we're not trying to access an invalid index (first column)
          var valueBefore = data[i][j - 1];
          Logger.log("Found match! Returning value from column before: " + valueBefore);
          return valueBefore;
        }
      }
    }
  }

  // If not found, return a message or empty string
  Logger.log("Asset tag not found.");
  return "Asset tag not found.";
}
