function Reminder_Email() {

    // function variables
    var newDay = new Date();
    var thisDay = newDay.toString().replace(/(\d\d:\d\d:\d\d)/, "00:00:00");
    thisDay = new Date(thisDay);
  
    var destinationFile = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Upcoming Deliveries");
    var values = destinationFile.getRange(2,1,destinationFile.getLastRow(),4).getValues();
  
    // loop through elements
    for (var i = 0; i < values.length; i++){
      if(values[i][0] == ""){
        break;
      }
  
      var eventMonthString = values[i][2];
      var dateParts = eventMonthString.split(/\./);
      var eventMonth = new Date(dateParts[1] +"/"+ dateParts[0] + "/" + thisDay.getFullYear() + " GMT +00");
  
      var monthOwner = values[i][0];
      var firstName = monthOwner.split(" ")[0];
  //    var firstName = nameSplit[0];
      var ownerEmail = values[i][1];
      var relevantMonthString = values[i][3];
  
    // email loops regular & reminder
  
      if (eventMonth.getTime() == thisDay.getTime()) {
        if(firstName.slice(-1) != "s") {
          MailApp.sendEmail(ownerEmail, firstName + "scheduled delivery [Automated email]", relevantMonthString + " is coming up in a few days! For any assistance, feel free to contact us!", {cc: "name@gmail"});
        } else {
           MailApp.sendEmail(ownerEmail, firstName + "scheduled delivery [Automated email]", relevantMonthString + " is coming up in a few days! For any assistance, feel free to contact us!", {cc: "name@gmail"});

        }
      }
  
