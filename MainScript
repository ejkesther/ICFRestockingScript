// THIS CODE AIMS TO AUTOMATE INVENTORY TRACKING
// SHOULD BE ROBUST ENOUGH TO ACCOUNT FOR STUDENT ERRORS 
// SET UP WEEKLY DIGEST OF INVENTORY STATS?

// FUNCTIONALITY:
  // EMAILS SERVICE LEAD WHEN ONE OR MORE ITEMS IS LOW IN STOCK
  // DOES NOT NOTIFY WHEN
    // WHEN COUNTS ARE STABLE / EQUIPMENT DOES NOT NEED TO BE ORDERED
    // STUDENT MISSES A SHIFT / NO EDITS MADE

// POTENTIAL TO-DO: 
  // EXTRA CELL MANAGEMENT OR FORM? SAYS YOU PURCHASED SO IT KNOWS NOT TO NOTIFY AGAIN

/*****************************************************************************/


// stores variable "lowStockStatus", where 0 means inventory stock is all good
// 1 means one or more items is low in stock (O-FALSE, 1-TRUE)
var properties = PropertiesService.getScriptProperties();

const currSheet = SpreadsheetApp.getActiveSheet();

// Dictionary of {days of the week}:{corresponding range}
var time = new Date();
var dayOfTheWeek = time.getDay();
var rangeOfDaysOfTheWeek = {
  1:'B27:B37', // Monday
  2:'C27:C37', // Tuesday
  3:'D27:C37', // Wednesday
  4:'E27:E37', // Thursday
  5:'F27:F37' // Friday
}

var previousDay;
if (dayOfTheWeek == 1) { // IF MONDAY, SEND REPORT OF FRIDAY'S STOCK
  previousDay = 5;
} else {
  previousDay = dayOfTheWeek - 1;
}

var previousDayRange = rangeOfDaysOfTheWeek[previousDay];
const inventoryRange = currSheet.getRange(previousDayRange);

// Table including all items and their item count (from the previous working day)
var lowStockItems = {
  "Microphones":inventoryRange.getCell(1,1).getValue(),
  "VGA Cables":inventoryRange.getCell(2,1).getValue(),
  "HDMI Cables":inventoryRange.getCell(3,1).getValue(),
  "9V batteries":inventoryRange.getCell(4,1).getValue(),
  "AA batteries":inventoryRange.getCell(5,1).getValue(),
  "AAA batteries":inventoryRange.getCell(6,1).getValue(),
  "HDMI - VGA":inventoryRange.getCell(7,1).getValue(),
  "USB-C - VGA":inventoryRange.getCell(8,1).getValue(),
  "MD - VGA":inventoryRange.getCell(9,1).getValue(),
  "USB-C - HDMI":inventoryRange.getCell(10,1).getValue(),
  "MD - HDMI":inventoryRange.getCell(11,1).getValue()
};

// a list of all items *excluding 9V, AA batteries, and HDMI - VGA* that are notified of low stock at <= 5 count
var rowToNonSpecialItemNames = { 
  1:'Microphones',
  2:'VGA Cables',
  3:'HDMI Cables',
  6:'AAA batteries',
  8:'USB-C - VGA',
  9:'MD - VGA',
  10:'USB-C - HDMI',
  11:'MD - HDMI'
}

// onEdit event trigger
// tracks if the edit made puts the item at a "LowStockStatus" status
function onEdit()
{  
  const range = SpreadsheetApp.getActiveSheet().getActiveRange();
  const row = range.getRow();
  const value = range.getValue();

  // else/if statements check validity of the value change according to equipment type
  if (value && row == 30 && value <= 15) { // 9V Battery, should be <= 15 before alert
    properties.setProperty('lowStockStatus', 1);

  } else if (value && row == 31 && value <= 10) { // AA Battery, should be <= 10 before alert
    properties.setProperty('lowStockStatus', 1);
  
  } else if (value && row == 33 && value <= 7) { // HDMI-VGA Adaptor, should be <= 7 before alert
    properties.setProperty('lowStockStatus', 1);
  
  } else if (value && value <= 5) { // everything else should alert when there are <= 5
    properties.setProperty('lowStockStatus', 1);    
  } 
}

// triggered every morning 8-9 AM PST, 
// if a change was made to sheet last night that indicated low inventory,
// an email is sent
function checkToPushUpdate() 
{
  var lowStockStatus = properties.getProperty('lowStockStatus');
  if (lowStockStatus == 1 && dayOfTheWeek != 0 && dayOfTheWeek != 6) { //DO NOT ALERT ON WEEKENDS/IF NO LOW-STOCK
    var itemsToPurchaseList = new Array();
    for (const [key, value] of Object.entries(lowStockItems)) {
      if (key == "9V batteries") {
        if (value <= 15) {
          itemsToPurchaseList.push('9V batteries');
        }
      } else if (key == "AA batteries") {
        if (value <= 10) {
          itemsToPurchaseList.push('AA batteries');
        }
      } else if (key == "HDMI - VGA") {
        if (value <= 7) {
          itemsToPurchaseList.push('HDMI - VGA');
        }
      } else if (value <= 5) {
        itemsToPurchaseList.push(key)
      }
    }
    var itemsToPurchaseStringList = itemsToPurchaseList.join(', ');
    update(itemsToPurchaseStringList);
  }
}

function update(itemsToPurchaseList) 
{
  // POP UP MESSAGE AFTER CONDITIONS ARE MET TO SEND EMAIL
  SpreadsheetApp.getUi().alert("One or more items in low stock -- Email sent to service lead!");

  // CONFIRMATION EMAIL SENT TO SERVICE LEAD NOTIFYING OF LOW STOCK
  SendEmail(itemsToPurchaseList);

  // RESET "lowStockStatus" status
  properties.setProperty('lowStockStatus', 0);
}

// CONVERT DICTIONARY OBJECTS TO STRING USING ARRAY STRING CONCATENATION
function dictToStringTable(dict) 
{
  var array = new Array();
  for (const [key, value] of Object.entries(dict)) {
  array.push(`${key}: ${value}`);
  }
  return array.join("\n");
}

// SENDS EMAIL TO SPECIFIC EMAIL ADDRESS
function SendEmail(itemsToPurchaseList) 
{
  // Set Recipient Email Address - CAN USE FOR TESTING
  var emailAddress = 'ejkesther@berkeley.edu';

  // Send Alert Email. CAN CHANGE SUBJECT OR MESSAGE.
  var message = 'Hello! \n \nYou are receiving this email due to low supply counts in HG6. \n \nPlease consider repurchasing the following items:\n' + itemsToPurchaseList + '\n' + '\n\n--- HG6 Inventory At a Glance --- \n' + dictToStringTable(lowStockItems) + '\n------------------------------------------- \nThanks! \n \nICF BOT \n(This is an automated message)';
  var subject = 'Inventory Low';

  MailApp.sendEmail(emailAddress, subject, message);
}

// (IN THE WORKS?) FUNCTIONALITY: COLLECT FORM RESPONSES (CHANGE: SUBMIT COUNTS OF FACILITY STOCK)
function getFormResponses()
{

}
