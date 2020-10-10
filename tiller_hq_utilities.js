/***
 *
 * * Apache Software License 2.0
 *
 * For use with TillerHQ (tillerhq.com) Google Spreadsheets. YMMV, as it assumes certain structures, sheet names, etc.
 *
 * Three functions as of now:
 * 1. findRecurring -- scans your transactions and finds, modulo a small delta, all transactions that are roughly alike and outputs them to a separate sheet.
 *  I find this useful for finding old subscriptions I forgot to cancel or otherwise grouping my expenses.  There is probably a better way to do this, as it runs the risk of timing out.
 *
 * 2. addCategory -- I never much cared for how Tiller Categorizes compared to how Yodlee does it underneath, so this method tries to reconcile this.  Note: I believe you have
 * to invoke some incantation in Tiller to get it to show you the Yodlee categories.
 *
 * 3. addAssessment -- OK, this one is a bit wierd, but it's an early attempt to codify the automatic assessment of the value of a purchase according to "Your Money or Your Live" by Joe Dominguez and Vickie Robin
 *   It requires you to do the work to figure out your real hourly wage (which is held in the Wage Assessment sheet) and then to divide that into the transaction.  Note: I also setup Google Sheets to visually display items that take up a lot of time
 *
 * onOpen and createTrigger are like they sound and create menus in Google Sheets.
 *
 *
 * TODO:
 * 1) Update so it doesn't scan the whole sheet all the time.
 *
 *
 */


/*
Add the "time" assessment to the Transactions sheet
=E26:E/'Wage Assessment'!B$17
*/
function addAssessment(){
  var sss = SpreadsheetApp;
  var ss  = sss.getActive().getSheetByName("Transactions");
  var lastRow = ss.getLastRow();
  var range = ss.getRange("Transactions!F1:F" + lastRow);
  Logger.log("Range: " + range.getA1Notation());
  i = 2;
  while (i < lastRow){
    var next = range.getCell(i, 1);
    Logger.log("Assessment: " + next.getA1Notation());
    next.setFormula("=E" + i + "/'Wage Assessment'!B$17");
    i++;
  }
}

//
/*
Add the updated category to the Transactions sheet
=IF(ISBLANK(H$2:H), I28:I, H$2:H )
*/
function addCategory(){
  var sss = SpreadsheetApp;
  var ss  = sss.getActive().getSheetByName("Transactions");
  var lastRow = ss.getLastRow();
  var range = ss.getRange("Transactions!G1:G" + lastRow);
  Logger.log("Range: " + range.getA1Notation());
  i = 2;
  while (i < lastRow){
    var next = range.getCell(i, 1);
    Logger.log("Category: " + next.getA1Notation());
    next.setFormula("=IF(ISBLANK(H$"+ i + "), I"+i+", H$"+i+")");
    i++;
  }
}

/*
Iterate through the transactions and find (near) duplicates
*/
function findRecurring() {
  var start = new Date();
  var ui = SpreadsheetApp.getUi();
  var sss = SpreadsheetApp;
  var ss  = sss.getActive().getSheetByName("Transactions");
  var range = ss.getDataRange();
  var vals = range.getValues();
  var dupes = {};
  for (i = 0; i < vals.length; i++){
    var orig = vals[i][4];
    var category = vals[i][6];
    var amount = Math.floor(orig) - (Math.floor(orig) % 2); //change this line to change the bucket size
    if (amount < 0){//this means it's a debit
      //first group by category
      var amountMap = dupes[category];
      if (!amountMap){//we've never seen this category before
        amountMap = {};
        dupes[category] = amountMap; // a list of maps of lists
      } else{
        //Logger.log("List: " + list.length);
      }
      //see if we've seen this amount
      var amountList = amountMap[amount];
      if (!amountList){
        amountList = [];
        amountMap[amount] = amountList;
      }
      amountList.push(vals[i][1] + "_" + vals[i][6] + "_" + vals[i][0] + "_" + orig);
    }
  }
  saveRecurring(dupes);
  var end = new Date();
  Logger.log("Execution time: " + (end - start));
}

/*
Write out the map of recurring items to a sheet
*/
function saveRecurring(recurring) {
  var ss  = SpreadsheetApp.getActive().getSheetByName("Recurrent Expenses");

  if (ss){
    ss.clear();
    for (var category in recurring){//category, then amount, then list
      var amountMap = recurring[category];
      if (amountMap){
        Logger.log("\nAppending Category: " + category + " AM: " + amountMap);
        for (var amount in amountMap){//for each amount in the category, append a row
          var appended = [category];
          Logger.log("Amount: " + amount);
          var amountList = amountMap[amount];
          appended.push(amount);
          Logger.log("Item app: " + amountList)
          if (amountList){
            appended.push(amountList.length);
            for (item in amountList){
              appended.push(amountList[item]);
            }
          }
          Logger.log(appended);
          var start = new Date();
          ss.appendRow(appended);
          var end = new Date();
          Logger.log("Insert time: " + (end - start));
        }

      }
    }
    ss.sort(1);
  } else {
    Logger.log("Can't find Sheet");
  }

}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Grant's Scripts")
      .addItem('Find Recurring', 'findRecurring')
      .addItem('Add Assessment', 'addAssessment')
      .addItem('Add Category', 'addCategory')
      .addToUi();
}

function createTrigger(){
  ScriptApp.newTrigger("addAssessment").timeBased().atHour(1).everyDays(1).create();
  ScriptApp.newTrigger("addCategory").timeBased().atHour(2).everyDays(1).create();
}