// 
// CURRENTLY USED
var core_tracker = SpreadsheetApp.openById("1_P9DiCAdwda0NSmDR6kpSGoVIg2zQfp2xLtqwGa6WhE");
// 
//

// OLD TRACKER
var tracker = SpreadsheetApp.openById("1Lbsn1qp7OLu-CaHlvuVQOYV2C7Xeb-x8dcFfCO88khY");

var core_old_tracker = SpreadsheetApp.openById("1wg2v8VaL_EBLwnybTrNccnZDQ2bE2Cq_iPXjR2zlWVk");
//Stuck Cases Tracker 
//var tracker = SpreadsheetApp.openById("1J9LXHaKwG3MEdphGw55Jtg5kWyj4OdYZUPmMh9Oprro");
var duplicate = core_tracker.getSheetByName("Duplicate surveys");

// Old BLS 2 Tracker
//var bls2 = SpreadsheetApp.openById("1xdBccZfcZfcuBsuHp5-aqAPMF9dYZbFKldBhww9EGSE");
// New BLS 2 Tracker
var bls2 = SpreadsheetApp.openById("1Lbsn1qp7OLu-CaHlvuVQOYV2C7Xeb-x8dcFfCO88khY");
// NEW TRACKER
//var tracker = SpreadsheetApp.openById("1J9LXHaKwG3MEdphGw55Jtg5kWyj4OdYZUPmMh9Oprro");
var all = tracker.getSheetByName("All BLS1 Surveys");
var all_bls2 = tracker.getSheetByName("All BLS2 Surveys");
var archive = tracker.getSheetByName("Archived BLS1 - June 2019");
var archive2 = tracker.getSheetByName("Archived BLS2 - June 2019");
var ind = tracker.getSheetByName("for_script");
var outbound = core_tracker.getSheetByName("Outbound Emails");
var outbound_new = core_old_tracker.getSheetByName("Outbound Emails");
var bls2_sh = bls2.getSheetByName("BLS2");


var all_core = core_tracker.getSheetByName("All Surveys");
var all_crust = core_tracker.getSheetByName("Crust Surveys");

// 
// CURRENTLY USED
var all_ewoq = core_tracker.getSheetByName("All Cases");
var all_ewoq_ne = core_tracker.getSheetByName("Assigned Cases");
// 
// 

var archive_core = core_tracker.getSheetByName("Archived CORE - July 2019");
var archive_core_august = core_tracker.getSheetByName("Archived CORE - August 2019");
var archive_old_core = core_old_tracker.getSheetByName("Archived CORE - July 2019");
var archive_old_core = core_old_tracker.getSheetByName("Archived CORE - August 2019");

var email = Session.getActiveUser().getEmail(),
  ldap = email.split("@")[0];


/////////////////
var FILEID = "1rEZga1JFOorWs5JGgbd5Gd3jmxwEQrgJ3rI_IjKCqhk";
function search(reference) {
  var file = SpreadsheetApp.openById(FILEID);
  var sheet = file.getSheetByName("Sheet1");
  var data = sheet.getDataRange().getValues();

  var result = [];
  for (var i = 0; i < data.length; i++) {
    if (similarity(data[i][2].toString().trim(), reference.toString().trim()) >= 0.90 || similarity(data[i][3].toString().trim(), reference.toString().trim()) >= 0.90) {
      result.push([data[i][0], data[i][2], data[i][3], data[i][4], data[i][5], data[i][6], data[i][7], data[i][8], data[i][9], data[i][10]]);
    }
  }

  return result
}


//EXTRA -------------------------------------------------------
function similarity(s1, s2) {
  var longer = s1;
  var shorter = s2;
  if (s1.length < s2.length) {
    longer = s2;
    shorter = s1;
  }
  var longerLength = longer.length;
  if (longerLength == 0) {
    return 1.0;
  }
  return (longerLength - editDistance(longer, shorter)) / parseFloat(longerLength);
}

function editDistance(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();

  var costs = new Array();
  for (var i = 0; i <= s1.length; i++) {
    var lastValue = i;
    for (var j = 0; j <= s2.length; j++) {
      if (i == 0)
        costs[j] = j;
      else {
        if (j > 0) {
          var newValue = costs[j - 1];
          if (s1.charAt(i - 1) != s2.charAt(j - 1))
            newValue = Math.min(Math.min(newValue, lastValue),
              costs[j]) + 1;
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0)
      costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}
///////////

function userProfile() {
  return (email);
}

function auto_archive() {
  for (var i = all_core.getLastRow() + 1; i >= 0; i--) {
    if (all_core.getRange(i, 2).getValue() == "reviewed" || all_core.getRange(i, 2).getValue() == "force-stopped") {
      all_core.getRange(i, 1, 1, 22).copyTo(archive_core.getRange(archive_core.getLastRow() + 1, 1, 1, 22));
      all_core.getRange(i, 1, 1, 22).copyTo(archive_old_core.getRange(archive_old_core.getLastRow() + 1, 1, 1, 22));
      all_core.deleteRow(i);
    }
  }
}

function auto_archive2() {
  for (var i = all_bls2.getLastRow() + 1; i >= 0; i--) {
    if (all_bls2.getRange(i, 8).getValue() == "reviewed" || all_bls2.getRange(i, 8).getValue() == "force-stopped") {
      all_bls2.getRange(i, 1, 1, 17).copyTo(archive2.getRange(archive2.getLastRow() + 1, 1, 1, 17));
      all_bls2.deleteRow(i);
    }
  }
}

function get_entity_id(parameters) {
  for (var x = ind.getRange(1, 1).getValue(); x < all.getLastRow() + 1; x++) {
    var name = all.getRange(x, 5).getValue();
    var re = /(?=(\d{8}))/g;
    var m, id = "";
    if ((m = re.exec(name)) != null) {
      all.getRange(x, 3).setValue(m[1]);

    }
  }
  ind.getRange(1, 1).setValue(all.getLastRow());
}

//Function for including external files into html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

function doGet(e) {
  Logger.log(Utilities.jsonStringify(e));
  if (!e.parameter.page) {
    // When no specific page requested, return "home page"
    if (e.parameter.caseID != null) {
      var htmlTemplate = HtmlService.createTemplateFromFile('response');
    }
    else
      var htmlTemplate = HtmlService.createTemplateFromFile('index');

    var htmlOutput = htmlTemplate.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    // appendDataToHtmlOutput modifies the html and returns the same htmlOutput object
    return htmlOutput;
  }
  // else, use page parameter to pick an html file from the script
  return HtmlService.createTemplateFromFile(e.parameter['page']).evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function decision_status_checker() {
  //checks other study ID on the list if processed/pending
}

function flashreport() {
  var list = all_core.getDataRange().getValues();
  var listcrust = all_crust.getDataRange().getValues();
  var pending = 0, ongoing = 0, reviewed = 0, all = 0, en = 0, nonen = 0, assigned_survey = "", min, sec, crust = 0;
  for (var i = 0; i < list.length; i++) {
    if (list[i][1] == "needs info" || list[i][1] == "unassigned" || list[i][1] == "pending") {
      if (list[i][8].indexOf("en") != -1) {
        en += 1;
      }
      else
        nonen += 1;
      pending += 1;
    }
    else if (list[i][1] == "decided" || list[i][1] == "ongoing") {
      ongoing += 1;
    }
    else if (list[i][1] == "reviewed" || list[i][1] == "force-stopped") {
      reviewed += 1;
    }
    all += 1;
  }
  for (var i = 0; i < listcrust.length; i++) {
    if (listcrust[i][1] == "needs info" || listcrust[i][1] == "unassigned" || listcrust[i][1] == "pending" || listcrust[i][1] == "ongoing" || listcrust[i][1] == "reviewed") {
      crust += 1;
    }
  }
  all -= 1;
  return ({
    pending: pending,
    ongoing: ongoing,
    reviewed: reviewed,
    processed: all,
    en: en,
    nonen: crust,
  });
}

function pushsurvey() {
  var list = all_core.getDataRange().getValues();
  var assigned_survey = "", min, sec, title, creator, ind_ = 0, queue, needs_info = "false";
  for (var i = 0; i < list.length; i++) {
    if (list[i][0] == email.split("@")[0] && (list[i][1] == "ongoing" || list[i][1] == "decided")) {
      ind_ = 1;
      assigned_survey = list[i][2];
      var date = new Date();
      var aht = Math.abs(new Date(date) - new Date(list[i][14]));
      var min = Math.floor((aht / 1000) / 60);
      var sec = Math.floor((aht / 1000) % 60);
      title = list[i][5];
      creator = list[i][6];
      queue = "core";
      if (list[i][13] == "needs info") needs_info = "true";
    }
  }
  list = all_crust.getDataRange().getValues();
  for (var i = 0; i < list.length; i++) {
    if (list[i][0] == email.split("@")[0] && (list[i][1] == "ongoing" || list[i][1] == "decided")) {
      ind_ = 1;
      assigned_survey = list[i][2];
      var date = new Date();
      var aht = Math.abs(new Date(date) - new Date(list[i][14]));
      var min = Math.floor((aht / 1000) / 60);
      var sec = Math.floor((aht / 1000) % 60);
      title = list[i][5];
      creator = list[i][6];
      queue = "crust";
      if (list[i][13] == "needs info") needs_info = "true";
    }
  }



  return ({
    assigned_survey: assigned_survey,
    aht_min: min,
    aht_sec: sec,
    title: title,
    creator: creator,
    queue: queue,
    needs_info: needs_info
  });
}

function get_notifications() {
  //   var list = all.getDataRange().getValues();
  //   list.reverse();
  //   var final = [];
  //   let x = 1;
  //   for(var i = 0; i < list.length; i++) {
  //     if(list[i][9] == "pending" || list[i][9] == "ongoing"|| list[i][9] == "decided"|| list[i][9] == "unassigned") {
  //       final.push([x,list[i][1],list[i][2],list[i][6],list[i][0],list[i][9],list[i][8]]);
  //       x++;
  //     }
  //   }
  //   return (JSON.stringify(final));
  const notif_tracker = SpreadsheetApp.openById("1mMMC-oOALm5shbdljtEdOb9XDciw72qu2z-HxZiH-iM");
  notif = notif_tracker.getSheetByName("notification");

  var list = notif.getDataRange().getValues();
  final = [];
  x = 1;
  Logger.log(list.length);
  for (i = 0; i < list.length; i++) {
    if (list[i][0] != "") {
      final.push([x, list[i][0], list[i][1], list[i][2], list[i][3], list[i][4], list[i][5]]);
      x++
    }
  }
  return (JSON.stringify(final));
}

function get_bls1() {
  var list = all.getDataRange().getValues();
  list.reverse();
  var final = [];
  var x = 1;
  for (var i = 0; i < list.length; i++) {
    if (list[i][9] == "pending" || list[i][9] == "ongoing" || list[i][9] == "decided" || list[i][9] == "unassigned") {
      final.push([x, list[i][1], list[i][2], list[i][6], list[i][0], list[i][9], list[i][8]]);
      x++;
    }
  }
  return (JSON.stringify(final));
}


function getCases() {
  var list = all_ewoq.getDataRange().getValues();
  list.shift();
  list.reverse();
  return (JSON.stringify(list));
}

function getNICases() {
  var list = all_ewoq_ne.getDataRange().getValues();
  list.shift();
  list.reverse();
  return (JSON.stringify(list));
}

function dynamic_categories(queue) {
  var tracker = SpreadsheetApp.openById("1WGFeTH-VVuY_lCpI9BoHe9N8mTHYQWFULGnzWwCNKco");
  var category = tracker.getSheetByName("Categories");
  var category_list = [], column = 1;
  if (queue == "Survey Review") column = 1;
  if (queue == "Support") column = 2;
  if (queue == "Publisher") column = 3;
  if (queue == "Partnership") column = 4;
  var list = category.getRange(1, column, category.getLastRow()).getValues();
  for (var i = 0; i < list.length; i++) {
    if (list[i] != "")
      category_list[i] = list[i];
  }

  return (category_list);
}

function get_bls2() {
  var list = all_bls2.getDataRange().getValues();
  var final = [];
  var x = 1;

  var filtered = data.filter(function (dataRow) {
    return dataRow[7] == 'pending' || dataRow[7] == 'ongoing' || dataRow[7] == 'decided' || dataRow[7] == 'unassigned';
  });
  //  for(var i = 0; i < list.length; i++) {
  //  if(list[i][7] == "pending" || list[i][7] == "ongoing"|| list[i][7] == "decided"|| list[i][7] == "unassigned") {
  //      final.push([x,list[i][4],list[i][3],list[i][0],list[i][7]]);
  //      x++;
  //    }
  //  }
  return (JSON.stringify(final));
}

function get_core() {
  var list = all_core.getDataRange().getValues();
  var final = [];
  var x = 1;

  var filtered = list.filter(function (dataRow) {
    return dataRow[1] == 'pending' || dataRow[1] == 'ongoing' || dataRow[1] == 'decided' || dataRow[1] == 'unassigned';
  });

  for (var i = 0; i < filtered.length; i++) {
    final.push([x, filtered[i][2], get_TAT(filtered[i][4]), filtered[i][5], filtered[i][8], filtered[i][0], filtered[i][1], filtered[i][13]]);
    x++;
  }
  return (JSON.stringify(final));
}



function get_crust() {
  var list = all_crust.getDataRange().getValues();
  var final = [];
  var x = 1;
  for (var i = 0; i < list.length; i++) {
    if (list[i][1] == "pending" || list[i][1] == "ongoing" || list[i][1] == "decided" || list[i][1] == "unassigned") {
      final.push([x, list[i][2], get_TAT_crust(list[i][4]), list[i][5], "_ANY_", list[i][0], list[i][1], list[i][13]]);
      x++;
    }
  }
  return (JSON.stringify(final));
}


function dateAdd(date, interval, units) {
  var ret = new Date(date); //don't change original date
  var checkRollover = function () { if (ret.getDate() != date.getDate()) ret.setDate(0); };
  switch (interval.toLowerCase()) {
    case 'year': ret.setFullYear(ret.getFullYear() + units); checkRollover(); break;
    case 'quarter': ret.setMonth(ret.getMonth() + 3 * units); checkRollover(); break;
    case 'month': ret.setMonth(ret.getMonth() + units); checkRollover(); break;
    case 'week': ret.setDate(ret.getDate() + 7 * units); break;
    case 'day': ret.setDate(ret.getDate() + units); break;
    case 'hour': ret.setTime(ret.getTime() + units * 3600000); break;
    case 'minute': ret.setTime(ret.getTime() + units * 60000); break;
    case 'second': ret.setTime(ret.getTime() + units * 1000); break;
    default: ret = undefined; break;
  }
  return ret;
}

function saveChanges(parameters) {
  var columnValues = all_ewoq.getRange(2, 3, all_ewoq.getLastRow()).getValues();
  var searchString = parameters.ref;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all_ewoq.getRange(searchResult + 2, 13).setValue(parameters.qaAuditor);
    all_ewoq.getRange(searchResult + 2, 14).setValue(parameters.qaScore);
    all_ewoq.getRange(searchResult + 2, 15).setValue(parameters.qaComment);
    all_ewoq.getRange(searchResult + 2, 16).setValue(parameters.qaVariance);
  }
  return ("success");
}

function get_TAT(sent) {
  var date = new Date();
  var current = date.getFullYear() + '-' + (date.getMonth() + 1) + "-" + date.getDate();
  current = current + " " + date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds();
  var snt = dateAdd(new Date(sent), 'hour', 15);
  var tat = Math.abs(new Date(current) - snt);
  var tat_days = Math.floor((tat / 1000) / 60 / 60 / 24);
  var tat_hours = Math.floor((tat / 1000) / 60 / 60);
  var tat_mins = Math.floor((tat / 1000) / 60);
  var actual_tat = "";
  if (tat_mins == 0) actual_tat = "1 hour";
  else if (tat_mins <= 60) {
    actual_tat = (60 - tat_mins) + " mins";
  }
  else if (tat_mins > 60 && tat_mins < 120) {
    actual_tat = (tat_mins - 60) + " mins ago"
  }
  else if (tat_mins == 120) {
    actual_tat = "1 hour ago";
  }
  else if (tat_mins > 120 && tat_days < 1) {
    actual_tat = tat_hours + " hours ago";
  }
  else if (tat_hours >= 24) {
    actual_tat = tat_days + " days " + (tat_hours - tat_days * 24) + " hours ago";
  }
  return actual_tat;
}


function get_TAT_crust(sent) {
  var date = new Date();
  var current = date.getFullYear() + '-' + (date.getMonth() + 1) + "-" + date.getDate();
  current = current + " " + date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds();
  var tat = Math.abs(new Date(current) - sent);
  var tat_days = Math.floor((tat / 1000) / 60 / 60 / 24);
  var tat_hours = Math.floor((tat / 1000) / 60 / 60);
  var tat_mins = Math.floor((tat / 1000) / 60);
  var actual_tat = "";
  if (tat_mins == 0) actual_tat = "1 hour";
  else if (tat_mins <= 60) {
    actual_tat = (60 - tat_mins) + " mins";
  }
  else if (tat_mins > 60 && tat_mins < 120) {
    actual_tat = (tat_mins - 60) + " mins ago"
  }
  else if (tat_mins == 120) {
    actual_tat = "1 hour ago";
  }
  else if (tat_mins > 120 && tat_days < 1) {
    actual_tat = tat_hours + " hours ago";
  }
  else if (tat_hours >= 24) {
    actual_tat = tat_days + " days " + (tat_hours - tat_days * 24) + " hours ago";
  }
  return actual_tat;
}

function assign_bls1(surveys) {
  var date = new Date();
  var column = 2;
  var columnValues = all.getRange(2, 2, all.getLastRow()).getValues();

  for (var i = 0; i < surveys.length; i++) {
    var searchString = surveys[i];
    var searchResult = columnValues.findIndex(searchString);

    if (searchResult != -1) {
      all.getRange(searchResult + 2, 1).setValue(email.split("@")[0]);
      all.getRange(searchResult + 2, 10).setValue("ongoing");
      // Check if needs info
      // Don't overwrite date if needs info
      if (all.getRange(searchResult + 2, 9).getValue() == 'needs info') {
        // all.getRange(searchResult + 2, 11).setValue('eyy');
      } else {
        all.getRange(searchResult + 2, 11).setValue(date);
      }

    }
    else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}

function setNeedsInfo(parameters) {
  all_ewoq_ne.appendRow([email.split("@")[0], "needsinfo", guid(), parameters["cases"], "", "", "", "", "", "", "", "", "", "", new Date(parameters["start_time"]), new Date(parameters["end_time"])]);

  return ("success");
}

function assign_bls2(surveys) {
  var date = new Date();
  var column = 2;
  var columnValues = all_bls2.getRange(2, 5, all_bls2.getLastRow()).getValues();

  for (var i = 0; i < surveys.length; i++) {
    var searchString = surveys[i];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      all_bls2.getRange(searchResult + 2, 1).setValue(email.split("@")[0]);
      all_bls2.getRange(searchResult + 2, 8).setValue("ongoing");
      // Check if needs info
      // Don't overwrite date if needs info
      if (all_bls2.getRange(searchResult + 2, 9).getValue() == 'needs info') {
        // all.getRange(searchResult + 2, 11).setValue('eyy');
      } else {
        all_bls2.getRange(searchResult + 2, 14).setValue(date);
      }

    }
    else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}

var bypass = "";
function assign_core(surveys, bypass) {
  var date = new Date();
  var column = 3;
  var needs_info = "false";
  var columnValues = all_core.getRange(2, column, all_core.getLastRow()).getValues();
  var title = "", creator = "", fifo_alert = 0;
  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    if ((all_core.getRange(searchResult + 2, 2).getValue() == "ongoing" || all_core.getRange(searchResult + 2, 2).getValue() == "decided") && all_core.getRange(searchResult + 2, 1).getValue() != email.split("@")[0]) {
      return ({ response: "Survey already taken.", title: "", creator: "", needs_info: "false" });
    }
    else if (all_core.getRange(searchResult + 2, 2).getValue() == "reviewed" || all_core.getRange(searchResult + 2, 2).getValue() == "force-stopped") {
      return ({ response: "Survey already finished.", title: "", creator: "", needs_info: "false" });
    } else {
      if (searchResult + 1 == 1 || all_core.getRange(searchResult + 2, 9).getValue().indexOf("en") == -1) {
        fifo_alert = 0;
        all_core.getRange(searchResult + 2, 1).setValue(email.split("@")[0]);
        all_core.getRange(searchResult + 2, 2).setValue("ongoing");
        title = all_core.getRange(searchResult + 2, 6).getValue();
        creator = all_core.getRange(searchResult + 2, 7).getValue();
        // Check if needs info
        // Don't overwrite date if needs info
        if (all_core.getRange(searchResult + 2, 14).getValue() == 'needs info') {
          // do something
          needs_info = "true";
        } else {
          all_core.getRange(searchResult + 2, 15).setValue(date);
        }

      } else {
        if (all_core.getRange(searchResult + 1, 2).getValue() == "pending" && bypass == "" && all_core.getRange(searchResult + 1, 9).getValue().indexOf("en") != -1) {
          fifo_alert = 1;
        } else if (all_core.getRange(searchResult + 1, 2).getValue() == "pending" && bypass == "true") {
          fifo_alert = 0;
          all_core.getRange(searchResult + 2, 1).setValue(email.split("@")[0]);
          all_core.getRange(searchResult + 2, 2).setValue("ongoing");
          title = all_core.getRange(searchResult + 2, 6).getValue();
          creator = all_core.getRange(searchResult + 2, 7).getValue();
          // Check if needs info
          // Don't overwrite date if needs info
          if (all_core.getRange(searchResult + 2, 14).getValue() == 'needs info') {
            needs_info = "true";
            // all.getRange(searchResult + 2, 11).setValue('eyy');
          } else {
            all_core.getRange(searchResult + 2, 15).setValue(date);
          }
        } else {
          fifo_alert = 0;
          all_core.getRange(searchResult + 2, 1).setValue(email.split("@")[0]);
          all_core.getRange(searchResult + 2, 2).setValue("ongoing");
          title = all_core.getRange(searchResult + 2, 6).getValue();
          creator = all_core.getRange(searchResult + 2, 7).getValue();
          // Check if needs info
          // Don't overwrite date if needs info
          if (all_core.getRange(searchResult + 2, 14).getValue() == 'needs info') {
            needs_info = "true";
            // all.getRange(searchResult + 2, 11).setValue('eyy');
          } else {
            all_core.getRange(searchResult + 2, 15).setValue(date);
          }

        }
      }
    }
  }
  else {
    return ({ response: surveys + " not found.", "title": "", "creator": "", needs_info: "false" });
  }
  return ({ response: "success", title: title, creator: creator, fifo_alert: fifo_alert, needs_info: needs_info });

}

function assign_crust(surveys, bypass) {
  var date = new Date();
  var column = 3;
  var columnValues = all_crust.getRange(2, column, all_crust.getLastRow()).getValues();
  var title = "", creator = "", fifo_alert = 0;
  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    if (all_crust.getRange(searchResult + 2, 2).getValue() == "ongoing" || all_crust.getRange(searchResult + 2, 2).getValue() == "decided") {
      return ({ response: "Survey already taken.", title: "", creator: "" });
    }
    else if (all_crust.getRange(searchResult + 2, 2).getValue() == "reviewed" || all_crust.getRange(searchResult + 2, 2).getValue() == "force-stopped") {
      return ({ response: "Survey already finished.", title: "", creator: "" });
    } else {
      if (searchResult + 1 == 1 || all_crust.getRange(searchResult + 2, 9).getValue().indexOf("en") == -1) {
        fifo_alert = 0;
        all_crust.getRange(searchResult + 2, 1).setValue(email.split("@")[0]);
        all_crust.getRange(searchResult + 2, 2).setValue("ongoing");
        title = all_crust.getRange(searchResult + 2, 6).getValue();
        creator = all_crust.getRange(searchResult + 2, 7).getValue();
        // Check if needs info
        // Don't overwrite date if needs info
        if (all_crust.getRange(searchResult + 2, 14).getValue() == 'needs info') {
          // do something
        } else {
          all_crust.getRange(searchResult + 2, 15).setValue(date);
        }

      } else {
        if (all_crust.getRange(searchResult + 1, 2).getValue() == "pending" && bypass == "" && all_crust.getRange(searchResult + 1, 9).getValue().indexOf("en") != -1) {
          fifo_alert = 1;
        } else if (all_crust.getRange(searchResult + 1, 2).getValue() == "pending" && bypass == "true") {
          fifo_alert = 0;
          all_crust.getRange(searchResult + 2, 1).setValue(email.split("@")[0]);
          all_crust.getRange(searchResult + 2, 2).setValue("ongoing");
          title = all_crust.getRange(searchResult + 2, 6).getValue();
          creator = all_crust.getRange(searchResult + 2, 7).getValue();
          // Check if needs info
          // Don't overwrite date if needs info
          if (all_crust.getRange(searchResult + 2, 14).getValue() == 'needs info') {
            // all.getRange(searchResult + 2, 11).setValue('eyy');
          } else {
            all_crust.getRange(searchResult + 2, 15).setValue(date);
          }
        } else {
          fifo_alert = 0;
          all_crust.getRange(searchResult + 2, 1).setValue(email.split("@")[0]);
          all_crust.getRange(searchResult + 2, 2).setValue("ongoing");
          title = all_crust.getRange(searchResult + 2, 6).getValue();
          creator = all_crust.getRange(searchResult + 2, 7).getValue();
          // Check if needs info
          // Don't overwrite date if needs info
          if (all_crust.getRange(searchResult + 2, 14).getValue() == 'needs info') {
            // all.getRange(searchResult + 2, 11).setValue('eyy');
          } else {
            all_crust.getRange(searchResult + 2, 15).setValue(date);
          }

        }
      }
    }
  }
  else {
    return ({ response: surveys + " not found.", "title": "", "creator": "" });
  }
  return ({ response: "success", title: title, creator: creator, fifo_alert: fifo_alert });

}

function unassign_bls1(surveys) {
  var date = new Date();
  var column = 2;
  var columnValues = all.getRange(2, 2, all.getLastRow()).getValues();

  for (var i = 0; i < surveys.length; i++) {
    var searchString = surveys[i];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      all.getRange(searchResult + 2, 10).setValue("unassigned");
      all.getRange(searchResult + 2, 11).setValue(date);
    }
    else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}


function assignNewCases(parameter){
  let payload = [],
      data = JSON.parse(parameter);
  data.forEach((submittedCase) => {
      payload[0] = submittedCase.ldap;
      payload[1] = guid();
      payload[2] = submittedCase.url;
      Logger.log("payload: " + payload);
      all_ewoq_ne.appendRow(payload);
  });
  
}


// ACTUAL FINISH FUNCTION
function removeNE(parameter) {
  
  parameter = JSON.parse(parameter);
//  Change search string to parameter.refID
  let searchString = parameter.refID,
    startRow = 1,
    refIDCol = 2,
    allData = all_ewoq_ne.getDataRange().getValues();

  allData.shift();
  
  var columnValues = all_ewoq_ne.getRange(startRow, refIDCol, all_ewoq_ne.getLastRow()).getValues();
  // Finds index 1st parameter = target, 2nd parameter = string to be searched
  var searchResult = find_index(columnValues, searchString);
  Logger.log("ad: " + allData);
  Logger.log("cV: " + columnValues);
  Logger.log("sR: " + searchResult);
  
  
//  Get data for append row
  let data = allData[(searchResult - 1)];
//  data.shift();
  Logger.log(data);
  
  data[0] = parameter.ldap;
  data[1] = "reviewed";
  data[2] = parameter.refID;
  data[3] = parameter.surveys;
  data[4] = parameter.language;
  data[5] = parameter.surveytype;
  data[6] = parameter.screenshot;
  data[7] = parameter.decision;
  data[8] = parameter.start_time;
  data[9] = parameter.end_time;
  data[10] = parameter.categories;
  data[11] = parameter.aht;

  if (searchResult != -1) {
    all_ewoq.appendRow(data);
    // all_ewoq.appendRow(JSON.stringify(parameter));



    // Delete row using Case ID from Assigned Tab
    all_ewoq_ne.deleteRow(searchResult + 1);
  }

  return "success";


}

function getShit() {
  let caseID = "case-5e03-62a2",
    startRow = 1,
    refIDCol = 2;


  var columnValues = all_ewoq_ne
    .getRange(startRow, refIDCol, all_ewoq_ne.getLastRow())
    .getValues();



  // Finds index 1st parameter = target, 2nd parameter = string to be searched
  var rowIndex = find_index(columnValues, caseID);

 let fuck = all_ewoq_ne.getRange(startRow, 1, all_ewoq_ne.getLastRow()).getValues();
//  let fuck = all_ewoq_ne.getDataRange().getValues();

  let allData = all_ewoq_ne.getDataRange().getValues();
  allData.shift();
  
  Logger.log("fuck: "+ fuck);
  Logger.log("allData: " + allData);
  Logger.log("col: " + columnValues + " " + rowIndex + rowIndex);
  
}

function testShit() {
  var list = all_ewoq_ne.getDataRange().getValues();
  // return (JSON.stringify(list));
  return JSON.stringify(list);
}

Logger.log(testShit());



function unassign_bls2(surveys) {
  var date = new Date();
  var column = 2;
  var columnValues = all_bls2.getRange(2, 5, all_bls2.getLastRow()).getValues();

  for (var i = 0; i < surveys.length; i++) {
    var searchString = surveys[i];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      all_bls2.getRange(searchResult + 2, 8).setValue("unassigned");
      all_bls2.getRange(searchResult + 2, 14).setValue(date);
    }
    else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}



function unassign_core(surveys) {
  var date = new Date();
  var column = 3;
  var columnValues = all_core.getRange(2, column, all_core.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all_core.getRange(searchResult + 2, 2).setValue("unassigned");
    all_core.getRange(searchResult + 2, 16).setValue(date);
  }
  else {
    return (surveys[i] + " not found.");
  }
  return ("success");

}

function unassign_crust(surveys) {
  var date = new Date();
  var column = 3;
  var columnValues = all_crust.getRange(2, column, all_crust.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all_crust.getRange(searchResult + 2, 2).setValue("unassigned");
    all_crust.getRange(searchResult + 2, 16).setValue(date);
  }
  else {
    return (surveys[i] + " not found.");
  }
  return ("success");

}

function compliant_bls1(surveys) {
  var date = new Date();
  var column = 2;
  var columnValues = all.getRange(2, 2, all.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    // Check if needs info
    // Don't overwrite date if needs info
    all.getRange(searchResult + 2, 10).setValue("decided");
    all.getRange(searchResult + 2, 9).setValue("compliant");
  }
  else {
    return (surveys + " not found.");
  }
  return ("success");

}

var action = CardService.newAction().setFunctionName("notificationCallback");
CardService.newTextButton().setText('Save').setOnClickAction(action);

// ...

function notificationCallback() {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText("Some info to display to user"))
    .build();
}

function markreviewed_core(surveys) {
  var date = new Date();
  var column = 3;
  var columnValues = all_core.getRange(2, column, all_core.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    // Check if needs info
    // Don't overwrite date if needs info
    all_core.getRange(searchResult + 2, 2).setValue("decided");
    all_core.getRange(searchResult + 2, 14).setValue("mark_reviewed");
  }
  else {
    return (surveys + " not found.");
  }
  return ("success");

}

function markreviewed_crust(surveys) {
  var date = new Date();
  var column = 3;
  var columnValues = all_crust.getRange(2, column, all_crust.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    // Check if needs info
    // Don't overwrite date if needs info
    if (all_crust.getRange(searchResult + 2, 14).getValue() == 'needs info') {
      // all.getRange(searchResult + 2, 11).setValue('eyy');
    } else {
      all_crust.getRange(searchResult + 2, 16).setValue(date);
    }
    all_crust.getRange(searchResult + 2, 2).setValue("decided");
    all_crust.getRange(searchResult + 2, 14).setValue("mark_reviewed");
  }
  else {
    return (surveys + " not found.");
  }
  return ("success");

}

function compliant_bls2(surveys) {
  var date = new Date();
  var column = 2;
  var columnValues = all_bls2.getRange(2, 5, all_bls2.getLastRow()).getValues();

  for (var i = 0; i < surveys.length; i++) {
    var searchString = surveys[i];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      // Check if needs info
      // Don't overwrite date if needs info
      // all_bls2.getRange(searchResult + 2, 14).setValue(date);
      all_bls2.getRange(searchResult + 2, 8).setValue("decided");
      all_bls2.getRange(searchResult + 2, 9).setValue("compliant");
      all_bls2.getRange(searchResult + 2, 15).setValue(date);
    }
    else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}



function compliant_core(surveys) {
  var date = new Date();
  var column = 3;
  var columnValues = all_core.getRange(2, column, all_core.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    // Check if needs info
    // Don't overwrite date if needs info
    all_core.getRange(searchResult + 2, 2).setValue("decided");
    all_core.getRange(searchResult + 2, 14).setValue("compliant");
  }
  else {
    return (surveys + " not found.");
  }
  return ("success");

}

function compliant_crust(surveys) {
  var date = new Date();
  var column = 3;
  var columnValues = all_crust.getRange(2, column, all_crust.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    // Check if needs info
    // Don't overwrite date if needs info
    if (all_crust.getRange(searchResult + 2, 14).getValue() == 'needs info') {
      // all.getRange(searchResult + 2, 11).setValue('eyy');
    } else {
      all_crust.getRange(searchResult + 2, 16).setValue(date);
    }
    all_crust.getRange(searchResult + 2, 2).setValue("decided");
    all_crust.getRange(searchResult + 2, 14).setValue("compliant");
  }
  else {
    return (surveys + " not found.");
  }
  return ("success");

}

function noncompliant_bls1(surveys) {
  var date = new Date();
  var column = 2;
  var columnValues = all.getRange(2, 2, all.getLastRow()).getValues();

  for (var i = 0; i < surveys.length; i++) {
    var searchString = surveys[i];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      all.getRange(searchResult + 2, 10).setValue("decided");
      all.getRange(searchResult + 2, 9).setValue("noncompliant");
      all.getRange(searchResult + 2, 12).setValue(date);
    }
    else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}



function noncompliant_bls2(surveys) {
  var date = new Date();
  var column = 2;
  var columnValues = all_bls2.getRange(2, 5, all_bls2.getLastRow()).getValues();

  for (var i = 0; i < surveys.length; i++) {
    var searchString = surveys[i];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      all_bls2.getRange(searchResult + 2, 8).setValue("decided");
      all_bls2.getRange(searchResult + 2, 9).setValue("noncompliant");
      all_bls2.getRange(searchResult + 2, 15).setValue(date);
    }
    else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}


function noncompliant_core(surveys) {
  var date = new Date();
  var column = 3;
  var columnValues = all_core.getRange(2, column, all_core.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all_core.getRange(searchResult + 2, 2).setValue("decided");
    all_core.getRange(searchResult + 2, 14).setValue("noncompliant");

  }
  else {
    return (surveys[i] + " not found.");
  }
  return ("success");

}

function noncompliant_crust(surveys) {
  var date = new Date();
  var column = 3;
  var columnValues = all_crust.getRange(2, column, all_crust.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all_crust.getRange(searchResult + 2, 2).setValue("decided");
    all_crust.getRange(searchResult + 2, 14).setValue("noncompliant");
    all_crust.getRange(searchResult + 2, 16).setValue(date);
  }
  else {
    return (surveys[i] + " not found.");
  }
  return ("success");

}

function forceStopped_bls1(surveys) {
  var date = new Date();
  var column = 2;
  var columnValues = all.getRange(2, 2, all.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all.getRange(searchResult + 2, 10).setValue("force-stopped");
    all.getRange(searchResult + 2, 9).setValue("force-stopped");
    all.getRange(searchResult + 2, 12).setValue(date);
  }
  else {
    return (surveys + " not found.");
  }
  return ("success");

}



function forceStopped_bls2(surveys) {
  var date = new Date();
  var column = 2;
  var columnValues = all_bls2.getRange(2, 5, all_bls2.getLastRow()).getValues();

  for (var i = 0; i < surveys.length; i++) {
    var searchString = surveys[i];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      all_bls2.getRange(searchResult + 2, 8).setValue("force-stopped");
      all_bls2.getRange(searchResult + 2, 9).setValue("force-stopped");
      all_bls2.getRange(searchResult + 2, 15).setValue(date);
    }
    else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}



function forceStopped_core(surveys) {
  var date = new Date();
  var column = 3;
  var columnValues = all_core.getRange(2, column, all_core.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all_core.getRange(searchResult + 2, 2).setValue("force-stopped");
    all_core.getRange(searchResult + 2, 14).setValue("force-stopped");
    all_core.getRange(searchResult + 2, 16).setValue(date);
  }
  else {
    return (surveys + " not found.");
  }
  return ("success");

}

function forceStopped_crust(surveys) {
  var date = new Date();
  var column = 3;
  var columnValues = all_crust.getRange(2, column, all_crust.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all_crust.getRange(searchResult + 2, 2).setValue("force-stopped");
    all_crust.getRange(searchResult + 2, 14).setValue("force-stopped");
    all_crust.getRange(searchResult + 2, 16).setValue(date);
  }
  else {
    return (surveys + " not found.");
  }
  return ("success");

}

function needsInfo_bls1(surveys) {

  var date = new Date();
  var column = 2;
  var columnValues = all.getRange(2, 2, all.getLastRow()).getValues();

  for (var i = 0; i < surveys.length; i++) {
    var searchString = surveys[i];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      all.getRange(searchResult + 2, 10).setValue("ongoing");
      all.getRange(searchResult + 2, 9).setValue("needs info");
      all.getRange(searchResult + 2, 12).setValue(date);
    } else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}


function needsInfo_core(surveys) {

  var date = new Date();
  var column = 3;
  var columnValues = all_core.getRange(2, column, all_core.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all_core.getRange(searchResult + 2, 2).setValue("ongoing");
    all_core.getRange(searchResult + 2, 14).setValue("needs info");
    all_core.getRange(searchResult + 2, 16).setValue(date);
  } else {
    return (surveys + " not found.");
  }
  return ("success");

}

function needsInfo_ewoq(surveys) {

  var date = new Date();
  var column = 3;
  var columnValues = all_ewoq.getRange(2, column, all_ewoq.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all_core.getRange(searchResult + 2, 2).setValue("");
    all_core.getRange(searchResult + 2, 14).setValue("needs info");
    all_core.getRange(searchResult + 2, 16).setValue(date);
  } else {
    return (surveys + " not found.");
  }
  return ("success");

}


function needsInfo_crust(surveys) {

  var date = new Date();
  var column = 3;
  var columnValues = all_crust.getRange(2, column, all_crust.getLastRow()).getValues();

  var searchString = surveys;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all_crust.getRange(searchResult + 2, 2).setValue("ongoing");
    all_crust.getRange(searchResult + 2, 14).setValue("needs info");
    all_crust.getRange(searchResult + 2, 16).setValue(date);
  } else {
    return (surveys + " not found.");
  }
  return ("success");

}


function needsInfo_bls2(surveys) {

  var date = new Date();
  var column = 2;
  var columnValues = all_bls2.getRange(2, 5, all_bls2.getLastRow()).getValues();

  for (var i = 0; i < surveys.length; i++) {
    var searchString = surveys[i];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      all_bls2.getRange(searchResult + 2, 8).setValue("ongoing");
      all_bls2.getRange(searchResult + 2, 9).setValue("needs info");
      all_bls2.getRange(searchResult + 2, 15).setValue(date);
    } else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}





function finish_bls1(parameters) {
  var date = new Date();
  var column = 2;
  var columnValues = all.getRange(2, 2, all.getLastRow()).getValues();

  for (var i = 0; i < parameters["surveys"].length; i++) {
    var searchString = parameters["surveys"][i];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      //   
      if (parameters["decision"] == 'force-stopped') {
        all.getRange(searchResult + 2, 10).setValue("force-stopped");
      } else {
        all.getRange(searchResult + 2, 10).setValue("reviewed");
      }

      all.getRange(searchResult + 2, 8).setValue(parameters["screenshot"]);
      all.getRange(searchResult + 2, 9).setValue(parameters["decision"]);
      all.getRange(searchResult + 2, 13).setValue(parameters["categories"]);
      all.getRange(searchResult + 2, 14).setValue(parameters["aht"]);
    }
    else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}

function finish_bls2(parameters) {
  var date = new Date();
  var column = 2;
  var columnValues = all_bls2.getRange(2, 5, all_bls2.getLastRow()).getValues();

  for (var i = 0; i < parameters["surveys"].length; i++) {
    var searchString = parameters["surveys"][i];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      //   
      if (parameters["decision"] == 'force-stopped') {
        all_bls2.getRange(searchResult + 2, 9).setValue("force-stopped");
      } else {
        all_bls2.getRange(searchResult + 2, 8).setValue("reviewed");
      }

      all_bls2.getRange(searchResult + 2, 10).setValue(parameters["violation"]);
      all_bls2.getRange(searchResult + 2, 12).setValue(parameters["screenshot2"]);
      all_bls2.getRange(searchResult + 2, 13).setValue(parameters["screenshot"]);
      all_bls2.getRange(searchResult + 2, 9).setValue(parameters["decision"]);
      all_bls2.getRange(searchResult + 2, 11).setValue(parameters["categories"]);
      all_bls2.getRange(searchResult + 2, 16).setValue(parameters["aht"]);
    }
    else {
      return (surveys[i] + " not found.");
    }
  }
  return ("success");

}

function search_survey(survey) {
  var column = 3;
  var columnValues = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(2, column, SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getLastRow()).getValues();
  var searchString = survey;
  var searchResult = find_index(columnValues, searchString);

  var ldap, refid, sent, title, creator, country, lang, times_rev;
  var surveytype, screenshot, decision, categories, tat, aht, comment_just, email, start, end, subject;

  if (searchResult != -1) {
    ldap = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 1).getValue();
    refid = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 3).getValue();
    sent = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 5).getValue().toString();
    title = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 6).getValue();
    creator = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 7).getValue();
    country = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 8).getValue();
    lang = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 9).getValue();
    times_rev = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 10).getValue();
    screenshot = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 13).getValue();
    decision = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 14).getValue();
    categories = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 17).getValue();
    aht = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 18).getValue().toString();
    surveytype = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 12).getValue();
    comment_just = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 21).getValue();
    start = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 15).getValue().toString();
    end = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 16).getValue().toString();
    if (SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getRange(searchResult + 2, 14).getValue() == "noncompliant") {
      var ne = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("needs edit email");
      var column = 1;
      var columnValues = ne.getRange(2, column, ne.getLastRow()).getValues();
      var searchString = survey;
      var searchResult = columnValues.findIndex(searchString);
      if (searchResult != -1) {
        email = ne.getRange(searchResult + 2, 8).getValue();
        subject = ne.getRange(searchResult + 2, 7).getValue();
      }
    }
  }
  else {

    var column2 = 3;
    var columnValues2 = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(2, column2, SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getLastRow()).getValues();
    var searchString2 = survey;
    var searchResult2 = find_index(columnValues2, searchString2);

    if (searchResult2 != -1) {
      ldap = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 1).getValue();
      refid = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 3).getValue();
      sent = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 5).getValue().toString();
      title = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 6).getValue();
      creator = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 7).getValue();
      country = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 8).getValue();
      lang = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 9).getValue();
      times_rev = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 10).getValue();
      screenshot = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 13).getValue();
      decision = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 14).getValue();
      categories = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 17).getValue();
      aht = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 18).getValue().toString();
      surveytype = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 12).getValue();
      comment_just = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 21).getValue();
      start = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 15).getValue().toString();
      end = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 16).getValue().toString();

      if (SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getRange(searchResult2 + 2, 14).getValue() == "noncompliant") {
        var ne = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("needs edit email");
        var column = 1;
        var columnValues = ne.getRange(2, column, ne.getLastRow()).getValues();
        var searchString = survey;
        var searchResult = columnValues.findIndex(searchString);
        if (searchResult != -1) {
          email = ne.getRange(searchResult + 2, 8).getValue();
          subject = ne.getRange(searchResult + 2, 7).getValue();
        }
      }
    } else {
      var column2 = 3;
      var columnValues2 = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(2, column2, SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getLastRow()).getValues();
      var searchString2 = survey;
      var searchResult2 = find_index(columnValues2, searchString2);

      if (searchResult2 != -1) {
        ldap = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 1).getValue();
        refid = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 3).getValue();
        sent = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 5).getValue().toString();
        title = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 6).getValue();
        creator = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 7).getValue();
        country = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 8).getValue();
        lang = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 9).getValue();
        times_rev = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 10).getValue();
        screenshot = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 13).getValue();
        decision = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 14).getValue();
        categories = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 17).getValue();
        aht = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 18).getValue().toString();
        surveytype = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 12).getValue();
        comment_just = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 21).getValue();
        start = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 15).getValue().toString();
        end = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 16).getValue().toString();

        if (SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getRange(searchResult2 + 2, 14).getValue() == "noncompliant") {
          var ne = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("needs edit email");
          var column = 1;
          var columnValues = ne.getRange(2, column, ne.getLastRow()).getValues();
          var searchString = survey;
          var searchResult = columnValues.findIndex(searchString);
          if (searchResult != -1) {
            email = ne.getRange(searchResult + 2, 8).getValue();
            subject = ne.getRange(searchResult + 2, 7).getValue();
          }
        }
      } else {
        return ({ status: survey + " not found." });
      }
    }
  }
  return ({
    ldap: ldap,
    refid: refid,
    sent: sent,
    title: title,
    creator: creator,
    country: country,
    lang: lang,
    times_rev: times_rev,
    screenshot: screenshot,
    decision: decision,
    categories: categories,
    aht: aht,
    surveytype: surveytype,
    comment_just: comment_just,
    email: email,
    subject: subject,
    start: start,
    end: end,
    surveyid: survey,
    status: "success"
  });

}

function finish_core(parameters) {
  try {
    var date = new Date();
    var column = 3;
    var columnValues = all_core.getRange(2, column, all_core.getLastRow()).getValues();
    var tat;
    var searchString = parameters["surveys"];
    var searchResult = find_index(columnValues, searchString);

    if (searchResult != -1) {
      //   
      if (parameters["decision"] == 'force-stopped') {
        all_core.getRange(searchResult + 2, 2).setValue("force-stopped");
      } else {
        all_core.getRange(searchResult + 2, 2).setValue("reviewed");
      }
      //      var current = date.getFullYear() + '-' + (date.getMonth() + 1) +  "-" + date.getDate();
      //      current = current + " " + date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds();
      //      var snt = dateAdd(new Date(all_core.getRange(searchResult + 2, 5).getValue().split(" P")[0]), 'hour', 15);
      //      var tat = Math.abs(new Date(current) - snt);
      //      var tat_days = Math.floor((tat/1000)/60/60/24);
      //      var tat_hours = Math.floor((tat/1000)/60/60);
      //      var tat_mins = Math.floor((tat/1000)/60);
      //      var tat_secs = Math.floor((tat/1000)%60);
      var tat_final = get_finish_TAT(all_core.getRange(searchResult + 2, 5).getValue() + " PDT -0700");
      all_core.getRange(searchResult + 2, 13).setValue(parameters["screenshot"]);
      all_core.getRange(searchResult + 2, 14).setValue(parameters["decision"]);
      all_core.getRange(searchResult + 2, 17).setValue(parameters["categories"]);
      all_core.getRange(searchResult + 2, 18).setValue(parameters["aht"]);
      all_core.getRange(searchResult + 2, 12).setValue(parameters["surveytype"]);
      all_core.getRange(searchResult + 2, 19).setValue(tat_final);
      all_core.getRange(searchResult + 2, 21).setValue(parameters["comment_just"]);

      if (all_core.getRange(searchResult + 2, 16).getValue() != '') {
        // all.getRange(searchResult + 2, 11).setValue('eyy');
      } else {
        all_core.getRange(searchResult + 2, 16).setValue(date);
      }
      var tracker = SpreadsheetApp.openById("1wg2v8VaL_EBLwnybTrNccnZDQ2bE2Cq_iPXjR2zlWVk");
      var email_sheet = tracker.getSheetByName("needs edit email");
      var columnValues = email_sheet.getRange(2, 1, email_sheet.getLastRow()).getValues();
      var tat;
      var searchString = parameters["surveys"];
      var searchResult = columnValues.findIndex(searchString);

      if (searchResult != -1) {
        var endDate = (date.getMonth() + 1) + '/' + date.getDate() + "/" + date.getFullYear();
        endDate = endDate + " " + date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds();
        email_sheet.getRange(searchResult + 2, 4).setValue(parameters["categories"]);
        email_sheet.getRange(searchResult + 2, 3).setValue(endDate);
      } else {
        if (parameters["decision"] == "noncompliant")
          return (parameters["surveys"] + " not found.");
      }

    }
    else {
      return (parameters["surveys"] + " not found.");
    }
    return ("success");
  } catch (err) {
    return ("error");
  }
}

function close_core(survey) {
  var date = new Date();
  var column = 3;
  var columnValues = all_core.getRange(2, column, all_core.getLastRow()).getValues();
  var tat;
  var searchString = survey;
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    all_core.getRange(searchResult + 2, 2).setValue("duplicate");
    duplicate.appendRow(all_core.getRange(searchResult + 2, 1, 1, 22).getValues()[0]);
    all_core.deleteRow(searchResult + 2);
  }
  else {
    return (survey + " not found.");
  }
  return ("success");

}

function finish_crust(parameters) {
  var date = new Date();
  var column = 3;
  var columnValues = all_crust.getRange(2, column, all_crust.getLastRow()).getValues();
  var tat;
  var searchString = parameters["surveys"];
  var searchResult = find_index(columnValues, searchString);

  if (searchResult != -1) {
    //   
    if (parameters["decision"] == 'force-stopped') {
      all_crust.getRange(searchResult + 2, 2).setValue("force-stopped");
    } else {
      all_crust.getRange(searchResult + 2, 2).setValue("reviewed");
    }
    //      var current = date.getFullYear() + '-' + (date.getMonth() + 1) +  "-" + date.getDate();
    //      current = current + " " + date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds();
    //      var snt = dateAdd(new Date(all_core.getRange(searchResult + 2, 5).getValue().split(" P")[0]), 'hour', 15);
    //      var tat = Math.abs(new Date(current) - snt);
    //      var tat_days = Math.floor((tat/1000)/60/60/24);
    //      var tat_hours = Math.floor((tat/1000)/60/60);
    //      var tat_mins = Math.floor((tat/1000)/60);
    //      var tat_secs = Math.floor((tat/1000)%60);
    var tat_final = get_finish_TAT(all_crust.getRange(searchResult + 2, 5).getValue());
    all_crust.getRange(searchResult + 2, 13).setValue(parameters["screenshot"]);
    all_crust.getRange(searchResult + 2, 14).setValue(parameters["decision"]);
    all_crust.getRange(searchResult + 2, 17).setValue(parameters["categories"]);
    all_crust.getRange(searchResult + 2, 18).setValue(parameters["aht"]);
    all_crust.getRange(searchResult + 2, 12).setValue(parameters["surveytype"]);
    all_crust.getRange(searchResult + 2, 19).setValue(tat_final);
    all_crust.getRange(searchResult + 2, 21).setValue(parameters["comment_just"]);

    var tracker = SpreadsheetApp.openById("1wg2v8VaL_EBLwnybTrNccnZDQ2bE2Cq_iPXjR2zlWVk");
    var email_sheet = tracker.getSheetByName("needs edit email");
    var columnValues = email_sheet.getRange(2, 1, email_sheet.getLastRow()).getValues();
    var tat;
    var searchString = parameters["surveys"];
    var searchResult = columnValues.findIndex(searchString);

    if (searchResult != -1) {
      var endDate = (date.getMonth() + 1) + '/' + date.getDate() + "/" + date.getFullYear();
      endDate = endDate + " " + date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds();
      email_sheet.getRange(searchResult + 2, 4).setValue(parameters["categories"]);
      email_sheet.getRange(searchResult + 2, 3).setValue(endDate);
    } else {
      if (parameters["decision"] == "noncompliant")
        return (parameters["surveys"] + " not found.");
    }

  }
  else {
    return (parameters["surveys"] + " not found.");
  }
  return ("success");
}

function send_email(parameters) {
  MailApp.sendEmail(parameters["recipient"], parameters["subject"], "", { 'htmlBody': parameters["email"], 'noReply': true });
  outbound.appendRow([email.split("@")[0], parameters["surveyid"], parameters["studyid"], parameters["recipient"], parameters["subject"], parameters["email"]]);
  outbound_new.appendRow([email.split("@")[0], parameters["surveyid"], parameters["studyid"], parameters["recipient"], parameters["subject"], parameters["email"]]);
  return "success";
}



function send_email_ne(parameters) {
  var date = new Date();
  //  var aliases = GmailApp.getAliases();
  var tracker = SpreadsheetApp.openById("1wg2v8VaL_EBLwnybTrNccnZDQ2bE2Cq_iPXjR2zlWVk");
  var email_sheet = tracker.getSheetByName("needs edit email");
  var tracker_new = SpreadsheetApp.openById("1WGFeTH-VVuY_lCpI9BoHe9N8mTHYQWFULGnzWwCNKco");
  var email_sheet_new = tracker_new.getSheetByName("needs edit email");
  MailApp.sendEmail(parameters["recipient"], parameters["subject"], "", { 'htmlBody': parameters["email"], 'noReply': true });
  email_sheet.appendRow([parameters["surveyid"], email.split("@")[0], "", "", date, parameters["recipient"], parameters["subject"], parameters["email"]]);
  email_sheet_new.appendRow([parameters["surveyid"], email.split("@")[0], "", "", date, parameters["recipient"], parameters["subject"], parameters["email"]]);
  return "success";
}

function sendFeedback(parameters) {
  var date = new Date();
  var fb = SpreadsheetApp.openById("1mOxzQL6RbkmQT1sa3apVr_LSp_QXObEnvMBBahkqDok");
  var all = fb.getSheetByName("CORE");
  MailApp.sendEmail(parameters["recipient"], parameters["subject"], "", { 'htmlBody': parameters["messageText"], 'noReply': true });
  all.appendRow(["", email.split("@")[0], date, parameters["surveyid"], parameters["message"], parameters["screenshot"], "new"]);
  return "success";
}


function get_finish_TAT(sent) {
  eval(UrlFetchApp.fetch('https://cdn.jsdelivr.net/npm/moment@2.22.2/moment.min.js').getContentText());
  var sDateTime = new Date(sent);
  var localConversion = sDateTime.toLocaleString().replace(",", "");
  var startTime = [moment(localConversion).format("HH:mm:ss")];

  var refstartTime = startTime.slice(1).reduce((prev, cur) => moment.duration(cur).add(prev), moment.duration(startTime[0]));

  var eTime = new Date();
  var endTime = [moment(eTime).format("HH:mm:ss")];

  var refendTime = endTime.slice(1).reduce((prev, cur) => moment.duration(cur).add(prev), moment.duration(endTime[0]));

  var tatCount = refendTime.asMilliseconds() - refstartTime.asMilliseconds();
  var tatCountFinal = moment.utc(tatCount).format("HH:mm:ss");

  return (tatCountFinal);  //set TAT
}

function create_announcement(parameters) {
  const notif_tracker = SpreadsheetApp.openById("1mMMC-oOALm5shbdljtEdOb9XDciw72qu2z-HxZiH-iM"),
    notif = notif_tracker.getSheetByName("notification");

  notif.appendRow([parameters["notif_id"], parameters["subject"], parameters["about"], parameters["screenshot"], parameters["ldap"], parameters["date"]]);
  return "success"
}

function add_log(parameters) {
  // var date = new Date();
  // var columnValues = all_core.getRange(2, 3, all_core.getLastRow()).getValues(); 
  //   var searchString = parameters["survey"];
  //   var searchResult = find_index(columnValues, searchString);

  //   if(searchResult != -1)
  //   { 
  //     if(all_core.getRange(searchResult + 2, 22).getValue() != "")
  //       all_core.getRange(searchResult + 2, 22).setValue(all_core.getRange(searchResult + 2, 22).getValue() + " | " + email.split("@")[0] + " " + parameters["log"] + "(" + date + ")");
  //     else
  //       all_core.getRange(searchResult + 2, 22).setValue(email.split("@")[0] + " " + parameters["log"] + "(" + date + ")");
  //   }
  Logger.log(parameters);

}

function get_dump() {
  var final = [];
  var list = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getDataRange().getValues();
  for (var i = 2; i < list.length; i++) {
    final.push([
      list[i][0]
      , list[i][2]
      , list[i][5]
      , list[i][6]
      , list[i][7]
      , list[i][8]
      , list[i][11]
      , new Date(list[i][15]).toLocaleString().split(",")[0]
      , list[i][13]
      , new Date(list[i][15]).toLocaleString()
    ]);
  }
  var list = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getDataRange().getValues();
  for (var i = 2; i < list.length; i++) {
    final.push([
      list[i][0]
      , list[i][2]
      , list[i][5]
      , list[i][6]
      , list[i][7]
      , list[i][8]
      , list[i][11]
      , new Date(list[i][15]).toLocaleString().split(",")[0]
      , list[i][13]
      , new Date(list[i][15]).toLocaleString()
    ]);
  }
  var list2 = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getDataRange().getValues();
  for (var i = 2; i < list2.length; i++) {
    final.push([
      list2[i][0]
      , list2[i][2]
      , list2[i][5]
      , list2[i][6]
      , list2[i][7]
      , list2[i][8]
      , list2[i][11]
      , new Date(list2[i][15]).toLocaleString().split(",")[0]
      , list2[i][13]
      , new Date(list2[i][15]).toLocaleString()
    ]);
  }
  return (JSON.stringify(final));
}

function get_overview() {
  var list = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - July 2019").getDataRange().getValues(), overall = 0, compliant = 0, noncompliant = 0, reviewed_more = 0;
  for (var i = 2; i < list.length; i++) {
    if (new Date(list[i][15]).toLocaleString().split(",")[0] == new Date().toLocaleString().split(",")[0]) {
      overall++;
      if (list[i][13] == "compliant") compliant++;
      else if (list[i][13] == "noncompliant") {
        noncompliant++;
        if (list[i][9] > 1)
          reviewed_more++;
      }

    }
  }
  var list2 = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - September 2019").getDataRange().getValues();
  for (var i = 2; i < list2.length; i++) {
    overall++;
    if (list2[i][13] == "compliant") compliant++;
    else if (list2[i][13] == "noncompliant") {
      noncompliant++;
      if (list2[i][9] > 1)
        reviewed_more++;
    }
  }
  var list2 = SpreadsheetApp.openById("1QQZZIyNLxQdKTkeAepVR-elO8D8PyGycNyQwGUyTD4Y").getSheetByName("Archived CORE - August 2019").getDataRange().getValues();
  for (var i = 2; i < list2.length; i++) {
    overall++;
    if (list2[i][13] == "compliant") compliant++;
    else if (list2[i][13] == "noncompliant") {
      noncompliant++;
      if (list2[i][9] > 1)
        reviewed_more++;
    }
  }
  overall -= 1;
  compliant -= 1;
  return ({
    overall: overall,
    compliant: compliant,
    noncompliant: noncompliant,
    reviewed_more: reviewed_more
  });
}


function finish_ewoq(parameters) {
  eval(UrlFetchApp.fetch('https://cdn.jsdelivr.net/npm/moment@2.22.2/moment.min.js').getContentText());

  const tracker = SpreadsheetApp.openById("1WGFeTH-VVuY_lCpI9BoHe9N8mTHYQWFULGnzWwCNKco"),
    arch = tracker.getSheetByName("All Cases");
  const tracker_far = SpreadsheetApp.openById("1YuOdhrsoJzjMm5b8fqm5X74_cAQHn88hGP9r0uoyE1Q"),
    arch_far = tracker_far.getSheetByName("Tracker - QA");
  let date = new Date();

  let tat,
    startTime,
    endTime,
    aht;

  var searchString = parameters["surveys"];
  startTime = parameters["start_time"];
  if (parameters["end_time"] != "") endTime = parameters["end_time"];
  else
    endTime = date;
  aht = moment.utc(moment(new Date(endTime), "DD/MM/YYYY HH:mm:ss").diff(moment(new Date(startTime), "DD/MM/YYYY HH:mm:ss"))).format("HH:mm:ss");
  let tool = "Cases 2.0";
  if (parameters["surveys"].indexOf("-") != -1) tool = "Cases 2.0";
  else tool = "EWOQ";

  let num_int = 0;
  if (String(parameters["surveys"]).indexOf("/") == -1 && String(parameters["surveys"]).indexOf("&") == -1 && String(parameters["surveys"]).indexOf(",") == -1 && String(parameters["surveys"]).indexOf("/") == -1 && String(parameters["surveys"]).indexOf("-") != -1)
    num_int = String(parameters["surveys"]).trim().match(/.{1,15}/g).length;
  else
    num_int = String(parameters["surveys"]).trim().split(/,|\/|&| /).filter(String).length;
  arch.appendRow([email.split("@")[0], "reviewed", guid(), "'" + parameters["surveys"], parameters["queuetype"], parameters["customertype"], tool, parameters["language"], parameters["country"], "", parameters["rmto"], parameters["surveytype"], parameters["screenshot"], parameters["decision"], new Date(startTime), dateAdd(new Date(startTime), "hour", 15), new Date(endTime), dateAdd(new Date(endTime), "hour", 15), parameters["categories"], "'" + aht, num_int])

  if (parameters["categories"].indexOf("Forum Access Request") != -1) {
    arch_far.appendRow(["", new Date().toLocaleString().split(",")[0], new Date(endTime), parameters["surveys"], email.split("@")[0], parameters["categories"]]);
  }

  return ("success");
}


function dateAdd(date, interval, units) {
  var ret = new Date(date); //don't change original date
  var checkRollover = function () { if (ret.getDate() != date.getDate()) ret.setDate(0); };
  switch (interval.toLowerCase()) {
    case 'year': ret.setFullYear(ret.getFullYear() + units); checkRollover(); break;
    case 'quarter': ret.setMonth(ret.getMonth() + 3 * units); checkRollover(); break;
    case 'month': ret.setMonth(ret.getMonth() + units); checkRollover(); break;
    case 'week': ret.setDate(ret.getDate() + 7 * units); break;
    case 'day': ret.setDate(ret.getDate() + units); break;
    case 'hour': ret.setTime(ret.getTime() - units * 3600000); break;
    case 'minute': ret.setTime(ret.getTime() + units * 60000); break;
    case 'second': ret.setTime(ret.getTime() + units * 1000); break;
    default: ret = undefined; break;
  }
  return ret;
}

function guid() {
  function s4() {
    return Math.floor((1 + Math.random()) * 0x10000)
      .toString(16)
      .substring(1);
  }
  return 'case' + '-' + s4() + '-' + s4();
}

function get_feedback() {
  var fb = SpreadsheetApp.openById("1mOxzQL6RbkmQT1sa3apVr_LSp_QXObEnvMBBahkqDok");
  var all = fb.getSheetByName("CORE").getDataRange().getValues();
  return (JSON.stringify(all));
}

function find_index(sheet, value) {
  return ArrayLib.indexOf(sheet, 0, value)
}

Array.prototype.findIndex = function (search) {
  if (search == "") return false;
  for (var i = 0; i < this.length; i++)
    if (this[i].indexOf(search) != -1) return i;

  return -1;
} 