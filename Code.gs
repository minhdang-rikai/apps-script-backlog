var baseApiUrl = "https://rikai.backlog.com/api/v2"
var apiKey = "YFSVss6ObEPNaudqRQr6BL0RxjEdeX4GoPhjtikQPFZvYfhKxq4bhDu1HS9yaEPo";
var projectId = 129818;
var priorityId = 2; // High
var sheetId = '1RO2fbyhqMYw1FoIB-D1fC3DbJKxZ0H3puTkLBoWTU10'
var sheetName = 'WBS(Tracking VN)'

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Backlog')
    .addItem('Add task', 'addTask')
    .addItem('Sync user', 'syncUsers')
    .addToUi();
}

function addTask() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createTemplateFromFile("index");
  html.mode = "init";

  ui.showModalDialog(html.evaluate(), 'Add task for Backlog');
}

function callBacklogAPI(summary, description, issuesType, assignee, startDate, dueDate, estimatedHours, row) {
  var apiUrl = baseApiUrl + "/issues?" +
    "apiKey=" + apiKey +
    "&projectId=" + projectId +
    "&summary=" + summary +
    "&issueTypeId=" + issuesType +
    "&assigneeId=" + assignee +
    "&description=" + description +
    "&startDate=" + startDate +
    "&dueDate=" + dueDate +
    "&estimatedHours=" + estimatedHours +
    "&priorityId=" + priorityId;

  var options = {
    "method": "POST",
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(apiUrl, options);

  Logger.log(response.getContentText());

  const jsonResponse = JSON.parse(response.getContentText());

  if (jsonResponse.errors) {
    const errorMessages = jsonResponse.errors.map(error => error.message);
    const errorMessageString = errorMessages.join('\n');

    SpreadsheetApp.getUi().alert(`${row}: ${errorMessageString}`)
  } else {
    SpreadsheetApp.getUi().alert(`${row}: Successfully!`)
  }
}

async function getUsers() {
  var apiUrl = baseApiUrl + `/projects/${projectId}/users?apiKey=${apiKey}`;

  var options = {
    "method": "GET",
    "muteHttpExceptions": true
  };

  var response = await UrlFetchApp.fetch(apiUrl, options);

  const jsonResponse = JSON.parse(response.getContentText());

  if (jsonResponse.errors) {
    return []
  }

  return jsonResponse.map(i => {
    return {
      id: i.id,
      name: i.name,
      mailAddress: i.mailAddress
    }
  })
}

async function syncUsers() {
  const users = await getUsers();
  formatColumnAssign(users);
}

function formatColumnAssign(users) {
  var spreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (users && users.length > 0) {
    var validationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(users.map(function (item) { return item.name; }), true)
      .build();

    var range = sheet.getRange('J5:J');
    range.setDataValidation(validationRule);
  }
}

async function getRow(rowsInput) {
  var rows = parseRowsInput(rowsInput);
  var users = await getUsers() 

  rows.forEach(function (row) {
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var sheet = spreadsheet.getSheetByName(sheetName);

    var rowData = sheet.getRange('A' + row + ':Z' + row).getValues()[0];

    var summary = rowData[4];
    var description = rowData[5];
    var issuesType = rowData[6];
    var assignee = rowData[9];
    var startDate = rowData[14];
    var dueDate = rowData[15];
    var estimatedHours = rowData[10];

    if (startDate) startDate = parseDate(startDate);
    if (dueDate) dueDate = parseDate(dueDate);
    if (issuesType) {
      const findIssueType = issuesTypes.find(x => x.name == issuesType)
      issuesType = findIssueType.id

      if (findIssueType.templateSummary) {
        summary = findIssueType.templateSummary + summary
      }
    }
    if (assignee) {
      const findAssignee = users.find(x => x.name == assignee)
      assignee = findAssignee.id
    }

    callBacklogAPI(summary, description, issuesType, assignee, startDate, dueDate, estimatedHours, row);
  });
}


function parseDate(dateString) {
  const date = new Date(dateString);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}

function parseRowsInput(rowsInput) {
  var rows = [];
  var parts = rowsInput.split(',');

  parts.forEach(function (part) {
    if (part.includes('-')) {
      var range = part.split('-');
      var start = parseInt(range[0], 10);
      var end = parseInt(range[1], 10);
      for (var i = start; i <= end; i++) {
        rows.push(i);
      }
    } else {
      rows.push(parseInt(part.trim(), 10));
    }
  });

  return rows;
}
