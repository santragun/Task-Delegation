function CollectFormValues() {
  SetUpTriggers();
  try {
    let EMPLOYEE_DB = SpreadsheetApp.openById(SpreadsheetId);

    let departments = EMPLOYEE_DB.getSheetByName('departments').getDataRange().getValues();
    let employees = EMPLOYEE_DB.getSheetByName('employees').getDataRange().getValues();

    var formQuestions = FormApp.getActiveForm().getItems();

    for (var i = 0; i < formQuestions.length; i++) {
      if (formQuestions[i].getTitle() == 'Task Type') {
        try {
          let list = formQuestions[i].asListItem();
          let choices = [];
          for (var j = 1; j < departments.length; j++) {
            choices.push(list.createChoice(departments[j][0]));
          }
          list.setChoices(choices);
        }
        catch ({ message }) {
          Logger.log('Departments list Exception:  %s', message);
        }
      }
      try {
        if (formQuestions[i].getTitle() == 'Assigned To') {
          let list = formQuestions[i].asListItem();
          let choices = [];
          for (var j = 1; j < employees.length; j++) {
            choices.push(list.createChoice(employees[j][0]));
          }
          list.setChoices(choices);
        }

        if (formQuestions[i].getTitle() == 'Assigned By') {
          let list = formQuestions[i].asListItem();
          let choices = [];
          for (var j = 1; j < employees.length; j++) {
            choices.push(list.createChoice(employees[j][0]));
          }
          list.setChoices(choices);
        }
      }
      catch ({ message }) {
        Logger.log('Employees list Exception:  %s', message);
      }
    }

    Logger.log('Completed');
    return 'done';
  }
  catch ({ message }) {
    Logger.log('CollectFormValues Exception:  %s', message);
    if (message.includes('You do not have permission to call')) {
      return 'authorize';
    }
  }
}

function SetUpTriggers() {
  if (!CheckTriggers(ScriptApp.EventType.ON_OPEN, 'CollectFormValues')) {
    let onOpenTrigger = documentProperties.getProperty('onOpen');
    if (onOpenTrigger) {
      documentProperties.deleteProperty('onOpen');
    }
    try {
      ScriptApp.newTrigger('CollectFormValues')
        .forForm(FormApp.getActiveForm())
        .onOpen()
        .create();
      documentProperties.setProperty('onOpen', 'enabled');
    }
    catch ({ message }) {
      Logger.log('onOpen trigger Exception:  %s', message);
    }
  }

  if (!CheckTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, 'onFormSubmit')) {
    try {
      ScriptApp.newTrigger('onFormSubmit')
        .forForm(FormApp.getActiveForm())
        .onFormSubmit()
        .create();
    }
    catch ({ message }) {
      Logger.log('onSubmit trigger Exception:  %s', message);
    }
  }
}


function onFormSubmit(e) {
  try {
    let formResponse = e.response;
    let id = formResponse.getId();
    Logger.log(id + 'response id');
    let itemResponses = formResponse.getItemResponses();

    let assignedTo = itemResponses[2].getResponse();
    let assignedToEmail = getEmployeeEmail(assignedTo);

    let form = FormApp.openById(responseFormID);
    let item = form.getItems()[1].asTextItem();

    let response = form.createResponse();
    let itemResponse = item.createResponse(id);
    let prefilledUrl = response.withItemResponse(itemResponse).toPrefilledUrl();

    if (assignedToEmail) {

      let template = HtmlService.createTemplateFromFile('assignedNotification');
      template.responses = itemResponses;
      // template.form = CreateReplyForm();
      template.form = prefilledUrl;

      let message = template.evaluate();

      MailApp.sendEmail(assignedToEmail,
        'Task has been assigned to you',
        '',
        { htmlBody: message.getContent() });
    }
  }
  catch ({ message }) {
    Logger.log('Notify Employee Exception: %s', message);
  }
}



function getEmployeeEmail(name) {
  let EMPLOYEE_DB = SpreadsheetApp.openById(SpreadsheetId);
  let employees = EMPLOYEE_DB.getSheetByName('employees').getDataRange().getValues();
  for (let i = 1; i < employees.length; i++) {
    if (employees[i][0] == name) {
      return employees[i][1];
    }
  }
}

function CheckTriggers(eventType, handlerFunction) {
  let triggers = ScriptApp.getUserTriggers(FormApp.getActiveForm());
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() == eventType && triggers[i].getHandlerFunction() == handlerFunction) {
      return true;
    }
  }
  return false;
}


function getAuthUrl() {
  var authInfo, message;

  authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  Logger.log(authInfo.getAuthorizationUrl());
  message = 'Task Delegation needs permission to use Apps Script Services.  Click ' +
    'this url to authorize: <br><br>' +
    '<a href="' + authInfo.getAuthorizationUrl() +
    '">Link to Authorization Dialog</a>' +
    '<br><br> Task Delegation needs to either ' +
    'be authorized or re-authorized.';

  return message;
}

