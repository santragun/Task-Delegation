function onOpen(e) {
  try {
    let menu = FormApp.getUi().createAddonMenu();

    if (!documentProperties.getProperty('onOpen')) {
      menu.addItem('Start Task Delegation Add-on', 'AddonStatus');
    }
    else {
      menu.addItem('Check Task Delegation Add-on', 'AddonStatus');
    }
    menu.addToUi();
  }
  catch (message) {
    Logger.log('onOpen:  %s', message);
  }
}

function onInstall(e) {
  onOpen(e);
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}


function AddonStatus() {
  // form = FormApp.getActiveForm();
  // MailApp.sendEmail()
  // UrlFetchApp.fetch('');

  let configure = CollectFormValues();
  if (configure == 'authorize') {
    var authorizationHtml = HtmlService.createTemplateFromFile('authorization').evaluate();
    return FormApp.getUi().showModalDialog(authorizationHtml, 'Authorization');
  }
  else {
    return FormApp.getUi().showModalDialog(HtmlService.createHtmlOutput('Add-on is running successfully'), 'Task Delegation Add-on Status');
  }
}

