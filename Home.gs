function CollectFormValues() {
  try {
    let departments = EMPLOYEE_DB.getSheetByName('departments').getDataRange().getValues();
    let employees = EMPLOYEE_DB.getSheetByName('employees').getDataRange().getValues();

    let formQuestions = FORM.getItems();

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
          Logger.log('Departments list Exception: ' + message);
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
        Logger.log('Employees list Exception: ' + message);
      }
    }

  }
  catch ({ message }) {
    Logger.log('CollectFormValues Exception: ' + message);
  }


}