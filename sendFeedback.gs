function pageMeister_sendFeedbackUi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var sendNotifications = ScriptProperties.getProperty('sendNotifications');
  var emailCol = "email";
  var app = UiApp.createApplication().setHeight(440).setWidth(400).setTitle('Send personalized emails to students');
  var panel = app.createVerticalPanel().setTitle('Set up file handling');
  var panel = app.createVerticalPanel().setId("settingsPanel").setWidth("390px"); 
  var waitingPanel = app.createVerticalPanel().setWidth("390px").setHeight("430px").setVisible(false);
  var waitingImage = app.createImage(this.IMAGEPATH).setHeight("200px").setWidth("200px");
  var waitingNote = app.createLabel('I\m busy sending emails...');
  waitingPanel.add(waitingImage);
  waitingPanel.add(waitingNote);
 
  var folderPanel = app.createVerticalPanel();
  var mainGrid = app.createGrid(10, 1).setId("mainGrid");
  //Help text below dynamically loads all field names from the sheet using normalized (camelCase) sheet headers
  var sheetFieldLabel = app.createLabel('Use these variables substitute spreadsheet values into any of the fields below.')
  var sheetFieldNames = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var normalizedSheetFieldNames = normalizeHeaders(sheetFieldNames);
  var fieldHelpTable = app.createFlexTable();
  for (var i = 0; i<normalizedSheetFieldNames.length; i++) {
    var variable = app.createLabel("$"+normalizedSheetFieldNames[i]).setStyleAttribute('color', 'blue');
    fieldHelpTable.setWidget(i, 0, variable)
  }
  var fieldHelpScrollPanel = app.createScrollPanel(fieldHelpTable).setHeight("100px");
  mainGrid.setWidget(0, 0, sheetFieldLabel);
  mainGrid.setWidget(1, 0, fieldHelpScrollPanel);
  var emailLabel = app.createLabel('Use variables to email custom notifications, grades and feedback').setStyleAttribute("backgroundColor", "grey").setStyleAttribute("color", "white").setWidth("100%");
  var emailFieldLabel = app.createLabel('Recipient email address(es).'); 
  var email = app.createTextBox().setWidth("100%").setName("feedbackEmailString").setValue("$"+emailCol);
  var feedbackEmailString = ScriptProperties.getProperty('feedbackEmailString');
  if (feedbackEmailString) {
    email.setValue(feedbackEmailString);
  } else {
    email.setValue("$" + emailCol);
  }
  var subjectLabel = app.createLabel('Email subject');
  var subjectBox = app.createTextBox().setWidth("100%").setName("feedbackSubjectString");
  var feedbackSubjectString = ScriptProperties.getProperty('feedbackSubjectString');
  if (feedbackSubjectString) {
    subjectBox.setValue(feedbackSubjectString);
  } else {
    subjectBox.setValue('Feedback on $pageTitle');
  }
  var emailBodyLabel = app.createLabel('Email body. HTML friendly. (Note: a link to the Site Page will be automatically included.)');
  var emailBodyArea = app.createTextArea().setWidth("100%").setHeight("90px").setName("feedbackBodyString");
  var feedbackBodyString = ScriptProperties.getProperty('feedbackBodyString');
  if ((feedbackBodyString)&&(feedbackBodyString!='')){
    emailBodyArea.setValue(feedbackBodyString);
  } else {
    emailBodyArea.setValue('$feedback');
  }
  mainGrid.setWidget(2, 0, emailLabel);
  mainGrid.setWidget(3, 0, emailFieldLabel);
  mainGrid.setWidget(4, 0, email);
  mainGrid.setWidget(5, 0, subjectLabel);
  mainGrid.setWidget(6, 0, subjectBox);
  mainGrid.setWidget(7, 0, emailBodyLabel);
  mainGrid.setWidget(8, 0, emailBodyArea);
  var buttonHandler = app.createServerHandler('sendFeedback').addCallbackElement(panel);
  var waitingHandler = app.createClientHandler().forTargets(panel).setVisible(false).forTargets(waitingPanel).setVisible(true);
  var button = app.createButton('Save and send emails').addClickHandler(buttonHandler).addClickHandler(waitingHandler);
  mainGrid.setWidget(9, 0, button);  
  panel.add(mainGrid);
  app.add(panel);
  app.add(waitingPanel);
  ss.show(app);
  return app;
}


function sendFeedback(e) {
  var app = UiApp.getActiveApplication();
  var sheetName = ScriptProperties.getProperty('sheetName');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var data = getRowsData(sheet);
  var feedbackEmailString = e.parameter.feedbackEmailString;
  var emailCol = 'email';  
  if (!feedbackEmailString) {
    feedbackEmailString = "$" + emailCol;
  }
  var feedbackSubjectString = e.parameter.feedbackSubjectString;
  var feedbackBodyString = e.parameter.feedbackBodyString;
  ScriptProperties.setProperty('feedbackEmailString', feedbackEmailString);
  ScriptProperties.setProperty('feedbackSubjectString', feedbackSubjectString);
  ScriptProperties.setProperty('feedbackBodyString', feedbackBodyString);
  var badEmails = new Array();
  var emailsSent = 0;
  for (var i=0; i<data.length; i++) {
    var url = data[i].linkToPage;
    var pageTitle = SitesApp.getPageByUrl(url).getTitle();
    var htmlLink = '<br/><br/>Link to your page: <a href = "'+url+'">'+pageTitle+'</a>';
    var recipient =  replaceFeedbackStringFields(feedbackEmailString, data[i],sheet);
    var subject = replaceFeedbackStringFields(feedbackSubjectString, data[i],sheet);
    var body = replaceFeedbackStringFields(feedbackBodyString, data[i],sheet);
    body+=htmlLink;
    try {
      MailApp.sendEmail(recipient, subject, '', {htmlBody: body});
      pageMeister_logFeedbackEmail();
      emailsSent++;
    } catch(err) {
      Logger.log(err);
      badEmails.push(recipient);
    }
  }
  var errMsg = '';
  if (badEmails.length>0) {
    errMsg += "There were problems with the following email addresses: " + badEmails.join(", ");
  }
  Browser.msgBox(emailsSent + " emails successfully sent." + errMsg);
  app.close();
  return app;
}


// This function subs in row values for $variables
function replaceFeedbackStringFields(string, rowData, sheet) {
  var newString = string;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var normalizedHeaders = normalizeHeaders(headers);
  var mergeTags = "$"+normalizedHeaders.join(",$");
  mergeTags = mergeTags.split(",");
  for (var i=0; i< mergeTags.length; i++) {
    var key = normalizedHeaders[i];
    var replacementValue = rowData[key];
    var replaceTag = mergeTags[i];
    replaceTag = replaceTag.replace("$","\\$") + "\\b";
    var find = new RegExp(replaceTag, "g");
    newString = newString.replace(find, replacementValue);
    newString = newString.replace(/(\r\n|\n|\r)/gm,"<br>");
  }
  return newString;
}
