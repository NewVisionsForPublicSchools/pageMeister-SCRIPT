var scriptTitle = "pageMeister Script V1.0.2 (9/11/13)";
// Written by Andrew Stillman for New Visions for Public Schools
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
// Support and contact at http://www.youpd.org/pagemeister  (screencast doesn't exist yet)

//Want to run autoCrat on a time-based trigger?  
//Set time triggers on the autoCrat_onFormSubmit function.
var scriptName = "pageMeiser"
var scriptTrackingId = "UA-40505612-1"

var pathToDrive = "https://googledrive.com/host/"  + '0B2vrNcqyzernM01qazZIQm1yYkE';
var IMAGEPATH = pathToDrive + "/pageMeister.gif";


function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuItems = [];
  menuItems[0] = {name: "What is pageMeister?", functionName: "pageMeister_whatIs"};
  menuItems[1] = {name: "Set up and run pageMeister", functionName: "pageMeister"};
  menuItems[2] = null;
  menuItems[3] = {name: "Get last page update time", functionName: "pageMeister_getLastEdit"};
  menuItems[4] = {name: "Send feedback emails", functionName: "pageMeister_sendFeedbackUi"};
  ss.addMenu('pageMeister', menuItems);
}

function pageMeister() {
  setpageMeisterSid();
  var app = UiApp.createApplication().setTitle("Page Creation Settings").setHeight(400).setWidth(480);
  var topLabel = app.createLabel("This script will only create new pages for rows with blank \"Status\" values.")
  app.add(topLabel);
  var waitingIcon = app.createImage(IMAGEPATH).setWidth('200px').setHeight('200px').setStyleAttribute('position', 'absolute').setStyleAttribute('left', '100px').setStyleAttribute('top', '100px').setId('waitingIcon').setVisible(false);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var test = sheet.getLastRow();
  if (sheet.getLastRow()<1) {
    sheet.getRange(1, 1, 1, 8).setValues([["First Name","Last Name","Email","Page Title","Status","Link to Page","Last Updated","Feedback"]]);
    sheet.getRange(1, 3).setNote("Do not change this column heading");
    sheet.getRange(1, 4).setNote("Do not change this column heading");
    sheet.getRange(1, 5).setNote("Do not change this column heading");
    sheet.getRange(1, 6).setNote("Do not change this column heading");
    sheet.getRange(1, 7).setNote("Do not change this column heading");
    sheet.getRange(2, 4).setFormula('=CONCATENATE(A2," ",B2)');
    sheet.getRange(1, 5, 1, 3).setBackground('black').setFontColor('white')
    sheet.setFrozenRows(1);
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var panel = app.createVerticalPanel().setSpacing(4).setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('margin', '5px');
  var siteURLLabel = app.createLabel("Full URL of Google Site under which you want student pages");
  var URLHandler = app.createServerHandler('parentPageRefresh').addCallbackElement(panel);
  var siteURLBox = app.createTextBox().setWidth("100%").setName('siteURL').addKeyUpHandler(URLHandler);
  var siteURL = ScriptProperties.getProperty('siteURL');
  var parentPageLabel = app.createLabel("PAGE under which you want student pages created");
  var parentPageList = app.createListBox().setName('parentPageName').setId('parentPageList').setWidth("200px");
  var templateLabel = app.createLabel("Page template to use");
  var templateList = app.createListBox().setName('template').setId('templateList').setWidth("200px");
  var accessPreferenceLabel = app.createLabel('What access level do you want students to recieve on the whole site?');
  var accessChangeHandler = app.createServerHandler('refreshAccessLabel').addCallbackElement(panel);
  var accessPreferenceBox = app.createListBox().setName('access').addChangeHandler(accessChangeHandler);
  accessPreferenceBox.addItem('Can view').addItem('Can edit');
  var access = ScriptProperties.getProperty('access');
  if (access=="Can view") {
    accessPreferenceBox.setSelectedIndex(0);
  }
  if (access=="Can edit") {
    accessPreferenceBox.setSelectedIndex(1);
  }
  var accessPanel = app.createHorizontalPanel();
  var accessLabel = app.createLabel('').setId('accessLabel');
  accessPanel.add(accessPreferenceBox).add(accessLabel);
  var notificationLabel = app.createLabel('Send notification email...');
  var notificationListBox = app.createListBox().setName('notification');
  var notification = ScriptProperties.getProperty('notification');
  var notificationTypes = ['new','all','none'];
  notificationListBox.addItem('for newly-created pages only','new').addItem('for all rows in sheet','all').addItem('not now','none');
  if (notification) {
    var notifyIndex = notificationTypes.indexOf(notification);
    notificationListBox.setSelectedIndex(notifyIndex);
  }
  var runHandler = app.createServerHandler('runPageCreator').addCallbackElement(panel);
  var waitingHandler = app.createClientHandler().forTargets(waitingIcon).setVisible(true).forTargets(panel).setVisible(false);
  var button = app.createButton("Create pages", runHandler).setId('button').setEnabled(false);
  button.addClickHandler(waitingHandler);
  panel.add(siteURLLabel).add(siteURLBox).add(parentPageLabel).add(parentPageList).add(templateLabel).add(templateList);
  panel.add(accessPreferenceLabel).add(accessPanel);
  panel.add(notificationLabel).add(notificationListBox).add(button);
  refreshAccessLabel();
  app.add(panel);
  if (siteURL) {
    siteURLBox.setValue(siteURL);
    parentPageRefresh();
    button.setEnabled(true);
  }
  app.add(waitingIcon);
  ss.show(app);
  return app;
}

function refreshAccessLabel(e) {
  var app = UiApp.getActiveApplication();
  var accessLabel = app.getElementById('accessLabel').setVisible(true);
  if (e) {
    var access = e.parameter.access;
  } else {
    var access = ScriptProperties.getProperty('access');
  }
  if (access=="Can edit") {
    accessLabel.setText("Students added as editors to whole site. Limitations of the Apps Script API prevent automatically adding students as page-level editors only.  Only use this setting if you don't mind students being able to edit each others' pages.").setStyleAttribute('backgroundColor', 'pink').setStyleAttribute('padding', '4px');
  } else {
    accessLabel.setText("Students added as viewers to whole site. Using \"page level permissions,\" once the script has run you can manually add each student as editor of their respective page.").setStyleAttribute('backgroundColor', 'yellow').setStyleAttribute('padding', '4px');
  }
  return app;
}



function parentPageRefresh(e) {
  var app = UiApp.getActiveApplication();
  var parentPageList = app.getElementById('parentPageList');
  var templateList = app.getElementById('templateList');
  templateList.clear();
  parentPageList.clear();
  if (e) {
    var siteURL = e.parameter.siteURL;
  } else {
    var siteURL = ScriptProperties.getProperty('siteURL');
    if ((siteURL=="")||(siteURL=="undefined")) {
      return app;
    }
  } 
  try {
    var site = SitesApp.getSiteByUrl(siteURL);
    var pages = site.getAllDescendants();
    parentPageList.addItem("Home", "home");
    var nameList = [];
    for (var i=0; i<pages.length; i++) {
      parentPageList.addItem(pages[i].getTitle(), pages[i].getUrl());
      nameList.push(pages[i].getUrl());
    }
    var parentPageName = ScriptProperties.getProperty('parentPageName');
    if (parentPageName) {  
      var index = nameList.indexOf(parentPageName);
      if (index!=-1) {
        parentPageList.setSelectedIndex(index+1);
      }
    }
    try {
    var templates = site.getTemplates();
    } catch(err) {
      templates = [];
    }
    var templateNames = [];
    var stockTemplates = ['Web page','Announcements','File cabinet','List'];
    for (var i=0; i<stockTemplates.length; i++) {
      templateList.addItem(stockTemplates[i],stockTemplates[i]);
      templateNames.push(stockTemplates[i]);
    }
    for (var i=0; i<templates.length; i++) {
      if (templates[i].isTemplate()) {
        templateList.addItem(templates[i].getTitle(),templates[i].getName());
        templateNames.push(templates[i].getName());
      }
    }
    var template = ScriptProperties.getProperty('template');
    if (template) {  
      var index = templateNames.indexOf(template);
      if (index!=-1) {
        templateList.setSelectedIndex(index);
      }
    }
    var button = app.getElementById('button').setEnabled(true);
    return app;
  } catch(err) {
    Logger.log(err.message);
    return app;
  }
}

function runPageCreator(e) {
  var app = UiApp.getActiveApplication();
  if(e) {
    var siteURL = e.parameter.siteURL
    var parentPageName = e.parameter.parentPageName;
    var template = e.parameter.template;
    var access = e.parameter.access;
    var notification = e.parameter.notification;
  }
  if (!siteURL) {
    var siteURL = ScriptProperties.getProperty('siteURL');
  } else {
    ScriptProperties.setProperty('siteURL', siteURL);
  }
  if (!parentPageName) {
    var parentPageName = ScriptProperties.getProperty('parentPageName');
  } else {
    ScriptProperties.setProperty('parentPageName', parentPageName);
  }
  if (!template) {
    var template = ScriptProperties.getProperty('template');
  } else {
    ScriptProperties.setProperty('template', template);
  }
  if (!access) {
    var access = ScriptProperties.getProperty('access');
  } else {
    ScriptProperties.setProperty('access', access);
  }
  if (!notification) {
    var notification = ScriptProperties.getProperty('notification');
  } else {
    ScriptProperties.setProperty('notification', notification);
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusCol = headers.indexOf("Status")+1;
  var linkCol = headers.indexOf("Link to Page")+1;
  if (sheet.getLastRow()>1) {
    var range = sheet.getRange(2, sheet.getLastColumn(), sheet.getLastRow()-1);
  } else {
    app.close();
    return app;
  }
  var data = getRowsData(sheet);
  var site = SitesApp.getSiteByUrl(siteURL);
  if (parentPageName=="home") {
    var parentPage = site;
  } else {
    var parentPage = SitesApp.getPageByUrl(parentPageName);
  }
  for (var i=0; i<data.length; i++) {
    if ((!data[i].status)&&(data[i].pageTitle!=' ')&&(data[i].pageTitle!='')) {
      var email = data[i].email;
      var pageTitle = data[i].pageTitle;
      var statusObj = createPage(email,pageTitle,template,parentPage,site,access,notification);
      sheet.getRange(i+2, statusCol).setValue(statusObj.message);
      sheet.getRange(i+2, linkCol).setValue(statusObj.url);
      SpreadsheetApp.flush();
    } else if ((data[i].linkToPage) && (notification=="all")) {
      var email = data[i].email;
      var pageTitle = data[i].pageTitle;
      MailApp.sendEmail(data[i].email, '', "Link to your site page", "Your page: " + pageTitle + " can be accessed at " + data[i].linkToPage);
    }
  }
  pageMeister_getLastEdit();
  app.close();
  return app;
}


function createPage(email,pageTitle,template,parentPage,site,access,notification) {
  var pageName = verifyUnique(site, normalizeHeader(pageTitle));
  var returnObject = new Object();
  switch(template)
  {
    case 'Web page':
      try {
        var page = parentPage.createWebPage(pageTitle, pageName, '');
        returnObject.message = "Created Web Page titled " + pageTitle;
        returnObject.url = page.getUrl();
        pageMeister_logPageCreation();
      } catch(err) {
        returnObject.message = err.message;
        returnObject.url = '';
      }
      break;
    case 'Announcements':
      try {
        var page = parentPage.createAnnouncementsPage(pageTitle, pageName,'');
        returnObject.message = "Created Announcements Page titled " + pageTitle;
        returnObject.url = page.getUrl();
        pageMeister_logPageCreation();
      } catch(err) {
        returnObject.message = err.message;
        returnObject.url = '';
      }
      break;
    case 'File cabinet':
      try {
        var page = parentPage.createFileCabinetPage(pageTitle, pageName,'');
        returnObject.message = "Created File cabinet page titled " + pageTitle;
        returnObject.url = page.getUrl();
      } catch(err) {
        returnObject.message = err.message;
        returnObject.url = '';
      }
      break;
    case 'List':
      try {
        var page = parentPage.createListPage(pageTitle, pageName,'');
        returnObject.message = "Created List page titled " + pageTitle;
        returnObject.url = page.getUrl();
        pageMeister_logPageCreation();
      } catch(err) {
        returnObject.message = err.message;
        returnObject.url = '';
      }
      break;
    default:
      try {
        var templates = site.getTemplates();
        for (var i=0; i<templates.length; i++) {
          if (templates[i].getName()==template) {
            var thisTemplate = templates[i];
          }
        }
        var page = parentPage.createPageFromTemplate(pageTitle, pageName, thisTemplate);
        returnObject.message = "Created " + thisTemplate.getTitle() + " page titled " + pageTitle + ".";
        returnObject.url = page.getUrl();
        pageMeister_logPageCreation();
      } catch(err) {
        returnObject.message = err.message;
        returnObject.url = '';
      }
  }
  try {
    var editors = site.getEditors();
    var editorEmails = [];
    for (var i=0; i<editors.length; i++) {
      editorEmails.push(editors[i].getEmail());
    }
    var viewers = site.getViewers();
    var viewerEmails = [];
    for (var i=0; i<viewers.length; i++) {
      viewerEmails.push(viewers[i].getEmail());
    }
    if (access=="Can edit") {
      if (editorEmails.indexOf(email)!=-1) {
        returnObject.message += " " + email + " was already a site editor.";
      }
      if (viewerEmails.indexOf(email)!=-1) {
        site.removeViewer(email);
        site.addEditor(email)
        returnObject.message += " " + email + " removed as viewer and added as site editor.";
      }
      if ((viewerEmails.indexOf(email)==-1)&&(editorEmails.indexOf(email)==-1)) {
        site.addEditor(email);
        returnObject.message += " Added " + email + " as site editor.";
      }
    }
    if (access=="Can view") {
      if (editorEmails.indexOf(email)!=-1) {
        site.removeEditor(email);
        site.addViewer(email);
        returnObject.message += " " + email + " removed as editor and added as site viewer.";
      }
      if (viewerEmails.indexOf(email)!=-1) {
        returnObject.message += " " + email + " was already a site viewer."; 
      }
      if ((viewerEmails.indexOf(email)==-1)&&(editorEmails.indexOf(email)==-1)) {
        site.addViewer(email);
        returnObject.message += " Added " + email + " as site viewer.";
      }
    }
    try {
      if (notification=="new") {
        MailApp.sendEmail(email, '', "Link to your new site page", "You " + access.toLowerCase() + " your new page: " + pageTitle + " at " + returnObject.url);
        returnObject.message += " Notification email sent.";
      }
    } catch(err1) {
      returnObject.message += err1.message;
    }
  } catch(err) {
    returnObject.message += err.message;
  }
  return returnObject;
}


function verifyUnique(site, pageName) {
  var pages = site.getAllDescendants();
  var allPageNames = [];
  for (var i=0; i<pages.length; i++) {
    allPageNames.push(pages[i].getName());
  }
  var j=0;
  var origPageName = pageName;
  while (allPageNames.indexOf(pageName)!=-1) {
    pageName = origPageName + "_" + j;
    j++;
  }
  return pageName;
}
