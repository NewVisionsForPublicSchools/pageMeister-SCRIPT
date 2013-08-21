// This code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit to report our impact to funders
// For original source see http://www.edcode.org

function pageMeister_logPageCreation()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Created%20Sites%20Page", scriptName, scriptTrackingId, systemName)
}


function pageMeister_logFeedbackEmail()
{
 var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Feedback%20Email%20Sent", scriptName, scriptTrackingId, systemName)
}

function pageMeister_logRepeatInstall()
{
 var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("Repeat%20Install", scriptName, scriptTrackingId, systemName)
}

function pageMeister_logFirstInstall()
{
  var systemName = ScriptProperties.getProperty("systemName")
  NVSL.log("First%20Install", scriptName, scriptTrackingId, systemName)
}

function setpageMeisterSid()
{ 
  var pageMeister_sid = ScriptProperties.getProperty("pageMeister_sid");
  if (pageMeister_sid == null || pageMeister_sid == "")
    {
      // user has never installed pageMeister before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
      ScriptProperties.setProperty("pageMeister_sid", ms_str);
      var pageMeister_uid = UserProperties.getProperty("pageMeister_uid");
      if (pageMeister_uid != null || pageMeister_uid != "") {
        pageMeister_logRepeatInstall();
      }else{
        pageMeister_logFirstInstall(); 
      }
    }
}
