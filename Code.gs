// Copyright 2013 Google Inc. All Rights Reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/**
 * Special function that handles HTTP GET requests to the published web app.
 * @return {HtmlOutput} The HTML page to be served.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Page').evaluate()
      .setTitle('Learning Analytics Package Generator')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}



/**
 * Adds a new task to the task list.
 * @param {String} userEmail user's email to share stuff with them
 * @param {String} sessionName The title of the session, folder, etc. (should be unique for this email).
 * @param {String} sesstionType The template to be copied and used as a base to generate the package
 */
function createPackage(userEmail, sessionName, sessionType) {
  return generateEnactmentSupport(userEmail, sessionName, sessionType)
}


function generateEnactmentSupport(userEmail, sessionName, sessionType) {
  // These are the spreadsheets with the design of the sessions
  //var designURLs = {
    //'lprisan@gmail.com': 'https://docs.google.com/spreadsheets/d/1tTP4on6c7SOm9jMkx27ZyHUd__9ml44Cma6lXOhee4k/edit?usp=sharing',
    //'luis.pablo.prieto@gmail.com': 'https://docs.google.com/spreadsheets/d/1iEwJYfkuErwOI7NLh-dRtm6YitoM8upr6mgsCdplMnY/edit',
    //'pihel.hunt@gmail.com': 'https://docs.google.com/spreadsheets/d/1MSSjKziHiL3-uM7IfzrW5pmbWE85MB0RtX8bVe7YskE/edit',
    //'stella.polikarpus@gmail.com': 'https://docs.google.com/spreadsheets/d/1Evm5zjBUM-fjUxySAaWgf__zdm6z564fTUdCBl2ZQsg/edit',
    //'eve@tlu.ee': 'https://docs.google.com/spreadsheets/d/180VKDCHy8cd7yK_vaLkeeIFuG1MZWBYwMP1G0vfKRjA/edit'
  //};

  var templateURLs = {
    'active-feedback': 'https://docs.google.com/a/tlu.ee/spreadsheets/d/1WmhcLQXXBZuSFyx5w6dbAswUXJOuRxD1FzFs6H0XjVc/edit',
    'living-lab': 'https://docs.google.com/a/tlu.ee/spreadsheets/d/1VoUWGEzw3B2z6sVcjGa7dARapjtkSF6cYq59gtXQrsU/edit'
  };

  var rootFolderId = '0Bzyi_E7EV3EjWVkxMmlDUU5nUTA';

  var template = templateURLs[sessionType];

  if(template && template.length>0){

    //Logger.log(email+': '+design);
    var ss = SpreadsheetApp.openByUrl(template);

    Logger.log('ss '+ss.toString());

    //We create the subfolder on which all objects will be created... if it already exists, it should throw an error
    var folder = createSubFolder(rootFolderId, sessionName, userEmail);
    Logger.log('folder '+folder.toString());

    //We copy the template corresponding to the type of session
    var newMasterSheet = SpreadsheetApp.openById(DriveApp.getFileById(ss.getId()).makeCopy(sessionName, folder).getId());
    Logger.log('newMasterSheet '+newMasterSheet.toString());

    //We create the forms, and return a dictionary/object with the form names and view URLs and responses spreadsheet
    var formsDict = generateForms(newMasterSheet, userEmail, folder, sessionName);
    Logger.log('formsDict '+formsDict.toString());

    //We generate the enactment support
    var website = generateWebsite(newMasterSheet, formsDict, userEmail, folder, sessionName);
    Logger.log('website '+website.toString());

    //We update the spreadsheet with the website and forms
    var emailCode = encodeURIComponent(userEmail).replace('%','').replace('.','');
    var dashboard = "https://luispprieto.shinyapps.io/dashboard5"+emailCode;
    var shorturl = "";//shortenUrl(website);

    updateSpreadsheet(newMasterSheet, formsDict, website, dashboard, shorturl);
    updateDoc(formsDict, website, dashboard, shorturl);

  }else{
    throw ("unknown session type "+sessionType);
  }

  var response = {
    forms: formsDict,
    website: website,
    dashboard: dashboard,
    mastersheet: newMasterSheet.getUrl(),
    folder: folder.getUrl()
  };

  return response;
}

// TODO: This does not work, for some reason!
function shortenUrl(longurl) {
  var url = UrlShortener.Url.insert({
    longUrl: longurl
  });
  Logger.log('Shortened URL is "%s".', url.id);
  return url.id;
}

function updateDoc(formsDict, website, dashboard, shorturl){
  var doc = DocumentApp.openById(website.id);

  substitute(doc, "{{dashboard}}", "dashboard", dashboard);
  if(shorturl.length>0){
    substitute(doc, "{{website}}", shorturl, shorturl);
  }else{
    substitute(doc, "{{website}}", doc.getUrl(), doc.getUrl());
  }
  for(f in formsDict){
    substitute(doc, "{{"+f+"}}", "this form", formsDict[f].view);
  }

  doc.saveAndClose();
}

function substitute(doc, pattern, targettext, targetUrl){
  var range = doc.getBody().findText(pattern);
  while(range){
    Logger.log(range);
    var start = range.getStartOffset();
    var text = range.getElement().asText();
    Logger.log(start+" "+text.getText());
    text.replaceText(pattern,targettext);
    Logger.log(text.getText());
    text.setLinkUrl(start, start+targettext.length-1, targetUrl);
    range = doc.getBody().findText(pattern);
  }

}

//Within the root folder, creates a folder with the name (if not already existing), and within that, creates a folder with the proposed name. If such folder name exists, returns an error
function createSubFolder(rootId, name, email){
  var root = DriveApp.getFolderById(rootId);
  var folder;
  try{
    var folders =  root.getFoldersByName(encodeURIComponent(email));
    if(folders && folders.hasNext()){
        folder = folders.next();
    }else{
        folder =  root.createFolder(encodeURIComponent(email));
    }
  }
  catch(e) {
    folder =  root.createFolder(encodeURIComponent(email));
  }
  folder.addEditor(email);

  var subfolder=null;
  try{
    var folders =  folder.getFoldersByName(name); //if successful, we throw an error!
    if(!folders || !folders.hasNext()){
      subfolder =  folder.createFolder(name);
    }
  }
  catch(e) {
    subfolder =  folder.createFolder(name);
  }
  if(subfolder==null) throw("The session name "+name+" already exists!");
  subfolder.addEditor(email);

  return subfolder;
}


Array.prototype.clean = function(deleteValue) {
  for (var i = 0; i < this.length; i++) {
    if (this[i] == deleteValue) {
      this.splice(i, 1);
      i--;
    }
  }
  return this;
};

//Updates the context of the newly created spreadsheet/design copy
//TODO: maybe other fields could also be updated?
function updateContext(sheet, name){
    var title = sheet.getSheetByName('Context').getRange('A2').setValue(name);
    var date = sheet.getSheetByName('Context').getRange('B2').setValue("");
}

function generateForms(ss, email, folder, sessionName) {
  Logger.log('generating forms ' +ss);
  //We read the title of the design
  updateContext(ss,sessionName);

  var formsDict = {}; //object where we will store the forms URLs

  var quests = ss.getSheetByName("Questions").getRange("A2:Z100").getValues();
  var currentForm = "";
  Logger.log(quests.length+" "+quests);
  for(var i=0; i<quests.length; i++){
    var quest = quests[i];
    var formCode = quest[0];
    if(formCode.length>0){ //If non-empty, we check that it is not the current form and we create it
      if(formsDict[formCode]){
        currentForm = formCode;
      }else if(formCode!=currentForm){
         formsDict[formCode] = {};
         var formName = sessionName+' - '+formCode;
         var form = FormApp.create(formName);
         var file = DriveApp.getFileById(form.getId());
         folder.addFile(file);
         file.addEditor(email);

         form.setTitle(formName);
         if(quest[quest.length-1].length>0) form.setDescription(quest[quest.length-1]);
         //Create the responses sheet
         var responsess = SpreadsheetApp.create(formName+' (Responses)');
         var file2 = DriveApp.getFileById(responsess.getId());
         folder.addFile(file2);
         file2.addEditor(email);


         form.setDestination(FormApp.DestinationType.SPREADSHEET, responsess.getId());
         formsDict[formCode].id = form.getId();
         formsDict[formCode].respss = DriveApp.getFileById(responsess.getId()).getUrl();
         currentForm = formCode;
      }
    } //if it is empty, we continue with the currentForm we had from previous iteration
    Logger.log(currentForm);
    var form = FormApp.openById(formsDict[currentForm].id);

    //We add the current question, depending on its type
    if(quest[2].length>0){
      if(quest[4]=='Short text'){
        var item = form.addTextItem().setTitle(quest[2]);
        if(quest[3]=='Yes') item.setRequired(true);
      }else if(quest[4]=='Long text'){
        var item = form.addParagraphTextItem().setTitle(quest[2]);
        if(quest[3]=='Yes') item.setRequired(true);
      }else if(quest[4]=='Scale'){
        var item = form.addScaleItem().setTitle(quest[2]).setBounds(parseInt(quest[5]),parseInt(quest[6]));
        if(quest[3]=='Yes') item.setRequired(true);
      }else if(quest[4]=='Multi-choice'){
        var item = form.addMultipleChoiceItem().setTitle(quest[2]);
        var choices = [];
        for(var k=5; k<quest.length-1; k++){
          if(quest[k].toString().length>0) choices.push(item.createChoice(quest[k].toString()));
        }
        Logger.log("setting choices for "+quest[2]+choices);
        item.setChoices(choices);
        if(quest[3]=='Yes') item.setRequired(true);
      }else if(quest[4]=='Checkboxes'){
        var item = form.addCheckboxItem().setTitle(quest[2]);
        var choices = [];
        for(var k=5; k<quest.length-1; k++){
          if(quest[k].toString.length>0) choices.push(item.createChoice(quest[k].toString()));
        }
        item.setChoices(choices);
        if(quest[3]=='Yes') item.setRequired(true);
      }else{
        Logger.log('Unknown question type!');
      }
    }
  }

  //We add to the forms array the viewing and editing URLs
  for(var f in formsDict){
    Logger.log('Enriching '+f);
    var form = FormApp.openById(formsDict[f].id);
    formsDict[f].view = form.getPublishedUrl();
    formsDict[f].edit = form.getEditUrl();
  }

  addRecurrentQuestions(ss, formsDict);

  Logger.log('created '+formsDict.length+' forms successfully!');
  return formsDict;
}


//Reads a different worksheet and adds a set of recurring questions to all forms
function addRecurrentQuestions(ss, formsDict){
  Logger.log('adding recurrent questions to forms '+formsDict);
  //Get the questions and specification
  var quests = ss.getSheetByName("RecurrentQuestions").getRange("A2:Z100").getValues();
  Logger.log(quests.length+" "+quests);

  if(quests && quests.length>0){

      for(var f in formsDict){
        Logger.log('Editing '+f)
        var form = FormApp.openById(formsDict[f].id);

        var addedHeader=false;

        for(var i=0; i<quests.length; i++){
          var quest = quests[i];


          //We add the current question, depending on its type
          if(quest[0].length>0){

            if(!addedHeader){//Add form section with recurrent questions
              form.addSectionHeaderItem().setTitle("A few additional questions");
              addedHeader=true;
            }



            if(quest[2]=='Short text'){
              var item = form.addTextItem().setTitle(quest[0]);
              if(quest[1]=='Yes') item.setRequired(true);
            }else if(quest[2]=='Long text'){
              var item = form.addParagraphTextItem().setTitle(quest[0]);
              if(quest[1]=='Yes') item.setRequired(true);
            }else if(quest[2]=='Scale'){
              var item = form.addScaleItem().setTitle(quest[0]).setBounds(parseInt(quest[3]),parseInt(quest[4]));
              if(quest[1]=='Yes') item.setRequired(true);
            }else if(quest[2]=='Multi-choice'){
              var item = form.addMultipleChoiceItem().setTitle(quest[0]);
              var choices = [];
              for(var k=3; k<quest.length-1; k++){
                if(quest[k].toString().length>0) choices.push(item.createChoice(quest[k].toString()));
              }
              Logger.log("setting choices for "+quest[0]+choices);
              item.setChoices(choices);
              if(quest[1]=='Yes') item.setRequired(true);
            }else if(quest[2]=='Checkboxes'){
              var item = form.addCheckboxItem().setTitle(quest[0]);
              var choices = [];
              for(var k=3; k<quest.length-1; k++){
                if(quest[k].toString.length>0) choices.push(item.createChoice(quest[k].toString()));
              }
              item.setChoices(choices);
              if(quest[1]=='Yes') item.setRequired(true);
            }else{
              Logger.log('Unknown question type!');
            }
          }//if question text


        }//for quests

      }// for forms




  }//if quest.length>0



}



function generateWebsite(ss, formsDict, email, folder, sessionName) {

  //This is the document we'll create with the links to forms, dashboard...
  var doc = DocumentApp.create(sessionName+' - Session description');
  var file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  file.addEditor(email);

  //Add title in the document
  doc.getBody().appendParagraph(sessionName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  //doc.getBody().appendParagraph(date).setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  // We extract the goals and add them to the doc
  var goals = ss.getSheetByName('Goals').getRange('A2:B8').getValues()
  Logger.log(goals);
  doc.getBody().appendParagraph("Goals").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  for(var i=0; i<goals.length; i++){
    var goal = goals[i];
    if(goal[0].length > 0){
      doc.getBody().appendListItem(goal[0]+" ("+goal[1]+")");
    }
  }
  // We extract the activities and add them to the doc
  var acts = ss.getSheetByName('Activities').getRange('A2:D11').getValues()
  Logger.log(acts);
  doc.getBody().appendParagraph("Activities").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  for(var i=0; i<acts.length; i++){
    var act = acts[i];
    if(act[0].length > 0){
      doc.getBody().appendListItem(act[0]+" ["+act[1]+"']");
      var descs = act[2].match(/[^\r\n]+/g); //Lines of description
      Logger.log(descs);
      for(var j=0; j<descs.length; j++){
        var desc = descs[j];
        doc.getBody().appendListItem(desc).setNestingLevel(1).setGlyphType(DocumentApp.GlyphType.BULLET);
      }
      if(act[3].length>0){ //Additional resources
        doc.getBody().appendListItem("Other resources").setNestingLevel(1).setGlyphType(DocumentApp.GlyphType.BULLET);
        var resources = act[3].match(/[^\r\n]+/g);
        for(var k=0; k<resources.length; k++){
          var resource = resources[k];
          var sp = resource.split("---");
          if(sp.length==1){
            Logger.log(sp[0]);
            //if(sp[0].startsWith("http")){//The resource is directly the link
              //doc.getBody().appendListItem(sp[0]).setNestingLevel(2).setGlyphType(DocumentApp.GlyphType.BULLET).setLinkUrl(sp[0]);
            //}else{ //The resource is a text only
              doc.getBody().appendListItem(sp[0]).setNestingLevel(2).setGlyphType(DocumentApp.GlyphType.BULLET)
            //}
          }else if(sp.length==2){//The resource is a pair text,link
            doc.getBody().appendListItem(sp[0]).setNestingLevel(2).setGlyphType(DocumentApp.GlyphType.BULLET).setLinkUrl(sp[1]);
          }

        }
      }
    }
  }

  //We share the document openly, and return the URL
  docId = doc.getId();
  var file = DriveApp.getFileById(docId)
  file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  var website = {};
  website.url = file.getUrl();
  website.id = docId;
  Logger.log(website);
  return website;
}

function updateSpreadsheet(ss, formsDict, website, dashboard, shorturl){
  //We open the Automation sheet, or else we create it
  var auto = ss.getSheetByName('Automation');
  if (auto == null) {
    auto = ss.insertSheet('Automation');
  }
  //We clear the sheet!!
  auto.clear();
  auto.getRange('A1').setBackground("grey").setValue('Resource');
  auto.getRange('B1').setBackground("grey").setValue('URL view');
  auto.getRange('C1').setBackground("grey").setValue('URL edit');
  auto.getRange('D1').setBackground("grey").setValue('URL responses');
  auto.getRange('E1').setBackground("grey").setValue('GDrive ID');

  auto.getRange('A2').setBackground("grey").setValue('Website');
  auto.getRange('B2').setBackground("grey").setValue(website.url);
  auto.getRange('C2').setBackground("grey").setValue('');
  auto.getRange('D2').setBackground("grey").setValue('');
  auto.getRange('E2').setBackground("grey").setValue(website.id);
  if(shorturl.length>0) auto.getRange('F2').setBackground("grey").setValue(shorturl);

  auto.getRange('A3').setBackground("grey").setValue('Dashboard');
  auto.getRange('B3').setBackground("grey").setValue(dashboard);
  auto.getRange('C3').setBackground("grey").setValue('');
  auto.getRange('D3').setBackground("grey").setValue('');
  auto.getRange('E3').setBackground("grey").setValue('');

  var counter=4;
  for(f in formsDict){
    auto.getRange('A'+counter).setBackground("grey").setValue(f);
    auto.getRange('B'+counter).setBackground("grey").setValue(formsDict[f].view);
    auto.getRange('C'+counter).setBackground("grey").setValue(formsDict[f].edit);
    auto.getRange('D'+counter).setBackground("grey").setValue(formsDict[f].respss);
    auto.getRange('E'+counter).setBackground("grey").setValue(formsDict[f].id);
    counter++;
  }
}
