var textStyle;

function generatePresentation(){
//  Get chosen presentations and the chosen file name to be saved
  var [srcPresentations, fileName] = getData_();
  
//  Destination folder to save the presentation
  folder = DriveApp.getFolderById("1nowuYSLNcQXi8GFBJspJXw0uZzPqI_Ki");
  
//  Check if a file with the same name already exists
  files = folder.getFilesByName(fileName)

//  If a file already exists with the same name, pop up a dialog to choose to overwrte or to cancel
  if (files.hasNext()) {
    var ui = SpreadsheetApp.getUi(); // Same variations.
    var result = ui.alert(
      'File with the same name already exists',
      'Do you wish to overwrite the existing file?',
      ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      // If "Yes", then delete the existing file
      Drive.Files.remove(files.next().getId());
    } else {
      // If "No", then exit
      ui.alert('Saving cancelled');
      return
    }
  }
    
//  Make a copy of the template presentation 
  var template = DriveApp.getFileById("1bcEBQgSAsI29v1JwYg3E9OS0TRK8HV_a9p5-kP51HhY");
  var templateCopy = template.makeCopy(fileName, folder)
////  Create a google slide file  
//  var tempPresentation = SlidesApp.create(fileName);
////  Locate the file using drive app
//  var tempFile = DriveApp.getFileById(tempPresentation.getId());
////  Save file id
//  var fileId = tempFile.getId();
////  Copy file(with the same id) to the relevant folder
//  folder.addFile(tempFile);
////  Remove the original file in the incorrect folder
//  DriveApp.getRootFolder().removeFile(tempFile);
//
////  Open created empty slide file by id
//  var master = SlidesApp.openById(fileId);

  var master = SlidesApp.openById(templateCopy.getId())
//  Get template text styles
  
  var ts = master.getSlides()[0].getShapes()[0].getText().getTextStyle()
  textStyle = {
    bold: ts.isBold(),
    fontFamily: ts.getFontFamily(),
    fontSize: ts.getFontSize()
  }
  
  var h = master.getPageHeight();
  var w = master.getPageWidth();
  
  //  Remove all existing slides
  master.getSlides().forEach(function(slide) {slide.remove()});
  
  //  Loop through source presentations
  srcPresentations.forEach(function (srcPresentation) {
    //    Insert slides if the file exists
    if ( srcPresentation.fileExists ) {
      var srcSlides = SlidesApp.openById(srcPresentation.id).getSlides()
      //      Get page height and width of master presentation
      
      //      Loop through source slides in a given source presentation
      srcSlides.forEach(function(srcSlide){
        slide = master.appendSlide(srcSlide);
        
        //        Set background as black
        var background = slide.getBackground();
        background.setSolidFill(0, 0, 0);
        
        //        Set text styles
        var textBox = slide.getShapes()[0];
        if (textBox)
          styleText_(textBox, w, h, [255, 255, 255], SlidesApp.ContentAlignment.TOP);  
        
        //        Add notes(bug fix: for some reason, google slides speakernote cannot handle spaces and so we replace spaces with another version
        if (srcPresentation.notes)
          srcSlide.getNotesPage().getSpeakerNotesShape().getText().setText(srcPresentation.notes.replace(/\s/g, '\u00A0'));      
      })
      
    }
    //    Insert a new slide if the file doesn't exist
    else {   
      //      Append a blank slide
      var newSlide = master.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      //      Create a textbox
      var textBox = newSlide.insertShape(SlidesApp.ShapeType.TEXT_BOX);
      //      Set text
      textBox.getText().setText("\""+srcPresentation.title+"\" file does not exist and needs to be created.");
      //      Style text
      styleText_(textBox, w, h, [255, 0, 0], SlidesApp.ContentAlignment.MIDDLE); 
      
//      Add notes
      if (srcPresentation.notes)
          newSlide.getNotesPage().getSpeakerNotesShape().getText().setText(srcPresentation.notes.replace(/\s/g, '\u00A0'));    
    }    
  })
//  Show google slides URL on form
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("googleSlidesUrl").setValue(master.getUrl());
  //  Show google slides URL on form
//  var file = DriveApp.getFileById(master.getId());
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("downloadUrl").setValue("https://docs.google.com/presentation/d/"+master.getId()+"/export/pptx");
}

function styleText_(textBox, w, h, rgbColour, alignVertical) {
  //  Set textbox size to fill the page and to align the text in the middle(vertically)
  textBox
  .setLeft(0)
  .setTop(0)
  .setWidth(w)
  .setHeight(h)
  .setContentAlignment(alignVertical);
  
  //      Get text range
  var textRange = textBox.getText();
  
  //      Change text style to bold and red
  textRange.getTextStyle()
  .setBold(textStyle.bold)
  .setForegroundColor(rgbColour[0], rgbColour[1], rgbColour[2])  
  .setFontFamily(textStyle.fontFamily)  
  .setFontSize(textStyle.fontSize);
  
  //      Align the text centre(horizontally)
  textRange.getParagraphs()[0].getRange().getParagraphStyle()
  .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

function getData_(){
  //  Get Active Sheet which should be the form sheet
  var sheetForm = SpreadsheetApp.getActiveSpreadsheet();
  var dataRaw = sheetForm.getRangeByName("dataRaw").getValues();

  var fileName = sheetForm.getRangeByName("fileName").getValue();
  if (!fileName) {
    var date = new Date();
    fileName = date.getDate() + "." + date.getMonth() + "." + date.getYear();
  }
      
  var dataPresentations = [];
  
  for (var i=0; i < dataRaw.length; i++) {
    //    Check if the entry is non empty
    if ( dataRaw[i][0] ) {
      dataPresentations.push(
        {
          title: dataRaw[i][0], // Song title
          notes: dataRaw[i][1], // Notes for the slide
          fileExists: dataRaw[i][3] == "No File" ? false : true, // Whether or not file is currently in the database
          id: dataRaw[i][4] // File id which will be empty if the file doesn't exist
        }
      )
    }
  }
  
  return [dataPresentations, fileName]
}

/*
* Copyright 2017 Mohsen Mesgarpour
* https://gist.github.com/mesgarpour/07317e81e9ee2b3f1699
*
* Licensed under the Apache License, Version 2.0 (the "License");
* you may not use this file except in compliance with the License.
* You may obtain a copy of the License at
*
*     http://www.apache.org/licenses/LICENSE-2.0
*
* Unless required by applicable law or agreed to in writing, software
* distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and
* limitations under the License.
*
* -----------------------------------------------------------------------------------------------------------
*
* ListFilesFolders script: It is a Google Apps Script, which lists all files and/or folders in a
*             Google Drive folder, and then writes the list into a spreadsheet in batches. The script uses a
*             caching mechanism to allow output recovery after possible crash; however, it won't continue
*             to recover the interrupted script and continue the search.
*             If you need to reduce/remove limitation on script runtime refer to the quotas for
*             Google Services: https://developers.google.com/apps-script/guides/services/quotas
*
* Functions: There are two accessible functions that you may call:
*    - 'run': It lists all folders and optionally files in the specified location, then writes them into
*            the selected spreadsheet.
*    - 'reset': It resets the script global cache. It must be run if the script was interrupted, to
*            clean out the cache and triggers. Moreover, since the outputs are cached and are written
*            after the whole list is created, if you run the script after the crash it would write all
*            the cached output into the sheet, then clears the cache.
*
* Configurations: The following configuration parameters must be configured before running the script.
*    - 'folderId' (type: string):
*          The folder ID. The folder ID is everything after the 'folders/' portion of the URL.
*    - 'searchDepthMax' (type: unsigned integer):
*          The maximum depth for the recursive search of folders.
*    - 'listFiles' (type: boolean):
*          It is a flag to indicate if the listing must include files.
*    - 'cacheTimeout' (type: unsigned integer, in milliseconds):
*          The maximum time that internal cache persists in memory.
*    - 'lockWaitTime' (type: unsigned integer, in milliseconds):
*          The maximum watiting time for the cache reader/writer to wait for the memory lock.
*    - 'appendToSheet' (type: boolean):
*          It is a flag for appending to the selected spreadsheet.
*    - 'writeBatchSize' (type: unsigned integer):
*          The batch size for writing into the spreadsheet.
*
* Algorithm: The written functions uses a recursive function to list files & folders, and it uses
*            a caching mechanisem to save the outputs in cache and write at the end.
*
* -----------------------------------------------------------------------------------------------------------
* Note-1: Because Apps Script services impose daily quotas and hard limitations on some features. If
*         you exceed a quota or limitation, your script will throw an exception and terminate execution.
*
* Note-2: Firstly, set your folder ID ('folderId' variable)! You may copy the folder ID from the
*         browser's address field. The folder ID is everything after the 'folders/' portion of the URL
*         for Google Drive folder.
*         Secondly, set the 'searchDepthMax' to a reasonable number, a very large number can
*         significantly delay the outputs or may cause unexpected termination.
*         Thirdly, set the 'listFiles' to 'false', if you only require the folders to be listed.
*         Finally, other configuration parameters are preconfigured and can be left as default.
*
* Note-3: Because, this is a script and not a certified app, you must grant it permission, to run.
*         When you run it for the first time, a pop-up window will open & asks you for permission.
*
* Note-4: Files and folders must NOT be modified in the selected path, as it may corrupt the
*         generated list.
*
* Note-5: If you interrupt the script you might have to wait a few seconds (maximum 6 minutes),
*         until you can re-run it.
*
* Note-6: It is recommended to first run the script on a small set of files and folders.
*
* Note-7: Make sure there are no other script in the current Google Sheet with similar function or
*         variable names.
*
* Note-8: Refer to version 1.0 of the script, for a simplified version of the ListFilesFolders script.
*
* Note-9: The "teamdrive" does not support "getSharingPermission()", therefore comment out the lines  
*          that uses this function.
*
* Note-10: Note that when you use business version of the Google Drive, there are still limits on how  
*          much time can be spent on auditing (refer to Google's quotas). For instance, the script would 
*          timeout after ten minutes search a folder, therefore avoid auditing very big or deep folders.
*
* -----------------------------------------------------------------------------------------------------------
*
* @version 2.3 (2018.10)
* @see     https://github.com/mesgarpour
*/


// Configurable variables
var folderId = '1LOuGqos7bdY8r-fxcShIrL6H88eaeawJ'; // The folder ID (everything after the 'folders/' portion of the URL).
var searchDepthMax = 100; // Max depth for recursive search of files and folders
var listFiles = true; // flag for listing files
var cacheTimeout = 24 * 60 * 60 * 1000; // set cache time-out
var lockWaitTime = 1 * 60 * 1000; // set maximium watiting time for the cache lock
var appendToSheet = false; // flag for appending to selected spreadsheet
var writeBatchSize = 100; // the write batch size

var sheetMain = SpreadsheetApp.openById("13AdvDvNUE7lxKbIR1OBBBLYrGERl8oIeFLa6aFlstpU");


// ===========================================================================================================
// Global variables
var cacheOutputs = 'InventoryScript_outputs';
var cacheKillFlag = 'InventoryScript_killFlag';


// ===========================================================================================================
// Reset the script cache if it is required to run from the beginning
function reset_() {
  sheetMain.toast('Reseting script...', 'Status', -1);
  
  // reset triggers and delete cache variables
  setKillFlag_(true, this.cacheTimeout);
  deleteTriggers_(this.loopResetGapTime);
  deleteCache_();
  
  sheetMain.toast('Reset is complete!', 'Status', -1);
}


// ===========================================================================================================
// List all folders and files, then write into the current spreadsheet.
function updateDatabase() {
  reset_();
  sheetMain.toast('Executing script...', 'Status', -1);
  
  // load cache
  setKillFlag_(false, this.cacheTimeout);
  var outputRows = getCache_(this.lockWaitTime);
  
  // get list
  if (outputRows === undefined || outputRows === null ||
    outputRows[0] === undefined || outputRows[0] === null) {
    outputRows = [];
    
    outputRows = getChildFiles_(null, DriveApp.getFolderById(this.folderId), 
                                listFiles, cacheTimeout, outputRows);
  }
  
  
  // write list
  writeFolderTree_(outputRows, this.appendToSheet);
  
  sheetMain.toast('Execution is complete!', 'Status', -1);
}


// ===========================================================================================================
// Write the list of folders and files into the spreadsheet
function writeFolderTree_(outputRows, appendToSheet) {
  var sheet = null;
  
  try {
    if (getKillFlag_() === false) {
      // Initialise the spreadsheet
      sheet = sheetMain.getSheetByName("List");
      
      // Write to the selected spreadsheet
      writeOutputs_(sheet, outputRows, appendToSheet);
      
      // reset cache
      reset_();
    }
  } catch (e) {
    sheetMain.toast('Timed out!', 'Status', -1);
  }
}

// ===========================================================================================================
// Get the list of files in the selected folder
function getChildFiles_(parentFolder, childFolder, listFiles, cacheTimeout, outputRows) {
  var childFiles = childFolder.getFiles();
  var childFile = null;
  var path = ""
  try{
    // List files inside the folder
    while (listFiles && childFiles.hasNext()) {
      childFile = childFiles.next();
      
      // derive path
      if (parentFolder === null){
        path = childFolder.getName() + "/" + childFile.getName()
      }else{
        path = parentFolder.getName() + "/" + childFolder.getName() + "/" + childFile.getName()
      }
      
      // Get file information
      //Logger.log("File Name: " + childFile.getName());
      outputRows.push([
        childFile.getName(),
        childFile.getId(),
        childFile.getUrl(),
        path,
        childFile.getName().split('.').pop(),
        childFile.getDateCreated(),
        childFile.getLastUpdated(),
        childFile.getDescription(),
        childFile.getSize(),
        childFile.getOwner().getEmail(),
        childFile.getSharingPermission(),
        childFile.getSharingAccess()
        //, '=HYPERLINK("' + childFile.getUrl() + '", IMAGE("' + Drive.Files.get(childFile.getId()).thumbnailLink + '",1))' //The 'Drive service' is a G-Suite service (commercial service)
      ]);
    }
    
    // Sort items alphabetically by the name
    outputRows.sort(function(a, b) {return a[0] > b[0] ? 1 : -1});
    
    // cache outputs
    setCache_(outputRows, lockWaitTime, cacheTimeout);
  } catch (e) {
    Logger.log('Timed out: Restarting! ' + e.toString());
    sheetMain.toast('Timed out!', 'Status', -1);
  }
  return outputRows;
}


// ===========================================================================================================
// Get the values from cache
function setCache_(outputRows, lockWaitTime, cacheTimeout) {
  try{
    var cache = CacheService.getScriptCache();
    var lock = LockService.getScriptLock();
    
    lock.waitLock(lockWaitTime);
    cache.put(cacheOutputs, JSON.stringify(outputRows), cacheTimeout);
    lock.releaseLock();
  } catch (e) {
    Logger.log('Timed out: Restarting! ' + e.toString());
    sheetMain.toast('Timed out!', 'Status', -1);
  }
}


// ===========================================================================================================
// Set the values in cache
function getCache_(lockWaitTime) {
  try{
    var outputRows = [];
    var cache = CacheService.getScriptCache();
    var lock = LockService.getScriptLock();
    
    lock.waitLock(lockWaitTime);
    outputRows =  JSON.parse(cache.get(cacheOutputs));
    if (outputRows === undefined || outputRows === null ||
        outputRows[0] === undefined || outputRows[0] === null) {
      outputRows = JSON.parse(cache.get(cacheOutputs));
    }
    lock.releaseLock();
  } catch (e) {
    sheetMain.toast('Timed out!', 'Status', -1);
  }
  return outputRows;
}


// ===========================================================================================================
// Write outputs to the selected spreadsheet
function writeOutputs_(sheet, outputRows, appendToSheet) {
  try{
    var range, rowStart, indexStart, indexEnd = null;
    var headerRow = ["Name", "Id", "URL", "Full Path", "Type", "Date", "Last Updated", "Description", "Size",
                     "Owner", "Sharing Permission", "Sharing Access"]; //, "Thumbnail"];
    sheetMain.toast('Writing outputs...', 'Status', -1);
    
    if (sheet !== null && outputRows.length > 0) {
      if (appendToSheet === false) {
        sheet.clear();
        sheet.appendRow(headerRow);
        rowStart = 2;
      } else {
        rowStart = getRowsFilled_(sheet, "A1:A") + 1;
      }
      
      indexStart = 0;
      indexEnd = Math.min(writeBatchSize, outputRows.length);
      
      while (indexStart < outputRows.length) {
        range = sheet.getRange(rowStart + indexStart, 1, indexEnd - indexStart, headerRow.length);
        range.setValues(outputRows.slice(indexStart, indexEnd));
        a = outputRows.slice(indexStart, indexEnd);
        
        indexStart = indexEnd;
        indexEnd =  Math.min(indexStart + writeBatchSize, outputRows.length);
      }
      
      //      range = sheet.getRange(getRowsFilled_(sheet, "A1:A") + 1, 1, 1, 1);
      //      range.setValues([["End of List!"]]);
    }
    
  } catch (e) {
    sheetMain.toast('Timed out!', 'Status', -1);
  }
}


// ===========================================================================================================
// Get number of rows filled in the selected spreadsheet
function getRowsFilled_(sheet, selectedRange) {
  var selectedMatrix = sheet.getRange(selectedRange).getValues();
  return selectedMatrix.filter(String).length;
}


// ===========================================================================================================
// Delete the global cache
function deleteCache_() {
  try{
    var cache = CacheService.getScriptCache();
    var lock = LockService.getScriptLock();
    
    lock.waitLock(this.lockWaitTime);
    cache = CacheService.getScriptCache();
    cache.remove(cacheOutputs);
    lock.releaseLock();
  } catch (e) {
    Logger.log('Failed to delete cache! ' + e.toString());
    sheetMain.toast('Failed to delete cache! Try again in a few minutes.');
  }
}


// ===========================================================================================================
// Delete triggers
function deleteTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  try{
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "run") {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  } catch (e) {
    Logger.log('Failed to delete triggers! ' + e.toString());
    sheetMain.toast('Failed to delete triggers! Try again in a few minutes.');
  }
}


// ===========================================================================================================
// Set kill flag
function setKillFlag_(state, cacheTimeout) {
  var lock = LockService.getScriptLock();
  try{
    lock.waitLock(this.lockWaitTime);
    cache = CacheService.getScriptCache();
    cache.put(cacheKillFlag, state, cacheTimeout);
    lock.releaseLock();
  } catch (e) {
    sheetMain.toast('Failed to set kill flag! Try again in a few minutes.');
  }
}


// ===========================================================================================================
// Get kill flag
function getKillFlag_() {
  killFlag = false;
  try {
    cache = CacheService.getScriptCache();
    //lock.waitLock(this.lockWaitTime);
    killFlag = cache.get(cacheKillFlag) === 'true';
    //lock.releaseLock();
  } catch (e) {
    sheetMain.toast('Failed to set kill flag! Try again in a few minutes.');
  }
  return killFlag;
}
