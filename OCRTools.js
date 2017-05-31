// AIO Team 
// 2017

//================= OPERATION ===========================

function executeOCRAllFileInFolder() {
  reset()
  var currentSheet = SpreadsheetApp.getActive(); //current spreadsheet
  var folder = DriveApp.getFileById(currentSheet.getId()).getParents().next().getName();
  var contents = DriveApp.getFileById(currentSheet.getId()).getParents().next().getFiles();
  var file;
  var data;
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();
  var rangeCellHighlight = sheet.getRange(1, 1, 1, 3);
  setCellColors(rangeCellHighlight,"white", "gray");//setCellColors("A1:C1", "white", "gray");
  sheet.appendRow(["No", "File Extracted Url", "Content"]);
  while(contents.hasNext()){
    file = contents.next();
    if (file.getMimeType() === MimeType.JPEG || MimeType.PNG || MimeType.PDF){
      if (file.getMimeType() !== MimeType.GOOGLE_SHEETS){
        if ( file.getMimeType() != MimeType.GOOGLE_DOCS){
        var row =  SpreadsheetApp.getActiveSheet().getLastRow();
        var result = executeOCR(file.getId());
          
          data = [ row,
                result[0],
                result[1]
        ];
      
        sheet.appendRow(data);
        updateDataCellLayout();
          
        }
      }
    }
  }
};

function reset(){
   //Delete folder that store all file extracted from image or PDF file.
   var folderResult = DriveApp.getFoldersByName("Extracted File");
   if (folderResult.hasNext()){
    var isDeleted = folderResult.next().setTrashed(true);
   }
}

function executeOCR(fileId) {
  
  var extractedFolderFile;
  var folderExist = DriveApp.getFoldersByName("Extracted File").hasNext();
  var currentSheet = SpreadsheetApp.getActive(); //current spreadsheet
  var folderId = DriveApp.getFileById(currentSheet.getId()).getParents().next().getId();
  var folder = DriveApp.getFolderById(folderId);
  //Checking folder OCR Resource. Existing.
  if (folderExist){
     extractedFolderFile = DriveApp.getFoldersByName("Extracted File").next();
  }else{
    extractedFolderFile = folder.createFolder("Extracted File");
  }
  
  var image = DriveApp.getFileById(fileId);
  var file = {
    title: fileId,
    "parents": [{'id':extractedFolderFile.getId()}],
    mimeType: 'image/png'
  };
  
  // OCR is supported for PDF and image formats
  
  file = Drive.Files.insert(file, image, {ocr: true});
  
  var doc = DocumentApp.openByUrl(file.embedLink);
  var body = doc.getBody().getText();
  return [file.embedLink,body];
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('OCR Tools')
      .addItem('Extract All File', 'executeOCRAllFileInFolder')
      .addSeparator()
      .addSubMenu(ui.createMenu('About US')
          .addItem('Infomation', 'aboutAIO'))
      .addToUi();
}

function aboutAIO() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('We are AIO Team');
}

//========================= Configuration appearance ==============================
function setCellColors(range, foregrd, backgrd) {
   var cell = range;
   cell.setFontColor(foregrd);                     // to set font and
   cell.setBackground(backgrd);                    // background colours.
   cell.setFontSize(12);
   cell.setHorizontalAlignment("center");
   cell.setVerticalAlignment("middle");
   cell. setFontWeight("bold");
 }

function updateDataCellLayout(){
  var sheet = SpreadsheetApp.getActiveSheet();                 // Select the first sheet.
  var allCell = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());
  allCell.setHorizontalAlignment("center");
  allCell.setVerticalAlignment("middle");
}