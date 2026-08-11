// ============================================================================
// GOOGLE APPS SCRIPT CLIENT BRIDGE FOR ZIPGRADE CONTROLLER LIBRARY
// ============================================================================
//
// INSTRUCTIONS:
// 1. Copy the code below and paste it into the 'Code.gs' file of your Google Sheet.
// 2. Add the ZipGrade Controller Library to your Google Apps Script project.
//    (Resource > Libraries > Add Library ID of ZipGrade Controller)
// 3. Set the library identifier to 'ZipGradeControllerMain'.

function onOpen() {
  ZipGradeControllerMain.onOpen();
}

function rebuildFormulasAndFormatting() {
  ZipGradeControllerMain.rebuildFormulasAndFormatting();
}

function cleanEmptyRowsAllSheets() {
  ZipGradeControllerMain.cleanEmptyRowsAllSheets();
}

function highlightBestStudentOnly() {
  ZipGradeControllerMain.highlightBestStudentOnly();
}

function toggleSortConsolidado() {
  ZipGradeControllerMain.toggleSortConsolidado();
}

function openHtmlManagerModal() {
  ZipGradeControllerMain.openHtmlManagerModal();
}

function openHtmlSessionManagerModal() {
  ZipGradeControllerMain.openHtmlSessionManagerModal();
}

function gestionarSesionesNativo() {
  ZipGradeControllerMain.gestionarSesionesNativo();
}

function openExportJsonModal() {
  ZipGradeControllerMain.openExportJsonModal();
}

function openImportJsonModal() {
  ZipGradeControllerMain.openImportJsonModal();
}

function noSheetsAlert() {
  ZipGradeControllerMain.noSheetsAlert();
}

function executeOption1() { ZipGradeControllerMain.processBySheetIndex(1); }
function executeOption2() { ZipGradeControllerMain.processBySheetIndex(2); }
function executeOption3() { ZipGradeControllerMain.processBySheetIndex(3); }
function executeOption4() { ZipGradeControllerMain.processBySheetIndex(4); }
function executeOption5() { ZipGradeControllerMain.processBySheetIndex(5); }
function executeOption6() { ZipGradeControllerMain.processBySheetIndex(6); }
function executeOption7() { ZipGradeControllerMain.processBySheetIndex(7); }
function executeOption8() { ZipGradeControllerMain.processBySheetIndex(8); }
function executeOption9() { ZipGradeControllerMain.processBySheetIndex(9); }
function executeOption10() { ZipGradeControllerMain.processBySheetIndex(10); }

// ============================================================================
// HTML DIALOG SERVER-SIDE BRIDGE FUNCTIONS
// ============================================================================

function executeCreateFromHtml(data) {
  return ZipGradeControllerMain.executeCreateFromHtml(data);
}

function executeFullSyncFromHtml(updatedSubjectsList) {
  return ZipGradeControllerMain.executeFullSyncFromHtml(updatedSubjectsList);
}

function createSessionFromHtml(name) {
  return ZipGradeControllerMain.createSessionFromHtml(name);
}

function deleteSessionFromHtml(name) {
  return ZipGradeControllerMain.deleteSessionFromHtml(name);
}

function syncSessionFromHtml(name) {
  return ZipGradeControllerMain.syncSessionFromHtml(name);
}

function executeJsonImport(jsonString) {
  return ZipGradeControllerMain.executeJsonImport(jsonString);
}

// ============================================================================
// UNIVERSAL DISPATCHER BRIDGE FUNCTION
// ============================================================================

/**
 * Universal dispatcher bridge to run any function from the library using google.script.run
 * without having to define individual bridge wrappers.
 * 
 * @param {string} functionName Name of the function in the library to run.
 * @param {Array} args Array of arguments to pass to the function.
 * @return {*} The result returned by the library function.
 */
function callLibrary(functionName, args) {
  if (typeof ZipGradeControllerMain !== 'undefined' && ZipGradeControllerMain[functionName]) {
    return ZipGradeControllerMain[functionName].apply(null, args || []);
  }
  throw new Error("The library or function '" + functionName + "' is not defined in the ZipGradeControllerMain library.");
}
