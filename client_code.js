// ============================================================================
// GOOGLE APPS SCRIPT CLIENT BRIDGE FOR ZIPGRADE CONTROLLER LIBRARY
// ============================================================================
//
// INSTRUCTIONS:
// 1. Copy the code below and paste it into the 'Code.gs' file of your Google Sheet.
// 2. Add the ZipGrade Controller Library to your Google Apps Script project.
//    (Resource > Libraries > Add Library ID of ZipGrade Controller)
// 3. Set the library identifier to 'ZipGradeController'.

function onOpen() {
  ZipGradeController.onOpen();
}

function rebuildFormulasAndFormatting() {
  ZipGradeController.rebuildFormulasAndFormatting();
}

function cleanEmptyRowsAllSheets() {
  ZipGradeController.cleanEmptyRowsAllSheets();
}

function highlightBestStudentOnly() {
  ZipGradeController.highlightBestStudentOnly();
}

function openHtmlManagerModal() {
  ZipGradeController.openHtmlManagerModal();
}

function openHtmlSessionManagerModal() {
  ZipGradeController.openHtmlSessionManagerModal();
}

function gestionarSesionesNativo() {
  ZipGradeController.gestionarSesionesNativo();
}

function openExportJsonModal() {
  ZipGradeController.openExportJsonModal();
}

function openImportJsonModal() {
  ZipGradeController.openImportJsonModal();
}

// ============================================================================
// HTML DIALOG SERVER-SIDE BRIDGE FUNCTIONS
// ============================================================================

function executeCreateFromHtml(data) {
  return ZipGradeController.executeCreateFromHtml(data);
}

function executeFullSyncFromHtml(updatedSubjectsList) {
  return ZipGradeController.executeFullSyncFromHtml(updatedSubjectsList);
}

function createSessionFromHtml(name) {
  return ZipGradeController.createSessionFromHtml(name);
}

function deleteSessionFromHtml(name) {
  return ZipGradeController.deleteSessionFromHtml(name);
}

function syncSessionFromHtml(name) {
  return ZipGradeController.syncSessionFromHtml(name);
}

function executeJsonImport(jsonString) {
  return ZipGradeController.executeJsonImport(jsonString);
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
  if (typeof ZipGradeController !== 'undefined' && ZipGradeController[functionName]) {
    return ZipGradeController[functionName].apply(null, args || []);
  }
  throw new Error("The library or function '" + functionName + "' is not defined in the ZipGradeController library.");
}
