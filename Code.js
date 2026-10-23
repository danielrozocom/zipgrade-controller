/**
 * EVALUATION INTEGRATED SYSTEM
 * ZipGrade Sync + Consolidated Auto-Formulas + Drag & Drop Management (UX/UI en Español)
 */

// ============================================================================
// 1. AUTOMATIC INITIALIZATION & MENUS
// ============================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  try {
    if (!isSpreadsheetStructureValid()) {
      rebuildFormulasAndFormatting();
    } else {
      // Sincronizar siempre el nombre del curso en A1 con el nombre actual del archivo
      updateCourseHeader();
    }
  } catch (e) {
    console.log("Deferred initialization: " + e.message);
  }

  // --- SUBMENU 1: ZipGrade ---
  var zipGradeMenu = ui.createMenu('📝 ZipGrade');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheets = getSessionTargetSheets();
  var addedCount = 0;

  for (var i = 0; i < targetSheets.length; i++) {
    var sheetName = targetSheets[i];
    addedCount++;
    if (addedCount <= 10) {
      zipGradeMenu.addItem('➡️ Enviar a: ' + sheetName, 'executeOption' + addedCount);
    }
  }

  if (addedCount === 0) {
    zipGradeMenu.addItem('⚠️ No hay hojas destino válidas', 'noSheetsAlert');
  }

  // --- SUBMENU 2: Configuración y Estructura ---
  var tableMenu = ui.createMenu('⚙️ Configuración y Estructura')
    .addItem('⚡ Recalcular y Corregir Fórmulas', 'rebuildFormulasAndFormatting')
    .addItem('🧹 Limpiar Filas Vacías', 'cleanEmptyRowsAllSheets')
    .addItem('🏆 Resaltar Mejor Promedio', 'highlightBestStudentOnly')
    .addItem('🔢 Ordenar por Total / Restaurar Alfabético', 'toggleSortConsolidado')
    .addSeparator()
    .addItem('📚 Gestión de Asignaturas', 'openHtmlManagerModal')
    .addItem('📂 Gestión de Sesiones', 'openHtmlSessionManagerModal')
    .addItem('🔍 Reporte de Inconsistencias de Estudiantes', 'gestionarSesionesNativo')
    .addSeparator()
    .addItem('📥 Exportar Configuración JSON', 'openExportJsonModal')
    .addItem('📤 Importar Configuración JSON', 'openImportJsonModal');

  zipGradeMenu.addToUi();
  tableMenu.addToUi();
}

function noSheetsAlert() {
  SpreadsheetApp.getUi().alert("No se encontraron hojas de destino válidas.");
}

function executeOption1() { processBySheetIndex(1); }
function executeOption2() { processBySheetIndex(2); }
function executeOption3() { processBySheetIndex(3); }
function executeOption4() { processBySheetIndex(4); }
function executeOption5() { processBySheetIndex(5); }
function executeOption6() { processBySheetIndex(6); }
function executeOption7() { processBySheetIndex(7); }
function executeOption8() { processBySheetIndex(8); }
function executeOption9() { processBySheetIndex(9); }
function executeOption10() { processBySheetIndex(10); }

function processBySheetIndex(index) {
  var validSheets = getSessionTargetSheets();
  if (index > validSheets.length) return;
  var targetSheetName = validSheets[index - 1];
  processZipGradeData(targetSheetName);
}

function getSessionTargetSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var validSheets = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().trim();
    if (/^S\d+$/i.test(name)) {
      validSheets.push(name);
    }
  }
  return validSheets;
}

// ============================================================================
// 2. ZIPGRADE MODULE (VALIDACIÓN S1 Y S2)
// ============================================================================

function getSourceDataSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  // 1. Preferencia: hoja llamada Sheet1 (cualquier combinación de mayúsculas/minúsculas)
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toUpperCase() === "SHEET1") {
      return sheets[i];
    }
  }

  // 2. Fallback: primera hoja que no sea sesión (S#) ni Consolidado
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().toUpperCase();
    if (name.indexOf("CONSOLID") !== -1) continue;
    if (/^S\d+$/.test(name)) continue;
    return sheets[i];
  }

  // 3. Último recurso: hoja activa
  return ss.getActiveSheet();
}

function processZipGradeData(targetSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = getSourceDataSheet();
  var targetSheet = ss.getSheetByName(targetSheetName);
  var ui = SpreadsheetApp.getUi();

  if (!sourceSheet || !targetSheet) return;

  var lastSourceRow = sourceSheet.getLastRow();
  if (lastSourceRow < 2) return;

  var rawData = sourceSheet.getRange(2, 1, lastSourceRow - 1, sourceSheet.getLastColumn()).getValues();

  var filteredData = rawData.filter(function (row) {
    return row[0] && row[0].toString().trim() !== "";
  });

  if (filteredData.length === 0) {
    ss.toast("No hay estudiantes válidos en " + sourceSheet.getName() + ".", "⚠️ Error", 4);
    return;
  }

  filteredData.sort(function (a, b) {
    return a[0].toString().localeCompare(b[0].toString());
  });

  var sessNumMatch = targetSheetName.match(/S(\d+)/i);
  var isSession1 = sessNumMatch && parseInt(sessNumMatch[1], 10) === 1;

  if (isSession1) {
    writeSheetData(targetSheet, filteredData, true);
    ss.toast("Procesado exitosamente en " + targetSheetName, "✅ Éxito", 3);
  } else {
    var sheetS1 = ss.getSheetByName("S1") || ss.getSheetByName("s1");
    if (!sheetS1) {
      writeSheetData(targetSheet, filteredData, false);
      ss.toast("Procesado en " + targetSheetName, "✅ Éxito", 3);
      return;
    }

    var lastRowS1 = sheetS1.getLastRow();
    if (lastRowS1 < 2) {
      ui.alert("⚠️ Advertencia: Primero debes procesar la Sesión 1 (S1) para establecer la lista maestra de estudiantes.");
      return;
    }

    var masterListS1 = sheetS1.getRange(2, 1, lastRowS1 - 1, 2).getValues();

    var sheet1Map = {};
    filteredData.forEach(function (row) {
      var key = fullNameKey(row[0], row[1]);
      sheet1Map[key] = row.slice(2);
    });

    var gradeColumnsCount = filteredData[0].slice(2).length;
    var adjustedDataS2 = [];
    var missingInS2 = [];

    masterListS1.forEach(function (studentS1) {
      var lastName = studentS1[0].toString().trim();
      var firstName = studentS1[1].toString().trim();
      var key = fullNameKey(lastName, firstName);

      if (sheet1Map[key]) {
        adjustedDataS2.push([lastName, firstName].concat(sheet1Map[key]));
        delete sheet1Map[key];
      } else {
        var zeroScores = Array(gradeColumnsCount).fill(0);
        adjustedDataS2.push([lastName, firstName].concat(zeroScores));
        missingInS2.push(lastName + " " + firstName);
      }
    });

    var extraStudents = [];
    Object.keys(sheet1Map).forEach(function (key) {
      extraStudents.push(key);
    });

    if (extraStudents.length > 0 || missingInS2.length > 0) {
      var alertMsg = "⚠️ Se detectaron inconsistencias de estudiantes al procesar " + targetSheetName + ":\n\n";
      if (missingInS2.length > 0) {
        alertMsg += "Estudiantes en S1 pero FALTAN en esta sesión (se les asignó nota 0):\n- " + missingInS2.join("\n- ") + "\n\n";
      }
      if (extraStudents.length > 0) {
        alertMsg += "Estudiantes en esta sesión pero NO registrados en S1 (no se guardarán):\n- " + extraStudents.join("\n- ") + "\n\n";
      }
      var response = ui.alert("⚠️ Inconsistencia Detectada", alertMsg + "¿Deseas continuar con el procesamiento de todos modos?", ui.ButtonSet.YES_NO);
      if (response !== ui.Button.YES) {
        ss.toast("Procesamiento cancelado. Revisa la lista de estudiantes.", "⚠️ Cancelado", 4);
        return;
      }
    }

    writeAlignedS2Data(targetSheet, adjustedDataS2, gradeColumnsCount);
    ss.toast("Procesado en " + targetSheetName + " (Alineado con S1)", "✅ Éxito", 3);
  }

  rebuildFormulasAndFormatting();
}

function writeSheetData(targetSheet, data, includeNames) {
  var names = [];
  var grades = [];

  for (var j = 0; j < data.length; j++) {
    names.push([data[j][0], data[j][1]]);
    grades.push(data[j].slice(2));
  }

  var requiredRows = data.length + 1;
  var maxRows = targetSheet.getMaxRows();
  var maxCols = targetSheet.getMaxColumns();
  var gradeColsCount = grades[0].length;
  var totalUsedCols = gradeColsCount + 2;

  // Asegurar que hay suficientes columnas en la hoja
  if (maxCols < totalUsedCols) {
    targetSheet.insertColumnsAfter(maxCols, totalUsedCols - maxCols);
    maxCols = targetSheet.getMaxColumns();
  }

  if (maxRows > 1) {
    targetSheet.getRange(1, 1, maxRows, maxCols).setBorder(false, false, false, false, false, false);
    if (includeNames) {
      targetSheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
    } else if (maxCols >= 3) {
      targetSheet.getRange(2, 3, maxRows - 1, maxCols - 2).clearContent();
    }
  }

  if (includeNames) {
    targetSheet.getRange(2, 1, names.length, 2).setValues(names);
  }
  targetSheet.getRange(2, 3, grades.length, gradeColsCount).setValues(grades);

  if (maxRows > requiredRows) {
    targetSheet.deleteRows(requiredRows + 1, maxRows - requiredRows);
  }

  applySessionSheetFormatting(targetSheet);
}

function writeAlignedS2Data(targetSheet, alignedData, gradeColsCount) {
  var grades = [];

  for (var j = 0; j < alignedData.length; j++) {
    grades.push(alignedData[j].slice(2));
  }

  var requiredRows = alignedData.length + 1;
  var maxRows = targetSheet.getMaxRows();
  var maxCols = targetSheet.getMaxColumns();
  var totalUsedCols = gradeColsCount + 2;

  // Asegurar que hay suficientes columnas en la hoja
  if (maxCols < totalUsedCols) {
    targetSheet.insertColumnsAfter(maxCols, totalUsedCols - maxCols);
    maxCols = targetSheet.getMaxColumns();
  }

  if (maxRows > 1) {
    targetSheet.getRange(1, 1, maxRows, maxCols).setBorder(false, false, false, false, false, false);
    if (maxCols >= 3) {
      targetSheet.getRange(2, 3, maxRows - 1, maxCols - 2).clearContent();
    }
  }

  targetSheet.getRange(2, 3, grades.length, gradeColsCount).setValues(grades);

  // Si la sesión usa nombres por fórmula, garantizar que las fórmulas cubran
  // todos los estudiantes de S1 (por si la lista maestra creció).
  ensureFormulaNames(targetSheet, alignedData.length);

  if (maxRows > requiredRows) {
    targetSheet.deleteRows(requiredRows + 1, maxRows - requiredRows);
  }

  applySessionSheetFormatting(targetSheet);
}

function ensureFormulaNames(sheet, studentCount) {
  try {
    var formulaCells = [];
    for (var r = 2; r <= studentCount + 1; r++) {
      formulaCells.push([
        "=IF('S1'!A" + r + "=\"\",\"\",'S1'!A" + r + ")",
        "=IF('S1'!B" + r + "=\"\",\"\",'S1'!B" + r + ")"
      ]);
    }
    if (formulaCells.length > 0) {
      sheet.getRange(2, 1, formulaCells.length, 2).setFormulas(formulaCells);
    }
  } catch (e) { }
}

function ensureDynamicNamesForAllSessions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetS1 = ss.getSheetByName("S1") || ss.getSheetByName("s1");
  if (!sheetS1) return;

  var lastRow1 = sheetS1.getLastRow();
  var numStudents = 0;
  if (lastRow1 >= 2) {
    var vals = sheetS1.getRange(2, 1, lastRow1 - 1, 2).getValues();
    for (var i = 0; i < vals.length; i++) {
      var l = String(vals[i][0] || '').trim();
      var f = String(vals[i][1] || '').trim();
      if (l !== '' || f !== '') numStudents = i + 1;
    }
  }
  if (numStudents < 1) return;

  var excluded = ["S1", "SHEET1", "CONSOLIDADO"];
  var sheets = ss.getSheets();
  sheets.forEach(function (s) {
    var name = s.getName().toUpperCase();
    if (excluded.indexOf(name) !== -1) return;
    ensureFormulaNames(s, numStudents);
  });
}

function normalizeNameKey(text) {
  return String(text || '')
    .toUpperCase()
    .replace(/[ÁÀÄÂÃ]/g, 'A')
    .replace(/[ÉÈËÊ]/g, 'E')
    .replace(/[ÍÌÏÎ]/g, 'I')
    .replace(/[ÓÒÖÔÕ]/g, 'O')
    .replace(/[ÚÙÜÛ]/g, 'U')
    .replace(/Ñ/g, 'N')
    .replace(/\s+/g, ' ')
    .trim();
}

// Clave invariante al reparto entre columnas (apellidos/nombres): usa el nombre completo.
function fullNameKey(last, first) {
  return normalizeNameKey(String(last || '') + " " + String(first || ''));
}

// ============================================================================
// 3. CENTRO DE GESTIÓN HTML (UX/UI EN ESPAÑOL)
// ============================================================================

function openHtmlManagerModal() {
  var ui = SpreadsheetApp.getUi();
  var subjectsData = getSubjectsBySession();

  var htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
          * { box-sizing: border-box; }
          body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 12px; color: #2c3e50; margin: 0; background: #ffffff; }
          
          .tabs { display: flex; border-bottom: 2px solid #3b5998; margin-bottom: 12px; background: #fff; border-radius: 6px 6px 0 0; }
          .tab-btn { flex: 1; padding: 10px; text-align: center; cursor: pointer; font-weight: 600; font-size: 12px; color: #555; border: none; background: none; }
          .tab-btn.active { color: #fff; background: #3b5998; border-radius: 4px 4px 0 0; }
          
          .tab-content { display: none; }
          .tab-content.active { display: block; }

          .form-group { margin-bottom: 12px; }
          label { display: block; font-weight: 600; margin-bottom: 4px; font-size: 11px; color: #1a237e; }
          input, select { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 12px; }

          .cards-list { background: #ffffff; border: 1px dashed #cbd5e1; border-radius: 6px; padding: 8px; height: 320px; overflow-y: auto; }
          .session-section-header { font-size: 12px; font-weight: bold; text-align: center; padding: 6px; border-radius: 4px; margin: 10px 0 8px 0; color: white; }
          .session-section-header:first-child { margin-top: 0; }
          .s1-title { background-color: #8D7BA6; }
          .s2-title { background-color: #B23A48; }

          .subject-card { background: #fff; border: 1px solid #e2e8f0; border-left: 4px solid #3b5998; padding: 8px; margin-bottom: 8px; border-radius: 4px; cursor: grab; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
          .subject-card:active { cursor: grabbing; transform: scale(0.98); }
          .subject-card.dragging { opacity: 0.4; }
          
          .card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; }
          .card-handle { font-size: 11px; font-weight: bold; color: #64748b; cursor: grab; }
          .card-inputs { display: flex; gap: 6px; margin-top: 4px; }
          .card-inputs input { padding: 4px; font-size: 11px; }

          .btn-danger { background: #ef4444; color: white; border: none; padding: 2px 6px; border-radius: 3px; cursor: pointer; font-size: 10px; }
          
          .footer-actions { text-align: right; margin-top: 15px; }
          .footer-actions button { display: inline-flex; align-items: center; justify-content: center; gap: 6px; }
          button.btn-primary { background-color: #3b5998; color: white; padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; font-size: 12px; }
          button.btn-primary:hover { opacity: 0.9; }
          button.btn-cancel { background-color: #e0e0e0; color: #333; padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; font-size: 12px; }
          button.btn-secondary { background-color: #10b981; color: white; padding: 8px 12px; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; font-size: 11px; }
          button.btn-secondary:hover { opacity: 0.9; }
          button.btn-secondary-export { background-color: #f59e0b; color: white; padding: 8px 12px; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; font-size: 11px; }
          button.btn-secondary-export:hover { opacity: 0.9; }
          
          #errorBox { color: #d32f2f; font-size: 11px; margin-top: 8px; display: none; background: #ffebee; padding: 6px; border-radius: 4px; }
        </style>
      </head>
      <body>
        
        <div style="background-color: #fff9e6; border-left: 4px solid #ffb020; padding: 8px; margin-bottom: 10px; font-size: 11px; border-radius: 4px; color: #7a5200; font-weight: 500;">
          ⚠️ <b>Nota importante:</b> Si tienes más de una cuenta de Google abierta en este navegador, Google Apps Script podría experimentar problemas al cargar o guardar cambios en este gestor.
        </div>

        <div class="tabs">
          <button class="tab-btn active" onclick="switchTab('reorderTab')"><i class="fa-solid fa-list-ol"></i> Reordenar / Editar</button>
          <button class="tab-btn" onclick="switchTab('createTab')"><i class="fa-solid fa-plus-circle"></i> Crear Asignatura</button>
        </div>

        <div id="reorderTab" class="tab-content active">
          <div style="font-size: 11px; color: #64748b; margin-bottom: 8px;">
            💡 <b>Arrastra y suelta</b> las tarjetas para ordenarlas. Edita el nombre o cantidad de preguntas in-situ:
          </div>

          <div class="cards-list" id="cardsList" ondragover="allowDrop(event)" ondrop="drop(event)">
            <!-- Tarjetas generadas dinámicamente -->
          </div>

          <div id="emptySubjectsMsg" style="display:none; background:#fff3cd; border-left:4px solid #ffc107; color:#856404; font-size:12px; padding:8px; border-radius:4px; margin:8px 0;">
            ⚠️ <b>No se detectaron asignaturas.</b> Cierra este gestor, ejecuta <b>⚡ Recalcular y Corregir Fórmulas</b> y vuelve a abrirlo.
          </div>

          <div class="footer-actions" style="display: flex; justify-content: space-between; align-items: center; margin-top: 15px;">
            <div style="display: flex; gap: 6px;">
              <button type="button" class="btn-secondary" onclick="triggerImportJson()" title="Importar JSON"><i class="fa-solid fa-file-import"></i> Importar JSON</button>
              <button type="button" class="btn-secondary-export" onclick="exportJson()" title="Exportar JSON"><i class="fa-solid fa-file-export"></i> Exportar JSON</button>
            </div>
            <div style="display: flex; gap: 8px;">
              <button type="button" class="btn-cancel" onclick="google.script.host.close()">Cancelar</button>
              <button class="btn-primary" id="saveReorderBtn" onclick="saveAllChanges()"><i class="fa-solid fa-floppy-disk"></i> Guardar Cambios</button>
            </div>
          </div>
        </div>

        <div id="createTab" class="tab-content">
          <form id="createForm">
            <div class="form-group">
              <label for="sessionNumber">1. Sesión de la Asignatura:</label>
              <select id="sessionNumber" name="sessionNumber" onchange="updatePositionOptions()" required>
                <!-- Opciones generadas dinámicamente -->
              </select>
            </div>

            <div class="form-group">
              <label for="subjectName">2. Nombre de la Asignatura:</label>
              <input type="text" id="subjectName" name="subjectName" placeholder="Asignatura" oninput="this.value = this.value.toUpperCase()" required autofocus>
            </div>

            <div class="form-group">
              <label for="questionCount">3. Cantidad de Preguntas:</label>
              <input type="number" id="questionCount" name="questionCount" placeholder="Cantidad de Preguntas" required min="1">
            </div>

            <div class="form-group">
              <label for="position">4. Ubicación en la Sesión:</label>
              <select id="position" name="position" required></select>
            </div>

            <div class="footer-actions">
              <button type="button" class="btn-cancel" onclick="google.script.host.close()">Cancelar</button>
              <button type="submit" id="createBtn" class="btn-primary">➕ Crear Asignatura</button>
            </div>
          </form>
        </div>

        <div id="errorBox"></div>
        <input type="file" id="importJsonFile" accept=".json" style="display:none;" onchange="handleImportJsonFile(event)" />

        <script>
          const initialData = ${JSON.stringify(subjectsData)};
          let subjectsList = [];
          let draggedCardId = null;
          let sessionKeys = [];

          function init() {
            subjectsList = [];
            sessionKeys = Object.keys(initialData).sort((a, b) => {
              return parseInt(a.replace('s', ''), 10) - parseInt(b.replace('s', ''), 10);
            });

            const sessionSelect = document.getElementById('sessionNumber');

            sessionSelect.innerHTML = '<option value="" disabled selected>Selecciona sesión</option>';

            sessionKeys.forEach(key => {
              const sessNum = key.replace('s', '');
              const sessName = 'SESIÓN ' + sessNum;

              // Generar opción select para la pestaña "Crear Asignatura"
              const opt = document.createElement('option');
              opt.value = sessNum;
              opt.textContent = sessName;
              sessionSelect.appendChild(opt);

              // Agregar asignaturas a la lista
              if (initialData[key]) {
                subjectsList = subjectsList.concat(initialData[key]);
              }
            });

            renderCards();
            updatePositionOptions();
          }

          function exportJson() {
            const cleanList = subjectsList.map(s => ({
              name: s.name,
              count: s.count,
              session: s.session
            }));
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(cleanList, null, 2));
            const downloadAnchor = document.createElement('a');
            downloadAnchor.setAttribute("href", dataStr);
            downloadAnchor.setAttribute("download", "zipgrade_config.json");
            document.body.appendChild(downloadAnchor);
            downloadAnchor.click();
            downloadAnchor.remove();
          }

          function triggerImportJson() {
            document.getElementById('importJsonFile').click();
          }

          function handleImportJsonFile(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
              try {
                const imported = JSON.parse(e.target.result);
                if (!Array.isArray(imported)) {
                  alert("El archivo JSON debe contener un arreglo de asignaturas.");
                  return;
                }

                const newSubjects = [];
                const existingMap = {};
                subjectsList.forEach(s => {
                  const key = s.name.toUpperCase().trim() + "_" + s.session;
                  existingMap[key] = s.colIndex;
                });

                const assignedCols = new Set();
                let nextNewColIndex = 1000;

                imported.forEach((item, idx) => {
                  if (!item.name || !item.session || isNaN(item.count)) {
                    throw new Error("Formato incorrecto en el elemento " + (idx + 1) + ". Debe tener name, count y session.");
                  }

                  const name = String(item.name).toUpperCase().trim();
                  const session = String(item.session).trim();
                  const count = parseInt(item.count, 10);
                  const key = name + "_" + session;

                  let colIndex = existingMap[key];
                  if (colIndex && !assignedCols.has(colIndex)) {
                    assignedCols.add(colIndex);
                  } else {
                    colIndex = nextNewColIndex++;
                  }

                  newSubjects.push({
                    colIndex: colIndex,
                    name: name,
                    count: count,
                    session: session,
                    colLetter: '?'
                  });
                });

                subjectsList = newSubjects;
                
                newSubjects.forEach(s => {
                  const sKey = 's' + s.session;
                  if (!sessionKeys.includes(sKey)) {
                    sessionKeys.push(sKey);
                  }
                });
                sessionKeys.sort((a, b) => parseInt(a.replace('s', ''), 10) - parseInt(b.replace('s', ''), 10));

                renderCards();
                updatePositionOptions();
                alert("Configuración importada con éxito en la vista previa. Haz clic en 'Guardar Cambios' para aplicarla.");
              } catch (err) {
                alert("Error al importar el JSON: " + err.message);
              }
            };
            reader.readAsText(file);
            event.target.value = '';
          }

          function switchTab(tabId) {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

            if (tabId === 'reorderTab') {
              document.querySelectorAll('.tab-btn')[0].classList.add('active');
              document.getElementById('reorderTab').classList.add('active');
            } else {
              document.querySelectorAll('.tab-btn')[1].classList.add('active');
              document.getElementById('createTab').classList.add('active');
            }
          }

          function renderCards() {
            const listContainer = document.getElementById('cardsList');
            listContainer.innerHTML = '';
            const renderedColIndexes = {};

            sessionKeys.forEach(key => {
              const sessNum = key.replace('s', '');
              const sessName = 'SESIÓN ' + sessNum;
              const titleClass = (parseInt(sessNum, 10) % 2 === 1) ? 's1-title' : 's2-title';

              // Encabezado de sección (no arrastrable)
              const header = document.createElement('div');
              header.className = 'session-section-header ' + titleClass;
              header.id = 'sessHeader_' + sessNum;
              header.innerText = sessName;
              listContainer.appendChild(header);

              const sessionSubjects = subjectsList.filter(s => s.session === sessNum);
              sessionSubjects.forEach(subj => {
                const card = document.createElement('div');
                card.className = 'subject-card';
                card.draggable = true;
                card.id = 'card_' + subj.colIndex;

                card.ondragstart = (e) => {
                  draggedCardId = card.id;
                  card.classList.add('dragging');
                };

                card.ondragend = (e) => { card.classList.remove('dragging'); };

                card.innerHTML = 
                  '<div class="card-header">' +
                  '  <span class="card-handle"><i class="fa-solid fa-grip-lines"></i> Col. ' + subj.colLetter + '</span>' +
                  '  <button class="btn-danger" type="button" onclick="deleteSubject(' + subj.colIndex + ')" title="Eliminar"><i class="fa-solid fa-trash-can"></i></button>' +
                  '</div>' +
                  '<div class="card-inputs">' +
                  '  <input type="text" style="flex: 2;" value="' + subj.name + '" oninput="this.value = this.value.toUpperCase()" onchange="updateSubjectField(' + subj.colIndex + ', \\'name\\', this.value)">' +
                  '  <input type="number" style="flex: 1;" value="' + subj.count + '" onchange="updateSubjectField(' + subj.colIndex + ', \\'count\\', parseInt(this.value, 10))">' +
                  '</div>';

                listContainer.appendChild(card);
                renderedColIndexes[subj.colIndex] = true;
              });
            });

            // Seguridad: si alguna asignatura quedara sin sesión válida, agregarla al final
            subjectsList.forEach(subj => {
              if (!renderedColIndexes[subj.colIndex]) {
                const card = document.createElement('div');
                card.className = 'subject-card';
                card.draggable = true;
                card.id = 'card_' + subj.colIndex;

                card.ondragstart = (e) => {
                  draggedCardId = card.id;
                  card.classList.add('dragging');
                };
                card.ondragend = (e) => { card.classList.remove('dragging'); };

                card.innerHTML = 
                  '<div class="card-header">' +
                  '  <span class="card-handle"><i class="fa-solid fa-grip-lines"></i> Col. ' + subj.colLetter + '</span>' +
                  '  <button class="btn-danger" type="button" onclick="deleteSubject(' + subj.colIndex + ')" title="Eliminar"><i class="fa-solid fa-trash-can"></i></button>' +
                  '</div>' +
                  '<div class="card-inputs">' +
                  '  <input type="text" style="flex: 2;" value="' + subj.name + '" oninput="this.value = this.value.toUpperCase()" onchange="updateSubjectField(' + subj.colIndex + ', \\'name\\', this.value)">' +
                  '  <input type="number" style="flex: 1;" value="' + subj.count + '" onchange="updateSubjectField(' + subj.colIndex + ', \\'count\\', parseInt(this.value, 10))">' +
                  '</div>';

                listContainer.appendChild(card);
              }
            });

            const emptyMsg = document.getElementById('emptySubjectsMsg');
            if (emptyMsg) {
              emptyMsg.style.display = (subjectsList.length === 0) ? 'block' : 'none';
            }
          }

          function updateSubjectField(colIndex, field, value) {
            const subj = subjectsList.find(s => s.colIndex === colIndex);
            if (subj) subj[field] = value;
          }

          function deleteSubject(colIndex) {
            if (confirm('¿Deseas eliminar esta asignatura del Consolidado?')) {
              subjectsList = subjectsList.filter(s => s.colIndex !== colIndex);
              renderCards();
            }
          }

          function allowDrop(e) { e.preventDefault(); }

          function drop(e) {
            e.preventDefault();
            const cardElem = document.getElementById(draggedCardId);
            if (!cardElem) return;
            const container = document.getElementById('cardsList');

            const afterElement = getDragAfterElement(container, e.clientY);
            if (afterElement == null) container.appendChild(cardElem);
            else container.insertBefore(cardElem, afterElement);

            updateSubjectsArrayOrder();
          }

          function getDragAfterElement(container, y) {
            const draggableElements = [...container.querySelectorAll('.subject-card:not(.dragging), .session-section-header')];
            return draggableElements.reduce((closest, child) => {
              const box = child.getBoundingClientRect();
              const offset = y - box.top - box.height / 2;
              if (offset < 0 && offset > closest.offset) {
                return { offset: offset, element: child };
              } else return closest;
            }, { offset: Number.NEGATIVE_INFINITY }).element;
          }

          function updateSubjectsArrayOrder() {
            const listContainer = document.getElementById('cardsList');
            const newOrder = [];
            let currentSession = sessionKeys.length > 0 ? sessionKeys[0].replace('s', '') : '1';

            const children = listContainer.children;
            for (let i = 0; i < children.length; i++) {
              const el = children[i];
              if (el.classList.contains('session-section-header')) {
                currentSession = el.id.replace('sessHeader_', '');
              } else if (el.classList.contains('subject-card')) {
                const colIdx = parseInt(el.id.replace('card_', ''), 10);
                const subj = subjectsList.find(s => s.colIndex === colIdx);
                if (subj) {
                  subj.session = String(currentSession);
                  newOrder.push(subj);
                }
              }
            }
            subjectsList = newOrder;
            renderCards();
          }

          function updatePositionOptions() {
            const session = document.getElementById('sessionNumber').value;
            const posSelect = document.getElementById('position');
            posSelect.innerHTML = '';

            if (!session) {
              posSelect.innerHTML = '<option value="">Selecciona una sesión primero</option>';
              return;
            }

            const list = subjectsList.filter(s => s.session === session);

            if (list.length === 0) {
              posSelect.innerHTML = '<option value="end">Al Inicio / Final de la Sesión</option>';
              return;
            }

            let opts = '<option value="end">Al Final de la Sesión</option>';
            opts += '<option value="start">Al Inicio de la Sesión</option>';
            list.forEach(subj => {
              opts += \`<option value="before_\${subj.colIndex}">Antes de: \${subj.name}</option>\`;
            });
            posSelect.innerHTML = opts;
          }

          function saveAllChanges() {
            const btn = document.getElementById('saveReorderBtn');
            const errBox = document.getElementById('errorBox');
            btn.disabled = true;
            btn.innerText = 'Guardando cambios...';
            errBox.style.display = 'none';

            google.script.run
              .withSuccessHandler(() => google.script.host.close())
              .withFailureHandler(err => {
                btn.disabled = false;
                btn.innerText = '💾 Guardar Cambios';
                errBox.innerText = 'Error: ' + err.message;
                errBox.style.display = 'block';
              })
              .executeFullSyncFromHtml(subjectsList);
          }

          document.getElementById('createForm').addEventListener('submit', function(e) {
            e.preventDefault();
            const btn = document.getElementById('createBtn');
            const errBox = document.getElementById('errorBox');
            btn.disabled = true;
            btn.innerText = 'Procesando...';
            errBox.style.display = 'none';

            const data = {
              session: document.getElementById('sessionNumber').value,
              name: document.getElementById('subjectName').value,
              count: parseInt(document.getElementById('questionCount').value, 10),
              position: document.getElementById('position').value
            };

            google.script.run
              .withSuccessHandler(() => google.script.host.close())
              .withFailureHandler(err => {
                btn.disabled = false;
                btn.innerText = '➕ Crear Asignatura';
                errBox.innerText = 'Error: ' + err.message;
                errBox.style.display = 'block';
              })
              .executeCreateFromHtml(data);
          });

          init();
        </script>
      </body>
    </html>
  `;

  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(460)
    .setHeight(480)
    .setTitle('⚙️ Gestión de Asignaturas');

  ui.showModalDialog(htmlOutput, '⚙️ Gestión de Asignaturas');
}

function executeCreateFromHtml(data) {
  var sheet = getConsolidatedSheet();
  if (!sheet) throw new Error("No se encontró la hoja Consolidado");

  var rawName = data.name.trim().toUpperCase();
  var cleanName = rawName.replace(/\s*\(\d+.*?\)/g, '').trim();
  if (!cleanName) throw new Error("Por favor ingresa un nombre válido.");

  var session = data.session;
  var pos = data.position;

  var targetCol = getTargetColumnForInsert(sheet, session, pos);
  sheet.insertColumnBefore(targetCol);

  // Escribir el identificador de sesión en la fila 1
  sheet.getRange(1, targetCol).setValue("SESIÓN " + session);

  setupSubjectHeader(sheet, targetCol, cleanName, '(' + data.count + ')', '#527FD9');
  sheet.setColumnWidth(targetCol, 100);

  rebuildFormulasAndFormatting();
  SpreadsheetApp.getActiveSpreadsheet().toast('✨ Asignatura "' + cleanName + '" creada con éxito.', '¡Éxito!', 5);
}

function executeFullSyncFromHtml(updatedSubjectsList) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getConsolidatedSheet();
  if (!sheet) throw new Error("No se encontró la hoja Consolidado");

  var studentEndRow = getStudentRowsEndIndex(sheet);
  var maxCols = sheet.getLastColumn();

  var totalColIndex = -1;
  for (var c = 3; c <= maxCols; c++) {
    var val2 = String(sheet.getRange(2, c).getValue() || '').toUpperCase();
    if (val2.includes('TOTAL')) {
      totalColIndex = c;
      break;
    }
  }

  if (totalColIndex === -1) totalColIndex = maxCols - 1;

  var existingSubjectsData = getSubjectsBySession();
  var existingAll = [];
  Object.keys(existingSubjectsData).forEach(function (key) {
    existingAll = existingAll.concat(existingSubjectsData[key]);
  });

  // Protección contra pérdida de datos: si el gestor no detectó asignaturas pero el
  // Consolidado sí tiene, cancelamos en lugar de borrar todas las columnas.
  if ((!updatedSubjectsList || updatedSubjectsList.length === 0) && existingAll.length > 0) {
    throw new Error("El gestor no cargó las asignaturas existentes. No se guardaron cambios para evitar pérdida de datos. Vuelve a abrir el gestor o ejecuta ⚡ Recalcular.");
  }

  var columnDataMap = {};

  existingAll.forEach(function (s) {
    columnDataMap[s.colIndex] = sheet.getRange(4, s.colIndex, studentEndRow - 3, 1).getValues();
  });

  var colsToDeleteCount = totalColIndex - 3;
  if (colsToDeleteCount > 0) {
    sheet.deleteColumns(3, colsToDeleteCount);
  }

  updatedSubjectsList.forEach(function (subj, idx) {
    var colToInsert = 3 + idx;
    sheet.insertColumnBefore(colToInsert);

    // Escribir el identificador de sesión en la fila 1
    sheet.getRange(1, colToInsert).setValue("SESIÓN " + subj.session);

    var cleanName = subj.name.replace(/\s*\(\d+.*?\)/g, '').trim();
    setupSubjectHeader(sheet, colToInsert, cleanName, '(' + subj.count + ')', "#527FD9");
    sheet.setColumnWidth(colToInsert, 100);

    if (columnDataMap[subj.colIndex]) {
      sheet.getRange(4, colToInsert, studentEndRow - 3, 1).setValues(columnDataMap[subj.colIndex]);
    } else {
      sheet.getRange(4, colToInsert, studentEndRow - 3, 1).setValues(Array(studentEndRow - 3).fill([0]));
    }
  });

  rebuildFormulasAndFormatting();
  ss.toast("🔄 Asignaturas reordenadas y sincronizadas con éxito.", "¡Éxito!", 5);
}

// ============================================================================
// 4. MOTOR ULTRARRÁPIDO DE FÓRMULAS Y FORMATO DE TABLA
// ============================================================================

function isSpreadsheetStructureValid() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getConsolidatedSheet();
  if (!sheet) return false;

  var studentEndRow = getStudentRowsEndIndex(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < studentEndRow + 1) return false;

  // Verificar que TOTAL y PROMEDIO estén en su posición correcta
  var totalVal = String(sheet.getRange(studentEndRow, 1).getValue() || '').trim().toUpperCase();
  var avgVal = String(sheet.getRange(studentEndRow + 1, 1).getValue() || '').trim().toUpperCase();
  if (totalVal !== 'TOTAL' || avgVal !== 'PROMEDIO') return false;

  // Verificar que los nombres de estudiantes sean fórmulas hacia S1
  try {
    var formulaA = sheet.getRange(4, 1).getFormula();
    var formulaB = sheet.getRange(4, 2).getFormula();
    if (!formulaA || formulaA.indexOf('S1') === -1 || !formulaB || formulaB.indexOf('S1') === -1) {
      return false;
    }
  } catch (e) {
    return false;
  }

  // Verificar que las filas TOTAL y PROMEDIO tengan fórmulas y formato de número
  try {
    var totalF = sheet.getRange(studentEndRow, 3).getFormula();
    var avgF = sheet.getRange(studentEndRow + 1, 3).getFormula();
    if (String(totalF || '').indexOf('=') !== 0 || String(avgF || '').indexOf('=') !== 0) {
      return false;
    }
    var avgFmt = sheet.getRange(studentEndRow + 1, 3).getNumberFormat();
    if (String(avgFmt || '').indexOf('0.00') === -1) {
      return false;
    }
  } catch (e) {
    return false;
  }

  // Verificar que existan las columnas TOTAL y PROMEDIO
  var lastCol = sheet.getLastColumn();
  var row1Vals = readRowValues(sheet, 1, lastCol);
  var row2Vals = readRowValues(sheet, 2, lastCol);
  var totalCol = findTotalColFromRows(row1Vals, row2Vals, lastCol);
  if (totalCol === -1) return false;

  var hasAvg = false;
  for (var c = 0; c < lastCol; c++) {
    if (String(row1Vals[c] || '').toUpperCase().indexOf('PROMEDIO') !== -1 ||
      String(row2Vals[c] || '').toUpperCase().indexOf('PROMEDIO') !== -1) {
      hasAvg = true;
      break;
    }
  }
  if (!hasAvg) return false;

  return true;
}

function rebuildFormulasAndFormatting() {
  var sheet = getConsolidatedSheet();
  if (!sheet) return;

  var numStudents = getValidStudentCount();
  var targetStudentRows = (numStudents > 0) ? numStudents : 40;
  var newTotalRow = 4 + targetStudentRows;

  // Buscar y limpiar las filas viejas de TOTAL/PROMEDIO si están en una posición incorrecta (ej. fila 21 y 22)
  try {
    var oldTotalRow = -1;
    var maxRows = sheet.getMaxRows();
    if (maxRows >= 4) {
      var valsA = sheet.getRange(1, 1, maxRows, 1).getValues();
      for (var r = 3; r < maxRows; r++) {
        var txt = String(valsA[r][0] || '').trim().toUpperCase();
        if (txt === 'TOTAL' || txt === 'PROMEDIO') {
          oldTotalRow = r + 1;
          break;
        }
      }
    }

    if (oldTotalRow > 0 && oldTotalRow !== newTotalRow) {
      var lastCol = sheet.getLastColumn();
      if (lastCol > 0) {
        // Limpiar contenido, formatos y descombinar A:B en las dos filas viejas de totales
        var oldRange = sheet.getRange(oldTotalRow, 1, 2, lastCol);
        oldRange.clearContent().clearFormat();
        sheet.getRange(oldTotalRow, 1, 1, 2).breakApart();
        sheet.getRange(oldTotalRow + 1, 1, 1, 2).breakApart();
      }
    }
  } catch (e) { }

  // 1. Limpiar/Alinear filas sobrantes o faltantes en todas las hojas primero
  cleanEmptyRowsAllSheets();

  // 2. Calcular índices y aplicar formatos y fórmulas
  var studentEndRow = getStudentRowsEndIndex(sheet);
  recalculateQuestionRangesAndFormulas(sheet);
  realignSessionHeaders(sheet);
  updateCalculatedColumns(sheet, studentEndRow);
  applyTableFormatting(sheet, studentEndRow);
  ensureTotalAverageRows(sheet, studentEndRow);

  // Formatear automáticamente todas las hojas de sesiones activas
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var excluded = ["SHEET1", "CONSOLIDADO"];
    sheets.forEach(function (s) {
      var name = s.getName();
      if (excluded.indexOf(name.toUpperCase()) === -1) {
        applySessionSheetFormatting(s);
      }
    });
  } catch (e) { }

  // Convertir los nombres de todas las sesiones (menos S1) a FÓRMULAS hacia S1,
  // para que apellidos y nombres sean SIEMPRE dinámicos.
  ensureDynamicNamesForAllSessions();

  // Actualizar texto inteligente del Curso (A1) con lo que esté después de '|'
  updateCourseHeader(sheet);
}

function extractCourseFromSpreadsheetName(name) {
  var rawName = String(name || '').trim();
  if (!rawName) return "CURSO";

  // 1. Extraer lo que esté después del delimitador '|' (no solo nombres específicos sino cualquier texto)
  if (rawName.indexOf("|") !== -1) {
    var parts = rawName.split("|").map(function (p) { return p.trim(); }).filter(Boolean);
    if (parts.length > 1) {
      // Buscar el segmento después del prefijo principal que NO sea un indicador de sesión (S#) o periodo (P#)
      for (var i = 1; i < parts.length; i++) {
        var token = parts[i];
        if (!/^(S\d+|P\d+)$/i.test(token)) {
          return token;
        }
      }
      return parts[1];
    }
  }

  // 2. Fallback si no tiene '|': detectar rango o grado (ej. "3° - 5°", "601", "1002", "9°A")
  var rangeMatch = rawName.match(/\d{1,2}\s*[º°ª]?[A-Za-z]?\s*(?:[-–]|a)\s*\d{1,2}\s*[º°ª]?[A-Za-z]?/i);
  if (rangeMatch && !/^(S\d+|P\d+)$/i.test(rangeMatch[0])) {
    return rangeMatch[0].trim();
  }
  var degMatch = rawName.match(/\b\d{1,2}[º°ª][A-Za-z]?\b/);
  if (degMatch) {
    return degMatch[0].trim();
  }
  var classCodeMatch = rawName.match(/\b(\d{3,4})\b/);
  if (classCodeMatch) {
    return classCodeMatch[1].trim();
  }

  return "CURSO";
}

function updateCourseHeader(sheet) {
  try {
    if (!sheet) sheet = getConsolidatedSheet();
    if (!sheet) return;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssName = ss ? ss.getName() : "";
    var cursoText = extractCourseFromSpreadsheetName(ssName);

    var cellA1 = sheet.getRange(1, 1);
    cellA1.setValue(cursoText);
    cellA1.setFontFamily("Nunito")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  } catch (e) {
    console.log("Error al actualizar encabezado del curso: " + e.message);
  }
}

function getConsolidatedSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Consolidado") || ss.getSheetByName("CONSOLIDADO");
  if (!sheet) {
    try {
      var allSheets = ss.getSheets();
      for (var i = 0; i < allSheets.length; i++) {
        if (allSheets[i].getName().toUpperCase().indexOf("CONSOLID") !== -1) {
          sheet = allSheets[i];
          break;
        }
      }
    } catch (e) { }
  }
  if (!sheet) sheet = ss.getActiveSheet();
  return sheet;
}

function getStudentRowsEndIndex(sheet) {
  var numStudents = getValidStudentCount();
  var targetStudentRows = (numStudents > 0) ? numStudents : 40;
  return 4 + targetStudentRows;
}

function getSubjectsBySession() {
  var sheet = getConsolidatedSheet();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var activeSessions = [];

  sheets.forEach(function (s) {
    var name = s.getName().toUpperCase();
    if (name !== "CONSOLIDADO" && name !== "SHEET1" && /^S\d+$/.test(name)) {
      activeSessions.push(name.replace("S", ""));
    }
  });

  activeSessions.sort(function (a, b) { return parseInt(a, 10) - parseInt(b, 10); });
  if (activeSessions.length === 0) activeSessions = ["1", "2"];

  var result = {};
  activeSessions.forEach(function (sessNum) {
    result["s" + sessNum] = [];
  });

  if (!sheet) return result;
  var maxCols = sheet.getLastColumn();
  if (maxCols < 3) return result;

  var row1Vals = sheet.getRange(1, 1, 1, maxCols).getValues()[0];
  var row2Values = sheet.getRange(2, 1, 1, maxCols).getValues()[0];
  var row4Formulas = [];
  try { row4Formulas = sheet.getRange(4, 1, 1, maxCols).getFormulas()[0]; } catch (e) { }
  var currentSession = "1";

  // Mapa de texto de sesión de la fila 1 para celdas combinadas
  var row1MergeMap = {};
  try {
    var merges = sheet.getRange(1, 3, 1, maxCols - 2).getMergedRanges();
    for (var mi = 0; mi < merges.length; mi++) {
      var mr = merges[mi];
      var mv = String(mr.getValue() || '').toUpperCase();
      for (var mc = mr.getColumn(); mc < mr.getColumn() + mr.getNumColumns(); mc++) {
        row1MergeMap[mc] = mv;
      }
    }
  } catch (e) { }

  for (var c = 3; c <= maxCols; c++) {
    var val2 = String(row2Values[c - 1] || '').trim();

    if (val2.toUpperCase().indexOf('TOTAL') !== -1) break;

    // Sesión: prioridad a la fórmula de la columna (=SUM('S1'!...)) que es la fuente más confiable
    var formulaSession = getSessionFromFormulaString(row4Formulas[c - 1]);
    if (formulaSession) {
      currentSession = formulaSession;
    } else {
      var cellText = row1MergeMap[c] || String(row1Vals[c - 1] || '').toUpperCase();
      var match = cellText.match(/SESIÓ[N\s]+(\d+)/) || cellText.match(/(\d+)\s*SESIÓ/) || cellText.match(/S(\d+)/);
      if (match) {
        currentSession = match[1];
      }
    }

    if (val2) {
      var cleanName = val2.split('\n')[0].replace(/\s*\(\d+.*?\)/g, '').trim();
      var countMatch = val2.match(/\((\d+)-(\d+)\)/) || val2.match(/\((\d+)\)/);
      var qCount = 10;
      if (countMatch) {
        if (countMatch[2]) {
          qCount = parseInt(countMatch[2], 10) - parseInt(countMatch[1], 10) + 1;
        } else {
          qCount = parseInt(countMatch[1], 10);
        }
      }

      var item = { colIndex: c, name: cleanName, count: qCount, colLetter: convertColToLetter(c), session: currentSession };
      var key = "s" + currentSession;
      if (!result[key]) {
        result[key] = [];
      }
      result[key].push(item);
    }
  }

  return result;
}

function getTargetColumnForInsert(sheet, sessionNumber, position) {
  var maxCols = sheet.getLastColumn();
  var row1Vals = readRowValues(sheet, 1, maxCols);
  var row2Vals = readRowValues(sheet, 2, maxCols);
  var totalCol = findTotalColFromRows(row1Vals, row2Vals, maxCols);
  if (totalCol === -1) totalCol = maxCols;

  if (position && position.startsWith("before_")) {
    var colBefore = parseInt(position.replace("before_", ""), 10);
    if (colBefore >= totalCol) return totalCol;
    return colBefore;
  }

  if (position === "start") {
    if (String(sessionNumber) === "1") return 3;
    var sessionMap = getSessionColumnMap(sheet, maxCols);
    for (var c = 3; c < totalCol; c++) {
      if (sessionMap[c] === String(sessionNumber)) {
        return c;
      }
    }
  }

  return findEndOfSessionColumn(sheet, sessionNumber);
}

function getSessionSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    var targetNum = String(sheetName).replace(/\D/g, '');
    var allS = ss.getSheets();
    for (var i = 0; i < allS.length; i++) {
      var nm = allS[i].getName().toUpperCase();
      if (nm.indexOf('CONSOLID') === -1 && nm.indexOf('SHEET1') === -1 && nm.replace(/\D/g, '') === targetNum && nm.indexOf('S') !== -1) {
        sheet = allS[i];
        break;
      }
    }
  }
  return sheet;
}

function ensureQuestionColumnsExist(sheetName, maxQuestion) {
  var sheet = getSessionSheet(sheetName);
  if (!sheet) return;

  var currentCols = sheet.getLastColumn();
  var neededCols = 2 + maxQuestion;

  if (currentCols < neededCols) {
    var headers = [];
    var startCol = currentCols + 1;
    for (var q = (currentCols - 2) + 1; q <= maxQuestion; q++) {
      headers.push("#" + q);
    }
    if (headers.length > 0) {
      sheet.getRange(1, startCol, 1, headers.length).setValues([headers])
        .setBackground("#35528C")
        .setFontColor("#FFFFFF")
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    }
  }
}

function trimSessionQuestionColumns(sheetName, maxQuestion) {
  var sheet = getSessionSheet(sheetName);
  if (!sheet) return;
  var neededCols = 2 + maxQuestion;
  var currentCols = sheet.getLastColumn();
  if (currentCols <= neededCols) return;

  var numRows = sheet.getLastRow();
  var allEmpty = true;

  // Revisar SOLO las filas de datos (desde la fila 2). Los encabezados #X no cuentan como contenido.
  if (numRows >= 2) {
    var extraRange = sheet.getRange(2, neededCols + 1, numRows - 1, currentCols - neededCols);
    var values = extraRange.getValues();
    for (var r = 0; r < values.length && allEmpty; r++) {
      for (var c = 0; c < values[r].length; c++) {
        var cell = values[r][c];
        if (cell !== null && cell !== undefined && String(cell).trim() !== '') {
          allEmpty = false;
          break;
        }
      }
    }
  }

  if (allEmpty) {
    sheet.deleteColumns(neededCols + 1, currentCols - neededCols);
  }
}

function findEndOfSessionColumn(sheet, sessionNumber) {
  var maxCols = sheet.getLastColumn();
  var row1Vals = readRowValues(sheet, 1, maxCols);
  var row2Vals = readRowValues(sheet, 2, maxCols);
  var totalCol = findTotalColFromRows(row1Vals, row2Vals, maxCols);
  if (totalCol === -1) totalCol = maxCols;

  var sessionMap = getSessionColumnMap(sheet, maxCols);
  var targetSessionInt = parseInt(sessionNumber, 10);

  for (var c = 3; c < totalCol; c++) {
    var sess = sessionMap[c];
    if (sess && parseInt(sess, 10) > targetSessionInt) {
      return c;
    }
  }
  return totalCol;
}

function getSessionFromFormulaString(formula) {
  var m = String(formula || '').match(/['"]?S(\d+)['"]?!/i);
  return m ? m[1] : null;
}

function getSessionFromFormula(sheet, colIndex) {
  try {
    return getSessionFromFormulaString(sheet.getRange(4, colIndex).getFormula());
  } catch (e) { }
  return null;
}

function readRowValues(sheet, rowNumber, maxCols) {
  try {
    return sheet.getRange(rowNumber, 1, 1, maxCols).getValues()[0];
  } catch (e) {
    return [];
  }
}

function findTotalColFromRows(row1, row2, maxCols) {
  for (var c = 0; c < maxCols; c++) {
    var v1 = String(row1[c] || '').toUpperCase();
    var v2 = String(row2[c] || '').toUpperCase();
    if (v1.indexOf('TOTAL') !== -1 || v2.indexOf('TOTAL') !== -1) {
      return c + 1;
    }
  }
  return -1;
}

function getSessionColumnMap(sheet, maxCols) {
  var map = {};
  var row1Vals = readRowValues(sheet, 1, maxCols);
  try {
    var merges = sheet.getRange(1, 1, 1, maxCols).getMergedRanges();
    for (var mi = 0; mi < merges.length; mi++) {
      var mr = merges[mi];
      var mv = String(mr.getValue() || '').toUpperCase();
      if (mv.indexOf('TOTAL') !== -1) continue;
      var mm = mv.match(/SESIÓ[N\s]+(\d+)/) || mv.match(/(\d+)\s*SESIÓ/) || mv.match(/S(\d+)/);
      if (mm) {
        for (var mc = mr.getColumn(); mc < mr.getColumn() + mr.getNumColumns(); mc++) {
          map[mc] = mm[1];
        }
      }
    }
  } catch (e) { }
  for (var c = 1; c <= maxCols; c++) {
    if (map[c]) continue;
    var v = String(row1Vals[c - 1] || '').toUpperCase();
    if (v.indexOf('TOTAL') !== -1) continue;
    var m2 = v.match(/SESIÓ[N\s]+(\d+)/) || v.match(/(\d+)\s*SESIÓ/) || v.match(/S(\d+)/);
    if (m2) map[c] = m2[1];
  }
  return map;
}

function recalculateQuestionRangesAndFormulas(sheet) {
  var maxCols = sheet.getLastColumn();
  var studentEndRow = getStudentRowsEndIndex(sheet);
  var numDataRows = studentEndRow - 4; // Solo las filas de estudiantes (sin incluir TOTAL/PROMEDIO)
  if (numDataRows < 1) return;

  var row1Vals = sheet.getRange(1, 1, 1, maxCols).getValues()[0];
  var row2Values = sheet.getRange(2, 1, 1, maxCols).getValues()[0];
  var row4Formulas = sheet.getRange(4, 1, 1, maxCols).getFormulas()[0];

  var currentSession = "1";
  var currentQuestionCounts = {}; // Mapa dinámico para contar preguntas por sesión: { "1": 0, "2": 0... }
  var sessionMaxQ = {}; // Máximo de preguntas por sesión, para recortar columnas sobrantes
  var subjectIndex = 0;

  var totalCol = findTotalColFromRows(row1Vals, row2Values, maxCols);
  if (totalCol === -1) totalCol = maxCols - 1;
  if (totalCol <= 3) return;

  // Mapa de texto de sesión de la fila 1 ANTES de desagrupar (para celdas combinadas)
  var row1MergeMap = {};
  try {
    var preMerges = sheet.getRange(1, 3, 1, totalCol - 3).getMergedRanges();
    for (var mi = 0; mi < preMerges.length; mi++) {
      var mr = preMerges[mi];
      var mv = String(mr.getValue() || '').toUpperCase();
      for (var mc = mr.getColumn(); mc < mr.getColumn() + mr.getNumColumns(); mc++) {
        row1MergeMap[mc] = mv;
      }
    }
  } catch (e) { }

  // Desagrupar la fila 1 completa para poder reescribir las sesiones de forma confiable
  try {
    var row1Merges = sheet.getRange(1, 3, 1, totalCol - 3).getMergedRanges();
    for (var mi2 = 0; mi2 < row1Merges.length; mi2++) {
      row1Merges[mi2].breakApart();
    }
  } catch (e) { }

  // Desagrupar de una vez todas las cabeceras de asignatura (filas 2-3) para evitar llamadas repetidas
  try {
    var hdrMerges = sheet.getRange(2, 3, 2, totalCol - 3).getMergedRanges();
    for (var hmi = 0; hmi < hdrMerges.length; hmi++) {
      hdrMerges[hmi].breakApart();
    }
  } catch (e) { }

  // Fuente de sesión más confiable: el nombre de la hoja dentro de cada fórmula (=SUM('S1'!...))
  var formulaSessionMap = {};
  for (var cf = 3; cf < totalCol; cf++) {
    var fs = getSessionFromFormulaString(row4Formulas[cf - 1]);
    if (fs) formulaSessionMap[cf] = fs;
  }

  // Etiquetas de sesión para reescribir la fila 1 de una sola vez al final
  var sessionLabels = [];
  for (var si = 0; si < totalCol - 3; si++) {
    sessionLabels.push(String(row1Vals[2 + si] || ''));
  }

  for (var c = 3; c < totalCol; c++) {
    var val2 = String(row2Values[c - 1] || '').trim();

    // Prioridad: fórmula de la columna > texto de la fila 1 > sesión arrastrada
    if (formulaSessionMap[c]) {
      currentSession = formulaSessionMap[c];
    } else {
      var cellText = row1MergeMap[c] || String(row1Vals[c - 1] || '').toUpperCase();
      var match = cellText.match(/SESIÓ[N\s]+(\d+)/) || cellText.match(/(\d+)\s*SESIÓ/) || cellText.match(/S(\d+)/);
      if (match) {
        currentSession = match[1];
      }
    }

    if (val2) {
      var cleanName = val2.split('\n')[0].replace(/\s*\(\d+.*?\)/g, '').trim();
      var countMatch = val2.match(/\((\d+)-(\d+)\)/) || val2.match(/\((\d+)\)/);

      if (!currentQuestionCounts[currentSession]) {
        currentQuestionCounts[currentSession] = 0;
      }

      var numQ = 10;
      if (countMatch) {
        if (countMatch[2]) {
          numQ = parseInt(countMatch[2], 10) - parseInt(countMatch[1], 10) + 1;
        } else {
          numQ = parseInt(countMatch[1], 10);
        }
      }

      var startQ = currentQuestionCounts[currentSession] + 1;
      var endQ = startQ + numQ - 1;

      currentQuestionCounts[currentSession] = Math.max(currentQuestionCounts[currentSession], endQ);
      sessionMaxQ[currentSession] = Math.max(sessionMaxQ[currentSession] || 0, endQ);
      ensureQuestionColumnsExist("S" + currentSession, endQ);

      // Auto-corregir el encabezado de sesión en la fila 1 (se escribirá en bloque al final)
      sessionLabels[c - 3] = "SESIÓN " + currentSession;

      var colorBg = (subjectIndex % 2 === 0) ? "#527FD9" : "#6C7D9F";
      subjectIndex++;

      var habText = extractHabilidades(val2);
      setupSubjectHeader(sheet, c, cleanName, '(' + startQ + '-' + endQ + ')', colorBg, habText);

      var colStartLetter = convertColToLetter(2 + startQ);
      var colEndLetter = convertColToLetter(2 + endQ);
      var sheetSource = "S" + currentSession;

      var colFormulas = [];
      for (var r = 4; r < studentEndRow; r++) {
        var sessionRow = r - 2; // Las notas en las hojas de sesión comienzan en la fila 2
        colFormulas.push(["=SUM('" + sheetSource + "'!" + colStartLetter + sessionRow + ":" + colEndLetter + sessionRow + ")"]);
      }

      var targetRange = sheet.getRange(4, c, numDataRows, 1);
      targetRange.setFormulas(colFormulas);
      targetRange.setFontColor('#000000').setFontWeight('normal');
      targetRange.setHorizontalAlignment('center').setVerticalAlignment('middle');
    }
  }

  // Escribir todas las etiquetas de sesión de la fila 1 de una sola vez
  try {
    sheet.getRange(1, 3, 1, totalCol - 3).setValues([sessionLabels]);
  } catch (e) { }

  // Borrar columnas de preguntas sobrantes (si la sesión solo llega a #20, eliminar las de más)
  Object.keys(sessionMaxQ).forEach(function (sess) {
    trimSessionQuestionColumns("S" + sess, sessionMaxQ[sess]);
  });
}

function extractHabilidades(headerText) {
  try {
    var lines = String(headerText || '').split('\n');
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (/^habilidades?:/i.test(line)) {
        return line.substring(line.indexOf(':') + 1).trim();
      }
    }
  } catch (e) { }
  return "";
}

function setupSubjectHeader(sheet, colIndex, name, rangeText, bgColor, habilidades) {
  var cleanNameOnly = name.replace(/\s*\(\d+.*?\)/g, '').trim();
  var title = rangeText ? cleanNameOnly + '\n' + rangeText : cleanNameOnly;
  var hab = String(habilidades || '').trim();
  if (hab) {
    title += '\nHabilidades: ' + hab;
  }
  var bg = bgColor || "#527FD9";
  var cell = sheet.getRange(2, colIndex, 2, 1); // Filas 2 y 3

  try {
    var mergedRanges = cell.getMergedRanges();
    for (var i = 0; i < mergedRanges.length; i++) {
      mergedRanges[i].breakApart();
    }
  } catch (e) { }

  try {
    cell.merge();
  } catch (e) { }

  cell.setValue(title)
    .setBackground(bg)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}

function realignSessionHeaders(sheet) {
  var maxCols = sheet.getLastColumn();
  var row1Full = readRowValues(sheet, 1, maxCols);
  var row2Full = readRowValues(sheet, 2, maxCols);

  var totalCol = findTotalColFromRows(row1Full, row2Full, maxCols);
  if (totalCol === -1) totalCol = maxCols - 1;
  if (totalCol <= 3) return;

  var subjectRange = sheet.getRange(1, 3, 1, totalCol - 3);

  // Preservar valores de celdas combinadas antes de desagrupar
  var mergeData = [];
  try {
    var existingMerges = subjectRange.getMergedRanges();
    for (var i = 0; i < existingMerges.length; i++) {
      var mr = existingMerges[i];
      mergeData.push({
        value: mr.getValue(),
        startCol: mr.getColumn(),
        numCols: mr.getNumColumns()
      });
    }
  } catch (e) { }

  // Desagrupar todas las combinaciones
  try {
    var allMerges = subjectRange.getMergedRanges();
    for (var i = 0; i < allMerges.length; i++) {
      allMerges[i].breakApart();
    }
  } catch (e) { }

  // Restaurar valores a cada celda individual
  mergeData.forEach(function (md) {
    for (var j = 0; j < md.numCols; j++) {
      try {
        sheet.getRange(1, md.startCol + j).setValue(md.value);
      } catch (e) { }
    }
  });

  // Auto-sanar celdas que queden vacías: tomar la sesión desde la fórmula de la columna
  var row4Formulas = [];
  try { row4Formulas = sheet.getRange(4, 1, 1, maxCols).getFormulas()[0]; } catch (e) { }
  var restoredVals = sheet.getRange(1, 3, 1, totalCol - 3).getValues()[0];
  for (var sc = 0; sc < restoredVals.length; sc++) {
    if (!String(restoredVals[sc] || '').trim()) {
      var colIdx = 3 + sc;
      var sessFromF = getSessionFromFormulaString(row4Formulas[colIdx - 1]);
      if (sessFromF) {
        try {
          sheet.getRange(1, colIdx).setValue("SESIÓN " + sessFromF);
        } catch (e) { }
      }
    }
  }

  // Leer valores individuales (ahora todos correctos) y agrupar
  var row1Values = sheet.getRange(1, 3, 1, totalCol - 3).getValues()[0];

  var startCol = 3;
  var currentSessionText = String(row1Values[0] || '').trim().toUpperCase();

  for (var i = 1; i < row1Values.length; i++) {
    var cellText = String(row1Values[i] || '').trim().toUpperCase();
    var colIndex = 3 + i;

    if (cellText !== currentSessionText) {
      mergeAndFormatSessionHeader(sheet, startCol, colIndex - 1, currentSessionText);
      startCol = colIndex;
      currentSessionText = cellText;
    }
  }
  mergeAndFormatSessionHeader(sheet, startCol, totalCol - 1, currentSessionText);
}

function mergeAndFormatSessionHeader(sheet, startCol, endCol, sessionText) {
  var range = sheet.getRange(1, startCol, 1, endCol - startCol + 1);
  if (endCol > startCol) {
    try {
      range.merge();
    } catch (e) { }
  }

  // Extraer número de sesión (ej: "SESIÓN 3" -> 3 -> Impar -> #8D7BA6)
  var sessNum = 1;
  var match = sessionText.match(/(\d+)/);
  if (match) {
    sessNum = parseInt(match[1], 10);
  }

  var label = "SESIÓN " + sessNum;
  var bgColor = (sessNum % 2 === 1) ? "#8D7BA6" : "#B23A48";

  range.setValue(label)
    .setBackground(bgColor)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

function updateCalculatedColumns(sheet, studentEndRow) {
  var maxCols = sheet.getLastColumn();
  var row1Vals = readRowValues(sheet, 1, maxCols);
  var row2Vals = readRowValues(sheet, 2, maxCols);

  var totalCol = -1;
  var avgCol = -1;
  for (var c = 0; c < maxCols; c++) {
    var v1 = String(row1Vals[c] || '').toUpperCase();
    var v2 = String(row2Vals[c] || '').toUpperCase();
    if (v1.indexOf('TOTAL') !== -1 || v2.indexOf('TOTAL') !== -1) totalCol = c + 1;
    if (v1.indexOf('PROMEDIO') !== -1 || v2.indexOf('PROMEDIO') !== -1) avgCol = c + 1;
  }

  if (totalCol === -1) return;

  var numDataRows = studentEndRow - 4; // Solo las filas de estudiantes (protegiendo las filas de totales inferiores)
  if (numDataRows < 1) return;

  var firstLetter = 'C';
  var lastLetter = convertColToLetter(totalCol - 1); // Columna justo antes de TOTAL

  var totalFormulas = [];
  var avgFormulas = [];

  for (var row = 4; row < studentEndRow; row++) { // Va hasta studentEndRow - 1
    totalFormulas.push(["=SUM(" + firstLetter + row + ":" + lastLetter + row + ")"]);
    avgFormulas.push(["=AVERAGE(" + firstLetter + row + ":" + lastLetter + row + ")"]);
  }

  if (totalCol > 0) {
    sheet.getRange(4, totalCol, numDataRows, 1).setFormulas(totalFormulas);
  }
  if (avgCol > 0) {
    sheet.getRange(4, avgCol, numDataRows, 1).setFormulas(avgFormulas).setNumberFormat("0.00");
  }

  // Re-escribir de forma dinámica las fórmulas de TOTAL y PROMEDIO por columnas en la parte inferior
  try {
    var colTotalRow = studentEndRow;
    var colAvgRow = studentEndRow + 1;

    // Limpiar desde la fila de TOTAL hacia abajo (incluyendo las propias filas 44 y 45)
    // para eliminar fórmulas antiguas que puedan estar incorrectamente posicionadas
    var lastRowOfSheet = sheet.getLastRow();
    if (lastRowOfSheet >= colTotalRow) {
      sheet.getRange(colTotalRow, 3, lastRowOfSheet - colTotalRow + 1, avgCol - 2).clearContent();
    }

    var colTotalFormulas = [];
    var colAvgFormulas = [];

    for (var col = 3; col <= avgCol; col++) {
      var colLetter = convertColToLetter(col);
      var rangeStr = colLetter + "4:" + colLetter + (studentEndRow - 1);
      colTotalFormulas.push("=SUM(" + rangeStr + ")");
      colAvgFormulas.push("=AVERAGE(" + rangeStr + ")");
    }

    sheet.getRange(colTotalRow, 3, 1, avgCol - 2).setFormulas([colTotalFormulas]);
    sheet.getRange(colAvgRow, 3, 1, avgCol - 2).setFormulas([colAvgFormulas]);
  } catch (e) { }
}

function ensureTotalAverageRows(sheet, studentEndRow) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 3) return;

  var totalRow = studentEndRow;
  var avgRow = studentEndRow + 1;

  // Asegurar suficientes filas antes de escribir las fórmulas
  var maxRows = sheet.getMaxRows();
  if (maxRows < avgRow) {
    try {
      sheet.insertRowsAfter(maxRows, avgRow - maxRows);
    } catch (e) { }
  }

  var avgCol = findAverageColIndex(sheet);
  var lastDataCol = (avgCol > 0) ? avgCol : lastCol;

  var totalFormulas = [];
  var avgFormulas = [];
  var endRow = totalRow - 1;
  for (var col = 3; col <= lastDataCol; col++) {
    var colLetter = convertColToLetter(col);
    totalFormulas.push("=SUM(" + colLetter + "4:" + colLetter + endRow + ")");
    avgFormulas.push("=AVERAGE(" + colLetter + "4:" + colLetter + endRow + ")");
  }
  if (totalFormulas.length === 0) return;

  var rangeT = sheet.getRange(totalRow, 3, 1, totalFormulas.length);
  var rangeA = sheet.getRange(avgRow, 3, 1, avgFormulas.length);

  try {
    rangeT.setFormulas([totalFormulas]);
    rangeA.setFormulas([avgFormulas]);
  } catch (e) { }

  try {
    rangeT.setBackground("#93ABD9").setFontColor("#FFFFFF").setFontWeight("normal")
      .setFontSize(12).setFontFamily("Nunito").setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
    rangeA.setBackground("#D9E1F2").setFontColor("#000000").setFontWeight("normal")
      .setFontSize(12).setFontFamily("Nunito").setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00");

    var row1Vals = readRowValues(sheet, 1, lastCol);
    var row2Vals = readRowValues(sheet, 2, lastCol);
    var totalCol = findTotalColFromRows(row1Vals, row2Vals, lastCol);
    if (totalCol > 0) {
      sheet.getRange(avgRow, totalCol)
        .setBackground("#93ABD9")
        .setFontColor("#FFFFFF")
        .setFontWeight("normal")
        .setFontSize(12)
        .setFontFamily("Nunito")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setNumberFormat("0.00");
    }
  } catch (e) { }
}

function applyTableFormatting(sheet, studentEndRow) {
  var lastCol = sheet.getLastColumn();
  var numRows = (studentEndRow + 1) - 3; // Incluye las filas de TOTAL y PROMEDIO
  if (numRows < 1) return;

  // Asegurar que la hoja tenga suficientes filas antes de formatear para evitar errores Out of Bounds
  var maxRows = sheet.getMaxRows();
  var neededRows = studentEndRow + 1;
  if (maxRows < neededRows) {
    try {
      sheet.insertRowsAfter(maxRows, neededRows - maxRows);
      maxRows = sheet.getMaxRows();
    } catch (e) { }
  }

  // Asignar color de pestaña azul oscuro #223459 para Consolidado
  try {
    sheet.setTabColor("#223459");
  } catch (e) { }

  // Fijar ancho fijo de 175 para las columnas A y B
  try {
    sheet.setColumnWidth(1, 175);
    sheet.setColumnWidth(2, 175);
  } catch (e) { }

  // Detectar columna TOTAL dinámicamente (lectura en bloque)
  var row1Vals = readRowValues(sheet, 1, lastCol);
  var row2Vals = readRowValues(sheet, 2, lastCol);
  var totalCol = -1;
  var avgCol = -1;
  for (var c = 0; c < lastCol; c++) {
    var v1 = String(row1Vals[c] || '').toUpperCase();
    var v2 = String(row2Vals[c] || '').toUpperCase();
    if (v1.indexOf('TOTAL') !== -1 || v2.indexOf('TOTAL') !== -1) totalCol = c + 1;
    if (v1.indexOf('PROMEDIO') !== -1 || v2.indexOf('PROMEDIO') !== -1) avgCol = c + 1;
  }

  // Bordes generales de datos (desde fila 4 hasta PROMEDIO)
  var dataRange = sheet.getRange(4, 1, numRows, lastCol);
  dataRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Copiar el formato (fuente, número, tamaño, alineación) de la fila 4 a todas las demás filas de estudiantes
  if (studentEndRow > 5) {
    try {
      var sourceRange = sheet.getRange(4, 1, 1, lastCol);
      var targetRange = sheet.getRange(5, 1, studentEndRow - 4, lastCol);
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    } catch (e) { }
  }

  // Toda la zona de datos (estudiantes + TOTAL + PROMEDIO) en fuente Nunito y tamaño 12
  try {
    sheet.getRange(4, 1, numRows, lastCol).setFontSize(12).setFontFamily("Nunito");
  } catch (e) { }

  // Fuente Nunito en las cabeceras (filas 1 a 3)
  try {
    sheet.getRange(1, 1, 3, lastCol).setFontFamily("Nunito");
  } catch (e) { }

  // Bordes negros a las cabeceras (filas 1, 2 y 3)
  var headerRange = sheet.getRange(1, 1, 3, lastCol);
  headerRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Alternar filas de estudiantes en toda la fila (A hasta última columna) — en un solo setBackgrounds
  var numStudentRows = studentEndRow - 4;
  if (numStudentRows > 0 && lastCol > 0) {
    var bgMatrix = [];
    var nameBgs = [];
    for (var row = 4; row < studentEndRow; row++) {
      var bg = (row % 2 === 0) ? "#FFFFFF" : "#F3F3F3";
      var rowArr = [];
      for (var col = 1; col <= lastCol; col++) {
        rowArr.push((col === totalCol) ? "#93ABD9" : bg);
      }
      bgMatrix.push(rowArr);
      nameBgs.push([bg, bg]);
    }
    sheet.getRange(4, 1, numStudentRows, lastCol).setBackgrounds(bgMatrix);
    sheet.getRange(4, 1, numStudentRows, 2).setBackgrounds(nameBgs);
  }

  // Columna TOTAL vertical: aplicar fondo #93ABD9 y letra blanca para filas de estudiantes
  if (totalCol > 0 && numStudentRows > 0) {
    sheet.getRange(4, totalCol, numStudentRows, 1)
      .setBackground("#93ABD9")
      .setFontColor("#FFFFFF") // Todo en blanco
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }

  // Asegurar fórmulas de nombres de estudiantes (A y B) desde S1 para evitar #REF! sobrantes
  try {
    var nameFormulas = [];
    for (var r = 4; r < studentEndRow; r++) {
      var s1Row = r - 2; // Fila 4 de Consolidado -> Fila 2 de S1
      nameFormulas.push([
        "=IF('S1'!A" + s1Row + "=\"\",\"\",'S1'!A" + s1Row + ")",
        "=IF('S1'!B" + s1Row + "=\"\",\"\",'S1'!B" + s1Row + ")"
      ]);
    }
    if (nameFormulas.length > 0) {
      sheet.getRange(4, 1, nameFormulas.length, 2).setFormulas(nameFormulas);
    }
  } catch (e) { }

  // Filas de TOTAL y PROMEDIO al final (posición dinámica según studentEndRow)
  try {
    // Fila de TOTAL: Título "TOTAL" en A:B (combinadas) en negrita, fondo #93ABD9 y letra blanca
    var totalTitleRange = sheet.getRange(studentEndRow, 1, 1, 2);
    try { totalTitleRange.breakApart(); } catch (e) { }
    try { totalTitleRange.merge(); } catch (e) { }
    totalTitleRange.setValue("TOTAL")
      .setBackground("#93ABD9")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    sheet.getRange(studentEndRow, 3, 1, lastCol - 2)
      .setBackground("#93ABD9")
      .setFontColor("#FFFFFF")
      .setFontWeight("normal")
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setNumberFormat("0");

    // Fila de PROMEDIO: Título "PROMEDIO" en A:B (combinadas) en negrita, fondo azul oscuro #223459 y letra blanca
    var avgTitleRange = sheet.getRange(studentEndRow + 1, 1, 1, 2);
    try { avgTitleRange.breakApart(); } catch (e) { }
    try { avgTitleRange.merge(); } catch (e) { }
    avgTitleRange.setValue("PROMEDIO")
      .setBackground("#223459")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Fila de PROMEDIO: Valores con fondo base #D9E1F2 y letra negra, sin negrita
    sheet.getRange(studentEndRow + 1, 3, 1, lastCol - 2)
      .setBackground("#D9E1F2")
      .setFontColor("#000000")
      .setFontWeight("normal")
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setNumberFormat("0.00");
  } catch (e) { }

  // Aplicar formato condicional de notas de forma automática
  applyConditionalFormattingRules(sheet, studentEndRow);

  // Garantizar que la columna TOTAL completa (desde la fila 4 hasta la fila PROMEDIO inclusive) tenga fondo #93ABD9 y letra blanca
  if (totalCol > 0) {
    sheet.getRange(4, totalCol, (studentEndRow + 1) - 3, 1)
      .setBackground("#93ABD9")
      .setFontColor("#FFFFFF")
      .setFontWeight("normal")
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }

  // Columna PROMEDIO vertical (estudiantes): tamaño de fuente 12
  if (avgCol > 0) {
    sheet.getRange(4, avgCol, studentEndRow - 4, 1)
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }

  // Failsafe: Limpiar cualquier residuo de formato o contenido debajo de la fila de PROMEDIO (fila studentEndRow + 1)
  try {
    var maxRowsOfSheet = sheet.getMaxRows();
    var startCleanupRow = studentEndRow + 2;
    if (maxRowsOfSheet >= startCleanupRow && lastCol > 0) {
      sheet.getRange(startCleanupRow, 1, maxRowsOfSheet - startCleanupRow + 1, lastCol).clearContent().clearFormat();
    }
  } catch (e) { }
}

function applyConditionalFormattingRules(sheet, studentEndRow) {
  var maxCols = sheet.getLastColumn();
  var row1Vals = readRowValues(sheet, 1, maxCols);
  var row2Vals = readRowValues(sheet, 2, maxCols);

  var totalCol = -1;
  var avgCol = -1;

  // Buscar columnas TOTAL y PROMEDIO (lectura en bloque)
  for (var c = 0; c < maxCols; c++) {
    var v1 = String(row1Vals[c] || '').toUpperCase();
    var v2 = String(row2Vals[c] || '').toUpperCase();
    if (v1.indexOf('TOTAL') !== -1 || v2.indexOf('TOTAL') !== -1) totalCol = c + 1;
    if (v1.indexOf('PROMEDIO') !== -1 || v2.indexOf('PROMEDIO') !== -1) avgCol = c + 1;
  }

  var ranges = [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return;
  var numRows = studentEndRow - 4; // Excluye las filas inferiores de totales y promedios de la regla de color

  // Columnas de notas/materias (desde la columna 3 hasta la columna anterior a TOTAL)
  if (totalCol > 3) {
    ranges.push(sheet.getRange(4, 3, numRows, totalCol - 3));
  }
  // Columna PROMEDIO (estudiantes)
  if (avgCol > 0) {
    ranges.push(sheet.getRange(4, avgCol, numRows, 1));
  }
  // Fila PROMEDIO horizontal inferior (Fila studentEndRow + 1, excluyendo la columna TOTAL)
  if (totalCol > 3) {
    ranges.push(sheet.getRange(studentEndRow + 1, 3, 1, totalCol - 3));
  }
  if (avgCol > 0) {
    ranges.push(sheet.getRange(studentEndRow + 1, avgCol, 1, 1));
  }

  if (ranges.length === 0) return;

  // Limpiar reglas existentes para evitar duplicados
  sheet.clearConditionalFormatRules();

  var rules = [];

  // Regla 1: 0.5 a 3.49 -> Rojo (#D96B6B)
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.5, 3.49)
    .setBackground("#E06666")
    .setFontColor("#000000")
    .setRanges(ranges)
    .build();
  rules.push(rule1);

  // Regla 2: 3.5 a 3.99 -> Naranja (#E69138)
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(3.5, 3.99)
    .setBackground("#E69238")
    .setFontColor("#000000")
    .setRanges(ranges)
    .build();
  rules.push(rule2);

  // Regla 3: 4 a 4.49 -> Amarillo (#F1C232)
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(4, 4.49)
    .setBackground("#FFD966")
    .setFontColor("#000000")
    .setRanges(ranges)
    .build();
  rules.push(rule3);

  // Regla 4: 4.5 a 5 -> Verde (#93C47D)
  var rule4 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(4.5, 5)
    .setBackground("#93C47D")
    .setFontColor("#000000")
    .setRanges(ranges)
    .build();
  rules.push(rule4);

  sheet.setConditionalFormatRules(rules);
}

function getValidStudentCount() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetS1 = ss.getSheetByName("S1") || ss.getSheetByName("s1");
  if (!sheetS1) return 0;
  var maxS1Rows = sheetS1.getLastRow();
  var numStudents = 0;
  if (maxS1Rows >= 2) {
    var vals = sheetS1.getRange(2, 1, maxS1Rows - 1, 2).getValues();
    for (var r = 0; r < vals.length; r++) {
      var last = String(vals[r][0] || '').trim();
      var first = String(vals[r][1] || '').trim();
      if (last !== '' || first !== '' && !last.includes('#REF!') && !first.includes('#REF!')) {
        numStudents = r + 1;
      }
    }
  }
  return numStudents;
}

function applySessionSheetFormatting(sheet) {
  var name = sheet.getName();
  var numStudents = getValidStudentCount();
  var targetStudentRows = (numStudents > 0) ? numStudents : 40;
  var targetLastRow = 1 + targetStudentRows;

  // Asegurar que la hoja tenga suficientes filas
  var maxRows = sheet.getMaxRows();
  if (maxRows < targetLastRow) {
    try {
      sheet.insertRowsAfter(maxRows, targetLastRow - maxRows);
      maxRows = sheet.getMaxRows();
    } catch (e) { }
  }

  var lastCol = Math.max(sheet.getLastColumn(), 2);

  // Asignar color de pestaña intercalado (Azul para S1/S3/etc. y Rojo para S2/S4/etc.)
  var sessionNum = parseInt(name.replace(/\D/g, ""), 10);
  if (!isNaN(sessionNum)) {
    var tabColor = (sessionNum % 2 === 1) ? "#8D7BA6" : "#B23A48";
    try {
      sheet.setTabColor(tabColor);
    } catch (e) { }
  }

  // Limpiar bordes viejos
  try {
    sheet.getRange(1, 1, maxRows, lastCol).setBorder(false, false, false, false, false, false);
  } catch (e) { }

  // Aplicar bordes negros en el rango del diseño (de la cabecera a targetLastRow)
  try {
    sheet.getRange(1, 1, targetLastRow, lastCol).setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  } catch (e) { }

  // Fuente Nunito en toda la hoja de sesión
  try {
    sheet.getRange(1, 1, targetLastRow, lastCol).setFontFamily("Nunito");
  } catch (e) { }

  // Centrar y dar formato a los encabezados de preguntas (fondo #35528C, letra blanca, negrita)
  try {
    if (lastCol >= 3) {
      sheet.getRange(1, 3, 1, lastCol - 2)
        .setBackground("#35528C")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      sheet.getRange(2, 3, targetLastRow - 1, lastCol - 2).setHorizontalAlignment("center");
    }
  } catch (e) { }

  // Replicar el formato (fuente, fondo, alineación, bordes) y anchos de columnas A y B desde S1
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetS1 = ss.getSheetByName("S1") || ss.getSheetByName("s1");
  if (sheetS1 && name.toUpperCase() !== "S1") {
    try {
      var copyLimit = Math.min(targetLastRow, sheetS1.getLastRow());
      if (copyLimit >= 1) {
        var sourceRange = sheetS1.getRange(1, 1, copyLimit, 2);
        var targetRange = sheet.getRange(1, 1, copyLimit, 2);
        sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      }
      sheet.setColumnWidth(1, 175);
      sheet.setColumnWidth(2, 175);
    } catch (e) { }
  } else if (name.toUpperCase() === "S1") {
    // Si es S1, aplicar el estilo correcto a los encabezados de A1 y B1
    try {
      sheet.getRange(1, 1, 1, 2)
        .setBackground("#35528C")
        .setFontColor("#FFFFFF")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
      sheet.setColumnWidth(1, 175);
      sheet.setColumnWidth(2, 175);
    } catch (e) { }
  }
}

function gestionarSesionesNativo() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // S1 siempre es la sesión maestra — no se pide al usuario
  var sheetS1 = ss.getSheetByName("S1") || ss.getSheetByName("s1");
  if (!sheetS1) {
    ui.alert("Error", "No se encontró la hoja 'S1'. La primera sesión siempre debe llamarse S1.", ui.ButtonSet.OK);
    return;
  }

  // Obtener estudiantes de S1
  var lastRow1 = sheetS1.getLastRow();
  var studentsS1 = [];
  if (lastRow1 >= 2) {
    var vals1 = sheetS1.getRange(2, 1, lastRow1 - 1, 2).getValues();
    vals1.forEach(function (row) {
      var lastName = String(row[0] || '').trim();
      var firstName = String(row[1] || '').trim();
      if (lastName || firstName) {
        studentsS1.push({
          last: lastName,
          first: firstName,
          fullName: (lastName + " " + firstName).trim(),
          key: fullNameKey(lastName, firstName)
        });
      }
    });
  }

  if (studentsS1.length === 0) {
    ui.alert("Sin datos", "S1 no tiene estudiantes registrados.", ui.ButtonSet.OK);
    return;
  }

  // Detectar todas las hojas de sesión distintas de S1 y Consolidado
  var excluded = ["S1", "SHEET1", "CONSOLIDADO"];
  var sessionSheets = [];
  ss.getSheets().forEach(function (s) {
    if (excluded.indexOf(s.getName().toUpperCase()) === -1) {
      sessionSheets.push(s);
    }
  });

  if (sessionSheets.length === 0) {
    ui.alert("Sin sesiones", "No hay otras hojas de sesión para validar. Solo existe S1.", ui.ButtonSet.OK);
    return;
  }

  var issuesFound = false;
  var report = "";

  sessionSheets.forEach(function (sheetSx) {
    var sxName = sheetSx.getName();
    var lastRowSx = sheetSx.getLastRow();
    var sxColCount = sheetSx.getLastColumn();

    var sxData = [];
    if (lastRowSx >= 2 && sxColCount >= 1) {
      sxData = sheetSx.getRange(2, 1, lastRowSx - 1, Math.max(sxColCount, 2)).getValues();
    }

    // Mapa: nombre completo normalizado -> notas (col 3+)
    var sxMap = {};
    sxData.forEach(function (row) {
      var last = String(row[0] || '').trim();
      var first = String(row[1] || '').trim();
      if (last || first) {
        sxMap[fullNameKey(last, first)] = row.slice(2);
      }
    });

    var s1Keys = {};
    studentsS1.forEach(function (s) { s1Keys[s.key] = true; });

    var missingInSx = [];
    var sinNotas = [];

    studentsS1.forEach(function (s) {
      if (!sxMap[s.key]) {
        missingInSx.push(s.fullName);
      } else if (sxColCount >= 3) {
        var grades = sxMap[s.key];
        var anyFilled = false;
        for (var g = 0; g < grades.length; g++) {
          var cell = grades[g];
          // Se considera "con nota" si la celda NO está vacía (un 0 también es nota escaneada)
          if (cell !== null && cell !== undefined && String(cell).trim() !== '') {
            anyFilled = true;
            break;
          }
        }
        if (!anyFilled) sinNotas.push(s.fullName);
      }
    });

    var extraInSx = Object.keys(sxMap).filter(function (k) { return !s1Keys[k]; });

    if (missingInSx.length > 0 || extraInSx.length > 0 || sinNotas.length > 0) {
      issuesFound = true;
      report += "📋 " + sxName + ":\n";
      if (missingInSx.length > 0) report += "  Faltan nombres: " + missingInSx.join(", ") + "\n";
      if (extraInSx.length > 0) report += "  Sobran nombres: " + extraInSx.join(", ") + "\n";
      if (sinNotas.length > 0) report += "  Sin notas (celdas vacías): " + sinNotas.join(", ") + "\n";
    }
  });

  if (!issuesFound) {
    ui.alert("✅ Sin Inconsistencias", "Todos los estudiantes coinciden con la lista maestra y tienen notas registradas.", ui.ButtonSet.OK);
    return;
  }

  ui.alert("⚠️ Inconsistencias Detectadas", report + "\nEste reporte es solo informativo. Revisa y corrige la data manualmente.", ui.ButtonSet.OK);
}

function convertColToLetter(colIndex) {
  var letter = '';
  var temp = colIndex;
  while (temp > 0) {
    var rem = (temp - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    temp = Math.floor((temp - rem) / 26);
  }
  return letter;
}

function checkSymmetryS1S2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheetS1 = ss.getSheetByName("S1") || ss.getSheetByName("s1");
  var sheetS2 = ss.getSheetByName("S2") || ss.getSheetByName("s2");
  if (!sheetS1 || !sheetS2) return;

  var lastRow1 = sheetS1.getLastRow();
  var lastRow2 = sheetS2.getLastRow();

  var studentsS1 = [];
  if (lastRow1 >= 2) {
    var vals1 = sheetS1.getRange(2, 1, lastRow1 - 1, 2).getValues();
    vals1.forEach(function (row) {
      var lastName = String(row[0] || '').trim();
      var firstName = String(row[1] || '').trim();
      if (lastName || firstName) {
        studentsS1.push({
          last: lastName,
          first: firstName,
          fullName: (lastName + " " + firstName).trim(),
          key: normalizeNameKey(lastName) + "|" + normalizeNameKey(firstName)
        });
      }
    });
  }

  var studentsS2 = [];
  if (lastRow2 >= 2) {
    var vals2 = sheetS2.getRange(2, 1, lastRow2 - 1, 2).getValues();
    vals2.forEach(function (row) {
      var lastName = String(row[0] || '').trim();
      var firstName = String(row[1] || '').trim();
      if (lastName || firstName) {
        studentsS2.push((lastName + " " + firstName).trim());
      }
    });
  }

  var mapS1 = {};
  studentsS1.forEach(function (name) { mapS1[name.toUpperCase()] = name; });

  var mapS2 = {};
  studentsS2.forEach(function (name) { mapS2[name.toUpperCase()] = name; });

  var missingInS2 = [];
  studentsS1.forEach(function (name) {
    if (!mapS2[name.toUpperCase()]) missingInS2.push(name);
  });

  var missingInS1 = [];
  studentsS2.forEach(function (name) {
    if (!mapS1[name.toUpperCase()]) missingInS1.push(name);
  });

  if (missingInS2.length > 0 || missingInS1.length > 0) {
    var alertMsg = "⚠️ Advertencia de Inconsistencia entre Sesiones:\n\n";
    if (missingInS2.length > 0) {
      alertMsg += "Estudiantes en S1 que FALTAN en S2:\n- " + missingInS2.join("\n- ") + "\n\n";
    }
    if (missingInS1.length > 0) {
      alertMsg += "Estudiantes en S2 que FALTAN en S1:\n- " + missingInS1.join("\n- ") + "\n\n";
    }
    ui.alert("⚠️ Inconsistencia de Estudiantes Detectada", alertMsg, ui.ButtonSet.OK);
  }
}

function openHtmlSessionManagerModal() {
  var ui = SpreadsheetApp.getUi();
  var sessionsData = getSessionSheetsData();

  // Calcular el nombre sugerido para la nueva sesión (ej. si existen S1 y S2, sugerir S3)
  var maxSessionNum = 0;
  sessionsData.forEach(function (s) {
    var num = parseInt(s.name.replace(/\D/g, ''), 10);
    if (!isNaN(num) && num > maxSessionNum) maxSessionNum = num;
  });
  var nextSessionName = 'S' + (maxSessionNum + 1);

  // Construir la lista de sesiones completamente en el servidor (HTML estático)
  var sessionListHtml = "";
  if (sessionsData.length === 0) {
    sessionListHtml = "<div style='font-size:11px;color:#888;text-align:center;padding:20px;'><i class='fa-solid fa-circle-info'></i> No hay hojas de sesión.</div>";
  } else {
    sessionsData.forEach(function (item) {
      sessionListHtml += "<div class='session-item'>";
      sessionListHtml += "  <div class='session-name'><i class='fa-solid fa-folder' style='color:#8D7BA6; margin-right: 6px;'></i> " + item.name + "</div>";
      sessionListHtml += "  <div class='actions-container'>";
      if (!item.isS1) {
        sessionListHtml += "    <button class='btn-danger' onclick='deleteSession(this, \"" + item.name + "\")'><i class='fa-solid fa-trash-can'></i> Eliminar</button>";
      } else {
        sessionListHtml += "    <span style='font-size:10px;color:#64748b;font-style:italic;'>Sesión Maestra</span>";
      }
      sessionListHtml += "  </div>";
      sessionListHtml += "</div>";
    });
  }

  var htmlContent =
    "<!DOCTYPE html>" +
    "<html><head><base target='_top'>" +
    "<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css'>" +
    "<style>" +
    "* { box-sizing: border-box; }" +
    "body { font-family: Segoe UI, sans-serif; padding: 12px; color: #2c3e50; margin: 0; background: #fff; }" +
    ".warning-box { background: #fff9e6; border-left: 4px solid #ffb020; padding: 8px; margin-bottom: 10px; font-size: 11px; border-radius: 4px; color: #7a5200; font-weight: 500; }" +
    ".section-title { font-size: 13px; font-weight: bold; margin-bottom: 8px; color: #1a237e; border-bottom: 1px solid #e2e8f0; padding-bottom: 4px; }" +
    ".form-group { margin-bottom: 12px; display: flex; gap: 8px; }" +
    "input { flex: 1; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 12px; }" +
    "button { padding: 8px 14px; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; font-size: 12px; }" +
    ".btn-primary { background: #3b5998; color: white; }" +
    ".btn-danger { background: #ef4444; color: white; padding: 4px 8px; font-size: 11px; border-radius: 3px; }" +
    ".btn-secondary { background: #10b981; color: white; padding: 4px 8px; font-size: 11px; border-radius: 3px; }" +
    ".session-list { margin-top: 10px; max-height: 180px; overflow-y: auto; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px; }" +
    ".session-item { display: flex; justify-content: space-between; align-items: center; padding: 6px; border-bottom: 1px solid #f1f5f9; }" +
    ".session-item:last-child { border-bottom: none; }" +
    ".session-name { font-size: 12px; font-weight: 600; }" +
    ".actions-container { display: flex; gap: 6px; }" +
    "#errorBox { color: #d32f2f; font-size: 11px; margin-top: 8px; display: none; background: #ffebee; padding: 6px; border-radius: 4px; }" +
    "#successBox { color: #155724; font-size: 11px; margin-top: 8px; display: none; background: #d4edda; padding: 6px; border-radius: 4px; }" +
    ".footer { text-align: right; margin-top: 15px; border-top: 1px solid #e2e8f0; padding-top: 10px; }" +
    ".btn-close { background: #e0e0e0; color: #333; }" +
    "</style></head><body>" +
    "<div class='warning-box'><i class='fa-solid fa-triangle-exclamation'></i> <b>Nota:</b> Si tienes más de una cuenta de Google abierta, puede haber problemas. Cierra sesiones extras si algo falla.</div>" +
    "<div class='section-title'><i class='fa-solid fa-circle-plus'></i> Crear Nueva Sesión</div>" +
    "<div class='form-group'>" +
    "  <input type='text' id='sessionNameInput' value='" + nextSessionName + "' readonly style='background-color: #f1f5f9; color: #64748b; cursor: not-allowed;' />" +
    "  <button class='btn-primary' id='btnCrear' onclick='createSession()'><i class='fa-solid fa-plus-circle'></i> Crear</button>" +
    "</div>" +
    "<div class='section-title'><i class='fa-solid fa-folder-open'></i> Sesiones Existentes</div>" +
    "<div class='session-list' id='sessionListContainer'>" +
    sessionListHtml +
    "</div>" +
    "<div id='errorBox'></div>" +
    "<div id='successBox'></div>" +
    "<div class='footer'><button class='btn-close' onclick='google.script.host.close()'>Cerrar</button></div>" +
    "<script>" +
    "function showMsg(id, msg) {" +
    "  var b = document.getElementById(id);" +
    "  b.innerText = msg; b.style.display = 'block';" +
    "  setTimeout(function(){ b.style.display = 'none'; }, 3500);" +
    "}" +
    "function createSession() {" +
    "  var btn = document.getElementById('btnCrear');" +
    "  var inp = document.getElementById('sessionNameInput');" +
    "  var name = (inp.value || '').trim().toUpperCase();" +
    "  if (!name) name = '" + nextSessionName + "';" +
    "  btn.disabled = true;" +
    "  btn.innerHTML = '<i class=\"fa-solid fa-spinner fa-spin\"></i> Creando...';" +
    "  google.script.run" +
    "    .withSuccessHandler(function(res) { " +
    "      showMsg('successBox', res); " +
    "      setTimeout(function(){ google.script.host.close(); }, 1500); " +
    "    })" +
    "    .withFailureHandler(function(e) { " +
    "      showMsg('errorBox', e.message); " +
    "      btn.disabled = false;" +
    "      btn.innerHTML = '<i class=\"fa-solid fa-plus-circle\"></i> Crear';" +
    "    })" +
    "    .createSessionFromHtml(name);" +
    "}" +
    "function deleteSession(btn, name) {" +
    "  if (!confirm('¿Eliminar sesión ' + name + '?')) return;" +
    "  btn.disabled = true;" +
    "  btn.innerHTML = '<i class=\"fa-solid fa-spinner fa-spin\"></i> Elimi...';" +
    "  google.script.run" +
    "    .withSuccessHandler(function(res) { " +
    "      showMsg('successBox', res); " +
    "      setTimeout(function(){ google.script.host.close(); }, 1500); " +
    "    })" +
    "    .withFailureHandler(function(e) { " +
    "      showMsg('errorBox', e.message); " +
    "      btn.disabled = false;" +
    "      btn.innerHTML = '<i class=\"fa-solid fa-trash-can\"></i> Eliminar';" +
    "    })" +
    "    .deleteSessionFromHtml(name);" +
    "}" +
    "function syncSession(btn, name) {" +
    "  btn.disabled = true;" +
    "  btn.innerHTML = '<i class=\"fa-solid fa-spinner fa-spin\"></i> Ali...';" +
    "  google.script.run" +
    "    .withSuccessHandler(function(res) { " +
    "      showMsg('successBox', res); " +
    "      btn.disabled = false;" +
    "      btn.innerHTML = '<i class=\"fa-solid fa-sync\"></i> Alinear';" +
    "    })" +
    "    .withFailureHandler(function(e) { " +
    "      showMsg('errorBox', e.message); " +
    "      btn.disabled = false;" +
    "      btn.innerHTML = '<i class=\"fa-solid fa-sync\"></i> Alinear';" +
    "    })" +
    "    .syncSessionFromHtml(name);" +
    "}" +
    "</script>" +
    "</body></html>";

  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(460)
    .setHeight(450)
    .setTitle("Gestión de Sesiones");
  ui.showModalDialog(htmlOutput, "Gestión de Sesiones");
}

function getSessionSheetsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var excluded = ["SHEET1", "CONSOLIDADO"];
  var list = [];

  sheets.forEach(function (s) {
    var name = s.getName();
    if (excluded.indexOf(name.toUpperCase()) === -1) {
      list.push({
        name: name,
        isS1: name.toUpperCase() === "S1"
      });
    }
  });
  return list;
}

function createSessionFromHtml(newSessName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var existingSheet = ss.getSheetByName(newSessName);
  if (existingSheet) {
    throw new Error("La hoja '" + newSessName + "' ya existe.");
  }

  // Determinar posición de inserción: justo antes del Consolidado
  var allSheets = ss.getSheets();
  var consolidadoIndex = -1;
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().toUpperCase() === "CONSOLIDADO") {
      consolidadoIndex = i;
      break;
    }
  }

  // Insertar la nueva hoja antes del Consolidado
  var insertPosition = (consolidadoIndex >= 0) ? consolidadoIndex : allSheets.length;
  var newSheet = ss.insertSheet(newSessName, insertPosition);

  // Cabecera estática (etiquetas de columnas A y B)
  newSheet.getRange(1, 1).setValue("APELLIDOS");
  newSheet.getRange(1, 2).setValue("NOMBRE(S)");

  // Los NOMBRES de estudiantes NO se copian como texto: se referencian con FÓRMULAS
  // hacia S1 (lista maestra) para que la sesión permanezca SIEMPRE dinámica.
  var sheetS1 = ss.getSheetByName("S1") || ss.getSheetByName("s1");
  if (sheetS1) {
    var lastRow1 = sheetS1.getLastRow();
    var numStudents = getValidStudentCount();
    var targetStudentRows = (numStudents > 0) ? numStudents : 40;
    var lastLimit = Math.max(lastRow1, targetStudentRows + 1);
    var formulaCells = [];
    for (var r = 2; r <= lastLimit; r++) {
      formulaCells.push([
        "=IF('S1'!A" + r + "=\"\",\"\",'S1'!A" + r + ")",
        "=IF('S1'!B" + r + "=\"\",\"\",'S1'!B" + r + ")"
      ]);
    }
    if (formulaCells.length > 0) {
      newSheet.getRange(2, 1, formulaCells.length, 2).setFormulas(formulaCells);
    }
  }

  // Aplicar el mismo estilo visual que S1
  applySessionSheetFormatting(newSheet);

  return "Sesión '" + newSessName + "' creada correctamente antes del Consolidado.";
}

function deleteSessionFromHtml(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetToDelete = ss.getSheetByName(name);
  if (!sheetToDelete) {
    throw new Error("No se encontró la hoja '" + name + "'.");
  }
  ss.deleteSheet(sheetToDelete);
  return "Sesión '" + name + "' eliminada con éxito.";
}

function sessionUsesFormulaNames(sheet) {
  try {
    var formula = sheet.getRange(2, 1).getFormula();
    return formula && formula.indexOf("=") === 0;
  } catch (e) { }
  return false;
}

function syncSessionFromHtml(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetS1 = ss.getSheetByName("S1") || ss.getSheetByName("s1");
  var sheetTarget = ss.getSheetByName(name);

  if (!sheetS1 || !sheetTarget) {
    throw new Error("No se encontró la hoja maestra S1 o la hoja de destino.");
  }

  var lastRow1 = sheetS1.getLastRow();
  var lastRowTarget = sheetTarget.getLastRow();

  // Obtener estudiantes de S1
  var studentsS1 = [];
  if (lastRow1 >= 2) {
    var vals1 = sheetS1.getRange(2, 1, lastRow1 - 1, 2).getValues();
    vals1.forEach(function (row) {
      var lastName = String(row[0] || '').trim();
      var firstName = String(row[1] || '').trim();
      if (lastName || firstName) {
        studentsS1.push({
          last: lastName,
          first: firstName,
          fullName: (lastName + " " + firstName).trim(),
          key: fullNameKey(lastName, firstName)
        });
      }
    });
  }

  // Guardar las notas actuales del destino en un mapa
  var targetGradesMap = {};
  var targetColCount = sheetTarget.getLastColumn();
  var gradeColsCount = Math.max(targetColCount - 2, 0);

  if (gradeColsCount > 0 && lastRowTarget >= 2) {
    var allTargetData = sheetTarget.getRange(2, 1, lastRowTarget - 1, targetColCount).getValues();
    allTargetData.forEach(function (row) {
      var key = fullNameKey(row[0], row[1]);
      targetGradesMap[key] = row.slice(2);
    });
  }

  var newTargetData = [];
  studentsS1.forEach(function (s) {
    var key = s.key;
    var grades = targetGradesMap[key] || Array(Math.max(gradeColsCount, 1)).fill(0);
    newTargetData.push([s.last, s.first].concat(grades));
  });

  if (newTargetData.length > 0) {
    var maxRowsTarget = sheetTarget.getMaxRows();
    var maxColsTarget = sheetTarget.getLastColumn();

    // Limpiar el contenido viejo
    if (maxRowsTarget > 1) {
      sheetTarget.getRange(2, 1, maxRowsTarget - 1, maxColsTarget).clearContent();
    }

    // SIEMPRE nombres dinámicos: escribir SOLO las notas (col 3+) y reescribir
    // las fórmulas de A/B hacia S1 (por si la lista maestra creció)
    var gradeRows = newTargetData.map(function (row) { return row.slice(2); });
    sheetTarget.getRange(2, 3, gradeRows.length, gradeRows[0].length).setValues(gradeRows);

    var formulaCells = [];
    for (var r = 2; r <= newTargetData.length + 1; r++) {
      formulaCells.push([
        "=IF('S1'!A" + r + "=\"\",\"\",'S1'!A" + r + ")",
        "=IF('S1'!B" + r + "=\"\",\"\",'S1'!B" + r + ")"
      ]);
    }
    sheetTarget.getRange(2, 1, formulaCells.length, 2).setFormulas(formulaCells);

    // Ajustar el número de filas
    var requiredRows = newTargetData.length + 1;
    if (maxRowsTarget > requiredRows) {
      sheetTarget.deleteRows(requiredRows + 1, maxRowsTarget - requiredRows);
    }

    applySessionSheetFormatting(sheetTarget);
    rebuildFormulasAndFormatting();
  }

  return "Sincronización completada con éxito para la sesión '" + name + "'.";

}

function highlightBestStudentOnly() {
  var sheet = getConsolidatedSheet();
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Error", "No se encontró la hoja Consolidado.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  var studentEndRow = findConsolidadoTotalRow(sheet);
  if (studentEndRow === -1) {
    studentEndRow = getStudentRowsEndIndex(sheet);
  }

  var avgCol = findAverageColIndex(sheet);
  if (avgCol === -1 || studentEndRow <= 4) {
    SpreadsheetApp.getUi().alert("Error", "No se encontró la columna de PROMEDIO o no hay estudiantes registrados.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var numStudents = studentEndRow - 4;
  var avgValues = sheet.getRange(4, avgCol, numStudents, 1).getValues();
  var maxVal = -1;

  // Encontrar el promedio máximo
  for (var i = 0; i < avgValues.length; i++) {
    var val = parseFloat(avgValues[i][0]);
    if (!isNaN(val) && val > maxVal) {
      maxVal = val;
    }
  }

  // Generar matriz de cebra limpia para columnas A y B
  var nameBgs = [];
  for (var i = 0; i < numStudents; i++) {
    var r = 4 + i;
    var bg = (r % 2 === 0) ? "#FFFFFF" : "#F3F3F3";
    nameBgs.push([bg, bg]);
  }

  // Si ya tiene estudiantes resaltados, toggle para remover el resaltado y restaurar la cebra
  if (hasBestStudentHighlight(sheet, numStudents)) {
    sheet.getRange(4, 1, numStudents, 2).setBackgrounds(nameBgs);
    SpreadsheetApp.getActiveSpreadsheet().toast('Resaltado removido. Se restauró el formato cebra.', 'Aviso', 3);
    return;
  }

  // Restablecer el fondo de las columnas A y B con la cebra base
  sheet.getRange(4, 1, numStudents, 2).setBackgrounds(nameBgs);

  // Si hay un promedio máximo válido, resaltar solo a los estudiantes con la nota máxima
  var count = 0;
  if (maxVal > 0) {
    for (var i = 0; i < avgValues.length; i++) {
      var val = parseFloat(avgValues[i][0]);
      if (!isNaN(val) && val === maxVal) {
        var studentRow = 4 + i;
        sheet.getRange(studentRow, 1, 1, 2).setBackground("#9FC5E8");
        count++;
      }
    }
    var msg = (count === 1)
      ? '🏆 Se ha resaltado al estudiante con el mejor promedio (' + maxVal.toFixed(2) + ').'
      : '🏆 Se han resaltado a ' + count + ' estudiantes empatados con el mejor promedio (' + maxVal.toFixed(2) + ').';
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '¡Éxito!', 4);
  } else {
    SpreadsheetApp.getUi().alert("Aviso", "No se encontraron promedios mayores a 0 para resaltar.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function toggleSortConsolidado() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getConsolidatedSheet();
  if (!sheet) {
    ss.toast("No se encontró la hoja Consolidado.", "⚠️ Error", 3);
    return;
  }

  var studentEndRow = findConsolidadoTotalRow(sheet);
  if (studentEndRow === -1) {
    studentEndRow = getStudentRowsEndIndex(sheet);
  }
  var numStudents = studentEndRow - 4;
  if (numStudents < 2) {
    ss.toast("No hay suficientes estudiantes para ordenar.", "⚠️ Aviso", 3);
    return;
  }

  var lastCol = sheet.getLastColumn();
  if (lastCol < 3) return;

  var row1Vals = readRowValues(sheet, 1, lastCol);
  var row2Vals = readRowValues(sheet, 2, lastCol);
  var totalCol = findTotalColFromRows(row1Vals, row2Vals, lastCol);
  if (totalCol === -1) {
    ss.toast("No se encontró la columna TOTAL.", "⚠️ Error", 3);
    return;
  }

  var preserveHighlight = hasBestStudentHighlight(sheet, numStudents);
  var bestAvgValue = preserveHighlight ? getBestAverageValue(sheet, studentEndRow) : null;

  var range = sheet.getRange(4, 1, numStudents, lastCol);

  if (isConsolidadoAlphabetical(sheet, numStudents)) {
    range.sort({ column: totalCol, ascending: false });
    reapplyStudentRowFormatting(sheet, studentEndRow, totalCol);
    ensureTotalAverageRows(sheet, studentEndRow);
    if (bestAvgValue !== null) reapplyBestStudentHighlight(sheet, studentEndRow, bestAvgValue);
    ss.toast("Consolidado ordenado por TOTAL (de mayor a menor).", "🔽 Orden por Total", 3);
  } else {
    range.sort({ column: 1, ascending: true });
    reapplyStudentRowFormatting(sheet, studentEndRow, totalCol);
    ensureTotalAverageRows(sheet, studentEndRow);
    if (bestAvgValue !== null) reapplyBestStudentHighlight(sheet, studentEndRow, bestAvgValue);
    ss.toast("Consolidado restaurado al orden alfabético por apellidos.", "🔤 Orden Alfabético", 3);
  }
}

function findConsolidadoTotalRow(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return -1;
  var valsA = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (var r = 3; r < lastRow; r++) {
    if (String(valsA[r][0] || '').trim().toUpperCase() === 'TOTAL') {
      return r + 1;
    }
  }
  return -1;
}

function hasBestStudentHighlight(sheet, numStudents) {
  try {
    var colors = sheet.getRange(4, 1, numStudents, 2).getBackgrounds();
    for (var i = 0; i < colors.length; i++) {
      if (String(colors[i][0]).toUpperCase() === "#9FC5E8" || String(colors[i][1]).toUpperCase() === "#9FC5E8") {
        return true;
      }
    }
  } catch (e) { }
  return false;
}

function findAverageColIndex(sheet) {
  var lastCol = sheet.getLastColumn();
  var row1Vals = readRowValues(sheet, 1, lastCol);
  var row2Vals = readRowValues(sheet, 2, lastCol);
  for (var c = 0; c < lastCol; c++) {
    var v1 = String(row1Vals[c] || '').toUpperCase();
    var v2 = String(row2Vals[c] || '').toUpperCase();
    if (v1.indexOf('PROMEDIO') !== -1 || v2.indexOf('PROMEDIO') !== -1) {
      return c + 1;
    }
  }
  return -1;
}

function getBestAverageValue(sheet, studentEndRow) {
  var avgCol = findAverageColIndex(sheet);
  if (avgCol === -1) return null;
  var numStudents = studentEndRow - 4;
  if (numStudents < 1) return null;
  var avgValues = sheet.getRange(4, avgCol, numStudents, 1).getValues();
  var maxVal = -1;
  for (var i = 0; i < avgValues.length; i++) {
    var val = parseFloat(avgValues[i][0]);
    if (!isNaN(val) && val > maxVal) maxVal = val;
  }
  return (maxVal > 0) ? maxVal : null;
}

function reapplyBestStudentHighlight(sheet, studentEndRow, bestVal) {
  var avgCol = findAverageColIndex(sheet);
  if (avgCol === -1) return;
  var numStudents = studentEndRow - 4;
  if (numStudents < 1) return;
  var avgValues = sheet.getRange(4, avgCol, numStudents, 1).getValues();
  for (var i = 0; i < avgValues.length; i++) {
    var val = parseFloat(avgValues[i][0]);
    if (!isNaN(val) && val === bestVal) {
      var studentRow = 4 + i;
      sheet.getRange(studentRow, 1, 1, 2).setBackground("#9FC5E8");
    }
  }
}

function isConsolidadoAlphabetical(sheet, numStudents) {
  var names = sheet.getRange(4, 1, numStudents, 1).getValues();
  for (var i = 1; i < names.length; i++) {
    var prev = normalizeNameKey(names[i - 1][0]);
    var curr = normalizeNameKey(names[i][0]);
    if (prev && curr && prev > curr) return false;
  }
  return true;
}

function reapplyStudentRowFormatting(sheet, studentEndRow, totalCol) {
  var numStudentRows = studentEndRow - 4;
  var lastCol = sheet.getLastColumn();
  if (numStudentRows <= 0 || lastCol <= 0) return;

  // Restaurar formato (fuente, alineación, bordes) copiando la fila 4 a las demás
  try {
    if (numStudentRows > 1) {
      var src = sheet.getRange(4, 1, 1, lastCol);
      var dst = sheet.getRange(5, 1, numStudentRows - 1, lastCol);
      src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
  } catch (e) { }

  // Fondo intercalado (blanco/gris) manteniendo la columna TOTAL en azul
  var bgMatrix = [];
  var nameBgs = [];
  for (var row = 4; row < studentEndRow; row++) {
    var bg = (row % 2 === 0) ? "#FFFFFF" : "#F3F3F3";
    var rowArr = [];
    for (var col = 1; col <= lastCol; col++) {
      rowArr.push((col === totalCol) ? "#93ABD9" : bg);
    }
    bgMatrix.push(rowArr);
    nameBgs.push([bg, bg]);
  }
  sheet.getRange(4, 1, numStudentRows, lastCol).setBackgrounds(bgMatrix);
  sheet.getRange(4, 1, numStudentRows, 2).setBackgrounds(nameBgs);

  // Fuente Nunito en toda la zona de estudiantes
  try {
    sheet.getRange(4, 1, numStudentRows, lastCol).setFontFamily("Nunito");
  } catch (e) { }

  // Bordes negros completos para no perder el subrayado/grid al reordenar
  try {
    sheet.getRange(4, 1, numStudentRows, lastCol)
      .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  } catch (e) { }

  // Columna TOTAL: fondo azul y letra blanca en las filas de estudiantes
  if (totalCol > 0) {
    try {
      sheet.getRange(4, totalCol, numStudentRows, 1)
        .setBackground("#93ABD9")
        .setFontColor("#FFFFFF")
        .setFontSize(12)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    } catch (e) { }
  }

  // Columna PROMEDIO vertical (estudiantes): tamaño de fuente 12
  var avgCol = findAverageColIndex(sheet);
  if (avgCol > 0) {
    try {
      sheet.getRange(4, avgCol, numStudentRows, 1)
        .setFontSize(12)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    } catch (e) { }
  }

  // Fila TOTAL horizontal: mantener fondo azul, letra blanca y tamaño 12
  try {
    sheet.getRange(studentEndRow, 3, 1, lastCol - 2)
      .setBackground("#93ABD9")
      .setFontColor("#FFFFFF")
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setNumberFormat("0");
  } catch (e) { }

  // Fila PROMEDIO horizontal: mantener fondo #D9E1F2, letra negra y tamaño 12, con columna TOTAL en #93ABD9 letra blanca
  try {
    sheet.getRange(studentEndRow + 1, 3, 1, lastCol - 2)
      .setBackground("#D9E1F2")
      .setFontColor("#000000")
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setNumberFormat("0.00");

    if (totalCol > 0) {
      sheet.getRange(studentEndRow + 1, totalCol)
        .setBackground("#93ABD9")
        .setFontColor("#FFFFFF")
        .setFontWeight("normal")
        .setFontSize(12)
        .setFontFamily("Nunito")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setNumberFormat("0.00");
    }
  } catch (e) { }
}

// ============================================================================
// 8. UTILIDADES
// ============================================================================

function cleanEmptyRowsAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var numStudents = getValidStudentCount();

  // Si no hay nombres de estudiantes en S1, dejamos 40 filas disponibles por defecto
  var targetStudentRows = (numStudents > 0) ? numStudents : 40;

  var sheets = ss.getSheets();
  sheets.forEach(function (sheet) {
    try {
      var name = sheet.getName().toUpperCase();
      var maxRows = sheet.getMaxRows();
      var maxCols = sheet.getLastColumn();

      // Quitar filtros activos para evitar errores de permisos al eliminar/modificar filas
      try {
        if (sheet.getFilter()) {
          sheet.getFilter().remove();
        }
      } catch (e) { }

      if (name === "CONSOLIDADO") {
        // En Consolidado: 3 filas de cabecera + targetStudentRows + 1 (TOTAL) + 1 (PROMEDIO)
        var targetLastRow = 3 + targetStudentRows + 2;

        if (maxRows < targetLastRow) {
          sheet.insertRowsAfter(maxRows, targetLastRow - maxRows);
          maxRows = sheet.getMaxRows();
        }

        if (maxRows > targetLastRow) {
          if (maxCols > 0) {
            sheet.getRange(targetLastRow + 1, 1, maxRows - targetLastRow, maxCols).clearContent().clearFormat();
          }
          sheet.deleteRows(targetLastRow + 1, maxRows - targetLastRow);
        }
      } else {
        // En hojas de sesión (S1, S2, etc): 1 fila cabecera + targetStudentRows
        var targetLastRow = 1 + targetStudentRows;

        if (maxRows < targetLastRow) {
          sheet.insertRowsAfter(maxRows, targetLastRow - maxRows);
          maxRows = sheet.getMaxRows();
        }

        if (maxRows > targetLastRow) {
          if (maxCols > 0) {
            // Limpiar contenido previo (por si hay #REF! o fórmulas sobrantes) antes de borrar
            sheet.getRange(targetLastRow + 1, 1, maxRows - targetLastRow, maxCols).clearContent().clearFormat();
          }
          sheet.deleteRows(targetLastRow + 1, maxRows - targetLastRow);
        }
      }
    } catch (e) {
      // Ignorar errores puntuales en hojas protegidas
    }
  });
}

// ============================================================================
// 9. IMPORT/EXPORT JSON MODULE
// ============================================================================

function openExportJsonModal() {
  var ui = SpreadsheetApp.getUi();
  var subjectsData = getSubjectsBySession();

  // Flatten subjects to a simple JSON list
  var list = [];
  Object.keys(subjectsData).forEach(function (key) {
    if (subjectsData[key]) {
      subjectsData[key].forEach(function (s) {
        list.push({
          name: s.name,
          count: s.count,
          session: s.session
        });
      });
    }
  });

  var jsonStr = JSON.stringify(list, null, 2);

  var htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
          * { box-sizing: border-box; }
          body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 15px; color: #2c3e50; margin: 0; background: #f8fafc; }
          .header { display: flex; align-items: center; gap: 8px; margin-bottom: 12px; }
          .header i { color: #f59e0b; font-size: 20px; }
          .header h3 { margin: 0; font-size: 16px; color: #1e293b; }
          .description { font-size: 12px; color: #64748b; margin-bottom: 12px; }
          textarea { width: 100%; height: 260px; font-family: 'Courier New', Courier, monospace; font-size: 11px; padding: 10px; border: 1px solid #e2e8f0; border-radius: 6px; resize: none; background: #ffffff; color: #334155; }
          .btn-group { display: flex; justify-content: space-between; align-items: center; margin-top: 15px; }
          button { padding: 8px 14px; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; font-size: 12px; display: flex; align-items: center; gap: 6px; }
          .btn-primary { background: #3b5998; color: white; }
          .btn-primary:hover { opacity: 0.9; }
          .btn-secondary { background: #f59e0b; color: white; }
          .btn-secondary:hover { opacity: 0.9; }
          .btn-close { background: #e0e0e0; color: #333; }
          #toast { display: none; background: #10b981; color: white; padding: 8px 12px; border-radius: 4px; font-size: 11px; text-align: center; margin-top: 10px; font-weight: 500; }
        </style>
      </head>
      <body>
        <div class="header">
          <i class="fa-solid fa-file-export"></i>
          <h3>Exportar Configuración de Asignaturas</h3>
        </div>
        <div class="description">
          Copia el siguiente texto JSON o descárgalo para usarlo como respaldo o importarlo en otro libro de calificaciones.
        </div>
        <textarea id="jsonArea" readonly>${jsonStr}</textarea>
        <div id="toast"></div>
        <div class="btn-group">
          <div style="display: flex; gap: 6px;">
            <button class="btn-primary" onclick="copyToClipboard()"><i class="fa-solid fa-copy"></i> Copiar</button>
            <button class="btn-secondary" onclick="downloadJson()"><i class="fa-solid fa-download"></i> Descargar</button>
          </div>
          <button class="btn-close" onclick="google.script.host.close()">Cerrar</button>
        </div>
        <script>
          function copyToClipboard() {
            var copyText = document.getElementById("jsonArea");
            copyText.select();
            copyText.setSelectionRange(0, 99999);
            navigator.clipboard.writeText(copyText.value).then(function() {
              showToast("¡Copiado al portapapeles!");
            }, function() {
              showToast("Error al copiar. Selecciona el texto manualmente.");
            });
          }
          
          function downloadJson() {
            var dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(document.getElementById("jsonArea").value);
            var dlAnchorElem = document.createElement('a');
            dlAnchorElem.setAttribute("href",     dataStr);
            dlAnchorElem.setAttribute("download", "zipgrade_config.json");
            dlAnchorElem.click();
          }
          
          function showToast(msg) {
            var t = document.getElementById("toast");
            t.innerText = msg;
            t.style.display = "block";
            setTimeout(function() { t.style.display = "none"; }, 3000);
          }
        </script>
      </body>
    </html>
  `;

  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(430)
    .setTitle('📥 Exportar Configuración (JSON)');
  ui.showModalDialog(htmlOutput, '📥 Exportar Configuración (JSON)');
}

function openImportJsonModal() {
  var ui = SpreadsheetApp.getUi();

  var htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
        <style>
          * { box-sizing: border-box; }
          body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 15px; color: #2c3e50; margin: 0; background: #f8fafc; }
          .header { display: flex; align-items: center; gap: 8px; margin-bottom: 12px; }
          .header i { color: #10b981; font-size: 20px; }
          .header h3 { margin: 0; font-size: 16px; color: #1e293b; }
          .description { font-size: 12px; color: #64748b; margin-bottom: 12px; }
          
          .drop-zone { border: 2px dashed #cbd5e1; border-radius: 6px; padding: 15px; text-align: center; background: #ffffff; cursor: pointer; transition: border-color 0.2s; margin-bottom: 12px; }
          .drop-zone:hover { border-color: #3b5998; }
          .drop-zone i { font-size: 24px; color: #94a3b8; margin-bottom: 6px; }
          .drop-zone p { margin: 0; font-size: 11px; color: #64748b; }
          
          textarea { width: 100%; height: 160px; font-family: 'Courier New', Courier, monospace; font-size: 11px; padding: 10px; border: 1px solid #e2e8f0; border-radius: 6px; resize: none; background: #ffffff; color: #334155; margin-bottom: 10px; }
          
          .btn-group { display: flex; justify-content: space-between; align-items: center; margin-top: 15px; }
          button { padding: 8px 14px; border: none; border-radius: 4px; cursor: pointer; font-weight: 600; font-size: 12px; display: flex; align-items: center; gap: 6px; }
          .btn-primary { background: #3b5998; color: white; }
          .btn-primary:hover { opacity: 0.9; }
          .btn-primary:disabled { background: #94a3b8; cursor: not-allowed; }
          .btn-close { background: #e0e0e0; color: #333; }
          
          #statusMsg { display: none; padding: 8px 12px; border-radius: 4px; font-size: 11px; margin-top: 10px; font-weight: 500; }
          .status-error { background: #fee2e2; color: #991b1b; border-left: 3px solid #ef4444; }
          .status-success { background: #dcfce7; color: #166534; border-left: 3px solid #22c55e; }
        </style>
      </head>
      <body>
        <div class="header">
          <i class="fa-solid fa-file-import"></i>
          <h3>Importar Configuración de Asignaturas</h3>
        </div>
        <div class="description">
          Arrastra y suelta un archivo JSON de configuración aquí, selecciónalo desde tu computadora o pega el texto JSON directamente en el cuadro de abajo.
        </div>
        
        <div class="drop-zone" onclick="triggerFileSelect()">
          <i class="fa-solid fa-cloud-arrow-up"></i>
          <p>Haz clic para buscar archivo o arrastra el JSON aquí</p>
          <input type="file" id="fileInput" accept=".json" style="display:none;" onchange="handleFileSelect(event)" />
        </div>
        
        <textarea id="jsonArea" placeholder='Pegar JSON de configuración aquí...' oninput="validateInput()"></textarea>
        
        <div id="statusMsg"></div>
        
        <div class="btn-group">
          <button class="btn-primary" id="btnImportar" onclick="startImport()" disabled><i class="fa-solid fa-circle-check"></i> Importar y Aplicar</button>
          <button class="btn-close" onclick="google.script.host.close()">Cancelar</button>
        </div>
        
        <script>
          function triggerFileSelect() {
            document.getElementById("fileInput").click();
          }
          
          function handleFileSelect(event) {
            var file = event.target.files[0];
            if (!file) return;
            var reader = new FileReader();
            reader.onload = function(e) {
              document.getElementById("jsonArea").value = e.target.result;
              validateInput();
            };
            reader.readAsText(file);
          }
          
          // Drag and drop events
          var dropZone = document.querySelector(".drop-zone");
          dropZone.addEventListener("dragover", function(e) {
            e.preventDefault();
            dropZone.style.borderColor = "#3b5998";
            dropZone.style.background = "#f1f5f9";
          });
          dropZone.addEventListener("dragleave", function(e) {
            e.preventDefault();
            dropZone.style.borderColor = "#cbd5e1";
            dropZone.style.background = "#ffffff";
          });
          dropZone.addEventListener("drop", function(e) {
            e.preventDefault();
            dropZone.style.borderColor = "#cbd5e1";
            dropZone.style.background = "#ffffff";
            var file = e.dataTransfer.files[0];
            if (file && file.name.endsWith(".json")) {
              var reader = new FileReader();
              reader.onload = function(evt) {
                document.getElementById("jsonArea").value = evt.target.result;
                validateInput();
              };
              reader.readAsText(file);
            } else {
              showStatus("Por favor arrastra únicamente archivos .json", false);
            }
          });
          
          function showStatus(msg, isSuccess) {
            var s = document.getElementById("statusMsg");
            s.innerText = msg;
            s.className = isSuccess ? "status-success" : "status-error";
            s.style.display = "block";
          }
          
          function validateInput() {
            var area = document.getElementById("jsonArea");
            var btn = document.getElementById("btnImportar");
            var txt = (area.value || '').trim();
            if (!txt) {
              btn.disabled = true;
              document.getElementById("statusMsg").style.display = "none";
              return;
            }
            
            try {
              var data = JSON.parse(txt);
              if (!Array.isArray(data)) {
                showStatus("La estructura debe ser un arreglo de asignaturas.", false);
                btn.disabled = true;
                return;
              }
              
              if (data.length === 0) {
                showStatus("El arreglo está vacío. No hay asignaturas que importar.", false);
                btn.disabled = true;
                return;
              }
              
              // Validar el formato de cada elemento
              for (var i = 0; i < data.length; i++) {
                var item = data[i];
                if (!item.name || !item.session || isNaN(item.count)) {
                  showStatus("Elemento " + (i + 1) + " inválido. Debe tener 'name', 'count' y 'session'.", false);
                  btn.disabled = true;
                  return;
                }
              }
              
              showStatus("JSON Válido. Detectadas " + data.length + " asignaturas configuradas.", true);
              btn.disabled = false;
            } catch (err) {
              showStatus("JSON Malformado: " + err.message, false);
              btn.disabled = true;
            }
          }
          
          function startImport() {
            var btn = document.getElementById("btnImportar");
            var area = document.getElementById("jsonArea");
            btn.disabled = true;
            btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Procesando...';
            
            google.script.run
              .withSuccessHandler(function(res) {
                showStatus(res, true);
                setTimeout(function() {
                  google.script.host.close();
                }, 2000);
              })
              .withFailureHandler(function(err) {
                showStatus("Error al importar: " + err.message, false);
                btn.disabled = false;
                btn.innerHTML = '<i class="fa-solid fa-circle-check"></i> Importar y Aplicar';
              })
              .executeJsonImport(area.value);
          }
        </script>
      </body>
    </html>
  `;

  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(480)
    .setTitle('📤 Importar Configuración (JSON)');
  ui.showModalDialog(htmlOutput, '📤 Importar Configuración (JSON)');
}

function executeJsonImport(jsonString) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(jsonString);

  if (!Array.isArray(data)) {
    throw new Error("El JSON debe ser un arreglo de asignaturas.");
  }

  // 1. Identificar sesiones únicas y asegurar que existan
  var sessions = {};
  data.forEach(function (item) {
    var sess = String(item.session).trim();
    sessions[sess] = true;
  });

  var createdSessions = [];
  Object.keys(sessions).forEach(function (sessNum) {
    var sessionName = "S" + sessNum;
    var existingSheet = ss.getSheetByName(sessionName);
    if (!existingSheet) {
      createSessionFromHtml(sessionName);
      createdSessions.push(sessionName);
    }
  });

  // 2. Obtener materias existentes para mapear colIndex
  var existingSubjectsData = getSubjectsBySession();
  var existingAll = [];
  Object.keys(existingSubjectsData).forEach(function (key) {
    existingAll = existingAll.concat(existingSubjectsData[key]);
  });

  // Mapa de existentes por nombre_sesion
  var existingMap = {};
  existingAll.forEach(function (s) {
    var key = s.name.toUpperCase().trim() + "_" + s.session;
    existingMap[key] = s.colIndex;
  });

  var assignedCols = {};
  var nextNewCol = 1000;

  // 3. Crear lista formateada para executeFullSync
  var formattedList = data.map(function (item) {
    var name = String(item.name).toUpperCase().trim();
    var session = String(item.session).trim();
    var count = parseInt(item.count, 10);
    var key = name + "_" + session;

    var colIndex = existingMap[key];
    if (colIndex && !assignedCols[colIndex]) {
      assignedCols[colIndex] = true;
    } else {
      colIndex = nextNewCol++;
    }

    return {
      name: name,
      count: count,
      session: session,
      colIndex: colIndex
    };
  });

  // 4. Ejecutar sincronización
  executeFullSyncFromHtml(formattedList);

  var msg = "¡Configuración aplicada con éxito!";
  if (createdSessions.length > 0) {
    msg += " Se crearon las sesiones: " + createdSessions.join(", ") + ".";
  }
  return msg;
}

function getFormulaSeparator() {
  try {
    var locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale() || "";
    var lang = locale.split(/[_-]/)[0].toLowerCase();
    // Idiomas comunes que usan coma decimal y punto y coma (;) como separador de argumentos
    var semicolonLangs = ["es", "fr", "de", "it", "pt", "ru", "nl", "sv", "no", "da", "fi", "pl", "cs", "sk", "hu", "ro", "tr", "el", "bg", "uk"];
    if (semicolonLangs.indexOf(lang) !== -1) {
      return ";";
    }
  } catch (e) { }
  return ",";
}

