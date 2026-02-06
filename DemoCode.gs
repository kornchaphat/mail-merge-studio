/**
 * Mail Merge Studio - Main Backend
 * Google Apps Script Web App with SSO Authentication
 * Version 1.0
 */

// ==================== WEB APP ENTRY POINTS ====================

function doGet(e) {
  var user = Session.getActiveUser().getEmail();
  
  // Check if user is authenticated
  if (!user || user === '') {
    return HtmlService.createHtmlOutputFromFile('Login')
      .setTitle('Mail Merge Studio - Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // User is authenticated, show main app
  return HtmlService.createHtmlOutputFromFile('App')
    .setTitle('Mail Merge Studio')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==================== AUTHENTICATION ====================

function getCurrentUser() {
  var email = Session.getActiveUser().getEmail();
  if (!email) {
    return { authenticated: false };
  }
  
  return {
    authenticated: true,
    email: email,
    name: getUserName(email)
  };
}

function getUserName(email) {
  try {
    var namePart = email.split('@')[0];
    return namePart.split('.').map(function(part) {
      return part.charAt(0).toUpperCase() + part.slice(1);
    }).join(' ');
  } catch (e) {
    return email;
  }
}

function getAuthUrl() {
  return ScriptApp.getService().getUrl();
}

// ==================== GOOGLE DRIVE OPERATIONS ====================

function getRecentSpreadsheets(maxResults) {
  maxResults = maxResults || 10;
  var files = [];
  
  try {
    var fileIterator = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
    var count = 0;
    
    while (fileIterator.hasNext() && count < maxResults) {
      var file = fileIterator.next();
      files.push({
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl(),
        lastUpdated: file.getLastUpdated().toISOString()
      });
      count++;
    }
  } catch (e) {
    console.error('Error getting spreadsheets:', e);
  }
  
  return files;
}

function getFolders() {
  var folders = [];
  
  try {
    var folderIterator = DriveApp.getFolders();
    
    while (folderIterator.hasNext()) {
      var folder = folderIterator.next();
      folders.push({
        id: folder.getId(),
        name: folder.getName(),
        url: folder.getUrl()
      });
    }
  } catch (e) {
    console.error('Error getting folders:', e);
  }
  
  return folders.sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });
}

function createFolder(name, parentId) {
  try {
    var parent = parentId ? DriveApp.getFolderById(parentId) : DriveApp.getRootFolder();
    var newFolder = parent.createFolder(name);
    return {
      success: true,
      id: newFolder.getId(),
      name: newFolder.getName(),
      url: newFolder.getUrl()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== SPREADSHEET OPERATIONS ====================

function readSpreadsheetData(fileId) {
  try {
    var ss = SpreadsheetApp.openById(fileId);
    var sheet = ss.getActiveSheet();
    var data = sheet.getDataRange().getDisplayValues();
    
    if (data.length < 2) {
      return { success: false, error: 'Spreadsheet must have headers and at least one data row' };
    }
    
    var headers = data[0].map(function(h) {
      return String(h).trim();
    });
    var rows = data.slice(1).filter(function(row) {
      return row.some(function(cell) {
        return cell !== '';
      });
    });
    
    return {
      success: true,
      headers: headers,
      rows: rows,
      totalRows: rows.length,
      sheetName: sheet.getName()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== TEMPLATE OPERATIONS ====================

function getTemplateById(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var mimeType = file.getMimeType();
    var templateType = 'doc'; // default
    var text = '';
    var name = '';
    var htmlContent = '';
    
    // Detect template type
    if (mimeType === 'application/vnd.google-apps.presentation') {
      templateType = 'slides';
      var presentation = SlidesApp.openById(fileId);
      name = presentation.getName();
      
      // Extract text from all slides
      var slides = presentation.getSlides();
      for (var i = 0; i < slides.length; i++) {
        var shapes = slides[i].getShapes();
        for (var j = 0; j < shapes.length; j++) {
          var shape = shapes[j];
          if (shape.getText) {
            text += shape.getText().asString() + '\n';
          }
        }
        var tables = slides[i].getTables();
        for (var t = 0; t < tables.length; t++) {
          var table = tables[t];
          for (var r = 0; r < table.getNumRows(); r++) {
            for (var c = 0; c < table.getNumColumns(); c++) {
              var cell = table.getCell(r, c);
              if (cell.getText) {
                text += cell.getText().asString() + '\n';
              }
            }
          }
        }
      }
      
      // Simple HTML preview for slides
      htmlContent = '<p style="color:#666;font-style:italic;">📊 Google Slides template with ' + slides.length + ' slide(s)</p>';
      htmlContent += '<p>' + text.substring(0, 500).replace(/\n/g, '<br>') + (text.length > 500 ? '...' : '') + '</p>';
      
    } else {
      // Google Docs
      templateType = 'doc';
      var doc = DocumentApp.openById(fileId);
      var body = doc.getBody();
      text = body.getText();
      name = doc.getName();
      
      // Get HTML content for better preview
      try {
        var blob = DriveApp.getFileById(fileId).getAs('text/html');
        htmlContent = blob.getDataAsString();
        var bodyMatch = htmlContent.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
        if (bodyMatch) {
          htmlContent = bodyMatch[1];
        }
      } catch (htmlError) {
        htmlContent = '<p>' + text.replace(/\n\n/g, '</p><p>').replace(/\n/g, '<br>') + '</p>';
      }
    }
    
    // Extract placeholders
    var placeholders = [];
    var regex = /\{\{([^}]+)\}\}/g;
    var match;
    while ((match = regex.exec(text)) !== null) {
      if (placeholders.indexOf(match[1]) === -1) {
        placeholders.push(match[1]);
      }
    }
    
    return {
      success: true,
      id: fileId,
      name: name,
      type: templateType,
      content: text,
      htmlContent: htmlContent,
      placeholders: placeholders
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== MAIL MERGE GENERATION ====================

function generateDocuments(config) {
  var results = {
    success: true,
    generated: 0,
    failed: 0,
    files: [],
    errors: [],
    logData: []
  };
  
  try {
    var dataRows = config.dataRows;
    var headers = config.headers;
    var templates = config.templates;
    var defaultTemplateId = config.defaultTemplateId;
    var rules = config.rules || [];
    var mapping = config.mapping || {};
    var filenamePattern = config.filenamePattern;
    var outputFormat = config.outputFormat;
    var folderId = config.folderId;
    var replaceExisting = config.replaceExisting;
    var startIndex = config.startIndex || 0;
    var useSubfolders = config.subfolders || false;
    var subfolderLevels = config.subfolderLevels || [];
    
    var outputFolder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
    
    // Cache for created subfolders to avoid recreating
    var subfolderCache = {};
    
    for (var i = 0; i < dataRows.length; i++) {
      var row = dataRows[i];
      var rowData = {};
      var actualRowNum = startIndex + i + 1;
      
      for (var j = 0; j < headers.length; j++) {
        rowData[headers[j]] = row[j] || '';
      }
      
      try {
        var templateId = getTemplateForRow(rowData, headers, row, rules, defaultTemplateId, templates);
        
        if (!templateId) {
          throw new Error('No template specified');
        }
        
        // Determine target folder (with subfolder support)
        var targetFolder = outputFolder;
        if (useSubfolders && subfolderLevels.length > 0) {
          targetFolder = getOrCreateSubfolder(outputFolder, rowData, subfolderLevels, subfolderCache);
        }
        
        var filename = generateFilename(filenamePattern, rowData, headers, actualRowNum);
        
        if (!filename || filename.trim() === '') {
          throw new Error('Generated filename is empty');
        }
        
        if (replaceExisting) {
          deleteExistingFile(targetFolder, filename + '.pdf');
          deleteExistingFile(targetFolder, filename + '.docx');
          deleteExistingFile(targetFolder, filename + '.pptx');
        }
        
        var templateFile = DriveApp.getFileById(templateId);
        var mimeType = templateFile.getMimeType();
        var isSlides = (mimeType === 'application/vnd.google-apps.presentation');
        var fileCopy = templateFile.makeCopy(filename, targetFolder);
        
        // Build replacement map
        var replacements = {};
        for (var placeholder in mapping) {
          if (mapping.hasOwnProperty(placeholder)) {
            var columnName = mapping[placeholder];
            replacements['{{' + placeholder + '}}'] = rowData[columnName] || '';
          }
        }
        for (var k = 0; k < headers.length; k++) {
          replacements['{{' + headers[k] + '}}'] = rowData[headers[k]] || '';
        }
        replacements['{{Year}}'] = new Date().getFullYear().toString();
        replacements['{{Date}}'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        replacements['{{Counter}}'] = padNumber(actualRowNum, 3);
        
        if (isSlides) {
          // Process Google Slides
          var presentation = SlidesApp.openById(fileCopy.getId());
          var slides = presentation.getSlides();
          
          for (var s = 0; s < slides.length; s++) {
            var slide = slides[s];
            
            // Replace in shapes
            var shapes = slide.getShapes();
            for (var sh = 0; sh < shapes.length; sh++) {
              var shape = shapes[sh];
              if (shape.getText) {
                var textRange = shape.getText();
                for (var placeholder in replacements) {
                  textRange.replaceAllText(placeholder, replacements[placeholder]);
                }
              }
            }
            
            // Replace in tables
            var tables = slide.getTables();
            for (var tb = 0; tb < tables.length; tb++) {
              var table = tables[tb];
              for (var r = 0; r < table.getNumRows(); r++) {
                for (var c = 0; c < table.getNumColumns(); c++) {
                  var cell = table.getCell(r, c);
                  if (cell.getText) {
                    var cellText = cell.getText();
                    for (var placeholder in replacements) {
                      cellText.replaceAllText(placeholder, replacements[placeholder]);
                    }
                  }
                }
              }
            }
          }
          
          presentation.saveAndClose();
          
        } else {
          // Process Google Docs
          var doc = DocumentApp.openById(fileCopy.getId());
          var body = doc.getBody();
          
          for (var placeholder in replacements) {
            var regexPattern = placeholder.replace(/\{/g, '\\{').replace(/\}/g, '\\}');
            body.replaceText(regexPattern, replacements[placeholder]);
          }
          
          doc.saveAndClose();
        }
        
        var generatedFiles = [];
        
        if (outputFormat === 'pdf' || outputFormat === 'both') {
          var pdfBlob = DriveApp.getFileById(fileCopy.getId()).getAs('application/pdf');
          pdfBlob.setName(filename + '.pdf');
          var pdfFile = targetFolder.createFile(pdfBlob);
          generatedFiles.push({ name: filename + '.pdf', id: pdfFile.getId(), url: pdfFile.getUrl() });
        }
        
        if (outputFormat === 'docx' || outputFormat === 'both') {
          fileCopy.setName(filename);
          generatedFiles.push({ name: filename, id: fileCopy.getId(), url: fileCopy.getUrl() });
        }
        
        if (outputFormat === 'pdf') {
          fileCopy.setTrashed(true);
        }
        
        results.generated++;
        for (var g = 0; g < generatedFiles.length; g++) {
          results.files.push(generatedFiles[g]);
        }
        results.logData.push({
          row: i + 1,
          filename: filename,
          status: 'success'
        });
        
      } catch (rowError) {
        results.failed++;
        results.errors.push({ row: i + 1, error: rowError.message });
        results.logData.push({
          row: i + 1,
          filename: '',
          status: 'failed',
          error: rowError.message
        });
      }
    }
    
  } catch (e) {
    results.success = false;
    results.error = e.message;
  }
  
  return results;
}

/**
 * Get or create nested subfolders based on column values
 */
function getOrCreateSubfolder(parentFolder, rowData, subfolderLevels, cache) {
  var currentFolder = parentFolder;
  var cacheKey = '';
  
  for (var i = 0; i < subfolderLevels.length; i++) {
    var columnName = subfolderLevels[i];
    var folderName = sanitizeFolderName(rowData[columnName] || 'Unknown');
    
    cacheKey += '/' + folderName;
    
    // Check cache first
    if (cache[cacheKey]) {
      currentFolder = cache[cacheKey];
    } else {
      // Look for existing folder or create new one
      var folders = currentFolder.getFoldersByName(folderName);
      if (folders.hasNext()) {
        currentFolder = folders.next();
      } else {
        currentFolder = currentFolder.createFolder(folderName);
      }
      cache[cacheKey] = currentFolder;
    }
  }
  
  return currentFolder;
}

/**
 * Sanitize folder name - remove invalid characters
 */
function sanitizeFolderName(name) {
  if (!name) return 'Unknown';
  // Remove characters not allowed in folder names
  return name.toString().replace(/[\/\\:*?"<>|]/g, '_').trim() || 'Unknown';
}

/**
 * Determine which template to use for a given row based on rules
 */
function getTemplateForRow(rowData, headers, row, rules, defaultTemplateId, templates) {
  // If no rules, use default
  if (!rules || rules.length === 0) {
    return defaultTemplateId || (templates && templates.length > 0 ? templates[0].id : null);
  }
  
  // Evaluate each rule in order (first match wins)
  for (var r = 0; r < rules.length; r++) {
    var rule = rules[r];
    var conditions = rule.conditions || [];
    var allMatch = true;
    
    // Check all conditions (AND logic)
    for (var c = 0; c < conditions.length; c++) {
      var cond = conditions[c];
      var colIndex = headers.indexOf(cond.col);
      var cellValue = colIndex >= 0 ? String(row[colIndex] || '').toLowerCase() : '';
      var testValue = (cond.val || '').toLowerCase();
      
      var match = evaluateCondition(cond.op, cellValue, testValue);
      if (!match) {
        allMatch = false;
        break;
      }
    }
    
    if (allMatch && rule.templateId) {
      return rule.templateId;
    }
  }
  
  // No rules matched, use default
  return defaultTemplateId || (templates && templates.length > 0 ? templates[0].id : null);
}

/**
 * Evaluate a single condition
 */
function evaluateCondition(op, cellValue, testValue) {
  switch (op) {
    case 'equals':
      return cellValue === testValue;
    case 'not_equals':
      return cellValue !== testValue;
    case 'contains':
      return cellValue.indexOf(testValue) !== -1;
    case 'not_contains':
      return cellValue.indexOf(testValue) === -1;
    case 'starts_with':
      return cellValue.indexOf(testValue) === 0;
    case 'ends_with':
      return cellValue.slice(-testValue.length) === testValue;
    case 'is_blank':
      return cellValue === '';
    case 'is_not_blank':
      return cellValue !== '';
    case 'greater_than':
      return parseFloat(cellValue) > parseFloat(testValue);
    case 'less_than':
      return parseFloat(cellValue) < parseFloat(testValue);
    default:
      return true;
  }
}

function generateFilename(pattern, rowData, headers, index) {
  var filename = pattern;
  
  for (var i = 0; i < headers.length; i++) {
    var key = headers[i];
    var regex = new RegExp('\\{\\{' + key + '\\}\\}', 'g');
    filename = filename.replace(regex, rowData[key] || '');
  }
  
  filename = filename.replace(/\{\{Year\}\}/g, new Date().getFullYear().toString());
  filename = filename.replace(/\{\{Date\}\}/g, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  filename = filename.replace(/\{\{Counter\}\}/g, padNumber(index, 3));
  filename = filename.replace(/[\/\\?%*:|"<>]/g, '-');
  
  return filename;
}

function padNumber(num, size) {
  var s = num.toString();
  while (s.length < size) s = '0' + s;
  return s;
}

function deleteExistingFile(folder, filename) {
  var files = folder.getFilesByName(filename);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

// ==================== ZIP DOWNLOAD ====================

function createZipDownload(fileIds) {
  try {
    var blobs = [];
    for (var i = 0; i < fileIds.length; i++) {
      var file = DriveApp.getFileById(fileIds[i]);
      blobs.push(file.getBlob());
    }
    
    var zipBlob = Utilities.zip(blobs, 'MailMerge_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.zip');
    var zipFile = DriveApp.createFile(zipBlob);
    
    return {
      success: true,
      id: zipFile.getId(),
      url: zipFile.getDownloadUrl(),
      name: zipFile.getName()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== CONFIGURATION ====================

function getConfig() {
  return {
    appName: 'Mail Merge Studio',
    version: '2.0'
  };
}

// ==================== DATABASE (Google Sheets) ====================

/**
 * Database Sheet ID - Set via Project Settings > Script Properties
 * Property Name: DB_SHEET_ID
 * Property Value: Your Google Sheet ID
 */
function getDbSheetId() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sheetId = scriptProperties.getProperty('DB_SHEET_ID');
  
  if (!sheetId) {
    throw new Error('Database not configured. Please set DB_SHEET_ID in Script Properties (Project Settings > Script Properties)');
  }
  
  return sheetId;
}

/**
 * Initialize database sheets with headers if they don't exist
 */
function initDatabase() {
  try {
    var ss = SpreadsheetApp.openById(getDbSheetId());
    
    // Initialize Presets sheet
    var presetsSheet = ss.getSheetByName('Presets');
    if (!presetsSheet) {
      presetsSheet = ss.insertSheet('Presets');
      presetsSheet.getRange(1, 1, 1, 6).setValues([[
        'id', 'user_email', 'preset_name', 'config_json', 'created_at', 'updated_at'
      ]]);
      presetsSheet.setFrozenRows(1);
      presetsSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }
    
    // Initialize Logs sheet
    var logsSheet = ss.getSheetByName('Logs');
    if (!logsSheet) {
      logsSheet = ss.insertSheet('Logs');
      logsSheet.getRange(1, 1, 1, 12).setValues([[
        'id', 'user_email', 'user_name', 'timestamp', 'data_source', 'records_count', 
        'templates', 'format', 'delivery', 'success_count', 'failed_count', 'duration_sec'
      ]]);
      logsSheet.setFrozenRows(1);
      logsSheet.getRange(1, 1, 1, 12).setFontWeight('bold');
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== PRESETS CRUD ====================

/**
 * Get all presets for the current user
 */
function getPresets() {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.openById(getDbSheetId());
    var sheet = ss.getSheetByName('Presets');
    
    if (!sheet) {
      initDatabase();
      return { success: true, presets: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var presets = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[1] === userEmail) { // user_email column
        presets.push({
          id: row[0],
          name: row[2],
          config: JSON.parse(row[3] || '{}'),
          createdAt: row[4],
          updatedAt: row[5]
        });
      }
    }
    
    return { success: true, presets: presets };
  } catch (e) {
    return { success: false, error: e.message, presets: [] };
  }
}

/**
 * Save a new preset or update existing
 */
function savePreset(presetName, configJson) {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.openById(getDbSheetId());
    var sheet = ss.getSheetByName('Presets');
    
    if (!sheet) {
      initDatabase();
      sheet = ss.getSheetByName('Presets');
    }
    
    var data = sheet.getDataRange().getValues();
    var now = new Date().toISOString();
    var existingRow = -1;
    
    // Check if preset with same name exists for this user
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === userEmail && data[i][2] === presetName) {
        existingRow = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }
    
    if (existingRow > 0) {
      // Update existing preset
      sheet.getRange(existingRow, 4).setValue(JSON.stringify(configJson)); // config_json
      sheet.getRange(existingRow, 6).setValue(now); // updated_at
      return { success: true, action: 'updated', id: data[existingRow - 1][0] };
    } else {
      // Create new preset
      var newId = 'preset_' + Date.now();
      sheet.appendRow([newId, userEmail, presetName, JSON.stringify(configJson), now, now]);
      return { success: true, action: 'created', id: newId };
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Delete a preset by ID
 */
function deletePreset(presetId) {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.openById(getDbSheetId());
    var sheet = ss.getSheetByName('Presets');
    
    if (!sheet) {
      return { success: false, error: 'Presets sheet not found' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === presetId && data[i][1] === userEmail) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    
    return { success: false, error: 'Preset not found or access denied' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== GENERATION LOGS ====================

/**
 * Log a generation run
 */
function logGeneration(logData) {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var userName = getUserName(userEmail);
    var ss = SpreadsheetApp.openById(getDbSheetId());
    var sheet = ss.getSheetByName('Logs');
    
    if (!sheet) {
      initDatabase();
      sheet = ss.getSheetByName('Logs');
    }
    
    var logId = 'log_' + Date.now();
    var timestamp = new Date().toISOString();
    
    sheet.appendRow([
      logId,
      userEmail,
      userName,
      timestamp,
      logData.dataSource || '',
      logData.recordsCount || 0,
      logData.templates || '',
      logData.format || '',
      logData.delivery || '',
      logData.successCount || 0,
      logData.failedCount || 0,
      logData.durationSec || 0
    ]);
    
    return { success: true, logId: logId };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Get generation logs for current user (last 50)
 */
function getMyLogs(limit) {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.openById(getDbSheetId());
    var sheet = ss.getSheetByName('Logs');
    
    if (!sheet) {
      return { success: true, logs: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var logs = [];
    
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      if (row[1] === userEmail) {
        logs.push({
          id: row[0],
          timestamp: row[3],
          dataSource: row[4],
          recordsCount: row[5],
          templates: row[6],
          format: row[7],
          delivery: row[8],
          successCount: row[9],
          failedCount: row[10],
          durationSec: row[11]
        });
        if (logs.length >= (limit || 50)) break;
      }
    }
    
    return { success: true, logs: logs };
  } catch (e) {
    return { success: false, error: e.message, logs: [] };
  }
}

/**
 * Get all logs (admin function)
 */
function getAllLogs(limit) {
  try {
    var ss = SpreadsheetApp.openById(getDbSheetId());
    var sheet = ss.getSheetByName('Logs');
    
    if (!sheet) {
      return { success: true, logs: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var logs = [];
    
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      logs.push({
        id: row[0],
        userEmail: row[1],
        userName: row[2],
        timestamp: row[3],
        dataSource: row[4],
        recordsCount: row[5],
        templates: row[6],
        format: row[7],
        delivery: row[8],
        successCount: row[9],
        failedCount: row[10],
        durationSec: row[11]
      });
      if (logs.length >= (limit || 100)) break;
    }
    
    return { success: true, logs: logs };
  } catch (e) {
    return { success: false, error: e.message, logs: [] };
  }
}

/**
 * Get usage statistics
 */
function getUsageStats() {
  try {
    var ss = SpreadsheetApp.openById(getDbSheetId());
    var sheet = ss.getSheetByName('Logs');
    
    if (!sheet) {
      return { success: true, stats: { totalRuns: 0, totalDocs: 0, uniqueUsers: 0 } };
    }
    
    var data = sheet.getDataRange().getValues();
    var totalRuns = data.length - 1;
    var totalDocs = 0;
    var users = {};
    
    for (var i = 1; i < data.length; i++) {
      totalDocs += (data[i][9] || 0); // success_count
      users[data[i][1]] = true; // user_email
    }
    
    return {
      success: true,
      stats: {
        totalRuns: totalRuns,
        totalDocs: totalDocs,
        uniqueUsers: Object.keys(users).length
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== OUTPUT FOLDER OPERATIONS ====================

/**
 * Create an output folder for mail merge results
 */
function createOutputFolder(folderName, parentFolderId) {
  try {
    var name = folderName || 'Mail Merge - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var folder;
    
    if (parentFolderId) {
      var parentFolder = DriveApp.getFolderById(parentFolderId);
      folder = parentFolder.createFolder(name);
    } else {
      folder = DriveApp.createFolder(name);
    }
    
    // Store folder ID for later use
    var scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('LAST_OUTPUT_FOLDER_ID', folder.getId());
    
    return {
      success: true,
      folderId: folder.getId(),
      folderUrl: folder.getUrl(),
      folderName: folder.getName()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Get the URL of the output folder
 */
function getOutputFolderUrl() {
  try {
    var scriptProps = PropertiesService.getScriptProperties();
    var folderId = scriptProps.getProperty('LAST_OUTPUT_FOLDER_ID');
    
    if (!folderId) {
      return { success: false, error: 'No output folder found' };
    }
    
    var folder = DriveApp.getFolderById(folderId);
    return {
      success: true,
      folderId: folderId,
      url: folder.getUrl(),
      name: folder.getName()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Extract folder ID from a Drive folder URL
 */
function extractFolderIdFromUrl(url) {
  if (!url) return null;
  var match = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Create a ZIP file from all files in the output folder
 */
function createOutputZip(folderId) {
  try {
    var folder;
    
    if (folderId) {
      folder = DriveApp.getFolderById(folderId);
    } else {
      var scriptProps = PropertiesService.getScriptProperties();
      var lastFolderId = scriptProps.getProperty('LAST_OUTPUT_FOLDER_ID');
      if (!lastFolderId) {
        return { success: false, error: 'No output folder found' };
      }
      folder = DriveApp.getFolderById(lastFolderId);
    }
    
    var files = folder.getFiles();
    var blobs = [];
    
    while (files.hasNext()) {
      var file = files.next();
      blobs.push(file.getBlob().setName(file.getName()));
    }
    
    if (blobs.length === 0) {
      return { success: false, error: 'No files in folder to ZIP' };
    }
    
    var zipBlob = Utilities.zip(blobs, folder.getName() + '.zip');
    var zipFile = folder.createFile(zipBlob);
    
    return {
      success: true,
      zipId: zipFile.getId(),
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + zipFile.getId(),
      fileName: zipFile.getName(),
      fileCount: blobs.length
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== GENERATION PROGRESS (for resume capability) ====================

/**
 * Save generation progress for resume capability
 */
function saveGenerationProgress(progressData) {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var scriptProps = PropertiesService.getScriptProperties();
    
    var progress = {
      sessionId: progressData.sessionId,
      userEmail: userEmail,
      folderId: progressData.folderId,
      folderUrl: progressData.folderUrl,
      totalRows: progressData.totalRows,
      processedCount: progressData.processedCount,
      successCount: progressData.successCount,
      failedCount: progressData.failedCount,
      updatedAt: new Date().toISOString()
    };
    
    scriptProps.setProperty('GENERATION_PROGRESS_' + userEmail, JSON.stringify(progress));
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Get saved generation progress for current user
 */
function getGenerationProgress() {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var scriptProps = PropertiesService.getScriptProperties();
    
    var progressJson = scriptProps.getProperty('GENERATION_PROGRESS_' + userEmail);
    if (!progressJson) {
      return { success: true, hasProgress: false };
    }
    
    var progress = JSON.parse(progressJson);
    
    // Check if progress is recent (within last hour)
    var updatedAt = new Date(progress.updatedAt);
    var hourAgo = new Date(Date.now() - 60 * 60 * 1000);
    
    if (updatedAt < hourAgo) {
      // Progress is stale, clear it
      scriptProps.deleteProperty('GENERATION_PROGRESS_' + userEmail);
      return { success: true, hasProgress: false };
    }
    
    // Check if folder still exists
    try {
      var folder = DriveApp.getFolderById(progress.folderId);
      progress.folderName = folder.getName();
    } catch (e) {
      // Folder doesn't exist, clear progress
      scriptProps.deleteProperty('GENERATION_PROGRESS_' + userEmail);
      return { success: true, hasProgress: false };
    }
    
    return { success: true, hasProgress: true, progress: progress };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Clear generation progress for current user
 */
function clearGenerationProgress() {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var scriptProps = PropertiesService.getScriptProperties();
    scriptProps.deleteProperty('GENERATION_PROGRESS_' + userEmail);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}