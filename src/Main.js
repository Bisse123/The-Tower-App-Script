function doGet() {
  const template = HtmlService.createTemplateFromFile('main');
  return template.evaluate()
    // .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getPickerHtml() {
  var template = HtmlService.createTemplateFromFile('pickerTest');
  template.origin = "https://script.google.com";

  return template.evaluate().getContent();
}

function getFileIdsArray() {
  const fileIds = [
    "1matne6T83-cQNHppR1xpdVP3rlkQb5kC7Qxz0sh0BZU",
    "1me78RwHtYprlCJflAuaIY_9gxoyb0zldss70bIQy40g",
    "1JyJ8ta_Py3akmWPCJbNkX1C_N2STU-mywEjlt5YLnXc",
    "1VhvINtUXcqcdHe5bvStGDRb6bzDdlHe1bApJhNgQrBA"
  ];
  return fileIds;
}

function getOAuthToken() {
  try {
    const token = ScriptApp.getOAuthToken();
    return {
      success: true,
      token: token,
      message: 'Token retrieved successfully'
    };
  } catch (error) {
    console.error('Error getting OAuth token:', error);
    return {
      success: false,
      token: null,
      message: error.toString()
    };
  }
}

function getSheetIds() {
  try {
    const fileIds = getFileIdsArray();
    return {
      success: true,
      fileIds: fileIds,
      message: 'Sheet IDs retrieved successfully'
    };
  } catch (error) {
    console.error('Error getting sheet IDs:', error);
    return {
      success: false,
      fileIds: [],
      message: error.toString()
    };
  }
}

function checkSheetAccess() {
  try {
    const fileIds = getFileIdsArray();
    const accessibleFiles = [];
    const inaccessibleFiles = [];
    const accessibleDetails = [];
    
    console.log('Checking access to', fileIds.length, 'predefined sheets...');
    
    fileIds.forEach(fileId => {
      try {
        const file = Drive.Files.get(fileId, {
          fields: 'id,name,mimeType,webViewLink'
        });
        
        console.log('Access confirmed for:', file.name, '(' + fileId + ')');
        accessibleFiles.push(fileId);
        accessibleDetails.push({
          id: file.id,
          name: file.name,
          mimeType: file.mimeType,
          webViewLink: file.webViewLink
        });
        
      } catch (error) {
        console.log('No access to file:', fileId, 'Error:', error.toString());
        let name = null;
        try {
          const ss = SpreadsheetApp.openById(fileId);
          name = ss.getName();
        } catch (e) {
        }
        inaccessibleFiles.push({ id: fileId, name: name });
      }
    });
    
    console.log('Access check complete. Accessible:', accessibleFiles.length, 'Inaccessible:', inaccessibleFiles.length);
    
    return {
      success: true,
      accessibleFiles: accessibleFiles,
      inaccessibleFiles: inaccessibleFiles,
      accessibleDetails: accessibleDetails,
      totalFiles: fileIds.length,
      message: `Access check complete. ${accessibleFiles.length} of ${fileIds.length} sheets are accessible.`
    };
    
  } catch (error) {
    console.error('Error checking sheet access:', error);
    return {
      success: false,
      accessibleFiles: [],
      inaccessibleFiles: [],
      accessibleDetails: [],
      message: error.toString()
    };
  }
}

function processSelectedFiles(fileIds) {
  try {
    console.log('Processing file IDs:', fileIds);
    const fileDetails = fileIds.map(fileId => {
      const file = Drive.Files.get(fileId, {
        fields: 'id,name,mimeType,size,modifiedTime,webViewLink,parents'
      });
      return {
        id: file.id,
        name: file.name,
        mimeType: file.mimeType,
        size: file.size,
        modifiedTime: file.modifiedTime,
        webViewLink: file.webViewLink,
        parents: file.parents
      };
    });
    
    return {
      success: true,
      message: `Processed ${fileIds.length} files successfully`,
      files: fileDetails
    };
    
  } catch (error) {
    console.error('Error processing files:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

function showImportDialog() {
  var template = HtmlService.createTemplateFromFile('pickerTest');
  template.origin = "https://docs.google.com";
  var html = template.evaluate()
    .setWidth(1500)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Tower Import Data');
}