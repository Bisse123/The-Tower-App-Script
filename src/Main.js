function doGet() {
  const fileIdsArray = getFileIdsArray();
  
  const template = HtmlService.createTemplateFromFile('pickerTest');
  template.fileIdsArray = fileIdsArray;
  
  return template.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
function getFileIdsArray() {
  const fileIds = [
    "1matne6T83-cQNHppR1xpdVP3rlkQb5kC7Qxz0sh0BZU",
    "1me78RwHtYprlCJflAuaIY_9gxoyb0zldss70bIQy40g",
    "1JyJ8ta_Py3akmWPCJbNkX1C_N2STU-mywEjlt5YLnXc"
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
function testDriveAccess() {
  try {
    const files = Drive.Files.list({
      pageSize: 1,
      fields: 'files(id,name)'
    });
    
    return {
      success: true,
      message: 'Drive access working',
      fileCount: files.files ? files.files.length : 0
    };
  } catch (error) {
    console.error('Drive access error:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Get sheet IDs for the picker and check which ones are already accessible
 */
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

/**
 * Check which predefined sheets the user already has access to
 * Returns accessible and inaccessible file IDs separately
 */
function checkSheetAccess() {
  try {
    const fileIds = getFileIdsArray();
    const accessibleFiles = [];
    const inaccessibleFiles = [];
    const accessibleDetails = [];
    
    console.log('Checking access to', fileIds.length, 'predefined sheets...');
    
    fileIds.forEach(fileId => {
      try {
        // Try to get basic file information - this will work if user has access
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
        inaccessibleFiles.push(fileId);
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
    
    // Your processing logic here
    // Since user selected these files, you have drive.file scope access
    
    // Example: Get file details using Drive API
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