// ============================================
// DriveOperations.gs - Work with Google Drive (ИСПРАВЛЕННАЯ ВЕРСИЯ)
// ============================================

// Upload files to Google Drive
function uploadFiles(filesData, invoiceNumber) {
  try {
    Logger.log('📎 Uploading ' + filesData.length + ' files for invoice ' + invoiceNumber);
    
    // Validate files count
    if (filesData.length > CONFIG.MAX_FILES_COUNT) {
      throw new Error('Максимальное количество файлов: ' + CONFIG.MAX_FILES_COUNT);
    }
    
    // Get main folder
    const mainFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    
    // Create subfolder for invoice
    const folderName = 'Invoice_' + invoiceNumber + '_' + new Date().toISOString().split('T')[0];
    const invoiceFolder = mainFolder.createFolder(folderName);
    
    const uploadedFiles = [];
    
    // Upload each file
    for (let i = 0; i < filesData.length; i++) {
      const fileData = filesData[i];
      
      try {
        // Validate file size
        if (fileData.size && fileData.size > CONFIG.MAX_FILE_SIZE) {
          const sizeMB = (fileData.size / 1024 / 1024).toFixed(2);
          const maxMB = (CONFIG.MAX_FILE_SIZE / 1024 / 1024).toFixed(2);
          Logger.log('⚠️ File ' + fileData.name + ' is too large: ' + sizeMB + 'MB (max: ' + maxMB + 'MB)');
          continue;
        }
        
        // Decode base64 and create blob
        const blob = Utilities.newBlob(
          Utilities.base64Decode(fileData.content),
          fileData.type,
          fileData.name
        );
        
        // Create file in folder
        const file = invoiceFolder.createFile(blob);
        
        // Set sharing permissions (anyone with link can view)
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // Save file info
        uploadedFiles.push({
          name: fileData.name,
          url: file.getUrl(),
          id: file.getId()
        });
        
        Logger.log('✅ File uploaded: ' + fileData.name);
        
      } catch (fileError) {
        Logger.log('❌ Error uploading file ' + fileData.name + ': ' + fileError);
      }
    }
    
    Logger.log('✅ Uploaded files: ' + uploadedFiles.length);
    return uploadedFiles;
    
  } catch (error) {
    Logger.log('❌ Error uploading files: ' + error);
    throw new Error('Failed to upload files: ' + error.message);
  }
}

// Create folder for invoice
function createInvoiceFolder(invoiceNumber) {
  try {
    const mainFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const folderName = 'Invoice_' + invoiceNumber + '_' + new Date().toISOString().split('T')[0];
    const newFolder = mainFolder.createFolder(folderName);
    
    Logger.log('📁 Created folder: ' + folderName);
    return newFolder.getId();
    
  } catch (error) {
    Logger.log('❌ Error creating folder: ' + error);
    throw error;
  }
}

// Get list of invoice files
function getInvoiceFiles(invoiceNumber) {
  try {
    const mainFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    
    // Search for folder by invoice number
    const folderSearchName = 'Invoice_' + invoiceNumber;
    const folders = mainFolder.getFolders();
    
    let invoiceFolder = null;
    while (folders.hasNext()) {
      const folder = folders.next();
      if (folder.getName().startsWith(folderSearchName)) {
        invoiceFolder = folder;
        break;
      }
    }
    
    if (!invoiceFolder) {
      Logger.log('📁 Folder not found for invoice: ' + invoiceNumber);
      return [];
    }
    
    const files = invoiceFolder.getFiles();
    const filesList = [];
    
    while (files.hasNext()) {
      const file = files.next();
      filesList.push({
        name: file.getName(),
        url: file.getUrl(),
        id: file.getId(),
        size: file.getSize(),
        mimeType: file.getMimeType()
      });
    }
    
    Logger.log('📎 Found ' + filesList.length + ' files');
    return filesList;
    
  } catch (error) {
    Logger.log('❌ Error getting files: ' + error);
    return [];
  }
}

// Delete file
function deleteFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
    Logger.log('🗑️ File deleted: ' + file.getName());
    return { success: true };
    
  } catch (error) {
    Logger.log('❌ Error deleting file: ' + error);
    return { success: false, error: error.message };
  }
}

// Get file by ID (for viewing/downloading)
function getFileById(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    return {
      name: file.getName(),
      url: file.getUrl(),
      downloadUrl: file.getDownloadUrl(),
      mimeType: file.getMimeType(),
      size: file.getSize()
    };
  } catch (error) {
    Logger.log('❌ Error getting file: ' + error);
    return null;
  }
}
