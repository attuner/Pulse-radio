/**
 * PULSE RADIO ENGINE - BACKEND ORCHESTRATOR (Code.gs)
 * Configuration: Set your target Google Drive Folder ID below.
 */
var CONFIG = {
  DRIVE_FOLDER_ID: "PASTE_YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE",
  USERS_SHEET: "USERS_DB",
  MANIFEST_SHEET: "RADIO_CONTENT_MANIFEST",
  MAX_ACTIVE_TRACKS: 50,
  SALT: "::pulse_radio_salt_2026"
};

/**
 * Handle HTTP GET Requests (Public Manifest Polling, Binary Audio Streaming & Heartbeat)
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getLiveManifest";
  var result = {};
  
  try {
    initDatabaseSheets();
    
    // Dedicated Audio Streaming Proxy to bypass Drive 302 redirects and virus-scan screens
    if (action === "streamAudio") {
      var fileId = (e && e.parameter && e.parameter.fileId) ? e.parameter.fileId : null;
      if (!fileId) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: "Missing fileId parameter for audio stream."
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      try {
        var file = DriveApp.getFileById(fileId);
        if (file.isTrashed()) {
          return ContentService.createTextOutput(JSON.stringify({
            success: false,
            error: "Requested file has been purged or trashed."
          })).setMimeType(ContentService.MimeType.JSON);
        }
        var blob = file.getBlob();
        var base64Data = Utilities.base64Encode(blob.getBytes());
        
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          mimeType: blob.getContentType(),
          base64: base64Data
        })).setMimeType(ContentService.MimeType.JSON);
      } catch (fileErr) {
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          error: "Failed to read audio bytes from Drive: " + fileErr.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    } else if (action === "getLiveManifest") {
      var listenerId = (e && e.parameter && e.parameter.listenerId) ? e.parameter.listenerId : null;
      result = handleGetLiveManifest(listenerId);
    } else if (action === "getBufferStats") {
      result = handleGetBufferStats();
    } else {
      result = { success: false, error: "Invalid GET action parameter" };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle HTTP POST Requests (Registration, Authentication, Uploads & Live Voiceover)
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  var result = {};
  
  try {
    lock.waitLock(30000);
    initDatabaseSheets();
    
    var requestBody = {};
    if (e && e.postData && e.postData.contents) {
      requestBody = JSON.parse(e.postData.contents);
    } else {
      throw new Error("Empty or malformed POST payload.");
    }
    
    var action = requestBody.action;
    
    if (action === "registerUser") {
      result = handleRegisterUser(requestBody);
    } else if (action === "authenticateUser") {
      result = handleAuthenticateUser(requestBody);
    } else if (action === "uploadAudioTrack") {
      result = handleUploadAudioTrack(requestBody);
    } else if (action === "broadcastLiveVoiceover") {
      result = handleBroadcastLiveVoiceover(requestBody);
    } else {
      result = { success: false, error: "Unrecognized POST action: " + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Sheet Verification and Auto-Initialization Routine
 */
function initDatabaseSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet 1: USERS_DB
  var usersSheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(CONFIG.USERS_SHEET);
    usersSheet.appendRow([
      "user_id",
      "username",
      "password_hash",
      "full_name",
      "registration_timestamp",
      "last_login_timestamp",
      "total_uploads_count",
      "account_status"
    ]);
    usersSheet.getRange("A1:H1").setFontWeight("bold");
    usersSheet.setFrozenRows(1);
  }
  
  // Sheet 2: RADIO_CONTENT_MANIFEST
  var manifestSheet = ss.getSheetByName(CONFIG.MANIFEST_SHEET);
  if (!manifestSheet) {
    manifestSheet = ss.insertSheet(CONFIG.MANIFEST_SHEET);
    manifestSheet.appendRow([
      "content_id",
      "file_id",
      "file_name",
      "title",
      "uploader_username",
      "uploader_name",
      "mime_type",
      "file_size_bytes",
      "audio_duration_seconds",
      "upload_timestamp",
      "playback_direct_url",
      "status",
      "play_count"
    ]);
    manifestSheet.getRange("A1:M1").setFontWeight("bold");
    manifestSheet.setFrozenRows(1);
  }
}

/**
 * Register User Controller
 */
function handleRegisterUser(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  var rows = sheet.getDataRange().getValues();
  
  var username = (data.username || "").toLowerCase().trim();
  var passwordHash = (data.passwordHash || "").trim();
  var fullName = (data.fullName || "").trim();
  
  if (!username || !passwordHash || !fullName) {
    return { success: false, error: "All registration fields are required." };
  }
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === username) {
      return { success: false, error: "Username is already taken. Please choose another." };
    }
  }
  
  var userId = Utilities.getUuid();
  var now = new Date().toISOString();
  var token = Utilities.base64EncodeWebSafe(userId + "::" + now);
  
  sheet.appendRow([
    userId,
    username,
    passwordHash,
    fullName,
    now,
    now,
    0,
    "ACTIVE"
  ]);
  
  return {
    success: true,
    data: {
      userId: userId,
      username: username,
      fullName: fullName,
      token: token
    }
  };
}

/**
 * Authenticate User Controller
 */
function handleAuthenticateUser(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  var rows = sheet.getDataRange().getValues();
  
  var username = (data.username || "").toLowerCase().trim();
  var passwordHash = (data.passwordHash || "").trim();
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === username) {
      if (rows[i][7] !== "ACTIVE") {
        return { success: false, error: "Account has been suspended or deactivated." };
      }
      if (rows[i][2] === passwordHash) {
        var now = new Date().toISOString();
        sheet.getRange(i + 1, 6).setValue(now);
        var userId = rows[i][0];
        var fullName = rows[i][3];
        var token = Utilities.base64EncodeWebSafe(userId + "::" + now);
        
        return {
          success: true,
          data: {
            userId: userId,
            username: username,
            fullName: fullName,
            totalUploads: rows[i][6],
            token: token
          }
        };
      } else {
        return { success: false, error: "Invalid password credentials." };
      }
    }
  }
  return { success: false, error: "User account not found." };
}

/**
 * Audio Upload and FIFO Pruning Engine (Cap: 50 active items)
 */
function handleUploadAudioTrack(data) {
  var auth = data.auth || {};
  var metadata = data.metadata || {};
  var fileBase64 = data.fileData;
  
  if (!auth.username || !fileBase64 || !metadata.mimeType) {
    return { success: false, error: "Missing required binary payload or parameters." };
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  var userRows = usersSheet.getDataRange().getValues();
  var authenticatedUser = null;
  var userRowIndex = -1;
  
  for (var u = 1; u < userRows.length; u++) {
    if (userRows[u][1].toString().toLowerCase() === auth.username.toLowerCase()) {
      authenticatedUser = {
        id: userRows[u][0],
        username: userRows[u][1],
        fullName: userRows[u][3]
      };
      userRowIndex = u + 1;
      break;
    }
  }
  
  if (!authenticatedUser) {
    return { success: false, error: "Authentication validation failed." };
  }
  
  var manifestSheet = ss.getSheetByName(CONFIG.MANIFEST_SHEET);
  var manifestRows = manifestSheet.getDataRange().getValues();
  
  // Enforce FIFO Limit to 50 active audio tracks
  var activeTracks = [];
  for (var m = 1; m < manifestRows.length; m++) {
    if (manifestRows[m][11] === "ACTIVE") {
      activeTracks.push({
        rowIndex: m + 1,
        contentId: manifestRows[m][0],
        fileId: manifestRows[m][1],
        uploadTimestamp: new Date(manifestRows[m][9]).getTime()
      });
    }
  }
  
  activeTracks.sort(function(a, b) {
    return a.uploadTimestamp - b.uploadTimestamp;
  });
  
  while (activeTracks.length >= CONFIG.MAX_ACTIVE_TRACKS) {
    var oldestTrack = activeTracks.shift();
    try {
      var oldFile = DriveApp.getFileById(oldestTrack.fileId);
      oldFile.setTrashed(true);
    } catch (e) {
      Logger.log("Drive file purge error or already deleted: " + e.toString());
    }
    manifestSheet.getRange(oldestTrack.rowIndex, 12).setValue("PURGED_FIFO");
  }
  
  // Ingest New Audio
  var decodedBytes = Utilities.base64Decode(fileBase64);
  var sanitizedTitle = (metadata.title || "Untitled Transmission").substring(0, 80);
  var diskFileName = "PULSE_" + Date.now() + "_" + (metadata.fileName || "audio.webm");
  var blob = Utilities.newBlob(decodedBytes, metadata.mimeType, diskFileName);
  
  var targetFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  var newDriveFile = targetFolder.createFile(blob);
  newDriveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  var fileId = newDriveFile.getId();
  var directStreamUrl = "https://drive.google.com/uc?export=download&id=" + fileId;
  var contentId = Utilities.getUuid();
  var nowIso = new Date().toISOString();
  
  // Record to Manifest Sheet
  manifestSheet.appendRow([
    contentId,
    fileId,
    diskFileName,
    sanitizedTitle,
    authenticatedUser.username,
    authenticatedUser.fullName,
    metadata.mimeType,
    metadata.sizeBytes || 0,
    metadata.durationSeconds || 0,
    nowIso,
    directStreamUrl,
    "ACTIVE",
    0
  ]);
  
  var currentUploadCount = parseInt(usersSheet.getRange(userRowIndex, 7).getValue(), 10) || 0;
  usersSheet.getRange(userRowIndex, 7).setValue(currentUploadCount + 1);
  
  return {
    success: true,
    data: {
      contentId: contentId,
      fileId: fileId,
      title: sanitizedTitle,
      playbackUrl: directStreamUrl,
      uploadTimestamp: nowIso
    }
  };
}

/**
 * Handle Push-To-Talk Live Voiceover Broadcast
 */
function handleBroadcastLiveVoiceover(data) {
  var auth = data.auth || {};
  var voiceBase64 = data.voiceData;
  var mimeType = data.mimeType || "audio/webm";
  
  if (!auth.username || !voiceBase64) {
    return { success: false, error: "Missing audio voice data or credentials." };
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  var userRows = usersSheet.getDataRange().getValues();
  var uploaderName = auth.username;
  
  for (var u = 1; u < userRows.length; u++) {
    if (userRows[u][1].toString().toLowerCase() === auth.username.toLowerCase()) {
      uploaderName = userRows[u][3];
      break;
    }
  }
  
  var voiceId = Utilities.getUuid();
  var timestamp = Date.now();
  var voiceoverPayload = {
    voiceId: voiceId,
    timestamp: timestamp,
    uploaderUsername: auth.username,
    uploaderName: uploaderName,
    mimeType: mimeType,
    base64: voiceBase64
  };
  
  var cache = CacheService.getScriptCache();
  cache.put("ACTIVE_LIVE_VOICEOVER", JSON.stringify(voiceoverPayload), 45); // Expires in 45 seconds
  
  return {
    success: true,
    data: {
      voiceId: voiceId,
      timestamp: timestamp,
      status: "BROADCASTING_LIVE"
    }
  };
}

/**
 * Record Listener Heartbeat in Cache & Compute Active Listener Count
 */
function recordListenerHeartbeat(listenerId) {
  if (!listenerId) return 1;
  var cache = CacheService.getScriptCache();
  var now = Date.now();
  var cachedRaw = cache.get("ACTIVE_LISTENERS_MAP");
  var listenersMap = {};
  
  if (cachedRaw) {
    try {
      listenersMap = JSON.parse(cachedRaw);
    } catch (e) {
      listenersMap = {};
    }
  }
  
  listenersMap[listenerId] = now;
  
  var activeCount = 0;
  var prunedMap = {};
  var threshold = now - 30000; // 30-second window
  
  for (var id in listenersMap) {
    if (listenersMap.hasOwnProperty(id)) {
      if (listenersMap[id] > threshold) {
        prunedMap[id] = listenersMap[id];
        activeCount++;
      }
    }
  }
  
  cache.put("ACTIVE_LISTENERS_MAP", JSON.stringify(prunedMap), 60);
  return Math.max(1, activeCount);
}

/**
 * Manifest Polling Query (Newest-First with Integrity Checks, Listener Count & Live Voiceover)
 */
function handleGetLiveManifest(listenerId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var manifestSheet = ss.getSheetByName(CONFIG.MANIFEST_SHEET);
  var rows = manifestSheet.getDataRange().getValues();
  var activeList = [];
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][11] === "ACTIVE") {
      var fileId = rows[i][1];
      var fileExists = true;
      
      try {
        var driveFile = DriveApp.getFileById(fileId);
        if (driveFile.isTrashed()) fileExists = false;
      } catch (err) {
        fileExists = false;
      }
      
      if (!fileExists) {
        manifestSheet.getRange(i + 1, 12).setValue("MISSING_OR_DELETED");
        continue;
      }
      
      activeList.push({
        contentId: rows[i][0],
        fileId: fileId,
        fileName: rows[i][2],
        title: rows[i][3],
        uploaderUsername: rows[i][4],
        uploaderName: rows[i][5],
        mimeType: rows[i][6],
        fileSizeBytes: rows[i][7],
        durationSeconds: rows[i][8],
        uploadTimestamp: rows[i][9],
        playbackDirectUrl: rows[i][10],
        playCount: rows[i][12]
      });
    }
  }
  
  activeList.sort(function(a, b) {
    return new Date(b.uploadTimestamp).getTime() - new Date(a.uploadTimestamp).getTime();
  });
  
  var activeListeners = recordListenerHeartbeat(listenerId);
  
  var cache = CacheService.getScriptCache();
  var liveVoiceoverRaw = cache.get("ACTIVE_LIVE_VOICEOVER");
  var liveVoiceover = null;
  if (liveVoiceoverRaw) {
    try {
      liveVoiceover = JSON.parse(liveVoiceoverRaw);
    } catch (e) {
      liveVoiceover = null;
    }
  }
  
  return {
    success: true,
    data: {
      manifest: activeList,
      totalActive: activeList.length,
      activeListeners: activeListeners,
      liveVoiceover: liveVoiceover,
      serverTime: new Date().toISOString()
    }
  };
}

/**
 * Buffer Status Query
 */
function handleGetBufferStats() {
  var manifestResult = handleGetLiveManifest(null);
  return {
    success: true,
    data: {
      activeSlotsUsed: manifestResult.data.totalActive,
      maxSlots: CONFIG.MAX_ACTIVE_TRACKS
    }
  };
}