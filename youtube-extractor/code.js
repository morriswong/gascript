/**
 * YouTube Channel Video Extractor - Multi-Channel Support
 * Extracts all videos from YouTube channels, each channel gets its own sheet
 */

// Configuration - adjust thumbnail size here
const THUMBNAIL_WIDTH = 120;
const THUMBNAIL_HEIGHT = 68;
const ROW_HEIGHT = 70;

/**
 * Runs when the add-on is installed
 */
function onInstall(e) {
  onOpen(e);
}

// Add custom menu when spreadsheet opens
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('YouTube Extractor')
    .addItem('➕ Add New Channel', 'addNewChannel')
    .addSeparator()
    .addItem('🔄 Load Latest Videos (Current Sheet)', 'loadLatestVideos')
    .addItem('🔄 Refresh All Channels', 'refreshAllChannels')
    .addSeparator()
    .addItem('🗑️ Remove Current Channel', 'removeCurrentChannel')
    .addSeparator()
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'YouTube Channel Extractor',
    'Version 1.0\n\nExtract and track videos from YouTube channels.\n\nDeveloped by [Your Name]',
    ui.ButtonSet.OK
  );
}

/**
 * Add a new YouTube channel (creates a new sheet)
 */
function addNewChannel() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Prompt for channel URL
  const response = ui.prompt(
    'Add New YouTube Channel',
    'Enter the YouTube channel URL, handle (@name), or Channel ID:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const input = response.getResponseText().trim();
  if (!input) {
    ui.alert('Error', 'Please enter a valid YouTube channel URL or ID.', ui.ButtonSet.OK);
    return;
  }

  try {
    spreadsheet.toast('Fetching channel information...', 'Please wait', -1);

    const channelId = getChannelId(input);
    if (!channelId) {
      ui.alert('Error', 'Could not find channel. Please check the URL or ID.', ui.ButtonSet.OK);
      return;
    }

    if (isChannelAlreadyAdded(channelId)) {
      ui.alert('Info', 'This channel has already been added. Use "Load Latest Videos" to update it.', ui.ButtonSet.OK);
      return;
    }

    const channelInfo = getChannelInfo(channelId);
    if (!channelInfo) {
      ui.alert('Error', 'Could not find videos for this channel.', ui.ButtonSet.OK);
      return;
    }

    const sheetName = getUniqueSheetName(spreadsheet, channelInfo.channelName);
    const newSheet = spreadsheet.insertSheet(sheetName);

    setupHeaders(newSheet, channelId, channelInfo.channelName);

    spreadsheet.toast('Fetching videos...', 'Please wait', -1);
    const videos = getAllVideosFromPlaylist(channelInfo.uploadsPlaylistId);

    if (videos.length === 0) {
      ui.alert('Info', `Sheet created for "${channelInfo.channelName}" but no videos found.`, ui.ButtonSet.OK);
      return;
    }

    spreadsheet.toast('Fetching video details...', 'Please wait', -1);
    const videoDetails = getVideoDetails(videos);

    writeVideosToSheet(newSheet, videoDetails);
    spreadsheet.setActiveSheet(newSheet);

    spreadsheet.toast(`Successfully added "${channelInfo.channelName}" with ${videoDetails.length} videos!`, 'Complete', 5);

  } catch (error) {
    ui.alert('Error', `An error occurred: ${error.message}`, ui.ButtonSet.OK);
    console.error(error);
  }
}

/**
 * Load only the latest videos for the current sheet
 */
function loadLatestVideos() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  try {
    const channelId = getStoredChannelId(sheet);

    if (!channelId) {
      ui.alert('Error', 'This sheet is not linked to a YouTube channel.\n\nUse "Add New Channel" to add a channel.', ui.ButtonSet.OK);
      return;
    }

    const result = refreshSheet(sheet, channelId);

    if (result.error) {
      ui.alert('Error', result.error, ui.ButtonSet.OK);
    } else if (result.newCount === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`"${result.channelName}" is up to date!`, 'Complete', 5);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Added ${result.newCount} new video(s) to "${result.channelName}"!`, 'Complete', 5);
    }

  } catch (error) {
    ui.alert('Error', `An error occurred: ${error.message}`, ui.ButtonSet.OK);
    console.error(error);
  }
}

/**
 * Refresh all channel sheets
 */
function refreshAllChannels() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  const channelSheets = [];

  for (const sheet of sheets) {
    const channelId = getStoredChannelId(sheet);
    if (channelId) {
      channelSheets.push({ sheet: sheet, channelId: channelId });
    }
  }

  if (channelSheets.length === 0) {
    ui.alert('Info', 'No YouTube channels found. Use "Add New Channel" to add channels.', ui.ButtonSet.OK);
    return;
  }

  const results = [];
  let totalNew = 0;

  for (let i = 0; i < channelSheets.length; i++) {
    const { sheet, channelId } = channelSheets[i];
    spreadsheet.toast(`Refreshing ${i + 1}/${channelSheets.length}: ${sheet.getName()}...`, 'Please wait', -1);

    const result = refreshSheet(sheet, channelId);
    results.push(result);

    if (!result.error) {
      totalNew += result.newCount;
    }
  }

  let summary = `Refreshed ${channelSheets.length} channel(s):\n\n`;
  for (const result of results) {
    if (result.error) {
      summary += `❌ ${result.channelName || 'Unknown'}: ${result.error}\n`;
    } else if (result.newCount > 0) {
      summary += `✅ ${result.channelName}: +${result.newCount} new video(s)\n`;
    } else {
      summary += `✓ ${result.channelName}: Up to date\n`;
    }
  }

  spreadsheet.toast(`Refresh complete! ${totalNew} new video(s) total.`, 'Complete', 5);
  ui.alert('Refresh Complete', summary, ui.ButtonSet.OK);
}

/**
 * Refresh a single sheet and return results
 */
function refreshSheet(sheet, channelId) {
  try {
    const channelInfo = getChannelInfo(channelId);
    if (!channelInfo) {
      return { error: 'Could not access channel information.', channelName: sheet.getName() };
    }

    const mostRecentVideoId = getMostRecentVideoId(sheet);
    const newVideos = getNewVideosFromPlaylist(channelInfo.uploadsPlaylistId, mostRecentVideoId);

    if (newVideos.length === 0) {
      return { newCount: 0, channelName: channelInfo.channelName };
    }

    const videoDetails = getVideoDetails(newVideos);
    insertNewVideosAtTop(sheet, videoDetails);

    return { newCount: videoDetails.length, channelName: channelInfo.channelName };

  } catch (error) {
    return { error: error.message, channelName: sheet.getName() };
  }
}

/**
 * Remove the current channel (delete sheet)
 */
function removeCurrentChannel() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  const channelId = getStoredChannelId(sheet);

  if (!channelId) {
    ui.alert('Error', 'This sheet is not linked to a YouTube channel.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    'Confirm Removal',
    `Are you sure you want to remove "${sheet.getName()}" and all its data?\n\nThis cannot be undone.`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteProperty('channelId_' + sheet.getSheetId());
  documentProperties.deleteProperty('channelName_' + sheet.getSheetId());

  spreadsheet.deleteSheet(sheet);
  spreadsheet.toast('Channel removed successfully.', 'Complete', 3);
}

/**
 * Check if a channel has already been added
 */
function isChannelAlreadyAdded(channelId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  for (const sheet of sheets) {
    const storedId = getStoredChannelId(sheet);
    if (storedId === channelId) {
      return true;
    }
  }
  return false;
}

/**
 * Get a unique sheet name (append number if duplicate)
 */
function getUniqueSheetName(spreadsheet, baseName) {
  const sanitized = sanitizeSheetName(baseName);
  const sheets = spreadsheet.getSheets();
  const existingNames = sheets.map(s => s.getName().toLowerCase());

  if (!existingNames.includes(sanitized.toLowerCase())) {
    return sanitized;
  }

  let counter = 2;
  while (existingNames.includes(`${sanitized} (${counter})`.toLowerCase())) {
    counter++;
  }
  return `${sanitized} (${counter})`;
}

/**
 * Get channel info including name and uploads playlist ID
 */
function getChannelInfo(channelId) {
  const response = YouTube.Channels.list('snippet,contentDetails', {
    id: channelId
  });

  if (response.items && response.items.length > 0) {
    return {
      channelId: channelId,
      channelName: response.items[0].snippet.title,
      uploadsPlaylistId: response.items[0].contentDetails.relatedPlaylists.uploads
    };
  }
  return null;
}

/**
 * Sanitize sheet name (remove invalid characters and limit length)
 */
function sanitizeSheetName(name) {
  let sanitized = name.replace(/[\\\/\?\*\[\]:]/g, '');
  if (sanitized.length > 100) {
    sanitized = sanitized.substring(0, 100);
  }
  return sanitized || 'YouTube Channel';
}

/**
 * Extract channel ID from various URL formats
 */
function getChannelId(input) {
  if (input.match(/^UC[\w-]{22}$/)) {
    return input;
  }

  let match = input.match(/youtube\.com\/channel\/(UC[\w-]{22})/);
  if (match) {
    return match[1];
  }

  match = input.match(/youtube\.com\/(@[\w.-]+|c\/[\w.-]+|user\/[\w.-]+)/);
  if (match) {
    const identifier = match[1];

    if (identifier.startsWith('@')) {
      const handle = identifier.substring(1);
      const response = YouTube.Channels.list('id', {
        forHandle: handle
      });
      if (response.items && response.items.length > 0) {
        return response.items[0].id;
      }
    } else {
      const searchTerm = identifier.replace(/^(c\/|user\/)/, '');
      const searchResponse = YouTube.Search.list('snippet', {
        q: searchTerm,
        type: 'channel',
        maxResults: 1
      });
      if (searchResponse.items && searchResponse.items.length > 0) {
        return searchResponse.items[0].snippet.channelId;
      }
    }
  }

  if (input.startsWith('@')) {
    const response = YouTube.Channels.list('id', {
      forHandle: input.substring(1)
    });
    if (response.items && response.items.length > 0) {
      return response.items[0].id;
    }
  }

  const searchResponse = YouTube.Search.list('snippet', {
    q: input,
    type: 'channel',
    maxResults: 1
  });
  if (searchResponse.items && searchResponse.items.length > 0) {
    return searchResponse.items[0].snippet.channelId;
  }

  return null;
}

/**
 * Get all videos from a playlist (handles pagination)
 */
function getAllVideosFromPlaylist(playlistId) {
  const videos = [];
  let nextPageToken = null;

  do {
    const options = {
      playlistId: playlistId,
      maxResults: 50
    };

    if (nextPageToken) {
      options.pageToken = nextPageToken;
    }

    const response = YouTube.PlaylistItems.list('snippet,contentDetails', options);

    if (response.items) {
      response.items.forEach(item => {
        // Skip deleted or private videos
        const title = item.snippet?.title;
        if (title === 'Deleted video' || title === 'Private video') {
          return;
        }

        // Skip if no video ID
        const videoId = item.contentDetails?.videoId;
        if (!videoId) {
          return;
        }

        videos.push({
          videoId: videoId,
          title: title || 'Untitled',
          description: item.snippet?.description || '',
          publishedAt: item.snippet?.publishedAt || new Date().toISOString(),
          thumbnailUrl: item.snippet?.thumbnails?.medium?.url || item.snippet?.thumbnails?.default?.url || ''
        });
      });
    }

    nextPageToken = response.nextPageToken;
    SpreadsheetApp.getActiveSpreadsheet().toast(`Fetched ${videos.length} videos...`, 'Please wait', -1);

  } while (nextPageToken);

  return videos;
}

/**
 * Get only new videos from playlist (stops when it finds an existing video)
 */
function getNewVideosFromPlaylist(playlistId, mostRecentVideoId) {
  const videos = [];
  let nextPageToken = null;
  let foundExisting = false;

  do {
    const options = {
      playlistId: playlistId,
      maxResults: 50
    };

    if (nextPageToken) {
      options.pageToken = nextPageToken;
    }

    const response = YouTube.PlaylistItems.list('snippet,contentDetails', options);

    if (response.items) {
      for (const item of response.items) {
        // Skip deleted or private videos
        const title = item.snippet?.title;
        if (title === 'Deleted video' || title === 'Private video') {
          continue;
        }

        const videoId = item.contentDetails?.videoId;
        if (!videoId) {
          continue;
        }

        if (videoId === mostRecentVideoId) {
          foundExisting = true;
          break;
        }

        videos.push({
          videoId: videoId,
          title: title || 'Untitled',
          description: item.snippet?.description || '',
          publishedAt: item.snippet?.publishedAt || new Date().toISOString(),
          thumbnailUrl: item.snippet?.thumbnails?.medium?.url || item.snippet?.thumbnails?.default?.url || ''
        });
      }
    }

    if (foundExisting) {
      break;
    }

    nextPageToken = response.nextPageToken;

  } while (nextPageToken);

  return videos;
}

/**
 * Get detailed statistics for videos (in batches of 50)
 */
function getVideoDetails(videos) {
  const detailedVideos = [];

  for (let i = 0; i < videos.length; i += 50) {
    const batch = videos.slice(i, i + 50);
    const videoIds = batch.map(v => v.videoId).join(',');

    try {
      const response = YouTube.Videos.list('statistics,contentDetails', {
        id: videoIds
      });

      if (response.items) {
        response.items.forEach((item) => {
          const originalVideo = batch.find(v => v.videoId === item.id);
          if (originalVideo) {
            detailedVideos.push({
              ...originalVideo,
              viewCount: parseInt(item.statistics?.viewCount || 0),
              likeCount: parseInt(item.statistics?.likeCount || 0),
              commentCount: parseInt(item.statistics?.commentCount || 0),
              duration: parseDuration(item.contentDetails?.duration)
            });
          }
        });
      }

      // Add videos that weren't found (deleted/private) with default values
      batch.forEach(video => {
        const found = detailedVideos.find(v => v.videoId === video.videoId);
        if (!found) {
          detailedVideos.push({
            ...video,
            viewCount: 0,
            likeCount: 0,
            commentCount: 0,
            duration: '0:00'
          });
        }
      });

    } catch (error) {
      console.error('Error fetching video details:', error);
      // Add batch with default values if API fails
      batch.forEach(video => {
        detailedVideos.push({
          ...video,
          viewCount: 0,
          likeCount: 0,
          commentCount: 0,
          duration: '0:00'
        });
      });
    }
  }

  return detailedVideos;
}

/**
 * Parse ISO 8601 duration to readable format
 */
function parseDuration(duration) {
  if (!duration) return '0:00';

  const match = duration.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!match) return '0:00';

  const hours = parseInt(match[1] || 0);
  const minutes = parseInt(match[2] || 0);
  const seconds = parseInt(match[3] || 0);

  if (hours > 0) {
    return `${hours}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
  }
  return `${minutes}:${seconds.toString().padStart(2, '0')}`;
}

/**
 * Build IMAGE formula for thumbnail
 */
function buildImageFormula(url) {
  if (!url) return '';
  // IMAGE(url, mode, height, width) - mode 4 = custom size
  return `=IMAGE("${url}", 4, ${THUMBNAIL_HEIGHT}, ${THUMBNAIL_WIDTH})`;
}

/**
 * Setup header row and store channel ID
 */
function setupHeaders(sheet, channelId, channelName) {
  const headers = [
    'Thumbnail',
    'Video Title',
    'Video URL',
    'Published Date',
    'Duration',
    'Views',
    'Likes',
    'Comments',
    'Description'
  ];

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#cc0000')
    .setFontColor('white')
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  // Set thumbnail column width
  sheet.setColumnWidth(1, THUMBNAIL_WIDTH + 10);

  // Set header row height
  sheet.setRowHeight(1, 30);

  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('channelId_' + sheet.getSheetId(), channelId);
  documentProperties.setProperty('channelName_' + sheet.getSheetId(), channelName);
}

/**
 * Get stored channel ID for the current sheet
 */
function getStoredChannelId(sheet) {
  const documentProperties = PropertiesService.getDocumentProperties();
  return documentProperties.getProperty('channelId_' + sheet.getSheetId());
}

/**
 * Get the most recent video ID from the sheet
 */
function getMostRecentVideoId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  // Video URL is now in column 3
  const url = sheet.getRange(2, 3).getValue();
  const match = url.toString().match(/[?&]v=([\w-]{11})/);
  return match ? match[1] : null;
}

/**
 * Insert new videos at the top of the sheet (below headers)
 */
function insertNewVideosAtTop(sheet, videos) {
  if (videos.length === 0) return;

  const data = videos.map(video => [
    buildImageFormula(video.thumbnailUrl),
    video.title,
    `https://www.youtube.com/watch?v=${video.videoId}`,
    new Date(video.publishedAt),
    video.duration,
    video.viewCount,
    video.likeCount,
    video.commentCount,
    video.description.substring(0, 500)
  ]);

  // Insert new rows below the header
  sheet.insertRowsAfter(1, data.length);
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

  // Set row heights for thumbnails
  for (let i = 0; i < data.length; i++) {
    sheet.setRowHeight(i + 2, ROW_HEIGHT);
  }

  // Format the new rows
  sheet.getRange(2, 6, data.length, 3).setNumberFormat('#,##0');
  sheet.getRange(2, 4, data.length, 1).setNumberFormat('yyyy-mm-dd');

  // Center align thumbnail column
  sheet.getRange(2, 1, data.length, 1).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Highlight new videos with light green
  sheet.getRange(2, 1, data.length, data[0].length).setBackground('#e6ffe6');
}

/**
 * Write video data to sheet
 */
function writeVideosToSheet(sheet, videos) {
  if (videos.length === 0) return;

  const data = videos.map(video => [
    buildImageFormula(video.thumbnailUrl),
    video.title,
    `https://www.youtube.com/watch?v=${video.videoId}`,
    new Date(video.publishedAt),
    video.duration,
    video.viewCount,
    video.likeCount,
    video.commentCount,
    video.description.substring(0, 500)
  ]);

  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

  // Set row heights for all data rows
  for (let i = 0; i < data.length; i++) {
    sheet.setRowHeight(i + 2, ROW_HEIGHT);
  }

  // Set column widths
  sheet.setColumnWidth(1, THUMBNAIL_WIDTH + 10);  // Thumbnail
  sheet.setColumnWidth(2, 300);                    // Title
  sheet.setColumnWidth(3, 280);                    // URL
  sheet.setColumnWidth(4, 100);                    // Date
  sheet.setColumnWidth(5, 70);                     // Duration
  sheet.setColumnWidth(6, 90);                     // Views
  sheet.setColumnWidth(7, 70);                     // Likes
  sheet.setColumnWidth(8, 80);                     // Comments
  sheet.setColumnWidth(9, 300);                    // Description

  // Format numbers
  sheet.getRange(2, 6, data.length, 3).setNumberFormat('#,##0');

  // Format date
  sheet.getRange(2, 4, data.length, 1).setNumberFormat('yyyy-mm-dd');

  // Center align thumbnail column
  sheet.getRange(2, 1, data.length, 1).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Vertical align all cells
  sheet.getRange(2, 2, data.length, data[0].length - 1).setVerticalAlignment('middle');
}
