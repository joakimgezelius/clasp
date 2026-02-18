/**
 * CONFIGURATION
 */
const ORG_UNIT_ID = 'orgunits/my_customer'; // Change to your specific OU ID if needed
const TOP_LEVEL_NAME = 'Company Bookmarks'; 
const PATH_SEPARATOR = ' > '; // The separator used in Column A (e.g. "HR > Benefits")

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Chrome Admin')
    .addItem('Preview Structure', 'previewPayload')
    .addSeparator()
    .addItem('🚀 Push to Chrome', 'pushToChrome')
    .addToUi();
}

/**
 * Reads the sheet and constructs the deep nested JSON structure.
 */
function buildBookmarksStructure() {
  const sheet = SpreadsheetApp.getActiveSheet();
  // Get all data (skip header row)
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  
  // The 'root' array that will hold our top-level items
  let rootItems = [];

  data.forEach(row => {
    let pathString = row[0].toString().trim();
    let name = row[1].toString().trim();
    let url = row[2].toString().trim();

    // Skip rows that don't have at least a Name and URL
    if (!name || !url) return;

    // Create the bookmark object
    let newBookmark = { "name": name, "url": url };

    // If no path, push to root
    if (!pathString) {
      rootItems.push(newBookmark);
      return;
    }

    // If there is a path, we must traverse/build the tree
    let pathParts = pathString.split(PATH_SEPARATOR);
    let currentLevel = rootItems;

    pathParts.forEach((folderName) => {
      // Look for an existing folder with this name in the current level
      let existingFolder = currentLevel.find(item => item.name === folderName && item.children);

      if (existingFolder) {
        // Folder exists, step into it
        currentLevel = existingFolder.children;
      } else {
        // Folder doesn't exist, create it
        let newFolder = { "name": folderName, "children": [] };
        currentLevel.push(newFolder);
        // Step into the new folder
        currentLevel = newFolder.children;
      }
    });

    // We are now inside the deepest folder, drop the bookmark here
    currentLevel.push(newBookmark);
  });

  // Wrap it in the official Chrome Policy envelope
  return [
    { "toplevel_name": TOP_LEVEL_NAME }, 
    { "top": rootItems }
  ];
}

/**
 * DEBUG: Shows the JSON structure in a popup so you can verify nesting.
 */
function previewPayload() {
  const structure = buildBookmarksStructure();
  const json = JSON.stringify(structure, null, 2);
  console.log(json); // Also logs to View > Execution Transcript
  SpreadsheetApp.getUi().alert("Check View > Execution Transcript for full log.\n\nPreview of start:\n" + json.substring(0, 500) + "...");
}

/**
 * ACTION: Pushes the structure to the Chrome Policy API.
 */
function pushToChrome() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
     '⚠️ Confirm Update',
     'This will overwrite all Managed Bookmarks for OU: ' + ORG_UNIT_ID + '.\nAre you sure?',
      ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  const bookmarksJson = buildBookmarksStructure();

  const payload = {
    "requests": [
      {
        "policyTargetKey": { "targetResource": ORG_UNIT_ID },
        "policyValue": {
          "policySchema": "chrome.users.ManagedBookmarks",
          "value": {
            "managedBookmarks": bookmarksJson
          }
        },
        "updateMask": { "paths": "managedBookmarks" }
      }
    ]
  };

  try {
    const url = 'https://chromepolicy.googleapis.com/v1/customers/my_customer/policies/orgunits/' + ORG_UNIT_ID.split('/')[1] + ':batchModify';
    const params = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, params);
    
    if (response.getResponseCode() === 200) {
      ui.alert('✅ Success! Bookmarks updated.');
    } else {
      ui.alert('❌ Error: ' + response.getContentText());
    }
  } catch (e) {
    ui.alert('Script Error: ' + e.toString());
  }
}