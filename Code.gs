//--------------------------------------------------------------------------------------------------
// BASCONF SHEET AND BACKEND
//--------------------------------------------------------------------------------------------------
function getActiveSnippet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BASECONF").getRange("C15").getValue();
}

function getValidToken() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B6").getValue();
}

function getFreshToken() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BASECONF");
  const token = sheet.getRange('B6').getValue();
  const expiryNote = sheet.getRange('B5').getValue();

  const match = expiryNote.match(/(\d+) seconds after (.+)/);
  if (!match) {
    logVlan('⚠️ Invalid expiry note format, requesting new token.');
    getSCMToken(); // force refresh
    return getFreshToken(); // retry
  }

  const issuedAt = new Date(match[2]);
  const validSeconds = parseInt(match[1]);
  const expiresAt = new Date(issuedAt.getTime() + validSeconds * 1000);
  const now = new Date();

  const timeRemaining = (expiresAt - now) / 1000;

  if (timeRemaining < 180) {
    logVlan(`🔄 Token expiring in ${Math.floor(timeRemaining)}s — refreshing.`);
    getSCMToken();
    return getFreshToken(); // retry
  }

  return {
    token,
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    }
  };
}

function getSCMToken() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BASECONF");
  const tsgId = sheet.getRange('B2').getValue().toString().trim();
  const clientId = sheet.getRange('B3').getValue().toString().trim();
  const clientSecret = sheet.getRange('B4').getValue().toString().trim();

  if (!tsgId || !clientId || !clientSecret) {
    SpreadsheetApp.getUi().alert('Please fill in TSG ID, client_id, and client_secret.');
    return;
  }

  const url = "https://auth.apps.paloaltonetworks.com/oauth2/access_token";
  const payload = `grant_type=client_credentials&scope=tsg_id:${encodeURIComponent(tsgId)}`;

  const headers = {
    "Authorization": "Basic " + Utilities.base64Encode(`${clientId}:${clientSecret}`),
    "Content-Type": "application/x-www-form-urlencoded"
  };

  const options = {
    method: "post",
    payload: payload,
    headers: headers,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  if (code !== 200) {
    SpreadsheetApp.getUi().alert("Token request failed:\n" + body);
    return;
  }

  const json = JSON.parse(body);
  const token = json.access_token;
  const expiresIn = parseInt(json.expires_in);
  const now = new Date();
  const formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const expiryNote = `${expiresIn} seconds after ${formattedTime}`;

  sheet.getRange('B6').setValue(token);
  sheet.getRange('B5').setValue(expiryNote);
}
/** Get snippet info from BASECONF */
function getBaseconfSnippetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BASECONF");
  if (!sheet) throw new Error("BASECONF sheet not found");

  const deviceName = sheet.getRange("C13").getValue();
  const serial = sheet.getRange("C14").getValue();
  const snippet = sheet.getRange("C15").getValue();

  logVlan(`📥 Read BASECONF: device=${deviceName}, serial=${serial}, snippet=${snippet}`);
  return { deviceName, serial, snippet };
}
/** Called from dialogs to suppress iframe warnings and confirm logging session */
function noopLogPing(type) {
  switch (type) {
    case 'vlan':
      logVlan("👋 noopLogPing called (VLAN)");
      break;
    case 'vrf':
      logVrf("👋 noopLogPing called (VRF)");
      break;
    case 'snatdhcp':
      logRoute("👋 noopLogPing called (SNAT-DHCP)");
      break;
    default:
      console.warn("⚠️ noopLogPing called with unknown type:", type);
      break;
  }
  return true;
}
///------------------------------------------------------------
/// VAR handling blocks
///-------------------------------------------------------------
function PullVars_From_SCM() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const debugLines = [];
  const apiBase = 'https://api.strata.paloaltonetworks.com/config/setup/v1';

  function logDebug(label, content) {
    const time = new Date().toISOString();
    debugLines.push(`=== ${label} @ ${time} ===\n${content}`);
  }

  function request(url, options, label) {
    const fullOpts = Object.assign({ muteHttpExceptions: true }, options);
    const response = UrlFetchApp.fetch(url, fullOpts);
    const status = response.getResponseCode();
    const body = response.getContentText();
    logDebug(`${label} [${status}]`, `URL: ${url}\n\n${body}`);
    return { status, body: JSON.parse(body || '{}') };
  }

  try {
    const confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BASECONF");
    const snippetName = confSheet.getRange("C15").getValue();  // use canonical source
    const { token, headers } = getFreshToken();                // use your global mechanism

    logDebug('INFO', `Looking for snippet "${snippetName}"`);
    const listResp = request(`${apiBase}/snippets`, { headers }, 'List Snippets');
    const snippets = listResp.body.data || [];
    const snippet = snippets.find(s => s.name === snippetName);

    if (!snippet) {
      ui.alert(`❌ Snippet "${snippetName}" not found in tenant.`);
      return;
    }

    logDebug('INFO', `✅ Snippet "${snippetName}" found`);
    const varResp = request(
      `${apiBase}/variables?snippet=${encodeURIComponent(snippetName)}`,
      { headers },
      'Get Snippet Variables'
    );
    const variables = varResp.body.data || [];

    logDebug('INFO', `Fetched ${variables.length} variables from snippet`);

    // Clear from row 17 downward in columns A–D
    const lastRow = sheet.getLastRow();
    if (lastRow >= 17) {
      sheet.getRange('A17:D' + lastRow).clearContent();
    }

    // Write variables to sheet starting at row 17
    const values = variables.map(v => [v.name, v.description || '', v.value, v.type]);
    if (values.length > 0) {
      sheet.getRange(17, 1, values.length, 4).setValues(values);
    }

    ui.alert(`✅ Pulled ${values.length} variables from snippet "${snippetName}"`);
  } catch (e) {
    const msg = '❌ ERROR: ' + e.message;
    logDebug('EXCEPTION', msg);
    SpreadsheetApp.getUi().alert(msg);
  } finally {
    // Show downloadable debug log
    const debugText = debugLines.join('\n\n');
    const base64 = Utilities.base64Encode(debugText, Utilities.Charset.UTF_8);
    const dataUrl = `data:text/plain;charset=utf-8;base64,${base64}`;
    const html = HtmlService.createHtmlOutput(
      `<p><b>✅ Debug log ready for download:</b></p>
       <a href="${dataUrl}" download="PullVars_DebugLog.txt">Click here to download log file</a>`
    ).setWidth(500).setHeight(200);
    ui.showModalDialog(html, 'PullVars_From_SCM Log Download');
  }
}
/// ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---
function PushVars_2_SCM() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const logCell = sheet.getRange('D3');
  const debugLines = [];
  const apiBase = 'https://api.strata.paloaltonetworks.com/config/setup/v1';

  debugLines.push(`=== START @ ${new Date().toISOString()} ===`);

  function logDebug(label, content) {
    const time = new Date().toISOString();
    debugLines.push(`=== ${label} @ ${time} ===\n${content}`);
  }

  function request(url, options, label) {
    const fullOpts = Object.assign({ muteHttpExceptions: true }, options);
    const response = UrlFetchApp.fetch(url, fullOpts);
    const status = response.getResponseCode();
    const body = response.getContentText();
    logDebug(`${label} [${status}]`, `URL: ${url}\n\n${body}`);
    return { status, body: JSON.parse(body || '{}') };
  }

  try {
    logCell.setValue('Starting...');

    // === Step 1: Initialization ===
    const baseconf = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BASECONF");
    const deviceName = baseconf.getRange('C13').getValue();
    const deviceSerial = baseconf.getRange('C14').getDisplayValue().trim();
    const snippetName = baseconf.getRange('C15').getValue();

    const { headers } = getFreshToken(); // uses updated logic

    // === Step 2: List all snippets from tenant ===
    logDebug('INFO', 'Listing all snippets from tenant');
    const listResp = request(`${apiBase}/snippets`, { headers }, 'List Snippets');
    const snippetList = listResp.body.data || [];
    logDebug('DATA', JSON.stringify(snippetList, null, 2));

    // === Step 3: Check if named snippet exists ===
    const foundSnippet = snippetList.find(s => s.name === snippetName);
    let snippet_variables = [];
    let snippetWasCreated = false;

    if (foundSnippet) {
      logDebug('INFO', `✅ Snippet "${snippetName}" found`);
      const varResp = request(`${apiBase}/variables?snippet=${encodeURIComponent(snippetName)}`, { headers }, 'Get Snippet Variables');
      snippet_variables = varResp.body.data || [];
      logDebug('INFO', `Fetched ${snippet_variables.length} variables from snippet`);
    } else {
      logDebug('INFO', `Snippet "${snippetName}" not found, creating it`);
      const createResp = request(`${apiBase}/snippets`, {
        method: 'POST',
        headers,
        payload: JSON.stringify({
          name: snippetName,
          description: 'Created by macro'
        })
      }, 'Create Snippet');

      if (createResp.status === 201) {
        logDebug('INFO', `✅ Snippet "${snippetName}" created successfully`);
        snippetWasCreated = true;
        snippet_variables = [];
      } else {
        throw new Error(`❌ Failed to create snippet "${snippetName}"`);
      }
    }

    // === Step 4: Read variables from sheet ===
    const varRows = sheet.getRange('A17:D' + sheet.getLastRow()).getValues();
    const sheet_variables = [];

    for (const [name, description, value, type ] of varRows) {
      if (!name) break;
      sheet_variables.push({ name, description, value, type });
    }

    logDebug('INFO', `Parsed ${sheet_variables.length} variables from sheet`);

    // === Step 5: Compare and Sync Variables ===
    const summary = { created: [], updated: [], deleted: [] };
    const currentMap = {};
    const seenNames = {};

    snippet_variables.forEach(v => {
      if (v.snippet !== snippetName) return;
      if (!v.name || !v.id) return;

      const trimmedName = v.name.trim();
      if (seenNames[trimmedName]) {
        logDebug('WARN', `Duplicate variable name detected: "${trimmedName}"`);
        return;
      }
      seenNames[trimmedName] = true;
      currentMap[trimmedName] = v;
    });

    const varsPushedMap = {};

    for (const v of sheet_variables) {
      const cleanName = v.name.trim();
      varsPushedMap[cleanName] = v;
      const old = currentMap[cleanName];

      const payload = {
        name: v.name,
        type: v.type,
        value: v.value,
        snippet: snippetName,
        description: v.description || ""
      };

      const changed = old && (String(old.value) !== String(v.value) || String(old.type) !== String(v.type));

      if (!old) {
        request(`${apiBase}/variables`, {
          method: 'POST',
          headers,
          payload: JSON.stringify(payload)
        }, `Create Var ${v.name}`);
        summary.created.push(v.name);
      } else if (changed) {
        if (!old.id) {
          logDebug('WARN', `Cannot update ${v.name}: missing Variable ID`);
        } else {
          request(`${apiBase}/variables/${old.id}`, {
            method: 'PUT',
            headers,
            payload: JSON.stringify(payload)
          }, `Update Var ${v.name}`);
          summary.updated.push(v.name);
        }
      }

      delete currentMap[v.name]; // Mark handled
    }

    // Delete obsolete vars
    for (const name in currentMap) {
      const variableToDelete = currentMap[name];
      if (variableToDelete?.id) {
        request(`${apiBase}/variables/${variableToDelete.id}`, {
          method: 'DELETE',
          headers
        }, `Delete Var ${name}`);
      } else {
        logDebug('WARN', `Cannot delete ${name}: missing Variable ID`);
      }
      summary.deleted.push(name);
    }

    // === Result ===
    let resultMsg;
    if (
      summary.created.length === 0 &&
      summary.updated.length === 0 &&
      summary.deleted.length === 0
    ) {
      resultMsg = `✅ Snippet: ${snippetName}\nDevice: ${deviceName} (${deviceSerial})\n\nNothing changed — snippet already up to date.`;
    } else {
      resultMsg = formatSummaryLog(snippetName, deviceName, deviceSerial, summary, varsPushedMap);
    }

    logCell.setValue('✅ Done');
    ui.alert('PushVars_2_SCM Summary', resultMsg, ui.ButtonSet.OK);

  } catch (e) {
    const msg = '❌ ERROR: ' + e.message;
    logCell.setValue(msg);
    logDebug('EXCEPTION', msg);
    ui.alert('Error', msg, ui.ButtonSet.OK);
  } finally {
    const debugText = debugLines.join('\n\n');
    const base64 = Utilities.base64Encode(debugText, Utilities.Charset.UTF_8);
    const dataUrl = `data:text/plain;charset=utf-8;base64,${base64}`;
    const html = HtmlService.createHtmlOutput(
      `<p><b>✅ Debug log ready for download:</b></p>
       <a href="${dataUrl}" download="PushVars_DebugLog.txt">Click here to download log file</a>`
    ).setWidth(500).setHeight(200);
    ui.showModalDialog(html, 'PushVars_2_SCM Log Download');
  }
}

function formatSummaryLog(snippetName, deviceName, deviceSerial, summary, varsPushedMap) {
  function formatGroup(title, names) {
    if (names.length === 0) return `${title}: none`;
    return `${title}:\n` + names.map(n => {
      const v = varsPushedMap[n];
      return `  - ${n} = "${v?.value}" (${v?.type})`;
    }).join('\n');
  }

  return (
    `✅ Snippet: ${snippetName}\nDevice: ${deviceName} (${deviceSerial})\n\n` +
    formatGroup('Created', summary.created) + '\n\n' +
    formatGroup('Updated', summary.updated) + '\n\n' +
    formatGroup('Deleted', summary.deleted)
  );
}
// for Object creation tab
function showObjectSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ObjectSidebar')
    .setTitle('Snippet Tags and Address Groups')
    .setWidth(300); // Sidebar width (px)

  SpreadsheetApp.getUi().showSidebar(html);
}
//@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
//--------------------------------------------------------------
// Object creation backend functions
// ------------------------------------------------------------

function fetchTags() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BASECONF');
  const snippetName = sheet.getRange('C15').getValue().trim();  // Trim spaces

  Logger.log(`📌 Snippet Name: ${snippetName}`);
  if (!snippetName) {
    throw new Error("❌ Snippet name is empty. Check BASECONF!B13");
  }

  const auth = getFreshToken();

  const url = `https://api.strata.paloaltonetworks.com/config/objects/v1/tags?snippet=${encodeURIComponent(snippetName)}`;
  Logger.log(`📌 Full URL: ${url}`);

  const options = {
    method: 'get',
    headers: auth.headers,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  Logger.log(`fetchTags() status: ${code}`);
  Logger.log(`fetchTags() body: ${body}`);

  if (code !== 200) {
    throw new Error(`Failed to fetch tags. Status: ${code}, Body: ${body}`);
  }

  const data = JSON.parse(body);
  const tags = data?.data || [];

  Logger.log("Fetched tags: " + JSON.stringify(tags));
  return tags.map(tag => ({
    name: tag.name,
    color: tag.color || '#808080'
  }));
}

// ------------------------------------------------------------
function assignTagsAndGroupToSheet(tags, group) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Object Creation");
  const cell = sheet.getActiveCell();
  const row = cell.getRow();

  if (row < 6) {
    throw new Error("❌ To transfer Address Groups and Tags from the sidebar dialogue into the main window, assign a name to the Address Object in Column B and leave the cursor in that row where you want the Address Group and/or Tags be placed / added !");
  }

  const nameCell = sheet.getRange(row, 2); // Column B
  const nameValue = nameCell.getValue();

  if (!nameValue || typeof nameValue !== 'string' || nameValue.trim() === "") {
    throw new Error("❌ Please define Address Object Name before assigning anything to it.");
  }

  // Write tags to Column F
  const tagString = tags.join(", ");
  sheet.getRange(row, 6).setValue(tagString);

  // Write address group name to Column G, type to Column H
  if (group && typeof group === 'object' && group.name) {
    sheet.getRange(row, 7).setValue(group.name);       // Column G
    sheet.getRange(row, 8).setValue(group.type || ""); // Column H
  } else {
    sheet.getRange(row, 7).clearContent(); // Clear G
    sheet.getRange(row, 8).clearContent(); // Clear H
  }
}

// ------------------------------------------------------------
function createNewTag(name, description) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BASECONF');
  const snippetName = sheet.getRange('C15').getValue();
  const auth = getFreshToken();

  const payload = {
    name: name,
    snippet: snippetName,
    comments: description || ""
  };

  const url = 'https://api.strata.paloaltonetworks.com/config/objects/v1/tags';
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: auth.headers,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  Logger.log(`createNewTag() status: ${code}`);
  Logger.log(`createNewTag() body: ${body}`);

  if (code !== 200 && code !==201) {
    throw new Error(`Failed to create tag. Status: ${code}, Body: ${body}`);
  }
}

// ------------------------------------------------------------
function fetchAddressGroupsBE() {
  Logger.log(`fetchAddressGroups() was triggered`);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BASECONF');
  const snippetName = sheet.getRange('C15').getValue(); 
  //const auth = getFreshToken(); 
  let auth;
  try {
    auth = getFreshToken(); 
    Logger.log(`✅ Token retrieved. First 14 chars: ${auth.headers.Authorization?.substring(0, 14)}...`);
  } catch (err) {
    Logger.log(`❌ Failed to get token: ${err.message}`);
    throw new Error(`❌ Token refresh failed: ${err.message}`);
  }

  const url = `https://api.strata.paloaltonetworks.com/config/objects/v1/address-groups?snippet=${encodeURIComponent(snippetName)}`;
  Logger.log(`📌 Token (first 10): ${auth.headers.Authorization?.substring(0, 14)}...`);
  Logger.log(`📌 URL: ${url}`);

  const options = {
    method: 'get',
    headers: auth.headers,
    muteHttpExceptions: true
  };
  Logger.log(`📌 Full URL: ${url}`);
  Logger.log(`📌 Headers: ${JSON.stringify(options.headers, null, 2)}`);
  
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  Logger.log(`fetchAddressGroups() status: ${code}`);
  Logger.log(`fetchAddressGroups() body: ${body}`);

  if (code !== 200) {
    throw new Error(`Failed to fetch address groups. Status: ${code}, Body: ${body}`);
  }

  const data = JSON.parse(body);
  const groups = data?.data || [];
  
  Logger.log(`✅ Returning ${groups.length} groups`);
  return groups.map(g => {
    const type = g.dynamic ? "dynamic" : (g.static ? "static" : "unknown");
    return {
      name: g.name,
      type: type
    };  
  });
}
//----------------------------------------------------------------------

function pushAddressObjectsToSCM() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName("BASECONF");
  const snippet = base.getRange("C15").getValue().trim();
  const headers = getFreshToken().headers;
  const apiBase = "https://api.strata.paloaltonetworks.com/config/objects/v1";

  Logger.log(`📌 Snippet Name: ${snippet}`);

  // Step 1: Fetch existing address objects
  const listUrl = `${apiBase}/addresses?snippet=${encodeURIComponent(snippet)}`;
  const res = UrlFetchApp.fetch(listUrl, { method: "GET", headers, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    ui.alert(`❌ Failed to fetch address objects. HTTP ${res.getResponseCode()}`, res.getContentText(), ui.ButtonSet.OK);
    return;
  }

  const existingObjects = JSON.parse(res.getContentText()).data || [];
  const existingMap = {};
  const existingNames = new Set();
  existingObjects.forEach(obj => {
    if (obj.name) {
      existingMap[obj.name] = obj;
      existingNames.add(obj.name);
    }
  });

  // Step 2: Parse sheet
  const parsed = parseAddressObjectSheet();

  // Step 3: Preview
  const sheetNames = new Set(parsed.map(row => row.name));
  const deletedObjects = [...existingNames].filter(name => !sheetNames.has(name));
  const logs = parsed.map(row => {
    return `Row ${row.rowNumber}: ${row.name}
 → Type: ${row.type}
 → Value: ${row.value}
 → Description: ${row.desc || "none"}
 → Tags: ${row.tags.length ? row.tags.join(", ") : "none"}
 → Group: ${row.group || "none"} (${row.groupType || "-"})\n`;
  });

  if (deletedObjects.length) {
    logs.push(`🗑️ Objects to be deleted from SCM:\n${deletedObjects.join("\n")}`);
  }

  const preview = logs.join("\n");
  const response = ui.alert("📋 Preview: Address Object Push", preview.substring(0, 3000), ui.ButtonSet.OK_CANCEL);
  if (response !== ui.Button.OK) {
    ui.alert("🚫 Push cancelled.");
    return;
  }

  // Step 4: Tags
  const existingTagNames = fetchTags().map(t => t.name);
  const createdTags = new Set();
  for (const row of parsed) {
    for (const tag of row.tags) {
      if (!existingTagNames.includes(tag) && !createdTags.has(tag)) {
        createNewTag(tag, "");
        createdTags.add(tag);
      }
    }
  }

  // Step 5: Create/update address objects
  const seenNames = new Set();
  parsed.forEach(row => {
    createOrUpdateAddressObject(row, snippet, headers, existingMap, seenNames);
  });

  // Step 6: Delete stale objects
  const toDelete = [...existingNames].filter(name => !seenNames.has(name));
  toDelete.forEach(name => {
    const id = existingMap[name]?.id;
    if (id) {
      const delUrl = `${apiBase}/addresses/${id}`;
      const delRes = UrlFetchApp.fetch(delUrl, { method: "DELETE", headers, muteHttpExceptions: true });
      logResult(`🗑️ Deleted object "${name}"`, delRes);
    }
  });

  // Step 7: Static group full sync
  const groupMap = {};
  for (const row of parsed) {
    if (row.group && row.groupType === "static") {
      if (!groupMap[row.group]) groupMap[row.group] = new Set();
      groupMap[row.group].add(row.name);
    }
  }

  const groupListUrl = `${apiBase}/address-groups?snippet=${encodeURIComponent(snippet)}`;
  const groupRes = UrlFetchApp.fetch(groupListUrl, { method: "GET", headers, muteHttpExceptions: true });
  const groupData = JSON.parse(groupRes.getContentText()).data || [];
  Logger.log(`📦 Fetched ${groupData.length} address groups`);

  const existingGroups = {};
  groupData.forEach(g => {
    if (g.name) existingGroups[g.name] = g;
  });

  // Track which groups were updated
  const touchedGroups = new Set();

  for (const [groupName, memberSet] of Object.entries(groupMap)) {
    const members = [...memberSet];
    const payload = {
      name: groupName,
      description: "",
      tag: [],
      static: members
    };

    if (existingGroups[groupName]) {
      touchedGroups.add(groupName);
      const groupId = existingGroups[groupName].id;
      const putUrl = `${apiBase}/address-groups/${groupId}`;
      const putRes = UrlFetchApp.fetch(putUrl, {
        method: "PUT",
        headers,
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      logResult(`🔁 Updated static group "${groupName}"`, putRes);
    } else {
      touchedGroups.add(groupName);
      const postUrl = `${apiBase}/address-groups`;
      payload.snippet = snippet;
      const postRes = UrlFetchApp.fetch(postUrl, {
        method: "POST",
        headers,
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      logResult(`📘 Created static group "${groupName}"`, postRes);
    }
  }

  // Step 8: Delete static groups not referenced in sheet
  for (const [groupName, group] of Object.entries(existingGroups)) {
    const isStatic = !!group.static;
    const wasTouched = touchedGroups.has(groupName);
    if (isStatic && !wasTouched) {
      const deleteUrl = `${apiBase}/address-groups/${group.id}`;
      const deleteRes = UrlFetchApp.fetch(deleteUrl, {
        method: "DELETE",
        headers,
        muteHttpExceptions: true
      });
      logResult(`🗑️ Deleted static group "${groupName}"`, deleteRes);
    }
  }
  syncDynamicAddressGroups(parsed, snippet, headers, existingGroups, existingTagNames, existingMap);
  ui.alert("✅ Address Objects and Static Groups synced.");
}

// ------------------------------------------------------------
function createOrUpdateStaticGroup(name, members, headers, snippet, groupMap) {
  const apiBase = "https://api.strata.paloaltonetworks.com/config/objects/v1/address-groups";
  const payload = {
    name,
    snippet,
    static: members
  };
  const body = JSON.stringify(payload);

  if (groupMap[name]) {
    const id = groupMap[name].id;
    const url = `${apiBase}/${id}`;
    const res = UrlFetchApp.fetch(url, {
      method: "PUT",
      headers,
      contentType: "application/json",
      payload: body,
      muteHttpExceptions: true
    });
    logResult(`🔁 Updated static group "${name}"`, res);
  } else {
    const res = UrlFetchApp.fetch(apiBase, {
      method: "POST",
      headers,
      contentType: "application/json",
      payload: body,
      muteHttpExceptions: true
    });
    logResult(`📘 Created static group "${name}"`, res);
  }
}

function parseAddressObjectSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Object Creation");
  const data = sheet.getDataRange().getValues();
  const rows = [];

  for (let row = 4; row < data.length; row++) {
    const [
      , name, value, type, desc, tagString, group, groupTypeRaw
    ] = data[row];

    const trimmedName = (name || "").toString().trim();
    const trimmedValue = (value || "").toString().trim();
    const trimmedType = (type || "").toString().trim().toLowerCase();
    const trimmedDesc = (desc || "").toString().trim();
    const trimmedGroup = (group || "").toString().trim();
    const trimmedGroupType = (groupTypeRaw || "").toString().trim().toLowerCase();
    const tags = (tagString || "").toString().split(",").map(t => t.trim()).filter(Boolean);

    if (!trimmedName || !trimmedValue || !trimmedType) continue;

    rows.push({
      name: trimmedName,
      value: trimmedValue,
      type: trimmedType,
      desc: trimmedDesc,
      tags,
      group: trimmedGroup,
      groupType: trimmedGroupType,
      rowNumber: row + 1 // for future error reporting
    });
  }

  return rows;
}
function previewObjectSyncPlan() {
  const parsed = parseAddressObjectSheet();
  const logs = [];

  for (const row of parsed) {
    logs.push(`Row ${row.rowNumber}: ✅ Create object "${row.name}" of type ${row.type} with value ${row.value}`);

    if (row.desc) {
      logs.push(` → Description: "${row.desc}"`);
    }

    logs.push(` → Tags: ${row.tags.length ? row.tags.join(", ") : "empty"}`);

    if (row.group) {
      if (!["static", "dynamic"].includes(row.groupType)) {
        logs.push(` → ❌ Group "${row.group}" has invalid or missing type.`);
      } else if (row.groupType === "static") {
        logs.push(` → Assign to STATIC group: ${row.group}`);
      } else if (row.groupType === "dynamic") {
        if (row.tags.length) {
          logs.push(` → Assign to DYNAMIC group: ${row.group} (match on tags)`);
        } else {
          logs.push(` → ❌ DYNAMIC group requires tag(s), none found.`);
        }
      }
    } else {
      logs.push(` → No group assignment.`);
    }
  }

  const preview = logs.join("\n");
  SpreadsheetApp.getUi().alert("📋 Preview: Planned Object Sync", preview.substring(0, 3000), SpreadsheetApp.getUi().ButtonSet.OK);
}

function getAllAddressObjects(snippet, headers) {
  const url = `https://api.strata.paloaltonetworks.com/config/objects/v1/addresses?snippet=${encodeURIComponent(snippet)}`;
  const res = UrlFetchApp.fetch(url, { method: "GET", headers, muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code !== 200) {
    logError(`❌ Failed to fetch addresses → HTTP ${code}`);
    logError(res.getContentText());
    return [];
  }
  const body = JSON.parse(res.getContentText());
  return body.data || [];
}
function fetchAddressGroups(snippet, headers) {
  const url = `https://api.strata.paloaltonetworks.com/config/objects/v1/address-groups?snippet=${encodeURIComponent(snippet)}`;
  const res = UrlFetchApp.fetch(url, { method: "GET", headers, muteHttpExceptions: true });
  const code = res.getResponseCode();
  if (code !== 200) {
    logError(`❌ Failed to fetch address groups. HTTP ${code}`);
    return [];
  }
  const data = JSON.parse(res.getContentText()).data || [];
  Logger.log(`📦 Fetched ${data.length} address groups`);
  return data;
}

///---------------------------------------------------------
function objectNeedsUpdate(existing, row) {
  const valMatch =
    getSCMValue(existing) === row.value;
  const descMatch =
    (existing.description || "").trim() === (row.desc || "").trim();
  const typeMatch =
    inferTypeFromSCM(existing) === row.type;

  const existingTags = (existing.tag || []).map(t => t.trim()).sort();
  const sheetTags = (row.tags || []).map(t => t.trim()).sort();
  const tagsMatch = JSON.stringify(existingTags) === JSON.stringify(sheetTags);

  return !(valMatch && descMatch && typeMatch && tagsMatch);
}

function inferTypeFromSCM(obj) {
  if (obj.ip_netmask) return "ip address";
  if (obj.range) return "ip range";
  if (obj.ip_wildcard) return "ip wildcard";
  if (obj.fqdn) return "fqdn";
  return "unknown";
}

function getSCMValue(obj) {
  return obj.ip_netmask || obj.range || obj.ip_wildcard || obj.fqdn || "";
}
function logResult(label, data ) {
  const ui = SpreadsheetApp.getUi();
  if (data && typeof data.getResponseCode === "function") {
    const status = data.getResponseCode();
    const body = data.getContentText();
    Logger.log(`${label} → HTTP ${status}`);
    Logger.log(body);
    if (status >= 400) {
      ui.alert(`${label}\n❌ HTTP ${status}\n\n${body}`);
    }
  } else {
    Logger.log(`${label}:`);
    Logger.log(JSON.stringify(data, null, 2));
  }
}

function logError(message) {
  Logger.log(`❌ ${message}`);
}

function createOrUpdateAddressObject(row, snippet, headers, existingMap, seenObjects) {
  const apiBase = "https://api.strata.paloaltonetworks.com/config/objects/v1";

  if (seenObjects) seenObjects.add(row.name);

  const existing = existingMap[row.name];

  const buildPayload = () => {
    const payload = {
      name: row.name,
      description: row.desc?.trim() || "description placeholder",
      tag: row.tags || [],
      snippet: snippet
    };
    const val = row.value;
    if (row.type === "ip address") payload.ip_netmask = val;
    else if (row.type === "ip range") payload.range = val;
    else if (row.type === "ip wildcard") payload.ip_wildcard = val;
    else if (row.type === "fqdn") payload.fqdn = val;
    else {
      logError(`❌ Unknown object type: ${row.type} for "${row.name}"`);
      return null;
    }
    return payload;
  };

  const payload = buildPayload();
  if (!payload) return;

  if (existing) {
    if (!objectNeedsUpdate(existing, row)) {
      Logger.log(`✅ Object "${row.name}" already up to date.`);
      return;
    }

    const updateUrl = `${apiBase}/addresses/${existing.id}`;
    const putRes = UrlFetchApp.fetch(updateUrl, {
      method: "PUT",
      headers,
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    logResult(`🔄 Updated object "${row.name}"`, putRes);
    Logger.log(`🔧 PUT URL: ${updateUrl}`);
    Logger.log(`📤 PUT Payload:`);
    Logger.log(JSON.stringify(payload, null, 2));

  } else {
    const postUrl = `${apiBase}/addresses`;
    const postRes = UrlFetchApp.fetch(postUrl, {
      method: "POST",
      headers,
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    logResult(`📦 Created object "${row.name}"`, postRes);
    Logger.log(`🔧 POST URL: ${postUrl}`);
    Logger.log(`📤 POST Payload:`);
    Logger.log(JSON.stringify(payload, null, 2));
  }
}



function deleteUnlistedAddressObjects(snippet, headers, allAddresses, createdNamesSet) {
  const apiBase = "https://api.strata.paloaltonetworks.com/config/objects/v1";

  const toDelete = allAddresses.filter(obj => !createdNamesSet.has(obj.name));
  if (toDelete.length === 0) {
    Logger.log("🧼 No orphaned address objects to delete.");
    return;
  }

  toDelete.forEach(obj => {
    const deleteUrl = `${apiBase}/addresses/${obj.id}`;
    const delRes = UrlFetchApp.fetch(deleteUrl, {
      method: "DELETE",
      headers,
      muteHttpExceptions: true
    });

    logResult(`🗑️ Deleted object "${obj.name}"`, delRes);
    Logger.log(`🔧 DELETE URL: ${deleteUrl}`);
  });
}

// ------------------------------------------------------------
function pullAddressObjectsToSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Object Creation");
  const baseconf = ss.getSheetByName("BASECONF");
  const snippetName = baseconf.getRange("C15").getValue().trim();
  const auth = getFreshToken();
  const headers = auth.headers;
  if (!snippetName) throw new Error(":x: Snippet name not found in BASECONF!C15");
  // Fetch Address Objects
  const aoUrl = `https://api.strata.paloaltonetworks.com/config/objects/v1/addresses?snippet=${encodeURIComponent(snippetName)}`;
  const aoResponse = UrlFetchApp.fetch(aoUrl, { method: 'get', headers, muteHttpExceptions: true });
  if (aoResponse.getResponseCode() !== 200) {
    throw new Error(`:x: Failed to fetch address objects. Code: ${aoResponse.getResponseCode()}, Body: ${aoResponse.getContentText()}`);
  }
  const aoJson = JSON.parse(aoResponse.getContentText());
  const addressObjects = aoJson?.data || [];
  Logger.log(`:receipt: Raw AO JSON: ${JSON.stringify(aoJson)}`);
  // Logger.log(`:white_check_mark: Received ${addressObjects.length} address objects from SCM`);
  // Fetch Address Groups
  const agUrl = `https://api.strata.paloaltonetworks.com/config/objects/v1/address-groups?snippet=${encodeURIComponent(snippetName)}`;
  const agResponse = UrlFetchApp.fetch(agUrl, { method: 'get', headers, muteHttpExceptions: true });
  if (agResponse.getResponseCode() !== 200) {
    throw new Error(`:x: Failed to fetch address groups. Code: ${agResponse.getResponseCode()}, Body: ${agResponse.getContentText()}`);
  }
  const groups = JSON.parse(agResponse.getContentText())?.data || [];
  const staticMemberships = {};   // Map: objectName → group name
  const dynamicTagMap = {};       // Map: tag → group name
  let hasDynamicGroups = false;
  groups.forEach(group => {
    const name = group.name;
    const isStatic = Array.isArray(group.static);
    const isDynamic = group.dynamic && typeof group.dynamic === 'object';
    Logger.log(`:mag: Group: ${name}, Static: ${isStatic}, Dynamic: ${isDynamic}`);
    if (isStatic) {
      group.static.forEach(objName => {
        staticMemberships[objName] = name;
      });
    }
    if (isDynamic && group.tag && Array.isArray(group.tag)) {
      hasDynamicGroups = true;
      group.tag.forEach(tag => {
        if (!dynamicTagMap[tag]) dynamicTagMap[tag] = name;
      });
    }
  });
  // :white_check_mark: Safely clear old rows (starting from row 5 down, not touching headers)
  const lastRow = Math.max(sheet.getLastRow(), 5);
  if (lastRow > 4) {
    sheet.getRange(`B5:H${lastRow}`).clearContent();
  }
  let row = 5;
  let count = 0;
  let multiTagObjDetected = false;
  addressObjects.forEach(obj => {
    const name = obj.name || "";
    const value = obj.ip_netmask || obj.fqdn || obj.range || "";
    const desc = obj.description || "";
    const tagsArray = obj.tag || [];
    const tags = tagsArray.join(", ");
    let groupName = "";
    let groupType = "";
    // :wrench: Derive type properly
    let type = "";
    if (obj.ip_netmask) type = "ip address";
    else if (obj.fqdn) type = "fqdn";
    else if (obj.range) type = "ip range";
    Logger.log(`:mag: AO: name="${name}" value="${value}" type="${type}" tags="${tags}"`);
    if (!name || !value || !type) return;
    if (tagsArray.length > 1) multiTagObjDetected = true;
    if (staticMemberships[name]) {
      groupName = staticMemberships[name];
      groupType = "static";
    } else {
      for (let i = 0; i < tagsArray.length; i++) {
        const tag = tagsArray[i];
        if (dynamicTagMap[tag]) {
          groupName = dynamicTagMap[tag];
          groupType = "dynamic";
          break;
        }
      }
    }
    // :white_check_mark: Write row
    sheet.getRange(row, 2).setValue(name);      // B: Name
    sheet.getRange(row, 3).setValue(value);     // C: Value
    sheet.getRange(row, 4).setValue(type);      // D: Type
    sheet.getRange(row, 5).setValue(desc);      // E: Description
    sheet.getRange(row, 6).setValue(tags);      // F: Tags
    sheet.getRange(row, 7).setValue(groupName); // G: Group
    sheet.getRange(row, 8).setValue(groupType); // H: Group Type
    row++;
    count++;
  });
  ui.alert(`:white_check_mark: Pulled ${count} address object(s) from snippet "${snippetName}".`);
  if (hasDynamicGroups && multiTagObjDetected) {
    ui.alert(":warning: Snippet contains Objects with multiple tags AND Dynamic Address Groups.\nWe do not have a full logic for this implemented.\nPlease check in SCM GUI that your Objects are correct.");
  }
}
//---------------------------------------------------------------------------------------------------
function syncDynamicAddressGroups(parsedRows, snippet, headers, existingGroups, existingTags, existingObjectsMap) {
  const ui = SpreadsheetApp.getUi();
  const apiBase = "https://api.strata.paloaltonetworks.com/config/objects/v1";

  const dynamicGroupMap = {};  // groupName → Set of tags (from objects)
  const tagToObjects = {};     // tag → Set of object names

  // Step 1: collect requested dynamic groups and tags
  for (const row of parsedRows) {
    if (row.group && row.groupType === "dynamic") {
      if (!dynamicGroupMap[row.group]) dynamicGroupMap[row.group] = new Set();
      if (row.tags.length === 0) {
        // 🚨 Step 1: Error — dynamic group requested, but no tags
        ui.alert(`❌ Object "${row.name}" references dynamic group "${row.group}" but has no tags.\n\n❗ I don't know what to do.\nUse the SCM GUI to fix the situation, or add tags to the object.`);
        continue;
      }
      row.tags.forEach(t => {
        dynamicGroupMap[row.group].add(t);
        if (!tagToObjects[t]) tagToObjects[t] = new Set();
        tagToObjects[t].add(row.name);
      });
    }
  }

  // Step 2: ensure each tag used exists
  for (const tag of Object.keys(tagToObjects)) {
    if (!existingTags.includes(tag)) {
      createNewTag(tag, "");
    }
  }

  // Step 3: create or update dynamic groups
  for (const [groupName, tagSet] of Object.entries(dynamicGroupMap)) {
    const tagArray = [...tagSet];
    const existingGroup = existingGroups[groupName];
    const filterFragment = tagArray.map(t => `'${t}'`).join(" or ");

    if (!existingGroup) {
      // 🆕 Group doesn't exist yet → create
      const payload = {
        name: groupName,
        snippet: snippet,
        dynamic: { filter: filterFragment },
        tag: tagArray
      };
      const res = UrlFetchApp.fetch(`${apiBase}/address-groups`, {
        method: "POST",
        headers,
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      logResult(`📘 Created dynamic group "${groupName}"`, res);
    } else if (existingGroup.dynamic) {
      // 🔁 Group exists → update only if needed
      const currentFilter = existingGroup.dynamic.filter;
      let updated = false;
      let updatedTags = new Set(existingGroup.tag || []);
      const existingTagSet = new Set(existingGroup.tag || []);
      const newTagsToAdd = tagArray.filter(t => !existingTagSet.has(t));

      // 3.1 OR-add tags to group filter (don’t remove anything)
      if (newTagsToAdd.length) {
        const additions = newTagsToAdd.map(t => `'${t}'`).join(" or ");
        const newFilter = `(${currentFilter}) or ${additions}`;
        const payload = {
          name: groupName,
          dynamic: {
            filter: newFilter
          },
          tag: [...new Set([...existingGroup.tag, ...newTagsToAdd])]
        };
        const res = UrlFetchApp.fetch(`${apiBase}/address-groups/${existingGroup.id}`, {
          method: "PUT",
          headers,
          contentType: "application/json",
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });
        logResult(`🔁 Updated dynamic group "${groupName}" filter`, res);
      }
    }
  }

  // Step 4: ensure objects have matching tags
  for (const row of parsedRows) {
    if (row.group && row.groupType === "dynamic") {
      const groupTags = dynamicGroupMap[row.group] || new Set();
      const objectTags = new Set(row.tags || []);
      const objectMissingTags = [...groupTags].filter(t => !objectTags.has(t));

      if (objectMissingTags.length > 0) {
        // Update the object: add tags required by the group
        const updatedTagList = [...new Set([...objectTags, ...groupTags])];
        const objectId = existingObjectsMap[row.name]?.id;
        const valueField = row.value.includes("/") ? "ip_netmask" : "fqdn";

        const payload = {
          name: row.name,
          description: row.desc || "",
          tag: updatedTagList,
          [valueField]: row.value,
          snippet: snippet
        };

        const res = UrlFetchApp.fetch(`${apiBase}/addresses/${objectId}`, {
          method: "PUT",
          headers,
          contentType: "application/json",
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });
        logResult(`🔁 Updated tags for "${row.name}"`, res);
      }
    }
  }

  // Step 5: delete unused dynamic groups (if all tags gone)
  for (const [groupName, group] of Object.entries(existingGroups)) {
    if (!group.dynamic) continue;
    const groupTags = group.tag || [];

    const allTagsUnused = groupTags.every(t => !tagToObjects[t] || tagToObjects[t].size === 0);
    if (allTagsUnused) {
      const delUrl = `${apiBase}/address-groups/${group.id}`;
      const res = UrlFetchApp.fetch(delUrl, {
        method: "DELETE",
        headers,
        muteHttpExceptions: true
      });
      logResult(`🗑️ Deleted dynamic group "${groupName}" (all tags unused)`, res);
    }
  }
}

//////////////////////////////////////////////////////////////
/// moving into frontend / backend dialogs
/// ---------------------------------------------------------
/// VLAN Dialog 
///----------------------------------------------------------
function showVlanDialog() {
  clearVlanLog();  // Wipe and start fresh
  logVlan("🟢 VLAN dialog opened");

  const html = HtmlService.createHtmlOutputFromFile("VlanDialogDocked")
    .setTitle("VLAN IF Manager")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

function clearVlanLog() {
  const sheet = getOrCreateVlanLogSheet();
  sheet.clearContents();
  sheet.appendRow(["Timestamp", "Message"]);
}

function logVlan(message) {
  const sheet = getOrCreateVlanLogSheet();
  const time = new Date().toISOString();
  sheet.appendRow([time, message]);
}

function getOrCreateVlanLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("VLAN_DEBUG_LOG");
  if (!sheet) {
    sheet = ss.insertSheet("VLAN_DEBUG_LOG");
    sheet.hideSheet();
  }
  return sheet;
}

/// ---------------------------------------------------------
/// VRF Dialog 
///----------------------------------------------------------
function showVrfDialog() {
  clearVrfLog();
  logVrf("🟢 VRF dialog opened");
  const html = HtmlService.createHtmlOutputFromFile("VrfDialogDocked")
    .setTitle("Manage VRFs")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getOrCreateVrfLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("VRF_DEBUG_LOG");
  if (!sheet) {
    sheet = ss.insertSheet("VRF_DEBUG_LOG");
    sheet.hideSheet();
  }
  return sheet;
}

function logVrf(message) {
  const sheet = getOrCreateVrfLogSheet();
  const time = new Date().toISOString();
  sheet.appendRow([time, message]);
}

function clearVrfLog() {
  const sheet = getOrCreateVrfLogSheet();
  sheet.clearContents();
  sheet.appendRow(["Timestamp", "Message"]);
}
///---------       ------------      ------------    ----------
/// Static Route popup

function showStaticRoutesDialogWithObject(routerJson) {
  const router = JSON.parse(routerJson);
  logVrf(`📡 Opening static route dialog for: ${JSON.stringify(router, null, 2)}`);
  PropertiesService.getUserProperties().setProperty("STATIC_ROUTE_NAME", router.name);
  PropertiesService.getUserProperties().setProperty("STATIC_ROUTE_UUID", router.id);
  PropertiesService.getUserProperties().setProperty("STATIC_ROUTE_OBJECT", routerJson);

  const html = HtmlService.createHtmlOutputFromFile("StaticRoutesDialogBox")
    .setWidth(1200)
    .setHeight(500)
    .setTitle("Static Routes");
    
  SpreadsheetApp.getUi().showModalDialog(html, "Manage Static Routes");
}

/// ---------------------------------------------------------
/// SNAT and DHCP Dialog 
///----------------------------------------------------------
function showSnatDhcpDialog() {
  // Create & show modal dialog
  const html = HtmlService.createHtmlOutputFromFile("SnatDhcpDialogDocked")
    .setTitle("NAT and DHCP Services");
  SpreadsheetApp.getUi().showSidebar(html);
  
  // Log that we initiated opening (frontend will also log when fully loaded)
  logSnatDhcp("🚀 showSnatDhcpDialog() invoked");
}
function openSnatDhcpServices() {
  showSnatDhcpDialog();
}

function getOrCreateSnatDhcpLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "SNAT-DHCP";
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, 2).setValues([["Timestamp", "Message"]]);
  }
  return sheet;
}

function logSnatDhcp(message) {
  try {
    const sheet = getOrCreateSnatDhcpLogSheet();
    const time = new Date().toISOString();
    sheet.appendRow([time, message]);
  } catch (e) {
    // Last-resort logging to executions log
    console.error("logSnatDhcp error:", e && e.message ? e.message : e);
  }
}

/** Optional: clear the SNAT-DHCP log intentionally (call on close if desired). */
function flushSnatDhcpLog() {
  const sheet = getOrCreateSnatDhcpLogSheet();
  // keep header row
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  logSnatDhcp("🧹 SNAT-DHCP log flushed");
}
/// ---------------------------------------------------------
/// Zone Dialog 
///----------------------------------------------------------
function showZonesDialog() {
  // Open the docked sidebar with your Zones HTML
  const html = HtmlService.createHtmlOutputFromFile("ZonesDialogDocked")
    .setTitle("Zones");
  SpreadsheetApp.getUi().showSidebar(html);

  // Initial breadcrumb (frontend can also log once fully loaded)
  logZone("🚀 showZonesDialog() invoked");
}

// Small wrapper so you can “Assign script…” to a drawing
function openZonesServices() {
  showZonesDialog();
}

/** Ensure the debug sheet exists (name: debug_Zones) and return it */
function getOrCreateZonesLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "debug_Zones";
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, 2).setValues([["Timestamp", "Message"]]);
  }
  return sheet;
}

/** Append a log line to debug_Zones */
function logZone(message) {
  try {
    const sheet = getOrCreateZonesLogSheet();
    const time = new Date().toISOString();
    sheet.appendRow([time, String(message)]);
  } catch (e) {
    console.error("logZone error:", e && e.message ? e.message : e);
  }
}

/** Optional: clear the log (keeps header) */
function flushZoneLog() {
  const sheet = getOrCreateZonesLogSheet();
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
  logZone("🧹 debug_Zones log flushed");
}


///----------------------------------------------------------
///----------------------------------------------------------
/// SCM API calls section
///----------------------------------------------------------
function getStaticRouteContext() {
  const name = PropertiesService.getUserProperties().getProperty("STATIC_ROUTE_NAME");
  const id = PropertiesService.getUserProperties().getProperty("STATIC_ROUTE_UUID");
  if (!name || !id) throw new Error("Missing VRF context in dialog");
  logVrf(`📦 Static Route Context: ${name} (UUID: ${id})`);

  let interfaces = [];
  try {
    const cached = getSelectedVrfRouter(); // uses STATIC_ROUTE_OBJECT
    interfaces = cached?.vrf?.[0]?.interface || [];
    logVrf(`🧰 Context interfaces (cached): ${interfaces.length} item(s)`);
  } catch (e) {
    logVrf(`ℹ️ No cached router available for interfaces, fetching full router...`);
    const { headers } = getFreshToken();
    const full = getLogicalRouterById(id, headers);
    interfaces = full?.vrf?.[0]?.interface || [];
    logVrf(`🧰 Context interfaces (fetched): ${interfaces.length} item(s)`);
  }
  logVrf(`📦 Static Route Context: ${name} (UUID: ${id})`);
  return { name, id, interfaces: [] };
}

function getVrfInterfacesForDialog() {
  const id = PropertiesService.getUserProperties().getProperty("STATIC_ROUTE_UUID");
  const name = PropertiesService.getUserProperties().getProperty("STATIC_ROUTE_NAME");
  if (!name || !id) throw new Error("No VRF context for interface refresh");

  const { headers } = getFreshToken();
  logVrf(`🔄 Refreshing interfaces for VRF '${name}' (${id})`);
  const router = getLogicalRouterById(id, headers);

  // Base interfaces (includes dotted subinterfaces if present in VRF)
  const baseIfs = router?.vrf?.[0]?.interface || [];
  logVrf(`🔢 Interfaces from VRF block: ${baseIfs.length}`);

  /// add tagged interfaces ///

  return baseIfs;
}

function getVrfNamesForDialog() {
  const { snippet } = getBaseconfSnippetData();
  const vrfs = fetchVrfs(snippet) || [];
  const names = vrfs.map(v => v.name).filter(Boolean).sort();
  logVrf(`📦 VRFs for dialog: ${names.join(", ")}`);
  return names;
}

function getSelectedVrfRouter() {
  const up = PropertiesService.getUserProperties();
  const raw = up.getProperty("STATIC_ROUTE_OBJECT");
  if (!raw) throw new Error("STATIC_ROUTE_OBJECT not set");
  try {
    const router = JSON.parse(raw);
    if (!router || !router.id || !router.name) throw new Error("Router object incomplete");
    return router;
  } catch (e) {
    throw new Error("Failed to parse STATIC_ROUTE_OBJECT JSON");
  }
}
function debugStaticRouteProps() {
  const up = PropertiesService.getUserProperties();
  const name = up.getProperty("STATIC_ROUTE_NAME");
  const id   = up.getProperty("STATIC_ROUTE_UUID");
  const obj  = up.getProperty("STATIC_ROUTE_OBJECT");
  logVrf(`🔎 STATIC_ROUTE_NAME=${name || "(none)"}`);
  logVrf(`🔎 STATIC_ROUTE_UUID=${id || "(none)"}`);
  logVrf(`🔎 STATIC_ROUTE_OBJECT length=${obj ? obj.length : 0}`);
  return { name, id, hasObject: !!obj };
}


function getStaticRoutesForDialog() {
  logVrf("🚀 getStaticRoutesForDialog() called");
  const { token, headers } = getFreshToken();
  const id = PropertiesService.getUserProperties().getProperty("STATIC_ROUTE_UUID");
  const name = PropertiesService.getUserProperties().getProperty("STATIC_ROUTE_NAME");
  const router = getLogicalRouterById(id, headers);

  logVrf(`📡 Raw router payload for ${name}: ${JSON.stringify(router).substring(0, 1200)}…`);
  logVrf(`🚀 getStaticRoutesForDialog() called — name="${name}" id="${id}"`);
  if (!name || !id) throw new Error(":x: No VRF context for static route dialog");
  const routes = router?.vrf?.[0]?.routing_table?.ip?.static_route || [];
  logVrf(`:package: Found ${routes.length} static route(s) in router "${name}"`);
  if (routes.length) {
    logVrf(`🔍 First route: ${JSON.stringify(routes[0])}`);
  }
  return routes;
}

/// ---------------------------------------------------------
/// IP 2 NIC Dialog 
///----------------------------------------------------------
function showIp2IfDialog() {
  clearIp2IfLog();  // Start fresh
  logIp2If(":large_green_circle: IP2NIC dialog opened");
  const html = HtmlService.createHtmlOutputFromFile("Ip2IfDialogDocked")
    .setTitle("IP ⇄ Interface Manager")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getOrCreateIp2IfLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("IP2NIC_DEBUG_LOG");
  if (!sheet) {
    sheet = ss.insertSheet("IP2NIC_DEBUG_LOG");
    sheet.hideSheet();
  }
  return sheet;
}

function logIp2If(message) {
  const sheet = getOrCreateIp2IfLogSheet();
  const time = new Date().toISOString();
  sheet.appendRow([time, message]);
}

function clearIp2IfLog() {
  const sheet = getOrCreateIp2IfLogSheet();
  sheet.clearContents();
  sheet.appendRow(["Timestamp", "Message"]);
}

function getSnippetVariablesForDialog(kind) {
  if (kind !== "ip" && kind !== "fqdn") {
    throw new Error("getSnippetVariablesForDialog: invalid kind (use 'ip' or 'fqdn').");
  }

  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();
  const url = `https://api.strata.paloaltonetworks.com/config/setup/v1/variables?snippet=${encodeURIComponent(snippet)}`;

  logVrf(`🔎 getSnippetVariablesForDialog(${kind}) → ${url}`);

  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers,
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();

  logVrf(`⬅️ [${code}] Variables list`);
  if (code < 200 || code >= 300) {
    throw new Error(`Variables list failed: ${code} ${body}`);
  }

  let json;
  try {
    json = JSON.parse(body);
  } catch (e) {
    throw new Error("Failed to parse variables response JSON.");
  }

  const items = Array.isArray(json) ? json : (json.data || []);
  const uniq = new Set();

  const lower = s => (s || "").toLowerCase();

  const looksLikeIpValue = v => {
    const val = (v.value || v.current || "").toString();
    return /^\d{1,3}(\.\d{1,3}){3}(\/\d{1,2})?$/.test(val); // 1.2.3.4 or 1.2.3.4/24
  };

  const looksLikeFqdnValue = v => {
    const val = (v.value || v.current || "").toString();
    return /^[a-z0-9.-]+\.[a-z]{2,}$/i.test(val); // simple fqdn heuristic
  };

  const isIpVar = v => {
    const t = lower(v.type || v.var_type);
    if (t.includes("ip")) return true;                       // ip, ip-address, ip_netmask, etc.
    return looksLikeIpValue(v);
  };

  const isFqdnVar = v => {
    const t = lower(v.type || v.var_type);
    if (t.includes("fqdn") || t.includes("domain") || t.includes("hostname")) return true;
    return looksLikeFqdnValue(v);
  };

  const pick = kind === "ip" ? isIpVar : isFqdnVar;

  items.forEach(v => {
    const name = v && v.name;
    if (!name) return;
    if (pick(v)) uniq.add(name);
  });

  const out = Array.from(uniq).sort((a, b) => a.localeCompare(b));
  logVrf(`📋 getSnippetVariablesForDialog(${kind}) → ${out.length} name(s)`);
  return out;
}



///=========================================================================================
function addNewIpVariable(ip, name) {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();

  const url = `https://api.strata.paloaltonetworks.com/config/setup/v1/variables?snippet=${encodeURIComponent(snippet)}`;
  logIp2If(`📤 Creating IP variable at: ${url}`);

  const payload = {
    name: name,
    type: "ip-netmask",
    value: ip,
    description: "",
    snippet: snippet
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const body = response.getContentText();
    logIp2If(`📥 Create IP variable response [${status}]: ${body}`);

    if (status === 201 || status === 200) {
      logIp2If(`✅ Variable ${name} created with IP ${ip}`);
      return true;
    } else {
      logIp2If(`❌ Failed to create variable: ${status}`);
      return false;
    }

  } catch (err) {
    logIp2If(`❌ Exception while creating variable: ${err.message}`);
    return false;
  }
}

/** Fetch NIC interface Variables for the current snippet */
///=========================================================================================
//function fetchInterfacesVlan() {
//  const { snippet } = getBaseconfSnippetData();
//  const { headers } = getFreshToken();
//
//  const url = `https://api.strata.paloaltonetworks.com/config/network/v1/ethernet-interfaces?snippet=${encodeURIComponent(snippet)}`;
//  logVlan(`📤 Fetching NICs from: ${url}`);

//  try {
//    const response = UrlFetchApp.fetch(url, {
//      headers,
//      muteHttpExceptions: true
//    });

//    const status = response.getResponseCode();
//    const body = response.getContentText();
//    logVlan(`📥 NIC response [${status}]: ${body}`);

//    if (status === 200) {
//      const data = JSON.parse(body);
//      const nicNames = (data?.data || []).map(obj => obj.name);
//      logVlan(`✅ Interfaces found: ${nicNames.join(", ")}`);
//      return nicNames;
//    } else {
//      logVlan(`❌ Failed to fetch interfaces: status ${status}`);
//      return [];
//    }
//  } catch (err) {
//    logVlan(`❌ Exception during NIC fetch: ${err.message}`);
//    return [];
//  }
//}
function fetchInterfacesVlan() {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();

  // 1) L3 ethernet interfaces
  const urlL3Eth = `https://api.strata.paloaltonetworks.com/config/network/v1/ethernet-interfaces?snippet=${encodeURIComponent(snippet)}`;
  logVrf(`🔎 fetchInterfacesVlan(): GET L3 ethernet interfaces → ${urlL3Eth}`);
  const resEth = UrlFetchApp.fetch(urlL3Eth, { method: "get", headers, muteHttpExceptions: true });
  const codeEth = resEth.getResponseCode();
  const bodyEth = resEth.getContentText();
  if (codeEth !== 200) throw new Error(`fetchInterfacesVlan: ethernet-interfaces failed [${codeEth}] ${bodyEth}`);
  const ethJson = JSON.parse(bodyEth);
  const ethList = Array.isArray(ethJson) ? ethJson : (ethJson.data || []);
  const ethNames = ethList.map(i => i && i.name).filter(Boolean);

  // 2) L3 subinterfaces (tagged)
  const urlL3Sub = `https://api.strata.paloaltonetworks.com/config/network/v1/layer3-subinterfaces?snippet=${encodeURIComponent(snippet)}`;
  logVrf(`🔎 fetchInterfacesVlan(): GET L3 subinterfaces → ${urlL3Sub}`);
  const resSub = UrlFetchApp.fetch(urlL3Sub, { method: "get", headers, muteHttpExceptions: true });
  const codeSub = resSub.getResponseCode();
  const bodySub = resSub.getContentText();
  if (codeSub !== 200) throw new Error(`fetchInterfacesVlan: layer3-subinterfaces failed [${codeSub}] ${bodySub}`);
  const subJson = JSON.parse(bodySub);
  const subList = Array.isArray(subJson) ? subJson : (subJson.data || []);
  const subNames = subList.map(s => s && s.name).filter(Boolean); // e.g. "$ENTERPRISE.810"

  // 3) Merge + de-dupe + sort
  const all = Array.from(new Set([...ethNames, ...subNames])).sort((a,b) => a.localeCompare(b));

  logVrf(`📋 fetchInterfacesVlan(): ${ethNames.length} L3, ${subNames.length} subif → ${all.length} total`);
  return all; // your frontend already filters out those in vrf[0].interface
}

function fetchInterfaces() {
  return fetchInterfacesGeneric(logIp2If);
}

function fetchInterfacesGeneric(logFn) {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();

  const url = `https://api.strata.paloaltonetworks.com/config/network/v1/ethernet-interfaces?snippet=${encodeURIComponent(snippet)}`;
  logFn(`📤 Fetching NICs from: ${url}`);

  try {
    const response = UrlFetchApp.fetch(url, {
      headers,
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const body = response.getContentText();
    logFn(`📥 NIC response [${status}]: ${body}`);

    if (status === 200) {
      const data = JSON.parse(body);
      const interfaces = (data?.data || []).map(obj => ({
        name: obj.name || "(unnamed)",
        assignedIps: (obj.layer3?.ip || []).map(ipObj => ipObj.name),
        favorite: obj.default_value || null ,
        uuid: obj.id || null
      }));
      logFn(`✅ Interfaces processed: ${interfaces.map(i => `${i.name} (fav: ${i.favorite || 'none'})`).join(", ")}`);
      return interfaces;
    } else {
      logFn(`❌ Failed to fetch interfaces: status ${status}`);
      return [];
    }
  } catch (err) {
    logFn(`❌ Exception during NIC fetch: ${err.message}`);
    return [];
  }
}
function createNicVariable(interfaceName, nicVarName, ipVarNames) {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();

  SpreadsheetApp.getUi().alert(
    "Creating NIC with:\n" +
    "Interface: " + interfaceName + "\n" +
    "NIC Variable: " + nicVarName + "\n" +
    "IP Variables: " + (ipVarNames && ipVarNames.length ? ipVarNames.join(", ") : "None")
  );

  const url = `https://api.strata.paloaltonetworks.com/config/network/v1/ethernet-interfaces`;
  logIp2If(`📤 Creating NIC via: ${url}`);

  const payload = {
    name: nicVarName,
    snippet: snippet,
    layer3: {
      ip: ipVarNames.map(name => ({ name }))
    }
  };
  if (interfaceName && interfaceName !== "none") {
    payload.default_value = interfaceName;
  }
  logIp2If(`📦 Payload: ${JSON.stringify(payload)}`);

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        ...headers,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const body = response.getContentText();
    logIp2If(`📥 NIC creation response [${status}]: ${body}`);

    if (status === 200 || status === 201) {
      logIp2If(`✅ NIC ${nicVarName} created on ${interfaceName}`);
    } else {
      logIp2If(`❌ NIC creation failed: status ${status}`);
    }
  } catch (err) {
    logIp2If(`❌ Exception during NIC creation: ${err.message}`);
  }
}
///=========================================================================================
function getAvailableIpVars(interfaceName, freeOnly) {
  const { snippet } = getBaseconfSnippetData();
  //const { headers } = getFreshToken();
  const ft = getFreshToken();
   const headers = {
    Authorization: 'Bearer ' + String(ft.token || ''),
    Accept: 'application/json'
  }

  const url = `https://api.strata.paloaltonetworks.com/config/setup/v1/variables?snippet=${encodeURIComponent(snippet)}`;
  logIp2If(`📤 Fetching snippet variables from: ${url}`);
  logIp2If(`🧪 headers keys=${Object.keys(headers).join(', ')}`);

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: headers,
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const body = response.getContentText();
    logIp2If(`📥 Variables response [${status}]: ${body}`);

    if (status !== 200) {
      logIp2If(`❌ Failed to fetch IP variables: status ${status}`);
      return [];
    }

    const data = JSON.parse(body);
    const allVars = data?.data || [];

    const ipVars = allVars.filter(v => v.type === "ip-netmask");

    let available;

    if (freeOnly) {
      available = ipVars.filter(v => (v.used_in || []).length === 0);
      logIp2If(`✅ Mode: Free only — ${available.length} IPs`);
    } else {
      available = ipVars.filter(v => {
        const usedIn = v.used_in || [];
        return !usedIn.some(u => u.name === interfaceName);
      });
      logIp2If(`✅ Mode: Assigned elsewhere ok — ${available.length} IPs`);
    }

    const result = available.map(v => ({ name: v.name, value: v.value }));
    logIp2If(`✅ Available IPs for ${interfaceName}: ${result.map(r => r.name).join(", ")}`);

    return result;

  } catch (err) {
    logIp2If(`❌ Exception during variable fetch: ${err.message}`);
    return [];
  }
}

///=========================================================================================
/** Delete NIC by UUID from current snippet */
function removeNicVariable(nicUuid) {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();

  if (!nicUuid) {
    logIp2If(`❌ No NIC UUID provided for deletion.`);
    return;
  }

  const url = `https://api.strata.paloaltonetworks.com/config/network/v1/ethernet-interfaces/${nicUuid}`;
  logIp2If(`📤 DELETE NIC by UUID: ${url}`);

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'delete',
      headers,
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const body = response.getContentText();
    logIp2If(`📥 NIC DELETE response [${status}]: ${body}`);

    if (status === 200 || status === 204) {
      logIp2If(`✅ NIC deleted successfully.`);
    } else {
      logIp2If(`❌ NIC delete failed with status ${status}`);
    }
  } catch (err) {
    logIp2If(`❌ Exception during NIC delete: ${err.message}`);
  }
}
///=========================================================================================
/** Delete NIC by name from current snippet */
function deleteInterfaceByName(nicName) {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();

  const listUrl = `https://api.strata.paloaltonetworks.com/config/network/v1/ethernet-interfaces?snippet=${encodeURIComponent(snippet)}`;
  logVlan(`🔍 Fetching NICs for deletion match: ${listUrl}`);

  try {
    const listResp = UrlFetchApp.fetch(listUrl, { headers, muteHttpExceptions: true });
    const listCode = listResp.getResponseCode();
    const listBody = listResp.getContentText();
    logVlan(`📥 Interface list response [${listCode}]: ${listBody}`);

    if (listCode !== 200) {
      logVlan(`❌ Could not retrieve interfaces to find UUID for ${nicName}`);
      return { status: listCode, body: listBody };
    }

    const interfaces = JSON.parse(listBody)?.data || [];
    const match = interfaces.find(iface => iface.name === nicName);

    if (!match) {
      logVlan(`❌ No interface found with name ${nicName}`);
      return { status: 404, body: `Interface ${nicName} not found` };
    }

    const uuid = match.id;
    const deleteUrl = `https://api.strata.paloaltonetworks.com/config/network/v1/ethernet-interfaces/${uuid}`;
    logVlan(`🗑 Deleting NIC ${nicName} (UUID: ${uuid})`);
    logVlan(`📤 DELETE URL: ${deleteUrl}`);

    const delResp = UrlFetchApp.fetch(deleteUrl, {
      method: "delete",
      headers,
      muteHttpExceptions: true
    });

    const delCode = delResp.getResponseCode();
    const delBody = delResp.getContentText();
    logVlan(`📥 Delete response [${delCode}]: ${delBody}`);

    if (delCode >= 200 && delCode < 300) {
      logVlan(`✅ NIC ${nicName} deleted successfully.`);
    } else {
      logVlan(`❌ Failed to delete NIC ${nicName}`);
    }

    return { status: delCode, body: delBody };

  } catch (err) {
    logVlan(`❌ Exception during NIC delete: ${err.message}`);
    return { status: 500, body: err.message };
  }
}

///=========================================================================================
/** Create NIC variable in snippet */
function createInterface(nicName, defaultValue, comment) {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();

  const url = 'https://api.strata.paloaltonetworks.com/config/network/v1/ethernet-interfaces';
  logVlan(`🆕 Creating NIC variable: ${nicName} → ${defaultValue}`);
  logVlan(`📤 POST URL: ${url}`);

  const payload = {
    name: nicName,
    "default-value": defaultValue,  
    snippet: snippet,
    comment: comment || "",
    "link-speed": "auto",
    "link-duplex": "auto",
    "link-state": "auto",
    poe: {
      "poe-enabled": false,
      "poe-rsvd-pwr": 0
    },
    layer3: {}  
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers,
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const body = response.getContentText();
    logVlan(`📥 Create NIC response [${status}]: ${body}`);

    if (status >= 200 && status < 300) {
      logVlan(`✅ NIC ${nicName} created successfully.`);
    } else {
      logVlan(`❌ NIC creation failed with status ${status}`);
      throw new Error("NIC creation failed: " + body);
    }

    return { status, body };
  } catch (err) {
    logVlan(`❌ Error creating NIC: ${err.message}`);
    throw err;
  }
}

///=========================================================================================
function getSubinterfacesForInterface(nicName) {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();

  const url = `https://api.strata.paloaltonetworks.com/config/network/v1/layer3-subinterfaces?snippet=${encodeURIComponent(snippet)}`;
  logVlan(`📤 Fetching subinterfaces from: ${url}`);

  try {
    const response = UrlFetchApp.fetch(url, {
      headers,
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const body = response.getContentText();
    logVlan(`📥 Subinterface response [${status}]: ${body}`);

    if (status === 200) {
      const data = JSON.parse(body);
      const subifs = (data?.data || []).filter(s => s.name.startsWith(nicName + "."));
      logVlan(`✅ Found subinterfaces: ${subifs.length}`);
      return subifs.map(s => ({
        uuid: s.id,
        name: s.name,
        tag: s['vlan-tag'] || '',
        description: s.comment || ''
      }));
    } else {
      logVlan(`❌ Failed to fetch subinterfaces: status ${status}`);
      return [];
    }
  } catch (err) {
    logVlan(`❌ Exception during subinterface fetch: ${err.message}`);
    return [];
  }
}
///=========================================================================================
function deleteSelectedSubinterfaces(uuids) {
  const { headers } = getFreshToken();
  const results = {
    successCount: 0,
    failures: []
  };

  uuids.forEach(uuid => {
    const url = `https://api.strata.paloaltonetworks.com/config/network/v1/layer3-subinterfaces/${uuid}`;
    logVlan(`🗑 Deleting subinterface UUID: ${uuid}`);
    logVlan(`📤 DELETE URL: ${url}`);

    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'delete',
        headers,
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      const body = response.getContentText();
      logVlan(`📥 Delete response [${code}]: ${body}`);

      if (code >= 200 && code < 300) {
        results.successCount++;
        logVlan(`✅ Subinterface UUID ${uuid} deleted`);
      } else {
        results.failures.push({ uuid, status: code, body });
        logVlan(`❌ Failed to delete subinterface UUID ${uuid}`);
      }

    } catch (err) {
      results.failures.push({ uuid, error: err.message });
      logVlan(`❌ Exception deleting subinterface UUID ${uuid}: ${err.message}`);
    }
  });

  return results;
}
///=========================================================================================

function getAvailableIpVars(arg1, arg2) {
  // ---- 1) Normalize args & choose logger -----------------------------------
  let ctx = 'vlan';            // default context
  let interfaceName = '';      // parent NIC (for VLAN dialog filtering)
  let freeOnly = false;        // "unused only" filter
  let log = typeof logVlan === 'function' ? logVlan : (m)=>{};

  if (typeof arg1 === 'object' && arg1 !== null) {
    // Old style: { for: "ip2nic" | "vlan", unusedOnly?: bool }
    ctx = (arg1.for || 'vlan').toLowerCase();
    freeOnly = !!arg1.unusedOnly;
  } else {
    // New style: (interfaceName: string, freeOnly: bool)
    interfaceName = String(arg1 || '');
    freeOnly = !!arg2;
  }

  if (ctx === 'ip2nic' && typeof logIp2If === 'function') log = logIp2If;
  // VLAN uses logVlan (default)

  log(`📡 getAvailableIpVars() called with ctx='${ctx}', parent='${interfaceName || "-"}', freeOnly=${freeOnly}`);

  // ---- 2) Build safe headers + URL -----------------------------------------
  const { snippet } = getBaseconfSnippetData();

  const ft = getFreshToken();
  const token = String(ft && ft.token ? ft.token : '');
  const headers = {
    Authorization: 'Bearer ' + token,
    Accept: 'application/json'
    // NOTE: do not set Content-Type on GET
  };

  const url = `https://api.strata.paloaltonetworks.com/config/setup/v1/variables?snippet=${encodeURIComponent(snippet)}`;
  log(`📡 GET ${url}`);
  try { log(`🧪 headers keys=${Object.keys(headers).join('|')} types=${Object.values(headers).map(v=>typeof v).join('|')}`); } catch(_) {}

  // ---- 3) Fetch + parse -----------------------------------------------------
  let status = 0, body = '';
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: headers,
      muteHttpExceptions: true
    });
    status = response.getResponseCode();
    body = response.getContentText();
    log(`📥 Variables response [${status}]`);
  } catch (e) {
    log(`❌ Exception during variable fetch: ${e && e.message ? e.message : e}`);
    return [];
  }

  if (status !== 200) {
    log(`❌ Failed to fetch IP variables: status ${status} bodyLen=${body ? body.length : 0}`);
    return [];
  }

  let data;
  try {
    data = JSON.parse(body);
  } catch (e) {
    log(`❌ JSON parse error: ${e && e.message ? e.message : e}`);
    return [];
  }

  const allVars = Array.isArray(data) ? data : (data.data || []);
  // Keep only IP-type variables
  const ipVars = allVars.filter(v => (v.type || '').toLowerCase() === 'ip-netmask');

  // ---- 4) Filter rules per context -----------------------------------------
  // ip2nic (old style): return all or only-unused (based on unusedOnly)
  // vlan (new style): if interfaceName provided, exclude IPs already used on that NIC unless freeOnly forces unused only
  let filtered;
  if (ctx === 'ip2nic') {
    filtered = freeOnly
      ? ipVars.filter(v => (v.used_in || []).length === 0)
      : ipVars.slice();
    log(`✅ ip2nic mode → ${filtered.length} IPs`);
  } else {
    if (freeOnly) {
      filtered = ipVars.filter(v => (v.used_in || []).length === 0);
      log(`✅ vlan mode (freeOnly) → ${filtered.length} IPs`);
    } else if (interfaceName) {
      filtered = ipVars.filter(v => {
        const usedIn = v.used_in || [];
        return !usedIn.some(u => u && u.name === interfaceName);
      });
      log(`✅ vlan mode (exclude already on '${interfaceName}') → ${filtered.length} IPs`);
    } else {
      filtered = ipVars.slice();
      log(`✅ vlan mode (no parent provided) → ${filtered.length} IPs`);
    }
  }

  // ---- 5) Return stable, simple array shape ---------------------------------
  const result = filtered.map(v => ({ name: v.name, value: v.value }));
  log(`📊 Returning ${result.length} IP vars`);
  return result;
}


///=========================================================================================
function createSubinterface(vlanTag, description, varName, parentInterface, ipArray) {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();

  // Strip leading $ from name only for use in the subinterface name (required by SCM)
  const cleanParent = parentInterface.replace(/^\$/, '');
  const fullName = `${parentInterface}.${vlanTag}`; // e.g. myneflame.999

  const payload = {
    name: fullName,                                 // Must NOT include $
    "vlan-tag": parseInt(vlanTag, 10),              // Required tag
    "parent-interface": parentInterface,            // Can include $, as it's a variable ref
    comment: description || "",
    snippet: snippet
  };

  if (Array.isArray(ipArray) && ipArray.length > 0 ) {
    payload["layer3"] = {
      ip: ipArray.map(ip => ({ name: ip }))
    };
  }

  const url = "https://api.strata.paloaltonetworks.com/config/network/v1/layer3-subinterfaces";
  logVlan(`🆕 Creating VLAN (L3) subinterface: ${JSON.stringify(payload)}`);
  logVlan(`📤 POST URL: ${url}`);

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      headers,
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const body = response.getContentText();
    logVlan(`📥 VLAN create response [${status}]: ${body}`);

    if (status >= 200 && status < 300) {
      logVlan(`✅ VLAN ${fullName} created successfully`);
    } else {
      logVlan(`❌ Failed to create VLAN: ${body}`);
      throw new Error("VLAN creation failed: " + body);
    }

    return { status, body };
  } catch (err) {
    logVlan(`❌ Exception during VLAN creation: ${err.message}`);
    throw err;
  }
}

///=========================================================================================
function getAvailableInterfacesForVrf(vrfInterfaces) {
  const { snippet } = getBaseconfSnippetData();  // Fetch from BASECONF sheet
  const allInterfaces = fetchInterfaces(snippet);  // Reuse VLAN helper

  // Return only interfaces NOT in the current VRF
  const unused = allInterfaces.filter(iface => !vrfInterfaces.includes(iface.name));
  return unused.map(i => i.name);  // Return just names
}

///=========================================================================================
function loadLogicalRouters() {
  const { snippet } = getBaseconfSnippetData();
  const { token, headers } = getFreshToken();
  const url = `https://api.strata.paloaltonetworks.com/config/network/v1/logical-routers?snippet=${encodeURIComponent(snippet)}`;
  const options = {
    method: "get",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  };
  logVrf(`➡️ GET Logical Routers: ${url}`);
  try {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const body = response.getContentText();
    logVrf(`⬅️ [${code}] Logical Router list response`);
    logVrf(body);
    if (code !== 200) {
      throw new Error(`Failed to load logical routers: ${code}`);
    }
    const json = JSON.parse(body);
    if (!json || !Array.isArray(json)) {
      logVrf(":warning: Response is not an array. Returning as-is for inspection.");
      return json;
    }
    return json;
  } catch (err) {
    logVrf(`:x: Exception in loadLogicalRouters: ${err}`);
    throw err;
  }
}
///=========================================================================================
function createLogicalRouter(name, description) {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();
  const url = "https://api.strata.paloaltonetworks.com/config/network/v1/logical-routers";
  const payload = {
    name: name,
    snippet: snippet,
    comment: description || "",
    "routing-stack": "advanced"
  };
  logVrf(`:new: Creating Logical Router: ${JSON.stringify(payload)}`);
  logVrf(`:outbox_tray: POST URL: ${url}`);
  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      headers,
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    const body = response.getContentText();
    logVrf(`:inbox_tray: Create VRF response [${code}]: ${body}`);
    if (code >= 200 && code < 300) {
      logVrf(`✅ Logical Router '${name}' created`);
      // 👇 Clear input fields on the frontend (NEW):
      const html = `
        <script>
          document.getElementById('logicalRouterName').value = "";
          document.getElementById('logicalRouterDesc').value = "";
        </script>`;
      SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), "Clear Fields");
      return { success: true, code, body };
    } else {
      logVrf(`:x: Failed to create Logical Router '${name}': ${code}`);
      throw new Error(`Create VRF failed: ${body}`);
    }
  } catch (err) {
    logVrf(`:x: Exception in createLogicalRouter(): ${err.message}`);
    throw err;
  }
}
///=========================================================================================
function getLogicalRouterById(uuid, headers) {
  const url = `https://api.strata.paloaltonetworks.com/config/network/v1/logical-routers/${uuid}`;
  logVrf(`📤 Fetch full Logical Router: ${url}`);

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers,
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  logVrf(`📥 Full router response [${code}]: ${body}`);

  if (code < 200 || code >= 300) {
    throw new Error(`❌ Failed to fetch full logical router [${code}]: ${body}`);
  }

  return JSON.parse(body);
}

///=========================================================================================
function fetchVrfs(snippet) {
  const { token, headers } = getFreshToken();
  const url = `https://api.strata.paloaltonetworks.com/config/network/v1/logical-routers?snippet=${encodeURIComponent(snippet)}`;

  logVrf(`➡️ GET VRFs for snippet: ${snippet}`);
  logVrf(`🔗 URL: ${url}`);

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers,
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const body = response.getContentText();

    logVrf(`⬅️ [${code}] VRF list response`);
    logVrf(body);

    if (code !== 200) {
      throw new Error(`Failed to fetch VRFs: ${code}`);
    }

    const json = JSON.parse(body);
    return Array.isArray(json) ? json : (json.data || []);
  } catch (err) {
    logVrf(`❌ Exception in fetchVrfs(): ${err.message}`);
    throw err;
  }
}
///=====================================================================================

///=========================================================================================
function addInterfacesToVrf(vrfName, interfaceNames) {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();
  logVrf(`🔧 Adding interfaces to VRF '${vrfName}': ${interfaceNames.join(', ')}`);

  // Step 1: Find logical router by name
  const allRouters = fetchVrfs(snippet);
  logVrf(`🔍 Searching for VRF name: "${vrfName}" in ${allRouters.length} routers`);
  allRouters.forEach(r => logVrf(`- Candidate: "${r.name}"`));

  const match = allRouters.find(r =>
    r.name.trim().toLowerCase() === vrfName.trim().toLowerCase()
  );
  if (!match || !match.id) throw new Error(`❌ VRF '${vrfName}' not found or missing ID`);
  logVrf(`✅ Matched router UUID: ${match.id}`);

  // Step 2: Fetch full structure via helper
  const router = getLogicalRouterById(match.id, headers);

  // Step 3: Merge interfaces into vrf[0]
  if (!router.vrf || !Array.isArray(router.vrf) || router.vrf.length === 0) {
    router.vrf = [{ name: "default", interface: [] }];
    logVrf(`🆕 Created empty VRF block`);
  }

  const currentIfs = router.vrf[0].interface || [];
  const combined = Array.from(new Set([...currentIfs, ...interfaceNames]));
  router.vrf[0].interface = combined;

  // Step 4: PUT update
  const urlPut = `https://api.strata.paloaltonetworks.com/config/network/v1/logical-routers/${match.id}`;
  const payload = JSON.stringify(router, null, 2);
  logVrf(`📤 PUT full router with updated interfaces`);
  logVrf(payload);

  const putResponse = UrlFetchApp.fetch(urlPut, {
    method: "put",
    headers,
    contentType: "application/json",
    payload,
    muteHttpExceptions: true
  });

  const putCode = putResponse.getResponseCode();
  const putBody = putResponse.getContentText();
  logVrf(`📥 PUT response [${putCode}]: ${putBody}`);

  if (putCode >= 200 && putCode < 300) {
    logVrf(`✅ Interfaces successfully added to VRF '${vrfName}'`);
    return { success: true, added: interfaceNames };
  } else {
    throw new Error(`❌ Failed to add interfaces: ${putCode} ${putBody}`);
  }
}
///==============================================================
function removeInterfacesFromVrf(vrfName, interfacesToRemove) {
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();
  logVrf(`🗑️ Removing interfaces from VRF '${vrfName}': ${interfacesToRemove.join(", ")}`);

  // Step 1: Find logical router by name
  const allRouters = fetchVrfs(snippet);
  logVrf(`🔍 Searching for VRF name: "${vrfName}" in ${allRouters.length} routers`);
  allRouters.forEach(r => logVrf(`- Candidate: "${r.name}"`));

  const match = allRouters.find(r =>
    r.name.trim().toLowerCase() === vrfName.trim().toLowerCase()
  );
  if (!match || !match.id) throw new Error(`❌ VRF '${vrfName}' not found or missing ID`);
  logVrf(`✅ Matched router UUID: ${match.id}`);

  // Step 2: Fetch full structure via helper
  const router = getLogicalRouterById(match.id, headers);

  // Step 3: Remove interfaces from vrf[0]
  if (!router.vrf || !Array.isArray(router.vrf) || router.vrf.length === 0) {
    throw new Error(`❌ VRF block missing in router '${vrfName}'`);
  }

  const currentIfs = router.vrf[0].interface || [];
  const updatedIfs = currentIfs.filter(i => !interfacesToRemove.includes(i));
  router.vrf[0].interface = updatedIfs;
  logVrf(`🧹 Updated interface list: ${updatedIfs.join(', ') || "(none left)"}`);

  // Step 4: PUT full payload
  const urlPut = `https://api.strata.paloaltonetworks.com/config/network/v1/logical-routers/${match.id}`;
  const payload = JSON.stringify(router, null, 2);
  logVrf(`📤 PUT full router with updated interface list`);
  logVrf(payload);

  const putResponse = UrlFetchApp.fetch(urlPut, {
    method: "put",
    headers,
    contentType: "application/json",
    payload,
    muteHttpExceptions: true
  });

  const putCode = putResponse.getResponseCode();
  const putBody = putResponse.getContentText();
  logVrf(`📥 PUT response [${putCode}]: ${putBody}`);

  if (putCode >= 200 && putCode < 300) {
    logVrf(`✅ Interfaces successfully removed from VRF '${vrfName}'`);
    return { success: true, removed: interfacesToRemove };
  } else {
    throw new Error(`❌ Failed to remove interfaces: ${putCode} ${putBody}`);
  }
}

///=========================================================================================
function deleteLogicalRouterById(uuid) {
  const { headers } = getFreshToken();
  const url = `https://api.strata.paloaltonetworks.com/config/network/v1/logical-routers/${uuid}`;

  logVrf(`🗑 Deleting logical router ID: ${uuid}`);
  logVrf(`📤 DELETE ${url}`);

  const res = UrlFetchApp.fetch(url, {
    method: "delete",
    headers,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  logVrf(`📥 Delete response [${code}]: ${body}`);

  if (code < 200 || code >= 300) {
    throw new Error(`Failed to delete VRF: ${body}`);
  }

  logVrf(`✅ VRF with UUID ${uuid} deleted`);
}

///=========================================================================================
function removeStaticRoutesFromVrf(vrfName, routeNamesToRemove) {
  if (!vrfName) throw new Error("❌ vrfName is required");
  if (!Array.isArray(routeNamesToRemove) || routeNamesToRemove.length === 0) {
    throw new Error("❌ routeNamesToRemove must be a non-empty array");
  }

  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();

  logVrf(`🗑️ Removing static routes from VRF '${vrfName}': ${routeNamesToRemove.join(", ")}`);

  // 1) Find router by name in this snippet
  const allRouters = fetchVrfs(snippet);
  const match = allRouters.find(r => (r.name || "").trim().toLowerCase() === vrfName.trim().toLowerCase());
  if (!match?.id) throw new Error(`❌ VRF '${vrfName}' not found or missing ID`);
  logVrf(`✅ Matched router UUID: ${match.id}`);

  // 2) Fetch full router
  const router = getLogicalRouterById(match.id, headers);

  // 3) Get static routes array safely
  const vrf0 = router?.vrf?.[0];
  if (!vrf0) throw new Error(`❌ VRF block missing in router '${vrfName}'`);

  const staticRoutes = vrf0?.routing_table?.ip?.static_route || [];
  if (!Array.isArray(staticRoutes)) throw new Error("❌ Unexpected static_route structure (not an array)");

  const toRemove = new Set(routeNamesToRemove.map(String));
  const keptRoutes = [];
  const removedNames = [];

  staticRoutes.forEach(rt => {
    if (toRemove.has(rt?.name)) removedNames.push(rt.name);
    else keptRoutes.push(rt);
  });

  if (removedNames.length === 0) {
    logVrf(`ℹ️ No matching routes found to remove in '${vrfName}'.`);
    return { success: true, removed: [], remainingCount: staticRoutes.length };
  }

  // 4) Put updated list back into payload
  if (!vrf0.routing_table) vrf0.routing_table = {};
  if (!vrf0.routing_table.ip) vrf0.routing_table.ip = {};
  vrf0.routing_table.ip.static_route = keptRoutes;

  // 5) PUT full router
  const urlPut = `https://api.strata.paloaltonetworks.com/config/network/v1/logical-routers/${match.id}`;
  const payload = JSON.stringify(router);

  logVrf(`📤 PUT full router with ${keptRoutes.length} remaining static route(s)`);
  const putResponse = UrlFetchApp.fetch(urlPut, {
    method: "put",
    headers,
    contentType: "application/json",
    payload,
    muteHttpExceptions: true
  });

  const putCode = putResponse.getResponseCode();
  const putBody = putResponse.getContentText();
  logVrf(`📥 PUT response [${putCode}]: ${putBody}`);

  if (putCode >= 200 && putCode < 300) {
    logVrf(`✅ Removed ${removedNames.length} route(s): ${removedNames.join(", ")}`);
    return { success: true, removed: removedNames, remainingCount: keptRoutes.length };
  }

  throw new Error(`❌ Failed to remove routes: ${putCode} ${putBody}`);
}

///=========================================================================================
function createStaticRouteForVrf(vrfName, input, optionsOverwrite) {
  const overwrite = optionsOverwrite === true;
  const { snippet } = getBaseconfSnippetData();
  const { headers } = getFreshToken();

  // --- Locate the logical router by VRF name (your existing pattern)
  logVrf(`➕ Adding static route to VRF '${vrfName}'`);
  const allRouters = fetchVrfs(snippet);
  const match = allRouters.find(r =>
    r.name && r.name.trim().toLowerCase() === vrfName.trim().toLowerCase()
  );
  if (!match || !match.id) {
    throw new Error(`❌ VRF '${vrfName}' not found or missing ID`);
  }
  logVrf(`✅ Matched router UUID: ${match.id}`);

  // --- Pull full LR
  const router = getLogicalRouterById(match.id, headers);

  // --- Make sure the VRF scaffolding exists
  if (!router.vrf || !Array.isArray(router.vrf) || router.vrf.length === 0) {
    router.vrf = [{ name: "default", interface: [] }];
    logVrf(`🧱 Created VRF[0] skeleton`);
  }
  const vrf0 = router.vrf[0];

  // --- (Optional) Validate interface is part of the VRF when provided
  if (input.iface) {
    const vrfIfaces = vrf0.interface || [];
    if (!vrfIfaces.includes(input.iface)) {
      throw new Error(`❌ Interface '${input.iface}' is not assigned to VRF '${vrfName}'. Add it first.`);
    }
  }

  // --- Ensure routing table scaffolding exists
  vrf0.routing_table = vrf0.routing_table || {};
  vrf0.routing_table.ip = vrf0.routing_table.ip || {};
  vrf0.routing_table.ip.static_route = vrf0.routing_table.ip.static_route || [];

  // --- Build the static route object from dialog input
  const routeObj = {
    name: (input.name || "").trim(),
    destination: (input.destination || "").trim(),
    metric: Number(input.metric) || 10,
    path_monitor: { enable: String(input.monitor).toLowerCase() === "yes" || input.monitor === true }
  };

  if (input.iface) routeObj.interface = input.iface;

  // Next hop mapping
  switch (input.nextHopType) {
    case "ip":
      if (!input.nextHop) throw new Error("Next Hop IP is required.");
      routeObj.nexthop = { ip_address: String(input.nextHop).trim() };
      break;
    case "fqdn":
      if (!input.nextHop) throw new Error("Next Hop FQDN is required.");
      routeObj.nexthop = { fqdn: String(input.nextHop).trim() };
      break;
    case "next_lr":
      if (!input.nextHop) throw new Error("Next Router name is required.");
      routeObj.nexthop = { next_lr: String(input.nextHop).trim() };
      break;
    case "discard":
      routeObj.nexthop = { discard: true };
      break;
    case "none":
    default:
      // No nexthop field
      break;
  }

  // Basic input checks
  if (!routeObj.name) throw new Error("Route Name is required.");
  if (!routeObj.destination) throw new Error("Destination is required (CIDR or variable).");

  // --- Upsert or prevent duplicate
  const list = vrf0.routing_table.ip.static_route;
  const idx = list.findIndex(r => (r.name || "").toLowerCase() === routeObj.name.toLowerCase());

  if (idx >= 0 && !overwrite) {
    throw new Error(`❌ A route named '${routeObj.name}' already exists. Use overwrite to replace it.`);
  }
  if (idx >= 0 && overwrite) {
    logVrf(`♻️ Overwriting existing route '${routeObj.name}'`);
    list[idx] = routeObj;
  } else {
    list.push(routeObj);
  }

  // --- PUT the full LR back
  const urlPut = `https://api.strata.paloaltonetworks.com/config/network/v1/logical-routers/${match.id}`;
  const payload = JSON.stringify(router, null, 2);
  logVrf(`📤 PUT logical router with new static route '${routeObj.name}'`);
  logVrf(payload);

  const putResponse = UrlFetchApp.fetch(urlPut, {
    method: "put",
    headers,
    contentType: "application/json",
    payload,
    muteHttpExceptions: true
  });

  const putCode = putResponse.getResponseCode();
  const putBody = putResponse.getContentText();
  logVrf(`📥 PUT response [${putCode}]: ${putBody}`);

  if (putCode < 200 || putCode >= 300) {
    throw new Error(`❌ Failed to add static route: ${putCode} ${putBody}`);
  }

  logVrf(`✅ Static route '${routeObj.name}' ${idx >= 0 ? "updated" : "added"} in VRF '${vrfName}'`);
  return {
    success: true,
    added: idx >= 0 ? undefined : routeObj.name,
    updated: idx >= 0 ? routeObj.name : undefined,
    total: list.length
  };
}


function getL3InterfacesWithIpForSnatDhcp() {
  logSnatDhcp("📡 getL3InterfacesWithIpForSnatDhcp() called");
  // TODO: replace with real fetch later.
  return [];
}
// Optional: remember the user’s selection 
function setSnatDhcpSelectedNic(nicName) {
  PropertiesService.getUserProperties().setProperty("SNAT_DHCP_SELECTED_NIC", nicName || "");
  logSnatDhcp(`🔖 Selected NIC set to: ${nicName || "(cleared)"}`);
  return true;
}
