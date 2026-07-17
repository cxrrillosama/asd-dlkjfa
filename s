
// Single source of truth for the backing spreadsheet — avoids getActiveSpreadsheet() context loss in web apps.
const SPREADSHEET_ID = "1EikuuKnQJeCLLsy8-TCxokIC_TZ2EzaKBWFluXtaBQ4";

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('GE Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Generic 2-column tab reader. Optional ss lets getStartupBundle reuse one open spreadsheet handle.
function fetchTwoColumnData(sheetName, keyName, valueName, ss) {
  try {
    const spreadsheet = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    return data.slice(1).map(row => ({
      [keyName]: row[0] ? row[0].toString().trim() : '',
      [valueName]: row[1] ? row[1].toString().trim() : ''
    })).filter(item => item[keyName] !== '' && item[valueName] !== '');
  } catch(e) {
    console.error(`Error fetching data from ${sheetName}: `, e);
    return [];
  }
}

// Everything the client needs at load, in ONE google.script.run round trip (opens the spreadsheet once).
// Client falls back to individual fetches if this fails, so a bad reader can't blank the whole app.
function getStartupBundle() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    fdPresets: getSpreadsheetPresets(ss),
    workNotesPresets: getWorkNotesPresets(ss),
    threadRequestsPresets: getThreadRequestsPresets(ss),
    carrierPhonebook: getCarrierPhonebookData(ss),
    portingIssues: getPortingIssuesData(ss),
    fraudGuide: getFraudGuideData(ss)
  };
}

function getSpreadsheetPresets(ss) { return fetchTwoColumnData('Templates', 'name', 'text', ss); }
function getWorkNotesPresets(ss)   { return fetchTwoColumnData('Categories', 'name', 'text', ss); }

// 3-column reader: name, text, Chat Destination. Blank destination = note-only preset.
// Non-blank must match a "CHAT_WEBHOOK_<destination>" Script Property (URL stays server-side).
function getThreadRequestsPresets(ss) {
  try {
    const spreadsheet = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('ThreadRequests');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    return data.slice(1).map(row => ({
      name: row[0] ? row[0].toString().trim() : '',
      text: row[1] ? row[1].toString().trim() : '',
      chatDestination: row[2] ? row[2].toString().trim() : ''
    })).filter(item => item.name !== '' && item.text !== '');
  } catch (e) {
    console.error('Error fetching Thread Requests presets: ', e);
    return [];
  }
}

// Posts to the Chat space mapped via Script Properties. Any failure returns {success:false, error}
// so the client can fall back to inserting into the note instead of dropping the message.
function sendThreadRequestToChat(messageText, destinationKey) {
  try {
    if (!destinationKey) {
      return { success: false, error: 'No Chat destination specified for this template.' };
    }

    const webhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_' + destinationKey);
    if (!webhookUrl) {
      return { success: false, error: 'No webhook configured for "' + destinationKey + '".' };
    }

    let submitterEmail = 'Unknown';
    try {
      submitterEmail = Session.getActiveUser().getEmail() || 'Unknown';
    } catch (emailErr) {
      console.error('Could not resolve active user email (likely missing userinfo.email scope): ', emailErr);
    }

    const stampedText = 'Sent by: ' + submitterEmail + '\n\n' + messageText;

    const response = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: stampedText }),
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    if (status >= 200 && status < 300) return { success: true };

    console.error('Chat webhook returned status ' + status + ': ' + response.getContentText());
    return { success: false, error: 'Chat returned an error (status ' + status + ').' };
  } catch (e) {
    console.error('Error sending Thread Request to Chat: ', e);
    return { success: false, error: e.toString() };
  }
}

// "Phonebook" tab, columns A-H: Carrier, NPC Phone, Care Phone, Hours, Port PIN, PIN Expiration, Account Number, Notes.
function getCarrierPhonebookData(ss) {
  try {
    const spreadsheet = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('Phonebook');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    return data.slice(1).map(row => ({
      carrier: row[0] ? row[0].toString().trim() : '',
      npcPhone: row[1] ? row[1].toString().trim() : '',
      carePhone: row[2] ? row[2].toString().trim() : '',
      hours: row[3] ? row[3].toString().trim() : '',
      portPin: row[4] ? row[4].toString().trim() : '',
      portPinExpiration: row[5] ? row[5].toString().trim() : '',
      accountNumber: row[6] ? row[6].toString().trim() : '',
      notes: row[7] ? row[7].toString().trim() : ''
    })).filter(item => item.carrier !== '');
  } catch (e) {
    console.error('Error fetching carrier phonebook data: ', e);
    return [];
  }
}

// "PortingIssues" tab, one row per step, grouped/sorted by Category ID + Order columns.
// Columns A-K: Category ID, Category Order, Category Title, Error Code, Step Order, Step Text,
// Template Copy Text, Escalate, Escalation Description, Escalation Type, Escalation Priority.
// Step Text syntax (parsed client-side): "- line" renders a bullet; {{label->category-id}} renders a jump link.
function getPortingIssuesData(ss) {
  try {
    const spreadsheet = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('PortingIssues');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const categoriesById = {};
    const seenOrder = [];

    data.slice(1).forEach(row => {
      const catId = row[0] ? row[0].toString().trim() : '';
      if (!catId) return;

      if (!categoriesById[catId]) {
        const orderVal = row[1];
        categoriesById[catId] = {
          id: catId,
          order: (orderVal !== '' && orderVal !== null && !isNaN(orderVal)) ? Number(orderVal) : 999,
          title: row[2] ? row[2].toString().trim() : '',
          errorCode: row[3] ? row[3].toString().trim() : '',
          steps: []
        };
        seenOrder.push(catId);
      }

      const stepText = row[5] ? row[5].toString().trim() : '';
      const templateText = row[6] ? row[6].toString() : '';
      if (!stepText && !templateText) return;

      const stepOrderVal = row[4];
      categoriesById[catId].steps.push({
        order: (stepOrderVal !== '' && stepOrderVal !== null && !isNaN(stepOrderVal))
          ? Number(stepOrderVal) : categoriesById[catId].steps.length + 1,
        text: stepText,
        template: templateText,
        escalate: row[7] ? row[7].toString().trim().toLowerCase() === 'yes' : false,
        escalationDescription: row[8] ? row[8].toString().trim() : '',
        escalationType: row[9] ? row[9].toString().trim() : '',
        escalationPriority: row[10] ? row[10].toString().trim() : ''
      });
    });

    const categories = Object.keys(categoriesById).map(id => categoriesById[id]);
    categories.forEach(cat => cat.steps.sort((a, b) => a.order - b.order));
    categories.sort((a, b) => a.order - b.order);
    return categories;
  } catch (e) {
    console.error('Error fetching porting issues data: ', e);
    return [];
  }
}

// "FraudTS" tab, one row per block (content is deeply nested, so raw HTML rather than plain steps).
// Columns A-N: Section ID, Section Order, Title, Subtitle, Persistent (Yes = always-visible intro),
// Subsection Title (blank continues previous), Block Order, Block Type, Content, Block ID,
// Escalate Description, Escalate Type, Escalate Priority, Template Text.
// Block Types: html | callout-warning | callout-redirect | callout-danger | escalate-def | template-def.
// def rows never render in sequence — they supply data for [[ESCALATE:id]] / [[TEMPLATE:id]] tokens
// inside html blocks, letting those components render at an exact nested position.
// {{label->section-id}} renders a cross-reference jump link.
function getFraudGuideData(ss) {
  try {
    const spreadsheet = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('FraudTS');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const sectionsById = {};
    const seenOrder = [];

    data.slice(1).forEach(row => {
      const sectionId = row[0] ? row[0].toString().trim() : '';
      if (!sectionId) return;

      if (!sectionsById[sectionId]) {
        const orderVal = row[1];
        sectionsById[sectionId] = {
          id: sectionId,
          order: (orderVal !== '' && orderVal !== null && !isNaN(orderVal)) ? Number(orderVal) : 999,
          title: row[2] ? row[2].toString().trim() : '',
          subtitle: row[3] ? row[3].toString().trim() : '',
          persistent: row[4] ? row[4].toString().trim().toLowerCase() === 'yes' : false,
          subsections: [],
          defs: {}
        };
        seenOrder.push(sectionId);
      }
      const section = sectionsById[sectionId];

      const subsectionTitle = row[5] ? row[5].toString().trim() : '';
      const blockOrderVal = row[6];
      const blockType = row[7] ? row[7].toString().trim() : 'html';
      const content = row[8] ? row[8].toString() : '';
      const blockId = row[9] ? row[9].toString().trim() : '';
      const escDescription = row[10] ? row[10].toString().trim() : '';
      const escType = row[11] ? row[11].toString().trim() : '';
      const escPriority = row[12] ? row[12].toString().trim() : '';
      const templateText = row[13] ? row[13].toString() : '';

      if (blockType === 'escalate-def') {
        if (blockId) section.defs[blockId] = { type: 'escalate', description: escDescription, issueType: escType, priority: escPriority };
        return;
      }
      if (blockType === 'template-def') {
        if (blockId) section.defs[blockId] = { type: 'template', text: templateText };
        return;
      }

      let subsection = section.subsections.length ? section.subsections[section.subsections.length - 1] : null;
      if (subsectionTitle) {
        subsection = { title: subsectionTitle, blocks: [] };
        section.subsections.push(subsection);
      } else if (!subsection) {
        subsection = { title: '', blocks: [] };
        section.subsections.push(subsection);
      }

      subsection.blocks.push({
        order: (blockOrderVal !== '' && blockOrderVal !== null && !isNaN(blockOrderVal)) ? Number(blockOrderVal) : subsection.blocks.length + 1,
        type: blockType,
        content: content
      });
    });

    const sections = Object.keys(sectionsById).map(id => sectionsById[id]);
    sections.forEach(section => section.subsections.forEach(sub => sub.blocks.sort((a, b) => a.order - b.order)));
    sections.sort((a, b) => a.order - b.order);
    return sections;
  } catch (e) {
    console.error('Error fetching fraud guide data: ', e);
    return [];
  }
}

function getEscalationGuideData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Escalation Guide');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    return data.slice(1).map(row => ({
      escalate: row[0] ? row[0].toString().trim() : '',
      category: row[1] ? row[1].toString().trim() : '',
      issueType: row[2] ? row[2].toString().trim() : '',
      priority: row[3] ? row[3].toString().trim() : '',
      rtsEligible: row[4] ? row[4].toString().trim() : '',
      definition: row[5] ? row[5].toString().trim() : ''
    })).filter(item => item.issueType !== '');
  } catch (e) {
    console.error('Error fetching escalation guide data: ', e);
    return [];
  }
}

function getSharedTemplates() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Updates1');
    if (!sheet) return ["Error: 'Updates1' sheet tab missing."];

    const data = sheet.getDataRange().getValues();
    if (data.length < 1) return ["No updates found."];

    return data
      .map(row => row[0] ? row[0].toString().trim() : '')
      .filter(txt => txt !== '');
  } catch(e) {
    return ["Error synchronizing stream: " + e.toString()];
  }
}

// Google Doc powering the Updates Feed: paste updates (text/images) separated by blank lines,
// newest at top. Set this to the Doc's ID from its URL.
const UPDATES_DOC_ID = "REPLACE_WITH_YOUR_DOC_ID";

// Splits the doc into feed items on blank paragraphs. Reads positioned/floating images too —
// pasting sometimes creates those instead of inline images, and they aren't reachable via getChild().
function getUpdatesDocFeed() {
  try {
    const doc = DocumentApp.openById(UPDATES_DOC_ID);
    const body = doc.getBody();
    const numChildren = body.getNumChildren();

    const feedItems = [];
    let currentText = '';
    let currentImages = [];

    const flushCurrentItem = () => {
      if (currentText.trim() !== '' || currentImages.length > 0) {
        feedItems.push({ text: currentText.trim(), images: currentImages });
      }
      currentText = '';
      currentImages = [];
    };

    for (let i = 0; i < numChildren; i++) {
      const el = body.getChild(i);
      if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

      const paragraph = el.asParagraph();
      const paragraphText = paragraph.getText();
      const paragraphImages = [];

      const numPieces = paragraph.getNumChildren();
      for (let j = 0; j < numPieces; j++) {
        const piece = paragraph.getChild(j);
        if (piece.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
          try {
            const blob = piece.asInlineImage().getBlob();
            const base64Data = Utilities.base64Encode(blob.getBytes());
            paragraphImages.push(`data:${blob.getContentType()};base64,${base64Data}`);
          } catch (imgErr) {
            console.error('Failed to read an inline image from the Updates Doc: ', imgErr);
          }
        }
      }

      const positionedIds = paragraph.getPositionedImageIds ? paragraph.getPositionedImageIds() : [];
      positionedIds.forEach(id => {
        try {
          const posImg = paragraph.getPositionedImage(id);
          const blob = posImg.getBlob();
          const base64Data = Utilities.base64Encode(blob.getBytes());
          paragraphImages.push(`data:${blob.getContentType()};base64,${base64Data}`);
        } catch (imgErr) {
          console.error('Failed to read a positioned image from the Updates Doc: ', imgErr);
        }
      });

      const isBlankParagraph = paragraphText.trim() === '' && paragraphImages.length === 0;

      if (isBlankParagraph) {

        flushCurrentItem();
        continue;
      }

      if (paragraphText.trim() !== '') {
        currentText += (currentText ? '\n' : '') + paragraphText;
      }
      currentImages = currentImages.concat(paragraphImages);
    }
    flushCurrentItem();

    feedItems.reverse();

    return {
      success: true,
      messages: feedItems.map(item => ({ text: item.text, sender: '', timestamp: '', images: item.images }))
    };
  } catch (e) {
    console.error('Error reading the Updates Doc: ', e);
    return { success: false, error: e.toString(), messages: [] };
  }
}

// All user-property writes route through here so quota failures surface as {success:false, error}.
function saveUserProperty(key, value) {
  try {
    PropertiesService.getUserProperties().setProperty(key, value);
    return { success: true };
  } catch (e) {
    console.error(`Error saving property ${key}: `, e);
    return { success: false, error: `Failed to save. You may have reached the storage limit. Error: ${e.message}` };
  }
}

function getUserNotes() { return PropertiesService.getUserProperties().getProperty('GE_SAVED_NOTES') || '[]'; }
function saveUserNotes(jsonString) { return saveUserProperty('GE_SAVED_NOTES', jsonString); }

function getQuickResponses() { return PropertiesService.getUserProperties().getProperty('GE_QUICK_RESPONSES') || '[]'; }
function saveQuickResponses(jsonString) { return saveUserProperty('GE_QUICK_RESPONSES', jsonString); }

function getCustomFolders() { return PropertiesService.getUserProperties().getProperty('GE_CUSTOM_FOLDERS') || '[]'; }
function saveCustomFolders(jsonString) { return saveUserProperty('GE_CUSTOM_FOLDERS', jsonString); }

function loadVaultData() { return PropertiesService.getUserProperties().getProperty('GE_SECURE_VAULT') || ''; }
function saveVaultData(scrambledPayload) { return saveUserProperty('GE_SECURE_VAULT', scrambledPayload); }

// Salt MUST live server-side: same password + different salt = different key, so a
// localStorage salt makes the vault undecryptable from any other browser/device.
function getVaultSalt() { return PropertiesService.getUserProperties().getProperty('GE_VAULT_SALT') || ''; }
function saveVaultSalt(saltB64) { return saveUserProperty('GE_VAULT_SALT', saltB64); }

function getCustomPresets() { return PropertiesService.getUserProperties().getProperty('GE_CUSTOM_PRESETS') || '[]'; }
function saveCustomPresets(jsonString) { return saveUserProperty('GE_CUSTOM_PRESETS', jsonString); }

function getSignatures() { return PropertiesService.getUserProperties().getProperty('GE_SIGNATURES') || '{}'; }
function saveSignatures(jsonString) { return saveUserProperty('GE_SIGNATURES', jsonString); }

function getQuickLinks() { return PropertiesService.getUserProperties().getProperty('GE_QUICK_LINKS') || '[]'; }
function saveQuickLinks(jsonString) { return saveUserProperty('GE_QUICK_LINKS', jsonString); }

function clearAllUserData() {
  const props = PropertiesService.getUserProperties();
  props.deleteProperty('GE_SAVED_NOTES');
  props.deleteProperty('GE_QUICK_RESPONSES');
  props.deleteProperty('GE_CUSTOM_FOLDERS');
  props.deleteProperty('GE_SECURE_VAULT');
  props.deleteProperty('GE_VAULT_SALT');
  props.deleteProperty('GE_CUSTOM_PRESETS');
  props.deleteProperty('GE_SIGNATURES');
  props.deleteProperty('GE_QUICK_LINKS');
}

// LockService-guarded append. Backfills the Account Number column on sheets created before it existed.
function logCallTrendingData(trendingData) {
  const lock = LockService.getScriptLock();

  try {

    lock.waitLock(3000);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Call Trending");
    const headers = ["Timestamp", "Submitted By", "Account Number", "Call Type", "Trend Reason", "Resolution", "Additional Details / Notes"];

    let submitterEmail = "Unknown";
    try {
      submitterEmail = Session.getActiveUser().getEmail() || "Unknown";
    } catch (emailErr) {
      console.error("Could not resolve active user email (likely missing userinfo.email scope): ", emailErr);
    }

    if (!sheet) {
      sheet = ss.insertSheet("Call Trending");
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#26262b").setFontColor("#f5f5f7");
    } else {
      let lastCol = Math.max(sheet.getLastColumn(), 1);
      let currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

      if (currentHeaders.indexOf("Submitted By") === -1) {
        sheet.insertColumnAfter(1);
        sheet.getRange(1, 2, 1, 1)
          .setValue("Submitted By")
          .setFontWeight("bold").setBackground("#26262b").setFontColor("#f5f5f7");
        lastCol = sheet.getLastColumn();
        currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      }

      if (currentHeaders.indexOf("Account Number") === -1) {
        const submittedByCol = currentHeaders.indexOf("Submitted By") + 1;
        const insertAfterCol = submittedByCol > 0 ? submittedByCol : 1;
        sheet.insertColumnAfter(insertAfterCol);
        sheet.getRange(1, insertAfterCol + 1, 1, 1)
          .setValue("Account Number")
          .setFontWeight("bold").setBackground("#26262b").setFontColor("#f5f5f7");
        lastCol = sheet.getLastColumn();
        currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      }

      if (currentHeaders.indexOf("Resolution") === -1) {
        const trendReasonCol = currentHeaders.indexOf("Trend Reason") + 1;
        const insertAfterCol = trendReasonCol > 0 ? trendReasonCol : lastCol;
        sheet.insertColumnAfter(insertAfterCol);
        sheet.getRange(1, insertAfterCol + 1, 1, 1)
          .setValue("Resolution")
          .setFontWeight("bold").setBackground("#26262b").setFontColor("#f5f5f7");
      }
    }

    sheet.appendRow([
      new Date(),
      submitterEmail,
      trendingData.accountNumber || "",
      trendingData.callType,
      trendingData.reason,
      trendingData.resolution || "",
      trendingData.notes
    ]);

    return { success: true };
  } catch (error) {
    console.error("Error logging call trending data: ", error);
    return { success: false, error: error.toString() };
  } finally {

    lock.releaseLock();
  }
}

// LockService-guarded append. Stamps submitter email server-side; falls back to "Unknown"
// rather than blocking the submission if the userinfo.email scope is unavailable.
function submitFeedbackData(feedbackData) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(3000);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Feedback");
    const headers = ["Timestamp", "Type", "Rating", "Feedback", "Submitted By"];

    if (!sheet) {
      sheet = ss.insertSheet("Feedback");
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#26262b").setFontColor("#f5f5f7");
    }

    let submitterEmail = "Unknown";
    try {
      submitterEmail = Session.getActiveUser().getEmail() || "Unknown";
    } catch (emailErr) {
      console.error("Could not resolve active user email (likely missing userinfo.email scope): ", emailErr);
    }

    sheet.appendRow([
      new Date(),
      feedbackData.type || "",
      feedbackData.rating || "",
      feedbackData.message || "",
      submitterEmail
    ]);

    return { success: true };
  } catch (error) {
    console.error("Error saving feedback: ", error);
    return { success: false, error: error.toString() };
  } finally {
    lock.releaseLock();
  }
}
