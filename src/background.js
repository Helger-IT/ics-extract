/*
 * Copyright (C) 2026 Philip Helger (www.helger.com)
 * philip[at]helger[dot]com
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *         http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

const MENU_OPEN = "open-ics";
const MENU_SAVE = "save-ics";

// Fixed name of the file that is prepared in the Downloads folder for "Open .ics" - overwritten for every invitation
const OPEN_FILENAME = "temp-ics-extract-invitation.ics";

// localStorage keys - localStorage is used, because it can be read synchronously inside a user input handler
const LS_PREPARED = "prepared";
const LS_TAB_PREFIX = "tab-";

// Object URLs of running downloads, keyed by download ID
const s_aObjectURLs = new Map ();

function _isCalendarContentType (sContentType) {
  const sCT = (sContentType || "").toLowerCase ();
  return sCT.startsWith ("text/calendar") || sCT.startsWith ("application/ics");
}

// Find the text of the first calendar part of a message, or null if there is none
async function _findCalendarText (nMessageID) {
  // Inline parts - this is where Outlook and Teams put the invitation
  const aInlineParts = await browser.messages.listInlineTextParts (nMessageID);
  const aInline = aInlineParts.find (aPart => _isCalendarContentType (aPart.contentType));
  if (aInline)
    return aInline.content;

  // Real attachments with a calendar content type
  const aAttachments = await browser.messages.listAttachments (nMessageID);
  const aAttachment = aAttachments.find (aAtt => _isCalendarContentType (aAtt.contentType));
  if (aAttachment) {
    const aFile = await browser.messages.getAttachmentFile (nMessageID, aAttachment.partName);
    return aFile.text ();
  }
  return null;
}

// Returns the single displayed message of a tab, or null if none or several are displayed
async function _getSingleDisplayedMessage (nTabID) {
  const aList = await browser.messageDisplay.getDisplayedMessages (nTabID);
  return aList.messages.length === 1 ? aList.messages[0] : null;
}

async function _createICS (sCalendarText) {
  const aSettings = await browser.storage.local.get (DEFAULT_SETTINGS);
  return buildICS (sCalendarText, aSettings.teamsURLAsLocation);
}

// Start a download of the given .ics content and return its ID - throws if the user cancels the "Save as" dialog
async function _downloadICS (sICS, sFilename, bSaveAs, sConflictAction) {
  const sObjectURL = URL.createObjectURL (new Blob ([sICS], { type: "text/calendar;charset=utf-8" }));
  try {
    const nDownloadID = await browser.downloads.download ({
      url: sObjectURL,
      filename: sFilename,
      saveAs: bSaveAs,
      conflictAction: sConflictAction
    });
    s_aObjectURLs.set (nDownloadID, sObjectURL);
    return nDownloadID;
  } catch (ex) {
    URL.revokeObjectURL (sObjectURL);
    throw ex;
  }
}

function _readPrepared () {
  try {
    return JSON.parse (localStorage.getItem (LS_PREPARED));
  } catch (ex) {
    return null;
  }
}

// Write the .ics of a message to the fixed file, so that "Open .ics" can open it synchronously
async function _prepareOpenFile (nTabID, aMessage, sCalendarText) {
  localStorage.setItem (LS_TAB_PREFIX + nTabID, String (aMessage.id));
  const nDownloadID = await _downloadICS (await _createICS (sCalendarText), OPEN_FILENAME, false, "overwrite");
  const aPrevious = _readPrepared ();
  localStorage.setItem (LS_PREPARED, JSON.stringify ({ downloadID: nDownloadID, messageID: aMessage.id }));
  // Keep the list of saved files clean - this does not delete the file itself
  if (aPrevious && aPrevious.downloadID !== nDownloadID)
    await browser.downloads.erase ({ id: aPrevious.downloadID });
}

// bForcePrepare: prepare the file for "Open .ics" even if pre-writing on display is disabled
async function _updateActionState (nTabID, aMessage, bForcePrepare) {
  let sCalendarText = null;
  if (aMessage) {
    try {
      sCalendarText = await _findCalendarText (aMessage.id);
    } catch (ex) {
      console.error ("Failed to inspect message " + aMessage.id, ex);
    }
  }
  if (sCalendarText === null) {
    localStorage.removeItem (LS_TAB_PREFIX + nTabID);
    await browser.messageDisplayAction.disable (nTabID);
    return;
  }
  await browser.messageDisplayAction.enable (nTabID);
  const aSettings = await browser.storage.local.get (DEFAULT_SETTINGS);
  if (!bForcePrepare && !aSettings.prepareOpenFile)
    return;
  try {
    await _prepareOpenFile (nTabID, aMessage, sCalendarText);
  } catch (ex) {
    console.error ("Failed to prepare the .ics of message " + aMessage.id, ex);
  }
}

// Must be called synchronously from the menu click - downloads.open only works inside a user input handler
function _openPrepared (aTab) {
  const aPrepared = _readPrepared ();
  const sTabMessageID = localStorage.getItem (LS_TAB_PREFIX + aTab.id);
  if (!aPrepared || String (aPrepared.messageID) !== sTabMessageID) {
    // Not prepared on display, or another tab displayed an invitation in the meantime - the next click opens it
    _getSingleDisplayedMessage (aTab.id).then (aMessage => _updateActionState (aTab.id, aMessage, true));
    return;
  }
  browser.downloads.open (aPrepared.downloadID).catch (ex => {
    console.error ("Failed to open download " + aPrepared.downloadID, ex);
  });
}

// Save the .ics of the displayed message via "Save as" dialog
async function _saveAs (aTab) {
  const aMessage = await _getSingleDisplayedMessage (aTab.id);
  if (!aMessage)
    return;
  const sCalendarText = await _findCalendarText (aMessage.id);
  if (sCalendarText === null)
    return;
  try {
    await _downloadICS (await _createICS (sCalendarText), buildFilename (aMessage.subject), true, "uniquify");
  } catch (ex) {
    // Also thrown when the user cancels the "Save as" dialog
  }
}

// Disabled unless the displayed message is known to contain a calendar part
browser.messageDisplayAction.disable ();

browser.runtime.onStartup.addListener (() => {
  // Message, tab and download IDs of the previous session are meaningless
  localStorage.clear ();
});

browser.runtime.onInstalled.addListener (async () => {
  localStorage.clear ();
  // Menu entries survive an update or reload of the extension
  await browser.menus.removeAll ();
  browser.menus.create ({ id: MENU_OPEN, title: browser.i18n.getMessage ("menuOpen"), contexts: ["message_display_action_menu"] });
  browser.menus.create ({ id: MENU_SAVE, title: browser.i18n.getMessage ("menuSave"), contexts: ["message_display_action_menu"] });
});

browser.messageDisplay.onMessagesDisplayed.addListener (async (aTab, aDisplayedMessages) => {
  const aMessages = aDisplayedMessages.messages;
  await _updateActionState (aTab.id, aMessages.length === 1 ? aMessages[0] : null, false);
});

// Deliberately not async - see _openPrepared
browser.menus.onClicked.addListener ((aInfo, aTab) => {
  if (aInfo.menuItemId === MENU_OPEN)
    _openPrepared (aTab);
  else if (aInfo.menuItemId === MENU_SAVE)
    _saveAs (aTab);
});

browser.downloads.onChanged.addListener (aDelta => {
  const sState = aDelta.state?.current;
  if (s_aObjectURLs.has (aDelta.id) && (sState === "complete" || sState === "interrupted")) {
    URL.revokeObjectURL (s_aObjectURLs.get (aDelta.id));
    s_aObjectURLs.delete (aDelta.id);
  }
});
