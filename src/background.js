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

async function _updateActionState (nTabID, aMessage) {
  let bEnable = false;
  if (aMessage) {
    try {
      bEnable = (await _findCalendarText (aMessage.id)) !== null;
    } catch (ex) {
      console.error ("Failed to inspect message " + aMessage.id, ex);
    }
  }
  if (bEnable)
    await browser.messageDisplayAction.enable (nTabID);
  else
    await browser.messageDisplayAction.disable (nTabID);
}

// Disabled unless the displayed message is known to contain a calendar part
browser.messageDisplayAction.disable ();

browser.messageDisplay.onMessagesDisplayed.addListener (async (aTab, aDisplayedMessages) => {
  const aMessages = aDisplayedMessages.messages;
  await _updateActionState (aTab.id, aMessages.length === 1 ? aMessages[0] : null);
});

browser.messageDisplayAction.onClicked.addListener (async aTab => {
  const aMessage = await _getSingleDisplayedMessage (aTab.id);
  if (!aMessage)
    return;
  const sCalendarText = await _findCalendarText (aMessage.id);
  if (sCalendarText === null)
    return;

  const aSettings = await browser.storage.local.get (DEFAULT_SETTINGS);
  const sICS = buildICS (sCalendarText, aSettings.teamsURLAsLocation);

  const sObjectURL = URL.createObjectURL (new Blob ([sICS], { type: "text/calendar;charset=utf-8" }));
  try {
    const nDownloadID = await browser.downloads.download ({
      url: sObjectURL,
      filename: buildFilename (aMessage.subject),
      saveAs: true
    });
    s_aObjectURLs.set (nDownloadID, sObjectURL);
  } catch (ex) {
    // Also thrown when the user cancels the "Save as" dialog
    URL.revokeObjectURL (sObjectURL);
  }
});

browser.downloads.onChanged.addListener (aDelta => {
  const sState = aDelta.state?.current;
  if (s_aObjectURLs.has (aDelta.id) && (sState === "complete" || sState === "interrupted")) {
    URL.revokeObjectURL (s_aObjectURLs.get (aDelta.id));
    s_aObjectURLs.delete (aDelta.id);
  }
});
