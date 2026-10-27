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

// Pure iCalendar text helpers - no WebExtension API usage, so they can be tested outside Thunderbird.

const ICS_MAX_LINE_LENGTH = 75;

// Normalise all line endings to CRLF as required by RFC 5545
function normaliseLineEndings (sICS) {
  return sICS.replace (/\r\n|\r|\n/g, "\r\n");
}

// Remove everything outside of the first BEGIN:VCALENDAR ... END:VCALENDAR block
function extractVCalendar (sICS) {
  const aMatch = sICS.match (/BEGIN:VCALENDAR[\s\S]*?END:VCALENDAR/);
  return aMatch ? aMatch[0] : sICS;
}

// Join folded content lines (RFC 5545, section 3.1)
function unfold (sICS) {
  return sICS.replace (/\r?\n[ \t]/g, "");
}

// Fold a single content line at 75 characters (RFC 5545, section 3.1)
function foldLine (sLine) {
  if (sLine.length <= ICS_MAX_LINE_LENGTH)
    return sLine;
  const aParts = [sLine.substring (0, ICS_MAX_LINE_LENGTH)];
  for (let nIndex = ICS_MAX_LINE_LENGTH; nIndex < sLine.length; nIndex += ICS_MAX_LINE_LENGTH - 1)
    aParts.push (" " + sLine.substring (nIndex, nIndex + ICS_MAX_LINE_LENGTH - 1));
  return aParts.join ("\r\n");
}

// Escape a TEXT property value (RFC 5545, section 3.3.11)
function escapeText (sValue) {
  return sValue.replace (/\\/g, "\\\\").replace (/;/g, "\\;").replace (/,/g, "\\,").replace (/\r?\n/g, "\\n");
}

// Find the Teams join URL - preferably from the dedicated Microsoft property, otherwise from the text
function findTeamsMeetingURL (sICS) {
  const sUnfolded = unfold (sICS);
  const aProperty = sUnfolded.match (/^X-MICROSOFT-SKYPETEAMSMEETINGURL[^:\r\n]*:(https:\/\/\S+)/m);
  if (aProperty)
    return aProperty[1];
  const aInText = sUnfolded.match (/https:\/\/teams\.(?:microsoft|live)\.com\/[^\s<>"\\]+/);
  return aInText ? aInText[0] : null;
}

// Replace the first LOCATION property with the Teams join URL, or add one to the first VEVENT if there is none
function replaceLocationWithTeamsURL (sICS) {
  const sURL = findTeamsMeetingURL (sICS);
  if (!sURL)
    return sICS;
  const sNewLine = foldLine ("LOCATION:" + escapeText (sURL));
  const rLocation = /^LOCATION[;:].*(?:\r\n[ \t].*)*/m;
  if (rLocation.test (sICS))
    return sICS.replace (rLocation, () => sNewLine);
  return sICS.replace (/^BEGIN:VEVENT\r\n/m, aMatch => aMatch + sNewLine + "\r\n");
}

// Full conversion pipeline from the text of a text/calendar part to the content of the .ics file
function buildICS (sCalendarText, bTeamsURLAsLocation) {
  let sICS = extractVCalendar (normaliseLineEndings (sCalendarText));
  if (bTeamsURLAsLocation)
    sICS = replaceLocationWithTeamsURL (sICS);
  return sICS.endsWith ("\r\n") ? sICS : sICS + "\r\n";
}

// Build a safe file name from the message subject
function buildFilename (sSubject) {
  const sClean = (sSubject || "").replace (/[\/\\:*?"<>|\x00-\x1f]/g, " ").replace (/\s+/g, " ").trim ().substring (0, 150);
  return (sClean || "invitation") + ".ics";
}

if (typeof module !== "undefined")
  module.exports = { normaliseLineEndings, extractVCalendar, unfold, foldLine, escapeText, findTeamsMeetingURL, replaceLocationWithTeamsURL, buildICS, buildFilename };
