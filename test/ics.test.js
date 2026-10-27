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

// Run with: node --test test/
const { test } = require ("node:test");
const assert = require ("node:assert");
const ics = require ("../src/ics.js");

const TEAMS_URL = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_ZjE2NDYwYTMtOTk1Zi00YjA0LWI5ZDEtNmQ0YjQ3ZWU1ZTEw%40thread.v2/0?context=%7b%22Tid%22%3a%22abc%22%7d";

function _teamsInvite (bWithProperty) {
  const aLines = ["BEGIN:VCALENDAR", "METHOD:REQUEST", "VERSION:2.0", "BEGIN:VEVENT",
                  "SUMMARY;LANGUAGE=de-DE:Weekly sync",
                  "DESCRIPTION;LANGUAGE=de-DE:Join here <" + TEAMS_URL.substring (0, 40),
                  " " + TEAMS_URL.substring (40) + ">\\n",
                  "LOCATION;LANGUAGE=de-DE:Microsoft Teams-Besprechung"];
  if (bWithProperty)
    aLines.push ("X-MICROSOFT-SKYPETEAMSMEETINGURL:" + TEAMS_URL.substring (0, 60), " " + TEAMS_URL.substring (60));
  aLines.push ("END:VEVENT", "END:VCALENDAR");
  return aLines.join ("\n");
}

test ("folded lines are at most 75 characters and unfold to the original", () => {
  const sLine = "LOCATION:" + TEAMS_URL;
  const sFolded = ics.foldLine (sLine);
  for (const sPart of sFolded.split ("\r\n"))
    assert.ok (sPart.length <= 75, sPart);
  assert.strictEqual (ics.unfold (sFolded), sLine);
});

test ("Teams URL from property and from description", () => {
  assert.strictEqual (ics.findTeamsMeetingURL (_teamsInvite (true)), TEAMS_URL);
  assert.strictEqual (ics.findTeamsMeetingURL (_teamsInvite (false)), TEAMS_URL);
  assert.strictEqual (ics.findTeamsMeetingURL ("BEGIN:VCALENDAR\r\nEND:VCALENDAR"), null);
});

test ("LOCATION is replaced when enabled and untouched when disabled", () => {
  const sOn = ics.buildICS (_teamsInvite (true), true);
  assert.ok (ics.unfold (sOn).includes ("\r\nLOCATION:" + TEAMS_URL + "\r\n"));
  assert.ok (!sOn.includes ("Teams-Besprechung"));
  assert.ok (!/[^\r]\n/.test (sOn), "only CRLF line endings");

  const sOff = ics.buildICS (_teamsInvite (true), false);
  assert.ok (sOff.includes ("LOCATION;LANGUAGE=de-DE:Microsoft Teams-Besprechung"));
});

test ("LOCATION is added if missing", () => {
  const sIn = _teamsInvite (false).replace (/\nLOCATION[^\n]*/, "");
  const sOut = ics.buildICS (sIn, true);
  assert.ok (ics.unfold (sOut).includes ("BEGIN:VEVENT\r\nLOCATION:" + TEAMS_URL + "\r\n"));
});

test ("text outside of VCALENDAR is removed", () => {
  const sOut = ics.buildICS ("garbage\nBEGIN:VCALENDAR\nVERSION:2.0\nEND:VCALENDAR\nmore garbage", true);
  assert.strictEqual (sOut, "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR\r\n");
});

test ("file names", () => {
  assert.strictEqual (ics.buildFilename ("Re: a/b \"c\"?"), "Re a b c.ics");
  assert.strictEqual (ics.buildFilename (""), "invitation.ics");
  assert.strictEqual (ics.buildFilename (undefined), "invitation.ics");
});
