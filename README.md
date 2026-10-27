# ics-extract

Thunderbird extension (Manifest V3, Thunderbird 128 or later) that saves the calendar invitation
contained in a message as an `.ics` file.

Outlook and Teams send meeting invitations as an inline `text/calendar` MIME part instead of a
regular attachment. When such a message is displayed, the "Save invitation as .ics" button in the
message header toolbar becomes active. Clicking it opens a "Save as" dialog for the `.ics` file.

Inspired by [thunderbird-outlook-appointment](https://github.com/Leworoe/thunderbird-outlook-appointment)
by Leon Wolfgang Rönnpagel.

## Features

* Finds inline `text/calendar` parts as well as `text/calendar` / `application/ics` attachments
* Uses the Thunderbird API for MIME decoding, so base64, quoted-printable and all charsets are supported
* Strips everything outside of `BEGIN:VCALENDAR` ... `END:VCALENDAR` and normalises line endings to CRLF
* Option (enabled by default): replace the `LOCATION` of Microsoft Teams invitations with the
  Teams join URL, so calendar apps show a clickable link

## Building

```sh
./build.sh          # creates dist/ics-extract-<version>.xpi
node --test test/   # unit tests of the iCalendar handling
```

For development, load `src/manifest.json` via "Add-ons and Themes" → gear icon → "Debug Add-ons"
→ "Load Temporary Add-on".

# News and noteworthy

v1.0.0 - work in progress
* Initial version
