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

async function _init () {
  for (const aElement of document.querySelectorAll ("[data-i18n]"))
    aElement.textContent = browser.i18n.getMessage (aElement.dataset.i18n);
  const aSettings = await browser.storage.local.get (DEFAULT_SETTINGS);
  for (const sKey of Object.keys (DEFAULT_SETTINGS)) {
    const aCheckbox = document.getElementById (sKey);
    aCheckbox.checked = aSettings[sKey];
    aCheckbox.addEventListener ("change", () => browser.storage.local.set ({ [sKey]: aCheckbox.checked }));
  }
}

document.addEventListener ("DOMContentLoaded", _init);
