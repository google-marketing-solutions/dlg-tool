/**
 * Copyright 2026 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * @fileoverview Dependency-free base64 decoder (`decodeBase64ToAscii`). Apps Script V8 has no `atob` or
 * `Buffer`, and `Utilities.base64Decode` is unavailable in unit tests, so this pure implementation is used
 * to decode Google Ads recommendation IDs. Shared by both `dist/ads/Code.gs` and `dist/sheet/Code.gs`.
 */

const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

/** Maps each base64 character (standard `+/` and URL-safe `-_` alphabets) to its 6-bit value. */
const BASE64_LOOKUP: Readonly<Record<string, number>> = (() => {
  const lookup: Record<string, number> = {'-': 62, '_': 63};
  for (let i = 0; i < BASE64_ALPHABET.length; i++) {
    lookup[BASE64_ALPHABET[i]] = i;
  }
  return lookup;
})();

/**
 * Decodes a base64 string (standard or URL-safe alphabet, padding optional) into printable ASCII text.
 * Returns `null` for anything that is not valid base64 or does not decode to printable ASCII
 * (`0x20`–`0x7E`), so callers can fall back safely on synthetic or garbage IDs.
 */
export function decodeBase64ToAscii(input: string): string | null {
  if (typeof input !== 'string') return null;
  const data = input.replace(/={1,2}$/, '');
  // A single leftover character carries only 6 bits, which can never encode a whole byte
  if (data.length % 4 === 1) return null;
  if (data.length !== input.length && input.length % 4 !== 0) return null;

  let out = '';
  let buffer = 0;
  let bits = 0;
  for (let i = 0; i < data.length; i++) {
    const ch = data[i];
    if (!Object.prototype.hasOwnProperty.call(BASE64_LOOKUP, ch)) return null;
    buffer = (buffer << 6) | BASE64_LOOKUP[ch];
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      const byte = (buffer >> bits) & 0xff;
      buffer &= (1 << bits) - 1;
      if (byte < 0x20 || byte > 0x7e) return null;
      out += String.fromCharCode(byte);
    }
  }
  return out;
}
