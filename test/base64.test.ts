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
 */

import {describe, expect, it} from 'vitest';
import {decodeBase64ToAscii} from '../src/shared/utils/base64';

describe('decodeBase64ToAscii', () => {
  it('decodes a Google Ads recommendation ID segment', () => {
    expect(decodeBase64ToAscii('NDQ1NTY2Nzc4OC0xMy0xNzAwMDAwMDAw')).toBe('4455667788-13-1700000000');
  });

  it.each([
    {input: 'YQ==', expected: 'a', padding: 'two padding characters'},
    {input: 'YWI=', expected: 'ab', padding: 'one padding character'},
    {input: 'YWJj', expected: 'abc', padding: 'no padding needed'},
  ])('decodes standard base64 with $padding', ({input, expected}) => {
    expect(decodeBase64ToAscii(input)).toBe(expected);
  });

  it.each([
    {input: 'YQ', expected: 'a'},
    {input: 'YWI', expected: 'ab'},
  ])('decodes $input with its padding omitted', ({input, expected}) => {
    expect(decodeBase64ToAscii(input)).toBe(expected);
  });

  it('decodes the standard alphabet (+ and /)', () => {
    expect(decodeBase64ToAscii('Pj4+Pz8/')).toBe('>>>???');
  });

  it('decodes the URL-safe alphabet (- and _)', () => {
    expect(decodeBase64ToAscii('Pj4-Pz8_')).toBe('>>>???');
  });

  it.each([
    {input: 'IA==', expected: ' ', boundary: 'the first printable character (0x20)'},
    {input: 'fg==', expected: '~', boundary: 'the last printable character (0x7E)'},
  ])('decodes $boundary', ({input, expected}) => {
    expect(decodeBase64ToAscii(input)).toBe(expected);
  });

  it('returns an empty string for empty input', () => {
    expect(decodeBase64ToAscii('')).toBe('');
  });

  it.each([
    {input: 'not base64!', reason: 'it contains characters outside both alphabets'},
    {input: 'ab$d', reason: 'it contains a character outside both alphabets'},
    {input: 'Y===', reason: 'it has three padding characters'},
    {input: 'Y=Q=', reason: 'it has padding in the middle'},
    {input: '=YWJ', reason: 'it starts with padding'},
    {input: 'YWJjZ', reason: 'a single trailing character cannot encode a whole byte'},
    {input: '1', reason: 'a single character cannot encode a whole byte'},
    {input: 'YQ=', reason: 'it is padded to a length that is not a multiple of 4'},
    {input: 'YWI==', reason: 'it has more padding than its length needs'},
    {input: '==', reason: 'it is only padding'},
  ])('returns null for "$input" because $reason', ({input}) => {
    expect(decodeBase64ToAscii(input)).toBeNull();
  });

  it.each([
    {input: 'Hw==', reason: 'it decodes to 0x1F, just below the printable range'},
    {input: 'fw==', reason: 'it decodes to 0x7F, just above the printable range'},
    {input: '/wCA', reason: 'it decodes to the bytes 0xFF 0x00 0x80'},
    {input: '500001', reason: 'it is a synthetic demo-data ID that decodes to binary'},
    {input: 'rec1', reason: 'it is a synthetic test ID that decodes to binary'},
  ])('returns null for "$input" because $reason', ({input}) => {
    expect(decodeBase64ToAscii(input)).toBeNull();
  });

  it.each([
    {input: undefined, type: 'undefined'},
    {input: null, type: 'null'},
    {input: 123, type: 'a number'},
  ])('returns null for $type', ({input}) => {
    expect(decodeBase64ToAscii(input as unknown as string)).toBeNull();
  });
});
