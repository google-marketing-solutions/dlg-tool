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

import {vi} from 'vitest';
import type {MockAdsAccount} from '../mocks/google-apps-script.mock';

/**
 * Reusable test data builders so unit tests only specify fields relevant to the scenario under test.
 */
export function buildMockAccount(overrides: Partial<MockAdsAccount> = {}): MockAdsAccount {
  return {
    customerId: overrides.customerId ?? '123-456-7890',
    name: overrides.name ?? 'Acme Retail Search',
    currencyCode: overrides.currencyCode ?? 'USD',
    timeZone: overrides.timeZone ?? 'America/New_York',
  };
}

/**
 * Freezes the fake system clock at `dateStr` (`YYYY-MM-DD`) and `hour` (`0–23` UTC) for deterministic date tests.
 */
export function setTestDate(dateStr: string, hour = 10): void {
  vi.useFakeTimers();
  vi.setSystemTime(new Date(`${dateStr}T${String(hour).padStart(2, '0')}:00:00Z`));
}
