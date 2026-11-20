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
 * @fileoverview Google Ads UI deep-link builders (`recommendationsDetailsUrl`, `campaignUrl`).
 * The Google Ads UI addresses accounts by `ocid`, a customer ID that differs from the 10-digit
 * Customer ID (CID). The trailing segment of a recommendation resource name
 * (`customers/<cid>/recommendations/<base64>`) decodes to `<ocid>-<recoTypeId>-...`, so links use the decoded
 * `ocid` (and `opp=<recoTypeId>` for the recommendation-type page) and fall back to the CID when the ID cannot
 * be decoded (e.g. synthetic demo IDs). Bundled only into `dist/sheet/Code.gs`.
 */

import {decodeBase64ToAscii} from '../../shared/utils/base64';
import {normalizeCid} from '../../shared/utils/cid';

const ADS_UI_BASE_URL = 'https://ads.google.com/aw';

/** Google Ads UI identifiers decoded from a recommendation resource name. */
export interface RecommendationIdInfo {
  ocid: string;
  recoTypeId: string;
}

/**
 * Decodes `{ocid, recoTypeId}` from the trailing segment of a recommendation resource name
 * (`customers/<cid>/recommendations/<base64(<ocid>-<recoTypeId>-...)>`), or returns `null` when the segment
 * is not valid base64 or does not start with two numeric `-`-separated fields.
 */
export function parseRecommendationIdInfo(recommendationId: string | undefined): RecommendationIdInfo | null {
  if (!recommendationId) return null;
  const encoded = String(recommendationId).split('/').pop()?.trim() ?? '';
  if (!encoded) return null;
  const decoded = decodeBase64ToAscii(encoded);
  if (!decoded) return null;
  const [ocid, recoTypeId] = decoded.split('-');
  if (!/^\d+$/.test(ocid ?? '') || /^0+$/.test(ocid) || !/^\d+$/.test(recoTypeId ?? '')) return null;
  return {ocid, recoTypeId};
}

/** Google Ads UI `ocid` for a row: decoded from `recommendationId`, falling back to the normalized CID. */
export function resolveAdsUiOcid(recommendationId: string | undefined, fallbackCid: unknown): string {
  return parseRecommendationIdInfo(recommendationId)?.ocid ?? normalizeCid(fallbackCid);
}

/**
 * Deep link to the recommendation type in the Google Ads UI (`recommendations?ocid=<ocid>&opp=<recoTypeId>`),
 * falling back to the account's Recommendations page (`ocid=<CID>`) when the ID cannot be decoded.
 */
export function buildRecommendationsDetailsUrl(recommendationId: string | undefined, accountCid: unknown): string {
  const info = parseRecommendationIdInfo(recommendationId);
  if (!info) return `${ADS_UI_BASE_URL}/recommendations?ocid=${normalizeCid(accountCid)}`;
  return `${ADS_UI_BASE_URL}/recommendations?ocid=${info.ocid}&opp=${info.recoTypeId}`;
}

/** Deep link to a campaign's Overview page, using the decoded `ocid` (falling back to the CID). */
export function buildCampaignUrl(
  campaignId: string,
  recommendationId: string | undefined,
  accountCid: unknown
): string {
  return `${ADS_UI_BASE_URL}/overview?campaignId=${campaignId}&ocid=${resolveAdsUiOcid(recommendationId, accountCid)}`;
}
