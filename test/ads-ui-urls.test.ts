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
import {
  buildCampaignUrl,
  buildRecommendationsDetailsUrl,
  parseRecommendationIdInfo,
  resolveAdsUiOcid,
} from '../src/appscript/processors/ads-ui-urls';

// Recommendation IDs end in base64(<ocid>-<recoTypeId>-<suffix>). Each test notes what its ID decodes to.

describe('parseRecommendationIdInfo', () => {
  it('decodes the ocid and recommendation type ID from the last resource-name segment', () => {
    // Decodes to 4455667788-13-1700000000
    const recommendationId = 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xMy0xNzAwMDAwMDAw';

    expect(parseRecommendationIdInfo(recommendationId)).toEqual({ocid: '4455667788', recoTypeId: '13'});
  });

  it('decodes a URL-safe segment', () => {
    // Decodes to 987654321-7-1700000000-ab?> (the trailing '-' is URL-safe for '+')
    const recommendationId = 'customers/1000000001/recommendations/OTg3NjU0MzIxLTctMTcwMDAwMDAwMC1hYj8-';

    expect(parseRecommendationIdInfo(recommendationId)).toEqual({ocid: '987654321', recoTypeId: '7'});
  });

  it('decodes a segment with its padding omitted', () => {
    // Decodes to 42-13-1 (padded form: NDItMTMtMQ==)
    const recommendationId = 'customers/1000000001/recommendations/NDItMTMtMQ';

    expect(parseRecommendationIdInfo(recommendationId)).toEqual({ocid: '42', recoTypeId: '13'});
  });

  it('decodes a bare encoded ID without the customers/<cid>/recommendations/ prefix', () => {
    // Decodes to 42-13-1
    expect(parseRecommendationIdInfo('NDItMTMtMQ==')).toEqual({ocid: '42', recoTypeId: '13'});
  });

  it('ignores whitespace around the encoded segment', () => {
    // Decodes to 42-13-1
    expect(parseRecommendationIdInfo('customers/1/recommendations/ NDItMTMtMQ== ')).toEqual({
      ocid: '42',
      recoTypeId: '13',
    });
  });

  it.each([
    {recommendationId: undefined, reason: 'it is undefined'},
    {recommendationId: '', reason: 'it is empty'},
    {recommendationId: 'N/A', reason: 'it is the N/A placeholder'},
    {recommendationId: 'customers/1/recommendations/', reason: 'the encoded segment is empty'},
    {recommendationId: '500001', reason: 'it is a synthetic demo-data ID'},
    {recommendationId: 'customers/100/recommendations/rec1', reason: 'it is a synthetic test ID'},
    {recommendationId: 'customers/1234567890/recommendations/99', reason: 'the segment is not base64 text'},
    {recommendationId: 'customers/1/recommendations/aGVsbG8td29ybGQ=', reason: 'it decodes to hello-world'},
    {recommendationId: 'customers/1/recommendations/NDQ1NTY2Nzc4OA==', reason: 'it decodes to an ocid only'},
    {
      recommendationId: 'customers/1/recommendations/NDQ1NTY2Nzc4OC1hYmMtMQ==',
      reason: 'it decodes to a non-numeric type ID (4455667788-abc-1)',
    },
    {recommendationId: 'customers/1/recommendations/MDAwMC0xMy0x', reason: 'it decodes to an all-zero ocid'},
  ])('returns null when $reason', ({recommendationId}) => {
    expect(parseRecommendationIdInfo(recommendationId)).toBeNull();
  });
});

describe('resolveAdsUiOcid', () => {
  it('returns the ocid decoded from the recommendation ID', () => {
    // Decodes to 4455667788-13-1700000000
    const recommendationId = 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xMy0xNzAwMDAwMDAw';

    expect(resolveAdsUiOcid(recommendationId, '100-000-0001')).toBe('4455667788');
  });

  it('falls back to the normalized CID when the recommendation ID cannot be decoded', () => {
    expect(resolveAdsUiOcid('customers/100/recommendations/rec1', '200-000-0002')).toBe('2000000002');
  });
});

describe('buildRecommendationsDetailsUrl', () => {
  it('links to the recommendation type with the decoded ocid and opp parameters', () => {
    // Decodes to 4455667788-13-1700000000
    const recommendationId = 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xMy0xNzAwMDAwMDAw';

    expect(buildRecommendationsDetailsUrl(recommendationId, '100-000-0001')).toBe(
      'https://ads.google.com/aw/recommendations?ocid=4455667788&opp=13'
    );
  });

  it("falls back to the account's Recommendations page when the ID cannot be decoded", () => {
    expect(buildRecommendationsDetailsUrl('500001', '101-202-3030')).toBe(
      'https://ads.google.com/aw/recommendations?ocid=1012023030'
    );
  });

  it('pads a short numeric CID from Sheets to 10 digits in the fallback link', () => {
    expect(buildRecommendationsDetailsUrl('500001', 12345678)).toBe(
      'https://ads.google.com/aw/recommendations?ocid=0012345678'
    );
  });
});

describe('buildCampaignUrl', () => {
  it('links to the campaign overview with the decoded ocid', () => {
    // Decodes to 4455667788-13-1700000000
    const recommendationId = 'customers/1000000001/recommendations/NDQ1NTY2Nzc4OC0xMy0xNzAwMDAwMDAw';

    expect(buildCampaignUrl('11', recommendationId, '100-000-0001')).toBe(
      'https://ads.google.com/aw/overview?campaignId=11&ocid=4455667788'
    );
  });

  it('falls back to the CID when the recommendation ID cannot be decoded', () => {
    expect(buildCampaignUrl('11', '500001', '100-000-0001')).toBe(
      'https://ads.google.com/aw/overview?campaignId=11&ocid=1000000001'
    );
  });
});
