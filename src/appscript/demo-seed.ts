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
 * @fileoverview Synthetic Demo MCC & Recommendation seeder (`seedSyntheticDemoData`) for testing
 * the Google Sheet `Config` -> `_Stage_*` -> `Data` pipeline and Data Studio Dashboard end-to-end
 * when live test MCCs have 0 active recommendations.
 */

import {writeMccSheetState} from '../shared/sheets/mcc-state';
import {
  type SpreadsheetLike,
  getOrCreateSheet,
} from '../shared/sheets/sheet-utils';
import {
  appendStagingBatch,
  clearMccFromStaging,
} from '../shared/sheets/staging-writer';
import type {StagingBatch} from '../shared/types';
import {normalizeCid} from '../shared/utils/cid';
import {resolveSpreadsheetDate} from '../shared/utils/date';
import {
  CID_COL,
  CID_TABLE_START_ROW,
  CONFIG_SHEET_NAME,
  MCC_COL,
  MCC_TABLE_END_ROW,
  MCC_TABLE_START_ROW,
  assertConfigLayout,
  sheetCol,
  syncDiscoveredAccountsToConfig,
} from './config-sheet';
import {publishFromStagingToData} from './publisher';
import {resetTodayPublisherState, withSheetLock} from './run-state';

/** Hyphenated Customer ID for the primary synthetic Demo MCC. */
export const SYNTHETIC_MCC_ID = '999-888-7777';
/** Normalized 10-digit Customer ID for the primary synthetic Demo MCC. */
export const SYNTHETIC_NORM_MCC_ID = '9998887777';
/** Display name for the primary synthetic Demo MCC. */
export const SYNTHETIC_MCC_NAME = '[Demo] Global Retail MCC';

/** Hyphenated Customer ID for the secondary synthetic Demo MCC. */
export const SYNTHETIC_SECOND_MCC_ID = '888-777-6666';
/** Normalized 10-digit Customer ID for the secondary synthetic Demo MCC. */
export const SYNTHETIC_SECOND_NORM_MCC_ID = '8887776666';
/** Display name for the secondary synthetic Demo MCC. */
export const SYNTHETIC_SECOND_MCC_NAME = '[Demo] DACH & Nordics Aviation MCC';

/** Synthetic child accounts registered under `SYNTHETIC_MCC_ID` (`USD`, `EUR`, `GBP`). */
export const SYNTHETIC_CHILD_ACCOUNTS = [
  {
    customerId: '101-202-3030',
    name: '[Demo] US Search & AI Max',
    currencyCode: 'USD',
  },
  {
    customerId: '202-303-4040',
    name: '[Demo] EMEA Retail & Shared Budgets',
    currencyCode: 'EUR',
  },
  {
    customerId: '303-404-5050',
    name: '[Demo] UK Performance & Portfolios',
    currencyCode: 'GBP',
  },
];

/** Synthetic child accounts registered under `SYNTHETIC_SECOND_MCC_ID` (`EUR`, `CHF`). */
export const SYNTHETIC_SECOND_MCC_CHILD_ACCOUNTS = [
  {
    customerId: '404-505-6060',
    name: '[Demo] DE Airlines Search & PMax',
    currencyCode: 'EUR',
  },
  {
    customerId: '505-606-7070',
    name: '[Demo] CH Alpine Routes & Portfolios',
    currencyCode: 'CHF',
  },
];

/**
 * Builds a realistic synthetic `StagingBatch` covering all 6 recommendation types,
 * all 3 scopes (`CAMPAIGN`, `SHARED_BUDGET`, `PORTFOLIO`), multi-currency (`USD`, `EUR`, `GBP`),
 * AI Max flags, Move Unused Budget source fields, Locally-Owned vs Manager-Owned portfolios,
 * a `SHOPPING` campaign (to test the EC Remedy toggle), budget recs on the campaigns of a portfolio that
 * also has a bid rec, and a standalone campaign with both a budget rec and a bid rec.
 */
export function buildSyntheticDemoBatch(dateStr: string): StagingBatch {
  const timestamp = `${dateStr} 00:00:00`;

  const accounts = SYNTHETIC_CHILD_ACCOUNTS.map((acc) => ({
    mccId: SYNTHETIC_MCC_ID,
    mccName: SYNTHETIC_MCC_NAME,
    accountId: acc.customerId,
    accountName: acc.name,
    currencyCode: acc.currencyCode,
    timestamp,
  }));

  const budgets = [
    // Account 1 (USD): 101-202-3030
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '101-202-3030',
      budgetResourceName: 'customers/1012023030/campaignBudgets/1001',
      campaignBudgetName: 'US - Brand Search Daily Budget',
      campaignBudgetAmount: 500,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: false,
    },
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '101-202-3030',
      budgetResourceName: 'customers/1012023030/campaignBudgets/1002',
      campaignBudgetName: 'US - Q4 Peak Demand PMax Budget',
      campaignBudgetAmount: 1200,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: false,
    },
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '101-202-3030',
      budgetResourceName: 'customers/1012023030/campaignBudgets/1003',
      campaignBudgetName: 'US - Non-Brand High Intent Budget',
      campaignBudgetAmount: 400,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: false,
    },
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '101-202-3030',
      budgetResourceName: 'customers/1012023030/campaignBudgets/1099',
      campaignBudgetName: 'US - Legacy Display Shared Budget (Donor)',
      campaignBudgetAmount: 650,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: true,
    },
    // Account 2 (EUR): 202-303-4040
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '202-303-4040',
      budgetResourceName: 'customers/2023034040/campaignBudgets/2001',
      campaignBudgetName: 'EMEA - Core Generic Shared Budget',
      campaignBudgetAmount: 900,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: true,
    },
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '202-303-4040',
      budgetResourceName: 'customers/2023034040/campaignBudgets/2002',
      campaignBudgetName: 'EMEA - Standard Shopping Budget',
      campaignBudgetAmount: 750,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: false,
    },
    // Account 3 (GBP): 303-404-5050
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '303-404-5050',
      budgetResourceName: 'customers/3034045050/campaignBudgets/3001',
      campaignBudgetName: 'UK - London Search Budget',
      campaignBudgetAmount: 350,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: false,
    },
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '303-404-5050',
      budgetResourceName: 'customers/3034045050/campaignBudgets/3002',
      campaignBudgetName: 'UK - Midlands Search Budget',
      campaignBudgetAmount: 280,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: false,
    },
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '303-404-5050',
      budgetResourceName: 'customers/3034045050/campaignBudgets/3003',
      campaignBudgetName: 'UK - Demand Gen Video Budget',
      campaignBudgetAmount: 420,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: false,
    },
  ];

  const campaigns = [
    // 1. US Brand Search with AI Max enabled (CAMPAIGN_BUDGET)
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '101-202-3030',
      campaignResourceName: 'customers/1012023030/campaigns/11001',
      campaignId: '11001',
      campaignName: 'US - Brand Search - AI Max',
      campaignType: 'SEARCH',
      campaignSubType: 'UNSPECIFIED',
      campaignIsAiMax: true,
      campaignAiMaxTextCustomizationEnabled: true,
      campaignAiMaxFinalUrlExpansionEnabled: true,
      budgetResourceName: 'customers/1012023030/campaignBudgets/1001',
      campaignIsPortfolioBiddingStrategy: false,
      campaignPortfolioBiddingStrategyName: 'N/A',
      portfolioStrategyResourceName: 'N/A',
      portfolioOwnerId: 'N/A',
      isLocallyOwnedPortfolio: false,
      campaignPortfolioIsManagerOwned: false,
      campaignBiddingStrategyType: 'MAXIMIZE_CONVERSIONS',
      campaignBiddingTargetRoas: null,
      campaignBiddingTargetCpa: 18.5,
      campaignStatsYesterdayCost: 495,
      campaignStats7DaysCost: 3420,
      campaign30DaysCost: 14600,
      campaign30DaysConversionsValue: 58400,
      campaign30DaysConversions: 810,
      campaign30DaysSearchRankLostImpressionShare: 0.08,
      campaign30DaysSearchBudgetLostImpressionShare: 0.34,
    },
    // 2. US Non-Feed PMax (FORECASTING_CAMPAIGN_BUDGET)
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '101-202-3030',
      campaignResourceName: 'customers/1012023030/campaigns/11002',
      campaignId: '11002',
      campaignName: 'US - Q4 Peak Demand PMax (Lead Gen)',
      campaignType: 'PERFORMANCE_MAX',
      campaignSubType: 'UNSPECIFIED',
      campaignIsAiMax: false,
      campaignAiMaxTextCustomizationEnabled: false,
      campaignAiMaxFinalUrlExpansionEnabled: false,
      budgetResourceName: 'customers/1012023030/campaignBudgets/1002',
      campaignIsPortfolioBiddingStrategy: false,
      campaignPortfolioBiddingStrategyName: 'N/A',
      portfolioStrategyResourceName: 'N/A',
      portfolioOwnerId: 'N/A',
      isLocallyOwnedPortfolio: false,
      campaignPortfolioIsManagerOwned: false,
      campaignBiddingStrategyType: 'MAXIMIZE_CONVERSION_VALUE',
      campaignBiddingTargetRoas: 3.6,
      campaignBiddingTargetCpa: null,
      campaignStatsYesterdayCost: 1190,
      campaignStats7DaysCost: 8250,
      campaign30DaysCost: 35100,
      campaign30DaysConversionsValue: 133380,
      campaign30DaysConversions: 1420,
      campaign30DaysSearchRankLostImpressionShare: null,
      campaign30DaysSearchBudgetLostImpressionShare: null,
    },
    // 3. US Non-Brand Search (MOVE_UNUSED_BUDGET)
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '101-202-3030',
      campaignResourceName: 'customers/1012023030/campaigns/11003',
      campaignId: '11003',
      campaignName: 'US - Non-Brand High Intent Search',
      campaignType: 'SEARCH',
      campaignSubType: 'UNSPECIFIED',
      campaignIsAiMax: true,
      campaignAiMaxTextCustomizationEnabled: true,
      campaignAiMaxFinalUrlExpansionEnabled: false,
      budgetResourceName: 'customers/1012023030/campaignBudgets/1003',
      campaignIsPortfolioBiddingStrategy: false,
      campaignPortfolioBiddingStrategyName: 'N/A',
      portfolioStrategyResourceName: 'N/A',
      portfolioOwnerId: 'N/A',
      isLocallyOwnedPortfolio: false,
      campaignPortfolioIsManagerOwned: false,
      campaignBiddingStrategyType: 'TARGET_CPA',
      campaignBiddingTargetRoas: null,
      campaignBiddingTargetCpa: 26,
      campaignStatsYesterdayCost: 398,
      campaignStats7DaysCost: 2760,
      campaign30DaysCost: 11850,
      campaign30DaysConversionsValue: 39105,
      campaign30DaysConversions: 465,
      campaign30DaysSearchRankLostImpressionShare: 0.14,
      campaign30DaysSearchBudgetLostImpressionShare: 0.41,
    },
    // 4a & 4b. EMEA Shared Budget across 2 campaigns (MARGINAL_ROI_CAMPAIGN_BUDGET -> SHARED_BUDGET scope)
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '202-303-4040',
      campaignResourceName: 'customers/2023034040/campaigns/22001',
      campaignId: '22001',
      campaignName: 'DE - Generic Search - Core',
      campaignType: 'SEARCH',
      campaignSubType: 'UNSPECIFIED',
      campaignIsAiMax: false,
      campaignAiMaxTextCustomizationEnabled: false,
      campaignAiMaxFinalUrlExpansionEnabled: false,
      budgetResourceName: 'customers/2023034040/campaignBudgets/2001',
      campaignIsPortfolioBiddingStrategy: false,
      campaignPortfolioBiddingStrategyName: 'N/A',
      portfolioStrategyResourceName: 'N/A',
      portfolioOwnerId: 'N/A',
      isLocallyOwnedPortfolio: false,
      campaignPortfolioIsManagerOwned: false,
      campaignBiddingStrategyType: 'MAXIMIZE_CONVERSIONS',
      campaignBiddingTargetRoas: null,
      campaignBiddingTargetCpa: 22,
      campaignStatsYesterdayCost: 540,
      campaignStats7DaysCost: 3750,
      campaign30DaysCost: 16200,
      campaign30DaysConversionsValue: 64800,
      campaign30DaysConversions: 750,
      campaign30DaysSearchRankLostImpressionShare: 0.11,
      campaign30DaysSearchBudgetLostImpressionShare: 0.29,
    },
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '202-303-4040',
      campaignResourceName: 'customers/2023034040/campaigns/22002',
      campaignId: '22002',
      campaignName: 'FR - Generic Search - Core',
      campaignType: 'SEARCH',
      campaignSubType: 'UNSPECIFIED',
      campaignIsAiMax: false,
      campaignAiMaxTextCustomizationEnabled: false,
      campaignAiMaxFinalUrlExpansionEnabled: false,
      budgetResourceName: 'customers/2023034040/campaignBudgets/2001',
      campaignIsPortfolioBiddingStrategy: false,
      campaignPortfolioBiddingStrategyName: 'N/A',
      portfolioStrategyResourceName: 'N/A',
      portfolioOwnerId: 'N/A',
      isLocallyOwnedPortfolio: false,
      campaignPortfolioIsManagerOwned: false,
      campaignBiddingStrategyType: 'MAXIMIZE_CONVERSIONS',
      campaignBiddingTargetRoas: null,
      campaignBiddingTargetCpa: 24,
      campaignStatsYesterdayCost: 355,
      campaignStats7DaysCost: 2480,
      campaign30DaysCost: 10600,
      campaign30DaysConversionsValue: 41340,
      campaign30DaysConversions: 445,
      campaign30DaysSearchRankLostImpressionShare: 0.13,
      campaign30DaysSearchBudgetLostImpressionShare: 0.27,
    },
    // 5. EMEA Standard Shopping (LOWER_TARGET_ROAS -> tests Exclude Shopping toggle)
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '202-303-4040',
      campaignResourceName: 'customers/2023034040/campaigns/22003',
      campaignId: '22003',
      campaignName: 'EMEA - Merchant Center Standard Shopping',
      campaignType: 'SHOPPING',
      campaignSubType: 'SHOPPING_SMART_ADS',
      campaignIsAiMax: false,
      campaignAiMaxTextCustomizationEnabled: false,
      campaignAiMaxFinalUrlExpansionEnabled: false,
      budgetResourceName: 'customers/2023034040/campaignBudgets/2002',
      campaignIsPortfolioBiddingStrategy: false,
      campaignPortfolioBiddingStrategyName: 'N/A',
      portfolioStrategyResourceName: 'N/A',
      portfolioOwnerId: 'N/A',
      isLocallyOwnedPortfolio: false,
      campaignPortfolioIsManagerOwned: false,
      campaignBiddingStrategyType: 'TARGET_ROAS',
      campaignBiddingTargetRoas: 4.5,
      campaignBiddingTargetCpa: null,
      campaignStatsYesterdayCost: 690,
      campaignStats7DaysCost: 4820,
      campaign30DaysCost: 20500,
      campaign30DaysConversionsValue: 96350,
      campaign30DaysConversions: 920,
      campaign30DaysSearchRankLostImpressionShare: 0.22,
      campaign30DaysSearchBudgetLostImpressionShare: 0.09,
    },
    // 6a & 6b. UK Locally-Owned Portfolio across 2 campaigns (RAISE_TARGET_CPA -> PORTFOLIO scope)
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '303-404-5050',
      campaignResourceName: 'customers/3034045050/campaigns/33001',
      campaignId: '33001',
      campaignName: 'UK - Search - London & South East',
      campaignType: 'SEARCH',
      campaignSubType: 'UNSPECIFIED',
      campaignIsAiMax: true,
      campaignAiMaxTextCustomizationEnabled: true,
      campaignAiMaxFinalUrlExpansionEnabled: true,
      budgetResourceName: 'customers/3034045050/campaignBudgets/3001',
      campaignIsPortfolioBiddingStrategy: true,
      campaignPortfolioBiddingStrategyName: 'UK - Target CPA Growth Portfolio',
      portfolioStrategyResourceName: 'customers/3034045050/biddingStrategies/9001',
      portfolioOwnerId: '3034045050',
      isLocallyOwnedPortfolio: true,
      campaignPortfolioIsManagerOwned: false,
      campaignBiddingStrategyType: 'TARGET_CPA',
      campaignBiddingTargetRoas: null,
      campaignBiddingTargetCpa: 24,
      campaignStatsYesterdayCost: 320,
      campaignStats7DaysCost: 2210,
      campaign30DaysCost: 9450,
      campaign30DaysConversionsValue: 37800,
      campaign30DaysConversions: 410,
      campaign30DaysSearchRankLostImpressionShare: 0.19,
      campaign30DaysSearchBudgetLostImpressionShare: 0.05,
    },
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '303-404-5050',
      campaignResourceName: 'customers/3034045050/campaigns/33002',
      campaignId: '33002',
      campaignName: 'UK - Search - Midlands & North',
      campaignType: 'SEARCH',
      campaignSubType: 'UNSPECIFIED',
      campaignIsAiMax: false,
      campaignAiMaxTextCustomizationEnabled: false,
      campaignAiMaxFinalUrlExpansionEnabled: false,
      budgetResourceName: 'customers/3034045050/campaignBudgets/3002',
      campaignIsPortfolioBiddingStrategy: true,
      campaignPortfolioBiddingStrategyName: 'UK - Target CPA Growth Portfolio',
      portfolioStrategyResourceName: 'customers/3034045050/biddingStrategies/9001',
      portfolioOwnerId: '3034045050',
      isLocallyOwnedPortfolio: true,
      campaignPortfolioIsManagerOwned: false,
      campaignBiddingStrategyType: 'TARGET_CPA',
      campaignBiddingTargetRoas: null,
      campaignBiddingTargetCpa: 24,
      campaignStatsYesterdayCost: 250,
      campaignStats7DaysCost: 1740,
      campaign30DaysCost: 7400,
      campaign30DaysConversionsValue: 29600,
      campaign30DaysConversions: 320,
      campaign30DaysSearchRankLostImpressionShare: 0.21,
      campaign30DaysSearchBudgetLostImpressionShare: 0.04,
    },
    // 7. UK Manager-Owned MCC Portfolio (LOWER_TARGET_ROAS -> PORTFOLIO scope)
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '303-404-5050',
      campaignResourceName: 'customers/3034045050/campaigns/33003',
      campaignId: '33003',
      campaignName: 'UK - Demand Gen Video & Discovery',
      campaignType: 'DEMAND_GEN',
      campaignSubType: 'UNSPECIFIED',
      campaignIsAiMax: false,
      campaignAiMaxTextCustomizationEnabled: false,
      campaignAiMaxFinalUrlExpansionEnabled: false,
      budgetResourceName: 'customers/3034045050/campaignBudgets/3003',
      campaignIsPortfolioBiddingStrategy: true,
      campaignPortfolioBiddingStrategyName: 'MCC Global ROAS Portfolio',
      portfolioStrategyResourceName: 'customers/9998887777/biddingStrategies/9900',
      portfolioOwnerId: '9998887777',
      isLocallyOwnedPortfolio: false,
      campaignPortfolioIsManagerOwned: true,
      campaignBiddingStrategyType: 'TARGET_ROAS',
      campaignBiddingTargetRoas: 5.2,
      campaignBiddingTargetCpa: null,
      campaignStatsYesterdayCost: 390,
      campaignStats7DaysCost: 2710,
      campaign30DaysCost: 11600,
      campaign30DaysConversionsValue: 61480,
      campaign30DaysConversions: 580,
      campaign30DaysSearchRankLostImpressionShare: null,
      campaign30DaysSearchBudgetLostImpressionShare: null,
    },
  ];

  const recommendations = [
    // Rec 1: CAMPAIGN_BUDGET (US Brand Search)
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '101-202-3030',
      recommendationId: '500001',
      recommendationType: 'CAMPAIGN_BUDGET',
      campaignResourceName: 'customers/1012023030/campaigns/11001',
      recommendationCurrentBudgetAmount: 500,
      recommendationNewBudgetAmount: 850,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 3500,
      recommendationPotentialCost: 5950,
      recommendationBaseClicks: 2800,
      recommendationPotentialClicks: 4650,
      recommendationBaseConversions: 195,
      recommendationPotentialConversions: 315,
      recommendationBaseConversionsValue: 14040,
      recommendationPotentialConversionsValue: 22680,
      recommendationBaseImpressions: 42000,
      recommendationPotentialImpressions: 68500,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    // Rec 2: FORECASTING_CAMPAIGN_BUDGET (US Q4 Peak Demand PMax)
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '101-202-3030',
      recommendationId: '500002',
      recommendationType: 'FORECASTING_CAMPAIGN_BUDGET',
      campaignResourceName: 'customers/1012023030/campaigns/11002',
      recommendationCurrentBudgetAmount: 1200,
      recommendationNewBudgetAmount: 1800,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 8400,
      recommendationPotentialCost: 12600,
      recommendationBaseClicks: 6100,
      recommendationPotentialClicks: 8950,
      recommendationBaseConversions: 340,
      recommendationPotentialConversions: 495,
      recommendationBaseConversionsValue: 31960,
      recommendationPotentialConversionsValue: 45540,
      recommendationBaseImpressions: 185000,
      recommendationPotentialImpressions: 270000,
      recommendationBaseVideoViews: 12400,
      recommendationPotentialVideoViews: 18900,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    // Rec 3: MOVE_UNUSED_BUDGET (US Non-Brand Search)
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '101-202-3030',
      recommendationId: '500003',
      recommendationType: 'MOVE_UNUSED_BUDGET',
      campaignResourceName: 'customers/1012023030/campaigns/11003',
      recommendationCurrentBudgetAmount: 400,
      recommendationNewBudgetAmount: 650,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 2800,
      recommendationPotentialCost: 4550,
      recommendationBaseClicks: 1900,
      recommendationPotentialClicks: 3050,
      recommendationBaseConversions: 108,
      recommendationPotentialConversions: 172,
      recommendationBaseConversionsValue: 9180,
      recommendationPotentialConversionsValue: 14620,
      recommendationBaseImpressions: 31000,
      recommendationPotentialImpressions: 49000,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: 'customers/1012023030/campaignBudgets/1099',
      moveBudgetSourceCampaigns: 'US - Upper Funnel Display | US - Remarketing Display',
      moveBudgetSourceTargetCpa: 32,
      moveBudgetSourceTargetRoas: 2.4,
    },
    // Rec 4a & 4b: MARGINAL_ROI_CAMPAIGN_BUDGET across 2 campaigns sharing EMEA Core Generic Shared Budget
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '202-303-4040',
      recommendationId: '500004',
      recommendationType: 'MARGINAL_ROI_CAMPAIGN_BUDGET',
      campaignResourceName: 'customers/2023034040/campaigns/22001',
      recommendationCurrentBudgetAmount: 900,
      recommendationNewBudgetAmount: 1350,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 6300,
      recommendationPotentialCost: 9450,
      recommendationBaseClicks: 4900,
      recommendationPotentialClicks: 7200,
      recommendationBaseConversions: 285,
      recommendationPotentialConversions: 415,
      recommendationBaseConversionsValue: 25650,
      recommendationPotentialConversionsValue: 37350,
      recommendationBaseImpressions: 78000,
      recommendationPotentialImpressions: 114000,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '202-303-4040',
      recommendationId: '500004',
      recommendationType: 'MARGINAL_ROI_CAMPAIGN_BUDGET',
      campaignResourceName: 'customers/2023034040/campaigns/22002',
      recommendationCurrentBudgetAmount: 900,
      recommendationNewBudgetAmount: 1350,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 6300,
      recommendationPotentialCost: 9450,
      recommendationBaseClicks: 4900,
      recommendationPotentialClicks: 7200,
      recommendationBaseConversions: 285,
      recommendationPotentialConversions: 415,
      recommendationBaseConversionsValue: 25650,
      recommendationPotentialConversionsValue: 37350,
      recommendationBaseImpressions: 78000,
      recommendationPotentialImpressions: 114000,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    // Rec 5: LOWER_TARGET_ROAS on EMEA Standard Shopping
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '202-303-4040',
      recommendationId: '500005',
      recommendationType: 'LOWER_TARGET_ROAS',
      campaignResourceName: 'customers/2023034040/campaigns/22003',
      recommendationCurrentBudgetAmount: 750,
      recommendationNewBudgetAmount: 750,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: 3.8,
      recommendationBaseCost: 4800,
      recommendationPotentialCost: 6650,
      recommendationBaseClicks: 5400,
      recommendationPotentialClicks: 7600,
      recommendationBaseConversions: 215,
      recommendationPotentialConversions: 295,
      recommendationBaseConversionsValue: 21600,
      recommendationPotentialConversionsValue: 28800,
      recommendationBaseImpressions: 140000,
      recommendationPotentialImpressions: 195000,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    // Rec 6a & 6b: RAISE_TARGET_CPA on UK Locally-Owned Portfolio across 2 campaigns
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '303-404-5050',
      recommendationId: '500006',
      recommendationType: 'RAISE_TARGET_CPA',
      campaignResourceName: 'customers/3034045050/campaigns/33001',
      recommendationCurrentBudgetAmount: 350,
      recommendationNewBudgetAmount: 350,
      recommendationNewTargetCpa: 29.5,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 3950,
      recommendationPotentialCost: 5400,
      recommendationBaseClicks: 3100,
      recommendationPotentialClicks: 4300,
      recommendationBaseConversions: 170,
      recommendationPotentialConversions: 225,
      recommendationBaseConversionsValue: 15640,
      recommendationPotentialConversionsValue: 20700,
      recommendationBaseImpressions: 52000,
      recommendationPotentialImpressions: 71000,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '303-404-5050',
      recommendationId: '500006',
      recommendationType: 'RAISE_TARGET_CPA',
      campaignResourceName: 'customers/3034045050/campaigns/33002',
      recommendationCurrentBudgetAmount: 280,
      recommendationNewBudgetAmount: 280,
      recommendationNewTargetCpa: 29.5,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 3950,
      recommendationPotentialCost: 5400,
      recommendationBaseClicks: 3100,
      recommendationPotentialClicks: 4300,
      recommendationBaseConversions: 170,
      recommendationPotentialConversions: 225,
      recommendationBaseConversionsValue: 15640,
      recommendationPotentialConversionsValue: 20700,
      recommendationBaseImpressions: 52000,
      recommendationPotentialImpressions: 71000,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    // Rec 7: LOWER_TARGET_ROAS on UK Manager-Owned MCC Portfolio (Demand Gen Video)
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '303-404-5050',
      recommendationId: '500007',
      recommendationType: 'LOWER_TARGET_ROAS',
      campaignResourceName: 'customers/3034045050/campaigns/33003',
      recommendationCurrentBudgetAmount: 420,
      recommendationNewBudgetAmount: 420,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: 4.4,
      recommendationBaseCost: 2700,
      recommendationPotentialCost: 3850,
      recommendationBaseClicks: 4200,
      recommendationPotentialClicks: 5900,
      recommendationBaseConversions: 135,
      recommendationPotentialConversions: 188,
      recommendationBaseConversionsValue: 14310,
      recommendationPotentialConversionsValue: 18920,
      recommendationBaseImpressions: 210000,
      recommendationPotentialImpressions: 295000,
      recommendationBaseVideoViews: 45000,
      recommendationPotentialVideoViews: 62000,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    // Rec 8 & 9: CAMPAIGN_BUDGET on each campaign of the UK Locally-Owned Portfolio (Rec 6). A portfolio
    // owns bidding, not budgets, so each budget rec stays a CAMPAIGN-scope row next to the PORTFOLIO row.
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '303-404-5050',
      recommendationId: '500008',
      recommendationType: 'CAMPAIGN_BUDGET',
      campaignResourceName: 'customers/3034045050/campaigns/33001',
      recommendationCurrentBudgetAmount: 350,
      recommendationNewBudgetAmount: 500,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 2200,
      recommendationPotentialCost: 3150,
      recommendationBaseClicks: 1750,
      recommendationPotentialClicks: 2450,
      recommendationBaseConversions: 96,
      recommendationPotentialConversions: 130,
      recommendationBaseConversionsValue: 8960,
      recommendationPotentialConversionsValue: 12100,
      recommendationBaseImpressions: 29000,
      recommendationPotentialImpressions: 40000,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '303-404-5050',
      recommendationId: '500009',
      recommendationType: 'CAMPAIGN_BUDGET',
      campaignResourceName: 'customers/3034045050/campaigns/33002',
      recommendationCurrentBudgetAmount: 280,
      recommendationNewBudgetAmount: 400,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 1730,
      recommendationPotentialCost: 2470,
      recommendationBaseClicks: 1350,
      recommendationPotentialClicks: 1900,
      recommendationBaseConversions: 75,
      recommendationPotentialConversions: 101,
      recommendationBaseConversionsValue: 6900,
      recommendationPotentialConversionsValue: 9400,
      recommendationBaseImpressions: 23000,
      recommendationPotentialImpressions: 32000,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    // Rec 10: RAISE_TARGET_CPA on US Brand Search, which also has a CAMPAIGN_BUDGET rec (Rec 1), so one
    // standalone campaign shows up on both the Budgets and the Bids dashboard pages.
    {
      mccId: SYNTHETIC_MCC_ID,
      accountId: '101-202-3030',
      recommendationId: '500010',
      recommendationType: 'RAISE_TARGET_CPA',
      campaignResourceName: 'customers/1012023030/campaigns/11001',
      recommendationCurrentBudgetAmount: 500,
      recommendationNewBudgetAmount: 500,
      recommendationNewTargetCpa: 21.5,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 3420,
      recommendationPotentialCost: 4100,
      recommendationBaseClicks: 2700,
      recommendationPotentialClicks: 3150,
      recommendationBaseConversions: 189,
      recommendationPotentialConversions: 215,
      recommendationBaseConversionsValue: 13600,
      recommendationPotentialConversionsValue: 15480,
      recommendationBaseImpressions: 41000,
      recommendationPotentialImpressions: 47500,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
  ];

  return {
    accounts,
    budgets,
    campaigns,
    recommendations,
  };
}

/**
 * Builds a second synthetic `StagingBatch` for `888-777-6666` (`[Demo] DACH & Nordics Aviation MCC`)
 * so Multi-MCC dropdown filtering and synchronization can be tested end-to-end in Data Studio.
 */
export function buildSyntheticSecondMccDemoBatch(dateStr: string): StagingBatch {
  const timestamp = `${dateStr} 00:00:00`;

  const accounts = SYNTHETIC_SECOND_MCC_CHILD_ACCOUNTS.map((acc) => ({
    mccId: SYNTHETIC_SECOND_MCC_ID,
    mccName: SYNTHETIC_SECOND_MCC_NAME,
    accountId: acc.customerId,
    accountName: acc.name,
    currencyCode: acc.currencyCode,
    timestamp,
  }));

  const budgets = [
    {
      mccId: SYNTHETIC_SECOND_MCC_ID,
      accountId: '404-505-6060',
      budgetResourceName: 'customers/4045056060/campaignBudgets/4001',
      campaignBudgetName: 'DE - Direct Flights Brand Budget',
      campaignBudgetAmount: 1100,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: false,
    },
    {
      mccId: SYNTHETIC_SECOND_MCC_ID,
      accountId: '404-505-6060',
      budgetResourceName: 'customers/4045056060/campaignBudgets/4002',
      campaignBudgetName: 'DE - Summer Holiday Routes PMax Budget',
      campaignBudgetAmount: 1600,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: false,
    },
    {
      mccId: SYNTHETIC_SECOND_MCC_ID,
      accountId: '505-606-7070',
      budgetResourceName: 'customers/5056067070/campaignBudgets/5001',
      campaignBudgetName: 'CH - Zurich & Geneva Hub Search Budget',
      campaignBudgetAmount: 820,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'STANDARD',
      campaignBudgetType: 'DAILY',
      campaignBudgetIsShared: false,
    },
  ];

  const campaigns = [
    {
      mccId: SYNTHETIC_SECOND_MCC_ID,
      accountId: '404-505-6060',
      campaignResourceName: 'customers/4045056060/campaigns/44001',
      campaignId: '44001',
      campaignName: 'DE - Direct Flights Brand Search - AI Max',
      campaignType: 'SEARCH',
      campaignSubType: 'UNSPECIFIED',
      campaignIsAiMax: true,
      campaignAiMaxTextCustomizationEnabled: true,
      campaignAiMaxFinalUrlExpansionEnabled: true,
      budgetResourceName: 'customers/4045056060/campaignBudgets/4001',
      campaignIsPortfolioBiddingStrategy: false,
      campaignPortfolioBiddingStrategyName: 'N/A',
      portfolioStrategyResourceName: 'N/A',
      portfolioOwnerId: 'N/A',
      isLocallyOwnedPortfolio: false,
      campaignPortfolioIsManagerOwned: false,
      campaignBiddingStrategyType: 'MAXIMIZE_CONVERSION_VALUE',
      campaignBiddingTargetRoas: 5.8,
      campaignBiddingTargetCpa: null,
      campaignStatsYesterdayCost: 1085,
      campaignStats7DaysCost: 7540,
      campaign30DaysCost: 32100,
      campaign30DaysConversionsValue: 192600,
      campaign30DaysConversions: 1280,
      campaign30DaysSearchRankLostImpressionShare: 0.06,
      campaign30DaysSearchBudgetLostImpressionShare: 0.31,
    },
    {
      mccId: SYNTHETIC_SECOND_MCC_ID,
      accountId: '404-505-6060',
      campaignResourceName: 'customers/4045056060/campaigns/44002',
      campaignId: '44002',
      campaignName: 'DE - Summer Holiday Routes PMax (Travel)',
      campaignType: 'PERFORMANCE_MAX',
      campaignSubType: 'UNSPECIFIED',
      campaignIsAiMax: false,
      campaignAiMaxTextCustomizationEnabled: false,
      campaignAiMaxFinalUrlExpansionEnabled: false,
      budgetResourceName: 'customers/4045056060/campaignBudgets/4002',
      campaignIsPortfolioBiddingStrategy: false,
      campaignPortfolioBiddingStrategyName: 'N/A',
      portfolioStrategyResourceName: 'N/A',
      portfolioOwnerId: 'N/A',
      isLocallyOwnedPortfolio: false,
      campaignPortfolioIsManagerOwned: false,
      campaignBiddingStrategyType: 'MAXIMIZE_CONVERSION_VALUE',
      campaignBiddingTargetRoas: 4.2,
      campaignBiddingTargetCpa: null,
      campaignStatsYesterdayCost: 1580,
      campaignStats7DaysCost: 10950,
      campaign30DaysCost: 46800,
      campaign30DaysConversionsValue: 205920,
      campaign30DaysConversions: 1640,
      campaign30DaysSearchRankLostImpressionShare: null,
      campaign30DaysSearchBudgetLostImpressionShare: null,
    },
    {
      mccId: SYNTHETIC_SECOND_MCC_ID,
      accountId: '505-606-7070',
      campaignResourceName: 'customers/5056067070/campaigns/55001',
      campaignId: '55001',
      campaignName: 'CH - Zurich & Geneva Hub Search',
      campaignType: 'SEARCH',
      campaignSubType: 'UNSPECIFIED',
      campaignIsAiMax: true,
      campaignAiMaxTextCustomizationEnabled: true,
      campaignAiMaxFinalUrlExpansionEnabled: false,
      budgetResourceName: 'customers/5056067070/campaignBudgets/5001',
      campaignIsPortfolioBiddingStrategy: true,
      campaignPortfolioBiddingStrategyName: 'CH Alpine Routes ROAS Portfolio',
      portfolioStrategyResourceName: 'customers/5056067070/biddingStrategies/9501',
      portfolioOwnerId: '5056067070',
      isLocallyOwnedPortfolio: true,
      campaignPortfolioIsManagerOwned: false,
      campaignBiddingStrategyType: 'TARGET_ROAS',
      campaignBiddingTargetRoas: 4.8,
      campaignBiddingTargetCpa: null,
      campaignStatsYesterdayCost: 790,
      campaignStats7DaysCost: 5460,
      campaign30DaysCost: 23400,
      campaign30DaysConversionsValue: 117000,
      campaign30DaysConversions: 690,
      campaign30DaysSearchRankLostImpressionShare: 0.17,
      campaign30DaysSearchBudgetLostImpressionShare: 0.07,
    },
  ];

  const recommendations = [
    {
      mccId: SYNTHETIC_SECOND_MCC_ID,
      accountId: '404-505-6060',
      recommendationId: '500008',
      recommendationType: 'CAMPAIGN_BUDGET',
      campaignResourceName: 'customers/4045056060/campaigns/44001',
      recommendationCurrentBudgetAmount: 1100,
      recommendationNewBudgetAmount: 1550,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 7700,
      recommendationPotentialCost: 10850,
      recommendationBaseClicks: 6400,
      recommendationPotentialClicks: 8900,
      recommendationBaseConversions: 310,
      recommendationPotentialConversions: 425,
      recommendationBaseConversionsValue: 46500,
      recommendationPotentialConversionsValue: 63750,
      recommendationBaseImpressions: 92000,
      recommendationPotentialImpressions: 128000,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    {
      mccId: SYNTHETIC_SECOND_MCC_ID,
      accountId: '404-505-6060',
      recommendationId: '500009',
      recommendationType: 'FORECASTING_CAMPAIGN_BUDGET',
      campaignResourceName: 'customers/4045056060/campaigns/44002',
      recommendationCurrentBudgetAmount: 1600,
      recommendationNewBudgetAmount: 2250,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: null,
      recommendationBaseCost: 11200,
      recommendationPotentialCost: 15750,
      recommendationBaseClicks: 8200,
      recommendationPotentialClicks: 11400,
      recommendationBaseConversions: 390,
      recommendationPotentialConversions: 535,
      recommendationBaseConversionsValue: 49140,
      recommendationPotentialConversionsValue: 66875,
      recommendationBaseImpressions: 215000,
      recommendationPotentialImpressions: 298000,
      recommendationBaseVideoViews: 14200,
      recommendationPotentialVideoViews: 19800,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
    {
      mccId: SYNTHETIC_SECOND_MCC_ID,
      accountId: '505-606-7070',
      recommendationId: '500010',
      recommendationType: 'LOWER_TARGET_ROAS',
      campaignResourceName: 'customers/5056067070/campaigns/55001',
      recommendationCurrentBudgetAmount: 820,
      recommendationNewBudgetAmount: 820,
      recommendationNewTargetCpa: null,
      recommendationNewTargetRoas: 4.1,
      recommendationBaseCost: 5500,
      recommendationPotentialCost: 7400,
      recommendationBaseClicks: 4100,
      recommendationPotentialClicks: 5550,
      recommendationBaseConversions: 165,
      recommendationPotentialConversions: 220,
      recommendationBaseConversionsValue: 27225,
      recommendationPotentialConversionsValue: 35200,
      recommendationBaseImpressions: 64000,
      recommendationPotentialImpressions: 86000,
      recommendationBaseVideoViews: 0,
      recommendationPotentialVideoViews: 0,
      moveBudgetSourceBudgetResourceName: '',
      moveBudgetSourceCampaigns: 'N/A',
      moveBudgetSourceTargetCpa: null,
      moveBudgetSourceTargetRoas: null,
    },
  ];

  return {
    accounts,
    budgets,
    campaigns,
    recommendations,
  };
}

/**
 * Seeds 2 synthetic Demo MCCs (`999-888-7777 [Demo] Global Retail MCC` and `888-777-6666 [Demo] DACH & Nordics Aviation MCC`),
 * their 5 multi-currency child accounts (`USD`, `EUR`, `GBP`, `CHF`), and 13 consolidated recommendations
 * into `_MCC_*`, `_Stage_*`, `Config`, and immediately publishes to `Data`.
 */
export function seedSyntheticDemoData(
  options: { spreadsheet?: SpreadsheetLike } = {}
): void {
  const ss: SpreadsheetLike = options.spreadsheet ?? SpreadsheetApp.getActiveSpreadsheet();
  const dateStr = resolveSpreadsheetDate(ss);

  const demoMccs = [
    {
      mccId: SYNTHETIC_MCC_ID,
      normMccId: SYNTHETIC_NORM_MCC_ID,
      mccName: SYNTHETIC_MCC_NAME,
      childAccounts: SYNTHETIC_CHILD_ACCOUNTS,
      batch: buildSyntheticDemoBatch(dateStr),
      recCountLabel: '10 recommendations',
    },
    {
      mccId: SYNTHETIC_SECOND_MCC_ID,
      normMccId: SYNTHETIC_SECOND_NORM_MCC_ID,
      mccName: SYNTHETIC_SECOND_MCC_NAME,
      childAccounts: SYNTHETIC_SECOND_MCC_CHILD_ACCOUNTS,
      batch: buildSyntheticSecondMccDemoBatch(dateStr),
      recCountLabel: '3 recommendations',
    },
  ];

  const demoNormMccSet = new Set(demoMccs.map((d) => d.normMccId));

  withSheetLock(ss, 'SHEET_SEED_DEMO', () => {
    const configSheet = getOrCreateSheet(ss, CONFIG_SHEET_NAME, { index: 0 });
    assertConfigLayout(configSheet);

    for (const demo of demoMccs) {
      clearMccFromStaging(ss, demo.normMccId);
      appendStagingBatch(ss, demo.batch, demo.normMccId);

      syncDiscoveredAccountsToConfig(
        configSheet,
        { mccId: demo.mccId, mccName: demo.mccName },
        demo.childAccounts
      );
    }

    // Check both synthetic MCCs and their 5 child accounts in Config so they publish immediately
    const mccCount = MCC_TABLE_END_ROW - MCC_TABLE_START_ROW + 1;
    const mccRows = configSheet
      .getRange(MCC_TABLE_START_ROW, 1, mccCount, sheetCol(MCC_COL.MCC_ID))
      .getValues();
    for (let i = 0; i < mccRows.length; i++) {
      if (demoNormMccSet.has(normalizeCid(mccRows[i][MCC_COL.MCC_ID]))) {
        configSheet.getRange(MCC_TABLE_START_ROW + i, sheetCol(MCC_COL.SELECT_ALL)).setValue(true);
      }
    }

    const lastRow = configSheet.getLastRow();
    if (lastRow >= CID_TABLE_START_ROW) {
      const cidCount = lastRow - CID_TABLE_START_ROW + 1;
      const cidRows = configSheet
        .getRange(CID_TABLE_START_ROW, 1, cidCount, sheetCol(CID_COL.PARENT_MCC_ID))
        .getValues();
      const nextCidColA = cidRows.map((r) => [
        demoNormMccSet.has(normalizeCid(r[CID_COL.PARENT_MCC_ID])) ? true : r[CID_COL.INCLUDE],
      ]);
      configSheet
        .getRange(CID_TABLE_START_ROW, sheetCol(CID_COL.INCLUDE), cidCount, 1)
        .setValues(nextCidColA);
    }

    // Reset today's publisher state (config lock, FX cache, publish signature) without touching real MCC
    // checkpoints, then write the COMPLETED checkpoint for both synthetic MCCs
    resetTodayPublisherState(ss, dateStr);
    for (const demo of demoMccs) {
      writeMccSheetState(ss, {
        mcc: { mccId: demo.mccId, mccName: demo.mccName },
        childAccounts: demo.childAccounts,
        statusMessage: `🧪 Demo MCC ${demo.mccId} seeded with ${demo.childAccounts.length} accounts & ${demo.recCountLabel} (${dateStr}).`,
        checkpoint: {
          mccId: demo.normMccId,
          dateStr,
          status: 'COMPLETED',
          completedCids: demo.childAccounts.map((a) => normalizeCid(a.customerId)),
          lastCompletedDate: dateStr,
        },
        lastCompletedDate: dateStr,
      });
    }
  });

  // The demo MCCs have COMPLETED checkpoints, so they publish fresh right away
  publishFromStagingToData({
    spreadsheet: ss,
    force: true,
  });
}
