//----------------------------------------------------------
//----------------------- IMPORTANT ------------------------
//----------------------------------------------------------
//------- YOU SHOULD NOT NEED TO EDIT ANYTHING BELOW -------
//----------------------------------------------------------

CONFIG = {
  // The template ID of the Looker Studio dashboard to be used for the report.
  // Don't change it unless you know what you're doing
  'LOOKER_TEMPLATE_ID': '0c3f445f-ae15-471c-a850-cfc1aaf53b83',

  // Set to true to see detailed logs, false for production.
  'debug_logs': false,

  // Set to true to see key timing details for API calls and processing.
  'show_timings_in_logs': true,

  // Set to true to see verbose timing details for debugging.
  'show_verbose_timings_in_logs': false,

  // Set to true to always send the setup email, even if the sheet already exists.
  'force_setup_email': false,

  // If false (default), the script will stop and ask you to share the existing
  // sheet with the current user.
  'force_new_sheet_on_access_error': false,
};



const MICROS = 1000000;

/**
 * Converts a currency value from micros to a standard unit.
 * @param {?number} micros The value in micros.
 * @return {number} The value in the standard currency unit, or 0 if the input
 * is invalid.
 */
function fromMicros(micros) {
  if (micros === null || typeof micros === 'undefined' || isNaN(micros)) {
    return 0;
  }
  return micros / MICROS;
}

/**
 * Logs a message to the console, respecting the DEBUG_MODE flag.
 * @param {string} message The message to log.
 * @param {string=} type The type of log, one of 'normal', 'debug', 'timing',
 * 'verbose_timing'.
 */
function log(message, type = 'normal') {
  if (type === 'debug' && !CONFIG.debug_logs) {
    return;
  }
  if (type === 'timing' && !CONFIG.show_timings_in_logs) {
    return;
  }
  if (type === 'verbose_timing' && !CONFIG.show_verbose_timings_in_logs) {
    return;
  }
  console.log(message);
}


const CUSTOM_RANGE_DAYS_AGO_START = 37;
const CUSTOM_RANGE_DAYS_AGO_END = 7;

const RECOMMENDATION_HANDLERS = {
  'CAMPAIGN_BUDGET': {
    /**
     * Extracts the recommendation target entities (in this case, campaigns)
     * from the recommendation.
     * @param {!Object} recommendation The recommendation object.
     * @param {!Object} recommendationDetails The pre-fetched details for this
     * recommendation.
     * @return {!Array<!Object>} An array of recommendation target objects.
     */
    getRecommendationTargets: function(recommendation, recommendationDetails) {
      if (recommendationDetails.campaigns &&
          recommendationDetails.campaigns.length > 0) {
        return recommendationDetails.campaigns.map(
            name => ({campaignResourceName: name}));
      } else if (recommendationDetails.campaign) {
        return [{campaignResourceName: recommendationDetails.campaign}];
      }
      return [];
    },

    /**
     * Adds details from the recommendationTarget object to the report row.
     * @param {!Object} rowData The report row to be enriched.
     * @param {!Object} recommendationTarget The specific target entity being
     * processed.
     * @param {!Object} recommendation The recommendation object.
     * @param {!Object} recommendationDetails The pre-fetched details for this
     * recommendation.
     * @return {!Object} The enriched rowData object.
     */
    addSpecificDetails: function(
        rowData, recommendationTarget, recommendation, recommendationDetails) {
      const budgetRecommendation =
          recommendationDetails.campaignBudgetRecommendation;
      if (budgetRecommendation) {
        rowData.recommendationCurrentBudgetAmount =
            fromMicros(budgetRecommendation.currentBudgetAmountMicros);
        rowData.recommendationNewBudgetAmount =
            fromMicros(budgetRecommendation.recommendedBudgetAmountMicros);
      }
      return rowData;
    },

    /**
     * Fetches the impact stats directly from the recommendation object.
     * @param {!Object} recommendation The recommendation object.
     * @param {!Object} recommendationDetails The pre-fetched details for this
     * recommendation.
     * @return {!Object} An object containing the base and potential metrics.
     */
    getImpactStats: function(recommendation, recommendationDetails) {
      const budgetRec = recommendationDetails.campaignBudgetRecommendation;
      if (budgetRec && budgetRec.budgetOptions) {
        const recommendedOption = budgetRec.budgetOptions.find(
            option => option.budgetAmountMicros ==
                budgetRec.recommendedBudgetAmountMicros);
        if (recommendedOption) {
          return {
            baseMetrics: recommendedOption.impact.baseMetrics,
            potentialMetrics: recommendedOption.impact.potentialMetrics,
          };
        }
      }
      // Return empty metrics if not found
      return {baseMetrics: null, potentialMetrics: null};
    },

    /**
     * Identifies any additional budget resource names required for this
     * recommendation type.
     * @param {!Object} recommendationDetails The pre-fetched details for this
     * recommendation.
     * @return {!Array<string>} An array of budget resource names.
     */
    getRequiredBudgetResourceNames: function(recommendationDetails) {
      // This type does not require fetching extra budget details.
      return [];
    }
  },
  'MOVE_UNUSED_BUDGET': {
    /**
     * Extracts the recommendationTarget entities (source and destination
     * campaigns) from the recommendation.
     * @param {!Object} recommendation The recommendation object.
     * @param {!Object} recommendationDetails The pre-fetched details for this
     * recommendation.
     * @return {!Array<!Object>} An array of recommendationTarget objects.
     */
    getRecommendationTargets: function(recommendation, recommendationDetails) {
      if (recommendationDetails.campaign) {
        return [{campaignResourceName: recommendationDetails.campaign}];
      }
      return [];
    },
    /**
     * Adds details from the recommendationTarget object to the report row.
     * @param {!Object} rowData The report row to be enriched.
     * @param {!Object} recommendationTarget The specific target entity being
     * processed.
     * @param {!Object} recommendation The recommendation object.
     * @param {!Object} recommendationDetails The pre-fetched details for this
     * recommendation.
     * @param {!Map<string, !Object>} budgetDataMap A map of budget resource
     * names to their details.
     * @param {!Map<string, !Array<string>>} budgetToCampaignsMap A map from
     * budget resource name to an array of campaign names.
     * @return {!Object} The enriched rowData object.
     */
    addSpecificDetails: function(
        rowData, recommendationTarget, recommendation, recommendationDetails,
        budgetDataMap, budgetToCampaignsMap) {
      const moveRec = recommendationDetails.moveUnusedBudgetRecommendation;
      if (moveRec) {
        if (moveRec.budgetRecommendation) {
          const budgetRec = moveRec.budgetRecommendation;
          rowData.recommendationCurrentBudgetAmount =
              fromMicros(budgetRec.currentBudgetAmountMicros);
          rowData.recommendationNewBudgetAmount =
              fromMicros(budgetRec.recommendedBudgetAmountMicros);
          rowData.moveBudgetAmount = fromMicros(
              budgetRec.recommendedBudgetAmountMicros -
              budgetRec.currentBudgetAmountMicros);
        }

        const excessBudgetResourceName = moveRec.excessCampaignBudget;
        if (excessBudgetResourceName &&
            budgetDataMap.has(excessBudgetResourceName)) {
          const budgetDetails = budgetDataMap.get(excessBudgetResourceName);
          rowData.moveBudgetSourceBudgetName = budgetDetails.campaignBudgetName;
          rowData.moveBudgetSourceBudgetType = budgetDetails.campaignBudgetType;
          rowData.moveBudgetSourceBudgetDeliveryMethod =
              budgetDetails.campaignBudgetDeliveryMethod;
          rowData.moveBudgetSourceBudgetIsShared =
              budgetDetails.campaignBudgetIsShared;
          rowData.moveBudgetSourceBudgetAmount =
              budgetDetails.campaignBudgetAmount;
        }
        if (excessBudgetResourceName &&
            budgetToCampaignsMap.has(excessBudgetResourceName)) {
          rowData.moveBudgetSourceCampaigns =
              budgetToCampaignsMap.get(excessBudgetResourceName).join(', ');
        }
      }
      return rowData;
    },

    /**
     * Fetches the impact stats from the nested details of the recommendation.
     * @param {!Object} recommendation The recommendation object.
     * @param {!Object} recommendationDetails The pre-fetched details for this
     * recommendation.
     * @return {!Object} An object containing the base and potential metrics.
     */
    getImpactStats: function(recommendation, recommendationDetails) {
      const moveRec = recommendationDetails.moveUnusedBudgetRecommendation;
      if (moveRec && moveRec.budgetRecommendation) {
        const budgetRec = moveRec.budgetRecommendation;
        const recommendedOption = budgetRec.budgetOptions.find(
            option => option.budgetAmountMicros ==
                budgetRec.recommendedBudgetAmountMicros);
        if (recommendedOption) {
          return {
            baseMetrics: recommendedOption.impact.baseMetrics,
            potentialMetrics: recommendedOption.impact.potentialMetrics,
          };
        }
      }
      // Return empty metrics if not found
      return {baseMetrics: null, potentialMetrics: null};
    },

    /**
     * Identifies any additional budget resource names required for this
     * recommendation type.
     * @param {!Object} recommendationDetails The pre-fetched details for this
     * recommendation.
     * @return {!Array<string>} An array of budget resource names.
     */
    getRequiredBudgetResourceNames: function(recommendationDetails) {
      if (recommendationDetails.moveUnusedBudgetRecommendation &&
          recommendationDetails.moveUnusedBudgetRecommendation
              .excessCampaignBudget) {
        return [recommendationDetails.moveUnusedBudgetRecommendation
                    .excessCampaignBudget];
      }
      return [];
    }
  },
  'RAISE_TARGET_CPA': {
    /**
     * Extracts the recommendation target entities (campaigns) from the
     * recommendation.
     * @param {!Object} recommendation The recommendation object.
     * @param {!Object} recommendationDetails The pre-fetched details for this
     *     recommendation.
     * @return {!Array<!Object>} An array of recommendation target objects.
     */
    getRecommendationTargets: function(recommendation, recommendationDetails) {
      if (recommendationDetails.campaigns &&
          recommendationDetails.campaigns.length > 0) {
        return recommendationDetails.campaigns.map(
            name => ({campaignResourceName: name}));
      } else if (recommendationDetails.campaign) {
        return [{campaignResourceName: recommendationDetails.campaign}];
      }
      return [];
    },
    addSpecificDetails: function(
        rowData, recommendationTarget, recommendation, recommendationDetails) {
      const cpaRec = recommendationDetails.raiseTargetCpaRecommendation?.targetAdjustment;
      if (cpaRec) {
        const currentTargetMicros = cpaRec.currentAverageTargetMicros;
        const recommendedMultiplier = cpaRec.recommendedTargetMultiplier;
        const recommendedTargetMicros = currentTargetMicros * recommendedMultiplier;
        rowData.recommendationNewTargetCpa = fromMicros(recommendedTargetMicros);
      }
      return rowData;
    },
    getImpactStats: function(recommendation, recommendationDetails) {
      const cpaImpact = recommendationDetails.impact;
      if (cpaImpact) {
        return {
          baseMetrics: cpaImpact.baseMetrics,
          potentialMetrics: cpaImpact.potentialMetrics,
        };
      }
      return {baseMetrics: null, potentialMetrics: null};
    },
    getRequiredBudgetResourceNames: function(recommendationDetails) {
      return [];
    }
  },
  'LOWER_TARGET_ROAS': {
    getRecommendationTargets: function(recommendation, recommendationDetails) {
      if (recommendationDetails.campaigns &&
          recommendationDetails.campaigns.length > 0) {
        return recommendationDetails.campaigns.map(
            name => ({campaignResourceName: name}));
      } else if (recommendationDetails.campaign) {
        return [{campaignResourceName: recommendationDetails.campaign}];
      }
      return [];
    },
    addSpecificDetails: function(
        rowData, recommendationTarget, recommendation, recommendationDetails) {
      const roasRec = recommendationDetails.lowerTargetRoasRecommendation?.targetAdjustment;
      if (roasRec) {
        const currentTargetMicros = roasRec.currentAverageTargetMicros;
        const recommendedMultiplier = roasRec.recommendedTargetMultiplier;
        const recommendedTargetMicros = currentTargetMicros * recommendedMultiplier;
        rowData.recommendationNewTargetRoas = fromMicros(recommendedTargetMicros);
      }
      return rowData;
    },
    getImpactStats: function(recommendation, recommendationDetails) {
      const roasImpact = recommendationDetails.impact;
      if (roasImpact) {
        return {
          baseMetrics: roasImpact.baseMetrics,
          potentialMetrics: roasImpact.potentialMetrics,
        };
      }
      return {baseMetrics: null, potentialMetrics: null};
    },
    getRequiredBudgetResourceNames: function(recommendationDetails) {
      return [];
    }
  }
};


const REPORT_SCHEMA_TEMPLATE = {
  // --- Basic Info (Should be always available) ---
  accountId: {defaultValue: 'N/A', format: '@'},
  accountName: {defaultValue: 'N/A', format: '@'},
  timestamp: {defaultValue: null, format: 'yyyy-mm-dd hh:mm:ss'},
  recommendationId: {defaultValue: 'N/A', format: '@'},
  recommendationType: {defaultValue: 'N/A', format: '@'},
  recommendationsDetailsUrl: {defaultValue: 'N/A', format: '@'},
  // --- Default values for all other fields ---
  campaignUrl: {defaultValue: 'N/A', format: '@'},
  campaignId: {defaultValue: 'N/A', format: '@'},
  campaignName: {defaultValue: 'N/A', format: '@'},
  campaignType: {defaultValue: 'N/A', format: '@'},
  campaignSubType: {defaultValue: 'N/A', format: '@'},
  campaignIsAiMax: {defaultValue: false, format: '@'},
  campaignAiMaxTextCustomizationEnabled: {defaultValue: false, format: '@'},
  campaignAiMaxFinalUrlExpansionEnabled: {defaultValue: false, format: '@'},
  currencyCode: {defaultValue: 'N/A', format: '@'},
  campaignBudgetName: {defaultValue: 'N/A', format: '@'},
  campaignBudgetAmount: {defaultValue: null, format: '#,##0.00'},
  campaignBudgetTotalAmount: {defaultValue: null, format: '#,##0.00'},
  campaignBudgetDeliveryMethod: {defaultValue: 'N/A', format: '@'},
  campaignBudgetType: {defaultValue: 'N/A', format: '@'},
  campaignBudgetIsShared: {defaultValue: 'N/A', format: '@'},
  campaignIsPortfolioBiddingStrategy: {defaultValue: 'N/A', format: '@'},
  campaignPortfolioBiddingStrategyName: {defaultValue: 'N/A', format: '@'},
  campaignPortfolioIsManagerOwned: {defaultValue: 'N/A', format: '@'},
  campaignBiddingStrategyType: {defaultValue: 'N/A', format: '@'},
  campaignBiddingTargetRoas: {defaultValue: null, format: '0.00'},
  campaignBiddingTargetCpa: {defaultValue: null, format: '#,##0.00'},
  campaignStatsYesterdayCost: {defaultValue: null, format: '#,##0.00'},
  campaignStats7DaysCost: {defaultValue: null, format: '#,##0.00'},
  campaign30DaysCost: {defaultValue: null, format: '#,##0.00'},
  campaign30DaysConversionsValue: {defaultValue: null, format: '#,##0.00'},
  campaign30DaysConversions: {defaultValue: null, format: '#,##0'},
  campaign30DaysRoas: {defaultValue: null, format: '0.00'},
  campaign30DaysCpa: {defaultValue: null, format: '#,##0.00'},
  campaign30DaysAvgCpm: {defaultValue: null, format: '#,##0.00'},
  campaign30DaysAvgCpv: {defaultValue: null, format: '#,##0.00'},
  campaign30DaysUniqueUsers: {defaultValue: null, format: '#,##0'},
  campaign30DaysAvgTargetRoas: {defaultValue: null, format: '0.00'},
  campaign30DaysAvgTargetCpa: {defaultValue: null, format: '#,##0.00'},
  campaign30DaysSearchRankLostImpressionShare:
      {defaultValue: null, format: '0.00%'},
  campaign30DaysSearchBudgetLostImpressionShare:
      {defaultValue: null, format: '0.00%'},
  campaignCustomRangeCost: {defaultValue: null, format: '#,##0.00'},
  campaignCustomRangeConversionsValue: {defaultValue: null, format: '#,##0.00'},
  campaignCustomRangeConversions: {defaultValue: null, format: '#,##0'},
  campaignCustomRangeRoas: {defaultValue: null, format: '0.00'},
  campaignCustomRangeCpa: {defaultValue: null, format: '#,##0.00'},
  campaignCustomRangeAvgCpm: {defaultValue: null, format: '#,##0.00'},
  campaignCustomRangeAvgCpv: {defaultValue: null, format: '#,##0.00'},
  campaignCustomRangeUniqueUsers: {defaultValue: null, format: '#,##0'},
  campaignCustomRangeAvgTargetRoas: {defaultValue: null, format: '0.00'},
  campaignCustomRangeAvgTargetCpa: {defaultValue: null, format: '#,##0.00'},
  campaignCustomRangeStartDate: {defaultValue: 'N/A', format: 'yyyy-mm-dd'},
  campaignCustomRangeEndDate: {defaultValue: 'N/A', format: 'yyyy-mm-dd'},
  campaignCustomRangeSearchRankLostImpressionShare:
      {defaultValue: null, format: '0.00%'},
  campaignCustomRangeSearchBudgetLostImpressionShare:
      {defaultValue: null, format: '0.00%'},
  recommendationCurrentBudgetAmount: {defaultValue: null, format: '#,##0.00'},
  recommendationNewBudgetAmount: {defaultValue: null, format: '#,##0.00'},
  recommendationNewTargetCpa: {defaultValue: null, format: '#,##0.00'},
  recommendationNewTargetRoas: {defaultValue: null, format: '0.00'},
  recommendationBaseCost: {defaultValue: null, format: '#,##0.00'},
  recommendationPotentialCost: {defaultValue: null, format: '#,##0.00'},
  recommendationBaseClicks: {defaultValue: null, format: '#,##0'},
  recommendationPotentialClicks: {defaultValue: null, format: '#,##0'},
  recommendationBaseConversions: {defaultValue: null, format: '#,##0'},
  recommendationPotentialConversions: {defaultValue: null, format: '#,##0'},
  recommendationBaseConversionsValue: {defaultValue: null, format: '#,##0'},
  recommendationPotentialConversionsValue: {defaultValue: null, format: '#,##0'},
  recommendationBaseCpa: {defaultValue: null, format: '#,##0.00'},
  recommendationPotentialCpa: {defaultValue: null, format: '#,##0.00'},
  recommendationBaseRoas: {defaultValue: null, format: '0.00'},
  recommendationPotentialRoas: {defaultValue: null, format: '0.00'},
  recommendationBaseImpressions: {defaultValue: null, format: '#,##0'},
  recommendationPotentialImpressions: {defaultValue: null, format: '#,##0'},
  recommendationBaseVideoViews: {defaultValue: null, format: '#,##0'},
  recommendationPotentialVideoViews: {defaultValue: null, format: '#,##0'},
  campaignCalculated30DaysPotentialConversionValue:
      {defaultValue: null, format: '#,##0.00'},
  campaignCalculatedCustomRangePotentialConversionValue:
      {defaultValue: null, format: '#,##0.00'},
  campaignPercentageOfBudgetUsedYesterday:
      {defaultValue: null, format: '0.00%'},
  weeklyCostIncrease: {defaultValue: null, format: '#,##0.00'},
  newWeeklyConversions: {defaultValue: null, format: '#,##0'},
  dailyBudget30Days: {defaultValue: null, format: '#,##0.00'},
  dailyBudgetDelta: {defaultValue: null, format: '#,##0.00'},
  targetRoasDelta: {defaultValue: null, format: '0.00'},
  targetCpaDelta: {defaultValue: null, format: '#,##0.00'},
  recommendationCpaDelta: {defaultValue: null, format: '#,##0.00'},
  recommendationRoasDelta: {defaultValue: null, format: '0.00'},
  // --- Fields for MOVE_UNUSED_BUDGET recommendations ---
  moveBudgetAmount: {defaultValue: 0, format: '#,##0.00'},
  moveBudgetSourceCampaigns: {defaultValue: 'N/A', format: '@'},
  moveBudgetSourceBudgetName: {defaultValue: 'N/A', format: '@'},
  moveBudgetSourceBudgetType: {defaultValue: 'N/A', format: '@'},
  moveBudgetSourceBudgetDeliveryMethod: {defaultValue: 'N/A', format: '@'},
  moveBudgetSourceBudgetIsShared: {defaultValue: 'N/A', format: '@'},
  moveBudgetSourceBudgetAmount: {defaultValue: 0, format: '#,##0.00'},
};


/**
 * Creates a new, blank object for a report row.
 *
 * @param {string} accountId The ID of the account.
 * @param {string} accountName The name of the account.
 * @param {!Object} recommendation The recommendation object.
 * @param {string} currencyCode The currency code of the account.
 * @return {!Object} The initialized row data object.
 */
function createReportRow(
    accountId, accountName, recommendation, currencyCode) {
  const row = {};
  for (const key in REPORT_SCHEMA_TEMPLATE) {
    if (REPORT_SCHEMA_TEMPLATE.hasOwnProperty(key)) {
      row[key] = REPORT_SCHEMA_TEMPLATE[key].defaultValue;
    }
  }

  row.accountId = accountId;
  row.accountName = accountName;
  row.currencyCode = currencyCode;
  row.timestamp = new Date();
  row.recommendationId = recommendation.getResourceName();
  row.recommendationType = recommendation.getType();
  const {ocid, recoTypeId} = getInfoFromRecoId(recommendation.getResourceName());
  row.recommendationsDetailsUrl = `https://ads.google.com/aw/recommendations?ocid=${ocid}&opp=${recoTypeId}`;
  return row;
}


/**
 * Fetches and stores all recommendation details in a map for efficient lookup.
 * @param {!Array<!Object>} recommendations An array of recommendation objects.
 * @return {!Map<string, !Object>} A map from recommendation resource name to
 * its details.
 */
function fetchAllRecommendationDetails(recommendations) {
  const recommendationDetailsMap = new Map();
  if (recommendations.length === 0) {
    return recommendationDetailsMap;
  }

  const recommendationIds =
      recommendations.map(r => `'${r.getResourceName()}'`).join(',');

  const query = `
      SELECT
        recommendation.resource_name,
        recommendation.campaign,
        recommendation.campaigns,
        recommendation.campaign_budget_recommendation,
        recommendation.move_unused_budget_recommendation,
        recommendation.raise_target_cpa_recommendation,
        recommendation.lower_target_roas_recommendation,
        recommendation.impact
      FROM recommendation
      WHERE recommendation.resource_name IN (${recommendationIds})`;

  try {
    const result = AdsApp.search(query);
    for (const row of result) {
      recommendationDetailsMap.set(
          row.recommendation.resourceName, row.recommendation);
    }
  } catch (e) {
    log(`Could not fetch recommendation details in bulk. Error: ${e.message}`);
  }
  return recommendationDetailsMap;
}


/**
 * Fetches and stores all budget details in a map for efficient lookup.
 * @param {!Array<string>} budgetResourceNames An array of budget resource
 * names.
 * @return {!Map<string, !Object>} A map from budget resource name to its
 * details.
 */
function fetchAllBudgetDetails(budgetResourceNames) {
  const budgetDataMap = new Map();
  if (budgetResourceNames.length === 0) {
    return budgetDataMap;
  }

  const budgetIds = budgetResourceNames.map(name => `'${name}'`).join(',');

  const query = `
      SELECT
        campaign_budget.resource_name,
        campaign_budget.amount_micros,
        campaign_budget.name,
        campaign_budget.period,
        campaign_budget.delivery_method,
        campaign_budget.explicitly_shared
      FROM campaign_budget
      WHERE campaign_budget.resource_name IN (${budgetIds})`;

  try {
    const result = AdsApp.search(query);
    for (const row of result) {
      const budgetDetails = extractBudgetDetails(row.campaignBudget);
      budgetDataMap.set(row.campaignBudget.resourceName, budgetDetails);
    }
  } catch (e) {
    log(`Could not fetch budget details in bulk. Error: ${e.message}`);
  }
  return budgetDataMap;
}


/**
 * Fetches campaign names associated with a given list of budget resource names.
 * @param {!Array<string>} budgetResourceNames An array of budget resource
 * names.
 * @return {!Map<string, !Array<string>>} A map where keys are budget resource
 * names and values are arrays of associated campaign names.
 */
function fetchCampaignsByBudgets(budgetResourceNames) {
  const budgetToCampaignsMap = new Map();
  if (budgetResourceNames.length === 0) {
    return budgetToCampaignsMap;
  }

  const budgetIds = budgetResourceNames.map(name => `'${name}'`).join(',');

  const query = `
      SELECT
        campaign.name,
        campaign_budget.resource_name
      FROM campaign
      WHERE campaign_budget.resource_name IN (${budgetIds})`;

  try {
    const result = AdsApp.search(query);
    for (const row of result) {
      const budgetResourceName = row.campaignBudget.resourceName;
      if (!budgetToCampaignsMap.has(budgetResourceName)) {
        budgetToCampaignsMap.set(budgetResourceName, []);
      }
      budgetToCampaignsMap.get(budgetResourceName).push(row.campaign.name);
    }
  } catch (e) {
    log(`Could not fetch campaigns by budgets in bulk. Error: ${e.message}`);
  }
  return budgetToCampaignsMap;
}


/**
 * Extracts and transforms budget details from a campaignBudget object.
 * @param {!Object} campaignBudget The budget object from a GAQL query result.
 * @return {!Object} A structured object with budget details.
 */
function extractBudgetDetails(campaignBudget) {
  if (!campaignBudget) {
    return {
      campaignBudgetAmount: null,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'N/A',
      campaignBudgetType: 'N/A',
      campaignBudgetIsShared: 'N/A',
      campaignBudgetName: 'N/A',
    };
  }
  return {
    campaignBudgetAmount: campaignBudget.amountMicros ?
        fromMicros(campaignBudget.amountMicros) :
        null,
    campaignBudgetTotalAmount: campaignBudget.totalAmountMicros ?
        fromMicros(campaignBudget.totalAmountMicros) :
        null,
    campaignBudgetDeliveryMethod: campaignBudget.deliveryMethod || 'N/A',
    campaignBudgetType: campaignBudget.period || 'N/A',
    campaignBudgetIsShared: campaignBudget.explicitlyShared,
    campaignBudgetName: campaignBudget.name || 'N/A',
  };
}


/**
 * Extracts and transforms the potential performance impact from a metrics
 * object into a structured format for the report.
 *
 * @param {{baseMetrics: !Object, potentialMetrics: !Object}} impactMetrics An
 * object containing the base and potential metrics, as returned by a
 * handler's getImpactStats function.
 * @return {!Object} An object containing the recommendation's base and
 * potential stats, formatted for the report.
 */
function extractRecommendationImpactStats(impactMetrics) {
  const startTime = new Date().getTime();
  const details = {
    recommendationBaseCost: null,
    recommendationPotentialCost: null,
    recommendationBaseClicks: null,
    recommendationPotentialClicks: null,
    recommendationBaseConversions: null,
    recommendationPotentialConversions: null,
    recommendationBaseImpressions: null,
    recommendationPotentialImpressions: null,
    recommendationBaseVideoViews: null,
    recommendationPotentialVideoViews: null,
    recommendationBaseConversionsValue: null,
    recommendationPotentialConversionsValue: null,
    recommendationBaseCpa: null,
    recommendationPotentialCpa: null,
    recommendationBaseRoas: null,
    recommendationPotentialRoas: null,
  };

  try {
    const {baseMetrics, potentialMetrics} = impactMetrics;

    if (baseMetrics && potentialMetrics) {
      details.recommendationBaseCost = fromMicros(baseMetrics.costMicros);
      details.recommendationPotentialCost =
          fromMicros(potentialMetrics.costMicros);
      details.recommendationBaseClicks = baseMetrics.clicks;
      details.recommendationPotentialClicks = potentialMetrics.clicks;
      details.recommendationBaseConversions = baseMetrics.conversions;
      details.recommendationPotentialConversions = potentialMetrics.conversions;
      details.recommendationBaseImpressions = baseMetrics.impressions;
      details.recommendationPotentialImpressions = potentialMetrics.impressions;
      details.recommendationBaseVideoViews = baseMetrics.videoViews;
      details.recommendationPotentialVideoViews = potentialMetrics.videoViews;
      details.recommendationBaseConversionsValue = baseMetrics.conversionsValue;
      details.recommendationPotentialConversionsValue = potentialMetrics.conversionsValue;
      details.recommendationBaseCpa = details.recommendationBaseConversions > 0 ? details.recommendationBaseCost / details.recommendationBaseConversions : 0;
      details.recommendationPotentialCpa = details.recommendationPotentialConversions > 0 ? details.recommendationPotentialCost / details.recommendationPotentialConversions : 0;
      details.recommendationBaseRoas = details.recommendationBaseCost > 0 ? details.recommendationBaseConversionsValue / details.recommendationBaseCost : 0;
      details.recommendationPotentialRoas = details.recommendationPotentialCost > 0 ? details.recommendationPotentialConversionsValue / details.recommendationPotentialCost : 0;
    }
  } catch (e) {
    log(`Could not extract base/potential stats. Error: ${e.message}`);
  }

  log(`extractRecommendationImpactStats took ${
          ((new Date().getTime() - startTime) / 1000).toFixed(2)}s`,
      'verbose_timing');
  return details;
}

/**
 * Calculates the potential conversion value based on current conversion value,
 * conversions, and potential conversions.
 * @param {number} conversionsValue The current conversion value.
 * @param {number} conversions The current number of conversions.
 * @param {number} potentialConversions The potential number of conversions.
 * @return {number|string} The calculated potential conversion value, or 'N/A'
 * if inputs are invalid.
 */
function calculatePotentialConversionValue(
    conversionsValue, conversions, potentialConversions) {
  // Safety Check: Ensure all inputs are valid numbers before performing
  // math.
  if (typeof conversionsValue === 'number' &&
      typeof conversions === 'number' && conversions > 0 &&
      typeof potentialConversions === 'number') {
    const averageConversionValue = conversionsValue / conversions;
    const potentialConversionValue =
        averageConversionValue * potentialConversions;
    log(`Calculated potential conversions value ${potentialConversionValue}.`,
        'debug');
    return potentialConversionValue;
  }

    return null;

  }


/**
 * Calculates the difference between two numbers.
 * @param {number} value1 The first number.
 * @param {number} value2 The second number.
 * @return {number|null} The difference between value1 and value2, or null if
 * either input is not a number.
 */
  function calculateDelta(value1, value2) {
    if (typeof value1 === 'number' && typeof value2 === 'number') {
      return value1 - value2;
    }
    return null;
  }

/**

* Fetches campaign metrics for a specific date range.

* @param {!Array<string>} campaignResourceNames The resource names of the
* campaigns.
* @param {string} dateCondition The GAQL date condition string (e.g., "DURING
* LAST_30_DAYS").
* @return {!Object} An object containing the campaign's metrics for the given
* date range.
*/
function fetchCampaignMetricsByDateRange(campaignResourceNames, dateCondition) {
  const startTime = new Date().getTime();
  const metricsMap = new Map();

  if (campaignResourceNames.length === 0) {
    return metricsMap;
  }

  const campaignIds = campaignResourceNames.map(name => `'${name}'`).join(',');

  try {
    const query = `
            SELECT
              campaign.resource_name,
              metrics.cost_micros,
              metrics.conversions_value,
              metrics.conversions,
              metrics.search_rank_lost_impression_share,
              metrics.search_budget_lost_impression_share,
              metrics.average_target_roas,
              metrics.average_target_cpa_micros,
              metrics.average_cpm,
              metrics.trueview_average_cpv,
              metrics.unique_users
            FROM campaign
            WHERE campaign.resource_name IN (${campaignIds})
            AND segments.date ${dateCondition}`;

    const result = AdsApp.search(query);
    for (const row of result) {
      const metrics = {
        cost: fromMicros(row.metrics.costMicros || 0),
        conversionsValue: row.metrics.conversionsValue || 0,
        conversions: row.metrics.conversions || 0,
        searchRankLostImpressionShare:
            row.metrics.searchRankLostImpressionShare || 0,
        searchBudgetLostImpressionShare:
            row.metrics.searchBudgetLostImpressionShare || 0,
        avgTargetRoas: row.metrics.averageTargetRoas || 0,
        avgTargetCpa: fromMicros(row.metrics.averageTargetCpaMicros || 0),
        averageCpm: fromMicros(row.metrics.averageCpm || 0),
        averageCpv: fromMicros(row.metrics.trueviewAverageCpv || 0),
        uniqueUsers: row.metrics.uniqueUsers || 0,
        roas: 0,
        cpa: 0,
      };
      if (metrics.cost > 0) {
        metrics.roas = metrics.conversionsValue / metrics.cost;
      }
      if (metrics.conversions > 0) {
        metrics.cpa = metrics.cost / metrics.conversions;
      }
      metricsMap.set(row.campaign.resourceName, metrics);
    }
  } catch (e) {
    log(`Warning: Could not fetch metrics with date condition '${
        dateCondition}'. Error: ${e.message}`);
  }
  log(`fetchCampaignMetricsByDateRange for ${dateCondition} took ${
          ((new Date().getTime() - startTime) / 1000).toFixed(2)}s`,
      'verbose_timing');
  return metricsMap;
}

/**
 * Fetches bidding strategy metrics for a specific date range.
 * @param {!Array<string>} biddingStrategyResourceNames The resource names of the
 * bidding strategies.
 * @param {string} dateCondition The GAQL date condition string (e.g., "DURING
 * LAST_30_DAYS").
 * @return {!Object} An object containing the bidding strategy's metrics for the
 * given date range.
 */
function fetchBiddingStrategyMetricsByDateRange(
    biddingStrategyResourceNames, dateCondition) {
  const startTime = new Date().getTime();
  const metricsMap = new Map();
  if (biddingStrategyResourceNames.length === 0) {
    return metricsMap;
  }
  log(`Fetching historical metrics for ${
      biddingStrategyResourceNames
          .length} portfolio strategies (${dateCondition})...`);
  const resourceNames =
      biddingStrategyResourceNames.map(name => `'${name}'`).join(',');
  try {
    const query = `
            SELECT
              bidding_strategy.resource_name,
              metrics.average_target_roas,
              metrics.average_target_cpa_micros
            FROM bidding_strategy
            WHERE bidding_strategy.resource_name IN (${resourceNames})
            AND segments.date ${dateCondition}`;

    const result = AdsApp.search(query);
    for (const row of result) {
      metricsMap.set(row.biddingStrategy.resourceName, {
        avgTargetRoas: row.metrics.averageTargetRoas,
        avgTargetCpa: row.metrics.averageTargetCpaMicros ?
            fromMicros(row.metrics.averageTargetCpaMicros) :
            null
      });
    }
  } catch (e) {
    log(`Warning: Could not fetch bidding strategy metrics with date condition '${
        dateCondition}'. Error: ${e.message}`);
  }
  log(`fetchBiddingStrategyMetricsByDateRange for ${dateCondition} took ${
          ((new Date().getTime() - startTime) / 1000).toFixed(2)}s`,
      'verbose_timing');
  return metricsMap;
}


/**
 * Main orchestrator for fetching all data for a single campaign.
 * the necessary data, and handles potential errors gracefully.
 * @param {!Array<string>} campaignResourceNames The resource names of the
 * campaigns.
 * @return {!Object} An object containing the campaign's details.
 */
function fetchAllCampaignData(campaignResourceNames) {
  const startTime = new Date().getTime();
  const campaignDataMap = new Map();

  if (campaignResourceNames.length === 0) {
    return campaignDataMap;
  }

  const today = new Date();
  const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  const startDate =
      new Date(today.getTime() - CUSTOM_RANGE_DAYS_AGO_START * MILLIS_PER_DAY);
  const endDate =
      new Date(today.getTime() - CUSTOM_RANGE_DAYS_AGO_END * MILLIS_PER_DAY);
  const timeZone = AdsApp.currentAccount().getTimeZone();
  const formattedStartDate =
      Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd');
  const formattedEndDate =
      Utilities.formatDate(endDate, timeZone, 'yyyy-MM-dd');

  const campaignIds = campaignResourceNames.map(name => `'${name}'`).join(',');

  // Initialize map with basic details
  for (const name of campaignResourceNames) {
    const campaignId = name.split('/')[3];
    campaignDataMap.set(name, {
      campaignId: campaignId,
      campaignName: 'N/A',
      campaignType: 'N/A',
      campaignSubType: 'N/A',
      campaignIsAiMax: false,
      campaignAiMaxTextCustomizationEnabled: false,
      campaignAiMaxFinalUrlExpansionEnabled: false,
      campaignBudgetAmount: null,
      campaignBudgetTotalAmount: null,
      campaignBudgetDeliveryMethod: 'N/A',
      campaignBudgetType: 'N/A',
      campaignBudgetIsShared: 'N/A',
      campaignIsPortfolioBiddingStrategy: 'N/A',
      campaignPortfolioBiddingStrategyName: 'N/A',
      campaignPortfolioIsManagerOwned: 'N/A',
      campaignBiddingStrategyType: 'N/A',
      campaignBiddingTargetRoas: null,
      campaignBiddingTargetCpa: null,
      campaignStatsYesterdayCost: null,
      campaignStats7DaysCost: null,
      campaign30DaysCost: null,
      campaign30DaysConversionsValue: null,
      campaign30DaysConversions: null,
      campaign30DaysRoas: null,
      campaign30DaysCpa: null,
      campaign30DaysAvgCpm: null,
      campaign30DaysAvgCpv: null,
      campaign30DaysUniqueUsers: null,
      campaign30DaysAvgTargetRoas: null,
      campaign30DaysAvgTargetCpa: null,
      campaign30DaysSearchRankLostImpressionShare: null,
      campaign30DaysSearchBudgetLostImpressionShare: null,
      campaignCustomRangeCost: null,
      campaignCustomRangeConversionsValue: null,
      campaignCustomRangeConversions: null,
      campaignCustomRangeRoas: null,
      campaignCustomRangeCpa: null,
      campaignCustomRangeAvgCpm: null,
      campaignCustomRangeAvgCpv: null,
      campaignCustomRangeUniqueUsers: null,
      campaignCustomRangeAvgTargetRoas: null,
      campaignCustomRangeAvgTargetCpa: null,
      campaignCustomRangeStartDate: formattedStartDate,
      campaignCustomRangeEndDate: formattedEndDate,
      campaignCustomRangeSearchRankLostImpressionShare: null,
      campaignCustomRangeSearchBudgetLostImpressionShare: null,
      portfolioStrategyResourceName: null,
      portfolioOwnerId: null,
      isLocallyOwnedPortfolio: false,
    });
  }


  // --- GAQL Query 1: Main Query ---
  const mainQuery = `
        SELECT
          customer.id,
          campaign.resource_name,
          campaign.name,
          campaign.advertising_channel_type,
          campaign.advertising_channel_sub_type,
          campaign.ai_max_setting.enable_ai_max,
          campaign.asset_automation_settings,
          campaign.bidding_strategy_type,
          campaign.bidding_strategy,
          campaign.maximize_conversions.target_cpa_micros,
          campaign.maximize_conversion_value.target_roas,
          campaign.target_cpa.target_cpa_micros,
          campaign.target_roas.target_roas,
          bidding_strategy.name,
          bidding_strategy.target_cpa.target_cpa_micros,
          bidding_strategy.target_roas.target_roas,
          campaign_budget.amount_micros,
          campaign_budget.total_amount_micros,
          campaign_budget.delivery_method,
          campaign_budget.type,
          campaign_budget.name,
          campaign_budget.period,
          campaign_budget.explicitly_shared,
          accessible_bidding_strategy.name,
          accessible_bidding_strategy.owner_customer_id,
          accessible_bidding_strategy.id,
          accessible_bidding_strategy.maximize_conversions.target_cpa_micros,
          accessible_bidding_strategy.maximize_conversion_value.target_roas,
          accessible_bidding_strategy.target_cpa.target_cpa_micros,
          accessible_bidding_strategy.target_roas.target_roas,
          accessible_bidding_strategy.type
        FROM campaign
        WHERE campaign.resource_name IN (${campaignIds})`;

  const mainQueryStartTime = new Date().getTime();
  try {
    const result = AdsApp.search(mainQuery);
    for (const row of result) {
      const campaignDetails = campaignDataMap.get(row.campaign.resourceName);
      campaignDetails.campaignName = row.campaign.name;
      campaignDetails.campaignType = row.campaign.advertisingChannelType;
      campaignDetails.campaignSubType = row.campaign.advertisingChannelSubType;
      campaignDetails.campaignIsAiMax =
          row.campaign.aiMaxSetting?.enableAiMax || false;

      // Extract Asset Automation Settings
      const rawSettings = row.campaign.assetAutomationSettings || [];
      let settings = [];
      if (Array.isArray(rawSettings)) {
        settings = rawSettings;
      } else if (typeof rawSettings === 'object' && rawSettings !== null) {
        settings = Object.values(rawSettings);
      }
      for (const setting of settings) {
        const type = setting.assetAutomationType;
        const status = setting.assetAutomationStatus;
        if (type === 'TEXT_ASSET_AUTOMATION') {
          campaignDetails.campaignAiMaxTextCustomizationEnabled =
              status === 'OPTED_IN';
        } else if (type === 'FINAL_URL_EXPANSION_TEXT_ASSET_AUTOMATION') {
          campaignDetails.campaignAiMaxFinalUrlExpansionEnabled =
              status === 'OPTED_IN';
        }
      }

      campaignDetails.campaignBiddingStrategyType =
          row.campaign.biddingStrategyType;

      if (row.campaignBudget) {
        const budgetDetails = extractBudgetDetails(row.campaignBudget);
        Object.assign(campaignDetails, budgetDetails);
      }

      if (row.campaign.biddingStrategy) {
        // Portfolio strategy is in use
        campaignDetails.campaignIsPortfolioBiddingStrategy = true;
        campaignDetails.portfolioStrategyResourceName =
            row.campaign.biddingStrategy;

        if (row.accessibleBiddingStrategy) {
          campaignDetails.campaignPortfolioBiddingStrategyName =
              row.accessibleBiddingStrategy.name || 'N/A';
          // Check strategy ownership
          if (row.accessibleBiddingStrategy.ownerCustomerId &&
              row.customer.id) {
            campaignDetails.portfolioOwnerId =
                row.accessibleBiddingStrategy.ownerCustomerId;
            if (row.accessibleBiddingStrategy.ownerCustomerId ==
                row.customer.id) {
              campaignDetails.isLocallyOwnedPortfolio = true;
              campaignDetails.campaignPortfolioIsManagerOwned = false;
            } else {
              campaignDetails.isLocallyOwnedPortfolio = false;
              campaignDetails.campaignPortfolioIsManagerOwned = true;
            }
          }
          // Extract targets from portfolio strategy
          if (row.accessibleBiddingStrategy.targetCpa) {
            campaignDetails.campaignBiddingTargetCpa = fromMicros(
                row.accessibleBiddingStrategy.targetCpa.targetCpaMicros);
          }
          if (row.accessibleBiddingStrategy.targetRoas) {
            campaignDetails.campaignBiddingTargetRoas =
                row.accessibleBiddingStrategy.targetRoas.targetRoas;
          }
          if (row.accessibleBiddingStrategy.maximizeConversions &&
              row.accessibleBiddingStrategy.maximizeConversions
                  .targetCpaMicros) {
            campaignDetails.campaignBiddingTargetCpa = fromMicros(
                row.accessibleBiddingStrategy.maximizeConversions
                    .targetCpaMicros);
          }
          if (row.accessibleBiddingStrategy.maximizeConversionValue &&
              row.accessibleBiddingStrategy.maximizeConversionValue
                  .targetRoas) {
            campaignDetails.campaignBiddingTargetRoas =
                row.accessibleBiddingStrategy.maximizeConversionValue
                    .targetRoas;
          }
        }
      } else {
        // Standard strategy is in use
        campaignDetails.campaignIsPortfolioBiddingStrategy = false;
        campaignDetails.campaignPortfolioIsManagerOwned = false;
        log(`tCPA from campaign object: ${row.campaign.targetCpa}`, 'debug');
        log(`tROAS from campaign object: ${row.campaign.targetRoas}`, 'debug');
        // Extract targets from campaign object for standard strategy
        if (row.campaign.targetCpa && row.campaign.targetCpa.targetCpaMicros) {
          campaignDetails.campaignBiddingTargetCpa =
              fromMicros(row.campaign.targetCpa.targetCpaMicros);
        }
        if (row.campaign.targetRoas && row.campaign.targetRoas.targetRoas) {
          campaignDetails.campaignBiddingTargetRoas =
              row.campaign.targetRoas.targetRoas;
        }
        if (row.campaign.maximizeConversions &&
            row.campaign.maximizeConversions.targetCpaMicros) {
          campaignDetails.campaignBiddingTargetCpa =
              fromMicros(row.campaign.maximizeConversions.targetCpaMicros);
        }
        if (row.campaign.maximizeConversionValue &&
            row.campaign.maximizeConversionValue.targetRoas) {
          campaignDetails.campaignBiddingTargetRoas =
              row.campaign.maximizeConversionValue.targetRoas;
        }
      }
    }
  } catch (e) {
    log(`Could not fetch core attributes with a single query. Error: ${
        e.message}`);
  }
  log(`Main GAQL query took ${
          ((new Date().getTime() - mainQueryStartTime) / 1000).toFixed(2)}s`,
      'verbose_timing');

  // --- Identify locally owned portfolio strategies for metric fetching ---
  const locallyOwnedPortfolioResourceNames = new Set();
  const managerOwnedPortfolioResourceNames = new Set();
  for (const details of campaignDataMap.values()) {
    if (details.campaignIsPortfolioBiddingStrategy) {
      if (details.isLocallyOwnedPortfolio) {
        locallyOwnedPortfolioResourceNames.add(
            details.portfolioStrategyResourceName);
      } else {
        managerOwnedPortfolioResourceNames.add(
            details.portfolioStrategyResourceName);
      }
    }
  }
  const locallyOwnedPortfolioIdsArray =
      Array.from(locallyOwnedPortfolioResourceNames);

  log(`Found ${
      locallyOwnedPortfolioIdsArray
          .length} locally owned portfolio strategies to query for metrics.`,
      'debug');
  if (managerOwnedPortfolioResourceNames.size > 0) {
    log(`Encountered ${
        managerOwnedPortfolioResourceNames
            .size} manager-owned portfolio strategies; ` +
        `historical tCPA/tROAS will be unavailable for campaigns using them.`);
  }

  // --- GAQL Query 1b: Get 30-Day Portfolio Metrics ---
  const thirtyDayPortfolioMetrics = fetchBiddingStrategyMetricsByDateRange(
      locallyOwnedPortfolioIdsArray, 'DURING LAST_30_DAYS');

  // --- GAQL Query 1c: Get Custom Range Portfolio Metrics ---
  const customRangePortfolioMetrics = fetchBiddingStrategyMetricsByDateRange(
      locallyOwnedPortfolioIdsArray,
      `BETWEEN '${formattedStartDate}' AND '${formattedEndDate}'`);

  // --- GAQL Query 2: Get 30-Day Metrics ---
  const thirtyDayMetrics = fetchCampaignMetricsByDateRange(
      campaignResourceNames, 'DURING LAST_30_DAYS');
  for (const [name, metrics] of thirtyDayMetrics) {
    const campaignDetails = campaignDataMap.get(name);
    campaignDetails.campaign30DaysCost = metrics.cost;
    campaignDetails.campaign30DaysConversionsValue = metrics.conversionsValue;
    campaignDetails.campaign30DaysConversions = metrics.conversions;
    campaignDetails.campaign30DaysRoas = metrics.roas;
    campaignDetails.campaign30DaysCpa = metrics.cpa;
    campaignDetails.campaign30DaysAvgCpm = metrics.averageCpm;
    campaignDetails.campaign30DaysAvgCpv = metrics.trueviewAverageCpv;
    campaignDetails.campaign30DaysUniqueUsers = metrics.uniqueUsers;
    if (!campaignDetails.campaignIsPortfolioBiddingStrategy) {
      // Standard strategy
      campaignDetails.campaign30DaysAvgTargetCpa = metrics.avgTargetCpa;
      campaignDetails.campaign30DaysAvgTargetRoas = metrics.avgTargetRoas;
    } else if (
        campaignDetails.isLocallyOwnedPortfolio &&
        thirtyDayPortfolioMetrics.has(
            campaignDetails.portfolioStrategyResourceName)) {
      // Locally-owned portfolio strategy with metrics found
      const portfolioMetrics = thirtyDayPortfolioMetrics.get(
          campaignDetails.portfolioStrategyResourceName);
      campaignDetails.campaign30DaysAvgTargetCpa = portfolioMetrics.avgTargetCpa;
      campaignDetails.campaign30DaysAvgTargetRoas =
          portfolioMetrics.avgTargetRoas;
    }
    campaignDetails.campaign30DaysSearchRankLostImpressionShare =
        metrics.searchRankLostImpressionShare;
    campaignDetails.campaign30DaysSearchBudgetLostImpressionShare =
        metrics.searchBudgetLostImpressionShare;
  }

  // --- GAQL Query 3: Get Custom Range Metrics ---
  const customRangeMetrics = fetchCampaignMetricsByDateRange(
      campaignResourceNames,
      `BETWEEN '${formattedStartDate}' AND '${formattedEndDate}'`);
  for (const [name, metrics] of customRangeMetrics) {
    const campaignDetails = campaignDataMap.get(name);
    campaignDetails.campaignCustomRangeCost = metrics.cost;
    campaignDetails.campaignCustomRangeConversionsValue =
        metrics.conversionsValue;
    campaignDetails.campaignCustomRangeConversions = metrics.conversions;
    campaignDetails.campaignCustomRangeRoas = metrics.roas;
    campaignDetails.campaignCustomRangeCpa = metrics.cpa;
    campaignDetails.campaignCustomRangeAvgCpm = metrics.averageCpm;
    campaignDetails.campaignCustomRangeAvgCpv = metrics.trueviewAverageCpv;
    campaignDetails.campaignCustomRangeUniqueUsers = metrics.uniqueUsers;
    if (!campaignDetails.campaignIsPortfolioBiddingStrategy) {
      // Standard strategy
      campaignDetails.campaignCustomRangeAvgTargetCpa = metrics.avgTargetCpa;
      campaignDetails.campaignCustomRangeAvgTargetRoas = metrics.avgTargetRoas;
    } else if (
        campaignDetails.isLocallyOwnedPortfolio &&
        customRangePortfolioMetrics.has(
            campaignDetails.portfolioStrategyResourceName)) {
      // Locally-owned portfolio strategy with metrics found
      const portfolioMetrics = customRangePortfolioMetrics.get(
          campaignDetails.portfolioStrategyResourceName);
      campaignDetails.campaignCustomRangeAvgTargetCpa =
          portfolioMetrics.avgTargetCpa;
      campaignDetails.campaignCustomRangeAvgTargetRoas =
          portfolioMetrics.avgTargetRoas;
    }
    campaignDetails.campaignCustomRangeSearchRankLostImpressionShare =
        metrics.searchRankLostImpressionShare;
    campaignDetails.campaignCustomRangeSearchBudgetLostImpressionShare =
        metrics.searchBudgetLostImpressionShare;
  }

  // --- GAQL Query 4: Get Yesterday Metrics ---
  const yesterdayMetrics =
      fetchCampaignMetricsByDateRange(
          campaignResourceNames, 'DURING YESTERDAY');
  for (const [name, metrics] of yesterdayMetrics) {
    const campaignDetails = campaignDataMap.get(name);
    campaignDetails.campaignStatsYesterdayCost = metrics.cost;
  }

  // --- GAQL Query 5: Get 7-Day Metrics ---
  const sevenDayMetrics =
      fetchCampaignMetricsByDateRange(
          campaignResourceNames, 'DURING LAST_7_DAYS');
  for (const [name, metrics] of sevenDayMetrics) {
    const campaignDetails = campaignDataMap.get(name);
    campaignDetails.campaignStats7DaysCost = metrics.cost;
  }


  log(`fetchAllCampaignData took ${((new Date().getTime() - startTime) / 1000).toFixed(2)}s`,
      'verbose_timing');
  return campaignDataMap;
}

/**
 * Decodes a Base64 encoded string into a regular string.
 * @param {string} base64Data The Base64 encoded string.
 * @return {string} The decoded string.
 */
function base64DecodeToString(base64Data) {
  const decoded = Utilities.base64Decode(base64Data);
  return Utilities.newBlob(decoded).getDataAsString();
}

/**
 * Extracts the OCID and recommendation type ID from a recommendation resource name.
 * The last segment of the resource name is a Base64 encoded string containing this info.
 * @param {string} recommendationId The full resource name of the recommendation.
 * @return {{ocid: string, recoTypeId: string}} An object with the OCID and
 *     recommendation type ID.
 */
function getInfoFromRecoId(recommendationId) {
  const recoResourceSegments = recommendationId.split('/');
  const recommendationIdDecoded = base64DecodeToString(
    recoResourceSegments[recoResourceSegments.length - 1]);
  const [ocid, recoTypeId, ...rest] = recommendationIdDecoded.split('-');
  return {ocid, recoTypeId};
}

/**
 * Fetches the OCID (Operating Customer ID) for a given Google Ads account ID.
 * The OCID is extracted from the optimization score URL.
 * @param {string} accountId The external customer ID of the account.
 * @return {!Object} An object containing the OCID.
 */
function getOcid(accountId) {
  const startTime = new Date().getTime();
  let ocid = 'N/A';

  try {
    log('Getting OCID for account:', 'debug');
    const ocidQuery = `
      SELECT
        customer.id,
        metrics.optimization_score_url
      FROM customer
      WHERE customer.id = '${accountId.replace(/-/g, '')}'`;


    const ocidResult = AdsApp.search(ocidQuery);

    if (ocidResult.hasNext()) {
      const row = ocidResult.next();
      if (row.metrics && row.metrics.optimizationScoreUrl) {
        const optiScoreUrl = row.metrics.optimizationScoreUrl;
        const match = optiScoreUrl.match(/[?&]ocid=([^&]*)/);
        ocid = match ? match[1] : 'N/A';
      }
    } else {
      log('Warning: No OCID found for ', accountId,
          '(unable to find OptiScore URL from which to extract the ID).');
    }
  } catch (err) {
    log(`ERROR - Unable to fetch OCID for ${accountId}: ${err.message}`);
  }

  log(`getOcid took ${((new Date().getTime() - startTime) / 1000).toFixed(2)}s`, 'verbose_timing');
  return {ocid: ocid};
}

/**
 * Orchestrates the entire processing for an account, turning recommendations
 * into the final report rows.
 *
 * @param {!AdsManagerApp.Account} account The Google Ads account object.
 * @param {number} counter The current account number being processed.
 * @param {number} totalAccounts The total number of accounts to process.
 * @return {!Array<!Object>} An array of objects, each representing a budget
 * recommendation and its associated campaign data.
 */
function processAccountRecommendations(account, counter, totalAccounts) {
  const accountStartTime = new Date().getTime();
  AdsManagerApp.select(account);
  const accountId = AdsApp.currentAccount().getCustomerId();
  const accountName = AdsApp.currentAccount().getName();
  const currencyCode = AdsApp.currentAccount().getCurrencyCode();
  log(progressBar(counter - 1, totalAccounts));
  log(`Processing account ${counter} of ${
      totalAccounts}: ${accountId} (${accountName})`);

  const supportedTypes =
      Object.keys(RECOMMENDATION_HANDLERS).map(type => `'${type}'`).join(',');
  const recommendationIterator =
      AdsApp.recommendations()
          .withCondition(`recommendation.type IN (${supportedTypes})`)
          .get();

  let processedRows = [];
  const recommendations = [];
  for (const recommendation of recommendationIterator) {
    recommendations.push(recommendation);
  }

  if (recommendations.length === 0) {
    log(`No supported recommendations found for this account. ` +
        `Account processing time: ${
            ((new Date().getTime() - accountStartTime) / 1000).toFixed(2)}s`,
        'timing');
    return processedRows;
  }

  // const ocid = getOcid(accountId).ocid;
  const {ocid} = getInfoFromRecoId(recommendations[0].getResourceName());

  log(`Found ${
      recommendations.length} recommendations to process.`);
  let recCounter = 1;

  // --- Phase 1: ID Gathering ---
  const idGatheringStartTime = new Date().getTime();
  const recommendationDetailsMap =
      fetchAllRecommendationDetails(recommendations);
  const allCampaignResourceNames = new Set();
  const allBudgetResourceNames = new Set();

  for (const recommendation of recommendations) {
    const handler = RECOMMENDATION_HANDLERS[recommendation.getType()];
    if (handler) {
      const recommendationDetails =
          recommendationDetailsMap.get(recommendation.getResourceName());
      if (recommendationDetails) {
        const recommendationTargets = handler.getRecommendationTargets(
            recommendation, recommendationDetails);
        for (const target of recommendationTargets) {
          if (target && target.campaignResourceName) {
            allCampaignResourceNames.add(target.campaignResourceName);
          } else {
            log(`WARNING: Found recommendation with an invalid or missing ` +
                `campaign target. Skipping target for recommendation: ${
                    recommendation.getResourceName()}`);
          }
        }
        const requiredBudgets =
            handler.getRequiredBudgetResourceNames(recommendationDetails);
        for (const budgetName of requiredBudgets) {
          allBudgetResourceNames.add(budgetName);
        }
      }
    }
  }
  log(`ID gathering took ${
          ((new Date().getTime() - idGatheringStartTime) / 1000).toFixed(2)}s`,
      'timing');


  // --- Phase 2: Bulk Data Fetching ---
  const bulkFetchStartTime = new Date().getTime();
  const campaignDataMap =
      fetchAllCampaignData(Array.from(allCampaignResourceNames));
  const budgetDataMap =
      fetchAllBudgetDetails(Array.from(allBudgetResourceNames));
  const budgetToCampaignsMap =
      fetchCampaignsByBudgets(Array.from(allBudgetResourceNames));
  log(`Bulk data fetching took ${
          ((new Date().getTime() - bulkFetchStartTime) / 1000).toFixed(2)}s`,
      'timing');


  // --- Phase 3: Data Processing and Enrichment ---
  const processingStartTime = new Date().getTime();
  for (const recommendation of recommendations) {
    const recStartTime = new Date().getTime();
    const type = recommendation.getType();
    log(progressBar(recCounter - 1, recommendations.length));
    log(`Processing recommendation ${type} (${recCounter} of ${
        recommendations.length}): ${recommendation.getResourceName()}`);

    const handler = RECOMMENDATION_HANDLERS[type];
    if (!handler) {
      log(`WARNING: No handler found for recommendation type '${
          type}'. Skipping.`);
      recCounter++;
      continue;
    }
    const recommendationDetails =
        recommendationDetailsMap.get(recommendation.getResourceName());
    if (!recommendationDetails) {
      log(`WARNING: Could not find pre-fetched details for ` +
          `recommendation. Skipping: ${recommendation.getResourceName()}`);
      recCounter++;
      continue;
    }

    // --- Step 1: Get recommendation-level data ONCE ---
    const getRecommendationTargetsStartTime = new Date().getTime();
    const recommendationTargets =
        handler.getRecommendationTargets(recommendation, recommendationDetails);
    log(`handler.getRecommendationTargets() took ${
            ((new Date().getTime() -
              getRecommendationTargetsStartTime) /
             1000)
                .toFixed(2)}s`,
        'verbose_timing');

    if (recommendationTargets.length === 0) {
      log(`WARNING: No recommendation targets found for recommendation. Skipping.`);
      recCounter++;
      continue;
    }

    // --- Step 2: Process each recommendation target entity (e.g., campaign)
    // ---
    let targetCounter = 1;
    for (const recommendationTarget of recommendationTargets) {
      const targetStartTime = new Date().getTime();
      log(`Processing recommendation target ${targetCounter} of ${
          recommendationTargets.length} for recommendation.`, 'debug');
      // Create the base report row
      let rowData =
          createReportRow(accountId, accountName, recommendation, currencyCode);

      // Add the type-specific details from the handler
      rowData = handler.addSpecificDetails(
          rowData, recommendationTarget, recommendation, recommendationDetails,
          budgetDataMap, budgetToCampaignsMap);

      // Add the common recommendation impact stats
      const impactMetrics =
          handler.getImpactStats(recommendation, recommendationDetails);
      const recommendationImpactStats =
          extractRecommendationImpactStats(impactMetrics);
      Object.assign(rowData, recommendationImpactStats);

      // Fetch and add all campaign-specific data
      const campaignDetails =
          campaignDataMap.get(recommendationTarget.campaignResourceName);
      if (campaignDetails) {
        Object.assign(rowData, campaignDetails);
        rowData.campaignUrl = `https://ads.google.com/aw/overview?campaignId=${
            campaignDetails.campaignId}&ocid=${ocid}`;
      }


      // Perform final calculations using the fully populated rowData
      const calculationsStartTime = new Date().getTime();
      rowData.campaignCalculatedCustomRangePotentialConversionValue =
          calculatePotentialConversionValue(
              rowData.campaignCustomRangeConversionsValue,
              rowData.campaignCustomRangeConversions,
              rowData.recommendationPotentialConversions);
      rowData.campaignCalculated30DaysPotentialConversionValue =
          calculatePotentialConversionValue(
              rowData.campaign30DaysConversionsValue,
              rowData.campaign30DaysConversions,
              rowData.recommendationPotentialConversions);

      if (rowData.recommendationCurrentBudgetAmount > 0) {
        rowData.campaignPercentageOfBudgetUsedYesterday =
            rowData.campaignStatsYesterdayCost /
            rowData.recommendationCurrentBudgetAmount;
      } else {
        rowData.campaignPercentageOfBudgetUsedYesterday = 0;
      }

      rowData.weeklyCostIncrease = calculateDelta(
          rowData.recommendationPotentialCost, rowData.recommendationBaseCost);
      rowData.newWeeklyConversions = calculateDelta(
          rowData.recommendationPotentialConversions,
          rowData.recommendationBaseConversions);
      rowData.dailyBudget30Days =
          typeof rowData.campaignBudgetAmount === 'number' ?
          rowData.campaignBudgetAmount * 30 :
          null;
      rowData.dailyBudgetDelta =
          calculateDelta(rowData.dailyBudget30Days, rowData.campaign30DaysCost);
      rowData.targetRoasDelta = calculateDelta(
          rowData.campaign30DaysRoas, rowData.campaignBiddingTargetRoas);
      rowData.targetCpaDelta = calculateDelta(
          rowData.campaign30DaysCpa, rowData.campaignBiddingTargetCpa);

      rowData.recommendationCpaDelta = calculateDelta(
        rowData.recommendationBaseCpa, rowData.recommendationPotentialCpa);
      recommendationRoasDelta = calculateDelta(
        rowData.recommendationBaseRoas, rowData.recommendationPotentialRoas);

      log(`Final calculations took ${
              ((new Date().getTime() - calculationsStartTime) / 1000)
                  .toFixed(2)}s`,
          'verbose_timing');

      processedRows.push(rowData);
      log(`Row for campaign ${
          campaignDetails ? campaignDetails.campaignId :
                            'N/A'} is complete.`,
          'debug');
      log(`Recommendation target processing took ${
              ((new Date().getTime() - targetStartTime) / 1000).toFixed(2)}s`,
          'verbose_timing');
      targetCounter++;
    }
    log(`Recommendation processing took ${
            ((new Date().getTime() - recStartTime) / 1000).toFixed(2)}s`,
        'verbose_timing');
    recCounter++;
  }
  log(`Data processing and enrichment took ${
          ((new Date().getTime() - processingStartTime) / 1000).toFixed(2)}s`,
      'timing');

  log(`Total account processing time: ${
          ((new Date().getTime() - accountStartTime) / 1000).toFixed(2)}s`,
      'timing');
  return processedRows;
}


/**
 * Writes a list of objects to a Google Sheet, including header extraction and
 * data transformation.
 *
 * @param {!Spreadsheet} spreadsheet The Google Spreadsheet object.
 * @param {string} sheetName The name of the worksheet to write to.
 * @param {!Array<!Object>} reportData An array of objects with uniform
 * properties to write.
 * @return {void}
 */
function writeDataToSheet(spreadsheet, sheetName, reportData) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    log(`Sheet "${sheetName}" created.`);
  }

  // Check for the default "Sheet1" and delete it if other sheets exist.
  const defaultSheet = spreadsheet.getSheetByName("Sheet1");
  if (defaultSheet && spreadsheet.getSheets().length > 1) {
    try {
      spreadsheet.deleteSheet(defaultSheet);
      log('Deleted default "Sheet1".');
    } catch (e) {
      log('Could not delete "Sheet1" (it might be the only sheet remaining).');
    }
  }

  if (reportData && reportData.length > 0) {
    // Clear the entire sheet before writing new data
    sheet.clear();

    const headers =
        Object.keys(reportData[0]);  // Extract headers from the first object
    const dataRows = [];

    dataRows.push(headers);  // Add headers as the first row

    // Transform object array into 2D data array
    for (let i = 0; i < reportData.length; i++) {
      const row = [];
      for (let j = 0; j < headers.length; j++) {
        row.push(reportData[i][headers[j]]);
      }
      dataRows.push(row);
    }

    const numRows = dataRows.length;
    const numCols = headers.length;
    const range = sheet.getRange(1, 1, numRows, numCols);

    range.setValues(dataRows);

    log(`Data written to sheet "${sheetName}", range: ${
        range.getA1Notation()}`);

    // --- Apply Column Formatting ---
    log('Applying column formatting...');
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      if (REPORT_SCHEMA_TEMPLATE[header] &&
          REPORT_SCHEMA_TEMPLATE[header].format) {
        const format = REPORT_SCHEMA_TEMPLATE[header].format;
        // Apply format to the entire column, starting from the second row
        const columnRange = sheet.getRange(2, i + 1, sheet.getMaxRows() - 1);
        try {
          columnRange.setNumberFormat(format);
          if (header === 'campaign30DaysAvgTargetCpa') {
             log(`*** DEBUG: Explicitly formatted 'campaign30DaysAvgTargetCpa' (Col ${i+1}) with format: '${format}' ***`, 'debug');
          }
          log(`  - Column '${header}' formatted as '${format}'.`, 'debug');
        } catch (e) {
          log(`  - WARNING: Failed to format column '${header}' with format '${format}'. Error: ${e.message}`);
        }
      }
    }
    log('Column formatting applied.');

  } else {
    log(`No data rows to write to sheet "${sheetName}".`);
  }
}


/**
 * Calculates and writes aggregated summary data to a new "Summary" sheet.
 *
 * @param {!Spreadsheet} spreadsheet The Google Spreadsheet object.
 * @param {!Array<!Object>} reportData The complete, final dataset from all
 * processed accounts.
 */
function calculateAndWriteSummaryData(spreadsheet, reportData) {
  const startTime = new Date().getTime();
  if (!reportData || reportData.length === 0) {
    log('No data available to generate summary. Skipping.');
    return;
  }

  // Metric 1: Potential Daily Budget Increase (Campaign Budget)
  const campaignBudgetRows =
      reportData.filter(row => row.recommendationType === 'CAMPAIGN_BUDGET');
  const sumOfCurrentBudget = campaignBudgetRows.reduce(
      (sum, row) => sum + (row.recommendationCurrentBudgetAmount || 0), 0);
  const sumOfNewBudget = campaignBudgetRows.reduce(
      (sum, row) => sum + (row.recommendationNewBudgetAmount || 0), 0);
  const potentialIncrease = sumOfNewBudget - sumOfCurrentBudget;

  // Metric 2: Recommended Moveable Budget
  const moveBudgetRows =
      reportData.filter(row => row.recommendationType === 'MOVE_UNUSED_BUDGET');
  const moveableBudget = moveBudgetRows.reduce(
      (sum, row) => sum + (row.moveBudgetAmount || 0), 0);

  // Metric 3: Forecasted Conversion Uplift (Move Unused Budget)
  const sumOfBaseConversions = moveBudgetRows.reduce(
      (sum, row) => sum + (row.recommendationBaseConversions || 0), 0);
  const sumOfPotentialConversions = moveBudgetRows.reduce(
      (sum, row) => sum + (row.recommendationPotentialConversions || 0), 0);
  const conversionUplift = sumOfPotentialConversions - sumOfBaseConversions;

  const summaryData = [
    {
      'Potential Daily Budget Increase (Campaign Budget)': potentialIncrease,
      'Recommended Moveable Budget': moveableBudget,
      'Forecasted Conversion Uplift (Move Unused Budget)': conversionUplift
    }
  ];

  writeDataToSheet(spreadsheet, 'Summary', summaryData);
  log(`calculateAndWriteSummaryData took ${
          ((new Date().getTime() - startTime) / 1000).toFixed(2)}s`,
      'timing');
}


/**
 * Loads a configuration object from the spreadsheet.
 *
 * @param {!Spreadsheet} spreadsheet The Google Spreadsheet object.
 * @return {!Object} A configuration object.
 */
function readConfig(spreadsheet) {
  const mccAccountId =
      spreadsheet.getRangeByName('mccAccountId').getValues().flat()[0];
  const excludedAccounts =
      spreadsheet.getRangeByName('excludedAccounts').getValues().flat();
  const excludedCampaigns =
      spreadsheet.getRangeByName('excludedCampaigns').getValues().flat();

  return {
    'mccAccountId': mccAccountId,
    'excludedAccounts': excludedAccounts.filter(item => item),
    'excludedCampaigns': excludedCampaigns.filter(item => item),
  };
}

/**
 * Detects the effective email of the user running the script by creating
 * and inspecting a temporary file. This workaround is necessary because
 * Session.getEffectiveUser() is not available in all Ads Script contexts.
 * @return {string} The email address of the script runner, or null if failed.
 */
function detectScriptRunnerEmail() {
  try {
    log("Attempting to detect script runner identity via temp file...", 'debug');
    const tempSs = SpreadsheetApp.create("Temp_Auth_Check_" + new Date().getTime());
    const tempId = tempSs.getId();
    const file = DriveApp.getFileById(tempId);
    const email = file.getOwner().getEmail();
    file.setTrashed(true);
    log(`Detected user: ${email}`, 'debug');
    return email;
  } catch (e) {
    log(`Failed to detect user email: ${e.message}`);
    return null;
  }
}

/**
 * Gets the existing spreadsheet from properties or creates a new one.
 * @param {string} reportName The desired name of the report/file.
 * @return {!Object} An object containing the spreadsheet and a flag indicating if it's new.
 */
function createOrGetSpreadsheet(reportName) {
  const props = PropertiesService.getScriptProperties();
  const storedUrl = props.getProperty('spreadsheet_url');

  if (storedUrl) {
    try {
      const ss = SpreadsheetApp.openByUrl(storedUrl);
      // Ensure the name matches the current account name
      if (ss.getName() !== reportName) {
        ss.rename(reportName);
      }
      log(`Found existing spreadsheet linked to script: ${storedUrl}`);
      return { spreadsheet: ss, isNew: false };
    } catch (e) {
      log(`Error accessing stored spreadsheet: ${storedUrl}. Error: ${e.message}`);

      if (!CONFIG.force_new_sheet_on_access_error) {
        // Attempt to identify the current user to give a helpful error message
        let currentUser = "the current script runner";
        const detectedEmail = detectScriptRunnerEmail();
        if (detectedEmail) {
            currentUser = detectedEmail;
        }

        const msg = `\n--------------------------------------------------------------------------------\n` +
                    `PERMISSION ERROR\n` +
                    `--------------------------------------------------------------------------------\n` +
                    `The script is trying to access the existing report, but access was denied.\n` +
                    `Spreadsheet URL: ${storedUrl}\n` +
                    `Script Runner:   ${currentUser}\n\n` +
                    `SOLUTION:\n` +
                    `1. Share the spreadsheet above with ${currentUser} as an Editor.\n` +
                    `2. OR, set CONFIG.force_new_sheet_on_access_error = true in the script to force a new file (WARNING: breaks Looker Studio links).\n` +
                    `--------------------------------------------------------------------------------`;
        throw new Error(msg);
      }

      log(`Proceeding to create a new spreadsheet because CONFIG.force_new_sheet_on_access_error is true.`);
    }
  }

  const ss = SpreadsheetApp.create(reportName);
  props.setProperty('spreadsheet_url', ss.getUrl());
  log(`Created new spreadsheet: ${ss.getUrl()}`);
  return { spreadsheet: ss, isNew: true };
}

/**
 * Sends the setup email with the magic link.
 * @param {string} sheetUrl The URL of the Google Sheet.
 * @param {string} reportName The name of the report.
 * @param {string} accountName The name of the Google Ads account.
 * @param {string} recipientEmail The email address of the recipient.
 */
function sendSetupEmail(sheetUrl, reportName, accountName, recipientEmail) {
  if (!recipientEmail) {
    log("No recipient email detected. Skipping email.");
    return;
  }

  const ss = SpreadsheetApp.openByUrl(sheetUrl);
  const spreadsheetId = ss.getId();

  // Logic adapted from provided snippet to handle specific aliases
  const baseUrl = "https://lookerstudio.google.com/reporting/create";
  const params = [];

  // 1. Add Report Config
  params.push(`c.reportId=${CONFIG.LOOKER_TEMPLATE_ID}`);
  params.push(`r.reportName=${encodeURIComponent(reportName)}`);

  // 2. Define datasources (Sheet Name -> DS Alias)
  const sheetDefinitions = {
    'Data': 'data_sheet',
    'Summary': 'summary_sheet'
  };

  // 3. Build params for each sheet
  Object.keys(sheetDefinitions).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const alias = sheetDefinitions[sheetName];
      const prefix = `ds.${alias}`;
      params.push(`${prefix}.connector=googleSheets`);
      params.push(`${prefix}.spreadsheetId=${spreadsheetId}`);
      params.push(`${prefix}.worksheetId=${sheet.getSheetId()}`);
    }
  });

  const lookerStudioLink = `${baseUrl}?${params.join('&')}`;

  const subject = `${accountName} - Your Budget Report Data is Ready for Setup`;

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; color: #333;">
      <h2 style="color: #4285f4;">Your Budget Report Data is Ready</h2>
      <p>Hi there,</p>
      <p>Your Google Ads budget data for account <strong>${accountName}</strong> has been successfully exported to Google Sheets.</p>

      <p>To finalize your report, click the button below. It will open Looker Studio with your data already connected and named correctly.</p>

      <p>Only do this <b>once</b>, otherwise you will end up with duplicate reports.</p>

      <br>
      <a href="${lookerStudioLink}" style="background-color: #4285f4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">
        CREATE REPORT IN LOOKER STUDIO
      </a>

      <br>
      <br>
      <p>Once your Looker report is set up, the contents will update automatically each time the Google Ads script runs.</p>
      <br>

      <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">

      <p style="font-size: 12px; color: #666;">
        <strong>Reference:</strong><br>
        <a href="${sheetUrl}" style="color: #666;">Open Google Sheet Source</a>
      </p>
       <p style="font-size: 12px; color: #999;">
        <em>(You can schedule this script to run daily to keep the data fresh).</em>
      </p>
    </div>
  `;

  MailApp.sendEmail({
    to: recipientEmail,
    subject: subject,
    htmlBody: htmlBody
  });

  log(`Email sent to ${recipientEmail}`);
}


/**
 * Creates a progress bar string.
 * @param {number} current The current item number.
 * @param {number} total The total number of items.
 * @param {number=} width The width of the progress bar in characters.
 * @return {string} The progress bar string.
 */
function progressBar(current, total, width = 54) {
  if (total === 0) {
    return '='.repeat(width);
  }
  const percentage = current / total;
  const progress = Math.round(width * percentage);
  const remaining = width - progress;
  return '='.repeat(progress) + '.'.repeat(remaining);
}

/**
 * Builds a map of internal CID -> Account object for all accessible accounts.
 * @param {!AdsManagerApp.Account} topLevelAccount The account running the script.
 * @return {!Map<number, !AdsManagerApp.Account>} Map of internal CID to account
 * object.
 */
function buildInternalIdToAccountMap(topLevelAccount) {
  log('Building internal ID to Account object map...', 'timing');
  const startTime = new Date().getTime();
  const accountIterator = AdsManagerApp.accounts().get();
  const internalIdToAccountMap = new Map();
  const accounts = [];
  while (accountIterator.hasNext()) {
    accounts.push(accountIterator.next());
  }
  accounts.push(topLevelAccount);  // also need to map top level
  log(`Found ${accounts.length} accounts to map.`, 'debug');

  for (const account of accounts) {
    try {
      log(`Mapping account: ${account.getCustomerId()} (${account.getName()})`,
          'debug');
      AdsManagerApp.select(account);
      const result = AdsApp.search('SELECT customer.id FROM customer LIMIT 1');
      if (result.hasNext()) {
        const internalId = result.next().customer.id;
        log(`  -> Internal ID: ${internalId}`, 'debug');
        internalIdToAccountMap.set(internalId, account);
      } else {
        log(`  -> No internal ID found for ${account.getCustomerId()}.`,
            'debug');
      }
    } catch (e) {
      log(`Could not get internal ID for account ${account.getCustomerId()}: ${
          e}`);
    }
  }
  log(`Building internal ID map took ${
          ((new Date().getTime() - startTime) / 1000).toFixed(2)}s`,
      'timing');
  log('Finished building internalIdToAccountMap:', 'debug');
  log(internalIdToAccountMap, 'debug');
  return internalIdToAccountMap;
}

/**
 * Fetches metrics for manager-owned portfolio strategies by switching account
 * context, and merges these metrics into the recommendation rows.
 * @param {!Array<!Object>} recommendations The recommendation rows processed so
 * far.
 * @param {!Map<number, !AdsManagerApp.Account>} internalIdToAccountMap Map of
 * internal CID to Account object.
 * @param {!AdsManagerApp.Account} topLevelAccount The account running the script.
 * @param {string} formattedStartDate Custom start date yyyy-MM-dd.
 * @param {string} formattedEndDate Custom end date yyyy-MM-dd.
 * @return {!Array<!Object>} The enriched recommendation rows.
 */
function enrichWithMccPortfolioMetrics(
    recommendations, internalIdToAccountMap, topLevelAccount,
    formattedStartDate, formattedEndDate) {
  const mccEnrichStartTime = new Date().getTime();
  // Group manager-owned strategies by owner
  // Map<ownerId, Set<strategyResourceName>>
  const ownerToStrategies = new Map();
  for (const row of recommendations) {
    if (row.campaignPortfolioIsManagerOwned === true && row.portfolioOwnerId &&
        row.portfolioStrategyResourceName) {
      if (!ownerToStrategies.has(row.portfolioOwnerId)) {
        ownerToStrategies.set(row.portfolioOwnerId, new Set());
      }
      ownerToStrategies.get(row.portfolioOwnerId)
          .add(row.portfolioStrategyResourceName);
    }
  }

  if (ownerToStrategies.size === 0) {
    log('No manager-owned portfolio strategies found that require ' +
        'metric enrichment.',
        'debug');
    return recommendations;
  }

  log(`Found ${
      ownerToStrategies
          .size} manager account(s) owning portfolio strategies used by ` +
      `client campaigns.`);

  // Map<strategyResourceName, {thirtyDay: Object, customRange: Object}>
  const metricsCache = new Map();

  for (const [ownerId, strategySet] of ownerToStrategies) {
    const ownerAccount = internalIdToAccountMap.get(ownerId);
    if (!ownerAccount) {
      log(`Warning: Cannot fetch metrics for portfolio strategies owned by ${
          ownerId}: manager account not found or accessible.`);
      continue;
    }

    try {
      log(`Switching to context of manager ${ownerId} (${
          ownerAccount.getName()}) to fetch portfolio metrics.`);
      AdsManagerApp.select(ownerAccount);
      const strategyList = Array.from(strategySet);
      const thirtyDayMetrics = fetchBiddingStrategyMetricsByDateRange(
          strategyList, 'DURING LAST_30_DAYS');
      const customRangeMetrics = fetchBiddingStrategyMetricsByDateRange(
          strategyList,
          `BETWEEN '${formattedStartDate}' AND '${formattedEndDate}'`);

      for (const strategyRn of strategyList) {
        metricsCache.set(strategyRn, {
          thirtyDay: thirtyDayMetrics.get(strategyRn),
          customRange: customRangeMetrics.get(strategyRn)
        });
      }
    } catch (e) {
      log(`ERROR: Failed during context switch or metric fetch for manager ${
          ownerId}: ${e}`);
    }
  }

  // Restore context to top-level account
  AdsManagerApp.select(topLevelAccount);

  // Merge metrics back into recommendations
  for (const row of recommendations) {
    if (row.campaignPortfolioIsManagerOwned === true &&
        metricsCache.has(row.portfolioStrategyResourceName)) {
      const cachedMetric = metricsCache.get(row.portfolioStrategyResourceName);
      if (cachedMetric.thirtyDay) {
        row.campaign30DaysAvgTargetCpa = cachedMetric.thirtyDay.avgTargetCpa;
        row.campaign30DaysAvgTargetRoas = cachedMetric.thirtyDay.avgTargetRoas;
      }
      if (cachedMetric.customRange) {
        row.campaignCustomRangeAvgTargetCpa =
            cachedMetric.customRange.avgTargetCpa;
        row.campaignCustomRangeAvgTargetRoas =
            cachedMetric.customRange.avgTargetRoas;
      }
    }
  }

  log(`MCC portfolio metric enrichment took ${
          ((new Date().getTime() - mccEnrichStartTime) / 1000).toFixed(2)}s`,
      'timing');
  return recommendations;
}


/**
 * Entry point to execute the script.
 */
function main() {
  const scriptStartTime = new Date().getTime();

  // 1. Get Account Details
  const topLevelAccount = AdsApp.currentAccount();  // The MCC running the script
  const accountName = topLevelAccount.getName();

  // Construct the dynamic Report Name: "[MCC Account name] Budget Recommendation Report"
  const DYNAMIC_REPORT_NAME = `[${accountName}] - Budget Recommendation Report`;

  log(`Processing Account: ${accountName}`);

  // 2. Create or Get Spreadsheet
  const sheetResult = createOrGetSpreadsheet(DYNAMIC_REPORT_NAME);
  const spreadsheet = sheetResult.spreadsheet;
  const isNewSheet = sheetResult.isNew;
  const spreadsheetUrl = spreadsheet.getUrl();

  const mccAccountTimeZone = topLevelAccount.getTimeZone();

  // Calculate dates early
  const today = new Date();
  const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  const startDate =
      new Date(today.getTime() - CUSTOM_RANGE_DAYS_AGO_START * MILLIS_PER_DAY);
  const endDate =
      new Date(today.getTime() - CUSTOM_RANGE_DAYS_AGO_END * MILLIS_PER_DAY);
  const formattedStartDate =
      Utilities.formatDate(startDate, mccAccountTimeZone, 'yyyy-MM-dd');
  const formattedEndDate =
      Utilities.formatDate(endDate, mccAccountTimeZone, 'yyyy-MM-dd');

  const accountSelector = AdsManagerApp.accounts();
  const accountIterator = accountSelector.get();
  const totalAccounts = accountIterator.totalNumEntities();
  log(`Found ${totalAccounts} sub-accounts to process.`);

  // Build account map
  const internalIdToAccountMap = buildInternalIdToAccountMap(topLevelAccount);

  let recommendations = [];
  let counter = 1;
  while (accountIterator.hasNext()) {
    const account = accountIterator.next();
    try {
      recommendations.push.apply(
          recommendations,
          processAccountRecommendations(account, counter, totalAccounts));
    } catch (e) {
      log(`Could not process account ${account.getCustomerId()}: ${e.message}`);
    }
    counter++;
  }
  log(recommendations, 'debug');

  // New phase: Enrich with MCC portfolio metrics
  recommendations = enrichWithMccPortfolioMetrics(
      recommendations, internalIdToAccountMap, topLevelAccount,
      formattedStartDate, formattedEndDate);

  writeDataToSheet(spreadsheet, 'Data', recommendations);
  calculateAndWriteSummaryData(spreadsheet, recommendations);

  // 4. Send Setup Email
  if (isNewSheet || CONFIG.force_setup_email) {
    let effectiveUserEmail = null;

    // If the sheet is brand new, the runner IS the owner.
    // If the sheet already exists but we are forcing an email, we verify identity via temp file
    // because the sheet owner might be someone else (e.g. the original creator).
    if (isNewSheet) {
        try {
            effectiveUserEmail = DriveApp.getFileById(spreadsheet.getId()).getOwner().getEmail();
        } catch(e) {
            log("Could not get owner of new sheet: " + e.message);
        }
    } else {
        effectiveUserEmail = detectScriptRunnerEmail();
    }

    sendSetupEmail(spreadsheetUrl, DYNAMIC_REPORT_NAME, accountName, effectiveUserEmail);
  } else {
    log("Spreadsheet already exists, skipping setup email.");
  }

  const scriptEndTime = new Date().getTime();
  log(`Total script execution time: ${
          ((scriptEndTime - scriptStartTime) / 1000).toFixed(2)}s`,
      'timing');
}