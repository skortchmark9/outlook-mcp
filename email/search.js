/**
 * Improved search emails functionality
 */
const config = require('../config');
const { callGraphAPI, callGraphAPIPaginated } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');
const { resolveFolderPath } = require('./folder-utils');

/**
 * Search emails handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchEmails(args) {
  const folder = args.folder || "inbox";
  const requestedCount = args.count || 10;
  const query = args.query || '';
  const from = args.from || '';
  const to = args.to || '';
  const subject = args.subject || '';
  const hasAttachments = args.hasAttachments;
  const unreadOnly = args.unreadOnly;
  
  try {
    // Get access token
    const accessToken = await ensureAuthenticated();
    
    // Resolve the folder path
    const endpoint = await resolveFolderPath(accessToken, folder);
    console.error(`Using endpoint: ${endpoint} for folder: ${folder}`);
    
    // Execute progressive search with pagination
    const response = await progressiveSearch(
      endpoint, 
      accessToken, 
      { query, from, to, subject },
      { hasAttachments, unreadOnly },
      requestedCount
    );
    
    return formatSearchResults(response);
  } catch (error) {
    // Handle authentication errors
    if (error.message === 'Authentication required') {
      return {
        content: [{ 
          type: "text", 
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }
    
    // General error response
    return {
      content: [{ 
        type: "text", 
        text: `Error searching emails: ${error.message}`
      }]
    };
  }
}

/**
 * Execute a search with progressively simpler fallback strategies
 * @param {string} endpoint - API endpoint
 * @param {string} accessToken - Access token
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 * @param {number} maxCount - Maximum number of results to retrieve
 * @returns {Promise<object>} - Search results
 */
async function progressiveSearch(endpoint, accessToken, searchTerms, filterTerms, maxCount) {
  // Track search strategies attempted
  const searchAttempts = [];

  // Did the caller actually ask to filter? If so, a zero-match result must be
  // reported as zero — NOT masked by dumping recent emails (which get labelled
  // "matching your search criteria" and read as false positives). The recent-
  // emails fallback is only honest when there were no search terms at all.
  const hasSearchTerms = !!(
    searchTerms.query || searchTerms.from || searchTerms.to || searchTerms.subject
  );

  // 1. Try combined search (most specific)
  try {
    const params = buildSearchParams(searchTerms, filterTerms, Math.min(50, maxCount));
    console.error("Attempting combined search with params:", params);
    searchAttempts.push("combined-search");
    
    const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, params, maxCount);
    if (response.value && response.value.length > 0) {
      console.error(`Combined search successful: found ${response.value.length} results`);
      return response;
    }
  } catch (error) {
    console.error(`Combined search failed: ${error.message}`);
  }
  
  // 2. Try each search term individually, starting with most specific
  const searchPriority = ['subject', 'from', 'to', 'query'];
  
  for (const term of searchPriority) {
    if (searchTerms[term]) {
      try {
        console.error(`Attempting search with only ${term}: "${searchTerms[term]}"`);
        searchAttempts.push(`single-term-${term}`);
        
        // For single term search, only use $search with that term
        // Graph API does not support $orderby or $filter with $search
        const simplifiedParams = {
          $top: Math.min(50, maxCount),
          $select: config.EMAIL_SELECT_FIELDS
        };

        // Build KQL terms for search. 'query' is free text (no field prefix).
        const kqlParts = [buildKqlTerm(term === 'query' ? null : term, searchTerms[term])];

        // Add boolean filters as KQL (can't use $filter with $search)
        addBooleanFiltersAsKQL(kqlParts, filterTerms);

        simplifiedParams.$search = `"${kqlParts.join(' AND ')}"`;
        
        const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, simplifiedParams, maxCount);
        if (response.value && response.value.length > 0) {
          console.error(`Search with ${term} successful: found ${response.value.length} results`);
          return response;
        }
      } catch (error) {
        console.error(`Search with ${term} failed: ${error.message}`);
      }
    }
  }
  
  // 3. Try with only boolean filters
  if (filterTerms.hasAttachments === true || filterTerms.unreadOnly === true) {
    try {
      console.error("Attempting search with only boolean filters");
      searchAttempts.push("boolean-filters-only");
      
      const filterOnlyParams = {
        $top: Math.min(50, maxCount),
        $select: config.EMAIL_SELECT_FIELDS,
        $orderby: 'receivedDateTime desc'
      };
      
      // Add the boolean filters
      addBooleanFilters(filterOnlyParams, filterTerms);
      
      const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, filterOnlyParams, maxCount);
      console.error(`Boolean filter search found ${response.value?.length || 0} results`);
      return response;
    } catch (error) {
      console.error(`Boolean filter search failed: ${error.message}`);
    }
  }
  
  // 4a. Caller supplied search terms but nothing matched. Report an honest zero
  // instead of falling back to unrelated recent emails (which would be presented
  // as matches). This is the fix for the silent false-positive fallback.
  if (hasSearchTerms) {
    console.error("Search terms provided but no matches found; returning empty (no recent-emails fallback)");
    return {
      value: [],
      '@odata.count': 0,
      _searchInfo: {
        attemptsCount: searchAttempts.length,
        strategies: searchAttempts,
        originalTerms: searchTerms,
        filterTerms: filterTerms,
        noFallback: true,
      },
    };
  }

  // 4b. Final fallback (no search terms): just get recent emails with pagination
  console.error("No search terms, returning recent emails");
  searchAttempts.push("recent-emails");

  const basicParams = {
    $top: Math.min(50, maxCount),
    $select: config.EMAIL_SELECT_FIELDS,
    $orderby: 'receivedDateTime desc'
  };
  
  const response = await callGraphAPIPaginated(accessToken, 'GET', endpoint, basicParams, maxCount);
  console.error(`Fallback to recent emails found ${response.value?.length || 0} results`);
  
  // Add a note to the response about the search attempts
  response._searchInfo = {
    attemptsCount: searchAttempts.length,
    strategies: searchAttempts,
    originalTerms: searchTerms,
    filterTerms: filterTerms
  };
  
  return response;
}

/**
 * Build a single KQL term for use inside a Graph $search expression.
 *
 * Two hard-won rules, both verified against the live Graph API:
 *  1. Never wrap a value in double quotes. `subject:"foo bar"` inside the
 *     already-quoted $search string produces nested quotes and Graph rejects it
 *     with `400 Syntax error: character '"' is not valid`.
 *  2. Multi-word values must be parenthesized: `subject:(foo bar)`. Without the
 *     parens, `subject:foo bar` scopes only "foo" to the subject and lets "bar"
 *     match ANY field as free text, so unrelated emails come back as "matches".
 *
 * @param {string|null} field - KQL property (subject/from/to) or null for free text
 * @param {string} value - user-supplied value
 * @returns {string} - a KQL term, e.g. `subject:(RateMate Analysis)`
 */
function buildKqlTerm(field, value) {
  const v = String(value).trim();
  const scoped = /\s/.test(v) ? `(${v})` : v;
  return field ? `${field}:${scoped}` : scoped;
}

/**
 * Build search parameters from search terms and filter terms
 * @param {object} searchTerms - Search terms (query, from, to, subject)
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 * @param {number} count - Maximum number of results
 * @returns {object} - Query parameters
 */
function buildSearchParams(searchTerms, filterTerms, count) {
  const params = {
    $top: count,
    $select: config.EMAIL_SELECT_FIELDS
  };

  // Handle search terms
  const kqlTerms = [];

  if (searchTerms.query)   kqlTerms.push(buildKqlTerm(null, searchTerms.query));
  if (searchTerms.subject) kqlTerms.push(buildKqlTerm('subject', searchTerms.subject));
  if (searchTerms.from)    kqlTerms.push(buildKqlTerm('from', searchTerms.from));
  if (searchTerms.to)      kqlTerms.push(buildKqlTerm('to', searchTerms.to));

  // Add $search if we have any search terms
  if (kqlTerms.length > 0) {
    // Graph API does not support $orderby or $filter with $search
    // Move boolean filters into KQL syntax instead
    addBooleanFiltersAsKQL(kqlTerms, filterTerms);
    // The whole KQL expression is wrapped in ONE pair of double quotes and the
    // terms are AND-joined. Individual terms must not carry their own quotes.
    params.$search = `"${kqlTerms.join(' AND ')}"`;
  } else {
    // No search terms — safe to use $orderby and $filter
    params.$orderby = 'receivedDateTime desc';
    addBooleanFilters(params, filterTerms);
  }
  
  return params;
}

/**
 * Add boolean filters to query parameters as OData $filter
 * Only use when $search is NOT present (they conflict in Graph API)
 * @param {object} params - Query parameters
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 */
function addBooleanFilters(params, filterTerms) {
  const filterConditions = [];
  
  if (filterTerms.hasAttachments === true) {
    filterConditions.push('hasAttachments eq true');
  }
  
  if (filterTerms.unreadOnly === true) {
    filterConditions.push('isRead eq false');
  }
  
  // Add $filter parameter if we have any filter conditions
  if (filterConditions.length > 0) {
    params.$filter = filterConditions.join(' and ');
  }
}

/**
 * Add boolean filters as KQL terms for use with $search
 * Use this instead of addBooleanFilters when $search is present
 * @param {string[]} kqlTerms - Array of KQL terms to append to
 * @param {object} filterTerms - Filter terms (hasAttachments, unreadOnly)
 */
function addBooleanFiltersAsKQL(kqlTerms, filterTerms) {
  if (filterTerms.hasAttachments === true) {
    kqlTerms.push('hasAttachments:true');
  }
  
  if (filterTerms.unreadOnly === true) {
    kqlTerms.push('isRead:false');
  }
}

/**
 * Format search results into a readable text format
 * @param {object} response - The API response object
 * @returns {object} - MCP response object
 */
function formatSearchResults(response) {
  if (!response.value || response.value.length === 0) {
    return {
      content: [{ 
        type: "text", 
        text: `No emails found matching your search criteria.`
      }]
    };
  }
  
  // Format results
  const emailList = response.value.map((email, index) => {
    const sender = email.from?.emailAddress || { name: 'Unknown', address: 'unknown' };
    const date = new Date(email.receivedDateTime).toLocaleString();
    const readStatus = email.isRead ? '' : '[UNREAD] ';
    
    return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})\nSubject: ${email.subject}\nID: ${email.id}\n`;
  }).join("\n");
  
  // Add search strategy info if available
  let additionalInfo = '';
  if (response._searchInfo) {
    additionalInfo = `\n(Search used ${response._searchInfo.strategies[response._searchInfo.strategies.length - 1]} strategy)`;
  }
  
  return {
    content: [{ 
      type: "text", 
      text: `Found ${response.value.length} emails matching your search criteria:${additionalInfo}\n\n${emailList}`
    }]
  };
}

module.exports = handleSearchEmails;
