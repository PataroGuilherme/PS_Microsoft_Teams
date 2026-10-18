
#1. GO TO GRAPH EXPLORER ON WEB and DO LOGIN
# https://developer.microsoft.com/en-us/graph/graph-explorer

#STEP 1 = CREATE THE SEARCH

(POST - beta) https://graph.microsoft.com/beta/security/auditLog/queries

# Go to Request Body, adn Apply the JSON below

{
  "displayName": "TITLE OF YOUR SEARCH",
  "filterStartDateTime": "2026-10-15T00:00:00Z",
  "filterEndDateTime": "2026-10-17T23:59:59Z",
  "recordTypeFilters": ["microsoftTeams"],
  "objectIdFilters": [
    "GROUP-OBJECT-ID",
    "ID-DO-PRIVATE-CHANNEL" #OR WITHOUT PRIVATE CHANNEL
  ]
}

#OR

{
  "displayName": "TITLE OF YOUR SEARCH",
  "filterStartDateTime": "2026-10-15T03:00:00Z",
  "filterEndDateTime": "2026-10-18T02:59:59Z",
  "recordTypeFilters": ["microsoftTeams"],
  "keywordFilter": "TEAM-NAME"
}

# THIS STEP WILL GENERATE A SEARCH ID. KEEP THE ID NUMBER SAFE!
# Example: "000-0000-000"


#STEP 2 = FOLLOW THE STATUS
# Repeat the 'Run query' untill the 'succeeded' status.

(GET - beta) https://graph.microsoft.com/beta/security/auditLog/queries/{SEARCH-ID}


#STEP 3 = CHECK THE RESULTS

(GET - beta) https://graph.microsoft.com/beta/security/auditLog/queries/{SEARCH-ID}/records?$top=100