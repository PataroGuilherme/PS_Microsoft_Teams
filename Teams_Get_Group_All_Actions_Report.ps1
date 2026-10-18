#1. GO TO GRAPH EXPLORER ON WEB and DO LOGIN
# https://developer.microsoft.com/en-us/graph/graph-explorer

# GET YOUR GROUP OBJECT-ID AND CHANGE IT on THE URL BELOW
# ALSO CHANGE THE SEARCH DATE PERIOD SCOPE (UTC)


(GET - V1.0) https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?$filter=activityDateTime ge 2026-10-15T03:00:00Z and activityDateTime le 2026-10-18T02:59:59Z and targetResources/any(t:t/id eq '3d771427-306c-4471-8525-3ed674b4ae24')&$orderby=activityDateTime desc&$top=100