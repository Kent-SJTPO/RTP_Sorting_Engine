SELECT
    DBNUM,
    ProjectName,
    COUNTY,
    SPONSOR,
    FUND,
    PHASE,
    2026 AS TIPYear,
    [2026] AS TIPAmount
FROM [250903_TIP]
WHERE [2026] Is Not Null AND [2026]<>0

UNION ALL

SELECT
    DBNUM,
    ProjectName,
    COUNTY,
    SPONSOR,
    FUND,
    PHASE,
    2027 AS TIPYear,
    [2027] AS TIPAmount
FROM [250903_TIP]
WHERE [2027] Is Not Null AND [2027]<>0

UNION ALL

SELECT
    DBNUM,
    ProjectName,
    COUNTY,
    SPONSOR,
    FUND,
    PHASE,
    2028 AS TIPYear,
    [2028] AS TIPAmount
FROM [250903_TIP]
WHERE [2028] Is Not Null AND [2028]<>0

UNION ALL

SELECT
    DBNUM,
    ProjectName,
    COUNTY,
    SPONSOR,
    FUND,
    PHASE,
    2029 AS TIPYear,
    [2029] AS TIPAmount
FROM [250903_TIP]
WHERE [2029] Is Not Null AND [2029]<>0

UNION ALL

SELECT
    DBNUM,
    ProjectName,
    COUNTY,
    SPONSOR,
    FUND,
    PHASE,
    2030 AS TIPYear,
    [2030] AS TIPAmount
FROM [250903_TIP]
WHERE [2030] Is Not Null AND [2030]<>0

UNION ALL

SELECT
    DBNUM,
    ProjectName,
    COUNTY,
    SPONSOR,
    FUND,
    PHASE,
    2031 AS TIPYear,
    [2031] AS TIPAmount
FROM [250903_TIP]
WHERE [2031] Is Not Null AND [2031]<>0

UNION ALL

SELECT
    DBNUM,
    ProjectName,
    COUNTY,
    SPONSOR,
    FUND,
    PHASE,
    2032 AS TIPYear,
    [2032] AS TIPAmount
FROM [250903_TIP]
WHERE [2032] Is Not Null AND [2032]<>0

UNION ALL

SELECT
    DBNUM,
    ProjectName,
    COUNTY,
    SPONSOR,
    FUND,
    PHASE,
    2033 AS TIPYear,
    [2033] AS TIPAmount
FROM [250903_TIP]
WHERE [2033] Is Not Null AND [2033]<>0

UNION ALL

SELECT
    DBNUM,
    ProjectName,
    COUNTY,
    SPONSOR,
    FUND,
    PHASE,
    2034 AS TIPYear,
    [2034] AS TIPAmount
FROM [250903_TIP]
WHERE [2034] Is Not Null AND [2034]<>0

UNION ALL SELECT
    DBNUM,
    ProjectName,
    COUNTY,
    SPONSOR,
    FUND,
    PHASE,
    2035 AS TIPYear,
    [2035] AS TIPAmount
FROM [250903_TIP]
WHERE [2035] Is Not Null AND [2035]<>0;

