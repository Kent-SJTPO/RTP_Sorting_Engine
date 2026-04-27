TRANSFORM Sum(Nz([fund_used],0) + Nz([pool_used],0)) AS AwardAmount
SELECT project_id, project_name, fund
FROM OutputAwardLog
GROUP BY project_id, project_name, fund
PIVOT award_year IN (
    2026,2027,2028,2029,2030,
    2031,2032,2033,2034,2035,
    2036,2037,2038,2039,2040,
    2041,2042,2043,2044,2045,
    2046,2047,2048,2049,2050
);

