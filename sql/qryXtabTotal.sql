TRANSFORM Sum(Nz([fund_used],0) + Nz([pool_used],0)) AS AwardAmount
SELECT Q.first_award_year, Q.project_id, Q.project_name, Q.fund
FROM (SELECT
            O.project_id,
            O.project_name,
            O.fund,
            O.award_year,
            O.fund_used,
            O.pool_used,
            DMin(
                "award_year",
                "OutputAwardLog",
                "project_id='" & [O].[project_id] & "'"
            ) AS first_award_year
        FROM OutputAwardLog AS O
    )  AS Q
GROUP BY Q.first_award_year, Q.project_id, Q.project_name, Q.fund
ORDER BY Q.first_award_year, Q.project_id
PIVOT Q.award_year IN (
    2026, 2027, 2028, 2029, 2030,
    2031, 2032, 2033, 2034, 2035,
    2036, 2037, 2038, 2039, 2040,
    2041, 2042, 2043, 2044, 2045,
    2046, 2047, 2048, 2049, 2050
);

