SELECT
    a.fund,
    Count(a.project_id) AS AwardCount,
    Sum(a.project_cost) AS TotalAwarded,
    Min(a.award_year) AS FirstYear,
    Max(a.award_year) AS LastYear,
    Max(a.award_year) - Min(a.award_year) + 1 AS YearSpan,
    Round(
        Sum(a.project_cost) / (Max(a.award_year) - Min(a.award_year) + 1),
        3
    ) AS AvgPerYear
FROM
    OutputAwardLog AS a
GROUP BY
    a.fund

UNION ALL SELECT
    "ALL FUNDS" AS fund,
    Count(a.project_id) AS AwardCount,
    Sum(a.project_cost) AS TotalAwarded,
    Min(a.award_year) AS FirstYear,
    Max(a.award_year) AS LastYear,
    Max(a.award_year) - Min(a.award_year) + 1 AS YearSpan,
    Round(
        Sum(a.project_cost) / (Max(a.award_year) - Min(a.award_year) + 1),
        3
    ) AS AvgPerYear
FROM
    OutputAwardLog AS a;

