SELECT X.fund, Count(X.year) AS YearCount, Sum(X.AvailableFunds) AS TotalAvailableFunds, Sum(X.AwardedAmount) AS TotalAwardedAmount, Sum(X.RemainingAmount) AS TotalRemainingAmount, Sum(IIf(X.AwardedAmount = 0, 1, 0)) AS UnusedYearCount, Round(         Sum(X.AwardedAmount) / IIf(Sum(X.AvailableFunds)=0, Null, Sum(X.AvailableFunds)),         3     ) AS OverallUtilizationRate, Round(Sum(X.AwardedAmount) / Count(X.year), 3) AS AvgAwardedPerYear, Round(Sum(X.AvailableFunds) / Count(X.year), 3) AS AvgAvailablePerYear
FROM (SELECT F.fund, F.year, F.projected_amount AS AvailableFunds, Nz(A.TotalAwarded,0) AS AwardedAmount, F.projected_amount - Nz(A.TotalAwarded,0) AS RemainingAmount FROM InputFunds AS F LEFT JOIN (SELECT fund, award_year, Sum(project_cost) AS TotalAwarded FROM OutputAwardLog GROUP BY fund, award_year)  AS A ON (F.fund = A.fund) AND (F.year = A.award_year))  AS X
GROUP BY X.fund
ORDER BY X.fund;

