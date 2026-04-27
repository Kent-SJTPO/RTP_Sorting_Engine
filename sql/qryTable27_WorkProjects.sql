SELECT wp.project_id AS [RTP ID NO], wp.project_name AS [Project Name], wp.County AS Sponsor, wp.score AS [Project Score], wp.year_eligible AS [Request Year], Round(wp.cost, 3) AS [Request Amount], wp.fund AS Fund, wp.award_year AS [Programmed Year], Round(wp.cost, 3) AS [Programmed Amount]
FROM WorkProjects AS wp
WHERE wp.award_year BETWEEN 2025 AND 2050
    AND wp.awarded = True
    AND wp.DBNUM NOT LIKE "S*"
ORDER BY wp.award_year, wp.fund, wp.score DESC;

