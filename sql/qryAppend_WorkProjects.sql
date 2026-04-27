INSERT INTO WorkProjects ( project_id, project_name, fund, cost, year_eligible, score, award_year, DBNUM, awarded )
SELECT I.project_id, I.project_name, I.fund, I.cost, I.year_eligible, I.score, 0 AS award_year, I.DBNUM, False AS awarded
FROM InputProjects AS I
WHERE I.project_id IS NOT NULL
    AND I.project_name IS NOT NULL
    AND I.fund IN (
        'STBGP-AC',
        'STBGP-B50K200K',
        'STBGP-B5K50K',
        'STBGP-L5K'
    )
    AND I.cost > 0
    AND I.year_eligible IS NOT NULL
    AND I.score IS NOT NULL;

