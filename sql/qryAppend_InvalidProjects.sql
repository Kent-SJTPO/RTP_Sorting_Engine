INSERT INTO InvalidProjects ( project_id, project_name, fund, cost, year_eligible, score, award_year, DBNUM, invalid_reason )
SELECT I.project_id, I.project_name, I.fund, I.cost, I.year_eligible, I.score, I.award_year, I.DBNUM, IIf(I.project_id Is Null, 'Missing project_id',
    IIf(I.project_name Is Null, 'Missing project_name',
    IIf(I.fund Is Null, 'Missing fund',
    IIf(I.cost Is Null, 'Missing cost',
    IIf(I.cost<=0, 'Cost must be > 0',
    IIf(I.year_eligible Is Null, 'Missing year_eligible',
    IIf(I.score Is Null, 'Missing score',
    IIf(I.fund Not In ('STBGP-AC','STBGP-B50K200K','STBGP-B5K50K','STBGP-L5K'),
        'Unsupported fund',
        'Unknown invalid record'
    )))))))) AS invalid_reason
FROM InputProjects AS I
WHERE I.project_id Is Null
    OR I.project_name Is Null
    OR I.fund Is Null
    OR I.cost Is Null
    OR I.cost<=0
    OR I.year_eligible Is Null
    OR I.score Is Null
    OR I.fund Not In ('STBGP-AC','STBGP-B50K200K','STBGP-B5K50K','STBGP-L5K');

