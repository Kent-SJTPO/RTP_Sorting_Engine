INSERT INTO
    InvalidProjects (
        project_id,
        project_name,
        fund,
        cost,
        year_eligible,
        score,
        award_year,
        DBNUM,
        invalid_reason
    )
SELECT
    I.project_id,
    I.project_name,
    I.fund,
    I.cost,
    I.year_eligible,
    I.score,
    I.award_year,
    I.DBNUM,
    IIf(
        I.project_id IS NULL,
        'Missing project_id',
        IIf(
            I.project_name IS NULL,
            'Missing project_name',
            IIf(
                I.fund IS NULL,
                'Missing fund',
                IIf(
                    I.cost IS NULL,
                    'Missing cost',
                    IIf(
                        I.cost <= 0,
                        'Cost must be > 0',
                        IIf(
                            I.year_eligible IS NULL,
                            'Missing year_eligible',
                            IIf(
                                I.score IS NULL,
                                'Missing score',
                                IIf(
                                    I.fund NOT IN (
                                        'STBGP-AC',
                                        'STBGP-B50K200K',
                                        'STBGP-B5K50K',
                                        'STBGP-L5K'
                                    ),
                                    'Unsupported fund',
                                    'Unknown invalid record'
                                )
                            )
                        )
                    )
                )
            )
        )
    ) AS invalid_reason
FROM
    InputProjects AS I
WHERE
    I.project_id IS NULL
    OR I.project_name IS NULL
    OR I.fund IS NULL
    OR I.cost IS NULL
    OR I.cost <= 0
    OR I.year_eligible IS NULL
    OR I.score IS NULL
    OR I.fund NOT IN (
        'STBGP-AC',
        'STBGP-B50K200K',
        'STBGP-B5K50K',
        'STBGP-L5K'
    );