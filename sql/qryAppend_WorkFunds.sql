INSERT INTO WorkFunds ( [year], fund, projected_amount )
SELECT F.[year], F.fund, Round(
        IIf(
            F.[year] < 2035,
            F.projected_amount,
            IIf(
                F.[year] = 2035,
                F.projected_amount * 1.3439163793,
                F.projected_amount * 1.3439163793 * (1.03 ^ (F.[year] - 2035))
            )
        ),
        3
    ) AS adjusted_amount
FROM InputFunds AS F;

