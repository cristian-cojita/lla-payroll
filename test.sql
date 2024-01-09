SELECT
    [i].[Id] AS [InvoiceId],
    [i].[ContactId],
    [i].[LocationId],
    [l].[Name] AS [LocationName],
    [c].[Company] AS [CompanyName],
    [c].[First] AS [FirstName],
    [c].[Last] AS [LastName],
    [c].[Latitude],
    [c].[Longitude],
    [i].[PaymentMethod],
    [i].[InvoiceNumber],
    [i].[InvoiceTime] AS [InvoiceDate],
    CONVERT(float, [i].[InvoiceTotal] - [i].[TaxTotal]) AS [TotalWithoutTax],
    CONVERT(float, [i].[InvoiceTotal]) AS [TotalWithTax],
    [i].[ContactType],
    COALESCE(
        SUM(
            CASE
                WHEN [i1].[Id] IS NOT NULL
                AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotal]
                ELSE 0.0
            END
        ),
        0.0
    ) AS [Invoiced],
    COALESCE(
        SUM(
            CASE
                WHEN [i1].[Id] IS NOT NULL
                AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotalCost]
                ELSE 0.0
            END
        ),
        0.0
    ) AS [Cost],
    CASE
        WHEN COALESCE(
            SUM(
                CASE
                    WHEN [i1].[Id] IS NOT NULL
                    AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotal]
                    ELSE 0.0
                END
            ),
            0.0
        ) = 0.0 THEN CAST(
            CASE
                WHEN COALESCE(
                    SUM(
                        CASE
                            WHEN [i1].[Id] IS NOT NULL
                            AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotalCost]
                            ELSE 0.0
                        END
                    ),
                    0.0
                ) = 0.0 THEN 0
                ELSE -9999
            END AS decimal(18, 4)
        )
        ELSE (
            (
                COALESCE(
                    SUM(
                        CASE
                            WHEN [i1].[Id] IS NOT NULL
                            AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotal]
                            ELSE 0.0
                        END
                    ),
                    0.0
                ) - COALESCE(
                    SUM(
                        CASE
                            WHEN [i1].[Id] IS NOT NULL
                            AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotalCost]
                            ELSE 0.0
                        END
                    ),
                    0.0
                )
            ) / COALESCE(
                SUM(
                    CASE
                        WHEN [i1].[Id] IS NOT NULL
                        AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotal]
                        ELSE 0.0
                    END
                ),
                0.0
            )
        ) * 100.0
    END AS [Margin],
    COUNT(
        CASE
            WHEN [i0].[Id] IS NOT NULL
            AND [i1].[Id] IS NOT NULL THEN 1
        END
    ) AS [SpLines]
FROM
    [Invoice] AS [i]
    LEFT JOIN [InvoicePackage] AS [i0] ON [i].[Id] = [i0].[InvoiceId]
    LEFT JOIN [InvoicePackageLine] AS [i1] ON [i0].[Id] = [i1].[InvoicePackageId]
    INNER JOIN [Location] AS [l] ON [i].[LocationId] = [l].[Id]
    INNER JOIN [Contact] AS [c] ON [i].[ContactId] = [c].[Id]
GROUP BY
    [i].[Id],
    [i].[ContactId],
    [i].[LocationId],
    [l].[Name],
    [c].[Company],
    [c].[First],
    [c].[Last],
    [c].[Latitude],
    [c].[Longitude],
    [i].[PaymentMethod],
    [i].[InvoiceNumber],
    [i].[InvoiceTime],
    [i].[InvoiceTotal],
    [i].[TaxTotal],
    [i].[ContactType]
HAVING
    [i].[InvoiceTime] >= '2023-12-01T00:00:00.000'
    AND [i].[InvoiceTime] <= '2023-12-31T23:59:59.000'
    AND CONVERT(float, [i].[InvoiceTotal]) >= 0.0E0
    AND CASE
        WHEN COALESCE(
            SUM(
                CASE
                    WHEN [i1].[Id] IS NOT NULL
                    AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotal]
                    ELSE 0.0
                END
            ),
            0.0
        ) = 0.0 THEN CAST(
            CASE
                WHEN COALESCE(
                    SUM(
                        CASE
                            WHEN [i1].[Id] IS NOT NULL
                            AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotalCost]
                            ELSE 0.0
                        END
                    ),
                    0.0
                ) = 0.0 THEN 0
                ELSE -9999
            END AS decimal(18, 4)
        )
        ELSE (
            (
                COALESCE(
                    SUM(
                        CASE
                            WHEN [i1].[Id] IS NOT NULL
                            AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotal]
                            ELSE 0.0
                        END
                    ),
                    0.0
                ) - COALESCE(
                    SUM(
                        CASE
                            WHEN [i1].[Id] IS NOT NULL
                            AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotalCost]
                            ELSE 0.0
                        END
                    ),
                    0.0
                )
            ) / COALESCE(
                SUM(
                    CASE
                        WHEN [i1].[Id] IS NOT NULL
                        AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotal]
                        ELSE 0.0
                    END
                ),
                0.0
            )
        ) * 100.0
    END > 0.0
    AND CASE
        WHEN COALESCE(
            SUM(
                CASE
                    WHEN [i1].[Id] IS NOT NULL
                    AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotal]
                    ELSE 0.0
                END
            ),
            0.0
        ) = 0.0 THEN CAST(
            CASE
                WHEN COALESCE(
                    SUM(
                        CASE
                            WHEN [i1].[Id] IS NOT NULL
                            AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotalCost]
                            ELSE 0.0
                        END
                    ),
                    0.0
                ) = 0.0 THEN 0
                ELSE -9999
            END AS decimal(18, 4)
        )
        ELSE (
            (
                COALESCE(
                    SUM(
                        CASE
                            WHEN [i1].[Id] IS NOT NULL
                            AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotal]
                            ELSE 0.0
                        END
                    ),
                    0.0
                ) - COALESCE(
                    SUM(
                        CASE
                            WHEN [i1].[Id] IS NOT NULL
                            AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotalCost]
                            ELSE 0.0
                        END
                    ),
                    0.0
                )
            ) / COALESCE(
                SUM(
                    CASE
                        WHEN [i1].[Id] IS NOT NULL
                        AND [i1].[IsInvoicing] = CAST(1 AS tinyint) THEN [i1].[ExtendedTotal]
                        ELSE 0.0
                    END
                ),
                0.0
            )
        ) * 100.0
    END <= 20.0
    AND COUNT(
        CASE
            WHEN [i0].[Id] IS NOT NULL
            AND [i1].[Id] IS NOT NULL THEN 1
        END
    ) > 0