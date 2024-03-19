DECLARE @__ToBeginOfDay_1 datetime = '2023-01-17T00:00:00.000';

DECLARE @__ToEndOfDay_2 datetime = '2024-02-09T23:59:59.000';

SELECT
    [c].[Id] AS [ContactId],
    [c].[First] AS [FirstName],
    [c].[Last] AS [LastName],
    [c].[Phone1] AS [Phone],
    [c].[Email],
    [c].[Company] AS [CompanyName],
    [c].[City] AS [AddressCity],
    [c].[Country] AS [AddressCountry],
    [c].[PostalCode] AS [AddressPostalCode],
    [c].[Province] AS [AddressProvince],
    [c].[Address] AS [AddressStreet],
    iv.TotalWithTax,
    iv.TotalWithoutTax,
    iv.InvoicesCount,
    iv.LastInvoiceTime,
    iv.invoiceTime,
    iv.locationId
FROM
    [Contact] AS [c]
    left join (
        select
            i.contactId,
            i.locationId,
            i.invoiceTime,
            iLast.LastInvoiceTime,
            iSum.TotalWithTax,
            iSum.TotalWithoutTax,
            iSum.InvoicesCount
        from
            invoice i
            inner join (
                select
                    iL.contactId,
                    iL.locationId,
                    MAX([iL].[InvoiceTime]) AS [LastInvoiceTime]
                from
                    invoice iL
                group by
                    iL.contactId,
                    iL.locationId
            ) as iLast on i.contactId = iLast.contactId
            and i.locationId = iLast.locationId
            inner join (
                select
                    iSum.contactId,
                    iSum.locationId,
                    Sum(iSum.InvoiceTotal) AS TotalWithTax,
                    Sum(iSum.InvoiceTotal - iSum.taxTotal) AS TotalWithoutTax,
                    Count(iSum.Id) as InvoicesCount
                from
                    invoice iSum
                group by
                    iSum.contactId,
                    iSum.locationId
            ) as iSum on i.contactId = iSum.contactId
            and i.locationId = iSum.locationId
        where
            i.invoiceTime >= @__ToBeginOfDay_1
            AND [i].[InvoiceTime] <= @__ToEndOfDay_2
    ) as iv on c.Id = iv.contactId
where
    c.ID = '42B043B4-1EDF-4012-A4CE-3245F5889DA3'