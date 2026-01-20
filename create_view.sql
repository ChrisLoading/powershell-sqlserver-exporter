USE [<YourDatabaseName>];
GO

CREATE OR ALTER VIEW <YourViewName>
AS
SELECT
    t.*,
    <TargetColumnName> =
        TRY_CONVERT(datetime2(0),
            CONCAT(
                CAST(CAST(LEFT(LTRIM(RTRIM(t.<SourceColumnName>)), 3) AS int) + 1911 AS char(4)),
                SUBSTRING(LTRIM(RTRIM(t.<SourceColumnName>)), 4, LEN(LTRIM(RTRIM(t.<SourceColumnName>))) - 3),
                CASE WHEN LEN(LTRIM(RTRIM(t.<SourceColumnName>))) = 15 THEN ':00' ELSE '' END
            ),
            120
        )
FROM <YourTableName> AS t;
GO
