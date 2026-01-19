USE [行政網路系統];
GO

CREATE OR ALTER VIEW dbo.vVisit_WithVisitDT
AS
SELECT
    t.*,
    VisitDT =
        TRY_CONVERT(datetime2(0),
            CONCAT(
                CAST(CAST(LEFT(LTRIM(RTRIM(t.拜訪日期)), 3) AS int) + 1911 AS char(4)),
                SUBSTRING(LTRIM(RTRIM(t.拜訪日期)), 4, LEN(LTRIM(RTRIM(t.拜訪日期))) - 3),
                CASE WHEN LEN(LTRIM(RTRIM(t.拜訪日期))) = 15 THEN ':00' ELSE '' END
            ),
            120
        )
FROM dbo.外訪表 AS t;
GO
