CREATE TABLE [log].[Log4Net] (
    [LogId]     BIGINT         IDENTITY (1, 1) NOT NULL,
    [Date]      DATETIME       NOT NULL,
    [Thread]    NVARCHAR (MAX) NULL,
    [Level]     NVARCHAR (MAX) NULL,
    [Logger]    NVARCHAR (MAX) NULL,
    [Message]   NVARCHAR (MAX) NULL,
    [Exception] NVARCHAR (MAX) NULL,
    CONSTRAINT [PK_log.Log4Net] PRIMARY KEY CLUSTERED ([LogId] ASC)
);

