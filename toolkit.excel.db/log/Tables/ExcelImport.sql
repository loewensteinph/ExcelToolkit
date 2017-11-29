CREATE TABLE [log].[ExcelImport] (
    [ImportId]                INT            IDENTITY (1, 1) NOT NULL,
    [ImportTimestamp]         DATETIME       NOT NULL,
    [RowsImported]            INT            NOT NULL,
    [RowsWithErrors]          INT            NOT NULL,
    [ResultStatus]            NVARCHAR (MAX) NULL,
    [Definition_DefinitionId] INT            NULL,
    CONSTRAINT [PK_log.ExcelImport] PRIMARY KEY CLUSTERED ([ImportId] ASC),
    CONSTRAINT [FK_log.ExcelImport_ctrl.ExcelDefinition_Definition_DefinitionId] FOREIGN KEY ([Definition_DefinitionId]) REFERENCES [ctrl].[ExcelDefinition] ([DefinitionId])
);


GO
CREATE NONCLUSTERED INDEX [IX_Definition_DefinitionId]
    ON [log].[ExcelImport]([Definition_DefinitionId] ASC);

