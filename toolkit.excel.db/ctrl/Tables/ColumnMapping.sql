CREATE TABLE [ctrl].[ColumnMapping] (
    [ColumnMappingId]              INT            IDENTITY (1, 1) NOT NULL,
    [SourceColumn]                 NVARCHAR (MAX) NULL,
    [TargetColumn]                 NVARCHAR (MAX) NULL,
    [ExcelDefinition_DefinitionId] INT            NULL,
    CONSTRAINT [PK_ctrl.ColumnMapping] PRIMARY KEY CLUSTERED ([ColumnMappingId] ASC),
    CONSTRAINT [FK_ctrl.ColumnMapping_ctrl.ExcelDefinition_ExcelDefinition_DefinitionId] FOREIGN KEY ([ExcelDefinition_DefinitionId]) REFERENCES [ctrl].[ExcelDefinition] ([DefinitionId])
);


GO
CREATE NONCLUSTERED INDEX [IX_ExcelDefinition_DefinitionId]
    ON [ctrl].[ColumnMapping]([ExcelDefinition_DefinitionId] ASC);

