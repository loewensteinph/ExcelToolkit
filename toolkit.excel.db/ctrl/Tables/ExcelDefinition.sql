CREATE TABLE [ctrl].[ExcelDefinition] (
    [DefinitionId]       INT            IDENTITY (1, 1) NOT NULL,
    [IsActive]           BIT            NOT NULL,
    [FileName]           NVARCHAR (MAX) NULL,
    [SheetName]          NVARCHAR (MAX) NULL,
    [Range]              NVARCHAR (MAX) NULL,
    [Annotation]         NVARCHAR (MAX) NULL,
    [RangeWidthAuto]     BIT            NOT NULL,
    [RangeHeightAuto]    BIT            NOT NULL,
    [TargetTable]        NVARCHAR (MAX) NULL,
    [ConnectionString]   NVARCHAR (MAX) NULL,
    [HasHeaderRow]       BIT            NOT NULL,
    [DeleteBeforeImport] BIT            NOT NULL,
    [BulkInsert]         BIT            NOT NULL,
    [ValidateDataTypes]  BIT            NOT NULL,
    [RollbackOnError]    BIT            NOT NULL,
    CONSTRAINT [PK_ctrl.ExcelDefinition] PRIMARY KEY CLUSTERED ([DefinitionId] ASC)
);

