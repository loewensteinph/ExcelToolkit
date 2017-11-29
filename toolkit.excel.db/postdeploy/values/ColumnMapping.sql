SET IDENTITY_INSERT ctrl.ColumnMapping ON;

INSERT ctrl.ColumnMapping
(
    ColumnMappingId,
    SourceColumn,
    TargetColumn,
    ExcelDefinition_DefinitionId
)
VALUES
(1, N'StringTest', N'StringTest', 1),
(2, N'DecimalTest', N'DecimalTest', 1),
(3, N'IntTest', N'IntTest', 1),
(4, N'GuidTest', N'GuidTest', 1),
(5, N'StringTest', N'StringTest', 2),
(6, N'DecimalTest', N'DecimalTest', 2),
(7, N'IntTest', N'IntTest', 2),
(8, N'GuidTest', N'GuidTest', 2),
(9, N'StringTest', N'StringTest1', 3),
(10, N'DecimalTest', N'DecimalTest1', 3),
(11, N'IntTest', N'IntTest1', 3),
(12, N'GuidTest', N'GuidTest1', 3),
(13, N'DateTest', N'DateTest', 3),
(14, N'StringTest', N'StringTest1', 4),
(15, N'DecimalTest', N'DecimalTest1', 4),
(16, N'IntTest', N'IntTest1', 4),
(17, N'GuidTest', N'GuidTest1', 4),
(18, N'DateTest', N'DateTest', 4);

SET IDENTITY_INSERT ctrl.ColumnMapping OFF;