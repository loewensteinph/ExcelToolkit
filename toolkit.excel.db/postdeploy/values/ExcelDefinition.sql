DELETE FROM ctrl.ColumnMapping
DELETE FROM ctrl.ExcelDefinition

SET IDENTITY_INSERT ctrl.ExcelDefinition ON;
INSERT ctrl.ExcelDefinition (
	DefinitionId,
    IsActive,
    FileName,
    SheetName,
    Range,
    Annotation,
    RangeWidthAuto,
    RangeHeightAuto,
    TargetTable,
    ConnectionString,
    HasHeaderRow,
    DeleteBeforeImport,
    BulkInsert,
    ValidateDataTypes,
    RollbackOnError
)
VALUES
( 1, 1, N'TestWorkbooks\UT1.xlsx', N'Sheet1', N'A1:D5', NULL, 0, 0, N'unittest.UT1', N'Data Source=.;Initial Catalog=ExcelDB;Integrated Security=true', 1, 1, 0, 1, 1 ), 
( 2, 1, N'TestWorkbooks\UT1.xlsx', N'Sheet2', N'M14:P18', NULL, 0, 0, N'unittest.UT1a', N'Data Source=.;Initial Catalog=ExcelDB;Integrated Security=true', 1, 1, 0, 0, 1 ), 
( 3, 1, N'TestWorkbooks\UT2.xlsx', N'Sheet1', N'A1:E5', NULL, 0, 0, N'unittest.UT2', N'Data Source=.;Initial Catalog=ExcelDB;Integrated Security=true', 1, 1, 1, 1, 1 ), 
( 4, 1, N'TestWorkbooks\UT3.xlsx', N'Sheet1', N'A1:A1', NULL, 1, 1, N'unittest.UT3', N'Data Source=.;Initial Catalog=ExcelDB;Integrated Security=true', 1, 1, 1, 1, 1 )
SET IDENTITY_INSERT ctrl.ExcelDefinition OFF