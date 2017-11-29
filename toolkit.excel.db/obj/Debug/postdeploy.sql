/*
Post-Deployment Script Template							
--------------------------------------------------------------------------------------
 This file contains SQL statements that will be appended to the build script.		
 Use SQLCMD syntax to include a file in the post-deployment script.			
 Example:      :r .\myfile.sql								
 Use SQLCMD syntax to reference a variable in the post-deployment script.		
 Example:      :setvar TableName MyTable							
               SELECT * FROM [$(TableName)]					
--------------------------------------------------------------------------------------
*/
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
GO
