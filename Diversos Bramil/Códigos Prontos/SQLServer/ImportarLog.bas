Attribute VB_Name = "Module2"
'****************************************************************
'Microsoft SQL Server 2000
'Visual Basic file generated for DTS Package
'File Name: C:\Meus documentos\LOGSIGA.bas
'Package Name: LOGSIGA
'Package Description: DTS package description
'Generated Date: 28/07/03
'Generated Time: 16:56:24
'****************************************************************

Option Explicit
Public goPackageOld As New DTS.Package
Public goPackage As DTS.Package2
Public INI As New GMclINI

Public strPasta As String
Public strDNSIP As String
Public strCatálogo As String
Public strTabela As String
Public strArquivo As String

Private Sub Main()
        strPasta = INI.INILê(App.Path & "\ImportarLog.ini", "Configuração", "Pasta")
        strArquivo = INI.INILê(App.Path & "\ImportarLog.ini", "Configuração", "Arquivo")
        strArquivo = Replace(strArquivo, "DDMMAA", Format(Date - 1, "ddmmyy"))
        
       
        strDNSIP = INI.INILê(App.Path & "\ImportarLog.ini", "Configuração", "DNS/IP")
        strCatálogo = INI.INILê(App.Path & "\ImportarLog.ini", "Configuração", "Catálogo")
        strTabela = INI.INILê(App.Path & "\ImportarLog.ini", "Configuração", "Tabela")

        Set goPackage = goPackageOld

        goPackage.Name = "LOGSIGA"
        goPackage.Description = "DTS package description"
        goPackage.WriteCompletionStatusToNTEventLog = False
        goPackage.FailOnError = False
        goPackage.PackagePriorityClass = 2
        goPackage.MaxConcurrentSteps = 4
        goPackage.LineageOptions = 0
        goPackage.UseTransaction = True
        goPackage.TransactionIsolationLevel = 4096
        goPackage.AutoCommitTransaction = True
        goPackage.RepositoryMetadataOptions = 0
        goPackage.UseOLEDBServiceComponents = True
        goPackage.LogToSQLServer = False
        goPackage.LogServerFlags = 0
        goPackage.FailPackageOnLogFailure = False
        goPackage.ExplicitGlobalVariables = False
        goPackage.PackageType = 0
        

Dim oConnProperty As DTS.OleDBProperty

'---------------------------------------------------------------------------
' create package connection information
'---------------------------------------------------------------------------

Dim oConnection As DTS.Connection2

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("DTSFlatFile")

        oConnection.ConnectionProperties("Data Source") = "E:\AP6\LOG\Ja importadas\sl280703.log"
        oConnection.ConnectionProperties("Mode") = 1
        oConnection.ConnectionProperties("Row Delimiter") = vbCrLf
        oConnection.ConnectionProperties("File Format") = 2
        oConnection.ConnectionProperties("Column Lengths") = "1,2,2,3,15,8,6,10,3,6,8,16"
        oConnection.ConnectionProperties("File Type") = 1
        oConnection.ConnectionProperties("Skip Rows") = 0
        oConnection.ConnectionProperties("First Row Column Name") = False
        oConnection.ConnectionProperties("Number of Column") = 12
        
        oConnection.Name = "Connection 1"
        oConnection.ID = 1
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = "E:\AP6\LOG\Ja importadas\sl280703.log"
        oConnection.ConnectionTimeout = 60
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("SQLOLEDB")

        oConnection.ConnectionProperties("Persist Security Info") = True
        oConnection.ConnectionProperties("User ID") = "sa"
        oConnection.ConnectionProperties("Initial Catalog") = "DADOSPRO"
        oConnection.ConnectionProperties("Data Source") = "SERVBDTESTE\PROTHEUS"
        oConnection.ConnectionProperties("Application Name") = "DTS  Import/Export Wizard"
        
        oConnection.Name = "Connection 2"
        oConnection.ID = 2
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = "SERVBDTESTE\PROTHEUS"
        oConnection.UserID = "sa"
        oConnection.ConnectionTimeout = 60
        oConnection.Catalog = "DADOSPRO"
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'---------------------------------------------------------------------------
' create package steps information
'---------------------------------------------------------------------------

Dim oStep As DTS.Step2
Dim oPrecConstraint As DTS.PrecedenceConstraint

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Create Table [DADOSPRO].[dbo].[SIGALOG] Step"
        oStep.Description = "Create Table [DADOSPRO].[dbo].[SIGALOG] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Create Table [DADOSPRO].[dbo].[SIGALOG] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Copy Data from sl280703 to [DADOSPRO].[dbo].[SIGALOG] Step"
        oStep.Description = "Copy Data from sl280703 to [DADOSPRO].[dbo].[SIGALOG] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copy Data from sl280703 to [DADOSPRO].[dbo].[SIGALOG] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a precedence constraint for steps defined below

Set oStep = goPackage.Steps("Copy Data from sl280703 to [DADOSPRO].[dbo].[SIGALOG] Step")
Set oPrecConstraint = oStep.PrecedenceConstraints.New("Create Table [DADOSPRO].[dbo].[SIGALOG] Step")
        oPrecConstraint.StepName = "Create Table [DADOSPRO].[dbo].[SIGALOG] Step"
        oPrecConstraint.PrecedenceBasis = 0
        oPrecConstraint.Value = 4
        
oStep.PrecedenceConstraints.Add oPrecConstraint
Set oPrecConstraint = Nothing

'---------------------------------------------------------------------------
' create package tasks information
'---------------------------------------------------------------------------

'------------- call Task_Sub1 for task Create Table [DADOSPRO].[dbo].[SIGALOG] Task (Create Table [DADOSPRO].[dbo].[SIGALOG] Task)
Call Task_Sub1(goPackage)

'------------- call Task_Sub2 for task Copy Data from sl280703 to [DADOSPRO].[dbo].[SIGALOG] Task (Copy Data from sl280703 to [DADOSPRO].[dbo].[SIGALOG] Task)
Call Task_Sub2(goPackage)

'---------------------------------------------------------------------------
' Save or execute package
'---------------------------------------------------------------------------

'goPackage.SaveToSQLServer "(local)", "sa", ""
goPackage.Execute
goPackage.UnInitialize
'to save a package instead of executing it, comment out the executing package line above and uncomment the saving package line
Set goPackage = Nothing

Set goPackageOld = Nothing

End Sub


'------------- define Task_Sub1 for task Create Table [DADOSPRO].[dbo].[SIGALOG] Task (Create Table [DADOSPRO].[dbo].[SIGALOG] Task)
Public Sub Task_Sub1(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
Set oCustomTask1 = oTask.CustomTask

        oCustomTask1.Name = "Create Table [DADOSPRO].[dbo].[SIGALOG] Task"
        oCustomTask1.Description = "Create Table [DADOSPRO].[dbo].[SIGALOG] Task"
        oCustomTask1.SQLStatement = "CREATE TABLE [DADOSPRO].[dbo].[SIGALOG] (" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_TPTRANSAC] char (1) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_EMPRESA] char (2) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_FILIAL] char (2) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_MODULO] char (3) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_USUARIO] char (15) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_PROGRAMA] char (8) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_HORA] char (6) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_REGISTRO] char (10) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_TABELA] char (3) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_CONTROLE] char (6) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_DTTRANSAC] char (8) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[ZZ_OBSERVAC] char (16) NULL" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & ")"
        oCustomTask1.ConnectionID = 2
        oCustomTask1.CommandTimeout = 0
        oCustomTask1.OutputAsRecordset = False
        
goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub2 for task Copy Data from sl280703 to [DADOSPRO].[dbo].[SIGALOG] Task (Copy Data from sl280703 to [DADOSPRO].[dbo].[SIGALOG] Task)
Public Sub Task_Sub2(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask2 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
Set oCustomTask2 = oTask.CustomTask

        oCustomTask2.Name = "Copy Data from sl280703 to [DADOSPRO].[dbo].[SIGALOG] Task"
        oCustomTask2.Description = "Copy Data from sl280703 to [DADOSPRO].[dbo].[SIGALOG] Task"
        oCustomTask2.SourceConnectionID = 1
        oCustomTask2.SourceObjectName = "E:\AP6\LOG\Ja importadas\sl280703.log"
        oCustomTask2.DestinationConnectionID = 2
        oCustomTask2.DestinationObjectName = "[DADOSPRO].[dbo].[SIGALOG]"
        oCustomTask2.ProgressRowCount = 1000
        oCustomTask2.MaximumErrorCount = 0
        oCustomTask2.FetchBufferSize = 1
        oCustomTask2.UseFastLoad = True
        oCustomTask2.InsertCommitSize = 0
        oCustomTask2.ExceptionFileColumnDelimiter = "|"
        oCustomTask2.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask2.AllowIdentityInserts = False
        oCustomTask2.FirstRow = 0
        oCustomTask2.LastRow = 0
        oCustomTask2.FastLoadOptions = 2
        oCustomTask2.ExceptionFileOptions = 1
        oCustomTask2.DataPumpOptions = 0
        
Call oCustomTask2_Trans_Sub1(oCustomTask2)
                
                
goPackage.Tasks.Add oTask
Set oCustomTask2 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask2_Trans_Sub1(ByVal oCustomTask2 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask2.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4
                
                Set oColumn = oTransformation.SourceColumns.New("Col001", 1)
                        oColumn.Name = "Col001"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 48
                        oColumn.Size = 1
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col002", 2)
                        oColumn.Name = "Col002"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 48
                        oColumn.Size = 2
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col003", 3)
                        oColumn.Name = "Col003"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 48
                        oColumn.Size = 2
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col004", 4)
                        oColumn.Name = "Col004"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 48
                        oColumn.Size = 3
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col005", 5)
                        oColumn.Name = "Col005"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 48
                        oColumn.Size = 15
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col006", 6)
                        oColumn.Name = "Col006"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 48
                        oColumn.Size = 8
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col007", 7)
                        oColumn.Name = "Col007"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 48
                        oColumn.Size = 6
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col008", 8)
                        oColumn.Name = "Col008"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 48
                        oColumn.Size = 10
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col009", 9)
                        oColumn.Name = "Col009"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 48
                        oColumn.Size = 3
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col010", 10)
                        oColumn.Name = "Col010"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 48
                        oColumn.Size = 6
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col011", 11)
                        oColumn.Name = "Col011"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 48
                        oColumn.Size = 8
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col012", 12)
                        oColumn.Name = "Col012"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 48
                        oColumn.Size = 16
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_TPTRANSAC", 1)
                        oColumn.Name = "ZZ_TPTRANSAC"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 48
                        oColumn.Size = 1
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_EMPRESA", 2)
                        oColumn.Name = "ZZ_EMPRESA"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 48
                        oColumn.Size = 2
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_FILIAL", 3)
                        oColumn.Name = "ZZ_FILIAL"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 48
                        oColumn.Size = 2
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_MODULO", 4)
                        oColumn.Name = "ZZ_MODULO"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 48
                        oColumn.Size = 3
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_USUARIO", 5)
                        oColumn.Name = "ZZ_USUARIO"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 48
                        oColumn.Size = 15
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_PROGRAMA", 6)
                        oColumn.Name = "ZZ_PROGRAMA"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 48
                        oColumn.Size = 8
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_HORA", 7)
                        oColumn.Name = "ZZ_HORA"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 48
                        oColumn.Size = 6
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_REGISTRO", 8)
                        oColumn.Name = "ZZ_REGISTRO"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 48
                        oColumn.Size = 10
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_TABELA", 9)
                        oColumn.Name = "ZZ_TABELA"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 48
                        oColumn.Size = 3
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_CONTROLE", 10)
                        oColumn.Name = "ZZ_CONTROLE"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 48
                        oColumn.Size = 6
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_DTTRANSAC", 11)
                        oColumn.Name = "ZZ_DTTRANSAC"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 48
                        oColumn.Size = 8
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ZZ_OBSERVAC", 12)
                        oColumn.Name = "ZZ_OBSERVAC"
                        oColumn.Ordinal = 12
                        oColumn.Flags = 48
                        oColumn.Size = 16
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties

                
        Set oTransProps = Nothing

        oCustomTask2.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub

