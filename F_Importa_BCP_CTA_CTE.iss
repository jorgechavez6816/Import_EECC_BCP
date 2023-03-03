'Desarrollado por Jorge M. Chávez
'Fecha: 01/03/2023

Sub Main
	IgnoreWarning(True)
	Call ReportReaderImport()	'D:\RUC1\DATA\Archivos fuente.ILB\202201_BCP.pdf
	Call AppendField()	'BCP2022.IMD
	Call AppendField1()	'BCP2022.IMD
	Call AppendField2()	'BCP2022.IMD
	Call AppendField4()	'BCP2022.IMD
	Call AppendField5()	'BCP2022.IMD
	Call AppendField6()	'BCP2022.IMD
	Call AppendField7()	'BCP2022.IMD
	Call AppendField8()	'BCP2022.IMD
	Call AppendField9()	'BCP2022.IMD
	Call AppendField10()	'BCP2022.IMD
	Call AppendField11()	'BCP2022.IMD
	Call DirectExtraction()	'BCP2022.IMD
	Call Summarization()	'BCP2022.IMD
	Client.RefreshFileExplorer
	Client.CloseAll
	Client.DeleteDatabase "BCP2022"
	Dim pm As Object
	Dim SourcePath As String
	Dim DestinationPath As String
	Set SourcePath = Client.WorkingDirectory
	Set DestinationPath = "D:\RUC1\DATA\_EECC"
	Client.RunAtServer False
	Set pm = Client.ProjectManagement
	pm.MoveDatabase SourcePath + "F_BCP2022.IMD", DestinationPath
	pm.MoveDatabase SourcePath + "F.1_Resumen_BCP.IMD", DestinationPath
	Set pm = Nothing
End Sub

' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport
	dbName = "BCP2022.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\BCP_CTA_CTE.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\2022_BCP.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function


' Anexar campo
Function AppendField
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "MED"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@IF(@Isini(""TLC""; DESC )=0; @IF(@Isini(""BPI""; DESC )=0; @IF(@Isini(""INT""; DESC )=0; @IF(@Isini(""VEN""; DESC )=0; ""CAJ""; ""VEN""); ""INT""); ""BPI"");""TLC"")"
	field.Length = 3
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anexar campo
Function AppendField1
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DESCRIPCION"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@SPLIT(DESC;MED;"""";1;1)"
	field.Length = 40
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anexar campo
Function AppendField2
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "HORA"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@IF(@RIGHT(@SPLIT(DESC;"":"";"""";1;1);2)+"":""+@LEFT((@SPLIT(DESC;"":"";"""";1;0);2))="":""; """"; @RIGHT(@SPLIT(DESC;"":"";"""";1;1);2)+"":""+@LEFT((@SPLIT(DESC;"":"";"""";1;0);2)))"
	field.Length = 5
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anexar campo
Function AppendField4
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ORIGEN"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@if( .NOT.  (DESC = ""REGULARIZACION ITF""); @IF(HORA<>"""";@LEFT(@ALLTRIM(@SPLIT(DESC;HORA;"""";1;0));6);"""");"""")"
	field.Length = 6
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anexar campo
Function AppendField5
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "LUGAR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Remove(@JUSTLETTERS(@SPLIT(DESC;MED;"""";1;0));""-"")"
	field.Length = 18
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anexar campo
Function AppendField6
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "PRUEBA1"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@SpacesToOne(@ALLTRIM(@Split( DESC ; DESCRIPCION ;"""";1;0)))"
	field.Length = 40
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anexar campo
Function AppendField7
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "SUC_AGE"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@IF(@MID(PRUEBA1; 5;7)= ""-""; """"; @IF(@MID(PRUEBA1; 5;7)= ""AG"" .OR. @MID(PRUEBA1; 5;7)= ""SUC"";  (@LEFT(@ALLTRIM(@SPLIT(PRUEBA1;LUGAR;"""";1;0));7);@MID(PRUEBA1; 5;7))))"
	field.Length = 7
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anexar campo
Function AppendField8
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NUM_OPE"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@IF(SUC_AGE= """"; """"; @ALLTRIM(@SPLIT(DESC; SUC_AGE; """"; 1;0)))"
	field.Length = 6
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anexar campo
Function AppendField9
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CARGO_ABONO_T"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@split(desc;"""";"" "";1;1)"
	field.Length = 18
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anexar campo
Function AppendField10
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CARGO_ABONO"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = "@IF(@RIGHT(CARGO_ABONO_T; 1) = ""-""; @VAL(@REMOVE(@Remove(CARGO_ABONO_T;""."");"",""))/-100.00 ;@VAL(@Remove(CARGO_ABONO_T;"","")))"
	field.Decimals = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anexar campo
Function AppendField11
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIPO"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@RIGHT(@ALLTRIM(@SPLIT(DESC;CARGO_ABONO_T;"""";1;1));4)"
	field.Length = 4
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Datos: Extracción directa
Function DirectExtraction
	Set db = Client.OpenDatabase("BCP2022.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "PERIODO"
	task.AddFieldToInc "FECHA_PROC"
	task.AddFieldToInc "DESCRIPCION"
	task.AddFieldToInc "MED"
	task.AddFieldToInc "LUGAR"
	task.AddFieldToInc "SUC_AGE"
	task.AddFieldToInc "NUM_OPE"
	task.AddFieldToInc "HORA"
	task.AddFieldToInc "ORIGEN"
	task.AddFieldToInc "TIPO"
	task.AddFieldToInc "CARGO_ABONO"
	task.AddFieldToInc "SALDO"
	task.AddFieldToInc "CUENTA"
	task.AddFieldToInc "MONEDA"
	task.AddFieldToInc "TITULAR_CUENTA"
	dbName = "F_BCP2022.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Análisis: Resumen
Function Summarization
	Set db = Client.OpenDatabase("F_BCP2022.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "PERIODO"
	task.AddFieldToSummarize "CUENTA"
	task.AddFieldToInc "MONEDA"
	task.AddFieldToTotal "CARGO_ABONO"
	dbName = "F.1_Resumen_BCP.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function