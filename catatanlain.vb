'=================== copy paste file / simple
Private Sub copyPasteFile()
    newPath = "E:\korban7.txt"
    oldPath = "c:\korban.txt"
    FileCopy oldPath, newPath
    
End Sub

Private Sub renameFile()
    newPath = "E:\korban7.txt"
    oldPath = "c:\korban.txt"
    name oldPath as  newPath
    
End Sub
'==================== delete file / simple
Private Sub deleteFile(newPath)
    'newPath= "E:\korban7.txt"
    On Error Resume Next 'digunakan jika nanti file yg di delete mmng tidak ada, tdk terjadi error
    	Kill newPath     ' syntak inti
    On Error GoTo 0
   
End Sub

Private Sub deleteFolder()
	newPath = "C:\Users\Ron\Test\"
	On Error Resume Next 'digunakan jika nanti file yg di delete mmng tidak ada, tdk terjadi error
	RmDir newPath
	On Error GoTo 0
end sub


'=====================	create folder - vba mode
Dim fileObj As Object

Private Sub createFolder()
	pathFolder = ThisWorkbook.Path & "\Data Out CFGMML"
    Set fileObj = CreateObject("Scripting.FileSystemObject")
    If fileObj.FolderExists(pathFolder) = False Then fileObj.createfolder pathFolder
	
End Sub

'===================== check file exist
Function isFileExist(pathFileCheck)
    isFileExist = CreateObject("Scripting.FileSystemObject").FileExists(pathFileCheck)
End Function

'===================== copy paste standard
sub copyPasteStd
	CH2.Range(Selection, Selection.End(xlDown)).Select
	CH2.Range(Selection, Selection.End(xlToRight)).Select
	Application.CutCopyMode = False
	Selection.SpecialCells(xlCellTypeVisible).Select 	'<- khusus yang ada di layar yg akan di copy
	Selection.Copy
 
	SH2.Select
	SH2.Range("A2").Select
	ActiveSheet.Paste
	Application.CutCopyMode = False
end sub
'==================== copy paste value :p
sub copPasVal
	Application.CutCopyMode = False
    Range("n8:n9").Select
    Selection.Copy
    Range("j5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
	
	Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
end sub
'=================== open file clasic
Sub OpenFile()
    Dim fileName

    fileName = Application.GetOpenFilename("excell lama (*.xls;*.xlsx;*.csv;*.xlsm;*xlsb),*.xls;*.xlsx;*.csv;*.xlsm;*xlsb")

    If fileName <> "False" Then
        Range("B2").Value = fileName
        Workbooks.Open fileName, Format:=2
    End If
End Sub
'================== open file Another
Sub browse()
    nameFileMaster = ThisWorkbook.Name
    
    Application.Dialogs(xlDialogOpen).Show
    If ActiveWorkbook.Name = nameFileMaster Then
        MsgBox ("anda tidak membuka file input yo")
        Exit Sub
    End If
    
    objIn1 = ActiveWorkbook.Name
    
    Windows(nameFileMaster).Activate
End sub

'=============== browse single file, use useing call pathBrowse("B4")
Private Sub pathBrowse(cellAddress)
    Set targetCell = Application.ActiveSheet.Range(cellAddress)

    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Select a File"
        .Filters.Clear
        .Filters.Add "Excel or Text Files", "*.xls;*.xlsx;*.csv;*.txt"
        .AllowMultiSelect = False

        If .Show = -1 Then
            targetCell.Value = .SelectedItems(1)
        End If
    End With

    Exit Sub
End Sub
						
'=============== open file clasic multiple
Sub OpenFileMultiple() 
    Dim lngCount As Long
     
     ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Multiple select"
		.Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.csv;*.txt"
        .AllowMultiSelect = True
        '.InitialFileName = ThisWorkbook.Path & "\Output\WPC\"
        If .Show = -1 Then
             ' Open the files
            For lngCount = 1 To .SelectedItems.Count
                Range("A" & lngCount).Value = .SelectedItems(lngCount)
				'Workbooks.Open .SelectedItems(lngCount), Format:=2
            Next lngCount
        End If
         
    End With
     
End Sub

Sub AddOpenFileMultiple()
    Dim lngCount As Long
    columnSelect = "R"
    maxBrowse = Range(columnSelect & Range(columnSelect & ":" & columnSelect).Rows.Count).End(xlUp).Row

     ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .Title = "Multiple select"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.csv;*.txt"
        .AllowMultiSelect = True
        '.InitialFileName = ThisWorkbook.Path & "\Output\WPC\"
        If .Show = -1 Then
             ' Open the files
            For lngCount = 1 To .SelectedItems.Count
                Range(columnSelect & lngCount + maxBrowse).Value = .SelectedItems(lngCount)
                'Workbooks.Open .SelectedItems(lngCount), Format:=2
            Next lngCount
        End If
         
    End With
     
End Sub

Sub clearBrowse()
    columnSelect = "R"
    maxBrowse = Range(columnSelect & Range(columnSelect & ":" & columnSelect).Rows.Count).End(xlUp).Row
    
    If maxBrowse > 1 Then
    Range(columnSelect & "2:" & columnSelect & maxBrowse).ClearContents
	end if
End Sub

'=============== open folder clasic multiple or not multiple
Sub OpenFolderMultiple() 
    Dim lngCount As Long
     
     ' Open the file dialog
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder"
        .AllowMultiSelect = false
        '.InitialFileName = strPath   '=== ini digunakan untuk spesific path when starting??
        If .Show = -1 Then
             ' Open the files
            For lngCount = 1 To .SelectedItems.Count
                Range("A" & lngCount).Value = .SelectedItems(lngCount)
            Next lngCount
        End If
         
    End With
     
End Sub

'======== loopChart
    Dim myChart As ChartObject
    For Each myChart In Sheets(ActiveSheet.Name).ChartObjects
 
        myChart.Select
        ActiveSheet.Shapes(myChart.Name).Width = 432
        ActiveSheet.Shapes(myChart.Name).Height = 216
    Next myChart
	
'======== remove filter
                If Sheets(ActiveSheet.Name).AutoFilterMode = True Then
                        Range("A1").Select
                        Selection.AutoFilter
                        
                End If




'================= function convert
Sub maxColumn()
        cellSelect = "A1"
        columnMax = Range(cellSelect).End(xlToRight).Column
        columnMaxLng = cnvrtColumn
        columnMaxStr = cnvrtColumn(columnMax + 0)
End Sub

Function cnvrtColumn(inC_F As long) As String ' from Int To Str
    Dim outC_Fpart1 As long
    Dim outC_Fpart2 As long
    
    outC_Fpart1 = ((inC_F - 1) / 26) - (((((inC_F - 1) / 26) * 100) Mod 100) / 100)
    
    outC_Fpart2 = inC_F Mod 26
    If outC_Fpart2 = 0 Then outC_Fpart2 = 26
    
    If outC_Fpart1 < 1 Then
        outC_F = (Chr(64 + outC_Fpart2))
    Else
        outC_F = (Chr(64 + outC_Fpart1) & Chr(64 + outC_Fpart2))
    End If
    
    cnvrtColumn = outC_F
    
End Function

Function cnvrtColumn(lngCol As Long) As String
	Dim vArr
	vArr = Split(Cells(1, lngCol).Address(True, False), "$")
	cnvrtColumn = vArr(0)
End Function

'================= collect filtering data uniqe
sub collectFilteringData()
	Columns( "A:A" ).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ShChart.Range("H1"), Unique:=True
	
	'collect filtering data, and put the result in H1
End sub

'================ tentang object
nameFileMaster = ThisWorkbook.Name 'file master
nameObj = ActiveWorkbook.Name 'initial objek
Windows(namaObj).Activate 'memilih windows yg aktif
Windows(namaObj).Close 'close objek
Windows(namaObj).save
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & yourfilename
Workbooks.Add ' new excel
Application.ScreenUpdating = False 'view

ActiveWorkbook.Save
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & yourfilenameWithExt , fileformat:=51 
ActiveWorkbook.SaveAs "C:\ron.xlsm", fileformat:=52 
note:
51 = xlOpenXMLWorkbook (without macro's in 2007-2010, xlsx)
52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2010, xlsm)
50 = xlExcel12 (Excel Binary Workbook in 2007-2010 with or without macro's, xlsb)
56 = xlExcel8 (97-2003 format in Excel 2007-2010, xls)
6 =  ".csv"
-4158 = ".txt"
36 = ".prn"

'====== get code ext
Function getCodeExtFile(strFile)
    
    jFnc = 0
    For iFnc = 1 To 20000
        If Mid(strFile, iFnc, 1) = "" Then Exit For
        If Mid(strFile, iFnc, 1) = "." Then
                jFnc = iFnc
        End If
    Next
    getCodeExtFile = Right(strFile, Len(strFile) - jFnc)
    
    If LCase(getCodeExtFile) = "xlsx" Then getCodeExtFile = 51
    If LCase(getCodeExtFile) = "xlsm" Then getCodeExtFile = 52
    If LCase(getCodeExtFile) = "xlsb" Then getCodeExtFile = 50
    If LCase(getCodeExtFile) = "xls" Then getCodeExtFile = 56
    If LCase(getCodeExtFile) = "csv" Then getCodeExtFile = 6
    If LCase(getCodeExtFile) = "txt" Then getCodeExtFile = -4158
    If LCase(getCodeExtFile) = "prn" Then getCodeExtFile = 36
    
End Function

'====== save file
    Application.Dialogs(xlDialogSaveAs).Show
	
'===== very last row
lastRow = Range("E" & Range("E:E").Rows.Count).End(xlUp).Row

'-- or
LastRow = ActiveSheet.UsedRange.Rows.Count

'===== very last column 'lastCell in long
With Range("A1").EntireRow
    LastCell = .Cells(1, .Columns.Count).End(xlToLeft).Column
End With

Dim lastColumn As Long
lastColumn = Sheet1.Cells(1, Columns.Count).End(xlToLeft).Column

'===== duplicate formula for one row
columnSelect = "A"
maxRow = Range(columnSelect & Range(columnSelect & ":" & columnSelect).Rows.Count).End(xlUp).Row

maxRow = Range("A" & Range("A:A").Rows.Count).End(xlUp).Row
Range("C2").Formula = "=A2 + B2"
Range("C2").AutoFill Destination:=Range("C2:C" & maxRow)

Range("C2:C" & maxRow).copy
Range("C2").select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Application.CutCopyMode = False


'===== dir??


strFile = Dir(ThisWorkbook.Path & "\")
msgbox(strfile)
for i = 1 to 20000
	If strFile = "" Then Exit For
	if left(strFile,12) = "file aku.txt" then
		'bla bla bla
	end if	
	strFile = Dir    ' Get next entry.
next

'====== ambil nama file
Function ambilTanpaExt(strFile) As String
    
    jFnc = 0
    For iFnc = 1 To 20000
        If Mid(strFile, iFnc, 1) = "" Then Exit For
        If Mid(strFile, iFnc, 1) = "." Then
                jFnc = iFnc
        End If
    Next
    ambilTanpaExt = Left(strFile, jFnc - 1)
	'in "abc.xls"
	'out "abc"
End Function

Function ambilNameFile(strFile As String) As String
    jFunc = 0
    For iFunc = 1 To 20000
        If Mid(strFile, iFunc, 1) = "" Then Exit For
        If Mid(strFile, iFunc, 1) = "\" Then
                jFunc = iFunc
        End If
        
    Next
    
    ambilNameFile = Right(strFile, Len(strFile) - jFunc)
	
	'in "C:\aa\abc.xls"
	'out "abc.xls"
End Function


Function ambilPathFolderFile(strFile As String) As String
	jFunc = 0
	For iFunc = 1 To 20000
		If Mid(strFile, iFunc, 1) = "" Then Exit For
		If Mid(strFile, iFunc, 1) = "\" Then
				jFunc = iFunc
		End If
		
	Next
	ambilPathFolderFile = Left(strFile, jFunc)
	'sampleOut: "F:\NDS\tools\raw sample\"
End Function

Function ambilExtFileOnly(strFile) As String
    
    jFnc = 0
    For iFnc = 1 To 20000
        If Mid(strFile, iFnc, 1) = "" Then Exit For
        If Mid(strFile, iFnc, 1) = "." Then
                jFnc = iFnc
        End If
    Next
    ambilExtFileOnly = Right(strFile, Len(strFile) - jFnc)
End Function


'======= select multiple shape
 Sub select2Shape()
    
    Set myrange = ActiveSheet.Shapes.Range(Array("Chart 2", _
    "Chart 6"))
    myrange.Select

 End Sub
 
 '===== array
Dim DynamicArray() As String
ReDim DynamicArray(1 To 10)

 '=====   delete sheet unused
    For iDSU = Sheets.Count To 1 Step -1
        
        If LCase(Left(Sheets(iDSU).Name, 5)) = "sheet" Then
            DeleteSheet (Sheets(iDSU).Name)
        End If
    Next
'====== hide prompt mesage delete sheet
Sub DeleteSheet(strSheetName As String)
' deletes a sheet named strSheetName in the active workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets(strSheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub


'====== add sheet and name it
Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "My New Worksheet" 

'====== name sheet now
ActiveSheet.Name

'===== rename sheet now
Sheets(ActiveSheet.Name).Name = "sheet apa aja boleh"
'===== pause calculate
Application.Calculation = xlManual 
 ' code here
Calculate 'calculate untuk tiap data yang belum beres di calculate
Application.Calculation = xlAutomatic 

Function jelekArraySplit(dataArraySplit As String, delimArraySplit As String) As Variant
    tmpFncArrayJAS = Split(dataArraySplit, delimArraySplit)

    Dim outJelekArraySplit()
    ReDim outJelekArraySplit(0 To 10000)
    jFncJAS = 0
    For iFncJAS = 0 To Application.CountA(tmpFncArrayJAS) - 1
        If Len(tmpFncArrayJAS(iFncJAS)) > 0 Then
            outJelekArraySplit(jFncJAS) = tmpFncArrayJAS(iFncJAS)
            jFncJAS = jFncJAS + 1
        End If
    Next
    
    Dim outFinalJelekArraySplit()
    ReDim outFinalJelekArraySplit(0 To jFncJAS - 1)
    For iFncJAS = 0 To jFncJAS - 1
        outFinalJelekArraySplit(iFncJAS) = outJelekArraySplit(iFncJAS)
    Next
    
    tmpFncArrayJAS = Null
    ReDim outJelekArraySplit(0 To 0)
    jelekArraySplit = outFinalJelekArraySplit
End Function



'====== running .exe file
Call Shell("C:\Program Files\apapun\delExcellProcess.exe", vbNormalFocus)


SiteID: IIf(InStr(1,[cellname],"_")<7 , Left([cellname],(InStr(1,[cellname],"_"))-1),Left([cellname],6))
NodeBName: Left([cellname],Len([cellname])-2)

'====== error handling with err number
On Error Resume Next
    N = 1 / 0    ' cause an error
    If Err.Number <> 0 Then
        N = 1
    End If
on error goto 0

'====== loop sheet
for i = 1 to sheets.count
sheetNow =sheets(i).name
	if sheets(sheetNow).visible = -1 then
		sheets(sheetNow).select
	end if
next

'====== check sheet exist
Function checkSheetExist(nameSheetSearch As String) As Boolean

    checkSheetExist = False
    For iCSE = 1 To Sheets.Count
        If Sheets(iCSE).Name = nameSheetSearch Then
            checkSheetExist = True
            Exit For
        End If
    Next
End Function


'====== collect all workbook
  For i = 1 To Workbooks.Count
      Range("A" & i).Value = Workbooks(i).Name
  Next

'====== convert to number and string
For Each rng In Selection.Cells
rng.value = rng.Value
Next rng
  
  
'====== call macro different excell
Dim CellXLS As excel.Application
Set CellXLS = CreateObject("Excel.Application")
CellXLS.Workbooks.Open "D:\HCPT_Reporting\Macros\ReportGeneratorMacro.xlsm"
CellXLS.Visible = True
CellXLS.Workbooks("ReportGeneratorMacro.xlsm").Activate
'CellXLS.ActiveWorkbook.Application.Run "GenerateRNCPerformance_PK"
CellXLS.ActiveWorkbook.Application.Run "generateReport_Failure_RNC_PK"
'CellXLS.ActiveWorkbook.Application.Run "generateReport_Failure_RNC_BH_PK"
CellXLS.ActiveWorkbook.Application.Run "generate_WPC_Daily_RNC_PK"
'CellXLS.ActiveWorkbook.Application.Run "generate_WPC_Daily_RNC_BH_PK"
CellXLS.ActiveWorkbook.Application.Run "generate_CQI_Report_RNC_PK"
CellXLS.ActiveWorkbook.Application.Run "ReportGeneratorMacro.xlsm!zeroTrafficPk"

CellXLS.Workbooks("ReportGeneratorMacro.xlsm").Close False

'============ call macro different excell
nameTemplate = "data master.xls"
macroName = "copyPasteAuto"
Application.Run ("'" & nameTemplate & "'!" & (macroName & ""))

bez!J-sh1nt4

'----
For i = 1 To Cells(Rows.Count,1).End(xlup).Row Step 1000 
    With .Cells(i, 1) 
        With .Resize(1000, 60).SpecialCells(xlCellTypeBlanks) 
            .Value = 0 
        End With 
    End With 
Next i 

'-======
=LEFT(ADDRESS(ROW(),1,4),LEN(ADDRESS(ROW(),1,4))-1)
=LEFT(ADDRESS(1,COLUMN(),4),LEN(ADDRESS(1,COLUMN(),4))-1)
=MIN(FIND({0,1,2,3,4,5,6,7,8,9},A15&"0123456789"))
'============
k = 2
    j = 14
    For i = 1 To 13
    
        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.SeriesCollection(i).Name = "='2G_SalesCluster'!$A$" & k + (i - 1) * 13
        ActiveChart.SeriesCollection(i).Values = "='2G_SalesCluster'!$C$" & j + (i - 1) * 13 & ":$BL$" & j + (i - 1) * 13
        ActiveChart.SeriesCollection(i).Select
        With Selection.Format.Line
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
        ActiveChart.SeriesCollection(i).XValues = "='2G_SalesCluster'!$C$1:$BL$1"
    
    Next
	
'-------
and =  A * B
not = - ( X - 1)
or  = -( A * B ) + A + B

'=========== cari find
			Rows("6:6").Select
            Set cari = Selection.Find(What:=tanggalText, After:=ActiveCell, LookIn:=xlFormulas _
                , LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
            
            If cari Is Nothing Then MsgBox "tanggal tidak ditemukan": Exit Sub
            
            cari.Activate
            Range("E2").Select
            Range(Selection, Selection.Offset(5000, cari.Column - 6)).ClearContents
            
            cari.Offset(-4, 1).Select
            Range(Selection, Selection.Offset(5000, (255 - cari.Column))).ClearContents
            cellArray(I) = cari.Address

'------------- visible
Application.ScreenUpdating = False

'=============== access

sub refreshPivot()
	ActiveWorkbook.RefreshAll
end sub

Sub timpaLinkChart()

    Dim oSl As Slide
    Dim oHl As Hyperlink
    Dim sSearchFor As String
    Dim sReplaceWith As String
    Dim oSh As Shape


'    If sSearchFor = "" Then
'        Exit Sub
'    End If

    sReplaceWith = "F:\Inspur\Meldi\Generic\raw_data_vas_performanceNew.xlsm"
    If sReplaceWith = "" Then
        Exit Sub
    End If

    
    For Each oSl In ActivePresentation.Slides
    
        For Each oHl In oSl.Hyperlinks
            oHl.Address = sReplaceWith
            oHl.SubAddress = sReplaceWith
        Next    ' hyperlink
        
        For Each oSh In oSl.Shapes
            If oSh.Type = msoLinkedOLEObject Or oSh.Type = msoMedia Or oSh.Type = 3 Then
                oSh.LinkFormat.SourceFullName = sReplaceWith
            End If
        Next

    Next    ' slide

End Sub

