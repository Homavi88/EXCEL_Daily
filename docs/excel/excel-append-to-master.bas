Attribute VB_Name = "modAppendToMaster"
'==============================================================================
' Append monthly ListObjects into one master workbook (for Power BI).
' Paste this module into the monthly SOURCE workbook (.xlsm).
' Edit MASTER_PATH below before first run.
'==============================================================================
Option Explicit

' --- Edit this path (master is .xlsx, no macros needed in master) ---
Private Const MASTER_PATH As String = "C:\Data\Lab_Master.xlsx"

' Sheet name = ListObject name
Private Const SHEET_LIST As String = "Stem_Flakes|DIET|CF|Kitchen|Ripper I|Other TI"
Private Const OTHER_TI As String = "Other TI"

'------------------------------------------------------------------------------
' Entry point: run from Macros dialog or assign to a button.
'------------------------------------------------------------------------------
Public Sub AppendTablesToMaster()
    Dim srcWb As Workbook
    Dim mstWb As Workbook
    Dim sheetNames() As String
    Dim i As Long
    Dim added As Long
    Dim skipped As Long
    Dim totalAdded As Long
    Dim totalSkipped As Long
    Dim report As String
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation
    Dim createdMaster As Boolean
    Dim errMsg As String
    Dim softFail As Boolean
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim whereProc As String

    softFail = False
    whereProc = "AppendTablesToMaster"
    errMsg = vbNullString

    Set srcWb = ActiveWorkbook
    If srcWb Is Nothing Then
        MsgBox "Нет активной книги-источника.", vbExclamation, "AppendToMaster"
        Exit Sub
    End If

    If Len(Trim$(MASTER_PATH)) = 0 Then
        MsgBox "Задайте константу MASTER_PATH в модуле modAppendToMaster.", _
               vbExclamation, "AppendToMaster"
        Exit Sub
    End If

    sheetNames = Split(SHEET_LIST, "|")

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Open or create master
    If FileExistsSafe(MASTER_PATH, errMsg) Then
        whereProc = "AppendTablesToMaster / Workbooks.Open"
        Set mstWb = Workbooks.Open(Filename:=MASTER_PATH, ReadOnly:=False)
        createdMaster = False
    Else
        If Len(errMsg) > 0 Then
            softFail = True
            GoTo CleanFail
        End If
        whereProc = "AppendTablesToMaster / CreateMasterWorkbook"
        Set mstWb = CreateMasterWorkbook(srcWb, sheetNames, errMsg)
        If mstWb Is Nothing Then
            ' Soft validation failure -- errMsg already set; do NOT use Err (Number=0)
            If Len(errMsg) = 0 Then
                errMsg = "Не удалось создать мастер-книгу."
            End If
            softFail = True
            GoTo CleanFail
        End If
        createdMaster = True
    End If

    report = IIf(createdMaster, "Создан новый мастер: " & MASTER_PATH & vbCrLf & vbCrLf, "")
    report = report & "Источник: " & srcWb.Name & vbCrLf & vbCrLf

    totalAdded = 0
    totalSkipped = 0

    For i = LBound(sheetNames) To UBound(sheetNames)
        added = 0
        skipped = 0
        errMsg = vbNullString
        whereProc = "AppendOneTable / " & Trim$(sheetNames(i))

        If Not AppendOneTable(srcWb, mstWb, Trim$(sheetNames(i)), added, skipped, errMsg) Then
            softFail = True
            If Len(errMsg) = 0 Then
                errMsg = "Не удалось обработать лист """ & Trim$(sheetNames(i)) & """."
            End If
            GoTo CleanFail
        End If

        totalAdded = totalAdded + added
        totalSkipped = totalSkipped + skipped
        report = report & Trim$(sheetNames(i)) & ": добавлено " & CStr(added) & _
                 ", пропущено (дубликаты) " & CStr(skipped) & vbCrLf
    Next i

    whereProc = "AppendTablesToMaster / Save"
    Application.DisplayAlerts = False
    If createdMaster Then
        mstWb.SaveAs Filename:=MASTER_PATH, FileFormat:=xlOpenXMLWorkbook ' .xlsx
    Else
        mstWb.Save
    End If
    Application.DisplayAlerts = True

    mstWb.Close SaveChanges:=False
    Set mstWb = Nothing

    report = report & vbCrLf & "Итого: добавлено " & CStr(totalAdded) & _
             ", пропущено " & CStr(totalSkipped) & vbCrLf & _
             "Мастер сохранён: " & MASTER_PATH

    MsgBox report, vbInformation, "AppendToMaster"
    GoTo CleanExit

CleanFail:
    ' Capture Err BEFORE any further On Error / Resume Next (clears Err)
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    If softFail Then
        ' Keep intentional errMsg; never replace with blank «Ошибка 0»
        If Len(Trim$(errMsg)) = 0 Then
            errMsg = FormatRuntimeError(errNum, errDesc, errSrc, whereProc)
        Else
            errMsg = errMsg & vbCrLf & vbCrLf & "Где: " & whereProc
        End If
    Else
        errMsg = FormatRuntimeError(errNum, errDesc, errSrc, whereProc)
    End If

    On Error Resume Next
    If Not mstWb Is Nothing Then
        mstWb.Close SaveChanges:=False
        Set mstWb = Nothing
    End If
    On Error GoTo 0

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    Application.DisplayAlerts = True

    MsgBox errMsg, IIf(softFail, vbExclamation, vbCritical), "AppendToMaster"
    Exit Sub

CleanExit:
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    Application.DisplayAlerts = True
End Sub

'------------------------------------------------------------------------------
' Build a non-blank error message even when Err.Number = 0.
'------------------------------------------------------------------------------
Private Function FormatRuntimeError(ByVal errNum As Long, _
                                    ByVal errDesc As String, _
                                    ByVal errSrc As String, _
                                    ByVal whereProc As String) As String
    Dim parts As String
    Dim desc As String

    desc = Trim$(errDesc)
    If Len(desc) = 0 Then
        If errNum = 0 Then
            desc = "(нет текста Err.Description -- сбой без runtime-ошибки VBA)"
        Else
            desc = "(пустое описание)"
        End If
    End If

    parts = "Ошибка " & CStr(errNum) & ": " & desc
    If Len(Trim$(errSrc)) > 0 Then
        parts = parts & vbCrLf & "Источник Err: " & errSrc
    End If
    If Len(Trim$(whereProc)) > 0 Then
        parts = parts & vbCrLf & "Процедура: " & whereProc
    End If

    FormatRuntimeError = parts
End Function

'------------------------------------------------------------------------------
' Safe file exists via Dir; captures Dir failures (network / bad path).
' Returns False and sets errMsg on Dir error; False + empty errMsg if missing.
'------------------------------------------------------------------------------
Private Function FileExistsSafe(ByVal filePath As String, ByRef errMsg As String) As Boolean
    Dim n As String
    Dim errNum As Long
    Dim errDesc As String

    errMsg = vbNullString
    On Error Resume Next
    n = Dir(filePath)
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0

    If errNum <> 0 Then
        errMsg = "Не удалось проверить файл мастера (Dir):" & vbCrLf & _
                 filePath & vbCrLf & vbCrLf & _
                 FormatRuntimeError(errNum, errDesc, vbNullString, "FileExistsSafe")
        FileExistsSafe = False
        Exit Function
    End If

    FileExistsSafe = (Len(n) > 0)
End Function

Private Function FolderExistsSafe(ByVal folderPath As String, ByRef errMsg As String) As Boolean
    Dim n As String
    Dim errNum As Long
    Dim errDesc As String
    Dim probe As String

    errMsg = vbNullString
    If Len(folderPath) = 0 Then
        FolderExistsSafe = False
        Exit Function
    End If

    ' Dir(path, vbDirectory) is flaky on some network paths; also try trailing slash
    probe = folderPath
    If Right$(probe, 1) <> "\" Then probe = probe & "\"

    On Error Resume Next
    n = Dir(probe, vbDirectory)
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0

    If errNum <> 0 Then
        errMsg = "Не удалось проверить папку мастера (Dir):" & vbCrLf & _
                 folderPath & vbCrLf & vbCrLf & _
                 FormatRuntimeError(errNum, errDesc, vbNullString, "FolderExistsSafe")
        FolderExistsSafe = False
        Exit Function
    End If

    FolderExistsSafe = (Len(n) > 0)
End Function

'------------------------------------------------------------------------------
' Create empty master with same sheet/table names and headers from source.
'------------------------------------------------------------------------------
Private Function CreateMasterWorkbook(ByVal srcWb As Workbook, _
                                      ByRef sheetNames() As String, _
                                      ByRef errMsg As String) As Workbook
    Dim mstWb As Workbook
    Dim i As Long
    Dim nm As String
    Dim srcLo As ListObject
    Dim ws As Worksheet
    Dim destLo As ListObject
    Dim hdr As Range
    Dim c As Long
    Dim folderPath As String
    Dim folderErr As String

    errMsg = vbNullString
    Set CreateMasterWorkbook = Nothing

    ' Ensure parent folder exists
    folderPath = Left$(MASTER_PATH, InStrRev(MASTER_PATH, "\") - 1)
    If Len(folderPath) > 0 Then
        If Not FolderExistsSafe(folderPath, folderErr) Then
            If Len(folderErr) > 0 Then
                errMsg = folderErr
            Else
                errMsg = "Папка для мастера не найдена:" & vbCrLf & folderPath
            End If
            Exit Function
        End If
    End If

    ' Validate all source tables exist before creating anything on disk
    For i = LBound(sheetNames) To UBound(sheetNames)
        nm = Trim$(sheetNames(i))
        If Not SheetExists(srcWb, nm) Then
            errMsg = "В источнике нет листа: """ & nm & """"
            Exit Function
        End If
        Set srcLo = FindListObject(srcWb.Worksheets(nm), nm)
        If srcLo Is Nothing Then
            errMsg = "На листе """ & nm & """ нет таблицы (ListObject) с именем """ & nm & """."
            Exit Function
        End If
        If Not ValidateKeyColumns(srcLo, nm, errMsg) Then
            Exit Function
        End If
    Next i

    Set mstWb = Workbooks.Add(xlWBATWorksheet)

    ' Remove extra default sheets, keep one to rename
    Application.DisplayAlerts = False
    Do While mstWb.Worksheets.Count > 1
        mstWb.Worksheets(mstWb.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True

    For i = LBound(sheetNames) To UBound(sheetNames)
        nm = Trim$(sheetNames(i))
        Set srcLo = FindListObject(srcWb.Worksheets(nm), nm)

        If i = LBound(sheetNames) Then
            Set ws = mstWb.Worksheets(1)
            ws.Name = nm
        Else
            Set ws = mstWb.Worksheets.Add(After:=mstWb.Worksheets(mstWb.Worksheets.Count))
            ws.Name = nm
        End If

        ' Copy header row only
        Set hdr = srcLo.HeaderRowRange
        For c = 1 To hdr.Columns.Count
            ws.Cells(1, c).Value = hdr.Cells(1, c).Value
        Next c

        Set destLo = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range(ws.Cells(1, 1), ws.Cells(1, hdr.Columns.Count)), _
            XlListObjectHasHeaders:=xlYes)
        destLo.Name = nm
        ' Excel often inserts one blank body row -- remove it
        On Error Resume Next
        If Not destLo.DataBodyRange Is Nothing Then
            destLo.DataBodyRange.Delete
        End If
        destLo.TableStyle = srcLo.TableStyle
        On Error GoTo 0
    Next i

    Set CreateMasterWorkbook = mstWb
End Function

'------------------------------------------------------------------------------
' Append new rows from source table into master table. Returns False on hard error.
'------------------------------------------------------------------------------
Private Function AppendOneTable(ByVal srcWb As Workbook, _
                                ByVal mstWb As Workbook, _
                                ByVal nm As String, _
                                ByRef added As Long, _
                                ByRef skipped As Long, _
                                ByRef errMsg As String) As Boolean
    Dim srcWs As Worksheet
    Dim mstWs As Worksheet
    Dim srcLo As ListObject
    Dim mstLo As ListObject
    Dim keyCols() As Long
    Dim existing As Object ' Scripting.Dictionary
    Dim r As Long
    Dim c As Long
    Dim key As String
    Dim srcData As Range
    Dim mstData As Range
    Dim newRow As ListRow
    Dim colCount As Long
    Dim vals As Variant
    Dim errNum As Long
    Dim errDesc As String

    added = 0
    skipped = 0
    AppendOneTable = False
    errMsg = vbNullString

    If Not SheetExists(srcWb, nm) Then
        errMsg = "В источнике нет листа: """ & nm & """"
        Exit Function
    End If
    If Not SheetExists(mstWb, nm) Then
        errMsg = "В мастере нет листа: """ & nm & """" & vbCrLf & _
                 "Ожидаются те же имена листов/таблиц, что в источнике."
        Exit Function
    End If

    Set srcWs = srcWb.Worksheets(nm)
    Set mstWs = mstWb.Worksheets(nm)

    Set srcLo = FindListObject(srcWs, nm)
    If srcLo Is Nothing Then
        errMsg = "На листе источника """ & nm & """ нет таблицы """ & nm & """."
        Exit Function
    End If

    Set mstLo = FindListObject(mstWs, nm)
    If mstLo Is Nothing Then
        errMsg = "На листе мастера """ & nm & """ нет таблицы """ & nm & """."
        Exit Function
    End If

    If srcLo.ListColumns.Count <> mstLo.ListColumns.Count Then
        errMsg = "Лист """ & nm & """: число столбцов в источнике (" & _
                 CStr(srcLo.ListColumns.Count) & ") и мастере (" & _
                 CStr(mstLo.ListColumns.Count) & ") не совпадает."
        Exit Function
    End If

    ' Header names must match (stable columns)
    For c = 1 To srcLo.ListColumns.Count
        If StrComp(CStr(srcLo.ListColumns(c).Name), CStr(mstLo.ListColumns(c).Name), vbTextCompare) <> 0 Then
            errMsg = "Лист """ & nm & """: заголовки не совпадают в столбце " & CStr(c) & _
                     " (источник: """ & srcLo.ListColumns(c).Name & _
                     """, мастер: """ & mstLo.ListColumns(c).Name & """)."
            Exit Function
        End If
    Next c

    If Not ResolveKeyColumnIndexes(srcLo, nm, keyCols, errMsg) Then
        Exit Function
    End If
    ' Same columns must exist in master (already matched by header names)
    If Not ResolveKeyColumnIndexes(mstLo, nm, keyCols, errMsg) Then
        Exit Function
    End If

    On Error Resume Next
    Set existing = CreateObject("Scripting.Dictionary")
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If existing Is Nothing Or errNum <> 0 Then
        errMsg = "Не удалось создать Scripting.Dictionary." & vbCrLf & _
                 "Проверьте, что на ПК доступен Microsoft Scripting Runtime / scrrun.dll." & vbCrLf & vbCrLf & _
                 FormatRuntimeError(errNum, errDesc, vbNullString, "AppendOneTable / CreateObject Dictionary")
        Exit Function
    End If
    existing.CompareMode = 1 ' TextCompare

    ' Index existing master keys
    If Not mstLo.DataBodyRange Is Nothing Then
        Set mstData = mstLo.DataBodyRange
        For r = 1 To mstData.Rows.Count
            key = BuildKeyFromRow(mstData, r, keyCols)
            If Len(key) > 0 Then
                If Not existing.Exists(key) Then existing.Add key, True
            End If
        Next r
    End If

    If srcLo.DataBodyRange Is Nothing Then
        AppendOneTable = True
        Exit Function
    End If

    Set srcData = srcLo.DataBodyRange
    colCount = srcLo.ListColumns.Count

    For r = 1 To srcData.Rows.Count
        key = BuildKeyFromRow(srcData, r, keyCols)
        If Len(key) = 0 Then
            ' Empty key: skip without counting as duplicate
            skipped = skipped + 1
        ElseIf existing.Exists(key) Then
            skipped = skipped + 1
        Else
            Set newRow = mstLo.ListRows.Add(AlwaysInsert:=True)
            vals = srcData.Rows(r).Value
            If colCount = 1 Then
                newRow.Range.Cells(1, 1).Value = vals
            Else
                newRow.Range.Value = vals
            End If
            existing.Add key, True
            added = added + 1
        End If
    Next r

    AppendOneTable = True
End Function

'------------------------------------------------------------------------------
' Key columns: Batch; for Other TI: Date + Cell + Point
'------------------------------------------------------------------------------
Private Function ResolveKeyColumnIndexes(ByVal lo As ListObject, _
                                         ByVal sheetName As String, _
                                         ByRef keyCols() As Long, _
                                         ByRef errMsg As String) As Boolean
    Dim names() As String
    Dim i As Long
    Dim col As ListColumn

    If StrComp(sheetName, OTHER_TI, vbTextCompare) = 0 Then
        names = Split("Date|Cell|Point", "|")
    Else
        names = Split("Batch", "|")
    End If

    ReDim keyCols(LBound(names) To UBound(names))

    For i = LBound(names) To UBound(names)
        Set col = Nothing
        On Error Resume Next
        Set col = lo.ListColumns(names(i))
        On Error GoTo 0
        If col Is Nothing Then
            errMsg = "На листе """ & sheetName & """ нет ключевого столбца """ & names(i) & """."
            ResolveKeyColumnIndexes = False
            Exit Function
        End If
        keyCols(i) = col.Index
    Next i

    ResolveKeyColumnIndexes = True
End Function

Private Function ValidateKeyColumns(ByVal lo As ListObject, _
                                    ByVal sheetName As String, _
                                    ByRef errMsg As String) As Boolean
    Dim dummy() As Long
    ValidateKeyColumns = ResolveKeyColumnIndexes(lo, sheetName, dummy, errMsg)
End Function

Private Function BuildKeyFromRow(ByVal dataRng As Range, _
                                 ByVal rowIndex As Long, _
                                 ByRef keyCols() As Long) As String
    Dim i As Long
    Dim parts() As String
    Dim v As Variant
    Dim s As String
    Dim allEmpty As Boolean

    ReDim parts(LBound(keyCols) To UBound(keyCols))
    allEmpty = True

    For i = LBound(keyCols) To UBound(keyCols)
        v = dataRng.Cells(rowIndex, keyCols(i)).Value
        If IsError(v) Then
            s = "#ERR"
        ElseIf IsDate(v) And VarType(v) <> vbString Then
            ' Normalize dates to ISO-like string for stable keys
            s = Format$(CDate(v), "yyyy-mm-dd")
        ElseIf IsEmpty(v) Then
            s = vbNullString
        Else
            s = Trim$(CStr(v))
        End If
        If Len(s) > 0 Then allEmpty = False
        parts(i) = LCase$(s)
    Next i

    If allEmpty Then
        BuildKeyFromRow = vbNullString
    Else
        BuildKeyFromRow = Join(parts, Chr$(1))
    End If
End Function

Private Function FindListObject(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
            Set FindListObject = lo
            Exit Function
        End If
    Next lo
    Set FindListObject = Nothing
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
