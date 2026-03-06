Option Explicit

'================================================================================
'
' This is a Excel 2000 VBA macro written by Rick Davin, 9/26/2003.
' Updated to Excel 2002.
' Added Office.FileDialog call.
'
' Dependencies:
'         you *** MUST *** reference (tools, references)
'
'                   Microsoft Scripting Runtime
'
' The two main subroutines that control the conversion process are:
'          Fox2Excel
'          Transform_FOB
'
'--------------------------------------------------------------------------------
'IMPORTANT QUESTION: What is a FOB?
'
'FOB is my shorthand for Fox Object Block.  We are processing a plain text file
'with lots of single lines of input records.  A FOB (Fox Object Block) is a block
'of input records that begin with "NAME " as the first record of the block and
'end with "END" as the last record in a block.  There may be a varying number of
'intermediate records in-between the NAME and END.
'
'With the exception of END, which contains no other characters on that record
'line, all previous records in the block will be of the format:
'
'             AttributeName  =  AttributeValue
'
'There may be 1 or more blanks padding the equal sign (=) but the equal sign
'is an important delimiter to split out the AttributeName, which becomes the
'column header, and the Attribute Value, which becomes the cell value under
'the AttributeName's column.  All we do with each input line is simple string
'parsing.
'
'Note there may be some preceding or trailing blanks on an input line.  This
'is presumably for readability by humans but for our purposes here all input
'lines will have a TRIM function applied.  A TRIM is also done each time we
'parse the attribute name and value.
'
'The first record of a FOB will be the NAME attribute.  The second record is
'usually the TYPE (e.g. AIN, AOUT, PIDA) but the code here checks for attribute
'of TYPE instead of assuming it's always record #2 or array element 2.
'
'The idea of the code is to read down an input file several records at a time
'to subsequently process an indeterminant number of records as a single FOB.
'For that singular FOB, we then will write it to an Excel sheet with the same
'name as the TYPE.  That single FOB will be a single row on that sheet with
'each AttributeName becoming a different column.  The AttributeValue will then
'be written to that particular row and column on that particular sheet.
'
'================================================================================

Private mws_Index As Excel.Worksheet

Private Type FoxProperty
    Name As String
    Value As String
End Type
    
Private Type udtSheets
    SheetName As String
    Columns As Integer      'ranges from 1-256
    Rows As Long            'ranges from 1-65536
End Type

Private mSheetsArray() As udtSheets
Private mSheetsIndex As Integer
Private mSheetsCurNdx As Integer

Public Sub Fox2Excel()

    ' <<<---- Hardcoded Filepath.  Change for each new run.
    Dim MyFile As String  ' = "C:\Share\gotGDU.txt"
    
    Dim vStartTime As Date
    Dim fso As FileSystemObject
    Dim ts As TextStream
    Dim TAttributes() As FoxProperty
    Dim szDefaultPath As String
    
    szDefaultPath = ActiveWorkbook.Path
    If Mid(szDefaultPath, Len(szDefaultPath), 1) <> "\" Then
        szDefaultPath = szDefaultPath & "\"
    End If
    
    MyFile = BrowseFileName(szDefaultPath)
    
    If MyFile = "" Then Exit Sub
        
    Application.Cursor = xlWait

    vStartTime = Now()
    
    ReDim mSheetsArray(0 To 100)
    mSheetsIndex = 0
    
    PrepWork
    Set mws_Index = Worksheets(1)

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(MyFile)
    
    DoEvents
    
    Do While (ts.AtEndOfStream = False)
    
        'Read as many lines of text that define a Fox Object
        Get_FOB ts, TAttributes()
        
        'Now that we have all the lines related to a single Fox Object
        'we can transform that object to a single row with many columns
        Transform_FOB TAttributes()
        
        'Index select columns from the current object
        Index_FOB TAttributes()
        
        'we are now done with the current Fox Object.
        'We can now delete the string array.
        Erase TAttributes
    
        DoEvents
        
    Loop
    
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    Sort_Sheets  'sort collection of sheet names alphabetically
    
    Sort_Each_Sheets 'sort rows by column 1 (NAME) on each individual sheet
    
    Post_Parse   'try to parse out Compound, Block, and Unit
    
    mws_Index.Move Before:=Worksheets(1)
    Set mws_Index = Nothing
    
    SheetStats   'create a new sheet with statistics
    
    Application.Cursor = xlDefault

    MsgBox "Finished!" & vbCrLf & vbCrLf & "Elapsed processing time (secs): " & DateDiff("s", vStartTime, Now) & _
           vbCrLf & vbCrLf & "This workbook contains " & Worksheets.Count & " worksheets."
    
End Sub


Private Sub Get_FOB(ByRef ts As TextStream, ByRef TAttributes() As FoxProperty)

    Dim i As Integer
    Dim szThisLine As String
    Const szStartToken = "NAME "  'very important: make sure a single BLANK follows NAME.
    Const szStopToken = "END"
    Dim x As Integer
    
   'Find the equal sign "=".  X marks the spot.
   'The attribute name is to the left of it (so subtract 1).
   'The value is to the right (so add 1).
   'Trim all results.
            
    ReDim TAttributes(1 To 256)
            
    'Be sure this current line starts with "NAME "  (be sure to pad a blank)
    'or else keep reading lines until you find a line that does
    Do While (ts.AtEndOfStream = False)
        szThisLine = Trim(ts.ReadLine)
        If Left(szThisLine, Len(szStartToken)) = szStartToken Then
            i = 1
            x = InStr(1, szThisLine, "=")
            TAttributes(i).Name = Trim(Left(szThisLine, x - 1))
            TAttributes(i).Value = Trim(Mid(szThisLine, x + 1))
            Exit Do
        End If
    Loop
    
    'to avoid later errors, we want to be sure string array is valid and
    'diminsioned but that it's upper bound is less than 1.
    'This is to protect for possible errors since we may later use code like
    '     e.g. For i = 1 to UBound(TAttributes)
    If i = 0 Then
        ReDim TAttributes(0 To 0)
        Exit Sub
    End If
    
    'now loop thru reading each line of this name block until you get
    ' to the a line containing "END" or else the end of text stream
    Do While (ts.AtEndOfStream = False) Or (szThisLine = szStopToken)
        szThisLine = Trim(ts.ReadLine)
        If szThisLine = szStopToken Then
            'if we read "END" then we can stop reading this Fox Object Block
            'we do not add "END" to the string array
            ReDim Preserve TAttributes(1 To i)
            Exit Do
        Else
            i = i + 1
            If i > UBound(TAttributes) Then
                'in case we exceed the current array's upper bound
                'increase in a sizeable chunk
                ReDim Preserve TAttributes(1 To (i + 100))
            End If
            x = InStr(1, szThisLine, "=")
            If x > 0 Then
                TAttributes(i).Name = Trim(Left(szThisLine, x - 1))
                TAttributes(i).Value = Trim(Mid(szThisLine, x + 1))
            Else
                TAttributes(i).Name = szThisLine
                TAttributes(i).Value = ""
            End If
        End If
    Loop
    
End Sub


Private Sub Transform_FOB(ByRef TAttributes() As FoxProperty)

    Dim i As Integer
    Dim szFoxType As String
    Dim ws As Worksheet
    Dim r As Long  'row
    
    'Each Fox object belongs to a specific TYPE (e.g. AIN, AOUT, PIDA, etc.)
    'We will write each unique object type to a worksheet of the same name.
    szFoxType = Get_FOB_Type(TAttributes())
    
    'This next statement returns sheet with current name.
    'If sheet with the new TYPE name does not exist, it will create it.
    'Think: Fox TYPE = Worksheet Name
    Set ws = Set_FoxType_Sheet(szFoxType)
    
    mSheetsCurNdx = Get_CurrentIndex(szFoxType)
    
    'Find the next available row on the selected sheet type
    r = Get_NextRow(mSheetsCurNdx)
    
    'Loop thru each input line
    For i = 1 To UBound(TAttributes)
        'for the next function, we pass a single line not the entire array
        Move_Line_To_Sheet ws, r, TAttributes(i)
    Next i

End Sub

Private Function Get_FOB_Type(ByRef TAttributes() As FoxProperty) As String

    Dim i As Integer
    
    'The Fox TYPE is demarked by the 2nd record (usually) but always will
    'begin with "TYPE".
    For i = 1 To UBound(TAttributes)
        If TAttributes(i).Name = "TYPE" Then
            Get_FOB_Type = TAttributes(i).Value
            Exit Function
        End If
    Next i

End Function

Private Function Set_FoxType_Sheet(ByVal szFoxType As String) As Worksheet

    Dim i As Integer
    
    'Search the worksheets collection seeing if any sheet is named after
    'the Fox Type
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = szFoxType Then
            Set Set_FoxType_Sheet = Worksheets(i)
            Exit Function
        End If
    Next i
    
    'A sheet not found for the latest Fox Type.  So we must
    'create a new sheet and name it after the new Type.
    Worksheets.Add.Move After:=Worksheets(Worksheets.Count)
    
    'Rename the brand new sheet here.
    Worksheets(Worksheets.Count).Name = szFoxType
    
    FreezeMySheet Worksheets(Worksheets.Count)
    
    Set Set_FoxType_Sheet = Worksheets(Worksheets.Count)
    
    
    'In addition to adding to the sheets collection,
    'we need to sync it with our array.
    mSheetsIndex = mSheetsIndex + 1
    
    If mSheetsIndex > UBound(mSheetsArray) Then
        ReDim Preserve mSheetsArray(0 To (mSheetsIndex + 100))
    End If
    
    mSheetsArray(mSheetsIndex).SheetName = szFoxType
    mSheetsArray(mSheetsIndex).Rows = 1     'for the header row
    mSheetsArray(mSheetsIndex).Columns = 0
    
    
End Function


Private Function Get_NextRow(ByVal i As Integer) As Long

    mSheetsArray(i).Rows = mSheetsArray(i).Rows + 1
    
    Get_NextRow = mSheetsArray(i).Rows
    
End Function

Private Function Get_CurrentIndex(ByVal szSheetName As String) As Long

    Dim i As Integer
    
    Get_CurrentIndex = -1       '0 is valid but -1 is totally BOGUS.  Totally.
    
    For i = 0 To mSheetsIndex
        If mSheetsArray(i).SheetName = szSheetName Then
            Get_CurrentIndex = i
            Exit Function
        End If
    Next i
    
End Function

Private Function Get_ColumnIndex(ByRef ws As Worksheet, ByVal szAttrib As String) As Integer
    
    'search Row 1 (the column header row) for the passed attribute name.
    Dim c As Integer
    Dim szThisCell As String
    
    For c = 1 To 255  'max is actually 256
        
        szThisCell = "" & ws.Cells(1, c).Value
        
        If szThisCell = "" Then
            'if a column header is empty, the attribute won't be found,
            'i.e. you have searched all current column headers so don't
            'keep searching thru all 256 columns.
            'Instead use the first empty cell in row 1 as the
            'new column header for the current attribute label.
            ws.Cells(1, c).Value = szAttrib
            szThisCell = szAttrib
        End If
        
        If szThisCell = szAttrib Then
            
            Get_ColumnIndex = c
            
            'while we are here let's check for max column number used.
            If c > mSheetsArray(mSheetsCurNdx).Columns Then
                 mSheetsArray(mSheetsCurNdx).Columns = c
            End If
            
            Exit Function
        
        End If
    Next c
    
    'To prevent later errors, set returned value to 0.
    'It is left to others to be sure this is not 0 before attempting writing.
    Get_ColumnIndex = 0

End Function


Private Sub Move_Line_To_Sheet(ByRef ws As Worksheet, ByVal r As Long, ByRef TAttribute As FoxProperty)
'Note that only a single line of text is passed, not entire string array

    Dim c As Integer
    
    c = Get_ColumnIndex(ws, TAttribute.Name)
        
    If (r > 0) And (c > 0) Then
        ws.Cells(r, c).Value = TAttribute.Value
    End If

End Sub

Public Sub Sort_Sheets()
'A slow but trustworthy bubble sort to alphabetize sheet names.

    Dim i As Integer
    Dim This As Worksheet
    Dim That As Worksheet
    Dim Sorted As Boolean
    
    Do While Not Sorted
        Sorted = True
        For i = 2 To Worksheets.Count
            Set This = Worksheets(i - 1)
            Set That = Worksheets(i)
            If That.Name < This.Name Then
                Sorted = False
                That.Move Before:=This
            End If
        Next i
        Set This = Nothing
        Set That = Nothing
    Loop
    
End Sub

Private Sub FreezeMySheet(ByRef ws As Worksheet)

    'Freezing is done to a Pane object which belongs to a Window object
    'The Window object is determined by the active worksheet
    ws.Activate
    
    ws.Columns(1).ColumnWidth = 25
    
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 1
        .SplitRow = 1
        .FreezePanes = True
    End With
    
End Sub

Public Sub FreezeAllSheets()
    
    Dim i As Integer
    
    For i = 1 To Worksheets.Count
        FreezeMySheet Worksheets(i)
    Next i
    
End Sub


Public Sub SheetStats()

    Dim wsStats As Worksheet
    Dim i As Integer
    Dim r As Long
    Dim c As Integer
    Dim RowPos As Long
    
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "_Dump_Stats_" Then
            Application.DisplayAlerts = False
            Worksheets(i).Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next i
    
    Worksheets.Add.Move Before:=Worksheets(1)
    'newly created sheet becomes new #1
    Set wsStats = Worksheets(1)
    wsStats.Name = "_Dump_Stats_"
    
    With wsStats
        .Cells(1, 1).Value = "Sheet"
        .Cells(1, 2).Value = "Rows"
        .Cells(1, 3).Value = "Columns"
    End With
    
    RowPos = 1  'this would be the header row
    
    Sort_SheetsArray
    
    For i = 1 To mSheetsIndex
        ' Skip any sheet name prefaced with underscore
        If Left(mSheetsArray(i).SheetName, 1) <> "_" Then
            RowPos = RowPos + 1
            r = mSheetsArray(i).Rows
            c = mSheetsArray(i).Columns
            'Subtract header row to return number of data records
            If r > 0 Then r = r - 1
            With wsStats
                .Cells(RowPos, 1).Value = mSheetsArray(i).SheetName
                .Cells(RowPos, 2).Value = r
                .Cells(RowPos, 3).Value = c
            End With
        End If
    Next i
    
    RowPos = RowPos + 1
    
    With wsStats
        .Columns(1).ColumnWidth = 20
        .Cells(RowPos, 1).Value = "# sheets = " & (RowPos - 2)
        .Cells(RowPos, 2).Formula = "=SUM(B2:B" & (RowPos - 1) & ")"
        .Cells(RowPos, 2).Calculate
    End With
    
    FreezeMySheet wsStats
    
    Set wsStats = Nothing

End Sub

'Private Sub GetRowColCount(ByRef ws As Worksheet, ByRef r As Long, ByRef c As Integer)
'
'    'Find first empty column in row 1
'    For c = 1 To 256
'        If ws.Cells(1, c).Value = "" Then
'            Exit For
'        End If
'    Next c
'    'this was empty column, so adjust to last good column
'    c = c - 1
'
'    'find first empty row in column A
'    r = Get_LastRow(ws.Name)
'    'this was empty row, so adjust to last good row
'    r = r - 1
'
'End Sub
'

Private Sub PrepWork()

    Dim ws As Worksheet
    Dim i As Integer
    Dim r As Long
    Dim c As Integer
    
    If Worksheets.Count = 1 And Worksheets(1).Name = "_INDEX_" Then
        Worksheets.Add.Move After:=Worksheets(Worksheets.Count)
        Set ws = Worksheets(Worksheets.Count)
        Randomize
        ws.Name = "ZZZ" & (Rnd(0) * 100000)
    End If
    
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "_INDEX_" Then
            Application.DisplayAlerts = False
            Worksheets(i).Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next i
    
    Worksheets.Add.Move Before:=Worksheets(1)
    'newly created sheet becomes new #1
    Set ws = Worksheets(1)
    ws.Name = "_INDEX_"
    
    mSheetsArray(0).SheetName = ws.Name
    mSheetsArray(0).Rows = 1   'row 0 reserved for header
    mSheetsArray(0).Columns = 6   'row 0 reserved for header
    
    ws.Cells(1, 1).Value = "NAME"
    ws.Cells(1, 2).Value = "TYPE"
    ws.Cells(1, 3).Value = "DESCRP"
    ws.Cells(1, 4).Value = "LOOPID"
    ws.Cells(1, 5).Value = "IOM_ID"   '--Optional -- all others required
    ws.Cells(1, 6).Value = "PNT_NO"   '--Optional -- all others required
    
    FreezeMySheet ws
    
    ws.Columns(2).ColumnWidth = 15
    ws.Columns(3).ColumnWidth = 30
    ws.Columns(4).ColumnWidth = 15
    ws.Columns(5).ColumnWidth = 25
    ws.Columns(6).ColumnWidth = 15

    Application.DisplayAlerts = False
    For i = Worksheets.Count To 2 Step -1
        Worksheets(i).Delete
    Next i
    Application.DisplayAlerts = True
    
    Set ws = Nothing
    
End Sub

Private Sub Index_FOB(ByRef TAttributes() As FoxProperty)

    Const MAXCOL = 6
    Const Threshold = (MAXCOL - 1)
    Dim c As Integer
    Dim r As Long
    Dim i As Integer
    Dim Votes As Integer
    
    If mws_Index Is Nothing Then
        Set mws_Index = Worksheets("_INDEX_")
    End If
    
    Votes = 0
    For c = 1 To MAXCOL
        Votes = Votes + AttributeHit(TAttributes(), "" & mws_Index.Cells(1, c).Value)
    Next c
    
    If Votes >= Threshold Then
        r = Get_NextRow(0)  '0 is for the _INDEX_ page
        'xxx where is the columns get???
        For c = 1 To MAXCOL
            i = AttributeRow(TAttributes(), "" & mws_Index.Cells(1, c).Value)
            If i > 0 Then
                mws_Index.Cells(r, c).Value = TAttributes(i).Value
            End If
        Next c
    End If
    
End Sub

Private Function AttributeHit(ByRef TAttributes() As FoxProperty, ByVal szAttribute As String) As Integer
' Don't want a pure boolean here.  We want an integer of 1 or 0 because
' we are summing the hits.
    
    If AttributeRow(TAttributes(), szAttribute) > 0 Then
        AttributeHit = 1
    Else
        AttributeHit = 0
    End If

End Function


Private Function AttributeRow(ByRef TAttributes() As FoxProperty, ByVal szAttribute As String) As Integer

    Dim i As Integer
    
    AttributeRow = 0
    
    szAttribute = UCase(Trim(szAttribute))
    
    If Len(szAttribute) = 0 Then Exit Function
    
    For i = 1 To UBound(TAttributes)
        If TAttributes(i).Name = szAttribute Then
            AttributeRow = i
            Exit For
        End If
    Next i


End Function

Public Sub Sort_Each_Sheets()

    Dim i As Integer
    Dim ws As Worksheet
    
    For i = 1 To Worksheets.Count
        Set ws = Worksheets(i)
        ws.Columns.Sort _
            Key1:=ws.Columns(1), _
            Order1:=xlAscending, _
            Header:=xlYes, _
            MatchCase:=False
        Set ws = Nothing
    Next i
    
End Sub


Private Function BrowseFileName(Optional ByVal szInitialFileName As String) As String

    'Declare a variable as a FileDialog object.
    Dim fd As Office.FileDialog

    'Create a FileDialog object as a File Picker dialog box.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .AllowMultiSelect = False
        .Filters.Add "Dumps", "*.txt", 1
        .FilterIndex = 1
        .InitialFileName = szInitialFileName  'or directory
        .InitialView = msoFileDialogViewDetails
        .Title = "Browse Fox Node Dumps"
    End With

    'Declare a variable to contain the path
    'of each selected item. Even though the path is a String,
    'the variable must be a Variant because For Each...Next
    'routines only work with Variants and Objects.
    Dim vrtSelectedItem As Variant

    'Use a With...End With block to reference the FileDialog object.
    With fd

        'Use the Show method to display the File Picker dialog box
        'The user pressed the action button.
        If .Show = -1 Then

            'Step through the FileDialogSelectedItems collection.
            For Each vrtSelectedItem In .SelectedItems

                'vrtSelectedItem contains the path of each selected item.
                'Here use any file I/O functions you want on the path.
                'This example simply displays the path in a message box.
                BrowseFileName = vrtSelectedItem

            Next vrtSelectedItem
        'The user pressed Cancel.
        Else
        End If
    End With

    'Set the object variable to Nothing.
    Set fd = Nothing

End Function
                
Private Sub Post_Parse()

    Dim wsThis As Worksheet
    Dim i As Integer
    Dim bRun As Boolean
    
    For i = 1 To Worksheets.Count
    
        bRun = False
        
        Set wsThis = Worksheets(i)
        
        ' Skip any sheet name prefaced with underscore, unless its "_INDEX_"
        If wsThis.Name = "_INDEX_" Then
            bRun = True
            Debug.Print "TRUE for _INDEX_"
        ElseIf Left(wsThis.Name, 1) <> "_" Then
            bRun = True
        End If
            
        If bRun Then
            ChopBlock wsThis
        End If
        
        Set wsThis = Nothing
    
    Next i

End Sub


Private Sub ChopBlock(ByRef ws As Worksheet)
'-----------------------------------------------------------------------------------------
'This should be called AFTER sheet has been completed filled from the node dump.
'-----------------------------------------------------------------------------------------
'
'For the specified sheet, 3 new columns will be added before column C.
'  (C) _compound_, (D) _block_, and (E) _unit_
'
'While the algorthm is exact in what it does, the results are NOT exact in what we want.
' (1) The NAME (Col A) is parsed into 2 parts.
' (2) The 1st part is the Compound, or anything found prior to a colon.
' (3) The 2nd part is the Block, or anything found after the colon.
' (4) The Block is then parsed to find the 1st two consecutive numerals (0-9), which is assumed to be the Unit.
'
'There's a lot of assumption going on and a lot of "if any" and "has none".

    Dim c As Integer
    Dim r As Long
    Dim LastRow As Long
    Dim rng As Excel.Range
    Dim x As Integer
    Dim szName As String
    Dim szCmpd As String
    Dim szBlck As String
    Dim szUnit As String

    On Error Resume Next
    
    ws.Columns(3).Insert shift:=xlShiftToRight
    ws.Columns(3).Insert shift:=xlShiftToRight
    ws.Columns(3).Insert shift:=xlShiftToRight

    For c = 3 To 5
    
        With ws.Columns(c)
            .ColumnWidth = 18
            .HorizontalAlignment = xlCenter
        End With
        
        Set rng = ws.Cells(1, c)
        
        rng.Font.Bold = True
        rng.Interior.Color = RGB(229, 229, 255)
        
        Select Case c
            Case 3:     rng.Value = "_compound_"
            Case 4:     rng.Value = "_block_"
            Case 5:     rng.Value = "_unit_"
        End Select
    
    Next c

    mSheetsCurNdx = Get_CurrentIndex(ws.Name)
    LastRow = mSheetsArray(mSheetsCurNdx).Rows

Debug.Print "CHOP [" & ws.Name & "] Row=" & LastRow & ", Col=" & c

    For r = 2 To LastRow
        
        szName = ""
        szCmpd = ""
        szBlck = ""
        szUnit = ""
        
        szName = Trim("" & ws.Cells(r, 1).Value)
        If szName = "" Then Exit For
            
        x = InStr(1, szName, ":", vbTextCompare)
        
        If x > 0 Then
            szCmpd = Trim(Mid(szName, 1, x - 1))
            szBlck = Trim(Mid(szName, x + 1))
        Else
            szCmpd = szName
        End If
        
        If Len(szBlck) > 0 Then
            szUnit = SniffUnit(szBlck)
        End If
    
        ws.Cells(r, 3).Value = szCmpd
        ws.Cells(r, 4).Value = szBlck
        ws.Cells(r, 5).Value = szUnit
    
    Next r

End Sub

Private Function SniffUnit(ByVal Block As String) As String
    
    Dim i As Integer
    Dim Digits As Integer
    Dim Char As String
    
    For i = 1 To Len(Block)
        
        Char = Mid(Block, i, 1)
        
        If InStr(1, "0123456789", Char) Then
            Digits = Digits + 1
        Else
            Digits = 0
        End If
        
        If Digits = 2 Then
            'Must preface unit with a single quote to force it to be a string when displayed in a cell.
            'Otherwise Excel would show Unit 04 as simply 4.
            SniffUnit = "'" & Mid(Block, i - 1, 2)
            Exit For
        End If
        
    Next i
    
End Function

Private Sub Sort_SheetsArray()
'slow but trusty bubblesort

    Dim i As Integer
    Dim Temp As udtSheets
    Dim Sorted As Boolean
    
    Sorted = False
    
    Do While Not Sorted
        Sorted = True
        For i = 1 To (mSheetsIndex - 1)
            If mSheetsArray(i).SheetName > mSheetsArray(i + 1).SheetName Then
                Sorted = False
                Temp = mSheetsArray(i)
                mSheetsArray(i) = mSheetsArray(i + 1)
                mSheetsArray(i + 1) = Temp
            End If
        Next i
    Loop

End Sub


