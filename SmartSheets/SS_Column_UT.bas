Attribute VB_Name = "SS_Column_UT"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.SmartSheets")

Private Assert As New Rubberduck.AssertClass
Private Const TestName As String = "TestWS"

Private Function WorksheetFromCodeName(ByVal CodeName As String) As Worksheet

    Dim WS As Worksheet
    For Each WS In ThisWorkbook.Worksheets
        If StrComp(WS.CodeName, CodeName, vbTextCompare) = 0 Then
            Set WorksheetFromCodeName = WS
            Exit Function
        End If
    Next WS

End Function

Private Function RndInt(ByVal Lower As Long, ByVal Upper As Long) As Long
    RndInt = Int((Upper - Lower + 1) * Rnd + Lower)
End Function

Private Function RndString(ByVal Length As Integer) As String
    Dim i As Integer

    For i = 1 To Length
        RndString = RndString & ChrW$(RndInt(32, 126))
    Next i

End Function

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
    Dim TestWS As Worksheet
    Set TestWS = WorksheetFromCodeName(TestName)
    If TestWS Is Nothing Then
        ThisWorkbook.Worksheets.Add After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set TestWS = ThisWorkbook.ActiveSheet
        ThisWorkbook.VBProject.VBComponents(TestWS.CodeName).Name = TestName
    Else
        With TestWS.Cells
            .ClearContents
            .ClearFormats
        End With
    End If

    ' Make sure all SS_Column methods and Tests account for escaped apostrophes ('')
    TestWS.Name = "LTG's Sheet"
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
    Dim TestWS As Worksheet
    Set TestWS = WorksheetFromCodeName(TestName)

    If Not TestWS Is Nothing Then
        Dim fDisplayAlerts As Boolean
        fDisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        TestWS.Delete
        Application.DisplayAlerts = fDisplayAlerts
    End If

End Sub

'@TestMethod("Init Checks")
Public Sub Test_Uninit()
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    'Arrange:
    Const ExpectedError As Long = ERROR_SSC.SSCE_ObjUninit
    Dim TestColumn As New SS_Column

    'Act/Assert:

    ' Property Init Checks
    Debug.Print TestColumn.Name
    Assert.AreEqual ExpectedError, Err.Number, "Property ""Name"" failed uninitialized check."
    Err.Clear

    Debug.Print TestColumn.Number
    Assert.AreEqual ExpectedError, Err.Number, "Property ""Number"" failed uninitialized check."
    Err.Clear

    Debug.Print TestColumn.Hidden
    Assert.AreEqual ExpectedError, Err.Number, "Property ""Hidden"" Get failed uninitialized check."
    Err.Clear

    TestColumn.Hidden = True
    Assert.AreEqual ExpectedError, Err.Number, "Property ""Hidden"" Let failed uninitialized check."
    Err.Clear

    Debug.Print TestColumn.Title
    Assert.AreEqual ExpectedError, Err.Number, "Property ""Title"" Get failed uninitialized check."
    Err.Clear

    TestColumn.Title = "Test Title"
    Assert.AreEqual ExpectedError, Err.Number, "Property ""Title"" Let failed uninitialized check."
    Err.Clear

    Debug.Print TestColumn.TitleCell.Value
    Assert.AreEqual ExpectedError, Err.Number, "Property ""TitleCell"" failed uninitialized check."
    Err.Clear

    Debug.Print TestColumn.Column
    Assert.AreEqual ExpectedError, Err.Number, "Property ""Column"" failed uninitialized check."
    Err.Clear

    Debug.Print TestColumn.LastRow
    Assert.AreEqual ExpectedError, Err.Number, "Property ""LastRow"" failed uninitialized check."
    Err.Clear

    ' Procedure/Function Init checks
    TestColumn.ShiftLeft
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""ShiftLeft"" failed uninitialized check."
    Err.Clear

    TestColumn.ShiftRight
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""ShiftRight"" failed uninitialized check."
    Err.Clear

    TestColumn.SetIndex 1
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""SetColumn"" failed uninitialized check."
    Err.Clear

    TestColumn.ColumnAddress
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""ColumnAddress"" failed uninitialized check."
    Err.Clear

    TestColumn.RowAddress 1
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""RowAddress"" failed uninitialized check."
    Err.Clear

    TestColumn.RangeAddress 1, 2
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""RangeAddress"" failed uninitialized check."
    Err.Clear

    TestColumn.AddressC1
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""AddressC1"" failed uninitialized check."
    Err.Clear

    TestColumn.AddressR1C1
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""AddressR1C1"" failed uninitialized check."
    Err.Clear

    TestColumn.Range
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""Range"" failed uninitialized check."
    Err.Clear

    TestColumn.FillDown
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""FillDown"" failed uninitialized check."
    Err.Clear

    TestColumn.FillUp
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""FillUp"" failed uninitialized check."
    Err.Clear

    TestColumn.Cell 2
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""Cell"" failed uninitialized check."
    Err.Clear

    TestColumn.ClearContents
    Assert.AreEqual ExpectedError, Err.Number, "Routine ""ClearContents"" failed uninitialized check."

End Sub

'@TestMethod("Name Checks")
Public Sub Test_Names()
    On Error GoTo TestFail

    'Arrange:
    Dim TestWS As Worksheet
    Dim TestColumn As SS_Column
    Dim Column As Long

    Set TestWS = WorksheetFromCodeName(TestName)
    Set TestColumn = New SS_Column
    TestColumn.Init TestWS, 1

    'Act/Assert:
    Assert.AreEqual TestWS.Columns(1).Address(ColumnAbsolute:=False), TestColumn.Name & ":" & TestColumn.Name, "Column naming broken on Column: " & TestColumn.Name & "(" & TestColumn.Number & ")"
    Assert.AreEqual 1&, TestColumn.Number, "Column numbering broken on Column: " & TestColumn.Name & "(" & TestColumn.Number & ")"

    For Column = 2 To ActiveSheet.Columns.Count
        TestColumn.ShiftRight
        Assert.AreEqual TestWS.Columns(Column).Address(ColumnAbsolute:=False), TestColumn.Name & ":" & TestColumn.Name, "Column naming broken on Column: " & TestColumn.Name & "(" & TestColumn.Number & ")"
        Assert.AreEqual Column, TestColumn.Number, "Column numbering broken on Column: " & TestColumn.Name & "(" & TestColumn.Number & ")"
    Next Column

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Address Tests")
Public Sub Test_Column_Addr()
    On Error GoTo TestFail

    'Arrange:
    Dim TestColumn As New SS_Column
    Dim nameWS As String
    Dim nameWB As String
    Dim formCRel As String
    Dim formCRelWS As String
    Dim formCRelWB As String
    Dim formCAbs As String
    Dim formCAbsWS As String
    Dim formCAbsWB As String

    Dim wsRef As String
    Dim wbRef As String
    Const relRef As String = "A:A"
    Const absRef As String = "$A:$A"

    TestColumn.Init ActiveSheet, 1
    nameWS = ThisWorkbook.ActiveSheet.Name
    nameWB = ThisWorkbook.Name
    wsRef = "'" & Replace(nameWS, "'", "''") & "'!"
    wbRef = "'[" & nameWB & "]" & Replace(nameWS, "'", "''") & "'!"

    'Act:
    formCRel = TestColumn.ColumnAddress(False)
    formCRelWS = TestColumn.ColumnAddress(False, True)
    formCRelWB = TestColumn.ColumnAddress(False, True, True)
    formCAbs = TestColumn.ColumnAddress(True)
    formCAbsWS = TestColumn.ColumnAddress(True, True)
    formCAbsWB = TestColumn.ColumnAddress(True, True, True)

    'Assert:
    Assert.AreEqual relRef, formCRel, "Relative column address generation is broken."
    Assert.AreEqual wsRef & relRef, formCRelWS, "Relative WS column address generation is broken."
    Assert.AreEqual wbRef & relRef, formCRelWB, "Relative WB column address generation is broken."

    Assert.AreEqual absRef, formCAbs, "Absolute column address generation is broken."
    Assert.AreEqual wsRef & absRef, formCAbsWS, "Absolute WS column address generation is broken."
    Assert.AreEqual wbRef & absRef, formCAbsWB, "Absolute WB column address generation is broken."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Address Tests")
Public Sub Test_Row_Addr()
    On Error GoTo TestFail

    'Arrange:
    Dim TestColumn As New SS_Column
    Dim nameWS As String
    Dim nameWB As String

    Dim formCRRelRel As String
    Dim formCRRelRelWS As String
    Dim formCRRelRelWB As String

    Dim formCRAbsRel As String
    Dim formCRAbsRelWS As String
    Dim formCRAbsRelWB As String

    Dim formCRRelAbs As String
    Dim formCRRelAbsWS As String
    Dim formCRRelAbsWB As String

    Dim formCRAbsAbs As String
    Dim formCRAbsAbsWS As String
    Dim formCRAbsAbsWB As String

    Dim wsRef As String
    Dim wbRef As String
    Const Row As Long = 1
    Const rrRef As String = "A1"
    Const arRef As String = "$A1"
    Const raRef As String = "A$1"
    Const aaRef As String = "$A$1"

    TestColumn.Init ActiveSheet, 1
    nameWS = ThisWorkbook.ActiveSheet.Name
    nameWB = ThisWorkbook.Name
    wsRef = "'" & Replace(nameWS, "'", "''") & "'!"
    wbRef = "'[" & nameWB & "]" & Replace(nameWS, "'", "''") & "'!"

    'Act:
    formCRRelRel = TestColumn.RowAddress(Row)
    formCRRelRelWS = TestColumn.RowAddress(Row, IncludeWS:=True)
    formCRRelRelWB = TestColumn.RowAddress(Row, IncludeWB:=True)

    formCRAbsRel = TestColumn.RowAddress(Row, True)
    formCRAbsRelWS = TestColumn.RowAddress(Row, True, IncludeWS:=True)
    formCRAbsRelWB = TestColumn.RowAddress(Row, True, IncludeWB:=True)

    formCRRelAbs = TestColumn.RowAddress(Row, False, True)
    formCRRelAbsWS = TestColumn.RowAddress(Row, False, True, IncludeWS:=True)
    formCRRelAbsWB = TestColumn.RowAddress(Row, False, True, IncludeWB:=True)

    formCRAbsAbs = TestColumn.RowAddress(Row, True, True)
    formCRAbsAbsWS = TestColumn.RowAddress(Row, True, True, IncludeWS:=True)
    formCRAbsAbsWB = TestColumn.RowAddress(Row, True, True, IncludeWB:=True)

    'Assert:
    Assert.AreEqual rrRef, formCRRelRel, "Relative column/Relative row address generation is broken."
    Assert.AreEqual wsRef & rrRef, formCRRelRelWS, "Relative column/Relative row WS column address generation is broken."
    Assert.AreEqual wbRef & rrRef, formCRRelRelWB, "Relative column/Relative row WB column address generation is broken."

    Assert.AreEqual arRef, formCRAbsRel, "Absolute column/Relative row address generation is broken."
    Assert.AreEqual wsRef & arRef, formCRAbsRelWS, "Absolute column/Relative row WS column address generation is broken."
    Assert.AreEqual wbRef & arRef, formCRAbsRelWB, "Absolute column/Relative row WB column address generation is broken."

    Assert.AreEqual raRef, formCRRelAbs, "Relative column/Absolute row address generation is broken."
    Assert.AreEqual wsRef & raRef, formCRRelAbsWS, "Relative column/Absolute row WS column address generation is broken."
    Assert.AreEqual wbRef & raRef, formCRRelAbsWB, "Relative column/Absolute row WB column address generation is broken."

    Assert.AreEqual aaRef, formCRAbsAbs, "Absolute column/Absolute row address generation is broken."
    Assert.AreEqual wsRef & aaRef, formCRAbsAbsWS, "Absolute column/Absolute row WS column address generation is broken."
    Assert.AreEqual wbRef & aaRef, formCRAbsAbsWB, "Absolute column/Absolute row WB column address generation is broken."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Address Tests")
Public Sub Test_Range_Addr()
    On Error GoTo TestFail

    'Arrange:
    Dim TestColumn As New SS_Column
    Dim nameWS As String
    Dim nameWB As String

    Dim formCRRelRel As String
    Dim formCRRelRelWS As String
    Dim formCRRelRelWB As String

    Dim formCRAbsRel As String
    Dim formCRAbsRelWS As String
    Dim formCRAbsRelWB As String

    Dim formCRRelAbs As String
    Dim formCRRelAbsWS As String
    Dim formCRRelAbsWB As String

    Dim formCRAbsAbs As String
    Dim formCRAbsAbsWS As String
    Dim formCRAbsAbsWB As String

    Dim wsRef As String
    Dim wbRef As String
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim rrRef As String
    Dim arRef As String
    Dim raRef As String
    Dim aaRef As String

    rowStart = Int(ActiveSheet.Rows.Count * Rnd + 1)
    rowEnd = Int((ActiveSheet.Rows.Count - rowStart + 1) * Rnd + rowStart)
    rrRef = "A" & rowStart & ":A" & rowEnd
    arRef = "$A" & rowStart & ":$A" & rowEnd
    raRef = "A$" & rowStart & ":A$" & rowEnd
    aaRef = "$A$" & rowStart & ":$A$" & rowEnd

    TestColumn.Init ActiveSheet, 1
    nameWS = ThisWorkbook.ActiveSheet.Name
    nameWB = ThisWorkbook.Name
    wsRef = "'" & Replace(nameWS, "'", "''") & "'!"
    wbRef = "'[" & nameWB & "]" & Replace(nameWS, "'", "''") & "'!"

    'Act:
    formCRRelRel = TestColumn.RangeAddress(rowStart, rowEnd)
    formCRRelRelWS = TestColumn.RangeAddress(rowStart, rowEnd, IncludeWS:=True)
    formCRRelRelWB = TestColumn.RangeAddress(rowStart, rowEnd, IncludeWB:=True)

    formCRAbsRel = TestColumn.RangeAddress(rowStart, rowEnd, True)
    formCRAbsRelWS = TestColumn.RangeAddress(rowStart, rowEnd, True, IncludeWS:=True)
    formCRAbsRelWB = TestColumn.RangeAddress(rowStart, rowEnd, True, IncludeWB:=True)

    formCRRelAbs = TestColumn.RangeAddress(rowStart, rowEnd, False, True)
    formCRRelAbsWS = TestColumn.RangeAddress(rowStart, rowEnd, False, True, IncludeWS:=True)
    formCRRelAbsWB = TestColumn.RangeAddress(rowStart, rowEnd, False, True, IncludeWB:=True)

    formCRAbsAbs = TestColumn.RangeAddress(rowStart, rowEnd, True, True)
    formCRAbsAbsWS = TestColumn.RangeAddress(rowStart, rowEnd, True, True, IncludeWS:=True)
    formCRAbsAbsWB = TestColumn.RangeAddress(rowStart, rowEnd, True, True, IncludeWB:=True)

    'Assert:
    Assert.AreEqual rrRef, formCRRelRel, "Relative column/Relative row range address generation is broken."
    Assert.AreEqual wsRef & rrRef, formCRRelRelWS, "Relative column/Relative row WS column range address generation is broken."
    Assert.AreEqual wbRef & rrRef, formCRRelRelWB, "Relative column/Relative row WB column range address generation is broken."

    Assert.AreEqual arRef, formCRAbsRel, "Absolute column/Relative row range address generation is broken."
    Assert.AreEqual wsRef & arRef, formCRAbsRelWS, "Absolute column/Relative row WS column range address generation is broken."
    Assert.AreEqual wbRef & arRef, formCRAbsRelWB, "Absolute column/Relative row WB column range address generation is broken."

    Assert.AreEqual raRef, formCRRelAbs, "Relative column/Absolute row range address generation is broken."
    Assert.AreEqual wsRef & raRef, formCRRelAbsWS, "Relative column/Absolute row WS column range address generation is broken."
    Assert.AreEqual wbRef & raRef, formCRRelAbsWB, "Relative column/Absolute row WB column range address generation is broken."

    Assert.AreEqual aaRef, formCRAbsAbs, "Absolute column/Absolute row range address generation is broken."
    Assert.AreEqual wsRef & aaRef, formCRAbsAbsWS, "Absolute column/Absolute row WS column range address generation is broken."
    Assert.AreEqual wbRef & aaRef, formCRAbsAbsWB, "Absolute column/Absolute row WB column range address generation is broken."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Address Tests")
Public Sub Test_Range_Same_StartEnd()
    On Error GoTo TestFail

    'Arrange:
    Dim TestColumn As New SS_Column
    Dim nameWS As String
    Dim nameWB As String

    Dim formCRRelRel As String
    Dim formCRRelRelWS As String
    Dim formCRRelRelWB As String

    Dim formCRAbsRel As String
    Dim formCRAbsRelWS As String
    Dim formCRAbsRelWB As String

    Dim formCRRelAbs As String
    Dim formCRRelAbsWS As String
    Dim formCRRelAbsWB As String

    Dim formCRAbsAbs As String
    Dim formCRAbsAbsWS As String
    Dim formCRAbsAbsWB As String

    Dim wsRef As String
    Dim wbRef As String
    Dim Row As Long
    Dim rrRef As String
    Dim arRef As String
    Dim raRef As String
    Dim aaRef As String

    Row = Int(ActiveSheet.Rows.Count * Rnd + 1)
    rrRef = "A" & Row
    arRef = "$A" & Row
    raRef = "A$" & Row
    aaRef = "$A$" & Row

    TestColumn.Init ActiveSheet, 1
    nameWS = ThisWorkbook.ActiveSheet.Name
    nameWB = ThisWorkbook.Name
    wsRef = "'" & Replace(nameWS, "'", "''") & "'!"
    wbRef = "'[" & nameWB & "]" & Replace(nameWS, "'", "''") & "'!"

    'Act:
    formCRRelRel = TestColumn.RangeAddress(Row, Row)
    formCRRelRelWS = TestColumn.RangeAddress(Row, Row, IncludeWS:=True)
    formCRRelRelWB = TestColumn.RangeAddress(Row, Row, IncludeWB:=True)

    formCRAbsRel = TestColumn.RangeAddress(Row, Row, True)
    formCRAbsRelWS = TestColumn.RangeAddress(Row, Row, True, IncludeWS:=True)
    formCRAbsRelWB = TestColumn.RangeAddress(Row, Row, True, IncludeWB:=True)

    formCRRelAbs = TestColumn.RangeAddress(Row, Row, False, True)
    formCRRelAbsWS = TestColumn.RangeAddress(Row, Row, False, True, IncludeWS:=True)
    formCRRelAbsWB = TestColumn.RangeAddress(Row, Row, False, True, IncludeWB:=True)

    formCRAbsAbs = TestColumn.RangeAddress(Row, Row, True, True)
    formCRAbsAbsWS = TestColumn.RangeAddress(Row, Row, True, True, IncludeWS:=True)
    formCRAbsAbsWB = TestColumn.RangeAddress(Row, Row, True, True, IncludeWB:=True)

    'Assert:
    Assert.AreEqual rrRef, formCRRelRel, "Relative column/Relative row range address generation with same row is broken."
    Assert.AreEqual wsRef & rrRef, formCRRelRelWS, "Relative column/Relative row WS column range address generation with same row is broken."
    Assert.AreEqual wbRef & rrRef, formCRRelRelWB, "Relative column/Relative row WB column range address generation with same row is broken."

    Assert.AreEqual arRef, formCRAbsRel, "Absolute column/Relative row range address generation with same row is broken."
    Assert.AreEqual wsRef & arRef, formCRAbsRelWS, "Absolute column/Relative row WS column range address generation with same row is broken."
    Assert.AreEqual wbRef & arRef, formCRAbsRelWB, "Absolute column/Relative row WB column range address generation with same row is broken."

    Assert.AreEqual raRef, formCRRelAbs, "Relative column/Absolute row range address generation with same row is broken."
    Assert.AreEqual wsRef & raRef, formCRRelAbsWS, "Relative column/Absolute row WS column range address generation with same row is broken."
    Assert.AreEqual wbRef & raRef, formCRRelAbsWB, "Relative column/Absolute row WB column range address generation with same row is broken."

    Assert.AreEqual aaRef, formCRAbsAbs, "Absolute column/Absolute row range address generation with same row is broken."
    Assert.AreEqual wsRef & aaRef, formCRAbsAbsWS, "Absolute column/Absolute row WS column range address generation with same row is broken."
    Assert.AreEqual wbRef & aaRef, formCRAbsAbsWB, "Absolute column/Absolute row WB column range address generation with same row is broken."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Address Tests")
Public Sub Test_C1_Addr()
    On Error GoTo TestFail

    'Arrange:
    Dim TestColumn As New SS_Column
    Dim nameWS As String
    Dim nameWB As String
    Dim formCRel As String
    Dim formCRelWS As String
    Dim formCRelWB As String
    Dim formCAbs As String
    Dim formCAbsWS As String
    Dim formCAbsWB As String

    Dim wsRef As String
    Dim wbRef As String
    Const relRef As String = "C"
    Const absRef As String = "C1"

    TestColumn.Init ActiveSheet, 1
    nameWS = ThisWorkbook.ActiveSheet.Name
    nameWB = ThisWorkbook.Name
    wsRef = "'" & Replace(nameWS, "'", "''") & "'!"
    wbRef = "'[" & nameWB & "]" & Replace(nameWS, "'", "''") & "'!"

    'Act:
    formCRel = TestColumn.AddressC1(False)
    formCRelWS = TestColumn.AddressC1(False, True)
    formCRelWB = TestColumn.AddressC1(False, True, True)
    formCAbs = TestColumn.AddressC1(True)
    formCAbsWS = TestColumn.AddressC1(True, True)
    formCAbsWB = TestColumn.AddressC1(True, True, True)

    'Assert:
    Assert.AreEqual relRef, formCRel, "Relative C1 column address generation is broken."
    Assert.AreEqual wsRef & relRef, formCRelWS, "Relative C1 WS column address generation is broken."
    Assert.AreEqual wbRef & relRef, formCRelWB, "Relative C1 WB column address generation is broken."

    Assert.AreEqual absRef, formCAbs, "Absolute C1 column address generation is broken."
    Assert.AreEqual wsRef & absRef, formCAbsWS, "Absolute C1 WS column address generation is broken."
    Assert.AreEqual wbRef & absRef, formCAbsWB, "Absolute C1 WB column address generation is broken."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Address Tests")
Public Sub Test_R1C1_PosRow()
    On Error GoTo TestFail

    'Arrange:
    Dim TestColumn As New SS_Column
    Dim nameWS As String
    Dim nameWB As String

    Dim formCRRelRel As String
    Dim formCRRelRelWS As String
    Dim formCRRelRelWB As String

    Dim formCRAbsRel As String
    Dim formCRAbsRelWS As String
    Dim formCRAbsRelWB As String

    Dim formCRRelAbs As String
    Dim formCRRelAbsWS As String
    Dim formCRRelAbsWB As String

    Dim formCRAbsAbs As String
    Dim formCRAbsAbsWS As String
    Dim formCRAbsAbsWB As String

    Dim wsRef As String
    Dim wbRef As String
    Const Row As Long = 1
    Const rrRef As String = "R[1]C"
    Const arRef As String = "R[1]C1"
    Const raRef As String = "R1C"
    Const aaRef As String = "R1C1"

    TestColumn.Init ActiveSheet, 1
    nameWS = ThisWorkbook.ActiveSheet.Name
    nameWB = ThisWorkbook.Name
    wsRef = "'" & Replace(nameWS, "'", "''") & "'!"
    wbRef = "'[" & nameWB & "]" & Replace(nameWS, "'", "''") & "'!"

    'Act:
    formCRRelRel = TestColumn.AddressR1C1(Row)
    formCRRelRelWS = TestColumn.AddressR1C1(Row, IncludeWS:=True)
    formCRRelRelWB = TestColumn.AddressR1C1(Row, IncludeWB:=True)

    formCRAbsRel = TestColumn.AddressR1C1(Row, True)
    formCRAbsRelWS = TestColumn.AddressR1C1(Row, True, IncludeWS:=True)
    formCRAbsRelWB = TestColumn.AddressR1C1(Row, True, IncludeWB:=True)

    formCRRelAbs = TestColumn.AddressR1C1(Row, False, True)
    formCRRelAbsWS = TestColumn.AddressR1C1(Row, False, True, IncludeWS:=True)
    formCRRelAbsWB = TestColumn.AddressR1C1(Row, False, True, IncludeWB:=True)

    formCRAbsAbs = TestColumn.AddressR1C1(Row, True, True)
    formCRAbsAbsWS = TestColumn.AddressR1C1(Row, True, True, IncludeWS:=True)
    formCRAbsAbsWB = TestColumn.AddressR1C1(Row, True, True, IncludeWB:=True)

    'Assert:
    Assert.AreEqual rrRef, formCRRelRel, "R1C1 Relative column/Relative row address generation is broken."
    Assert.AreEqual wsRef & rrRef, formCRRelRelWS, "R1C1 Relative column/Relative row WS column address generation is broken."
    Assert.AreEqual wbRef & rrRef, formCRRelRelWB, "R1C1 Relative column/Relative row WB column address generation is broken."

    Assert.AreEqual arRef, formCRAbsRel, "R1C1 Absolute column/Relative row address generation is broken."
    Assert.AreEqual wsRef & arRef, formCRAbsRelWS, "R1C1 Absolute column/Relative row WS column address generation is broken."
    Assert.AreEqual wbRef & arRef, formCRAbsRelWB, "R1C1 Absolute column/Relative row WB column address generation is broken."

    Assert.AreEqual raRef, formCRRelAbs, "R1C1 Relative column/Absolute row address generation is broken."
    Assert.AreEqual wsRef & raRef, formCRRelAbsWS, "R1C1 Relative column/Absolute row WS column address generation is broken."
    Assert.AreEqual wbRef & raRef, formCRRelAbsWB, "R1C1 Relative column/Absolute row WB column address generation is broken."

    Assert.AreEqual aaRef, formCRAbsAbs, "R1C1 Absolute column/Absolute row address generation is broken."
    Assert.AreEqual wsRef & aaRef, formCRAbsAbsWS, "R1C1 Absolute column/Absolute row WS column address generation is broken."
    Assert.AreEqual wbRef & aaRef, formCRAbsAbsWB, "R1C1 Absolute column/Absolute row WB column address generation is broken."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Address Tests")
Public Sub Test_R1C1_NegRow()
    On Error GoTo TestFail

    'Arrange:
    Dim TestColumn As New SS_Column
    Dim nameWS As String
    Dim nameWB As String

    Dim formCRRelRel As String
    Dim formCRRelRelWS As String
    Dim formCRRelRelWB As String

    Dim formCRAbsRel As String
    Dim formCRAbsRelWS As String
    Dim formCRAbsRelWB As String

    Dim wsRef As String
    Dim wbRef As String
    Const Row As Long = -1
    Const rrRef As String = "R[-1]C"
    Const arRef As String = "R[-1]C1"
    Const ExpectedError As Long = ERROR_SSC.SSCE_InvalParam

    TestColumn.Init ActiveSheet, 1
    nameWS = ThisWorkbook.ActiveSheet.Name
    nameWB = ThisWorkbook.Name
    wsRef = "'" & Replace(nameWS, "'", "''") & "'!"
    wbRef = "'[" & nameWB & "]" & Replace(nameWS, "'", "''") & "'!"

    'Act1:
    formCRRelRel = TestColumn.AddressR1C1(Row)
    formCRRelRelWS = TestColumn.AddressR1C1(Row, IncludeWS:=True)
    formCRRelRelWB = TestColumn.AddressR1C1(Row, IncludeWB:=True)

    formCRAbsRel = TestColumn.AddressR1C1(Row, True)
    formCRAbsRelWS = TestColumn.AddressR1C1(Row, True, IncludeWS:=True)
    formCRAbsRelWB = TestColumn.AddressR1C1(Row, True, IncludeWB:=True)

    'Assert1:
    Assert.AreEqual rrRef, formCRRelRel, "R1C1 Relative column/Relative row address generation is broken."
    Assert.AreEqual wsRef & rrRef, formCRRelRelWS, "R1C1 Relative column/Relative row WS column address generation is broken."
    Assert.AreEqual wbRef & rrRef, formCRRelRelWB, "R1C1 Relative column/Relative row WB column address generation is broken."

    Assert.AreEqual arRef, formCRAbsRel, "R1C1 Absolute column/Relative row address generation is broken."
    Assert.AreEqual wsRef & arRef, formCRAbsRelWS, "R1C1 Absolute column/Relative row WS column address generation is broken."
    Assert.AreEqual wbRef & arRef, formCRAbsRelWB, "R1C1 Absolute column/Relative row WB column address generation is broken."

    ' Act2:
    ' All should fail because Negative rows are not allowed for absolute row references.
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    TestColumn.AddressR1C1 Row, False, True
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 negative row is mistakenly allowed."
    Err.Clear

    TestColumn.AddressR1C1 Row, False, True, IncludeWS:=True
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 negative row is mistakenly allowed."
    Err.Clear

    TestColumn.AddressR1C1 Row, False, True, IncludeWB:=True
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 negative row is mistakenly allowed."
    Err.Clear

    TestColumn.AddressR1C1 Row, True, True
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 negative row is mistakenly allowed."
    Err.Clear

    TestColumn.AddressR1C1 Row, True, True, IncludeWS:=True
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 negative row is mistakenly allowed."
    Err.Clear

    Call TestColumn.AddressR1C1(Row, True, True, IncludeWB:=True)
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 negative row is mistakenly allowed."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Address Tests")
Public Sub Test_R1C1_ZeroRow()
    On Error GoTo TestFail

    'Arrange:
    Dim TestColumn As New SS_Column
    Dim nameWS As String
    Dim nameWB As String

    Dim formCRRelRel As String
    Dim formCRRelRelWS As String
    Dim formCRRelRelWB As String

    Dim formCRAbsRel As String
    Dim formCRAbsRelWS As String
    Dim formCRAbsRelWB As String

    Dim wsRef As String
    Dim wbRef As String
    Const Row As Long = 0
    Const rrRef As String = "RC"
    Const arRef As String = "RC1"
    Const ExpectedError As Long = ERROR_SSC.SSCE_InvalParam

    TestColumn.Init ActiveSheet, 1
    nameWS = ThisWorkbook.ActiveSheet.Name
    nameWB = ThisWorkbook.Name
    wsRef = "'" & Replace(nameWS, "'", "''") & "'!"
    wbRef = "'[" & nameWB & "]" & Replace(nameWS, "'", "''") & "'!"

    'Act1:
    formCRRelRel = TestColumn.AddressR1C1(Row)
    formCRRelRelWS = TestColumn.AddressR1C1(Row, IncludeWS:=True)
    formCRRelRelWB = TestColumn.AddressR1C1(Row, IncludeWB:=True)

    formCRAbsRel = TestColumn.AddressR1C1(Row, True)
    formCRAbsRelWS = TestColumn.AddressR1C1(Row, True, IncludeWS:=True)
    formCRAbsRelWB = TestColumn.AddressR1C1(Row, True, IncludeWB:=True)

    'Assert1:
    Assert.AreEqual rrRef, formCRRelRel, "R1C1 Relative column/Relative row address generation is broken."
    Assert.AreEqual wsRef & rrRef, formCRRelRelWS, "R1C1 Relative column/Relative row WS column address generation is broken."
    Assert.AreEqual wbRef & rrRef, formCRRelRelWB, "R1C1 Relative column/Relative row WB column address generation is broken."

    Assert.AreEqual arRef, formCRAbsRel, "R1C1 Absolute column/Relative row address generation is broken."
    Assert.AreEqual wsRef & arRef, formCRAbsRelWS, "R1C1 Absolute column/Relative row WS column address generation is broken."
    Assert.AreEqual wbRef & arRef, formCRAbsRelWB, "R1C1 Absolute column/Relative row WB column address generation is broken."

    ' Act2:
    ' All should fail because row 0 is not allowed for absolute row references.

    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    TestColumn.AddressR1C1 Row, False, True
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 zero row is mistakenly allowed."
    Err.Clear

    TestColumn.AddressR1C1 Row, False, True, IncludeWS:=True
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 zero row is mistakenly allowed."
    Err.Clear

    TestColumn.AddressR1C1 Row, False, True, IncludeWB:=True
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 zero row is mistakenly allowed."
    Err.Clear

    TestColumn.AddressR1C1 Row, True, True
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 zero row is mistakenly allowed."
    Err.Clear

    TestColumn.AddressR1C1 Row, True, True, IncludeWS:=True
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 zero row is mistakenly allowed."
    Err.Clear

    TestColumn.AddressR1C1 Row, True, True, IncludeWB:=True
    Assert.AreEqual ExpectedError, Err.Number, "R1C1 zero row is mistakenly allowed."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Range Tests")
Public Sub Test_Range_Cell()
    On Error GoTo TestFail

    'Arrange:
    Dim TestWS As Worksheet
    Dim TestColumn As New SS_Column
    Dim rngOutput As Range
    Dim rngTest As Range

    Set TestWS = WorksheetFromCodeName(TestName)
    TestColumn.Init TestWS, 1
    With TestWS
        Set rngTest = .Range(.Cells(2, 1), .Cells(1000, 1))
    End With

    'Act:
    Set rngOutput = TestColumn.Range(2, 1000)

    'Assert:
    Assert.AreEqual rngTest.Address, rngOutput.Address, "Column range function broken."

    'Assert2:
    On Error GoTo TestError
    Set rngOutput = TestColumn.Range(1, ActiveSheet.Rows.Count + 1)
    Assert.Fail "Row Range out of bounds succeeding."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Exit Sub
TestError:
    If Err.Number = ERROR_SSC.SSCE_InvalParam Then _
        Exit Sub
    Assert.Fail "Row test assertion raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Content Tests")
Public Sub Test_Clear()
    On Error GoTo TestFail

    'Arrange1:
    Dim TestWS As Worksheet
    Dim TestColumn As New SS_Column
    Dim rngTest As Range
    Dim rngCell As Variant

    Set TestWS = WorksheetFromCodeName(TestName)
    Set rngTest = TestWS.Range("A2:A50")
    TestColumn.Init TestWS, 1

    TestColumn.Title = "Title Cell"
    For Each rngCell In rngTest
        rngCell.Value = RndString(RndInt(0, 20))
    Next rngCell

    'Act1:
    ' Shouldn't clear Title cell.
    TestColumn.ClearContents

    'Assert1:
    Assert.AreEqual "Title Cell", TestColumn.Title, "Title cell mistakenly cleared on param-less .ClearContents call."
    For Each rngCell In rngTest
        Assert.AreEqual vbNullString, rngCell.Value, "ClearContents not clearing all filled cells. Cell: " & rngCell.Address(False, False)
    Next rngCell

    'Arrange1:
    TestColumn.Title = "Title Cell"
    For Each rngCell In rngTest
        rngCell.Value = RndString(RndInt(0, 20))
    Next rngCell

    'Act2:
    ' Should clear Title cell now.
    TestColumn.ClearContents 1

    'Assert2:
    Assert.AreEqual vbNullString, TestColumn.Title, "Title cell not cleared on .ClearContents call incuding row 1."
    For Each rngCell In rngTest
        Assert.AreEqual vbNullString, rngCell.Value, "ClearContents not clearing all filled cells. Cell: " & rngCell.Address(False, False)
    Next rngCell

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Content Tests")
Public Sub Test_Fill()
    On Error GoTo TestFail

    'Arrange1:
    Dim TestWS As Worksheet
    Dim TestColumn As New SS_Column
    Dim TestString As String
    Dim rngTest As Range
    Dim rngCell As Variant

    Set TestWS = WorksheetFromCodeName(TestName)
    Set rngTest = TestWS.Range("A2:A50")
    TestColumn.Init TestWS, 1

    TestString = RndString(RndInt(5, 20))
    TestWS.Cells(2, 1).Value = TestString

    'Act1:
    TestColumn.FillDown 2, 50

    'Assert1:
    For Each rngCell In rngTest
        Assert.AreEqual TestString, rngCell.Value, "FillDown not filling all cells. Cell: " & rngCell.Address(False, False)
    Next rngCell

    'Arrange1:
    TestString = RndString(RndInt(5, 20))
    TestWS.Cells(50, 1).Value = TestString

    'Act2:
    TestColumn.FillUp 2, 50

    'Assert2:
    For Each rngCell In rngTest
        Assert.AreEqual TestString, rngCell.Value, "FillUp not filling all cells. Cell: " & rngCell.Address(False, False)
    Next rngCell

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
