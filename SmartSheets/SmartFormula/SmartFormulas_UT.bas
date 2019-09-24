Attribute VB_Name = "SmartFormulas_UT"
Option Explicit

'@TestModule
'@Folder("Tests.SmartSheets")

Private Assert As New Rubberduck.AssertClass
Private Test_SS As SmartSheet

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Test_SS = New SmartSheet
    Test_SS.Init pSheet:=ThisWorkbook.Worksheets(1)
    Test_SS.Add "TestCol", "TestCol2", "TestCol3"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    Set Test_SS = Nothing
End Sub

Private Function Permutate(ByVal sBegin As String, Optional ByVal sEnd As String, Optional ByRef arrPerms As Variant, Optional ByRef ixPerm As Long = 1, Optional ByVal permDepth As Long) As Variant
    Dim ixString As Long
    Dim lenEnd As Long

    lenEnd = Len(sBegin)
    
    If IsMissing(arrPerms) Then
        Dim cPerm As Long
        cPerm = 1
        For ixString = 1 To lenEnd
            cPerm = cPerm * ixString
        Next ixString
        ReDim arrPerms(1 To cPerm) As String
    End If

    If lenEnd < 2 Then
        arrPerms(ixPerm) = sEnd & sBegin
        ixPerm = ixPerm + 1
        Exit Function
    End If

    For ixString = 1 To lenEnd
        Permutate Left$(sBegin, ixString - 1) + Right$(sBegin, lenEnd - ixString), sEnd + Mid$(sBegin, ixString, 1), arrPerms, ixPerm, permDepth + 1
    Next ixString

    If permDepth = 0 Then _
        Permutate = arrPerms

End Function

'@TestMethod("Output Validation")
Public Sub Single_Cell_Formula()
    On Error GoTo TestFail
    
    'Arrange:
    Dim formrColrRow As String '  A2
    Dim formrColaRow As String '  A$2
    Dim formaColrRow As String ' $A2
    Dim formaColaRow As String ' $A$2

    'Act:
    formrColrRow = BuildFormula("=%cr", Test_SS("TestCol"), 2)
    formrColaRow = BuildFormula("=%cR", Test_SS("TestCol"), 2)
    formaColrRow = BuildFormula("=%Cr", Test_SS("TestCol"), 2)
    formaColaRow = BuildFormula("=%CR", Test_SS("TestCol"), 2)

    'Assert:
    Assert.AreEqual "=A2", formrColrRow, "Relative columns or relative rows broken."
    Assert.AreEqual "=A$2", formrColaRow, "Relative columns or absolute row broken."
    Assert.AreEqual "=$A2", formaColrRow, "Absolute columns or relative rows broken."
    Assert.AreEqual "=$A$2", formaColaRow, "Absolute columns or absolute rows broken."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Output Validation")
Public Sub WS_WB_Names()
    On Error GoTo TestFail
    
    'Arrange:
    Dim nameWB As String
    Dim nameWS As String

    Dim formWS As String
    Dim formWBWS As String
    nameWB = Test_SS.Workbook.name
    nameWS = Test_SS.Worksheet.name

    'Act:
    formWS = BuildFormula("=%sCr", Test_SS("TestCol"), 2)
    formWBWS = BuildFormula("=%bsCr", Test_SS("TestCol"), 2)

    'Assert:
    Assert.AreEqual "='" & nameWS & "'!$A2", formWS, "Worksheet flag not functioning."
    Assert.AreEqual "='[" & nameWB & "]" & nameWS & "'!$A2", formWBWS, "Workbook flag not functioning."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Output Validation")
Public Sub Certain_Orders_Irrelevant()
    On Error GoTo TestFail
    
    'Arrange:
    Dim nameWB As String
    Dim nameWS As String
    Dim permResCell As String
    Dim permResWS As String
    Dim permResWBWS As String
    Dim sPerm As Variant
    Dim argArray As Variant

    ' Permutation results of C, R, S, B
    nameWB = Test_SS.Workbook.name
    nameWS = Test_SS.Worksheet.name
    permResCell = "=A2"
    permResWS = "='" & nameWS & "'!A2"
    permResWBWS = "='[" & nameWB & "]" & nameWS & "'!A2"

    'Act/Assert:

    ' Cell Reference ordering
    For Each sPerm In Permutate("cr")
        If InStr(sPerm, "r") > InStr(sPerm, "c") Then
            argArray = Array(Test_SS("TestCol"), 2)
        Else
            argArray = Array(2, Test_SS("TestCol"))
        End If
        Assert.AreEqual permResCell, BuildFormulaFromArray("=%" & sPerm, argArray), "Cell reference flags irrelevant ordering broken."
    Next sPerm

    For Each sPerm In Permutate("scr")
        If InStr(sPerm, "r") > InStr(sPerm, "c") Then
            argArray = Array(Test_SS("TestCol"), 2)
        Else
            argArray = Array(2, Test_SS("TestCol"))
        End If
        Assert.AreEqual permResWS, BuildFormulaFromArray("=%" & sPerm, argArray), "Worksheet reference flags irrelevant ordering broken."
    Next sPerm
    
    For Each sPerm In Permutate("bscr")
        If InStr(sPerm, "r") > InStr(sPerm, "c") Then
            argArray = Array(Test_SS("TestCol"), 2)
        Else
            argArray = Array(2, Test_SS("TestCol"))
        End If
        Assert.AreEqual permResWBWS, BuildFormulaFromArray("=%" & sPerm, argArray), "Workbook reference flags irrelevant ordering broken."
    Next sPerm

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Output Validation")
Public Sub Reference_Escaping()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sescC As String, sescR As String, sescL As String
    Dim sescS As String, sescB As String
    Dim sescPerc As String, sescSlash As String

    'Act:
    sescC = BuildFormula("=%Cr\c", Test_SS("TestCol"), 2)
    sescR = BuildFormula("=%Cr\r", Test_SS("TestCol"), 2)
    sescL = BuildFormula("=%Cr\l", Test_SS("TestCol"), 2)
    sescS = BuildFormula("=%Cr\s", Test_SS("TestCol"), 2)
    sescB = BuildFormula("=%Cr\b", Test_SS("TestCol"), 2)
    sescPerc = BuildFormula("=%%%Cr%%", Test_SS("TestCol"), 2)
    sescSlash = BuildFormula("=%Cr\\", Test_SS("TestCol"), 2)

    'Assert:
    Assert.AreEqual "=$A2c", sescC, "C flag escaping is broken."
    Assert.AreEqual "=$A2r", sescR, "R flag escaping is broken."
    Assert.AreEqual "=$A2l", sescL, "L flag escaping is broken."
    Assert.AreEqual "=$A2s", sescS, "S flag escaping is broken."
    Assert.AreEqual "=$A2b", sescB, "B flag escaping is broken."
    Assert.AreEqual "=%$A2%", sescPerc, "Percent flag escaping is broken."
    Assert.AreEqual "=$A2\", sescSlash, "Slash escaping is broken."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Output Validation")
Public Sub Half_References()
    On Error GoTo TestFail
    
    'Arrange:
    Dim colAbsRef As String, rowAbsRef As String
    Dim colRelRef As String, rowRelRef As String
    Dim colWSRef As String, rowWSRef As String
    Dim colWBRef As String, rowWBRef As String
    Dim nameWB As String, nameWS As String
    nameWB = Test_SS.Workbook.name
    nameWS = Test_SS.Worksheet.name

    'Act:
    colAbsRef = BuildFormula("=%C", Test_SS("TestCol"))
    rowAbsRef = BuildFormula("=%R", 2)

    colRelRef = BuildFormula("=%c", Test_SS("TestCol"))
    rowRelRef = BuildFormula("=%r", 2)

    colWSRef = BuildFormula("=%sC", Test_SS("TestCol"))
    rowWSRef = BuildFormula("=%sR", Test_SS, 2)

    colWBRef = BuildFormula("=%sbC", Test_SS("TestCol"))
    rowWBRef = BuildFormula("=%sbR", Test_SS, 2)

    'Assert:
    Assert.AreEqual "=$A:$A", colAbsRef, "Single column absolute references broken."
    Assert.AreEqual "=$2:$2", rowAbsRef, "Single row absolute references broken."

    Assert.AreEqual "=A:A", colRelRef, "Single column relative references broken."
    Assert.AreEqual "=2:2", rowRelRef, "Single row relative references broken."

    Assert.AreEqual "='" & nameWS & "'!$A:$A", colWSRef, "Worksheet column references broken."
    Assert.AreEqual "='" & nameWS & "'!$2:$2", rowWSRef, "Worksheet row references broken."

    Assert.AreEqual "='[" & nameWB & "]" & nameWS & "'!$A:$A", colWBRef, "Workbook column references broken."
    Assert.AreEqual "='[" & nameWB & "]" & nameWS & "'!$2:$2", rowWBRef, "Workbook row references broken."

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

Private Function Assert_On_Success(ByVal sFormat As String, Optional ByVal Params As Variant, Optional ByVal AssertMsg As String) As String
    Assert_On_Success = BuildFormulaFromArray(sFormat, Params)
    If AssertMsg <> vbNullString Then
        Assert.Fail AssertMsg
    Else
        Assert.Fail "Improper BuildFormula format string(""" & sFormat & """) succeeded at building."
    End If
End Function

'@TestMethod("Input Validation")
Public Sub Bad_FmtStr_Input()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpectedError As Long
    Dim arrArgs As Variant

    arrArgs = Array(Test_SS("TestCol"), 2)

    'Act:
    ExpectedError = ERROR_BF.BFE_BadChar
    Assert_On_Success "=%bbCr", arrArgs, AssertMsg:="Duplicate ""b"" flag is passing."
    Assert_On_Success "=%ssCr", arrArgs, AssertMsg:="Duplicate ""s"" flag is passing."
    Assert_On_Success "=%llCr", arrArgs, AssertMsg:="Duplicate ""l"" flag is passing."
    Assert_On_Success "=%cCr", arrArgs, AssertMsg:="Duplicate ""c"" flag is passing."
    Assert_On_Success "=%Crr", arrArgs, AssertMsg:="Duplicate ""r"" flag is passing."

    ExpectedError = ERROR_BF.BFE_BadFlag
    Assert_On_Success "=%lf", arrArgs, AssertMsg:="Bad flag after ""l"" flag is passing."

    ExpectedError = ERROR_BF.BFE_MissFlag
    Assert_On_Success "=%cl", arrArgs, AssertMsg:="Missing flag after ""l"" flag is passing."

    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then _
        Resume Next
    Err.Clear
End Sub

'@TestMethod("Input Validation")
Public Sub Bad_Array_Input()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ExpectedError As Long

    'Act:
    ExpectedError = ERROR_BF.BFE_BadCol
    Assert_On_Success "=%Cr", Array(2, 5), AssertMsg:="Invalid SS_Column parameter is passing."
    Assert_On_Success "=INDEX(%C,%rC)", Array(Test_SS("TestCol"), 2), AssertMsg:="Missing SS_Column parameter is passing."

    ExpectedError = ERROR_BF.BFE_BadRow
    Assert_On_Success "=%r", Array(0), AssertMsg:="Invalid row parameter is passing."
    Assert_On_Success "=%r", Array(-1), AssertMsg:="Invalid row parameter is passing."
    Assert_On_Success "=%Cr", Array(Test_SS("TestCol")), AssertMsg:="Missing row parameter is passing."

    ExpectedError = ERROR_BF.BFE_BadSS
    Assert_On_Success "=%sR", Array(Test_SS("TestCol"), 2), AssertMsg:="Invalid SmartSheet parameter is passing."

    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then _
        Resume Next
    Err.Clear
End Sub

