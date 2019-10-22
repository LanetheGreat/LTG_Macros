VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SS_Column"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("SmartSheets")

'@Ignore ConstantNotUsed
Const Version As String = "1.0.0"

Private Const ClassName As String = "SS_Column"

Private Const descSSCE_ObjUninit As String = ClassName & " object not initialized."
Private Const descSSCE_ObjInit As String = ClassName & " has already been initialized."

Private ParentSheet As Worksheet
Private ColumnName As String
Private ColumnNumber As Long

Public Enum ERROR_SSC
    [_start] = vbObjectError + 512
    SSCE_ObjUninit  ' Class not initialized
    SSCE_ObjInit    ' Class already initialized
    SSCE_InvalParam ' Invalid parameter or property assignment
End Enum

' Events triggered by the SS_Column Class

Private Sub Class_Terminate()
    Set ParentSheet = Nothing
End Sub


' Private routines used by the SS_Column Class

Private Function Num2Char(ByVal Number As Long) As String
    If Number < 1 Or Number > 26 Then Exit Function
    Num2Char = ChrW$(64 + Number)
End Function

Private Function CalculateName(ByVal ColNumber As Long) As String
    Const n As Long = 26
    Dim cN1 As Long
    Dim cN2 As Long
    Dim cN3 As Long
    Dim cN4 As Long

    cN1 = (ColNumber - 1) Mod n + 1

    If ColNumber > n Then _
        cN2 = Fix((ColNumber - (n + 1)) / n) Mod n + 1

    If ColNumber > n * n + n Then _
        cN3 = Fix((ColNumber - (n * (n + 1) + 1)) / (n * n)) Mod n + 1

    If ColNumber > n * n * n + (n * n + n) Then _
        cN4 = Fix((ColNumber - (n * (n * (n + 1) + 1) + 1)) / (n * n * n)) Mod n + 1

    CalculateName = Num2Char(cN4) & Num2Char(cN3) & Num2Char(cN2) & Num2Char(cN1)
End Function

Private Sub CheckInit(Optional ByVal Src As String)

    If ParentSheet Is Nothing Then
        If Src = vbNullString Then
            Src = ClassName
        Else
            Src = ClassName & "." & Src
        End If

        Err.Raise ERROR_SSC.SSCE_ObjUninit, Src, descSSCE_ObjUninit
    End If

End Sub

Private Sub CheckRow(ByVal Row As Long, Optional ByVal Src As String)

    If Row < 1 Or Row > ParentSheet.Rows.Count Then
        If Src = vbNullString Then
            Src = ClassName
        Else
            Src = ClassName & "." & Src
        End If

        Err.Raise ERROR_SSC.SSCE_InvalParam, Src, "Invalid parameter: " & Row & " is not a valid row number."
    End If

End Sub

Private Sub CheckColumn(ByVal Column As Variant, Optional ByVal Src As String, Optional ByVal WS As Worksheet)

    If WS Is Nothing Then _
        Set WS = ParentSheet

    On Error GoTo ERROR_InvalidColumn
    Set Column = WS.Columns(Column)
    Exit Sub

ERROR_InvalidColumn:
    Err.Clear
    If Src = vbNullString Then
        Src = ClassName
    Else
        Src = ClassName & "." & Src
    End If
    Err.Raise ERROR_SSC.SSCE_InvalParam, Src, "Invalid parameter: " & Column & " is not a valid column value."
End Sub

Private Function ExternalRef(ByRef IncludeWS As Boolean, ByVal IncludeWB As Boolean) As String

    If IncludeWB Then
        ExternalRef = "'[" & ParentSheet.Parent.Name & "]"
        IncludeWS = True
    ElseIf IncludeWS Then
        ExternalRef = "'"
    End If

    If IncludeWS Then _
        ExternalRef = ExternalRef & Replace(ParentSheet.Name, "'", "''") & "'!"

End Function

Private Function privLastRow() As Long
    With ParentSheet.UsedRange.Rows
        privLastRow = .Row + .Count - 1
    End With
End Function


' Properties supported by the SS_Column Class

' Property: IsInit (Read only)
Public Property Get IsInit() As Boolean
    IsInit = Not ParentSheet Is Nothing
End Property

' Property: Name (Read Only)
'@DefaultMember
Public Property Get Name() As String
Attribute Name.VB_UserMemId = 0
    CheckInit "Name"
    Name = ColumnName
End Property

' Property: Number (Read only)
Public Property Get Number() As Long
    CheckInit "Number"
    Number = ColumnNumber
End Property

' Property: Hidden (Read/Write)
Public Property Get Hidden() As Boolean
    CheckInit "Hidden"
    Hidden = ParentSheet.Columns(ColumnNumber).EntireColumn.Hidden
End Property

Public Property Let Hidden(ByVal Hide As Boolean)
    CheckInit "Hidden"
    ParentSheet.Columns(ColumnNumber).EntireColumn.Hidden = Hide
End Property

' Property: Title (Read/Write)
Public Property Get Title() As String
    CheckInit "Title"
    Title = CStr(ParentSheet.Cells(1, ColumnNumber).Value)
End Property

Public Property Let Title(ByVal NewTitle As String)
    CheckInit "Title"
    ParentSheet.Cells(1, ColumnNumber).Value = NewTitle
End Property

' Property: TitleCell (Read only)
Public Property Get TitleCell() As Range
    CheckInit "TitleCell"
    Set TitleCell = ParentSheet.Cells(1, ColumnNumber)
End Property

' Property: Column (Read only)
Public Property Get Column() As Range
    CheckInit "Column"
    Set Column = ParentSheet.Columns(ColumnNumber)
End Property

' Property: LastRow (Read only)
Public Property Get LastRow() As Long
    CheckInit "LastRow"

    With ParentSheet.Cells(ParentSheet.Rows.Count, ColumnNumber)
        If Not IsEmpty(.Value) Then
            LastRow = .Row
        Else
            LastRow = .End(xlUp).Row
        End If
    End With
End Property


' Public routines provided by the SS_Column Class

Public Sub Init(ByVal Parent As Variant, ByVal Column As Variant)

    If Not ParentSheet Is Nothing Then _
        Err.Raise ERROR_SSC.SSCE_ObjInit, ClassName & ".Init", descSSCE_ObjInit

    If TypeOf Parent Is Workbook Then
        Set Parent = Parent.ActiveSheet
    ElseIf Not TypeOf Parent Is Worksheet Then
        Err.Raise ERROR_SSC.SSCE_InvalParam, ClassName & ".Init", "Parent is not a valid Workbook or Worksheet."
    End If

    CheckColumn Column, "Init", Parent

    ColumnNumber = Parent.Columns(Column).Column
    ColumnName = CalculateName(ColumnNumber)
    Set ParentSheet = Parent

End Sub

Public Sub ShiftLeft()

    CheckInit "ShiftLeft"

    If ColumnNumber = 1 Then _
        Err.Raise 9, ClassName & ".ShiftLeft"

    ColumnNumber = ColumnNumber - 1
    ColumnName = CalculateName(ColumnNumber)

End Sub

Public Sub ShiftRight()

    CheckInit "ShiftRight"

    If ColumnNumber = ParentSheet.Columns.Count Then _
        Err.Raise 9, ClassName & ".ShiftRight"

    ColumnNumber = ColumnNumber + 1
    ColumnName = CalculateName(ColumnNumber)

End Sub

Public Sub SetIndex(ByVal ColumnIndex As Variant)

    CheckInit "SetIndex"
    CheckColumn Column, "SetIndex"

    ColumnNumber = ParentSheet.Columns(ColumnIndex).Column
    ColumnName = CalculateName(ColumnNumber)

End Sub

Public Function ColumnAddress( _
    Optional ByVal Absolute As Boolean = False, _
    Optional ByVal IncludeWS As Boolean = False, _
    Optional ByVal IncludeWB As Boolean = False _
) As String

    CheckInit "ColumnAddress"
    ColumnAddress = ExternalRef(IncludeWS, IncludeWB)

    If Absolute Then
        ColumnAddress = ColumnAddress & "$" & ColumnName & ":$" & ColumnName
    Else
        ColumnAddress = ColumnAddress & ColumnName & ":" & ColumnName
    End If

End Function

Public Function RowAddress( _
    ByVal Row As Long, _
    Optional ByVal AbsColumn As Boolean = False, _
    Optional ByVal AbsRow As Boolean = False, _
    Optional ByVal IncludeWS As Boolean = False, _
    Optional ByVal IncludeWB As Boolean = False _
) As String

    CheckInit "RowAddress"
    CheckRow Row, "RowAddress"
    RowAddress = ExternalRef(IncludeWS, IncludeWB)

    If AbsColumn Then _
        RowAddress = RowAddress & "$"
    RowAddress = RowAddress & ColumnName

    If AbsRow Then _
        RowAddress = RowAddress & "$"
    RowAddress = RowAddress & Row

End Function

Public Function RangeAddress( _
    ByVal StartRow As Long, _
    ByVal EndRow As Long, _
    Optional ByVal AbsColumn As Boolean = False, _
    Optional ByVal AbsRow As Boolean = False, _
    Optional ByVal IncludeWS As Boolean = False, _
    Optional ByVal IncludeWB As Boolean = False _
) As String

    CheckInit "RangeAddress"
    CheckRow StartRow, "RangeAddress"
    CheckRow EndRow, "RangeAddress"
    RangeAddress = ExternalRef(IncludeWS, IncludeWB)

    If AbsColumn Then _
        RangeAddress = RangeAddress & "$"
    RangeAddress = RangeAddress & ColumnName

    If AbsRow Then _
        RangeAddress = RangeAddress & "$"
    RangeAddress = RangeAddress & StartRow

    If StartRow <> EndRow Then
        RangeAddress = RangeAddress & ":"

        If AbsColumn Then _
            RangeAddress = RangeAddress & "$"
        RangeAddress = RangeAddress & ColumnName

        If AbsRow Then _
            RangeAddress = RangeAddress & "$"
        RangeAddress = RangeAddress & EndRow
    End If

End Function

Public Function AddressC1( _
    Optional ByVal Absolute As Boolean = False, _
    Optional ByVal IncludeWS As Boolean = False, _
    Optional ByVal IncludeWB As Boolean = False _
) As String

    CheckInit "AddressC1"
    AddressC1 = ExternalRef(IncludeWS, IncludeWB) & "C"

    If Absolute Then _
        AddressC1 = AddressC1 & ColumnNumber

End Function

Public Function AddressR1C1( _
    Optional ByVal Row As Long, _
    Optional ByVal AbsColumn As Boolean = False, _
    Optional ByVal AbsRow As Boolean = False, _
    Optional ByVal IncludeWS As Boolean = False, _
    Optional ByVal IncludeWB As Boolean = False _
) As String

    CheckInit "AddressR1C1"
    AddressR1C1 = ExternalRef(IncludeWS, IncludeWB)

    AddressR1C1 = AddressR1C1 & "R"
    If AbsRow Then
        CheckRow Row, "AddressR1C1"
        AddressR1C1 = AddressR1C1 & Row
    ElseIf Row <> 0 Then
        AddressR1C1 = AddressR1C1 & "[" & Row & "]"
    End If

    AddressR1C1 = AddressR1C1 & "C"
    If AbsColumn Then _
        AddressR1C1 = AddressR1C1 & ColumnNumber

End Function

Public Function Cell(ByVal Row As Long) As Range

    CheckInit "Cell"
    CheckRow Row, "Cell"

    Set Cell = ParentSheet.Cells(Row, ColumnNumber)

End Function

Public Function Range(Optional ByVal StartRow As Long = 2, Optional ByVal EndRow As Variant) As Range

    CheckInit "Range"
    CheckRow StartRow, "Range"

    If IsMissing(EndRow) Then _
        EndRow = privLastRow()
    CheckRow EndRow, "Range"

    Set Range = ParentSheet.Range(ColumnName & StartRow & ":" & ColumnName & EndRow)

End Function

'TODO: Check that StartRow > Endrow (Also Unit Test)
Public Sub ClearContents(Optional ByVal StartRow As Long = 2, Optional ByVal EndRow As Variant)

    CheckInit "ClearContents"
    CheckRow StartRow, "ClearContents"

    If IsMissing(EndRow) Then _
        EndRow = privLastRow()
    CheckRow EndRow, "ClearContents"

    With ParentSheet
        .Range(.Cells(StartRow, ColumnNumber), .Cells(EndRow, ColumnNumber)).ClearContents
    End With

End Sub

'TODO: Check that StartRow > Endrow (Also Unit Test)
Public Sub FillDown(Optional ByVal StartRow As Long = 2, Optional ByVal EndRow As Variant)

    CheckInit "FillDown"
    CheckRow StartRow, "FillDown"

    If IsMissing(EndRow) Then _
        EndRow = privLastRow()
    CheckRow EndRow, "FillDown"

    With ParentSheet
        .Range(.Cells(StartRow, ColumnNumber), .Cells(EndRow, ColumnNumber)).FillDown
    End With

End Sub

Public Sub FillUp(Optional ByVal StartRow As Long = 2, Optional ByVal EndRow As Variant)

    CheckInit "FillUp"
    CheckRow StartRow, "FillUp"

    If IsMissing(EndRow) Then _
        EndRow = privLastRow()
    CheckRow EndRow, "FillUp"

    With ParentSheet
        .Range(.Cells(StartRow, ColumnNumber), .Cells(EndRow, ColumnNumber)).FillUp
    End With

End Sub

'TODO: Add Unit Tests for .Delete subroutine.
Public Sub Delete()

    CheckInit "Delete"

    ParentSheet.Columns(ColumnNumber).Delete
    Set ParentSheet = Nothing

End Sub

