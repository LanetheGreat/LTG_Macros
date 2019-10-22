VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SmartSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("SmartSheets")
'References "Microsoft Scripting Runtime"
'References "Microsoft VBScript Regular Expressions 5.5"

'@Ignore ConstantNotUsed
Const Version As String = "1.0.0"

Private Const ClassName As String = "SmartSheet"
Private Const descSSE_ObjUninit As String = ClassName & "object not initialized."
Private Const descSSE_ObjInit As String = ClassName & " has already been initialized."

Private ParentSheet As Worksheet
Private SmartColumns As Scripting.Dictionary
Private NextColumn As Long

Public Enum ERROR_SS
    [_start] = vbObjectError + 512
    SSE_ObjUninit
    SSE_ObjInit
    SSE_InvalParam
    SSE_LabelExists
    SSE_LabelNotFound
End Enum


' Private routines used by the SmartSheet Class

Private Sub CheckInit(Optional ByVal Src As String)

    If ParentSheet Is Nothing Then
        If Src = vbNullString Then
            Src = ClassName
        Else
            Src = ClassName & "." & Src
        End If

        Err.Raise ERROR_SS.SSE_ObjUninit, Src, descSSE_ObjUninit
    End If

End Sub

Private Sub CheckExists(ByVal Label As String, Optional ByVal Src As String)

    If Not SmartColumns.Exists(Label) Then
        If Src = vbNullString Then
            Src = ClassName
        Else
            Src = ClassName & "." & Src
        End If

        Err.Raise ERROR_SS.SSE_LabelNotFound, Src, "Label """ & Label & """ doesn't exist in " & ClassName & "."
    End If

End Sub

Private Sub CheckNotExists(ByVal Label As String, Optional ByVal Src As String)

    If SmartColumns.Exists(Label) Then
        If Src = vbNullString Then
            Src = ClassName
        Else
            Src = ClassName & "." & Src
        End If

        Err.Raise ERROR_SS.SSE_LabelExists, Src, "Label """ & Label & """ already exists in " & ClassName & "."
    End If

End Sub

Private Sub CheckColumnNumber(ByVal ColumnNumber As Long, Optional ByVal Src As String)

    Dim Label As Variant
    For Each Label In SmartColumns.Keys
        If SmartColumns(Label).Number = ColumnNumber Then
            If Src = vbNullString Then
                Src = ClassName
            Else
                Src = ClassName & "." & Src
            End If

            Err.Raise ERROR_SS.SSE_LabelExists, Src, "Column " & ColumnNumber & " already has a label. (""" & Label & """)"
        End If
    Next Label

    On Error GoTo ERROR_InvalidColumn

    Dim Column As Range
    Set Column = ParentSheet.Columns(ColumnNumber)

    Exit Sub

ERROR_InvalidColumn:
    Err.Clear
    If Src = vbNullString Then
        Src = ClassName
    Else
        Src = ClassName & "." & Src
    End If
    Err.Raise ERROR_SSC.SSCE_InvalParam, Src, "Invalid parameter: " & ColumnNumber & " is not a valid column value."

End Sub

Private Sub CheckRow(ByVal Row As Long, Optional ByVal Src As String)

    If Row < 1 Or Row > ParentSheet.rows.Count Then
        If Src = vbNullString Then
            Src = ClassName
        Else
            Src = ClassName & "." & Src
        End If

        Err.Raise ERROR_SS.SSE_InvalParam, Src, "Invalid parameter: " & Row & " is not a valid row number."
    End If

End Sub

Private Sub UpdateNextColumn()
    Dim NewNext As Long
    Dim ColumnNumber As Long
    Dim Label As Variant

    NewNext = 1
    For Each Label In SmartColumns
        ColumnNumber = SmartColumns(Label).Number
        If ColumnNumber > NewNext Then _
            NewNext = ColumnNumber
    Next Label
    NextColumn = NewNext

End Sub

Private Sub privAddFromArray(ByVal Labels As Variant, ByVal Src As String)

    Dim Label As Variant

    For Each Label In Labels
        If VarType(Label) <> vbString Then _
            Err.Raise SSE_InvalParam, ClassName & "." & Src, """" & TypeName(Label) & """ is not allowed as a label."

        CheckNotExists Label, Src
    Next Label

    Dim Column As SS_Column
    For Each Label In Labels
        Set Column = New SS_Column
        Column.Init ParentSheet, NextColumn

        SmartColumns.Add Label, Column

        NextColumn = NextColumn + 1
    Next Label

End Sub

' Creates a mapping of columns headers to their column number. (Ignores extra columns with duplicate headers though)
Private Sub privCacheTitles(ByRef TitleCache As Scripting.Dictionary)

    Dim Index As Long
    Dim Title As String

    For Index = 1 To ParentSheet.UsedRange.Columns.Count
        Title = CStr(ParentSheet.Cells(1, Index).Value)
        If Not TitleCache.Exists(Title) Then _
            TitleCache.Add Title, Index
    Next Index

End Sub

Private Function privExternalRef(ByRef IncludeWS As Boolean, ByVal IncludeWB As Boolean) As String

    If IncludeWB Then
        privExternalRef = "'[" & ParentSheet.Parent.Name & "]"
        IncludeWS = True
    ElseIf IncludeWS Then
        privExternalRef = "'"
    End If

    If IncludeWS Then _
        privExternalRef = privExternalRef & Replace(ParentSheet.Name, "'", "''") & "'!"

End Function


' Events triggered by the SS_Column Class

Private Sub Class_Initialize()
    Set SmartColumns = New Scripting.Dictionary
    NextColumn = 1
End Sub

Private Sub Class_Terminate()
    Set SmartColumns = Nothing
    Set ParentSheet = Nothing
End Sub


' Properties supported by the SmartSheet Class

' Property: Workbook (Read only)
Public Property Get Workbook() As Workbook
    CheckInit "Workbook"
    Set Workbook = ParentSheet.Parent
End Property

' Property: Worksheet (Read only)
Public Property Get Worksheet() As Worksheet
    CheckInit "Worksheet"
    Set Worksheet = ParentSheet
End Property

' Property: Count (Read only)
Public Property Get Count() As Long
    CheckInit "Count"
    Count = SmartColumns.Count
End Property

' Property: Labels (Read only)
Public Property Get Labels() As Variant
    CheckInit "Labels"
    Labels = SmartColumns.Keys
End Property

' Property: FirstRow (Read only)
Public Property Get FirstRow() As Long
    CheckInit "FirstRow"
    FirstRow = ParentSheet.UsedRange.rows.Row
End Property

' Property: LastRow (Read only)
Public Property Get LastRow() As Long
    CheckInit "LastRow"
    With ParentSheet.UsedRange.rows
        LastRow = .Row + .Count - 1
    End With
End Property

' Property: FirstColumn (Read only)
Public Property Get FirstColumn() As SS_Column
    CheckInit "FirstColumn"

    Dim Column As Variant
    For Each Column In SmartColumns.Items

        If FirstColumn Is Nothing Then
            Set FirstColumn = Column
        Else
            If Column.Number < FirstColumn.Number Then _
                Set FirstColumn = Column
        End If

    Next Column

End Property

' Property: FirstLabel (Read only)
Public Property Get FirstLabel() As String
    CheckInit "FirstLabel"

    Dim Label As Variant

    For Each Label In SmartColumns.Keys

        If FirstLabel = vbNullString Then
            FirstLabel = Label
        Else
            If SmartColumns(Label).Number < SmartColumns(FirstLabel).Number Then _
                FirstLabel = Label
        End If

    Next Label

End Property

' Property: LastColumn (Read only)
Public Property Get LastColumn() As SS_Column
    CheckInit "LastColumn"

    Dim Column As Variant
    For Each Column In SmartColumns.Items

        If LastColumn Is Nothing Then
            Set LastColumn = Column
        Else
            If Column.Number > LastColumn.Number Then _
                Set LastColumn = Column
        End If

    Next Column

End Property

' Property: LastLabel (Read only)
Public Property Get LastLabel() As String
    CheckInit "LastLabel"

    Dim Label As Variant

    For Each Label In SmartColumns.Keys

        If LastLabel = vbNullString Then
            LastLabel = Label
        Else
            If SmartColumns(Label).Number < SmartColumns(LastLabel).Number Then _
                LastLabel = Label
        End If

    Next Label

End Property

' Worksheet Wrapper Properties

' Property: Number (Read only)
Public Property Get Name() As String
    CheckInit "Name"
    Name = ParentSheet.Name
End Property

Public Property Let Name(ByVal NewName As String)
    CheckInit "Name"
    ParentSheet.Name = NewName
End Property


' Public routines provided by the SmartSheet Class

' Initialization Method (Must be called before the object can be used)
Public Sub Init(ByVal Parent As Variant)

    If Not ParentSheet Is Nothing Then _
        Err.Raise ERROR_SS.SSE_ObjInit, ClassName & ".Init", descSSE_ObjInit

    If TypeOf Parent Is Workbook Then
        Set Parent = Parent.ActiveSheet
    ElseIf Not TypeOf Parent Is Worksheet Then
        Err.Raise ERROR_SSC.SSCE_InvalParam, ClassName & ".Init", "Parent is not a valid Workbook or Worksheet."
    End If

    Set ParentSheet = Parent
    'RndKeySeed = -Timer()
End Sub


' Public routines provided by the SmartSheet Class

'@DefaultMember
Public Function Item(ByVal Label As String) As SS_Column
Attribute Item.VB_UserMemId = 0

    CheckInit "Item"
    CheckExists Label, "Item"

    Set Item = SmartColumns.Item(Label)

End Function

Public Function HasLabel(ByVal Label As String) As Boolean

    CheckInit "HasLabel"

    HasLabel = SmartColumns.Exists(Label)

End Function

Public Function HasAll(ByVal Labels As Variant) As Boolean

    CheckInit "HasAll"

    Dim Label As Variant
    For Each Label In Labels
        If Not SmartColumns.Exists(Label) Then _
            Exit Function
    Next Label
    HasAll = True

End Function

Public Function HasAny(ByVal Labels As Variant) As Boolean

    CheckInit "HasAny"

    Dim Label As Variant
    For Each Label In Labels
        If SmartColumns.Exists(Label) Then
            HasAny = True
            Exit For
        End If
    Next Label

End Function

Public Function Clone(ByVal Parent As Worksheet) As SmartSheet

    CheckInit "Clone"

    Dim Label As Variant
    Set Clone = New SmartSheet
    Clone.Init Parent

    For Each Label In SmartColumns.Keys
        Clone.AddByIndex Label, SmartColumns(Label).Number
    Next Label

End Function

Public Sub AddFromArray(ByVal Labels As Variant)

    CheckInit "AddFromArray"
    privAddFromArray Labels, "AddFromArray"

End Sub

Public Sub Add(ParamArray Labels())

    CheckInit "Add"
    privAddFromArray CVar(Labels), "Add"

End Sub

Public Sub AddByIndex(ByVal Label As String, ByVal Index As Long)

    CheckInit "AddByIndex"
    CheckNotExists Label, "AddByIndex"
    CheckColumnNumber Index, "AddByIndex"

    Dim Column As New SS_Column
    Column.Init ParentSheet, Index

    SmartColumns.Add Label, Column
    UpdateNextColumn

End Sub

Public Sub InsertByIndex(ByVal Label As String, ByVal Index As Long)

    CheckInit "InsertByIndex"
    CheckNotExists Label, "InsertByIndex"

    Dim Column As Variant

    ParentSheet.Columns(Index).Insert Shift:=xlToRight

    For Each Column In SmartColumns.Items
        If Column.Number >= Index Then Column.ShiftRight
    Next Column

    Set Column = New SS_Column
    Column.Init ParentSheet, Index

    SmartColumns.Add Label, Column
    UpdateNextColumn

End Sub

Public Sub InsertAfter(ByVal InsertLabel As String, ByVal AfterLabel As String)

    CheckInit "InsertAfter"
    CheckNotExists InsertLabel, "InsertAfter"
    CheckExists AfterLabel, "InsertAfter"

    Dim Index As Long
    Dim Column As Variant

    Index = SmartColumns(AfterLabel).Number + 1
    ParentSheet.Columns(Index).Insert Shift:=xlToRight

    For Each Column In SmartColumns.Items
        If Column.Number >= Index Then Column.ShiftRight
    Next Column

    Set Column = New SS_Column
    Column.Init ParentSheet, Index

    SmartColumns.Add InsertLabel, Column
    UpdateNextColumn

End Sub

Public Sub InsertBefore(ByVal InsertLabel As String, ByVal BeforeLabel As String)

    CheckInit "InsertBefore"
    CheckNotExists InsertLabel, "InsertBefore"
    CheckExists BeforeLabel, "InsertBefore"

    Dim Index As Long
    Dim Column As Variant

    Index = SmartColumns(BeforeLabel).Number
    ParentSheet.Columns(Index).Insert Shift:=xlToRight

    For Each Column In SmartColumns.Items
        If Column.Number >= Index Then Column.ShiftRight
    Next Column

    Set Column = New SS_Column
    Column.Init ParentSheet, Index

    SmartColumns.Add InsertLabel, Column
    UpdateNextColumn

End Sub

Public Sub MoveLabel(ByVal Label As String, ByVal Index As Long)

    CheckInit "MoveLabel"
    CheckExists Label, "MoveLabel"
    CheckColumnNumber Index, "MoveLabel"

    Dim MovingColumn As SS_Column
    Dim Column As Variant

    Set MovingColumn = SmartColumns(Label)
    If MovingColumn.Number = Index Then Exit Sub

    ParentSheet.Columns(MovingColumn.Number).Cut
    If Index < MovingColumn.Number Then
        ParentSheet.Columns(Index).Insert Shift:=xlToRight
    Else
        ParentSheet.Columns(Index + 1).Insert Shift:=xlToRight
    End If

    For Each Column In SmartColumns.Items

        If Not Column Is MovingColumn Then
            If Index > MovingColumn.Number Then
                If Column.Number <= Index And Column.Number > MovingColumn.Number Then Column.ShiftLeft
            Else
                If Column.Number >= Index And Column.Number < MovingColumn.Number Then Column.ShiftRight
            End If
        End If

    Next Column

    MovingColumn.SetIndex Index
    UpdateNextColumn

End Sub

Public Sub MoveAfter(ByVal MoveLabel As String, ByVal AfterLabel As String)

    CheckInit "MoveAfter"
    CheckExists MoveLabel, "MoveAfter"
    CheckExists AfterLabel, "MoveAfter"

    Dim MovingColumn As SS_Column
    Dim AfterIndex As Long
    Dim Column As Variant

    Set MovingColumn = SmartColumns(MoveLabel)
    AfterIndex = SmartColumns(AfterLabel).Number

    If MoveLabel = AfterLabel Or MovingColumn.Number = AfterIndex + 1 Then _
        Exit Sub

    ParentSheet.Columns(MovingColumn.Number).Cut
    ParentSheet.Columns(AfterIndex + 1).Insert Shift:=xlToRight

    For Each Column In SmartColumns.Items

        If Not Column Is MovingColumn Then
            If AfterIndex > MovingColumn.Number Then
                If Column.Number <= AfterIndex And Column.Number > MovingColumn.Number Then Column.ShiftLeft
            Else
                If Column.Number > AfterIndex And Column.Number < MovingColumn.Number Then Column.ShiftRight
            End If
        End If

    Next Column

    If AfterIndex > MovingColumn.Number Then
        MovingColumn.SetIndex AfterIndex
    Else
        MovingColumn.SetIndex AfterIndex + 1
    End If

    UpdateNextColumn

End Sub

Public Sub MoveBefore(ByVal MoveLabel As String, ByVal BeforeLabel As String)

    CheckInit "MoveBefore"
    CheckExists MoveLabel, "MoveBefore"
    CheckExists BeforeLabel, "MoveBefore"

    Dim MovingColumn As SS_Column
    Dim AfterIndex As Long
    Dim Column As Variant

    Set MovingColumn = SmartColumns(MoveLabel)
    AfterIndex = SmartColumns(BeforeLabel).Number

    If MoveLabel = BeforeLabel Or MovingColumn.Number = AfterIndex - 1 Then _
        Exit Sub

    ParentSheet.Columns(MovingColumn.Number).Cut
    ParentSheet.Columns(AfterIndex).Insert Shift:=xlToRight

    For Each Column In SmartColumns.Items

        If Not Column Is MovingColumn Then
            If AfterIndex > MovingColumn.Number Then
                If Column.Number < AfterIndex And Column.Number > MovingColumn.Number Then Column.ShiftLeft
            Else
                If Column.Number >= AfterIndex And Column.Number < MovingColumn.Number Then Column.ShiftRight
            End If
        End If

    Next Column

    If AfterIndex > MovingColumn.Number Then
        MovingColumn.SetIndex AfterIndex - 1
    Else
        MovingColumn.SetIndex AfterIndex
    End If

    UpdateNextColumn

End Sub

Public Sub ReverseMapTitles(ByVal TitleMap As Scripting.Dictionary)
    CheckInit "ReverseMapTitles"

    Dim Label As Variant
    Dim Column As SS_Column

    Dim TitleCache As New Scripting.Dictionary
    privCacheTitles TitleCache

    For Each Label In TitleMap.Keys
        If Not SmartColumns.Exists(Label) Then
            If TitleCache.Exists(TitleMap(Label)) Then
                Set Column = New SS_Column
                Column.Init ParentSheet, TitleCache(TitleMap(Label))
                SmartColumns.Add Label, Column
            End If
        End If
    Next Label

    UpdateNextColumn

End Sub

Public Function InsertMissingByTitles(ByVal TitleMap As Scripting.Dictionary, Optional ByVal Offset As Long = 1, Optional ByVal Include As Variant, Optional ByVal Exclude As Variant) As Collection
    CheckInit "InsertMissingByTitles"
    CheckColumnNumber Offset, "InsertMissingByTitles"

    Set InsertMissingByTitles = New Collection

    Dim TitleCache As New Scripting.Dictionary
    privCacheTitles TitleCache

    Dim Index As Long
    Dim Label As Variant
    Dim FilterLabel As Variant
    Dim Column As Variant
    Dim addLabel As Boolean
    Index = Offset

    For Each Label In TitleMap.Keys
        If Not SmartColumns.Exists(Label) Then
            If Not TitleCache.Exists(TitleMap(Label)) Then

                If Not IsMissing(Include) Then
                    addLabel = False
                    For Each FilterLabel In Include
                        If Label = FilterLabel Then
                            addLabel = True
                            Exit For
                        End If
                    Next FilterLabel

                ElseIf Not IsMissing(Exclude) Then
                    addLabel = True
                    For Each FilterLabel In Exclude
                        If Label = FilterLabel Then
                            addLabel = False
                            Exit For
                        End If
                    Next FilterLabel

                Else
                    addLabel = True
                End If

                If addLabel Then

                    For Each Column In SmartColumns.Items
                        If Column.Number >= Index Then _
                            Column.ShiftRight
                    Next Column

                    ParentSheet.Columns(Index).Insert Shift:=xlToRight
                    ParentSheet.Columns(Index).validation.Delete

                    Set Column = New SS_Column
                    Column.Init ParentSheet, Index
                    SmartColumns.Add Label, Column

                    Column.Title = TitleMap(Label)
                    TitleCache.Add TitleMap(Label), Index
                    InsertMissingByTitles.Add Label

                End If
            End If
        End If

        Index = Index + 1
    Next Label

    UpdateNextColumn

End Function

'TODO: Check for StartRow > EndRow. (Also Unit Test)
Public Function RangeAddress( _
    ByVal StartLabel As String, _
    ByVal EndLabel As String, _
    Optional ByVal StartRow As Long = 1, _
    Optional ByVal EndRow As Variant, _
    Optional ByVal AbsColumn As Boolean = False, _
    Optional ByVal AbsRow As Boolean = False, _
    Optional ByVal IncludeWS As Boolean = False, _
    Optional ByVal IncludeWB As Boolean = False _
) As String

    CheckInit "RangeAddress"
    CheckExists StartLabel, "RangeAddress"
    CheckExists EndLabel, "RangeAddress"
    CheckRow StartRow, "RangeAddress"

    If IsMissing(EndRow) Then
        EndRow = Me.LastRow
    Else
        EndRow = CLng(EndRow)
    End If
    CheckRow EndRow, "RangeAddress"

    RangeAddress = privExternalRef(IncludeWS, IncludeWB)

    If AbsColumn Then _
        RangeAddress = RangeAddress & "$"
    RangeAddress = RangeAddress & SmartColumns(StartLabel).Name

    If AbsRow Then _
        RangeAddress = RangeAddress & "$"
    RangeAddress = RangeAddress & StartRow

    If StartRow <> EndRow Or StartLabel <> EndLabel Then
        RangeAddress = RangeAddress & ":"

        If AbsColumn Then _
            RangeAddress = RangeAddress & "$"
        RangeAddress = RangeAddress & SmartColumns(EndLabel).Name

        If AbsRow Then _
            RangeAddress = RangeAddress & "$"
        RangeAddress = RangeAddress & EndRow
    End If

End Function

Public Function RowAddress(ByVal Row As Long, Optional ByVal Absolute As Boolean = False, Optional ByVal IncludeWS As Boolean = False, Optional ByVal IncludeWB As Boolean = False) As String

    CheckInit "RowAddress"
    CheckRow Row, "RowAddress"

    RowAddress = privExternalRef(IncludeWS, IncludeWB)

    If Absolute Then
        RowAddress = RowAddress & "$" & RowAddress & ":$" & RowAddress
    Else
        RowAddress = RowAddress & RowAddress & ":" & RowAddress
    End If

End Function

'TODO: Check for StartRow > EndRow (Also Unit Test)
Public Sub FillDownFromTo( _
    Optional ByVal StartLabel As Variant, _
    Optional ByVal EndLabel As Variant, _
    Optional ByVal StartRow As Long = 2, _
    Optional ByVal EndRow As Variant _
)

    CheckInit "FillDownFromTo"

    If IsMissing(StartLabel) Then StartLabel = Me.FirstLabel _
    Else: StartLabel = CStr(StartLabel)

    If IsMissing(EndLabel) Then EndLabel = Me.LastLabel _
    Else: EndLabel = CStr(EndLabel)

    If IsMissing(EndRow) Then EndRow = Me.LastRow _
    Else: EndRow = CLng(EndRow)

    CheckExists StartLabel, "FillDownFromTo"
    CheckExists EndLabel, "FillDownFromTo"
    CheckRow StartRow, "FillDownFromTo"
    CheckRow EndRow, "FillDownFromTo"

    If StartRow <> EndRow Then

        Dim Column As Variant
        Dim StartIndex As Long
        Dim EndIndex As Long

        StartIndex = SmartColumns(StartLabel).Number
        EndIndex = SmartColumns(EndLabel).Number

        If StartIndex > EndIndex Then
            Dim TempIndex As Long
            TempIndex = StartIndex
            StartIndex = EndIndex
            EndIndex = TempIndex
        End If

        For Each Column In SmartColumns.Items
            If StartIndex <= Column.Number And Column.Number <= EndIndex Then
                With ParentSheet
                    .Range(.Cells(StartRow, Column.Number), .Cells(EndRow, Column.Number)).FillDown
                End With
            End If
        Next Column

    End If

End Sub

Public Sub FillDownArray(ByVal Labels As Variant, Optional ByVal StartRow As Long = 2, Optional ByVal EndRow As Variant)

    CheckInit "FillDownArray"

    If IsMissing(EndRow) Then EndRow = Me.LastRow _
    Else: EndRow = CLng(EndRow)

    CheckRow StartRow, "FillDownArray"
    CheckRow EndRow, "FillDownArray"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "FillDownArray"
    Next Label

    For Each Label In Labels
        SmartColumns(CStr(Label)).FillDown StartRow, EndRow
    Next Label

End Sub

Private Function privCheckDynLng(ByVal Arr As Variant, ByRef Index As Long, ByRef Value As Variant, Optional ByVal Default As Variant) As Boolean
    If Index <= UBound(Arr) Then
        If IsNumeric(Arr(Index)) Then
            If VarType(Arr(Index)) <> vbString Then
                Value = CLng(Arr(Index))
                Index = Index + 1
                privCheckDynLng = True
            End If
        ElseIf Not IsMissing(Default) Then
            Value = Default
        End If
    ElseIf Not IsMissing(Default) Then
        Value = Default
    End If
End Function

Public Sub FillDownFrom(ParamArray Labels() As Variant)

    CheckInit "FillDownFrom"

    Dim KeyIndex As Long
    Dim StartRow As Long
    Dim EndRow As Long

    KeyIndex = LBound(Labels)
    If privCheckDynLng(Labels, KeyIndex, StartRow, 2) Then
        privCheckDynLng Labels, KeyIndex, EndRow, Me.LastRow
    Else
        EndRow = Me.LastRow
    End If

    CheckRow StartRow, "FillDownFrom"
    CheckRow EndRow, "FillDownFrom"

    Dim OldIndex As Long
    OldIndex = KeyIndex

    For KeyIndex = OldIndex To UBound(Labels)
        CheckExists CStr(Labels(KeyIndex)), "FillDownFrom"
    Next KeyIndex

    For KeyIndex = OldIndex To UBound(Labels)
        SmartColumns(CStr(Labels(KeyIndex))).FillDown StartRow, EndRow
    Next KeyIndex

End Sub

Public Sub FillDown(Optional ByVal StartRow As Long = 2, Optional ByVal EndRow As Variant)

    CheckInit "FillDown"

    If IsMissing(EndRow) Then EndRow = Me.LastRow _
    Else: EndRow = CLng(EndRow)

    CheckRow StartRow, "FillDown"
    CheckRow EndRow, "FillDown"

    Dim Column As Variant
    For Each Column In SmartColumns.Items
        Column.FillDown StartRow, EndRow
    Next Column

End Sub

Public Sub RemoveLabelArray(ByVal Labels As Variant)

    CheckInit "RemoveLabelArray"

    If SmartColumns.Count = 0 Then Exit Sub

    Dim Label As Variant
    For Each Label In Labels
        Label = CStr(Label)
        If SmartColumns.Exists(Label) Then _
            SmartColumns.Remove Label
    Next Label

    UpdateNextColumn

End Sub

Public Sub RemoveLabels(ParamArray Labels())

    CheckInit "RemoveLabels"

    If SmartColumns.Count = 0 Then Exit Sub

    Dim Label As Variant
    For Each Label In Labels
        Label = CStr(Label)
        If SmartColumns.Exists(Label) Then _
            SmartColumns.Remove Label
    Next Label

    UpdateNextColumn

End Sub

Public Sub RemoveLabel(ByVal Label As String)

    CheckInit "RemoveLabel"

    If SmartColumns.Count = 0 Then Exit Sub

    Label = CStr(Label)
    If SmartColumns.Exists(Label) Then _
        SmartColumns.Remove Label

    UpdateNextColumn

End Sub

Public Sub ChangeLabel(ByVal OldLabel As String, ByVal NewLabel As String)

    CheckInit "ChangeLabel"
    CheckExists OldLabel, "ChangeLabel"
    CheckNotExists NewLabel, "ChangeLabel"

    If StrComp(OldLabel, NewLabel) = 0 Then Exit Sub

    Set SmartColumns(NewLabel) = SmartColumns(OldLabel)
    SmartColumns.Remove OldLabel

End Sub

Public Sub KeepLabelArray(ByVal Labels As Variant)

    CheckInit "KeepLabelArray"

    If SmartColumns.Count = 0 Then Exit Sub

    Dim KeepLabel As Variant
    Dim Label As Variant
    Dim keep As Boolean

    For Each Label In SmartColumns.Keys
        keep = False
        For Each KeepLabel In Labels
            If StrComp(Label, KeepLabel) = 0 Then
                keep = True
                Exit For
            End If
        Next KeepLabel

        If Not keep Then SmartColumns.Remove Label
    Next Label

    UpdateNextColumn

End Sub

Public Sub KeepLabels(ParamArray Labels())

    CheckInit "KeepLabels"

    If SmartColumns.Count = 0 Then Exit Sub

    Dim KeepLabel As Variant
    Dim Label As Variant
    Dim keep As Boolean

    For Each Label In SmartColumns.Keys
        keep = False
        For Each KeepLabel In Labels
            If StrComp(Label, KeepLabel) = 0 Then
                keep = True
                Exit For
            End If
        Next KeepLabel

        If Not keep Then SmartColumns.Remove Label
    Next Label

    UpdateNextColumn

End Sub

Public Sub KeepLabel(ByVal KeepLabel As String)

    CheckInit "KeepLabel"

    If SmartColumns.Count = 0 Then Exit Sub

    Dim Label As Variant
    For Each Label In SmartColumns.Keys
        If StrComp(Label, KeepLabel) <> 0 Then _
            SmartColumns.Remove Label
    Next Label

    UpdateNextColumn

End Sub

Public Sub Destroy()

    If Not ParentSheet Is Nothing Then
        Dim OldAlerts As Boolean

        OldAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False

        ParentSheet.Delete
        Set ParentSheet = Nothing
        Set SmartColumns = New Scripting.Dictionary

        Application.DisplayAlerts = OldAlerts
    End If

End Sub

Public Sub DeleteColumn(ByVal Label As String)

    CheckInit "DeleteColumn"
    CheckExists Label, "DeleteColumn"

    Dim ColumnIndex As Long
    Dim Column As Variant

    ColumnIndex = SmartColumns(Label).Number
    ParentSheet.Columns(ColumnIndex).Delete Shift:=xlToLeft
    SmartColumns.Remove Label

    For Each Column In SmartColumns.Items
        If Column.Number > ColumnIndex Then Column.ShiftLeft
    Next Column

    UpdateNextColumn

End Sub

Public Sub SetNumberFormatsArray(ByVal FormatStr As String, ByVal Labels As Variant)

    CheckInit "SetNumberFormatsArray"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "SetNumberFormatsArray"
    Next Label

    For Each Label In Labels
        SmartColumns(CStr(Label)).Column.NumberFormat = FormatStr
    Next Label

End Sub

Public Sub SetNumberFormats(ByVal FormatStr As String, ParamArray Labels())

    CheckInit "SetNumberFormats"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "SetNumberFormats"
    Next Label

    For Each Label In Labels
        SmartColumns(CStr(Label)).Column.NumberFormat = FormatStr
    Next Label

End Sub

Public Sub SetStylesArray(ByVal StyleStr As String, ByVal Labels As Variant)

    CheckInit "SetStylesArray"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "SetStylesArray"
    Next Label

    For Each Label In SmartColumns.Keys
        SmartColumns(CStr(Label)).Column.Style = StyleStr
    Next Label

End Sub

Public Sub SetStyles(ByVal StyleStr As String, ParamArray Labels())

    CheckInit "SetStyles"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "SetStyles"
    Next Label

    For Each Label In SmartColumns.Keys
        SmartColumns(CStr(Label)).Column.Style = StyleStr
    Next Label

End Sub

Public Sub SetValidationArray( _
    ByVal ValidType As XlDVType, _
    Optional ByVal AlertStyle As XlDVAlertStyle, _
    Optional ByVal Operator As XlFormatConditionOperator, _
    Optional ByVal Formula1 As String, _
    Optional ByVal Formula2 As String, _
    Optional ByVal Labels As Variant _
)

    CheckInit "SetValidationArray"

    If IsEmpty(Labels) Then Exit Sub
    If AlertStyle = 0 Then AlertStyle = xlValidAlertInformation

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "SetValidationArray"
    Next Label

    For Each Label In Labels

        With SmartColumns(CStr(Label)).Column.validation
            .Delete
            .Add ValidType, AlertStyle, Operator, Formula1, Formula2
        End With

        SmartColumns(CStr(Label)).Cell(1).validation.Delete

    Next Label

End Sub

Public Sub SetTitleBoldArray(ByVal Bolden As Boolean, ByVal Labels As Variant)

    CheckInit "SetTitleBoldArray"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "SetTitleBoldArray"
    Next Label

    For Each Label In Labels
        SmartColumns(CStr(Label)).TitleCell.Font.Bold = Bolden
    Next Label

End Sub

Public Sub SetTitleBold(ByVal Bolden As Boolean, ParamArray Labels())

    CheckInit "SetTitleBold"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "SetTitleBoldArray"
    Next Label

    For Each Label In Labels
        SmartColumns(CStr(Label)).TitleCell.Font.Bold = Bolden
    Next Label

End Sub

Public Sub DeleteUnlabeled()

    CheckInit "DeleteUnlabeled"

    Dim IndexCache As Scripting.Dictionary
    Dim Column As Variant
    Set IndexCache = New Scripting.Dictionary

    ' Create a cache of our column index for fast referencing
    For Each Column In SmartColumns.Items
        IndexCache.Add Column.Number, Empty
    Next

    Dim EndIndex As Long
    Dim Index As Long

    Index = 1
    EndIndex = ParentSheet.UsedRange.Columns.Count

    Do While Index <= EndIndex

        If Not IndexCache.Exists(Index) Then
            ParentSheet.Columns(Index).Delete Shift:=xlToLeft

            For Each Column In SmartColumns.Items
                If Column.Number > Index Then Column.ShiftLeft
            Next Column

            EndIndex = EndIndex - 1
        Else
            Index = Index + 1
        End If
    Loop

    UpdateNextColumn

End Sub

Public Function Row(ByVal Index As Long) As Range
    CheckInit "Row"
    Set Row = ParentSheet.rows(Index)
End Function

Public Sub DeleteLastRow()
    CheckInit "DeleteLastRow"
    ParentSheet.rows(Me.LastRow).Delete Shift:=xlUp
End Sub

Public Sub SortByColumns(ByVal Labels As Variant, Optional ByVal Directions As Variant)

    CheckInit "SortByColumns"

    Dim DirIndex As Long
    If Not IsMissing(Directions) Then _
        DirIndex = LBound(Directions)

    Dim Label As Variant
    Dim LabelIndex As Long
    Dim ColumnIndex As Long
    Dim Direction As XlSortOrder

    For Each Label In Labels
        CheckExists CStr(Label), "SortByColumns"
    Next Label

    ParentSheet.Sort.SortFields.Clear
    For LabelIndex = LBound(Labels) To UBound(Labels)

        Direction = xlAscending
        If Not IsMissing(Directions) Then
            If DirIndex <= UBound(Directions) Then
                Direction = Directions(DirIndex)
                DirIndex = DirIndex + 1
            End If
        End If

        ColumnIndex = SmartColumns(CStr(Labels(LabelIndex))).Number
        ParentSheet.Sort.SortFields.Add _
            key:=ParentSheet.Range(ParentSheet.Cells(2, ColumnIndex), ParentSheet.Cells(Me.LastRow, ColumnIndex)), _
            SortOn:=xlSortOnValues, _
            Order:=Direction, _
            DataOption:=xlSortTextAsNumbers

    Next LabelIndex

    With ParentSheet.Sort
        .SetRange ParentSheet.Range(ParentSheet.Cells(1, Me.FirstColumn.Number), ParentSheet.Cells(Me.LastRow, Me.LastColumn.Number))
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

' NOTE: Uses sorting as an optimization, will modify row ordering.
Public Sub DeleteRowsByValue(ByVal Label As String, ByVal Value As Variant)

    CheckInit "DeleteRowsByValue"
    CheckExists Label, "DeleteRowsByValue"

    Dim FirstCell As Range
    Dim LastCell As Range
    Dim Column As SS_Column

    Me.SortByColumns Array(Label)
    Set Column = SmartColumns(Label)

    Set FirstCell = Column.Column.Find( _
        What:=Value, _
        After:=Column.Cell(1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False _
    )

    If FirstCell Is Nothing Then _
        Exit Sub

    '@Ignore ExcelMemberMayReturnNothing
    Set LastCell = Column.Column.Find( _
        What:=Value, _
        After:=Column.Cell(1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False, _
        SearchFormat:=False _
    )

    ParentSheet.Range(FirstCell, LastCell).EntireRow.Delete Shift:=xlUp

End Sub

Public Sub DeleteRowsByValues(ByVal Label As String, ByVal Values As Variant, Optional ByVal DelErrors As Boolean = True)

    CheckInit "DeleteRowsByValues"
    CheckExists Label, "DeleteRowsByValues"

    Dim FirstCell As Range
    Dim LastCell As Range
    Dim Column As SS_Column
    Dim Value As Variant
    Dim LastRow As Long

    Me.SortByColumns Array(Label)
    Set Column = SmartColumns(Label)

    If DelErrors Then
        Set FirstCell = Nothing

        On Error Resume Next
        Set FirstCell = Column.Column.SpecialCells(xlCellTypeFormulas, xlErrors)
        On Error GoTo 0

        If Not FirstCell Is Nothing Then _
            FirstCell.EntireRow.Delete Shift:=xlUp

    End If

    LastRow = Me.LastRow

    For Each Value In Values
        Set FirstCell = Column.Column.Find( _
            What:=Value, _
            After:=Column.Cell(1), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            SearchFormat:=False _
        )

        If Not FirstCell Is Nothing Then

            '@Ignore ExcelMemberMayReturnNothing
            Set LastCell = Column.Column.Find( _
                What:=Value, _
                After:=Column.Cell(1), _
                LookIn:=xlValues, _
                LookAt:=xlWhole, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False, _
                SearchFormat:=False _
            )

            ParentSheet.Range(FirstCell, LastCell).EntireRow.Delete Shift:=xlUp
        End If

    Next Value

End Sub

' NOTE: Uses sorting as an optimization, will modify row ordering.
Public Sub DeleteRowsByRegex(ByVal Label As String, ByVal Expression As String, Optional ByVal IgnoreCase As Boolean = True)

    CheckInit "DeleteRowsByRegex"
    CheckExists Label, "DeleteRowsByRegex"

    Dim FirstCell As Range
    Dim LastCell As Range
    Dim Row As Long
    Dim LastRow As Long
    Dim Regex As New RegExp

    Dim TempColumn As New SS_Column
    With TempColumn
        .Init ParentSheet, ParentSheet.UsedRange.Columns.Count + 1
        .Title = "Filter Column"
        .Column.NumberFormat = "General"
    End With

    On Error GoTo RegexDelete_Error

    LastRow = Me.LastRow

    With Regex
        .Pattern = Expression
        .IgnoreCase = IgnoreCase
    End With

    For Row = 2 To LastRow
        TempColumn.Cell(Row).Value = Regex.Test(CStr(SmartColumns(Label).Cell(Row).Value))
    Next Row

    With ParentSheet
        .Sort.SortFields.Clear
        .Sort.SortFields.Add _
            key:=.Range(.Cells(2, TempColumn.Number), .Cells(Me.LastRow, TempColumn.Number)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortTextAsNumbers
    End With

    With ParentSheet.Sort
        .SetRange ParentSheet.UsedRange
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Set FirstCell = TempColumn.Column.Find( _
        What:=True, _
        After:=TempColumn.Cell(1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False _
    )

    If FirstCell Is Nothing Then _
        GoTo RegexDelete_Done

    '@Ignore ExcelMemberMayReturnNothing
    Set LastCell = TempColumn.Column.Find( _
        What:=True, _
        After:=TempColumn.Cell(1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False, _
        SearchFormat:=False _
    )
    ParentSheet.Range(FirstCell, LastCell).EntireRow.Delete Shift:=xlUp

RegexDelete_Done:
    Err.Clear
RegexDelete_Error:

    TempColumn.Delete

    If Err.Number <> 0 Then _
        Err.Raise Err.Number

End Sub

Public Sub KeepRowsByValue(ByVal Label As String, ByVal Value As Variant)

    CheckInit "KeepRowsByValue"
    CheckExists Label, "KeepRowsByValue"

    Dim FirstCell As Range
    Dim LastCell As Range
    Dim Column As SS_Column
    Dim LastRow As Long

    Me.SortByColumns Array(Label)
    Set Column = SmartColumns(Label)
    LastRow = Me.LastRow

    Set FirstCell = Column.Column.Find( _
        What:=Value, _
        After:=Column.Cell(1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False _
    )

    If FirstCell Is Nothing Then
        If LastRow <> 1 Then _
            ParentSheet.rows("2:" & LastRow).Delete
        Exit Sub
    End If

    '@Ignore ExcelMemberMayReturnNothing
    Set LastCell = Column.Column.Find( _
        What:=Value, _
        After:=Column.Cell(1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False, _
        SearchFormat:=False _
    )

    If FirstCell.Row = 2 Then
        ParentSheet.Range(ParentSheet.Cells(LastCell.Row + 1, Column.Number), ParentSheet.Cells(LastRow, Column.Number)).EntireRow.Delete Shift:=xlUp
    ElseIf LastCell.Row = LastRow Then
        ParentSheet.Range(ParentSheet.Cells(2, Column.Number), ParentSheet.Cells(FirstCell.Row - 1, Column.Number)).EntireRow.Delete Shift:=xlUp
    Else
        ParentSheet.Range(ParentSheet.Cells(LastCell.Row + 1, Column.Number), ParentSheet.Cells(LastRow, Column.Number)).EntireRow.Delete Shift:=xlUp
        ParentSheet.Range(ParentSheet.Cells(2, Column.Number), ParentSheet.Cells(FirstCell.Row - 1, Column.Number)).EntireRow.Delete Shift:=xlUp
    End If

End Sub

Public Sub KeepRowsByValues(ByVal Label As String, ByVal Values As Variant, Optional ByVal DelErrors As Boolean = True)

    CheckInit "KeepRowsByValues"
    CheckExists Label, "KeepRowsByValues"

    Dim FirstCell As Range
    Dim LastCell As Range
    Dim Column As SS_Column
    Dim Value As Variant
    Dim LastRow As Long

    Me.SortByColumns Array(Label)
    Set Column = SmartColumns(Label)

    If DelErrors Then
        Set FirstCell = Nothing

        On Error Resume Next
        Set FirstCell = Column.Column.SpecialCells(xlCellTypeFormulas, xlErrors)
        On Error GoTo 0

        If Not FirstCell Is Nothing Then _
            FirstCell.EntireRow.Delete Shift:=xlUp

    End If

    LastRow = Me.LastRow

    For Each Value In Values
        Set FirstCell = Column.Column.Find( _
            What:=Value, _
            After:=Column.Cell(1), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            SearchFormat:=False _
        )

        If FirstCell Is Nothing Then
            If LastRow <> 1 Then _
                ParentSheet.rows("2:" & LastRow).Delete
            Exit Sub
        End If

        '@Ignore ExcelMemberMayReturnNothing
        Set LastCell = Column.Column.Find( _
            What:=Value, _
            After:=Column.Cell(1), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False, _
            SearchFormat:=False _
        )

        If FirstCell.Row = 2 Then
            ParentSheet.Range(ParentSheet.Cells(LastCell.Row + 1, Column.Number), ParentSheet.Cells(LastRow, Column.Number)).EntireRow.Delete Shift:=xlUp
        ElseIf LastCell.Row = LastRow Then
            ParentSheet.Range(ParentSheet.Cells(2, Column.Number), ParentSheet.Cells(FirstCell.Row - 1, Column.Number)).EntireRow.Delete Shift:=xlUp
        Else
            ParentSheet.Range(ParentSheet.Cells(LastCell.Row + 1, Column.Number), ParentSheet.Cells(LastRow, Column.Number)).EntireRow.Delete Shift:=xlUp
            ParentSheet.Range(ParentSheet.Cells(2, Column.Number), ParentSheet.Cells(FirstCell.Row - 1, Column.Number)).EntireRow.Delete Shift:=xlUp
        End If

    Next Value

End Sub

Public Sub KeepRowsByRegEx(ByVal Label As String, ByVal Expression As String, Optional ByVal IgnoreCase As Boolean = True)

    CheckInit "KeepRowsByRegEx"
    CheckExists Label, "KeepRowsByRegEx"

    Dim FirstCell As Range
    Dim LastCell As Range
    Dim Row As Long
    Dim LastRow As Long
    Dim Regex As New RegExp

    Dim TempColumn As New SS_Column
    With TempColumn
        .Init ParentSheet, ParentSheet.UsedRange.Columns.Count + 1
        .Title = "Filter Column"
        .Column.NumberFormat = "General"
    End With

    On Error GoTo RegexDelete_Error

    LastRow = Me.LastRow

    With Regex
        .Pattern = Expression
        .IgnoreCase = IgnoreCase
    End With

    For Row = 2 To LastRow
        TempColumn.Cell(Row).Value = Regex.Test(CStr(SmartColumns(Label).Cell(Row).Value))
    Next Row

    With ParentSheet
        .Sort.SortFields.Clear
        .Sort.SortFields.Add _
            key:=.Range(.Cells(2, TempColumn.Number), .Cells(Me.LastRow, TempColumn.Number)), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortTextAsNumbers
    End With

    With ParentSheet.Sort
        .SetRange ParentSheet.UsedRange
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Set FirstCell = TempColumn.Column.Find( _
        What:=True, _
        After:=TempColumn.Cell(1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False _
    )

    If FirstCell Is Nothing Then
        If LastRow <> 1 Then _
            ParentSheet.rows("2:" & LastRow).Delete
        GoTo RegexDelete_Done
    End If

    '@Ignore ExcelMemberMayReturnNothing
    Set LastCell = TempColumn.Column.Find( _
        What:=True, _
        After:=TempColumn.Cell(1), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False, _
        SearchFormat:=False _
    )

    If FirstCell.Row = 2 Then
        ParentSheet.Range(ParentSheet.Cells(LastCell.Row + 1, TempColumn.Number), ParentSheet.Cells(LastRow, TempColumn.Number)).EntireRow.Delete Shift:=xlUp
    ElseIf LastCell.Row = LastRow Then
        ParentSheet.Range(ParentSheet.Cells(2, TempColumn.Number), ParentSheet.Cells(FirstCell.Row - 1, TempColumn.Number)).EntireRow.Delete Shift:=xlUp
    Else
        ParentSheet.Range(ParentSheet.Cells(LastCell.Row + 1, TempColumn.Number), ParentSheet.Cells(LastRow, TempColumn.Number)).EntireRow.Delete Shift:=xlUp
        ParentSheet.Range(ParentSheet.Cells(2, TempColumn.Number), ParentSheet.Cells(FirstCell.Row - 1, TempColumn.Number)).EntireRow.Delete Shift:=xlUp
    End If

RegexDelete_Done:
    Err.Clear
RegexDelete_Error:

    TempColumn.Delete

    If Err.Number <> 0 Then _
        Err.Raise Err.Number

End Sub

Public Sub FlattenColumnArray(ByVal Labels As Variant)

    CheckInit "FlattenColumnArray"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "FlattenColumnArray"
    Next Label

    For Each Label In Labels
        With SmartColumns(CStr(Label)).Column
            .Formula = .Value
        End With
    Next Label

End Sub

Public Sub FlattenColumns(ParamArray Labels())

    CheckInit "FlattenColumns"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "FlattenColumns"
    Next Label

    For Each Label In Labels
        With SmartColumns(CStr(Label)).Column
            .Formula = .Value
        End With
    Next Label

End Sub

Public Sub FlattenRangesArray(ByVal Labels As Variant)

    CheckInit "FlattenRangesArray"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "FlattenRangesArray"
    Next Label

    For Each Label In Labels
        With SmartColumns(CStr(Label)).Range
            .Formula = .Value
        End With
    Next Label

End Sub

Public Sub FlattenRanges(ParamArray Labels())

    CheckInit "FlattenRanges"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "FlattenRanges"
    Next Label

    For Each Label In Labels
        With SmartColumns(CStr(Label)).Range
            .Formula = .Value
        End With
    Next Label

End Sub

Public Sub HideColumnArray(ByVal Labels As Variant)

    CheckInit "HideColumnArray"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "HideColumnArray"
    Next Label

    For Each Label In Labels
        SmartColumns(CStr(Label)).Hidden = True
    Next Label

End Sub

Public Sub HideColumns(ParamArray Labels())

    CheckInit "HideColumns"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "HideColumns"
    Next Label

    For Each Label In Labels
        SmartColumns(CStr(Label)).Hidden = True
    Next Label

End Sub

Public Sub UnhideColumnArray(ByVal Labels As Variant)

    CheckInit "UnhideColumnArray"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "UnhideColumnArray"
    Next Label

    For Each Label In Labels
        SmartColumns(CStr(Label)).Hidden = False
    Next Label

End Sub

Public Sub UnhideColumns(ParamArray Labels())

    CheckInit "UnhideColumns"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "UnhideColumns"
    Next Label

    For Each Label In Labels
        SmartColumns(CStr(Label)).Hidden = False
    Next Label

End Sub

Public Sub AutoFitColumnArray(ByVal Labels As Variant)

    CheckInit "AutoFitColumnArray"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "AutoFitColumnArray"
    Next Label

    For Each Label In Labels
        With SmartColumns(CStr(Label))
            If .Hidden Then
                .Column.AutoFit
                .Hidden = True
            Else
                .Column.AutoFit
            End If
        End With
    Next Label

End Sub

Public Sub AutoFitColumns(ParamArray Labels())

    CheckInit "AutoFitColumns"

    Dim Label As Variant
    For Each Label In Labels
        CheckExists CStr(Label), "AutoFitColumns"
    Next Label

    For Each Label In Labels
        With SmartColumns(CStr(Label))
            If .Hidden Then
                .Column.AutoFit
                .Hidden = True
            Else
                .Column.AutoFit
            End If
        End With
    Next Label

End Sub

Public Sub RemoveDuplicatesArray(ByVal Individual As Boolean, ByVal Labels As Variant)

    CheckInit "RemoveDuplicatesArray"

    Dim Label As Variant

    If Individual Then
        For Each Label In Labels
            CheckExists CStr(Label), "RemoveDuplicatesArray"
        Next Label

        For Each Label In Labels
            ParentSheet.UsedRange.RemoveDuplicates Array(SmartColumns(CStr(Label)).Number), xlYes
        Next Label

    Else
        Dim ColumnIndex As Long
        Dim FirstColumn As Long
        Dim Indexes() As Long
        ReDim Indexes(LBound(Labels) To UBound(Labels))

        ColumnIndex = LBound(Labels)
        FirstColumn = Me.FirstColumn.Number

        For Each Label In Labels
            Label = CStr(Label)

            CheckExists Label, "RemoveDuplicatesArray"
            Indexes(ColumnIndex) = SmartColumns(Label).Number - FirstColumn + 1

            ColumnIndex = ColumnIndex + 1
        Next Label

        ParentSheet.UsedRange.RemoveDuplicates Indexes, xlYes
    End If

End Sub

Public Sub RemoveDuplicates(ByVal Individual As Boolean, ParamArray Labels())

    CheckInit "RemoveDuplicates"

    Dim Label As Variant

    If Individual Then
        For Each Label In Labels
            CheckExists CStr(Label), "RemoveDuplicates"
        Next Label

        For Each Label In Labels
            ParentSheet.UsedRange.RemoveDuplicates Array(SmartColumns(CStr(Label)).Number), xlYes
        Next Label

    Else
        Dim ColumnIndex As Long
        Dim FirstColumn As Long
        Dim Indexes() As Long
        ReDim Indexes(LBound(Labels) To UBound(Labels))

        ColumnIndex = LBound(Labels)
        FirstColumn = Me.FirstColumn.Number

        For Each Label In Labels
            Label = CStr(Label)

            CheckExists Label, "RemoveDuplicates"
            Indexes(ColumnIndex) = SmartColumns(Label).Number - FirstColumn + 1

            ColumnIndex = ColumnIndex + 1
        Next Label

        ParentSheet.UsedRange.RemoveDuplicates Indexes, xlYes
    End If

End Sub

Public Sub AddBordersArray(ByVal Borders As Variant, ByVal LineStyle As XlLineStyle, ByVal ColorIndex As XlColorIndex, ByVal TintShade As Long, ByVal Weight As XlBorderWeight, ByVal Labels As Variant)

    CheckInit "AddBordersArray"

    Dim Label As Variant

    For Each Label In Labels
        CheckExists CStr(Label), "AddBordersArray"
    Next Label

    For Each Label In Labels
        Label = CStr(Label)

        If IsArray(Borders) Then
            Dim bType As Variant

            For Each bType In Borders
                With SmartColumns(Label).Column.Borders(bType)
                    .LineStyle = LineStyle
                    .ColorIndex = ColorIndex
                    .TintAndShade = TintShade
                    .Weight = Weight
                End With
            Next bType

        Else
            With SmartColumns(Label).Column.Borders(Borders)
                .LineStyle = LineStyle
                .ColorIndex = ColorIndex
                .TintAndShade = TintShade
                .Weight = Weight
            End With
        End If

    Next Label

End Sub

Public Sub AddBorders(ByVal Borders As Variant, ByVal LineStyle As XlLineStyle, ByVal ColorIndex As XlColorIndex, ByVal TintShade As Long, ByVal Weight As XlBorderWeight, ParamArray Labels())

    CheckInit "AddBorders"

    Dim Label As Variant

    For Each Label In Labels
        CheckExists CStr(Label), "AddBorders"
    Next Label

    For Each Label In Labels
        Label = CStr(Label)

        If IsArray(Borders) Then
            Dim bType As Variant

            For Each bType In Borders
                With SmartColumns(Label).Column.Borders(bType)
                    .LineStyle = LineStyle
                    .ColorIndex = ColorIndex
                    .TintAndShade = TintShade
                    .Weight = Weight
                End With
            Next bType

        Else
            With SmartColumns(Label).Column.Borders(Borders)
                .LineStyle = LineStyle
                .ColorIndex = ColorIndex
                .TintAndShade = TintShade
                .Weight = Weight
            End With
        End If

    Next Label

End Sub

Public Sub FreezeCells(Optional ByVal Row As Long, Optional ByVal Column As Variant)

    If IsMissing(Column) Then _
        Column = 0

    ParentSheet.Activate

    With ActiveWindow
        If .WindowState = xlMinimized Then .WindowState = xlNormal
        .ScrollRow = Row
        .ScrollColumn = Column
        .SplitRow = Row
        .SplitColumn = Column
        .FreezePanes = True
    End With

    ParentSheet.rows(1).Font.Bold = True

End Sub


