Attribute VB_Name = "SmartFormulas"
Option Explicit

'@Folder("SmartSheets.SmartFormulas")

Public Enum ERROR_BF
    [_start] = vbObjectError + 512
    BFE_HandleRef
    BFE_BadChar
    BFE_BadFlag
    BFE_BadCol
    BFE_BadRow
    BFE_BadSS
    BFE_MissFlag
End Enum

' Supported format specifiers:
'  - % : Starts processing a cell reference (escape using double sign %%)
'  - c or C: Column Reference (relative or absolute)
'  - r or R: Row Referemce (relative or absolute)
'  ^ At least 1 of the previous flags is required.
'  * Use \ to escape any c or r characters that come immediately after the reference,
'    that don't make up part of the formatting characters, ex. "=%C\r" == "=$A:$Ar"
' -Prefixes:
'  - s : Include worksheet name in reference
'  - b : Include workbook name in reference (impies s flag automatically)
'  - l : Use the last column/row reference used in the formula already
Public Function BuildFormulaFromArray(ByVal sFormat As String, Optional ByVal Params As Variant, Optional ByVal Src As String = "BuildFormulaFromArray") As String
    Dim soutFormParts() As Variant

    Dim sindexFormat As Integer ' Index into the format string
    Dim sindexOut As Integer    ' Index into the output string
    Dim indexParam As Integer ' Current index into the format string parameters
    Dim char As String  ' The last character parsed

    Dim fAbsCol As Boolean  ' Output an absolute column reference
    Dim fAbsRow As Boolean  ' Output an absolute row reference

    Dim fIncWB As Boolean    ' Include workbook name in output
    Dim fIncWS As Boolean    ' Include worksheet name in output
    Dim fIncCol As Boolean   ' Include the column reference in the output
    Dim fIncRow As Boolean   ' Include the row reference in the output

    Dim fUseLRow As Boolean  ' Do we use the last row number used earlier from the formatting parameters?
    Dim fUseLCol As Boolean  ' Do we use the last column object used earlier from the formatting parameters?
    Dim fColFirst As Boolean ' Do we look for an SS_Col object when checking the next parameter or a SmartSheet and/or integer?

    Dim fHasRefCh As Boolean ' Denotes if we're on the first reference character or not. (for detecting %%)
    Dim fInRef As Boolean    ' Denotes whether we've start reference parsing or are just passing strings
    Dim fInLast As Boolean   ' Denotes that the l flag was specified before our current flag/character
    Dim fRefDone As Boolean  ' Denotes that we've finished parsing a reference and need to append it's name to the join list and clear flags
    Dim fReparse As Boolean  ' Orders us to parse again in the case of 2 references 1 after another, ex. %Cr%Cr

    Dim lColumn As SS_Column ' Last SS_Column object used for L flag column references.
    Dim lSS As SmartSheet    ' Last SmartSheet object used for row only references, ex. "=%sR" == "Config!$1:$1"
    Dim lRow As Variant      ' Last row used for L flag row references.

    ReDim soutFormParts(0 To Len(sFormat))
    If IsArray(Params) Then _
        indexParam = LBound(Params)

    On Error GoTo BFFA_RefHandler

    For sindexFormat = 1 To Len(sFormat)
        char = Mid$(sFormat, sindexFormat, 1)

BFFA_Reparse:
        fReparse = False
        Select Case char
            Case "%":
                If fInRef Then
                    If fHasRefCh Then
                        fRefDone = True
                        fReparse = True
                    Else
                        fInRef = False
                        fHasRefCh = False
                        soutFormParts(sindexFormat) = char
                    End If
                Else
                    fInRef = True
                End If
            Case "l", "L":
                If fInRef Then
                    If fInLast Then _
                        Err.Raise ERROR_BF.BFE_BadChar, Src, "The L character is only allowed before an R/r or C/c flag." ' The L character was specified twice Ex. %ll
                    fInLast = True
                    fHasRefCh = True
                Else
                    soutFormParts(sindexFormat) = char
                End If
            Case "b", "B":
                If fInRef Then
                    If fIncWB Then _
                        Err.Raise ERROR_BF.BFE_BadChar, Src, "B flag specified more than once."
                    If fInLast Then _
                        Err.Raise ERROR_BF.BFE_BadChar, Src, "The L character is only allowed before an R/r or C/c flag." ' The L character has been specified before a valid char, ex. %lb
                    fIncWB = True
                    fHasRefCh = True
                Else
                    soutFormParts(sindexFormat) = char
                End If
            Case "s", "S":
                If fInRef Then
                    If fIncWS Then _
                        Err.Raise ERROR_BF.BFE_BadChar, Src, "S flag specified more than once."
                    If fInLast Then _
                        Err.Raise ERROR_BF.BFE_BadChar, Src, "The L character is only allowed before an R/r or C/c sequence." ' The L character has been specified before a valid char, ex. %ls
                    fIncWS = True
                    fHasRefCh = True
                Else
                    soutFormParts(sindexFormat) = char
                End If
            Case "\":
                If fInRef Then
                    fRefDone = True
                    If sindexFormat = Len(sFormat) Then _
                        soutFormParts(sindexFormat) = char
                Else
                    soutFormParts(sindexFormat) = char
                End If
            Case "r", "R":
                If Not fInRef Then
                    soutFormParts(sindexFormat) = char
                Else
                    If fIncRow Then _
                        Err.Raise ERROR_BF.BFE_BadChar, Src, "R flag specified more than once."
                    fHasRefCh = True
                    fAbsRow = (char = "R") ' Absolute row for uppercase
                    fIncRow = True
                    If fInLast Then
                        fUseLRow = True
                        fInLast = False
                    End If
                End If
            Case "c", "C":
                If Not fInRef Then
                    soutFormParts(sindexFormat) = char
                Else
                    If fIncCol Then _
                        Err.Raise ERROR_BF.BFE_BadChar, Src, "R flag specified more than once."
                    fHasRefCh = True
                    fAbsCol = (char = "C") ' Absolute column for uppercase
                    fIncCol = True
                    If Not fIncRow Then _
                        fColFirst = True
                    If fInLast Then
                        fUseLCol = True
                        fInLast = False
                    End If
                End If
            Case Else:
                If fInRef Then
                    If fInLast Then _
                        Err.Raise ERROR_BF.BFE_BadFlag, Src, "Invalid flag character after L flag specified: """ & char & """."
                    fRefDone = True
                    soutFormParts(sindexFormat) = char
                Else
                    soutFormParts(sindexFormat) = char
                End If
        End Select
        If fRefDone Then
            sindexOut = sindexFormat - 1 ' Output our reference a position back since we may have already parsed a valid character.
            If fInLast Then _
                Err.Raise ERROR_BF.BFE_MissFlag, Src, "Flag missing after L flag."
            Err.Raise ERROR_BF.BFE_HandleRef
        End If
    Next

    If fInRef Then
        sindexOut = sindexFormat - 1 ' Output our reference a position back since we may have already parsed a valid character.
        If fInLast Then _
            Err.Raise ERROR_BF.BFE_MissFlag, Src, "Flag missing after L flag."
        Err.Raise ERROR_BF.BFE_HandleRef
    End If

    BuildFormulaFromArray = Join$(soutFormParts, vbNullString)
    Exit Function

BFFA_RefHandler:
    If Err.Number <> ERROR_BF.BFE_HandleRef Then _
        Err.Raise Err.Number

    If fIncCol Then
        ' Output our column first if required.
        If Not fUseLCol Then
            If IsEmpty(Params) Or Not IsArray(Params) Then _
                Err.Raise ERROR_BF.BFE_BadCol, Src, "Missing SS_Column parameter for column info."
            If Not fColFirst And Not fUseLRow Then
                ' Pull the row information first, b/c of R being specified before C (colfirst).
                If indexParam > UBound(Params) Then Err.Raise ERROR_BF.BFE_BadRow, Src, "Missing integer parameter for row number."
                Select Case VarType(Params(indexParam))
                    Case vbByte, vbInteger, vbLong, vbLongLong:
                    Case Else:
                        Err.Raise ERROR_BF.BFE_BadRow, Src, "Invalid parameter type. Integer parameter required for row number. Value: " & CStr(Params(indexParam))
                End Select
                lRow = Params(indexParam)
                indexParam = indexParam + 1
            End If
            ' Pull the column info now.
            If indexParam > UBound(Params) Then Err.Raise ERROR_BF.BFE_BadCol, Src, "Missing SS_Column parameter for column info."
            If Not TypeOf Params(indexParam) Is SS_Column Then _
                Err.Raise ERROR_BF.BFE_BadCol, Src, "Invalid parameter type. SS_Column parameter required for column info."
            Set lColumn = Params(indexParam)
            indexParam = indexParam + 1
        ElseIf lColumn Is Nothing Then
            Err.Raise ERROR_BF.BFE_BadCol, Src, "Can't use last column, because no column was given prior."
        End If
        If fIncRow Then
            If fColFirst And Not fUseLRow Then
                ' Pull the row information now if we didn't earlier.
                If indexParam > UBound(Params) Then Err.Raise ERROR_BF.BFE_BadRow, Src, "Missing integer parameter for row number."
                Select Case VarType(Params(indexParam))
                    Case vbByte, vbInteger, vbLong, vbLongLong:
                    Case Else:
                        Err.Raise ERROR_BF.BFE_BadRow, Src, "Invalid parameter type. Integer parameter required for row number. Value: " & CStr(Params(indexParam))
                End Select
                lRow = Params(indexParam)
                indexParam = indexParam + 1
            End If
            If IsEmpty(lRow) Then Err.Raise ERROR_BF.BFE_BadRow, Src, "Can't use last row, because no row number was given prior."
            If lRow < 1 Then Err.Raise ERROR_BF.BFE_BadRow, Src, "Bad row number specified. Row: """ & lRow & """"
            soutFormParts(sindexOut) = soutFormParts(sindexOut) & lColumn.RowAddress((lRow), fAbsCol, fAbsRow, fIncWS, fIncWB)
        Else
            soutFormParts(sindexOut) = soutFormParts(sindexOut) & lColumn.ColumnAddress(fAbsCol, fIncWS, fIncWB)
        End If
    ElseIf fIncRow Then
        If Not fUseLRow Then
            ' Pull the SmartSheet information if we include either the workbook or worksheet name.
            If fIncWB Or fIncWS Then
                If indexParam > UBound(Params) Then Err.Raise ERROR_BF.BFE_BadSS, Src, "Missing SmartSheet parameter for worksheet/workbook info."
                If Not TypeOf Params(indexParam) Is SmartSheet Then _
                    Err.Raise ERROR_BF.BFE_BadSS, Src, "Invalid parameter type. SmartSheet parameter required for worksheet/workbook info."
                Set lSS = Params(indexParam)
                indexParam = indexParam + 1
            End If
            If indexParam > UBound(Params) Then Err.Raise ERROR_BF.BFE_BadRow, Src, "Missing integer parameter for row number."
            Select Case VarType(Params(indexParam))
                Case vbByte, vbInteger, vbLong, vbLongLong:
                Case Else:
                    Err.Raise ERROR_BF.BFE_BadRow, Src, "Invalid parameter type. Integer parameter required for row number. Value: " & CStr(Params(indexParam))
            End Select
            lRow = Params(indexParam)
            indexParam = indexParam + 1
        End If
        If IsEmpty(lRow) Then Err.Raise ERROR_BF.BFE_BadRow, Src, "Can't use last row, because no row number was given prior."
        If lRow < 1 Then Err.Raise ERROR_BF.BFE_BadRow, Src, "Bad row number specified. Row: """ & lRow & """"
        If fIncWS Or fIncWB Then
            soutFormParts(sindexOut) = soutFormParts(sindexOut) & lSS.RowName((lRow), fAbsRow, fIncWS, fIncWB)
        ElseIf fAbsRow Then
            soutFormParts(sindexOut) = soutFormParts(sindexOut) & "$" & lRow & ":$" & lRow
        Else
            soutFormParts(sindexOut) = soutFormParts(sindexOut) & lRow & ":" & lRow
        End If
    Else
        Err.Raise ERROR_BF.BFE_MissFlag, Src, "Missing both column and row flags."
    End If

    ' Reset our flags now that our reference is output.
    fHasRefCh = False
    fInRef = False
    fAbsCol = False
    fAbsRow = False
    fIncWB = False
    fIncWS = False
    fIncCol = False
    fIncRow = False
    fUseLRow = False
    fUseLCol = False
    fColFirst = False
    fRefDone = False
    If fReparse Then Resume BFFA_Reparse
    Resume Next
End Function

Public Function BuildFormula(ByVal sFormat As String, ParamArray Params()) As String
    BuildFormula = BuildFormulaFromArray(sFormat, CVar(Params), "BuildFormula")
End Function
