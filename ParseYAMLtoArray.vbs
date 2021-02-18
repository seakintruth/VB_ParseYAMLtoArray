Public Function ParseYAMLtoArray(ByVal filePath) ' as array
' Version 1.0.4
' Dependencies: NONE
' Modified from this post: https://stackoverflow.com/a/40659701/1146659
' License: - CC BY-SA 4.0 - <https://creativecommons.org/licenses/by-sa/4.0/>
' Contributors: cgasp <https://stackoverflow.com/users/1862421/cgasp>; Jeremy D. Gerdes <jeremy.gerdes@navy.mil>;
' Reference: https://yaml.org/refcard.html
' Usage Example: debug.print ParseYAMLtoArray(GetCurrentFileFolder() & "\" & "documentation" & "\" & "exampleNested.yaml")(1,3)
' Notes: Using late binding to run for all vb engines
' -------------------------------
' Known ParserIssues:
'   - Niave: This is not spec conforming, just usefull enough, use another parser if you need more features.
'     See spec at: http://yaml.org/spec/1.2/spec.html
'   - A block scalar indicator should include all subsequent rows that have the same white space intentation past the current line
'     this parser fails to do this if any of those following row contains a ":"
'   - YAML denotes nesting via indent delimitation (white space), this parser attempts to record nested "{level=n}" in the data
'     column for each empty Category, and ignores all other nesting.
'   -This parser ignores all cast data types like "!!float " whatever is accepting the results of this Public Function will
'     have to handle any type casting in the YAML document.

Const ForReading = 1
Dim arryReturn() ' As variant
Dim text ' As String
Dim textline ' As String
Dim objFSO 'As Scripting.FileSystemObject
Dim objFile 'As Scripting.TextStream
Dim intLastLineWhiteSpace 'As Integer
Dim dataArray 'As Variant
Dim sizeArray 'As Long
Dim oneline 'As String
Dim Data 'As Variant
Dim Key 'As Variant
Dim intRow 'as integer
Dim intColumn 'as integer
Dim intNestingLevel 'As Integer
Dim intLastNestingSpaces 'As Integer
Dim intCurrentNestingSpaces 'As Integer
Dim intThisLineWhiteSpace
Dim fIsNestedHeader
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    ' verify if file exists
    If objFSO.FileExists(filePath) Then
        Set objFile = objFSO.GetFile(filePath).OpenAsTextStream(ForReading)
        'Open FilePath For Input As #1
        intRow = 0
        intNestingLevel = 1
        Do Until objFile.AtEndOfStream
            intThisLineWhiteSpace = Len(textline) - Len(LTrim(textline))
            textline = objFile.ReadLine
            oneline = Trim(textline) 'remove leading/trailing spaces
            ' test if line doesn't start with --- or #
            If Left(oneline, 3) <> "---" And Left(oneline, 1) <> "#" Then
                dataArray = Split(oneline, ":", 2)
                sizeArray = UBound(dataArray, 1) - LBound(dataArray, 1) + 1
                ' Verification Empty Lines and Split don't occur
                If Not Len(oneline) = 0 And Not sizeArray = 0 Then
                    fIsNestedHeader = False
                    If sizeArray = 1 And intThisLineWhiteSpace > intLastLineWhiteSpace Then  ' HEADER
                        fIsNestedHeader = True
                    ElseIf sizeArray = 2 Then  ' HEADER: <NULL>
                        fIsNestedHeader = Len(Trim(dataArray(0))) <> 0 And Len(Trim(dataArray(1))) = 0
                    End If
                    If sizeArray = 1 And intThisLineWhiteSpace >= intLastLineWhiteSpace And Len(Trim(dataArray(0))) > 0 Then ' semicolins in a block breaks this parser
                        'assume we are continuing the data from previous line
                        intRow = intRow - 1 ' use previous row in the array
                        Data = Trim(dataArray(0))
                        'remove leading block annotation | or >
                        If arryReturn(1, intRow) = "|" Or arryReturn(1, intRow) = ">" Then
                            If Len(arryReturn(1, intRow)) = 1 Then
                                arryReturn(1, intRow) = vbNullString
                            Else
                                arryReturn(1, intRow) = Right(arryReturn(1, intRow), Len(arryReturn(1, intRow) - 1))
                            End If
                        End If
                        arryReturn(1, intRow) = arryReturn(1, intRow) & vbCrLf & Data
                    ElseIf fIsNestedHeader Then
                        'Category/Header
                        Key = Trim(dataArray(0))
                        ReDim Preserve arryReturn(1, intRow)
                        arryReturn(0, intRow) = Key
                        ' calculate nesting level - just kind of works,
                        ' doesn't really map to what's in the YAML as nesting back up is actually dependent on the number of spaces not previous nesting...
                        intCurrentNestingSpaces = intThisLineWhiteSpace
                        If intThisLineWhiteSpace = 0 Then
                            'We are back at level 1
                            intNestingLevel = 1
                        Else
                            If intCurrentNestingSpaces > intLastNestingSpaces Then
                                intNestingLevel = intNestingLevel + 1
                            ElseIf intCurrentNestingSpaces < intLastNestingSpaces Then
                                intNestingLevel = intNestingLevel - 1
                            'Else 'should be equal so intNestingLevel, stays the same
                                'intCurrentNestingSpaces = intLastNestingSpaces
                            End If
                        End If
                        arryReturn(1, intRow) = "{level=" & intNestingLevel & "}"
                        intLastNestingSpaces = intThisLineWhiteSpace
                    Else
                        Data = Trim(dataArray(1))
                        Key = Trim(dataArray(0))
                        ReDim Preserve arryReturn(1, intRow)
                        arryReturn(0, intRow) = Key
                        arryReturn(1, intRow) = Data
                    End If
                    intRow = intRow + 1
                End If
            End If
            intLastLineWhiteSpace = Len(textline) - Len(LTrim(textline))
        Loop
        objFile.Close
        Dim arryReturnTemp
        'Must build array in Array(column,row) format to be able to append rows in VBScript, now transform to the standard Array(row,column) format
        If TransposeArray(arryReturn, arryReturnTemp) Then
            ParseYAMLtoArray = arryReturnTemp
        Else
            Err.Raise vbObjectError + 667, "ParseYAML", "Failed to Transform array"
        End If
    Else
        Err.Raise vbObjectError + 666, "ParseYAML", "Config file not found"
    End If
End Function

Public Function TransposeArray(ByRef InputArr, ByRef OutputArr) 'As Variant, ByRef OutputArr As Variant) As Boolean
    ' Version 1.0.0
    ' Dependencies: NONE
    ' Note: The following Public Function has been modified by jeremy.gerdes@navy.mil to conform to VBScipt from:
    '   http://www.cpearson.com/excel/vbaarrays.htm
    ' License: Charles H. Pearson. All of the formulas and VBA code are explicitly granted to the Public Domain. You may use the formulas and VBA code on this site for any purpose you see fit without permission from me. This includes inclusion in commercial works and works for hire. By using the formula and code on this site, you agree to hold Charles H. Pearson and Pearson Software Consulting, LLC, free of any liability. The formulas and code are presented as is and the author makes no warranty, express or implied, of their fitness for use. You assume all responsibility for testing and ensuring that the code works properly in your environment

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' TransposeArray
    ' This transposes a two-dimensional array. It returns True if successful or
    ' False if an error occurs. InputArr must be two-dimensions. OutputArr must be
    ' a dynamic array. It will be Erased and resized, so any existing content will
    ' be destroyed.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim RowNdx ' As Long
    Dim ColNdx ' As Long
    Dim LB1  ' As Long
    Dim LB2 ' As Long
    Dim UB1 ' As Long
    Dim UB2 ' As Long

    '''''''''''''''''''''''''''''''''''
    ' Ensure InputArr is an array
    '''''''''''''''''''''''''''''''''''
    If (IsArray(InputArr) = False) Then
        TransposeArray = False
        Exit Function
    End If

    '''''''''''''''''''''''''''''''''''''''
    ' Get the Lower and Upper bounds of
    ' InputArr.
    '''''''''''''''''''''''''''''''''''''''
    LB1 = LBound(InputArr, 1)
    LB2 = LBound(InputArr, 2)
    UB1 = UBound(InputArr, 1)
    UB2 = UBound(InputArr, 2)

    '''''''''''''''''''''''''''''''''''''''''
    ' Erase and ReDim OutputArr
    '''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    'If it's an array empty it, if not then it's empty
    Erase OutputArr
    On Error GoTo 0
    'In VBS we can't ReDim Array(LowBound To HighBound) all arrays must conform to Lbound = 0
    If LB1 <> 0 Or LB2 <> 0 Then
        TransposeArray = False
        Exit Function
    End If


    ReDim OutputArr(UB2, UB1)

    For RowNdx = LBound(InputArr, 2) To UBound(InputArr, 2)
        For ColNdx = LBound(InputArr, 1) To UBound(InputArr, 1)
            OutputArr(RowNdx, ColNdx) = InputArr(ColNdx, RowNdx)
        Next ' ColNdx
    Next ' RowNdx

    TransposeArray = True

End Function
