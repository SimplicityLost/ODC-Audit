Attribute VB_Name = "B__Vendor_Match_and_Cat_Exile"
Function categoryexiler()

Dim lookupws As Worksheet
Dim dataws As Worksheet
Dim lastrow As Long

Set lookupws = Sheets("Lookups")
Set dataws = Sheets("Compiled")

    'Read strings from string sheet into array
    Dim SearchArray(), SearchStrings
    ReDim SearchArray(1 To 5)
    For i = 1 To 5
        SearchStrings = Split(lookupws.Range("H" & i + 1) & ";" & lookupws.Range("I" & i + 1), ";")
        SearchArray(i) = SearchStrings
    Next i
    
    lastrow = dataws.Cells(dataws.Rows.Count, "A").End(xlUp).Row
    
    'For each category and type
    For Each cattype In SearchArray()
        For Each srchstr In Split(cattype(1), ",")
            lookupws.Range("A13") = cattype(0)
            lookupws.Range("A14") = srchstr
            
            dataws.Range("A:AD").AdvancedFilter _
                Action:=xlFilterInPlace, _
                CriteriaRange:=lookupws.Range("A13:A14"), _
                Unique:=False
            
            If dataws.Range("A1:A" & lastrow).SpecialCells(xlCellTypeVisible).Count > 1 Then
                For Each cell In dataws.Range("O2:O" & lastrow).SpecialCells(xlCellTypeVisible)
                    cell.Value = 86
                Next cell
            End If

            dataws.ShowAllData
        Next srchstr
    Next cattype
        dataws.Range("A:AD").Sort key1:=dataws.Range("O1"), Header:=xlYes
        
        dataws.Range("A:AD").AdvancedFilter _
                    Action:=xlFilterInPlace, _
                    CriteriaRange:=lookupws.Range("A16:A17"), _
                    Unique:=False
            
        If dataws.Range("A1:A" & lastrow).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Sheets("Exiles").Range("A2:A" & dataws.Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).Count + 1).EntireRow.Insert shift:=xlDown
            dataws.Range("A2:AD" & lastrow).SpecialCells(xlCellTypeVisible).Copy
            Sheets("Exiles").Paste Destination:=Sheets("Exiles").Range("A2")
            dataws.Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
    
        dataws.ShowAllData
    'For each string
    'Search for string
End Function

Function vendormatch()
    Dim lookupws As Worksheet
    Dim dataws As Worksheet
    Dim lastrow As Long
    Dim venws As Worksheet
    Dim findarray() As Integer
    Dim confarray() As Integer
    Dim highestscore()
    
    Set lookupws = Sheets("Lookups")
    Set venws = Sheets("Vendor List")
    Set dataws = Sheets("Compiled")
    lastrow = dataws.Cells(dataws.Rows.Count, "A").End(xlUp).Row
    ReDim findarray(11, 1)
    ReDim confarray(1, 1)
    ReDim highestscore(1)
'--------------------------------------------------------
    'For each row in the ODC file
'--------------------------------------------------------
For i = 2 To lastrow
Application.StatusBar = Format(i / lastrow, "Percent")
DoEvents
    'Look to see if a match is found
    'If so, check if that match is already in array
    'If not, check to see if there is room for it
    'If not, make another row in the array to hold the data
    'Move to next thingy
        
'--------------------------------------------------------
    'Look up ODC Payee in Vendor Name 1 and 2
'--------------------------------------------------------
If Not IsError(dataws.Range("R" & i).Value) Then
    If Len(dataws.Range("R" & i).Value) > 3 Then
        If Asc(dataws.Range("R" & i).Value) > 32 Then
        Set matchfoundcell = venws.Range("B:B").Find(dataws.Range("R" & i).Value)
        
        Do
            If Not matchfoundcell Is Nothing Then
                getout = 0
                MatchFound = matchfoundcell.Row
                For j = 0 To UBound(findarray, 2)
                    If findarray(0, j) = 0 Then Exit For
                    If findarray(0, j) = MatchFound Then getout = 1: Exit For
                Next j
                
                If getout = 0 Then
                    If findarray(0, UBound(findarray, 2)) > 0 Then
                        fasize = UBound(findarray, 2) + 1
                        ReDim Preserve findarray(11, fasize)
                    End If
                    
                    For j = 0 To UBound(findarray, 2)
                        If findarray(0, j) = 0 Then
                            findarray(0, j) = MatchFound
                            Exit For
                        End If
                    Next j
                End If
                
                Set matchfoundcell = venws.Range("B:B").FindNext(matchfoundcell)
            
            Else
                getout = 1
            End If
        Loop Until getout = 1
        
        Set matchfoundcell = venws.Range("C:C").Find(dataws.Range("R" & i).Value)

        Do
            If Not matchfoundcell Is Nothing Then
                getout = 0
                MatchFound = matchfoundcell.Row
                For j = 0 To UBound(findarray, 2)
                    If findarray(1, j) = 0 Then Exit For
                    If findarray(1, j) = MatchFound Then getout = 1: Exit For
                Next j
                
                If getout = 0 Then
                    If findarray(1, UBound(findarray, 2)) > 0 Then
                        fasize = UBound(findarray, 2) + 1
                        ReDim Preserve findarray(11, fasize)
                    End If
                    
                    For j = 0 To UBound(findarray, 2)
                        If findarray(1, j) = 0 Then
                            findarray(1, j) = MatchFound
                            Exit For
                        End If
                    Next j
                End If
                
                Set matchfoundcell = venws.Range("C:C").FindNext(matchfoundcell)
            
            Else
                getout = 1
            End If
        Loop Until getout = 1
        End If
    End If
End If
'--------------------------------------------------------
    'Look up ODC Payee 2 in Vendor 1 and 2
'--------------------------------------------------------
If Not IsError(dataws.Range("S" & i).Value) Then
    If Len(dataws.Range("S" & i).Value) > 3 Then
       If Asc(dataws.Range("S" & i).Value) > 32 Then
        Set matchfoundcell = venws.Range("B:B").Find(dataws.Range("S" & i).Value)

        Do
            If Not matchfoundcell Is Nothing Then
                getout = 0
                MatchFound = matchfoundcell.Row
                For j = 0 To UBound(findarray, 2)
                    If findarray(2, j) = 0 Then Exit For
                    If findarray(2, j) = MatchFound Then getout = 1: Exit For
                Next j
                
                If getout = 0 Then
                    If findarray(2, UBound(findarray, 2)) > 0 Then
                        fasize = UBound(findarray, 2) + 1
                        ReDim Preserve findarray(11, fasize)
                    End If
                    
                    For j = 0 To UBound(findarray, 2)
                        If findarray(2, j) = 0 Then
                            findarray(2, j) = MatchFound
                            Exit For
                        End If
                    Next j
                End If
                
                Set matchfoundcell = venws.Range("B:B").FindNext(matchfoundcell)
            
            Else
                getout = 1
            End If
        Loop Until getout = 1


        Set matchfoundcell = venws.Range("C:C").Find(dataws.Range("S" & i).Value)

        Do
            If Not matchfoundcell Is Nothing Then
                getout = 0
                MatchFound = matchfoundcell.Row
                For j = 0 To UBound(findarray, 2)
                    If findarray(3, j) = 0 Then Exit For
                    If findarray(3, j) = MatchFound Then getout = 1: Exit For
                Next j
                
                If getout = 0 Then
                    If findarray(3, UBound(findarray, 2)) > 0 Then
                        fasize = UBound(findarray, 2) + 1
                        ReDim Preserve findarray(11, fasize)
                    End If
                    
                    For j = 0 To UBound(findarray, 2)
                        If findarray(3, j) = 0 Then
                            findarray(3, j) = MatchFound
                            Exit For
                        End If
                    Next j
                End If
                
                Set matchfoundcell = venws.Range("C:C").FindNext(matchfoundcell)
            
            Else
                getout = 1
            End If
        Loop Until getout = 1
        End If
    End If
End If


'--------------------------------------------------------
        'Det_Desc in Name 1 and 2
'--------------------------------------------------------
If Not IsError(dataws.Range("P" & i).Value) Then
    If Len(dataws.Range("P" & i).Value) > 3 Then
       If Asc(dataws.Range("P" & i).Value) > 32 Then
        Set matchfoundcell = venws.Range("B:B").Find(dataws.Range("P" & i).Value)

        Do
            If Not matchfoundcell Is Nothing Then
                getout = 0
                MatchFound = matchfoundcell.Row
                For j = 0 To UBound(findarray, 2)
                    If findarray(6, j) = 0 Then Exit For
                    If findarray(6, j) = MatchFound Then getout = 1: Exit For
                Next j
                
                If getout = 0 Then
                    If findarray(6, UBound(findarray, 2)) > 0 Then
                        fasize = UBound(findarray, 2) + 1
                        ReDim Preserve findarray(11, fasize)
                    End If
                    
                    For j = 0 To UBound(findarray, 2)
                        If findarray(6, j) = 0 Then
                            findarray(6, j) = MatchFound
                            Exit For
                        End If
                    Next j
                End If
                
                Set matchfoundcell = venws.Range("B:B").FindNext(matchfoundcell)
            
            Else
                getout = 1
            End If
        Loop Until getout = 1
        
        
        Set matchfoundcell = venws.Range("C:C").Find(dataws.Range("P" & i).Value)

        Do
            If Not matchfoundcell Is Nothing Then
                getout = 0
                MatchFound = matchfoundcell.Row
                For j = 0 To UBound(findarray, 2)
                    If findarray(7, j) = 0 Then Exit For
                    If findarray(7, j) = MatchFound Then getout = 1: Exit For
                Next j
                
                If getout = 0 Then
                    If findarray(7, UBound(findarray, 2)) > 0 Then
                        fasize = UBound(findarray, 2) + 1
                        ReDim Preserve findarray(11, fasize)
                    End If
                    
                    For j = 0 To UBound(findarray, 2)
                        If findarray(7, j) = 0 Then
                            findarray(7, j) = MatchFound
                            Exit For
                        End If
                    Next j
                End If
                
                Set matchfoundcell = venws.Range("C:C").FindNext(matchfoundcell)
            
            Else
                getout = 1
            End If
        Loop Until getout = 1
        End If
    End If
End If
'--------------------------------------------------------
        'Payee in Vendor #
'--------------------------------------------------------
If Not IsError(dataws.Range("Q" & i).Value) Then
    If Len(dataws.Range("Q" & i).Value) > 3 Then
       If Asc(dataws.Range("q" & i).Value) > 32 Then
        Set matchfoundcell = venws.Range("A:A").Find(dataws.Range("Q" & i).Value, lookat:=xlWhole)

        Do
            If Not matchfoundcell Is Nothing Then
                getout = 0
                MatchFound = matchfoundcell.Row
                For j = 0 To UBound(findarray, 2)
                    If findarray(8, j) = 0 Then Exit For
                    If findarray(8, j) = MatchFound Then getout = 1: Exit For
                Next j
                
                If getout = 0 Then
                    If findarray(8, UBound(findarray, 2)) > 0 Then
                        fasize = UBound(findarray, 2) + 1
                        ReDim Preserve findarray(11, fasize)
                    End If
                    
                    For j = 0 To UBound(findarray, 2)
                        If findarray(8, j) = 0 Then
                            findarray(8, j) = MatchFound
                            Exit For
                        End If
                    Next j
                End If
                
                Set matchfoundcell = venws.Range("A:A").FindNext(matchfoundcell)
            
            Else
                getout = 1
            End If
        Loop Until getout = 1
        End If
        
    End If
End If
'--------------------------------------------------------
        'Look up ODC Address in Address
'--------------------------------------------------------
If Not IsError(dataws.Range("U" & i).Value) Then
    If Len(dataws.Range("U" & i).Value) > 3 Then
        If Asc(dataws.Range("U" & i).Value) > 32 Then
        Set matchfoundcell = venws.Range("D:D").Find(dataws.Range("U" & i).Value, lookat:=xlPart)

        Do
            If Not matchfoundcell Is Nothing Then
                getout = 0
                MatchFound = matchfoundcell.Row
                For j = 0 To UBound(findarray, 2)
                    If findarray(9, j) = 0 Then Exit For
                    If findarray(9, j) = MatchFound Then getout = 1: Exit For
                Next j
                
                If getout = 0 Then
                    If findarray(9, UBound(findarray, 2)) > 0 Then
                        fasize = UBound(findarray, 2) + 1
                        ReDim Preserve findarray(11, fasize)
                    End If
                    
                    For j = 0 To UBound(findarray, 2)
                        If findarray(9, j) = 0 Then
                            findarray(9, j) = MatchFound
                            Exit For
                        End If
                    Next j
                End If
                
                Set matchfoundcell = venws.Range("D:D").FindNext(matchfoundcell)
            
            Else
                getout = 1
            End If
        Loop Until getout = 1
        End If
    End If
End If
'--------------------------------------------------------
        'ODC City in City
'--------------------------------------------------------
If Not IsError(dataws.Range("V" & i).Value) Then
    If Len(dataws.Range("V" & i).Value) > 3 Then
               If Asc(dataws.Range("V" & i).Value) > 32 Then
        Set matchfoundcell = venws.Range("E:E").Find(dataws.Range("V" & i).Value)

        Do
            If Not matchfoundcell Is Nothing Then
                getout = 0
                MatchFound = matchfoundcell.Row
                For j = 0 To UBound(findarray, 2)
                    If findarray(10, j) = 0 Then Exit For
                    If findarray(10, j) = MatchFound Then getout = 1: Exit For
                Next j

                If getout = 0 Then
                    If findarray(10, UBound(findarray, 2)) > 0 Then
                        fasize = UBound(findarray, 2) + 1
                        ReDim Preserve findarray(11, fasize)
                    End If

                    For j = 0 To UBound(findarray, 2)
                        If findarray(10, j) = 0 Then
                            findarray(10, j) = MatchFound
                            Exit For
                        End If
                    Next j
                End If

                Set matchfoundcell = venws.Range("E:E").FindNext(matchfoundcell)

            Else
                getout = 1
            End If
        Loop Until getout = 1
        End If
    End If
End If
'--------------------------------------------------------
        'ODC Zip in ZIP
'--------------------------------------------------------
If Not IsError(dataws.Range("X" & i).Value) Then
    If Len(dataws.Range("X" & i).Value) > 3 Then
        If Asc(dataws.Range("X" & i).Value) > 32 Then
        Set matchfoundcell = venws.Range("G:G").Find(dataws.Range("X" & i).Value)

        Do
            If Not matchfoundcell Is Nothing Then
                getout = 0
                MatchFound = matchfoundcell.Row
                For j = 0 To UBound(findarray, 2)
                    If findarray(11, j) = 0 Then Exit For
                    If findarray(11, j) = MatchFound Then getout = 1: Exit For
                Next j

                If getout = 0 Then
                    If findarray(11, UBound(findarray, 2)) > 0 Then
                        fasize = UBound(findarray, 2) + 1
                        ReDim Preserve findarray(11, fasize)
                    End If

                    For j = 0 To UBound(findarray, 2)
                        If findarray(11, j) = 0 Then
                            findarray(11, j) = MatchFound
                            Exit For
                        End If
                    Next j
                End If

                Set matchfoundcell = venws.Range("G:G").FindNext(matchfoundcell)

            Else
                getout = 1
            End If
        Loop Until getout = 1
        End If
    End If
End If

'--------------------------------------------------------
    'Walk through entire array and compare found
    'row #s for commonalities
'--------------------------------------------------------

'Generate confidence score
    'Start walking through array
    u = 0
    getout = 0
    For s = 0 To UBound(findarray, 1)
        For t = 0 To UBound(findarray, 2)
            'Move to next slot if this one is 0
            If findarray(s, t) = 0 Then: Exit For
            
            'Check for duplicates
            For Z = 0 To UBound(confarray, 2)
                If confarray(0, Z) = findarray(s, t) Then getout = 1: Exit For
            Next Z
            
            'Make room in the array if it's full
            If confarray(0, UBound(confarray)) > 0 And getout = 0 Then
                casize = UBound(confarray, 2) + 1
                ReDim Preserve confarray(1, casize)
            End If
            

            
            If getout = 0 Then
                'Assign row number to array
                confarray(0, u) = findarray(s, t)
    
                'increment array index
                u = u + 1
            End If
            getout = 0
        Next t
    Next s
    
    For u = 0 To UBound(confarray, 2)
        For s = 0 To UBound(findarray, 1)
            For t = 0 To UBound(findarray, 2)
            If findarray(s, t) = confarray(0, u) And findarray(s, t) <> 0 Then
                confarray(1, u) = confarray(1, u) + confscore(s)
            End If
            
            Next t
        Next s
        
        If confarray(1, u) > highestscore(1) Then
            highestscore(0) = venws.Range("A" & confarray(0, u)).Value
            highestscore(1) = confarray(1, u)
        End If
    Next u
    
'    For u = 0 To UBound(confarray, 2)
'        If confarray(1, u) > highestscore(1) Then
'            highestscore(0) = confarray(0, u)
'            highestscore(1) = confarray(1, u)
'        End If
'    Next u
    
    If highestscore(1) > 5 Then
        'Do the thing
        Sheets("Compiled").Range("O" & i).Value = highestscore(0)
    End If
    
    'Assign priorities to each matchup
    'Create array of confidence scores
    'Figure out highest row
    
    'Do the next row/entry
    ReDim findarray(11, 1)
    ReDim highestscore(1)
    ReDim confarray(1, 1)
    Next i
End Function
Function confscore(colin)
Select Case colin
Case 0 To 7
    confscore = 8
Case 8
    confscore = 6
Case 9
    confscore = 4
Case 10, 11
    confscore = 1
End Select
End Function


