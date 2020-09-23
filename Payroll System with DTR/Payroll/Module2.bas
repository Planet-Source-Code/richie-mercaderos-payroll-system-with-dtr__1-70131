Attribute VB_Name = "Module2"
Public Enum FindOptions
    PartOfWord = 0
    MatchCase = 1
    WholeWordOnly = 3
End Enum

Public Function FindLVItem(ByRef vLV As ListView, sCriteria As String, Optional iOption As FindOptions = 0, Optional MultiSelect As Boolean = False, Optional InverseSelection As Boolean = False, Optional FindNext As Boolean = False)

    Dim i As Integer
    Dim isFound As Boolean
    Dim li As Integer
    Dim StartPos As Integer
    
'On Error GoTo eh
    
    If vLV.ListItems.Count < 1 Then Exit Function

    If FindNext = True And vLV.SelectedItem.Index < vLV.ListItems.Count Then
        For li = 1 To vLV.SelectedItem.Index
            vLV.ListItems(li).Selected = False
        Next
        StartPos = vLV.SelectedItem.Index + 1
    Else
        For li = 1 To vLV.ListItems.Count
            vLV.ListItems(li).Selected = False
        Next
        StartPos = 1
    End If
    
    'set flag to default
    isFound = False
    
    For li = StartPos To vLV.ListItems.Count
        
        Select Case iOption
            
            Case FindOptions.PartOfWord  'normal

                If InStr(1, LCase(vLV.ListItems(li).Text), LCase(sCriteria)) > 0 Then
                                        
                    isFound = True

                Else

                    'check subitems
                    For i = 1 To vLV.ListItems(li).ListSubItems.Count
                        If InStr(1, LCase(vLV.ListItems(li).ListSubItems(i)), LCase(sCriteria)) > 0 Then
                            
                            isFound = True
                            Exit For
                        
                        End If
                    Next
                                        
                End If
                
            Case FindOptions.MatchCase  'match case
            
            Case FindOptions.WholeWordOnly  ' whole word only
                
            
        End Select
        
        
        
        
        If isFound Then
            
            vLV.ListItems(li).Selected = CBool(True - InverseSelection)
            vLV.ListItems(li).EnsureVisible
            
            If Not MultiSelect Then Exit For
        
        Else
            vLV.ListItems(li).Selected = CBool(False - InverseSelection)
        End If
        
    Next
    
    If FindNext = True And isFound = False And StartPos > 1 Then
        
        For li = 1 To StartPos
            
            Select Case iOption
                
                Case FindOptions.PartOfWord  'normal
    
                    If InStr(1, LCase(vLV.ListItems(li).Text), LCase(sCriteria)) > 0 Then
                                            
                        isFound = True
    
                    Else
    
                        'check subitems
                        For i = 1 To vLV.ListItems(li).ListSubItems.Count
                            If InStr(1, LCase(vLV.ListItems(li).ListSubItems(i)), LCase(sCriteria)) > 0 Then
                                
                                isFound = True
                                Exit For
                            
                            End If
                        Next
                                            
                    End If
                    
                Case FindOptions.MatchCase  'match case
                
                Case FindOptions.WholeWordOnly  ' whole word only
                    
                
            End Select
            
            
            
            
            If isFound Then
                
                vLV.ListItems(li).Selected = CBool(True - InverseSelection)
                vLV.ListItems(li).EnsureVisible
                
                If Not MultiSelect Then Exit For
            
            Else
                vLV.ListItems(li).Selected = CBool(False - InverseSelection)
            End If
            
        Next
    End If
'On Error Resume Next
Exit Function
eh:
    MsgBox Err.Description
    Resume Next
End Function


Public Function cSentenceCase(sText As String) As String
    
    Dim splitText() As String
    Dim newWord As String
    Dim i As Integer
    
    'check if null---------------
    If Len(sText) < 1 Then
        cSentenceCase = ""
        Exit Function
    End If
    'end Null --------------------
    
    'convert
    sText = Trim(sText)
    
    splitText = Split(sText, " ")
    
    For i = 0 To UBound(splitText)
        If Len(Trim(splitText(i))) > 0 Then
            newWord = UCase(Left(Trim(splitText(i)), 1)) & LCase(Right(Trim(splitText(i)), Len(Trim(splitText(i))) - 1))
            cSentenceCase = cSentenceCase & " " & newWord
        End If
    Next
    
    cSentenceCase = Trim(cSentenceCase)
End Function

