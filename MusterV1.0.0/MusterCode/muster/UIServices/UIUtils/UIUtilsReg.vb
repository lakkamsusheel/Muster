Public Class UIUtilsReg
    'This function is to filter an array list based on a specific value/parent
    'This is currently used in filtering tank secondary options based on the tank material of construction
    'It accepts the array list to be filtered 'MasterChildArray' and the value on which it has to be filtered
    ''nSelectedParent' and returns a filtered arraylist
    'Friend Shared Function FilterChildren(ByVal MasterChildArray As ArrayList, ByVal nSelectedParent As Integer, Optional ByVal reg As Registration = Nothing) As ArrayList
    '    Dim FilteredChildren As New ArrayList
    '    Dim LstItem As InfoRepository.LookupProperty
    '    Dim ChildEnumerator As System.Collections.IEnumerator = MasterChildArray.GetEnumerator()

    '    FilterChildren = Nothing
    '    Try
    '        While ChildEnumerator.MoveNext()
    '            LstItem = ChildEnumerator.Current
    '            If LstItem.ParentId = nSelectedParent Then
    '                If LstItem.Id <> 341 Then
    '                    FilteredChildren.Add(LstItem)
    '                End If

    '            End If
    '        End While
    '        FilterChildren = FilteredChildren
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        ChildEnumerator = Nothing
    '    End Try
    'End Function
    'This function is to filter an array list based on a specific value/parent
    'Friend Shared Function FilterTnkReleaseDetectionList(ByVal MasterChildArray As ArrayList, ByVal nSelectedParent As Integer, Optional ByVal reg As Registration = Nothing) As ArrayList
    '    Dim FilteredChildren As New ArrayList
    '    Dim LstItem As InfoRepository.LookupProperty
    '    Dim strReleaseDetection As String
    '    Dim ChildEnumerator As System.Collections.IEnumerator = MasterChildArray.GetEnumerator()

    '    FilterTnkReleaseDetectionList = Nothing
    '    Try
    '        While ChildEnumerator.MoveNext()
    '            LstItem = ChildEnumerator.Current
    '            If LstItem.ParentId = nSelectedParent Then
    '                If Not reg Is Nothing Then

    '                    Dim nTnkCapacity As Integer
    '                    Dim nTnkCompCapacity As Integer
    '                    nTnkCapacity = Integer.Parse(IIf(reg.txtNonCompTankCapacity.Text <> "", reg.txtNonCompTankCapacity.Text, 0))
    '                    nTnkCompCapacity = Integer.Parse(IIf(reg.txtTankCapacity.Text <> "", reg.txtTankCapacity.Text, 0))
    '                    If reg.chkEmergencyPower.Checked Then
    '                        If Not LstItem.Type = strReleaseDetection Then
    '                            If reg.chkTankCompartment.Checked = False And nTnkCapacity < 2000 Then
    '                                FilteredChildren.Add(LstItem)
    '                                strReleaseDetection = LstItem.Type

    '                            Else
    '                                If reg.chkTankCompartment.Checked = True And nTnkCompCapacity < 2000 Then
    '                                    FilteredChildren.Add(LstItem)
    '                                    strReleaseDetection = LstItem.Type

    '                                Else
    '                                    If LstItem.Id <> 337 Then
    '                                        FilteredChildren.Add(LstItem)
    '                                        strReleaseDetection = LstItem.Type
    '                                    End If
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                    If reg.chkTankCompartment.Checked = False Then

    '                        If nTnkCapacity >= 2000 And LstItem.Id <> 337 And reg.chkEmergencyPower.Checked = False Then
    '                            If LstItem.Id <> 341 Then
    '                                FilteredChildren.Add(LstItem)
    '                            End If
    '                        End If
    '                    Else
    '                        If nTnkCompCapacity >= 2000 And LstItem.Id <> 337 And reg.chkEmergencyPower.Checked = False Then
    '                            If LstItem.Id <> 341 Then
    '                                FilteredChildren.Add(LstItem)
    '                            End If
    '                        End If
    '                    End If
    '                Else
    '                    If LstItem.Id <> 341 Then
    '                        FilteredChildren.Add(LstItem)
    '                    End If
    '                End If
    '            End If
    '        End While
    '        FilterTnkReleaseDetectionList = FilteredChildren
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        ChildEnumerator = Nothing
    '    End Try
    'End Function
End Class