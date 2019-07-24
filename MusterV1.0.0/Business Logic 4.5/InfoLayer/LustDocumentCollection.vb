'-------------------------------------------------------------------------------
' MUSTER.Info.LustDocumentCollection
'   Provides a stongly-typed collection for storing LustDocument objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC     03/10/05      Original class definition.
'
' Function          Description
' Item(ID)          Gets/Sets the LustDocumentInfo requested by the string arg ID
' Values()          Returns the collection of LustDocumentInfo in the LustDocumentCollection
' GetKeys()         Returns an array of string containing the keys in the LustDocumentCollection
' Add(LustDocumentInfo)       Adds the LustDocumentInfo supplied to the Entity Collection
' Contains(LustDocumentInfo)  Returns True/False to indicate if the supplied LustDocumentInfo is contained
'                   in the LustDocumentCollection
' Contains(Key)     Returns True/False to indicate if the supplied LustDocumentInfo key is contained
'                   in the LustDocumentCollection
' Remove(LustDocumentInfo)    Removes the LustDocumentInfo supplied from the LustDocumentCollection
'
'

Namespace MUSTER.Info
    Public Class LustDocumentCollection
        Inherits DictionaryBase
        Public Delegate Sub LustDocumentColChangedEventHandler()
        ' Raised whenever a member is inserted into or removed from the LustEvent collection
        Public Event LustDocumentColChanged As LustDocumentColChangedEventHandler
        Public Sub Add(ByVal value As MUSTER.Info.LustDocumentInfo)
            ' #Region "XDEOperation" ' Begin Template Expansion{1E386D57-39D2-4728-AD5B-51B94C2E0671}
            Me.Item(value.ID.ToString) = value
            ' #End Region ' XDEOperation End Template Expansion{1E386D57-39D2-4728-AD5B-51B94C2E0671}
        End Sub
        Public Sub ChangeKey(ByVal OldKey As String, ByVal NewKey As String)
            ' #Region "XDEOperation" ' Begin Template Expansion{AB52984F-E58B-40AF-9A9D-53B137F514E5}
            If Me.Contains(OldKey) Then
                Dim MyInfo As New Object
                MyInfo = MyBase.Dictionary.Item(OldKey)
                Me.Remove(OldKey)
                Me.Add(MyInfo)
            End If
            ' #End Region ' XDEOperation End Template Expansion{AB52984F-E58B-40AF-9A9D-53B137F514E5}
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.LustDocumentInfo) As Boolean
            ' #Region "XDEOperation" ' Begin Template Expansion{C8D49447-D902-495F-AC9D-03F0AE805E0C}
            Return MyBase.Dictionary.Contains(value.ID.ToString)
            ' #End Region ' XDEOperation End Template Expansion{C8D49447-D902-495F-AC9D-03F0AE805E0C}
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            ' #Region "XDEOperation" ' Begin Template Expansion{CA920BE8-F55C-4DA4-B3DF-3CC9E8D8408A}
            Return MyBase.Dictionary.Contains(Key)
            ' #End Region ' XDEOperation End Template Expansion{CA920BE8-F55C-4DA4-B3DF-3CC9E8D8408A}
        End Function
        Public ReadOnly Property GetKeys() As String()
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{BC9C5530-0675-476C-9B6D-975DDD7268AC}
                Dim KeyCol(MyBase.Dictionary.Keys.Count - 1) As String
                MyBase.Dictionary.Keys.CopyTo(KeyCol, 0)
                Array.Sort(KeyCol)
                Return KeyCol
                ' #End Region ' XDEOperation End Template Expansion{BC9C5530-0675-476C-9B6D-975DDD7268AC}
                ' #Region "XDEOperation" ' Begin Template Expansion{0AE692A7-B0A4-46C4-9999-E1F6BB81213E}
                'Dim KeyCol(MyBase.Dictionary.Keys.Count - 1) As String
                'MyBase.Dictionary.Keys.CopyTo(KeyCol, 0)
                'Array.Sort(KeyCol)
                'Return KeyCol
                ' #End Region ' XDEOperation End Template Expansion{0AE692A7-B0A4-46C4-9999-E1F6BB81213E}
            End Get
        End Property
        Default Public Property Item(ByVal index As String) As MUSTER.Info.LustDocumentInfo
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{71588C6A-EF74-4200-98A5-6FBD1587D3FA}
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.LustDocumentInfo)
                ' #End Region ' XDEOperation End Template Expansion{71588C6A-EF74-4200-98A5-6FBD1587D3FA}
            End Get
            Set(ByVal Value As MUSTER.Info.LustDocumentInfo)
                ' #Region "XDEOperation" ' Begin Template Expansion{C6F883B7-69F6-470E-A497-4178E6DB8496}
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent LustDocumentColChanged()
                End If
                ' #End Region ' XDEOperation End Template Expansion{C6F883B7-69F6-470E-A497-4178E6DB8496}
            End Set
        End Property
        Public Sub Remove(ByVal value As MUSTER.Info.LustDocumentInfo)
            ' #Region "XDEOperation" ' Begin Template Expansion{1511F5E9-E717-43A0-8D2D-263967215B13}
            MyBase.Dictionary.Remove(value.ID.ToString)
            RaiseEvent LustDocumentColChanged()
            ' #End Region ' XDEOperation End Template Expansion{1511F5E9-E717-43A0-8D2D-263967215B13}
        End Sub
        Public Sub Remove(ByVal value As String)
            ' #Region "XDEOperation" ' Begin Template Expansion{A08BE081-839F-4606-9971-55F005749562}
            MyBase.Dictionary.Remove(value)
            RaiseEvent LustDocumentColChanged()
            ' #End Region ' XDEOperation End Template Expansion{A08BE081-839F-4606-9971-55F005749562}
        End Sub
        Public ReadOnly Property Values() As ICollection
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{3F64C9B6-73DE-4AAE-9C7F-BF5153D04D5F}
                Return MyBase.Dictionary.Values
                ' #End Region ' XDEOperation End Template Expansion{3F64C9B6-73DE-4AAE-9C7F-BF5153D04D5F}
            End Get
        End Property
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
        End Sub
    End Class
End Namespace
