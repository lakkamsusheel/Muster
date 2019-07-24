'-------------------------------------------------------------------------------
' MUSTER.Info.LustRemediationCollection
'   Provides a stongly-typed collection for storing LustRemediation objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC     03/02/05      Original class definition.
'
' Function          Description
' Item(ID)          Gets/Sets the EntityInfo requested by the string arg ID
' Values()          Returns the collection of Entities in the EntityCollection
' GetKeys()         Returns an array of string containing the keys in the EntityCollection
' Add(Entity)       Adds the Entity supplied to the Entity Collection
' Contains(Entity)  Returns True/False to indicate if the supplied Entity is contained
'                   in the Entity Collection
' Contains(Key)     Returns True/False to indicate if the supplied Entity key is contained
'                   in the Entity Collection
' Remove(Entity)    Removes the Entity supplied from the Entity Collection
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    Public Class LustRemediationCollection
        Inherits DictionaryBase
        Public Delegate Sub LustRemediationColChangedEventHandler()
        ' Raised whenever a member is inserted into or removed from the LustEvent collection
        Public Event ProtoColChanged As LustRemediationColChangedEventHandler

        Public Sub Add(ByVal value As MUSTER.Info.LustRemediationInfo)
            Me.Item(value.ID.ToString) = value
        End Sub
        Public Sub ChangeKey(ByVal OldKey As String, ByVal NewKey As String)
            If Me.Contains(OldKey) Then
                Dim MyInfo As New Object
                MyInfo = MyBase.Dictionary.Item(OldKey)
                Me.Remove(OldKey)
                Me.Add(MyInfo)
            End If
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.LustRemediationInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public ReadOnly Property GetKeys() As String()
            Get
                Dim KeyCol(MyBase.Dictionary.Keys.Count - 1) As String
                MyBase.Dictionary.Keys.CopyTo(KeyCol, 0)
                Array.Sort(KeyCol)
                Return KeyCol
            End Get
        End Property
        Default Public Property Item(ByVal index As String) As MUSTER.Info.LustRemediationInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.LustRemediationInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.LustRemediationInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                Else
                    MyBase.Dictionary.Item(index) = Value
                End If
            End Set
        End Property

        Public Sub Remove(ByVal value As MUSTER.Info.LustRemediationInfo)
            MyBase.Dictionary.Remove(value.ID.ToString)
        End Sub
        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
        End Sub
        Public ReadOnly Property Values() As ICollection
            Get
                Return MyBase.Dictionary.Values
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
