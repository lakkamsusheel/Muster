'-------------------------------------------------------------------------------
' MUSTER.Info.ContactStructCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      KKM         03/30/2005  Class definition
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
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    Public Class ContactStructCollection
        Inherits DictionaryBase
#Region "Public Events"
        Public Event ContactStructColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.ContactStructInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.ContactStructInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.ContactStructInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.entityAssocID.ToString, Value)
                    RaiseEvent ContactStructColChanged(True)
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent ContactStructColChanged(True)
                End If
            End Set
        End Property
        Public ReadOnly Property Values() As ICollection
            Get
                Return MyBase.Dictionary.Values
            End Get
        End Property
        Public ReadOnly Property GetKeys() As String()
            Get
                Dim KeyCol(MyBase.Dictionary.Keys.Count - 1) As String
                MyBase.Dictionary.Keys.CopyTo(KeyCol, 0)
                Array.Sort(KeyCol)
                Return KeyCol
            End Get
        End Property
        Public Sub Add(ByVal value As MUSTER.Info.ContactStructInfo)
            Me.Item(value.entityAssocID.ToString) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.ContactStructInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.entityAssocID.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.ContactStructInfo)
            MyBase.Dictionary.Remove(value.entityAssocID.ToString)
        End Sub
        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
        End Sub
        Public Sub ChangeKey(ByVal OldKey As String, ByVal NewKey As String)
            If Me.Contains(OldKey) Then
                Dim MyInfo As New Object
                MyInfo = MyBase.Dictionary.Item(OldKey)
                Me.Remove(OldKey)
                Me.Add(MyInfo)
            End If
        End Sub
#End Region
#Region "Overloaded Operators"
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ContactStructInfo)) Then
                Throw New ArgumentException("Only ContactStructs may be inserted into a ContactStruct collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ContactStructInfo)) Then
                Throw New ArgumentException("Only ContactStructs may be removed from a ContactStruct collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.ContactStructInfo)) Then
                Throw New ArgumentException("Only ContactStructs updated in a ContactStruct collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.ContactStructInfo)) Then
                Throw New ArgumentException("Only ContactStructs may be validated in a ContactStruct collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace
