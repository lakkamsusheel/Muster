'-------------------------------------------------------------------------------
' MUSTER.Info.UserGroupCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        AN      11/29/04    Original class definition.
'  1.1        JC      12/28/04    Added event UserColChanged to signal that collection 
'                                   was modified
'  1.2        MNR     01/25/05      Added Array.Sort(KeyCol) in GetKeys() Property
'
' Function          Description
' Item(NAME)        Gets/Sets the EntityInfo requested by the string arg NAME
' Values()          Returns the collection of Entities in the EntityCollection
' GetKeys()         Returns an array of string containing the keys in the EntityCollection
' Add(UserGroupInfo)   Adds the Entity supplied to the internal UserGroupCollection
' Contains(UserGroupInfo) Returns True/False to indicate if the supplied EntityInfo is contained
'                           in the internal EntityCollection
' Contains(Key)     Returns True/False to indicate if the supplied EntityInfo key is contained
'                           in the internal EntityCollection
' Remove(UserGroupInfo) Removes the EntityInfo supplied from the internal EntityCollection
'
' Events
' Name              Description
' UserColChanged    Alerts client that the collection was modified
'-------------------------------------------------------------------------------
'
' TODO - 12/29 - Check list of operations
'
Namespace MUSTER.Info
    Public Class UserGroupCollection
        '
        Inherits DictionaryBase
#Region "Public Events"
        Public Event UserColChanged()
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.UserGroupInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.UserGroupInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.UserGroupInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                    RaiseEvent UserColChanged()
                Else
                    MyBase.Dictionary.Item(index) = Value
                    If Value.IsDirty Then
                        RaiseEvent UserColChanged()
                    End If
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

        Public Sub Add(ByVal value As MUSTER.Info.UserGroupInfo)
            Me.Item(value.ID.ToString) = value
        End Sub

        Public Function Contains(ByVal value As MUSTER.Info.UserGroupInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function

        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function

        Public Sub Remove(ByVal value As MUSTER.Info.UserGroupInfo)
            MyBase.Dictionary.Remove(value.ID.ToString)
        End Sub

        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
        End Sub
#End Region
#Region "Overloaded Operators"
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.UserGroupInfo)) Then
                Throw New ArgumentException("Only User Groups may be inserted into a user group collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.UserGroupInfo)) Then
                Throw New ArgumentException("Only User Groups may be removed from a user group collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.UserGroupInfo)) Then
                Throw New ArgumentException("Only User Groups updated in a user group collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.UserGroupInfo)) Then
                Throw New ArgumentException("Only User Groups may be validated in a user group collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace



