'-------------------------------------------------------------------------------
' MUSTER.Info.UserGroupCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        AN      11/29/04    Original class definition.
'  1.1        MNR     01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
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
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class UserAuditPointCollection
        Inherits DictionaryBase
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.UserAuditPoint
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.UserAuditPoint)
            End Get
            Set(ByVal Value As MUSTER.Info.UserAuditPoint)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.GUID, Value)
                Else
                    MyBase.Dictionary.Item(index) = Value
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

        Public Sub Add(ByVal value As MUSTER.Info.UserAuditPoint)
            Me.Item(value.GUID) = value
        End Sub

        Public Function Contains(ByVal value As MUSTER.Info.UserAuditPoint) As Boolean
            Return MyBase.Dictionary.Contains(value.GUID)
        End Function

        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function

        Public Sub Remove(ByVal value As MUSTER.Info.UserAuditPoint)
            MyBase.Dictionary.Remove(value.GUID)
        End Sub


        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
        End Sub

#End Region
#Region "Overloaded Operators"
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.UserAuditPoint)) Then
                Throw New ArgumentException("Only User Groups may be inserted into a user group collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.UserAuditPoint)) Then
                Throw New ArgumentException("Only User Groups may be removed from a user group collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.UserAuditPoint)) Then
                Throw New ArgumentException("Only User Groups updated in a user group collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.UserAuditPoint)) Then
                Throw New ArgumentException("Only User Groups may be validated in a user group collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace
