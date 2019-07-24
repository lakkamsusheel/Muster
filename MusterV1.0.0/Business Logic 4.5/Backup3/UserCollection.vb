'-------------------------------------------------------------------------------
' MUSTER.Info.UserCollection
'   Provides a stongly-typed collection for storing User objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC2      12/04/04    Original class definition.
'  1.1        MNR       01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
'
' Function          Description
' Item(NAME)        Gets/Sets the UserInfo requested by the string arg NAME
' Values()          Returns the collection of Users in the UserCollection
' GetKeys()         Returns an array of string containing the keys in the UserCollection
' Add(UserInfo)   Adds the Entity supplied to the internal UserCollection
' Contains(UserInfo) Returns True/False to indicate if the supplied UserInfo is contained
'                           in the internal UserCollection
' Contains(Key)     Returns True/False to indicate if the supplied UserInfo key is contained
'                           in the internal UserCollection
' Remove(UserInfo) Removes the UserInfo supplied from the internal UserCollection
'-------------------------------------------------------------------------------
'
' TODO - Add to app 1/3/2005 - JVC 2
'

Namespace MUSTER.Info
    Public Class UserCollection
        '
        Inherits DictionaryBase
#Region "Public Events"
        Public Event UserColChanged()
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.UserInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.UserInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.UserInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.UserKey.ToString, Value)
                    RaiseEvent UserColChanged()
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent UserColChanged()
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

        Public Sub Add(ByVal value As MUSTER.Info.UserInfo)
            Me.Item(value.UserKey.ToString) = value
        End Sub

        Public Function Contains(ByVal value As MUSTER.Info.UserInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.UserKey.ToString)
        End Function

        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function

        Public Sub Remove(ByVal value As MUSTER.Info.UserInfo)
            MyBase.Dictionary.Remove(value.UserKey.ToString)
        End Sub


        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
        End Sub

#End Region
#Region "Overloaded Operators"
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.UserInfo)) Then
                Throw New ArgumentException("Only User may be inserted into a User collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.UserInfo)) Then
                Throw New ArgumentException("Only User may be removed from a User collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.UserInfo)) Then
                Throw New ArgumentException("Only Users updated in a User collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.UserInfo)) Then
                Throw New ArgumentException("Only Users may be validated in a User collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace
