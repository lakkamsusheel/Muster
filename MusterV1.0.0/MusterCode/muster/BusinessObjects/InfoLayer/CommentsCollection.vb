
'-------------------------------------------------------------------------------
' MUSTER.Info.CommentsCollection
'   Provides a stongly-typed collection for storing Comments objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        PN        12/13/04    Original class definition.
'  1.1        MNR       01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
'  1.2        MNR       01/27/05    Added ToString in Subs calling Value.ID
'  1.3        AN        02/02/05      Added ChangeKey operation for altering keys
' Function          Description
' Item(NAME)        Gets/Sets the CommentsInfo requested by the string arg NAME
' Values()          Returns the collection of Comments in the CommentsCollection
' GetKeys()         Returns an array of string containing the keys in the CommentsCollection
' Add(CommentsInfo)   Adds the Entity supplied to the internal CommentsCollection
' Contains(CommentsInfo) Returns True/False to indicate if the supplied UserInfo is contained
'                           in the internal CommentsCollection
' Contains(Key)     Returns True/False to indicate if the supplied CommentsInfo key is contained
'                           in the internal CommentsCollection
' Remove(CommetnsInfo) Removes the CommentsInfo supplied from the internal CommentsCollection
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class CommentsCollection
        '
        Inherits DictionaryBase
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.CommentsInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.CommentsInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.CommentsInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
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

        Public Sub Add(ByVal value As MUSTER.Info.CommentsInfo)
            Me.Item(value.ID.ToString) = value
        End Sub

        Public Function Contains(ByVal value As MUSTER.Info.CommentsInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function

        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function

        Public Sub Remove(ByVal value As MUSTER.Info.CommentsInfo)
            MyBase.Dictionary.Remove(value.ID.ToString)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.CommentsInfo)) Then
                Throw New ArgumentException("Only User may be inserted into a User collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.CommentsInfo)) Then
                Throw New ArgumentException("Only User may be removed from a User collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.CommentsInfo)) Then
                Throw New ArgumentException("Only Users updated in a User collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.CommentsInfo)) Then
                Throw New ArgumentException("Only Users may be validated in a User collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace

