'-------------------------------------------------------------------------------
' MUSTER.Info.LetterCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        EN      12/16/04      Original class definition.
'  1.1        MNR     01/25/05      Added Array.Sort(KeyCol) in GetKeys() Property
'  1.2        MNR     01/26/05      Modified Array size in GetKeys() Property
'  1.3        MNR     01/27/05      Added ToString in Subs calling Value.ID
'
' Function          Description
' Item(NAME)        Gets/Sets the Letterinfo requested by the string arg NAME
' Values()          Returns the collection of Letter in the LetterCollection
' GetKeys()         Returns an array of string containing the keys in the LetterCollection
' Add(Letterinfo)   Adds the Letter supplied to the internal LetterCollection
' Contains(Letterinfo) Returns True/False to indicate if the supplied Letterinfo is contained
'                           in the internal LetterCollection
' Contains(Key)     Returns True/False to indicate if the supplied Letterinfo key is contained
'                           in the internal LetterCollection
' Remove(Letterinfo) Removes the Letterinfo supplied from the internal LetterCollectionf
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class LetterCollection
        Inherits DictionaryBase
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.LetterInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.LetterInfo)
            End Get

            Set(ByVal Value As MUSTER.Info.LetterInfo)
                If MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Item(index) = Value
                Else
                    MyBase.Dictionary.Add(CType(Value.ID, String), Value)
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

        Public Sub Add(ByVal value As MUSTER.Info.LetterInfo)
            Me.Item(CType(value.ID, String)) = value
        End Sub

        Public Function Contains(ByVal value As MUSTER.Info.LetterInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function

        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function

        Public Sub Remove(ByVal value As MUSTER.Info.LetterInfo)
            MyBase.Dictionary.Remove(CType(value.ID, String))
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
            If Not value.GetType() Is (GetType(MUSTER.Info.LetterInfo)) Then
                Throw New ArgumentException("Only Letter may be inserted into a Lettercollection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.LetterInfo)) Then

                Throw New ArgumentException("Only Letter may be removed from a Lettercollection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.LetterInfo)) Then
                Throw New ArgumentException("Only Letter updated in a Lettercollection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.LetterInfo)) Then
                Throw New ArgumentException("Only Letter may be validated in a user Lettercollection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace



