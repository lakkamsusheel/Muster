' -------------------------------------------------------------------------------
' MUSTER.Info.FinancialCommitmentCollection
' Provides the container to persist MUSTER FinancialCommitment state
' 
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0        AB       06/24/05    Original class definition.
' 
' Function          Description
' ---

Namespace MUSTER.Info

    Public Class FinancialCommitmentCollection
        Inherits DictionaryBase
        Public Delegate Sub FinancialColChangedEventHandler()
        Public Delegate Sub FinancialColErrorEventHandler(ByVal msg As String)

        Public Event FinancialColChanged As FinancialColChangedEventHandler
        Public Event FinancialColError As FinancialColErrorEventHandler

        Public Sub Add(ByVal value As MUSTER.Info.FinancialCommitmentInfo)
            Me.Item(value.CommitmentID.ToString) = value
        End Sub
        Public Sub ChangeKey(ByVal OldKey As String, ByVal NewKey As String)
            If Me.Contains(OldKey) Then
                Dim MyInfo As New Object
                MyInfo = MyBase.Dictionary.Item(OldKey)
                Me.Remove(OldKey)
                Me.Add(MyInfo)
            End If
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.FinancialCommitmentInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.CommitmentID.ToString)
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
        Default Public Property Item(ByVal index As String) As MUSTER.Info.FinancialCommitmentInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.FinancialCommitmentInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.FinancialCommitmentInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.CommitmentID.ToString, Value)
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent FinancialColChanged()
                End If
            End Set
        End Property
        Public Sub Remove(ByVal value As MUSTER.Info.FinancialCommitmentInfo)
            MyBase.Dictionary.Remove(value.CommitmentID.ToString)
            RaiseEvent FinancialColChanged()
        End Sub
        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
            RaiseEvent FinancialColChanged()
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
