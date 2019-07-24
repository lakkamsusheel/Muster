' -------------------------------------------------------------------------------
' MUSTER.Info.FinancialReimbursementCollection
' Provides the container to persist MUSTER FinancialReimbursement state
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
    Public Class FinancialReimbursementCollection
        Inherits DictionaryBase
        Public Delegate Sub FinancialReimbursementColChangedEventHandler()
        Public Delegate Sub FinancialReimbursementColErrorEventHandler(ByVal msg As String)

        Public Event FinancialReimbursementColChanged As FinancialReimbursementColChangedEventHandler
        Public Event FinancialReimbursementColError As FinancialReimbursementColErrorEventHandler

        Public Sub Add(ByVal value As MUSTER.Info.FinancialReimbursementInfo)
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
        Public Function Contains(ByVal value As MUSTER.Info.FinancialReimbursementInfo) As Boolean
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
        Default Public Property Item(ByVal index As String) As MUSTER.Info.FinancialReimbursementInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.FinancialReimbursementInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.FinancialReimbursementInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent FinancialReimbursementColChanged()
                End If
            End Set
        End Property
        Public Sub Remove(ByVal value As MUSTER.Info.FinancialReimbursementInfo)
            MyBase.Dictionary.Remove(value.ID.ToString)
            RaiseEvent FinancialReimbursementColChanged()
        End Sub
        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
            RaiseEvent FinancialReimbursementColChanged()
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
