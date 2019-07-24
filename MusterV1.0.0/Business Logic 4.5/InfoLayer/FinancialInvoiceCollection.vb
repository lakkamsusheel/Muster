' -------------------------------------------------------------------------------
' MUSTER.Info.FinancialInvoiceCollection
' Provides the container to persist MUSTER FinancialInvoice state
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

    Public Class FinancialInvoiceCollection
        Inherits DictionaryBase
        Public Delegate Sub FinancialInvoiceColChangedEventHandler()
        Public Delegate Sub FinancialInvoiceColErrorEventHandler(ByVal msg As String)

        Public Event FinancialInvoiceColChanged As FinancialInvoiceColChangedEventHandler
        Public Event FinancialInvoiceColError As FinancialInvoiceColErrorEventHandler

        Public Sub Add(ByVal value As MUSTER.Info.FinancialInvoiceInfo)
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
        Public Function Contains(ByVal value As MUSTER.Info.FinancialInvoiceInfo) As Boolean
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
        Default Public Property Item(ByVal index As String) As MUSTER.Info.FinancialInvoiceInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.FinancialInvoiceInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.FinancialInvoiceInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent FinancialInvoiceColChanged()
                End If
            End Set
        End Property
        Public Sub Remove(ByVal value As MUSTER.Info.FinancialInvoiceInfo)
            MyBase.Dictionary.Remove(value.ID.ToString)
            RaiseEvent FinancialInvoiceColChanged()
        End Sub
        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
            RaiseEvent FinancialInvoiceColChanged()
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

