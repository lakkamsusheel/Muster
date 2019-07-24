
Namespace MUSTER.Info
    ' -------------------------------------------------------------------------------
    ' MUSTER.Info.FeeInvoiceCollection
    '       Provides a stongly-typed collection for storing Fee Basis objects
    ' 
    '   Copyright (C) 2004 CIBER, Inc.
    '   All rights reserved.
    ' 
    '   Release   Initials    Date        Description
    '     1.0           JC      06/14/05   Original class definition
    ' -------------------------------------------------------------------------------
    Public Class FeeInvoiceCollection
        '
        Inherits DictionaryBase
#Region "Public Events"
        Public Event FeeInvoiceColChanged(ByVal strSrc As String)
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.FeeInvoiceInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.FeeInvoiceInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.FeeInvoiceInfo)
                If MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent FeeInvoiceColChanged(Me.ToString)
                Else
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                    RaiseEvent FeeInvoiceColChanged(Me.ToString)
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
        Public Sub Add(ByVal value As MUSTER.Info.FeeInvoiceInfo)
            Me.Item(value.ID.ToString) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.FeeInvoiceInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.FeeInvoiceInfo)
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
            'If Not value.GetType() Is (GetType(MUSTER.Info.FeeInvoiceInfo)) Then
            '    Throw New ArgumentException("Only Facility Info may be inserted into a Facility collection!", "value")
            'End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            'If Not value.GetType() Is (GetType(MUSTER.Info.FeeInvoiceInfo)) Then

            '    Throw New ArgumentException("Only Facility may be removed from a Facility collection!", "value")
            'End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            'If Not newvalue.GetType() Is (GetType(MUSTER.Info.FeeInvoiceInfo)) Then
            '    Throw New ArgumentException("Only Facility updated in a Facility collection!", "value")
            'End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            'If Not value.GetType() Is (GetType(MUSTER.Info.FeeInvoiceInfo)) Then
            '    Throw New ArgumentException("Only Facility may be validated in a Facility collection!", "value")
            'End If
        End Sub
#End Region
    End Class
End Namespace

