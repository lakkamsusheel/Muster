
Namespace MUSTER.Info
    ' -------------------------------------------------------------------------------
    ' MUSTER.Info._ProtoCollection
    ' Provides a stongly-typed collection for storing _ProtoInfo objects
    ' 
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    ' 
    ' Release   Initials    Date        Description
    ' 1.0        JVC	    ??/??/05    Original class definition.
    ' 
    ' Function          Description
    ' ------------------------------------------------------------------------------------
    ' 
    Public Class FinancialTextCollection
        Inherits DictionaryBase
        Public Delegate Sub FinTextColChangedEventHandler()
        Public Delegate Sub FinTextColErrorEventHandler(ByVal msg As String)
        ' Raised whenever a member is inserted into or removed from the TecDoc collection
        Public Event FinTextColChanged As FinTextColChangedEventHandler
        Public Event FinTextColError As FinTextColErrorEventHandler
        Private oFinancialTextInfo As FinancialTextInfo = New MUSTER.Info.FinancialTextInfo
        Public Sub Add(ByVal value As MUSTER.Info.FinancialTextInfo)
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
        Public Function Contains(ByVal value As MUSTER.Info.FinancialTextInfo) As Boolean
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
        Default Public Property Item(ByVal index As String) As MUSTER.Info.FinancialTextInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.FinancialTextInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.FinancialTextInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent FinTextColChanged()
                End If
            End Set
        End Property
        Public Sub Remove(ByVal value As MUSTER.Info.FinancialTextInfo)
            MyBase.Dictionary.Remove(value.ID.ToString)
            RaiseEvent FinTextColChanged()
        End Sub
        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
            RaiseEvent FinTextColChanged()
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
