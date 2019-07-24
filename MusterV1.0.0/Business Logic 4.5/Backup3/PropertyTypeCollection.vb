'-------------------------------------------------------------------------------
' MUSTER.Info.MusterPropertyTypeCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       Elango      11/19/04    Original class definition.
'  1.1        MNR        01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
'
' Function          Description
' 
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class PropertyTypeCollection
        Inherits DictionaryBase
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.PropertyTypeInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.PropertyTypeInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.PropertyTypeInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                Else
                    MyBase.Dictionary.Item(index.ToString) = Value
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
        Public Sub Add(ByVal value As MUSTER.Info.PropertyTypeInfo)
            Me.Item(value.ID.ToString) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.PropertyTypeInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function
        Public Function Contains(ByVal Name As String) As Boolean
            Return MyBase.Dictionary.Contains(Name)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.PropertyTypeInfo)
            MyBase.Dictionary.Remove(value.ID.ToString)
        End Sub
#End Region
#Region "Overloaded Operators"
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.PropertyTypeInfo)) Then
                Throw New ArgumentException("Only Property may be validated in an Property collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.PropertyTypeInfo)) Then
                Throw New ArgumentException("Only Property may be validated in an Property collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.PropertyTypeInfo)) Then
                Throw New ArgumentException("Only Property may be validated in an Property collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.PropertyTypeInfo)) Then
                Throw New ArgumentException("Only Property may be validated in an Property collection!", "value")
            End If
        End Sub
#End Region
#Region "Public Events"
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
    End Class
End Namespace




