'-------------------------------------------------------------------------------
' MUSTER.Info.TCompanyCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'
' Function          Description
' Item(ID)          Gets/Sets the EntityInfo requested by the string arg ID
' Values()          Returns the collection of Entities in the EntityCollection
' GetKeys()         Returns an array of string containing the keys in the EntityCollection
' Add(Entity)       Adds the Entity supplied to the Entity Collection
' Contains(Entity)  Returns True/False to indicate if the supplied Entity is contained
'                   in the Entity Collection
' Contains(Key)     Returns True/False to indicate if the supplied Entity key is contained
'                   in the Entity Collection
' Remove(Entity)    Removes the Entity supplied from the Entity Collection
'
' NOTE: This file to be used as Company to build other objects.
'       Replace keyword "Company" with respective Object name.
'       Also - don't forget to update the revision history above!!!
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class CompanyCollection
        Inherits DictionaryBase
#Region "Public Events"
        Public Event CompanyColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.CompanyInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.CompanyInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.CompanyInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                    RaiseEvent CompanyColChanged(True)
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent CompanyColChanged(True)
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
        Public Sub Add(ByVal value As MUSTER.Info.CompanyInfo)
            Me.Item(value.ID.ToString) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.CompanyInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.CompanyInfo)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.CompanyInfo)) Then
                Throw New ArgumentException("Only Company may be inserted into a Company collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.CompanyInfo)) Then
                Throw New ArgumentException("Only Company may be removed from a Company collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.CompanyInfo)) Then
                Throw New ArgumentException("Only Company updated in a Company collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.CompanyInfo)) Then
                Throw New ArgumentException("Only Company may be validated in a Company collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace
