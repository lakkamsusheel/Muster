'-------------------------------------------------------------------------------
' MUSTER.Info.AppFlagsCollection
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
' NOTE: This file to be used as AppFlag to build other objects.
'       Replace keyword "AppFlag" with respective Object name.
'       Also - don't forget to update the revision history above!!!
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class AppFlagsCollection
        Inherits DictionaryBase
#Region "Public Events"
        Public Event AppFlagColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.AppFlagInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.AppFlagInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.AppFlagInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.Key.ToString, Value)
                    RaiseEvent AppFlagColChanged(True)
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent AppFlagColChanged(True)
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
        Public Sub Add(ByVal value As MUSTER.Info.AppFlagInfo)
            Me.Item(value.Key.ToString) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.AppFlagInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.Key.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.AppFlagInfo)
            MyBase.Dictionary.Remove(value.Key.ToString)
        End Sub
        Public Sub Remove(ByVal value As String)
            MyBase.Dictionary.Remove(value)
        End Sub
#End Region
#Region "Overloaded Operators"
        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.AppFlagInfo)) Then
                Throw New ArgumentException("Only AppFlags may be inserted into a AppFlag collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.AppFlagInfo)) Then
                Throw New ArgumentException("Only AppFlags may be removed from a AppFlag collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.AppFlagInfo)) Then
                Throw New ArgumentException("Only AppFlags updated in a AppFlag collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.AppFlagInfo)) Then
                Throw New ArgumentException("Only AppFlags may be validated in a AppFlag collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace
