'-------------------------------------------------------------------------------
' MUSTER.Info.PipesCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MNR       12/03/04    Original class definition.
'  1.1        EN        01/06/05    Added Events and raised the Events.
'  1.2        EN        01/19/05    Added Source Column in Event. 
'  1.3        MNR       01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
'  1.4        MNR       01/27/05    Added ToString in Subs calling Value.ID
'  1.5        MNR       01/28/05    Added ChangeKey operation for altering keys
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
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    Public Class PipesCollection
        '
        Inherits DictionaryBase
#Region "Public Events"
        Public Event InfoChanged(ByVal strSrc As String)
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.PipeInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.PipeInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.PipeInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(Value.ID, Value)
                    RaiseEvent InfoChanged(Me.ToString)
                Else
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent InfoChanged(Me.ToString)
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
        Public Sub Add(ByVal value As MUSTER.Info.PipeInfo)
            Me.Item(value.ID) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.PipeInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.PipeInfo)
            MyBase.Dictionary.Remove(value.ID)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.PipeInfo)) Then
                Throw New ArgumentException("Only Pipes may be inserted into a Pipe collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.PipeInfo)) Then
                Throw New ArgumentException("Only Pipes may be removed from a Pipe collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.PipeInfo)) Then
                Throw New ArgumentException("Only Pipes updated in a Pipe collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.PipeInfo)) Then
                Throw New ArgumentException("Only Pipes may be validated in a Pipe collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace
