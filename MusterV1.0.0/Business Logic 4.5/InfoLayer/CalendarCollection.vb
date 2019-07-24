'-------------------------------------------------------------------------------
' MUSTER.Info.CalendarCollection
'   Provides a stongly-typed collection for storing Calendar objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        KJ        12/10/04    Original class definition.
'  1.1        KJ        01/06/05    Added Events for infoChanged. Also changed the Dictionary.Add
'  1.2        MNR       01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
'  1.3        MNR       01/27/05    Added ToString in Subs calling Value.ID
'  1.4        MR        01/28/05    Added ChangeKey operation for altering keys
'
' Function              Description
' Item(CalendarInfoId)       Gets/Sets the CalendarInfo requested by the arg CalendarInfoId
' Values()              Returns the collection of Calendar in the CalendarCollection
' GetKeys()             Returns an array of string containing the keys in the CalendarCollection
' Add(CalendarInfo)      Adds the Address supplied to the internal CalendarCollection
' Contains(CalendarInfo) Returns True/False to indicate if the supplied CalendarInfo is contained
'                           in the internal CalendarCollection
' Contains(Key)         Returns True/False to indicate if the supplied CalendarInfo key is contained
'                           in the internal CalendarCollection
' Remove(CalendarInfo)   Removes the CalendarInfo supplied from the internal CalendarCollection
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    Public Class CalendarCollection
        Inherits DictionaryBase
#Region "Public Events"
        Public Event InfoChanged()
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.CalendarInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.CalendarInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.CalendarInfo)
                If Not MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Add(CType(Value.CalendarInfoId, String), Value)
                    RaiseEvent InfoChanged()
                Else
                    MyBase.Dictionary.Item(index) = Value
                    If Value.IsDirty Then
                        RaiseEvent InfoChanged()
                    End If
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

        Public Sub Add(ByVal value As MUSTER.Info.CalendarInfo)
            Me.Item(CType(value.CalendarInfoId, String)) = value
        End Sub

        Public Function Contains(ByVal value As MUSTER.Info.CalendarInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.CalendarInfoId.ToString)
        End Function

        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function

        Public Sub Remove(ByVal value As MUSTER.Info.CalendarInfo)
            MyBase.Dictionary.Remove(value.CalendarInfoId.ToString)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.CalendarInfo)) Then
                Throw New ArgumentException("Only Reports may be inserted into a report collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.CalendarInfo)) Then
                Throw New ArgumentException("Only Reports may be removed from a report collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.CalendarInfo)) Then
                Throw New ArgumentException("Only Reports updated in a report collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.CalendarInfo)) Then
                Throw New ArgumentException("Only Reports may be validated in a report collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace

