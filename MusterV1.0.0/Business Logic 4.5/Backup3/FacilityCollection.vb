'-------------------------------------------------------------------------------
' MUSTER.Info.FacilityCollection
'   Provides a stongly-typed collection for storing Entity objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        EN        12/03/04    Original class definition.
'  1.1        MNR       01/13/05    Added Events
'  1.2        MNR       01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
'  1.3        MNR       01/26/05    Modified Array size in GetKeys() Property
'  1.4        MNR       01/27/05    Added ToString in Subs calling Value.ID
'  1.5        MNR       01/28/05    Added ChangeKey operation for altering keys
'
' Function          Description
' Item(NAME)        Gets/Sets the Facilityinfo requested by the string arg NAME
' Values()          Returns the collection of Facilities in the FacilityCollection
' GetKeys()         Returns an array of string containing the keys in the FacilityCollection
' Add(FacilityInfo)   Adds the Entity supplied to the internal FacilityCollection
' Contains(Facilityinfo) Returns True/False to indicate if the supplied Facilityinfo is contained
'                           in the internal FacilityCollection
' Contains(Key)     Returns True/False to indicate if the supplied Facilityinfo key is contained
'                           in the internal FacilityCollection
' Remove(Facilityinfo) Removes the Facilityinfo supplied from the internal FacilityCollection
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    Public Class FacilityCollection
        '
        Inherits DictionaryBase
#Region "Public Events"
        Public Event FacilityColChanged(ByVal strSrc As String)
#End Region
#Region "Exposed Operations"
        Default Public Property Item(ByVal index As String) As MUSTER.Info.FacilityInfo
            Get
                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.FacilityInfo)
            End Get
            Set(ByVal Value As MUSTER.Info.FacilityInfo)
                If MyBase.Dictionary.Contains(index) Then
                    MyBase.Dictionary.Item(index) = Value
                    RaiseEvent FacilityColChanged(Me.ToString)
                Else
                    MyBase.Dictionary.Add(Value.ID.ToString, Value)
                    RaiseEvent FacilityColChanged(Me.ToString)
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
        Public Sub Add(ByVal value As MUSTER.Info.FacilityInfo)
            Me.Item(value.ID.ToString) = value
        End Sub
        Public Function Contains(ByVal value As MUSTER.Info.FacilityInfo) As Boolean
            Return MyBase.Dictionary.Contains(value.ID.ToString)
        End Function
        Public Function Contains(ByVal Key As String) As Boolean
            Return MyBase.Dictionary.Contains(Key)
        End Function
        Public Sub Remove(ByVal value As MUSTER.Info.FacilityInfo)
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
            If Not value.GetType() Is (GetType(MUSTER.Info.FacilityInfo)) Then
                Throw New ArgumentException("Only Facility Info may be inserted into a Facility collection!", "value")
            End If
        End Sub
        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.FacilityInfo)) Then

                Throw New ArgumentException("Only Facility may be removed from a Facility collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
            If Not newvalue.GetType() Is (GetType(MUSTER.Info.FacilityInfo)) Then
                Throw New ArgumentException("Only Facility updated in a Facility collection!", "value")
            End If
        End Sub

        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
            If Not value.GetType() Is (GetType(MUSTER.Info.FacilityInfo)) Then
                Throw New ArgumentException("Only Facility may be validated in a Facility collection!", "value")
            End If
        End Sub
#End Region
    End Class
End Namespace
