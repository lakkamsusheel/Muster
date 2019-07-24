'-------------------------------------------------------------------------------
' MUSTER.Info.AdvancedSearchCollection
'   Provides a stongly-typed collection for storing AdvancedSearch objects
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        EN        12/10/04    Original class definition.
'  1.1        MNR       01/25/05    Added Array.Sort(KeyCol) in GetKeys() Property
'
' Function          Description
' Item(NAME)        Gets/Sets the AdvancedSearchObject requested by the string arg NAME
' Values()          Returns the collection of AdvancedSearches in the AdvancedSearchCollection
' GetKeys()         Returns an array of string containing the keys in the AdvancedSearchCollection
' Add(AdvancedSearchInfo)   Adds the Entity supplied to the internal AdvancedSearchCollection
' Contains(AdvancedSearchInfo) Returns True/False to indicate if the supplied AdvancedSearchInfo is contained
'                           in the internal AdvancedSearchCollection
' Contains(Key)     Returns True/False to indicate if the supplied AdvancedSearchInfo key is contained
'                           in the internal AdvancedSearchCollection
' Remove(AdvancedSearchInfo) Removes the AdvancedSearchInfo supplied from the internal AdvancedSearchCollection
'-------------------------------------------------------------------------------------------------------------------------
'Namespace MUSTER.Info
'    Public Class AdvancedSearchCollection
'        '
'        Inherits DictionaryBase
'#Region "Exposed Operations"
'        Default Public Property Item(ByVal index As String) As MUSTER.Info.AdvancedSearchInfo
'            Get
'                Return CType(MyBase.Dictionary.Item(index), MUSTER.Info.AdvancedSearchInfo)
'            End Get
'            Set(ByVal Value As MUSTER.Info.FacilityInfo)
'                If MyBase.Dictionary.Contains(CInt(index)) Then
'                    MyBase.Dictionary.Item(CInt(index)) = Value
'                Else
'                    MyBase.Dictionary.Add(Value.ID, Value)
'                End If
'            End Set

'        End Property

'        Public ReadOnly Property Values() As ICollection
'            Get
'                Return MyBase.Dictionary.Values
'            End Get
'        End Property

'        Public ReadOnly Property GetKeys() As String()
'            Get
'                Dim KeyCol(MyBase.Dictionary.Keys.Count - 1) As String
'                MyBase.Dictionary.Keys.CopyTo(KeyCol, 0)
'                Array.Sort(KeyCol)
'                Return KeyCol
'            End Get
'        End Property

'        Public Sub Add(ByVal value As MUSTER.Info.AdvancedSearchInfo)
'            Me.Item(value.ID) = value
'        End Sub

'        Public Function Contains(ByVal value As MUSTER.Info.AdvancedSearchInfo) As Boolean
'            Return MyBase.Dictionary.Contains(value.ID)
'        End Function

'        Public Function Contains(ByVal Key As String) As Boolean
'            Return MyBase.Dictionary.Contains(Key)
'        End Function

'        Public Sub Remove(ByVal value As MUSTER.Info.AdvancedSearchInfo)
'            MyBase.Dictionary.Remove(value.ID)
'        End Sub

'        Public Sub Remove(ByVal value As String)
'            MyBase.Dictionary.Remove(value)
'        End Sub

'#End Region
'#Region "Overloaded Operators"
'        Protected Overloads Sub OnInsert(ByVal Key As Object, ByVal value As Object)
'            If Not value.GetType() Is (GetType(MUSTER.Info.FacilityInfo)) Then
'                Throw New ArgumentException("Only User Groups may be inserted into a user group collection!", "value")
'            End If
'        End Sub

'        Protected Overloads Sub OnRemove(ByVal Key As Object, ByVal value As Object)
'            If Not value.GetType() Is (GetType(MUSTER.Info.FacilityInfo)) Then
'                Throw New ArgumentException("Only User Groups may be removed from a user group collection!", "value")
'            End If
'        End Sub

'        Protected Overloads Sub OnSet(ByVal Key As Object, ByVal oldvalue As Object, ByVal newvalue As Object)
'            If Not newvalue.GetType() Is (GetType(MUSTER.Info.FacilityInfo)) Then
'                Throw New ArgumentException("Only User Groups updated in a user group collection!", "value")
'            End If
'        End Sub

'        Protected Overloads Sub OnValidate(ByVal Key As Object, ByVal value As Object)
'            If Not value.GetType() Is (GetType(MUSTER.Info.FacilityInfo)) Then
'                Throw New ArgumentException("Only User Groups may be validated in a user group collection!", "value")
'            End If
'        End Sub
'#End Region
'    End Class
'End Namespace



